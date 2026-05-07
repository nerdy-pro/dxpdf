//! Extract floating (anchor) images and shapes from paragraph inlines, and
//! resolve their positions in the caller-specified coordinate frame.

use crate::model::{self, Paragraph};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::layout::section::{FloatingImage, FloatingImageY, FloatingShape};
use crate::render::resolve::shape_geometry::build_geometry;
use crate::render::resolve::shape_visuals::resolve_shape_visuals;

use super::{BuildContext, BuildState};

/// Coordinate frame in which an anchor's position is resolved.
///
/// The choice of frame determines both the origin used as the zero of the
/// horizontal axis and how §20.4.2 vertical references map onto the
/// `FloatingImageY` ADT.
///
/// * `Page` — page-absolute coordinates. The horizontal origin is the page
///   left edge and §20.4.2.10 `AnchorRelativeFrom` references resolve against
///   the page's own margins. Callers use this frame when the emitted command
///   is appended directly to the page command list without a further shift.
///
/// * `Stack` — relative to a stack frame origin (table cell top-left, or
///   header/footer content-area top-left). All horizontal references collapse
///   to the frame's left edge, and vertical offsets are stored as
///   `RelativeToParagraph` so the stacker anchors them to the owning
///   paragraph. Callers use this frame when the emitted command passes
///   through `stack_blocks` and will be shifted by the caller into page
///   coordinates.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum AnchorFrame {
    Page,
    Stack,
}

/// Extract floating (anchor) images from a paragraph's inlines.
///
/// Positions are resolved in the coordinate system implied by `frame`
/// (see [`AnchorFrame`]).
pub(super) fn extract_floating_images(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &BuildState,
    frame: AnchorFrame,
) -> Vec<FloatingImage> {
    use crate::model::{GraphicContent, ImagePlacement, Inline};

    fn find_anchor_images<'a>(inlines: &'a [Inline], out: &mut Vec<&'a crate::model::Image>) {
        for inline in inlines {
            match inline {
                // Images with a WordProcessingShape graphic are handled by
                // `extract_floating_shapes`; skip them here so the shape
                // branch owns their layout path end-to-end.
                Inline::Image(img)
                    if matches!(img.placement, ImagePlacement::Anchor(_))
                        && !matches!(img.graphic, Some(GraphicContent::WordProcessingShape(_))) =>
                {
                    out.push(img);
                }
                Inline::Hyperlink(link) => find_anchor_images(&link.content, out),
                Inline::Field(f) => find_anchor_images(&f.content, out),
                Inline::AlternateContent(ac) => {
                    if let Some(ref fb) = ac.fallback {
                        find_anchor_images(fb, out);
                    }
                }
                _ => {}
            }
        }
    }

    let mut anchor_imgs = Vec::new();
    find_anchor_images(&para.content, &mut anchor_imgs);

    let mut images = Vec::new();
    for img in anchor_imgs {
        let ImagePlacement::Anchor(ref anchor) = img.placement else {
            continue;
        };
        let Some(rel_id) = crate::render::resolve::images::extract_image_rel_id(img) else {
            continue;
        };
        let Some(image_data) = ctx.resolved.media.get(rel_id).cloned() else {
            log::warn!(
                "anchor image: rel_id={} missing from media table ({} entries)",
                rel_id.as_str(),
                ctx.resolved.media.len(),
            );
            continue;
        };

        let w = Pt::from(img.extent.width);
        let h = Pt::from(img.extent.height);
        let (x, y) = resolve_anchor_position(anchor, w, h, state, frame);

        images.push(FloatingImage {
            image_data,
            size: PtSize::new(w, h),
            x,
            y,
            wrap_mode: crate::render::layout::section::WrapMode::from_model(&anchor.wrap),
            dist_left: Pt::from(anchor.distance.left),
            dist_right: Pt::from(anchor.distance.right),
            behind_doc: anchor.behind_text,
        });
    }

    images
}

// ── Floating shape extraction ──────────────────────────────────────────────

/// Extract floating (anchor) DrawingML shapes from a paragraph's inlines,
/// resolve their geometry + visuals, and compute their positions in the
/// coordinate frame implied by `frame`. Pure: takes immutable references to
/// `ctx` / `state`.
pub(super) fn extract_floating_shapes(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &BuildState,
    frame: AnchorFrame,
) -> Vec<FloatingShape> {
    use crate::model::{GraphicContent, ImagePlacement, Inline};

    fn find_anchor_shapes<'a>(inlines: &'a [Inline], out: &mut Vec<&'a crate::model::Image>) {
        for inline in inlines {
            match inline {
                Inline::Image(img)
                    if matches!(img.placement, ImagePlacement::Anchor(_))
                        && matches!(img.graphic, Some(GraphicContent::WordProcessingShape(_))) =>
                {
                    out.push(img);
                }
                Inline::Hyperlink(link) => find_anchor_shapes(&link.content, out),
                Inline::Field(f) => find_anchor_shapes(&f.content, out),
                // MCE §M.1.2: shapes live inside the `<mc:Choice Requires="wps">`
                // branch; the `<mc:Fallback>` carries the VML equivalent. We
                // scan both: the first choice that yields a shape wins, else
                // we try the fallback (which will be ignored at build time
                // anyway because VML has no `WordProcessingShape` graphic).
                Inline::AlternateContent(ac) => {
                    let before = out.len();
                    for choice in &ac.choices {
                        find_anchor_shapes(&choice.content, out);
                    }
                    if out.len() == before {
                        if let Some(ref fb) = ac.fallback {
                            find_anchor_shapes(fb, out);
                        }
                    }
                }
                _ => {}
            }
        }
    }

    let mut shape_imgs = Vec::new();
    find_anchor_shapes(&para.content, &mut shape_imgs);

    let mut shapes = Vec::new();
    for img in shape_imgs {
        let ImagePlacement::Anchor(ref anchor) = img.placement else {
            continue;
        };
        let wsp = match img.graphic.as_ref() {
            Some(GraphicContent::WordProcessingShape(w)) => w,
            _ => continue,
        };
        let shape_props = wsp.shape_properties.as_ref();
        let geometry = match shape_props.and_then(|p| p.geometry.as_ref()) {
            Some(g) => g,
            None => continue, // No geometry → nothing to draw.
        };

        let w = Pt::from(img.extent.width);
        let h = Pt::from(img.extent.height);
        let extent = PtSize::new(w, h);

        let shape_path = match build_geometry(geometry, extent) {
            Some(p) => p,
            None => continue, // Unimplemented preset or empty geometry.
        };

        let visuals = resolve_shape_visuals(
            shape_props,
            wsp.style_line_ref.as_ref(),
            wsp.style_effect_ref.as_ref(),
            ctx.resolved.theme.as_ref(),
        );

        // §20.1.7.6 transform attributes (rotation/flip) live on the shape's
        // `spPr/xfrm`; anchor position is independent.
        let (rotation, flip_h, flip_v) = shape_props
            .and_then(|p| p.transform.as_ref())
            .map(|t| {
                (
                    t.rotation
                        .unwrap_or_else(|| crate::model::dimension::Dimension::new(0)),
                    t.flip_h.unwrap_or(false),
                    t.flip_v.unwrap_or(false),
                )
            })
            .unwrap_or((crate::model::dimension::Dimension::new(0), false, false));

        let (x, y) = resolve_anchor_position(anchor, w, h, state, frame);

        shapes.push(FloatingShape {
            x,
            y,
            size: extent,
            rotation,
            flip_h,
            flip_v,
            wrap_mode: crate::render::layout::section::WrapMode::from_model(&anchor.wrap),
            dist_left: Pt::from(anchor.distance.left),
            dist_right: Pt::from(anchor.distance.right),
            behind_doc: anchor.behind_text,
            paths: shape_path.paths,
            fill: visuals.fill,
            stroke: visuals.stroke,
            effects: visuals.effects,
        });
    }

    // VML primitives (`<v:rect>` and friends) coexist in the same
    // paragraph and resolve to the same `FloatingShape` shape format.
    // We append them here so both DrawingML and VML floats live in one
    // ordered list passed downstream.
    extract_vml_primitive_shapes(&para.content, state, frame, &mut shapes);

    shapes
}

/// Walk the inlines for `Inline::Pict` containers and emit a
/// [`FloatingShape`] for every renderable VML primitive variant.
/// Phase B handles `<v:rect>`; later phases can extend this to
/// `RoundRect`, `Oval`, `Line`, `PolyLine`, `Image`, and grouped
/// children.
///
/// Position resolution uses [`vml_absolute_position`] — currently
/// page-relative when `position:absolute` and `margin-left`/
/// `margin-top` are present. The vorlage gray-bar pattern fits this
/// shape exactly.
fn extract_vml_primitive_shapes(
    inlines: &[crate::model::Inline],
    state: &BuildState,
    frame: AnchorFrame,
    out: &mut Vec<FloatingShape>,
) {
    use crate::model::Inline;
    for inline in inlines {
        match inline {
            Inline::Pict(pict) => {
                for primitive in &pict.primitives {
                    extract_vml_primitive(primitive, state, frame, out);
                }
            }
            Inline::Hyperlink(link) => {
                extract_vml_primitive_shapes(&link.content, state, frame, out)
            }
            Inline::Field(f) => extract_vml_primitive_shapes(&f.content, state, frame, out),
            // §M.1.2: `<mc:AlternateContent>` carries the same shape
            // twice — modern DrawingML in `<mc:Choice>` and a VML
            // fallback in `<mc:Fallback>` — for older clients. As a
            // modern renderer we honor the Choice (the DrawingML
            // walker `find_anchor_shapes` already extracts it
            // alongside us in `extract_floating_shapes`) and skip the
            // Fallback. Walking it would emit a duplicate shape for
            // the same logical rectangle. If a future case needs the
            // Fallback (e.g. Choice that fails our `Requires`
            // namespace test), we can add a Choice-fed signal here.
            Inline::AlternateContent(_) => {}
            _ => {}
        }
    }
}

fn extract_vml_primitive(
    primitive: &model::VmlPrimitive,
    state: &BuildState,
    frame: AnchorFrame,
    out: &mut Vec<FloatingShape>,
) {
    use crate::model::VmlPrimitive;
    match primitive {
        VmlPrimitive::Rect(r) => {
            if let Some(shape) = build_vml_rect_shape(&r.common, state, frame) {
                out.push(shape);
            }
        }
        // §14.1.2.9: groups carry their own coord system; for now we
        // recurse without applying the transform — most groups in
        // practice use the page coord system anyway. Phase D will
        // honor `coordsize`/`coordorigin`.
        VmlPrimitive::Group(g) => {
            for child in &g.children {
                extract_vml_primitive(child, state, frame, out);
            }
        }
        // Other variants are modeled but not yet emitted as shapes.
        // Their text-box content is still consumed by the inline
        // fragment collector, which keeps `<v:shape>`-style text
        // boxes working as before.
        VmlPrimitive::Shape(_)
        | VmlPrimitive::RoundRect(_)
        | VmlPrimitive::Oval(_)
        | VmlPrimitive::Line(_)
        | VmlPrimitive::PolyLine(_)
        | VmlPrimitive::Arc(_)
        | VmlPrimitive::Curve(_)
        | VmlPrimitive::Image(_) => {}
    }
}

/// Build a [`FloatingShape`] for a `<v:rect>`-like primitive whose
/// `common` carries an absolute position + width/height in its
/// `style`. Returns `None` when the spec-required attributes are
/// absent so the rect can't be placed (in which case Word would
/// silently skip it too).
///
/// The returned shape's `x` lives in the coordinate frame implied by
/// `frame`: in `AnchorFrame::Stack` the downstream emitter (e.g.
/// `render_footer`) adds `margins.left` back, so we subtract it here
/// to keep the page-relative `margin-left` honest.
fn build_vml_rect_shape(
    common: &model::VmlCommonAttrs,
    state: &BuildState,
    frame: AnchorFrame,
) -> Option<FloatingShape> {
    use crate::render::geometry::PtOffset;
    use crate::render::layout::draw_command::ResolvedFill;
    use crate::render::resolve::drawing_color::Rgba;
    use crate::render::resolve::shape_geometry::{PathVerb, SubPath};

    // Position via `position:absolute` + `margin-left/top`. VML's
    // `margin-left` is page-relative when the shape's
    // `mso-position-horizontal-relative` is `page` (the vorlage gray
    // bar's case) — we don't model that style attribute yet, so
    // assume page-relative and let phase D add the discriminator.
    //
    // In `AnchorFrame::Stack` the eventual emitter shifts every
    // command by `margins.left` to convert stack→page; we subtract
    // it from `x` up front so the round-trip preserves the
    // page-relative offset.
    let (page_x, y) = vml_absolute_position(&common.style)?;
    let x = match frame {
        AnchorFrame::Page => page_x,
        AnchorFrame::Stack => page_x - state.page_config.margins.left,
    };

    // Size via `style.width` / `style.height`. A rect with no extent
    // can't meaningfully render.
    let width = common.style.width.map(vml_length_to_pt)?;
    let height = common.style.height.map(vml_length_to_pt)?;
    if width <= Pt::ZERO || height <= Pt::ZERO {
        return None;
    }
    let extent = PtSize::new(width, height);

    // Resolve fill from `@fillcolor`. Tier 0 only handles solid
    // colors; gradients and patterns log-and-fall-through elsewhere.
    let fill = match common.fill_color {
        Some(crate::model::VmlColor::Rgb(r, g, b)) => ResolvedFill::Solid(Rgba {
            r: r as f32 / 255.0,
            g: g as f32 / 255.0,
            b: b as f32 / 255.0,
            a: 1.0,
        }),
        Some(crate::model::VmlColor::Named(_)) | None => ResolvedFill::None,
    };

    // Build a closed-rectangle path in shape-local Pt. The painter
    // applies the `(x, y)` and `size` to position the path.
    let paths = vec![SubPath {
        verbs: vec![
            PathVerb::MoveTo(PtOffset::new(Pt::ZERO, Pt::ZERO)),
            PathVerb::LineTo(PtOffset::new(extent.width, Pt::ZERO)),
            PathVerb::LineTo(PtOffset::new(extent.width, extent.height)),
            PathVerb::LineTo(PtOffset::new(Pt::ZERO, extent.height)),
            PathVerb::Close,
        ],
        fill_mode: crate::model::PathFillMode::Norm,
        stroked: matches!(common.stroked, Some(true)),
    }];

    // Vertical position resolution depends on the
    // `mso-position-vertical-relative` style attribute:
    // * `page`/`margin` → absolute page coordinates
    // * `text`/`paragraph` (Word's default in body and footer) →
    //   relative to the owning paragraph
    //
    // Phase B only honors the latter (the vorlage gray-bar case);
    // page-anchored vertical resolution lands in phase D when the
    // full position resolver is wired in. Until then we treat the
    // y offset as relative to the host paragraph — which matches
    // Word's default and lets the footer rect render correctly.
    let y_image = FloatingImageY::RelativeToParagraph(y);

    Some(FloatingShape {
        x,
        y: y_image,
        size: extent,
        rotation: crate::model::dimension::Dimension::new(0),
        flip_h: false,
        flip_v: false,
        // §14.1.2.16: VML rects don't usually wrap surrounding text
        // — they sit at an absolute z-index. Treat them as wrapNone.
        wrap_mode: crate::render::layout::section::WrapMode::None,
        dist_left: Pt::ZERO,
        dist_right: Pt::ZERO,
        // §14.1.2 z-index drives layering. For Tier 0 we treat all
        // VML primitives as non-behind-text (drawn in document order).
        behind_doc: false,
        paths,
        fill,
        stroke: None,
        effects: vec![],
    })
}

/// Shared anchor-position resolver used by both `extract_floating_images`
/// and `extract_floating_shapes`. Returns `(x, y)` in the coordinate system
/// implied by `frame`.
///
/// See [`AnchorFrame`] for the semantics of each frame. Both horizontal and
/// vertical axes are resolved per §20.4.2.10 `AnchorRelativeFrom` when
/// `frame = Page`; in `Stack` the frame origin collapses every horizontal
/// reference to the frame's left edge (matching the body's left margin) and
/// every vertical offset is carried as `RelativeToParagraph` so the stacker
/// anchors the float to the owning paragraph.
fn resolve_anchor_position(
    anchor: &crate::model::AnchorProperties,
    content_w: Pt,
    content_h: Pt,
    state: &BuildState,
    frame: AnchorFrame,
) -> (Pt, FloatingImageY) {
    let x = resolve_anchor_x(anchor, content_w, state, frame);
    let y = resolve_anchor_y(anchor, content_h, state, frame);
    (x, y)
}

/// Horizontal axis of `resolve_anchor_position`. Split out so the two axes
/// can be read independently.
fn resolve_anchor_x(
    anchor: &crate::model::AnchorProperties,
    content_w: Pt,
    state: &BuildState,
    frame: AnchorFrame,
) -> Pt {
    use crate::model::{AnchorAlignment, AnchorPosition, AnchorRelativeFrom};

    let pc = &state.page_config;
    // The "frame-zeroed" margins are used when the anchor is
    // margin-relative: in `Stack` we treat the body left margin as
    // origin (margin_left = 0) so the result is stack-relative;
    // `render_header`/`render_footer` re-adds the real margin later.
    //
    // Page-relative anchors (`AnchorRelativeFrom::Page`) need a
    // different convention: the spec puts the offset in *page*
    // coordinates, but downstream still adds `margins.left` back, so
    // we subtract the real margin here to keep the round-trip honest.
    let (page_width, margin_left, margin_right) = match frame {
        AnchorFrame::Page => (pc.page_size.width, pc.margins.left, pc.margins.right),
        AnchorFrame::Stack => (Pt::ZERO, Pt::ZERO, Pt::ZERO),
    };
    let content_width = (page_width - margin_left - margin_right).max(Pt::ZERO);

    // The compensation we apply for Page-relative anchors so the
    // post-shift value lands on the actual page coordinate. Zero for
    // `AnchorFrame::Page` (no shift happens) and `-pc.margins.left`
    // for `AnchorFrame::Stack`.
    let page_anchor_offset = match frame {
        AnchorFrame::Page => Pt::ZERO,
        AnchorFrame::Stack => -pc.margins.left,
    };

    match &anchor.horizontal_position {
        AnchorPosition::Offset {
            relative_from,
            offset,
        } => match relative_from {
            AnchorRelativeFrom::Page => page_anchor_offset + Pt::from(*offset),
            _ => margin_left + Pt::from(*offset),
        },
        AnchorPosition::Align {
            relative_from,
            alignment,
        } => {
            let (area_left, area_width) = match relative_from {
                AnchorRelativeFrom::Page => (page_anchor_offset, page_width.max(pc.page_size.width)),
                AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => {
                    (margin_left, content_width)
                }
                _ => (margin_left, content_width),
            };
            match alignment {
                AnchorAlignment::Left => area_left,
                AnchorAlignment::Right => area_left + area_width - content_w,
                AnchorAlignment::Center => area_left + (area_width - content_w) * 0.5,
                _ => area_left,
            }
        }
    }
}

/// Vertical axis of `resolve_anchor_position`. In `Stack` every offset is
/// paragraph-relative because the stacker — not the anchor — decides the
/// absolute page-y of the owning paragraph.
fn resolve_anchor_y(
    anchor: &crate::model::AnchorProperties,
    content_h: Pt,
    state: &BuildState,
    frame: AnchorFrame,
) -> FloatingImageY {
    use crate::model::{AnchorAlignment, AnchorPosition, AnchorRelativeFrom};

    let pc = &state.page_config;

    match &anchor.vertical_position {
        AnchorPosition::Offset {
            relative_from,
            offset,
        } => match frame {
            AnchorFrame::Stack => FloatingImageY::RelativeToParagraph(Pt::from(*offset)),
            AnchorFrame::Page => match relative_from {
                AnchorRelativeFrom::Page => FloatingImageY::Absolute(Pt::from(*offset)),
                AnchorRelativeFrom::Margin => {
                    FloatingImageY::Absolute(pc.margins.top + Pt::from(*offset))
                }
                // §20.4.2.11: topMargin — offset from page top.
                AnchorRelativeFrom::TopMargin => FloatingImageY::Absolute(Pt::from(*offset)),
                // §20.4.2.11: bottomMargin — offset from bottom margin edge.
                AnchorRelativeFrom::BottomMargin => FloatingImageY::Absolute(
                    pc.page_size.height - pc.margins.bottom + Pt::from(*offset),
                ),
                AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                    FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                }
                _ => FloatingImageY::Absolute(pc.margins.top + Pt::from(*offset)),
            },
        },
        AnchorPosition::Align {
            relative_from,
            alignment,
        } => {
            let (margin_top, page_height, margin_bottom) = match frame {
                AnchorFrame::Page => (pc.margins.top, pc.page_size.height, pc.margins.bottom),
                AnchorFrame::Stack => (Pt::ZERO, Pt::ZERO, Pt::ZERO),
            };
            let (area_top, area_height) = match relative_from {
                AnchorRelativeFrom::Page => (Pt::ZERO, page_height),
                AnchorRelativeFrom::Margin => (
                    margin_top,
                    (page_height - margin_top - margin_bottom).max(Pt::ZERO),
                ),
                // §20.4.2.11: topMargin = area from page top to top margin edge.
                AnchorRelativeFrom::TopMargin => (Pt::ZERO, margin_top),
                // §20.4.2.11: bottomMargin = area from bottom margin edge to page bottom.
                AnchorRelativeFrom::BottomMargin => (page_height - margin_bottom, margin_bottom),
                _ => (
                    margin_top,
                    (page_height - margin_top - margin_bottom).max(Pt::ZERO),
                ),
            };
            let y_pos = match alignment {
                AnchorAlignment::Top => area_top,
                AnchorAlignment::Bottom => area_top + area_height - content_h,
                AnchorAlignment::Center => area_top + (area_height - content_h) * 0.5,
                _ => area_top,
            };
            FloatingImageY::Absolute(y_pos)
        }
    }
}

// ── VML position helpers ───────────────────────────────────────────────────

/// Search an inline (and AlternateContent fallback) for a VML text box with
/// absolute positioning.
pub(super) fn find_vml_absolute_position(inline: &model::Inline) -> Option<(Pt, Pt)> {
    match inline {
        model::Inline::Pict(pict) => find_vml_pos_in_pict(pict),
        model::Inline::AlternateContent(ac) => {
            if let Some(ref fallback) = ac.fallback {
                for inner in fallback {
                    if let Some(pos) = find_vml_absolute_position(inner) {
                        return Some(pos);
                    }
                }
            }
            None
        }
        _ => None,
    }
}

fn find_vml_pos_in_pict(pict: &model::Pict) -> Option<(Pt, Pt)> {
    for shape in pict.shapes() {
        if shape.common.text_box.is_some() {
            if let Some(pos) = vml_absolute_position(&shape.common.style) {
                return Some(pos);
            }
        }
    }
    None
}

/// Extract absolute page-relative position from a VML shape style, in points.
fn vml_absolute_position(style: &model::VmlStyle) -> Option<(Pt, Pt)> {
    use crate::model::CssPosition;
    if style.position != Some(CssPosition::Absolute) {
        return None;
    }
    let x = style.margin_left.map(vml_length_to_pt)?;
    let y = style.margin_top.map(vml_length_to_pt)?;
    Some((x, y))
}

/// Convert a VML CSS length to points.
fn vml_length_to_pt(len: model::VmlLength) -> Pt {
    use crate::model::VmlLengthUnit;
    let value = len.value as f32;
    Pt::new(match len.unit {
        VmlLengthUnit::Pt => value,
        VmlLengthUnit::In => value * 72.0,
        VmlLengthUnit::Cm => value * 72.0 / 2.54,
        VmlLengthUnit::Mm => value * 72.0 / 25.4,
        VmlLengthUnit::Px => value * 0.75, // 96dpi → 72pt/in
        VmlLengthUnit::None => value / 914400.0 * 72.0, // bare number = EMU
        _ => value,                        // Em, Percent — fallback to raw value
    })
}
