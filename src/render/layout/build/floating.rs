//! Extract floating (anchor) images and shapes from paragraph inlines, and
//! resolve their positions in the caller-specified coordinate frame.

use crate::model::{self, Paragraph};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::layout::section::{
    FloatingImage, FloatingImageX, FloatingImageY, FloatingShape, PageParity,
};
use crate::render::resolve::shape_geometry::build_geometry;
use crate::render::resolve::shape_visuals::resolve_shape_visuals;

use super::convert::vml_style_length_to_pt;
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

/// Filter applied at extraction time so a shape's vertical anchor type can
/// route it through the correct frame.
///
/// Shapes whose vertical position is `paragraph`/`line` are paragraph-bound
/// (their absolute y is a function of the host paragraph's y), so they
/// travel on the owning paragraph through `stack_blocks`. Shapes with
/// `page`/`margin` vertical anchors resolve to a fixed page-y and must
/// bypass the stacker's per-paragraph anchoring — header/footer routes
/// them through a separate page-level vec, similar to how floating images
/// are split out.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum ShapeAnchorClass {
    /// Take every wsp shape regardless of its vertical anchor. Used by
    /// body paragraphs and table cells, where the legacy single-channel
    /// behaviour is still in effect.
    All,
    /// Only shapes whose vertical anchor is `paragraph` or `line`.
    ParagraphAnchored,
    /// Only shapes whose vertical anchor is `page` / `margin` /
    /// `topMargin` / `bottomMargin` / `insideMargin` / `outsideMargin`.
    PageAnchored,
}

/// Extract floating (anchor) images from a paragraph's inlines.
///
/// Positions are resolved in the coordinate system implied by `frame`
/// (see [`AnchorFrame`]).
use crate::render::layout::{live_mc_branch, McBranch};

fn find_anchor_images<'a>(
    inlines: &'a [crate::model::Inline],
    out: &mut Vec<&'a crate::model::Image>,
) {
    use crate::model::{GraphicContent, ImagePlacement, Inline};

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
            Inline::AlternateContent(ac) => match live_mc_branch(ac) {
                McBranch::Choices(choices) => {
                    for choice in choices {
                        find_anchor_images(&choice.content, out);
                    }
                }
                McBranch::Fallback(fallback) => find_anchor_images(fallback, out),
                McBranch::Neither => {}
            },
            _ => {}
        }
    }
}

pub(super) fn extract_floating_images(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &BuildState,
    frame: AnchorFrame,
) -> Vec<FloatingImage> {
    use crate::model::ImagePlacement;

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
            src_rect: crate::render::resolve::images::extract_src_rect(img),
            x,
            y,
            wrap_mode: crate::render::layout::section::WrapMode::from_model(&anchor.wrap),
            dist_left: Pt::from(anchor.distance.left),
            dist_right: Pt::from(anchor.distance.right),
            behind_doc: anchor.behind_text,
        });
    }

    // VML primitives that resolve to images (`<v:image>` or
    // `<v:shape type="#_x0000_t75">` carrying `<v:imagedata>`) ride
    // through the same `FloatingImage` channel as DrawingML images.
    extract_vml_floating_images(&para.content, state, frame, ctx, &mut images);

    images
}

// ── Floating shape extraction ──────────────────────────────────────────────

/// Extract floating (anchor) DrawingML shapes from a paragraph's inlines,
/// resolve their geometry + visuals, and compute their positions in the
/// coordinate frame implied by `frame`. Pure: takes immutable references to
/// `ctx` / `state`.
fn find_anchor_shapes<'a>(
    inlines: &'a [crate::model::Inline],
    out: &mut Vec<&'a crate::model::Image>,
) {
    use crate::model::{GraphicContent, ImagePlacement, Inline};

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
            // branch and the `<mc:Fallback>` carries the VML equivalent. The
            // branch is chosen by `live_mc_branch`, the same call the image
            // walker makes, so the two cannot disagree about which one the
            // paragraph contains.
            Inline::AlternateContent(ac) => match live_mc_branch(ac) {
                McBranch::Choices(choices) => {
                    for choice in choices {
                        find_anchor_shapes(&choice.content, out);
                    }
                }
                McBranch::Fallback(fallback) => find_anchor_shapes(fallback, out),
                McBranch::Neither => {}
            },
            _ => {}
        }
    }
}

pub(super) fn extract_floating_shapes(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &mut BuildState,
    frame: AnchorFrame,
    restrict: ShapeAnchorClass,
) -> Vec<FloatingShape> {
    use crate::model::{GraphicContent, ImagePlacement};

    let mut shape_imgs = Vec::new();
    find_anchor_shapes(&para.content, &mut shape_imgs);

    let mut shapes = Vec::new();
    for img in shape_imgs {
        let ImagePlacement::Anchor(ref anchor) = img.placement else {
            continue;
        };
        // Filter by anchor class so each call site only sees the shapes whose
        // vertical anchor matches the frame it's resolving in.
        let class_match = match restrict {
            ShapeAnchorClass::All => true,
            ShapeAnchorClass::ParagraphAnchored => anchors_to_paragraph(anchor),
            ShapeAnchorClass::PageAnchored => !anchors_to_paragraph(anchor),
        };
        if !class_match {
            continue;
        }
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
            wsp.style_fill_ref.as_ref(),
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

        // §17.17.1: lay out the shape's text-box content (`wps:txbx`) into
        // shape-local Pt commands. Both paragraph- and page-anchored shapes
        // benefit from the typed sub-layout — the consumer shifts the
        // commands by the shape's resolved origin (whether `RelativeToParagraph`
        // or `Absolute`), so text always lands on the shape's fill.
        let text_commands = build_shape_text_commands(wsp, extent, ctx, state);

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
            text_commands,
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
/// Walk inlines for VML primitives that resolve to images
/// (`<v:image>` or a `<v:shape type="#_x0000_t75">` whose only child
/// is `<v:imagedata>`). Both forms reach this code path through
/// `Inline::Pict.primitives`, including the `Inline::Pict` a live
/// `<mc:Fallback>` carries — see [`live_mc_branch`].
fn extract_vml_floating_images(
    inlines: &[crate::model::Inline],
    state: &BuildState,
    frame: AnchorFrame,
    ctx: &BuildContext,
    out: &mut Vec<FloatingImage>,
) {
    use crate::model::Inline;
    for inline in inlines {
        match inline {
            Inline::Pict(pict) => {
                for primitive in &pict.primitives {
                    extract_vml_primitive_image(primitive, state, frame, ctx, out);
                }
            }
            Inline::Hyperlink(link) => {
                extract_vml_floating_images(&link.content, state, frame, ctx, out)
            }
            Inline::Field(f) => extract_vml_floating_images(&f.content, state, frame, ctx, out),
            // §M.1.2: exactly one branch is live and [`live_mc_branch`] is the
            // only thing that decides which. This arm used to be `{}` — a
            // third answer, given by not answering — so a Choice we could not
            // draw plus a Fallback holding a VML image left the anchor with no
            // geometry at all while its text still reached the page.
            Inline::AlternateContent(ac) => match live_mc_branch(ac) {
                McBranch::Choices(choices) => {
                    for choice in choices {
                        extract_vml_floating_images(&choice.content, state, frame, ctx, out);
                    }
                }
                McBranch::Fallback(fallback) => {
                    extract_vml_floating_images(fallback, state, frame, ctx, out)
                }
                McBranch::Neither => {}
            },
            _ => {}
        }
    }
}

fn extract_vml_primitive_image(
    primitive: &model::VmlPrimitive,
    state: &BuildState,
    frame: AnchorFrame,
    ctx: &BuildContext,
    out: &mut Vec<FloatingImage>,
) {
    use crate::model::VmlPrimitive;
    match primitive {
        VmlPrimitive::Image(img) => {
            if let Some(fi) = build_vml_floating_image(&img.common, state, frame, ctx) {
                out.push(fi);
            }
        }
        // §14.1.2.19 — `<v:shape type="#_x0000_t75"><v:imagedata r:id=…/>`
        // is the standard pre-DrawingML way to embed an image. We treat
        // any VmlShape whose `image_data` is set and whose `text_box` is
        // empty as image-bearing; shapes that *also* host text are handled
        // elsewhere via the inline fragment collector.
        VmlPrimitive::Shape(s) if s.common.image_data.is_some() && s.common.text_box.is_none() => {
            if let Some(fi) = build_vml_floating_image(&s.common, state, frame, ctx) {
                out.push(fi);
            }
        }
        VmlPrimitive::Group(g) => {
            for child in &g.children {
                extract_vml_primitive_image(child, state, frame, ctx, out);
            }
        }
        _ => {}
    }
}

fn build_vml_floating_image(
    common: &model::VmlCommonAttrs,
    state: &BuildState,
    frame: AnchorFrame,
    ctx: &BuildContext,
) -> Option<FloatingImage> {
    use crate::render::layout::section::WrapMode;

    let rel_id = common.image_data.as_ref()?.rel_id.as_ref()?;
    let image_data = ctx.resolved.media.get(rel_id).cloned()?;

    let (page_x, y) = vml_absolute_position(&common.style)?;
    // VML has no `inside`/`outside` equivalent — `margin-left` is one number.
    let x = FloatingImageX::Absolute(match frame {
        AnchorFrame::Page => page_x,
        AnchorFrame::Stack => page_x - state.page_config.margins.left,
    });

    let width = common.style.width.and_then(vml_style_length_to_pt)?;
    let height = common.style.height.and_then(vml_style_length_to_pt)?;
    if width <= Pt::ZERO || height <= Pt::ZERO {
        return None;
    }

    Some(FloatingImage {
        image_data,
        size: PtSize::new(width, height),
        // VML crops (`v:imagedata` crop*) are a separate mechanism; not modelled.
        src_rect: None,
        x,
        y: FloatingImageY::RelativeToParagraph(y),
        wrap_mode: WrapMode::None,
        dist_left: Pt::ZERO,
        dist_right: Pt::ZERO,
        behind_doc: false,
    })
}

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
            // §M.1.2: `<mc:AlternateContent>` carries the same shape twice —
            // modern DrawingML in `<mc:Choice>`, a VML fallback in
            // `<mc:Fallback>` — so drawing both emits one rectangle twice.
            // Skipping the element outright avoided that and created the
            // opposite bug: a Choice we cannot draw left the Fallback's rect
            // unrendered, which is the "future case" the note here used to
            // anticipate. [`live_mc_branch`] is now that Choice-fed signal,
            // and the same one `find_anchor_shapes` asks.
            Inline::AlternateContent(ac) => match live_mc_branch(ac) {
                McBranch::Choices(choices) => {
                    for choice in choices {
                        extract_vml_primitive_shapes(&choice.content, state, frame, out);
                    }
                }
                McBranch::Fallback(fallback) => {
                    extract_vml_primitive_shapes(fallback, state, frame, out)
                }
                McBranch::Neither => {}
            },
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
        // §14.1.2.17 `<v:roundrect>`: rounded-corner variant. The
        // corner radius is `arcsize * min(width, height) / 2` —
        // small radii barely distinguish from a plain rect at typical
        // doc resolutions, so for Tier 0 we render as a plain
        // rectangle. The `arcsize` value survives in the model for a
        // future rounded-path build.
        VmlPrimitive::RoundRect(r) => {
            if let Some(shape) = build_vml_rect_shape(&r.common, state, frame) {
                out.push(shape);
            }
        }
        // §14.1.2.9: groups carry their own coord system; we walk
        // children so any absolute-positioned descendants reach the
        // page. The local-coord transform (`coordsize`/`coordorigin`
        // → page coords) for relative-positioned children is left
        // for a future phase — most authored groups in headers and
        // footers use absolute positioning and don't depend on it.
        VmlPrimitive::Group(g) => {
            for child in &g.children {
                extract_vml_primitive(child, state, frame, out);
            }
        }
        // Image variants don't produce a `FloatingShape` — they go
        // through the parallel `extract_vml_floating_images` path
        // and emit `FloatingImage` instead.
        VmlPrimitive::Image(_) => {}
        // Long-tail primitives modeled in Phase A but not yet
        // emitted as shapes (oval/line/polyline/arc/curve). Phase D
        // will dispatch them to `DrawCommand::Path` / `Line`. Their
        // text-box content (where applicable) is still picked up by
        // the inline fragment collector.
        VmlPrimitive::Shape(_)
        | VmlPrimitive::Oval(_)
        | VmlPrimitive::Line(_)
        | VmlPrimitive::PolyLine(_)
        | VmlPrimitive::Arc(_)
        | VmlPrimitive::Curve(_) => {}
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
    // VML has no `inside`/`outside` equivalent — `margin-left` is one number.
    let x = FloatingImageX::Absolute(match frame {
        AnchorFrame::Page => page_x,
        AnchorFrame::Stack => page_x - state.page_config.margins.left,
    });

    // Size via `style.width` / `style.height`. A rect with no extent
    // can't meaningfully render.
    let width = common.style.width.and_then(vml_style_length_to_pt)?;
    let height = common.style.height.and_then(vml_style_length_to_pt)?;
    if width <= Pt::ZERO || height <= Pt::ZERO {
        return None;
    }
    let extent = PtSize::new(width, height);

    // Fill resolution per §14.1.2.5: a `<v:fill>` child overrides
    // `@fillcolor`. We honor solid fills natively and degrade
    // gradient/tile/pattern/frame to `ResolvedFill::None` with a
    // one-time log so the rest of the shape still renders.
    let fill = resolve_vml_solid_fill(common);

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
        // VML rect text-box content is still picked up at the host paragraph
        // y by the inline-fragment collector. Sub-layout into shape-local
        // commands isn't wired for VML primitives yet.
        text_commands: Vec::new(),
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
) -> (FloatingImageX, FloatingImageY) {
    let x = resolve_anchor_x(anchor, content_w, state, frame);
    let y = resolve_anchor_y(anchor, content_h, state, frame);
    (x, y)
}

/// §20.4.3.4 `ST_RelFromH` / §20.4.3.5 `ST_RelFromV`: a strip of the sheet
/// along one axis, expressed in the coordinates of the [`AnchorFrame`] that
/// produced it.
#[derive(Clone, Copy, Debug, PartialEq)]
struct AnchorSpan {
    /// The strip's near edge — x of its left edge horizontally, y of its top
    /// edge vertically.
    start: Pt,
    /// Its extent along that axis. Zero is legal and meaningful: a page with no
    /// left margin has an empty `leftMargin` strip, and in
    /// [`AnchorFrame::Stack`] the container's extent is simply not known here.
    extent: Pt,
}

/// The region an anchor measures against — or, for the two references whose
/// region depends on which page the object lands on, the pair it chooses from.
///
/// Shared by both axes: §20.4.3.4 and §20.4.3.5 mirror the same way, one about
/// left and right and the other about top and bottom.
#[derive(Clone, Copy, Debug, PartialEq)]
enum AnchorRegion {
    /// The same strip on every page.
    Fixed(AnchorSpan),
    /// §20.4.3.4 / §20.4.3.5 `insideMargin` / `outsideMargin`: "inside" is the
    /// left margin on an odd (recto) page and the right margin on an even one —
    /// and, vertically, the top margin on an odd page and the bottom on an even
    /// one. Which strip this names is a function of the page number.
    Mirrored { odd: AnchorSpan, even: AnchorSpan },
}

impl AnchorRegion {
    /// The strip this reference names on a page of the given parity.
    fn on(self, parity: PageParity) -> AnchorSpan {
        match self {
            Self::Fixed(span) => span,
            Self::Mirrored { odd, even } => match parity {
                PageParity::Odd => odd,
                PageParity::Even => even,
            },
        }
    }
}

/// The page as an [`AnchorFrame`] measures it — the input to
/// [`horizontal_region`].
///
/// §20.4.3.4's references split in two, and that split is what makes `Stack`
/// tractable:
///
/// * **Page-derived** — `page` and the four margin strips are pure functions of
///   the sheet and its margins, so they are knowable in either frame;
///   `page_left` is what carries them into frame coordinates.
/// * **Container-derived** — `margin` and `column` name the area the object's
///   *container* gives it. In `Page` that is the sheet's text area. In `Stack`
///   the container is a table cell or a header whose extent never reaches this
///   function, so the region collapses onto the frame origin. That collapse is
///   pre-existing and deliberate (see [`AnchorFrame`]); giving it a real extent
///   would move every header float.
#[derive(Clone, Copy, Debug)]
struct FrameGeometry {
    /// x of the sheet's left edge. Zero in `Page`; `-margins.left` in `Stack`,
    /// where the frame origin is the body's left margin and the caller adds
    /// that margin back on the way into page coordinates.
    page_left: Pt,
    /// The sheet's own width, always real — the page exists in either frame.
    page_width: Pt,
    /// The sheet's own margins, likewise always real.
    margin_left: Pt,
    margin_right: Pt,
    /// The container region, already in frame coordinates.
    container: AnchorSpan,
}

impl FrameGeometry {
    fn new(pc: &crate::render::layout::page::PageConfig, frame: AnchorFrame) -> Self {
        let (margin_left, margin_right) = (pc.margins.left, pc.margins.right);
        let page_width = pc.page_size.width;
        let (page_left, container) = match frame {
            AnchorFrame::Page => (
                Pt::ZERO,
                AnchorSpan {
                    start: margin_left,
                    extent: (page_width - margin_left - margin_right).max(Pt::ZERO),
                },
            ),
            AnchorFrame::Stack => (
                -margin_left,
                AnchorSpan {
                    start: Pt::ZERO,
                    extent: Pt::ZERO,
                },
            ),
        };
        Self {
            page_left,
            page_width,
            margin_left,
            margin_right,
            container,
        }
    }

    /// §20.4.3.4 `page` — the whole sheet, margins included.
    fn page(&self) -> AnchorSpan {
        AnchorSpan {
            start: self.page_left,
            extent: self.page_width,
        }
    }

    /// §20.4.3.4 `leftMargin` — the strip from the sheet's left edge to the
    /// left margin edge. The margin *itself*, not the text area beside it.
    fn left_margin(&self) -> AnchorSpan {
        AnchorSpan {
            start: self.page_left,
            extent: self.margin_left,
        }
    }

    /// §20.4.3.4 `rightMargin` — the mirror strip at the other edge.
    fn right_margin(&self) -> AnchorSpan {
        AnchorSpan {
            start: self.page_left + self.page_width - self.margin_right,
            extent: self.margin_right,
        }
    }
}

/// §20.4.3.4 `ST_RelFromH`: the region `from` names.
///
/// Total over `AnchorRelativeFrom` on purpose — there is deliberately no
/// catch-all, so a new spec variant becomes a build error rather than another
/// silent landing in the text area, which is how four distinct margin strips
/// came to share one arm in the first place.
fn horizontal_region(from: crate::model::AnchorRelativeFrom, geom: &FrameGeometry) -> AnchorRegion {
    use crate::model::AnchorRelativeFrom as From;
    use AnchorRegion::{Fixed, Mirrored};

    match from {
        From::Page => Fixed(geom.page()),
        // §20.4.3.4 `column` is the text column. With one column that is the
        // text area; a multi-column section anchors its floats per column, and
        // no float takes that path today.
        From::Margin | From::Column => Fixed(geom.container),
        From::LeftMargin => Fixed(geom.left_margin()),
        From::RightMargin => Fixed(geom.right_margin()),
        From::InsideMargin => Mirrored {
            odd: geom.left_margin(),
            even: geom.right_margin(),
        },
        From::OutsideMargin => Mirrored {
            odd: geom.right_margin(),
            even: geom.left_margin(),
        },
        // §20.4.3.4 `character` is not a region at all: it is the anchor's own
        // position within the text run. Floats are extracted before that run is
        // laid out, so the position does not exist here. Fall back to the text
        // area — and say so, because a silent fallback is how this went
        // unnoticed while sharing an arm with `margin`.
        From::Character => {
            log::warn!(
                "anchor: relativeFrom=\"character\" needs the anchor's position in \
                 the run, which float extraction runs before — positioning \
                 against the text area instead"
            );
            Fixed(geom.container)
        }
        // §20.4.3.5-only values. One `AnchorRelativeFrom` serves both axes, so
        // these are reachable only from a document that put a vertical
        // reference on `wp:positionH`.
        From::Paragraph | From::Line | From::TopMargin | From::BottomMargin => {
            log::warn!(
                "anchor: relativeFrom={from:?} is not a horizontal reference \
                 (§20.4.3.4) — positioning against the text area instead"
            );
            Fixed(geom.container)
        }
    }
}

/// §20.4.3.5 `ST_RelFromV`: the vertical strip `from` names.
///
/// Only ever asked in [`AnchorFrame::Page`] — a `Stack`-framed anchor has no
/// resolved container to measure against and returns before reaching here — so
/// this takes the page directly rather than a [`FrameGeometry`].
///
/// Total over `AnchorRelativeFrom` for the same reason `horizontal_region` is:
/// a new spec variant must be a build error, not another silent landing in the
/// margin box.
fn vertical_region(
    from: crate::model::AnchorRelativeFrom,
    pc: &crate::render::layout::page::PageConfig,
) -> AnchorRegion {
    use crate::model::AnchorRelativeFrom as From;
    use AnchorRegion::{Fixed, Mirrored};

    let (margin_top, margin_bottom) = (pc.margins.top, pc.margins.bottom);
    let page_height = pc.page_size.height;
    let margin_box = AnchorSpan {
        start: margin_top,
        extent: (page_height - margin_top - margin_bottom).max(Pt::ZERO),
    };
    // §20.4.2.11: the strip from the sheet's top edge to the top margin edge —
    // the margin itself, not the text area below it. `bottomMargin` mirrors it.
    let top_margin = AnchorSpan {
        start: Pt::ZERO,
        extent: margin_top,
    };
    let bottom_margin = AnchorSpan {
        start: page_height - margin_bottom,
        extent: margin_bottom,
    };

    match from {
        From::Page => Fixed(AnchorSpan {
            start: Pt::ZERO,
            extent: page_height,
        }),
        From::Margin => Fixed(margin_box),
        From::TopMargin => Fixed(top_margin),
        From::BottomMargin => Fixed(bottom_margin),
        // §20.4.3.5 `insideMargin`/`outsideMargin`. Settled against a Word
        // render of `test-files/issue-165-floatv.docx` (issue #165): vertically
        // "inside" is the *top* margin strip on an odd (recto) page and the
        // *bottom* one on an even page — the top/bottom analogue of the
        // left/right mirror `horizontal_region` applies. See the test module's
        // §20.4.3.2 section for the six-page table this comes from.
        From::InsideMargin => Mirrored {
            odd: top_margin,
            even: bottom_margin,
        },
        From::OutsideMargin => Mirrored {
            odd: bottom_margin,
            even: top_margin,
        },
        // §20.4.3.5 `paragraph`/`line` name an area the stacker has not placed
        // yet. An *offset* against them is carried as `RelativeToParagraph` and
        // never reaches here; an *alignment* has nothing to align within, so it
        // collapses onto the margin box.
        From::Paragraph | From::Line => Fixed(margin_box),
        // §20.4.3.4-only values on `wp:positionV` — one `AnchorRelativeFrom`
        // serves both axes, so these are reachable only from a document that
        // put a horizontal reference on the vertical position.
        From::Column | From::Character | From::LeftMargin | From::RightMargin => {
            log::warn!(
                "anchor: relativeFrom={from:?} is not a vertical reference \
                 (§20.4.3.5) — positioning against the margin box"
            );
            Fixed(margin_box)
        }
    }
}

/// Horizontal axis of `resolve_anchor_position`. Split out so the two axes
/// can be read independently.
///
/// Parity reaches the result through *two* channels — the region
/// (`insideMargin`/`outsideMargin`) and the alignment (`inside`/`outside`) —
/// so rather than thread it through the arithmetic, the whole position is
/// evaluated once per parity and the two readings handed to
/// [`FloatingImageX::from_pages`]. An anchor that uses neither channel produces
/// two equal readings and collapses back to `Absolute`, which is why a
/// single-sided document carries no deferral at all.
fn resolve_anchor_x(
    anchor: &crate::model::AnchorProperties,
    content_w: Pt,
    state: &BuildState,
    frame: AnchorFrame,
) -> FloatingImageX {
    use crate::model::{AnchorAlignment, AnchorPosition};

    let geom = FrameGeometry::new(&state.page_config, frame);

    let at = |parity: PageParity| -> Pt {
        match &anchor.horizontal_position {
            // §20.4.2.12: an offset is measured from the region's own left edge.
            AnchorPosition::Offset {
                relative_from,
                offset,
            } => horizontal_region(*relative_from, &geom).on(parity).start + Pt::from(*offset),
            // §20.4.3.1: an alignment places the object within the region.
            AnchorPosition::Align {
                relative_from,
                alignment,
            } => {
                let span = horizontal_region(*relative_from, &geom).on(parity);
                let near = span.start;
                let far = span.start + span.extent - content_w;
                match alignment {
                    AnchorAlignment::Left => near,
                    AnchorAlignment::Right => far,
                    AnchorAlignment::Center => span.start + (span.extent - content_w) * 0.5,
                    // §20.4.3.1: "inside" is the binding edge — left on an odd
                    // (recto) page, right on an even one — and "outside" is the
                    // trimmed edge opposite it.
                    AnchorAlignment::Inside => match parity {
                        PageParity::Odd => near,
                        PageParity::Even => far,
                    },
                    AnchorAlignment::Outside => match parity {
                        PageParity::Odd => far,
                        PageParity::Even => near,
                    },
                    // §20.4.3.2 vertical alignments on `wp:positionH`. One
                    // `AnchorAlignment` serves both axes, so these are
                    // reachable only from a malformed document.
                    AnchorAlignment::Top | AnchorAlignment::Bottom => {
                        log::warn!(
                            "anchor: align={alignment:?} is not a horizontal alignment \
                             (§20.4.3.1) — placing at the region's left edge instead"
                        );
                        near
                    }
                }
            }
        }
    };

    FloatingImageX::from_pages(at(PageParity::Odd), at(PageParity::Even))
}

/// Vertical axis of `resolve_anchor_position`. In `Stack` every offset is
/// paragraph-relative because the stacker — not the anchor — decides the
/// absolute page-y of the owning paragraph.
///
/// Parity reaches the result through the same two channels as on the
/// horizontal axis — the region (`insideMargin`/`outsideMargin`) and the
/// alignment (`inside`/`outside`) — so the position is evaluated once per
/// parity and the two readings handed to
/// [`FloatingImageY::absolute_from_pages`], which collapses them back to
/// `Absolute` when they agree. They agree for every anchor that is not
/// `inside`/`outside`.
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
                // §20.4.3.5 `paragraph`/`line`: the stacker, not the anchor,
                // knows where the owning paragraph lands.
                AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                    FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                }
                // §20.4.2.11: every other reference is a strip of the sheet,
                // and the offset is measured from that strip's own top edge.
                // Listed rather than caught so a new spec variant is a build
                // error here as well as in `vertical_region`.
                AnchorRelativeFrom::Page
                | AnchorRelativeFrom::Margin
                | AnchorRelativeFrom::TopMargin
                | AnchorRelativeFrom::BottomMargin
                | AnchorRelativeFrom::InsideMargin
                | AnchorRelativeFrom::OutsideMargin
                | AnchorRelativeFrom::Column
                | AnchorRelativeFrom::Character
                | AnchorRelativeFrom::LeftMargin
                | AnchorRelativeFrom::RightMargin => {
                    let region = vertical_region(*relative_from, pc);
                    let at = |parity: PageParity| region.on(parity).start + Pt::from(*offset);
                    FloatingImageY::absolute_from_pages(at(PageParity::Odd), at(PageParity::Even))
                }
            },
        },
        AnchorPosition::Align {
            relative_from,
            alignment,
        } => {
            // `Stack` has no resolved container extent at extraction time —
            // the stacker decides the frame origin later — so there is no
            // area to align within. Collapse to the paragraph origin, the
            // same convention the `Offset` arm above uses. Emitting an
            // `Absolute` page coordinate here would be doubly wrong: the
            // caller shifts stack coordinates into page space afterwards,
            // and with the frame's margins zeroed `Bottom`/`Center` resolve
            // to negative y — off the top of the page.
            if frame == AnchorFrame::Stack {
                return FloatingImageY::RelativeToParagraph(Pt::ZERO);
            }
            let region = vertical_region(*relative_from, pc);
            let at = |parity: PageParity| -> Pt {
                let span = region.on(parity);
                let near = span.start;
                let far = span.start + span.extent - content_h;
                match alignment {
                    AnchorAlignment::Top => near,
                    AnchorAlignment::Bottom => far,
                    AnchorAlignment::Center => span.start + (span.extent - content_h) * 0.5,
                    // §20.4.3.2 `inside`/`outside`: vertically, "inside" is the
                    // region's top edge on an odd (recto) page and its bottom
                    // edge on an even one, "outside" the opposite. Measured
                    // from a Word render — see the test module's §20.4.3.2
                    // section for the six-page table.
                    AnchorAlignment::Inside => match parity {
                        PageParity::Odd => near,
                        PageParity::Even => far,
                    },
                    AnchorAlignment::Outside => match parity {
                        PageParity::Odd => far,
                        PageParity::Even => near,
                    },
                    // §20.4.3.1 horizontal alignments on `wp:positionV` —
                    // malformed, the same way round as the horizontal axis.
                    AnchorAlignment::Left | AnchorAlignment::Right => {
                        log::warn!(
                            "anchor: align={alignment:?} is not a vertical alignment \
                             (§20.4.3.2) — aligning to the region's top instead"
                        );
                        near
                    }
                }
            };
            FloatingImageY::absolute_from_pages(at(PageParity::Odd), at(PageParity::Even))
        }
    }
}

// ── VML position helpers ───────────────────────────────────────────────────

/// Search an inline (and AlternateContent fallback) for a VML text box with
/// absolute positioning.
pub(super) fn find_vml_absolute_position(inline: &model::Inline) -> Option<(Pt, Pt)> {
    match inline {
        model::Inline::Pict(pict) => find_vml_pos_in_pict(pict),
        // Only the VML fallback carries an absolute position, and it is
        // meaningful only when the fallback is the branch we render. A
        // drawable Choice makes it inert, and consuming it anyway turns the
        // fallback's origin into the header's and pushes paragraph-anchored
        // content off-page.
        model::Inline::AlternateContent(ac) => match crate::render::layout::live_mc_branch(ac) {
            crate::render::layout::McBranch::Fallback(fallback) => {
                fallback.iter().find_map(find_vml_absolute_position)
            }
            crate::render::layout::McBranch::Choices(_)
            | crate::render::layout::McBranch::Neither => None,
        },
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
    let x = style.margin_left.and_then(vml_style_length_to_pt)?;
    let y = style.margin_top.and_then(vml_style_length_to_pt)?;
    Some((x, y))
}

/// Compute the effective solid `ResolvedFill` for a VML primitive
/// per §14.1.2.5. The `<v:fill>` child wins over `@fillcolor`. Only
/// the `Solid` fill type is honored here; the others log once and
/// degrade to `ResolvedFill::None` so the shape's outline / text
/// content still renders.
fn resolve_vml_solid_fill(
    common: &model::VmlCommonAttrs,
) -> crate::render::layout::draw_command::ResolvedFill {
    use crate::model::{VmlColor, VmlFillType};
    use crate::render::layout::draw_command::ResolvedFill;
    use crate::render::resolve::drawing_color::Rgba;

    let to_solid = |c: &VmlColor| -> Option<ResolvedFill> {
        match c {
            VmlColor::Rgb(r, g, b) => Some(ResolvedFill::Solid(Rgba {
                r: *r as f32 / 255.0,
                g: *g as f32 / 255.0,
                b: *b as f32 / 255.0,
                a: 1.0,
            })),
            // Named/system colors aren't yet resolved — fall through.
            VmlColor::Named(_) => None,
        }
    };

    if let Some(ref fill) = common.fill {
        match fill.fill_type {
            VmlFillType::Solid => {
                if let Some(c) = fill.color.as_ref().and_then(to_solid) {
                    return c;
                }
                // Solid type with no `@color` — fall back to attribute.
            }
            VmlFillType::Gradient
            | VmlFillType::GradientRadial
            | VmlFillType::Tile
            | VmlFillType::Frame
            | VmlFillType::Pattern => {
                log::warn!(
                    "vml: unsupported fill type {:?} — rendering as no-fill",
                    fill.fill_type
                );
                return ResolvedFill::None;
            }
        }
    }

    common
        .fill_color
        .as_ref()
        .and_then(to_solid)
        .unwrap_or(ResolvedFill::None)
}

/// True when the anchor's vertical position resolves relative to the host
/// paragraph or line — the only case where the shape's eventual page-y is a
/// function of `paragraph_top` rather than an absolute frame. The shape-text
/// sub-layout assumes paragraph-relative placement (the stacker ties
/// `text_commands` to the same `(fs.x, shape_y)` it uses for the path), so
/// page-anchored shapes get their text emitted via the legacy inline-fragment
/// collector instead. Tier 1 follow-up: extend the sub-layout to honor
/// page/margin frames so all wsp shapes can use the typed path.
fn anchors_to_paragraph(anchor: &crate::model::AnchorProperties) -> bool {
    use crate::model::{AnchorPosition, AnchorRelativeFrom};
    let relative_from = match &anchor.vertical_position {
        AnchorPosition::Offset { relative_from, .. } => relative_from,
        AnchorPosition::Align { relative_from, .. } => relative_from,
    };
    matches!(
        relative_from,
        AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line
    )
}

/// §20.1.10.60 `ST_TextAnchoringType`: where a shape's text body sits inside
/// the box its insets leave.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum BodyAnchor {
    Top,
    Center,
    Bottom,
}

impl BodyAnchor {
    /// Total over `TextAnchoringType`, with the §20.1.2.1.1 default for an
    /// absent attribute.
    fn resolve(anchor: Option<crate::model::TextAnchoringType>) -> Self {
        use crate::model::TextAnchoringType as T;
        match anchor {
            None | Some(T::Top) => Self::Top,
            Some(T::Center) => Self::Center,
            Some(T::Bottom) => Self::Bottom,
            // §20.1.10.60 `just`/`dist` stretch the *inter-line* spacing to
            // fill the box, which this sub-layout has no line-level control
            // over. Degrading to `Top` is the closest honest reading: a
            // justified body also begins at the top, it simply is not
            // stretched to reach the bottom.
            Some(anchor @ (T::Justified | T::Distributed)) => {
                log::warn!(
                    "shape text: anchor={anchor:?} distributes lines to fill the body \
                     (§20.1.10.60), which is not modelled — anchoring to the top instead"
                );
                Self::Top
            }
        }
    }

    /// How far below the top inset a `text_height`-tall body sits in a
    /// `box_height`-tall box.
    ///
    /// The slack is floored at zero, so a body taller than its box anchors to
    /// the top and overflows *downward* whatever the attribute says.
    /// `@vertOverflow` defaults to `overflow` — Word draws overflowing shape
    /// text rather than clipping it — and centring a body that does not fit
    /// would put its first lines above the shape, over whatever sits there. So
    /// the choice here is only ever about where the spare room goes. A body
    /// that asks for `clip` is trimmed afterwards, in `overflow_keeps`, not by
    /// moving it.
    fn offset(self, box_height: Pt, text_height: Pt) -> Pt {
        let slack = (box_height - text_height).max(Pt::ZERO);
        match self {
            Self::Top => Pt::ZERO,
            Self::Center => slack * 0.5,
            Self::Bottom => slack,
        }
    }
}

/// §17.17.1 / §20.1.2.1.1: lay out a shape's `wps:txbx/w:txbxContent` into
/// shape-local draw commands. Output is in shape-local Pt with origin at the
/// shape's top-left; the stacker shifts by `(fs.x, shape_y)` when emitting,
/// so the text appears inside the shape's bounding box on top of the fill.
///
/// Body insets (§20.1.2.1.1 lIns/tIns/rIns/bIns) deflate the box the body is
/// laid out in, and §20.1.10.60 `anchor` places it within that box. The
/// function also runs a fresh `BuildState` for list/footnote counters so the
/// shape's interior doesn't pollute the outer document; `field_ctx` is copied
/// so PAGE/NUMPAGES inside a shape's text body still resolve against the host
/// page.
pub(super) fn build_shape_text_commands(
    wsp: &crate::model::WordProcessingShape,
    extent: PtSize,
    ctx: &BuildContext,
    state: &BuildState,
) -> Vec<crate::render::layout::draw_command::DrawCommand> {
    if wsp.txbx_content.is_empty() {
        return Vec::new();
    }

    // §20.1.2.1.1 spec defaults: 91440 EMU horizontal, 45720 EMU vertical.
    let default_lr = Pt::new(91440.0 / 12700.0); // ≈ 7.2pt
    let default_tb = Pt::new(45720.0 / 12700.0); // ≈ 3.6pt
    let (left_inset, top_inset, right_inset, bot_inset) =
        wsp.body_pr
            .as_ref()
            .map_or((default_lr, default_tb, default_lr, default_tb), |bp| {
                (
                    bp.left_inset.map_or(default_lr, Pt::from),
                    bp.top_inset.map_or(default_tb, Pt::from),
                    bp.right_inset.map_or(default_lr, Pt::from),
                    bp.bottom_inset.map_or(default_tb, Pt::from),
                )
            });

    let content_width = (extent.width - left_inset - right_inset).max(Pt::ZERO);
    if content_width <= Pt::ZERO {
        return Vec::new();
    }

    // §20.1.4.1.17: a wps:style/fontRef sets the shape's default text color and
    // theme font collection for the text-box content. Resolved here and threaded
    // through BuildState so the paragraph cascade uses them as the base (an
    // explicit run/style color still wins).
    let theme = ctx.resolved.theme.as_ref();
    let (shape_default_text_color, shape_default_font_family) = match &wsp.style_font_ref {
        Some(fr) => {
            let color = fr.color.as_ref().map(|c| {
                let dc = crate::render::resolve::drawing_color::DrawingColorContext::new(theme);
                let rgba = crate::render::resolve::drawing_color::resolve_drawing_color(c, &dc);
                crate::render::resolve::color::rgb_from_u32(rgba.to_rgb24())
            });
            let family = theme.and_then(|t| {
                let fam = match fr.collection {
                    crate::model::FontCollectionIndex::Major => t.major_font.latin.clone(),
                    crate::model::FontCollectionIndex::Minor => t.minor_font.latin.clone(),
                    crate::model::FontCollectionIndex::None => String::new(),
                };
                (!fam.is_empty()).then_some(fam)
            });
            (color, family)
        }
        None => (None, None),
    };

    // §20.1.2.1.18: the shrink Word already computed for this body. Read once
    // here and carried on the sub-state, so every size the cascade resolves
    // inside the body — runs, list labels, field substitutions, blank lines —
    // goes through it.
    let auto_fit = crate::render::layout::ShapeAutoFit::from_body(
        wsp.body_pr.as_ref().and_then(|bp| bp.auto_fit),
    );

    // Sub-state with the host's page dimensions and field context. Counters
    // are reset so a footnote/list inside a shape body doesn't bump the
    // outer counters.
    let mut sub_state = BuildState {
        // §17.3.1.19: a heading inside a text box is not a position in the
        // document's main story — see `OutlineCollector`.
        outline: crate::render::layout::build::OutlineCollector::Excluded,
        shape_auto_fit: auto_fit,
        page_config: state.page_config.clone(),
        footnotes: Default::default(),
        endnote_counter: 0,
        list_counters: std::collections::HashMap::new(),
        field_ctx: state.field_ctx,
        shape_default_text_color,
        shape_default_font_family,
        // Its own set: this sub-state is built from a `&BuildState`, so the
        // parent's set can't be borrowed mutably here. A border style used
        // only inside a shape text box is therefore reported once per shape
        // rather than once per render — bounded over-reporting in a rare case,
        // preferred over making `state` mutable through ten signatures.
        warned_border_styles: std::collections::HashSet::new(),
        warned_row_cell_spacing: false,
        warned_orphan_vmerge: false,
    };

    let hf = super::build_header_footer_content(&wsp.txbx_content, ctx, &mut sub_state);
    // §20.1.2.1.18: the body's own shrink also applies to the fallback line
    // height, which is what an empty paragraph and an image-only line fall back
    // to — otherwise a shrunk body would keep full-size blank lines.
    let line_height = auto_fit.scale_font(super::default_line_height(ctx));
    // Shape text is laid out at *build* time, before the shape is placed on a
    // page, so a §20.4.3.1 `inside`/`outside` float nested inside a shape's
    // text box has no parity to resolve against and takes the odd-page
    // reading. The same structural limit as the table-cell path in
    // `layout_cell`, and rarer still.
    let result = crate::render::layout::section::stack_blocks(
        &hf.blocks,
        content_width,
        line_height,
        None,
        PageParity::Odd,
    );

    // §20.1.10.60: `bIns` closes off the bottom of the box the body sits in,
    // and `anchor` decides where in that box it sits. Both were previously
    // dropped, which pinned every body to the top.
    let content_height = (extent.height - top_inset - bot_inset).max(Pt::ZERO);
    let anchor = BodyAnchor::resolve(wsp.body_pr.as_ref().and_then(|bp| bp.anchor));
    let body_top = top_inset + anchor.offset(content_height, result.height);

    // `@vertOverflow` decides what happens to the part of the body that does
    // not fit. `Overflow` — the spec default, and the only value the corpus
    // asks for — keeps everything, so the common path is untouched.
    let overflow = wsp
        .body_pr
        .as_ref()
        .and_then(|bp| bp.vert_overflow)
        .unwrap_or_default();
    let box_bottom = top_inset + content_height;

    let mut commands = Vec::with_capacity(result.commands.len());
    for mut cmd in result.commands {
        cmd.shift(left_inset, body_top);
        if !overflow_keeps(overflow, &cmd, box_bottom) {
            continue;
        }
        commands.push(cmd);
    }
    commands
}

/// Whether `@vertOverflow` keeps `cmd`, given the bottom of the body's box.
///
/// Total over [`TextVertOverflow`] with no catch-all, so a new value of the
/// attribute has to state its own behaviour here.
///
/// **This drops whole commands, which is a line-granular approximation of what
/// Word does.** Word clips at the pixel, so a line straddling the box edge
/// shows its top sliver; here it disappears. Real clipping needs a canvas clip
/// that survives into paint, and draw commands are flattened into one flat
/// per-page list with no scoping — so it would mean a new `DrawCommand`
/// wrapper variant and an arm in every consumer. Dropping is the safe
/// direction (`clip`'s contract is that nothing paints outside the box), and
/// no corpus document asks for `clip` at all — 1 explicit `overflow`, 12
/// `bodyPr` with the attribute absent, zero `clip` or `ellipsis` — so this is
/// worth revisiting only once a real document needs the sliver.
fn overflow_keeps(
    overflow: crate::model::TextVertOverflow,
    cmd: &crate::render::layout::draw_command::DrawCommand,
    box_bottom: Pt,
) -> bool {
    use crate::model::TextVertOverflow;

    match overflow {
        TextVertOverflow::Overflow => true,
        // `ellipsis` is `clip` plus an indicator on the last visible line.
        // Choosing that line and refitting it around the ellipsis glyph is a
        // decision this sub-layout does not make, so the indicator is dropped
        // and the clipping is honoured — the same text as `clip`, which is far
        // closer to Word than not clipping at all.
        TextVertOverflow::Clip | TextVertOverflow::Ellipsis => cmd
            .vertical_span()
            .is_none_or(|(_, bottom)| bottom <= box_bottom),
    }
}

#[cfg(test)]
mod tests {
    use super::find_vml_absolute_position;
    use crate::model::dimension::Dimension;
    use crate::model::geometry::{EdgeInsets, Size};
    use crate::model::{
        AlternateContent, AnchorPosition, AnchorProperties, AnchorRelativeFrom, DocProperties,
        GraphicContent, Image, ImagePlacement, Inline, McChoice, McRequires, TextWrap,
        WordProcessingShape,
    };
    use crate::render::layout::{live_mc_branch, McBranch};

    /// A minimally-populated anchored `wps:wsp` shape (as `Inline::Image`).
    fn anchored_wps_image() -> Image {
        Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "shape".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: Some(GraphicContent::WordProcessingShape(WordProcessingShape {
                cnv_pr: None,
                shape_properties: None,
                style_line_ref: None,
                style_effect_ref: None,
                style_fill_ref: None,
                style_font_ref: None,
                body_pr: None,
                txbx_content: vec![],
            })),
            placement: ImagePlacement::Anchor(AnchorProperties {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
                simple_pos: None,
                use_simple_pos: None,
                horizontal_position: AnchorPosition::Offset {
                    relative_from: AnchorRelativeFrom::Margin,
                    offset: Dimension::new(0),
                },
                vertical_position: AnchorPosition::Offset {
                    relative_from: AnchorRelativeFrom::Paragraph,
                    offset: Dimension::new(0),
                },
                wrap: TextWrap::None,
                behind_text: false,
                lock_anchor: false,
                allow_overlap: true,
                relative_height: 0,
                layout_in_cell: None,
                hidden: None,
            }),
        }
    }

    fn ac_with_wps_choice() -> AlternateContent {
        AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::Image(Box::new(anchored_wps_image()))],
            }],
            // A non-empty fallback that would otherwise be searched.
            fallback: Some(vec![Inline::InstrText(String::new())]),
        }
    }

    #[test]
    fn an_anchored_shape_in_a_choice_makes_that_choice_live() {
        assert!(matches!(
            live_mc_branch(&ac_with_wps_choice()),
            McBranch::Choices(_)
        ));
    }

    /// A Choice whose `Requires` we meet but whose content yields no anchored
    /// object is not drawable. Liveness is a question about content: widening
    /// it to a namespace check would light up this Choice and strand a
    /// Fallback that is the only branch with anything to draw.
    #[test]
    fn a_choice_with_nothing_anchored_yields_to_the_fallback() {
        let ac = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::InstrText(String::new())],
            }],
            fallback: Some(vec![Inline::InstrText(String::new())]),
        };
        assert!(matches!(live_mc_branch(&ac), McBranch::Fallback(_)));
    }

    #[test]
    fn no_drawable_choice_and_no_fallback_is_neither() {
        let ac = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::InstrText(String::new())],
            }],
            fallback: None,
        };
        assert!(matches!(live_mc_branch(&ac), McBranch::Neither));
    }

    /// An anchored *picture* is as live as an anchored shape. The suppression
    /// sites used to ask a narrower question — "does a Choice hold a `wps:wsp`
    /// shape?" — so a picture Choice was drawn as a float while its Fallback's
    /// text was *also* collected inline: two branches of one element on one
    /// page, which §M.1.2 does not allow.
    #[test]
    fn an_anchored_picture_in_a_choice_is_just_as_live_as_a_shape() {
        let ac = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wpg],
                content: vec![Inline::Image(Box::new(anchored_picture()))],
            }],
            fallback: Some(vec![Inline::InstrText(String::new())]),
        };
        assert!(matches!(live_mc_branch(&ac), McBranch::Choices(_)));
    }

    /// §M.1.2's content model for a branch is `drawing | pict`, so this shape
    /// cannot come from a document — but the recursion has to be right anyway,
    /// because every walker reads the same answer and an outer element that
    /// disagreed with its inner one would put the two on different branches.
    #[test]
    fn a_nested_alternate_content_resolves_innermost_first() {
        let nested = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::Image(Box::new(anchored_wps_image()))],
            }],
            fallback: None,
        };
        let outer = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::AlternateContent(nested)],
            }],
            fallback: Some(vec![Inline::InstrText(String::new())]),
        };
        assert!(matches!(live_mc_branch(&outer), McBranch::Choices(_)));
    }

    /// …and the converse: an inner element that draws nothing leaves the outer
    /// Choice undrawable, so the outer Fallback goes live.
    #[test]
    fn a_nested_alternate_content_that_draws_nothing_frees_the_outer_fallback() {
        let nested = AlternateContent {
            choices: vec![],
            fallback: Some(vec![Inline::InstrText(String::new())]),
        };
        let outer = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::AlternateContent(nested)],
            }],
            fallback: Some(vec![Inline::InstrText(String::new())]),
        };
        assert!(matches!(live_mc_branch(&outer), McBranch::Fallback(_)));
    }

    /// A `<w:pict>` holding one absolutely-positioned VML text box at
    /// (100pt, 40pt) — a position `find_vml_absolute_position` will find, so a
    /// test that expects `None` fails for the right reason rather than because
    /// there was nothing to find.
    fn positioned_vml_text_box() -> Inline {
        use crate::model::{
            CssPosition, Pict, VmlCommonAttrs, VmlLength, VmlLengthUnit, VmlPrimitive, VmlShape,
            VmlStyle, VmlTextBox,
        };

        let pt = |value| {
            Some(VmlLength {
                value,
                unit: VmlLengthUnit::Pt,
            })
        };
        Inline::Pict(Pict {
            shape_type: None,
            primitives: vec![VmlPrimitive::Shape(VmlShape {
                common: VmlCommonAttrs {
                    style: VmlStyle {
                        position: Some(CssPosition::Absolute),
                        margin_left: pt(100.0),
                        margin_top: pt(40.0),
                        ..VmlStyle::default()
                    },
                    text_box: Some(VmlTextBox {
                        style: VmlStyle::default(),
                        inset: None,
                        content: vec![],
                    }),
                    ..VmlCommonAttrs::default()
                },
                shape_type_ref: None,
                vml_path: None,
            })],
        })
    }

    /// The probe finds a position when the pict is what the paragraph holds —
    /// the precondition every suppression assertion below rests on.
    #[test]
    fn a_positioned_vml_text_box_has_an_absolute_position() {
        assert!(find_vml_absolute_position(&positioned_vml_text_box()).is_some());
    }

    /// Regression: when the Choice is a DrawingML shape we render, the VML
    /// fallback's absolute position must be ignored — otherwise it hijacks
    /// the header origin and pushes paragraph-anchored content off-page.
    #[test]
    fn a_drawable_choice_suppresses_its_fallbacks_absolute_position() {
        let mut ac = ac_with_wps_choice();
        ac.fallback = Some(vec![positioned_vml_text_box()]);
        assert!(find_vml_absolute_position(&Inline::AlternateContent(ac)).is_none());
    }

    /// …and when no Choice is drawable the fallback *is* the document, so its
    /// position is the one to use. Suppressing here would leave a legacy
    /// VML-only header with no origin at all.
    #[test]
    fn a_live_fallbacks_absolute_position_is_the_one_that_counts() {
        let ac = AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wps],
                content: vec![Inline::InstrText(String::new())],
            }],
            fallback: Some(vec![positioned_vml_text_box()]),
        };
        assert!(find_vml_absolute_position(&Inline::AlternateContent(ac)).is_some());
    }

    // ── §20.4.2.10 vertical anchor resolution ────────────────────────────

    use super::{resolve_anchor_y, AnchorFrame};
    use crate::model::AnchorAlignment;
    use crate::render::dimension::Pt;
    use crate::render::layout::build::BuildState;
    use crate::render::layout::section::FloatingImageY;
    use crate::render::layout::section::{FloatingImageX, PageParity};

    fn default_state() -> BuildState {
        BuildState {
            page_config: Default::default(),
            outline: Default::default(),
            shape_auto_fit: crate::render::layout::ShapeAutoFit::NONE,
            footnotes: Default::default(),
            endnote_counter: 0,
            list_counters: Default::default(),
            field_ctx: Default::default(),
            warned_border_styles: Default::default(),
            warned_row_cell_spacing: false,
            warned_orphan_vmerge: false,
            shape_default_text_color: None,
            shape_default_font_family: None,
        }
    }

    fn anchor_with_v(vertical_position: AnchorPosition) -> AnchorProperties {
        let ImagePlacement::Anchor(mut a) = anchored_wps_image().placement else {
            unreachable!("fixture is anchored")
        };
        a.vertical_position = vertical_position;
        a
    }

    fn v_align(alignment: AnchorAlignment) -> AnchorProperties {
        anchor_with_v(AnchorPosition::Align {
            relative_from: AnchorRelativeFrom::Margin,
            alignment,
        })
    }

    /// Regression: the `Align` arm used to ignore `AnchorFrame`, returning an
    /// `Absolute` *page* coordinate for stack-framed anchors. With the frame's
    /// margins zeroed that put `Center`/`Bottom` at negative y — so an anchored
    /// float in a table cell, header, or footer using `<wp:align>` rendered at
    /// or above the top of the page instead of next to its paragraph.
    #[test]
    fn stack_frame_align_is_paragraph_relative() {
        let state = default_state();
        for alignment in [
            AnchorAlignment::Top,
            AnchorAlignment::Center,
            AnchorAlignment::Bottom,
        ] {
            let y = resolve_anchor_y(
                &v_align(alignment),
                Pt::new(50.0),
                &state,
                AnchorFrame::Stack,
            );
            let FloatingImageY::RelativeToParagraph(offset) = y else {
                panic!("{alignment:?} in Stack frame must be paragraph-relative");
            };
            assert_eq!(offset, Pt::ZERO, "{alignment:?} collapses to the paragraph");
        }
    }

    /// The `Offset` arm already honored the frame — pinned so the two arms
    /// can't drift apart again.
    #[test]
    fn stack_frame_offset_is_paragraph_relative() {
        let anchor = anchor_with_v(AnchorPosition::Offset {
            relative_from: AnchorRelativeFrom::Margin,
            offset: Dimension::new(914400), // 1 inch in EMU
        });
        let y = resolve_anchor_y(&anchor, Pt::new(50.0), &default_state(), AnchorFrame::Stack);
        let FloatingImageY::RelativeToParagraph(offset) = y else {
            panic!("Offset in Stack frame must be paragraph-relative");
        };
        assert!((offset.raw() - 72.0).abs() < 1e-3, "1in = 72pt");
    }

    /// `Page` frame still resolves alignment against the real margin box.
    /// Default page: 792pt tall, 72pt margins → content area 72..720.
    #[test]
    fn page_frame_align_resolves_against_margin_box() {
        let state = default_state();
        let content_h = Pt::new(50.0);
        let cases = [
            (AnchorAlignment::Top, 72.0),
            (AnchorAlignment::Center, 72.0 + (648.0 - 50.0) * 0.5),
            (AnchorAlignment::Bottom, 72.0 + 648.0 - 50.0),
        ];
        for (alignment, expected) in cases {
            let y = resolve_anchor_y(&v_align(alignment), content_h, &state, AnchorFrame::Page);
            let FloatingImageY::Absolute(got) = y else {
                panic!("{alignment:?} in Page frame must be absolute");
            };
            assert!(
                (got.raw() - expected).abs() < 1e-3,
                "{alignment:?}: expected {expected}, got {}",
                got.raw()
            );
        }
    }

    // ── §20.4.3.2 / §20.4.3.5 vertical inside/outside ────────────────────
    //
    // Settled by a Word render of `test-files/issue-165-floatv.docx` (issue
    // #165), a two-sided (`w:mirrorMargins`) document whose six pages each
    // carry one anchored 36pt image, on a 612×792pt sheet with a deliberately
    // asymmetric 72pt top / 144pt bottom margin so a top/bottom mirror is
    // visible at all. Word placed them:
    //
    // | page | anchor                       | Word            |
    // |------|------------------------------|-----------------|
    // | 1 odd  | `margin` + `align=inside`  | top             |
    // | 2 even | `margin` + `align=inside`  | bottom          |
    // | 3 odd  | `insideMargin` + offset 0  | glued to page top |
    // | 4 even | `insideMargin` + offset 0  | below the bottom margin |
    // | 5 odd  | `margin` + `align=outside` | bottom (like 2) |
    // | 6 even | `margin` + `align=outside` | top (like 1)    |
    //
    // So vertically `inside` is the **top** on an odd page and the **bottom**
    // on an even one, and `outside` is the complement — the same page-parity
    // mirror the horizontal axis applies to left and right. That falsifies the
    // reading these arms carried until now ("a two-sided document mirrors left
    // and right, not top and bottom"), which was a guess made in the absence of
    // a render.
    //
    // One qualification on page 4, the only observation that is not exact: it
    // fixes the *strip* (below the bottom margin) but not the position within
    // it. Page 3 is what settles that — "glued to the page top", i.e. offset 0
    // measured from the strip's own start, not centred in it — and the two
    // strips are read the same way, as `topMargin`/`bottomMargin` already are.
    // A render that put page 4's object anywhere but flush against the bottom
    // margin edge would contradict page 3, so this is derived rather than
    // assumed; a measurement of that one page would confirm it outright.
    //
    // The tests below use the default page (612×792, 72pt margins all round)
    // rather than the fixture's geometry; the fixture itself is checked
    // end-to-end in `tests/floating_anchor_parity.rs`.

    fn v_align_from(
        relative_from: AnchorRelativeFrom,
        alignment: AnchorAlignment,
    ) -> AnchorProperties {
        anchor_with_v(AnchorPosition::Align {
            relative_from,
            alignment,
        })
    }

    fn v_offset(relative_from: AnchorRelativeFrom, offset: i64) -> AnchorProperties {
        anchor_with_v(AnchorPosition::Offset {
            relative_from,
            offset: Dimension::new(offset),
        })
    }

    /// `resolve_anchor_y` for a 50pt-tall object on a `Page`-framed page of
    /// `parity`.
    fn y_on(anchor: &AnchorProperties, parity: PageParity) -> f32 {
        resolve_anchor_y(anchor, Pt::new(50.0), &default_state(), AnchorFrame::Page)
            .at(parity, Pt::ZERO)
            .raw()
    }

    fn assert_y(got: f32, expected: f32, what: &str) {
        assert!(
            (got - expected).abs() < 1e-3,
            "{what}: expected {expected}, got {got}"
        );
    }

    /// §20.4.3.2: `inside` is the top of the region on an odd page and its
    /// bottom on an even one; `outside` is the opposite edge. Pages 1/2 and
    /// 5/6 of the Word render above.
    #[test]
    fn inside_and_outside_alignments_mirror_top_and_bottom() {
        // Margin box 72..720; a 50pt object is flush top at 72, flush bottom
        // at 670.
        for (alignment, odd, even) in [
            (AnchorAlignment::Inside, 72.0, 670.0),
            (AnchorAlignment::Outside, 670.0, 72.0),
        ] {
            let anchor = v_align_from(AnchorRelativeFrom::Margin, alignment);
            assert_y(
                y_on(&anchor, PageParity::Odd),
                odd,
                &format!("{alignment:?} on an odd page"),
            );
            assert_y(
                y_on(&anchor, PageParity::Even),
                even,
                &format!("{alignment:?} on an even page"),
            );
        }
    }

    /// §20.4.3.5: vertically, `insideMargin` names the **top** margin strip on
    /// an odd page and the **bottom** one on an even page, `outsideMargin` the
    /// complement — and an offset is measured from the named strip's own start,
    /// the way `topMargin`/`bottomMargin` already are. Pages 3/4 of the render:
    /// odd puts the object at the very top of the sheet, even puts it just
    /// below the bottom margin edge.
    #[test]
    fn inside_and_outside_margin_references_mirror_top_and_bottom() {
        for (from, odd, even) in [
            (AnchorRelativeFrom::InsideMargin, 0.0, 720.0),
            (AnchorRelativeFrom::OutsideMargin, 720.0, 0.0),
        ] {
            let anchor = v_offset(from, 0);
            assert_y(
                y_on(&anchor, PageParity::Odd),
                odd,
                &format!("{from:?} odd"),
            );
            assert_y(
                y_on(&anchor, PageParity::Even),
                even,
                &format!("{from:?} even"),
            );
        }
    }

    /// Both parity channels compose on this axis too: an `inside` alignment
    /// *within* an `insideMargin` region mirrors through region and alignment
    /// at once, and must not double-mirror back to the odd-page answer. The
    /// horizontal twin is `a_mirrored_alignment_inside_a_mirrored_region_mirrors_once`.
    #[test]
    fn a_mirrored_vertical_alignment_inside_a_mirrored_region_mirrors_once() {
        let anchor = v_align_from(AnchorRelativeFrom::InsideMargin, AnchorAlignment::Inside);
        // Odd: insideMargin = the 0..72 strip, inside-aligned = its top edge.
        assert_y(y_on(&anchor, PageParity::Odd), 0.0, "odd");
        // Even: insideMargin = the 720..792 strip, inside-aligned = its
        // *bottom* edge, less the object's height.
        assert_y(y_on(&anchor, PageParity::Even), 720.0 + 72.0 - 50.0, "even");
    }

    /// An anchor that uses neither parity channel collapses to `Absolute`, so
    /// an ordinary document carries no vertical deferral at all — the property
    /// that keeps this ADT free on the documents that are nearly all of them.
    #[test]
    fn an_unmirrored_vertical_anchor_carries_no_deferral() {
        let cases: [(&str, AnchorProperties); 5] = [
            ("align=top", v_align(AnchorAlignment::Top)),
            ("align=center", v_align(AnchorAlignment::Center)),
            ("align=bottom", v_align(AnchorAlignment::Bottom)),
            (
                "offset from margin",
                v_offset(AnchorRelativeFrom::Margin, 914400),
            ),
            (
                "offset from topMargin",
                v_offset(AnchorRelativeFrom::TopMargin, 0),
            ),
        ];
        for (what, anchor) in cases {
            let y = resolve_anchor_y(&anchor, Pt::new(50.0), &default_state(), AnchorFrame::Page);
            assert!(
                matches!(y, FloatingImageY::Absolute(_)),
                "{what} is parity-independent, got {y:?}"
            );
        }
        for (what, anchor) in [
            (
                "align=inside",
                v_align_from(AnchorRelativeFrom::Margin, AnchorAlignment::Inside),
            ),
            (
                "offset from insideMargin",
                v_offset(AnchorRelativeFrom::InsideMargin, 0),
            ),
        ] {
            let y = resolve_anchor_y(&anchor, Pt::new(50.0), &default_state(), AnchorFrame::Page);
            assert!(
                matches!(y, FloatingImageY::PageParity { .. }),
                "{what} is parity-dependent, got {y:?}"
            );
        }
    }

    /// The `Stack` frame has no page to mirror against — a header, footer or
    /// table cell float stays paragraph-relative whichever way it is anchored.
    /// Pinned because the mirror is applied in the same two arms that
    /// `stack_frame_align_is_paragraph_relative` guards.
    #[test]
    fn stack_frame_never_mirrors() {
        for anchor in [
            v_align_from(AnchorRelativeFrom::Margin, AnchorAlignment::Inside),
            v_align_from(AnchorRelativeFrom::InsideMargin, AnchorAlignment::Outside),
            v_offset(AnchorRelativeFrom::OutsideMargin, 0),
        ] {
            let y = resolve_anchor_y(&anchor, Pt::new(50.0), &default_state(), AnchorFrame::Stack);
            assert!(
                matches!(y, FloatingImageY::RelativeToParagraph(_)),
                "Stack frame must not mirror, got {y:?}"
            );
        }
    }

    // ── §20.1.10.60 shape text-body anchoring ────────────────────────────

    use super::build_shape_text_commands;
    use crate::model::{
        BodyProperties, Paragraph as ModelParagraph, ParagraphProperties, RunElement,
        RunProperties, TextAnchoringType, TextRun,
    };
    use crate::render::fonts::FontRegistry;
    use crate::render::geometry::PtSize;
    use crate::render::layout::build::BuildContext;
    use crate::render::layout::measurer::TextMeasurer;
    use crate::render::resolve::ResolvedDocument;

    fn empty_resolved() -> ResolvedDocument {
        use std::collections::HashMap;
        ResolvedDocument {
            sections: Vec::new(),
            styles: HashMap::new(),
            numbering: HashMap::new(),
            font_families: Vec::new(),
            media: HashMap::new(),
            embedded_fonts: Vec::new(),
            pic_bullets: HashMap::new(),
            theme: None,
            doc_defaults_paragraph: ParagraphProperties::default(),
            doc_defaults_run: RunProperties::default(),
            default_paragraph_style_id: None,
            default_table_style_id: None,
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            even_and_odd_headers: false,
            default_tab_stop: Dimension::new(720),
            show_ins_del_marks: true,
            show_comment_marks: true,
            revision_colors: Default::default(),
            comments: Default::default(),
        }
    }

    /// A shape whose text body is one short line, with the given `bodyPr`.
    fn wsp_with_text(body_pr: Option<BodyProperties>) -> WordProcessingShape {
        WordProcessingShape {
            cnv_pr: None,
            shape_properties: None,
            style_line_ref: None,
            style_effect_ref: None,
            style_fill_ref: None,
            style_font_ref: None,
            body_pr,
            txbx_content: vec![crate::model::Block::Paragraph(Box::new(ModelParagraph {
                style_id: None,
                properties: ParagraphProperties::default(),
                mark_run_properties: None,
                content: vec![Inline::TextRun(Box::new(TextRun {
                    style_id: None,
                    properties: RunProperties::default(),
                    content: vec![RunElement::Text("hi".into())],
                    rsids: crate::model::RevisionIds::default(),
                    revision: None,
                    comment: None,
                }))],
                rsids: crate::model::ParagraphRevisionIds::default(),
                mark_deleted: false,
            }))],
        }
    }

    fn body_pr(anchor: Option<TextAnchoringType>, inset_emu: i64) -> BodyProperties {
        BodyProperties {
            rotation: None,
            vert_overflow: None,
            vert: None,
            wrap: None,
            left_inset: Some(Dimension::new(inset_emu)),
            top_inset: Some(Dimension::new(inset_emu)),
            right_inset: Some(Dimension::new(inset_emu)),
            bottom_inset: Some(Dimension::new(inset_emu)),
            anchor,
            auto_fit: None,
        }
    }

    /// The y of the first emitted text command, in shape-local points.
    fn shape_text_y(wsp: &WordProcessingShape, extent: PtSize) -> f32 {
        let resolved = empty_resolved();
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let state = BuildState::default();
        let commands = build_shape_text_commands(wsp, extent, &ctx, &state);
        commands
            .iter()
            .find_map(|c| match c {
                crate::render::layout::draw_command::DrawCommand::Text { position, .. } => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .expect("the shape body emits text")
    }

    /// §20.1.10.60: `anchor` places the body within the box `bIns` closes off.
    /// `t` pins it under the top inset, `ctr` centres it in the box, `b` sits
    /// it on the bottom inset.
    ///
    /// Every assertion is a *difference* between two runs. The emitted y is a
    /// baseline, so its absolute value carries Skia's real ascent for whatever
    /// font the host resolves — differencing cancels it out and leaves only the
    /// anchoring this test is about.
    #[test]
    fn body_anchor_places_text_within_the_inset_box() {
        // 50800 EMU = 4pt on every side of a 200x120pt shape, so the text box
        // is 112pt tall.
        const INSET: f32 = 4.0;
        const BOX_HEIGHT: f32 = 120.0 - 2.0 * INSET;
        let extent = PtSize::new(Pt::new(200.0), Pt::new(120.0));
        let y = |anchor| shape_text_y(&wsp_with_text(Some(body_pr(Some(anchor), 50800))), extent);
        let (top, centre, bottom) = (
            y(TextAnchoringType::Top),
            y(TextAnchoringType::Center),
            y(TextAnchoringType::Bottom),
        );

        assert!(
            top < centre && centre < bottom,
            "t < ctr < b, got {top} / {centre} / {bottom}"
        );
        // Bottom-anchoring pushes the body down by the whole slack, centring by
        // half of it — exactly a 1:2 ratio, whatever the line height is.
        let (half, full) = (centre - top, bottom - top);
        assert!(
            (full - 2.0 * half).abs() < 1e-3,
            "the centre offset is half the bottom offset, got {half} / {full}"
        );
        // And the slack is the box height less the one line in it, so the line
        // height falls out of `full` — a real number to check the box against.
        let line_height = BOX_HEIGHT - full;
        assert!(
            line_height > 0.0 && line_height < BOX_HEIGHT,
            "one line fits inside the 112pt box, implied height {line_height}"
        );
    }

    /// The top inset shifts the body one-for-one — which is what says the box
    /// is measured from it, rather than the anchoring happening to land there.
    #[test]
    fn the_top_inset_shifts_a_top_anchored_body_one_for_one() {
        let extent = PtSize::new(Pt::new(200.0), Pt::new(120.0));
        let anchor = Some(TextAnchoringType::Top);
        // 0 EMU vs 50800 EMU (4pt).
        let flush = shape_text_y(&wsp_with_text(Some(body_pr(anchor, 0))), extent);
        let inset = shape_text_y(&wsp_with_text(Some(body_pr(anchor, 50800))), extent);
        assert!(
            (inset - flush - 4.0).abs() < 1e-3,
            "a 4pt top inset moves the body 4pt down, got {flush} → {inset}"
        );
    }

    /// The bottom inset shifts a bottom-anchored body one-for-one — which is
    /// what says `bIns` closes off the box, rather than the anchoring reaching
    /// the shape's own bottom edge. The 1:2 centre/bottom ratio above holds
    /// either way, so nothing else here would notice `bIns` being dropped.
    #[test]
    fn the_bottom_inset_shifts_a_bottom_anchored_body_one_for_one() {
        let extent = PtSize::new(Pt::new(200.0), Pt::new(120.0));
        let anchor = Some(TextAnchoringType::Bottom);
        let flush = shape_text_y(&wsp_with_text(Some(body_pr(anchor, 0))), extent);
        let inset = shape_text_y(&wsp_with_text(Some(body_pr(anchor, 50800))), extent);
        // 4pt of bottom inset lifts the body 4pt; the 4pt of *top* inset the
        // same fixture adds cannot reach a bottom-anchored body at all.
        assert!(
            (flush - inset - 4.0).abs() < 1e-3,
            "a 4pt bottom inset lifts the body 4pt, got {flush} → {inset}"
        );
    }

    /// No `bodyPr` at all keeps the §20.1.2.1.1 defaults — top anchoring under
    /// a 45720 EMU top inset. This is the behaviour every shape had before
    /// anchoring existed, so nothing that omits `anchor` moves.
    #[test]
    fn a_shape_without_body_properties_keeps_the_spec_defaults() {
        let extent = PtSize::new(Pt::new(200.0), Pt::new(120.0));
        let bare = shape_text_y(&wsp_with_text(None), extent);
        let explicit = shape_text_y(
            &wsp_with_text(Some(body_pr(Some(TextAnchoringType::Top), 45720))),
            extent,
        );
        assert!(
            (bare - explicit).abs() < 1e-3,
            "an absent bodyPr matches the spec defaults spelled out, \
             got {bare} vs {explicit}"
        );
    }

    /// `@vertOverflow` defaults to `overflow` — and every `bodyPr` in the
    /// corpus that names it says so explicitly — so a body taller than its box
    /// is *not* clipped. It anchors to the top and overflows downward, which is
    /// also what it did before anchoring existed: the change reaches only
    /// bodies that fit. Centring an overflowing body would draw it above the
    /// shape, over whatever sits there.
    #[test]
    fn an_overflowing_body_is_not_pushed_above_the_shape() {
        // A 6pt-tall shape cannot hold a line of text at any font size.
        let extent = PtSize::new(Pt::new(200.0), Pt::new(6.0));
        let y = |anchor| shape_text_y(&wsp_with_text(Some(body_pr(Some(anchor), 50800))), extent);
        let top = y(TextAnchoringType::Top);
        for anchor in [TextAnchoringType::Center, TextAnchoringType::Bottom] {
            assert!(
                (y(anchor) - top).abs() < 1e-3,
                "{anchor:?} on an overflowing body places as `t` does, got {} vs {top}",
                y(anchor)
            );
        }
    }

    // ── MCE §M.1.2 branch selection ──────────────────────────────────────

    use super::{find_anchor_images, find_anchor_shapes};
    use crate::model::{Blip, BlipFill, BlipFillKind, NvPicProperties, Picture, RelId};

    /// An anchored DrawingML *picture* — the other half of what an anchor can
    /// hold, and the half `find_anchor_images` owns.
    fn anchored_picture() -> Image {
        let mut img = anchored_wps_image();
        img.graphic = Some(GraphicContent::Picture(Picture {
            nv_pic_pr: NvPicProperties {
                cnv_pr: DocProperties {
                    id: 2,
                    name: "picture".into(),
                    description: None,
                    hidden: None,
                    title: None,
                },
                cnv_pic_pr: None,
            },
            blip_fill: BlipFill {
                rotate_with_shape: None,
                dpi: None,
                blip: Some(Blip {
                    embed: Some(RelId::new("rId7")),
                    link: None,
                    compression: None,
                }),
                src_rect: None,
                fill_kind: BlipFillKind::Unspecified,
            },
            shape_properties: None,
        }));
        img
    }

    /// `<mc:AlternateContent>` whose Choice holds `choice` and whose Fallback
    /// holds `fallback`.
    fn ac_of(choice: Vec<Inline>, fallback: Vec<Inline>) -> Inline {
        Inline::AlternateContent(AlternateContent {
            choices: vec![McChoice {
                requires: vec![McRequires::Wpg],
                content: choice,
            }],
            fallback: Some(fallback),
        })
    }

    /// MCE §M.1.2 selects exactly one branch, so the two walkers — two halves
    /// of one DrawingML extraction, split by graphic type — have to select the
    /// same one. Here the Choice holds a picture and the Fallback a shape: ask
    /// the image walker and the paragraph contains nothing, ask the shape
    /// walker and it contains a shape from the *other* branch.
    #[test]
    fn both_anchor_walkers_read_the_same_alternate_content_branch() {
        let content = vec![ac_of(
            vec![Inline::Image(Box::new(anchored_picture()))],
            vec![Inline::Image(Box::new(anchored_wps_image()))],
        )];

        let mut images = Vec::new();
        find_anchor_images(&content, &mut images);
        let mut shapes = Vec::new();
        find_anchor_shapes(&content, &mut shapes);

        assert_eq!(images.len(), 1, "the Choice's picture is live");
        assert_eq!(
            shapes.len(),
            0,
            "the Fallback's shape is not — the Choice was selected"
        );
    }

    /// The dominant shape of the element in the wild: a DrawingML Choice and a
    /// VML Fallback. The Choice's picture must render; today the image walker
    /// only ever looks in the Fallback, and the VML walkers skip
    /// `AlternateContent` outright, so the picture reaches neither and the
    /// anchor draws nothing at all.
    #[test]
    fn an_anchored_picture_in_a_choice_is_found() {
        let content = vec![ac_of(
            vec![Inline::Image(Box::new(anchored_picture()))],
            vec![Inline::InstrText(String::new())],
        )];

        let mut images = Vec::new();
        find_anchor_images(&content, &mut images);
        assert_eq!(images.len(), 1, "the Choice's anchored picture is rendered");
    }

    /// With no renderable Choice the Fallback is live, for both walkers. This
    /// is the branch the image walker used to take unconditionally, and it has
    /// to keep working — a legacy VML-only anchor has nowhere else to come
    /// from.
    #[test]
    fn an_unrenderable_choice_hands_the_document_to_the_fallback() {
        let content = vec![ac_of(
            vec![Inline::InstrText(String::new())],
            vec![Inline::Image(Box::new(anchored_picture()))],
        )];

        let mut images = Vec::new();
        find_anchor_images(&content, &mut images);
        assert_eq!(images.len(), 1, "the Fallback's picture is live");
    }

    /// Regression guard for the case the shape walker already handled: a wps
    /// Choice wins and its VML Fallback stays inert, so the same rectangle is
    /// not drawn twice.
    #[test]
    fn a_wps_choice_still_suppresses_its_vml_fallback() {
        let content = vec![ac_of(
            vec![Inline::Image(Box::new(anchored_wps_image()))],
            vec![Inline::Image(Box::new(anchored_picture()))],
        )];

        let mut images = Vec::new();
        find_anchor_images(&content, &mut images);
        let mut shapes = Vec::new();
        find_anchor_shapes(&content, &mut shapes);

        assert_eq!(shapes.len(), 1, "the Choice's shape renders");
        assert_eq!(images.len(), 0, "the Fallback's picture does not");
    }

    // ── §20.4.3.4 horizontal anchor resolution ───────────────────────────
    //
    // The default page is US Letter with 1in margins: 612 x 792pt, text area
    // 72..540. Every expectation below is written against those numbers.

    use super::resolve_anchor_x;

    /// One inch in EMU — the unit `wp:posOffset` is expressed in.
    const INCH: i64 = 914400;

    fn anchor_with_h(horizontal_position: AnchorPosition) -> AnchorProperties {
        let ImagePlacement::Anchor(mut a) = anchored_wps_image().placement else {
            unreachable!("fixture is anchored")
        };
        a.horizontal_position = horizontal_position;
        a
    }

    fn h_offset(relative_from: AnchorRelativeFrom, offset: i64) -> AnchorProperties {
        anchor_with_h(AnchorPosition::Offset {
            relative_from,
            offset: Dimension::new(offset),
        })
    }

    fn h_align(relative_from: AnchorRelativeFrom, alignment: AnchorAlignment) -> AnchorProperties {
        anchor_with_h(AnchorPosition::Align {
            relative_from,
            alignment,
        })
    }

    /// `resolve_anchor_x` for a 100pt-wide object on a page of `parity`.
    fn x_on(anchor: &AnchorProperties, frame: AnchorFrame, parity: PageParity) -> f32 {
        resolve_anchor_x(anchor, Pt::new(100.0), &default_state(), frame)
            .resolve(parity)
            .raw()
    }

    /// The odd-page reading — which is *the* reading for every anchor that is
    /// not `inside`/`outside`.
    fn x_of(anchor: &AnchorProperties, frame: AnchorFrame) -> f32 {
        x_on(anchor, frame, PageParity::Odd)
    }

    fn assert_x(got: f32, expected: f32, what: &str) {
        assert!(
            (got - expected).abs() < 1e-3,
            "{what}: expected {expected}, got {got}"
        );
    }

    /// §20.4.2.12: a `page`-relative offset is a page coordinate, so in the
    /// page frame it passes through untouched.
    #[test]
    fn page_frame_page_relative_offset_is_a_page_coordinate() {
        assert_x(
            x_of(&h_offset(AnchorRelativeFrom::Page, INCH), AnchorFrame::Page),
            72.0,
            "1in from the page's left edge",
        );
    }

    /// A `margin`-relative offset is measured from the text area's left edge.
    #[test]
    fn page_frame_margin_relative_offset_starts_at_the_margin() {
        assert_x(
            x_of(
                &h_offset(AnchorRelativeFrom::Margin, INCH),
                AnchorFrame::Page,
            ),
            144.0,
            "1in past the 1in left margin",
        );
    }

    /// In `Stack` the caller adds the body's left margin back, so a page
    /// coordinate has to be pre-compensated or the round-trip double-counts.
    #[test]
    fn stack_frame_page_relative_offset_backs_out_the_left_margin() {
        assert_x(
            x_of(
                &h_offset(AnchorRelativeFrom::Page, INCH),
                AnchorFrame::Stack,
            ),
            0.0,
            "72pt page coordinate minus the 72pt margin the caller re-adds",
        );
    }

    /// A margin-relative offset needs no compensation: the frame origin
    /// already *is* the left margin.
    #[test]
    fn stack_frame_margin_relative_offset_is_frame_relative() {
        assert_x(
            x_of(
                &h_offset(AnchorRelativeFrom::Margin, INCH),
                AnchorFrame::Stack,
            ),
            72.0,
            "1in from the frame origin",
        );
    }

    /// §20.4.3.1: `left`/`center`/`right` place the object inside the region
    /// named by `relativeFrom` — here the 72..540 text area.
    #[test]
    fn page_frame_align_resolves_against_the_text_area() {
        for (alignment, expected) in [
            (AnchorAlignment::Left, 72.0),
            (AnchorAlignment::Center, 72.0 + (468.0 - 100.0) * 0.5),
            (AnchorAlignment::Right, 72.0 + 468.0 - 100.0),
        ] {
            let got = x_of(
                &h_align(AnchorRelativeFrom::Margin, alignment),
                AnchorFrame::Page,
            );
            assert_x(
                got,
                expected,
                &format!("{alignment:?} within the text area"),
            );
        }
    }

    /// `page` aligns against the whole sheet, margins included.
    #[test]
    fn page_frame_page_align_resolves_against_the_whole_page() {
        for (alignment, expected) in [
            (AnchorAlignment::Left, 0.0),
            (AnchorAlignment::Center, (612.0 - 100.0) * 0.5),
            (AnchorAlignment::Right, 612.0 - 100.0),
        ] {
            let got = x_of(
                &h_align(AnchorRelativeFrom::Page, alignment),
                AnchorFrame::Page,
            );
            assert_x(got, expected, &format!("{alignment:?} within the page"));
        }
    }

    /// `Stack` has no container extent at extraction time, so a margin-relative
    /// alignment has nothing to align *within* and collapses onto the frame
    /// origin. `Right`/`Center` therefore run negative by the object's own
    /// width. Pinned because every new region has to keep respecting it —
    /// giving one of them a real extent here would shift header floats.
    #[test]
    fn stack_frame_align_collapses_to_the_frame_origin() {
        for (alignment, expected) in [
            (AnchorAlignment::Left, 0.0),
            (AnchorAlignment::Center, -50.0),
            (AnchorAlignment::Right, -100.0),
        ] {
            let got = x_of(
                &h_align(AnchorRelativeFrom::Margin, alignment),
                AnchorFrame::Stack,
            );
            assert_x(got, expected, &format!("{alignment:?} in a stack frame"));
        }
    }

    /// §20.4.3.4 `leftMargin` — the 0..72 strip, *not* the text area beside
    /// it. A `left`-aligned object in it sits flush with the sheet's edge, and
    /// an offset counts from that edge.
    #[test]
    fn left_margin_is_the_strip_from_the_page_edge_to_the_margin() {
        let from = AnchorRelativeFrom::LeftMargin;
        assert_x(
            x_of(&h_offset(from, INCH), AnchorFrame::Page),
            72.0,
            "1in from the sheet's left edge",
        );
        assert_x(
            x_of(&h_align(from, AnchorAlignment::Left), AnchorFrame::Page),
            0.0,
            "flush with the sheet's left edge",
        );
        // A 100pt object is wider than the 72pt strip, so right-aligning it
        // hangs it past the margin edge — which is what Word draws.
        assert_x(
            x_of(&h_align(from, AnchorAlignment::Right), AnchorFrame::Page),
            -28.0,
            "right edge on the margin edge",
        );
    }

    /// §20.4.3.4 `rightMargin` — the mirror strip, 540..612.
    #[test]
    fn right_margin_is_the_strip_from_the_margin_to_the_page_edge() {
        let from = AnchorRelativeFrom::RightMargin;
        assert_x(
            x_of(&h_offset(from, INCH), AnchorFrame::Page),
            612.0,
            "1in past the right margin edge, i.e. the sheet's right edge",
        );
        assert_x(
            x_of(&h_align(from, AnchorAlignment::Left), AnchorFrame::Page),
            540.0,
            "flush with the right margin edge",
        );
    }

    /// §20.4.3.4 `insideMargin`/`outsideMargin` are page-parity dependent, and
    /// the page a float lands on is not known at extraction time. The odd-page
    /// reading — inside = left, outside = right — is the Tier-0 collapse, and
    /// it is exact for the single-sided documents that are nearly all of them.
    #[test]
    fn parity_margins_take_their_odd_page_reading() {
        let inside = x_of(
            &h_align(AnchorRelativeFrom::InsideMargin, AnchorAlignment::Left),
            AnchorFrame::Page,
        );
        let outside = x_of(
            &h_align(AnchorRelativeFrom::OutsideMargin, AnchorAlignment::Left),
            AnchorFrame::Page,
        );
        assert_x(inside, 0.0, "inside = the left margin on an odd page");
        assert_x(outside, 540.0, "outside = the right margin on an odd page");
    }

    /// The deferred pair is a genuine mirror — whichever page the object lands
    /// on, `inside` and `outside` name opposite strips. Asserted as a relation
    /// between the two readings rather than through their coordinates, so it
    /// still holds for a page whose margins are not the ones this fixture uses.
    #[test]
    fn inside_and_outside_margins_mirror_each_other() {
        use super::{horizontal_region, AnchorRegion::Mirrored, FrameGeometry};

        let geom = FrameGeometry::new(&default_state().page_config, AnchorFrame::Page);
        let (
            Mirrored {
                odd: inside_odd,
                even: inside_even,
            },
            Mirrored {
                odd: outside_odd,
                even: outside_even,
            },
        ) = (
            horizontal_region(AnchorRelativeFrom::InsideMargin, &geom),
            horizontal_region(AnchorRelativeFrom::OutsideMargin, &geom),
        )
        else {
            panic!("both references are parity-dependent");
        };

        assert_eq!(inside_odd, outside_even, "inside on odd = outside on even");
        assert_eq!(inside_even, outside_odd, "inside on even = outside on odd");
        assert_ne!(inside_odd, inside_even, "the two pages differ");
    }

    /// §20.4.3.1: `inside` is the binding edge — left on an odd (recto) page,
    /// right on an even one — and `outside` is the trimmed edge opposite it.
    /// Both readings are carried, because the page is not known here.
    #[test]
    fn inside_and_outside_alignments_mirror_on_even_pages() {
        let margin = AnchorRelativeFrom::Margin;
        // Text area 72..540; a 100pt object is flush left at 72, flush right
        // at 440.
        for (alignment, odd, even) in [
            (AnchorAlignment::Inside, 72.0, 440.0),
            (AnchorAlignment::Outside, 440.0, 72.0),
        ] {
            let anchor = h_align(margin, alignment);
            assert_x(
                x_on(&anchor, AnchorFrame::Page, PageParity::Odd),
                odd,
                &format!("{alignment:?} on an odd page"),
            );
            assert_x(
                x_on(&anchor, AnchorFrame::Page, PageParity::Even),
                even,
                &format!("{alignment:?} on an even page"),
            );
        }
    }

    /// The two parity channels compose: an `inside` alignment *within* an
    /// `insideMargin` region mirrors through both at once, and must not
    /// double-mirror back to the odd-page answer.
    #[test]
    fn a_mirrored_alignment_inside_a_mirrored_region_mirrors_once() {
        let anchor = h_align(AnchorRelativeFrom::InsideMargin, AnchorAlignment::Inside);
        // Odd: insideMargin = the 0..72 strip, inside-aligned = its left edge.
        assert_x(
            x_on(&anchor, AnchorFrame::Page, PageParity::Odd),
            0.0,
            "odd",
        );
        // Even: insideMargin = the 540..612 strip, inside-aligned = its
        // *right* edge, less the object's width.
        assert_x(
            x_on(&anchor, AnchorFrame::Page, PageParity::Even),
            540.0 + 72.0 - 100.0,
            "even",
        );
    }

    /// An anchor that uses neither parity channel collapses to `Absolute`, so
    /// a single-sided document carries no deferral at all and every downstream
    /// `resolve` is a no-op. This is what keeps the ADT from costing anything
    /// on the documents that are nearly all of them.
    #[test]
    fn an_unmirrored_anchor_carries_no_deferral() {
        for alignment in [
            AnchorAlignment::Left,
            AnchorAlignment::Center,
            AnchorAlignment::Right,
        ] {
            let x = resolve_anchor_x(
                &h_align(AnchorRelativeFrom::Margin, alignment),
                Pt::new(100.0),
                &default_state(),
                AnchorFrame::Page,
            );
            assert!(
                matches!(x, FloatingImageX::Absolute(_)),
                "{alignment:?} is parity-independent, got {x:?}"
            );
        }
        let mirrored = resolve_anchor_x(
            &h_align(AnchorRelativeFrom::Margin, AnchorAlignment::Inside),
            Pt::new(100.0),
            &default_state(),
            AnchorFrame::Page,
        );
        assert!(
            matches!(mirrored, FloatingImageX::PageParity { .. }),
            "inside is parity-dependent, got {mirrored:?}"
        );
    }

    /// §20.4.3.4 `character` is the anchor's position in the text run, which
    /// float extraction runs before. It falls back to the text area — the
    /// behaviour it had while sharing a catch-all with `margin`, now reached by
    /// a named arm that logs.
    #[test]
    fn character_relative_falls_back_to_the_text_area() {
        let from = AnchorRelativeFrom::Character;
        assert_x(
            x_of(&h_offset(from, INCH), AnchorFrame::Page),
            144.0,
            "offset",
        );
        assert_x(
            x_of(&h_align(from, AnchorAlignment::Left), AnchorFrame::Page),
            72.0,
            "align",
        );
    }

    /// The margin strips are page-derived, so a header float can reach them:
    /// `Stack` expresses them relative to the frame origin, and the caller's
    /// `+ margins.left` shift lands them on the page coordinate they name.
    #[test]
    fn stack_frame_margin_strips_survive_the_round_trip_into_page_space() {
        const BODY_MARGIN: f32 = 72.0;
        for (from, page_x) in [
            (AnchorRelativeFrom::LeftMargin, 0.0),
            (AnchorRelativeFrom::RightMargin, 540.0),
        ] {
            let got = x_of(&h_align(from, AnchorAlignment::Left), AnchorFrame::Stack);
            assert_x(
                got + BODY_MARGIN,
                page_x,
                &format!("{from:?} after the caller's shift"),
            );
        }
    }

    // ── §14.1.2.5 VML fill resolution ────────────────────────────────────

    use super::{build_vml_rect_shape, model, resolve_vml_solid_fill};
    use crate::model::{VmlColor, VmlFill, VmlFillType, VmlLength, VmlLengthUnit, VmlNamedColor};
    use crate::render::layout::draw_command::ResolvedFill;

    fn solid_rgb(fill: &ResolvedFill) -> (f32, f32, f32) {
        let ResolvedFill::Solid(c) = fill else {
            panic!("expected a solid fill, got {fill:?}");
        };
        (c.r, c.g, c.b)
    }

    /// §14.1.2.5: the `<v:fill>` child element wins over `@fillcolor`.
    #[test]
    fn vml_fill_child_overrides_the_fillcolor_attribute() {
        let common = model::VmlCommonAttrs {
            fill_color: Some(VmlColor::Rgb(255, 0, 0)),
            fill: Some(VmlFill {
                fill_type: VmlFillType::Solid,
                color: Some(VmlColor::Rgb(0, 255, 0)),
                ..Default::default()
            }),
            ..Default::default()
        };
        assert_eq!(solid_rgb(&resolve_vml_solid_fill(&common)), (0.0, 1.0, 0.0));
    }

    /// With no child element the attribute is the whole story.
    #[test]
    fn vml_fillcolor_attribute_applies_without_a_fill_child() {
        let common = model::VmlCommonAttrs {
            fill_color: Some(VmlColor::Rgb(0, 0, 255)),
            ..Default::default()
        };
        assert_eq!(solid_rgb(&resolve_vml_solid_fill(&common)), (0.0, 0.0, 1.0));
    }

    /// A `<v:fill type="solid"/>` carrying no `@color` is not an override —
    /// it falls through to the attribute rather than blanking the shape.
    #[test]
    fn vml_solid_fill_without_a_color_falls_back_to_the_attribute() {
        let common = model::VmlCommonAttrs {
            fill_color: Some(VmlColor::Rgb(255, 0, 0)),
            fill: Some(VmlFill {
                fill_type: VmlFillType::Solid,
                color: None,
                ..Default::default()
            }),
            ..Default::default()
        };
        assert_eq!(solid_rgb(&resolve_vml_solid_fill(&common)), (1.0, 0.0, 0.0));
    }

    /// A non-solid fill degrades to no-fill and — deliberately — does *not*
    /// fall back to `@fillcolor`: the attribute is the gradient's own start
    /// colour, so painting the shape flat in it would be worse than leaving
    /// the outline and text to carry the shape.
    #[test]
    fn vml_non_solid_fills_degrade_to_no_fill_without_falling_back() {
        for fill_type in [
            VmlFillType::Gradient,
            VmlFillType::GradientRadial,
            VmlFillType::Tile,
            VmlFillType::Frame,
            VmlFillType::Pattern,
        ] {
            let common = model::VmlCommonAttrs {
                fill_color: Some(VmlColor::Rgb(255, 0, 0)),
                fill: Some(VmlFill {
                    fill_type,
                    color: Some(VmlColor::Rgb(0, 255, 0)),
                    ..Default::default()
                }),
                ..Default::default()
            };
            assert!(
                matches!(resolve_vml_solid_fill(&common), ResolvedFill::None),
                "{fill_type:?} must not paint",
            );
        }
    }

    /// Named colours are parsed but not yet resolved to RGB, so they leave the
    /// shape unfilled rather than guessing.
    #[test]
    fn vml_named_colors_are_not_resolved_yet() {
        let common = model::VmlCommonAttrs {
            fill_color: Some(VmlColor::Named(VmlNamedColor::Black)),
            ..Default::default()
        };
        assert!(matches!(
            resolve_vml_solid_fill(&common),
            ResolvedFill::None
        ));
    }

    // ── §14.1.2.19 VML rect construction ─────────────────────────────────

    fn vml_pt(value: f64) -> VmlLength {
        VmlLength {
            value,
            unit: VmlLengthUnit::Pt,
        }
    }

    /// A `position:absolute` rect at `(x, y)` sized `w x h`, all in points.
    fn vml_rect(x: f64, y: f64, w: f64, h: f64) -> model::VmlCommonAttrs {
        model::VmlCommonAttrs {
            style: model::VmlStyle {
                position: Some(crate::model::CssPosition::Absolute),
                margin_left: Some(vml_pt(x)),
                margin_top: Some(vml_pt(y)),
                width: Some(vml_pt(w)),
                height: Some(vml_pt(h)),
                ..Default::default()
            },
            ..Default::default()
        }
    }

    /// The path is a closed rectangle in *shape-local* points — the painter
    /// applies `(x, y)`, so baking the position into the path would double it.
    #[test]
    fn vml_rect_is_a_closed_rectangle_in_shape_local_points() {
        use crate::render::resolve::shape_geometry::PathVerb;

        let shape = build_vml_rect_shape(
            &vml_rect(30.0, 40.0, 200.0, 10.0),
            &default_state(),
            AnchorFrame::Page,
        )
        .expect("a positioned, sized rect builds");

        assert_x(
            shape.x.resolve(PageParity::Odd).raw(),
            30.0,
            "page-frame x is the style's margin-left",
        );
        let FloatingImageY::RelativeToParagraph(y) = shape.y else {
            panic!("VML rects anchor to the host paragraph");
        };
        assert_x(y.raw(), 40.0, "y is the style's margin-top");
        assert_x(shape.size.width.raw(), 200.0, "width");
        assert_x(shape.size.height.raw(), 10.0, "height");

        let [sub] = &shape.paths[..] else {
            panic!("one sub-path, got {}", shape.paths.len());
        };
        assert_eq!(sub.verbs.len(), 5, "4 corners + close");
        assert!(matches!(sub.verbs[0], PathVerb::MoveTo(o) if o.x == Pt::ZERO && o.y == Pt::ZERO));
        assert!(matches!(sub.verbs[4], PathVerb::Close));
    }

    /// Same round-trip as the DrawingML path: the stack emitter shifts by the
    /// body's left margin, so the page-relative x is pre-compensated.
    #[test]
    fn stack_frame_vml_rect_backs_out_the_left_margin() {
        let shape = build_vml_rect_shape(
            &vml_rect(80.0, 0.0, 10.0, 10.0),
            &default_state(),
            AnchorFrame::Stack,
        )
        .expect("a positioned, sized rect builds");
        assert_eq!(
            shape.x,
            FloatingImageX::Absolute(Pt::new(8.0)),
            "80pt page x minus the 72pt margin"
        );
    }

    /// Without `position:absolute` there is no position to honour, and a rect
    /// with no extent cannot render — both drop out rather than emitting a
    /// degenerate shape at the origin.
    #[test]
    fn vml_rect_needs_absolute_positioning_and_a_positive_extent() {
        let mut unpositioned = vml_rect(30.0, 40.0, 200.0, 10.0);
        unpositioned.style.position = None;
        assert!(
            build_vml_rect_shape(&unpositioned, &default_state(), AnchorFrame::Page).is_none(),
            "no position:absolute"
        );

        for (w, h) in [(0.0, 10.0), (200.0, 0.0), (-5.0, 10.0)] {
            assert!(
                build_vml_rect_shape(
                    &vml_rect(30.0, 40.0, w, h),
                    &default_state(),
                    AnchorFrame::Page
                )
                .is_none(),
                "{w} x {h} has no drawable area"
            );
        }
    }

    /// `@stroked` reaches the sub-path so the painter knows whether to outline.
    #[test]
    fn vml_rect_carries_the_stroked_flag_onto_its_path() {
        for stroked in [None, Some(false), Some(true)] {
            let mut common = vml_rect(0.0, 0.0, 10.0, 10.0);
            common.stroked = stroked;
            let shape =
                build_vml_rect_shape(&common, &default_state(), AnchorFrame::Page).expect("builds");
            assert_eq!(
                shape.paths[0].stroked,
                stroked == Some(true),
                "stroked={stroked:?}"
            );
        }
    }
}
