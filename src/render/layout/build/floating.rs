//! Extract floating (anchor) images from paragraph inlines and resolve their
//! absolute page positions.

use crate::model::{self, Paragraph};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::layout::section::{FloatingImage, FloatingImageY, FloatingShape};
use crate::render::resolve::shape_geometry::build_geometry;
use crate::render::resolve::shape_visuals::resolve_shape_visuals;

use super::{BuildContext, BuildState};

/// Extract floating (anchor) images from a paragraph's inlines.
/// When `cell_context` is true, positions are resolved relative to the cell
/// origin (0,0) instead of the page margins.
pub(super) fn extract_floating_images(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &BuildState,
    cell_context: bool,
) -> Vec<FloatingImage> {
    use crate::model::{
        AnchorAlignment, AnchorPosition, AnchorRelativeFrom, ImagePlacement, Inline,
    };

    let mut images = Vec::new();

    fn find_anchor_images<'a>(inlines: &'a [Inline], out: &mut Vec<&'a crate::model::Image>) {
        for inline in inlines {
            match inline {
                Inline::Image(img) => {
                    if matches!(img.placement, ImagePlacement::Anchor(_)) {
                        out.push(img);
                    }
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

    for img in &anchor_imgs {
        if let ImagePlacement::Anchor(ref anchor) = img.placement {
            let rel_id = match crate::render::resolve::images::extract_image_rel_id(img) {
                Some(id) => id,
                None => {
                    eprintln!(
                        "  -> no rel_id, graphic.is_some()={}",
                        img.graphic.is_some()
                    );
                    continue;
                }
            };

            let image_data = match ctx.resolved.media.get(rel_id) {
                Some(entry) => entry.clone(),
                None => {
                    eprintln!(
                        "Anchor image: rel_id={} NOT FOUND in media (media has {} entries)",
                        rel_id.as_str(),
                        ctx.resolved.media.len()
                    );
                    continue;
                }
            };

            let w = Pt::from(img.extent.width);
            let h = Pt::from(img.extent.height);
            let pc = &state.page_config;

            // Resolve horizontal position.
            // In cell context, positions are relative to the cell origin.
            let (page_width, margin_left, margin_right) = if cell_context {
                (Pt::ZERO, Pt::ZERO, Pt::ZERO)
            } else {
                (pc.page_size.width, pc.margins.left, pc.margins.right)
            };
            let content_width = if cell_context {
                Pt::ZERO
            } else {
                page_width - margin_left - margin_right
            };

            let x = match &anchor.horizontal_position {
                AnchorPosition::Offset {
                    relative_from,
                    offset,
                } => {
                    let base = match relative_from {
                        AnchorRelativeFrom::Page => Pt::ZERO,
                        AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => margin_left,
                        _ => margin_left,
                    };
                    base + Pt::from(*offset)
                }
                AnchorPosition::Align {
                    relative_from,
                    alignment,
                } => {
                    let (area_left, area_width) = match relative_from {
                        AnchorRelativeFrom::Page => (Pt::ZERO, page_width),
                        AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => {
                            (margin_left, content_width)
                        }
                        _ => (margin_left, content_width),
                    };
                    match alignment {
                        AnchorAlignment::Left => area_left,
                        AnchorAlignment::Right => area_left + area_width - w,
                        AnchorAlignment::Center => area_left + (area_width - w) * 0.5,
                        _ => area_left,
                    }
                }
            };

            // Resolve vertical position.
            let y = match &anchor.vertical_position {
                AnchorPosition::Offset {
                    relative_from,
                    offset,
                } => {
                    let margin_top = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.margins.top
                    };
                    if cell_context {
                        // In cell context, all positions are relative to cell origin.
                        FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                    } else {
                        match relative_from {
                            AnchorRelativeFrom::Page => FloatingImageY::Absolute(Pt::from(*offset)),
                            AnchorRelativeFrom::Margin => {
                                FloatingImageY::Absolute(margin_top + Pt::from(*offset))
                            }
                            // §20.4.2.11: topMargin — offset from page top.
                            AnchorRelativeFrom::TopMargin => {
                                FloatingImageY::Absolute(Pt::from(*offset))
                            }
                            // §20.4.2.11: bottomMargin — offset from bottom margin edge.
                            AnchorRelativeFrom::BottomMargin => {
                                let page_height = pc.page_size.height;
                                let margin_bottom = pc.margins.bottom;
                                FloatingImageY::Absolute(
                                    page_height - margin_bottom + Pt::from(*offset),
                                )
                            }
                            AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                                FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                            }
                            _ => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                        }
                    }
                }
                AnchorPosition::Align {
                    relative_from,
                    alignment,
                } => {
                    let margin_top = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.margins.top
                    };
                    let page_height = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.page_size.height
                    };
                    let margin_bottom = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.margins.bottom
                    };
                    let (area_top, area_height) = match relative_from {
                        AnchorRelativeFrom::Page => (Pt::ZERO, page_height),
                        AnchorRelativeFrom::Margin => {
                            (margin_top, page_height - margin_top - margin_bottom)
                        }
                        // §20.4.2.11: topMargin = area from page top to top margin edge.
                        AnchorRelativeFrom::TopMargin => (Pt::ZERO, margin_top),
                        // §20.4.2.11: bottomMargin = area from bottom margin edge to page bottom.
                        AnchorRelativeFrom::BottomMargin => {
                            (page_height - margin_bottom, margin_bottom)
                        }
                        _ => (margin_top, page_height - margin_top - margin_bottom),
                    };
                    let y_pos = match alignment {
                        AnchorAlignment::Top => area_top,
                        AnchorAlignment::Bottom => area_top + area_height - h,
                        AnchorAlignment::Center => area_top + (area_height - h) * 0.5,
                        _ => area_top,
                    };
                    FloatingImageY::Absolute(y_pos)
                }
            };

            images.push(FloatingImage {
                image_data,
                size: crate::render::geometry::PtSize::new(w, h),
                x,
                y,
                wrap_top_and_bottom: matches!(
                    anchor.wrap,
                    crate::model::TextWrap::TopAndBottom { .. }
                ),
                dist_left: Pt::from(anchor.distance.left),
                dist_right: Pt::from(anchor.distance.right),
                behind_doc: anchor.behind_text,
            });
        }
    }

    images
}

// ── Floating shape extraction ──────────────────────────────────────────────

/// Extract floating (anchor) DrawingML shapes from a paragraph's inlines,
/// resolve their geometry + visuals, and compute their absolute page
/// positions. Pure: takes immutable references to ctx/state.
pub(super) fn extract_floating_shapes(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &BuildState,
    cell_context: bool,
) -> Vec<FloatingShape> {
    use crate::model::{GraphicContent, ImagePlacement, Inline};

    fn find_anchor_shapes<'a>(inlines: &'a [Inline], out: &mut Vec<&'a crate::model::Image>) {
        for inline in inlines {
            match inline {
                Inline::Image(img) => {
                    if matches!(img.placement, ImagePlacement::Anchor(_))
                        && matches!(img.graphic, Some(GraphicContent::WordProcessingShape(_)))
                    {
                        out.push(img);
                    }
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

        let visuals = resolve_shape_visuals(shape_props, ctx.resolved.theme.as_ref());

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

        let (x, y) = resolve_anchor_position(anchor, w, h, state, cell_context);

        shapes.push(FloatingShape {
            x,
            y,
            size: extent,
            rotation,
            flip_h,
            flip_v,
            wrap_top_and_bottom: matches!(anchor.wrap, crate::model::TextWrap::TopAndBottom { .. }),
            dist_left: Pt::from(anchor.distance.left),
            dist_right: Pt::from(anchor.distance.right),
            behind_doc: anchor.behind_text,
            paths: shape_path.paths,
            fill: visuals.fill,
            stroke: visuals.stroke,
            effects: visuals.effects,
        });
    }

    shapes
}

/// Shared anchor-position resolver used by both `extract_floating_images`
/// and `extract_floating_shapes`. Returns `(x, y)` in page or cell-relative
/// coordinates depending on `cell_context`.
fn resolve_anchor_position(
    anchor: &crate::model::AnchorProperties,
    content_w: Pt,
    content_h: Pt,
    state: &BuildState,
    cell_context: bool,
) -> (Pt, FloatingImageY) {
    use crate::model::{AnchorAlignment, AnchorPosition, AnchorRelativeFrom};

    let pc = &state.page_config;
    let (page_width, margin_left, margin_right) = if cell_context {
        (Pt::ZERO, Pt::ZERO, Pt::ZERO)
    } else {
        (pc.page_size.width, pc.margins.left, pc.margins.right)
    };
    let content_width = if cell_context {
        Pt::ZERO
    } else {
        page_width - margin_left - margin_right
    };

    let x = match &anchor.horizontal_position {
        AnchorPosition::Offset {
            relative_from,
            offset,
        } => {
            let base = match relative_from {
                AnchorRelativeFrom::Page => Pt::ZERO,
                AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => margin_left,
                _ => margin_left,
            };
            base + Pt::from(*offset)
        }
        AnchorPosition::Align {
            relative_from,
            alignment,
        } => {
            let (area_left, area_width) = match relative_from {
                AnchorRelativeFrom::Page => (Pt::ZERO, page_width),
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
    };

    let y = match &anchor.vertical_position {
        AnchorPosition::Offset {
            relative_from,
            offset,
        } => {
            let margin_top = if cell_context {
                Pt::ZERO
            } else {
                pc.margins.top
            };
            if cell_context {
                FloatingImageY::RelativeToParagraph(Pt::from(*offset))
            } else {
                match relative_from {
                    AnchorRelativeFrom::Page => FloatingImageY::Absolute(Pt::from(*offset)),
                    AnchorRelativeFrom::Margin => {
                        FloatingImageY::Absolute(margin_top + Pt::from(*offset))
                    }
                    AnchorRelativeFrom::TopMargin => FloatingImageY::Absolute(Pt::from(*offset)),
                    AnchorRelativeFrom::BottomMargin => {
                        let page_height = pc.page_size.height;
                        let margin_bottom = pc.margins.bottom;
                        FloatingImageY::Absolute(page_height - margin_bottom + Pt::from(*offset))
                    }
                    AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                        FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                    }
                    _ => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                }
            }
        }
        AnchorPosition::Align {
            relative_from,
            alignment,
        } => {
            let margin_top = if cell_context {
                Pt::ZERO
            } else {
                pc.margins.top
            };
            let page_height = if cell_context {
                Pt::ZERO
            } else {
                pc.page_size.height
            };
            let margin_bottom = if cell_context {
                Pt::ZERO
            } else {
                pc.margins.bottom
            };
            let (area_top, area_height) = match relative_from {
                AnchorRelativeFrom::Page => (Pt::ZERO, page_height),
                AnchorRelativeFrom::Margin => {
                    (margin_top, page_height - margin_top - margin_bottom)
                }
                AnchorRelativeFrom::TopMargin => (Pt::ZERO, margin_top),
                AnchorRelativeFrom::BottomMargin => (page_height - margin_bottom, margin_bottom),
                _ => (margin_top, page_height - margin_top - margin_bottom),
            };
            let y_pos = match alignment {
                AnchorAlignment::Top => area_top,
                AnchorAlignment::Bottom => area_top + area_height - content_h,
                AnchorAlignment::Center => area_top + (area_height - content_h) * 0.5,
                _ => area_top,
            };
            FloatingImageY::Absolute(y_pos)
        }
    };

    (x, y)
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
    for shape in &pict.shapes {
        if shape.text_box.is_some() {
            if let Some(pos) = vml_absolute_position(&shape.style) {
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
