//! Extract floating (anchor) images from paragraph inlines and resolve their
//! absolute page positions.

use crate::model::{self, Paragraph};
use crate::render::dimension::Pt;
use crate::render::layout::section::{FloatingImage, FloatingImageY};

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
