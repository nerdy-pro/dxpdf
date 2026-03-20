use crate::model::*;
use crate::units::*;

use super::fragment::*;
use super::measurer;
use super::{DrawCommand, LayoutConfig, LayoutedPage};

/// Render header and footer content on each page.
pub(super) fn render_headers_footers(
    pages: &mut [LayoutedPage],
    doc_defaults: &DocDefaultsLayout,
    default_tab_stop_pt: f32,
    config: &LayoutConfig,
) {
    let measurer = measurer::TextMeasurer::new();

    for page in pages.iter_mut() {
        let page_width = page.page_width;
        let page_height = page.page_height;
        let margin_left = config.margin_left;
        let margin_right = config.margin_right;
        let content_width = page_width - margin_left - margin_right;
        let header_y = config.header_margin;
        // Footer starts at the body's bottom margin boundary
        let footer_y = page_height - config.margin_bottom;

        let margin_top = config.margin_top;
        let margin_bottom = config.margin_bottom;

        // Render header
        if let Some(ref header) = doc_defaults.default_header {
            let (commands, _header_bottom) = layout_header_footer_blocks(
                &header.blocks,
                margin_left,
                header_y,
                content_width,
                margin_top,
                page_height, // don't clip header text — let it extend if needed
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
            );
            // Insert header commands at the beginning so they render behind body
            let body_commands = std::mem::take(&mut page.commands);
            page.commands = commands;
            page.commands.extend(body_commands);
        }

        // Render footer
        if let Some(ref footer) = doc_defaults.default_footer {
            let (commands, _) = layout_header_footer_blocks(
                &footer.blocks,
                margin_left,
                footer_y,
                content_width,
                margin_bottom,
                page_height,
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
            );
            page.commands.extend(commands);
        }
    }
}

/// Layout header/footer blocks at a fixed position.
/// Returns (draw_commands, max_y_extent) where max_y_extent is the bottommost
/// y coordinate used by any content including float images.
pub(super) fn layout_header_footer_blocks(
    blocks: &[Block],
    x_start: f32,
    y_start: f32,
    content_width: f32,
    margin_extent: f32,
    y_limit: f32,
    defaults: &DocDefaultsLayout,
    measurer: &measurer::TextMeasurer,
    default_tab_stop_pt: f32,
) -> (Vec<DrawCommand>, f32) {
    use super::fragment::*;

    let mut commands = Vec::new();
    let mut cursor_y = y_start;
    let mut max_y = y_start;

    for block in blocks {
        // Stop rendering if we've exceeded the allowed area
        if cursor_y >= y_limit {
            break;
        }
        if let Block::Paragraph(para) = block {
            let spacing = match para.properties.spacing {
                Some(s) => s,
                None => Spacing::default(),
            };
            cursor_y += spacing.before_pt();

            // Render floating images with alignment support
            for float in &para.floats {
                if float.data.is_empty() {
                    continue;
                }
                let scale = f32::min(
                    1.0,
                    content_width / float.width_pt.max(1.0),
                );
                let img_w = float.width_pt * scale;
                let img_h = float.height_pt * scale;
                // Use alignment if specified, otherwise offset
                let img_x = match float.align_h.as_deref() {
                    Some("right") => x_start + content_width - img_w,
                    Some("center") => x_start + (content_width - img_w) / 2.0,
                    Some("left") => x_start,
                    _ => x_start + float.offset_x_pt,
                };
                let img_y = match float.align_v.as_deref() {
                    Some("center") => (margin_extent - img_h) / 2.0,
                    Some("bottom") => margin_extent - img_h,
                    Some("top") => 0.0,
                    _ => cursor_y + float.offset_y_pt,
                };
                max_y = max_y.max(img_y + img_h);
                commands.push(DrawCommand::Image {
                    x: img_x,
                    y: img_y,
                    width: img_w,
                    height: img_h,
                    data: float.data.clone(),
                });
            }

            let fragments = collect_fragments(
                &para.runs,
                content_width,
                HF_MAX_IMAGE_HEIGHT,
                defaults,
                measurer,
            );

            if fragments.is_empty() {
                cursor_y += spacing.line_pt();
                cursor_y += spacing.after_pt();
                continue;
            }

            let indent = para.properties.indentation.unwrap_or_default();
            let mut line_start = 0;
            let mut is_first_line = true;

            while line_start < fragments.len() {
                // Skip leading spaces
                if !is_first_line {
                    while line_start < fragments.len() {
                        if let Fragment::Text { ref text, .. } = fragments[line_start] {
                            if text.trim().is_empty() {
                                line_start += 1;
                                continue;
                            }
                        }
                        break;
                    }
                    if line_start >= fragments.len() {
                        break;
                    }
                }

                let avail = content_width - indent.left_pt() - indent.right_pt();
                let (line_end, _) = fit_fragments(&fragments[line_start..], avail);
                let actual_end = line_start + line_end.max(1);

                let frag_height = fragments[line_start..actual_end]
                    .iter()
                    .map(|f| f.height())
                    .fold(0.0_f32, f32::max);
                let line_height = match spacing.line_pt_opt() {
                    Some(lh) => frag_height.max(lh),
                    None => frag_height,
                };
                cursor_y += line_height;

                let used_width = measure_fragments(&fragments[line_start..actual_end]);
                let x_offset = match para.properties.alignment {
                    Some(Alignment::Center) => (avail - used_width) / 2.0,
                    Some(Alignment::Right) => avail - used_width,
                    _ => 0.0,
                };
                let mut x = x_start + indent.left_pt() + x_offset;

                for frag in &fragments[line_start..actual_end] {
                    match frag {
                        Fragment::Text {
                            text, font_family, font_size, bold, italic,
                            underline, color, shading, char_spacing_pt, measured_width, ..
                        } => {
                            let c = color.map(|c| (c.r, c.g, c.b)).unwrap_or((0, 0, 0));
                            if let Some(bg) = shading {
                                commands.push(DrawCommand::Rect {
                                    x,
                                    y: cursor_y - line_height,
                                    width: *measured_width,
                                    height: line_height,
                                    color: (bg.r, bg.g, bg.b),
                                });
                            }
                            commands.push(DrawCommand::Text {
                                x, y: cursor_y, text: text.clone(),
                                font_family: font_family.clone(),
                                char_spacing_pt: *char_spacing_pt,
                                font_size: *font_size, bold: *bold, italic: *italic,
                                color: c,
                            });
                            if *underline {
                                commands.push(DrawCommand::Underline {
                                    x1: x, y1: cursor_y + crate::units::UNDERLINE_Y_OFFSET,
                                    x2: x + measured_width,
                                    y2: cursor_y + crate::units::UNDERLINE_Y_OFFSET,
                                    color: c, width: crate::units::UNDERLINE_STROKE_WIDTH,
                                });
                            }
                            x += measured_width;
                        }
                        Fragment::Image { width, height, data } => {
                            commands.push(DrawCommand::Image {
                                x, y: cursor_y - height,
                                width: *width, height: *height,
                                data: data.clone(),
                            });
                            x += width;
                        }
                        Fragment::Tab { .. } => {
                            let rel_x = x - x_start;
                            let next = find_next_tab_stop(
                                rel_x, &para.properties.tab_stops, default_tab_stop_pt,
                            );
                            x = x_start + next;
                        }
                        Fragment::LineBreak { .. } => {}
                    }
                }

                line_start = actual_end;
                is_first_line = false;
            }

            cursor_y += spacing.after_pt();
        }
    }

    max_y = max_y.max(cursor_y);
    (commands, max_y)
}

pub(super) fn to_roman(mut n: u32) -> String {
    let table = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ];
    let mut result = String::new();
    for &(value, numeral) in &table {
        while n >= value {
            result.push_str(numeral);
            n -= value;
        }
    }
    result
}
