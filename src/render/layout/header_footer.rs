use crate::model::*;

use super::fragment::*;
use super::measurer;
use super::{offset_command, DrawCommand, LayoutConfig, LayoutedPage};

/// Render header and footer content on each page.
pub(super) fn render_headers_footers(
    pages: &mut [LayoutedPage],
    doc_defaults: &DocDefaultsLayout,
    default_tab_stop_pt: f32,
    config: &LayoutConfig,
    font_mgr: &skia_safe::FontMgr,
    image_cache: &super::ImageCache,
) {
    let measurer = measurer::TextMeasurer::with_font_mgr(font_mgr.clone());
    let num_pages = pages.len() as u32;

    for (page_idx, page) in pages.iter_mut().enumerate() {
        let page_number = (page_idx + 1) as u32;
        let page_width = page.page_width;
        let page_height = page.page_height;
        let margin_left = config.margin_left;
        let margin_right = config.margin_right;
        let content_width = page_width - margin_left - margin_right;
        let header_y = config.header_margin;
        let footer_y = page_height - config.margin_bottom;

        let margin_top = config.margin_top;
        let margin_bottom = config.margin_bottom;

        let field_ctx = FieldContext {
            page_number,
            num_pages,
        };

        // Render header
        if let Some(ref header) = doc_defaults.default_header {
            let (commands, _header_bottom) = layout_header_footer_blocks(
                &header.blocks,
                margin_left,
                header_y,
                content_width,
                margin_top,
                page_height,
                page_width,
                page_height,
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
                Some(&field_ctx),
                image_cache,
            );
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
                page_width,
                page_height,
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
                Some(&field_ctx),
                image_cache,
            );
            page.commands.extend(commands);
        }
    }
}

/// Layout header/footer blocks at a fixed position.
/// Returns (draw_commands, max_y_extent).
pub(super) fn layout_header_footer_blocks(
    blocks: &[Block],
    x_start: f32,
    y_start: f32,
    content_width: f32,
    margin_extent: f32,
    y_limit: f32,
    page_width: f32,
    page_height: f32,
    defaults: &DocDefaultsLayout,
    measurer: &measurer::TextMeasurer,
    default_tab_stop_pt: f32,
    field_ctx: Option<&FieldContext>,
    image_cache: &super::ImageCache,
) -> (Vec<DrawCommand>, f32) {
    let mut commands = Vec::new();
    let mut cursor_y = y_start;
    let mut max_y = y_start;

    for block in blocks {
        if cursor_y >= y_limit {
            break;
        }
        if let Block::Paragraph(para) = block {
            let spacing = para.properties.spacing.unwrap_or_default();
            cursor_y += spacing.before_pt();

            // Render floating images with alignment support
            for float in &para.floats {
                if !image_cache.contains(&float.rel_id) {
                    continue;
                }
                let scale = f32::min(1.0, content_width / float.width_pt.max(1.0));
                let img_w = float.width_pt * scale;
                let img_h = float.height_pt * scale;
                let img_x = if let Some(pct) = float.pct_pos_h {
                    pct as f32 / 100_000.0 * page_width
                } else {
                    match float.align_h.as_deref() {
                        Some("right") => x_start + content_width - img_w,
                        Some("center") => x_start + (content_width - img_w) / 2.0,
                        Some("left") => x_start,
                        _ => x_start + float.offset_x_pt,
                    }
                };
                let img_y = if let Some(pct) = float.pct_pos_v {
                    pct as f32 / 100_000.0 * page_height
                } else {
                    match float.align_v.as_deref() {
                        Some("center") => (margin_extent - img_h) / 2.0,
                        Some("bottom") => margin_extent - img_h,
                        Some("top") => 0.0,
                        _ => cursor_y + float.offset_y_pt,
                    }
                };
                max_y = max_y.max(img_y + img_h);
                let image = image_cache.get(&float.rel_id);
                commands.push(DrawCommand::Image {
                    x: img_x,
                    y: img_y,
                    width: img_w,
                    height: img_h,
                    image,
                });
            }

            // MEASURE: collect fragments and produce measured lines
            let fragments = collect_fragments_with_fields(
                &para.runs,
                content_width,
                margin_extent,
                defaults,
                measurer,
                field_ctx,
                image_cache,
            );

            if fragments.is_empty() {
                cursor_y += spacing.line_pt();
                cursor_y += spacing.after_pt();
                continue;
            }

            let indent = para.properties.indentation.unwrap_or_default();
            let avail = content_width - indent.left_pt() - indent.right_pt();

            let measured = measure_lines(
                &fragments,
                x_start + indent.left_pt(),
                avail,
                0.0,
                para.properties.alignment,
                spacing.line_spacing(),
                &para.properties.tab_stops,
                default_tab_stop_pt,
                image_cache,
            );

            // PAINT: offset measured commands by cursor_y
            for line in &measured.lines {
                for cmd in &line.commands {
                    commands.push(offset_command(cmd, cursor_y));
                }
            }
            cursor_y += measured.total_height;
            cursor_y += spacing.after_pt();
        }
    }

    max_y = max_y.max(cursor_y);
    (commands, max_y)
}

pub(super) fn to_roman(mut n: u32) -> String {
    let table = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
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
