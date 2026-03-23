use crate::dimension::Pt;
use crate::geometry::PtRect;
use crate::model::*;

use super::context::LayoutConstraints;
use super::fragment::*;
use super::measurer;
use super::{offset_command, DrawCommand, LayoutConfig, LayoutedPage};

/// Render header and footer content on each page.
pub(super) fn render_headers_footers(
    pages: &mut [LayoutedPage],
    doc_defaults: &DocDefaultsLayout,
    default_tab_stop_pt: Pt,
    config: &LayoutConfig,
    font_mgr: &skia_safe::FontMgr,
    image_cache: &super::ImageCache,
) {
    let measurer = measurer::TextMeasurer::new(font_mgr.clone());
    let num_pages = pages.len() as u32;

    let hf_constraints = LayoutConstraints::for_header_footer(config, config.margins.top);

    for (page_idx, page) in pages.iter_mut().enumerate() {
        let page_number = (page_idx + 1) as u32;
        let field_ctx = FieldContext {
            page_number,
            num_pages,
        };

        // Render header
        if let Some(ref header) = doc_defaults.default_header {
            let (commands, _header_bottom) = layout_header_footer_blocks(
                &header.blocks,
                &hf_constraints,
                config.header_margin,
                config.margins.top,
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
            let footer_y = hf_constraints.page_size.height - config.margins.bottom;
            let (commands, _) = layout_header_footer_blocks(
                &footer.blocks,
                &hf_constraints,
                footer_y,
                config.margins.bottom,
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
    constraints: &LayoutConstraints,
    y_start: Pt,
    margin_extent: Pt,
    defaults: &DocDefaultsLayout,
    measurer: &measurer::TextMeasurer,
    default_tab_stop_pt: Pt,
    field_ctx: Option<&FieldContext>,
    image_cache: &super::ImageCache,
) -> (Vec<DrawCommand>, Pt) {
    let x_start = constraints.x_origin;
    let content_width = constraints.available_width;
    let page_width = constraints.page_size.width;
    let page_height = constraints.page_size.height;

    let mut commands = Vec::new();
    let mut cursor_y = y_start;
    let mut max_y = y_start;

    for block in blocks {
        if cursor_y >= page_height {
            break;
        }
        if let Block::Paragraph(para) = block {
            let spacing = para.properties.spacing.unwrap_or_default();
            cursor_y += spacing.before.map(Pt::from).unwrap_or(Pt::ZERO);

            // Render floating images with alignment support
            for float in &para.floats {
                if !image_cache.contains(&float.rel_id) {
                    continue;
                }
                let fw = float.size.width;
                let fh = float.size.height;
                let scale = f32::min(1.0, content_width / fw.max(Pt::new(1.0)));
                let img_w = fw * scale;
                let img_h = fh * scale;
                let img_x = if let Some(pct) = float.pct_pos_h {
                    page_width * (pct as f32 / 100_000.0)
                } else {
                    match float.align_h.as_deref() {
                        Some("right") => x_start + content_width - img_w,
                        Some("center") => x_start + (content_width - img_w) / 2.0,
                        Some("left") => x_start,
                        _ => x_start + float.offset.x,
                    }
                };
                let img_y = if let Some(pct) = float.pct_pos_v {
                    page_height * (pct as f32 / 100_000.0)
                } else {
                    match float.align_v.as_deref() {
                        Some("center") => (margin_extent - img_h) / 2.0,
                        Some("bottom") => margin_extent - img_h,
                        Some("top") => Pt::ZERO,
                        _ => cursor_y + float.offset.y,
                    }
                };
                max_y = max_y.max(img_y + img_h);
                let image = image_cache.get(&float.rel_id);
                commands.push(DrawCommand::Image {
                    rect: PtRect::from_xywh(img_x, img_y, img_w, img_h),
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
                cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
                continue;
            }

            let indent = para.properties.indentation.unwrap_or_default();
            let avail = content_width
                - indent.left.map(Pt::from).unwrap_or(Pt::ZERO)
                - indent.right.map(Pt::from).unwrap_or(Pt::ZERO);

            let measured = measure_lines(
                &fragments,
                x_start + indent.left.map(Pt::from).unwrap_or(Pt::ZERO),
                avail,
                Pt::ZERO,
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
            cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
        }
    }

    max_y = max_y.max(cursor_y);
    (commands, max_y)
}
