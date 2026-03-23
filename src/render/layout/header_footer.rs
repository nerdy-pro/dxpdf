use crate::dimension::Pt;
use crate::geometry::PtRect;
use crate::model::*;

use super::context::LayoutConstraints;
use super::fragment::*;
use super::measure::compute_column_widths;
use super::measurer;
use super::{DrawCommand, LayoutConfig, LayoutedPage};

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

        // Render footer: measure at y=0 first, then shift so content
        // ends at footer_margin from the page bottom edge.
        if let Some(ref footer) = doc_defaults.default_footer {
            let page_height = hf_constraints.page_size().height;
            let (mut commands, content_bottom) = layout_header_footer_blocks(
                &footer.blocks,
                &hf_constraints,
                Pt::ZERO,
                page_height, // no clipping during measurement
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
                Some(&field_ctx),
                image_cache,
            );
            // Position so the footer content sits above the page bottom,
            // starting at footer_margin from the edge.
            let footer_top = page_height - config.footer_margin - content_bottom;
            for cmd in &mut commands {
                cmd.shift_y(footer_top);
            }
            page.commands.extend(commands);
        }
    }
}

/// Layout header/footer blocks at a fixed position.
/// Returns (draw_commands, max_y_extent).
#[allow(clippy::too_many_arguments)]
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
    let x_start = constraints.x_origin();
    let content_width = constraints.available_width();
    let page_width = constraints.page_size().width;
    let page_height = constraints.page_size().height;

    let mut commands = Vec::new();
    let mut cursor_y = y_start;
    let mut max_y = y_start;

    for block in blocks {
        if cursor_y >= page_height {
            break;
        }
        if let Block::Table(table) = block {
            let num_cols = table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
            if num_cols == 0 {
                continue;
            }
            let col_widths = compute_column_widths(&table.grid_cols, num_cols, content_width);

            for row in &table.rows {
                let mut row_height = Pt::ZERO;
                let mut row_commands = Vec::new();
                let mut grid_col_idx = 0;

                for cell in &row.cells {
                    let span = cell.grid_span.max(1) as usize;
                    let cell_x = x_start + col_widths[..grid_col_idx].iter().copied().sum::<Pt>();
                    let cell_w: Pt = (grid_col_idx..grid_col_idx + span)
                        .map(|i| col_widths.get(i).copied().unwrap_or(Pt::ZERO))
                        .sum();
                    grid_col_idx += span;

                    let mut cell_y = Pt::ZERO;
                    for cell_block in &cell.blocks {
                        if let Block::Paragraph(p) = cell_block {
                            let sp = p.properties.spacing.unwrap_or_default();
                            cell_y += sp.before_pt();
                            let frags = collect_fragments_with_fields(
                                &p.runs,
                                constraints,
                                defaults,
                                measurer,
                                field_ctx,
                                image_cache,
                            );
                            if frags.is_empty() {
                                cell_y += sp.line_pt();
                                cell_y += sp.after_pt();
                                continue;
                            }
                            let measured = measure_lines(
                                &frags,
                                cell_x,
                                cell_w,
                                Pt::ZERO,
                                p.properties.alignment,
                                sp.line_spacing(),
                                &p.properties.tab_stops,
                                default_tab_stop_pt,
                                image_cache,
                                None,
                            );
                            for line in &measured.lines {
                                for cmd in &line.commands {
                                    row_commands.push(cmd.offset_y(cell_y));
                                }
                            }
                            cell_y += measured.total_height;
                            cell_y += sp.after_pt();
                        }
                    }
                    row_height = row_height.max(cell_y);
                }

                // Offset all row commands by cursor_y
                for cmd in &row_commands {
                    commands.push(cmd.offset_y(cursor_y));
                }
                cursor_y += row_height;
            }
        } else if let Block::Paragraph(para) = block {
            let spacing = para.properties.spacing.unwrap_or_default();
            cursor_y += spacing.before_pt();

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
                constraints,
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
                Pt::ZERO,
                para.properties.alignment,
                spacing.line_spacing(),
                &para.properties.tab_stops,
                default_tab_stop_pt,
                image_cache,
                None,
            );

            // PAINT: offset measured commands by cursor_y
            for line in &measured.lines {
                for cmd in &line.commands {
                    commands.push(cmd.offset_y(cursor_y));
                }
            }
            cursor_y += measured.total_height;
            cursor_y += spacing.after_pt();
        }
    }

    max_y = max_y.max(cursor_y);
    (commands, max_y)
}
