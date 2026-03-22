use crate::model::*;
use crate::units::*;

use super::fragment::*;
use super::{offset_command, DrawCommand, Layouter};

// ============================================================
// Types for the measure→layout→paint pipeline
// ============================================================

/// Result of measuring a single cell's content.
struct MeasuredCell {
    commands: Vec<DrawCommand>,
    content_height: f32,
    col_width: f32,
}

/// Result of measuring all cells in a row.
struct MeasuredRow {
    cells: Vec<MeasuredCell>,
    col_x_positions: Vec<f32>,
    min_height: f32,
}

/// A vMerge span tracking entry.
struct VmergeSpan {
    total_height: f32,
    start_row: usize,
    row_count: usize,
}

// ============================================================
// Helper functions
// ============================================================

fn resolve_cell_margins(
    cell: &TableCell,
    table_default: &Option<CellMargins>,
    doc_default: &CellMargins,
) -> CellMargins {
    cell.cell_margins.or(*table_default).unwrap_or(*doc_default)
}

fn resolve_border(cell_border: Option<BorderDef>, table_border: BorderDef) -> BorderDef {
    cell_border.unwrap_or(table_border)
}

fn emit_border(
    commands: &mut Vec<DrawCommand>,
    border: &BorderDef,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
) {
    if border.is_visible() {
        commands.push(DrawCommand::Line {
            x1,
            y1,
            x2,
            y2,
            color: border.color_rgb(),
            width: crate::dimension::Pt::from(border.size).raw(),
        });
    }
}

// ============================================================
// Main entry point
// ============================================================

impl Layouter {
    pub(super) fn layout_table(&mut self, table: &Table, next_is_table: bool) {
        if table.rows.is_empty() {
            return;
        }
        let num_cols = table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
        if num_cols == 0 {
            return;
        }

        let content_width = self.config.content_width();
        let doc_cell_margins = self.doc_defaults.default_cell_margins;

        let col_widths: Vec<f32> = if !table.grid_cols.is_empty() {
            let grid_total: f32 = table.grid_cols.iter().map(|w| twips_to_pt(*w)).sum();
            let scale = if grid_total > 0.0 {
                content_width / grid_total
            } else {
                1.0
            };
            table
                .grid_cols
                .iter()
                .map(|w| twips_to_pt(*w) * scale)
                .collect()
        } else {
            vec![content_width / num_cols as f32; num_cols]
        };

        // ============================
        // PASS 1: MEASURE all cells
        // ============================
        let measured_rows: Vec<MeasuredRow> = table
            .rows
            .iter()
            .map(|row| {
                let min_height = row.height.map(twips_to_pt).unwrap_or(0.0);
                let row_height_limit = self.config.content_height();

                let mut col_x_positions = Vec::with_capacity(row.cells.len());
                let mut cell_widths_computed = Vec::with_capacity(row.cells.len());
                let mut grid_col_idx = 0;
                for cell in &row.cells {
                    let x =
                        self.config.margin_left + col_widths[..grid_col_idx].iter().sum::<f32>();
                    let w = self.cell_width(grid_col_idx, cell, &col_widths);
                    col_x_positions.push(x);
                    cell_widths_computed.push(w);
                    grid_col_idx += cell.grid_span.max(1) as usize;
                }

                let mut measured_cells = Vec::with_capacity(row.cells.len());

                for (col_idx, cell) in row.cells.iter().enumerate() {
                    let col_width = cell_widths_computed[col_idx];
                    let cell_x = col_x_positions[col_idx];
                    let margins =
                        resolve_cell_margins(cell, &table.default_cell_margins, &doc_cell_margins);
                    let pad_left = margins.left_pt();
                    let pad_right = margins.right_pt();
                    let pad_top = margins.top_pt();
                    let pad_bottom = margins.bottom_pt();
                    let cell_content_width = (col_width - pad_left - pad_right).max(1.0);

                    if cell.is_vmerge_continue() {
                        measured_cells.push(MeasuredCell {
                            commands: Vec::new(),
                            content_height: 0.0,
                            col_width,
                        });
                        continue;
                    }

                    let mut commands = Vec::new();
                    let mut cell_y = pad_top;

                    for block in &cell.blocks {
                        if let Block::Paragraph(p) = block {
                            let spacing =
                                self.resolve_cell_spacing(p.properties.spacing, table.cell_spacing);
                            cell_y += spacing.before_pt();

                            // Floating images in cell
                            for float in &p.floats {
                                if !self.image_cache.contains(&float.rel_id) {
                                    continue;
                                }
                                let fw = float.width.raw();
                                let fh = float.height.raw();
                                let scale = f32::min(
                                    1.0,
                                    f32::min(
                                        cell_content_width / fw.max(1.0),
                                        row_height_limit.max(1.0) / fh.max(1.0),
                                    ),
                                );
                                let img_w = fw * scale;
                                let img_h = fh * scale;
                                let img_x = cell_x + (col_width - img_w) / 2.0;
                                let image = self.image_cache.get(&float.rel_id);
                                commands.push(DrawCommand::Image {
                                    x: img_x,
                                    y: cell_y,
                                    width: img_w,
                                    height: img_h,
                                    image,
                                });
                                cell_y += img_h;
                            }

                            let fragments = collect_fragments(
                                &p.runs,
                                cell_content_width,
                                self.config.content_height(),
                                &self.doc_defaults,
                                &self.measurer,
                                &self.image_cache,
                            );

                            if fragments.is_empty() && p.floats.is_empty() {
                                let default_size =
                                    crate::dimension::Pt::from(self.doc_defaults.font_size).raw();
                                let natural = self.measurer.line_height(
                                    &self.doc_defaults.font_family,
                                    default_size,
                                    false,
                                    false,
                                );
                                cell_y += resolve_line_height(natural, spacing.line_spacing());
                                cell_y += spacing.after_pt();
                                continue;
                            }
                            if fragments.is_empty() {
                                cell_y += spacing.after_pt();
                                continue;
                            }

                            // Use shared measure_lines for fragment → command conversion
                            let measured = measure_lines(
                                &fragments,
                                cell_x + pad_left,
                                cell_content_width,
                                0.0, // no first-line indent in cells
                                p.properties.alignment,
                                spacing.line_spacing(),
                                &p.properties.tab_stops,
                                self.default_tab_stop_pt,
                                &self.image_cache,
                            );

                            // Offset measured commands by cell_y and accumulate
                            for line in &measured.lines {
                                for cmd in &line.commands {
                                    commands.push(offset_command(cmd, cell_y));
                                }
                            }
                            cell_y += measured.total_height;
                            cell_y += spacing.after_pt();
                        }
                    }

                    // Ensure minimum bottom padding to account for font descender space
                    // not captured by line height metrics.
                    let effective_pad_bottom = pad_bottom.max(1.0);
                    let content_height = cell_y + effective_pad_bottom;

                    measured_cells.push(MeasuredCell {
                        commands,
                        content_height,
                        col_width,
                    });
                }

                MeasuredRow {
                    cells: measured_cells,
                    col_x_positions,
                    min_height,
                }
            })
            .collect();

        // ============================
        // PASS 2: LAYOUT — compute row heights with vMerge distribution
        // ============================

        let mut vmerge_spans: Vec<VmergeSpan> = Vec::new();
        for (row_idx, row) in table.rows.iter().enumerate() {
            for (col_idx, cell) in row.cells.iter().enumerate() {
                if cell.vertical_merge == Some(VerticalMerge::Restart) {
                    let span_count = 1 + table.rows[row_idx + 1..]
                        .iter()
                        .take_while(|r| {
                            r.cells.get(col_idx).is_some_and(|c| c.is_vmerge_continue())
                        })
                        .count();
                    let content_height = measured_rows[row_idx].cells[col_idx].content_height;
                    vmerge_spans.push(VmergeSpan {
                        total_height: content_height,
                        start_row: row_idx,
                        row_count: span_count,
                    });
                }
            }
        }

        // Compute base row heights from non-merged cells only.
        let mut row_heights: Vec<f32> = measured_rows
            .iter()
            .enumerate()
            .map(|(row_idx, mr)| {
                let mut h = mr.min_height;
                for (col_idx, mc) in mr.cells.iter().enumerate() {
                    let cell = &table.rows[row_idx].cells[col_idx];
                    if cell.vertical_merge.is_none() && mc.content_height > h {
                        h = mc.content_height;
                    }
                }
                h
            })
            .collect();

        // Distribute vMerge span heights across rows
        for span in &vmerge_spans {
            let current_total: f32 = row_heights[span.start_row..span.start_row + span.row_count]
                .iter()
                .sum();
            if span.total_height > current_total {
                let deficit = span.total_height - current_total;
                let per_row = deficit / span.row_count as f32;
                for i in 0..span.row_count {
                    row_heights[span.start_row + i] += per_row;
                }
            }
        }

        // ============================
        // PASS 3: PAINT — emit draw commands at computed positions
        // ============================

        let tbl_borders = table
            .borders
            .unwrap_or(self.doc_defaults.default_table_borders);
        let num_rows = table.rows.len();

        for (row_idx, (mrow, row)) in measured_rows.iter().zip(table.rows.iter()).enumerate() {
            let row_height = row_heights[row_idx];

            if self.cursor_y + row_height > self.content_bottom() {
                self.new_page();
            }

            let row_top = self.cursor_y;
            let row_bottom = row_top + row_height;
            let num_cells = row.cells.len();

            for (col_idx, (mc, cell)) in mrow.cells.iter().zip(row.cells.iter()).enumerate() {
                let cell_x = mrow.col_x_positions[col_idx];
                let cw = mc.col_width;

                // 1. Cell shading
                if let Some(color) = &cell.shading {
                    self.current_page.commands.push(DrawCommand::Rect {
                        x: cell_x,
                        y: row_top,
                        width: cw,
                        height: row_height,
                        color: (color.r, color.g, color.b),
                    });
                }

                // 2. Cell content (offset by row_top)
                for cmd in &mc.commands {
                    self.current_page
                        .commands
                        .push(offset_command(cmd, row_top));
                }

                // 3. Cell borders
                let cell_b = cell.cell_borders.unwrap_or_default();
                let is_first_row = row_idx == 0;
                let is_last_row = row_idx == num_rows - 1;
                let is_first_col = col_idx == 0;
                let is_last_col = col_idx == num_cells - 1;

                if !cell.is_vmerge_continue() {
                    let tbl_top = if is_first_row {
                        tbl_borders.top
                    } else {
                        tbl_borders.inside_h
                    };
                    emit_border(
                        &mut self.current_page.commands,
                        &resolve_border(cell_b.top, tbl_top),
                        cell_x,
                        row_top,
                        cell_x + cw,
                        row_top,
                    );
                }

                let tbl_left = if is_first_col {
                    tbl_borders.left
                } else {
                    tbl_borders.inside_v
                };
                emit_border(
                    &mut self.current_page.commands,
                    &resolve_border(cell_b.left, tbl_left),
                    cell_x,
                    row_top,
                    cell_x,
                    row_bottom,
                );

                let tbl_right = if is_last_col {
                    tbl_borders.right
                } else {
                    tbl_borders.inside_v
                };
                emit_border(
                    &mut self.current_page.commands,
                    &resolve_border(cell_b.right, tbl_right),
                    cell_x + cw,
                    row_top,
                    cell_x + cw,
                    row_bottom,
                );

                let tbl_bottom = if is_last_row {
                    tbl_borders.bottom
                } else {
                    tbl_borders.inside_h
                };
                emit_border(
                    &mut self.current_page.commands,
                    &resolve_border(cell_b.bottom, tbl_bottom),
                    cell_x,
                    row_bottom,
                    cell_x + cw,
                    row_bottom,
                );
            }

            self.cursor_y += row_height;
        }

        // After a table, Word applies the document-default paragraph after-spacing.
        if !next_is_table {
            self.cursor_y += self.doc_defaults.default_spacing.after_pt();
        }
    }
}
