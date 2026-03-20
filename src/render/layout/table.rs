use crate::model::*;
use crate::units::*;

use super::fragment::*;
use super::{DrawCommand, Layouter};

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
    /// Column index where the span starts.
    _col_idx: usize,
    /// Total content height of the restart cell.
    total_height: f32,
    /// Row index where the span starts.
    start_row: usize,
    /// Number of rows in the span.
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
    cell.cell_margins
        .or(*table_default)
        .unwrap_or(*doc_default)
}

fn resolve_border(
    cell_border: Option<BorderDef>,
    table_border: BorderDef,
) -> BorderDef {
    cell_border.unwrap_or(table_border)
}

fn emit_border(
    commands: &mut Vec<DrawCommand>,
    border: &BorderDef,
    x1: f32, y1: f32, x2: f32, y2: f32,
) {
    if border.is_visible() {
        commands.push(DrawCommand::Line {
            x1, y1, x2, y2,
            color: border.color_rgb(),
            width: border.width_pt(),
        });
    }
}

fn offset_command(cmd: &DrawCommand, row_top: f32) -> DrawCommand {
    match cmd {
        DrawCommand::Text {
            x, y, text, font_family, char_spacing_pt, font_size,
            bold, italic, color,
        } => DrawCommand::Text {
            x: *x, y: row_top + y,
            text: text.clone(), font_family: font_family.clone(),
            char_spacing_pt: *char_spacing_pt,
            font_size: *font_size, bold: *bold, italic: *italic, color: *color,
        },
        DrawCommand::Underline { x1, y1, x2, y2, color, width } => DrawCommand::Underline {
            x1: *x1, y1: row_top + y1, x2: *x2, y2: row_top + y2,
            color: *color, width: *width,
        },
        DrawCommand::Image { x, y, width, height, data } => DrawCommand::Image {
            x: *x, y: row_top + y, width: *width, height: *height, data: data.clone(),
        },
        DrawCommand::Rect { x, y, width, height, color } => DrawCommand::Rect {
            x: *x, y: row_top + y, width: *width, height: *height, color: *color,
        },
        DrawCommand::Line { .. } => cmd.clone(),
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
            let scale = if grid_total > 0.0 { content_width / grid_total } else { 1.0 };
            table.grid_cols.iter().map(|w| twips_to_pt(*w) * scale).collect()
        } else {
            vec![content_width / num_cols as f32; num_cols]
        };

        // ============================
        // PASS 1: MEASURE all cells
        // ============================
        let measured_rows: Vec<MeasuredRow> = table.rows.iter().enumerate().map(|(_row_idx, row)| {
            let min_height = row.height.map(twips_to_pt).unwrap_or(MIN_ROW_HEIGHT_PT);
            let row_height_limit = self.config.content_height();

            let mut col_x_positions = Vec::with_capacity(row.cells.len());
            let mut cell_widths_computed = Vec::with_capacity(row.cells.len());
            let mut grid_col_idx = 0;
            for cell in &row.cells {
                let x = self.config.margin_left
                    + col_widths[..grid_col_idx].iter().sum::<f32>();
                let w = self.cell_width(grid_col_idx, cell, &col_widths);
                col_x_positions.push(x);
                cell_widths_computed.push(w);
                grid_col_idx += cell.grid_span.max(1) as usize;
            }

            let mut measured_cells = Vec::with_capacity(row.cells.len());

            for (col_idx, cell) in row.cells.iter().enumerate() {
                let col_width = cell_widths_computed[col_idx];
                let cell_x = col_x_positions[col_idx];
                let margins = resolve_cell_margins(cell, &table.default_cell_margins, &doc_cell_margins);
                let pad_left = margins.left_pt();
                let pad_right = margins.right_pt();
                let pad_top = margins.top_pt();
                let pad_bottom = margins.bottom_pt();
                let cell_content_width = (col_width - pad_left - pad_right).max(1.0);

                let mut commands = Vec::new();
                let mut cell_y = pad_top;

                if cell.is_vmerge_continue() {
                    measured_cells.push(MeasuredCell {
                        commands,
                        content_height: 0.0,
                        col_width,
                    });
                    continue;
                }

                // Measure cell content (same logic as before, but no cursor_y mutation)
                for block in &cell.blocks {
                    if let Block::Paragraph(p) = block {
                        let spacing = self.resolve_cell_spacing(
                            p.properties.spacing, table.cell_spacing,
                        );
                        cell_y += spacing.before_pt();

                        // Floating images in cell
                        for float in &p.floats {
                            if float.data.is_empty() { continue; }
                            let scale = f32::min(1.0, f32::min(
                                cell_content_width / float.width_pt.max(1.0),
                                row_height_limit.max(1.0) / float.height_pt.max(1.0),
                            ));
                            let img_w = float.width_pt * scale;
                            let img_h = float.height_pt * scale;
                            let img_x = cell_x + (col_width - img_w) / 2.0;
                            commands.push(DrawCommand::Image {
                                x: img_x, y: cell_y,
                                width: img_w, height: img_h,
                                data: float.data.clone(),
                            });
                            cell_y += img_h;
                        }

                        let fragments = collect_fragments(
                            &p.runs, cell_content_width,
                            self.config.content_height(),
                            &self.doc_defaults, &self.measurer,
                        );

                        if fragments.is_empty() && p.floats.is_empty() {
                            let default_size = self.doc_defaults.font_size_half_pts as f32
                                / HALF_POINTS_PER_POINT;
                            let natural = self.measurer.line_height(
                                &self.doc_defaults.font_family, default_size, false, false,
                            );
                            cell_y += resolve_line_height(natural, spacing.line_spacing());
                            cell_y += spacing.after_pt();
                            continue;
                        }
                        if fragments.is_empty() {
                            cell_y += spacing.after_pt();
                            continue;
                        }

                        let mut line_start = 0;
                        let mut is_first_line = true;
                        while line_start < fragments.len() {
                            if !is_first_line {
                                while line_start < fragments.len() {
                                    if let Fragment::Text { ref text, .. } = fragments[line_start] {
                                        if text.trim().is_empty() { line_start += 1; continue; }
                                    }
                                    break;
                                }
                                if line_start >= fragments.len() { break; }
                            }

                            let (line_end, _) = fit_fragments(
                                &fragments[line_start..], cell_content_width,
                            );
                            let actual_end = line_start + line_end.max(1);

                            let frag_height = fragments[line_start..actual_end]
                                .iter().map(|f| f.height()).fold(0.0_f32, f32::max);
                            let line_height = resolve_line_height(
                                frag_height, spacing.line_spacing(),
                            );
                            cell_y += line_height;

                            let used_width = measure_fragments(&fragments[line_start..actual_end]);
                            let align_offset = match p.properties.alignment {
                                Some(Alignment::Center) => (cell_content_width - used_width) / 2.0,
                                Some(Alignment::Right) => cell_content_width - used_width,
                                _ => 0.0,
                            };
                            let mut x = cell_x + pad_left + align_offset;

                            for frag in &fragments[line_start..actual_end] {
                                match frag {
                                    Fragment::Text {
                                        text, font_family, font_size, bold, italic,
                                        underline, color, shading, char_spacing_pt,
                                        measured_width, ..
                                    } => {
                                        let c = color.map(|c| (c.r, c.g, c.b)).unwrap_or((0, 0, 0));
                                        if let Some(bg) = shading {
                                            commands.push(DrawCommand::Rect {
                                                x, y: cell_y - line_height,
                                                width: *measured_width, height: line_height,
                                                color: (bg.r, bg.g, bg.b),
                                            });
                                        }
                                        commands.push(DrawCommand::Text {
                                            x, y: cell_y,
                                            text: text.clone(),
                                            font_family: font_family.clone(),
                                            char_spacing_pt: *char_spacing_pt,
                                            font_size: *font_size, bold: *bold, italic: *italic,
                                            color: c,
                                        });
                                        if *underline {
                                            commands.push(DrawCommand::Underline {
                                                x1: x, y1: cell_y + UNDERLINE_Y_OFFSET,
                                                x2: x + measured_width,
                                                y2: cell_y + UNDERLINE_Y_OFFSET,
                                                color: c, width: UNDERLINE_STROKE_WIDTH,
                                            });
                                        }
                                        x += measured_width;
                                    }
                                    Fragment::Image { width, height, data } => {
                                        commands.push(DrawCommand::Image {
                                            x, y: cell_y - height,
                                            width: *width, height: *height,
                                            data: data.clone(),
                                        });
                                        x += width;
                                    }
                                    Fragment::Tab { .. } => {
                                        let rel_x = x - (cell_x + pad_left);
                                        let next_stop = find_next_tab_stop(
                                            rel_x, &p.properties.tab_stops,
                                            self.default_tab_stop_pt,
                                        );
                                        x = cell_x + pad_left + next_stop;
                                    }
                                    Fragment::LineBreak { .. } => {}
                                }
                            }
                            line_start = actual_end;
                            is_first_line = false;
                        }
                        cell_y += spacing.after_pt();
                    }
                }

                let effective_pad_bottom = pad_bottom.max(MIN_CELL_BOTTOM_PAD_PT);
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
        }).collect();

        // ============================
        // PASS 2: LAYOUT — compute row heights with vMerge distribution
        // ============================

        // First, find all vMerge spans
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
                        _col_idx: col_idx,
                        total_height: content_height,
                        start_row: row_idx,
                        row_count: span_count,
                    });
                }
            }
        }

        // Compute base row heights from non-merged cells only.
        // Cells with vertical_merge (Restart or Continue) are handled by vMerge distribution.
        let mut row_heights: Vec<f32> = measured_rows.iter().enumerate().map(|(row_idx, mr)| {
            let mut h = mr.min_height;
            for (col_idx, mc) in mr.cells.iter().enumerate() {
                let cell = &table.rows[row_idx].cells[col_idx];
                if cell.vertical_merge.is_none() && mc.content_height > h {
                    h = mc.content_height;
                }
            }
            h
        }).collect();

        // Distribute vMerge span heights across rows
        for span in &vmerge_spans {
            let current_total: f32 = row_heights[span.start_row..span.start_row + span.row_count]
                .iter().sum();
            if span.total_height > current_total {
                // Need to add extra height. Distribute evenly.
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

        let tbl_borders = table.borders
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
                        x: cell_x, y: row_top, width: cw, height: row_height,
                        color: (color.r, color.g, color.b),
                    });
                }

                // 2. Cell content (offset by row_top)
                for cmd in &mc.commands {
                    self.current_page.commands.push(offset_command(cmd, row_top));
                }

                // 3. Cell borders
                let cell_b = cell.cell_borders.unwrap_or_default();
                let is_first_row = row_idx == 0;
                let is_last_row = row_idx == num_rows - 1;
                let is_first_col = col_idx == 0;
                let is_last_col = col_idx == num_cells - 1;

                if !cell.is_vmerge_continue() {
                    let tbl_top = if is_first_row { tbl_borders.top } else { tbl_borders.inside_h };
                    emit_border(&mut self.current_page.commands,
                        &resolve_border(cell_b.top, tbl_top),
                        cell_x, row_top, cell_x + cw, row_top);
                }

                let tbl_left = if is_first_col { tbl_borders.left } else { tbl_borders.inside_v };
                emit_border(&mut self.current_page.commands,
                    &resolve_border(cell_b.left, tbl_left),
                    cell_x, row_top, cell_x, row_bottom);

                let tbl_right = if is_last_col { tbl_borders.right } else { tbl_borders.inside_v };
                emit_border(&mut self.current_page.commands,
                    &resolve_border(cell_b.right, tbl_right),
                    cell_x + cw, row_top, cell_x + cw, row_bottom);

                let tbl_bottom = if is_last_row { tbl_borders.bottom } else { tbl_borders.inside_h };
                emit_border(&mut self.current_page.commands,
                    &resolve_border(cell_b.bottom, tbl_bottom),
                    cell_x, row_bottom, cell_x + cw, row_bottom);
            }

            self.cursor_y += row_height;
        }

        // After a table, Word applies the document-default paragraph after-spacing.
        if !next_is_table {
            self.cursor_y += self.doc_defaults.default_spacing.after_pt();
        }
    }
}
