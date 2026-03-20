use crate::model::*;
use crate::units::*;

use super::fragment::*;
use super::{DrawCommand, Layouter};

/// Resolve cell margins for a specific cell: per-cell > table default > document default.
fn resolve_cell_margins(
    cell: &TableCell,
    table_default: &Option<CellMargins>,
    doc_default: &CellMargins,
) -> CellMargins {
    cell.cell_margins
        .or(*table_default)
        .unwrap_or(*doc_default)
}

impl Layouter {
    pub(super) fn layout_table(&mut self, table: &Table) {
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
            let grid_total: f32 =
                table.grid_cols.iter().map(|w| twips_to_pt(*w)).sum();
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

        for row in &table.rows {
            let mut cell_layouts: Vec<Vec<DrawCommand>> = Vec::new();
            let mut row_height = MIN_ROW_HEIGHT_PT;

            let mut col_x_positions: Vec<f32> = Vec::with_capacity(row.cells.len());
            let mut cell_widths_computed: Vec<f32> =
                Vec::with_capacity(row.cells.len());
            let mut grid_col_idx = 0;
            for cell in &row.cells {
                let x = self.config.margin_left
                    + col_widths[..grid_col_idx].iter().sum::<f32>();
                let w = self.cell_width(grid_col_idx, cell, &col_widths);
                col_x_positions.push(x);
                cell_widths_computed.push(w);
                grid_col_idx += cell.grid_span.max(1) as usize;
            }

            for (col_idx, cell) in row.cells.iter().enumerate() {
                let cell_x = col_x_positions[col_idx];
                let col_width = cell_widths_computed[col_idx];
                let margins = resolve_cell_margins(
                    cell,
                    &table.default_cell_margins,
                    &doc_cell_margins,
                );
                let pad_left = margins.left_pt();
                let pad_right = margins.right_pt();
                let pad_top = margins.top_pt();
                let pad_bottom = margins.bottom_pt();
                let cell_content_width =
                    (col_width - pad_left - pad_right).max(1.0);
                let mut commands = Vec::new();
                let mut cell_y = pad_top;

                if cell.is_vmerge_continue() {
                    cell_layouts.push(commands);
                    continue;
                }

                for block in &cell.blocks {
                    if let Block::Paragraph(p) = block {
                        let spacing = self.resolve_cell_spacing(
                            p.properties.spacing,
                            table.cell_spacing,
                        );
                        cell_y += spacing.before_pt();

                        for float in &p.floats {
                            if float.data.is_empty() {
                                continue;
                            }
                            let scale = f32::min(
                                1.0,
                                f32::min(
                                    cell_content_width / float.width_pt.max(1.0),
                                    self.config.content_height()
                                        / float.height_pt.max(1.0),
                                ),
                            );
                            let img_w = float.width_pt * scale;
                            let img_h = float.height_pt * scale;
                            let img_x = if float.offset_x_pt > 0.0
                                && float.offset_x_pt + img_w <= cell_content_width
                            {
                                cell_x + pad_left + float.offset_x_pt
                            } else {
                                cell_x + (col_width - img_w) / 2.0
                            };
                            commands.push(DrawCommand::Image {
                                x: img_x,
                                y: cell_y,
                                width: img_w,
                                height: img_h,
                                data: float.data.clone(),
                            });
                            cell_y += img_h;
                        }

                        let fragments = collect_fragments(
                            &p.runs,
                            cell_content_width,
                            self.config.content_height(),
                            &self.doc_defaults,
                            &self.measurer,
                        );

                        if fragments.is_empty() && p.floats.is_empty() {
                            cell_y += spacing.line_pt();
                            cell_y += spacing.after_pt();
                            continue;
                        }

                        if fragments.is_empty() {
                            cell_y += spacing.after_pt();
                            continue;
                        }

                        let mut line_start = 0;
                        while line_start < fragments.len() {
                            let (line_end, _) = fit_fragments(
                                &fragments[line_start..],
                                cell_content_width,
                            );
                            let actual_end = line_start + line_end.max(1);

                            let line_height = fragments[line_start..actual_end]
                                .iter()
                                .map(|f| f.height())
                                .fold(spacing.line_pt(), f32::max);

                            cell_y += line_height;

                            let used_width = measure_fragments(
                                &fragments[line_start..actual_end],
                            );
                            let align_offset = match p.properties.alignment {
                                Some(Alignment::Center) => {
                                    (cell_content_width - used_width) / 2.0
                                }
                                Some(Alignment::Right) => {
                                    cell_content_width - used_width
                                }
                                _ => 0.0,
                            };
                            let mut x = cell_x + pad_left + align_offset;

                            for frag in &fragments[line_start..actual_end] {
                                match frag {
                                    Fragment::Text {
                                        text,
                                        font_family,
                                        font_size,
                                        bold,
                                        italic,
                                        underline,
                                        color,
                                        measured_width,
                                        ..
                                    } => {
                                        let c = color
                                            .map(|c| (c.r, c.g, c.b))
                                            .unwrap_or((0, 0, 0));

                                        commands.push(DrawCommand::Text {
                                            x,
                                            y: cell_y,
                                            text: text.clone(),
                                            font_family: font_family.clone(),
                                            font_size: *font_size,
                                            bold: *bold,
                                            italic: *italic,
                                            color: c,
                                        });

                                        if *underline {
                                            commands.push(
                                                DrawCommand::Underline {
                                                    x1: x,
                                                    y1: cell_y + UNDERLINE_Y_OFFSET,
                                                    x2: x + measured_width,
                                                    y2: cell_y + UNDERLINE_Y_OFFSET,
                                                    color: c,
                                                    width: TABLE_BORDER_WIDTH,
                                                },
                                            );
                                        }

                                        x += measured_width;
                                    }
                                    Fragment::Image {
                                        width,
                                        height,
                                        data,
                                    } => {
                                        commands.push(DrawCommand::Image {
                                            x,
                                            y: cell_y - height,
                                            width: *width,
                                            height: *height,
                                            data: data.clone(),
                                        });
                                        x += width;
                                    }
                                    Fragment::Tab { .. } => {
                                        let rel_x = x - (cell_x + pad_left);
                                        let next_stop = find_next_tab_stop(
                                            rel_x,
                                            &p.properties.tab_stops,
                                            self.default_tab_stop_pt,
                                        );
                                        x = cell_x + pad_left + next_stop;
                                    }
                                    Fragment::LineBreak { .. } => {}
                                }
                            }

                            line_start = actual_end;
                        }

                        cell_y += spacing.after_pt();
                    }
                }

                // Ensure minimum bottom padding to prevent text descenders
                // from overlapping the cell border
                let effective_pad_bottom = pad_bottom.max(MIN_CELL_BOTTOM_PAD_PT);
                let total_cell_height = cell_y + effective_pad_bottom;
                if total_cell_height > row_height {
                    row_height = total_cell_height;
                }
                cell_layouts.push(commands);
            }

            if self.cursor_y + row_height > self.content_bottom() {
                self.new_page();
            }

            let row_top = self.cursor_y;

            // Emit cell borders and content
            for (col_idx, commands) in cell_layouts.iter().enumerate() {
                let cell_x = col_x_positions[col_idx];
                let cw = cell_widths_computed[col_idx];
                let cell = &row.cells[col_idx];
                let row_bottom = row_top + row_height;

                if !cell.is_vmerge_continue() {
                    self.current_page.commands.push(DrawCommand::Line {
                        x1: cell_x,
                        y1: row_top,
                        x2: cell_x + cw,
                        y2: row_top,
                        color: (0, 0, 0),
                        width: TABLE_BORDER_WIDTH,
                    });
                }

                self.current_page.commands.push(DrawCommand::Line {
                    x1: cell_x,
                    y1: row_top,
                    x2: cell_x,
                    y2: row_bottom,
                    color: (0, 0, 0),
                    width: 0.5,
                });

                self.current_page.commands.push(DrawCommand::Line {
                    x1: cell_x + cw,
                    y1: row_top,
                    x2: cell_x + cw,
                    y2: row_bottom,
                    color: (0, 0, 0),
                    width: 0.5,
                });

                self.current_page.commands.push(DrawCommand::Line {
                    x1: cell_x,
                    y1: row_bottom,
                    x2: cell_x + cw,
                    y2: row_bottom,
                    color: (0, 0, 0),
                    width: 0.5,
                });

                for cmd in commands {
                    let adjusted = offset_command(cmd, row_top);
                    self.current_page.commands.push(adjusted);
                }
            }

            self.cursor_y += row_height;
        }

        self.cursor_y += TABLE_AFTER_SPACING_PT;
    }
}

/// Offset a draw command's y coordinates by `row_top`.
fn offset_command(cmd: &DrawCommand, row_top: f32) -> DrawCommand {
    match cmd {
        DrawCommand::Text {
            x,
            y,
            text,
            font_family,
            font_size,
            bold,
            italic,
            color,
        } => DrawCommand::Text {
            x: *x,
            y: row_top + y,
            text: text.clone(),
            font_family: font_family.clone(),
            font_size: *font_size,
            bold: *bold,
            italic: *italic,
            color: *color,
        },
        DrawCommand::Underline {
            x1,
            y1,
            x2,
            y2,
            color,
            width,
        } => DrawCommand::Underline {
            x1: *x1,
            y1: row_top + y1,
            x2: *x2,
            y2: row_top + y2,
            color: *color,
            width: *width,
        },
        DrawCommand::Image {
            x,
            y,
            width,
            height,
            data,
        } => DrawCommand::Image {
            x: *x,
            y: row_top + y,
            width: *width,
            height: *height,
            data: data.clone(),
        },
        DrawCommand::Line { .. } => cmd.clone(),
    }
}
