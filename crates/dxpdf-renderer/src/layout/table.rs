//! Table layout — 3-pass column sizing, cell layout, border rendering.
//!
//! Pass 1: Compute column widths from grid definitions or equal distribution.
//! Pass 2: Lay out each cell with tight width constraints, determine row heights.
//! Pass 3: Position cells and emit border commands.

use crate::dimension::Pt;
use crate::geometry::{PtEdgeInsets, PtLineSegment, PtOffset, PtRect, PtSize};
use crate::resolve::color::RgbColor;

use super::cell::{layout_cell, CellBlock, CellLayout};
use super::draw_command::DrawCommand;
use super::BoxConstraints;

/// A table row for layout.
pub struct TableRowInput {
    pub cells: Vec<TableCellInput>,
    /// Minimum row height (from trHeight).
    pub min_height: Option<Pt>,
}

/// A single cell for layout.
pub struct TableCellInput {
    pub blocks: Vec<CellBlock>,
    pub margins: PtEdgeInsets,
    /// Number of grid columns this cell spans (gridSpan, default 1).
    pub grid_span: u32,
    /// Background color for cell shading.
    pub shading: Option<RgbColor>,
}

/// Result of laying out a table.
#[derive(Debug)]
pub struct TableLayout {
    /// Draw commands positioned relative to the table's top-left origin.
    pub commands: Vec<DrawCommand>,
    /// Total size of the table.
    pub size: PtSize,
}

/// Lay out a table: compute column widths, lay out cells, emit borders.
pub fn layout_table(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    _constraints: &BoxConstraints,
    default_line_height: Pt,
    draw_borders: bool,
) -> TableLayout {
    if rows.is_empty() || col_widths.is_empty() {
        return TableLayout {
            commands: Vec::new(),
            size: PtSize::ZERO,
        };
    }

    let table_width: Pt = col_widths.iter().copied().sum();
    let mut commands = Vec::new();
    let mut cursor_y = Pt::ZERO;
    let mut row_heights = Vec::with_capacity(rows.len());

    // Pass 2: lay out each cell, determine row heights.
    let mut row_cell_layouts: Vec<Vec<(CellLayout, Pt, Pt)>> = Vec::new(); // (layout, x, width)

    for row in rows {
        let mut cell_layouts = Vec::new();
        let mut max_height = Pt::ZERO;
        let mut grid_idx = 0;

        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let cell_width: Pt = (grid_idx..grid_idx + span)
                .map(|i| col_widths.get(i).copied()// §17.4.14: missing grid column treated as zero width.
                .unwrap_or(Pt::ZERO))
                .sum();
            let cell_x: Pt = (0..grid_idx)
                .map(|i| col_widths.get(i).copied().unwrap_or(Pt::ZERO))
                .sum();

            let cell_layout = layout_cell(
                &cell.blocks,
                cell_width,
                &cell.margins,
                default_line_height,
            );

            let total_cell_height =
                cell_layout.content_height + cell.margins.vertical();
            max_height = max_height.max(total_cell_height);

            cell_layouts.push((cell_layout, cell_x, cell_width));
            grid_idx += span;
        }

        // Apply minimum row height
        if let Some(min_h) = row.min_height {
            max_height = max_height.max(min_h);
        }

        row_heights.push(max_height);
        row_cell_layouts.push(cell_layouts);
    }

    // Pass 3: position cells, emit shading and commands.
    for (row_idx, (_row, cell_layouts)) in rows.iter().zip(row_cell_layouts.iter()).enumerate() {
        let row_height = row_heights[row_idx];

        for ((cell_layout, cell_x, cell_width), cell_input) in
            cell_layouts.iter().zip(rows[row_idx].cells.iter())
        {
            // Cell shading
            if let Some(color) = cell_input.shading {
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(*cell_x, cursor_y, *cell_width, row_height),
                    color,
                });
            }

            // Cell content commands — offset by cell position + row position
            for cmd in &cell_layout.commands {
                let mut cmd = cmd.clone();
                cmd.shift_y(cursor_y);
                shift_x(&mut cmd, *cell_x);
                commands.push(cmd);
            }
        }

        // Borders — draw at cell boundaries, not grid column boundaries.
        if draw_borders {
            // Top border of this row
            commands.push(DrawCommand::Line {
                line: PtLineSegment::new(
                    PtOffset::new(Pt::ZERO, cursor_y),
                    PtOffset::new(table_width, cursor_y),
                ),
                // §17.4.38: default border color is auto (black) per spec.
                // §17.4.38: default border width is 0.5pt (4 eighths of a point).
                color: RgbColor::BLACK,
                width: Pt::new(0.5),
            });

            // Left edge
            commands.push(DrawCommand::Line {
                line: PtLineSegment::new(
                    PtOffset::new(Pt::ZERO, cursor_y),
                    PtOffset::new(Pt::ZERO, cursor_y + row_height),
                ),
                // §17.4.38: default border color is auto (black) per spec.
                // §17.4.38: default border width is 0.5pt (4 eighths of a point).
                color: RgbColor::BLACK,
                width: Pt::new(0.5),
            });

            // Vertical borders at each cell's right edge
            let mut grid_idx = 0;
            for cell_input in &rows[row_idx].cells {
                let span = cell_input.grid_span.max(1) as usize;
                grid_idx += span;
                let vx: Pt = (0..grid_idx)
                    .map(|i| col_widths.get(i).copied().unwrap_or(Pt::ZERO))
                    .sum();
                commands.push(DrawCommand::Line {
                    line: PtLineSegment::new(
                        PtOffset::new(vx, cursor_y),
                        PtOffset::new(vx, cursor_y + row_height),
                    ),
                    color: RgbColor::BLACK,
                    width: Pt::new(0.5),
                });
            }
        }

        cursor_y += row_height;
    }

    // Bottom border
    if draw_borders {
        commands.push(DrawCommand::Line {
            line: PtLineSegment::new(
                PtOffset::new(Pt::ZERO, cursor_y),
                PtOffset::new(table_width, cursor_y),
            ),
            color: RgbColor::BLACK,
            width: Pt::new(0.5),
        });
    }

    TableLayout {
        commands,
        size: PtSize::new(table_width, cursor_y),
    }
}

/// Compute column widths by scaling grid column values to fit available width.
/// If `grid_cols` is empty, distributes equally.
pub fn compute_column_widths(grid_cols: &[Pt], num_cols: usize, available_width: Pt) -> Vec<Pt> {
    if !grid_cols.is_empty() {
        let total: Pt = grid_cols.iter().copied().sum();
        let scale = if total > Pt::ZERO {
            available_width / total
        } else {
            1.0
        };
        grid_cols.iter().map(|w| *w * scale).collect()
    } else if num_cols > 0 {
        vec![available_width / num_cols as f32; num_cols]
    } else {
        vec![]
    }
}

fn shift_x(cmd: &mut DrawCommand, dx: Pt) {
    match cmd {
        DrawCommand::Text { position, .. } => position.x += dx,
        DrawCommand::Underline { line, .. } => {
            line.start.x += dx;
            line.end.x += dx;
        }
        DrawCommand::Image { rect, .. } => rect.origin.x += dx,
        DrawCommand::Rect { rect, .. } => rect.origin.x += dx,
        DrawCommand::LinkAnnotation { rect, .. } => rect.origin.x += dx,
        DrawCommand::Line { line, .. } => {
            line.start.x += dx;
            line.end.x += dx;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::layout::fragment::{FontProps, Fragment};
    use crate::layout::paragraph::ParagraphStyle;
    use std::rc::Rc;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            text: text.into(),
            font: FontProps {
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: false,
                italic: false,
                underline: false,
                char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width),
            height: Pt::new(14.0),
            ascent: Pt::new(10.0),
            hyperlink_url: None,
            shading: None, baseline_offset: Pt::ZERO,
        }
    }

    fn simple_cell(text: &str) -> TableCellInput {
        TableCellInput {
            blocks: vec![CellBlock {
                fragments: vec![text_frag(text, 30.0)],
                style: ParagraphStyle::default(),
            }],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
        }
    }

    fn body_constraints() -> BoxConstraints {
        BoxConstraints::loose(PtSize::new(Pt::new(400.0), Pt::new(1000.0)))
    }

    // ── compute_column_widths ────────────────────────────────────────────

    #[test]
    fn equal_distribution_when_no_grid() {
        let widths = compute_column_widths(&[], 3, Pt::new(300.0));
        assert_eq!(widths.len(), 3);
        assert_eq!(widths[0].raw(), 100.0);
        assert_eq!(widths[1].raw(), 100.0);
        assert_eq!(widths[2].raw(), 100.0);
    }

    #[test]
    fn grid_cols_scaled_to_fit() {
        let grid = vec![Pt::new(100.0), Pt::new(200.0)];
        let widths = compute_column_widths(&grid, 2, Pt::new(600.0));
        // Scale = 600/300 = 2.0
        assert_eq!(widths[0].raw(), 200.0);
        assert_eq!(widths[1].raw(), 400.0);
    }

    #[test]
    fn grid_cols_already_fit() {
        let grid = vec![Pt::new(150.0), Pt::new(150.0)];
        let widths = compute_column_widths(&grid, 2, Pt::new(300.0));
        assert_eq!(widths[0].raw(), 150.0);
        assert_eq!(widths[1].raw(), 150.0);
    }

    #[test]
    fn zero_cols_empty_result() {
        let widths = compute_column_widths(&[], 0, Pt::new(300.0));
        assert!(widths.is_empty());
    }

    // ── layout_table ─────────────────────────────────────────────────────

    #[test]
    fn empty_table() {
        let result = layout_table(&[], &[], &body_constraints(), Pt::new(14.0), false);
        assert!(result.commands.is_empty());
        assert_eq!(result.size, PtSize::ZERO);
    }

    #[test]
    fn single_cell_table() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("hello")],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        assert_eq!(result.size.width.raw(), 200.0);
        assert_eq!(result.size.height.raw(), 14.0);

        let text_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    #[test]
    fn two_by_two_table() {
        let rows = vec![
            TableRowInput {
                cells: vec![simple_cell("a"), simple_cell("b")],
                min_height: None,
            },
            TableRowInput {
                cells: vec![simple_cell("c"), simple_cell("d")],
                min_height: None,
            },
        ];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        assert_eq!(result.size.width.raw(), 200.0);
        assert_eq!(result.size.height.raw(), 28.0); // 2 rows * 14pt

        let text_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 4);
    }

    #[test]
    fn row_height_is_max_of_cells() {
        // Cell A has 1 line (14pt), Cell B has 2 lines (28pt) because text wraps
        let rows = vec![TableRowInput {
            cells: vec![
                simple_cell("short"),
                TableCellInput {
                    blocks: vec![CellBlock {
                        fragments: vec![text_frag("long ", 60.0), text_frag("text", 60.0)],
                        style: ParagraphStyle::default(),
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                },
            ],
            min_height: None,
        }];
        // Column B is only 80 wide, so "long " + "text" (120) wraps
        let col_widths = vec![Pt::new(200.0), Pt::new(80.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        assert_eq!(result.size.height.raw(), 28.0, "row height = tallest cell");
    }

    #[test]
    fn min_row_height_respected() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("x")],
            min_height: Some(Pt::new(40.0)),
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        assert_eq!(result.size.height.raw(), 40.0, "min height > content height");
    }

    #[test]
    fn cell_shading_emits_rect() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![CellBlock {
                    fragments: vec![text_frag("x", 10.0)],
                    style: ParagraphStyle::default(),
                }],
                margins: PtEdgeInsets::ZERO,
                grid_span: 1,
                shading: Some(RgbColor { r: 200, g: 200, b: 200 }),
            }],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        let rect_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Rect { .. }))
            .count();
        assert_eq!(rect_count, 1, "shading produces a Rect command");
    }

    #[test]
    fn borders_emit_lines() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("a"), simple_cell("b")],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), true);

        let line_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Line { .. }))
            .count();
        // Top border (1) + 3 vertical borders (left, middle, right) + bottom border (1) = 5
        assert_eq!(line_count, 5);
    }

    #[test]
    fn grid_span_widens_cell() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![CellBlock {
                    fragments: vec![text_frag("spanning", 30.0)],
                    style: ParagraphStyle::default(),
                }],
                margins: PtEdgeInsets::ZERO,
                grid_span: 2, // spans both columns
                shading: None,
            }],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        // Cell gets full 200pt width, text should still render
        assert_eq!(result.size.width.raw(), 200.0);
        let text_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    #[test]
    fn cell_margins_affect_layout() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![CellBlock {
                    fragments: vec![text_frag("text", 30.0)],
                    style: ParagraphStyle::default(),
                }],
                margins: PtEdgeInsets::new(Pt::new(5.0), Pt::new(10.0), Pt::new(5.0), Pt::new(10.0)),
                grid_span: 1,
                shading: None,
            }],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), false);

        // Row height = content(14) + top(5) + bottom(5) = 24
        assert_eq!(result.size.height.raw(), 24.0);

        // Text should be offset by left margin
        if let Some(DrawCommand::Text { position, .. }) = result.commands.first() {
            assert_eq!(position.x.raw(), 10.0, "left margin");
        }
    }
}
