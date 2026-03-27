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
    /// §17.7.6: per-cell resolved borders from conditional formatting.
    pub cell_borders: Option<CellBorderConfig>,
}

/// Per-cell border configuration (resolved from conditional formatting).
/// Each field: `None` = no override (use table-level), `Some(None)` = nil
/// (explicitly no border), `Some(Some(line))` = specific border.
#[derive(Clone, Debug)]
pub struct CellBorderConfig {
    pub top: Option<Option<TableBorderLine>>,
    pub bottom: Option<Option<TableBorderLine>>,
    pub left: Option<Option<TableBorderLine>>,
    pub right: Option<Option<TableBorderLine>>,
}

/// Resolved table border configuration.
#[derive(Clone, Debug)]
pub struct TableBorderConfig {
    pub top: Option<TableBorderLine>,
    pub bottom: Option<TableBorderLine>,
    pub left: Option<TableBorderLine>,
    pub right: Option<TableBorderLine>,
    pub inside_h: Option<TableBorderLine>,
    pub inside_v: Option<TableBorderLine>,
}

/// A single table border line.
#[derive(Clone, Copy, Debug)]
pub struct TableBorderLine {
    pub width: Pt,
    pub color: RgbColor,
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
    borders: Option<&TableBorderConfig>,
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

        // §17.4.38 / §17.7.6: draw borders.
        // Per-cell borders (from conditional formatting) take priority over
        // table-level borders. Vertical borders extend to cover horizontal
        // border thickness at corners.
        {
            let mut grid_idx = 0;
            for (cell_ci, cell_input) in rows[row_idx].cells.iter().enumerate() {
                let span = cell_input.grid_span.max(1) as usize;
                let cell_x: Pt = (0..grid_idx)
                    .map(|i| col_widths.get(i).copied().unwrap_or(Pt::ZERO))
                    .sum();
                let cell_w: Pt = (grid_idx..grid_idx + span)
                    .map(|i| col_widths.get(i).copied().unwrap_or(Pt::ZERO))
                    .sum();
                grid_idx += span;

                // Resolve effective borders for this cell.
                let (b_top, b_bottom, b_left, b_right) =
                    resolve_cell_effective_borders(cell_input, borders, row_idx, cell_ci,
                        rows.len(), rows[row_idx].cells.len());

                // Horizontal borders.
                let h_top_half = b_top.as_ref().map(|b| b.width * 0.5).unwrap_or(Pt::ZERO);
                let h_bot_half = b_bottom.as_ref().map(|b| b.width * 0.5).unwrap_or(Pt::ZERO);

                if let Some(ref b) = b_top {
                    emit_h_border(&mut commands, b, cell_x, cell_x + cell_w, cursor_y);
                }
                if let Some(ref b) = b_bottom {
                    emit_h_border(&mut commands, b, cell_x, cell_x + cell_w, cursor_y + row_height);
                }

                // Vertical borders (extend for clean corners).
                let v_top = cursor_y - h_top_half;
                let v_bot = cursor_y + row_height + h_bot_half;
                if let Some(ref b) = b_left {
                    emit_v_border(&mut commands, b, cell_x, v_top, v_bot);
                }
                if let Some(ref b) = b_right {
                    emit_v_border(&mut commands, b, cell_x + cell_w, v_top, v_bot);
                }
            }
        }

        cursor_y += row_height;
    }

    // Bottom border is now drawn per-cell in the loop above.

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

/// §17.4.38 / §17.7.6: resolve effective borders for a cell.
/// Per-cell borders (from conditional formatting) override table-level borders.
/// Table-level insideH/insideV are mapped to cell edges based on position.
fn resolve_cell_effective_borders(
    cell: &TableCellInput,
    table_borders: Option<&TableBorderConfig>,
    row_idx: usize,
    col_idx: usize,
    num_rows: usize,
    num_cols: usize,
) -> (Option<TableBorderLine>, Option<TableBorderLine>, Option<TableBorderLine>, Option<TableBorderLine>) {
    // Start with table-level borders mapped to cell edges.
    let tb = table_borders;
    let is_first_row = row_idx == 0;
    let is_last_row = row_idx == num_rows - 1;
    let is_first_col = col_idx == 0;
    let is_last_col = col_idx == num_cols - 1;

    let mut top = if is_first_row { tb.and_then(|b| b.top) } else { tb.and_then(|b| b.inside_h) };
    let mut bottom = if is_last_row { tb.and_then(|b| b.bottom) } else { tb.and_then(|b| b.inside_h) };
    let mut left = if is_first_col { tb.and_then(|b| b.left) } else { tb.and_then(|b| b.inside_v) };
    let mut right = if is_last_col { tb.and_then(|b| b.right) } else { tb.and_then(|b| b.inside_v) };

    // Per-cell borders from conditional formatting override.
    // Some(None) = nil (explicitly remove border), Some(Some(line)) = override border.
    if let Some(ref cb) = cell.cell_borders {
        if let Some(v) = cb.top { top = v; }
        if let Some(v) = cb.bottom { bottom = v; }
        if let Some(v) = cb.left { left = v; }
        if let Some(v) = cb.right { right = v; }
    }

    (top, bottom, left, right)
}

fn emit_h_border(commands: &mut Vec<DrawCommand>, b: &TableBorderLine, x1: Pt, x2: Pt, y: Pt) {
    commands.push(DrawCommand::Line {
        line: PtLineSegment::new(PtOffset::new(x1, y), PtOffset::new(x2, y)),
        color: b.color,
        width: b.width,
    });
}

fn emit_v_border(commands: &mut Vec<DrawCommand>, b: &TableBorderLine, x: Pt, y1: Pt, y2: Pt) {
    commands.push(DrawCommand::Line {
        line: PtLineSegment::new(PtOffset::new(x, y1), PtOffset::new(x, y2)),
        color: b.color,
        width: b.width,
    });
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
            width: Pt::new(width), trimmed_width: Pt::new(width),
            height: Pt::new(14.0),
            ascent: Pt::new(10.0),
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO,
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
            cell_borders: None,
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
        let result = layout_table(&[], &[], &body_constraints(), Pt::new(14.0), None);
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
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

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
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

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
                    shading: None, cell_borders: None,
                },
            ],
            min_height: None,
        }];
        // Column B is only 80 wide, so "long " + "text" (120) wraps
        let col_widths = vec![Pt::new(200.0), Pt::new(80.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

        assert_eq!(result.size.height.raw(), 28.0, "row height = tallest cell");
    }

    #[test]
    fn min_row_height_respected() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("x")],
            min_height: Some(Pt::new(40.0)),
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

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
                cell_borders: None,
            }],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

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
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), Some(&super::TableBorderConfig { top: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK }), bottom: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK }), left: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK }), right: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK }), inside_h: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK }), inside_v: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK }) }));

        let line_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Line { .. }))
            .count();
        // Top border (1) + 3 vertical borders (left, middle, right) + bottom border (1) = 5
        // Per-cell borders: 2 cells × 4 borders = 8 lines (shared edges drawn by both cells).
        assert_eq!(line_count, 8);
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
                shading: None, cell_borders: None,
            }],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

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
                shading: None, cell_borders: None,
            }],
            min_height: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

        // Row height = content(14) + top(5) + bottom(5) = 24
        assert_eq!(result.size.height.raw(), 24.0);

        // Text should be offset by left margin
        if let Some(DrawCommand::Text { position, .. }) = result.commands.first() {
            assert_eq!(position.x.raw(), 10.0, "left margin");
        }
    }
}
