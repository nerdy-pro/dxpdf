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

/// §17.4.81: row height rule.
#[derive(Clone, Copy, Debug)]
pub enum RowHeightRule {
    /// Row height is at least this value; grows to fit content.
    AtLeast(Pt),
    /// Row height is exactly this value; content may clip.
    Exact(Pt),
}

/// A table row for layout.
pub struct TableRowInput {
    pub cells: Vec<TableCellInput>,
    /// §17.4.81: row height constraint.
    pub height_rule: Option<RowHeightRule>,
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
    /// §17.4.85: vertical merge state.
    pub vertical_merge: Option<VerticalMergeState>,
}

/// §17.4.85: vertical merge state for a cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalMergeState {
    /// This cell starts a new vertical merge group.
    Restart,
    /// This cell continues from the cell above (content is skipped).
    Continue,
}

/// §17.7.6: a conditional border override for a single cell edge.
#[derive(Clone, Copy, Debug)]
pub enum CellBorderOverride {
    /// §17.4.38 val="nil": explicitly no border on this edge.
    Nil,
    /// A specific border line on this edge.
    Border(TableBorderLine),
}

/// Per-cell border configuration (resolved from conditional formatting).
/// `None` = no override (use table-level default for this edge).
#[derive(Clone, Debug)]
pub struct CellBorderConfig {
    pub top: Option<CellBorderOverride>,
    pub bottom: Option<CellBorderOverride>,
    pub left: Option<CellBorderOverride>,
    pub right: Option<CellBorderOverride>,
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
    /// §17.4.38: border style (single, double, etc.)
    pub style: TableBorderStyle,
}

/// Supported table border styles.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableBorderStyle {
    Single,
    Double,
}

/// Per-cell layout result with positioning info from pass 2.
struct CellLayoutEntry {
    layout: CellLayout,
    cell_x: Pt,
    cell_w: Pt,
    /// Starting grid column index (for vMerge neighbor lookup).
    grid_col: usize,
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
    let mut row_cell_layouts: Vec<Vec<CellLayoutEntry>> = Vec::new();

    for row in rows {
        let mut entries = Vec::new();
        let mut max_height = Pt::ZERO;
        let mut grid_idx = 0;

        for cell in &row.cells {
            let span = cell.grid_span.max(1) as usize;
            let cell_w: Pt = col_widths[grid_idx..grid_idx + span].iter().copied().sum();
            let cell_x: Pt = col_widths[..grid_idx].iter().copied().sum();

            // §17.4.85: vMerge=continue cells have no content.
            let is_continue = cell.vertical_merge == Some(VerticalMergeState::Continue);
            let layout = if is_continue {
                CellLayout { commands: Vec::new(), content_height: Pt::ZERO }
            } else {
                layout_cell(&cell.blocks, cell_w, &cell.margins, default_line_height)
            };

            // §17.4.85: Continue cells don't contribute to row height.
            if !is_continue {
                max_height = max_height.max(layout.content_height + cell.margins.vertical());
            }

            entries.push(CellLayoutEntry { layout, cell_x, cell_w, grid_col: grid_idx });
            grid_idx += span;
        }

        match row.height_rule {
            Some(RowHeightRule::AtLeast(min_h)) => {
                max_height = max_height.max(min_h);
            }
            Some(RowHeightRule::Exact(h)) => {
                max_height = h;
            }
            None => {}
        }

        row_heights.push(max_height);
        row_cell_layouts.push(entries);
    }

    // §17.4.85: expand the last row of each vertical merge group so the
    // Restart cell's content fits within the combined spanned row heights.
    expand_rows_for_vmerge(rows, &row_cell_layouts, &mut row_heights);

    // Pass 3: resolve per-cell borders, then apply §17.4.43 conflict
    // resolution at shared edges so each edge is drawn exactly once.
    let num_rows = rows.len();

    // 3a: resolve raw borders for every cell.
    let cell_borders: Vec<Vec<CellBorders>> = rows
        .iter()
        .enumerate()
        .map(|(row_idx, row)| {
            row.cells
                .iter()
                .enumerate()
                .zip(row_cell_layouts[row_idx].iter())
                .map(|((cell_ci, cell_input), entry)| {
                    let (mut b_top, mut b_bottom, b_left, b_right) =
                        resolve_cell_effective_borders(
                            cell_input, borders, row_idx, cell_ci,
                            num_rows, row.cells.len(),
                        );
                    // §17.4.85: suppress horizontal borders inside vMerge groups.
                    if cell_input.vertical_merge == Some(VerticalMergeState::Continue) {
                        b_top = None;
                    }
                    if row_idx + 1 < num_rows
                        && is_vmerge_continue(&rows[row_idx + 1], entry.grid_col)
                    {
                        b_bottom = None;
                    }
                    CellBorders { top: b_top, bottom: b_bottom, left: b_left, right: b_right }
                })
                .collect()
        })
        .collect();

    // 3b: §17.4.43 conflict resolution at shared edges.
    // Horizontal interior edges: cell[r].bottom vs cell[r+1].top → winner kept
    // on cell[r].bottom, cell[r+1].top set to None.
    // Vertical interior edges: cell[c].right vs cell[c+1].left → winner kept
    // on cell[c].right, cell[c+1].left set to None.
    let mut resolved_borders = cell_borders;

    for row_idx in 0..num_rows {
        let num_cells = rows[row_idx].cells.len();
        for cell_ci in 0..num_cells {
            // Vertical: resolve right vs next cell's left.
            if cell_ci + 1 < num_cells {
                let right = resolved_borders[row_idx][cell_ci].right;
                let left = resolved_borders[row_idx][cell_ci + 1].left;
                let winner = resolve_border_conflict(right, left);
                resolved_borders[row_idx][cell_ci].right = winner;
                resolved_borders[row_idx][cell_ci + 1].left = None;
            }
            // Horizontal: resolve bottom vs next row's top.
            if row_idx + 1 < num_rows {
                // Find the cell in the next row at the same grid column.
                let grid_col = row_cell_layouts[row_idx][cell_ci].grid_col;
                if let Some(below_ci) = cell_index_at_grid_col(&rows[row_idx + 1], grid_col) {
                    let bottom = resolved_borders[row_idx][cell_ci].bottom;
                    let top = resolved_borders[row_idx + 1][below_ci].top;
                    let winner = resolve_border_conflict(bottom, top);
                    resolved_borders[row_idx][cell_ci].bottom = winner;
                    resolved_borders[row_idx + 1][below_ci].top = None;
                }
            }
        }
    }

    // 3c: emit in layered order — shading below content below borders.
    // This prevents shading rectangles from overlapping adjacent borders.
    let mut content_commands = Vec::new();
    let mut border_commands = Vec::new();

    for (row_idx, entries) in row_cell_layouts.iter().enumerate() {
        let row_height = row_heights[row_idx];

        for (cell_ci, (entry, cell_input)) in
            entries.iter().zip(rows[row_idx].cells.iter()).enumerate()
        {
            // Layer 1: shading (bottom).
            if let Some(color) = cell_input.shading {
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(entry.cell_x, cursor_y, entry.cell_w, row_height),
                    color,
                });
            }
            // Layer 2: content (middle).
            for cmd in &entry.layout.commands {
                let mut cmd = cmd.clone();
                cmd.shift(entry.cell_x, cursor_y);
                content_commands.push(cmd);
            }
            // Layer 3: borders (top).
            let b = &resolved_borders[row_idx][cell_ci];
            emit_cell_borders(
                &mut border_commands,
                CellBorders { top: b.top, bottom: b.bottom, left: b.left, right: b.right },
                entry.cell_x, entry.cell_w, cursor_y, row_height,
            );
        }

        cursor_y += row_height;
    }

    commands.append(&mut content_commands);
    commands.append(&mut border_commands);

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
    if let Some(ref cb) = cell.cell_borders {
        if let Some(v) = &cb.top { top = resolve_override(v); }
        if let Some(v) = &cb.bottom { bottom = resolve_override(v); }
        if let Some(v) = &cb.left { left = resolve_override(v); }
        if let Some(v) = &cb.right { right = resolve_override(v); }
    }

    (top, bottom, left, right)
}

/// §17.4.85: expand the last row of each vertical merge group so the
/// Restart cell's content fits within the combined spanned row heights.
fn expand_rows_for_vmerge(
    rows: &[TableRowInput],
    row_cell_layouts: &[Vec<CellLayoutEntry>],
    row_heights: &mut [Pt],
) {
    for (row_idx, row) in rows.iter().enumerate() {
        for (cell_ci, cell) in row.cells.iter().enumerate() {
            if cell.vertical_merge != Some(VerticalMergeState::Restart) {
                continue;
            }

            let entry = &row_cell_layouts[row_idx][cell_ci];
            let content_h = entry.layout.content_height + cell.margins.vertical();

            // Find last row in this merge group.
            let mut last_merged_row = row_idx;
            for (r, row_below) in rows.iter().enumerate().skip(row_idx + 1) {
                if is_vmerge_continue(row_below, entry.grid_col) {
                    last_merged_row = r;
                } else {
                    break;
                }
            }
            if last_merged_row == row_idx {
                continue;
            }

            // Expand last row if content overflows.
            let spanned: Pt = row_heights[row_idx..=last_merged_row].iter().copied().sum();
            if content_h > spanned {
                row_heights[last_merged_row] += content_h - spanned;
            }
        }
    }
}

/// Check if the cell at `grid_col` in `row` is a vMerge Continue cell.
fn is_vmerge_continue(row: &TableRowInput, grid_col: usize) -> bool {
    find_cell_at_grid_col(row, grid_col)
        .is_some_and(|c| c.vertical_merge == Some(VerticalMergeState::Continue))
}

/// Find the cell in a row that covers the given grid column index.
fn find_cell_at_grid_col(row: &TableRowInput, target_grid_col: usize) -> Option<&TableCellInput> {
    let mut col = 0;
    for cell in &row.cells {
        let span = cell.grid_span.max(1) as usize;
        if target_grid_col < col + span {
            return Some(cell);
        }
        col += span;
    }
    None
}

/// Return the cell index (not grid column) for the cell covering `grid_col`.
fn cell_index_at_grid_col(row: &TableRowInput, target_grid_col: usize) -> Option<usize> {
    let mut col = 0;
    for (i, cell) in row.cells.iter().enumerate() {
        let span = cell.grid_span.max(1) as usize;
        if target_grid_col < col + span {
            return Some(i);
        }
        col += span;
    }
    None
}

/// §17.4.43: resolve a border conflict between two competing borders on
/// a shared edge.  Returns the winning border (or `None` if both are `None`).
///
/// Algorithm per [MS-OI29500] §17.4.66:
///   1. `none` yields to the opposing border; `nil` suppresses both.
///   2. Weight = border_width_eighths × border_number.  Higher wins.
///   3. Equal weight: style precedence list.
///   4. Equal style: darker color wins (R+B+2G, then B+2G, then G).
fn resolve_border_conflict(
    a: Option<TableBorderLine>,
    b: Option<TableBorderLine>,
) -> Option<TableBorderLine> {
    match (a, b) {
        (None, b) => b,
        (a, None) => a,
        (Some(a), Some(b)) => Some(if border_weight(&b) > border_weight(&a) { b } else { a }),
    }
}

/// §17.4.43: compute border weight = width_eighths × style_number.
/// We only support Single (1) and Double (3).
fn border_weight(b: &TableBorderLine) -> f32 {
    let eighths = b.width.raw() * 8.0; // width is in points, sz is eighths
    let style_number = match b.style {
        TableBorderStyle::Single => 1.0,
        TableBorderStyle::Double => 3.0,
    };
    eighths * style_number
}

fn resolve_override(ovr: &CellBorderOverride) -> Option<TableBorderLine> {
    match ovr {
        CellBorderOverride::Nil => None,
        CellBorderOverride::Border(line) => Some(*line),
    }
}

/// Resolved borders for one cell edge.
struct CellBorders {
    top: Option<TableBorderLine>,
    bottom: Option<TableBorderLine>,
    left: Option<TableBorderLine>,
    right: Option<TableBorderLine>,
}

/// Emit all four borders for a cell, extending verticals for clean corners.
fn emit_cell_borders(
    commands: &mut Vec<DrawCommand>,
    b: CellBorders,
    cell_x: Pt, cell_w: Pt, row_y: Pt, row_h: Pt,
) {
    let h_top_half = b.top.map(|b| b.width * 0.5).unwrap_or(Pt::ZERO);
    let h_bot_half = b.bottom.map(|b| b.width * 0.5).unwrap_or(Pt::ZERO);

    if let Some(ref border) = b.top {
        emit_border(commands, border,
            PtOffset::new(cell_x, row_y),
            PtOffset::new(cell_x + cell_w, row_y));
    }
    if let Some(ref border) = b.bottom {
        emit_border(commands, border,
            PtOffset::new(cell_x, row_y + row_h),
            PtOffset::new(cell_x + cell_w, row_y + row_h));
    }

    // Vertical borders extend past horizontal border thickness for clean corners.
    let v_top = row_y - h_top_half;
    let v_bot = row_y + row_h + h_bot_half;
    if let Some(ref border) = b.left {
        emit_border(commands, border,
            PtOffset::new(cell_x, v_top),
            PtOffset::new(cell_x, v_bot));
    }
    if let Some(ref border) = b.right {
        emit_border(commands, border,
            PtOffset::new(cell_x + cell_w, v_top),
            PtOffset::new(cell_x + cell_w, v_bot));
    }
}

/// Emit a single border line (or two lines for §17.4.38 double style).
fn emit_border(commands: &mut Vec<DrawCommand>, b: &TableBorderLine, p1: PtOffset, p2: PtOffset) {
    match b.style {
        TableBorderStyle::Single => {
            commands.push(DrawCommand::Line {
                line: PtLineSegment::new(p1, p2),
                color: b.color,
                width: b.width,
            });
        }
        TableBorderStyle::Double => {
            // §17.4.38: total = w:sz, each sub-line = sz/3, gap = sz/3.
            let sub_w = b.width * (1.0 / 3.0);
            let off = sub_w; // half sub-line + half gap = sz/3
            // Offset perpendicular to the line direction.
            let dx = if p1.x == p2.x { off } else { Pt::ZERO };
            let dy = if p1.y == p2.y { off } else { Pt::ZERO };
            commands.push(DrawCommand::Line {
                line: PtLineSegment::new(
                    PtOffset::new(p1.x - dx, p1.y - dy),
                    PtOffset::new(p2.x - dx, p2.y - dy)),
                color: b.color,
                width: sub_w,
            });
            commands.push(DrawCommand::Line {
                line: PtLineSegment::new(
                    PtOffset::new(p1.x + dx, p1.y + dy),
                    PtOffset::new(p2.x + dx, p2.y + dy)),
                color: b.color,
                width: sub_w,
            });
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
            blocks: vec![CellBlock::Paragraph {
                fragments: vec![text_frag(text, 30.0)],
                style: ParagraphStyle::default(),
            }],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
            cell_borders: None, vertical_merge: None,
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
            height_rule: None,
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
                height_rule: None,
            },
            TableRowInput {
                cells: vec![simple_cell("c"), simple_cell("d")],
                height_rule: None,
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
                    blocks: vec![CellBlock::Paragraph {
                        fragments: vec![text_frag("long ", 60.0), text_frag("text", 60.0)],
                        style: ParagraphStyle::default(),
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None, cell_borders: None, vertical_merge: None,
                },
            ],
            height_rule: None,
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
            height_rule: Some(RowHeightRule::AtLeast(Pt::new(40.0))),
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), None);

        assert_eq!(result.size.height.raw(), 40.0, "min height > content height");
    }

    #[test]
    fn cell_shading_emits_rect() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![CellBlock::Paragraph {
                    fragments: vec![text_frag("x", 10.0)],
                    style: ParagraphStyle::default(),
                }],
                margins: PtEdgeInsets::ZERO,
                grid_span: 1,
                shading: Some(RgbColor { r: 200, g: 200, b: 200 }),
                cell_borders: None, vertical_merge: None,
            }],
            height_rule: None,
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
            height_rule: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(&rows, &col_widths, &body_constraints(), Pt::new(14.0), Some(&super::TableBorderConfig { top: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, style: super::TableBorderStyle::Single }), bottom: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, style: super::TableBorderStyle::Single }), left: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, style: super::TableBorderStyle::Single }), right: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, style: super::TableBorderStyle::Single }), inside_h: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, style: super::TableBorderStyle::Single }), inside_v: Some(super::TableBorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, style: super::TableBorderStyle::Single }) }));

        let line_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Line { .. }))
            .count();
        // §17.4.43: shared edges drawn once after conflict resolution.
        // Top(2) + bottom(2) + left(1) + insideV(1) + right(1) = 7 lines.
        assert_eq!(line_count, 7);
    }

    #[test]
    fn grid_span_widens_cell() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![CellBlock::Paragraph {
                    fragments: vec![text_frag("spanning", 30.0)],
                    style: ParagraphStyle::default(),
                }],
                margins: PtEdgeInsets::ZERO,
                grid_span: 2, // spans both columns
                shading: None, cell_borders: None, vertical_merge: None,
            }],
            height_rule: None,
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
                blocks: vec![CellBlock::Paragraph {
                    fragments: vec![text_frag("text", 30.0)],
                    style: ParagraphStyle::default(),
                }],
                margins: PtEdgeInsets::new(Pt::new(5.0), Pt::new(10.0), Pt::new(5.0), Pt::new(10.0)),
                grid_span: 1,
                shading: None, cell_borders: None, vertical_merge: None,
            }],
            height_rule: None,
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
