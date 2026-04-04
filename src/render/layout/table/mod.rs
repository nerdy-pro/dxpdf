//! Table layout — 3-pass column sizing, cell layout, border rendering.
//!
//! Pass 1: Compute column widths from grid definitions or equal distribution.
//! Pass 2: Lay out each cell with tight width constraints, determine row heights.
//! Pass 3: Position cells and emit border commands.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtEdgeInsets, PtRect, PtSize};
use crate::render::resolve::color::RgbColor;

use super::cell::{layout_cell, CellLayout};
use super::draw_command::DrawCommand;
use super::BoxConstraints;

mod borders;
mod grid;

pub use grid::compute_column_widths;

use borders::{border_width, emit_cell_borders, resolve_cell_effective_borders, resolve_border_conflict, CellBorders};
use grid::{build_row_groups, cell_index_at_grid_col, expand_rows_for_vmerge, is_vmerge_continue};

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
    /// §17.4.49: row repeats as header on each continuation page.
    pub is_header: Option<bool>,
    /// §17.4.1: if true, row cannot be split across pages.
    pub cant_split: Option<bool>,
}

/// §17.4.84: cell vertical alignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellVAlign {
    Top,
    Center,
    Bottom,
}

/// A single cell for layout.
pub struct TableCellInput {
    pub blocks: Vec<super::section::LayoutBlock>,
    pub margins: PtEdgeInsets,
    /// Number of grid columns this cell spans (gridSpan, default 1).
    pub grid_span: u32,
    /// Background color for cell shading.
    pub shading: Option<RgbColor>,
    /// §17.7.6: per-cell resolved borders from conditional formatting.
    pub cell_borders: Option<CellBorderConfig>,
    /// §17.4.85: vertical merge state.
    pub vertical_merge: Option<VerticalMergeState>,
    /// §17.4.84: vertical alignment of content within the cell.
    pub vertical_align: CellVAlign,
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

/// Per-row measurement data from the table measurement phase.
/// Contains everything needed to emit draw commands for this row.
struct MeasuredRow {
    entries: Vec<CellLayoutEntry>,
    borders: Vec<CellBorders>,
    height: Pt,
    /// §17.4.38: maximum bottom border width for gap between this row and the next.
    border_gap_below: Pt,
}

/// Result of the table measurement phase.
struct MeasuredTable {
    rows: Vec<MeasuredRow>,
    table_width: Pt,
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

/// Measure all table rows: resolve borders, lay out cell content, compute heights.
/// This is the shared measurement phase used by both `layout_table` (monolithic)
/// and `layout_table_paginated` (page-splitting).
///
/// §17.4.38: `suppress_first_row_top` — when `true`, the top border of the first
/// row is suppressed. Used for adjacent table border collapse: consecutive tables
/// with the same style are treated as a single merged table, so the second table's
/// top border would duplicate the first table's bottom border.
fn measure_table_rows(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    suppress_first_row_top: bool,
) -> MeasuredTable {
    let table_width: Pt = col_widths.iter().copied().sum();
    let num_rows = rows.len();
    let mut row_heights = Vec::with_capacity(num_rows);

    // Pass 2a: resolve borders for every cell.
    let mut resolved_borders: Vec<Vec<CellBorders>> = Vec::new();
    {
        let mut grid_indices: Vec<Vec<usize>> = Vec::new();
        for (row_idx, row) in rows.iter().enumerate() {
            let mut row_borders = Vec::new();
            let mut row_grid = Vec::new();
            let mut grid_idx = 0;
            for (cell_ci, cell_input) in row.cells.iter().enumerate() {
                let (mut b_top, mut b_bottom, b_left, b_right) = resolve_cell_effective_borders(
                    cell_input,
                    borders,
                    row_idx,
                    cell_ci,
                    num_rows,
                    row.cells.len(),
                );
                if cell_input.vertical_merge == Some(VerticalMergeState::Continue) {
                    b_top = None;
                }
                if row_idx + 1 < num_rows && is_vmerge_continue(&rows[row_idx + 1], grid_idx) {
                    b_bottom = None;
                }
                row_borders.push(CellBorders {
                    top: b_top,
                    bottom: b_bottom,
                    left: b_left,
                    right: b_right,
                });
                row_grid.push(grid_idx);
                grid_idx += cell_input.grid_span.max(1) as usize;
            }
            resolved_borders.push(row_borders);
            grid_indices.push(row_grid);
        }

        // §17.4.43: conflict resolution at shared edges.
        for row_idx in 0..num_rows {
            let num_cells = rows[row_idx].cells.len();
            for cell_ci in 0..num_cells {
                if cell_ci + 1 < num_cells {
                    let right = resolved_borders[row_idx][cell_ci].right;
                    let left = resolved_borders[row_idx][cell_ci + 1].left;
                    let winner = resolve_border_conflict(right, left);
                    resolved_borders[row_idx][cell_ci].right = winner;
                    resolved_borders[row_idx][cell_ci + 1].left = None;
                }
                if row_idx + 1 < num_rows {
                    let start_grid = grid_indices[row_idx][cell_ci];
                    let span = rows[row_idx].cells[cell_ci].grid_span.max(1) as usize;
                    let mut resolved_once = false;
                    for gc in start_grid..start_grid + span {
                        if let Some(below_ci) = cell_index_at_grid_col(&rows[row_idx + 1], gc) {
                            if !resolved_once {
                                let bottom = resolved_borders[row_idx][cell_ci].bottom;
                                let top = resolved_borders[row_idx + 1][below_ci].top;
                                let winner = resolve_border_conflict(bottom, top);
                                resolved_borders[row_idx][cell_ci].bottom = winner;
                                resolved_once = true;
                            }
                            resolved_borders[row_idx + 1][below_ci].top = None;
                        }
                    }
                }
            }
        }

        // §17.4.38: suppress first-row top borders for adjacent table collapse.
        if suppress_first_row_top && !resolved_borders.is_empty() {
            for b in &mut resolved_borders[0] {
                b.top = None;
            }
        }
    }

    // Pass 2b: lay out each cell.
    let mut row_cell_layouts: Vec<Vec<CellLayoutEntry>> = Vec::new();

    for (row_idx, row) in rows.iter().enumerate() {
        let mut entries = Vec::new();
        let mut max_height = Pt::ZERO;
        let mut grid_idx = 0;

        for (cell_ci, cell) in row.cells.iter().enumerate() {
            let span = cell.grid_span.max(1) as usize;
            let cell_w: Pt = col_widths[grid_idx..grid_idx + span].iter().copied().sum();
            let cell_x: Pt = col_widths[..grid_idx].iter().copied().sum();

            let b = &resolved_borders[row_idx][cell_ci];
            let extra_left = (border_width(b.left) - cell.margins.left).max(Pt::ZERO);
            let extra_right = (border_width(b.right) - cell.margins.right).max(Pt::ZERO);
            let layout_w = (cell_w - extra_left - extra_right).max(Pt::ZERO);

            let is_continue = cell.vertical_merge == Some(VerticalMergeState::Continue);
            let layout = if is_continue {
                CellLayout {
                    commands: Vec::new(),
                    content_height: Pt::ZERO,
                }
            } else {
                layout_cell(
                    &cell.blocks,
                    layout_w,
                    &cell.margins,
                    default_line_height,
                    measure_text,
                )
            };

            if cell.vertical_merge.is_none() {
                max_height = max_height.max(layout.content_height + cell.margins.vertical());
            }

            entries.push(CellLayoutEntry {
                layout,
                cell_x,
                cell_w,
                grid_col: grid_idx,
            });
            grid_idx += span;
        }

        match row.height_rule {
            Some(RowHeightRule::AtLeast(min_h)) => max_height = max_height.max(min_h),
            Some(RowHeightRule::Exact(h)) => max_height = h,
            None => {}
        }

        row_heights.push(max_height);
        row_cell_layouts.push(entries);
    }

    // §17.4.85: distribute vMerge overflow.
    expand_rows_for_vmerge(rows, &row_cell_layouts, &mut row_heights);

    // Compute border gaps and assemble measured rows.
    let measured_rows: Vec<MeasuredRow> = row_cell_layouts
        .into_iter()
        .zip(resolved_borders)
        .zip(row_heights.iter())
        .enumerate()
        .map(|(row_idx, ((entries, borders), &height))| {
            let border_gap_below = if row_idx + 1 < num_rows {
                borders
                    .iter()
                    .map(|b| border_width(b.bottom))
                    .fold(Pt::ZERO, Pt::max)
            } else {
                Pt::ZERO
            };
            MeasuredRow {
                entries,
                borders,
                height,
                border_gap_below,
            }
        })
        .collect();

    MeasuredTable {
        rows: measured_rows,
        table_width,
    }
}

/// Layered command buffers for table rendering: shading, content, borders.
struct TableCommandBuffers<'a> {
    commands: &'a mut Vec<DrawCommand>,
    content_commands: &'a mut Vec<DrawCommand>,
    border_commands: &'a mut Vec<DrawCommand>,
}

/// Emit draw commands for a range of measured rows.
///
/// `top_border_override`: if `Some`, the first row in the range gets this border
/// as its top edge. Used for page-split tables where the measured top borders were
/// suppressed (adjacent table collapse) or resolved away (conflict resolution),
/// but the continuation slice still needs a visible top boundary.
fn emit_table_rows(
    measured: &MeasuredTable,
    rows: &[TableRowInput],
    row_range: std::ops::Range<usize>,
    cursor_y: &mut Pt,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
) {
    let commands = &mut *bufs.commands;
    let content_commands = &mut *bufs.content_commands;
    let border_commands = &mut *bufs.border_commands;
    let num_rows = measured.rows.len();
    let range_start = row_range.start;
    for row_idx in row_range {
        let mr = &measured.rows[row_idx];
        let row_height = mr.height;
        let is_first_in_range = row_idx == range_start;

        for (cell_ci, (entry, cell_input)) in mr
            .entries
            .iter()
            .zip(rows[row_idx].cells.iter())
            .enumerate()
        {
            if let Some(color) = cell_input.shading {
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(entry.cell_x, *cursor_y, entry.cell_w, row_height),
                    color,
                });
            }

            // §17.4.38: restore top border for the first row on a page slice.
            let cell_top = if is_first_in_range && mr.borders[cell_ci].top.is_none() {
                top_border_override
            } else {
                mr.borders[cell_ci].top
            };
            let b_left = mr.borders[cell_ci].left;
            let b_right = mr.borders[cell_ci].right;
            let b_bottom = mr.borders[cell_ci].bottom;

            let dx = (border_width(b_left) - cell_input.margins.left).max(Pt::ZERO);
            let dy_border = (border_width(cell_top) - cell_input.margins.top).max(Pt::ZERO);

            let content_h = entry.layout.content_height + cell_input.margins.vertical();
            let dy_valign = match cell_input.vertical_align {
                CellVAlign::Bottom => (row_height - content_h - dy_border).max(Pt::ZERO),
                CellVAlign::Center => ((row_height - content_h - dy_border) * 0.5).max(Pt::ZERO),
                CellVAlign::Top => Pt::ZERO,
            };

            for cmd in &entry.layout.commands {
                let mut cmd = cmd.clone();
                cmd.shift(entry.cell_x + dx, *cursor_y + dy_border + dy_valign);
                content_commands.push(cmd);
            }

            let bottom_border_gap = if row_idx + 1 < num_rows {
                border_width(b_bottom)
            } else {
                Pt::ZERO
            };
            emit_cell_borders(
                border_commands,
                CellBorders {
                    top: cell_top,
                    bottom: b_bottom,
                    left: b_left,
                    right: b_right,
                },
                entry.cell_x,
                entry.cell_w,
                *cursor_y,
                row_height + bottom_border_gap,
            );
        }

        *cursor_y += row_height + mr.border_gap_below;
    }
}

/// Lay out a table: compute column widths, lay out cells, emit borders.
///
/// §17.4.38: `suppress_first_row_top` suppresses the top border of the first row
/// for adjacent table border collapse.
pub fn layout_table(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    _constraints: &BoxConstraints,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    suppress_first_row_top: bool,
) -> TableLayout {
    if rows.is_empty() || col_widths.is_empty() {
        return TableLayout {
            commands: Vec::new(),
            size: PtSize::ZERO,
        };
    }

    let measured = measure_table_rows(
        rows,
        col_widths,
        default_line_height,
        borders,
        measure_text,
        suppress_first_row_top,
    );

    let mut commands = Vec::new();
    let mut content_commands = Vec::new();
    let mut border_commands = Vec::new();
    let mut cursor_y = Pt::ZERO;

    // Monolithic table: no top border override needed — borders are resolved correctly.
    emit_table_rows(
        &measured,
        rows,
        0..measured.rows.len(),
        &mut cursor_y,
        &mut TableCommandBuffers {
            commands: &mut commands,
            content_commands: &mut content_commands,
            border_commands: &mut border_commands,
        },
        None,
    );

    commands.append(&mut content_commands);
    commands.append(&mut border_commands);

    TableLayout {
        commands,
        size: PtSize::new(measured.table_width, cursor_y),
    }
}

/// One page-slice of a table, produced by `layout_table_paginated`.
#[derive(Debug)]
pub struct TableSlice {
    /// Draw commands positioned relative to this slice's top-left origin (0,0).
    pub commands: Vec<DrawCommand>,
    /// Size of this slice.
    pub size: PtSize,
}

/// Pagination parameters for `layout_table_paginated`.
pub struct TablePaginationConfig {
    /// Available height on the first page.
    pub available_height: Pt,
    /// Full page height for continuation pages.
    pub page_height: Pt,
    /// Whether to suppress the first row's top border (adjacent table collapse).
    pub suppress_first_row_top: bool,
}

/// Lay out a table with page splitting at row boundaries.
///
/// §17.4.49: header rows repeat on each continuation page.
/// §17.4.1: `cantSplit` rows are kept together (moved to next page if needed).
///
/// Returns one `TableSlice` per page.
pub fn layout_table_paginated(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    _constraints: &BoxConstraints,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    pagination: &TablePaginationConfig,
) -> Vec<TableSlice> {
    let available_height = pagination.available_height;
    let page_height = pagination.page_height;
    let suppress_first_row_top = pagination.suppress_first_row_top;
    if rows.is_empty() || col_widths.is_empty() {
        return vec![TableSlice {
            commands: Vec::new(),
            size: PtSize::ZERO,
        }];
    }

    let measured = measure_table_rows(
        rows,
        col_widths,
        default_line_height,
        borders,
        measure_text,
        suppress_first_row_top,
    );

    // §17.4.49: contiguous header rows from index 0.
    let header_count = rows
        .iter()
        .take_while(|r| r.is_header == Some(true))
        .count();
    let header_height: Pt = measured.rows[..header_count]
        .iter()
        .map(|mr| mr.height + mr.border_gap_below)
        .sum();

    let groups = build_row_groups(rows, &measured);

    // Pack groups into page slices.
    // Each slice is a Vec<Range<usize>> of row ranges to emit.
    let mut slices: Vec<Vec<std::ops::Range<usize>>> = Vec::new();
    let mut current_slice: Vec<std::ops::Range<usize>> = Vec::new();
    let mut remaining = available_height;
    let is_first_slice = |slices: &Vec<Vec<_>>| slices.is_empty();

    for group in &groups {
        // Header rows on the first slice are emitted as normal rows.
        if is_first_slice(&slices) && group.start < header_count {
            current_slice.push(group.start..group.end);
            remaining -= group.height;
            continue;
        }

        if group.height <= remaining {
            current_slice.push(group.start..group.end);
            remaining -= group.height;
        } else {
            // Close current slice.
            slices.push(std::mem::take(&mut current_slice));
            // New page: start with header rows (if any).
            remaining = page_height;
            if header_count > 0 {
                current_slice.push(0..header_count);
                remaining -= header_height;
            }
            if group.height > remaining {
                log::warn!(
                    "[table] row group {}-{} ({:.1}pt) exceeds page height ({:.1}pt available)",
                    group.start,
                    group.end,
                    group.height.raw(),
                    remaining.raw(),
                );
            }
            current_slice.push(group.start..group.end);
            remaining -= group.height;
        }
    }
    slices.push(current_slice);

    // §17.4.38: the table's outer top border, used to restore the top edge
    // on continuation page slices where border conflict resolution or adjacent
    // table collapse removed it.
    let outer_top_border = borders.and_then(|b| b.top);

    // Emit draw commands for each slice.
    slices
        .iter()
        .enumerate()
        .map(|(slice_idx, row_ranges)| {
            let mut commands = Vec::new();
            let mut content_commands = Vec::new();
            let mut border_commands = Vec::new();
            let mut cursor_y = Pt::ZERO;
            for (range_idx, range) in row_ranges.iter().enumerate() {
                // The first range on each continuation slice (slice_idx > 0) needs
                // a top border override, since internal conflict resolution removed
                // it. On the first slice, only apply if suppress_first_row_top was
                // used (adjacent table collapse) — in that case the first range
                // should NOT get the override (the whole point is to suppress it).
                let top_override = if slice_idx > 0 && range_idx == 0 {
                    outer_top_border
                } else {
                    None
                };
                emit_table_rows(
                    &measured,
                    rows,
                    range.clone(),
                    &mut cursor_y,
                    &mut TableCommandBuffers {
                        commands: &mut commands,
                        content_commands: &mut content_commands,
                        border_commands: &mut border_commands,
                    },
                    top_override,
                );
            }
            commands.append(&mut content_commands);
            commands.append(&mut border_commands);
            TableSlice {
                commands,
                size: PtSize::new(measured.table_width, cursor_y),
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
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
                char_spacing: Pt::ZERO,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
        }
    }

    fn simple_cell(text: &str) -> TableCellInput {
        TableCellInput {
            blocks: vec![LayoutBlock::Paragraph {
                fragments: vec![text_frag(text, 30.0)],
                style: ParagraphStyle::default(),
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
            }],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
            cell_borders: None,
            vertical_merge: None,
            vertical_align: CellVAlign::Top,
        }
    }

    fn body_constraints() -> BoxConstraints {
        BoxConstraints::loose(PtSize::new(Pt::new(400.0), Pt::new(1000.0)))
    }

    // ── layout_table ─────────────────────────────────────────────────────

    #[test]
    fn empty_table() {
        let result = layout_table(
            &[],
            &[],
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );
        assert!(result.commands.is_empty());
        assert_eq!(result.size, PtSize::ZERO);
    }

    #[test]
    fn single_cell_table() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("hello")],
            height_rule: None,
            is_header: None,
            cant_split: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

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
                is_header: None,
                cant_split: None,
            },
            TableRowInput {
                cells: vec![simple_cell("c"), simple_cell("d")],
                height_rule: None,
                is_header: None,
                cant_split: None,
            },
        ];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

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
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: vec![text_frag("long ", 60.0), text_frag("text", 60.0)],
                        style: ParagraphStyle::default(),
                        page_break_before: false,
                        footnotes: vec![],
                        floating_images: vec![],
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: None,
                    vertical_align: CellVAlign::Top,
                },
            ],
            height_rule: None,
            is_header: None,
            cant_split: None,
        }];
        // Column B is only 80 wide, so "long " + "text" (120) wraps
        let col_widths = vec![Pt::new(200.0), Pt::new(80.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

        assert_eq!(result.size.height.raw(), 28.0, "row height = tallest cell");
    }

    #[test]
    fn min_row_height_respected() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("x")],
            height_rule: Some(RowHeightRule::AtLeast(Pt::new(40.0))),
            is_header: None,
            cant_split: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

        assert_eq!(
            result.size.height.raw(),
            40.0,
            "min height > content height"
        );
    }

    #[test]
    fn cell_shading_emits_rect() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![LayoutBlock::Paragraph {
                    fragments: vec![text_frag("x", 10.0)],
                    style: ParagraphStyle::default(),
                    page_break_before: false,
                    footnotes: vec![],
                    floating_images: vec![],
                }],
                margins: PtEdgeInsets::ZERO,
                grid_span: 1,
                shading: Some(RgbColor {
                    r: 200,
                    g: 200,
                    b: 200,
                }),
                cell_borders: None,
                vertical_merge: None,
                vertical_align: CellVAlign::Top,
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
        }];
        let col_widths = vec![Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let rect_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Rect { .. }))
            .count();
        assert_eq!(rect_count, 1, "shading produces a Rect command");
    }

    #[test]
    fn grid_span_widens_cell() {
        let rows = vec![TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![LayoutBlock::Paragraph {
                    fragments: vec![text_frag("spanning", 30.0)],
                    style: ParagraphStyle::default(),
                    page_break_before: false,
                    footnotes: vec![],
                    floating_images: vec![],
                }],
                margins: PtEdgeInsets::ZERO,
                grid_span: 2, // spans both columns
                shading: None,
                cell_borders: None,
                vertical_merge: None,
                vertical_align: CellVAlign::Top,
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

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
                blocks: vec![LayoutBlock::Paragraph {
                    fragments: vec![text_frag("text", 30.0)],
                    style: ParagraphStyle::default(),
                    page_break_before: false,
                    footnotes: vec![],
                    floating_images: vec![],
                }],
                margins: PtEdgeInsets::new(
                    Pt::new(5.0),
                    Pt::new(10.0),
                    Pt::new(5.0),
                    Pt::new(10.0),
                ),
                grid_span: 1,
                shading: None,
                cell_borders: None,
                vertical_merge: None,
                vertical_align: CellVAlign::Top,
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

        // Row height = content(14) + top(5) + bottom(5) = 24
        assert_eq!(result.size.height.raw(), 24.0);

        // Text should be offset by left margin
        if let Some(DrawCommand::Text { position, .. }) = result.commands.first() {
            assert_eq!(position.x.raw(), 10.0, "left margin");
        }
    }

    #[test]
    fn suppress_first_row_top_removes_top_borders() {
        let border_line = super::TableBorderLine {
            width: Pt::new(0.5),
            color: RgbColor::BLACK,
            style: super::TableBorderStyle::Single,
        };
        let borders = super::TableBorderConfig {
            top: Some(border_line),
            bottom: Some(border_line),
            left: Some(border_line),
            right: Some(border_line),
            inside_h: None,
            inside_v: None,
        };
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("a")],
            height_rule: None,
            is_header: None,
            cant_split: None,
        }];
        let col_widths = vec![Pt::new(100.0)];

        // Without suppression: top border present.
        let normal = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            Some(&borders),
            None,
            false,
        );
        let normal_borders: Vec<_> = normal
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Rect { color, .. } if *color == RgbColor::BLACK))
            .collect();

        // With suppression: top border removed.
        let suppressed = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            Some(&borders),
            None,
            true,
        );
        let suppressed_borders: Vec<_> = suppressed
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Rect { color, .. } if *color == RgbColor::BLACK))
            .collect();

        // Normal has 4 borders (top, bottom, left, right).
        assert_eq!(normal_borders.len(), 4, "all 4 borders present");
        // Suppressed has 3 borders (bottom, left, right — no top).
        assert_eq!(suppressed_borders.len(), 3, "top border suppressed");
    }
}
