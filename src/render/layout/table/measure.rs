//! Table measurement phase — cell layout and border resolution.

use crate::render::dimension::Pt;

use crate::render::layout::cell::{layout_cell, CellLayout};

use super::borders::{
    border_width, resolve_border_conflict, resolve_cell_effective_borders, CellBorders,
};
use super::grid::{cell_index_at_grid_col, expand_rows_for_vmerge, is_vmerge_continue};
use super::types::{
    CellLayoutEntry, MeasuredRow, MeasuredTable, RowHeightRule, TableBorderConfig, TableRowInput,
    VerticalMergeState,
};

/// Measure all table rows: resolve borders, lay out cell content, compute heights.
/// This is the shared measurement phase used by both `layout_table` (monolithic)
/// and `layout_table_paginated` (page-splitting).
///
/// §17.4.38: `suppress_first_row_top` — when `true`, the top border of the first
/// row is suppressed. Used for adjacent table border collapse: consecutive tables
/// with the same style are treated as a single merged table, so the second table's
/// top border would duplicate the first table's bottom border.
pub(super) fn measure_table_rows(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: crate::render::layout::paragraph::MeasureTextFn<'_>,
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
