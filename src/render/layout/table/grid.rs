use crate::render::dimension::Pt;

use super::{CellLayoutEntry, MeasuredTable, TableCellInput, TableRowInput, VerticalMergeState};

/// A group of rows that must stay together during page splitting.
pub(super) struct RowGroup {
    pub(super) start: usize,
    pub(super) end: usize, // exclusive
    pub(super) height: Pt,
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

/// Build atomic row groups for pagination.
///
/// Groups are formed by: vMerge groups (Restart through consecutive Continue
/// rows) and §17.4.1 cantSplit rows. Each group is an indivisible unit for
/// page-break decisions.
pub(super) fn build_row_groups(rows: &[TableRowInput], measured: &MeasuredTable) -> Vec<RowGroup> {
    let mut groups = Vec::new();
    let mut i = 0;
    while i < rows.len() {
        let start = i;
        // Extend through vMerge Continue rows.
        i += 1;
        while i < rows.len() {
            let has_continue = rows[i]
                .cells
                .iter()
                .any(|c| c.vertical_merge == Some(VerticalMergeState::Continue));
            if has_continue {
                i += 1;
            } else {
                break;
            }
        }
        let height: Pt = measured.rows[start..i]
            .iter()
            .map(|mr| mr.height + mr.border_gap_below)
            .sum();
        groups.push(RowGroup {
            start,
            end: i,
            height,
        });
    }
    groups
}

/// §17.4.85: expand the last row of each vertical merge group so the
/// Restart cell's content fits within the combined spanned row heights.
pub(super) fn expand_rows_for_vmerge(
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

            // Distribute overflow evenly across all rows in the merge group.
            let spanned: Pt = row_heights[row_idx..=last_merged_row].iter().copied().sum();
            if content_h > spanned {
                let overflow = content_h - spanned;
                let num_rows = (last_merged_row - row_idx + 1) as f32;
                let per_row = overflow / num_rows;
                for h in &mut row_heights[row_idx..=last_merged_row] {
                    *h += per_row;
                }
            }
        }
    }
}

/// Find the cell in a row that covers the given grid column index.
pub(super) fn find_cell_at_grid_col(
    row: &TableRowInput,
    target_grid_col: usize,
) -> Option<&TableCellInput> {
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

/// Check if the cell at `grid_col` in `row` is a vMerge Continue cell.
pub(super) fn is_vmerge_continue(row: &TableRowInput, grid_col: usize) -> bool {
    find_cell_at_grid_col(row, grid_col)
        .is_some_and(|c| c.vertical_merge == Some(VerticalMergeState::Continue))
}

/// Return the cell index (not grid column) for the cell covering `grid_col`.
pub(super) fn cell_index_at_grid_col(row: &TableRowInput, target_grid_col: usize) -> Option<usize> {
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
