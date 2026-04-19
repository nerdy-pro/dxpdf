//! Table command emission — positions cells and emits border commands.

use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

use crate::render::layout::draw_command::DrawCommand;

use super::borders::{border_width, emit_cell_borders, CellBorders};
use super::grid::is_vmerge_continue;
use super::types::{
    CellVAlign, MeasuredRow, MeasuredTable, TableBorderLine, TableRowInput, VerticalMergeState,
};

/// Layered command buffers for table rendering: shading, content, borders.
pub(super) struct TableCommandBuffers<'a> {
    pub(super) commands: &'a mut Vec<DrawCommand>,
    pub(super) content_commands: &'a mut Vec<DrawCommand>,
    pub(super) border_commands: &'a mut Vec<DrawCommand>,
}

/// Emit draw commands for a range of measured rows.
///
/// `top_border_override`: if `Some`, the first row in the range gets this border
/// as its top edge. Used for page-split tables where the measured top borders were
/// suppressed (adjacent table collapse) or resolved away (conflict resolution),
/// but the continuation slice still needs a visible top boundary.
pub(super) fn emit_table_rows(
    measured: &MeasuredTable,
    rows: &[TableRowInput],
    row_range: std::ops::Range<usize>,
    cursor_y: &mut Pt,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
) {
    let num_rows = measured.rows.len();
    let range_start = row_range.start;
    for row_idx in row_range {
        let mr = &measured.rows[row_idx];
        let is_first_in_range = row_idx == range_start;
        let has_next_in_slice = row_idx + 1 < num_rows;
        emit_one_row(
            mr,
            &rows[row_idx],
            cursor_y,
            bufs,
            if is_first_in_range {
                top_border_override
            } else {
                None
            },
            has_next_in_slice,
            Some((measured, rows, row_idx)),
        );
    }
}

/// Emit a custom `MeasuredRow` (produced by `split::split_row_at`). Unlike
/// the range-based emit above, this takes a single already-built
/// `MeasuredRow` and the matching `TableRowInput`. `vmerge_ctx` is always
/// `None` here — split rows can't contain vMerge (the group is flagged
/// not splittable if any cell is merged).
pub(super) fn emit_split_row(
    mr: &MeasuredRow,
    row: &TableRowInput,
    cursor_y: &mut Pt,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
    has_next_in_slice: bool,
) {
    emit_one_row(
        mr,
        row,
        cursor_y,
        bufs,
        top_border_override,
        has_next_in_slice,
        None,
    );
}

fn emit_one_row(
    mr: &MeasuredRow,
    row: &TableRowInput,
    cursor_y: &mut Pt,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
    has_next_in_slice: bool,
    // For vMerge=Restart cells, resolve the full merged span height.
    // `None` disables the lookup (used for split rows and for standalone
    // emission paths that don't carry a merge context).
    vmerge_ctx: Option<(&MeasuredTable, &[TableRowInput], usize)>,
) {
    let row_height = mr.height;
    for (cell_ci, (entry, cell_input)) in mr.entries.iter().zip(row.cells.iter()).enumerate() {
        if let Some(color) = cell_input.shading {
            bufs.commands.push(DrawCommand::Rect {
                rect: PtRect::from_xywh(entry.cell_x, *cursor_y, entry.cell_w, row_height),
                color,
            });
        }

        // §17.4.38: restore top border when this row is the first in its
        // slice and the resolved top was removed (conflict resolution or
        // adjacent-table collapse).
        let cell_top = if mr.borders[cell_ci].top.is_none() {
            top_border_override
        } else {
            mr.borders[cell_ci].top
        };
        let b_left = mr.borders[cell_ci].left;
        let b_right = mr.borders[cell_ci].right;
        let b_bottom = mr.borders[cell_ci].bottom;

        let dx = (border_width(b_left) - cell_input.margins.left).max(Pt::ZERO);
        let dy_border = (border_width(cell_top) - cell_input.margins.top).max(Pt::ZERO);

        // §17.4.85: for vMerge=Restart cells, vAlign operates over the
        // whole merged span, not just the starting row.
        let effective_h = if cell_input.vertical_merge == Some(VerticalMergeState::Restart) {
            vmerge_ctx
                .map(|(m, rs, row_idx)| merged_span_height(m, rs, row_idx, entry.grid_col))
                .unwrap_or(row_height)
        } else {
            row_height
        };

        let content_h = entry.layout.content_height + cell_input.margins.vertical();
        let dy_valign = match cell_input.vertical_align {
            CellVAlign::Bottom => (effective_h - content_h - dy_border).max(Pt::ZERO),
            CellVAlign::Center => ((effective_h - content_h - dy_border) * 0.5).max(Pt::ZERO),
            CellVAlign::Top => Pt::ZERO,
        };

        for cmd in &entry.layout.commands {
            let mut cmd = cmd.clone();
            cmd.shift(entry.cell_x + dx, *cursor_y + dy_border + dy_valign);
            bufs.content_commands.push(cmd);
        }

        let bottom_border_gap = if has_next_in_slice {
            border_width(b_bottom)
        } else {
            Pt::ZERO
        };
        emit_cell_borders(
            bufs.border_commands,
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

/// Total vertical space owned by a vMerge=Restart cell at `grid_col`.
/// Includes the restart row's height and every `Continue` row below it,
/// plus the `border_gap_below` of intermediate rows (the cell's own top/
/// bottom borders between merged rows were suppressed in measurement, so
/// the gap is driven by sibling columns only).
fn merged_span_height(
    measured: &MeasuredTable,
    rows: &[TableRowInput],
    start_row: usize,
    grid_col: usize,
) -> Pt {
    let mut total = measured.rows[start_row].height;
    let mut row = start_row + 1;
    while row < rows.len() && is_vmerge_continue(&rows[row], grid_col) {
        total += measured.rows[row - 1].border_gap_below;
        total += measured.rows[row].height;
        row += 1;
    }
    total
}
