//! Table command emission — positions cells and emits border commands.

use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

use crate::render::layout::draw_command::DrawCommand;

use super::borders::{border_width, emit_cell_borders, CellBorders};
use super::grid::is_vmerge_continue;
use super::types::{CellVAlign, MeasuredTable, TableBorderLine, TableRowInput, VerticalMergeState};

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

            // §17.4.85: for vMerge=Restart cells, vAlign operates over the
            // whole merged span, not just the starting row. Sum the heights
            // (and intermediate border gaps) of all rows this cell spans.
            let effective_h = if cell_input.vertical_merge == Some(VerticalMergeState::Restart) {
                merged_span_height(measured, rows, row_idx, entry.grid_col)
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
