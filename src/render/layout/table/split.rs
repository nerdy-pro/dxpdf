//! Row splitting for page pagination.
//!
//! §17.4.1: a row without `cantSplit` may have its content broken across
//! page boundaries. This module derives safe cut points from a row's
//! already-laid-out cell commands and partitions those commands into a
//! first slice (stays on the current page) and a second slice (flows to
//! the next page). Both halves preserve the cell's top/bottom margins so
//! the text doesn't collide with the cut-edge border.

use crate::render::dimension::Pt;
use crate::render::layout::draw_command::DrawCommand;

use super::borders::CellBorders;
use super::types::{CellLayoutEntry, MeasuredRow, TableRowInput};

/// Options for finding a row cut.
pub(super) struct RowCutInput<'a> {
    pub(super) mr: &'a MeasuredRow,
    pub(super) row: &'a TableRowInput,
    /// Space available for the row on the current page, measured from the
    /// row's top edge.
    pub(super) available: Pt,
}

/// Per-cell split decision: where to partition the commands and how far
/// to shift the tail commands so the continuation half has the cell's
/// natural top margin.
struct CellCut {
    /// Partition threshold: commands with primary-Y strictly less than
    /// this value stay on the first half.
    content_cut_y: Pt,
    /// Amount by which second-half commands are shifted up so that the
    /// first surviving line lands at `margin_top + ascent` in the
    /// continuation cell.
    shift: Pt,
}

impl CellCut {
    /// "Don't split" sentinel — all commands stay on the first half.
    fn keep_all() -> Self {
        Self {
            content_cut_y: Pt::new(f32::INFINITY),
            shift: Pt::ZERO,
        }
    }
}

/// Aggregate split decision for a whole row.
pub(super) struct SplitCut {
    /// Row-level first-half height (max across cells of per-cell first
    /// half heights). Each cell's visible box in the first half extends
    /// from the row top to this Y.
    first_half_height: Pt,
    /// One per cell, in the same order as `row.cells`.
    cells: Vec<CellCut>,
}

/// Pick a row cut that fits in `available`, honoring each cell's top and
/// bottom margins so the cut-edge border gets the natural padding Word
/// preserves on split cells.
///
/// Returns `None` when no cell has at least one line that fits within
/// `available - margins.vertical()`.
pub(super) fn find_row_cut(input: &RowCutInput<'_>) -> Option<SplitCut> {
    let mut cells: Vec<CellCut> = Vec::with_capacity(input.row.cells.len());
    let mut first_half_height = Pt::ZERO;
    let mut any_fits = false;

    for (entry, cell) in input.mr.entries.iter().zip(&input.row.cells) {
        match cut_for_cell(
            entry,
            cell.margins.top,
            cell.margins.bottom,
            input.available,
        ) {
            Some((cut, half_h)) => {
                any_fits = true;
                if half_h > first_half_height {
                    first_half_height = half_h;
                }
                cells.push(cut);
            }
            None => cells.push(CellCut::keep_all()),
        }
    }

    if !any_fits {
        return None;
    }
    Some(SplitCut {
        first_half_height,
        cells,
    })
}

/// For a single cell, choose the largest prefix of lines that fits in
/// `available`. Returns the `CellCut` and the first-half height this cell
/// needs (midpoint between the last-fit baseline and the next, plus the
/// cell's bottom margin — the visual equivalent of half a line gap of
/// breathing room before the cut-edge border).
fn cut_for_cell(
    entry: &CellLayoutEntry,
    _margin_top: Pt,
    margin_bottom: Pt,
    available: Pt,
) -> Option<(CellCut, Pt)> {
    let mut baselines: Vec<Pt> = entry
        .layout
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { position, .. } => Some(position.y),
            _ => None,
        })
        .collect();
    baselines.sort_by(|a, b| {
        a.raw()
            .partial_cmp(&b.raw())
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    baselines.dedup_by(|a, b| (a.raw() - b.raw()).abs() < 0.01);

    // Need at least two baselines to have somewhere to cut between.
    if baselines.len() < 2 {
        return None;
    }

    // Bottom of the first-half box must land below the last-fit line
    // with room for the cell's bottom margin.
    let budget = available - margin_bottom;
    if budget <= Pt::ZERO {
        return None;
    }

    let first = baselines[0];
    // Find the largest k such that the midpoint between baselines[k] and
    // baselines[k+1] fits in `budget`. The midpoint is where the border
    // will sit (giving half a line-gap of padding above it).
    let mut best_k: Option<usize> = None;
    for k in 0..baselines.len() - 1 {
        let midpoint = Pt::new((baselines[k].raw() + baselines[k + 1].raw()) * 0.5);
        if midpoint <= budget {
            best_k = Some(k);
        } else {
            break;
        }
    }
    let k = best_k?;

    let midpoint = Pt::new((baselines[k].raw() + baselines[k + 1].raw()) * 0.5);
    // Shift the second half so line k+1 lands at the cell's natural first-
    // line baseline (`margin_top + ascent` = `baselines[0]`).
    let shift = baselines[k + 1] - first;
    let half_h = midpoint + margin_bottom;
    Some((
        CellCut {
            content_cut_y: baselines[k + 1],
            shift,
        },
        half_h,
    ))
}

/// A row cut into two halves. Each half is a full `MeasuredRow` ready to
/// pass to `emit_one_row`.
pub(super) struct SplitRow {
    pub(super) first: MeasuredRow,
    pub(super) second: MeasuredRow,
}

/// Split a row using `cut`. Each cell's commands are partitioned at the
/// cell's own `content_cut_y`; the tail is shifted so its first line
/// lands at the cell's natural `margin_top + ascent` in the
/// continuation.
///
/// Borders: Word preserves the full border box on each half. First half
/// keeps all original borders. Second half inherits a top border from
/// the original top if set, otherwise falls back to the original bottom
/// (same inside-horizontal style used by conflict resolution on
/// mid-table rows).
pub(super) fn split_row_at(mr: &MeasuredRow, cut: &SplitCut) -> SplitRow {
    let first_h = cut.first_half_height;
    // The largest shift across cells is the amount of content consumed by
    // the first half. Whatever's left in the original row height is what
    // the continuation needs to render (top/bottom margins are already
    // included in `mr.height`, so no additional padding is needed).
    let max_shift = cut.cells.iter().map(|c| c.shift).fold(Pt::ZERO, Pt::max);
    let second_h = (mr.height - max_shift).max(Pt::ZERO);

    let mut first_entries: Vec<CellLayoutEntry> = Vec::with_capacity(mr.entries.len());
    let mut second_entries: Vec<CellLayoutEntry> = Vec::with_capacity(mr.entries.len());

    for (entry, cc) in mr.entries.iter().zip(cut.cells.iter()) {
        let (first_cmds, second_cmds) =
            partition_commands(&entry.layout.commands, cc.content_cut_y, cc.shift);
        first_entries.push(CellLayoutEntry {
            layout: crate::render::layout::cell::CellLayout {
                commands: first_cmds,
                content_height: entry.layout.content_height.min(first_h),
            },
            cell_x: entry.cell_x,
            cell_w: entry.cell_w,
            grid_col: entry.grid_col,
        });
        second_entries.push(CellLayoutEntry {
            layout: crate::render::layout::cell::CellLayout {
                commands: second_cmds,
                content_height: (entry.layout.content_height - cc.shift).max(Pt::ZERO),
            },
            cell_x: entry.cell_x,
            cell_w: entry.cell_w,
            grid_col: entry.grid_col,
        });
    }

    let first_borders: Vec<CellBorders> = mr.borders.to_vec();
    let second_borders: Vec<CellBorders> = mr
        .borders
        .iter()
        .map(|b| CellBorders {
            top: b.top.or(b.bottom),
            bottom: b.bottom,
            left: b.left,
            right: b.right,
        })
        .collect();

    SplitRow {
        first: MeasuredRow {
            entries: first_entries,
            borders: first_borders,
            height: first_h,
            border_gap_below: Pt::ZERO,
        },
        second: MeasuredRow {
            entries: second_entries,
            borders: second_borders,
            height: second_h,
            border_gap_below: mr.border_gap_below,
        },
    }
}

/// Split a command list at `cut_y`. Commands whose primary Y < `cut_y`
/// go to the first half; the rest land in the second half shifted up by
/// `shift` so they start at the continuation cell's natural top margin.
fn partition_commands(
    commands: &[DrawCommand],
    cut_y: Pt,
    shift: Pt,
) -> (Vec<DrawCommand>, Vec<DrawCommand>) {
    let mut first = Vec::new();
    let mut second = Vec::new();
    for cmd in commands {
        if command_primary_y(cmd) < cut_y {
            first.push(cmd.clone());
        } else {
            let mut c = cmd.clone();
            c.shift_y(-shift);
            second.push(c);
        }
    }
    (first, second)
}

/// The Y used to decide which side of the cut a command belongs to. For
/// `Text` we use the baseline; for rect/line/image we use the top edge.
fn command_primary_y(cmd: &DrawCommand) -> Pt {
    match cmd {
        DrawCommand::Text { position, .. } | DrawCommand::NamedDestination { position, .. } => {
            position.y
        }
        DrawCommand::Underline { line, .. } | DrawCommand::Line { line, .. } => line.start.y,
        DrawCommand::Image { rect, .. }
        | DrawCommand::Rect { rect, .. }
        | DrawCommand::LinkAnnotation { rect, .. }
        | DrawCommand::InternalLink { rect, .. } => rect.origin.y,
        DrawCommand::Path { origin, .. } => origin.y,
    }
}
