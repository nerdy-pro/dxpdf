//! Table borders — §17.4.66 resolution, and the drawing of what it resolved.
//!
//! Everything starts from [`declare_cell_borders`], which answers what each cell
//! *declares* on its four edges and looks across none of them:
//! [`resolve_cell_effective_borders`] maps the table's borders onto a cell's
//! position in its row and lays the cell's own `w:tcBorders` on top. [`CellEdge`]'s
//! three states are what makes that expressible — an omitted edge and a
//! `val="nil"` one paint the same nothing but inherit differently.
//!
//! Two readers then ask two different questions of those declarations, and the
//! split is the point of this module:
//!
//! * [`resolve_table_cell_borders`] — **how much of each edge is charged to each
//!   cell**, which insets its content box. A per-cell question with a per-cell
//!   answer.
//! * [`plan_table_borders`] — **what line stands on each line of the grid**,
//!   which is what reaches the page. A collapsed border sits on an edge two
//!   cells share, so it belongs to neither; [`BorderPlan`] is indexed by grid
//!   line and has no notion of an owning cell at all.
//!
//! Both resolve with [`resolve_border_conflict`]; only the pairs they feed it
//! differ. The doc on `resolve_table_cell_borders` says where the two disagree
//! and why that disagreement is still open.
//!
//! Emission follows the same shape as the geometry, and there are two:
//! [`rasterize_border_grid`] paints a collapsed table's grid as junctions and
//! the segments between them — every rect disjoint from every other by
//! construction — while a §17.4.45-spaced table has no grid to collapse onto and
//! takes [`emit_cell_frame`] per cell plus [`emit_table_outline`] for the
//! rectangle its cells no longer reach.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};

use super::grid::{cell_index_at_grid_col, is_vmerge_continue};
use super::types::{
    CellBorderOverride, TableBorderConfig, TableBorderLine, TableBorderStyle, TableCellInput,
    TableRowInput, VerticalMergeState,
};
use crate::render::layout::draw_command::DrawCommand;

/// One cell edge during and after §17.4.38 resolution.
///
/// Three states rather than `Option<TableBorderLine>`, because [MS-OI29500]
/// §17.4.66 distinguishes "nothing said about this edge" from "declared
/// `val="nil"`". The difference is entirely about **inheritance**: an omitted
/// or `none` edge falls back to the table style, then `tblPrEx`, then
/// `tblBorders`; `nil` declines that fallback and stays empty.
///
/// It is *not* about outranking the facing cell. `nil` removes this cell's
/// border and nothing else — see [`resolve_border_conflict`].
///
/// The distinction survives resolution for one downstream reader: the page-split
/// top-border restore in `emit.rs` may revive an `Absent` top but must not
/// revive a `Suppressed` one. For painting they are identical, which is what
/// [`CellEdge::line`] expresses.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(super) enum CellEdge {
    /// Nothing said about this edge — or it was declared `val="none"`, which
    /// §17.4.66 treats identically. Inherits, then yields.
    Absent,
    /// Declared `val="nil"`: no border here, and no inheritance either.
    Suppressed,
    /// A border to resolve against the opposing edge, and paint if it wins.
    Line(TableBorderLine),
}

impl CellEdge {
    /// The line to paint, if any. Both `Absent` and `Suppressed` paint nothing.
    pub(super) fn line(self) -> Option<TableBorderLine> {
        match self {
            Self::Line(l) => Some(l),
            Self::Absent | Self::Suppressed => None,
        }
    }

    /// Whether two *resolved* edges paint the same thing.
    ///
    /// Not `==`: by this point the `Absent`/`Suppressed` distinction is not
    /// observable to the painter, and letting it in would split a run of columns
    /// that paints one continuous line. Callers asking "can one cell draw this
    /// whole span in a single stroke?" mean *this* question.
    pub(super) fn paints_same(self, other: Self) -> bool {
        self.line() == other.line()
    }
}

impl From<Option<TableBorderLine>> for CellEdge {
    /// Table-level borders have no way to express `nil` — an edge is either
    /// configured or not — so an absent one is `Absent`, never `Suppressed`.
    fn from(b: Option<TableBorderLine>) -> Self {
        match b {
            Some(l) => Self::Line(l),
            None => Self::Absent,
        }
    }
}

/// Resolved borders for one cell.
#[derive(Clone)]
pub(super) struct CellBorders {
    pub(super) top: CellEdge,
    pub(super) bottom: CellEdge,
    pub(super) left: CellEdge,
    pub(super) right: CellEdge,
}

/// §17.4.38 / §17.7.6: what each cell *declares* on its four edges, before any
/// question of who paints a shared one.
///
/// The first of the two passes both readers below start from: table borders
/// mapped onto the cell's position in its row via
/// [`resolve_cell_effective_borders`], the cell's own `w:tcBorders` on top, then
/// §17.4.84's two vertical-merge clearings. Nothing here looks across an edge.
struct Declarations {
    /// One per cell, in row order, indexed the same way `rows[r].cells` is.
    cells: Vec<Vec<CellBorders>>,
    /// The grid column each cell starts at, indexed the same way.
    grid_indices: Vec<Vec<usize>>,
}

fn declare_cell_borders(
    rows: &[TableRowInput],
    borders: Option<&TableBorderConfig>,
    cell_spacing: Pt,
) -> Declarations {
    let num_rows = rows.len();
    let mut cells: Vec<Vec<CellBorders>> = Vec::new();
    let mut grid_indices: Vec<Vec<usize>> = Vec::new();
    for (row_idx, row) in rows.iter().enumerate() {
        let mut row_borders = Vec::new();
        let mut row_grid = Vec::new();
        // §17.4.15: gridBefore — the row's first cell starts at grid_col
        // `grid_before`, leaving the leftmost columns empty.
        let mut grid_idx = row.grid_before as usize;
        // §17.4.60: a row may carry per-row border overrides
        // (`<w:tblPrEx><w:tblBorders/></w:tblPrEx>`). When set,
        // it's the *fully merged* effective table borders for this
        // row — the build layer already overlaid the override on
        // the table's own borders so the model-layer
        // "explicitly none" vs "not specified" distinction is
        // preserved during conversion. Use it verbatim; otherwise
        // fall back to the table-wide config.
        let row_table_borders = row.border_overrides.as_ref().or(borders);
        let cells_in_row = row.cells.len();
        for (cell_i, cell_input) in row.cells.iter().enumerate() {
            let span = cell_input.grid_span.max(1) as usize;
            let mut cell_borders = resolve_cell_effective_borders(
                cell_input,
                row_table_borders,
                CellPosition {
                    row: row_idx,
                    num_rows,
                    first_in_row: cell_i == 0,
                    last_in_row: cell_i + 1 == cells_in_row,
                },
                cell_spacing > Pt::ZERO,
            );
            if cell_input.vertical_merge == Some(VerticalMergeState::Continue) {
                cell_borders.top = CellEdge::Absent;
            }
            if row_idx + 1 < num_rows && is_vmerge_continue(&rows[row_idx + 1], grid_idx) {
                cell_borders.bottom = CellEdge::Absent;
            }
            row_borders.push(cell_borders);
            row_grid.push(grid_idx);
            grid_idx += span;
        }
        cells.push(row_borders);
        grid_indices.push(row_grid);
    }
    Declarations {
        cells,
        grid_indices,
    }
}

/// §17.4.66: how much of each edge is charged to each cell — the **measurement**
/// question, and only that one.
///
/// This used to be the whole of border resolution and is now half of it. What it
/// answers is how far a cell's content box is inset by the borders around it
/// (`measure_table_rows`), which is a per-*cell* question and therefore has a
/// per-cell answer. Where the line each edge resolves to actually goes on the
/// page is a different question with a different shape, and
/// [`plan_table_borders`] answers that one — a collapsed border stands on an
/// edge two cells share, so it belongs to neither of them.
///
/// The two are deliberately not folded together, because they disagree and the
/// disagreement is **unsettled**. Here a shared vertical is charged wholly to
/// the cell on its left (the winner is written to that cell's `right` and the
/// facing `left` is cleared); a centred border puts half its width in each. The
/// second reading is what Word's collapsed model implies and what
/// [`plan_table_borders`] paints, but changing what a cell is *charged* moves
/// text in every bordered table in the corpus, and `tests/table_cell_content_box.rs`
/// pins today's rule against a reasoned defect history. **Word reference render
/// needed**: one table whose `w:sz` steps 4 → 48 across otherwise-identical
/// rows, measuring where the first glyph lands, which separates "the full width
/// is inside the cell" from "half of it is".
///
/// `num_grid_cols` is the table-wide grid column count (`col_widths.len()`),
/// which is what makes a cell "at the table edge" rather than merely last in
/// its row (§17.4.15 `gridBefore` separates the two).
pub(super) fn resolve_table_cell_borders(
    rows: &[TableRowInput],
    num_grid_cols: usize,
    borders: Option<&TableBorderConfig>,
    // §17.4.45 `tblCellSpacing`, already resolved to points. Non-zero means the
    // cells share no edges, which decides both the seeding and whether the
    // collapse pass runs at all.
    cell_spacing: Pt,
    // §17.4.38: adjacent-table collapse — see `measure_table_rows`.
    suppress_first_row_top: bool,
) -> ResolvedTableBorders {
    let num_rows = rows.len();
    let Declarations {
        cells: mut resolved_borders,
        grid_indices,
    } = declare_cell_borders(rows, borders, cell_spacing);

    // [MS-OI29500] §17.4.66: *"If the cell spacing is nonzero ... then all
    // cell borders and outer table borders display."* With a gap between
    // them, adjacent cells share no edge, so there is no conflict to
    // resolve and every cell keeps its own four borders. Collapsing them
    // here would delete borders that must be drawn, and drawing both sides
    // of a collapsed edge would double every line.
    let collapse_borders = cell_spacing <= Pt::ZERO;
    if collapse_borders {
        // §17.4.66: conflict resolution at vertical shared edges (a cell's
        // right vs. its right neighbour's left). Drawn once on the left cell.
        for row_idx in 0..num_rows {
            let num_cells = rows[row_idx].cells.len();
            for cell_ci in 0..num_cells.saturating_sub(1) {
                let right = resolved_borders[row_idx][cell_ci].right;
                let left = resolved_borders[row_idx][cell_ci + 1].left;
                let winner = resolve_border_conflict(right, left);
                resolved_borders[row_idx][cell_ci].right = winner;
                resolved_borders[row_idx][cell_ci + 1].left = CellEdge::Absent;
            }
        }

        // [MS-OI29500] §17.4.66: conflict resolution at horizontal shared edges (row R's
        // bottom vs. row R+1's top). Resolved *per grid column* because a
        // `gridSpan` cell in one row can face several cells in the other:
        //   • wide upper cell over several lower cells — resolving only the
        //     first lower cell (and nulling the rest) drops their borders;
        //   • wide lower cell under several upper cells — a nil spacer among
        //     them must not punch a gap through the lower cell's border.
        //
        // The whole edge is then drawn from *one* side (all upper bottoms, or
        // all lower tops). This matters visually: an upper-row bottom sits in
        // the inter-row gap while a lower-row top sits just below it, so
        // splitting a single line between the two sides would offset segments
        // by the border width. A cell paints one border across its width, so
        // a side can own the edge only if each of its cells spans a run of
        // columns whose resolved border is uniform; upper is preferred (it
        // keeps the aligned-grid path and page-split top restoration valid).
        // `saturating_sub`, not `- 1`: an empty table would underflow.
        for upper in 0..num_rows.saturating_sub(1) {
            let lower = upper + 1;

            // Per-column resolved border for this inter-row edge.
            let resolved: Vec<CellEdge> = (0..num_grid_cols)
                .map(|gc| {
                    let edge = |row: usize, pick: fn(&CellBorders) -> CellEdge| {
                        cell_index_at_grid_col(&rows[row], gc)
                            .map(|ci| pick(&resolved_borders[row][ci]))
                            .unwrap_or(CellEdge::Absent)
                    };
                    resolve_border_conflict(edge(upper, |b| b.bottom), edge(lower, |b| b.top))
                })
                .collect();

            // A row can paint the whole edge iff (a) it has a cell over every
            // column that carries a border — a row whose `gridSpan` leaves a
            // bordered column uncovered (its gridAfter gap) can't draw that
            // column, so the other row must — and (b) each of its cells spans
            // a uniform run of resolved columns (a cell paints one border
            // across its width). Without (a), a partly-covered cell would draw
            // its own top *and* the covering row its bottom → a doubled line.
            let can_own = |row_idx: usize| -> bool {
                let covers_bordered_cols = (0..num_grid_cols).all(|gc| {
                    resolved[gc].line().is_none()
                        || cell_index_at_grid_col(&rows[row_idx], gc).is_some()
                });
                covers_bordered_cols
                    && grid_indices[row_idx]
                        .iter()
                        .enumerate()
                        .all(|(ci, &start)| {
                            let span = rows[row_idx].cells[ci].grid_span.max(1) as usize;
                            let end = (start + span).min(num_grid_cols);
                            start >= end
                                || (start..end).all(|gc| resolved[gc].paints_same(resolved[start]))
                        })
            };

            if !can_own(upper) && can_own(lower) {
                // Wide upper cell can't paint the mixed edge; draw it entirely
                // from the finer lower row so the line stays at one y (e.g. a
                // label cell right of a nil spacer under a gridSpan header).
                for (ci, &start) in grid_indices[lower].iter().enumerate() {
                    let span = rows[lower].cells[ci].grid_span.max(1) as usize;
                    let end = (start + span).min(num_grid_cols);
                    if start < end {
                        resolved_borders[lower][ci].top = resolved[start];
                    }
                }
                for b in resolved_borders[upper].iter_mut() {
                    b.bottom = CellEdge::Absent;
                }
            } else {
                // Upper row owns the edge: each upper cell paints its uniform
                // run (a nil spacer above a gridSpan cell resolves to that
                // cell's inherited border, so no gap), and lower tops it
                // covers are cleared. Columns an upper cell can't paint
                // uniformly (only reachable in the both-non-uniform fallback)
                // fall through to the lower cell.
                let mut covered = vec![false; num_grid_cols];
                for (ci, &start) in grid_indices[upper].iter().enumerate() {
                    let span = rows[upper].cells[ci].grid_span.max(1) as usize;
                    let end = (start + span).min(num_grid_cols);
                    if start >= end {
                        continue;
                    }
                    if (start..end).all(|gc| resolved[gc].paints_same(resolved[start])) {
                        resolved_borders[upper][ci].bottom = resolved[start];
                        for c in covered.iter_mut().take(end).skip(start) {
                            *c = true;
                        }
                    } else {
                        resolved_borders[upper][ci].bottom = CellEdge::Absent;
                    }
                }
                for (ci, &start) in grid_indices[lower].iter().enumerate() {
                    let span = rows[lower].cells[ci].grid_span.max(1) as usize;
                    let end = (start + span).min(num_grid_cols);
                    if start >= end {
                        continue;
                    }
                    if (start..end).any(|gc| covered[gc]) {
                        // Any column already painted from above → defer the
                        // whole cell so a partly-covered span can't double up.
                        // The columns it *would* have covered alone become a
                        // band fill below.
                        resolved_borders[lower][ci].top = CellEdge::Absent;
                    } else if (start..end).all(|gc| resolved[gc].paints_same(resolved[start])) {
                        resolved_borders[lower][ci].top = resolved[start];
                        for c in covered.iter_mut().take(end).skip(start) {
                            *c = true;
                        }
                    } else {
                        resolved_borders[lower][ci].top = CellEdge::Absent;
                    }
                }

                // Whatever neither side could paint is simply not charged to
                // either — `covered` is dropped here. It is not lost: the run
                // still resolves to a line, and [`plan_table_borders`] puts that
                // line on the grid boundary it stands on, which has no owning
                // row to fall between. That is the whole of what the `BandFill`
                // machinery this replaced was for.
                let _ = &covered;
            }
        }
    }

    // §17.4.38: suppress first-row top borders for adjacent table collapse.
    if suppress_first_row_top && !resolved_borders.is_empty() {
        for b in &mut resolved_borders[0] {
            b.top = CellEdge::Absent;
        }
    }

    ResolvedTableBorders {
        cells: resolved_borders,
    }
}

/// What §17.4.66 charges to each cell: every cell's four edges.
pub(super) struct ResolvedTableBorders {
    pub(super) cells: Vec<Vec<CellBorders>>,
}

/// §17.4.66: what line stands on each line of the table's border grid — the
/// **painting** question.
///
/// Indexed by grid line, never by cell, and that is the whole point. A collapsed
/// border stands on an edge two cells *share*, so asking a cell to paint it has
/// no well-defined answer: the same line is derivable from both sides, and
/// whichever one paints decides where it sits. Three defects were reported
/// against the per-cell emitter this replaced, all of them squares at a
/// boundary that the cells touching it had each been emptied of.
///
/// Two things that used to need machinery fall out of the indexing:
///
/// * A row gapped by §17.4.15 `gridBefore` above one gapped by §17.4.14
///   `gridAfter` leaves runs of a boundary that *neither row's cells* can paint,
///   because a cell paints one border across its whole width. Here there is no
///   owning row to fall between — `h[r][c]` is per grid column, and a run is a
///   run.
/// * A `w:gridSpan` cell has no vertical inside its own span, so those grid
///   lines are simply `Absent` rather than something a cell has to decline.
pub(super) struct BorderPlan {
    cols: usize,
    rows: usize,
    /// `(cols + 1) * rows`, row-major within each grid line: the vertical on
    /// grid line `c` alongside row `r` is at `c * rows + r`.
    v: Vec<CellEdge>,
    /// `(rows + 1) * cols`, column-major within each boundary: the horizontal on
    /// row boundary `r` over grid column `c` is at `r * cols + c`.
    h: Vec<CellEdge>,
}

impl BorderPlan {
    /// Number of grid columns; there are `cols + 1` vertical grid lines.
    pub(super) fn cols(&self) -> usize {
        self.cols
    }

    /// The vertical on grid line `c` (`0..=cols`), alongside row `r`.
    pub(super) fn vertical(&self, c: usize, r: usize) -> CellEdge {
        if c > self.cols || r >= self.rows {
            return CellEdge::Absent;
        }
        self.v[c * self.rows + r]
    }

    /// The horizontal on row boundary `r` (`0..=rows`), over grid column `c`.
    pub(super) fn horizontal(&self, r: usize, c: usize) -> CellEdge {
        if r > self.rows || c >= self.cols {
            return CellEdge::Absent;
        }
        self.h[r * self.cols + c]
    }
}

/// §17.4.66: resolve the whole border grid of a **collapsed** table.
///
/// Cell spacing is not a parameter because a spaced table has no grid to
/// collapse onto — its cells share no edges, so it takes the other constructor
/// of [`TableBorderGeometry`] and keeps a closed frame per cell.
///
/// Both families resolve the same way and with the same function: each grid
/// segment has at most two declarations facing it, and
/// [`resolve_border_conflict`] picks between them. What differs is only which
/// two.
pub(super) fn plan_table_borders(
    rows: &[TableRowInput],
    num_grid_cols: usize,
    borders: Option<&TableBorderConfig>,
    // §17.4.38: adjacent-table collapse — see `measure_table_rows`.
    suppress_first_row_top: bool,
) -> BorderPlan {
    let num_rows = rows.len();
    let Declarations {
        cells,
        grid_indices,
    } = declare_cell_borders(rows, borders, Pt::ZERO);

    // Verticals. On grid line `c` in row `r`, the two facing declarations are
    // the `right` of the cell whose span *ends* at `c` and the `left` of the
    // cell whose span *starts* there. A line strictly inside a `w:gridSpan`
    // has neither, and a row gapped at that end has only one.
    let mut v = vec![CellEdge::Absent; (num_grid_cols + 1) * num_rows];
    for (r, starts) in grid_indices.iter().enumerate() {
        for (ci, &start) in starts.iter().enumerate() {
            let span = rows[r].cells[ci].grid_span.max(1) as usize;
            let end = (start + span).min(num_grid_cols);
            if start >= end {
                continue;
            }
            let b = &cells[r][ci];
            for (c, edge) in [(start, b.left), (end, b.right)] {
                let slot = &mut v[c * num_rows + r];
                *slot = resolve_border_conflict(*slot, edge);
            }
        }
    }

    // Horizontals. On boundary `r` over grid column `c`, the two facing
    // declarations are the `bottom` of the cell covering `c` in row `r - 1` and
    // the `top` of the cell covering it in row `r`. Either may be missing — at
    // the table's own two boundaries, and wherever a row's cells do not reach
    // that column.
    let mut h = vec![CellEdge::Absent; (num_rows + 1) * num_grid_cols];
    for r in 0..=num_rows {
        for c in 0..num_grid_cols {
            let above = r
                .checked_sub(1)
                .and_then(|up| cell_index_at_grid_col(&rows[up], c).map(|ci| cells[up][ci].bottom))
                .unwrap_or(CellEdge::Absent);
            let below = (r < num_rows)
                .then(|| cell_index_at_grid_col(&rows[r], c).map(|ci| cells[r][ci].top))
                .flatten()
                .unwrap_or(CellEdge::Absent);
            h[r * num_grid_cols + c] = resolve_border_conflict(above, below);
        }
    }

    // §17.4.38: adjacent-table collapse removes the whole of the table's top
    // boundary, the one the table above already painted as its bottom.
    if suppress_first_row_top {
        for e in h.iter_mut().take(num_grid_cols) {
            *e = CellEdge::Absent;
        }
    }

    BorderPlan {
        cols: num_grid_cols,
        rows: num_rows,
        v,
        h,
    }
}

/// Which edges of the table, and of its own row, one cell touches.
///
/// Deliberately **not** a grid position. The border rule turned out to be about
/// the row's cells rather than the grid: a row's first `<w:tc>` takes the
/// table's `w:left` whether or not it starts at grid column 0, because nothing
/// faces its leading edge either way. Carrying `col`/`span`/`num_grid_cols`
/// here invited the grid-column reading that two Word renders have now refuted,
/// and nothing read them once the rule was right.
#[derive(Clone, Copy)]
pub(super) struct CellPosition {
    pub(super) row: usize,
    pub(super) num_rows: usize,
    /// Whether this cell is the first `<w:tc>` of its `<w:tr>` — *not* whether
    /// it starts at grid column 0. The two differ exactly when §17.4.15
    /// `gridBefore` gaps the row, and the border rule follows this one.
    pub(super) first_in_row: bool,
    /// The same for the last `<w:tc>`, which §17.4.14 `gridAfter` — or a row
    /// whose cells simply do not reach the grid's end — separates from reaching
    /// the last grid column.
    pub(super) last_in_row: bool,
}

/// §17.4.38 / §17.7.6: resolve effective borders for a cell.
/// Per-cell borders (from conditional formatting) override table-level borders.
/// Table-level insideH/insideV are mapped to cell edges based on position.
pub(super) fn resolve_cell_effective_borders(
    cell: &TableCellInput,
    table_borders: Option<&TableBorderConfig>,
    at: CellPosition,
    // §17.4.45: whether this table has a non-zero `w:tblCellSpacing`. See the
    // `outer` closure below — it is the whole reason this parameter exists.
    spaced: bool,
) -> CellBorders {
    // Start with table-level borders mapped to cell edges.
    let tb = table_borders;
    let is_first_row = at.row == 0;
    // `row + 1 == num_rows`, not `row == num_rows - 1`: the latter underflows on
    // an empty table. No caller passes `num_rows == 0` today, but the field is
    // free and the guard would live entirely in the callers.
    let is_last_row = at.row + 1 == at.num_rows;

    // §17.4.45 / issue #168: with a non-zero cell spacing the outer edges are
    // **not** seeded from the table's own borders. A spaced cell is inset from
    // the table's boundary, so a table border painted on it lands in the wrong
    // place — and once `emit_table_outline` draws that border where it belongs,
    // seeding it here as well would paint it twice.
    //
    // The interior seeding is deliberately left alone. With a gap there is no
    // shared edge for `insideH`/`insideV` to sit on either, so what they mean
    // for a spaced table is a real question — but [MS-OI29500] §17.4.66 names
    // only "cell borders and outer table borders", and answering a second
    // unsettled question inside this one is how a fix stops being reviewable.
    //
    // **Word reference render needed** (issue #165 has the batch): a spaced
    // table with `insideV` set whose cells carry no `w:tcBorders`. If Word
    // draws one line per cell edge, today's behaviour is right; if it draws one
    // line in the gap, or none, this seeding has to change too.
    let outer = |line: Option<TableBorderLine>| -> CellEdge {
        if spaced {
            CellEdge::Absent
        } else {
            line.into()
        }
    };
    let mut top: CellEdge = if is_first_row {
        outer(tb.and_then(|b| b.top))
    } else {
        tb.and_then(|b| b.inside_h).into()
    };
    let mut bottom: CellEdge = if is_last_row {
        outer(tb.and_then(|b| b.bottom))
    } else {
        tb.and_then(|b| b.inside_h).into()
    };
    // §17.4.36 against §17.4.15/§17.4.14: the question is about the row's
    // **cells**, not about the grid.
    //
    // §17.4.66 resolves a cell edge against "cell borders and outer table
    // borders". A row's first `<w:tc>` has no cell facing its leading edge —
    // §17.4.15 `gridBefore` leaves those grid columns blank — so the table's own
    // border is what faces it, wherever across the grid that edge happens to
    // fall. `gridBefore` moves where the edge *is*, not what it is. Same for the
    // last cell and §17.4.14 `gridAfter`.
    //
    // So this is deliberately keyed on the cell's index in its row and not on
    // its grid column, and the two differ exactly when a row is gapped.
    // **Measured, not reasoned**: Word renders `grid-gap-borders.docx` with a
    // 3pt red `w:left` on the leading edge of `D`, `F` and `G` and a 3pt green
    // `w:right` on the trailing edge of `E` and `F`. It renders the same edge
    // bare in `bidi-visual-table.docx`, whose `w:left` is `nil` — the same rule,
    // since `nil` paints nothing.
    //
    // Two other readings were held here and both were wrong, which is why the
    // grid-column form is not worth trying again: `insideV` (grid columns exist
    // to the left, so the edge is interior) is refuted by the `nil` render, and
    // *nothing at all* (§17.4.35 places `w:left` "around the table", and this
    // edge is 50pt inside it) is refuted by this one. The second was argued from
    // the spec's wording and fit every measurement available at the time.
    let mut left: CellEdge = if at.first_in_row {
        outer(tb.and_then(|b| b.left))
    } else {
        tb.and_then(|b| b.inside_v).into()
    };
    let mut right: CellEdge = if at.last_in_row {
        outer(tb.and_then(|b| b.right))
    } else {
        tb.and_then(|b| b.inside_v).into()
    };

    // Per-cell overrides. Only `nil` and a real border reach here — an explicit
    // `none` was mapped to "no override" upstream (§17.4.66: it inherits
    // exactly like an omitted edge), so it correctly leaves the table-level
    // border above untouched instead of erasing it.
    if let Some(ref cb) = cell.cell_borders {
        if let Some(v) = &cb.top {
            top = resolve_override(v);
        }
        if let Some(v) = &cb.bottom {
            bottom = resolve_override(v);
        }
        if let Some(v) = &cb.left {
            left = resolve_override(v);
        }
        if let Some(v) = &cb.right {
            right = resolve_override(v);
        }
    }

    CellBorders {
        top,
        bottom,
        left,
        right,
    }
}

/// Resolve a border conflict between two competing borders on a shared edge.
/// Returns the winning border (or `None` if both are `None`).
///
/// The algorithm is **not in ISO/IEC 29500-1** — the standard only says a method
/// exists. It is spelled out in [MS-OI29500] §17.4.66 (`tcBorders`, note a),
/// which is the authority for every step below:
///   1. An edge with no border yields to one that has it. `none` counts as
///      no border, per *"If the conflicting table cell border is `none` (no
///      border), then the opposing border shall be displayed."*
///   2. Weight = width in eighths of a point × style number. Higher wins.
///   3. Equal weight: the style **earlier in the spec's precedence list** wins —
///      `Single` over `Double`. See `style_precedence_index`.
///   4. Equal style: darker colour wins (`R+B+2G`, then `B+2G`, then `G`).
///
/// **What `nil` does, and what it does not.** The note adds *"If the conflicting
/// table cell border is `nil`, then no border shall be displayed"*, which reads
/// as `nil` beating everything on the far side of the edge. It does not, and
/// implementing it that way deleted borders Word draws. `nil` acts on **its own
/// cell only**: it is how a cell declines the inheritance the note describes one
/// step earlier (style → `tblPrEx` → `tblBorders`), which is the whole of its
/// difference from `none`. The facing cell's border is untouched, so
/// `Suppressed` yields here exactly like `Absent`.
///
/// Three independent facts in `IP 05 Trenches` fix the reading, and no evidence
/// contradicts it:
///
/// * a cell declaring `<w:bottom w:val="single"/>` above one declaring
///   `<w:top w:val="nil"/>` — Word draws the line, as does macOS's own DOCX
///   renderer on the same markup;
/// * a cell that *inherits* its bottom from `insideH`, faced by a `gridSpan=2`
///   spacer cell whose `nil` was aimed at the neighbouring column — Word draws
///   that line too, and it could not do otherwise: a cell paints one border
///   across its whole width, so a wide cell's `nil` cannot punch a hole in the
///   cell above it;
/// * down the document's spacer columns the generator writes `nil` on **both**
///   sides of every shared edge. Writing both is only necessary because one
///   alone does not suppress.
///
/// `nil` is still not a no-op: with nothing facing it — a table's outer edge, or
/// a facing cell that is also `nil` — declining inheritance is exactly what
/// removes the border. Both halves are pinned by
/// `tests/table_border_conflict.rs`.
///
/// **The comparison is a total order, and that is the point.** The caller feeds
/// this (upper row's bottom, lower row's top) and (left cell's right, right
/// cell's left), so a rule that stops at step 2 leaves the winner decided by
/// *which side of the edge a border came from* — an implementation detail. Ties
/// used to fall through to whichever argument came first, which meant an
/// equal-weight 3pt single beat a 1pt double or lost to it depending on
/// argument order, and of two equal borders differing only in colour the paler
/// one won half the time. `resolve_border_conflict(a, b)` now always equals
/// `resolve_border_conflict(b, a)`.
///
/// Suppression is still a *third* state, which is why the argument type is
/// [`CellEdge`] and not `Option<TableBorderLine>`: when neither side paints,
/// returning `Suppressed` rather than `Absent` keeps the two distinguishable for
/// the caller — a suppressed edge must not be revived by the page-split
/// top-border restore in `emit.rs`, whereas an absent one should be.
pub(super) fn resolve_border_conflict(a: CellEdge, b: CellEdge) -> CellEdge {
    match (a, b) {
        (CellEdge::Line(la), CellEdge::Line(lb)) => {
            match border_precedence(&la).cmp(&border_precedence(&lb)) {
                std::cmp::Ordering::Less => b,
                _ => a,
            }
        }
        // One side paints: it does so regardless of what the other side says.
        // A facing `nil` removed *its* border, not this one.
        (CellEdge::Line(_), _) => a,
        (_, CellEdge::Line(_)) => b,
        // Neither side paints. Carry suppression forward so the page-split
        // restore cannot revive an edge the author explicitly emptied.
        (CellEdge::Suppressed, _) | (_, CellEdge::Suppressed) => CellEdge::Suppressed,
        (CellEdge::Absent, CellEdge::Absent) => CellEdge::Absent,
    }
}

/// Sort key for [MS-OI29500] §17.4.66 conflict resolution — greater wins.
///
/// Returns integers so the key is `Ord`: comparing `f32` weights directly would
/// need `partial_cmp`, and a `NaN` width (unreachable, but the type permits it)
/// would silently make the comparison non-transitive and reintroduce the
/// order-dependence this exists to remove.
///
/// **Two fields are inverted, and for the same reason.** The spec states both
/// style and colour as "lower value wins" rankings — earliest in the precedence
/// list, and smallest brightness. This key is "greater wins", so each is
/// subtracted from its type's maximum. Inverting one and not the other is the
/// defect this layout is meant to make obvious.
fn border_precedence(b: &TableBorderLine) -> (u32, u8, u32, u32, u32) {
    let (l0, l1, l2) = colour_luminance(b);
    (
        // Weight in eighths of a point, rounded — the spec's `sz` unit.
        (border_weight(b) * 8.0).round().max(0.0) as u32,
        u8::MAX - style_precedence_index(b.style),
        u32::MAX - l0,
        u32::MAX - l1,
        u32::MAX - l2,
    )
}

/// [MS-OI29500] §17.4.66 style precedence: at equal weight, *"the higher of the
/// two on this precedence list shall be displayed"*, the list being
///
/// > single, thick, double, dotted, dashed, dotDash, dotDotDash, triple,
/// > thinThickSmallGap, … outset, inset
///
/// "Higher on the list" means **earlier**, so this returns the 0-based index
/// into it and **lower wins** — `border_precedence` inverts it.
///
/// So `Single` beats `Double` at equal weight, which is worth stating plainly
/// because the intuition runs the other way: a double border has the greater
/// *style number* (3 vs 1) and therefore the greater weight at equal width, and
/// it is easy to carry that ordering into the tie-break, where the spec
/// reverses it. Equal weight means the single is three times wider — a 3pt
/// solid line against two 0.33pt hairlines — and the spec prefers the single.
///
/// Only `Single` and `Double` reach layout (the other 24 §17.18.2 `ST_Border` styles are
/// approximated as `Single` — see `convert_model_border`), so only their two
/// positions are modelled: single is first, double is third.
fn style_precedence_index(style: TableBorderStyle) -> u8 {
    match style {
        TableBorderStyle::Single => 0,
        TableBorderStyle::Double => 2,
    }
}

/// [MS-OI29500] §17.4.66 darkness keys, compared in order: `R+B+2G`, then
/// `B+2G`, then `G`. Lower is darker.
fn colour_luminance(b: &TableBorderLine) -> (u32, u32, u32) {
    let (r, g, bl) = (b.color.r as u32, b.color.g as u32, b.color.b as u32);
    (r + bl + 2 * g, bl + 2 * g, g)
}

/// Where the line on one placed boundary comes from.
///
/// A page slice does not always put the plan's two neighbours on either side of
/// a y, so which of the plan's boundaries a placed one shows is not simply its
/// row index. Two shapes do it: the seam under a §17.4.49 repeated header, whose
/// next row is not the header's successor and which therefore shows the
/// *header's own* lower boundary; and a continuation slice's first row, whose
/// predecessor is on the page before.
#[derive(Clone, Copy)]
pub(super) struct BoundarySource {
    /// Which of the plan's boundaries (`0..=rows`) holds the line for this y.
    pub(super) plan_boundary: usize,
    /// §17.4.38: the line to draw where that boundary is `Absent`, for a
    /// continuation slice whose top edge conflict resolution gave to the page
    /// before. `None` everywhere else.
    ///
    /// Only `Absent` falls through to it. An edge the author set to `nil` stays
    /// empty — they asked for no border — and [`CellEdge`]'s third state is what
    /// keeps the two distinguishable this far down.
    pub(super) restore: Option<TableBorderLine>,
}

/// One row as a page slice placed it: which plan row it is, and the two
/// boundaries it sits between.
///
/// The boundaries are y values in table-local coordinates, and consecutive
/// placed rows share one — `placed[i].bottom == placed[i + 1].top` — so each is
/// rasterized once.
pub(super) struct PlacedRow {
    pub(super) plan_row: usize,
    pub(super) top: Pt,
    pub(super) bottom: Pt,
    /// Where the line on the boundary at `top` comes from. The one at `bottom`
    /// belongs to the row below, except on the slice's last row, where
    /// [`rasterize_border_grid`] takes the plan's boundary under it — the cut
    /// closes the table off exactly as its own edge would.
    pub(super) top_source: BoundarySource,
}

/// The width a resolved edge paints, zero where it paints nothing.
fn edge_width(e: CellEdge) -> Pt {
    e.line().map(|l| l.width).unwrap_or(Pt::ZERO)
}

/// §17.4.66: paint one page slice's share of a table's border grid.
///
/// **Every rect this emits is disjoint from every other, by construction**, and
/// together they cover the whole network. That is the property three reported
/// corner defects were each a violation of, and it is here a consequence of the
/// decomposition rather than something the code has to be argued into:
///
/// * a **junction** is the neighbourhood of a node — one grid line crossed with
///   one row boundary;
/// * a **horizontal segment** lives in the open x-interval *between* two nodes
///   on one boundary;
/// * a **vertical segment** lives in the open y-interval between two nodes on
///   one grid line.
///
/// Each family is disjoint from itself (different nodes, different intervals),
/// and from the other two — a node's neighbourhood is exactly what the segments
/// have removed from theirs, which is `subtract` below rather than a trim at
/// each segment's own two ends, because a junction can be wider than the column
/// beside it.
///
/// The single exception is two *segments* on parallel lines closer together than
/// the lines are thick, which is the author's geometry being impossible rather
/// than this decomposition's: see `is_parallel_crowding` in
/// `tests/table_border_corners.rs`, which audits the invariant over the whole
/// corpus and allows exactly that.
///
/// Where a border sits on its line — straddling a shared one, inside one shared
/// with nothing — is decided by `inside`, below.
///
/// `x` holds the `cols + 1` vertical grid lines; `placed` is this slice's rows,
/// top to bottom.
pub(super) fn rasterize_border_grid(
    commands: &mut Vec<DrawCommand>,
    plan: &BorderPlan,
    x: &[Pt],
    placed: &[PlacedRow],
    // The slice's own box. Only the four lines bounding it are affected — see
    // `inside` below.
    box_size: PtSize,
) {
    if placed.is_empty() || x.len() < 2 {
        return;
    }
    let cols = plan.cols();

    // The boundaries this slice paints, top to bottom: one above each placed
    // row, then one below the last. Which of the plan's boundaries each one
    // shows is `BoundarySource`'s question, not this loop's — a continuation
    // slice and a repeated header both break the correspondence.
    let lines_of = |source: BoundarySource| -> Vec<CellEdge> {
        (0..cols)
            .map(|c| match plan.horizontal(source.plan_boundary, c) {
                CellEdge::Absent => source.restore.into(),
                resolved => resolved,
            })
            .collect()
    };
    let mut boundaries: Vec<(Pt, Vec<CellEdge>)> = Vec::with_capacity(placed.len() + 1);
    for row in placed {
        boundaries.push((row.top, lines_of(row.top_source)));
    }
    // The slice's foot is the plan's boundary under its last row — a page cut
    // closes the table off with the same line its own edge would, which is what
    // the row above the cut would have painted had the next row followed it.
    let last = &placed[placed.len() - 1];
    boundaries.push((
        last.bottom,
        lines_of(BoundarySource {
            plan_boundary: last.plan_row + 1,
            restore: None,
        }),
    ));

    // §17.4.66 / issue #157: a row of **zero height** puts two boundaries at one
    // y, and an empty `<w:tr/>` — a row with no cells at all — is the shape that
    // does it. Two boundaries at one y are one boundary: resolve the lower into
    // the upper and leave it empty, exactly as two cells facing across a shared
    // edge are resolved. Painting both put the same rect on the page twice and
    // left its colour to emission order.
    //
    // This settles only the double-*paint*. Whether an empty row should separate
    // its neighbours at all — giving that boundary two lines a row apart rather
    // than one — is what `test-files/issue-157-empty-row-edge.docx` asks, and it
    // is still **open**: the row has no height here, so the question does not
    // arise geometrically.
    for b in 1..boundaries.len() {
        if boundaries[b].0 != boundaries[b - 1].0 {
            continue;
        }
        let lower = std::mem::replace(&mut boundaries[b].1, vec![CellEdge::Absent; cols]);
        for (c, edge) in lower.into_iter().enumerate() {
            let upper = boundaries[b - 1].1[c];
            boundaries[b - 1].1[c] = resolve_border_conflict(upper, edge);
        }
    }

    // Width of the horizontal reaching each node from its left and its right.
    // A node at grid line `c` on boundary `b` is met by columns `c - 1` and `c`.
    let h_at_node = |b: usize, c: usize| -> Pt {
        let lines = &boundaries[b].1;
        let left = c
            .checked_sub(1)
            .map(|i| lines[i])
            .unwrap_or(CellEdge::Absent);
        let right = lines.get(c).copied().unwrap_or(CellEdge::Absent);
        edge_width(left).max(edge_width(right))
    };
    // The same for the verticals reaching it from above and below — the placed
    // rows on either side of the boundary, which need not be plan-adjacent.
    let v_at_node = |b: usize, c: usize| -> Pt {
        let above = b
            .checked_sub(1)
            .map(|i| plan.vertical(c, placed[i].plan_row))
            .unwrap_or(CellEdge::Absent);
        let below = placed
            .get(b)
            .map(|r| plan.vertical(c, r.plan_row))
            .unwrap_or(CellEdge::Absent);
        edge_width(above).max(edge_width(below))
    };

    // §17.4.66: **a border straddles a line two cells share, and sits inside an
    // unshared one.** The interior of the grid is all shared edges, so a border
    // there is centred on its line and a 1pt `insideV` comes out concentric with
    // a 3pt `w:left` meeting it. The four lines bounding the slice are shared
    // with nothing — there is no cell beyond them — so a border on one of them
    // goes *inside*, which is the same rule [`emit_cell_frame`] applies to every
    // edge of a spaced table, where no edge is shared at all.
    //
    // Centring those four as well is what this fixed. The table's ink then ran
    // half a border past its own declared box on every side: `TableSlice::size`
    // stopped containing what the slice draws, and §17.4.63's auto-width guard —
    // which is drawn at the *paper* edge — let 0.2pt of a full-width table off
    // the page (`tests/table_auto_width.rs`).
    //
    // **Word reference render needed** for the half this does not settle: whether
    // Word's own table box contains its outer borders or straddles them, which
    // decides whether `w:tblInd` measures to the border's outer edge or to its
    // centre. Both readings keep the ink on the paper, so nothing here is
    // evidence between them; a table with `w:left` at 0.5pt and at 6pt, measuring
    // the first column's text x, would be.
    //
    // Expressed against the slice's box rather than against the *index* of the
    // outermost line, and the difference is a case the index gets wrong: a page
    // cut's boundary is the last one this slice paints but sits half a reserved
    // strip above the slice's foot, so it is interior after all and its border
    // belongs centred on it. Shift, never clamp — a border keeps the width it
    // declared and moves, so a row too short to hold its own borders still
    // paints both of them (at full width, overlapping) rather than two slivers.
    let inside = |centre: Pt, w: Pt, limit: Pt| -> Pt {
        let lo = centre - w * 0.5;
        if lo < Pt::ZERO {
            Pt::ZERO
        } else if lo + w > limit {
            limit - w
        } else {
            lo
        }
    };

    // 1. Junctions, resolved but **not yet emitted**. A node needs one only
    //    where both axes reach it: with one axis alone the segments on either
    //    side lost nothing to it and already abut at the node.
    //
    //    Held rather than emitted because two things need them. A segment is
    //    what the junctions it runs into leave of its interval, so they must all
    //    be known before any segment is cut. And a junction is emitted *among*
    //    the segments of the axis that won it — see the passes below.
    struct Junction {
        node: (usize, usize),
        x: (Pt, Pt),
        y: (Pt, Pt),
        line: TableBorderLine,
        along_vertical: bool,
    }
    let mut junctions: Vec<Junction> = Vec::new();
    for (b, (y, _)) in boundaries.iter().enumerate() {
        for (c, &gx) in x.iter().enumerate() {
            let (vw, hw) = (v_at_node(b, c), h_at_node(b, c));
            if vw <= Pt::ZERO || hw <= Pt::ZERO {
                continue;
            }
            let Some((line, along_vertical)) = junction_line(plan, placed, &boundaries, b, c)
            else {
                continue;
            };
            let (jx, jy) = (
                inside(gx, vw, box_size.width),
                inside(*y, hw, box_size.height),
            );
            junctions.push(Junction {
                node: (b, c),
                x: (jx, jx + vw),
                y: (jy, jy + hw),
                line,
                along_vertical,
            });
        }
    }
    let emit_junction = |commands: &mut Vec<DrawCommand>, j: &Junction| {
        emit_border_rect(
            commands,
            &j.line,
            PtRect::from_xywh(j.x.0, j.y.0, j.x.1 - j.x.0, j.y.1 - j.y.0),
            !j.along_vertical,
        );
    };
    let junction_at = |b: usize, c: usize, vertical: bool| -> Option<&Junction> {
        junctions
            .iter()
            .find(|j| j.node == (b, c) && j.along_vertical == vertical)
    };

    // 2 and 3. The segments: each one's interval **minus every junction whose
    //    square it runs into**.
    //
    //    Subtraction against the junctions it actually meets, not a trim at its
    //    own two ends, and the difference is not academic. A junction is as wide
    //    as the vertical standing in it, so a grid line closer to its neighbour
    //    than half that width *reaches past* it — into the next column's
    //    horizontal, or into the next grid line's vertical, neither of which a
    //    two-ended trim would touch. Both shapes are real rather than contrived:
    //    the spacer columns of Word's own `MediumShading` table styles are 0.7pt
    //    wide against 3pt borders, which is six overlaps on one page of
    //    `sample-docx-files-sample1.docx` and eighteen more across the local
    //    corpus.
    //
    //    The one overlap this cannot remove is between two *segments* — two
    //    parallel lines whose boundaries are closer together than the lines are
    //    thick. That is the author's geometry being impossible (a `hRule="exact"`
    //    row shorter than its own borders, or an empty `<w:tr/>`), not the
    //    decomposition's, and `tests/table_border_corners.rs` allows exactly it.
    let cuts_across = |band: (Pt, Pt), vertical: bool| -> Vec<(Pt, Pt)> {
        junctions
            .iter()
            .filter(|j| {
                let across = if vertical { j.x } else { j.y };
                across.1 > band.0 && band.1 > across.0
            })
            .map(|j| if vertical { j.y } else { j.x })
            .collect()
    };

    // Each pass walks its own axis **in order**, emitting each junction it owns
    // just before the segment that abuts it. The order is not cosmetic: a
    // junction and the segment beside it are usually the same colour and always
    // share an edge, and `coalesce_abutting_rects` fuses such a pair only when
    // the two are *consecutive* commands. Emitted in two separate passes they
    // never are, and every one of them reaches the page as a seam under a
    // rasterizer that anti-aliases each fill on its own —
    // `tests/table_shading_seams.rs` is that defect's audit and caught exactly
    // this.
    for (b, (y, lines)) in boundaries.iter().enumerate() {
        for (c, edge) in lines.iter().enumerate() {
            if let Some(j) = junction_at(b, c, false) {
                emit_junction(commands, j);
            }
            let Some(line) = edge.line() else { continue };
            let y0 = inside(*y, line.width, box_size.height);
            let band = (y0, y0 + line.width);
            for (x0, x1) in subtract(x[c], x[c + 1], &cuts_across(band, false)) {
                emit_border_rect(
                    commands,
                    &line,
                    PtRect::from_xywh(x0, band.0, x1 - x0, line.width),
                    true,
                );
            }
        }
        if let Some(j) = junction_at(b, lines.len(), false) {
            emit_junction(commands, j);
        }
    }

    for (c, &gx) in x.iter().enumerate() {
        for (b, row) in placed.iter().enumerate() {
            if let Some(j) = junction_at(b, c, true) {
                emit_junction(commands, j);
            }
            let Some(line) = plan.vertical(c, row.plan_row).line() else {
                continue;
            };
            let x0 = inside(gx, line.width, box_size.width);
            let band = (x0, x0 + line.width);
            for (y0, y1) in subtract(row.top, row.bottom, &cuts_across(band, true)) {
                emit_border_rect(
                    commands,
                    &line,
                    PtRect::from_xywh(band.0, y0, line.width, y1 - y0),
                    false,
                );
            }
        }
        if let Some(j) = junction_at(placed.len(), c, true) {
            emit_junction(commands, j);
        }
    }
}

/// `[a, b]` with every interval in `cuts` removed, as the surviving runs.
///
/// `cuts` need be neither sorted nor disjoint, and runs of zero or negative
/// length are dropped — a segment entirely covered by junctions yields nothing,
/// which is exactly right for a grid column narrower than the borders at its
/// two ends.
fn subtract(a: Pt, b: Pt, cuts: &[(Pt, Pt)]) -> Vec<(Pt, Pt)> {
    let mut runs = vec![(a, b)];
    for &(c0, c1) in cuts {
        let mut next = Vec::with_capacity(runs.len() + 1);
        for (r0, r1) in runs {
            if c1 <= r0 || c0 >= r1 {
                next.push((r0, r1));
                continue;
            }
            if r0 < c0 {
                next.push((r0, c0));
            }
            if c1 < r1 {
                next.push((c1, r1));
            }
        }
        runs = next;
    }
    runs.retain(|&(r0, r1)| r1 - r0 > Pt::ZERO);
    runs
}

/// Which of the (up to four) borders meeting at a node paints its square.
///
/// **ECMA-376 does not settle this, and neither does [MS-OI29500].** The
/// standard specifies no stroke geometry at all, and §17.4.66's precedence
/// list is about *conflicting* declarations on one edge — a junction is not a
/// conflict, since all four segments meeting there are correct and all four
/// want the square.
///
/// So the rule is this engine's, and it is chosen to be the one ordering the
/// spec does supply rather than an invented one: the square goes to the
/// [`border_precedence`] winner among the segments that reach it. Heavier wins,
/// then the earlier style, then the darker colour; identical lines are
/// indistinguishable and the choice between them is not observable.
///
/// **Word reference render needed** to confirm it: a crossing where a red
/// vertical and a blue horizontal tie on weight, reading the square's colour off
/// the page. `test-files/grid-gap-borders.docx` is one 3pt horizontal away from
/// being that probe.
///
/// The bool says which **axis** the square is drawn along, which matters only
/// for §17.18.2 `double`: a double splits into two sub-lines across its own
/// short side, so a square inherited from a vertical must split side by side and
/// one from a horizontal must stack. At equal precedence the vertical takes it,
/// which is a tie-break with no evidence behind it — and the case that exposes
/// the limit is a double crossing a double, where the honest answer is neither
/// axis but a 2 × 2 lattice of ink with the two gaps running through it. This
/// engine draws two sub-lines there and not four squares.
fn junction_line(
    plan: &BorderPlan,
    placed: &[PlacedRow],
    boundaries: &[(Pt, Vec<CellEdge>)],
    b: usize,
    c: usize,
) -> Option<(TableBorderLine, bool)> {
    let lines = &boundaries[b].1;
    let incident = [
        (c.checked_sub(1).and_then(|i| lines[i].line()), false),
        (lines.get(c).and_then(|e| e.line()), false),
        (
            b.checked_sub(1)
                .and_then(|i| plan.vertical(c, placed[i].plan_row).line()),
            true,
        ),
        (
            placed
                .get(b)
                .and_then(|r| plan.vertical(c, r.plan_row).line()),
            true,
        ),
    ];
    incident
        .into_iter()
        .filter_map(|(l, vertical)| l.map(|l| (l, vertical)))
        .max_by_key(|(l, vertical)| (border_precedence(l), *vertical))
}

/// §17.4.45: the four borders of one cell, drawn inside its own box.
///
/// The **spaced** constructor's emitter, and only that one. [MS-OI29500]
/// §17.4.66: *"If the cell spacing is nonzero ... then all cell borders and
/// outer table borders display."* With a gap between them adjacent cells share
/// no edge, so there is no grid line for a border to stand on and nothing to
/// centre — each cell keeps its four borders wholly inside itself, and the
/// table's own rectangle is drawn separately by [`emit_table_outline`].
///
/// Every corner square of the box is painted by exactly one of the two edges
/// that meet there: the horizontals own the corners, because they span the full
/// cell width, and the verticals fill only what is left between them. Both
/// halves are load-bearing — painting a corner twice lets the second rect win it
/// when the two edges differ in colour, and painting it not at all leaves a hole
/// one border wide.
pub(super) fn emit_cell_frame(commands: &mut Vec<DrawCommand>, b: &CellBorders, cell: CellBox) {
    // Resolution is over by now, so `Suppressed` and `Absent` are the same
    // thing here: nothing to paint.
    let (top, bottom) = (b.top.line(), b.bottom.line());
    let top_w = top.map(|l| l.width).unwrap_or(Pt::ZERO);
    let bot_w = bottom.map(|l| l.width).unwrap_or(Pt::ZERO);

    // A spaced row reserves no strip below it — the gap between rows *is* the
    // spacing — so a bottom border is always inset into the cell's own foot.
    let (top_y, bottom_y) = (cell.y, cell.y + cell.h - bot_w);

    if let Some(ref line) = top {
        emit_border_rect(
            commands,
            line,
            PtRect::from_xywh(cell.x, top_y, cell.w, top_w),
            true,
        );
    }
    if let Some(ref line) = bottom {
        emit_border_rect(
            commands,
            line,
            PtRect::from_xywh(cell.x, bottom_y, cell.w, bot_w),
            true,
        );
    }

    let (v_top, v_bottom) = (top_y + top_w, bottom_y);
    let v_height = v_bottom - v_top;
    if v_height <= Pt::ZERO {
        return;
    }
    if let Some(line) = b.left.line() {
        let rect = PtRect::from_xywh(cell.x, v_top, line.width, v_height);
        emit_border_rect(commands, &line, rect, false);
    }
    if let Some(line) = b.right.line() {
        let rect = PtRect::from_xywh(cell.x + cell.w - line.width, v_top, line.width, v_height);
        emit_border_rect(commands, &line, rect, false);
    }
}

/// One cell's box, as [`emit_cell_frame`] needs it.
#[derive(Clone, Copy)]
pub(super) struct CellBox {
    pub(super) x: Pt,
    pub(super) w: Pt,
    /// Top of the row's content box.
    pub(super) y: Pt,
    /// Height of the row's content box.
    pub(super) h: Pt,
}

/// §17.4.45 / issue #168: draw the table's own outer border, for a table whose
/// `w:tblCellSpacing` is non-zero.
///
/// [MS-OI29500] §17.4.66: *"If the cell spacing is nonzero ... then all cell
/// borders and outer table borders display."* Everywhere else in this engine a
/// table border exists only as a **cell** edge, which is exactly right while the
/// spacing is zero — the outer cells' edges are then the table's edges. Once
/// there is a gap they are not, and nothing else in the pipeline draws the
/// table's own rectangle.
///
/// `rect` is the slice's box in table-local coordinates. `draw_top` and
/// `draw_bottom` are false on the sides where a paginated table continues:
/// an intermediate slice ends at a page cut, not at the table's edge, so it
/// gets left and right only.
///
/// Geometry mirrors [`emit_cell_frame`] exactly — horizontals span the full
/// width and own the corners, verticals are inset between them — so an outline
/// and a cell edge of the same width meet the same way a cell edge meets its
/// neighbour.
pub(super) fn emit_table_outline(
    commands: &mut Vec<DrawCommand>,
    borders: Option<&TableBorderConfig>,
    rect: PtRect,
    draw_top: bool,
    draw_bottom: bool,
) {
    let Some(cfg) = borders else {
        return;
    };
    let top = if draw_top { cfg.top } else { None };
    let bottom = if draw_bottom { cfg.bottom } else { None };
    let (x, y) = (rect.origin.x, rect.origin.y);
    let (w, h) = (rect.size.width, rect.size.height);

    let top_w = top.map(|b| b.width).unwrap_or(Pt::ZERO);
    let bot_w = bottom.map(|b| b.width).unwrap_or(Pt::ZERO);

    if let Some(ref border) = top {
        emit_border_rect(commands, border, PtRect::from_xywh(x, y, w, top_w), true);
    }
    if let Some(ref border) = bottom {
        emit_border_rect(
            commands,
            border,
            PtRect::from_xywh(x, y + h - bot_w, w, bot_w),
            true,
        );
    }

    let v_height = h - top_w - bot_w;
    if v_height > Pt::ZERO {
        if let Some(ref border) = cfg.left {
            emit_border_rect(
                commands,
                border,
                PtRect::from_xywh(x, y + top_w, border.width, v_height),
                false,
            );
        }
        if let Some(ref border) = cfg.right {
            emit_border_rect(
                commands,
                border,
                PtRect::from_xywh(x + w - border.width, y + top_w, border.width, v_height),
                false,
            );
        }
    }
}

/// [MS-OI29500] §17.4.66: border weight = width × style number, in points.
///
/// The spec states the rule in eighths of a point (`w:sz`), but every use is a
/// *comparison* between two weights, and converting both to eighths scales both
/// by the same 8 — so the factor cancels. Keeping it in points avoids implying
/// that a unit conversion is load-bearing here. `border_precedence` scales to
/// eighths once, where rounding to an integer sort key does depend on the unit.
fn border_weight(b: &TableBorderLine) -> f32 {
    let style_number = match b.style {
        TableBorderStyle::Single => 1.0,
        TableBorderStyle::Double => 3.0,
    };
    b.width.raw() * style_number
}

/// Width of the line this edge paints, or zero when it paints none — which
/// includes a suppressed edge, since suppression reserves no space.
pub(super) fn border_width(b: CellEdge) -> Pt {
    b.line().map(|b| b.width).unwrap_or(Pt::ZERO)
}

fn resolve_override(ovr: &CellBorderOverride) -> CellEdge {
    match ovr {
        CellBorderOverride::Suppress => CellEdge::Suppressed,
        // The cell's own `<w:tcBorders>` — the provenance that beats a facing
        // `nil` in `resolve_border_conflict`.
        CellBorderOverride::Border(line) => CellEdge::Line(*line),
    }
}

/// Emit a border as filled rectangle(s).
/// `is_horizontal` controls double-border sub-rect orientation.
fn emit_border_rect(
    commands: &mut Vec<DrawCommand>,
    b: &TableBorderLine,
    rect: PtRect,
    is_horizontal: bool,
) {
    match b.style {
        TableBorderStyle::Single => {
            commands.push(DrawCommand::Rect {
                rect,
                color: b.color,
            });
        }
        TableBorderStyle::Double => {
            // §17.4.38: total = w:sz, each sub-line = sz/3, gap = sz/3.
            let sub = b.width * (1.0 / 3.0);
            if is_horizontal {
                // Two horizontal sub-rects: top and bottom of the border area.
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(rect.origin.x, rect.origin.y, rect.size.width, sub),
                    color: b.color,
                });
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(
                        rect.origin.x,
                        rect.origin.y + rect.size.height - sub,
                        rect.size.width,
                        sub,
                    ),
                    color: b.color,
                });
            } else {
                // Two vertical sub-rects: left and right of the border area.
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(rect.origin.x, rect.origin.y, sub, rect.size.height),
                    color: b.color,
                });
                commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(
                        rect.origin.x + rect.size.width - sub,
                        rect.origin.y,
                        sub,
                        rect.size.height,
                    ),
                    color: b.color,
                });
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::render::dimension::Pt;
    use crate::render::fonts::Toggle;
    use crate::render::geometry::PtEdgeInsets;
    use crate::render::layout::draw_command::DrawCommand;
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
    use crate::render::layout::table::TableSlice;
    use crate::render::layout::table::{
        layout_table, CellVAlign, TableBorderConfig, TableBorderLine, TableBorderStyle,
        TableCellInput, TableRowInput,
    };
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after: crate::render::layout::fragment::fixture_break_after(text),
            font: Rc::new(FontProps {
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            }),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
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
                floating_shapes: vec![],
            }],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
            cell_borders: None,
            vertical_merge: None,
            vertical_align: CellVAlign::Top,
        }
    }

    /// Every border rect of a 1×2 table, at its exact position — the whole
    /// decomposition [`rasterize_border_grid`] produces, written out.
    ///
    /// The network is three grid lines (x = 0, 100, 200) crossing two boundaries
    /// (y = 0, 15), and it comes out as **13** rects in three families that
    /// tile it exactly: 6 junction squares at the 6 nodes, 4 horizontal segments
    /// in the gaps between nodes along the two boundaries, 3 vertical segments
    /// in the gaps along the three grid lines. A count alone could not tell a
    /// correct 13 from a wrong one, so every number is derived: the columns are
    /// 100pt and every border 0.5pt, so a border straddles its line by 0.25 each
    /// side; the row is 15pt — one 14pt default line plus the 0.5pt each of the
    /// two horizontals is *charged* for, since `resolve_table_cell_borders`
    /// insets the content box by the full width even though only half of it lies
    /// inside (see that function on why the two disagree and why the charging
    /// half is the unsettled one).
    ///
    /// [MS-OI29500] §17.4.66: the shared vertical at x = 100 appears **once**.
    /// That is the property this file's older shape asserted as "seven rects,
    /// not eight", and it survives the decomposition — a per-cell emitter could
    /// paint it twice, and the grid has nowhere to put a second one.
    ///
    /// Note the outer borders reach 0.25pt outside `result.size` on all four
    /// sides, because a border centred on the table's own edge is half outside
    /// it. Whether Word's table box contains its outer borders or straddles them
    /// is **open** — see `resolve_table_cell_borders` — and this test pins
    /// today's answer rather than endorsing it.
    #[test]
    fn borders_emit_lines() {
        let line = TableBorderLine {
            width: Pt::new(0.5),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("a"), simple_cell("b")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            Some(&TableBorderConfig {
                top: Some(line),
                bottom: Some(line),
                left: Some(line),
                right: Some(line),
                inside_h: Some(line),
                inside_v: Some(line),
            }),
            None,
            false,
        );

        assert_eq!(
            result.size,
            crate::render::geometry::PtSize::new(Pt::new(200.0), Pt::new(15.0))
        );
        // Every border is 0.5pt (`w`) and half of one is `h`. The four lines
        // bounding the table go **inside** it; the shared line at x = 100
        // straddles. So the grid's four ordinates on the page are:
        let (w, h) = (0.5_f32, 0.25_f32);
        let (left, mid, right) = (0.0, 100.0 - h, 200.0 - w);
        let (top, bottom) = (0.0, 15.0 - w);
        // The order is asserted along with the geometry, and is not incidental:
        // each junction is emitted **among the segments of the axis that won
        // it**, immediately before the one it abuts, so `coalesce_abutting_rects`
        // can fuse the pair. Every border here is identical, so every junction
        // is a tie and every tie goes to the vertical — hence four horizontals
        // first, then each grid line as junction / segment / junction.
        assert_eq!(
            rects(&result.commands),
            vec![
                // The two boundaries, one segment per column, each stopping
                // where the junctions at its ends begin.
                (left + w, top, mid - left - w, w),
                (mid + w, top, right - mid - w, w),
                (left + w, bottom, mid - left - w, w),
                (mid + w, bottom, right - mid - w, w),
                // Then the three grid lines, each with the junction at its head,
                // its segment, and the junction at its foot. The one at x = 100
                // is the shared edge, drawn once.
                (left, top, w, w),
                (left, top + w, w, bottom - top - w),
                (left, bottom, w, w),
                (mid, top, w, w),
                (mid, top + w, w, bottom - top - w),
                (mid, bottom, w, w),
                (right, top, w, w),
                (right, top + w, w, bottom - top - w),
                (right, bottom, w, w),
            ],
        );
    }

    /// §17.4.60 tblPrEx — when a row carries a `tblBorders` override,
    /// it fully replaces the table's tblBorders for *that row only*.
    /// Here row 0 sets every side to "no border", row 1 doesn't.
    /// The table-wide config has all sides set to single. Expectation:
    /// row 0's cell contributes zero border rects, while row 1's cell
    /// produces the usual top/left/right/bottom set.
    #[test]
    fn row_border_override_replaces_table_borders_for_that_row() {
        let single = TableBorderLine {
            width: Pt::new(0.5),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let all_single = TableBorderConfig {
            top: Some(single),
            bottom: Some(single),
            left: Some(single),
            right: Some(single),
            inside_h: Some(single),
            inside_v: Some(single),
        };
        let no_borders = TableBorderConfig {
            top: None,
            bottom: None,
            left: None,
            right: None,
            inside_h: None,
            inside_v: None,
        };
        let rows = vec![
            TableRowInput {
                cells: vec![simple_cell("opt-out")],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: Some(no_borders),
            },
            TableRowInput {
                cells: vec![simple_cell("normal")],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
            },
        ];
        let col_widths = vec![Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            Some(&all_single),
            None,
            false,
        );

        // Group border rects by their y position. The opt-out row is
        // first (lower y), the normal row second. We know the order
        // because layout_table walks rows top-down.
        let border_rects: Vec<_> = result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color } if *color == RgbColor::BLACK => Some(*rect),
                _ => None,
            })
            .collect();

        // No rect should sit entirely within row 0's vertical span —
        // not the cell's top, not its sides, not its bottom (with row 1
        // having a top border, conflict resolution gives row 0 a
        // bottom from row 1's top, but that's drawn at the boundary,
        // not inside row 0).
        // We exercise this by asserting that no rect's *vertical*
        // extent falls within (epsilon, row_0_height - epsilon) — the
        // strict interior of row 0.
        let row_0_height = Pt::new(14.0);
        let interior_eps = Pt::new(0.1);
        let interior_top = interior_eps;
        let interior_bottom = row_0_height - interior_eps;
        for rect in &border_rects {
            let r_top = rect.origin.y;
            let r_bottom = rect.origin.y + rect.size.height;
            let entirely_inside = r_top >= interior_top && r_bottom <= interior_bottom;
            assert!(
                !entirely_inside,
                "row 0 (border-override = all None) must not host a \
                 black border rect strictly inside its content area; got rect \
                 y=[{:.2}..{:.2}] (interior was ({:.2}..{:.2}))",
                r_top.raw(),
                r_bottom.raw(),
                interior_top.raw(),
                interior_bottom.raw(),
            );
        }
    }

    // ── issue #168: the outer table border of a spaced table ────────────────

    fn all_borders(width: f32) -> TableBorderConfig {
        let line = TableBorderLine {
            width: Pt::new(width),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        TableBorderConfig {
            top: Some(line),
            bottom: Some(line),
            left: Some(line),
            right: Some(line),
            inside_h: Some(line),
            inside_v: Some(line),
        }
    }

    /// One `Rect` command as `(x, y, w, h)`.
    type R = (f32, f32, f32, f32);

    /// Every `Rect` command, flattened to plain numbers so a failing assertion
    /// prints geometry rather than a wall of `Pt` wrappers.
    fn rects(cmds: &[DrawCommand]) -> Vec<R> {
        cmds.iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, .. } => Some((
                    rect.origin.x.raw(),
                    rect.origin.y.raw(),
                    rect.size.width.raw(),
                    rect.size.height.raw(),
                )),
                _ => None,
            })
            .collect()
    }

    /// The **union** of the ink a ray meets, as sorted disjoint intervals.
    /// `vertical` sends the ray down at `x = at` and returns y intervals;
    /// otherwise it runs right at `y = at` and they are x.
    ///
    /// This is how a claim about *what a band contains* is asked of a command
    /// stream that is free to split every line at its junctions: the ray sees
    /// the union, so the decomposition is invisible to it and only the geometry
    /// is asserted. Touching intervals merge, because two abutting rects are one
    /// line to any reader — which is exactly what a junction and its two
    /// segments are.
    fn ink_along(cmds: &[DrawCommand], at: f32, vertical: bool) -> Vec<(f32, f32)> {
        let mut runs: Vec<(f32, f32)> = rects(cmds)
            .into_iter()
            .filter_map(|(x, y, w, h)| {
                let (across, along) = if vertical {
                    ((x, x + w), (y, y + h))
                } else {
                    ((y, y + h), (x, x + w))
                };
                (across.0 <= at && at <= across.1).then_some(along)
            })
            .collect();
        runs.sort_by(|a, b| a.0.total_cmp(&b.0));
        let mut merged: Vec<(f32, f32)> = Vec::new();
        for (a, b) in runs {
            match merged.last_mut() {
                Some(last) if a <= last.1 + 1e-4 => last.1 = last.1.max(b),
                _ => merged.push((a, b)),
            }
        }
        merged
    }

    fn two_rows() -> Vec<TableRowInput> {
        (0..2)
            .map(|_| TableRowInput {
                cells: vec![simple_cell("a"), simple_cell("b")],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
            })
            .collect()
    }

    const NEAR: f32 = 0.01;

    /// The defect. [MS-OI29500] §17.4.66: *"If the cell spacing is nonzero ...
    /// then all cell borders and outer table borders display."* Before this
    /// fix nothing was ever drawn at the table's own bounds — table borders
    /// existed only as cell edges, which are inset by the spacing.
    #[test]
    fn a_spaced_table_draws_its_outer_border_at_the_table_bounds() {
        let cfg = all_borders(1.0);
        let result = layout_table(
            &two_rows(),
            &[Pt::new(100.0), Pt::new(100.0)],
            Pt::new(20.0),
            Pt::new(14.0),
            Some(&cfg),
            None,
            false,
        );
        let (w, h) = (result.size.width.raw(), result.size.height.raw());
        let r = rects(&result.commands);

        assert!(
            r.iter().any(|t| t.1.abs() < NEAR && (t.2 - w).abs() < NEAR),
            "no top outline spanning the table width at y=0; rects={r:?} (w={w}, h={h})"
        );
        assert!(
            r.iter()
                .any(|t| (t.1 + t.3 - h).abs() < NEAR && (t.2 - w).abs() < NEAR),
            "no bottom outline at the table's bottom edge; rects={r:?} (w={w}, h={h})"
        );
        assert!(
            r.iter().any(|t| t.0.abs() < NEAR && t.3 > h * 0.5),
            "no left outline down the table's left edge; rects={r:?} (w={w}, h={h})"
        );
        assert!(
            r.iter()
                .any(|t| (t.0 + t.2 - w).abs() < NEAR && t.3 > h * 0.5),
            "no right outline down the table's right edge; rects={r:?} (w={w}, h={h})"
        );
    }

    /// The guarantee the whole corpus rests on. With no spacing the outer
    /// cells' edges *are* the table's edges, the existing mapping is correct,
    /// and this fix must not add a single rect.
    ///
    /// Two rows deliberately: in a one-row table the row's own height equals
    /// the table's, so a full-height vertical rect would be ambiguous. With two
    /// rows only an outline can span the whole table.
    #[test]
    fn a_zero_spacing_table_draws_no_outline() {
        let cfg = all_borders(1.0);
        let result = layout_table(
            &two_rows(),
            &[Pt::new(100.0), Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&cfg),
            None,
            false,
        );
        let h = result.size.height.raw();
        let spanning: Vec<_> = rects(&result.commands)
            .into_iter()
            .filter(|t| t.3 > h * 0.9 && t.2 < 5.0)
            .collect();
        assert!(
            spanning.is_empty(),
            "a zero-spacing table must draw no table-height outline, got {spanning:?}"
        );
    }

    /// A spaced table split across pages: left and right bound every slice,
    /// but the table's top edge exists only where the table starts and its
    /// bottom edge only where it ends. An intermediate slice stops at a page
    /// cut, not at the table's boundary, and drawing a horizontal rule there
    /// would draw a table edge that does not exist.
    #[test]
    fn a_paginated_spaced_table_splits_its_outline_across_slices() {
        use crate::render::layout::table::{layout_table_paginated, TablePaginationConfig};

        let cfg = all_borders(1.0);
        // Six rows against a short page, so the table needs at least three
        // slices and therefore has a middle one with neither horizontal edge.
        let rows: Vec<TableRowInput> = (0..6)
            .map(|_| TableRowInput {
                cells: vec![simple_cell("a")],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
            })
            .collect();
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(100.0)],
            Pt::new(20.0),
            Pt::new(14.0),
            Some(&cfg),
            None,
            &TablePaginationConfig {
                available_height: Pt::new(80.0),
                page_height: Pt::new(80.0),
                suppress_first_row_top: false,
            },
        );
        assert!(
            slices.len() >= 3,
            "need a middle slice to test; got {}",
            slices.len()
        );

        let last = slices.len() - 1;
        for (i, slice) in slices.iter().enumerate() {
            let (w, h) = (slice.size.width.raw(), slice.size.height.raw());
            let r = rects(&slice.commands);
            let spans_width = |y: f32| {
                r.iter()
                    .any(|t| (t.1 - y).abs() < NEAR && (t.2 - w).abs() < NEAR)
            };

            assert_eq!(
                spans_width(0.0),
                i == 0,
                "slice {i}: top edge should be present only on the first slice"
            );
            let bottom_present = r
                .iter()
                .any(|t| (t.1 + t.3 - h).abs() < NEAR && (t.2 - w).abs() < NEAR);
            assert_eq!(
                bottom_present,
                i == last,
                "slice {i}: bottom edge should be present only on the last slice"
            );
            assert!(
                r.iter().any(|t| t.0.abs() < NEAR && t.3 > h * 0.4),
                "slice {i}: left edge must bound every slice; rects={r:?}"
            );
            assert!(
                r.iter()
                    .any(|t| (t.0 + t.2 - w).abs() < NEAR && t.3 > h * 0.4),
                "slice {i}: right edge must bound every slice; rects={r:?}"
            );
        }
    }

    /// The other half of §17.4.66's sentence: the cells' own borders keep
    /// drawing, at the cells' rectangles, alongside the outline. A fix that
    /// moved the border out to the table bounds and dropped the cell edges
    /// would satisfy the first test and still be wrong.
    #[test]
    fn a_spaced_table_draws_cell_borders_as_well_as_the_outline() {
        let cfg = all_borders(1.0);
        let result = layout_table(
            &two_rows(),
            &[Pt::new(100.0), Pt::new(100.0)],
            Pt::new(20.0),
            Pt::new(14.0),
            Some(&cfg),
            None,
            false,
        );
        let w = result.size.width.raw();
        // A rect that touches neither the left nor the right table edge can
        // only belong to a cell.
        let interior = rects(&result.commands)
            .into_iter()
            .filter(|t| t.0 > NEAR && (t.0 + t.2) < w - NEAR)
            .count();
        assert!(
            interior > 0,
            "the cells' own borders vanished; only the outline is left"
        );
    }

    /// A cell of exactly `width × height` with no content, so a row's geometry
    /// follows from `RowHeightRule::Exact` alone and not from text metrics.
    fn sized_row(height: f32, cells: usize) -> TableRowInput {
        TableRowInput {
            cells: (0..cells)
                .map(|_| TableCellInput {
                    blocks: vec![],
                    ..simple_cell("")
                })
                .collect(),
            height_rule: Some(crate::render::layout::table::RowHeightRule::Exact(Pt::new(
                height,
            ))),
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
        }
    }

    fn double(width: f32) -> TableBorderLine {
        TableBorderLine {
            width: Pt::new(width),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Double,
        }
    }

    // ── §17.18.2 `double`: two sub-lines, not one line ───────────────────────

    /// The declared `w:sz` of a `double` border is the **total** width of the
    /// pair — two lines of `sz/3` separated by a gap of `sz/3` — not the width
    /// of each line. So a 3pt double is two 1pt lines a 1pt apart, filling
    /// exactly the same 3pt band a 3pt `single` would fill, and
    /// [`border_width`] reports 3pt for both.
    ///
    /// The split is along the edge's **short** axis: a horizontal border's two
    /// lines are stacked vertically and each spans the full cell width, a
    /// vertical border's sit side by side and each spans the full edge height.
    /// Getting that backwards would produce two lines running the wrong way
    /// through the border band, and a test that only counted rects would not
    /// notice.
    ///
    /// One 100 × 20pt cell (`Exact` keeps the arithmetic free of text metrics),
    /// every edge a 3pt double.
    ///
    /// Asserted by *probing* the ink rather than by listing rects, because the
    /// rect list is a property of the grid decomposition (which splits every
    /// line at its junctions) and this claim is not: it is about what the band
    /// contains, at any one point along it. Probing across a boundary's middle
    /// asks exactly that question and does not care how many pieces the line
    /// arrived in.
    #[test]
    fn a_double_border_paints_two_sub_lines_of_a_third_the_declared_width() {
        let d = double(3.0);
        let result = layout_table(
            &[sized_row(20.0, 1)],
            &[Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&TableBorderConfig {
                top: Some(d),
                bottom: Some(d),
                left: Some(d),
                right: Some(d),
                inside_h: None,
                inside_v: None,
            }),
            None,
            false,
        );

        // The ink intervals a ray at `x` meets going down, and vice versa.
        let down = |x: f32| ink_along(&result.commands, x, true);
        let across = |y: f32| ink_along(&result.commands, y, false);

        // Down the middle of the cell: the two horizontals, each two 1pt lines a
        // 1pt apart. These are the table's own boundaries, so each 3pt band goes
        // inside — 0..3 at the top, 17..20 at the foot of a 20pt row.
        assert_eq!(
            down(50.0),
            vec![(0.0, 1.0), (2.0, 3.0), (17.0, 18.0), (19.0, 20.0)],
            "a horizontal double splits along its short axis — stacked, not side by side"
        );
        // Across the middle of the row: the two verticals, split the other way,
        // in the bands 0..3 and 97..100.
        assert_eq!(
            across(10.0),
            vec![(0.0, 1.0), (2.0, 3.0), (97.0, 98.0), (99.0, 100.0)],
            "a vertical double splits along *its* short axis"
        );
    }

    /// The band a `double` occupies is the band `border_width` reserves for it,
    /// so a double and a single of the same `w:sz` meet their neighbours
    /// identically — only the interior of the band differs.
    ///
    /// Asserted against the single as a control rather than against literals:
    /// the claim is a relation between the two styles, and pinning the double's
    /// coordinates alone could not tell "same band" from "same numbers I typed
    /// twice". Two halves, and both are needed — the outer extent is equal, and
    /// the double paints exactly two thirds of it, which is what says the gap is
    /// `sz/3` and not merely that a gap exists.
    #[test]
    fn a_double_border_occupies_the_same_band_as_a_single_of_the_same_width() {
        let render = |style: TableBorderStyle| {
            let line = TableBorderLine {
                width: Pt::new(3.0),
                color: RgbColor::BLACK,
                style,
            };
            layout_table(
                &[sized_row(20.0, 1)],
                &[Pt::new(100.0)],
                Pt::ZERO,
                Pt::new(14.0),
                Some(&TableBorderConfig {
                    top: Some(line),
                    bottom: Some(line),
                    left: Some(line),
                    right: Some(line),
                    inside_h: None,
                    inside_v: None,
                }),
                None,
                false,
            )
        };
        let (dbl, sgl) = (
            render(TableBorderStyle::Double),
            render(TableBorderStyle::Single),
        );

        // Same four bands: the outer extent of the ink a ray meets is identical,
        // in both directions and at a point on every one of the four lines.
        for (probe, at) in [(true, 50.0_f32), (false, 10.0_f32)] {
            let extent = |slice: &TableSlice| {
                let runs = ink_along(&slice.commands, at, probe);
                (runs.first().map(|r| r.0), runs.last().map(|r| r.1))
            };
            assert_eq!(
                extent(&dbl),
                extent(&sgl),
                "same band, probe vertical={probe}"
            );
        }

        // And two thirds of it is ink: `sz/3` line, `sz/3` gap, `sz/3` line.
        let ink =
            |slice: &TableSlice| -> f32 { rects(&slice.commands).iter().map(|r| r.2 * r.3).sum() };
        let (a, b) = (ink(&dbl), ink(&sgl));
        assert!(
            (a - b * 2.0 / 3.0).abs() < 1e-3,
            "a double paints two thirds of the band a single fills: {a} vs {b}"
        );
    }

    // ── §17.4.66 at a vertical edge, across a `w:gridSpan` boundary ──────────

    /// Within a row, cell `ci`'s right edge and cell `ci+1`'s left edge always
    /// meet at the same grid column, however wide either cell is — the walk
    /// advances by `grid_span`, so cell adjacency and grid adjacency are the
    /// same question here. (They are *not* at a horizontal edge, which is why
    /// that pass resolves per grid column instead.)
    ///
    /// What a `gridSpan` changes is where the edge lands and how many there
    /// are: the grid boundary **inside** the span carries no vertical at all,
    /// and the one boundary that survives is resolved once and drawn by the
    /// left cell.
    ///
    /// Three 50pt columns, row `[span-2 | plain]`. The span cell's own right is
    /// `insideV` at 0.5pt; the plain cell declares a 2pt left. §17.4.66 step 2
    /// gives the edge to the heavier border, once — the whole point being that
    /// the two cells do not each draw their own.
    ///
    /// Asserted as the interval each vertical covers, expressed from the grid
    /// line it stands on: a shared line is straddled and the table's own two are
    /// gone inside, so what identifies a border is its line plus which of the
    /// two it is. Writing bare origins instead made a test about *which* borders
    /// exist fail whenever their thickness moved.
    #[test]
    fn a_gridspan_cell_resolves_one_vertical_edge_at_its_far_side() {
        let mut wide = TableCellInput {
            blocks: vec![],
            ..simple_cell("")
        };
        wide.grid_span = 2;
        let mut narrow = TableCellInput {
            blocks: vec![],
            ..simple_cell("")
        };
        narrow.cell_borders = Some(crate::render::layout::table::CellBorderConfig {
            top: None,
            bottom: None,
            left: Some(crate::render::layout::table::CellBorderOverride::Border(
                TableBorderLine {
                    width: Pt::new(2.0),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                },
            )),
            right: None,
        });

        let outer = TableBorderLine {
            width: Pt::new(1.0),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let rows = vec![TableRowInput {
            cells: vec![wide, narrow],
            ..sized_row(20.0, 0)
        }];
        let result = layout_table(
            &rows,
            &[Pt::new(50.0), Pt::new(50.0), Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&TableBorderConfig {
                top: None,
                bottom: None,
                left: Some(outer),
                right: Some(outer),
                inside_h: None,
                inside_v: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    ..outer
                }),
            }),
            None,
            false,
        );

        let verticals: Vec<(f32, f32)> = rects(&result.commands)
            .into_iter()
            .filter(|&(_, _, w, h)| h > w)
            .map(|(x, _, w, _)| (x, w))
            .collect();
        assert_eq!(
            verticals,
            vec![
                // The table's own left edge, on grid line 0 and wholly inside it.
                (0.0, 1.0),
                // The one interior vertical: the 2pt winner, straddling grid
                // line 2 at x = 100 — the span cell's far side.
                (100.0 - 1.0, 2.0),
                // The table's own right edge, on grid line 3 and inside it.
                (150.0 - 1.0, 1.0),
            ],
            "nothing may be painted at x = 50 — that grid boundary is interior \
             to the span — and the shared edge is drawn once, not once per cell"
        );
    }

    // ── A row shorter than its own borders (§17.4.80 `hRule="exact"`) ────────

    /// `<w:trHeight w:hRule="exact"/>` sets the row height outright, so a row
    /// can be declared shorter than the borders it carries. Borders are drawn
    /// *inward from the cell edge*, which for a 2pt row with 3pt horizontals
    /// means the two of them overlap and there is nothing between them.
    ///
    /// Two things follow, and both are the point of this test. No rect is
    /// emitted with a non-positive extent — the guard on each segment's length
    /// is what stops the verticals from becoming inverted rectangles, which the
    /// painter would render as nothing or as a smear depending on the backend.
    /// And the two horizontals still paint at full width, so neither declared
    /// border is silently dropped.
    ///
    /// The verticals *are* dropped, and there is nowhere to put them: the span
    /// they would occupy has negative height. What a renderer should instead do
    /// with a row shorter than its borders — grow it, or clip the borders into
    /// it — is not something §17.4.80 or §17.4.66 settles, and this test
    /// deliberately does not pin an answer to it.
    ///
    /// **This is the one case where two border rects legitimately overlap**, and
    /// the only exception to the invariant the rasterizer otherwise guarantees
    /// by construction (`tests/table_border_corners.rs`). The two boundaries are
    /// 2pt apart and each carries a 3pt line centred on it, so the bands cross
    /// whatever the model: it is the author's geometry that is impossible, not
    /// the decomposition's. Any audit of the overlap invariant has to allow it.
    #[test]
    fn a_row_shorter_than_its_own_borders_drops_no_horizontal_and_inverts_nothing() {
        let thick = TableBorderLine {
            width: Pt::new(3.0),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let thin = TableBorderLine {
            width: Pt::new(1.0),
            ..thick
        };
        let result = layout_table(
            &[sized_row(2.0, 1)],
            &[Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&TableBorderConfig {
                top: Some(thick),
                bottom: Some(thick),
                left: Some(thin),
                right: Some(thin),
                inside_h: None,
                inside_v: None,
            }),
            None,
            false,
        );
        let r = rects(&result.commands);

        assert!(
            r.iter().all(|(_, _, w, h)| *w > 0.0 && *h > 0.0),
            "no rect may be emitted with a non-positive extent, got {r:?}"
        );
        // Every rect belongs to one of the two boundary bands, each 3pt tall and
        // pushed inside the 2pt row from its own end — 0..3 from the top, −1..2
        // from the foot. Nothing taller, which is what says the verticals were
        // dropped rather than emitted clamped or inverted.
        assert!(
            r.iter()
                .all(|&(_, y, _, h)| h == 3.0 && (y == 0.0 || y == -1.0)),
            "only the two 3pt boundary bands may be painted, got {r:?}"
        );
        // And both are there at full width. The segments stop where the
        // junctions at their ends begin (the 1pt `w:left`/`w:right`), and the
        // junctions fill exactly that, so each band reaches 0..100 as a union.
        for (label, y) in [("top", 0.5_f32), ("bottom", 1.5_f32)] {
            assert_eq!(
                ink_along(&result.commands, y, false),
                vec![(0.0, 100.0)],
                "the {label} border spans the cell, junctions included"
            );
        }
        // The two bands overlap, because 2pt of row cannot hold 3pt of border on
        // each of its edges. Asserted rather than tolerated — see the doc above.
        assert_eq!(
            ink_along(&result.commands, 50.0, true),
            vec![(-1.0, 3.0)],
            "the two bands cross, so a ray down the cell meets one 4pt run"
        );
    }

    // ── §17.4.45 / issue #168: a spaced table, at exact coordinates ──────────

    /// Every rect a spaced table emits, derived from its inputs.
    ///
    /// Two 100pt slots at a 20pt `w:tblCellSpacing`. The slots were already
    /// shrunk by `build/table.rs::reserve_cell_spacing`, so the table's own
    /// width is 200 + 20 = 220, and each cell is inset one spacing: cell 0 at
    /// x = 20 with width 80, cell 1 at x = 120. Each row box is one 14pt line,
    /// plus the 1pt horizontal border that lies inside it, plus its own 20pt
    /// leading gap; the table adds one trailing gap at its bottom edge:
    /// 35 + 35 + 20 = 90.
    ///
    /// Each row holds exactly *one* horizontal border and not two, which is why
    /// 35 and not 36: with a spacing the table's own top and bottom belong to
    /// the outline, so row 0 has only its `insideH` bottom and row 1 only its
    /// `insideH` top.
    ///
    /// [MS-OI29500] §17.4.66 — *"if the cell spacing is nonzero … all cell
    /// borders and outer table borders display"* — so both appear, and neither
    /// stands in for the other: the outline is the table's own 220 × 88
    /// rectangle, the cell edges are the cells'. The outline is emitted last,
    /// after every row.
    ///
    /// §17.4.38: **every one of a cell's borders is inside its own box**, which
    /// with a spacing is the whole of the cell — there is no shared edge and no
    /// band reserved for one, as `measure_table_rows` says where it declines to
    /// reserve it. Row 0's bottom therefore sits at 34..35, flush with the
    /// inside of its box, exactly as its top sits flush with the inside of the
    /// other end. It used to be drawn at 34..35 of a box that ended at 34, one
    /// width *below* the box and so inside the 20pt spacing, because the
    /// emitter extended a cell's border box by its own bottom border's width
    /// whenever a row followed it — a rule meant for the band between two rows
    /// that share an edge, applied where there is no band and no shared edge.
    /// A spacing narrower than the border would have put a cell's bottom border
    /// inside the next row's box.
    ///
    /// Being inside the box is also why the box has to be tall enough to hold
    /// it. The row is 15pt of box for 14pt of line, and the border occupies the
    /// last point of it rather than the last point of the content.
    #[test]
    fn a_spaced_table_paints_its_cell_edges_and_its_outline_at_exact_coordinates() {
        let result = layout_table(
            &two_rows(),
            &[Pt::new(100.0), Pt::new(100.0)],
            Pt::new(20.0),
            Pt::new(14.0),
            Some(&all_borders(1.0)),
            None,
            false,
        );
        assert_eq!(
            result.size,
            crate::render::geometry::PtSize::new(Pt::new(220.0), Pt::new(90.0))
        );

        assert_eq!(
            rects(&result.commands),
            vec![
                // Row 0 (box 20..35, its bottom border inside it). Its top and
                // left are the table's own edges, which a spaced cell does not
                // take — they belong to the outline.
                (20.0, 34.0, 80.0, 1.0),  // cell 0 bottom (insideH)
                (99.0, 20.0, 1.0, 14.0),  // cell 0 right  (insideV), inset above it
                (120.0, 34.0, 80.0, 1.0), // cell 1 bottom
                (120.0, 20.0, 1.0, 14.0), // cell 1 left
                // Row 1 (box 55..70). Its bottom is the table's own edge.
                (20.0, 55.0, 80.0, 1.0),  // cell 0 top (insideH)
                (99.0, 56.0, 1.0, 14.0),  // cell 0 right, inset under its top
                (120.0, 55.0, 80.0, 1.0), // cell 1 top
                (120.0, 56.0, 1.0, 14.0), // cell 1 left
                // The table's own rectangle, drawn the way a cell's edges are:
                // horizontals span the full width and own the corners, the
                // verticals fill the 88pt between them.
                (0.0, 0.0, 220.0, 1.0),
                (0.0, 89.0, 220.0, 1.0),
                (0.0, 1.0, 1.0, 88.0),
                (219.0, 1.0, 1.0, 88.0),
            ],
        );
    }
}

#[cfg(test)]
mod conflict_tests {
    use super::*;
    use crate::render::resolve::color::RgbColor;

    const BLACK: RgbColor = RgbColor { r: 0, g: 0, b: 0 };
    const PALE: RgbColor = RgbColor {
        r: 220,
        g: 220,
        b: 220,
    };

    fn line(width: f32, style: TableBorderStyle, color: RgbColor) -> TableBorderLine {
        TableBorderLine {
            width: Pt::new(width),
            color,
            style,
        }
    }

    /// A representative spread: both styles, several widths, both colours.
    fn sample_borders() -> Vec<TableBorderLine> {
        let mut v = Vec::new();
        for &w in &[0.5f32, 1.0, 2.0, 3.0, 6.0] {
            for &s in &[TableBorderStyle::Single, TableBorderStyle::Double] {
                for &c in &[BLACK, PALE] {
                    v.push(line(w, s, c));
                }
            }
        }
        v
    }

    /// **The property that matters.** The caller passes (upper row's bottom,
    /// lower row's top) and (left cell's right, right cell's left), so a
    /// resolution that depends on argument order makes the rendered border
    /// depend on which *side of the edge* it was declared on.
    ///
    /// Before this was a total order, ties fell through to whichever argument
    /// came first: an equal-weight 3pt single beat a 1pt double or lost to it
    /// depending on the call, and of two borders differing only in colour the
    /// paler one won half the time.
    #[test]
    fn resolution_is_independent_of_argument_order() {
        let borders = sample_borders();
        for a in &borders {
            for b in &borders {
                let ab = resolve_border_conflict(CellEdge::Line(*a), CellEdge::Line(*b));
                let ba = resolve_border_conflict(CellEdge::Line(*b), CellEdge::Line(*a));
                assert_eq!(
                    (ab.line().map(|x| (x.width, x.style, x.color))),
                    (ba.line().map(|x| (x.width, x.style, x.color))),
                    "order-dependent for {a:?} vs {b:?}"
                );
            }
        }
    }

    /// Step 2 — the heavier border wins outright.
    #[test]
    fn heavier_weight_wins() {
        let thin = line(0.5, TableBorderStyle::Single, BLACK);
        let thick = line(2.0, TableBorderStyle::Single, BLACK);
        assert_eq!(
            resolve_border_conflict(CellEdge::Line(thin), CellEdge::Line(thick))
                .line()
                .map(|b| b.width),
            Some(Pt::new(2.0))
        );
        assert_eq!(
            resolve_border_conflict(CellEdge::Line(thick), CellEdge::Line(thin))
                .line()
                .map(|b| b.width),
            Some(Pt::new(2.0))
        );
    }

    /// Step 3 — equal weight, so position in the spec's precedence list decides,
    /// and **`Single` wins**. 3pt single and 1pt double both weigh 3
    /// (width × style number), which is exactly the tie the pre-E5b#2 code
    /// resolved by argument position.
    ///
    /// This test previously asserted the opposite, and was mutation-checked in
    /// that state — the code and the assertion shared one error, so no mutation
    /// could expose it. [MS-OI29500] §17.4.66 orders the list
    /// `single, thick, double, …` and displays *"the higher of the two on this
    /// precedence list"*, i.e. the earlier one.
    #[test]
    fn equal_weight_prefers_the_earlier_style_in_the_precedence_list() {
        let single = line(3.0, TableBorderStyle::Single, BLACK);
        let double = line(1.0, TableBorderStyle::Double, BLACK);
        assert_eq!(
            border_weight(&single),
            border_weight(&double),
            "same weight"
        );

        for (a, b) in [(single, double), (double, single)] {
            assert_eq!(
                resolve_border_conflict(CellEdge::Line(a), CellEdge::Line(b))
                    .line()
                    .map(|x| x.style),
                Some(TableBorderStyle::Single),
                "Single is earlier in the precedence list, so it wins at equal weight"
            );
        }
    }

    /// The tie-break must not leak into the *weight* comparison: a double
    /// border of equal width still outweighs a single (style number 3 vs 1) and
    /// wins at step 2, before precedence is consulted.
    ///
    /// Pins the two steps apart. Ranking `Single` above `Double` is only correct
    /// as a tie-break; applied one step earlier it would invert every ordinary
    /// single-vs-double edge in a table.
    #[test]
    fn precedence_does_not_override_weight() {
        let single = line(1.0, TableBorderStyle::Single, BLACK);
        let double = line(1.0, TableBorderStyle::Double, BLACK);
        assert!(
            border_weight(&double) > border_weight(&single),
            "equal width, double is heavier"
        );

        for (a, b) in [(single, double), (double, single)] {
            assert_eq!(
                resolve_border_conflict(CellEdge::Line(a), CellEdge::Line(b))
                    .line()
                    .map(|x| x.style),
                Some(TableBorderStyle::Double),
                "the heavier border wins outright, regardless of precedence"
            );
        }
    }

    /// Step 4 — equal weight and style, so the darker colour decides.
    #[test]
    fn equal_weight_and_style_prefers_the_darker_colour() {
        let dark = line(1.0, TableBorderStyle::Single, BLACK);
        let pale = line(1.0, TableBorderStyle::Single, PALE);
        for (a, b) in [(dark, pale), (pale, dark)] {
            assert_eq!(
                resolve_border_conflict(CellEdge::Line(a), CellEdge::Line(b))
                    .line()
                    .map(|x| x.color),
                Some(BLACK),
                "darker colour wins regardless of argument order"
            );
        }
    }

    /// The §17.4.66 darkness keys are compared in order `R+B+2G`, then `B+2G`,
    /// then `G` — so two colours with the same total brightness are separated by
    /// the later keys rather than by argument order.
    #[test]
    fn darkness_tie_breaks_on_the_secondary_keys() {
        // R+B+2G equal (both 255*2 = 510... constructed to match), differing in
        // the B+2G term.
        let a = line(
            1.0,
            TableBorderStyle::Single,
            RgbColor { r: 100, g: 0, b: 0 },
        );
        let b = line(
            1.0,
            TableBorderStyle::Single,
            RgbColor { r: 0, g: 0, b: 100 },
        );
        assert_eq!(
            colour_luminance(&a).0,
            colour_luminance(&b).0,
            "primary key ties"
        );
        // a has B+2G = 0, b has B+2G = 100 → a is "darker" by the second key.
        let winner = resolve_border_conflict(CellEdge::Line(a), CellEdge::Line(b))
            .line()
            .expect("some");
        assert_eq!(winner.color, RgbColor { r: 100, g: 0, b: 0 });
        // And symmetric.
        assert_eq!(
            resolve_border_conflict(CellEdge::Line(b), CellEdge::Line(a))
                .line()
                .map(|x| x.color),
            Some(RgbColor { r: 100, g: 0, b: 0 })
        );
    }

    /// Step 1 — an absent border yields to a present one, in both directions,
    /// and two absent borders stay absent.
    #[test]
    fn absent_yields_to_present() {
        let some = line(1.0, TableBorderStyle::Single, BLACK);
        assert_eq!(
            resolve_border_conflict(CellEdge::Absent, CellEdge::Line(some))
                .line()
                .map(|b| b.width),
            Some(Pt::new(1.0))
        );
        assert_eq!(
            resolve_border_conflict(CellEdge::Line(some), CellEdge::Absent)
                .line()
                .map(|b| b.width),
            Some(Pt::new(1.0))
        );
        assert_eq!(
            resolve_border_conflict(CellEdge::Absent, CellEdge::Absent),
            CellEdge::Absent
        );
    }

    /// **`nil` does not reach across the edge.** [MS-OI29500] §17.4.66 says
    /// *"If the conflicting table cell border is nil, then no border shall be
    /// displayed"*, and read literally that is wrong: `nil` empties its own
    /// cell's edge and leaves the facing cell's border alone. It loses from
    /// either side and at any weight — even a hairline survives it.
    ///
    /// `IP 05 Trenches` is the reference. `<w:bottom w:val="single"/>` above
    /// `<w:top w:val="nil"/>` draws in Word and in macOS's DOCX renderer; so
    /// does an *inherited* bottom faced by a `gridSpan` spacer cell's `nil`, and
    /// it must — a cell paints one border across its whole width, so a wide
    /// cell's `nil` cannot punch a hole in the cell above it.
    #[test]
    fn nil_yields_to_the_facing_border() {
        let hair = line(0.25, TableBorderStyle::Single, BLACK);
        for (a, b) in [
            (CellEdge::Suppressed, CellEdge::Line(hair)),
            (CellEdge::Line(hair), CellEdge::Suppressed),
        ] {
            assert_eq!(
                resolve_border_conflict(a, b).line(),
                Some(hair),
                "the facing border must survive the nil: {a:?} vs {b:?}"
            );
        }
    }

    /// …and yet `nil` is not a no-op, because it declined **inheritance**
    /// upstream in `resolve_cell_effective_borders`. With nothing facing it —
    /// another `nil`, or an edge nobody spoke for — nothing is painted, and the
    /// result stays `Suppressed` rather than collapsing to `Absent`.
    ///
    /// That last part is load-bearing: `emit.rs` may revive an `Absent` top when
    /// a row starts a page slice, and must not revive an emptied one.
    #[test]
    fn nil_stays_suppressed_when_nothing_faces_it() {
        for (a, b) in [
            (CellEdge::Suppressed, CellEdge::Absent),
            (CellEdge::Absent, CellEdge::Suppressed),
            (CellEdge::Suppressed, CellEdge::Suppressed),
        ] {
            assert_eq!(
                resolve_border_conflict(a, b),
                CellEdge::Suppressed,
                "suppression must survive where nothing paints: {a:?} vs {b:?}"
            );
        }
        assert_eq!(
            resolve_border_conflict(CellEdge::Absent, CellEdge::Absent),
            CellEdge::Absent,
            "…but two silent edges stay restorable"
        );
    }

    /// The counterpart, and the half that is easy to get wrong when fixing the
    /// other: an edge declared `none` is **not** suppression. §17.4.66 puts it
    /// with the omitted case — *"If the conflicting table cell border is none
    /// (no border), then the opposing border shall be displayed."*
    ///
    /// `none` never reaches the resolver as its own state; it arrives as
    /// `Absent` because `convert_cell_border_override` maps it to "no override".
    /// This test pins the consequence at the level the resolver sees.
    #[test]
    fn an_absent_edge_never_suppresses() {
        let border = line(1.0, TableBorderStyle::Single, BLACK);
        assert_eq!(
            resolve_border_conflict(CellEdge::Absent, CellEdge::Line(border)),
            CellEdge::Line(border),
            "absent (which is what `none` becomes) must yield, not suppress"
        );
    }

    /// Identical borders resolve to themselves — the reflexive case, which a
    /// comparison built on `partial_cmp` of `f32` could get wrong.
    #[test]
    fn identical_borders_resolve_to_themselves() {
        for b in sample_borders() {
            let r = resolve_border_conflict(CellEdge::Line(b), CellEdge::Line(b))
                .line()
                .expect("some");
            assert_eq!((r.width, r.style, r.color), (b.width, b.style, b.color));
        }
    }
}

/// §17.4.38 edge mapping: which of the six table-level borders each cell edge
/// draws from, given the cell's position in the grid.
#[cfg(test)]
mod edge_mapping_tests {
    use super::*;
    use crate::render::geometry::PtEdgeInsets;
    use crate::render::layout::table::CellVAlign;
    use crate::render::resolve::color::RgbColor;

    /// Every edge gets its own width, so a resolved border names the config
    /// field it came from.
    const TOP: f32 = 1.0;
    const BOTTOM: f32 = 2.0;
    const LEFT: f32 = 3.0;
    const RIGHT: f32 = 4.0;
    const INSIDE_H: f32 = 5.0;
    const INSIDE_V: f32 = 6.0;

    fn edge(width: f32) -> Option<TableBorderLine> {
        Some(TableBorderLine {
            width: Pt::new(width),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        })
    }

    fn config() -> TableBorderConfig {
        TableBorderConfig {
            top: edge(TOP),
            bottom: edge(BOTTOM),
            left: edge(LEFT),
            right: edge(RIGHT),
            inside_h: edge(INSIDE_H),
            inside_v: edge(INSIDE_V),
        }
    }

    fn plain_cell() -> TableCellInput {
        TableCellInput {
            blocks: vec![],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
            cell_borders: None,
            vertical_merge: None,
            vertical_align: CellVAlign::Top,
        }
    }

    /// `(top, bottom, left, right)` widths, so a failure reads as which edges
    /// were mis-mapped rather than as four separate assertions.
    fn edge_widths(
        row_idx: usize,
        grid_col: usize,
        num_rows: usize,
        num_grid_cols: usize,
        spaced: bool,
    ) -> (Option<f32>, Option<f32>, Option<f32>, Option<f32>) {
        let b = resolve_cell_effective_borders(
            &plain_cell(),
            Some(&config()),
            CellPosition {
                row: row_idx,
                num_rows,
                // A full row of single-column cells, which is what these cases
                // model: cell index and grid column coincide, so no gap exists
                // and "first in row" and "at grid column 0" have one answer.
                first_in_row: grid_col == 0,
                last_in_row: grid_col + 1 == num_grid_cols,
            },
            spaced,
        );
        let w = |e: CellEdge| e.line().map(|e| e.width.raw());
        (w(b.top), w(b.bottom), w(b.left), w(b.right))
    }

    fn widths(
        row_idx: usize,
        grid_col: usize,
        num_rows: usize,
        num_grid_cols: usize,
    ) -> (Option<f32>, Option<f32>, Option<f32>, Option<f32>) {
        edge_widths(row_idx, grid_col, num_rows, num_grid_cols, false)
    }

    /// The same corner cell, in a table whose `w:tblCellSpacing` is non-zero.
    fn spaced_edges(
        row_idx: usize,
        grid_col: usize,
        num_rows: usize,
        num_grid_cols: usize,
    ) -> (Option<f32>, Option<f32>, Option<f32>, Option<f32>) {
        edge_widths(row_idx, grid_col, num_rows, num_grid_cols, true)
    }

    /// Issue #168. A spaced cell is inset from the table's boundary, so the
    /// table's own borders must not be painted on it — `emit_table_outline`
    /// draws them at the table's bounds instead, and seeding them here too
    /// would paint each one twice.
    #[test]
    fn a_spaced_cell_takes_no_outer_border_from_the_table() {
        // Top-left corner of a 2x2 table: outer on top and left, interior on
        // bottom and right.
        let (t, b, l, r) = spaced_edges(0, 0, 2, 2);
        assert_eq!(
            t, None,
            "top is the table's own edge and belongs to the outline"
        );
        assert_eq!(
            l, None,
            "left is the table's own edge and belongs to the outline"
        );
        assert_eq!(
            (b, r),
            (Some(INSIDE_H), Some(INSIDE_V)),
            "interior edges are untouched by spacing — see the comment in \
             resolve_cell_effective_borders for why that question is left open"
        );
    }

    /// And the unspaced mapping is exactly as it was, which is what every
    /// document in the corpus depends on.
    #[test]
    fn an_unspaced_cell_still_takes_the_tables_outer_borders() {
        assert_eq!(
            widths(0, 0, 2, 2),
            (Some(TOP), Some(INSIDE_H), Some(LEFT), Some(INSIDE_V))
        );
    }

    /// A 3×3 grid: the corners take the outer borders, the middle takes
    /// `insideH`/`insideV` on all four sides.
    #[test]
    fn outer_edges_use_outer_borders_and_interior_edges_use_inside() {
        assert_eq!(
            widths(0, 0, 3, 3),
            (Some(TOP), Some(INSIDE_H), Some(LEFT), Some(INSIDE_V)),
            "top-left cell"
        );
        assert_eq!(
            widths(1, 1, 3, 3),
            (
                Some(INSIDE_H),
                Some(INSIDE_H),
                Some(INSIDE_V),
                Some(INSIDE_V)
            ),
            "centre cell"
        );
        assert_eq!(
            widths(2, 2, 3, 3),
            (Some(INSIDE_H), Some(BOTTOM), Some(INSIDE_V), Some(RIGHT)),
            "bottom-right cell"
        );
    }

    /// A single-row, single-column table is both first and last on both axes,
    /// so it takes all four outer borders and neither inside border.
    #[test]
    fn a_one_cell_table_takes_all_four_outer_borders() {
        assert_eq!(
            widths(0, 0, 1, 1),
            (Some(TOP), Some(BOTTOM), Some(LEFT), Some(RIGHT))
        );
    }

    /// E5b#7. `num_rows == 0` is unreachable through `layout_table` — it returns
    /// early on empty input, and every other caller is inside a row loop — but
    /// `num_rows` is a free parameter of a `pub(super)` function, so the
    /// last-row test must not depend on a caller having checked it.
    /// `row_idx == num_rows - 1` underflows here; `row_idx + 1 == num_rows`
    /// answers "no row is the last row of an empty table".
    #[test]
    fn an_empty_table_does_not_underflow_the_last_row_check() {
        assert_eq!(
            widths(0, 0, 0, 3),
            (Some(TOP), Some(INSIDE_H), Some(LEFT), Some(INSIDE_V))
        );
    }
}
