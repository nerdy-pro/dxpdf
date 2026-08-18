//! Table measurement phase — cell layout, row heights, cell geometry.
//!
//! Border *resolution* is not here. `borders.rs` answers what each cell's four
//! edges are (§17.4.66, including who owns each shared edge); this phase asks
//! for that answer once, then spends its length on the question it is named
//! for: how tall each row is and where each cell sits.

use crate::render::dimension::Pt;

use crate::render::layout::cell::{layout_cell, CellLayout};

use super::borders::{border_width, plan_table_borders, resolve_table_cell_borders};
use super::grid::{expand_rows_for_vmerge, is_vmerge_continue};
use super::types::{
    CellLayoutEntry, MeasuredRow, MeasuredTable, RowHeightRule, TableBorderConfig, TableRowInput,
    VerticalMergeState,
};

/// Measure all table rows: lay out cell content, compute row heights and cell
/// positions, carrying each cell's resolved borders (from `borders.rs`) along.
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
    // §17.4.45 `tblCellSpacing`, already resolved to points. Zero for every
    // table that does not set it, which is the overwhelming majority.
    cell_spacing: Pt,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: crate::render::layout::paragraph::MeasureTextFn<'_>,
    suppress_first_row_top: bool,
) -> MeasuredTable {
    // §17.4.45: the slots were shrunk by one `cell_spacing` before they got
    // here, so adding it back recovers the table's own outer width — the
    // spacing is carved out of the table, not added to it.
    //
    // §17.4.45 says the spacing sits "between adjacent cells and the edges of
    // the table" without saying whether an edge gets the same amount as the gap
    // between two cells. **Settled** (issue #165) by a Word render of
    // `test-files/issue-165-cellspacing.docx`: the two are equal — one full
    // spacing everywhere, which is what is implemented here and what HTML's
    // `cellspacing` does. The half-gap reading is out.
    //
    // The same render showed the gaps at twice this engine's width, which is a
    // separate and still-open question — see
    // `build::table::resolve_cell_spacing`, which explains why the answer is
    // probably not a factor (ONLYOFFICE's renderer applies none) and probably is
    // that probe declaring `tblCellSpacing` at both table and row level. The
    // equality above is unaffected either way, being a ratio.
    //
    // Note this is *carved out of* the table rather than added to it, so the
    // cells shrink as the spacing grows and the table keeps its declared
    // `w:tblW`. ONLYOFFICE does the same, insetting each cell within its grid
    // slot. `test-files/issue-165-cellspacing-scale.docx` shows both at once.
    let table_width: Pt = col_widths.iter().copied().sum::<Pt>() + cell_spacing;
    let num_rows = rows.len();
    let mut row_heights = Vec::with_capacity(num_rows);

    // Pass 2a: §17.4.66 border resolution — what each cell declares, then who
    // owns each shared edge. Both live in `borders.rs`; this phase only needs
    // the answer.
    let resolved = resolve_table_cell_borders(
        rows,
        col_widths.len(),
        borders,
        cell_spacing,
        suppress_first_row_top,
    );
    let resolved_borders = resolved.cells;

    // §17.4.66: what stands on each line of the grid — the painting answer, kept
    // separate from the charging one above. See `borders.rs`' module doc.
    let plan = plan_table_borders(rows, col_widths.len(), borders, suppress_first_row_top);

    // x of every vertical grid line. §17.4.45's spacing offsets the whole grid
    // by one, matching `cell_x` below; a collapsed table adds zero.
    let grid_x: Vec<Pt> = (0..=col_widths.len())
        .map(|c| col_widths[..c].iter().copied().sum::<Pt>() + cell_spacing)
        .collect();

    // Pass 2b: lay out each cell.
    let mut row_cell_layouts: Vec<Vec<CellLayoutEntry>> = Vec::new();

    for (row_idx, row) in rows.iter().enumerate() {
        let mut entries = Vec::new();
        let mut max_height = Pt::ZERO;
        // §17.4.15: gridBefore — first cell offset.
        let mut grid_idx = row.grid_before as usize;
        // §17.4.38: whether the strip below this row will be reserved for its
        // bottom borders. The condition is repeated verbatim at
        // `border_gap_below` below, and it decides which side of the cell box
        // the bottom border is drawn on: in the strip when there is one, inset
        // into the box's foot when there is not (`emit_cell_borders`). Only the
        // second case takes room from the content.
        let reserves_band_below = row_idx + 1 < num_rows && cell_spacing <= Pt::ZERO;

        for (cell_ci, cell) in row.cells.iter().enumerate() {
            let span = cell.grid_span.max(1) as usize;
            // Defensive clamp: malformed DOCX where gridBefore + spans + gridAfter
            // exceed the grid would otherwise panic in the slice index below.
            // Both ends need clamping — clamping only `grid_end` inverts the
            // range (`start > end`), which panics just as an out-of-bounds end
            // would. Mirrors the same clamp in `build/table.rs`.
            let grid_start = grid_idx.min(col_widths.len());
            let grid_end = (grid_start + span).min(col_widths.len());
            // §17.4.45: the grid slots were already shrunk so they sum to
            // `table_width - cell_spacing`; offsetting every cell by one
            // spacing and taking one off its width then leaves exactly
            // `cell_spacing` between adjacent cells *and* at both table edges,
            // without changing the table's own width. A `gridSpan` cell
            // absorbs the interior gaps it covers, which is what a merged cell
            // should do.
            let slots: Pt = col_widths[grid_start..grid_end].iter().copied().sum();
            let cell_w: Pt = (slots - cell_spacing).max(Pt::ZERO);
            let cell_x: Pt = col_widths[..grid_start].iter().copied().sum::<Pt>() + cell_spacing;

            // §17.4.39/§17.4.66 against §17.4.41/§17.4.42: a border is drawn
            // *inside* the cell box, so the content box starts at
            // `max(border, margin)` from each edge — the margin part is applied
            // by `layout_cell`, and this is the rest. All four sides obey the
            // same rule, and all four must be charged here, because `emit`
            // charges all four when it places the content. Charging only the
            // horizontal pair is what let a row whose top border was thicker
            // than its top cell margin overflow its own box by the difference,
            // straight into the strip where the bottom border paints.
            //
            // **How much of a border is "inside" depends on whether the edge is
            // shared**, and it is the same rule `rasterize_border_grid` paints
            // by: a collapsed border straddles a line two cells share, so each
            // is charged half, and it goes wholly inside a line shared with
            // nothing, so the table's own two edges are charged in full.
            //
            // **Measured**, against `test-files/border-content-charge.docx`: with
            // zero cell margins and a shared border stepping 0.5 → 12pt, Word
            // draws the following cell's glyph flush against the border's inner
            // edge at every weight. Both other readings are refuted by that one
            // render — charging it nothing puts the glyph on the grid line with
            // the border painted through it, and charging it the whole width
            // leaves a gap of half the border. `tests/table_cell_content_box.rs`
            // holds the assertion.
            //
            // The charge comes from the **plan** and not from `resolved_borders`
            // for the same reason: resolution hands a shared edge to one of the
            // two cells and clears the other, so the loser's own `left` says
            // `Absent` and would be charged nothing. Both cells stand on one
            // line and both must see it.
            let b = &resolved_borders[row_idx][cell_ci];
            let charge = |edge, outer: bool, margin: Pt| {
                let w = border_width(edge);
                ((if outer { w } else { w * 0.5 }) - margin).max(Pt::ZERO)
            };
            // §17.4.45: a spaced table has no shared edges at all — each cell
            // keeps its four borders wholly inside itself, so each is charged in
            // full, and the plan (which resolves as though collapsed) does not
            // describe it.
            let (extra_left, extra_right) = if cell_spacing > Pt::ZERO {
                (
                    (border_width(b.left) - cell.margins.left).max(Pt::ZERO),
                    (border_width(b.right) - cell.margins.right).max(Pt::ZERO),
                )
            } else {
                (
                    charge(
                        plan.vertical(grid_start, row_idx),
                        grid_start == 0,
                        cell.margins.left,
                    ),
                    charge(
                        plan.vertical(grid_end, row_idx),
                        grid_end == col_widths.len(),
                        cell.margins.right,
                    ),
                )
            };
            // The horizontal twin needs no such split, and the asymmetry is
            // worth stating rather than leaving to be rediscovered. §17.4.38
            // reserves a strip *between* two rows' content boxes wide enough for
            // the border on their shared boundary (`border_gap_below`), so an
            // interior horizontal already takes its room from neither box —
            // which is the half-each rule delivered by geometry instead of by a
            // charge. `extra_top` therefore only ever sees the table's own top
            // edge, where the whole border is inside and the full charge is
            // right. The one shape that escapes this is a boundary the *lower*
            // row owns (`can_own` prefers the upper, and only a `gridSpan`
            // mismatch overrides it): there the strip is zero and the lower row
            // is charged the full width. No render has measured that case.
            let extra_top = (border_width(b.top) - cell.margins.top).max(Pt::ZERO);
            let extra_bottom = if reserves_band_below {
                Pt::ZERO
            } else {
                (border_width(b.bottom) - cell.margins.bottom).max(Pt::ZERO)
            };
            let layout_w = (cell_w - extra_left - extra_right).max(Pt::ZERO);

            let is_continue = cell.vertical_merge == Some(VerticalMergeState::Continue);
            let layout = if is_continue {
                CellLayout {
                    commands: Vec::new(),
                    content_height: Pt::ZERO,
                    lines: Vec::new(),
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

            // §17.4.84: a merged cell's height is normally decided by
            // `expand_rows_for_vmerge` over the whole span, not here — folding a
            // `Restart` cell's full content into its *first* row would double-count
            // it against the rows below.
            //
            // Unless the span is a span of one. A `Restart` with no `Continue`
            // under it is an ordinary cell, and `expand_rows_for_vmerge` skips it
            // (it returns early when the group is a single row), so if this branch
            // skipped it too the row would get **no** height from any path while
            // still emitting its content — following blocks then draw on top of
            // the table. Word treats a restart with nothing continuing as a plain
            // cell, which is what this reproduces.
            let continues_below =
                row_idx + 1 < num_rows && is_vmerge_continue(&rows[row_idx + 1], grid_idx);
            let is_lone_restart =
                cell.vertical_merge == Some(VerticalMergeState::Restart) && !continues_below;
            if cell.vertical_merge.is_none() || is_lone_restart {
                max_height = max_height.max(
                    layout.content_height + cell.margins.vertical() + extra_top + extra_bottom,
                );
            }

            entries.push(CellLayoutEntry {
                content_dx: extra_left,
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

        // §17.4.45: the row's box reserves its own leading gap, mirroring the
        // horizontal inset above — `emit_one_row` places content one spacing
        // below the cursor, so consecutive rows end up exactly `cell_spacing`
        // apart. `RowHeightRule` is applied to the *content* height first, so a
        // `trHeight` still means the height of the row's content, not of the
        // content plus a gap the author never asked for.
        row_heights.push(max_height + cell_spacing);
        row_cell_layouts.push(entries);
    }

    // §17.4.84: distribute vMerge overflow.
    expand_rows_for_vmerge(rows, &row_cell_layouts, &mut row_heights);

    // Compute border gaps and assemble measured rows.
    let measured_rows: Vec<MeasuredRow> = row_cell_layouts
        .into_iter()
        .zip(resolved_borders)
        .zip(row_heights.iter())
        .enumerate()
        .map(|(row_idx, ((entries, borders), &height))| {
            // With cell spacing there is no shared edge to reserve room for:
            // every cell draws its own bottom border inside its own box, and
            // the gap between rows is the spacing itself.
            let border_gap_below = if row_idx + 1 < num_rows && cell_spacing <= Pt::ZERO {
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
                leading_gap: cell_spacing,
                border_gap_below,
                plan_row: row_idx,
            }
        })
        .collect();

    MeasuredTable {
        rows: measured_rows,
        table_width,
        plan,
        grid_x,
    }
}

#[cfg(test)]
mod tests {
    use super::super::borders::CellEdge;
    use super::super::types::{
        CellBorderConfig, CellBorderOverride, CellVAlign, RowHeightRule, TableBorderConfig,
        TableBorderLine, TableBorderStyle, TableCellInput, TableRowInput, VerticalMergeState,
    };
    use super::measure_table_rows;
    use crate::render::dimension::Pt;
    use crate::render::geometry::PtEdgeInsets;
    use crate::render::resolve::color::RgbColor;

    fn single(w: f32) -> TableBorderLine {
        TableBorderLine {
            width: Pt::new(w),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        }
    }

    /// A table style like `Tabellenraster`: every side plus insideH/insideV.
    fn all_single() -> TableBorderConfig {
        let s = single(0.5);
        TableBorderConfig {
            top: Some(s),
            bottom: Some(s),
            left: Some(s),
            right: Some(s),
            inside_h: Some(s),
            inside_v: Some(s),
        }
    }

    fn cb(top: Option<CellBorderOverride>, bottom: Option<CellBorderOverride>) -> CellBorderConfig {
        CellBorderConfig {
            top,
            bottom,
            left: None,
            right: None,
        }
    }

    fn cell(span: u32, borders: Option<CellBorderConfig>) -> TableCellInput {
        TableCellInput {
            blocks: vec![],
            margins: PtEdgeInsets::ZERO,
            grid_span: span,
            shading: None,
            cell_borders: borders,
            vertical_merge: None,
            vertical_align: CellVAlign::Top,
        }
    }

    fn row(cells: Vec<TableCellInput>) -> TableRowInput {
        TableRowInput {
            cells,
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
        }
    }

    /// §17.4.80: `row`, plus a `trHeight` constraint.
    fn row_sized(cells: Vec<TableCellInput>, rule: RowHeightRule) -> TableRowInput {
        TableRowInput {
            height_rule: Some(rule),
            ..row(cells)
        }
    }

    /// §17.4.15: `row`, starting `grid_before` grid columns in.
    fn row_at(cells: Vec<TableCellInput>, grid_before: u32) -> TableRowInput {
        TableRowInput {
            grid_before,
            ..row(cells)
        }
    }

    /// One 14 pt line per `words` entry that does not fit beside its
    /// predecessor. `default_line_height` is 10 pt everywhere below, but the
    /// fragment's own 10 pt ascent + 4 pt descent wins, so each emitted line is
    /// exactly 14 pt tall.
    fn text_block(words: &[(&str, f32)]) -> crate::render::layout::section::LayoutBlock {
        use crate::render::fonts::Toggle;
        use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
        use std::rc::Rc;

        let font = Rc::new(FontProps {
            rtl: Toggle::Absent,
            family: Rc::from("Test"),
            size: Pt::new(12.0),
            bold: Toggle::Absent,
            italic: Toggle::Absent,
            underline: false,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        });
        crate::render::layout::section::LayoutBlock::Paragraph {
            fragments: words
                .iter()
                .map(|&(text, width)| Fragment::Text {
                    shaped: None,
                    level: crate::i18n::bidi::BidiLevel::LTR,
                    text: text.into(),
                    break_after: crate::render::layout::fragment::fixture_break_after(text),
                    font: font.clone(),
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
                })
                .collect(),
            style: crate::render::layout::paragraph::ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    /// A one-column cell holding `lines` lines of text — each word is 30 pt
    /// wide and breakable, so a 30 pt content width puts one per line.
    fn cell_of(lines: usize) -> TableCellInput {
        TableCellInput {
            blocks: vec![text_block(
                &(0..lines).map(|_| ("aa ", 30.0)).collect::<Vec<_>>(),
            )],
            ..cell(1, None)
        }
    }

    /// [MS-OI29500] §17.4.66 regression: a `gridSpan` upper cell facing several
    /// lower cells must not drop the later cells' top borders (previously only
    /// the first lower cell was resolved and the rest nulled), and the whole
    /// shared edge must be drawn from a single side so the line does not split
    /// across two y positions. Mirrors the real doc's
    /// `[spacer | Function: | Qualitätssicherung]` row under a `gridSpan` header.
    ///
    /// Non-uniformity comes from a **heavier border on one column**, not from a
    /// `nil`. It used to come from a nil, which worked only while nil resolved
    /// to "absent": now that nil suppresses (§17.4.66), a nil under the wide
    /// cell makes its run uniformly suppressed and the upper row owns the edge —
    /// the opposite branch, so the configuration no longer reaches the bug this
    /// test exists for. Suppression is covered separately by
    /// `nil_suppresses_across_a_gridspan_mismatch`.
    #[test]
    fn wide_upper_cell_draws_whole_edge_from_lower_row() {
        let s = single(0.5);
        let heavy = single(2.0);
        let rows = vec![
            // Row 0: gridSpan=2 header over spacer+Function, then two single
            // cells over the Qualitätssicherung span. All bottoms inherit.
            row(vec![cell(2, None), cell(1, None), cell(1, None)]),
            // Row 1: [spacer | Function (heavy top) | Q (gridSpan=2)]. The heavy
            // top on column 1 alone makes the wide upper cell's run non-uniform.
            row(vec![
                cell(1, None),
                cell(1, Some(cb(Some(CellBorderOverride::Border(heavy)), None))),
                cell(2, None),
            ]),
        ];
        let cols = vec![Pt::new(100.0); 4];
        let m = measure_table_rows(
            &rows,
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        // Whole edge drawn from the lower row → every upper bottom cleared,
        // so Function and Qualitätssicherung tops share one y position.
        for b in &m.rows[0].borders {
            assert_eq!(
                b.bottom,
                CellEdge::Absent,
                "upper bottoms cleared (edge owned below)"
            );
        }
        assert_eq!(
            m.rows[1].borders[0].top.line(),
            Some(s),
            "spacer column keeps the inherited insideH"
        );
        assert_eq!(
            m.rows[1].borders[1].top.line(),
            Some(heavy),
            "Function keeps its heavier top border across the gridSpan mismatch"
        );
        assert_eq!(
            m.rows[1].borders[2].top.line(),
            Some(s),
            "Qualitätssicherung top drawn from the same (lower) side as Function"
        );
    }

    /// §17.4.66 step 0 across a `gridSpan` mismatch: a `nil` bottom on a wide
    /// upper cell suppresses the columns it spans that fall back to the table —
    /// but **not** the column where the lower cell has a border of its own,
    /// which survives and is drawn from below. One nil bottom therefore resolves to two different
    /// answers along its own width, which is exactly the case that forces the
    /// per-column resolution this pass does.
    ///
    /// This is the configuration `wide_upper_cell_draws_whole_edge_from_lower_row`
    /// used to carry, kept here for what it now demonstrates.
    #[test]
    fn nil_suppresses_across_a_gridspan_mismatch() {
        let s = single(0.5);
        let rows = vec![
            row(vec![
                cell(2, Some(cb(None, Some(CellBorderOverride::Suppress)))),
                cell(1, None),
                cell(1, None),
            ]),
            row(vec![
                cell(1, Some(cb(Some(CellBorderOverride::Suppress), None))),
                cell(1, Some(cb(Some(CellBorderOverride::Border(s)), None))),
                cell(2, None),
            ]),
        ];
        let cols = vec![Pt::new(100.0); 4];
        let m = measure_table_rows(
            &rows,
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        // The nil bottom spans columns 0-1 and resolves differently on each:
        // column 0 faces another nil and stays suppressed, column 1 faces
        // Function's *declared* top and loses to it. A cell paints one border
        // across its width, so that split hands the whole edge to the lower row.
        for b in &m.rows[0].borders {
            assert_eq!(
                b.bottom,
                CellEdge::Absent,
                "upper bottoms cleared — the nil span is not uniform, so the edge is owned below"
            );
        }
        assert_eq!(
            m.rows[1].borders[0].top,
            CellEdge::Suppressed,
            "column 0: nil against nil stays suppressed"
        );
        assert_eq!(
            m.rows[1].borders[1].top.line(),
            Some(s),
            "column 1: Function has a top of its own, so the nil above does not erase it"
        );
        // Columns outside the nil span keep the inherited insideH, drawn from
        // the same (lower) side so the line sits at one y.
        assert_eq!(m.rows[1].borders[2].top.line(), Some(s));
    }

    /// A cell can paint one border across its whole span when every column under
    /// it *paints the same thing* — not when every column resolved to the same
    /// `CellEdge`. `Absent` and `Suppressed` both paint nothing, so a run
    /// mixing them is uniform; comparing with `==` would split it and hand the
    /// edge to the other row for no visible reason.
    ///
    /// Built so the two columns differ **only** in that: with no `insideH` to
    /// inherit, column 0 resolves to `Suppressed` (the lower cell wrote `nil`)
    /// and column 1 to `Absent` (nobody said anything). Neither paints.
    ///
    /// The upper cell must end up owning the edge, and carrying `Suppressed`
    /// while doing so — `emit.rs` restores an `Absent` top at a page split and
    /// must not restore an emptied one, so which of the two lands there is a
    /// real distinction, not bookkeeping.
    #[test]
    fn a_uniform_run_is_not_split_by_absent_versus_suppressed() {
        let hair = single(0.5);
        let borders = TableBorderConfig {
            top: Some(hair),
            bottom: Some(hair),
            left: Some(hair),
            right: Some(hair),
            // Nothing to inherit on the inter-row edge — so the only states in
            // play there are `Absent` and `Suppressed`.
            inside_h: None,
            inside_v: Some(hair),
        };
        let rows = vec![
            // One wide cell saying nothing about its bottom.
            row(vec![cell(2, None)]),
            // Column 0 writes `nil`; column 1 says nothing.
            row(vec![
                cell(1, Some(cb(Some(CellBorderOverride::Suppress), None))),
                cell(1, None),
            ]),
        ];
        let cols = vec![Pt::new(100.0), Pt::new(100.0)];
        let m = measure_table_rows(
            &rows,
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&borders),
            None,
            false,
        );

        assert_eq!(
            m.rows[0].borders[0].bottom,
            CellEdge::Suppressed,
            "the wide cell owns the whole edge, carrying the suppression"
        );
        assert_eq!(
            m.rows[1].borders[0].top,
            CellEdge::Absent,
            "so the lower row's tops are cleared"
        );
        assert_eq!(m.rows[1].borders[1].top, CellEdge::Absent);
    }

    /// §17.4.66: a `nil` among the cells on one side of an edge **cannot** punch
    /// a gap through a wide `gridSpan` cell facing it. The wide cell inherits
    /// `insideH` for its whole width and paints one border across it; a
    /// neighbour's `nil` empties only that neighbour's own edge.
    ///
    /// This is `IP 05 Trenches`' `Date/Time:` cell with the sides swapped — the
    /// real document has the wide spacer cell *below*, its `nil` aimed at the
    /// narrow spacer column, and the label cell above still draws the bottom it
    /// inherited. The assertion here was inverted twice: once when `nil`
    /// wrongly collapsed to "absent" (so it also wrongly *inherited*), and once
    /// when `nil` was made to win the conflict outright. Declining inheritance
    /// and overruling the neighbour are different powers; `nil` has only the
    /// first.
    #[test]
    fn nil_spacer_cannot_punch_a_gap_through_a_wide_facing_cell() {
        let s = single(0.5);
        let rows = vec![
            // Row 0: [inherits single | nil spacer | inherits single].
            row(vec![
                cell(1, None),
                cell(1, Some(cb(None, Some(CellBorderOverride::Suppress)))),
                cell(1, None),
            ]),
            // Row 1: one gridSpan=3 cell inheriting insideH as its top.
            row(vec![cell(3, None)]),
        ];
        let cols = vec![Pt::new(100.0), Pt::new(100.0), Pt::new(100.0)];
        let m = measure_table_rows(
            &rows,
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        // Every column resolves to the same line, so the edge is uniform and one
        // side paints it whole. The nil column is not a hole in it.
        assert_eq!(
            m.rows[0].borders[1].bottom.line(),
            Some(s),
            "the wide cell below still supplies this column's border"
        );
        assert_eq!(m.rows[0].borders[0].bottom.line(), Some(s));
        assert_eq!(m.rows[0].borders[2].bottom.line(), Some(s));
        // …and it is drawn exactly once, from the upper row.
        assert_eq!(m.rows[1].borders[0].top.line(), None);
    }

    /// [MS-OI29500] §17.4.66 regression: an upper `gridSpan` cell that leaves the last
    /// column uncovered (its gridAfter gap) must not "own" the edge, or a
    /// lower cell straddling that boundary would draw its own top over the
    /// upper bottom → a doubled line. Mirrors the real doc's `gridSpan=9`
    /// section row above the `Observations` (`gridSpan=2`) header.
    #[test]
    fn upper_grid_after_gap_yields_edge_to_lower_row() {
        let s = single(0.5);
        let rows = vec![
            // Row 0: one gridSpan=2 cell over cols 0-1; col 2 is its gridAfter.
            row(vec![cell(2, None)]),
            // Row 1: [cell | gridSpan=2 cell straddling covered col 1 + col 2].
            row(vec![cell(1, None), cell(2, None)]),
        ];
        let cols = vec![Pt::new(100.0); 3];
        let m = measure_table_rows(
            &rows,
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        // Upper can't cover col 2, so the lower row owns the whole edge:
        // its bottom is cleared (no doubling), the lower tops carry the line.
        assert_eq!(
            m.rows[0].borders[0].bottom.line(),
            None,
            "upper bottom cleared so the straddling lower cell isn't doubled"
        );
        assert_eq!(m.rows[1].borders[0].top.line(), Some(s));
        assert_eq!(m.rows[1].borders[1].top.line(), Some(s));
    }

    /// Aligned grids keep the pre-existing "upper cell owns the shared edge"
    /// behaviour: the lower cell's top is cleared, the upper bottom carries it.
    #[test]
    fn aligned_grid_upper_cell_owns_horizontal_edge() {
        let s = single(0.5);
        let rows = vec![
            row(vec![cell(1, None), cell(1, None)]),
            row(vec![cell(1, None), cell(1, None)]),
        ];
        let cols = vec![Pt::new(100.0), Pt::new(100.0)];
        let m = measure_table_rows(
            &rows,
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );
        for ci in 0..2 {
            assert_eq!(m.rows[0].borders[ci].bottom.line(), Some(s));
            assert_eq!(m.rows[1].borders[ci].top.line(), None);
        }
    }

    /// §17.4.45 geometry. The invariant that matters is a *uniform* gap: exactly
    /// one `cell_spacing` between adjacent cells **and** at both table edges,
    /// with the table's own width unchanged.
    ///
    /// Slots are pre-shrunk by `reserve_cell_spacing` (build side), so this
    /// feeds slots summing to `width - spacing` and checks the resulting edges.
    #[test]
    fn cell_spacing_leaves_a_uniform_gap_and_keeps_the_table_width() {
        let spacing = Pt::new(10.0);
        // Table is 100pt wide; slots therefore sum to 90.
        let slots = vec![Pt::new(45.0), Pt::new(45.0)];
        let rows = vec![row(vec![cell(1, None), cell(1, None)])];
        let m = measure_table_rows(
            &rows,
            &slots,
            spacing,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        assert_eq!(
            m.table_width,
            Pt::new(100.0),
            "spacing comes out of the table"
        );

        let e = &m.rows[0].entries;
        let left_edge = e[0].cell_x;
        let gap_between = e[1].cell_x - (e[0].cell_x + e[0].cell_w);
        let right_edge = m.table_width - (e[1].cell_x + e[1].cell_w);
        assert_eq!(left_edge, spacing, "gap at the table's left edge");
        assert_eq!(gap_between, spacing, "gap between adjacent cells");
        assert_eq!(right_edge, spacing, "gap at the table's right edge");
    }

    /// A `gridSpan` cell absorbs the interior gaps it covers — it is one cell,
    /// so the spacing between the columns it spans belongs to it.
    #[test]
    fn a_gridspan_cell_absorbs_the_gaps_it_covers() {
        let spacing = Pt::new(10.0);
        let slots = vec![Pt::new(30.0); 3]; // 90 + 10 spacing = 100 wide
        let rows = vec![
            row(vec![cell(3, None)]),
            row(vec![cell(1, None), cell(1, None), cell(1, None)]),
        ];
        let m = measure_table_rows(
            &rows,
            &slots,
            spacing,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        let wide = &m.rows[0].entries[0];
        let narrow = &m.rows[1].entries;
        assert_eq!(wide.cell_x, spacing);
        assert_eq!(wide.cell_w, Pt::new(80.0), "spans 90 of slot minus one gap");
        // The wide cell covers both interior gaps plus the two narrow cells
        // between them, so its right edge matches the last narrow cell's.
        assert_eq!(
            wide.cell_x + wide.cell_w,
            narrow[2].cell_x + narrow[2].cell_w
        );
    }

    /// Vertically the rule is the same: one spacing between row boxes, and one
    /// above the first row. The row's own height carries its leading gap.
    #[test]
    fn cell_spacing_separates_rows_vertically() {
        let spacing = Pt::new(6.0);
        let slots = vec![Pt::new(94.0)];
        let rows = vec![row(vec![cell(1, None)]), row(vec![cell(1, None)])];
        let m = measure_table_rows(
            &rows,
            &slots,
            spacing,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        for r in &m.rows {
            assert_eq!(r.leading_gap, spacing);
            assert_eq!(
                r.border_gap_below,
                Pt::ZERO,
                "no shared edge to reserve for once cells are separated"
            );
        }
        // These cells are empty, so each row's height is its reserved gap —
        // added exactly once per row, not once per cell — plus the one
        // horizontal border that lies inside its box. A spaced table's own top
        // and bottom belong to the outline, so row 0 has only its `insideH`
        // bottom and row 1 only its `insideH` top: one 0.5pt border each, not
        // two and not none.
        for r in &m.rows {
            assert_eq!(r.height, spacing + Pt::new(0.5));
        }
    }

    /// [MS-OI29500] §17.4.66: *"If the cell spacing is nonzero ... then all cell
    /// borders and outer table borders display."* Separated cells share no edge,
    /// so conflict resolution is skipped and **both** sides keep their border —
    /// where collapsing would have kept one and cleared the other.
    #[test]
    fn nonzero_cell_spacing_disables_border_collapsing() {
        let slots = vec![Pt::new(45.0), Pt::new(45.0)];
        let rows = vec![
            row(vec![cell(1, None), cell(1, None)]),
            row(vec![cell(1, None), cell(1, None)]),
        ];
        let measure = |spacing: Pt| {
            measure_table_rows(
                &rows,
                &slots,
                spacing,
                Pt::new(10.0),
                Some(&all_single()),
                None,
                false,
            )
        };

        // Collapsed (no spacing): the right/left pair resolves to one border on
        // the left cell, and the shared horizontal edge is owned by one row.
        let collapsed = measure(Pt::ZERO);
        assert!(collapsed.rows[0].borders[1].left.line().is_none());
        assert!(collapsed.rows[1].borders[0].top.line().is_none());

        // Spaced: no edge is resolved away. This is what the test is about —
        // where collapsing kept one side of a shared edge and cleared the
        // other, both sides now survive.
        let spaced = measure(Pt::new(8.0));
        assert!(spaced.rows[0].borders[1].left.line().is_some(), "left kept");
        assert!(
            spaced.rows[0].borders[0].right.line().is_some(),
            "right kept"
        );
        assert!(spaced.rows[1].borders[0].top.line().is_some(), "top kept");
        assert!(
            spaced.rows[0].borders[0].bottom.line().is_some(),
            "bottom kept"
        );

        // Issue #168: the *outer* edges of a spaced table are no longer drawn
        // on the cells. They are the table's own border, and a cell inset by
        // the spacing is the wrong place for it — `emit_table_outline` paints
        // it at the table's bounds instead. This assertion used to read "every
        // cell keeps all four of its own borders", which was true of the
        // implementation but said more than the §17.4.66 sentence above does:
        // the sentence distinguishes *cell* borders from *outer table* borders
        // precisely so both can display, each in its own place.
        assert!(spaced.rows[0].borders[0].top.line().is_none());
        assert!(spaced.rows[0].borders[0].left.line().is_none());
        assert!(spaced.rows[1].borders[1].bottom.line().is_none());
        assert!(spaced.rows[1].borders[1].right.line().is_none());
    }

    // ── §17.4.80 row height rules ────────────────────────────────────────
    //
    // `RowHeightRule` reached layout with no layout test of its own: `mod.rs`
    // pins `AtLeast` in the one direction where it wins, and `Exact` had only a
    // parse test. Both arms of the `match` below are decided here, together
    // with the two orderings around them — the rule runs *before* the cell
    // spacing gap is added and *before* `expand_rows_for_vmerge`.

    /// A one-column table 50 pt wide, so `cell_of(n)` is exactly `n` lines.
    fn sized_table(rows: &[TableRowInput]) -> super::MeasuredTable {
        measure_table_rows(
            rows,
            &[Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(10.0),
            None,
            None,
            false,
        )
    }

    /// The measurement the rules are compared against: three 14 pt lines.
    #[test]
    fn an_unconstrained_row_is_as_tall_as_its_content() {
        let m = sized_table(&[row(vec![cell_of(3)])]);
        assert_eq!(m.rows[0].height, Pt::new(42.0), "three 14 pt lines");
    }

    /// §17.4.80 `hRule="exact"`: *"the height of the row shall be exactly the
    /// value specified"*. Shorter content does not shrink it.
    #[test]
    fn an_exact_row_height_holds_when_the_content_is_shorter() {
        let m = sized_table(&[row_sized(
            vec![cell_of(1)],
            RowHeightRule::Exact(Pt::new(40.0)),
        )]);
        assert_eq!(
            m.rows[0].height,
            Pt::new(40.0),
            "14 pt of content, 40 pt row"
        );
    }

    /// …and taller content does not grow it, which is the half `atLeast` does
    /// not share: `exact` is a ceiling as well as a floor. All three rows below
    /// hold more than the 20 pt they declare — 42, 112 and 280 pt of it, the
    /// first calibrated by `an_unconstrained_row_is_as_tall_as_its_content` —
    /// so the equality cannot be met by any rule that reads the content.
    ///
    /// What this deliberately does **not** assert is where the overflowing
    /// content goes. dxpdf does not clip it: the lines that do not fit are
    /// painted below the row's box and overprint whatever follows the table.
    /// That is a known open defect rather than a decision — §17.4.80 states the
    /// height and says nothing about the overflow, LibreOffice clips, and no
    /// Word render is on record — so pinning the overflow here would pin a
    /// guess. The height, which the section does state, is pinned.
    #[test]
    fn an_exact_row_height_holds_when_the_content_is_taller() {
        let heights: Vec<Pt> = [3usize, 8, 20]
            .iter()
            .map(|&n| {
                sized_table(&[row_sized(
                    vec![cell_of(n)],
                    RowHeightRule::Exact(Pt::new(20.0)),
                )])
                .rows[0]
                    .height
            })
            .collect();
        assert_eq!(
            heights,
            vec![Pt::new(20.0); 3],
            "an exact row is its declared height whatever it holds"
        );
    }

    /// §17.4.80 `hRule="atLeast"` is a floor and only a floor — the same two
    /// contents that leave an `exact` row unmoved decide an `atLeast` one.
    #[test]
    fn an_at_least_row_height_is_a_floor_that_taller_content_overrides() {
        let at_least = |n| {
            sized_table(&[row_sized(
                vec![cell_of(n)],
                RowHeightRule::AtLeast(Pt::new(20.0)),
            )])
            .rows[0]
                .height
        };
        assert_eq!(at_least(1), Pt::new(20.0), "14 pt of content raised to 20");
        assert_eq!(
            at_least(3),
            Pt::new(42.0),
            "42 pt of content is not cut to 20"
        );
    }

    /// §17.4.80 × §17.4.45: the height rule sizes the row's **content**, and
    /// the leading gap is reserved on top of it. So a row does not lose the
    /// height its author declared to a gap the author did not — its content box
    /// is the declared height at every spacing, and only the row's outer box
    /// grows.
    ///
    /// Stated as the difference between two spacings rather than as one number,
    /// because that is the invariant: `height - leading_gap` is the rule's own
    /// answer, unchanged.
    #[test]
    fn an_exact_row_height_is_measured_before_the_cell_spacing_gap() {
        let spacing = Pt::new(6.0);
        let rows = [row_sized(
            vec![cell_of(1)],
            RowHeightRule::Exact(Pt::new(40.0)),
        )];
        // The slot is pre-shrunk by `reserve_cell_spacing` on the build side.
        let spaced = measure_table_rows(
            &rows,
            &[Pt::new(44.0)],
            spacing,
            Pt::new(10.0),
            None,
            None,
            false,
        );
        assert_eq!(spaced.rows[0].leading_gap, spacing);
        assert_eq!(
            spaced.rows[0].height - spaced.rows[0].leading_gap,
            Pt::new(40.0),
            "the declared height is the content box, not the outer box"
        );
        assert_eq!(spaced.rows[0].height, Pt::new(46.0));
    }

    /// §17.4.80 × §17.4.84: the height rule runs **before**
    /// `expand_rows_for_vmerge`, and that pass puts a span's shortfall on the
    /// span's *last* row (issue #165). Together those two facts decide the case
    /// neither settles alone: an `exact` restart row keeps exactly its declared
    /// height, and every point the merged content still needs lands below it.
    ///
    /// The span totals the restart cell's content height exactly, which is what
    /// makes this more than a restatement of `expand_puts_overflow_on_the_last_
    /// row_of_the_span` — that test starts from two rows of natural height, and
    /// here the first row's height is the rule's answer rather than its
    /// content's.
    #[test]
    fn an_exact_height_on_a_vmerge_restart_row_is_kept_and_the_span_grows_below() {
        let restart = TableCellInput {
            vertical_merge: Some(VerticalMergeState::Restart),
            ..cell_of(3)
        };
        let continued = TableCellInput {
            vertical_merge: Some(VerticalMergeState::Continue),
            ..cell(1, None)
        };
        let m = sized_table(&[
            row_sized(vec![restart], RowHeightRule::Exact(Pt::new(10.0))),
            row(vec![continued]),
        ]);

        assert_eq!(
            m.rows[0].height,
            Pt::new(10.0),
            "the restart row keeps the height it declared"
        );
        assert_eq!(
            m.rows[1].height,
            Pt::new(32.0),
            "the last row of the span absorbs the remaining 42 − 10"
        );
        assert_eq!(
            m.rows[0].height + m.rows[1].height,
            Pt::new(42.0),
            "and the span totals the merged cell's content exactly"
        );
    }

    // ── §17.4.15 × §17.4.45 ──────────────────────────────────────────────

    /// `gridBefore` and `tblCellSpacing` both displace a cell rightward, and no
    /// test crossed them: every `gridBefore` case ran at zero spacing and every
    /// spacing case at zero `gridBefore`.
    ///
    /// They compose **by construction**, which is what is asserted: the columns
    /// a `gridBefore` skips are grid columns like any other, so a cell in
    /// column 1 must land exactly where column 1's cell lands in a row that
    /// declares no `gridBefore`. A formula that dropped either term — or applied
    /// the spacing twice for a displaced cell — fails that parity.
    #[test]
    fn grid_before_and_cell_spacing_place_a_cell_in_the_same_column() {
        let spacing = Pt::new(10.0);
        // Three 30 pt slots plus one spacing is a 100 pt table.
        let slots = vec![Pt::new(30.0); 3];
        let rows = vec![
            row(vec![cell(1, None), cell(1, None), cell(1, None)]),
            row_at(vec![cell(1, None), cell(1, None)], 1),
        ];
        let m = measure_table_rows(
            &rows,
            &slots,
            spacing,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );

        let full = &m.rows[0].entries;
        let offset = &m.rows[1].entries;
        for (i, e) in offset.iter().enumerate() {
            assert_eq!(
                (e.cell_x, e.cell_w),
                (full[i + 1].cell_x, full[i + 1].cell_w),
                "the displaced row's cell {i} must sit in grid column {}",
                i + 1
            );
        }
        // Absolute, so the parity above cannot be satisfied by moving both rows
        // together: one whole slot in, then one spacing, with one gap taken out
        // of the slot's own width.
        assert_eq!(offset[0].cell_x, Pt::new(40.0), "30 pt of slot + one gap");
        assert_eq!(
            offset[0].cell_w,
            Pt::new(20.0),
            "the 30 pt slot less one gap"
        );
        // The skipped column is still part of the table, so the gap between the
        // table's left edge and the first *drawn* cell is one column wider than
        // in the undisplaced row — and the table's own width is untouched.
        assert_eq!(m.table_width, Pt::new(100.0));
    }

    // ── §17.4.21 / §17.4.80: a row with no cells ─────────────────────────

    /// `<w:tr/>` with no `<w:tc>` is well-formed. §17.4.21 says a row's height
    /// is determined by the glyphs in its cells, so a row with none has none —
    /// and having none it must displace nothing *sideways*: its neighbours' cells
    /// land at exactly the x they land at in a table that never contained it.
    ///
    /// The populated rows carry a line of text so that "zero" is a measurement
    /// rather than the default every empty cell already produces.
    ///
    /// **The row below it is half a point taller, and that is not this row's
    /// height leaking.** §17.4.66 resolves a shared horizontal edge to one of
    /// the two cells facing across it, and an empty row leaves the row beneath
    /// facing nothing — so where the control's second row has its top resolved
    /// to `Absent` (the row above owns that edge and paints it in the strip
    /// between them), the gapped table's third row keeps a top border of its own
    /// and paints it *inside* its box. `measure_table_rows` charges a border
    /// inside the box to the height, so the height follows the ownership.
    ///
    /// That ownership difference predates the height following it, and it is
    /// visible without any of this: the gapped table paints 1pt of border at
    /// that boundary where the control paints 0.5pt. Whether an empty row should
    /// break the shared edge at all is the real question, and it is a
    /// border-resolution one — asserted here as it stands so that changing it is
    /// a deliberate act.
    #[test]
    fn a_row_with_no_cells_is_zero_height_and_moves_nothing() {
        let slots = vec![Pt::new(50.0), Pt::new(50.0)];
        let populated = || row(vec![cell_of(1), cell_of(1)]);
        let measure = |rows: &[TableRowInput]| {
            measure_table_rows(
                rows,
                &slots,
                Pt::ZERO,
                Pt::new(10.0),
                Some(&all_single()),
                None,
                false,
            )
        };
        let gapped = measure(&[populated(), row(vec![]), populated()]);
        let control = measure(&[populated(), populated()]);

        assert!(gapped.rows[1].entries.is_empty(), "no cells, no entries");
        assert_eq!(gapped.rows[1].height, Pt::ZERO, "no cells, no height");
        assert_eq!(gapped.table_width, control.table_width);

        for (mine, theirs) in [(0usize, 0usize), (2, 1)] {
            let got: Vec<_> = gapped.rows[mine]
                .entries
                .iter()
                .map(|e| (e.cell_x, e.cell_w))
                .collect();
            let want: Vec<_> = control.rows[theirs]
                .entries
                .iter()
                .map(|e| (e.cell_x, e.cell_w))
                .collect();
            assert_eq!(got, want, "row {mine} moved");
        }
        // The row above the gap is untouched: whatever the empty row does, it
        // does it downward.
        assert_eq!(gapped.rows[0].height, control.rows[0].height);

        // And the row below it differs by exactly one border width — the top
        // edge it owns and the control's does not. Asserted as the difference
        // rather than as two heights, so it stays a statement about the border
        // and not about the line.
        assert_eq!(
            gapped.rows[2].height - control.rows[1].height,
            Pt::new(0.5),
            "the row below an empty one keeps its own top border"
        );

        // …and the populated rows really are 14pt of line plus the table's own
        // 0.5pt top border, which is drawn inside the first row's box, so the
        // zero above is a difference and not the table's uniform answer.
        assert_eq!(gapped.rows[0].height, Pt::new(14.5));
    }

    // ── §17.4.66: a border wider than the margin it sits behind ──────────

    /// The cell's content is laid out inside `cell_w` less however much its own
    /// borders stick out past its margins, so text and border cannot overlap.
    /// Nothing pinned that deflation: `cell_w` is unchanged by it, so it is
    /// only visible in what the content does with the width it is given.
    ///
    /// Two words of 30 pt in a 62 pt cell: they fit on one line with no border,
    /// and a 10 pt left border leaves 52 pt, which is one word short.
    #[test]
    fn a_border_wider_than_the_margin_narrows_the_content_and_can_force_a_wrap() {
        let heavy = single(10.0);
        let with_left_border = TableCellInput {
            cell_borders: Some(CellBorderConfig {
                top: None,
                bottom: None,
                left: Some(CellBorderOverride::Border(heavy)),
                right: None,
            }),
            ..cell_of(2)
        };
        let measure = |c: TableCellInput| {
            measure_table_rows(
                &[row(vec![c])],
                &[Pt::new(62.0)],
                Pt::ZERO,
                Pt::new(10.0),
                None,
                None,
                false,
            )
        };

        let plain = measure(cell_of(2));
        let bordered = measure(with_left_border);
        assert_eq!(
            plain.rows[0].height,
            Pt::new(14.0),
            "60 pt of text fits across a 62 pt cell"
        );
        assert_eq!(
            bordered.rows[0].height,
            Pt::new(28.0),
            "…and not across the 52 pt the 10 pt border leaves"
        );
        assert_eq!(
            bordered.rows[0].entries[0].cell_w, plain.rows[0].entries[0].cell_w,
            "the cell's own box is unchanged — only its content width narrows"
        );
    }

    // ── §17.4.48: a `gridSpan` the grid cannot hold ──────────────────────

    /// A row may address more grid columns than the grid declares. Upstream
    /// `seat_every_cell` grows the grid so this cannot reach layout from a
    /// document, but the clamp here is what keeps a malformed input from
    /// panicking on the slice — and a clamp that inverted the range would panic
    /// just as an unclamped end would.
    ///
    /// The cell gets every column that exists, which is the most width it could
    /// have had: its box ends exactly at the table's right edge.
    #[test]
    fn a_span_past_the_end_of_the_grid_takes_every_column_there_is() {
        let cols = vec![Pt::new(100.0), Pt::new(100.0)];
        let m = measure_table_rows(
            &[row(vec![cell(5, None)])],
            &cols,
            Pt::ZERO,
            Pt::new(10.0),
            Some(&all_single()),
            None,
            false,
        );
        let e = &m.rows[0].entries[0];
        assert_eq!(e.cell_x, Pt::ZERO);
        assert_eq!(e.cell_w, Pt::new(200.0), "both declared columns, not five");
        assert_eq!(e.cell_x + e.cell_w, m.table_width, "flush with the table");
    }
}
