//! Table command emission — positions cells and emits border commands.

use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

use crate::render::layout::draw_command::DrawCommand;

use super::borders::{
    border_width, emit_cell_frame, BoundarySource, CellBorders, CellBox, CellEdge, PlacedRow,
};
use super::grid::is_vmerge_continue;
use super::types::{
    CellVAlign, MeasuredRow, MeasuredTable, TableBorderLine, TableRowInput, VerticalMergeState,
};

/// Layered command buffers for table rendering: shading, content, borders.
///
/// Concatenated in that order — shading, then content, then borders — and the
/// order is load-bearing: a cell's background is painted before its
/// neighbours' borders, so a background can never cover a shared edge, and
/// text is never hidden under a background.
pub(super) struct TableCommandBuffers<'a> {
    pub(super) commands: &'a mut Vec<DrawCommand>,
    pub(super) content_commands: &'a mut Vec<DrawCommand>,
    pub(super) border_commands: &'a mut Vec<DrawCommand>,
}

/// The vertical state a page slice accumulates as rows are emitted into it.
///
/// Two things travel together because neither survives the slice: where the next
/// row's box starts, and every row placed so far. The second is what makes the
/// border grid expressible at all — a border stands on a boundary *between* two
/// rows, and which two rows those are is a fact about this page rather than
/// about the table. A continuation slice puts a repeated header above a row that
/// does not follow it, and the other half of a split row above nothing at all.
pub(super) struct SliceCursor {
    /// Where the next row's box starts, in table-local coordinates.
    pub(super) y: Pt,
    /// Every row this slice has placed, top to bottom.
    pub(super) placed: Vec<PlacedRow>,
}

impl SliceCursor {
    /// A fresh slice: at its top, with nothing placed.
    pub(super) fn new() -> Self {
        Self {
            y: Pt::ZERO,
            placed: Vec::new(),
        }
    }
}

/// Where a row sits in its table, for the one §17.4.84 question a row cannot
/// answer from its own `MeasuredRow`: how tall a `vMerge="restart"` cell's whole
/// merged span is.
///
/// `emit_one_row` takes it as an `Option`, and `None` is not "the caller did not
/// bother". A split row's half is the only caller that passes it, and
/// `build_row_groups` flags a group non-splittable when any of its cells is
/// merged — so a split half is never part of a merge and there is nothing for
/// this lookup to find.
#[derive(Clone, Copy)]
struct RowContext<'a> {
    measured: &'a MeasuredTable,
    rows: &'a [TableRowInput],
    row_idx: usize,
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
    cursor: &mut SliceCursor,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
) {
    let num_rows = measured.rows.len();
    let range_start = row_range.start;
    for row_idx in row_range {
        let mr = &measured.rows[row_idx];
        let is_first_in_range = row_idx == range_start;
        // Deliberately the *table's* row count, not the range's: this mirrors
        // `measure_table_rows`' own condition for reserving `border_gap_below`
        // (`row_idx + 1 < num_rows`), so the two always agree about whether a
        // gap exists. A row that ends a page slice but not the table still has
        // its reserved gap, and its bottom border belongs in it.
        let has_reserved_bottom_gap = row_idx + 1 < num_rows;
        emit_one_row(
            mr,
            &rows[row_idx],
            cursor,
            bufs,
            if is_first_in_range {
                top_border_override
            } else {
                None
            },
            has_reserved_bottom_gap,
            Some(RowContext {
                measured,
                rows,
                row_idx,
            }),
        );
    }
}

/// Emit a custom `MeasuredRow` (produced by `split::split_row_at`). Unlike
/// the range-based emit above, this takes a single already-built
/// `MeasuredRow` and the matching `TableRowInput`. Its [`RowContext`] is always
/// `None` — split rows can't contain vMerge (the group is flagged
/// not splittable if any cell is merged).
pub(super) fn emit_split_row(
    mr: &MeasuredRow,
    row: &TableRowInput,
    cursor: &mut SliceCursor,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
    has_reserved_bottom_gap: bool,
) {
    emit_one_row(
        mr,
        row,
        cursor,
        bufs,
        top_border_override,
        has_reserved_bottom_gap,
        None,
    );
}

fn emit_one_row(
    mr: &MeasuredRow,
    row: &TableRowInput,
    cursor: &mut SliceCursor,
    bufs: &mut TableCommandBuffers<'_>,
    top_border_override: Option<TableBorderLine>,
    // Whether space was reserved below this row for its bottom border, so the
    // border is drawn *under* the content rather than inset into it.
    //
    // Not a positional fact, which is why the two callers derive it
    // differently and both are right: `emit_table_rows` repeats
    // `measure_table_rows`' reservation condition, while a split row's halves
    // carry their own (`split_row_at` zeroes the first half's gap — a cut edge
    // has none — and passes the original's to the second).
    has_reserved_bottom_gap: bool,
    // §17.4.84: where this row sits in its table, for the two lookups a row
    // cannot answer alone. `None` is not missing information — see
    // [`RowContext`].
    row_ctx: Option<RowContext<'_>>,
) {
    // §17.4.45: the row's box starts one cell-spacing below the cursor; its
    // content box is what remains. Both are zero-cost when no spacing is set.
    // §17.4.66: borders collapse exactly when the table sets no
    // `w:tblCellSpacing`, and `leading_gap` is that spacing — `measure_table_rows`
    // puts it on every row. It decides where a border sits, not merely whether
    // the resolution pass ran: collapsed, adjacent cells share an edge and the
    // border is centred on it; spaced, each keeps its border inside its own box.
    let collapsed = mr.leading_gap <= Pt::ZERO;
    let leading = mr.leading_gap;
    let row_top = cursor.y + leading;
    let row_height = mr.height - leading;
    // §17.4.38: the band under this row, if one was reserved. Not a positional
    // fact — see `has_reserved_bottom_gap`.
    let band_below = if has_reserved_bottom_gap {
        mr.border_gap_below
    } else {
        Pt::ZERO
    };

    // §17.4.66: where this row's two boundaries fall, and where the line on the
    // upper one comes from. `bottom` is the middle of the strip §17.4.38
    // reserved, because *the strip is the boundary* — it exists to hold that
    // edge and takes its height from the widest border on it. Where none was
    // reserved (the table's foot, a page cut) the boundary is the content box's
    // own edge.
    let top_source = BoundarySource {
        plan_boundary: match cursor.placed.last() {
            // A §17.4.49 seam — a repeated header above a row that does not
            // follow it. The line is the header's *own* lower boundary, not the
            // one the row below would name.
            Some(prev) if prev.plan_row + 1 != mr.plan_row => prev.plan_row + 1,
            // Adjacent in the plan, or the slice's first row: this row's own
            // upper boundary.
            _ => mr.plan_row,
        },
        // §17.4.38: only the slice's first row can have had its top edge left on
        // the page before, so only it takes a restore.
        restore: cursor
            .placed
            .is_empty()
            .then_some(top_border_override)
            .flatten(),
    };
    // The row above already fixed this row's top boundary; the slice's first row
    // has none above it and takes its own content edge.
    let top = cursor.placed.last().map_or(row_top, |prev| prev.bottom);
    cursor.placed.push(PlacedRow {
        plan_row: mr.plan_row,
        top,
        bottom: row_top + row_height + band_below * 0.5,
        top_source,
    });

    for (cell_ci, (entry, cell_input)) in mr.entries.iter().zip(row.cells.iter()).enumerate() {
        // §17.4.84: the merged span, used below for vAlign and here for
        // shading. Hoisted above the shading so both read the same height —
        // shading used `row_height` while vAlign used the span, so a shaded
        // merged cell was coloured across its first row only.
        let effective_h = if cell_input.vertical_merge == Some(VerticalMergeState::Restart) {
            row_ctx
                .map(|ctx| merged_span_height(ctx.measured, ctx.rows, ctx.row_idx, entry.grid_col))
                .unwrap_or(row_height)
        } else {
            row_height
        };

        // §17.4.32 / §17.4.84: a merged cell's shading covers the whole span,
        // and the `Continue` rows do not paint their own. Word treats the
        // continuation cells as part of the `Restart` cell, so its `<w:shd>`
        // governs the merged region; letting a continuation paint over the span
        // would let a stale or differing `shd` on a row that has no independent
        // existence win for that row.
        // Cells in a row abut exactly, so a run of same-coloured ones reaches
        // the page as N rects sharing N−1 edges — a seam under any rasterizer
        // that anti-aliases each fill on its own. They are fused once the page
        // is finished, by
        // [`coalesce_abutting_rects`](crate::render::layout::draw_command::coalesce_abutting_rects),
        // which owns that rule for every producer rather than each producer
        // half-owning it.
        if cell_input.vertical_merge != Some(VerticalMergeState::Continue) {
            if let Some(color) = cell_input.shading {
                bufs.commands.push(DrawCommand::Rect {
                    rect: PtRect::from_xywh(entry.cell_x, row_top, entry.cell_w, effective_h),
                    color,
                });
            }
        }

        // §17.4.38: restore the top border when this row starts a slice and the
        // resolved top was removed by conflict resolution or adjacent-table
        // collapse. An edge the author set to `nil` must NOT be restored — they
        // asked for no border — and `CellEdge` already says which is which, so
        // this reads the resolved edge instead of re-deriving intent from
        // `cell_input`. That also fixes the `none` case for free: `none` now
        // resolves to `Absent` and is restorable, where the old re-derivation
        // lumped it in with `nil` and left a continuation slice with no top.
        let cell_top = match mr.borders[cell_ci].top {
            CellEdge::Absent => top_border_override.into(),
            resolved => resolved,
        };
        let b_left = mr.borders[cell_ci].left;
        let b_right = mr.borders[cell_ci].right;
        let b_bottom = mr.borders[cell_ci].bottom;

        let dx = entry.content_dx;
        let dy_border = (border_width(cell_top) - cell_input.margins.top).max(Pt::ZERO);
        // The foot of the content box, mirroring `dy_border` at the head. A
        // bottom border only takes room from the content where it is drawn
        // *inside* the box — `emit_cell_borders` puts it there exactly when no
        // strip was reserved below — so the branch is on `band_below` and not
        // on the border alone. `measure_table_rows` reserved the same amount;
        // subtracting it here is what keeps `Bottom` and `Center` from pushing
        // content back down into the border that room was made for.
        let dy_bottom = if band_below > Pt::ZERO {
            Pt::ZERO
        } else {
            (border_width(b_bottom) - cell_input.margins.bottom).max(Pt::ZERO)
        };

        // §17.4.84: for vMerge=Restart cells, vAlign operates over the whole
        // merged span (`effective_h`, computed above with the shading).
        let content_h = entry.layout.content_height + cell_input.margins.vertical();
        let slack = (effective_h - content_h - dy_border - dy_bottom).max(Pt::ZERO);
        let dy_valign = match cell_input.vertical_align {
            CellVAlign::Bottom => slack,
            CellVAlign::Center => slack * 0.5,
            CellVAlign::Top => Pt::ZERO,
        };

        for cmd in &entry.layout.commands {
            let mut cmd = cmd.clone();
            cmd.shift(entry.cell_x + dx, row_top + dy_border + dy_valign);
            bufs.content_commands.push(cmd);
        }

        // §17.4.45: a spaced table has no grid to collapse onto, so each cell
        // closes its own frame here. A collapsed one paints nothing per cell —
        // its borders stand on grid lines and are rasterized once for the whole
        // slice, by `SliceBuilder::finish`.
        if !collapsed {
            emit_cell_frame(
                bufs.border_commands,
                &CellBorders {
                    top: cell_top,
                    bottom: b_bottom,
                    left: b_left,
                    right: b_right,
                },
                CellBox {
                    x: entry.cell_x,
                    w: entry.cell_w,
                    y: row_top,
                    h: row_height,
                },
            );
        }
    }

    cursor.y += mr.height + mr.border_gap_below;
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
    // The span starts below its own leading gap; every row it swallows
    // contributes that row's full height, gap included, because a merged cell
    // covers the spacing between the rows it spans.
    let mut total = measured.rows[start_row].height - measured.rows[start_row].leading_gap;
    let mut row = start_row + 1;
    while row < rows.len() && is_vmerge_continue(&rows[row], grid_col) {
        total += measured.rows[row - 1].border_gap_below;
        total += measured.rows[row].height;
        row += 1;
    }
    total
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::geometry::{PtEdgeInsets, PtSize};
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
    use crate::render::layout::table::types::TableCellInput;
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    const GREY: RgbColor = RgbColor {
        r: 200,
        g: 200,
        b: 200,
    };

    fn frag(text: &str) -> Fragment {
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
            width: Pt::new(30.0),
            trimmed_width: Pt::new(30.0),
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

    fn cell(
        n_lines: usize,
        vmerge: Option<VerticalMergeState>,
        shading: Option<RgbColor>,
    ) -> TableCellInput {
        TableCellInput {
            blocks: (0..n_lines)
                .map(|i| LayoutBlock::Paragraph {
                    fragments: vec![frag(&format!("L{i}"))],
                    style: ParagraphStyle::default(),
                    page_break_before: false,
                    footnotes: vec![],
                    floating_images: vec![],
                    floating_shapes: vec![],
                })
                .collect(),
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading,
            cell_borders: None,
            vertical_merge: vmerge,
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

    fn shading_rects(commands: &[DrawCommand]) -> Vec<(f32, f32)> {
        commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color } if *color == GREY => {
                    Some((rect.origin.y.raw(), rect.size.height.raw()))
                }
                _ => None,
            })
            .collect()
    }

    /// Two rows: col 0 is a shaded `Restart` over a `Continue`, col 1 carries
    /// two 2-line cells so each row measures 28pt — a 56pt merged span.
    fn merged_shading_table(continue_shading: Option<RgbColor>) -> Vec<TableRowInput> {
        vec![
            row(vec![
                cell(1, Some(VerticalMergeState::Restart), Some(GREY)),
                cell(2, None, None),
            ]),
            row(vec![
                cell(0, Some(VerticalMergeState::Continue), continue_shading),
                cell(2, None, None),
            ]),
        ]
    }

    /// §17.4.84 + §17.4.32: a shaded merged cell is shaded across the whole
    /// span, not just its first row.
    ///
    /// The shading rect used `row_height` while vAlign used the span height, so
    /// only the first row was coloured — the lower half of a shaded merged cell
    /// rendered as unshaded background.
    #[test]
    fn shaded_vmerge_restart_cell_shades_the_whole_span() {
        let rows = merged_shading_table(None);
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0), Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        assert_eq!(table.size, PtSize::new(Pt::new(100.0), Pt::new(56.0)));
        assert_eq!(
            shading_rects(&table.commands),
            vec![(0.0, 56.0)],
            "one rect covering the full 56pt merged span"
        );
    }

    /// A `Continue` cell does not paint its own shading: Word treats it as part
    /// of the `Restart` cell, whose `<w:shd>` governs the merged region. Without
    /// this the continuation would paint over the span rect for its own row.
    #[test]
    fn vmerge_continue_cell_does_not_paint_its_own_shading() {
        let rows = merged_shading_table(Some(RgbColor { r: 1, g: 2, b: 3 }));
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0), Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        // The span rect survives whole …
        assert_eq!(shading_rects(&table.commands), vec![(0.0, 56.0)]);
        // … and the continuation's own colour is never emitted.
        assert!(
            !table.commands.iter().any(|c| matches!(
                c,
                DrawCommand::Rect {
                    color: RgbColor { r: 1, g: 2, b: 3 },
                    ..
                }
            )),
            "a Continue cell's shading must not override the merged cell's"
        );
    }

    /// An unmerged shaded cell is unaffected — the fix must not widen every
    /// cell's shading to some span.
    #[test]
    fn unmerged_shaded_cell_shades_only_its_own_row() {
        let rows = vec![
            row(vec![cell(1, None, Some(GREY))]),
            row(vec![cell(1, None, None)]),
        ];
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );
        assert_eq!(shading_rects(&table.commands), vec![(0.0, 14.0)]);
    }

    /// `merged_span_height` stops at the first row that does not continue *this*
    /// grid column — the row below the span is not absorbed into it.
    ///
    /// Observed through the shading rect, since that (with vAlign) is what the
    /// function feeds; it does not affect row heights. Rows 0-1 form the span
    /// and split the restart cell's 14pt of content 7/7 (§17.4.84
    /// distribution), so the span is 14pt while the table is 28pt. A walk that
    /// ran past the `Continue` would shade all 28.
    #[test]
    fn merged_span_height_stops_at_the_first_non_continue_row() {
        let rows = vec![
            row(vec![cell(1, Some(VerticalMergeState::Restart), Some(GREY))]),
            row(vec![cell(0, Some(VerticalMergeState::Continue), None)]),
            row(vec![cell(1, None, None)]),
        ];
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        assert_eq!(
            table.size.height,
            Pt::new(28.0),
            "span 14pt + row 2 at 14pt"
        );
        assert_eq!(
            shading_rects(&table.commands),
            vec![(0.0, 14.0)],
            "the span ends at the Continue row; row 2 is not part of it"
        );
    }

    // ── Buffer layering and the nil/override interaction (E5b#3) ─────────

    const RED: RgbColor = RgbColor { r: 255, g: 0, b: 0 };

    fn single(width: f32, color: RgbColor) -> TableBorderLine {
        TableBorderLine {
            width: Pt::new(width),
            color,
            style: crate::render::layout::table::types::TableBorderStyle::Single,
        }
    }

    fn all_edges(line: TableBorderLine) -> crate::render::layout::table::TableBorderConfig {
        crate::render::layout::table::TableBorderConfig {
            top: Some(line),
            bottom: Some(line),
            left: Some(line),
            right: Some(line),
            inside_h: Some(line),
            inside_v: Some(line),
        }
    }

    /// Index of the first command matching `pred`.
    fn index_of(cmds: &[DrawCommand], pred: impl Fn(&DrawCommand) -> bool) -> usize {
        cmds.iter()
            .position(pred)
            .expect("expected command not emitted")
    }

    /// Emit a range of rows into a fresh slice and rasterize the border grid it
    /// placed — the two halves `SliceBuilder` puts together.
    ///
    /// A test that reads border commands needs both, and only both: placement
    /// decides which of the plan's boundaries this page shows and where, and
    /// rasterization turns that into rects. Neither half emits a border on its
    /// own, which is the point of the split.
    fn slice_border_commands(
        measured: &MeasuredTable,
        rows: &[TableRowInput],
        range: std::ops::Range<usize>,
        top_border_override: Option<TableBorderLine>,
    ) -> Vec<DrawCommand> {
        let (mut commands, mut content, mut borders) = (Vec::new(), Vec::new(), Vec::new());
        let mut cursor = SliceCursor::new();
        emit_table_rows(
            measured,
            rows,
            range,
            &mut cursor,
            &mut TableCommandBuffers {
                commands: &mut commands,
                content_commands: &mut content,
                border_commands: &mut borders,
            },
            top_border_override,
        );
        crate::render::layout::table::borders::rasterize_border_grid(
            &mut borders,
            &measured.plan,
            &measured.grid_x,
            &cursor.placed,
            crate::render::geometry::PtSize::new(measured.table_width, cursor.y),
        );
        borders
    }

    /// The three buffers concatenate as shading → content → borders, and that
    /// order is load-bearing: a cell's background is painted before its
    /// neighbours' borders, so a background can never cover a shared edge, and
    /// text is never hidden under a background.
    ///
    /// Asserted on the *final* command list, which is what the painter walks —
    /// the buffers themselves are an implementation detail.
    #[test]
    fn commands_layer_shading_then_content_then_borders() {
        let rows = vec![row(vec![cell(1, None, Some(GREY))])];
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&all_edges(single(1.0, RED))),
            None,
            false,
        );

        let shading = index_of(
            &table.commands,
            |c| matches!(c, DrawCommand::Rect { color, .. } if *color == GREY),
        );
        let text = index_of(&table.commands, |c| matches!(c, DrawCommand::Text { .. }));
        let border = index_of(
            &table.commands,
            |c| matches!(c, DrawCommand::Rect { color, .. } if *color == RED),
        );

        assert!(
            shading < text && text < border,
            "expected shading({shading}) < content({text}) < borders({border})"
        );
    }

    /// §17.4.38: a continuation slice restores a top border that *resolution*
    /// removed — but never one the author removed with `<w:top w:val="nil"/>`.
    ///
    /// Both cells resolve to no top border, so both are candidates for the
    /// override; only the one without an explicit nil may take it. Driving
    /// `emit_table_rows` directly is what makes the two cases comparable in a
    /// single row.
    #[test]
    fn top_border_override_skips_a_cell_that_explicitly_suppressed_its_top() {
        use crate::render::layout::table::types::{CellBorderConfig, CellBorderOverride};

        let nil_top = CellBorderConfig {
            top: Some(CellBorderOverride::Suppress),
            bottom: None,
            left: None,
            right: None,
        };
        let mut suppressed = cell(1, None, None);
        suppressed.cell_borders = Some(nil_top);

        // Cell 0 says "no top border, deliberately"; cell 1 says nothing.
        let rows = vec![row(vec![suppressed, cell(1, None, None)])];
        // No table borders at all, so both cells resolve `top: None`.
        let measured = crate::render::layout::table::measure::measure_table_rows(
            &rows,
            &[Pt::new(50.0), Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );
        assert!(
            measured.rows[0]
                .borders
                .iter()
                .all(|b| b.top.line().is_none()),
            "both cells must resolve to no top border for this test to mean anything"
        );

        let borders = slice_border_commands(&measured, &rows, 0..1, Some(single(3.0, RED)));

        // The restored line as an x-interval. The grid columns are 50pt each, so
        // "cell 1 only" is 50..100 — asserted as the span rather than as a rect
        // count, because the rasterizer is free to split one line at a junction
        // and a count would then measure the decomposition instead of the line.
        let restored: Vec<(f32, f32)> = borders
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color } if *color == RED => Some((
                    rect.origin.x.raw(),
                    rect.origin.x.raw() + rect.size.width.raw(),
                )),
                _ => None,
            })
            .collect();
        let reach = |x: f32| restored.iter().any(|&(a, b)| a <= x && x <= b);
        assert!(
            !reach(25.0),
            "cell 0 asked for no top border and must not get one: {restored:?}"
        );
        assert!(
            reach(75.0),
            "cell 1 said nothing, so §17.4.38 restores its top: {restored:?}"
        );
    }

    /// The override applies only to the *first* row of the emitted range — a
    /// continuation slice gets one restored top edge, not one per row.
    #[test]
    fn top_border_override_applies_only_to_the_first_row_of_the_range() {
        let rows = vec![
            row(vec![cell(1, None, None)]),
            row(vec![cell(1, None, None)]),
        ];
        let measured = crate::render::layout::table::measure::measure_table_rows(
            &rows,
            &[Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let borders = slice_border_commands(&measured, &rows, 0..2, Some(single(3.0, RED)));

        // The restored edge as the set of y values it reaches. One row boundary,
        // not two: whichever y the override lands on, it must be a single one,
        // and it must be the slice's top — asserted as `min`, so a second
        // restored edge lower down shows up as a differing set rather than as a
        // count this decomposition could satisfy by accident.
        let mut ys: Vec<f32> = borders
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color } if *color == RED => Some(rect.origin.y.raw()),
                _ => None,
            })
            .collect();
        ys.sort_by(f32::total_cmp);
        ys.dedup();
        assert_eq!(
            ys.len(),
            1,
            "exactly one restored boundary, not one per row: {ys:?}"
        );
        assert!(
            ys[0] < Pt::ZERO.raw() + f32::EPSILON,
            "and it is the slice's own top: {ys:?}"
        );
    }

    /// A row that ends a page slice but not the *table* keeps its reserved
    /// bottom-border gap.
    ///
    /// `has_reserved_bottom_gap` is derived from the table's row count, not the
    /// emitted range's — it has to be, because `measure_table_rows` reserves
    /// `border_gap_below` on exactly that condition and `cursor_y` advances by
    /// it. Deriving it per slice instead would inset the border into the last
    /// row of every page, leaving the reserved gap empty.
    ///
    /// Three 14pt rows with 2pt inside borders, split so rows 0-1 land on the
    /// first slice: row 1 ends that slice, and its border must still sit in the
    /// gap below it (y = 30..32), not inside its content box.
    #[test]
    fn a_row_ending_a_slice_keeps_its_reserved_bottom_gap() {
        let line = single(2.0, RED);
        let borders = crate::render::layout::table::TableBorderConfig {
            top: None,
            bottom: Some(line),
            left: None,
            right: None,
            inside_h: Some(line),
            inside_v: None,
        };
        let rows: Vec<TableRowInput> = (0..3).map(|_| row(vec![cell(1, None, None)])).collect();
        let slices = crate::render::layout::table::layout_table_paginated(
            &rows,
            &[Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&borders),
            None,
            &crate::render::layout::table::TablePaginationConfig {
                available_height: Pt::new(34.0),
                page_height: Pt::new(34.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2, "3 rows at 16pt each over a 34pt page");

        let border_bands: Vec<(f32, f32)> = slices[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color } if *color == RED => Some((
                    rect.origin.y.raw(),
                    rect.origin.y.raw() + rect.size.height.raw(),
                )),
                _ => None,
            })
            .collect();

        assert!(
            border_bands.contains(&(30.0, 32.0)),
            "row 1 ends the slice at y=30; its bottom border belongs in the \
             reserved gap 30..32, got {border_bands:?}"
        );
    }

    // ── The corner square at a row boundary ─────────────────────────────────

    /// The spacer column a generated form uses as a visual gutter: a narrow
    /// cell declaring §17.4.66 `nil` on its top and bottom, saying nothing
    /// about its sides (they inherit `insideV`).
    fn nil_top_bottom_cell() -> TableCellInput {
        use crate::render::layout::table::types::{CellBorderConfig, CellBorderOverride};
        let mut c = cell(1, None, None);
        c.cell_borders = Some(CellBorderConfig {
            top: Some(CellBorderOverride::Suppress),
            bottom: Some(CellBorderOverride::Suppress),
            left: None,
            right: None,
        });
        c
    }

    /// Two rows of `[cell | nil-top-and-bottom spacer | cell]` under a
    /// `Tabellenraster`-shaped border config — every side plus insideH/insideV
    /// at 0.5pt. Columns 100 / 20 / 100, so the spacer's right edge (the
    /// vertical it wins from `insideV`) is painted at x = 119.5..120.
    fn spacer_table() -> crate::render::layout::table::TableSlice {
        let rows = vec![
            row(vec![
                cell(1, None, None),
                nil_top_bottom_cell(),
                cell(1, None, None),
            ]),
            row(vec![
                cell(1, None, None),
                nil_top_bottom_cell(),
                cell(1, None, None),
            ]),
        ];
        crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(100.0), Pt::new(20.0), Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&all_edges(single(0.5, RED))),
            None,
            false,
        )
    }

    /// Every border rect as `(x0, x1, y0, y1)`.
    fn border_rects(cmds: &[DrawCommand]) -> Vec<(f32, f32, f32, f32)> {
        cmds.iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color } if *color == RED => Some((
                    rect.origin.x.raw(),
                    rect.origin.x.raw() + rect.size.width.raw(),
                    rect.origin.y.raw(),
                    rect.origin.y.raw() + rect.size.height.raw(),
                )),
                _ => None,
            })
            .collect()
    }

    /// The y ranges within `0..height` that **no** border rect paints along the
    /// vertical line at `probe_x`. A continuous border line leaves none.
    fn unpainted_along(cmds: &[DrawCommand], probe_x: f32, height: f32) -> Vec<(f32, f32)> {
        let mut spans: Vec<(f32, f32)> = border_rects(cmds)
            .into_iter()
            .filter(|(x0, x1, _, _)| *x0 <= probe_x && probe_x <= *x1)
            .map(|(_, _, y0, y1)| (y0, y1))
            .collect();
        spans.sort_by(|a, b| a.0.total_cmp(&b.0));
        let mut gaps = Vec::new();
        let mut reached = 0.0f32;
        for (y0, y1) in spans {
            if y0 > reached + 0.001 {
                gaps.push((reached, y0));
            }
            reached = reached.max(y1);
        }
        if reached < height - 0.001 {
            gaps.push((reached, height));
        }
        gaps
    }

    /// The reported defect: at every interior row boundary of the user's form,
    /// a 0.5 × 0.5pt square of the spacer column's vertical border was painted
    /// by nobody, showing as a 1-2px hole at the top-left corner of the cell
    /// to its right.
    ///
    /// [MS-OI29500] §17.4.66 leaves each cell's four edges to be painted, and
    /// says nothing about corners — that is this engine's own convention
    /// (see [`emit_cell_borders`]): a horizontal owns the corners of its cell
    /// because it spans the full cell width, and the verticals are inset
    /// between them. The convention is only sound while a horizontal *paints*.
    /// Where the author suppressed it with `nil`, no horizontal owns that
    /// corner and the vertical must not yield it.
    ///
    /// A row's bottom border lives in a band **below** the row's content box,
    /// reserved by `measure_table_rows` at the widest bottom border in the row.
    /// A cell whose own bottom paints nothing contributes nothing to that band
    /// and used to stop short of it, leaving the band unpainted at its own
    /// vertical edges — the hole.
    ///
    /// Asserted over **every** vertical edge in the table rather than the
    /// spacer's alone: in this table every column boundary carries `insideV`
    /// down both rows, so each is one unbroken line from the table's top to its
    /// bottom, and any convention that leaves a corner to nobody breaks one of
    /// them.
    #[test]
    fn every_vertical_border_is_unbroken_from_the_tables_top_to_its_bottom() {
        let table = spacer_table();
        let h = table.size.height.raw();
        assert_eq!(h, 56.5, "28pt row + 0.5pt reserved band + 28pt row");

        let rects = border_rects(&table.commands);
        let mut probes: Vec<String> = Vec::new();
        for (x0, x1, y0, y1) in rects.iter().copied() {
            if x1 - x0 >= y1 - y0 {
                continue; // horizontal
            }
            let probe_x = (x0 + x1) * 0.5;
            let gaps = unpainted_along(&table.commands, probe_x, h);
            if !gaps.is_empty() {
                probes.push(format!("x={probe_x}: unpainted {gaps:?}"));
            }
        }
        // The four column boundaries: the table's own two sides and the two
        // interior edges `insideV` supplies (one of them the spacer's right).
        let xs: std::collections::BTreeSet<i32> = rects
            .iter()
            .filter(|(x0, x1, y0, y1)| x1 - x0 < y1 - y0)
            .map(|(x0, _, _, _)| (x0 * 10.0).round() as i32)
            .collect();
        assert_eq!(
            xs.len(),
            4,
            "expected four vertical edges to probe, got {xs:?}"
        );
        assert!(
            probes.is_empty(),
            "border lines broken: {probes:?} in rects {rects:?}"
        );
    }

    /// The other half of the same invariant: a corner is painted by *exactly*
    /// one edge. Extending every vertical through the reserved band would
    /// satisfy the test above and paint each bordered cell's corners twice —
    /// invisible while the two edges share a colour, and wrong the moment they
    /// do not, since the later rect would win the corner.
    #[test]
    fn no_two_border_rects_overlap() {
        for (name, table) in [("spacer", spacer_table()), ("gutter", gutter_row_table())] {
            let rects = border_rects(&table.commands);
            for (i, a) in rects.iter().enumerate() {
                for b in &rects[i + 1..] {
                    let overlap =
                        a.1.min(b.1) - a.0.max(b.0) > 0.001 && a.3.min(b.3) - a.2.max(b.2) > 0.001;
                    assert!(!overlap, "{name}: border rects overlap: {a:?} and {b:?}");
                }
            }
        }
    }

    // ── The junction, which is where all three corner defects were ──────────

    /// A cell declaring §17.4.66 `nil` on both of its **sides**, so the row it
    /// sits in paints no vertical border there at all.
    fn nil_sides_cell() -> TableCellInput {
        use crate::render::layout::table::types::{CellBorderConfig, CellBorderOverride};
        let mut c = cell(1, None, None);
        c.cell_borders = Some(CellBorderConfig {
            top: None,
            bottom: None,
            left: Some(CellBorderOverride::Suppress),
            right: Some(CellBorderOverride::Suppress),
        });
        c
    }

    /// The spacer cell of a **gutter row**: `nil` on all four edges, so neither
    /// a horizontal nor a vertical paints anywhere on it.
    fn nil_all_edges_cell() -> TableCellInput {
        use crate::render::layout::table::types::{CellBorderConfig, CellBorderOverride};
        let mut c = cell(1, None, None);
        c.cell_borders = Some(CellBorderConfig {
            top: Some(CellBorderOverride::Suppress),
            bottom: Some(CellBorderOverride::Suppress),
            left: Some(CellBorderOverride::Suppress),
            right: Some(CellBorderOverride::Suppress),
        });
        c
    }

    /// `spacer_table` with a **gutter row** in the middle — a short row whose
    /// cells declare `nil` on both sides, which the reporting document uses as
    /// the horizontal counterpart of its spacer column. The row still paints
    /// the bottom borders that separate it from the row below, so a band is
    /// reserved under it; what it does not paint is a single vertical.
    ///
    /// That is what makes the junction below the gutter row's spacer different
    /// from every case already covered: the square where the spacer column's
    /// vertical crosses that band belongs, geometrically, to a cell whose own
    /// bottom and own right are both empty — while the vertical that needs it
    /// is one row *below* and the horizontal that stops at it is one column to
    /// the *right*.
    fn gutter_row_table() -> crate::render::layout::table::TableSlice {
        let normal = || {
            row(vec![
                cell(1, None, None),
                nil_top_bottom_cell(),
                cell(1, None, None),
            ])
        };
        let gutter = || {
            row(vec![
                nil_sides_cell(),
                nil_all_edges_cell(),
                nil_sides_cell(),
            ])
        };
        let rows = vec![normal(), gutter(), normal()];
        crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(100.0), Pt::new(20.0), Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&all_edges(single(0.5, RED))),
            None,
            false,
        )
    }

    /// Every junction square that **no** rect paints.
    ///
    /// A junction is where a vertical border and a horizontal one meet: the
    /// vertical's x-band crossed with the horizontal's y-band, for any pair that
    /// touch or overlap on both axes. Two borders that meet there are two lines
    /// joining, and the square they join in has to be ink — whichever of them
    /// paints it. A grid that leaves one empty shows the 1–2px notch three
    /// separate reports have now described.
    ///
    /// Deliberately *not* a per-cell check: every one of those defects was a
    /// square whose two owning edges were both empty while the borders needing
    /// it were in the neighbouring row and the neighbouring column. Asking the
    /// question of the rects alone is what makes it blind to which cell they
    /// came from.
    ///
    /// A square counts as painted when **one** rect covers it, which is the
    /// stronger reading and sound here: these tables have a single grid, so no
    /// junction can straddle the seam between two abutting horizontals. The
    /// corpus-wide audit in `tests/table_border_corners.rs` tests the union
    /// instead, because two tables' grids can differ by a fraction of a point
    /// and do.
    fn unpainted_junctions(cmds: &[DrawCommand]) -> Vec<(f32, f32, f32, f32)> {
        const EPS: f32 = 0.001;
        let rects = border_rects(cmds);
        let (vertical, horizontal): (Vec<_>, Vec<_>) = rects
            .iter()
            .copied()
            .partition(|(x0, x1, y0, y1)| x1 - x0 < y1 - y0);

        let mut missing = Vec::new();
        for (vx0, vx1, vy0, vy1) in vertical {
            for (hx0, hx1, hy0, hy1) in horizontal.iter().copied() {
                // Touching or overlapping on both axes: the two lines meet.
                if vx1 < hx0 - EPS || vx0 > hx1 + EPS || hy1 < vy0 - EPS || hy0 > vy1 + EPS {
                    continue;
                }
                let square = (vx0, vx1, hy0, hy1);
                let covered = rects.iter().any(|(x0, x1, y0, y1)| {
                    *x0 <= square.0 + EPS
                        && *x1 >= square.1 - EPS
                        && *y0 <= square.2 + EPS
                        && *y1 >= square.3 - EPS
                });
                if !covered && !missing.contains(&square) {
                    missing.push(square);
                }
            }
        }
        missing
    }

    /// The junction invariant on the table every existing corner test uses:
    /// a control, so a failure below is about the gutter row and not about the
    /// audit itself.
    #[test]
    fn the_spacer_tables_junctions_are_all_painted() {
        let missing = unpainted_junctions(&spacer_table().commands);
        assert!(
            missing.is_empty(),
            "unpainted junction squares (x0,x1,y0,y1): {missing:?}"
        );
    }

    /// §17.4.38 / §17.18.2: a `double` border crossing a band keeps its own
    /// division — **two lines side by side**, never two stacked.
    ///
    /// A band crossing is a junction, and a junction is the *product* of its two
    /// axes' rules (`borders::junction_axes`): the vertical divides it across x
    /// and the horizontal across y. Divided on the wrong axis the crossing still
    /// fills its x — so continuity, the corner audit and the overlap check all
    /// pass while the double border turns into a pair of rungs across the gap.
    ///
    /// `spacer_table`'s two rows at 3pt double borders instead of 0.5pt single
    /// ones. A 3pt double is **9pt** of page — two 3pt rules with a 3pt gap
    /// (`borders::drawn_width`) — so the §17.4.38 strip between the two rows is
    /// 9pt, y 28..37, and the spacer's right edge crosses it on grid line 2 at
    /// x = 120, straddling it from 115.5 to 124.5. Both borders are double, so
    /// the crossing is the 2 x 2 lattice Word draws: two 3pt columns of ink,
    /// each broken by the horizontal's own 3pt gap.
    #[test]
    fn a_double_border_crosses_a_band_as_two_lines_side_by_side() {
        let rows = vec![
            row(vec![
                cell(1, None, None),
                nil_top_bottom_cell(),
                cell(1, None, None),
            ]),
            row(vec![
                cell(1, None, None),
                nil_top_bottom_cell(),
                cell(1, None, None),
            ]),
        ];
        let double = TableBorderLine {
            width: Pt::new(3.0),
            color: RED,
            style: crate::render::layout::table::types::TableBorderStyle::Double,
        };
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(100.0), Pt::new(20.0), Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&all_edges(double)),
            None,
            false,
        );

        // The boundary between the two rows is the middle of the 3pt strip they
        // reserve: 28pt of content, then half of 3. The junction square sits on
        // it, 3pt tall, and the crossing is at grid line 2 — x = 120.
        let (band_top, band_bottom) = (28.0, 37.0);
        // Sorted, because this test is about the crossing's *geometry* and the
        // stream's order is a separate claim with its own owner: the rasterizer
        // emits rule by rule so that each junction rule lands beside the segment
        // it abuts and `coalesce_abutting_rects` can fuse the pair
        // (`tests/table_shading_seams.rs`).
        let mut in_band: Vec<(f32, f32, f32, f32)> = border_rects(&table.commands)
            .into_iter()
            .filter(|(x0, x1, y0, y1)| {
                *y0 >= band_top - 0.001 && *y1 <= band_bottom + 0.001 && *x0 > 114.0 && *x1 < 126.0
            })
            .collect();
        in_band.sort_by(|a, b| a.0.total_cmp(&b.0).then(a.2.total_cmp(&b.2)));
        assert_eq!(
            in_band,
            vec![
                (115.5, 118.5, 28.0, 31.0),
                (115.5, 118.5, 34.0, 37.0),
                (121.5, 124.5, 28.0, 31.0),
                (121.5, 124.5, 34.0, 37.0),
            ],
            "the spacer's right edge crosses the boundary as its own two \
             rules, each a whole `sz` wide and side by side — not as rungs across \
             the gap, which is what splitting on the wrong axis draws"
        );
    }

    /// The reported defect, third manifestation: a junction square below a row
    /// that paints **no** vertical of its own.
    ///
    /// The gutter row's spacer cell has `nil` on all four edges, so the square
    /// where the spacer column's vertical crosses the band under that row is
    /// painted by neither of the two edges that geometrically own it. The
    /// vertical that arrives there comes up from the row *below*, and the
    /// horizontal that stops at it belongs to the cell one column to the
    /// *right* — neither of which the owning cell can see. The user reports it
    /// as a missing corner at the top-left of the cell to the right of the
    /// gutter.
    #[test]
    fn a_junction_below_a_row_that_paints_no_vertical_is_still_painted() {
        let table = gutter_row_table();
        let missing = unpainted_junctions(&table.commands);
        assert!(
            missing.is_empty(),
            "unpainted junction squares (x0,x1,y0,y1): {missing:?} \
             in rects {:?}",
            border_rects(&table.commands)
        );
    }

    // ── §17.4.83 vAlign, and the §17.4.38 border inset it composes with ─────

    /// Every text command as `(x, y)`, sorted by x — so a cell is identified by
    /// the column it sits in rather than by its position in the command list.
    fn text_positions(cmds: &[DrawCommand]) -> Vec<(f32, f32)> {
        let mut v: Vec<(f32, f32)> = cmds
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some((position.x.raw(), position.y.raw())),
                _ => None,
            })
            .collect();
        v.sort_by(|a, b| a.0.total_cmp(&b.0));
        v
    }

    /// §17.4.83: `top`, `center` and `bottom` place a cell's content at
    /// `0`, `(row_h − content_h) / 2` and `row_h − content_h` below the row's
    /// top edge. On a **plain** cell — no `vMerge` — `row_h` is the row's own
    /// height, which `RowHeightRule::Exact` fixes at 60pt here so the arithmetic
    /// does not depend on how tall a line happens to be.
    ///
    /// The three cells are in **one row**, so they share `row_h` by
    /// construction and the two offsets are differences against the `top` cell
    /// rather than against a literal baseline. `content_h` is taken from the
    /// same table without the height rule, where the row *is* its content — so
    /// the 23 and 46 below are derived, not measured off the output.
    #[test]
    fn valign_places_a_plain_cells_content_at_exact_offsets_in_the_row() {
        let aligned = |align: CellVAlign| TableCellInput {
            vertical_align: align,
            ..cell(1, None, None)
        };
        let build = |rule: Option<crate::render::layout::table::RowHeightRule>| {
            let rows = vec![TableRowInput {
                cells: vec![
                    aligned(CellVAlign::Top),
                    aligned(CellVAlign::Center),
                    aligned(CellVAlign::Bottom),
                ],
                height_rule: rule,
                ..row(vec![])
            }];
            crate::render::layout::table::layout_table(
                &rows,
                &[Pt::new(50.0), Pt::new(50.0), Pt::new(50.0)],
                Pt::ZERO,
                Pt::new(14.0),
                None,
                None,
                false,
            )
        };

        // Without a height rule the row is exactly its content, so its height
        // *is* `content_h`. vAlign has nothing to distribute and all three sit
        // at the same y.
        let natural = build(None);
        let content_h = natural.size.height.raw();
        assert_eq!(content_h, 14.0, "one default line");
        let ys: Vec<f32> = text_positions(&natural.commands)
            .iter()
            .map(|p| p.1)
            .collect();
        assert_eq!(
            ys,
            vec![ys[0]; 3],
            "with no spare height every alignment lands in the same place"
        );

        let tall = build(Some(crate::render::layout::table::RowHeightRule::Exact(
            Pt::new(60.0),
        )));
        assert_eq!(tall.size.height, Pt::new(60.0));
        let p = text_positions(&tall.commands);
        assert_eq!(
            p.iter().map(|q| q.0).collect::<Vec<_>>(),
            vec![0.0, 50.0, 100.0],
            "one cell per column, so the y values below are in top/center/bottom order"
        );
        let (top, center, bottom) = (p[0].1, p[1].1, p[2].1);
        assert_eq!(top, ys[0], "`top` is unaffected by the row's spare height");
        assert_eq!(
            center - top,
            (60.0 - content_h) / 2.0,
            "`center` takes half the spare 46pt"
        );
        assert_eq!(
            bottom - top,
            60.0 - content_h,
            "`bottom` takes all 46pt of it"
        );
    }

    /// §17.4.38: a cell's border is drawn inside its box, so a top border wider
    /// than the cell's own top margin pushes the content down by the
    /// **difference** — the margin already holds part of the border's width.
    /// A border no wider than the margin moves nothing.
    ///
    /// Three cells in one row, all with a 2pt top margin: no border, a 1pt
    /// border (narrower than the margin), and a 5pt one. Only the third moves,
    /// and by 5 − 2 = 3pt. Asserted as differences against the first, so the
    /// test does not depend on where a baseline sits inside its line box.
    #[test]
    fn a_top_border_wider_than_the_cell_margin_pushes_content_down_by_the_difference() {
        use crate::render::layout::table::types::{CellBorderConfig, CellBorderOverride};

        let with_top = |width: Option<f32>| TableCellInput {
            margins: PtEdgeInsets::new(Pt::new(2.0), Pt::ZERO, Pt::new(2.0), Pt::ZERO),
            cell_borders: width.map(|w| CellBorderConfig {
                top: Some(CellBorderOverride::Border(TableBorderLine {
                    width: Pt::new(w),
                    color: RED,
                    style: crate::render::layout::table::types::TableBorderStyle::Single,
                })),
                bottom: None,
                left: None,
                right: None,
            }),
            ..cell(1, None, None)
        };

        let rows = vec![row(vec![
            with_top(None),
            with_top(Some(1.0)),
            with_top(Some(5.0)),
        ])];
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0), Pt::new(50.0), Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let p = text_positions(&table.commands);
        assert_eq!(
            p.iter().map(|q| q.0).collect::<Vec<_>>(),
            vec![0.0, 50.0, 100.0]
        );
        assert_eq!(
            p[1].1, p[0].1,
            "a 1pt border fits inside the 2pt margin and moves nothing"
        );
        assert_eq!(
            p[2].1 - p[0].1,
            3.0,
            "a 5pt border against a 2pt margin pushes the content down by 3pt"
        );
    }

    /// …and a **bottom**-aligned cell's content does not move, because the
    /// border inset is taken out of the same spare height the alignment
    /// distributes: `dy_border + (row_h − content_h − dy_border)` is `row_h −
    /// content_h` whatever the border is. §17.4.83 `bottom` means flush with
    /// the cell's bottom, and a top border is not the bottom's business.
    ///
    /// The companion to the test above — that one pins the inset applying, this
    /// one pins it cancelling — and together they are why the two terms are
    /// added rather than either one replacing the other.
    #[test]
    fn a_top_border_does_not_move_bottom_aligned_content() {
        use crate::render::layout::table::types::{CellBorderConfig, CellBorderOverride};

        let bottom_aligned = |width: Option<f32>| TableCellInput {
            vertical_align: CellVAlign::Bottom,
            cell_borders: width.map(|w| CellBorderConfig {
                top: Some(CellBorderOverride::Border(TableBorderLine {
                    width: Pt::new(w),
                    color: RED,
                    style: crate::render::layout::table::types::TableBorderStyle::Single,
                })),
                bottom: None,
                left: None,
                right: None,
            }),
            ..cell(1, None, None)
        };

        let rows = vec![TableRowInput {
            cells: vec![bottom_aligned(None), bottom_aligned(Some(5.0))],
            height_rule: Some(crate::render::layout::table::RowHeightRule::Exact(Pt::new(
                60.0,
            ))),
            ..row(vec![])
        }];
        let table = crate::render::layout::table::layout_table(
            &rows,
            &[Pt::new(50.0), Pt::new(50.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let p = text_positions(&table.commands);
        assert_eq!(p.len(), 2);
        assert_eq!(
            p[1].1, p[0].1,
            "bottom-aligned content is flush with the cell's bottom whether or \
             not the cell has a top border"
        );
    }

    /// The control on how a vertical and the row boundary it crosses divide the
    /// square where they meet: the vertical yields **exactly** the junction and
    /// not a point more.
    ///
    /// `spacer_table`'s columns are 100pt and every border 0.5pt, so grid line 1
    /// is at x = 100 and the vertical there straddles it, 99.75..100.25. Down
    /// that line the ink comes in three pieces — segment, junction, segment —
    /// and what this asserts is that the two segments stop where the junction
    /// starts. A segment that stopped short would leave the hole three reports
    /// have been filed about; one that ran on would paint the square twice and
    /// leave its colour to emission order.
    #[test]
    fn a_bordered_cell_still_yields_the_reserved_band_to_its_own_bottom_border() {
        let table = spacer_table();
        let mut down_the_line: Vec<(f32, f32)> = border_rects(&table.commands)
            .into_iter()
            .filter(|&(x0, x1, ..)| x0 <= 100.0 && 100.0 <= x1)
            .map(|(_, _, y0, y1)| (y0, y1))
            .collect();
        down_the_line.sort_by(|a, b| a.0.total_cmp(&b.0));

        // The boundary between the two rows: 28pt of content, then half of the
        // 0.5pt strip reserved for the widest bottom border on it.
        let boundary = 28.25_f32;
        let half = 0.25_f32;
        // Five pieces, because the table's own top and bottom boundaries have a
        // junction here too: junction, segment, junction, segment, junction.
        assert_eq!(down_the_line.len(), 5, "got {down_the_line:?}");
        assert_eq!(
            down_the_line[2],
            (boundary - half, boundary + half),
            "the junction is the border's own width, centred on the boundary"
        );
        assert_eq!(
            (down_the_line[1].1, down_the_line[3].0),
            (boundary - half, boundary + half),
            "and the two verticals stop exactly at it — neither short nor over"
        );
    }
}
