//! Table layout — 3-pass column sizing, cell layout, border rendering.
//!
//! Pass 1: Compute column widths from grid definitions or equal distribution.
//! Pass 2: Lay out each cell with tight width constraints, determine row heights.
//! Pass 3: Position cells and emit border commands.
//!
//! Widths are fixed *first*, from the declared grid rather than from content,
//! because the three depend on each other in a circle: column width affects
//! how cell content wraps, wrapped content determines row height, and there is
//! no width-independent way to measure a cell. Fixing width from the grid
//! breaks the cycle instead of resolving it.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::layout::draw_command::DrawCommand;

pub(crate) mod borders;
mod emit;
mod grid;
mod measure;
mod split;
mod types;

pub use grid::compute_column_widths;
pub use types::*;

use borders::{emit_table_outline, rasterize_border_grid};
use emit::{emit_split_row, emit_table_rows, SliceCursor, TableCommandBuffers};
use grid::{build_row_groups, row_group_end};
use measure::measure_table_rows;
use split::{find_row_cut, split_row_at, RowCutInput};

/// Measure the first atomic paginator row group without emitting table commands.
///
/// A lookahead for the section stacker's page-fit decisions (does the next
/// atomic group fit before a page/column break, keepNext chains, and so on) —
/// it needs a height to decide with, not commands to draw.
///
/// The following row is included only to resolve the group's shared bottom
/// border; later rows and groups are neither measured nor paginated.
pub(crate) fn measure_leading_table_group_height(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    // §17.4.45 `tblCellSpacing`, resolved to points (zero when unset). The grid
    // slots must already be shrunk by this amount — see
    // `build/table.rs::reserve_cell_spacing`.
    cell_spacing: Pt,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    suppress_first_row_top: bool,
) -> Option<Pt> {
    if rows.is_empty() || col_widths.is_empty() {
        return None;
    }

    let group_end = row_group_end(rows, 0);
    let measured_end = (group_end + 1).min(rows.len());
    let measured = measure_table_rows(
        &rows[..measured_end],
        col_widths,
        cell_spacing,
        default_line_height,
        borders,
        measure_text,
        suppress_first_row_top,
    );

    Some(
        measured.rows[..group_end]
            .iter()
            .map(|row| row.height + row.border_gap_below)
            .sum(),
    )
}

/// The buffers one slice is emitted into, and the tail that closes it.
///
/// Content and borders are collected apart and concatenated only at the end, so
/// every border paints over every cell's content regardless of the order the
/// rows were emitted in. Both emit paths — a table laid out whole and one laid
/// out per page slice — fill the same three buffers and finish the same way,
/// which is what this type is for.
struct SliceBuilder {
    commands: Vec<DrawCommand>,
    content_commands: Vec<DrawCommand>,
    border_commands: Vec<DrawCommand>,
    /// Vertical state, advanced by each `emit_*` call.
    cursor: SliceCursor,
}

impl SliceBuilder {
    fn new() -> Self {
        Self {
            commands: Vec::new(),
            content_commands: Vec::new(),
            border_commands: Vec::new(),
            cursor: SliceCursor::new(),
        }
    }

    /// The cursor and the three buffers, borrowed disjointly so one `emit_*`
    /// call can take both at once.
    fn emit_into(&mut self) -> (&mut SliceCursor, TableCommandBuffers<'_>) {
        let Self {
            commands,
            content_commands,
            border_commands,
            cursor,
        } = self;
        (
            cursor,
            TableCommandBuffers {
                commands,
                content_commands,
                border_commands,
            },
        )
    }

    /// Close the slice.
    ///
    /// `carries_top` / `carries_bottom` say whether this slice holds the
    /// **table's** own top and bottom edges — true for the first and last slice
    /// respectively, and both true for a table laid out whole, which is the
    /// one-slice case. Two things follow from them:
    ///
    /// §17.4.45: each row reserves its own leading gap, so the only gap left to
    /// add is the trailing one at the table's bottom edge. An intermediate slice
    /// ends at a page cut rather than at the table's edge and gets none —
    /// without that, a table that happens to paginate would lose the bottom gap
    /// the same table keeps when it fits on one page.
    ///
    /// §17.4.66: the two constructors part here, and they are exclusive. A
    /// **collapsed** table (`cell_spacing == 0`) has one grid its cells share,
    /// and the whole slice's borders are rasterized from it in one pass — every
    /// rect disjoint from every other, see [`rasterize_border_grid`]. A
    /// **spaced** one has no shared edges at all: each cell closed its own frame
    /// as it was emitted, and only the table's own rectangle is left.
    ///
    /// Issue #168: a spaced table's outer border is its own rectangle, because
    /// the cells no longer touch the table's edges. Left and right bound every
    /// slice; top and bottom follow the same two conditions the gap does.
    /// §17.4.38's `suppress_first_row_top` is deliberately not consulted here:
    /// with a gap between two tables there is no shared edge to collapse, so
    /// there is nothing for it to suppress.
    fn finish(
        mut self,
        measured: &MeasuredTable,
        cell_spacing: Pt,
        borders: Option<&TableBorderConfig>,
        carries_top: bool,
        carries_bottom: bool,
    ) -> TableSlice {
        let table_width = measured.table_width;
        let height = self.cursor.y
            + if carries_bottom {
                cell_spacing
            } else {
                Pt::ZERO
            };

        if cell_spacing > Pt::ZERO {
            emit_table_outline(
                &mut self.border_commands,
                borders,
                PtRect::from_xywh(Pt::ZERO, Pt::ZERO, table_width, height),
                carries_top,
                carries_bottom,
            );
        } else {
            rasterize_border_grid(
                &mut self.border_commands,
                &measured.plan,
                &measured.grid_x,
                &self.cursor.placed,
                PtSize::new(table_width, height),
            );
        }

        self.commands.append(&mut self.content_commands);
        self.commands.append(&mut self.border_commands);
        TableSlice {
            commands: self.commands,
            size: PtSize::new(table_width, height),
        }
    }
}

/// Lay out a table: compute column widths, lay out cells, emit borders.
///
/// Monolithic — the whole table fits on one page, so unlike
/// [`layout_table_paginated`] there is no top-border override to thread
/// through: every border is already resolved correctly in one pass. The result
/// is a [`TableSlice`] because that is what a one-slice table is.
///
/// §17.4.38: `suppress_first_row_top` suppresses the top border of the first row
/// for adjacent table border collapse.
pub fn layout_table(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    // §17.4.45 `tblCellSpacing`, resolved to points (zero when unset). The grid
    // slots must already be shrunk by this amount — see
    // `build/table.rs::reserve_cell_spacing`.
    cell_spacing: Pt,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    suppress_first_row_top: bool,
) -> TableSlice {
    if rows.is_empty() || col_widths.is_empty() {
        return TableSlice {
            commands: Vec::new(),
            size: PtSize::ZERO,
        };
    }

    let measured = measure_table_rows(
        rows,
        col_widths,
        cell_spacing,
        default_line_height,
        borders,
        measure_text,
        suppress_first_row_top,
    );

    let mut builder = SliceBuilder::new();
    // Monolithic table: no top border override needed — borders are resolved correctly.
    let (cursor, mut buffers) = builder.emit_into();
    emit_table_rows(
        &measured,
        rows,
        0..measured.rows.len(),
        cursor,
        &mut buffers,
        None,
    );

    // The whole table is one slice, so it carries both of the table's edges.
    builder.finish(&measured, cell_spacing, borders, true, true)
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

pub(crate) struct TablePaginationHeights<F> {
    pub(crate) available_height: Pt,
    pub(crate) suppress_first_row_top: bool,
    pub(crate) page_height_for_slice: F,
}

/// Lay out a table with page splitting at row boundaries.
///
/// §17.4.49: header rows repeat on each continuation page.
/// §17.4.6: `cantSplit` rows are kept together (moved to next page if needed).
///
/// Returns one `TableSlice` per page.
pub fn layout_table_paginated(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    // §17.4.45 `tblCellSpacing`, resolved to points (zero when unset). The grid
    // slots must already be shrunk by this amount — see
    // `build/table.rs::reserve_cell_spacing`.
    cell_spacing: Pt,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    pagination: &TablePaginationConfig,
) -> Vec<TableSlice> {
    let page_height = pagination.page_height;
    layout_table_paginated_with_page_heights(
        rows,
        col_widths,
        cell_spacing,
        default_line_height,
        borders,
        measure_text,
        TablePaginationHeights {
            available_height: pagination.available_height,
            suppress_first_row_top: pagination.suppress_first_row_top,
            page_height_for_slice: |_| page_height,
        },
    )
}

/// As [`layout_table_paginated`], but with a per-page height rather than one
/// fixed page height — continuation pages can have different header/footer
/// clearance than the first, so the height available to a row varies by which
/// page it lands on.
pub(crate) fn layout_table_paginated_with_page_heights(
    rows: &[TableRowInput],
    col_widths: &[Pt],
    // §17.4.45 `tblCellSpacing`, resolved to points (zero when unset). The grid
    // slots must already be shrunk by this amount — see
    // `build/table.rs::reserve_cell_spacing`.
    cell_spacing: Pt,
    default_line_height: Pt,
    borders: Option<&TableBorderConfig>,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    pagination: TablePaginationHeights<impl FnMut(usize) -> Pt>,
) -> Vec<TableSlice> {
    let TablePaginationHeights {
        available_height,
        suppress_first_row_top,
        mut page_height_for_slice,
    } = pagination;
    if rows.is_empty() || col_widths.is_empty() {
        return vec![TableSlice {
            commands: Vec::new(),
            size: PtSize::ZERO,
        }];
    }

    let measured = measure_table_rows(
        rows,
        col_widths,
        cell_spacing,
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

    // Each slice is a list of items to emit in order: either a range of
    // measured rows (the common case) or a custom (split) row with its own
    // MeasuredRow data.
    let mut slices: Vec<Vec<SliceItem>> = Vec::new();
    let mut current_slice: Vec<SliceItem> = Vec::new();
    let mut remaining = available_height;

    for group in &groups {
        if group.height <= remaining {
            current_slice.push(SliceItem::Range(group.start..group.end));
            remaining -= group.height;
            continue;
        }

        // §17.4.49: header rows are atomic with respect to splitting — they
        // must remain intact so the same row content can repeat verbatim on
        // continuation slices. Non-header rows fall through to the normal
        // §17.4.6 split path.
        let is_header = group.start < header_count;

        // Doesn't fit. Try to split (§17.4.6) before spilling the whole
        // group to the next page. Only non-header single-row groups are
        // splittable — vMerge spans and cantSplit rows set `splittable=false`.
        if !is_header && group.splittable && group.end - group.start == 1 {
            let row_idx = group.start;
            let cut_input = RowCutInput {
                mr: &measured.rows[row_idx],
                row: &rows[row_idx],
                available: remaining,
            };
            if let Some(cut) = find_row_cut(&cut_input) {
                let parts = split_row_at(&measured.rows[row_idx], &cut);
                current_slice.push(SliceItem::Split {
                    row_idx,
                    mr: parts.first,
                });
                slices.push(std::mem::take(&mut current_slice));
                // New page: start with header rows (if any).
                remaining = page_height_for_slice(slices.len());
                if header_count > 0 {
                    current_slice.push(SliceItem::Range(0..header_count));
                    remaining -= header_height;
                }

                // Iteratively place the continuation, splitting again each
                // time it exceeds the new page's remaining space.
                let mut pending = parts.second;
                loop {
                    if pending.height <= remaining {
                        remaining -= pending.height;
                        current_slice.push(SliceItem::Continuation {
                            row_idx,
                            mr: pending,
                        });
                        break;
                    }
                    let sub_cut_input = RowCutInput {
                        mr: &pending,
                        row: &rows[row_idx],
                        available: remaining,
                    };
                    match find_row_cut(&sub_cut_input) {
                        Some(sub_cut) => {
                            let sub = split_row_at(&pending, &sub_cut);
                            current_slice.push(SliceItem::Continuation {
                                row_idx,
                                mr: sub.first,
                            });
                            slices.push(std::mem::take(&mut current_slice));
                            remaining = page_height_for_slice(slices.len());
                            if header_count > 0 {
                                current_slice.push(SliceItem::Range(0..header_count));
                                remaining -= header_height;
                            }
                            pending = sub.second;
                        }
                        None => {
                            // Not even one line of the continuation fits.
                            // This should be rare — a row taller than a
                            // full page of content. Emit it anyway and log.
                            log::warn!(
                                "[table] row {} continuation ({:.1}pt) exceeds \
                                 page content height ({:.1}pt available)",
                                row_idx,
                                pending.height.raw(),
                                remaining.raw(),
                            );
                            current_slice.push(SliceItem::Continuation {
                                row_idx,
                                mr: pending,
                            });
                            break;
                        }
                    }
                }
                continue;
            }
        }

        // No split possible — move the whole group to the next page, but only
        // if the next page is actually roomier than what is left here.
        //
        // A group taller than a whole page never fits anywhere. Advancing
        // unconditionally then abandons `current_slice` while it is still
        // empty — the caller turns that empty leading slice into a page push,
        // so a table starting at the top of a page emitted a **blank page**
        // and the group overflowed the next one just the same (the warning
        // below fired either way). Same shape as the E4c floating-table
        // spillover: acting on a condition the action cannot change.
        //
        // Comparing against the *post-header* room is what makes this exact —
        // repeating headers can leave a fresh page with less usable space than
        // the current one, and in that case staying put overflows by less.
        // `slices.len() + 1` because this decision happens *before* the push
        // that would append the current slice — the index being queried is the
        // page the group would move onto, not the one it is leaving.
        let next_page_height = page_height_for_slice(slices.len() + 1);
        let next_remaining = if !is_header && header_count > 0 {
            next_page_height - header_height
        } else {
            next_page_height
        };
        if next_remaining > remaining {
            slices.push(std::mem::take(&mut current_slice));
            remaining = next_page_height;
            // §17.4.49: prepend the repeating header rows only when this group
            // sits past the headers. When advancing because a header row itself
            // doesn't fit, the row is part of the table's first appearance —
            // emitting `Range(0..header_count)` here would duplicate it.
            if !is_header && header_count > 0 {
                current_slice.push(SliceItem::Range(0..header_count));
                remaining -= header_height;
            }
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
        current_slice.push(SliceItem::Range(group.start..group.end));
        remaining -= group.height;
    }
    slices.push(current_slice);

    // §17.4.38: the table's outer top border, used to restore the top edge
    // on continuation page slices where border conflict resolution or adjacent
    // table collapse removed it.
    let outer_top_border = borders.and_then(|b| b.top);

    // Emit draw commands for each slice.
    let last_slice_idx = slices.len().saturating_sub(1);
    slices
        .iter()
        .enumerate()
        .map(|(slice_idx, items)| {
            let mut builder = SliceBuilder::new();
            for (item_idx, item) in items.iter().enumerate() {
                // First item on each continuation slice (slice_idx > 0) needs
                // its top border restored if it was resolved away. The first
                // slice does NOT get an override — `suppress_first_row_top`
                // semantics are preserved there.
                let top_override = if slice_idx > 0 && item_idx == 0 {
                    outer_top_border
                } else {
                    None
                };
                let (cursor, mut buffers) = builder.emit_into();
                match item {
                    SliceItem::Range(range) => {
                        emit_table_rows(
                            &measured,
                            rows,
                            range.clone(),
                            cursor,
                            &mut buffers,
                            top_override,
                        );
                    }
                    SliceItem::Split { row_idx, mr } | SliceItem::Continuation { row_idx, mr } => {
                        // A split half's bottom border sits in a reserved gap
                        // only when something follows it on *this* slice. The
                        // last item on a page ends at a cut or page edge, where
                        // `split_row_at` reserved no gap, so the border is
                        // inset into the row instead.
                        let has_reserved_bottom_gap = item_idx + 1 < items.len();
                        emit_split_row(
                            mr,
                            &rows[*row_idx],
                            cursor,
                            &mut buffers,
                            top_override,
                            has_reserved_bottom_gap,
                        );
                    }
                }
            }
            // The table's own top edge is on the first slice and its own bottom
            // edge on the last; every slice in between ends at a page cut.
            builder.finish(
                &measured,
                cell_spacing,
                borders,
                slice_idx == 0,
                slice_idx == last_slice_idx,
            )
        })
        .collect()
}

/// One item inside a page slice's emit list.
enum SliceItem {
    /// Emit the contiguous range of measured rows (normal case).
    Range(std::ops::Range<usize>),
    /// Emit a partial row (first half) at the bottom of a page. Shares the
    /// row's `TableRowInput` with `row_idx` but carries partitioned
    /// commands and modified borders.
    Split {
        row_idx: usize,
        mr: types::MeasuredRow,
    },
    /// Emit the continuation (second half) of a split row at the top of a
    /// continuation page.
    Continuation {
        row_idx: usize,
        mr: types::MeasuredRow,
    },
}

#[cfg(test)]
mod tests {
    use super::super::draw_command::DrawCommand;
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::geometry::PtEdgeInsets;
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
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

    // ── layout_table ─────────────────────────────────────────────────────

    #[test]
    fn empty_table() {
        let result = layout_table(&[], &[], Pt::ZERO, Pt::new(14.0), None, None, false);
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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
                grid_before: 0,
                border_overrides: None,
                bidi_override: None,
            },
            TableRowInput {
                cells: vec![simple_cell("c"), simple_cell("d")],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
                bidi_override: None,
            },
        ];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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
                        floating_shapes: vec![],
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        // Column B is only 80 wide, so "long " + "text" (120) wraps
        let col_widths = vec![Pt::new(200.0), Pt::new(80.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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
                    floating_shapes: vec![],
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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
                    floating_shapes: vec![],
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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

    /// Helper: collect all Text command x-positions in command order.
    fn text_x_positions(commands: &[DrawCommand]) -> Vec<f32> {
        commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some(position.x.raw()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn grid_before_offsets_first_cell_x() {
        // §17.4.15: gridBefore=1 + wBefore skips the first grid column. With
        // a 4-column grid [10, 100, 200, 10] and a row with gridBefore=1 and
        // gridAfter=1, the two cells must occupy columns 1 and 2 — not 0 and 1.
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("A"), simple_cell("B")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 1,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(10.0), Pt::new(100.0), Pt::new(200.0), Pt::new(10.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let xs = text_x_positions(&result.commands);
        assert_eq!(xs.len(), 2, "two text fragments expected");
        assert_eq!(xs[0], 10.0, "first cell starts at col 1's left edge (10pt)");
        assert_eq!(
            xs[1], 110.0,
            "second cell starts after col 1 (10 + 100 = 110pt)"
        );

        // The table's overall width is unchanged by gridBefore/gridAfter —
        // they only leave whitespace within rows.
        assert_eq!(result.size.width.raw(), 320.0);
    }

    /// §17.4.14: a row whose cells don't reach the last grid column leaves the
    /// remainder empty — the `gridAfter` region — without stretching into it.
    ///
    /// The row declares no `gridAfter`: layout derives the right edge from
    /// `grid_before` plus the cells' spans, so the trailing gap is a
    /// *consequence* of the cells, not a separate input. (This test previously
    /// set a `grid_after` field that no layout code read, so it asserted the
    /// same positions whether the field was right, wrong, or absent.)
    #[test]
    fn row_shorter_than_the_grid_leaves_the_trailing_columns_empty() {
        // Shaded cells so the emitted rects expose each cell's *width* — text
        // x-positions alone cannot see a cell stretching rightward, since the
        // second cell's x is fixed by the first cell's grid column either way.
        let shaded = |text: &str| TableCellInput {
            shading: Some(RgbColor {
                r: 200,
                g: 200,
                b: 200,
            }),
            ..simple_cell(text)
        };
        let rows = vec![TableRowInput {
            cells: vec![shaded("X"), shaded("Y")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(10.0), Pt::new(100.0), Pt::new(200.0), Pt::new(10.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let xs = text_x_positions(&result.commands);
        assert_eq!(xs, vec![0.0, 10.0], "cells sit at grid columns 0 and 1");

        // Each cell occupies exactly its own column — not the 210pt of unused
        // grid to the right.
        let rects: Vec<(f32, f32)> = result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, .. } => {
                    Some((rect.origin.x.raw(), rect.size.width.raw()))
                }
                _ => None,
            })
            .collect();
        assert_eq!(
            rects,
            vec![(0.0, 10.0), (10.0, 100.0)],
            "cells must not stretch into the trailing grid columns"
        );

        // The table still spans the whole declared grid; the gap is empty, not
        // removed.
        assert_eq!(result.size.width.raw(), 320.0);
    }

    /// §17.4.15 + §17.4.66: a row inset from **both** table edges takes the
    /// outer `left`/`right` at its own two ends, and `inside_v` only between
    /// its cells.
    ///
    /// `grid_before = 1` moves the first cell off the left edge; the two cells
    /// then span only grid columns 1–2 of 4, so the last one stops short of the
    /// right edge as well. §17.4.66 resolves an edge against "cell borders and
    /// outer table borders" — neither of this row's outer ends has a cell
    /// facing it, so the table's border is what faces them, wherever across the
    /// grid they fall.
    ///
    /// Distinct widths identify which border was applied: `left`/`right` are
    /// 4pt and `inside_v` is 1pt, so the whole claim is three rects and their x.
    ///
    /// **This test asserted the opposite until a Word render of
    /// `test-files/grid-gap-borders.docx` settled it** — that a gapped row's
    /// ends are interior and take `inside_v`. It is inverted rather than
    /// deleted because the geometry it pins is the same; only the answer
    /// changed. Its old form merely *forbade* a 4pt rect without saying what
    /// stood at those ends, which is why it also passed under a third reading
    /// (nothing at all) that was equally wrong.
    #[test]
    fn row_inset_from_both_edges_takes_the_outer_borders_at_its_own_ends() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("A"), simple_cell("B")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 1,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(10.0), Pt::new(50.0), Pt::new(50.0), Pt::new(10.0)];
        let line = |w: f32| {
            Some(TableBorderLine {
                width: Pt::new(w),
                color: RgbColor::BLACK,
                style: TableBorderStyle::Single,
            })
        };
        let borders = TableBorderConfig {
            top: None,
            bottom: None,
            left: line(4.0),
            right: line(4.0),
            inside_h: None,
            inside_v: line(1.0),
        };
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            Some(&borders),
            None,
            false,
        );

        // Every vertical border rect, as (x, width). With no top/bottom borders
        // configured there are no horizontals to filter out beyond the height
        // test.
        let mut verticals: Vec<(f32, f32)> = result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, color }
                    if *color == RgbColor::BLACK && rect.size.height.raw() > 1.0 =>
                {
                    Some((rect.origin.x.raw(), rect.size.width.raw()))
                }
                _ => None,
            })
            .collect();
        verticals.sort_by(|a, b| a.0.total_cmp(&b.0));

        // Cell A occupies 10..60 and B 60..110 (grid 10/50/50/10, one column
        // skipped). So three verticals, on grid lines 10, 60 and 110: the row's
        // leading edge, the shared A|B edge resolved once, and the trailing
        // edge. Asserted as centres, because §17.4.66 centres a collapsed border
        // on its grid line — which is what makes the 4pt and the 1pt comparable
        // at all, since their rects begin at different distances from it.
        let centres: Vec<(f32, f32)> = verticals.iter().map(|&(x, w)| (x + w * 0.5, w)).collect();
        assert_eq!(
            centres,
            vec![(10.0, 4.0), (60.0, 1.0), (110.0, 4.0)],
            "leading 4pt w:left, interior 1pt insideV, trailing 4pt w:right"
        );
    }

    #[test]
    fn vmerge_across_rows_with_different_grid_before() {
        // §17.4.84 + §17.4.15: a cell at grid_col 1 in row A (gridBefore=1)
        // can merge vertically with a Continue cell at grid_col 1 in row B
        // (gridBefore=0) — the merge is per absolute grid column. Row B's
        // cell at grid_col 0 has no above-cell (it's in row A's gridBefore
        // region) and must layout independently.
        let row_a = TableRowInput {
            cells: vec![
                TableCellInput {
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: vec![text_frag("Restart", 30.0)],
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
                    vertical_merge: Some(VerticalMergeState::Restart),
                    vertical_align: CellVAlign::Top,
                },
                simple_cell("header"),
            ],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 1,
            border_overrides: None,
            bidi_override: None,
        };
        let row_b = TableRowInput {
            cells: vec![
                simple_cell("row1col0"),
                TableCellInput {
                    blocks: vec![],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: Some(VerticalMergeState::Continue),
                    vertical_align: CellVAlign::Top,
                },
                simple_cell("row1col2"),
            ],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        };
        let col_widths = vec![Pt::new(50.0), Pt::new(100.0), Pt::new(150.0)];
        let result = layout_table(
            &[row_a, row_b],
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let position_of = |needle: &str| -> Option<(f32, f32)> {
            result.commands.iter().find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if text.as_ref() == needle => {
                    Some((position.x.raw(), position.y.raw()))
                }
                _ => None,
            })
        };

        let (restart_x, _) = position_of("Restart").expect("Restart text present");
        let (header_x, header_y) = position_of("header").expect("header text present");
        let (col0_x, col0_y) = position_of("row1col0").expect("row1col0 text present");
        let (col2_x, _) = position_of("row1col2").expect("row1col2 text present");

        // Row A respects gridBefore=1: first cell at col 1's left edge (50pt).
        assert_eq!(restart_x, 50.0, "Restart cell starts at grid col 1");
        assert_eq!(header_x, 150.0, "header cell starts at grid col 2 (50+100)");

        // Row B has its own grid_before=0 — the col-0 cell exists and lays
        // out independently. The col-1 cell is a Continue (no content).
        // Row B's col-2 cell is unaffected by the merge.
        assert_eq!(col0_x, 0.0, "row1col0 starts at grid col 0");
        assert_eq!(col2_x, 150.0, "row1col2 starts at grid col 2 (50+100)");

        // Row B's col-0 cell sits in row 2 (y > 0); it's not stretched by
        // the vMerge happening at col 1.
        assert!(col0_y > 0.0, "row1col0 is on the second row");
        assert_eq!(
            col0_y,
            header_y + 14.0,
            "row1col0 sits exactly one row-height below the row 0 header"
        );
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
                    floating_shapes: vec![],
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(200.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            Pt::ZERO,
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
        let border_line = TableBorderLine {
            width: Pt::new(0.5),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let borders = TableBorderConfig {
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
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }];
        let col_widths = vec![Pt::new(100.0)];

        let render = |suppress| {
            layout_table(
                &rows,
                &col_widths,
                Pt::ZERO,
                Pt::new(14.0),
                Some(&borders),
                None,
                suppress,
            )
        };
        let normal = render(false);
        let suppressed = render(true);

        // §17.4.66: how far ink reaches along a boundary, as the union of the
        // x-intervals of every border rect crossing that y.
        //
        // A *count* of rects would not do, and used to: the grid rasterizer
        // paints each junction square as its own command, so "four borders" is
        // eight rects and the number says nothing about where any of them is.
        // What suppression means is that one boundary is bare and the other
        // three are not, which is what this measures — and it measures it as a
        // difference between two renders of one table, so no coordinate but the
        // 0.5pt border width is pinned.
        let ink_at_y = |slice: &TableSlice, y: Pt| -> Pt {
            slice
                .commands
                .iter()
                .filter_map(|c| match c {
                    DrawCommand::Rect { rect, color } if *color == RgbColor::BLACK => {
                        (rect.origin.y <= y && y <= rect.origin.y + rect.size.height)
                            .then_some(rect.size.width)
                    }
                    _ => None,
                })
                .fold(Pt::ZERO, |a, w| a + w)
        };
        // The top boundary is sampled inside its own band. A *horizontal* on the
        // table's own edge lies wholly within the box (§17.4.66 — see `place_y`
        // in `rasterize_border_grid`), so half a border-width down is inside the
        // top border where one exists. Its two verticals do not: they straddle
        // the box's own edges (`place_x`, measured against
        // `test-files/border-outer-box.docx`), half of each falling outside. So
        // with a top border the band reaches the whole 100pt column plus half a
        // border of overhang at each end, and without one only the two verticals
        // cross that y.
        let top = border_line.width * 0.5;
        let bottom = normal.size.height;
        let mid = bottom * 0.5;

        assert_eq!(
            ink_at_y(&normal, top),
            col_widths[0] + border_line.width,
            "the unsuppressed table paints its own top boundary, right across \
             the column and over the halves of the two verticals straddling its \
             ends"
        );
        assert_eq!(
            ink_at_y(&suppressed, top),
            border_line.width + border_line.width,
            "§17.4.38: suppression leaves the top boundary bare — only the two \
             verticals cross that y"
        );
        // The controls. Suppression reaches that boundary and nothing else: both
        // renders still paint their bottom edge and both verticals, measured at
        // each table's *own* foot because the two are no longer the same height.
        //
        // And the height difference is itself the assertion, not an allowance —
        // it is exactly one border, because a table's own *horizontal* edge is
        // charged to the content box it insets (`measure_table_rows`), so
        // removing one gives that room back. Its verticals are charged nothing,
        // hanging outside the box, and neither of them is what this measures.
        assert_eq!(
            normal.size.height - suppressed.size.height,
            border_line.width,
            "the suppressed top gives its charged width back to the row"
        );
        for (label, y) in [("verticals", mid), ("bottom", bottom)] {
            let (a, b) = (
                ink_at_y(&normal, y),
                ink_at_y(
                    &suppressed,
                    y - (normal.size.height - suppressed.size.height),
                ),
            );
            assert_eq!(a, b, "suppression must not touch the {label}");
            assert!(b > Pt::ZERO, "{label} still painted");
        }
    }

    #[test]
    fn vmerge_outer_border_crosses_sibling_horizontal_border_gap() {
        let border = TableBorderLine {
            width: Pt::new(1.0),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let borders = TableBorderConfig {
            top: Some(border),
            bottom: Some(border),
            left: Some(border),
            right: Some(border),
            inside_h: Some(border),
            inside_v: Some(border),
        };
        let merged_cell = |state| TableCellInput {
            blocks: vec![],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
            cell_borders: None,
            vertical_merge: Some(state),
            vertical_align: CellVAlign::Top,
        };
        let rows = vec![
            TableRowInput {
                cells: vec![
                    simple_cell("recipient"),
                    merged_cell(VerticalMergeState::Restart),
                ],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
                bidi_override: None,
            },
            TableRowInput {
                cells: vec![
                    simple_cell("recipient bank"),
                    merged_cell(VerticalMergeState::Continue),
                ],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
                bidi_override: None,
            },
            TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 2,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: None,
                    vertical_align: CellVAlign::Top,
                }],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
                bidi_override: None,
            },
        ];

        let result = layout_table(
            &rows,
            &[Pt::new(100.0), Pt::new(100.0)],
            Pt::ZERO,
            Pt::new(14.0),
            Some(&borders),
            None,
            false,
        );
        // The ink a ray down the table's right edge (grid line 2, x = 200) meets,
        // merged into runs. Two runs would mean a gap where the merged cell
        // crosses the boundary its sibling column paints a horizontal on; one
        // run means the line is whole.
        //
        // Asserted as the union rather than as a pair of segments, because the
        // grid rasterizer splits every line at its junctions: how many rects
        // arrive is a property of the decomposition, and continuity is not.
        let mut runs: Vec<(f32, f32)> = result
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Rect { rect, color } if *color == RgbColor::BLACK => {
                    let (x0, x1) = (
                        rect.origin.x.raw(),
                        rect.origin.x.raw() + rect.size.width.raw(),
                    );
                    (x0 <= 200.0 && 200.0 <= x1).then(|| {
                        (
                            rect.origin.y.raw(),
                            rect.origin.y.raw() + rect.size.height.raw(),
                        )
                    })
                }
                _ => None,
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
        assert_eq!(
            merged.len(),
            1,
            "a vertically merged outer border must cross the row boundary its \
             sibling column paints a horizontal on: {merged:?}"
        );
        assert!(
            merged[0].1 - merged[0].0 >= result.size.height.raw(),
            "and it must run the whole height of the table: {merged:?}"
        );
    }

    #[test]
    fn valign_bottom_on_vmerge_restart_uses_span_height() {
        // §17.4.84 + §17.4.83: vAlign on a vMerge=Restart cell should
        // apply across the whole merged span, not just the first row.
        //
        // Table: 2 rows × 2 cols.
        //  Row 0: [restart "Total", top-align "header"]
        //  Row 1: [continue,         top-align "value"]
        // Default single-line height is 14pt per row → span = 28pt.
        // Restart cell has bottom alignment, content height 14pt → text
        // should sit 14pt below the cell's top, i.e. inside row 1.
        let row0 = TableRowInput {
            cells: vec![
                TableCellInput {
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: vec![text_frag("Total", 20.0)],
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
                    vertical_merge: Some(VerticalMergeState::Restart),
                    vertical_align: CellVAlign::Bottom,
                },
                simple_cell("header"),
            ],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        };
        let row1 = TableRowInput {
            cells: vec![
                TableCellInput {
                    blocks: vec![],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: Some(VerticalMergeState::Continue),
                    vertical_align: CellVAlign::Top,
                },
                simple_cell("value"),
            ],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        };
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &[row0, row1],
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        // Text position for "Total" comes first (row 0, cell 0).
        let total_y = result
            .commands
            .iter()
            .find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if text.as_ref() == "Total" => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .expect("Total text present");
        let header_y = result
            .commands
            .iter()
            .find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if text.as_ref() == "header" => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .expect("header text present");

        // "header" is top-aligned in row 0, so it marks the span's top. The
        // span is 28pt (two 14pt rows) holding 14pt of content, so §17.4.83
        // `bottom` puts "Total" exactly 14pt lower — one whole row down, which
        // is only expressible because the alignment sees the *span* and not
        // the restart row. A loose inequality here would also pass for a
        // renderer that aligned against row 0 alone and merely overshot.
        assert_eq!(
            total_y - header_y,
            14.0,
            "Total sits `span − content` = 28 − 14 below the span's top; \
             total_y={total_y}, header_y={header_y}"
        );
    }

    // ── Row splitting (§17.4.6) ──────────────────────────────────────────

    /// Build a single-row table with one cell whose paragraph contains
    /// `n_lines` narrow text fragments that wrap to separate lines.
    fn tall_row(n_lines: usize) -> TableRowInput {
        // Width=30 for each fragment; cell content width ≈ 40 ⇒ each
        // fragment wraps to its own line (default line height 14pt).
        let fragments: Vec<Fragment> = (0..n_lines)
            .map(|i| text_frag(&format!("L{i} "), 30.0))
            .collect();
        TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![LayoutBlock::Paragraph {
                    fragments,
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
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }
    }

    /// A paragraph of `n_lines` narrow fragments (each wraps to its own 14pt
    /// line) carrying `style`, for in-cell split tests.
    fn styled_para(n_lines: usize, style: ParagraphStyle) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: (0..n_lines)
                .map(|i| text_frag(&format!("L{i} "), 30.0))
                .collect(),
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    /// A single-cell, single-column row whose cell holds `blocks`.
    fn one_cell_row(blocks: Vec<LayoutBlock>) -> TableRowInput {
        TableRowInput {
            cells: vec![TableCellInput {
                blocks,
                margins: PtEdgeInsets::ZERO,
                grid_span: 1,
                shading: None,
                cell_borders: None,
                vertical_merge: None,
                vertical_align: CellVAlign::Top,
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        }
    }

    /// Number of `Text` draw commands in a slice (one per fitted line here).
    fn slice_line_count(slice: &TableSlice) -> usize {
        slice
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count()
    }

    #[test]
    fn splittable_row_breaks_across_pages() {
        // Row with 6 lines (84pt). Available = 50pt on page 1 ⇒ only ~3
        // lines fit. The row should split: first slice has ~3 lines,
        // second slice has the rest.
        let rows = vec![tall_row(6)];
        let col_widths = vec![Pt::new(40.0)];
        let slices = layout_table_paginated(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(50.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );

        assert!(
            slices.len() >= 2,
            "expected at least 2 slices, got {}",
            slices.len()
        );

        let text_y = |slice: &TableSlice| -> Vec<f32> {
            slice
                .commands
                .iter()
                .filter_map(|c| match c {
                    DrawCommand::Text { position, .. } => Some(position.y.raw()),
                    _ => None,
                })
                .collect()
        };

        let s0 = text_y(&slices[0]);
        let s1 = text_y(&slices[1]);
        assert!(!s0.is_empty(), "slice 0 should contain some lines");
        assert!(!s1.is_empty(), "slice 1 should contain continuation lines");
        assert_eq!(
            s0.len() + s1.len(),
            6,
            "every line should be emitted exactly once across slices"
        );
        // Slice 1's lines should sit near the top (near y=0) since we
        // rebased them.
        let min_s1_y = s1.iter().copied().fold(f32::INFINITY, f32::min);
        assert!(
            min_s1_y < 20.0,
            "slice 1 top text should be near y=0, was {min_s1_y}"
        );
    }

    #[test]
    fn cant_split_row_moves_whole_to_next_page() {
        // cantSplit=true ⇒ entire row moves when it doesn't fit.
        let mut row = tall_row(6);
        row.cant_split = Some(true);
        let rows = vec![row];
        let col_widths = vec![Pt::new(40.0)];
        let slices = layout_table_paginated(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(50.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );

        assert_eq!(slices.len(), 2, "should still produce 2 slices");
        // Slice 0 is empty (or at most has no text); all 6 lines land on slice 1.
        let count0 = slices[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        let count1 = slices[1]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(count0, 0, "first slice has no text with cantSplit");
        assert_eq!(count1, 6, "all 6 lines on second slice");
    }

    #[test]
    fn splittable_row_spans_three_or_more_pages() {
        // Row with 15 lines (≈210pt at 14pt line height).
        // Page 1 has 50pt → ~3 lines fit.
        // Page 2 and on have ~70pt → ~5 lines fit each.
        // Expected: 3+ slices, every line emitted exactly once.
        let rows = vec![tall_row(15)];
        let col_widths = vec![Pt::new(40.0)];
        let slices = layout_table_paginated(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(50.0),
                page_height: Pt::new(70.0),
                suppress_first_row_top: false,
            },
        );

        assert!(
            slices.len() >= 3,
            "expected ≥3 slices for a 15-line row split across pages, got {}",
            slices.len()
        );

        // Every original line should appear exactly once across all slices.
        let total_lines: usize = slices
            .iter()
            .map(|s| {
                s.commands
                    .iter()
                    .filter(|c| matches!(c, DrawCommand::Text { .. }))
                    .count()
            })
            .sum();
        assert_eq!(
            total_lines, 15,
            "all 15 lines should be emitted across the slices exactly once"
        );
    }

    #[test]
    fn vmerge_span_is_not_split_mid_cell() {
        // A vMerge span must stay atomic even when its content would split.
        let row0 = TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![LayoutBlock::Paragraph {
                    fragments: (0..4).map(|i| text_frag(&format!("L{i} "), 30.0)).collect(),
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
                vertical_merge: Some(VerticalMergeState::Restart),
                vertical_align: CellVAlign::Top,
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        };
        let row1 = TableRowInput {
            cells: vec![TableCellInput {
                blocks: vec![],
                margins: PtEdgeInsets::ZERO,
                grid_span: 1,
                shading: None,
                cell_borders: None,
                vertical_merge: Some(VerticalMergeState::Continue),
                vertical_align: CellVAlign::Top,
            }],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
            bidi_override: None,
        };
        let col_widths = vec![Pt::new(40.0)];
        let slices = layout_table_paginated(
            &[row0, row1],
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(30.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );

        // Should still page: the merge group doesn't fit on page 1, so it
        // moves intact to page 2 — never split mid-cell.
        let count0 = slices[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(count0, 0, "vMerge span must not split across pages");
    }

    /// Regression: a table whose every row has `tblHeader=1` (a Word template
    /// pattern, e.g. the trailing "Anhang" tables in the Volvo Annahme-Protokoll)
    /// must still respect `available_height`. The previous implementation
    /// unconditionally added every header row to the first slice, overflowing
    /// the page when the table arrived near the bottom of a stacked page.
    #[test]
    fn all_header_rows_paginate_when_exceeding_available() {
        let mut r0 = tall_row(2);
        r0.is_header = Some(true);
        r0.cant_split = Some(true);
        let mut r1 = tall_row(2);
        r1.is_header = Some(true);
        r1.cant_split = Some(true);
        let mut r2 = tall_row(2);
        r2.is_header = Some(true);
        r2.cant_split = Some(true);

        let rows = vec![r0, r1, r2];
        let col_widths = vec![Pt::new(40.0)];

        // Available 30pt fits at most one row; the remaining rows must spill.
        let slices = layout_table_paginated(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(30.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );

        assert!(
            slices.len() >= 2,
            "expected ≥2 slices, got {}",
            slices.len()
        );
        assert!(
            slices[0].size.height <= Pt::new(30.0),
            "slice 0 height {:.1}pt exceeds available 30pt — header rows \
             ignoring the fitting check",
            slices[0].size.height.raw()
        );

        // Each header row appears exactly once — no double-emission, since
        // an all-header table has no "subsequent" rows to head.
        let total_text: usize = slices
            .iter()
            .map(|s| {
                s.commands
                    .iter()
                    .filter(|c| matches!(c, DrawCommand::Text { .. }))
                    .count()
            })
            .sum();
        assert_eq!(total_text, 6, "all rows emitted exactly once");
    }

    /// Regression: when `available_height` is zero (cursor at the page
    /// bottom), an all-header table must produce an empty first slice and
    /// emit content on a fresh continuation page.
    #[test]
    fn all_header_rows_with_no_space_advances_page() {
        let mut r0 = tall_row(2);
        r0.is_header = Some(true);
        r0.cant_split = Some(true);

        let rows = vec![r0];
        let col_widths = vec![Pt::new(40.0)];

        let slices = layout_table_paginated(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::ZERO,
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );

        assert_eq!(slices.len(), 2, "empty first slice + content slice");
        let count0 = slices[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        let count1 = slices[1]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(count0, 0, "first slice empty when no space available");
        assert_eq!(count1, 2, "all content moves to continuation");
    }

    /// Regression: when a non-header row triggers a page break, header rows
    /// must still be re-emitted at the top of the continuation slice.
    #[test]
    fn header_repeats_on_continuation_when_body_overflows() {
        let mut header = tall_row(1); // 14pt
        header.is_header = Some(true);
        header.cant_split = Some(true);
        let body0 = tall_row(2); // 28pt
        let mut body1 = tall_row(2); // 28pt
        body1.cant_split = Some(true);

        let rows = vec![header, body0, body1];
        let col_widths = vec![Pt::new(40.0)];

        // Available 50pt: header (14) + body0 (28) = 42 fits; body1 (28) spills.
        let slices = layout_table_paginated(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(50.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );

        assert_eq!(slices.len(), 2);
        let count0 = slices[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        let count1 = slices[1]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(count0, 3, "slice 0: header (1) + body0 (2)");
        assert_eq!(count1, 3, "slice 1: header repeated (1) + body1 (2)");
    }

    #[test]
    fn continuation_slices_use_their_own_page_heights() {
        let rows = (0..4)
            .map(|_| {
                let mut row = tall_row(2);
                row.cant_split = Some(true);
                row
            })
            .collect::<Vec<_>>();
        let col_widths = vec![Pt::new(40.0)];

        let slices = layout_table_paginated_with_page_heights(
            &rows,
            &col_widths,
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            TablePaginationHeights {
                available_height: Pt::new(30.0),
                suppress_first_row_top: false,
                page_height_for_slice: |slice_index| match slice_index {
                    1 => Pt::new(30.0),
                    _ => Pt::new(60.0),
                },
            },
        );

        let row_counts = slices
            .iter()
            .map(|slice| {
                slice
                    .commands
                    .iter()
                    .filter(|command| matches!(command, DrawCommand::Text { .. }))
                    .count()
                    / 2
            })
            .collect::<Vec<_>>();
        assert_eq!(row_counts, vec![1, 1, 2]);
    }

    // ── §17.4.84 lone vMerge=Restart ─────────────────────────────────────

    /// A `Restart` cell with no `Continue` row under it is an ordinary cell.
    ///
    /// It used to fall through *both* height paths: `measure_table_rows`
    /// skipped every merged cell (deferring to the span calculation) and
    /// `expand_rows_for_vmerge` returned early because the "span" is one row.
    /// The row therefore got zero height while still emitting its content, so
    /// whatever followed the table drew on top of it.
    ///
    /// Asserted against the unmerged control rather than a literal, so the
    /// test states the actual rule — a lone restart behaves like no merge at
    /// all — instead of pinning today's line height.
    #[test]
    fn lone_vmerge_restart_row_gets_the_same_height_as_an_unmerged_cell() {
        let build = |vmerge: Option<VerticalMergeState>| {
            let mut row = tall_row(3);
            row.cells[0].vertical_merge = vmerge;
            layout_table(
                &[row],
                &[Pt::new(40.0)],
                Pt::ZERO,
                Pt::new(14.0),
                None,
                None,
                false,
            )
        };

        let lone_restart = build(Some(VerticalMergeState::Restart));
        let control = build(None);

        assert!(
            control.size.height > Pt::ZERO,
            "control must have real height for this test to mean anything"
        );
        assert_eq!(
            lone_restart.size.height, control.size.height,
            "a restart with nothing continuing below it is an ordinary cell"
        );
    }

    /// The companion guard: a *genuine* merge span must still take its height
    /// from `expand_rows_for_vmerge` across the whole span. Folding a
    /// `Restart` cell into its first row unconditionally — the naive fix for
    /// the case above — double-counts it against the rows below and inflates
    /// the table.
    #[test]
    fn genuine_vmerge_span_height_is_not_double_counted() {
        let mut restart = tall_row(6); // 6 lines ≈ 84pt of content
        restart.cells[0].vertical_merge = Some(VerticalMergeState::Restart);
        let mut cont = tall_row(0);
        cont.cells[0].vertical_merge = Some(VerticalMergeState::Continue);

        let result = layout_table(
            &[restart, cont],
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );

        // The span is exactly the restart cell's content: 6 lines × 14pt.
        assert!(
            (result.size.height.raw() - 84.0).abs() < 0.01,
            "merged span should total the restart cell's content height (84pt), got {}",
            result.size.height.raw()
        );
    }

    // ── §17.4.6 page advance must gain space ─────────────────────────────

    /// A row group taller than a whole page fits nowhere, so advancing to a
    /// fresh page cannot help — it only abandons an empty leading slice, which
    /// the section layer turns into a blank page.
    ///
    /// `available_height == page_height` models the table starting at the top
    /// of a page, which is exactly when the old code emitted the blank.
    #[test]
    fn oversized_group_at_page_top_does_not_emit_an_empty_leading_slice() {
        let mut row = tall_row(20); // ≈ 280pt
        row.cant_split = Some(true);

        let slices = layout_table_paginated(
            &[row],
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(100.0),
                page_height: Pt::new(100.0), // a fresh page is no roomier
                suppress_first_row_top: false,
            },
        );

        assert_eq!(slices.len(), 1, "no page to gain, so no page to push");
        assert!(
            !slices[0].commands.is_empty(),
            "the single slice must carry the content, not be an abandoned empty"
        );
    }

    /// The converse, so the fix can't be "never advance": when the table
    /// starts part-way down a page, a fresh page *is* roomier and the group
    /// must still move.
    #[test]
    fn oversized_group_still_advances_when_the_next_page_is_roomier() {
        let mut row = tall_row(5); // ≈ 70pt
        row.cant_split = Some(true);

        let slices = layout_table_paginated(
            &[row],
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(30.0), // little left here
                page_height: Pt::new(100.0),     // but a full page next
                suppress_first_row_top: false,
            },
        );

        assert_eq!(slices.len(), 2, "advancing gains 70pt, so it must advance");
        assert!(
            slices[0].commands.is_empty(),
            "leading slice is the advance"
        );
        assert!(!slices[1].commands.is_empty());
    }

    // ── In-cell paragraph splitting semantics (§17.3.1.14/.15/.44) ───────

    #[test]
    fn widow_control_keeps_two_cell_lines_on_each_side_of_a_split() {
        // §17.3.1.44: a 6-line cell paragraph where 5 lines would fit must not
        // strand a single-line widow — the split leaves 4 on page 1 and 2 on
        // page 2 (not 5 + 1).
        let rows = vec![one_cell_row(vec![styled_para(
            6,
            ParagraphStyle::default(),
        )])];
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(75.0), // 5 lines fit geometrically
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2);
        assert_eq!(slice_line_count(&slices[0]), 4, "orphan side keeps >= 2");
        assert_eq!(slice_line_count(&slices[1]), 2, "widow side keeps >= 2");
    }

    #[test]
    fn without_widow_control_a_cell_paragraph_may_strand_one_line() {
        // Same geometry, widow control off (§17.3.1.44 disabled) ⇒ the cell
        // splits as far as it fits: 5 lines on page 1, 1 on page 2. This is the
        // behaviour the widow-control test above suppresses.
        let style = ParagraphStyle {
            widow_control: false,
            ..Default::default()
        };
        let rows = vec![one_cell_row(vec![styled_para(6, style)])];
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(75.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2);
        assert_eq!(slice_line_count(&slices[0]), 5);
        assert_eq!(slice_line_count(&slices[1]), 1);
    }

    #[test]
    fn keep_lines_cell_paragraph_moves_whole_instead_of_splitting() {
        // §17.3.1.14: a keepLines paragraph is never divided. Only 3 of its 6
        // lines fit on page 1, but rather than split it moves whole to page 2
        // (which can hold all 6).
        let style = ParagraphStyle {
            keep_lines: true,
            ..Default::default()
        };
        let rows = vec![one_cell_row(vec![styled_para(6, style)])];
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(50.0), // 3 lines would fit
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2);
        assert_eq!(
            slice_line_count(&slices[0]),
            0,
            "keepLines: nothing splits off"
        );
        assert_eq!(slice_line_count(&slices[1]), 6, "whole paragraph moved");
    }

    #[test]
    fn cell_splits_at_paragraph_boundary_when_neither_paragraph_can_split() {
        // Two 3-line paragraphs: neither can split under widow control (a
        // 3-line paragraph has no ≥2/≥2 cut). With 4 lines of room the cell
        // must break at the *paragraph boundary* — 3 lines (para A) on page 1,
        // 3 (para B) on page 2 — not mid-paragraph.
        let rows = vec![one_cell_row(vec![
            styled_para(3, ParagraphStyle::default()),
            styled_para(3, ParagraphStyle::default()),
        ])];
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(60.0), // room for ~4 lines
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2);
        assert_eq!(slice_line_count(&slices[0]), 3, "para A stays intact");
        assert_eq!(slice_line_count(&slices[1]), 3, "para B stays intact");
    }

    #[test]
    fn keep_next_forbids_splitting_a_cell_at_that_paragraph_boundary() {
        // §17.3.1.15: para A is keepNext, so the cell may not break between A
        // and B. Neither 3-line paragraph can split internally (widow control),
        // so — unlike the boundary test above — the whole cell moves to page 2.
        let keep_next = ParagraphStyle {
            keep_next: true,
            ..Default::default()
        };
        let rows = vec![one_cell_row(vec![
            styled_para(3, keep_next),
            styled_para(3, ParagraphStyle::default()),
        ])];
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(60.0),
                page_height: Pt::new(200.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2);
        assert_eq!(
            slice_line_count(&slices[0]),
            0,
            "keepNext binds A to B: the boundary cut is illegal, cell moves whole"
        );
        assert_eq!(slice_line_count(&slices[1]), 6);
    }

    #[test]
    fn cell_paragraph_split_across_three_pages_never_widows() {
        // §17.3.1.44 must hold at *every* break, not just the first. An 11-line
        // cell paragraph over pages that each hold 5 lines splits 5/4/2 — never
        // the 5/5/1 a naive re-split would produce.
        let rows = vec![one_cell_row(vec![styled_para(
            11,
            ParagraphStyle::default(),
        )])];
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(75.0),
                page_height: Pt::new(75.0),
                suppress_first_row_top: false,
            },
        );
        assert!(
            slices.len() >= 3,
            "expected >= 3 slices, got {}",
            slices.len()
        );
        let counts: Vec<usize> = slices.iter().map(slice_line_count).collect();
        assert_eq!(counts.iter().sum::<usize>(), 11, "every line emitted once");
        assert!(
            counts.iter().all(|&c| c >= 2),
            "no single-line widow/orphan segment: {counts:?}"
        );
    }

    /// §17.4.45: the bottom-edge gap belongs to the table, not to the page it
    /// happens to land on. The monolithic path adds it (`cursor_y +
    /// cell_spacing`); the paginated path did not, so the *same table* kept or
    /// lost its bottom gap depending only on whether it fitted on one page.
    ///
    /// Asserted against absolute values and against the **monolithic** path,
    /// which is a genuinely separate code path. Comparing one paginated layout
    /// to another cannot see this: a mutation that drops the gap drops it from
    /// both sides and the comparison still holds.
    ///
    /// Empty cells make the arithmetic exact — each row's whole height is its
    /// own reserved leading gap (see
    /// `measure::tests::cell_spacing_separates_rows_vertically`), so two rows
    /// are `2 × spacing` of content plus one trailing gap.
    #[test]
    fn cell_spacing_bottom_edge_gap_survives_pagination() {
        let spacing = Pt::new(6.0);
        let rows = vec![one_cell_row(vec![]), one_cell_row(vec![])];
        let widths = [Pt::new(94.0)];

        let whole = layout_table(&rows, &widths, spacing, Pt::new(14.0), None, None, false);
        assert_eq!(
            whole.size.height,
            spacing * 3.0,
            "two leading gaps plus the table's own bottom edge"
        );

        let split = layout_table_paginated(
            &rows,
            &widths,
            spacing,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: spacing * 1.5,
                page_height: spacing * 1.5,
                suppress_first_row_top: false,
            },
        );
        assert_eq!(split.len(), 2, "precondition: the table paginated");
        assert_eq!(
            split[0].size.height, spacing,
            "an intermediate slice ends at a page cut, not at the table's edge"
        );
        assert_eq!(
            split[1].size.height,
            spacing * 2.0,
            "the final slice owns the trailing gap"
        );

        let total: Pt = split.iter().map(|s| s.size.height).sum();
        assert_eq!(
            total, whole.size.height,
            "a paginated table owns the same total height as the same table \
             laid out whole"
        );
    }

    /// §17.4.45: the width a table *occupies* is the slot sum **plus** one
    /// `tblCellSpacing`, because `build/table.rs::reserve_cell_spacing` already
    /// took that spacing out of the slots. Reported the same way by both paths
    /// and on every slice, so a caller placing the table never has to re-derive
    /// it — which is what `section/layout.rs` used to do, and got wrong.
    #[test]
    fn the_reported_width_is_the_outer_width_including_cell_spacing() {
        let spacing = Pt::new(6.0);
        let rows = vec![one_cell_row(vec![]), one_cell_row(vec![])];
        let widths = [Pt::new(94.0)];
        let outer = Pt::new(100.0);

        let whole = layout_table(&rows, &widths, spacing, Pt::new(14.0), None, None, false);
        assert_eq!(whole.size.width, outer, "94pt of slots + 6pt of spacing");

        let split = layout_table_paginated(
            &rows,
            &widths,
            spacing,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: spacing * 1.5,
                page_height: spacing * 1.5,
                suppress_first_row_top: false,
            },
        );
        assert_eq!(split.len(), 2, "precondition: the table paginated");
        for (i, slice) in split.iter().enumerate() {
            assert_eq!(slice.size.width, outer, "slice {i} spans the same width");
        }
    }

    // ── The paginator's lookahead must agree with the paginator ─────────────

    /// The lookahead and the group builder, over the same table.
    fn lookahead_and_first_group(
        rows: &[TableRowInput],
        borders: Option<&TableBorderConfig>,
    ) -> (Pt, Pt) {
        let cols = [Pt::new(40.0)];
        let measured =
            measure_table_rows(rows, &cols, Pt::ZERO, Pt::new(14.0), borders, None, false);
        let groups = build_row_groups(rows, &measured);
        let lookahead = measure_leading_table_group_height(
            rows,
            &cols,
            Pt::ZERO,
            Pt::new(14.0),
            borders,
            None,
            false,
        )
        .expect("a non-empty table has a leading group");
        (lookahead, groups[0].height)
    }

    /// `measure_leading_table_group_height` derives the height of the table's
    /// first atomic row group **independently** of `build_row_groups` — over a
    /// truncated table, of the group plus the one row that follows it, which is
    /// the least that can resolve the group's shared bottom border. Nothing
    /// asserted the two agreed.
    ///
    /// A divergence is silent and expensive: `section/layout.rs` asks the
    /// lookahead whether the group fits before a break, and then the paginator
    /// lays out a group of a different height. Every test in the suite would
    /// stay green while keepNext chains and page fits were decided against a
    /// number the table never has.
    ///
    /// Asserted for the three shapes `build_row_groups` distinguishes: a plain
    /// row (a group of one), a §17.4.84 vMerge span (a group of many), and a
    /// §17.4.6 `cantSplit` row. Each carries `insideH` borders, because it is
    /// `border_gap_below` — reserved from the *next* row's share of the shared
    /// edge — that the truncation is most able to get wrong. Without the
    /// following row the group's last row looks like the table's last row, its
    /// bottom falls back to `tblBorders/bottom` (unset here), and the reserved
    /// gap silently disappears.
    #[test]
    fn the_leading_group_lookahead_matches_the_paginators_own_first_group() {
        let inside_h = TableBorderConfig {
            top: None,
            bottom: None,
            left: None,
            right: None,
            inside_h: Some(TableBorderLine {
                width: Pt::new(2.0),
                color: RgbColor::BLACK,
                style: TableBorderStyle::Single,
            }),
            inside_v: None,
        };

        // A plain row: the group is one row, 14pt of content plus the 2pt gap
        // reserved below it for the `insideH` edge it owns.
        let plain: Vec<TableRowInput> = (0..3).map(|_| tall_row(1)).collect();
        let (lookahead, group) = lookahead_and_first_group(&plain, Some(&inside_h));
        assert_eq!(group, Pt::new(16.0), "14pt row + 2pt reserved band");
        assert_eq!(lookahead, group, "plain leading row");

        // A 3-row merge span, followed by a plain row. The group is the whole
        // span, so the lookahead has to measure four rows to answer about three.
        let mut restart = tall_row(4);
        restart.cells[0].vertical_merge = Some(VerticalMergeState::Restart);
        let mut cont1 = tall_row(0);
        cont1.cells[0].vertical_merge = Some(VerticalMergeState::Continue);
        let mut cont2 = tall_row(0);
        cont2.cells[0].vertical_merge = Some(VerticalMergeState::Continue);
        let span = vec![restart, cont1, cont2, tall_row(1)];
        let (lookahead, group) = lookahead_and_first_group(&span, Some(&inside_h));
        assert!(group > Pt::new(16.0), "a span is taller than one row");
        assert_eq!(lookahead, group, "3-row vMerge span");

        // §17.4.6 `cantSplit` — still a group of one, but one the paginator may
        // not divide, so the lookahead's number is the only thing standing
        // between it and an overflowing page.
        let mut atomic: Vec<TableRowInput> = (0..3).map(|_| tall_row(2)).collect();
        atomic[0].cant_split = Some(true);
        let (lookahead, group) = lookahead_and_first_group(&atomic, Some(&inside_h));
        assert_eq!(group, Pt::new(30.0), "two 14pt lines + the 2pt band");
        assert_eq!(lookahead, group, "cantSplit leading row");
    }

    // ── §17.4.49 repeated headers ───────────────────────────────────────────

    /// A row whose one cell holds one 14pt line per entry — each fragment is
    /// 30pt wide in a 40pt column, so no two share a line.
    fn labelled_row(lines: &[&str]) -> TableRowInput {
        one_cell_row(vec![LayoutBlock::Paragraph {
            fragments: lines.iter().map(|t| text_frag(t, 30.0)).collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }])
    }

    /// Every `Text` command as `(text, x, y)`.
    fn text_at(slice: &TableSlice) -> Vec<(String, f32, f32)> {
        slice
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, position, .. } => {
                    Some((text.to_string(), position.x.raw(), position.y.raw()))
                }
                _ => None,
            })
            .collect()
    }

    /// §17.4.49: a repeated header is the *same* row drawn again, so it must
    /// land at the same x and the same y on every slice — a continuation slice
    /// begins at its own origin, and the header is the first thing emitted into
    /// it.
    ///
    /// Three slices, so "every continuation" means more than one. The body row
    /// under the header must then start exactly one header-row height down, on
    /// each of them — taken from a monolithic layout of the header row alone
    /// rather than typed in, so the assertion is that the repeat *occupies* the
    /// height it measured.
    #[test]
    fn a_repeated_header_lands_at_the_same_place_on_every_slice() {
        let mut header = labelled_row(&["HDR "]);
        header.is_header = Some(true);
        header.cant_split = Some(true);
        let body = |t: &str| {
            let mut r = labelled_row(&[t, "x "]);
            r.cant_split = Some(true);
            r
        };
        let rows = vec![header, body("A0 "), body("B0 "), body("C0 ")];

        let header_only = layout_table(
            &rows[..1],
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            false,
        );
        assert_eq!(header_only.size.height, Pt::new(14.0));

        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
            Pt::ZERO,
            Pt::new(14.0),
            None,
            None,
            &TablePaginationConfig {
                available_height: Pt::new(45.0),
                page_height: Pt::new(50.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 3, "header + one 28pt body row per page");

        let mut header_positions = Vec::new();
        for (i, slice) in slices.iter().enumerate() {
            let texts = text_at(slice);
            let hdr: Vec<_> = texts.iter().filter(|t| t.0 == "HDR ").collect();
            assert_eq!(hdr.len(), 1, "slice {i}: the header appears exactly once");
            header_positions.push((hdr[0].1, hdr[0].2));

            let body_top = texts
                .iter()
                .filter(|t| t.0 != "HDR ")
                .map(|t| t.2)
                .fold(f32::INFINITY, f32::min);
            assert_eq!(
                body_top - hdr[0].2,
                header_only.size.height.raw(),
                "slice {i}: the body starts one header-row height below the header"
            );
        }
        assert_eq!(
            header_positions,
            vec![header_positions[0]; 3],
            "the repeat is the same row at the same place, not a re-flow"
        );

        // …and the bodies themselves are neither lost nor repeated.
        let bodies: Vec<String> = slices
            .iter()
            .flat_map(|s| text_at(s).into_iter().map(|t| t.0))
            .filter(|t| t.ends_with("0 "))
            .filter(|t| t != "HDR ")
            .collect();
        assert_eq!(bodies, vec!["A0 ", "B0 ", "C0 "]);
    }

    /// §17.4.49 against a page too short for the header itself, where
    /// `remaining -= header_height` goes negative before a single body row is
    /// placed and `next_page_height - header_height` is negative too.
    ///
    /// The arithmetic is allowed to go negative — there is genuinely less than
    /// no room — but two things must survive it: the paginator terminates, and
    /// no body row is lost or drawn twice. Swept across page heights either
    /// side of the header's own, so the properties are asserted where the
    /// header fits, where it exactly fills a page, and where it does not fit at
    /// all.
    #[test]
    fn header_repetition_conserves_body_content_at_every_page_height() {
        let mut header = labelled_row(&["H0 ", "H1 ", "H2 ", "H3 ", "H4 "]); // 70pt
        header.is_header = Some(true);
        header.cant_split = Some(true);
        let body = |t: &str| {
            let mut r = labelled_row(&[t, "x "]); // 28pt
            r.cant_split = Some(true);
            r
        };
        let rows = vec![header, body("A0 "), body("B0 "), body("C0 ")];

        for page in [20.0f32, 50.0, 70.0, 71.0, 100.0, 200.0] {
            let slices = layout_table_paginated(
                &rows,
                &[Pt::new(40.0)],
                Pt::ZERO,
                Pt::new(14.0),
                None,
                None,
                &TablePaginationConfig {
                    available_height: Pt::new(100.0),
                    page_height: Pt::new(page),
                    suppress_first_row_top: false,
                },
            );

            let all: Vec<String> = slices
                .iter()
                .flat_map(|s| text_at(s).into_iter().map(|t| t.0))
                .collect();
            let bodies: Vec<&String> = all
                .iter()
                .filter(|t| t.starts_with(['A', 'B', 'C']))
                .collect();
            assert_eq!(
                bodies,
                vec!["A0 ", "B0 ", "C0 "],
                "page={page}: every body row exactly once, in order"
            );
            // The header's first line appears once per slice — never twice on
            // one, never missing from a continuation.
            let headers_per_slice: Vec<usize> = slices
                .iter()
                .map(|s| text_at(s).iter().filter(|t| t.0 == "H0 ").count())
                .collect();
            assert_eq!(
                headers_per_slice,
                vec![1; slices.len()],
                "page={page}: one header per slice"
            );
        }
    }

    // ── Slice assembly ──────────────────────────────────────────────────────

    /// The shading → content → borders layering is a property of *every* slice,
    /// not only of a table that fits on one page: `SliceBuilder::finish`
    /// concatenates the three buffers once per slice, so a paginated table gets
    /// the guarantee three times over rather than once at the end.
    ///
    /// Pinned on the final command list, which is what the painter walks. Were
    /// the order to come out per row instead of per slice, a cell's background
    /// would be painted over the previous row's shared bottom edge.
    #[test]
    fn every_slice_layers_shading_then_content_then_borders() {
        const GREY: RgbColor = RgbColor {
            r: 200,
            g: 200,
            b: 200,
        };
        const RED: RgbColor = RgbColor { r: 255, g: 0, b: 0 };
        let line = TableBorderLine {
            width: Pt::new(1.0),
            color: RED,
            style: TableBorderStyle::Single,
        };
        let rows: Vec<TableRowInput> = (0..4)
            .map(|i| {
                let mut r = labelled_row(&[if i % 2 == 0 { "even " } else { "odd " }]);
                r.cells[0].shading = Some(GREY);
                r.cant_split = Some(true);
                r
            })
            .collect();

        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(40.0)],
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
            &TablePaginationConfig {
                available_height: Pt::new(40.0),
                page_height: Pt::new(40.0),
                suppress_first_row_top: false,
            },
        );
        assert!(slices.len() >= 2, "got {} slices", slices.len());

        for (i, slice) in slices.iter().enumerate() {
            let find = |pred: &dyn Fn(&DrawCommand) -> bool| slice.commands.iter().position(pred);
            let shading = find(&|c| matches!(c, DrawCommand::Rect { color, .. } if *color == GREY))
                .unwrap_or_else(|| panic!("slice {i}: no shading"));
            let text = find(&|c| matches!(c, DrawCommand::Text { .. }))
                .unwrap_or_else(|| panic!("slice {i}: no content"));
            let border = find(&|c| matches!(c, DrawCommand::Rect { color, .. } if *color == RED))
                .unwrap_or_else(|| panic!("slice {i}: no borders"));
            assert!(
                shading < text && text < border,
                "slice {i}: expected shading({shading}) < content({text}) < borders({border})"
            );
        }
    }

    /// §17.4.45 / issue #168: the outer rectangle of a spaced table, at exact
    /// coordinates on both slices of a two-page split.
    ///
    /// `carries_top` and `carries_bottom` are what the flags mean geometrically.
    /// The first slice holds the table's top edge and not its bottom, the last
    /// the reverse, and each slice's verticals run the whole way to whichever
    /// side is a page cut rather than a table edge — inset under a horizontal
    /// only where one exists. That last part is why the two slices' verticals
    /// differ in *both* origin and length rather than just in length.
    ///
    /// One 100pt slot at a 20pt spacing: the table is 120pt wide, each row box
    /// is one 14pt line plus the horizontal borders inside it plus its own 20pt
    /// leading gap, and only the last slice adds the trailing gap at the
    /// table's bottom edge — 71pt against 91pt. The cells sit at x = 20..100,
    /// so a rect touching x = 0 or x = 120 can only belong to the table's own
    /// rectangle.
    ///
    /// The four rows are not all the same height, and that is the spacing's
    /// doing: the table's own top and bottom belong to the outline, so row 0
    /// carries only its `insideH` bottom and row 3 only its `insideH` top — 35pt
    /// each — while the two interior rows carry both and are 36pt. A slice is
    /// therefore 35 + 36, not twice anything.
    #[test]
    fn a_paginated_spaced_tables_outline_is_exact_on_each_slice() {
        let line = TableBorderLine {
            width: Pt::new(1.0),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        let rows: Vec<TableRowInput> = (0..4).map(|_| labelled_row(&["r "])).collect();
        let slices = layout_table_paginated(
            &rows,
            &[Pt::new(100.0)],
            Pt::new(20.0),
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
            &TablePaginationConfig {
                available_height: Pt::new(90.0),
                page_height: Pt::new(90.0),
                suppress_first_row_top: false,
            },
        );
        assert_eq!(slices.len(), 2, "a 35pt and a 36pt row box per 90pt page");

        // The table's own rectangle: every rect that touches its left or right
        // bound. A cell's edges live between x = 20 and x = 100.
        let outline = |slice: &TableSlice| -> Vec<(f32, f32, f32, f32)> {
            slice
                .commands
                .iter()
                .filter_map(|c| match c {
                    DrawCommand::Rect { rect, .. } => {
                        let (x, w) = (rect.origin.x.raw(), rect.size.width.raw());
                        (x == 0.0 || x + w == 120.0).then_some((
                            x,
                            rect.origin.y.raw(),
                            w,
                            rect.size.height.raw(),
                        ))
                    }
                    _ => None,
                })
                .collect()
        };

        assert_eq!(slices[0].size, PtSize::new(Pt::new(120.0), Pt::new(71.0)));
        assert_eq!(
            outline(&slices[0]),
            vec![
                (0.0, 0.0, 120.0, 1.0), // the table's top edge, on the first slice only
                (0.0, 1.0, 1.0, 70.0),  // left, inset under the top, running to the cut
                (119.0, 1.0, 1.0, 70.0),
            ],
        );

        assert_eq!(slices[1].size, PtSize::new(Pt::new(120.0), Pt::new(91.0)));
        assert_eq!(
            outline(&slices[1]),
            vec![
                (0.0, 90.0, 120.0, 1.0), // the bottom edge, on the last slice only
                (0.0, 0.0, 1.0, 90.0),   // left, from the cut down to that bottom
                (119.0, 0.0, 1.0, 90.0),
            ],
        );
    }
}
