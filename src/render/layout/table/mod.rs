//! Table layout — 3-pass column sizing, cell layout, border rendering.
//!
//! Pass 1: Compute column widths from grid definitions or equal distribution.
//! Pass 2: Lay out each cell with tight width constraints, determine row heights.
//! Pass 3: Position cells and emit border commands.

use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;

use super::BoxConstraints;

mod borders;
mod emit;
mod grid;
mod measure;
mod split;
mod types;

pub use grid::compute_column_widths;
pub use types::*;

use emit::{emit_split_row, emit_table_rows, TableCommandBuffers};
use grid::build_row_groups;
use measure::measure_table_rows;
use split::{find_row_cut, split_row_at, RowCutInput};

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
    pub commands: Vec<super::draw_command::DrawCommand>,
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
        // §17.4.1 split path.
        let is_header = group.start < header_count;

        // Doesn't fit. Try to split (§17.4.1) before spilling the whole
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
                remaining = page_height;
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
                            remaining = page_height;
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

        // No split possible — move the whole group to the next page.
        slices.push(std::mem::take(&mut current_slice));
        remaining = page_height;
        // §17.4.49: prepend the repeating header rows only when this group
        // sits past the headers. When advancing because a header row itself
        // doesn't fit, the row is part of the table's first appearance —
        // emitting `Range(0..header_count)` here would duplicate it.
        if !is_header && header_count > 0 {
            current_slice.push(SliceItem::Range(0..header_count));
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
        current_slice.push(SliceItem::Range(group.start..group.end));
        remaining -= group.height;
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
        .map(|(slice_idx, items)| {
            let mut commands = Vec::new();
            let mut content_commands = Vec::new();
            let mut border_commands = Vec::new();
            let mut cursor_y = Pt::ZERO;
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
                match item {
                    SliceItem::Range(range) => {
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
                    SliceItem::Split { row_idx, mr } | SliceItem::Continuation { row_idx, mr } => {
                        let has_next = item_idx + 1 < items.len();
                        emit_split_row(
                            mr,
                            &rows[*row_idx],
                            &mut cursor_y,
                            &mut TableCommandBuffers {
                                commands: &mut commands,
                                content_commands: &mut content_commands,
                                border_commands: &mut border_commands,
                            },
                            top_override,
                            has_next,
                        );
                    }
                }
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
    use crate::render::geometry::PtEdgeInsets;
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
    use crate::render::resolve::color::RgbColor;
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
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            },
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
            grid_before: 0,
            grid_after: 0,
            border_overrides: None,
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
                grid_before: 0,
                grid_after: 0,
                border_overrides: None,
            },
            TableRowInput {
                cells: vec![simple_cell("c"), simple_cell("d")],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                grid_after: 0,
                border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
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
            grid_before: 0,
            grid_after: 0,
            border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
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
        // §17.4.17: gridBefore=1 + wBefore skips the first grid column. With
        // a 4-column grid [10, 100, 200, 10] and a row with gridBefore=1 and
        // gridAfter=1, the two cells must occupy columns 1 and 2 — not 0 and 1.
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("A"), simple_cell("B")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 1,
            grid_after: 1,
            border_overrides: None,
        }];
        let col_widths = vec![Pt::new(10.0), Pt::new(100.0), Pt::new(200.0), Pt::new(10.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
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

    #[test]
    fn grid_after_does_not_overflow() {
        // §17.4.16: gridAfter=2 leaves the rightmost two columns of a 4-column
        // grid empty for this row. Cells must fit within the leftmost two
        // columns and not overflow into the gridAfter region.
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("X"), simple_cell("Y")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            grid_after: 2,
            border_overrides: None,
        }];
        let col_widths = vec![Pt::new(10.0), Pt::new(100.0), Pt::new(200.0), Pt::new(10.0)];
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            None,
            None,
            false,
        );

        let xs = text_x_positions(&result.commands);
        assert_eq!(xs.len(), 2);
        assert_eq!(xs[0], 0.0, "first cell at col 0");
        assert_eq!(xs[1], 10.0, "second cell at col 1");
        // The cells together occupy 10 + 100 = 110pt; the right 210pt are
        // empty (gridAfter region). Total table width stays at 320pt.
        assert_eq!(result.size.width.raw(), 320.0);
    }

    #[test]
    fn grid_before_first_cell_uses_inside_v_left_border() {
        // §17.4.17 + §17.4.38: with gridBefore>0, the row's first cell is not
        // at the table's left edge, so its left border must come from
        // `inside_v`, not `left`. Mirror for gridAfter>0 and the right edge.
        //
        // We use distinct border widths to identify which table border was
        // applied: `left`/`right` = 4pt, `inside_v` = 1pt. With grid_before=1
        // and grid_after=1 in a 4-column grid, both cells must use only the
        // 1pt borders for left/right edges; no 4pt border rect should appear.
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("A"), simple_cell("B")],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 1,
            grid_after: 1,
            border_overrides: None,
        }];
        let col_widths = vec![Pt::new(10.0), Pt::new(50.0), Pt::new(50.0), Pt::new(10.0)];
        let borders = TableBorderConfig {
            top: None,
            bottom: None,
            left: Some(TableBorderLine {
                width: Pt::new(4.0),
                color: RgbColor::BLACK,
                style: TableBorderStyle::Single,
            }),
            right: Some(TableBorderLine {
                width: Pt::new(4.0),
                color: RgbColor::BLACK,
                style: TableBorderStyle::Single,
            }),
            inside_h: None,
            inside_v: Some(TableBorderLine {
                width: Pt::new(1.0),
                color: RgbColor::BLACK,
                style: TableBorderStyle::Single,
            }),
        };
        let result = layout_table(
            &rows,
            &col_widths,
            &body_constraints(),
            Pt::new(14.0),
            Some(&borders),
            None,
            false,
        );

        // Vertical border rects have width equal to the border thickness
        // (depth) and height >= 1. Find any 4pt-thick border rect.
        let has_thick_vertical = result.commands.iter().any(|c| match c {
            DrawCommand::Rect { rect, color } if *color == RgbColor::BLACK => {
                rect.size.width.raw() == 4.0 && rect.size.height.raw() > 1.0
            }
            _ => false,
        });
        assert!(
            !has_thick_vertical,
            "no 4pt-thick vertical border should appear: gridBefore/gridAfter \
             mean cells aren't at the table's left/right edges, so left/right \
             borders are not applied"
        );

        // Inside_v (1pt) should appear at the boundary between cell A and cell B.
        let has_inside_v = result.commands.iter().any(|c| match c {
            DrawCommand::Rect { rect, color } if *color == RgbColor::BLACK => {
                rect.size.width.raw() == 1.0
            }
            _ => false,
        });
        assert!(
            has_inside_v,
            "1pt inside_v border between cells must appear"
        );
    }

    #[test]
    fn vmerge_across_rows_with_different_grid_before() {
        // §17.4.85 + §17.4.17: a cell at grid_col 1 in row A (gridBefore=1)
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
            grid_after: 0,
            border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
        };
        let col_widths = vec![Pt::new(50.0), Pt::new(100.0), Pt::new(150.0)];
        let result = layout_table(
            &[row_a, row_b],
            &col_widths,
            &body_constraints(),
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
            grid_after: 0,
            border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
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

    #[test]
    fn valign_bottom_on_vmerge_restart_uses_span_height() {
        // §17.4.85 + §17.4.84: vAlign on a vMerge=Restart cell should
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
            grid_after: 0,
            border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
        };
        let col_widths = vec![Pt::new(100.0), Pt::new(100.0)];
        let result = layout_table(
            &[row0, row1],
            &col_widths,
            &body_constraints(),
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

        // "header" is top-aligned in row 0; "Total" should be bottom-
        // aligned across the 2-row span, so roughly one row-height below.
        assert!(
            total_y > header_y + 10.0,
            "Total (bottom-valigned merged) should sit well below header (top-valigned row 0): \
             total_y={total_y}, header_y={header_y}"
        );
    }

    // ── Row splitting (§17.4.1) ──────────────────────────────────────────

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
            grid_after: 0,
            border_overrides: None,
        }
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
            &body_constraints(),
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
            &body_constraints(),
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
            &body_constraints(),
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
            grid_after: 0,
            border_overrides: None,
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
            grid_after: 0,
            border_overrides: None,
        };
        let col_widths = vec![Pt::new(40.0)];
        let slices = layout_table_paginated(
            &[row0, row1],
            &col_widths,
            &body_constraints(),
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
            &body_constraints(),
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
            &body_constraints(),
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
            &body_constraints(),
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
}
