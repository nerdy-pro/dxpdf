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
mod types;

pub use grid::compute_column_widths;
pub use types::*;

use emit::{emit_table_rows, TableCommandBuffers};
use grid::build_row_groups;
use measure::measure_table_rows;

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

    // Pack groups into page slices.
    // Each slice is a Vec<Range<usize>> of row ranges to emit.
    let mut slices: Vec<Vec<std::ops::Range<usize>>> = Vec::new();
    let mut current_slice: Vec<std::ops::Range<usize>> = Vec::new();
    let mut remaining = available_height;
    let is_first_slice = |slices: &Vec<Vec<_>>| slices.is_empty();

    for group in &groups {
        // Header rows on the first slice are emitted as normal rows.
        if is_first_slice(&slices) && group.start < header_count {
            current_slice.push(group.start..group.end);
            remaining -= group.height;
            continue;
        }

        if group.height <= remaining {
            current_slice.push(group.start..group.end);
            remaining -= group.height;
        } else {
            // Close current slice.
            slices.push(std::mem::take(&mut current_slice));
            // New page: start with header rows (if any).
            remaining = page_height;
            if header_count > 0 {
                current_slice.push(0..header_count);
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
            current_slice.push(group.start..group.end);
            remaining -= group.height;
        }
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
        .map(|(slice_idx, row_ranges)| {
            let mut commands = Vec::new();
            let mut content_commands = Vec::new();
            let mut border_commands = Vec::new();
            let mut cursor_y = Pt::ZERO;
            for (range_idx, range) in row_ranges.iter().enumerate() {
                // The first range on each continuation slice (slice_idx > 0) needs
                // a top border override, since internal conflict resolution removed
                // it. On the first slice, only apply if suppress_first_row_top was
                // used (adjacent table collapse) — in that case the first range
                // should NOT get the override (the whole point is to suppress it).
                let top_override = if slice_idx > 0 && range_idx == 0 {
                    outer_top_border
                } else {
                    None
                };
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
            commands.append(&mut content_commands);
            commands.append(&mut border_commands);
            TableSlice {
                commands,
                size: PtSize::new(measured.table_width, cursor_y),
            }
        })
        .collect()
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
            },
            TableRowInput {
                cells: vec![simple_cell("c"), simple_cell("d")],
                height_rule: None,
                is_header: None,
                cant_split: None,
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
}
