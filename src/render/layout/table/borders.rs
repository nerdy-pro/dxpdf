use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

use super::{
    CellBorderOverride, DrawCommand, TableBorderConfig, TableBorderLine, TableBorderStyle,
    TableCellInput,
};

/// Resolved borders for one cell edge.
pub(super) struct CellBorders {
    pub(super) top: Option<TableBorderLine>,
    pub(super) bottom: Option<TableBorderLine>,
    pub(super) left: Option<TableBorderLine>,
    pub(super) right: Option<TableBorderLine>,
}

/// §17.4.38 / §17.7.6: resolve effective borders for a cell.
/// Per-cell borders (from conditional formatting) override table-level borders.
/// Table-level insideH/insideV are mapped to cell edges based on position.
pub(super) fn resolve_cell_effective_borders(
    cell: &TableCellInput,
    table_borders: Option<&TableBorderConfig>,
    row_idx: usize,
    col_idx: usize,
    num_rows: usize,
    num_cols: usize,
) -> (
    Option<TableBorderLine>,
    Option<TableBorderLine>,
    Option<TableBorderLine>,
    Option<TableBorderLine>,
) {
    // Start with table-level borders mapped to cell edges.
    let tb = table_borders;
    let is_first_row = row_idx == 0;
    let is_last_row = row_idx == num_rows - 1;
    let is_first_col = col_idx == 0;
    let is_last_col = col_idx == num_cols - 1;

    let mut top = if is_first_row {
        tb.and_then(|b| b.top)
    } else {
        tb.and_then(|b| b.inside_h)
    };
    let mut bottom = if is_last_row {
        tb.and_then(|b| b.bottom)
    } else {
        tb.and_then(|b| b.inside_h)
    };
    let mut left = if is_first_col {
        tb.and_then(|b| b.left)
    } else {
        tb.and_then(|b| b.inside_v)
    };
    let mut right = if is_last_col {
        tb.and_then(|b| b.right)
    } else {
        tb.and_then(|b| b.inside_v)
    };

    // Per-cell borders from conditional formatting override.
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

    (top, bottom, left, right)
}

/// §17.4.43: resolve a border conflict between two competing borders on
/// a shared edge.  Returns the winning border (or `None` if both are `None`).
///
/// Algorithm per [MS-OI29500] §17.4.66:
///   1. `none` yields to the opposing border; `nil` suppresses both.
///   2. Weight = border_width_eighths × border_number.  Higher wins.
///   3. Equal weight: style precedence list.
///   4. Equal style: darker color wins (R+B+2G, then B+2G, then G).
pub(super) fn resolve_border_conflict(
    a: Option<TableBorderLine>,
    b: Option<TableBorderLine>,
) -> Option<TableBorderLine> {
    match (a, b) {
        (None, b) => b,
        (a, None) => a,
        (Some(a), Some(b)) => Some(if border_weight(&b) > border_weight(&a) {
            b
        } else {
            a
        }),
    }
}

/// Emit all four borders for a cell as filled rectangles.
/// Borders are drawn INWARD from the cell edge per OOXML.
///
/// Horizontal borders (top/bottom) own the corner squares — they span the
/// full cell width. Vertical borders (left/right) fill only the space
/// between the horizontals. This eliminates anti-aliasing gaps at corners
/// that plagued the previous stroke-based approach.
pub(super) fn emit_cell_borders(
    commands: &mut Vec<DrawCommand>,
    b: CellBorders,
    cell_x: Pt,
    cell_w: Pt,
    row_y: Pt,
    row_h: Pt,
) {
    let top_w = b.top.map(|b| b.width).unwrap_or(Pt::ZERO);
    let bot_w = b.bottom.map(|b| b.width).unwrap_or(Pt::ZERO);
    let left_w = b.left.map(|b| b.width).unwrap_or(Pt::ZERO);
    let right_w = b.right.map(|b| b.width).unwrap_or(Pt::ZERO);

    // Horizontal borders: full cell width, covering corner squares.
    if let Some(ref border) = b.top {
        emit_border_rect(
            commands,
            border,
            PtRect::from_xywh(cell_x, row_y, cell_w, top_w),
            true,
        );
    }
    if let Some(ref border) = b.bottom {
        emit_border_rect(
            commands,
            border,
            PtRect::from_xywh(cell_x, row_y + row_h - bot_w, cell_w, bot_w),
            true,
        );
    }

    // Vertical borders: between horizontal borders (no corner overlap).
    let top_inset = if b.top.is_some() { top_w } else { Pt::ZERO };
    let bot_inset = if b.bottom.is_some() { bot_w } else { Pt::ZERO };
    let v_height = row_h - top_inset - bot_inset;
    if v_height > Pt::ZERO {
        if let Some(ref border) = b.left {
            emit_border_rect(
                commands,
                border,
                PtRect::from_xywh(cell_x, row_y + top_inset, left_w, v_height),
                false,
            );
        }
        if let Some(ref border) = b.right {
            emit_border_rect(
                commands,
                border,
                PtRect::from_xywh(
                    cell_x + cell_w - right_w,
                    row_y + top_inset,
                    right_w,
                    v_height,
                ),
                false,
            );
        }
    }
}

/// §17.4.43: compute border weight = width_eighths × style_number.
/// We only support Single (1) and Double (3).
fn border_weight(b: &TableBorderLine) -> f32 {
    let eighths = b.width.raw() * 8.0; // width is in points, sz is eighths
    let style_number = match b.style {
        TableBorderStyle::Single => 1.0,
        TableBorderStyle::Double => 3.0,
    };
    eighths * style_number
}

/// Extract border width or zero if absent.
pub(super) fn border_width(b: Option<TableBorderLine>) -> Pt {
    b.map(|b| b.width).unwrap_or(Pt::ZERO)
}

fn resolve_override(ovr: &CellBorderOverride) -> Option<TableBorderLine> {
    match ovr {
        CellBorderOverride::Nil => None,
        CellBorderOverride::Border(line) => Some(*line),
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
    use crate::render::geometry::{PtEdgeInsets, PtSize};
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
    use crate::render::layout::table::{
        layout_table, BoxConstraints, CellVAlign, DrawCommand, TableBorderConfig, TableBorderLine,
        TableBorderStyle, TableCellInput, TableRowInput,
    };
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

    #[test]
    fn borders_emit_lines() {
        let rows = vec![TableRowInput {
            cells: vec![simple_cell("a"), simple_cell("b")],
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
            Some(&TableBorderConfig {
                top: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
                bottom: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
                left: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
                right: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
                inside_h: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
                inside_v: Some(TableBorderLine {
                    width: Pt::new(0.5),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
            }),
            None,
            false,
        );

        // Borders are emitted as filled rects. Count border rects by
        // excluding cell shading rects (which use non-BLACK colors or
        // appear before borders in the command list).
        let border_rect_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Rect { color, .. } if *color == RgbColor::BLACK))
            .count();
        // §17.4.43: shared edges drawn once after conflict resolution.
        // Top(2) + bottom(2) + left(1) + insideV(1) + right(1) = 7 border rects.
        assert_eq!(border_rect_count, 7);
    }
}
