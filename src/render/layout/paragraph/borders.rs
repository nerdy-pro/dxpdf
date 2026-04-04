//! Paragraph border and shading rendering.

use super::super::draw_command::DrawCommand;
use super::super::BoxConstraints;
use super::line_emit::resolve_line_height;
use super::types::ParagraphStyle;
use crate::render::dimension::Pt;
use crate::render::geometry::PtOffset;

/// Emit paragraph border and shading commands, then advance the cursor for
/// border bottom space, `space_after`, and the empty-paragraph minimum height.
///
/// Returns the updated `cursor_y` after all spacing has been applied.
pub(super) fn emit_paragraph_borders_and_shading(
    commands: &mut Vec<DrawCommand>,
    style: &ParagraphStyle,
    constraints: &BoxConstraints,
    cursor_y: Pt,
    default_line_height: Pt,
    no_lines: bool,
) -> Pt {
    // §17.3.1.24: paragraph border and shading coordinate system.
    // Borders sit at the paragraph indent edges. The border `space` is the
    // distance between the border line and the text content. Top/bottom
    // border space expands the bordered area vertically.
    let border_space_top = style
        .borders
        .as_ref()
        .and_then(|b| b.top.as_ref())
        .map(|b| b.space)
        .unwrap_or(Pt::ZERO);
    let border_space_bottom = style
        .borders
        .as_ref()
        .and_then(|b| b.bottom.as_ref())
        .map(|b| b.space)
        .unwrap_or(Pt::ZERO);
    let para_left = style.indent_left;
    let para_right = constraints.max_width - style.indent_right;
    let para_top = style.space_before - border_space_top;
    let para_bottom = cursor_y + border_space_bottom;

    // §17.3.1.31: render paragraph shading (fills the border area).
    if let Some(bg_color) = style.shading {
        commands.insert(
            0,
            DrawCommand::Rect {
                rect: crate::render::geometry::PtRect::from_xywh(
                    para_left,
                    para_top,
                    para_right - para_left,
                    para_bottom - para_top,
                ),
                color: bg_color,
            },
        );
    }

    // §17.3.1.24: render paragraph borders at the indent edges.
    if let Some(ref borders) = style.borders {
        if let Some(ref top) = borders.top {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, para_top),
                    PtOffset::new(para_right, para_top),
                ),
                color: top.color,
                width: top.width,
            });
        }
        if let Some(ref bottom) = borders.bottom {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, para_bottom),
                    PtOffset::new(para_right, para_bottom),
                ),
                color: bottom.color,
                width: bottom.width,
            });
        }
        if let Some(ref left) = borders.left {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, para_top),
                    PtOffset::new(para_left, para_bottom),
                ),
                color: left.color,
                width: left.width,
            });
        }
        if let Some(ref right) = borders.right {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(para_right, para_top),
                    PtOffset::new(para_right, para_bottom),
                ),
                color: right.color,
                width: right.width,
            });
        }
    }

    // §17.3.1.24: bottom border space adds to paragraph height.
    let mut cursor_y = cursor_y + border_space_bottom + style.space_after;

    // If no lines, still consume default height + spacing.
    // Apply the paragraph's line spacing rule to the default line height.
    if no_lines {
        let line_h = resolve_line_height(
            default_line_height,
            default_line_height,
            &style.line_spacing,
        );
        cursor_y = style.space_before + line_h + style.space_after;
    }

    cursor_y
}
