//! Paragraph border and shading rendering.

use super::super::draw_command::DrawCommand;
use super::super::BoxConstraints;
use super::types::ParagraphStyle;
use crate::render::dimension::Pt;
use crate::render::geometry::PtOffset;

/// §17.3.1.24: which page segment of a paragraph is being emitted, i.e. which
/// horizontal border edges and vertical spacing it owns. Side borders and
/// shading span every segment; the top edge (and space above) belongs to the
/// first segment, the bottom edge (and `space_after`) to the last.
#[derive(Clone, Copy)]
pub(super) enum SegmentEdges {
    /// A whole, un-split paragraph — both edges.
    Whole,
    /// First of several segments — top edge only.
    First,
    /// Interior segment — neither edge.
    Middle,
    /// Final segment — bottom edge only.
    Last,
}

impl SegmentEdges {
    /// The segment kind for a `[first, last)` range flagged first/last.
    pub(super) fn for_segment(is_first: bool, is_last: bool) -> Self {
        match (is_first, is_last) {
            (true, true) => Self::Whole,
            (true, false) => Self::First,
            (false, false) => Self::Middle,
            (false, true) => Self::Last,
        }
    }

    fn draw_top(self) -> bool {
        matches!(self, Self::Whole | Self::First)
    }

    fn draw_bottom(self) -> bool {
        matches!(self, Self::Whole | Self::Last)
    }
}

/// Emit border and shading for a whole (un-split) paragraph, then advance the
/// cursor for bottom border space, `space_after`, and the empty-paragraph
/// minimum height. A thin wrapper over [`emit_segment_borders_and_shading`]
/// with both edges drawn.
pub(super) fn emit_paragraph_borders_and_shading(
    commands: &mut Vec<DrawCommand>,
    style: &ParagraphStyle,
    constraints: &BoxConstraints,
    cursor_y: Pt,
    default_line_height: Pt,
    no_lines: bool,
) -> Pt {
    emit_segment_borders_and_shading(
        commands,
        style,
        constraints,
        cursor_y,
        default_line_height,
        no_lines,
        SegmentEdges::Whole,
    )
}

/// §17.3.1.24/§17.3.1.31: emit border and shading for one page segment of a
/// paragraph, returning the cursor after the segment's trailing spacing.
///
/// `content_bottom` is the y after the segment's last line. `edges` selects
/// which horizontal border edges this segment owns (see [`SegmentEdges`]): side
/// borders and shading span every segment, so a paragraph split across pages
/// keeps a continuous box down each page, while the top/bottom edges and their
/// spacing land only on the first/last segment. `SegmentEdges::Whole`
/// reproduces the un-split geometry exactly.
pub(super) fn emit_segment_borders_and_shading(
    commands: &mut Vec<DrawCommand>,
    style: &ParagraphStyle,
    constraints: &BoxConstraints,
    content_bottom: Pt,
    default_line_height: Pt,
    no_lines: bool,
    edges: SegmentEdges,
) -> Pt {
    let draw_top = edges.draw_top();
    let draw_bottom = edges.draw_bottom();
    // §17.3.1.24: the border `space` is the distance between the border line and
    // the text content; top/bottom border space expands the box vertically.
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
    // The box top sits above the first line by the top border space on the first
    // segment; a continuation segment's box starts flush at the segment top.
    let box_top = if draw_top {
        style.space_before - border_space_top
    } else {
        Pt::ZERO
    };
    let box_bottom = if draw_bottom {
        content_bottom + border_space_bottom
    } else {
        content_bottom
    };

    // §17.3.1.31: render paragraph shading (fills the border area).
    if let Some(bg_color) = style.shading {
        commands.insert(
            0,
            DrawCommand::Rect {
                rect: crate::render::geometry::PtRect::from_xywh(
                    para_left,
                    box_top,
                    para_right - para_left,
                    box_bottom - box_top,
                ),
                color: bg_color,
            },
        );
    }

    // §17.3.1.24: render paragraph borders at the indent edges. Top/bottom only
    // on the owning segment; sides on every segment.
    if let Some(ref borders) = style.borders {
        if draw_top {
            if let Some(ref top) = borders.top {
                commands.push(DrawCommand::Line {
                    line: crate::render::geometry::PtLineSegment::new(
                        PtOffset::new(para_left, box_top),
                        PtOffset::new(para_right, box_top),
                    ),
                    color: top.color,
                    width: top.width,
                });
            }
        }
        if draw_bottom {
            if let Some(ref bottom) = borders.bottom {
                commands.push(DrawCommand::Line {
                    line: crate::render::geometry::PtLineSegment::new(
                        PtOffset::new(para_left, box_bottom),
                        PtOffset::new(para_right, box_bottom),
                    ),
                    color: bottom.color,
                    width: bottom.width,
                });
            }
        }
        if let Some(ref left) = borders.left {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, box_top),
                    PtOffset::new(para_left, box_bottom),
                ),
                color: left.color,
                width: left.width,
            });
        }
        if let Some(ref right) = borders.right {
            commands.push(DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    PtOffset::new(para_right, box_top),
                    PtOffset::new(para_right, box_bottom),
                ),
                color: right.color,
                width: right.width,
            });
        }
    }

    // §17.3.1.24: only the final segment carries the bottom border space and
    // space_after (§17.3.1.33).
    let mut cursor_y = content_bottom;
    if draw_bottom {
        cursor_y = content_bottom + border_space_bottom + style.space_after;
    }

    // If no lines, still consume default height + spacing (whole-paragraph
    // path). Gridded like any other line (§17.6.5): an empty paragraph on a
    // gridded page holds a slot open, exactly as Word's empty 行 does.
    if no_lines {
        let line_h =
            super::line_emit::resolve_line_metrics(default_line_height, default_line_height, style)
                .advance;
        cursor_y = style.space_before + line_h + style.space_after;
    }

    cursor_y
}
