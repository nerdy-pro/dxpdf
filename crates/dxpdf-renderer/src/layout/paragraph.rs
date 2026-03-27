//! Paragraph layout — line fitting, alignment, spacing, borders, shading.
//!
//! Implements the LayoutBox protocol: receives BoxConstraints, returns PtSize,
//! emits DrawCommands at absolute offsets during paint.

use dxpdf_docx_model::model::Alignment;

use crate::dimension::Pt;
use crate::geometry::{PtOffset, PtSize};
use crate::resolve::color::RgbColor;
use super::draw_command::DrawCommand;
use super::fragment::Fragment;
use super::BoxConstraints;

/// Configuration for paragraph layout.
#[derive(Clone, Debug)]
pub struct ParagraphStyle {
    pub alignment: Alignment,
    pub space_before: Pt,
    pub space_after: Pt,
    pub indent_left: Pt,
    pub indent_right: Pt,
    pub indent_first_line: Pt,
    pub line_spacing: LineSpacingRule,
    /// Drop cap to render at the start of this paragraph.
    pub drop_cap: Option<DropCapInfo>,
    /// §17.3.1.24: paragraph borders.
    pub borders: Option<ParagraphBorderStyle>,
    /// §17.3.1.31: paragraph shading (background fill).
    pub shading: Option<RgbColor>,
    /// §17.4.58: floating element beside this paragraph.
    /// (float_width, remaining_float_height) — lines within this height get
    /// narrower width; lines below get full width.
    pub float_beside: Option<(Pt, Pt)>,
}

/// Resolved paragraph border style for rendering.
#[derive(Clone, Debug)]
pub struct ParagraphBorderStyle {
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
    pub right: Option<BorderLine>,
}

/// A single border line for rendering.
#[derive(Clone, Copy, Debug)]
pub struct BorderLine {
    pub width: Pt,
    pub color: RgbColor,
    /// §17.3.1.24: space between border and text in points.
    pub space: Pt,
}

/// Drop cap letter to float at the start of a paragraph.
#[derive(Clone, Debug)]
pub struct DropCapInfo {
    /// The drop cap fragments (usually a single large letter).
    pub fragments: Vec<Fragment>,
    /// §17.3.1.11 @w:lines: number of body text lines the drop cap spans.
    pub lines: u32,
    /// Total width of the drop cap (measured).
    pub width: Pt,
    /// Total height of the drop cap (measured).
    pub height: Pt,
    /// Ascent of the drop cap font (for baseline positioning).
    pub ascent: Pt,
    /// §17.3.1.11 @w:hSpace: horizontal distance from surrounding text.
    pub h_space: Pt,
}

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            alignment: Alignment::Start,
            space_before: Pt::ZERO,
            space_after: Pt::ZERO,
            indent_left: Pt::ZERO,
            indent_right: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            line_spacing: LineSpacingRule::Auto(1.0),
            drop_cap: None,
            borders: None,
            shading: None,
            float_beside: None,
        }
    }
}

// Workaround for clippy: ParagraphStyle has many fields but they all map 1:1 to spec properties.
// A builder pattern would add complexity without value here.

/// Line spacing rules matching OOXML semantics.
#[derive(Clone, Copy, Debug)]
pub enum LineSpacingRule {
    /// Proportional: multiplier on natural line height (1.0 = single, 1.5 = 1.5x, etc.)
    Auto(f32),
    /// Exact line height in points.
    Exact(Pt),
    /// Minimum line height in points.
    AtLeast(Pt),
}

/// Result of laying out a paragraph.
#[derive(Debug)]
pub struct ParagraphLayout {
    /// Draw commands positioned relative to the paragraph's top-left origin.
    pub commands: Vec<DrawCommand>,
    /// Total size consumed by this paragraph (including spacing).
    pub size: PtSize,
}

/// Lay out a paragraph: fit fragments into lines, apply alignment and spacing.
///
/// Returns draw commands positioned relative to (0, 0). The caller positions
/// the paragraph by adding its offset during the paint phase.
pub fn layout_paragraph(
    fragments: &[Fragment],
    constraints: &BoxConstraints,
    style: &ParagraphStyle,
    default_line_height: Pt,
) -> ParagraphLayout {
    // §17.3.1.11: drop cap text frame. The width includes any w:hSpace from framePr.
    let drop_cap_indent = style
        .drop_cap
        .as_ref()
        .map(|dc| dc.width + dc.h_space)
        .unwrap_or(Pt::ZERO);
    let drop_cap_lines = style
        .drop_cap
        .as_ref()
        .map(|dc| dc.lines as usize)
        .unwrap_or(0);

    // §17.3.1.24: border space is the distance between the border line and the text.
    // Only the space reduces the text area, not the border line width.
    let border_space_left = style
        .borders
        .as_ref()
        .and_then(|b| b.left.as_ref())
        .map(|b| b.space)
        .unwrap_or(Pt::ZERO);
    let border_space_right = style
        .borders
        .as_ref()
        .and_then(|b| b.right.as_ref())
        .map(|b| b.space)
        .unwrap_or(Pt::ZERO);
    let content_width = constraints.max_width
        - style.indent_left
        - style.indent_right
        - border_space_left
        - border_space_right;
    // §17.3.1.12: first-line indent adjusts the first line's available width.
    // Positive = narrower (indent), negative = wider (hanging indent).
    // Drop cap indent also reduces width for the first N lines.
    // §17.4.58: if a float is beside this paragraph, lines within the float
    // height use narrower width; lines below use full content_width.
    let float_reduction = style
        .float_beside
        .map(|(fw, _)| fw)
        .unwrap_or(Pt::ZERO);
    let narrow_width = (content_width - float_reduction).max(Pt::ZERO);
    let wide_width = content_width;

    let first_line_adjustment = style.indent_first_line + drop_cap_indent;
    let first_line_width = (narrow_width - first_line_adjustment).max(Pt::ZERO);
    let remaining_narrow = if drop_cap_indent > Pt::ZERO {
        narrow_width - drop_cap_indent
    } else {
        narrow_width
    };

    // Fit lines: if float_beside is set, fit at narrow width first, then refit
    // remaining fragments at full width once past the float height.
    // Number of lines fit at narrow width (beside a float). Lines after this
    // index were fit at full width and should not get float_offset in rendering.
    let mut narrow_line_count: usize = usize::MAX; // no float = all lines treated uniformly

    let lines = if let Some((_, float_remaining_height)) = style.float_beside {
        let mut all_lines = Vec::new();
        let narrow_lines = super::line::fit_lines_with_first(fragments, first_line_width, remaining_narrow);

        let mut accum_height = Pt::ZERO;
        let mut last_narrow_end = fragments.len();

        for line in &narrow_lines {
            let natural = if line.height > Pt::ZERO { line.height } else { default_line_height };
            let lh = resolve_line_height(natural, &style.line_spacing);
            // A line is beside the float if its TOP (accum_height) is within
            // the float area. The line may extend past the float bottom — that's
            // OK, the next line will be full-width.
            if accum_height >= float_remaining_height {
                last_narrow_end = line.start;
                break;
            }
            all_lines.push(super::line::FittedLine {
                start: line.start,
                end: line.end,
                width: line.width,
                height: line.height,
                ascent: line.ascent,
                has_break: line.has_break,
            });
            accum_height += lh;
        }

        narrow_line_count = all_lines.len();

        // Refit remaining fragments at full width.
        if last_narrow_end < fragments.len() {
            let remaining_frags = &fragments[last_narrow_end..];
            let wide_lines = super::line::fit_lines_with_first(remaining_frags, wide_width, wide_width);
            for wl in wide_lines {
                all_lines.push(super::line::FittedLine {
                    start: wl.start + last_narrow_end,
                    end: wl.end + last_narrow_end,
                    width: wl.width,
                    height: wl.height,
                    ascent: wl.ascent,
                    has_break: wl.has_break,
                });
            }
        }
        all_lines
    } else {
        super::line::fit_lines_with_first(fragments, first_line_width, remaining_narrow)
    };

    let mut commands = Vec::new();
    let mut cursor_y = style.space_before;

    // §17.3.1.11: compute the baseline of the Nth body text line for drop cap positioning.
    // The drop cap baseline must align with the baseline of the last line it spans.
    let drop_cap_baseline_y = if let Some(ref dc) = style.drop_cap {
        let n = dc.lines.max(1) as usize;
        let mut y = cursor_y;
        for (i, fitted_line) in lines.iter().enumerate().take(n) {
            let natural = if fitted_line.height > Pt::ZERO {
                fitted_line.height
            } else {
                default_line_height
            };
            let lh = resolve_line_height(natural, &style.line_spacing);
            if i == n - 1 {
                // Nth line: baseline = current y + this line's ascent.
                y += fitted_line.ascent;
                break;
            }
            y += lh;
        }
        Some(y)
    } else {
        None
    };

    // Render drop cap at the computed baseline.
    // §17.3.1.11: the drop cap baseline aligns with the Nth body line's baseline,
    // but the glyph's top must not extend above the paragraph start (cursor_y).
    // When cross-engine font metric differences cause the glyph to be taller than
    // the N-line span, clamp so the top aligns with the paragraph start instead.
    if let (Some(ref dc), Some(nth_baseline_y)) = (&style.drop_cap, drop_cap_baseline_y) {
        let baseline_y = nth_baseline_y.max(cursor_y + dc.ascent);
        for frag in &dc.fragments {
            if let Fragment::Text {
                text, font, color, ..
            } = frag
            {
                commands.push(DrawCommand::Text {
                    position: PtOffset::new(style.indent_left, baseline_y),
                    text: text.clone(),
                    font_family: font.family.clone(),
                    char_spacing: font.char_spacing,
                    font_size: font.size,
                    bold: font.bold,
                    italic: font.italic,
                    color: *color,
                });
            }
        }
    }

    let mut accum_line_height = Pt::ZERO;

    for (line_idx, line) in lines.iter().enumerate() {
        // Drop cap lines get extra indent; after that, refit remaining lines at full width.
        let dc_offset = if line_idx < drop_cap_lines {
            drop_cap_indent
        } else {
            Pt::ZERO
        };

        // §17.4.58: lines beside the float get extra left indent.
        // Use narrow_line_count (from fitting) to determine which lines are
        // beside the float, ensuring fitting and rendering agree on the boundary.
        let float_offset = if let Some((fw, _)) = style.float_beside {
            if line_idx < narrow_line_count { fw } else { Pt::ZERO }
        } else {
            Pt::ZERO
        };

        let indent = if line_idx == 0 {
            style.indent_left + style.indent_first_line + dc_offset + float_offset
        } else {
            style.indent_left + dc_offset + float_offset
        };

        let natural_height = if line.height > Pt::ZERO {
            line.height
        } else {
            default_line_height
        };
        let line_height = resolve_line_height(natural_height, &style.line_spacing);

        // Alignment offset — computed relative to the line's available width.
        // First line may be narrower (first-line indent), lines beside float are
        // narrower, lines past float use full width.
        let line_available = if float_offset > Pt::ZERO {
            narrow_width
        } else if line_idx == 0 {
            first_line_width
        } else {
            wide_width
        };
        let remaining = (line_available - line.width).max(Pt::ZERO);
        let align_offset = match style.alignment {
            Alignment::Center => remaining * 0.5,
            Alignment::End => remaining,
            Alignment::Both if !line.has_break && line_idx < lines.len() - 1 => Pt::ZERO, // justify handled separately
            _ => Pt::ZERO,
        };

        let x_start = indent + align_offset;

        // Emit text commands for this line
        let mut x = x_start;
        for frag in &fragments[line.start..line.end] {
            match frag {
                Fragment::Text {
                    text,
                    font,
                    color,
                    shading,
                    border,
                    width,
                    hyperlink_url,
                    baseline_offset,
                    ..
                } => {
                    // §17.3.2.32: render run-level shading behind text.
                    if let Some(bg_color) = shading {
                        commands.push(DrawCommand::Rect {
                            rect: crate::geometry::PtRect::from_xywh(
                                x,
                                cursor_y,
                                *width,
                                line_height,
                            ),
                            color: *bg_color,
                        });
                    }

                    // §17.3.2.4: render run-level border (box around text).
                    if let Some(bdr) = border {
                        let bx = x - bdr.space;
                        let by = cursor_y;
                        let bw = *width + bdr.space * 2.0;
                        let bh = line_height;
                        let half = bdr.width * 0.5;
                        // Top
                        commands.push(DrawCommand::Line {
                            line: crate::geometry::PtLineSegment::new(
                                PtOffset::new(bx, by + half),
                                PtOffset::new(bx + bw, by + half),
                            ),
                            color: bdr.color, width: bdr.width,
                        });
                        // Bottom
                        commands.push(DrawCommand::Line {
                            line: crate::geometry::PtLineSegment::new(
                                PtOffset::new(bx, by + bh - half),
                                PtOffset::new(bx + bw, by + bh - half),
                            ),
                            color: bdr.color, width: bdr.width,
                        });
                        // Left
                        commands.push(DrawCommand::Line {
                            line: crate::geometry::PtLineSegment::new(
                                PtOffset::new(bx + half, by),
                                PtOffset::new(bx + half, by + bh),
                            ),
                            color: bdr.color, width: bdr.width,
                        });
                        // Right
                        commands.push(DrawCommand::Line {
                            line: crate::geometry::PtLineSegment::new(
                                PtOffset::new(bx + bw - half, by),
                                PtOffset::new(bx + bw - half, by + bh),
                            ),
                            color: bdr.color, width: bdr.width,
                        });
                    }

                    let y = cursor_y + line.ascent + *baseline_offset;
                    commands.push(DrawCommand::Text {
                        position: PtOffset::new(x, y),
                        text: text.clone(),
                        font_family: font.family.clone(),
                        char_spacing: font.char_spacing,
                        font_size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        color: *color,
                    });

                    if let Some(url) = hyperlink_url {
                        commands.push(DrawCommand::LinkAnnotation {
                            rect: crate::geometry::PtRect::from_xywh(
                                x,
                                cursor_y,
                                *width,
                                line_height,
                            ),
                            url: url.clone(),
                        });
                    }

                    if font.underline {
                        // §17.3.2.40: underline position and thickness from font metrics.
                        // Skia provides underlinePosition (negative = below baseline)
                        // and underlineThickness.
                        let underline_y = y - font.underline_position;
                        let stroke_width = font.underline_thickness;
                        commands.push(DrawCommand::Underline {
                            line: crate::geometry::PtLineSegment::new(
                                PtOffset::new(x, underline_y),
                                PtOffset::new(x + *width, underline_y),
                            ),
                            color: *color,
                            width: stroke_width,
                        });
                    }

                    x += *width;
                }
                Fragment::Image {
                    size, image_data, ..
                } => {
                    if let Some(data) = image_data {
                        commands.push(DrawCommand::Image {
                            rect: crate::geometry::PtRect::from_xywh(
                                x,
                                cursor_y,
                                size.width,
                                size.height,
                            ),
                            image_data: data.clone(),
                        });
                    }
                    x += size.width;
                }
                Fragment::Tab { .. } => {
                    x += super::fragment::MIN_TAB_WIDTH;
                }
                Fragment::LineBreak { .. } => {}
            }
        }

        cursor_y += line_height;
        accum_line_height += line_height;
    }

    // §17.3.1.24: paragraph border and shading coordinate system.
    // Borders sit at the paragraph indent edges. The border `space` is the
    // distance inward from the border to the text. Shading fills the area
    // enclosed by the borders.
    let para_left = style.indent_left;
    let para_right = constraints.max_width - style.indent_right;
    let para_top = style.space_before;
    let para_bottom = cursor_y;

    // §17.3.1.31: render paragraph shading (fills the border area).
    if let Some(bg_color) = style.shading {
        commands.insert(0, DrawCommand::Rect {
            rect: crate::geometry::PtRect::from_xywh(
                para_left,
                para_top,
                para_right - para_left,
                para_bottom - para_top,
            ),
            color: bg_color,
        });
    }

    // §17.3.1.24: render paragraph borders at the indent edges.
    if let Some(ref borders) = style.borders {
        if let Some(ref top) = borders.top {
            commands.push(DrawCommand::Line {
                line: crate::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, para_top),
                    PtOffset::new(para_right, para_top),
                ),
                color: top.color,
                width: top.width,
            });
        }
        if let Some(ref bottom) = borders.bottom {
            commands.push(DrawCommand::Line {
                line: crate::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, para_bottom),
                    PtOffset::new(para_right, para_bottom),
                ),
                color: bottom.color,
                width: bottom.width,
            });
        }
        if let Some(ref left) = borders.left {
            commands.push(DrawCommand::Line {
                line: crate::geometry::PtLineSegment::new(
                    PtOffset::new(para_left, para_top),
                    PtOffset::new(para_left, para_bottom),
                ),
                color: left.color,
                width: left.width,
            });
        }
        if let Some(ref right) = borders.right {
            commands.push(DrawCommand::Line {
                line: crate::geometry::PtLineSegment::new(
                    PtOffset::new(para_right, para_top),
                    PtOffset::new(para_right, para_bottom),
                ),
                color: right.color,
                width: right.width,
            });
        }
    }

    cursor_y += style.space_after;

    // If no lines, still consume default height + spacing.
    if lines.is_empty() {
        cursor_y = style.space_before + default_line_height + style.space_after;
    }

    let total_height = constraints.constrain(PtSize::new(constraints.max_width, cursor_y)).height;

    ParagraphLayout {
        commands,
        size: PtSize::new(constraints.max_width, total_height),
    }
}

fn resolve_line_height(natural: Pt, rule: &LineSpacingRule) -> Pt {
    match rule {
        LineSpacingRule::Auto(multiplier) => natural * *multiplier,
        LineSpacingRule::Exact(h) => *h,
        LineSpacingRule::AtLeast(min) => natural.max(*min),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::layout::fragment::FontProps;
    use crate::resolve::color::RgbColor;
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
            width: Pt::new(width), trimmed_width: Pt::new(width),
            height: Pt::new(14.0),
            ascent: Pt::new(10.0),
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO,
        }
    }

    fn body_constraints(width: f32) -> BoxConstraints {
        BoxConstraints::new(
            Pt::ZERO, Pt::new(width),
            Pt::ZERO, Pt::new(1000.0),
        )
    }

    #[test]
    fn empty_paragraph_has_default_height() {
        let result = layout_paragraph(
            &[],
            &body_constraints(400.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
        );
        assert_eq!(result.size.height.raw(), 14.0, "default line height");
        assert!(result.commands.is_empty());
    }

    #[test]
    fn single_line_produces_text_command() {
        let frags = vec![text_frag("hello", 30.0)];
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
        );

        assert_eq!(result.commands.len(), 1);
        if let DrawCommand::Text { text, position, .. } = &result.commands[0] {
            assert_eq!(text, "hello");
            assert_eq!(position.x.raw(), 0.0); // left aligned, no indent
        }
    }

    #[test]
    fn center_alignment_shifts_text() {
        let frags = vec![text_frag("hi", 20.0)];
        let style = ParagraphStyle {
            alignment: Alignment::Center,
            ..Default::default()
        };
        let result = layout_paragraph(&frags, &body_constraints(100.0), &style, Pt::new(14.0));

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 40.0); // (100 - 20) / 2
        }
    }

    #[test]
    fn end_alignment_right_aligns() {
        let frags = vec![text_frag("hi", 20.0)];
        let style = ParagraphStyle {
            alignment: Alignment::End,
            ..Default::default()
        };
        let result = layout_paragraph(&frags, &body_constraints(100.0), &style, Pt::new(14.0));

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 80.0); // 100 - 20
        }
    }

    #[test]
    fn indentation_shifts_text() {
        let frags = vec![text_frag("text", 40.0)];
        let style = ParagraphStyle {
            indent_left: Pt::new(36.0),
            ..Default::default()
        };
        let result = layout_paragraph(&frags, &body_constraints(400.0), &style, Pt::new(14.0));

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 36.0);
        }
    }

    #[test]
    fn first_line_indent() {
        let frags = vec![
            text_frag("first ", 40.0),
            text_frag("second", 40.0),
        ];
        let style = ParagraphStyle {
            indent_first_line: Pt::new(24.0),
            ..Default::default()
        };
        let result = layout_paragraph(&frags, &body_constraints(400.0), &style, Pt::new(14.0));

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 24.0, "first line indented");
        }
    }

    #[test]
    fn space_before_and_after() {
        let frags = vec![text_frag("text", 30.0)];
        let style = ParagraphStyle {
            space_before: Pt::new(10.0),
            space_after: Pt::new(8.0),
            ..Default::default()
        };
        let result = layout_paragraph(&frags, &body_constraints(400.0), &style, Pt::new(14.0));

        // Height should be: space_before(10) + line_height(14) + space_after(8) = 32
        assert_eq!(result.size.height.raw(), 32.0);

        // Text y should include space_before
        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert!(position.y.raw() >= 10.0, "y should account for space_before");
        }
    }

    #[test]
    fn line_spacing_exact() {
        let frags = vec![text_frag("line1 ", 60.0), text_frag("line2", 60.0)];
        let style = ParagraphStyle {
            line_spacing: LineSpacingRule::Exact(Pt::new(20.0)),
            ..Default::default()
        };
        // With max_width=80, they'll break into 2 lines
        let result = layout_paragraph(&frags, &body_constraints(80.0), &style, Pt::new(14.0));

        assert_eq!(result.size.height.raw(), 40.0, "2 lines * 20pt each");
    }

    #[test]
    fn line_spacing_at_least_with_larger_natural() {
        let frags = vec![text_frag("text", 30.0)];
        let style = ParagraphStyle {
            line_spacing: LineSpacingRule::AtLeast(Pt::new(10.0)),
            ..Default::default()
        };
        let result = layout_paragraph(&frags, &body_constraints(400.0), &style, Pt::new(14.0));

        // Natural height is 14, at-least is 10 → should be 14
        assert_eq!(result.size.height.raw(), 14.0);
    }

    #[test]
    fn wrapping_produces_multiple_lines() {
        let frags = vec![
            text_frag("word1 ", 45.0),
            text_frag("word2 ", 45.0),
            text_frag("word3", 45.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(80.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
        );

        // Should have 3 text commands (one per word, each on its own line)
        let text_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 3);
        // Height: 3 lines * 14pt = 42pt
        assert_eq!(result.size.height.raw(), 42.0);
    }

    #[test]
    fn resolve_line_height_auto() {
        assert_eq!(resolve_line_height(Pt::new(14.0), &LineSpacingRule::Auto(1.0)).raw(), 14.0);
        assert_eq!(resolve_line_height(Pt::new(14.0), &LineSpacingRule::Auto(1.5)).raw(), 21.0);
    }

    #[test]
    fn resolve_line_height_exact_overrides() {
        assert_eq!(
            resolve_line_height(Pt::new(14.0), &LineSpacingRule::Exact(Pt::new(20.0))).raw(),
            20.0
        );
    }

    #[test]
    fn resolve_line_height_at_least() {
        assert_eq!(
            resolve_line_height(Pt::new(14.0), &LineSpacingRule::AtLeast(Pt::new(10.0))).raw(),
            14.0,
            "natural > minimum"
        );
        assert_eq!(
            resolve_line_height(Pt::new(8.0), &LineSpacingRule::AtLeast(Pt::new(10.0))).raw(),
            10.0,
            "minimum > natural"
        );
    }
}
