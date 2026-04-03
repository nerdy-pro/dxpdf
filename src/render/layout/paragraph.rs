//! Paragraph layout — line fitting, alignment, spacing, borders, shading.
//!
//! Implements the LayoutBox protocol: receives BoxConstraints, returns PtSize,
//! emits DrawCommands at absolute offsets during paint.

use std::rc::Rc;

use crate::model::Alignment;

use super::draw_command::DrawCommand;
use super::fragment::Fragment;
use super::BoxConstraints;
use crate::render::dimension::Pt;
use crate::render::geometry::{PtOffset, PtSize};
use crate::render::resolve::color::RgbColor;

// ── Tab leader rendering constants ────────────────────────────────────────────

/// Maximum font size used when measuring/drawing tab leader characters (pt).
/// Caps at 12pt so leaders remain legible regardless of paragraph line height.
const LEADER_FONT_SIZE_CAP: Pt = Pt::new(12.0);

/// Fallback width for a single leader character when no measurer is available (pt).
const LEADER_CHAR_WIDTH_FALLBACK: Pt = Pt::new(4.0);

/// A fitted line together with the per-line float adjustments that were active when it was placed.
struct LinePlacement {
    line: super::line::FittedLine,
    /// Width stolen from the left by an active float.
    float_left: Pt,
    /// Width stolen from the right by an active float.
    float_right: Pt,
}

/// §17.3.1.38: a resolved tab stop for layout.
#[derive(Clone, Debug)]
pub struct TabStopDef {
    /// Absolute position from paragraph left edge.
    pub position: Pt,
    /// §17.18.81: tab alignment (left, center, right, decimal).
    pub alignment: crate::model::TabAlignment,
    /// §17.18.82: leader character (dot, hyphen, underscore, etc.).
    pub leader: crate::model::TabLeader,
}

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
    /// §17.3.1.38: custom tab stops.
    pub tabs: Vec<TabStopDef>,
    /// Drop cap to render at the start of this paragraph.
    pub drop_cap: Option<DropCapInfo>,
    /// §17.3.1.24: paragraph borders.
    pub borders: Option<ParagraphBorderStyle>,
    /// §17.3.1.31: paragraph shading (background fill).
    pub shading: Option<RgbColor>,
    /// §17.3.1.14: keep this paragraph on the same page as the next.
    pub keep_next: bool,
    /// §17.3.1.9: suppress spacing between paragraphs of the same style.
    pub contextual_spacing: bool,
    /// Style ID for contextual spacing comparison.
    pub style_id: Option<crate::model::StyleId>,
    /// Active floats for per-line width adjustment.
    pub page_floats: Vec<super::float::ActiveFloat>,
    /// Absolute y position of this paragraph on the page (for float overlap checks).
    pub page_y: Pt,
    /// Left margin x position (for float_adjustments computation).
    pub page_x: Pt,
    /// Total content width (for float_adjustments computation).
    pub page_content_width: Pt,
}

/// Resolved paragraph border style for rendering.
#[derive(Clone, Debug, PartialEq)]
pub struct ParagraphBorderStyle {
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
    pub right: Option<BorderLine>,
}

/// A single border line for rendering.
#[derive(Clone, Copy, Debug, PartialEq)]
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
    /// §17.3.1.11: true = Margin mode (drop cap in margin), false = Drop mode (in text area).
    pub margin_mode: bool,
    /// Left indent of the drop cap paragraph (from its own style cascade).
    pub indent: Pt,
    /// Frame height from the drop cap paragraph's spacing (lineRule="exact").
    pub frame_height: Option<Pt>,
    /// §17.3.2.19: vertical baseline offset in points (negative = down).
    pub position_offset: Pt,
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
            tabs: Vec::new(),
            drop_cap: None,
            borders: None,
            shading: None,
            keep_next: false,
            contextual_spacing: false,
            style_id: None,
            page_floats: Vec::new(),
            page_y: Pt::ZERO,
            page_x: Pt::ZERO,
            page_content_width: Pt::ZERO,
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

/// Optional text measurement callback for accurate per-character splitting.
pub type MeasureTextFn<'a> =
    Option<&'a dyn Fn(&str, &super::fragment::FontProps) -> (Pt, super::fragment::TextMetrics)>;

/// Lay out a paragraph: fit fragments into lines, apply alignment and spacing.
///
/// Returns draw commands positioned relative to (0, 0). The caller positions
/// the paragraph by adding its offset during the paint phase.
pub fn layout_paragraph(
    fragments: &[Fragment],
    constraints: &BoxConstraints,
    style: &ParagraphStyle,
    default_line_height: Pt,
    measure_text: MeasureTextFn<'_>,
) -> ParagraphLayout {
    // §17.3.1.11: drop cap text frame.
    // Drop mode: body text indented by drop cap position + width + hSpace.
    // Margin mode: drop cap is in the margin, body text is NOT indented.
    let drop_cap_indent = style
        .drop_cap
        .as_ref()
        .filter(|dc| !dc.margin_mode)
        .map(|dc| dc.indent + dc.width + dc.h_space)
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
    let first_line_adjustment = style.indent_first_line + drop_cap_indent;

    // Split oversized text fragments into per-character fragments so narrow
    // cells get character-level line breaking.
    let min_avail = (content_width - first_line_adjustment).max(Pt::ZERO);
    let split_frags;
    let fragments = if min_avail > Pt::ZERO {
        split_frags = split_oversized_fragments(fragments, min_avail, measure_text);
        &split_frags
    } else {
        fragments
    };

    // Per-line float adjustment: fit one line at a time, computing the available
    // width for each line based on its absolute y position on the page.
    // Each line stores its (float_left, float_right) adjustments for rendering.
    let line_placements = compute_line_placements(
        fragments,
        style,
        content_width,
        first_line_adjustment,
        drop_cap_indent,
        drop_cap_lines,
        default_line_height,
    );

    let mut commands = Vec::new();
    let mut cursor_y = style.space_before;

    // §17.3.1.11: compute the drop cap baseline.
    // When frame_height is set (lineRule="exact"), use:
    //   baseline = frame_top + frame_height - descent + position_offset
    // Otherwise fall back to aligning with the Nth body line's baseline.
    let drop_cap_baseline_y = if let Some(ref dc) = style.drop_cap {
        if let Some(fh) = dc.frame_height {
            let baseline = cursor_y + fh + dc.position_offset;
            Some(baseline)
        } else {
            // Fallback: align with Nth body line baseline.
            let n = dc.lines.max(1) as usize;
            let mut y = cursor_y;
            for (i, lp) in line_placements.iter().enumerate().take(n) {
                let natural = if lp.line.height > Pt::ZERO {
                    lp.line.height
                } else {
                    default_line_height
                };
                let text_h = if lp.line.text_height > Pt::ZERO {
                    lp.line.text_height
                } else {
                    default_line_height
                };
                let lh = resolve_line_height(natural, text_h, &style.line_spacing);
                if i == n - 1 {
                    y += lp.line.ascent;
                    break;
                }
                y += lh;
            }
            Some(y)
        }
    } else {
        None
    };

    // Render drop cap at the computed baseline.
    if let (Some(ref dc), Some(baseline_y)) = (&style.drop_cap, drop_cap_baseline_y) {
        // §17.3.1.11: position the drop cap using its own paragraph's indent.
        // Drop mode: at the drop cap paragraph's indent (inside text area).
        // Margin mode: in the page margin, to the left of text.
        let dc_x = if dc.margin_mode {
            dc.indent - dc.width - dc.h_space
        } else {
            dc.indent
        };
        for frag in &dc.fragments {
            if let Fragment::Text {
                text, font, color, ..
            } = frag
            {
                commands.push(DrawCommand::Text {
                    position: PtOffset::new(dc_x, baseline_y),
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

    emit_line_commands(
        &mut commands,
        &mut cursor_y,
        &line_placements,
        fragments,
        style,
        content_width,
        first_line_adjustment,
        drop_cap_indent,
        drop_cap_lines,
        default_line_height,
        measure_text,
    );

    // §17.3.1.24: paragraph border and shading coordinate system.
    // Borders sit at the paragraph indent edges. The border `space` is the
    // distance between the border line and the text content. Top/bottom
    // border space expands the bordered area vertically.
    cursor_y = emit_paragraph_borders_and_shading(
        &mut commands,
        style,
        constraints,
        cursor_y,
        default_line_height,
        line_placements.is_empty(),
    );

    let total_height = constraints
        .constrain(PtSize::new(constraints.max_width, cursor_y))
        .height;

    ParagraphLayout {
        commands,
        size: PtSize::new(constraints.max_width, total_height),
    }
}

// ── Extracted helpers ─────────────────────────────────────────────────────────

/// Fit all fragments into lines, computing per-line float adjustments when
/// active floats are present.
///
/// When `style.page_floats` is non-empty each line is fitted individually so
/// the available width can vary depending on the line's absolute y position.
/// When there are no floats a single call to `fit_lines_with_first` is used.
#[allow(clippy::too_many_arguments)]
fn compute_line_placements(
    fragments: &[Fragment],
    style: &ParagraphStyle,
    content_width: Pt,
    first_line_adjustment: Pt,
    drop_cap_indent: Pt,
    drop_cap_lines: usize,
    default_line_height: Pt,
) -> Vec<LinePlacement> {
    if style.page_floats.is_empty() {
        // No floats — use standard line fitting.
        let first_line_width = (content_width - first_line_adjustment).max(Pt::ZERO);
        let remaining_width = if drop_cap_indent > Pt::ZERO {
            content_width - drop_cap_indent
        } else {
            content_width
        };
        return super::line::fit_lines_with_first(fragments, first_line_width, remaining_width)
            .into_iter()
            .map(|line| LinePlacement {
                line,
                float_left: Pt::ZERO,
                float_right: Pt::ZERO,
            })
            .collect();
    }

    // Active floats — fit one line at a time so available width can change.
    let mut placements = Vec::new();
    let mut frag_idx = 0;
    let mut line_y = style.space_before;

    while frag_idx < fragments.len() {
        let abs_y = style.page_y + line_y;
        let (fl, fr) = super::float::float_adjustments_with_height(
            &style.page_floats,
            abs_y,
            default_line_height,
            style.page_x,
            style.page_content_width,
        );
        let float_reduction = fl + fr;
        let available = (content_width - float_reduction).max(Pt::ZERO);

        let is_first = placements.is_empty();
        let dc_adj = if placements.len() < drop_cap_lines {
            drop_cap_indent
        } else {
            Pt::ZERO
        };
        let line_width = if is_first {
            (available - first_line_adjustment).max(Pt::ZERO)
        } else {
            (available - dc_adj).max(Pt::ZERO)
        };

        // Fit one line at this width.
        let remaining = &fragments[frag_idx..];
        let fitted = super::line::fit_lines_with_first(remaining, line_width, line_width);
        let fitted_line = if let Some(first) = fitted.into_iter().next() {
            super::line::FittedLine {
                start: first.start + frag_idx,
                end: first.end + frag_idx,
                width: first.width,
                height: first.height,
                text_height: first.text_height,
                ascent: first.ascent,
                has_break: first.has_break,
            }
        } else {
            break;
        };

        let natural = if fitted_line.height > Pt::ZERO {
            fitted_line.height
        } else {
            default_line_height
        };
        let text_h = if fitted_line.text_height > Pt::ZERO {
            fitted_line.text_height
        } else {
            default_line_height
        };
        let lh = resolve_line_height(natural, text_h, &style.line_spacing);

        frag_idx = fitted_line.end;
        placements.push(LinePlacement {
            line: fitted_line,
            float_left: fl,
            float_right: fr,
        });
        line_y += lh;
    }
    placements
}

/// Emit `DrawCommand`s for all lines in a paragraph, advancing `cursor_y`
/// by the total line height consumed.
///
/// Handles per-fragment shading, run borders, text, hyperlinks, underlines,
/// images, tab stops (with leaders), bookmarks, and drop-cap indent offsets.
#[allow(clippy::too_many_arguments)]
fn emit_line_commands(
    commands: &mut Vec<DrawCommand>,
    cursor_y: &mut Pt,
    line_placements: &[LinePlacement],
    fragments: &[Fragment],
    style: &ParagraphStyle,
    content_width: Pt,
    first_line_adjustment: Pt,
    drop_cap_indent: Pt,
    drop_cap_lines: usize,
    default_line_height: Pt,
    measure_text: MeasureTextFn<'_>,
) {
    for (line_idx, lp) in line_placements.iter().enumerate() {
        let line = &lp.line;

        // Drop cap lines get extra indent; after that, refit remaining lines at full width.
        let dc_offset = if line_idx < drop_cap_lines {
            drop_cap_indent
        } else {
            Pt::ZERO
        };

        // §17.4.58: per-line float left indent.
        let float_offset = lp.float_left;

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
        let text_height = if line.text_height > Pt::ZERO {
            line.text_height
        } else {
            default_line_height
        };
        let line_height = resolve_line_height(natural_height, text_height, &style.line_spacing);

        // Alignment offset — computed relative to the line's available width.
        let float_reduction = lp.float_left + lp.float_right;
        let line_available = if line_idx == 0 {
            (content_width - float_reduction - first_line_adjustment).max(Pt::ZERO)
        } else {
            (content_width - float_reduction - dc_offset).max(Pt::ZERO)
        };
        let remaining = (line_available - line.width).max(Pt::ZERO);
        let align_offset = match style.alignment {
            Alignment::Center => remaining * 0.5,
            Alignment::End => remaining,
            Alignment::Both if !line.has_break && line_idx < line_placements.len() - 1 => Pt::ZERO,
            _ => Pt::ZERO,
        };

        let x_start = indent + align_offset;

        // Emit text commands for this line
        let mut x = x_start;
        for (frag_idx, frag) in (line.start..line.end).zip(&fragments[line.start..line.end]) {
            match frag {
                Fragment::Text {
                    text,
                    font,
                    color,
                    shading,
                    border,
                    width,
                    metrics,
                    hyperlink_url,
                    baseline_offset,
                    text_offset,
                    ..
                } => {
                    // §17.3.2.32: render run-level shading behind text.
                    // Uses text bounds (ascent+descent), not full line height.
                    if let Some(bg_color) = shading {
                        let text_top = *cursor_y + line.ascent - metrics.ascent;
                        commands.push(DrawCommand::Rect {
                            rect: crate::render::geometry::PtRect::from_xywh(
                                x,
                                text_top,
                                *width,
                                metrics.height(),
                            ),
                            color: *bg_color,
                        });
                    }

                    // §17.3.2.4: render run-level border (box around text).
                    // Uses text bounds, not full line height.
                    if let Some(bdr) = border {
                        let text_top = *cursor_y + line.ascent - metrics.ascent;
                        let bx = x - bdr.space;
                        let by = text_top;
                        let bw = *width + bdr.space * 2.0;
                        let bh = metrics.height();
                        let half = bdr.width * 0.5;
                        // Top
                        commands.push(DrawCommand::Line {
                            line: crate::render::geometry::PtLineSegment::new(
                                PtOffset::new(bx, by + half),
                                PtOffset::new(bx + bw, by + half),
                            ),
                            color: bdr.color,
                            width: bdr.width,
                        });
                        // Bottom
                        commands.push(DrawCommand::Line {
                            line: crate::render::geometry::PtLineSegment::new(
                                PtOffset::new(bx, by + bh - half),
                                PtOffset::new(bx + bw, by + bh - half),
                            ),
                            color: bdr.color,
                            width: bdr.width,
                        });
                        // Left
                        commands.push(DrawCommand::Line {
                            line: crate::render::geometry::PtLineSegment::new(
                                PtOffset::new(bx + half, by),
                                PtOffset::new(bx + half, by + bh),
                            ),
                            color: bdr.color,
                            width: bdr.width,
                        });
                        // Right
                        commands.push(DrawCommand::Line {
                            line: crate::render::geometry::PtLineSegment::new(
                                PtOffset::new(bx + bw - half, by),
                                PtOffset::new(bx + bw - half, by + bh),
                            ),
                            color: bdr.color,
                            width: bdr.width,
                        });
                    }

                    let y = *cursor_y + line.ascent + *baseline_offset;
                    commands.push(DrawCommand::Text {
                        position: PtOffset::new(x + *text_offset, y),
                        text: text.clone(),
                        font_family: font.family.clone(),
                        char_spacing: font.char_spacing,
                        font_size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        color: *color,
                    });

                    if let Some(url) = hyperlink_url {
                        let rect = crate::render::geometry::PtRect::from_xywh(
                            x,
                            *cursor_y,
                            *width,
                            line_height,
                        );
                        if url.starts_with("http://")
                            || url.starts_with("https://")
                            || url.starts_with("mailto:")
                            || url.starts_with("ftp://")
                        {
                            commands.push(DrawCommand::LinkAnnotation {
                                rect,
                                url: url.clone(),
                            });
                        } else {
                            // Internal bookmark link.
                            commands.push(DrawCommand::InternalLink {
                                rect,
                                destination: url.clone(),
                            });
                        }
                    }

                    if font.underline {
                        // §17.3.2.40: underline position and thickness from font metrics.
                        // Skia provides underlinePosition (negative = below baseline)
                        // and underlineThickness.
                        let underline_y = y - font.underline_position;
                        let stroke_width = font.underline_thickness;
                        commands.push(DrawCommand::Underline {
                            line: crate::render::geometry::PtLineSegment::new(
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
                            rect: crate::render::geometry::PtRect::from_xywh(
                                x,
                                *cursor_y,
                                size.width,
                                size.height,
                            ),
                            image_data: data.clone(),
                        });
                    }
                    x += size.width;
                }
                Fragment::Tab { .. } => {
                    // §17.3.1.37: resolve to the next tab stop.
                    // Tab stop positions are absolute from the paragraph's
                    // left edge, not relative to the text indent.
                    let (tab_pos, tab_stop) = find_next_tab_stop(x, &style.tabs, line_available);

                    let new_x = if let Some(ts) = tab_stop {
                        use crate::model::TabAlignment;
                        // §17.3.1.37: for right/center tabs, compute the width
                        // of content in this tab's zone — from here to the next
                        // Tab fragment or line end, whichever comes first.
                        let zone_end = fragments[frag_idx + 1..line.end]
                            .iter()
                            .position(|f| matches!(f, Fragment::Tab { .. }))
                            .map_or(line.end, |i| frag_idx + 1 + i);
                        match ts.alignment {
                            TabAlignment::Right => {
                                let zone_width: Pt = fragments[frag_idx + 1..zone_end]
                                    .iter()
                                    .map(|f| f.width())
                                    .sum();
                                (tab_pos - zone_width).max(x)
                            }
                            TabAlignment::Center => {
                                let zone_width: Pt = fragments[frag_idx + 1..zone_end]
                                    .iter()
                                    .map(|f| f.width())
                                    .sum();
                                (tab_pos - zone_width * 0.5).max(x)
                            }
                            _ => tab_pos,
                        }
                    } else {
                        tab_pos
                    };

                    // Emit leader characters between tab start and tab position.
                    if let Some(ts) = tab_stop {
                        emit_tab_leader(
                            commands,
                            ts.leader,
                            x,
                            new_x,
                            *cursor_y + line.ascent,
                            line_height,
                            measure_text,
                            default_line_height,
                        );
                    }

                    x = new_x;
                }
                Fragment::LineBreak { .. } | Fragment::ColumnBreak => {}
                Fragment::Bookmark { name } => {
                    commands.push(DrawCommand::NamedDestination {
                        position: PtOffset::new(x, *cursor_y),
                        name: name.clone(),
                    });
                }
            }
        }

        *cursor_y += line_height;
    }
}

/// Emit paragraph border and shading commands, then advance the cursor for
/// border bottom space, `space_after`, and the empty-paragraph minimum height.
///
/// Returns the updated `cursor_y` after all spacing has been applied.
fn emit_paragraph_borders_and_shading(
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

/// Split text fragments wider than `max_width` into per-character fragments.
/// Uses accurate measurements when a measurer is provided, otherwise
/// falls back to uniform width distribution.
fn split_oversized_fragments(
    fragments: &[Fragment],
    max_width: Pt,
    measure: MeasureTextFn<'_>,
) -> Vec<Fragment> {
    let mut result = Vec::with_capacity(fragments.len());
    let mut any_split = false;
    for frag in fragments {
        match frag {
            Fragment::Text {
                text,
                width,
                font,
                color,
                shading,
                border,
                metrics,
                hyperlink_url,
                baseline_offset,
                ..
            } if *width > max_width && text.chars().count() > 1 => {
                any_split = true;
                for ch in text.chars() {
                    let ch_str = ch.to_string();
                    let (w, char_metrics) = if let Some(m) = measure {
                        m(&ch_str, font)
                    } else {
                        let per_char = *width / text.chars().count() as f32;
                        (per_char, *metrics)
                    };
                    result.push(Fragment::Text {
                        text: Rc::from(ch_str.as_str()),
                        font: font.clone(),
                        color: *color,
                        shading: *shading,
                        border: *border,
                        width: w,
                        trimmed_width: w,
                        metrics: char_metrics,
                        hyperlink_url: hyperlink_url.clone(),
                        baseline_offset: *baseline_offset,
                        text_offset: Pt::ZERO,
                    });
                }
            }
            _ => result.push(frag.clone()),
        }
    }
    if !any_split {
        return fragments.to_vec();
    }
    result
}

/// §17.3.1.37: find the next tab stop position greater than `current_x`.
/// Returns (position, optional tab stop definition).
/// If no custom tab stop matches, uses default tab stops every 36pt (0.5 inch).
fn find_next_tab_stop(
    current_x: Pt,
    tabs: &[TabStopDef],
    line_width: Pt,
) -> (Pt, Option<&TabStopDef>) {
    // §17.15.1.25: default tab stop interval is 36pt (0.5 inch).
    const DEFAULT_TAB_INTERVAL: f32 = 36.0;

    // Find the first custom tab stop past current position.
    for ts in tabs {
        if ts.position > current_x {
            return (ts.position, Some(ts));
        }
    }

    // No custom tab stop — use default interval.
    let next = ((current_x.raw() / DEFAULT_TAB_INTERVAL).floor() + 1.0) * DEFAULT_TAB_INTERVAL;
    (Pt::new(next.min(line_width.raw())), None)
}

/// Emit leader characters (dots, hyphens, etc.) between tab start and end.
#[allow(clippy::too_many_arguments)]
fn emit_tab_leader(
    commands: &mut Vec<DrawCommand>,
    leader: crate::model::TabLeader,
    x_start: Pt,
    x_end: Pt,
    baseline_y: Pt,
    _line_height: Pt,
    measure_text: MeasureTextFn<'_>,
    default_line_height: Pt,
) {
    use crate::model::TabLeader;

    let leader_char = match leader {
        TabLeader::Dot => ".",
        TabLeader::Hyphen => "-",
        TabLeader::Underscore => "_",
        TabLeader::MiddleDot => "\u{00B7}",
        TabLeader::Heavy => "_",
        TabLeader::None => return,
    };

    let gap = x_end - x_start;
    if gap <= Pt::ZERO {
        return;
    }

    // Build a string of leader characters that fills the gap.
    // Use a small font to get the leader char width.
    let leader_font = super::fragment::FontProps {
        family: std::rc::Rc::from("Times New Roman"),
        size: default_line_height.min(LEADER_FONT_SIZE_CAP),
        bold: false,
        italic: false,
        underline: false,
        char_spacing: Pt::ZERO,
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    };

    let char_width = if let Some(m) = measure_text {
        m(leader_char, &leader_font).0
    } else {
        LEADER_CHAR_WIDTH_FALLBACK
    };

    if char_width <= Pt::ZERO {
        return;
    }

    let count = ((gap / char_width) as usize).min(500);
    if count == 0 {
        return;
    }

    let leader_text: String = leader_char.repeat(count);
    let leader_width = char_width * count as f32;

    // Right-align the leader dots within the gap (looks cleaner).
    let leader_x = x_end - leader_width;

    commands.push(DrawCommand::Text {
        position: PtOffset::new(leader_x.max(x_start), baseline_y),
        text: Rc::from(leader_text.as_str()),
        font_family: leader_font.family,
        char_spacing: Pt::ZERO,
        font_size: leader_font.size,
        bold: false,
        italic: false,
        color: crate::render::resolve::color::RgbColor::BLACK,
    });
}

/// §17.3.1.33: resolve the effective line height from the natural height
/// and the line spacing rule.
///
/// For Auto mode, the multiplier applies only to text metrics — inline
/// images use their natural height without scaling. The final line height
/// is `max(text_height * multiplier, total_height)`.
fn resolve_line_height(natural: Pt, text_height: Pt, rule: &LineSpacingRule) -> Pt {
    match rule {
        LineSpacingRule::Auto(multiplier) => {
            let scaled_text = text_height * *multiplier;
            // Use the scaled text height or the full natural height (which
            // includes images), whichever is larger.
            scaled_text.max(natural)
        }
        LineSpacingRule::Exact(h) => *h,
        LineSpacingRule::AtLeast(min) => natural.max(*min),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::layout::fragment::{FontProps, TextMetrics};
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

    fn body_constraints(width: f32) -> BoxConstraints {
        BoxConstraints::new(Pt::ZERO, Pt::new(width), Pt::ZERO, Pt::new(1000.0))
    }

    #[test]
    fn empty_paragraph_has_default_height() {
        let result = layout_paragraph(
            &[],
            &body_constraints(400.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
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
            None,
        );

        assert_eq!(result.commands.len(), 1);
        if let DrawCommand::Text { text, position, .. } = &result.commands[0] {
            assert_eq!(&**text, "hello");
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
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &style,
            Pt::new(14.0),
            None,
        );

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
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &style,
            Pt::new(14.0),
            None,
        );

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
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 36.0);
        }
    }

    #[test]
    fn first_line_indent() {
        let frags = vec![text_frag("first ", 40.0), text_frag("second", 40.0)];
        let style = ParagraphStyle {
            indent_first_line: Pt::new(24.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

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
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        // Height should be: space_before(10) + line_height(14) + space_after(8) = 32
        assert_eq!(result.size.height.raw(), 32.0);

        // Text y should include space_before
        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert!(
                position.y.raw() >= 10.0,
                "y should account for space_before"
            );
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
        let result = layout_paragraph(&frags, &body_constraints(80.0), &style, Pt::new(14.0), None);

        assert_eq!(result.size.height.raw(), 40.0, "2 lines * 20pt each");
    }

    #[test]
    fn line_spacing_at_least_with_larger_natural() {
        let frags = vec![text_frag("text", 30.0)];
        let style = ParagraphStyle {
            line_spacing: LineSpacingRule::AtLeast(Pt::new(10.0)),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

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
            None,
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
    fn resolve_line_height_auto_text_only() {
        // Text-only line: multiplier applies to text_height.
        assert_eq!(
            resolve_line_height(Pt::new(14.0), Pt::new(14.0), &LineSpacingRule::Auto(1.0)).raw(),
            14.0
        );
        assert_eq!(
            resolve_line_height(Pt::new(14.0), Pt::new(14.0), &LineSpacingRule::Auto(1.5)).raw(),
            21.0
        );
    }

    #[test]
    fn resolve_line_height_auto_image_line() {
        // Image-only line: natural=325 (image), text_height=0 (no text).
        // The multiplier does NOT inflate the image height.
        let h = resolve_line_height(Pt::new(325.0), Pt::ZERO, &LineSpacingRule::Auto(1.08));
        assert_eq!(h.raw(), 325.0, "image height should not be multiplied");
    }

    #[test]
    fn resolve_line_height_auto_mixed_line() {
        // Line with text (14pt) and image (100pt): multiplier scales text only.
        // max(14*1.5=21, 100) = 100.
        let h = resolve_line_height(Pt::new(100.0), Pt::new(14.0), &LineSpacingRule::Auto(1.5));
        assert_eq!(h.raw(), 100.0, "image dominates");
    }

    #[test]
    fn resolve_line_height_exact_overrides() {
        assert_eq!(
            resolve_line_height(
                Pt::new(14.0),
                Pt::new(14.0),
                &LineSpacingRule::Exact(Pt::new(20.0))
            )
            .raw(),
            20.0
        );
    }

    #[test]
    fn resolve_line_height_at_least() {
        assert_eq!(
            resolve_line_height(
                Pt::new(14.0),
                Pt::new(14.0),
                &LineSpacingRule::AtLeast(Pt::new(10.0))
            )
            .raw(),
            14.0,
            "natural > minimum"
        );
        assert_eq!(
            resolve_line_height(
                Pt::new(8.0),
                Pt::new(8.0),
                &LineSpacingRule::AtLeast(Pt::new(10.0))
            )
            .raw(),
            10.0,
            "minimum > natural"
        );
    }
}
