//! Line placement, command emission, and related helpers.

use std::rc::Rc;

use super::super::draw_command::DrawCommand;
use super::super::fragment::Fragment;
use super::types::{LineSpacingRule, MeasureTextFn, ParagraphStyle, TabStopDef};
use super::{LineLayoutParams, LinePlacement, LEADER_CHAR_WIDTH_FALLBACK, LEADER_FONT_SIZE_CAP};
use crate::render::dimension::Pt;
use crate::render::geometry::PtOffset;

/// Fit all fragments into lines, computing per-line float adjustments when
/// active floats are present.
///
/// When `style.page_floats` is non-empty each line is fitted individually so
/// the available width can vary depending on the line's absolute y position.
/// When there are no floats a single call to `fit_lines_with_first` is used.
pub(super) fn compute_line_placements(
    fragments: &[Fragment],
    style: &ParagraphStyle,
    params: &LineLayoutParams,
) -> Vec<LinePlacement> {
    let content_width = params.content_width;
    let first_line_adjustment = params.first_line_adjustment;
    let drop_cap_indent = params.drop_cap_indent;
    let drop_cap_lines = params.drop_cap_lines;
    let default_line_height = params.default_line_height;
    if style.page_floats.is_empty() {
        // No floats — use standard line fitting.
        let first_line_width = (content_width - first_line_adjustment).max(Pt::ZERO);
        let remaining_width = if drop_cap_indent > Pt::ZERO {
            content_width - drop_cap_indent
        } else {
            content_width
        };
        return super::super::line::fit_lines_with_first(
            fragments,
            first_line_width,
            remaining_width,
        )
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
        let (fl, fr) = super::super::float::float_adjustments_with_height(
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
        let fitted = super::super::line::fit_lines_with_first(remaining, line_width, line_width);
        let fitted_line = if let Some(first) = fitted.into_iter().next() {
            super::super::line::FittedLine {
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
pub(super) fn emit_line_commands(
    commands: &mut Vec<DrawCommand>,
    cursor_y: &mut Pt,
    line_placements: &[LinePlacement],
    fragments: &[Fragment],
    style: &ParagraphStyle,
    params: &LineLayoutParams,
    measure_text: MeasureTextFn<'_>,
) {
    let content_width = params.content_width;
    let first_line_adjustment = params.first_line_adjustment;
    let drop_cap_indent = params.drop_cap_indent;
    let drop_cap_lines = params.drop_cap_lines;
    let default_line_height = params.default_line_height;
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
        // §17.3.1.37: when a line contains tab characters, tab stops control
        // horizontal positioning — paragraph alignment does not apply.
        let line_has_tabs = fragments[line.start..line.end]
            .iter()
            .any(|f| matches!(f, Fragment::Tab { .. }));
        let align_offset = if line_has_tabs {
            Pt::ZERO
        } else {
            match style.alignment {
                crate::model::Alignment::Center => remaining * 0.5,
                crate::model::Alignment::End => remaining,
                crate::model::Alignment::Both
                    if !line.has_break && line_idx < line_placements.len() - 1 =>
                {
                    Pt::ZERO
                }
                _ => Pt::ZERO,
            }
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

/// Split text fragments wider than `max_width` into per-character fragments.
/// Uses accurate measurements when a measurer is provided, otherwise
/// falls back to uniform width distribution.
pub(super) fn split_oversized_fragments(
    fragments: &[Fragment],
    max_width: Pt,
    measure: MeasureTextFn<'_>,
) -> Vec<Fragment> {
    // Fast path: check if any fragment actually needs splitting.
    let needs_split = fragments.iter().any(
        |f| matches!(f, Fragment::Text { width, text, .. } if *width > max_width && text.len() > 1),
    );
    if !needs_split {
        return fragments.to_vec();
    }

    let mut result = Vec::with_capacity(fragments.len());
    // Reusable buffer for single-character measurement (avoids per-char heap allocation).
    let mut ch_buf = [0u8; 4];
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
                let char_count = text.chars().count();
                let per_char_fallback = *width / char_count as f32;
                for ch in text.chars() {
                    let ch_str = ch.encode_utf8(&mut ch_buf);
                    let (w, char_metrics) = if let Some(m) = measure {
                        m(ch_str, font)
                    } else {
                        (per_char_fallback, *metrics)
                    };
                    result.push(Fragment::Text {
                        text: Rc::from(&*ch_str),
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
    result
}

/// §17.3.1.37: find the next tab stop position greater than `current_x`.
/// Returns (position, optional tab stop definition).
/// If no custom tab stop matches, uses default tab stops every 36pt (0.5 inch).
pub(super) fn find_next_tab_stop(
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
pub(super) fn emit_tab_leader(
    commands: &mut Vec<DrawCommand>,
    leader: crate::model::TabLeader,
    x_start: Pt,
    x_end: Pt,
    baseline_y: Pt,
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
    let leader_font = super::super::fragment::FontProps {
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
pub(super) fn resolve_line_height(natural: Pt, text_height: Pt, rule: &LineSpacingRule) -> Pt {
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
