use std::rc::Rc;

use super::measurer::TextMeasurer;
use super::ImageCache;
use crate::dimension::Pt;
use crate::geometry::{PtLineSegment, PtOffset, PtRect, PtSize};
use crate::model::*;

/// Minimum tab fragment width for line fitting.
/// Prevents tabs from collapsing to zero width during line breaking.
pub const MIN_TAB_WIDTH: Pt = Pt::new(12.0);

/// Fallback tab advance when no stops and no default interval.
pub const TAB_FALLBACK: Pt = Pt::new(36.0);

/// Super/subscript font size as a fraction of the base size.
/// Matches Word's default rendering (~58% of base).
const VERT_ALIGN_SIZE_FACTOR: f32 = 0.58;

/// Superscript baseline shift as a fraction of the base font size (negative = up).
const SUPERSCRIPT_BASELINE_SHIFT: f32 = -0.33;

/// Subscript baseline shift as a fraction of the base font size (positive = down).
const SUBSCRIPT_BASELINE_SHIFT: f32 = 0.08;

/// Underline offset below the text baseline.
pub const UNDERLINE_Y_OFFSET: Pt = Pt::new(2.0);
use super::DrawCommand;

// DocDefaultsLayout is defined in measure.rs and re-exported here for backward compatibility.
pub use super::measure::DocDefaultsLayout;

/// Font properties that travel together for text rendering.
pub struct FontProps {
    pub font_family: Rc<str>,
    pub font_size: Pt,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// Character spacing in points (positive = expand).
    pub char_spacing_pt: Pt,
}

/// A flattened fragment for layout — either text, an image, a tab, or a line break.
pub enum Fragment {
    Text {
        text: String,
        font: FontProps,
        color: Option<Color>,
        shading: Option<Color>,
        measured_width: Pt,
        measured_height: Pt,
        /// Distance from top of line box to baseline (always positive).
        ascent: Pt,
        /// Hyperlink URL, if this text is part of a link.
        hyperlink_url: Option<String>,
        /// Baseline offset in points (negative = up for superscript, positive = down for subscript).
        baseline_offset: Pt,
    },
    Image {
        size: PtSize,
        rel_id: String,
    },
    Tab {
        line_height: Pt,
    },
    LineBreak {
        line_height: Pt,
    },
}

impl Fragment {
    pub fn width(&self) -> Pt {
        match self {
            Fragment::Text { measured_width, .. } => *measured_width,
            Fragment::Image { size, .. } => size.width,
            Fragment::Tab { .. } => MIN_TAB_WIDTH,
            Fragment::LineBreak { .. } => Pt::ZERO,
        }
    }

    pub fn height(&self) -> Pt {
        match self {
            Fragment::Text {
                measured_height, ..
            } => *measured_height,
            Fragment::Image { size, .. } => size.height,
            Fragment::Tab { line_height } | Fragment::LineBreak { line_height } => *line_height,
        }
    }

    pub fn is_line_break(&self) -> bool {
        matches!(self, Fragment::LineBreak { .. })
    }
}

/// Compute underline stroke width based on font size and weight.
/// Bold text gets a thicker underline proportional to the font size.
pub fn underline_width(font_size: Pt, bold: bool) -> Pt {
    let base = font_size * 0.05; // ~5% of font size
    if bold {
        base * 1.5
    } else {
        base
    }
}

/// Context for evaluating field codes during fragment collection.
pub struct FieldContext {
    pub page_number: u32,
    pub num_pages: u32,
}

pub fn collect_fragments_with_fields(
    runs: &[Inline],
    constraints: &super::context::LayoutConstraints,
    defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    field_ctx: Option<&FieldContext>,
    image_cache: &ImageCache,
) -> Vec<Fragment> {
    let content_width = constraints.available_width();
    let content_height = constraints.available_height();
    let mut fragments = Vec::new();
    for run in runs {
        match run {
            Inline::TextRun(tr) => {
                let font_family = tr
                    .properties
                    .font_family
                    .clone()
                    .unwrap_or_else(|| defaults.font_family.clone());
                let base_font_size =
                    Pt::from(tr.properties.font_size.unwrap_or(defaults.font_size));
                let bold = tr.properties.bold;
                let italic = tr.properties.italic;
                let char_spacing_pt = tr.properties.char_spacing.map(Pt::from).unwrap_or(Pt::ZERO);

                // Super/subscript: reduce font size and compute baseline offset
                let (font_size, baseline_offset) = match tr.properties.vert_align {
                    Some(VertAlign::Superscript) => {
                        let reduced = base_font_size * VERT_ALIGN_SIZE_FACTOR;
                        let offset = base_font_size * SUPERSCRIPT_BASELINE_SHIFT;
                        (reduced, offset)
                    }
                    Some(VertAlign::Subscript) => {
                        let reduced = base_font_size * VERT_ALIGN_SIZE_FACTOR;
                        let offset = base_font_size * SUBSCRIPT_BASELINE_SHIFT;
                        (reduced, offset)
                    }
                    None => (base_font_size, Pt::ZERO),
                };

                let base_font = measurer.font(&font_family, base_font_size, bold, italic);
                let fm = base_font.metrics();
                let line_height = fm.line_height;
                let ascent = fm.ascent;
                let render_font = measurer.font(&font_family, font_size, bold, italic);
                for part in split_words_and_spaces(&tr.text) {
                    let base_width = render_font.measure_width(part);
                    // Character spacing expands each character's advance width
                    let char_count = part.chars().count() as f32;
                    let measured_width = base_width + char_spacing_pt * char_count;
                    let is_space = part.chars().all(|c| c == ' ');
                    fragments.push(Fragment::Text {
                        text: part.to_string(),
                        font: FontProps {
                            font_family: font_family.clone(),
                            font_size,
                            bold,
                            italic,
                            underline: tr.properties.underline,
                            char_spacing_pt,
                        },
                        color: tr.properties.color,
                        shading: if is_space {
                            None
                        } else {
                            tr.properties.shading
                        },
                        measured_width,
                        measured_height: line_height,
                        ascent,
                        hyperlink_url: if is_space {
                            None
                        } else {
                            tr.hyperlink_url.clone()
                        },
                        baseline_offset,
                    });
                }
            }
            Inline::Image(img) if image_cache.contains(&img.rel_id) => {
                let size = scale_to_fit(img.size, content_width, content_height);
                fragments.push(Fragment::Image {
                    size,
                    rel_id: img.rel_id.to_string(),
                });
            }
            Inline::Tab => {
                let default_size = Pt::from(defaults.font_size);
                let lh = measurer
                    .font(&defaults.font_family, default_size, false, false)
                    .metrics()
                    .line_height;
                fragments.push(Fragment::Tab { line_height: lh });
            }
            Inline::LineBreak => {
                let default_size = Pt::from(defaults.font_size);
                let lh = measurer
                    .font(&defaults.font_family, default_size, false, false)
                    .metrics()
                    .line_height;
                fragments.push(Fragment::LineBreak { line_height: lh });
            }
            Inline::Image(_) => {}
            Inline::Field(fc) => {
                let text = match (&fc.field_type, field_ctx) {
                    (FieldType::Page, Some(ctx)) => ctx.page_number.to_string(),
                    (FieldType::NumPages, Some(ctx)) => ctx.num_pages.to_string(),
                    _ => "?".to_string(),
                };
                let rp = &fc.properties;
                let font_family = rp
                    .font_family
                    .clone()
                    .unwrap_or_else(|| defaults.font_family.clone());
                let font_size = Pt::from(rp.font_size.unwrap_or(defaults.font_size));
                let bold = rp.bold;
                let italic = rp.italic;
                let char_spacing_pt = rp.char_spacing.map(Pt::from).unwrap_or(Pt::ZERO);
                let f = measurer.font(&font_family, font_size, bold, italic);
                let w = f.measure_width(&text);
                let char_count = text.chars().count() as f32;
                let measured_width = w + char_spacing_pt * char_count;
                let fm = f.metrics();
                let lh = fm.line_height;
                let ascent = fm.ascent;
                fragments.push(Fragment::Text {
                    text,
                    font: FontProps {
                        font_family,
                        font_size,
                        bold,
                        italic,
                        underline: rp.underline,
                        char_spacing_pt,
                    },
                    color: rp.color,
                    shading: rp.shading,
                    measured_width,
                    measured_height: lh,
                    ascent,
                    hyperlink_url: None,
                    baseline_offset: Pt::ZERO,
                });
            }
        }
    }
    fragments
}

/// Find the next tab stop position (in points, relative to paragraph left edge).
pub fn find_next_tab_stop(current_x: Pt, custom_stops: &[TabStop], default_interval: Pt) -> Pt {
    for stop in custom_stops {
        let pos = Pt::from(stop.position);
        if pos > current_x + Pt::new(1.0) {
            return pos;
        }
    }
    if default_interval > Pt::ZERO {
        let next_multiple = ((current_x / default_interval).floor() + 1.0) * default_interval;
        return next_multiple;
    }
    current_x + TAB_FALLBACK
}

/// Scale an image to fit within max_width × max_height, preserving aspect ratio.
fn scale_to_fit(size: PtSize, max_width: Pt, max_height: Pt) -> PtSize {
    let scale = f32::min(
        1.0,
        f32::min(
            max_width / size.width.max(Pt::new(1.0)),
            max_height / size.height.max(Pt::new(1.0)),
        ),
    );
    PtSize::new(size.width * scale, size.height * scale)
}

/// Split text into segments that can break at spaces and after hyphens.
/// E.g., "Funktions-kleinspannungs-Stromkreise" → ["Funktions-", "kleinspannungs-", "Stromkreise"]
/// and "RCD Messung" → ["RCD", " ", "Messung"]
fn split_words_and_spaces(text: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0;
    let bytes = text.as_bytes();

    for (i, &b) in bytes.iter().enumerate() {
        if b == b' ' {
            // Split before space: emit word, then space segment
            if start < i {
                parts.push(&text[start..i]);
            }
            // Find end of space run
            let space_end = bytes[i..]
                .iter()
                .position(|&c| c != b' ')
                .map(|p| i + p)
                .unwrap_or(bytes.len());
            parts.push(&text[i..space_end]);
            start = space_end;
        } else if b == b'-' && i + 1 < bytes.len() && bytes[i + 1] != b' ' && start < i + 1 {
            // Split after hyphen: "word-" becomes a breakable fragment
            parts.push(&text[start..i + 1]);
            start = i + 1;
        }
    }
    if start < text.len() {
        parts.push(&text[start..]);
    }
    parts
}

/// Resolve line height from fragment height and spacing rule.
pub fn resolve_line_height(frag_height: Pt, spacing: Option<LineSpacing>) -> Pt {
    match spacing {
        Some(LineSpacing::Multiplier(m)) => frag_height * m,
        Some(LineSpacing::Fixed(pt)) => pt,
        Some(LineSpacing::AtLeast(pt)) => frag_height.max(pt),
        None => frag_height,
    }
}

pub fn measure_fragments(fragments: &[Fragment]) -> Pt {
    fragments.iter().map(|f| f.width()).sum()
}

/// Find how many fragments fit within the available width.
/// Trailing space fragments at a line break are included in the count
/// but not in the returned width.
pub fn fit_fragments(fragments: &[Fragment], available_width: Pt) -> (usize, Pt) {
    let mut total_width = Pt::ZERO;
    let mut last_break_point = 0;
    let mut width_at_break = Pt::ZERO;

    for (i, frag) in fragments.iter().enumerate() {
        if frag.is_line_break() {
            return (i + 1, total_width);
        }

        let w = frag.width();
        let is_space = matches!(frag, Fragment::Text { ref text, .. } if text.trim().is_empty());

        if total_width + w > available_width && i > 0 {
            if last_break_point > 0 {
                return (last_break_point, width_at_break);
            }
            return (i, total_width);
        }

        total_width += w;

        if is_space {
            last_break_point = i + 1;
            width_at_break = total_width - w;
        } else if matches!(frag, Fragment::Text { ref text, .. } if text.ends_with('-')) {
            // Allow breaking after hyphens (hyphen stays on the line)
            last_break_point = i + 1;
            width_at_break = total_width;
        }
    }
    (fragments.len(), total_width)
}

// ============================================================
// Shared measure→paint infrastructure for paragraphs
// ============================================================

/// A single laid-out line with draw commands at positions relative to an origin.
/// The y-coordinates in commands are relative: 0.0 = top of the paragraph/cell content.
pub struct MeasuredLine {
    pub commands: Vec<DrawCommand>,
    pub height: Pt,
}

/// Result of measuring paragraph text content (lines only, no spacing or floats).
pub struct MeasuredLines {
    pub lines: Vec<MeasuredLine>,
    pub total_height: Pt,
}

/// Measure paragraph fragments into lines with relative-positioned draw commands.
///
/// All y-coordinates in the returned commands are relative to 0.0.
pub fn measure_lines(
    fragments: &[Fragment],
    x_origin: Pt,
    available_width: Pt,
    first_line_offset: Pt,
    alignment: Option<Alignment>,
    line_spacing: Option<LineSpacing>,
    tab_stops: &[TabStop],
    default_tab_stop_pt: Pt,
    image_cache: &super::ImageCache,
    float_adjuster: Option<&dyn Fn(Pt, Pt) -> (Pt, Pt)>,
) -> MeasuredLines {
    let breaks = break_into_lines(
        fragments,
        available_width,
        first_line_offset,
        line_spacing,
        float_adjuster,
    );

    let mut lines = Vec::with_capacity(breaks.len());
    let mut cursor_y = Pt::ZERO;

    for lb in &breaks {
        cursor_y += lb.height;
        let commands = paint_line(
            &fragments[lb.start..lb.end],
            x_origin + lb.first_line_offset + lb.float_x_shift,
            cursor_y,
            lb.height,
            lb.available_width,
            alignment,
            tab_stops,
            default_tab_stop_pt,
            image_cache,
        );
        lines.push(MeasuredLine {
            commands,
            height: lb.height,
        });
    }

    MeasuredLines {
        total_height: cursor_y,
        lines,
    }
}

// ============================================================
// Line breaking (decides WHICH fragments go on which line)
// ============================================================

/// A resolved line break: fragment range, height, and layout adjustments.
struct LineBreakInfo {
    start: usize,
    end: usize,
    height: Pt,
    available_width: Pt,
    first_line_offset: Pt,
    float_x_shift: Pt,
}

/// Break fragments into lines based on available width, spacing, and float adjustments.
fn break_into_lines(
    fragments: &[Fragment],
    available_width: Pt,
    first_line_offset: Pt,
    line_spacing: Option<LineSpacing>,
    float_adjuster: Option<&dyn Fn(Pt, Pt) -> (Pt, Pt)>,
) -> Vec<LineBreakInfo> {
    let mut breaks = Vec::new();
    let mut cursor_y = Pt::ZERO;
    let mut line_start = 0;
    let mut first_line = true;

    while line_start < fragments.len() {
        // Skip leading spaces at the start of continuation lines
        if !first_line {
            while line_start < fragments.len() {
                if let Fragment::Text { ref text, .. } = fragments[line_start] {
                    if text.trim().is_empty() {
                        line_start += 1;
                        continue;
                    }
                }
                break;
            }
            if line_start >= fragments.len() {
                break;
            }
        }

        let fl_offset = if first_line {
            first_line_offset
        } else {
            Pt::ZERO
        };

        // Float adjustment: optionally narrow width based on overlapping floats
        let (float_x_shift, float_width_reduction) = if let Some(adjuster) = float_adjuster {
            let tentative_height = fragments[line_start..]
                .first()
                .map(|f| f.height())
                .unwrap_or(Pt::ZERO);
            let tentative_line_height = resolve_line_height(tentative_height, line_spacing);
            adjuster(cursor_y, cursor_y + tentative_line_height)
        } else {
            (Pt::ZERO, Pt::ZERO)
        };

        let line_avail = (available_width - fl_offset - float_width_reduction).max(Pt::ZERO);

        let (line_end, _) = fit_fragments(&fragments[line_start..], line_avail);
        let actual_end = line_start + line_end.max(1);

        let frag_height = fragments[line_start..actual_end]
            .iter()
            .map(|f| f.height())
            .fold(Pt::ZERO, Pt::max);
        let line_height = resolve_line_height(frag_height, line_spacing);

        cursor_y += line_height;

        breaks.push(LineBreakInfo {
            start: line_start,
            end: actual_end,
            height: line_height,
            available_width: line_avail,
            first_line_offset: fl_offset,
            float_x_shift,
        });
        line_start = actual_end;
        first_line = false;
    }

    breaks
}

// ============================================================
// Line painting (converts fragments into positioned DrawCommands)
// ============================================================

/// Convert a line's fragments into positioned draw commands.
///
/// `x_start` is the left edge (already includes first-line offset and float shift).
/// `cursor_y` is the bottom of this line (relative to paragraph top).
/// `line_height` is the height of this line.
/// `available_width` is the width available for alignment calculations.
fn paint_line(
    line_fragments: &[Fragment],
    x_start: Pt,
    cursor_y: Pt,
    line_height: Pt,
    available_width: Pt,
    alignment: Option<Alignment>,
    tab_stops: &[TabStop],
    default_tab_stop_pt: Pt,
    image_cache: &super::ImageCache,
) -> Vec<DrawCommand> {
    let used_width = measure_fragments(line_fragments);
    let x_offset = match alignment {
        Some(Alignment::Center) => (available_width - used_width) / 2.0,
        Some(Alignment::Right) => available_width - used_width,
        _ => Pt::ZERO,
    };

    let x_origin = x_start;
    let mut commands = Vec::new();
    let mut x = x_origin + x_offset;
    let line_top = cursor_y - line_height;

    for frag in line_fragments {
        match frag {
            Fragment::Text {
                text,
                font,
                color,
                shading,
                measured_width,
                ascent,
                hyperlink_url,
                baseline_offset,
                ..
            } => {
                let c = color.unwrap_or(Color::BLACK);
                let baseline_y = line_top + *ascent;
                if let Some(bg) = shading {
                    commands.push(DrawCommand::Rect {
                        rect: PtRect::from_xywh(x, line_top, *measured_width, line_height),
                        color: *bg,
                    });
                }
                commands.push(DrawCommand::Text {
                    position: PtOffset::new(x, baseline_y + *baseline_offset),
                    text: text.clone(),
                    font_family: font.font_family.clone(),
                    char_spacing_pt: font.char_spacing_pt,
                    font_size: font.font_size,
                    bold: font.bold,
                    italic: font.italic,
                    color: c,
                });
                if font.underline {
                    let uw = underline_width(font.font_size, font.bold);
                    let underline_y = baseline_y + UNDERLINE_Y_OFFSET;
                    commands.push(DrawCommand::Underline {
                        line: PtLineSegment::new(
                            PtOffset::new(x, underline_y),
                            PtOffset::new(x + *measured_width, underline_y),
                        ),
                        color: c,
                        width: uw,
                    });
                }
                if let Some(url) = hyperlink_url {
                    commands.push(DrawCommand::LinkAnnotation {
                        rect: PtRect::from_xywh(x, line_top, *measured_width, line_height),
                        url: url.clone(),
                    });
                }
                x += *measured_width;
            }
            Fragment::Image { size, rel_id } => {
                let image = image_cache.get(rel_id);
                commands.push(DrawCommand::Image {
                    rect: PtRect::from_xywh(x, cursor_y - size.height, size.width, size.height),
                    image,
                });
                x += size.width;
            }
            Fragment::Tab { .. } => {
                let rel_x = x - x_origin;
                let next_stop = find_next_tab_stop(rel_x, tab_stops, default_tab_stop_pt);
                x = x_origin + next_stop;
            }
            Fragment::LineBreak { .. } => {}
        }
    }

    commands
}
