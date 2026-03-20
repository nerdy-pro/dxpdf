use crate::model::*;

use super::measurer::TextMeasurer;

/// Document-level defaults for layout.
pub struct DocDefaultsLayout {
    pub font_size_half_pts: u32,
    pub font_family: String,
    pub default_spacing: Spacing,
}

/// A flattened fragment for layout — either text, an image, a tab, or a line break.
pub enum Fragment {
    Text {
        text: String,
        font_family: String,
        font_size: f32,
        bold: bool,
        italic: bool,
        underline: bool,
        color: Option<Color>,
        measured_width: f32,
        measured_height: f32,
    },
    Image {
        width: f32,
        height: f32,
        data: Vec<u8>,
    },
    Tab {
        line_height: f32,
    },
    LineBreak {
        line_height: f32,
    },
}

impl Fragment {
    pub fn width(&self) -> f32 {
        match self {
            Fragment::Text { measured_width, .. } => *measured_width,
            Fragment::Image { width, .. } => *width,
            Fragment::Tab { .. } => 12.0,
            Fragment::LineBreak { .. } => 0.0,
        }
    }

    pub fn height(&self) -> f32 {
        match self {
            Fragment::Text { measured_height, .. } => *measured_height,
            Fragment::Image { height, .. } => *height,
            Fragment::Tab { line_height } | Fragment::LineBreak { line_height } => *line_height,
        }
    }

    pub fn is_line_break(&self) -> bool {
        matches!(self, Fragment::LineBreak { .. })
    }
}

pub fn collect_fragments(
    runs: &[Inline],
    content_width: f32,
    content_height: f32,
    defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
) -> Vec<Fragment> {
    let mut fragments = Vec::new();
    for run in runs {
        match run {
            Inline::TextRun(tr) => {
                let collapsed = collapse_spaces(&tr.text);
                let font_family = tr
                    .properties
                    .font_family
                    .clone()
                    .unwrap_or_else(|| defaults.font_family.clone());
                let font_size = tr
                    .properties
                    .font_size_pt_with_default(defaults.font_size_half_pts);
                let bold = tr.properties.bold;
                let italic = tr.properties.italic;
                let line_height =
                    measurer.line_height(&font_family, font_size, bold, italic);
                for part in split_words_and_spaces(&collapsed) {
                    let measured_width = measurer.measure_width(
                        part,
                        &font_family,
                        font_size,
                        bold,
                        italic,
                    );
                    let is_space = part.chars().all(|c| c == ' ');
                    fragments.push(Fragment::Text {
                        text: part.to_string(),
                        font_family: font_family.clone(),
                        font_size,
                        bold,
                        italic,
                        underline: !is_space && tr.properties.underline,
                        color: tr.properties.color,
                        measured_width,
                        measured_height: line_height,
                    });
                }
            }
            Inline::Image(img) if !img.data.is_empty() => {
                let scale = f32::min(
                    1.0,
                    f32::min(
                        content_width / img.width_pt.max(1.0),
                        content_height / img.height_pt.max(1.0),
                    ),
                );
                fragments.push(Fragment::Image {
                    width: img.width_pt * scale,
                    height: img.height_pt * scale,
                    data: img.data.clone(),
                });
            }
            Inline::Tab => {
                let default_size = defaults.font_size_half_pts as f32 / 2.0;
                let lh = measurer.line_height(
                    &defaults.font_family,
                    default_size,
                    false,
                    false,
                );
                fragments.push(Fragment::Tab { line_height: lh });
            }
            Inline::LineBreak => {
                let default_size = defaults.font_size_half_pts as f32 / 2.0;
                let lh = measurer.line_height(
                    &defaults.font_family,
                    default_size,
                    false,
                    false,
                );
                fragments.push(Fragment::LineBreak { line_height: lh });
            }
            Inline::Image(_) => {}
        }
    }
    fragments
}

/// Find the next tab stop position (in points, relative to paragraph left edge).
pub fn find_next_tab_stop(
    current_x: f32,
    custom_stops: &[TabStop],
    default_interval: f32,
) -> f32 {
    for stop in custom_stops {
        let pos = stop.position_pt();
        if pos > current_x + 1.0 {
            return pos;
        }
    }
    if default_interval > 0.0 {
        let next_multiple =
            ((current_x / default_interval).floor() + 1.0) * default_interval;
        return next_multiple;
    }
    current_x + 36.0
}

/// Split text into alternating word and space segments.
fn split_words_and_spaces(text: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0;
    let bytes = text.as_bytes();
    let mut in_space = bytes.first() == Some(&b' ');

    for (i, &b) in bytes.iter().enumerate() {
        let is_space = b == b' ';
        if is_space != in_space {
            if start < i {
                parts.push(&text[start..i]);
            }
            start = i;
            in_space = is_space;
        }
    }
    if start < text.len() {
        parts.push(&text[start..]);
    }
    parts
}

/// Collapse runs of more than 2 consecutive spaces into a single space.
fn collapse_spaces(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    let mut space_count = 0;
    for ch in text.chars() {
        if ch == ' ' {
            space_count += 1;
            if space_count <= 2 {
                result.push(ch);
            }
        } else {
            space_count = 0;
            result.push(ch);
        }
    }
    result
}

pub fn measure_fragments(fragments: &[Fragment]) -> f32 {
    fragments.iter().map(|f| f.width()).sum()
}

/// Find how many fragments fit within the available width.
/// Trailing space fragments at a line break are included in the count
/// but not in the returned width.
pub fn fit_fragments(fragments: &[Fragment], available_width: f32) -> (usize, f32) {
    let mut total_width = 0.0;
    let mut last_break_point = 0;
    let mut width_at_break = 0.0;

    for (i, frag) in fragments.iter().enumerate() {
        if frag.is_line_break() {
            return (i + 1, total_width);
        }

        let w = frag.width();
        let is_space =
            matches!(frag, Fragment::Text { ref text, .. } if text.trim().is_empty());

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
        }
    }
    (fragments.len(), total_width)
}
