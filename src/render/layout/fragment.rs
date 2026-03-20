use std::rc::Rc;

use crate::model::*;
use crate::units::*;

use super::measurer::TextMeasurer;

/// Document-level defaults for layout.
pub struct DocDefaultsLayout {
    pub font_size_half_pts: u32,
    pub font_family: Rc<str>,
    pub default_spacing: Spacing,
    pub default_cell_margins: CellMargins,
    pub table_cell_spacing: Spacing,
    pub default_table_borders: TableBorders,
    pub default_header: Option<HeaderFooter>,
    pub default_footer: Option<HeaderFooter>,
    pub numbering: NumberingMap,
}

impl DocDefaultsLayout {
    pub fn from_document(doc: &crate::model::Document) -> Self {
        Self {
            font_size_half_pts: doc.default_font_size,
            font_family: Rc::clone(&doc.default_font_family),
            default_spacing: doc.default_spacing,
            default_cell_margins: doc.default_cell_margins,
            table_cell_spacing: doc.table_cell_spacing,
            default_table_borders: doc.default_table_borders,
            default_header: doc.default_header.clone(),
            default_footer: doc.default_footer.clone(),
            numbering: doc.numbering.clone(),
        }
    }
}

/// A flattened fragment for layout — either text, an image, a tab, or a line break.
pub enum Fragment {
    Text {
        text: String,
        font_family: Rc<str>,
        font_size: f32,
        bold: bool,
        italic: bool,
        underline: bool,
        color: Option<Color>,
        shading: Option<Color>,
        /// Character spacing in points (positive = expand).
        char_spacing_pt: f32,
        measured_width: f32,
        measured_height: f32,
    },
    Image {
        width: f32,
        height: f32,
        data: ImageData,
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
            Fragment::Tab { .. } => MIN_TAB_WIDTH_PT,
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
                let char_spacing_pt = tr.properties.char_spacing
                    .map(|cs| twips_to_pt_signed(cs))
                    .unwrap_or(0.0);
                let line_height =
                    measurer.line_height(&font_family, font_size, bold, italic);
                for part in split_words_and_spaces(&tr.text) {
                    let base_width = measurer.measure_width(
                        part,
                        &font_family,
                        font_size,
                        bold,
                        italic,
                    );
                    // Character spacing adds extra width per character
                    let char_count = part.chars().count() as f32;
                    let measured_width = base_width + char_spacing_pt * char_count;
                    let is_space = part.chars().all(|c| c == ' ');
                    fragments.push(Fragment::Text {
                        text: part.to_string(),
                        font_family: font_family.clone(),
                        font_size,
                        bold,
                        italic,
                        underline: !is_space && tr.properties.underline,
                        color: tr.properties.color,
                        shading: if is_space { None } else { tr.properties.shading },
                        char_spacing_pt,
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
                let default_size = defaults.font_size_half_pts as f32 / HALF_POINTS_PER_POINT;
                let lh = measurer.line_height(
                    &defaults.font_family,
                    default_size,
                    false,
                    false,
                );
                fragments.push(Fragment::Tab { line_height: lh });
            }
            Inline::LineBreak => {
                let default_size = defaults.font_size_half_pts as f32 / HALF_POINTS_PER_POINT;
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
    current_x + TAB_FALLBACK_PT
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
pub fn resolve_line_height(frag_height: f32, spacing: Option<LineSpacing>) -> f32 {
    match spacing {
        Some(LineSpacing::Multiplier(m)) => frag_height * m,
        Some(LineSpacing::Fixed(pt)) => pt,
        Some(LineSpacing::AtLeast(pt)) => frag_height.max(pt),
        None => frag_height,
    }
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
        } else if matches!(frag, Fragment::Text { ref text, .. } if text.ends_with('-')) {
            // Allow breaking after hyphens (hyphen stays on the line)
            last_break_point = i + 1;
            width_at_break = total_width;
        }
    }
    (fragments.len(), total_width)
}
