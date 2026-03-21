use std::rc::Rc;

use crate::model::*;
use crate::units::*;

use super::measurer::TextMeasurer;
use super::DrawCommand;

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
        /// Hyperlink URL, if this text is part of a link.
        hyperlink_url: Option<String>,
        /// Baseline offset in points (negative = up for superscript, positive = down for subscript).
        baseline_offset: f32,
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

/// Context for evaluating field codes during fragment collection.
pub struct FieldContext {
    pub page_number: u32,
    pub num_pages: u32,
}

pub fn collect_fragments(
    runs: &[Inline],
    content_width: f32,
    content_height: f32,
    defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
) -> Vec<Fragment> {
    collect_fragments_with_fields(runs, content_width, content_height, defaults, measurer, None)
}

pub fn collect_fragments_with_fields(
    runs: &[Inline],
    content_width: f32,
    content_height: f32,
    defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    field_ctx: Option<&FieldContext>,
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
                let base_font_size = tr
                    .properties
                    .font_size_pt_with_default(defaults.font_size_half_pts);
                let bold = tr.properties.bold;
                let italic = tr.properties.italic;
                let char_spacing_pt = tr.properties.char_spacing
                    .map(|cs| twips_to_pt_signed(cs))
                    .unwrap_or(0.0);

                // Super/subscript: reduce font size and compute baseline offset
                let (font_size, baseline_offset) = match tr.properties.vert_align {
                    Some(VertAlign::Superscript) => {
                        let reduced = base_font_size * 0.58;
                        // Shift up by ~33% of the original line height
                        let offset = -(base_font_size * 0.33);
                        (reduced, offset)
                    }
                    Some(VertAlign::Subscript) => {
                        let reduced = base_font_size * 0.58;
                        // Shift down by ~8% of the original line height
                        let offset = base_font_size * 0.08;
                        (reduced, offset)
                    }
                    None => (base_font_size, 0.0),
                };

                let line_height =
                    measurer.line_height(&font_family, base_font_size, bold, italic);
                for part in split_words_and_spaces(&tr.text) {
                    let base_width = measurer.measure_width(
                        part,
                        &font_family,
                        font_size,
                        bold,
                        italic,
                    );
                    // Character spacing expands each character's advance width
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
                        hyperlink_url: if is_space { None } else { tr.hyperlink_url.clone() },
                        baseline_offset,
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
            Inline::Field(fc) => {
                let text = match (&fc.field_type, field_ctx) {
                    (FieldType::Page, Some(ctx)) => ctx.page_number.to_string(),
                    (FieldType::NumPages, Some(ctx)) => ctx.num_pages.to_string(),
                    _ => "?".to_string(),
                };
                let rp = &fc.properties;
                let font_family = rp.font_family.clone()
                    .unwrap_or_else(|| defaults.font_family.clone());
                let font_size = rp.font_size_pt_with_default(defaults.font_size_half_pts);
                let bold = rp.bold;
                let italic = rp.italic;
                let char_spacing_pt = rp.char_spacing
                    .map(|cs| twips_to_pt_signed(cs))
                    .unwrap_or(0.0);
                let w = measurer.measure_width(&text, &font_family, font_size, bold, italic);
                let char_count = text.chars().count() as f32;
                let measured_width = w + char_spacing_pt * char_count;
                let lh = measurer.line_height(&font_family, font_size, bold, italic);
                fragments.push(Fragment::Text {
                    text,
                    font_family,
                    font_size,
                    bold,
                    italic,
                    underline: rp.underline,
                    color: rp.color,
                    shading: rp.shading,
                    char_spacing_pt,
                    measured_width,
                    measured_height: lh,
                    hyperlink_url: None,
                    baseline_offset: 0.0,
                });
            }
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

// ============================================================
// Shared measure→paint infrastructure for paragraphs
// ============================================================

/// A single laid-out line with draw commands at positions relative to an origin.
/// The y-coordinates in commands are relative: 0.0 = top of the paragraph/cell content.
pub struct MeasuredLine {
    pub commands: Vec<DrawCommand>,
    pub height: f32,
}

/// Result of measuring paragraph text content (lines only, no spacing or floats).
pub struct MeasuredLines {
    pub lines: Vec<MeasuredLine>,
    pub total_height: f32,
}

/// Measure paragraph runs into lines with relative-positioned draw commands.
///
/// This is the single source of truth for converting fragments into draw commands.
/// Used by: body paragraphs, table cell paragraphs, header/footer paragraphs.
///
/// `x_origin` is the left edge for content (e.g., margin_left + indent.left).
/// `available_width` is the width for line fitting (content_width - indents).
/// All y-coordinates in the returned commands are relative to 0.0.
pub fn measure_lines(
    fragments: &[Fragment],
    x_origin: f32,
    available_width: f32,
    first_line_offset: f32,
    alignment: Option<Alignment>,
    line_spacing: Option<LineSpacing>,
    tab_stops: &[TabStop],
    default_tab_stop_pt: f32,
) -> MeasuredLines {
    let mut lines = Vec::new();
    let mut cursor_y = 0.0_f32;
    let mut line_start = 0;
    let mut first_line = true;

    while line_start < fragments.len() {
        // Skip leading space fragments at the start of a new line
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

        let fl_offset = if first_line { first_line_offset } else { 0.0 };
        let line_avail = (available_width - fl_offset).max(0.0);

        let (line_end, _) = fit_fragments(&fragments[line_start..], line_avail);
        let actual_end = line_start + line_end.max(1);

        let frag_height = fragments[line_start..actual_end]
            .iter()
            .map(|f| f.height())
            .fold(0.0_f32, f32::max);
        let line_height = resolve_line_height(frag_height, line_spacing);

        let used_width = if actual_end > line_start {
            measure_fragments(&fragments[line_start..actual_end])
        } else {
            0.0
        };
        let x_offset = match alignment {
            Some(Alignment::Center) => (line_avail - used_width) / 2.0,
            Some(Alignment::Right) => line_avail - used_width,
            _ => 0.0,
        };

        cursor_y += line_height;
        let mut commands = Vec::new();
        let mut x = x_origin + fl_offset + x_offset;

        for frag in &fragments[line_start..actual_end] {
            match frag {
                Fragment::Text {
                    text, font_family, font_size, bold, italic,
                    underline, color, shading, char_spacing_pt,
                    measured_width, hyperlink_url, baseline_offset, ..
                } => {
                    let c = color.map(|c| (c.r, c.g, c.b)).unwrap_or((0, 0, 0));
                    if let Some(bg) = shading {
                        commands.push(DrawCommand::Rect {
                            x,
                            y: cursor_y - line_height,
                            width: *measured_width,
                            height: line_height,
                            color: (bg.r, bg.g, bg.b),
                        });
                    }
                    commands.push(DrawCommand::Text {
                        x,
                        y: cursor_y + baseline_offset,
                        text: text.clone(),
                        font_family: font_family.clone(),
                        char_spacing_pt: *char_spacing_pt,
                        font_size: *font_size,
                        bold: *bold,
                        italic: *italic,
                        color: c,
                    });
                    if *underline {
                        commands.push(DrawCommand::Underline {
                            x1: x,
                            y1: cursor_y + UNDERLINE_Y_OFFSET,
                            x2: x + measured_width,
                            y2: cursor_y + UNDERLINE_Y_OFFSET,
                            color: c,
                            width: UNDERLINE_STROKE_WIDTH,
                        });
                    }
                    if let Some(url) = hyperlink_url {
                        commands.push(DrawCommand::LinkAnnotation {
                            x,
                            y: cursor_y - line_height,
                            width: *measured_width,
                            height: line_height,
                            url: url.clone(),
                        });
                    }
                    x += measured_width;
                }
                Fragment::Image { width, height, data } => {
                    commands.push(DrawCommand::Image {
                        x,
                        y: cursor_y - height,
                        width: *width,
                        height: *height,
                        data: data.clone(),
                    });
                    x += width;
                }
                Fragment::Tab { .. } => {
                    let rel_x = x - x_origin;
                    let next_stop = find_next_tab_stop(rel_x, tab_stops, default_tab_stop_pt);
                    x = x_origin + next_stop;
                }
                Fragment::LineBreak { .. } => {}
            }
        }

        lines.push(MeasuredLine { commands, height: line_height });
        line_start = actual_end;
        first_line = false;
    }

    MeasuredLines { total_height: cursor_y, lines }
}
