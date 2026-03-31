//! Fragment conversion — transform Inline content into measured Fragments
//! for the line-fitting algorithm.

use std::rc::Rc;

use crate::model::{Block, FieldCharType, Inline, RunElement, RunProperties, VerticalAlign};

use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::resolve::color::RgbColor;
use crate::render::resolve::fonts::effective_font;

/// Font properties needed for rendering a text fragment.
#[derive(Clone, Debug)]
pub struct FontProps {
    pub family: Rc<str>,
    pub size: Pt,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub char_spacing: Pt,
    /// Underline position from font metrics (positive = below baseline).
    pub underline_position: Pt,
    /// Underline thickness from font metrics.
    pub underline_thickness: Pt,
}

/// Font metrics for a specific font at a specific size.
/// Evaluated once by the measurer and carried through the pipeline.
#[derive(Clone, Copy, Debug)]
pub struct TextMetrics {
    /// Distance from baseline to top of glyphs (positive upward).
    pub ascent: Pt,
    /// Distance from baseline to bottom of glyphs (positive downward).
    pub descent: Pt,
}

impl TextMetrics {
    /// Total text height (ascent + descent).
    pub fn height(&self) -> Pt {
        self.ascent + self.descent
    }
}

/// §17.3.2.4: run-level border for rendering.
#[derive(Clone, Copy, Debug)]
pub struct FragmentBorder {
    pub width: Pt,
    pub color: RgbColor,
    pub space: Pt,
}

/// A measured fragment — the atomic unit for line fitting.
#[derive(Clone, Debug)]
pub enum Fragment {
    Text {
        text: String,
        font: FontProps,
        color: RgbColor,
        /// §17.3.2.32: run-level shading (background color behind text).
        shading: Option<RgbColor>,
        /// §17.3.2.4: run-level border (box around text).
        border: Option<FragmentBorder>,
        /// Full width including trailing whitespace (used for positioning).
        width: Pt,
        /// Width excluding trailing whitespace (used for line-break overflow checking).
        /// Trailing whitespace is allowed to hang past the margin per Word behavior.
        trimmed_width: Pt,
        /// Font metrics (ascent + descent = text height).
        metrics: TextMetrics,
        hyperlink_url: Option<String>,
        baseline_offset: Pt,
        /// Horizontal offset for drawing text within the fragment width.
        /// Used for right/center-justified list labels where the text is
        /// positioned within a wider fragment. Default: Pt::ZERO.
        text_offset: Pt,
    },
    Image {
        size: PtSize,
        rel_id: String,
        image_data: Option<std::rc::Rc<[u8]>>,
    },
    Tab {
        line_height: Pt,
        /// Override minimum width for line fitting (default: MIN_TAB_WIDTH).
        fitting_width: Option<Pt>,
    },
    LineBreak {
        line_height: Pt,
    },
    /// §17.3.3.1: column break — forces content to the next column.
    ColumnBreak,
    /// Named destination (bookmark target) — zero-width marker.
    Bookmark {
        name: String,
    },
}

impl Fragment {
    pub fn width(&self) -> Pt {
        match self {
            Fragment::Text { width, .. } => *width,
            Fragment::Image { size, .. } => size.width,
            Fragment::Tab { fitting_width, .. } => fitting_width.unwrap_or(MIN_TAB_WIDTH),
            Fragment::LineBreak { .. } | Fragment::ColumnBreak | Fragment::Bookmark { .. } => {
                Pt::ZERO
            }
        }
    }

    /// Width for overflow checking — excludes trailing whitespace on text fragments.
    pub fn trimmed_width(&self) -> Pt {
        match self {
            Fragment::Text { trimmed_width, .. } => *trimmed_width,
            other => other.width(),
        }
    }

    pub fn height(&self) -> Pt {
        match self {
            Fragment::Text { metrics, .. } => metrics.height(),
            Fragment::Image { size, .. } => size.height,
            Fragment::Tab { line_height, .. } | Fragment::LineBreak { line_height } => *line_height,
            Fragment::ColumnBreak | Fragment::Bookmark { .. } => Pt::ZERO,
        }
    }

    pub fn is_line_break(&self) -> bool {
        matches!(self, Fragment::LineBreak { .. } | Fragment::ColumnBreak)
    }

    /// Get font properties if this is a text fragment.
    pub fn font_props(&self) -> Option<&FontProps> {
        match self {
            Fragment::Text { font, .. } => Some(font),
            _ => None,
        }
    }
}

/// §17.3.1.37: minimum tab fragment width for line fitting.
/// Tabs resolve to tab stops defined on the paragraph; this constant is only
/// used as the fragment width during line breaking (actual tab position is
/// computed during paragraph layout).
pub const MIN_TAB_WIDTH: Pt = Pt::new(1.0);

/// Extract font properties from RunProperties with a default font family fallback.
pub fn font_props_from_run(
    rp: &RunProperties,
    default_family: &str,
    default_size: Pt,
) -> FontProps {
    let family = effective_font(&rp.fonts).unwrap_or(default_family);

    let size = rp.font_size.map(Pt::from).unwrap_or(default_size);

    let char_spacing = rp.spacing.map(Pt::from).unwrap_or(Pt::ZERO);

    FontProps {
        family: Rc::from(family),
        size,
        bold: rp.bold.unwrap_or(false),
        italic: rp.italic.unwrap_or(false),
        underline: rp.underline.is_some(),
        char_spacing,
        // Populated by the measurer from Skia font metrics.
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    }
}

/// §17.18.40 ST_HighlightColor: map highlight enum to RGB.
/// These are the fixed palette colors defined in the OOXML spec.
fn resolve_highlight_color(hl: crate::model::HighlightColor) -> RgbColor {
    use crate::model::HighlightColor;
    match hl {
        HighlightColor::Black => RgbColor { r: 0, g: 0, b: 0 },
        HighlightColor::Blue => RgbColor { r: 0, g: 0, b: 255 },
        HighlightColor::Cyan => RgbColor {
            r: 0,
            g: 255,
            b: 255,
        },
        HighlightColor::DarkBlue => RgbColor { r: 0, g: 0, b: 139 },
        HighlightColor::DarkCyan => RgbColor {
            r: 0,
            g: 139,
            b: 139,
        },
        HighlightColor::DarkGray => RgbColor {
            r: 169,
            g: 169,
            b: 169,
        },
        HighlightColor::DarkGreen => RgbColor { r: 0, g: 100, b: 0 },
        HighlightColor::DarkMagenta => RgbColor {
            r: 139,
            g: 0,
            b: 139,
        },
        HighlightColor::DarkRed => RgbColor { r: 139, g: 0, b: 0 },
        HighlightColor::DarkYellow => RgbColor {
            r: 139,
            g: 139,
            b: 0,
        },
        HighlightColor::Green => RgbColor { r: 0, g: 255, b: 0 },
        HighlightColor::LightGray => RgbColor {
            r: 211,
            g: 211,
            b: 211,
        },
        HighlightColor::Magenta => RgbColor {
            r: 255,
            g: 0,
            b: 255,
        },
        HighlightColor::Red => RgbColor { r: 255, g: 0, b: 0 },
        HighlightColor::White => RgbColor {
            r: 255,
            g: 255,
            b: 255,
        },
        HighlightColor::Yellow => RgbColor {
            r: 255,
            g: 255,
            b: 0,
        },
    }
}

/// Split text into word-level chunks for line breaking.
/// Whitespace is kept attached to the preceding word: "hello world" → ["hello ", "world"].
/// This allows the line fitter to break between fragments at word boundaries.
/// Convert a number to lowercase Roman numerals.
pub fn to_roman_lower(mut n: u32) -> String {
    const VALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut s = String::new();
    for &(val, sym) in &VALS {
        while n >= val {
            s.push_str(sym);
            n -= val;
        }
    }
    s
}

fn split_into_words(text: &str) -> Vec<&str> {
    let mut words = Vec::new();
    let mut start = 0;

    for (i, ch) in text.char_indices() {
        match ch {
            // Whitespace: include with the preceding word.
            ' ' | '\t' => {
                let end = i + ch.len_utf8();
                if end > start {
                    words.push(&text[start..end]);
                    start = end;
                }
            }
            // Hyphen/dash: break AFTER the hyphen (UAX #14).
            // The hyphen stays with the preceding word.
            '-' | '\u{2010}' | '\u{2011}' | '\u{2012}' | '\u{2013}' | '\u{2014}' => {
                let end = i + ch.len_utf8();
                if end > start {
                    words.push(&text[start..end]);
                    start = end;
                }
            }
            _ => {}
        }
    }

    // Remaining text (last word without trailing space)
    if start < text.len() {
        words.push(&text[start..]);
    }

    words
}

/// Split text into word-level fragments and push to the output vec.
#[allow(clippy::too_many_arguments)]
fn emit_text_fragments<F>(
    text: &str,
    font: &FontProps,
    color: RgbColor,
    shading: Option<RgbColor>,
    border: Option<FragmentBorder>,
    hyperlink_url: Option<&str>,
    measure_text: &F,
    baseline_offset: Pt,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    // §2.1 XML spec: C0 control characters (U+0000–U+001F) other than
    // HT (U+0009), LF (U+000A), CR (U+000D) are invalid in XML but some
    // producers embed LF/CR in w:t content. Strip all non-tab controls
    // so they don't render as tofu/question-mark glyphs.
    let cleaned: String = text
        .chars()
        .filter(|&c| !c.is_control() || c == '\t')
        .collect();
    if cleaned.is_empty() {
        return;
    }
    for word in split_into_words(&cleaned) {
        let (w, m) = measure_text(word, font);
        let trimmed = word.trim_end();
        let tw = if trimmed.len() < word.len() {
            measure_text(trimmed, font).0
        } else {
            w
        };
        fragments.push(Fragment::Text {
            text: word.to_string(),
            font: font.clone(),
            color,
            shading,
            border,
            width: w,
            trimmed_width: tw,
            metrics: m,
            hyperlink_url: hyperlink_url.map(String::from),
            baseline_offset,
            text_offset: Pt::ZERO,
        });
    }
}

/// §17.16.4.1: context for evaluating dynamic fields (PAGE, NUMPAGES).
#[derive(Clone, Copy, Default)]
pub struct FieldContext {
    /// Current page number (1-based).
    pub page_number: Option<usize>,
    /// Total page count in the document.
    pub num_pages: Option<usize>,
}

/// §17.16.4.1: evaluate a parsed field instruction against the current context.
/// Returns the substituted text for PAGE/NUMPAGES, or None for other fields
/// or when no context is available.
fn evaluate_field_instruction(
    instruction: &crate::field::FieldInstruction,
    ctx: FieldContext,
) -> Option<String> {
    match instruction {
        crate::field::FieldInstruction::Page { .. } => ctx.page_number.map(|n| n.to_string()),
        crate::field::FieldInstruction::NumPages { .. } => ctx.num_pages.map(|n| n.to_string()),
        _ => None,
    }
}

/// Build a text fragment for a substituted field value, using the paragraph's
/// default font properties.
fn make_field_text_fragment<F>(
    text: String,
    default_family: &str,
    default_size: Pt,
    default_color: crate::render::resolve::color::RgbColor,
    measure_text: &F,
) -> Fragment
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let font = FontProps {
        family: Rc::from(default_family),
        size: default_size,
        bold: false,
        italic: false,
        underline: false,
        char_spacing: Pt::ZERO,
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    };
    let (w, m) = measure_text(&text, &font);
    Fragment::Text {
        text,
        font,
        color: default_color,
        shading: None,
        border: None,
        width: w,
        trimmed_width: w,
        metrics: m,
        hyperlink_url: None,
        baseline_offset: Pt::ZERO,
        text_offset: Pt::ZERO,
    }
}

/// Walk inline content and collect fragments.
/// `measure_text` is a callback that measures text width/height/ascent for a given font.
/// `resolved_styles` is used to look up character styles (w:rStyle) on text runs.
///
/// Returns fragments suitable for the line-fitting algorithm.
#[allow(clippy::too_many_arguments)]
pub fn collect_fragments<F>(
    inlines: &[Inline],
    default_family: &str,
    default_size: Pt,
    default_color: RgbColor,
    hyperlink_url: Option<&str>,
    measure_text: &F,
    resolved_styles: Option<
        &std::collections::HashMap<
            crate::model::StyleId,
            crate::render::resolve::styles::ResolvedStyle,
        >,
    >,
    // §17.3.1: paragraph style's run properties, merged as base for all runs.
    paragraph_run_defaults: Option<&RunProperties>,
    footnote_counter: &mut u32,
    endnote_counter: &mut u32,
    field_ctx: FieldContext,
    theme: Option<&crate::model::Theme>,
) -> Vec<Fragment>
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics), // (width, metrics)
{
    let mut fragments = Vec::new();
    let mut field_depth: i32 = 0; // tracks nested complex field state
    let mut field_instr = String::new(); // accumulated instruction text for current complex field
                                         // §17.16.19: field substitution state for complex fields.
                                         // Pending = substitution text waiting for the first result TextRun's formatting.
                                         // Emitted = substitution was rendered, skip remaining result TextRuns until End.
    let mut field_sub_pending: Option<String> = None;
    let mut field_sub_emitted = false;
    for inline in inlines {
        match inline {
            Inline::TextRun(tr) => {
                // Skip field instruction text (between Begin and Separate).
                if field_depth > 0 {
                    continue;
                }
                // Skip remaining result runs after substitution was emitted.
                if field_sub_emitted {
                    continue;
                }
                // Run property cascade per §17.7.2:
                // direct run properties → character style (w:rStyle) → paragraph style run defaults.
                let mut effective_props = tr.properties.clone();
                // §17.3.2.26: resolve theme font references before merging,
                // so theme-derived names take precedence over explicit names
                // from lower-priority levels in the cascade.
                if let Some(th) = theme {
                    crate::render::resolve::fonts::resolve_font_set_themes(
                        &mut effective_props.fonts,
                        th,
                    );
                }
                if let (Some(ref style_id), Some(styles)) = (&tr.style_id, resolved_styles) {
                    if let Some(resolved_style) = styles.get(style_id) {
                        crate::render::resolve::properties::merge_run_properties(
                            &mut effective_props,
                            &resolved_style.run,
                        );
                    }
                }
                if let Some(para_run) = paragraph_run_defaults {
                    crate::render::resolve::properties::merge_run_properties(
                        &mut effective_props,
                        para_run,
                    );
                }

                let mut font = font_props_from_run(&effective_props, default_family, default_size);
                let color = effective_props
                    .color
                    .map(|c| {
                        crate::render::resolve::color::resolve_color(
                            c,
                            crate::render::resolve::color::ColorContext::Text,
                        )
                    })
                    .unwrap_or(default_color);
                // §17.3.2.32: run-level shading (background behind text).
                // §17.3.2.15: highlight color (fixed palette) takes effect when shading is absent.
                let shading = effective_props
                    .shading
                    .as_ref()
                    .map(|s| {
                        crate::render::resolve::color::resolve_color(
                            s.fill,
                            crate::render::resolve::color::ColorContext::Background,
                        )
                    })
                    .or_else(|| effective_props.highlight.map(resolve_highlight_color));

                // §17.3.2.42: vertical alignment (superscript/subscript).
                // The spec states these are "application-defined" — the values below
                // match Word's rendering: 58% font size reduction, superscript shifted
                // up by 33% of base ascent, subscript shifted down by 8% of base height.
                // These ratios are documented in the OpenXML SDK reference.
                let mut baseline_offset = match effective_props.vertical_align {
                    Some(VerticalAlign::Superscript) => {
                        let (_, base_m) = measure_text("X", &font);
                        font.size = font.size * 0.58;
                        -(base_m.ascent * 0.33)
                    }
                    Some(VerticalAlign::Subscript) => {
                        let (_, base_m) = measure_text("X", &font);
                        font.size = font.size * 0.58;
                        base_m.height() * 0.08
                    }
                    _ => Pt::ZERO,
                };
                // §17.3.2.19: w:position — vertical baseline offset in half-points.
                if let Some(pos) = effective_props.position {
                    baseline_offset += Pt::from(pos);
                }

                // §17.3.2.4: run-level border.
                let border = effective_props.border.as_ref().map(|b| FragmentBorder {
                    width: Pt::from(b.width),
                    color: crate::render::resolve::color::resolve_color(
                        b.color,
                        crate::render::resolve::color::ColorContext::Text,
                    ),
                    space: Pt::new(b.space.raw() as f32),
                });

                // §17.16.19: if a field substitution is pending, use the
                // substituted text with this TextRun's resolved formatting.
                if field_sub_pending.is_some() {
                    let sub = field_sub_pending.take().unwrap();
                    field_sub_emitted = true;
                    emit_text_fragments(
                        &sub,
                        &font,
                        color,
                        shading,
                        border,
                        hyperlink_url,
                        measure_text,
                        baseline_offset,
                        &mut fragments,
                    );
                } else {
                    for element in &tr.content {
                        match element {
                            RunElement::Text(text) => {
                                emit_text_fragments(
                                    text,
                                    &font,
                                    color,
                                    shading,
                                    border,
                                    hyperlink_url,
                                    measure_text,
                                    baseline_offset,
                                    &mut fragments,
                                );
                            }
                            RunElement::Tab => {
                                fragments.push(Fragment::Tab {
                                    line_height: font.size,
                                    fitting_width: None,
                                });
                            }
                            RunElement::LineBreak(_) => {
                                fragments.push(Fragment::LineBreak {
                                    line_height: font.size,
                                });
                            }
                            RunElement::PageBreak => {
                                fragments.push(Fragment::LineBreak {
                                    line_height: font.size,
                                });
                            }
                            RunElement::ColumnBreak => {
                                fragments.push(Fragment::ColumnBreak);
                            }
                            RunElement::LastRenderedPageBreak => {}
                        }
                    }
                }
            }
            Inline::Image(img) => {
                // Only render INLINE images as fragments.
                // Anchor (floating) images are handled separately in build.rs.
                if matches!(img.placement, crate::model::ImagePlacement::Inline { .. }) {
                    if let Some(rel_id) = crate::render::resolve::images::extract_image_rel_id(img)
                    {
                        let w = Pt::from(img.extent.width);
                        let h = Pt::from(img.extent.height);
                        fragments.push(Fragment::Image {
                            size: PtSize::new(w, h),
                            rel_id: rel_id.as_str().to_string(),
                            image_data: None,
                        });
                    }
                }
            }
            Inline::Hyperlink(link) => {
                let url: Option<&str> = match &link.target {
                    crate::model::HyperlinkTarget::External(rel_id) => Some(rel_id.as_str()),
                    crate::model::HyperlinkTarget::Internal { anchor } => Some(anchor.as_str()),
                };
                let mut sub = collect_fragments(
                    &link.content,
                    default_family,
                    default_size,
                    default_color,
                    url,
                    measure_text,
                    resolved_styles,
                    paragraph_run_defaults,
                    footnote_counter,
                    endnote_counter,
                    field_ctx,
                    theme,
                );
                fragments.append(&mut sub);
            }
            Inline::Field(field) => {
                // §17.16.18: simple field — check for dynamic substitution.
                let substituted = evaluate_field_instruction(&field.instruction, field_ctx);
                if let Some(text) = substituted {
                    fragments.push(make_field_text_fragment(
                        text,
                        default_family,
                        default_size,
                        default_color,
                        measure_text,
                    ));
                } else {
                    let mut sub = collect_fragments(
                        &field.content,
                        default_family,
                        default_size,
                        default_color,
                        hyperlink_url,
                        measure_text,
                        resolved_styles,
                        paragraph_run_defaults,
                        footnote_counter,
                        endnote_counter,
                        field_ctx,
                        theme,
                    );
                    fragments.append(&mut sub);
                }
            }
            Inline::FieldChar(fc) => {
                // §17.16.18: complex field state machine:
                // Begin → InstrText... → Separate → result runs → End
                match fc.field_char_type {
                    FieldCharType::Begin => {
                        field_depth += 1;
                        field_instr.clear();
                        field_sub_pending = None;
                        field_sub_emitted = false;
                    }
                    FieldCharType::Separate => {
                        // §17.16.4.1: parse accumulated instruction, evaluate
                        // PAGE/NUMPAGES if field context is available.
                        if let Ok(parsed) = crate::field::parse(&field_instr) {
                            field_sub_pending = evaluate_field_instruction(&parsed, field_ctx);
                        }
                        field_depth -= 1; // now collect result runs (unless substituted)
                    }
                    FieldCharType::End => {
                        // If a substitution was pending but no result TextRun
                        // was present to provide formatting, emit with defaults.
                        if let Some(text) = field_sub_pending.take() {
                            fragments.push(make_field_text_fragment(
                                text,
                                default_family,
                                default_size,
                                default_color,
                                measure_text,
                            ));
                        }
                        field_sub_emitted = false;
                    }
                }
            }
            Inline::InstrText(text) => {
                // Accumulate instruction text for complex field parsing.
                if field_depth > 0 {
                    field_instr.push_str(text);
                }
            }
            Inline::AlternateContent(ac) => {
                // Pick fallback content (safest for PDF rendering)
                if let Some(ref fallback) = ac.fallback {
                    let mut sub = collect_fragments(
                        fallback,
                        default_family,
                        default_size,
                        default_color,
                        hyperlink_url,
                        measure_text,
                        resolved_styles,
                        paragraph_run_defaults,
                        footnote_counter,
                        endnote_counter,
                        field_ctx,
                        theme,
                    );
                    fragments.append(&mut sub);
                }
            }
            Inline::Symbol(sym) => {
                let font = FontProps {
                    family: Rc::from(sym.font.as_str()),
                    size: default_size,
                    bold: false,
                    italic: false,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                };
                let ch = char::from_u32(sym.char_code as u32).unwrap_or('\u{FFFD}');
                let text = ch.to_string();
                let (w, m) = measure_text(&text, &font);
                fragments.push(Fragment::Text {
                    text,
                    font,
                    color: RgbColor::BLACK,
                    shading: None,
                    border: None,
                    width: w,
                    trimmed_width: w,
                    metrics: m,
                    hyperlink_url: hyperlink_url.map(String::from),
                    baseline_offset: Pt::ZERO,
                    text_offset: Pt::ZERO,
                });
            }
            // Bookmark target — emit as zero-width named destination.
            Inline::BookmarkStart { name, .. } => {
                fragments.push(Fragment::Bookmark { name: name.clone() });
            }
            // Non-visual inlines — skip
            Inline::BookmarkEnd(_)
            | Inline::Separator
            | Inline::ContinuationSeparator
            | Inline::FootnoteRefMark
            | Inline::EndnoteRefMark => {}
            // §17.11.12: footnote reference — render as superscript number.
            Inline::FootnoteRef(_note_id) => {
                *footnote_counter += 1;
                let num_text = format!("{}", *footnote_counter);
                // §17.11.12: footnote reference uses superscript at 58% size.
                let ref_size = default_size * 0.58;
                let ref_font = FontProps {
                    family: std::rc::Rc::from(default_family),
                    size: ref_size,
                    bold: false,
                    italic: false,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                };
                let (w, m) = measure_text(&num_text, &ref_font);
                // Superscript baseline offset: raise by ~40% of the full-size ascent.
                let baseline_offset = -(default_size * 0.4);
                fragments.push(Fragment::Text {
                    text: num_text,
                    font: ref_font,
                    color: default_color,
                    shading: None,
                    border: None,
                    width: w,
                    trimmed_width: w,
                    metrics: m,
                    hyperlink_url: None,
                    baseline_offset,
                    text_offset: Pt::ZERO,
                });
            }
            // §17.11.2: endnote reference — render as superscript Roman numeral.
            Inline::EndnoteRef(_note_id) => {
                *endnote_counter += 1;
                let num_text = to_roman_lower(*endnote_counter);
                let ref_size = default_size * 0.58;
                let ref_font = FontProps {
                    family: std::rc::Rc::from(default_family),
                    size: ref_size,
                    bold: false,
                    italic: false,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                };
                let (w, m) = measure_text(&num_text, &ref_font);
                let baseline_offset = -(default_size * 0.4);
                fragments.push(Fragment::Text {
                    text: num_text,
                    font: ref_font,
                    color: default_color,
                    shading: None,
                    border: None,
                    width: w,
                    trimmed_width: w,
                    metrics: m,
                    hyperlink_url: None,
                    baseline_offset,
                    text_offset: Pt::ZERO,
                });
            }
            Inline::Pict(pict) => {
                // Render text content from VML text box shapes inline.
                // Does not handle absolute positioning — text appears inline
                // with the surrounding paragraph.
                for shape in &pict.shapes {
                    if let Some(ref text_box) = shape.text_box {
                        for block in &text_box.content {
                            if let Block::Paragraph(p) = block {
                                let para_run_defaults = p.mark_run_properties.as_ref();
                                let mut sub = collect_fragments(
                                    &p.content,
                                    default_family,
                                    default_size,
                                    default_color,
                                    hyperlink_url,
                                    measure_text,
                                    resolved_styles,
                                    para_run_defaults,
                                    footnote_counter,
                                    endnote_counter,
                                    field_ctx,
                                    theme,
                                );
                                fragments.append(&mut sub);
                            }
                        }
                    }
                }
            }
        }
    }

    fragments
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::{Dimension, HalfPoints};
    use crate::model::*;

    /// Dummy measurer: width = text.len() * 6.0, ascent = 10.0, descent = 2.0
    fn dummy_measure(text: &str, _font: &FontProps) -> (Pt, TextMetrics) {
        (
            Pt::new(text.len() as f32 * 6.0),
            TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(2.0),
            },
        )
    }

    fn text_run(text: &str) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    fn text_run_with_font(text: &str, font: &str, size: i64) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties {
                fonts: FontSet {
                    ascii: Some(font.into()),
                    ..Default::default()
                },
                font_size: Some(Dimension::<HalfPoints>::new(size)),
                ..Default::default()
            },
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    #[test]
    fn single_text_run() {
        let inlines = vec![text_run("hello")];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 1);
        assert_eq!(frags[0].width().raw(), 30.0); // 5 * 6
        assert_eq!(frags[0].height().raw(), 12.0);
    }

    #[test]
    fn text_run_uses_run_font() {
        let inlines = vec![text_run_with_font("hi", "Arial", 24)];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(10.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        if let Fragment::Text { font, .. } = &frags[0] {
            assert_eq!(&*font.family, "Arial");
            assert_eq!(font.size.raw(), 12.0); // 24 half-points = 12pt
        } else {
            panic!("expected Text fragment");
        }
    }

    #[test]
    fn tab_produces_tab_fragment() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Tab],
            rsids: RevisionIds::default(),
        }))];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 1);
        assert!(matches!(frags[0], Fragment::Tab { .. }));
    }

    #[test]
    fn line_break_produces_break_fragment() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::LineBreak(BreakKind::TextWrapping)],
            rsids: RevisionIds::default(),
        }))];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 1);
        assert!(frags[0].is_line_break());
    }

    #[test]
    fn hyperlink_recurses_into_content() {
        let inlines = vec![Inline::Hyperlink(Hyperlink {
            target: HyperlinkTarget::External(RelId::new("rId1")),
            content: vec![text_run("click me")],
        })];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 2, "split into 'click ' and 'me'");
        if let Fragment::Text {
            hyperlink_url,
            text,
            ..
        } = &frags[0]
        {
            assert_eq!(text, "click ");
            assert_eq!(hyperlink_url.as_deref(), Some("rId1"));
        } else {
            panic!("expected Text fragment");
        }
    }

    #[test]
    fn complex_field_skips_instructions_collects_result() {
        // FieldChar::Begin -> InstrText("PAGE") -> FieldChar::Separate -> TextRun("3") -> FieldChar::End
        let inlines = vec![
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::Begin,
                dirty: None,
                fld_lock: None,
            }),
            Inline::InstrText("PAGE".into()),
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::Separate,
                dirty: None,
                fld_lock: None,
            }),
            text_run("3"),
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::End,
                dirty: None,
                fld_lock: None,
            }),
        ];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        // Should only have the "3" result, not "PAGE"
        assert_eq!(frags.len(), 1);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(text, "3");
        }
    }

    #[test]
    fn bookmarks_and_separators_skipped() {
        let inlines = vec![
            Inline::BookmarkStart {
                id: BookmarkId::new(1),
                name: "bm1".into(),
            },
            text_run("visible"),
            Inline::BookmarkEnd(BookmarkId::new(1)),
            Inline::Separator,
            Inline::ContinuationSeparator,
            Inline::FootnoteRefMark,
            Inline::EndnoteRefMark,
            // LastRenderedPageBreak is now inside RunElement, not Inline
        ];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        // BookmarkStart produces a Bookmark fragment, text run produces a Text fragment.
        assert_eq!(
            frags.len(),
            2,
            "bookmark + text run should produce fragments"
        );
        assert!(matches!(frags[0], Fragment::Bookmark { .. }));
        assert!(matches!(frags[1], Fragment::Text { .. }));
    }

    #[test]
    fn alternate_content_uses_fallback() {
        let inlines = vec![Inline::AlternateContent(AlternateContent {
            choices: vec![McChoice {
                requires: McRequires::Wps,
                content: vec![text_run("choice")],
            }],
            fallback: Some(vec![text_run("fallback")]),
        })];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(text, "fallback");
        }
    }

    #[test]
    fn empty_text_run_produces_no_fragment() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Text(String::new())],
            rsids: RevisionIds::default(),
        }))];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );
        assert!(frags.is_empty());
    }

    #[test]
    fn font_props_default_fallback() {
        let rp = RunProperties::default();
        let fp = font_props_from_run(&rp, "Helvetica", Pt::new(12.0));
        assert_eq!(&*fp.family, "Helvetica");
        assert_eq!(fp.size.raw(), 12.0);
        assert!(!fp.bold);
        assert!(!fp.italic);
    }

    #[test]
    fn symbol_produces_text_fragment() {
        let inlines = vec![Inline::Symbol(Symbol {
            font: "Wingdings".into(),
            char_code: 0x46, // 'F'
        })];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { font, text, .. } = &frags[0] {
            assert_eq!(&*font.family, "Wingdings");
            assert_eq!(text, "F");
        }
    }

    #[test]
    fn simple_field_collects_content() {
        let inlines = vec![Inline::Field(Field {
            instruction: crate::field::FieldInstruction::Page {
                switches: Default::default(),
            },
            content: vec![text_run("5")],
        })];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(text, "5");
        }
    }

    // ── split_into_words ─────────────────────────────────────────────────

    #[test]
    fn split_single_word() {
        assert_eq!(split_into_words("hello"), vec!["hello"]);
    }

    #[test]
    fn split_two_words() {
        assert_eq!(split_into_words("hello world"), vec!["hello ", "world"]);
    }

    #[test]
    fn split_trailing_space() {
        assert_eq!(split_into_words("hello "), vec!["hello "]);
    }

    #[test]
    fn split_multiple_words() {
        assert_eq!(
            split_into_words("the quick brown fox"),
            vec!["the ", "quick ", "brown ", "fox"]
        );
    }

    #[test]
    fn split_empty() {
        let result: Vec<&str> = split_into_words("");
        assert!(result.is_empty());
    }

    #[test]
    fn multi_word_text_run_splits_into_fragments() {
        let inlines = vec![text_run("hello world foo")];
        let frags = collect_fragments(
            &inlines,
            "Default",
            Pt::new(12.0),
            RgbColor::BLACK,
            None,
            &dummy_measure,
            None,
            None,
            &mut 0,
            &mut 0,
            FieldContext::default(),
            None,
        );

        assert_eq!(frags.len(), 3);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(text, "hello ");
        }
        if let Fragment::Text { text, .. } = &frags[1] {
            assert_eq!(text, "world ");
        }
        if let Fragment::Text { text, .. } = &frags[2] {
            assert_eq!(text, "foo");
        }
    }
}
