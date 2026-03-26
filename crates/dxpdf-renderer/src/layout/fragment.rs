//! Fragment conversion — transform Inline content into measured Fragments
//! for the line-fitting algorithm.

use std::rc::Rc;

use dxpdf_docx_model::model::{FieldCharType, Inline, RunProperties};

use crate::dimension::Pt;
use crate::geometry::PtSize;
use crate::resolve::color::RgbColor;
use crate::resolve::fonts::effective_font;

/// Font properties needed for rendering a text fragment.
#[derive(Clone, Debug)]
pub struct FontProps {
    pub family: Rc<str>,
    pub size: Pt,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub char_spacing: Pt,
}

/// A measured fragment — the atomic unit for line fitting.
#[derive(Clone, Debug)]
pub enum Fragment {
    Text {
        text: String,
        font: FontProps,
        color: RgbColor,
        width: Pt,
        height: Pt,
        ascent: Pt,
        hyperlink_url: Option<String>,
        baseline_offset: Pt,
    },
    Image {
        size: PtSize,
        rel_id: String,
        image_data: Option<std::rc::Rc<[u8]>>,
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
            Fragment::Text { width, .. } => *width,
            Fragment::Image { size, .. } => size.width,
            Fragment::Tab { .. } => MIN_TAB_WIDTH,
            Fragment::LineBreak { .. } => Pt::ZERO,
        }
    }

    pub fn height(&self) -> Pt {
        match self {
            Fragment::Text { height, .. } => *height,
            Fragment::Image { size, .. } => size.height,
            Fragment::Tab { line_height } | Fragment::LineBreak { line_height } => *line_height,
        }
    }

    pub fn is_line_break(&self) -> bool {
        matches!(self, Fragment::LineBreak { .. })
    }
}

/// Minimum tab fragment width for line fitting.
pub const MIN_TAB_WIDTH: Pt = Pt::new(12.0);

/// Extract font properties from RunProperties with a default font family fallback.
pub fn font_props_from_run(
    rp: &RunProperties,
    default_family: &str,
    default_size: Pt,
) -> FontProps {
    let family = effective_font(&rp.fonts)
        .unwrap_or(default_family);

    let size = rp
        .font_size
        .map(Pt::from)
        .unwrap_or(default_size);

    let char_spacing = rp
        .spacing
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);

    FontProps {
        family: Rc::from(family),
        size,
        bold: rp.bold.unwrap_or(false),
        italic: rp.italic.unwrap_or(false),
        underline: rp.underline.is_some(),
        char_spacing,
    }
}

/// Split text into word-level chunks for line breaking.
/// Whitespace is kept attached to the preceding word: "hello world" → ["hello ", "world"].
/// This allows the line fitter to break between fragments at word boundaries.
fn split_into_words(text: &str) -> Vec<&str> {
    let mut words = Vec::new();
    let mut start = 0;

    for (i, ch) in text.char_indices() {
        if ch == ' ' || ch == '\t' {
            // Include the whitespace with the word that precedes it
            let end = i + ch.len_utf8();
            if end > start {
                words.push(&text[start..end]);
                start = end;
            }
        }
    }

    // Remaining text (last word without trailing space)
    if start < text.len() {
        words.push(&text[start..]);
    }

    words
}

/// Walk inline content and collect fragments.
/// `measure_text` is a callback that measures text width/height/ascent for a given font.
///
/// Returns fragments suitable for the line-fitting algorithm.
pub fn collect_fragments<F>(
    inlines: &[Inline],
    default_family: &str,
    default_size: Pt,
    hyperlink_url: Option<&str>,
    measure_text: &F,
) -> Vec<Fragment>
where
    F: Fn(&str, &FontProps) -> (Pt, Pt, Pt), // (width, height, ascent)
{
    let mut fragments = Vec::new();
    let mut field_depth: i32 = 0; // tracks nested complex field state

    for inline in inlines {
        match inline {
            Inline::TextRun(tr) => {
                // Skip field instruction text (between begin and separate)
                if field_depth > 0 {
                    continue;
                }
                let font = font_props_from_run(&tr.properties, default_family, default_size);
                let color = tr
                    .properties
                    .color
                    .map(|c| crate::resolve::color::resolve_color(c, crate::resolve::color::ColorContext::Text))
                    .unwrap_or(RgbColor::BLACK);

                if !tr.text.is_empty() {
                    // Split text into word-level fragments so the line fitter
                    // can break between words. Whitespace is kept as a trailing
                    // part of the preceding word (e.g., "hello " + "world").
                    for word in split_into_words(&tr.text) {
                        let (w, h, a) = measure_text(word, &font);
                        fragments.push(Fragment::Text {
                            text: word.to_string(),
                            font: font.clone(),
                            color,
                            width: w,
                            height: h,
                            ascent: a,
                            hyperlink_url: hyperlink_url.map(String::from),
                            baseline_offset: Pt::ZERO,
                        });
                    }
                }
            }
            Inline::Tab => {
                fragments.push(Fragment::Tab {
                    line_height: default_size,
                });
            }
            Inline::LineBreak(_) => {
                fragments.push(Fragment::LineBreak {
                    line_height: default_size,
                });
            }
            Inline::PageBreak | Inline::ColumnBreak => {
                // Treated as line breaks in fragment collection
                fragments.push(Fragment::LineBreak {
                    line_height: default_size,
                });
            }
            Inline::Image(img) => {
                if let Some(rel_id) = crate::resolve::images::extract_image_rel_id(img) {
                    let w = Pt::from(img.extent.width);
                    let h = Pt::from(img.extent.height);
                    fragments.push(Fragment::Image {
                        size: PtSize::new(w, h),
                        rel_id: rel_id.as_str().to_string(),
                        image_data: None, // populated by caller with media bytes
                    });
                }
            }
            Inline::Hyperlink(link) => {
                let url = match &link.target {
                    dxpdf_docx_model::model::HyperlinkTarget::External(rel_id) => {
                        Some(rel_id.as_str())
                    }
                    dxpdf_docx_model::model::HyperlinkTarget::Internal { anchor } => {
                        Some(anchor.as_str())
                    }
                };
                let mut sub = collect_fragments(
                    &link.content,
                    default_family,
                    default_size,
                    url,
                    measure_text,
                );
                fragments.append(&mut sub);
            }
            Inline::Field(field) => {
                // Simple field — collect content (the cached result)
                let mut sub = collect_fragments(
                    &field.content,
                    default_family,
                    default_size,
                    hyperlink_url,
                    measure_text,
                );
                fragments.append(&mut sub);
            }
            Inline::FieldChar(fc) => {
                // Complex field state machine:
                // Begin -> InstrText... -> Separate -> result runs -> End
                match fc.field_char_type {
                    FieldCharType::Begin => field_depth += 1,
                    FieldCharType::Separate => field_depth -= 1, // now collect result runs
                    FieldCharType::End => {} // no-op, result already collected
                }
            }
            Inline::InstrText(_) => {
                // Skipped — field instruction text, not rendered
            }
            Inline::AlternateContent(ac) => {
                // Pick fallback content (safest for PDF rendering)
                if let Some(ref fallback) = ac.fallback {
                    let mut sub = collect_fragments(
                        fallback,
                        default_family,
                        default_size,
                        hyperlink_url,
                        measure_text,
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
                };
                let ch = char::from_u32(sym.char_code as u32).unwrap_or('\u{FFFD}');
                let text = ch.to_string();
                let (w, h, a) = measure_text(&text, &font);
                fragments.push(Fragment::Text {
                    text,
                    font,
                    color: RgbColor::BLACK,
                    width: w,
                    height: h,
                    ascent: a,
                    hyperlink_url: hyperlink_url.map(String::from),
                    baseline_offset: Pt::ZERO,
                });
            }
            // Non-visual inlines — skip
            Inline::BookmarkStart { .. }
            | Inline::BookmarkEnd(_)
            | Inline::Separator
            | Inline::ContinuationSeparator
            | Inline::FootnoteRefMark
            | Inline::EndnoteRefMark
            | Inline::LastRenderedPageBreak => {}
            // Not yet handled — skip silently
            Inline::FootnoteRef(_)
            | Inline::EndnoteRef(_)
            | Inline::Pict(_) => {}
        }
    }

    fragments
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::dimension::{Dimension, HalfPoints};
    use dxpdf_docx_model::model::*;

    /// Dummy measurer: width = text.len() * 6.0, height = 12.0, ascent = 10.0
    fn dummy_measure(text: &str, _font: &FontProps) -> (Pt, Pt, Pt) {
        (
            Pt::new(text.len() as f32 * 6.0),
            Pt::new(12.0),
            Pt::new(10.0),
        )
    }

    fn text_run(text: &str) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            text: text.into(),
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
            text: text.into(),
            rsids: RevisionIds::default(),
        }))
    }

    #[test]
    fn single_text_run() {
        let inlines = vec![text_run("hello")];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

        assert_eq!(frags.len(), 1);
        assert_eq!(frags[0].width().raw(), 30.0); // 5 * 6
        assert_eq!(frags[0].height().raw(), 12.0);
    }

    #[test]
    fn text_run_uses_run_font() {
        let inlines = vec![text_run_with_font("hi", "Arial", 24)];
        let frags = collect_fragments(&inlines, "Default", Pt::new(10.0), None, &dummy_measure);

        if let Fragment::Text { font, .. } = &frags[0] {
            assert_eq!(&*font.family, "Arial");
            assert_eq!(font.size.raw(), 12.0); // 24 half-points = 12pt
        } else {
            panic!("expected Text fragment");
        }
    }

    #[test]
    fn tab_produces_tab_fragment() {
        let inlines = vec![Inline::Tab];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

        assert_eq!(frags.len(), 1);
        assert!(matches!(frags[0], Fragment::Tab { .. }));
    }

    #[test]
    fn line_break_produces_break_fragment() {
        let inlines = vec![Inline::LineBreak(BreakKind::TextWrapping)];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

        assert_eq!(frags.len(), 1);
        assert!(frags[0].is_line_break());
    }

    #[test]
    fn hyperlink_recurses_into_content() {
        let inlines = vec![Inline::Hyperlink(Hyperlink {
            target: HyperlinkTarget::External(RelId::new("rId1")),
            content: vec![text_run("click me")],
        })];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

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
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

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
            Inline::LastRenderedPageBreak,
        ];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

        assert_eq!(frags.len(), 1, "only the text run should produce a fragment");
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
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

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
            text: String::new(),
            rsids: RevisionIds::default(),
        }))];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);
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
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { font, text, .. } = &frags[0] {
            assert_eq!(&*font.family, "Wingdings");
            assert_eq!(text, "F");
        }
    }

    #[test]
    fn simple_field_collects_content() {
        let inlines = vec![Inline::Field(Field {
            instruction: dxpdf_field::FieldInstruction::Page {
                switches: Default::default(),
            },
            content: vec![text_run("5")],
        })];
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

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
        let frags = collect_fragments(&inlines, "Default", Pt::new(12.0), None, &dummy_measure);

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
