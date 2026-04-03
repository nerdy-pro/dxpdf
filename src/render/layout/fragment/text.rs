use std::rc::Rc;

use crate::render::dimension::Pt;
use crate::render::resolve::color::RgbColor;

use super::{Fragment, FragmentBorder, FontProps, TextMetrics};

/// §17.18.40 ST_HighlightColor: map highlight enum to RGB.
/// These are the fixed palette colors defined in the OOXML spec.
pub(super) fn resolve_highlight_color(hl: crate::model::HighlightColor) -> RgbColor {
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
pub(super) fn emit_text_fragments<F>(
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
            text: Rc::from(word),
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
