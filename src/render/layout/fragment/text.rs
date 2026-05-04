use std::rc::Rc;

use crate::render::dimension::Pt;
use crate::render::emoji::cluster::{self, EmojiCluster, InlineCluster};
use crate::render::emoji::resolve::{EmojiFamily, EmojiTypeface};
use crate::render::layout::measurer::TextMeasurer;
use crate::render::resolve::color::RgbColor;

use super::{FontProps, Fragment, FragmentBorder, TextMetrics};

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

/// Resolved styling for a single text fragment.
pub(super) struct TextRunStyle {
    pub color: RgbColor,
    pub shading: Option<RgbColor>,
    pub border: Option<FragmentBorder>,
    pub baseline_offset: Pt,
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
///
/// When `measurer` is `Some`, the text is first split into grapheme clusters
/// (UAX #29) and each cluster is classified per UTS #51. Emoji clusters that
/// resolve to a host color emoji typeface become [`Fragment::Emoji`]; clusters
/// without a resolved typeface fall through to the text path with a one-time
/// warning per cluster. When `measurer` is `None` (used by unit tests that
/// don't construct a font registry), the input is passed straight to the
/// existing word-split + measure path — preserving prior behaviour.
pub(super) fn emit_text_fragments<F>(
    text: &str,
    font: &FontProps,
    style: &TextRunStyle,
    hyperlink_url: Option<&str>,
    measure_text: &F,
    measurer: Option<&TextMeasurer<'_>>,
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

    let Some(measurer) = measurer else {
        emit_text_words(
            &cleaned,
            font,
            style,
            hyperlink_url,
            measure_text,
            fragments,
        );
        return;
    };

    // Classify into clusters and route emoji clusters through the raster
    // pipeline; text spans go through the existing word-split path.
    for cluster in cluster::classify(&cleaned) {
        match cluster {
            InlineCluster::Text(span) => {
                emit_text_words(span, font, style, hyperlink_url, measure_text, fragments);
            }
            InlineCluster::Emoji(emoji) => {
                emit_emoji_or_fallback(
                    &emoji,
                    font,
                    style,
                    hyperlink_url,
                    measure_text,
                    measurer,
                    fragments,
                );
            }
        }
    }
}

/// Word-split + measure path. Identical to the prior body of
/// [`emit_text_fragments`]; factored out so emoji-cluster fallback can reuse it.
pub(super) fn emit_text_words<F>(
    text: &str,
    font: &FontProps,
    style: &TextRunStyle,
    hyperlink_url: Option<&str>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    if text.is_empty() {
        return;
    }
    for word in split_into_words(text) {
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
            color: style.color,
            shading: style.shading,
            border: style.border,
            width: w,
            trimmed_width: tw,
            metrics: m,
            hyperlink_url: hyperlink_url.map(String::from),
            baseline_offset: style.baseline_offset,
            text_offset: Pt::ZERO,
        });
    }
}

/// Resolve a host color emoji typeface for an emoji cluster and emit a
/// [`Fragment::Emoji`]. On `Unavailable`, log a one-time warning and route
/// the cluster through the text path so its codepoints still appear in the
/// PDF text stream (per the no-bundle / no-silent-degradation policy in
/// `docs/emoji-rendering.md`).
pub(super) fn emit_emoji_or_fallback<F>(
    cluster: &EmojiCluster<'_>,
    font: &FontProps,
    style: &TextRunStyle,
    hyperlink_url: Option<&str>,
    measure_text: &F,
    measurer: &TextMeasurer<'_>,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    // §17.3.2.26: the run's font name acts as a hint. If it names a known
    // color emoji family, prefer it; otherwise we fall through to the
    // host-default chain (e.g. Calibri-tagged runs containing 📞 still
    // resolve to Apple Color Emoji on macOS).
    let requested = EmojiFamily::from_name_ci(&font.family);
    match measurer.resolve_emoji(requested) {
        EmojiTypeface::Resolved {
            entry: typeface, ..
        } => {
            let (advance, metrics) =
                measurer.measure_with_typeface(cluster.text, &typeface, font.size);
            // Line-height contribution uses the run's text-font metrics so
            // an inline emoji doesn't bloat the line. Color emoji typefaces
            // have ≈1.25× ascent which would otherwise stretch every line
            // that contains an emoji vs surrounding text-only lines.
            let (_, line_metrics) = measure_text("X", font);
            fragments.push(Fragment::Emoji {
                text: cluster.text.to_string(),
                typeface,
                size: font.size,
                presentation: cluster.presentation,
                structure: cluster.structure,
                advance,
                metrics,
                line_metrics,
                baseline_offset: style.baseline_offset,
            });
        }
        EmojiTypeface::Unavailable { attempted } => {
            measurer.warn_emoji_unavailable_once(cluster.text, &attempted);
            emit_text_words(
                cluster.text,
                font,
                style,
                hyperlink_url,
                measure_text,
                fragments,
            );
        }
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

    // ── L1, L2, L4: emoji cluster integration ────────────────────────────

    use crate::render::emoji::cluster::EmojiPresentation;
    use crate::render::fonts::FontRegistry;
    use crate::render::layout::measurer::TextMeasurer;
    use skia_safe::FontMgr;
    use std::rc::Rc;

    fn font(family: &str, size: f32) -> FontProps {
        FontProps {
            family: Rc::from(family),
            size: Pt::new(size),
            bold: false,
            italic: false,
            underline: false,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        }
    }

    fn style() -> TextRunStyle {
        TextRunStyle {
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
        }
    }

    /// L1 — Run "hi 📞" produces `[Text("hi "), Emoji(...)]` when an emoji
    /// typeface is resolvable on the host. Skipped on hosts without one.
    #[test]
    fn l1_emoji_run_splits_into_text_and_emoji_fragments() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        // Bail if the host has no color emoji — Phase 3 is platform-aware.
        use crate::render::emoji::resolve::EmojiTypeface;
        if matches!(
            measurer.resolve_emoji(None),
            EmojiTypeface::Unavailable { .. }
        ) {
            eprintln!("skipping L1: no color emoji typeface on this host");
            return;
        }
        let mut fragments = Vec::new();
        let measure = |text: &str, fp: &FontProps| measurer.measure(text, fp);
        emit_text_fragments(
            "hi \u{1F4DE}",
            &font("Calibri", 12.0),
            &style(),
            None,
            &measure,
            Some(&measurer),
            &mut fragments,
        );
        assert_eq!(
            fragments.len(),
            2,
            "expected 2 fragments (Text + Emoji), got {fragments:#?}"
        );
        match &fragments[0] {
            Fragment::Text { text, .. } => assert_eq!(&**text, "hi "),
            other => panic!("first fragment must be Text, got {other:?}"),
        }
        match &fragments[1] {
            Fragment::Emoji {
                text,
                presentation,
                advance,
                ..
            } => {
                assert_eq!(text, "\u{1F4DE}");
                assert_eq!(*presentation, EmojiPresentation::Emoji);
                // L2 — advance must be > 0 when the typeface is resolved.
                assert!(
                    advance.raw() > 0.0,
                    "advance must be positive, got {advance}"
                );
            }
            other => panic!("second fragment must be Emoji, got {other:?}"),
        }
    }

    /// L4 — When `measurer` is `None` (no emoji pipeline available), emoji
    /// codepoints flow through the existing text path unchanged. This matches
    /// the no-bundle / no-silent-degradation policy: the codepoint is still
    /// preserved in the PDF's text stream.
    #[test]
    fn l4_no_measurer_routes_emoji_through_text_path() {
        let mut fragments = Vec::new();
        let measure = |text: &str, _fp: &FontProps| {
            (
                Pt::new(text.len() as f32 * 6.0),
                TextMetrics {
                    ascent: Pt::new(10.0),
                    descent: Pt::new(2.0),
                    leading: Pt::ZERO,
                },
            )
        };
        emit_text_fragments(
            "hi \u{1F4DE}",
            &font("Calibri", 12.0),
            &style(),
            None,
            &measure,
            None,
            &mut fragments,
        );
        // No measurer → the whole input is fed to the text path. There must
        // be zero Emoji fragments and the original codepoint must appear in
        // exactly one Text fragment.
        for f in &fragments {
            assert!(
                !matches!(f, Fragment::Emoji { .. }),
                "no emoji fragments must be produced when measurer is None"
            );
        }
        let joined: String = fragments
            .iter()
            .filter_map(|f| match f {
                Fragment::Text { text, .. } => Some(&**text),
                _ => None,
            })
            .collect();
        assert!(
            joined.contains('\u{1F4DE}'),
            "emoji codepoint must survive through the text path"
        );
    }

    // ── split_into_words (existing tests) ────────────────────────────────

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
