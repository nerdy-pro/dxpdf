//! §17.3.2.35 / §17.3.1.13: the unit that inter-character spacing is applied
//! *between*.
//!
//! Two unrelated features insert horizontal space inside a run:
//!
//! ```xml
//! <w:rPr><w:spacing w:val="40"/></w:rPr>     <!-- §17.3.2.35: +2pt everywhere -->
//! <w:pPr><w:jc w:val="distribute"/></w:pPr>  <!-- §17.3.1.13: share the line's slack -->
//! ```
//!
//! They differ in where the amount comes from — `w:spacing` is authored, while
//! `distribute` is derived per line from the width left over — but they land in
//! the same place, add up (`font.char_spacing + distribution_extra`), and reach
//! the painter as the single `char_spacing` field of `DrawCommand::Text`.
//!
//! So both need one answer to one question: **where may space go?**
//!
//! # The unit is a grapheme cluster
//!
//! This module owns that answer, and it is the UAX #29 extended grapheme
//! cluster. [`unit_count`] says how many helpings of spacing a run earns;
//! [`units`] says which substrings the painter draws between them. Layout and
//! paint call the same two functions, which is the whole reason the module
//! exists — a run that measures to one width and paints at another leaves its
//! underline, its run border and its hyperlink rect behind.
//!
//! Counting Unicode **scalars** instead, which is what this code did until
//! issue #82, breaks any cluster spelled with more than one:
//!
//! | Text | Scalars | Clusters | Scalar counting drew |
//! |---|---|---|---|
//! | `e` + `U+0301` | 2 | 1 | the acute accent a spacing step right of its `e` |
//! | `1` `U+FE0F` `U+20E3` | 3 | 1 | a digit, a gap, and a bare keycap ring |
//! | `U+1F1E9` `U+1F1EA` | 2 | 1 | two lettered squares instead of a German flag |
//! | `U+0915` `U+094D` `U+0937` | 3 | 1 | a conjunct pulled apart at its virama |
//!
//! # Who asks
//!
//! | Caller | Uses it for |
//! |---|---|
//! | `layout::measurer::measure` | §17.3.2.35 — adds `char_spacing × unit_count` to the measured width |
//! | `layout::paragraph::line_emit::distribution_unit_count` | §17.3.1.13 — how many units a fragment contributes, hence how many gaps a line has to fill |
//! | `render::painter`, the `char_spacing != 0` arm | draws one unit at a time, advancing by the unit's own advance plus `char_spacing` |
//! | `layout::fragment::split` | per-unit fragments for an over-wide word |
//!
//! The last is not spacing at all, and is governed here for the same reason:
//! splitting an over-wide word per *scalar* gave the line-fitter permission to
//! break between a letter and its own accent and carry the mark to the next
//! line.
//!
//! # For a shaped run, the unit is the shaped cluster
//!
//! This module's grapheme cluster is the unit **for text the painter draws
//! from its string**: `draw_str` and `TextBlob::from_str` map codepoints to
//! glyphs through the cmap alone, with no GSUB, so no finer or coarser
//! boundary than UAX #29's is honourable there — and that is still the path
//! every Latin, Cyrillic, Greek, CJK and Thai run takes (and Hebrew *script*
//! could take, but a Hebrew run at its ordinary right-to-left level is shaped
//! for glyph order — `layout::fragment::shape`'s reason 2 — so in practice
//! Hebrew text uses the shaped unit below), which is why nothing in this file
//! changed for issue #153.
//!
//! A run [`crate::render::shape::needs_shaping`] picks out — cursive-joining
//! since issue #131, reordering (Indic) since issue #153 — or that sits at a
//! right-to-left level is shaped through HarfBuzz, and its unit is the
//! **shaped cluster**: the glyphs sharing one HarfBuzz cluster value. A conjunct's glyphs share one; a ligature merges
//! what were separate grapheme clusters; a prebase matra's glyph belongs to a
//! cluster whose text it precedes. Grapheme boundaries would fall inside all
//! three, which is exactly the "spacing inside a unit" defect the table above
//! shows for scalars. The same three functions of the question exist for that
//! unit, one per asker:
//!
//! | Caller | Shaped counterpart |
//! |---|---|
//! | measurement | [`ShapedRun::unit_count`](crate::render::shape::ShapedRun::unit_count), via `TextMeasurer::shaped_measurement` — `layout::fragment::shape` swaps both the advance and the spacing term's count |
//! | distribution | the count stored on the fragment (`fragment::Shaping::unit_count`), because `line_emit` has no measurer |
//! | paint | the painter's own shaping pass, grouped by cluster value (`painter::cluster_spacing_shifts`) |
//!
//! All three are one number because all three come from the same shaper over
//! the same typeface — the same argument that keeps `units`/`unit_count`
//! below a single definition for the cmap path.
//!
//! One caller keeps the grapheme unit even for a shaped run: the emergency
//! splitter (`layout::fragment::split`). A shaped cluster is never *finer*
//! than a grapheme cluster — the shape module folds cluster values to
//! grapheme starts — so a grapheme seam falls either *between* shaped
//! clusters or across a ligature that re-shaping the halves degrades legibly:
//! a lam-alef loses its ligature, and a conjunct in the scripts whose viramas
//! UAX #29's InCB derivation leaves out (Tamil, Kannada, Gurmukhi, Sinhala —
//! GB9c keeps the *other* Indic conjuncts one grapheme) degrades to halant
//! forms. Both are that module's documented last resort, and each piece is
//! re-measured shaped so its box still matches what is painted.

use unicode_segmentation::UnicodeSegmentation;

/// True when one byte of `text` is exactly one spacing unit, so the grapheme
/// segmenter can be skipped.
///
/// `\r` is excluded even though it is ASCII: UAX #29 GB3 keeps CRLF together as
/// a *single* cluster. Text reaching layout has its C0 controls stripped
/// (`fragment::text::emit_text_fragments`), but this module is also called from
/// the painter and must not inherit that assumption.
fn is_one_byte_per_unit(text: &str) -> bool {
    text.is_ascii() && !text.as_bytes().contains(&b'\r')
}

/// The spacing units of `text`, in order — the substrings that spacing is
/// inserted *between*, and that the painter draws one at a time.
pub fn units(text: &str) -> impl Iterator<Item = &str> + '_ {
    // Two concrete iterators behind one `impl Iterator`: the ASCII path walks
    // byte slices, which is what the overwhelming majority of runs take.
    let (ascii, graphemes) = if is_one_byte_per_unit(text) {
        (Some((0..text.len()).map(|i| &text[i..i + 1])), None)
    } else {
        (None, Some(text.graphemes(true)))
    };
    ascii
        .into_iter()
        .flatten()
        .chain(graphemes.into_iter().flatten())
}

/// How many spacing units `text` holds — i.e. how many times a per-unit amount
/// is added when the run is measured.
pub fn unit_count(text: &str) -> usize {
    if is_one_byte_per_unit(text) {
        text.len()
    } else {
        text.graphemes(true).count()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The defect this module exists to fix: a combining mark must never be a
    /// unit of its own, or spacing lands between the letter and its accent.
    #[test]
    fn combining_mark_joins_its_base() {
        let accented = "e\u{301}";
        assert_eq!(accented.chars().count(), 2, "two scalars");
        assert_eq!(unit_count(accented), 1, "but one spacing unit");
        assert_eq!(units(accented).collect::<Vec<_>>(), vec!["e\u{301}"]);
    }

    /// The same rule across the multi-scalar sequences a DOCX actually carries:
    /// keycap, variation selector, ZWJ family, regional-indicator flag.
    #[test]
    fn multi_scalar_sequences_are_single_units() {
        for (label, text) in [
            ("keycap", "1\u{FE0F}\u{20E3}"),
            ("variation selector", "\u{2764}\u{FE0F}"),
            ("ZWJ family", "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}"),
            ("regional indicator", "\u{1F1E9}\u{1F1EA}"),
            ("devanagari akshara", "\u{0915}\u{094D}\u{0937}"),
        ] {
            assert_eq!(unit_count(text), 1, "{label} must be one spacing unit");
        }
    }

    /// UAX #29 GB3 — the reason the ASCII fast path rejects `\r`. A CRLF split
    /// into two units would put spacing inside a single cluster, the very thing
    /// this module forbids.
    #[test]
    fn crlf_is_one_unit_despite_being_ascii() {
        assert!(!is_one_byte_per_unit("a\r\nb"));
        assert_eq!(unit_count("\r\n"), 1);
        assert_eq!(units("a\r\nb").collect::<Vec<_>>(), vec!["a", "\r\n", "b"]);
    }

    /// The fast path is an optimisation, not a second definition: it must agree
    /// with the segmenter everywhere it is taken.
    #[test]
    fn ascii_fast_path_agrees_with_the_segmenter() {
        for text in ["", "a", "hello world", "tab\there", "a-b-c", "  "] {
            assert!(is_one_byte_per_unit(text), "{text:?} should take fast path");
            assert_eq!(
                unit_count(text),
                text.graphemes(true).count(),
                "unit_count disagrees with the segmenter on {text:?}"
            );
            assert_eq!(
                units(text).collect::<Vec<_>>(),
                text.graphemes(true).collect::<Vec<_>>(),
                "units disagrees with the segmenter on {text:?}"
            );
        }
    }

    /// Non-ASCII text without combining marks still counts one unit per
    /// character — spacing is not silently suppressed for Cyrillic or CJK.
    #[test]
    fn simple_non_ascii_counts_one_unit_per_character() {
        assert_eq!(unit_count("привет"), 6);
        assert_eq!(unit_count("日本語"), 3);
    }

    /// `units` and `unit_count` are one definition seen two ways; a caller that
    /// measures with one and paints with the other must not drift.
    #[test]
    fn units_and_unit_count_agree() {
        for text in [
            "",
            "abc",
            "e\u{301}x",
            "日本語",
            "a\r\nb",
            "\u{1F1E9}\u{1F1EA}!",
        ] {
            assert_eq!(
                units(text).count(),
                unit_count(text),
                "count mismatch on {text:?}"
            );
            assert_eq!(
                units(text).collect::<String>(),
                text,
                "units must reconstruct the input exactly for {text:?}"
            );
        }
    }
}
