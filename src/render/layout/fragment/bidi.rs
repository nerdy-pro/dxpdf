//! UAX #9 level resolution over a paragraph's finished fragments.
//!
//! The bridge between [`crate::i18n::bidi`], which knows the algorithm, and
//! [`crate::render::layout::paragraph`], which reorders a line. It runs once
//! per paragraph and leaves behind the one thing the reorder needs: every
//! fragment carrying exactly one embedding level.
//!
//! # Why a pass over fragments, and not a split axis in `segment`
//!
//! #130 added its UAX #14 boundaries inside `segment::JoinedTextSegment`,
//! whose scope is a contiguous stretch of text runs. That scope is too small
//! here. UAX #9 resolves a neutral or weak character from the strong ones
//! around it, and "around" reaches across a tab, an inline image, a field
//! result — everything a paragraph can hold. So the analysis string this
//! module builds spans the *whole* paragraph, with a placeholder standing in
//! for each non-text fragment:
//!
//! | Fragment | Contributes | Because |
//! |---|---|---|
//! | `Text` | its own text | |
//! | `Tab`, `PTab` | `U+0009` | class S — rule L1 resets it to the paragraph level |
//! | `Image`, `Emoji` | `U+FFFC` | class ON — an object is a neutral, and its neighbours must resolve against one |
//! | `LineBreak`, `PageBreak`, `ColumnBreak` | `U+000A` | class B — rule P1 makes a forced break a new bidi paragraph, as CSS does |
//! | `Bookmark` | nothing | zero width, and not text |
//!
//! §17.3.2.30 `w:rtl` enters the same string as an isolate: a run whose
//! toggle is on is wrapped in `RLI`…`PDI`, one explicitly off in `LRI`…`PDI`.
//! Consecutive fragments sharing a toggle share one isolate, so a word split
//! across `<w:r>` boundaries is not fenced off from its own other half.
//!
//! # What it leaves behind
//!
//! A fragment a level boundary falls inside is **split**, because the reorder
//! is a permutation of fragments and a fragment that is half one direction
//! cannot be placed. Splitting keeps #130's invariant for run boundaries: the
//! last piece inherits the parent's `break_after`, every earlier piece is
//! [`BreakAfter::Prohibited`], since a level change is not a line-break
//! opportunity.
//!
//! Rule L4 (mirroring) is applied here rather than at paint, so that a
//! mirrored glyph's *own* advance is the one measured — `(` and `)` are not
//! obliged to be the same width, and a run whose measured width disagrees with
//! what is painted leaves its underline and its hyperlink rect behind.

use std::rc::Rc;

use crate::i18n::bidi::{self, BaseDirection, BidiLevel, POP_ISOLATE};
use crate::render::dimension::Pt;
use crate::render::fonts::Toggle;

use super::{BreakAfter, FontProps, Fragment, TextMetrics};

/// Stand-in for a fragment that is not text, in the analysis string.
///
/// Each is chosen for its UAX #9 bidi class rather than for what it looks
/// like; the table in this module's doc says which and why.
const OBJECT: char = '\u{FFFC}';

/// Resolve every fragment's embedding level, splitting and mirroring as UAX #9
/// requires.
///
/// Call once per paragraph, on the **finished** vector — after a list label
/// (§17.9.22) or a note body's number (§17.11.12) has been prefixed, since
/// those are as much part of the paragraph's text as anything the document
/// wrote, and a label left out of the analysis would be placed at the wrong
/// end of a `w:bidi` line.
///
/// Returns without touching anything when the paragraph is left-to-right and
/// holds no right-to-left character — which is every paragraph of most
/// documents, and the reason this costs them a scan and no allocation.
pub fn assign_bidi_levels<F>(fragments: &mut Vec<Fragment>, base: BaseDirection, measure_text: &F)
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    if !needs_pass(fragments, base) {
        return;
    }

    let (analysis, ranges) = build_analysis(fragments);
    let levels = bidi::resolve_levels(&analysis, base);

    let mut out = Vec::with_capacity(fragments.len());
    for (idx, fragment) in fragments.drain(..).enumerate() {
        match ranges[idx] {
            Some(ref range) if !range.is_empty() => {
                split_at_level_boundaries(fragment, &levels[range.clone()], measure_text, &mut out);
            }
            // Not text, or text that contributed nothing: nothing to resolve.
            // `Fragment::bidi_level` gives these the paragraph's base level at
            // reorder time.
            _ => out.push(fragment),
        }
    }
    *fragments = out;
}

/// Whether this paragraph has any bidirectional text to resolve at all.
///
/// Three ways in, and the cheap one first: a `w:bidi` paragraph always needs
/// the pass, because even its neutrals resolve right-to-left. Otherwise it
/// takes a §17.3.2.30 `w:rtl` run, or an actual right-to-left character.
fn needs_pass(fragments: &[Fragment], base: BaseDirection) -> bool {
    base == BaseDirection::Rtl
        || fragments.iter().any(|f| match f {
            Fragment::Text { text, font, .. } => {
                font.rtl != Toggle::Absent || bidi::needs_analysis(text)
            }
            _ => false,
        })
}

/// Build the paragraph's analysis string, and each fragment's byte range in it.
///
/// The range is `None` for a fragment whose contribution is a placeholder or
/// nothing at all — only text has a level of its own to read back.
#[allow(clippy::type_complexity)]
fn build_analysis(fragments: &[Fragment]) -> (String, Vec<Option<std::ops::Range<usize>>>) {
    let mut text = String::new();
    let mut ranges = Vec::with_capacity(fragments.len());
    // §17.3.2.30: the isolate currently open, if any. Tracked across fragments
    // so consecutive ones sharing a toggle share an isolate — otherwise a run
    // split by #130 into one fragment per word would fence each word off from
    // its neighbours and defeat the point of resolving over the paragraph.
    let mut open: Option<BaseDirection> = None;

    for fragment in fragments {
        match fragment {
            Fragment::Text { text: t, font, .. } => {
                let want = match font.rtl {
                    Toggle::On => Some(BaseDirection::Rtl),
                    Toggle::Off => Some(BaseDirection::Ltr),
                    Toggle::Absent => None,
                };
                if want != open {
                    if open.is_some() {
                        text.push(POP_ISOLATE);
                    }
                    if let Some(dir) = want {
                        text.push(dir.isolate());
                    }
                    open = want;
                }
                let start = text.len();
                text.push_str(t);
                ranges.push(Some(start..text.len()));
            }
            other => {
                // Non-text fragments stay inside whatever isolate is open: an
                // image in the middle of a `w:rtl` run belongs to that run.
                match other {
                    Fragment::Tab { .. } | Fragment::PTab { .. } => text.push('\t'),
                    Fragment::Image { .. } | Fragment::Emoji { .. } => text.push(OBJECT),
                    Fragment::LineBreak { .. }
                    | Fragment::PageBreak { .. }
                    | Fragment::ColumnBreak => text.push('\n'),
                    Fragment::Bookmark { .. } | Fragment::Text { .. } => {}
                }
                ranges.push(None);
            }
        }
    }
    if open.is_some() {
        text.push(POP_ISOLATE);
    }
    (text, ranges)
}

/// Emit `fragment` as one piece per run of equal level in `levels`.
///
/// `levels` is one entry per byte of the fragment's own text, so a boundary can
/// only fall where a level changes — and never inside a character, since every
/// byte of one carries the same level.
fn split_at_level_boundaries<F>(
    fragment: Fragment,
    levels: &[BidiLevel],
    measure_text: &F,
    out: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let Fragment::Text {
        text,
        font,
        color,
        shading,
        border,
        break_after,
        width,
        trimmed_width,
        metrics,
        hyperlink_url,
        baseline_offset,
        text_offset,
        is_footnote_ref,
        ..
    } = fragment
    else {
        out.push(fragment);
        return;
    };

    let runs = level_runs(levels);
    let mirrored = runs
        .iter()
        .any(|&(start, end, level)| level.is_rtl() && text[start..end].chars().any(has_mirror));

    // The common case even in a right-to-left paragraph: one level, nothing to
    // mirror. Stamp the level and keep every measurement already taken.
    if runs.len() == 1 && !mirrored {
        out.push(Fragment::Text {
            shaped: None,
            level: runs[0].2,
            text,
            font,
            color,
            shading,
            border,
            break_after,
            width,
            trimmed_width,
            metrics,
            hyperlink_url,
            baseline_offset,
            text_offset,
            is_footnote_ref,
        });
        return;
    }

    let last = runs.len() - 1;
    for (i, (start, end, level)) in runs.into_iter().enumerate() {
        let piece = mirror_text(&text[start..end], level);
        let (w, m) = measure_text(&piece, &font);
        let trimmed = piece.trim_end();
        let tw = if trimmed.len() < piece.len() {
            measure_text(trimmed, &font).0
        } else {
            w
        };
        out.push(Fragment::Text {
            shaped: None,
            level,
            text: Rc::from(&*piece),
            font: Rc::clone(&font),
            color,
            shading,
            border,
            // A level change is not a line-break opportunity — UAX #14 knew
            // nothing about it. Only the last piece may keep whatever
            // opportunity the whole fragment had earned at its trailing edge.
            break_after: if i == last {
                break_after
            } else {
                BreakAfter::Prohibited
            },
            width: w,
            trimmed_width: tw,
            metrics: m,
            hyperlink_url: hyperlink_url.clone(),
            baseline_offset,
            // §17.9.7: a list label's own justification offset belongs to the
            // label as a whole. A label that had to be split is the pathological
            // case; giving the offset to the first piece keeps the label's left
            // edge where it was rather than applying the shift once per piece.
            text_offset: if i == 0 { text_offset } else { Pt::ZERO },
            is_footnote_ref,
        });
    }
}

/// Byte ranges of maximal runs of equal level: `(start, end, level)`.
fn level_runs(levels: &[BidiLevel]) -> Vec<(usize, usize, BidiLevel)> {
    let mut runs = Vec::new();
    let mut start = 0;
    for i in 1..levels.len() {
        if levels[i] != levels[start] {
            runs.push((start, i, levels[start]));
            start = i;
        }
    }
    runs.push((start, levels.len(), levels[start]));
    runs
}

fn has_mirror(c: char) -> bool {
    bidi::mirror(c).is_some()
}

/// Rule L4: mirror every character that has a mirror, at an odd level.
///
/// Borrows unless something actually changes, so the ordinary case — text with
/// no paired punctuation in it — allocates nothing.
fn mirror_text(text: &str, level: BidiLevel) -> std::borrow::Cow<'_, str> {
    if !level.is_rtl() || !text.chars().any(has_mirror) {
        return std::borrow::Cow::Borrowed(text);
    }
    std::borrow::Cow::Owned(text.chars().map(|c| bidi::mirror(c).unwrap_or(c)).collect())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::resolve::color::RgbColor;

    fn font(rtl: Toggle) -> Rc<FontProps> {
        Rc::new(FontProps {
            effects: Default::default(),
            family: Rc::from("Test"),
            size: Pt::new(12.0),
            bold: Toggle::Absent,
            italic: Toggle::Absent,
            underline: false,
            rtl,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        })
    }

    /// One point per character, so a width assertion reads as a length.
    fn measure(text: &str, _: &FontProps) -> (Pt, TextMetrics) {
        (
            Pt::new(text.chars().count() as f32),
            TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
        )
    }

    fn text_frag(text: &str, rtl: Toggle) -> Fragment {
        let (width, metrics) = measure(text, &font(rtl));
        Fragment::Text {
            shaped: None,
            level: BidiLevel::LTR,
            text: Rc::from(text),
            font: font(rtl),
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            break_after: BreakAfter::Opportunity,
            width,
            trimmed_width: width,
            metrics,
            hyperlink_url: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    fn run(mut frags: Vec<Fragment>, base: BaseDirection) -> Vec<(String, u8, BreakAfter)> {
        assign_bidi_levels(&mut frags, base, &measure);
        frags
            .iter()
            .filter_map(|f| match f {
                Fragment::Text {
                    text,
                    level,
                    break_after,
                    ..
                } => Some((
                    text.to_string(),
                    if level.is_rtl() { 1 } else { 0 },
                    *break_after,
                )),
                _ => None,
            })
            .collect()
    }

    // ── the fast path ─────────────────────────────────────────────────────

    #[test]
    fn a_left_to_right_paragraph_with_no_rtl_text_is_untouched() {
        let mut frags = vec![
            text_frag("Nicht ", Toggle::Absent),
            text_frag("gefunden", Toggle::Absent),
        ];
        let before = format!("{frags:?}");
        assign_bidi_levels(&mut frags, BaseDirection::Ltr, &measure);
        assert_eq!(format!("{frags:?}"), before, "nothing may change");
    }

    // ── levels ────────────────────────────────────────────────────────────

    #[test]
    fn hebrew_words_come_back_at_an_odd_level() {
        assert_eq!(
            run(
                vec![
                    text_frag("a ", Toggle::Absent),
                    text_frag("שלום", Toggle::Absent)
                ],
                BaseDirection::Ltr,
            ),
            [
                ("a ".to_string(), 0, BreakAfter::Opportunity),
                ("שלום".to_string(), 1, BreakAfter::Opportunity),
            ],
        );
    }

    /// The reason the analysis string spans the whole paragraph: a word split
    /// across `<w:r>` boundaries must resolve as one, and a fragment in the
    /// middle must see the strong characters on both sides of it.
    #[test]
    fn levels_are_resolved_across_fragment_boundaries() {
        // "שלום" then " " then "עולם": the space between two Hebrew words is
        // a neutral that must take their level, not the paragraph's.
        let got = run(
            vec![
                text_frag("שלום", Toggle::Absent),
                text_frag(" ", Toggle::Absent),
                text_frag("עולם", Toggle::Absent),
            ],
            BaseDirection::Ltr,
        );
        assert_eq!(
            got.iter().map(|(_, l, _)| *l).collect::<Vec<_>>(),
            [1, 1, 1],
            "the space between two Hebrew words is Hebrew",
        );
    }

    // ── splitting ─────────────────────────────────────────────────────────

    /// A fragment a level boundary falls inside cannot be placed as one unit,
    /// so it is cut — and the cut is not a line-break opportunity.
    #[test]
    fn a_fragment_spanning_a_level_boundary_is_split() {
        assert_eq!(
            run(
                vec![text_frag("abשלום", Toggle::Absent)],
                BaseDirection::Ltr
            ),
            [
                ("ab".to_string(), 0, BreakAfter::Prohibited),
                ("שלום".to_string(), 1, BreakAfter::Opportunity),
            ],
            "earlier pieces are prohibited; the last keeps the parent's status",
        );
    }

    #[test]
    fn split_pieces_are_remeasured() {
        let mut frags = vec![text_frag("abשלום", Toggle::Absent)];
        assign_bidi_levels(&mut frags, BaseDirection::Ltr, &measure);
        let widths: Vec<f32> = frags
            .iter()
            .map(|f| match f {
                Fragment::Text { width, .. } => width.raw(),
                _ => unreachable!(),
            })
            .collect();
        assert_eq!(widths, [2.0, 4.0], "each piece carries its own width");
    }

    // ── §17.3.2.30 w:rtl ──────────────────────────────────────────────────

    /// A run of neutrals marked `w:rtl` resolves right-to-left, which is the
    /// whole of what the toggle buys over letting the characters speak.
    #[test]
    fn an_rtl_run_of_neutrals_resolves_right_to_left() {
        let got = run(
            vec![
                text_frag("a ", Toggle::Absent),
                text_frag("(1)", Toggle::On),
                text_frag(" b", Toggle::Absent),
            ],
            BaseDirection::Ltr,
        );
        assert_eq!(got[1].1, 1, "the marked run is right-to-left: {got:?}");
        assert_eq!(got[0].1, 0, "and its neighbours are not");
    }

    /// Consecutive fragments of one `w:rtl` run share an isolate. If each were
    /// fenced separately, a word #130 split into per-word fragments would stop
    /// resolving as one piece of text.
    #[test]
    fn consecutive_rtl_fragments_share_one_isolate() {
        let got = run(
            vec![
                text_frag("שלום ", Toggle::On),
                text_frag("עולם", Toggle::On),
            ],
            BaseDirection::Ltr,
        );
        assert_eq!(
            got.iter().map(|(_, l, _)| *l).collect::<Vec<_>>(),
            [1, 1],
            "both halves of the run resolve together",
        );
    }

    // ── rule L4 ───────────────────────────────────────────────────────────

    #[test]
    fn brackets_mirror_at_an_odd_level() {
        let got = run(
            vec![text_frag("(שלום)", Toggle::Absent)],
            BaseDirection::Rtl,
        );
        let joined: String = got.iter().map(|(t, ..)| t.as_str()).collect();
        assert_eq!(
            joined, ")שלום(",
            "an opening bracket at an odd level paints as a closing one",
        );
    }

    #[test]
    fn brackets_do_not_mirror_at_an_even_level() {
        let got = run(
            vec![
                text_frag("שלום ", Toggle::Absent),
                text_frag("(ok)", Toggle::Absent),
            ],
            BaseDirection::Ltr,
        );
        assert!(
            got.iter().any(|(t, ..)| t == "(ok)"),
            "left-to-right text keeps its brackets: {got:?}",
        );
    }

    // ── non-text fragments ────────────────────────────────────────────────

    /// An image is a neutral (U+FFFC, class ON), so the text on either side of
    /// it must resolve as though it were there — and it must survive the pass.
    #[test]
    fn non_text_fragments_survive_and_stand_in_as_neutrals() {
        let mut frags = vec![
            text_frag("שלום", Toggle::Absent),
            Fragment::Bookmark {
                name: "b".to_string(),
            },
            text_frag("עולם", Toggle::Absent),
        ];
        assign_bidi_levels(&mut frags, BaseDirection::Rtl, &measure);
        assert_eq!(frags.len(), 3);
        assert!(matches!(frags[1], Fragment::Bookmark { .. }));
    }

    #[test]
    fn an_empty_paragraph_is_handled() {
        let mut frags: Vec<Fragment> = Vec::new();
        assign_bidi_levels(&mut frags, BaseDirection::Rtl, &measure);
        assert!(frags.is_empty());
    }
}
