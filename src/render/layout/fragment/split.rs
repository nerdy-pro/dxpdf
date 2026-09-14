//! Cluster-level splitting of over-wide text fragments.
//!
//! A word wider than the space available can't be broken at a normal break
//! opportunity, so it is split into one fragment per grapheme cluster and the
//! line-fitter breaks between them. Used for narrow table cells and deep
//! indents.
//!
//! The unit is the same one spacing is inserted between
//! ([`crate::render::spacing`]) — and for the same reason. Splitting per
//! *scalar*, which this module used to do, let the line-fitter break between a
//! letter and its own combining mark and carry the accent to the next line.

use std::rc::Rc;

use super::{BreakAfter, FontProps, Fragment, TextMetrics};
use crate::render::dimension::Pt;
use crate::render::spacing;

/// The measuring capability the split (and the paragraph layer's tab
/// machinery) is handed — `paragraph::MeasureTextFn` aliases the same trait
/// object, defined here so neither layer depends on the other.
///
/// Two operations because a piece of a shaped fragment must be measured the
/// way it will be painted (issue #153): the painter re-shapes each piece
/// through HarfBuzz, and for an Indic conjunct — one grapheme, so one piece —
/// the cmap sum of its consonants can run half again wider than the ligature
/// the shaper draws. Measuring such a piece with [`SplitMeasure::measure`]
/// alone reserves the cmap width and strands the underline past the glyphs.
pub trait SplitMeasure {
    /// Cmap measurement — `TextMeasurer::measure`: advance (scaled, with
    /// §17.3.2.35 spacing per grapheme cluster) and font metrics.
    fn measure(&self, text: &str, font: &FontProps) -> (Pt, TextMetrics);

    /// Shaped measurement for one piece of a shaped fragment —
    /// `TextMeasurer::shaped_measurement`: the HarfBuzz advance (scaled, no
    /// spacing) and the piece's shaped-cluster count. `None` when shaping is
    /// unavailable or failed; the caller falls back to [`Self::measure`],
    /// which is also what the painter's cmap fallback will draw.
    fn shaped_piece(
        &self,
        text: &str,
        font: &FontProps,
        direction: crate::render::shape::RunDirection,
    ) -> Option<(Pt, usize)> {
        let _ = (text, font, direction);
        None
    }
}

/// Adapter for tests and callers that measure with a plain closure — the
/// closure supplies the cmap half and [`SplitMeasure::shaped_piece`] keeps
/// its `None` default.
pub struct MeasureWith<F>(pub F);

impl<F: Fn(&str, &FontProps) -> (Pt, TextMetrics)> SplitMeasure for MeasureWith<F> {
    fn measure(&self, text: &str, font: &FontProps) -> (Pt, TextMetrics) {
        (self.0)(text, font)
    }
}

/// Optional measuring capability, `None` in width-approximating tests.
pub type MeasureFn<'a> = Option<&'a dyn SplitMeasure>;

/// True for a text fragment that is both too wide and actually splittable.
///
/// "Too wide" is measured **without trailing whitespace**, the same width line
/// fitting checks overflow against: trailing space is allowed to hang past the
/// margin, so a fragment whose visible text fits is not over-wide however much
/// space trails it. Documents lay out with runs of spaces — one in
/// `test-cases/` pads a header with 130 of them — and UAX #14 keeps such a run
/// in a single unit, because [LB7] forbids a break *before* a space. Measuring
/// the full width would send that unit here to be cut into 130 one-space
/// fragments, and the visible word ahead of them would be dragged onto its own
/// line for want of room that was never needed.
///
/// The *cluster* count is what matters for splittability: a single cluster
/// cannot be split however many scalars or bytes it occupies. An earlier
/// byte-length (`text.len() > 1`) spelling of this test disagreed with the
/// split itself for any non-ASCII single character — it reported "needs
/// split", then split nothing, and the caller paid for a full clone of the
/// fragment vector. A scalar count has the same disagreement for a
/// one-cluster accented letter.
///
/// [LB7]: https://www.unicode.org/reports/tr14/#LB7
fn needs_split(fragment: &Fragment, max_width: Pt) -> bool {
    matches!(
        fragment,
        Fragment::Text { trimmed_width, text, .. }
            if *trimmed_width > max_width && spacing::unit_count(text) > 1
    )
}

/// Split text fragments wider than `max_width` into per-cluster fragments.
///
/// Returns `None` when nothing needs splitting — the common case. That lets a
/// caller holding an owned `Vec` keep it untouched, and a caller holding a
/// slice hand back `Cow::Borrowed`; neither pays for a copy on the fast path.
///
/// Per-cluster widths come from `measure` when one is supplied. Without a
/// measurer the fragment's width is divided evenly across its clusters,
/// which is only a positioning approximation — the total is preserved.
pub fn split_oversized_fragments(
    fragments: &[Fragment],
    max_width: Pt,
    measure: MeasureFn<'_>,
) -> Option<Vec<Fragment>> {
    // A non-positive budget can't be met by any split, and dividing by it
    // below would be meaningless.
    if max_width <= Pt::ZERO {
        return None;
    }
    if !fragments.iter().any(|f| needs_split(f, max_width)) {
        return None;
    }

    let mut result = Vec::with_capacity(fragments.len());
    for frag in fragments {
        let Fragment::Text {
            text,
            width,
            font,
            color,
            shading,
            border,
            metrics,
            hyperlink_url,
            baseline_offset,
            level,
            shaped,
            ..
        } = frag
        else {
            result.push(frag.clone());
            continue;
        };
        if !needs_split(frag, max_width) {
            result.push(frag.clone());
            continue;
        }

        let unit_count = spacing::unit_count(text);
        let per_unit_fallback = *width / unit_count as f32;
        for unit in spacing::units(text) {
            // A piece of a shaped fragment is measured the way the painter
            // will draw it — shaped. The cut still costs an Arabic word its
            // joining (each piece re-shapes in isolation, the documented last
            // resort), but a piece that is itself a ligature — an Indic
            // conjunct is one grapheme, hence one piece — keeps its shaped
            // width, where the cmap sum of its consonants can be half again
            // wider than the glyph and would strand the underline past it.
            let shaped_piece = match (measure, shaped) {
                (Some(m), Some(s)) => m.shaped_piece(unit, font, s.direction),
                _ => None,
            };
            let (w, unit_metrics, piece_units) = match (shaped_piece, measure) {
                (Some((advance, units)), Some(m)) => (
                    // §17.3.2.35: spacing per shaped cluster, the same
                    // formula `fragment::shape` charges a whole fragment.
                    advance + font.char_spacing * units as f32,
                    m.measure(unit, font).1,
                    units,
                ),
                (None, Some(m)) => {
                    let (w, metrics) = m.measure(unit, font);
                    (w, metrics, 1)
                }
                (_, None) => (per_unit_fallback, *metrics, 1),
            };
            result.push(Fragment::Text {
                // The unit count is re-stamped, not copied: the piece's own
                // shaped-cluster count when shaping measured it, else one —
                // a piece is one grapheme cluster, and a shaped cluster is
                // never finer than one where UAX #29 GB9c holds (it does not
                // in Tamil, Kannada, Gurmukhi or Sinhala, whose viramas the
                // pinned segmenter leaves outside InCB — there a conjunct
                // spans two pieces and the cut degrades it to halant forms,
                // the same accepted degradation as the lost Arabic joining).
                shaped: shaped.map(|s| super::Shaping {
                    unit_count: piece_units,
                    ..s
                }),
                // UAX #9: a cluster of the parent's text is at the parent's
                // level. Cutting a word never crosses a level boundary —
                // `fragment::bidi` has already split at every one of those, so
                // whatever reaches here is level-uniform by the time it does.
                level: *level,
                text: Rc::from(unit),
                font: font.clone(),
                color: *color,
                shading: *shading,
                border: *border,
                // Still prohibited, even though this split exists precisely to
                // let the fitter break here. The fitter breaks at the *last*
                // opportunity before an overflow, so calling each cluster one
                // would drag the head of the word up onto the previous line —
                // "Nicht gefunden" in a narrow cell becomes "Nicht ge" /
                // "funden" — when the word would have fitted whole on the next
                // line. With no opportunity to fall back to, the fitter breaks
                // immediately before whichever cluster overflows, which is only
                // reached once the word has a line to itself and still doesn't
                // fit. That is the order Word applies too: move the word down
                // first, cut it only if it still won't fit.
                break_after: BreakAfter::Prohibited,
                width: w,
                trimmed_width: w,
                metrics: unit_metrics,
                hyperlink_url: hyperlink_url.clone(),
                baseline_offset: *baseline_offset,
                text_offset: Pt::ZERO,
                // Per-cluster split of an over-wide word — not a mark.
                is_footnote_ref: false,
            });
        }
    }
    Some(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::resolve::color::RgbColor;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: Rc::from(text),
            break_after: super::super::fixture_break_after(text),
            font: Rc::new(FontProps {
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            }),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    fn texts(frags: &[Fragment]) -> Vec<String> {
        frags
            .iter()
            .map(|f| match f {
                Fragment::Text { text, .. } => text.to_string(),
                _ => panic!("expected Text fragment"),
            })
            .collect()
    }

    /// Issue #153: a piece of a shaped fragment is one grapheme cluster, and a
    /// shaped cluster is never finer than one — so the piece's spacing-unit
    /// count is re-stamped to 1, not inherited from the whole word, or every
    /// piece would claim the parent's full complement of distribution gaps.
    #[test]
    fn a_shaped_fragment_splits_into_single_unit_pieces() {
        let mut frags = vec![text_frag("مرحبا", 60.0)];
        if let Fragment::Text { shaped, .. } = &mut frags[0] {
            *shaped = Some(super::super::Shaping {
                direction: crate::render::shape::RunDirection::RightToLeft,
                unit_count: 5,
            });
        }
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(result.len(), 5);
        for frag in &result {
            let Fragment::Text {
                shaped: Some(s), ..
            } = frag
            else {
                panic!("pieces stay marked shaped");
            };
            assert_eq!(s.unit_count, 1, "one grapheme piece, one spacing unit");
            assert_eq!(
                s.direction,
                crate::render::shape::RunDirection::RightToLeft,
                "direction survives the split"
            );
        }
    }

    /// The review finding behind the `shaped_piece` operation: a piece of a
    /// shaped fragment must get its *shaped* width, not the cmap sum — for a
    /// Devanagari conjunct the two differ by the whole ligature saving, and
    /// the painter draws the shaped form. A measurer offering both answers
    /// (cmap 30pt, shaped 17pt per piece, 1 cluster) must see the shaped one
    /// win, with §17.3.2.35 spacing charged per shaped cluster on top.
    #[test]
    fn a_shaped_piece_is_measured_shaped_not_by_cmap() {
        struct FakeMeasure;
        impl SplitMeasure for FakeMeasure {
            fn measure(&self, _: &str, _: &FontProps) -> (Pt, TextMetrics) {
                (
                    Pt::new(30.0),
                    TextMetrics {
                        ascent: Pt::new(9.0),
                        descent: Pt::new(3.0),
                        leading: Pt::ZERO,
                    },
                )
            }
            fn shaped_piece(
                &self,
                _: &str,
                _: &FontProps,
                _: crate::render::shape::RunDirection,
            ) -> Option<(Pt, usize)> {
                Some((Pt::new(17.0), 1))
            }
        }

        let mut frags = vec![text_frag("क्षमता", 90.0)];
        if let Fragment::Text { shaped, font, .. } = &mut frags[0] {
            *shaped = Some(super::super::Shaping {
                direction: crate::render::shape::RunDirection::LeftToRight,
                unit_count: 3,
            });
            *font = Rc::new(FontProps {
                char_spacing: Pt::new(2.0),
                ..(**font).clone()
            });
        }
        let result =
            split_oversized_fragments(&frags, Pt::new(20.0), Some(&FakeMeasure)).expect("splits");
        assert_eq!(result.len(), 3, "क्ष / म / ता — three grapheme pieces");
        for frag in &result {
            let Fragment::Text {
                width,
                shaped: Some(s),
                ..
            } = frag
            else {
                panic!("pieces stay marked shaped");
            };
            assert!(
                (width.raw() - 19.0).abs() < 1e-4,
                "shaped 17pt + 2pt spacing × 1 cluster, not cmap 30pt; got {width:?}"
            );
            assert_eq!(s.unit_count, 1);
        }

        // The same fragment with no shaped mark keeps the cmap measurement —
        // shaped_piece must not leak onto the cmap path.
        let mut plain = frags.clone();
        if let Fragment::Text { shaped, .. } = &mut plain[0] {
            *shaped = None;
        }
        let result =
            split_oversized_fragments(&plain, Pt::new(20.0), Some(&FakeMeasure)).expect("splits");
        for frag in &result {
            let Fragment::Text { width, shaped, .. } = frag else {
                unreachable!();
            };
            assert!(shaped.is_none());
            assert!(
                (width.raw() - 30.0).abs() < 1e-4,
                "cmap width, got {width:?}"
            );
        }
    }

    #[test]
    fn splits_into_one_fragment_per_character() {
        // "ab" at 60pt is wider than max_width=20pt → two 30pt characters
        // (uniform fallback — no measurer provided).
        let frags = vec![text_frag("ab", 60.0)];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(
            texts(&result),
            ["a", "b"],
            "one fragment per character, in order"
        );
        for frag in &result {
            let Fragment::Text { width, .. } = frag else {
                unreachable!()
            };
            assert!((width.raw() - 30.0).abs() < 1e-4, "uniform fallback 60/2");
        }
    }

    #[test]
    fn measurer_supplies_per_character_widths() {
        // A measurer that gives 'w' a wider advance than 'i' — the uniform
        // fallback would give both the same width.
        let measure = |t: &str, _: &FontProps| {
            let w = if t == "w" { 40.0 } else { 10.0 };
            (
                Pt::new(w),
                TextMetrics {
                    ascent: Pt::new(9.0),
                    descent: Pt::new(3.0),
                    leading: Pt::ZERO,
                },
            )
        };
        let frags = vec![text_frag("wi", 50.0)];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), Some(&MeasureWith(measure)))
            .expect("splits");
        let widths: Vec<f32> = result
            .iter()
            .map(|f| match f {
                Fragment::Text { width, .. } => width.raw(),
                _ => unreachable!(),
            })
            .collect();
        assert_eq!(widths, [40.0, 10.0], "measured, not evenly divided");
        // Metrics come from the measurer too, not the parent fragment.
        let Fragment::Text { metrics, .. } = &result[0] else {
            unreachable!()
        };
        assert_eq!(metrics.ascent.raw(), 9.0);
    }

    #[test]
    fn nothing_to_split_returns_none() {
        let frags = vec![text_frag("hi", 10.0)];
        assert!(split_oversized_fragments(&frags, Pt::new(100.0), None).is_none());
    }

    #[test]
    fn single_character_is_never_split() {
        // Wider than max_width, but one character — nothing to break.
        let frags = vec![text_frag("M", 200.0)];
        assert!(split_oversized_fragments(&frags, Pt::new(10.0), None).is_none());
    }

    /// The character count, not the byte length, decides splittability. A
    /// byte-length test reports "needs split" for these and then splits
    /// nothing, costing the caller a pointless clone of the whole vector.
    #[test]
    fn multi_byte_single_character_is_never_split() {
        for ch in ["é", "😀", "字"] {
            let frags = vec![text_frag(ch, 200.0)];
            assert!(
                split_oversized_fragments(&frags, Pt::new(10.0), None).is_none(),
                "{ch:?} is one character regardless of its byte length"
            );
        }
    }

    /// A unit whose *visible* text fits is not over-wide, however much space
    /// trails it. Real documents pad with runs of spaces, and UAX #14 keeps a
    /// run of them in one unit ([LB7] forbids a break before a space) — so
    /// this fragment reaches the splitter looking 200pt wide when only 5pt of
    /// it will ever be drawn inside the margin.
    ///
    /// [LB7]: https://www.unicode.org/reports/tr14/#LB7
    #[test]
    fn trailing_whitespace_does_not_make_a_fragment_over_wide() {
        let mut frag = text_frag("H                    ", 200.0);
        if let Fragment::Text { trimmed_width, .. } = &mut frag {
            *trimmed_width = Pt::new(5.0);
        }
        assert!(
            split_oversized_fragments(&[frag], Pt::new(20.0), None).is_none(),
            "only the visible width counts; the spaces hang past the margin",
        );
    }

    #[test]
    fn non_positive_max_width_returns_none() {
        let frags = vec![text_frag("ab", 60.0)];
        assert!(split_oversized_fragments(&frags, Pt::ZERO, None).is_none());
        assert!(split_oversized_fragments(&frags, Pt::new(-5.0), None).is_none());
    }

    #[test]
    fn fragments_that_fit_pass_through_unchanged() {
        let frags = vec![
            text_frag("ab", 60.0), // splits
            text_frag("ok", 5.0),  // fits — must survive whole
        ];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(texts(&result), ["a", "b", "ok"]);
    }

    /// A combining mark must travel with its base. Splitting per scalar put the
    /// accent in its own fragment, which the line-fitter is then free to break
    /// before — carrying a bare `U+0301` to the start of the next line.
    #[test]
    fn split_never_separates_a_combining_mark_from_its_base() {
        let frags = vec![text_frag("e\u{301}x", 60.0)];
        let result = split_oversized_fragments(&frags, Pt::new(20.0), None).expect("splits");
        assert_eq!(
            texts(&result),
            ["e\u{301}", "x"],
            "the accent stays in the same fragment as its base letter"
        );
    }

    /// The cluster count decides splittability, so a single accented letter is
    /// as unsplittable as a single plain one — `needs_split` must not report
    /// work it cannot then do.
    #[test]
    fn multi_scalar_single_cluster_is_never_split() {
        for text in ["e\u{301}", "1\u{FE0F}\u{20E3}", "\u{1F1E9}\u{1F1EA}"] {
            let frags = vec![text_frag(text, 200.0)];
            assert!(
                split_oversized_fragments(&frags, Pt::new(10.0), None).is_none(),
                "{text:?} is one grapheme cluster regardless of its scalar count"
            );
        }
    }
}
