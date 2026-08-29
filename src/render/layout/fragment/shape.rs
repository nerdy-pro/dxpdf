//! Marking the fragments that must be shaped, and re-measuring them.
//!
//! The counterpart to [`crate::render::shape`] on the fragment side, and the
//! same division [`crate::i18n::bidi`] and [`super::bidi`] keep: one module
//! knows *how*, this one knows *which*, and *what to do with the answer*.
//!
//! # Two reasons a run cannot be painted from its string
//!
//! The painter's ordinary path walks a string and places one glyph per
//! codepoint, left to right. A run goes through the shaper instead when that
//! walk would put a glyph in the wrong place, and there are exactly two ways
//! it can:
//!
//! 1. **The script joins.** A letter's glyph depends on its neighbours, which
//!    a cmap lookup cannot see — [`crate::render::shape::needs_shaping`] is
//!    that test, and Arabic is the case that motivates it.
//! 2. **The run is right-to-left.** Its first character belongs at the run's
//!    *right* edge, and walking the string left to right puts it at the left.
//!    Measured, not assumed: `draw_str` paints `שלום ` as
//!    `[ש, ל, ו, ם, space]` from x=0, where the shaper returns
//!    `[space, ם, ו, ל, ש]` — every glyph on the wrong side, and the trailing
//!    space at the wrong end of the word.
//!
//! The second is why Hebrew is in scope here even though it never joins. Rule
//! L2 reordering (`layout::paragraph::line_emit`) puts a line's *fragments* in
//! visual order; nothing but the shaper puts the glyphs *inside* one there.
//!
//! # What this costs the rest of the corpus
//!
//! Nothing, for a document with neither right-to-left nor reordering-script
//! text in it: `needs_shaping` is false for Latin, Cyrillic, Greek, CJK,
//! Hebrew and Thai, and every fragment in such a document is at an even level
//! because `super::bidi` returned without touching it. So this pass walks the
//! vector, finds nothing, and returns. That is what keeps such a document's
//! pixels identical — shaping applies GPOS kerning and standard ligatures as
//! well, so a shaped Latin corpus would reflow everywhere. A document that
//! *does* hold Devanagari (or another `script_reorders` script) now pays for
//! shaping those runs even with no right-to-left text anywhere — that is
//! issue #153's point, not a regression: their cmap widths were widths of the
//! wrong glyph sequence.
//!
//! # §17.3.2.35 `w:spacing` on a shaped run
//!
//! Applied here, at the shaped cluster (issue #153). The cmap measurement this
//! pass overwrites had added `char_spacing` per UAX #29 grapheme cluster; a
//! shaped run's spacing units are the clusters HarfBuzz reports instead — a
//! conjunct's glyphs share one, a ligature merges what were separate grapheme
//! clusters — so the re-measure swaps both terms: the advance *and* the unit
//! count the spacing multiplies. The count is stored on the fragment
//! ([`super::Shaping::unit_count`]) so that §17.3.1.13 distribution counts the
//! same units later without re-shaping, and the painter groups its own shaped
//! glyphs by cluster into the same number of units — see
//! [`crate::render::spacing`] for why all three must be one number.

use crate::render::layout::measurer::TextMeasurer;
use crate::render::shape::{needs_shaping, RunDirection};

use super::{Fragment, Shaping};

/// Mark every fragment that cannot be painted from its string, and re-measure
/// it against the shaped advance.
///
/// Call once per paragraph, immediately after [`super::assign_bidi_levels`] —
/// the level each fragment carries by then is what decides both *whether* it
/// is shaped and *which way round*, and a fragment that pass split has each
/// half measured on its own.
pub fn shape_complex_scripts(fragments: &mut [Fragment], measurer: &TextMeasurer) {
    for fragment in fragments {
        let Fragment::Text {
            text,
            font,
            level,
            shaped,
            width,
            trimmed_width,
            ..
        } = fragment
        else {
            continue;
        };
        // Either reason is enough; see this module's doc for why they are two
        // and not one.
        if !level.is_rtl() && !needs_shaping(text) {
            continue;
        }

        let direction = RunDirection::from(*level);

        // A shaping failure leaves the cmap-measured width in place. The run
        // then measures and paints the same way it did before issue #131 —
        // wrong in the way it was already wrong, rather than measured one way
        // and painted another, which is the failure that strands an underline.
        // The grapheme-cluster unit count is kept for the same reason: it is
        // the count the cmap width already charged spacing for, and the count
        // the painter's cmap fallback will draw.
        if let Some(m) = measurer.shaped_measurement(text, font, direction) {
            // §17.3.2.35: the cmap measurement charged `char_spacing` per
            // grapheme cluster; the shaped run's unit is the shaped cluster,
            // so the spacing term is rebuilt with the shaped count.
            let spacing = font.char_spacing * (m.unit_count as f32);
            // Trailing whitespace is not part of a joining sequence and never
            // merges into a shaped cluster, so its cost — cmap advance plus
            // its own spacing units, whichever model counted them — carries
            // over as the same absolute difference, without a second shaping
            // pass over a string that differs only in spaces. The difference
            // is legitimately negative under strongly condensed §17.3.2.35
            // spacing (each trailing space then *shrinks* the run by more
            // than its own advance), so it is carried signed: clamping it
            // would hand the trimmed width the clamp's error and let line
            // fitting admit a fragment whose visible glyphs overrun.
            let trailing = *width - *trimmed_width;
            *width = m.advance + spacing;
            *trimmed_width = *width - trailing;
            *shaped = Some(Shaping {
                direction,
                unit_count: m.unit_count,
            });
        } else {
            *shaped = Some(Shaping {
                direction,
                unit_count: crate::render::spacing::unit_count(text),
            });
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::i18n::bidi::BidiLevel;
    use crate::render::dimension::Pt;
    use crate::render::fonts::{FontRegistry, Toggle};
    use crate::render::layout::fragment::{BreakAfter, FontProps, TextMetrics};
    use crate::render::resolve::color::RgbColor;
    use skia_safe::FontMgr;
    use std::rc::Rc;

    fn test_font(char_spacing: Pt) -> FontProps {
        FontProps {
            family: Rc::from("Test"),
            size: Pt::new(12.0),
            bold: Toggle::Absent,
            italic: Toggle::Absent,
            underline: false,
            rtl: Toggle::Absent,
            char_spacing,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        }
    }

    fn frag(text: &str, level: BidiLevel) -> Fragment {
        Fragment::Text {
            shaped: None,
            level,
            text: Rc::from(text),
            font: Rc::new(test_font(Pt::ZERO)),
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            break_after: BreakAfter::Opportunity,
            width: Pt::new(50.0),
            trimmed_width: Pt::new(50.0),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    fn shaped_of(f: &Fragment) -> Option<RunDirection> {
        match f {
            Fragment::Text { shaped, .. } => shaped.map(|s| s.direction),
            _ => None,
        }
    }

    #[test]
    fn latin_is_never_marked_for_shaping() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let mut frags = vec![frag("Nicht gefunden", BidiLevel::LTR)];
        shape_complex_scripts(&mut frags, &measurer);
        assert_eq!(shaped_of(&frags[0]), None);
    }

    /// Hebrew never joins, so `needs_shaping` is false for it — and it is
    /// shaped anyway, because a right-to-left run's glyphs have to come out
    /// in the opposite order to its string. This is reason 2.
    #[test]
    fn a_right_to_left_run_is_shaped_even_though_its_script_never_joins() {
        assert!(
            !crate::render::shape::needs_shaping("שלום"),
            "the premise: Hebrew has no positional forms",
        );
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let mut frags = vec![frag("שלום", BidiLevel::from_number(1))];
        shape_complex_scripts(&mut frags, &measurer);
        assert_eq!(shaped_of(&frags[0]), Some(RunDirection::RightToLeft));
    }

    /// And the same text at an even level is not — a Hebrew word inside a
    /// left-to-right run reads the way the string does.
    #[test]
    fn the_same_text_at_an_even_level_is_left_alone() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let mut frags = vec![frag("שלום", BidiLevel::LTR)];
        shape_complex_scripts(&mut frags, &measurer);
        assert_eq!(shaped_of(&frags[0]), None);
    }

    /// The direction comes from the resolved level, not from the text — which
    /// is why this pass runs after `bidi` and not instead of it.
    #[test]
    fn the_direction_comes_from_the_resolved_level() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let mut frags = vec![
            frag("مرحبا", BidiLevel::from_number(1)),
            frag("مرحبا", BidiLevel::from_number(2)),
        ];
        shape_complex_scripts(&mut frags, &measurer);
        assert_eq!(shaped_of(&frags[0]), Some(RunDirection::RightToLeft));
        assert_eq!(
            shaped_of(&frags[1]),
            Some(RunDirection::LeftToRight),
            "an even level lays out left to right whatever its script",
        );
    }

    /// Issue #153: a Devanagari fragment — LTR, no positional forms — is now
    /// marked for shaping. This is the predicate half; the reorder itself is
    /// pinned against a real face in `render::shape`'s tests.
    #[test]
    fn a_reordering_script_is_marked_for_shaping_at_an_even_level() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let mut frags = vec![frag("हिन्दी", BidiLevel::LTR)];
        shape_complex_scripts(&mut frags, &measurer);
        assert_eq!(shaped_of(&frags[0]), Some(RunDirection::LeftToRight));
    }

    /// §17.3.2.35 on a shaped run (issue #153): the re-measured width is the
    /// shaped advance plus `char_spacing` × the *shaped* cluster count, and
    /// that count is stored on the fragment for distribution to reuse. The
    /// text is lam + alef — Unicode's mandatory ligature, two grapheme
    /// clusters that shape to one — precisely so the shaped count *differs*
    /// from the grapheme count and a formula charging the wrong unit cannot
    /// pass. Asserted against `shaped_measurement` itself, so no glyph metric
    /// is pinned; a host that cannot shape the run, or whose face leaves the
    /// ligature unformed, skips.
    #[test]
    fn spacing_on_a_shaped_run_is_charged_per_shaped_cluster() {
        const LAM_ALEF: &str = "\u{0644}\u{0627}";
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let spacing = Pt::new(2.0);
        // The fixture family resolves to the host default, which need not
        // cover Arabic — ask the host which family does, exactly as the
        // per-glyph fallback pass would have before this pass runs.
        let Some(arabic_family) = FontMgr::new()
            .match_family_style_character(
                "",
                skia_safe::FontStyle::normal(),
                &[],
                '\u{0644}' as i32,
            )
            .map(|tf| tf.family_name())
        else {
            eprintln!("skipping: no face on this host covers Arabic");
            return;
        };
        let font = FontProps {
            family: Rc::from(arabic_family.as_str()),
            ..test_font(spacing)
        };
        let mut frags = vec![frag(LAM_ALEF, BidiLevel::from_number(1))];
        if let Fragment::Text { font: f, .. } = &mut frags[0] {
            *f = Rc::new(font.clone());
        }
        let Some(m) = measurer.shaped_measurement(LAM_ALEF, &font, RunDirection::RightToLeft)
        else {
            eprintln!("skipping: this host cannot shape the run");
            return;
        };
        assert_eq!(crate::render::spacing::unit_count(LAM_ALEF), 2);
        if m.unit_count != 1 {
            eprintln!("skipping: this host's face leaves lam-alef unligated");
            return;
        }
        shape_complex_scripts(&mut frags, &measurer);
        let Fragment::Text {
            shaped: Some(s),
            width,
            ..
        } = &frags[0]
        else {
            panic!("fragment must be marked shaped");
        };
        assert_eq!(s.unit_count, 1, "stored count is the shaped count");
        let expected = m.advance + spacing;
        assert!(
            (f32::from(*width) - f32::from(expected)).abs() < 0.01,
            "width {width:?} must be advance {:?} + spacing × 1 — spacing must \
             be charged per shaped cluster, not per grapheme cluster",
            m.advance,
        );
    }

    /// Trailing whitespace carries over as the same absolute cost: the gap
    /// between `width` and `trimmed_width` survives the shaped re-measure
    /// unchanged, because spaces neither join nor merge into shaped clusters.
    #[test]
    fn the_trailing_whitespace_cost_survives_the_re_measure() {
        let registry = FontRegistry::new(FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let font = test_font(Pt::ZERO);
        let mut frags = vec![frag("مرحبا ", BidiLevel::from_number(1))];
        if let Fragment::Text {
            width,
            trimmed_width,
            ..
        } = &mut frags[0]
        {
            *width = Pt::new(50.0);
            *trimmed_width = Pt::new(41.5);
        }
        if measurer
            .shaped_measurement("مرحبا ", &font, RunDirection::RightToLeft)
            .is_none()
        {
            eprintln!("skipping: this host cannot shape the run");
            return;
        }
        shape_complex_scripts(&mut frags, &measurer);
        let Fragment::Text {
            width,
            trimmed_width,
            ..
        } = &frags[0]
        else {
            unreachable!();
        };
        assert!(
            (f32::from(*width - *trimmed_width) - 8.5).abs() < 0.01,
            "trailing gap must stay 8.5pt, got {:?}",
            *width - *trimmed_width,
        );
    }
}
