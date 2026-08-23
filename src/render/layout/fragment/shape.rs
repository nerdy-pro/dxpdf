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
//! Nothing. Neither reason can fire for a document with no right-to-left text
//! in it: `needs_shaping` is false for Latin, Cyrillic, Greek, CJK, Hebrew and
//! Thai, and every fragment in such a document is at an even level because
//! `super::bidi` returned without touching it. So this pass walks the vector,
//! finds nothing, and returns. That is what keeps such a document's pixels
//! identical — shaping applies GPOS kerning and standard ligatures as well, so
//! a shaped Latin corpus would reflow everywhere.
//!
//! One thing it does not do is §17.3.2.35 `w:spacing`. Inserting space
//! *between* units of a shaped run means knowing where the shaped cluster
//! boundaries are, and [`crate::render::spacing`] — which owns that question —
//! answers it in UAX #29 grapheme clusters, on the stated grounds that the
//! painter cannot honour anything finer. For a shaped run it now could, and
//! that module's own doc names itself as the seam that changes when the paint
//! path shapes. Until it does, a `w:spacing` run in Arabic is measured and
//! painted without the spacing rather than with it applied at boundaries that
//! do not survive shaping.

use crate::render::layout::measurer::TextMeasurer;
use crate::render::shape::{needs_shaping, RunDirection};

use super::Fragment;

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
        *shaped = Some(direction);

        // A shaping failure leaves the cmap-measured width in place. The run
        // then measures and paints the same way it did before issue #131 —
        // wrong in the way it was already wrong, rather than measured one way
        // and painted another, which is the failure that strands an underline.
        if let Some(w) = measurer.shaped_advance(text, font, direction) {
            // Trailing whitespace is not part of a joining sequence, so the
            // trimmed width scales with the same ratio rather than costing a
            // second shaping pass over a string that differs only in spaces.
            let ratio = if *width > crate::render::dimension::Pt::ZERO {
                *trimmed_width / *width
            } else {
                1.0
            };
            *width = w;
            *trimmed_width = w * ratio;
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

    fn frag(text: &str, level: BidiLevel) -> Fragment {
        Fragment::Text {
            shaped: None,
            level,
            text: Rc::from(text),
            font: Rc::new(FontProps {
                effects: Default::default(),
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                rtl: Toggle::Absent,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            }),
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
            Fragment::Text { shaped, .. } => *shaped,
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
}
