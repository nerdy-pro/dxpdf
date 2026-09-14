//! Office Math (§22.1) → fragments: runs become ordinary [`Fragment::Text`]s
//! in the math face, superscripts reuse the footnote-reference raise, and a
//! fraction becomes one pre-measured [`Fragment::MathFraction`] stack.

use std::rc::Rc;

use crate::model::{MathBlock, MathElement, DEFAULT_MATH_FONT};
use crate::render::dimension::Pt;
use crate::render::fonts::Toggle;

use super::text::{emit_text_words, TextRunStyle};
use super::{
    BreakAfter, FontProps, Fragment, FragmentCtx, LinkTarget, MathRow, TextMetrics,
    FRACTION_GAP_RATIO, FRACTION_RULE_RATIO, FRACTION_SIDE_PAD_RATIO, MATH_AXIS_RATIO,
    SUPERSCRIPT_ASCENT_OFFSET_RATIO, SUPERSCRIPT_FONT_SIZE_RATIO,
};

/// Emit one `m:oMath` into the paragraph's fragment stream.
///
/// `hyperlink_url` is the target of the enclosing `w:hyperlink`, if any — an
/// `m:oMath` shares `ParaChildXml`'s content model with ordinary runs, so it
/// can legally sit inside one, and every glyph the equation draws gets the
/// same annotation an ordinary run in the same hyperlink would.
pub(super) fn emit_math_fragments<F>(
    math: &MathBlock,
    ctx: &FragmentCtx<'_>,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let font = math_font(ctx.default_size);
    emit_elements(
        &math.content,
        &font,
        Pt::ZERO,
        ctx,
        hyperlink_url,
        measure_text,
        fragments,
    );
}

/// The math face at a given size. Word renders math in Cambria Math; the
/// italic look of variables comes from Unicode mathematical-alphabet
/// codepoints (see [`map_math_italic`]), not from an italic face — Cambria
/// Math has none.
fn math_font(size: Pt) -> FontProps {
    FontProps {
        strike_lines: 0,
        strike_position: Pt::ZERO,
        strike_thickness: Pt::ZERO,
        rtl: Toggle::Absent,
        family: Rc::from(DEFAULT_MATH_FONT),
        size,
        bold: Toggle::Absent,
        italic: Toggle::Absent,
        underline: false,
        char_spacing: Pt::ZERO,
        text_scale: 1.0,
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    }
}

fn emit_elements<F>(
    elements: &[MathElement],
    font: &FontProps,
    baseline_offset: Pt,
    ctx: &FragmentCtx<'_>,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    for element in elements {
        match element {
            MathElement::Run(run) => {
                let mapped = map_math_italic(&run.text);
                let style = TextRunStyle {
                    color: ctx.default_color,
                    shading: None,
                    border: None,
                    baseline_offset,
                };
                emit_text_words(
                    &mapped,
                    font,
                    &style,
                    hyperlink_url,
                    measure_text,
                    fragments,
                );
            }
            MathElement::Superscript { base, sup } => {
                let base_start = fragments.len();
                emit_elements(
                    base,
                    font,
                    baseline_offset,
                    ctx,
                    hyperlink_url,
                    measure_text,
                    fragments,
                );
                // The exponent belongs to its base: no line break between.
                // Guarded on `fragments.len() > base_start` because `base`
                // can legally emit nothing — §22.1 defaults `SSupXml::base`
                // to `vec![]`, and a text-less `m:r` is dropped before this
                // point — in which case there is no base fragment to glue
                // and `fragments.last_mut()` would otherwise reach back into
                // whatever unrelated fragment preceded this equation.
                //
                // Matches both `Fragment::Text` (an ordinary base) and
                // `Fragment::MathFraction` (a fraction as a base, e.g.
                // `(1/2)²`) — the only two variants `emit_elements` can leave
                // behind here. `line.rs`'s line-fitter has to honor
                // `MathFraction`'s own `break_after` for this to matter; see
                // its `is_break_point` match.
                if fragments.len() > base_start {
                    match fragments.last_mut().expect("checked non-empty above") {
                        Fragment::Text { break_after, .. }
                        | Fragment::MathFraction { break_after, .. } => {
                            *break_after = BreakAfter::Prohibited;
                        }
                        _ => {}
                    }
                }
                let (_, base_metrics) = measure_text("X", font);
                let mut sup_font = font.clone();
                sup_font.size = font.size * SUPERSCRIPT_FONT_SIZE_RATIO;
                let sup_offset =
                    baseline_offset - base_metrics.ascent * SUPERSCRIPT_ASCENT_OFFSET_RATIO;
                emit_elements(
                    sup,
                    &sup_font,
                    sup_offset,
                    ctx,
                    hyperlink_url,
                    measure_text,
                    fragments,
                );
            }
            MathElement::Fraction { num, den } => {
                fragments.push(fraction_fragment(
                    num,
                    den,
                    font,
                    baseline_offset,
                    ctx,
                    hyperlink_url,
                    measure_text,
                ));
            }
        }
    }
}

/// Build the pre-measured numerator/rule/denominator stack.
///
/// The ratios live next to the fragment definition in `fragment/mod.rs`;
/// both rows keep the full font size (`m:smallFrac` defaults to off —
/// display style).
fn fraction_fragment<F>(
    num: &[MathElement],
    den: &[MathElement],
    font: &FontProps,
    baseline_offset: Pt,
    ctx: &FragmentCtx<'_>,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
) -> Fragment
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let num_text = map_math_italic(&flatten_plain_text(num));
    let den_text = map_math_italic(&flatten_plain_text(den));
    let (num_width, num_metrics) = measure_text(&num_text, font);
    let (den_width, den_metrics) = measure_text(&den_text, font);

    let geometry = fraction_geometry(font.size, num_width, num_metrics, den_width, den_metrics);

    let row = |text: String, row_width: Pt, row_metrics: TextMetrics| MathRow {
        text: Rc::from(text.as_str()),
        font: Rc::new(font.clone()),
        width: row_width,
        metrics: row_metrics,
    };
    Fragment::MathFraction {
        num: row(num_text, num_width, num_metrics),
        den: row(den_text, den_width, den_metrics),
        color: ctx.default_color,
        width: geometry.width,
        metrics: geometry.metrics,
        baseline_offset,
        break_after: BreakAfter::Opportunity,
        hyperlink_url: hyperlink_url.cloned(),
    }
}

/// Every quantity derived from the `MATH_*`/`FRACTION_*` ratios that either
/// the initial layout pass or the paint pass needs to place a fraction
/// stack — the layout pass only wants `width`/`metrics`, the paint pass only
/// wants `rule`/`gap`/`pad`/`axis`, but both must read them off the same
/// call so the space line-fitting reserves and what paint actually draws
/// cannot independently drift (see [`fraction_geometry`]).
pub(crate) struct FractionGeometry {
    /// max(row widths) plus side padding — the fragment's line-fitting width.
    pub width: Pt,
    /// Synthesized ascent/descent covering both rows plus the rule and gaps.
    pub metrics: TextMetrics,
    /// Thickness of the fraction rule.
    pub rule: Pt,
    /// Vertical clearance between the rule and each row.
    pub gap: Pt,
    /// Horizontal padding on each side of the wider row.
    pub pad: Pt,
    /// Distance from the row's baseline up to the math axis the rule is
    /// centered on — paint derives the rule's y as `baseline - axis`.
    pub axis: Pt,
}

/// The fraction's geometry from its two rows' own width/metrics.
///
/// Shared between initial construction, [`super::fallback::apply_font_
/// fallback`]'s repair of a row whose face got substituted, and the paint
/// arm in `paragraph::line_emit` — all three must derive the stack's
/// geometry from the same `size`/`FRACTION_*` ratios, or the space
/// line-fitting reserves, what fallback repair leaves behind, and what paint
/// actually draws could disagree with each other.
pub(crate) fn fraction_geometry(
    size: Pt,
    num_width: Pt,
    num_metrics: TextMetrics,
    den_width: Pt,
    den_metrics: TextMetrics,
) -> FractionGeometry {
    let axis = size * MATH_AXIS_RATIO;
    let rule = size * FRACTION_RULE_RATIO;
    let gap = size * FRACTION_GAP_RATIO;
    let pad = size * FRACTION_SIDE_PAD_RATIO;

    let width = num_width.max(den_width) + pad * 2.0;
    let metrics = TextMetrics {
        ascent: axis + rule * 0.5 + gap + num_metrics.height(),
        descent: (rule * 0.5 + gap + den_metrics.height() - axis).max(Pt::ZERO),
        leading: Pt::ZERO,
    };
    FractionGeometry {
        width,
        metrics,
        rule,
        gap,
        pad,
        axis,
    }
}

/// Flatten a fraction argument to plain text. The minimal scope renders
/// nested structure linearly — a nested fraction becomes `num/den` — with a
/// warning, so the content is never silently lost.
fn flatten_plain_text(elements: &[MathElement]) -> String {
    let mut out = String::new();
    for element in elements {
        match element {
            MathElement::Run(run) => out.push_str(&run.text),
            MathElement::Superscript { base, sup } => {
                log::warn!("OMML: superscript inside a fraction argument renders linearly");
                out.push_str(&flatten_plain_text(base));
                out.push_str(&flatten_plain_text(sup));
            }
            MathElement::Fraction { num, den } => {
                log::warn!("OMML: nested fraction renders linearly as num/den");
                out.push_str(&flatten_plain_text(num));
                out.push('/');
                out.push_str(&flatten_plain_text(den));
            }
        }
    }
    out
}

/// What Word actually draws for math variables: ASCII letters mapped to the
/// Unicode Mathematical Italic alphabet (U+1D434…, U+1D44E…), with `h` on its
/// Letterlike exception U+210E ℎ. Digits and operators pass through upright,
/// matching §22.1's own default for them.
///
/// Every other character — non-ASCII letters included — also passes through
/// upright here, but that is this function's own limit, not §22.1's: Word
/// italicizes math variables in other alphabetic scripts (Greek, for
/// instance) by default too, and `m:sty` overrides are not consumed either,
/// so neither the real default nor an explicit override reaches a non-ASCII
/// letter yet. Extending coverage needs each script's own Mathematical
/// Italic sub-range, not a blanket pass-through.
fn map_math_italic(text: &str) -> String {
    text.chars()
        .map(|c| match c {
            'h' => '\u{210E}',
            'a'..='z' => {
                char::from_u32(0x1D44E + (c as u32 - 'a' as u32)).expect("math italic lowercase")
            }
            'A'..='Z' => {
                char::from_u32(0x1D434 + (c as u32 - 'A' as u32)).expect("math italic uppercase")
            }
            other => other,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::MathRun;

    fn dummy_measure(text: &str, font: &FontProps) -> (Pt, TextMetrics) {
        let width = font.size * 0.5 * text.chars().count() as f32;
        (
            width,
            TextMetrics {
                ascent: font.size * 0.8,
                descent: font.size * 0.2,
                leading: Pt::ZERO,
            },
        )
    }

    fn ctx() -> FragmentCtx<'static> {
        FragmentCtx {
            default_family: "Calibri",
            default_size: Pt::new(12.0),
            default_color: crate::render::resolve::color::RgbColor::BLACK,
            resolved_styles: None,
            paragraph_run_defaults: None,
            theme: None,
            measurer: None,
            auto_fit: crate::render::layout::ShapeAutoFit::NONE,
            locale_tag: None,
            revision_palette: None,
            comment_marks: false,
        }
    }

    fn run(text: &str) -> MathElement {
        MathElement::Run(MathRun { text: text.into() })
    }

    /// x² emits the base and a smaller, raised, unbreakable-from-base sup.
    #[test]
    fn superscript_raises_and_shrinks_the_exponent() {
        let math = MathBlock {
            content: vec![MathElement::Superscript {
                base: vec![run("x")],
                sup: vec![run("2")],
            }],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), None, &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 2, "base + exponent");
        match (&fragments[0], &fragments[1]) {
            (
                Fragment::Text {
                    font: base_font,
                    break_after,
                    baseline_offset: base_off,
                    ..
                },
                Fragment::Text {
                    font: sup_font,
                    baseline_offset: sup_off,
                    ..
                },
            ) => {
                assert_eq!(*break_after, BreakAfter::Prohibited, "base glued to sup");
                assert!(
                    (sup_font.size.raw() - base_font.size.raw() * SUPERSCRIPT_FONT_SIZE_RATIO)
                        .abs()
                        < 1e-4
                );
                assert!(sup_off.raw() < base_off.raw(), "exponent raised");
            }
            other => panic!("expected two text fragments, got {other:?}"),
        }
    }

    /// A fraction can be a superscript's base (`(1/2)²` — `m:sSup` whose
    /// `m:e` is `m:f`) — the exponent must still glue to it. `line.rs`'s
    /// `is_break_point` has to honor `Fragment::MathFraction`'s own
    /// `break_after` for this to matter; that half is covered in
    /// `render::layout::line`'s own tests.
    #[test]
    fn superscript_glues_to_a_fraction_base() {
        let math = MathBlock {
            content: vec![MathElement::Superscript {
                base: vec![MathElement::Fraction {
                    num: vec![run("1")],
                    den: vec![run("2")],
                }],
                sup: vec![run("2")],
            }],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), None, &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 2, "fraction base + exponent");
        match &fragments[0] {
            Fragment::MathFraction { break_after, .. } => {
                assert_eq!(
                    *break_after,
                    BreakAfter::Prohibited,
                    "fraction base glued to sup"
                );
            }
            other => panic!("expected the fraction base, got {other:?}"),
        }
    }

    /// §22.1 lets `SSupXml::base` be empty — dropped by `convert_children`
    /// when its only content is a text-less `m:r`, and reachable directly
    /// too. The glue must not then reach backward past the equation into
    /// whatever fragment preceded it in the same paragraph.
    #[test]
    fn superscript_with_an_empty_base_does_not_corrupt_a_preceding_fragment() {
        let math = MathBlock {
            content: vec![MathElement::Superscript {
                base: vec![],
                sup: vec![run("2")],
            }],
        };
        let mut fragments = Vec::new();
        // Ordinary prose preceding the equation in the same paragraph, with
        // an ordinary break opportunity after it.
        let style = TextRunStyle {
            color: crate::render::resolve::color::RgbColor::BLACK,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
        };
        emit_text_words(
            "word",
            &math_font(Pt::new(12.0)),
            &style,
            None,
            &dummy_measure,
            &mut fragments,
        );
        let before = fragments.len();

        emit_math_fragments(&math, &ctx(), None, &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), before + 1, "just the bare exponent");
        match &fragments[0] {
            Fragment::Text { break_after, .. } => {
                assert_eq!(
                    *break_after,
                    BreakAfter::Opportunity,
                    "an empty superscript base must not glue backward into unrelated prose"
                );
            }
            other => panic!("expected the preceding word untouched, got {other:?}"),
        }
    }

    /// A fraction is one stacked fragment whose synthesized metrics cover
    /// both rows plus the rule and gaps.
    #[test]
    fn fraction_stacks_into_one_fragment() {
        let math = MathBlock {
            content: vec![MathElement::Fraction {
                num: vec![run("1")],
                den: vec![run("12")],
            }],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), None, &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 1);
        match &fragments[0] {
            Fragment::MathFraction {
                num,
                den,
                width,
                metrics,
                ..
            } => {
                assert_eq!(&*num.text, "1");
                assert_eq!(&*den.text, "12");
                assert!(
                    *width > num.width.max(den.width),
                    "side padding widens the stack"
                );
                assert!(
                    metrics.ascent > num.metrics.height(),
                    "numerator sits fully above the baseline area"
                );
                assert!(metrics.descent > Pt::ZERO);
            }
            other => panic!("expected a fraction, got {other:?}"),
        }
    }

    /// Variables italicize via the Unicode math alphabet; digits and
    /// operators stay upright. `h` takes its Letterlike exception.
    #[test]
    fn math_italic_maps_letters_only() {
        assert_eq!(map_math_italic("x2"), "\u{1D465}2");
        assert_eq!(map_math_italic("h"), "\u{210E}");
        assert_eq!(map_math_italic("A + 1"), "\u{1D434} + 1");
    }

    /// A known gap, pinned rather than left to be assumed away: Word
    /// italicizes non-ASCII math variables by default too, but this
    /// function only maps the ASCII alphabet, so a Greek variable passes
    /// through upright. If this ever starts failing because someone taught
    /// `map_math_italic` a non-ASCII range, update this test and the
    /// function's doc comment together — don't just delete the assertion.
    #[test]
    fn math_italic_leaves_non_ascii_letters_upright() {
        assert_eq!(map_math_italic("α"), "α");
    }

    /// An `m:oMath` sharing `ParaChildXml`'s content model with ordinary runs
    /// means it can sit inside a `w:hyperlink` — every glyph the run
    /// produces must carry the same annotation an ordinary run in the same
    /// hyperlink would.
    #[test]
    fn a_math_run_inside_a_hyperlink_carries_its_target() {
        let target = LinkTarget::External(Rc::from("https://example.invalid"));
        let math = MathBlock {
            content: vec![run("x")],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), Some(&target), &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 1);
        match &fragments[0] {
            Fragment::Text { hyperlink_url, .. } => {
                assert_eq!(hyperlink_url.as_ref(), Some(&target));
            }
            other => panic!("expected text, got {other:?}"),
        }
    }

    /// A superscript's base and exponent are two separate `Fragment::Text`s
    /// (see `superscript_raises_and_shrinks_the_exponent`) — both must carry
    /// the hyperlink, not just the base.
    #[test]
    fn a_math_superscript_inside_a_hyperlink_carries_its_target_on_both_parts() {
        let target = LinkTarget::External(Rc::from("https://example.invalid"));
        let math = MathBlock {
            content: vec![MathElement::Superscript {
                base: vec![run("x")],
                sup: vec![run("2")],
            }],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), Some(&target), &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 2);
        for fragment in &fragments {
            match fragment {
                Fragment::Text { hyperlink_url, .. } => {
                    assert_eq!(hyperlink_url.as_ref(), Some(&target));
                }
                other => panic!("expected text, got {other:?}"),
            }
        }
    }

    /// A fraction is one pre-measured atom (`Fragment::MathFraction`, not
    /// `Fragment::Text`) — it needs its own `hyperlink_url` field, carried
    /// the same way.
    #[test]
    fn a_math_fraction_inside_a_hyperlink_carries_its_target() {
        let target = LinkTarget::External(Rc::from("https://example.invalid"));
        let math = MathBlock {
            content: vec![MathElement::Fraction {
                num: vec![run("1")],
                den: vec![run("2")],
            }],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), Some(&target), &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 1);
        match &fragments[0] {
            Fragment::MathFraction { hyperlink_url, .. } => {
                assert_eq!(hyperlink_url.as_ref(), Some(&target));
            }
            other => panic!("expected a fraction, got {other:?}"),
        }
    }

    /// The stated behaviour when there is no hyperlink: unchanged from before
    /// this field existed.
    #[test]
    fn math_without_a_hyperlink_carries_no_target() {
        let math = MathBlock {
            content: vec![
                run("x"),
                MathElement::Fraction {
                    num: vec![run("1")],
                    den: vec![run("2")],
                },
            ],
        };
        let mut fragments = Vec::new();
        emit_math_fragments(&math, &ctx(), None, &dummy_measure, &mut fragments);

        assert_eq!(fragments.len(), 2);
        for fragment in &fragments {
            match fragment {
                Fragment::Text { hyperlink_url, .. } => assert_eq!(*hyperlink_url, None),
                Fragment::MathFraction { hyperlink_url, .. } => assert_eq!(*hyperlink_url, None),
                other => panic!("expected text or a fraction, got {other:?}"),
            }
        }
    }
}
