//! Office Math (§22.1) → fragments: runs become ordinary [`Fragment::Text`]s
//! in the math face, superscripts reuse the footnote-reference raise, and a
//! fraction becomes one pre-measured [`Fragment::MathFraction`] stack.

use std::rc::Rc;

use crate::model::{MathBlock, MathElement, DEFAULT_MATH_FONT};
use crate::render::dimension::Pt;
use crate::render::fonts::Toggle;

use super::text::{emit_text_words, TextRunStyle};
use super::{
    BreakAfter, FontProps, Fragment, FragmentCtx, MathRow, TextMetrics, FRACTION_GAP_RATIO,
    FRACTION_RULE_RATIO, FRACTION_SIDE_PAD_RATIO, MATH_AXIS_RATIO, SUPERSCRIPT_ASCENT_OFFSET_RATIO,
    SUPERSCRIPT_FONT_SIZE_RATIO,
};

/// Emit one `m:oMath` into the paragraph's fragment stream.
pub(super) fn emit_math_fragments<F>(
    math: &MathBlock,
    ctx: &FragmentCtx<'_>,
    measure_text: &F,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let font = math_font(ctx.default_size);
    emit_elements(&math.content, &font, Pt::ZERO, ctx, measure_text, fragments);
}

/// The math face at a given size. Word renders math in Cambria Math; the
/// italic look of variables comes from Unicode mathematical-alphabet
/// codepoints (see [`map_math_italic`]), not from an italic face — Cambria
/// Math has none.
fn math_font(size: Pt) -> FontProps {
    FontProps {
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
                emit_text_words(&mapped, font, &style, None, measure_text, fragments);
            }
            MathElement::Superscript { base, sup } => {
                emit_elements(base, font, baseline_offset, ctx, measure_text, fragments);
                // The exponent belongs to its base: no line break between.
                if let Some(Fragment::Text { break_after, .. }) = fragments.last_mut() {
                    *break_after = BreakAfter::Prohibited;
                }
                let (_, base_metrics) = measure_text("X", font);
                let mut sup_font = font.clone();
                sup_font.size = font.size * SUPERSCRIPT_FONT_SIZE_RATIO;
                let sup_offset =
                    baseline_offset - base_metrics.ascent * SUPERSCRIPT_ASCENT_OFFSET_RATIO;
                emit_elements(sup, &sup_font, sup_offset, ctx, measure_text, fragments);
            }
            MathElement::Fraction { num, den } => {
                fragments.push(fraction_fragment(
                    num,
                    den,
                    font,
                    baseline_offset,
                    ctx,
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
    measure_text: &F,
) -> Fragment
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let num_text = map_math_italic(&flatten_plain_text(num));
    let den_text = map_math_italic(&flatten_plain_text(den));
    let (num_width, num_metrics) = measure_text(&num_text, font);
    let (den_width, den_metrics) = measure_text(&den_text, font);

    let size = font.size;
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
        width,
        metrics,
        baseline_offset,
        break_after: BreakAfter::Opportunity,
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
/// Letterlike exception U+210E ℎ. Digits, operators and everything else pass
/// through upright, per §22.1 defaults (`m:sty` overrides are not consumed).
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
        emit_math_fragments(&math, &ctx(), &dummy_measure, &mut fragments);

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
        emit_math_fragments(&math, &ctx(), &dummy_measure, &mut fragments);

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
}
