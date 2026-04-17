//! Geometry guide expression evaluator (§20.1.9.11 CT_GeomGuide).
//!
//! OOXML guide formulas are postfix-like prefix expressions. Each formula is
//! a single operator followed by one to three operand tokens, where each
//! operand is either a decimal literal, a guide-name reference, or one of a
//! small set of spec-defined named constants.
//!
//! The spec enumerates 18 operators (§20.1.9.11 ST_GeomGuideFormula). All
//! are implemented here. Resolution is pure: given a `GuideContext` and a
//! slice of `GeomGuide`s in document order, `evaluate_guides` returns a
//! `GuideValues` map keyed by guide name.
//!
//! Units: formula operands and results are f64. Angles are expressed in
//! 60000ths of a degree (§20.1.10.3 ST_PositiveFixedAngle convention);
//! dimensions are in the path's local EMU coordinate space.

use std::collections::HashMap;
use std::f64::consts::PI;

use crate::model::{AdjAngle, AdjCoord, GeomGuide};

/// The coordinate context a guide expression is evaluated against.
///
/// Fields correspond to the spec's named constants:
///  * `w`, `h`      — path extent.
///  * `ss`          — shortest side (`min(w, h)`).
///  * `ls`          — longest side (`max(w, h)`).
///  * `hc`, `vc`    — horizontal/vertical center.
///  * `t`, `b`, `l`, `r` — top/bottom/left/right edges.
///  * `wd2`, `wd4`, `wd8`, `wd16`, `wd32` — width divisions.
///  * `hd2`, `hd4`, `hd8`, `hd16`, `hd32` — height divisions.
///  * `cd2`, `cd4`, `cd8` — 360° / N angle divisions (in 60000ths of a deg).
///  * `3cd4`, `3cd8`, `5cd8`, `7cd8` — compound angle constants.
#[derive(Clone, Copy, Debug)]
pub struct GuideContext {
    pub w: f64,
    pub h: f64,
}

impl GuideContext {
    pub fn new(w: f64, h: f64) -> Self {
        Self { w, h }
    }
}

/// A name → evaluated value map produced by [`evaluate_guides`].
pub type GuideValues = HashMap<String, f64>;

/// Evaluate a list of guides in document order. Later guides may reference
/// earlier guide names via [`AdjCoord::Guide`]. Returns a map of resolved
/// values; callers pass this to [`resolve_adj_coord`] / [`resolve_adj_angle`]
/// when reading path commands.
///
/// Unknown guide references resolve to 0.0 (spec says behaviour is undefined
/// for forward references; we default to 0 so we never panic on malformed
/// but structurally-valid input).
pub fn evaluate_guides(guides: &[GeomGuide], ctx: GuideContext) -> GuideValues {
    let mut values = GuideValues::new();
    for guide in guides {
        if let Some(v) = evaluate_formula(&guide.formula, &values, ctx) {
            values.insert(guide.name.clone(), v);
        }
    }
    values
}

/// Resolve an `AdjCoord` (literal or guide reference) to a concrete value.
pub fn resolve_adj_coord(c: &AdjCoord, values: &GuideValues, ctx: GuideContext) -> f64 {
    match c {
        AdjCoord::Lit(n) => *n as f64,
        AdjCoord::Guide(name) => values
            .get(name.as_str())
            .copied()
            .or_else(|| named_constant(name, ctx))
            .unwrap_or(0.0),
    }
}

/// `AdjAngle` is `AdjCoord` per §20.1.10.4 — same resolution.
pub fn resolve_adj_angle(a: &AdjAngle, values: &GuideValues, ctx: GuideContext) -> f64 {
    resolve_adj_coord(a, values, ctx)
}

/// Parse + evaluate a single formula string.
///
/// Returns `None` if the formula cannot be parsed (malformed) or if an
/// operator's arity is wrong. The caller leaves the guide unassigned, which
/// surfaces as 0.0 downstream.
fn evaluate_formula(formula: &str, values: &GuideValues, ctx: GuideContext) -> Option<f64> {
    let mut tokens = formula.split_whitespace();
    let op = tokens.next()?;
    let a = tokens.next();
    let b = tokens.next();
    let c = tokens.next();
    let v = |t: &str| -> f64 { resolve_token(t, values, ctx) };

    Some(match op {
        // §20.1.9.11 three-operand operators
        "*/" => {
            let (a, b, c) = (v(a?), v(b?), v(c?));
            if c == 0.0 {
                0.0
            } else {
                (a * b) / c
            }
        }
        "+-" => v(a?) + v(b?) - v(c?),
        "+/" => {
            let (a, b, c) = (v(a?), v(b?), v(c?));
            if c == 0.0 {
                0.0
            } else {
                (a + b) / c
            }
        }
        "?:" => {
            // If a > 0 then b else c.
            if v(a?) > 0.0 {
                v(b?)
            } else {
                v(c?)
            }
        }
        "cat2" => {
            // val1 * cos(arctan2(val3, val2))
            let (a, b, c) = (v(a?), v(b?), v(c?));
            a * (c.atan2(b)).cos()
        }
        "mod" => {
            // sqrt(val1² + val2² + val3²) — despite the name, this is
            // vector magnitude per the spec.
            let (a, b, c) = (v(a?), v(b?), v(c?));
            (a * a + b * b + c * c).sqrt()
        }
        "pin" => {
            // clamp(val2, val1, val3): lower=val1, value=val2, upper=val3.
            let (a, b, c) = (v(a?), v(b?), v(c?));
            b.max(a).min(c)
        }
        "sat2" => {
            // val1 * sin(arctan2(val3, val2))
            let (a, b, c) = (v(a?), v(b?), v(c?));
            a * (c.atan2(b)).sin()
        }

        // Two-operand operators
        "at2" => {
            // arctan2(val2, val1), result in 60000ths of a degree.
            let (a, b) = (v(a?), v(b?));
            b.atan2(a) * 180.0 / PI * 60_000.0
        }
        "cos" => {
            // val1 * cos(val2); val2 in 60000ths of a degree.
            let (a, b) = (v(a?), v(b?));
            a * (b * PI / 180.0 / 60_000.0).cos()
        }
        "max" => v(a?).max(v(b?)),
        "min" => v(a?).min(v(b?)),
        "sin" => {
            let (a, b) = (v(a?), v(b?));
            a * (b * PI / 180.0 / 60_000.0).sin()
        }
        "tan" => {
            let (a, b) = (v(a?), v(b?));
            a * (b * PI / 180.0 / 60_000.0).tan()
        }

        // One-operand operators
        "abs" => v(a?).abs(),
        "sqrt" => v(a?).sqrt(),
        "val" => v(a?),

        _ => return None,
    })
}

fn resolve_token(token: &str, values: &GuideValues, ctx: GuideContext) -> f64 {
    if let Ok(n) = token.parse::<i64>() {
        return n as f64;
    }
    if let Ok(n) = token.parse::<f64>() {
        return n;
    }
    if let Some(v) = named_constant(token, ctx) {
        return v;
    }
    values.get(token).copied().unwrap_or(0.0)
}

/// §20.1.9.11 spec-named constants available inside any formula.
fn named_constant(name: &str, ctx: GuideContext) -> Option<f64> {
    Some(match name {
        "w" => ctx.w,
        "h" => ctx.h,
        "ss" => ctx.w.min(ctx.h),
        "ls" => ctx.w.max(ctx.h),
        "hc" => ctx.w / 2.0,
        "vc" => ctx.h / 2.0,
        "t" | "l" => 0.0,
        "b" => ctx.h,
        "r" => ctx.w,
        "wd2" => ctx.w / 2.0,
        "wd4" => ctx.w / 4.0,
        "wd8" => ctx.w / 8.0,
        "wd16" => ctx.w / 16.0,
        "wd32" => ctx.w / 32.0,
        "hd2" => ctx.h / 2.0,
        "hd4" => ctx.h / 4.0,
        "hd8" => ctx.h / 8.0,
        "hd16" => ctx.h / 16.0,
        "hd32" => ctx.h / 32.0,
        // Angle constants — 360° expressed in 60000ths of a degree.
        "cd2" => 360.0 / 2.0 * 60_000.0,
        "cd4" => 360.0 / 4.0 * 60_000.0,
        "cd8" => 360.0 / 8.0 * 60_000.0,
        "3cd4" => 3.0 * 360.0 / 4.0 * 60_000.0,
        "3cd8" => 3.0 * 360.0 / 8.0 * 60_000.0,
        "5cd8" => 5.0 * 360.0 / 8.0 * 60_000.0,
        "7cd8" => 7.0 * 360.0 / 8.0 * 60_000.0,
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx() -> GuideContext {
        GuideContext::new(100.0, 50.0)
    }

    fn g(name: &str, formula: &str) -> GeomGuide {
        GeomGuide {
            name: name.into(),
            formula: formula.into(),
        }
    }

    #[test]
    fn val_literal() {
        let vs = evaluate_guides(&[g("a", "val 42")], ctx());
        assert_eq!(vs["a"], 42.0);
    }

    #[test]
    fn val_named_constant_w() {
        let vs = evaluate_guides(&[g("x", "val w")], ctx());
        assert_eq!(vs["x"], 100.0);
    }

    #[test]
    fn val_named_constant_hc_vc() {
        let vs = evaluate_guides(&[g("cx", "val hc"), g("cy", "val vc")], ctx());
        assert_eq!(vs["cx"], 50.0);
        assert_eq!(vs["cy"], 25.0);
    }

    #[test]
    fn val_ss_ls() {
        let vs = evaluate_guides(&[g("s", "val ss"), g("l", "val ls")], ctx());
        assert_eq!(vs["s"], 50.0);
        assert_eq!(vs["l"], 100.0);
    }

    #[test]
    fn muldiv() {
        // (10 * 3) / 2 = 15
        let vs = evaluate_guides(&[g("r", "*/ 10 3 2")], ctx());
        assert_eq!(vs["r"], 15.0);
    }

    #[test]
    fn muldiv_by_zero_returns_zero() {
        let vs = evaluate_guides(&[g("r", "*/ 10 3 0")], ctx());
        assert_eq!(vs["r"], 0.0);
    }

    #[test]
    fn plus_minus() {
        // 10 + 5 - 3 = 12
        let vs = evaluate_guides(&[g("r", "+- 10 5 3")], ctx());
        assert_eq!(vs["r"], 12.0);
    }

    #[test]
    fn plus_div() {
        // (10 + 6) / 4 = 4
        let vs = evaluate_guides(&[g("r", "+/ 10 6 4")], ctx());
        assert_eq!(vs["r"], 4.0);
    }

    #[test]
    fn ternary_selects_by_sign() {
        // if 1 > 0 then 42 else 99
        let vs = evaluate_guides(&[g("a", "?: 1 42 99"), g("b", "?: -1 42 99")], ctx());
        assert_eq!(vs["a"], 42.0);
        assert_eq!(vs["b"], 99.0);
    }

    #[test]
    fn pin_clamp() {
        // clamp(5, 0, 10) = 5; clamp(-1, 0, 10) = 0; clamp(20, 0, 10) = 10.
        let vs = evaluate_guides(
            &[
                g("a", "pin 0 5 10"),
                g("b", "pin 0 -1 10"),
                g("c", "pin 0 20 10"),
            ],
            ctx(),
        );
        assert_eq!(vs["a"], 5.0);
        assert_eq!(vs["b"], 0.0);
        assert_eq!(vs["c"], 10.0);
    }

    #[test]
    fn min_max() {
        let vs = evaluate_guides(&[g("mn", "min 3 7"), g("mx", "max 3 7")], ctx());
        assert_eq!(vs["mn"], 3.0);
        assert_eq!(vs["mx"], 7.0);
    }

    #[test]
    fn abs_and_sqrt() {
        let vs = evaluate_guides(&[g("a", "abs -9"), g("s", "sqrt 16")], ctx());
        assert_eq!(vs["a"], 9.0);
        assert_eq!(vs["s"], 4.0);
    }

    #[test]
    fn mod_is_vector_magnitude() {
        // sqrt(3² + 4² + 12²) = sqrt(169) = 13
        let vs = evaluate_guides(&[g("m", "mod 3 4 12")], ctx());
        assert!((vs["m"] - 13.0).abs() < 1e-9);
    }

    #[test]
    fn trig_cos_sin_zero_angle() {
        // cos(0) = 1, sin(0) = 0. val1 = 100 scales.
        let vs = evaluate_guides(&[g("c", "cos 100 0"), g("s", "sin 100 0")], ctx());
        assert!((vs["c"] - 100.0).abs() < 1e-9);
        assert!(vs["s"].abs() < 1e-9);
    }

    #[test]
    fn trig_cos_90deg() {
        // cos(90°) ≈ 0, sin(90°) ≈ 1. 90° = 5_400_000 in 60000ths of a deg.
        let vs = evaluate_guides(
            &[g("c", "cos 100 5400000"), g("s", "sin 100 5400000")],
            ctx(),
        );
        assert!(vs["c"].abs() < 1e-6);
        assert!((vs["s"] - 100.0).abs() < 1e-6);
    }

    #[test]
    fn at2_returns_angle_in_60k_deg() {
        // arctan2(100, 0) = 90° = 5_400_000.
        let vs = evaluate_guides(&[g("a", "at2 0 100")], ctx());
        assert!((vs["a"] - 5_400_000.0).abs() < 1e-3);
    }

    #[test]
    fn cat2_sat2_polar_projection() {
        // cat2(r, 0, r) = r * cos(arctan2(r, 0)) = r * cos(90°) = 0
        // sat2(r, 0, r) = r * sin(arctan2(r, 0)) = r * sin(90°) = r
        let vs = evaluate_guides(
            &[g("cx", "cat2 100 0 100"), g("cy", "sat2 100 0 100")],
            ctx(),
        );
        assert!(vs["cx"].abs() < 1e-6);
        assert!((vs["cy"] - 100.0).abs() < 1e-6);
    }

    #[test]
    fn guide_references_preceding_guide() {
        let vs = evaluate_guides(
            &[g("half_w", "*/ w 1 2"), g("quarter_w", "*/ half_w 1 2")],
            ctx(),
        );
        assert_eq!(vs["half_w"], 50.0);
        assert_eq!(vs["quarter_w"], 25.0);
    }

    #[test]
    fn forward_reference_resolves_to_zero() {
        // First guide references second (not yet computed) → 0.
        let vs = evaluate_guides(&[g("first", "val later"), g("later", "val 10")], ctx());
        assert_eq!(vs["first"], 0.0);
        assert_eq!(vs["later"], 10.0);
    }

    #[test]
    fn malformed_formula_is_skipped() {
        let vs = evaluate_guides(&[g("a", "bogus_op 1 2")], ctx());
        assert!(!vs.contains_key("a"));
    }

    #[test]
    fn resolve_adj_coord_lit_vs_guide() {
        let mut values = GuideValues::new();
        values.insert("x".into(), 42.0);
        assert_eq!(resolve_adj_coord(&AdjCoord::Lit(7), &values, ctx()), 7.0);
        assert_eq!(
            resolve_adj_coord(&AdjCoord::Guide("x".into()), &values, ctx()),
            42.0
        );
        // Named constant fallback.
        assert_eq!(
            resolve_adj_coord(&AdjCoord::Guide("w".into()), &values, ctx()),
            100.0
        );
        // Unknown → 0.
        assert_eq!(
            resolve_adj_coord(&AdjCoord::Guide("nope".into()), &values, ctx()),
            0.0
        );
    }

    #[test]
    fn angle_constants_cd4_is_90deg() {
        let vs = evaluate_guides(&[g("a", "val cd4")], ctx());
        assert_eq!(vs["a"], 5_400_000.0);
    }
}
