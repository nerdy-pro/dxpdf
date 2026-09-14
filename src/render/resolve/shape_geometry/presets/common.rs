//! Shared arithmetic for the preset generators.
//!
//! Every §20.1.9.18 preset is defined in the spec's `presetShapeDefinitions`
//! as guide formulas over the same handful of inputs — the extent, `ss`
//! (the shorter side), and the shape's adjust values — so the generators
//! share the two pieces every formula needs: reading an `avLst` override,
//! and drawing a quarter-ellipse corner. Corners are cubic Béziers rather
//! than [`PathVerb::ArcTo`]: the constant below reproduces a circular arc to
//! within 0.03% of the radius, and a cubic's endpoints are explicit — the
//! generator states exactly where the corner starts and ends instead of
//! deriving it from the painter's arc convention.

use crate::model::PresetGeometryDef;
use crate::render::dimension::Pt;
use crate::render::geometry::PtOffset;
use crate::render::resolve::shape_geometry::PathVerb;

/// The cubic-Bézier circle constant: control-point distance as a fraction of
/// the radius for a 90° arc, `4/3·(√2 − 1)`.
pub(super) const KAPPA: f32 = 0.552_284_8;

/// §20.1.9.5: read an adjust value from the shape's `avLst`, falling back to
/// the preset's spec default. Adjust formulas are always `val N`; anything
/// else (a guide reference, which `avLst` cannot legally hold) reads as the
/// default.
pub(super) fn adjust(def: &PresetGeometryDef, name: &str, default: f32) -> f32 {
    def.adjust_values
        .iter()
        .find(|g| g.name == name)
        .and_then(|g| g.formula.strip_prefix("val "))
        .and_then(|v| v.trim().parse::<f32>().ok())
        .unwrap_or(default)
}

/// §20.1.9.18's ubiquitous `pin` guide: clamp `v` into `[lo, hi]`.
pub(super) fn pin(lo: f32, v: f32, hi: f32) -> f32 {
    v.clamp(lo, hi)
}

/// A 90° elliptical corner from the current pen position to `end`, bulging
/// toward `corner` (the box corner the arc rounds). The pen must sit at the
/// arc's start.
pub(super) fn corner_arc(
    verbs: &mut Vec<PathVerb>,
    from: PtOffset,
    corner: PtOffset,
    end: PtOffset,
) {
    let c1 = PtOffset::new(
        from.x + (corner.x - from.x) * KAPPA,
        from.y + (corner.y - from.y) * KAPPA,
    );
    let c2 = PtOffset::new(
        end.x + (corner.x - end.x) * KAPPA,
        end.y + (corner.y - end.y) * KAPPA,
    );
    verbs.push(PathVerb::CubicTo(c1, c2, end));
}

/// Convenience: `PtOffset` from raw f32 points.
pub(super) fn pt(x: f32, y: f32) -> PtOffset {
    PtOffset::new(Pt::new(x), Pt::new(y))
}
