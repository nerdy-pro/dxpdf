//! §20.1.9.18 `roundRect` — a rectangle with four quarter-ellipse corners.
//!
//! Spec guides: `adj` defaults to 16667; the corner radius is
//! `ss · pin(0, adj, 50000) / 100000` with `ss = min(w, h)`. The text
//! rectangle is inset by `x1 · 29289/100000` per side — the sagitta of the
//! 45° point of the corner arc, so text clears the rounded corner.

use crate::model::{PathFillMode, PresetGeometryDef};
use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::resolve::shape_geometry::{PathVerb, ShapePath, SubPath};

use super::common::{adjust, corner_arc, pin, pt};

pub fn build(def: &PresetGeometryDef, extent: PtSize) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    let ss = w.min(h);
    let a = pin(0.0, adjust(def, "adj", 16667.0), 50000.0);
    let rad = ss * a / 100000.0;

    let mut verbs = vec![PathVerb::MoveTo(pt(rad, 0.0))];
    verbs.push(PathVerb::LineTo(pt(w - rad, 0.0)));
    corner_arc(&mut verbs, pt(w - rad, 0.0), pt(w, 0.0), pt(w, rad));
    verbs.push(PathVerb::LineTo(pt(w, h - rad)));
    corner_arc(&mut verbs, pt(w, h - rad), pt(w, h), pt(w - rad, h));
    verbs.push(PathVerb::LineTo(pt(rad, h)));
    corner_arc(&mut verbs, pt(rad, h), pt(0.0, h), pt(0.0, h - rad));
    verbs.push(PathVerb::LineTo(pt(0.0, rad)));
    corner_arc(&mut verbs, pt(0.0, rad), pt(0.0, 0.0), pt(rad, 0.0));
    verbs.push(PathVerb::Close);

    let inset = rad * 0.29289;
    ShapePath {
        paths: vec![SubPath {
            verbs,
            fill_mode: PathFillMode::Norm,
            stroked: true,
        }],
        text_rect: Some(PtRect::from_xywh(
            Pt::new(inset),
            Pt::new(inset),
            Pt::new((w - 2.0 * inset).max(0.0)),
            Pt::new((h - 2.0 * inset).max(0.0)),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::GeomGuide;
    use crate::model::PresetShapeType;

    fn def(adj: Option<i64>) -> PresetGeometryDef {
        PresetGeometryDef {
            preset: PresetShapeType::RoundRect,
            adjust_values: adj
                .map(|v| {
                    vec![GeomGuide {
                        name: "adj".into(),
                        formula: format!("val {v}"),
                    }]
                })
                .unwrap_or_default(),
        }
    }

    fn first_move(p: &ShapePath) -> (f32, f32) {
        let PathVerb::MoveTo(o) = p.paths[0].verbs[0] else {
            panic!()
        };
        (o.x.raw(), o.y.raw())
    }

    /// Default adj 16667 on a 60×30 box: radius = 30 · 0.16667 ≈ 5.
    #[test]
    fn default_radius_is_a_sixth_of_the_short_side() {
        let p = build(&def(None), PtSize::new(Pt::new(60.0), Pt::new(30.0)));
        let (x, y) = first_move(&p);
        assert!((x - 5.0).abs() < 0.01, "start x = radius, got {x}");
        assert_eq!(y, 0.0);
    }

    /// adj 50000 (the clamp maximum) makes the radius half the short side —
    /// a stadium shape on a wide box.
    #[test]
    fn adjust_scales_and_clamps_the_radius() {
        let p = build(&def(Some(50000)), PtSize::new(Pt::new(60.0), Pt::new(30.0)));
        assert_eq!(first_move(&p).0, 15.0);
        let p = build(&def(Some(99999)), PtSize::new(Pt::new(60.0), Pt::new(30.0)));
        assert_eq!(first_move(&p).0, 15.0, "over-max adj clamps to 50000");
        let p = build(&def(Some(0)), PtSize::new(Pt::new(60.0), Pt::new(30.0)));
        assert_eq!(first_move(&p).0, 0.0, "adj 0 degenerates to a rect");
    }

    /// Four corners, each one cubic: 1 move + 4 lines + 4 cubics + close.
    #[test]
    fn path_has_four_rounded_corners() {
        let p = build(&def(None), PtSize::new(Pt::new(60.0), Pt::new(30.0)));
        let cubics = p.paths[0]
            .verbs
            .iter()
            .filter(|v| matches!(v, PathVerb::CubicTo(..)))
            .count();
        assert_eq!(cubics, 4);
        assert!(matches!(p.paths[0].verbs.last(), Some(PathVerb::Close)));
    }

    /// §20.1.9.18: the text box clears the corner arc by the 45° sagitta.
    #[test]
    fn text_rect_is_inset_by_the_corner_sagitta() {
        let p = build(&def(Some(50000)), PtSize::new(Pt::new(60.0), Pt::new(30.0)));
        let tr = p.text_rect.unwrap();
        // radius 15 → inset 15 · 0.29289 ≈ 4.393
        assert!((tr.origin.x.raw() - 4.393).abs() < 0.01);
        assert!((tr.origin.y.raw() - 4.393).abs() < 0.01);
    }
}
