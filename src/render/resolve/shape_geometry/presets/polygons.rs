//! §20.1.9.18 straight-edged presets — `chevron`, `homePlate`, `diamond`,
//! `triangle` (the spec's `isocTriangle` spelling is a separate value; both
//! land here).
//!
//! Each is its spec definition's point list evaluated directly; the guide
//! formulas are quoted at each builder. `ss = min(w, h)` throughout.

use crate::model::{PathFillMode, PresetGeometryDef};
use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::resolve::shape_geometry::{PathVerb, ShapePath, SubPath};

use super::common::{adjust, pin, pt};

fn polygon(points: &[(f32, f32)], text_rect: Option<PtRect>) -> ShapePath {
    let mut verbs = Vec::with_capacity(points.len() + 1);
    for (i, &(x, y)) in points.iter().enumerate() {
        verbs.push(if i == 0 {
            PathVerb::MoveTo(pt(x, y))
        } else {
            PathVerb::LineTo(pt(x, y))
        });
    }
    verbs.push(PathVerb::Close);
    ShapePath {
        paths: vec![SubPath {
            verbs,
            fill_mode: PathFillMode::Norm,
            stroked: true,
        }],
        text_rect,
    }
}

fn rect(x0: f32, y0: f32, x1: f32, y1: f32) -> Option<PtRect> {
    Some(PtRect::from_xywh(
        Pt::new(x0),
        Pt::new(y0),
        Pt::new((x1 - x0).max(0.0)),
        Pt::new((y1 - y0).max(0.0)),
    ))
}

/// `chevron`: `adj` default 50000, `maxAdj = 100000·w/ss`,
/// `x1 = ss·pin(0, adj, maxAdj)/100000`, `x2 = r − x1`; points
/// `(l,t) (x2,t) (r,vc) (x2,b) (l,b) (x1,vc)`.
pub fn chevron(def: &PresetGeometryDef, extent: PtSize) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    let ss = w.min(h).max(f32::EPSILON);
    let a = pin(0.0, adjust(def, "adj", 50000.0), 100000.0 * w / ss);
    let x1 = ss * a / 100000.0;
    let x2 = w - x1;
    let vc = h / 2.0;
    polygon(
        &[(0.0, 0.0), (x2, 0.0), (w, vc), (x2, h), (0.0, h), (x1, vc)],
        // The spec's il/ir guides place text between the two notches — and
        // guard them: `?: dx x1 l` falls back to the full bounding box the
        // moment the notches meet or cross (any chevron with w ≤ h).
        if x2 > x1 {
            rect(x1, 0.0, x2, h)
        } else {
            rect(0.0, 0.0, w, h)
        },
    )
}

/// `homePlate`: same guides as chevron but the tail edge is flat —
/// points `(l,t) (x1,t) (r,vc) (x1,b) (l,b)` with `x1 = r − ss·a/100000`.
/// The spec's text rect reaches halfway into the point: `ir = (x1+r)/2`.
pub fn home_plate(def: &PresetGeometryDef, extent: PtSize) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    let ss = w.min(h).max(f32::EPSILON);
    let a = pin(0.0, adjust(def, "adj", 50000.0), 100000.0 * w / ss);
    let x1 = w - ss * a / 100000.0;
    polygon(
        &[(0.0, 0.0), (x1, 0.0), (w, h / 2.0), (x1, h), (0.0, h)],
        rect(0.0, 0.0, (x1 + w) / 2.0, h),
    )
}

/// `diamond`: the box's four edge midpoints; text in the inscribed
/// half-size box (`il = wd4 … ib = hd4·3`).
pub fn diamond(_def: &PresetGeometryDef, extent: PtSize) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    polygon(
        &[(0.0, h / 2.0), (w / 2.0, 0.0), (w, h / 2.0), (w / 2.0, h)],
        rect(w / 4.0, h / 4.0, w * 0.75, h * 0.75),
    )
}

/// `triangle` (isoceles): `adj` default 50000 places the apex at
/// `w·pin(0, adj, 100000)/100000`; base on the bottom edge. The spec's
/// text rect spans the lower half between the two mid-slope points —
/// `il = apex/2`, `ir = (apex + r)/2`, `t = vc` — which is the quarter box
/// only at the centered default.
pub fn triangle(def: &PresetGeometryDef, extent: PtSize) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    let a = pin(0.0, adjust(def, "adj", 50000.0), 100000.0);
    let apex = w * a / 100000.0;
    polygon(
        &[(0.0, h), (apex, 0.0), (w, h)],
        rect(apex / 2.0, h / 2.0, (apex + w) / 2.0, h),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::PresetShapeType;

    fn def(preset: PresetShapeType) -> PresetGeometryDef {
        PresetGeometryDef {
            preset,
            adjust_values: vec![],
        }
    }

    fn points(p: &ShapePath) -> Vec<(f32, f32)> {
        p.paths[0]
            .verbs
            .iter()
            .filter_map(|v| match v {
                PathVerb::MoveTo(o) | PathVerb::LineTo(o) => Some((o.x.raw(), o.y.raw())),
                _ => None,
            })
            .collect()
    }

    /// 100×40 chevron, defaults: notch depth = ss·0.5 = 20.
    #[test]
    fn chevron_notches_by_half_the_short_side() {
        let p = chevron(
            &def(PresetShapeType::Chevron),
            PtSize::new(Pt::new(100.0), Pt::new(40.0)),
        );
        assert_eq!(
            points(&p),
            vec![
                (0.0, 0.0),
                (80.0, 0.0),
                (100.0, 20.0),
                (80.0, 40.0),
                (0.0, 40.0),
                (20.0, 20.0)
            ]
        );
    }

    #[test]
    fn home_plate_has_a_flat_tail_and_a_pointed_nose() {
        let p = home_plate(
            &def(PresetShapeType::HomePlate),
            PtSize::new(Pt::new(100.0), Pt::new(40.0)),
        );
        assert_eq!(
            points(&p),
            vec![
                (0.0, 0.0),
                (80.0, 0.0),
                (100.0, 20.0),
                (80.0, 40.0),
                (0.0, 40.0)
            ]
        );
    }

    #[test]
    fn diamond_touches_the_four_midpoints() {
        let p = diamond(
            &def(PresetShapeType::Diamond),
            PtSize::new(Pt::new(60.0), Pt::new(40.0)),
        );
        assert_eq!(
            points(&p),
            vec![(0.0, 20.0), (30.0, 0.0), (60.0, 20.0), (30.0, 40.0)]
        );
    }

    /// The default apex is centered; an adjust slides it along the top edge.
    #[test]
    fn triangle_apex_follows_its_adjust() {
        let p = triangle(
            &def(PresetShapeType::Triangle),
            PtSize::new(Pt::new(60.0), Pt::new(40.0)),
        );
        assert_eq!(points(&p)[1], (30.0, 0.0));
        let skewed = PresetGeometryDef {
            preset: PresetShapeType::Triangle,
            adjust_values: vec![crate::model::GeomGuide {
                name: "adj".into(),
                formula: "val 0".into(),
            }],
        };
        let p = triangle(&skewed, PtSize::new(Pt::new(60.0), Pt::new(40.0)));
        assert_eq!(points(&p)[1], (0.0, 0.0), "adj 0 = right triangle");
    }
}
