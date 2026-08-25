//! §20.1.9.18 `ellipse` — the full ellipse inscribed in the bounding box,
//! four quarter-arc cubics. The spec's text rectangle is the box of the
//! ellipse's ±45° points: inset `(1 − cos 45°)/2 ≈ 0.14645` of each
//! dimension per side (guides `il = wd2·29289/100000` etc.).

use crate::model::{PathFillMode, PresetGeometryDef};
use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::resolve::shape_geometry::{PathVerb, ShapePath, SubPath};

use super::common::{corner_arc, pt};

pub fn build(_def: &PresetGeometryDef, extent: PtSize) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    let (hc, vc) = (w / 2.0, h / 2.0);

    // Start at 3 o'clock, clockwise; each quarter bulges toward its box
    // corner.
    let mut verbs = vec![PathVerb::MoveTo(pt(w, vc))];
    corner_arc(&mut verbs, pt(w, vc), pt(w, h), pt(hc, h));
    corner_arc(&mut verbs, pt(hc, h), pt(0.0, h), pt(0.0, vc));
    corner_arc(&mut verbs, pt(0.0, vc), pt(0.0, 0.0), pt(hc, 0.0));
    corner_arc(&mut verbs, pt(hc, 0.0), pt(w, 0.0), pt(w, vc));
    verbs.push(PathVerb::Close);

    let (ix, iy) = (w * 0.14645, h * 0.14645);
    ShapePath {
        paths: vec![SubPath {
            verbs,
            fill_mode: PathFillMode::Norm,
            stroked: true,
        }],
        text_rect: Some(PtRect::from_xywh(
            Pt::new(ix),
            Pt::new(iy),
            Pt::new(w - 2.0 * ix),
            Pt::new(h - 2.0 * iy),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::PresetShapeType;

    fn build_default(w: f32, h: f32) -> ShapePath {
        build(
            &PresetGeometryDef {
                preset: PresetShapeType::Ellipse,
                adjust_values: vec![],
            },
            PtSize::new(Pt::new(w), Pt::new(h)),
        )
    }

    #[test]
    fn four_quarter_arcs_close_the_ellipse() {
        let p = build_default(80.0, 40.0);
        let cubics = p.paths[0]
            .verbs
            .iter()
            .filter(|v| matches!(v, PathVerb::CubicTo(..)))
            .count();
        assert_eq!(cubics, 4);
        // Starts and ends at the 3 o'clock point.
        let PathVerb::MoveTo(start) = p.paths[0].verbs[0] else {
            panic!()
        };
        assert_eq!((start.x.raw(), start.y.raw()), (80.0, 20.0));
        let PathVerb::CubicTo(_, _, end) = p.paths[0].verbs[4] else {
            panic!()
        };
        assert_eq!((end.x.raw(), end.y.raw()), (80.0, 20.0));
    }

    /// The quarter-arc through 6 o'clock passes through the box's bottom
    /// center — the cubic's endpoint, exact by construction.
    #[test]
    fn arcs_touch_the_box_midpoints() {
        let p = build_default(80.0, 40.0);
        let PathVerb::CubicTo(_, _, bottom) = p.paths[0].verbs[1] else {
            panic!()
        };
        assert_eq!((bottom.x.raw(), bottom.y.raw()), (40.0, 40.0));
    }

    #[test]
    fn text_rect_is_the_inscribed_45_degree_box() {
        let p = build_default(100.0, 100.0);
        let tr = p.text_rect.unwrap();
        assert!((tr.origin.x.raw() - 14.645).abs() < 0.01);
        assert!((tr.size.width.raw() - 70.71).abs() < 0.01);
    }
}
