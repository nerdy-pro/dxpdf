//! §20.1.9.18 block arrows — `rightArrow`, `leftArrow`, `upArrow`,
//! `downArrow`.
//!
//! One spec definition, four orientations. For `rightArrow` the guides are:
//! `adj1` (default 50000) — shaft thickness as a fraction of the height,
//! `adj2` (default 50000) — head length in units of `ss = min(w, h)`,
//! clamped to `maxAdj2 = 100000·w/ss` so the head never outgrows the shape:
//!
//! ```text
//! dx1 = ss·a2/100000        head length
//! x1  = r − dx1             where the head begins
//! dy1 = h·a1/200000         half the shaft thickness
//! y1, y2 = vc ∓ dy1         shaft top/bottom
//! ```
//!
//! The other three transpose/mirror the same seven points. The spec's text
//! rectangle is the shaft plus the wedge of head between the shaft's edges
//! (`l, y1, x1 + y1·dx1/hd2, y2` for rightArrow).

use crate::model::{PathFillMode, PresetGeometryDef};
use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::resolve::shape_geometry::{PathVerb, ShapePath, SubPath};

use super::common::{adjust, pin, pt};

/// Which way the head points.
#[derive(Clone, Copy)]
pub enum Direction {
    Right,
    Left,
    Up,
    Down,
}

pub fn build(def: &PresetGeometryDef, extent: PtSize, dir: Direction) -> ShapePath {
    let (w, h) = (extent.width.raw(), extent.height.raw());
    let ss = w.min(h).max(f32::EPSILON);
    let a1 = pin(0.0, adjust(def, "adj1", 50000.0), 100000.0);

    // The head runs along the arrow's long axis; the shaft across it.
    let (along, across) = match dir {
        Direction::Right | Direction::Left => (w, h),
        Direction::Up | Direction::Down => (h, w),
    };
    let max_adj2 = 100000.0 * along / ss;
    let a2 = pin(0.0, adjust(def, "adj2", 50000.0), max_adj2);
    let head = ss * a2 / 100000.0;
    let half_shaft = across * a1 / 200000.0;
    let (c1, c2) = (across / 2.0 - half_shaft, across / 2.0 + half_shaft);
    let x1 = along - head;

    // Seven points in (along, across) space, tail first, clockwise around
    // the head.
    let points = [
        (0.0, c1),
        (x1, c1),
        (x1, 0.0),
        (along, across / 2.0),
        (x1, across),
        (x1, c2),
        (0.0, c2),
    ];
    // Map (along, across) into the box for this orientation.
    let map = |(a, c): (f32, f32)| match dir {
        Direction::Right => (a, c),
        Direction::Left => (w - a, c),
        Direction::Down => (c, a),
        Direction::Up => (c, h - a),
    };

    let mut verbs = Vec::with_capacity(9);
    for (i, p) in points.iter().enumerate() {
        let (x, y) = map(*p);
        verbs.push(if i == 0 {
            PathVerb::MoveTo(pt(x, y))
        } else {
            PathVerb::LineTo(pt(x, y))
        });
    }
    verbs.push(PathVerb::Close);

    // §20.1.9.18 rightArrow rect: past the shaft into the head, to where
    // the shaft's edge meets the head slope — `dx2 = y1·dx1/hd2`,
    // `r = x1 + dx2` — transposed into the (along, across) frame.
    let half = across / 2.0;
    let dx2 = if half > 0.0 { c1 * head / half } else { 0.0 };
    let text_end = (x1 + dx2).min(along);
    let (sa, sb) = map((0.0, c1));
    let (ea, eb) = map((text_end, c2));
    let (tx0, tx1) = (sa.min(ea), sa.max(ea));
    let (ty0, ty1) = (sb.min(eb), sb.max(eb));

    ShapePath {
        paths: vec![SubPath {
            verbs,
            fill_mode: PathFillMode::Norm,
            stroked: true,
        }],
        text_rect: Some(PtRect::from_xywh(
            Pt::new(tx0),
            Pt::new(ty0),
            Pt::new(tx1 - tx0),
            Pt::new(ty1 - ty0),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::PresetShapeType;

    fn def() -> PresetGeometryDef {
        PresetGeometryDef {
            preset: PresetShapeType::RightArrow,
            adjust_values: vec![],
        }
    }

    fn xy(v: &PathVerb) -> (f32, f32) {
        match v {
            PathVerb::MoveTo(o) | PathVerb::LineTo(o) => (o.x.raw(), o.y.raw()),
            _ => panic!("polygon verbs only"),
        }
    }

    /// 100×40 rightArrow, defaults: head = ss·0.5 = 20 → begins at x=80;
    /// shaft half-thickness = 40·0.25 = 10 → shaft spans y 10..30; tip at
    /// the middle of the right edge.
    #[test]
    fn right_arrow_default_geometry() {
        let p = build(
            &def(),
            PtSize::new(Pt::new(100.0), Pt::new(40.0)),
            Direction::Right,
        );
        let v = &p.paths[0].verbs;
        assert_eq!(v.len(), 8, "seven points + close");
        assert_eq!(xy(&v[0]), (0.0, 10.0), "tail top");
        assert_eq!(xy(&v[1]), (80.0, 10.0), "shaft meets head");
        assert_eq!(xy(&v[2]), (80.0, 0.0), "head top barb");
        assert_eq!(xy(&v[3]), (100.0, 20.0), "tip");
        assert_eq!(xy(&v[6]), (0.0, 30.0), "tail bottom");
        let tr = p.text_rect.unwrap();
        assert_eq!(
            (
                tr.origin.x.raw(),
                tr.origin.y.raw(),
                tr.size.width.raw(),
                tr.size.height.raw()
            ),
            (0.0, 10.0, 90.0, 20.0),
            "text spans the shaft and the head wedge (r = 80 + 10·20/20)"
        );
    }

    /// The four orientations are exact transposes/mirrors: tips sit at the
    /// middle of the pointed edge.
    #[test]
    fn each_direction_points_its_own_way() {
        let sz = PtSize::new(Pt::new(100.0), Pt::new(40.0));
        let tip = |d| {
            let p = build(&def(), sz, d);
            xy(&p.paths[0].verbs[3])
        };
        assert_eq!(tip(Direction::Right), (100.0, 20.0));
        assert_eq!(tip(Direction::Left), (0.0, 20.0));
        assert_eq!(tip(Direction::Down), (50.0, 40.0));
        assert_eq!(tip(Direction::Up), (50.0, 0.0));
    }

    /// adj2 is clamped so the head cannot be longer than the shape: on a
    /// 30×60 down-arrow, ss = 30 and maxAdj2 = 100000·60/30, so adj2 =
    /// 300000 clamps to a head of 60 — the whole height, a pure triangle.
    #[test]
    fn head_length_clamps_to_the_long_axis() {
        let d = PresetGeometryDef {
            preset: PresetShapeType::DownArrow,
            adjust_values: vec![crate::model::GeomGuide {
                name: "adj2".into(),
                formula: "val 999999".into(),
            }],
        };
        let p = build(
            &d,
            PtSize::new(Pt::new(30.0), Pt::new(60.0)),
            Direction::Down,
        );
        let v = &p.paths[0].verbs;
        // Head begins at along = 60 − 60 = 0 → the barbs sit at y = 0.
        assert_eq!(xy(&v[1]).1, 0.0);
    }
}
