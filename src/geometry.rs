//! Geometry types for spatial layout: Offset, Size, Rect, EdgeInsets, LineSegment.
//!
//! All types are generic over the unit marker, reusing `Dimension<U>` from the
//! `dimension` module. This means `Offset<PtUnit>` stores two `Pt` values,
//! `Size<TwipsUnit>` stores two `Twips` values, etc.

use std::fmt;

use crate::dimension::{Dimension, IntegerUnit, Pt, PtUnit, TwipsUnit, Unit, UnitLabel};

// ---------------------------------------------------------------------------
// Offset — a position or translation (x, y)
// ---------------------------------------------------------------------------

#[derive(Clone, Copy)]
pub struct Offset<U: Unit> {
    pub x: Dimension<U>,
    pub y: Dimension<U>,
}

impl<U: Unit> PartialEq for Offset<U> {
    fn eq(&self, other: &Self) -> bool {
        self.x == other.x && self.y == other.y
    }
}

impl<U: IntegerUnit> Eq for Offset<U> {}

impl<U: Unit> Offset<U> {
    pub fn new(x: Dimension<U>, y: Dimension<U>) -> Self {
        Self { x, y }
    }
}

impl Offset<PtUnit> {
    pub const ZERO: Self = Self {
        x: Pt::ZERO,
        y: Pt::ZERO,
    };

    /// Return a copy with y shifted by `dy`.
    pub fn offset_y(self, dy: Pt) -> Self {
        Self {
            x: self.x,
            y: self.y + dy,
        }
    }
}

impl<U: Unit + UnitLabel> fmt::Debug for Offset<U>
where
    Dimension<U>: fmt::Debug,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Offset({:?}, {:?})", self.x, self.y)
    }
}

// ---------------------------------------------------------------------------
// Size — width and height
// ---------------------------------------------------------------------------

#[derive(Clone, Copy)]
pub struct Size<U: Unit> {
    pub width: Dimension<U>,
    pub height: Dimension<U>,
}

impl<U: Unit> PartialEq for Size<U> {
    fn eq(&self, other: &Self) -> bool {
        self.width == other.width && self.height == other.height
    }
}

impl<U: IntegerUnit> Eq for Size<U> {}

impl<U: Unit> Size<U> {
    pub fn new(width: Dimension<U>, height: Dimension<U>) -> Self {
        Self { width, height }
    }
}

impl Size<PtUnit> {
    pub const ZERO: Self = Self {
        width: Pt::ZERO,
        height: Pt::ZERO,
    };
}

impl<U: Unit + UnitLabel> fmt::Debug for Size<U>
where
    Dimension<U>: fmt::Debug,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Size({:?} × {:?})", self.width, self.height)
    }
}

// ---------------------------------------------------------------------------
// Rect — position + size
// ---------------------------------------------------------------------------

#[derive(Clone, Copy)]
pub struct Rect<U: Unit> {
    pub origin: Offset<U>,
    pub size: Size<U>,
}

impl<U: Unit> PartialEq for Rect<U> {
    fn eq(&self, other: &Self) -> bool {
        self.origin == other.origin && self.size == other.size
    }
}

impl<U: IntegerUnit> Eq for Rect<U> {}

impl<U: Unit> Rect<U> {
    pub fn new(origin: Offset<U>, size: Size<U>) -> Self {
        Self { origin, size }
    }

    pub fn from_xywh(
        x: Dimension<U>,
        y: Dimension<U>,
        width: Dimension<U>,
        height: Dimension<U>,
    ) -> Self {
        Self {
            origin: Offset::new(x, y),
            size: Size::new(width, height),
        }
    }
}

impl Rect<PtUnit> {
    /// Return a copy with origin.y shifted by `dy`.
    pub fn offset_y(self, dy: Pt) -> Self {
        Self {
            origin: self.origin.offset_y(dy),
            size: self.size,
        }
    }
}

impl<U: Unit + UnitLabel> fmt::Debug for Rect<U>
where
    Dimension<U>: fmt::Debug,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "Rect({:?}, {:?}, {:?}, {:?})",
            self.origin.x, self.origin.y, self.size.width, self.size.height
        )
    }
}

// ---------------------------------------------------------------------------
// EdgeInsets — top, right, bottom, left
// ---------------------------------------------------------------------------

#[derive(Clone, Copy)]
pub struct EdgeInsets<U: Unit> {
    pub top: Dimension<U>,
    pub right: Dimension<U>,
    pub bottom: Dimension<U>,
    pub left: Dimension<U>,
}

impl<U: Unit> PartialEq for EdgeInsets<U> {
    fn eq(&self, other: &Self) -> bool {
        self.top == other.top
            && self.right == other.right
            && self.bottom == other.bottom
            && self.left == other.left
    }
}

impl<U: IntegerUnit> Eq for EdgeInsets<U> {}

impl<U: Unit> EdgeInsets<U> {
    pub fn new(
        top: Dimension<U>,
        right: Dimension<U>,
        bottom: Dimension<U>,
        left: Dimension<U>,
    ) -> Self {
        Self {
            top,
            right,
            bottom,
            left,
        }
    }
}

impl<U: Unit + UnitLabel> fmt::Debug for EdgeInsets<U>
where
    Dimension<U>: fmt::Debug,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "EdgeInsets(top: {:?}, right: {:?}, bottom: {:?}, left: {:?})",
            self.top, self.right, self.bottom, self.left
        )
    }
}

// ---------------------------------------------------------------------------
// LineSegment — two endpoints
// ---------------------------------------------------------------------------

#[derive(Clone, Copy)]
pub struct LineSegment<U: Unit> {
    pub start: Offset<U>,
    pub end: Offset<U>,
}

impl<U: Unit> PartialEq for LineSegment<U> {
    fn eq(&self, other: &Self) -> bool {
        self.start == other.start && self.end == other.end
    }
}

impl<U: IntegerUnit> Eq for LineSegment<U> {}

impl<U: Unit> LineSegment<U> {
    pub fn new(start: Offset<U>, end: Offset<U>) -> Self {
        Self { start, end }
    }
}

impl LineSegment<PtUnit> {
    /// Return a copy with both endpoints' y shifted by `dy`.
    pub fn offset_y(self, dy: Pt) -> Self {
        Self {
            start: self.start.offset_y(dy),
            end: self.end.offset_y(dy),
        }
    }
}

impl<U: Unit + UnitLabel> fmt::Debug for LineSegment<U>
where
    Dimension<U>: fmt::Debug,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "LineSegment({:?} → {:?})", self.start, self.end)
    }
}

// ---------------------------------------------------------------------------
// From conversions: integer-unit geometry → Pt geometry
// ---------------------------------------------------------------------------

impl<U: IntegerUnit> From<Offset<U>> for Offset<PtUnit>
where
    Pt: From<Dimension<U>>,
{
    fn from(o: Offset<U>) -> Self {
        Self {
            x: Pt::from(o.x),
            y: Pt::from(o.y),
        }
    }
}

impl<U: IntegerUnit> From<Size<U>> for Size<PtUnit>
where
    Pt: From<Dimension<U>>,
{
    fn from(s: Size<U>) -> Self {
        Self {
            width: Pt::from(s.width),
            height: Pt::from(s.height),
        }
    }
}

impl<U: IntegerUnit> From<Rect<U>> for Rect<PtUnit>
where
    Pt: From<Dimension<U>>,
{
    fn from(r: Rect<U>) -> Self {
        Self {
            origin: Offset::<PtUnit>::from(r.origin),
            size: Size::<PtUnit>::from(r.size),
        }
    }
}

impl<U: IntegerUnit> From<EdgeInsets<U>> for EdgeInsets<PtUnit>
where
    Pt: From<Dimension<U>>,
{
    fn from(e: EdgeInsets<U>) -> Self {
        Self {
            top: Pt::from(e.top),
            right: Pt::from(e.right),
            bottom: Pt::from(e.bottom),
            left: Pt::from(e.left),
        }
    }
}

// ---------------------------------------------------------------------------
// Skia interop
// ---------------------------------------------------------------------------

impl From<Offset<PtUnit>> for skia_safe::Point {
    fn from(o: Offset<PtUnit>) -> Self {
        skia_safe::Point::new(f32::from(o.x), f32::from(o.y))
    }
}

impl From<Size<PtUnit>> for skia_safe::Size {
    fn from(s: Size<PtUnit>) -> Self {
        skia_safe::Size::new(f32::from(s.width), f32::from(s.height))
    }
}

impl From<Rect<PtUnit>> for skia_safe::Rect {
    fn from(r: Rect<PtUnit>) -> Self {
        skia_safe::Rect::from_xywh(
            f32::from(r.origin.x),
            f32::from(r.origin.y),
            f32::from(r.size.width),
            f32::from(r.size.height),
        )
    }
}

// ---------------------------------------------------------------------------
// Type aliases
// ---------------------------------------------------------------------------

pub type PtOffset = Offset<PtUnit>;
pub type PtSize = Size<PtUnit>;
pub type PtRect = Rect<PtUnit>;
pub type PtEdgeInsets = EdgeInsets<PtUnit>;
pub type PtLineSegment = LineSegment<PtUnit>;

pub type TwipsSize = Size<TwipsUnit>;
pub type TwipsEdgeInsets = EdgeInsets<TwipsUnit>;

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::dimension::{Pt, Twips};

    // -- Offset --

    #[test]
    fn offset_new_and_fields() {
        let o = PtOffset::new(Pt::new(10.0), Pt::new(20.0));
        assert_eq!(f32::from(o.x), 10.0);
        assert_eq!(f32::from(o.y), 20.0);
    }

    #[test]
    fn offset_zero() {
        assert_eq!(PtOffset::ZERO.x, Pt::ZERO);
        assert_eq!(PtOffset::ZERO.y, Pt::ZERO);
    }

    #[test]
    fn offset_equality() {
        let a = PtOffset::new(Pt::new(1.0), Pt::new(2.0));
        let b = PtOffset::new(Pt::new(1.0), Pt::new(2.0));
        let c = PtOffset::new(Pt::new(3.0), Pt::new(2.0));
        assert_eq!(a, b);
        assert_ne!(a, c);
    }

    #[test]
    fn offset_debug() {
        let o = PtOffset::new(Pt::new(10.0), Pt::new(20.0));
        assert_eq!(format!("{:?}", o), "Offset(10pt, 20pt)");
    }

    #[test]
    fn offset_is_copy() {
        let o = PtOffset::new(Pt::new(1.0), Pt::new(2.0));
        let o2 = o;
        assert_eq!(o, o2);
    }

    // -- Size --

    #[test]
    fn size_new_and_fields() {
        let s = PtSize::new(Pt::new(612.0), Pt::new(792.0));
        assert_eq!(f32::from(s.width), 612.0);
        assert_eq!(f32::from(s.height), 792.0);
    }

    #[test]
    fn size_zero() {
        assert_eq!(PtSize::ZERO.width, Pt::ZERO);
        assert_eq!(PtSize::ZERO.height, Pt::ZERO);
    }

    #[test]
    fn size_debug() {
        let s = PtSize::new(Pt::new(100.0), Pt::new(50.0));
        assert_eq!(format!("{:?}", s), "Size(100pt × 50pt)");
    }

    #[test]
    fn twips_size_and_fields() {
        let s = TwipsSize::new(Twips::new(12240), Twips::new(15840));
        assert_eq!(i64::from(s.width), 12240);
        assert_eq!(i64::from(s.height), 15840);
    }

    #[test]
    fn twips_size_equality() {
        let a = TwipsSize::new(Twips::new(100), Twips::new(200));
        let b = TwipsSize::new(Twips::new(100), Twips::new(200));
        assert_eq!(a, b);
    }

    // -- Rect --

    #[test]
    fn rect_from_xywh() {
        let r = PtRect::from_xywh(Pt::new(10.0), Pt::new(20.0), Pt::new(100.0), Pt::new(50.0));
        assert_eq!(f32::from(r.origin.x), 10.0);
        assert_eq!(f32::from(r.origin.y), 20.0);
        assert_eq!(f32::from(r.size.width), 100.0);
        assert_eq!(f32::from(r.size.height), 50.0);
    }

    #[test]
    fn rect_new_from_offset_and_size() {
        let origin = PtOffset::new(Pt::new(5.0), Pt::new(10.0));
        let size = PtSize::new(Pt::new(200.0), Pt::new(100.0));
        let r = PtRect::new(origin, size);
        assert_eq!(r.origin, origin);
        assert_eq!(r.size, size);
    }

    #[test]
    fn rect_debug() {
        let r = PtRect::from_xywh(Pt::new(0.0), Pt::new(0.0), Pt::new(612.0), Pt::new(792.0));
        assert_eq!(format!("{:?}", r), "Rect(0pt, 0pt, 612pt, 792pt)");
    }

    // -- EdgeInsets --

    #[test]
    fn edge_insets_new_and_fields() {
        let e = TwipsEdgeInsets::new(
            Twips::new(1440),
            Twips::new(1440),
            Twips::new(1440),
            Twips::new(1440),
        );
        assert_eq!(i64::from(e.top), 1440);
        assert_eq!(i64::from(e.right), 1440);
        assert_eq!(i64::from(e.bottom), 1440);
        assert_eq!(i64::from(e.left), 1440);
    }

    #[test]
    fn edge_insets_debug() {
        let e = PtEdgeInsets::new(Pt::new(72.0), Pt::new(72.0), Pt::new(72.0), Pt::new(72.0));
        assert_eq!(
            format!("{:?}", e),
            "EdgeInsets(top: 72pt, right: 72pt, bottom: 72pt, left: 72pt)"
        );
    }

    // -- LineSegment --

    #[test]
    fn line_segment_new_and_fields() {
        let seg = PtLineSegment::new(
            PtOffset::new(Pt::new(0.0), Pt::new(10.0)),
            PtOffset::new(Pt::new(100.0), Pt::new(10.0)),
        );
        assert_eq!(f32::from(seg.start.x), 0.0);
        assert_eq!(f32::from(seg.end.x), 100.0);
        assert_eq!(seg.start.y, seg.end.y);
    }

    #[test]
    fn line_segment_debug() {
        let seg = PtLineSegment::new(
            PtOffset::new(Pt::new(0.0), Pt::new(0.0)),
            PtOffset::new(Pt::new(100.0), Pt::new(50.0)),
        );
        assert_eq!(
            format!("{:?}", seg),
            "LineSegment(Offset(0pt, 0pt) → Offset(100pt, 50pt))"
        );
    }

    // -- Cross-unit conversions --

    #[test]
    fn twips_size_to_pt_size() {
        let ts = TwipsSize::new(Twips::new(12240), Twips::new(15840));
        let ps: PtSize = ts.into();
        assert!((f32::from(ps.width) - 612.0).abs() < 0.01);
        assert!((f32::from(ps.height) - 792.0).abs() < 0.01);
    }

    #[test]
    fn twips_edge_insets_to_pt() {
        let te = TwipsEdgeInsets::new(
            Twips::new(1440),
            Twips::new(1440),
            Twips::new(1440),
            Twips::new(1440),
        );
        let pe: PtEdgeInsets = te.into();
        assert_eq!(f32::from(pe.top), 72.0);
        assert_eq!(f32::from(pe.right), 72.0);
    }

    #[test]
    fn twips_offset_to_pt() {
        use crate::dimension::TwipsUnit;
        let to = Offset::<TwipsUnit>::new(Twips::new(1440), Twips::new(2880));
        let po: PtOffset = to.into();
        assert_eq!(f32::from(po.x), 72.0);
        assert_eq!(f32::from(po.y), 144.0);
    }

    #[test]
    fn twips_rect_to_pt() {
        use crate::dimension::TwipsUnit;
        let tr = Rect::<TwipsUnit>::from_xywh(
            Twips::new(0),
            Twips::new(0),
            Twips::new(12240),
            Twips::new(15840),
        );
        let pr: PtRect = tr.into();
        assert!((f32::from(pr.size.width) - 612.0).abs() < 0.01);
    }

    // -- Skia interop --

    #[test]
    fn pt_offset_into_skia_point() {
        let o = PtOffset::new(Pt::new(72.0), Pt::new(36.0));
        let p: skia_safe::Point = o.into();
        assert_eq!(p.x, 72.0);
        assert_eq!(p.y, 36.0);
    }

    #[test]
    fn pt_size_into_skia_size() {
        let s = PtSize::new(Pt::new(612.0), Pt::new(792.0));
        let ss: skia_safe::Size = s.into();
        assert_eq!(ss.width, 612.0);
        assert_eq!(ss.height, 792.0);
    }

    #[test]
    fn pt_rect_into_skia_rect() {
        let r = PtRect::from_xywh(Pt::new(10.0), Pt::new(20.0), Pt::new(100.0), Pt::new(50.0));
        let sr: skia_safe::Rect = r.into();
        assert_eq!(sr.left, 10.0);
        assert_eq!(sr.top, 20.0);
        assert_eq!(sr.width(), 100.0);
        assert_eq!(sr.height(), 50.0);
    }
}
