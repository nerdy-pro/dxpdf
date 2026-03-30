//! Conversions between dxpdf renderer types and Skia types.

use crate::geometry::{PtLineSegment, PtOffset, PtRect, PtSize};
use crate::resolve::color::RgbColor;

pub fn to_point(o: PtOffset) -> skia_safe::Point {
    skia_safe::Point::new(f32::from(o.x), f32::from(o.y))
}

pub fn to_size(s: PtSize) -> skia_safe::Size {
    skia_safe::Size::new(f32::from(s.width), f32::from(s.height))
}

pub fn to_rect(r: PtRect) -> skia_safe::Rect {
    skia_safe::Rect::from_xywh(
        f32::from(r.origin.x),
        f32::from(r.origin.y),
        f32::from(r.size.width),
        f32::from(r.size.height),
    )
}

pub fn to_color4f(c: RgbColor) -> skia_safe::Color4f {
    const MAX: f32 = u8::MAX as f32;
    skia_safe::Color4f::new(c.r as f32 / MAX, c.g as f32 / MAX, c.b as f32 / MAX, 1.0)
}

pub fn to_line(l: PtLineSegment) -> (skia_safe::Point, skia_safe::Point) {
    (to_point(l.start), to_point(l.end))
}
