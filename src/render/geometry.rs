//! Point-space geometry types for layout and rendering.

use crate::render::dimension::Pt;

/// A 2D offset in points.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PtOffset {
    pub x: Pt,
    pub y: Pt,
}

impl PtOffset {
    pub const ZERO: Self = Self {
        x: Pt::ZERO,
        y: Pt::ZERO,
    };

    pub const fn new(x: Pt, y: Pt) -> Self {
        Self { x, y }
    }
}

impl std::ops::Add for PtOffset {
    type Output = Self;
    fn add(self, rhs: Self) -> Self {
        Self {
            x: self.x + rhs.x,
            y: self.y + rhs.y,
        }
    }
}

/// A 2D size in points.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PtSize {
    pub width: Pt,
    pub height: Pt,
}

impl PtSize {
    pub const ZERO: Self = Self {
        width: Pt::ZERO,
        height: Pt::ZERO,
    };

    pub const fn new(width: Pt, height: Pt) -> Self {
        Self { width, height }
    }
}

/// Axis-aligned rectangle in points.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PtRect {
    pub origin: PtOffset,
    pub size: PtSize,
}

impl PtRect {
    pub fn from_xywh(x: Pt, y: Pt, w: Pt, h: Pt) -> Self {
        Self {
            origin: PtOffset::new(x, y),
            size: PtSize::new(w, h),
        }
    }
}

/// Edge insets (padding/margins) in points.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PtEdgeInsets {
    pub top: Pt,
    pub right: Pt,
    pub bottom: Pt,
    pub left: Pt,
}

impl PtEdgeInsets {
    pub const ZERO: Self = Self {
        top: Pt::ZERO,
        right: Pt::ZERO,
        bottom: Pt::ZERO,
        left: Pt::ZERO,
    };

    pub const fn new(top: Pt, right: Pt, bottom: Pt, left: Pt) -> Self {
        Self {
            top,
            right,
            bottom,
            left,
        }
    }

    pub fn horizontal(&self) -> Pt {
        self.left + self.right
    }

    pub fn vertical(&self) -> Pt {
        self.top + self.bottom
    }
}

/// A line segment in points.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct PtLineSegment {
    pub start: PtOffset,
    pub end: PtOffset,
}

impl PtLineSegment {
    pub const fn new(start: PtOffset, end: PtOffset) -> Self {
        Self { start, end }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn offset_add() {
        let a = PtOffset::new(Pt::new(1.0), Pt::new(2.0));
        let b = PtOffset::new(Pt::new(3.0), Pt::new(4.0));
        let c = a + b;
        assert_eq!(c.x.raw(), 4.0);
        assert_eq!(c.y.raw(), 6.0);
    }

    #[test]
    fn edge_insets_horizontal_vertical() {
        let ei = PtEdgeInsets::new(Pt::new(10.0), Pt::new(20.0), Pt::new(30.0), Pt::new(40.0));
        assert_eq!(ei.horizontal().raw(), 60.0);
        assert_eq!(ei.vertical().raw(), 40.0);
    }

    #[test]
    fn rect_from_xywh() {
        let r = PtRect::from_xywh(Pt::new(1.0), Pt::new(2.0), Pt::new(3.0), Pt::new(4.0));
        assert_eq!(r.origin.x.raw(), 1.0);
        assert_eq!(r.origin.y.raw(), 2.0);
        assert_eq!(r.size.width.raw(), 3.0);
        assert_eq!(r.size.height.raw(), 4.0);
    }
}
