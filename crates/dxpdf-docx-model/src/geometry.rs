use crate::dimension::{Dimension, Unit};

/// A 2D offset (x, y) in a given unit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Offset<U: Unit> {
    pub x: Dimension<U>,
    pub y: Dimension<U>,
}

impl<U: Unit> Offset<U> {
    pub const ZERO: Self = Self {
        x: Dimension::ZERO,
        y: Dimension::ZERO,
    };

    pub const fn new(x: Dimension<U>, y: Dimension<U>) -> Self {
        Self { x, y }
    }
}

/// A 2D size (width, height) in a given unit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Size<U: Unit> {
    pub width: Dimension<U>,
    pub height: Dimension<U>,
}

impl<U: Unit> Size<U> {
    pub const ZERO: Self = Self {
        width: Dimension::ZERO,
        height: Dimension::ZERO,
    };

    pub const fn new(width: Dimension<U>, height: Dimension<U>) -> Self {
        Self { width, height }
    }
}

/// A rectangle defined by origin + size.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Rect<U: Unit> {
    pub origin: Offset<U>,
    pub size: Size<U>,
}

impl<U: Unit> Rect<U> {
    pub const fn new(origin: Offset<U>, size: Size<U>) -> Self {
        Self { origin, size }
    }
}

/// Insets from each edge (top, right, bottom, left) — used for margins and padding.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct EdgeInsets<U: Unit> {
    pub top: Dimension<U>,
    pub right: Dimension<U>,
    pub bottom: Dimension<U>,
    pub left: Dimension<U>,
}

impl<U: Unit> EdgeInsets<U> {
    pub const ZERO: Self = Self {
        top: Dimension::ZERO,
        right: Dimension::ZERO,
        bottom: Dimension::ZERO,
        left: Dimension::ZERO,
    };

    pub const fn new(
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

    pub const fn uniform(value: Dimension<U>) -> Self {
        Self {
            top: value,
            right: value,
            bottom: value,
            left: value,
        }
    }
}
