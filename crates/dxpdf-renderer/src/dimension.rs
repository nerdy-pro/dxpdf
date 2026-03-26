//! Point-based dimension type for layout and rendering.
//!
//! `Pt` is the f32-based unit used throughout the layout pipeline.
//! All OOXML integer units (Twips, EMU, HalfPoints, EighthPoints) are
//! converted to Pt before layout.

use std::fmt;
use std::iter::Sum;
use std::ops::{Add, AddAssign, Div, Mul, Neg, Sub, SubAssign};

/// A dimension in typographic points (1pt = 1/72 inch).
#[derive(Clone, Copy, Default, PartialEq, PartialOrd)]
pub struct Pt(f32);

impl Pt {
    pub const ZERO: Self = Self(0.0);
    pub const INFINITY: Self = Self(f32::INFINITY);

    pub const fn new(v: f32) -> Self {
        Self(v)
    }

    pub fn abs(self) -> Self {
        Self(self.0.abs())
    }

    pub fn max(self, other: Self) -> Self {
        Self(self.0.max(other.0))
    }

    pub fn min(self, other: Self) -> Self {
        Self(self.0.min(other.0))
    }

    pub fn raw(self) -> f32 {
        self.0
    }
}

impl fmt::Debug for Pt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:.2}pt", self.0)
    }
}

impl fmt::Display for Pt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:.2}pt", self.0)
    }
}

impl From<Pt> for f32 {
    fn from(pt: Pt) -> f32 {
        pt.0
    }
}

impl Add for Pt {
    type Output = Self;
    fn add(self, rhs: Self) -> Self {
        Self(self.0 + rhs.0)
    }
}

impl AddAssign for Pt {
    fn add_assign(&mut self, rhs: Self) {
        self.0 += rhs.0;
    }
}

impl Sub for Pt {
    type Output = Self;
    fn sub(self, rhs: Self) -> Self {
        Self(self.0 - rhs.0)
    }
}

impl SubAssign for Pt {
    fn sub_assign(&mut self, rhs: Self) {
        self.0 -= rhs.0;
    }
}

impl Neg for Pt {
    type Output = Self;
    fn neg(self) -> Self {
        Self(-self.0)
    }
}

impl Mul<f32> for Pt {
    type Output = Self;
    fn mul(self, rhs: f32) -> Self {
        Self(self.0 * rhs)
    }
}

impl Div<f32> for Pt {
    type Output = Self;
    fn div(self, rhs: f32) -> Self {
        Self(self.0 / rhs)
    }
}

impl Div<Pt> for Pt {
    type Output = f32;
    fn div(self, rhs: Pt) -> f32 {
        self.0 / rhs.0
    }
}

impl Sum for Pt {
    fn sum<I: Iterator<Item = Self>>(iter: I) -> Self {
        Self(iter.map(|p| p.0).sum())
    }
}

// ── Conversions from OOXML units ─────────────────────────────────────────

use dxpdf_docx_model::dimension::{Dimension, EighthPoints, Emu, HalfPoints, Twips};

impl From<Dimension<Twips>> for Pt {
    /// 1 twip = 1/20 pt.
    fn from(d: Dimension<Twips>) -> Self {
        Self(d.raw() as f32 / 20.0)
    }
}

impl From<Dimension<HalfPoints>> for Pt {
    /// 1 half-point = 0.5 pt.
    fn from(d: Dimension<HalfPoints>) -> Self {
        Self(d.raw() as f32 / 2.0)
    }
}

impl From<Dimension<Emu>> for Pt {
    /// 1 EMU = 1/12700 pt.
    fn from(d: Dimension<Emu>) -> Self {
        Self(d.raw() as f32 / 12700.0)
    }
}

impl From<Dimension<EighthPoints>> for Pt {
    /// 1 eighth-point = 1/8 pt.
    fn from(d: Dimension<EighthPoints>) -> Self {
        Self(d.raw() as f32 / 8.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pt_zero() {
        assert_eq!(Pt::ZERO.raw(), 0.0);
    }

    #[test]
    fn pt_arithmetic() {
        let a = Pt::new(10.0);
        let b = Pt::new(3.0);
        assert_eq!((a + b).raw(), 13.0);
        assert_eq!((a - b).raw(), 7.0);
        assert_eq!((a * 2.0).raw(), 20.0);
        assert_eq!((a / 2.0).raw(), 5.0);
        assert_eq!(a / b, 10.0 / 3.0);
    }

    #[test]
    fn pt_add_assign() {
        let mut a = Pt::new(5.0);
        a += Pt::new(3.0);
        assert_eq!(a.raw(), 8.0);
    }

    #[test]
    fn pt_sub_assign() {
        let mut a = Pt::new(5.0);
        a -= Pt::new(2.0);
        assert_eq!(a.raw(), 3.0);
    }

    #[test]
    fn pt_neg() {
        assert_eq!((-Pt::new(5.0)).raw(), -5.0);
    }

    #[test]
    fn pt_abs() {
        assert_eq!(Pt::new(-3.0).abs().raw(), 3.0);
        assert_eq!(Pt::new(3.0).abs().raw(), 3.0);
    }

    #[test]
    fn pt_max_min() {
        let a = Pt::new(5.0);
        let b = Pt::new(10.0);
        assert_eq!(a.max(b).raw(), 10.0);
        assert_eq!(a.min(b).raw(), 5.0);
    }

    #[test]
    fn pt_sum() {
        let pts = vec![Pt::new(1.0), Pt::new(2.0), Pt::new(3.0)];
        let total: Pt = pts.into_iter().sum();
        assert_eq!(total.raw(), 6.0);
    }

    #[test]
    fn twips_to_pt() {
        // 720 twips = 36 pt
        assert_eq!(Pt::from(Dimension::<Twips>::new(720)).raw(), 36.0);
    }

    #[test]
    fn half_points_to_pt() {
        // 24 half-points = 12 pt
        assert_eq!(Pt::from(Dimension::<HalfPoints>::new(24)).raw(), 12.0);
    }

    #[test]
    fn emu_to_pt() {
        // 914400 EMU = 72 pt (1 inch)
        let pt = Pt::from(Dimension::<Emu>::new(914400));
        assert!((pt.raw() - 72.0).abs() < 0.01);
    }

    #[test]
    fn eighth_points_to_pt() {
        // 8 eighth-points = 1 pt
        assert_eq!(Pt::from(Dimension::<EighthPoints>::new(8)).raw(), 1.0);
    }

    #[test]
    fn pt_display() {
        assert_eq!(format!("{}", Pt::new(12.5)), "12.50pt");
    }

    #[test]
    fn pt_infinity() {
        assert!(Pt::INFINITY.raw().is_infinite());
    }
}
