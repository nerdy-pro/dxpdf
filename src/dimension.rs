//! Type-safe dimensional units for OOXML and layout values.
//!
//! Each OOXML unit (twips, half-points, eighth-points, EMU) is stored as an
//! integer to preserve the original lossless representation from the XML.
//! Conversion to points (`Pt`) for rendering produces `f32` values.

use std::fmt;
use std::marker::PhantomData;
use std::ops::{Add, Div, Mul, Neg, Sub};

// ---------------------------------------------------------------------------
// Unit markers (zero-sized types)
// ---------------------------------------------------------------------------

/// Twips — 1/20 of a point. The primary OOXML structural unit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TwipsUnit;

/// Half-points — 1/2 of a point. Used for font sizes (`w:sz`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HalfPointsUnit;

/// Eighth-points — 1/8 of a point. Used for border widths.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EighthPointsUnit;

/// English Metric Units — 1/914400 of an inch. Used for DrawingML dimensions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EmuUnit;

/// Typographic points — 1/72 of an inch. The layout/rendering unit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PtUnit;

// ---------------------------------------------------------------------------
// Dimension<U> — the core type
// ---------------------------------------------------------------------------

/// A dimensional value parameterized by its unit of measurement.
///
/// OOXML units store an `i64` to preserve the original integer values from XML.
/// `Pt` stores `f32` since it is the output of conversion for rendering.
#[derive(Clone, Copy)]
pub struct Dimension<U: Unit> {
    value: DimStorage<U>,
    _unit: PhantomData<U>,
}

impl<U: Unit> PartialEq for Dimension<U> {
    fn eq(&self, other: &Self) -> bool {
        self.value == other.value
    }
}

// Integer-based dimensions support Eq (f32 does not).
impl<U: IntegerUnit> Eq for Dimension<U> {}

// ---------------------------------------------------------------------------
// Storage: integer for OOXML units, f32 for Pt
// ---------------------------------------------------------------------------

/// Internal storage trait — not public, just selects i64 vs f32.
pub trait Unit {
    type Storage: Copy + PartialEq;
}

/// Marker trait for units backed by integer storage.
pub trait IntegerUnit: Unit<Storage = i64> {}

impl Unit for TwipsUnit {
    type Storage = i64;
}
impl IntegerUnit for TwipsUnit {}

impl Unit for HalfPointsUnit {
    type Storage = i64;
}
impl IntegerUnit for HalfPointsUnit {}

impl Unit for EighthPointsUnit {
    type Storage = i64;
}
impl IntegerUnit for EighthPointsUnit {}

impl Unit for EmuUnit {
    type Storage = i64;
}
impl IntegerUnit for EmuUnit {}

impl Unit for PtUnit {
    type Storage = f32;
}

/// Type-level storage selector.
///
/// For integer units this is `i64`, for `PtUnit` it is `f32`.
type DimStorage<U> = <U as Unit>::Storage;

// ---------------------------------------------------------------------------
// Type aliases for ergonomic use
// ---------------------------------------------------------------------------

/// Dimension in twips (1/20 pt).
pub type Twips = Dimension<TwipsUnit>;

/// Dimension in half-points (1/2 pt), used for font sizes.
pub type HalfPoints = Dimension<HalfPointsUnit>;

/// Dimension in eighth-points (1/8 pt), used for border widths.
pub type EighthPoints = Dimension<EighthPointsUnit>;

/// Dimension in English Metric Units (1/914400 in).
pub type Emu = Dimension<EmuUnit>;

/// Dimension in typographic points (1/72 in).
pub type Pt = Dimension<PtUnit>;

// ---------------------------------------------------------------------------
// Constructors
// ---------------------------------------------------------------------------

impl<U: IntegerUnit> Dimension<U> {
    pub const fn new(value: i64) -> Self {
        Self {
            value,
            _unit: PhantomData,
        }
    }

    pub const fn is_positive(self) -> bool {
        self.value > 0
    }
}

impl Pt {
    pub fn new(value: f32) -> Self {
        Self {
            value,
            _unit: PhantomData,
        }
    }
}

// ---------------------------------------------------------------------------
// Into primitive: extract the inner value
// ---------------------------------------------------------------------------

impl<U: IntegerUnit> From<Dimension<U>> for i64 {
    fn from(d: Dimension<U>) -> i64 {
        d.value
    }
}

impl From<Pt> for f32 {
    fn from(p: Pt) -> f32 {
        p.value
    }
}

// ---------------------------------------------------------------------------
// Arithmetic: Add, Sub, Neg (same-unit only)
// ---------------------------------------------------------------------------

impl<U: IntegerUnit> Add for Dimension<U> {
    type Output = Self;
    fn add(self, rhs: Self) -> Self {
        Self::new(self.value + rhs.value)
    }
}

impl Add for Pt {
    type Output = Self;
    fn add(self, rhs: Self) -> Self {
        Self::new(self.value + rhs.value)
    }
}

impl<U: IntegerUnit> Sub for Dimension<U> {
    type Output = Self;
    fn sub(self, rhs: Self) -> Self {
        Self::new(self.value - rhs.value)
    }
}

impl Sub for Pt {
    type Output = Self;
    fn sub(self, rhs: Self) -> Self {
        Self::new(self.value - rhs.value)
    }
}

impl<U: IntegerUnit> Neg for Dimension<U> {
    type Output = Self;
    fn neg(self) -> Self {
        Self::new(-self.value)
    }
}

impl Neg for Pt {
    type Output = Self;
    fn neg(self) -> Self {
        Self::new(-self.value)
    }
}

// ---------------------------------------------------------------------------
// Pt-specific arithmetic: scalar multiply/divide, ordering
// ---------------------------------------------------------------------------

/// Scale a point value by a dimensionless factor.
impl Mul<f32> for Pt {
    type Output = Self;
    fn mul(self, rhs: f32) -> Self {
        Self::new(self.value * rhs)
    }
}

/// Scale a point value by a dimensionless factor (f32 * Pt).
impl Mul<Pt> for f32 {
    type Output = Pt;
    fn mul(self, rhs: Pt) -> Pt {
        Pt::new(self * rhs.value)
    }
}

/// Divide a point value by a dimensionless factor.
impl Div<f32> for Pt {
    type Output = Self;
    fn div(self, rhs: f32) -> Self {
        Self::new(self.value / rhs)
    }
}

/// Divide two point values to get a dimensionless ratio.
impl Div<Pt> for Pt {
    type Output = f32;
    fn div(self, rhs: Pt) -> f32 {
        self.value / rhs.value
    }
}

impl PartialOrd for Pt {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        self.value.partial_cmp(&other.value)
    }
}

// ---------------------------------------------------------------------------
// From conversions: OOXML units → Pt (and transitively → f32)
// ---------------------------------------------------------------------------

impl From<Twips> for Pt {
    fn from(t: Twips) -> Self {
        Pt::new(t.value as f32 / 20.0)
    }
}

impl From<HalfPoints> for Pt {
    fn from(hp: HalfPoints) -> Self {
        Pt::new(hp.value as f32 / 2.0)
    }
}

impl From<EighthPoints> for Pt {
    fn from(ep: EighthPoints) -> Self {
        Pt::new(ep.value as f32 / 8.0)
    }
}

impl From<Emu> for Pt {
    fn from(e: Emu) -> Self {
        // 914400 EMU = 1 inch = 72 pt → 1 EMU = 72/914400 pt
        Pt::new(e.value as f32 * 72.0 / 914_400.0)
    }
}

// ---------------------------------------------------------------------------
// Shortcut conversions: OOXML units → f32 (via Pt)
// ---------------------------------------------------------------------------

impl From<Twips> for f32 {
    fn from(t: Twips) -> f32 {
        f32::from(Pt::from(t))
    }
}

impl From<HalfPoints> for f32 {
    fn from(hp: HalfPoints) -> f32 {
        f32::from(Pt::from(hp))
    }
}

impl From<EighthPoints> for f32 {
    fn from(ep: EighthPoints) -> f32 {
        f32::from(Pt::from(ep))
    }
}

impl From<Emu> for f32 {
    fn from(e: Emu) -> f32 {
        f32::from(Pt::from(e))
    }
}

// ---------------------------------------------------------------------------
// Debug & Display
// ---------------------------------------------------------------------------

impl<U: IntegerUnit> fmt::Debug for Dimension<U>
where
    U: UnitLabel,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", self.value, U::LABEL)
    }
}

impl fmt::Debug for Pt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}pt", self.value)
    }
}

impl<U: IntegerUnit> fmt::Display for Dimension<U>
where
    U: UnitLabel,
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", self.value, U::LABEL)
    }
}

impl fmt::Display for Pt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}pt", self.value)
    }
}

/// Label suffix for Debug/Display output.
pub trait UnitLabel {
    const LABEL: &'static str;
}

impl UnitLabel for TwipsUnit {
    const LABEL: &'static str = "tw";
}

impl UnitLabel for HalfPointsUnit {
    const LABEL: &'static str = "hp";
}

impl UnitLabel for EighthPointsUnit {
    const LABEL: &'static str = "ep";
}

impl UnitLabel for EmuUnit {
    const LABEL: &'static str = "emu";
}

impl UnitLabel for PtUnit {
    const LABEL: &'static str = "pt";
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    // -- Construction & raw access --

    #[test]
    fn twips_new_and_raw() {
        let t = Twips::new(1440);
        assert_eq!(i64::from(t), 1440);
    }

    #[test]
    fn half_points_new_and_raw() {
        let hp = HalfPoints::new(24);
        assert_eq!(i64::from(hp), 24);
    }

    #[test]
    fn eighth_points_new_and_raw() {
        let ep = EighthPoints::new(8);
        assert_eq!(i64::from(ep), 8);
    }

    #[test]
    fn emu_new_and_raw() {
        let e = Emu::new(914400);
        assert_eq!(i64::from(e), 914400);
    }

    #[test]
    fn pt_new_and_raw() {
        let p = Pt::new(72.0);
        assert_eq!(f32::from(p), 72.0);
    }

    #[test]
    fn negative_values() {
        let t = Twips::new(-360);
        assert_eq!(i64::from(t), -360);

        let e = Emu::new(-457200);
        assert_eq!(i64::from(e), -457200);
    }

    #[test]
    fn zero_values() {
        assert_eq!(i64::from(Twips::new(0)), 0);
        assert_eq!(f32::from(Pt::new(0.0)), 0.0);
    }

    // -- Arithmetic --

    #[test]
    fn twips_add() {
        let a = Twips::new(720);
        let b = Twips::new(360);
        assert_eq!(i64::from(a + b), 1080);
    }

    #[test]
    fn twips_sub() {
        let a = Twips::new(720);
        let b = Twips::new(360);
        assert_eq!(i64::from(a - b), 360);
    }

    #[test]
    fn twips_neg() {
        let t = Twips::new(720);
        assert_eq!(i64::from(-t), -720);
    }

    #[test]
    fn pt_add() {
        let a = Pt::new(36.0);
        let b = Pt::new(36.0);
        assert_eq!(f32::from(a + b), 72.0);
    }

    #[test]
    fn pt_sub() {
        let a = Pt::new(72.0);
        let b = Pt::new(36.0);
        assert_eq!(f32::from(a - b), 36.0);
    }

    #[test]
    fn pt_neg() {
        let p = Pt::new(12.0);
        assert_eq!(f32::from(-p), -12.0);
    }

    #[test]
    fn emu_add() {
        let a = Emu::new(914400);
        let b = Emu::new(914400);
        assert_eq!(i64::from(a + b), 1828800);
    }

    // -- Conversions: OOXML → Pt via From --

    #[test]
    fn twips_to_pt() {
        // 1440 twips = 72 pt (1 inch)
        let pt: Pt = Twips::new(1440).into();
        assert_eq!(f32::from(pt), 72.0);
    }

    #[test]
    fn twips_to_pt_one_twip() {
        let pt: Pt = Twips::new(1).into();
        assert!((f32::from(pt) - 0.05).abs() < 0.001);
    }

    #[test]
    fn twips_to_pt_negative() {
        let pt: Pt = Twips::new(-240).into();
        assert_eq!(f32::from(pt), -12.0);
    }

    #[test]
    fn twips_to_pt_zero() {
        let pt: Pt = Twips::new(0).into();
        assert_eq!(f32::from(pt), 0.0);
    }

    #[test]
    fn half_points_to_pt() {
        // 24 half-points = 12 pt
        let pt: Pt = HalfPoints::new(24).into();
        assert_eq!(f32::from(pt), 12.0);
    }

    #[test]
    fn half_points_to_pt_odd() {
        // 13 half-points = 6.5 pt
        let pt: Pt = HalfPoints::new(13).into();
        assert_eq!(f32::from(pt), 6.5);
    }

    #[test]
    fn eighth_points_to_pt() {
        // 8 eighth-points = 1 pt
        let pt: Pt = EighthPoints::new(8).into();
        assert_eq!(f32::from(pt), 1.0);
    }

    #[test]
    fn eighth_points_to_pt_border_default() {
        // 4 eighth-points = 0.5 pt (default border width)
        let pt: Pt = EighthPoints::new(4).into();
        assert_eq!(f32::from(pt), 0.5);
    }

    #[test]
    fn emu_to_pt() {
        // 914400 EMU = 1 inch = 72 pt
        let pt: Pt = Emu::new(914400).into();
        assert!((f32::from(pt) - 72.0).abs() < 0.01);
    }

    #[test]
    fn emu_to_pt_zero() {
        let pt: Pt = Emu::new(0).into();
        assert_eq!(f32::from(pt), 0.0);
    }

    #[test]
    fn emu_to_pt_negative() {
        let pt: Pt = Emu::new(-914400).into();
        assert!((f32::from(pt) + 72.0).abs() < 0.01);
    }

    // -- Conversions via Pt::from --

    #[test]
    fn pt_from_twips() {
        let pt = Pt::from(Twips::new(240));
        assert_eq!(f32::from(pt), 12.0);
    }

    #[test]
    fn pt_from_half_points() {
        let pt = Pt::from(HalfPoints::new(48));
        assert_eq!(f32::from(pt), 24.0);
    }

    #[test]
    fn pt_from_eighth_points() {
        let pt = Pt::from(EighthPoints::new(16));
        assert_eq!(f32::from(pt), 2.0);
    }

    #[test]
    fn pt_from_emu() {
        let pt = Pt::from(Emu::new(457200));
        assert!((f32::from(pt) - 36.0).abs() < 0.01);
    }

    // -- Equality --

    #[test]
    fn twips_equality() {
        assert_eq!(Twips::new(720), Twips::new(720));
        assert_ne!(Twips::new(720), Twips::new(360));
    }

    #[test]
    fn pt_equality() {
        assert_eq!(Pt::new(12.0), Pt::new(12.0));
        assert_ne!(Pt::new(12.0), Pt::new(11.0));
    }

    // -- is_positive --

    #[test]
    fn positive_integer_dimension() {
        assert!(Twips::new(1).is_positive());
        assert!(EighthPoints::new(4).is_positive());
    }

    #[test]
    fn zero_is_not_positive() {
        assert!(!Twips::new(0).is_positive());
        assert!(!Emu::new(0).is_positive());
    }

    #[test]
    fn negative_is_not_positive() {
        assert!(!Twips::new(-1).is_positive());
        assert!(!HalfPoints::new(-10).is_positive());
    }

    // -- Debug & Display --

    #[test]
    fn twips_debug() {
        assert_eq!(format!("{:?}", Twips::new(1440)), "1440tw");
    }

    #[test]
    fn half_points_debug() {
        assert_eq!(format!("{:?}", HalfPoints::new(24)), "24hp");
    }

    #[test]
    fn eighth_points_debug() {
        assert_eq!(format!("{:?}", EighthPoints::new(8)), "8ep");
    }

    #[test]
    fn emu_debug() {
        assert_eq!(format!("{:?}", Emu::new(914400)), "914400emu");
    }

    #[test]
    fn pt_debug() {
        assert_eq!(format!("{:?}", Pt::new(72.0)), "72pt");
    }

    #[test]
    fn twips_display() {
        assert_eq!(format!("{}", Twips::new(720)), "720tw");
    }

    #[test]
    fn pt_display() {
        assert_eq!(format!("{}", Pt::new(12.5)), "12.5pt");
    }

    // -- Copy & Clone --

    #[test]
    fn dimensions_are_copy() {
        let t = Twips::new(100);
        let t2 = t; // copy
        assert_eq!(t, t2); // original still usable
    }

    #[test]
    fn pt_is_copy() {
        let p = Pt::new(12.0);
        let p2 = p;
        assert_eq!(p, p2);
    }

    // -- Pt scalar arithmetic --

    #[test]
    fn pt_mul_scalar() {
        let p = Pt::new(12.0);
        assert_eq!(f32::from(p * 2.0), 24.0);
    }

    #[test]
    fn pt_scalar_mul() {
        let p = Pt::new(12.0);
        assert_eq!(f32::from(2.0 * p), 24.0);
    }

    #[test]
    fn pt_div_scalar() {
        let p = Pt::new(72.0);
        assert_eq!(f32::from(p / 2.0), 36.0);
    }

    #[test]
    fn pt_div_pt_gives_ratio() {
        let a = Pt::new(36.0);
        let b = Pt::new(72.0);
        assert_eq!(a / b, 0.5);
    }

    // -- Pt ordering --

    #[test]
    fn pt_ordering() {
        assert!(Pt::new(12.0) < Pt::new(13.0));
        assert!(Pt::new(13.0) > Pt::new(12.0));
        assert!(Pt::new(12.0) <= Pt::new(12.0));
        assert!(Pt::new(12.0) >= Pt::new(12.0));
    }
}
