use std::fmt;
use std::marker::PhantomData;
use std::ops::{Add, Div, Mul, Neg, Sub};

/// Marker trait for dimension unit types.
pub trait Unit: Copy + Clone + fmt::Debug + PartialEq + Eq {
    const NAME: &'static str;
}

/// A dimension value parameterized by its unit of measurement.
/// Integer storage for lossless OOXML round-tripping.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Dimension<U: Unit> {
    raw: i64,
    _unit: PhantomData<U>,
}

impl<U: Unit> Dimension<U> {
    pub const ZERO: Self = Self {
        raw: 0,
        _unit: PhantomData,
    };

    pub const fn new(raw: i64) -> Self {
        Self {
            raw,
            _unit: PhantomData,
        }
    }

    pub const fn raw(self) -> i64 {
        self.raw
    }
}

impl<U: Unit> fmt::Debug for Dimension<U> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", self.raw, U::NAME)
    }
}

impl<U: Unit> Default for Dimension<U> {
    fn default() -> Self {
        Self::ZERO
    }
}

// Ordering depends only on the stored value, never on the phantom unit, so
// `Dimension<U>` is `Ord` for *every* `U: Unit`. Deriving these would instead
// synthesize a spurious `U: Ord` bound — and no unit marker implements `Ord` —
// leaving the derived impl unusable and forcing callers to compare `.raw()` by
// hand. These manual impls (like the manual `Default`/arithmetic above) drop the
// bound.
impl<U: Unit> PartialOrd for Dimension<U> {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl<U: Unit> Ord for Dimension<U> {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        self.raw.cmp(&other.raw)
    }
}

impl<U: Unit> Add for Dimension<U> {
    type Output = Self;
    fn add(self, rhs: Self) -> Self {
        Self::new(self.raw + rhs.raw)
    }
}

impl<U: Unit> Sub for Dimension<U> {
    type Output = Self;
    fn sub(self, rhs: Self) -> Self {
        Self::new(self.raw - rhs.raw)
    }
}

impl<U: Unit> Mul<i64> for Dimension<U> {
    type Output = Self;
    fn mul(self, rhs: i64) -> Self {
        Self::new(self.raw * rhs)
    }
}

impl<U: Unit> Div<i64> for Dimension<U> {
    type Output = Self;
    fn div(self, rhs: i64) -> Self {
        Self::new(self.raw / rhs)
    }
}

impl<U: Unit> Neg for Dimension<U> {
    type Output = Self;
    fn neg(self) -> Self {
        Self::new(-self.raw)
    }
}

// --- Unit markers ---

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Twips;
impl Unit for Twips {
    const NAME: &'static str = "twip";
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct HalfPoints;
impl Unit for HalfPoints {
    const NAME: &'static str = "hp";
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Emu;
impl Unit for Emu {
    const NAME: &'static str = "emu";
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct EighthPoints;
impl Unit for EighthPoints {
    const NAME: &'static str = "ep";
}

/// 1/4096th of a point — used for document grid character spacing (§17.6.5).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct FractionPoints;
impl Unit for FractionPoints {
    const NAME: &'static str = "fp4096";
}

/// §17.18.68 ST_PointMeasure — whole points (1/72nd of an inch).
/// Used for border spacing (§17.3.4 w:space).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Points;
impl Unit for Points {
    const NAME: &'static str = "pt";
}

/// §20.1.10.68 ST_TextFontSize — hundredths of a point, DrawingML text's
/// own size unit (`a:rPr/@sz`, 1100 = 11pt). A fourth point-fraction unit
/// beside half/eighth/4096th because nothing but the type says which
/// divisor a raw `sz` needs.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CentiPoints;
impl Unit for CentiPoints {
    const NAME: &'static str = "cpt";
}

/// Percentage in 1/1000th of a percent (OOXML ST_DecimalNumberOrPercent).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ThousandthPercent;
impl Unit for ThousandthPercent {
    const NAME: &'static str = "‰%";
}

/// §17.18.90 ST_TblWidth `pct` — percentage in **fiftieths** of a percent, so
/// 5000 is 100%.
///
/// A separate unit from [`ThousandthPercent`] because the two divisors differ
/// by 20× and nothing but the type tells them apart: a `<w:tblW w:type="pct"
/// w:w="5000"/>` read on the thousandth scale is 5%, not 100%, and a table
/// meant to span the text column comes out a twentieth of it. That is exactly
/// the mistake the unit system exists to make unrepresentable.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct FiftiethPercent;
impl Unit for FiftiethPercent {
    const NAME: &'static str = "/50%";
}

/// DrawingML angle in 60,000ths of a degree (§20.1.10.3 ST_Angle).
/// 0 = no rotation, 5400000 = 90° clockwise.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct SixtieThousandthDeg;
impl Unit for SixtieThousandthDeg {
    const NAME: &'static str = "∠60k";
}

// --- Conversions ---

impl Dimension<Twips> {
    /// 1 twip = 1/20 of a point.
    pub fn to_half_points(self) -> Dimension<HalfPoints> {
        Dimension::new(self.raw / 10)
    }

    /// 1 twip = 635 EMU.
    pub fn to_emu(self) -> Dimension<Emu> {
        Dimension::new(self.raw * 635)
    }

    pub fn to_points_f32(self) -> f32 {
        self.raw as f32 / 20.0
    }
}

impl Dimension<HalfPoints> {
    pub fn to_twips(self) -> Dimension<Twips> {
        Dimension::new(self.raw * 10)
    }

    pub fn to_points_f32(self) -> f32 {
        self.raw as f32 / 2.0
    }
}

impl Dimension<Emu> {
    /// 1 EMU = 1/914400 inch = 1/635 twip.
    pub fn to_twips(self) -> Dimension<Twips> {
        Dimension::new(self.raw / 635)
    }

    pub fn to_points_f32(self) -> f32 {
        self.raw as f32 / 12700.0
    }
}

impl Dimension<EighthPoints> {
    pub fn to_half_points(self) -> Dimension<HalfPoints> {
        Dimension::new(self.raw / 4)
    }

    pub fn to_points_f32(self) -> f32 {
        self.raw as f32 / 8.0
    }
}

impl Dimension<ThousandthPercent> {
    /// Returns the percentage as a fraction (e.g., 50000 → 0.5).
    ///
    /// The single home for the ST_Percentage / CT_RelativeRect scale (100% =
    /// 100000); every consumer that needs a `[0, 1]` fraction routes through
    /// here rather than open-coding the divisor.
    pub fn to_fraction(self) -> f32 {
        self.raw as f32 / 100_000.0
    }
}

impl Dimension<FiftiethPercent> {
    /// 100% on the ST_TblWidth `pct` scale — the width of a table that spans
    /// whatever it is measured against.
    pub const FULL: Self = Self::new(5000);

    /// Returns the percentage as a fraction (e.g., 2500 → 0.5).
    ///
    /// The single home for the ST_TblWidth `pct` scale, as
    /// `Dimension::<ThousandthPercent>::to_fraction` is for ST_Percentage.
    pub fn to_fraction(self) -> f32 {
        self.raw as f32 / 5000.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// f32 equality with a small tolerance (clippy forbids bare `==` on floats).
    fn approx(a: f32, b: f32) {
        assert!((a - b).abs() < 1e-6, "{a} != {b}");
    }

    #[test]
    fn twips_to_half_points() {
        // 20 twips = 1 pt = 2 half-points; the conversion divides by 10.
        assert_eq!(Dimension::<Twips>::new(20).to_half_points().raw(), 2);
        assert_eq!(Dimension::<Twips>::new(1440).to_half_points().raw(), 144);
    }

    #[test]
    fn twips_emu_roundtrip() {
        // 1 twip = 635 EMU; 1 inch = 1440 twip = 914400 EMU.
        assert_eq!(Dimension::<Twips>::new(1440).to_emu().raw(), 914_400);
        assert_eq!(Dimension::<Emu>::new(914_400).to_twips().raw(), 1440);
    }

    #[test]
    fn to_points_conversions() {
        approx(Dimension::<Twips>::new(1440).to_points_f32(), 72.0);
        approx(Dimension::<Twips>::new(240).to_points_f32(), 12.0);
        approx(Dimension::<HalfPoints>::new(24).to_points_f32(), 12.0);
        approx(Dimension::<Emu>::new(12_700).to_points_f32(), 1.0);
        approx(Dimension::<EighthPoints>::new(8).to_points_f32(), 1.0);
    }

    #[test]
    fn half_points_to_twips() {
        assert_eq!(Dimension::<HalfPoints>::new(24).to_twips().raw(), 240);
    }

    #[test]
    fn eighth_points_to_half_points() {
        // 8 eighth-points = 1 pt = 2 half-points; divides by 4.
        assert_eq!(Dimension::<EighthPoints>::new(8).to_half_points().raw(), 2);
    }

    #[test]
    fn thousandth_percent_to_fraction() {
        approx(
            Dimension::<ThousandthPercent>::new(50_000).to_fraction(),
            0.5,
        );
        approx(
            Dimension::<ThousandthPercent>::new(100_000).to_fraction(),
            1.0,
        );
        approx(Dimension::<ThousandthPercent>::new(0).to_fraction(), 0.0);
    }

    /// §17.18.90: the `pct` scale is 20× coarser than ST_Percentage's, which is
    /// the whole reason the two are separate units. `5000` on this scale is a
    /// full-width table; on the other it would be 5% of one.
    #[test]
    fn fiftieth_percent_to_fraction() {
        approx(Dimension::<FiftiethPercent>::new(2500).to_fraction(), 0.5);
        approx(Dimension::<FiftiethPercent>::FULL.to_fraction(), 1.0);
        approx(Dimension::<FiftiethPercent>::new(0).to_fraction(), 0.0);
        approx(
            Dimension::<ThousandthPercent>::new(5000).to_fraction(),
            0.05,
        );
    }

    #[test]
    fn arithmetic_ops() {
        let a = Dimension::<Twips>::new(100);
        let b = Dimension::<Twips>::new(30);
        assert_eq!((a + b).raw(), 130);
        assert_eq!((a - b).raw(), 70);
        assert_eq!((a * 3).raw(), 300);
        assert_eq!((a / 4).raw(), 25); // integer division truncates
        assert_eq!((-a).raw(), -100);
    }

    #[test]
    fn zero_and_default() {
        assert_eq!(Dimension::<Twips>::ZERO.raw(), 0);
        assert_eq!(Dimension::<Emu>::default().raw(), 0);
    }

    #[test]
    fn ordering_and_equality() {
        assert!(Dimension::<Twips>::new(10) < Dimension::<Twips>::new(20));
        assert_eq!(Dimension::<Twips>::new(5), Dimension::<Twips>::new(5));
    }
}
