use std::fmt;
use std::marker::PhantomData;
use std::ops::{Add, Div, Mul, Neg, Sub};

/// Marker trait for dimension unit types.
pub trait Unit: Copy + Clone + fmt::Debug + PartialEq + Eq {
    const NAME: &'static str;
}

/// A dimension value parameterized by its unit of measurement.
/// Integer storage for lossless OOXML round-tripping.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
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

/// Percentage in 1/1000th of a percent (OOXML ST_DecimalNumberOrPercent).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ThousandthPercent;
impl Unit for ThousandthPercent {
    const NAME: &'static str = "‰%";
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
    pub fn to_fraction(self) -> f64 {
        self.raw as f64 / 100_000.0
    }
}
