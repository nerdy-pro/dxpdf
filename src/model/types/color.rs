//! Color types.

/// A color value as specified in the XML (§17.18.5 ST_HexColor).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Color {
    /// sRGB color parsed from a 6-digit hex string (0xRRGGBB).
    Rgb(u32),
    /// The special "auto" color — meaning context-dependent.
    Auto,
}

impl Color {
    pub const BLACK: Self = Self::Rgb(0x000000);
    pub const WHITE: Self = Self::Rgb(0xFFFFFF);
}
