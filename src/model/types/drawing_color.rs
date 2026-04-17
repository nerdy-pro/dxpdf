//! DrawingML color model (§20.1.2.3 EG_ColorChoice).
//!
//! The `DrawingColor` ADT models the six base color representations defined
//! by the spec, each followed by an ordered list of color transforms. All
//! values are stored in their spec-defined units (percent in 1000ths, angle
//! in 60000ths of a degree) — no normalization at parse time.

use crate::model::dimension::{Dimension, SixtieThousandthDeg, ThousandthPercent};

// ── EG_ColorChoice (§20.1.2.3) ──────────────────────────────────────────────

/// A DrawingML color specification.
#[derive(Clone, Debug, PartialEq)]
pub enum DrawingColor {
    /// §20.1.2.3.30 CT_ScRgbColor — linear RGB as percentages.
    ScRgb {
        r: Dimension<ThousandthPercent>,
        g: Dimension<ThousandthPercent>,
        b: Dimension<ThousandthPercent>,
        transforms: Vec<ColorTransform>,
    },
    /// §20.1.2.3.32 CT_SRgbColor — sRGB 6-hex-digit value (`val="RRGGBB"`).
    Srgb {
        rgb: u32,
        transforms: Vec<ColorTransform>,
    },
    /// §20.1.2.3.14 CT_HslColor.
    Hsl {
        hue: Dimension<SixtieThousandthDeg>,
        sat: Dimension<ThousandthPercent>,
        lum: Dimension<ThousandthPercent>,
        transforms: Vec<ColorTransform>,
    },
    /// §20.1.2.3.33 CT_SystemColor — OS color with cached RGB.
    Sys {
        name: SystemColorVal,
        last_clr: Option<u32>,
        transforms: Vec<ColorTransform>,
    },
    /// §20.1.2.3.29 CT_SchemeColor — theme color reference.
    Scheme {
        name: SchemeColorVal,
        transforms: Vec<ColorTransform>,
    },
    /// §20.1.2.3.22 CT_PresetColor — named color from the preset palette.
    Prst {
        name: PresetColorVal,
        transforms: Vec<ColorTransform>,
    },
}

impl DrawingColor {
    /// Returns the transform list for any variant.
    pub fn transforms(&self) -> &[ColorTransform] {
        match self {
            Self::ScRgb { transforms, .. }
            | Self::Srgb { transforms, .. }
            | Self::Hsl { transforms, .. }
            | Self::Sys { transforms, .. }
            | Self::Scheme { transforms, .. }
            | Self::Prst { transforms, .. } => transforms,
        }
    }
}

// ── Color transforms (§20.1.2.3.*) ──────────────────────────────────────────

/// A single color-transform child of a color-choice element.
///
/// Transforms apply in document order. Each variant carries the spec's value
/// type and range exactly; clamping/conversion is the resolver's job.
#[derive(Clone, Debug, PartialEq)]
pub enum ColorTransform {
    // Intensity / lightness
    /// §20.1.2.3.34 tint — lighten toward white by `val` ∈ [0, 100%].
    Tint(Dimension<ThousandthPercent>),
    /// §20.1.2.3.31 shade — darken toward black by `val` ∈ [0, 100%].
    Shade(Dimension<ThousandthPercent>),
    /// §20.1.2.3.7 comp — complementary (180° hue rotation).
    Comp,
    /// §20.1.2.3.17 inv — RGB channel inversion.
    Inv,
    /// §20.1.2.3.8 gray — luminance-preserving desaturation.
    Gray,
    /// §20.1.2.3.9 gamma — sRGB → linear gamma correction.
    Gamma,
    /// §20.1.2.3.18 invGamma — linear → sRGB gamma correction.
    InvGamma,

    // Alpha
    /// §20.1.2.3.1 alpha — absolute alpha, `val` ∈ [0, 100%].
    Alpha(Dimension<ThousandthPercent>),
    /// §20.1.2.3.3 alphaOff — alpha offset, `val` ∈ [-100%, 100%].
    AlphaOff(Dimension<ThousandthPercent>),
    /// §20.1.2.3.2 alphaMod — alpha multiplier, `val` ∈ [0, ∞).
    AlphaMod(Dimension<ThousandthPercent>),

    // Hue / Saturation / Luminance (HSL space)
    /// §20.1.2.3.15 hue — absolute hue, `val` ∈ [0, 360°).
    Hue(Dimension<SixtieThousandthDeg>),
    /// §20.1.2.3.16 hueOff — hue offset (any angle).
    HueOff(Dimension<SixtieThousandthDeg>),
    /// §20.1.2.3.14 hueMod — hue multiplier, `val` ∈ [0, ∞).
    HueMod(Dimension<ThousandthPercent>),
    /// §20.1.2.3.26 sat — absolute saturation.
    Sat(Dimension<ThousandthPercent>),
    /// §20.1.2.3.27 satOff — saturation offset.
    SatOff(Dimension<ThousandthPercent>),
    /// §20.1.2.3.28 satMod — saturation multiplier.
    SatMod(Dimension<ThousandthPercent>),
    /// §20.1.2.3.19 lum — absolute luminance.
    Lum(Dimension<ThousandthPercent>),
    /// §20.1.2.3.20 lumOff — luminance offset.
    LumOff(Dimension<ThousandthPercent>),
    /// §20.1.2.3.21 lumMod — luminance multiplier.
    LumMod(Dimension<ThousandthPercent>),

    // Per-channel (RGB space)
    /// §20.1.2.3.23 red — absolute red channel.
    Red(Dimension<ThousandthPercent>),
    /// §20.1.2.3.24 redOff — red offset.
    RedOff(Dimension<ThousandthPercent>),
    /// §20.1.2.3.25 redMod — red multiplier.
    RedMod(Dimension<ThousandthPercent>),
    /// §20.1.2.3.10 green — absolute green channel.
    Green(Dimension<ThousandthPercent>),
    /// §20.1.2.3.11 greenOff — green offset.
    GreenOff(Dimension<ThousandthPercent>),
    /// §20.1.2.3.12 greenMod — green multiplier.
    GreenMod(Dimension<ThousandthPercent>),
    /// §20.1.2.3.4 blue — absolute blue channel.
    Blue(Dimension<ThousandthPercent>),
    /// §20.1.2.3.5 blueOff — blue offset.
    BlueOff(Dimension<ThousandthPercent>),
    /// §20.1.2.3.6 blueMod — blue multiplier.
    BlueMod(Dimension<ThousandthPercent>),
}

// ── ST_SchemeColorVal (§20.1.10.54) ─────────────────────────────────────────

/// §20.1.10.54 ST_SchemeColorVal — named theme color slot.
///
/// `bg1`/`tx1`/`bg2`/`tx2` and `dk1`/`lt1`/`dk2`/`lt2` refer to the same
/// four core slots but with different conventions:
///  * `dk1`/`lt1`/`dk2`/`lt2` reference the theme scheme directly.
///  * `bg1`/`tx1`/`bg2`/`tx2` reference the scheme via the host's
///    foreground/background convention (§17.3.1.31). In Word, light
///    themes map `tx1 → dk1` and `bg1 → lt1`; the resolver folds this.
///
/// `phClr` is a placeholder for master/layout parts; not meaningful in body.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum SchemeColorVal {
    Bg1,
    Tx1,
    Bg2,
    Tx2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hlink,
    FolHlink,
    PhClr,
    Dk1,
    Lt1,
    Dk2,
    Lt2,
}

// ── ST_SystemColorVal (§20.1.10.57) ─────────────────────────────────────────

/// §20.1.10.57 ST_SystemColorVal — OS-defined color name.
///
/// The cached RGB in `CT_SystemColor/@lastClr` is the runtime-resolved value
/// at last save. Renderers without OS access should prefer `lastClr` when
/// present and fall back to a fixed defaults table.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum SystemColorVal {
    ScrollBar,
    Background,
    ActiveCaption,
    InactiveCaption,
    Menu,
    Window,
    WindowFrame,
    MenuText,
    WindowText,
    CaptionText,
    ActiveBorder,
    InactiveBorder,
    AppWorkspace,
    Highlight,
    HighlightText,
    BtnFace,
    BtnShadow,
    GrayText,
    BtnText,
    InactiveCaptionText,
    BtnHighlight,
    ThreeDDkShadow,
    ThreeDLight,
    InfoText,
    InfoBk,
    HotLight,
    GradientActiveCaption,
    GradientInactiveCaption,
    MenuHighlight,
    MenuBar,
}

impl SystemColorVal {
    /// Default RGB fallback when `lastClr` is absent. Based on the Windows
    /// default theme, since OOXML system colors originate from GDI.
    pub fn default_rgb(self) -> u32 {
        match self {
            Self::ScrollBar => 0xC8C8C8,
            Self::Background => 0x000000,
            Self::ActiveCaption => 0x99B4D1,
            Self::InactiveCaption => 0xBFCDDB,
            Self::Menu => 0xF0F0F0,
            Self::Window => 0xFFFFFF,
            Self::WindowFrame => 0x646464,
            Self::MenuText => 0x000000,
            Self::WindowText => 0x000000,
            Self::CaptionText => 0x000000,
            Self::ActiveBorder => 0xB4B4B4,
            Self::InactiveBorder => 0xF4F7FC,
            Self::AppWorkspace => 0xABABAB,
            Self::Highlight => 0x3399FF,
            Self::HighlightText => 0xFFFFFF,
            Self::BtnFace => 0xF0F0F0,
            Self::BtnShadow => 0xA0A0A0,
            Self::GrayText => 0x6D6D6D,
            Self::BtnText => 0x000000,
            Self::InactiveCaptionText => 0x000000,
            Self::BtnHighlight => 0xFFFFFF,
            Self::ThreeDDkShadow => 0x696969,
            Self::ThreeDLight => 0xE3E3E3,
            Self::InfoText => 0x000000,
            Self::InfoBk => 0xFFFFE1,
            Self::HotLight => 0x0066CC,
            Self::GradientActiveCaption => 0xB9D1EA,
            Self::GradientInactiveCaption => 0xD7E4F2,
            Self::MenuHighlight => 0x3399FF,
            Self::MenuBar => 0xF0F0F0,
        }
    }
}

// ── ST_PresetColorVal (§20.1.10.47) ─────────────────────────────────────────

/// §20.1.10.47 ST_PresetColorVal — named preset color.
///
/// A fixed palette of 140 named colors derived from X11 / CSS color names.
/// The spec specifies the RGB value for each; see `rgb()`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum PresetColorVal {
    AliceBlue,
    AntiqueWhite,
    Aqua,
    Aquamarine,
    Azure,
    Beige,
    Bisque,
    Black,
    BlanchedAlmond,
    Blue,
    BlueViolet,
    Brown,
    BurlyWood,
    CadetBlue,
    Chartreuse,
    Chocolate,
    Coral,
    CornflowerBlue,
    Cornsilk,
    Crimson,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGoldenrod,
    DarkGray,
    DarkGreen,
    DarkGrey,
    DarkKhaki,
    DarkMagenta,
    DarkOliveGreen,
    DarkOrange,
    DarkOrchid,
    DarkRed,
    DarkSalmon,
    DarkSeaGreen,
    DarkSlateBlue,
    DarkSlateGray,
    DarkSlateGrey,
    DarkTurquoise,
    DarkViolet,
    DeepPink,
    DeepSkyBlue,
    DimGray,
    DimGrey,
    DkBlue,
    DkCyan,
    DkGoldenrod,
    DkGray,
    DkGreen,
    DkGrey,
    DkKhaki,
    DkMagenta,
    DkOliveGreen,
    DkOrange,
    DkOrchid,
    DkRed,
    DkSalmon,
    DkSeaGreen,
    DkSlateBlue,
    DkSlateGray,
    DkSlateGrey,
    DkTurquoise,
    DkViolet,
    DodgerBlue,
    Firebrick,
    FloralWhite,
    ForestGreen,
    Fuchsia,
    Gainsboro,
    GhostWhite,
    Gold,
    Goldenrod,
    Gray,
    Green,
    GreenYellow,
    Grey,
    Honeydew,
    HotPink,
    IndianRed,
    Indigo,
    Ivory,
    Khaki,
    Lavender,
    LavenderBlush,
    LawnGreen,
    LemonChiffon,
    LightBlue,
    LightCoral,
    LightCyan,
    LightGoldenrodYellow,
    LightGray,
    LightGreen,
    LightGrey,
    LightPink,
    LightSalmon,
    LightSeaGreen,
    LightSkyBlue,
    LightSlateGray,
    LightSlateGrey,
    LightSteelBlue,
    LightYellow,
    Lime,
    LimeGreen,
    Linen,
    LtBlue,
    LtCoral,
    LtCyan,
    LtGoldenrodYellow,
    LtGray,
    LtGreen,
    LtGrey,
    LtPink,
    LtSalmon,
    LtSeaGreen,
    LtSkyBlue,
    LtSlateGray,
    LtSlateGrey,
    LtSteelBlue,
    LtYellow,
    Magenta,
    Maroon,
    MedAquamarine,
    MedBlue,
    MedOrchid,
    MedPurple,
    MedSeaGreen,
    MedSlateBlue,
    MedSpringGreen,
    MedTurquoise,
    MedVioletRed,
    MediumAquamarine,
    MediumBlue,
    MediumOrchid,
    MediumPurple,
    MediumSeaGreen,
    MediumSlateBlue,
    MediumSpringGreen,
    MediumTurquoise,
    MediumVioletRed,
    MidnightBlue,
    MintCream,
    MistyRose,
    Moccasin,
    NavajoWhite,
    Navy,
    OldLace,
    Olive,
    OliveDrab,
    Orange,
    OrangeRed,
    Orchid,
    PaleGoldenrod,
    PaleGreen,
    PaleTurquoise,
    PaleVioletRed,
    PapayaWhip,
    PeachPuff,
    Peru,
    Pink,
    Plum,
    PowderBlue,
    Purple,
    Red,
    RosyBrown,
    RoyalBlue,
    SaddleBrown,
    Salmon,
    SandyBrown,
    SeaGreen,
    SeaShell,
    Sienna,
    Silver,
    SkyBlue,
    SlateBlue,
    SlateGray,
    SlateGrey,
    Snow,
    SpringGreen,
    SteelBlue,
    Tan,
    Teal,
    Thistle,
    Tomato,
    Turquoise,
    Violet,
    Wheat,
    White,
    WhiteSmoke,
    Yellow,
    YellowGreen,
}

impl PresetColorVal {
    /// §20.1.10.47 spec-defined RGB for each preset color.
    ///
    /// The `dk*` / `lt*` / `med*` variants are spec aliases for their
    /// `dark*` / `light*` / `medium*` counterparts and share RGB values.
    pub fn rgb(self) -> u32 {
        match self {
            Self::AliceBlue => 0xF0F8FF,
            Self::AntiqueWhite => 0xFAEBD7,
            Self::Aqua => 0x00FFFF,
            Self::Aquamarine => 0x7FFFD4,
            Self::Azure => 0xF0FFFF,
            Self::Beige => 0xF5F5DC,
            Self::Bisque => 0xFFE4C4,
            Self::Black => 0x000000,
            Self::BlanchedAlmond => 0xFFEBCD,
            Self::Blue => 0x0000FF,
            Self::BlueViolet => 0x8A2BE2,
            Self::Brown => 0xA52A2A,
            Self::BurlyWood => 0xDEB887,
            Self::CadetBlue => 0x5F9EA0,
            Self::Chartreuse => 0x7FFF00,
            Self::Chocolate => 0xD2691E,
            Self::Coral => 0xFF7F50,
            Self::CornflowerBlue => 0x6495ED,
            Self::Cornsilk => 0xFFF8DC,
            Self::Crimson => 0xDC143C,
            Self::Cyan => 0x00FFFF,
            Self::DarkBlue | Self::DkBlue => 0x00008B,
            Self::DarkCyan | Self::DkCyan => 0x008B8B,
            Self::DarkGoldenrod | Self::DkGoldenrod => 0xB8860B,
            Self::DarkGray | Self::DkGray | Self::DarkGrey | Self::DkGrey => 0xA9A9A9,
            Self::DarkGreen | Self::DkGreen => 0x006400,
            Self::DarkKhaki | Self::DkKhaki => 0xBDB76B,
            Self::DarkMagenta | Self::DkMagenta => 0x8B008B,
            Self::DarkOliveGreen | Self::DkOliveGreen => 0x556B2F,
            Self::DarkOrange | Self::DkOrange => 0xFF8C00,
            Self::DarkOrchid | Self::DkOrchid => 0x9932CC,
            Self::DarkRed | Self::DkRed => 0x8B0000,
            Self::DarkSalmon | Self::DkSalmon => 0xE9967A,
            Self::DarkSeaGreen | Self::DkSeaGreen => 0x8FBC8F,
            Self::DarkSlateBlue | Self::DkSlateBlue => 0x483D8B,
            Self::DarkSlateGray | Self::DkSlateGray | Self::DarkSlateGrey | Self::DkSlateGrey => {
                0x2F4F4F
            }
            Self::DarkTurquoise | Self::DkTurquoise => 0x00CED1,
            Self::DarkViolet | Self::DkViolet => 0x9400D3,
            Self::DeepPink => 0xFF1493,
            Self::DeepSkyBlue => 0x00BFFF,
            Self::DimGray | Self::DimGrey => 0x696969,
            Self::DodgerBlue => 0x1E90FF,
            Self::Firebrick => 0xB22222,
            Self::FloralWhite => 0xFFFAF0,
            Self::ForestGreen => 0x228B22,
            Self::Fuchsia => 0xFF00FF,
            Self::Gainsboro => 0xDCDCDC,
            Self::GhostWhite => 0xF8F8FF,
            Self::Gold => 0xFFD700,
            Self::Goldenrod => 0xDAA520,
            Self::Gray | Self::Grey => 0x808080,
            Self::Green => 0x008000,
            Self::GreenYellow => 0xADFF2F,
            Self::Honeydew => 0xF0FFF0,
            Self::HotPink => 0xFF69B4,
            Self::IndianRed => 0xCD5C5C,
            Self::Indigo => 0x4B0082,
            Self::Ivory => 0xFFFFF0,
            Self::Khaki => 0xF0E68C,
            Self::Lavender => 0xE6E6FA,
            Self::LavenderBlush => 0xFFF0F5,
            Self::LawnGreen => 0x7CFC00,
            Self::LemonChiffon => 0xFFFACD,
            Self::LightBlue | Self::LtBlue => 0xADD8E6,
            Self::LightCoral | Self::LtCoral => 0xF08080,
            Self::LightCyan | Self::LtCyan => 0xE0FFFF,
            Self::LightGoldenrodYellow | Self::LtGoldenrodYellow => 0xFAFAD2,
            Self::LightGray | Self::LtGray | Self::LightGrey | Self::LtGrey => 0xD3D3D3,
            Self::LightGreen | Self::LtGreen => 0x90EE90,
            Self::LightPink | Self::LtPink => 0xFFB6C1,
            Self::LightSalmon | Self::LtSalmon => 0xFFA07A,
            Self::LightSeaGreen | Self::LtSeaGreen => 0x20B2AA,
            Self::LightSkyBlue | Self::LtSkyBlue => 0x87CEFA,
            Self::LightSlateGray | Self::LtSlateGray | Self::LightSlateGrey | Self::LtSlateGrey => {
                0x778899
            }
            Self::LightSteelBlue | Self::LtSteelBlue => 0xB0C4DE,
            Self::LightYellow | Self::LtYellow => 0xFFFFE0,
            Self::Lime => 0x00FF00,
            Self::LimeGreen => 0x32CD32,
            Self::Linen => 0xFAF0E6,
            Self::Magenta => 0xFF00FF,
            Self::Maroon => 0x800000,
            Self::MedAquamarine | Self::MediumAquamarine => 0x66CDAA,
            Self::MedBlue | Self::MediumBlue => 0x0000CD,
            Self::MedOrchid | Self::MediumOrchid => 0xBA55D3,
            Self::MedPurple | Self::MediumPurple => 0x9370DB,
            Self::MedSeaGreen | Self::MediumSeaGreen => 0x3CB371,
            Self::MedSlateBlue | Self::MediumSlateBlue => 0x7B68EE,
            Self::MedSpringGreen | Self::MediumSpringGreen => 0x00FA9A,
            Self::MedTurquoise | Self::MediumTurquoise => 0x48D1CC,
            Self::MedVioletRed | Self::MediumVioletRed => 0xC71585,
            Self::MidnightBlue => 0x191970,
            Self::MintCream => 0xF5FFFA,
            Self::MistyRose => 0xFFE4E1,
            Self::Moccasin => 0xFFE4B5,
            Self::NavajoWhite => 0xFFDEAD,
            Self::Navy => 0x000080,
            Self::OldLace => 0xFDF5E6,
            Self::Olive => 0x808000,
            Self::OliveDrab => 0x6B8E23,
            Self::Orange => 0xFFA500,
            Self::OrangeRed => 0xFF4500,
            Self::Orchid => 0xDA70D6,
            Self::PaleGoldenrod => 0xEEE8AA,
            Self::PaleGreen => 0x98FB98,
            Self::PaleTurquoise => 0xAFEEEE,
            Self::PaleVioletRed => 0xDB7093,
            Self::PapayaWhip => 0xFFEFD5,
            Self::PeachPuff => 0xFFDAB9,
            Self::Peru => 0xCD853F,
            Self::Pink => 0xFFC0CB,
            Self::Plum => 0xDDA0DD,
            Self::PowderBlue => 0xB0E0E6,
            Self::Purple => 0x800080,
            Self::Red => 0xFF0000,
            Self::RosyBrown => 0xBC8F8F,
            Self::RoyalBlue => 0x4169E1,
            Self::SaddleBrown => 0x8B4513,
            Self::Salmon => 0xFA8072,
            Self::SandyBrown => 0xF4A460,
            Self::SeaGreen => 0x2E8B57,
            Self::SeaShell => 0xFFF5EE,
            Self::Sienna => 0xA0522D,
            Self::Silver => 0xC0C0C0,
            Self::SkyBlue => 0x87CEEB,
            Self::SlateBlue => 0x6A5ACD,
            Self::SlateGray | Self::SlateGrey => 0x708090,
            Self::Snow => 0xFFFAFA,
            Self::SpringGreen => 0x00FF7F,
            Self::SteelBlue => 0x4682B4,
            Self::Tan => 0xD2B48C,
            Self::Teal => 0x008080,
            Self::Thistle => 0xD8BFD8,
            Self::Tomato => 0xFF6347,
            Self::Turquoise => 0x40E0D0,
            Self::Violet => 0xEE82EE,
            Self::Wheat => 0xF5DEB3,
            Self::White => 0xFFFFFF,
            Self::WhiteSmoke => 0xF5F5F5,
            Self::Yellow => 0xFFFF00,
            Self::YellowGreen => 0x9ACD32,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn drawing_color_transforms_accessor() {
        let c = DrawingColor::Srgb {
            rgb: 0xAABBCC,
            transforms: vec![ColorTransform::Comp, ColorTransform::Inv],
        };
        assert_eq!(c.transforms().len(), 2);

        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Accent1,
            transforms: vec![],
        };
        assert!(c.transforms().is_empty());
    }

    #[test]
    fn preset_color_aliases_share_rgb() {
        assert_eq!(PresetColorVal::DarkBlue.rgb(), PresetColorVal::DkBlue.rgb());
        assert_eq!(PresetColorVal::Gray.rgb(), PresetColorVal::Grey.rgb());
        assert_eq!(
            PresetColorVal::MediumAquamarine.rgb(),
            PresetColorVal::MedAquamarine.rgb()
        );
        assert_eq!(
            PresetColorVal::LightSlateGray.rgb(),
            PresetColorVal::LtSlateGrey.rgb()
        );
    }

    #[test]
    fn preset_color_black_and_white() {
        assert_eq!(PresetColorVal::Black.rgb(), 0x000000);
        assert_eq!(PresetColorVal::White.rgb(), 0xFFFFFF);
    }

    #[test]
    fn preset_color_spec_values() {
        assert_eq!(PresetColorVal::AliceBlue.rgb(), 0xF0F8FF);
        assert_eq!(PresetColorVal::Red.rgb(), 0xFF0000);
        assert_eq!(PresetColorVal::Chocolate.rgb(), 0xD2691E);
    }

    #[test]
    fn system_color_default_window() {
        assert_eq!(SystemColorVal::Window.default_rgb(), 0xFFFFFF);
        assert_eq!(SystemColorVal::WindowText.default_rgb(), 0x000000);
    }
}
