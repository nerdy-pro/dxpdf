//! DrawingML color schema (§20.1.2.3 EG_ColorChoice).
//!
//! Each color-choice element (`srgbClr`, `scRgbClr`, `hslClr`, `sysClr`,
//! `schemeClr`, `prstClr`) carries attributes for its base color plus an
//! ordered list of color-transform children (`tint`, `shade`, `alpha`, etc.).
//!
//! During Phase 5A these schemas are additive — consumers (fill, stroke,
//! effect) still use the procedural `drawing/color.rs` parser. When those
//! parsers migrate in Phase 5B+, they will switch to `DrawingColorXml`.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::dimension::{Dimension, SixtieThousandthDeg, ThousandthPercent};
use crate::docx::model::{
    ColorTransform, DrawingColor, PresetColorVal, SchemeColorVal, SystemColorVal,
};
use crate::docx::parse::primitives::colors::RgbHexU32;

// ── Top-level choice ──────────────────────────────────────────────────────

/// §20.1.2.3 EG_ColorChoice — exactly one of six color representations.
#[derive(Debug, Deserialize)]
pub enum DrawingColorXml {
    #[serde(rename = "srgbClr")]
    Srgb(SrgbClrXml),
    #[serde(rename = "scRgbClr")]
    ScRgb(ScRgbClrXml),
    #[serde(rename = "hslClr")]
    Hsl(HslClrXml),
    #[serde(rename = "sysClr")]
    Sys(SysClrXml),
    #[serde(rename = "schemeClr")]
    Scheme(SchemeClrXml),
    #[serde(rename = "prstClr")]
    Prst(PrstClrXml),
}

impl From<DrawingColorXml> for DrawingColor {
    fn from(c: DrawingColorXml) -> Self {
        match c {
            DrawingColorXml::Srgb(x) => Self::Srgb {
                rgb: x.val.0,
                transforms: convert_transforms(x.transforms),
            },
            DrawingColorXml::ScRgb(x) => Self::ScRgb {
                r: x.r,
                g: x.g,
                b: x.b,
                transforms: convert_transforms(x.transforms),
            },
            DrawingColorXml::Hsl(x) => Self::Hsl {
                hue: x.hue,
                sat: x.sat,
                lum: x.lum,
                transforms: convert_transforms(x.transforms),
            },
            DrawingColorXml::Sys(x) => Self::Sys {
                name: x.val.into(),
                last_clr: x.last_clr.map(|c| c.0),
                transforms: convert_transforms(x.transforms),
            },
            DrawingColorXml::Scheme(x) => Self::Scheme {
                name: x.val.into(),
                transforms: convert_transforms(x.transforms),
            },
            DrawingColorXml::Prst(x) => Self::Prst {
                name: x.val.into(),
                transforms: convert_transforms(x.transforms),
            },
        }
    }
}

// ── Base color variants ───────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct SrgbClrXml {
    /// §20.1.10.41 ST_HexColorRGB — 6 hex digits.
    #[serde(rename = "@val")]
    pub val: RgbHexU32,
    #[serde(rename = "$value", default)]
    pub transforms: Vec<ColorTransformXml>,
}

#[derive(Debug, Deserialize)]
pub struct ScRgbClrXml {
    #[serde(rename = "@r")]
    pub r: Dimension<ThousandthPercent>,
    #[serde(rename = "@g")]
    pub g: Dimension<ThousandthPercent>,
    #[serde(rename = "@b")]
    pub b: Dimension<ThousandthPercent>,
    #[serde(rename = "$value", default)]
    pub transforms: Vec<ColorTransformXml>,
}

#[derive(Debug, Deserialize)]
pub struct HslClrXml {
    #[serde(rename = "@hue")]
    pub hue: Dimension<SixtieThousandthDeg>,
    #[serde(rename = "@sat")]
    pub sat: Dimension<ThousandthPercent>,
    #[serde(rename = "@lum")]
    pub lum: Dimension<ThousandthPercent>,
    #[serde(rename = "$value", default)]
    pub transforms: Vec<ColorTransformXml>,
}

#[derive(Debug, Deserialize)]
pub struct SysClrXml {
    #[serde(rename = "@val")]
    pub val: StSystemColorVal,
    #[serde(rename = "@lastClr", default)]
    pub last_clr: Option<RgbHexU32>,
    #[serde(rename = "$value", default)]
    pub transforms: Vec<ColorTransformXml>,
}

#[derive(Debug, Deserialize)]
pub struct SchemeClrXml {
    #[serde(rename = "@val")]
    pub val: StSchemeColorVal,
    #[serde(rename = "$value", default)]
    pub transforms: Vec<ColorTransformXml>,
}

#[derive(Debug, Deserialize)]
pub struct PrstClrXml {
    #[serde(rename = "@val")]
    pub val: StPresetColorVal,
    #[serde(rename = "$value", default)]
    pub transforms: Vec<ColorTransformXml>,
}

// ── Color transforms (§20.1.2.3.*) ────────────────────────────────────────

/// Ordered list of transform children; each element is its own variant.
#[derive(Debug, Deserialize)]
pub enum ColorTransformXml {
    // Parameterless transforms — deserialize from empty elements.
    #[serde(rename = "comp")]
    Comp,
    #[serde(rename = "inv")]
    Inv,
    #[serde(rename = "gray")]
    Gray,
    #[serde(rename = "gamma")]
    Gamma,
    #[serde(rename = "invGamma")]
    InvGamma,

    // Percentage-valued (thousandth-percent)
    #[serde(rename = "tint")]
    Tint(PctVal),
    #[serde(rename = "shade")]
    Shade(PctVal),
    #[serde(rename = "alpha")]
    Alpha(PctVal),
    #[serde(rename = "alphaOff")]
    AlphaOff(PctVal),
    #[serde(rename = "alphaMod")]
    AlphaMod(PctVal),
    #[serde(rename = "hueMod")]
    HueMod(PctVal),
    #[serde(rename = "sat")]
    Sat(PctVal),
    #[serde(rename = "satOff")]
    SatOff(PctVal),
    #[serde(rename = "satMod")]
    SatMod(PctVal),
    #[serde(rename = "lum")]
    Lum(PctVal),
    #[serde(rename = "lumOff")]
    LumOff(PctVal),
    #[serde(rename = "lumMod")]
    LumMod(PctVal),
    #[serde(rename = "red")]
    Red(PctVal),
    #[serde(rename = "redOff")]
    RedOff(PctVal),
    #[serde(rename = "redMod")]
    RedMod(PctVal),
    #[serde(rename = "green")]
    Green(PctVal),
    #[serde(rename = "greenOff")]
    GreenOff(PctVal),
    #[serde(rename = "greenMod")]
    GreenMod(PctVal),
    #[serde(rename = "blue")]
    Blue(PctVal),
    #[serde(rename = "blueOff")]
    BlueOff(PctVal),
    #[serde(rename = "blueMod")]
    BlueMod(PctVal),

    // Angle-valued (60000ths of a degree)
    #[serde(rename = "hue")]
    Hue(AngleVal),
    #[serde(rename = "hueOff")]
    HueOff(AngleVal),
}

#[derive(Debug, Deserialize)]
pub struct PctVal {
    #[serde(rename = "@val")]
    pub val: Dimension<ThousandthPercent>,
}

#[derive(Debug, Deserialize)]
pub struct AngleVal {
    #[serde(rename = "@val")]
    pub val: Dimension<SixtieThousandthDeg>,
}

fn convert_transforms(xs: Vec<ColorTransformXml>) -> Vec<ColorTransform> {
    xs.into_iter().map(Into::into).collect()
}

impl From<ColorTransformXml> for ColorTransform {
    fn from(c: ColorTransformXml) -> Self {
        use ColorTransformXml as X;
        match c {
            X::Comp => Self::Comp,
            X::Inv => Self::Inv,
            X::Gray => Self::Gray,
            X::Gamma => Self::Gamma,
            X::InvGamma => Self::InvGamma,
            X::Tint(v) => Self::Tint(v.val),
            X::Shade(v) => Self::Shade(v.val),
            X::Alpha(v) => Self::Alpha(v.val),
            X::AlphaOff(v) => Self::AlphaOff(v.val),
            X::AlphaMod(v) => Self::AlphaMod(v.val),
            X::HueMod(v) => Self::HueMod(v.val),
            X::Sat(v) => Self::Sat(v.val),
            X::SatOff(v) => Self::SatOff(v.val),
            X::SatMod(v) => Self::SatMod(v.val),
            X::Lum(v) => Self::Lum(v.val),
            X::LumOff(v) => Self::LumOff(v.val),
            X::LumMod(v) => Self::LumMod(v.val),
            X::Red(v) => Self::Red(v.val),
            X::RedOff(v) => Self::RedOff(v.val),
            X::RedMod(v) => Self::RedMod(v.val),
            X::Green(v) => Self::Green(v.val),
            X::GreenOff(v) => Self::GreenOff(v.val),
            X::GreenMod(v) => Self::GreenMod(v.val),
            X::Blue(v) => Self::Blue(v.val),
            X::BlueOff(v) => Self::BlueOff(v.val),
            X::BlueMod(v) => Self::BlueMod(v.val),
            X::Hue(v) => Self::Hue(v.val),
            X::HueOff(v) => Self::HueOff(v.val),
        }
    }
}

// ── ST enums ──────────────────────────────────────────────────────────────

/// §20.1.10.54 ST_SchemeColorVal — theme color slot.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StSchemeColorVal {
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

impl From<StSchemeColorVal> for SchemeColorVal {
    fn from(s: StSchemeColorVal) -> Self {
        match s {
            StSchemeColorVal::Bg1 => Self::Bg1,
            StSchemeColorVal::Tx1 => Self::Tx1,
            StSchemeColorVal::Bg2 => Self::Bg2,
            StSchemeColorVal::Tx2 => Self::Tx2,
            StSchemeColorVal::Accent1 => Self::Accent1,
            StSchemeColorVal::Accent2 => Self::Accent2,
            StSchemeColorVal::Accent3 => Self::Accent3,
            StSchemeColorVal::Accent4 => Self::Accent4,
            StSchemeColorVal::Accent5 => Self::Accent5,
            StSchemeColorVal::Accent6 => Self::Accent6,
            StSchemeColorVal::Hlink => Self::Hlink,
            StSchemeColorVal::FolHlink => Self::FolHlink,
            StSchemeColorVal::PhClr => Self::PhClr,
            StSchemeColorVal::Dk1 => Self::Dk1,
            StSchemeColorVal::Lt1 => Self::Lt1,
            StSchemeColorVal::Dk2 => Self::Dk2,
            StSchemeColorVal::Lt2 => Self::Lt2,
        }
    }
}

/// §20.1.10.57 ST_SystemColorVal. `3d*` variants need explicit rename
/// because `rename_all="camelCase"` would produce `threeDDkShadow`, not the
/// spec value `3dDkShadow`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StSystemColorVal {
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
    #[serde(rename = "3dDkShadow")]
    ThreeDDkShadow,
    #[serde(rename = "3dLight")]
    ThreeDLight,
    InfoText,
    InfoBk,
    HotLight,
    GradientActiveCaption,
    GradientInactiveCaption,
    MenuHighlight,
    MenuBar,
}

impl From<StSystemColorVal> for SystemColorVal {
    fn from(s: StSystemColorVal) -> Self {
        use StSystemColorVal as X;
        match s {
            X::ScrollBar => Self::ScrollBar,
            X::Background => Self::Background,
            X::ActiveCaption => Self::ActiveCaption,
            X::InactiveCaption => Self::InactiveCaption,
            X::Menu => Self::Menu,
            X::Window => Self::Window,
            X::WindowFrame => Self::WindowFrame,
            X::MenuText => Self::MenuText,
            X::WindowText => Self::WindowText,
            X::CaptionText => Self::CaptionText,
            X::ActiveBorder => Self::ActiveBorder,
            X::InactiveBorder => Self::InactiveBorder,
            X::AppWorkspace => Self::AppWorkspace,
            X::Highlight => Self::Highlight,
            X::HighlightText => Self::HighlightText,
            X::BtnFace => Self::BtnFace,
            X::BtnShadow => Self::BtnShadow,
            X::GrayText => Self::GrayText,
            X::BtnText => Self::BtnText,
            X::InactiveCaptionText => Self::InactiveCaptionText,
            X::BtnHighlight => Self::BtnHighlight,
            X::ThreeDDkShadow => Self::ThreeDDkShadow,
            X::ThreeDLight => Self::ThreeDLight,
            X::InfoText => Self::InfoText,
            X::InfoBk => Self::InfoBk,
            X::HotLight => Self::HotLight,
            X::GradientActiveCaption => Self::GradientActiveCaption,
            X::GradientInactiveCaption => Self::GradientInactiveCaption,
            X::MenuHighlight => Self::MenuHighlight,
            X::MenuBar => Self::MenuBar,
        }
    }
}

/// §20.1.10.47 ST_PresetColorVal — 140+ named colors.
///
/// All variants match `rename_all="camelCase"` since the OOXML spec uses
/// camelCase consistently for this palette.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPresetColorVal {
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

impl From<StPresetColorVal> for PresetColorVal {
    fn from(s: StPresetColorVal) -> Self {
        use StPresetColorVal as X;
        match s {
            X::AliceBlue => Self::AliceBlue,
            X::AntiqueWhite => Self::AntiqueWhite,
            X::Aqua => Self::Aqua,
            X::Aquamarine => Self::Aquamarine,
            X::Azure => Self::Azure,
            X::Beige => Self::Beige,
            X::Bisque => Self::Bisque,
            X::Black => Self::Black,
            X::BlanchedAlmond => Self::BlanchedAlmond,
            X::Blue => Self::Blue,
            X::BlueViolet => Self::BlueViolet,
            X::Brown => Self::Brown,
            X::BurlyWood => Self::BurlyWood,
            X::CadetBlue => Self::CadetBlue,
            X::Chartreuse => Self::Chartreuse,
            X::Chocolate => Self::Chocolate,
            X::Coral => Self::Coral,
            X::CornflowerBlue => Self::CornflowerBlue,
            X::Cornsilk => Self::Cornsilk,
            X::Crimson => Self::Crimson,
            X::Cyan => Self::Cyan,
            X::DarkBlue => Self::DarkBlue,
            X::DarkCyan => Self::DarkCyan,
            X::DarkGoldenrod => Self::DarkGoldenrod,
            X::DarkGray => Self::DarkGray,
            X::DarkGreen => Self::DarkGreen,
            X::DarkGrey => Self::DarkGrey,
            X::DarkKhaki => Self::DarkKhaki,
            X::DarkMagenta => Self::DarkMagenta,
            X::DarkOliveGreen => Self::DarkOliveGreen,
            X::DarkOrange => Self::DarkOrange,
            X::DarkOrchid => Self::DarkOrchid,
            X::DarkRed => Self::DarkRed,
            X::DarkSalmon => Self::DarkSalmon,
            X::DarkSeaGreen => Self::DarkSeaGreen,
            X::DarkSlateBlue => Self::DarkSlateBlue,
            X::DarkSlateGray => Self::DarkSlateGray,
            X::DarkSlateGrey => Self::DarkSlateGrey,
            X::DarkTurquoise => Self::DarkTurquoise,
            X::DarkViolet => Self::DarkViolet,
            X::DeepPink => Self::DeepPink,
            X::DeepSkyBlue => Self::DeepSkyBlue,
            X::DimGray => Self::DimGray,
            X::DimGrey => Self::DimGrey,
            X::DkBlue => Self::DkBlue,
            X::DkCyan => Self::DkCyan,
            X::DkGoldenrod => Self::DkGoldenrod,
            X::DkGray => Self::DkGray,
            X::DkGreen => Self::DkGreen,
            X::DkGrey => Self::DkGrey,
            X::DkKhaki => Self::DkKhaki,
            X::DkMagenta => Self::DkMagenta,
            X::DkOliveGreen => Self::DkOliveGreen,
            X::DkOrange => Self::DkOrange,
            X::DkOrchid => Self::DkOrchid,
            X::DkRed => Self::DkRed,
            X::DkSalmon => Self::DkSalmon,
            X::DkSeaGreen => Self::DkSeaGreen,
            X::DkSlateBlue => Self::DkSlateBlue,
            X::DkSlateGray => Self::DkSlateGray,
            X::DkSlateGrey => Self::DkSlateGrey,
            X::DkTurquoise => Self::DkTurquoise,
            X::DkViolet => Self::DkViolet,
            X::DodgerBlue => Self::DodgerBlue,
            X::Firebrick => Self::Firebrick,
            X::FloralWhite => Self::FloralWhite,
            X::ForestGreen => Self::ForestGreen,
            X::Fuchsia => Self::Fuchsia,
            X::Gainsboro => Self::Gainsboro,
            X::GhostWhite => Self::GhostWhite,
            X::Gold => Self::Gold,
            X::Goldenrod => Self::Goldenrod,
            X::Gray => Self::Gray,
            X::Green => Self::Green,
            X::GreenYellow => Self::GreenYellow,
            X::Grey => Self::Grey,
            X::Honeydew => Self::Honeydew,
            X::HotPink => Self::HotPink,
            X::IndianRed => Self::IndianRed,
            X::Indigo => Self::Indigo,
            X::Ivory => Self::Ivory,
            X::Khaki => Self::Khaki,
            X::Lavender => Self::Lavender,
            X::LavenderBlush => Self::LavenderBlush,
            X::LawnGreen => Self::LawnGreen,
            X::LemonChiffon => Self::LemonChiffon,
            X::LightBlue => Self::LightBlue,
            X::LightCoral => Self::LightCoral,
            X::LightCyan => Self::LightCyan,
            X::LightGoldenrodYellow => Self::LightGoldenrodYellow,
            X::LightGray => Self::LightGray,
            X::LightGreen => Self::LightGreen,
            X::LightGrey => Self::LightGrey,
            X::LightPink => Self::LightPink,
            X::LightSalmon => Self::LightSalmon,
            X::LightSeaGreen => Self::LightSeaGreen,
            X::LightSkyBlue => Self::LightSkyBlue,
            X::LightSlateGray => Self::LightSlateGray,
            X::LightSlateGrey => Self::LightSlateGrey,
            X::LightSteelBlue => Self::LightSteelBlue,
            X::LightYellow => Self::LightYellow,
            X::Lime => Self::Lime,
            X::LimeGreen => Self::LimeGreen,
            X::Linen => Self::Linen,
            X::LtBlue => Self::LtBlue,
            X::LtCoral => Self::LtCoral,
            X::LtCyan => Self::LtCyan,
            X::LtGoldenrodYellow => Self::LtGoldenrodYellow,
            X::LtGray => Self::LtGray,
            X::LtGreen => Self::LtGreen,
            X::LtGrey => Self::LtGrey,
            X::LtPink => Self::LtPink,
            X::LtSalmon => Self::LtSalmon,
            X::LtSeaGreen => Self::LtSeaGreen,
            X::LtSkyBlue => Self::LtSkyBlue,
            X::LtSlateGray => Self::LtSlateGray,
            X::LtSlateGrey => Self::LtSlateGrey,
            X::LtSteelBlue => Self::LtSteelBlue,
            X::LtYellow => Self::LtYellow,
            X::Magenta => Self::Magenta,
            X::Maroon => Self::Maroon,
            X::MedAquamarine => Self::MedAquamarine,
            X::MedBlue => Self::MedBlue,
            X::MedOrchid => Self::MedOrchid,
            X::MedPurple => Self::MedPurple,
            X::MedSeaGreen => Self::MedSeaGreen,
            X::MedSlateBlue => Self::MedSlateBlue,
            X::MedSpringGreen => Self::MedSpringGreen,
            X::MedTurquoise => Self::MedTurquoise,
            X::MedVioletRed => Self::MedVioletRed,
            X::MediumAquamarine => Self::MediumAquamarine,
            X::MediumBlue => Self::MediumBlue,
            X::MediumOrchid => Self::MediumOrchid,
            X::MediumPurple => Self::MediumPurple,
            X::MediumSeaGreen => Self::MediumSeaGreen,
            X::MediumSlateBlue => Self::MediumSlateBlue,
            X::MediumSpringGreen => Self::MediumSpringGreen,
            X::MediumTurquoise => Self::MediumTurquoise,
            X::MediumVioletRed => Self::MediumVioletRed,
            X::MidnightBlue => Self::MidnightBlue,
            X::MintCream => Self::MintCream,
            X::MistyRose => Self::MistyRose,
            X::Moccasin => Self::Moccasin,
            X::NavajoWhite => Self::NavajoWhite,
            X::Navy => Self::Navy,
            X::OldLace => Self::OldLace,
            X::Olive => Self::Olive,
            X::OliveDrab => Self::OliveDrab,
            X::Orange => Self::Orange,
            X::OrangeRed => Self::OrangeRed,
            X::Orchid => Self::Orchid,
            X::PaleGoldenrod => Self::PaleGoldenrod,
            X::PaleGreen => Self::PaleGreen,
            X::PaleTurquoise => Self::PaleTurquoise,
            X::PaleVioletRed => Self::PaleVioletRed,
            X::PapayaWhip => Self::PapayaWhip,
            X::PeachPuff => Self::PeachPuff,
            X::Peru => Self::Peru,
            X::Pink => Self::Pink,
            X::Plum => Self::Plum,
            X::PowderBlue => Self::PowderBlue,
            X::Purple => Self::Purple,
            X::Red => Self::Red,
            X::RosyBrown => Self::RosyBrown,
            X::RoyalBlue => Self::RoyalBlue,
            X::SaddleBrown => Self::SaddleBrown,
            X::Salmon => Self::Salmon,
            X::SandyBrown => Self::SandyBrown,
            X::SeaGreen => Self::SeaGreen,
            X::SeaShell => Self::SeaShell,
            X::Sienna => Self::Sienna,
            X::Silver => Self::Silver,
            X::SkyBlue => Self::SkyBlue,
            X::SlateBlue => Self::SlateBlue,
            X::SlateGray => Self::SlateGray,
            X::SlateGrey => Self::SlateGrey,
            X::Snow => Self::Snow,
            X::SpringGreen => Self::SpringGreen,
            X::SteelBlue => Self::SteelBlue,
            X::Tan => Self::Tan,
            X::Teal => Self::Teal,
            X::Thistle => Self::Thistle,
            X::Tomato => Self::Tomato,
            X::Turquoise => Self::Turquoise,
            X::Violet => Self::Violet,
            X::Wheat => Self::Wheat,
            X::White => Self::White,
            X::WhiteSmoke => Self::WhiteSmoke,
            X::Yellow => Self::Yellow,
            X::YellowGreen => Self::YellowGreen,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> DrawingColor {
        // Wrap in a parent so the enum dispatches on the child's element name.
        let wrapped = format!(r#"<wrap xmlns:a="urn:a">{}</wrap>"#, xml);
        let (_tag, c): (String, DrawingColorXml) = {
            #[derive(Deserialize)]
            struct Wrap {
                #[serde(rename = "$value")]
                c: DrawingColorXml,
            }
            let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
            (String::new(), w.c)
        };
        c.into()
    }

    #[test]
    fn srgb_with_no_transforms() {
        let c = parse(r#"<srgbClr val="FF0000"/>"#);
        assert!(matches!(c, DrawingColor::Srgb { rgb: 0xFF0000, .. }));
        assert!(c.transforms().is_empty());
    }

    #[test]
    fn srgb_with_transforms_preserves_order() {
        let c = parse(
            r#"<srgbClr val="4F81BD">
                <tint val="50000"/>
                <shade val="75000"/>
                <comp/>
            </srgbClr>"#,
        );
        match c {
            DrawingColor::Srgb { rgb, transforms } => {
                assert_eq!(rgb, 0x4F81BD);
                assert!(matches!(transforms[0], ColorTransform::Tint(d) if d.raw() == 50000));
                assert!(matches!(transforms[1], ColorTransform::Shade(d) if d.raw() == 75000));
                assert!(matches!(transforms[2], ColorTransform::Comp));
            }
            other => panic!("expected Srgb, got {other:?}"),
        }
    }

    #[test]
    fn scheme_color_with_lum_mods() {
        let c = parse(
            r#"<schemeClr val="accent1">
                <lumMod val="60000"/>
                <lumOff val="40000"/>
            </schemeClr>"#,
        );
        match c {
            DrawingColor::Scheme { name, transforms } => {
                assert_eq!(name, SchemeColorVal::Accent1);
                assert_eq!(transforms.len(), 2);
            }
            other => panic!("expected Scheme, got {other:?}"),
        }
    }

    #[test]
    fn sys_color_with_last_clr() {
        let c = parse(r#"<sysClr val="windowText" lastClr="000000"/>"#);
        match c {
            DrawingColor::Sys { name, last_clr, .. } => {
                assert_eq!(name, SystemColorVal::WindowText);
                assert_eq!(last_clr, Some(0));
            }
            other => panic!("expected Sys, got {other:?}"),
        }
    }

    #[test]
    fn sys_3d_dk_shadow_spec_rename() {
        let c = parse(r#"<sysClr val="3dDkShadow"/>"#);
        match c {
            DrawingColor::Sys { name, .. } => assert_eq!(name, SystemColorVal::ThreeDDkShadow),
            other => panic!("expected Sys, got {other:?}"),
        }
    }

    #[test]
    fn scrgb_color_components() {
        let c = parse(r#"<scRgbClr r="50000" g="25000" b="10000"/>"#);
        match c {
            DrawingColor::ScRgb { r, g, b, .. } => {
                assert_eq!(r.raw(), 50000);
                assert_eq!(g.raw(), 25000);
                assert_eq!(b.raw(), 10000);
            }
            other => panic!("expected ScRgb, got {other:?}"),
        }
    }

    #[test]
    fn hsl_color_components() {
        let c = parse(r#"<hslClr hue="14400000" sat="50000" lum="50000"/>"#);
        match c {
            DrawingColor::Hsl { hue, sat, lum, .. } => {
                assert_eq!(hue.raw(), 14_400_000);
                assert_eq!(sat.raw(), 50000);
                assert_eq!(lum.raw(), 50000);
            }
            other => panic!("expected Hsl, got {other:?}"),
        }
    }

    #[test]
    fn preset_color_alice_blue() {
        let c = parse(r#"<prstClr val="aliceBlue"/>"#);
        match c {
            DrawingColor::Prst { name, .. } => assert_eq!(name, PresetColorVal::AliceBlue),
            other => panic!("expected Prst, got {other:?}"),
        }
    }

    #[test]
    fn preset_color_dk_alias() {
        let c = parse(r#"<prstClr val="dkBlue"/>"#);
        match c {
            DrawingColor::Prst { name, .. } => assert_eq!(name, PresetColorVal::DkBlue),
            other => panic!("expected Prst, got {other:?}"),
        }
    }

    #[test]
    fn hue_transform_preserves_angle() {
        let c = parse(r#"<srgbClr val="FF0000"><hue val="3600000"/></srgbClr>"#);
        match c {
            DrawingColor::Srgb { transforms, .. } => {
                assert!(matches!(&transforms[0], ColorTransform::Hue(d) if d.raw() == 3_600_000));
            }
            other => panic!("expected Srgb, got {other:?}"),
        }
    }

    #[test]
    fn parameterless_transforms() {
        let c = parse(
            r#"<srgbClr val="FF0000">
                <inv/>
                <gray/>
                <gamma/>
                <invGamma/>
            </srgbClr>"#,
        );
        let transforms = match c {
            DrawingColor::Srgb { transforms, .. } => transforms,
            _ => panic!(),
        };
        assert_eq!(transforms.len(), 4);
        assert!(matches!(transforms[0], ColorTransform::Inv));
        assert!(matches!(transforms[1], ColorTransform::Gray));
        assert!(matches!(transforms[2], ColorTransform::Gamma));
        assert!(matches!(transforms[3], ColorTransform::InvGamma));
    }

    #[test]
    fn unknown_preset_is_strict() {
        #[derive(Deserialize)]
        struct W {
            #[serde(rename = "@val")]
            #[allow(dead_code)]
            val: StPresetColorVal,
        }
        let r: Result<W, _> = quick_xml::de::from_str(r#"<x val="unknownColor"/>"#);
        assert!(r.is_err());
    }
}
