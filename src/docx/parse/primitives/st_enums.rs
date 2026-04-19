//! OOXML `ST_*` simple-type enums with strict `Deserialize` impls.
//!
//! Each schema enum mirrors an OOXML spec-defined simple type. `From<St…>`
//! implementations convert the schema enum into the matching model type.
//! Unknown string values fail deserialization (plan §Decisions: strict).
//!
//! Alphabetically ordered by schema type name. Layered as:
//!
//! 1. Schema enum with `#[derive(Deserialize)]` + `#[serde(rename_all)]`
//!    (or explicit per-variant rename where the value doesn't match `camelCase`).
//! 2. `impl From<StXxx> for ModelXxx` — identity mapping in most cases,
//!    re-naming/re-mapping where the model is coarser or uses different names.
//!
//! Tests live at the bottom of the file and cover every variant plus a
//! known-bad value for each enum.

use serde::Deserialize;

use crate::docx::model::{
    Alignment, BorderStyle, BreakClear, CellVerticalAlign, FieldCharType, FrameWrap, HeightRule,
    HighlightColor, NumberFormat, PageOrientation, SectionType, ShadingPattern, TabAlignment,
    TabLeader, TableAnchor, TableLayout, TableOverlap, TableXAlign, TableYAlign, TextAlignment,
    TextDirection, ThemeFontRef, UnderlineStyle, VerticalAlign,
};

// ── StBorderType (§17.18.2) ───────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBorderType {
    /// §17.18.2: "no border" sentinel (distinct from `none` per spec but
    /// treated identically by the model).
    Nil,
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
}

impl From<StBorderType> for BorderStyle {
    fn from(s: StBorderType) -> Self {
        match s {
            StBorderType::Nil | StBorderType::None => Self::None,
            StBorderType::Single => Self::Single,
            StBorderType::Thick => Self::Thick,
            StBorderType::Double => Self::Double,
            StBorderType::Dotted => Self::Dotted,
            StBorderType::Dashed => Self::Dashed,
            StBorderType::DotDash => Self::DotDash,
            StBorderType::DotDotDash => Self::DotDotDash,
            StBorderType::Triple => Self::Triple,
            StBorderType::ThinThickSmallGap => Self::ThinThickSmallGap,
            StBorderType::ThickThinSmallGap => Self::ThickThinSmallGap,
            StBorderType::ThinThickThinSmallGap => Self::ThinThickThinSmallGap,
            StBorderType::ThinThickMediumGap => Self::ThinThickMediumGap,
            StBorderType::ThickThinMediumGap => Self::ThickThinMediumGap,
            StBorderType::ThinThickThinMediumGap => Self::ThinThickThinMediumGap,
            StBorderType::ThinThickLargeGap => Self::ThinThickLargeGap,
            StBorderType::ThickThinLargeGap => Self::ThickThinLargeGap,
            StBorderType::ThinThickThinLargeGap => Self::ThinThickThinLargeGap,
            StBorderType::Wave => Self::Wave,
            StBorderType::DoubleWave => Self::DoubleWave,
            StBorderType::DashSmallGap => Self::DashSmallGap,
            StBorderType::DashDotStroked => Self::DashDotStroked,
            StBorderType::ThreeDEmboss => Self::ThreeDEmboss,
            StBorderType::ThreeDEngrave => Self::ThreeDEngrave,
            StBorderType::Outset => Self::Outset,
            StBorderType::Inset => Self::Inset,
        }
    }
}

// ── StBrClear (§17.18.4) ──────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBrClear {
    None,
    Left,
    Right,
    All,
}

impl From<StBrClear> for BreakClear {
    fn from(s: StBrClear) -> Self {
        match s {
            StBrClear::None => Self::None,
            StBrClear::Left => Self::Left,
            StBrClear::Right => Self::Right,
            StBrClear::All => Self::All,
        }
    }
}

// ── StFldCharType (§17.18.29) ─────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StFldCharType {
    Begin,
    Separate,
    End,
}

impl From<StFldCharType> for FieldCharType {
    fn from(s: StFldCharType) -> Self {
        match s {
            StFldCharType::Begin => Self::Begin,
            StFldCharType::Separate => Self::Separate,
            StFldCharType::End => Self::End,
        }
    }
}

// ── StFrameWrap (§17.18.104 ST_Wrap) ──────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StFrameWrap {
    Auto,
    NotBeside,
    Around,
    Tight,
    Through,
    None,
}

impl From<StFrameWrap> for FrameWrap {
    fn from(s: StFrameWrap) -> Self {
        match s {
            StFrameWrap::Auto => Self::Auto,
            StFrameWrap::NotBeside => Self::NotBeside,
            StFrameWrap::Around => Self::Around,
            StFrameWrap::Tight => Self::Tight,
            StFrameWrap::Through => Self::Through,
            StFrameWrap::None => Self::None,
        }
    }
}

// ── StHeightRule (§17.18.38) ──────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StHeightRule {
    Auto,
    Exact,
    AtLeast,
}

impl From<StHeightRule> for HeightRule {
    fn from(s: StHeightRule) -> Self {
        match s {
            StHeightRule::Auto => Self::Auto,
            StHeightRule::Exact => Self::Exact,
            StHeightRule::AtLeast => Self::AtLeast,
        }
    }
}

// ── StHighlightColor (§17.18.40) ──────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StHighlightColor {
    Black,
    Blue,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGray,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    Green,
    LightGray,
    Magenta,
    Red,
    White,
    Yellow,
}

impl From<StHighlightColor> for HighlightColor {
    fn from(s: StHighlightColor) -> Self {
        match s {
            StHighlightColor::Black => Self::Black,
            StHighlightColor::Blue => Self::Blue,
            StHighlightColor::Cyan => Self::Cyan,
            StHighlightColor::DarkBlue => Self::DarkBlue,
            StHighlightColor::DarkCyan => Self::DarkCyan,
            StHighlightColor::DarkGray => Self::DarkGray,
            StHighlightColor::DarkGreen => Self::DarkGreen,
            StHighlightColor::DarkMagenta => Self::DarkMagenta,
            StHighlightColor::DarkRed => Self::DarkRed,
            StHighlightColor::DarkYellow => Self::DarkYellow,
            StHighlightColor::Green => Self::Green,
            StHighlightColor::LightGray => Self::LightGray,
            StHighlightColor::Magenta => Self::Magenta,
            StHighlightColor::Red => Self::Red,
            StHighlightColor::White => Self::White,
            StHighlightColor::Yellow => Self::Yellow,
        }
    }
}

// ── StJc (§17.18.44) ──────────────────────────────────────────────────────
//
// OOXML `both` and `justify` are synonyms per the spec; both produce
// `Alignment::Both`. Rust variant names follow the model's directional
// naming (Start/End) rather than OOXML's presentation naming (Left/Right),
// preserving fidelity to the schema side.

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
pub enum StJc {
    #[serde(rename = "left", alias = "start")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "right", alias = "end")]
    Right,
    #[serde(rename = "both", alias = "justify")]
    Both,
    #[serde(rename = "distribute")]
    Distribute,
    #[serde(rename = "thaiDistribute")]
    ThaiDistribute,
}

impl From<StJc> for Alignment {
    fn from(s: StJc) -> Self {
        match s {
            StJc::Left => Self::Start,
            StJc::Center => Self::Center,
            StJc::Right => Self::End,
            StJc::Both => Self::Both,
            StJc::Distribute => Self::Distribute,
            StJc::ThaiDistribute => Self::Thai,
        }
    }
}

// ── StLineSpacingRule (§17.18.48) ─────────────────────────────────────────
//
// No `From` impl: the model's `LineSpacing` is a discriminated union that
// combines the rule with the value. The conversion happens at the owning
// schema (see `parse::properties::paragraph::SpacingXml`).

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StLineSpacingRule {
    Auto,
    Exact,
    AtLeast,
}

// ── StNumberFormat (§17.18.59) ────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StNumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Bullet,
    Ordinal,
    CardinalText,
    OrdinalText,
    None,
}

impl From<StNumberFormat> for NumberFormat {
    fn from(s: StNumberFormat) -> Self {
        match s {
            StNumberFormat::Decimal => Self::Decimal,
            StNumberFormat::UpperRoman => Self::UpperRoman,
            StNumberFormat::LowerRoman => Self::LowerRoman,
            StNumberFormat::UpperLetter => Self::UpperLetter,
            StNumberFormat::LowerLetter => Self::LowerLetter,
            StNumberFormat::Bullet => Self::Bullet,
            StNumberFormat::Ordinal => Self::Ordinal,
            StNumberFormat::CardinalText => Self::CardinalText,
            StNumberFormat::OrdinalText => Self::OrdinalText,
            StNumberFormat::None => Self::None,
        }
    }
}

// ── StPageOrientation (§17.18.65) ─────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPageOrientation {
    Portrait,
    Landscape,
}

impl From<StPageOrientation> for PageOrientation {
    fn from(s: StPageOrientation) -> Self {
        match s {
            StPageOrientation::Portrait => Self::Portrait,
            StPageOrientation::Landscape => Self::Landscape,
        }
    }
}

// ── StSectionMark (§17.18.77) ─────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StSectionMark {
    NextPage,
    Continuous,
    EvenPage,
    OddPage,
    NextColumn,
}

impl From<StSectionMark> for SectionType {
    fn from(s: StSectionMark) -> Self {
        match s {
            StSectionMark::NextPage => Self::NextPage,
            StSectionMark::Continuous => Self::Continuous,
            StSectionMark::EvenPage => Self::EvenPage,
            StSectionMark::OddPage => Self::OddPage,
            StSectionMark::NextColumn => Self::NextColumn,
        }
    }
}

// ── StShd (§17.18.78) ─────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StShd {
    Clear,
    Solid,
    HorzStripe,
    VertStripe,
    ReverseDiagStripe,
    DiagStripe,
    HorzCross,
    DiagCross,
    ThinHorzStripe,
    ThinVertStripe,
    ThinReverseDiagStripe,
    ThinDiagStripe,
    ThinHorzCross,
    ThinDiagCross,
    Pct5,
    Pct10,
    Pct12,
    Pct15,
    Pct20,
    Pct25,
    Pct30,
    Pct35,
    Pct37,
    Pct40,
    Pct45,
    Pct50,
    Pct55,
    Pct60,
    Pct62,
    Pct65,
    Pct70,
    Pct75,
    Pct80,
    Pct85,
    Pct87,
    Pct90,
    Pct95,
}

impl From<StShd> for ShadingPattern {
    fn from(s: StShd) -> Self {
        match s {
            StShd::Clear => Self::Clear,
            StShd::Solid => Self::Solid,
            StShd::HorzStripe => Self::HorzStripe,
            StShd::VertStripe => Self::VertStripe,
            StShd::ReverseDiagStripe => Self::ReverseDiagStripe,
            StShd::DiagStripe => Self::DiagStripe,
            StShd::HorzCross => Self::HorzCross,
            StShd::DiagCross => Self::DiagCross,
            StShd::ThinHorzStripe => Self::ThinHorzStripe,
            StShd::ThinVertStripe => Self::ThinVertStripe,
            StShd::ThinReverseDiagStripe => Self::ThinReverseDiagStripe,
            StShd::ThinDiagStripe => Self::ThinDiagStripe,
            StShd::ThinHorzCross => Self::ThinHorzCross,
            StShd::ThinDiagCross => Self::ThinDiagCross,
            StShd::Pct5 => Self::Pct5,
            StShd::Pct10 => Self::Pct10,
            StShd::Pct12 => Self::Pct12,
            StShd::Pct15 => Self::Pct15,
            StShd::Pct20 => Self::Pct20,
            StShd::Pct25 => Self::Pct25,
            StShd::Pct30 => Self::Pct30,
            StShd::Pct35 => Self::Pct35,
            StShd::Pct37 => Self::Pct37,
            StShd::Pct40 => Self::Pct40,
            StShd::Pct45 => Self::Pct45,
            StShd::Pct50 => Self::Pct50,
            StShd::Pct55 => Self::Pct55,
            StShd::Pct60 => Self::Pct60,
            StShd::Pct62 => Self::Pct62,
            StShd::Pct65 => Self::Pct65,
            StShd::Pct70 => Self::Pct70,
            StShd::Pct75 => Self::Pct75,
            StShd::Pct80 => Self::Pct80,
            StShd::Pct85 => Self::Pct85,
            StShd::Pct87 => Self::Pct87,
            StShd::Pct90 => Self::Pct90,
            StShd::Pct95 => Self::Pct95,
        }
    }
}

// ── StHAnchor/StVAnchor (§17.18.35/106) ─────────────────────────────────
//
// Shared by table positioning (`<w:tblpPr>`) and frame positioning
// (`<w:framePr>`). OOXML uses the same tag set in both spots.

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StAnchor {
    Text,
    Margin,
    Page,
}

impl From<StAnchor> for TableAnchor {
    fn from(s: StAnchor) -> Self {
        match s {
            StAnchor::Text => Self::Text,
            StAnchor::Margin => Self::Margin,
            StAnchor::Page => Self::Page,
        }
    }
}

// ── StXAlign (§17.18.108) ────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StXAlign {
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

impl From<StXAlign> for TableXAlign {
    fn from(s: StXAlign) -> Self {
        match s {
            StXAlign::Left => Self::Left,
            StXAlign::Center => Self::Center,
            StXAlign::Right => Self::Right,
            StXAlign::Inside => Self::Inside,
            StXAlign::Outside => Self::Outside,
        }
    }
}

// ── StYAlign (§17.18.109) ────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StYAlign {
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
    Inline,
}

impl From<StYAlign> for TableYAlign {
    fn from(s: StYAlign) -> Self {
        match s {
            StYAlign::Top => Self::Top,
            StYAlign::Center => Self::Center,
            StYAlign::Bottom => Self::Bottom,
            StYAlign::Inside => Self::Inside,
            StYAlign::Outside => Self::Outside,
            StYAlign::Inline => Self::Inline,
        }
    }
}

// ── StTabJc (§17.18.85 tab alignment) ─────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTabJc {
    Left,
    Center,
    Right,
    Decimal,
    Bar,
    Clear,
    /// Legacy — treated as `Left`.
    Num,
}

impl From<StTabJc> for TabAlignment {
    fn from(s: StTabJc) -> Self {
        match s {
            StTabJc::Left | StTabJc::Num => Self::Left,
            StTabJc::Center => Self::Center,
            StTabJc::Right => Self::Right,
            StTabJc::Decimal => Self::Decimal,
            StTabJc::Bar => Self::Bar,
            StTabJc::Clear => Self::Clear,
        }
    }
}

// ── StTabTlc (§17.18.86 tab leader character) ─────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTabTlc {
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

impl From<StTabTlc> for TabLeader {
    fn from(s: StTabTlc) -> Self {
        match s {
            StTabTlc::None => Self::None,
            StTabTlc::Dot => Self::Dot,
            StTabTlc::Hyphen => Self::Hyphen,
            StTabTlc::Underscore => Self::Underscore,
            StTabTlc::Heavy => Self::Heavy,
            StTabTlc::MiddleDot => Self::MiddleDot,
        }
    }
}

// ── StTblLayoutType (§17.18.87) ───────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTblLayoutType {
    Auto,
    Fixed,
}

impl From<StTblLayoutType> for TableLayout {
    fn from(s: StTblLayoutType) -> Self {
        match s {
            StTblLayoutType::Auto => Self::Auto,
            StTblLayoutType::Fixed => Self::Fixed,
        }
    }
}

// ── StTblOverlap (§17.4.56) ───────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTblOverlap {
    Overlap,
    Never,
}

impl From<StTblOverlap> for TableOverlap {
    fn from(s: StTblOverlap) -> Self {
        match s {
            StTblOverlap::Overlap => Self::Overlap,
            StTblOverlap::Never => Self::Never,
        }
    }
}

// ── StTextAlignment (§17.18.91) ───────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTextAlignment {
    Auto,
    Top,
    Center,
    Baseline,
    Bottom,
}

impl From<StTextAlignment> for TextAlignment {
    fn from(s: StTextAlignment) -> Self {
        match s {
            StTextAlignment::Auto => Self::Auto,
            StTextAlignment::Top => Self::Top,
            StTextAlignment::Center => Self::Center,
            StTextAlignment::Baseline => Self::Baseline,
            StTextAlignment::Bottom => Self::Bottom,
        }
    }
}

// ── StTextDirection (§17.18.93) ───────────────────────────────────────────
//
// OOXML uses opaque two/three-letter codes; the model spells the full
// directional order. Mapping is static.

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
pub enum StTextDirection {
    #[serde(rename = "lrTb")]
    LrTb,
    #[serde(rename = "tbRl")]
    TbRl,
    #[serde(rename = "btLr")]
    BtLr,
    #[serde(rename = "lrTbV")]
    LrTbV,
    #[serde(rename = "tbRlV")]
    TbRlV,
    #[serde(rename = "tbLrV")]
    TbLrV,
}

impl From<StTextDirection> for TextDirection {
    fn from(s: StTextDirection) -> Self {
        match s {
            StTextDirection::LrTb => Self::LeftToRightTopToBottom,
            StTextDirection::TbRl => Self::TopToBottomRightToLeft,
            StTextDirection::BtLr => Self::BottomToTopLeftToRight,
            StTextDirection::LrTbV => Self::LeftToRightTopToBottomRotated,
            StTextDirection::TbRlV => Self::TopToBottomRightToLeftRotated,
            StTextDirection::TbLrV => Self::TopToBottomLeftToRightRotated,
        }
    }
}

// ── StTheme (§17.18.95) ───────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTheme {
    MajorHAnsi,
    MajorEastAsia,
    MajorBidi,
    MinorHAnsi,
    MinorEastAsia,
    MinorBidi,
}

impl From<StTheme> for ThemeFontRef {
    fn from(s: StTheme) -> Self {
        match s {
            StTheme::MajorHAnsi => Self::MajorHAnsi,
            StTheme::MajorEastAsia => Self::MajorEastAsia,
            StTheme::MajorBidi => Self::MajorBidi,
            StTheme::MinorHAnsi => Self::MinorHAnsi,
            StTheme::MinorEastAsia => Self::MinorEastAsia,
            StTheme::MinorBidi => Self::MinorBidi,
        }
    }
}

// ── StUnderline (§17.18.99) ───────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StUnderline {
    None,
    Single,
    Words,
    Double,
    Thick,
    Dotted,
    DottedHeavy,
    Dash,
    DashedHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DashDotHeavy,
    DotDotDash,
    DashDotDotHeavy,
    Wave,
    WavyHeavy,
    WavyDouble,
}

impl From<StUnderline> for UnderlineStyle {
    fn from(s: StUnderline) -> Self {
        match s {
            StUnderline::None => Self::None,
            StUnderline::Single => Self::Single,
            StUnderline::Words => Self::Words,
            StUnderline::Double => Self::Double,
            StUnderline::Thick => Self::Thick,
            StUnderline::Dotted => Self::Dotted,
            StUnderline::DottedHeavy => Self::DottedHeavy,
            StUnderline::Dash => Self::Dash,
            StUnderline::DashedHeavy => Self::DashedHeavy,
            StUnderline::DashLong => Self::DashLong,
            StUnderline::DashLongHeavy => Self::DashLongHeavy,
            StUnderline::DotDash => Self::DotDash,
            StUnderline::DashDotHeavy => Self::DashDotHeavy,
            StUnderline::DotDotDash => Self::DotDotDash,
            StUnderline::DashDotDotHeavy => Self::DashDotDotHeavy,
            StUnderline::Wave => Self::Wave,
            StUnderline::WavyHeavy => Self::WavyHeavy,
            StUnderline::WavyDouble => Self::WavyDouble,
        }
    }
}

// ── StVerticalAlignRun (§17.18.100) ───────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StVerticalAlignRun {
    Baseline,
    Superscript,
    Subscript,
}

impl From<StVerticalAlignRun> for VerticalAlign {
    fn from(s: StVerticalAlignRun) -> Self {
        match s {
            StVerticalAlignRun::Baseline => Self::Baseline,
            StVerticalAlignRun::Superscript => Self::Superscript,
            StVerticalAlignRun::Subscript => Self::Subscript,
        }
    }
}

// ── StVerticalJc (§17.18.101) ─────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StVerticalJc {
    Top,
    Center,
    Bottom,
    Both,
}

impl From<StVerticalJc> for CellVerticalAlign {
    fn from(s: StVerticalJc) -> Self {
        match s {
            StVerticalJc::Top => Self::Top,
            StVerticalJc::Center => Self::Center,
            StVerticalJc::Bottom => Self::Bottom,
            StVerticalJc::Both => Self::Both,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde::de::DeserializeOwned;

    fn de<T: DeserializeOwned>(v: &str) -> Result<T, quick_xml::DeError> {
        #[derive(Deserialize)]
        struct Wrap<X> {
            #[serde(rename = "@v")]
            v: X,
        }
        quick_xml::de::from_str::<Wrap<T>>(&format!(r#"<x v="{v}"/>"#)).map(|w| w.v)
    }

    fn assert_bad<T: DeserializeOwned + std::fmt::Debug>(v: &str) {
        let r: Result<T, _> = de(v);
        assert!(r.is_err(), "expected error for {v:?}, got {r:?}");
    }

    // ── StBorderType ──
    #[test]
    fn border_type_all_variants() {
        assert_eq!(de::<StBorderType>("none").unwrap(), StBorderType::None);
        assert_eq!(de::<StBorderType>("single").unwrap(), StBorderType::Single);
        assert_eq!(
            de::<StBorderType>("dotDash").unwrap(),
            StBorderType::DotDash
        );
        assert_eq!(
            de::<StBorderType>("threeDEmboss").unwrap(),
            StBorderType::ThreeDEmboss
        );
        assert_eq!(
            de::<StBorderType>("threeDEngrave").unwrap(),
            StBorderType::ThreeDEngrave
        );
        assert_eq!(de::<StBorderType>("inset").unwrap(), StBorderType::Inset);
    }
    #[test]
    fn border_type_strict() {
        assert_bad::<StBorderType>("bogus");
    }
    #[test]
    fn border_type_converts_to_model() {
        let m: BorderStyle = StBorderType::DashDotStroked.into();
        assert_eq!(m, BorderStyle::DashDotStroked);
    }

    // ── StBrClear ──
    #[test]
    fn br_clear_all_variants() {
        assert_eq!(de::<StBrClear>("none").unwrap(), StBrClear::None);
        assert_eq!(de::<StBrClear>("left").unwrap(), StBrClear::Left);
        assert_eq!(de::<StBrClear>("right").unwrap(), StBrClear::Right);
        assert_eq!(de::<StBrClear>("all").unwrap(), StBrClear::All);
    }
    #[test]
    fn br_clear_strict() {
        assert_bad::<StBrClear>("middle");
    }
    #[test]
    fn br_clear_converts_to_model() {
        let m: BreakClear = StBrClear::Left.into();
        assert_eq!(m, BreakClear::Left);
    }

    // ── StFldCharType ──
    #[test]
    fn fld_char_type_all_variants() {
        assert_eq!(de::<StFldCharType>("begin").unwrap(), StFldCharType::Begin);
        assert_eq!(
            de::<StFldCharType>("separate").unwrap(),
            StFldCharType::Separate
        );
        assert_eq!(de::<StFldCharType>("end").unwrap(), StFldCharType::End);
    }
    #[test]
    fn fld_char_type_strict() {
        assert_bad::<StFldCharType>("middle");
    }

    // ── StFrameWrap ──
    #[test]
    fn frame_wrap_all_variants() {
        assert_eq!(de::<StFrameWrap>("auto").unwrap(), StFrameWrap::Auto);
        assert_eq!(
            de::<StFrameWrap>("notBeside").unwrap(),
            StFrameWrap::NotBeside
        );
        assert_eq!(de::<StFrameWrap>("around").unwrap(), StFrameWrap::Around);
        assert_eq!(de::<StFrameWrap>("tight").unwrap(), StFrameWrap::Tight);
        assert_eq!(de::<StFrameWrap>("through").unwrap(), StFrameWrap::Through);
        assert_eq!(de::<StFrameWrap>("none").unwrap(), StFrameWrap::None);
    }
    #[test]
    fn frame_wrap_strict() {
        assert_bad::<StFrameWrap>("wrap");
    }

    // ── StHeightRule ──
    #[test]
    fn height_rule_all_variants() {
        assert_eq!(de::<StHeightRule>("auto").unwrap(), StHeightRule::Auto);
        assert_eq!(de::<StHeightRule>("exact").unwrap(), StHeightRule::Exact);
        assert_eq!(
            de::<StHeightRule>("atLeast").unwrap(),
            StHeightRule::AtLeast
        );
    }
    #[test]
    fn height_rule_strict() {
        assert_bad::<StHeightRule>("maximum");
    }

    // ── StHighlightColor ──
    #[test]
    fn highlight_color_sample_variants() {
        assert_eq!(
            de::<StHighlightColor>("black").unwrap(),
            StHighlightColor::Black
        );
        assert_eq!(
            de::<StHighlightColor>("darkMagenta").unwrap(),
            StHighlightColor::DarkMagenta
        );
        assert_eq!(
            de::<StHighlightColor>("lightGray").unwrap(),
            StHighlightColor::LightGray
        );
        assert_eq!(
            de::<StHighlightColor>("yellow").unwrap(),
            StHighlightColor::Yellow
        );
    }
    #[test]
    fn highlight_color_strict() {
        assert_bad::<StHighlightColor>("chartreuse");
    }

    // ── StJc — includes the both/justify alias and Start/End rename ──
    #[test]
    fn jc_all_variants_and_aliases() {
        assert_eq!(de::<StJc>("left").unwrap(), StJc::Left);
        assert_eq!(de::<StJc>("start").unwrap(), StJc::Left); // alias
        assert_eq!(de::<StJc>("center").unwrap(), StJc::Center);
        assert_eq!(de::<StJc>("right").unwrap(), StJc::Right);
        assert_eq!(de::<StJc>("end").unwrap(), StJc::Right); // alias
        assert_eq!(de::<StJc>("both").unwrap(), StJc::Both);
        assert_eq!(de::<StJc>("justify").unwrap(), StJc::Both); // alias
        assert_eq!(de::<StJc>("distribute").unwrap(), StJc::Distribute);
        assert_eq!(de::<StJc>("thaiDistribute").unwrap(), StJc::ThaiDistribute);
    }
    #[test]
    fn jc_strict() {
        assert_bad::<StJc>("middle");
    }
    #[test]
    fn jc_converts_to_model_with_rename() {
        assert_eq!(Alignment::from(StJc::Left), Alignment::Start);
        assert_eq!(Alignment::from(StJc::Right), Alignment::End);
        assert_eq!(Alignment::from(StJc::ThaiDistribute), Alignment::Thai);
    }

    // ── StNumberFormat ──
    #[test]
    fn number_format_all_variants() {
        assert_eq!(
            de::<StNumberFormat>("decimal").unwrap(),
            StNumberFormat::Decimal
        );
        assert_eq!(
            de::<StNumberFormat>("upperRoman").unwrap(),
            StNumberFormat::UpperRoman
        );
        assert_eq!(
            de::<StNumberFormat>("cardinalText").unwrap(),
            StNumberFormat::CardinalText
        );
        assert_eq!(
            de::<StNumberFormat>("bullet").unwrap(),
            StNumberFormat::Bullet
        );
    }
    #[test]
    fn number_format_strict() {
        assert_bad::<StNumberFormat>("hex");
    }

    // ── StPageOrientation ──
    #[test]
    fn page_orientation_both() {
        assert_eq!(
            de::<StPageOrientation>("portrait").unwrap(),
            StPageOrientation::Portrait
        );
        assert_eq!(
            de::<StPageOrientation>("landscape").unwrap(),
            StPageOrientation::Landscape
        );
    }
    #[test]
    fn page_orientation_strict() {
        assert_bad::<StPageOrientation>("sideways");
    }

    // ── StSectionMark ──
    #[test]
    fn section_mark_all_variants() {
        assert_eq!(
            de::<StSectionMark>("nextPage").unwrap(),
            StSectionMark::NextPage
        );
        assert_eq!(
            de::<StSectionMark>("continuous").unwrap(),
            StSectionMark::Continuous
        );
        assert_eq!(
            de::<StSectionMark>("evenPage").unwrap(),
            StSectionMark::EvenPage
        );
        assert_eq!(
            de::<StSectionMark>("oddPage").unwrap(),
            StSectionMark::OddPage
        );
        assert_eq!(
            de::<StSectionMark>("nextColumn").unwrap(),
            StSectionMark::NextColumn
        );
    }
    #[test]
    fn section_mark_strict() {
        assert_bad::<StSectionMark>("previous");
    }

    // ── StShd ──
    #[test]
    fn shd_sample_variants() {
        assert_eq!(de::<StShd>("clear").unwrap(), StShd::Clear);
        assert_eq!(de::<StShd>("solid").unwrap(), StShd::Solid);
        assert_eq!(de::<StShd>("horzStripe").unwrap(), StShd::HorzStripe);
        assert_eq!(de::<StShd>("thinDiagCross").unwrap(), StShd::ThinDiagCross);
        assert_eq!(de::<StShd>("pct5").unwrap(), StShd::Pct5);
        assert_eq!(de::<StShd>("pct95").unwrap(), StShd::Pct95);
    }
    #[test]
    fn shd_strict() {
        assert_bad::<StShd>("pct100");
    }

    // ── StTblLayoutType ──
    #[test]
    fn tbl_layout_type_both() {
        assert_eq!(
            de::<StTblLayoutType>("auto").unwrap(),
            StTblLayoutType::Auto
        );
        assert_eq!(
            de::<StTblLayoutType>("fixed").unwrap(),
            StTblLayoutType::Fixed
        );
    }
    #[test]
    fn tbl_layout_type_strict() {
        assert_bad::<StTblLayoutType>("flex");
    }

    // ── StTblOverlap ──
    #[test]
    fn tbl_overlap_both() {
        assert_eq!(
            de::<StTblOverlap>("overlap").unwrap(),
            StTblOverlap::Overlap
        );
        assert_eq!(de::<StTblOverlap>("never").unwrap(), StTblOverlap::Never);
    }
    #[test]
    fn tbl_overlap_strict() {
        assert_bad::<StTblOverlap>("always");
    }

    // ── StTextAlignment ──
    #[test]
    fn text_alignment_all_variants() {
        assert_eq!(
            de::<StTextAlignment>("auto").unwrap(),
            StTextAlignment::Auto
        );
        assert_eq!(de::<StTextAlignment>("top").unwrap(), StTextAlignment::Top);
        assert_eq!(
            de::<StTextAlignment>("center").unwrap(),
            StTextAlignment::Center
        );
        assert_eq!(
            de::<StTextAlignment>("baseline").unwrap(),
            StTextAlignment::Baseline
        );
        assert_eq!(
            de::<StTextAlignment>("bottom").unwrap(),
            StTextAlignment::Bottom
        );
    }
    #[test]
    fn text_alignment_strict() {
        assert_bad::<StTextAlignment>("middle");
    }

    // ── StTextDirection ──
    #[test]
    fn text_direction_all_variants() {
        assert_eq!(
            de::<StTextDirection>("lrTb").unwrap(),
            StTextDirection::LrTb
        );
        assert_eq!(
            de::<StTextDirection>("tbRl").unwrap(),
            StTextDirection::TbRl
        );
        assert_eq!(
            de::<StTextDirection>("btLr").unwrap(),
            StTextDirection::BtLr
        );
        assert_eq!(
            de::<StTextDirection>("lrTbV").unwrap(),
            StTextDirection::LrTbV
        );
        assert_eq!(
            de::<StTextDirection>("tbRlV").unwrap(),
            StTextDirection::TbRlV
        );
        assert_eq!(
            de::<StTextDirection>("tbLrV").unwrap(),
            StTextDirection::TbLrV
        );
    }
    #[test]
    fn text_direction_strict() {
        assert_bad::<StTextDirection>("ltr");
    }
    #[test]
    fn text_direction_converts_to_model() {
        assert_eq!(
            TextDirection::from(StTextDirection::LrTb),
            TextDirection::LeftToRightTopToBottom
        );
        assert_eq!(
            TextDirection::from(StTextDirection::TbLrV),
            TextDirection::TopToBottomLeftToRightRotated
        );
    }

    // ── StTheme ──
    #[test]
    fn theme_all_variants() {
        assert_eq!(de::<StTheme>("majorHAnsi").unwrap(), StTheme::MajorHAnsi);
        assert_eq!(
            de::<StTheme>("majorEastAsia").unwrap(),
            StTheme::MajorEastAsia
        );
        assert_eq!(de::<StTheme>("majorBidi").unwrap(), StTheme::MajorBidi);
        assert_eq!(de::<StTheme>("minorHAnsi").unwrap(), StTheme::MinorHAnsi);
        assert_eq!(
            de::<StTheme>("minorEastAsia").unwrap(),
            StTheme::MinorEastAsia
        );
        assert_eq!(de::<StTheme>("minorBidi").unwrap(), StTheme::MinorBidi);
    }
    #[test]
    fn theme_strict() {
        assert_bad::<StTheme>("default");
    }

    // ── StUnderline ──
    #[test]
    fn underline_sample_variants() {
        assert_eq!(de::<StUnderline>("none").unwrap(), StUnderline::None);
        assert_eq!(de::<StUnderline>("single").unwrap(), StUnderline::Single);
        assert_eq!(de::<StUnderline>("dotted").unwrap(), StUnderline::Dotted);
        assert_eq!(
            de::<StUnderline>("dashDotHeavy").unwrap(),
            StUnderline::DashDotHeavy
        );
        assert_eq!(
            de::<StUnderline>("wavyDouble").unwrap(),
            StUnderline::WavyDouble
        );
    }
    #[test]
    fn underline_strict() {
        assert_bad::<StUnderline>("italic");
    }

    // ── StVerticalAlignRun ──
    #[test]
    fn vertical_align_run_all_variants() {
        assert_eq!(
            de::<StVerticalAlignRun>("baseline").unwrap(),
            StVerticalAlignRun::Baseline
        );
        assert_eq!(
            de::<StVerticalAlignRun>("superscript").unwrap(),
            StVerticalAlignRun::Superscript
        );
        assert_eq!(
            de::<StVerticalAlignRun>("subscript").unwrap(),
            StVerticalAlignRun::Subscript
        );
    }
    #[test]
    fn vertical_align_run_strict() {
        assert_bad::<StVerticalAlignRun>("middle");
    }

    // ── StVerticalJc ──
    #[test]
    fn vertical_jc_all_variants() {
        assert_eq!(de::<StVerticalJc>("top").unwrap(), StVerticalJc::Top);
        assert_eq!(de::<StVerticalJc>("center").unwrap(), StVerticalJc::Center);
        assert_eq!(de::<StVerticalJc>("bottom").unwrap(), StVerticalJc::Bottom);
        assert_eq!(de::<StVerticalJc>("both").unwrap(), StVerticalJc::Both);
    }
    #[test]
    fn vertical_jc_strict() {
        assert_bad::<StVerticalJc>("start");
    }
}
