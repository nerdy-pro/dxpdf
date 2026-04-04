//! Shared formatting primitives — borders, shading, tabs, alignment, and enums
//! used across paragraph, table, and run properties.

use crate::model::dimension::{Dimension, EighthPoints, Points, Twips};

use super::color::Color;

// ── Alignment ────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Alignment {
    Start,
    Center,
    End,
    Both,
    Distribute,
    Thai,
}

// ── Number Format ────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NumberFormat {
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

// ── Height Rule ──────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HeightRule {
    Auto,
    Exact,
    AtLeast,
}

// ── Borders ──────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ParagraphBorders {
    pub top: Option<Border>,
    pub bottom: Option<Border>,
    pub left: Option<Border>,
    pub right: Option<Border>,
    pub between: Option<Border>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Border {
    pub style: BorderStyle,
    /// §17.3.4: border width in eighths of a point (ST_EighthPointMeasure).
    pub width: Dimension<EighthPoints>,
    /// §17.3.4: spacing offset (ST_PointMeasure §17.18.68).
    pub space: Dimension<Points>,
    pub color: Color,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderStyle {
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

// ── Shading ──────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Shading {
    pub fill: Color,
    pub pattern: ShadingPattern,
    pub color: Color,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ShadingPattern {
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

// ── Tabs ─────────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TabStop {
    pub position: Dimension<Twips>,
    pub alignment: TabAlignment,
    pub leader: TabLeader,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TabAlignment {
    Left,
    Center,
    Right,
    Decimal,
    Bar,
    Clear,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

// ── Conditional Formatting ───────────────────────────────────────────────────

bitflags::bitflags! {
    /// §17.3.1.8: conditional formatting region flags indicating which table
    /// style regions apply to an element (paragraph, row, or cell).
    ///
    /// The 12 bits correspond to the positional regions defined in ST_CnfType.
    /// The legacy `val` binary string (e.g. `"100000000000"`) maps to these
    /// bits left-to-right: bit 0 = firstRow, …, bit 11 = lastRowLastColumn.
    #[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
    pub struct CnfStyle: u16 {
        const FIRST_ROW              = 1 << 0;
        const LAST_ROW               = 1 << 1;
        const FIRST_COLUMN           = 1 << 2;
        const LAST_COLUMN            = 1 << 3;
        const ODD_V_BAND             = 1 << 4;
        const EVEN_V_BAND            = 1 << 5;
        const ODD_H_BAND             = 1 << 6;
        const EVEN_H_BAND            = 1 << 7;
        const FIRST_ROW_FIRST_COLUMN = 1 << 8;
        const FIRST_ROW_LAST_COLUMN  = 1 << 9;
        const LAST_ROW_FIRST_COLUMN  = 1 << 10;
        const LAST_ROW_LAST_COLUMN   = 1 << 11;
    }
}

impl CnfStyle {
    /// Parse the legacy 12-character `val` binary string (§17.3.1.8).
    ///
    /// Each character position maps to a flag left-to-right: `'1'` sets the
    /// flag, `'0'` or any other character leaves it unset. Characters beyond
    /// position 11 are ignored.
    pub fn from_val_str(s: &str) -> Self {
        let flags = [
            CnfStyle::FIRST_ROW,
            CnfStyle::LAST_ROW,
            CnfStyle::FIRST_COLUMN,
            CnfStyle::LAST_COLUMN,
            CnfStyle::ODD_V_BAND,
            CnfStyle::EVEN_V_BAND,
            CnfStyle::ODD_H_BAND,
            CnfStyle::EVEN_H_BAND,
            CnfStyle::FIRST_ROW_FIRST_COLUMN,
            CnfStyle::FIRST_ROW_LAST_COLUMN,
            CnfStyle::LAST_ROW_FIRST_COLUMN,
            CnfStyle::LAST_ROW_LAST_COLUMN,
        ];
        s.bytes()
            .zip(flags.iter())
            .fold(
                CnfStyle::empty(),
                |acc, (ch, &flag)| {
                    if ch == b'1' {
                        acc | flag
                    } else {
                        acc
                    }
                },
            )
    }
}

// ── Text Alignment ───────────────────────────────────────────────────────────

/// §17.18.91 ST_TextAlignment — vertical alignment of characters on a line.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAlignment {
    Auto,
    Top,
    Center,
    Baseline,
    Bottom,
}

// ── Positioning enums (shared by table and frame) ────────────────────────────

/// §17.18.106 ST_VAnchor — vertical/horizontal anchor for table positioning.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableAnchor {
    Text,
    Margin,
    Page,
}

/// §17.18.108 ST_XAlign — horizontal alignment for floating table.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableXAlign {
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

/// §17.18.109 ST_YAlign — vertical alignment for floating table.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableYAlign {
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
    Inline,
}
