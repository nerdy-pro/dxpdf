//! Shared formatting primitives — borders, shading, tabs, alignment, and enums
//! used across paragraph, table, and run properties.

use crate::dimension::{Dimension, EighthPoints, Twips};

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
    pub width: Dimension<EighthPoints>,
    pub space: Dimension<Twips>,
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

/// §17.3.1.8: conditional formatting bit flags indicating which table style
/// regions apply to an element (paragraph, row, or cell).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CnfStyle {
    /// Legacy 12-character binary string (e.g., "101000000000").
    pub val: Option<String>,
    pub first_row: Option<bool>,
    pub last_row: Option<bool>,
    pub first_column: Option<bool>,
    pub last_column: Option<bool>,
    pub odd_v_band: Option<bool>,
    pub even_v_band: Option<bool>,
    pub odd_h_band: Option<bool>,
    pub even_h_band: Option<bool>,
    pub first_row_first_column: Option<bool>,
    pub first_row_last_column: Option<bool>,
    pub last_row_first_column: Option<bool>,
    pub last_row_last_column: Option<bool>,
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
