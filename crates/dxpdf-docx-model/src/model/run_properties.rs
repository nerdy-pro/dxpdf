//! Run (character) properties.

use crate::dimension::{Dimension, HalfPoints, Twips};

use super::color::Color;
use super::formatting::{Border, Shading};

/// Run properties — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Debug, PartialEq, Default)]
pub struct RunProperties {
    pub fonts: FontSet,
    pub font_size: Option<Dimension<HalfPoints>>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<UnderlineStyle>,
    pub strike: Option<StrikeStyle>,
    pub color: Option<Color>,
    pub highlight: Option<HighlightColor>,
    pub shading: Option<Shading>,
    pub vertical_align: Option<VerticalAlign>,
    pub spacing: Option<Dimension<Twips>>,
    pub kerning: Option<Dimension<HalfPoints>>,
    pub all_caps: Option<bool>,
    pub small_caps: Option<bool>,
    pub vanish: Option<bool>,
    /// §17.3.2.21: suppress spell/grammar checking for this run.
    pub no_proof: Option<bool>,
    /// §17.3.2.44: hidden when displayed as a web page, visible in print view.
    pub web_hidden: Option<bool>,
    pub rtl: Option<bool>,
    pub emboss: Option<bool>,
    pub imprint: Option<bool>,
    pub outline: Option<bool>,
    pub shadow: Option<bool>,
    /// §17.3.2.19: vertical position offset of text baseline, in half-points.
    /// Positive raises, negative lowers.
    pub position: Option<Dimension<HalfPoints>>,
    /// §17.3.2.20: proofing languages per script category (BCP 47 tags).
    pub lang: Option<Lang>,
    /// §17.3.2.4: border around run content.
    pub border: Option<Border>,
}

/// §17.3.2.20: proofing language specification per script category.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Lang {
    /// Language for Latin text (e.g., "en-US").
    pub val: Option<String>,
    /// Language for East Asian text (e.g., "zh-CN").
    pub east_asia: Option<String>,
    /// Language for complex script text (e.g., "ar-SA").
    pub bidi: Option<String>,
}

/// Font family names for each script category.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct FontSet {
    pub ascii: Option<String>,
    pub high_ansi: Option<String>,
    pub east_asian: Option<String>,
    pub complex_script: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnderlineStyle {
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StrikeStyle {
    None,
    Single,
    Double,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

/// Highlight colors — fixed palette per OOXML spec (ST_HighlightColor).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HighlightColor {
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
