//! Paragraph types — paragraph properties, spacing, indentation, frames.

use crate::dimension::{Dimension, Twips};

use super::formatting::{
    Alignment, CnfStyle, HeightRule, ParagraphBorders, Shading, TabStop, TableAnchor,
    TableXAlign, TableYAlign, TextAlignment,
};
use super::identifiers::{ParagraphRevisionIds, StyleId};
use super::run_properties::RunProperties;

use super::content::Inline;

#[derive(Clone, Debug)]
pub struct Paragraph {
    /// Style ID reference (e.g., "Heading1"). Resolve via `Document.styles`.
    pub style_id: Option<StyleId>,
    pub properties: ParagraphProperties,
    /// Run properties specified on the paragraph mark (w:rPr inside w:pPr).
    pub mark_run_properties: Option<RunProperties>,
    pub content: Vec<Inline>,
    pub rsids: ParagraphRevisionIds,
}

/// Paragraph properties — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Debug, Default)]
pub struct ParagraphProperties {
    pub alignment: Option<Alignment>,
    pub indentation: Option<Indentation>,
    pub spacing: Option<ParagraphSpacing>,
    pub numbering: Option<NumberingReference>,
    pub tabs: Vec<TabStop>,
    pub borders: Option<ParagraphBorders>,
    pub shading: Option<Shading>,
    pub keep_next: Option<bool>,
    pub keep_lines: Option<bool>,
    pub widow_control: Option<bool>,
    pub page_break_before: Option<bool>,
    pub suppress_auto_hyphens: Option<bool>,
    /// §17.3.1.9: suppress spacing when adjacent paragraph has same style.
    pub contextual_spacing: Option<bool>,
    pub bidi: Option<bool>,
    /// §17.3.1.45: allow line breaking between any characters for East Asian text.
    pub word_wrap: Option<bool>,
    pub outline_level: Option<OutlineLevel>,
    /// §17.3.1.39: vertical alignment of text on each line (ST_TextAlignment).
    pub text_alignment: Option<TextAlignment>,
    /// §17.3.1.8: table conditional formatting applied to this paragraph.
    pub cnf_style: Option<CnfStyle>,
    /// §17.3.1.11: text frame (legacy positioned text region).
    pub frame_properties: Option<FrameProperties>,
    /// §17.3.1.2: auto-space East Asian text with Latin text.
    pub auto_space_de: Option<bool>,
    /// §17.3.1.3: auto-space East Asian text with numbers.
    pub auto_space_dn: Option<bool>,
}

/// §17.3.1.11: text frame properties — legacy floating positioned text.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FrameProperties {
    /// §17.18.16 ST_DropCap: drop cap type.
    pub drop_cap: Option<DropCap>,
    /// Number of lines the drop cap spans.
    pub lines: Option<u32>,
    /// Frame width in twips.
    pub width: Option<Dimension<Twips>>,
    /// Frame height in twips.
    pub height: Option<Dimension<Twips>>,
    /// §17.18.37 ST_HeightRule: how to interpret the height value.
    pub height_rule: Option<HeightRule>,
    /// Horizontal distance from surrounding text in twips.
    pub h_space: Option<Dimension<Twips>>,
    /// Vertical distance from surrounding text in twips.
    pub v_space: Option<Dimension<Twips>>,
    /// §17.18.104 ST_Wrap: text wrapping mode.
    pub wrap: Option<FrameWrap>,
    /// §17.18.35 ST_HAnchor: horizontal anchor.
    pub h_anchor: Option<TableAnchor>,
    /// §17.18.106 ST_VAnchor: vertical anchor.
    pub v_anchor: Option<TableAnchor>,
    /// Absolute horizontal position in twips.
    pub x: Option<Dimension<Twips>>,
    /// §17.18.108 ST_XAlign: horizontal alignment.
    pub x_align: Option<TableXAlign>,
    /// Absolute vertical position in twips.
    pub y: Option<Dimension<Twips>>,
    /// §17.18.109 ST_YAlign: vertical alignment.
    pub y_align: Option<TableYAlign>,
}

/// §17.18.16 ST_DropCap
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DropCap {
    None,
    Drop,
    Margin,
}

/// §17.18.104 ST_Wrap — text wrapping for frames.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FrameWrap {
    Auto,
    NotBeside,
    Around,
    Tight,
    Through,
    None,
}

/// Heading outline level (0–8, where 0 = Heading 1 in OOXML).
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct OutlineLevel(u8);

impl OutlineLevel {
    /// Create an outline level. Panics if `level` is 0 or > 9.
    pub fn new(level: u8) -> Self {
        assert!((1..=9).contains(&level), "outline level must be 1..=9");
        Self(level)
    }

    /// Create from OOXML raw value (0-based). Returns None if > 8.
    pub fn from_ooxml(val: u8) -> Option<Self> {
        if val <= 8 {
            Some(Self(val + 1))
        } else {
            None
        }
    }

    pub fn value(self) -> u8 {
        self.0
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct Indentation {
    pub start: Option<Dimension<Twips>>,
    pub end: Option<Dimension<Twips>>,
    pub first_line: Option<FirstLineIndent>,
    pub mirror: Option<bool>,
}

/// First-line indent: either hanging (negative) or first-line (positive).
/// These are mutually exclusive in OOXML.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FirstLineIndent {
    None,
    FirstLine(Dimension<Twips>),
    Hanging(Dimension<Twips>),
}

/// Paragraph spacing — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct ParagraphSpacing {
    pub before: Option<Dimension<Twips>>,
    pub after: Option<Dimension<Twips>>,
    pub line: Option<LineSpacing>,
    pub before_auto_spacing: Option<bool>,
    pub after_auto_spacing: Option<bool>,
}

/// Line spacing rule — the three OOXML modes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineSpacing {
    /// Automatic (proportional). Value is in 240ths of a line (240 = single).
    Auto(Dimension<Twips>),
    /// Exact line height.
    Exact(Dimension<Twips>),
    /// Minimum line height (at least this much).
    AtLeast(Dimension<Twips>),
}

/// Raw numbering reference on a paragraph (w:numPr).
/// Resolve via `Document.numbering` using `num_id` + `level`.
#[derive(Clone, Copy, Debug)]
pub struct NumberingReference {
    pub num_id: i64,
    pub level: u8,
}
