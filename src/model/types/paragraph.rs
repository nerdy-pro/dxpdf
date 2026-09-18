//! Paragraph types — paragraph properties, spacing, indentation, frames.

use crate::model::dimension::{Dimension, Twips};
use crate::model::Dup;

use super::formatting::{
    Alignment, CnfStyle, HeightRule, ParagraphBorders, Shading, TabStop, TableAnchor, TableXAlign,
    TableYAlign, TextAlignment,
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

/// Paragraph properties — only fields explicitly present in the XML are set.
///
/// Every non-toggle child of `<w:pPr>` is a [`Dup`]: the schema allows it once,
/// producers repeat it anyway, and the model carries every occurrence so the
/// choice of which one wins belongs to the reader rather than the parser. See
/// `model::dup`. The `Option<bool>` toggles below are the deliberate exception
/// — §17.7.2 defines last-wins for them, so `last_toggle` applies the *spec's*
/// rule at the seam and there is nothing left to carry.
#[derive(Clone, Debug, Default)]
pub struct ParagraphProperties {
    pub alignment: Dup<Alignment>,
    pub indentation: Dup<Indentation>,
    pub spacing: Dup<ParagraphSpacing>,
    pub numbering: Dup<NumberingReference>,
    /// §17.3.1.38 `<w:tabs>`. The one pPr child that stays a plain `Vec`: it is
    /// a *container*, so absence is an empty `Vec` rather than `None`, and
    /// `Dup<Vec<TabStop>>` would make "the document set no tabs" and "the
    /// document set an empty `<w:tabs/>`" the same value. A repeated `<w:tabs>`
    /// therefore collapses at the seam, last-wins like everything else.
    ///
    /// Word reference render needed: whether Word takes the last `<w:tabs>` or
    /// unions the tab stops of both. §17.3.1.38 describes one container and
    /// does not say. A document with two `<w:tabs>` holding different positions
    /// would settle it — if Word unions them, this becomes `Dup` and the read
    /// site concatenates.
    pub tabs: Vec<TabStop>,
    /// §17.3.1.24 `<w:pBdr>`. The sides *within* one `<w:pBdr>` stay plain so
    /// [`ParagraphBorders`] keeps `Copy`. See `model::dup` for where the line is.
    pub borders: Dup<ParagraphBorders>,
    pub shading: Dup<Shading>,
    pub keep_next: Option<bool>,
    pub keep_lines: Option<bool>,
    pub widow_control: Option<bool>,
    /// §17.3.1.32: use the section's document grid for inter-line spacing.
    /// Absent everywhere in the style hierarchy means *on* — the paragraph
    /// snaps whenever its section defines an applied grid.
    pub snap_to_grid: Option<bool>,
    pub page_break_before: Option<bool>,
    pub suppress_auto_hyphens: Option<bool>,
    /// §17.3.1.9: suppress spacing when adjacent paragraph has same style.
    pub contextual_spacing: Option<bool>,
    /// §17.3.1.6: this paragraph is right-to-left.
    ///
    /// The base embedding level for UAX #9 (issue #131), and with it which
    /// physical edge `Alignment::Start` and `Indentation::start` mean — see
    /// `render::layout::build::convert::base_direction`. Never inferred from
    /// the text: a document that states its direction outranks a heuristic
    /// reading of its own characters.
    pub bidi: Option<bool>,
    /// §17.3.1.45: allow line breaking between any characters for East Asian text.
    pub word_wrap: Option<bool>,
    pub outline_level: Dup<OutlineLevel>,
    /// §17.3.1.39: vertical alignment of text on each line (ST_TextAlignment).
    pub text_alignment: Dup<TextAlignment>,
    /// §17.3.1.8: table conditional formatting applied to this paragraph.
    pub cnf_style: Dup<CnfStyle>,
    /// §17.3.1.11: text frame (legacy positioned text region).
    pub frame_properties: Dup<FrameKind>,
    /// §17.3.1.2: auto-space East Asian text with Latin text.
    pub auto_space_de: Option<bool>,
    /// §17.3.1.3: auto-space East Asian text with numbers.
    pub auto_space_dn: Option<bool>,
}

/// §17.3.1.11: frame kind — a paragraph is either a drop cap or a floating text box.
///
/// OOXML `w:framePr` conflates these two uses; the presence and value of `w:dropCap`
/// determines which kind a given `framePr` represents.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FrameKind {
    /// The paragraph's first character(s) are rendered as a drop cap.
    DropCap {
        /// Whether the cap is inline (`Drop`) or in the margin (`Margin`).
        style: DropCap,
        /// Number of body-text lines the cap letter spans (default 3).
        lines: u32,
        /// Horizontal space between cap and body text, in twips.
        h_space: Option<Dimension<Twips>>,
    },
    /// Legacy floating text frame — positioned outside normal flow.
    TextBox(TextBoxPositioning),
}

/// §17.18.16 ST_DropCap — the two active drop-cap styles.
/// (`dropCap="none"` is represented by the absence of `FrameKind::DropCap`.)
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DropCap {
    /// Cap is inline, to the left of the paragraph text.
    Drop,
    /// Cap is placed in the page margin.
    Margin,
}

/// Positioning attributes for a legacy floating text frame (`w:framePr` without drop cap).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct TextBoxPositioning {
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

    /// Create from a §17.3.1.19 `w:outlineLvl/@w:val`, which is 0-based.
    ///
    /// Returns `None` for 9, and that is the spec rather than a bound: value 9
    /// is "body text" — an explicit statement that the paragraph has **no**
    /// outline level, not a ninth heading level. Word writes it when a
    /// paragraph's outline level is reset to Body Text, so a heading style's
    /// level can be overridden back off. `None` therefore means "not a
    /// heading", whether the attribute was absent or present-and-9. Not
    /// theoretical: `sample-emoji.docx` in the test corpus declares
    /// `w:outlineLvl w:val="9"` twice.
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn outline_level_from_ooxml_is_one_based() {
        // OOXML `w:outlineLvl` is 0-based (0 = Heading 1); storage is 1-based.
        assert_eq!(OutlineLevel::from_ooxml(0).unwrap().value(), 1);
        assert_eq!(OutlineLevel::from_ooxml(8).unwrap().value(), 9);
    }

    #[test]
    fn outline_level_from_ooxml_rejects_out_of_range() {
        assert!(OutlineLevel::from_ooxml(9).is_none());
        assert!(OutlineLevel::from_ooxml(255).is_none());
    }

    #[test]
    fn outline_level_new_accepts_valid_range() {
        assert_eq!(OutlineLevel::new(1).value(), 1);
        assert_eq!(OutlineLevel::new(9).value(), 9);
    }

    #[test]
    #[should_panic(expected = "outline level must be 1..=9")]
    fn outline_level_new_rejects_zero() {
        OutlineLevel::new(0);
    }

    #[test]
    #[should_panic(expected = "outline level must be 1..=9")]
    fn outline_level_new_rejects_ten() {
        OutlineLevel::new(10);
    }
}
