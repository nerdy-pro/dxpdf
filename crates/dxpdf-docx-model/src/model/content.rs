//! Content model — Block and Inline enums, hyperlinks, fields, alternate content.

use dxpdf_field::FieldInstruction;

use super::drawing::Image;
use super::identifiers::{BookmarkId, NoteId, RevisionIds, StyleId};
use super::paragraph::Paragraph;
use super::run_properties::RunProperties;
use super::section::SectionProperties;
use super::table::Table;
use super::vml::Pict;

// ── Blocks ───────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub enum Block {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
    /// A section break that applies to all preceding content since the last break.
    SectionBreak(Box<SectionProperties>),
}

// ── Inline content ───────────────────────────────────────────────────────────

/// A child element within a `<w:r>` run. All elements in a run share
/// the same `RunProperties` (font, size, color, etc.).
#[derive(Clone, Debug)]
pub enum RunElement {
    Text(String),
    Tab,
    LineBreak(BreakKind),
    ColumnBreak,
    PageBreak,
    /// §17.3.3.13: rendering hint, not a content break.
    LastRenderedPageBreak,
}

#[derive(Clone, Debug)]
pub enum Inline {
    TextRun(Box<TextRun>),
    Image(Box<Image>),
    FootnoteRef(NoteId),
    EndnoteRef(NoteId),
    Hyperlink(Hyperlink),
    Field(Field),
    BookmarkStart {
        id: BookmarkId,
        name: String,
    },
    BookmarkEnd(BookmarkId),
    Symbol(Symbol),
    /// §17.11.23: footnote/endnote separator line.
    Separator,
    /// §17.11.3: continuation separator for notes spanning pages.
    ContinuationSeparator,
    /// §17.16.18: complex field character (begin/separate/end marker).
    FieldChar(FieldChar),
    /// §17.16.23: field instruction text (appears between begin and separate).
    InstrText(String),
    /// §17.11.13: footnote reference mark (auto-number rendered in the footnote body).
    FootnoteRefMark,
    /// §17.11.6: endnote reference mark (auto-number rendered in the endnote body).
    EndnoteRefMark,
    /// §17.3.3.19: VML picture/shape container (legacy drawing).
    Pict(Pict),
    /// MCE §M.2.1: markup compatibility alternate content.
    AlternateContent(AlternateContent),
}

/// §17.16.18: complex field character marker.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FieldChar {
    /// §17.18.29 ST_FldCharType: begin, separate, or end.
    pub field_char_type: FieldCharType,
    /// Field result needs recalculation.
    pub dirty: Option<bool>,
    /// Field is locked from updates.
    pub fld_lock: Option<bool>,
}

/// §17.18.29 ST_FldCharType
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FieldCharType {
    Begin,
    Separate,
    End,
}

#[derive(Clone, Debug)]
pub struct TextRun {
    /// Character style ID reference (e.g., "Hyperlink"). Resolve via `Document.styles`.
    pub style_id: Option<StyleId>,
    pub properties: RunProperties,
    /// Children of this run: text segments, breaks, and tabs.
    /// All share the run's properties.
    pub content: Vec<RunElement>,
    pub rsids: RevisionIds,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BreakKind {
    TextWrapping,
    /// Clears left, right, or both float areas.
    Clear(BreakClear),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BreakClear {
    None,
    Left,
    Right,
    All,
}

/// A symbol character from a specific font (w:sym).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Symbol {
    pub font: String,
    pub char_code: u16,
}

// ── Hyperlink ────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Hyperlink {
    pub target: HyperlinkTarget,
    pub content: Vec<Inline>,
}

/// A hyperlink either targets an external URL (via relationship) or an internal bookmark.
#[derive(Clone, Debug)]
pub enum HyperlinkTarget {
    External(super::identifiers::RelId),
    Internal { anchor: String },
}

// ── Field ────────────────────────────────────────────────────────────────────

/// A simple field (w:fldSimple). Stores the parsed field instruction.
#[derive(Clone, Debug)]
pub struct Field {
    /// Parsed field instruction (e.g., `FieldInstruction::Page`, `FieldInstruction::Toc { .. }`).
    pub instruction: FieldInstruction,
    pub content: Vec<Inline>,
}

// ── Alternate Content ────────────────────────────────────────────────────────

/// MCE §M.2.1: alternate content for markup compatibility.
#[derive(Clone, Debug)]
pub struct AlternateContent {
    /// §M.2.2: ordered list of preferred choices.
    pub choices: Vec<McChoice>,
    /// §M.2.3: fallback content when no choice is supported.
    pub fallback: Option<Vec<Inline>>,
}

/// MCE §M.2.2: a single choice in alternate content.
#[derive(Clone, Debug)]
pub struct McChoice {
    /// Required namespace/feature identifier.
    pub requires: McRequires,
    /// Inline content for this choice.
    pub content: Vec<Inline>,
}

/// MCE §M.2.2: namespace prefixes used in `mc:Choice Requires`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum McRequires {
    /// Word Processing Shape (wps).
    Wps,
    /// Word Processing Group (wpg).
    Wpg,
    /// Word Processing Canvas (wpc).
    Wpc,
    /// Word Processing Ink (wpi).
    Wpi,
    /// Math (m).
    Math,
    /// DrawingML 2010 (a14).
    A14,
    /// Word 2010 extensions (w14).
    W14,
    /// Word 2012 extensions (w15).
    W15,
    /// Word 2016 extensions (w16).
    W16,
}
