//! Content model — Block and Inline enums, hyperlinks, fields, alternate content.

use crate::field::FieldInstruction;

use super::drawing::Image;
use super::identifiers::{BookmarkId, CommentId, NoteId, RevisionIds, StyleId};
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
    /// §17.3.1.30: absolute-position tab (`<w:ptab>`).
    PositionTab(super::formatting::PositionTab),
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
    /// `<w:commentReference>` (issue #154): the comment's anchor — where the
    /// balloon's connector attaches. Zero-width, like a bookmark target.
    CommentRef(CommentId),
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
    /// §17.13.5.14 / §17.13.5.18: the unaccepted tracked change this run sits
    /// inside, when it does — the parse flattens the `<w:ins>`/`<w:del>`
    /// wrapper and stamps each run it contained (issue #154). `None` for the
    /// overwhelming majority of runs. Unlike `rsids` above — save-session
    /// metadata nothing renders — this is document content: an unaccepted
    /// deletion's text exists in the document *as a deletion*, and which way
    /// it reaches the page (struck through, or suppressed) is the renderer's
    /// `w:revisionView` decision, not the parser's.
    pub revision: Option<RunRevision>,
    /// Issue #154: the comment range this run sits inside, when it does —
    /// the parse tracks `commentRangeStart`/`End` markers (which can span
    /// paragraphs) and stamps the runs between them, the same shape as
    /// `revision` above. The first-opened range wins when ranges overlap.
    pub comment: Option<CommentId>,
}

/// The tracked change a [`TextRun`] belongs to.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RunRevision {
    pub kind: RevisionKind,
    /// `@w:author` of the wrapper — what revision marks are colored by.
    /// Empty when the document omits it.
    pub author: String,
}

/// Which side of a tracked change: `<w:ins>` or `<w:del>`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum RevisionKind {
    Inserted,
    Deleted,
}

/// One comment from the comments part (issue #154): who wrote it and its
/// block content — the balloon body. The `@w:date` attribute stays
/// uncaptured: the balloon shows author and text, and nothing else reads it.
#[derive(Clone, Debug)]
pub struct Comment {
    /// `@w:author` — balloons are labeled and colored by it.
    pub author: String,
    /// `@w:initials`, when present — the compact label Word puts on the
    /// anchor; empty when absent.
    pub initials: String,
    pub content: Vec<Block>,
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

/// A hyperlink target. External links start life as an unresolved relationship
/// id (`w:hyperlink/@r:id`) and are rewritten to a concrete URL by the
/// parse-time `resolve_hyperlinks` pass; the two states are distinct variants so
/// no code has to guess whether an `External` payload is an rId or a URL.
#[derive(Clone, Debug)]
pub enum HyperlinkTarget {
    /// Unresolved external relationship id, before the rels lookup.
    ExternalRel(super::identifiers::RelId),
    /// Resolved external URL, after `resolve_hyperlinks`.
    ExternalUrl(String),
    /// Internal bookmark anchor.
    Internal { anchor: String },
}

// ── Field ────────────────────────────────────────────────────────────────────

/// A simple field (w:fldSimple). Stores the parsed field instruction.
#[derive(Clone, Debug)]
pub struct Field {
    /// Parsed field instruction (e.g., `FieldInstruction::Page`, `FieldInstruction::Toc { .. }`).
    pub instruction: FieldInstruction,
    /// The cached result inlines from when the document was last saved —
    /// Word's own evaluation. Used as-is for fields this engine doesn't
    /// dynamically evaluate, and as the fallback when `FieldContext` lacks a
    /// value for one it does.
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
    /// Required namespace/feature identifiers. `@Requires` is a space-separated
    /// list of prefixes; the choice is usable only if every one is understood.
    pub requires: Vec<McRequires>,
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
