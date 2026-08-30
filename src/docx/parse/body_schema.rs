//! Serde schema for `<w:body>` / `<w:hdr>` / `<w:ftr>` / `<w:footnote>` contents.
//!
//! The body grammar is mixed-content: paragraphs contain runs, hyperlinks,
//! fields, and bookmarks in arbitrary order; runs themselves contain text,
//! tabs, breaks, drawings, pictures, and field characters in arbitrary order.
//! We use `#[serde(rename = "$value")]` + untagged enums to preserve order,
//! then flatten to the model's `Vec<Inline>` in `From` impls.
//!
//! `<w:drawing>` and `<w:pict>` sub-trees are handed to the legacy DrawingML
//! and VML parsers via a two-pass approach (see `parse/body.rs`); the schema
//! treats them as placeholders and the merge step fills them in.
//!
//! Serde deserializes into many fields that are only read during the
//! `From`-style conversion in `body.rs`; the `allow(dead_code)` silences
//! the spurious warnings.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::dimension::{Dimension, Twips};
use crate::docx::parse::primitives::st_enums::{
    StBrClear, StFldCharType, StPTabAlignment, StPTabLeader, StPTabRelativeTo,
};
use crate::docx::parse::primitives::units::deserialize_optional_nonnegative_dimension;
use crate::docx::parse::properties::schema::paragraph::PPrXml;
use crate::docx::parse::properties::schema::run::RPrXml;
use crate::docx::parse::properties::schema::section::SectPrXml;
use crate::docx::parse::properties::schema::table::{TblPrExXml, TblPrXml, TcPrXml, TrPrXml};

// ── root-level ─────────────────────────────────────────────────────────────

/// Top-level container (body / header / footer / footnote / endnote body).
#[derive(Deserialize, Default)]
pub(crate) struct BlockContainerXml {
    #[serde(rename = "$value", default)]
    pub children: Vec<BlockChildXml>,
}

/// Union of direct children of a body-level container. Most variants beyond
/// paragraph/table/sectPr are structural elements (bookmarks, comment range
/// markers, proofing errors, SDT wrappers) that OOXML allows at block level
/// but the renderer discards. We model them so serde can skip them cleanly.
#[derive(Deserialize)]
pub(crate) enum BlockChildXml {
    #[serde(rename = "p")]
    Paragraph(Box<ParaXml>),
    #[serde(rename = "tbl")]
    Table(Box<TableXml>),
    #[serde(rename = "sectPr")]
    SectPr(Box<SectPrXml>),
    /// Block-level bookmark start (spans multiple blocks).
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(BookmarkStartXml),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(BookmarkEndXml),
    /// Comment range markers — ignored.
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(BookmarkEndXml),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(BookmarkEndXml),
    /// Proofing error markers — ignored.
    #[serde(rename = "proofErr")]
    ProofErr(IgnoredXml),
    /// Structured document tag wrappers — contents are treated as blocks.
    #[serde(rename = "sdt")]
    Sdt(Box<SdtBlockXml>),
    /// Catch-all for other OOXML elements we don't yet model.
    #[serde(other)]
    Other,
}

/// Placeholder type for elements we accept but don't process.
#[derive(Deserialize, Default)]
pub(crate) struct IgnoredXml {}

/// `<w:sdt>` block-level structured document tag — extract the content
/// from `<w:sdtContent>` and treat it as block-level children.
#[derive(Deserialize, Default)]
pub(crate) struct SdtBlockXml {
    #[serde(rename = "sdtContent", default)]
    pub content: Vec<SdtBlockContentXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct SdtBlockContentXml {
    #[serde(rename = "$value", default)]
    pub children: Vec<BlockChildXml>,
}

// ── paragraph ──────────────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
pub(crate) struct ParaXml {
    #[serde(rename = "@rsidR", default)]
    pub rsid_r: Option<String>,
    #[serde(rename = "@rsidRDefault", default)]
    pub rsid_r_default: Option<String>,
    #[serde(rename = "@rsidP", default)]
    pub rsid_p: Option<String>,
    #[serde(rename = "@rsidRPr", default)]
    pub rsid_r_pr: Option<String>,
    #[serde(rename = "@rsidDel", default)]
    pub rsid_del: Option<String>,

    #[serde(rename = "pPr", default)]
    pub p_pr: Vec<PPrXml>,
    #[serde(rename = "$value", default)]
    pub content: Vec<ParaChildXml>,
}

/// Children of `<w:p>` excluding `<w:pPr>` (which is captured separately).
/// This is `EG_PContent` — also the content model of `<w:hyperlink>`,
/// `<w:fldSimple>`, and the revision/structural wrappers below.
///
/// OOXML wraps run content in revision-tracking (`ins`/`del`/`moveFrom`/
/// `moveTo`) and structural (`smartTag`/`customXml`) elements. These are
/// modelled so `convert_para_children` can flatten them: insert-side and
/// structural wrappers are rendered, delete-side wrappers are dropped (an
/// "accept all changes" / final view). Comment range markers are modelled so
/// the runs between them can be stamped (issue #154). Remaining annotation
/// elements (proofErr, permStart/End, sdt, ...) hit the `Other` catch-all
/// and are discarded cleanly.
#[derive(Deserialize)]
pub(crate) enum ParaChildXml {
    #[serde(rename = "r")]
    Run(RunXml),
    #[serde(rename = "hyperlink")]
    Hyperlink(HyperlinkXml),
    #[serde(rename = "fldSimple")]
    FldSimple(FldSimpleXml),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(BookmarkStartXml),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(BookmarkEndXml),
    /// Issue #154: comment range markers — toggles for the run stamp.
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(BookmarkEndXml),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(BookmarkEndXml),
    /// §17.13.5.18 `<w:ins>` — insert-side revision wrapper; content is kept.
    #[serde(rename = "ins")]
    Ins(RunTrackChangeXml),
    /// §17.13.5.14 `<w:del>` — delete-side revision wrapper; content is dropped.
    #[serde(rename = "del")]
    Del(RunTrackChangeXml),
    /// §17.13.5.22 `<w:moveTo>` — destination side of a move; content is kept.
    #[serde(rename = "moveTo")]
    MoveTo(RunTrackChangeXml),
    /// §17.13.5.19 `<w:moveFrom>` — source side of a move; content is dropped.
    #[serde(rename = "moveFrom")]
    MoveFrom(RunTrackChangeXml),
    /// §17.5.1.9 `<w:smartTag>` — structural run wrapper; content is flattened.
    #[serde(rename = "smartTag")]
    SmartTag(RunTrackChangeXml),
    /// §17.5.1.6 `<w:customXml>` (CT_CustomXmlRun) — content is flattened.
    #[serde(rename = "customXml")]
    CustomXml(RunTrackChangeXml),
    /// `<w:pPr>` is captured on `ParaXml` directly, but serde's untagged
    /// enum still has to handle it if it appears in `$value` ordering.
    #[serde(rename = "pPr")]
    PPr(Box<PPrXml>),
    #[serde(other)]
    Other,
}

/// A run-level revision (`ins`/`del`/`moveFrom`/`moveTo`, CT_RunTrackChange)
/// or structural (`smartTag`/`customXml`) wrapper. Its content model is
/// `EG_PContent` — the same children as `<w:p>` — so it nests recursively.
/// The paragraph-level analogue of `RowTrackChangeXml`. Any `customXmlPr`
/// child is absorbed by `$value` and ignored via `ParaChildXml::Other`.
#[derive(Deserialize, Default)]
pub(crate) struct RunTrackChangeXml {
    /// §17.13.5.14/.18 `@w:author` — who made the change; revision marks are
    /// colored by it. `@w:id` and `@w:date` stay uncaptured: nothing renders
    /// them, and the id numbers revisions within one save, not the document.
    #[serde(rename = "@author")]
    pub author: Option<String>,
    #[serde(rename = "$value", default)]
    pub content: Vec<ParaChildXml>,
}

// ── run ────────────────────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
pub(crate) struct RunXml {
    #[serde(rename = "@rsidR", default)]
    pub rsid_r: Option<String>,
    #[serde(rename = "@rsidRPr", default)]
    pub rsid_r_pr: Option<String>,
    #[serde(rename = "@rsidDel", default)]
    pub rsid_del: Option<String>,

    #[serde(rename = "rPr", default)]
    pub r_pr: Vec<RPrXml>,
    #[serde(rename = "$value", default)]
    pub content: Vec<RunChildXml>,
}

/// Children of `<w:r>`. Includes both "run element" kinds (text, tab, break
/// — collected into a single `TextRun`) and "sibling inline" kinds (drawing,
/// pict, sym, etc. — each becomes its own `Inline` at the parent level).
#[derive(Deserialize)]
pub(crate) enum RunChildXml {
    #[serde(rename = "t")]
    Text(TextXml),
    #[serde(rename = "delText")]
    DelText(TextXml),
    #[serde(rename = "tab")]
    Tab,
    #[serde(rename = "ptab")]
    PTab(PTabXml),
    #[serde(rename = "br")]
    Br(BrXml),
    #[serde(rename = "cr")]
    Cr,
    #[serde(rename = "softHyphen")]
    SoftHyphen,
    #[serde(rename = "noBreakHyphen")]
    NoBreakHyphen,
    #[serde(rename = "lastRenderedPageBreak")]
    LastRenderedPageBreak,
    #[serde(rename = "drawing")]
    Drawing(DrawingXml),
    #[serde(rename = "pict")]
    Pict(crate::docx::parse::vml::schema::PictXml),
    #[serde(rename = "sym")]
    Sym(SymXml),
    #[serde(rename = "instrText")]
    InstrText(TextXml),
    #[serde(rename = "fldChar")]
    FldChar(FldCharXml),
    #[serde(rename = "footnoteReference")]
    FootnoteRef(NoteRefXml),
    #[serde(rename = "endnoteReference")]
    EndnoteRef(NoteRefXml),
    #[serde(rename = "footnoteRef")]
    FootnoteRefMark,
    #[serde(rename = "endnoteRef")]
    EndnoteRefMark,
    #[serde(rename = "separator")]
    Separator,
    #[serde(rename = "continuationSeparator")]
    ContinuationSeparator,
    #[serde(rename = "AlternateContent")]
    AlternateContent(AltContentXml),
    /// `<w:commentReference>` — comment anchor inside a run. Comments are not
    /// rendered (matching Word's default print output); parsed only so a
    /// commented run doesn't fail the document. The range markers
    /// (`commentRangeStart/End`) are handled at the paragraph level above.
    #[serde(rename = "commentReference")]
    CommentReference(BookmarkEndXml),
    /// `<w:annotationRef>` — the reference mark inside a comment's own body
    /// (issue #154). Dropped: the balloon labels itself with the author.
    #[serde(rename = "annotationRef")]
    AnnotationRef(IgnoredXml),
    /// `<w:rPr>` captured separately; included here for serde ordering.
    #[serde(rename = "rPr")]
    RPr(Box<RPrXml>),
}

// ── inline sub-types ───────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
pub(crate) struct TextXml {
    #[serde(rename = "@xml:space", default)]
    pub space: Option<String>,
    #[serde(rename = "$text", default)]
    pub content: String,
}

#[derive(Deserialize, Default)]
pub(crate) struct BrXml {
    #[serde(rename = "@type", default)]
    pub ty: Option<StBrType>,
    #[serde(rename = "@clear", default)]
    pub clear: Option<StBrClear>,
}

#[derive(Deserialize, Clone, Copy, Debug)]
#[serde(rename_all = "camelCase")]
pub(crate) enum StBrType {
    Page,
    Column,
    TextWrapping,
}

/// §17.3.1.30 `<w:ptab>` — absolute-position tab. All three attributes are
/// required by the schema; we default leniently (left / indent / none) so a
/// malformed run does not abort the whole parse.
#[derive(Deserialize)]
pub(crate) struct PTabXml {
    #[serde(rename = "@alignment", default = "default_ptab_alignment")]
    pub alignment: StPTabAlignment,
    #[serde(rename = "@relativeTo", default = "default_ptab_relative_to")]
    pub relative_to: StPTabRelativeTo,
    #[serde(rename = "@leader", default = "default_ptab_leader")]
    pub leader: StPTabLeader,
}

fn default_ptab_alignment() -> StPTabAlignment {
    StPTabAlignment::Left
}

fn default_ptab_relative_to() -> StPTabRelativeTo {
    StPTabRelativeTo::Indent
}

fn default_ptab_leader() -> StPTabLeader {
    StPTabLeader::None
}

impl From<PTabXml> for crate::docx::model::PositionTab {
    fn from(x: PTabXml) -> Self {
        Self {
            alignment: x.alignment.into(),
            relative_to: x.relative_to.into(),
            leader: x.leader.into(),
        }
    }
}

#[derive(Deserialize, Default)]
pub(crate) struct SymXml {
    #[serde(rename = "@font", default)]
    pub font: String,
    /// Hex string like `"F0A2"` — parsed to u16 at From time.
    #[serde(rename = "@char", default)]
    pub char: String,
}

#[derive(Deserialize)]
pub(crate) struct FldCharXml {
    #[serde(rename = "@fldCharType")]
    pub fld_char_type: StFldCharType,
    #[serde(rename = "@dirty", default)]
    pub dirty: Option<AttrBool>,
    #[serde(rename = "@fldLock", default)]
    pub fld_lock: Option<AttrBool>,
}

#[derive(Deserialize)]
pub(crate) struct NoteRefXml {
    #[serde(rename = "@id")]
    pub id: i64,
}

#[derive(Deserialize)]
pub(crate) struct BookmarkStartXml {
    #[serde(rename = "@id")]
    pub id: i64,
    #[serde(rename = "@name", default)]
    pub name: String,
}

#[derive(Deserialize)]
pub(crate) struct BookmarkEndXml {
    #[serde(rename = "@id")]
    pub id: i64,
}

/// `<w:drawing>` wrapper — contains exactly one `<wp:inline>` or
/// `<wp:anchor>` child (both modelled in `drawing/schema/anchor.rs`).
#[derive(Deserialize)]
pub(crate) struct DrawingXml {
    #[serde(rename = "inline", default)]
    pub inline: Vec<crate::docx::parse::drawing::schema::anchor::InlineXml>,
    #[serde(rename = "anchor", default)]
    pub anchor: Vec<crate::docx::parse::drawing::schema::anchor::AnchorXml>,
}

// ── hyperlink ──────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct HyperlinkXml {
    #[serde(rename = "@id", default)]
    pub r_id: Option<String>,
    #[serde(rename = "@anchor", default)]
    pub anchor: Option<String>,
    #[serde(rename = "$value", default)]
    pub content: Vec<ParaChildXml>,
}

// ── simple field ───────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct FldSimpleXml {
    #[serde(rename = "@instr", default)]
    pub instr: String,
    #[serde(rename = "$value", default)]
    pub content: Vec<ParaChildXml>,
}

// ── alternate content ──────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
pub(crate) struct AltContentXml {
    #[serde(rename = "Choice", default)]
    pub choices: Vec<ChoiceXml>,
    #[serde(rename = "Fallback", default)]
    pub fallback: Vec<FallbackXml>,
}

#[derive(Deserialize)]
pub(crate) struct ChoiceXml {
    #[serde(rename = "@Requires", default)]
    pub requires: String,
    #[serde(rename = "$value", default)]
    pub content: Vec<McContentXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct FallbackXml {
    #[serde(rename = "$value", default)]
    pub content: Vec<McContentXml>,
}

/// Legacy parser only supports `drawing` and `pict` inside mc:Choice /
/// mc:Fallback — other children are ignored.
#[derive(Deserialize)]
pub(crate) enum McContentXml {
    #[serde(rename = "drawing")]
    Drawing(DrawingXml),
    #[serde(rename = "pict")]
    Pict(crate::docx::parse::vml::schema::PictXml),
}

// ── table ──────────────────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
pub(crate) struct TableXml {
    #[serde(rename = "tblPr", default)]
    pub tbl_pr: Vec<TblPrXml>,
    #[serde(rename = "tblGrid", default)]
    pub tbl_grid: Vec<TblGridXml>,
    /// All direct children of `<w:tbl>` in document order. `<w:tblPr>` and
    /// `<w:tblGrid>` are also captured here (serde's `$value` collects every
    /// child) and must be ignored when extracting rows; the dedicated
    /// `tbl_pr` / `tbl_grid` fields above are canonical.
    #[serde(rename = "$value", default)]
    pub children: Vec<TableChildXml>,
}

/// Direct children of `<w:tbl>` modelled per ECMA-376 §17.4.38 (CT_Tbl).
///
/// After `<w:tblPr>` and `<w:tblGrid>`, OOXML's content model is a
/// `<xsd:choice>` repeated 0..* (`EG_ContentRowContent`): `<w:tr>`,
/// `<w:customXml>` (CT_CustomXmlRow), `<w:sdt>` (CT_SdtRow), or
/// EG_RunLevelElts (proofErr/permStart/permEnd; ins/del/moveFrom/moveTo as
/// CT_RowTrackChange; EG_RangeMarkupElements: bookmarkStart/End,
/// commentRange*; EG_MathContent). Modelling every spec-defined variant
/// lets row extraction recurse into wrappers; non-row variants are dropped
/// during conversion since they have no rendered effect at table level.
#[derive(Deserialize)]
pub(crate) enum TableChildXml {
    /// `<w:tr>` — a table row.
    #[serde(rename = "tr")]
    Row(Box<TableRowXml>),
    /// `<w:sdt>` — table-level structured document tag (CT_SdtRow).
    #[serde(rename = "sdt")]
    Sdt(Box<SdtRowXml>),
    /// `<w:ins>` revision-tracked inserted rows (CT_RowTrackChange).
    #[serde(rename = "ins")]
    Ins(Box<RowTrackChangeXml>),
    /// `<w:del>` revision-tracked deleted rows.
    #[serde(rename = "del")]
    Del(Box<RowTrackChangeXml>),
    /// `<w:moveFrom>` / `<w:moveTo>` revision-tracked moved rows.
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<RowTrackChangeXml>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<RowTrackChangeXml>),
    /// `<w:customXml>` — row-level custom-XML wrapper (CT_CustomXmlRow).
    #[serde(rename = "customXml")]
    CustomXml(Box<CustomXmlRowXml>),
    /// Range markup — bookmarks and comment ranges may span multiple rows.
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(BookmarkStartXml),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(BookmarkEndXml),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(BookmarkEndXml),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(BookmarkEndXml),
    /// Proofreading and permission markers — ignored.
    #[serde(rename = "proofErr")]
    ProofErr(IgnoredXml),
    #[serde(rename = "permStart")]
    PermStart(IgnoredXml),
    #[serde(rename = "permEnd")]
    PermEnd(IgnoredXml),
    /// `<w:tblPr>` / `<w:tblGrid>` are captured on the parent directly;
    /// these variants exist only so `$value` can absorb the duplicate
    /// match. Compare with `ParaChildXml::PPr`.
    #[serde(rename = "tblPr")]
    TblPr(Box<TblPrXml>),
    #[serde(rename = "tblGrid")]
    TblGrid(Box<TblGridXml>),
    /// Catch-all for unmodelled OOXML elements (e.g., math content,
    /// `customXml*RangeStart/End`).
    #[serde(other)]
    Other,
}

/// CT_SdtRow §17.5.2.30 — `<w:sdt>` at table level. Wrapped rows live in
/// `<w:sdtContent>` (CT_SdtContentRow).
#[derive(Deserialize, Default)]
pub(crate) struct SdtRowXml {
    #[serde(rename = "sdtContent", default)]
    pub content: Vec<SdtRowContentXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct SdtRowContentXml {
    #[serde(rename = "$value", default)]
    pub children: Vec<TableChildXml>,
}

/// CT_RowTrackChange §17.13.5.16/.7/.22 — used for `<w:ins>` / `<w:del>` /
/// `<w:moveFrom>` / `<w:moveTo>` at table level. Per the spec, contains
/// only `<w:tr>` children plus the inherited revision attributes.
#[derive(Deserialize, Default)]
pub(crate) struct RowTrackChangeXml {
    #[serde(rename = "@id", default)]
    pub id: Option<i64>,
    #[serde(rename = "@author", default)]
    pub author: Option<String>,
    #[serde(rename = "@date", default)]
    pub date: Option<String>,
    #[serde(rename = "tr", default)]
    pub rows: Vec<TableRowXml>,
}

/// CT_CustomXmlRow §17.5.1.5–.9 — recursive wrapper whose content model
/// mirrors CT_Tbl's row-level choice group, hence `Vec<TableChildXml>`.
#[derive(Deserialize, Default)]
pub(crate) struct CustomXmlRowXml {
    #[serde(rename = "customXmlPr", default)]
    pub custom_xml_pr: Vec<IgnoredXml>,
    #[serde(rename = "$value", default)]
    pub children: Vec<TableChildXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TblGridXml {
    #[serde(rename = "gridCol", default)]
    pub cols: Vec<GridColXml>,
}

#[derive(Deserialize)]
pub(crate) struct GridColXml {
    #[serde(
        rename = "@w",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    pub w: Option<Dimension<Twips>>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TableRowXml {
    #[serde(rename = "@rsidR", default)]
    pub rsid_r: Option<String>,
    #[serde(rename = "@rsidRPr", default)]
    pub rsid_r_pr: Option<String>,
    #[serde(rename = "@rsidDel", default)]
    pub rsid_del: Option<String>,
    #[serde(rename = "@rsidTr", default)]
    pub rsid_tr: Option<String>,

    /// §17.4.61 `<w:tblPrEx>` — per-row exceptions to table-level
    /// properties. Spec mandates that when present, it appears *before*
    /// `<w:trPr>` inside `<w:tr>`. Also absorbed by `$value` below (via
    /// `RowChildXml::TblPrEx`); this dedicated field is canonical.
    #[serde(rename = "tblPrEx", default)]
    pub tbl_pr_ex: Vec<TblPrExXml>,
    #[serde(rename = "trPr", default)]
    pub tr_pr: Vec<TrPrXml>,
    /// All direct children of `<w:tr>` in document order. `<w:tc>` cells may
    /// be interleaved with `<w:sdt>` / `<w:customXml>` wrappers (each of
    /// which nests further cells), so — like `<w:tbl>` — the row cannot use a
    /// plain `Vec<TableCellXml>`: `quick_xml` reports a duplicate `tc` field
    /// when cells are non-contiguous. Cells are flattened out during
    /// conversion by `collect_row_cells`.
    #[serde(rename = "$value", default)]
    pub children: Vec<RowChildXml>,
}

/// Direct children of `<w:tr>` modelled per ECMA-376 §17.4 (CT_Row).
///
/// After `<w:tblPrEx>` and `<w:trPr>`, the content model is the
/// `EG_ContentCellContent` group repeated 0..*: `<w:tc>` (CT_Tc),
/// `<w:customXml>` (CT_CustomXmlCell), `<w:sdt>` (CT_SdtCell), or
/// EG_RunLevelElts (proofErr/permStart/permEnd) and EG_RangeMarkupElements
/// (bookmarkStart/End, commentRange*). The wrappers nest further cells;
/// non-cell variants are dropped during conversion since they have no
/// rendered effect at row level. Mirrors `TableChildXml`.
#[derive(Deserialize)]
pub(crate) enum RowChildXml {
    /// `<w:tc>` — a table cell.
    #[serde(rename = "tc")]
    Cell(Box<TableCellXml>),
    /// `<w:sdt>` — cell-level structured document tag (CT_SdtCell).
    #[serde(rename = "sdt")]
    Sdt(Box<SdtCellXml>),
    /// `<w:customXml>` — cell-level custom-XML wrapper (CT_CustomXmlCell).
    #[serde(rename = "customXml")]
    CustomXml(Box<CustomXmlCellXml>),
    /// Range markup — bookmarks and comment ranges may span multiple cells.
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(BookmarkStartXml),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(BookmarkEndXml),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(BookmarkEndXml),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(BookmarkEndXml),
    /// Proofreading and permission markers — ignored.
    #[serde(rename = "proofErr")]
    ProofErr(IgnoredXml),
    #[serde(rename = "permStart")]
    PermStart(IgnoredXml),
    #[serde(rename = "permEnd")]
    PermEnd(IgnoredXml),
    /// `<w:tblPrEx>` / `<w:trPr>` are captured on the parent directly; these
    /// variants exist only so `$value` can absorb the duplicate match.
    /// Compare with `TableChildXml::TblPr`.
    #[serde(rename = "tblPrEx")]
    TblPrEx(Box<TblPrExXml>),
    #[serde(rename = "trPr")]
    TrPr(Box<TrPrXml>),
    /// Catch-all for unmodelled OOXML elements (e.g., math content).
    #[serde(other)]
    Other,
}

/// CT_SdtCell §17.5.2 — `<w:sdt>` at cell level. Wrapped cells live in
/// `<w:sdtContent>` (CT_SdtContentCell); the cell-level analogue of
/// `SdtRowXml`.
#[derive(Deserialize, Default)]
pub(crate) struct SdtCellXml {
    #[serde(rename = "sdtContent", default)]
    pub content: Vec<SdtCellContentXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct SdtCellContentXml {
    #[serde(rename = "$value", default)]
    pub children: Vec<RowChildXml>,
}

/// CT_CustomXmlCell §17.5.1 — recursive wrapper whose content model
/// mirrors CT_Row's cell-level choice group, hence `Vec<RowChildXml>`; the
/// cell-level analogue of `CustomXmlRowXml`.
#[derive(Deserialize, Default)]
pub(crate) struct CustomXmlCellXml {
    #[serde(rename = "customXmlPr", default)]
    pub custom_xml_pr: Vec<IgnoredXml>,
    #[serde(rename = "$value", default)]
    pub children: Vec<RowChildXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TableCellXml {
    #[serde(rename = "tcPr", default)]
    pub tc_pr: Vec<TcPrXml>,
    #[serde(rename = "$value", default)]
    pub content: Vec<BlockChildXml>,
}

// ── helpers ────────────────────────────────────────────────────────────────

pub(crate) use crate::docx::parse::primitives::AttrBool;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{PTabAlignment, PTabLeader, PTabRelativeTo, PositionTab};

    fn ptab(xml: &str) -> PositionTab {
        let x: PTabXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn ptab_center_margin() {
        let p = ptab(r#"<ptab relativeTo="margin" alignment="center" leader="none"/>"#);
        assert_eq!(p.alignment, PTabAlignment::Center);
        assert_eq!(p.relative_to, PTabRelativeTo::Margin);
        assert_eq!(p.leader, PTabLeader::None);
    }

    #[test]
    fn ptab_right_margin_with_dot_leader() {
        let p = ptab(r#"<ptab relativeTo="margin" alignment="right" leader="dot"/>"#);
        assert_eq!(p.alignment, PTabAlignment::Right);
        assert_eq!(p.relative_to, PTabRelativeTo::Margin);
        assert_eq!(p.leader, PTabLeader::Dot);
    }

    #[test]
    fn ptab_defaults_are_lenient() {
        // All three attributes are required by the schema; a malformed run
        // that omits them defaults to left / indent / none instead of failing.
        let p = ptab(r#"<ptab/>"#);
        assert_eq!(p.alignment, PTabAlignment::Left);
        assert_eq!(p.relative_to, PTabRelativeTo::Indent);
        assert_eq!(p.leader, PTabLeader::None);
    }

    #[test]
    fn ptab_dispatches_as_run_child() {
        // The whole reason for the fix: `<w:ptab>` must be a recognized run
        // child, not an unknown-variant deserialize error.
        let r: RunXml = quick_xml::de::from_str(
            r#"<r><t>a</t><ptab relativeTo="margin" alignment="right" leader="none"/><t>b</t></r>"#,
        )
        .unwrap();
        assert!(
            r.content.iter().any(|c| matches!(c, RunChildXml::PTab(_))),
            "run content should contain a PTab child"
        );
    }
}
