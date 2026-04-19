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
use crate::docx::parse::primitives::st_enums::{StBrClear, StFldCharType};
use crate::docx::parse::properties::schema::paragraph::PPrXml;
use crate::docx::parse::properties::schema::run::RPrXml;
use crate::docx::parse::properties::schema::section::SectPrXml;
use crate::docx::parse::properties::schema::table::{TblPrXml, TcPrXml, TrPrXml};

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
    pub content: Option<SdtBlockContentXml>,
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
    pub p_pr: Option<PPrXml>,
    #[serde(rename = "$value", default)]
    pub content: Vec<ParaChildXml>,
}

/// Children of `<w:p>` excluding `<w:pPr>` (which is captured separately).
///
/// OOXML allows many annotation and revision-tracking elements at this level
/// (proofErr, smartTag, ins/del, moveFrom/moveTo, commentRangeStart/End,
/// permStart/End, customXml, sdt, ...). We only model the ones we render;
/// the `Other` catch-all lets serde discard everything else cleanly.
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
    /// `<w:pPr>` is captured on `ParaXml` directly, but serde's untagged
    /// enum still has to handle it if it appears in `$value` ordering.
    #[serde(rename = "pPr")]
    PPr(Box<PPrXml>),
    #[serde(other)]
    Other,
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
    pub r_pr: Option<RPrXml>,
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
    #[serde(rename = "br")]
    Br(BrXml),
    #[serde(rename = "cr")]
    Cr,
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
    pub inline: Option<crate::docx::parse::drawing::schema::anchor::InlineXml>,
    #[serde(rename = "anchor", default)]
    pub anchor: Option<crate::docx::parse::drawing::schema::anchor::AnchorXml>,
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
    pub fallback: Option<FallbackXml>,
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
    pub tbl_pr: Option<TblPrXml>,
    #[serde(rename = "tblGrid", default)]
    pub tbl_grid: Option<TblGridXml>,
    #[serde(rename = "tr", default)]
    pub rows: Vec<TableRowXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TblGridXml {
    #[serde(rename = "gridCol", default)]
    pub cols: Vec<GridColXml>,
}

#[derive(Deserialize)]
pub(crate) struct GridColXml {
    #[serde(rename = "@w", default)]
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

    #[serde(rename = "trPr", default)]
    pub tr_pr: Option<TrPrXml>,
    #[serde(rename = "tc", default)]
    pub cells: Vec<TableCellXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TableCellXml {
    #[serde(rename = "tcPr", default)]
    pub tc_pr: Option<TcPrXml>,
    #[serde(rename = "$value", default)]
    pub content: Vec<BlockChildXml>,
}

// ── helpers ────────────────────────────────────────────────────────────────

pub(crate) use crate::docx::parse::primitives::AttrBool;
