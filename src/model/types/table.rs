//! Table types — table, row, cell properties, borders, positioning.

use crate::model::dimension::{Dimension, FiftiethPercent, Twips};
use crate::model::geometry::{EdgeInsets, PartialEdgeInsets};
use crate::model::Dup;

use super::content::Block;
use super::formatting::{
    Alignment, Border, CnfStyle, HeightRule, Shading, TableAnchor, TableXAlign, TableYAlign,
};
use super::identifiers::TableRowRevisionIds;

#[derive(Clone, Debug)]
pub struct Table {
    pub properties: TableProperties,
    pub grid: Vec<GridColumn>,
    pub rows: Vec<TableRow>,
}

#[derive(Clone, Debug, Default)]
pub struct TableProperties {
    /// §17.4.62: table style reference. Stays `Option`: `split` returns it
    /// separately because the cascade applies it *before* direct formatting,
    /// so it is consumed at the seam rather than carried.
    pub style_id: Option<super::identifiers::StyleId>,
    pub alignment: Dup<Alignment>,
    pub width: Dup<TableMeasure>,
    pub layout: Dup<TableLayout>,
    pub indent: Dup<TableMeasure>,
    pub borders: Dup<TableBorders>,
    pub cell_margins: Dup<EdgeInsets<Twips>>,
    pub cell_spacing: Dup<TableMeasure>,
    pub look: Dup<TableLook>,
    /// §17.7.6.7: number of rows in each row band for conditional formatting.
    pub style_row_band_size: Dup<u32>,
    /// §17.7.6.5: number of columns in each column band for conditional formatting.
    pub style_col_band_size: Dup<u32>,
    /// §17.4.57: floating table positioning properties.
    pub positioning: Dup<TablePositioning>,
    /// §17.4.56: whether this floating table can overlap other floating tables.
    pub overlap: Dup<TableOverlap>,
    /// §17.4.1 `w:bidiVisual`: the table's columns display right to left —
    /// the row's first cell is the rightmost one.
    ///
    /// Applied from the `<w:tbl>`'s own `<w:tblPr>` only, like `tblW` and
    /// `tblLook`: [MS-OI29500] §2.1.250(a)/§2.1.249(a) list it among the
    /// elements Word does not read from a style's `tblPr` — see the cascade
    /// note in `render::layout::build::table::build_table`. The style merge
    /// still folds it through `basedOn` with every other field, because
    /// §17.7.4.3 style-to-style inheritance is a different question; a value
    /// both levels carry and no one applies is inert.
    pub bidi_visual: Dup<bool>,
}

/// §17.4.57: floating table positioning.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TablePositioning {
    pub left_from_text: Option<Dimension<Twips>>,
    pub right_from_text: Option<Dimension<Twips>>,
    pub top_from_text: Option<Dimension<Twips>>,
    pub bottom_from_text: Option<Dimension<Twips>>,
    /// §17.18.100: vertical anchor (text, margin, page).
    pub vert_anchor: Option<TableAnchor>,
    /// §17.18.35: horizontal anchor (text, margin, page).
    pub horz_anchor: Option<TableAnchor>,
    /// §22.9.2.18 `ST_XAlign`: horizontal alignment relative to anchor.
    pub x_align: Option<TableXAlign>,
    /// §22.9.2.20 `ST_YAlign`: vertical alignment relative to anchor.
    pub y_align: Option<TableYAlign>,
    /// Absolute horizontal offset from anchor.
    pub x: Option<Dimension<Twips>>,
    /// Absolute vertical offset from anchor.
    pub y: Option<Dimension<Twips>>,
}

/// §17.18.88 `ST_TblOverlap` — floating table overlap behavior, carried by
/// `w:tblOverlap` (§17.4.56).
///
/// This block previously claimed the repo's §17.4.x citations followed
/// ISO/IEC 29500-1 1st Edition, and that MS-OI29500's edition ran exactly one
/// lower — so an off-by-one against that document was the other edition, not a
/// bug. **That is refuted by the repo's own numbers.** An edition-wide shift
/// would move every §17.4 subclause together, but `tblBorders` (§17.4.38),
/// `tblLook` (§17.4.55), `tblHeader` (§17.4.49), `tblW` (§17.4.63) and
/// `tcBorders` (§17.4.66) all agreed with ECMA-376-1:2016 already, while only
/// `tblpPr`/`tblOverlap` sat one high. A shift cannot be selective, so the
/// numbers were simply wrong; they have been corrected against the section
/// headings of the primary spec text. Numbering here is ECMA-376-1:2016
/// throughout, which is also the edition MS-OI29500's `§17.4.x` notes annotate.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableOverlap {
    Overlap,
    Never,
}

/// A dimension for table/cell widths — may be auto, fixed, or percentage.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableMeasure {
    Auto,
    Twips(Dimension<Twips>),
    /// §17.18.90: percentage in 50ths of a percent (OOXML `pct` type).
    Pct(Dimension<FiftiethPercent>),
    /// Nil — explicitly zero.
    Nil,
}

/// §17.18.87 `ST_TblLayoutType`, one variant per enumerated value.
///
/// `Autofit`, not `Auto`: the type's two values are `fixed` and `autofit`, and
/// the `auto` §17.4.52's prose names as the default is a typo for the latter —
/// see
/// [`StTblLayoutType`](crate::docx::parse::primitives::st_enums::StTblLayoutType),
/// which states the evidence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableLayout {
    Autofit,
    Fixed,
}

#[derive(Clone, Copy, Debug)]
pub struct GridColumn {
    pub width: Dimension<Twips>,
}

#[derive(Clone, Debug)]
pub struct TableRow {
    pub properties: TableRowProperties,
    pub cells: Vec<TableCell>,
    pub rsids: TableRowRevisionIds,
    /// §17.4.60 `<w:tblPrEx>` — table-level property *exceptions*
    /// scoped to this row. The spec admits the full vocabulary of
    /// `<w:tblPr>` minus tblStyle/tblpPr; we carry only the slice we
    /// honor today (currently borders) and grow as needed. Values
    /// here override the corresponding entries on the table's
    /// `TableProperties` for this row's cells only.
    pub property_exceptions: Option<TableRowPropertyExceptions>,
}

/// §17.4.60 — the subset of `<w:tblPrEx>` we currently honor. Each
/// field, when `Some`, replaces the matching field on the parent
/// table's `TableProperties` for the row that owns this struct;
/// `None` falls through to the table's value.
#[derive(Clone, Debug, Default)]
pub struct TableRowPropertyExceptions {
    /// Per-row replacement for `TableProperties.borders`.
    pub borders: Option<TableBorders>,
    /// §17.4.44: per-row replacement for `TableProperties.cell_spacing`.
    pub cell_spacing: Option<TableMeasure>,
}

#[derive(Clone, Debug, Default)]
pub struct TableRowProperties {
    pub height: Dup<TableRowHeight>,
    pub is_header: Option<bool>,
    pub cant_split: Option<bool>,
    /// §17.4.27: alignment of the row with respect to text margins (uses ST_Jc).
    pub justification: Dup<Alignment>,
    /// §17.4.7: table conditional formatting applied to this row.
    pub cnf_style: Dup<CnfStyle>,
    /// §17.4.15: number of grid columns to skip before the first cell of the row.
    /// Default 0 when omitted from `<w:trPr>`.
    pub grid_before: u32,
    /// §17.4.43: row-level override of the table's `tblCellSpacing`.
    pub cell_spacing: Dup<TableMeasure>,
    /// §17.4.86: preferred width of the leading space before the first cell.
    /// Note: when present and divergent from the corresponding `tblGrid` columns'
    /// summed widths, treated as informational only — column widths are not
    /// overridden per row in this implementation; a `warn!` is logged on mismatch.
    pub w_before: Dup<TableMeasure>,
    /// §17.4.14: number of grid columns to skip after the last cell of the row.
    /// Default 0 when omitted from `<w:trPr>`.
    pub grid_after: u32,
    /// §17.4.85: preferred width of the trailing space after the last cell.
    /// Same width-override caveat as `w_before`.
    pub w_after: Dup<TableMeasure>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableRowHeight {
    pub value: Dimension<Twips>,
    pub rule: HeightRule,
}

#[derive(Clone, Debug)]
pub struct TableCell {
    pub properties: TableCellProperties,
    pub content: Vec<Block>,
}

/// Table cell properties — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Debug, Default)]
pub struct TableCellProperties {
    pub width: Dup<TableMeasure>,
    pub borders: Dup<TableCellBorders>,
    pub shading: Dup<Shading>,
    /// §17.4.68 `<w:tcMar>` — per-cell margin override. Each side is `Some`
    /// only when explicitly present in the XML; missing sides inherit from
    /// the table-level `<w:tblCellMar>` via `PartialEdgeInsets::resolve_against`.
    pub margins: Dup<PartialEdgeInsets<Twips>>,
    pub vertical_align: Dup<CellVerticalAlign>,
    /// Vertical merge (w:vMerge): absent = not present, else Restart or Continue.
    pub vertical_merge: Dup<VerticalMerge>,
    /// Horizontal span (w:gridSpan): absent = not present, else spans n columns.
    pub grid_span: Dup<u32>,
    pub text_direction: Dup<TextDirection>,
    /// §17.7.2 toggle — collapsed at the seam by `last_toggle`, because for a
    /// toggle last-wins is the spec's own rule rather than this parser's choice.
    pub no_wrap: Option<bool>,
    /// §17.4.8: table conditional formatting applied to this cell.
    pub cnf_style: Dup<CnfStyle>,
}

/// Vertical merge state from `w:vMerge` attribute.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalMerge {
    /// `w:vMerge val="restart"` — this cell starts a new vertical merge group.
    Restart,
    /// `w:vMerge` (no val or val="continue") — this cell continues from above.
    Continue,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellVerticalAlign {
    Top,
    Center,
    Bottom,
    Both,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextDirection {
    LeftToRightTopToBottom,
    TopToBottomRightToLeft,
    BottomToTopLeftToRight,
    LeftToRightTopToBottomRotated,
    TopToBottomRightToLeftRotated,
    TopToBottomLeftToRightRotated,
}

#[derive(Clone, Copy, Debug)]
pub struct TableBorders {
    pub top: Option<Border>,
    pub bottom: Option<Border>,
    pub left: Option<Border>,
    pub right: Option<Border>,
    pub inside_h: Option<Border>,
    pub inside_v: Option<Border>,
}

#[derive(Clone, Copy, Debug)]
pub struct TableCellBorders {
    pub top: Option<Border>,
    pub bottom: Option<Border>,
    pub left: Option<Border>,
    pub right: Option<Border>,
    pub inside_h: Option<Border>,
    pub inside_v: Option<Border>,
    pub tl2br: Option<Border>,
    pub tr2bl: Option<Border>,
}

/// Table conditional formatting flags (ST_TblLook).
#[derive(Clone, Copy, Debug, Default)]
pub struct TableLook {
    pub first_row: Option<bool>,
    pub last_row: Option<bool>,
    pub first_column: Option<bool>,
    pub last_column: Option<bool>,
    pub no_h_band: Option<bool>,
    pub no_v_band: Option<bool>,
}
