//! Table types — table, row, cell properties, borders, positioning.

use crate::dimension::{Dimension, ThousandthPercent, Twips};
use crate::geometry::EdgeInsets;

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
    /// §17.4.63: table style reference.
    pub style_id: Option<super::identifiers::StyleId>,
    pub alignment: Option<Alignment>,
    pub width: Option<TableMeasure>,
    pub layout: Option<TableLayout>,
    pub indent: Option<TableMeasure>,
    pub borders: Option<TableBorders>,
    pub cell_margins: Option<EdgeInsets<Twips>>,
    pub cell_spacing: Option<TableMeasure>,
    pub look: Option<TableLook>,
    /// §17.4.68: number of rows in each row band for conditional formatting.
    pub style_row_band_size: Option<u32>,
    /// §17.4.67: number of columns in each column band for conditional formatting.
    pub style_col_band_size: Option<u32>,
    /// §17.4.58: floating table positioning properties.
    pub positioning: Option<TablePositioning>,
    /// §17.4.56: whether this floating table can overlap other floating tables.
    pub overlap: Option<TableOverlap>,
}

/// §17.4.58: floating table positioning.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TablePositioning {
    pub left_from_text: Option<Dimension<Twips>>,
    pub right_from_text: Option<Dimension<Twips>>,
    pub top_from_text: Option<Dimension<Twips>>,
    pub bottom_from_text: Option<Dimension<Twips>>,
    /// §17.18.106: vertical anchor (text, margin, page).
    pub vert_anchor: Option<TableAnchor>,
    /// §17.18.106: horizontal anchor (text, margin, page).
    pub horz_anchor: Option<TableAnchor>,
    /// §17.18.108: horizontal alignment relative to anchor.
    pub x_align: Option<TableXAlign>,
    /// §17.18.109: vertical alignment relative to anchor.
    pub y_align: Option<TableYAlign>,
    /// Absolute horizontal offset from anchor.
    pub x: Option<Dimension<Twips>>,
    /// Absolute vertical offset from anchor.
    pub y: Option<Dimension<Twips>>,
}

/// §17.4.56 ST_TblOverlap — floating table overlap behavior.
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
    /// Percentage in 50ths of a percent (OOXML `pct` type).
    Pct(Dimension<ThousandthPercent>),
    /// Nil — explicitly zero.
    Nil,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableLayout {
    Auto,
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
}

#[derive(Clone, Debug, Default)]
pub struct TableRowProperties {
    pub height: Option<TableRowHeight>,
    pub is_header: Option<bool>,
    pub cant_split: Option<bool>,
    /// §17.4.29: alignment of the row with respect to text margins (uses ST_Jc).
    pub justification: Option<Alignment>,
    /// §17.3.1.8: table conditional formatting applied to this row.
    pub cnf_style: Option<CnfStyle>,
    /// §17.4.14: number of grid columns in the trailing grid units after the last cell.
    pub grid_after: Option<u32>,
    /// §17.4.87: preferred width of the trailing space after the last cell.
    pub w_after: Option<TableMeasure>,
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
    pub width: Option<TableMeasure>,
    pub borders: Option<TableCellBorders>,
    pub shading: Option<Shading>,
    pub margins: Option<EdgeInsets<Twips>>,
    pub vertical_align: Option<CellVerticalAlign>,
    /// Vertical merge (w:vMerge): None = not present, Some(Restart) or Some(Continue).
    pub vertical_merge: Option<VerticalMerge>,
    /// Horizontal span (w:gridSpan): None = not present, Some(n) = spans n columns.
    pub grid_span: Option<u32>,
    pub text_direction: Option<TextDirection>,
    pub no_wrap: Option<bool>,
    /// §17.3.1.8: table conditional formatting applied to this cell.
    pub cnf_style: Option<CnfStyle>,
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
