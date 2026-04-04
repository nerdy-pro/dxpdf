//! Type definitions for table layout.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtEdgeInsets, PtSize};
use crate::render::resolve::color::RgbColor;

use crate::render::layout::cell::CellLayout;
use crate::render::layout::draw_command::DrawCommand;

/// §17.4.81: row height rule.
#[derive(Clone, Copy, Debug)]
pub enum RowHeightRule {
    /// Row height is at least this value; grows to fit content.
    AtLeast(Pt),
    /// Row height is exactly this value; content may clip.
    Exact(Pt),
}

/// A table row for layout.
pub struct TableRowInput {
    pub cells: Vec<TableCellInput>,
    /// §17.4.81: row height constraint.
    pub height_rule: Option<RowHeightRule>,
    /// §17.4.49: row repeats as header on each continuation page.
    pub is_header: Option<bool>,
    /// §17.4.1: if true, row cannot be split across pages.
    pub cant_split: Option<bool>,
}

/// §17.4.84: cell vertical alignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellVAlign {
    Top,
    Center,
    Bottom,
}

/// A single cell for layout.
pub struct TableCellInput {
    pub blocks: Vec<crate::render::layout::section::LayoutBlock>,
    pub margins: PtEdgeInsets,
    /// Number of grid columns this cell spans (gridSpan, default 1).
    pub grid_span: u32,
    /// Background color for cell shading.
    pub shading: Option<RgbColor>,
    /// §17.7.6: per-cell resolved borders from conditional formatting.
    pub cell_borders: Option<CellBorderConfig>,
    /// §17.4.85: vertical merge state.
    pub vertical_merge: Option<VerticalMergeState>,
    /// §17.4.84: vertical alignment of content within the cell.
    pub vertical_align: CellVAlign,
}

/// §17.4.85: vertical merge state for a cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalMergeState {
    /// This cell starts a new vertical merge group.
    Restart,
    /// This cell continues from the cell above (content is skipped).
    Continue,
}

/// §17.7.6: a conditional border override for a single cell edge.
#[derive(Clone, Copy, Debug)]
pub enum CellBorderOverride {
    /// §17.4.38 val="nil": explicitly no border on this edge.
    Nil,
    /// A specific border line on this edge.
    Border(TableBorderLine),
}

/// Per-cell border configuration (resolved from conditional formatting).
/// `None` = no override (use table-level default for this edge).
#[derive(Clone, Debug)]
pub struct CellBorderConfig {
    pub top: Option<CellBorderOverride>,
    pub bottom: Option<CellBorderOverride>,
    pub left: Option<CellBorderOverride>,
    pub right: Option<CellBorderOverride>,
}

/// Resolved table border configuration.
#[derive(Clone, Debug)]
pub struct TableBorderConfig {
    pub top: Option<TableBorderLine>,
    pub bottom: Option<TableBorderLine>,
    pub left: Option<TableBorderLine>,
    pub right: Option<TableBorderLine>,
    pub inside_h: Option<TableBorderLine>,
    pub inside_v: Option<TableBorderLine>,
}

/// A single table border line.
#[derive(Clone, Copy, Debug)]
pub struct TableBorderLine {
    pub width: Pt,
    pub color: RgbColor,
    /// §17.4.38: border style (single, double, etc.)
    pub style: TableBorderStyle,
}

/// Supported table border styles.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableBorderStyle {
    Single,
    Double,
}

/// Result of laying out a table.
#[derive(Debug)]
pub struct TableLayout {
    /// Draw commands positioned relative to the table's top-left origin.
    pub commands: Vec<DrawCommand>,
    /// Total size of the table.
    pub size: PtSize,
}

/// Per-row measurement data from the table measurement phase.
/// Contains everything needed to emit draw commands for this row.
pub(super) struct MeasuredRow {
    pub(super) entries: Vec<CellLayoutEntry>,
    pub(super) borders: Vec<super::borders::CellBorders>,
    pub(super) height: Pt,
    /// §17.4.38: maximum bottom border width for gap between this row and the next.
    pub(super) border_gap_below: Pt,
}

/// Result of the table measurement phase.
pub(super) struct MeasuredTable {
    pub(super) rows: Vec<MeasuredRow>,
    pub(super) table_width: Pt,
}

/// Per-cell layout result with positioning info from pass 2.
pub(super) struct CellLayoutEntry {
    pub(super) layout: CellLayout,
    pub(super) cell_x: Pt,
    pub(super) cell_w: Pt,
    /// Starting grid column index (for vMerge neighbor lookup).
    pub(super) grid_col: usize,
}
