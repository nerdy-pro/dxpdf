//! Type definitions for table layout.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtEdgeInsets, PtSize};
use crate::render::resolve::color::RgbColor;

use crate::render::layout::cell::CellLayout;
use crate::render::layout::draw_command::DrawCommand;

/// §17.4.80: row height rule.
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
    /// §17.4.80: row height constraint.
    pub height_rule: Option<RowHeightRule>,
    /// §17.4.49: row repeats as header on each continuation page.
    ///
    /// `Option<bool>` and not `bool`, even though `Some(false)` and `None` are
    /// indistinguishable to every reader today — both mean "not a header". The
    /// tri-state is here for the cascade that is not wired yet: §17.7.6 lets a
    /// `<w:tblStylePr w:type="firstRow">` carry a `<w:trPr>`, and once a style
    /// can declare `tblHeader`, a row needs a way to say *explicitly off* and
    /// override it. That is `Some(false)`; `None` is "this level said nothing",
    /// which is what falls through to the style. Collapsing to `bool` now would
    /// erase the distinction and have to be undone then.
    pub is_header: Option<bool>,
    /// §17.4.6: if true, row cannot be split across pages. `Option<bool>` for
    /// the same reason as `is_header` above — `<w:cantSplit>` is a `<w:trPr>`
    /// child, so a table style's conditional `<w:trPr>` can declare it and a row
    /// will need `Some(false)` to turn it back off.
    pub cant_split: Option<bool>,
    /// §17.4.15: number of grid columns to skip at the row's start. The first
    /// cell's leftmost grid column is `grid_before`, not 0.
    ///
    /// There is deliberately no `grid_after` counterpart. §17.4.14 `gridAfter`
    /// is *derivable* — the row's right edge is `grid_before` plus the sum of
    /// its cells' `grid_span`s — and every consumer here works from that
    /// running grid column already, including the border resolution that
    /// decides whether the last cell sits at the table edge (`insideV`) or not.
    /// The model keeps `gridAfter` as parsed (`model::TableRowProperties`);
    /// carrying it into layout added a field nothing read.
    pub grid_before: u32,
    /// §17.4.60 `<w:tblPrEx><w:tblBorders/></w:tblPrEx>` — per-row
    /// override of the table-level `tblBorders`. When set, each side
    /// independently replaces the corresponding side of the table's
    /// `TableBorderConfig` for this row only; sides absent from the
    /// override fall through to the table's value.
    ///
    /// Note that an override side of `None` means "explicitly no
    /// border" (Word's `<w:top w:val="none"/>`) — the parser converts
    /// `val="none"` to `None` here, matching the layout's invariant
    /// that `None` is "draw nothing" everywhere else in this module.
    pub border_overrides: Option<TableBorderConfig>,
}

/// §17.4.83: cell vertical alignment.
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
    /// §17.4.84: vertical merge state.
    pub vertical_merge: Option<VerticalMergeState>,
    /// §17.4.83: vertical alignment of content within the cell.
    pub vertical_align: CellVAlign,
}

/// §17.4.84: vertical merge state for a cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalMergeState {
    /// This cell starts a new vertical merge group.
    Restart,
    /// This cell continues from the cell above (content is skipped).
    Continue,
}

/// §17.7.6: a border override for a single cell edge.
///
/// There is deliberately **no variant for `val="none"`**. Per [MS-OI29500]
/// §17.4.66 an explicit `none` is indistinguishable from an omitted edge — both
/// inherit and both yield — so `convert_cell_border_override` maps it to
/// `Option::None` and it never reaches here. Only `nil`, which suppresses,
/// needs to be represented.
#[derive(Clone, Copy, Debug)]
pub enum CellBorderOverride {
    /// §17.18.2 `val="nil"`: draw nothing on this edge *and* win the conflict
    /// against whatever the adjacent cell declares.
    Suppress,
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
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TableBorderLine {
    pub width: Pt,
    pub color: RgbColor,
    /// §17.18.2: border style (single, double, etc.)
    pub style: TableBorderStyle,
}

/// Supported table border styles.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableBorderStyle {
    Single,
    Double,
}

/// One page-slice of a table.
///
/// A table laid out whole (`layout_table`) is the one-slice case, so it is the
/// same type: there is nothing a monolithic result carries that a slice does
/// not, and two names for one shape only invited the caller to guess which one
/// a given function returns.
#[derive(Debug)]
pub struct TableSlice {
    /// Draw commands positioned relative to this slice's top-left origin (0,0).
    pub commands: Vec<DrawCommand>,
    /// Size of this slice.
    pub size: PtSize,
}

/// Per-row measurement data from the table measurement phase.
/// Contains everything needed to emit draw commands for this row.
pub(super) struct MeasuredRow {
    pub(super) entries: Vec<CellLayoutEntry>,
    pub(super) borders: Vec<super::borders::CellBorders>,
    /// Total vertical space the row owns, **including** `leading_gap`.
    pub(super) height: Pt,
    /// §17.4.45: cell spacing reserved above this row's content, so the gap to
    /// the row above is exactly one `tblCellSpacing`. Zero without spacing,
    /// which is every table that does not set it.
    pub(super) leading_gap: Pt,
    /// §17.4.38: maximum bottom border width for gap between this row and the next.
    pub(super) border_gap_below: Pt,
    /// §17.4.39: `(x0, x1, line)` runs of the boundary **below** this row that
    /// neither it nor the row under it will paint, because a cell paints one
    /// border across its whole width and the two rows break at different
    /// columns. Painted into the strip by `emit`, so the line stays at one y.
    /// Empty for every row of every table whose rows share their column breaks,
    /// which is all but the gapped-against-gapped case.
    pub(super) band_fills: Vec<(Pt, Pt, TableBorderLine)>,
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
