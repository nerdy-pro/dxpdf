//! Pure data types for section layout.

use super::super::draw_command::{LayoutedPage, ResolvedEffect, ResolvedFill, ResolvedStroke};
use super::super::fragment::Fragment;
use super::super::paragraph::ParagraphStyle;
use super::super::table::TableRowInput;
use crate::model::dimension::{Dimension, SixtieThousandthDeg};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::resolve::images::MediaEntry;
use crate::render::resolve::shape_geometry::SubPath;

/// §17.4.58 / §17.4.59: positioning data for a floating table.
#[derive(Debug, Clone)]
pub struct TableFloatInfo {
    /// Gap between the table's right edge and surrounding text.
    pub right_gap: Pt,
    /// Gap between the table's bottom edge and surrounding text.
    pub bottom_gap: Pt,
    /// §17.4.58: horizontal alignment override (tblpXSpec).
    pub x_align: Option<crate::model::TableXAlign>,
    /// §17.4.59: absolute Y offset from the vertical anchor.
    pub y_offset: Pt,
    /// §17.4.58: vertical anchor reference (text / margin / page).
    pub vert_anchor: crate::model::TableAnchor,
}

/// A floating (anchor) image to be positioned absolutely on the page.
#[derive(Clone)]
pub struct FloatingImage {
    pub image_data: MediaEntry,
    pub size: PtSize,
    /// Resolved absolute x position on the page.
    pub x: Pt,
    /// Resolved absolute y position on the page (may be relative to paragraph).
    pub y: FloatingImageY,
    /// §20.4.2.18: wrapTopAndBottom — text only above/below, not beside.
    pub wrap_top_and_bottom: bool,
    /// §20.4.2.3 distL/distR: horizontal distance from surrounding text.
    pub dist_left: Pt,
    pub dist_right: Pt,
    /// §20.4.2.3 @behindDoc: image is painted behind document text.
    pub behind_doc: bool,
}

/// Vertical position for a floating image.
#[derive(Clone, Copy)]
pub enum FloatingImageY {
    /// Absolute page position.
    Absolute(Pt),
    /// Relative to the paragraph's y position (offset added to cursor_y).
    RelativeToParagraph(Pt),
}

/// A floating (anchor) DrawingML shape to be positioned absolutely on the
/// page. The geometry is already evaluated into path-local Pt subpaths; the
/// painter applies origin + rotation + flip.
///
/// Parallels `FloatingImage` — the stacker treats them the same way for
/// placement, differing only in the emitted `DrawCommand` variant.
#[derive(Clone)]
pub struct FloatingShape {
    /// Resolved absolute x position on the page (top-left of the shape).
    pub x: Pt,
    /// Resolved absolute y position on the page.
    pub y: FloatingImageY,
    /// Shape bounding-box size in Pt.
    pub size: PtSize,
    /// §20.1.7.6 @rot — rotation around the shape's center, in 60000ths
    /// of a degree.
    pub rotation: Dimension<SixtieThousandthDeg>,
    /// §20.1.7.6 @flipH — mirror horizontally.
    pub flip_h: bool,
    /// §20.1.7.6 @flipV — mirror vertically.
    pub flip_v: bool,
    /// §20.4.2.18: wrapTopAndBottom — content only above/below.
    pub wrap_top_and_bottom: bool,
    /// §20.4.2.3 distL/distR — horizontal distance from surrounding text.
    pub dist_left: Pt,
    pub dist_right: Pt,
    /// §20.4.2.3 @behindDoc — painted behind document text.
    pub behind_doc: bool,
    /// Path subpaths in shape-local Pt (already scaled into `size`).
    pub paths: Vec<SubPath>,
    /// Resolved fill.
    pub fill: ResolvedFill,
    /// Optional resolved stroke.
    pub stroke: Option<ResolvedStroke>,
    /// Resolved post-processing effects (painter may defer all in Tier 0).
    pub effects: Vec<ResolvedEffect>,
}

/// A block ready for layout — either a paragraph or a table.
///
/// The `Paragraph` variant is intentionally larger than `Table` (it carries
/// fragments, floats, and resolved style inline). Boxing would add an
/// allocation per block without a measurable benefit for this codebase,
/// where the Vec holding these is the dominant allocation.
#[allow(clippy::large_enum_variant)]
pub enum LayoutBlock {
    Paragraph {
        fragments: Vec<Fragment>,
        style: ParagraphStyle,
        /// §17.3.1.23: force a page break before this paragraph.
        page_break_before: bool,
        /// Footnotes referenced in this paragraph — rendered at page bottom.
        footnotes: Vec<(Vec<Fragment>, ParagraphStyle)>,
        /// §20.4.2.3: floating images anchored to this paragraph.
        floating_images: Vec<FloatingImage>,
        /// §14.5 / §20.1.2.2.35: floating DrawingML shapes anchored to this paragraph.
        floating_shapes: Vec<FloatingShape>,
    },
    Table {
        rows: Vec<TableRowInput>,
        col_widths: Vec<Pt>,
        /// §17.4.38: resolved table border configuration.
        border_config: Option<super::super::table::TableBorderConfig>,
        /// §17.4.51: table indentation from left margin.
        indent: Pt,
        /// §17.4.28: table horizontal alignment.
        alignment: Option<crate::model::Alignment>,
        /// §17.4.58: floating table positioning — if present, text wraps around it.
        float_info: Option<TableFloatInfo>,
        /// §17.4.38: table style reference for adjacent table border collapse.
        style_id: Option<crate::model::StyleId>,
    },
}

/// §17.6.22: continuation state for `Continuous` section breaks.
/// Allows a new section to continue on the current page.
pub struct ContinuationState {
    pub page: LayoutedPage,
    pub cursor_y: Pt,
}
