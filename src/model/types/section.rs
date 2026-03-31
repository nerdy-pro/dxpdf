//! Section properties — page size, margins, columns, headers/footers.

use crate::model::dimension::{Dimension, FractionPoints, Twips};

use super::formatting::NumberFormat;
use super::identifiers::{RelId, SectionRevisionIds};

#[derive(Clone, Debug, Default)]
pub struct SectionProperties {
    pub page_size: Option<PageSize>,
    pub page_margins: Option<PageMargins>,
    pub columns: Option<Columns>,
    /// §17.6.5: document grid for East Asian typography and line pitch.
    pub doc_grid: Option<DocGrid>,
    pub header_refs: SectionHeaderFooterRefs,
    pub footer_refs: SectionHeaderFooterRefs,
    pub title_page: Option<bool>,
    pub section_type: Option<SectionType>,
    /// §17.6.12: page numbering settings for this section.
    pub page_number_type: Option<PageNumberType>,
    pub rsids: SectionRevisionIds,
}

/// §17.6.12: page numbering settings.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PageNumberType {
    /// §17.18.59 ST_NumberFormat: page number format.
    pub format: Option<NumberFormat>,
    /// Starting page number (overrides sequential).
    pub start: Option<u32>,
    /// Heading style level for chapter numbering (1-indexed).
    pub chap_style: Option<u32>,
    /// §17.18.6 ST_ChapSep: separator between chapter and page number.
    pub chap_sep: Option<ChapterSeparator>,
}

/// §17.18.6 ST_ChapSep — separator between chapter number and page number.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChapterSeparator {
    Hyphen,
    Period,
    Colon,
    EmDash,
    EnDash,
}

/// §17.6.5: document grid — controls character and line pitch.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct DocGrid {
    /// §17.18.14 ST_DocGrid: type of grid.
    pub grid_type: Option<DocGridType>,
    /// Distance between lines in twips.
    pub line_pitch: Option<Dimension<Twips>>,
    /// Additional character pitch in 4096ths of a point.
    pub char_space: Option<Dimension<FractionPoints>>,
}

/// §17.18.14 ST_DocGrid
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocGridType {
    Default,
    Lines,
    LinesAndChars,
    SnapToChars,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SectionType {
    NextPage,
    Continuous,
    EvenPage,
    OddPage,
    NextColumn,
}

#[derive(Clone, Copy, Debug)]
pub struct PageSize {
    pub width: Option<Dimension<Twips>>,
    pub height: Option<Dimension<Twips>>,
    pub orientation: Option<PageOrientation>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

#[derive(Clone, Copy, Debug)]
pub struct PageMargins {
    pub top: Option<Dimension<Twips>>,
    pub right: Option<Dimension<Twips>>,
    pub bottom: Option<Dimension<Twips>>,
    pub left: Option<Dimension<Twips>>,
    pub header: Option<Dimension<Twips>>,
    pub footer: Option<Dimension<Twips>>,
    pub gutter: Option<Dimension<Twips>>,
}

#[derive(Clone, Debug)]
pub struct Columns {
    pub count: Option<u32>,
    pub space: Option<Dimension<Twips>>,
    pub equal_width: Option<bool>,
    /// §17.6.3: individual column definitions. Empty when `equal_width` is true/absent.
    pub columns: Vec<ColumnDefinition>,
}

/// §17.6.3: a single column definition within a multi-column section.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ColumnDefinition {
    /// Column width in twips.
    pub width: Option<Dimension<Twips>>,
    /// Spacing after this column in twips.
    pub space: Option<Dimension<Twips>>,
}

/// Header/footer references for a section, by position type.
#[derive(Clone, Debug, Default)]
pub struct SectionHeaderFooterRefs {
    pub default: Option<RelId>,
    pub first: Option<RelId>,
    pub even: Option<RelId>,
}
