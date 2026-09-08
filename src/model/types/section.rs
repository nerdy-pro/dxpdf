//! Section properties — page size, margins, columns, headers/footers.

use crate::model::dimension::{Dimension, FractionPoints, Twips};
use crate::model::Dup;

use super::formatting::NumberFormat;
use super::identifiers::{RelId, SectionRevisionIds};

/// §17.6.18 `<w:sectPr>`.
///
/// Unlike the `w:` property bags, this one was never fatal on a repeated
/// child: it deserializes through a `$value` catch-all, which simply
/// overwrote a local. Carrying [`Dup`] here buys losslessness, not tolerance —
/// the earlier occurrences now reach the model instead of being dropped in the
/// fold. See `model::dup`.
#[derive(Clone, Debug, Default)]
pub struct SectionProperties {
    pub page_size: Dup<PageSize>,
    pub page_margins: Dup<PageMargins>,
    pub columns: Dup<Columns>,
    /// §17.6.5: document grid for East Asian typography and line pitch.
    pub doc_grid: Dup<DocGrid>,
    /// §17.6.10: `<w:headerReference>`, one per `@w:type`. Genuinely
    /// repeatable — the spec expects up to three (default, first, even) — so
    /// this is a keyed set, not a [`Dup`]: repetition here is the schema
    /// working, not a producer violating it.
    pub header_refs: SectionHeaderFooterRefs,
    /// §17.6.5: `<w:footerReference>`. See [`SectionProperties::header_refs`].
    pub footer_refs: SectionHeaderFooterRefs,
    pub title_page: Option<bool>,
    /// §17.6.6 `w:bidi`: right-to-left section layout — which side of the
    /// content area is the **leading** one.
    ///
    /// Distinct from §17.4.1 `w:bidiVisual`, which reverses the cells inside a
    /// single table; a document may set either without the other. What reads
    /// this is `render::layout::page::PageConfig::from_section`, which resolves
    /// it to a [`BaseDirection`](crate::i18n::bidi::BaseDirection) — the
    /// engine's existing word for the same question at paragraph level.
    ///
    /// `Option<bool>` rather than `Dup<bool>` to match `title_page`: both are
    /// toggles the `$value` walk above collapses last-wins as it goes.
    pub bidi: Option<bool>,
    pub section_type: Dup<SectionType>,
    /// §17.6.12: page numbering settings for this section.
    pub page_number_type: Dup<PageNumberType>,
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
    /// §17.6.4 `w:sep`: draw a vertical rule between columns.
    ///
    /// Parsed and carried so the document is mirrored faithfully; **not drawn**
    /// — see `render::layout::page::compute_columns`, which records it as a
    /// Tier-0 gap. Column positions are unaffected; only the divider is absent.
    pub separator: Option<bool>,
    /// §17.6.3: individual column definitions. Consulted only when
    /// `equal_width` is explicitly `Some(false)` — §17.6.4 defaults the
    /// attribute to `true`, so an absent value means equal widths.
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
