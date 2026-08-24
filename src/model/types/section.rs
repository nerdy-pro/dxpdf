//! Section properties — page size, margins, columns, headers/footers.

use crate::model::dimension::{Dimension, FractionPoints, Points, Twips};
use crate::model::Dup;

use super::formatting::{Border, NumberFormat};
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
    /// §17.6.10 `<w:pgBorders>`: borders around the page.
    pub page_borders: Dup<PageBorders>,
    /// `<w:headerReference>` (§17.10), one per `@w:type`. Genuinely
    /// repeatable — the spec expects up to three (default, first, even) — so
    /// this is a keyed set, not a [`Dup`]: repetition here is the schema
    /// working, not a producer violating it. (The §17.6.10 / §17.6.5 numbers
    /// these two fields used to carry belong to `pgBorders` and `docGrid`;
    /// per AGENTS.md, a disputed number is dropped rather than guessed.)
    pub header_refs: SectionHeaderFooterRefs,
    /// `<w:footerReference>` (§17.10). See [`SectionProperties::header_refs`].
    pub footer_refs: SectionHeaderFooterRefs,
    pub title_page: Option<bool>,
    pub section_type: Dup<SectionType>,
    /// §17.6.12: page numbering settings for this section.
    pub page_number_type: Dup<PageNumberType>,
    pub rsids: SectionRevisionIds,
}

/// §17.6.10 `<w:pgBorders>` — borders around each page of the section.
///
/// The attribute defaults are a Word matter, not a schema one: the strict
/// schema gives none, and [MS-OE376] §2.6.10 records what Word assumes when
/// they are absent — `offsetFrom="text"`, `zOrder="front"`, and the §17.6.10
/// prose itself gives `display="allPages"`. The model keeps them optional and
/// the renderer applies those defaults, so the document is mirrored as
/// written.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PageBorders {
    /// `@w:offsetFrom` — what `space` on each edge is measured from.
    pub offset_from: Option<PageBorderOffset>,
    /// `@w:display` — which pages of the section show the borders.
    pub display: Option<PageBorderDisplay>,
    /// `@w:zOrder` — paint the borders over or under intersecting content.
    pub z_order: Option<PageBorderZOrder>,
    pub top: Option<PageBorderEdge>,
    pub left: Option<PageBorderEdge>,
    pub bottom: Option<PageBorderEdge>,
    pub right: Option<PageBorderEdge>,
}

/// One edge of `<w:pgBorders>`.
///
/// §17.18.2 `ST_Border` is two vocabularies in one type: the 27 line styles
/// every border shares, and the ~165 art names (`apples`…`zigZagStitch`) that
/// are meaningful only on page borders. A line edge reuses [`Border`]; an art
/// edge keeps its name verbatim rather than growing the shared
/// [`super::formatting::BorderStyle`] by 165 variants no other border can
/// carry — and because art `sz` is measured in whole points (1–31), not the
/// line styles' eighths, the two cannot share a width field either.
#[derive(Clone, Debug, PartialEq)]
pub enum PageBorderEdge {
    Line(Border),
    /// An art border, carried losslessly; not rendered.
    Art {
        /// The `@w:val` art name, verbatim.
        name: Box<str>,
        /// `@w:sz` — for art borders, the width in points (1–31).
        width: Option<Dimension<Points>>,
        /// `@w:space` — offset in points, same as a line edge's.
        space: Option<Dimension<Points>>,
    },
}

/// §17.18.63 `ST_PageBorderOffset`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageBorderOffset {
    /// `space` is the distance from the page edge, border drawn inward.
    Page,
    /// `space` is the distance from the text margin, border drawn outward.
    Text,
}

/// §17.18.62 `ST_PageBorderDisplay`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageBorderDisplay {
    AllPages,
    FirstPage,
    NotFirstPage,
}

/// §17.18.64 `ST_PageBorderZOrder`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageBorderZOrder {
    Front,
    Back,
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
