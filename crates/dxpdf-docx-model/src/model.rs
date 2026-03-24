//! Complete DOCX document model — all types are fully resolved ADTs.
//! No unparsed strings, no style indirection, no invalid states.

use std::collections::HashMap;

use dxpdf_field::FieldInstruction;

use crate::dimension::{
    Dimension, EighthPoints, Emu, FractionPoints, HalfPoints, SixtieThousandthDeg,
    ThousandthPercent, Twips,
};
use crate::geometry::{EdgeInsets, Offset, Size};

// ── Identifiers ──────────────────────────────────────────────────────────────

/// A relationship ID (e.g., "rId1") — opaque, interned from the .rels files.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct RelId(String);

impl RelId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Footnote or endnote numeric ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NoteId(i64);

impl NoteId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// A bookmark ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct BookmarkId(i64);

impl BookmarkId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// A style ID (e.g., "Heading1", "Normal") — reference into `Document.styles`.
/// Per §17.7.4.17, this is the `w:styleId` attribute value.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct StyleId(String);

impl StyleId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// VML shape identifier (e.g., "_x0000_t202"). Used as a `v:shapetype` `id`
/// and referenced by `v:shape` `type` (with a leading `#` prefix).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct VmlShapeId(String);

impl VmlShapeId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Revision Save ID — identifies which editing session produced a change.
/// Stored as a 32-bit value parsed from an 8-digit hex string (e.g., "00A2B3C4").
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RevisionSaveId(u32);

impl RevisionSaveId {
    pub fn value(self) -> u32 {
        self.0
    }

    /// Parse from an OOXML hex string. Returns None if invalid.
    pub fn from_hex(s: &str) -> Option<Self> {
        u32::from_str_radix(s, 16).ok().map(Self)
    }
}

/// Revision tracking IDs attached to an element.
/// Each field records which editing session performed that type of change.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct RevisionIds {
    /// Session that added this element.
    pub r: Option<RevisionSaveId>,
    /// Session that last modified this element's properties.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that deleted this element (for tracked deletions).
    pub del: Option<RevisionSaveId>,
}

/// Revision tracking IDs specific to paragraphs.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct ParagraphRevisionIds {
    /// Session that added this paragraph.
    pub r: Option<RevisionSaveId>,
    /// Session that added the default run content.
    pub r_default: Option<RevisionSaveId>,
    /// Session that last modified paragraph properties.
    pub p: Option<RevisionSaveId>,
    /// Session that last modified run properties on the paragraph mark.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that deleted this paragraph.
    pub del: Option<RevisionSaveId>,
}

/// Revision tracking IDs specific to table rows.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct TableRowRevisionIds {
    /// Session that added this row.
    pub r: Option<RevisionSaveId>,
    /// Session that last modified row properties.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that deleted this row.
    pub del: Option<RevisionSaveId>,
    /// Session that last modified this row's table-level formatting.
    pub tr: Option<RevisionSaveId>,
}

/// Revision tracking IDs specific to section properties.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct SectionRevisionIds {
    /// Session that added this section.
    pub r: Option<RevisionSaveId>,
    /// Session that last modified section run properties.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that last modified section properties.
    pub sect: Option<RevisionSaveId>,
}

// ── Color ────────────────────────────────────────────────────────────────────

/// A color value as specified in the XML (§17.18.5 ST_HexColor).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Color {
    /// sRGB color parsed from a 6-digit hex string (0xRRGGBB).
    Rgb(u32),
    /// The special "auto" color — meaning context-dependent.
    Auto,
}

impl Color {
    pub const BLACK: Self = Self::Rgb(0x000000);
    pub const WHITE: Self = Self::Rgb(0xFFFFFF);
}

// ── Theme ────────────────────────────────────────────────────────────────────

/// Resolved theme data from `theme1.xml`.
#[derive(Clone, Debug, Default)]
pub struct Theme {
    pub color_scheme: ThemeColorScheme,
    pub major_font: ThemeFontScheme,
    pub minor_font: ThemeFontScheme,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct ThemeColorScheme {
    pub dark1: u32,
    pub light1: u32,
    pub dark2: u32,
    pub light2: u32,
    pub accent1: u32,
    pub accent2: u32,
    pub accent3: u32,
    pub accent4: u32,
    pub accent5: u32,
    pub accent6: u32,
    pub hyperlink: u32,
    pub followed_hyperlink: u32,
}

impl ThemeColorScheme {
    /// Resolve a theme color index to an RGB value.
    pub fn resolve(&self, idx: ThemeColorIndex) -> u32 {
        match idx {
            ThemeColorIndex::Dark1 => self.dark1,
            ThemeColorIndex::Light1 => self.light1,
            ThemeColorIndex::Dark2 => self.dark2,
            ThemeColorIndex::Light2 => self.light2,
            ThemeColorIndex::Accent1 => self.accent1,
            ThemeColorIndex::Accent2 => self.accent2,
            ThemeColorIndex::Accent3 => self.accent3,
            ThemeColorIndex::Accent4 => self.accent4,
            ThemeColorIndex::Accent5 => self.accent5,
            ThemeColorIndex::Accent6 => self.accent6,
            ThemeColorIndex::Hyperlink => self.hyperlink,
            ThemeColorIndex::FollowedHyperlink => self.followed_hyperlink,
        }
    }
}

/// Index into the theme color scheme (ST_ThemeColor).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ThemeColorIndex {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

#[derive(Clone, Debug, Default)]
pub struct ThemeFontScheme {
    pub latin: String,
    pub east_asian: String,
    pub complex_script: String,
}

// ── Document ─────────────────────────────────────────────────────────────────

/// The fully parsed and resolved DOCX document.
#[derive(Clone, Debug)]
pub struct Document {
    pub settings: DocumentSettings,
    pub theme: Option<Theme>,
    /// Style definitions from `word/styles.xml`, with `basedOn` references intact.
    pub styles: StyleSheet,
    /// Numbering definitions from `word/numbering.xml`.
    pub numbering: NumberingDefinitions,
    pub body: Vec<Block>,
    /// Final section properties (from the last `w:sectPr` in `w:body`).
    pub final_section: SectionProperties,
    /// Header content keyed by relationship ID.
    pub headers: HashMap<RelId, Vec<Block>>,
    /// Footer content keyed by relationship ID.
    pub footers: HashMap<RelId, Vec<Block>>,
    pub footnotes: HashMap<NoteId, Vec<Block>>,
    pub endnotes: HashMap<NoteId, Vec<Block>>,
    /// Embedded media (images) — raw bytes keyed by relationship ID.
    pub media: HashMap<RelId, Vec<u8>>,
}

// ── Style Sheet ──────────────────────────────────────────────────────────────

/// Raw style definitions as parsed from `word/styles.xml`.
/// No inheritance resolution — `basedOn` references are preserved as-is.
#[derive(Clone, Debug, Default)]
pub struct StyleSheet {
    /// Document-level default paragraph properties (from `w:docDefaults/w:pPrDefault`).
    pub doc_defaults_paragraph: ParagraphProperties,
    /// Document-level default run properties (from `w:docDefaults/w:rPrDefault`).
    pub doc_defaults_run: RunProperties,
    /// All style definitions keyed by style ID.
    pub styles: HashMap<StyleId, Style>,
    /// §17.7.4.5: default properties for styles not explicitly defined.
    pub latent_styles: Option<LatentStyles>,
}

/// §17.7.4.5: default properties for latent (not explicitly defined) styles.
#[derive(Clone, Debug)]
pub struct LatentStyles {
    pub default_locked_state: Option<bool>,
    pub default_ui_priority: Option<u32>,
    pub default_semi_hidden: Option<bool>,
    pub default_unhide_when_used: Option<bool>,
    pub default_q_format: Option<bool>,
    pub count: Option<u32>,
    /// §17.7.4.8: per-style exceptions to the defaults.
    pub exceptions: Vec<LatentStyleException>,
}

/// §17.7.4.8: exception to latent style defaults for a specific style name.
#[derive(Clone, Debug)]
pub struct LatentStyleException {
    pub name: Option<String>,
    pub locked: Option<bool>,
    pub ui_priority: Option<u32>,
    pub semi_hidden: Option<bool>,
    pub unhide_when_used: Option<bool>,
    pub q_format: Option<bool>,
}

/// A single style definition.
#[derive(Clone, Debug)]
pub struct Style {
    pub name: Option<String>,
    pub style_type: StyleType,
    /// Parent style ID. Properties not specified here should be inherited from this style.
    pub based_on: Option<StyleId>,
    pub is_default: bool,
    pub paragraph_properties: Option<ParagraphProperties>,
    pub run_properties: Option<RunProperties>,
    pub table_properties: Option<TableProperties>,
    /// §17.7.6.6: conditional formatting overrides for table regions.
    pub table_style_overrides: Vec<TableStyleOverride>,
}

/// §17.7.6.6: formatting override for a specific table region.
#[derive(Clone, Debug)]
pub struct TableStyleOverride {
    /// Which region this override applies to.
    pub override_type: TableStyleOverrideType,
    pub paragraph_properties: Option<ParagraphProperties>,
    pub run_properties: Option<RunProperties>,
    pub table_properties: Option<TableProperties>,
    pub table_row_properties: Option<TableRowProperties>,
    pub table_cell_properties: Option<TableCellProperties>,
}

/// §17.18.89 ST_TblStyleOverrideType — table region for conditional formatting.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableStyleOverrideType {
    FirstRow,
    LastRow,
    FirstCol,
    LastCol,
    Band1Vert,
    Band2Vert,
    Band1Horz,
    Band2Horz,
    NeCell,
    NwCell,
    SeCell,
    SwCell,
    WholeTable,
}

/// The type of a style definition (§17.7.4.17).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

// ── Numbering Definitions ────────────────────────────────────────────────────

/// Raw numbering definitions as parsed from `word/numbering.xml`.
#[derive(Clone, Debug, Default)]
pub struct NumberingDefinitions {
    /// Abstract numbering definitions keyed by abstract numbering ID.
    pub abstract_nums: HashMap<i64, AbstractNumbering>,
    /// Numbering instances keyed by numbering ID.
    pub numbering_instances: HashMap<i64, NumberingInstance>,
}

/// An abstract numbering definition.
#[derive(Clone, Debug)]
pub struct AbstractNumbering {
    pub levels: Vec<NumberingLevelDefinition>,
}

/// A single level within an abstract numbering definition.
#[derive(Clone, Debug)]
pub struct NumberingLevelDefinition {
    pub level: u8,
    pub format: Option<NumberFormat>,
    pub level_text: String,
    pub start: Option<u32>,
    /// §17.9.7: justification of the numbering symbol (uses ST_Jc).
    pub justification: Option<Alignment>,
    pub indentation: Option<Indentation>,
    pub run_properties: Option<RunProperties>,
}

/// A numbering instance — maps to an abstract numbering, with optional level overrides.
#[derive(Clone, Debug)]
pub struct NumberingInstance {
    pub abstract_num_id: i64,
    pub level_overrides: Vec<NumberingLevelDefinition>,
}

// ── Settings ─────────────────────────────────────────────────────────────────

#[derive(Clone, Debug, Default)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// The rsid of the original editing session that created this document.
    pub rsid_root: Option<RevisionSaveId>,
    /// All revision save IDs recorded in this document's history.
    pub rsids: Vec<RevisionSaveId>,
}

// ── Section ──────────────────────────────────────────────────────────────────

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

// ── Blocks ───────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub enum Block {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
    /// A section break that applies to all preceding content since the last break.
    SectionBreak(Box<SectionProperties>),
}

// ── Paragraph ────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Paragraph {
    /// Style ID reference (e.g., "Heading1"). Resolve via `Document.styles`.
    pub style_id: Option<StyleId>,
    pub properties: ParagraphProperties,
    /// Run properties specified on the paragraph mark (w:rPr inside w:pPr).
    pub mark_run_properties: Option<RunProperties>,
    pub content: Vec<Inline>,
    pub rsids: ParagraphRevisionIds,
}

/// Paragraph properties — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Debug, Default)]
pub struct ParagraphProperties {
    pub alignment: Option<Alignment>,
    pub indentation: Option<Indentation>,
    pub spacing: Option<ParagraphSpacing>,
    pub numbering: Option<NumberingReference>,
    pub tabs: Vec<TabStop>,
    pub borders: Option<ParagraphBorders>,
    pub shading: Option<Shading>,
    pub keep_next: Option<bool>,
    pub keep_lines: Option<bool>,
    pub widow_control: Option<bool>,
    pub page_break_before: Option<bool>,
    pub suppress_auto_hyphens: Option<bool>,
    /// §17.3.1.9: suppress spacing when adjacent paragraph has same style.
    pub contextual_spacing: Option<bool>,
    pub bidi: Option<bool>,
    /// §17.3.1.45: allow line breaking between any characters for East Asian text.
    pub word_wrap: Option<bool>,
    pub outline_level: Option<OutlineLevel>,
    /// §17.3.1.39: vertical alignment of text on each line (ST_TextAlignment).
    pub text_alignment: Option<TextAlignment>,
    /// §17.3.1.8: table conditional formatting applied to this paragraph.
    pub cnf_style: Option<CnfStyle>,
    /// §17.3.1.11: text frame (legacy positioned text region).
    pub frame_properties: Option<FrameProperties>,
}

/// §17.3.1.11: text frame properties — legacy floating positioned text.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FrameProperties {
    /// §17.18.16 ST_DropCap: drop cap type.
    pub drop_cap: Option<DropCap>,
    /// Number of lines the drop cap spans.
    pub lines: Option<u32>,
    /// Frame width in twips.
    pub width: Option<Dimension<Twips>>,
    /// Frame height in twips.
    pub height: Option<Dimension<Twips>>,
    /// §17.18.37 ST_HeightRule: how to interpret the height value.
    pub height_rule: Option<HeightRule>,
    /// Horizontal distance from surrounding text in twips.
    pub h_space: Option<Dimension<Twips>>,
    /// Vertical distance from surrounding text in twips.
    pub v_space: Option<Dimension<Twips>>,
    /// §17.18.104 ST_Wrap: text wrapping mode.
    pub wrap: Option<FrameWrap>,
    /// §17.18.35 ST_HAnchor: horizontal anchor.
    pub h_anchor: Option<TableAnchor>,
    /// §17.18.106 ST_VAnchor: vertical anchor.
    pub v_anchor: Option<TableAnchor>,
    /// Absolute horizontal position in twips.
    pub x: Option<Dimension<Twips>>,
    /// §17.18.108 ST_XAlign: horizontal alignment.
    pub x_align: Option<TableXAlign>,
    /// Absolute vertical position in twips.
    pub y: Option<Dimension<Twips>>,
    /// §17.18.109 ST_YAlign: vertical alignment.
    pub y_align: Option<TableYAlign>,
}

/// §17.18.16 ST_DropCap
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DropCap {
    None,
    Drop,
    Margin,
}

/// §17.18.104 ST_Wrap — text wrapping for frames.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FrameWrap {
    Auto,
    NotBeside,
    Around,
    Tight,
    Through,
    None,
}

/// §17.3.1.8: conditional formatting bit flags indicating which table style
/// regions apply to an element (paragraph, row, or cell).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CnfStyle {
    /// Legacy 12-character binary string (e.g., "101000000000").
    pub val: Option<String>,
    pub first_row: Option<bool>,
    pub last_row: Option<bool>,
    pub first_column: Option<bool>,
    pub last_column: Option<bool>,
    pub odd_v_band: Option<bool>,
    pub even_v_band: Option<bool>,
    pub odd_h_band: Option<bool>,
    pub even_h_band: Option<bool>,
    pub first_row_first_column: Option<bool>,
    pub first_row_last_column: Option<bool>,
    pub last_row_first_column: Option<bool>,
    pub last_row_last_column: Option<bool>,
}

/// §17.18.91 ST_TextAlignment — vertical alignment of characters on a line.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAlignment {
    Auto,
    Top,
    Center,
    Baseline,
    Bottom,
}

/// Heading outline level (0–8, where 0 = Heading 1 in OOXML).
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct OutlineLevel(u8);

impl OutlineLevel {
    /// Create an outline level. Panics if `level` is 0 or > 9.
    pub fn new(level: u8) -> Self {
        assert!((1..=9).contains(&level), "outline level must be 1..=9");
        Self(level)
    }

    /// Create from OOXML raw value (0-based). Returns None if > 8.
    pub fn from_ooxml(val: u8) -> Option<Self> {
        if val <= 8 {
            Some(Self(val + 1))
        } else {
            None
        }
    }

    pub fn value(self) -> u8 {
        self.0
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Alignment {
    Start,
    Center,
    End,
    Both,
    Distribute,
    Thai,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct Indentation {
    pub start: Option<Dimension<Twips>>,
    pub end: Option<Dimension<Twips>>,
    pub first_line: Option<FirstLineIndent>,
    pub mirror: Option<bool>,
}

/// First-line indent: either hanging (negative) or first-line (positive).
/// These are mutually exclusive in OOXML.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FirstLineIndent {
    None,
    FirstLine(Dimension<Twips>),
    Hanging(Dimension<Twips>),
}

/// Paragraph spacing — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct ParagraphSpacing {
    pub before: Option<Dimension<Twips>>,
    pub after: Option<Dimension<Twips>>,
    pub line: Option<LineSpacing>,
    pub before_auto_spacing: Option<bool>,
    pub after_auto_spacing: Option<bool>,
}

/// Line spacing rule — the three OOXML modes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineSpacing {
    /// Automatic (proportional). Value is in 240ths of a line (240 = single).
    Auto(Dimension<Twips>),
    /// Exact line height.
    Exact(Dimension<Twips>),
    /// Minimum line height (at least this much).
    AtLeast(Dimension<Twips>),
}

// ── Numbering ────────────────────────────────────────────────────────────────

/// Raw numbering reference on a paragraph (w:numPr).
/// Resolve via `Document.numbering` using `num_id` + `level`.
#[derive(Clone, Copy, Debug)]
pub struct NumberingReference {
    pub num_id: i64,
    pub level: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Bullet,
    Ordinal,
    CardinalText,
    OrdinalText,
    None,
}

// ── Tabs ─────────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TabStop {
    pub position: Dimension<Twips>,
    pub alignment: TabAlignment,
    pub leader: TabLeader,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TabAlignment {
    Left,
    Center,
    Right,
    Decimal,
    Bar,
    Clear,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

// ── Borders ──────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ParagraphBorders {
    pub top: Option<Border>,
    pub bottom: Option<Border>,
    pub left: Option<Border>,
    pub right: Option<Border>,
    pub between: Option<Border>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Border {
    pub style: BorderStyle,
    pub width: Dimension<EighthPoints>,
    pub space: Dimension<Twips>,
    pub color: Color,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderStyle {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
}

// ── Shading ──────────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Shading {
    pub fill: Color,
    pub pattern: ShadingPattern,
    pub color: Color,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ShadingPattern {
    Clear,
    Solid,
    HorzStripe,
    VertStripe,
    ReverseDiagStripe,
    DiagStripe,
    HorzCross,
    DiagCross,
    ThinHorzStripe,
    ThinVertStripe,
    ThinReverseDiagStripe,
    ThinDiagStripe,
    ThinHorzCross,
    ThinDiagCross,
    Pct5,
    Pct10,
    Pct12,
    Pct15,
    Pct20,
    Pct25,
    Pct30,
    Pct35,
    Pct37,
    Pct40,
    Pct45,
    Pct50,
    Pct55,
    Pct60,
    Pct62,
    Pct65,
    Pct70,
    Pct75,
    Pct80,
    Pct85,
    Pct87,
    Pct90,
    Pct95,
}

// ── Inline content ───────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub enum Inline {
    TextRun(Box<TextRun>),
    Tab,
    LineBreak(BreakKind),
    ColumnBreak,
    PageBreak,
    /// §17.3.3.13: position where the previous application rendered a page break.
    /// This is a rendering hint, not a content break.
    LastRenderedPageBreak,
    Image(Box<Image>),
    FootnoteRef(NoteId),
    EndnoteRef(NoteId),
    Hyperlink(Hyperlink),
    Field(Field),
    BookmarkStart {
        id: BookmarkId,
        name: String,
    },
    BookmarkEnd(BookmarkId),
    Symbol(Symbol),
    /// §17.11.23: footnote/endnote separator line.
    Separator,
    /// §17.11.3: continuation separator for notes spanning pages.
    ContinuationSeparator,
    /// §17.16.18: complex field character (begin/separate/end marker).
    FieldChar(FieldChar),
    /// §17.16.23: field instruction text (appears between begin and separate).
    InstrText(String),
    /// §17.11.13: footnote reference mark (auto-number rendered in the footnote body).
    FootnoteRefMark,
    /// §17.11.6: endnote reference mark (auto-number rendered in the endnote body).
    EndnoteRefMark,
    /// §17.3.3.19: VML picture/shape container (legacy drawing).
    Pict(Pict),
    /// MCE §M.2.1: markup compatibility alternate content.
    AlternateContent(AlternateContent),
}

/// §17.16.18: complex field character marker.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FieldChar {
    /// §17.18.29 ST_FldCharType: begin, separate, or end.
    pub field_char_type: FieldCharType,
    /// Field result needs recalculation.
    pub dirty: Option<bool>,
    /// Field is locked from updates.
    pub fld_lock: Option<bool>,
}

/// §17.18.29 ST_FldCharType
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FieldCharType {
    Begin,
    Separate,
    End,
}

#[derive(Clone, Debug)]
pub struct TextRun {
    /// Character style ID reference (e.g., "Hyperlink"). Resolve via `Document.styles`.
    pub style_id: Option<StyleId>,
    pub properties: RunProperties,
    pub text: String,
    pub rsids: RevisionIds,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BreakKind {
    TextWrapping,
    /// Clears left, right, or both float areas.
    Clear(BreakClear),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BreakClear {
    None,
    Left,
    Right,
    All,
}

/// A symbol character from a specific font (w:sym).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Symbol {
    pub font: String,
    pub char_code: u16,
}

// ── Run Properties ───────────────────────────────────────────────────────────

/// Run properties — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Debug, PartialEq, Default)]
pub struct RunProperties {
    pub fonts: FontSet,
    pub font_size: Option<Dimension<HalfPoints>>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<UnderlineStyle>,
    pub strike: Option<StrikeStyle>,
    pub color: Option<Color>,
    pub highlight: Option<HighlightColor>,
    pub shading: Option<Shading>,
    pub vertical_align: Option<VerticalAlign>,
    pub spacing: Option<Dimension<Twips>>,
    pub kerning: Option<Dimension<HalfPoints>>,
    pub all_caps: Option<bool>,
    pub small_caps: Option<bool>,
    pub vanish: Option<bool>,
    /// §17.3.2.21: suppress spell/grammar checking for this run.
    pub no_proof: Option<bool>,
    /// §17.3.2.44: hidden when displayed as a web page, visible in print view.
    pub web_hidden: Option<bool>,
    pub rtl: Option<bool>,
    pub emboss: Option<bool>,
    pub imprint: Option<bool>,
    pub outline: Option<bool>,
    pub shadow: Option<bool>,
    /// §17.3.2.19: vertical position offset of text baseline, in half-points.
    /// Positive raises, negative lowers.
    pub position: Option<Dimension<HalfPoints>>,
    /// §17.3.2.20: proofing languages per script category (BCP 47 tags).
    pub lang: Option<Lang>,
}

/// §17.3.2.20: proofing language specification per script category.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Lang {
    /// Language for Latin text (e.g., "en-US").
    pub val: Option<String>,
    /// Language for East Asian text (e.g., "zh-CN").
    pub east_asia: Option<String>,
    /// Language for complex script text (e.g., "ar-SA").
    pub bidi: Option<String>,
}

/// Font family names for each script category.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct FontSet {
    pub ascii: Option<String>,
    pub high_ansi: Option<String>,
    pub east_asian: Option<String>,
    pub complex_script: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnderlineStyle {
    None,
    Single,
    Words,
    Double,
    Thick,
    Dotted,
    DottedHeavy,
    Dash,
    DashedHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DashDotHeavy,
    DotDotDash,
    DashDotDotHeavy,
    Wave,
    WavyHeavy,
    WavyDouble,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StrikeStyle {
    None,
    Single,
    Double,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

/// Highlight colors — fixed palette per OOXML spec (ST_HighlightColor).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HighlightColor {
    Black,
    Blue,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGray,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    Green,
    LightGray,
    Magenta,
    Red,
    White,
    Yellow,
}

// ── Image / Drawing ──────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Image {
    /// §20.4.2.7: drawing extent.
    pub extent: Size<Emu>,
    /// §20.4.2.6: additional extent for effects.
    pub effect_extent: Option<EdgeInsets<Emu>>,
    /// §20.1.2.2.8: non-visual drawing properties (wp:docPr).
    pub doc_properties: DocProperties,
    /// §20.4.2.4: graphic frame locking properties.
    pub graphic_frame_locks: Option<GraphicFrameLocks>,
    /// §19.3.1.37: picture content (a:graphic > a:graphicData > pic:pic).
    pub picture: Option<Picture>,
    /// Inline or anchor placement.
    pub placement: ImagePlacement,
}

/// How the image is placed in the document flow.
#[derive(Clone, Debug)]
pub enum ImagePlacement {
    /// §20.4.2.8: inline with text — no wrapping.
    Inline {
        /// Distance from surrounding text.
        distance: EdgeInsets<Emu>,
    },
    /// §20.4.2.3: floating/anchored with text wrapping.
    Anchor(AnchorProperties),
}

/// §20.1.2.2.8 CT_NonVisualDrawingProps — shared by wp:docPr and pic:cNvPr.
#[derive(Clone, Debug)]
pub struct DocProperties {
    /// Unique identifier.
    pub id: u32,
    /// Element name.
    pub name: String,
    /// Alternative text description.
    pub description: Option<String>,
    /// Whether the element is hidden.
    pub hidden: Option<bool>,
    /// Title (Office 2010+).
    pub title: Option<String>,
}

/// §20.1.2.2.19 CT_GraphicalObjectFrameLocking.
#[derive(Clone, Copy, Debug)]
pub struct GraphicFrameLocks {
    pub no_change_aspect: Option<bool>,
    pub no_drilldown: Option<bool>,
    pub no_grp: Option<bool>,
    pub no_move: Option<bool>,
    pub no_resize: Option<bool>,
    pub no_select: Option<bool>,
}

/// §19.3.1.37 pic:pic — a picture element.
#[derive(Clone, Debug)]
pub struct Picture {
    /// §19.3.1.32: non-visual picture properties.
    pub nv_pic_pr: NvPicProperties,
    /// §20.1.8.14: blip fill (picture data + crop + fill mode).
    pub blip_fill: BlipFill,
    /// §20.1.2.2.35: shape properties (transform, geometry, outline).
    pub shape_properties: Option<ShapeProperties>,
}

/// §19.3.1.32 pic:nvPicPr — non-visual picture properties.
#[derive(Clone, Debug)]
pub struct NvPicProperties {
    /// §20.1.2.2.8 pic:cNvPr.
    pub cnv_pr: DocProperties,
    /// §19.3.1.4 pic:cNvPicPr.
    pub cnv_pic_pr: Option<CnvPicProperties>,
}

/// §19.3.1.4 pic:cNvPicPr — non-visual picture drawing properties.
#[derive(Clone, Debug)]
pub struct CnvPicProperties {
    pub prefer_relative_resize: Option<bool>,
    /// §20.1.2.2.31: picture locking.
    pub pic_locks: Option<PicLocks>,
}

/// §20.1.2.2.31 a:picLocks — picture locking constraints.
#[derive(Clone, Copy, Debug)]
pub struct PicLocks {
    pub no_change_aspect: Option<bool>,
    pub no_crop: Option<bool>,
    pub no_resize: Option<bool>,
    pub no_move: Option<bool>,
    pub no_rot: Option<bool>,
    pub no_select: Option<bool>,
    pub no_edit_points: Option<bool>,
    pub no_adjust_handles: Option<bool>,
    pub no_change_arrowheads: Option<bool>,
    pub no_change_shape_type: Option<bool>,
    pub no_grp: Option<bool>,
}

/// §20.1.8.14 pic:blipFill — picture fill properties.
#[derive(Clone, Debug)]
pub struct BlipFill {
    pub rotate_with_shape: Option<bool>,
    pub dpi: Option<u32>,
    /// §20.1.8.13: blip reference.
    pub blip: Option<Blip>,
    /// §20.1.10.48: source rectangle (crop).
    pub src_rect: Option<RelativeRect>,
    /// §20.1.8.56: stretch fill mode.
    pub stretch: Option<StretchFill>,
}

/// §20.1.8.13 a:blip — reference to image data.
#[derive(Clone, Debug)]
pub struct Blip {
    /// r:embed — relationship ID for embedded image.
    pub embed: Option<RelId>,
    /// r:link — relationship ID for linked (external) image.
    pub link: Option<RelId>,
    /// §20.1.10.7: compression state.
    pub compression: Option<BlipCompression>,
}

/// §20.1.10.7 ST_BlipCompression.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BlipCompression {
    Email,
    Hqprint,
    None,
    Print,
    Screen,
}

/// §20.1.10.48 CT_RelativeRect — relative rectangle (thousandths of percent).
/// Used for a:srcRect and a:fillRect.
#[derive(Clone, Copy, Debug)]
pub struct RelativeRect {
    pub left: Option<Dimension<ThousandthPercent>>,
    pub top: Option<Dimension<ThousandthPercent>>,
    pub right: Option<Dimension<ThousandthPercent>>,
    pub bottom: Option<Dimension<ThousandthPercent>>,
}

/// §20.1.8.56 a:stretch — stretch fill mode.
#[derive(Clone, Copy, Debug)]
pub struct StretchFill {
    /// §20.1.10.48: fill rectangle.
    pub fill_rect: Option<RelativeRect>,
}

/// §20.1.2.2.35 CT_ShapeProperties — shape visual properties.
#[derive(Clone, Debug)]
pub struct ShapeProperties {
    /// §20.1.10.10: black-and-white mode.
    pub bw_mode: Option<BlackWhiteMode>,
    /// §20.1.7.6: 2D transform.
    pub transform: Option<Transform2D>,
    /// §20.1.9.18: preset geometry.
    pub preset_geometry: Option<PresetGeometryDef>,
    /// Fill type (noFill, solidFill, etc.).
    pub fill: Option<DrawingFill>,
    /// §20.1.2.2.24: outline/line properties.
    pub outline: Option<Outline>,
}

/// §20.1.7.6 CT_Transform2D — 2D transform.
#[derive(Clone, Copy, Debug)]
pub struct Transform2D {
    /// Rotation in 60,000ths of a degree (§20.1.10.3).
    pub rotation: Option<Dimension<SixtieThousandthDeg>>,
    pub flip_h: Option<bool>,
    pub flip_v: Option<bool>,
    /// §20.1.7.4: offset (x, y).
    pub offset: Option<Offset<Emu>>,
    /// §20.1.7.3: extent (cx, cy).
    pub extent: Option<Size<Emu>>,
}

/// §20.1.9.18 CT_PresetGeometry2D — preset shape geometry.
#[derive(Clone, Debug)]
pub struct PresetGeometryDef {
    /// §20.1.10.56: preset shape type.
    pub preset: PresetShapeType,
    /// §20.1.9.5: adjustment values.
    pub adjust_values: Vec<GeomGuide>,
}

/// §20.1.9.11 CT_GeomGuide — geometry guide (named formula).
#[derive(Clone, Debug)]
pub struct GeomGuide {
    pub name: String,
    pub formula: String,
}

/// §20.1.10.56 ST_ShapeType — preset shape types (subset).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PresetShapeType {
    Rect,
    RoundRect,
    Ellipse,
    Triangle,
    RtTriangle,
    Diamond,
    Parallelogram,
    Trapezoid,
    Pentagon,
    Hexagon,
    Octagon,
    Star4,
    Star5,
    Star6,
    Star8,
    Star10,
    Star12,
    Star16,
    Star24,
    Star32,
    Line,
    Plus,
    Can,
    Cube,
    Donut,
    NoSmoking,
    BlockArc,
    Heart,
    Sun,
    Moon,
    SmileyFace,
    LightningBolt,
    Cloud,
    Arc,
    Plaque,
    Frame,
    Bevel,
    FoldedCorner,
    Chevron,
    HomePlate,
    Ribbon,
    Ribbon2,
    Pie,
    PieWedge,
    Chord,
    Teardrop,
    Arrow,
    LeftArrow,
    RightArrow,
    UpArrow,
    DownArrow,
    LeftRightArrow,
    UpDownArrow,
    QuadArrow,
    BentArrow,
    UturnArrow,
    CircularArrow,
    CurvedRightArrow,
    CurvedLeftArrow,
    CurvedUpArrow,
    CurvedDownArrow,
    StripedRightArrow,
    NotchedRightArrow,
    BentUpArrow,
    LeftUpArrow,
    LeftRightUpArrow,
    LeftArrowCallout,
    RightArrowCallout,
    UpArrowCallout,
    DownArrowCallout,
    LeftRightArrowCallout,
    UpDownArrowCallout,
    QuadArrowCallout,
    SwooshArrow,
    LeftCircularArrow,
    LeftRightCircularArrow,
    Callout1,
    Callout2,
    Callout3,
    AccentCallout1,
    AccentCallout2,
    AccentCallout3,
    BorderCallout1,
    BorderCallout2,
    BorderCallout3,
    AccentBorderCallout1,
    AccentBorderCallout2,
    AccentBorderCallout3,
    WedgeRectCallout,
    WedgeRoundRectCallout,
    WedgeEllipseCallout,
    CloudCallout,
    LeftBracket,
    RightBracket,
    LeftBrace,
    RightBrace,
    BracketPair,
    BracePair,
    StraightConnector1,
    BentConnector2,
    BentConnector3,
    BentConnector4,
    BentConnector5,
    CurvedConnector2,
    CurvedConnector3,
    CurvedConnector4,
    CurvedConnector5,
    FlowChartProcess,
    FlowChartDecision,
    FlowChartInputOutput,
    FlowChartPredefinedProcess,
    FlowChartInternalStorage,
    FlowChartDocument,
    FlowChartMultidocument,
    FlowChartTerminator,
    FlowChartPreparation,
    FlowChartManualInput,
    FlowChartManualOperation,
    FlowChartConnector,
    FlowChartPunchedCard,
    FlowChartPunchedTape,
    FlowChartSummingJunction,
    FlowChartOr,
    FlowChartCollate,
    FlowChartSort,
    FlowChartExtract,
    FlowChartMerge,
    FlowChartOfflineStorage,
    FlowChartOnlineStorage,
    FlowChartMagneticTape,
    FlowChartMagneticDisk,
    FlowChartMagneticDrum,
    FlowChartDisplay,
    FlowChartDelay,
    FlowChartAlternateProcess,
    FlowChartOffpageConnector,
    ActionButtonBlank,
    ActionButtonHome,
    ActionButtonHelp,
    ActionButtonInformation,
    ActionButtonForwardNext,
    ActionButtonBackPrevious,
    ActionButtonEnd,
    ActionButtonBeginning,
    ActionButtonReturn,
    ActionButtonDocument,
    ActionButtonSound,
    ActionButtonMovie,
    IrregularSeal1,
    IrregularSeal2,
    Wave,
    DoubleWave,
    EllipseRibbon,
    EllipseRibbon2,
    VerticalScroll,
    HorizontalScroll,
    LeftRightRibbon,
    Gear6,
    Gear9,
    Funnel,
    MathPlus,
    MathMinus,
    MathMultiply,
    MathDivide,
    MathEqual,
    MathNotEqual,
    CornerTabs,
    SquareTabs,
    PlaqueTabs,
    ChartX,
    ChartStar,
    ChartPlus,
    HalfFrame,
    Corner,
    DiagStripe,
    NonIsoscelesTrapezoid,
    Heptagon,
    Decagon,
    Dodecagon,
    Round1Rect,
    Round2SameRect,
    Round2DiagRect,
    SnipRoundRect,
    Snip1Rect,
    Snip2SameRect,
    Snip2DiagRect,
    /// Unrecognized shape type — preserved as raw string.
    Other(String),
}

/// §20.1.10.10 ST_BlackWhiteMode.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BlackWhiteMode {
    Auto,
    Black,
    BlackGray,
    BlackWhite,
    Clr,
    Gray,
    GrayWhite,
    Hidden,
    InvGray,
    LtGray,
    White,
}

/// Drawing fill type.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DrawingFill {
    /// §20.1.8.44: no fill.
    NoFill,
    // SolidFill, GradFill, etc. — extend as needed.
}

/// §20.1.2.2.24 CT_LineProperties — outline/line properties.
#[derive(Clone, Debug)]
pub struct Outline {
    /// Line width in EMUs.
    pub width: Option<Dimension<Emu>>,
    /// §20.1.10.31: line cap style.
    pub cap: Option<LineCap>,
    /// §20.1.10.15: compound line type.
    pub compound: Option<CompoundLine>,
    /// §20.1.10.39: pen alignment.
    pub alignment: Option<PenAlignment>,
    /// Line fill.
    pub fill: Option<DrawingFill>,
}

/// §20.1.10.31 ST_LineCap.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineCap {
    Flat,
    Round,
    Square,
}

/// §20.1.10.15 ST_CompoundLine.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CompoundLine {
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple,
}

/// §20.1.10.39 ST_PenAlignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PenAlignment {
    Center,
    Inset,
}

/// §20.4.2.3 CT_Anchor — anchor/floating drawing properties.
#[derive(Clone, Copy, Debug)]
pub struct AnchorProperties {
    /// §20.4.2.3: distance from surrounding text.
    pub distance: EdgeInsets<Emu>,
    /// §20.4.2.13: simple positioning point.
    pub simple_pos: Option<Offset<Emu>>,
    /// §20.4.2.3 @simplePos: whether to use simplePos coordinates.
    pub use_simple_pos: Option<bool>,
    /// §20.4.2.10: horizontal position.
    pub horizontal_position: AnchorPosition,
    /// §20.4.2.11: vertical position.
    pub vertical_position: AnchorPosition,
    /// Text wrapping mode.
    pub wrap: TextWrap,
    /// §20.4.2.3 @behindDoc: behind document text.
    pub behind_text: bool,
    /// §20.4.2.3 @locked: anchor is locked to position.
    pub lock_anchor: bool,
    /// §20.4.2.3 @allowOverlap: can overlap other anchored objects.
    pub allow_overlap: bool,
    /// §20.4.2.3 @relativeHeight: z-ordering value.
    pub relative_height: u32,
    /// §20.4.2.3 @layoutInCell: allow layout inside table cell.
    pub layout_in_cell: Option<bool>,
    /// §20.4.2.3 @hidden: whether the anchor is hidden.
    pub hidden: Option<bool>,
}

/// §20.4.2.10 / §20.4.2.11: anchor position (offset or alignment).
#[derive(Clone, Copy, Debug)]
pub enum AnchorPosition {
    /// §20.4.2.12: position by EMU offset from relativeFrom.
    Offset {
        relative_from: AnchorRelativeFrom,
        offset: Dimension<Emu>,
    },
    /// §20.4.2.1: position by alignment within relativeFrom.
    Align {
        relative_from: AnchorRelativeFrom,
        alignment: AnchorAlignment,
    },
}

/// §20.4.3.4 ST_RelFromH / §20.4.3.5 ST_RelFromV — relative-from values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AnchorRelativeFrom {
    Page,
    Margin,
    Column,
    Character,
    Paragraph,
    Line,
    InsideMargin,
    OutsideMargin,
    TopMargin,
    BottomMargin,
    LeftMargin,
    RightMargin,
}

/// §20.4.3.1 ST_AlignH / §20.4.3.2 ST_AlignV — alignment values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AnchorAlignment {
    Left,
    Center,
    Right,
    Inside,
    Outside,
    Top,
    Bottom,
}

/// Text wrapping mode for anchored drawings.
#[derive(Clone, Copy, Debug)]
pub enum TextWrap {
    /// §20.4.2.15: no wrapping.
    None,
    /// §20.4.2.17: square wrapping.
    Square {
        distance: EdgeInsets<Emu>,
        wrap_text: WrapText,
    },
    /// §20.4.2.16: tight wrapping.
    Tight {
        distance: EdgeInsets<Emu>,
        wrap_text: WrapText,
    },
    /// §20.4.2.18: text above and below only.
    TopAndBottom {
        distance_top: Dimension<Emu>,
        distance_bottom: Dimension<Emu>,
    },
    /// §20.4.2.14: through wrapping.
    Through {
        distance: EdgeInsets<Emu>,
        wrap_text: WrapText,
    },
}

/// §20.4.3.7 ST_WrapText — which sides text wraps on.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WrapText {
    BothSides,
    Left,
    Right,
    Largest,
}

// ── Hyperlink ────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Hyperlink {
    pub target: HyperlinkTarget,
    pub content: Vec<Inline>,
}

/// A hyperlink either targets an external URL (via relationship) or an internal bookmark.
#[derive(Clone, Debug)]
pub enum HyperlinkTarget {
    External(RelId),
    Internal { anchor: String },
}

// ── Field ────────────────────────────────────────────────────────────────────

/// A simple field (w:fldSimple). Stores the parsed field instruction.
#[derive(Clone, Debug)]
pub struct Field {
    /// Parsed field instruction (e.g., `FieldInstruction::Page`, `FieldInstruction::Toc { .. }`).
    pub instruction: FieldInstruction,
    pub content: Vec<Inline>,
}

// ── VML / Pict ──────────────────────────────────────────────────────────────

/// §17.3.3.19: VML picture container. Wraps legacy VML shape content.
#[derive(Clone, Debug)]
pub struct Pict {
    /// VML §14.1.2.20: optional reusable shape type definition.
    pub shape_type: Option<VmlShapeType>,
    /// VML §14.1.2.19: shape instances.
    pub shapes: Vec<VmlShape>,
}

/// VML §14.2.1.6: a single path command in the shape path language.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlPathCommand {
    /// `m x,y` — move to absolute position.
    MoveTo { x: i64, y: i64 },
    /// `l x,y` — line to absolute position.
    LineTo { x: i64, y: i64 },
    /// `c x1,y1,x2,y2,x,y` — cubic bezier to absolute position.
    CurveTo {
        x1: i64,
        y1: i64,
        x2: i64,
        y2: i64,
        x: i64,
        y: i64,
    },
    /// `r dx,dy` — relative line to.
    RLineTo { dx: i64, dy: i64 },
    /// `v dx1,dy1,dx2,dy2,dx,dy` — relative cubic bezier.
    RCurveTo {
        dx1: i64,
        dy1: i64,
        dx2: i64,
        dy2: i64,
        dx: i64,
        dy: i64,
    },
    /// `t dx,dy` — relative move to.
    RMoveTo { dx: i64, dy: i64 },
    /// `x` — close subpath.
    Close,
    /// `e` — end of path.
    End,
    /// `qx x,y` — elliptical quadrant, x-axis first.
    QuadrantX { x: i64, y: i64 },
    /// `qy x,y` — elliptical quadrant, y-axis first.
    QuadrantY { x: i64, y: i64 },
    /// `nf` — no fill for following subpath.
    NoFill,
    /// `ns` — no stroke for following subpath.
    NoStroke,
    /// `wa/wr/at/ar x1,y1,x2,y2,x3,y3,x4,y4` — arc commands.
    Arc {
        /// `wa` (angle clockwise), `wr` (angle counter-clockwise),
        /// `at` (to clockwise), `ar` (to counter-clockwise).
        kind: VmlArcKind,
        bounding_x1: i64,
        bounding_y1: i64,
        bounding_x2: i64,
        bounding_y2: i64,
        start_x: i64,
        start_y: i64,
        end_x: i64,
        end_y: i64,
    },
}

/// VML §14.2.1.6: arc sub-type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlArcKind {
    /// `wa` — clockwise arc (angle).
    WA,
    /// `wr` — counter-clockwise arc (angle).
    WR,
    /// `at` — clockwise arc (to point).
    AT,
    /// `ar` — counter-clockwise arc (to point).
    AR,
}

/// VML §14.1.2.6: a single formula equation (`v:f eqn="..."`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VmlFormula {
    pub operation: VmlFormulaOp,
    pub args: [VmlFormulaArg; 3],
}

/// VML §14.1.2.6: formula operations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlFormulaOp {
    /// `val` — returns arg1.
    Val,
    /// `sum` — arg1 + arg2 - arg3.
    Sum,
    /// `product` — arg1 * arg2 / arg3.
    Product,
    /// `mid` — (arg1 + arg2) / 2.
    Mid,
    /// `abs` — |arg1|.
    Abs,
    /// `min` — min(arg1, arg2).
    Min,
    /// `max` — max(arg1, arg2).
    Max,
    /// `if` — if arg1 > 0 then arg2 else arg3.
    If,
    /// `sqrt` — sqrt(arg1).
    Sqrt,
    /// `mod` — sqrt(arg1² + arg2² + arg3²).
    Mod,
    /// `sin` — arg1 * sin(arg2).
    Sin,
    /// `cos` — arg1 * cos(arg2).
    Cos,
    /// `tan` — arg1 * tan(arg2).
    Tan,
    /// `atan2` — atan2(arg1, arg2) in fd units.
    Atan2,
    /// `sinatan2` — arg1 * sin(atan2(arg2, arg3)).
    SinAtan2,
    /// `cosatan2` — arg1 * cos(atan2(arg2, arg3)).
    CosAtan2,
    /// `sumangle` — arg1 + arg2 * 2^16 - arg3 * 2^16 (angle arithmetic).
    SumAngle,
    /// `ellipse` — arg3 * sqrt(1 - (arg1/arg2)²).
    Ellipse,
}

/// VML §14.1.2.6: formula argument — a reference or literal.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlFormulaArg {
    /// Integer literal.
    Literal(i64),
    /// `#n` — adjustment value reference.
    AdjRef(u32),
    /// `@n` — formula result reference.
    FormulaRef(u32),
    /// Named guide value (width, height, xcenter, ycenter, etc.).
    Guide(VmlGuide),
}

/// VML §14.1.2.6: named guide constants available in formulas.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlGuide {
    Width,
    Height,
    XCenter,
    YCenter,
    XRange,
    YRange,
    PixelWidth,
    PixelHeight,
    PixelLineWidth,
    EmuWidth,
    EmuHeight,
    EmuWidth2,
    EmuHeight2,
}

/// VML Vector2D — unitless coordinate pair (e.g., coordsize="21600,21600").
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct VmlVector2D {
    pub x: i64,
    pub y: i64,
}

/// VML §14.1.2.20: shape type definition (reusable template for shapes).
#[derive(Clone, Debug)]
pub struct VmlShapeType {
    /// Shape type identifier (e.g., "_x0000_t202").
    pub id: Option<VmlShapeId>,
    /// Coordinate space for the shape (VML Vector2D, e.g., 21600,21600).
    pub coord_size: Option<VmlVector2D>,
    /// o:spt — shape type number (Office extension, xsd:float).
    pub spt: Option<f32>,
    /// Adjustment values for parameterized shapes (comma-separated integers).
    pub adj: Vec<i64>,
    /// VML §14.2.1.6: parsed shape path commands.
    pub path: Vec<VmlPathCommand>,
    /// Whether the shape is filled by default.
    pub filled: Option<bool>,
    /// Whether the shape is stroked by default.
    pub stroked: Option<bool>,
    /// VML §14.1.2.21: stroke child element.
    pub stroke: Option<VmlStroke>,
    /// VML §14.1.2.14: path child element.
    pub vml_path: Option<VmlPath>,
    /// VML §14.1.2.6: formula definitions.
    pub formulas: Vec<VmlFormula>,
}

/// VML §14.1.2.19: a shape instance.
#[derive(Clone, Debug)]
pub struct VmlShape {
    /// Shape identifier.
    pub id: Option<VmlShapeId>,
    /// Reference to a shape type (e.g., "#_x0000_t202").
    pub shape_type_ref: Option<VmlShapeId>,
    /// Parsed CSS2 style properties.
    pub style: VmlStyle,
    /// Fill color.
    pub fill_color: Option<VmlColor>,
    /// Whether the shape has a stroke.
    pub stroked: Option<bool>,
    /// VML §14.1.2.21: stroke child element.
    pub stroke: Option<VmlStroke>,
    /// VML §14.1.2.14: path child element.
    pub vml_path: Option<VmlPath>,
    /// VML §14.1.2.22: text box child element.
    pub text_box: Option<VmlTextBox>,
}

/// VML style — parsed CSS2 properties from the `style` attribute (§14.1.2.19).
#[derive(Clone, Debug, Default)]
pub struct VmlStyle {
    /// CSS `position`.
    pub position: Option<CssPosition>,
    /// CSS `left`.
    pub left: Option<VmlLength>,
    /// CSS `top`.
    pub top: Option<VmlLength>,
    /// CSS `width`.
    pub width: Option<VmlLength>,
    /// CSS `height`.
    pub height: Option<VmlLength>,
    /// CSS `margin-left`.
    pub margin_left: Option<VmlLength>,
    /// CSS `margin-top`.
    pub margin_top: Option<VmlLength>,
    /// CSS `margin-right`.
    pub margin_right: Option<VmlLength>,
    /// CSS `margin-bottom`.
    pub margin_bottom: Option<VmlLength>,
    /// CSS `z-index`.
    pub z_index: Option<i64>,
    /// CSS `rotation` (degrees).
    pub rotation: Option<f64>,
    /// VML `flip`.
    pub flip: Option<VmlFlip>,
    /// CSS `visibility`.
    pub visibility: Option<CssVisibility>,
    /// Office `mso-position-horizontal`.
    pub mso_position_horizontal: Option<MsoPositionH>,
    /// Office `mso-position-horizontal-relative`.
    pub mso_position_horizontal_relative: Option<MsoPositionHRelative>,
    /// Office `mso-position-vertical`.
    pub mso_position_vertical: Option<MsoPositionV>,
    /// Office `mso-position-vertical-relative`.
    pub mso_position_vertical_relative: Option<MsoPositionVRelative>,
    /// Office `mso-wrap-distance-left`.
    pub mso_wrap_distance_left: Option<VmlLength>,
    /// Office `mso-wrap-distance-right`.
    pub mso_wrap_distance_right: Option<VmlLength>,
    /// Office `mso-wrap-distance-top`.
    pub mso_wrap_distance_top: Option<VmlLength>,
    /// Office `mso-wrap-distance-bottom`.
    pub mso_wrap_distance_bottom: Option<VmlLength>,
    /// Office `mso-wrap-style`.
    pub mso_wrap_style: Option<MsoWrapStyle>,
}

/// VML color value (§14.1.2.1 ST_ColorType).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum VmlColor {
    /// Hex-specified sRGB color (r, g, b).
    Rgb(u8, u8, u8),
    /// Predefined named color.
    Named(VmlNamedColor),
}

/// VML/CSS named colors (§14.1.2.1, CSS2.1 §4.3.6, and SVG/CSS3 extended colors).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum VmlNamedColor {
    // CSS2.1 §4.3.6 — the 17 standard colors.
    Black,
    Silver,
    Gray,
    White,
    Maroon,
    Red,
    Purple,
    Fuchsia,
    Green,
    Lime,
    Olive,
    Yellow,
    Navy,
    Blue,
    Teal,
    Aqua,
    Orange,
    // SVG/CSS3 extended named colors used in Office documents.
    AliceBlue,
    AntiqueWhite,
    Beige,
    Bisque,
    BlanchedAlmond,
    BlueViolet,
    Brown,
    BurlyWood,
    CadetBlue,
    Chartreuse,
    Chocolate,
    Coral,
    CornflowerBlue,
    Cornsilk,
    Crimson,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGoldenrod,
    DarkGray,
    DarkGreen,
    DarkKhaki,
    DarkMagenta,
    DarkOliveGreen,
    DarkOrange,
    DarkOrchid,
    DarkRed,
    DarkSalmon,
    DarkSeaGreen,
    DarkSlateBlue,
    DarkSlateGray,
    DarkTurquoise,
    DarkViolet,
    DeepPink,
    DeepSkyBlue,
    DimGray,
    DodgerBlue,
    Firebrick,
    FloralWhite,
    ForestGreen,
    Gainsboro,
    GhostWhite,
    Gold,
    Goldenrod,
    GreenYellow,
    Honeydew,
    HotPink,
    IndianRed,
    Indigo,
    Ivory,
    Khaki,
    Lavender,
    LavenderBlush,
    LawnGreen,
    LemonChiffon,
    LightBlue,
    LightCoral,
    LightCyan,
    LightGoldenrodYellow,
    LightGray,
    LightGreen,
    LightPink,
    LightSalmon,
    LightSeaGreen,
    LightSkyBlue,
    LightSlateGray,
    LightSteelBlue,
    LightYellow,
    LimeGreen,
    Linen,
    Magenta,
    MediumAquamarine,
    MediumBlue,
    MediumOrchid,
    MediumPurple,
    MediumSeaGreen,
    MediumSlateBlue,
    MediumSpringGreen,
    MediumTurquoise,
    MediumVioletRed,
    MidnightBlue,
    MintCream,
    MistyRose,
    Moccasin,
    NavajoWhite,
    OldLace,
    OliveDrab,
    OrangeRed,
    Orchid,
    PaleGoldenrod,
    PaleGreen,
    PaleTurquoise,
    PaleVioletRed,
    PapayaWhip,
    PeachPuff,
    Peru,
    Pink,
    Plum,
    PowderBlue,
    RosyBrown,
    RoyalBlue,
    SaddleBrown,
    Salmon,
    SandyBrown,
    SeaGreen,
    Seashell,
    Sienna,
    SkyBlue,
    SlateBlue,
    SlateGray,
    Snow,
    SpringGreen,
    SteelBlue,
    Tan,
    Thistle,
    Tomato,
    Turquoise,
    Violet,
    Wheat,
    WhiteSmoke,
    YellowGreen,
    // VML §14.1.2.1 system colors.
    ButtonFace,
    ButtonHighlight,
    ButtonShadow,
    ButtonText,
    CaptionText,
    GrayText,
    Highlight,
    HighlightText,
    InactiveBorder,
    InactiveCaption,
    InactiveCaptionText,
    InfoBackground,
    InfoText,
    Menu,
    MenuText,
    Scrollbar,
    ThreeDDarkShadow,
    ThreeDFace,
    ThreeDHighlight,
    ThreeDLightShadow,
    ThreeDShadow,
    Window,
    WindowFrame,
    WindowText,
}

/// CSS2 `position` property values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CssPosition {
    Static,
    Relative,
    Absolute,
}

/// CSS2 `visibility` property values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CssVisibility {
    Visible,
    Hidden,
    Inherit,
}

/// VML `flip` attribute values (§14.1.2.19).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlFlip {
    /// Flip along the x-axis.
    X,
    /// Flip along the y-axis.
    Y,
    /// Flip along both axes.
    XY,
}

/// Office `mso-position-horizontal` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionH {
    Absolute,
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

/// Office `mso-position-horizontal-relative` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionHRelative {
    Margin,
    Page,
    Text,
    Char,
    LeftMarginArea,
    RightMarginArea,
    InnerMarginArea,
    OuterMarginArea,
}

/// Office `mso-position-vertical` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionV {
    Absolute,
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
}

/// Office `mso-position-vertical-relative` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionVRelative {
    Margin,
    Page,
    Text,
    Line,
    TopMarginArea,
    BottomMarginArea,
    InnerMarginArea,
    OuterMarginArea,
}

/// Office `mso-wrap-style` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoWrapStyle {
    Square,
    None,
    Tight,
    Through,
}

/// A CSS length value with unit.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct VmlLength {
    pub value: f64,
    pub unit: VmlLengthUnit,
}

/// CSS length unit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlLengthUnit {
    /// Points (pt).
    Pt,
    /// Inches (in).
    In,
    /// Centimeters (cm).
    Cm,
    /// Millimeters (mm).
    Mm,
    /// Pixels (px).
    Px,
    /// Em units (em).
    Em,
    /// Percentage (%).
    Percent,
    /// No unit (bare number, treated as EMU in VML context).
    None,
}

/// VML §14.1.2.21: stroke styling.
#[derive(Clone, Debug)]
pub struct VmlStroke {
    /// Dash pattern.
    pub dash_style: Option<VmlDashStyle>,
    /// Line join style.
    pub join_style: Option<VmlJoinStyle>,
}

/// VML §14.1.2.21: stroke dash patterns.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlDashStyle {
    Solid,
    ShortDash,
    ShortDot,
    ShortDashDot,
    ShortDashDotDot,
    Dot,
    Dash,
    LongDash,
    DashDot,
    LongDashDot,
    LongDashDotDot,
}

/// VML §14.1.2.21: line join styles.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlJoinStyle {
    Round,
    Bevel,
    Miter,
}

/// VML §14.1.2.14: path properties.
#[derive(Clone, Debug)]
pub struct VmlPath {
    /// o:gradientshapeok — whether the path supports gradient fill.
    pub gradient_shape_ok: Option<bool>,
    /// o:connecttype — connection point type.
    pub connect_type: Option<VmlConnectType>,
    /// o:extrusionok — whether the path supports 3D extrusion.
    pub extrusion_ok: Option<bool>,
}

/// VML o:connecttype — how connection points are defined on the shape.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlConnectType {
    /// No connection points.
    None,
    /// Four connection points at the rectangle midpoints.
    Rect,
    /// Connection points derived from path segments.
    Segments,
    /// Custom connection points.
    Custom,
}

/// VML §14.1.2.22: text box inset margins (comma-separated CSS lengths).
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct VmlTextBoxInset {
    pub left: Option<VmlLength>,
    pub top: Option<VmlLength>,
    pub right: Option<VmlLength>,
    pub bottom: Option<VmlLength>,
}

/// VML §14.1.2.22: text box within a shape.
#[derive(Clone, Debug)]
pub struct VmlTextBox {
    /// Parsed CSS2 style properties.
    pub style: VmlStyle,
    /// VML §14.1.2.22: inset margins (top, left, bottom, right).
    pub inset: Option<VmlTextBoxInset>,
    /// §17.17.1: block-level content from w:txbxContent.
    pub content: Vec<Block>,
}

// ── Alternate Content ───────────────────────────────────────────────────────

/// MCE §M.2.1: alternate content for markup compatibility.
#[derive(Clone, Debug)]
pub struct AlternateContent {
    /// §M.2.2: ordered list of preferred choices.
    pub choices: Vec<McChoice>,
    /// §M.2.3: fallback content when no choice is supported.
    pub fallback: Option<Vec<Inline>>,
}

/// MCE §M.2.2: a single choice in alternate content.
#[derive(Clone, Debug)]
pub struct McChoice {
    /// Required namespace/feature identifier.
    pub requires: McRequires,
    /// Inline content for this choice.
    pub content: Vec<Inline>,
}

/// MCE §M.2.2: namespace prefixes used in `mc:Choice Requires`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum McRequires {
    /// Word Processing Shape (wps).
    Wps,
    /// Word Processing Group (wpg).
    Wpg,
    /// Word Processing Canvas (wpc).
    Wpc,
    /// Word Processing Ink (wpi).
    Wpi,
    /// Math (m).
    Math,
    /// DrawingML 2010 (a14).
    A14,
    /// Word 2010 extensions (w14).
    W14,
    /// Word 2012 extensions (w15).
    W15,
    /// Word 2016 extensions (w16).
    W16,
}

// ── Table ────────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Table {
    pub properties: TableProperties,
    pub grid: Vec<GridColumn>,
    pub rows: Vec<TableRow>,
}

#[derive(Clone, Debug, Default)]
pub struct TableProperties {
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

/// §17.18.106 ST_VAnchor — vertical/horizontal anchor for table positioning.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableAnchor {
    Text,
    Margin,
    Page,
}

/// §17.18.108 ST_XAlign — horizontal alignment for floating table.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableXAlign {
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

/// §17.18.109 ST_YAlign — vertical alignment for floating table.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableYAlign {
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
    Inline,
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
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableRowHeight {
    pub value: Dimension<Twips>,
    pub rule: HeightRule,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HeightRule {
    Auto,
    Exact,
    AtLeast,
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
