//! Complete DOCX document model — all types are fully resolved ADTs.
//! No unparsed strings, no style indirection, no invalid states.

use std::collections::HashMap;

use crate::dimension::{Dimension, EighthPoints, Emu, HalfPoints, ThousandthPercent, Twips};
use crate::geometry::{EdgeInsets, Size};

// ── Identifiers ──────────────────────────────────────────────────────────────

/// A relationship ID (e.g., "rId1") — opaque, interned from the .rels files.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct RelId(pub(crate) String);

impl RelId {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Footnote or endnote numeric ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NoteId(pub(crate) i64);

impl NoteId {
    pub fn value(self) -> i64 {
        self.0
    }
}

/// A bookmark ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct BookmarkId(pub(crate) i64);

/// Revision Save ID — identifies which editing session produced a change.
/// Stored as a 32-bit value parsed from an 8-digit hex string (e.g., "00A2B3C4").
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Rsid(pub(crate) u32);

impl Rsid {
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
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct RevisionIds {
    /// Session that added this element.
    pub r: Option<Rsid>,
    /// Session that last modified this element's properties.
    pub r_pr: Option<Rsid>,
    /// Session that deleted this element (for tracked deletions).
    pub del: Option<Rsid>,
}

/// Revision tracking IDs specific to paragraphs.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ParagraphRevisionIds {
    /// Session that added this paragraph.
    pub r: Option<Rsid>,
    /// Session that added the default run content.
    pub r_default: Option<Rsid>,
    /// Session that last modified paragraph properties.
    pub p: Option<Rsid>,
    /// Session that last modified run properties on the paragraph mark.
    pub r_pr: Option<Rsid>,
    /// Session that deleted this paragraph.
    pub del: Option<Rsid>,
}

/// Revision tracking IDs specific to table rows.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableRowRevisionIds {
    /// Session that added this row.
    pub r: Option<Rsid>,
    /// Session that last modified row properties.
    pub r_pr: Option<Rsid>,
    /// Session that deleted this row.
    pub del: Option<Rsid>,
    /// Session that last modified this row's table-level formatting.
    pub tr: Option<Rsid>,
}

/// Revision tracking IDs specific to section properties.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SectionRevisionIds {
    /// Session that added this section.
    pub r: Option<Rsid>,
    /// Session that last modified section run properties.
    pub r_pr: Option<Rsid>,
    /// Session that last modified section properties.
    pub sect: Option<Rsid>,
}

// ── Color ────────────────────────────────────────────────────────────────────

/// A fully resolved color — no theme references survive parsing.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Color {
    /// sRGB color (0xRRGGBB).
    Rgb(u32),
    /// The special "auto" color — meaning context-dependent (usually black for text).
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

#[derive(Clone, Debug, Default)]
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
    pub styles: HashMap<String, Style>,
}

/// A single style definition.
#[derive(Clone, Debug)]
pub struct Style {
    pub name: Option<String>,
    pub style_type: StyleType,
    /// Parent style ID. Properties not specified here should be inherited from this style.
    pub based_on: Option<String>,
    pub is_default: bool,
    pub paragraph_properties: Option<ParagraphProperties>,
    pub run_properties: Option<RunProperties>,
    pub table_properties: Option<TableProperties>,
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
    pub format: NumberFormat,
    pub level_text: String,
    pub start: Option<u32>,
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

#[derive(Clone, Debug)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// The rsid of the original editing session that created this document.
    pub rsid_root: Option<Rsid>,
    /// All revision save IDs recorded in this document's history.
    pub rsids: Vec<Rsid>,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            default_tab_stop: Dimension::new(720),
            even_and_odd_headers: false,
            rsid_root: None,
            rsids: Vec::new(),
        }
    }
}

// ── Section ──────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct SectionProperties {
    pub page_size: PageSize,
    pub page_margins: PageMargins,
    pub columns: Columns,
    pub header_refs: SectionHeaderFooterRefs,
    pub footer_refs: SectionHeaderFooterRefs,
    /// If true, the first page uses distinct header/footer.
    pub title_page: bool,
    pub section_type: SectionType,
    pub rsids: SectionRevisionIds,
}

impl Default for SectionProperties {
    fn default() -> Self {
        Self {
            page_size: PageSize::default(),
            page_margins: PageMargins::default(),
            columns: Columns::default(),
            header_refs: SectionHeaderFooterRefs::default(),
            footer_refs: SectionHeaderFooterRefs::default(),
            title_page: false,
            section_type: SectionType::NextPage,
            rsids: SectionRevisionIds::default(),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SectionType {
    NextPage,
    Continuous,
    EvenPage,
    OddPage,
    NextColumn,
}

#[derive(Clone, Debug)]
pub struct PageSize {
    pub width: Dimension<Twips>,
    pub height: Dimension<Twips>,
    pub orientation: PageOrientation,
}

impl Default for PageSize {
    fn default() -> Self {
        // US Letter
        Self {
            width: Dimension::new(12240),  // 8.5 inches
            height: Dimension::new(15840), // 11 inches
            orientation: PageOrientation::Portrait,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum PageOrientation {
    #[default]
    Portrait,
    Landscape,
}

#[derive(Clone, Debug)]
pub struct PageMargins {
    pub top: Dimension<Twips>,
    pub right: Dimension<Twips>,
    pub bottom: Dimension<Twips>,
    pub left: Dimension<Twips>,
    pub header: Dimension<Twips>,
    pub footer: Dimension<Twips>,
    pub gutter: Dimension<Twips>,
}

impl Default for PageMargins {
    fn default() -> Self {
        Self {
            top: Dimension::new(1440), // 1 inch
            right: Dimension::new(1440),
            bottom: Dimension::new(1440),
            left: Dimension::new(1440),
            header: Dimension::new(720), // 0.5 inch
            footer: Dimension::new(720),
            gutter: Dimension::ZERO,
        }
    }
}

#[derive(Clone, Debug)]
pub struct Columns {
    pub count: u32,
    pub space: Dimension<Twips>,
    pub equal_width: bool,
}

impl Default for Columns {
    fn default() -> Self {
        Self {
            count: 1,
            space: Dimension::new(720),
            equal_width: true,
        }
    }
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
    pub style_id: Option<String>,
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
    pub bidi: Option<bool>,
    pub outline_level: Option<OutlineLevel>,
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

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum Alignment {
    #[default]
    Start,
    Center,
    End,
    Both,
    Distribute,
    Thai,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Indentation {
    pub start: Option<Dimension<Twips>>,
    pub end: Option<Dimension<Twips>>,
    pub first_line: Option<FirstLineIndent>,
    pub mirror: Option<bool>,
}

/// First-line indent: either hanging (negative) or first-line (positive).
/// These are mutually exclusive in OOXML.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum FirstLineIndent {
    #[default]
    None,
    FirstLine(Dimension<Twips>),
    Hanging(Dimension<Twips>),
}

/// Paragraph spacing — only fields explicitly present in the XML are `Some`.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
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
#[derive(Clone, Debug)]
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

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TabAlignment {
    #[default]
    Left,
    Center,
    Right,
    Decimal,
    Bar,
    Clear,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TabLeader {
    #[default]
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

// ── Borders ──────────────────────────────────────────────────────────────────

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParagraphBorders {
    pub top: Option<Border>,
    pub bottom: Option<Border>,
    pub left: Option<Border>,
    pub right: Option<Border>,
    pub between: Option<Border>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Border {
    pub style: BorderStyle,
    pub width: Dimension<EighthPoints>,
    pub space: Dimension<Twips>,
    pub color: Color,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum BorderStyle {
    #[default]
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

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Shading {
    pub fill: Color,
    pub pattern: ShadingPattern,
    pub color: Color,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ShadingPattern {
    #[default]
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
    TextRun(TextRun),
    Tab,
    LineBreak(BreakKind),
    ColumnBreak,
    PageBreak,
    Image(Image),
    FootnoteRef(NoteId),
    EndnoteRef(NoteId),
    Hyperlink(Hyperlink),
    Field(Field),
    BookmarkStart { id: BookmarkId, name: String },
    BookmarkEnd(BookmarkId),
    Symbol(Symbol),
}

#[derive(Clone, Debug)]
pub struct TextRun {
    /// Character style ID reference (e.g., "Hyperlink"). Resolve via `Document.styles`.
    pub style_id: Option<String>,
    pub properties: RunProperties,
    pub text: String,
    pub rsids: RevisionIds,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum BreakKind {
    #[default]
    TextWrapping,
    /// Clears left, right, or both float areas.
    Clear(BreakClear),
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum BreakClear {
    #[default]
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
#[derive(Clone, Debug, Default, PartialEq)]
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
    pub rtl: Option<bool>,
    pub emboss: Option<bool>,
    pub imprint: Option<bool>,
    pub outline: Option<bool>,
    pub shadow: Option<bool>,
}

/// Font family names for each script category.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct FontSet {
    pub ascii: Option<String>,
    pub high_ansi: Option<String>,
    pub east_asian: Option<String>,
    pub complex_script: Option<String>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum UnderlineStyle {
    #[default]
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

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum StrikeStyle {
    #[default]
    None,
    Single,
    Double,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum VerticalAlign {
    #[default]
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

// ── Image ────────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Image {
    pub rel_id: RelId,
    pub extent: Size<Emu>,
    pub placement: ImagePlacement,
    pub description: Option<String>,
}

/// How the image is placed in the document flow.
#[derive(Clone, Debug)]
pub enum ImagePlacement {
    /// Inline with text — no wrapping.
    Inline,
    /// Floating/anchored with text wrapping.
    Anchor(AnchorProperties),
}

#[derive(Clone, Debug)]
pub struct AnchorProperties {
    pub horizontal_position: AnchorPosition,
    pub vertical_position: AnchorPosition,
    pub wrap: TextWrap,
    pub behind_text: bool,
    pub lock_anchor: bool,
    pub allow_overlap: bool,
    pub relative_height: u32,
}

#[derive(Clone, Debug)]
pub enum AnchorPosition {
    /// Offset from the anchor base.
    Offset {
        relative_from: AnchorRelativeFrom,
        offset: Dimension<Emu>,
    },
    /// Aligned to an edge of the relative area.
    Align {
        relative_from: AnchorRelativeFrom,
        alignment: AnchorAlignment,
    },
}

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

#[derive(Clone, Debug)]
pub enum TextWrap {
    /// No wrapping — text does not flow around the image.
    None,
    /// Wrap on both sides.
    Square { distance: EdgeInsets<Emu> },
    /// Wrap tightly along the shape outline.
    Tight { distance: EdgeInsets<Emu> },
    /// Text appears above and below only.
    TopAndBottom {
        distance_top: Dimension<Emu>,
        distance_bottom: Dimension<Emu>,
    },
    /// Behind text or in front of text — no wrapping displacement.
    Through { distance: EdgeInsets<Emu> },
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

/// A resolved field (e.g., PAGE, NUMPAGES, TOC, etc.).
#[derive(Clone, Debug)]
pub struct Field {
    pub kind: FieldKind,
    pub content: Vec<Inline>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FieldKind {
    Page,
    NumPages,
    Date,
    Time,
    FileName,
    Author,
    Title,
    /// Table of contents field.
    Toc,
    /// An unrecognized field — stores the raw field code for forward compatibility.
    Other(String),
}

// ── Table ────────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct Table {
    pub properties: TableProperties,
    pub grid: Vec<GridColumn>,
    pub rows: Vec<TableRow>,
}

#[derive(Clone, Debug)]
pub struct TableProperties {
    pub alignment: Alignment,
    pub width: TableMeasure,
    pub layout: TableLayout,
    pub indent: Option<TableMeasure>,
    pub borders: Option<TableBorders>,
    pub cell_margins: Option<EdgeInsets<Twips>>,
    pub cell_spacing: Option<TableMeasure>,
    pub look: TableLook,
}

impl Default for TableProperties {
    fn default() -> Self {
        Self {
            alignment: Alignment::Start,
            width: TableMeasure::Auto,
            layout: TableLayout::Auto,
            indent: None,
            borders: None,
            cell_margins: None,
            cell_spacing: None,
            look: TableLook::default(),
        }
    }
}

/// A dimension for table/cell widths — may be auto, fixed, or percentage.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum TableMeasure {
    #[default]
    Auto,
    Twips(Dimension<Twips>),
    /// Percentage in 50ths of a percent (OOXML `pct` type).
    Pct(Dimension<ThousandthPercent>),
    /// Nil — explicitly zero.
    Nil,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TableLayout {
    #[default]
    Auto,
    Fixed,
}

#[derive(Clone, Debug)]
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
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableRowHeight {
    pub value: Dimension<Twips>,
    pub rule: HeightRule,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum HeightRule {
    #[default]
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
}

/// Vertical merge state from `w:vMerge` attribute.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VerticalMerge {
    /// `w:vMerge val="restart"` — this cell starts a new vertical merge group.
    Restart,
    /// `w:vMerge` (no val or val="continue") — this cell continues from above.
    Continue,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum CellVerticalAlign {
    #[default]
    Top,
    Center,
    Bottom,
    Both,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TextDirection {
    #[default]
    LeftToRightTopToBottom,
    TopToBottomRightToLeft,
    BottomToTopLeftToRight,
    LeftToRightTopToBottomRotated,
    TopToBottomRightToLeftRotated,
    TopToBottomLeftToRightRotated,
}

#[derive(Clone, Debug)]
pub struct TableBorders {
    pub top: Option<Border>,
    pub bottom: Option<Border>,
    pub left: Option<Border>,
    pub right: Option<Border>,
    pub inside_h: Option<Border>,
    pub inside_v: Option<Border>,
}

#[derive(Clone, Debug)]
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
