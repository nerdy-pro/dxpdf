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

// ── Settings ─────────────────────────────────────────────────────────────────

#[derive(Clone, Debug)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// Document-level default paragraph properties.
    pub default_paragraph_properties: ParagraphProperties,
    /// Document-level default run properties.
    pub default_run_properties: RunProperties,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            default_tab_stop: Dimension::new(720),
            even_and_odd_headers: false,
            default_paragraph_properties: ParagraphProperties::default(),
            default_run_properties: RunProperties::default(),
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
    pub properties: ParagraphProperties,
    pub content: Vec<Inline>,
}

#[derive(Clone, Debug)]
pub struct ParagraphProperties {
    pub alignment: Alignment,
    pub indentation: Indentation,
    pub spacing: ParagraphSpacing,
    pub numbering: Option<NumberingProperties>,
    pub tabs: Vec<TabStop>,
    pub borders: Option<ParagraphBorders>,
    pub shading: Option<Shading>,
    pub keep_next: bool,
    pub keep_lines: bool,
    pub widow_control: bool,
    pub page_break_before: bool,
    pub suppress_auto_hyphens: bool,
    pub bidi: bool,
    pub outline_level: Option<OutlineLevel>,
}

impl Default for ParagraphProperties {
    fn default() -> Self {
        Self {
            alignment: Alignment::Start,
            indentation: Indentation::default(),
            spacing: ParagraphSpacing::default(),
            numbering: None,
            tabs: Vec::new(),
            borders: None,
            shading: None,
            keep_next: false,
            keep_lines: false,
            widow_control: true, // OOXML default is true
            page_break_before: false,
            suppress_auto_hyphens: false,
            bidi: false,
            outline_level: None,
        }
    }
}

/// Heading outline level (0–8, where 0 = Heading 1 in OOXML).
/// We normalize: Level(1) = Heading 1 .. Level(9) = Heading 9.
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
    pub start: Dimension<Twips>,
    pub end: Dimension<Twips>,
    pub first_line: FirstLineIndent,
    pub mirror: bool,
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ParagraphSpacing {
    pub before: Dimension<Twips>,
    pub after: Dimension<Twips>,
    pub line: LineSpacing,
    pub before_auto_spacing: bool,
    pub after_auto_spacing: bool,
}

impl Default for ParagraphSpacing {
    fn default() -> Self {
        Self {
            before: Dimension::ZERO,
            after: Dimension::ZERO,
            line: LineSpacing::default(),
            before_auto_spacing: false,
            after_auto_spacing: false,
        }
    }
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

impl Default for LineSpacing {
    fn default() -> Self {
        // OOXML default: single spacing = 240 twips in auto mode
        Self::Auto(Dimension::new(240))
    }
}

// ── Numbering ────────────────────────────────────────────────────────────────

/// Fully resolved numbering properties for a paragraph.
#[derive(Clone, Debug)]
pub struct NumberingProperties {
    pub level: u8,
    pub format: NumberFormat,
    pub level_text: String,
    pub indent: Indentation,
    pub run_properties: Option<RunProperties>,
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
    pub properties: RunProperties,
    pub text: String,
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

#[derive(Clone, Debug, PartialEq)]
pub struct RunProperties {
    pub fonts: FontSet,
    pub font_size: Dimension<HalfPoints>,
    pub bold: bool,
    pub italic: bool,
    pub underline: UnderlineStyle,
    pub strike: StrikeStyle,
    pub color: Color,
    pub highlight: Option<HighlightColor>,
    pub shading: Option<Shading>,
    pub vertical_align: VerticalAlign,
    pub spacing: Dimension<Twips>,
    pub kerning: Option<Dimension<HalfPoints>>,
    pub all_caps: bool,
    pub small_caps: bool,
    pub vanish: bool,
    pub rtl: bool,
    pub emboss: bool,
    pub imprint: bool,
    pub outline: bool,
    pub shadow: bool,
}

impl Default for RunProperties {
    fn default() -> Self {
        Self {
            fonts: FontSet::default(),
            font_size: Dimension::new(20), // 10pt
            bold: false,
            italic: false,
            underline: UnderlineStyle::None,
            strike: StrikeStyle::None,
            color: Color::Auto,
            highlight: None,
            shading: None,
            vertical_align: VerticalAlign::Baseline,
            spacing: Dimension::ZERO,
            kerning: None,
            all_caps: false,
            small_caps: false,
            vanish: false,
            rtl: false,
            emboss: false,
            imprint: false,
            outline: false,
            shadow: false,
        }
    }
}

/// Resolved font family names for each script category.
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
}

#[derive(Clone, Debug, Default)]
pub struct TableRowProperties {
    pub height: Option<TableRowHeight>,
    pub is_header: bool,
    pub cant_split: bool,
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

#[derive(Clone, Debug)]
pub struct TableCellProperties {
    pub width: TableMeasure,
    pub borders: Option<TableCellBorders>,
    pub shading: Option<Shading>,
    pub margins: Option<EdgeInsets<Twips>>,
    pub vertical_align: CellVerticalAlign,
    pub merge: CellMerge,
    pub text_direction: TextDirection,
    pub no_wrap: bool,
}

impl Default for TableCellProperties {
    fn default() -> Self {
        Self {
            width: TableMeasure::Auto,
            borders: None,
            shading: None,
            margins: None,
            vertical_align: CellVerticalAlign::Top,
            merge: CellMerge::None,
            text_direction: TextDirection::LeftToRightTopToBottom,
            no_wrap: false,
        }
    }
}

/// Cell merge state — makes vertical/horizontal merge explicit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum CellMerge {
    #[default]
    None,
    /// This cell starts a vertical merge group.
    VerticalStart,
    /// This cell continues a vertical merge from above.
    VerticalContinue,
    /// Horizontal span count (grid_span > 1).
    HorizontalSpan(u32),
    /// Both vertical start and horizontal span.
    VerticalStartWithSpan(u32),
    /// Both vertical continue and horizontal span.
    VerticalContinueWithSpan(u32),
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
    pub first_row: bool,
    pub last_row: bool,
    pub first_column: bool,
    pub last_column: bool,
    pub no_h_band: bool,
    pub no_v_band: bool,
}
