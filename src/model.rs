use std::rc::Rc;

use crate::dimension::{EighthPoints, HalfPoints, Pt, Twips};
use crate::geometry::{PtOffset, PtSize, TwipsEdgeInsets, TwipsSize};
/// Default font family when nothing is specified.
pub const DEFAULT_FONT_FAMILY: &str = "Helvetica";

use std::collections::HashMap;

/// Image store mapping relationship IDs to raw image bytes.
/// Produced during parsing, consumed during layout/rendering.
pub type ImageStore = HashMap<String, Vec<u8>>;

/// Resolved paragraph style properties.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ResolvedParagraphStyle {
    pub alignment: Option<Alignment>,
    pub spacing: Option<Spacing>,
    pub indentation: Option<Indentation>,
    /// Run properties from the style's rPr (applied to all runs in the paragraph).
    pub run_props: ResolvedRunStyle,
}

impl ResolvedParagraphStyle {
    /// Fill in `None` fields from a base (parent) style.
    pub fn merge_from(&mut self, base: &ResolvedParagraphStyle) {
        if self.alignment.is_none() {
            self.alignment = base.alignment;
        }
        if self.spacing.is_none() {
            self.spacing = base.spacing;
        }
        if self.indentation.is_none() {
            self.indentation = base.indentation;
        }
        self.run_props.merge_from(&base.run_props);
    }
}

/// Resolved run (character) style properties.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ResolvedRunStyle {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub font_size: Option<HalfPoints>,
    pub font_family: Option<Rc<str>>,
    pub color: Option<Color>,
}

impl ResolvedRunStyle {
    /// Fill in `None` fields from a base (parent) style.
    pub fn merge_from(&mut self, base: &ResolvedRunStyle) {
        if self.bold.is_none() {
            self.bold = base.bold;
        }
        if self.italic.is_none() {
            self.italic = base.italic;
        }
        if self.underline.is_none() {
            self.underline = base.underline;
        }
        if self.font_size.is_none() {
            self.font_size = base.font_size;
        }
        if self.font_family.is_none() {
            self.font_family = base.font_family.clone();
        }
        if self.color.is_none() {
            self.color = base.color;
        }
    }
}

/// Map of style IDs to resolved properties.
pub type StyleMap = HashMap<String, ResolvedParagraphStyle>;

/// Number format type for list items.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NumberFormat {
    Bullet(String), // The bullet character
    Decimal,        // 1, 2, 3
    LowerLetter,    // a, b, c
    UpperLetter,    // A, B, C
    LowerRoman,     // i, ii, iii
    UpperRoman,     // I, II, III
}

/// A single numbering level definition.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingLevel {
    pub format: NumberFormat,
    /// Format pattern (e.g., "%1." for "1.")
    pub level_text: String,
    pub start: u32,
    /// Indentation: left margin.
    pub indent_left: Twips,
    /// Indentation: hanging indent.
    pub indent_hanging: Twips,
}

/// A numbering definition (abstractNum + num mapping).
#[derive(Debug, Clone, PartialEq)]
pub struct NumberingDef {
    pub levels: Vec<NumberingLevel>,
}

/// Map of numId -> NumberingDef.
pub type NumberingMap = HashMap<u32, NumberingDef>;

/// List reference on a paragraph.
#[derive(Debug, Clone, PartialEq)]
pub struct ListRef {
    pub num_id: u32,
    pub level: u32,
}

/// A section: a group of blocks sharing the same page geometry and header/footer.
#[derive(Debug, Clone, PartialEq)]
pub struct Section {
    pub properties: SectionProperties,
    pub blocks: Vec<Block>,
}

/// Document-wide defaults from `word/styles.xml`, `word/settings.xml`, and theme.
#[derive(Debug, Clone, PartialEq)]
pub struct DocumentDefaults {
    /// Default tab stop interval.
    pub tab_stop: Twips,
    /// Default font size in half-points.
    pub font_size: HalfPoints,
    /// Default font family.
    pub font_family: Rc<str>,
    /// Default paragraph spacing.
    pub spacing: Spacing,
    /// Default table cell margins.
    pub cell_margins: CellMargins,
    /// Default paragraph spacing inside table cells.
    pub table_cell_spacing: Spacing,
    /// Default table borders (from table style).
    pub table_borders: TableBorders,
    /// Named paragraph/run styles resolved from `word/styles.xml`.
    pub styles: StyleMap,
    /// Numbering definitions from `word/numbering.xml`.
    pub numbering: NumberingMap,
}

impl Default for DocumentDefaults {
    fn default() -> Self {
        Self {
            tab_stop: Document::DEFAULT_TAB_STOP,
            font_size: Document::DEFAULT_FONT_SIZE,
            font_family: Rc::from(DEFAULT_FONT_FAMILY),
            spacing: Spacing::default(),
            cell_margins: CellMargins::default(),
            table_cell_spacing: Spacing {
                after: Some(Twips::new(0)),
                ..Default::default()
            },
            table_borders: TableBorders::default(),
            styles: StyleMap::new(),
            numbering: NumberingMap::new(),
        }
    }
}

/// Root of the document tree.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Document {
    pub sections: Vec<Section>,
    /// Document-wide defaults (fonts, spacing, styles, numbering).
    pub defaults: DocumentDefaults,
    /// Raw image bytes keyed by relationship ID.
    pub images: ImageStore,
}

impl Document {
    /// Default font size: 24 half-points (12pt).
    /// Matches Microsoft Word's default; not mandated by the OOXML spec.
    pub const DEFAULT_FONT_SIZE: HalfPoints = HalfPoints::new(24);

    /// Default tab stop interval: 720 twips (0.5 inch).
    pub const DEFAULT_TAB_STOP: Twips = Twips::new(720);

    /// Iterate over all section properties in document order.
    pub fn section_properties(&self) -> impl Iterator<Item = &SectionProperties> {
        self.sections.iter().map(|s| &s.properties)
    }

    /// Iterate over all blocks across all sections.
    pub fn all_blocks(&self) -> impl Iterator<Item = &Block> {
        self.sections.iter().flat_map(|s| &s.blocks)
    }

    /// Collect all unique font families referenced in this document.
    pub fn font_families(&self) -> Vec<Rc<str>> {
        use std::collections::HashSet;
        let mut families = HashSet::new();
        families.insert(self.defaults.font_family.clone());

        // From styles
        for style in self.defaults.styles.values() {
            if let Some(ref f) = style.run_props.font_family {
                families.insert(f.clone());
            }
        }

        fn collect_from_blocks(blocks: &[Block], families: &mut HashSet<Rc<str>>) {
            for block in blocks {
                match block {
                    Block::Paragraph(p) => {
                        for inline in &p.runs {
                            if let Inline::TextRun(tr) = inline {
                                if let Some(ref f) = tr.properties.font_family {
                                    families.insert(f.clone());
                                }
                            }
                        }
                    }
                    Block::Table(t) => {
                        for row in &t.rows {
                            for cell in &row.cells {
                                collect_from_blocks(&cell.blocks, families);
                            }
                        }
                    }
                }
            }
        }

        for section in &self.sections {
            collect_from_blocks(&section.blocks, &mut families);
            if let Some(ref hf) = section.properties.header {
                collect_from_blocks(&hf.blocks, &mut families);
            }
            if let Some(ref hf) = section.properties.footer {
                collect_from_blocks(&hf.blocks, &mut families);
            }
        }

        families.into_iter().collect()
    }
}

/// A block-level element.
#[derive(Debug, Clone, PartialEq)]
pub enum Block {
    Paragraph(Box<Paragraph>),
    Table(Box<Table>),
}

/// Page size — a type alias for `TwipsSize`.
pub type PageSize = TwipsSize;

/// Page margins.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageMargins {
    pub top: Twips,
    pub right: Twips,
    pub bottom: Twips,
    pub left: Twips,
    /// Distance from page top to header content.
    pub header: Twips,
    /// Distance from page bottom to footer content.
    pub footer: Twips,
}

/// Header or footer content — same structure as the document body.
#[derive(Debug, Clone, PartialEq)]
pub struct HeaderFooter {
    pub blocks: Vec<Block>,
}

/// Section properties from `w:sectPr`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectionProperties {
    pub page_size: Option<PageSize>,
    pub page_margins: Option<PageMargins>,
    /// Default header content (type="default").
    pub header: Option<HeaderFooter>,
    /// Default footer content (type="default").
    pub footer: Option<HeaderFooter>,
    /// Relationship ID for the default header (used during parsing to resolve content).
    pub header_rel_id: Option<String>,
    /// Relationship ID for the default footer.
    pub footer_rel_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Paragraph {
    pub properties: ParagraphProperties,
    pub runs: Vec<Inline>,
    pub floats: Vec<FloatingImage>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphProperties {
    pub alignment: Option<Alignment>,
    pub spacing: Option<Spacing>,
    pub indentation: Option<Indentation>,
    pub tab_stops: Vec<TabStop>,
    /// Paragraph background shading from `w:shd`.
    pub shading: Option<Color>,
    /// Referenced paragraph style ID (e.g., "Kopfzeile").
    pub style_id: Option<String>,
    /// List reference (numId + level) from `w:numPr`.
    pub list_ref: Option<ListRef>,
    /// Paragraph borders from `w:pBdr`.
    pub paragraph_borders: Option<ParagraphBorders>,
}

/// Paragraph border edges from `w:pBdr`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphBorders {
    pub top: Option<BorderDef>,
    pub bottom: Option<BorderDef>,
    pub left: Option<BorderDef>,
    pub right: Option<BorderDef>,
}

/// A tab stop alignment type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabStopType {
    Left,
    Center,
    Right,
    Decimal,
}

/// A custom tab stop definition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TabStop {
    pub position: Twips,
    pub stop_type: TabStopType,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

/// Line spacing rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineRule {
    /// `line` value is a multiplier: 240 = single (100%), 480 = double (200%).
    #[default]
    Auto,
    /// `line` value is an exact height in twips.
    Exact,
    /// `line` value is a minimum height in twips.
    AtLeast,
}

/// Resolved line spacing value.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacing {
    /// Multiplier of font's natural line height (1.0 = single, 1.15 = 1.15x).
    Multiplier(f32),
    /// Fixed height.
    Fixed(Pt),
    /// Minimum height.
    AtLeast(Pt),
}

/// Paragraph spacing.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Spacing {
    pub before: Option<Twips>,
    pub after: Option<Twips>,
    pub line: Option<Twips>,
    pub line_rule: LineRule,
}

/// Paragraph indentation.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Indentation {
    pub left: Option<Twips>,
    pub right: Option<Twips>,
    pub first_line: Option<Twips>,
}

impl Indentation {
    pub fn left_pt(&self) -> Pt {
        self.left.map(Pt::from).unwrap_or(Pt::ZERO)
    }

    pub fn right_pt(&self) -> Pt {
        self.right.map(Pt::from).unwrap_or(Pt::ZERO)
    }

    pub fn first_line_pt(&self) -> Pt {
        self.first_line.map(Pt::from).unwrap_or(Pt::ZERO)
    }
}

/// A type-safe wrapper for OOXML relationship IDs (e.g., "rId5").
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RelId(String);

impl RelId {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl std::ops::Deref for RelId {
    type Target = str;
    fn deref(&self) -> &str {
        &self.0
    }
}

impl<T: Into<String>> From<T> for RelId {
    fn from(s: T) -> Self {
        Self(s.into())
    }
}

/// A field code that can be evaluated at render time.
#[derive(Debug, Clone, PartialEq)]
pub enum FieldType {
    /// Current page number.
    Page,
    /// Total number of pages.
    NumPages,
}

/// A field with its type and the run properties from the surrounding XML.
#[derive(Debug, Clone, PartialEq)]
pub struct FieldCode {
    pub field_type: FieldType,
    pub properties: RunProperties,
}

/// An inline-level element within a paragraph.
#[derive(Debug, Clone, PartialEq)]
pub enum Inline {
    TextRun(TextRun),
    LineBreak,
    Tab,
    Image(InlineImage),
    /// A field code evaluated at render time (e.g., PAGE, NUMPAGES).
    Field(FieldCode),
}

#[derive(Debug, Clone, PartialEq)]
pub struct InlineImage {
    pub rel_id: RelId,
    pub size: PtSize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WrapSide {
    Left,
    Right,
    BothSides,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FloatingImage {
    pub rel_id: RelId,
    pub size: PtSize,
    pub offset: PtOffset,
    /// Horizontal alignment (e.g., "left", "right", "center") — alternative to offset.
    pub align_h: Option<String>,
    /// Vertical alignment (e.g., "top", "center", "bottom") — alternative to offset.
    pub align_v: Option<String>,
    pub wrap_side: WrapSide,
    /// Percentage-based horizontal position (wp14:pctPosHOffset).
    /// Value is percentage × 1000 (e.g., 5000 = 5% of page width).
    pub pct_pos_h: Option<i32>,
    /// Percentage-based vertical position (wp14:pctPosVOffset).
    /// Value is percentage × 1000 (e.g., 3000 = 3% of page height).
    pub pct_pos_v: Option<i32>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TextRun {
    pub text: String,
    pub properties: RunProperties,
    /// URL for hyperlinked text (resolved from w:hyperlink r:id).
    pub hyperlink_url: Option<String>,
}

/// Vertical alignment of text (superscript/subscript).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VertAlign {
    Superscript,
    Subscript,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RunProperties {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// Font size in half-points (OOXML native for `w:sz`).
    pub font_size: Option<HalfPoints>,
    pub font_family: Option<Rc<str>>,
    pub color: Option<Color>,
    /// Character spacing adjustment (positive = expand, negative = condense).
    pub char_spacing: Option<Twips>,
    /// Background shading color from `w:shd`.
    pub shading: Option<Color>,
    /// Superscript or subscript positioning.
    pub vert_align: Option<VertAlign>,
    /// Character style ID (e.g., "Hyperlink").
    pub style_id: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Color {
    pub const BLACK: Color = Color { r: 0, g: 0, b: 0 };

    pub const fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }
}

impl From<Color> for skia_safe::Color4f {
    fn from(c: Color) -> Self {
        const MAX: f32 = u8::MAX as f32;
        skia_safe::Color4f::new(c.r as f32 / MAX, c.g as f32 / MAX, c.b as f32 / MAX, 1.0)
    }
}

/// Cell margins (padding) — a newtype over `TwipsEdgeInsets`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellMargins(pub TwipsEdgeInsets);

impl CellMargins {
    /// Default left/right cell margin: 108 twips (~0.075 inch).
    pub const DEFAULT_LR: Twips = Twips::new(108);

    /// Create new cell margins from individual values.
    pub fn new(top: Twips, right: Twips, bottom: Twips, left: Twips) -> Self {
        Self(TwipsEdgeInsets::new(top, right, bottom, left))
    }
}

impl Default for CellMargins {
    fn default() -> Self {
        Self::new(
            Twips::new(0),
            Self::DEFAULT_LR,
            Twips::new(0),
            Self::DEFAULT_LR,
        )
    }
}

impl std::ops::Deref for CellMargins {
    type Target = TwipsEdgeInsets;
    fn deref(&self) -> &TwipsEdgeInsets {
        &self.0
    }
}

impl std::ops::DerefMut for CellMargins {
    fn deref_mut(&mut self) -> &mut TwipsEdgeInsets {
        &mut self.0
    }
}

/// Border line style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderStyle {
    None,
    Single,
    Double,
    Dashed,
    Dotted,
}

/// A single border edge definition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct BorderDef {
    pub style: BorderStyle,
    /// Width in eighths of a point (OOXML native for `w:sz`).
    pub size: EighthPoints,
    pub color: Color,
    /// Space between border and content in points (OOXML `w:space`).
    pub space: Pt,
}

impl BorderDef {
    /// Default border width when `w:sz` is absent: 4 eighth-points (0.5pt).
    /// Not defined by OOXML spec; matches Microsoft Word's behavior.
    pub const DEFAULT_SIZE: EighthPoints = EighthPoints::new(4);

    /// Create a single-style border with given size (eighths of a point) and color.
    pub fn single(size: i64, color: Color) -> Self {
        Self {
            style: BorderStyle::Single,
            size: EighthPoints::new(size),
            color,
            space: Pt::ZERO,
        }
    }

    /// Returns true if this border should be drawn.
    pub fn is_visible(&self) -> bool {
        self.style != BorderStyle::None && self.size.is_positive()
    }
}

impl Default for BorderDef {
    fn default() -> Self {
        Self {
            style: BorderStyle::Single,
            size: Self::DEFAULT_SIZE,
            color: Color::BLACK,
            space: Pt::ZERO,
        }
    }
}

/// Table-level border definitions.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct TableBorders {
    pub top: BorderDef,
    pub bottom: BorderDef,
    pub left: BorderDef,
    pub right: BorderDef,
    /// Horizontal borders between rows.
    pub inside_h: BorderDef,
    /// Vertical borders between columns.
    pub inside_v: BorderDef,
}

/// Per-cell border overrides.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct CellBorders {
    pub top: Option<BorderDef>,
    pub bottom: Option<BorderDef>,
    pub left: Option<BorderDef>,
    pub right: Option<BorderDef>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub grid_cols: Vec<Twips>,
    pub default_cell_margins: Option<CellMargins>,
    pub cell_spacing: Option<Spacing>,
    pub borders: Option<TableBorders>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// Minimum row height from `w:trHeight`.
    pub height: Option<Twips>,
}

/// Vertical merge state for a table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalMerge {
    Restart,
    Continue,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TableCell {
    pub blocks: Vec<Block>,
    pub width: Option<Twips>,
    pub grid_span: u32,
    pub vertical_merge: Option<VerticalMerge>,
    pub cell_margins: Option<CellMargins>,
    pub cell_borders: Option<CellBorders>,
    /// Background fill color from `w:shd`.
    pub shading: Option<Color>,
}

impl TableCell {
    pub fn is_vmerge_continue(&self) -> bool {
        self.vertical_merge == Some(VerticalMerge::Continue)
    }
}

impl Color {
    pub fn from_hex(hex: &str) -> Option<Self> {
        if hex.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
        let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
        let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
        Some(Color { r, g, b })
    }
}

impl Spacing {
    /// Spacing before in points, defaulting to zero.
    pub fn before_pt(&self) -> Pt {
        self.before.map(Pt::from).unwrap_or(Pt::ZERO)
    }

    /// Spacing after in points, defaulting to zero.
    pub fn after_pt(&self) -> Pt {
        self.after.map(Pt::from).unwrap_or(Pt::ZERO)
    }

    /// Line spacing semantic value.
    /// For `Auto`: multiplier (1.0 = single spacing).
    /// For `Exact`/`AtLeast`: value in points.
    /// Returns `None` if not explicitly set.
    pub fn line_spacing(&self) -> Option<LineSpacing> {
        self.line.map(|v| match self.line_rule {
            LineRule::Auto => LineSpacing::Multiplier(i64::from(v) as f32 / 240.0),
            LineRule::Exact => LineSpacing::Fixed(Pt::from(v)),
            LineRule::AtLeast => LineSpacing::AtLeast(Pt::from(v)),
        })
    }

    /// Line spacing in points with a fixed fallback of single-line (240 twips = 12pt).
    pub fn line_pt(&self) -> Pt {
        let single_line = Pt::from(Twips::new(240));
        match self.line_spacing() {
            Some(LineSpacing::Fixed(pt) | LineSpacing::AtLeast(pt)) => pt,
            Some(LineSpacing::Multiplier(m)) => single_line * m,
            None => single_line,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn color_from_hex_valid() {
        let c = Color::from_hex("FF8000").unwrap();
        assert_eq!(
            c,
            Color {
                r: 255,
                g: 128,
                b: 0
            }
        );
    }

    #[test]
    fn color_from_hex_invalid() {
        assert!(Color::from_hex("GGGGGG").is_none());
        assert!(Color::from_hex("FF00").is_none());
    }

    #[test]
    fn font_size_conversion() {
        use crate::dimension::HalfPoints;
        let hp = HalfPoints::new(24);
        assert_eq!(f32::from(hp), 12.0);
    }

    #[test]
    fn default_font_size() {
        assert_eq!(f32::from(Document::DEFAULT_FONT_SIZE), 12.0);
    }

    #[test]
    fn emu_to_pt_conversion() {
        use crate::dimension::{Emu, Pt};
        let pt = f32::from(Pt::from(Emu::new(914400)));
        assert!((pt - 72.0).abs() < 0.01);
        assert_eq!(f32::from(Pt::from(Emu::new(0))), 0.0);
    }

    #[test]
    fn spacing_conversion() {
        let s = Spacing {
            before: Some(Twips::new(240)),
            after: Some(Twips::new(120)),
            line: Some(Twips::new(360)),
            ..Default::default()
        };
        assert_eq!(f32::from(s.before_pt()), 12.0);
        assert_eq!(f32::from(s.after_pt()), 6.0);
        assert_eq!(f32::from(s.line_pt()), 18.0);
    }

    #[test]
    fn rel_id_from_str() {
        let rid: RelId = "rId5".into();
        assert_eq!(rid.as_str(), "rId5");
        let rid2: RelId = String::from("rId6").into();
        assert_eq!(rid2.as_str(), "rId6");
    }
}
