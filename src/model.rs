use crate::units::{self, twips_to_pt, twips_to_pt_signed, DEFAULT_CELL_MARGIN_LR_TWIPS,
    DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_HALF_PTS, DEFAULT_TAB_STOP_TWIPS};

/// Root of the document tree.
#[derive(Debug, Clone, PartialEq)]
pub struct Document {
    pub blocks: Vec<Block>,
    /// Final section properties (from `w:body/w:sectPr`).
    pub final_section: Option<SectionProperties>,
    /// Default tab stop interval in twips.
    pub default_tab_stop: u32,
    /// Default font size in half-points.
    pub default_font_size: u32,
    /// Default font family.
    pub default_font_family: String,
    /// Default paragraph spacing.
    pub default_spacing: Spacing,
    /// Default table cell margins.
    pub default_cell_margins: CellMargins,
    /// Default paragraph spacing inside table cells.
    pub table_cell_spacing: Spacing,
    /// Default table borders (from table style).
    pub default_table_borders: TableBorders,
}

impl Default for Document {
    fn default() -> Self {
        Self {
            blocks: Vec::new(),
            final_section: None,
            default_tab_stop: DEFAULT_TAB_STOP_TWIPS,
            default_font_size: DEFAULT_FONT_SIZE_HALF_PTS,
            default_font_family: DEFAULT_FONT_FAMILY.to_string(),
            default_spacing: Spacing::default(),
            default_cell_margins: CellMargins::default(),
            table_cell_spacing: Spacing { after: Some(0), ..Default::default() },
            default_table_borders: TableBorders::default(),
        }
    }
}

/// A block-level element.
#[derive(Debug, Clone, PartialEq)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
}

/// Page size in twips.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageSize {
    pub width: u32,
    pub height: u32,
}

impl PageSize {
    pub fn width_pt(&self) -> f32 {
        twips_to_pt(self.width)
    }

    pub fn height_pt(&self) -> f32 {
        twips_to_pt(self.height)
    }
}

/// Page margins in twips.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageMargins {
    pub top: u32,
    pub right: u32,
    pub bottom: u32,
    pub left: u32,
}

impl PageMargins {
    pub fn top_pt(&self) -> f32 {
        twips_to_pt(self.top)
    }
    pub fn right_pt(&self) -> f32 {
        twips_to_pt(self.right)
    }
    pub fn bottom_pt(&self) -> f32 {
        twips_to_pt(self.bottom)
    }
    pub fn left_pt(&self) -> f32 {
        twips_to_pt(self.left)
    }
}

/// Section properties from `w:sectPr`.
#[derive(Debug, Clone, PartialEq)]
pub struct SectionProperties {
    pub page_size: Option<PageSize>,
    pub page_margins: Option<PageMargins>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Paragraph {
    pub properties: ParagraphProperties,
    pub runs: Vec<Inline>,
    pub floats: Vec<FloatingImage>,
    pub section_properties: Option<SectionProperties>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphProperties {
    pub alignment: Option<Alignment>,
    pub spacing: Option<Spacing>,
    pub indentation: Option<Indentation>,
    pub tab_stops: Vec<TabStop>,
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
    /// Position in twips.
    pub position: u32,
    pub stop_type: TabStopType,
}

impl TabStop {
    pub fn position_pt(&self) -> f32 {
        twips_to_pt(self.position)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

/// Spacing in twips.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Spacing {
    pub before: Option<u32>,
    pub after: Option<u32>,
    pub line: Option<u32>,
}

/// Indentation in twips.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Indentation {
    pub left: Option<u32>,
    pub right: Option<u32>,
    pub first_line: Option<i32>,
}

/// A type-safe wrapper for OOXML relationship IDs (e.g., "rId5").
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RelId(pub String);

impl RelId {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl<T: Into<String>> From<T> for RelId {
    fn from(s: T) -> Self {
        Self(s.into())
    }
}

/// A type-safe wrapper for image format hints (e.g., "png", "jpeg").
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FormatHint(pub String);

impl FormatHint {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl<T: Into<String>> From<T> for FormatHint {
    fn from(s: T) -> Self {
        Self(s.into())
    }
}

/// An inline-level element within a paragraph.
#[derive(Debug, Clone, PartialEq)]
pub enum Inline {
    TextRun(TextRun),
    LineBreak,
    Tab,
    Image(InlineImage),
}

#[derive(Debug, Clone, PartialEq)]
pub struct InlineImage {
    pub rel_id: RelId,
    pub width_pt: f32,
    pub height_pt: f32,
    pub data: Vec<u8>,
    pub format_hint: FormatHint,
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
    pub width_pt: f32,
    pub height_pt: f32,
    pub data: Vec<u8>,
    pub format_hint: FormatHint,
    pub offset_x_pt: f32,
    pub offset_y_pt: f32,
    pub wrap_side: WrapSide,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TextRun {
    pub text: String,
    pub properties: RunProperties,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RunProperties {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    /// Font size in half-points (OOXML native for w:sz).
    pub font_size: Option<u32>,
    pub font_family: Option<String>,
    pub color: Option<Color>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

/// Cell margins (padding) in twips.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellMargins {
    pub top: u32,
    pub right: u32,
    pub bottom: u32,
    pub left: u32,
}

impl Default for CellMargins {
    fn default() -> Self {
        Self {
            top: 0,
            right: DEFAULT_CELL_MARGIN_LR_TWIPS,
            bottom: 0,
            left: DEFAULT_CELL_MARGIN_LR_TWIPS,
        }
    }
}

impl CellMargins {
    pub fn top_pt(&self) -> f32 {
        twips_to_pt(self.top)
    }
    pub fn right_pt(&self) -> f32 {
        twips_to_pt(self.right)
    }
    pub fn bottom_pt(&self) -> f32 {
        twips_to_pt(self.bottom)
    }
    pub fn left_pt(&self) -> f32 {
        twips_to_pt(self.left)
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
    /// Width in eighths of a point (OOXML native for w:sz).
    pub size: u32,
    pub color: Color,
}

impl BorderDef {
    /// Width in points.
    pub fn width_pt(&self) -> f32 {
        self.size as f32 / 8.0
    }

    /// Returns true if this border should be drawn.
    pub fn is_visible(&self) -> bool {
        self.style != BorderStyle::None && self.size > 0
    }

    /// Color as an RGB tuple.
    pub fn color_rgb(&self) -> (u8, u8, u8) {
        (self.color.r, self.color.g, self.color.b)
    }
}

impl Default for BorderDef {
    fn default() -> Self {
        Self {
            style: BorderStyle::Single,
            size: 4, // 0.5pt
            color: Color { r: 0, g: 0, b: 0 },
        }
    }
}

/// Table-level border definitions.
#[derive(Debug, Clone, Copy, PartialEq)]
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

impl Default for TableBorders {
    fn default() -> Self {
        Self {
            top: BorderDef::default(),
            bottom: BorderDef::default(),
            left: BorderDef::default(),
            right: BorderDef::default(),
            inside_h: BorderDef::default(),
            inside_v: BorderDef::default(),
        }
    }
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
    pub grid_cols: Vec<u32>,
    pub default_cell_margins: Option<CellMargins>,
    pub cell_spacing: Option<Spacing>,
    pub borders: Option<TableBorders>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// Minimum row height in twips from `w:trHeight`.
    pub height: Option<u32>,
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
    pub width: Option<u32>,
    pub grid_span: u32,
    pub vertical_merge: Option<VerticalMerge>,
    pub cell_margins: Option<CellMargins>,
    pub cell_borders: Option<CellBorders>,
    /// Background fill color from `w:shd`.
    pub shading: Option<Color>,
}

impl TableCell {
    pub fn width_pt(&self) -> Option<f32> {
        self.width.map(twips_to_pt)
    }

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

impl RunProperties {
    pub fn font_size_pt(&self) -> f32 {
        self.font_size.map(|s| s as f32 / 2.0).unwrap_or(
            DEFAULT_FONT_SIZE_HALF_PTS as f32 / 2.0,
        )
    }

    pub fn font_size_pt_with_default(&self, default_half_pts: u32) -> f32 {
        self.font_size.unwrap_or(default_half_pts) as f32 / 2.0
    }
}

impl Spacing {
    pub fn before_pt(&self) -> f32 {
        self.before.map(twips_to_pt).unwrap_or(0.0)
    }

    pub fn after_pt(&self) -> f32 {
        self.after.map(twips_to_pt).unwrap_or(0.0)
    }

    /// Line spacing in points, or `None` if not explicitly set
    /// (meaning the font's natural line height should be used).
    pub fn line_pt_opt(&self) -> Option<f32> {
        self.line.map(twips_to_pt)
    }

    /// Line spacing in points with a fixed fallback of 12pt.
    /// Prefer `line_pt_opt()` when natural font height should be used.
    pub fn line_pt(&self) -> f32 {
        self.line_pt_opt().unwrap_or(twips_to_pt(240))
    }
}

impl Indentation {
    pub fn left_pt(&self) -> f32 {
        self.left.map(twips_to_pt).unwrap_or(0.0)
    }

    pub fn right_pt(&self) -> f32 {
        self.right.map(twips_to_pt).unwrap_or(0.0)
    }

    pub fn first_line_pt(&self) -> f32 {
        self.first_line.map(twips_to_pt_signed).unwrap_or(0.0)
    }
}

// Re-export emu_to_pt from units for backward compatibility
pub use units::emu_to_pt;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn color_from_hex_valid() {
        let c = Color::from_hex("FF8000").unwrap();
        assert_eq!(c, Color { r: 255, g: 128, b: 0 });
    }

    #[test]
    fn color_from_hex_invalid() {
        assert!(Color::from_hex("GGGGGG").is_none());
        assert!(Color::from_hex("FF00").is_none());
    }

    #[test]
    fn font_size_conversion() {
        let rp = RunProperties { font_size: Some(24), ..Default::default() };
        assert_eq!(rp.font_size_pt(), 12.0);
    }

    #[test]
    fn default_font_size() {
        let rp = RunProperties::default();
        assert_eq!(rp.font_size_pt(), 12.0);
    }

    #[test]
    fn emu_to_pt_conversion() {
        let pt = emu_to_pt(914400);
        assert!((pt - 72.0).abs() < 0.01);
        assert_eq!(emu_to_pt(0), 0.0);
    }

    #[test]
    fn spacing_conversion() {
        let s = Spacing { before: Some(240), after: Some(120), line: Some(360) };
        assert_eq!(s.before_pt(), 12.0);
        assert_eq!(s.after_pt(), 6.0);
        assert_eq!(s.line_pt(), 18.0);
    }

    #[test]
    fn rel_id_from_str() {
        let rid: RelId = "rId5".into();
        assert_eq!(rid.as_str(), "rId5");
        let rid2: RelId = String::from("rId6").into();
        assert_eq!(rid2.as_str(), "rId6");
    }

    #[test]
    fn format_hint_from_str() {
        let fh: FormatHint = "png".into();
        assert_eq!(fh.as_str(), "png");
    }
}
