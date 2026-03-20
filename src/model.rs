/// Root of the document tree.
#[derive(Debug, Clone, PartialEq)]
pub struct Document {
    pub blocks: Vec<Block>,
    /// Final section properties (from `w:body/w:sectPr`), applies to
    /// all content after the last mid-document section break.
    pub final_section: Option<SectionProperties>,
    /// Default tab stop interval in twips (from `word/settings.xml`).
    /// Defaults to 720 twips (0.5 inch) if not specified.
    pub default_tab_stop: u32,
    /// Default font size in half-points (from `word/styles.xml` docDefaults).
    /// Defaults to 24 (12pt) if not specified.
    pub default_font_size: u32,
    /// Default font family (from `word/styles.xml` docDefaults).
    pub default_font_family: String,
    /// Default paragraph spacing (from `word/styles.xml` docDefaults).
    pub default_spacing: Spacing,
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
    /// Width in twips.
    pub width: u32,
    /// Height in twips.
    pub height: u32,
}

impl PageSize {
    pub fn width_pt(&self) -> f32 {
        self.width as f32 / 20.0
    }

    pub fn height_pt(&self) -> f32 {
        self.height as f32 / 20.0
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
        self.top as f32 / 20.0
    }
    pub fn right_pt(&self) -> f32 {
        self.right as f32 / 20.0
    }
    pub fn bottom_pt(&self) -> f32 {
        self.bottom as f32 / 20.0
    }
    pub fn left_pt(&self) -> f32 {
        self.left as f32 / 20.0
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
    /// Floating/anchored images attached to this paragraph.
    pub floats: Vec<FloatingImage>,
    /// If present, this paragraph ends a section with these properties.
    /// All content from the previous section break up to (and including)
    /// this paragraph uses this section's page size and margins.
    pub section_properties: Option<SectionProperties>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphProperties {
    pub alignment: Option<Alignment>,
    pub spacing: Option<Spacing>,
    pub indentation: Option<Indentation>,
    /// Custom tab stops defined for this paragraph.
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
        self.position as f32 / 20.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

/// Spacing in twentieths of a point (OOXML native unit).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Spacing {
    pub before: Option<u32>,
    pub after: Option<u32>,
    pub line: Option<u32>,
}

/// Indentation in twentieths of a point.
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
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A type-safe wrapper for image format hints (e.g., "png", "jpeg").
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FormatHint(pub String);

impl FormatHint {
    pub fn new(hint: impl Into<String>) -> Self {
        Self(hint.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
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
    /// Relationship ID referencing the image file (e.g., "rId5").
    pub rel_id: RelId,
    /// Width in points (converted from EMUs at parse time).
    pub width_pt: f32,
    /// Height in points (converted from EMUs at parse time).
    pub height_pt: f32,
    /// Raw image bytes, populated after archive extraction.
    pub data: Vec<u8>,
    /// File extension hint for decoding (e.g., "png", "jpeg").
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
    /// Horizontal offset from margin in points.
    pub offset_x_pt: f32,
    /// Vertical offset from paragraph in points.
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

#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    pub rows: Vec<TableRow>,
    /// Column widths from `w:tblGrid` in twips.
    pub grid_cols: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TableCell {
    pub blocks: Vec<Block>,
    /// Cell width in twips from `w:tcW`, if specified.
    pub width: Option<u32>,
}

impl TableCell {
    pub fn width_pt(&self) -> Option<f32> {
        self.width.map(|w| w as f32 / 20.0)
    }
}

/// Convert English Metric Units to points (914400 EMU = 1 inch = 72 points).
pub fn emu_to_pt(emu: u64) -> f32 {
    emu as f32 / 914400.0 * 72.0
}

impl Color {
    /// Parse a 6-digit hex color string (e.g., "FF0000" for red).
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
    /// Font size in points (converts from half-points).
    /// Falls back to 12pt if not specified.
    pub fn font_size_pt(&self) -> f32 {
        self.font_size.map(|s| s as f32 / 2.0).unwrap_or(12.0)
    }

    /// Font size in points using a document-level default (in half-points).
    pub fn font_size_pt_with_default(&self, default_half_pts: u32) -> f32 {
        self.font_size
            .unwrap_or(default_half_pts) as f32 / 2.0
    }
}

impl Spacing {
    /// Convert `before` from twips to points.
    pub fn before_pt(&self) -> f32 {
        self.before.map(|v| v as f32 / 20.0).unwrap_or(0.0)
    }

    /// Convert `after` from twips to points.
    pub fn after_pt(&self) -> f32 {
        self.after.map(|v| v as f32 / 20.0).unwrap_or(0.0)
    }

    /// Convert `line` spacing from twips to points. Default is single spacing (240 twips = 12pt).
    pub fn line_pt(&self) -> f32 {
        self.line.map(|v| v as f32 / 20.0).unwrap_or(12.0)
    }
}

impl Indentation {
    pub fn left_pt(&self) -> f32 {
        self.left.map(|v| v as f32 / 20.0).unwrap_or(0.0)
    }

    pub fn right_pt(&self) -> f32 {
        self.right.map(|v| v as f32 / 20.0).unwrap_or(0.0)
    }

    pub fn first_line_pt(&self) -> f32 {
        self.first_line.map(|v| v as f32 / 20.0).unwrap_or(0.0)
    }
}

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
        // 914400 EMU = 1 inch = 72 points
        let pt = emu_to_pt(914400);
        assert!((pt - 72.0).abs() < 0.01);
        // 0 EMU = 0 points
        assert_eq!(emu_to_pt(0), 0.0);
    }

    #[test]
    fn spacing_conversion() {
        let s = Spacing { before: Some(240), after: Some(120), line: Some(360) };
        assert_eq!(s.before_pt(), 12.0);
        assert_eq!(s.after_pt(), 6.0);
        assert_eq!(s.line_pt(), 18.0);
    }
}
