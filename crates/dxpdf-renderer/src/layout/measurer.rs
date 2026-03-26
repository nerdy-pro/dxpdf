//! Text measurer — wraps Skia FontMgr for real font metrics.

use skia_safe::FontMgr;

use crate::dimension::Pt;
use crate::fonts;

use super::fragment::FontProps;

/// Measures text using Skia fonts. Wraps a FontMgr reference.
pub struct TextMeasurer {
    font_mgr: FontMgr,
}

impl TextMeasurer {
    pub fn new(font_mgr: FontMgr) -> Self {
        Self { font_mgr }
    }

    /// Measure a text string with the given font properties.
    /// Returns (width, line_height, ascent).
    pub fn measure(&self, text: &str, font_props: &FontProps) -> (Pt, Pt, Pt) {
        let font = fonts::make_font(
            &self.font_mgr,
            &font_props.family,
            font_props.size,
            font_props.bold,
            font_props.italic,
        );

        let (width, _bounds) = font.measure_str(text, None);
        let (_, metrics) = font.metrics();
        let line_height = Pt::new(-metrics.ascent + metrics.descent + metrics.leading);
        let ascent = Pt::new(-metrics.ascent);

        (Pt::new(width), line_height, ascent)
    }

    /// Get line height for the default font (used for empty paragraphs).
    pub fn default_line_height(&self, family: &str, size: Pt) -> Pt {
        let font = fonts::make_font(&self.font_mgr, family, size, false, false);
        let (_, metrics) = font.metrics();
        Pt::new(-metrics.ascent + metrics.descent + metrics.leading)
    }
}
