use skia_safe::FontMgr;

use crate::render::fonts;

/// Measures text using Skia font metrics.
pub struct TextMeasurer {
    font_mgr: FontMgr,
}

impl TextMeasurer {
    pub fn new() -> Self {
        Self {
            font_mgr: FontMgr::new(),
        }
    }

    /// Measure the width of a text string in points.
    pub fn measure_width(
        &self,
        text: &str,
        font_family: &str,
        font_size: f32,
        bold: bool,
        italic: bool,
    ) -> f32 {
        let font = fonts::make_font(&self.font_mgr, font_family, font_size, bold, italic);
        let (width, _) = font.measure_str(text, None);
        width
    }

    /// Get the line height (ascent + descent + leading) for a font.
    pub fn line_height(
        &self,
        font_family: &str,
        font_size: f32,
        bold: bool,
        italic: bool,
    ) -> f32 {
        let font = fonts::make_font(&self.font_mgr, font_family, font_size, bold, italic);
        let (_, metrics) = font.metrics();
        -metrics.ascent + metrics.descent + metrics.leading
    }
}
