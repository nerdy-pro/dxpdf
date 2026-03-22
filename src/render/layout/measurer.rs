use skia_safe::FontMgr;

use crate::dimension::Pt;
use crate::render::fonts;

/// Measures text using Skia font metrics.
/// Requires a shared `FontMgr` — use the same instance as the layout/paint pipeline.
pub struct TextMeasurer {
    font_mgr: FontMgr,
}

impl TextMeasurer {
    pub fn new(font_mgr: FontMgr) -> Self {
        Self { font_mgr }
    }

    /// Measure the width of a text string in points.
    pub fn measure_width(
        &self,
        text: &str,
        font_family: &str,
        font_size: Pt,
        bold: bool,
        italic: bool,
    ) -> Pt {
        let font = fonts::make_font(&self.font_mgr, font_family, font_size, bold, italic);
        let (width, _) = font.measure_str(text, None);
        Pt::new(width)
    }

    /// Get the line height (ascent + descent + leading) for a font.
    pub fn line_height(&self, font_family: &str, font_size: Pt, bold: bool, italic: bool) -> Pt {
        let font = fonts::make_font(&self.font_mgr, font_family, font_size, bold, italic);
        let (_, metrics) = font.metrics();
        Pt::new(-metrics.ascent + metrics.descent + metrics.leading)
    }

    /// Get the ascent (distance from baseline to top of line) for a font.
    /// Always positive.
    pub fn ascent(&self, font_family: &str, font_size: Pt, bold: bool, italic: bool) -> Pt {
        let font = fonts::make_font(&self.font_mgr, font_family, font_size, bold, italic);
        let (_, metrics) = font.metrics();
        Pt::new(-metrics.ascent)
    }
}
