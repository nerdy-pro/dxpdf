use skia_safe::{Font, FontMgr};

use crate::dimension::Pt;
use crate::render::fonts;

/// A resolved font ready for measurement.
/// Created once per (family, size, bold, italic) combination; use it for
/// both metrics and width measurements to avoid redundant font creation.
pub struct MeasuredFont {
    font: Font,
}

/// Font metrics for a specific font configuration.
#[derive(Debug, Clone, Copy)]
pub struct FontMetrics {
    /// Distance from baseline to top of line (always positive).
    pub ascent: Pt,
    /// Total line height (ascent + descent + leading).
    pub line_height: Pt,
}

impl MeasuredFont {
    /// Get font metrics (ascent, line height).
    pub fn metrics(&self) -> FontMetrics {
        let (_, m) = self.font.metrics();
        FontMetrics {
            ascent: Pt::new(-m.ascent),
            line_height: Pt::new(-m.ascent + m.descent + m.leading),
        }
    }

    /// Measure the width of a text string.
    pub fn measure_width(&self, text: &str) -> Pt {
        let (width, _) = self.font.measure_str(text, None);
        Pt::new(width)
    }
}

/// Resolves fonts and measures text using Skia.
/// Requires a shared `FontMgr` — use the same instance as the layout/paint pipeline.
pub struct TextMeasurer {
    font_mgr: FontMgr,
}

impl TextMeasurer {
    pub fn new(font_mgr: FontMgr) -> Self {
        Self { font_mgr }
    }

    /// Resolve a font for the given properties. The returned `MeasuredFont`
    /// can be used for both metrics and width measurements.
    pub fn font(&self, font_family: &str, font_size: Pt, bold: bool, italic: bool) -> MeasuredFont {
        let font = fonts::make_font(&self.font_mgr, font_family, font_size, bold, italic);
        MeasuredFont { font }
    }
}
