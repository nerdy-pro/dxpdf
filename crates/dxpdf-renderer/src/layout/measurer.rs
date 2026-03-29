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
    /// Also populates underline metrics on the FontProps.
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
        // §17.3.1.33: line height = ascent + descent (without leading).
        // Word's Auto line spacing base is ascent+descent. The leading
        // (sTypoLineGap) is NOT included — it would make table cells and
        // single-spaced paragraphs too tall.
        let line_height = Pt::new(-metrics.ascent + metrics.descent);
        let ascent = Pt::new(-metrics.ascent);

        (Pt::new(width), line_height, ascent)
    }

    /// Query font metrics for underline positioning.
    /// Returns (underline_position, underline_thickness) in points.
    /// Position is positive below baseline per Skia convention.
    pub fn underline_metrics(&self, font_props: &FontProps) -> (Pt, Pt) {
        let font = fonts::make_font(
            &self.font_mgr,
            &font_props.family,
            font_props.size,
            font_props.bold,
            font_props.italic,
        );
        let (_, metrics) = font.metrics();
        // Skia: underline_position() returns a negative value (below baseline).
        // We negate to get a positive offset below the baseline.
        // If the font doesn't provide metrics, log a warning and use descent as fallback.
        let raw_pos = metrics.underline_position();
        let raw_thick = metrics.underline_thickness();
        if raw_pos.is_none() || raw_thick.is_none() {
            log::warn!(
                "font '{}' ({:?}) missing underline metrics, using descent as fallback",
                font_props.family, font_props.size
            );
        }
        let position = Pt::new(-raw_pos.unwrap_or(metrics.descent));
        // Thickness fallback: 1pt (smallest visible line at 72dpi).
        let thickness = Pt::new(raw_thick.unwrap_or(1.0));
        (position, thickness)
    }

    /// Get line height for the default font (used for empty paragraphs).
    pub fn default_line_height(&self, family: &str, size: Pt) -> Pt {
        let font = fonts::make_font(&self.font_mgr, family, size, false, false);
        let (_, metrics) = font.metrics();
        Pt::new(-metrics.ascent + metrics.descent)
    }
}
