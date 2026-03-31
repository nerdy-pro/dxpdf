//! Text measurer — wraps Skia FontMgr for real font metrics.

use skia_safe::FontMgr;

use crate::render::dimension::Pt;
use crate::render::fonts;

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
    /// Returns (width, TextMetrics).
    pub fn measure(
        &self,
        text: &str,
        font_props: &FontProps,
    ) -> (Pt, super::fragment::TextMetrics) {
        let font = fonts::make_font(
            &self.font_mgr,
            &font_props.family,
            font_props.size,
            font_props.bold,
            font_props.italic,
        );

        let (width, _bounds) = font.measure_str(text, None);
        let (_, metrics) = font.metrics();
        let text_metrics = super::fragment::TextMetrics {
            ascent: Pt::new(-metrics.ascent),
            descent: Pt::new(metrics.descent),
        };

        // §17.3.2.35: include character spacing in the measured width
        // so line fitting accounts for the extra inter-character space.
        let char_count = text.chars().count();
        let spacing_extra = if char_count > 0 {
            font_props.char_spacing * (char_count as f32)
        } else {
            Pt::ZERO
        };

        (Pt::new(width) + spacing_extra, text_metrics)
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
                font_props.family,
                font_props.size
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
        // ascent + descent (without leading)
        Pt::new(-metrics.ascent + metrics.descent)
    }
}
