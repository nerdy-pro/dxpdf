use skia_safe::{Font, FontMgr, FontStyle};

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

    /// Get a Skia Font for the given properties.
    fn make_font(&self, font_family: &str, font_size: f32, bold: bool, italic: bool) -> Font {
        let style = match (bold, italic) {
            (true, true) => FontStyle::bold_italic(),
            (true, false) => FontStyle::bold(),
            (false, true) => FontStyle::italic(),
            (false, false) => FontStyle::normal(),
        };

        let typeface = self
            .font_mgr
            .match_family_style(font_family, style)
            .or_else(|| self.font_mgr.match_family_style("Helvetica", style))
            .or_else(|| self.font_mgr.legacy_make_typeface(None::<&str>, style))
            .expect("no fallback typeface available");

        Font::from_typeface(typeface, font_size)
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
        let font = self.make_font(font_family, font_size, bold, italic);
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
        let font = self.make_font(font_family, font_size, bold, italic);
        let (_, metrics) = font.metrics();
        -metrics.ascent + metrics.descent + metrics.leading
    }
}
