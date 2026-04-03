//! Fragment conversion — transform Inline content into measured Fragments
//! for the line-fitting algorithm.

use std::rc::Rc;

use crate::model::RunProperties;

use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::resolve::color::RgbColor;
use crate::render::resolve::fonts::effective_font;

mod collect;
mod text;

pub use collect::{collect_fragments, FieldContext};

// ── Superscript / subscript rendering constants ───────────────────────────────
// §17.3.2.42: these ratios are "application-defined" per the spec; the values
// below match Word's rendering as documented in the OpenXML SDK reference.

/// Font size of super/subscript text as a fraction of the base font size.
pub(super) const SUPERSCRIPT_FONT_SIZE_RATIO: f32 = 0.58;

/// Superscript baseline shift: fraction of base ascent to raise the text by.
pub(super) const SUPERSCRIPT_ASCENT_OFFSET_RATIO: f32 = 0.33;

/// Subscript baseline shift: fraction of base character height to lower the text by.
pub(super) const SUBSCRIPT_HEIGHT_OFFSET_RATIO: f32 = 0.08;

/// Font properties needed for rendering a text fragment.
#[derive(Clone, Debug)]
pub struct FontProps {
    pub family: Rc<str>,
    pub size: Pt,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub char_spacing: Pt,
    /// Underline position from font metrics (positive = below baseline).
    pub underline_position: Pt,
    /// Underline thickness from font metrics.
    pub underline_thickness: Pt,
}

/// Font metrics for a specific font at a specific size.
/// Evaluated once by the measurer and carried through the pipeline.
#[derive(Clone, Copy, Debug)]
pub struct TextMetrics {
    /// Distance from baseline to top of glyphs (positive upward).
    pub ascent: Pt,
    /// Distance from baseline to bottom of glyphs (positive downward).
    pub descent: Pt,
}

impl TextMetrics {
    /// Total text height (ascent + descent).
    pub fn height(&self) -> Pt {
        self.ascent + self.descent
    }
}

/// §17.3.2.4: run-level border for rendering.
#[derive(Clone, Copy, Debug)]
pub struct FragmentBorder {
    pub width: Pt,
    pub color: RgbColor,
    pub space: Pt,
}

/// A measured fragment — the atomic unit for line fitting.
#[derive(Clone, Debug)]
pub enum Fragment {
    Text {
        text: Rc<str>,
        font: FontProps,
        color: RgbColor,
        /// §17.3.2.32: run-level shading (background color behind text).
        shading: Option<RgbColor>,
        /// §17.3.2.4: run-level border (box around text).
        border: Option<FragmentBorder>,
        /// Full width including trailing whitespace (used for positioning).
        width: Pt,
        /// Width excluding trailing whitespace (used for line-break overflow checking).
        /// Trailing whitespace is allowed to hang past the margin per Word behavior.
        trimmed_width: Pt,
        /// Font metrics (ascent + descent = text height).
        metrics: TextMetrics,
        hyperlink_url: Option<String>,
        baseline_offset: Pt,
        /// Horizontal offset for drawing text within the fragment width.
        /// Used for right/center-justified list labels where the text is
        /// positioned within a wider fragment. Default: Pt::ZERO.
        text_offset: Pt,
    },
    Image {
        size: PtSize,
        rel_id: String,
        image_data: Option<std::rc::Rc<[u8]>>,
    },
    Tab {
        line_height: Pt,
        /// Override minimum width for line fitting (default: MIN_TAB_WIDTH).
        fitting_width: Option<Pt>,
    },
    LineBreak {
        line_height: Pt,
    },
    /// §17.3.3.1: column break — forces content to the next column.
    ColumnBreak,
    /// Named destination (bookmark target) — zero-width marker.
    Bookmark {
        name: String,
    },
}

impl Fragment {
    pub fn width(&self) -> Pt {
        match self {
            Fragment::Text { width, .. } => *width,
            Fragment::Image { size, .. } => size.width,
            Fragment::Tab { fitting_width, .. } => fitting_width.unwrap_or(MIN_TAB_WIDTH),
            Fragment::LineBreak { .. } | Fragment::ColumnBreak | Fragment::Bookmark { .. } => {
                Pt::ZERO
            }
        }
    }

    /// Width for overflow checking — excludes trailing whitespace on text fragments.
    pub fn trimmed_width(&self) -> Pt {
        match self {
            Fragment::Text { trimmed_width, .. } => *trimmed_width,
            other => other.width(),
        }
    }

    pub fn height(&self) -> Pt {
        match self {
            Fragment::Text { metrics, .. } => metrics.height(),
            Fragment::Image { size, .. } => size.height,
            Fragment::Tab { line_height, .. } | Fragment::LineBreak { line_height } => *line_height,
            Fragment::ColumnBreak | Fragment::Bookmark { .. } => Pt::ZERO,
        }
    }

    pub fn is_line_break(&self) -> bool {
        matches!(self, Fragment::LineBreak { .. } | Fragment::ColumnBreak)
    }

    /// Get font properties if this is a text fragment.
    pub fn font_props(&self) -> Option<&FontProps> {
        match self {
            Fragment::Text { font, .. } => Some(font),
            _ => None,
        }
    }
}

/// §17.3.1.37: minimum tab fragment width for line fitting.
/// Tabs resolve to tab stops defined on the paragraph; this constant is only
/// used as the fragment width during line breaking (actual tab position is
/// computed during paragraph layout).
pub const MIN_TAB_WIDTH: Pt = Pt::new(1.0);

/// Extract font properties from RunProperties with a default font family fallback.
pub fn font_props_from_run(
    rp: &RunProperties,
    default_family: &str,
    default_size: Pt,
) -> FontProps {
    let family = effective_font(&rp.fonts).unwrap_or(default_family);

    let size = rp.font_size.map(Pt::from).unwrap_or(default_size);

    let char_spacing = rp.spacing.map(Pt::from).unwrap_or(Pt::ZERO);

    FontProps {
        family: Rc::from(family),
        size,
        bold: rp.bold.unwrap_or(false),
        italic: rp.italic.unwrap_or(false),
        underline: rp.underline.is_some(),
        char_spacing,
        // Populated by the measurer from Skia font metrics.
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    }
}

/// Convert a number to lowercase Roman numerals.
pub fn to_roman_lower(mut n: u32) -> String {
    const VALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut s = String::new();
    for &(val, sym) in &VALS {
        while n >= val {
            s.push_str(sym);
            n -= val;
        }
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn font_props_default_fallback() {
        let rp = RunProperties::default();
        let fp = font_props_from_run(&rp, "Helvetica", Pt::new(12.0));
        assert_eq!(&*fp.family, "Helvetica");
        assert_eq!(fp.size.raw(), 12.0);
        assert!(!fp.bold);
        assert!(!fp.italic);
    }
}
