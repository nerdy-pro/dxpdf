//! Unit conversion constants and helpers for OOXML measurements.

/// Number of twips per point (1 point = 20 twips).
pub const TWIPS_PER_POINT: f32 = 20.0;

/// Number of EMUs per inch.
pub const EMU_PER_INCH: f32 = 914400.0;

/// Number of points per inch.
pub const POINTS_PER_INCH: f32 = 72.0;

/// Convert twips to points.
pub fn twips_to_pt(twips: u32) -> f32 {
    twips as f32 / TWIPS_PER_POINT
}

/// Convert signed twips to points.
pub fn twips_to_pt_signed(twips: i32) -> f32 {
    twips as f32 / TWIPS_PER_POINT
}

/// Convert English Metric Units to points.
pub fn emu_to_pt(emu: u64) -> f32 {
    emu as f32 / EMU_PER_INCH * POINTS_PER_INCH
}

// --- Document defaults ---

/// Default font family when nothing is specified.
pub const DEFAULT_FONT_FAMILY: &str = "Helvetica";

/// Default font size in half-points (24 = 12pt).
pub const DEFAULT_FONT_SIZE_HALF_PTS: u32 = 24;

/// Default tab stop interval in twips (720 = 0.5 inch).
pub const DEFAULT_TAB_STOP_TWIPS: u32 = 720;

/// Default page margin in twips (1440 = 1 inch).
pub const DEFAULT_PAGE_MARGIN_TWIPS: u32 = 1440;

/// Default cell margin left/right in twips (108 = 5.4pt).
pub const DEFAULT_CELL_MARGIN_LR_TWIPS: u32 = 108;

// --- Layout defaults ---

/// US Letter page width in points (8.5 inches).
pub const US_LETTER_WIDTH_PT: f32 = 612.0;

/// US Letter page height in points (11 inches).
pub const US_LETTER_HEIGHT_PT: f32 = 792.0;

/// Default page margin in points (1 inch).
pub const DEFAULT_PAGE_MARGIN_PT: f32 = 72.0;

/// Gap between a floating image and adjacent text in points.
pub const FLOAT_TEXT_GAP_PT: f32 = 4.0;

/// Minimum tab fragment width for line fitting (in points).
pub const MIN_TAB_WIDTH_PT: f32 = 12.0;

/// Fallback tab advance when no stops and no default interval (in points).
pub const TAB_FALLBACK_PT: f32 = 36.0;

/// Underline offset below the text baseline (in points).
pub const UNDERLINE_Y_OFFSET: f32 = 2.0;

/// Underline stroke width (in points).
pub const UNDERLINE_STROKE_WIDTH: f32 = 0.5;

/// Table cell border stroke width (in points).
pub const TABLE_BORDER_WIDTH: f32 = 0.5;

/// Minimum table row height (in points).
pub const MIN_ROW_HEIGHT_PT: f32 = 12.0;

/// Minimum bottom padding in cells to prevent descender overlap (in points).
pub const MIN_CELL_BOTTOM_PAD_PT: f32 = 2.0;

/// Spacing after a table (in points).
pub const TABLE_AFTER_SPACING_PT: f32 = 8.0;

/// Maximum consecutive spaces before collapsing.
pub const SPACE_COLLAPSE_THRESHOLD: usize = 2;

/// OOXML width type for twips.
pub const WIDTH_TYPE_DXA: &str = "dxa";

/// OOXML underline value for "no underline".
pub const UNDERLINE_NONE: &str = "none";
