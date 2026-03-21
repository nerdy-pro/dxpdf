//! Unit conversion constants and helpers for OOXML measurements.
//!
//! All constants in this module are either:
//! - Defined by the OOXML / ISO 29500 specification
//! - Standard typographic / paper size values
//! - Clearly documented rendering defaults with rationale

// --- OOXML unit conversion factors (ISO 29500) ---

/// Number of twips per point (1 point = 20 twips).
pub const TWIPS_PER_POINT: f32 = 20.0;

/// Number of EMUs per inch (ISO 29500: 1 inch = 914400 EMUs).
pub const EMU_PER_INCH: f32 = 914400.0;

/// Number of points per inch (standard typographic unit).
pub const POINTS_PER_INCH: f32 = 72.0;

/// Half-points per point (OOXML font sizes are in half-points).
pub const HALF_POINTS_PER_POINT: f32 = 2.0;

/// Border size units per point (OOXML w:sz is in eighths of a point).
pub const BORDER_SIZE_UNITS_PER_POINT: f32 = 8.0;

// --- Unit conversion functions ---

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

/// Convert signed English Metric Units to points.
pub fn emu_to_pt_signed(emu: i64) -> f32 {
    emu as f32 / EMU_PER_INCH * POINTS_PER_INCH
}

// --- OOXML document defaults (ISO 29500) ---

/// Default font family when nothing is specified.
pub const DEFAULT_FONT_FAMILY: &str = "Helvetica";

/// Default font size in half-points (24 = 12pt, OOXML default).
pub const DEFAULT_FONT_SIZE_HALF_PTS: u32 = 24;

/// Default tab stop interval in twips (720 = 0.5 inch, OOXML default).
pub const DEFAULT_TAB_STOP_TWIPS: u32 = 720;

/// Default page margin in twips (1440 = 1 inch, OOXML default).
pub const DEFAULT_PAGE_MARGIN_TWIPS: u32 = 1440;

/// Default cell margin left/right in twips (108 twips, OOXML w:tblCellMar default).
pub const DEFAULT_CELL_MARGIN_LR_TWIPS: u32 = 108;

/// OOXML width type for twips.
pub const WIDTH_TYPE_DXA: &str = "dxa";

/// OOXML underline value for "no underline".
pub const UNDERLINE_NONE: &str = "none";

// --- Standard page dimensions ---

/// US Letter page width in points (8.5 inches).
pub const US_LETTER_WIDTH_PT: f32 = 612.0;

/// US Letter page height in points (11 inches).
pub const US_LETTER_HEIGHT_PT: f32 = 792.0;

/// Default page margin in points (1 inch).
pub const DEFAULT_PAGE_MARGIN_PT: f32 = 72.0;

// --- Rendering defaults ---
// These are not from the OOXML spec but are reasonable rendering choices.
// TODO: FLOAT_TEXT_GAP_PT should be replaced by parsing wp:distL/distR/distT/distB
// from the floating image element.

/// Gap between a floating image and adjacent text in points.
/// Workaround: OOXML specifies this per-image via wp:distL/distR attributes,
/// which we don't parse yet.
pub const FLOAT_TEXT_GAP_PT: f32 = 4.0;

/// Minimum tab fragment width for line fitting (in points).
/// Prevents tabs from collapsing to zero width during line breaking.
pub const MIN_TAB_WIDTH_PT: f32 = 12.0;

/// Fallback tab advance when no stops and no default interval (in points).
pub const TAB_FALLBACK_PT: f32 = 36.0;

/// Underline offset below the text baseline (in points).
pub const UNDERLINE_Y_OFFSET: f32 = 2.0;
