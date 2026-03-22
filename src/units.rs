//! OOXML string constants and rendering defaults.
//!
//! Unit conversion is handled by the `dimension` module. This module retains
//! string constants used during XML parsing and rendering defaults that are
//! not yet derived from the document.

use crate::dimension::Pt;

// --- Legacy conversion function (re-exported from model for backward compat) ---

/// Convert English Metric Units to points.
pub fn emu_to_pt(emu: u64) -> f32 {
    emu as f32 / 914_400.0 * 72.0
}

// --- OOXML document defaults ---

/// Default font family when nothing is specified.
pub const DEFAULT_FONT_FAMILY: &str = "Helvetica";

/// OOXML width type for twips.
pub const WIDTH_TYPE_DXA: &str = "dxa";

/// OOXML underline value for "no underline".
pub const UNDERLINE_NONE: &str = "none";

// --- Rendering defaults ---
// These are not from the OOXML spec but are reasonable rendering choices.
// TODO: FLOAT_TEXT_GAP should be replaced by parsing wp:distL/distR/distT/distB
// from the floating image element.

/// Gap between a floating image and adjacent text.
/// Workaround: OOXML specifies this per-image via wp:distL/distR attributes,
/// which we don't parse yet.
pub const FLOAT_TEXT_GAP: Pt = Pt::new(4.0);

/// Minimum tab fragment width for line fitting.
/// Prevents tabs from collapsing to zero width during line breaking.
pub const MIN_TAB_WIDTH: Pt = Pt::new(12.0);

/// Fallback tab advance when no stops and no default interval.
pub const TAB_FALLBACK: Pt = Pt::new(36.0);

/// Underline offset below the text baseline.
pub const UNDERLINE_Y_OFFSET: Pt = Pt::new(2.0);
