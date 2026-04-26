//! GSUB-aware text shaping for the emoji raster pipeline.
//!
//! Skia's `canvas.draw_str` performs only cmap-level codepoint→glyph
//! mapping; it does not apply OpenType GSUB lookups. For multi-codepoint
//! emoji sequences (`1️⃣` keycap, `👍🏿` modifier, `👨‍👩‍👧` ZWJ family),
//! the rasterizer needs the *ligated* single glyph, not the constituent
//! glyphs side-by-side.
//!
//! This module wraps `rustybuzz` (a pure-Rust HarfBuzz port) to produce
//! a closed-ADT [`ShapedRun`] from font bytes + text. The rasterizer
//! consumes that and calls `canvas.draw_glyphs` with the shaped glyph IDs.
//!
//! No system dependencies; rustybuzz is pure Rust.

use rustybuzz::{Face, UnicodeBuffer};
use thiserror::Error;

use crate::render::dimension::Pt;

// ─── Public ADTs ─────────────────────────────────────────────────────────────

/// One glyph in a [`ShapedRun`]. All offsets are in pixels at the
/// requested rasterization size; `cluster_byte` is the byte index in
/// the source `&str` of the codepoint that produced this glyph (used
/// for cluster→glyph mapping when needed).
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ShapedGlyph {
    /// Skia-compatible glyph id (u16). Skia stores glyph ids as u16 even
    /// though OpenType allows up to 65535; for fonts with > 65535 glyphs
    /// we degrade gracefully (see `shape_text` error path).
    pub id: u16,
    /// Pen advance in pixels.
    pub advance: Pt,
    /// Per-glyph horizontal offset in pixels (positive moves right).
    pub x_offset: Pt,
    /// Per-glyph vertical offset in pixels (positive moves up in HarfBuzz
    /// convention; Skia's y-axis is flipped — the rasterizer adapts).
    pub y_offset: Pt,
    /// Source byte index of the cluster this glyph belongs to.
    pub cluster_byte: u32,
}

/// Output of shaping one run of text.
#[derive(Clone, Debug)]
pub struct ShapedRun {
    pub glyphs: Vec<ShapedGlyph>,
    /// Sum of glyph advances in pixels — the rasterizer uses this to
    /// size the offscreen surface.
    pub total_advance: Pt,
}

#[derive(Debug, Error)]
pub enum ShapeError {
    #[error("rustybuzz refused to parse the font bytes")]
    InvalidFontData,
    #[error("font reports zero units_per_em — cannot scale glyph positions")]
    ZeroUnitsPerEm,
    #[error("glyph id {0} exceeds Skia's u16 range")]
    GlyphIdOutOfRange(u32),
}

// ─── Public API ──────────────────────────────────────────────────────────────

/// Shape `text` against `font_bytes` at `size_px`, returning the glyph
/// sequence + positions.
///
/// `size_px` is in raw pixels — already pre-multiplied by any super-sample
/// scale the caller wants.
pub fn shape_text(font_bytes: &[u8], text: &str, size_px: f32) -> Result<ShapedRun, ShapeError> {
    let face = Face::from_slice(font_bytes, 0).ok_or(ShapeError::InvalidFontData)?;
    let upem = face.units_per_em();
    if upem <= 0 {
        return Err(ShapeError::ZeroUnitsPerEm);
    }
    let scale = size_px / upem as f32;

    let mut buf = UnicodeBuffer::new();
    buf.push_str(text);
    let glyph_buf = rustybuzz::shape(&face, &[], buf);

    let infos = glyph_buf.glyph_infos();
    let positions = glyph_buf.glyph_positions();
    let mut glyphs = Vec::with_capacity(infos.len());
    let mut total = 0.0f32;

    for (info, pos) in infos.iter().zip(positions.iter()) {
        let id_u16 = u16::try_from(info.glyph_id)
            .map_err(|_| ShapeError::GlyphIdOutOfRange(info.glyph_id))?;
        let advance = pos.x_advance as f32 * scale;
        let x_offset = pos.x_offset as f32 * scale;
        let y_offset = pos.y_offset as f32 * scale;
        glyphs.push(ShapedGlyph {
            id: id_u16,
            advance: Pt::new(advance),
            x_offset: Pt::new(x_offset),
            y_offset: Pt::new(y_offset),
            cluster_byte: info.cluster,
        });
        total += advance;
    }

    Ok(ShapedRun {
        glyphs,
        total_advance: Pt::new(total),
    })
}

// ─── Tests ───────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use skia_safe::{FontMgr, FontStyle, Typeface};

    /// Pull bytes from any host system font for ASCII tests. Skips with
    /// a message if no system typeface is available.
    fn any_system_font_bytes() -> Option<Vec<u8>> {
        let mgr = FontMgr::new();
        let tf = mgr.legacy_make_typeface(None::<&str>, FontStyle::normal())?;
        tf.to_font_data().map(|(b, _)| b)
    }

    /// Pull bytes from a host emoji typeface, or skip the test cleanly.
    fn host_emoji_font_bytes() -> Option<Vec<u8>> {
        let mgr = FontMgr::new();
        for name in [
            "Apple Color Emoji",
            "Segoe UI Emoji",
            "Noto Color Emoji",
            "Twitter Color Emoji",
        ] {
            if let Some(tf) = mgr.match_family_style(name, FontStyle::normal()) {
                if tf.family_name().eq_ignore_ascii_case(name) {
                    if let Some((bytes, _)) = tf.to_font_data() {
                        return Some(bytes);
                    }
                }
            }
        }
        None
    }

    /// Skia accepts glyph ids that rustybuzz produces — invariant we rely
    /// on when we feed shape_text output into `canvas.draw_glyphs`.
    fn skia_typeface_for(bytes: &[u8]) -> Typeface {
        let mgr = FontMgr::new();
        mgr.new_from_data(&skia_safe::Data::new_copy(bytes), 0)
            .expect("skia must accept the same bytes rustybuzz parses")
    }

    /// Y1 — ASCII does not ligate. "abc" produces three glyphs against
    /// any real font. Regression guard: shaping must not ligate non-emoji
    /// runs.
    #[test]
    fn y1_ascii_does_not_ligate() {
        let bytes = match any_system_font_bytes() {
            Some(b) => b,
            None => {
                eprintln!("skipping Y1: no system typeface");
                return;
            }
        };
        let run = shape_text(&bytes, "abc", 24.0).expect("shape");
        assert_eq!(run.glyphs.len(), 3, "ASCII text must not ligate");
        for g in &run.glyphs {
            assert!(g.advance.raw() > 0.0);
        }
    }

    /// Y2 — Keycap "1\u{FE0F}\u{20E3}" ligates to one glyph against
    /// the host emoji typeface (skipped if no emoji font on host).
    #[test]
    fn y2_keycap_ligates_to_one_glyph() {
        let bytes = match host_emoji_font_bytes() {
            Some(b) => b,
            None => {
                eprintln!("skipping Y2: no host color emoji typeface");
                return;
            }
        };
        let run = shape_text(&bytes, "1\u{FE0F}\u{20E3}", 24.0).expect("shape");
        assert_eq!(
            run.glyphs.len(),
            1,
            "keycap must ligate to one glyph (got {} glyphs)",
            run.glyphs.len()
        );
        assert!(run.glyphs[0].advance.raw() > 0.0);
    }

    /// Y3 — Modifier sequence "👍🏿" ligates to one glyph.
    #[test]
    fn y3_modifier_sequence_ligates() {
        let bytes = match host_emoji_font_bytes() {
            Some(b) => b,
            None => {
                eprintln!("skipping Y3: no host color emoji typeface");
                return;
            }
        };
        let run = shape_text(&bytes, "\u{1F44D}\u{1F3FF}", 24.0).expect("shape");
        assert_eq!(run.glyphs.len(), 1, "modifier sequence must ligate");
    }

    /// Y4 — ZWJ family ligates to one glyph.
    #[test]
    fn y4_zwj_family_ligates() {
        let bytes = match host_emoji_font_bytes() {
            Some(b) => b,
            None => {
                eprintln!("skipping Y4: no host color emoji typeface");
                return;
            }
        };
        let run =
            shape_text(&bytes, "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}", 24.0).expect("shape");
        assert_eq!(run.glyphs.len(), 1, "ZWJ family must ligate");
    }

    /// Y5 — Garbage input fails cleanly.
    #[test]
    fn y5_invalid_font_bytes_errors() {
        let garbage = vec![0xFFu8; 64];
        let result = shape_text(&garbage, "hi", 12.0);
        assert!(matches!(result, Err(ShapeError::InvalidFontData)));
    }

    /// Y6 — empty input returns an empty run with zero advance.
    #[test]
    fn y6_empty_text_yields_empty_run() {
        let bytes = match any_system_font_bytes() {
            Some(b) => b,
            None => {
                eprintln!("skipping Y6: no system typeface");
                return;
            }
        };
        let run = shape_text(&bytes, "", 24.0).expect("shape");
        assert!(run.glyphs.is_empty());
        assert_eq!(run.total_advance.raw(), 0.0);
    }

    /// Y7 — the glyph ids that rustybuzz produces are valid for Skia's
    /// `canvas.draw_glyphs`. Verified by checking that the same Skia
    /// typeface (built from the same bytes) reports a non-zero advance
    /// for each glyph id.
    #[test]
    fn y7_glyph_ids_are_skia_compatible() {
        let bytes = match any_system_font_bytes() {
            Some(b) => b,
            None => {
                eprintln!("skipping Y7: no system typeface");
                return;
            }
        };
        let tf = skia_typeface_for(&bytes);
        let font = skia_safe::Font::from_typeface(tf, 24.0);
        let run = shape_text(&bytes, "abc", 24.0).expect("shape");
        let ids: Vec<u16> = run.glyphs.iter().map(|g| g.id).collect();
        let mut widths = vec![0f32; ids.len()];
        font.get_widths(&ids, &mut widths);
        for w in widths {
            assert!(
                w > 0.0,
                "every shaped glyph id must have a positive Skia advance"
            );
        }
    }
}
