//! Paint phase — iterate DrawCommands and emit Skia PDF canvas operations.
//!
//! Text is rendered via `TextBlob` with explicit per-glyph x-positions
//! to avoid Skia's PDF backend rounding glyph advances to integers.
//! This is critical on Linux where metric-compatible substitute fonts
//! (Carlito, Caladea) trigger per-character `Tj` emission with
//! integer-rounded advances, producing visible gaps within words.

use skia_safe::{pdf, Data, Font, FontMgr, GlyphId, Paint, TextBlobBuilder};

use crate::render::dimension::Pt;
use crate::render::error::RenderError;
use crate::render::fonts;
use crate::render::geometry::PtOffset;
use crate::render::layout::draw_command::{DrawCommand, LayoutedPage};
use crate::render::skia_conv::{to_color4f, to_line, to_point, to_rect, to_size};

/// Compute per-glyph x-positions from glyph advance widths.
///
/// §17.3.2.35 w:spacing — character spacing is added between every glyph.
/// The measurer applies spacing per character (Unicode scalar), so we
/// must apply it per glyph here to match. When a ligature maps multiple
/// characters to one glyph, the glyph count is smaller and the spacing
/// adjusts naturally (fewer inter-glyph gaps = less total spacing).
fn compute_glyph_positions(
    glyphs: &[GlyphId],
    font: &Font,
    base_x: f32,
    char_spacing: f32,
) -> Vec<f32> {
    let glyph_count = glyphs.len();
    let mut widths = vec![0.0f32; glyph_count];
    font.get_widths(glyphs, &mut widths);

    let mut positions = Vec::with_capacity(glyph_count);
    let mut x = base_x;
    for (i, &advance) in widths.iter().enumerate() {
        positions.push(x);
        if i < glyph_count - 1 {
            x += advance + char_spacing;
        }
    }
    positions
}

/// Render a text fragment as a positioned `TextBlob`.
///
/// Uses `TextBlobBuilder::alloc_run_pos_h` to embed glyph IDs with
/// fractional x-positions directly into the PDF content stream.
/// Positions are absolute page coordinates; the blob is drawn at origin.
fn draw_text_blob(
    canvas: &skia_safe::Canvas,
    text: &str,
    position: PtOffset,
    font: &Font,
    char_spacing: Pt,
    paint: &Paint,
) {
    let glyphs = font.str_to_glyphs_vec(text);
    if glyphs.is_empty() {
        log::warn!("[paint] text produced no glyphs: {:?}", &text[..text.len().min(40)]);
        return;
    }

    let x_positions = compute_glyph_positions(
        &glyphs,
        font,
        f32::from(position.x),
        f32::from(char_spacing),
    );

    let y = f32::from(position.y);

    let mut builder = TextBlobBuilder::new();
    let (glyph_buf, pos_buf) = builder.alloc_run_pos_h(font, glyphs.len(), y, None);
    glyph_buf.copy_from_slice(&glyphs);
    pos_buf.copy_from_slice(&x_positions);

    let blob = builder
        .make()
        .expect("TextBlobBuilder::make failed for non-empty glyph run");

    log::trace!(
        "[paint] blob {} glyphs, x=[{:.2}..{:.2}] y={:.2}",
        glyphs.len(),
        x_positions.first().unwrap(),
        x_positions.last().unwrap(),
        y,
    );

    canvas.draw_text_blob(&blob, (0.0, 0.0), paint);
}

/// Render laid-out pages to PDF bytes via Skia.
pub fn render_to_pdf(pages: &[LayoutedPage], font_mgr: &FontMgr) -> Result<Vec<u8>, RenderError> {
    let mut pdf_bytes: Vec<u8> = Vec::new();
    let mut doc = pdf::new_document(&mut pdf_bytes, None);

    for page in pages {
        let mut on_page = doc.begin_page(to_size(page.page_size), None);
        {
            let canvas = on_page.canvas();
            render_page(canvas, page, font_mgr);
        }
        doc = on_page.end_page();
    }

    doc.close();
    Ok(pdf_bytes)
}

fn render_page(canvas: &skia_safe::Canvas, page: &LayoutedPage, font_mgr: &FontMgr) {
    for cmd in &page.commands {
        match cmd {
            DrawCommand::Text {
                position,
                text,
                font_family,
                char_spacing,
                font_size,
                bold,
                italic,
                color,
            } => {
                let font = fonts::make_font(font_mgr, font_family, *font_size, *bold, *italic);
                log::trace!(
                    "[paint] '{}' → font='{}' size={:.1}pt bold={} italic={}",
                    &text[..text.len().min(30)],
                    font.typeface().family_name(),
                    font_size.raw(),
                    bold,
                    italic,
                );
                let mut paint = Paint::default();
                paint.set_anti_alias(true);
                paint.set_color4f(to_color4f(*color), None);

                draw_text_blob(canvas, text, *position, &font, *char_spacing, &paint);
            }
            DrawCommand::Underline { line, color, width }
            | DrawCommand::Line { line, color, width } => {
                let mut paint = Paint::default();
                paint.set_anti_alias(true);
                paint.set_stroke(true);
                paint.set_stroke_width(f32::from(*width));
                paint.set_color4f(to_color4f(*color), None);

                let (start, end) = to_line(*line);
                canvas.draw_line(start, end, &paint);
            }
            DrawCommand::Image { rect, image_data } => {
                let skia_data = Data::new_copy(image_data);
                if let Some(image) = skia_safe::Image::from_encoded(skia_data) {
                    canvas.draw_image_rect(image, None, to_rect(*rect), &Paint::default());
                }
            }
            DrawCommand::Rect { rect, color } => {
                let mut paint = Paint::default();
                paint.set_anti_alias(false);
                paint.set_color4f(to_color4f(*color), None);
                canvas.draw_rect(to_rect(*rect), &paint);
            }
            DrawCommand::LinkAnnotation { rect, url } => {
                let mut url_bytes = url.as_bytes().to_vec();
                url_bytes.push(0);
                let url_data = Data::new_copy(&url_bytes);
                canvas.annotate_rect_with_url(to_rect(*rect), &url_data);
            }
            DrawCommand::InternalLink { rect, destination } => {
                let mut name_bytes = destination.as_bytes().to_vec();
                name_bytes.push(0);
                let name_data = Data::new_copy(&name_bytes);
                canvas.annotate_link_to_destination(to_rect(*rect), &name_data);
            }
            DrawCommand::NamedDestination { position, name } => {
                let mut name_bytes = name.as_bytes().to_vec();
                name_bytes.push(0);
                let name_data = Data::new_copy(&name_bytes);
                canvas.annotate_named_destination(to_point(*position), &name_data);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::geometry::PtSize;
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    fn test_font_mgr() -> FontMgr {
        FontMgr::new()
    }

    fn test_font(font_mgr: &FontMgr) -> Font {
        fonts::make_font(font_mgr, "Helvetica", Pt::new(12.0), false, false)
    }

    // ── compute_glyph_positions ─────────────────────────────────────

    #[test]
    fn glyph_positions_single_glyph() {
        let font_mgr = test_font_mgr();
        let font = test_font(&font_mgr);
        let glyphs = font.str_to_glyphs_vec("A");
        assert_eq!(glyphs.len(), 1);

        let positions = compute_glyph_positions(&glyphs, &font, 100.0, 0.0);
        assert_eq!(positions.len(), 1);
        assert_eq!(positions[0], 100.0, "single glyph starts at base_x");
    }

    #[test]
    fn glyph_positions_multiple_glyphs_no_spacing() {
        let font_mgr = test_font_mgr();
        let font = test_font(&font_mgr);
        let glyphs = font.str_to_glyphs_vec("ABC");

        let positions = compute_glyph_positions(&glyphs, &font, 50.0, 0.0);
        assert_eq!(positions.len(), 3);
        assert_eq!(positions[0], 50.0, "first glyph at base_x");
        assert!(
            positions[1] > positions[0],
            "second glyph advances from first"
        );
        assert!(
            positions[2] > positions[1],
            "third glyph advances from second"
        );
    }

    #[test]
    fn glyph_positions_char_spacing_adds_per_glyph() {
        let font_mgr = test_font_mgr();
        let font = test_font(&font_mgr);
        let glyphs = font.str_to_glyphs_vec("AB");

        let spacing = 5.0;
        let without = compute_glyph_positions(&glyphs, &font, 0.0, 0.0);
        let with = compute_glyph_positions(&glyphs, &font, 0.0, spacing);

        assert_eq!(without[0], 0.0);
        assert_eq!(with[0], 0.0);
        let delta = with[1] - without[1];
        assert!(
            (delta - spacing).abs() < 0.01,
            "second glyph shifted by char_spacing: delta={delta}, expected={spacing}"
        );
    }

    #[test]
    fn glyph_positions_spacing_accumulates() {
        let font_mgr = test_font_mgr();
        let font = test_font(&font_mgr);
        let glyphs = font.str_to_glyphs_vec("ABCD");

        let spacing = 3.0;
        let without = compute_glyph_positions(&glyphs, &font, 0.0, 0.0);
        let with = compute_glyph_positions(&glyphs, &font, 0.0, spacing);

        // Each glyph after the first should be shifted by i * spacing
        for i in 1..4 {
            let expected_shift = spacing * i as f32;
            let actual_shift = with[i] - without[i];
            assert!(
                (actual_shift - expected_shift).abs() < 0.01,
                "glyph {i}: shift={actual_shift:.2}, expected={expected_shift:.2}"
            );
        }
    }

    #[test]
    fn glyph_positions_fractional_precision() {
        let font_mgr = test_font_mgr();
        let font = test_font(&font_mgr);
        let glyphs = font.str_to_glyphs_vec("Hello world");

        let positions = compute_glyph_positions(&glyphs, &font, 72.5, 0.0);

        // All positions must be fractional (not rounded to integers).
        // At least some interior positions should have non-zero fractional parts.
        let has_fractional = positions
            .iter()
            .any(|p| (p - p.round()).abs() > 0.001);
        assert!(
            has_fractional,
            "positions should preserve fractional precision: {:?}",
            positions
        );
    }

    // ── draw_text_blob integration ──────────────────────────────────

    #[test]
    fn render_text_command_produces_pdf() {
        let font_mgr = test_font_mgr();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "Hello world".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::ZERO,
                font_size: Pt::new(12.0),
                bold: false,
                italic: false,
                color: RgbColor::BLACK,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &font_mgr).expect("render_to_pdf must succeed");
        assert!(pdf_bytes.len() > 100, "PDF output must be non-trivial");
        assert_eq!(&pdf_bytes[..5], b"%PDF-", "output must be valid PDF");
    }

    #[test]
    fn render_text_with_char_spacing_produces_pdf() {
        let font_mgr = test_font_mgr();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "Spaced".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::new(2.0),
                font_size: Pt::new(14.0),
                bold: true,
                italic: false,
                color: RgbColor::BLACK,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &font_mgr).expect("render_to_pdf must succeed");
        assert!(pdf_bytes.len() > 100);
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }

    #[test]
    fn render_empty_text_produces_pdf() {
        let font_mgr = test_font_mgr();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: String::new(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::ZERO,
                font_size: Pt::new(12.0),
                bold: false,
                italic: false,
                color: RgbColor::BLACK,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &font_mgr).expect("empty text must not panic");
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }

    #[test]
    fn render_unicode_text_produces_pdf() {
        let font_mgr = test_font_mgr();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "Ärzte für Ökologie — 日本語".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::ZERO,
                font_size: Pt::new(11.0),
                bold: false,
                italic: false,
                color: RgbColor::BLACK,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &font_mgr).expect("unicode text must not panic");
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }
}
