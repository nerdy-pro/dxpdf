//! Paint phase — iterate DrawCommands and emit Skia PDF canvas operations.

use std::collections::HashMap;
use std::rc::Rc;

use skia_safe::{pdf, Data, FontMgr, Paint};

use crate::render::dimension::Pt;
use crate::render::error::RenderError;
use crate::render::fonts;
use crate::render::layout::draw_command::{DrawCommand, LayoutedPage};
use crate::render::skia_conv::{to_color4f, to_line, to_point, to_rect, to_size};

/// Render laid-out pages to PDF bytes via Skia.
pub fn render_to_pdf(pages: &[LayoutedPage], font_mgr: &FontMgr) -> Result<Vec<u8>, RenderError> {
    let mut pdf_bytes: Vec<u8> = Vec::new();
    let mut doc = pdf::new_document(&mut pdf_bytes, None);
    let mut font_cache = fonts::FontCache::new();
    // Cache decoded Skia images across pages, keyed by Rc pointer identity.
    // Avoids re-copying and re-decoding the same image bytes on every page
    // (e.g. a logo repeated in headers/footers).
    let mut image_cache: HashMap<*const [u8], skia_safe::Image> = HashMap::new();

    for page in pages {
        let mut on_page = doc.begin_page(to_size(page.page_size), None);
        {
            let canvas = on_page.canvas();
            render_page(canvas, page, font_mgr, &mut font_cache, &mut image_cache);
        }
        doc = on_page.end_page();
    }

    doc.close();
    Ok(pdf_bytes)
}

fn render_page(
    canvas: &skia_safe::Canvas,
    page: &LayoutedPage,
    font_mgr: &FontMgr,
    font_cache: &mut fonts::FontCache,
    image_cache: &mut HashMap<*const [u8], skia_safe::Image>,
) {
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
                let font =
                    font_cache.get(font_mgr, font_family, *font_size, *bold, *italic);
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

                if char_spacing.abs() > Pt::ZERO {
                    // §17.3.2.35 w:spacing — draw each character with
                    // explicit spacing to match the measured fragment width.
                    // Batch: convert text → glyphs → widths in two Skia calls
                    // instead of per-char measure_str + String allocation.
                    let glyphs = font.text_to_glyphs_vec(text.as_str());
                    let mut widths = vec![0f32; glyphs.len()];
                    font.get_widths(&glyphs, &mut widths);

                    let mut cursor = *position;
                    let mut buf = [0u8; 4];
                    for (i, ch) in text.chars().enumerate() {
                        let s = ch.encode_utf8(&mut buf);
                        canvas.draw_str(s, to_point(cursor), font, &paint);
                        // Use pre-computed glyph width (falls back to 0 for
                        // complex scripts where glyph count != char count).
                        let w = widths.get(i).copied().unwrap_or(0.0);
                        cursor.x += Pt::new(w) + *char_spacing;
                    }
                } else {
                    canvas.draw_str(text, to_point(*position), font, &paint);
                }
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
                let ptr_key: *const [u8] = Rc::as_ptr(image_data);
                if let Some(image) = image_cache.get(&ptr_key) {
                    canvas.draw_image_rect(image, None, to_rect(*rect), &Paint::default());
                } else {
                    let skia_data = Data::new_copy(image_data);
                    if let Some(image) = skia_safe::Image::from_encoded(skia_data) {
                        canvas.draw_image_rect(&image, None, to_rect(*rect), &Paint::default());
                        image_cache.insert(ptr_key, image);
                    }
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
    use crate::render::geometry::{PtOffset, PtSize};
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    fn test_font_mgr() -> FontMgr {
        FontMgr::new()
    }

    // ── render_to_pdf integration ───────────────────────────────────

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
