//! Paint phase — iterate DrawCommands and emit Skia PDF canvas operations.

use std::collections::HashMap;
use std::rc::Rc;

use skia_safe::{pdf, Data, FontMgr, Paint};

use crate::render::dimension::Pt;
use crate::render::error::RenderError;
use crate::render::fonts;
use crate::render::layout::draw_command::{DrawCommand, LayoutedPage};
use crate::render::resolve::images::MediaEntry;
use crate::render::skia_conv::{to_color4f, to_line, to_point, to_rect, to_size};

/// Target resolution for embedded images (pixels per inch).
/// 150 DPI balances print quality with file size; matches LibreOffice.
const IMAGE_TARGET_DPI: f32 = 150.0;
/// Conversion factor from PDF points to target pixels.
const IMAGE_DPI_SCALE: f32 = IMAGE_TARGET_DPI / 72.0;

/// Render laid-out pages to PDF bytes via Skia.
pub fn render_to_pdf(pages: &[LayoutedPage], font_mgr: &FontMgr) -> Result<Vec<u8>, RenderError> {
    let mut pdf_bytes: Vec<u8> = Vec::new();
    let pdf_metadata = pdf::Metadata {
        encoding_quality: Some(85),
        ..Default::default()
    };
    let mut doc = pdf::new_document(&mut pdf_bytes, Some(&pdf_metadata));
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
                let font = font_cache.get(font_mgr, font_family, *font_size, *bold, *italic);
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
                    let char_count = text.chars().count();
                    let glyphs = font.text_to_glyphs_vec(&**text);
                    // Batch path: use text_to_glyphs + get_widths when glyph
                    // count matches char count (common Latin/CJK text).
                    // Fallback to per-char measure_str for ligatures or
                    // complex scripts where counts diverge.
                    let batch_widths = if glyphs.len() == char_count {
                        let mut widths = vec![0f32; glyphs.len()];
                        font.get_widths(&glyphs, &mut widths);
                        Some(widths)
                    } else {
                        None
                    };

                    let mut cursor = *position;
                    let mut buf = [0u8; 4];
                    for (i, ch) in text.chars().enumerate() {
                        let s = ch.encode_utf8(&mut buf);
                        let w = if let Some(ref widths) = batch_widths {
                            widths[i]
                        } else {
                            font.measure_str(&*s, None).0
                        };
                        canvas.draw_str(&*s, to_point(cursor), font, &paint);
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
                let ptr_key: *const [u8] = Rc::as_ptr(&image_data.data);
                if let Some(image) = image_cache.get(&ptr_key) {
                    canvas.draw_image_rect(image, None, to_rect(*rect), &Paint::default());
                } else {
                    let decoded = decode_image(image_data);
                    if let Some(image) = decoded {
                        let image = downsample_if_oversize(image, *rect);
                        canvas.draw_image_rect(&image, None, to_rect(*rect), &Paint::default());
                        image_cache.insert(ptr_key, image);
                    } else {
                        let magic = &image_data.data[..image_data.data.len().min(4)];
                        log::warn!(
                            "[paint] unsupported image format {:?} — could not decode {} bytes \
                             (magic: {:02x?}); image will be blank",
                            image_data.format,
                            image_data.data.len(),
                            magic,
                        );
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

/// Decode a `MediaEntry` to a Skia image, dispatching on format.
///
/// Returns `None` if the format is unsupported or the data is malformed.
fn decode_image(entry: &MediaEntry) -> Option<skia_safe::Image> {
    use crate::model::ImageFormat;
    match entry.format {
        ImageFormat::Emf => crate::render::emf::decode_emf_bitmap(&entry.data),
        // All other formats are handled by Skia's built-in decoder.
        _ => skia_safe::Image::from_encoded(Data::new_copy(&entry.data)),
    }
}

/// Downsample an image if its native pixel dimensions significantly exceed
/// the display dimensions at `IMAGE_TARGET_DPI`. Uses Mitchell-Netravali
/// cubic filtering for high-quality results.
fn downsample_if_oversize(
    image: skia_safe::Image,
    rect: crate::render::geometry::PtRect,
) -> skia_safe::Image {
    use skia_safe::CubicResampler;
    use skia_safe::{AlphaType, ColorType, ImageInfo, SamplingOptions};

    let target_w = (rect.size.width.raw() * IMAGE_DPI_SCALE).ceil() as i32;
    let target_h = (rect.size.height.raw() * IMAGE_DPI_SCALE).ceil() as i32;
    if image.width() > target_w && image.height() > target_h && target_w > 0 && target_h > 0 {
        log::debug!(
            "[paint] downsampling image {}×{} → {}×{} (display {:.0}×{:.0}pt @ {:.0} DPI)",
            image.width(),
            image.height(),
            target_w,
            target_h,
            rect.size.width.raw(),
            rect.size.height.raw(),
            IMAGE_TARGET_DPI,
        );
        // Draw scaled image onto an opaque surface so Skia applies JPEG
        // encoding (encoding_quality) instead of lossless FlateDecode.
        let info = ImageInfo::new(
            (target_w, target_h),
            ColorType::RGBA8888,
            AlphaType::Opaque,
            None,
        );
        let sampling = SamplingOptions::from(CubicResampler::mitchell());
        if let Some(mut surface) = skia_safe::surfaces::raster(&info, None, None) {
            let dst = skia_safe::Rect::from_iwh(target_w, target_h);
            surface.canvas().draw_image_rect_with_sampling_options(
                &image,
                None,
                dst,
                sampling,
                &Paint::default(),
            );
            surface.image_snapshot()
        } else {
            image
        }
    } else {
        image
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
                text: Rc::from(""),
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
