//! Paint phase — iterate DrawCommands and emit Skia PDF canvas operations.

use skia_safe::{pdf, Data, FontMgr, Paint};

use crate::dimension::Pt;
use crate::error::RenderError;
use crate::fonts;
use crate::layout::draw_command::{DrawCommand, LayoutedPage};
use crate::skia_conv::{to_color4f, to_line, to_point, to_rect, to_size};

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

                if char_spacing.abs() > Pt::ZERO {
                    let mut cursor = *position;
                    for ch in text.chars() {
                        let s = ch.to_string();
                        canvas.draw_str(&s, to_point(cursor), &font, &paint);
                        let (w, _) = font.measure_str(&s, None);
                        cursor.x += Pt::new(w) + *char_spacing;
                    }
                } else {
                    canvas.draw_str(text, to_point(*position), &font, &paint);
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
