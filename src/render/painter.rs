use skia_safe::{pdf, Data, FontMgr, Paint};

use super::fonts;
use super::layout::{DrawCommand, LayoutedPage};
use super::skia_conv::{to_color4f, to_line, to_point, to_rect, to_size};
use crate::dimension::Pt;
use crate::error::Error;
use crate::geometry::{PtLineSegment, PtOffset, PtRect};
use crate::model::Color;

/// Render laid-out pages to a PDF byte buffer.
pub fn render_to_pdf_with_font_mgr(
    pages: &[LayoutedPage],
    font_mgr: &FontMgr,
) -> Result<Vec<u8>, Error> {
    let mut pdf_bytes: Vec<u8> = Vec::new();
    let mut doc = pdf::new_document(&mut pdf_bytes, None);

    for page in pages {
        let mut on_page = doc.begin_page(to_size(page.page_size), None);
        {
            let canvas = on_page.canvas();
            render_page(canvas, page, font_mgr)?;
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
) -> Result<(), Error> {
    for cmd in &page.commands {
        match cmd {
            DrawCommand::Text {
                position,
                text,
                font_family,
                char_spacing_pt,
                font_size,
                bold,
                italic,
                color,
            } => {
                draw_text(
                    canvas,
                    font_mgr,
                    *position,
                    text,
                    font_family,
                    *font_size,
                    *bold,
                    *italic,
                    *color,
                    *char_spacing_pt,
                );
            }
            DrawCommand::Underline { line, color, width }
            | DrawCommand::Line { line, color, width } => {
                draw_line(canvas, *line, *color, *width);
            }
            DrawCommand::Image { rect, image } => {
                canvas.draw_image_rect(image, None, to_rect(*rect), &Paint::default());
            }
            DrawCommand::Rect { rect, color } => {
                draw_rect(canvas, *rect, *color);
            }
            DrawCommand::LinkAnnotation { rect, url } => {
                let mut url_bytes = url.as_bytes().to_vec();
                url_bytes.push(0);
                let url_data = Data::new_copy(&url_bytes);
                canvas.annotate_rect_with_url(to_rect(*rect), &url_data);
            }
        }
    }

    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn draw_text(
    canvas: &skia_safe::Canvas,
    font_mgr: &FontMgr,
    position: PtOffset,
    text: &str,
    font_family: &str,
    font_size: Pt,
    bold: bool,
    italic: bool,
    color: Color,
    char_spacing: Pt,
) {
    let font = fonts::make_font(font_mgr, font_family, font_size, bold, italic);
    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_color4f(to_color4f(color), None);

    if char_spacing.abs() > Pt::ZERO {
        let mut cursor = position;
        for ch in text.chars() {
            let s = ch.to_string();
            canvas.draw_str(&s, to_point(cursor), &font, &paint);
            let (w, _) = font.measure_str(&s, None);
            cursor.x += Pt::new(w) + char_spacing;
        }
        return;
    }

    canvas.draw_str(text, to_point(position), &font, &paint);
}

fn draw_line(canvas: &skia_safe::Canvas, line: PtLineSegment, color: Color, width: Pt) {
    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_stroke(true);
    paint.set_stroke_width(f32::from(width));
    paint.set_color4f(to_color4f(color), None);

    let (start, end) = to_line(line);
    canvas.draw_line(start, end, &paint);
}

fn draw_rect(canvas: &skia_safe::Canvas, rect: PtRect, color: Color) {
    let mut paint = Paint::default();
    paint.set_anti_alias(false);
    paint.set_color4f(to_color4f(color), None);
    canvas.draw_rect(to_rect(rect), &paint);
}
