use skia_safe::{pdf, Data, FontMgr, Paint, Rect};

use super::fonts;
use super::layout::{DrawCommand, LayoutedPage};
use crate::dimension::Pt;
use crate::error::Error;
use crate::geometry::{PtLineSegment, PtOffset, PtRect};
use crate::model::Color;

/// Render laid-out pages to a PDF byte buffer.
pub fn render_to_pdf(pages: &[LayoutedPage]) -> Result<Vec<u8>, Error> {
    render_to_pdf_with_font_mgr(pages, &FontMgr::new())
}

/// Render laid-out pages to a PDF byte buffer, reusing an existing FontMgr.
pub fn render_to_pdf_with_font_mgr(
    pages: &[LayoutedPage],
    font_mgr: &FontMgr,
) -> Result<Vec<u8>, Error> {
    let mut pdf_bytes: Vec<u8> = Vec::new();
    let mut doc = pdf::new_document(&mut pdf_bytes, None);

    for page in pages {
        let mut on_page = doc.begin_page(page.page_size, None);
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
                let skia_rect: Rect = (*rect).into();
                canvas.draw_image_rect(image, None, skia_rect, &Paint::default());
            }
            DrawCommand::Rect { rect, color } => {
                draw_rect(canvas, *rect, *color);
            }
            DrawCommand::LinkAnnotation { rect, url } => {
                let skia_rect: Rect = (*rect).into();
                let mut url_bytes = url.as_bytes().to_vec();
                url_bytes.push(0);
                let url_data = Data::new_copy(&url_bytes);
                canvas.annotate_rect_with_url(skia_rect, &url_data);
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
    paint.set_color4f(skia_safe::Color4f::from(color), None);

    if char_spacing.abs() > Pt::ZERO {
        let mut cursor = position;
        for ch in text.chars() {
            let s = ch.to_string();
            let pt: skia_safe::Point = cursor.into();
            canvas.draw_str(&s, pt, &font, &paint);
            let (w, _) = font.measure_str(&s, None);
            cursor.x += Pt::new(w) + char_spacing;
        }
        return;
    }

    let pt: skia_safe::Point = position.into();
    canvas.draw_str(text, pt, &font, &paint);
}

fn draw_line(canvas: &skia_safe::Canvas, line: PtLineSegment, color: Color, width: Pt) {
    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_stroke(true);
    paint.set_stroke_width(f32::from(width));
    paint.set_color4f(skia_safe::Color4f::from(color), None);

    let start: skia_safe::Point = line.start.into();
    let end: skia_safe::Point = line.end.into();
    canvas.draw_line(start, end, &paint);
}

fn draw_rect(canvas: &skia_safe::Canvas, rect: PtRect, color: Color) {
    let mut paint = Paint::default();
    paint.set_anti_alias(false);
    paint.set_color4f(skia_safe::Color4f::from(color), None);
    let skia_rect: Rect = rect.into();
    canvas.draw_rect(skia_rect, &paint);
}
