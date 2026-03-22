use skia_safe::{pdf, Color4f, Data, FontMgr, Paint, Rect};

use super::fonts;
use super::layout::{DrawCommand, LayoutedPage};
use crate::error::Error;

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
        let mut on_page = doc.begin_page(
            (f32::from(page.page_width), f32::from(page.page_height)),
            None,
        );
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
                x,
                y,
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
                    f32::from(*x),
                    f32::from(*y),
                    text,
                    font_family,
                    f32::from(*font_size),
                    *bold,
                    *italic,
                    *color,
                    f32::from(*char_spacing_pt),
                );
            }
            DrawCommand::Underline {
                x1,
                y1,
                x2,
                y2,
                color,
                width,
            }
            | DrawCommand::Line {
                x1,
                y1,
                x2,
                y2,
                color,
                width,
            } => {
                draw_line(
                    canvas,
                    f32::from(*x1),
                    f32::from(*y1),
                    f32::from(*x2),
                    f32::from(*y2),
                    *color,
                    f32::from(*width),
                );
            }
            DrawCommand::Image {
                x,
                y,
                width,
                height,
                image,
            } => {
                let rect = Rect::from_xywh(
                    f32::from(*x),
                    f32::from(*y),
                    f32::from(*width),
                    f32::from(*height),
                );
                canvas.draw_image_rect(image, None, rect, &Paint::default());
            }
            DrawCommand::Rect {
                x,
                y,
                width,
                height,
                color,
            } => {
                draw_rect(
                    canvas,
                    f32::from(*x),
                    f32::from(*y),
                    f32::from(*width),
                    f32::from(*height),
                    *color,
                );
            }
            DrawCommand::LinkAnnotation {
                x,
                y,
                width,
                height,
                url,
            } => {
                let rect = Rect::from_xywh(
                    f32::from(*x),
                    f32::from(*y),
                    f32::from(*width),
                    f32::from(*height),
                );
                let mut url_bytes = url.as_bytes().to_vec();
                url_bytes.push(0);
                let url_data = Data::new_copy(&url_bytes);
                canvas.annotate_rect_with_url(rect, &url_data);
            }
        }
    }

    Ok(())
}

fn draw_text(
    canvas: &skia_safe::Canvas,
    font_mgr: &FontMgr,
    x: f32,
    y: f32,
    text: &str,
    font_family: &str,
    font_size: f32,
    bold: bool,
    italic: bool,
    color: (u8, u8, u8),
    char_spacing_pt: f32,
) {
    let font = fonts::make_font(font_mgr, font_family, font_size, bold, italic);
    if char_spacing_pt.abs() > f32::EPSILON {
        let mut paint = Paint::default();
        paint.set_anti_alias(true);
        paint.set_color4f(color_to_4f(color), None);
        let mut cx = x;
        for ch in text.chars() {
            let s = ch.to_string();
            canvas.draw_str(&s, (cx, y), &font, &paint);
            let (w, _) = font.measure_str(&s, None);
            cx += w + char_spacing_pt;
        }
        return;
    }
    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_color4f(color_to_4f(color), None);

    canvas.draw_str(text, (x, y), &font, &paint);
}

fn draw_line(
    canvas: &skia_safe::Canvas,
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
    color: (u8, u8, u8),
    width: f32,
) {
    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_stroke(true);
    paint.set_stroke_width(width);
    paint.set_color4f(color_to_4f(color), None);

    canvas.draw_line((x1, y1), (x2, y2), &paint);
}

fn draw_rect(
    canvas: &skia_safe::Canvas,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    color: (u8, u8, u8),
) {
    let rect = Rect::from_xywh(x, y, width, height);
    let mut paint = Paint::default();
    paint.set_anti_alias(false);
    paint.set_color4f(color_to_4f(color), None);
    canvas.draw_rect(rect, &paint);
}

fn color_to_4f(color: (u8, u8, u8)) -> Color4f {
    const MAX_U8: f32 = u8::MAX as f32;
    Color4f::new(
        color.0 as f32 / MAX_U8,
        color.1 as f32 / MAX_U8,
        color.2 as f32 / MAX_U8,
        1.0,
    )
}
