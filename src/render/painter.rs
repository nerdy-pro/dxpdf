use skia_safe::{pdf, Color4f, Data, Font, FontMgr, FontStyle, Image, Paint, Rect};

use super::layout::{DrawCommand, LayoutConfig, LayoutedPage};
use crate::error::Error;

/// Render laid-out pages to a PDF byte buffer.
pub fn render_to_pdf(pages: &[LayoutedPage], config: &LayoutConfig) -> Result<Vec<u8>, Error> {
    let font_mgr = FontMgr::new();
    let mut pdf_bytes: Vec<u8> = Vec::new();
    let mut doc = pdf::new_document(&mut pdf_bytes, None);

    for page in pages {
        let mut on_page = doc.begin_page((config.page_width, config.page_height), None);
        {
            let canvas = on_page.canvas();
            render_page(canvas, page, &font_mgr)?;
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
            DrawCommand::Text { x, y, text, font_family, font_size, bold, italic, color } => {
                draw_text(canvas, font_mgr, *x, *y, text, font_family, *font_size, *bold, *italic, *color);
            }
            DrawCommand::Underline { x1, y1, x2, y2, color, width }
            | DrawCommand::Line { x1, y1, x2, y2, color, width } => {
                draw_line(canvas, *x1, *y1, *x2, *y2, *color, *width);
            }
            DrawCommand::Image { x, y, width, height, data } => {
                draw_image(canvas, *x, *y, *width, *height, data);
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
) {
    let style = match (bold, italic) {
        (true, true) => FontStyle::bold_italic(),
        (true, false) => FontStyle::bold(),
        (false, true) => FontStyle::italic(),
        (false, false) => FontStyle::normal(),
    };

    let typeface = font_mgr
        .match_family_style(font_family, style)
        .or_else(|| font_mgr.match_family_style("Helvetica", style))
        .or_else(|| font_mgr.legacy_make_typeface(None::<&str>, style))
        .expect("no fallback typeface available");

    let font = Font::from_typeface(typeface, font_size);
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

fn draw_image(
    canvas: &skia_safe::Canvas,
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    data: &[u8],
) {
    let skia_data = Data::new_copy(data);
    if let Some(image) = Image::from_encoded(skia_data) {
        let rect = Rect::from_xywh(x, y, width, height);
        canvas.draw_image_rect(image, None, rect, &Paint::default());
    }
}

fn color_to_4f(color: (u8, u8, u8)) -> Color4f {
    Color4f::new(
        color.0 as f32 / 255.0,
        color.1 as f32 / 255.0,
        color.2 as f32 / 255.0,
        1.0,
    )
}
