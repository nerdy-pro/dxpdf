//! Header/footer layout — render headers and footers on each page.
//!
//! Headers and footers are laid out in a separate constraint frame
//! (between page edge and body margin), then their draw commands are
//! prepended to each page's command list.
//!
//! Content is built per-page so that PAGE / NUMPAGES fields (§17.16.4.1)
//! evaluate to the correct values on each page.

use dxpdf_docx_model::model::Block;

use crate::dimension::Pt;

use super::build::{BuildContext, HeaderFooterContent, build_header_footer_content};
use super::draw_command::{DrawCommand, LayoutedPage};
use super::page::PageConfig;
use super::section::stack_blocks;

/// Render headers and footers onto already-laid-out pages.
///
/// `header_blocks` / `footer_blocks` are the raw DOCX blocks; content is
/// rebuilt per-page so that field values (PAGE, NUMPAGES) are correct.
/// `page_base` is the 0-based index of the first page in `pages` within
/// the overall document (for multi-section PAGE numbering).
#[allow(clippy::too_many_arguments)]
pub fn render_headers_footers(
    pages: &mut [LayoutedPage],
    config: &PageConfig,
    header_blocks: Option<&[Block]>,
    footer_blocks: Option<&[Block]>,
    ctx: &BuildContext,
    default_line_height: Pt,
    page_base: usize,
    total_pages: usize,
) {
    let content_width = config.content_width();

    for (page_idx, page) in pages.iter_mut().enumerate() {
        let page_number = page_base + page_idx + 1; // 1-based

        // Header
        if let Some(blocks) = header_blocks {
            // Set per-page field context for PAGE/NUMPAGES evaluation.
            ctx.field_ctx_cell.set(crate::layout::fragment::FieldContext {
                page_number: Some(page_number),
                num_pages: Some(total_pages),
            });

            let hf = build_header_footer_content(blocks, ctx);
            render_header(page, config, &hf, content_width, default_line_height);
        }

        // Footer
        if let Some(blocks) = footer_blocks {
            ctx.field_ctx_cell.set(crate::layout::fragment::FieldContext {
                page_number: Some(page_number),
                num_pages: Some(total_pages),
            });

            let hf = build_header_footer_content(blocks, ctx);
            render_footer(page, config, &hf, content_width, default_line_height);
        }
    }

    // Reset field context after header/footer rendering.
    ctx.field_ctx_cell.set(crate::layout::fragment::FieldContext::default());
}

/// Render a single header onto a page.
fn render_header(
    page: &mut LayoutedPage,
    config: &PageConfig,
    hf: &HeaderFooterContent,
    content_width: Pt,
    default_line_height: Pt,
) {
    if hf.blocks.is_empty() {
        return;
    }

    let (offset_x, offset_y) = if let Some((abs_x, abs_y)) = hf.absolute_position {
        (abs_x, abs_y)
    } else {
        (config.margins.left, config.header_margin)
    };

    let result = stack_blocks(&hf.blocks, content_width, default_line_height, None);

    let mut header_cmds: Vec<DrawCommand> = result
        .commands
        .into_iter()
        .map(|mut cmd| {
            cmd.shift(offset_x, offset_y);
            cmd
        })
        .collect();

    // Render floating images from the header (page-relative positions).
    for fi in &hf.floating_images {
        let img_y = match fi.y {
            super::section::FloatingImageY::Absolute(y) => y,
            super::section::FloatingImageY::RelativeToParagraph(offset) => offset_y + offset,
        };
        header_cmds.push(DrawCommand::Image {
            rect: crate::geometry::PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
            image_data: fi.image_data.clone(),
        });
    }

    // Prepend header commands before body content.
    header_cmds.append(&mut page.commands);
    page.commands = header_cmds;
}

/// Render a single footer onto a page.
fn render_footer(
    page: &mut LayoutedPage,
    config: &PageConfig,
    hf: &HeaderFooterContent,
    content_width: Pt,
    default_line_height: Pt,
) {
    if hf.blocks.is_empty() {
        return;
    }

    let result = stack_blocks(&hf.blocks, content_width, default_line_height, None);

    let footer_y = config.page_size.height - config.footer_margin - result.height;
    for mut cmd in result.commands {
        cmd.shift(config.margins.left, footer_y);
        page.commands.push(cmd);
    }

    // Render floating images from the footer.
    for fi in &hf.floating_images {
        let img_y = match fi.y {
            super::section::FloatingImageY::Absolute(y) => y,
            super::section::FloatingImageY::RelativeToParagraph(offset) => footer_y + offset,
        };
        page.commands.push(DrawCommand::Image {
            rect: crate::geometry::PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
            image_data: fi.image_data.clone(),
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{PtEdgeInsets, PtOffset, PtSize};
    use crate::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::layout::paragraph::ParagraphStyle;
    use crate::layout::section::LayoutBlock;
    use crate::resolve::color::RgbColor;
    use std::rc::Rc;

    fn make_hf(frags: Vec<Fragment>) -> HeaderFooterContent {
        HeaderFooterContent {
            blocks: vec![LayoutBlock::Paragraph {
                fragments: frags,
                style: ParagraphStyle::default(),
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
            }],
            absolute_position: None,
            floating_images: vec![],
        }
    }

    fn text_frag(s: &str) -> Fragment {
        let font = FontProps {
            family: Rc::from("Test"),
            size: Pt::new(12.0),
            bold: false,
            italic: false,
            underline: false,
            char_spacing: Pt::ZERO,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        };
        Fragment::Text {
            text: s.to_string(),
            font,
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            width: Pt::new(40.0),
            trimmed_width: Pt::new(40.0),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
            },
            hyperlink_url: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
        }
    }

    fn test_config() -> PageConfig {
        use crate::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
            margins: PtEdgeInsets::new(Pt::new(72.0), Pt::new(72.0), Pt::new(72.0), Pt::new(72.0)),
            header_margin: Pt::new(36.0),
            footer_margin: Pt::new(36.0),
            columns: vec![ColumnGeometry { x_offset: Pt::ZERO, width: Pt::new(468.0) }],
        }
    }

    #[test]
    fn no_header_footer_leaves_page_unchanged() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];
        pages[0].commands.push(DrawCommand::Text {
            text: "body".into(),
            position: PtOffset::new(Pt::ZERO, Pt::ZERO),
            font_family: Rc::from("T"),
            font_size: Pt::new(12.0),
            char_spacing: Pt::ZERO,
            bold: false, italic: false,
            color: RgbColor::BLACK,
        });

        let config = test_config();
        // Direct call to render_header / render_footer with empty content.
        let hf = HeaderFooterContent { blocks: vec![], absolute_position: None, floating_images: vec![] };
        render_header(&mut pages[0], &config, &hf, config.content_width(), Pt::new(14.0));
        render_footer(&mut pages[0], &config, &hf, config.content_width(), Pt::new(14.0));

        assert_eq!(pages[0].commands.len(), 1, "no changes");
    }

    #[test]
    fn header_prepended_to_page() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];
        pages[0].commands.push(DrawCommand::Text {
            text: "body".into(),
            position: PtOffset::new(Pt::ZERO, Pt::ZERO),
            font_family: Rc::from("T"),
            font_size: Pt::new(12.0),
            char_spacing: Pt::ZERO,
            bold: false, italic: false,
            color: RgbColor::BLACK,
        });

        let config = test_config();
        let header = make_hf(vec![text_frag("Header")]);
        render_header(&mut pages[0], &config, &header, config.content_width(), Pt::new(14.0));

        assert!(pages[0].commands.len() > 1);
        // First command should be the header text
        if let DrawCommand::Text { text, .. } = &pages[0].commands[0] {
            assert_eq!(text, "Header");
        }
    }

    #[test]
    fn footer_appended_to_page() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];

        let config = test_config();
        let footer = make_hf(vec![text_frag("Footer")]);
        render_footer(&mut pages[0], &config, &footer, config.content_width(), Pt::new(14.0));

        assert_eq!(pages[0].commands.len(), 1);
        if let DrawCommand::Text { text, position, .. } = &pages[0].commands[0] {
            assert_eq!(text, "Footer");
            // Footer y should be near the bottom of the page.
            assert!(position.y.raw() > 700.0, "footer y={}", position.y.raw());
        }
    }

    #[test]
    fn header_applied_to_all_pages() {
        let mut pages = vec![
            LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0))),
            LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0))),
        ];

        let config = test_config();
        let header = make_hf(vec![text_frag("H")]);
        for page in pages.iter_mut() {
            render_header(page, &config, &header, config.content_width(), Pt::new(14.0));
        }

        // Both pages should have header
        for page in &pages {
            assert!(!page.commands.is_empty());
        }
    }

    #[test]
    fn header_y_position_uses_header_margin() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];
        let config = test_config();
        let header = make_hf(vec![text_frag("H")]);
        render_header(&mut pages[0], &config, &header, config.content_width(), Pt::new(14.0));

        if let DrawCommand::Text { position, .. } = &pages[0].commands[0] {
            // Header y should be near header_margin (36) + ascent
            assert!(
                position.y.raw() > 36.0 && position.y.raw() < 72.0,
                "header y={} should be between header_margin and top margin",
                position.y.raw()
            );
        }
    }
}
