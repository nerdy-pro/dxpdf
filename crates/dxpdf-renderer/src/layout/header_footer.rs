//! Header/footer layout — render headers and footers on each page.
//!
//! Headers and footers are laid out in a separate constraint frame
//! (between page edge and body margin), then their draw commands are
//! prepended to each page's command list.

use crate::dimension::Pt;

use super::draw_command::{DrawCommand, LayoutedPage};
use super::fragment::Fragment;
use super::page::PageConfig;
use super::paragraph::{layout_paragraph, ParagraphStyle};
use super::BoxConstraints;

/// Render headers and footers onto already-laid-out pages.
///
/// `header_fragments` / `footer_fragments`: pre-collected fragments for the
/// header/footer content. `page_number_field` is substituted into PAGE fields.
pub fn render_headers_footers(
    pages: &mut [LayoutedPage],
    config: &PageConfig,
    header_fragments: Option<&[Fragment]>,
    footer_fragments: Option<&[Fragment]>,
    default_line_height: Pt,
) {
    let content_width = config.content_width();

    for page in pages.iter_mut() {

        // Header
        if let Some(frags) = header_fragments {
            let constraints = BoxConstraints::tight_width(content_width, config.margins.top);
            let para = layout_paragraph(frags, &constraints, &ParagraphStyle::default(), default_line_height, None);

            let header_y = config.header_margin;
            let mut header_cmds: Vec<DrawCommand> = para
                .commands
                .into_iter()
                .map(|mut cmd| {
                    cmd.shift(config.margins.left, header_y);
                    cmd
                })
                .collect();

            // Prepend header commands before body content
            header_cmds.append(&mut page.commands);
            page.commands = header_cmds;
        }

        // Footer
        if let Some(frags) = footer_fragments {
            let constraints = BoxConstraints::tight_width(content_width, config.margins.bottom);
            let para = layout_paragraph(frags, &constraints, &ParagraphStyle::default(), default_line_height, None);

            let footer_y = config.page_size.height - config.footer_margin - para.size.height;
            for mut cmd in para.commands {
                cmd.shift(config.margins.left, footer_y);
                page.commands.push(cmd);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{PtEdgeInsets, PtOffset, PtSize};
    use crate::layout::fragment::FontProps;
    use crate::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str) -> Fragment {
        Fragment::Text {
            text: text.into(),
            font: FontProps {
                family: Rc::from("Test"),
                size: Pt::new(10.0),
                bold: false,
                italic: false,
                underline: false,
                char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(40.0), trimmed_width: Pt::new(40.0),
            height: Pt::new(12.0),
            ascent: Pt::new(9.0),
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO, text_offset: Pt::ZERO,
        }
    }

    fn test_config() -> PageConfig {
        PageConfig {
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
            margins: PtEdgeInsets::new(Pt::new(72.0), Pt::new(72.0), Pt::new(72.0), Pt::new(72.0)),
            header_margin: Pt::new(36.0),
            footer_margin: Pt::new(36.0),
        }
    }

    #[test]
    fn no_header_footer_leaves_pages_unchanged() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];
        pages[0].commands.push(DrawCommand::Text {
            position: PtOffset::new(Pt::new(72.0), Pt::new(82.0)),
            text: "body".into(),
            font_family: Rc::from("Test"),
            char_spacing: Pt::ZERO,
            font_size: Pt::new(12.0),
            bold: false,
            italic: false,
            color: RgbColor::BLACK,
        });

        let config = test_config();
        render_headers_footers(&mut pages, &config, None, None, Pt::new(14.0));

        assert_eq!(pages[0].commands.len(), 1, "no changes");
    }

    #[test]
    fn header_prepended_to_page() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];
        pages[0].commands.push(DrawCommand::Text {
            position: PtOffset::new(Pt::new(72.0), Pt::new(82.0)),
            text: "body".into(),
            font_family: Rc::from("Test"),
            char_spacing: Pt::ZERO,
            font_size: Pt::new(12.0),
            bold: false,
            italic: false,
            color: RgbColor::BLACK,
        });

        let config = test_config();
        let header_frags = vec![text_frag("Header")];
        render_headers_footers(&mut pages, &config, Some(&header_frags), None, Pt::new(14.0));

        assert!(pages[0].commands.len() > 1);
        // First command should be the header text
        if let DrawCommand::Text { text, position, .. } = &pages[0].commands[0] {
            assert_eq!(text, "Header");
            assert!(position.y.raw() < 72.0, "header above body margin");
        }
    }

    #[test]
    fn footer_appended_to_page() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];

        let config = test_config();
        let footer_frags = vec![text_frag("Footer")];
        render_headers_footers(&mut pages, &config, None, Some(&footer_frags), Pt::new(14.0));

        assert_eq!(pages[0].commands.len(), 1);
        if let DrawCommand::Text { text, position, .. } = &pages[0].commands[0] {
            assert_eq!(text, "Footer");
            assert!(position.y.raw() > 700.0, "footer near bottom of page");
        }
    }

    #[test]
    fn header_footer_on_multiple_pages() {
        let mut pages = vec![
            LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0))),
            LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0))),
        ];

        let config = test_config();
        let header_frags = vec![text_frag("H")];
        render_headers_footers(&mut pages, &config, Some(&header_frags), None, Pt::new(14.0));

        // Both pages should have header
        for page in &pages {
            let has_header = page.commands.iter().any(|c| matches!(c, DrawCommand::Text { text, .. } if text == "H"));
            assert!(has_header, "each page gets a header");
        }
    }

    #[test]
    fn header_positioned_at_header_margin() {
        let mut pages = vec![LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)))];
        let config = test_config();
        let header_frags = vec![text_frag("H")];
        render_headers_footers(&mut pages, &config, Some(&header_frags), None, Pt::new(14.0));

        if let DrawCommand::Text { position, .. } = &pages[0].commands[0] {
            // Header y should be near header_margin (36) + ascent
            assert!(position.y.raw() >= 36.0);
            assert!(position.y.raw() < 72.0, "header within margin area");
            assert_eq!(position.x.raw(), 72.0, "at left margin");
        }
    }
}
