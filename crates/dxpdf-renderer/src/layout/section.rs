//! Section layout — sequence blocks vertically into pages.
//!
//! Takes measured blocks (paragraphs with fragments, tables with cells),
//! fits them into pages respecting page size and margins, handles page breaks.

use crate::dimension::Pt;
use super::draw_command::{DrawCommand, LayoutedPage};
use super::fragment::Fragment;
use super::page::PageConfig;
use super::paragraph::{layout_paragraph, ParagraphStyle};
use super::table::{layout_table, TableRowInput};
use super::BoxConstraints;

/// A block ready for layout — either a paragraph or a table.
pub enum LayoutBlock {
    Paragraph {
        fragments: Vec<Fragment>,
        style: ParagraphStyle,
    },
    Table {
        rows: Vec<TableRowInput>,
        col_widths: Vec<Pt>,
        draw_borders: bool,
    },
}

/// Lay out a sequence of blocks into pages.
pub fn layout_section(
    blocks: &[LayoutBlock],
    config: &PageConfig,
    default_line_height: Pt,
) -> Vec<LayoutedPage> {
    let content_width = config.content_width();
    let content_height = config.content_height();
    let constraints = BoxConstraints::new(
        Pt::ZERO,
        content_width,
        Pt::ZERO,
        content_height,
    );

    let mut pages: Vec<LayoutedPage> = Vec::new();
    let mut current_page = LayoutedPage::new(config.page_size);
    let mut cursor_y = config.margins.top;
    let bottom = config.page_size.height - config.margins.bottom;

    for block in blocks {
        match block {
            LayoutBlock::Paragraph { fragments, style } => {
                let para = layout_paragraph(
                    fragments,
                    &constraints,
                    style,
                    default_line_height,
                );

                // Page break if paragraph doesn't fit
                if cursor_y + para.size.height > bottom && cursor_y > config.margins.top {
                    pages.push(std::mem::replace(
                        &mut current_page,
                        LayoutedPage::new(config.page_size),
                    ));
                    cursor_y = config.margins.top;
                }

                // Offset commands to absolute page position
                for mut cmd in para.commands {
                    cmd.shift_y(cursor_y);
                    shift_x(&mut cmd, config.margins.left);
                    current_page.commands.push(cmd);
                }

                cursor_y += para.size.height;
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                draw_borders,
            } => {
                let table = layout_table(
                    rows,
                    col_widths,
                    &constraints,
                    default_line_height,
                    *draw_borders,
                );

                // Page break if table doesn't fit
                if cursor_y + table.size.height > bottom && cursor_y > config.margins.top {
                    pages.push(std::mem::replace(
                        &mut current_page,
                        LayoutedPage::new(config.page_size),
                    ));
                    cursor_y = config.margins.top;
                }

                for mut cmd in table.commands {
                    cmd.shift_y(cursor_y);
                    shift_x(&mut cmd, config.margins.left);
                    current_page.commands.push(cmd);
                }

                cursor_y += table.size.height;
            }
        }
    }

    // Push the last page (even if empty — ensure at least one page)
    pages.push(current_page);

    pages
}

fn shift_x(cmd: &mut DrawCommand, dx: Pt) {
    match cmd {
        DrawCommand::Text { position, .. } => position.x += dx,
        DrawCommand::Underline { line, .. } => {
            line.start.x += dx;
            line.end.x += dx;
        }
        DrawCommand::Image { rect, .. } => rect.origin.x += dx,
        DrawCommand::Rect { rect, .. } => rect.origin.x += dx,
        DrawCommand::LinkAnnotation { rect, .. } => rect.origin.x += dx,
        DrawCommand::Line { line, .. } => {
            line.start.x += dx;
            line.end.x += dx;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{PtEdgeInsets, PtSize};
    use crate::layout::cell::CellBlock;
    use crate::layout::fragment::FontProps;
    use crate::layout::table::TableCellInput;
    use crate::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str, width: f32, height: f32) -> Fragment {
        Fragment::Text {
            text: text.into(),
            font: FontProps {
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: false,
                italic: false,
                underline: false,
                char_spacing: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width),
            height: Pt::new(height),
            ascent: Pt::new(height * 0.7),
            hyperlink_url: None,
            baseline_offset: Pt::ZERO,
        }
    }

    fn para_block(text: &str, width: f32) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, width, 14.0)],
            style: ParagraphStyle::default(),
        }
    }

    fn small_config() -> PageConfig {
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
        }
    }

    #[test]
    fn empty_blocks_produces_one_empty_page() {
        let pages = layout_section(&[], &small_config(), Pt::new(14.0));
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn single_paragraph_on_one_page() {
        let blocks = vec![para_block("hello", 30.0)];
        let pages = layout_section(&blocks, &small_config(), Pt::new(14.0));

        assert_eq!(pages.len(), 1);
        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    #[test]
    fn text_positioned_at_margins() {
        let blocks = vec![para_block("hello", 30.0)];
        let config = small_config();
        let pages = layout_section(&blocks, &config, Pt::new(14.0));

        if let Some(DrawCommand::Text { position, .. }) = pages[0].commands.first() {
            assert!(
                position.x.raw() >= config.margins.left.raw(),
                "x should be at least left margin"
            );
            assert!(
                position.y.raw() >= config.margins.top.raw(),
                "y should be at least top margin"
            );
        }
    }

    #[test]
    fn page_break_when_content_overflows() {
        // Page: 100pt tall, margins 10 each → 80pt content area
        // Each paragraph: 14pt tall
        // 6 paragraphs = 84pt > 80pt → should break to 2 pages
        let blocks: Vec<_> = (0..6).map(|i| para_block(&format!("p{i}"), 30.0)).collect();
        let pages = layout_section(&blocks, &small_config(), Pt::new(14.0));

        assert_eq!(pages.len(), 2, "should overflow to 2 pages");

        let page1_texts: Vec<_> = pages[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect();
        let page2_texts: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect();

        assert_eq!(page1_texts.len(), 5, "5 paras fit on page 1 (5*14=70 < 80)");
        assert_eq!(page2_texts.len(), 1, "1 para on page 2");
    }

    #[test]
    fn page_size_set_on_layouted_page() {
        let config = small_config();
        let pages = layout_section(&[], &config, Pt::new(14.0));
        assert_eq!(pages[0].page_size, config.page_size);
    }

    #[test]
    fn many_paragraphs_produce_multiple_pages() {
        // 20 paragraphs at 14pt each = 280pt
        // Content area = 80pt → need 4 pages (80/14 = 5.7 paras per page)
        let blocks: Vec<_> = (0..20).map(|i| para_block(&format!("p{i}"), 30.0)).collect();
        let pages = layout_section(&blocks, &small_config(), Pt::new(14.0));

        assert_eq!(pages.len(), 4);
    }

    #[test]
    fn table_on_page() {
        let blocks = vec![LayoutBlock::Table {
            rows: vec![TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![CellBlock {
                        fragments: vec![text_frag("cell", 30.0, 14.0)],
                        style: ParagraphStyle::default(),
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                }],
                min_height: None,
            }],
            col_widths: vec![Pt::new(100.0)],
            draw_borders: false,
        }];

        let pages = layout_section(&blocks, &small_config(), Pt::new(14.0));
        assert_eq!(pages.len(), 1);

        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }
}
