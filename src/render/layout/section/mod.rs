//! Section layout — sequence blocks vertically into pages.
//!
//! Takes measured blocks (paragraphs with fragments, tables with cells),
//! fits them into pages respecting page size and margins, handles page breaks.

mod helpers;
mod layout;
mod stacker;
mod types;

pub use layout::layout_section;
pub use stacker::{stack_blocks, StackResult};
pub use types::*;

// ── Footnote rendering constants ─────────────────────────────────────────────

/// §17.11.23: footnote separator width as a fraction of the content area.
/// Word renders the separator at one-third of the text column width.
const FOOTNOTE_SEPARATOR_RATIO: f32 = 0.33;

/// Thickness of the footnote separator line (pt).
const FOOTNOTE_SEPARATOR_LINE_WIDTH: crate::render::dimension::Pt =
    crate::render::dimension::Pt::new(0.5);

/// Vertical gap between the footnote separator and the first footnote paragraph.
/// Also used as the initial height budget for the separator region (pt).
const FOOTNOTE_SEPARATOR_GAP: crate::render::dimension::Pt = crate::render::dimension::Pt::new(4.0);

// ── Float deduplication ───────────────────────────────────────────────────────

/// Position tolerance for deduplicating floating images (pt).
/// Two float entries within this distance on every axis are treated as identical.
const FLOAT_DEDUP_EPSILON_PT: f32 = 0.1;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::dimension::Pt;
    use crate::render::geometry::{PtEdgeInsets, PtSize};
    use crate::render::layout::draw_command::DrawCommand;
    use crate::render::layout::fragment::Fragment;
    use crate::render::layout::fragment::{FontProps, TextMetrics};
    use crate::render::layout::page::PageConfig;
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::table::{TableCellInput, TableRowInput};
    use crate::render::resolve::color::RgbColor;
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
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(height * 0.7),
                descent: Pt::new(height * 0.3),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
        }
    }

    fn para_block(text: &str, width: f32) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, width, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
        }
    }

    fn small_config() -> PageConfig {
        use crate::render::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
            columns: vec![ColumnGeometry {
                x_offset: Pt::ZERO,
                width: Pt::new(180.0),
            }],
        }
    }

    #[test]
    fn empty_blocks_produces_one_empty_page() {
        let pages = layout_section(&[], &small_config(), None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn single_paragraph_on_one_page() {
        let blocks = vec![para_block("hello", 30.0)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

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
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

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
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

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
        let pages = layout_section(&[], &config, None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages[0].page_size, config.page_size);
    }

    #[test]
    fn many_paragraphs_produce_multiple_pages() {
        // 20 paragraphs at 14pt each = 280pt
        // Content area = 80pt → need 4 pages (80/14 = 5.7 paras per page)
        let blocks: Vec<_> = (0..20)
            .map(|i| para_block(&format!("p{i}"), 30.0))
            .collect();
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 4);
    }

    #[test]
    fn table_on_page() {
        let blocks = vec![LayoutBlock::Table {
            rows: vec![TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: vec![text_frag("cell", 30.0, 14.0)],
                        style: ParagraphStyle::default(),
                        page_break_before: false,
                        footnotes: vec![],
                        floating_images: vec![],
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: None,
                    vertical_align: crate::render::layout::table::CellVAlign::Top,
                }],
                height_rule: None,
                is_header: None,
                cant_split: None,
            }],
            col_widths: vec![Pt::new(100.0)],
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        }];

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        assert_eq!(pages.len(), 1);

        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    // ── §17.3.1.33 space_before suppression tests ──────────────────────

    #[test]
    fn space_before_suppressed_for_first_paragraph_of_section() {
        let style = ParagraphStyle {
            space_before: Pt::new(24.0),
            ..Default::default()
        };
        let blocks = vec![LayoutBlock::Paragraph {
            fragments: vec![text_frag("heading", 50.0, 14.0)],
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
        }];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        // First paragraph on the section's initial page: space_before suppressed.
        if let Some(DrawCommand::Text { position, .. }) = pages[0].commands.first() {
            assert!(
                position.y.raw() < config.margins.top.raw() + 24.0,
                "space_before should be suppressed: y={}",
                position.y.raw()
            );
        }
    }

    #[test]
    fn space_before_preserved_for_page_break_before() {
        let heading_style = ParagraphStyle {
            space_before: Pt::new(24.0),
            ..Default::default()
        };

        let blocks = vec![
            para_block("first page", 30.0),
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("heading", 50.0, 14.0)],
                style: heading_style,
                page_break_before: true,
                footnotes: vec![],
                floating_images: vec![],
            },
        ];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        assert!(pages.len() >= 2, "should have at least 2 pages");
        let heading_y = pages[1]
            .commands
            .iter()
            .find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if &**text == "heading" => {
                    Some(position.y)
                }
                _ => None,
            })
            .expect("heading should be on page 2");
        // §17.3.1.33: space_before is preserved — pageBreakBefore paragraphs
        // are not the structural first of the section.
        assert!(
            heading_y.raw() > config.margins.top.raw() + 20.0,
            "space_before should be preserved for pageBreakBefore: y={}",
            heading_y.raw(),
        );
    }

    // ── §17.3.1.24 paragraph border grouping tests ─────────────────────

    #[test]
    fn identical_borders_suppress_second_top() {
        use crate::render::layout::paragraph::{BorderLine, ParagraphBorderStyle};
        let border = Some(ParagraphBorderStyle {
            top: Some(BorderLine {
                width: Pt::new(0.5),
                color: RgbColor::BLACK,
                space: Pt::new(1.0),
            }),
            bottom: None,
            left: None,
            right: None,
        });
        let style1 = ParagraphStyle {
            borders: border.clone(),
            ..Default::default()
        };
        let style2 = ParagraphStyle {
            borders: border,
            ..Default::default()
        };

        let blocks = vec![
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("para1", 30.0, 14.0)],
                style: style1,
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
            },
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("para2", 30.0, 14.0)],
                style: style2,
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
            },
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        // Count Line draw commands (border lines).
        // Only the first paragraph should draw its top border; the second's
        // top border is suppressed by §17.3.1.24 grouping.
        let line_cmds: Vec<_> = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Line { .. }))
            .collect();
        assert_eq!(
            line_cmds.len(),
            1,
            "only one top border line (grouped): got {}",
            line_cmds.len()
        );
    }
}
