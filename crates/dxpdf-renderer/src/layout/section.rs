//! Section layout — sequence blocks vertically into pages.
//!
//! Takes measured blocks (paragraphs with fragments, tables with cells),
//! fits them into pages respecting page size and margins, handles page breaks.

use std::rc::Rc;

use crate::dimension::Pt;
use crate::geometry::{PtRect, PtSize};
use super::draw_command::{DrawCommand, LayoutedPage};
use super::fragment::Fragment;
use super::page::PageConfig;
use super::paragraph::{layout_paragraph, ParagraphStyle};
use super::table::{layout_table, TableRowInput};
use super::BoxConstraints;

/// A floating (anchor) image to be positioned absolutely on the page.
#[derive(Clone)]
pub struct FloatingImage {
    pub image_data: Rc<[u8]>,
    pub size: PtSize,
    /// Resolved absolute x position on the page.
    pub x: Pt,
    /// Resolved absolute y position on the page (may be relative to paragraph).
    pub y: FloatingImageY,
}

/// Vertical position for a floating image.
#[derive(Clone, Copy)]
pub enum FloatingImageY {
    /// Absolute page position.
    Absolute(Pt),
    /// Relative to the paragraph's y position (offset added to cursor_y).
    RelativeToParagraph(Pt),
}

/// A block ready for layout — either a paragraph or a table.
pub enum LayoutBlock {
    Paragraph {
        fragments: Vec<Fragment>,
        style: ParagraphStyle,
        /// §17.3.1.23: force a page break before this paragraph.
        page_break_before: bool,
        /// Footnotes referenced in this paragraph — rendered at page bottom.
        footnotes: Vec<(Vec<Fragment>, ParagraphStyle)>,
        /// §20.4.2.3: floating images anchored to this paragraph.
        floating_images: Vec<FloatingImage>,
    },
    Table {
        rows: Vec<TableRowInput>,
        col_widths: Vec<Pt>,
        /// §17.4.38: resolved table border configuration.
        border_config: Option<super::table::TableBorderConfig>,
        /// §17.4.51: table indentation from left margin.
        indent: Pt,
        /// §17.4.28: table horizontal alignment.
        alignment: Option<dxpdf_docx_model::model::Alignment>,
        /// §17.4.58: floating table — text wraps around it.
        /// (table_width, right_gap, bottom_gap) for float positioning.
        float_info: Option<(Pt, Pt, Pt)>,
    },
}

/// Lay out a sequence of blocks into pages.
pub fn layout_section(
    blocks: &[LayoutBlock],
    config: &PageConfig,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
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
    let page_bottom = config.page_size.height - config.margins.bottom;
    // Effective bottom boundary — reduced by footnote height.
    let mut bottom = page_bottom;
    // Footnotes collected for the current page.
    let mut page_footnotes: Vec<(&[Fragment], &ParagraphStyle)> = Vec::new();
    // Unified float tracking: tables + images.
    let mut page_floats: Vec<super::float::ActiveFloat> = Vec::new();
    // Per-page absolute float cache: absolute floats from paragraphs on the
    // current page (including upcoming ones via forward scan). Reset on page break.
    let mut current_page_abs_floats: Vec<super::float::ActiveFloat> = Vec::new();
    // Index of the first block on the current page (for forward scanning).
    let mut page_start_block: usize = 0;
    // Whether we need to rescan absolute floats for this page.
    let mut abs_floats_dirty = true;

    for (block_idx, block) in blocks.iter().enumerate() {
        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                page_break_before,
                footnotes,
                floating_images,
            } => {
                // §17.3.1.23: force a new page before this paragraph.
                if *page_break_before && cursor_y > config.margins.top {
                    if !page_footnotes.is_empty() {
                        render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                        page_footnotes.clear();
                    }
                    pages.push(std::mem::replace(
                        &mut current_page,
                        LayoutedPage::new(config.page_size),
                    ));
                    cursor_y = config.margins.top;
                    bottom = page_bottom;
                    page_start_block = block_idx;
                    abs_floats_dirty = true;
                }

                let mut effective_style = style.clone();

                // Register floating images (both relative and absolute).
                for fi in floating_images.iter() {
                    let (y_start, y_end) = match fi.y {
                        FloatingImageY::RelativeToParagraph(offset) => {
                            (cursor_y + offset, cursor_y + offset + fi.size.height)
                        }
                        FloatingImageY::Absolute(img_y) => {
                            (img_y, img_y + fi.size.height)
                        }
                    };
                    let dist_h = Pt::new(9.0);
                    page_floats.push(super::float::ActiveFloat {
                        page_x: fi.x - dist_h,
                        page_y_start: y_start,
                        page_y_end: y_end,
                        width: fi.size.width + dist_h * 2.0,
                    });
                }

                // Prune expired floats, then pass float context for per-line adjustment.
                super::float::prune_floats(&mut page_floats, cursor_y);

                // Forward-scan absolute floats from upcoming paragraphs on the
                // current page. These may need to constrain text before the owning
                // paragraph is processed. Only rescan when the page changes.
                if abs_floats_dirty {
                    current_page_abs_floats.clear();
                    for (fi_idx, future_block) in blocks[page_start_block..].iter().enumerate() {
                        if let LayoutBlock::Paragraph { floating_images: fi_list, page_break_before, .. } = future_block {
                            // Stop scanning at the next explicit page break
                            // (skip the first block since it may be the one that triggered the new page).
                            if *page_break_before && fi_idx > 0 {
                                break;
                            }
                            for fi in fi_list {
                                if let FloatingImageY::Absolute(img_y) = fi.y {
                                    let dist_h = Pt::new(9.0);
                                    current_page_abs_floats.push(super::float::ActiveFloat {
                                        page_x: fi.x - dist_h,
                                        page_y_start: img_y,
                                        page_y_end: img_y + fi.size.height,
                                        width: fi.size.width + dist_h * 2.0,
                                    });
                                }
                            }
                        }
                    }
                    abs_floats_dirty = false;
                }

                // Merge page_floats with forward-scanned absolute floats (dedup).
                let mut effective_floats = page_floats.clone();
                for af in &current_page_abs_floats {
                    let already_present = effective_floats.iter().any(|pf|
                        (pf.page_x - af.page_x).raw().abs() < 0.1
                        && (pf.page_y_start - af.page_y_start).raw().abs() < 0.1
                        && (pf.page_y_end - af.page_y_end).raw().abs() < 0.1
                    );
                    if !already_present {
                        effective_floats.push(af.clone());
                    }
                }
                effective_style.page_floats = effective_floats;
                effective_style.page_y = cursor_y;
                effective_style.page_x = config.margins.left;
                effective_style.page_content_width = content_width;

                let para = layout_paragraph(
                    fragments,
                    &constraints,
                    &effective_style,
                    default_line_height,
                    measure_text,
                );

                // Page break if paragraph doesn't fit
                if cursor_y + para.size.height > bottom && cursor_y > config.margins.top {
                    if !page_footnotes.is_empty() {
                        render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                        page_footnotes.clear();
                    }
                    pages.push(std::mem::replace(
                        &mut current_page,
                        LayoutedPage::new(config.page_size),
                    ));
                    cursor_y = config.margins.top;
                    bottom = page_bottom;
                    page_start_block = block_idx;
                    abs_floats_dirty = true;
                }

                // Offset commands to absolute page position
                for mut cmd in para.commands {
                    cmd.shift_y(cursor_y);
                    cmd.shift_x(config.margins.left);
                    current_page.commands.push(cmd);
                }

                cursor_y += para.size.height;

                // §20.4.2.3: emit floating images (already registered above).
                for fi in floating_images {
                    let img_y = match fi.y {
                        FloatingImageY::Absolute(y) => y,
                        FloatingImageY::RelativeToParagraph(offset) => {
                            cursor_y - para.size.height + offset
                        }
                    };
                    current_page.commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                        image_data: fi.image_data.clone(),
                    });
                }

                // Collect footnotes for this page and reserve space at bottom.
                if !footnotes.is_empty() {
                    let fn_constraints = super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
                    let sep_height = Pt::new(4.0); // separator line + gap
                    for (fn_frags, fn_style) in footnotes {
                        let fn_para = super::paragraph::layout_paragraph(
                            fn_frags, &fn_constraints, fn_style, default_line_height, measure_text,
                        );
                        // Only add separator space for the first footnote on this page.
                        if page_footnotes.is_empty() {
                            bottom -= sep_height;
                        }
                        bottom -= fn_para.size.height;
                        page_footnotes.push((fn_frags, fn_style));
                    }
                }
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                border_config,
                indent,
                alignment,
                float_info,
            } => {
                let table = layout_table(
                    rows,
                    col_widths,
                    &constraints,
                    default_line_height,
                    border_config.as_ref(),
                    measure_text,
                );

                // §17.4.28 / §17.4.51: compute table x position from
                // alignment and indent.
                let table_x = table_x_offset(
                    *alignment, *indent, table.size.width, content_width,
                    config.margins.left,
                );

                // §17.4.58: floating table — render at current position and
                // register as a float so subsequent text wraps around it.
                if let Some((_, right_gap, bottom_gap)) = float_info {
                    if cursor_y + table.size.height > bottom && cursor_y > config.margins.top {
                        if !page_footnotes.is_empty() {
                            render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                            page_footnotes.clear();
                        }
                        pages.push(std::mem::replace(
                            &mut current_page,
                            LayoutedPage::new(config.page_size),
                        ));
                        cursor_y = config.margins.top;
                        bottom = page_bottom;
                        page_start_block = block_idx;
                        abs_floats_dirty = true;
                    }

                    let float_y_start = cursor_y;
                    let float_y_end = cursor_y + table.size.height + *bottom_gap;

                    for mut cmd in table.commands {
                        cmd.shift_y(cursor_y);
                        cmd.shift_x(table_x);
                        current_page.commands.push(cmd);
                    }

                    // Register as unified float for text wrapping.
                    page_floats.push(super::float::ActiveFloat {
                        page_x: table_x,
                        page_y_start: float_y_start,
                        page_y_end: float_y_end,
                        width: table.size.width + *right_gap,
                    });
                    continue;
                }

                // Non-floating table: page break if table doesn't fit.
                if cursor_y + table.size.height > bottom && cursor_y > config.margins.top {
                    if !page_footnotes.is_empty() {
                        render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                        page_footnotes.clear();
                    }
                    pages.push(std::mem::replace(
                        &mut current_page,
                        LayoutedPage::new(config.page_size),
                    ));
                    cursor_y = config.margins.top;
                    bottom = page_bottom;
                    page_start_block = block_idx;
                    abs_floats_dirty = true;
                }

                for mut cmd in table.commands {
                    cmd.shift_y(cursor_y);
                    cmd.shift_x(table_x);
                    current_page.commands.push(cmd);
                }

                cursor_y += table.size.height;
            }
        }
    }

    // Render footnotes on the current (last) page.
    if !page_footnotes.is_empty() {
        render_page_footnotes(
            &mut current_page, config, &page_footnotes,
            default_line_height, measure_text, separator_indent,
        );
    }

    // Push the last page (even if empty — ensure at least one page)
    pages.push(current_page);

    pages
}

/// Render footnotes at the bottom of a page.
fn render_page_footnotes(
    page: &mut LayoutedPage,
    config: &PageConfig,
    footnotes: &[(&[Fragment], &ParagraphStyle)],
    default_line_height: Pt,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
) {
    use super::draw_command::DrawCommand;
    use super::paragraph::layout_paragraph;

    let content_width = config.content_width();
    let constraints = super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
    let page_bottom = config.page_size.height - config.margins.bottom;

    // Layout all footnotes to compute total height.
    let mut footnote_layouts = Vec::new();
    let mut total_height = Pt::new(4.0); // separator + gap
    for (frags, style) in footnotes {
        let para = layout_paragraph(frags, &constraints, style, default_line_height, measure_text);
        total_height += para.size.height;
        footnote_layouts.push(para);
    }

    let footnote_top = page_bottom - total_height;

    // §17.11.23: separator line positioned per default paragraph indent.
    let sep_x = config.margins.left + separator_indent;
    let sep_width = content_width * 0.33;
    page.commands.push(DrawCommand::Line {
        line: crate::geometry::PtLineSegment::new(
            crate::geometry::PtOffset::new(sep_x, footnote_top),
            crate::geometry::PtOffset::new(sep_x + sep_width, footnote_top),
        ),
        color: crate::resolve::color::RgbColor::BLACK,
        width: Pt::new(0.5),
    });

    // Render footnote paragraphs.
    let mut cursor_y = footnote_top + Pt::new(4.0);
    for para in footnote_layouts {
        for mut cmd in para.commands {
            cmd.shift_y(cursor_y);
            cmd.shift_x(config.margins.left);
            page.commands.push(cmd);
        }
        cursor_y += para.size.height;
    }
}

/// §17.4.28: compute the table's x offset based on alignment and indent.
fn table_x_offset(
    alignment: Option<dxpdf_docx_model::model::Alignment>,
    indent: Pt,
    table_width: Pt,
    content_width: Pt,
    margin_left: Pt,
) -> Pt {
    use dxpdf_docx_model::model::Alignment;
    match alignment {
        Some(Alignment::Center) => {
            margin_left + (content_width - table_width) * 0.5
        }
        Some(Alignment::End) => {
            margin_left + content_width - table_width
        }
        _ => margin_left + indent,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{PtEdgeInsets, PtSize};
    use crate::layout::cell::CellBlock;
    use crate::layout::draw_command::DrawCommand;
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
                char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width), trimmed_width: Pt::new(width),
            height: Pt::new(height),
            ascent: Pt::new(height * 0.7),
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO,
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
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
        }
    }

    #[test]
    fn empty_blocks_produces_one_empty_page() {
        let pages = layout_section(&[], &small_config(), None, Pt::ZERO, Pt::new(14.0));
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn single_paragraph_on_one_page() {
        let blocks = vec![para_block("hello", 30.0)];
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0));

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
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0));

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
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0));

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
        let pages = layout_section(&[], &config, None, Pt::ZERO, Pt::new(14.0));
        assert_eq!(pages[0].page_size, config.page_size);
    }

    #[test]
    fn many_paragraphs_produce_multiple_pages() {
        // 20 paragraphs at 14pt each = 280pt
        // Content area = 80pt → need 4 pages (80/14 = 5.7 paras per page)
        let blocks: Vec<_> = (0..20).map(|i| para_block(&format!("p{i}"), 30.0)).collect();
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0));

        assert_eq!(pages.len(), 4);
    }

    #[test]
    fn table_on_page() {
        let blocks = vec![LayoutBlock::Table {
            rows: vec![TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![CellBlock::Paragraph {
                        fragments: vec![text_frag("cell", 30.0, 14.0)],
                        style: ParagraphStyle::default(),
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None, cell_borders: None, vertical_merge: None, vertical_align: crate::layout::table::CellVAlign::Top,
                }],
                height_rule: None,
            }],
            col_widths: vec![Pt::new(100.0)],
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
        }];

        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0));
        assert_eq!(pages.len(), 1);

        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }
}
