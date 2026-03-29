//! PDF renderer for dxpdf — measure, layout, and paint pipeline.
//!
//! Takes a parsed `Document` from `dxpdf-docx` and produces PDF bytes.
//!
//! # Pipeline
//!
//! 1. **Resolve** — flatten style inheritance, split sections, extract images/fonts
//! 2. **Layout** — fit content into pages using constraint-based layout
//! 3. **Paint** — emit draw commands to Skia PDF canvas (requires `skia-safe`)

pub mod dimension;
pub mod error;
pub mod fonts;
pub mod geometry;
pub mod layout;
pub mod painter;
pub mod resolve;
pub mod skia_conv;

use dxpdf_docx_model::model::Document;

use crate::layout::build::{BuildContext, build_section_blocks, collect_fragments_from_blocks, default_line_height};
use crate::layout::draw_command::LayoutedPage;
use crate::layout::header_footer::render_headers_footers;
use crate::layout::page::PageConfig;
use crate::layout::section::layout_section;
use crate::resolve::ResolvedDocument;

/// Render a parsed DOCX document to PDF bytes.
///
/// Estimate the cursor_y position from the last page's draw commands.
/// Used to determine where a continuous section should start on the page.
fn estimate_cursor_y(page: &layout::draw_command::LayoutedPage, config: &layout::page::PageConfig) -> dimension::Pt {
    let mut max_y = config.margins.top;
    for cmd in &page.commands {
        let bottom = match cmd {
            layout::draw_command::DrawCommand::Text { position, font_size, .. } => {
                position.y + *font_size
            }
            layout::draw_command::DrawCommand::Image { rect, .. } => {
                rect.origin.y + rect.size.height
            }
            layout::draw_command::DrawCommand::Rect { rect, .. } => {
                rect.origin.y + rect.size.height
            }
            layout::draw_command::DrawCommand::Line { line, .. } => {
                line.end.y
            }
            _ => continue,
        };
        if bottom > max_y {
            max_y = bottom;
        }
    }
    max_y
}

/// Full pipeline: resolve → preload fonts → layout → paint.
pub fn render(doc: &Document) -> Result<Vec<u8>, error::RenderError> {
    let font_mgr = skia_safe::FontMgr::new();
    render_with_font_mgr(doc, &font_mgr)
}

/// Render with a pre-configured FontMgr (for reuse across calls).
pub fn render_with_font_mgr(
    doc: &Document,
    font_mgr: &skia_safe::FontMgr,
) -> Result<Vec<u8>, error::RenderError> {
    let resolved = resolve::resolve(doc);
    fonts::register_embedded_fonts(font_mgr, &doc.embedded_fonts);
    fonts::preload_fonts(font_mgr, &resolved.font_families);
    let pages = layout_document(&resolved, font_mgr);
    painter::render_to_pdf(&pages, font_mgr)
}

/// Resolve and lay out a document without painting to PDF.
/// Uses a real FontMgr for text measurement.
pub fn resolve_and_layout(doc: &Document) -> (ResolvedDocument, Vec<LayoutedPage>) {
    let font_mgr = skia_safe::FontMgr::new();
    let resolved = resolve::resolve(doc);
    fonts::preload_fonts(&font_mgr, &resolved.font_families);
    let pages = layout_document(&resolved, &font_mgr);
    (resolved, pages)
}

/// Lay out a resolved document using Skia font metrics.
pub fn layout_document(
    resolved: &ResolvedDocument,
    font_mgr: &skia_safe::FontMgr,
) -> Vec<LayoutedPage> {
    let measurer = layout::measurer::TextMeasurer::new(font_mgr.clone());
    let ctx = BuildContext {
        measurer: &measurer,
        resolved,
        footnote_counter: std::cell::Cell::new(0),
        endnote_counter: std::cell::Cell::new(0),
        list_counters: std::cell::RefCell::new(std::collections::HashMap::new()),
    };
    let dlh = default_line_height(&ctx);
    let mut all_pages = Vec::new();
    let mut all_endnotes = Vec::new();
    let mut last_config = PageConfig::default();

    // §17.11.23: footnote separator indent from default paragraph style.
    let separator_indent = resolved.default_paragraph_style_id.as_ref()
        .and_then(|id| resolved.styles.get(id))
        .and_then(|s| s.paragraph.indentation)
        .and_then(|ind| ind.first_line)
        .map(|fl| match fl {
            dxpdf_docx_model::model::FirstLineIndent::FirstLine(v) => dimension::Pt::from(v),
            _ => dimension::Pt::ZERO,
        })
        .unwrap_or(dimension::Pt::ZERO);

    // §17.6.22: track continuation state for `Continuous` section breaks.
    let mut pending_continuation: Option<layout::section::ContinuationState> = None;

    for section in &resolved.sections {
        let config = PageConfig::from_section(&section.properties);
        let built = build_section_blocks(section, &config, &ctx);
        let measure_fn = |text: &str, font: &layout::fragment::FontProps| -> (dimension::Pt, layout::fragment::TextMetrics) {
            measurer.measure(text, font)
        };

        // §17.6.22: continuous sections continue on the current page.
        let continuation = if section.properties.section_type == Some(dxpdf_docx_model::model::SectionType::Continuous) {
            pending_continuation.take()
        } else {
            pending_continuation = None;
            None
        };

        let mut pages = layout_section(&built.blocks, &config, Some(&measure_fn), separator_indent, dlh, continuation);

        // Collect endnotes for rendering at document end.
        all_endnotes.extend(built.endnotes);

        // Render headers/footers onto each page.
        let header_content = section.header.as_ref()
            .map(|blocks| collect_fragments_from_blocks(blocks, &ctx));
        let footer_content = section.footer.as_ref()
            .map(|blocks| collect_fragments_from_blocks(blocks, &ctx));
        render_headers_footers(
            &mut pages, &config,
            header_content.as_ref(), footer_content.as_ref(),
            dlh,
        );

        last_config = config;

        // Check if the NEXT section is continuous — if so, save the last page
        // as continuation state instead of appending it.
        // (We peek ahead by checking the section index.)
        let next_is_continuous = {
            let section_idx = resolved.sections.iter().position(|s| std::ptr::eq(s, section));
            section_idx.and_then(|i| resolved.sections.get(i + 1))
                .is_some_and(|next| next.properties.section_type == Some(dxpdf_docx_model::model::SectionType::Continuous))
        };

        if next_is_continuous && !pages.is_empty() {
            let last_page = pages.pop().unwrap();
            // Estimate cursor_y from the last command on the page.
            let cursor_y = estimate_cursor_y(&last_page, &last_config);
            pending_continuation = Some(layout::section::ContinuationState {
                page: last_page,
                cursor_y,
            });
        }

        all_pages.append(&mut pages);
    }

    // Render endnotes on a new page at the end of the document.
    if !all_endnotes.is_empty() {
        let measure_fn = |text: &str, font: &layout::fragment::FontProps| -> (dimension::Pt, layout::fragment::TextMetrics) {
            measurer.measure(text, font)
        };
        let mut endnote_page = LayoutedPage::new(last_config.page_size);
        let content_width = last_config.content_width();
        let constraints = layout::BoxConstraints::tight_width(content_width, dimension::Pt::INFINITY);
        let mut cursor_y = last_config.margins.top;

        // Separator line.
        let sep_width = content_width * 0.33;
        let sep_x = last_config.margins.left + separator_indent;
        endnote_page.commands.push(layout::draw_command::DrawCommand::Line {
            line: crate::geometry::PtLineSegment::new(
                crate::geometry::PtOffset::new(sep_x, cursor_y),
                crate::geometry::PtOffset::new(sep_x + sep_width, cursor_y),
            ),
            color: crate::resolve::color::RgbColor::BLACK,
            width: dimension::Pt::new(0.5),
        });
        cursor_y += dimension::Pt::new(4.0);

        for (_, frags, style) in &all_endnotes {
            let para = layout::paragraph::layout_paragraph(
                frags, &constraints, style, dlh, Some(&measure_fn),
            );
            for mut cmd in para.commands {
                cmd.shift_y(cursor_y);
                cmd.shift_x(last_config.margins.left);
                endnote_page.commands.push(cmd);
            }
            cursor_y += para.size.height;
        }
        all_pages.push(endnote_page);
    }

    if all_pages.is_empty() {
        all_pages.push(LayoutedPage::new(PageConfig::default().page_size));
    }

    all_pages
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::model::*;
    use std::collections::HashMap;

    fn empty_doc() -> Document {
        Document {
            settings: DocumentSettings::default(),
            theme: None,
            styles: StyleSheet::default(),
            numbering: NumberingDefinitions::default(),
            body: vec![],
            final_section: SectionProperties::default(),
            headers: HashMap::new(),
            footers: HashMap::new(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            media: HashMap::new(),
            embedded_fonts: vec![],
        }
    }

    fn para(text: &str) -> Block {
        Block::Paragraph(Box::new(Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: vec![Inline::TextRun(Box::new(TextRun {
                style_id: None,
                properties: RunProperties::default(),
                text: text.to_string(),
                rsids: RevisionIds::default(),
            }))],
            rsids: ParagraphRevisionIds::default(),
        }))
    }

    #[test]
    fn resolve_and_layout_empty_doc() {
        let doc = empty_doc();
        let (resolved, pages) = resolve_and_layout(&doc);

        assert_eq!(resolved.sections.len(), 1);
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn resolve_and_layout_with_paragraphs() {
        let mut doc = empty_doc();
        doc.body = vec![para("hello"), para("world")];

        let (_, pages) = resolve_and_layout(&doc);

        assert_eq!(pages.len(), 1);
        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, layout::draw_command::DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 2);
    }

    #[test]
    fn resolve_and_layout_with_table() {
        let mut doc = empty_doc();
        doc.body = vec![Block::Table(Box::new(Table {
            properties: TableProperties::default(),
            grid: vec![
                GridColumn {
                    width: dxpdf_docx_model::dimension::Dimension::new(4680),
                },
                GridColumn {
                    width: dxpdf_docx_model::dimension::Dimension::new(4680),
                },
            ],
            rows: vec![TableRow {
                properties: TableRowProperties::default(),
                cells: vec![
                    TableCell {
                        properties: TableCellProperties::default(),
                        content: vec![para("A")],
                    },
                    TableCell {
                        properties: TableCellProperties::default(),
                        content: vec![para("B")],
                    },
                ],
                rsids: TableRowRevisionIds::default(),
            }],
        }))];

        let (_, pages) = resolve_and_layout(&doc);
        assert_eq!(pages.len(), 1);

        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, layout::draw_command::DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 2, "two cells = two text commands");
    }

    #[test]
    fn layout_respects_page_size() {
        let mut doc = empty_doc();
        doc.final_section = SectionProperties {
            page_size: Some(PageSize {
                width: Some(dxpdf_docx_model::dimension::Dimension::new(12240)),
                height: Some(dxpdf_docx_model::dimension::Dimension::new(15840)),
                orientation: None,
            }),
            ..Default::default()
        };

        let (_, pages) = resolve_and_layout(&doc);
        assert_eq!(pages[0].page_size.width.raw(), 612.0);
        assert_eq!(pages[0].page_size.height.raw(), 792.0);
    }
}
