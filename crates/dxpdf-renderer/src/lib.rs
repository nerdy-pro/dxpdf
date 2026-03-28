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

    for section in &resolved.sections {
        let config = PageConfig::from_section(&section.properties);
        let built = build_section_blocks(section, &config, &ctx);
        let measure_fn = |text: &str, font: &layout::fragment::FontProps| -> (dimension::Pt, dimension::Pt, dimension::Pt) {
            measurer.measure(text, font)
        };
        let mut pages = layout_section(&built.blocks, &config, Some(&measure_fn), separator_indent, dlh);

        // Collect endnotes for rendering at document end.
        all_endnotes.extend(built.endnotes);

        // Render headers/footers onto each page.
        let header_frags = section.header.as_ref()
            .map(|blocks| collect_fragments_from_blocks(blocks, &ctx));
        let footer_frags = section.footer.as_ref()
            .map(|blocks| collect_fragments_from_blocks(blocks, &ctx));
        render_headers_footers(
            &mut pages, &config,
            header_frags.as_deref(), footer_frags.as_deref(),
            dlh,
        );

        last_config = config;
        all_pages.append(&mut pages);
    }

    // Render endnotes at the end of the document.
    if !all_endnotes.is_empty() {
        let measure_fn = |text: &str, font: &layout::fragment::FontProps| -> (dimension::Pt, dimension::Pt, dimension::Pt) {
            measurer.measure(text, font)
        };
        render_endnotes(&mut all_pages, &last_config, &all_endnotes, dlh, Some(&measure_fn));
    }

    if all_pages.is_empty() {
        all_pages.push(LayoutedPage::new(PageConfig::default().page_size));
    }

    all_pages
}

/// Render endnotes at the bottom of the last page.
fn render_endnotes(
    pages: &mut [LayoutedPage],
    config: &PageConfig,
    footnotes: &[(String, Vec<layout::fragment::Fragment>, layout::paragraph::ParagraphStyle)],
    default_line_height: dimension::Pt,
    measure_text: layout::paragraph::MeasureTextFn<'_>,
) {
    use layout::draw_command::DrawCommand;
    use layout::paragraph::layout_paragraph;

    if pages.is_empty() || footnotes.is_empty() {
        return;
    }

    let page = pages.last_mut().unwrap();
    let content_width = config.content_width();
    let constraints = layout::BoxConstraints::tight_width(content_width, dimension::Pt::INFINITY);

    // Start from the bottom margin, building upward.
    let page_bottom = config.page_size.height - config.margins.bottom;

    // Layout all footnotes to compute total height.
    let mut footnote_layouts = Vec::new();
    let mut total_height = dimension::Pt::new(4.0); // separator line + gap
    for (_, frags, style) in footnotes {
        let para = layout_paragraph(frags, &constraints, style, default_line_height, measure_text);
        total_height += para.size.height;
        footnote_layouts.push(para);
    }

    let footnote_top = page_bottom - total_height;

    // Draw separator line (short horizontal rule).
    let sep_y = footnote_top;
    let sep_width = content_width * 0.33;
    page.commands.push(DrawCommand::Line {
        line: crate::geometry::PtLineSegment::new(
            crate::geometry::PtOffset::new(config.margins.left, sep_y),
            crate::geometry::PtOffset::new(config.margins.left + sep_width, sep_y),
        ),
        color: crate::resolve::color::RgbColor::BLACK,
        width: dimension::Pt::new(0.5),
    });

    // Render footnote paragraphs below the separator.
    let mut cursor_y = sep_y + dimension::Pt::new(4.0);
    for para in footnote_layouts {
        for mut cmd in para.commands {
            cmd.shift_y(cursor_y);
            cmd.shift_x(config.margins.left);
            page.commands.push(cmd);
        }
        cursor_y += para.size.height;
    }
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
