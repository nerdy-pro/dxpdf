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
pub mod geometry;
pub mod layout;
pub mod resolve;

use dxpdf_docx_model::model::Document;

use crate::dimension::Pt;
use crate::layout::draw_command::LayoutedPage;
use crate::layout::page::PageConfig;
use crate::layout::section::{layout_section, LayoutBlock};
use crate::resolve::ResolvedDocument;

/// Default line height when no font metrics are available.
const DEFAULT_LINE_HEIGHT: Pt = Pt::new(14.0);

/// Resolve and lay out a document, producing positioned pages with draw commands.
///
/// This performs phases 1 (resolve) and 3 (layout) of the pipeline.
/// Phase 2 (measure with Skia fonts) and phase 4 (paint to PDF) require
/// `skia-safe` and are handled separately.
pub fn resolve_and_layout(doc: &Document) -> (ResolvedDocument, Vec<LayoutedPage>) {
    let resolved = resolve::resolve(doc);
    let pages = layout_document(&resolved);
    (resolved, pages)
}

/// Lay out a resolved document into pages.
pub fn layout_document(resolved: &ResolvedDocument) -> Vec<LayoutedPage> {
    let mut all_pages = Vec::new();

    for section in &resolved.sections {
        let config = PageConfig::from_section(&section.properties);

        // Convert resolved blocks to layout blocks.
        // For now, this is a simplified bridge — a full implementation would
        // use Skia font metrics for text measurement.
        let layout_blocks = build_layout_blocks(section, &config);

        let mut pages = layout_section(&layout_blocks, &config, DEFAULT_LINE_HEIGHT);
        all_pages.append(&mut pages);
    }

    // Ensure at least one page
    if all_pages.is_empty() {
        all_pages.push(LayoutedPage::new(PageConfig::default().page_size));
    }

    all_pages
}

/// Build layout blocks from a resolved section.
/// This is the bridge between the resolve output and the layout input.
fn build_layout_blocks(
    section: &resolve::sections::ResolvedSection,
    config: &PageConfig,
) -> Vec<LayoutBlock> {
    use dxpdf_docx_model::model::Block;
    use layout::cell::CellBlock;
    use layout::fragment::{collect_fragments, FontProps};
    use layout::table::{compute_column_widths, TableCellInput, TableRowInput};

    let default_family = "Helvetica";
    let default_size = Pt::new(12.0);

    // Dummy text measurer — width = len * 6, height = size, ascent = size * 0.8
    // A real implementation uses Skia FontMgr.
    let measure = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
        let w = Pt::new(text.len() as f32 * font.size.raw() * 0.5);
        let h = font.size;
        let a = font.size * 0.8;
        (w, h, a)
    };

    let mut blocks = Vec::new();

    for block in &section.blocks {
        match block {
            Block::Paragraph(p) => {
                let fragments = collect_fragments(
                    &p.content,
                    default_family,
                    default_size,
                    None,
                    &measure,
                );
                let style = paragraph_style_from_props(&p.properties);
                blocks.push(LayoutBlock::Paragraph { fragments, style });
            }
            Block::Table(t) => {
                let num_cols = t.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
                let grid_cols: Vec<Pt> = t.grid.iter().map(|g| Pt::from(g.width)).collect();
                let col_widths =
                    compute_column_widths(&grid_cols, num_cols, config.content_width());

                let rows: Vec<TableRowInput> = t
                    .rows
                    .iter()
                    .map(|row| {
                        let cells: Vec<TableCellInput> = row
                            .cells
                            .iter()
                            .map(|cell| {
                                let cell_blocks: Vec<CellBlock> = cell
                                    .content
                                    .iter()
                                    .filter_map(|b| {
                                        if let Block::Paragraph(p) = b {
                                            let frags = collect_fragments(
                                                &p.content,
                                                default_family,
                                                default_size,
                                                None,
                                                &measure,
                                            );
                                            Some(CellBlock {
                                                fragments: frags,
                                                style: paragraph_style_from_props(&p.properties),
                                            })
                                        } else {
                                            None
                                        }
                                    })
                                    .collect();

                                TableCellInput {
                                    blocks: cell_blocks,
                                    margins: geometry::PtEdgeInsets::new(
                                        Pt::new(2.0),
                                        Pt::new(5.0),
                                        Pt::new(2.0),
                                        Pt::new(5.0),
                                    ),
                                    grid_span: cell.properties.grid_span.unwrap_or(1),
                                    shading: None,
                                }
                            })
                            .collect();

                        TableRowInput {
                            cells,
                            min_height: None,
                        }
                    })
                    .collect();

                blocks.push(LayoutBlock::Table {
                    rows,
                    col_widths,
                    draw_borders: true,
                });
            }
            Block::SectionBreak(_) => {} // already split by resolve
        }
    }

    blocks
}

fn paragraph_style_from_props(
    props: &dxpdf_docx_model::model::ParagraphProperties,
) -> layout::paragraph::ParagraphStyle {
    use dxpdf_docx_model::model::{FirstLineIndent, LineSpacing};
    use layout::paragraph::{LineSpacingRule, ParagraphStyle};

    let indent_left = props
        .indentation
        .and_then(|i| i.start)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);

    let indent_right = props
        .indentation
        .and_then(|i| i.end)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);

    let indent_first_line = props
        .indentation
        .and_then(|i| i.first_line)
        .map(|fl| match fl {
            FirstLineIndent::FirstLine(v) => Pt::from(v),
            FirstLineIndent::Hanging(v) => -Pt::from(v),
            FirstLineIndent::None => Pt::ZERO,
        })
        .unwrap_or(Pt::ZERO);

    let space_before = props
        .spacing
        .and_then(|s| s.before)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);

    let space_after = props
        .spacing
        .and_then(|s| s.after)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);

    let line_spacing = props
        .spacing
        .and_then(|s| s.line)
        .map(|ls| match ls {
            LineSpacing::Auto(v) => LineSpacingRule::Auto(Pt::from(v).raw() / 12.0),
            LineSpacing::Exact(v) => LineSpacingRule::Exact(Pt::from(v)),
            LineSpacing::AtLeast(v) => LineSpacingRule::AtLeast(Pt::from(v)),
        })
        .unwrap_or(LineSpacingRule::Auto(1.0));

    ParagraphStyle {
        alignment: props.alignment.unwrap_or(dxpdf_docx_model::model::Alignment::Start),
        space_before,
        space_after,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
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
