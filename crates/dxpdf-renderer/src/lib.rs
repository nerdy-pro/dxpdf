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

use crate::dimension::Pt;
use crate::layout::draw_command::LayoutedPage;
use crate::layout::measurer::TextMeasurer;
use crate::layout::page::PageConfig;
use crate::layout::section::{layout_section, LayoutBlock};
use crate::resolve::ResolvedDocument;

const DEFAULT_FAMILY: &str = "Helvetica";
const DEFAULT_SIZE: Pt = Pt::new(12.0);

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
    fonts::preload_fonts(font_mgr, &resolved.font_families);
    let pages = layout_document_with_fonts(&resolved, font_mgr);
    painter::render_to_pdf(&pages, font_mgr)
}

/// Resolve and lay out a document without painting to PDF.
/// Uses a real FontMgr for text measurement.
pub fn resolve_and_layout(doc: &Document) -> (ResolvedDocument, Vec<LayoutedPage>) {
    let font_mgr = skia_safe::FontMgr::new();
    let resolved = resolve::resolve(doc);
    fonts::preload_fonts(&font_mgr, &resolved.font_families);
    let pages = layout_document_with_fonts(&resolved, &font_mgr);
    (resolved, pages)
}

/// Lay out a resolved document using Skia font metrics.
pub fn layout_document_with_fonts(
    resolved: &ResolvedDocument,
    font_mgr: &skia_safe::FontMgr,
) -> Vec<LayoutedPage> {
    let measurer = TextMeasurer::new(font_mgr.clone());
    let default_line_height = measurer.default_line_height(DEFAULT_FAMILY, DEFAULT_SIZE);
    let mut all_pages = Vec::new();

    for section in &resolved.sections {
        let config = PageConfig::from_section(&section.properties);
        let layout_blocks = build_layout_blocks(section, &config, &measurer, &resolved.styles);
        let mut pages = layout_section(&layout_blocks, &config, default_line_height);
        all_pages.append(&mut pages);
    }

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
    measurer: &TextMeasurer,
    styles: &std::collections::HashMap<dxpdf_docx_model::model::StyleId, resolve::styles::ResolvedStyle>,
) -> Vec<LayoutBlock> {
    use dxpdf_docx_model::model::Block;
    use layout::cell::CellBlock;
    use layout::fragment::{collect_fragments, FontProps};
    use layout::table::{compute_column_widths, TableCellInput, TableRowInput};

    let measure = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
        measurer.measure(text, font)
    };

    let mut blocks = Vec::new();
    // Drop cap fragments to prepend to the next paragraph.
    let mut pending_dropcap: Option<Vec<layout::fragment::Fragment>> = None;

    for block in &section.blocks {
        match block {
            Block::Paragraph(p) => {
                let (default_family, default_size, merged_props) =
                    resolve_paragraph_defaults(p, styles);

                // Detect drop cap: paragraph with frame_properties.drop_cap = Drop or Margin.
                let is_dropcap = merged_props
                    .frame_properties
                    .and_then(|fp| fp.drop_cap)
                    .is_some_and(|dc| {
                        matches!(
                            dc,
                            dxpdf_docx_model::model::DropCap::Drop
                                | dxpdf_docx_model::model::DropCap::Margin
                        )
                    });

                let mut fragments = collect_fragments(
                    &p.content,
                    &default_family,
                    default_size,
                    None,
                    &measure,
                );

                if is_dropcap {
                    // Stash drop cap fragments — they'll be prepended to the next paragraph.
                    pending_dropcap = Some(fragments);
                    continue;
                }

                // Prepend any pending drop cap fragments.
                if let Some(mut dc_frags) = pending_dropcap.take() {
                    dc_frags.append(&mut fragments);
                    fragments = dc_frags;
                }

                let style = paragraph_style_from_props(&merged_props);
                blocks.push(LayoutBlock::Paragraph { fragments, style });
            }
            Block::Table(t) => {
                // Use grid column count (not cell count) — cells may span multiple grid columns.
                let num_cols = if t.grid.is_empty() {
                    t.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
                } else {
                    t.grid.len()
                };
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
                                            let (df, ds, mp) =
                                                resolve_paragraph_defaults(p, styles);
                                            let frags = collect_fragments(
                                                &p.content,
                                                &df,
                                                ds,
                                                None,
                                                &measure,
                                            );
                                            Some(CellBlock {
                                                fragments: frags,
                                                style: paragraph_style_from_props(&mp),
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

/// Resolve a paragraph's effective defaults by looking up its style_id
/// in the resolved styles map and merging with direct properties.
///
/// Returns (default_font_family, default_font_size, merged_paragraph_properties).
fn resolve_paragraph_defaults(
    para: &dxpdf_docx_model::model::Paragraph,
    styles: &std::collections::HashMap<dxpdf_docx_model::model::StyleId, resolve::styles::ResolvedStyle>,
) -> (String, Pt, dxpdf_docx_model::model::ParagraphProperties) {
    use resolve::fonts::effective_font;
    use resolve::properties::merge_paragraph_properties;

    let mut para_props = para.properties.clone();
    let mut default_family = DEFAULT_FAMILY.to_string();
    let mut default_size = DEFAULT_SIZE;

    // Look up the paragraph's style and merge its properties as base.
    if let Some(ref style_id) = para.style_id {
        if let Some(resolved_style) = styles.get(style_id) {
            // Merge: direct paragraph properties override style properties.
            merge_paragraph_properties(&mut para_props, &resolved_style.paragraph);

            // Use the style's run properties as defaults for font family/size.
            if let Some(f) = effective_font(&resolved_style.run.fonts) {
                default_family = f.to_string();
            }
            if let Some(fs) = resolved_style.run.font_size {
                default_size = Pt::from(fs);
            }
        }
    }

    (default_family, default_size, para_props)
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
