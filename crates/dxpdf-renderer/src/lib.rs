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

/// §17.8.3.2: OOXML fallback font when no theme or doc defaults specify one.
const SPEC_FALLBACK_FONT: &str = "Times New Roman";
/// §17.3.2.14: default font size when not specified (10pt = 20 half-points per ECMA-376 §17.3.2.14).
const SPEC_DEFAULT_FONT_SIZE: Pt = Pt::new(10.0);

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
    // Register embedded fonts first — they take priority over system fonts.
    fonts::register_embedded_fonts(font_mgr, &doc.embedded_fonts);
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
    use layout::fragment::FontProps;
    use layout::header_footer::render_headers_footers;

    let measurer = TextMeasurer::new(font_mgr.clone());

    // Derive the default font family and size from the document.
    let doc_font_family = resolved
        .theme
        .as_ref()
        .map(|t| t.minor_font.latin.as_str())
        .filter(|s| !s.is_empty())
        .unwrap_or(SPEC_FALLBACK_FONT);
    let doc_font_size = resolved
        .doc_defaults_run
        .font_size
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE);
    let default_line_height = measurer.default_line_height(doc_font_family, doc_font_size);
    let measure = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
        measurer.measure(text, font)
    };
    let mut all_pages = Vec::new();

    for section in &resolved.sections {
        let config = PageConfig::from_section(&section.properties);
        let layout_blocks = build_layout_blocks(section, &config, &measurer, resolved);
        let mut pages = layout_section(&layout_blocks, &config, default_line_height);

        // Render headers/footers onto each page in this section.
        let header_frags = section.header.as_ref().map(|blocks| {
            collect_fragments_from_blocks(blocks, &measure, resolved)
        });
        let footer_frags = section.footer.as_ref().map(|blocks| {
            collect_fragments_from_blocks(blocks, &measure, resolved)
        });
        render_headers_footers(
            &mut pages,
            &config,
            header_frags.as_deref(),
            footer_frags.as_deref(),
            default_line_height,
        );

        all_pages.append(&mut pages);
    }

    if all_pages.is_empty() {
        all_pages.push(LayoutedPage::new(PageConfig::default().page_size));
    }

    all_pages
}

/// Collect fragments from a list of blocks (used for headers/footers).
fn collect_fragments_from_blocks<F>(
    blocks: &[dxpdf_docx_model::model::Block],
    measure: &F,
    resolved: &ResolvedDocument,
) -> Vec<layout::fragment::Fragment>
where
    F: Fn(&str, &layout::fragment::FontProps) -> (Pt, Pt, Pt),
{
    use dxpdf_docx_model::model::Block;
    use layout::fragment::collect_fragments;

    let mut all_frags = Vec::new();
    for block in blocks {
        if let Block::Paragraph(p) = block {
            let (df, ds, dc, _) = resolve_paragraph_defaults(p, resolved);
            let mut frags = collect_fragments(&p.content, &df, ds, dc, None, measure);
            all_frags.append(&mut frags);
        }
    }
    all_frags
}

/// Build layout blocks from a resolved section.
/// This is the bridge between the resolve output and the layout input.
fn build_layout_blocks(
    section: &resolve::sections::ResolvedSection,
    config: &PageConfig,
    measurer: &TextMeasurer,
    resolved: &ResolvedDocument,
) -> Vec<LayoutBlock> {
    use dxpdf_docx_model::model::Block;
    use layout::cell::CellBlock;
    use layout::fragment::{collect_fragments, FontProps};
    use layout::table::{compute_column_widths, TableCellInput, TableRowInput};

    let media = &resolved.media;

    let measure = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
        measurer.measure(text, font)
    };

    let mut blocks = Vec::new();
    // Drop cap info to attach to the next paragraph.
    let mut pending_dropcap: Option<layout::paragraph::DropCapInfo> = None;

    for block in &section.blocks {
        match block {
            Block::Paragraph(p) => {
                let (default_family, default_size, default_color, merged_props) =
                    resolve_paragraph_defaults(p, resolved);

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

                let drop_cap_lines = merged_props
                    .frame_properties
                    .and_then(|fp| fp.lines)
                    .unwrap_or(3);

                let mut fragments = collect_fragments(
                    &p.content,
                    &default_family,
                    default_size,
                    default_color,
                    None,
                    &measure,
                );
                populate_image_data(&mut fragments, media);
                populate_underline_metrics(&mut fragments, measurer);

                if is_dropcap {
                    let width: Pt = fragments.iter().map(|f| f.width()).sum();
                    let height: Pt = fragments.iter().map(|f| f.height()).fold(Pt::ZERO, Pt::max);
                    let ascent: Pt = fragments.iter().map(|f| {
                        if let layout::fragment::Fragment::Text { ascent, .. } = f { *ascent } else { Pt::ZERO }
                    }).fold(Pt::ZERO, Pt::max);
                    // §17.3.1.11: hSpace from frame properties.
                    let h_space = merged_props
                        .frame_properties
                        .and_then(|fp| fp.h_space)
                        .map(Pt::from)
                        .unwrap_or(Pt::ZERO);
                    pending_dropcap = Some(layout::paragraph::DropCapInfo {
                        fragments,
                        lines: drop_cap_lines,
                        ascent,
                        h_space,
                        width,
                        height,
                    });
                    continue;
                }

                let mut style = paragraph_style_from_props(&merged_props);

                // Attach pending drop cap to this paragraph.
                if let Some(dc) = pending_dropcap.take() {
                    style.drop_cap = Some(dc);
                }

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
                                            let (df, ds, dc, mp) =
                                                resolve_paragraph_defaults(p, resolved);
                                            let mut frags = collect_fragments(
                                                &p.content,
                                                &df,
                                                ds,
                                                dc,
                                                None,
                                                &measure,
                                            );
                                            populate_image_data(&mut frags, media);
                                            populate_underline_metrics(&mut frags, measurer);
                                            Some(CellBlock {
                                                fragments: frags,
                                                style: paragraph_style_from_props(&mp),
                                            })
                                        } else {
                                            None
                                        }
                                    })
                                    .collect();

                                // §17.4.42: cell margins cascade: cell → table → 0.
                                let cell_margins = cell
                                    .properties
                                    .margins
                                    .or(t.properties.cell_margins)
                                    .map(|m| geometry::PtEdgeInsets::new(
                                        Pt::from(m.top),
                                        Pt::from(m.right),
                                        Pt::from(m.bottom),
                                        Pt::from(m.left),
                                    ))
                                    .unwrap_or(geometry::PtEdgeInsets::ZERO);

                                // Resolve cell shading.
                                let shading = cell.properties.shading.map(|s| {
                                    resolve::color::resolve_color(
                                        s.fill,
                                        resolve::color::ColorContext::Background,
                                    )
                                });

                                TableCellInput {
                                    blocks: cell_blocks,
                                    margins: cell_margins,
                                    grid_span: cell.properties.grid_span.unwrap_or(1),
                                    shading,
                                }
                            })
                            .collect();

                        TableRowInput {
                            cells,
                            min_height: row.properties.height.map(|h| Pt::from(h.value)),
                        }
                    })
                    .collect();

                // Draw borders if the table has any border properties defined.
                let has_borders = t.properties.borders.is_some();

                blocks.push(LayoutBlock::Table {
                    rows,
                    col_widths,
                    draw_borders: has_borders,
                });
            }
            Block::SectionBreak(_) => {} // already split by resolve
        }
    }

    blocks
}

/// Populate image data on Fragment::Image fragments by looking up the media map.
fn populate_image_data(
    fragments: &mut [layout::fragment::Fragment],
    media: &std::collections::HashMap<dxpdf_docx_model::model::RelId, Vec<u8>>,
) {
    use dxpdf_docx_model::model::RelId;
    for frag in fragments.iter_mut() {
        if let layout::fragment::Fragment::Image {
            rel_id, image_data, ..
        } = frag
        {
            if image_data.is_none() {
                if let Some(bytes) = media.get(&RelId::new(rel_id.as_str())) {
                    *image_data = Some(bytes.as_slice().into());
                }
            }
        }
    }
}

/// Populate underline position/thickness from Skia font metrics.
fn populate_underline_metrics(
    fragments: &mut [layout::fragment::Fragment],
    measurer: &TextMeasurer,
) {
    for frag in fragments.iter_mut() {
        if let layout::fragment::Fragment::Text { font, .. } = frag {
            if font.underline {
                let (pos, thickness) = measurer.underline_metrics(font);
                font.underline_position = pos;
                font.underline_thickness = thickness;
            }
        }
    }
}

/// Resolve a paragraph's effective defaults by looking up its style_id
/// in the resolved styles map and merging with direct properties.
/// Cascade: direct properties → style properties → doc defaults.
///
/// Returns (default_font_family, default_font_size, default_color, merged_paragraph_properties).
fn resolve_paragraph_defaults(
    para: &dxpdf_docx_model::model::Paragraph,
    resolved: &ResolvedDocument,
) -> (String, Pt, resolve::color::RgbColor, dxpdf_docx_model::model::ParagraphProperties) {
    use resolve::color::{resolve_color, ColorContext, RgbColor};
    use resolve::fonts::effective_font;
    use resolve::properties::merge_paragraph_properties;

    let mut para_props = para.properties.clone();
    let mut run_defaults = resolved.doc_defaults_run.clone();

    // Derive default font from: doc defaults → theme minor font → spec fallback.
    let mut default_family = resolved
        .theme
        .as_ref()
        .map(|t| t.minor_font.latin.as_str())
        .filter(|s| !s.is_empty())
        .unwrap_or(SPEC_FALLBACK_FONT)
        .to_string();
    let mut default_size = resolved
        .doc_defaults_run
        .font_size
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE);
    let mut default_color = RgbColor::BLACK;

    // Look up the paragraph's style and merge its properties as base.
    if let Some(ref style_id) = para.style_id {
        if let Some(resolved_style) = resolved.styles.get(style_id) {
            merge_paragraph_properties(&mut para_props, &resolved_style.paragraph);
            run_defaults = resolved_style.run.clone();
        }
    }

    // Always merge doc defaults as the lowest-priority fallback.
    merge_paragraph_properties(&mut para_props, &resolved.doc_defaults_paragraph);

    // Style's run font overrides the document-level default.
    if let Some(f) = effective_font(&run_defaults.fonts) {
        default_family = f.to_string();
    }
    if let Some(fs) = run_defaults.font_size {
        default_size = Pt::from(fs);
    }
    if let Some(c) = run_defaults.color {
        default_color = resolve_color(c, ColorContext::Text);
    }

    (default_family, default_size, default_color, para_props)
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

    // §17.3.1.33: space after defaults to 0 when not specified in the cascade.
    // The doc defaults / style cascade should have already merged any intended
    // default (e.g., Normal style with after=200twip). If it's still None after
    // the full cascade, spec says 0.
    let space_after = props
        .spacing
        .and_then(|s| s.after)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);

    // §17.3.1.33: line spacing defaults to single (auto, 240 twips = 1.0x) per spec.
    // Auto line spacing is in 240ths of a line: 240 = single, 480 = double.
    let line_spacing = props
        .spacing
        .and_then(|s| s.line)
        .map(|ls| match ls {
            // §17.3.1.33: auto value is in 240ths of a line.
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
        drop_cap: None,
        borders: resolve_paragraph_borders(props),
    }
}

/// §17.3.1.24: resolve paragraph borders for rendering.
fn resolve_paragraph_borders(
    props: &dxpdf_docx_model::model::ParagraphProperties,
) -> Option<layout::paragraph::ParagraphBorderStyle> {
    use layout::paragraph::{BorderLine, ParagraphBorderStyle};

    let pbdr = props.borders.as_ref()?;

    let convert = |b: &dxpdf_docx_model::model::Border| -> BorderLine {
        BorderLine {
            width: Pt::from(b.width),
            color: resolve::color::resolve_color(b.color, resolve::color::ColorContext::Text),
            space: Pt::from(b.space),
        }
    };

    let style = ParagraphBorderStyle {
        top: pbdr.top.as_ref().map(convert),
        bottom: pbdr.bottom.as_ref().map(convert),
        left: pbdr.left.as_ref().map(convert),
        right: pbdr.right.as_ref().map(convert),
    };

    // Only return if at least one border is set.
    if style.top.is_some() || style.bottom.is_some() || style.left.is_some() || style.right.is_some() {
        Some(style)
    } else {
        None
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
