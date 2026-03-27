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
            let (df, ds, dc, _, rd) = resolve_paragraph_defaults(p, resolved);
            let mut frags = collect_fragments(&p.content, &df, ds, dc, None, measure, Some(&resolved.styles), Some(&rd));
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
                let (default_family, default_size, default_color, merged_props, run_defaults) =
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
                    Some(&resolved.styles),
                    Some(&run_defaults),
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

                let page_break_before = merged_props.page_break_before.unwrap_or(false);
                blocks.push(LayoutBlock::Paragraph { fragments, style, page_break_before });
            }
            Block::Table(t) => {
                // Use grid column count (not cell count) — cells may span multiple grid columns.
                let num_cols = if t.grid.is_empty() {
                    t.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
                } else {
                    t.grid.len()
                };
                let grid_cols: Vec<Pt> = t.grid.iter().map(|g| Pt::from(g.width)).collect();

                // §17.4.63: table width determines whether to scale columns.
                // Auto or absent: use natural grid column widths (no scaling).
                // Pct/fixed: scale to fill available width.
                use dxpdf_docx_model::model::TableMeasure;
                let is_auto_width = matches!(
                    t.properties.width,
                    None | Some(TableMeasure::Auto) | Some(TableMeasure::Nil)
                );
                let col_widths = if is_auto_width && !grid_cols.is_empty() {
                    grid_cols.clone()
                } else {
                    compute_column_widths(&grid_cols, num_cols, config.content_width())
                };

                // §17.7.6: resolve table style for borders and conditional formatting.
                // Look up the raw style definition for table style overrides (tblStylePr).
                let raw_table_style = t
                    .properties
                    .style_id
                    .as_ref()
                    .and_then(|sid| {
                        // The raw stylesheet is not in ResolvedDocument, but the resolved style
                        // carries the table properties with borders.
                        resolved.styles.get(sid)
                    });

                // §17.7.6: resolve conditional formatting for each cell.
                let style_overrides = raw_table_style
                    .map(|s| s.table_style_overrides.as_slice())
                    .unwrap_or(&[]);
                let tbl_look = t.properties.look.as_ref();
                let row_band_size = t.properties.style_row_band_size.unwrap_or(1);
                let col_band_size = t.properties.style_col_band_size.unwrap_or(1);
                let num_rows = t.rows.len();

                // §17.4.42: cell margins from table style.
                let style_cell_margins = raw_table_style
                    .and_then(|s| s.table.as_ref())
                    .and_then(|tp| tp.cell_margins);

                let rows: Vec<TableRowInput> = t
                    .rows
                    .iter()
                    .enumerate()
                    .map(|(row_idx, row)| {
                        let num_cols = row.cells.len();
                        let cells: Vec<TableCellInput> = row
                            .cells
                            .iter()
                            .enumerate()
                            .map(|(col_idx, cell)| {
                                // §17.7.6: resolve conditional formatting for this cell.
                                let cond = resolve::conditional::resolve_cell_conditional(
                                    row_idx, col_idx, num_rows, num_cols,
                                    tbl_look, style_overrides, row_band_size, col_band_size,
                                );

                                let cell_blocks: Vec<CellBlock> = cell
                                    .content
                                    .iter()
                                    .filter_map(|b| {
                                        if let Block::Paragraph(p) = b {
                                            // §17.7.2: table style paragraph properties
                                            // take priority over Normal for table cells.
                                            let mut table_para = p.clone();
                                            if let Some(ts) = raw_table_style {
                                                resolve::properties::merge_paragraph_properties(
                                                    &mut table_para.properties, &ts.paragraph,
                                                );
                                            }
                                            // §17.7.6: conditional paragraph overrides.
                                            if let Some(ref pp) = cond.paragraph_properties {
                                                resolve::properties::merge_paragraph_properties(
                                                    &mut table_para.properties, pp,
                                                );
                                            }

                                            let (df, mut ds, mut dc, mp, mut rd) =
                                                resolve_paragraph_defaults(&table_para, resolved);

                                            // §17.7.2: table style run properties override Normal.
                                            if let Some(ts) = raw_table_style {
                                                if let Some(fs) = ts.run.font_size {
                                                    ds = Pt::from(fs);
                                                    rd.font_size = Some(fs);
                                                }
                                            }

                                            // §17.7.6: conditional run property overrides.
                                            if let Some(ref rp) = cond.run_properties {
                                                resolve::properties::merge_run_properties(&mut rd, rp);
                                                if let Some(c) = rd.color {
                                                    dc = resolve::color::resolve_color(c, resolve::color::ColorContext::Text);
                                                }
                                            }

                                            let mut frags = collect_fragments(
                                                &p.content,
                                                &df,
                                                ds,
                                                dc,
                                                None,
                                                &measure,
                                                Some(&resolved.styles),
                                                Some(&rd),
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

                                // §17.4.42: cell margins cascade.
                                let cell_margins = cell
                                    .properties
                                    .margins
                                    .or(t.properties.cell_margins)
                                    .or(style_cell_margins)
                                    .map(|m| geometry::PtEdgeInsets::new(
                                        Pt::from(m.top),
                                        Pt::from(m.right),
                                        Pt::from(m.bottom),
                                        Pt::from(m.left),
                                    ))
                                    .unwrap_or(geometry::PtEdgeInsets::ZERO);

                                // §17.7.6: resolve cell shading.
                                // Priority: direct cell → conditional → none.
                                let shading = cell.properties.shading
                                    .map(|s| resolve::color::resolve_color(s.fill, resolve::color::ColorContext::Background))
                                    .or_else(|| {
                                        cond.cell_properties.as_ref()
                                            .and_then(|tcp| tcp.shading.as_ref())
                                            .map(|s| resolve::color::resolve_color(s.fill, resolve::color::ColorContext::Background))
                                    });

                                // §17.7.6: resolve per-cell borders from conditional formatting.
                                let cell_borders = cond.cell_properties.as_ref()
                                    .and_then(|tcp| tcp.borders.as_ref())
                                    .map(|tcb| {
                                        use layout::table::{CellBorderConfig, TableBorderLine};
                                        let convert_border = |b: &dxpdf_docx_model::model::Border| -> TableBorderLine {
                                            TableBorderLine {
                                                width: Pt::from(b.width),
                                                color: resolve::color::resolve_color(
                                                    b.color, resolve::color::ColorContext::Text,
                                                ),
                                            }
                                        };
                                        // §17.4.38: convert tcBorders to CellBorderConfig.
                                        // None = not mentioned (no override).
                                        // Some(None) = val="nil" (explicitly remove border).
                                        // Some(Some(line)) = specific border.
                                        let convert_cell_border = |b: &Option<dxpdf_docx_model::model::Border>| -> Option<Option<TableBorderLine>> {
                                            b.as_ref().map(|b| {
                                                if b.style == dxpdf_docx_model::model::BorderStyle::None {
                                                    None // val="nil"
                                                } else {
                                                    Some(convert_border(b))
                                                }
                                            })
                                        };
                                        CellBorderConfig {
                                            top: convert_cell_border(&tcb.top),
                                            bottom: convert_cell_border(&tcb.bottom),
                                            left: convert_cell_border(&tcb.left),
                                            right: convert_cell_border(&tcb.right),
                                        }
                                    });

                                TableCellInput {
                                    blocks: cell_blocks,
                                    margins: cell_margins,
                                    grid_span: cell.properties.grid_span.unwrap_or(1),
                                    shading,
                                    cell_borders,
                                }
                            })
                            .collect();

                        TableRowInput {
                            cells,
                            min_height: row.properties.height.map(|h| Pt::from(h.value)),
                        }
                    })
                    .collect();

                // §17.4.38: resolve table borders from direct properties or table style.
                let tbl_borders = t
                    .properties
                    .borders
                    .as_ref()
                    .or_else(|| {
                        raw_table_style
                            .and_then(|s| s.table.as_ref())
                            .and_then(|tp| tp.borders.as_ref())
                    });

                let border_config = tbl_borders.map(|b| {
                    use layout::table::{TableBorderConfig, TableBorderLine};
                    let convert = |border: &dxpdf_docx_model::model::Border| -> TableBorderLine {
                        TableBorderLine {
                            width: Pt::from(border.width),
                            color: resolve::color::resolve_color(
                                border.color,
                                resolve::color::ColorContext::Text,
                            ),
                        }
                    };
                    TableBorderConfig {
                        top: b.top.as_ref().map(&convert),
                        bottom: b.bottom.as_ref().map(&convert),
                        left: b.left.as_ref().map(&convert),
                        right: b.right.as_ref().map(&convert),
                        inside_h: b.inside_h.as_ref().map(&convert),
                        inside_v: b.inside_v.as_ref().map(&convert),
                    }
                });

                // §17.4.58: floating table positioning.
                let float_info = t.properties.positioning.as_ref().map(|pos| {
                    let table_width: Pt = col_widths.iter().copied().sum();
                    let right_gap = pos
                        .right_from_text
                        .map(Pt::from)
                        .unwrap_or(Pt::ZERO);
                    let bottom_gap = pos
                        .bottom_from_text
                        .map(Pt::from)
                        .unwrap_or(Pt::ZERO);
                    (table_width, right_gap, bottom_gap)
                });

                blocks.push(LayoutBlock::Table {
                    rows,
                    col_widths,
                    border_config,
                    float_info,
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
/// Returns (default_font_family, default_font_size, default_color, merged_paragraph_properties, run_defaults).
fn resolve_paragraph_defaults(
    para: &dxpdf_docx_model::model::Paragraph,
    resolved: &ResolvedDocument,
) -> (String, Pt, resolve::color::RgbColor, dxpdf_docx_model::model::ParagraphProperties, dxpdf_docx_model::model::RunProperties) {
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
    // §17.7.4.17: if no style is specified, use the default paragraph style.
    let effective_style_id = para
        .style_id
        .as_ref()
        .or(resolved.default_paragraph_style_id.as_ref());

    if let Some(style_id) = effective_style_id {
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

    (default_family, default_size, default_color, para_props, run_defaults)
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
        // §17.3.1.31: paragraph shading.
        shading: props.shading.as_ref().map(|s| {
            resolve::color::resolve_color(s.fill, resolve::color::ColorContext::Background)
        }),
        float_beside: None,
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
            // §17.3.4: w:sz is ST_EighthPointMeasure.
            width: Pt::from(b.width),
            color: resolve::color::resolve_color(b.color, resolve::color::ColorContext::Text),
            // §17.3.4: w:space is ST_PointMeasure (§17.18.68).
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
