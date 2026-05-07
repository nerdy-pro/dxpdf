//! PDF renderer for dxpdf — measure, layout, and paint pipeline.
//!
//! Takes a parsed `Document` from `dxpdf-docx` and produces PDF bytes.
//!
//! # Pipeline
//!
//! 1. **Resolve** — flatten style inheritance, split sections, extract images/fonts
//! 2. **Layout** — fit content into pages using constraint-based layout
//! 3. **Subset** *(optional, gated by `subset-fonts`)* — collect glyph usage
//!    and replace each typeface with a subsetted variant before paint
//! 4. **Paint** — emit draw commands to Skia PDF canvas (requires `skia-safe`)

pub mod dimension;
pub(crate) mod emf;
pub mod emoji;
pub mod error;
pub mod fonts;
pub mod geometry;
pub mod layout;
pub mod painter;
pub mod resolve;
pub mod skia_conv;
#[cfg(feature = "subset-fonts")]
pub mod subset;

use crate::model::Document;

use crate::model::Block;
use crate::render::layout::build::{
    build_section_blocks, default_line_height, BuildContext, BuildState,
};
use crate::render::layout::draw_command::LayoutedPage;
use crate::render::layout::header_footer::{render_headers_footers, HeaderFooterBlocks, PageRange};
use crate::render::layout::page::PageConfig;
use crate::render::layout::section::layout_section;
use crate::render::resolve::ResolvedDocument;

/// Render a parsed DOCX document to PDF bytes.
///
/// Estimate the cursor_y position from the last page's draw commands.
/// Used to determine where a continuous section should start on the page.
fn estimate_cursor_y(
    page: &layout::draw_command::LayoutedPage,
    config: &layout::page::PageConfig,
) -> dimension::Pt {
    let mut max_y = config.margins.top;
    for cmd in &page.commands {
        let bottom = match cmd {
            layout::draw_command::DrawCommand::Text {
                position,
                font_size,
                ..
            } => position.y + *font_size,
            layout::draw_command::DrawCommand::Image { rect, .. } => {
                rect.origin.y + rect.size.height
            }
            layout::draw_command::DrawCommand::Rect { rect, .. } => {
                rect.origin.y + rect.size.height
            }
            layout::draw_command::DrawCommand::Line { line, .. } => line.end.y,
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
    #[allow(unused_mut)] // mut required only when subset-fonts is enabled
    let mut registry = fonts::FontRegistry::build(
        font_mgr.clone(),
        &doc.embedded_fonts,
        &resolved.font_families,
    );
    let pages = layout_document(&resolved, &registry);

    #[cfg(feature = "subset-fonts")]
    {
        let usage = subset::collect(&pages, &registry);
        let report = subset::apply(usage, &mut registry);
        log::info!("font subset: {report}");
    }

    painter::render_to_pdf(&pages, &registry)
}

/// Resolve and lay out a document without painting to PDF.
/// Uses a real FontMgr for text measurement.
pub fn resolve_and_layout(doc: &Document) -> (ResolvedDocument, Vec<LayoutedPage>) {
    let font_mgr = skia_safe::FontMgr::new();
    let resolved = resolve::resolve(doc);
    let registry =
        fonts::FontRegistry::build(font_mgr, &doc.embedded_fonts, &resolved.font_families);
    let pages = layout_document(&resolved, &registry);
    (resolved, pages)
}

/// Lay out a resolved document using Skia font metrics resolved through
/// the supplied [`fonts::FontRegistry`].
pub fn layout_document(
    resolved: &ResolvedDocument,
    registry: &fonts::FontRegistry,
) -> Vec<LayoutedPage> {
    let measurer = layout::measurer::TextMeasurer::new(registry);
    let ctx = BuildContext {
        measurer: &measurer,
        resolved,
    };
    let mut state = BuildState::default();
    let dlh = default_line_height(&ctx);
    let mut all_pages = Vec::new();
    let mut all_endnotes = Vec::new();
    let mut last_config = PageConfig::default();
    // Per-section metadata for deferred header/footer rendering.
    // Carries the section's resolved slot sets, `<w:titlePg/>` flag,
    // and logical page number of the section's first page (§17.6.12);
    // the global `<w:evenAndOddHeaders/>` setting is read once below.
    struct SectionHfInfo<'a> {
        page_range: std::ops::Range<usize>,
        config: PageConfig,
        headers: &'a crate::render::resolve::header_footer::HeaderFooterSet<Vec<Block>>,
        footers: &'a crate::render::resolve::header_footer::HeaderFooterSet<Vec<Block>>,
        title_pg: bool,
        logical_page_base: usize,
    }
    let mut section_hf: Vec<SectionHfInfo> = Vec::new();
    // §17.6.12: logical PAGE numbering accumulates across sections,
    // resetting wherever a section sets `pgNumType.start`. Document
    // starts at logical 1 unless the first section overrides it.
    let mut next_logical: usize = 1;

    // §17.11.23: footnote separator indent from default paragraph style.
    let separator_indent = resolved
        .default_paragraph_style_id
        .as_ref()
        .and_then(|id| resolved.styles.get(id))
        .and_then(|s| s.paragraph.indentation)
        .and_then(|ind| ind.first_line)
        .map(|fl| match fl {
            crate::model::FirstLineIndent::FirstLine(v) => dimension::Pt::from(v),
            _ => dimension::Pt::ZERO,
        })
        .unwrap_or(dimension::Pt::ZERO);

    // §17.6.22: track continuation state for `Continuous` section breaks.
    let mut pending_continuation: Option<layout::section::ContinuationState> = None;

    // Phase 1: layout all sections to determine total page count.
    for (section_idx, section) in resolved.sections.iter().enumerate() {
        // §17.6.2: if header/footer content extends past the body margin,
        // push the body start down (or bottom up) so content doesn't overlap.
        let config = adjust_margins_for_header_footer(
            PageConfig::from_section(&section.properties),
            section,
            &ctx,
            &mut state,
            dlh,
        );

        state.page_config = config.clone();
        let built = build_section_blocks(section, &config, &ctx, &mut state);
        let measure_fn = |text: &str,
                          font: &layout::fragment::FontProps|
         -> (dimension::Pt, layout::fragment::TextMetrics) {
            measurer.measure(text, font)
        };

        // §17.6.22: continuous sections continue on the current page.
        let continuation =
            if section.properties.section_type == Some(crate::model::SectionType::Continuous) {
                pending_continuation.take()
            } else {
                pending_continuation = None;
                None
            };

        let mut pages = layout_section(
            &built.blocks,
            &config,
            Some(&measure_fn),
            separator_indent,
            dlh,
            continuation,
        );

        // Collect endnotes for rendering at document end.
        all_endnotes.extend(built.endnotes);

        last_config = config.clone();

        // Check if the NEXT section is continuous — if so, save the last page
        // as continuation state instead of appending it.
        // (Peek ahead by checking the section index.)
        let next_is_continuous = resolved.sections.get(section_idx + 1).is_some_and(|next| {
            next.properties.section_type == Some(crate::model::SectionType::Continuous)
        });

        if next_is_continuous && !pages.is_empty() {
            let last_page = pages.pop().unwrap();
            let cursor_y = estimate_cursor_y(&last_page, &last_config);
            pending_continuation = Some(layout::section::ContinuationState {
                page: last_page,
                cursor_y,
            });
        }

        let page_start = all_pages.len();
        all_pages.append(&mut pages);
        let logical_page_base = layout::header_footer::next_logical_page_base(
            next_logical,
            section.properties.page_number_type.as_ref(),
        );
        let pages_in_section = all_pages.len() - page_start;
        next_logical = logical_page_base + pages_in_section;
        section_hf.push(SectionHfInfo {
            page_range: page_start..all_pages.len(),
            config,
            headers: &section.headers,
            footers: &section.footers,
            title_pg: section.properties.title_page.unwrap_or(false),
            logical_page_base,
        });
    }

    // Phase 2: render headers/footers with correct NUMPAGES (total page count).
    let total_pages = all_pages.len();
    let even_and_odd = resolved.even_and_odd_headers;
    for info in &section_hf {
        state.page_config = info.config.clone();
        render_headers_footers(
            &mut all_pages[info.page_range.clone()],
            &info.config,
            &HeaderFooterBlocks {
                headers: info.headers,
                footers: info.footers,
                title_pg: info.title_pg,
                even_and_odd,
            },
            &ctx,
            &mut state,
            dlh,
            &PageRange {
                page_base: info.page_range.start,
                logical_page_base: info.logical_page_base,
                total_pages,
            },
        );
    }

    // Render endnotes on a new page at the end of the document.
    if !all_endnotes.is_empty() {
        let measure_fn = |text: &str,
                          font: &layout::fragment::FontProps|
         -> (dimension::Pt, layout::fragment::TextMetrics) {
            measurer.measure(text, font)
        };
        let mut endnote_page = LayoutedPage::new(last_config.page_size);
        let content_width = last_config.content_width();
        let constraints =
            layout::BoxConstraints::tight_width(content_width, dimension::Pt::INFINITY);
        let mut cursor_y = last_config.margins.top;

        // Separator line.
        let sep_width = content_width * 0.33;
        let sep_x = last_config.margins.left + separator_indent;
        endnote_page
            .commands
            .push(layout::draw_command::DrawCommand::Line {
                line: crate::render::geometry::PtLineSegment::new(
                    crate::render::geometry::PtOffset::new(sep_x, cursor_y),
                    crate::render::geometry::PtOffset::new(sep_x + sep_width, cursor_y),
                ),
                color: crate::render::resolve::color::RgbColor::BLACK,
                width: dimension::Pt::new(0.5),
            });
        cursor_y += dimension::Pt::new(4.0);

        for (_, frags, style) in &all_endnotes {
            let para = layout::paragraph::layout_paragraph(
                frags,
                &constraints,
                style,
                dlh,
                Some(&measure_fn),
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

/// §17.6.2: if header or footer content extends past the body margin,
/// adjust margins so body text starts after the header / ends before
/// the footer.
///
/// Each section can supply up to three header parts (default / first /
/// even) and three footer parts. To keep the body region consistent
/// across pages within a section, we compute the extent as the
/// **maximum** across every populated slot — the body starts low
/// enough to clear the tallest header, and ends high enough to clear
/// the tallest footer. Per-page slot variation in height is
/// uncommon in real documents (Word styles all three slots
/// identically by default), so this conservative allocation is
/// indistinguishable from per-page adjustment in the typical case
/// while keeping body geometry stable.
fn adjust_margins_for_header_footer(
    mut config: PageConfig,
    section: &crate::render::resolve::sections::ResolvedSection,
    ctx: &layout::build::BuildContext,
    state: &mut BuildState,
    default_line_height: dimension::Pt,
) -> PageConfig {
    let content_width = config.content_width();
    let header_slots = [
        section.headers.default.as_deref(),
        section.headers.first.as_deref(),
        section.headers.even.as_deref(),
    ];
    let footer_slots = [
        section.footers.default.as_deref(),
        section.footers.first.as_deref(),
        section.footers.even.as_deref(),
    ];

    let mut max_header_bottom = dimension::Pt::ZERO;
    for blocks in header_slots.iter().flatten() {
        let hf = layout::build::build_header_footer_content(blocks, ctx, state);
        let result =
            layout::section::stack_blocks(&hf.blocks, content_width, default_line_height, None);

        // §17.6.2: header_bottom is the greater of stacked block content and
        // wrapTopAndBottom floating image extent, both measured from the header
        // margin edge. wrapNone/behindDoc images are backgrounds and don't
        // push body content down.
        let blocks_bottom = config.header_margin + result.height;
        let floats_bottom = hf
            .floating_images
            .iter()
            .filter(|fi| fi.is_wrap_top_and_bottom())
            .map(|fi| {
                let y = match fi.y {
                    layout::section::FloatingImageY::Absolute(y) => y,
                    layout::section::FloatingImageY::RelativeToParagraph(off) => {
                        config.header_margin + off
                    }
                };
                y + fi.size.height
            })
            .fold(dimension::Pt::ZERO, |a, b| a.max(b));
        max_header_bottom = max_header_bottom.max(blocks_bottom.max(floats_bottom));
    }
    if max_header_bottom > config.margins.top {
        config.margins.top = max_header_bottom;
    }

    let mut max_footer_extent = dimension::Pt::ZERO;
    for blocks in footer_slots.iter().flatten() {
        let hf = layout::build::build_header_footer_content(blocks, ctx, state);
        let result =
            layout::section::stack_blocks(&hf.blocks, content_width, default_line_height, None);

        let blocks_extent = config.footer_margin + result.height;
        let floats_extent = hf
            .floating_images
            .iter()
            .filter(|fi| fi.is_wrap_top_and_bottom())
            .map(|fi| match fi.y {
                layout::section::FloatingImageY::Absolute(y) => config.page_size.height - y,
                layout::section::FloatingImageY::RelativeToParagraph(off) => {
                    config.footer_margin + off + fi.size.height
                }
            })
            .fold(dimension::Pt::ZERO, |a, b| a.max(b));
        max_footer_extent = max_footer_extent.max(blocks_extent.max(floats_extent));
    }
    if max_footer_extent > config.margins.bottom {
        config.margins.bottom = max_footer_extent;
    }

    config
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::*;
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
                content: vec![RunElement::Text(text.to_string())],
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
                    width: crate::model::dimension::Dimension::new(4680),
                },
                GridColumn {
                    width: crate::model::dimension::Dimension::new(4680),
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
                property_exceptions: None,
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
                width: Some(crate::model::dimension::Dimension::new(12240)),
                height: Some(crate::model::dimension::Dimension::new(15840)),
                orientation: None,
            }),
            ..Default::default()
        };

        let (_, pages) = resolve_and_layout(&doc);
        assert_eq!(pages[0].page_size.width.raw(), 612.0);
        assert_eq!(pages[0].page_size.height.raw(), 792.0);
    }
}
