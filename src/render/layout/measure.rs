//! Step 1 of the pipeline: measure text, resolve fonts, decode images.
//!
//! Transforms a `Document` (model) into a `MeasuredDocument` (tree with metrics).
//! No layout decisions are made here — only data enrichment.

use crate::dimension::Pt;
use crate::model::*;

use super::fragment::{collect_fragments, DocDefaultsLayout, Fragment};
use super::measurer::TextMeasurer;
use super::ImageCache;
use super::LayoutConfig;

/// A document tree enriched with text metrics and decoded images.
/// Produced by the measure step, consumed by the layout step.
pub(crate) struct MeasuredDocument {
    pub blocks: Vec<MeasuredBlock>,
    /// Section configs derived from inline section breaks + final section.
    pub section_configs: Vec<LayoutConfig>,
    pub doc_defaults: DocDefaultsLayout,
    pub default_tab_stop: Pt,
    pub image_cache: ImageCache,
    pub header: Option<Vec<MeasuredBlock>>,
    pub footer: Option<Vec<MeasuredBlock>>,
}

/// A measured block-level element.
pub(crate) enum MeasuredBlock {
    Paragraph(MeasuredParagraph),
    Table(MeasuredTable),
}

/// A paragraph with pre-measured fragments.
pub(crate) struct MeasuredParagraph {
    pub fragments: Vec<Fragment>,
    pub properties: ParagraphProperties,
    pub floats: Vec<FloatingImage>,
    pub section_properties: Option<SectionProperties>,
    /// Resolved list label: (label_text, indent_left, indent_hanging) in Twips.
    pub list_label: Option<(String, crate::dimension::Twips, crate::dimension::Twips)>,
}

/// A table with pre-measured cells.
pub(crate) struct MeasuredTable {
    pub rows: Vec<MeasuredTableRow>,
    pub grid_cols: Vec<crate::dimension::Twips>,
    pub default_cell_margins: Option<CellMargins>,
    pub cell_spacing: Option<Spacing>,
    pub borders: Option<TableBorders>,
}

/// A table row with pre-measured cells.
pub(crate) struct MeasuredTableRow {
    pub cells: Vec<MeasuredTableCell>,
    pub height: Option<crate::dimension::Twips>,
}

/// A table cell with pre-measured paragraph blocks.
pub(crate) struct MeasuredTableCell {
    pub blocks: Vec<MeasuredBlock>,
    pub width: Option<crate::dimension::Twips>,
    pub grid_span: u32,
    pub vertical_merge: Option<VerticalMerge>,
    pub cell_margins: Option<CellMargins>,
    pub cell_borders: Option<CellBorders>,
    pub shading: Option<Color>,
}

/// Measure a document: resolve fonts, measure text, decode images.
/// This is Step 1 of the pipeline.
pub(crate) fn measure(doc: &Document, font_mgr: &skia_safe::FontMgr) -> MeasuredDocument {
    let measurer = TextMeasurer::new(font_mgr.clone());
    let doc_defaults = DocDefaultsLayout::from_document(doc);
    let image_cache = ImageCache::new(&doc.images);
    let default_tab_stop = Pt::from(doc.default_tab_stop);

    // Build section configs from inline section breaks + final section
    let mut section_configs: Vec<LayoutConfig> = Vec::new();
    for block in &doc.blocks {
        if let Block::Paragraph(p) = block {
            if let Some(ref sect) = p.section_properties {
                section_configs.push(LayoutConfig::from_section(sect));
            }
        }
    }
    if let Some(ref sect) = doc.final_section {
        section_configs.push(LayoutConfig::from_section(sect));
    }

    // List counter state for numbering
    let mut list_counters: std::collections::HashMap<(u32, u32), u32> =
        std::collections::HashMap::new();

    let blocks = measure_blocks(
        &doc.blocks,
        &doc_defaults,
        &measurer,
        &image_cache,
        &mut list_counters,
    );

    let header = doc.default_header.as_ref().map(|hf| {
        measure_blocks(
            &hf.blocks,
            &doc_defaults,
            &measurer,
            &image_cache,
            &mut list_counters,
        )
    });

    let footer = doc.default_footer.as_ref().map(|hf| {
        measure_blocks(
            &hf.blocks,
            &doc_defaults,
            &measurer,
            &image_cache,
            &mut list_counters,
        )
    });

    MeasuredDocument {
        blocks,
        section_configs,
        doc_defaults,
        default_tab_stop,
        image_cache,
        header,
        footer,
    }
}

fn measure_blocks(
    blocks: &[Block],
    doc_defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    image_cache: &ImageCache,
    list_c ounters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> Vec<MeasuredBlock> {
    blocks
        .iter()
        .map(|block| match block {
            Block::Paragraph(p) => MeasuredBlock::Paragraph(measure_paragraph(
                p,
                doc_defaults,
                measurer,
                image_cache,
                list_counters,
            )),
            Block::Table(t) => MeasuredBlock::Table(measure_table(
                t,
                doc_defaults,
                measurer,
                image_cache,
                list_counters,
            )),
        })
        .collect()
}

fn measure_paragraph(
    para: &Paragraph,
    doc_defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    image_cache: &ImageCache,
    list_counters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> MeasuredParagraph {
    // Use a generous content width for initial measurement — layout will re-measure
    // with actual available width if needed (e.g., after float adjustment).
    // For now, use a large value so no text is prematurely truncated.
    let content_width = Pt::new(10000.0);
    let content_height = Pt::new(10000.0);

    let fragments = collect_fragments(
        &para.runs,
        content_width,
        content_height,
        doc_defaults,
        measurer,
        image_cache,
    );

    // Resolve list label
    let list_label = para
        .properties
        .list_ref
        .as_ref()
        .and_then(|lr| resolve_list_label(lr, doc_defaults, list_counters));

    MeasuredParagraph {
        fragments,
        properties: para.properties.clone(),
        floats: para.floats.clone(),
        section_properties: para.section_properties.clone(),
        list_label,
    }
}

fn measure_table(
    table: &Table,
    doc_defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    image_cache: &ImageCache,
    list_counters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> MeasuredTable {
    let rows = table
        .rows
        .iter()
        .map(|row| MeasuredTableRow {
            cells: row
                .cells
                .iter()
                .map(|cell| MeasuredTableCell {
                    blocks: measure_blocks(
                        &cell.blocks,
                        doc_defaults,
                        measurer,
                        image_cache,
                        list_counters,
                    ),
                    width: cell.width,
                    grid_span: cell.grid_span,
                    vertical_merge: cell.vertical_merge,
                    cell_margins: cell.cell_margins,
                    cell_borders: cell.cell_borders,
                    shading: cell.shading,
                })
                .collect(),
            height: row.height,
        })
        .collect();

    MeasuredTable {
        rows,
        grid_cols: table.grid_cols.clone(),
        default_cell_margins: table.default_cell_margins,
        cell_spacing: table.cell_spacing,
        borders: table.borders,
    }
}

/// Resolve list label text and indentation.
/// Moved from Layouter to the measure step since it only needs numbering definitions.
fn resolve_list_label(
    list_ref: &ListRef,
    doc_defaults: &DocDefaultsLayout,
    list_counters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> Option<(String, crate::dimension::Twips, crate::dimension::Twips)> {
    let def = doc_defaults.numbering.get(&list_ref.num_id)?;
    let level = def.levels.get(list_ref.level as usize)?;

    let label = match &level.format {
        NumberFormat::Bullet(ch) => ch.clone(),
        NumberFormat::Decimal => {
            let counter = list_counters
                .entry((list_ref.num_id, list_ref.level))
                .or_insert(level.start.saturating_sub(1));
            *counter += 1;
            level.level_text.replace("%1", &counter.to_string())
        }
        NumberFormat::LowerLetter => {
            let counter = list_counters
                .entry((list_ref.num_id, list_ref.level))
                .or_insert(level.start.saturating_sub(1));
            *counter += 1;
            let ch = (b'a' + ((*counter - 1) % 26) as u8) as char;
            level.level_text.replace("%1", &ch.to_string())
        }
        NumberFormat::UpperLetter => {
            let counter = list_counters
                .entry((list_ref.num_id, list_ref.level))
                .or_insert(level.start.saturating_sub(1));
            *counter += 1;
            let ch = (b'A' + ((*counter - 1) % 26) as u8) as char;
            level.level_text.replace("%1", &ch.to_string())
        }
        NumberFormat::LowerRoman | NumberFormat::UpperRoman => {
            let counter = list_counters
                .entry((list_ref.num_id, list_ref.level))
                .or_insert(level.start.saturating_sub(1));
            *counter += 1;
            let roman = super::header_footer::to_roman(*counter);
            let roman = if matches!(level.format, NumberFormat::LowerRoman) {
                roman.to_lowercase()
            } else {
                roman
            };
            level.level_text.replace("%1", &roman)
        }
    };

    Some((label, level.indent_left, level.indent_hanging))
}
