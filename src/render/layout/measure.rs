//! Step 1 of the pipeline: measure text, resolve fonts, decode images.
//!
//! Transforms a `Document` (model) into a `MeasuredDocument` (tree with metrics).
//! No layout decisions are made here — only data enrichment.

use std::rc::Rc;

use crate::dimension::{HalfPoints, Pt};
use crate::model::*;

use super::context::LayoutConstraints;
use super::fragment::{collect_fragments_with_fields, Fragment};
use super::measurer::TextMeasurer;
use super::ImageCache;
use super::LayoutConfig;

/// Document-level defaults for layout.
pub struct DocDefaultsLayout {
    pub font_size: HalfPoints,
    pub font_family: Rc<str>,
    pub default_spacing: Spacing,
    pub default_cell_margins: CellMargins,
    pub table_cell_spacing: Spacing,
    pub default_table_borders: TableBorders,
    pub default_header: Option<HeaderFooter>,
    pub default_footer: Option<HeaderFooter>,
    pub numbering: NumberingMap,
    /// Pre-computed line height of an empty paragraph using the default font.
    pub default_line_height: Pt,
}

impl DocDefaultsLayout {
    pub fn from_document(doc: &Document, measurer: &TextMeasurer) -> Self {
        let font_size = doc.default_font_size;
        let font_family = Rc::clone(&doc.default_font_family);
        let default_size_pt = Pt::from(font_size);
        let default_line_height = measurer
            .font(&font_family, default_size_pt, false, false)
            .metrics()
            .line_height;

        // Resolve default header/footer from the first section that defines one.
        let mut default_header = None;
        let mut default_footer = None;
        for sect in doc.sections() {
            if default_header.is_none() {
                default_header = sect.header.clone();
            }
            if default_footer.is_none() {
                default_footer = sect.footer.clone();
            }
            if default_header.is_some() && default_footer.is_some() {
                break;
            }
        }

        Self {
            font_size,
            font_family,
            default_spacing: doc.default_spacing,
            default_cell_margins: doc.default_cell_margins,
            table_cell_spacing: doc.table_cell_spacing,
            default_table_borders: doc.default_table_borders,
            default_header,
            default_footer,
            numbering: doc.numbering.clone(),
            default_line_height,
        }
    }
}

/// A document tree enriched with text metrics and decoded images.
/// Produced by the measure step, consumed by the layout step.
pub struct MeasuredDocument {
    pub blocks: Vec<MeasuredBlock>,
    /// Section configs derived from inline section breaks + final section.
    pub(crate) section_configs: Vec<LayoutConfig>,
    pub(crate) doc_defaults: DocDefaultsLayout,
    pub(crate) default_tab_stop: Pt,
    pub(crate) image_cache: ImageCache,
}

/// Compute column widths by scaling grid columns to fit the available width.
/// If `grid_cols` is empty, distributes `available_width` equally across `num_cols`.
pub(super) fn compute_column_widths(
    grid_cols: &[crate::dimension::Twips],
    num_cols: usize,
    available_width: Pt,
) -> Vec<Pt> {
    if !grid_cols.is_empty() {
        let grid_total: Pt = grid_cols.iter().map(|w| Pt::from(*w)).sum();
        let scale = if grid_total > Pt::ZERO {
            available_width / grid_total
        } else {
            1.0
        };
        grid_cols.iter().map(|w| Pt::from(*w) * scale).collect()
    } else if num_cols > 0 {
        vec![available_width / num_cols as f32; num_cols]
    } else {
        vec![]
    }
}

/// A measured block-level element.
pub enum MeasuredBlock {
    Paragraph(Box<MeasuredParagraph>),
    Table(Box<MeasuredTable>),
}

/// A paragraph with pre-measured fragments.
pub struct MeasuredParagraph {
    pub fragments: Vec<Fragment>,
    pub properties: ParagraphProperties,
    pub floats: Vec<FloatingImage>,
    pub section_properties: Option<SectionProperties>,
    /// Resolved list label: (label_text, indent_left, indent_hanging) in Twips.
    pub list_label: Option<(String, crate::dimension::Twips, crate::dimension::Twips)>,
}

/// A table with pre-measured cells.
pub struct MeasuredTable {
    pub rows: Vec<MeasuredTableRow>,
    pub grid_cols: Vec<crate::dimension::Twips>,
    pub default_cell_margins: Option<CellMargins>,
    pub cell_spacing: Option<Spacing>,
    pub borders: Option<TableBorders>,
}

/// A table row with pre-measured cells.
pub struct MeasuredTableRow {
    pub cells: Vec<MeasuredTableCell>,
    pub height: Option<crate::dimension::Twips>,
}

/// A table cell with pre-measured paragraph blocks.
pub struct MeasuredTableCell {
    pub blocks: Vec<MeasuredBlock>,
    pub width: Option<crate::dimension::Twips>,
    pub grid_span: u32,
    pub vertical_merge: Option<VerticalMerge>,
    pub cell_margins: Option<CellMargins>,
    pub cell_borders: Option<CellBorders>,
    pub shading: Option<Color>,
}

impl MeasuredTableCell {
    pub fn is_vmerge_continue(&self) -> bool {
        self.vertical_merge == Some(VerticalMerge::Continue)
    }
}

/// Measure a document: resolve fonts, measure text, decode images.
/// This is Step 1 of the pipeline.
pub fn measure(doc: &Document, font_mgr: &skia_safe::FontMgr) -> MeasuredDocument {
    let measurer = TextMeasurer::new(font_mgr.clone());
    let doc_defaults = DocDefaultsLayout::from_document(doc, &measurer);
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

    // Derive page constraints from the first section config (or defaults)
    let initial_config = section_configs.first().copied().unwrap_or_default();
    let mut constraints = LayoutConstraints::for_page(&initial_config);

    // List counter state for numbering
    let mut list_counters: std::collections::HashMap<(u32, u32), u32> =
        std::collections::HashMap::new();

    let blocks = measure_blocks(
        &doc.blocks,
        &mut constraints,
        &doc_defaults,
        &measurer,
        &image_cache,
        &mut list_counters,
    );

    MeasuredDocument {
        blocks,
        section_configs,
        doc_defaults,
        default_tab_stop,
        image_cache,
    }
}

fn measure_blocks(
    blocks: &[Block],
    constraints: &mut LayoutConstraints,
    doc_defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    image_cache: &ImageCache,
    list_counters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> Vec<MeasuredBlock> {
    let mut result = Vec::with_capacity(blocks.len());
    for block in blocks {
        let measured = match block {
            Block::Paragraph(p) => MeasuredBlock::Paragraph(Box::new(measure_paragraph(
                p,
                constraints,
                doc_defaults,
                measurer,
                image_cache,
                list_counters,
            ))),
            Block::Table(t) => MeasuredBlock::Table(Box::new(measure_table(
                t,
                constraints,
                doc_defaults,
                measurer,
                image_cache,
                list_counters,
            ))),
        };
        result.push(measured);
    }
    result
}

fn measure_paragraph(
    para: &Paragraph,
    constraints: &LayoutConstraints,
    doc_defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    image_cache: &ImageCache,
    list_counters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> MeasuredParagraph {
    let fragments = collect_fragments_with_fields(
        &para.runs,
        constraints,
        doc_defaults,
        measurer,
        None,
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
    constraints: &mut LayoutConstraints,
    doc_defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
    image_cache: &ImageCache,
    list_counters: &mut std::collections::HashMap<(u32, u32), u32>,
) -> MeasuredTable {
    let content_width = constraints.available_width();
    let available_height = constraints.available_height();
    let num_cols = table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
    let col_widths = compute_column_widths(&table.grid_cols, num_cols, content_width);

    let doc_cell_margins = doc_defaults.default_cell_margins;

    let mut rows = Vec::with_capacity(table.rows.len());
    for row in &table.rows {
        let mut grid_col_idx = 0;
        let mut cells = Vec::with_capacity(row.cells.len());
        for cell in &row.cells {
            // Compute cell width from grid columns
            let span = cell.grid_span.max(1) as usize;
            let cell_width: Pt = (grid_col_idx..grid_col_idx + span)
                .map(|i| col_widths.get(i).copied().unwrap_or(Pt::new(72.0)))
                .sum();
            grid_col_idx += span;

            // Compute cell content constraints
            let margins = cell
                .cell_margins
                .or(table.default_cell_margins)
                .unwrap_or(doc_cell_margins);
            let pad_left = Pt::from(margins.left);
            let pad_right = Pt::from(margins.right);
            let cell_content_width = (cell_width - pad_left - pad_right).max(Pt::new(1.0));

            constraints.push_cell(Pt::ZERO, cell_content_width, available_height);
            let blocks = measure_blocks(
                &cell.blocks,
                constraints,
                doc_defaults,
                measurer,
                image_cache,
                list_counters,
            );
            constraints.pop();

            cells.push(MeasuredTableCell {
                blocks,
                width: cell.width,
                grid_span: cell.grid_span,
                vertical_merge: cell.vertical_merge,
                cell_margins: cell.cell_margins,
                cell_borders: cell.cell_borders,
                shading: cell.shading,
            });
        }
        rows.push(MeasuredTableRow {
            cells,
            height: row.height,
        });
    }

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
            let roman = to_roman(*counter);
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

pub(super) fn to_roman(mut n: u32) -> String {
    let table = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut result = String::new();
    for &(value, numeral) in &table {
        while n >= value {
            result.push_str(numeral);
            n -= value;
        }
    }
    result
}
