//! Recursive tree-walk that converts a resolved document model into layout blocks.
//!
//! The document tree (Section → Block → Paragraph | Table → Cell → Block…) is
//! processed by recursive descent.  Each `Block::Table` recurses into its cells'
//! content, which may contain nested tables.

use std::collections::HashMap;

use dxpdf_docx_model::model::{
    self, Block, FirstLineIndent, LineSpacing, Paragraph, Table, TableCell,
};

use crate::dimension::Pt;
use crate::geometry;
use crate::layout::cell::CellBlock;
use crate::layout::fragment::{collect_fragments, FontProps, Fragment};
use crate::layout::measurer::TextMeasurer;
use crate::layout::page::PageConfig;
use crate::layout::paragraph::{
    BorderLine, DropCapInfo, LineSpacingRule, ParagraphBorderStyle, ParagraphStyle,
};
use crate::layout::section::LayoutBlock;
use crate::layout::table::{
    CellBorderConfig, CellBorderOverride, TableBorderConfig, TableBorderLine, TableBorderStyle,
    TableCellInput, TableRowInput, compute_column_widths, layout_table,
};
use crate::resolve::color::{ColorContext, RgbColor, resolve_color};
use crate::resolve::conditional::{CellConditionalFormatting, resolve_cell_conditional};
use crate::resolve::fonts::effective_font;
use crate::resolve::properties::{merge_paragraph_properties, merge_run_properties};
use crate::resolve::sections::ResolvedSection;
use crate::resolve::styles::ResolvedStyle;
use crate::resolve::ResolvedDocument;

/// §17.8.3.2: OOXML fallback font when no theme or doc defaults specify one.
const SPEC_FALLBACK_FONT: &str = "Times New Roman";
/// §17.3.2.14: default font size (10pt = 20 half-points per ECMA-376 §17.3.2.14).
const SPEC_DEFAULT_FONT_SIZE: Pt = Pt::new(10.0);

// ── Context ─────────────────────────────────────────────────────────────────

/// Immutable context threaded through the recursive tree walk.
pub struct BuildContext<'a> {
    pub measurer: &'a TextMeasurer,
    pub resolved: &'a ResolvedDocument,
}

impl BuildContext<'_> {
    fn media(&self) -> &HashMap<model::RelId, Vec<u8>> {
        &self.resolved.media
    }
}

// ── Public entry point ──────────────────────────────────────────────────────

/// Build layout blocks for one section by recursing into its block tree.
pub fn build_section_blocks(
    section: &ResolvedSection,
    config: &PageConfig,
    ctx: &BuildContext,
) -> Vec<LayoutBlock> {
    let mut pending_dropcap: Option<DropCapInfo> = None;
    section
        .blocks
        .iter()
        .filter_map(|block| build_block(block, config.content_width(), ctx, &mut pending_dropcap))
        .collect()
}

/// Collect fragments from header/footer blocks.
pub fn collect_fragments_from_blocks(
    blocks: &[Block],
    ctx: &BuildContext,
) -> Vec<Fragment> {
    blocks
        .iter()
        .filter_map(|block| {
            if let Block::Paragraph(p) = block {
                let (frags, _) = build_fragments(p, ctx, None, None);
                Some(frags)
            } else {
                None
            }
        })
        .flatten()
        .collect()
}

/// Default line height derived from document-level font settings.
pub fn default_line_height(ctx: &BuildContext) -> Pt {
    let family = doc_font_family(ctx);
    let size = doc_font_size(ctx);
    ctx.measurer.default_line_height(&family, size)
}

// ── Recursive block processing ──────────────────────────────────────────────

/// Recursively process a single model block into a layout block.
///
/// Returns `None` for drop cap paragraphs (consumed by the next paragraph)
/// and section breaks (already handled by resolve).
fn build_block(
    block: &Block,
    available_width: Pt,
    ctx: &BuildContext,
    pending_dropcap: &mut Option<DropCapInfo>,
) -> Option<LayoutBlock> {
    match block {
        Block::Paragraph(p) => build_paragraph_block(p, ctx, pending_dropcap),
        Block::Table(t) => {
            let built = build_table(t, available_width, ctx);
            Some(LayoutBlock::Table {
                rows: built.rows,
                col_widths: built.col_widths,
                border_config: built.border_config,
                float_info: built.float_info,
            })
        }
        Block::SectionBreak(_) => None,
    }
}

// ── Paragraph building ──────────────────────────────────────────────────────

/// Build a top-level paragraph into a layout block.
/// Handles drop cap detection (§17.3.1.11).
fn build_paragraph_block(
    p: &Paragraph,
    ctx: &BuildContext,
    pending_dropcap: &mut Option<DropCapInfo>,
) -> Option<LayoutBlock> {
    let (fragments, merged_props) = build_fragments(p, ctx, None, None);

    // §17.3.1.11: detect drop cap paragraph.
    let is_dropcap = merged_props
        .frame_properties
        .and_then(|fp| fp.drop_cap)
        .is_some_and(|dc| {
            matches!(dc, model::DropCap::Drop | model::DropCap::Margin)
        });

    if is_dropcap {
        let drop_cap_lines = merged_props
            .frame_properties
            .and_then(|fp| fp.lines)
            .unwrap_or(3);
        let width: Pt = fragments.iter().map(|f| f.width()).sum();
        let height: Pt = fragments.iter().map(|f| f.height()).fold(Pt::ZERO, Pt::max);
        let ascent: Pt = fragments
            .iter()
            .map(|f| match f {
                Fragment::Text { ascent, .. } => *ascent,
                _ => Pt::ZERO,
            })
            .fold(Pt::ZERO, Pt::max);
        let h_space = merged_props
            .frame_properties
            .and_then(|fp| fp.h_space)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO);
        *pending_dropcap = Some(DropCapInfo {
            fragments,
            lines: drop_cap_lines,
            ascent,
            h_space,
            width,
            height,
        });
        return None;
    }

    let mut style = paragraph_style_from_props(&merged_props);

    // Attach pending drop cap to this paragraph.
    if let Some(dc) = pending_dropcap.take() {
        style.drop_cap = Some(dc);
    }

    let page_break_before = merged_props.page_break_before.unwrap_or(false);
    Some(LayoutBlock::Paragraph {
        fragments,
        style,
        page_break_before,
    })
}

/// Build fragments and resolved paragraph properties for a paragraph.
///
/// Handles the full cascade: table style → conditional → paragraph style →
/// doc defaults → fragment collection → image/underline population.
fn build_fragments(
    para: &Paragraph,
    ctx: &BuildContext,
    table_style: Option<&ResolvedStyle>,
    cond: Option<&CellConditionalFormatting>,
) -> (Vec<Fragment>, model::ParagraphProperties) {
    // Clone paragraph for table style / conditional merge.
    let mut effective_para = para.clone();

    // §17.7.2: table style paragraph properties as base.
    if let Some(ts) = table_style {
        merge_paragraph_properties(&mut effective_para.properties, &ts.paragraph);
    }
    // §17.7.6: conditional paragraph overrides.
    if let Some(c) = cond {
        if let Some(ref pp) = c.paragraph_properties {
            merge_paragraph_properties(&mut effective_para.properties, pp);
        }
    }

    let (default_family, mut default_size, mut default_color, merged_props, mut run_defaults) =
        resolve_paragraph_defaults(&effective_para, ctx.resolved);

    // §17.7.2: table style run properties override Normal.
    if let Some(ts) = table_style {
        if let Some(fs) = ts.run.font_size {
            default_size = Pt::from(fs);
            run_defaults.font_size = Some(fs);
        }
    }

    // §17.7.6: conditional run property overrides.
    if let Some(c) = cond {
        if let Some(ref rp) = c.run_properties {
            merge_run_properties(&mut run_defaults, rp);
            if let Some(color) = run_defaults.color {
                default_color = resolve_color(color, ColorContext::Text);
            }
        }
    }

    let measure = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
        ctx.measurer.measure(text, font)
    };

    let mut fragments = collect_fragments(
        &para.content,
        &default_family,
        default_size,
        default_color,
        None,
        &measure,
        Some(&ctx.resolved.styles),
        Some(&run_defaults),
    );
    populate_image_data(&mut fragments, ctx.media());
    populate_underline_metrics(&mut fragments, ctx.measurer);

    (fragments, merged_props)
}

// ── Table building ──────────────────────────────────────────────────────────

/// Result of building a table from the model.
struct BuiltTable {
    rows: Vec<TableRowInput>,
    col_widths: Vec<Pt>,
    border_config: Option<TableBorderConfig>,
    float_info: Option<(Pt, Pt, Pt)>,
}

/// Recursively build a table: resolve styles, conditional formatting, and
/// recurse into each cell's content blocks.
fn build_table(t: &Table, available_width: Pt, ctx: &BuildContext) -> BuiltTable {
    // §17.4.14: grid column widths.
    let num_cols = if t.grid.is_empty() {
        t.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    } else {
        t.grid.len()
    };
    let grid_cols: Vec<Pt> = t.grid.iter().map(|g| Pt::from(g.width)).collect();

    // §17.4.63: table width determines whether to scale columns.
    let is_auto_width = matches!(
        t.properties.width,
        None | Some(model::TableMeasure::Auto) | Some(model::TableMeasure::Nil)
    );
    let col_widths = if is_auto_width && !grid_cols.is_empty() {
        grid_cols.clone()
    } else {
        compute_column_widths(&grid_cols, num_cols, available_width)
    };

    // §17.7.6: table style and conditional formatting context.
    let raw_table_style = t
        .properties
        .style_id
        .as_ref()
        .and_then(|sid| ctx.resolved.styles.get(sid));
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

    // Build rows by iterating cells and recursing into their content.
    let rows: Vec<TableRowInput> = t
        .rows
        .iter()
        .enumerate()
        .map(|(row_idx, row)| {
            let num_cells = row.cells.len();
            let cells: Vec<TableCellInput> = row
                .cells
                .iter()
                .enumerate()
                .map(|(col_idx, cell)| {
                    let cond = resolve_cell_conditional(
                        row_idx, col_idx, num_rows, num_cells,
                        tbl_look, style_overrides, row_band_size, col_band_size,
                    );

                    // Compute available width for nested content.
                    let span = cell.properties.grid_span.unwrap_or(1) as usize;
                    let mut grid_start = 0;
                    for ci in 0..col_idx {
                        grid_start += row.cells[ci].properties.grid_span.unwrap_or(1) as usize;
                    }
                    let cell_width: Pt = col_widths[grid_start..grid_start + span]
                        .iter()
                        .copied()
                        .sum();
                    let cell_margins_h = cell
                        .properties
                        .margins
                        .or(t.properties.cell_margins)
                        .or(style_cell_margins)
                        .map(|m| Pt::from(m.left) + Pt::from(m.right))
                        .unwrap_or(Pt::ZERO);
                    let inner_width = (cell_width - cell_margins_h).max(Pt::ZERO);

                    build_table_cell(
                        cell, &t.properties, raw_table_style, style_cell_margins,
                        &cond, inner_width, ctx,
                    )
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
    let border_config = tbl_borders.map(convert_table_border_config);

    // §17.4.58: floating table positioning.
    let float_info = t.properties.positioning.as_ref().map(|pos| {
        let table_width: Pt = col_widths.iter().copied().sum();
        let right_gap = pos.right_from_text.map(Pt::from).unwrap_or(Pt::ZERO);
        let bottom_gap = pos.bottom_from_text.map(Pt::from).unwrap_or(Pt::ZERO);
        (table_width, right_gap, bottom_gap)
    });

    BuiltTable { rows, col_widths, border_config, float_info }
}

/// Build a single table cell: resolve content blocks, margins, shading, borders.
fn build_table_cell(
    cell: &TableCell,
    table_props: &model::TableProperties,
    table_style: Option<&ResolvedStyle>,
    style_cell_margins: Option<dxpdf_docx_model::geometry::EdgeInsets<dxpdf_docx_model::dimension::Twips>>,
    cond: &CellConditionalFormatting,
    inner_width: Pt,
    ctx: &BuildContext,
) -> TableCellInput {
    // Recurse into cell content blocks.
    let cell_blocks = build_cell_blocks(&cell.content, table_style, cond, inner_width, ctx);

    // §17.4.42: cell margins cascade: cell → table → table style.
    let cell_margins = cell
        .properties
        .margins
        .or(table_props.cell_margins)
        .or(style_cell_margins)
        .map(|m| {
            geometry::PtEdgeInsets::new(
                Pt::from(m.top), Pt::from(m.right),
                Pt::from(m.bottom), Pt::from(m.left),
            )
        })
        .unwrap_or(geometry::PtEdgeInsets::ZERO);

    // §17.7.6: resolve cell shading.  Priority: direct → conditional → none.
    let shading = cell
        .properties
        .shading
        .map(|s| resolve_color(s.fill, ColorContext::Background))
        .or_else(|| {
            cond.cell_properties
                .as_ref()
                .and_then(|tcp| tcp.shading.as_ref())
                .map(|s| resolve_color(s.fill, ColorContext::Background))
        });

    // §17.7.6: per-cell borders from conditional formatting.
    let cell_borders = cond
        .cell_properties
        .as_ref()
        .and_then(|tcp| tcp.borders.as_ref())
        .map(|tcb| CellBorderConfig {
            top: convert_cell_border_override(&tcb.top),
            bottom: convert_cell_border_override(&tcb.bottom),
            left: convert_cell_border_override(&tcb.left),
            right: convert_cell_border_override(&tcb.right),
        });

    TableCellInput {
        blocks: cell_blocks,
        margins: cell_margins,
        grid_span: cell.properties.grid_span.unwrap_or(1),
        shading,
        cell_borders,
        vertical_merge: cell.properties.vertical_merge.map(|vm| match vm {
            model::VerticalMerge::Restart => crate::layout::table::VerticalMergeState::Restart,
            model::VerticalMerge::Continue => crate::layout::table::VerticalMergeState::Continue,
        }),
    }
}

/// Recursively build cell content blocks.
///
/// Paragraphs are resolved with table style + conditional overrides.
/// Nested tables recurse via `build_table()` → `layout_table()`.
fn build_cell_blocks(
    content: &[Block],
    table_style: Option<&ResolvedStyle>,
    cond: &CellConditionalFormatting,
    inner_width: Pt,
    ctx: &BuildContext,
) -> Vec<CellBlock> {
    let dlh = default_line_height(ctx);
    content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => {
                let (frags, merged_props) = build_fragments(p, ctx, table_style, Some(cond));
                Some(CellBlock::Paragraph {
                    fragments: frags,
                    style: paragraph_style_from_props(&merged_props),
                })
            }
            Block::Table(nested_t) => {
                let built = build_table(nested_t, inner_width, ctx);
                let result = layout_table(
                    &built.rows,
                    &built.col_widths,
                    &crate::layout::BoxConstraints::unbounded(),
                    dlh,
                    built.border_config.as_ref(),
                );
                Some(CellBlock::NestedTable {
                    commands: result.commands,
                    size: result.size,
                })
            }
            _ => None,
        })
        .collect()
}

// ── Paragraph property resolution ───────────────────────────────────────────

/// Resolve a paragraph's effective defaults.
/// Cascade: direct → style → doc defaults.
///
/// Returns (font_family, font_size, color, merged_paragraph_props, run_defaults).
fn resolve_paragraph_defaults(
    para: &Paragraph,
    resolved: &ResolvedDocument,
) -> (String, Pt, RgbColor, model::ParagraphProperties, model::RunProperties) {
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

    // Always merge doc defaults as lowest-priority fallback.
    merge_paragraph_properties(&mut para_props, &resolved.doc_defaults_paragraph);

    // Style's run font overrides document-level default.
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

// ── Conversion helpers ──────────────────────────────────────────────────────

fn doc_font_family(ctx: &BuildContext) -> String {
    ctx.resolved
        .theme
        .as_ref()
        .map(|t| t.minor_font.latin.as_str())
        .filter(|s| !s.is_empty())
        .unwrap_or(SPEC_FALLBACK_FONT)
        .to_string()
}

fn doc_font_size(ctx: &BuildContext) -> Pt {
    ctx.resolved
        .doc_defaults_run
        .font_size
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE)
}

/// Convert a model paragraph properties into a layout ParagraphStyle.
fn paragraph_style_from_props(props: &model::ParagraphProperties) -> ParagraphStyle {
    let indent_left = props.indentation.and_then(|i| i.start).map(Pt::from).unwrap_or(Pt::ZERO);
    let indent_right = props.indentation.and_then(|i| i.end).map(Pt::from).unwrap_or(Pt::ZERO);
    let indent_first_line = props
        .indentation
        .and_then(|i| i.first_line)
        .map(|fl| match fl {
            FirstLineIndent::FirstLine(v) => Pt::from(v),
            FirstLineIndent::Hanging(v) => -Pt::from(v),
            FirstLineIndent::None => Pt::ZERO,
        })
        .unwrap_or(Pt::ZERO);

    let space_before = props.spacing.and_then(|s| s.before).map(Pt::from).unwrap_or(Pt::ZERO);
    // §17.3.1.33: space after defaults to 0 when not specified in the cascade.
    let space_after = props.spacing.and_then(|s| s.after).map(Pt::from).unwrap_or(Pt::ZERO);

    // §17.3.1.33: line spacing defaults to single (auto, 240 twips = 1.0x).
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
        alignment: props.alignment.unwrap_or(model::Alignment::Start),
        space_before,
        space_after,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        drop_cap: None,
        borders: resolve_paragraph_borders(props),
        shading: props.shading.as_ref().map(|s| resolve_color(s.fill, ColorContext::Background)),
        float_beside: None,
    }
}

/// §17.3.1.24: resolve paragraph borders.
fn resolve_paragraph_borders(
    props: &model::ParagraphProperties,
) -> Option<ParagraphBorderStyle> {
    let pbdr = props.borders.as_ref()?;

    let convert = |b: &model::Border| -> BorderLine {
        BorderLine {
            width: Pt::from(b.width),
            color: resolve_color(b.color, ColorContext::Text),
            space: Pt::from(b.space),
        }
    };

    let style = ParagraphBorderStyle {
        top: pbdr.top.as_ref().map(convert),
        bottom: pbdr.bottom.as_ref().map(convert),
        left: pbdr.left.as_ref().map(convert),
        right: pbdr.right.as_ref().map(convert),
    };

    if style.top.is_some() || style.bottom.is_some() || style.left.is_some() || style.right.is_some() {
        Some(style)
    } else {
        None
    }
}

/// Convert a model `Border` to a layout `TableBorderLine`.
fn convert_model_border(b: &model::Border) -> TableBorderLine {
    TableBorderLine {
        width: Pt::from(b.width),
        color: resolve_color(b.color, ColorContext::Text),
        style: match b.style {
            model::BorderStyle::Double => TableBorderStyle::Double,
            _ => TableBorderStyle::Single,
        },
    }
}

/// Convert a model cell border to a `CellBorderOverride`.
fn convert_cell_border_override(b: &Option<model::Border>) -> Option<CellBorderOverride> {
    b.as_ref().map(|b| {
        if b.style == model::BorderStyle::None {
            CellBorderOverride::Nil
        } else {
            CellBorderOverride::Border(convert_model_border(b))
        }
    })
}

/// Convert model `TableBorders` to a layout `TableBorderConfig`.
fn convert_table_border_config(b: &model::TableBorders) -> TableBorderConfig {
    TableBorderConfig {
        top: b.top.as_ref().map(convert_model_border),
        bottom: b.bottom.as_ref().map(convert_model_border),
        left: b.left.as_ref().map(convert_model_border),
        right: b.right.as_ref().map(convert_model_border),
        inside_h: b.inside_h.as_ref().map(convert_model_border),
        inside_v: b.inside_v.as_ref().map(convert_model_border),
    }
}

/// Populate image data on Fragment::Image fragments from the media map.
fn populate_image_data(
    fragments: &mut [Fragment],
    media: &HashMap<model::RelId, Vec<u8>>,
) {
    for frag in fragments.iter_mut() {
        if let Fragment::Image { rel_id, image_data, .. } = frag {
            if image_data.is_none() {
                if let Some(bytes) = media.get(&model::RelId::new(rel_id.as_str())) {
                    *image_data = Some(bytes.as_slice().into());
                }
            }
        }
    }
}

/// Populate underline position/thickness from Skia font metrics.
fn populate_underline_metrics(fragments: &mut [Fragment], measurer: &TextMeasurer) {
    for frag in fragments.iter_mut() {
        if let Fragment::Text { font, .. } = frag {
            if font.underline {
                let (pos, thickness) = measurer.underline_metrics(font);
                font.underline_position = pos;
                font.underline_thickness = thickness;
            }
        }
    }
}
