use crate::model::{self, Block, Table, TableCell};
use crate::render::dimension::Pt;
use crate::render::geometry;
use crate::render::layout::paragraph::DropCapInfo;
use crate::render::layout::section::LayoutBlock;
use crate::render::layout::table::{
    compute_column_widths, CellBorderConfig, CellBorderOverride, TableCellInput, TableRowInput,
};
use crate::render::resolve::color::{resolve_color, ColorContext};
use crate::render::resolve::conditional::{
    resolve_cell_conditional, CellConditionalFormatting, CellGridPosition,
};
use crate::render::resolve::styles::ResolvedStyle;

use super::block::build_paragraph_block;
use super::convert::{
    convert_cell_border_override, convert_table_border_config, merge_table_borders,
    split_oversized_fragments,
};
use super::{BuildContext, BuildState};

/// Result of building a table from the model.
pub(super) struct BuiltTable {
    pub(super) rows: Vec<TableRowInput>,
    pub(super) col_widths: Vec<Pt>,
    pub(super) border_config: Option<crate::render::layout::table::TableBorderConfig>,
    /// §17.4.51: table indentation from left margin.
    pub(super) indent: Pt,
    /// §17.4.28: table horizontal alignment (left/center/right).
    pub(super) alignment: Option<model::Alignment>,
    pub(super) float_info: Option<super::super::section::TableFloatInfo>,
}

/// Recursively build a table: resolve styles, conditional formatting, and
/// recurse into each cell's content blocks.
pub(super) fn build_table(
    t: &Table,
    available_width: Pt,
    ctx: &BuildContext,
    state: &mut BuildState,
) -> BuiltTable {
    // §17.4.14: grid column widths.
    let num_cols = if t.grid.is_empty() {
        t.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    } else {
        t.grid.len()
    };
    let grid_cols: Vec<Pt> = t.grid.iter().map(|g| Pt::from(g.width)).collect();

    // §17.7.6: table style for conditional formatting, borders, cell margins.
    let raw_table_style = t
        .properties
        .style_id
        .as_ref()
        .and_then(|sid| ctx.resolved.styles.get(sid));

    // §17.4.42: default cell margins from table style cascade.
    let style_cell_margins = raw_table_style
        .and_then(|s| s.table.as_ref())
        .and_then(|tp| tp.cell_margins);
    // Per-edge merge: direct tblCellMar overrides style per-edge, with
    // unspecified edges (value 0) falling back to the style's value.
    // Word merges per-edge rather than replacing the entire set.
    let default_cell_margins = match (t.properties.cell_margins, style_cell_margins) {
        (Some(direct), Some(style)) => {
            use crate::model::geometry::EdgeInsets;
            Some(EdgeInsets {
                top: if direct.top.raw() != 0 {
                    direct.top
                } else {
                    style.top
                },
                bottom: if direct.bottom.raw() != 0 {
                    direct.bottom
                } else {
                    style.bottom
                },
                left: if direct.left.raw() != 0 {
                    direct.left
                } else {
                    style.left
                },
                right: if direct.right.raw() != 0 {
                    direct.right
                } else {
                    style.right
                },
            })
        }
        (Some(direct), None) => Some(direct),
        (None, Some(style)) => Some(style),
        (None, None) => None,
    };

    // §17.4.63: resolve table width from tblW.
    let is_auto_width = matches!(
        t.properties.width,
        None | Some(model::TableMeasure::Auto) | Some(model::TableMeasure::Nil)
    );
    let cell_margins_h = default_cell_margins
        .map(|m| Pt::from(m.left) + Pt::from(m.right))
        .unwrap_or(Pt::ZERO);
    let target_width = match t.properties.width {
        Some(model::TableMeasure::Pct(pct)) => {
            // §17.4.63: percentage in fiftieths of a percent.
            // 5000 = 100%. Full-width tables extend by cell margins so
            // cell content aligns with paragraph text at the margins.
            let ratio = pct.raw() as f32 / 5000.0;
            let base = if pct.raw() >= 5000 {
                available_width + cell_margins_h
            } else {
                available_width
            };
            base * ratio
        }
        Some(model::TableMeasure::Twips(tw)) => Pt::from(tw),
        _ => available_width, // auto/nil: use grid cols or available width
    };
    // §17.4.53: fixed layout uses exact grid column widths — no scaling.
    // Auto layout scales columns to fit the target width.
    let is_fixed = t.properties.layout == Some(model::TableLayout::Fixed);
    let col_widths = if (is_auto_width || is_fixed) && !grid_cols.is_empty() {
        grid_cols.clone()
    } else {
        compute_column_widths(&grid_cols, num_cols, target_width)
    };
    let style_overrides = raw_table_style
        .map(|s| s.table_style_overrides.as_slice())
        .unwrap_or(&[]);
    let tbl_look = t.properties.look.as_ref();
    let row_band_size = t.properties.style_row_band_size.unwrap_or(1);
    let col_band_size = t.properties.style_col_band_size.unwrap_or(1);
    let num_rows = t.rows.len();

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
                        &CellGridPosition {
                            row_idx,
                            col_idx,
                            num_rows,
                            num_cols: num_cells,
                            row_band_size,
                            col_band_size,
                        },
                        tbl_look,
                        style_overrides,
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
                        .or(default_cell_margins)
                        .map(|m| Pt::from(m.left) + Pt::from(m.right))
                        .unwrap_or(Pt::ZERO);
                    let inner_width = (cell_width - cell_margins_h).max(Pt::ZERO);

                    build_table_cell(
                        cell,
                        raw_table_style,
                        default_cell_margins,
                        &cond,
                        inner_width,
                        ctx,
                        state,
                    )
                })
                .collect();

            TableRowInput {
                cells,
                height_rule: row.properties.height.map(|h| {
                    use crate::model::HeightRule;
                    use crate::render::layout::table::RowHeightRule;
                    match h.rule {
                        HeightRule::Exact => RowHeightRule::Exact(Pt::from(h.value)),
                        _ => RowHeightRule::AtLeast(Pt::from(h.value)),
                    }
                }),
                is_header: row.properties.is_header,
                cant_split: row.properties.cant_split,
            }
        })
        .collect();

    // §17.4.38: resolve table borders — merge direct properties over table style.
    // Direct tblBorders may specify only a subset of edges (e.g. insideH=none);
    // unspecified edges inherit from the table style.
    let style_borders = raw_table_style
        .and_then(|s| s.table.as_ref())
        .and_then(|tp| tp.borders.as_ref());
    let tbl_borders = match (t.properties.borders.as_ref(), style_borders) {
        (Some(direct), Some(style)) => Some(merge_table_borders(direct, style)),
        (Some(direct), None) => Some(*direct),
        (None, Some(style)) => Some(*style),
        (None, None) => None,
    };
    let border_config = tbl_borders.as_ref().map(convert_table_border_config);

    // §17.4.58: floating table positioning.
    let float_info = t.properties.positioning.as_ref().map(|pos| {
        super::super::section::TableFloatInfo {
            right_gap: pos.right_from_text.map(Pt::from).unwrap_or(Pt::ZERO),
            bottom_gap: pos.bottom_from_text.map(Pt::from).unwrap_or(Pt::ZERO),
            x_align: pos.x_align,
            // §17.4.59: tblpY — absolute Y offset from the vertical anchor.
            y_offset: pos.y.map(Pt::from).unwrap_or(Pt::ZERO),
            // §17.4.58: default vertical anchor is "text".
            vert_anchor: pos.vert_anchor.unwrap_or(crate::model::TableAnchor::Text),
        }
    });

    // §17.4.51: table indentation from left margin.
    // For full-width left-aligned tables, MS Word shifts the table left
    // by the default cell margin so cell content aligns with paragraph text.
    let is_full_width = matches!(
        t.properties.width,
        Some(model::TableMeasure::Pct(pct)) if pct.raw() >= 5000
    );
    let is_left_aligned = !matches!(
        t.properties.alignment,
        Some(model::Alignment::Center) | Some(model::Alignment::End)
    );
    let indent = match t.properties.indent {
        Some(model::TableMeasure::Twips(tw)) => Pt::from(tw),
        _ if is_full_width && is_left_aligned => -default_cell_margins
            .map(|m| Pt::from(m.left))
            .unwrap_or(Pt::ZERO),
        _ => Pt::ZERO,
    };

    BuiltTable {
        rows,
        col_widths,
        border_config,
        indent,
        alignment: t.properties.alignment,
        float_info,
    }
}

/// Build a single table cell: resolve content blocks, margins, shading, borders.
fn build_table_cell(
    cell: &TableCell,
    table_style: Option<&ResolvedStyle>,
    style_cell_margins: Option<crate::model::geometry::EdgeInsets<crate::model::dimension::Twips>>,
    cond: &CellConditionalFormatting,
    inner_width: Pt,
    ctx: &BuildContext,
    state: &mut BuildState,
) -> TableCellInput {
    // §17.4.42: cell margins cascade: cell-level tcMar → pre-merged table default.
    let cell_margins = cell
        .properties
        .margins
        .or(style_cell_margins)
        .map(|m| {
            geometry::PtEdgeInsets::new(
                Pt::from(m.top),
                Pt::from(m.right),
                Pt::from(m.bottom),
                Pt::from(m.left),
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

    // §17.4.66: cell borders cascade — direct cell borders (highest priority)
    // → conditional formatting → table-level borders (resolved in layout).
    let cond_borders = cond
        .cell_properties
        .as_ref()
        .and_then(|tcp| tcp.borders.as_ref());
    let direct_borders = cell.properties.borders.as_ref();

    let cell_borders = match (direct_borders, cond_borders) {
        (Some(db), _) => {
            // Direct cell borders: highest priority.  Fall through to
            // conditional for edges not specified directly.
            Some(CellBorderConfig {
                top: convert_cell_border_override(&db.top)
                    .or_else(|| cond_borders.and_then(|cb| convert_cell_border_override(&cb.top))),
                bottom: convert_cell_border_override(&db.bottom).or_else(|| {
                    cond_borders.and_then(|cb| convert_cell_border_override(&cb.bottom))
                }),
                left: convert_cell_border_override(&db.left)
                    .or_else(|| cond_borders.and_then(|cb| convert_cell_border_override(&cb.left))),
                right: convert_cell_border_override(&db.right).or_else(|| {
                    cond_borders.and_then(|cb| convert_cell_border_override(&cb.right))
                }),
            })
        }
        (None, Some(cb)) => Some(CellBorderConfig {
            top: convert_cell_border_override(&cb.top),
            bottom: convert_cell_border_override(&cb.bottom),
            left: convert_cell_border_override(&cb.left),
            right: convert_cell_border_override(&cb.right),
        }),
        (None, None) => None,
    };

    // §17.4.84: vertical alignment — direct cell, conditional, or default top.
    let valign = cell
        .properties
        .vertical_align
        .or_else(|| {
            cond.cell_properties
                .as_ref()
                .and_then(|tcp| tcp.vertical_align)
        })
        .map(|va| match va {
            model::CellVerticalAlign::Bottom => crate::render::layout::table::CellVAlign::Bottom,
            model::CellVerticalAlign::Center => crate::render::layout::table::CellVAlign::Center,
            _ => crate::render::layout::table::CellVAlign::Top,
        })
        .unwrap_or(crate::render::layout::table::CellVAlign::Top);

    // Estimate border insets to compute effective content width for
    // character-level splitting of oversized fragments.
    let border_w = |ovr: &Option<CellBorderOverride>| -> Pt {
        match ovr {
            Some(CellBorderOverride::Border(b)) => b.width,
            _ => Pt::ZERO,
        }
    };
    let border_inset_h = cell_borders
        .as_ref()
        .map(|cb| {
            let bl = (border_w(&cb.left) - cell_margins.left).max(Pt::ZERO);
            let br = (border_w(&cb.right) - cell_margins.right).max(Pt::ZERO);
            bl + br
        })
        .unwrap_or(Pt::ZERO);
    let content_width = (inner_width - border_inset_h).max(Pt::ZERO);

    // Recurse into cell content blocks.
    let cell_blocks =
        build_cell_blocks(&cell.content, table_style, cond, content_width, ctx, state);

    TableCellInput {
        blocks: cell_blocks,
        margins: cell_margins,
        grid_span: cell.properties.grid_span.unwrap_or(1),
        shading,
        cell_borders,
        vertical_merge: cell.properties.vertical_merge.map(|vm| match vm {
            model::VerticalMerge::Restart => {
                crate::render::layout::table::VerticalMergeState::Restart
            }
            model::VerticalMerge::Continue => {
                crate::render::layout::table::VerticalMergeState::Continue
            }
        }),
        vertical_align: valign,
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
    state: &mut BuildState,
) -> Vec<LayoutBlock> {
    let mut blocks = Vec::new();
    let mut pending_dropcap: Option<DropCapInfo> = None;

    for (i, block) in content.iter().enumerate() {
        match block {
            Block::Paragraph(p) => {
                // §17.4.66: every cell must end with a paragraph. When the
                // last block is an empty paragraph following a table, it is
                // structural — Word renders it with zero height.
                if p.content.is_empty()
                    && i > 0
                    && matches!(content[i - 1], Block::Table(_))
                    && i == content.len() - 1
                {
                    continue;
                }
                if let Some(lb) = build_paragraph_block(
                    p,
                    ctx,
                    state,
                    &mut pending_dropcap,
                    table_style,
                    Some(cond),
                ) {
                    // Split oversized text fragments for narrow cells.
                    let lb = if let LayoutBlock::Paragraph {
                        fragments,
                        style,
                        page_break_before,
                        footnotes,
                        floating_images,
                    } = lb
                    {
                        let fragments = split_oversized_fragments(fragments, inner_width, ctx);
                        LayoutBlock::Paragraph {
                            fragments,
                            style,
                            page_break_before,
                            footnotes,
                            floating_images,
                        }
                    } else {
                        lb
                    };
                    blocks.push(lb);
                }
            }
            Block::Table(nested_t) => {
                let built = build_table(nested_t, inner_width, ctx, state);
                blocks.push(LayoutBlock::Table {
                    rows: built.rows,
                    col_widths: built.col_widths,
                    border_config: built.border_config,
                    indent: built.indent,
                    alignment: built.alignment,
                    float_info: built.float_info,
                    style_id: nested_t.properties.style_id.clone(),
                });
            }
            _ => {}
        }
    }

    blocks
}
