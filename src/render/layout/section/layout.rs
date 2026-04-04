//! Core section layout — the `layout_section()` function and its private state types.

use super::super::draw_command::{DrawCommand, LayoutedPage};
use super::super::float;
use super::super::page::PageConfig;
use super::super::paragraph::{layout_paragraph, ParagraphBorderStyle, ParagraphStyle};
use super::super::table::{layout_table, TablePaginationConfig};
use super::super::BoxConstraints;
use super::helpers::{render_page_footnotes, split_at_column_breaks, table_x_offset};
use super::types::{ContinuationState, FloatingImageY, LayoutBlock};
use super::FLOAT_DEDUP_EPSILON_PT;
use super::FOOTNOTE_SEPARATOR_GAP;
use crate::model::StyleId;
use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

// ── Layout context and mutable page state ────────────────────────────────────

/// Read-only context passed to every section-layout helper.
/// Bundles the parameters that are constant for the lifetime of a `layout_section` call.
struct LayoutCtx<'cx> {
    config: &'cx PageConfig,
    measure_text: super::super::paragraph::MeasureTextFn<'cx>,
    separator_indent: Pt,
    default_line_height: Pt,
    /// Absolute bottom boundary of the printable area (page height − bottom margin).
    page_bottom: Pt,
}

/// All mutable paging state threaded through `layout_section`.
/// Extracted from the function to make ownership and page-break resets explicit.
struct PageLayoutState<'doc> {
    /// Fully laid-out pages emitted so far.
    pages: Vec<LayoutedPage>,
    /// The page currently being assembled.
    current_page: LayoutedPage,
    /// Current vertical cursor position on the page.
    cursor_y: Pt,
    /// §17.6.4: current column index (0-based).
    current_col: usize,
    /// §17.6.4: y at which columns start on the current page.
    column_top: Pt,
    /// Effective bottom boundary — reduced as footnotes are reserved.
    bottom: Pt,
    /// Footnotes accumulated for the current page.
    page_footnotes: Vec<(
        &'doc [super::super::fragment::Fragment],
        &'doc ParagraphStyle,
    )>,
    /// §17.3.1.33: true until the structural first content block is placed.
    first_on_section_page: bool,
    /// §17.3.1.9: space_after of the previous paragraph for spacing collapse.
    prev_space_after: Pt,
    /// §17.3.1.9: style_id of the previous paragraph for contextual spacing.
    prev_style_id: Option<StyleId>,
    /// §17.3.1.24: borders of the previous paragraph for border grouping.
    prev_borders: Option<ParagraphBorderStyle>,
    /// Active floats on the current page (text wraps around these).
    page_floats: Vec<float::ActiveFloat>,
    /// Forward-scanned absolute floats from future paragraphs on this page.
    current_page_abs_floats: Vec<float::ActiveFloat>,
    /// True when `current_page_abs_floats` needs rebuilding (e.g. after a page break).
    abs_floats_dirty: bool,
    /// Index of the first block on the current page (for forward scanning).
    page_start_block: usize,
    /// §17.4.59: y-start of the most recent paragraph (floating-table anchor).
    last_para_start_y: Pt,
    /// §17.4.38: style_id of the previous table for adjacent border collapse.
    prev_table_style_id: Option<StyleId>,
}

impl<'doc> PageLayoutState<'doc> {
    fn new(config: &PageConfig, continuation: Option<ContinuationState>, page_bottom: Pt) -> Self {
        let (current_page, cursor_y) = match continuation {
            Some(c) => (c.page, c.cursor_y),
            None => (LayoutedPage::new(config.page_size), config.margins.top),
        };
        PageLayoutState {
            pages: Vec::new(),
            column_top: cursor_y,
            last_para_start_y: cursor_y,
            current_page,
            cursor_y,
            current_col: 0,
            bottom: page_bottom,
            page_footnotes: Vec::new(),
            first_on_section_page: true,
            prev_space_after: Pt::ZERO,
            prev_style_id: None,
            prev_borders: None,
            page_floats: Vec::new(),
            current_page_abs_floats: Vec::new(),
            abs_floats_dirty: true,
            page_start_block: 0,
            prev_table_style_id: None,
        }
    }

    /// Render accumulated footnotes onto the current page and clear the list.
    fn flush_footnotes(&mut self, ctx: &LayoutCtx<'_>) {
        if !self.page_footnotes.is_empty() {
            render_page_footnotes(
                &mut self.current_page,
                ctx.config,
                &self.page_footnotes,
                ctx.default_line_height,
                ctx.measure_text,
                ctx.separator_indent,
            );
            self.page_footnotes.clear();
        }
    }

    /// Commit the current page and start a fresh one, resetting all per-page state.
    /// Callers that also need `prev_space_after = Pt::ZERO` must set that separately.
    fn push_new_page(&mut self, block_idx: usize, ctx: &LayoutCtx<'_>) {
        self.flush_footnotes(ctx);
        self.pages.push(std::mem::replace(
            &mut self.current_page,
            LayoutedPage::new(ctx.config.page_size),
        ));
        self.cursor_y = ctx.config.margins.top;
        self.column_top = ctx.config.margins.top;
        self.current_col = 0;
        self.bottom = ctx.page_bottom;
        self.page_start_block = block_idx;
        self.abs_floats_dirty = true;
        self.page_floats.clear();
    }

    /// Flush any remaining footnotes, push the last page, and return all pages.
    fn finalize(mut self, ctx: &LayoutCtx<'_>) -> Vec<LayoutedPage> {
        self.flush_footnotes(ctx);
        self.pages.push(self.current_page);
        self.pages
    }
}

/// Lay out a sequence of blocks into pages.
///
/// If `continuation` is provided, the section starts on the given page at the
/// given cursor_y (for `SectionType::Continuous` sections).
pub fn layout_section(
    blocks: &[LayoutBlock],
    config: &PageConfig,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
    default_line_height: Pt,
    continuation: Option<ContinuationState>,
) -> Vec<LayoutedPage> {
    let content_width = config.content_width();
    let num_cols = config.num_columns();
    let page_bottom = config.page_size.height - config.margins.bottom;

    let ctx = LayoutCtx {
        config,
        measure_text,
        separator_indent,
        default_line_height,
        page_bottom,
    };
    let mut state = PageLayoutState::new(config, continuation, page_bottom);

    // Column-aware constraints and x-offset for the current column.
    let col_constraints = |col: usize| -> BoxConstraints {
        let col_width = config.columns[col].width;
        BoxConstraints::new(Pt::ZERO, col_width, Pt::ZERO, config.content_height())
    };
    let col_x = |col: usize| -> Pt { config.margins.left + config.columns[col].x_offset };

    for (block_idx, block) in blocks.iter().enumerate() {
        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                page_break_before,
                footnotes,
                floating_images,
            } => {
                // §17.3.1.23: force a new page before this paragraph.
                if *page_break_before && state.cursor_y > config.margins.top {
                    state.push_new_page(block_idx, &ctx);
                    state.prev_space_after = Pt::ZERO;
                }

                let mut effective_style = style.clone();

                // Log paragraph info.
                let first_text = fragments
                    .iter()
                    .find_map(|f| {
                        if let super::super::fragment::Fragment::Text { text, .. } = f {
                            Some(&**text)
                        } else {
                            None
                        }
                    })
                    .unwrap_or("");
                log::debug!(
                    "[layout] block[{block_idx}] para style={:?} text={:?} cursor_y={:.1} col={} floats={} fwd_floats={}",
                    effective_style.style_id, &first_text[..first_text.len().min(30)],
                    state.cursor_y.raw(), state.current_col,
                    state.page_floats.len(), state.current_page_abs_floats.len()
                );

                // Save page state for potential keepNext rollback.
                let cmds_before_para = state.current_page.commands.len();

                // §17.3.1.33: suppress space_before for the structural first
                // paragraph of a section on its initial page.
                if state.cursor_y <= state.column_top && state.first_on_section_page {
                    effective_style.space_before = Pt::ZERO;
                }
                // §17.3.1.24: paragraph border grouping — consecutive paragraphs
                // with identical borders suppress interior top borders.
                if effective_style.borders.is_some()
                    && effective_style.borders == state.prev_borders
                {
                    if let Some(ref mut b) = effective_style.borders {
                        b.top = None;
                    }
                }
                // §17.3.1.9: spacing collapse (must happen before float registration).
                if effective_style.contextual_spacing
                    && effective_style.style_id.is_some()
                    && effective_style.style_id == state.prev_style_id
                {
                    state.cursor_y -= state.prev_space_after + effective_style.space_before;
                } else {
                    let collapse = state.prev_space_after.min(effective_style.space_before);
                    state.cursor_y -= collapse;
                }

                // Register floating images (both relative and absolute).
                // §20.4.2.18: wrapTopAndBottom images are emitted immediately
                // and cursor_y advances past them — they act as block spacers.
                // §20.4.2.10: paragraph-relative floats use the content area
                // top (after space_before), not the total paragraph box top.
                let content_top = state.cursor_y + effective_style.space_before;
                for fi in floating_images.iter() {
                    let (y_start, y_end) = match fi.y {
                        FloatingImageY::RelativeToParagraph(offset) => {
                            (content_top + offset, content_top + offset + fi.size.height)
                        }
                        FloatingImageY::Absolute(img_y) => (img_y, img_y + fi.size.height),
                    };
                    if fi.wrap_top_and_bottom {
                        // §20.4.2.18: emit now and advance cursor past the image.
                        let img_y = match fi.y {
                            FloatingImageY::Absolute(y) => y,
                            FloatingImageY::RelativeToParagraph(offset) => content_top + offset,
                        };
                        state.current_page.commands.push(DrawCommand::Image {
                            rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                            image_data: fi.image_data.clone(),
                        });
                        if y_end > state.cursor_y {
                            state.cursor_y = y_end;
                        }
                    } else {
                        // §20.4.2.3: use distL/distR for text distance from float.
                        let float_entry = float::ActiveFloat {
                            page_x: fi.x - fi.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fi.size.width + fi.dist_left + fi.dist_right,
                            source: float::FloatSource::Image,
                        };
                        log::debug!(
                            "[layout]   register image float: x={:.1} y={:.1}-{:.1} w={:.1}",
                            float_entry.page_x.raw(),
                            y_start.raw(),
                            y_end.raw(),
                            float_entry.width.raw()
                        );
                        state.page_floats.push(float_entry);
                    }
                }

                // Prune expired floats.
                float::prune_floats(&mut state.page_floats, state.cursor_y);

                // Forward-scan absolute floats from upcoming paragraphs on the
                // current page. Only rescan when the page changes.
                if state.abs_floats_dirty {
                    state.current_page_abs_floats.clear();
                    for (fi_idx, future_block) in
                        blocks[state.page_start_block..].iter().enumerate()
                    {
                        if let LayoutBlock::Paragraph {
                            floating_images: fi_list,
                            page_break_before,
                            ..
                        } = future_block
                        {
                            // Stop scanning at the next explicit page break
                            // (skip the first block — it may have triggered this page).
                            if *page_break_before && fi_idx > 0 {
                                break;
                            }
                            for fi in fi_list {
                                if fi.wrap_top_and_bottom {
                                    continue; // handled as block spacers, not floats
                                }
                                if let FloatingImageY::Absolute(img_y) = fi.y {
                                    state.current_page_abs_floats.push(float::ActiveFloat {
                                        page_x: fi.x - fi.dist_left,
                                        page_y_start: img_y,
                                        page_y_end: img_y + fi.size.height,
                                        width: fi.size.width + fi.dist_left + fi.dist_right,
                                        source: float::FloatSource::Image,
                                    });
                                }
                            }
                        }
                    }
                    state.abs_floats_dirty = false;
                }

                // Merge page_floats with forward-scanned absolute floats (dedup).
                // Only include absolute floats whose y range starts at or above
                // the current cursor — floats below shouldn't affect text above.
                let mut effective_floats = state.page_floats.clone();
                let y_threshold = state.cursor_y + effective_style.space_before;
                let deduped: Vec<float::ActiveFloat> = state
                    .current_page_abs_floats
                    .iter()
                    .filter(|af| af.page_y_start <= y_threshold)
                    .filter(|af| {
                        !effective_floats.iter().any(|pf| {
                            (pf.page_x - af.page_x).raw().abs() < FLOAT_DEDUP_EPSILON_PT
                                && (pf.page_y_start - af.page_y_start).raw().abs()
                                    < FLOAT_DEDUP_EPSILON_PT
                                && (pf.page_y_end - af.page_y_end).raw().abs()
                                    < FLOAT_DEDUP_EPSILON_PT
                        })
                    })
                    .cloned()
                    .collect();
                effective_floats.extend(deduped);
                if !effective_floats.is_empty() {
                    for (i, f) in effective_floats.iter().enumerate() {
                        log::debug!(
                            "[layout]   effective_float[{i}]: x={:.1} y={:.1}-{:.1} w={:.1} src={:?}",
                            f.page_x.raw(), f.page_y_start.raw(),
                            f.page_y_end.raw(), f.width.raw(), f.source
                        );
                    }
                }
                // §17.4.56: advance past any full-width float that blocks all text.
                let col_width = config.columns[state.current_col].width;
                let page_x = col_x(state.current_col);
                for ef in &effective_floats {
                    if ef.overlaps_y(state.cursor_y) && ef.width >= col_width {
                        state.cursor_y = state.cursor_y.max(ef.page_y_end);
                    }
                }
                float::prune_floats(&mut effective_floats, state.cursor_y);

                effective_style.page_floats = effective_floats;
                effective_style.page_y = state.cursor_y;
                effective_style.page_x = page_x;
                effective_style.page_content_width = col_width;

                // §17.6.4: split paragraph at column breaks for multi-column layout.
                let frag_chunks = split_at_column_breaks(fragments);
                let mut para_start_y = state.cursor_y;
                state.last_para_start_y = state.cursor_y;

                for (chunk_idx, chunk) in frag_chunks.iter().enumerate() {
                    // Advance to the next column for chunks after a column break.
                    if chunk_idx > 0 {
                        if state.current_col + 1 < num_cols {
                            state.current_col += 1;
                        } else {
                            // All columns full — new page, reset to column 0.
                            state.push_new_page(block_idx, &ctx);
                        }
                        state.cursor_y = state.column_top;
                        effective_style.page_x = col_x(state.current_col);
                        effective_style.page_content_width =
                            config.columns[state.current_col].width;
                    }

                    let constraints = col_constraints(state.current_col);
                    let para = layout_paragraph(
                        chunk,
                        &constraints,
                        &effective_style,
                        ctx.default_line_height,
                        ctx.measure_text,
                    );

                    // Column/page overflow: advance column, then page.
                    if state.cursor_y + para.size.height > state.bottom
                        && state.cursor_y > state.column_top
                    {
                        if state.current_col + 1 < num_cols {
                            state.current_col += 1;
                            state.cursor_y = state.column_top;
                        } else {
                            state.push_new_page(block_idx, &ctx);
                        }
                        // Update para_start_y after page/column change so
                        // floating images use the correct position.
                        para_start_y = state.cursor_y;
                    }

                    log::debug!(
                        "[layout]   chunk[{chunk_idx}] placed at y={:.1} x={:.1} height={:.1}",
                        state.cursor_y.raw(),
                        col_x(state.current_col).raw(),
                        para.size.height.raw()
                    );
                    for mut cmd in para.commands {
                        cmd.shift_y(state.cursor_y);
                        cmd.shift_x(col_x(state.current_col));
                        state.current_page.commands.push(cmd);
                    }
                    state.cursor_y += para.size.height;
                }

                // §17.3.1.14: keepNext — if this paragraph has keep_next, check
                // whether the next block fits. If not, undo placement and page-break.
                if effective_style.keep_next
                    && state.cursor_y > state.column_top
                    && block_idx + 1 < blocks.len()
                    && num_cols <= 1
                // skip for multi-column (too complex)
                {
                    let next_fits = match &blocks[block_idx + 1] {
                        LayoutBlock::Paragraph {
                            fragments: next_frags,
                            style: next_style,
                            ..
                        } => {
                            let mut next_eff = next_style.clone();
                            let next_collapse =
                                effective_style.space_after.min(next_eff.space_before);
                            let next_cursor = state.cursor_y - next_collapse;
                            let next_constraints = col_constraints(state.current_col);
                            next_eff.page_y = next_cursor;
                            next_eff.page_x = col_x(state.current_col);
                            next_eff.page_content_width = config.columns[state.current_col].width;
                            let next_para = layout_paragraph(
                                next_frags,
                                &next_constraints,
                                &next_eff,
                                ctx.default_line_height,
                                ctx.measure_text,
                            );
                            next_cursor + next_para.size.height <= state.bottom
                        }
                        LayoutBlock::Table { .. } => true, // don't keepNext with tables
                    };

                    if !next_fits {
                        // Undo this paragraph's commands, then page-break.
                        state.current_page.commands.truncate(cmds_before_para);
                        state.push_new_page(block_idx, &ctx);

                        // Re-layout the paragraph on the fresh page.
                        // keepNext paragraphs at page top retain their space_before.
                        let constraints = col_constraints(state.current_col);
                        effective_style.page_y = state.cursor_y;
                        effective_style.page_x = col_x(state.current_col);
                        effective_style.page_content_width =
                            config.columns[state.current_col].width;
                        effective_style.page_floats = Vec::new();
                        let para = layout_paragraph(
                            fragments,
                            &constraints,
                            &effective_style,
                            ctx.default_line_height,
                            ctx.measure_text,
                        );
                        for mut cmd in para.commands {
                            cmd.shift_y(state.cursor_y);
                            cmd.shift_x(col_x(state.current_col));
                            state.current_page.commands.push(cmd);
                        }
                        state.cursor_y += para.size.height;
                    }
                }

                state.first_on_section_page = false;
                state.prev_borders = style.borders.clone();
                state.prev_space_after = effective_style.space_after;
                state.prev_style_id = effective_style.style_id.clone();
                state.prev_table_style_id = None; // paragraph breaks adjacent table chain

                // §20.4.2.3: emit non-wrapTopAndBottom floating images.
                // (wrapTopAndBottom images were emitted immediately above.)
                for fi in floating_images {
                    if fi.wrap_top_and_bottom {
                        continue;
                    }
                    let img_y = match fi.y {
                        FloatingImageY::Absolute(y) => y,
                        FloatingImageY::RelativeToParagraph(offset) => {
                            para_start_y + effective_style.space_before + offset
                        }
                    };
                    state.current_page.commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                        image_data: fi.image_data.clone(),
                    });
                }

                // Collect footnotes for this page and reduce the available bottom.
                if !footnotes.is_empty() {
                    let fn_constraints = BoxConstraints::tight_width(content_width, Pt::INFINITY);
                    let sep_height = FOOTNOTE_SEPARATOR_GAP;
                    for (fn_frags, fn_style) in footnotes {
                        let fn_para = layout_paragraph(
                            fn_frags,
                            &fn_constraints,
                            fn_style,
                            ctx.default_line_height,
                            ctx.measure_text,
                        );
                        // Reserve separator space only for the first footnote on this page.
                        if state.page_footnotes.is_empty() {
                            state.bottom -= sep_height;
                        }
                        state.bottom -= fn_para.size.height;
                        state.page_footnotes.push((fn_frags, fn_style));
                    }
                }
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                border_config,
                indent,
                alignment,
                float_info,
                style_id,
            } => {
                // §17.4.58: floating table — render and register as a float so
                // subsequent text wraps around it.  Floating tables are absolutely
                // positioned and do not participate in adjacent border collapse.
                if let Some(fi) = float_info {
                    let table = layout_table(
                        rows,
                        col_widths,
                        &col_constraints(state.current_col),
                        ctx.default_line_height,
                        border_config.as_ref(),
                        ctx.measure_text,
                        false,
                    );

                    // §17.4.28 / §17.4.51: compute table x position.
                    let table_x = table_x_offset(
                        *alignment,
                        *indent,
                        table.size.width,
                        content_width,
                        config.margins.left,
                    );

                    // Floating table breaks the adjacent table chain.
                    state.prev_table_style_id = None;
                    // §17.4.58: apply tblpXSpec horizontal alignment override.
                    let table_x = match fi.x_align {
                        Some(crate::model::TableXAlign::Center) => {
                            config.margins.left + (content_width - table.size.width) * 0.5
                        }
                        Some(crate::model::TableXAlign::Right) => {
                            config.margins.left + content_width - table.size.width
                        }
                        _ => table_x,
                    };
                    if state.cursor_y + table.size.height > state.bottom
                        && state.cursor_y > config.margins.top
                    {
                        state.push_new_page(block_idx, &ctx);
                        state.prev_space_after = Pt::ZERO;
                    }

                    // §17.4.59: tblpY is the absolute Y offset from the vertical
                    // anchor. The table must not start before cursor_y.
                    let float_y_start = if fi.y_offset > Pt::ZERO {
                        let anchor_y = match fi.vert_anchor {
                            crate::model::TableAnchor::Text => {
                                state.last_para_start_y + fi.y_offset
                            }
                            crate::model::TableAnchor::Margin => config.margins.top + fi.y_offset,
                            crate::model::TableAnchor::Page => fi.y_offset,
                        };
                        anchor_y.max(state.cursor_y)
                    } else {
                        state.cursor_y
                    };
                    // §17.4.56: float y_end uses only the table's visual height.
                    let float_y_end = float_y_start + table.size.height;

                    for mut cmd in table.commands {
                        cmd.shift_y(float_y_start);
                        cmd.shift_x(table_x);
                        state.current_page.commands.push(cmd);
                    }

                    log::debug!(
                        "[layout]   register table float: x={:.1} y={:.1}-{:.1} w={:.1} block_idx={block_idx}",
                        table_x.raw(), float_y_start.raw(), float_y_end.raw(),
                        (table.size.width + fi.right_gap).raw()
                    );
                    // §17.4.56: register as float for text wrapping.
                    state.page_floats.push(float::ActiveFloat {
                        page_x: table_x,
                        page_y_start: float_y_start,
                        page_y_end: float_y_end,
                        width: table.size.width + fi.right_gap,
                        source: float::FloatSource::Table {
                            owner_block_idx: block_idx,
                        },
                    });
                    continue;
                }

                // §17.4.38: consecutive non-floating tables with the same style
                // are treated as one merged table — the second table's top border
                // is suppressed so the shared edge is drawn once.
                let suppress_top = style_id.is_some() && *style_id == state.prev_table_style_id;

                // Non-floating table: paginated row-level splitting.
                // §17.4.49 / §17.4.1: split at row boundaries, repeat headers.
                let available = state.bottom - state.cursor_y;
                let slices = super::super::table::layout_table_paginated(
                    rows,
                    col_widths,
                    &col_constraints(state.current_col),
                    ctx.default_line_height,
                    border_config.as_ref(),
                    ctx.measure_text,
                    &TablePaginationConfig {
                        available_height: available,
                        page_height: config.content_height(),
                        suppress_first_row_top: suppress_top,
                    },
                );

                // §17.4.28 / §17.4.51: compute table x position.
                let table_width: Pt = col_widths.iter().copied().sum();
                let table_x = table_x_offset(
                    *alignment,
                    *indent,
                    table_width,
                    content_width,
                    config.margins.left,
                );

                for (slice_idx, slice) in slices.into_iter().enumerate() {
                    if slice_idx > 0 {
                        // Continuation slice — start a new page.
                        state.push_new_page(block_idx, &ctx);
                    }
                    for mut cmd in slice.commands {
                        cmd.shift_y(state.cursor_y);
                        cmd.shift_x(table_x);
                        state.current_page.commands.push(cmd);
                    }
                    state.cursor_y += slice.size.height;
                }
                state.first_on_section_page = false;
                state.prev_borders = None; // table breaks border grouping
                state.prev_space_after = Pt::ZERO;
                state.prev_style_id = None;
                state.prev_table_style_id = style_id.clone();
            }
        }
    }

    // Flush remaining footnotes and push the last page.
    state.finalize(&ctx)
}
