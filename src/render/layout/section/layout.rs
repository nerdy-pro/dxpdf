//! Core section layout — the `layout_section()` function and its private state types.

use super::super::draw_command::{DrawCommand, LayoutedPage};
use super::super::float;
use super::super::fragment::Fragment;
use super::super::header_footer::{HeaderFooterClearance, PageBodyBounds};
use super::super::page::PageConfig;
use super::super::paragraph::{
    layout_paragraph, place_paragraph, ParagraphBorderStyle, ParagraphStyle,
};
use super::super::table::{
    layout_table, layout_table_paginated_with_page_heights, measure_leading_table_group_height,
    TablePaginationHeights,
};
use super::super::BoxConstraints;
use super::floating_table::{
    plan_floating_table_pages_with_page_tops, resolve_floating_anchor, FloatingTableAnchor,
    FloatingTablePagePlacement,
};
use super::helpers::{
    render_page_footnotes, split_at_column_breaks, split_at_page_breaks, table_x_offset,
};
use super::types::{
    ContinuationState, FloatingImage, FloatingShape, LayoutBlock, PageParity, WrapMode,
};
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
    clearance: &'cx HeaderFooterClearance,
    last_page: LastPageOwner,
    measure_text: super::super::paragraph::MeasureTextFn<'cx>,
    separator_indent: Pt,
    default_line_height: Pt,
}

impl LayoutCtx<'_> {
    fn page_bounds(&self, section_page_index: usize) -> PageBodyBounds {
        match self.last_page.bounds_override() {
            // §17.6.22: this page is shared with a following Continuous
            // section, which owns its header and footer (see `SectionStart`).
            // Every *other* page keeps this section's own clearance — which is
            // why the override is indexed rather than applied to "the last
            // page seen so far": content that spills off the shared page must
            // not drag the override along with it.
            Some(FinalPageBounds { index, bounds }) if index == section_page_index => bounds,
            _ => self.clearance.for_page(section_page_index),
        }
    }
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
    /// 0-based physical page index within the current section.
    page_index: usize,
    /// §17.10.6: logical number of this section's first page, with
    /// `w:pgNumType/@start` applied. With `page_index` it gives the current
    /// page's logical number, and so the §20.4.3.1 parity an `inside`/
    /// `outside` float mirrors on.
    logical_page_base: usize,
    /// Effective top boundary for the selected header slot on this page.
    page_top: Pt,
    /// §17.6.4: current column index (0-based).
    current_col: usize,
    /// §17.6.4: y at which columns start on the current page.
    column_top: Pt,
    /// §17.6.22: true while `column_top` is a *continuation* cursor rather than
    /// a page top — the column set inherited from a continuous break, which
    /// starts part-way down a page the preceding section already filled.
    ///
    /// Such a column is strictly shorter than the ones that follow it, which
    /// makes it the one case where being at a column's top does not mean having
    /// as much room as anywhere else. See
    /// [`at_full_height_column_top`](PageLayoutState::at_full_height_column_top).
    inherited_short_columns: bool,
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
    /// §17.4.57: y-start of the most recent paragraph (floating-table anchor).
    /// Set after that paragraph's spacing collapse but before its
    /// `space_before` is added.
    last_para_start_y: Pt,
    /// §17.4.38: style_id of the previous table for adjacent border collapse.
    prev_table_style_id: Option<StyleId>,
    /// §17.3.3.1: an inline page break at the end of a paragraph defers the
    /// page break to the start of the next block.
    pending_page_break: bool,
    /// §17.11.23: notes reserved by a *previous* section on this same page,
    /// carried across a continuous break unflushed so one separator covers
    /// them all. Owned, because the blocks they came from are gone.
    carried_footnotes: Vec<(Vec<super::super::fragment::Fragment>, ParagraphStyle)>,
}

impl<'doc> PageLayoutState<'doc> {
    fn new(
        config: &PageConfig,
        continuation: Option<ContinuationState>,
        bounds: PageBodyBounds,
        logical_page_base: usize,
    ) -> Self {
        // §17.6.22: a continuation resumes *inside* a page, so it inherits
        // that page's in-progress state rather than a fresh page's defaults.
        // Columns are the deliberate exception — see below.
        match continuation {
            Some(c) => PageLayoutState {
                pages: Vec::new(),
                column_top: c.cursor_y,
                last_para_start_y: c.last_para_start_y,
                current_page: c.page,
                cursor_y: c.cursor_y,
                page_index: 0,
                logical_page_base,
                page_top: bounds.top,
                // §17.6.4: a continuous break is the usual way to *change* the
                // column count, so the incoming index may not exist in this
                // section's config — carrying it panics the `config.columns
                // [state.current_col]` below (3 columns before the break, 1
                // after: "len is 1 but the index is 2"). Restarting at column 0
                // from the shared cursor is both safe and what the change
                // means.
                //
                // Word additionally *balances* the preceding section's columns
                // at such a break, so its content is spread evenly above the
                // new column set rather than left as it fell. That is not
                // implemented: it needs the preceding section's last page
                // re-flowed to an equal-height target, which is a different
                // problem from the clearance relayout around it.
                current_col: 0,
                // The inherited column set starts at the shared cursor, so it
                // is shorter than a fresh page's — *strictly* shorter, which is
                // the property the predicate's reasoning rests on. A break
                // landing exactly on the page top inherits a full-height column
                // and needs no exception.
                inherited_short_columns: c.cursor_y > bounds.top,
                bottom: c.bottom,
                page_footnotes: Vec::new(),
                carried_footnotes: c.page_footnotes,
                // §17.3.1.33: the page top is above the preceding section's
                // content, not here — space_before must not be suppressed.
                first_on_section_page: false,
                prev_space_after: c.prev_space_after,
                prev_style_id: c.prev_style_id,
                prev_borders: c.prev_borders,
                page_floats: c.page_floats,
                current_page_abs_floats: Vec::new(),
                abs_floats_dirty: true,
                page_start_block: 0,
                prev_table_style_id: c.prev_table_style_id,
                pending_page_break: c.pending_page_break,
            },
            None => PageLayoutState {
                pages: Vec::new(),
                column_top: bounds.top,
                last_para_start_y: bounds.top,
                current_page: LayoutedPage::new(config.page_size),
                cursor_y: bounds.top,
                page_index: 0,
                logical_page_base,
                page_top: bounds.top,
                current_col: 0,
                inherited_short_columns: false,
                bottom: bounds.bottom,
                page_footnotes: Vec::new(),
                carried_footnotes: Vec::new(),
                first_on_section_page: true,
                prev_space_after: Pt::ZERO,
                prev_style_id: None,
                prev_borders: None,
                page_floats: Vec::new(),
                current_page_abs_floats: Vec::new(),
                abs_floats_dirty: true,
                page_start_block: 0,
                prev_table_style_id: None,
                pending_page_break: false,
            },
        }
    }

    /// True when the cursor sits at the top of a column that already offers as
    /// much height as any column this section could move the content to — so
    /// moving it again cannot help, and the caller must place it in-line and
    /// let it overflow rather than loop.
    ///
    /// §17.6.4 makes a fresh column equivalent to a fresh page, which is why
    /// the test is against `column_top` and not the page top. §17.6.22 is the
    /// one exception: a continuation's column set starts at the *shared
    /// cursor*, part-way down an inherited page, so it is strictly shorter than
    /// the ones after it and moving content off it genuinely does help.
    ///
    /// Reading the raw `cursor_y <= column_top` instead strands content on the
    /// shared page: a §17.3.1.15 `keepNext` chain that starts there is **torn**,
    /// its head left behind while the rest moves on — which the same document
    /// without the break does not do.
    ///
    /// Used at the two sites that move a whole block that does not fit. The
    /// paragraph-*splitting* site deliberately keeps the raw test; the reason
    /// it trades differently is recorded there.
    ///
    /// Termination is unaffected. The exception only ever *permits* one move,
    /// onto a page whose own column set is full height — where the flag is
    /// clear and the ordinary guard applies again.
    fn at_full_height_column_top(&self) -> bool {
        self.cursor_y <= self.column_top && !self.inherited_short_columns
    }

    /// Render accumulated footnotes onto the current page and clear the list.
    /// §20.4.3.1: parity of the page currently being assembled.
    fn parity(&self) -> PageParity {
        PageParity::of_page(self.logical_page_base + self.page_index)
    }

    /// §17.11.23: render every note owed by this page under **one** separator.
    ///
    /// Notes carried across a continuous break are rendered together with this
    /// section's own, because a separator belongs to a page, not to a section:
    /// flushing per section draws a second rule and positions the second block
    /// from the page's unreduced bottom, laying it over the first.
    ///
    /// The vectors are taken out first so the borrow of `current_page` and the
    /// borrows of the note fragments do not overlap; both are cleared by the
    /// flush regardless.
    fn flush_footnotes(&mut self, ctx: &LayoutCtx<'_>) {
        if self.page_footnotes.is_empty() && self.carried_footnotes.is_empty() {
            return;
        }
        let carried = std::mem::take(&mut self.carried_footnotes);
        let own = std::mem::take(&mut self.page_footnotes);
        let mut all: Vec<(&[super::super::fragment::Fragment], &ParagraphStyle)> = carried
            .iter()
            .map(|(frags, style)| (frags.as_slice(), style))
            .collect();
        all.extend(own.iter().copied());
        render_page_footnotes(
            &mut self.current_page,
            ctx.config,
            &all,
            ctx.default_line_height,
            ctx.measure_text,
            ctx.separator_indent,
            ctx.page_bounds(self.page_index).bottom,
        );
    }

    /// Commit the current page and start a fresh one, resetting all per-page state.
    /// Callers that also need `prev_space_after = Pt::ZERO` must set that separately.
    fn push_new_page(&mut self, block_idx: usize, ctx: &LayoutCtx<'_>) {
        self.flush_footnotes(ctx);
        self.pages.push(std::mem::replace(
            &mut self.current_page,
            LayoutedPage::new(ctx.config.page_size),
        ));
        self.page_index += 1;
        let bounds = ctx.page_bounds(self.page_index);
        self.page_top = bounds.top;
        self.cursor_y = bounds.top;
        self.column_top = bounds.top;
        self.current_col = 0;
        // §17.6.22: whatever page we came from, this one starts at its own top,
        // so the inherited short column set is behind us.
        self.inherited_short_columns = false;
        self.bottom = bounds.bottom;
        self.page_start_block = block_idx;
        self.abs_floats_dirty = true;
        self.page_floats.clear();
    }

    /// Flush any remaining footnotes, push the last page, and return all pages
    /// together with the cursor the section left off at.
    ///
    /// The cursor is *returned* rather than recovered from the emitted commands
    /// because the two disagree: the last command's extent is not the flow
    /// position — trailing `space_after`, an empty final paragraph and a
    /// footnote block at the page bottom all move one without the other. A
    /// §17.6.22 continuation resumes at the flow position.
    fn finalize(mut self, ctx: &LayoutCtx<'_>) -> SectionLayout {
        match ctx.last_page {
            LastPageOwner::Own => {
                self.flush_footnotes(ctx);
                self.pages.push(self.current_page);
                SectionLayout {
                    pages: self.pages,
                    tail: SectionTail::Complete,
                }
            }
            // §17.6.22: the last page belongs to the section that follows. It
            // is handed over *unfinished* — footnotes deliberately unflushed,
            // so the succeeding section's flush covers both under one
            // separator — and stays out of `pages` so no caller can commit it
            // twice.
            LastPageOwner::SharedWithNext { .. } => {
                let carried = std::mem::take(&mut self.carried_footnotes);
                let own: Vec<_> = self
                    .page_footnotes
                    .iter()
                    .map(|(frags, style)| (frags.to_vec(), (*style).clone()))
                    .collect();
                SectionLayout {
                    pages: self.pages,
                    tail: SectionTail::SharedWithNext(ContinuationState {
                        page: self.current_page,
                        cursor_y: self.cursor_y,
                        bottom: self.bottom,
                        page_footnotes: carried.into_iter().chain(own).collect(),
                        page_floats: self.page_floats,
                        prev_space_after: self.prev_space_after,
                        prev_style_id: self.prev_style_id,
                        prev_borders: self.prev_borders,
                        prev_table_style_id: self.prev_table_style_id,
                        last_para_start_y: self.last_para_start_y,
                        pending_page_break: self.pending_page_break,
                    }),
                }
            }
        }
    }

    /// The floats affecting text at the current cursor on the current
    /// page/column: registered page floats (§20.4.2) plus forward-scanned
    /// absolute floats from upcoming blocks on this page (rebuilt once per page
    /// while `abs_floats_dirty`), advancing the cursor past any full-width float
    /// that blocks all text (§ not verified: `w:tblOverlap` governs table-vs-
    /// table overlap only, and no section states this cursor rule for a float
    /// of any kind). Shared by the main block loop and the
    /// across-page split re-fit (§4) so a continuation wraps around whatever
    /// floats live on *its* page, not the paragraph's starting page.
    ///
    /// Word runs multi-pass layout, where every float on a page affects all of
    /// that page's text regardless of document order. This single-pass
    /// renderer approximates that by forward-scanning upcoming blocks for
    /// absolute floats when a page starts, rather than only ever knowing about
    /// floats already encountered.
    fn effective_floats_at_cursor(
        &mut self,
        blocks: &[LayoutBlock],
        relocated_absolute_float_blocks: &std::collections::HashSet<usize>,
        num_cols: usize,
        space_before: Pt,
        col_width: Pt,
    ) -> Vec<float::ActiveFloat> {
        let parity = self.parity();
        // Forward-scan absolute floats from upcoming paragraphs on the current
        // page. Only rescan when the page changes. The scan tracks inline
        // column/page boundaries (`scan_inline_page_boundary`) so a float past
        // an in-paragraph break isn't attributed to this page, and skips blocks
        // whose absolute float has been relocated to a later page (#86's
        // relocation replay), so wrapped text on the replayed page ignores it.
        if self.abs_floats_dirty {
            self.current_page_abs_floats.clear();
            let mut scan_col = self.current_col;
            for (fi_idx, future_block) in blocks[self.page_start_block..].iter().enumerate() {
                let future_block_idx = self.page_start_block + fi_idx;
                if relocated_absolute_float_blocks.contains(&future_block_idx) {
                    if fi_idx == 0 {
                        continue;
                    }
                    break;
                }
                if let LayoutBlock::Paragraph {
                    floating_images: fi_list,
                    page_break_before,
                    fragments,
                    ..
                } = future_block
                {
                    // Stop scanning at the next explicit page break (skip the
                    // first block — it may have triggered this page).
                    if *page_break_before && fi_idx > 0 {
                        break;
                    }
                    let boundary = scan_inline_page_boundary(fragments, &mut scan_col, num_cols);
                    if boundary != ForwardScanBoundary::BeforeParagraphFloat {
                        for fi in fi_list {
                            // §20.4.2.15/.18: only wrap-enabled modes narrow
                            // text. `TopAndBottom` is a block spacer and
                            // `None` is a pure overlay — matches the gate in
                            // `has_absolute_wrap_float` below.
                            if !fi.wrap_mode.registers_as_wrap_float() {
                                continue;
                            }
                            if let Some(img_y) = fi.y.absolute(parity) {
                                self.current_page_abs_floats.push(float::ActiveFloat {
                                    page_x: fi.x.resolve(parity) - fi.dist_left,
                                    page_y_start: img_y,
                                    page_y_end: img_y + fi.size.height,
                                    width: fi.size.width + fi.dist_left + fi.dist_right,
                                    source: float::FloatSource::Image,
                                    wrap_text: fi.wrap_mode.wrap_text().into(),
                                });
                            }
                        }
                    }
                    if boundary != ForwardScanBoundary::None {
                        break;
                    }
                }
            }
            self.abs_floats_dirty = false;
        }

        // Merge page_floats with forward-scanned absolute floats (dedup). Only
        // include absolute floats whose y range starts at or above the current
        // cursor — floats below shouldn't affect text above.
        let mut effective_floats = self.page_floats.clone();
        let y_threshold = self.cursor_y + space_before;
        let deduped: Vec<float::ActiveFloat> = self
            .current_page_abs_floats
            .iter()
            .filter(|af| af.page_y_start <= y_threshold)
            .filter(|af| {
                !effective_floats.iter().any(|pf| {
                    (pf.page_x - af.page_x).raw().abs() < FLOAT_DEDUP_EPSILON_PT
                        && (pf.page_y_start - af.page_y_start).raw().abs() < FLOAT_DEDUP_EPSILON_PT
                        && (pf.page_y_end - af.page_y_end).raw().abs() < FLOAT_DEDUP_EPSILON_PT
                })
            })
            .cloned()
            .collect();
        effective_floats.extend(deduped);

        // Advance past any full-width float that blocks all text — see
        // `effective_floats_at_cursor` for why this cites no section.
        for ef in &effective_floats {
            if ef.overlaps_y(self.cursor_y) && ef.width >= col_width {
                self.cursor_y = self.cursor_y.max(ef.page_y_end);
            }
        }
        float::prune_floats(&mut effective_floats, self.cursor_y);
        effective_floats
    }
}

struct ParagraphFloatCheckpoint {
    command_count: usize,
    page_floats: Vec<float::ActiveFloat>,
    cursor_y: Pt,
}

struct PageReplayCheckpoint<'doc> {
    current_page: LayoutedPage,
    cursor_y: Pt,
    page_index: usize,
    page_top: Pt,
    current_col: usize,
    column_top: Pt,
    /// Saved with `column_top`, which it qualifies: restoring one without the
    /// other would describe a column set that never existed.
    inherited_short_columns: bool,
    bottom: Pt,
    page_footnotes: Vec<(
        &'doc [super::super::fragment::Fragment],
        &'doc ParagraphStyle,
    )>,
    first_on_section_page: bool,
    prev_space_after: Pt,
    prev_style_id: Option<StyleId>,
    prev_borders: Option<ParagraphBorderStyle>,
    page_floats: Vec<float::ActiveFloat>,
    current_page_abs_floats: Vec<float::ActiveFloat>,
    abs_floats_dirty: bool,
    page_start_block: usize,
    last_para_start_y: Pt,
    prev_table_style_id: Option<StyleId>,
    pending_page_break: bool,
}

impl<'doc> PageReplayCheckpoint<'doc> {
    fn capture(state: &PageLayoutState<'doc>) -> Self {
        Self {
            current_page: state.current_page.clone(),
            cursor_y: state.cursor_y,
            page_index: state.page_index,
            page_top: state.page_top,
            current_col: state.current_col,
            column_top: state.column_top,
            inherited_short_columns: state.inherited_short_columns,
            bottom: state.bottom,
            page_footnotes: state.page_footnotes.clone(),
            first_on_section_page: state.first_on_section_page,
            prev_space_after: state.prev_space_after,
            prev_style_id: state.prev_style_id.clone(),
            prev_borders: state.prev_borders.clone(),
            page_floats: state.page_floats.clone(),
            current_page_abs_floats: state.current_page_abs_floats.clone(),
            abs_floats_dirty: state.abs_floats_dirty,
            page_start_block: state.page_start_block,
            last_para_start_y: state.last_para_start_y,
            prev_table_style_id: state.prev_table_style_id.clone(),
            pending_page_break: state.pending_page_break,
        }
    }

    fn restore(&self, state: &mut PageLayoutState<'doc>) {
        state.current_page.clone_from(&self.current_page);
        state.cursor_y = self.cursor_y;
        state.page_index = self.page_index;
        state.page_top = self.page_top;
        state.current_col = self.current_col;
        state.column_top = self.column_top;
        state.inherited_short_columns = self.inherited_short_columns;
        state.bottom = self.bottom;
        state.page_footnotes.clone_from(&self.page_footnotes);
        state.first_on_section_page = self.first_on_section_page;
        state.prev_space_after = self.prev_space_after;
        state.prev_style_id.clone_from(&self.prev_style_id);
        state.prev_borders.clone_from(&self.prev_borders);
        state.page_floats.clone_from(&self.page_floats);
        state
            .current_page_abs_floats
            .clone_from(&self.current_page_abs_floats);
        state.abs_floats_dirty = self.abs_floats_dirty;
        state.page_start_block = self.page_start_block;
        state.last_para_start_y = self.last_para_start_y;
        state
            .prev_table_style_id
            .clone_from(&self.prev_table_style_id);
        state.pending_page_break = self.pending_page_break;
    }
}

fn refresh_page_replay_checkpoint<'doc>(
    state: &PageLayoutState<'doc>,
    block_idx: usize,
    checkpoint: &mut PageReplayCheckpoint<'doc>,
    checkpoint_page_index: &mut usize,
    replay_block_idx: &mut usize,
) {
    if state.page_index != *checkpoint_page_index {
        *checkpoint = PageReplayCheckpoint::capture(state);
        *checkpoint_page_index = state.page_index;
        *replay_block_idx = block_idx;
    }
}

impl ParagraphFloatCheckpoint {
    fn capture(state: &PageLayoutState<'_>) -> Self {
        Self {
            command_count: state.current_page.commands.len(),
            page_floats: state.page_floats.clone(),
            cursor_y: state.cursor_y,
        }
    }

    fn restore(&self, state: &mut PageLayoutState<'_>) {
        state.current_page.commands.truncate(self.command_count);
        state.page_floats.clone_from(&self.page_floats);
        state.cursor_y = self.cursor_y;
    }
}

/// §17.17.1 / §20.1.2.1.1: emit a floating shape's `wps:txbx` text over its
/// fill. Each command is in shape-local coordinates; shift by the shape's
/// resolved page origin `(fs.x, shape_y)`. Mirrors the stacker's shape-text
/// emission (`section::stacker`) so body/page-anchored shapes render their text,
/// not just the fill/stroke.
fn emit_shape_text(state: &mut PageLayoutState<'_>, fs: &FloatingShape, shape_y: Pt) {
    let parity = state.parity();
    for mut cmd in fs.text_commands.iter().cloned() {
        cmd.shift(fs.x.resolve(parity), shape_y);
        state.current_page.commands.push(cmd);
    }
}

fn register_paragraph_floats(
    state: &mut PageLayoutState<'_>,
    floating_images: &[FloatingImage],
    floating_shapes: &[FloatingShape],
    content_top: Pt,
) {
    let parity = state.parity();
    for fi in floating_images {
        let y_start = fi.y.at(parity, content_top);
        let y_end = y_start + fi.size.height;
        if fi.is_wrap_top_and_bottom() {
            state.current_page.commands.push(DrawCommand::Image {
                rect: PtRect::from_xywh(
                    fi.x.resolve(parity),
                    y_start,
                    fi.size.width,
                    fi.size.height,
                ),
                image_data: fi.image_data.clone(),
                src_rect: fi.src_rect,
            });
            if y_end > state.cursor_y {
                state.cursor_y = y_end;
            }
        } else if fi.wrap_mode.registers_as_wrap_float() {
            let float_entry = float::ActiveFloat {
                page_x: fi.x.resolve(parity) - fi.dist_left,
                page_y_start: y_start,
                page_y_end: y_end,
                width: fi.size.width + fi.dist_left + fi.dist_right,
                source: float::FloatSource::Image,
                wrap_text: fi.wrap_mode.wrap_text().into(),
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

    for fs in floating_shapes {
        if matches!(fs.wrap_mode, WrapMode::None) {
            continue;
        }
        let y_start = fs.y.at(parity, content_top);
        let y_end = y_start + fs.size.height;
        if fs.is_wrap_top_and_bottom() {
            let shape_y = y_start;
            state.current_page.commands.push(DrawCommand::Path {
                origin: crate::render::geometry::PtOffset::new(fs.x.resolve(parity), shape_y),
                rotation: fs.rotation,
                flip_h: fs.flip_h,
                flip_v: fs.flip_v,
                extent: fs.size,
                paths: fs.paths.clone(),
                fill: fs.fill.clone(),
                stroke: fs.stroke.clone(),
                effects: fs.effects.clone(),
            });
            emit_shape_text(state, fs, shape_y);
            if y_end > state.cursor_y {
                state.cursor_y = y_end;
            }
        } else {
            let float_entry = float::ActiveFloat {
                page_x: fs.x.resolve(parity) - fs.dist_left,
                page_y_start: y_start,
                page_y_end: y_end,
                width: fs.size.width + fs.dist_left + fs.dist_right,
                source: float::FloatSource::Shape,
                wrap_text: fs.wrap_mode.wrap_text().into(),
            };
            log::debug!(
                "[layout]   register shape float: x={:.1} y={:.1}-{:.1} w={:.1} mode={:?}",
                float_entry.page_x.raw(),
                y_start.raw(),
                y_end.raw(),
                float_entry.width.raw(),
                fs.wrap_mode
            );
            state.page_floats.push(float_entry);
        }
    }
}

fn register_destination_paragraph_floats(
    state: &mut PageLayoutState<'_>,
    floating_images: &[FloatingImage],
    floating_shapes: &[FloatingShape],
    space_before: Pt,
    col_width: Pt,
) {
    let content_top = state.cursor_y + space_before;
    register_paragraph_floats(state, floating_images, floating_shapes, content_top);
    for active_float in &state.page_floats {
        if active_float.overlaps_y(state.cursor_y) && active_float.width >= col_width {
            state.cursor_y = state.cursor_y.max(active_float.page_y_end);
        }
    }
    float::prune_floats(&mut state.page_floats, state.cursor_y);
}

fn has_absolute_wrap_float(floating_images: &[FloatingImage]) -> bool {
    floating_images
        .iter()
        .any(|image| image.y.is_page_absolute() && image.wrap_mode.registers_as_wrap_float())
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ForwardScanBoundary {
    None,
    BeforeParagraphFloat,
    AfterParagraphFloat,
}

fn scan_inline_page_boundary(
    fragments: &[super::super::fragment::Fragment],
    current_col: &mut usize,
    num_cols: usize,
) -> ForwardScanBoundary {
    for (index, fragment) in fragments.iter().enumerate() {
        match fragment {
            super::super::fragment::Fragment::PageBreak { .. } => {
                let has_following_content = fragments[index + 1..].iter().any(|fragment| {
                    !matches!(fragment, super::super::fragment::Fragment::PageBreak { .. })
                });
                return if has_following_content {
                    ForwardScanBoundary::BeforeParagraphFloat
                } else {
                    ForwardScanBoundary::AfterParagraphFloat
                };
            }
            super::super::fragment::Fragment::ColumnBreak => {
                if *current_col + 1 < num_cols {
                    *current_col += 1;
                } else {
                    *current_col = 0;
                    return ForwardScanBoundary::BeforeParagraphFloat;
                }
            }
            _ => {}
        }
    }

    ForwardScanBoundary::None
}

fn mark_absolute_float_relocation(
    paragraph_content_placed: bool,
    moves_to_new_page: bool,
    block_idx: usize,
    replay_block_idx: usize,
    floating_images: &[FloatingImage],
    relocated_blocks: &mut std::collections::HashSet<usize>,
) -> bool {
    !paragraph_content_placed
        && moves_to_new_page
        && block_idx > replay_block_idx
        && has_absolute_wrap_float(floating_images)
        && relocated_blocks.insert(block_idx)
}

fn paragraph_keep_next(block: &LayoutBlock) -> bool {
    matches!(block, LayoutBlock::Paragraph { style, .. } if style.keep_next)
}

/// §17.3.1.15: does a keepNext run *start* at `block_idx`?
///
/// True when the block itself is keepNext and either it is the first block,
/// it forces a `pageBreakBefore`, or the previous block is not keepNext — any
/// of those means nothing above it is chained to it.
fn starts_keep_next_chain(blocks: &[LayoutBlock], block_idx: usize) -> bool {
    paragraph_keep_next(&blocks[block_idx])
        && (block_idx == 0
            || matches!(
                blocks[block_idx],
                LayoutBlock::Paragraph {
                    page_break_before: true,
                    ..
                }
            )
            || !paragraph_keep_next(&blocks[block_idx - 1]))
}

/// Walk a keepNext chain forward from `start` looking for a non-floating table
/// that terminates it.
///
/// Returns `None` if the chain hits a `pageBreakBefore`, a paragraph that
/// isn't itself keepNext, or a *floating* table — a floating table is
/// positioned independently of the flow (§17.4.57 `tblpPr`) and so cannot
/// anchor a chain the way an in-flow table can. `None` here means the chain
/// either ends in an ordinary paragraph or doesn't terminate in something this
/// function can peel around; the caller falls back to measuring the whole
/// paragraph run instead.
fn keep_next_terminal_table(blocks: &[LayoutBlock], start: usize) -> Option<&LayoutBlock> {
    let mut index = start;

    while let Some(block) = blocks.get(index) {
        match block {
            LayoutBlock::Paragraph {
                style,
                page_break_before,
                ..
            } => {
                if index > start && *page_break_before {
                    return None;
                }
                if !style.keep_next {
                    return None;
                }
                index += 1;
            }
            LayoutBlock::Table {
                float_info: None, ..
            } => return Some(block),
            LayoutBlock::Table {
                float_info: Some(_),
                ..
            } => return None,
        }
    }

    None
}

fn fresh_page_contextual_space_before(
    block: &LayoutBlock,
    previous_style_id: &Option<StyleId>,
) -> Pt {
    let LayoutBlock::Paragraph { style, .. } = block else {
        return Pt::ZERO;
    };
    let effective = style.clone_for_layout();
    if effective.contextual_spacing
        && effective.style_id.is_some()
        && effective.style_id.as_ref() == previous_style_id.as_ref()
    {
        effective.space_before
    } else {
        Pt::ZERO
    }
}

struct KeepNextGroupMeasurement {
    body_height: Pt,
    footnote_height: Pt,
    has_footnotes: bool,
}

impl KeepNextGroupMeasurement {
    fn total_height(&self, separator_already_reserved: bool) -> Pt {
        self.body_height
            + self.footnote_height
            + if self.has_footnotes && !separator_already_reserved {
                FOOTNOTE_SEPARATOR_GAP
            } else {
                Pt::ZERO
            }
    }
}

/// Pre-measure a keepNext chain starting at `start`, so the fit decision at
/// the call site is made once against a single number rather than re-measured
/// per candidate. The group ends at the first block that is a table, or a
/// paragraph that is not itself keepNext (its `page_break_before` already
/// having been checked at the previous iteration); `None` if a
/// `pageBreakBefore` interrupts the chain before it can close.
///
/// An oversized group that fits on neither the current nor a fresh page falls
/// through to ordinary in-line placement at the call site rather than being
/// specially handled here — the pagination loop always makes progress this
/// way, the same guarantee [`decide_paragraph_split`] gives the plain
/// (non-keepNext) splitting path.
fn measure_keep_next_group(
    blocks: &[LayoutBlock],
    start: usize,
    constraints: &BoxConstraints,
    default_line_height: Pt,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
) -> Option<KeepNextGroupMeasurement> {
    let mut measurement = KeepNextGroupMeasurement {
        body_height: Pt::ZERO,
        footnote_height: Pt::ZERO,
        has_footnotes: false,
    };
    let mut previous_space_after = Pt::ZERO;
    let mut previous_style_id = None;
    let mut index = start;

    while let Some(block) = blocks.get(index) {
        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                page_break_before,
                footnotes,
                ..
            } => {
                if index > start && *page_break_before {
                    return None;
                }
                let effective = style.clone_for_layout();
                let collapsed = if effective.contextual_spacing
                    && effective.style_id.is_some()
                    && effective.style_id == previous_style_id
                {
                    previous_space_after + effective.space_before
                } else {
                    previous_space_after.min(effective.space_before)
                };
                let layout = layout_paragraph(
                    fragments,
                    constraints,
                    &effective,
                    default_line_height,
                    measure_text,
                );
                measurement.body_height += layout.size.height - collapsed;
                for (footnote_fragments, footnote_style) in footnotes {
                    let footnote = layout_paragraph(
                        footnote_fragments,
                        &BoxConstraints::tight_width(constraints.max_width, Pt::INFINITY),
                        footnote_style,
                        default_line_height,
                        measure_text,
                    );
                    measurement.footnote_height += footnote.size.height;
                    measurement.has_footnotes = true;
                }
                previous_space_after = effective.space_after;
                previous_style_id = effective.style_id.clone();
                if !effective.keep_next {
                    return Some(measurement);
                }
                index += 1;
            }
            LayoutBlock::Table { .. } => {
                return Some(measurement);
            }
        }
    }
    None
}

/// §17.3.1.15 (keepNext tightening): can the paragraph that *starts* a keepNext
/// chain be split so a widow-legal head stays on the current page while its tail
/// travels to the fresh page with the rest of the group?
///
/// keepNext forbids a page break only *between* a paragraph and the next block —
/// it does not pin the paragraph's earlier lines. So when the whole group would
/// otherwise be moved to a fresh page (leaving the current page under-filled),
/// a splittable leading paragraph can instead fill the current page and carry
/// its `>= 2`-line tail onward: the caller has already checked the whole group
/// fits on a fresh page, and the remainder after peeling a `>= 2`-line head is
/// strictly smaller, so it still fits there and no keepNext boundary is broken.
///
/// True only for a body paragraph that can actually split — no keepLines
/// (§17.3.1.14), no floating objects (they anchor to one page), footnotes only
/// on an unbroken chunk, and enough lines to leave `>= 2` on each side under
/// §17.3.1.44 widow control. A `false` result keeps the conservative
/// whole-group move.
fn leading_keep_next_paragraph_splittable(
    block: &LayoutBlock,
    constraints: &BoxConstraints,
    default_line_height: Pt,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
) -> bool {
    let LayoutBlock::Paragraph {
        fragments,
        style,
        footnotes,
        floating_images,
        floating_shapes,
        ..
    } = block
    else {
        return false;
    };
    let effective = style.clone_for_layout();
    // `single_chunk` is a whole-paragraph property, so it can be derived here
    // exactly as the placement gate derives it from its loop state: both
    // splitters are pure functions of the fragment list, and when the paragraph
    // has more than one page chunk the flag is false regardless of column
    // chunking.
    let single_chunk =
        split_at_page_breaks(fragments).len() == 1 && split_at_column_breaks(fragments).len() == 1;
    if !paragraph_breakable(
        &effective,
        footnotes,
        floating_images,
        floating_shapes,
        single_chunk,
    ) {
        return false;
    }
    let placed = place_paragraph(
        fragments,
        constraints,
        &effective,
        default_line_height,
        measure_text,
    );
    // Widow control (§17.3.1.44) needs `>= 2` lines on each side of the break;
    // without it a single-line head is legal. Deliberately stricter than the
    // placement gate's flat `>= 2`: this predicts whether a split would yield a
    // *useful* head, and under-predicting only costs a conservative move.
    let min_lines = if effective.widow_control { 4 } else { 2 };
    placed.line_count() >= min_lines
}

/// §17.3.1.14 / §17.3.1.15: the conditions under which a paragraph may be
/// broken across a page or column boundary, shared by the placement gate
/// (`can_split`) and the keepNext predictor above.
///
/// Extracted because the two disagreed: the predictor omitted the footnote
/// condition, so a keepNext-chain head carrying footnotes across several
/// page/column chunks was reported splittable, the whole-group move was
/// skipped, and the placement path then refused to split it and fell back to
/// atomic — legal output on an under-filled page, but not the pagination either
/// side intended.
///
/// Line counts are **not** included. The gate needs `>= 2`; the predictor needs
/// enough to leave a legal head *and* tail. Folding them together would either
/// weaken the predictor or tighten the gate, so each keeps its own.
fn paragraph_breakable(
    style: &crate::render::layout::paragraph::ParagraphStyle,
    footnotes: &[(
        Vec<Fragment>,
        crate::render::layout::paragraph::ParagraphStyle,
    )],
    floating_images: &[FloatingImage],
    floating_shapes: &[FloatingShape],
    single_chunk: bool,
) -> bool {
    if style.keep_lines || !floating_images.is_empty() || !floating_shapes.is_empty() {
        return false;
    }
    // Footnotes are reserved per segment, but only for a single unbroken chunk:
    // with explicit page/column breaks a reference's segment is ambiguous, so
    // those paragraphs keep the atomic reservation.
    footnotes.is_empty() || single_chunk
}

/// §17.3.1.14 / §17.3.1.44: how to break a paragraph across a page boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParagraphSplit {
    /// Emit all lines on the current page (they fit, or must overflow because no
    /// legal split exists and the paragraph already starts at the page top).
    All,
    /// Emit the first `head` lines here and carry the remaining `total - head`
    /// to the next page. Invariants: `1 <= head < total`; under widow control,
    /// additionally `head >= 2 && total - head >= 2`.
    Break { head: usize },
    /// No legal break at this fill level, and the paragraph does not start at the
    /// page top — move the whole paragraph to the next page and re-decide there.
    MoveWhole,
}

/// How many of a placement's `total` lines fit in `avail`.
///
/// Three things are charged, and which ones apply to which line is the whole
/// of the rule:
///
/// * `space_before` — once, ahead of the first line, and only when this is the
///   paragraph's first segment (§17.3.1.33).
/// * each line's own height.
/// * `trailing_extra` — once, with the **last** line, and only if this
///   placement carries it.
///
/// What is deliberately *not* charged is the paragraph's `space_after`; see
/// the call site for why (issue #126). `trailing_extra` today is the bottom
/// border's space, which is drawn and so must genuinely fit.
fn lines_that_fit(
    total: usize,
    line_height: impl Fn(usize) -> Pt,
    space_before: Pt,
    trailing_extra: Pt,
    avail: Pt,
) -> usize {
    let mut used = space_before;
    let mut n_fit = 0;
    for i in 0..total {
        let mut needed = used + line_height(i);
        if i + 1 == total {
            needed += trailing_extra;
        }
        if needed > avail {
            break;
        }
        used += line_height(i);
        n_fit += 1;
    }
    n_fit
}

/// Decide how a paragraph whose first `n_fit` of `total` lines fit in the
/// remaining space should break (§17.3.1.14 keepLines is handled by the caller,
/// which only calls this for splittable paragraphs).
///
/// `widow_control` (§17.3.1.44) forbids leaving a single line stranded: a legal
/// break keeps `>= 2` lines on each side. `at_page_top` is true when the
/// paragraph begins at the very top of the page column; there, the remaining
/// space already equals a full page, so "move whole" cannot help — a paragraph
/// taller than a page must split ("where possible"), and one that cannot split
/// legally is emitted whole and allowed to overflow. This guarantees the stacker
/// always makes progress (never loops).
fn decide_paragraph_split(
    n_fit: usize,
    total: usize,
    widow_control: bool,
    at_page_top: bool,
) -> ParagraphSplit {
    debug_assert!(total >= 1, "a paragraph has at least one line");
    if n_fit >= total {
        return ParagraphSplit::All;
    }

    // The paragraph does not fully fit; find the largest legal head.
    let head = if widow_control {
        // Leave >= 2 lines for the tail (widow) and keep >= 2 here (orphan).
        let capped = n_fit.min(total.saturating_sub(2));
        if capped >= 2 {
            capped
        } else {
            0 // no widow/orphan-legal break at this fill level
        }
    } else {
        // Without widow control a single line on either side is allowed; place
        // as many as fit, leaving at least one for the tail (n_fit < total, so
        // n_fit <= total - 1 already holds).
        n_fit
    };

    if head >= 1 {
        return ParagraphSplit::Break { head };
    }

    // No legal break at this fill level.
    if at_page_top {
        // Remaining space is a full page and the paragraph still cannot fit or
        // split legally — emit it whole and let it overflow.
        ParagraphSplit::All
    } else {
        ParagraphSplit::MoveWhole
    }
}

/// §17.11.23: reserve `footnotes` on the current page — measure each, subtract
/// its height (and the separator gap for the first footnote on the page) from
/// the available bottom, and queue it for rendering. Shared by the atomic
/// placement path and the per-segment split path.
fn reserve_footnotes<'doc>(
    state: &mut PageLayoutState<'doc>,
    ctx: &LayoutCtx<'_>,
    content_width: Pt,
    footnotes: &'doc [(Vec<Fragment>, ParagraphStyle)],
) {
    if footnotes.is_empty() {
        return;
    }
    let fn_constraints = BoxConstraints::tight_width(content_width, Pt::INFINITY);
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
            state.bottom -= FOOTNOTE_SEPARATOR_GAP;
        }
        state.bottom -= fn_para.size.height;
        state.page_footnotes.push((fn_frags, fn_style));
    }
}

/// §17.6.4: advance the split cursor to the next column, or — when the last
/// column is full — to the top of a fresh page. A fresh column, like a fresh
/// page, offers the full column height, so the split fit treats both as "at the
/// column top".
fn advance_column_or_page(
    state: &mut PageLayoutState<'_>,
    block_idx: usize,
    ctx: &LayoutCtx<'_>,
    num_cols: usize,
    para_start_y: &mut Pt,
) {
    if state.current_col + 1 < num_cols {
        state.current_col += 1;
        state.cursor_y = state.column_top;
    } else {
        state.push_new_page(block_idx, ctx);
    }
    *para_start_y = state.cursor_y;
}

/// Place a splittable paragraph, breaking its lines across pages as needed.
///
/// The caller establishes splittability — see the `can_split` gate at the call
/// site, which is the authority. It requires: no §17.3.1.14 `keepLines`, no
/// floating images or shapes (those anchor to one page), `>= 2` fitted lines,
/// and footnotes only within a single unbroken chunk (with explicit
/// page/column breaks a reference→segment mapping is ambiguous, so those keep
/// the atomic reservation).
///
/// Deliberately *not* required — each was allowed by a later change and this
/// list is the historical trip hazard:
/// - **Borders, shading and drop caps** split fine; they are drawn per segment
///   (`emit_segment_borders_and_shading`).
/// - **Multiple columns** split fine; §17.6.4 unequal-width columns work
///   because each segment re-fits against its own column's width.
///
/// Each iteration fits as many remaining lines as the current page holds,
/// applies §17.3.1.44 widow/orphan control via [`decide_paragraph_split`],
/// emits that segment, and page-breaks to continue. `line_start` advances by
/// `>= 1` on every emitted segment and a `MoveWhole` always lands on a fresh
/// page where progress is forced, so the loop terminates.
#[allow(clippy::too_many_arguments)] // stacker state + placement inputs are cohesive
fn emit_split_paragraph<'doc>(
    state: &mut PageLayoutState<'doc>,
    fragments: &[Fragment],
    style: &ParagraphStyle,
    block_idx: usize,
    config: &PageConfig,
    ctx: &LayoutCtx<'_>,
    para_start_y: &mut Pt,
    footnotes: &'doc [(Vec<Fragment>, ParagraphStyle)],
    content_width: Pt,
    blocks: &[LayoutBlock],
    relocated_absolute_float_blocks: &std::collections::HashSet<usize>,
) {
    let num_cols = config.num_columns();
    let col_x = |col: usize| config.margins.left + config.columns[col].x_offset;
    let widow_control = style.widow_control;

    // §4: the remainder to place on the current page, plus the style to place
    // it with. Both are re-derived per page/column so a continuation wraps
    // around whatever floats live on *its* page (not the paragraph's starting
    // page) and uses that page's width. Owned so they can be re-sliced/re-styled
    // across page breaks. The first page reuses the caller's `style` (which
    // already carries the starting page's floats/width).
    let mut remaining: Vec<Fragment> = fragments.to_vec();
    let mut cont_style: ParagraphStyle = style.clone();
    let mut first_segment = true;
    // §17.11.12: footnotes reserved in document order; this many are placed.
    let mut footnote_cursor = 0;

    loop {
        let col_width = config.columns[state.current_col].width;
        let page_height = (state.bottom - state.page_top).max(Pt::ZERO);
        let constraints = BoxConstraints::new(Pt::ZERO, col_width, Pt::ZERO, page_height);
        let placed = place_paragraph(
            &remaining,
            &constraints,
            &cont_style,
            ctx.default_line_height,
            ctx.measure_text,
        );
        let total = placed.line_count();
        if total == 0 {
            break;
        }

        // §17.6.4: a fresh column offers full height, like a fresh page.
        //
        // §17.6.22: deliberately the raw test, *not*
        // `at_full_height_column_top`. An inherited column really is shorter,
        // but the branch this feeds trades differently from the two block-move
        // sites that do use the predicate. Here `at_page_top` selects "emit the
        // whole paragraph and let it overflow" over "move it", and the overflow
        // it admits is bounded.
        let at_page_top = state.cursor_y <= state.column_top;
        let avail = (state.bottom - state.cursor_y).max(Pt::ZERO);
        // §17.3.1.24/§17.3.1.33: the bottom border space is spent once, on the
        // segment carrying the paragraph's last line.
        //
        // **`space_after` is deliberately not here** (issue #126). Trailing
        // spacing is discarded at a fragmentation boundary — it exists to
        // separate this paragraph from the next one, and at a page or column
        // bottom there is nothing to separate it from, so requiring it to fit
        // rejects a line that is perfectly visible. That is the whole of #126:
        // a 16.85pt line whose bottom cleared the content limit by 6.76pt was
        // pushed to the next page because the paragraph's own 10pt `w:after`
        // did not also fit, stranding it there alone. Word and LibreOffice
        // both keep the line and drop the space.
        //
        // The bottom border stays charged, because unlike whitespace it is
        // drawn: a border that does not fit would be clipped or land in the
        // margin.
        //
        // This only relaxes the *fit* test. The cursor still advances by the
        // full `space_after` below, so content following on the same page is
        // pushed down exactly as before, and the inter-paragraph gap is still
        // charged when deciding whether the *next* paragraph fits — which is
        // what correctly keeps a following paragraph off a page whose
        // remaining room is smaller than the gap plus one line.
        let trailing_extra = placed.bottom_border_space();

        // Count how many of this placement's lines fit, charging space_before
        // (first segment only, §17.3.1.33) and the trailing spacing on the last.
        let n_fit = lines_that_fit(
            total,
            |i| placed.line_height(i),
            if first_segment {
                cont_style.space_before
            } else {
                Pt::ZERO
            },
            trailing_extra,
            avail,
        );

        // §17.3.1.44 widow/orphan, then §17.3.1.11/§17.3.1.44 unbreakable-prefix
        // clamp (drop cap / float-wrapped run kept intact on the first segment).
        let head = match decide_paragraph_split(n_fit, total, widow_control, at_page_top) {
            ParagraphSplit::All => Some(total),
            ParagraphSplit::Break { head } => Some(head),
            ParagraphSplit::MoveWhole => None,
        }
        .and_then(|head| {
            prefix_adjusted_head(
                head,
                total,
                placed.unbreakable_prefix_lines(),
                first_segment,
                at_page_top,
            )
        });

        // No legal placement here: move the whole remainder to the next
        // column/page and re-fit at the top (nothing emitted, so the drop cap
        // and space_before are preserved for the real first segment).
        let Some(head) = head else {
            drop(placed);
            advance_column_or_page(state, block_idx, ctx, num_cols, para_start_y);
            cont_style = continuation_style(
                state,
                blocks,
                style,
                config,
                relocated_absolute_float_blocks,
                first_segment,
            );
            continue;
        };

        let is_last = head == total;
        let segment = placed.emit_split_segment(0, head, first_segment, is_last);
        if !(first_segment && is_last) {
            log::debug!(
                "[layout]   paragraph split: {head}/{total} lines at y={:.1}",
                state.cursor_y.raw()
            );
        }
        let segment_x = col_x(state.current_col);
        for mut cmd in segment.commands {
            cmd.shift_y(state.cursor_y);
            cmd.shift_x(segment_x);
            state.current_page.commands.push(cmd);
        }
        state.cursor_y += segment.size.height;

        // §17.11.12: reserve this segment's footnotes on the page its reference
        // marks landed on, before any page break.
        if !footnotes.is_empty() {
            let refs = placed.footnote_refs_in(0, head);
            let end = (footnote_cursor + refs).min(footnotes.len());
            reserve_footnotes(state, ctx, content_width, &footnotes[footnote_cursor..end]);
            footnote_cursor = end;
        }

        if is_last {
            break;
        }

        // §4: carry the unplaced fragments to the next column/page and re-fit
        // them there against that page's floats and width.
        let next_remaining = placed.fragments_from_line(head);
        drop(placed);
        first_segment = false;
        advance_column_or_page(state, block_idx, ctx, num_cols, para_start_y);
        cont_style = continuation_style(
            state,
            blocks,
            style,
            config,
            relocated_absolute_float_blocks,
            first_segment,
        );
        remaining = next_remaining;
    }
}

/// Build the paragraph style for a split segment after advancing to a new
/// column/page (§4 continuation re-fit). Refreshes the float set for the new
/// page (`effective_floats_at_cursor`, which may also advance the cursor past a
/// full-width float) and repositions the paragraph. When something has already
/// been emitted (`first_segment == false`), the continuation drops
/// `space_before` (§17.3.1.33), the drop cap (§17.3.1.11), and the first-line
/// indent (§17.3.1.12); a whole-paragraph move that emitted nothing keeps them
/// for the eventual real first segment.
fn continuation_style(
    state: &mut PageLayoutState<'_>,
    blocks: &[LayoutBlock],
    style: &ParagraphStyle,
    config: &PageConfig,
    relocated_absolute_float_blocks: &std::collections::HashSet<usize>,
    first_segment: bool,
) -> ParagraphStyle {
    let col = state.current_col;
    let col_width = config.columns[col].width;
    let page_x = config.margins.left + config.columns[col].x_offset;
    let space_before = if first_segment {
        style.space_before
    } else {
        Pt::ZERO
    };
    let floats = state.effective_floats_at_cursor(
        blocks,
        relocated_absolute_float_blocks,
        config.num_columns(),
        space_before,
        col_width,
    );
    let mut cs = style.clone();
    if !first_segment {
        cs.drop_cap = None;
        cs.indent_first_line = Pt::ZERO;
        cs.space_before = Pt::ZERO;
    }
    cs.page_floats = floats;
    cs.page_y = state.cursor_y;
    cs.page_x = page_x;
    cs.page_content_width = col_width;
    cs
}

/// Constrain a chosen split `head` so it never falls inside the paragraph's
/// unbreakable prefix (§17.3.1.11 drop cap and/or §17.3.1.44 float-wrapped run,
/// the first `prefix_lines` lines). Breaking there would tear a drop-cap glyph
/// or leave float-narrowed lines to be reused full-width on the next page.
///
/// Returns `Some(head)` — unchanged when it already clears the prefix — or
/// `None` meaning "move the whole paragraph to a fresh page" (a break inside the
/// prefix with room above). At the page top the prefix cannot move any higher,
/// so it is emitted whole (`Some(prefix_lines)`) and allowed to overflow. The
/// guard only applies to the first segment (the prefix lives at the start).
fn prefix_adjusted_head(
    head: usize,
    remaining: usize,
    prefix_lines: usize,
    first_segment: bool,
    at_page_top: bool,
) -> Option<usize> {
    let breaks_prefix =
        first_segment && prefix_lines > 1 && head < prefix_lines && head < remaining;
    if !breaks_prefix {
        return Some(head);
    }
    if at_page_top {
        Some(prefix_lines.min(remaining))
    } else {
        None
    }
}

/// Where a section begins in the document's page sequence.
///
/// The three travel together because they answer one question — which page
/// this section's first page *is*. `continuation` says whether it shares a
/// page with the section before it, `clearance` how much of that page the
/// header and footer take, and `logical_page_base` what §17.10.6 calls it.
pub(crate) struct SectionStart<'a> {
    /// §17.6.22: an in-progress page to continue on, for a `Continuous` break.
    pub continuation: Option<ContinuationState>,
    /// Header/footer clearances, selected per physical page in the section.
    pub clearance: &'a HeaderFooterClearance,
    /// §17.6.22: what becomes of this section's last page.
    pub last_page: LastPageOwner,
    /// §17.10.6: logical number of the section's first page, with
    /// `w:pgNumType/@start` applied. Drives §20.4.3.1 float mirroring.
    pub logical_page_base: usize,
}

/// §17.6.22: who owns a section's last page.
///
/// One decision, one type. As a `bool` beside an `Option<FinalPageBounds>` the
/// two could disagree — "shared but no bounds", "bounds but not shared" — and
/// both spellings mean something different to `finalize`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum LastPageOwner {
    /// Committed to the document like any other page.
    Own,
    /// A following `Continuous` section shares it and owns its header and
    /// footer. `bounds` is that section's `for_page(0)`, known only once this
    /// section's page count is — so it is `None` on the first pass and `Some`
    /// on any relayout.
    SharedWithNext { bounds: Option<FinalPageBounds> },
}

impl LastPageOwner {
    fn bounds_override(self) -> Option<FinalPageBounds> {
        match self {
            LastPageOwner::SharedWithNext { bounds } => bounds,
            LastPageOwner::Own => None,
        }
    }
}

/// §17.6.22: the body bounds one page of a section must use because a following
/// `Continuous` section shares it.
///
/// A page holding a continuous break is drawn with the header and footer of the
/// **last** section on it — see the ownership rule at
/// `render::shared_page_owner_bounds`, which is where it is decided.
///
/// Index and bounds travel as one value because they are one decision: an index
/// without bounds, or bounds without an index, cannot be acted on, and as two
/// `Option`s they could disagree.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) struct FinalPageBounds {
    /// 0-based physical page index within the section.
    pub index: usize,
    /// The succeeding section's `for_page(0)` bounds.
    pub bounds: PageBodyBounds,
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
    let clearance = HeaderFooterClearance::uniform(config);
    layout_section_with_clearance(
        blocks,
        config,
        measure_text,
        separator_indent,
        default_line_height,
        SectionStart {
            continuation,
            clearance: &clearance,
            // Laid out on its own, nothing follows on the last page.
            last_page: LastPageOwner::Own,
            // A section laid out on its own starts at page 1 — the §17.10.6
            // renumbering only exists across a document's section list.
            logical_page_base: 1,
        },
    )
    .pages
}

/// One section's laid-out pages, plus what happened to its last one.
pub(crate) struct SectionLayout {
    /// Pages this section commits. When `tail` is
    /// [`SectionTail::SharedWithNext`] the shared page is **not** here — it
    /// lives inside the tail, so "popped off `pages` and stashed elsewhere" is
    /// not a state a caller can construct or forget to handle.
    pub pages: Vec<LayoutedPage>,
    pub tail: SectionTail,
}

/// §17.6.22: what became of a section's last page.
///
/// `SharedWithNext` is much larger than `Complete` — it carries a whole page.
/// Boxing it would add an allocation per section to shrink a value that exists
/// once per section and is immediately destructured, so the size difference is
/// accepted for the same reason [`LayoutBlock`]'s is.
#[allow(clippy::large_enum_variant)]
pub(crate) enum SectionTail {
    /// Nothing follows on it; every page is in `pages`.
    Complete,
    /// A `Continuous` section continues on it. Carries the page and the
    /// in-progress state that section resumes from.
    SharedWithNext(ContinuationState),
}

/// Lay out a sequence of blocks using header/footer clearances selected for
/// each physical page in the section.
///
/// Per-page rather than one clearance for the whole section because pages can
/// have differently-sized headers/footers (first-page, even/odd, or a
/// `Continuous` section's shared page) — using one clearance for all of them
/// would give the wrong body bounds on every page but the one it was measured
/// from.
pub(crate) fn layout_section_with_clearance(
    blocks: &[LayoutBlock],
    config: &PageConfig,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
    default_line_height: Pt,
    start: SectionStart<'_>,
) -> SectionLayout {
    let SectionStart {
        continuation,
        clearance,
        last_page,
        logical_page_base,
    } = start;
    let content_width = config.content_width();
    let num_cols = config.num_columns();

    let ctx = LayoutCtx {
        config,
        clearance,
        last_page,
        measure_text,
        separator_indent,
        default_line_height,
    };
    let mut state =
        PageLayoutState::new(config, continuation, ctx.page_bounds(0), logical_page_base);

    // Column-aware constraints and x-offset for the current column.
    let col_constraints = |col: usize, page_height: Pt| -> BoxConstraints {
        let col_width = config.columns[col].width;
        BoxConstraints::new(Pt::ZERO, col_width, Pt::ZERO, page_height)
    };
    let col_x = |col: usize| -> Pt { config.margins.left + config.columns[col].x_offset };

    // A forward-scanned absolute float can later move to another page. Keep
    // those owners out of the source page's next scan while replaying it.
    let mut relocated_absolute_float_blocks = std::collections::HashSet::new();
    let mut page_replay_state = PageReplayCheckpoint::capture(&state);
    let mut checkpoint_page_index = state.page_index;
    let mut replay_block_idx = 0;
    let mut block_idx = 0;

    'blocks: while block_idx < blocks.len() {
        let block = &blocks[block_idx];
        // §17.3.3.1: a deferred inline page break from the previous block
        // forces this block onto a new page.
        if state.pending_page_break {
            state.pending_page_break = false;
            if state.cursor_y > state.page_top {
                state.push_new_page(block_idx, &ctx);
                state.prev_space_after = Pt::ZERO;
            }
        }
        refresh_page_replay_checkpoint(
            &state,
            block_idx,
            &mut page_replay_state,
            &mut checkpoint_page_index,
            &mut replay_block_idx,
        );

        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                page_break_before,
                footnotes,
                floating_images,
                floating_shapes,
            } => {
                // §17.3.1.23: force a new page before this paragraph.
                if *page_break_before && state.cursor_y > state.page_top {
                    state.push_new_page(block_idx, &ctx);
                    state.prev_space_after = Pt::ZERO;
                }

                if num_cols == 1
                    && starts_keep_next_chain(blocks, block_idx)
                    && state.page_floats.is_empty()
                {
                    let constraints = col_constraints(
                        state.current_col,
                        (state.bottom - state.page_top).max(Pt::ZERO),
                    );
                    if let Some(group) = measure_keep_next_group(
                        blocks,
                        block_idx,
                        &constraints,
                        ctx.default_line_height,
                        ctx.measure_text,
                    ) {
                        let current_group_height =
                            group.total_height(!state.page_footnotes.is_empty());
                        let current_group_top = match &blocks[block_idx] {
                            LayoutBlock::Paragraph { style, .. }
                                if style.contextual_spacing
                                    && style.style_id.is_some()
                                    && style.style_id == state.prev_style_id =>
                            {
                                state.cursor_y - state.prev_space_after - style.space_before
                            }
                            LayoutBlock::Paragraph { style, .. } => {
                                state.cursor_y - state.prev_space_after.min(style.space_before)
                            }
                            LayoutBlock::Table { .. } => state.cursor_y,
                        };
                        let full_page_height = ctx.page_bounds(state.page_index + 1).height();
                        let fresh_page_group_height = group.total_height(false)
                            - fresh_page_contextual_space_before(
                                &blocks[block_idx],
                                &state.prev_style_id,
                            );
                        let should_move = match keep_next_terminal_table(blocks, block_idx) {
                            Some(LayoutBlock::Table {
                                rows,
                                col_widths,
                                cell_spacing,
                                border_config,
                                ..
                            }) if !rows.is_empty() => {
                                let current_available =
                                    state.bottom - current_group_top - current_group_height;
                                let full_page_available =
                                    full_page_height - fresh_page_group_height;
                                let leading_group_height = measure_leading_table_group_height(
                                    rows,
                                    col_widths,
                                    *cell_spacing,
                                    ctx.default_line_height,
                                    border_config.as_ref(),
                                    ctx.measure_text,
                                    false,
                                );
                                leading_group_height.is_some_and(|height| {
                                    height <= full_page_available && height > current_available
                                })
                            }
                            _ => {
                                // §17.3.1.15 (11.2): the group doesn't fit here
                                // but fits on a fresh page. Rather than move the
                                // whole group, let a splittable leading
                                // paragraph fill this page and carry its
                                // widow-legal tail onward — the remainder is
                                // strictly smaller than the (fresh-page-fitting)
                                // group, so the keepNext boundary stays intact.
                                // Only for paragraph terminals: a table terminal
                                // is excluded from the group measurement, so its
                                // leading row is not covered by that guarantee
                                // and keeps the whole-move (handled above).
                                fresh_page_group_height <= full_page_height
                                    && current_group_top + current_group_height > state.bottom
                                    && !leading_keep_next_paragraph_splittable(
                                        &blocks[block_idx],
                                        &constraints,
                                        ctx.default_line_height,
                                        ctx.measure_text,
                                    )
                            }
                        };
                        if should_move && !state.at_full_height_column_top() {
                            state.push_new_page(block_idx, &ctx);
                            state.prev_space_after = Pt::ZERO;
                        }
                    }
                }
                refresh_page_replay_checkpoint(
                    &state,
                    block_idx,
                    &mut page_replay_state,
                    &mut checkpoint_page_index,
                    &mut replay_block_idx,
                );

                let mut effective_style = style.clone_for_layout();

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
                let float_checkpoint = ParagraphFloatCheckpoint::capture(&state);
                let content_top = state.cursor_y + effective_style.space_before;
                register_paragraph_floats(
                    &mut state,
                    floating_images,
                    floating_shapes,
                    content_top,
                );

                // Prune expired floats.
                float::prune_floats(&mut state.page_floats, state.cursor_y);

                // §20.4.2: floats affecting text at the cursor —
                // registered page floats plus the boundary-/relocation-aware
                // forward scan of upcoming blocks — advancing past any
                // full-width blocker.
                let col_width = config.columns[state.current_col].width;
                let page_x = col_x(state.current_col);
                let effective_floats = state.effective_floats_at_cursor(
                    blocks,
                    &relocated_absolute_float_blocks,
                    num_cols,
                    effective_style.space_before,
                    col_width,
                );

                effective_style.page_floats = effective_floats;
                effective_style.page_y = state.cursor_y;
                effective_style.page_x = page_x;
                effective_style.page_content_width = col_width;

                // §17.3.3.1: split paragraph at inline page breaks first,
                // then §17.6.4: split each page-chunk at column breaks.
                let page_chunks = split_at_page_breaks(fragments);
                let mut para_start_y = state.cursor_y;
                state.last_para_start_y = state.cursor_y;
                // Issue #86 (float relocation): true once any of this paragraph's
                // content has been placed, so a later overflow relocates its
                // absolute float instead of double-wrapping earlier text.
                let mut paragraph_content_placed = false;
                // §17.11.12: set once a split segment reserves this paragraph's
                // footnotes per page, so the atomic after-loop reservation is
                // skipped (avoids double-reserving).
                let mut footnotes_reserved = false;

                // §17.3.3.1: track whether an unresolved page break remains
                // after processing all chunks, so it can be deferred to the
                // next block.
                let mut unresolved_page_break = false;

                for (page_chunk_idx, page_chunk) in page_chunks.iter().enumerate() {
                    if page_chunk_idx > 0 {
                        unresolved_page_break = true;
                    }

                    // §17.3.3.1: skip empty chunks — they carry no renderable
                    // content. A page break whose leading or trailing side is
                    // empty simply means "nothing on this side of the break."
                    if page_chunk.is_empty() {
                        continue;
                    }

                    // §17.3.3.1: force a new page for non-empty chunks that
                    // follow a page break.
                    if unresolved_page_break {
                        if mark_absolute_float_relocation(
                            paragraph_content_placed,
                            true,
                            block_idx,
                            replay_block_idx,
                            floating_images,
                            &mut relocated_absolute_float_blocks,
                        ) {
                            page_replay_state.restore(&mut state);
                            block_idx = replay_block_idx;
                            continue 'blocks;
                        }
                        if !paragraph_content_placed {
                            float_checkpoint.restore(&mut state);
                        }
                        state.push_new_page(block_idx, &ctx);
                        state.prev_space_after = Pt::ZERO;
                        para_start_y = state.cursor_y;
                        if !paragraph_content_placed {
                            let col_width = config.columns[state.current_col].width;
                            register_destination_paragraph_floats(
                                &mut state,
                                floating_images,
                                floating_shapes,
                                effective_style.space_before,
                                col_width,
                            );
                        }
                        effective_style.page_y = state.cursor_y;
                        effective_style.page_x = col_x(state.current_col);
                        effective_style.page_content_width =
                            config.columns[state.current_col].width;
                        effective_style.page_floats = state.page_floats.clone();
                    }

                    let col_chunks = split_at_column_breaks(page_chunk);

                    for (chunk_idx, chunk) in col_chunks.iter().enumerate() {
                        // Advance to the next column for chunks after a column break.
                        if chunk_idx > 0 {
                            let starts_new_page = state.current_col + 1 >= num_cols;
                            if mark_absolute_float_relocation(
                                paragraph_content_placed,
                                starts_new_page,
                                block_idx,
                                replay_block_idx,
                                floating_images,
                                &mut relocated_absolute_float_blocks,
                            ) {
                                page_replay_state.restore(&mut state);
                                block_idx = replay_block_idx;
                                continue 'blocks;
                            }
                            if !paragraph_content_placed && starts_new_page {
                                float_checkpoint.restore(&mut state);
                            }
                            if !starts_new_page {
                                state.current_col += 1;
                            } else {
                                // All columns full — new page, reset to column 0.
                                state.push_new_page(block_idx, &ctx);
                            }
                            state.cursor_y = state.column_top;
                            if !paragraph_content_placed && starts_new_page {
                                let col_width = config.columns[state.current_col].width;
                                register_destination_paragraph_floats(
                                    &mut state,
                                    floating_images,
                                    floating_shapes,
                                    effective_style.space_before,
                                    col_width,
                                );
                            }
                            effective_style.page_y = state.cursor_y;
                            effective_style.page_x = col_x(state.current_col);
                            effective_style.page_content_width =
                                config.columns[state.current_col].width;
                            if starts_new_page {
                                effective_style.page_floats = state.page_floats.clone();
                            }
                        }

                        let constraints = col_constraints(
                            state.current_col,
                            (state.bottom - state.page_top).max(Pt::ZERO),
                        );
                        let placed = place_paragraph(
                            chunk,
                            &constraints,
                            &effective_style,
                            ctx.default_line_height,
                            ctx.measure_text,
                        );

                        // §17.3.1.14: a paragraph may break across a page
                        // boundary when keepLines is unset. Borders/shading, drop
                        // caps, and lines wrapped around an active float are
                        // handled per segment (the float-wrapped run is kept on
                        // the first segment by the unbreakable-prefix guard, so
                        // the tail reuses its fitted lines). Paragraphs that own
                        // floating objects stay atomic (the object anchors to one
                        // page) and take the #86 relocation path in the `else`
                        // branch below. Footnotes are reserved per segment, but
                        // only for a single, unbroken chunk — with explicit
                        // page/column breaks their reference→segment mapping is
                        // ambiguous, so those keep the atomic reservation.
                        // §17.6.4: splitting re-fits the
                        // remainder against each column's own width
                        // (`emit_split_paragraph`), so unequal-width columns split
                        // correctly — no equal-width gate.
                        let single_chunk = page_chunks.len() == 1 && col_chunks.len() == 1;
                        let can_split = paragraph_breakable(
                            &effective_style,
                            footnotes,
                            floating_images,
                            floating_shapes,
                            single_chunk,
                        ) && placed.line_count() >= 2;

                        if can_split {
                            if !footnotes.is_empty() {
                                footnotes_reserved = true;
                            }
                            drop(placed);
                            emit_split_paragraph(
                                &mut state,
                                chunk,
                                &effective_style,
                                block_idx,
                                config,
                                &ctx,
                                &mut para_start_y,
                                footnotes,
                                content_width,
                                blocks,
                                &relocated_absolute_float_blocks,
                            );
                        } else {
                            let mut para = placed.emit_full();
                            // Column/page overflow: advance column, then page.
                            //
                            // Measured **without** the paragraph's trailing
                            // space (issue #126), for the reason given at the
                            // `lines_that_fit` call site: `emit_full`'s height
                            // is `space_before + Σ lines + space_after`, and
                            // charging that last term here rejects a paragraph
                            // whose text fits, over whitespace that a page
                            // bottom discards. The split path had the same
                            // defect; this is the single-segment twin of it.
                            //
                            // `cursor_y` below still advances by the full
                            // height, so anything following on this page is
                            // placed exactly as before.
                            let fit_height = para.size.height - effective_style.space_after;
                            if state.cursor_y + fit_height > state.bottom
                                && !state.at_full_height_column_top()
                            {
                                // `placed` (which borrows `effective_style`) is no
                                // longer needed; drop it before mutating the style
                                // and re-placing at the destination page, whose
                                // max_height differs and re-clamps the height.
                                drop(placed);
                                // #86: if this atomic paragraph owns an absolute
                                // wrap float and moves to a new page before any of
                                // its own content is placed, earlier source-page
                                // text may already have wrapped around that future
                                // float. Replay the page without it before placing
                                // the owner on its destination page.
                                let moves_to_new_page = state.current_col + 1 >= num_cols;
                                if mark_absolute_float_relocation(
                                    paragraph_content_placed,
                                    moves_to_new_page,
                                    block_idx,
                                    replay_block_idx,
                                    floating_images,
                                    &mut relocated_absolute_float_blocks,
                                ) {
                                    page_replay_state.restore(&mut state);
                                    block_idx = replay_block_idx;
                                    continue 'blocks;
                                }
                                if !paragraph_content_placed {
                                    float_checkpoint.restore(&mut state);
                                }
                                if state.current_col + 1 < num_cols {
                                    state.current_col += 1;
                                    state.cursor_y = state.column_top;
                                } else {
                                    state.push_new_page(block_idx, &ctx);
                                }
                                let destination_para_start_y = state.cursor_y;
                                // Re-register this paragraph's own floats at the
                                // destination when none of its content has landed
                                // yet (§20.4.2).
                                if !paragraph_content_placed {
                                    let col_width = config.columns[state.current_col].width;
                                    register_destination_paragraph_floats(
                                        &mut state,
                                        floating_images,
                                        floating_shapes,
                                        effective_style.space_before,
                                        col_width,
                                    );
                                }
                                // Update para_start_y after page/column change so
                                // floating images use the correct position.
                                para_start_y = destination_para_start_y;
                                effective_style.page_y = state.cursor_y;
                                effective_style.page_x = col_x(state.current_col);
                                effective_style.page_content_width =
                                    config.columns[state.current_col].width;
                                effective_style.page_floats = state.page_floats.clone();
                                let destination_constraints = col_constraints(
                                    state.current_col,
                                    (state.bottom - state.page_top).max(Pt::ZERO),
                                );
                                para = place_paragraph(
                                    chunk,
                                    &destination_constraints,
                                    &effective_style,
                                    ctx.default_line_height,
                                    ctx.measure_text,
                                )
                                .emit_full();
                            }

                            log::debug!(
                                "[layout]   page_chunk[{page_chunk_idx}] col_chunk[{chunk_idx}] placed at y={:.1} x={:.1} height={:.1}",
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
                        // #86: mark that content of this paragraph has landed
                        // (either split segments or the atomic block), so a
                        // later block's absolute-float overflow relocates rather
                        // than double-wrapping this already-placed text.
                        paragraph_content_placed |= !chunk.is_empty();
                    }
                    // The page break has been consumed by this non-empty chunk.
                    unresolved_page_break = false;
                }

                // §17.3.3.1: if the paragraph ended with a page break and
                // no non-empty chunk followed, defer the break to the next block.
                if unresolved_page_break {
                    state.pending_page_break = true;
                }

                state.first_on_section_page = false;
                state.prev_borders = style.borders.clone();
                state.prev_space_after = effective_style.space_after;
                state.prev_style_id = effective_style.style_id.clone();
                state.prev_table_style_id = None; // paragraph breaks adjacent table chain

                // §20.4.2.3: emit non-wrapTopAndBottom floating images.
                // (wrapTopAndBottom images were emitted immediately above.)
                let parity = state.parity();
                for fi in floating_images {
                    if fi.is_wrap_top_and_bottom() {
                        continue;
                    }
                    let img_y = fi.y.at(parity, para_start_y + effective_style.space_before);
                    state.current_page.commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(
                            fi.x.resolve(parity),
                            img_y,
                            fi.size.width,
                            fi.size.height,
                        ),
                        image_data: fi.image_data.clone(),
                        src_rect: fi.src_rect,
                    });
                }

                // §20.4.2: emit floating DrawingML shapes after the
                // paragraph's text so they paint on top (for `behindDoc=0`).
                // `wrapTopAndBottom` shapes were emitted pre-layout along
                // with their cursor advance — skip them here.
                for fs in floating_shapes {
                    if fs.is_wrap_top_and_bottom() {
                        continue;
                    }
                    let shape_y = fs.y.at(parity, para_start_y + effective_style.space_before);
                    state.current_page.commands.push(DrawCommand::Path {
                        origin: crate::render::geometry::PtOffset::new(
                            fs.x.resolve(parity),
                            shape_y,
                        ),
                        rotation: fs.rotation,
                        flip_h: fs.flip_h,
                        flip_v: fs.flip_v,
                        extent: fs.size,
                        paths: fs.paths.clone(),
                        fill: fs.fill.clone(),
                        stroke: fs.stroke.clone(),
                        effects: fs.effects.clone(),
                    });
                    emit_shape_text(&mut state, fs, shape_y);
                }

                // Collect footnotes for this page and reduce the available
                // bottom. A split paragraph already reserved them per segment
                // (on the page each reference landed on), so skip here.
                if !footnotes_reserved {
                    reserve_footnotes(&mut state, &ctx, content_width, footnotes);
                }
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                cell_spacing,
                border_config,
                indent,
                alignment,
                float_info,
                style_id,
            } => {
                // §17.4.57: floating table — render and register as a float so
                // subsequent text wraps around it. Floating tables are absolutely
                // positioned and do not participate in adjacent border collapse.
                if let Some(fi) = float_info {
                    // Run an un-paginated layout once to get the table's
                    // width (for x positioning + alignment overrides) and
                    // total height (for the §17.4.57 page-push heuristic).
                    // The actual emission uses `layout_table_paginated`
                    // below so rows that overflow split across pages.
                    let table = layout_table(
                        rows,
                        col_widths,
                        *cell_spacing,
                        ctx.default_line_height,
                        border_config.as_ref(),
                        ctx.measure_text,
                        false,
                    );

                    // §17.4.28 / §17.4.50: compute table x position.
                    let table_x = table_x_offset(
                        *alignment,
                        *indent,
                        table.size.width,
                        content_width,
                        config.margins.left,
                    );
                    // §17.4.57: apply tblpXSpec horizontal alignment override.
                    let table_x = match fi.x_align {
                        Some(crate::model::TableXAlign::Center) => {
                            config.margins.left + (content_width - table.size.width) * 0.5
                        }
                        Some(crate::model::TableXAlign::Right) => {
                            config.margins.left + content_width - table.size.width
                        }
                        _ => table_x,
                    };

                    // §17.4.57: anchor-page heuristic — if the cursor isn't
                    // already at top of page and the table won't fit below
                    // it, push the table to the next page before resolving
                    // the anchor. This matches Word's "don't anchor on a
                    // page that's already mostly full" behavior.
                    if state.cursor_y + table.size.height > state.bottom
                        && state.cursor_y > state.page_top
                    {
                        state.push_new_page(block_idx, &ctx);
                        state.prev_space_after = Pt::ZERO;
                    }

                    // §17.4.57: resolve `tblpY` on the (possibly new)
                    // current page, then §17.4.56 resolve collisions with
                    // prior floats. On `Spillover`, push to next page and
                    // re-resolve with the new (empty) float list.
                    //
                    // This loop terminates because `Spillover` requires a
                    // collision-induced shift and `push_new_page` clears
                    // `page_floats`: the retry has no floats to collide
                    // with, so it cannot spill again. A table too tall for
                    // the body comes back as `OnCurrentPage` and is sliced
                    // by the pagination call below, never re-resolved.
                    let float_y_start = loop {
                        let requested_y = if fi.y_offset > Pt::ZERO {
                            let anchor_y = match fi.vert_anchor {
                                crate::model::TableAnchor::Text => {
                                    state.last_para_start_y + fi.y_offset
                                }
                                crate::model::TableAnchor::Margin => state.page_top + fi.y_offset,
                                crate::model::TableAnchor::Page => fi.y_offset,
                            };
                            // The table must not start above the cursor —
                            // preceding content already occupies the space
                            // above it.
                            anchor_y.max(state.cursor_y)
                        } else {
                            state.cursor_y
                        };

                        match resolve_floating_anchor(
                            requested_y,
                            table.size.height,
                            fi.overlap,
                            &state.page_floats,
                            state.bottom,
                        ) {
                            FloatingTableAnchor::OnCurrentPage(y) => break y,
                            FloatingTableAnchor::Shifted { from, to } => {
                                log::debug!(
                                    "[layout]   shift float past prior (overlap=Never): {:.1} -> {:.1}",
                                    from.raw(),
                                    to.raw(),
                                );
                                break to;
                            }
                            FloatingTableAnchor::Spillover => {
                                log::debug!(
                                    "[layout]   float spill to next page (overlap=Never): block_idx={block_idx}",
                                );
                                state.push_new_page(block_idx, &ctx);
                                state.prev_space_after = Pt::ZERO;
                                // Loop: re-resolve on the fresh page (empty
                                // float list, cursor at the selected page top).
                            }
                        }
                    };

                    // Floating table breaks the adjacent table chain.
                    state.prev_table_style_id = None;

                    // §17.4.57: paginate at row boundaries when the table
                    // would overflow. First slice gets the anchor page's
                    // remaining height (`bottom - float_y_start`);
                    // continuation slices get the selected body height for
                    // each subsequent page.
                    let available_first = (state.bottom - float_y_start).max(Pt::ZERO);
                    let section_page_index = state.page_index;
                    let slices = layout_table_paginated_with_page_heights(
                        rows,
                        col_widths,
                        *cell_spacing,
                        ctx.default_line_height,
                        border_config.as_ref(),
                        ctx.measure_text,
                        TablePaginationHeights {
                            available_height: available_first,
                            suppress_first_row_top: false,
                            page_height_for_slice: |slice_index| {
                                ctx.page_bounds(section_page_index + slice_index).height()
                            },
                        },
                    );

                    // §17.4.57: anchor only the first slice; continuation
                    // slices flow at the top of subsequent pages. Encoded
                    // by the `Anchor` / `Continuation` enum variants in
                    // the placement plan.
                    let plan = plan_floating_table_pages_with_page_tops(
                        slices,
                        float_y_start,
                        |slice_index| ctx.page_bounds(section_page_index + slice_index).top,
                    );

                    let table_width = table.size.width;
                    for (page_idx, placement) in plan.pages.into_iter().enumerate() {
                        if page_idx > 0 {
                            state.push_new_page(block_idx, &ctx);
                            state.prev_space_after = Pt::ZERO;
                        }

                        let (y_start, slice, is_anchor) = match placement {
                            FloatingTablePagePlacement::Anchor { y_start, slice } => {
                                (y_start, slice, true)
                            }
                            FloatingTablePagePlacement::Continuation { y_start, slice } => {
                                (y_start, slice, false)
                            }
                        };
                        let slice_height = slice.size.height;

                        for mut cmd in slice.commands {
                            cmd.shift_y(y_start);
                            cmd.shift_x(table_x);
                            state.current_page.commands.push(cmd);
                        }

                        // §17.4.57 / §17.4.56: register every slice as a
                        // float on its respective page. The anchor slice
                        // drives text wrapping for body paragraphs that
                        // follow; continuation slices are registered so
                        // subsequent floating tables can see them during
                        // collision resolution (§17.4.56 `tblOverlap`).
                        log::debug!(
                            "[layout]   register table float ({}): x={:.1} y={:.1}-{:.1} w={:.1} block_idx={block_idx}",
                            if is_anchor { "anchor" } else { "continuation" },
                            table_x.raw(),
                            y_start.raw(),
                            (y_start + slice_height).raw(),
                            (table_width + fi.right_gap).raw(),
                        );
                        state.page_floats.push(float::ActiveFloat {
                            page_x: table_x,
                            page_y_start: y_start,
                            page_y_end: y_start + slice_height,
                            width: table_width + fi.right_gap,
                            source: float::FloatSource::Table {
                                owner_block_idx: block_idx,
                            },
                            // §17.4.57: floating tables default to
                            // bothSides; no dedicated wrapText
                            // attribute exists for tables.
                            wrap_text: float::WrapTextSide::BothSides,
                        });
                        // Suppress unused warning when `is_anchor` is no
                        // longer the discriminant for registration.
                        let _ = is_anchor;
                    }
                    block_idx += 1;
                    continue;
                }

                // §17.4.38: consecutive non-floating tables with the same style
                // are treated as one merged table — the second table's top border
                // is suppressed so the shared edge is drawn once.
                let suppress_top = style_id.is_some() && *style_id == state.prev_table_style_id;

                // Non-floating table: paginated row-level splitting.
                // §17.4.49 / §17.4.6: split at row boundaries, repeat headers.
                let available = state.bottom - state.cursor_y;
                let section_page_index = state.page_index;
                let slices = layout_table_paginated_with_page_heights(
                    rows,
                    col_widths,
                    *cell_spacing,
                    ctx.default_line_height,
                    border_config.as_ref(),
                    ctx.measure_text,
                    TablePaginationHeights {
                        available_height: available,
                        suppress_first_row_top: suppress_top,
                        page_height_for_slice: |slice_index| {
                            ctx.page_bounds(section_page_index + slice_index).height()
                        },
                    },
                );

                // §17.4.28 / §17.4.50: compute table x position.
                //
                // The width to align on is the table's *outer* width, which
                // every slice already reports — §17.4.45 spacing was carved
                // out of the slots by `build/table.rs::reserve_cell_spacing`,
                // so `sum(col_widths)` is the table's width *minus* one
                // `tblCellSpacing` and centring on it lands the table half a
                // spacing off. Every slice spans the same width; take the
                // first (there is always one).
                let table_width = slices.first().map_or(Pt::ZERO, |s| s.size.width);
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
        block_idx += 1;
    }

    // Flush remaining footnotes and push the last page.
    state.finalize(&ctx)
}

/// Issue #126: what a page bottom charges for, and what it does not.
#[cfg(test)]
mod fit_tests {
    use super::*;

    /// Ten 10pt lines, so the arithmetic is readable.
    fn fit(space_before: f32, trailing: f32, avail: f32) -> usize {
        lines_that_fit(
            10,
            |_| Pt::new(10.0),
            Pt::new(space_before),
            Pt::new(trailing),
            Pt::new(avail),
        )
    }

    #[test]
    fn lines_fill_the_available_space() {
        assert_eq!(fit(0.0, 0.0, 100.0), 10, "all ten fit exactly");
        assert_eq!(fit(0.0, 0.0, 99.9), 9, "one short of the tenth");
        assert_eq!(fit(0.0, 0.0, 0.0), 0);
    }

    /// §17.3.1.33: charged once, ahead of the first line — so it costs the
    /// paragraph a line only when it pushes past a boundary.
    #[test]
    fn space_before_is_charged_once_at_the_front() {
        assert_eq!(fit(10.0, 0.0, 100.0), 9, "space_before displaces one line");
        assert_eq!(fit(10.0, 0.0, 110.0), 10);
    }

    /// **The defect behind #126.** The paragraph's own trailing space must not
    /// decide whether its last line is placed — this asserts the rule through
    /// the parameter that survived: with no `trailing_extra`, a last line that
    /// fits on its own is kept.
    ///
    /// Before the fix `trailing_extra` also carried `space_after`, so on the
    /// reporter's document a 16.85pt line whose bottom cleared the content
    /// limit by 6.76pt was rejected for want of 10pt of invisible whitespace,
    /// and ended up alone on a page of its own.
    #[test]
    fn a_last_line_that_fits_is_kept_when_nothing_trails_it() {
        assert_eq!(fit(0.0, 0.0, 100.0), 10, "exactly enough for ten lines");
    }

    /// …and a bottom border still reserves its space, because unlike
    /// whitespace it is drawn. `trailing_extra` is exactly that today.
    #[test]
    fn trailing_extra_still_costs_the_last_line() {
        assert_eq!(fit(0.0, 5.0, 100.0), 9, "the border does not fit with #10");
        assert_eq!(fit(0.0, 5.0, 105.0), 10, "give it room and #10 returns");
    }

    /// It is charged against the **last** line only, not every line: nine
    /// lines plus a border fit in less than ten lines' worth of space.
    #[test]
    fn trailing_extra_is_charged_once_not_per_line() {
        assert_eq!(fit(0.0, 5.0, 95.0), 9, "9 lines + 5pt border = 95");
    }

    /// A paragraph too tall for the space still reports what fits, so the
    /// caller can split it rather than loop.
    #[test]
    fn a_partial_fit_reports_the_lines_that_fit() {
        assert_eq!(fit(0.0, 0.0, 55.0), 5);
        assert_eq!(fit(0.0, 100.0, 55.0), 5, "trailing only affects the last");
    }
}

#[cfg(test)]
mod keep_next_chain_tests {
    use super::*;
    use crate::render::layout::paragraph::ParagraphStyle;

    fn paragraph(keep_next: bool, page_break_before: bool) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: Vec::new(),
            style: ParagraphStyle {
                keep_next,
                ..Default::default()
            },
            page_break_before,
            footnotes: Vec::new(),
            floating_images: Vec::new(),
            floating_shapes: Vec::new(),
        }
    }

    #[test]
    fn page_break_before_starts_a_new_keep_next_chain() {
        let blocks = [paragraph(true, false), paragraph(true, true)];

        assert!(starts_keep_next_chain(&blocks, 1));
    }
}

/// §17.3.1.14 / §17.3.1.15 — `paragraph_breakable`, the predicate shared by the
/// placement gate and the keepNext predictor. Extracted precisely because the
/// two had drifted apart on the footnote condition (E4a#3).
#[cfg(test)]
mod paragraph_breakable_tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::layout::paragraph::ParagraphStyle;

    fn footnote() -> (Vec<Fragment>, ParagraphStyle) {
        (Vec::new(), ParagraphStyle::default())
    }

    fn plain() -> ParagraphStyle {
        ParagraphStyle::default()
    }

    #[test]
    fn a_plain_paragraph_is_breakable() {
        assert!(paragraph_breakable(&plain(), &[], &[], &[], true));
        assert!(paragraph_breakable(&plain(), &[], &[], &[], false));
    }

    #[test]
    fn keep_lines_forbids_breaking() {
        let style = ParagraphStyle {
            keep_lines: true,
            ..Default::default()
        };
        assert!(!paragraph_breakable(&style, &[], &[], &[], true));
    }

    /// **The condition the keepNext predictor was missing.** Footnotes are
    /// reserved per segment, but only for a single unbroken chunk — with
    /// explicit page/column breaks a reference's segment is ambiguous.
    ///
    /// Both directions matter: footnotes on one chunk stay breakable (or every
    /// footnote-bearing paragraph would become atomic), and footnotes across
    /// several chunks do not.
    #[test]
    fn footnotes_are_breakable_only_on_a_single_chunk() {
        let notes = [footnote()];
        assert!(
            paragraph_breakable(&plain(), &notes, &[], &[], true),
            "footnotes on one unbroken chunk may still split"
        );
        assert!(
            !paragraph_breakable(&plain(), &notes, &[], &[], false),
            "footnotes spanning page/column chunks keep the atomic reservation"
        );
    }

    /// `single_chunk` is irrelevant without footnotes — it must not become a
    /// blanket restriction on multi-chunk paragraphs.
    #[test]
    fn single_chunk_only_matters_when_footnotes_are_present() {
        assert!(paragraph_breakable(&plain(), &[], &[], &[], false));
    }

    // ── The predictor's own derivation of `single_chunk` ─────────────────

    use crate::render::geometry::PtSize;
    use crate::render::layout::fragment::{FontProps, TextMetrics};
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after: crate::render::layout::fragment::fixture_break_after(text),
            font: Rc::new(FontProps {
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
                strike_lines: 0,
                strike_position: Pt::ZERO,
                strike_thickness: Pt::ZERO,
            }),
            color: RgbColor::BLACK,
            width: Pt::new(30.0),
            trimmed_width: Pt::new(30.0),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    fn page_break() -> Fragment {
        Fragment::PageBreak {
            line_height: Pt::new(14.0),
        }
    }

    /// Six wrapping lines, optionally interrupted by `brk`, optionally carrying
    /// a footnote.
    fn keep_next_head(brk: Option<Fragment>, with_footnote: bool) -> LayoutBlock {
        let mut fragments: Vec<Fragment> = (0..3).map(|i| text_frag(&format!("L{i} "))).collect();
        if let Some(b) = brk {
            fragments.push(b);
        }
        fragments.extend((3..6).map(|i| text_frag(&format!("L{i} "))));
        LayoutBlock::Paragraph {
            fragments,
            style: ParagraphStyle {
                keep_next: true,
                ..Default::default()
            },
            page_break_before: false,
            footnotes: if with_footnote {
                vec![footnote()]
            } else {
                Vec::new()
            },
            floating_images: Vec::new(),
            floating_shapes: Vec::new(),
        }
    }

    fn splittable(block: &LayoutBlock) -> bool {
        // Width 40 against 30pt fragments → one line each.
        let constraints = BoxConstraints::loose(PtSize::new(Pt::new(40.0), Pt::new(1000.0)));
        leading_keep_next_paragraph_splittable(block, &constraints, Pt::new(14.0), None)
    }

    /// The predictor must derive `single_chunk` itself, not assume it.
    ///
    /// A keepNext-chain head carrying footnotes **across a page break** is not
    /// splittable: the placement gate will refuse it, so reporting it splittable
    /// skips the whole-group move and then falls back to atomic anyway — an
    /// under-filled page. Hard-coding `single_chunk = true` here is precisely
    /// the bug this closes, and only this test sees it.
    #[test]
    fn footnote_bearing_head_with_an_internal_break_is_not_splittable() {
        // §17.3.3.1 page break and §17.6.3 column break both end a chunk, so
        // both must be consulted — the derivation is an `&&` of two splitters
        // and testing only one leaves half of it unverified.
        for (label, brk) in [
            ("page break", page_break()),
            ("column break", Fragment::ColumnBreak),
        ] {
            assert!(
                !splittable(&keep_next_head(Some(brk), true)),
                "footnotes + an internal {label} → atomic"
            );
        }
    }

    /// The controls, so the test above can't pass for the wrong reason: each
    /// ingredient alone still permits a split.
    #[test]
    fn a_break_or_a_footnote_alone_still_permits_a_split() {
        assert!(
            splittable(&keep_next_head(None, true)),
            "footnotes on a single chunk are fine"
        );
        assert!(
            splittable(&keep_next_head(Some(page_break()), false)),
            "a page break without footnotes is fine"
        );
        assert!(
            splittable(&keep_next_head(Some(Fragment::ColumnBreak), false)),
            "a column break without footnotes is fine"
        );
        assert!(
            splittable(&keep_next_head(None, false)),
            "plain multi-line head is splittable"
        );
    }
}

#[cfg(test)]
mod paragraph_split_tests {
    use super::{decide_paragraph_split, prefix_adjusted_head, ParagraphSplit};

    // ── Whole paragraph fits ──────────────────────────────────────────────

    #[test]
    fn all_lines_fit_is_all() {
        assert_eq!(
            decide_paragraph_split(5, 5, true, false),
            ParagraphSplit::All
        );
        // More space than needed also counts as fitting.
        assert_eq!(
            decide_paragraph_split(9, 5, true, false),
            ParagraphSplit::All
        );
    }

    // ── Widow/orphan control on (§17.3.1.44) ──────────────────────────────

    #[test]
    fn widow_control_keeps_two_lines_on_each_side() {
        // 6 lines, 4 fit: place 4, carry 2 — both sides satisfy the >= 2 rule.
        assert_eq!(
            decide_paragraph_split(4, 6, true, false),
            ParagraphSplit::Break { head: 4 }
        );
    }

    #[test]
    fn widow_control_caps_head_to_leave_a_non_widow_tail() {
        // 4 lines, 3 fit: placing 3 would strand 1 (widow), so cap head to 2.
        assert_eq!(
            decide_paragraph_split(3, 4, true, false),
            ParagraphSplit::Break { head: 2 }
        );
    }

    #[test]
    fn widow_control_rejects_single_orphan_line_and_moves_whole() {
        // Only 1 line fits: an orphan. Not at page top → move the whole para.
        assert_eq!(
            decide_paragraph_split(1, 6, true, false),
            ParagraphSplit::MoveWhole
        );
    }

    #[test]
    fn widow_control_three_line_paragraph_cannot_split() {
        // No head in 1..=2 leaves >= 2 on both sides → move whole.
        assert_eq!(
            decide_paragraph_split(2, 3, true, false),
            ParagraphSplit::MoveWhole
        );
    }

    #[test]
    fn widow_control_paragraph_taller_than_page_splits_two_by_two() {
        // 10 lines, 4 fit per page: 4 / 4 / 2 across three pages.
        assert_eq!(
            decide_paragraph_split(4, 10, true, true),
            ParagraphSplit::Break { head: 4 }
        );
        assert_eq!(
            decide_paragraph_split(4, 6, true, true),
            ParagraphSplit::Break { head: 4 }
        );
        assert_eq!(
            decide_paragraph_split(4, 2, true, true),
            ParagraphSplit::All
        );
    }

    // ── Widow control off ─────────────────────────────────────────────────

    #[test]
    fn without_widow_control_a_single_line_may_split() {
        assert_eq!(
            decide_paragraph_split(1, 6, false, false),
            ParagraphSplit::Break { head: 1 }
        );
        assert_eq!(
            decide_paragraph_split(3, 4, false, false),
            ParagraphSplit::Break { head: 3 }
        );
    }

    #[test]
    fn without_widow_control_nothing_fits_moves_whole() {
        assert_eq!(
            decide_paragraph_split(0, 4, false, false),
            ParagraphSplit::MoveWhole
        );
    }

    // ── Degenerate cases at the page top guarantee progress ────────────────

    #[test]
    fn at_page_top_unsplittable_paragraph_overflows_whole() {
        // 3-line paragraph, only 2 fit even on a full page, widow control on:
        // cannot split legally, already at the top → emit whole (overflow).
        assert_eq!(
            decide_paragraph_split(2, 3, true, true),
            ParagraphSplit::All
        );
    }

    #[test]
    fn at_page_top_single_oversized_line_overflows_whole() {
        // Even one line does not fit a full page → emit it and overflow.
        assert_eq!(
            decide_paragraph_split(0, 1, true, true),
            ParagraphSplit::All
        );
        assert_eq!(
            decide_paragraph_split(0, 3, false, true),
            ParagraphSplit::All
        );
    }

    // ── §17.3.1.11 drop-cap prefix guard ──────────────────────────────────

    #[test]
    fn drop_cap_head_clearing_the_prefix_is_unchanged() {
        // head >= drop_cap_lines: the whole prefix is in the first segment.
        assert_eq!(prefix_adjusted_head(5, 8, 3, true, true), Some(5));
        // No drop cap.
        assert_eq!(prefix_adjusted_head(2, 8, 0, true, false), Some(2));
        // Single-line drop cap can't be split within.
        assert_eq!(prefix_adjusted_head(1, 8, 1, true, false), Some(1));
        // Not the first segment — the prefix is behind us.
        assert_eq!(prefix_adjusted_head(1, 8, 3, false, false), Some(1));
        // head == remaining is the whole paragraph (no break inside the prefix).
        assert_eq!(prefix_adjusted_head(2, 2, 3, true, false), Some(2));
    }

    #[test]
    fn drop_cap_break_inside_prefix_moves_whole_when_not_at_top() {
        // head 2 < drop_cap_lines 3 → tearing the glyph; move the whole para.
        assert_eq!(prefix_adjusted_head(2, 8, 3, true, false), None);
        assert_eq!(prefix_adjusted_head(1, 8, 3, true, false), None);
    }

    #[test]
    fn drop_cap_break_inside_prefix_overflows_whole_at_top() {
        // At the page top the prefix can't move up: emit the whole prefix
        // (drop_cap_lines) and let it overflow, keeping the glyph intact.
        assert_eq!(prefix_adjusted_head(1, 8, 3, true, true), Some(3));
        // A paragraph shorter than the prefix clamps to its line count.
        assert_eq!(prefix_adjusted_head(1, 2, 3, true, true), Some(2));
    }
}
