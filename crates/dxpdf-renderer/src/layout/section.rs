//! Section layout — sequence blocks vertically into pages.
//!
//! Takes measured blocks (paragraphs with fragments, tables with cells),
//! fits them into pages respecting page size and margins, handles page breaks.

use std::rc::Rc;

use crate::dimension::Pt;
use crate::geometry::{PtRect, PtSize};
use super::draw_command::{DrawCommand, LayoutedPage};
use super::fragment::Fragment;
use super::page::PageConfig;
use super::paragraph::{layout_paragraph, ParagraphStyle};
use super::table::{layout_table, TableRowInput};
use super::BoxConstraints;

/// §17.4.58 / §17.4.59: positioning data for a floating table.
#[derive(Debug, Clone)]
pub struct TableFloatInfo {
    /// Gap between the table's right edge and surrounding text.
    pub right_gap: Pt,
    /// Gap between the table's bottom edge and surrounding text.
    pub bottom_gap: Pt,
    /// §17.4.58: horizontal alignment override (tblpXSpec).
    pub x_align: Option<dxpdf_docx_model::model::TableXAlign>,
    /// §17.4.59: absolute Y offset from the vertical anchor.
    pub y_offset: Pt,
    /// §17.4.58: vertical anchor reference (text / margin / page).
    pub vert_anchor: dxpdf_docx_model::model::TableAnchor,
}

/// A floating (anchor) image to be positioned absolutely on the page.
#[derive(Clone)]
pub struct FloatingImage {
    pub image_data: Rc<[u8]>,
    pub size: PtSize,
    /// Resolved absolute x position on the page.
    pub x: Pt,
    /// Resolved absolute y position on the page (may be relative to paragraph).
    pub y: FloatingImageY,
    /// §20.4.2.18: wrapTopAndBottom — text only above/below, not beside.
    pub wrap_top_and_bottom: bool,
    /// §20.4.2.3 distL/distR: horizontal distance from surrounding text.
    pub dist_left: Pt,
    pub dist_right: Pt,
}

/// Vertical position for a floating image.
#[derive(Clone, Copy)]
pub enum FloatingImageY {
    /// Absolute page position.
    Absolute(Pt),
    /// Relative to the paragraph's y position (offset added to cursor_y).
    RelativeToParagraph(Pt),
}

/// A block ready for layout — either a paragraph or a table.
pub enum LayoutBlock {
    Paragraph {
        fragments: Vec<Fragment>,
        style: ParagraphStyle,
        /// §17.3.1.23: force a page break before this paragraph.
        page_break_before: bool,
        /// Footnotes referenced in this paragraph — rendered at page bottom.
        footnotes: Vec<(Vec<Fragment>, ParagraphStyle)>,
        /// §20.4.2.3: floating images anchored to this paragraph.
        floating_images: Vec<FloatingImage>,
    },
    Table {
        rows: Vec<TableRowInput>,
        col_widths: Vec<Pt>,
        /// §17.4.38: resolved table border configuration.
        border_config: Option<super::table::TableBorderConfig>,
        /// §17.4.51: table indentation from left margin.
        indent: Pt,
        /// §17.4.28: table horizontal alignment.
        alignment: Option<dxpdf_docx_model::model::Alignment>,
        /// §17.4.58: floating table positioning — if present, text wraps around it.
        float_info: Option<TableFloatInfo>,
        /// §17.4.38: table style reference for adjacent table border collapse.
        style_id: Option<dxpdf_docx_model::model::StyleId>,
    },
}

/// §17.6.22: continuation state for `Continuous` section breaks.
/// Allows a new section to continue on the current page.
pub struct ContinuationState {
    pub page: LayoutedPage,
    pub cursor_y: Pt,
}

/// Lay out a sequence of blocks into pages.
///
/// If `continuation` is provided, the section starts on the given page at the
/// given cursor_y (for `SectionType::Continuous` sections).
pub fn layout_section(
    blocks: &[LayoutBlock],
    config: &PageConfig,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
    default_line_height: Pt,
    continuation: Option<ContinuationState>,
) -> Vec<LayoutedPage> {
    let content_width = config.content_width();
    let num_cols = config.num_columns();

    let mut pages: Vec<LayoutedPage> = Vec::new();
    let (mut current_page, mut cursor_y) = match continuation {
        Some(c) => (c.page, c.cursor_y),
        None => (LayoutedPage::new(config.page_size.clone()), config.margins.top),
    };
    let page_bottom = config.page_size.height - config.margins.bottom;
    // Effective bottom boundary — reduced by footnote height.
    let mut bottom = page_bottom;
    // §17.6.4: current column index (0-based).
    let mut current_col: usize = 0;
    // §17.6.4: y position where columns start on the current page.
    // On a continuation page, columns start at cursor_y; on fresh pages, at margins.top.
    let mut column_top = cursor_y;
    // Footnotes collected for the current page.
    let mut page_footnotes: Vec<(&[Fragment], &ParagraphStyle)> = Vec::new();
    // §17.3.1.33: only the structural first paragraph of a section on its
    // initial page has space_before suppressed. Paragraphs arriving at page
    // top via pageBreakBefore or overflow retain their space_before.
    let mut first_on_section_page = true;
    // §17.3.1.9: track previous paragraph for contextual spacing collapsing.
    let mut prev_space_after = Pt::ZERO;
    let mut prev_style_id: Option<dxpdf_docx_model::model::StyleId> = None;
    // §17.3.1.24: track previous paragraph borders for grouping.
    // Consecutive paragraphs with identical borders form a group — interior
    // paragraphs suppress their top border.
    let mut prev_borders: Option<super::paragraph::ParagraphBorderStyle> = None;
    // Unified float tracking: tables + images.
    let mut page_floats: Vec<super::float::ActiveFloat> = Vec::new();
    // Per-page absolute float cache: absolute floats from paragraphs on the
    // current page (including upcoming ones via forward scan). Reset on page break.
    let mut current_page_abs_floats: Vec<super::float::ActiveFloat> = Vec::new();
    // Index of the first block on the current page (for forward scanning).
    let mut page_start_block: usize = 0;
    // Whether we need to rescan absolute floats for this page.
    let mut abs_floats_dirty = true;
    // §17.4.59: y position at the start of the most recently rendered paragraph,
    // used as the anchor reference for floating tables with vertAnchor="text".
    let mut last_para_start_y = cursor_y;
    // §17.4.38: track previous table for adjacent table border collapse.
    // Consecutive tables with the same style are treated as a merged table:
    // the second table's top border is suppressed, and an insideH-equivalent
    // border gap is inserted between them.
    let mut prev_table_style_id: Option<dxpdf_docx_model::model::StyleId> = None;
    let mut prev_table_border_gap: Pt = Pt::ZERO;

    // Column-aware constraint: use current column's width.
    let col_constraints = |col: usize| -> BoxConstraints {
        let col_width = config.columns[col].width;
        BoxConstraints::new(Pt::ZERO, col_width, Pt::ZERO, config.content_height())
    };
    let col_x = |col: usize| -> Pt {
        config.margins.left + config.columns[col].x_offset
    };

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
                if *page_break_before && cursor_y > config.margins.top {
                    if !page_footnotes.is_empty() {
                        render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                        page_footnotes.clear();
                    }
                    pages.push(std::mem::replace(
                        &mut current_page,
                        LayoutedPage::new(config.page_size),
                    ));
                    cursor_y = config.margins.top;
                    column_top = config.margins.top;
                    current_col = 0;
                    prev_space_after = Pt::ZERO;
                    bottom = page_bottom;
                    page_start_block = block_idx;
                    abs_floats_dirty = true;
                    page_floats.clear();
                }

                let mut effective_style = style.clone();

                // Log paragraph info.
                let first_text = fragments.iter().find_map(|f| {
                    if let Fragment::Text { text, .. } = f { Some(text.as_str()) } else { None }
                }).unwrap_or("");
                log::debug!(
                    "[layout] block[{block_idx}] para style={:?} text={:?} cursor_y={:.1} col={current_col} floats={} fwd_floats={}",
                    effective_style.style_id, &first_text[..first_text.len().min(30)],
                    cursor_y.raw(), page_floats.len(), current_page_abs_floats.len()
                );

                // Save page state for potential keepNext rollback.
                let cmds_before_para = current_page.commands.len();

                // §17.3.1.33: suppress space_before for the "first paragraph
                // in a body/text story that begins on a page." Only the
                // structural first content block of a section is suppressed —
                // paragraphs at page top via pageBreakBefore or overflow
                // retain their space_before.
                if cursor_y <= column_top && first_on_section_page {
                    effective_style.space_before = Pt::ZERO;
                }
                // §17.3.1.24: paragraph border grouping — consecutive paragraphs
                // with identical borders form a group. Interior paragraphs
                // suppress their top border to avoid doubling.
                if effective_style.borders.is_some() && effective_style.borders == prev_borders {
                    if let Some(ref mut b) = effective_style.borders {
                        b.top = None;
                    }
                }
                // §17.3.1.9 / §17.3.1.33: spacing collapse must happen BEFORE
                // float registration so float y coordinates match actual paragraph position.
                if effective_style.contextual_spacing
                    && effective_style.style_id.is_some()
                    && effective_style.style_id == prev_style_id
                {
                    cursor_y -= prev_space_after + effective_style.space_before;
                } else {
                    let collapse = prev_space_after.min(effective_style.space_before);
                    cursor_y -= collapse;
                }

                // Register floating images (both relative and absolute).
                // §20.4.2.18: wrapTopAndBottom images are emitted immediately
                // and cursor_y advances past them — they act as block spacers.
                // §20.4.2.10: paragraph-relative floats use the content area
                // top (after space_before), not the total paragraph box top.
                let content_top = cursor_y + effective_style.space_before;
                for fi in floating_images.iter() {
                    let (y_start, y_end) = match fi.y {
                        FloatingImageY::RelativeToParagraph(offset) => {
                            (content_top + offset, content_top + offset + fi.size.height)
                        }
                        FloatingImageY::Absolute(img_y) => {
                            (img_y, img_y + fi.size.height)
                        }
                    };
                    if fi.wrap_top_and_bottom {
                        // §20.4.2.18: emit the image now and advance past it.
                        let img_y = match fi.y {
                            FloatingImageY::Absolute(y) => y,
                            FloatingImageY::RelativeToParagraph(offset) => content_top + offset,
                        };
                        current_page.commands.push(DrawCommand::Image {
                            rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                            image_data: fi.image_data.clone(),
                        });
                        // Push cursor_y past the image so text starts below it.
                        if y_end > cursor_y {
                            cursor_y = y_end;
                        }
                    } else {
                        // §20.4.2.3: use distL/distR for text distance from float.
                        let float_entry = super::float::ActiveFloat {
                            page_x: fi.x - fi.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fi.size.width + fi.dist_left + fi.dist_right,
                            source: super::float::FloatSource::Image,
                        };
                        log::debug!(
                            "[layout]   register image float: x={:.1} y={:.1}-{:.1} w={:.1}",
                            float_entry.page_x.raw(), y_start.raw(), y_end.raw(), float_entry.width.raw()
                        );
                        page_floats.push(float_entry);
                    }
                }

                // Prune expired floats, then pass float context for per-line adjustment.
                super::float::prune_floats(&mut page_floats, cursor_y);

                // Forward-scan absolute floats from upcoming paragraphs on the
                // current page. These may need to constrain text before the owning
                // paragraph is processed. Only rescan when the page changes.
                if abs_floats_dirty {
                    current_page_abs_floats.clear();
                    for (fi_idx, future_block) in blocks[page_start_block..].iter().enumerate() {
                        if let LayoutBlock::Paragraph { floating_images: fi_list, page_break_before, .. } = future_block {
                            // Stop scanning at the next explicit page break
                            // (skip the first block since it may be the one that triggered the new page).
                            if *page_break_before && fi_idx > 0 {
                                break;
                            }
                            for fi in fi_list {
                                if fi.wrap_top_and_bottom {
                                    continue; // handled as block spacers, not floats
                                }
                                if let FloatingImageY::Absolute(img_y) = fi.y {
                                    current_page_abs_floats.push(super::float::ActiveFloat {
                                        page_x: fi.x - fi.dist_left,
                                        page_y_start: img_y,
                                        page_y_end: img_y + fi.size.height,
                                        width: fi.size.width + fi.dist_left + fi.dist_right,
                                        source: super::float::FloatSource::Image,
                                    });
                                }
                            }
                        }
                    }
                    abs_floats_dirty = false;
                }

                // Merge page_floats with forward-scanned absolute floats (dedup).
                // Only include forward-scanned floats whose y range starts
                // above or at the current cursor — floats below shouldn't
                // affect paragraphs above them.
                let mut effective_floats = page_floats.clone();
                for af in &current_page_abs_floats {
                    // Only include forward-scanned floats that start above or
                    // at the current paragraph's content area top.
                    if af.page_y_start > cursor_y + effective_style.space_before {
                        continue;
                    }
                    let already_present = effective_floats.iter().any(|pf|
                        (pf.page_x - af.page_x).raw().abs() < 0.1
                        && (pf.page_y_start - af.page_y_start).raw().abs() < 0.1
                        && (pf.page_y_end - af.page_y_end).raw().abs() < 0.1
                    );
                    if !already_present {
                        effective_floats.push(af.clone());
                    }
                }
                if !effective_floats.is_empty() {
                    for (i, f) in effective_floats.iter().enumerate() {
                        log::debug!(
                            "[layout]   effective_float[{i}]: x={:.1} y={:.1}-{:.1} w={:.1} src={:?}",
                            f.page_x.raw(), f.page_y_start.raw(), f.page_y_end.raw(), f.width.raw(), f.source
                        );
                    }
                }
                // §17.4.56: if active floats leave no horizontal space at the
                // current cursor_y, advance past them so text can start below.
                // This handles full-width floating tables that act as block spacers
                // for non-owner paragraphs.
                let col_width = config.columns[current_col].width;
                let page_x = col_x(current_col);
                for ef in &effective_floats {
                    if ef.overlaps_y(cursor_y) && ef.width >= col_width {
                        cursor_y = cursor_y.max(ef.page_y_end);
                    }
                }
                super::float::prune_floats(&mut effective_floats, cursor_y);

                effective_style.page_floats = effective_floats;
                effective_style.page_y = cursor_y;
                effective_style.page_x = page_x;
                effective_style.page_content_width = col_width;

                // §17.6.4: split paragraph at column breaks for multi-column layout.
                let frag_chunks = split_at_column_breaks(fragments);
                let mut para_start_y = cursor_y;
                last_para_start_y = cursor_y;

                for (chunk_idx, chunk) in frag_chunks.iter().enumerate() {
                    // Advance to next column for chunks after a column break.
                    if chunk_idx > 0 {
                        if current_col + 1 < num_cols {
                            current_col += 1;
                        } else {
                            // All columns full — new page, reset to column 0.
                            if !page_footnotes.is_empty() {
                                render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                                page_footnotes.clear();
                            }
                            pages.push(std::mem::replace(
                                &mut current_page,
                                LayoutedPage::new(config.page_size.clone()),
                            ));
                            current_col = 0;
                            column_top = config.margins.top;
                            bottom = page_bottom;
                            page_start_block = block_idx;
                            abs_floats_dirty = true;
                    page_floats.clear();
                        }
                        cursor_y = column_top;
                        effective_style.page_x = col_x(current_col);
                        effective_style.page_content_width = config.columns[current_col].width;
                    }

                    let constraints = col_constraints(current_col);
                    let para = layout_paragraph(
                        chunk,
                        &constraints,
                        &effective_style,
                        default_line_height,
                        measure_text,
                    );

                    // Column/page overflow: advance column, then page.
                    if cursor_y + para.size.height > bottom && cursor_y > column_top {
                        if current_col + 1 < num_cols {
                            current_col += 1;
                            cursor_y = column_top;
                        } else {
                            if !page_footnotes.is_empty() {
                                render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                                page_footnotes.clear();
                            }
                            pages.push(std::mem::replace(
                                &mut current_page,
                                LayoutedPage::new(config.page_size.clone()),
                            ));
                            cursor_y = config.margins.top;
                            column_top = config.margins.top;
                            current_col = 0;
                            bottom = page_bottom;
                            page_start_block = block_idx;
                            abs_floats_dirty = true;
                    page_floats.clear();
                        }
                        // Update para_start_y after page/column change so
                        // floating images use the correct position.
                        para_start_y = cursor_y;
                    }

                    // Offset commands to absolute page position
                    log::debug!(
                        "[layout]   chunk[{chunk_idx}] placed at y={:.1} x={:.1} height={:.1}",
                        cursor_y.raw(), col_x(current_col).raw(), para.size.height.raw()
                    );
                    for mut cmd in para.commands {
                        cmd.shift_y(cursor_y);
                        cmd.shift_x(col_x(current_col));
                        current_page.commands.push(cmd);
                    }

                    cursor_y += para.size.height;
                }

                // §17.3.1.14: keepNext — if this paragraph has keep_next,
                // check if the next block fits on the same page. If not,
                // push both to the next page by undoing this paragraph's
                // placement and inserting a page break.
                if effective_style.keep_next
                    && cursor_y > column_top
                    && block_idx + 1 < blocks.len()
                    && num_cols <= 1 // skip for multi-column (too complex)
                {
                    let next_fits = match &blocks[block_idx + 1] {
                        LayoutBlock::Paragraph { fragments: next_frags, style: next_style, .. } => {
                            // Speculatively lay out the next paragraph to check height.
                            let mut next_eff = next_style.clone();
                            // Collapse spacing between current and next.
                            let next_collapse = effective_style.space_after.min(next_eff.space_before);
                            let next_cursor = cursor_y - next_collapse;
                            let next_constraints = col_constraints(current_col);
                            next_eff.page_y = next_cursor;
                            next_eff.page_x = col_x(current_col);
                            next_eff.page_content_width = config.columns[current_col].width;
                            let next_para = layout_paragraph(
                                next_frags, &next_constraints, &next_eff,
                                default_line_height, measure_text,
                            );
                            // Check if at least the first line fits.
                            next_cursor + next_para.size.height <= bottom
                        }
                        LayoutBlock::Table { .. } => true, // don't keepNext with tables
                    };

                    if !next_fits {
                        // Undo this paragraph: remove its commands from the page.
                        current_page.commands.truncate(cmds_before_para);
                        // Push page break.
                        if !page_footnotes.is_empty() {
                            render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                            page_footnotes.clear();
                        }
                        pages.push(std::mem::replace(
                            &mut current_page,
                            LayoutedPage::new(config.page_size.clone()),
                        ));
                        cursor_y = config.margins.top;
                        column_top = config.margins.top;
                        current_col = 0;
                        bottom = page_bottom;
                        page_start_block = block_idx;
                        abs_floats_dirty = true;
                    page_floats.clear();

                        // Re-layout the current paragraph on the new page.
                        // keepNext paragraphs at page top are not the structural
                        // first — their space_before is preserved.
                        let constraints = col_constraints(current_col);
                        effective_style.page_y = cursor_y;
                        effective_style.page_x = col_x(current_col);
                        effective_style.page_content_width = config.columns[current_col].width;
                        effective_style.page_floats = Vec::new();
                        let para = layout_paragraph(
                            fragments, &constraints, &effective_style,
                            default_line_height, measure_text,
                        );
                        for mut cmd in para.commands {
                            cmd.shift_y(cursor_y);
                            cmd.shift_x(col_x(current_col));
                            current_page.commands.push(cmd);
                        }
                        cursor_y += para.size.height;
                    }
                }

                first_on_section_page = false;
                prev_borders = style.borders.clone();
                prev_space_after = effective_style.space_after;
                prev_style_id = effective_style.style_id.clone();
                prev_table_style_id = None; // paragraph breaks adjacent table chain
                prev_table_border_gap = Pt::ZERO;

                // §20.4.2.3: emit floating images (already registered above).
                // wrapTopAndBottom images were already emitted before paragraph layout.
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
                    current_page.commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                        image_data: fi.image_data.clone(),
                    });
                }

                // Collect footnotes for this page and reserve space at bottom.
                if !footnotes.is_empty() {
                    let fn_constraints = super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
                    let sep_height = Pt::new(4.0); // separator line + gap
                    for (fn_frags, fn_style) in footnotes {
                        let fn_para = super::paragraph::layout_paragraph(
                            fn_frags, &fn_constraints, fn_style, default_line_height, measure_text,
                        );
                        // Only add separator space for the first footnote on this page.
                        if page_footnotes.is_empty() {
                            bottom -= sep_height;
                        }
                        bottom -= fn_para.size.height;
                        page_footnotes.push((fn_frags, fn_style));
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
                // §17.4.58: floating table — render at current position and
                // register as a float so subsequent text wraps around it.
                // Floating tables are absolutely positioned and do not
                // participate in §17.4.38 adjacent table border collapse.
                if let Some(fi) = float_info {
                    let table = layout_table(
                        rows,
                        col_widths,
                        &col_constraints(current_col),
                        default_line_height,
                        border_config.as_ref(),
                        measure_text,
                        false,
                    );

                    // §17.4.28 / §17.4.51: compute table x position.
                    let table_x = table_x_offset(
                        *alignment, *indent, table.size.width, content_width,
                        config.margins.left,
                    );

                    // Floating table breaks the adjacent table chain.
                    prev_table_style_id = None;
                    prev_table_border_gap = Pt::ZERO;
                    // §17.4.58: apply tblpXSpec horizontal alignment.
                    let table_x = match fi.x_align {
                        Some(dxpdf_docx_model::model::TableXAlign::Center) => {
                            config.margins.left + (content_width - table.size.width) * 0.5
                        }
                        Some(dxpdf_docx_model::model::TableXAlign::Right) => {
                            config.margins.left + content_width - table.size.width
                        }
                        _ => table_x,
                    };
                    if cursor_y + table.size.height > bottom && cursor_y > config.margins.top {
                        if !page_footnotes.is_empty() {
                            render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                            page_footnotes.clear();
                        }
                        pages.push(std::mem::replace(
                            &mut current_page,
                            LayoutedPage::new(config.page_size),
                        ));
                        cursor_y = config.margins.top;
                        column_top = config.margins.top;
                        current_col = 0;
                        prev_space_after = Pt::ZERO;
                        bottom = page_bottom;
                        page_start_block = block_idx;
                        abs_floats_dirty = true;
                    page_floats.clear();
                    }

                    // §17.4.59: tblpY is the absolute Y offset from the vertical
                    // anchor. The table must not start before cursor_y (preceding
                    // content already occupies space above it).
                    let float_y_start = if fi.y_offset > Pt::ZERO {
                        let anchor_y = match fi.vert_anchor {
                            dxpdf_docx_model::model::TableAnchor::Text => last_para_start_y + fi.y_offset,
                            dxpdf_docx_model::model::TableAnchor::Margin => config.margins.top + fi.y_offset,
                            dxpdf_docx_model::model::TableAnchor::Page => fi.y_offset,
                        };
                        anchor_y.max(cursor_y)
                    } else {
                        cursor_y
                    };
                    // §17.4.56: float y_end uses only the table's visual height.
                    // The bottom_gap is distance from text, not part of the float
                    // region — it prevents text from getting too close but doesn't
                    // extend the wrapping zone.
                    let float_y_end = float_y_start + table.size.height;

                    for mut cmd in table.commands {
                        cmd.shift_y(float_y_start);
                        cmd.shift_x(table_x);
                        current_page.commands.push(cmd);
                    }

                    log::debug!(
                        "[layout]   register table float: x={:.1} y={:.1}-{:.1} w={:.1} block_idx={block_idx}",
                        table_x.raw(), float_y_start.raw(), float_y_end.raw(), (table.size.width + fi.right_gap).raw()
                    );
                    // §17.4.56: register as float for text wrapping.
                    page_floats.push(super::float::ActiveFloat {
                        page_x: table_x,
                        page_y_start: float_y_start,
                        page_y_end: float_y_end,
                        width: table.size.width + fi.right_gap,
                        source: super::float::FloatSource::Table { owner_block_idx: block_idx },
                    });
                    continue;
                }

                // §17.4.38: consecutive non-floating tables with the same style
                // are treated as a single merged table — suppress the second
                // table's top border to avoid doubling at the shared edge,
                // and insert a border gap (equivalent to insideH) between them.
                let suppress_top = style_id.is_some()
                    && *style_id == prev_table_style_id;
                if suppress_top {
                    cursor_y += prev_table_border_gap;
                }

                // Non-floating table: paginated row-level splitting.
                // §17.4.49 / §17.4.1: split at row boundaries, repeat headers.
                let available = bottom - cursor_y;
                let slices = super::table::layout_table_paginated(
                    rows, col_widths,
                    &col_constraints(current_col),
                    default_line_height,
                    border_config.as_ref(),
                    measure_text,
                    available,
                    config.content_height(),
                    suppress_top,
                );

                // §17.4.28 / §17.4.51: compute table x position from
                // alignment and indent.
                let table_width: Pt = col_widths.iter().copied().sum();
                let table_x = table_x_offset(
                    *alignment, *indent, table_width,
                    content_width, config.margins.left,
                );

                for (slice_idx, slice) in slices.into_iter().enumerate() {
                    if slice_idx > 0 {
                        // Continuation slice → new page.
                        if !page_footnotes.is_empty() {
                            render_page_footnotes(&mut current_page, config, &page_footnotes, default_line_height, measure_text, separator_indent);
                            page_footnotes.clear();
                        }
                        pages.push(std::mem::replace(
                            &mut current_page,
                            LayoutedPage::new(config.page_size),
                        ));
                        cursor_y = config.margins.top;
                        column_top = config.margins.top;
                        current_col = 0;
                        bottom = page_bottom;
                        page_start_block = block_idx;
                        abs_floats_dirty = true;
                        page_floats.clear();
                    }

                    for mut cmd in slice.commands {
                        cmd.shift_y(cursor_y);
                        cmd.shift_x(table_x);
                        current_page.commands.push(cmd);
                    }
                    cursor_y += slice.size.height;
                }
                first_on_section_page = false;
                prev_borders = None; // table breaks border grouping
                prev_space_after = Pt::ZERO;
                prev_style_id = None;
                prev_table_style_id = style_id.clone();
                // §17.4.38: the gap between adjacent merged tables is the
                // table's bottom border width (equivalent to insideH in the
                // merged table). This gap separates rows visually and affects
                // page-break decisions.
                prev_table_border_gap = border_config
                    .as_ref()
                    .and_then(|b| b.bottom)
                    .map_or(Pt::ZERO, |b| b.width);
            }
        }
    }

    // Render footnotes on the current (last) page.
    if !page_footnotes.is_empty() {
        render_page_footnotes(
            &mut current_page, config, &page_footnotes,
            default_line_height, measure_text, separator_indent,
        );
    }

    // Push the last page (even if empty — ensure at least one page)
    pages.push(current_page);

    pages
}

// ── Shared block stacker ────────────────────────────────────────────────────

/// Result of stacking blocks vertically.
pub struct StackResult {
    /// Draw commands positioned relative to the stacking origin (0,0).
    pub commands: Vec<DrawCommand>,
    /// Total height consumed by all blocks.
    pub height: Pt,
}

/// Stack blocks vertically within a fixed-width area.
///
/// This is the shared core used by both page-level layout (`layout_section`)
/// and cell-level layout. It handles:
/// - Paragraph layout with spacing collapse and space_before suppression
/// - Table layout
/// - Floating image registration and text wrapping
///
/// It does NOT handle page breaks, column breaks, or footnote collection —
/// those are page-level concerns managed by `layout_section`.
pub fn stack_blocks(
    blocks: &[LayoutBlock],
    content_width: Pt,
    default_line_height: Pt,
    measure_text: super::paragraph::MeasureTextFn<'_>,
) -> StackResult {
    let constraints = super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
    let mut commands = Vec::new();
    let mut cursor_y = Pt::ZERO;
    let mut prev_space_after = Pt::ZERO;
    let mut prev_style_id: Option<dxpdf_docx_model::model::StyleId> = None;
    let mut page_floats: Vec<super::float::ActiveFloat> = Vec::new();

    for block in blocks {
        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                floating_images,
                ..
            } => {
                let mut effective_style = style.clone();

                // Spacing collapse.
                if effective_style.contextual_spacing
                    && effective_style.style_id.is_some()
                    && effective_style.style_id == prev_style_id
                {
                    cursor_y -= prev_space_after + effective_style.space_before;
                } else {
                    let collapse = prev_space_after.min(effective_style.space_before);
                    cursor_y -= collapse;
                }

                // Register floating images.
                let content_top = cursor_y + effective_style.space_before;
                for fi in floating_images.iter() {
                    let (y_start, y_end) = match fi.y {
                        FloatingImageY::RelativeToParagraph(offset) => {
                            (content_top + offset, content_top + offset + fi.size.height)
                        }
                        FloatingImageY::Absolute(img_y) => {
                            (img_y, img_y + fi.size.height)
                        }
                    };
                    if fi.wrap_top_and_bottom {
                        let img_y = match fi.y {
                            FloatingImageY::Absolute(y) => y,
                            FloatingImageY::RelativeToParagraph(offset) => content_top + offset,
                        };
                        commands.push(DrawCommand::Image {
                            rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                            image_data: fi.image_data.clone(),
                        });
                        if y_end > cursor_y {
                            cursor_y = y_end;
                        }
                    } else {
                        page_floats.push(super::float::ActiveFloat {
                            page_x: fi.x - fi.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fi.size.width + fi.dist_left + fi.dist_right,
                            source: super::float::FloatSource::Image,
                        });
                    }
                }

                super::float::prune_floats(&mut page_floats, cursor_y);

                effective_style.page_floats = page_floats.clone();
                effective_style.page_y = cursor_y;
                effective_style.page_x = Pt::ZERO;
                effective_style.page_content_width = content_width;

                let para = layout_paragraph(
                    fragments,
                    &constraints,
                    &effective_style,
                    default_line_height,
                    measure_text,
                );

                for mut cmd in para.commands {
                    cmd.shift_y(cursor_y);
                    commands.push(cmd);
                }

                cursor_y += para.size.height;

                // Emit non-wrapTopAndBottom floating images.
                let para_content_top = cursor_y - para.size.height + effective_style.space_before;
                for fi in floating_images {
                    if fi.wrap_top_and_bottom {
                        continue;
                    }
                    let img_y = match fi.y {
                        FloatingImageY::Absolute(y) => y,
                        FloatingImageY::RelativeToParagraph(offset) => para_content_top + offset,
                    };
                    commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                        image_data: fi.image_data.clone(),
                    });
                    // Extend cursor to encompass the image so table cells
                    // expand to contain floating images.
                    let img_bottom = img_y + fi.size.height;
                    if img_bottom > cursor_y {
                        cursor_y = img_bottom;
                    }
                }

                prev_space_after = effective_style.space_after;
                prev_style_id = effective_style.style_id.clone();
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                border_config,
                indent,
                alignment,
                ..
            } => {
                // stack_blocks is used for table cells and header/footer —
                // no adjacent table collapse in these contexts.
                let table = layout_table(
                    rows,
                    col_widths,
                    &constraints,
                    default_line_height,
                    border_config.as_ref(),
                    measure_text,
                    false,
                );

                let table_x = table_x_offset(
                    *alignment, *indent, table.size.width, content_width,
                    Pt::ZERO,
                );

                for mut cmd in table.commands {
                    cmd.shift_y(cursor_y);
                    cmd.shift_x(table_x);
                    commands.push(cmd);
                }

                cursor_y += table.size.height;
                prev_space_after = Pt::ZERO;
                prev_style_id = None;
            }
        }
    }

    StackResult {
        commands,
        height: cursor_y,
    }
}

// ── Helper functions ────────────────────────────────────────────────────────

/// Split a fragment slice at `Fragment::ColumnBreak` markers.
/// Returns a vec of slices; the column break fragments themselves are excluded.
fn split_at_column_breaks(fragments: &[Fragment]) -> Vec<&[Fragment]> {
    let has_break = fragments.iter().any(|f| matches!(f, Fragment::ColumnBreak));
    if !has_break {
        return vec![fragments];
    }
    let mut chunks = Vec::new();
    let mut start = 0;
    for (i, frag) in fragments.iter().enumerate() {
        if matches!(frag, Fragment::ColumnBreak) {
            chunks.push(&fragments[start..i]);
            start = i + 1;
        }
    }
    chunks.push(&fragments[start..]);
    chunks
}

fn render_page_footnotes(
    page: &mut LayoutedPage,
    config: &PageConfig,
    footnotes: &[(&[Fragment], &ParagraphStyle)],
    default_line_height: Pt,
    measure_text: super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
) {
    use super::draw_command::DrawCommand;
    use super::paragraph::layout_paragraph;

    let content_width = config.content_width();
    let constraints = super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
    let page_bottom = config.page_size.height - config.margins.bottom;

    // Layout all footnotes to compute total height.
    let mut footnote_layouts = Vec::new();
    let mut total_height = Pt::new(4.0); // separator + gap
    for (frags, style) in footnotes {
        let para = layout_paragraph(frags, &constraints, style, default_line_height, measure_text);
        total_height += para.size.height;
        footnote_layouts.push(para);
    }

    let footnote_top = page_bottom - total_height;

    // §17.11.23: separator line positioned per default paragraph indent.
    let sep_x = config.margins.left + separator_indent;
    let sep_width = content_width * 0.33;
    page.commands.push(DrawCommand::Line {
        line: crate::geometry::PtLineSegment::new(
            crate::geometry::PtOffset::new(sep_x, footnote_top),
            crate::geometry::PtOffset::new(sep_x + sep_width, footnote_top),
        ),
        color: crate::resolve::color::RgbColor::BLACK,
        width: Pt::new(0.5),
    });

    // Render footnote paragraphs.
    let mut cursor_y = footnote_top + Pt::new(4.0);
    for para in footnote_layouts {
        for mut cmd in para.commands {
            cmd.shift_y(cursor_y);
            cmd.shift_x(config.margins.left);
            page.commands.push(cmd);
        }
        cursor_y += para.size.height;
    }
}

/// §17.4.28: compute the table's x offset based on alignment and indent.
fn table_x_offset(
    alignment: Option<dxpdf_docx_model::model::Alignment>,
    indent: Pt,
    table_width: Pt,
    content_width: Pt,
    margin_left: Pt,
) -> Pt {
    use dxpdf_docx_model::model::Alignment;
    match alignment {
        Some(Alignment::Center) => {
            margin_left + (content_width - table_width) * 0.5
        }
        Some(Alignment::End) => {
            margin_left + content_width - table_width
        }
        _ => margin_left + indent,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{PtEdgeInsets, PtSize};
    use crate::layout::draw_command::DrawCommand;
    use crate::layout::fragment::{FontProps, TextMetrics};
    use crate::layout::table::TableCellInput;
    use crate::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str, width: f32, height: f32) -> Fragment {
        Fragment::Text {
            text: text.into(),
            font: FontProps {
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: false,
                italic: false,
                underline: false,
                char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width), trimmed_width: Pt::new(width),
            metrics: TextMetrics { ascent: Pt::new(height * 0.7), descent: Pt::new(height * 0.3) },
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO, text_offset: Pt::ZERO,
        }
    }

    fn para_block(text: &str, width: f32) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, width, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
        }
    }

    fn small_config() -> PageConfig {
        use crate::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
            columns: vec![ColumnGeometry { x_offset: Pt::ZERO, width: Pt::new(180.0) }],
        }
    }

    #[test]
    fn empty_blocks_produces_one_empty_page() {
        let pages = layout_section(&[], &small_config(), None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn single_paragraph_on_one_page() {
        let blocks = vec![para_block("hello", 30.0)];
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0), None);

        assert_eq!(pages.len(), 1);
        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    #[test]
    fn text_positioned_at_margins() {
        let blocks = vec![para_block("hello", 30.0)];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        if let Some(DrawCommand::Text { position, .. }) = pages[0].commands.first() {
            assert!(
                position.x.raw() >= config.margins.left.raw(),
                "x should be at least left margin"
            );
            assert!(
                position.y.raw() >= config.margins.top.raw(),
                "y should be at least top margin"
            );
        }
    }

    #[test]
    fn page_break_when_content_overflows() {
        // Page: 100pt tall, margins 10 each → 80pt content area
        // Each paragraph: 14pt tall
        // 6 paragraphs = 84pt > 80pt → should break to 2 pages
        let blocks: Vec<_> = (0..6).map(|i| para_block(&format!("p{i}"), 30.0)).collect();
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0), None);

        assert_eq!(pages.len(), 2, "should overflow to 2 pages");

        let page1_texts: Vec<_> = pages[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect();
        let page2_texts: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect();

        assert_eq!(page1_texts.len(), 5, "5 paras fit on page 1 (5*14=70 < 80)");
        assert_eq!(page2_texts.len(), 1, "1 para on page 2");
    }

    #[test]
    fn page_size_set_on_layouted_page() {
        let config = small_config();
        let pages = layout_section(&[], &config, None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages[0].page_size, config.page_size);
    }

    #[test]
    fn many_paragraphs_produce_multiple_pages() {
        // 20 paragraphs at 14pt each = 280pt
        // Content area = 80pt → need 4 pages (80/14 = 5.7 paras per page)
        let blocks: Vec<_> = (0..20).map(|i| para_block(&format!("p{i}"), 30.0)).collect();
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0), None);

        assert_eq!(pages.len(), 4);
    }

    #[test]
    fn table_on_page() {
        let blocks = vec![LayoutBlock::Table {
            rows: vec![TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: vec![text_frag("cell", 30.0, 14.0)],
                        style: ParagraphStyle::default(),
                        page_break_before: false,
                        footnotes: vec![],
                        floating_images: vec![],
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None, cell_borders: None, vertical_merge: None, vertical_align: crate::layout::table::CellVAlign::Top,
                }],
                height_rule: None,
            is_header: None,
            cant_split: None,
            }],
            col_widths: vec![Pt::new(100.0)],
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        }];

        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages.len(), 1);

        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    // ── §17.3.1.33 space_before suppression tests ──────────────────────

    #[test]
    fn space_before_suppressed_for_first_paragraph_of_section() {
        let mut style = ParagraphStyle::default();
        style.space_before = Pt::new(24.0);
        let blocks = vec![LayoutBlock::Paragraph {
            fragments: vec![text_frag("heading", 50.0, 14.0)],
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
        }];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        // First paragraph on the section's initial page: space_before suppressed.
        if let Some(DrawCommand::Text { position, .. }) = pages[0].commands.first() {
            assert!(
                position.y.raw() < config.margins.top.raw() + 24.0,
                "space_before should be suppressed: y={}",
                position.y.raw()
            );
        }
    }

    #[test]
    fn space_before_preserved_for_page_break_before() {
        let mut heading_style = ParagraphStyle::default();
        heading_style.space_before = Pt::new(24.0);

        let blocks = vec![
            para_block("first page", 30.0),
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("heading", 50.0, 14.0)],
                style: heading_style,
                page_break_before: true,
                footnotes: vec![],
                floating_images: vec![],
            },
        ];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        assert!(pages.len() >= 2, "should have at least 2 pages");
        let heading_y = pages[1].commands.iter()
            .find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if text == "heading" => Some(position.y),
                _ => None,
            })
            .expect("heading should be on page 2");
        // §17.3.1.33: space_before is preserved — pageBreakBefore paragraphs
        // are not the structural first of the section.
        assert!(
            heading_y.raw() > config.margins.top.raw() + 20.0,
            "space_before should be preserved for pageBreakBefore: y={}",
            heading_y.raw(),
        );
    }

    // ── §17.3.1.24 paragraph border grouping tests ─────────────────────

    #[test]
    fn identical_borders_suppress_second_top() {
        use crate::layout::paragraph::{ParagraphBorderStyle, BorderLine};
        let border = Some(ParagraphBorderStyle {
            top: Some(BorderLine { width: Pt::new(0.5), color: RgbColor::BLACK, space: Pt::new(1.0) }),
            bottom: None, left: None, right: None,
        });
        let mut style1 = ParagraphStyle::default();
        style1.borders = border.clone();
        let mut style2 = ParagraphStyle::default();
        style2.borders = border;

        let blocks = vec![
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("para1", 30.0, 14.0)],
                style: style1,
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
            },
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("para2", 30.0, 14.0)],
                style: style2,
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
            },
        ];
        let pages = layout_section(&blocks, &small_config(), None, Pt::ZERO, Pt::new(14.0), None);

        // Count Line draw commands (border lines).
        // Only the first paragraph should draw its top border; the second's
        // top border is suppressed by §17.3.1.24 grouping.
        let line_cmds: Vec<_> = pages[0].commands.iter()
            .filter(|c| matches!(c, DrawCommand::Line { .. }))
            .collect();
        assert_eq!(line_cmds.len(), 1, "only one top border line (grouped): got {}", line_cmds.len());
    }
}
