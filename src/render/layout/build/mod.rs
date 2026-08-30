//! Recursive tree-walk that converts a resolved document model into layout blocks.
//!
//! The document tree (Section → Block → Paragraph | Table → Cell → Block…) is
//! processed by recursive descent.  Each `Block::Table` recurses into its cells'
//! content, which may contain nested tables.

pub(super) mod block;
pub(super) mod convert;
pub(super) mod floating;
pub(super) mod list_label;
pub(super) mod table;

use std::collections::HashMap;

use crate::model::{self, Block};
use crate::render::dimension::Pt;
use crate::render::layout::fragment::Fragment;
use crate::render::layout::measurer::TextMeasurer;
use crate::render::layout::page::PageConfig;
use crate::render::layout::paragraph::ParagraphStyle;
use crate::render::layout::section::LayoutBlock;
use crate::render::resolve::images::MediaEntry;
use crate::render::resolve::sections::ResolvedSection;
use crate::render::resolve::ResolvedDocument;

use block::{build_block, build_fragments, collect_endnotes};
use convert::{
    doc_font_family, doc_font_size, paragraph_style_from_props, resolve_paragraph_defaults,
};
use floating::{extract_floating_images, find_vml_absolute_position, AnchorFrame};
use table::build_table;

/// §17.8.3.2: OOXML fallback font when no theme or doc defaults specify one.
pub(super) const SPEC_FALLBACK_FONT: &str = "Times New Roman";
/// §17.3.2.14: default font size (10pt = 20 half-points per ECMA-376 §17.3.2.14).
pub(super) const SPEC_DEFAULT_FONT_SIZE: Pt = Pt::new(10.0);

// ── Context ─────────────────────────────────────────────────────────────────

/// Immutable context threaded through the recursive tree walk.
pub struct BuildContext<'a> {
    pub measurer: &'a TextMeasurer<'a>,
    pub resolved: &'a ResolvedDocument,
}

impl BuildContext<'_> {
    pub(super) fn media(&self) -> &HashMap<model::RelId, MediaEntry> {
        &self.resolved.media
    }
}

/// Mutable state threaded through the recursive tree walk.
#[derive(Default)]
pub struct BuildState {
    /// Page configuration for the current section.
    pub page_config: crate::render::layout::page::PageConfig,
    /// §17.11.12: footnote display numbering plus the ordered record of which
    /// notes each paragraph referenced. Advanced by `collect_fragments` and
    /// drained per paragraph — see
    /// [`FootnoteTracker`](crate::render::layout::fragment::FootnoteTracker).
    pub footnotes: crate::render::layout::fragment::FootnoteTracker,
    /// Sequential endnote display number (i, ii, iii...).
    pub endnote_counter: u32,
    /// Per-(numId, level) running counters for list labels.
    pub list_counters: HashMap<(model::NumId, u8), u32>,
    /// §17.3.1.19: the document outline being accumulated.
    pub outline: OutlineCollector,
    /// Field evaluation context (page number, total pages).
    pub field_ctx: crate::render::layout::fragment::FieldContext,
    /// §20.1.4.1.17: default text color for the current shape text box (from
    /// `wps:style/fontRef`), applied as the base below the run/style cascade.
    /// `None` outside a shape text box.
    pub shape_default_text_color: Option<crate::render::resolve::color::RgbColor>,
    /// §20.1.4.1.17: default font family for the current shape text box (the
    /// theme major/minor collection selected by `fontRef@idx`). `None` outside
    /// a shape text box.
    pub shape_default_font_family: Option<String>,
    /// §20.1.2.1.18: the `a:normAutofit` shrink in force for the current shape
    /// text body — the scale Word computed and wrote into the file.
    /// [`ShapeAutoFit::NONE`](crate::render::layout::ShapeAutoFit::NONE)
    /// outside a shape text box, and for a body that declares
    /// `noAutofit`/`spAutoFit`.
    pub shape_auto_fit: crate::render::layout::ShapeAutoFit,
    /// §17.4.38: border styles already reported as unsupported this render.
    ///
    /// The layout layer draws only `Single` and `Double`, so the other 24
    /// `BorderStyle` variants are approximated by a solid line. Warning at
    /// every call would emit one line per cell edge — thousands for a large
    /// table — so each distinct style is reported once. Scoped to the render
    /// (like `Measurer::warned_emoji`) rather than to the process, so a second
    /// document still reports its own.
    pub warned_border_styles: std::collections::HashSet<model::BorderStyle>,
    /// §17.4.41 / §17.4.42: a row overrode the table's `tblCellSpacing`, which
    /// layout applies per *table*, not per row. Reported once per render for
    /// the same reason as `warned_border_styles`.
    pub warned_row_cell_spacing: bool,
    /// §17.4.85: a `w:vMerge w:val="continue"` cell had no `restart` above it
    /// and was rendered as an ordinary cell. Malformed input, so it is worth a
    /// `RUST_LOG=warn` line — but once per render, not once per cell: a table
    /// built by a producer that gets this wrong tends to get it wrong in every
    /// row. See `build::table::promote_orphan_vmerge_continues`.
    pub warned_orphan_vmerge: bool,
}

// ── Public entry point ──────────────────────────────────────────────────────

/// Built section output — layout blocks.
pub struct BuiltSection {
    pub blocks: Vec<LayoutBlock>,
}

/// Build layout blocks for one section by recursing into its block tree.
pub fn build_section_blocks(
    section: &ResolvedSection,
    config: &PageConfig,
    ctx: &BuildContext,
    state: &mut BuildState,
) -> BuiltSection {
    let mut pending_dropcap: Option<crate::render::layout::paragraph::DropCapInfo> = None;
    let blocks: Vec<LayoutBlock> = section
        .blocks
        .iter()
        .filter_map(|block| {
            build_block(
                block,
                config.content_width(),
                ctx,
                state,
                &mut pending_dropcap,
            )
        })
        .collect();

    BuiltSection { blocks }
}

/// §17.11.2: build the document's endnote content, rendered once at the end of
/// the document.
///
/// Endnotes are **document-scoped**, not section-scoped: `ResolvedDocument`
/// holds a single endnote map. Calling this per section would emit every
/// endnote once per section, so it deliberately sits outside the section loop.
pub fn build_document_endnotes(
    ctx: &BuildContext,
    state: &mut BuildState,
) -> Vec<(String, Vec<Fragment>, ParagraphStyle)> {
    let mut endnotes = Vec::new();
    collect_endnotes(ctx, state, &mut endnotes);
    endnotes
}

/// Issue #154: build every comment's body once, for the balloon pass — the
/// same shape as the endnote prebuild above. Keyed by comment id; paragraphs
/// come back as (fragments, style) pairs ready for `layout_paragraph` at
/// balloon width. Ids are processed in order so any shared build state stays
/// deterministic.
pub fn build_comment_bodies(
    ctx: &BuildContext,
    state: &mut BuildState,
) -> std::collections::HashMap<crate::model::CommentId, Vec<(Vec<Fragment>, ParagraphStyle)>> {
    let mut out = std::collections::HashMap::new();
    let mut ids: Vec<_> = ctx.resolved.comments.keys().copied().collect();
    ids.sort();
    for id in ids {
        let comment = &ctx.resolved.comments[&id];
        // The note builder lays out paragraphs only; a table in a comment
        // body is dropped, like a table in a footnote. Say so rather than
        // silently shrinking the balloon.
        if comment
            .content
            .iter()
            .any(|b| !matches!(b, crate::model::Block::Paragraph(_)))
        {
            log::warn!(
                "comment {:?}: non-paragraph content (a table?) is not rendered in the balloon",
                id
            );
        }
        let paras = block::build_note_content("", &comment.content, ctx, state)
            .into_iter()
            .map(|(_, frags, style)| (frags, style))
            .collect();
        out.insert(id, paras);
    }
    out
}

/// Collected header/footer content with layout metadata.
pub struct HeaderFooterContent {
    /// Layout blocks (paragraphs and tables) for stacking.
    pub blocks: Vec<LayoutBlock>,
    /// Absolute page-relative position from a VML text box, if present.
    pub absolute_position: Option<(Pt, Pt)>,
    /// Floating (anchor) images from header/footer paragraphs.
    ///
    /// Images are collected outside the layout-block tree so the renderer
    /// can split behindDoc and in-front-of-text images around the text
    /// commands for correct z-ordering. Floating shapes, by contrast, ride
    /// on `LayoutBlock::Paragraph.floating_shapes` so the stacker can anchor
    /// them to the owning paragraph's y position.
    pub floating_images: Vec<crate::render::layout::section::FloatingImage>,
    /// Page-anchored floating shapes (vertical `relativeFrom` is
    /// `page` / `margin` / `topMargin` / `bottomMargin` / etc.).
    ///
    /// Their `y` is `FloatingImageY::Absolute(...)` resolved in `Page`
    /// frame, and the renderer emits them directly without applying the
    /// header/footer stack shift — so a `<wp:positionV relativeFrom="page">`
    /// shape in a footer (e.g. a "Seite X von Y" indicator pinned just
    /// below the gray address bar) lands at its authored page-y.
    /// Paragraph- / line-anchored shapes still travel on
    /// `LayoutBlock::Paragraph.floating_shapes` so they follow the host
    /// paragraph's y.
    pub floating_shapes: Vec<crate::render::layout::section::FloatingShape>,
}

/// Build header/footer content from blocks.
///
/// Produces `LayoutBlock` entries for both paragraphs and tables, and
/// extracts floating images separately (they are positioned page-relative
/// rather than stack-relative).
/// Lay out content that is **not** the document's main story: a header, a
/// footer, or a shape text body.
///
/// §17.3.1.19: the outline collector is suspended for the whole call and
/// restored after. Doing it here rather than at each paragraph is what makes it
/// correct — a `Block::Table` in a header is handed to the *body* table
/// builder, so its cell paragraphs reach the ordinary paragraph builder with
/// the document's state. Guarding only the paragraph arm let a heading in a
/// header table into the outline once per page the header was drawn on.
///
/// Save-and-restore rather than a fresh flag because the call nests: a shape
/// text body inside a header goes through here twice.
pub fn build_header_footer_content(
    blocks: &[Block],
    ctx: &BuildContext,
    state: &mut BuildState,
) -> HeaderFooterContent {
    let outer = std::mem::replace(&mut state.outline, OutlineCollector::Excluded);
    let content = build_non_story_content(blocks, ctx, state);
    state.outline = outer;
    content
}

fn build_non_story_content(
    blocks: &[Block],
    ctx: &BuildContext,
    state: &mut BuildState,
) -> HeaderFooterContent {
    let mut layout_blocks = Vec::new();
    let mut all_floating_images = Vec::new();
    let mut all_page_anchored_shapes = Vec::new();
    let mut absolute_position: Option<(Pt, Pt)> = None;

    let available_width = state.page_config.content_width();

    let block_count = blocks.len();
    for (block_i, block) in blocks.iter().enumerate() {
        match block {
            Block::Paragraph(p) => {
                let (mut frags, props) = build_fragments(p, ctx, state, None, None);
                // §17.11.12: headers/footers don't render footnote bodies, but
                // they must still drain — otherwise a reference inside one
                // would be attributed to the next body paragraph.
                let _ = state.footnotes.take_pending();
                // §17.3.1.6 / UAX #9. A header carries no list label, so its
                // fragment vector is already final here.
                block::resolve_paragraph_bidi(&mut frags, &props, ctx);
                let style = paragraph_style_from_props(
                    &props,
                    Pt::from(ctx.resolved.default_tab_stop),
                    state.shape_auto_fit,
                    convert::paragraph_locale(p, ctx.resolved),
                    // §17.3.1.19: always `None` here — the collector is
                    // suspended for this whole call. Asking rather than passing
                    // `None` keeps one rule: whether a paragraph is a heading is
                    // decided in one place, for every path.
                    convert::paragraph_outline(p, &props, state),
                );

                // Check for VML absolute positioning in Pict inlines.
                if absolute_position.is_none() {
                    for inline in &p.content {
                        if let Some(pos) = find_vml_absolute_position(inline) {
                            absolute_position = Some(pos);
                            break;
                        }
                    }
                }
                // §20.4.2.3: extract floating (anchor) images. Images are
                // collected at the header/footer level and emitted by
                // `render_header`/`render_footer` *outside* of `stack_blocks`
                // — that's how z-ordering against the text (behindDoc) is
                // implemented. Because the caller does not apply the stack
                // shift to these commands, their coordinates must be
                // page-absolute.
                let para_floats = extract_floating_images(p, ctx, state, AnchorFrame::Page);
                let has_float_images = !para_floats.is_empty();
                all_floating_images.extend(para_floats);
                // Shapes, by contrast, travel on the owning paragraph so
                // `stack_blocks` can anchor them to the paragraph's y
                // coordinate. `stack_blocks` emits every command in
                // stack-frame-relative space and `render_footer` /
                // `render_header` later shift the whole batch by
                // `margins.left` — so shapes must be resolved in the stack
                // frame (not page-absolute), otherwise `margins.left` would
                // be applied twice. §17.10.1: z-ordering against the
                // surrounding text in Tier 0 is paragraph-granular, matching
                // the stack-frame emission order.
                let paragraph_shapes = floating::extract_floating_shapes(
                    p,
                    ctx,
                    state,
                    AnchorFrame::Stack,
                    floating::ShapeAnchorClass::ParagraphAnchored,
                );

                // §17.10.1 / §20.4.2.10: shapes whose vertical position is
                // page/margin-relative resolve to a fixed absolute y. They
                // can't ride on the per-paragraph anchoring used by
                // `paragraph_shapes`, so collect them at the
                // header/footer level — `render_header`/`render_footer`
                // emits them outside `stack_blocks` (no shift) so the
                // resolved page-y is honored.
                let page_anchored_shapes = floating::extract_floating_shapes(
                    p,
                    ctx,
                    state,
                    AnchorFrame::Page,
                    floating::ShapeAnchorClass::PageAnchored,
                );
                all_page_anchored_shapes.extend(page_anchored_shapes);

                // §17.10.1: empty non-last paragraphs in headers/footers still
                // occupy a line height (from the paragraph mark's font size).
                //
                // Exception: paragraphs whose only content is floating shape
                // or image anchors are treated as zero-height anchors —
                // Word positions the shape relative to the paragraph top,
                // which must coincide with the preceding paragraph's bottom
                // (otherwise the shape displaces text it should flank).
                let has_floating_anchor = has_float_images || !paragraph_shapes.is_empty();
                if frags.is_empty() && block_i + 1 < block_count && !has_floating_anchor {
                    let (family, mut size, ..) =
                        resolve_paragraph_defaults(p, ctx.resolved, false, None, None);
                    if let Some(ref mrp) = p.mark_run_properties {
                        if let Some(fs) = mrp.font_size.cloned() {
                            size = Pt::from(fs);
                        }
                    }
                    let line_height = ctx.measurer.default_line_height(&family, size);
                    frags.push(Fragment::LineBreak { line_height });
                }

                layout_blocks.push(LayoutBlock::Paragraph {
                    fragments: frags,
                    style,
                    page_break_before: false,
                    footnotes: vec![],
                    floating_images: vec![], // handled separately above
                    floating_shapes: paragraph_shapes,
                });
            }
            Block::Table(t) => {
                let built = build_table(t, available_width, ctx, state);
                layout_blocks.push(LayoutBlock::Table {
                    rows: built.rows,
                    col_widths: built.col_widths,
                    cell_spacing: built.cell_spacing,
                    border_config: built.border_config,
                    indent: built.indent,
                    alignment: built.alignment,
                    float_info: built.float_info,
                    style_id: t.properties.style_id.clone(),
                });
            }
            Block::SectionBreak(_) => {}
        }
    }

    HeaderFooterContent {
        blocks: layout_blocks,
        absolute_position,
        floating_images: all_floating_images,
        floating_shapes: all_page_anchored_shapes,
    }
}

/// Default line height derived from document-level font settings.
pub fn default_line_height(ctx: &BuildContext) -> Pt {
    let family = doc_font_family(ctx);
    let size = doc_font_size(ctx);
    ctx.measurer.default_line_height(&family, size)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::Dimension;
    use crate::render::fonts::FontRegistry;
    use crate::render::layout::measurer::TextMeasurer;

    fn empty_resolved() -> ResolvedDocument {
        ResolvedDocument {
            sections: Vec::new(),
            styles: HashMap::new(),
            numbering: HashMap::new(),
            font_families: Vec::new(),
            media: HashMap::new(),
            embedded_fonts: Vec::new(),
            pic_bullets: HashMap::new(),
            theme: None,
            doc_defaults_paragraph: model::ParagraphProperties::default(),
            doc_defaults_run: model::RunProperties::default(),
            default_paragraph_style_id: None,
            default_table_style_id: None,
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            even_and_odd_headers: false,
            default_tab_stop: Dimension::new(720),
            show_ins_del_marks: true,
            show_comment_marks: true,
            revision_colors: Default::default(),
            comments: Default::default(),
        }
    }

    fn empty_para() -> Block {
        Block::Paragraph(Box::new(model::Paragraph {
            style_id: None,
            properties: model::ParagraphProperties::default(),
            mark_run_properties: None,
            content: Vec::new(),
            rsids: model::ParagraphRevisionIds::default(),
            mark_deleted: false,
        }))
    }

    fn line_break_counts(blocks: &[LayoutBlock]) -> Vec<usize> {
        blocks
            .iter()
            .map(|b| match b {
                LayoutBlock::Paragraph { fragments, .. } => fragments
                    .iter()
                    .filter(|f| matches!(f, Fragment::LineBreak { .. }))
                    .count(),
                _ => 0,
            })
            .collect()
    }

    /// §17.10.1: an empty paragraph in a header/footer holds a line of height —
    /// *except* the last one, which would otherwise push the whole stack.
    #[test]
    fn empty_header_paragraphs_hold_a_line_except_the_last() {
        let resolved = empty_resolved();
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut state = BuildState::default();

        let blocks = vec![empty_para(), empty_para(), empty_para()];
        let hf = build_header_footer_content(&blocks, &ctx, &mut state);

        assert_eq!(hf.blocks.len(), 3);
        assert_eq!(
            line_break_counts(&hf.blocks),
            vec![1, 1, 0],
            "the trailing empty paragraph contributes no line"
        );
    }

    /// A single empty paragraph is also the last one, so it contributes nothing.
    #[test]
    fn a_lone_empty_header_paragraph_holds_no_line() {
        let resolved = empty_resolved();
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut state = BuildState::default();

        let hf = build_header_footer_content(&[empty_para()], &ctx, &mut state);
        assert_eq!(line_break_counts(&hf.blocks), vec![0]);
    }

    // ── §17.6.22 speculative build (plan §1) ────────────────────────────────

    /// A speculative build must leave every **document-order** counter exactly
    /// as it found it. Measuring the *next* section's header to learn its
    /// clearance builds that header's content, and
    /// [`build_header_footer_content`] deliberately drains footnote references
    /// and advances list counters — so without rollback, peeking ahead would
    /// renumber §17.9 lists and §17.11.2 endnotes in document order.
    #[test]
    fn speculatively_rolls_back_document_order_counters() {
        let mut state = BuildState::default();
        let num = (model::NumId::new(1), 0u8);
        state.endnote_counter = 7;
        state.list_counters.insert(num, 3);
        state.outline = OutlineCollector::Collecting(5);
        let before_width = state.page_config.content_width();

        state.speculatively(|s| {
            s.endnote_counter += 1;
            *s.list_counters.entry(num).or_default() += 1;
            s.list_counters.insert((model::NumId::new(99), 0), 42);
            s.next_outline_node_id();
            s.page_config = crate::render::layout::page::PageConfig::default();
            s.page_config.margins.left = Pt::new(999.0);
        });

        assert_eq!(state.endnote_counter, 7, "§17.11.2 endnote number restored");
        assert_eq!(
            state.list_counters.get(&num).copied(),
            Some(3),
            "§17.9 list counter restored"
        );
        assert!(
            !state
                .list_counters
                .contains_key(&(model::NumId::new(99), 0)),
            "a counter the speculative build *created* must not survive it"
        );
        assert_eq!(
            state.outline,
            OutlineCollector::Collecting(5),
            "§17.3.1.19 node IDs must not be consumed speculatively"
        );
        assert_eq!(
            state.page_config.content_width(),
            before_width,
            "the peek measures the *next* section's config; the current one \
             must survive it"
        );
    }

    /// Warn-once sets are deliberately **not** rolled back. The peek builds the
    /// same content the real measurement will, so restoring them would report
    /// the same unsupported feature twice — the exact duplication the
    /// warn-once sets exist to prevent.
    #[test]
    fn speculatively_keeps_warn_once_state() {
        let mut state = BuildState::default();
        state.speculatively(|s| {
            s.warned_row_cell_spacing = true;
            s.warned_border_styles.insert(model::BorderStyle::Dotted);
        });
        assert!(
            state.warned_row_cell_spacing,
            "a warning already emitted must not be re-armed"
        );
        assert!(state
            .warned_border_styles
            .contains(&model::BorderStyle::Dotted));
    }

    /// The measured clearance is the whole point of the peek, so the value
    /// crosses the rollback boundary.
    #[test]
    fn speculatively_returns_the_closure_value() {
        let mut state = BuildState::default();
        let measured = state.speculatively(|s| {
            s.endnote_counter = 3;
            Pt::new(61.0)
        });
        assert_eq!(measured, Pt::new(61.0));
        assert_eq!(state.endnote_counter, 0, "and still rolls back");
    }

    /// Section breaks inside header/footer content are structural, not laid out.
    #[test]
    fn header_section_breaks_produce_no_blocks() {
        let resolved = empty_resolved();
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut state = BuildState::default();

        let blocks = vec![Block::SectionBreak(Box::default()), empty_para()];
        let hf = build_header_footer_content(&blocks, &ctx, &mut state);
        assert_eq!(hf.blocks.len(), 1, "only the paragraph survives");
    }
}

/// §17.3.1.19: whether a build contributes to the PDF document outline.
///
/// A PDF structure node ID must be unique across the whole document, so the
/// counter is document-wide — like `list_counters`, and for the same reason.
/// The variant matters because a shape text body is laid out by a **fresh**
/// `BuildState`: left `Collecting`, its counter would restart at 1 and hand out
/// IDs already owned by body headings, silently retargeting their outline
/// entries at the shape.
///
/// It is also the right answer independently. A heading inside a text box is
/// not a position in the document's main story, which is what an outline
/// enumerates — Word's own navigation pane leaves them out too.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OutlineCollector {
    /// The document body. Holds how many headings it has taken IDs for.
    Collecting(i32),
    /// A sub-layout whose headings are not document positions.
    Excluded,
}

impl Default for OutlineCollector {
    fn default() -> Self {
        Self::Collecting(0)
    }
}

/// Everything in [`BuildState`] that is a *position in the document* rather
/// than a property of the render as a whole.
///
/// Captured by destructuring so the exhaustiveness is enforced by the compiler:
/// a new `BuildState` field does not compile here until someone decides which
/// of the two categories it belongs to.
struct DocumentPosition {
    page_config: crate::render::layout::page::PageConfig,
    footnotes: crate::render::layout::fragment::FootnoteTracker,
    endnote_counter: u32,
    list_counters: HashMap<(model::NumId, u8), u32>,
    outline: OutlineCollector,
    field_ctx: crate::render::layout::fragment::FieldContext,
    shape_default_text_color: Option<crate::render::resolve::color::RgbColor>,
    shape_default_font_family: Option<String>,
    shape_auto_fit: crate::render::layout::ShapeAutoFit,
}

impl DocumentPosition {
    fn capture(state: &BuildState) -> Self {
        let BuildState {
            page_config,
            footnotes,
            endnote_counter,
            list_counters,
            outline,
            field_ctx,
            shape_default_text_color,
            shape_default_font_family,
            shape_auto_fit,
            // Warn-once state is deliberately *not* position state — see
            // `BuildState::speculatively` for why it is not rolled back.
            warned_border_styles: _,
            warned_row_cell_spacing: _,
            warned_orphan_vmerge: _,
        } = state;
        Self {
            page_config: page_config.clone(),
            footnotes: footnotes.clone(),
            endnote_counter: *endnote_counter,
            list_counters: list_counters.clone(),
            outline: *outline,
            field_ctx: *field_ctx,
            shape_default_text_color: *shape_default_text_color,
            shape_default_font_family: shape_default_font_family.clone(),
            shape_auto_fit: *shape_auto_fit,
        }
    }

    fn restore(self, state: &mut BuildState) {
        state.page_config = self.page_config;
        state.footnotes = self.footnotes;
        state.endnote_counter = self.endnote_counter;
        state.list_counters = self.list_counters;
        state.outline = self.outline;
        state.field_ctx = self.field_ctx;
        state.shape_default_text_color = self.shape_default_text_color;
        state.shape_default_font_family = self.shape_default_font_family;
        state.shape_auto_fit = self.shape_auto_fit;
    }
}

impl BuildState {
    /// Run `f` against this state and roll back every document-order counter it
    /// advanced, returning whatever `f` produced.
    ///
    /// §17.6.22: laying out a section needs to know the bounds of the page its
    /// *last* page will share with a following `Continuous` section, which
    /// means measuring that following section's header before reaching it.
    /// Measuring builds the header's content, and
    /// [`build_header_footer_content`] deliberately drains footnote references
    /// (§17.11.12) and advances list counters (§17.9) — so an unguarded peek
    /// renumbers the document in a way that depends on the renderer's
    /// look-ahead rather than on the file.
    ///
    /// A scope rather than a `capture`/`restore` pair, so a caller cannot hold
    /// a snapshot and forget to put it back: the only way to peek is to peek
    /// speculatively.
    ///
    /// **Warn-once state is not rolled back.** `warned_border_styles` and
    /// `warned_row_cell_spacing` record that a diagnostic has already been
    /// emitted this render. The peek builds the same content the real
    /// measurement will build moments later, so restoring them would report the
    /// same unsupported feature twice — the exact duplication those sets exist
    /// to prevent. They are render-scoped, not position-scoped.
    pub fn speculatively<R>(&mut self, f: impl FnOnce(&mut Self) -> R) -> R {
        let saved = DocumentPosition::capture(self);
        let result = f(self);
        saved.restore(self);
        result
    }

    /// §17.3.1.19: reserve the next PDF structure node ID, or `None` when this
    /// build contributes nothing to the outline.
    pub fn next_outline_node_id(&mut self) -> Option<i32> {
        match &mut self.outline {
            OutlineCollector::Collecting(count) => {
                // 1-based: `skia_safe::pdf::node_id` reserves 0 for "no node"
                // and every negative value for artifacts, so a 0-based counter
                // would silently leave the first heading unmarked.
                *count += 1;
                Some(*count)
            }
            OutlineCollector::Excluded => None,
        }
    }
}
