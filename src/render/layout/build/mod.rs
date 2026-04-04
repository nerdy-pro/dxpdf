//! Recursive tree-walk that converts a resolved document model into layout blocks.
//!
//! The document tree (Section → Block → Paragraph | Table → Cell → Block…) is
//! processed by recursive descent.  Each `Block::Table` recurses into its cells'
//! content, which may contain nested tables.

pub(super) mod block;
pub(super) mod convert;
pub(super) mod table;

use std::collections::HashMap;
use std::rc::Rc;

use crate::model::{self, Block};
use crate::render::dimension::Pt;
use crate::render::layout::fragment::Fragment;
use crate::render::layout::measurer::TextMeasurer;
use crate::render::layout::page::PageConfig;
use crate::render::layout::paragraph::ParagraphStyle;
use crate::render::layout::section::LayoutBlock;
use crate::render::resolve::sections::ResolvedSection;
use crate::render::resolve::ResolvedDocument;

use block::{
    build_block, build_fragments, collect_endnotes, extract_floating_images,
    find_vml_absolute_position,
};
use convert::{
    doc_font_family, doc_font_size, paragraph_style_from_props, resolve_paragraph_defaults,
};
use table::build_table;

/// §17.8.3.2: OOXML fallback font when no theme or doc defaults specify one.
pub(super) const SPEC_FALLBACK_FONT: &str = "Times New Roman";
/// §17.3.2.14: default font size (10pt = 20 half-points per ECMA-376 §17.3.2.14).
pub(super) const SPEC_DEFAULT_FONT_SIZE: Pt = Pt::new(10.0);

// ── Context ─────────────────────────────────────────────────────────────────

/// Immutable context threaded through the recursive tree walk.
pub struct BuildContext<'a> {
    pub measurer: &'a TextMeasurer,
    pub resolved: &'a ResolvedDocument,
}

impl BuildContext<'_> {
    pub(super) fn media(&self) -> &HashMap<model::RelId, Rc<[u8]>> {
        &self.resolved.media
    }
}

/// Mutable state threaded through the recursive tree walk.
#[derive(Default)]
pub struct BuildState {
    /// Page configuration for the current section.
    pub page_config: crate::render::layout::page::PageConfig,
    /// Sequential footnote display number (1, 2, 3...).
    pub footnote_counter: u32,
    /// Sequential endnote display number (i, ii, iii...).
    pub endnote_counter: u32,
    /// Per-(numId, level) running counters for list labels.
    pub list_counters: HashMap<(model::NumId, u8), u32>,
    /// Field evaluation context (page number, total pages).
    pub field_ctx: crate::render::layout::fragment::FieldContext,
}

// ── Public entry point ──────────────────────────────────────────────────────

/// Built section output — layout blocks plus endnotes.
pub struct BuiltSection {
    pub blocks: Vec<LayoutBlock>,
    /// Endnote content (display number, fragments, style) — rendered at document end.
    pub endnotes: Vec<(String, Vec<Fragment>, ParagraphStyle)>,
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

    // Collect endnotes (rendered at document end).
    let mut endnotes = Vec::new();
    collect_endnotes(ctx, state, &mut endnotes);

    BuiltSection { blocks, endnotes }
}

/// Collected header/footer content with layout metadata.
pub struct HeaderFooterContent {
    /// Layout blocks (paragraphs and tables) for stacking.
    pub blocks: Vec<LayoutBlock>,
    /// Absolute page-relative position from a VML text box, if present.
    pub absolute_position: Option<(Pt, Pt)>,
    /// Floating (anchor) images from header/footer paragraphs.
    pub floating_images: Vec<crate::render::layout::section::FloatingImage>,
}

/// Build header/footer content from blocks.
///
/// Produces `LayoutBlock` entries for both paragraphs and tables, and
/// extracts floating images separately (they are positioned page-relative
/// rather than stack-relative).
pub fn build_header_footer_content(
    blocks: &[Block],
    ctx: &BuildContext,
    state: &mut BuildState,
) -> HeaderFooterContent {
    let mut layout_blocks = Vec::new();
    let mut all_floating_images = Vec::new();
    let mut absolute_position: Option<(Pt, Pt)> = None;

    let available_width = state.page_config.content_width();

    let block_count = blocks.len();
    for (block_i, block) in blocks.iter().enumerate() {
        match block {
            Block::Paragraph(p) => {
                let (mut frags, props) = build_fragments(p, ctx, state, None, None);
                let style = paragraph_style_from_props(&props);

                // Check for VML absolute positioning in Pict inlines.
                if absolute_position.is_none() {
                    for inline in &p.content {
                        if let Some(pos) = find_vml_absolute_position(inline) {
                            absolute_position = Some(pos);
                            break;
                        }
                    }
                }
                // Extract floating (anchor) images — positioned page-relative.
                let floats = extract_floating_images(p, ctx, state, false);
                all_floating_images.extend(floats);

                // §17.10.1: empty non-last paragraphs in headers/footers still
                // occupy a line height (from the paragraph mark's font size).
                if frags.is_empty() && block_i + 1 < block_count {
                    let (family, mut size, ..) = resolve_paragraph_defaults(p, ctx.resolved, false);
                    if let Some(ref mrp) = p.mark_run_properties {
                        if let Some(fs) = mrp.font_size {
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
                });
            }
            Block::Table(t) => {
                let built = build_table(t, available_width, ctx, state);
                layout_blocks.push(LayoutBlock::Table {
                    rows: built.rows,
                    col_widths: built.col_widths,
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
    }
}

/// Default line height derived from document-level font settings.
pub fn default_line_height(ctx: &BuildContext) -> Pt {
    let family = doc_font_family(ctx);
    let size = doc_font_size(ctx);
    ctx.measurer.default_line_height(&family, size)
}
