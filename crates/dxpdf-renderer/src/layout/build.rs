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
use crate::geometry::{self, PtSize};
use crate::layout::fragment::{collect_fragments, FontProps, Fragment};
use crate::layout::measurer::TextMeasurer;
use crate::layout::page::PageConfig;
use crate::layout::paragraph::{
    BorderLine, DropCapInfo, LineSpacingRule, ParagraphBorderStyle, ParagraphStyle,
};
use crate::layout::paragraph::TabStopDef;
use crate::layout::section::LayoutBlock;
use crate::layout::table::{
    CellBorderConfig, CellBorderOverride, TableBorderConfig, TableBorderLine, TableBorderStyle,
    TableCellInput, TableRowInput, compute_column_widths,
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

/// Context threaded through the recursive tree walk.
pub struct BuildContext<'a> {
    pub measurer: &'a TextMeasurer,
    pub resolved: &'a ResolvedDocument,
    /// Page configuration for the current section.
    pub page_config: std::cell::RefCell<crate::layout::page::PageConfig>,
    /// Sequential footnote display number (1, 2, 3...).
    pub footnote_counter: std::cell::Cell<u32>,
    /// Sequential endnote display number (i, ii, iii...).
    pub endnote_counter: std::cell::Cell<u32>,
    /// Per-(numId, level) running counters for list labels.
    pub list_counters: std::cell::RefCell<HashMap<(model::NumId, u8), u32>>,
    /// Field evaluation context (page number, total pages).
    /// Uses Cell so header/footer rendering can set per-page values.
    pub field_ctx_cell: std::cell::Cell<crate::layout::fragment::FieldContext>,
}

impl BuildContext<'_> {
    fn media(&self) -> &HashMap<model::RelId, Vec<u8>> {
        &self.resolved.media
    }
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
) -> BuiltSection {
    let mut pending_dropcap: Option<DropCapInfo> = None;
    let blocks: Vec<LayoutBlock> = section
        .blocks
        .iter()
        .filter_map(|block| build_block(block, config.content_width(), ctx, &mut pending_dropcap))
        .collect();

    // Collect endnotes (rendered at document end).
    let mut endnotes = Vec::new();
    collect_endnotes(ctx, &mut endnotes);

    BuiltSection { blocks, endnotes }
}

/// Build note content (footnotes or endnotes) with a display number prefix.
fn build_note_content(
    _note_id_value: i64,
    display_num: &str,
    content: &[Block],
    ctx: &BuildContext,
) -> Vec<(String, Vec<Fragment>, ParagraphStyle)> {
    let mut results = Vec::new();
    for (i, block) in content.iter().enumerate() {
        if let model::Block::Paragraph(p) = block {
            let (mut frags, merged_props) = build_fragments(p, ctx, None, None);

            // Prepend display number to the first paragraph.
            if i == 0 && !frags.is_empty() {
                let num_text = format!("{}  ", display_num);
                let font = frags[0].font_props().cloned().unwrap_or_else(|| FontProps {
                    family: std::rc::Rc::from("Times New Roman"),
                    size: Pt::new(10.0),
                    bold: false, italic: false, underline: false,
                    char_spacing: Pt::ZERO,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                });
                let ref_size = font.size * 0.58;
                let ref_font = FontProps { size: ref_size, ..font };
                let (w, m) = ctx.measurer.measure(&num_text, &ref_font);
                frags.insert(0, Fragment::Text {
                    text: num_text,
                    font: ref_font,
                    color: RgbColor::BLACK,
                    shading: None, border: None,
                    width: w, trimmed_width: w,
                    metrics: m,
                    hyperlink_url: None,
                    baseline_offset: -(font.size * 0.4),
                    text_offset: Pt::ZERO,
                });
            }
            let style = paragraph_style_from_props(&merged_props);
            results.push((display_num.to_string(), frags, style));
        }
    }
    results
}

/// Collect endnotes from the resolved document.
fn collect_endnotes(
    ctx: &BuildContext,
    endnotes: &mut Vec<(String, Vec<Fragment>, ParagraphStyle)>,
) {
    // IDs 0 and 1 are reserved for separator and continuation separator.
    let mut en_ids: Vec<_> = ctx.resolved.endnotes.keys()
        .filter(|id| id.value() > 1)
        .collect();
    en_ids.sort_by_key(|id| id.value());
    for (i, note_id) in en_ids.iter().enumerate() {
        let display = crate::layout::fragment::to_roman_lower((i + 1) as u32);
        if let Some(content) = ctx.resolved.endnotes.get(note_id) {
            endnotes.extend(build_note_content(note_id.value(), &display, content, ctx));
        }
    }
}

/// Collected header/footer content with layout metadata.
pub struct HeaderFooterContent {
    /// Layout blocks (paragraphs and tables) for stacking.
    pub blocks: Vec<LayoutBlock>,
    /// Absolute page-relative position from a VML text box, if present.
    pub absolute_position: Option<(Pt, Pt)>,
    /// Floating (anchor) images from header/footer paragraphs.
    pub floating_images: Vec<crate::layout::section::FloatingImage>,
}

/// Build header/footer content from blocks.
///
/// Produces `LayoutBlock` entries for both paragraphs and tables, and
/// extracts floating images separately (they are positioned page-relative
/// rather than stack-relative).
pub fn build_header_footer_content(
    blocks: &[Block],
    ctx: &BuildContext,
) -> HeaderFooterContent {
    let mut layout_blocks = Vec::new();
    let mut all_floating_images = Vec::new();
    let mut absolute_position: Option<(Pt, Pt)> = None;

    let available_width = ctx.page_config.borrow().content_width();

    let block_count = blocks.len();
    for (block_i, block) in blocks.iter().enumerate() {
        match block {
            Block::Paragraph(p) => {
                let (mut frags, props) = build_fragments(p, ctx, None, None);
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
                let floats = extract_floating_images(p, ctx, false);
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
                let built = build_table(t, available_width, ctx);
                layout_blocks.push(LayoutBlock::Table {
                    rows: built.rows,
                    col_widths: built.col_widths,
                    border_config: built.border_config,
                    indent: built.indent,
                    alignment: built.alignment,
                    float_info: built.float_info,
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

/// Search an inline (and AlternateContent fallback) for a VML text box with
/// absolute positioning.
fn find_vml_absolute_position(inline: &model::Inline) -> Option<(Pt, Pt)> {
    match inline {
        model::Inline::Pict(pict) => find_vml_pos_in_pict(pict),
        model::Inline::AlternateContent(ac) => {
            if let Some(ref fallback) = ac.fallback {
                for inner in fallback {
                    if let Some(pos) = find_vml_absolute_position(inner) {
                        return Some(pos);
                    }
                }
            }
            None
        }
        _ => None,
    }
}

fn find_vml_pos_in_pict(pict: &model::Pict) -> Option<(Pt, Pt)> {
    for shape in &pict.shapes {
        if shape.text_box.is_some() {
            if let Some(pos) = vml_absolute_position(&shape.style) {
                return Some(pos);
            }
        }
    }
    None
}

/// Extract absolute page-relative position from a VML shape style, in points.
fn vml_absolute_position(style: &model::VmlStyle) -> Option<(Pt, Pt)> {
    use dxpdf_docx_model::model::CssPosition;
    if style.position != Some(CssPosition::Absolute) {
        return None;
    }
    let x = style.margin_left.map(vml_length_to_pt)?;
    let y = style.margin_top.map(vml_length_to_pt)?;
    Some((x, y))
}

/// Convert a VML CSS length to points.
fn vml_length_to_pt(len: model::VmlLength) -> Pt {
    use dxpdf_docx_model::model::VmlLengthUnit;
    let value = len.value as f32;
    Pt::new(match len.unit {
        VmlLengthUnit::Pt => value,
        VmlLengthUnit::In => value * 72.0,
        VmlLengthUnit::Cm => value * 72.0 / 2.54,
        VmlLengthUnit::Mm => value * 72.0 / 25.4,
        VmlLengthUnit::Px => value * 0.75, // 96dpi → 72pt/in
        VmlLengthUnit::None => value / 914400.0 * 72.0, // bare number = EMU
        _ => value, // Em, Percent — fallback to raw value
    })
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
        Block::Paragraph(p) => build_paragraph_block(p, ctx, pending_dropcap, None, None),
        Block::Table(t) => {
            let built = build_table(t, available_width, ctx);
            Some(LayoutBlock::Table {
                rows: built.rows,
                col_widths: built.col_widths,
                border_config: built.border_config,
                indent: built.indent,
                alignment: built.alignment,
                float_info: built.float_info,
            })
        }
        Block::SectionBreak(_) => None,
    }
}

// ── Paragraph building ──────────────────────────────────────────────────────

/// Build a paragraph into a layout block.
/// Handles drop cap detection (§17.3.1.11), list labels, floating images.
/// For table cells, pass `table_style` and `cond` to apply table formatting cascade.
fn build_paragraph_block(
    p: &Paragraph,
    ctx: &BuildContext,
    pending_dropcap: &mut Option<DropCapInfo>,
    table_style: Option<&ResolvedStyle>,
    cond: Option<&CellConditionalFormatting>,
) -> Option<LayoutBlock> {
    let (mut fragments, mut merged_props) = build_fragments(p, ctx, table_style, cond);

    // §17.9.22: inject list label if paragraph has a numbering reference.
    if let Some(ref num_ref) = merged_props.numbering {
        let num_id = model::NumId::new(num_ref.num_id);
        let level = num_ref.level;
        if let Some(levels) = ctx.resolved.numbering.get(&num_id) {
            // Update counters: increment this level, reset deeper levels.
            {
                let mut counters = ctx.list_counters.borrow_mut();
                let count = counters.entry((num_id, level)).or_insert_with(|| {
                    levels.get(level as usize).map(|l| l.start).unwrap_or(1) - 1
                });
                *count += 1;
                // Reset deeper levels.
                let max_level = levels.len() as u8;
                for deeper in (level + 1)..max_level {
                    counters.remove(&(num_id, deeper));
                }
            }

            let level_def = levels.get(level as usize);

            // §17.9.10: check for picture bullet before text label.
            let pic_bullet_injected = level_def
                .and_then(|l| l.lvl_pic_bullet_id)
                .and_then(|pic_id| ctx.resolved.pic_bullets.get(&pic_id))
                .and_then(|bullet| {
                    let rel_id = bullet.pict.as_ref()?
                        .shapes.first()?
                        .image_data.as_ref()?
                        .rel_id.as_ref()?;
                    let image_bytes = ctx.media().get(rel_id)?;
                    // Size from VML shape style (width/height), default 9pt.
                    let size = pic_bullet_size(bullet);
                    let label_frag = Fragment::Image {
                        size,
                        rel_id: rel_id.as_str().to_string(),
                        image_data: Some(image_bytes.as_slice().into()),
                    };
                    Some((label_frag, size.height))
                });

            if let Some((label_frag, label_height)) = pic_bullet_injected {
                let hanging = level_def
                    .and_then(|l| l.indentation.as_ref())
                    .and_then(|ind| ind.first_line)
                    .map(|fl| match fl {
                        model::FirstLineIndent::Hanging(v) => Pt::from(v),
                        _ => Pt::ZERO,
                    })
                    .unwrap_or(Pt::ZERO);
                let tab_frag = Fragment::Tab {
                    line_height: label_height,
                    fitting_width: Some(hanging),
                };
                fragments.insert(0, tab_frag);
                fragments.insert(0, label_frag);

                if let Some(lvl_left) = level_def
                    .and_then(|l| l.indentation.as_ref())
                    .and_then(|ind| ind.start)
                {
                    merged_props.tabs.insert(0, dxpdf_docx_model::model::TabStop {
                        position: lvl_left,
                        alignment: dxpdf_docx_model::model::TabAlignment::Left,
                        leader: dxpdf_docx_model::model::TabLeader::None,
                    });
                }
            } else {
            let counters = ctx.list_counters.borrow();
            if let Some(label_text) = crate::resolve::numbering::format_list_label(
                levels, level, &counters, num_id,
            ) {
                // Resolve label font from level run_properties or paragraph defaults.
                let (default_family, default_size, default_color, _, _) =
                    resolve_paragraph_defaults(p, ctx.resolved, false);
                let level_font_family = level_def
                    .and_then(|l| l.run_properties.as_ref())
                    .and_then(|rp| crate::resolve::fonts::effective_font(&rp.fonts))
                    .unwrap_or("");

                // Remap PUA codepoints from legacy Symbol/Wingdings encoding
                // to standard Unicode per official mapping tables, and use
                // a standard font. This is portable across all platforms.
                let (label_text, label_family) = remap_legacy_font_chars(
                    &label_text, level_font_family, &default_family,
                );
                let label_family: std::rc::Rc<str> = std::rc::Rc::from(label_family.as_str());

                let label_size = level_def
                    .and_then(|l| l.run_properties.as_ref())
                    .and_then(|rp| rp.font_size)
                    .map(Pt::from)
                    .unwrap_or(default_size);
                let label_bold = level_def
                    .and_then(|l| l.run_properties.as_ref())
                    .and_then(|rp| rp.bold)
                    .unwrap_or(false);
                let label_italic = level_def
                    .and_then(|l| l.run_properties.as_ref())
                    .and_then(|rp| rp.italic)
                    .unwrap_or(false);

                let label_font = FontProps {
                    family: label_family,
                    size: label_size,
                    bold: label_bold,
                    italic: label_italic,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                };
                let (w, m) = ctx.measurer.measure(&label_text, &label_font);
                let h = m.height();
                // Tab after label: advances to indent_left via the implicit
                // tab stop. Fitting width = hanging so that label + tab
                // consume exactly the hanging indent space during fitting,
                // leaving content_width for the body text.
                let hanging = level_def
                    .and_then(|l| l.indentation.as_ref())
                    .and_then(|ind| ind.first_line)
                    .map(|fl| match fl {
                        model::FirstLineIndent::Hanging(v) => Pt::from(v),
                        _ => Pt::ZERO,
                    })
                    .unwrap_or(Pt::ZERO);
                // §17.9.7: lvlJc controls label justification within the
                // hanging indent area. The label fragment occupies `w`
                // points but the text is drawn at `text_offset` within it.
                // The tab then advances from `w` to the indent tab stop,
                // producing a natural gap = hanging − w.
                let jc = level_def.and_then(|l| l.justification);
                let text_offset = match jc {
                    Some(dxpdf_docx_model::model::Alignment::End) => -w,
                    Some(dxpdf_docx_model::model::Alignment::Center) => w * -0.5,
                    _ => Pt::ZERO,
                };
                // The label fragment width is the text width only.
                // text_offset shifts where the text is drawn (for right/
                // center justification) but x advances by w, leaving
                // room for the tab to fill hanging − w to the stop.
                let label_width = w;
                let label_frag = Fragment::Text {
                    text: label_text,
                    font: label_font.clone(),
                    color: default_color,
                    shading: None,
                    border: None,
                    width: label_width,
                    trimmed_width: label_width,
                    metrics: m,
                    hyperlink_url: None,
                    baseline_offset: Pt::ZERO,
                    text_offset,
                };
                let tab_fitting = (hanging - label_width).max(Pt::ZERO);
                let tab_frag = Fragment::Tab {
                    line_height: h,
                    fitting_width: Some(tab_fitting),
                };
                fragments.insert(0, tab_frag);
                fragments.insert(0, label_frag);

                // Add implicit tab stop at numLvl.left so the tab lands
                // at the body text position.
                let lvl_left = level_def
                    .and_then(|l| l.indentation.as_ref())
                    .and_then(|ind| ind.start);
                if let Some(lvl_left) = lvl_left {
                    merged_props.tabs.insert(0, dxpdf_docx_model::model::TabStop {
                        position: lvl_left,
                        alignment: dxpdf_docx_model::model::TabAlignment::Left,
                        leader: dxpdf_docx_model::model::TabLeader::None,
                    });
                }
            }
            }

            // §17.9.23: numbering level pPr overrides the paragraph style.
            // Only the paragraph's direct ind overrides the numbering level.
            if let Some(lvl_ind) = levels.get(level as usize).and_then(|l| l.indentation.as_ref()) {
                let mut ind = *lvl_ind;
                if let Some(direct) = p.properties.indentation {
                    if let Some(start) = direct.start {
                        ind.start = Some(start);
                    }
                    if let Some(end) = direct.end {
                        ind.end = Some(end);
                    }
                    if let Some(first_line) = direct.first_line {
                        ind.first_line = Some(first_line);
                    }
                }
                merged_props.indentation = Some(ind);
            }
        }
    }

    // Word suppresses Hyperlink character style (blue/underline) for ToC
    // entries in print view. Strip visual hyperlink styling but keep the
    // click annotation URL.
    if p.style_id.as_ref()
        .is_some_and(|id| id.as_str().starts_with("TOC") || id.as_str().starts_with("toc"))
    {
        for frag in &mut fragments {
            if let Fragment::Text { font, color, hyperlink_url, .. } = frag {
                if hyperlink_url.is_some() {
                    *color = RgbColor::BLACK;
                    font.underline = false;
                }
            }
        }
    }

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
                Fragment::Text { metrics, .. } => metrics.ascent,
                _ => Pt::ZERO,
            })
            .fold(Pt::ZERO, Pt::max);
        let h_space = merged_props
            .frame_properties
            .and_then(|fp| fp.h_space)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO);
        let margin_mode = merged_props
            .frame_properties
            .and_then(|fp| fp.drop_cap)
            .is_some_and(|dc| matches!(dc, model::DropCap::Margin));
        // The drop cap paragraph's own indent determines the x position.
        // This includes indent_left + indent_first_line from the cascade.
        let dc_indent_left = merged_props.indentation
            .and_then(|i| i.start).map(Pt::from).unwrap_or(Pt::ZERO);
        let dc_indent_first = merged_props.indentation
            .and_then(|i| i.first_line)
            .map(|fl| match fl {
                model::FirstLineIndent::FirstLine(v) => Pt::from(v),
                model::FirstLineIndent::Hanging(v) => -Pt::from(v),
                model::FirstLineIndent::None => Pt::ZERO,
            })
            .unwrap_or(Pt::ZERO);
        // §17.3.1.33: frame height from drop cap paragraph's exact line spacing.
        let frame_height = merged_props.spacing
            .and_then(|s| s.line)
            .and_then(|ls| match ls {
                model::LineSpacing::Exact(v) => Some(Pt::from(v)),
                _ => None,
            });
        // §17.3.2.19: position offset from the drop cap run.
        let position_offset = fragments.first()
            .and_then(|f| match f {
                Fragment::Text { baseline_offset, .. } => Some(*baseline_offset),
                _ => None,
            })
            .unwrap_or(Pt::ZERO);
        *pending_dropcap = Some(DropCapInfo {
            fragments,
            lines: drop_cap_lines,
            ascent,
            h_space,
            width,
            height,
            margin_mode,
            indent: dc_indent_left + dc_indent_first,
            frame_height,
            position_offset,
        });
        return None;
    }

    let mut style = paragraph_style_from_props(&merged_props);
    style.style_id = p.style_id.clone();

    // Attach pending drop cap to this paragraph.
    if let Some(dc) = pending_dropcap.take() {
        style.drop_cap = Some(dc);
    }

    let page_break_before = merged_props.page_break_before.unwrap_or(false);

    // Collect footnotes referenced in this paragraph.
    // The footnote_counter was already incremented during fragment collection,
    // so we count backwards to get the display number for each reference.
    let fn_refs: Vec<_> = p.content.iter()
        .filter_map(|i| if let model::Inline::FootnoteRef(id) = i { Some(id) } else { None })
        .collect();
    let fn_base = ctx.footnote_counter.get() - fn_refs.len() as u32;
    let mut para_footnotes = Vec::new();
    for (i, note_id) in fn_refs.iter().enumerate() {
        let display = format!("{}", fn_base + i as u32 + 1);
        if let Some(content) = ctx.resolved.footnotes.get(note_id) {
            let notes = build_note_content(note_id.value(), &display, content, ctx);
            for (_, frags, style) in notes {
                para_footnotes.push((frags, style));
            }
        }
    }

    // §20.4.2.3: extract floating (anchor) images from this paragraph.
    // In cell context, positions are cell-relative instead of page-relative.
    let cell_context = table_style.is_some();
    let floating_images = extract_floating_images(p, ctx, cell_context);

    Some(LayoutBlock::Paragraph {
        fragments,
        style,
        page_break_before,
        footnotes: para_footnotes,
        floating_images,
    })
}

/// Extract floating (anchor) images from a paragraph's inlines.
/// When `cell_context` is true, positions are resolved relative to the cell
/// origin (0,0) instead of the page margins.
fn extract_floating_images(
    para: &Paragraph,
    ctx: &BuildContext,
    cell_context: bool,
) -> Vec<crate::layout::section::FloatingImage> {
    use dxpdf_docx_model::model::{ImagePlacement, AnchorPosition, AnchorRelativeFrom, AnchorAlignment, Inline};
    use crate::layout::section::{FloatingImage, FloatingImageY};

    let mut images = Vec::new();

    fn find_anchor_images<'a>(inlines: &'a [Inline], out: &mut Vec<&'a dxpdf_docx_model::model::Image>) {
        for inline in inlines {
            match inline {
                Inline::Image(img) => {
                    if matches!(img.placement, ImagePlacement::Anchor(_)) {
                        out.push(img);
                    }
                }
                Inline::Hyperlink(link) => find_anchor_images(&link.content, out),
                Inline::Field(f) => find_anchor_images(&f.content, out),
                Inline::AlternateContent(ac) => {
                    if let Some(ref fb) = ac.fallback {
                        find_anchor_images(fb, out);
                    }
                }
                _ => {}
            }
        }
    }

    let mut anchor_imgs = Vec::new();
    find_anchor_images(&para.content, &mut anchor_imgs);

    for img in &anchor_imgs {
        if let ImagePlacement::Anchor(ref anchor) = img.placement {
            let rel_id = match crate::resolve::images::extract_image_rel_id(img) {
                Some(id) => id,
                None => {
                    eprintln!("  -> no rel_id, graphic.is_some()={}", img.graphic.is_some());
                    continue;
                }
            };

            let image_data = match ctx.resolved.media.get(rel_id) {
                Some(bytes) => std::rc::Rc::from(bytes.as_slice()),
                None => {
                    eprintln!("Anchor image: rel_id={} NOT FOUND in media (media has {} entries)", rel_id.as_str(), ctx.resolved.media.len());
                    continue;
                }
            };

                let w = Pt::from(img.extent.width);
                let h = Pt::from(img.extent.height);
                let pc = ctx.page_config.borrow();

                // Resolve horizontal position.
                // In cell context, positions are relative to the cell origin.
                let (page_width, margin_left, margin_right) = if cell_context {
                    (Pt::ZERO, Pt::ZERO, Pt::ZERO)
                } else {
                    (pc.page_size.width, pc.margins.left, pc.margins.right)
                };
                let content_width = if cell_context { Pt::ZERO } else { page_width - margin_left - margin_right };

                let x = match &anchor.horizontal_position {
                    AnchorPosition::Offset { relative_from, offset } => {
                        let base = match relative_from {
                            AnchorRelativeFrom::Page => Pt::ZERO,
                            AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => margin_left,
                            _ => margin_left,
                        };
                        base + Pt::from(*offset)
                    }
                    AnchorPosition::Align { relative_from, alignment } => {
                        let (area_left, area_width) = match relative_from {
                            AnchorRelativeFrom::Page => (Pt::ZERO, page_width),
                            AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => (margin_left, content_width),
                            _ => (margin_left, content_width),
                        };
                        match alignment {
                            AnchorAlignment::Left => area_left,
                            AnchorAlignment::Right => area_left + area_width - w,
                            AnchorAlignment::Center => area_left + (area_width - w) * 0.5,
                            _ => area_left,
                        }
                    }
                };

                // Resolve vertical position.
                let y = match &anchor.vertical_position {
                    AnchorPosition::Offset { relative_from, offset } => {
                        let margin_top = if cell_context { Pt::ZERO } else { pc.margins.top };
                        if cell_context {
                            // In cell context, all positions are relative to cell origin.
                            FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                        } else {
                            match relative_from {
                                AnchorRelativeFrom::Page => FloatingImageY::Absolute(Pt::from(*offset)),
                                AnchorRelativeFrom::Margin => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                                // §20.4.2.11: topMargin — offset from page top.
                                AnchorRelativeFrom::TopMargin => FloatingImageY::Absolute(Pt::from(*offset)),
                                // §20.4.2.11: bottomMargin — offset from bottom margin edge.
                                AnchorRelativeFrom::BottomMargin => {
                                    let page_height = pc.page_size.height;
                                    let margin_bottom = pc.margins.bottom;
                                    FloatingImageY::Absolute(page_height - margin_bottom + Pt::from(*offset))
                                }
                                AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                                    FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                                }
                                _ => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                            }
                        }
                    }
                    AnchorPosition::Align { relative_from, alignment } => {
                        let margin_top = if cell_context { Pt::ZERO } else { pc.margins.top };
                        let page_height = if cell_context { Pt::ZERO } else { pc.page_size.height };
                        let margin_bottom = if cell_context { Pt::ZERO } else { pc.margins.bottom };
                        let (area_top, area_height) = match relative_from {
                            AnchorRelativeFrom::Page => (Pt::ZERO, page_height),
                            AnchorRelativeFrom::Margin => (margin_top, page_height - margin_top - margin_bottom),
                            // §20.4.2.11: topMargin = area from page top to top margin edge.
                            AnchorRelativeFrom::TopMargin => (Pt::ZERO, margin_top),
                            // §20.4.2.11: bottomMargin = area from bottom margin edge to page bottom.
                            AnchorRelativeFrom::BottomMargin => (page_height - margin_bottom, margin_bottom),
                            _ => (margin_top, page_height - margin_top - margin_bottom),
                        };
                        let y_pos = match alignment {
                            AnchorAlignment::Top => area_top,
                            AnchorAlignment::Bottom => area_top + area_height - h,
                            AnchorAlignment::Center => area_top + (area_height - h) * 0.5,
                            _ => area_top,
                        };
                        FloatingImageY::Absolute(y_pos)
                    }
                };

                images.push(FloatingImage {
                    image_data,
                    size: crate::geometry::PtSize::new(w, h),
                    x,
                    y,
                    wrap_top_and_bottom: matches!(anchor.wrap, dxpdf_docx_model::model::TextWrap::TopAndBottom { .. }),
                    dist_left: Pt::from(anchor.distance.left),
                    dist_right: Pt::from(anchor.distance.right),
                });
            }
        }

    images
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
    // Clone paragraph for style resolution.
    let effective_para = para.clone();

    // §17.7.2: resolve paragraph defaults (direct → paragraph style).
    // Doc defaults are deferred so table style/conditional can be inserted
    // between paragraph style and doc defaults in the cascade.
    let (default_family, mut default_size, mut default_color, mut merged_props, mut run_defaults) =
        resolve_paragraph_defaults(&effective_para, ctx.resolved, table_style.is_some());

    // §17.7.2: table conditional formatting — lower priority than paragraph style.
    if let Some(c) = cond {
        if let Some(ref pp) = c.paragraph_properties {
            merge_paragraph_properties(&mut merged_props, pp);
        }
    }
    // §17.7.2: table style paragraph properties — lower priority than conditional.
    if let Some(ts) = table_style {
        merge_paragraph_properties(&mut merged_props, &ts.paragraph);
    }
    // §17.7.2: doc defaults — lowest priority, deferred from resolve_paragraph_defaults.
    if table_style.is_some() {
        merge_paragraph_properties(&mut merged_props, &ctx.resolved.doc_defaults_paragraph);
    }

    // §17.7.2: table style run properties override Normal.
    if let Some(ts) = table_style {
        if let Some(fs) = ts.run.font_size {
            default_size = Pt::from(fs);
            run_defaults.font_size = Some(fs);
        }
    }

    // §17.7.6: conditional run property overrides — higher priority than
    // table style and paragraph style. Overlay (not merge): conditional
    // values replace existing ones.
    if let Some(c) = cond {
        if let Some(ref rp) = c.run_properties {
            // Overlay: for each Some field in rp, replace in run_defaults.
            let mut overlay = rp.clone();
            merge_run_properties(&mut overlay, &run_defaults);
            run_defaults = overlay;
            if let Some(fs) = run_defaults.font_size {
                default_size = Pt::from(fs);
            }
            if let Some(color) = run_defaults.color {
                default_color = resolve_color(color, ColorContext::Text);
            }
        }
    }

    let measure = |text: &str, font: &FontProps| -> (Pt, crate::layout::fragment::TextMetrics) {
        ctx.measurer.measure(text, font)
    };

    let mut fn_counter = ctx.footnote_counter.get();
    let mut en_counter = ctx.endnote_counter.get();
    let mut fragments = collect_fragments(
        &para.content,
        &default_family,
        default_size,
        default_color,
        None,
        &measure,
        Some(&ctx.resolved.styles),
        Some(&run_defaults),
        &mut fn_counter,
        &mut en_counter,
        ctx.field_ctx_cell.get(),
        ctx.resolved.theme.as_ref(),
    );
    ctx.footnote_counter.set(fn_counter);
    ctx.endnote_counter.set(en_counter);
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
    /// §17.4.51: table indentation from left margin.
    indent: Pt,
    /// §17.4.28: table horizontal alignment (left/center/right).
    alignment: Option<model::Alignment>,
    float_info: Option<super::section::TableFloatInfo>,
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
    let default_cell_margins = t.properties.cell_margins.or(style_cell_margins);

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
    let col_widths = if is_auto_width && !grid_cols.is_empty() {
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
                height_rule: row.properties.height.map(|h| {
                    use dxpdf_docx_model::model::HeightRule;
                    use crate::layout::table::RowHeightRule;
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
        super::section::TableFloatInfo {
            right_gap: pos.right_from_text.map(Pt::from).unwrap_or(Pt::ZERO),
            bottom_gap: pos.bottom_from_text.map(Pt::from).unwrap_or(Pt::ZERO),
            x_align: pos.x_align,
            // §17.4.59: tblpY — absolute Y offset from the vertical anchor.
            y_offset: pos.y.map(Pt::from).unwrap_or(Pt::ZERO),
            // §17.4.58: default vertical anchor is "text".
            vert_anchor: pos.vert_anchor.unwrap_or(dxpdf_docx_model::model::TableAnchor::Text),
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
        _ if is_full_width && is_left_aligned => {
            -default_cell_margins.map(|m| Pt::from(m.left)).unwrap_or(Pt::ZERO)
        }
        _ => Pt::ZERO,
    };

    BuiltTable { rows, col_widths, border_config, indent, alignment: t.properties.alignment, float_info }
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
                bottom: convert_cell_border_override(&db.bottom)
                    .or_else(|| cond_borders.and_then(|cb| convert_cell_border_override(&cb.bottom))),
                left: convert_cell_border_override(&db.left)
                    .or_else(|| cond_borders.and_then(|cb| convert_cell_border_override(&cb.left))),
                right: convert_cell_border_override(&db.right)
                    .or_else(|| cond_borders.and_then(|cb| convert_cell_border_override(&cb.right))),
            })
        }
        (None, Some(cb)) => {
            Some(CellBorderConfig {
                top: convert_cell_border_override(&cb.top),
                bottom: convert_cell_border_override(&cb.bottom),
                left: convert_cell_border_override(&cb.left),
                right: convert_cell_border_override(&cb.right),
            })
        }
        (None, None) => None,
    };

    // §17.4.84: vertical alignment — direct cell, conditional, or default top.
    let valign = cell.properties.vertical_align
        .or_else(|| cond.cell_properties.as_ref().and_then(|tcp| tcp.vertical_align))
        .map(|va| match va {
            model::CellVerticalAlign::Bottom => crate::layout::table::CellVAlign::Bottom,
            model::CellVerticalAlign::Center => crate::layout::table::CellVAlign::Center,
            _ => crate::layout::table::CellVAlign::Top,
        })
        .unwrap_or(crate::layout::table::CellVAlign::Top);

    // Estimate border insets to compute effective content width for
    // character-level splitting of oversized fragments.
    let border_w = |ovr: &Option<CellBorderOverride>| -> Pt {
        match ovr {
            Some(CellBorderOverride::Border(b)) => b.width,
            _ => Pt::ZERO,
        }
    };
    let border_inset_h = cell_borders.as_ref()
        .map(|cb| {
            let bl = (border_w(&cb.left) - cell_margins.left).max(Pt::ZERO);
            let br = (border_w(&cb.right) - cell_margins.right).max(Pt::ZERO);
            bl + br
        })
        .unwrap_or(Pt::ZERO);
    let content_width = (inner_width - border_inset_h).max(Pt::ZERO);

    // Recurse into cell content blocks.
    let cell_blocks = build_cell_blocks(&cell.content, table_style, cond, content_width, ctx);

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
                    p, ctx, &mut pending_dropcap,
                    table_style, Some(cond),
                ) {
                    // Split oversized text fragments for narrow cells.
                    let lb = if let LayoutBlock::Paragraph { fragments, style, page_break_before, footnotes, floating_images } = lb {
                        let fragments = split_oversized_fragments(fragments, inner_width, ctx);
                        LayoutBlock::Paragraph { fragments, style, page_break_before, footnotes, floating_images }
                    } else {
                        lb
                    };
                    blocks.push(lb);
                }
            }
            Block::Table(nested_t) => {
                let built = build_table(nested_t, inner_width, ctx);
                blocks.push(LayoutBlock::Table {
                    rows: built.rows,
                    col_widths: built.col_widths,
                    border_config: built.border_config,
                    indent: built.indent,
                    alignment: built.alignment,
                    float_info: built.float_info,
                });
            }
            _ => {}
        }
    }

    blocks
}

// ── Paragraph property resolution ───────────────────────────────────────────

/// Resolve a paragraph's effective defaults.
/// Cascade: direct → style → doc defaults.
///
/// Returns (font_family, font_size, color, merged_paragraph_props, run_defaults).
///
/// When `defer_doc_defaults` is true, doc defaults are NOT merged into the
/// paragraph properties — the caller is responsible for merging them after
/// inserting table style / conditional formatting in the cascade.
fn resolve_paragraph_defaults(
    para: &Paragraph,
    resolved: &ResolvedDocument,
    defer_doc_defaults: bool,
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

    // Merge doc defaults as lowest-priority fallback (unless deferred for table cascade).
    if !defer_doc_defaults {
        merge_paragraph_properties(&mut para_props, &resolved.doc_defaults_paragraph);
    }

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

    // §17.3.1.33: when autoSpacing is true, use 14pt instead of explicit value.
    let space_before = if props.spacing.and_then(|s| s.before_auto_spacing) == Some(true) {
        Pt::new(14.0)
    } else {
        props.spacing.and_then(|s| s.before).map(Pt::from).unwrap_or(Pt::ZERO)
    };
    let space_after = if props.spacing.and_then(|s| s.after_auto_spacing) == Some(true) {
        Pt::new(14.0)
    } else {
        props.spacing.and_then(|s| s.after).map(Pt::from).unwrap_or(Pt::ZERO)
    };

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

    // §17.3.1.38: convert tab stops to layout format.
    // Clear entries are directives consumed during style merging, not layout stops.
    let tabs: Vec<TabStopDef> = props.tabs.iter()
        .filter(|t| t.alignment != model::TabAlignment::Clear)
        .map(|t| TabStopDef {
            position: Pt::from(t.position),
            alignment: t.alignment,
            leader: t.leader,
        })
        .collect();

    ParagraphStyle {
        alignment: props.alignment.unwrap_or(model::Alignment::Start),
        space_before,
        space_after,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        tabs,
        drop_cap: None,
        borders: resolve_paragraph_borders(props),
        shading: props.shading.as_ref().map(|s| resolve_color(s.fill, ColorContext::Background)),
        keep_next: props.keep_next.unwrap_or(false),
        contextual_spacing: props.contextual_spacing.unwrap_or(false),
        style_id: None, // set by caller when available
        page_floats: Vec::new(),
        page_y: crate::dimension::Pt::ZERO,
        page_x: crate::dimension::Pt::ZERO,
        page_content_width: crate::dimension::Pt::ZERO,
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
/// §17.4.38: borders with `val="none"` or `val="nil"` are suppressed.
fn convert_table_border_config(b: &model::TableBorders) -> TableBorderConfig {
    let convert = |border: &Option<model::Border>| -> Option<TableBorderLine> {
        border.as_ref().and_then(|b| {
            if b.style == model::BorderStyle::None {
                None
            } else {
                Some(convert_model_border(b))
            }
        })
    };
    TableBorderConfig {
        top: convert(&b.top),
        bottom: convert(&b.bottom),
        left: convert(&b.left),
        right: convert(&b.right),
        inside_h: convert(&b.inside_h),
        inside_v: convert(&b.inside_v),
    }
}

/// Split text fragments wider than `max_width` into per-character fragments
/// with individually measured widths. Used in narrow table cells for
/// character-level line breaking.
fn split_oversized_fragments(
    fragments: Vec<Fragment>,
    max_width: Pt,
    ctx: &BuildContext,
) -> Vec<Fragment> {
    if max_width <= Pt::ZERO {
        return fragments;
    }
    let mut result = Vec::with_capacity(fragments.len());
    for frag in fragments {
        match &frag {
            Fragment::Text { text, width, font, .. }
                if *width > max_width && text.chars().count() > 1 =>
            {
                // Re-measure each character individually.
                for ch in text.chars() {
                    let ch_str = ch.to_string();
                    let (w, m) = ctx.measurer.measure(&ch_str, font);
                    if let Fragment::Text {
                        color, shading, border, hyperlink_url,
                        baseline_offset, ..
                    } = &frag
                    {
                        result.push(Fragment::Text {
                            text: ch_str,
                            font: font.clone(),
                            color: *color,
                            shading: *shading,
                            border: *border,
                            width: w,
                            trimmed_width: w,
                            metrics: m,
                            hyperlink_url: hyperlink_url.clone(),
                            baseline_offset: *baseline_offset,
                            text_offset: Pt::ZERO,
                        });
                    }
                }
            }
            _ => result.push(frag),
        }
    }
    result
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

/// Remap PUA codepoints (0xF0xx) from legacy Symbol/Wingdings encoding
/// to standard Unicode, and return a portable font to render them with.
///
/// OOXML stores Symbol/Wingdings characters as PUA codepoints. These are
/// not portable across platforms — different OS font versions have different
/// cmap coverage. The standard approach (used by LibreOffice, Google Docs)
/// is to remap to Unicode equivalents from the official mapping tables:
/// - Symbol: unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
/// - Wingdings: standard Microsoft Wingdings-to-Unicode mapping
///
/// Returns (remapped_text, font_family).
fn remap_legacy_font_chars(text: &str, font_family: &str, fallback_family: &str) -> (String, String) {
    let is_symbol = font_family.eq_ignore_ascii_case("Symbol");
    let is_wingdings = font_family.eq_ignore_ascii_case("Wingdings");

    if !is_symbol && !is_wingdings {
        let family = if font_family.is_empty() { fallback_family } else { font_family };
        return (text.to_string(), family.to_string());
    }

    let remapped: String = text.chars().map(|ch| {
        let code = ch as u32;
        if is_symbol && (0xF020..=0xF0FF).contains(&code) {
            // Symbol font PUA mapping per unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
            match code {
                0xF020 => '\u{0020}', // SPACE
                0xF021 => '\u{0021}', // EXCLAMATION MARK
                0xF025 => '\u{0025}', // PERCENT SIGN
                0xF028 => '\u{0028}', // LEFT PARENTHESIS
                0xF029 => '\u{0029}', // RIGHT PARENTHESIS
                0xF02B => '\u{002B}', // PLUS SIGN
                0xF02E => '\u{002E}', // FULL STOP
                0xF030..=0xF039 => char::from_u32(code - 0xF000).unwrap_or(ch), // DIGITS
                0xF03C => '\u{003C}', // LESS-THAN SIGN
                0xF03D => '\u{003D}', // EQUALS SIGN
                0xF03E => '\u{003E}', // GREATER-THAN SIGN
                0xF05B => '\u{005B}', // LEFT SQUARE BRACKET
                0xF05D => '\u{005D}', // RIGHT SQUARE BRACKET
                0xF07B => '\u{007B}', // LEFT CURLY BRACKET
                0xF07C => '\u{007C}', // VERTICAL LINE
                0xF07D => '\u{007D}', // RIGHT CURLY BRACKET
                0xF07E => '\u{223C}', // TILDE OPERATOR
                0xF0A0 => '\u{20AC}', // EURO SIGN
                0xF0A5 => '\u{221E}', // INFINITY
                0xF0A7 => '\u{2663}', // BLACK CLUB SUIT
                0xF0A8 => '\u{2666}', // BLACK DIAMOND SUIT
                0xF0A9 => '\u{2665}', // BLACK HEART SUIT
                0xF0AA => '\u{2660}', // BLACK SPADE SUIT
                0xF0AB => '\u{2194}', // LEFT RIGHT ARROW
                0xF0AC => '\u{2190}', // LEFTWARDS ARROW
                0xF0AD => '\u{2191}', // UPWARDS ARROW
                0xF0AE => '\u{2192}', // RIGHTWARDS ARROW
                0xF0AF => '\u{2193}', // DOWNWARDS ARROW
                0xF0B0 => '\u{00B0}', // DEGREE SIGN
                0xF0B1 => '\u{00B1}', // PLUS-MINUS SIGN
                0xF0B2 => '\u{2033}', // DOUBLE PRIME
                0xF0B3 => '\u{2265}', // GREATER-THAN OR EQUAL TO
                0xF0B4 => '\u{00D7}', // MULTIPLICATION SIGN
                0xF0B5 => '\u{221D}', // PROPORTIONAL TO
                0xF0B7 => '\u{2022}', // BULLET
                0xF0B8 => '\u{00F7}', // DIVISION SIGN
                0xF0B9 => '\u{2260}', // NOT EQUAL TO
                0xF0BA => '\u{2261}', // IDENTICAL TO
                0xF0BB => '\u{2248}', // ALMOST EQUAL TO
                0xF0BC => '\u{2026}', // HORIZONTAL ELLIPSIS
                0xF0C0 => '\u{2135}', // ALEF SYMBOL
                0xF0C1 => '\u{2111}', // BLACK-LETTER CAPITAL I
                0xF0C2 => '\u{211C}', // BLACK-LETTER CAPITAL R
                0xF0C3 => '\u{2118}', // SCRIPT CAPITAL P
                0xF0C5 => '\u{2297}', // CIRCLED TIMES
                0xF0C6 => '\u{2295}', // CIRCLED PLUS
                0xF0C7 => '\u{2205}', // EMPTY SET
                0xF0C8 => '\u{2229}', // INTERSECTION
                0xF0C9 => '\u{222A}', // UNION
                0xF0CB => '\u{2283}', // SUPERSET OF
                0xF0CC => '\u{2287}', // SUPERSET OF OR EQUAL TO
                0xF0CD => '\u{2284}', // NOT A SUBSET OF
                0xF0CE => '\u{2282}', // SUBSET OF
                0xF0CF => '\u{2286}', // SUBSET OF OR EQUAL TO
                0xF0D0 => '\u{2208}', // ELEMENT OF
                0xF0D1 => '\u{2209}', // NOT AN ELEMENT OF
                0xF0D5 => '\u{220F}', // N-ARY PRODUCT
                0xF0D6 => '\u{221A}', // SQUARE ROOT
                0xF0D7 => '\u{22C5}', // DOT OPERATOR
                0xF0D8 => '\u{00AC}', // NOT SIGN
                0xF0D9 => '\u{2227}', // LOGICAL AND
                0xF0DA => '\u{2228}', // LOGICAL OR
                0xF0E0 => '\u{21D0}', // LEFTWARDS DOUBLE ARROW
                0xF0E1 => '\u{21D1}', // UPWARDS DOUBLE ARROW
                0xF0E2 => '\u{21D2}', // RIGHTWARDS DOUBLE ARROW
                0xF0E3 => '\u{21D3}', // DOWNWARDS DOUBLE ARROW
                0xF0E4 => '\u{21D4}', // LEFT RIGHT DOUBLE ARROW
                0xF0E5 => '\u{2329}', // LEFT-POINTING ANGLE BRACKET
                0xF0F1 => '\u{232A}', // RIGHT-POINTING ANGLE BRACKET
                0xF0F2 => '\u{222B}', // INTEGRAL
                _ => ch,
            }
        } else if is_wingdings && (0xF020..=0xF0FF).contains(&code) {
            // Wingdings PUA mapping per Microsoft Wingdings-to-Unicode table
            match code {
                0xF021 => '\u{270E}', // LOWER RIGHT PENCIL
                0xF022 => '\u{2702}', // BLACK SCISSORS
                0xF023 => '\u{2701}', // UPPER BLADE SCISSORS
                0xF028 => '\u{1F4CB}',// CLIPBOARD (may need supplementary)
                0xF029 => '\u{1F4CB}',// CLIPBOARD
                0xF041 => '\u{FE4E}', // WAVY LOW LINE (approximate)
                0xF046 => '\u{1F44D}',// THUMBS UP SIGN
                0xF04C => '\u{2639}', // WHITE FROWNING FACE
                0xF04A => '\u{263A}', // WHITE SMILING FACE
                0xF06C => '\u{25CF}', // BLACK CIRCLE
                0xF06D => '\u{274D}', // SHADOWED WHITE CIRCLE
                0xF06E => '\u{25A0}', // BLACK SQUARE
                0xF06F => '\u{25A1}', // WHITE SQUARE
                0xF070 => '\u{25A1}', // WHITE SQUARE (alt)
                0xF071 => '\u{2751}', // LOWER RIGHT SHADOWED WHITE SQUARE
                0xF072 => '\u{2752}', // UPPER RIGHT SHADOWED WHITE SQUARE
                0xF073 => '\u{25C6}', // BLACK DIAMOND
                0xF074 => '\u{2756}', // BLACK DIAMOND MINUS WHITE X
                0xF076 => '\u{2756}', // BLACK DIAMOND MINUS WHITE X
                0xF09F => '\u{2708}', // AIRPLANE
                0xF0A1 => '\u{270C}', // VICTORY HAND
                0xF0A4 => '\u{261C}', // WHITE LEFT POINTING INDEX
                0xF0A5 => '\u{261E}', // WHITE RIGHT POINTING INDEX
                0xF0A7 => '\u{25AA}', // BLACK SMALL SQUARE
                0xF0A8 => '\u{25FB}', // WHITE MEDIUM SQUARE
                0xF0D5 => '\u{232B}', // ERASE TO THE LEFT
                0xF0D8 => '\u{27A2}', // THREE-D TOP-LIGHTED RIGHTWARDS ARROWHEAD
                0xF0E8 => '\u{2B22}', // BLACK HEXAGON (approximate)
                0xF0F0 => '\u{2B1A}', // DOTTED SQUARE (approximate)
                0xF0FC => '\u{2714}', // HEAVY CHECK MARK
                0xF0FB => '\u{2718}', // HEAVY BALLOT X
                0xF0FE => '\u{2612}', // BALLOT BOX WITH X (approximate)
                _ => ch,
            }
        } else {
            ch
        }
    }).collect();

    // After PUA→Unicode remapping, the original Symbol/Wingdings font
    // cannot render the standard Unicode codepoints (legacy fonts lack
    // Unicode cmaps). Use the document's fallback font for the remapped
    // glyphs — standard Unicode bullets (U+2022, U+25AA) are in most
    // text fonts.
    (remapped, fallback_family.to_string())
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

/// Extract the display size for a picture bullet from its VML shape style.
/// Falls back to 9pt × 9pt (common Word default for picture bullets).
fn pic_bullet_size(bullet: &model::NumPicBullet) -> PtSize {
    use dxpdf_docx_model::model::VmlLengthUnit;

    let default = PtSize::new(Pt::new(9.0), Pt::new(9.0));
    let shape = match bullet.pict.as_ref().and_then(|p| p.shapes.first()) {
        Some(s) => s,
        None => return default,
    };

    let to_pt = |len: &dxpdf_docx_model::model::VmlLength| -> Pt {
        let val = len.value as f32;
        match len.unit {
            VmlLengthUnit::Pt => Pt::new(val),
            VmlLengthUnit::In => Pt::new(val * 72.0),
            VmlLengthUnit::Cm => Pt::new(val * 28.3465),
            VmlLengthUnit::Mm => Pt::new(val * 2.83465),
            VmlLengthUnit::Px => Pt::new(val * 0.75),
            _ => Pt::new(val),
        }
    };

    let w = shape.style.width.as_ref().map(to_pt).unwrap_or(default.width);
    let h = shape.style.height.as_ref().map(to_pt).unwrap_or(default.height);
    PtSize::new(w, h)
}
