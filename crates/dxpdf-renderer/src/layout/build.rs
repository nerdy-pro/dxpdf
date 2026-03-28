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
use crate::layout::paragraph::TabStopDef;
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

/// Context threaded through the recursive tree walk.
pub struct BuildContext<'a> {
    pub measurer: &'a TextMeasurer,
    pub resolved: &'a ResolvedDocument,
    /// Sequential footnote display number (1, 2, 3...).
    pub footnote_counter: std::cell::Cell<u32>,
    /// Sequential endnote display number (i, ii, iii...).
    pub endnote_counter: std::cell::Cell<u32>,
    /// Per-(numId, level) running counters for list labels.
    pub list_counters: std::cell::RefCell<HashMap<(model::NumId, u8), u32>>,
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
                let (w, h, a) = ctx.measurer.measure(&num_text, &ref_font);
                frags.insert(0, Fragment::Text {
                    text: num_text,
                    font: ref_font,
                    color: RgbColor::BLACK,
                    shading: None, border: None,
                    width: w, trimmed_width: w,
                    height: h, ascent: a,
                    hyperlink_url: None,
                    baseline_offset: -(font.size * 0.4),
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
    let mut en_ids: Vec<_> = ctx.resolved.endnotes.keys()
        .filter(|id| id.value() > 0)
        .collect();
    en_ids.sort_by_key(|id| id.value());
    for (i, note_id) in en_ids.iter().enumerate() {
        let display = crate::layout::fragment::to_roman_lower((i + 1) as u32);
        if let Some(content) = ctx.resolved.endnotes.get(note_id) {
            endnotes.extend(build_note_content(note_id.value(), &display, content, ctx));
        }
    }
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
                indent: built.indent,
                alignment: built.alignment,
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
    let (mut fragments, mut merged_props) = build_fragments(p, ctx, None, None);

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

            let counters = ctx.list_counters.borrow();
            if let Some(label_text) = crate::resolve::numbering::format_list_label(
                levels, level, &counters, num_id,
            ) {
                // Resolve label font from level run_properties or paragraph defaults.
                let (default_family, default_size, default_color, _, _) =
                    resolve_paragraph_defaults(p, ctx.resolved);
                let level_def = levels.get(level as usize);
                let level_font_family = level_def
                    .and_then(|l| l.run_properties.as_ref())
                    .and_then(|rp| crate::resolve::fonts::effective_font(&rp.fonts))
                    .unwrap_or("");

                // Remap PUA codepoints from legacy Symbol/Wingdings encoding to
                // Unicode equivalents so standard fonts can render them.
                // Symbol mapping: unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
                // Wingdings mapping: standard Microsoft Wingdings-to-Unicode table.
                let (label_text, label_family) = remap_symbol_bullets(
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
                let (w, h, a) = ctx.measurer.measure(&label_text, &label_font);
                let label_frag = Fragment::Text {
                    text: label_text,
                    font: label_font,
                    color: default_color,
                    shading: None,
                    border: None,
                    width: w,
                    trimmed_width: w,
                    height: h,
                    ascent: a,
                    hyperlink_url: None,
                    baseline_offset: Pt::ZERO,
                };
                // Tab after label: advances to indent_left via the implicit
                // tab stop. Fitting width = hanging - label_width so the
                // fitter and renderer agree on consumed space.
                let hanging = levels.get(level as usize)
                    .and_then(|l| l.indentation.as_ref())
                    .and_then(|ind| ind.first_line)
                    .map(|fl| match fl {
                        model::FirstLineIndent::Hanging(v) => Pt::from(v),
                        _ => Pt::ZERO,
                    })
                    .unwrap_or(Pt::ZERO);
                let tab_fitting = (hanging - w).max(Pt::ZERO);
                let tab_frag = Fragment::Tab {
                    line_height: h,
                    fitting_width: Some(tab_fitting),
                };
                fragments.insert(0, tab_frag);
                fragments.insert(0, label_frag);

                // Add implicit tab stop at numLvl.left so the tab lands
                // at the body text position.
                if let Some(lvl_left) = levels.get(level as usize)
                    .and_then(|l| l.indentation.as_ref())
                    .and_then(|ind| ind.start)
                {
                    merged_props.tabs.insert(0, dxpdf_docx_model::model::TabStop {
                        position: lvl_left,
                        alignment: dxpdf_docx_model::model::TabAlignment::Left,
                        leader: dxpdf_docx_model::model::TabLeader::None,
                    });
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
                Fragment::Text { ascent, .. } => *ascent,
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
    let floating_images = extract_floating_images(p, ctx);

    Some(LayoutBlock::Paragraph {
        fragments,
        style,
        page_break_before,
        footnotes: para_footnotes,
        floating_images,
    })
}

/// Extract floating (anchor) images from a paragraph's inlines.
fn extract_floating_images(
    para: &Paragraph,
    ctx: &BuildContext,
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

                // Resolve horizontal position.
                // TODO: get actual page config here. For now use US Letter defaults.
                let page_width = Pt::new(612.0);
                let margin_left = Pt::new(72.0);
                let margin_right = Pt::new(72.0);
                let content_width = page_width - margin_left - margin_right;

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
                        let margin_top = Pt::new(72.0);
                        match relative_from {
                            AnchorRelativeFrom::Page => FloatingImageY::Absolute(Pt::from(*offset)),
                            AnchorRelativeFrom::Margin => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                            AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                                FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                            }
                            _ => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                        }
                    }
                    AnchorPosition::Align { relative_from, alignment } => {
                        let margin_top = Pt::new(72.0);
                        let page_height = Pt::new(792.0);
                        let margin_bottom = Pt::new(72.0);
                        let (area_top, area_height) = match relative_from {
                            AnchorRelativeFrom::Page => (Pt::ZERO, page_height),
                            AnchorRelativeFrom::Margin => (margin_top, page_height - margin_top - margin_bottom),
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

    let measure = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
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
) -> Vec<CellBlock> {
    let dlh = default_line_height(ctx);
    content
        .iter()
        .enumerate()
        .filter_map(|(i, block)| match block {
            Block::Paragraph(p) => {
                // §17.4.66: every cell must end with a paragraph. When the
                // last block is an empty paragraph following a table, it is
                // structural — Word renders it with zero height.
                if p.content.is_empty()
                    && i > 0
                    && matches!(content[i - 1], Block::Table(_))
                    && i == content.len() - 1
                {
                    return None;
                }
                let (frags, merged_props) = build_fragments(p, ctx, table_style, Some(cond));
                // Split oversized text fragments into per-character fragments
                // so narrow table cells get character-level line breaking.
                let frags = split_oversized_fragments(frags, inner_width, ctx);
                Some(CellBlock::Paragraph {
                    fragments: frags,
                    style: paragraph_style_from_props(&merged_props),
                })
            }
            Block::Table(nested_t) => {
                let built = build_table(nested_t, inner_width, ctx);
                let measure_fn = |text: &str, font: &FontProps| -> (Pt, Pt, Pt) {
                    ctx.measurer.measure(text, font)
                };
                let result = layout_table(
                    &built.rows,
                    &built.col_widths,
                    &crate::layout::BoxConstraints::unbounded(),
                    dlh,
                    built.border_config.as_ref(),
                    Some(&measure_fn),
                );
                // §17.4.28: apply nested table alignment within the cell.
                let align_offset = match built.alignment {
                    Some(model::Alignment::Center) => {
                        (inner_width - result.size.width) * 0.5
                    }
                    Some(model::Alignment::End) => {
                        inner_width - result.size.width
                    }
                    _ => built.indent,
                };
                let commands = if align_offset != Pt::ZERO {
                    result.commands.into_iter().map(|mut cmd| {
                        cmd.shift_x(align_offset);
                        cmd
                    }).collect()
                } else {
                    result.commands
                };
                Some(CellBlock::NestedTable {
                    commands,
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

    // §17.3.1.38: convert tab stops to layout format.
    let tabs: Vec<TabStopDef> = props.tabs.iter().map(|t| TabStopDef {
        position: Pt::from(t.position),
        alignment: t.alignment,
        leader: t.leader,
    }).collect();

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
                    let (w, h, a) = ctx.measurer.measure(&ch_str, font);
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
                            height: h,
                            ascent: a,
                            hyperlink_url: hyperlink_url.clone(),
                            baseline_offset: *baseline_offset,
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

/// Remap PUA codepoints from legacy Symbol/Wingdings encoding to Unicode.
///
/// Symbol font mapping per unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt.
/// Wingdings mapping per the standard Microsoft Wingdings-to-Unicode table.
///
/// Returns the remapped text and the font family to use for rendering.
fn remap_symbol_bullets(text: &str, font_family: &str, fallback_family: &str) -> (String, String) {
    let is_symbol = font_family.eq_ignore_ascii_case("Symbol");
    let is_wingdings = font_family.eq_ignore_ascii_case("Wingdings");

    if !is_symbol && !is_wingdings {
        // Not a legacy symbol font — use as-is.
        let family = if font_family.is_empty() { fallback_family } else { font_family };
        return (text.to_string(), family.to_string());
    }

    let remapped: String = text.chars().map(|ch| {
        let code = ch as u32;
        if is_symbol && (0xF020..=0xF0FF).contains(&code) {
            // Symbol font: offset 0xF000 from encoding position.
            match code {
                0xF0B7 => '\u{2022}', // BULLET
                0xF0A8 => '\u{2666}', // BLACK DIAMOND SUIT
                0xF0B0 => '\u{00B0}', // DEGREE SIGN
                0xF0D7 => '\u{00D7}', // MULTIPLICATION SIGN
                0xF0B1 => '\u{00B1}', // PLUS-MINUS SIGN
                _ => ch, // keep as-is for unmapped
            }
        } else if is_wingdings && (0xF020..=0xF0FF).contains(&code) {
            match code {
                0xF0A7 => '\u{25AA}', // BLACK SMALL SQUARE
                0xF0D8 => '\u{2794}', // HEAVY WIDE-HEADED RIGHTWARDS ARROW
                0xF0FC => '\u{2714}', // HEAVY CHECK MARK
                0xF076 => '\u{2756}', // BLACK DIAMOND MINUS WHITE X
                0xF06C => '\u{25CF}', // BLACK CIRCLE
                _ => ch,
            }
        } else {
            ch
        }
    }).collect();

    // Use a standard font for the Unicode equivalents.
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
