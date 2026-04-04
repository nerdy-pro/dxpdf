use std::rc::Rc;

use crate::model::{self, Block, Paragraph};
use crate::render::dimension::Pt;
use crate::render::layout::fragment::{collect_fragments, FontProps, Fragment, FragmentCtx};
use crate::render::layout::paragraph::DropCapInfo;
use crate::render::layout::section::LayoutBlock;
use crate::render::resolve::color::{resolve_color, ColorContext, RgbColor};
use crate::render::resolve::conditional::CellConditionalFormatting;
use crate::render::resolve::properties::{merge_paragraph_properties, merge_run_properties};
use crate::render::resolve::styles::ResolvedStyle;

use super::convert::{
    paragraph_style_from_props, populate_image_data, populate_underline_metrics,
    resolve_paragraph_defaults,
};
use super::floating::extract_floating_images;
use super::table::build_table;
use super::{BuildContext, BuildState};

/// Recursively process a single model block into a layout block.
///
/// Returns `None` for drop cap paragraphs (consumed by the next paragraph)
/// and section breaks (already handled by resolve).
pub(super) fn build_block(
    block: &Block,
    available_width: Pt,
    ctx: &BuildContext,
    state: &mut BuildState,
    pending_dropcap: &mut Option<DropCapInfo>,
) -> Option<LayoutBlock> {
    match block {
        Block::Paragraph(p) => build_paragraph_block(p, ctx, state, pending_dropcap, None, None),
        Block::Table(t) => {
            let built = build_table(t, available_width, ctx, state);
            Some(LayoutBlock::Table {
                rows: built.rows,
                col_widths: built.col_widths,
                border_config: built.border_config,
                indent: built.indent,
                alignment: built.alignment,
                float_info: built.float_info,
                style_id: t.properties.style_id.clone(),
            })
        }
        Block::SectionBreak(_) => None,
    }
}

// ── Paragraph building ──────────────────────────────────────────────────────

/// Build a paragraph into a layout block.
/// Handles drop cap detection (§17.3.1.11), list labels, floating images.
/// For table cells, pass `table_style` and `cond` to apply table formatting cascade.
pub(super) fn build_paragraph_block(
    p: &Paragraph,
    ctx: &BuildContext,
    state: &mut BuildState,
    pending_dropcap: &mut Option<DropCapInfo>,
    table_style: Option<&ResolvedStyle>,
    cond: Option<&CellConditionalFormatting>,
) -> Option<LayoutBlock> {
    let (mut fragments, mut merged_props) = build_fragments(p, ctx, state, table_style, cond);

    // §17.9.22: inject list label if paragraph has a numbering reference.
    super::list_label::inject_list_label(p, &mut fragments, &mut merged_props, ctx, state);

    // Word suppresses Hyperlink character style (blue/underline) for ToC
    // entries in print view. Strip visual hyperlink styling but keep the
    // click annotation URL.
    if p.style_id
        .as_ref()
        .is_some_and(|id| id.as_str().starts_with("TOC") || id.as_str().starts_with("toc"))
    {
        for frag in &mut fragments {
            if let Fragment::Text {
                font,
                color,
                hyperlink_url,
                ..
            } = frag
            {
                if hyperlink_url.is_some() {
                    *color = RgbColor::BLACK;
                    font.underline = false;
                }
            }
        }
    }

    // §17.3.1.11: detect drop cap paragraph.
    if let Some(model::FrameKind::DropCap {
        style,
        lines,
        h_space: dc_h_space,
    }) = merged_props.frame_properties
    {
        let drop_cap_lines = lines;
        let width: Pt = fragments.iter().map(|f| f.width()).sum();
        let height: Pt = fragments.iter().map(|f| f.height()).fold(Pt::ZERO, Pt::max);
        let ascent: Pt = fragments
            .iter()
            .map(|f| match f {
                Fragment::Text { metrics, .. } => metrics.ascent,
                _ => Pt::ZERO,
            })
            .fold(Pt::ZERO, Pt::max);
        let h_space = dc_h_space.map(Pt::from).unwrap_or(Pt::ZERO);
        let margin_mode = matches!(style, model::DropCap::Margin);
        // The drop cap paragraph's own indent determines the x position.
        // This includes indent_left + indent_first_line from the cascade.
        let dc_indent_left = merged_props
            .indentation
            .and_then(|i| i.start)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO);
        let dc_indent_first = merged_props
            .indentation
            .and_then(|i| i.first_line)
            .map(|fl| match fl {
                model::FirstLineIndent::FirstLine(v) => Pt::from(v),
                model::FirstLineIndent::Hanging(v) => -Pt::from(v),
                model::FirstLineIndent::None => Pt::ZERO,
            })
            .unwrap_or(Pt::ZERO);
        // §17.3.1.33: frame height from drop cap paragraph's exact line spacing.
        let frame_height = merged_props
            .spacing
            .and_then(|s| s.line)
            .and_then(|ls| match ls {
                model::LineSpacing::Exact(v) => Some(Pt::from(v)),
                _ => None,
            });
        // §17.3.2.19: position offset from the drop cap run.
        let position_offset = fragments
            .first()
            .and_then(|f| match f {
                Fragment::Text {
                    baseline_offset, ..
                } => Some(*baseline_offset),
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
    let fn_refs: Vec<_> = p
        .content
        .iter()
        .filter_map(|i| {
            if let model::Inline::FootnoteRef(id) = i {
                Some(id)
            } else {
                None
            }
        })
        .collect();
    let fn_base = state.footnote_counter - fn_refs.len() as u32;
    let mut para_footnotes = Vec::new();
    for (i, note_id) in fn_refs.iter().enumerate() {
        let display = format!("{}", fn_base + i as u32 + 1);
        if let Some(content) = ctx.resolved.footnotes.get(note_id) {
            let notes = build_note_content(note_id.value(), &display, content, ctx, state);
            for (_, frags, style) in notes {
                para_footnotes.push((frags, style));
            }
        }
    }

    // §20.4.2.3: extract floating (anchor) images from this paragraph.
    // In cell context, positions are cell-relative instead of page-relative.
    let cell_context = table_style.is_some();
    let floating_images = extract_floating_images(p, ctx, state, cell_context);

    Some(LayoutBlock::Paragraph {
        fragments,
        style,
        page_break_before,
        footnotes: para_footnotes,
        floating_images,
    })
}

/// Build note content (footnotes or endnotes) with a display number prefix.
pub(super) fn build_note_content(
    _note_id_value: i64,
    display_num: &str,
    content: &[Block],
    ctx: &BuildContext,
    state: &mut BuildState,
) -> Vec<(
    String,
    Vec<Fragment>,
    crate::render::layout::paragraph::ParagraphStyle,
)> {
    let mut results = Vec::new();
    for (i, block) in content.iter().enumerate() {
        if let model::Block::Paragraph(p) = block {
            let (mut frags, merged_props) = build_fragments(p, ctx, state, None, None);

            // Prepend display number to the first paragraph.
            if i == 0 && !frags.is_empty() {
                let num_text = format!("{}  ", display_num);
                let font = frags[0].font_props().cloned().unwrap_or_else(|| FontProps {
                    family: std::rc::Rc::from("Times New Roman"),
                    size: Pt::new(10.0),
                    bold: false,
                    italic: false,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                });
                let ref_size = font.size * 0.58;
                let ref_font = FontProps {
                    size: ref_size,
                    ..font
                };
                let (w, m) = ctx.measurer.measure(&num_text, &ref_font);
                frags.insert(
                    0,
                    Fragment::Text {
                        text: Rc::from(num_text.as_str()),
                        font: ref_font,
                        color: RgbColor::BLACK,
                        shading: None,
                        border: None,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: None,
                        baseline_offset: -(font.size * 0.4),
                        text_offset: Pt::ZERO,
                    },
                );
            }
            let style = paragraph_style_from_props(&merged_props);
            results.push((display_num.to_string(), frags, style));
        }
    }
    results
}

/// Collect endnotes from the resolved document.
pub(super) fn collect_endnotes(
    ctx: &BuildContext,
    state: &mut BuildState,
    endnotes: &mut Vec<(
        String,
        Vec<Fragment>,
        crate::render::layout::paragraph::ParagraphStyle,
    )>,
) {
    // IDs 0 and 1 are reserved for separator and continuation separator.
    let mut en_ids: Vec<_> = ctx
        .resolved
        .endnotes
        .keys()
        .filter(|id| id.value() > 1)
        .collect();
    en_ids.sort_by_key(|id| id.value());
    for (i, note_id) in en_ids.iter().enumerate() {
        let display = crate::render::layout::fragment::to_roman_lower((i + 1) as u32);
        if let Some(content) = ctx.resolved.endnotes.get(note_id) {
            endnotes.extend(build_note_content(
                note_id.value(),
                &display,
                content,
                ctx,
                state,
            ));
        }
    }
}

/// Build fragments and resolved paragraph properties for a paragraph.
///
/// Handles the full cascade: table style → conditional → paragraph style →
/// doc defaults → fragment collection → image/underline population.
pub(super) fn build_fragments(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &mut BuildState,
    table_style: Option<&ResolvedStyle>,
    cond: Option<&CellConditionalFormatting>,
) -> (Vec<Fragment>, model::ParagraphProperties) {
    // §17.7.2: resolve paragraph defaults (direct → paragraph style).
    // Doc defaults are deferred so table style/conditional can be inserted
    // between paragraph style and doc defaults in the cascade.
    let (default_family, mut default_size, mut default_color, mut merged_props, mut run_defaults) =
        resolve_paragraph_defaults(para, ctx.resolved, table_style.is_some());

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

    let measure =
        |text: &str, font: &FontProps| -> (Pt, crate::render::layout::fragment::TextMetrics) {
            ctx.measurer.measure(text, font)
        };

    let frag_ctx = FragmentCtx {
        default_family: &default_family,
        default_size,
        default_color,
        resolved_styles: Some(&ctx.resolved.styles),
        paragraph_run_defaults: Some(&run_defaults),
        theme: ctx.resolved.theme.as_ref(),
    };
    let mut fragments = collect_fragments(
        &para.content,
        &frag_ctx,
        None,
        &measure,
        &mut state.footnote_counter,
        &mut state.endnote_counter,
        state.field_ctx,
    );
    populate_image_data(&mut fragments, ctx.media());
    populate_underline_metrics(&mut fragments, ctx.measurer);

    (fragments, merged_props)
}
