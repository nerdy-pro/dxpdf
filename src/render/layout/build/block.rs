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
    paragraph_style_from_props, pic_bullet_size, populate_image_data, populate_underline_metrics,
    remap_legacy_font_chars, resolve_paragraph_defaults,
};
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
    if let Some(ref num_ref) = merged_props.numbering {
        let num_id = model::NumId::new(num_ref.num_id);
        let level = num_ref.level;
        if let Some(levels) = ctx.resolved.numbering.get(&num_id) {
            // Update counters: increment this level, reset deeper levels.
            {
                let counters = &mut state.list_counters;
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
                    let rel_id = bullet
                        .pict
                        .as_ref()?
                        .shapes
                        .first()?
                        .image_data
                        .as_ref()?
                        .rel_id
                        .as_ref()?;
                    let image_bytes = ctx.media().get(rel_id)?;
                    // Size from VML shape style (width/height), default 9pt.
                    let size = pic_bullet_size(bullet);
                    let label_frag = Fragment::Image {
                        size,
                        rel_id: rel_id.as_str().to_string(),
                        image_data: Some(image_bytes.clone()),
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
                    merged_props.tabs.insert(
                        0,
                        crate::model::TabStop {
                            position: lvl_left,
                            alignment: crate::model::TabAlignment::Left,
                            leader: crate::model::TabLeader::None,
                        },
                    );
                }
            } else {
                let counters = &state.list_counters;
                if let Some(label_text) = crate::render::resolve::numbering::format_list_label(
                    levels, level, counters, num_id,
                ) {
                    // Resolve label font from level run_properties or paragraph defaults.
                    let (default_family, default_size, default_color, _, _) =
                        resolve_paragraph_defaults(p, ctx.resolved, false);
                    let level_font_family = level_def
                        .and_then(|l| l.run_properties.as_ref())
                        .and_then(|rp| crate::render::resolve::fonts::effective_font(&rp.fonts))
                        .unwrap_or("");

                    // Remap PUA codepoints from legacy Symbol/Wingdings encoding
                    // to standard Unicode per official mapping tables, and use
                    // a standard font. This is portable across all platforms.
                    let (label_text, label_family) =
                        remap_legacy_font_chars(&label_text, level_font_family, &default_family);
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
                        Some(crate::model::Alignment::End) => -w,
                        Some(crate::model::Alignment::Center) => w * -0.5,
                        _ => Pt::ZERO,
                    };
                    // The label fragment width is the text width only.
                    // text_offset shifts where the text is drawn (for right/
                    // center justification) but x advances by w, leaving
                    // room for the tab to fill hanging − w to the stop.
                    let label_width = w;
                    let label_frag = Fragment::Text {
                        text: Rc::from(label_text.as_str()),
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
                        merged_props.tabs.insert(
                            0,
                            crate::model::TabStop {
                                position: lvl_left,
                                alignment: crate::model::TabAlignment::Left,
                                leader: crate::model::TabLeader::None,
                            },
                        );
                    }
                }
            }

            // §17.9.23: numbering level pPr overrides the paragraph style.
            // Only the paragraph's direct ind overrides the numbering level.
            if let Some(lvl_ind) = levels
                .get(level as usize)
                .and_then(|l| l.indentation.as_ref())
            {
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

/// Extract floating (anchor) images from a paragraph's inlines.
/// When `cell_context` is true, positions are resolved relative to the cell
/// origin (0,0) instead of the page margins.
pub(super) fn extract_floating_images(
    para: &Paragraph,
    ctx: &BuildContext,
    state: &BuildState,
    cell_context: bool,
) -> Vec<crate::render::layout::section::FloatingImage> {
    use crate::model::{
        AnchorAlignment, AnchorPosition, AnchorRelativeFrom, ImagePlacement, Inline,
    };
    use crate::render::layout::section::{FloatingImage, FloatingImageY};

    let mut images = Vec::new();

    fn find_anchor_images<'a>(inlines: &'a [Inline], out: &mut Vec<&'a crate::model::Image>) {
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
            let rel_id = match crate::render::resolve::images::extract_image_rel_id(img) {
                Some(id) => id,
                None => {
                    eprintln!(
                        "  -> no rel_id, graphic.is_some()={}",
                        img.graphic.is_some()
                    );
                    continue;
                }
            };

            let image_data = match ctx.resolved.media.get(rel_id) {
                Some(entry) => entry.clone(),
                None => {
                    eprintln!(
                        "Anchor image: rel_id={} NOT FOUND in media (media has {} entries)",
                        rel_id.as_str(),
                        ctx.resolved.media.len()
                    );
                    continue;
                }
            };

            let w = Pt::from(img.extent.width);
            let h = Pt::from(img.extent.height);
            let pc = &state.page_config;

            // Resolve horizontal position.
            // In cell context, positions are relative to the cell origin.
            let (page_width, margin_left, margin_right) = if cell_context {
                (Pt::ZERO, Pt::ZERO, Pt::ZERO)
            } else {
                (pc.page_size.width, pc.margins.left, pc.margins.right)
            };
            let content_width = if cell_context {
                Pt::ZERO
            } else {
                page_width - margin_left - margin_right
            };

            let x = match &anchor.horizontal_position {
                AnchorPosition::Offset {
                    relative_from,
                    offset,
                } => {
                    let base = match relative_from {
                        AnchorRelativeFrom::Page => Pt::ZERO,
                        AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => margin_left,
                        _ => margin_left,
                    };
                    base + Pt::from(*offset)
                }
                AnchorPosition::Align {
                    relative_from,
                    alignment,
                } => {
                    let (area_left, area_width) = match relative_from {
                        AnchorRelativeFrom::Page => (Pt::ZERO, page_width),
                        AnchorRelativeFrom::Margin | AnchorRelativeFrom::Column => {
                            (margin_left, content_width)
                        }
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
                AnchorPosition::Offset {
                    relative_from,
                    offset,
                } => {
                    let margin_top = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.margins.top
                    };
                    if cell_context {
                        // In cell context, all positions are relative to cell origin.
                        FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                    } else {
                        match relative_from {
                            AnchorRelativeFrom::Page => FloatingImageY::Absolute(Pt::from(*offset)),
                            AnchorRelativeFrom::Margin => {
                                FloatingImageY::Absolute(margin_top + Pt::from(*offset))
                            }
                            // §20.4.2.11: topMargin — offset from page top.
                            AnchorRelativeFrom::TopMargin => {
                                FloatingImageY::Absolute(Pt::from(*offset))
                            }
                            // §20.4.2.11: bottomMargin — offset from bottom margin edge.
                            AnchorRelativeFrom::BottomMargin => {
                                let page_height = pc.page_size.height;
                                let margin_bottom = pc.margins.bottom;
                                FloatingImageY::Absolute(
                                    page_height - margin_bottom + Pt::from(*offset),
                                )
                            }
                            AnchorRelativeFrom::Paragraph | AnchorRelativeFrom::Line => {
                                FloatingImageY::RelativeToParagraph(Pt::from(*offset))
                            }
                            _ => FloatingImageY::Absolute(margin_top + Pt::from(*offset)),
                        }
                    }
                }
                AnchorPosition::Align {
                    relative_from,
                    alignment,
                } => {
                    let margin_top = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.margins.top
                    };
                    let page_height = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.page_size.height
                    };
                    let margin_bottom = if cell_context {
                        Pt::ZERO
                    } else {
                        pc.margins.bottom
                    };
                    let (area_top, area_height) = match relative_from {
                        AnchorRelativeFrom::Page => (Pt::ZERO, page_height),
                        AnchorRelativeFrom::Margin => {
                            (margin_top, page_height - margin_top - margin_bottom)
                        }
                        // §20.4.2.11: topMargin = area from page top to top margin edge.
                        AnchorRelativeFrom::TopMargin => (Pt::ZERO, margin_top),
                        // §20.4.2.11: bottomMargin = area from bottom margin edge to page bottom.
                        AnchorRelativeFrom::BottomMargin => {
                            (page_height - margin_bottom, margin_bottom)
                        }
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
                size: crate::render::geometry::PtSize::new(w, h),
                x,
                y,
                wrap_top_and_bottom: matches!(
                    anchor.wrap,
                    crate::model::TextWrap::TopAndBottom { .. }
                ),
                dist_left: Pt::from(anchor.distance.left),
                dist_right: Pt::from(anchor.distance.right),
                behind_doc: anchor.behind_text,
            });
        }
    }

    images
}

/// Search an inline (and AlternateContent fallback) for a VML text box with
/// absolute positioning.
pub(super) fn find_vml_absolute_position(inline: &model::Inline) -> Option<(Pt, Pt)> {
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
    use crate::model::CssPosition;
    if style.position != Some(CssPosition::Absolute) {
        return None;
    }
    let x = style.margin_left.map(vml_length_to_pt)?;
    let y = style.margin_top.map(vml_length_to_pt)?;
    Some((x, y))
}

/// Convert a VML CSS length to points.
fn vml_length_to_pt(len: model::VmlLength) -> Pt {
    use crate::model::VmlLengthUnit;
    let value = len.value as f32;
    Pt::new(match len.unit {
        VmlLengthUnit::Pt => value,
        VmlLengthUnit::In => value * 72.0,
        VmlLengthUnit::Cm => value * 72.0 / 2.54,
        VmlLengthUnit::Mm => value * 72.0 / 25.4,
        VmlLengthUnit::Px => value * 0.75, // 96dpi → 72pt/in
        VmlLengthUnit::None => value / 914400.0 * 72.0, // bare number = EMU
        _ => value,                        // Em, Percent — fallback to raw value
    })
}
