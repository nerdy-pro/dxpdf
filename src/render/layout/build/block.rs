use crate::model::Dup;
use std::rc::Rc;

use crate::i18n::bidi::BidiLevel;
use crate::model::{self, Block, Paragraph};
use crate::render::dimension::Pt;
use crate::render::layout::fragment::{
    collect_fragments, font_props_from_run, FontProps, Fragment, FragmentCtx, MarkLine,
};
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
use super::floating::{extract_floating_images, AnchorFrame};
use super::table::build_table;
use super::{BuildContext, BuildState};
use crate::render::fonts::Toggle;
use crate::render::layout::ShapeAutoFit;

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
                cell_spacing: built.cell_spacing,
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
/// Whether every visible inline of `p` is a deletion-stamped run — nothing
/// would survive the final view's suppression filter. Invisible markers
/// (bookmarks, comment anchors) don't count as survivors; anything else
/// visible does, conservatively.
fn paragraph_fully_deleted(p: &Paragraph) -> bool {
    fn deleted(inlines: &[model::Inline]) -> bool {
        inlines.iter().all(|i| match i {
            model::Inline::TextRun(tr) => tr
                .revision
                .as_ref()
                .is_some_and(|r| r.kind == model::RevisionKind::Deleted),
            model::Inline::Hyperlink(h) => deleted(&h.content),
            model::Inline::BookmarkStart { .. }
            | model::Inline::BookmarkEnd(_)
            | model::Inline::CommentRef(_) => true,
            _ => false,
        })
    }
    deleted(&p.content)
}

pub(super) fn build_paragraph_block(
    p: &Paragraph,
    ctx: &BuildContext,
    state: &mut BuildState,
    pending_dropcap: &mut Option<DropCapInfo>,
    table_style: Option<&ResolvedStyle>,
    cond: Option<&CellConditionalFormatting>,
) -> Option<LayoutBlock> {
    // §17.13.5.15 (issue #154): in the final view, a paragraph whose mark is
    // deleted and whose visible content is entirely deletion-stamped merges
    // away — emitting it would leave a phantom blank line where Word shows
    // nothing, shifting everything below by a line height. (True merging —
    // surviving content joining the *next* paragraph when only the mark is
    // deleted — is not modelled; such a paragraph keeps its line.)
    if !ctx.resolved.show_ins_del_marks && p.mark_deleted && paragraph_fully_deleted(p) {
        return None;
    }
    let (mut fragments, mut merged_props) = build_fragments(p, ctx, state, table_style, cond);
    // Drain immediately: this paragraph owns exactly the references its own
    // fragment collection recorded. Draining before the drop-cap early return
    // below keeps them from leaking into the next paragraph's batch, and
    // before `build_note_content` re-enters `build_fragments` (a footnote body
    // may itself carry references) from mixing the two levels together.
    let fn_refs = state.footnotes.take_pending();

    // §17.9.22: inject list label if paragraph has a numbering reference.
    super::list_label::inject_list_label(p, &mut fragments, &mut merged_props, ctx, state);

    // §17.3.1.6 / UAX #9 — after the label, which is part of the text.
    resolve_paragraph_bidi(&mut fragments, &merged_props, ctx);

    // §17.3.1.29: a paragraph with no runs still occupies one line — the
    // paragraph mark (¶) has a font-sized line height. Inject a `LineBreak` so
    // the layout phase treats it as a real line. Without this, the fragment
    // list is empty, `split_at_page_breaks` yields a single empty page-chunk,
    // `section::layout_section` drops it, and the line collapses to zero.
    //
    // The test is "the final segment carries nothing that draws", not "the
    // paragraph is empty" (issue #126). A paragraph whose only content is
    // `<w:br w:type="page"/>` has a *fragment*, so it slips past an
    // `is_empty()` check while still having nothing to show. [`MarkLine`]
    // makes that judgement and carries the reasoning for where the line goes.
    //
    // Table cells use §17.4.66 (trailing-empty-after-table is structural
    // and suppressed) in `build_cell_blocks` — that skip runs before this
    // function, so genuinely structural terminators never reach us.
    // Headers/footers have their own injection in
    // `build_header_footer_content` (§17.10.1) and do not call this.
    if MarkLine::of(&fragments) == MarkLine::NeedsOwnLine {
        let (family, size, ..) = resolve_paragraph_defaults(
            p,
            ctx.resolved,
            table_style.is_some(),
            state.shape_default_text_color,
            state.shape_default_font_family.as_deref(),
        );
        // §17.3.1.29 says the mark is formatted by `w:pPr/w:rPr`, and that is
        // the *whole* of it — `w:rFonts` as much as `w:sz`. Taking only the
        // size and leaving the family at the paragraph default measures the
        // mark's line in the wrong face, which is not a rounding error: on
        // `ELH_2025-12-18.docx` every one of these paragraphs asks for Arial
        // and got Calibri, whose line box is ~6% taller at the same size, so
        // each mark line came out 13.43 pt instead of 12.65 pt. That document
        // leaves 27.45 pt below its photo table for two such lines — 27.20 pt
        // of Arial, 28.87 pt of Calibri — so the wrong face was the difference
        // between fitting and spilling a page that holds nothing but the
        // running footer, ten times over.
        //
        // `font_props_from_run` is the same resolution every ordinary run
        // goes through, so the mark cannot drift from the text around it.
        let default_mark_props = model::RunProperties::default();
        let mark_font = font_props_from_run(
            p.mark_run_properties
                .as_ref()
                .unwrap_or(&default_mark_props),
            &family,
            size,
            ShapeAutoFit::NONE,
        );
        let line_height = ctx
            .measurer
            .default_line_height(&mark_font.family, mark_font.size);
        // **At the front, ahead of any break** — the position is load-bearing.
        // `split_at_page_breaks` cuts at the break, so `[LineBreak, PageBreak]`
        // yields `[[LineBreak], []]`: the mark's line is offered to the current
        // page under the normal fit test, and the empty trailing chunk defers
        // the break to the following block.
        //
        // Appending instead — putting the mark on the page the break lands on,
        // which is what a literal reading of §17.3.3.1 suggests, since the mark
        // is the last thing in the paragraph — was tried against
        // `ELH_2025-12-18.docx` and is worse: it adds a line above every table
        // that follows a break, pushing 602 pt of photo table into a 629 pt
        // band until the *next* paragraph spills instead. 34 pages against 23.
        // The ordering here is the one both renderers agree on.
        fragments.insert(0, Fragment::LineBreak { line_height });
    }

    // Word suppresses Hyperlink character style (blue/underline) for ToC
    // entries in print view. Strip visual hyperlink styling but keep the
    // click annotation URL.
    //
    // §17.7.4.9: identified by the resolved style's *primary style name*
    // (`toc 1` … `toc 9`, locale-independent), not by the `w:styleId`
    // spelling — a `starts_with("TOC")` test both over-matches (an unrelated
    // user style `TOCustom`) and under-matches (producers that don't spell
    // their ToC style IDs `TOC1`).
    let is_toc_entry = p
        .style_id
        .as_ref()
        .and_then(|id| ctx.resolved.styles.get(id))
        .is_some_and(|s| s.is_toc_entry);
    if is_toc_entry {
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
                    // Rare (TOC hyperlinks only); clones this fragment's shared
                    // font on write.
                    Rc::make_mut(font).underline = false;
                }
            }
        }
    }

    // §17.3.1.11: detect drop cap paragraph.
    if let Some(model::FrameKind::DropCap {
        style,
        lines,
        h_space: dc_h_space,
    }) = merged_props.frame_properties.cloned()
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
            .get()
            .and_then(|i| i.start)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO);
        let dc_indent_first = merged_props
            .indentation
            .get()
            .and_then(|i| i.first_line)
            .map(|fl| match fl {
                model::FirstLineIndent::FirstLine(v) => Pt::from(v),
                model::FirstLineIndent::Hanging(v) => -Pt::from(v),
                model::FirstLineIndent::None => Pt::ZERO,
            })
            .unwrap_or(Pt::ZERO);
        // §17.3.1.33: frame height from drop cap paragraph's exact line spacing.
        let frame_height =
            merged_props
                .spacing
                .get()
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

    let outline = super::convert::paragraph_outline(p, &merged_props, state);
    let mut style = paragraph_style_from_props(
        &merged_props,
        Pt::from(ctx.resolved.default_tab_stop),
        state.shape_auto_fit,
        super::convert::paragraph_locale(p, ctx.resolved),
        outline,
    );
    style.style_id = p.style_id.clone();

    // Attach pending drop cap to this paragraph.
    if let Some(dc) = pending_dropcap.take() {
        style.drop_cap = Some(dc);
    }

    let page_break_before = merged_props.page_break_before.unwrap_or(false);

    // §17.11.12: render a body for each footnote this paragraph referenced.
    // `collect_fragments` recorded them — id *and* the display number it
    // emitted as the superscript — as it walked, so the body and the mark
    // cannot disagree. Draining here also covers references nested inside
    // hyperlinks, fields, and text boxes, which the previous flat scan of
    // `p.content` missed entirely.
    let mut para_footnotes = Vec::new();
    for note in fn_refs {
        if let Some(content) = ctx.resolved.footnotes.get(&note.id) {
            let display = format!("{}", note.display);
            let notes = build_note_content(&display, content, ctx, state);
            for (_, frags, style) in notes {
                para_footnotes.push((frags, style));
            }
        }
    }

    // §20.4.2.3: extract floating (anchor) images and shapes from this
    // paragraph. Table cells emit their commands through `stack_blocks`,
    // which shifts them into page coordinates, so anchors inside a cell use
    // the stack frame. Body paragraphs emit in page-absolute coordinates.
    let frame = if table_style.is_some() {
        AnchorFrame::Stack
    } else {
        AnchorFrame::Page
    };
    let floating_images = extract_floating_images(p, ctx, state, frame);
    let floating_shapes = super::floating::extract_floating_shapes(
        p,
        ctx,
        state,
        frame,
        super::floating::ShapeAnchorClass::All,
    );

    Some(LayoutBlock::Paragraph {
        fragments,
        style,
        page_break_before,
        footnotes: para_footnotes,
        floating_images,
        floating_shapes,
    })
}

/// Build note content (footnotes or endnotes) with a display number prefix.
pub(super) fn build_note_content(
    display_num: &str,
    content: &[Block],
    ctx: &BuildContext,
    state: &mut BuildState,
) -> Vec<(
    String,
    Vec<Fragment>,
    crate::render::layout::paragraph::ParagraphStyle,
)> {
    // §17.3.1.19: a footnote or endnote body is not the document's main story,
    // so nothing in it is an outline position. Suspended for the whole call and
    // restored, for the same reason `build_header_footer_content` does it there
    // rather than at the paragraph — a `Block::Table` in a note reaches the
    // ordinary body builders.
    let outer = std::mem::replace(
        &mut state.outline,
        crate::render::layout::build::OutlineCollector::Excluded,
    );
    let results = build_note_blocks(display_num, content, ctx, state);
    state.outline = outer;
    results
}

fn build_note_blocks(
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
            // §17.11.12: a footnote body may itself carry references. We don't
            // render nested footnote bodies (matching the previous behaviour),
            // but the references must be drained so they aren't attributed to
            // the paragraph that hosts this note.
            let _ = state.footnotes.take_pending();

            // Prepend display number to the first paragraph. A comment body
            // reuses this builder with an empty display and no prefix.
            if i == 0 && !display_num.is_empty() && !frags.is_empty() {
                let num_text = format!("{}  ", display_num);
                // §17.8.3.2 / §17.3.2.14: fall back to the document-level spec
                // defaults rather than restating a font name here.
                let font = frags[0].font_props().cloned().unwrap_or_else(|| FontProps {
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: std::rc::Rc::from(super::SPEC_FALLBACK_FONT),
                    size: super::SPEC_DEFAULT_FONT_SIZE,
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
                });
                let ref_size =
                    font.size * crate::render::layout::fragment::SUPERSCRIPT_FONT_SIZE_RATIO;
                let ref_font = FontProps {
                    size: ref_size,
                    ..font
                };
                let (w, m) = ctx.measurer.measure(&num_text, &ref_font);
                frags.insert(
                    0,
                    Fragment::Text {
                        shaped: None,
                        level: BidiLevel::LTR,
                        text: Rc::from(num_text.as_str()),
                        font: Rc::new(ref_font),
                        color: RgbColor::BLACK,
                        shading: None,
                        border: None,
                        // §17.11.12: the number is followed by the two spaces
                        // built into `num_text`, so the note body may start on
                        // the next line if its first word doesn't fit.
                        break_after: crate::render::layout::fragment::BreakAfter::Opportunity,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: None,
                        baseline_offset: -(font.size
                            * crate::render::layout::fragment::NOTE_REF_BASELINE_OFFSET_RATIO),
                        text_offset: Pt::ZERO,
                        is_footnote_ref: false,
                    },
                );
            }
            // §17.3.1.6 / UAX #9 — after the note number, which is part of
            // the body's text.
            resolve_paragraph_bidi(&mut frags, &merged_props, ctx);
            let style = paragraph_style_from_props(
                &merged_props,
                Pt::from(ctx.resolved.default_tab_stop),
                state.shape_auto_fit,
                super::convert::paragraph_locale(p, ctx.resolved),
                // §17.3.1.19: always `None` — the collector is suspended for
                // this whole call. Asking rather than passing `None` keeps the
                // heading decision in one place for every path.
                super::convert::paragraph_outline(p, &merged_props, state),
            );
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
            endnotes.extend(build_note_content(&display, content, ctx, state));
        }
    }
}

/// Build fragments and resolved paragraph properties for a paragraph.
///
/// Handles the full cascade: table style → conditional → paragraph style →
/// doc defaults → fragment collection → image/underline population.
/// UAX #9: resolve every fragment's embedding level, once the paragraph's
/// fragment vector is **final**.
///
/// "Final" is the whole reason this is a separate step called from three
/// places rather than the tail of [`build_fragments`]: a list label
/// (§17.9.22) and a note body's number (§17.11.12) are prefixed afterwards,
/// and they are as much part of the paragraph's text as anything the document
/// wrote. A label left out of the analysis keeps the base level while the text
/// around it does not, and rule L2 then places it at the wrong end of a
/// `w:bidi` line.
pub(super) fn resolve_paragraph_bidi(
    fragments: &mut Vec<Fragment>,
    props: &model::ParagraphProperties,
    ctx: &BuildContext,
) {
    let measure =
        |text: &str, font: &FontProps| -> (Pt, crate::render::layout::fragment::TextMetrics) {
            ctx.measurer.measure(text, font)
        };
    crate::render::layout::fragment::assign_bidi_levels(
        fragments,
        super::convert::base_direction(props),
        &measure,
    );
    // Issue #139: then give every cluster a face that can actually draw it.
    // After bidi because a coverage boundary cannot change an embedding level,
    // so each piece inherits the one its fragment already carries; before
    // shaping because shaping re-measures against the resolved typeface and
    // has to see the family this pass may have changed.
    crate::render::layout::fragment::apply_font_fallback(
        fragments,
        ctx.measurer.fallback(),
        &measure,
    );
    // Then shaping, which needs the levels the line above resolved: a run's
    // direction is its embedding level, and nothing about the run's own
    // characters says whether it is a Latin phrase quoted inside Arabic.
    crate::render::layout::fragment::shape_complex_scripts(fragments, ctx.measurer);
}

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
        resolve_paragraph_defaults(
            para,
            ctx.resolved,
            table_style.is_some(),
            state.shape_default_text_color,
            state.shape_default_font_family.as_deref(),
        );

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
        if let Some(fs) = ts.run.font_size.cloned() {
            default_size = Pt::from(fs);
            run_defaults.font_size = Dup::from(Some(fs));
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
            if let Some(fs) = run_defaults.font_size.cloned() {
                default_size = Pt::from(fs);
            }
            if let Some(color) = run_defaults.color.cloned() {
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
        measurer: Some(ctx.measurer),
        auto_fit: state.shape_auto_fit,
        // §17.3.2.20: resolved here rather than taken from the caller
        // because fragments are built before the paragraph's style is —
        // `paragraph_locale` runs later, off this same cascade.
        locale_tag: super::convert::resolve_lang_tag(para, ctx.resolved),
        // Issue #154: `Some` renders revision marks; `None` is the
        // `w:revisionView` final view.
        revision_palette: ctx
            .resolved
            .show_ins_del_marks
            .then_some(&ctx.resolved.revision_colors),
        comment_marks: ctx.resolved.show_comment_marks,
    };
    let mut fragments = collect_fragments(
        &para.content,
        &frag_ctx,
        None,
        &measure,
        &mut state.footnotes,
        &mut state.endnote_counter,
        state.field_ctx,
    );
    populate_image_data(&mut fragments, ctx.media());
    populate_underline_metrics(&mut fragments, ctx.measurer);

    (fragments, merged_props)
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use super::*;
    use crate::model::dimension::Dimension;
    use crate::render::fonts::FontRegistry;
    use crate::render::layout::measurer::TextMeasurer;
    use crate::render::resolve::ResolvedDocument;

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

    fn para(content: Vec<model::Inline>) -> Paragraph {
        Paragraph {
            style_id: None,
            properties: model::ParagraphProperties::default(),
            mark_run_properties: None,
            content,
            rsids: model::ParagraphRevisionIds::default(),
            mark_deleted: false,
        }
    }

    /// A run whose only content is §17.3.3.1's explicit page break.
    fn page_break_run() -> model::Inline {
        model::Inline::TextRun(Box::new(model::TextRun {
            style_id: None,
            properties: model::RunProperties::default(),
            content: vec![model::RunElement::PageBreak],
            rsids: model::RevisionIds::default(),
            revision: None,
            comment: None,
        }))
    }

    fn text_run(s: &str) -> model::Inline {
        model::Inline::TextRun(Box::new(model::TextRun {
            style_id: None,
            properties: model::RunProperties::default(),
            content: vec![model::RunElement::Text(s.to_string())],
            rsids: model::RevisionIds::default(),
            revision: None,
            comment: None,
        }))
    }

    /// Run a closure with a live `BuildContext` + `BuildState`. A real Skia
    /// measurer is used, so assertions below are structural — never on
    /// platform-dependent metric values.
    fn with_ctx<R>(
        resolved: &ResolvedDocument,
        f: impl FnOnce(&BuildContext, &mut BuildState) -> R,
    ) -> R {
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved,
        };
        let mut state = BuildState::default();
        f(&ctx, &mut state)
    }

    fn resolved_style(paragraph: model::ParagraphProperties) -> ResolvedStyle {
        ResolvedStyle {
            paragraph,
            run: model::RunProperties::default(),
            table: None,
            table_style_overrides: Vec::new(),
            is_toc_entry: false,
        }
    }

    #[test]
    fn section_break_produces_no_layout_block() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let block = Block::SectionBreak(Box::default());
            let mut pending = None;
            assert!(
                build_block(&block, Pt::new(400.0), ctx, state, &mut pending).is_none(),
                "section breaks are consumed by resolve, not laid out"
            );
        });
    }

    /// §17.3.1.29: a paragraph with no runs still occupies one line. Without an
    /// injected `LineBreak` the fragment list is empty and `layout_section`
    /// drops the paragraph, collapsing it to zero height.
    #[test]
    fn empty_paragraph_gets_a_line_break_fragment() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let mut pending = None;
            let block = build_paragraph_block(&para(vec![]), ctx, state, &mut pending, None, None)
                .expect("empty paragraph still lays out");
            let LayoutBlock::Paragraph { fragments, .. } = block else {
                panic!("expected a paragraph block");
            };
            assert!(
                matches!(fragments.as_slice(), [Fragment::LineBreak { line_height }] if line_height.raw() > 0.0),
                "exactly one LineBreak with a real height"
            );
        });
    }

    /// Issue #126: a paragraph whose only content is `<w:br w:type="page"/>`
    /// is empty in every sense that matters — it draws nothing — but it does
    /// have a fragment, so an `is_empty()` test misses it and the paragraph
    /// collapses to zero height. Its mark still owns a line.
    ///
    /// The injected line goes **before** the break, and that is the assertion
    /// with teeth: `split_at_page_breaks` cuts here, so this order offers the
    /// mark's line to the current page and defers the break to the next block.
    /// Reversed, the break fires first and the document renders one page short
    /// of Word and LibreOffice.
    #[test]
    fn a_break_only_paragraph_gets_a_line_ahead_of_its_break() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let mut pending = None;
            let block = build_paragraph_block(
                &para(vec![page_break_run()]),
                ctx,
                state,
                &mut pending,
                None,
                None,
            )
            .expect("a break-only paragraph still lays out");
            let LayoutBlock::Paragraph { fragments, .. } = block else {
                panic!("expected a paragraph block");
            };
            assert!(
                matches!(
                    fragments.as_slice(),
                    [Fragment::LineBreak { line_height }, Fragment::PageBreak { .. }]
                        if line_height.raw() > 0.0
                ),
                "line then break, in that order: {fragments:?}",
            );
        });
    }

    /// §17.3.1.29: the mark's line is measured in the mark's *own* font, which
    /// `w:pPr/w:rPr` gives as a family as well as a size. Only the size used to
    /// be read, so a paragraph that asked for one face had its mark measured in
    /// another — on `ELH_2025-12-18.docx`, Arial asked for and Calibri
    /// measured, 13.43 pt where 12.65 pt was due. That document had 27.45 pt
    /// below a photo table for two such lines, so the difference cost it ten
    /// pages holding nothing but a running footer.
    ///
    /// The assertion is that the injected height is what the measurer reports
    /// for the *mark's* family, not the paragraph default's. It can only bite
    /// on a host that has two faces of different metrics, so it says which it
    /// used and steps aside when they measure alike rather than claiming to
    /// have checked something it could not.
    #[test]
    fn the_mark_line_is_measured_in_the_marks_own_font() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let (default_family, size, ..) =
                resolve_paragraph_defaults(&para(vec![]), ctx.resolved, false, None, None);
            let mark_family = "Courier New";
            let default_h = ctx.measurer.default_line_height(&default_family, size);
            let mark_h = ctx.measurer.default_line_height(mark_family, size);
            if (default_h.raw() - mark_h.raw()).abs() < 0.01 {
                eprintln!(
                    "skipped: this host measures {default_family} and {mark_family} \
                     identically ({default_h:?}), so the two cannot be told apart"
                );
                return;
            }

            let mut p = para(vec![page_break_run()]);
            p.mark_run_properties = Some(model::RunProperties {
                fonts: model::FontSet {
                    ascii: model::FontSlot::from_name(mark_family),
                    ..model::FontSet::default()
                },
                ..model::RunProperties::default()
            });

            let mut pending = None;
            let block = build_paragraph_block(&p, ctx, state, &mut pending, None, None)
                .expect("a break-only paragraph still lays out");
            let LayoutBlock::Paragraph { fragments, .. } = block else {
                panic!("expected a paragraph block");
            };
            let Some(Fragment::LineBreak { line_height }) = fragments.first() else {
                panic!("expected an injected mark line: {fragments:?}");
            };
            assert_eq!(
                line_height.raw(),
                mark_h.raw(),
                "mark line measured in {default_family} ({default_h:?}) instead of \
                 the requested {mark_family} ({mark_h:?})",
            );
        });
    }

    /// …and a paragraph that has real text *plus* a break keeps exactly what
    /// it had. The text already owns the line; injecting another would add a
    /// blank one.
    #[test]
    fn a_paragraph_with_text_and_a_break_gets_no_injected_line() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let mut pending = None;
            let block = build_paragraph_block(
                &para(vec![text_run("hi"), page_break_run()]),
                ctx,
                state,
                &mut pending,
                None,
                None,
            )
            .expect("lays out");
            let LayoutBlock::Paragraph { fragments, .. } = block else {
                panic!("expected a paragraph block");
            };
            assert!(
                !matches!(fragments.first(), Some(Fragment::LineBreak { .. })),
                "no line injected ahead of real text: {fragments:?}",
            );
        });
    }

    #[test]
    fn paragraph_with_content_gets_no_injected_line_break() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let mut pending = None;
            let block = build_paragraph_block(
                &para(vec![text_run("hi")]),
                ctx,
                state,
                &mut pending,
                None,
                None,
            )
            .expect("lays out");
            let LayoutBlock::Paragraph { fragments, .. } = block else {
                panic!("expected a paragraph block");
            };
            assert!(
                !fragments
                    .iter()
                    .any(|f| matches!(f, Fragment::LineBreak { .. })),
                "no LineBreak injected when the paragraph has runs"
            );
        });
    }

    /// §17.3.1.11: a drop-cap paragraph emits no block of its own — it is held
    /// aside and attached to the *following* paragraph.
    #[test]
    fn drop_cap_paragraph_is_deferred_onto_the_next_paragraph() {
        let resolved = empty_resolved();
        with_ctx(&resolved, |ctx, state| {
            let mut cap = para(vec![text_run("D")]);
            cap.properties.frame_properties = Dup::from(Some(model::FrameKind::DropCap {
                style: model::DropCap::Drop,
                lines: 3,
                h_space: None,
            }));

            let mut pending = None;
            assert!(
                build_paragraph_block(&cap, ctx, state, &mut pending, None, None).is_none(),
                "the drop-cap paragraph itself produces no block"
            );
            let held = pending
                .as_ref()
                .expect("drop cap held for the next paragraph");
            assert_eq!(held.lines, 3, "line span carried through");

            // The next paragraph consumes it.
            let next = build_paragraph_block(
                &para(vec![text_run("body")]),
                ctx,
                state,
                &mut pending,
                None,
                None,
            )
            .expect("lays out");
            let LayoutBlock::Paragraph { style, .. } = next else {
                panic!("expected a paragraph block");
            };
            assert!(
                style.drop_cap.is_some(),
                "attached to the following paragraph"
            );
            assert!(pending.is_none(), "and taken, so it attaches only once");
        });
    }

    /// §17.7.2 precedence inside a table cell: conditional formatting outranks
    /// the table style, which outranks document defaults. Asserted on the
    /// merged paragraph properties returned by `build_fragments`.
    #[test]
    fn conditional_formatting_outranks_table_style() {
        let mut resolved = empty_resolved();
        resolved.doc_defaults_paragraph.alignment = Dup::from(Some(model::Alignment::Start));
        let table_style = resolved_style(model::ParagraphProperties {
            alignment: Dup::from(Some(model::Alignment::Center)),
            ..Default::default()
        });

        with_ctx(&resolved, |ctx, state| {
            // Table style alone.
            let (_, props) = build_fragments(
                &para(vec![text_run("x")]),
                ctx,
                state,
                Some(&table_style),
                None,
            );
            assert_eq!(
                props.alignment,
                Dup::from(Some(model::Alignment::Center)),
                "table style beats doc defaults"
            );

            // Conditional formatting on top of it.
            let cond = CellConditionalFormatting {
                cell_properties: None,
                run_properties: None,
                paragraph_properties: Some(model::ParagraphProperties {
                    alignment: Dup::from(Some(model::Alignment::End)),
                    ..Default::default()
                }),
            };
            let (_, props) = build_fragments(
                &para(vec![text_run("x")]),
                ctx,
                state,
                Some(&table_style),
                Some(&cond),
            );
            assert_eq!(
                props.alignment,
                Dup::from(Some(model::Alignment::End)),
                "conditional formatting beats the table style"
            );
        });
    }

    /// The paragraph's own direct formatting outranks every table-level layer.
    #[test]
    fn direct_paragraph_properties_outrank_conditional_formatting() {
        let resolved = empty_resolved();
        let table_style = resolved_style(model::ParagraphProperties {
            alignment: Dup::from(Some(model::Alignment::Center)),
            ..Default::default()
        });
        let cond = CellConditionalFormatting {
            cell_properties: None,
            run_properties: None,
            paragraph_properties: Some(model::ParagraphProperties {
                alignment: Dup::from(Some(model::Alignment::End)),
                ..Default::default()
            }),
        };

        with_ctx(&resolved, |ctx, state| {
            let mut p = para(vec![text_run("x")]);
            p.properties.alignment = Dup::from(Some(model::Alignment::Both));
            let (_, props) = build_fragments(&p, ctx, state, Some(&table_style), Some(&cond));
            assert_eq!(
                props.alignment,
                Dup::from(Some(model::Alignment::Both)),
                "direct pPr wins over every table layer"
            );
        });
    }
}
