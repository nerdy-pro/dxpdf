use std::rc::Rc;

use crate::model::{
    Block, BorderStyle, FieldCharType, Inline, RunElement, RunProperties, TextRun, VerticalAlign,
};
use crate::render::dimension::Pt;
use crate::render::emoji::cluster::EmojiCluster;
use crate::render::geometry::PtSize;
use crate::render::resolve::color::RgbColor;

use super::segment::{build_inline_units, InlineUnit, SegmentPiece};
use super::text::{
    emit_emoji_or_fallback, emit_text_fragments, emit_text_words, resolve_highlight_color,
    TextRunStyle,
};
use super::{
    font_props_from_run, to_roman_lower, FontProps, Fragment, FragmentBorder, TextMetrics,
    SUBSCRIPT_HEIGHT_OFFSET_RATIO, SUPERSCRIPT_ASCENT_OFFSET_RATIO, SUPERSCRIPT_FONT_SIZE_RATIO,
};

/// §17.3.2.4: convert a run-level [`crate::model::Border`] into a render-side
/// [`FragmentBorder`], filtering out the spec's "no border" sentinel
/// ([`BorderStyle::None`]).
///
/// `<w:bdr w:val="nil"/>` and `<w:bdr w:val="none"/>` (§17.18.2 ST_Border)
/// both signal "no border"; the parser collapses them to `BorderStyle::None`
/// in a `Some(Border { ... })`. The model preserves the explicit `Some` so
/// it can override an inherited border in the §17.7.2 cascade — but at the
/// render boundary we drop the variant, otherwise the painter would draw
/// a hairline box around every word.
pub(super) fn run_border_to_fragment(
    border: Option<&crate::model::Border>,
) -> Option<FragmentBorder> {
    let b = border?;
    if b.style == BorderStyle::None {
        return None;
    }
    Some(FragmentBorder {
        width: Pt::from(b.width),
        color: crate::render::resolve::color::resolve_color(
            b.color,
            crate::render::resolve::color::ColorContext::Text,
        ),
        space: Pt::new(b.space.raw() as f32),
    })
}

/// §17.7.2: resolve the effective styling of a single run by walking the
/// cascade (direct → character style → paragraph run defaults), then
/// translating to render-side `FontProps` + `TextRunStyle`.
///
/// This is the single source of truth for run-level styling. Both the
/// per-run path (`Discrete TextRun`) and the per-segment-piece path
/// (cross-run cluster reassembly via `segment.rs`) call it — for cross-
/// run clusters, the *base run*'s styling drives the entire piece per
/// the design in `docs/cross-run-cluster-reassembly.md`.
#[allow(clippy::too_many_arguments)] // the cascade has many independent inputs by spec
fn resolve_run_styling<F>(
    tr: &TextRun,
    default_family: &str,
    default_size: Pt,
    default_color: RgbColor,
    resolved_styles: Option<
        &std::collections::HashMap<
            crate::model::StyleId,
            crate::render::resolve::styles::ResolvedStyle,
        >,
    >,
    paragraph_run_defaults: Option<&RunProperties>,
    theme: Option<&crate::model::Theme>,
    measure_text: &F,
) -> (FontProps, TextRunStyle)
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let mut effective_props = tr.properties.clone();
    // §17.3.2.26: resolve theme font references before merging.
    if let Some(th) = theme {
        crate::render::resolve::fonts::resolve_font_set_themes(&mut effective_props.fonts, th);
    }
    if let (Some(ref style_id), Some(styles)) = (&tr.style_id, resolved_styles) {
        if let Some(resolved_style) = styles.get(style_id) {
            crate::render::resolve::properties::merge_run_properties(
                &mut effective_props,
                &resolved_style.run,
            );
        }
    }
    if let Some(para_run) = paragraph_run_defaults {
        crate::render::resolve::properties::merge_run_properties(&mut effective_props, para_run);
    }

    let mut font = font_props_from_run(&effective_props, default_family, default_size);
    let color = effective_props
        .color
        .map(|c| {
            crate::render::resolve::color::resolve_color(
                c,
                crate::render::resolve::color::ColorContext::Text,
            )
        })
        .unwrap_or(default_color);
    // §17.3.2.32 / §17.3.2.15: shading or highlight as background.
    let shading = effective_props
        .shading
        .as_ref()
        .map(|s| {
            crate::render::resolve::color::resolve_color(
                s.fill,
                crate::render::resolve::color::ColorContext::Background,
            )
        })
        .or_else(|| effective_props.highlight.map(resolve_highlight_color));

    // §17.3.2.42: vertical alignment (super/sub).
    let mut baseline_offset = match effective_props.vertical_align {
        Some(VerticalAlign::Superscript) => {
            let (_, base_m) = measure_text("X", &font);
            font.size = font.size * SUPERSCRIPT_FONT_SIZE_RATIO;
            -(base_m.ascent * SUPERSCRIPT_ASCENT_OFFSET_RATIO)
        }
        Some(VerticalAlign::Subscript) => {
            let (_, base_m) = measure_text("X", &font);
            font.size = font.size * SUPERSCRIPT_FONT_SIZE_RATIO;
            base_m.height() * SUBSCRIPT_HEIGHT_OFFSET_RATIO
        }
        _ => Pt::ZERO,
    };
    // §17.3.2.19: w:position — vertical baseline offset in half-points.
    if let Some(pos) = effective_props.position {
        baseline_offset += Pt::from(pos);
    }

    // §17.3.2.4: run-level border (filtered to drop BorderStyle::None).
    let border = run_border_to_fragment(effective_props.border.as_ref());

    let text_style = TextRunStyle {
        color,
        shading,
        border,
        baseline_offset,
    };
    (font, text_style)
}

/// §17.16.4.1: context for evaluating dynamic fields (PAGE, NUMPAGES).
#[derive(Clone, Copy, Default)]
pub struct FieldContext {
    /// Current page number (1-based).
    pub page_number: Option<usize>,
    /// Total page count in the document.
    pub num_pages: Option<usize>,
}

/// §17.16.4.1: evaluate a parsed field instruction against the current context.
/// Returns the substituted text for PAGE/NUMPAGES, or None for other fields
/// or when no context is available.
fn evaluate_field_instruction(
    instruction: &crate::field::FieldInstruction,
    ctx: FieldContext,
) -> Option<String> {
    match instruction {
        crate::field::FieldInstruction::Page { .. } => ctx.page_number.map(|n| n.to_string()),
        crate::field::FieldInstruction::NumPages { .. } => ctx.num_pages.map(|n| n.to_string()),
        _ => None,
    }
}

/// Build a text fragment for a substituted field value, using the paragraph's
/// default font properties.
fn make_field_text_fragment<F>(
    text: Rc<str>,
    default_family: &str,
    default_size: Pt,
    default_color: crate::render::resolve::color::RgbColor,
    measure_text: &F,
) -> Fragment
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let font = FontProps {
        family: Rc::from(default_family),
        size: default_size,
        bold: false,
        italic: false,
        underline: false,
        char_spacing: Pt::ZERO,
        text_scale: 1.0,
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    };
    let (w, m) = measure_text(&text, &font);
    Fragment::Text {
        text,
        font,
        color: default_color,
        shading: None,
        border: None,
        width: w,
        trimmed_width: w,
        metrics: m,
        hyperlink_url: None,
        baseline_offset: Pt::ZERO,
        text_offset: Pt::ZERO,
    }
}

/// Invariant context threaded through all recursive `collect_fragments` calls.
pub struct FragmentCtx<'a> {
    pub default_family: &'a str,
    pub default_size: Pt,
    pub default_color: RgbColor,
    pub resolved_styles: Option<
        &'a std::collections::HashMap<
            crate::model::StyleId,
            crate::render::resolve::styles::ResolvedStyle,
        >,
    >,
    pub paragraph_run_defaults: Option<&'a RunProperties>,
    pub theme: Option<&'a crate::model::Theme>,
    /// Measurer used by the emoji pipeline for typeface resolution and
    /// raster-backend metrics. `None` disables the emoji path entirely —
    /// callers without a font registry (most unit tests) pass `None` and
    /// emoji codepoints flow through the existing text path unchanged.
    pub measurer: Option<&'a crate::render::layout::measurer::TextMeasurer<'a>>,
}

/// Walk inline content and collect fragments.
/// `measure_text` is a callback that measures text width/height/ascent for a given font.
/// `resolved_styles` is used to look up character styles (w:rStyle) on text runs.
///
/// Returns fragments suitable for the line-fitting algorithm.
pub fn collect_fragments<F>(
    inlines: &[Inline],
    ctx: &FragmentCtx<'_>,
    hyperlink_url: Option<&str>,
    measure_text: &F,
    footnote_counter: &mut u32,
    endnote_counter: &mut u32,
    field_ctx: FieldContext,
) -> Vec<Fragment>
where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics), // (width, metrics)
{
    let default_family = ctx.default_family;
    let default_size = ctx.default_size;
    let default_color = ctx.default_color;
    let resolved_styles = ctx.resolved_styles;
    let paragraph_run_defaults = ctx.paragraph_run_defaults;
    let theme = ctx.theme;
    let mut fragments = Vec::new();
    let mut field_depth: i32 = 0; // tracks nested complex field state
    let mut field_instr = String::new(); // accumulated instruction text for current complex field
                                         // §17.16.19: field substitution state for complex fields.
                                         // Pending = substitution text waiting for the first result TextRun's formatting.
                                         // Emitted = substitution was rendered, skip remaining result TextRuns until End.
    let mut field_sub_pending: Option<String> = None;
    let mut field_sub_emitted = false;
    // Pre-pass: join consecutive text-only TextRuns into segments so
    // UAX #29 grapheme clusters reassemble across `<w:rFonts>`-induced
    // run splits (keycap `1️⃣`, ZWJ family, modifier sequence, …).
    // See `docs/cross-run-cluster-reassembly.md`.
    let units = build_inline_units(inlines);
    for unit in units {
        match unit {
            InlineUnit::TextSegment(seg) => {
                // Field state (mirrors the per-run logic below). Field chars
                // appear as Discrete Inlines and break segment joining, so
                // a TextSegment is always entirely inside one field zone.
                if field_depth > 0 || field_sub_emitted {
                    continue;
                }

                // §17.16.19: pending substitution uses the segment's first run
                // for formatting (per cross-run cluster cascade rule).
                if let Some(sub) = field_sub_pending.take() {
                    let base_run = seg.char_runs()[0];
                    let (font, text_style) = resolve_run_styling(
                        base_run,
                        default_family,
                        default_size,
                        default_color,
                        resolved_styles,
                        paragraph_run_defaults,
                        theme,
                        measure_text,
                    );
                    field_sub_emitted = true;
                    emit_text_fragments(
                        &sub,
                        &font,
                        &text_style,
                        hyperlink_url,
                        measure_text,
                        ctx.measurer,
                        &mut fragments,
                    );
                    continue;
                }

                // Normal segment: classify and emit each piece using its
                // own (or for emoji, base) run's resolved styling.
                for piece in seg.classify() {
                    match piece {
                        SegmentPiece::Text { run, text } => {
                            let (font, text_style) = resolve_run_styling(
                                run,
                                default_family,
                                default_size,
                                default_color,
                                resolved_styles,
                                paragraph_run_defaults,
                                theme,
                                measure_text,
                            );
                            // Pre-classified text: bypass cluster::classify
                            // by going straight to the word-split path.
                            emit_text_words(
                                &text,
                                &font,
                                &text_style,
                                hyperlink_url,
                                measure_text,
                                &mut fragments,
                            );
                        }
                        SegmentPiece::Emoji {
                            base_run,
                            text,
                            presentation,
                            structure,
                        } => {
                            let (font, text_style) = resolve_run_styling(
                                base_run,
                                default_family,
                                default_size,
                                default_color,
                                resolved_styles,
                                paragraph_run_defaults,
                                theme,
                                measure_text,
                            );
                            if let Some(measurer) = ctx.measurer {
                                let cluster = EmojiCluster {
                                    text: &text,
                                    presentation,
                                    structure,
                                };
                                emit_emoji_or_fallback(
                                    &cluster,
                                    &font,
                                    &text_style,
                                    hyperlink_url,
                                    measure_text,
                                    measurer,
                                    &mut fragments,
                                );
                            } else {
                                // No measurer (test path): fall through to
                                // text — the cluster's codepoints survive
                                // in the PDF text stream verbatim.
                                emit_text_words(
                                    &text,
                                    &font,
                                    &text_style,
                                    hyperlink_url,
                                    measure_text,
                                    &mut fragments,
                                );
                            }
                        }
                    }
                }
            }
            InlineUnit::Discrete(inline) => match inline {
                Inline::TextRun(tr) => {
                    // A text-only TextRun would have been a TextSegment; this
                    // branch handles runs whose content includes Tab,
                    // LineBreak, PageBreak, ColumnBreak, or
                    // LastRenderedPageBreak.
                    if field_depth > 0 || field_sub_emitted {
                        continue;
                    }

                    let (font, text_style) = resolve_run_styling(
                        tr,
                        default_family,
                        default_size,
                        default_color,
                        resolved_styles,
                        paragraph_run_defaults,
                        theme,
                        measure_text,
                    );

                    if field_sub_pending.is_some() {
                        let sub = field_sub_pending.take().unwrap();
                        field_sub_emitted = true;
                        emit_text_fragments(
                            &sub,
                            &font,
                            &text_style,
                            hyperlink_url,
                            measure_text,
                            ctx.measurer,
                            &mut fragments,
                        );
                    } else {
                        for element in &tr.content {
                            match element {
                                RunElement::Text(text) => {
                                    emit_text_fragments(
                                        text,
                                        &font,
                                        &text_style,
                                        hyperlink_url,
                                        measure_text,
                                        ctx.measurer,
                                        &mut fragments,
                                    );
                                }
                                RunElement::Tab => {
                                    fragments.push(Fragment::Tab {
                                        line_height: font.size,
                                        fitting_width: None,
                                    });
                                }
                                RunElement::LineBreak(_) => {
                                    fragments.push(Fragment::LineBreak {
                                        line_height: font.size,
                                    });
                                }
                                RunElement::PageBreak => {
                                    fragments.push(Fragment::PageBreak {
                                        line_height: font.size,
                                    });
                                }
                                RunElement::ColumnBreak => {
                                    fragments.push(Fragment::ColumnBreak);
                                }
                                RunElement::LastRenderedPageBreak => {}
                            }
                        }
                    }
                }
                Inline::Image(img) => {
                    // Only render INLINE images as fragments.
                    // Anchor (floating) images are handled separately in build.rs.
                    if matches!(img.placement, crate::model::ImagePlacement::Inline { .. }) {
                        if let Some(rel_id) =
                            crate::render::resolve::images::extract_image_rel_id(img)
                        {
                            let w = Pt::from(img.extent.width);
                            let h = Pt::from(img.extent.height);
                            fragments.push(Fragment::Image {
                                size: PtSize::new(w, h),
                                rel_id: rel_id.as_str().to_string(),
                                image_data: None,
                            });
                        }
                    }
                }
                Inline::Hyperlink(link) => {
                    let url: Option<&str> = match &link.target {
                        crate::model::HyperlinkTarget::External(rel_id) => Some(rel_id.as_str()),
                        crate::model::HyperlinkTarget::Internal { anchor } => Some(anchor.as_str()),
                    };
                    let mut sub = collect_fragments(
                        &link.content,
                        ctx,
                        url,
                        measure_text,
                        footnote_counter,
                        endnote_counter,
                        field_ctx,
                    );
                    fragments.append(&mut sub);
                }
                Inline::Field(field) => {
                    // §17.16.18: simple field — check for dynamic substitution.
                    let substituted = evaluate_field_instruction(&field.instruction, field_ctx);
                    if let Some(text) = substituted {
                        fragments.push(make_field_text_fragment(
                            Rc::from(text.as_str()),
                            default_family,
                            default_size,
                            default_color,
                            measure_text,
                        ));
                    } else {
                        let mut sub = collect_fragments(
                            &field.content,
                            ctx,
                            hyperlink_url,
                            measure_text,
                            footnote_counter,
                            endnote_counter,
                            field_ctx,
                        );
                        fragments.append(&mut sub);
                    }
                }
                Inline::FieldChar(fc) => {
                    // §17.16.18: complex field state machine:
                    // Begin → InstrText... → Separate → result runs → End
                    match fc.field_char_type {
                        FieldCharType::Begin => {
                            field_depth += 1;
                            field_instr.clear();
                            field_sub_pending = None;
                            field_sub_emitted = false;
                        }
                        FieldCharType::Separate => {
                            // §17.16.4.1: parse accumulated instruction, evaluate
                            // PAGE/NUMPAGES if field context is available.
                            if let Ok(parsed) = crate::field::parse(&field_instr) {
                                field_sub_pending = evaluate_field_instruction(&parsed, field_ctx);
                            }
                            field_depth -= 1; // now collect result runs (unless substituted)
                        }
                        FieldCharType::End => {
                            // If a substitution was pending but no result TextRun
                            // was present to provide formatting, emit with defaults.
                            if let Some(text) = field_sub_pending.take() {
                                fragments.push(make_field_text_fragment(
                                    Rc::from(text.as_str()),
                                    default_family,
                                    default_size,
                                    default_color,
                                    measure_text,
                                ));
                            }
                            field_sub_emitted = false;
                        }
                    }
                }
                Inline::InstrText(text) => {
                    // Accumulate instruction text for complex field parsing.
                    if field_depth > 0 {
                        field_instr.push_str(text);
                    }
                }
                Inline::AlternateContent(ac) => {
                    // Pick fallback content (safest for PDF rendering)
                    if let Some(ref fallback) = ac.fallback {
                        let mut sub = collect_fragments(
                            fallback,
                            ctx,
                            hyperlink_url,
                            measure_text,
                            footnote_counter,
                            endnote_counter,
                            field_ctx,
                        );
                        fragments.append(&mut sub);
                    }
                }
                Inline::Symbol(sym) => {
                    let font = FontProps {
                        family: Rc::from(sym.font.as_str()),
                        size: default_size,
                        bold: false,
                        italic: false,
                        underline: false,
                        char_spacing: Pt::ZERO,
                        text_scale: 1.0,
                        underline_position: Pt::ZERO,
                        underline_thickness: Pt::ZERO,
                    };
                    let ch = char::from_u32(sym.char_code as u32).unwrap_or('\u{FFFD}');
                    let text = ch.to_string();
                    let (w, m) = measure_text(&text, &font);
                    fragments.push(Fragment::Text {
                        text: Rc::from(text.as_str()),
                        font,
                        color: RgbColor::BLACK,
                        shading: None,
                        border: None,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: hyperlink_url.map(String::from),
                        baseline_offset: Pt::ZERO,
                        text_offset: Pt::ZERO,
                    });
                }
                // Bookmark target — emit as zero-width named destination.
                Inline::BookmarkStart { name, .. } => {
                    fragments.push(Fragment::Bookmark { name: name.clone() });
                }
                // Non-visual inlines — skip
                Inline::BookmarkEnd(_)
                | Inline::Separator
                | Inline::ContinuationSeparator
                | Inline::FootnoteRefMark
                | Inline::EndnoteRefMark => {}
                // §17.11.12: footnote reference — render as superscript number.
                Inline::FootnoteRef(_note_id) => {
                    *footnote_counter += 1;
                    let num_text = format!("{}", *footnote_counter);
                    // §17.11.12: footnote reference uses superscript at 58% size.
                    let ref_size = default_size * 0.58;
                    let ref_font = FontProps {
                        family: std::rc::Rc::from(default_family),
                        size: ref_size,
                        bold: false,
                        italic: false,
                        underline: false,
                        char_spacing: Pt::ZERO,
                        text_scale: 1.0,
                        underline_position: Pt::ZERO,
                        underline_thickness: Pt::ZERO,
                    };
                    let (w, m) = measure_text(&num_text, &ref_font);
                    // Superscript baseline offset: raise by ~40% of the full-size ascent.
                    let baseline_offset = -(default_size * 0.4);
                    fragments.push(Fragment::Text {
                        text: Rc::from(num_text.as_str()),
                        font: ref_font,
                        color: default_color,
                        shading: None,
                        border: None,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: None,
                        baseline_offset,
                        text_offset: Pt::ZERO,
                    });
                }
                // §17.11.2: endnote reference — render as superscript Roman numeral.
                Inline::EndnoteRef(_note_id) => {
                    *endnote_counter += 1;
                    let num_text = to_roman_lower(*endnote_counter);
                    let ref_size = default_size * 0.58;
                    let ref_font = FontProps {
                        family: std::rc::Rc::from(default_family),
                        size: ref_size,
                        bold: false,
                        italic: false,
                        underline: false,
                        char_spacing: Pt::ZERO,
                        text_scale: 1.0,
                        underline_position: Pt::ZERO,
                        underline_thickness: Pt::ZERO,
                    };
                    let (w, m) = measure_text(&num_text, &ref_font);
                    let baseline_offset = -(default_size * 0.4);
                    fragments.push(Fragment::Text {
                        text: Rc::from(num_text.as_str()),
                        font: ref_font,
                        color: default_color,
                        shading: None,
                        border: None,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: None,
                        baseline_offset,
                        text_offset: Pt::ZERO,
                    });
                }
                Inline::Pict(pict) => {
                    // Render text content from VML text box shapes inline.
                    // Does not handle absolute positioning — text appears inline
                    // with the surrounding paragraph.
                    for shape in &pict.shapes {
                        if let Some(ref text_box) = shape.text_box {
                            for block in &text_box.content {
                                if let Block::Paragraph(p) = block {
                                    let pict_ctx = FragmentCtx {
                                        default_family,
                                        default_size,
                                        default_color,
                                        resolved_styles,
                                        paragraph_run_defaults: p.mark_run_properties.as_ref(),
                                        theme,
                                        measurer: ctx.measurer,
                                    };
                                    let mut sub = collect_fragments(
                                        &p.content,
                                        &pict_ctx,
                                        hyperlink_url,
                                        measure_text,
                                        footnote_counter,
                                        endnote_counter,
                                        field_ctx,
                                    );
                                    fragments.append(&mut sub);
                                }
                            }
                        }
                    }
                }
            },
        }
    }

    fragments
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::{Dimension, HalfPoints};
    use crate::model::*;

    /// Dummy measurer: width = text.len() * 6.0, ascent = 10.0, descent = 2.0
    fn dummy_measure(text: &str, _font: &FontProps) -> (Pt, TextMetrics) {
        (
            Pt::new(text.len() as f32 * 6.0),
            TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(2.0),
                leading: Pt::ZERO,
            },
        )
    }

    fn default_ctx(size: f32) -> FragmentCtx<'static> {
        FragmentCtx {
            default_family: "Default",
            default_size: Pt::new(size),
            default_color: RgbColor::BLACK,
            resolved_styles: None,
            paragraph_run_defaults: None,
            theme: None,
            measurer: None,
        }
    }

    // ── §17.3.2.4 / §17.18.2 run-level border tri-state ─────────────────
    //
    // The cascade may carry a child run whose `<w:bdr w:val="nil"/>`
    // (or "none") explicitly turns off an inherited border. The model
    // preserves this as `Some(Border { style: BorderStyle::None, .. })`
    // so the §17.7.2 merge can distinguish "explicit no border" from
    // "field absent → inherit". At the render boundary we must drop the
    // sentinel; otherwise the painter draws a hairline box around every
    // word in any Word-saved doc (Word emits `<w:bdr w:val="nil"/>` in
    // the default rPrDefault for the entire document).

    fn border_with_style(style: BorderStyle) -> crate::model::Border {
        crate::model::Border {
            style,
            width: Dimension::new(0),
            space: Dimension::new(0),
            color: crate::model::Color::Auto,
        }
    }

    #[test]
    fn run_border_absent_yields_no_fragment_border() {
        assert!(run_border_to_fragment(None).is_none());
    }

    #[test]
    fn run_border_explicit_none_yields_no_fragment_border() {
        let b = border_with_style(BorderStyle::None);
        assert!(
            run_border_to_fragment(Some(&b)).is_none(),
            "<w:bdr w:val=\"nil\"/> / \"none\" must NOT produce a render-side border"
        );
    }

    #[test]
    fn run_border_actual_style_yields_fragment_border() {
        let b = border_with_style(BorderStyle::Single);
        assert!(
            run_border_to_fragment(Some(&b)).is_some(),
            "explicit Single border must reach the painter"
        );
    }

    fn text_run(text: &str) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    fn text_run_with_font(text: &str, font: &str, size: i64) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties {
                fonts: FontSet {
                    ascii: FontSlot::from_name(font),
                    ..Default::default()
                },
                font_size: Some(Dimension::<HalfPoints>::new(size)),
                ..Default::default()
            },
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    #[test]
    fn single_text_run() {
        let inlines = vec![text_run("hello")];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        assert_eq!(frags[0].width().raw(), 30.0); // 5 * 6
        assert_eq!(frags[0].height().raw(), 12.0);
    }

    #[test]
    fn text_run_uses_run_font() {
        let inlines = vec![text_run_with_font("hi", "Arial", 24)];
        let ctx = default_ctx(10.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        if let Fragment::Text { font, .. } = &frags[0] {
            assert_eq!(&*font.family, "Arial");
            assert_eq!(font.size.raw(), 12.0); // 24 half-points = 12pt
        } else {
            panic!("expected Text fragment");
        }
    }

    #[test]
    fn tab_produces_tab_fragment() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Tab],
            rsids: RevisionIds::default(),
        }))];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        assert!(matches!(frags[0], Fragment::Tab { .. }));
    }

    #[test]
    fn line_break_produces_break_fragment() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::LineBreak(BreakKind::TextWrapping)],
            rsids: RevisionIds::default(),
        }))];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        assert!(frags[0].is_line_break());
    }

    #[test]
    fn hyperlink_recurses_into_content() {
        let inlines = vec![Inline::Hyperlink(Hyperlink {
            target: HyperlinkTarget::External(RelId::new("rId1")),
            content: vec![text_run("click me")],
        })];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 2, "split into 'click ' and 'me'");
        if let Fragment::Text {
            hyperlink_url,
            text,
            ..
        } = &frags[0]
        {
            assert_eq!(&**text, "click ");
            assert_eq!(hyperlink_url.as_deref(), Some("rId1"));
        } else {
            panic!("expected Text fragment");
        }
    }

    #[test]
    fn complex_field_skips_instructions_collects_result() {
        // FieldChar::Begin -> InstrText("PAGE") -> FieldChar::Separate -> TextRun("3") -> FieldChar::End
        let inlines = vec![
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::Begin,
                dirty: None,
                fld_lock: None,
            }),
            Inline::InstrText("PAGE".into()),
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::Separate,
                dirty: None,
                fld_lock: None,
            }),
            text_run("3"),
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::End,
                dirty: None,
                fld_lock: None,
            }),
        ];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        // Should only have the "3" result, not "PAGE"
        assert_eq!(frags.len(), 1);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(&**text, "3");
        }
    }

    #[test]
    fn bookmarks_and_separators_skipped() {
        let inlines = vec![
            Inline::BookmarkStart {
                id: BookmarkId::new(1),
                name: "bm1".into(),
            },
            text_run("visible"),
            Inline::BookmarkEnd(BookmarkId::new(1)),
            Inline::Separator,
            Inline::ContinuationSeparator,
            Inline::FootnoteRefMark,
            Inline::EndnoteRefMark,
            // LastRenderedPageBreak is now inside RunElement, not Inline
        ];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        // BookmarkStart produces a Bookmark fragment, text run produces a Text fragment.
        assert_eq!(
            frags.len(),
            2,
            "bookmark + text run should produce fragments"
        );
        assert!(matches!(frags[0], Fragment::Bookmark { .. }));
        assert!(matches!(frags[1], Fragment::Text { .. }));
    }

    #[test]
    fn alternate_content_uses_fallback() {
        let inlines = vec![Inline::AlternateContent(AlternateContent {
            choices: vec![McChoice {
                requires: McRequires::Wps,
                content: vec![text_run("choice")],
            }],
            fallback: Some(vec![text_run("fallback")]),
        })];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(&**text, "fallback");
        }
    }

    #[test]
    fn empty_text_run_produces_no_fragment() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Text(String::new())],
            rsids: RevisionIds::default(),
        }))];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );
        assert!(frags.is_empty());
    }

    #[test]
    fn symbol_produces_text_fragment() {
        let inlines = vec![Inline::Symbol(Symbol {
            font: "Wingdings".into(),
            char_code: 0x46, // 'F'
        })];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { font, text, .. } = &frags[0] {
            assert_eq!(&*font.family, "Wingdings");
            assert_eq!(&**text, "F");
        }
    }

    #[test]
    fn simple_field_collects_content() {
        let inlines = vec![Inline::Field(Field {
            instruction: crate::field::FieldInstruction::Page {
                switches: Default::default(),
            },
            content: vec![text_run("5")],
        })];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(&**text, "5");
        }
    }

    #[test]
    fn multi_word_text_run_splits_into_fragments() {
        let inlines = vec![text_run("hello world foo")];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut 0,
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 3);
        if let Fragment::Text { text, .. } = &frags[0] {
            assert_eq!(&**text, "hello ");
        }
        if let Fragment::Text { text, .. } = &frags[1] {
            assert_eq!(&**text, "world ");
        }
        if let Fragment::Text { text, .. } = &frags[2] {
            assert_eq!(&**text, "foo");
        }
    }
}
