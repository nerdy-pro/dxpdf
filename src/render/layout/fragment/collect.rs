use std::rc::Rc;

use crate::model::{Block, FieldCharType, Inline, RunElement, RunProperties, VerticalAlign};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::resolve::color::RgbColor;

use super::{
    font_props_from_run, Fragment, FragmentBorder, FontProps, TextMetrics,
    SUBSCRIPT_HEIGHT_OFFSET_RATIO, SUPERSCRIPT_ASCENT_OFFSET_RATIO, SUPERSCRIPT_FONT_SIZE_RATIO,
    to_roman_lower,
};
use super::text::{emit_text_fragments, resolve_highlight_color, TextRunStyle};

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
    for inline in inlines {
        match inline {
            Inline::TextRun(tr) => {
                // Skip field instruction text (between Begin and Separate).
                if field_depth > 0 {
                    continue;
                }
                // Skip remaining result runs after substitution was emitted.
                if field_sub_emitted {
                    continue;
                }
                // Run property cascade per §17.7.2:
                // direct run properties → character style (w:rStyle) → paragraph style run defaults.
                let mut effective_props = tr.properties.clone();
                // §17.3.2.26: resolve theme font references before merging,
                // so theme-derived names take precedence over explicit names
                // from lower-priority levels in the cascade.
                if let Some(th) = theme {
                    crate::render::resolve::fonts::resolve_font_set_themes(
                        &mut effective_props.fonts,
                        th,
                    );
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
                    crate::render::resolve::properties::merge_run_properties(
                        &mut effective_props,
                        para_run,
                    );
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
                // §17.3.2.32: run-level shading (background behind text).
                // §17.3.2.15: highlight color (fixed palette) takes effect when shading is absent.
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

                // §17.3.2.42: vertical alignment (superscript/subscript).
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

                // §17.3.2.4: run-level border.
                let border = effective_props.border.as_ref().map(|b| FragmentBorder {
                    width: Pt::from(b.width),
                    color: crate::render::resolve::color::resolve_color(
                        b.color,
                        crate::render::resolve::color::ColorContext::Text,
                    ),
                    space: Pt::new(b.space.raw() as f32),
                });

                // §17.16.19: if a field substitution is pending, use the
                // substituted text with this TextRun's resolved formatting.
                let text_style = TextRunStyle { color, shading, border, baseline_offset };
                if field_sub_pending.is_some() {
                    let sub = field_sub_pending.take().unwrap();
                    field_sub_emitted = true;
                    emit_text_fragments(
                        &sub,
                        &font,
                        &text_style,
                        hyperlink_url,
                        measure_text,
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
                                fragments.push(Fragment::LineBreak {
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
                    if let Some(rel_id) = crate::render::resolve::images::extract_image_rel_id(img)
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
        }
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
