use std::rc::Rc;

use crate::model::{
    Block, FieldCharType, Inline, NoteId, RunElement, RunProperties, TextRun, VerticalAlign,
};
use crate::render::dimension::Pt;
use crate::render::emoji::cluster::EmojiCluster;
use crate::render::geometry::PtSize;
use crate::render::resolve::color::RgbColor;

use super::segment::{build_inline_units, InlineUnit, SegmentPiece};
use super::text::{
    emit_emoji_or_fallback, emit_text_fragments, emit_text_unit, emit_text_words,
    resolve_highlight_color, TextRunStyle,
};
use super::{
    font_props_from_run, to_roman_lower, BreakAfter, FontProps, Fragment, FragmentBorder,
    LinkTarget, TextMetrics, SUBSCRIPT_HEIGHT_OFFSET_RATIO, SUPERSCRIPT_ASCENT_OFFSET_RATIO,
    SUPERSCRIPT_FONT_SIZE_RATIO,
};
use crate::i18n::bidi::BidiLevel;
use crate::render::fonts::Toggle;

/// §17.11.12: a footnote reference recorded while walking a paragraph's inlines.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RecordedFootnote {
    /// The referenced note.
    pub id: NoteId,
    /// The sequential display number emitted as this reference's superscript.
    pub display: u32,
}

/// §17.11.12: footnote numbering state threaded through fragment collection.
///
/// The display counter and the record of *which* notes were referenced advance
/// at a single site (the private `record` method), so the footnote body a
/// caller renders can never disagree with the superscript mark this walk
/// emitted.
///
/// This replaced a design where the counter was advanced by this **recursive**
/// walk (which descends into hyperlinks, non-substituted fields, MCE
/// fallbacks, and VML text boxes) while callers re-derived the reference list
/// from a **flat** scan of the paragraph's top-level inlines. The two walkers
/// disagreed for any nested reference: its body was never emitted, and when it
/// preceded a top-level reference the body's number no longer matched the mark.
/// `Clone` exists for §17.6.22 speculative builds: measuring a *following*
/// section's header to learn its clearance walks that header's content, and
/// that walk records footnote references like any other. The peek snapshots
/// this tracker and restores it, so those references never reach the document.
/// See `BuildState::speculatively`.
#[derive(Clone, Default, Debug)]
pub struct FootnoteTracker {
    /// Display number of the most recently emitted reference.
    last_display: u32,
    /// References recorded since the last [`FootnoteTracker::take_pending`].
    pending: Vec<RecordedFootnote>,
}

impl FootnoteTracker {
    /// Assign the next display number to `id`, record it, and return the number.
    fn record(&mut self, id: NoteId) -> u32 {
        self.last_display += 1;
        self.pending.push(RecordedFootnote {
            id,
            display: self.last_display,
        });
        self.last_display
    }

    /// Take the references recorded since the last call, in encounter order.
    ///
    /// Every caller of [`collect_fragments`] must drain — including those that
    /// don't render footnote bodies (headers/footers, footnote bodies
    /// themselves). Otherwise their references leak into the next paragraph's
    /// batch and are rendered against the wrong host.
    pub fn take_pending(&mut self) -> Vec<RecordedFootnote> {
        std::mem::take(&mut self.pending)
    }
}

/// §17.3.2.4: convert a run-level [`crate::model::Border`] into a render-side
/// [`FragmentBorder`], filtering out the spec's "no border" styles.
///
/// `<w:bdr w:val="nil"/>` and `<w:bdr w:val="none"/>` (§17.18.2 ST_Border) are
/// kept apart by the model, because table border conflict resolution needs the
/// difference — but a run border has no adjacent edge to conflict with, so here
/// they are genuinely equivalent and `draws_nothing` covers both. The model
/// preserves the explicit `Some` so it can override an inherited border in the
/// §17.7.2 cascade; at the render boundary we drop it, otherwise the painter
/// would draw a hairline box around every word.
pub(super) fn run_border_to_fragment(
    border: Option<&crate::model::Border>,
) -> Option<FragmentBorder> {
    let b = border?;
    if b.style.draws_nothing() {
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
/// the design in `fragment/segment.rs`.
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
    auto_fit: crate::render::layout::ShapeAutoFit,
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

    let mut font = font_props_from_run(&effective_props, default_family, default_size, auto_fit);
    let color = effective_props
        .color
        .map(|c| {
            crate::render::resolve::color::resolve_color(
                c,
                crate::render::resolve::color::ColorContext::Text,
            )
        })
        .cloned()
        .unwrap_or(default_color);
    // §17.3.2.32 / §17.3.2.15: shading or highlight as background.
    let shading = effective_props
        .shading
        .get()
        .map(|s| {
            crate::render::resolve::color::resolve_color(
                s.fill,
                crate::render::resolve::color::ColorContext::Background,
            )
        })
        // §17.18.40: HighlightColor::None is the explicit "no highlight"
        // override and yields no fill, so use `and_then` to thread the
        // Option through.
        .or_else(|| {
            effective_props
                .highlight
                .cloned()
                .and_then(resolve_highlight_color)
        });

    // §17.3.2.42: vertical alignment (super/sub).
    let mut baseline_offset = match effective_props.vertical_align.cloned() {
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
    if let Some(pos) = effective_props.position.cloned() {
        baseline_offset += Pt::from(pos);
    }

    // §17.3.2.4: run-level border (filtered to drop the no-border styles).
    let border = run_border_to_fragment(effective_props.border.get());

    let text_style = TextRunStyle {
        color,
        shading,
        border,
        baseline_offset,
    };
    (font, text_style)
}

/// §17.16.4.1: context for evaluating dynamic fields (PAGE, NUMPAGES).
///
/// The layout-local subset, not the full field evaluator's context —
/// [`crate::field::FieldContext`] is the larger struct `crate::field::eval`
/// uses (different field names, `u32` rather than `usize`, mail-merge/date
/// /bookmark data this one doesn't carry).
#[derive(Clone, Copy, Default)]
pub struct FieldContext {
    /// Current page number (1-based).
    pub page_number: Option<usize>,
    /// Total page count in the document.
    pub num_pages: Option<usize>,
    /// §17.16.5.13: the moment this render started, in UTC — the same value
    /// for every `DATE` field in the document, so a render that spans
    /// midnight doesn't date two pages differently. Seeded once in
    /// `render::layout_document` (see `crate::field::now`); `None` in
    /// contexts that never evaluate a date field, where the field keeps its
    /// cached text.
    pub date: Option<crate::field::context::Date>,
    /// §17.16.5.76: `TIME`'s half of the same instant.
    pub time: Option<crate::field::context::Time>,
}

/// §17.16.4.1: evaluate a parsed field instruction against the current context.
/// Returns the substituted text, or `None` for a field this evaluator does not
/// compute — which leaves the field's cached `content` in place.
///
/// `locale_tag` is the §17.3.2.20 `w:lang` of the paragraph the field sits in;
/// only the §17.16.4.2 date pictures read it, for their month and weekday
/// names.
///
/// `crate::field::eval` has a fuller evaluator covering most
/// `FieldInstruction` variants (TOC, HYPERLINK, REF, SEQ, ...), but it isn't
/// wired into layout — this is the only field evaluation that happens during
/// rendering, and anything it doesn't answer falls through to the cached
/// `content`. The general `\*`/`\#` switches that evaluator applies are not
/// applied here for any field, DATE included: PAGE and NUMPAGES have always
/// ignored them on this path, and quietly honouring them for one field type
/// would make the same picture behave differently depending on which
/// evaluator ran.
fn evaluate_field_instruction(
    instruction: &crate::field::FieldInstruction,
    ctx: FieldContext,
    locale_tag: Option<&str>,
) -> Option<String> {
    use crate::field::FieldInstruction;
    match instruction {
        FieldInstruction::Page { .. } => ctx.page_number.map(|n| n.to_string()),
        FieldInstruction::NumPages { .. } => ctx.num_pages.map(|n| n.to_string()),
        // §17.16.4.2: `switches.date_format` is the `\@ "picture"` argument,
        // already extracted during parse. The defaults are Word's for a
        // picture-less field.
        // §17.16.4.2 is one picture grammar, so both field types render
        // through the same function and each may use the other's tokens; the
        // `?` still gates on the source the field is *named* for.
        FieldInstruction::Date { switches, .. } => {
            let date = ctx.date.as_ref()?;
            Some(match switches.date_format.as_deref() {
                Some(picture) => crate::field::format::format_datetime(
                    Some(date),
                    ctx.time.as_ref(),
                    picture,
                    locale_tag,
                ),
                // §17.16.5.13: no picture is a different question from an
                // empty one — the locale's own short date, not a hardcoded
                // American one. See `format::default_date`.
                None => crate::field::format::default_date(date, locale_tag),
            })
        }
        FieldInstruction::Time { switches, .. } => {
            let time = ctx.time.as_ref()?;
            Some(match switches.date_format.as_deref() {
                Some(picture) => crate::field::format::format_datetime(
                    ctx.date.as_ref(),
                    Some(time),
                    picture,
                    locale_tag,
                ),
                None => crate::field::format::default_time(time, locale_tag),
            })
        }
        _ => None,
    }
}

/// §17.16.19 MERGEFORMAT — source of formatting for a complex field's
/// substituted dynamic value. Resolved when the `Separate` fldChar is
/// reached so the lookup honors the OOXML "first result run wins" rule
/// regardless of how the inline-units pre-pass packaged the result zone:
/// an empty `<w:t></w:t>` placeholder run carries `<w:rPr>` but does not
/// surface as its own unit (segment joining drops it as 0 chars), yet
/// is still the spec's first-result-run for formatting purposes.
#[derive(Clone, Copy)]
pub(super) enum FieldFormatSource<'a> {
    /// First TextRun encountered between `Separate` and the matching
    /// `End` at the outer field's nesting level. Its `<w:rPr>` provides
    /// font family, size, bold, italic, color per §17.16.19.
    FirstResultRun(&'a TextRun),
    /// No result TextRun is present at the outer level. The
    /// substitution falls back to paragraph default font properties at
    /// emission time.
    ParagraphDefaults,
}

/// Locate the formatting source for the complex field whose `Separate`
/// fldChar sits at `inlines[separate_idx]`. Walks raw inlines (not
/// unit-packaged), tracking nesting via `Begin` / `End` counts, and
/// returns at the first top-level `TextRun` or at the matching `End`,
/// whichever comes first.
///
/// "Top level" = the depth of the field whose `Separate` triggered the
/// lookup. Text runs that sit inside a nested field's own result zone
/// belong to that nested field's substitution and are skipped — they
/// are not the outer field's first result run.
///
/// Malformed input (no matching `End`) returns `ParagraphDefaults`
/// rather than panicking.
pub(super) fn resolve_field_format_source(
    inlines: &[Inline],
    separate_idx: usize,
) -> FieldFormatSource<'_> {
    let mut depth: i32 = 0;
    for inline in &inlines[separate_idx + 1..] {
        match inline {
            Inline::FieldChar(fc) => match fc.field_char_type {
                FieldCharType::Begin => depth += 1,
                FieldCharType::End => {
                    if depth == 0 {
                        return FieldFormatSource::ParagraphDefaults;
                    }
                    depth -= 1;
                }
                FieldCharType::Separate => {
                    // Belongs to a nested field; the outer scan ignores it.
                }
            },
            Inline::TextRun(tr) if depth == 0 => {
                return FieldFormatSource::FirstResultRun(tr.as_ref());
            }
            _ => {}
        }
    }
    FieldFormatSource::ParagraphDefaults
}

/// Emit the substituted text of a complex field using the formatting
/// resolved at `Separate` (§17.16.19). When the source is
/// [`FieldFormatSource::FirstResultRun`] the substitution inherits font
/// family, size, bold, italic, color, etc. from that run's `<w:rPr>` —
/// matching what Word renders when it updates a dynamic field in place.
/// When no result run was present in the field zone the substitution
/// falls back to paragraph defaults.
#[allow(clippy::too_many_arguments)]
fn emit_field_substitution<F>(
    text: &str,
    source: Option<&FieldFormatSource<'_>>,
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
    auto_fit: crate::render::layout::ShapeAutoFit,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    measurer: Option<&crate::render::layout::measurer::TextMeasurer<'_>>,
    fragments: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let (font, text_style) = match source {
        Some(FieldFormatSource::FirstResultRun(tr)) => resolve_run_styling(
            tr,
            default_family,
            default_size,
            default_color,
            resolved_styles,
            paragraph_run_defaults,
            theme,
            auto_fit,
            measure_text,
        ),
        _ => (
            FontProps {
                effects: Default::default(),
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from(default_family),
                size: auto_fit.scale_font(default_size),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
            },
            TextRunStyle {
                color: default_color,
                shading: None,
                border: None,
                baseline_offset: Pt::ZERO,
            },
        ),
    };
    emit_text_fragments(
        text,
        &font,
        &text_style,
        hyperlink_url,
        measure_text,
        measurer,
        fragments,
    );
}

/// Build a text fragment for a substituted **simple field**
/// (`w:fldSimple`) value, using the paragraph's default font properties.
///
/// Not the complex-field path: a complex field's MERGEFORMAT substitution
/// goes through [`emit_field_substitution`] instead, which prefers the
/// field's own first result run over paragraph defaults.
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
        effects: Default::default(),
        rtl: crate::render::fonts::Toggle::Absent,
        family: Rc::from(default_family),
        size: default_size,
        bold: Toggle::Absent,
        italic: Toggle::Absent,
        underline: false,
        char_spacing: Pt::ZERO,
        text_scale: 1.0,
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    };
    let (w, m) = measure_text(&text, &font);
    Fragment::Text {
        shaped: None,
        level: BidiLevel::LTR,
        text,
        font: Rc::new(font),
        color: default_color,
        shading: None,
        border: None,
        // A substituted field value is one unbreakable unit — unchanged from
        // before UAX #14 arrived, and adequate for what these fields produce
        // (a page number, a date). A long result that has to wrap still can:
        // `split_oversized_fragments` cuts it once it exceeds the line.
        break_after: BreakAfter::Prohibited,
        width: w,
        trimmed_width: w,
        metrics: m,
        hyperlink_url: None,
        baseline_offset: Pt::ZERO,
        text_offset: Pt::ZERO,
        is_footnote_ref: false,
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
    /// §20.1.2.1.18: the `a:normAutofit` shrink of the enclosing shape text
    /// body. [`ShapeAutoFit::NONE`] everywhere else.
    ///
    /// [`ShapeAutoFit::NONE`]: crate::render::layout::ShapeAutoFit::NONE
    pub auto_fit: crate::render::layout::ShapeAutoFit,
    /// §17.3.2.20: the `w:lang` tag in effect for this paragraph, resolved
    /// from the same cascade `build::convert::paragraph_locale` reads (see
    /// `build::convert::resolve_lang_tag`). Carried as the raw tag rather
    /// than a `Locale` because its one consumer here — §17.16.4.2 date
    /// pictures — needs the region a `Locale` bucket discards.
    pub locale_tag: Option<&'a str>,
}

/// §17.3.2 `w:vanish`: whether this run is hidden text, and so contributes
/// nothing to the page.
///
/// Walks the §17.7.2 cascade in the order [`resolve_run_styling`] merges it —
/// direct `w:rPr`, then the character style named by `w:rStyle`, then the
/// paragraph style's run properties — but reads the one field rather than
/// merging the whole bag, because this is asked of every run in every document
/// and that merge clones a `RunProperties`. The two orders have to agree;
/// `a_run_can_override_a_hidden_style_back_to_visible` is what notices if they
/// drift apart.
///
/// Reads the *value*, not the presence: `<w:vanish w:val="0"/>` is how a run
/// un-hides itself inside a hidden style, and is the reason this cannot be an
/// `is_some()` test.
fn run_is_hidden(tr: &TextRun, ctx: &FragmentCtx<'_>) -> bool {
    if let Some(vanish) = tr.properties.vanish {
        return vanish;
    }
    if let (Some(style_id), Some(styles)) = (&tr.style_id, ctx.resolved_styles) {
        if let Some(vanish) = styles.get(style_id).and_then(|s| s.run.vanish) {
            return vanish;
        }
    }
    ctx.paragraph_run_defaults
        .and_then(|defaults| defaults.vanish)
        .unwrap_or(false)
}

/// Walk inline content and collect fragments.
/// `measure_text` is a callback that measures text width/height/ascent for a given font.
/// `resolved_styles` is used to look up character styles (w:rStyle) on text runs.
///
/// Returns fragments suitable for the line-fitting algorithm.
pub fn collect_fragments<F>(
    inlines: &[Inline],
    ctx: &FragmentCtx<'_>,
    hyperlink_url: Option<&LinkTarget>,
    measure_text: &F,
    footnotes: &mut FootnoteTracker,
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
    let auto_fit = ctx.auto_fit;

    // §17.3.2 `w:vanish`: hidden runs are removed here, before anything
    // measures or joins them, so the visible text either side closes up rather
    // than leaving a gap where the run was. Everything downstream — the field
    // pre-pass below, `build_inline_units`, the line fitter — then sees a
    // stream that simply does not contain them.
    //
    // The scan runs first so the allocation happens only for a document that
    // actually hides something, which is a small minority; for every other one
    // this costs two `Option` checks per run and no clone. A `Vec<&Inline>`
    // would avoid the clone in that minority too, at the price of threading a
    // borrow through `build_inline_units` and the field pre-pass — not worth it
    // for a path taken this rarely.
    let hidden = |inline: &Inline| matches!(inline, Inline::TextRun(tr) if run_is_hidden(tr, ctx));
    let visible: Option<Vec<Inline>> = inlines
        .iter()
        .any(hidden)
        .then(|| inlines.iter().filter(|i| !hidden(i)).cloned().collect());
    let inlines: &[Inline] = visible.as_deref().unwrap_or(inlines);

    let mut fragments = Vec::new();
    let mut field_depth: i32 = 0; // tracks nested complex field state
    let mut field_instr = String::new(); // accumulated instruction text for current complex field
                                         // §17.16.19: field substitution state for complex fields.
                                         // Pending = substitution text waiting for the first result TextRun's formatting.
                                         // Emitted = substitution was rendered, skip remaining result TextRuns until End.
    let mut field_sub_pending: Option<String> = None;
    let mut field_sub_emitted = false;

    // §17.16.19 MERGEFORMAT — pre-resolve formatting for each complex
    // field's substitution against raw inlines, so empty placeholder
    // result runs (`<w:t></w:t>` — swallowed by `build_inline_units`
    // because they contribute 0 chars) still surface their `<w:rPr>`.
    // One entry per `Separate` fldChar, consumed in order.
    let field_format_sources: Vec<FieldFormatSource<'_>> = inlines
        .iter()
        .enumerate()
        .filter_map(|(idx, inl)| match inl {
            Inline::FieldChar(fc) if matches!(fc.field_char_type, FieldCharType::Separate) => {
                Some(resolve_field_format_source(inlines, idx))
            }
            _ => None,
        })
        .collect();
    let mut field_format_idx: usize = 0;
    let mut current_field_format: Option<FieldFormatSource<'_>> = None;
    // Pre-pass: join consecutive text-only TextRuns into segments so
    // UAX #29 grapheme clusters reassemble across `<w:rFonts>`-induced
    // run splits (keycap `1️⃣`, ZWJ family, modifier sequence, …).
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
                        auto_fit,
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
                //
                // `classify` divides a run's text at every UAX #14 boundary,
                // so consecutive pieces overwhelmingly come from the *same*
                // `<w:r>`: resolving its styling once per piece would walk the
                // §17.7.2 cascade once per word instead of once per run, which
                // measured as +19% on the layout phase of `sample4.docx`. The
                // resolved styling is memoized against the run's identity —
                // `classify` never interleaves runs, so a one-entry memo
                // catches every repeat.
                let mut run_styling: Option<(&TextRun, Rc<FontProps>, TextRunStyle)> = None;
                for piece in seg.classify() {
                    match piece {
                        SegmentPiece::Text {
                            run,
                            text,
                            break_after,
                        } => {
                            let (font, text_style) = match &run_styling {
                                Some((cached, font, style)) if std::ptr::eq(*cached, run) => {
                                    (font, style)
                                }
                                _ => {
                                    let (font, style) = resolve_run_styling(
                                        run,
                                        default_family,
                                        default_size,
                                        default_color,
                                        resolved_styles,
                                        paragraph_run_defaults,
                                        theme,
                                        auto_fit,
                                        measure_text,
                                    );
                                    let (_, font, style) =
                                        run_styling.insert((run, Rc::new(font), style));
                                    (&*font, &*style)
                                }
                            };
                            // Pre-classified *and* pre-segmented: `classify`
                            // has already applied UAX #29, UTS #51 and UAX #14
                            // across the whole joined text, so this piece is
                            // one break unit and re-segmenting it would cost a
                            // second pass to find nothing.
                            emit_text_unit(
                                &text,
                                break_after,
                                font,
                                text_style,
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
                                auto_fit,
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
                        auto_fit,
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
                                        // §17.3.1.38: a leader on this tab is
                                        // drawn in the tab run's own formatting.
                                        font: Rc::new(font.clone()),
                                        color: text_style.color,
                                        fitting_width: None,
                                    });
                                }
                                RunElement::PositionTab(ptab) => {
                                    fragments.push(Fragment::PTab {
                                        align: ptab.alignment,
                                        relative_to: ptab.relative_to,
                                        leader: ptab.leader.into(),
                                        line_height: font.size,
                                        font: Rc::new(font.clone()),
                                        color: text_style.color,
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
                                src_rect: crate::render::resolve::images::extract_src_rect(img),
                            });
                        }
                    }
                }
                Inline::Hyperlink(link) => {
                    // Preserve the external/internal kind as a closed ADT so
                    // the emitter routes external→URI and internal→GoTo without
                    // guessing from the string (§17.16.22).
                    // One allocation per `w:hyperlink`, shared by every word
                    // fragment inside it and every command they emit.
                    let target: Option<LinkTarget> = match &link.target {
                        crate::model::HyperlinkTarget::ExternalUrl(url) => {
                            Some(LinkTarget::External(Rc::from(url.as_str())))
                        }
                        crate::model::HyperlinkTarget::Internal { anchor } => {
                            Some(LinkTarget::Internal(Rc::from(anchor.as_str())))
                        }
                        // An unresolved rId (no matching relationship) has no link.
                        crate::model::HyperlinkTarget::ExternalRel(_) => None,
                    };
                    let mut sub = collect_fragments(
                        &link.content,
                        ctx,
                        target.as_ref(),
                        measure_text,
                        footnotes,
                        endnote_counter,
                        field_ctx,
                    );
                    fragments.append(&mut sub);
                }
                Inline::Field(field) => {
                    // §17.16.18: simple field — check for dynamic substitution.
                    let substituted =
                        evaluate_field_instruction(&field.instruction, field_ctx, ctx.locale_tag);
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
                            footnotes,
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
                            // §17.16.4.1: parse accumulated instruction, then
                            // evaluate it if this context can (PAGE/NUMPAGES,
                            // DATE/TIME).
                            if let Ok(parsed) = crate::field::parse(&field_instr) {
                                field_sub_pending =
                                    evaluate_field_instruction(&parsed, field_ctx, ctx.locale_tag);
                            }
                            // §17.16.19: bind the formatting source resolved
                            // against raw inlines, so the End fallback path
                            // can recover an empty placeholder run's rPr
                            // even though it was dropped by segment joining.
                            current_field_format =
                                field_format_sources.get(field_format_idx).copied();
                            field_format_idx += 1;
                            field_depth -= 1; // now collect result runs (unless substituted)
                        }
                        FieldCharType::End => {
                            // Substitution still pending at End: the unit
                            // stream never carried a result run (either the
                            // placeholder was empty and got swallowed by
                            // segment joining, or the field has no result
                            // content at all). Use the pre-resolved format
                            // source — §17.16.19 first-result-run when
                            // present, paragraph defaults otherwise.
                            if let Some(text) = field_sub_pending.take() {
                                emit_field_substitution(
                                    &text,
                                    current_field_format.as_ref(),
                                    default_family,
                                    default_size,
                                    default_color,
                                    resolved_styles,
                                    paragraph_run_defaults,
                                    theme,
                                    auto_fit,
                                    hyperlink_url,
                                    measure_text,
                                    ctx.measurer,
                                    &mut fragments,
                                );
                            }
                            current_field_format = None;
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
                    use crate::render::layout::{live_mc_branch, McBranch};
                    // §M.1.2 / §17.17.1: only the live branch contributes, and
                    // `live_mc_branch` is the one place that decides which.
                    //
                    // A drawable Choice is drawn as float geometry by the
                    // floating extractor — and for a wps shape, its `txbx`
                    // contents are laid out into shape-local commands emitted
                    // on top of the shape's path. Walking the Fallback here as
                    // well would duplicate that text into the host paragraph at
                    // the wrong y.
                    //
                    // A live Fallback does come through here, and must: its VML
                    // geometry goes to the float walkers, its text box to this
                    // collector, landing at the host paragraph y as the Tier 0
                    // placeholder it has always been.
                    match live_mc_branch(ac) {
                        McBranch::Fallback(fallback) => {
                            let mut sub = collect_fragments(
                                fallback,
                                ctx,
                                hyperlink_url,
                                measure_text,
                                footnotes,
                                endnote_counter,
                                field_ctx,
                            );
                            fragments.append(&mut sub);
                        }
                        McBranch::Choices(_) | McBranch::Neither => {}
                    }
                }
                Inline::Symbol(sym) => {
                    let font = FontProps {
                        effects: Default::default(),
                        rtl: crate::render::fonts::Toggle::Absent,
                        family: Rc::from(sym.font.as_str()),
                        size: default_size,
                        bold: Toggle::Absent,
                        italic: Toggle::Absent,
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
                        shaped: None,
                        level: BidiLevel::LTR,
                        text: Rc::from(text.as_str()),
                        font: Rc::new(font),
                        color: RgbColor::BLACK,
                        shading: None,
                        border: None,
                        // §17.3.3.30: one glyph from a symbol font, whose
                        // code point carries the symbol font's meaning and
                        // not Unicode's — so UAX #14 has nothing to say
                        // about it. Joined to what follows, as before.
                        break_after: BreakAfter::Prohibited,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: hyperlink_url.cloned(),
                        baseline_offset: Pt::ZERO,
                        text_offset: Pt::ZERO,
                        is_footnote_ref: false,
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
                Inline::FootnoteRef(note_id) => {
                    let num_text = format!("{}", footnotes.record(*note_id));
                    // §17.11.12: footnote reference uses superscript sizing.
                    let ref_size = default_size * super::SUPERSCRIPT_FONT_SIZE_RATIO;
                    let ref_font = FontProps {
                        effects: Default::default(),
                        rtl: crate::render::fonts::Toggle::Absent,
                        family: std::rc::Rc::from(default_family),
                        size: ref_size,
                        bold: Toggle::Absent,
                        italic: Toggle::Absent,
                        underline: false,
                        char_spacing: Pt::ZERO,
                        text_scale: 1.0,
                        underline_position: Pt::ZERO,
                        underline_thickness: Pt::ZERO,
                    };
                    let (w, m) = measure_text(&num_text, &ref_font);
                    // Raise the mark clear of the baseline (see the constant).
                    let baseline_offset = -(default_size * super::NOTE_REF_BASELINE_OFFSET_RATIO);
                    fragments.push(Fragment::Text {
                        shaped: None,
                        level: BidiLevel::LTR,
                        text: Rc::from(num_text.as_str()),
                        font: Rc::new(ref_font),
                        color: default_color,
                        shading: None,
                        border: None,
                        // §17.11.12: a reference mark belongs to the word it
                        // follows and must not be stranded on the next line.
                        break_after: BreakAfter::Prohibited,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: None,
                        baseline_offset,
                        text_offset: Pt::ZERO,
                        // §17.11.12: tag so across-page splitting reserves this
                        // footnote on the page its reference mark lands on.
                        is_footnote_ref: true,
                    });
                }
                // §17.11.2: endnote reference — render as superscript Roman numeral.
                Inline::EndnoteRef(_note_id) => {
                    *endnote_counter += 1;
                    let num_text = to_roman_lower(*endnote_counter);
                    let ref_size = default_size * super::SUPERSCRIPT_FONT_SIZE_RATIO;
                    let ref_font = FontProps {
                        effects: Default::default(),
                        rtl: crate::render::fonts::Toggle::Absent,
                        family: std::rc::Rc::from(default_family),
                        size: ref_size,
                        bold: Toggle::Absent,
                        italic: Toggle::Absent,
                        underline: false,
                        char_spacing: Pt::ZERO,
                        text_scale: 1.0,
                        underline_position: Pt::ZERO,
                        underline_thickness: Pt::ZERO,
                    };
                    let (w, m) = measure_text(&num_text, &ref_font);
                    let baseline_offset = -(default_size * super::NOTE_REF_BASELINE_OFFSET_RATIO);
                    fragments.push(Fragment::Text {
                        shaped: None,
                        level: BidiLevel::LTR,
                        text: Rc::from(num_text.as_str()),
                        font: Rc::new(ref_font),
                        color: default_color,
                        shading: None,
                        border: None,
                        // §17.11.12: a reference mark belongs to the word it
                        // follows and must not be stranded on the next line.
                        break_after: BreakAfter::Prohibited,
                        width: w,
                        trimmed_width: w,
                        metrics: m,
                        hyperlink_url: None,
                        baseline_offset,
                        text_offset: Pt::ZERO,
                        is_footnote_ref: false,
                    });
                }
                Inline::Pict(pict) => {
                    // Render text content from VML text-box-bearing
                    // primitives inline. Every primitive variant
                    // (`<v:shape>`, `<v:rect>`, `<v:roundrect>`,
                    // `<v:oval>`, …) admits a `<v:textbox>` child via
                    // `VmlCommonAttrs.text_box`; the previous code
                    // only walked the `Shape` variant and silently
                    // dropped text from rect / roundrect / oval text
                    // boxes (the case footer3.xml of the vorlage doc
                    // exercised — the gray bar is a `<v:rect>`).
                    //
                    // Does not handle absolute positioning — text
                    // appears inline with the surrounding paragraph.
                    for primitive in &pict.primitives {
                        let common = primitive.common();
                        if let Some(ref text_box) = common.text_box {
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
                                        auto_fit: ctx.auto_fit,
                                        // A VML text box carries no language
                                        // of its own; it reads as part of the
                                        // paragraph that anchors it.
                                        locale_tag: ctx.locale_tag,
                                    };
                                    let mut sub = collect_fragments(
                                        &p.content,
                                        &pict_ctx,
                                        hyperlink_url,
                                        measure_text,
                                        footnotes,
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
    use crate::render::fonts::Toggle;

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
            auto_fit: crate::render::layout::ShapeAutoFit::NONE,
            locale_tag: None,
        }
    }

    // ── §17.3.2 `w:vanish` hidden text ───────────────────────────────────
    //
    // A run marked hidden contributes nothing to the page: no glyphs, no
    // width, and no line height of its own. The surrounding text closes up
    // around it, so hiding is a *removal* from the inline stream rather than a
    // draw-time skip — that is the difference between "Visible tail" and
    // "Visible  tail".
    //
    // The cascade is §17.7.2's, the same one `resolve_run_styling` walks:
    // direct `w:rPr` beats the character style (`w:rStyle`), which beats the
    // paragraph style's run properties and `docDefaults`. `w:val="0"` on a
    // run is therefore how a document un-hides one run of a hidden style, and
    // it has to keep working — a `.is_some()` test in place of a value test
    // would silently invert it.
    //
    // Not covered, deliberately, and stated where it bites rather than only
    // here: the paragraph **mark**'s own `w:pPr/w:rPr/w:vanish`. That hides the
    // mark, which merges the paragraph into the next one — a pagination change,
    // not a run filter — and the mark's properties never reach this cascade
    // anyway (`resolve_paragraph_defaults` builds `paragraph_run_defaults` from
    // the paragraph *style*, not from `w:pPr/w:rPr`).

    fn hidden_run(text: &str, vanish: Option<bool>, style_id: Option<&str>) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: style_id.map(crate::model::StyleId::new),
            properties: RunProperties {
                vanish,
                ..Default::default()
            },
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    /// A character style that hides whatever it is applied to.
    fn styles_with_hidden(
        id: &str,
    ) -> std::collections::HashMap<
        crate::model::StyleId,
        crate::render::resolve::styles::ResolvedStyle,
    > {
        let style = crate::render::resolve::styles::ResolvedStyle {
            paragraph: Default::default(),
            run: RunProperties {
                vanish: Some(true),
                ..Default::default()
            },
            table: None,
            table_style_overrides: Vec::new(),
            is_toc_entry: false,
        };
        let mut map = std::collections::HashMap::new();
        map.insert(crate::model::StyleId::new(id), style);
        map
    }

    fn collect(inlines: &[Inline], ctx: &FragmentCtx<'_>) -> Vec<Fragment> {
        collect_fragments(
            inlines,
            ctx,
            None,
            &dummy_measure,
            &mut FootnoteTracker::default(),
            &mut 0,
            FieldContext::default(),
        )
    }

    /// The defect: a hidden run drew its text like any other.
    #[test]
    fn a_hidden_run_contributes_no_fragments() {
        let inlines = vec![hidden_run("SECRET", Some(true), None)];
        assert!(
            collect(&inlines, &default_ctx(12.0)).is_empty(),
            "a hidden run must not reach layout at all"
        );
    }

    /// Removal, not a zero-width draw: the visible text either side has to end
    /// up adjacent, which is only true if the hidden run never becomes a
    /// fragment.
    #[test]
    fn text_closes_up_around_a_hidden_run() {
        let inlines = vec![
            text_run("ab"),
            hidden_run("SECRET", Some(true), None),
            text_run("cd"),
        ];
        let frags = collect(&inlines, &default_ctx(12.0));
        let width: f32 = frags.iter().map(|f| f.width().raw()).sum();
        assert_eq!(
            width, 24.0,
            "only the four visible characters may take width, got {frags:?}"
        );
    }

    /// §17.7.2: a character style can hide the runs that reference it.
    #[test]
    fn a_character_style_hides_the_runs_that_use_it() {
        let styles = styles_with_hidden("Secret");
        let ctx = FragmentCtx {
            resolved_styles: Some(&styles),
            ..default_ctx(12.0)
        };
        let inlines = vec![hidden_run("SECRET", None, Some("Secret"))];
        assert!(
            collect(&inlines, &ctx).is_empty(),
            "the style's w:vanish must reach the run"
        );
    }

    /// §17.7.2: and the run can turn it back off. `w:vanish w:val="0"` is the
    /// documented way to un-hide one run of a hidden style, so the filter has
    /// to read the *value*, not merely the presence of the element.
    #[test]
    fn a_run_can_override_a_hidden_style_back_to_visible() {
        let styles = styles_with_hidden("Secret");
        let ctx = FragmentCtx {
            resolved_styles: Some(&styles),
            ..default_ctx(12.0)
        };
        let inlines = vec![hidden_run("ab", Some(false), Some("Secret"))];
        assert_eq!(
            collect(&inlines, &ctx).len(),
            1,
            "explicit w:val=\"0\" outranks the style"
        );
    }

    /// §17.7.2: the paragraph style's run properties are the bottom of the
    /// cascade, and reach a run that states nothing itself.
    #[test]
    fn paragraph_run_defaults_can_hide_a_run() {
        let defaults = RunProperties {
            vanish: Some(true),
            ..Default::default()
        };
        let ctx = FragmentCtx {
            paragraph_run_defaults: Some(&defaults),
            ..default_ctx(12.0)
        };
        assert!(
            collect(&[text_run("SECRET")], &ctx).is_empty(),
            "an inherited w:vanish hides a run that states nothing"
        );
        let inlines = vec![hidden_run("ab", Some(false), None)];
        assert_eq!(
            collect(&inlines, &ctx).len(),
            1,
            "and the run still outranks it"
        );
    }

    /// A hidden run's tabs and breaks go with it — they are its content, and
    /// Word does not leave a tab stop or a line break behind when it hides the
    /// run that carried them.
    #[test]
    fn a_hidden_run_takes_its_tabs_and_breaks_with_it() {
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties {
                vanish: Some(true),
                ..Default::default()
            },
            content: vec![
                RunElement::Tab,
                RunElement::Text("SECRET".into()),
                RunElement::LineBreak(crate::model::BreakKind::TextWrapping),
            ],
            rsids: RevisionIds::default(),
        }))];
        assert!(
            collect(&inlines, &default_ctx(12.0)).is_empty(),
            "the whole run goes, not just its text"
        );
    }

    /// Hidden runs inside a `w:hyperlink` are hidden too — the recursion has to
    /// carry the filter, not just the top level.
    #[test]
    fn a_hidden_run_inside_a_hyperlink_is_hidden() {
        let inlines = vec![Inline::Hyperlink(crate::model::Hyperlink {
            target: crate::model::HyperlinkTarget::ExternalUrl("https://example.com".into()),
            content: vec![hidden_run("SECRET", Some(true), None), text_run("ab")],
        })];
        let frags = collect(&inlines, &default_ctx(12.0));
        let width: f32 = frags.iter().map(|f| f.width().raw()).sum();
        assert_eq!(
            width, 12.0,
            "only the visible half of the link, got {frags:?}"
        );
    }

    /// **Known limit, characterized rather than fixed.** Word hides a `w:sym`,
    /// `w:drawing` or `w:pict` in a hidden run along with its text. This engine
    /// cannot: `docx::parse::body::extend_from_run` flushes those children into
    /// sibling `Inline`s of their own, and `Inline::Symbol` / `Inline::Image` /
    /// `Inline::Pict` carry no run properties — so by the time the filter runs,
    /// the `w:vanish` that governed them is gone.
    ///
    /// Closing it is a model change (carry the run's `w:rPr` onto those
    /// inlines), not a change here. This test exists so that change announces
    /// itself instead of silently contradicting a passing suite.
    #[test]
    fn a_hidden_symbol_still_draws_because_the_model_drops_its_run_properties() {
        let inlines = vec![
            hidden_run("SECRET", Some(true), None),
            Inline::Symbol(crate::model::Symbol {
                font: "Wingdings".into(),
                char_code: 0xF0FC,
            }),
        ];
        assert_eq!(
            collect(&inlines, &default_ctx(12.0)).len(),
            1,
            "the symbol survives its hidden run — see this test's doc comment"
        );
    }

    // ── §17.3.2.4 / §17.18.2 run-level border tri-state ─────────────────
    //
    // The cascade may carry a child run whose `<w:bdr w:val="nil"/>`
    // (or "none") explicitly turns off an inherited border. The model
    // preserves this as `Some(Border { style: nil-or-none, .. })`
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
                font_size: Dup::from(Some(Dimension::<HalfPoints>::new(size))),
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        assert!(matches!(frags[0], Fragment::Tab { .. }));
    }

    #[test]
    fn position_tab_produces_ptab_fragment() {
        use crate::model::{PTabAlignment, PTabLeader, PTabRelativeTo, PositionTab};
        let inlines = vec![Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::PositionTab(PositionTab {
                alignment: PTabAlignment::Right,
                relative_to: PTabRelativeTo::Margin,
                leader: PTabLeader::Dot,
            })],
            rsids: RevisionIds::default(),
        }))];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut FootnoteTracker::default(),
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        assert!(matches!(
            frags[0],
            Fragment::PTab {
                align: PTabAlignment::Right,
                relative_to: PTabRelativeTo::Margin,
                leader: crate::model::TabLeader::Dot,
                ..
            }
        ));
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
            &mut FootnoteTracker::default(),
            &mut 0,
            FieldContext::default(),
        );

        assert_eq!(frags.len(), 1);
        assert!(frags[0].is_line_break());
    }

    #[test]
    fn hyperlink_recurses_into_content() {
        let inlines = vec![Inline::Hyperlink(Hyperlink {
            target: HyperlinkTarget::ExternalUrl("https://example.com".into()),
            content: vec![text_run("click me")],
        })];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut FootnoteTracker::default(),
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
            assert_eq!(
                hyperlink_url,
                &Some(LinkTarget::External("https://example.com".into()))
            );
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
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
                requires: vec![McRequires::Wps],
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
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
            &mut FootnoteTracker::default(),
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

    // ── §17.16.19 MERGEFORMAT — field result format source ───────────────
    //
    // `resolve_field_format_source` walks raw inlines forward from a
    // `Separate` fldChar to locate the formatting that should be applied
    // to a substituted dynamic value (PAGE/NUMPAGES/...). Decoupling
    // this from `build_inline_units` is the key correctness property:
    // an empty `<w:t></w:t>` placeholder result run carries `<w:rPr>`
    // but is swallowed by segment joining (it contributes 0 chars), so
    // we cannot rely on units to surface it.

    fn fld_char(kind: FieldCharType) -> Inline {
        Inline::FieldChar(FieldChar {
            field_char_type: kind,
            dirty: None,
            fld_lock: None,
        })
    }

    fn bold_text_run(text: &str) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties {
                bold: Some(true),
                ..Default::default()
            },
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    /// Helper: pull `&TextRun` out of `FieldFormatSource::FirstResultRun`,
    /// or panic with a descriptive message.
    fn expect_first_run<'a>(src: FieldFormatSource<'a>) -> &'a TextRun {
        match src {
            FieldFormatSource::FirstResultRun(tr) => tr,
            FieldFormatSource::ParagraphDefaults => {
                panic!("expected FirstResultRun, got ParagraphDefaults")
            }
        }
    }

    /// Extract the concatenated text from a TextRun's content. Lets the
    /// assertions read like prose without depending on `PartialEq` for
    /// the `RunElement` ADT.
    fn run_text(tr: &TextRun) -> String {
        tr.content
            .iter()
            .filter_map(|e| match e {
                RunElement::Text(t) => Some(t.as_str()),
                _ => None,
            })
            .collect()
    }

    /// Canonical complex field shape with a non-empty result run.
    /// Inline layout: `[Begin, InstrText, Separate, TextRun("3"), End]`.
    #[test]
    fn format_source_finds_text_run_after_separate() {
        let inlines = vec![
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate),
            bold_text_run("3"),
            fld_char(FieldCharType::End),
        ];
        let src = resolve_field_format_source(&inlines, 2);
        let tr = expect_first_run(src);
        assert_eq!(run_text(tr), "3");
        assert_eq!(tr.properties.bold, Some(true));
    }

    /// The original bug: an empty `<w:t></w:t>` placeholder result run
    /// must still be discoverable as the format source, because its
    /// `<w:rPr>` is the only place the substitution can find its
    /// formatting.
    #[test]
    fn format_source_finds_empty_placeholder_result_run() {
        let inlines = vec![
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate),
            bold_text_run(""), // empty placeholder, bold rPr
            fld_char(FieldCharType::End),
        ];
        let src = resolve_field_format_source(&inlines, 2);
        let tr = expect_first_run(src);
        assert!(run_text(tr).is_empty(), "expected empty text content");
        assert_eq!(
            tr.properties.bold,
            Some(true),
            "empty placeholder's bold rPr must be reachable"
        );
    }

    /// §17.16.19 — the FIRST result run wins when multiple are present.
    #[test]
    fn format_source_uses_first_when_multiple_result_runs() {
        let inlines = vec![
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate),
            bold_text_run("first"),
            text_run("second"), // not bold
            fld_char(FieldCharType::End),
        ];
        let src = resolve_field_format_source(&inlines, 2);
        let tr = expect_first_run(src);
        assert_eq!(run_text(tr), "first");
        assert_eq!(tr.properties.bold, Some(true));
    }

    /// `Separate` immediately followed by `End` — no result run exists,
    /// so the resolver returns `ParagraphDefaults` and the substitution
    /// will fall back to paragraph defaults at emission time.
    #[test]
    fn format_source_returns_defaults_for_empty_result_zone() {
        let inlines = vec![
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate),
            fld_char(FieldCharType::End),
        ];
        let src = resolve_field_format_source(&inlines, 2);
        assert!(matches!(src, FieldFormatSource::ParagraphDefaults));
    }

    /// Text runs inside a NESTED field's own result zone belong to that
    /// nested field's substitution, not the outer one's. The outer
    /// resolver must skip them and look for runs at its own nesting
    /// level — here, `outer_run` after the nested End.
    #[test]
    fn format_source_skips_runs_inside_nested_field() {
        let inlines = vec![
            fld_char(FieldCharType::Begin), // outer
            Inline::InstrText("OUTER".into()),
            fld_char(FieldCharType::Separate), // index 2
            // ── nested field at depth 1 ──
            fld_char(FieldCharType::Begin),
            Inline::InstrText("INNER".into()),
            fld_char(FieldCharType::Separate),
            bold_text_run("inner-result"), // inside nested zone — skip
            fld_char(FieldCharType::End),
            // ── back at outer's level ──
            bold_text_run("outer-result"), // this is the one we want
            fld_char(FieldCharType::End),
        ];
        let src = resolve_field_format_source(&inlines, 2);
        let tr = expect_first_run(src);
        assert_eq!(
            run_text(tr),
            "outer-result",
            "must skip runs inside nested field"
        );
    }

    /// Content past the matching `End` belongs to a later field (or to
    /// the surrounding paragraph). The resolver stops at `End` and does
    /// not leak formatting from outside.
    #[test]
    fn format_source_stops_at_matching_end() {
        let inlines = vec![
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate), // index 2
            fld_char(FieldCharType::End),
            bold_text_run("trailing"), // outside the field — must not be picked up
        ];
        let src = resolve_field_format_source(&inlines, 2);
        assert!(
            matches!(src, FieldFormatSource::ParagraphDefaults),
            "trailing run after End must not become the source"
        );
    }

    /// Malformed inlines: no matching `End` after `Separate`. Treat as
    /// "no result" rather than panicking or returning a partial result.
    #[test]
    fn format_source_handles_missing_end_gracefully() {
        let inlines = vec![
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate), // index 2
                                               // (no End — malformed)
        ];
        let src = resolve_field_format_source(&inlines, 2);
        // No TextRun present and no End either — defaults is correct.
        assert!(matches!(src, FieldFormatSource::ParagraphDefaults));
    }

    // ── End-to-end: empty-placeholder result run formatting ──────────────
    //
    // Mirrors the Fotodokumentation Test header structure:
    //   `Seite ` [Begin → InstrText("PAGE") → Separate → <w:t></w:t> bold → End]
    // The substituted page number must inherit bold from the empty
    // placeholder result run's `<w:rPr>`.

    #[test]
    fn page_field_with_empty_placeholder_inherits_bold() {
        let inlines = vec![
            bold_text_run("Seite "),
            fld_char(FieldCharType::Begin),
            Inline::InstrText("PAGE".into()),
            fld_char(FieldCharType::Separate),
            bold_text_run(""), // placeholder result with bold rPr
            fld_char(FieldCharType::End),
        ];
        let ctx = default_ctx(12.0);
        let frags = collect_fragments(
            &inlines,
            &ctx,
            None,
            &dummy_measure,
            &mut FootnoteTracker::default(),
            &mut 0,
            FieldContext {
                page_number: Some(7),
                ..Default::default()
            },
        );
        // Two visible text fragments: "Seite " and the substituted "7".
        let texts: Vec<(&str, Toggle)> = frags
            .iter()
            .filter_map(|f| match f {
                Fragment::Text { text, font, .. } => Some((text.as_ref(), font.bold)),
                _ => None,
            })
            .collect();
        assert_eq!(
            texts,
            vec![("Seite ", Toggle::On), ("7", Toggle::On)],
            "PAGE substitution must inherit bold from empty placeholder run"
        );
    }

    // ── §17.6.22 speculative build support (plan §1) ─────────────────────────

    /// A speculative build (peeking at the next section's header to learn its
    /// clearance) snapshots the tracker and restores it afterwards, so the
    /// clone must be fully independent: both the display counter *and* the
    /// pending list. Sharing either would let a peeked reference surface as a
    /// real one, rendering a footnote body against the wrong host paragraph.
    #[test]
    fn clone_is_independent_in_counter_and_pending_list() {
        let mut original = FootnoteTracker::default();
        assert_eq!(original.record(NoteId::new(1)), 1);

        let snapshot = original.clone();

        assert_eq!(original.record(NoteId::new(2)), 2, "original advances");
        assert_eq!(original.record(NoteId::new(3)), 3);

        // The snapshot saw exactly one reference and still numbers from there.
        let mut restored = snapshot;
        assert_eq!(
            restored.take_pending().len(),
            1,
            "the clone keeps only the references recorded before it was taken"
        );
        assert_eq!(
            restored.record(NoteId::new(4)),
            2,
            "and resumes numbering where the snapshot was taken, not where the \
             speculative walk left off"
        );
    }
}
