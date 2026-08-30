use std::collections::HashMap;

use crate::i18n::bidi::BaseDirection;
use crate::model::{self, FirstLineIndent, LineSpacing};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::layout::build::BuildState;
use crate::render::layout::fragment::Fragment;
use crate::render::layout::measurer::TextMeasurer;
use crate::render::layout::paragraph::TabStopDef;
use crate::render::layout::paragraph::{
    BorderLine, LineSpacingRule, ParagraphBorderStyle, ParagraphStyle,
};
use crate::render::layout::table::{
    CellBorderOverride, TableBorderConfig, TableBorderLine, TableBorderStyle,
};
use crate::render::resolve::color::{resolve_color, ColorContext, RgbColor};
use crate::render::resolve::fonts::effective_font;
use crate::render::resolve::images::MediaEntry;
use crate::render::resolve::properties::merge_paragraph_properties;
use crate::render::resolve::ResolvedDocument;

use super::{BuildContext, SPEC_DEFAULT_FONT_SIZE, SPEC_FALLBACK_FONT};

/// Resolve a paragraph's effective defaults.
/// Cascade: direct → style → doc defaults.
///
/// Returns (font_family, font_size, color, merged_paragraph_props, run_defaults).
///
/// When `defer_doc_defaults` is true, doc defaults are NOT merged into the
/// paragraph properties — the caller is responsible for merging them after
/// inserting table style / conditional formatting in the cascade.
pub(super) fn resolve_paragraph_defaults(
    para: &model::Paragraph,
    resolved: &ResolvedDocument,
    defer_doc_defaults: bool,
    // §20.1.4.1.17: base color / font family for a shape text box (from
    // `wps:style/fontRef`). They sit *below* the style/run cascade — an
    // explicit run or style value still wins. `None` uses the normal defaults.
    base_color: Option<RgbColor>,
    base_family: Option<&str>,
) -> (
    String,
    Pt,
    RgbColor,
    model::ParagraphProperties,
    model::RunProperties,
) {
    let mut para_props = para.properties.clone();
    let mut run_defaults = resolved.doc_defaults_run.clone();

    // Derive default font from: fontRef base → doc defaults → theme minor → spec.
    let mut default_family = match base_family.filter(|s| !s.is_empty()) {
        Some(f) => f.to_string(),
        None => resolved
            .theme
            .as_ref()
            .map(|t| t.minor_font.latin.as_str())
            .filter(|s| !s.is_empty())
            .unwrap_or(SPEC_FALLBACK_FONT)
            .to_string(),
    };
    let mut default_size = resolved
        .doc_defaults_run
        .font_size
        .cloned()
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE);
    let mut default_color = base_color.unwrap_or(RgbColor::BLACK);

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
    if let Some(fs) = run_defaults.font_size.cloned() {
        default_size = Pt::from(fs);
    }
    if let Some(c) = run_defaults.color.cloned() {
        default_color = resolve_color(c, ColorContext::Text);
    }

    (
        default_family,
        default_size,
        default_color,
        para_props,
        run_defaults,
    )
}

/// §17.3.2.20 / §17.7.2: the language in effect for a paragraph's own layout
/// decisions.
///
/// The cascade is the paragraph **mark**'s `w:rPr`, then the paragraph style's
/// run defaults, then `docDefaults` — the same order §17.7.2 resolves any other
/// run property in, minus the runs themselves.
///
/// Leaving the runs out is deliberate, and is the one judgement call in
/// threading locale. The decisions this answers belong to the *paragraph*: a
/// §17.18.85 `decimal` stop is declared in `w:pPr/w:tabs`, and its zone may span
/// several runs that each declare a different `w:lang` — there is no single
/// run to ask. The paragraph mark is the nearest thing the file offers to
/// "the language of this paragraph", and it is already what the §17.9.22 list
/// label reads its formatting from, so the two agree.
///
/// Word itself is reported to take a decimal tab's separator from the host's
/// regional settings rather than from the document at all, which a converter
/// cannot reproduce and should not want to: the same file would render
/// differently on two machines. This reading is deterministic and keyed to the
/// document instead. **Word reference render**: this is an assumption, not a
/// verified match — if a Word reference render ever settles the question
/// differently, this cascade is the one place to change.
///
/// Returns both the coarse [`Locale`](crate::render::resolve::locale::Locale)
/// bucket (still primary-subtag-only — it also picks which language's number
/// words a §17.9.27 label is spelled in, a question region doesn't change)
/// and the §17.18.85 decimal separator
/// itself (issue #128), resolved from the *same* tag rather than derived from
/// the bucket: a bucket built from the primary subtag alone cannot tell
/// `de-DE` from `de-CH`, which is exactly the distinction the separator
/// needs. [`crate::i18n::decimal_separator_for_tag`] answers from real CLDR
/// data first; `Locale::decimal_separator` is the fallback when ICU4X has
/// nothing for the tag (or the cascade declared none at all).
pub(super) fn paragraph_locale(
    para: &model::Paragraph,
    resolved: &ResolvedDocument,
) -> (crate::render::resolve::locale::Locale, char) {
    use crate::render::resolve::locale::Locale;

    let tag = resolve_lang_tag(para, resolved);
    let locale = tag.map_or(Locale::English, Locale::from_tag);
    let decimal_separator = tag
        .and_then(crate::i18n::decimal_separator_for_tag)
        .unwrap_or_else(|| locale.decimal_separator());

    (locale, decimal_separator)
}

/// The raw §17.3.2.20 tag [`paragraph_locale`] resolves from, before it is
/// reduced to a `Locale` bucket or a separator.
///
/// Split out because a second consumer needs the tag itself, not either
/// answer derived from it: §17.16.4.2 date-picture names
/// (`crate::i18n::month_name_for_tag`) are looked up per tag, and a `DATE`
/// field's text is built during fragment collection, where the paragraph's
/// resolved style is not yet in hand. Both callers walk the same cascade
/// through this one function rather than assembling the layer list twice —
/// the layers, and their order, are the part that must not drift.
pub(super) fn resolve_lang_tag<'a>(
    para: &'a model::Paragraph,
    resolved: &'a ResolvedDocument,
) -> Option<&'a str> {
    let style_run = para
        .style_id
        .as_ref()
        .or(resolved.default_paragraph_style_id.as_ref())
        .and_then(|id| resolved.styles.get(id))
        .map(|s| &s.run);

    crate::render::resolve::locale::Locale::first_tag(
        [
            para.mark_run_properties.as_ref(),
            style_run,
            Some(&resolved.doc_defaults_run),
        ]
        .into_iter()
        .flatten(),
    )
}

/// §17.3.1.19: this paragraph's PDF outline entry, if it is a heading.
///
/// `props` is the **cascaded** (§17.7.2) paragraph properties, so a level
/// inherited from a heading style counts exactly as a direct `w:outlineLvl`
/// does — which is
/// the point: §17.3.1.19 is a paragraph property, and "is a heading" is not a
/// question about style names. `OutlineLevel::from_ooxml` has already rejected
/// value 9 ("body text"), so `None` here means "not a heading" for either
/// reason.
///
/// Returns `None` for a heading with no text: the title is the only thing a
/// reader can navigate by, and an untitled outline row points somewhere
/// unnameable. Consuming a node ID for it would also leave a gap in the
/// sequence, so the counter is only advanced once a heading is certain.
pub(super) fn paragraph_outline(
    para: &model::Paragraph,
    props: &model::ParagraphProperties,
    state: &mut BuildState,
) -> Option<crate::render::layout::draw_command::OutlineHeading> {
    let level = *props.outline_level.get()?;
    let title = outline_title(&para.content);
    if title.is_empty() {
        return None;
    }
    // Reserved last, so a body-text paragraph or an untitled heading never
    // consumes an ID and leaves a hole in the sequence.
    let node_id = state.next_outline_node_id()?;
    Some(crate::render::layout::draw_command::OutlineHeading {
        node_id,
        level,
        title: std::rc::Rc::from(title.as_str()),
    })
}

/// The text of a heading paragraph, as the outline should title it.
///
/// Runs are joined, and hyperlinks and field *results* are walked into, because
/// all three are ordinary text as far as a reader is concerned — a title cut at
/// the first formatting change would be wrong far more often than not.
///
/// The §17.9.22 list label is deliberately absent: it is injected at layout,
/// not part of the paragraph, and both Word and LibreOffice title a numbered
/// heading with its text alone (measured against the reference PDF attached to
/// issue #90).
fn outline_title(inlines: &[model::Inline]) -> String {
    fn walk(inlines: &[model::Inline], out: &mut String) {
        for inline in inlines {
            match inline {
                model::Inline::TextRun(run) => {
                    for element in &run.content {
                        if let model::RunElement::Text(text) = element {
                            out.push_str(text);
                        }
                    }
                }
                model::Inline::Symbol(sym) => {
                    out.push(char::from_u32(sym.char_code as u32).unwrap_or('\u{FFFD}'));
                }
                model::Inline::Hyperlink(link) => walk(&link.content, out),
                model::Inline::Field(field) => walk(&field.content, out),
                _ => {}
            }
        }
    }
    let mut out = String::new();
    walk(inlines, &mut out);
    out.trim().to_string()
}

pub(super) fn doc_font_family(ctx: &BuildContext) -> String {
    ctx.resolved
        .theme
        .as_ref()
        .map(|t| t.minor_font.latin.as_str())
        .filter(|s| !s.is_empty())
        .unwrap_or(SPEC_FALLBACK_FONT)
        .to_string()
}

pub(super) fn doc_font_size(ctx: &BuildContext) -> Pt {
    ctx.resolved
        .doc_defaults_run
        .font_size
        .cloned()
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE)
}

/// Convert a model paragraph properties into a layout ParagraphStyle.
///
/// `auto_fit` is the §20.1.2.1.18 `a:normAutofit` shrink of the enclosing shape
/// text body. It is carried onto the style rather than applied here: its
/// `@lnSpcReduction` half has to reach the line height *after* §17.3.1.33
/// resolution — see [`ShapeAutoFit::scale_line_height`](crate::render::layout::ShapeAutoFit::scale_line_height).
/// [`ShapeAutoFit::NONE`](crate::render::layout::ShapeAutoFit::NONE) outside a
/// shape text box.
/// §17.3.1.6 `w:bidi` — the paragraph's base embedding direction.
///
/// One reading of the property, because two things ask: level resolution
/// (`fragment::bidi`, before the paragraph has a style) and the style itself.
/// `props.bidi` has been parsed and cascaded correctly since long before
/// either; this is the function that stops discarding it.
///
/// Absent means left-to-right, and is never inferred from the text — see
/// [`BaseDirection`] for why UAX #9's own P2/P3 derivation is overridden.
pub(super) fn base_direction(props: &model::ParagraphProperties) -> BaseDirection {
    if props.bidi.unwrap_or(false) {
        BaseDirection::Rtl
    } else {
        BaseDirection::Ltr
    }
}

pub(super) fn paragraph_style_from_props(
    props: &model::ParagraphProperties,
    default_tab_stop: Pt,
    auto_fit: crate::render::layout::ShapeAutoFit,
    (locale, decimal_separator): (crate::render::resolve::locale::Locale, char),
    outline: Option<crate::render::layout::draw_command::OutlineHeading>,
) -> ParagraphStyle {
    let indent_left = props
        .indentation
        .get()
        .and_then(|i| i.start)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);
    let indent_right = props
        .indentation
        .get()
        .and_then(|i| i.end)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);
    let indent_first_line = props
        .indentation
        .get()
        .and_then(|i| i.first_line)
        .map(|fl| match fl {
            FirstLineIndent::FirstLine(v) => Pt::from(v),
            FirstLineIndent::Hanging(v) => -Pt::from(v),
            FirstLineIndent::None => Pt::ZERO,
        })
        .unwrap_or(Pt::ZERO);

    // §17.3.1.33: when autoSpacing is true, use 14pt instead of explicit value.
    let space_before = if props.spacing.get().and_then(|s| s.before_auto_spacing) == Some(true) {
        Pt::new(14.0)
    } else {
        props
            .spacing
            .get()
            .and_then(|s| s.before)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO)
    };
    let space_after = if props.spacing.get().and_then(|s| s.after_auto_spacing) == Some(true) {
        Pt::new(14.0)
    } else {
        props
            .spacing
            .get()
            .and_then(|s| s.after)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO)
    };

    // §17.3.1.33: line spacing defaults to single (auto, 240 twips = 1.0x).
    let line_spacing = props
        .spacing
        .get()
        .and_then(|s| s.line)
        .map(|ls| match ls {
            LineSpacing::Auto(v) => LineSpacingRule::Auto(Pt::from(v).raw() / 12.0),
            LineSpacing::Exact(v) => LineSpacingRule::Exact(Pt::from(v)),
            LineSpacing::AtLeast(v) => LineSpacingRule::AtLeast(Pt::from(v)),
        })
        .unwrap_or(LineSpacingRule::Auto(1.0));

    // §17.3.1.38: convert tab stops to layout format.
    // Clear entries are directives consumed during style merging, not layout stops.
    let tabs: Vec<TabStopDef> = props
        .tabs
        .iter()
        .filter(|t| t.alignment != model::TabAlignment::Clear)
        .map(|t| TabStopDef {
            position: Pt::from(t.position),
            alignment: t.alignment,
            leader: t.leader,
        })
        .collect();

    ParagraphStyle {
        auto_fit,
        alignment: props.alignment.cloned().unwrap_or(model::Alignment::Start),
        base_direction: base_direction(props),
        space_before,
        space_after,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        tabs,
        default_tab_stop,
        locale,
        decimal_separator,
        outline,
        drop_cap: None,
        borders: resolve_paragraph_borders(props),
        shading: props
            .shading
            .get()
            .map(|s| resolve_color(s.fill, ColorContext::Background)),
        keep_next: props.keep_next.unwrap_or(false),
        keep_lines: props.keep_lines.unwrap_or(false),
        // §17.3.1.44: Word enables widow/orphan control by default.
        widow_control: props.widow_control.unwrap_or(true),
        contextual_spacing: props.contextual_spacing.unwrap_or(false),
        style_id: None, // set by caller when available
        page_floats: Vec::new(),
        page_y: crate::render::dimension::Pt::ZERO,
        page_x: crate::render::dimension::Pt::ZERO,
        page_content_width: crate::render::dimension::Pt::ZERO,
    }
}

/// §17.3.1.24: resolve paragraph borders.
///
/// Each side that arrives as `Some(Border { style: BorderStyle::None, .. })`
/// is the spec's explicit "no border" override (`<w:top w:val="nil"/>` or
/// `<w:top w:val="none"/>`, §17.18.2 ST_Border). The model preserves the
/// `Some` so the §17.7.2 cascade can override an inherited side, but at
/// the render boundary we drop the sentinel — otherwise an empty `<w:pBdr>`
/// with all sides "nil" (Word's default for any paragraph cascade) would
/// frame every paragraph in the document with a hairline rectangle.
pub(super) fn resolve_paragraph_borders(
    props: &model::ParagraphProperties,
) -> Option<ParagraphBorderStyle> {
    let pbdr = props.borders.get()?;

    let convert = |b: &model::Border| -> Option<BorderLine> {
        if b.style.draws_nothing() {
            return None;
        }
        Some(BorderLine {
            width: Pt::from(b.width),
            color: resolve_color(b.color, ColorContext::Text),
            space: Pt::from(b.space),
        })
    };

    let style = ParagraphBorderStyle {
        top: pbdr.top.as_ref().and_then(convert),
        bottom: pbdr.bottom.as_ref().and_then(convert),
        left: pbdr.left.as_ref().and_then(convert),
        right: pbdr.right.as_ref().and_then(convert),
    };

    if style.top.is_some()
        || style.bottom.is_some()
        || style.left.is_some()
        || style.right.is_some()
    {
        Some(style)
    } else {
        None
    }
}

/// Convert a model `Border` to a layout `TableBorderLine`.
///
/// §17.4.38 defines 26 border styles; the painter draws two. `Double` maps
/// through, and the remaining 24 — `Dotted`, `Dashed`, `Triple`, `Wave`,
/// `DashDotStroked`, the `ThinThick*` family and the rest — are approximated
/// by a solid line of the declared width and colour. The approximation is
/// bounded (position, thickness and colour are all preserved; only the stroke
/// pattern is lost) but it is not free, so each distinct style is reported once
/// per render via `state.warned_border_styles`.
fn convert_model_border(b: &model::Border, state: &mut BuildState) -> TableBorderLine {
    let style = match b.style {
        model::BorderStyle::Double => TableBorderStyle::Double,
        model::BorderStyle::Single => TableBorderStyle::Single,
        other => {
            if state.warned_border_styles.insert(other) {
                log::warn!(
                    "[build] §17.4.38 border style {other:?} is not drawn; \
                     approximating with a single solid line"
                );
            }
            TableBorderStyle::Single
        }
    };
    TableBorderLine {
        width: Pt::from(b.width),
        color: resolve_color(b.color, ColorContext::Text),
        style,
    }
}

/// Convert a model cell border to a `CellBorderOverride`.
///
/// Three inputs, three outcomes ([MS-OI29500] §17.4.66):
///
/// | Markup | Result | Meaning |
/// |---|---|---|
/// | edge omitted | `None` | inherit: style → `tblPrEx` → table borders |
/// | `val="none"` | `None` | **same as omitted** — the spec's cascade starts with *"if the cell border is omitted or none"* |
/// | `val="nil"` | `Some(Suppress)` | draw nothing, and win the shared-edge conflict |
/// | anything else | `Some(Border(..))` | that border, subject to conflict resolution |
///
/// `none` collapsing to `None` is the whole point: it makes an explicit
/// `<w:top w:val="none"/>` behave as if the author had written nothing, which
/// is what the spec says and the opposite of suppressing the edge.
pub(super) fn convert_cell_border_override(
    b: &Option<model::Border>,
    state: &mut BuildState,
) -> Option<CellBorderOverride> {
    b.as_ref().and_then(|b| match b.style {
        model::BorderStyle::Nil => Some(CellBorderOverride::Suppress),
        model::BorderStyle::None => None,
        _ => Some(CellBorderOverride::Border(convert_model_border(b, state))),
    })
}

/// §17.4.38: merge direct table borders over style borders.
/// Each edge in `direct` overrides the corresponding edge in `style`;
/// unspecified edges (`None`) inherit from the style.
pub(super) fn merge_table_borders(
    direct: &model::TableBorders,
    style: &model::TableBorders,
) -> model::TableBorders {
    model::TableBorders {
        top: direct.top.or(style.top),
        bottom: direct.bottom.or(style.bottom),
        left: direct.left.or(style.left),
        right: direct.right.or(style.right),
        inside_h: direct.inside_h.or(style.inside_h),
        inside_v: direct.inside_v.or(style.inside_v),
    }
}

/// Convert model `TableBorders` to a layout `TableBorderConfig`.
/// §17.4.38: borders with `val="none"` or `val="nil"` are suppressed.
///
/// Filtered here, *before* [`convert_model_border`] ever runs — that function
/// has no "draw nothing" style of its own, so an unsuppressed `none`/`nil`
/// would fall into its unhandled-style arm and be approximated as a visible
/// solid line instead of correctly drawing nothing.
pub(super) fn convert_table_border_config(
    b: &model::TableBorders,
    state: &mut BuildState,
) -> TableBorderConfig {
    let mut convert = |border: &Option<model::Border>| -> Option<TableBorderLine> {
        border.as_ref().and_then(|b| {
            if b.style.draws_nothing() {
                None
            } else {
                Some(convert_model_border(b, state))
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

/// Populate image data on Fragment::Image fragments from the media map.
pub(super) fn populate_image_data(
    fragments: &mut [Fragment],
    media: &HashMap<model::RelId, MediaEntry>,
) {
    for frag in fragments.iter_mut() {
        if let Fragment::Image {
            rel_id, image_data, ..
        } = frag
        {
            if image_data.is_none() {
                if let Some(entry) = media.get(&model::RelId::new(rel_id.as_str())) {
                    *image_data = Some(entry.clone());
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
pub(super) fn remap_legacy_font_chars(
    text: &str,
    font_family: &str,
    fallback_family: &str,
) -> (String, String) {
    let is_symbol = font_family.eq_ignore_ascii_case("Symbol");
    let is_wingdings = font_family.eq_ignore_ascii_case("Wingdings");

    if !is_symbol && !is_wingdings {
        let family = if font_family.is_empty() {
            fallback_family
        } else {
            font_family
        };
        return (text.to_string(), family.to_string());
    }

    let mut remapped = String::with_capacity(text.len());
    // True once a PUA codepoint is found that we have no Unicode mapping for.
    // Such a character survives as a raw PUA scalar, which *only* the original
    // legacy font can render — see the family choice below.
    let mut unmapped_pua = false;

    for ch in text.chars() {
        let code = ch as u32;
        if !LEGACY_PUA_RANGE.contains(&code) {
            // Not legacy-encoded; pass through untouched.
            remapped.push(ch);
            continue;
        }
        let mapped = if is_symbol {
            symbol_pua_to_unicode(code)
        } else {
            wingdings_pua_to_unicode(code)
        };
        match mapped {
            Some(unicode) => remapped.push(unicode),
            None => {
                unmapped_pua = true;
                remapped.push(ch);
            }
        }
    }

    // Font choice follows what actually happened to the text:
    //
    // * Everything mapped → the text is now standard Unicode, which the legacy
    //   font cannot render (Symbol/Wingdings carry no Unicode cmap). Switch to
    //   the document's fallback font; standard bullets (U+2022, U+25AA) and
    //   Greek are present in most text fonts.
    // * Something didn't map → that character is still a PUA scalar, and the
    //   *only* font that can render it is the legacy one it came from. Keeping
    //   the original family renders correctly wherever the real Symbol /
    //   Wingdings face is installed; handing a PUA scalar to a text font is a
    //   guaranteed `.notdef` box.
    let family = if unmapped_pua {
        font_family
    } else {
        fallback_family
    };
    (remapped, family.to_string())
}

/// The PUA block Word uses for legacy symbol-font encodings: character code
/// `0xNN` is stored as `0xF0NN` (§17.3.3.30 `w:sym`).
const LEGACY_PUA_RANGE: std::ops::RangeInclusive<u32> = 0xF020..=0xF0FF;

/// Symbol-font PUA codepoint → Unicode, per the Adobe mapping table
/// (unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt).
///
/// `None` means "no mapping known" — deliberately distinct from "maps to
/// itself", because the caller uses it to decide whether the remapped text can
/// safely move to a Unicode text font.
fn symbol_pua_to_unicode(code: u32) -> Option<char> {
    // Greek is the Symbol font's primary content: the letter positions carry
    // the alphabet, with four glyph-variant slots (theta1, sigma1, phi1,
    // omega1) that do not follow the plain A–Z / a–z order.
    const GREEK_UPPER: [char; 26] = [
        '\u{0391}', // A Alpha
        '\u{0392}', // B Beta
        '\u{03A7}', // C Chi
        '\u{0394}', // D Delta
        '\u{0395}', // E Epsilon
        '\u{03A6}', // F Phi
        '\u{0393}', // G Gamma
        '\u{0397}', // H Eta
        '\u{0399}', // I Iota
        '\u{03D1}', // J theta1 (GREEK THETA SYMBOL)
        '\u{039A}', // K Kappa
        '\u{039B}', // L Lambda
        '\u{039C}', // M Mu
        '\u{039D}', // N Nu
        '\u{039F}', // O Omicron
        '\u{03A0}', // P Pi
        '\u{0398}', // Q Theta
        '\u{03A1}', // R Rho
        '\u{03A3}', // S Sigma
        '\u{03A4}', // T Tau
        '\u{03A5}', // U Upsilon
        '\u{03C2}', // V sigma1 (FINAL SIGMA)
        '\u{03A9}', // W Omega
        '\u{039E}', // X Xi
        '\u{03A8}', // Y Psi
        '\u{0396}', // Z Zeta
    ];
    const GREEK_LOWER: [char; 26] = [
        '\u{03B1}', // a alpha
        '\u{03B2}', // b beta
        '\u{03C7}', // c chi
        '\u{03B4}', // d delta
        '\u{03B5}', // e epsilon
        '\u{03C6}', // f phi
        '\u{03B3}', // g gamma
        '\u{03B7}', // h eta
        '\u{03B9}', // i iota
        '\u{03D5}', // j phi1 (GREEK PHI SYMBOL)
        '\u{03BA}', // k kappa
        '\u{03BB}', // l lambda
        '\u{03BC}', // m mu
        '\u{03BD}', // n nu
        '\u{03BF}', // o omicron
        '\u{03C0}', // p pi
        '\u{03B8}', // q theta
        '\u{03C1}', // r rho
        '\u{03C3}', // s sigma
        '\u{03C4}', // t tau
        '\u{03C5}', // u upsilon
        '\u{03D6}', // v omega1 (GREEK PI SYMBOL)
        '\u{03C9}', // w omega
        '\u{03BE}', // x xi
        '\u{03C8}', // y psi
        '\u{03B6}', // z zeta
    ];

    Some(match code {
        0xF020 => '\u{0020}',                              // SPACE
        0xF021 => '\u{0021}',                              // EXCLAMATION MARK
        0xF025 => '\u{0025}',                              // PERCENT SIGN
        0xF028 => '\u{0028}',                              // LEFT PARENTHESIS
        0xF029 => '\u{0029}',                              // RIGHT PARENTHESIS
        0xF02B => '\u{002B}',                              // PLUS SIGN
        0xF02E => '\u{002E}',                              // FULL STOP
        0xF030..=0xF039 => char::from_u32(code - 0xF000)?, // DIGITS
        0xF03C => '\u{003C}',                              // LESS-THAN SIGN
        0xF03D => '\u{003D}',                              // EQUALS SIGN
        0xF03E => '\u{003E}',                              // GREATER-THAN SIGN
        0xF041..=0xF05A => GREEK_UPPER[(code - 0xF041) as usize],
        0xF05B => '\u{005B}', // LEFT SQUARE BRACKET
        0xF05D => '\u{005D}', // RIGHT SQUARE BRACKET
        0xF061..=0xF07A => GREEK_LOWER[(code - 0xF061) as usize],
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
        _ => return None,
    })
}

/// Wingdings PUA codepoint → Unicode, per the Microsoft Wingdings-to-Unicode
/// table. `None` means "no mapping known" — see [`symbol_pua_to_unicode`].
///
/// Wingdings is a dingbat font with no canonical full mapping; this covers the
/// glyphs that appear as list bullets, which is what routes through here.
fn wingdings_pua_to_unicode(code: u32) -> Option<char> {
    Some(match code {
        0xF021 => '\u{270E}',  // LOWER RIGHT PENCIL
        0xF022 => '\u{2702}',  // BLACK SCISSORS
        0xF023 => '\u{2701}',  // UPPER BLADE SCISSORS
        0xF028 => '\u{1F4CB}', // CLIPBOARD
        0xF029 => '\u{1F4CB}', // CLIPBOARD
        0xF041 => '\u{FE4E}',  // WAVY LOW LINE (approximate)
        0xF046 => '\u{1F44D}', // THUMBS UP SIGN
        0xF04A => '\u{263A}',  // WHITE SMILING FACE
        0xF04C => '\u{2639}',  // WHITE FROWNING FACE
        0xF06C => '\u{25CF}',  // BLACK CIRCLE
        0xF06D => '\u{274D}',  // SHADOWED WHITE CIRCLE
        0xF06E => '\u{25A0}',  // BLACK SQUARE
        0xF06F => '\u{25A1}',  // WHITE SQUARE
        0xF070 => '\u{25A1}',  // WHITE SQUARE (alt)
        0xF071 => '\u{2751}',  // LOWER RIGHT SHADOWED WHITE SQUARE
        0xF072 => '\u{2752}',  // UPPER RIGHT SHADOWED WHITE SQUARE
        0xF073 => '\u{25C6}',  // BLACK DIAMOND
        0xF074 => '\u{2756}',  // BLACK DIAMOND MINUS WHITE X
        0xF076 => '\u{2756}',  // BLACK DIAMOND MINUS WHITE X
        0xF09F => '\u{2708}',  // AIRPLANE
        0xF0A1 => '\u{270C}',  // VICTORY HAND
        0xF0A4 => '\u{261C}',  // WHITE LEFT POINTING INDEX
        0xF0A5 => '\u{261E}',  // WHITE RIGHT POINTING INDEX
        0xF0A7 => '\u{25AA}',  // BLACK SMALL SQUARE
        0xF0A8 => '\u{25FB}',  // WHITE MEDIUM SQUARE
        0xF0D5 => '\u{232B}',  // ERASE TO THE LEFT
        0xF0D8 => '\u{27A2}',  // THREE-D TOP-LIGHTED RIGHTWARDS ARROWHEAD
        0xF0E8 => '\u{2B22}',  // BLACK HEXAGON (approximate)
        0xF0F0 => '\u{2B1A}',  // DOTTED SQUARE (approximate)
        0xF0FB => '\u{2718}',  // HEAVY BALLOT X
        0xF0FC => '\u{2714}',  // HEAVY CHECK MARK
        0xF0FE => '\u{2612}',  // BALLOT BOX WITH X (approximate)
        _ => return None,
    })
}

/// Populate underline and strike position/thickness from Skia font metrics.
pub(super) fn populate_underline_metrics(fragments: &mut [Fragment], measurer: &TextMeasurer) {
    for frag in fragments.iter_mut() {
        if let Fragment::Text { font, .. } = frag {
            if font.underline || font.strike_lines > 0 {
                // Decorated runs only; `make_mut` clones the shared font so
                // the metrics can be written back.
                let fp = std::rc::Rc::make_mut(font);
                if fp.underline {
                    let (pos, thickness) = measurer.underline_metrics(fp);
                    fp.underline_position = pos;
                    fp.underline_thickness = thickness;
                }
                if fp.strike_lines > 0 {
                    let (pos, thickness) = measurer.strike_metrics(fp);
                    fp.strike_position = pos;
                    fp.strike_thickness = thickness;
                }
            }
        }
    }
}

/// Extract the display size for a picture bullet from its VML shape style.
/// Falls back to 9pt × 9pt (common Word default for picture bullets).
pub(super) fn pic_bullet_size(bullet: &model::NumPicBullet) -> PtSize {
    let default = PtSize::new(Pt::new(9.0), Pt::new(9.0));
    let shape = match bullet.pict.as_ref().and_then(|p| p.shapes().next()) {
        Some(s) => s,
        None => return default,
    };

    let w = shape
        .common
        .style
        .width
        .and_then(vml_style_length_to_pt)
        .unwrap_or(default.width);
    let h = shape
        .common
        .style
        .height
        .and_then(vml_style_length_to_pt)
        .unwrap_or(default.height);
    PtSize::new(w, h)
}

/// §14.1.2: convert a VML **style** measurement (`width`, `height`,
/// `margin-left`, …) to points.
///
/// The absolute units come from [`model::VmlLength::to_absolute_points`], the
/// one unit table shared with the parse layer. What this function adds is the
/// policy for the units that table leaves unresolved, which is specific to
/// style measurements:
///
/// * a **bare number** is EMU — the interpretation the model documents for
///   this context (on a primitive's `from`/`to` the same bare number would
///   instead be local coordinate units, which is why the rule can't live on
///   the model type);
/// * **`%`** and **`em`** are relative to a containing box or font size that
///   isn't available here, so they yield `None` rather than a fabricated
///   number. Returning the raw value — the previous behaviour — turned
///   `width:50%` into 50pt.
pub(super) fn vml_style_length_to_pt(len: model::VmlLength) -> Option<Pt> {
    use crate::model::VmlLengthUnit;
    if let Some(pt) = len.to_absolute_points() {
        return Some(Pt::new(pt));
    }
    match len.unit {
        VmlLengthUnit::None => Some(Pt::new(len.value as f32 / 914400.0 * 72.0)),
        // Unresolvable without a containing box / font size.
        VmlLengthUnit::Em | VmlLengthUnit::Percent => None,
        // `to_absolute_points` already handled every absolute unit.
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::Dimension;
    use crate::model::Dup;
    use crate::model::{Border, BorderStyle, Color, ParagraphBorders, ParagraphProperties};
    use crate::render::resolve::color::rgb_from_u32;

    /// An all-empty ResolvedDocument for exercising `resolve_paragraph_defaults`.
    fn empty_resolved() -> ResolvedDocument {
        use std::collections::HashMap;
        ResolvedDocument {
            sections: Vec::new(),
            styles: HashMap::new(),
            numbering: HashMap::new(),
            font_families: Vec::new(),
            media: HashMap::new(),
            embedded_fonts: Vec::new(),
            pic_bullets: HashMap::new(),
            theme: None,
            doc_defaults_paragraph: ParagraphProperties::default(),
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
        }
    }

    fn bare_para() -> model::Paragraph {
        model::Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: Vec::new(),
            rsids: model::ParagraphRevisionIds::default(),
        }
    }

    #[test]
    fn shape_font_ref_base_color_and_family_apply_as_defaults() {
        // §20.1.4.1.17: a shape's fontRef supplies the base text color / family
        // used when the run/style specify none.
        let resolved = empty_resolved();
        let para = bare_para();

        let (_, _, color, _, _) =
            resolve_paragraph_defaults(&para, &resolved, false, Some(rgb_from_u32(0xFF0000)), None);
        assert_eq!(color, rgb_from_u32(0xFF0000), "fontRef base color applies");

        // Without a base, the default stays black (unchanged behavior).
        let (_, _, black, _, _) = resolve_paragraph_defaults(&para, &resolved, false, None, None);
        assert_eq!(black, rgb_from_u32(0x000000));

        // Base family applies when set.
        let (family, _, _, _, _) =
            resolve_paragraph_defaults(&para, &resolved, false, None, Some("Foo Sans"));
        assert_eq!(family, "Foo Sans");
    }

    fn border_with_style(style: BorderStyle) -> Border {
        Border {
            style,
            width: Dimension::new(0),
            space: Dimension::new(0),
            color: Color::Auto,
        }
    }

    /// §17.18.2: every side carrying `BorderStyle::None` (the parser's
    /// landing for `<w:top w:val="nil"/>` and `<w:top w:val="none"/>`)
    /// is dropped, even though the side's `Option<Border>` is `Some(_)`.
    #[test]
    fn paragraph_borders_explicit_none_yields_no_render_borders() {
        let props = ParagraphProperties {
            borders: Dup::from(Some(ParagraphBorders {
                top: Some(border_with_style(BorderStyle::None)),
                bottom: Some(border_with_style(BorderStyle::None)),
                left: Some(border_with_style(BorderStyle::None)),
                right: Some(border_with_style(BorderStyle::None)),
                between: Some(border_with_style(BorderStyle::None)),
            })),
            ..Default::default()
        };
        assert!(
            resolve_paragraph_borders(&props).is_none(),
            "all-sides-nil pPr must produce no ParagraphBorderStyle"
        );
    }

    /// One real side stays — the others (None) are still filtered.
    #[test]
    fn paragraph_borders_mixed_keeps_only_actual_sides() {
        let props = ParagraphProperties {
            borders: Dup::from(Some(ParagraphBorders {
                top: Some(border_with_style(BorderStyle::Single)),
                bottom: Some(border_with_style(BorderStyle::None)),
                left: Some(border_with_style(BorderStyle::None)),
                right: Some(border_with_style(BorderStyle::None)),
                between: None,
            })),
            ..Default::default()
        };
        let resolved = resolve_paragraph_borders(&props).expect("top side must survive");
        assert!(resolved.top.is_some());
        assert!(resolved.bottom.is_none());
        assert!(resolved.left.is_none());
        assert!(resolved.right.is_none());
    }

    /// All-Single → carries through to render-side struct.
    #[test]
    fn paragraph_borders_all_single_round_trips() {
        let props = ParagraphProperties {
            borders: Dup::from(Some(ParagraphBorders {
                top: Some(border_with_style(BorderStyle::Single)),
                bottom: Some(border_with_style(BorderStyle::Single)),
                left: Some(border_with_style(BorderStyle::Single)),
                right: Some(border_with_style(BorderStyle::Single)),
                between: None,
            })),
            ..Default::default()
        };
        let resolved = resolve_paragraph_borders(&props).expect("must produce Some");
        assert!(resolved.top.is_some());
        assert!(resolved.bottom.is_some());
        assert!(resolved.left.is_some());
        assert!(resolved.right.is_some());
    }

    /// No `<w:pBdr>` at all → no render-side borders.
    #[test]
    fn paragraph_borders_absent_yields_none() {
        let props = ParagraphProperties::default();
        assert!(resolve_paragraph_borders(&props).is_none());
    }

    // ── §17.3.3.30 legacy symbol-font PUA remapping ──────────────────────

    /// The Symbol font's letter positions carry the Greek alphabet — its
    /// primary content, and previously absent from the table entirely, so
    /// every Greek character survived as a raw PUA scalar.
    #[test]
    fn symbol_pua_maps_the_greek_alphabet() {
        // Plain alphabetical positions, spot-checked at both ends.
        for (code, expected) in [
            (0xF041, '\u{0391}'), // A → Alpha
            (0xF044, '\u{0394}'), // D → Delta
            (0xF05A, '\u{0396}'), // Z → Zeta
            (0xF061, '\u{03B1}'), // a → alpha
            (0xF06D, '\u{03BC}'), // m → mu
            (0xF07A, '\u{03B6}'), // z → zeta
        ] {
            assert_eq!(symbol_pua_to_unicode(code), Some(expected), "U+{code:04X}");
        }
    }

    /// Four Symbol slots hold glyph *variants* rather than the letter their
    /// position would suggest — the easiest thing to get wrong when filling
    /// in the alphabet mechanically.
    #[test]
    fn symbol_pua_greek_variant_slots_are_not_alphabetical() {
        for (code, expected, what) in [
            (0xF04A, '\u{03D1}', "J is theta1, not Iota-then-Kappa order"),
            (0xF056, '\u{03C2}', "V is final sigma, not Upsilon+1"),
            (0xF06A, '\u{03D5}', "j is phi1 (phi symbol)"),
            (0xF076, '\u{03D6}', "v is omega1 (pi symbol), not omega"),
        ] {
            assert_eq!(symbol_pua_to_unicode(code), Some(expected), "{what}");
        }
        // The plain letters those variants are easily confused with.
        assert_eq!(
            symbol_pua_to_unicode(0xF051),
            Some('\u{0398}'),
            "Q is Theta"
        );
        assert_eq!(symbol_pua_to_unicode(0xF066), Some('\u{03C6}'), "f is phi");
        assert_eq!(
            symbol_pua_to_unicode(0xF077),
            Some('\u{03C9}'),
            "w is omega"
        );
    }

    /// Adding the Greek ranges must not have displaced the punctuation and
    /// math symbols that sit between them.
    #[test]
    fn symbol_pua_keeps_non_greek_mappings() {
        for (code, expected) in [
            (0xF05B, '\u{005B}'), // [ — just past uppercase Greek
            (0xF05D, '\u{005D}'), // ]
            (0xF07B, '\u{007B}'), // { — just past lowercase Greek
            (0xF0B7, '\u{2022}'), // bullet
            (0xF030, '\u{0030}'), // digit zero
            (0xF0F2, '\u{222B}'), // integral
        ] {
            assert_eq!(symbol_pua_to_unicode(code), Some(expected), "U+{code:04X}");
        }
    }

    /// An unmapped codepoint is `None`, not "maps to itself" — the caller
    /// needs the distinction to choose a font.
    #[test]
    fn unmapped_pua_is_none() {
        assert_eq!(symbol_pua_to_unicode(0xF0FF), None);
        assert_eq!(wingdings_pua_to_unicode(0xF0FF), None);
    }

    /// When every PUA character maps, the text is standard Unicode and must
    /// move to the fallback text font — the legacy face has no Unicode cmap.
    #[test]
    fn fully_mapped_text_switches_to_the_fallback_font() {
        let (text, family) = remap_legacy_font_chars("\u{F0B7}", "Symbol", "Calibri");
        assert_eq!(text, "\u{2022}", "bullet remapped");
        assert_eq!(family, "Calibri");
    }

    /// Regression: an *unmapped* PUA character used to be handed to the
    /// fallback text font anyway, which is a guaranteed `.notdef` box. The
    /// only font that can render a surviving PUA scalar is the legacy face it
    /// came from, so the original family must be kept.
    #[test]
    fn unmapped_pua_keeps_the_legacy_font() {
        let (text, family) = remap_legacy_font_chars("\u{F0FF}", "Symbol", "Calibri");
        assert_eq!(text, "\u{F0FF}", "unmapped codepoint survives unchanged");
        assert_eq!(
            family, "Symbol",
            "keep the legacy face — it is the only one that can render this"
        );
    }

    /// Mixed input: one unmapped character is enough to keep the legacy font,
    /// because the family is a whole-string decision.
    #[test]
    fn one_unmapped_char_keeps_the_legacy_font_for_the_whole_label() {
        let (_, family) = remap_legacy_font_chars("\u{F0B7}\u{F0FF}", "Symbol", "Calibri");
        assert_eq!(family, "Symbol");
    }

    /// Non-legacy families are passed through untouched.
    #[test]
    fn non_legacy_family_is_untouched() {
        let (text, family) = remap_legacy_font_chars("abc", "Arial", "Calibri");
        assert_eq!(text, "abc");
        assert_eq!(family, "Arial");
    }

    // ── §14.1.2 style-measurement policy ─────────────────────────────────

    fn vml_len(value: f64, unit: model::VmlLengthUnit) -> model::VmlLength {
        model::VmlLength { value, unit }
    }

    #[test]
    fn style_length_passes_absolute_units_through() {
        use model::VmlLengthUnit;
        assert_eq!(
            vml_style_length_to_pt(vml_len(9.0, VmlLengthUnit::Pt)),
            Some(Pt::new(9.0))
        );
        assert_eq!(
            vml_style_length_to_pt(vml_len(1.0, VmlLengthUnit::In)),
            Some(Pt::new(72.0))
        );
    }

    /// Regression: the two render-layer converters disagreed here — one read a
    /// bare number as EMU, the other as points, so the same `width:9` meant
    /// 0.0007pt in one code path and 9pt in the other. EMU is the rule the
    /// model documents for `style` measurements.
    #[test]
    fn style_length_reads_a_bare_number_as_emu() {
        let got = vml_style_length_to_pt(vml_len(914400.0, model::VmlLengthUnit::None))
            .expect("unitless resolves in a style measurement");
        assert!(
            (got.raw() - 72.0).abs() < 1e-3,
            "914400 EMU is one inch, got {}",
            got.raw()
        );
    }

    /// Regression: `%` and `em` used to fall through to the raw value, so
    /// `width:50%` became 50pt. They are unresolvable here; the caller
    /// substitutes its own default.
    #[test]
    fn style_length_rejects_units_needing_a_container() {
        for unit in [model::VmlLengthUnit::Percent, model::VmlLengthUnit::Em] {
            assert_eq!(
                vml_style_length_to_pt(vml_len(50.0, unit)),
                None,
                "{unit:?} needs a containing box / font size"
            );
        }
    }
}

#[cfg(test)]
mod border_style_tests {
    use super::*;

    fn border(style: model::BorderStyle) -> model::Border {
        model::Border {
            style,
            width: crate::model::dimension::Dimension::new(8),
            color: model::Color::Auto,
            space: crate::model::dimension::Dimension::new(0),
        }
    }

    /// §17.4.38: only `Single` and `Double` are drawn; the other 24 styles are
    /// approximated by a solid line of the declared width and colour.
    #[test]
    fn unsupported_styles_collapse_to_single_preserving_width_and_colour() {
        let mut state = BuildState::default();
        for style in [
            model::BorderStyle::Dotted,
            model::BorderStyle::Dashed,
            model::BorderStyle::Triple,
            model::BorderStyle::Wave,
            model::BorderStyle::DashDotStroked,
            model::BorderStyle::ThinThickLargeGap,
        ] {
            let line = convert_model_border(&border(style), &mut state);
            assert_eq!(
                line.style,
                TableBorderStyle::Single,
                "{style:?} should approximate as a single line"
            );
            assert_eq!(line.width, Pt::from(border(style).width), "width preserved");
        }
    }

    #[test]
    fn supported_styles_are_not_approximated() {
        let mut state = BuildState::default();
        assert_eq!(
            convert_model_border(&border(model::BorderStyle::Double), &mut state).style,
            TableBorderStyle::Double
        );
        assert_eq!(
            convert_model_border(&border(model::BorderStyle::Single), &mut state).style,
            TableBorderStyle::Single
        );
        assert!(
            state.warned_border_styles.is_empty(),
            "styles that render faithfully must not be reported as approximated"
        );
    }

    /// Each distinct style is recorded once, so the warning fires once per
    /// render instead of once per cell edge — thousands of lines for a large
    /// table with dotted borders.
    #[test]
    fn each_unsupported_style_is_recorded_exactly_once() {
        let mut state = BuildState::default();
        for _ in 0..50 {
            convert_model_border(&border(model::BorderStyle::Dotted), &mut state);
            convert_model_border(&border(model::BorderStyle::Dashed), &mut state);
        }
        assert_eq!(
            state.warned_border_styles.len(),
            2,
            "one entry per distinct style, regardless of occurrence count"
        );
        assert!(state
            .warned_border_styles
            .contains(&model::BorderStyle::Dotted));
        assert!(state
            .warned_border_styles
            .contains(&model::BorderStyle::Dashed));
    }
}
