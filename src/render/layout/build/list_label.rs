//! List label injection — prepend bullet/number labels to paragraph fragments.
//!
//! §17.9.22: when a paragraph carries a numbering reference, resolve the label
//! text (or picture bullet) and inject it as the first fragment, followed by
//! the level's §17.9.29 separator — by default a tab that advances to the
//! body-text indent. For centered and right-aligned paragraphs the separator
//! is a fixed-advance spacer instead of a tab, so the §17.3.1.37 tab-line
//! rule cannot suppress the paragraph's alignment.
//!
//! Because the label is the paragraph's first fragment, it is first-segment
//! -only by construction when the paragraph splits across pages — it lives on
//! line 0, which always belongs to the first segment.
//!
//! # Under `w:bidi`
//!
//! The label is an ordinary fragment in the paragraph's stream, so UAX #9 sees
//! it: `build::block::resolve_paragraph_bidi` runs *after* this module, and a
//! label followed by a §17.9.29 separator of `space` or `nothing` reorders with
//! the line like any other text — in a right-to-left paragraph it ends up at
//! the right, which is where it belongs.
//!
//! With the default separator, which is a **tab**, it does not, and that is the
//! same boundary §17.3.1.37 has everywhere else in this engine rather than a
//! second gap: `line_emit::visual_order` segments a line at every tab and does
//! not mirror the stop positions the pen jumps to, so a label before a tab
//! stays at the physical left. Closing it means mirroring tab geometry under
//! `w:bidi` — the stops, the pen, and the alignment-suppression rule together
//! — and there is no reference corpus here to check that against.

use crate::model::Dup;
use std::rc::Rc;

use crate::i18n::bidi::BidiLevel;
use crate::model::{self, ParagraphProperties};
use crate::render::dimension::Pt;
use crate::render::layout::fragment::Fragment;

use super::convert::{pic_bullet_size, remap_legacy_font_chars, resolve_paragraph_defaults};
use super::{BuildContext, BuildState};

/// Inject list label fragments into a paragraph if it has a numbering reference.
///
/// Updates `fragments` (prepends label + tab), `merged_props` (overrides indentation
/// from the numbering level), and `state.list_counters` (increments/resets counters).
pub(super) fn inject_list_label(
    para: &model::Paragraph,
    fragments: &mut Vec<Fragment>,
    merged_props: &mut ParagraphProperties,
    ctx: &BuildContext,
    state: &mut BuildState,
) {
    let num_ref = match merged_props.numbering.get() {
        Some(nr) => nr,
        None => return,
    };

    let num_id = model::NumId::new(num_ref.num_id);
    let level = num_ref.level;

    let levels = match ctx.resolved.numbering.get(&num_id) {
        Some(levels) => levels,
        None => return,
    };

    // Update counters: increment this level, reset deeper levels.
    //
    // The map holds the *current* (last-emitted) value, which is what
    // `format_list_label` reads. §17.9.28 permits `w:start="0"` (and
    // §17.9.27 `w:startOverride="0"`), so the first item must seed the
    // counter with `start` directly — seeding `start - 1` and adding one
    // underflows `u32` at zero, panicking debug builds and silently
    // wrapping in release.
    {
        use std::collections::hash_map::Entry;
        let counters = &mut state.list_counters;
        let level_start = |lvl: u8| levels.get(lvl as usize).map(|l| l.start).unwrap_or(1);
        // Word semantics: an item at level L instantiates every not-yet-active
        // ancestor counter at its start value without consuming it — after
        // "1.1.1"-style items, the next top-level item continues as "2.", not
        // "1.". Only ancestors missing from the map are touched.
        for ancestor in 0..level {
            counters
                .entry((num_id, ancestor))
                .or_insert_with(|| level_start(ancestor));
        }
        match counters.entry((num_id, level)) {
            Entry::Vacant(slot) => {
                slot.insert(level_start(level));
            }
            Entry::Occupied(mut slot) => {
                let next = slot.get().saturating_add(1);
                slot.insert(next);
            }
        }
        // Reset deeper levels. `saturating_add`: `w:ilvl` is parsed as a raw
        // u8, so a crafted ilvl=255 must yield an empty range here — `level + 1`
        // would overflow (panic in debug, wrap to 0 in release and wipe every
        // counter this numId owns, including the ancestors seeded above).
        let max_level = levels.len() as u8;
        for deeper in level.saturating_add(1)..max_level {
            counters.remove(&(num_id, deeper));
        }
    }

    let level_def = levels.get(level as usize);

    // §17.9.23: the level's indentation with the paragraph's direct overrides
    // applied. `None` when the level defines no indentation — the paragraph
    // then keeps its style-cascade value, which is exactly what
    // `merged_props.indentation` already holds.
    let effective_ind = effective_indentation(para, level_def);
    // Label geometry — hanging width, implicit tab stop — must follow the
    // paragraph's *final* indentation wherever it came from, or the suffix tab
    // lands somewhere other than where the body text wraps. Only the §17.9.23
    // overwrite below stays keyed to the level-derived value.
    let layout_ind = effective_ind.or(merged_props.indentation.cloned());

    // §17.9.10: check for picture bullet before text label.
    let pic_bullet_injected = level_def
        .and_then(|l| l.lvl_pic_bullet_id)
        .and_then(|pic_id| ctx.resolved.pic_bullets.get(&pic_id))
        .and_then(|bullet| {
            let rel_id = bullet
                .pict
                .as_ref()?
                .shapes()
                .next()?
                .common
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
                src_rect: None,
            };
            Some((label_frag, size.height))
        });

    if let Some((label_frag, label_height)) = pic_bullet_injected {
        let hanging = extract_hanging(layout_ind);
        // §17.9.29: `Nothing` drops the separator entirely. `Tab` and `Space`
        // both advance via a tab here — a picture bullet has no text font to emit
        // a literal space with, and image-bullet + space is vanishingly rare.
        let drop_separator =
            level_def.map(|l| l.suffix) == Some(crate::model::LevelSuffix::Nothing);
        if !drop_separator {
            // §17.3.1.38: a leader on this separator is drawn in the
            // formatting in effect at the tab. A picture bullet has no text
            // run of its own, so the paragraph defaults *are* that formatting.
            let (fam, size, color, _, _) =
                resolve_paragraph_defaults(para, ctx.resolved, false, None, None);
            let sep_font = crate::render::layout::fragment::font_props_from_run(
                &model::RunProperties::default(),
                &fam,
                size,
                state.shape_auto_fit,
            );
            if alignment_suppresses_tab(merged_props) {
                // An image's advance is its drawn size, so the gap cannot be
                // baked into the label fragment like the text path does.
                let gap = (hanging - label_frag.width()).max(Pt::ZERO);
                fragments.insert(0, spacer_fragment(gap, &sep_font, color, ctx.measurer));
            } else {
                let tab_frag = Fragment::Tab {
                    line_height: label_height,
                    font: Rc::new(sep_font),
                    color,
                    fitting_width: Some(hanging),
                };
                fragments.insert(0, tab_frag);
            }
        }
        fragments.insert(0, label_frag);

        if !drop_separator {
            insert_implicit_tab_stop(merged_props, layout_ind);
        }
    } else {
        inject_text_label(
            para,
            fragments,
            merged_props,
            ctx,
            &state.list_counters,
            levels,
            level,
            level_def,
            state.shape_auto_fit,
            layout_ind,
        );
    }

    // §17.9.23: numbering level pPr overrides the paragraph style.
    // Only the paragraph's direct ind overrides the numbering level.
    if let Some(ind) = effective_ind {
        merged_props.indentation = Dup::from(Some(ind));
    }
}

/// §17.9.23: the numbering level's indentation with the paragraph's *direct*
/// `w:ind` overrides applied field-by-field. `None` when the level defines no
/// indentation at all (the paragraph then keeps its style-cascade value).
fn effective_indentation(
    para: &model::Paragraph,
    level_def: Option<&crate::render::resolve::numbering::ResolvedNumberingLevel>,
) -> Option<model::Indentation> {
    let mut ind = *level_def?.indentation.as_ref()?;
    if let Some(direct) = para.properties.indentation.get() {
        if let Some(start) = direct.start {
            ind.start = Some(start);
        }
        if let Some(end) = direct.end {
            ind.end = Some(end);
        }
        if let Some(first_line) = direct.first_line {
            ind.first_line = Some(first_line);
        }
        if let Some(mirror) = direct.mirror {
            ind.mirror = Some(mirror);
        }
    }
    Some(ind)
}

/// §17.3.1.37 suppresses paragraph alignment on lines that contain tabs, but
/// the label's suffix tab is layout machinery, not typed content — Word still
/// centers "1.  Heading". For centered and right-aligned paragraphs the
/// separator is emitted as a fixed-advance spacer (see [`spacer_fragment`])
/// instead of a tab, so alignment applies to the whole line while the
/// label→text distance stays the hanging width, as in the left-aligned case.
fn alignment_suppresses_tab(merged_props: &ParagraphProperties) -> bool {
    matches!(
        merged_props.alignment.get(),
        Some(model::Alignment::Center | model::Alignment::End)
    )
}

/// Fixed-advance spacer standing in for the suffix tab on centered/right
/// lines.
///
/// - `trimmed_width` is zero: when nothing follows the label, the gap is
///   trailing whitespace, which per Word does not participate in alignment —
///   the label's glyphs center exactly, and on end-alignment they sit flush
///   at the margin with the gap hanging past it (see
///   [`Fragment::trimmed_width`]).
/// - Never underlined: a `<w:lvl><w:rPr>` underline covers the label's
///   glyphs, not the separator.
/// - A gap clamped to zero (no hanging indent, or a label wider than it)
///   still yields one space width, so the label never glues to the text.
fn spacer_fragment(
    gap: Pt,
    font: &crate::render::layout::fragment::FontProps,
    color: crate::render::resolve::color::RgbColor,
    measurer: &crate::render::layout::measurer::TextMeasurer,
) -> Fragment {
    let mut font = font.clone();
    font.underline = false;
    let (space_width, metrics) = measurer.measure(" ", &font);
    let gap = if gap > Pt::ZERO { gap } else { space_width };
    Fragment::Text {
        shaped: None,
        level: BidiLevel::LTR,
        text: Rc::from(""),
        font: Rc::new(font),
        color,
        shading: None,
        border: None,
        // §17.9.6: the label and the first line of its paragraph are one
        // line by construction — the gap exists so they don't touch, not so
        // the body can start on the line below.
        break_after: crate::render::layout::fragment::BreakAfter::Prohibited,
        width: gap,
        trimmed_width: Pt::ZERO,
        metrics,
        hyperlink_url: None,
        baseline_offset: Pt::ZERO,
        text_offset: Pt::ZERO,
        is_footnote_ref: false,
    }
}

/// Add the implicit tab stop at the effective text-indent position so the
/// label's suffix tab lands where the body text wraps (Word puts an implicit
/// stop at the paragraph's left indent for numbered paragraphs).
fn insert_implicit_tab_stop(
    merged_props: &mut ParagraphProperties,
    layout_ind: Option<model::Indentation>,
) {
    if let Some(left) = layout_ind.and_then(|ind| ind.start) {
        merged_props.tabs.insert(
            0,
            crate::model::TabStop {
                position: left,
                alignment: crate::model::TabAlignment::Left,
                leader: crate::model::TabLeader::None,
            },
        );
    }
}

/// Inject a text label (non-picture bullet) into the paragraph fragments.
#[allow(clippy::too_many_arguments)]
fn inject_text_label(
    para: &model::Paragraph,
    fragments: &mut Vec<Fragment>,
    merged_props: &mut ParagraphProperties,
    ctx: &BuildContext,
    counters: &std::collections::HashMap<(model::NumId, u8), u32>,
    levels: &[crate::render::resolve::numbering::ResolvedNumberingLevel],
    level: u8,
    level_def: Option<&crate::render::resolve::numbering::ResolvedNumberingLevel>,
    auto_fit: crate::render::layout::ShapeAutoFit,
    layout_ind: Option<model::Indentation>,
) {
    let num_id = model::NumId::new(merged_props.numbering.get().unwrap().num_id);

    let (default_family, default_size, default_color, _, paragraph_style_run) =
        resolve_paragraph_defaults(para, ctx.resolved, false, None, None);

    // §17.9.23 / §17.3.1.29: assemble the label's character-property
    // cascade. Order matters — level rPr beats paragraph-mark rPr beats
    // paragraph-style run defaults. Document defaults arrive below via
    // `default_family`/`default_size` scalars consumed by
    // `font_props_from_run`.
    let cascade = ListLabelRunPropertyCascade {
        level: level_def.and_then(|l| l.run_properties.as_ref()),
        paragraph_mark: para.mark_run_properties.as_ref(),
        paragraph_style: Some(&paragraph_style_run),
    };

    // §17.3.2.20: the label's language, read from the label's *own* cascade
    // rather than the paragraph's — a `<w:lvl><w:rPr><w:lang>` is as much a
    // property of the level as its font is, and §17.9.27's three text formats
    // are the only thing that consults it. Document defaults close the cascade,
    // the layer `default_family`/`default_size` carry for the other fields.
    let locale = crate::render::resolve::locale::Locale::from_cascade(
        cascade
            .iter()
            .chain(std::iter::once(&ctx.resolved.doc_defaults_run)),
    );

    let label_text = match crate::render::resolve::numbering::format_list_label(
        levels, level, counters, num_id, locale,
    ) {
        Some(t) => t,
        None => return,
    };

    // §17.3.2.6: color follows the same cascade as the font fields,
    // resolved as a separate scalar because it isn't part of
    // `FontProps`.
    let label_color = cascade
        .pick(|rp| rp.color.cloned())
        .map(|c| {
            crate::render::resolve::color::resolve_color(
                c,
                crate::render::resolve::color::ColorContext::Text,
            )
        })
        .unwrap_or(default_color);

    // Legacy Symbol/Wingdings remapping needs the cascade-resolved
    // family before deciding whether to remap PUA codepoints back to
    // their bullet glyphs.
    let cascade_family = cascade
        .iter()
        .find_map(|rp| crate::render::resolve::fonts::effective_font(&rp.fonts))
        .unwrap_or("");
    let (label_text, label_family) =
        remap_legacy_font_chars(&label_text, cascade_family, &default_family);

    // Build the label font through the canonical path so every
    // label-relevant field (underline, char_spacing, text_scale, etc.)
    // flows from cascade.resolve() into `font_props_from_run`. After
    // remapping, override the family if it changed.
    let mut label_font = build_label_font_props(&cascade, &default_family, default_size, auto_fit);
    if label_family != *label_font.family {
        label_font.family = Rc::from(label_family.as_str());
    }
    // §17.3.2.40: populate underline metrics from font metrics now that
    // the bool is settled — `populate_underline_metrics` (used by
    // `build_fragments`) ran before label injection, so this fragment
    // must populate its own metrics.
    populate_label_underline_metrics(&mut label_font, ctx.measurer);

    let (w, m) = ctx.measurer.measure(&label_text, &label_font);
    let h = m.height();

    let hanging = extract_hanging(layout_ind);
    // §17.9.7: lvlJc controls label justification within the hanging indent area.
    let jc = level_def.and_then(|l| l.justification);
    let text_offset = match jc {
        Some(crate::model::Alignment::End) => -w,
        Some(crate::model::Alignment::Center) => w * -0.5,
        _ => Pt::ZERO,
    };
    let label_width = w;
    let label_frag = Fragment::Text {
        shaped: None,
        level: BidiLevel::LTR,
        text: Rc::from(label_text.as_str()),
        font: Rc::new(label_font.clone()),
        color: label_color,
        shading: None,
        border: None,
        // The label leads its paragraph's first line; what separates it from
        // the body is the suffix that follows (tab, space or nothing), and
        // that fragment carries the break status.
        break_after: crate::render::layout::fragment::BreakAfter::Prohibited,
        width: label_width,
        trimmed_width: label_width,
        metrics: m,
        hyperlink_url: None,
        baseline_offset: Pt::ZERO,
        text_offset,
        is_footnote_ref: false,
    };
    // §17.9.29: the separator between the label and the body text depends on the
    // level's `suff`. Tab (default) advances to the body-text indent via a tab
    // stop; Space emits a single space; Nothing puts the text flush against the
    // label.
    use crate::model::LevelSuffix;
    match level_def.map(|l| l.suffix).unwrap_or_default() {
        LevelSuffix::Tab => {
            let gap = (hanging - label_width).max(Pt::ZERO);
            if alignment_suppresses_tab(merged_props) {
                fragments.insert(
                    0,
                    spacer_fragment(gap, &label_font, label_color, ctx.measurer),
                );
            } else {
                fragments.insert(
                    0,
                    Fragment::Tab {
                        line_height: h,
                        // §17.3.1.38: the label's own formatting is what a
                        // leader on this separator is drawn in.
                        font: Rc::new(label_font.clone()),
                        color: label_color,
                        fitting_width: Some(gap),
                    },
                );
            }
            fragments.insert(0, label_frag);
            // Implicit tab stop at the body-text indent. Also installed on the
            // spacer path: the line has no tab there, so the stop is inert
            // unless the run content itself contains one.
            insert_implicit_tab_stop(merged_props, layout_ind);
        }
        LevelSuffix::Space => {
            let (sw, sm) = ctx.measurer.measure(" ", &label_font);
            fragments.insert(
                0,
                Fragment::Text {
                    shaped: None,
                    level: BidiLevel::LTR,
                    text: Rc::from(" "),
                    font: Rc::new(label_font.clone()),
                    color: label_color,
                    shading: None,
                    border: None,
                    // §17.9.30 `w:suff` of `space`: a real space, so UAX #14
                    // breaks after it like any other.
                    break_after: crate::render::layout::fragment::BreakAfter::Opportunity,
                    width: sw,
                    trimmed_width: sw,
                    metrics: sm,
                    hyperlink_url: None,
                    baseline_offset: Pt::ZERO,
                    text_offset: Pt::ZERO,
                    is_footnote_ref: false,
                },
            );
            fragments.insert(0, label_frag);
        }
        LevelSuffix::Nothing => {
            fragments.insert(0, label_frag);
        }
    }
}

/// Extract the hanging indent from an effective indentation.
fn extract_hanging(ind: Option<model::Indentation>) -> Pt {
    ind.and_then(|ind| ind.first_line)
        .map(|fl| match fl {
            model::FirstLineIndent::Hanging(v) => Pt::from(v),
            _ => Pt::ZERO,
        })
        .unwrap_or(Pt::ZERO)
}

/// §17.9.23 — derive a label's [`FontProps`] from the resolved
/// character-property cascade. Delegates to [`font_props_from_run`]
/// so every label-relevant field (bold, italic, underline,
/// char_spacing, text_scale, font family, font size) takes the same
/// code path as ordinary text-run formatting. The returned
/// [`FontProps`] still has `underline_position`/`underline_thickness`
/// at zero — those come from font metrics and require a measurer.
/// Use [`populate_label_underline_metrics`] to fill them.
pub(super) fn build_label_font_props(
    cascade: &ListLabelRunPropertyCascade<'_>,
    default_family: &str,
    default_size: Pt,
    auto_fit: crate::render::layout::ShapeAutoFit,
) -> crate::render::layout::fragment::FontProps {
    let effective = cascade.resolve();
    crate::render::layout::fragment::font_props_from_run(
        &effective,
        default_family,
        default_size,
        auto_fit,
    )
}

/// Populate `underline_position` and `underline_thickness` on a
/// label's [`FontProps`] from the measurer's font metrics. No-op when
/// the label is not underlined. Idempotent.
pub(super) fn populate_label_underline_metrics(
    font: &mut crate::render::layout::fragment::FontProps,
    measurer: &crate::render::layout::measurer::TextMeasurer,
) {
    if font.underline {
        let (pos, thickness) = measurer.underline_metrics(font);
        font.underline_position = pos;
        font.underline_thickness = thickness;
    }
}

/// §17.9.23 / §17.3.1.29 — character-property cascade for a list
/// label's text. Sources are listed in priority order (highest first);
/// each field of the resolved properties is the topmost `Some` value
/// across `level → paragraph_mark → paragraph_style`.
///
/// The cascade is the spec-defined chain: a `<w:lvl><w:rPr>` overrides
/// the paragraph-mark `<w:pPr><w:rPr>`, which in turn overrides the
/// paragraph style's run defaults. Document defaults and theme
/// fallbacks come in below as scalar `default_family`/`default_size`
/// values supplied to `font_props_from_run`.
pub(super) struct ListLabelRunPropertyCascade<'a> {
    /// `<w:lvl><w:rPr>` — top of the cascade (§17.9.23).
    pub level: Option<&'a model::RunProperties>,
    /// `<w:pPr><w:rPr>` — formatting on the paragraph mark, applied
    /// where the level layer doesn't set a given field (§17.3.1.29).
    pub paragraph_mark: Option<&'a model::RunProperties>,
    /// Paragraph-style run defaults — bottom of the cascade.
    pub paragraph_style: Option<&'a model::RunProperties>,
}

impl<'a> ListLabelRunPropertyCascade<'a> {
    /// Iterate cascade sources in priority order, skipping `None`
    /// layers. Used by `pick` and by callers that need direct access
    /// to the cascade for non-`Copy` fields like `FontSet`.
    pub(super) fn iter(&self) -> impl Iterator<Item = &'a model::RunProperties> + '_ {
        [self.level, self.paragraph_mark, self.paragraph_style]
            .into_iter()
            .flatten()
    }

    /// Return the first `Some` value of a given field across the
    /// cascade, or `None` when no layer sets it. Restricted to `Copy`
    /// fields so the lookup is allocation-free; use `iter()` for fields
    /// that require cloning (e.g. `FontSet`).
    pub(super) fn pick<T: Copy>(
        &self,
        get: impl Fn(&model::RunProperties) -> Option<T>,
    ) -> Option<T> {
        self.iter().find_map(get)
    }

    /// Materialize a single effective `RunProperties` whose label-
    /// relevant fields (those consumed by `font_props_from_run` plus
    /// `color`) are picked from the highest-priority cascade source
    /// that sets each field. Fields no source sets are left at their
    /// default (`None` / empty `FontSet`) so downstream code can apply
    /// paragraph-level fallbacks.
    pub(super) fn resolve(&self) -> model::RunProperties {
        // `FontSet` is non-`Copy`. Pick from the first cascade source
        // whose font is explicitly set (any slot resolves under
        // `effective_font`); otherwise leave empty so downstream
        // `font_props_from_run` falls back to `default_family`.
        let fonts = self
            .iter()
            .find(|rp| crate::render::resolve::fonts::effective_font(&rp.fonts).is_some())
            .map(|rp| rp.fonts.clone())
            .unwrap_or_default();

        model::RunProperties {
            fonts,
            font_size: Dup::from(self.pick(|rp| rp.font_size.cloned())),
            bold: self.pick(|rp| rp.bold),
            italic: self.pick(|rp| rp.italic),
            underline: Dup::from(self.pick(|rp| rp.underline.cloned())),
            color: Dup::from(self.pick(|rp| rp.color.cloned())),
            spacing: Dup::from(self.pick(|rp| rp.spacing.cloned())),
            text_scale: Dup::from(self.pick(|rp| rp.text_scale.cloned())),
            ..Default::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::{Dimension, HalfPoints, Twips};
    use crate::model::{FontSet, FontSlot, RunProperties, TextScale, UnderlineStyle};
    use crate::render::fonts::Toggle;

    // ── §17.9.22 label emission ──────────────────────────────────────────
    //
    // The 16 tests below this block all exercise the property cascade. These
    // exercise the path that *uses* it: counter arithmetic, the §17.9.29
    // separator switch, §17.9.7 justification, and the hanging indent.

    use crate::model::{
        Alignment, FirstLineIndent, Indentation, LevelSuffix, NumId, NumberFormat,
        NumberingReference,
    };
    use crate::render::fonts::FontRegistry;
    use crate::render::layout::measurer::TextMeasurer;
    use crate::render::resolve::numbering::ResolvedNumberingLevel;
    use crate::render::resolve::ResolvedDocument;
    use std::collections::HashMap;

    fn level(format: NumberFormat, level_text: &str) -> ResolvedNumberingLevel {
        ResolvedNumberingLevel {
            format,
            level_text: level_text.into(),
            start: 1,
            run_properties: None,
            indentation: None,
            justification: None,
            lvl_pic_bullet_id: None,
            suffix: LevelSuffix::Tab,
            is_legal: false,
        }
    }

    /// A decimal level whose label is just its own counter.
    fn decimal_level() -> ResolvedNumberingLevel {
        level(NumberFormat::Decimal, "%1")
    }

    fn resolved_with(levels: Vec<ResolvedNumberingLevel>) -> ResolvedDocument {
        let mut numbering = HashMap::new();
        numbering.insert(NumId::new(7), levels);
        ResolvedDocument {
            sections: Vec::new(),
            styles: HashMap::new(),
            numbering,
            font_families: Vec::new(),
            media: HashMap::new(),
            embedded_fonts: Vec::new(),
            pic_bullets: HashMap::new(),
            theme: None,
            doc_defaults_paragraph: ParagraphProperties::default(),
            doc_defaults_run: RunProperties::default(),
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

    fn numbered_para() -> model::Paragraph {
        model::Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: Vec::new(),
            rsids: model::ParagraphRevisionIds::default(),
        }
    }

    fn props_at(level: u8) -> ParagraphProperties {
        ParagraphProperties {
            numbering: Dup::from(Some(NumberingReference { num_id: 7, level })),
            ..Default::default()
        }
    }

    /// Run `inject_list_label` over a live measurer, returning the fragments it
    /// prepended and the properties it rewrote.
    fn inject(
        resolved: &ResolvedDocument,
        state: &mut BuildState,
        level: u8,
    ) -> (Vec<Fragment>, ParagraphProperties) {
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved,
        };
        let mut fragments = Vec::new();
        let mut props = props_at(level);
        inject_list_label(&numbered_para(), &mut fragments, &mut props, &ctx, state);
        (fragments, props)
    }

    fn label_text(fragments: &[Fragment]) -> String {
        match fragments.first() {
            Some(Fragment::Text { text, .. }) => text.to_string(),
            other => panic!("expected a label fragment, got {other:?}"),
        }
    }

    /// §17.9.22: the counter seeds at `start` on the first item and advances by
    /// one thereafter. Seeding `start - 1` and adding one is the shape that
    /// underflows on `w:start="0"`, so this pins the seed as well as the step.
    #[test]
    fn the_counter_seeds_at_start_then_increments() {
        let resolved = resolved_with(vec![decimal_level()]);
        let mut state = BuildState::default();
        let labels: Vec<String> = (0..3)
            .map(|_| label_text(&inject(&resolved, &mut state, 0).0))
            .collect();
        assert_eq!(labels, ["1", "2", "3"]);
    }

    /// §17.9.28 permits `w:start="0"`, which the seed must survive — it
    /// underflows `u32` if the counter is seeded one below `start`.
    #[test]
    fn a_zero_start_does_not_underflow() {
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            start: 0,
            ..decimal_level()
        }]);
        let mut state = BuildState::default();
        let labels: Vec<String> = (0..2)
            .map(|_| label_text(&inject(&resolved, &mut state, 0).0))
            .collect();
        assert_eq!(labels, ["0", "1"]);
    }

    /// §17.9.22: an item at one level restarts every level *below* it, so a
    /// nested list numbers `1 / 1 / 2 / 1` rather than carrying the sub-counter
    /// across. Levels *above* are untouched.
    #[test]
    fn a_deeper_level_restarts_when_its_parent_advances() {
        // `%N` names level N−1's counter, not "this level's" — so the level-1
        // template has to be `%2` to show its own number. With `%1` the test
        // reads level 0's counter and passes against a broken reset.
        let resolved = resolved_with(vec![decimal_level(), level(NumberFormat::Decimal, "%2")]);
        let mut state = BuildState::default();

        // Evaluated in order, each call advancing `state` — the sequence is
        // the assertion.
        let seen = vec![
            label_text(&inject(&resolved, &mut state, 0).0), // "1"
            label_text(&inject(&resolved, &mut state, 1).0), // "1" — first sub-item
            label_text(&inject(&resolved, &mut state, 1).0), // "2"
            label_text(&inject(&resolved, &mut state, 0).0), // "2" — resets level 1
            label_text(&inject(&resolved, &mut state, 1).0), // "1" again
        ];

        assert_eq!(seen, ["1", "1", "2", "2", "1"]);
    }

    /// Like `inject`, but with a caller-supplied paragraph and properties —
    /// for tests exercising direct `w:ind` overrides and paragraph alignment.
    fn inject_full(
        resolved: &ResolvedDocument,
        state: &mut BuildState,
        para: &model::Paragraph,
        mut props: ParagraphProperties,
    ) -> (Vec<Fragment>, ParagraphProperties) {
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved,
        };
        let mut fragments = Vec::new();
        inject_list_label(para, &mut fragments, &mut props, &ctx, state);
        (fragments, props)
    }

    fn ind(left: i64, hanging: i64) -> Indentation {
        Indentation {
            start: Some(Dimension::<Twips>::new(left)),
            end: None,
            first_line: Some(FirstLineIndent::Hanging(Dimension::<Twips>::new(hanging))),
            mirror: None,
        }
    }

    fn para_with(props: ParagraphProperties) -> model::Paragraph {
        model::Paragraph {
            style_id: None,
            properties: props,
            mark_run_properties: None,
            content: Vec::new(),
            rsids: model::ParagraphRevisionIds::default(),
        }
    }

    /// §17.9.23: the paragraph's direct `w:ind` overrides the numbering
    /// level's — for the implicit tab stop too, not just for body-text
    /// wrapping. With level left=2631 and direct left=567, the suffix tab
    /// must land at 567, or the first line renders with a huge gap while
    /// continuation lines wrap at the direct indent.
    #[test]
    fn direct_ind_overrides_level_for_the_implicit_tab_stop() {
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            indentation: Some(ind(2631, 504)),
            ..decimal_level()
        }]);
        let para = para_with(ParagraphProperties {
            indentation: Dup::from(Some(ind(567, 567))),
            ..Default::default()
        });
        let (_, props) = inject_full(&resolved, &mut BuildState::default(), &para, props_at(0));

        assert_eq!(
            props.tabs.first().map(|t| t.position.raw()),
            Some(567),
            "implicit tab stop must sit at the direct indent, not the level's"
        );
        let merged = props.indentation.get().expect("indentation merged");
        assert_eq!(merged.start.map(|d| d.raw()), Some(567));
        assert_eq!(
            merged.first_line,
            Some(FirstLineIndent::Hanging(Dimension::<Twips>::new(567)))
        );
    }

    /// Word semantics: an item at a deep level instantiates its ancestor
    /// counters at their start values without consuming them — after
    /// "1.1"-style items the next top-level item continues as "2", not "1".
    #[test]
    fn deep_level_items_instantiate_ancestor_counters() {
        let resolved = resolved_with(vec![decimal_level(), level(NumberFormat::Decimal, "%1.%2")]);
        let mut state = BuildState::default();

        let seen = vec![
            label_text(&inject(&resolved, &mut state, 1).0), // "1.1"
            label_text(&inject(&resolved, &mut state, 1).0), // "1.2"
            label_text(&inject(&resolved, &mut state, 0).0), // "2" — not "1"
        ];
        assert_eq!(seen, ["1.1", "1.2", "2"]);
    }

    /// §17.3.1.37 suppresses paragraph alignment on lines with tabs, but the
    /// label's suffix tab is layout machinery — a centered numbered heading
    /// must stay centered. For center/end alignment the separator is a
    /// fixed-advance spacer, not a tab: label advance + spacer advance equal
    /// the hanging width (label starts at left−hanging, text at left, as in
    /// the left-aligned case), and the spacer's `trimmed_width` is zero so a
    /// label alone on its line aligns by its glyphs, the gap hanging like
    /// trailing whitespace.
    #[test]
    fn centered_paragraph_emits_spacer_instead_of_tab() {
        for alignment in [Alignment::Center, Alignment::End] {
            let resolved = resolved_with(vec![ResolvedNumberingLevel {
                indentation: Some(ind(360, 360)),
                ..decimal_level()
            }]);
            let props = ParagraphProperties {
                alignment: Dup::from(Some(alignment)),
                ..props_at(0)
            };
            let (fragments, props) = inject_full(
                &resolved,
                &mut BuildState::default(),
                &numbered_para(),
                props,
            );

            assert_eq!(fragments.len(), 2, "{alignment:?}: label + spacer");
            let label_width = match &fragments[0] {
                Fragment::Text { width, .. } => *width,
                other => panic!("{alignment:?}: expected label text fragment, got {other:?}"),
            };
            match &fragments[1] {
                Fragment::Text {
                    text,
                    width,
                    trimmed_width,
                    ..
                } => {
                    assert!(text.is_empty(), "{alignment:?}: spacer carries no glyphs");
                    // 360 twips = 18 pt total advance.
                    assert!(
                        ((label_width.raw() + width.raw()) - 18.0).abs() < 0.01,
                        "{alignment:?}: label + spacer must advance the hanging width"
                    );
                    assert_eq!(
                        trimmed_width.raw(),
                        0.0,
                        "{alignment:?}: the gap is trailing whitespace when nothing follows"
                    );
                }
                other => panic!("{alignment:?}: expected spacer text fragment, got {other:?}"),
            }
            // The implicit stop is installed but inert — the line has no tab.
            assert_eq!(
                props.tabs.first().map(|t| t.position.raw()),
                Some(360),
                "{alignment:?}: implicit stop still installed"
            );
        }
    }

    /// A zero hanging indent (or a label wider than it) must not glue the
    /// label to the body text on the spacer path: the gap floors at one
    /// space width of the label font.
    #[test]
    fn centered_spacer_floors_at_a_space_width_when_hanging_is_zero() {
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            indentation: Some(ind(360, 0)),
            ..decimal_level()
        }]);
        let props = ParagraphProperties {
            alignment: Dup::from(Some(Alignment::Center)),
            ..props_at(0)
        };
        let (fragments, _) = inject_full(
            &resolved,
            &mut BuildState::default(),
            &numbered_para(),
            props,
        );
        match &fragments[1] {
            Fragment::Text { width, .. } => {
                assert!(
                    width.raw() > 0.0,
                    "spacer must keep a positive advance, got {width:?}"
                );
            }
            other => panic!("expected spacer, got {other:?}"),
        }
    }

    /// A `<w:lvl><w:rPr>` underline covers the label's glyphs only — the
    /// spacer standing in for the suffix tab must not stretch the underline
    /// across the gap.
    #[test]
    fn centered_spacer_is_never_underlined() {
        let level_rpr = RunProperties {
            underline: Dup::from(Some(UnderlineStyle::Single)),
            ..Default::default()
        };
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            indentation: Some(ind(360, 360)),
            run_properties: Some(level_rpr),
            ..decimal_level()
        }]);
        let props = ParagraphProperties {
            alignment: Dup::from(Some(Alignment::Center)),
            ..props_at(0)
        };
        let (fragments, _) = inject_full(
            &resolved,
            &mut BuildState::default(),
            &numbered_para(),
            props,
        );
        match (&fragments[0], &fragments[1]) {
            (Fragment::Text { font: label, .. }, Fragment::Text { font: spacer, .. }) => {
                assert!(label.underline, "precondition: the label is underlined");
                assert!(!spacer.underline, "the spacer must not be underlined");
            }
            other => panic!("expected label + spacer, got {other:?}"),
        }
    }

    /// A crafted `w:ilvl="255"` must not overflow the deeper-level reset
    /// range (`level + 1` would panic in debug and wrap in release, wiping
    /// every counter for the numId — including seeded ancestors).
    #[test]
    fn ilvl_255_does_not_overflow_or_wipe_counters() {
        let resolved = resolved_with(vec![decimal_level()]);
        let mut state = BuildState::default();
        assert_eq!(label_text(&inject(&resolved, &mut state, 0).0), "1");
        // No level 255 exists: no label, and — critically — no panic and no
        // counter wipe.
        let (fragments, _) = inject(&resolved, &mut state, 255);
        assert!(fragments.is_empty(), "no label for an undefined level");
        assert_eq!(
            label_text(&inject(&resolved, &mut state, 0).0),
            "2",
            "the level-0 counter must survive the crafted ilvl"
        );
    }

    /// §17.9.23 direct override carries `w:mirrorIndents` too, not just the
    /// start/end/first-line fields.
    #[test]
    fn direct_mirror_override_survives_the_level_merge() {
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            indentation: Some(ind(2631, 504)),
            ..decimal_level()
        }]);
        let para = para_with(ParagraphProperties {
            indentation: Dup::from(Some(Indentation {
                mirror: Some(true),
                ..ind(567, 567)
            })),
            ..Default::default()
        });
        let (_, props) = inject_full(&resolved, &mut BuildState::default(), &para, props_at(0));
        assert_eq!(props.indentation.get().unwrap().mirror, Some(true));
    }

    /// When the numbering level defines no `w:ind`, label geometry falls back
    /// to the paragraph's cascade indentation: the implicit tab stop lands at
    /// its start and the tab's fitting width reflects its hanging — while the
    /// §17.9.23 overwrite is skipped (the cascade value is already in place).
    #[test]
    fn level_without_ind_falls_back_to_cascade_indentation() {
        let resolved = resolved_with(vec![decimal_level()]); // no level ind
        let props = ParagraphProperties {
            indentation: Dup::from(Some(ind(720, 360))),
            ..props_at(0)
        };
        let (fragments, props) = inject_full(
            &resolved,
            &mut BuildState::default(),
            &numbered_para(),
            props,
        );

        assert_eq!(
            props.tabs.first().map(|t| t.position.raw()),
            Some(720),
            "implicit stop must land at the cascade start"
        );
        let label_width = match &fragments[0] {
            Fragment::Text { width, .. } => *width,
            other => panic!("expected label, got {other:?}"),
        };
        match &fragments[1] {
            Fragment::Tab { fitting_width, .. } => {
                // 360 twips = 18 pt hanging.
                let expected = (Pt::new(18.0) - label_width).max(Pt::ZERO);
                assert!(
                    (fitting_width.unwrap().raw() - expected.raw()).abs() < 0.01,
                    "fitting width must reflect the cascade hanging"
                );
            }
            other => panic!("expected separator tab, got {other:?}"),
        }
        // The cascade value stays as-is; the §17.9.23 overwrite is for
        // level-derived indentation only.
        assert_eq!(
            props.indentation.get().unwrap().start.map(|d| d.raw()),
            Some(720)
        );
    }

    /// The picture-bullet branch honors alignment suppression too: a centered
    /// picture-bulleted item emits a spacer, not a tab, so the line keeps its
    /// alignment.
    #[test]
    fn centered_picture_bullet_emits_spacer_instead_of_tab() {
        use crate::model::{
            NumPicBullet, NumPicBulletId, Pict, RelId, VmlCommonAttrs, VmlImageData, VmlPrimitive,
            VmlShape,
        };
        let mut resolved = resolved_with(vec![ResolvedNumberingLevel {
            lvl_pic_bullet_id: Some(NumPicBulletId::new(1)),
            indentation: Some(ind(360, 360)),
            ..decimal_level()
        }]);
        resolved.pic_bullets.insert(
            NumPicBulletId::new(1),
            NumPicBullet {
                id: NumPicBulletId::new(1),
                pict: Some(Pict {
                    shape_type: None,
                    primitives: vec![VmlPrimitive::Shape(VmlShape {
                        common: VmlCommonAttrs {
                            image_data: Some(VmlImageData {
                                rel_id: Some(RelId::new("rId9")),
                                title: None,
                            }),
                            ..Default::default()
                        },
                        shape_type_ref: None,
                        vml_path: None,
                    })],
                }),
            },
        );
        resolved.media.insert(
            RelId::new("rId9"),
            crate::render::resolve::images::MediaEntry {
                data: std::sync::Arc::from(&b"\x89PNG"[..]),
                format: crate::model::ImageFormat::Png,
            },
        );

        for (alignment, expect_tab) in [
            (Some(Alignment::Center), false),
            (Some(Alignment::End), false),
            (None, true),
        ] {
            let props = ParagraphProperties {
                alignment: Dup::from(alignment),
                ..props_at(0)
            };
            let (fragments, _) = inject_full(
                &resolved,
                &mut BuildState::default(),
                &numbered_para(),
                props,
            );
            assert!(
                matches!(fragments[0], Fragment::Image { .. }),
                "{alignment:?}: picture bullet first"
            );
            match (&fragments[1], expect_tab) {
                (Fragment::Tab { .. }, true) => {}
                (Fragment::Text { text, width, .. }, false) => {
                    assert!(text.is_empty(), "{alignment:?}: spacer carries no glyphs");
                    assert!(width.raw() > 0.0, "{alignment:?}: spacer has an advance");
                }
                (other, expected) => {
                    panic!("{alignment:?}: expected tab={expected}, got {other:?}")
                }
            }
        }
    }

    /// §17.9.29: `Tab` (the default) puts a tab after the label, `Space` a
    /// literal space, `Nothing` neither.
    #[test]
    fn the_suffix_switch_picks_the_separator() {
        for (suffix, expect) in [
            (LevelSuffix::Tab, "tab"),
            (LevelSuffix::Space, "space"),
            (LevelSuffix::Nothing, "none"),
        ] {
            let resolved = resolved_with(vec![ResolvedNumberingLevel {
                suffix,
                ..decimal_level()
            }]);
            let (fragments, _) = inject(&resolved, &mut BuildState::default(), 0);
            assert_eq!(label_text(&fragments), "1", "{suffix:?}");

            let got = match fragments.get(1) {
                None => "none",
                Some(Fragment::Tab { .. }) => "tab",
                Some(Fragment::Text { text, .. }) if &**text == " " => "space",
                other => panic!("{suffix:?} emitted {other:?}"),
            };
            assert_eq!(got, expect, "{suffix:?}");
        }
    }

    /// §17.9.29: only the `Tab` suffix installs the implicit tab stop at the
    /// level's `start` indent — a `Space`/`Nothing` label has no tab to land.
    #[test]
    fn only_a_tab_suffix_installs_the_implicit_tab_stop() {
        for (suffix, expect_stop) in [
            (LevelSuffix::Tab, true),
            (LevelSuffix::Space, false),
            (LevelSuffix::Nothing, false),
        ] {
            let resolved = resolved_with(vec![ResolvedNumberingLevel {
                suffix,
                indentation: Some(Indentation {
                    start: Some(Dimension::new(720)),
                    ..Default::default()
                }),
                ..decimal_level()
            }]);
            let (_, props) = inject(&resolved, &mut BuildState::default(), 0);
            assert_eq!(
                props.tabs.iter().any(|t| t.position == Dimension::new(720)),
                expect_stop,
                "{suffix:?}"
            );
        }
    }

    /// §17.9.7 `lvlJc`: the label is placed by shifting it within the hanging
    /// indent area — `start` not at all, `end` by its whole width, `center` by
    /// half. Expressed as a ratio so the host font's metrics cancel.
    #[test]
    fn lvl_jc_offsets_the_label_by_its_own_width() {
        let offset_for = |jc: Option<Alignment>| {
            let resolved = resolved_with(vec![ResolvedNumberingLevel {
                justification: jc,
                ..decimal_level()
            }]);
            let (fragments, _) = inject(&resolved, &mut BuildState::default(), 0);
            match fragments.first() {
                Some(Fragment::Text {
                    text_offset, width, ..
                }) => (text_offset.raw(), width.raw()),
                other => panic!("expected a label, got {other:?}"),
            }
        };

        let (start, w) = offset_for(Some(Alignment::Start));
        assert_eq!(start, 0.0, "start-justified labels are not shifted");
        assert!(w > 0.0, "the label has a measured width");

        let (end, _) = offset_for(Some(Alignment::End));
        assert!((end + w).abs() < 1e-3, "end shifts left by the full width");

        let (centre, _) = offset_for(Some(Alignment::Center));
        assert!((centre + w * 0.5).abs() < 1e-3, "center shifts by half");

        let (absent, _) = offset_for(None);
        assert_eq!(absent, 0.0, "an absent lvlJc behaves as start");
    }

    /// §17.9.3: the tab after the label is asked to span the hanging indent
    /// *less* what the label already occupies, so the body text lands at the
    /// level's indent rather than one label-width past it.
    #[test]
    fn the_separator_tab_spans_the_hanging_indent_less_the_label() {
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            indentation: Some(Indentation {
                start: Some(Dimension::new(720)),
                first_line: Some(FirstLineIndent::Hanging(Dimension::new(360))), // 18pt
                ..Default::default()
            }),
            ..decimal_level()
        }]);
        let (fragments, _) = inject(&resolved, &mut BuildState::default(), 0);

        let label_width = match &fragments[0] {
            Fragment::Text { width, .. } => width.raw(),
            other => panic!("expected a label, got {other:?}"),
        };
        match &fragments[1] {
            Fragment::Tab { fitting_width, .. } => {
                let got = fitting_width
                    .expect("the label tab carries a fitting width")
                    .raw();
                assert!(
                    (got - (18.0 - label_width)).abs() < 1e-3,
                    "18pt hanging indent less a {label_width}pt label, got {got}"
                );
            }
            other => panic!("expected the separator tab, got {other:?}"),
        }
    }

    /// `extract_hanging` reads only the hanging case: a *first-line* indent is
    /// not a hanging indent, and neither is an absent one. Observed here
    /// through the tab's fitting width; the centered/right-aligned path
    /// consumes the same value as the spacer fragment's advance instead.
    #[test]
    fn only_a_hanging_first_line_indent_is_a_hanging_indent() {
        let fitting_for = |first_line: Option<FirstLineIndent>| {
            let resolved = resolved_with(vec![ResolvedNumberingLevel {
                indentation: Some(Indentation {
                    first_line,
                    ..Default::default()
                }),
                ..decimal_level()
            }]);
            let (fragments, _) = inject(&resolved, &mut BuildState::default(), 0);
            match &fragments[1] {
                Fragment::Tab { fitting_width, .. } => fitting_width.unwrap().raw(),
                other => panic!("expected the separator tab, got {other:?}"),
            }
        };

        // A hanging indent gives the tab something to span; the other two
        // clamp to zero, because `hanging - label_width` goes negative.
        assert!(fitting_for(Some(FirstLineIndent::Hanging(Dimension::new(720)))) > 0.0);
        assert_eq!(
            fitting_for(Some(FirstLineIndent::FirstLine(Dimension::new(720)))),
            0.0
        );
        assert_eq!(fitting_for(Some(FirstLineIndent::None)), 0.0);
        assert_eq!(fitting_for(None), 0.0);
    }

    /// §17.9.22: a paragraph with no numbering reference is left completely
    /// alone — no label, no rewritten indentation, no counter touched.
    #[test]
    fn a_paragraph_without_numbering_is_untouched() {
        let resolved = resolved_with(vec![decimal_level()]);
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut state = BuildState::default();
        let mut fragments = Vec::new();
        let mut props = ParagraphProperties::default();
        inject_list_label(
            &numbered_para(),
            &mut fragments,
            &mut props,
            &ctx,
            &mut state,
        );

        assert!(fragments.is_empty(), "no label injected");
        assert!(props.indentation.get().is_none(), "indentation untouched");
        assert!(state.list_counters.is_empty(), "no counter advanced");
    }

    /// §17.9.2: a `numId` with no definition is a dangling reference. It must
    /// not inject a label *or* advance a counter — the paragraph simply is not
    /// numbered.
    #[test]
    fn an_unknown_num_id_injects_nothing() {
        let resolved = resolved_with(vec![decimal_level()]);
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut state = BuildState::default();
        let mut fragments = Vec::new();
        let mut props = ParagraphProperties {
            numbering: Dup::from(Some(NumberingReference {
                num_id: 999,
                level: 0,
            })),
            ..Default::default()
        };
        inject_list_label(
            &numbered_para(),
            &mut fragments,
            &mut props,
            &ctx,
            &mut state,
        );

        assert!(fragments.is_empty(), "no label injected");
        assert!(state.list_counters.is_empty(), "no counter advanced");
    }

    /// §17.9.23: the level's indentation replaces the paragraph's, but the
    /// paragraph's *direct* indentation still wins field by field.
    #[test]
    fn direct_paragraph_indentation_beats_the_level() {
        let resolved = resolved_with(vec![ResolvedNumberingLevel {
            indentation: Some(Indentation {
                start: Some(Dimension::new(720)),
                end: Some(Dimension::new(100)),
                ..Default::default()
            }),
            ..decimal_level()
        }]);
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut para = numbered_para();
        para.properties.indentation = Dup::from(Some(Indentation {
            start: Some(Dimension::new(1440)),
            ..Default::default()
        }));
        let mut fragments = Vec::new();
        let mut props = props_at(0);
        inject_list_label(
            &para,
            &mut fragments,
            &mut props,
            &ctx,
            &mut BuildState::default(),
        );

        let ind = props
            .indentation
            .get()
            .expect("the level supplies indentation");
        assert_eq!(ind.start, Some(Dimension::new(1440)), "direct start wins");
        assert_eq!(ind.end, Some(Dimension::new(100)), "level end survives");
    }

    fn rp_with_bold(b: bool) -> RunProperties {
        RunProperties {
            bold: Some(b),
            ..Default::default()
        }
    }

    fn rp_with_underline(u: UnderlineStyle) -> RunProperties {
        RunProperties {
            underline: Dup::from(Some(u)),
            ..Default::default()
        }
    }

    fn rp_with_font(name: &str) -> RunProperties {
        RunProperties {
            fonts: FontSet {
                ascii: FontSlot::from_name(name),
                ..Default::default()
            },
            ..Default::default()
        }
    }

    /// Highest-priority source wins when more than one layer sets the
    /// same field.
    #[test]
    fn cascade_pick_level_overrides_mark_for_bold() {
        let level = rp_with_bold(true);
        let mark = rp_with_bold(false);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: Some(&mark),
            paragraph_style: None,
        };
        assert_eq!(cascade.pick(|rp| rp.bold), Some(true));
    }

    /// When the level layer doesn't set a field, the next layer down
    /// supplies it. §17.3.1.29.
    #[test]
    fn cascade_pick_falls_through_to_mark_when_level_field_absent() {
        let level = rp_with_bold(true); // sets bold but NOT underline
        let mark = rp_with_underline(UnderlineStyle::Single);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: Some(&mark),
            paragraph_style: None,
        };
        assert_eq!(cascade.pick(|rp| rp.bold), Some(true));
        assert_eq!(
            cascade.pick(|rp| rp.underline.cloned()),
            Some(UnderlineStyle::Single)
        );
    }

    /// Paragraph-style run defaults are the bottom layer — used only
    /// when neither level nor mark sets a field.
    #[test]
    fn cascade_pick_falls_through_to_paragraph_style_when_others_absent() {
        let style = rp_with_underline(UnderlineStyle::Double);
        let cascade = ListLabelRunPropertyCascade {
            level: None,
            paragraph_mark: None,
            paragraph_style: Some(&style),
        };
        assert_eq!(
            cascade.pick(|rp| rp.underline.cloned()),
            Some(UnderlineStyle::Double)
        );
    }

    /// No layer sets the field → `None`. Caller decides the default.
    #[test]
    fn cascade_pick_returns_none_when_no_source_sets_field() {
        let level = rp_with_bold(true);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None,
            paragraph_style: None,
        };
        assert_eq!(cascade.pick(|rp| rp.italic), None);
    }

    /// `iter()` skips `None` layers and preserves priority order.
    #[test]
    fn cascade_iter_skips_none_layers_in_order() {
        let level = rp_with_bold(true);
        let style = rp_with_bold(false);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None, // skipped
            paragraph_style: Some(&style),
        };
        let bolds: Vec<_> = cascade.iter().map(|rp| rp.bold).collect();
        assert_eq!(bolds, vec![Some(true), Some(false)]);
    }

    /// The key correctness property for the bug at hand: a level rPr
    /// underline must surface in the resolved properties so
    /// `font_props_from_run` reads it.
    #[test]
    fn cascade_resolve_materializes_underline_from_level() {
        let level = rp_with_underline(UnderlineStyle::Single);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None,
            paragraph_style: None,
        };
        let effective = cascade.resolve();
        assert_eq!(effective.underline, Dup::from(Some(UnderlineStyle::Single)));
    }

    /// `FontSet` (non-`Copy`) cascade rule: the resolved `fonts` is
    /// the first cascade source that has an explicitly set font slot.
    #[test]
    fn cascade_resolve_fonts_picks_first_explicit_source() {
        let level = RunProperties::default(); // no fonts
        let mark = rp_with_font("Verdana");
        let style = rp_with_font("Arial");
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: Some(&mark),
            paragraph_style: Some(&style),
        };
        let effective = cascade.resolve();
        assert_eq!(effective.fonts.ascii.explicit.as_deref(), Some("Verdana"));
    }

    // ── §17.9.23 — `build_label_font_props` ──────────────────────────────
    //
    // The label's `FontProps` is built by passing the cascade-resolved
    // `RunProperties` through `font_props_from_run`. These tests pin
    // the spec-driven invariants the previous hand-rolled construction
    // violated (notably: dropped underline, char_spacing, text_scale).

    /// THE BUG: an underline in the level rPr (`<w:lvl><w:rPr><w:u/>`)
    /// must surface as `font.underline = true`. Previously this code
    /// path hardcoded `underline: false`.
    #[test]
    fn label_font_inherits_underline_from_level() {
        let level = rp_with_underline(UnderlineStyle::Single);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None,
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(
            font.underline,
            "level rPr <w:u/> must become font.underline"
        );
    }

    /// Cascade depth: when the level layer doesn't set underline but
    /// the paragraph-mark rPr does, the label inherits the underline.
    /// §17.3.1.29. (Previously broken: hand-rolled code never read
    /// either layer's underline.)
    #[test]
    fn label_font_inherits_underline_from_mark_when_level_absent() {
        let mark = rp_with_underline(UnderlineStyle::Single);
        let cascade = ListLabelRunPropertyCascade {
            level: None,
            paragraph_mark: Some(&mark),
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(font.underline);
    }

    /// Explicit `<w:u w:val="none"/>` in the level must turn underline
    /// OFF even if a lower cascade layer has it on — per spec, "none"
    /// is an explicit override, not an "inherit" signal. The cascade
    /// picks the level's `Some(None)` over the mark's `Some(Single)`,
    /// and `font_props_from_run` collapses it to `underline = false`.
    #[test]
    fn label_font_underline_false_when_level_explicitly_none() {
        let level = rp_with_underline(UnderlineStyle::None);
        let mark = rp_with_underline(UnderlineStyle::Single);
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: Some(&mark),
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(
            !font.underline,
            "explicit UnderlineStyle::None must override lower layers"
        );
    }

    /// `<w:spacing>` (char spacing) was silently dropped by the
    /// hand-rolled construction. With cascade routing it must flow
    /// into `FontProps::char_spacing`.
    #[test]
    fn label_font_inherits_char_spacing_from_level() {
        let level = RunProperties {
            spacing: Dup::from(Some(Dimension::<Twips>::new(40))),
            ..Default::default()
        };
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None,
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        // 40 twips = 2 pt.
        assert!((font.char_spacing.raw() - 2.0).abs() < 1e-4);
    }

    /// `<w:w>` (text scale) was silently fixed to 1.0 in the
    /// hand-rolled construction.
    #[test]
    fn label_font_inherits_text_scale_from_level() {
        let level = RunProperties {
            text_scale: Dup::from(Some(TextScale::new(150))),
            ..Default::default()
        };
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None,
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!((font.text_scale - 1.5).abs() < 1e-4);
    }

    /// Bold/italic/size pass through, with the cascade-resolved font
    /// family taking precedence over `default_family`.
    #[test]
    fn label_font_pass_through_basic_fields() {
        let level = RunProperties {
            bold: Some(true),
            italic: Some(true),
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(24))), // 12 pt
            fonts: FontSet {
                ascii: FontSlot::from_name("Verdana"),
                ..Default::default()
            },
            ..Default::default()
        };
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: None,
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(10.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert_eq!(font.bold, Toggle::On);
        assert_eq!(font.italic, Toggle::On);
        assert_eq!(font.size.raw(), 12.0);
        assert_eq!(&*font.family, "Verdana");
    }

    /// When no cascade layer sets a font, `default_family` wins.
    #[test]
    fn label_font_falls_back_to_default_family() {
        let cascade = ListLabelRunPropertyCascade {
            level: None,
            paragraph_mark: None,
            paragraph_style: None,
        };
        let font = build_label_font_props(
            &cascade,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert_eq!(&*font.family, "Helvetica");
    }

    /// Empty cascade — resolve returns default (all `None` / empty).
    #[test]
    fn cascade_resolve_returns_defaults_when_cascade_empty() {
        let cascade = ListLabelRunPropertyCascade {
            level: None,
            paragraph_mark: None,
            paragraph_style: None,
        };
        let effective = cascade.resolve();
        assert_eq!(effective, RunProperties::default());
    }

    /// All label-relevant fields composed across all three layers.
    /// Documents the full set of fields the resolver pulls (any
    /// future field becoming label-relevant should fail this test
    /// until added).
    #[test]
    fn cascade_resolve_composes_all_label_fields() {
        use crate::model::Color;

        let level = RunProperties {
            bold: Some(true),
            underline: Dup::from(Some(UnderlineStyle::Single)),
            ..Default::default()
        };
        let mark = RunProperties {
            italic: Some(true),
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(24))),
            spacing: Dup::from(Some(Dimension::<Twips>::new(40))),
            ..Default::default()
        };
        let style = RunProperties {
            color: Dup::from(Some(Color::Rgb(0x112233))),
            text_scale: Dup::from(Some(TextScale::new(120))),
            fonts: FontSet {
                ascii: FontSlot::from_name("Calibri"),
                ..Default::default()
            },
            ..Default::default()
        };
        let cascade = ListLabelRunPropertyCascade {
            level: Some(&level),
            paragraph_mark: Some(&mark),
            paragraph_style: Some(&style),
        };
        let effective = cascade.resolve();
        assert_eq!(effective.bold, Some(true), "from level");
        assert_eq!(
            effective.underline,
            Dup::from(Some(UnderlineStyle::Single)),
            "from level"
        );
        assert_eq!(effective.italic, Some(true), "from mark");
        assert_eq!(
            effective.font_size,
            Dup::from(Some(Dimension::<HalfPoints>::new(24))),
            "from mark"
        );
        assert_eq!(
            effective.spacing,
            Dup::from(Some(Dimension::<Twips>::new(40))),
            "from mark"
        );
        assert_eq!(
            effective.color,
            Dup::from(Some(Color::Rgb(0x112233))),
            "from style"
        );
        assert_eq!(
            effective.text_scale,
            Dup::from(Some(TextScale::new(120))),
            "from style"
        );
        assert_eq!(
            effective.fonts.ascii.explicit.as_deref(),
            Some("Calibri"),
            "from style (lowest source that sets fonts)"
        );
    }
}
