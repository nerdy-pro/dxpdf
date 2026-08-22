//! Option-field merging for RunProperties and ParagraphProperties.
//!
//! "Merge" means: for each field, if `self` is `None`, take the value from `base`.
//! This implements the OOXML style inheritance cascade.

use crate::model::Dup;
use crate::model::{
    ParagraphProperties, RunProperties, TabAlignment, TableCellProperties, TableProperties,
    TableStyleOverride,
};

/// Fill any `None` fields in `$target` from the corresponding fields in `$base`.
///
/// Expands to a sequence of `merge_opt(&mut $target.$field, &$base.$field)` calls.
macro_rules! merge_fields {
    ($target:expr, $base:expr, $($field:ident),+ $(,)?) => {
        $(merge_opt(&mut $target.$field, &$base.$field);)+
    };
}

/// Merge `base` into `target`: any `None` field in `target` gets filled from `base`.
pub fn merge_run_properties(target: &mut RunProperties, base: &RunProperties) {
    // FontSet: delegate to FontSlot::merge_from so each slot handles its own fields.
    target.fonts.ascii.merge_from(&base.fonts.ascii);
    target.fonts.high_ansi.merge_from(&base.fonts.high_ansi);
    target.fonts.east_asian.merge_from(&base.fonts.east_asian);
    target
        .fonts
        .complex_script
        .merge_from(&base.fonts.complex_script);

    merge_fields!(
        target,
        base,
        font_size,
        bold,
        italic,
        underline,
        strike,
        color,
        highlight,
        shading,
        vertical_align,
        spacing,
        kerning,
        all_caps,
        small_caps,
        vanish,
        no_proof,
        web_hidden,
        rtl,
        emboss,
        imprint,
        outline,
        shadow,
        position,
        lang,
        border,
        text_scale,
    );
}

/// Merge `base` into `target`: any `None` field in `target` gets filled from `base`.
pub fn merge_paragraph_properties(target: &mut ParagraphProperties, base: &ParagraphProperties) {
    merge_opt(&mut target.alignment, &base.alignment);
    // §17.3.1.12: merge indentation sub-fields individually so partial
    // overrides (e.g., left from child style, firstLine from parent) combine.
    // Both sides resolve to their effective occurrence first: the cascade
    // combines the values the document means, not every one it wrote.
    match (target.indentation.get_mut(), base.indentation.get()) {
        (Some(ti), Some(bi)) => {
            merge_opt(&mut ti.start, &bi.start);
            merge_opt(&mut ti.end, &bi.end);
            merge_opt(&mut ti.first_line, &bi.first_line);
            merge_opt(&mut ti.mirror, &bi.mirror);
        }
        (None, Some(_)) => target.indentation = base.indentation.clone(),
        _ => {}
    }
    // §17.3.1.33: merge spacing sub-fields individually so partial
    // overrides (e.g., line from table style, after from paragraph style)
    // combine correctly.
    match (target.spacing.get_mut(), base.spacing.get()) {
        (Some(ts), Some(bs)) => {
            merge_opt(&mut ts.before, &bs.before);
            merge_opt(&mut ts.after, &bs.after);
            merge_opt(&mut ts.line, &bs.line);
            merge_opt(&mut ts.before_auto_spacing, &bs.before_auto_spacing);
            merge_opt(&mut ts.after_auto_spacing, &bs.after_auto_spacing);
        }
        (None, Some(_)) => target.spacing = base.spacing.clone(),
        _ => {}
    }
    merge_fields!(
        target,
        base,
        numbering,
        borders,
        shading,
        keep_next,
        keep_lines,
        widow_control,
        page_break_before,
        suppress_auto_hyphens,
        contextual_spacing,
        bidi,
        word_wrap,
        outline_level,
        text_alignment,
        cnf_style,
        frame_properties,
        auto_space_de,
        auto_space_dn,
    );

    // §17.3.1.38: merge tab stops at the individual-stop level.
    // Child Clear entries remove matching positions from the parent.
    // Child non-Clear entries are added. Result is sorted by position.
    if !base.tabs.is_empty() || !target.tabs.is_empty() {
        let child_tabs = std::mem::take(&mut target.tabs);
        // Start from parent tabs.
        target.tabs.clone_from(&base.tabs);
        // Remove positions that the child clears.
        for clear in child_tabs
            .iter()
            .filter(|t| t.alignment == TabAlignment::Clear)
        {
            target.tabs.retain(|t| t.position != clear.position);
        }
        // Add child's non-Clear tabs (replacing any at the same position).
        for tab in child_tabs
            .iter()
            .filter(|t| t.alignment != TabAlignment::Clear)
        {
            target.tabs.retain(|t| t.position != tab.position);
            target.tabs.push(*tab);
        }
        target.tabs.sort_by_key(|t| t.position.raw());
    }
}

/// §17.7.2 / §17.7.4.3: merge a parent table style's `<w:tblPr>` into a child's.
///
/// The unit of inheritance is the **property**, not the `<w:tblPr>` element that
/// carries it: `basedOn` says the child "inherits all of the properties of the
/// base style", and the child then layers its own on top. So a child that states
/// one `tblPr` child element keeps the parent's every other property.
///
/// This used to merge `cell_margins` alone, on the stated grounds that the rest
/// was "resolved separately through the table-level cascade in `build_table`".
/// That was never true — there is no `basedOn` walk there, only a single map
/// lookup by the leaf `style_id` — so a `<w:tblLayout>` in the child was enough
/// to erase the parent's borders outright.
///
/// `style_id` is the one exclusion, for the reason
/// [`overlay_table_properties`] excludes it: a `tblStyle` reference inside a
/// style's own `tblPr` would be a second inheritance edge competing with
/// `basedOn`, which §17.7 does not define and which invites a resolution loop.
/// Nothing reads `ResolvedStyle::table`'s `style_id` — `build_table` takes the
/// reference off the `<w:tbl>` — so leaving it unmerged costs nothing.
pub fn merge_table_properties(
    target: &mut Option<TableProperties>,
    base: &Option<TableProperties>,
) {
    let Some(b) = base.as_ref() else { return };
    // A child with no `<w:tblPr>` at all inherits the parent's outright, which
    // is the same rule with nothing on the left of it — hence one path, so the
    // two cases cannot drift apart.
    let t = target.get_or_insert_with(TableProperties::default);
    merge_fields!(
        t,
        b,
        alignment,
        width,
        layout,
        indent,
        borders,
        cell_margins,
        cell_spacing,
        look,
        style_row_band_size,
        style_col_band_size,
        positioning,
        overlap,
        bidi_visual,
    );
}

/// §17.7.2: merge a parent's `<w:tcPr>` into a child's, property by property.
///
/// Only reached through [`merge_table_style_overrides`] — a `tcPr` on a *cell*
/// has no parent to inherit from, it *is* the exception over the table default.
pub fn merge_table_cell_properties(target: &mut TableCellProperties, base: &TableCellProperties) {
    merge_fields!(
        target,
        base,
        width,
        borders,
        shading,
        margins,
        vertical_align,
        vertical_merge,
        grid_span,
        text_direction,
        no_wrap,
        cnf_style,
    );
}

/// §17.7.6 + §17.7.4.3: layer a child table style's `<w:tblStylePr>` set over
/// its parent's.
///
/// Conditional formatting is part of the style *definition*, so `basedOn`
/// carries it like any other property — a user style derived from a banded
/// built-in inherits that built-in's bands. The layers are keyed by `w:type`,
/// and where both levels define the same type they merge **property by
/// property**: the child's values win and the parent's fill the gaps. That is
/// the granularity §17.7.2 uses at every other level of the cascade, and it is
/// what keeps a child that restates one region's shading from also discarding
/// that region's inherited run formatting.
///
/// `trPr` is the exception, merged whole rather than per field:
/// `TableRowProperties` carries `grid_before`/`grid_after` as plain `u32` with
/// no "unset" state, so a per-field merge could not tell an inherited value from
/// a declared zero. Nothing reads a style-level `trPr` today, so the choice is
/// currently unobservable; giving those two fields a carrier is what would let
/// this merge like the rest.
///
/// The parent's layers come first in the result, in the parent's own order,
/// with child-only types appended — [`resolve_cell_conditional`] looks each
/// type up by key, so the order only has to be deterministic, not meaningful.
///
/// [`resolve_cell_conditional`]: super::conditional::resolve_cell_conditional
pub fn merge_table_style_overrides(
    child: Vec<TableStyleOverride>,
    parent: &[TableStyleOverride],
) -> Vec<TableStyleOverride> {
    if parent.is_empty() {
        return child;
    }
    let mut out = parent.to_vec();
    for mut layer in child {
        let Some(base) = out
            .iter_mut()
            .find(|o| o.override_type == layer.override_type)
        else {
            out.push(layer);
            continue;
        };
        merge_bag(
            &mut layer.paragraph_properties,
            &base.paragraph_properties,
            merge_paragraph_properties,
        );
        merge_bag(
            &mut layer.run_properties,
            &base.run_properties,
            merge_run_properties,
        );
        merge_table_properties(&mut layer.table_properties, &base.table_properties);
        merge_bag(
            &mut layer.table_cell_properties,
            &base.table_cell_properties,
            merge_table_cell_properties,
        );
        merge_opt(&mut layer.table_row_properties, &base.table_row_properties);
        *base = layer;
    }
    out
}

/// Merge one `Option`-carried property bag from `base` into `target`: `merge`
/// runs when both levels have one, and an absent target simply takes the base's.
fn merge_bag<T: Clone>(target: &mut Option<T>, base: &Option<T>, merge: fn(&mut T, &T)) {
    match (target.as_mut(), base.as_ref()) {
        (Some(t), Some(b)) => merge(t, b),
        (None, Some(_)) => target.clone_from(base),
        _ => {}
    }
}

/// §17.7.6: overlay a `wholeTable` conditional layer onto a table style's own
/// table properties. Every field the overlay specifies wins; the rest fall
/// through to `base`.
///
/// This is the opposite direction from [`merge_table_properties`], which is
/// *inheritance* (the target keeps what it has and the base fills gaps).
/// `wholeTable` is conditional formatting layered on top of the style, so it
/// overrides rather than backfills.
///
/// `style_id` is deliberately not overlaid: a `tblStylePr` carrying a
/// `tblStyle` reference would point a style at another style from inside its
/// own conditional formatting, which §17.7.6 does not define and which would
/// invite a resolution loop.
pub fn overlay_table_properties(
    base: Option<TableProperties>,
    overlay: &TableProperties,
) -> TableProperties {
    let mut out = base.unwrap_or_default();
    // Listed field by field rather than by `..overlay.clone()` so that adding a
    // field to `TableProperties` without deciding how it layers is a visible
    // omission here, not a silent one. `table_properties_overlay_covers_every_field`
    // pins it.
    take(&mut out.alignment, overlay.alignment.clone());
    take(&mut out.width, overlay.width.clone());
    take(&mut out.layout, overlay.layout.clone());
    take(&mut out.indent, overlay.indent.clone());
    take(&mut out.borders, overlay.borders.clone());
    take(&mut out.cell_margins, overlay.cell_margins.clone());
    take(&mut out.cell_spacing, overlay.cell_spacing.clone());
    take(&mut out.look, overlay.look.clone());
    take(
        &mut out.style_row_band_size,
        overlay.style_row_band_size.clone(),
    );
    take(
        &mut out.style_col_band_size,
        overlay.style_col_band_size.clone(),
    );
    take(&mut out.positioning, overlay.positioning.clone());
    take(&mut out.overlap, overlay.overlap.clone());
    take(&mut out.bidi_visual, overlay.bidi_visual.clone());
    out
}

/// §17.7.6: if `overlay` specifies a value, it replaces `target`.
///
/// The overlay twin of [`merge_opt`], and generic over the same [`Unset`] trait
/// so it reads both carriers. For a [`Dup`] the replacement is wholesale: an
/// overlay that set the property brings *all* of its occurrences and drops the
/// target's, which is what "replaces" has always meant here.
fn take<T: Unset>(target: &mut T, overlay: T) {
    if !overlay.is_unset() {
        *target = overlay;
    }
}

/// §17.7.2 inheritance: "this level did not set the property".
///
/// Implemented for both carriers a property can have — `Option` for a value the
/// seam collapsed, [`Dup`] for one that still holds every occurrence the
/// document wrote. Absent means the same thing either way, so the cascade does
/// not have to know which it is looking at.
trait Unset {
    fn is_unset(&self) -> bool;
}

impl<T> Unset for Option<T> {
    fn is_unset(&self) -> bool {
        self.is_none()
    }
}

impl<T> Unset for Dup<T> {
    fn is_unset(&self) -> bool {
        self.is_absent()
    }
}

/// If `target` did not set this property, clone `base` into it.
///
/// **What this decides for `Dup`:** a child that set the property wins with
/// *all* of its occurrences — the parent's are not appended, and the child's
/// losers are not replaced by the parent's winner. That is the faithful
/// generalisation of the `Option` rule it replaced ("the child wins if it said
/// anything"), and it keeps the two questions separate: `Dup::get` resolves a
/// duplicate *within* one bag, this resolves *across* cascade levels. Merging
/// the two lists would conflate them.
fn merge_opt<T: Clone + Unset>(target: &mut T, base: &T) {
    if target.is_unset() {
        *target = base.clone();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::{Dimension, HalfPoints, Twips};
    use crate::model::*;

    // ── RunProperties merging ────────────────────────────────────────────

    #[test]
    fn merge_run_empty_target_takes_all_from_base() {
        let mut target = RunProperties::default();
        let base = RunProperties {
            bold: Some(true),
            italic: Some(true),
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(24))),
            color: Dup::from(Some(Color::Rgb(0xFF0000))),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);

        assert_eq!(target.bold, Some(true));
        assert_eq!(target.italic, Some(true));
        assert_eq!(
            target.font_size,
            Dup::from(Some(Dimension::<HalfPoints>::new(24)))
        );
        assert_eq!(target.color, Dup::from(Some(Color::Rgb(0xFF0000))));
    }

    #[test]
    fn merge_run_target_values_not_overwritten() {
        let mut target = RunProperties {
            bold: Some(false),
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(20))),
            ..Default::default()
        };
        let base = RunProperties {
            bold: Some(true),
            italic: Some(true),
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(24))),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);

        assert_eq!(target.bold, Some(false), "target's bold should win");
        assert_eq!(
            target.font_size,
            Dup::from(Some(Dimension::<HalfPoints>::new(20))),
            "target's size should win"
        );
        assert_eq!(target.italic, Some(true), "italic should come from base");
    }

    #[test]
    fn merge_run_both_empty_stays_empty() {
        let mut target = RunProperties::default();
        let base = RunProperties::default();
        merge_run_properties(&mut target, &base);

        assert_eq!(target.bold, None);
        assert_eq!(target.italic, None);
        assert_eq!(target.font_size, Dup::from(None));
    }

    #[test]
    fn merge_run_explicit_underline_none_overrides_parent_single() {
        // §17.7.2 + §17.3.2.40: <w:u w:val="none"/> on the child explicitly
        // turns off an underline that the parent style would otherwise
        // inherit. The merge must preserve `Some(UnderlineStyle::None)`
        // rather than overwriting it with the parent's `Single`.
        let mut target = RunProperties {
            underline: Dup::from(Some(UnderlineStyle::None)),
            ..Default::default()
        };
        let base = RunProperties {
            underline: Dup::from(Some(UnderlineStyle::Single)),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);
        assert_eq!(
            target.underline,
            Dup::from(Some(UnderlineStyle::None)),
            "explicit <w:u w:val=\"none\"/> must override the parent's Single"
        );
    }

    #[test]
    fn merge_run_absent_underline_inherits_from_parent() {
        // Mirror invariant: when the child is silent, parent's underline wins.
        let mut target = RunProperties::default();
        let base = RunProperties {
            underline: Dup::from(Some(UnderlineStyle::Single)),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);
        assert_eq!(target.underline, Dup::from(Some(UnderlineStyle::Single)));
    }

    #[test]
    fn merge_run_fonts_merged_field_by_field() {
        let mut target = RunProperties {
            fonts: FontSet {
                ascii: FontSlot::from_name("Arial"),
                ..Default::default()
            },
            ..Default::default()
        };
        let base = RunProperties {
            fonts: FontSet {
                ascii: FontSlot::from_name("Times"),
                east_asian: FontSlot::from_name("SimSun"),
                ..Default::default()
            },
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);

        assert_eq!(
            target.fonts.ascii.explicit.as_deref(),
            Some("Arial"),
            "target's ascii should win"
        );
        assert_eq!(
            target.fonts.east_asian.explicit.as_deref(),
            Some("SimSun"),
            "east_asian should come from base"
        );
    }

    #[test]
    fn merge_run_all_fields_covered() {
        // Ensure merge touches every field by setting all in base, none in target.
        let base = RunProperties {
            fonts: FontSet {
                ascii: FontSlot::from_name("F"),
                high_ansi: FontSlot::from_name("F"),
                east_asian: FontSlot::from_name("F"),
                complex_script: FontSlot::from_name("F"),
            },
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(24))),
            bold: Some(true),
            italic: Some(true),
            underline: Dup::from(Some(UnderlineStyle::Single)),
            strike: Some(StrikeStyle::Single),
            color: Dup::from(Some(Color::Rgb(0))),
            highlight: Dup::from(Some(HighlightColor::Yellow)),
            shading: Dup::from(Some(Shading {
                fill: Color::Rgb(0),
                pattern: ShadingPattern::Clear,
                color: Color::Rgb(0),
            })),
            vertical_align: Dup::from(Some(VerticalAlign::Superscript)),
            spacing: Dup::from(Some(Dimension::<Twips>::new(10))),
            kerning: Dup::from(Some(Dimension::<HalfPoints>::new(2))),
            all_caps: Some(true),
            small_caps: Some(true),
            vanish: Some(true),
            no_proof: Some(true),
            web_hidden: Some(true),
            rtl: Some(true),
            emboss: Some(true),
            imprint: Some(true),
            outline: Some(true),
            shadow: Some(true),
            position: Dup::from(Some(Dimension::<HalfPoints>::new(5))),
            lang: Dup::from(Some(Lang {
                val: Some("en".into()),
                east_asia: None,
                bidi: None,
            })),
            border: Dup::from(Some(Border {
                style: BorderStyle::Single,
                width: Dimension::new(0),
                space: Dimension::new(0),
                color: Color::BLACK,
            })),
            text_scale: Dup::from(Some(TextScale::new(80))),
        };
        let mut target = RunProperties::default();
        merge_run_properties(&mut target, &base);

        assert!(target.bold.is_some());
        assert!(target.italic.is_some());
        assert!(target.underline.cloned().is_some());
        assert!(target.strike.is_some());
        assert!(target.color.cloned().is_some());
        assert!(target.highlight.cloned().is_some());
        assert!(target.shading.cloned().is_some());
        assert!(target.vertical_align.cloned().is_some());
        assert!(target.spacing.cloned().is_some());
        assert!(target.kerning.cloned().is_some());
        assert!(target.all_caps.is_some());
        assert!(target.small_caps.is_some());
        assert!(target.vanish.is_some());
        assert!(target.no_proof.is_some());
        assert!(target.web_hidden.is_some());
        assert!(target.rtl.is_some());
        assert!(target.emboss.is_some());
        assert!(target.imprint.is_some());
        assert!(target.outline.is_some());
        assert!(target.shadow.is_some());
        assert!(target.position.cloned().is_some());
        assert!(target.lang.cloned().is_some());
        assert!(target.border.cloned().is_some());
        assert!(target.fonts.ascii.explicit.is_some());
        assert!(target.fonts.high_ansi.explicit.is_some());
        assert!(target.fonts.east_asian.explicit.is_some());
        assert!(target.fonts.complex_script.explicit.is_some());
        assert!(target.font_size.cloned().is_some());
        assert_eq!(target.text_scale, Dup::from(Some(TextScale::new(80))));
    }

    #[test]
    fn merge_run_text_scale_child_wins() {
        // §17.7.2: when both parent and child specify text_scale, the child's
        // value wins. Mirrors the standard merge semantics for run properties.
        let mut target = RunProperties {
            text_scale: Dup::from(Some(TextScale::new(120))),
            ..Default::default()
        };
        let base = RunProperties {
            text_scale: Dup::from(Some(TextScale::new(80))),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);
        assert_eq!(target.text_scale, Dup::from(Some(TextScale::new(120))));
    }

    #[test]
    fn merge_run_text_scale_inherited_from_parent() {
        // No child override → parent's text_scale propagates.
        let mut target = RunProperties::default();
        let base = RunProperties {
            text_scale: Dup::from(Some(TextScale::new(80))),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);
        assert_eq!(target.text_scale, Dup::from(Some(TextScale::new(80))));
    }

    // ── ParagraphProperties merging ──────────────────────────────────────

    #[test]
    fn merge_para_empty_target_takes_from_base() {
        let mut target = ParagraphProperties::default();
        let base = ParagraphProperties {
            alignment: Dup::from(Some(Alignment::Center)),
            keep_next: Some(true),
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert_eq!(target.alignment, Dup::from(Some(Alignment::Center)));
        assert_eq!(target.keep_next, Some(true));
    }

    #[test]
    fn merge_para_target_values_not_overwritten() {
        let mut target = ParagraphProperties {
            alignment: Dup::from(Some(Alignment::End)),
            ..Default::default()
        };
        let base = ParagraphProperties {
            alignment: Dup::from(Some(Alignment::Center)),
            keep_next: Some(true),
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert_eq!(
            target.alignment,
            Dup::from(Some(Alignment::End)),
            "target should win"
        );
        assert_eq!(target.keep_next, Some(true), "keep_next from base");
    }

    #[test]
    fn merge_tabs_child_adds_to_parent() {
        // §17.3.1.38: child non-Clear tabs are added to inherited parent tabs.
        let mut target = ParagraphProperties {
            tabs: vec![TabStop {
                position: Dimension::<Twips>::new(720),
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }],
            ..Default::default()
        };
        let base = ParagraphProperties {
            tabs: vec![TabStop {
                position: Dimension::<Twips>::new(1440),
                alignment: TabAlignment::Right,
                leader: TabLeader::Dot,
            }],
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert_eq!(target.tabs.len(), 2, "both tabs should be present");
        assert_eq!(
            target.tabs[0].position,
            Dimension::<Twips>::new(720),
            "sorted: left@720 first"
        );
        assert_eq!(
            target.tabs[1].position,
            Dimension::<Twips>::new(1440),
            "sorted: right@1440 second"
        );
    }

    #[test]
    fn merge_tabs_inherited_when_target_empty() {
        let mut target = ParagraphProperties::default();
        let base = ParagraphProperties {
            tabs: vec![TabStop {
                position: Dimension::<Twips>::new(1440),
                alignment: TabAlignment::Right,
                leader: TabLeader::Dot,
            }],
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert_eq!(target.tabs.len(), 1);
        assert_eq!(target.tabs[0].position, Dimension::<Twips>::new(1440));
    }

    #[test]
    fn merge_tabs_clear_removes_parent_stop() {
        // §17.3.1.38: val="clear" removes an inherited tab at that position.
        let mut target = ParagraphProperties {
            tabs: vec![
                TabStop {
                    position: Dimension::<Twips>::new(4536),
                    alignment: TabAlignment::Clear,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: Dimension::<Twips>::new(1701),
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
            ],
            ..Default::default()
        };
        let base = ParagraphProperties {
            tabs: vec![
                TabStop {
                    position: Dimension::<Twips>::new(4536),
                    alignment: TabAlignment::Center,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: Dimension::<Twips>::new(9072),
                    alignment: TabAlignment::Right,
                    leader: TabLeader::None,
                },
            ],
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        // center@4536 removed by clear, right@9072 inherited, left@1701 added.
        assert_eq!(target.tabs.len(), 2);
        assert_eq!(target.tabs[0].position, Dimension::<Twips>::new(1701));
        assert_eq!(target.tabs[0].alignment, TabAlignment::Left);
        assert_eq!(target.tabs[1].position, Dimension::<Twips>::new(9072));
        assert_eq!(target.tabs[1].alignment, TabAlignment::Right);
    }

    #[test]
    fn merge_tabs_clear_only_no_additions() {
        // Child only clears a parent tab, adds nothing.
        let mut target = ParagraphProperties {
            tabs: vec![TabStop {
                position: Dimension::<Twips>::new(4536),
                alignment: TabAlignment::Clear,
                leader: TabLeader::None,
            }],
            ..Default::default()
        };
        let base = ParagraphProperties {
            tabs: vec![TabStop {
                position: Dimension::<Twips>::new(4536),
                alignment: TabAlignment::Center,
                leader: TabLeader::None,
            }],
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert!(
            target.tabs.is_empty(),
            "cleared tab should be gone, nothing added"
        );
    }

    #[test]
    fn merge_para_all_fields_covered() {
        let base = ParagraphProperties {
            alignment: Dup::from(Some(Alignment::Start)),
            indentation: Dup::from(Some(Indentation::default())),
            spacing: Dup::from(Some(ParagraphSpacing::default())),
            numbering: Dup::from(Some(NumberingReference {
                num_id: 1,
                level: 0,
            })),
            tabs: vec![TabStop {
                position: Dimension::new(720),
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }],
            borders: Dup::from(Some(ParagraphBorders {
                top: None,
                bottom: None,
                left: None,
                right: None,
                between: None,
            })),
            shading: Dup::from(Some(Shading {
                fill: Color::Rgb(0),
                pattern: ShadingPattern::Clear,
                color: Color::Rgb(0),
            })),
            keep_next: Some(true),
            keep_lines: Some(true),
            widow_control: Some(true),
            page_break_before: Some(true),
            suppress_auto_hyphens: Some(true),
            contextual_spacing: Some(true),
            bidi: Some(true),
            word_wrap: Some(true),
            outline_level: Dup::from(Some(OutlineLevel::new(1))),
            text_alignment: Dup::from(Some(TextAlignment::Center)),
            cnf_style: Dup::from(Some(CnfStyle::FIRST_ROW)),
            frame_properties: Dup::from(None),
            auto_space_de: Some(true),
            auto_space_dn: Some(true),
        };
        let mut target = ParagraphProperties::default();
        merge_paragraph_properties(&mut target, &base);

        assert!(target.alignment.get().is_some());
        assert!(target.indentation.get().is_some());
        assert!(target.spacing.get().is_some());
        assert!(target.numbering.get().is_some());
        assert!(!target.tabs.is_empty());
        assert!(!target.borders.is_absent());
        assert!(target.shading.get().is_some());
        assert!(target.keep_next.is_some());
        assert!(target.keep_lines.is_some());
        assert!(target.widow_control.is_some());
        assert!(target.page_break_before.is_some());
        assert!(target.suppress_auto_hyphens.is_some());
        assert!(target.contextual_spacing.is_some());
        assert!(target.bidi.is_some());
        assert!(target.word_wrap.is_some());
        assert!(target.outline_level.get().is_some());
        assert!(target.text_alignment.get().is_some());
        assert!(target.cnf_style.get().is_some());
        assert!(target.auto_space_de.is_some());
        assert!(target.auto_space_dn.is_some());
    }

    // ── table-property cascade (§17.7.2 inheritance, §17.7.6 overlay) ────

    /// A `TableProperties` with **every** field set, so a test can assert that
    /// a cascade step carried all of them. Shared by the inheritance and
    /// overlay coverage tests below — one place to add a new field to, which is
    /// what makes forgetting one a failure in both directions.
    fn every_table_property() -> TableProperties {
        use crate::model::geometry::EdgeInsets;
        TableProperties {
            style_id: Some(StyleId::new("Ignored")),
            alignment: Dup::from(Some(Alignment::Center)),
            width: Dup::from(Some(TableMeasure::Auto)),
            layout: Dup::from(Some(TableLayout::Fixed)),
            indent: Dup::from(Some(TableMeasure::Twips(Dimension::new(120)))),
            borders: Dup::from(Some(TableBorders {
                top: None,
                bottom: None,
                left: None,
                right: None,
                inside_h: None,
                inside_v: None,
            })),
            cell_margins: Dup::from(Some(EdgeInsets {
                top: Dimension::new(1),
                right: Dimension::new(2),
                bottom: Dimension::new(3),
                left: Dimension::new(4),
            })),
            cell_spacing: Dup::from(Some(TableMeasure::Twips(Dimension::new(30)))),
            look: Dup::from(Some(TableLook {
                first_row: Some(true),
                last_row: None,
                first_column: None,
                last_column: None,
                no_h_band: None,
                no_v_band: None,
            })),
            style_row_band_size: Dup::from(Some(2)),
            style_col_band_size: Dup::from(Some(3)),
            positioning: Dup::from(Some(TablePositioning {
                left_from_text: None,
                right_from_text: None,
                top_from_text: None,
                bottom_from_text: None,
                vert_anchor: None,
                horz_anchor: None,
                x_align: None,
                y_align: None,
                x: Some(Dimension::new(5)),
                y: None,
            })),
            overlap: Dup::from(Some(TableOverlap::Never)),
            bidi_visual: Dup::from(Some(true)),
        }
    }

    /// Every field must be listed in [`merge_table_properties`]. A field added
    /// to `TableProperties` and not to the merge would be unreachable from a
    /// parent table style — silently, since the child would simply render
    /// without it. This fails instead.
    #[test]
    fn table_properties_merge_covers_every_field() {
        let base = every_table_property();
        // `Some(default)` is the case the old code got wrong: a child that has
        // a `<w:tblPr>` at all, but states none of these properties in it.
        let mut target = Some(TableProperties::default());
        merge_table_properties(&mut target, &Some(base.clone()));
        let out = target.expect("target stays present");

        assert_eq!(out.alignment, base.alignment);
        assert_eq!(out.width, base.width);
        assert_eq!(out.layout, base.layout);
        assert_eq!(out.indent, base.indent);
        assert!(out.borders.cloned().is_some(), "borders inherited"); // no PartialEq on TableBorders
        assert_eq!(out.cell_margins, base.cell_margins);
        assert_eq!(out.cell_spacing, base.cell_spacing);
        assert!(out.look.cloned().is_some(), "look inherited"); // no PartialEq on TableLook
        assert_eq!(out.style_row_band_size, base.style_row_band_size);
        assert_eq!(out.style_col_band_size, base.style_col_band_size);
        assert_eq!(out.positioning, base.positioning);
        assert_eq!(out.overlap, base.overlap);
        assert_eq!(out.bidi_visual, base.bidi_visual);
        // The one deliberate exclusion: `basedOn` is the inheritance edge, and
        // a `tblStyle` inside a style's own `tblPr` must not become a second.
        assert_eq!(
            out.style_id, None,
            "style_id must not be inherited — see merge_table_properties"
        );
    }

    /// A child that declares no `<w:tblPr>` at all takes the parent's, and by
    /// the same rule — so the two arms cannot drift apart.
    #[test]
    fn a_table_style_with_no_tbl_pr_inherits_the_parents() {
        let base = every_table_property();
        let mut target = None;
        merge_table_properties(&mut target, &Some(base.clone()));
        let out = target.expect("an absent tblPr becomes the parent's");
        assert_eq!(out.width, base.width);
        assert_eq!(out.alignment, base.alignment);
        assert_eq!(out.style_id, None, "…still without the style reference");
    }

    /// §17.7.2: what the child states wins, per property — the parent fills
    /// only the gaps.
    #[test]
    fn table_properties_merge_lets_the_child_win_per_property() {
        let base = every_table_property();
        let mut target = Some(TableProperties {
            alignment: Dup::from(Some(Alignment::End)),
            ..Default::default()
        });
        merge_table_properties(&mut target, &Some(base.clone()));
        let out = target.expect("target stays present");
        assert_eq!(
            out.alignment,
            Dup::from(Some(Alignment::End)),
            "the child's own alignment survives"
        );
        assert_eq!(
            out.width, base.width,
            "…and the parent still fills what the child left unset"
        );
    }

    /// Every field must be listed in `overlay_table_properties`. If a field is
    /// added to `TableProperties` and not to the overlay, a `wholeTable` style
    /// setting it would be silently dropped — this fails instead.
    #[test]
    fn table_properties_overlay_covers_every_field() {
        let overlay = every_table_property();
        let out = overlay_table_properties(None, &overlay);

        assert_eq!(out.alignment, overlay.alignment);
        assert_eq!(out.width, overlay.width);
        assert_eq!(out.layout, overlay.layout);
        assert_eq!(out.indent, overlay.indent);
        assert!(out.borders.cloned().is_some(), "borders overlaid"); // no PartialEq on TableBorders
        assert_eq!(out.cell_margins, overlay.cell_margins);
        assert_eq!(out.cell_spacing, overlay.cell_spacing);
        assert!(out.look.cloned().is_some(), "look overlaid"); // no PartialEq on TableLook
        assert_eq!(out.style_row_band_size, overlay.style_row_band_size);
        assert_eq!(out.style_col_band_size, overlay.style_col_band_size);
        assert_eq!(out.positioning, overlay.positioning);
        assert_eq!(out.overlap, overlay.overlap);
        assert_eq!(out.bidi_visual, overlay.bidi_visual);
        // The one deliberate exclusion: a conditional layer must not re-point
        // the style at another style.
        assert_eq!(
            out.style_id, None,
            "style_id must not be overlaid — see overlay_table_properties"
        );
    }

    // ── §17.7.6 conditional layers through `basedOn` ─────────────────────

    /// A `tblStylePr` layer carrying a shading fill and/or a bold toggle.
    fn layer(
        override_type: TableStyleOverrideType,
        fill: Option<u32>,
        bold: Option<bool>,
    ) -> TableStyleOverride {
        TableStyleOverride {
            override_type,
            paragraph_properties: None,
            run_properties: bold.map(|b| RunProperties {
                bold: Some(b),
                ..Default::default()
            }),
            table_properties: None,
            table_row_properties: None,
            table_cell_properties: fill.map(|f| TableCellProperties {
                shading: Dup::from(Some(Shading {
                    fill: Color::Rgb(f),
                    pattern: ShadingPattern::Clear,
                    color: Color::Auto,
                })),
                ..Default::default()
            }),
        }
    }

    fn fill_of(layer: &TableStyleOverride) -> Option<Color> {
        layer
            .table_cell_properties
            .as_ref()
            .and_then(|p| p.shading.get())
            .map(|s| s.fill)
    }

    /// A child that defines no layer of that type inherits the parent's whole.
    #[test]
    fn conditional_layers_descend_when_the_child_defines_none() {
        let parent = vec![layer(
            TableStyleOverrideType::FirstRow,
            Some(0xFF0000),
            None,
        )];
        let out = merge_table_style_overrides(vec![], &parent);
        assert_eq!(out.len(), 1);
        assert_eq!(fill_of(&out[0]), Some(Color::Rgb(0xFF0000)));
    }

    /// Two layers of the same type merge per property rather than replacing:
    /// the child's shading wins, the parent's run formatting survives.
    #[test]
    fn conditional_layers_of_the_same_type_merge_property_by_property() {
        let parent = vec![layer(
            TableStyleOverrideType::FirstRow,
            Some(0xFF0000),
            Some(true),
        )];
        let child = vec![layer(
            TableStyleOverrideType::FirstRow,
            Some(0x00FF00),
            None,
        )];
        let out = merge_table_style_overrides(child, &parent);

        assert_eq!(out.len(), 1, "one layer per type, not two");
        assert_eq!(
            fill_of(&out[0]),
            Some(Color::Rgb(0x00FF00)),
            "the child's own shading wins"
        );
        assert_eq!(
            out[0].run_properties.as_ref().and_then(|r| r.bold),
            Some(true),
            "the parent's run formatting, which the child never restates, survives"
        );
    }

    /// …and the same holds *inside* a layer's `<w:tcPr>`: a child that states
    /// one cell property keeps the parent's others for that region, rather than
    /// its `tcPr` replacing the parent's whole.
    #[test]
    fn a_layers_cell_properties_merge_property_by_property_too() {
        let parent = vec![layer(
            TableStyleOverrideType::FirstRow,
            Some(0xFF0000),
            None,
        )];
        let mut child_layer = layer(TableStyleOverrideType::FirstRow, None, None);
        child_layer.table_cell_properties = Some(TableCellProperties {
            vertical_align: Dup::from(Some(CellVerticalAlign::Center)),
            ..Default::default()
        });
        let out = merge_table_style_overrides(vec![child_layer], &parent);

        assert_eq!(out.len(), 1);
        assert_eq!(
            out[0]
                .table_cell_properties
                .as_ref()
                .and_then(|p| p.vertical_align.cloned()),
            Some(CellVerticalAlign::Center),
            "the child's own vAlign applies"
        );
        assert_eq!(
            fill_of(&out[0]),
            Some(Color::Rgb(0xFF0000)),
            "…and the parent's shading for the same region is not discarded with it"
        );
    }

    /// A layer's `<w:trPr>` descends whole, which is the documented exception:
    /// `TableRowProperties` has no "unset" state for `grid_before`/`grid_after`,
    /// so it cannot merge per field. Pinned so the exception stays a decision.
    #[test]
    fn a_layers_row_properties_descend_whole() {
        let mut parent_layer = layer(TableStyleOverrideType::FirstRow, None, None);
        parent_layer.table_row_properties = Some(TableRowProperties {
            is_header: Some(true),
            ..Default::default()
        });
        let child = vec![layer(
            TableStyleOverrideType::FirstRow,
            Some(0x00FF00),
            None,
        )];
        let out = merge_table_style_overrides(child, &[parent_layer]);

        assert_eq!(
            out[0]
                .table_row_properties
                .as_ref()
                .and_then(|r| r.is_header),
            Some(true),
            "a child layer that states no trPr takes the parent's"
        );
    }

    /// Types only one level defines are kept from both — the merge is keyed by
    /// `w:type`, not a replacement of the set.
    #[test]
    fn conditional_layers_are_keyed_by_type() {
        let parent = vec![
            layer(TableStyleOverrideType::FirstRow, Some(0xFF0000), None),
            layer(TableStyleOverrideType::Band1Horz, Some(0x0000FF), None),
        ];
        let child = vec![layer(TableStyleOverrideType::LastRow, Some(0x00FF00), None)];
        let out = merge_table_style_overrides(child, &parent);

        let by_type = |t| out.iter().find(|o| o.override_type == t).and_then(fill_of);
        assert_eq!(
            by_type(TableStyleOverrideType::FirstRow),
            Some(Color::Rgb(0xFF0000))
        );
        assert_eq!(
            by_type(TableStyleOverrideType::Band1Horz),
            Some(Color::Rgb(0x0000FF))
        );
        assert_eq!(
            by_type(TableStyleOverrideType::LastRow),
            Some(Color::Rgb(0x00FF00))
        );
        assert_eq!(out.len(), 3);
    }

    /// A styleless parent is the common case — it must be a pass-through, and
    /// must not disturb the child's own order.
    #[test]
    fn conditional_layers_pass_through_an_empty_parent() {
        let child = vec![
            layer(TableStyleOverrideType::LastRow, Some(0x00FF00), None),
            layer(TableStyleOverrideType::FirstRow, Some(0xFF0000), None),
        ];
        let out = merge_table_style_overrides(child.clone(), &[]);
        assert_eq!(out.len(), child.len());
        assert_eq!(out[0].override_type, TableStyleOverrideType::LastRow);
        assert_eq!(out[1].override_type, TableStyleOverrideType::FirstRow);
    }

    /// Every field must be listed in [`merge_table_cell_properties`], or a
    /// parent layer's value for it would be unreachable from a child layer of
    /// the same type.
    #[test]
    fn table_cell_properties_merge_covers_every_field() {
        use crate::model::geometry::PartialEdgeInsets;
        let base = TableCellProperties {
            width: Dup::from(Some(TableMeasure::Auto)),
            borders: Dup::from(Some(TableCellBorders {
                top: None,
                bottom: None,
                left: None,
                right: None,
                tl2br: None,
                tr2bl: None,
                inside_h: None,
                inside_v: None,
            })),
            shading: Dup::from(Some(Shading {
                fill: Color::Rgb(0xFF0000),
                pattern: ShadingPattern::Clear,
                color: Color::Auto,
            })),
            margins: Dup::from(Some(PartialEdgeInsets {
                top: Some(Dimension::new(1)),
                right: None,
                bottom: None,
                left: None,
            })),
            vertical_align: Dup::from(Some(CellVerticalAlign::Center)),
            vertical_merge: Dup::from(Some(VerticalMerge::Restart)),
            grid_span: Dup::from(Some(2)),
            text_direction: Dup::from(Some(TextDirection::TopToBottomRightToLeft)),
            no_wrap: Some(true),
            cnf_style: Dup::from(Some(CnfStyle::FIRST_ROW)),
        };
        let mut target = TableCellProperties::default();
        merge_table_cell_properties(&mut target, &base);

        assert_eq!(target.width, base.width);
        assert!(target.borders.cloned().is_some(), "borders"); // no PartialEq
        assert_eq!(target.shading, base.shading);
        assert_eq!(target.margins, base.margins);
        assert_eq!(target.vertical_align, base.vertical_align);
        assert_eq!(target.vertical_merge, base.vertical_merge);
        assert_eq!(target.grid_span, base.grid_span);
        assert_eq!(target.text_direction, base.text_direction);
        assert_eq!(target.no_wrap, base.no_wrap);
        assert_eq!(target.cnf_style, base.cnf_style);
    }

    /// The overlay direction is the opposite of `merge_table_properties`:
    /// what the overlay specifies wins, and what it omits falls through.
    #[test]
    fn table_properties_overlay_wins_but_only_where_specified() {
        use crate::model::geometry::EdgeInsets;
        let base = TableProperties {
            alignment: Dup::from(Some(Alignment::Start)),
            cell_margins: Dup::from(Some(EdgeInsets {
                top: Dimension::new(9),
                right: Dimension::new(9),
                bottom: Dimension::new(9),
                left: Dimension::new(9),
            })),
            ..Default::default()
        };
        let overlay = TableProperties {
            alignment: Dup::from(Some(Alignment::Center)),
            ..Default::default()
        };
        let out = overlay_table_properties(Some(base.clone()), &overlay);
        assert_eq!(
            out.alignment,
            Dup::from(Some(Alignment::Center)),
            "overlay wins"
        );
        assert_eq!(
            out.cell_margins, base.cell_margins,
            "what the overlay omits falls through to the base"
        );
    }
}
