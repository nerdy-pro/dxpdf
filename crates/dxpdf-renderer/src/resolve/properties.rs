//! Option-field merging for RunProperties and ParagraphProperties.
//!
//! "Merge" means: for each field, if `self` is `None`, take the value from `base`.
//! This implements the OOXML style inheritance cascade.

use dxpdf_docx_model::model::{ParagraphProperties, RunProperties, TabAlignment};

/// Merge `base` into `target`: any `None` field in `target` gets filled from `base`.
pub fn merge_run_properties(target: &mut RunProperties, base: &RunProperties) {
    // FontSet: merge each sub-field independently
    merge_opt(&mut target.fonts.ascii, &base.fonts.ascii);
    merge_opt(&mut target.fonts.high_ansi, &base.fonts.high_ansi);
    merge_opt(&mut target.fonts.east_asian, &base.fonts.east_asian);
    merge_opt(&mut target.fonts.complex_script, &base.fonts.complex_script);

    merge_opt(&mut target.font_size, &base.font_size);
    merge_opt(&mut target.bold, &base.bold);
    merge_opt(&mut target.italic, &base.italic);
    merge_opt(&mut target.underline, &base.underline);
    merge_opt(&mut target.strike, &base.strike);
    merge_opt(&mut target.color, &base.color);
    merge_opt(&mut target.highlight, &base.highlight);
    merge_opt(&mut target.shading, &base.shading);
    merge_opt(&mut target.vertical_align, &base.vertical_align);
    merge_opt(&mut target.spacing, &base.spacing);
    merge_opt(&mut target.kerning, &base.kerning);
    merge_opt(&mut target.all_caps, &base.all_caps);
    merge_opt(&mut target.small_caps, &base.small_caps);
    merge_opt(&mut target.vanish, &base.vanish);
    merge_opt(&mut target.no_proof, &base.no_proof);
    merge_opt(&mut target.web_hidden, &base.web_hidden);
    merge_opt(&mut target.rtl, &base.rtl);
    merge_opt(&mut target.emboss, &base.emboss);
    merge_opt(&mut target.imprint, &base.imprint);
    merge_opt(&mut target.outline, &base.outline);
    merge_opt(&mut target.shadow, &base.shadow);
    merge_opt(&mut target.position, &base.position);
    merge_opt(&mut target.lang, &base.lang);
    merge_opt(&mut target.border, &base.border);
}

/// Merge `base` into `target`: any `None` field in `target` gets filled from `base`.
pub fn merge_paragraph_properties(target: &mut ParagraphProperties, base: &ParagraphProperties) {
    merge_opt(&mut target.alignment, &base.alignment);
    // §17.3.1.12: merge indentation sub-fields individually so partial
    // overrides (e.g., left from child style, firstLine from parent) combine.
    match (&mut target.indentation, &base.indentation) {
        (Some(ref mut ti), Some(bi)) => {
            merge_opt(&mut ti.start, &bi.start);
            merge_opt(&mut ti.end, &bi.end);
            merge_opt(&mut ti.first_line, &bi.first_line);
            merge_opt(&mut ti.mirror, &bi.mirror);
        }
        (None, Some(_)) => target.indentation = base.indentation,
        _ => {}
    }
    // §17.3.1.33: merge spacing sub-fields individually so partial
    // overrides (e.g., line from table style, after from paragraph style)
    // combine correctly.
    match (&mut target.spacing, &base.spacing) {
        (Some(ref mut ts), Some(bs)) => {
            merge_opt(&mut ts.before, &bs.before);
            merge_opt(&mut ts.after, &bs.after);
            merge_opt(&mut ts.line, &bs.line);
            merge_opt(&mut ts.before_auto_spacing, &bs.before_auto_spacing);
            merge_opt(&mut ts.after_auto_spacing, &bs.after_auto_spacing);
        }
        (None, Some(_)) => target.spacing = base.spacing,
        _ => {}
    }
    merge_opt(&mut target.numbering, &base.numbering);
    merge_opt(&mut target.borders, &base.borders);
    merge_opt(&mut target.shading, &base.shading);
    merge_opt(&mut target.keep_next, &base.keep_next);
    merge_opt(&mut target.keep_lines, &base.keep_lines);
    merge_opt(&mut target.widow_control, &base.widow_control);
    merge_opt(&mut target.page_break_before, &base.page_break_before);
    merge_opt(&mut target.suppress_auto_hyphens, &base.suppress_auto_hyphens);
    merge_opt(&mut target.contextual_spacing, &base.contextual_spacing);
    merge_opt(&mut target.bidi, &base.bidi);
    merge_opt(&mut target.word_wrap, &base.word_wrap);
    merge_opt(&mut target.outline_level, &base.outline_level);
    merge_opt(&mut target.text_alignment, &base.text_alignment);
    merge_opt(&mut target.cnf_style, &base.cnf_style);
    merge_opt(&mut target.frame_properties, &base.frame_properties);
    merge_opt(&mut target.auto_space_de, &base.auto_space_de);
    merge_opt(&mut target.auto_space_dn, &base.auto_space_dn);

    // §17.3.1.38: merge tab stops at the individual-stop level.
    // Child Clear entries remove matching positions from the parent.
    // Child non-Clear entries are added. Result is sorted by position.
    if !base.tabs.is_empty() || !target.tabs.is_empty() {
        let child_tabs = std::mem::take(&mut target.tabs);
        // Start from parent tabs.
        target.tabs.clone_from(&base.tabs);
        // Remove positions that the child clears.
        for clear in child_tabs.iter().filter(|t| t.alignment == TabAlignment::Clear) {
            target.tabs.retain(|t| t.position != clear.position);
        }
        // Add child's non-Clear tabs (replacing any at the same position).
        for tab in child_tabs.iter().filter(|t| t.alignment != TabAlignment::Clear) {
            target.tabs.retain(|t| t.position != tab.position);
            target.tabs.push(*tab);
        }
        target.tabs.sort_by(|a, b| a.position.raw().cmp(&b.position.raw()));
    }
}

/// If `target` is `None`, clone `base` into it.
fn merge_opt<T: Clone>(target: &mut Option<T>, base: &Option<T>) {
    if target.is_none() {
        *target = base.clone();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::dimension::{Dimension, HalfPoints, Twips};
    use dxpdf_docx_model::model::*;

    // ── RunProperties merging ────────────────────────────────────────────

    #[test]
    fn merge_run_empty_target_takes_all_from_base() {
        let mut target = RunProperties::default();
        let base = RunProperties {
            bold: Some(true),
            italic: Some(true),
            font_size: Some(Dimension::<HalfPoints>::new(24)),
            color: Some(Color::Rgb(0xFF0000)),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);

        assert_eq!(target.bold, Some(true));
        assert_eq!(target.italic, Some(true));
        assert_eq!(target.font_size, Some(Dimension::<HalfPoints>::new(24)));
        assert_eq!(target.color, Some(Color::Rgb(0xFF0000)));
    }

    #[test]
    fn merge_run_target_values_not_overwritten() {
        let mut target = RunProperties {
            bold: Some(false),
            font_size: Some(Dimension::<HalfPoints>::new(20)),
            ..Default::default()
        };
        let base = RunProperties {
            bold: Some(true),
            italic: Some(true),
            font_size: Some(Dimension::<HalfPoints>::new(24)),
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);

        assert_eq!(target.bold, Some(false), "target's bold should win");
        assert_eq!(target.font_size, Some(Dimension::<HalfPoints>::new(20)), "target's size should win");
        assert_eq!(target.italic, Some(true), "italic should come from base");
    }

    #[test]
    fn merge_run_both_empty_stays_empty() {
        let mut target = RunProperties::default();
        let base = RunProperties::default();
        merge_run_properties(&mut target, &base);

        assert_eq!(target.bold, None);
        assert_eq!(target.italic, None);
        assert_eq!(target.font_size, None);
    }

    #[test]
    fn merge_run_fonts_merged_field_by_field() {
        let mut target = RunProperties {
            fonts: FontSet {
                ascii: Some("Arial".into()),
                ..Default::default()
            },
            ..Default::default()
        };
        let base = RunProperties {
            fonts: FontSet {
                ascii: Some("Times".into()),
                east_asian: Some("SimSun".into()),
                ..Default::default()
            },
            ..Default::default()
        };
        merge_run_properties(&mut target, &base);

        assert_eq!(target.fonts.ascii.as_deref(), Some("Arial"), "target's ascii should win");
        assert_eq!(target.fonts.east_asian.as_deref(), Some("SimSun"), "east_asian should come from base");
    }

    #[test]
    fn merge_run_all_fields_covered() {
        // Ensure merge touches every field by setting all in base, none in target.
        let base = RunProperties {
            fonts: FontSet {
                ascii: Some("F".into()),
                high_ansi: Some("F".into()),
                east_asian: Some("F".into()),
                complex_script: Some("F".into()),
            },
            font_size: Some(Dimension::<HalfPoints>::new(24)),
            bold: Some(true),
            italic: Some(true),
            underline: Some(UnderlineStyle::Single),
            strike: Some(StrikeStyle::Single),
            color: Some(Color::Rgb(0)),
            highlight: Some(HighlightColor::Yellow),
            shading: Some(Shading {
                fill: Color::Rgb(0),
                pattern: ShadingPattern::Clear,
                color: Color::Rgb(0),
            }),
            vertical_align: Some(VerticalAlign::Superscript),
            spacing: Some(Dimension::<Twips>::new(10)),
            kerning: Some(Dimension::<HalfPoints>::new(2)),
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
            position: Some(Dimension::<HalfPoints>::new(5)),
            lang: Some(Lang {
                val: Some("en".into()),
                east_asia: None,
                bidi: None,
            }),
            border: Some(Border {
                style: BorderStyle::Single,
                width: Dimension::new(0),
                space: Dimension::new(0),
                color: Color::BLACK,
            }),
        };
        let mut target = RunProperties::default();
        merge_run_properties(&mut target, &base);

        assert!(target.bold.is_some());
        assert!(target.italic.is_some());
        assert!(target.underline.is_some());
        assert!(target.strike.is_some());
        assert!(target.color.is_some());
        assert!(target.highlight.is_some());
        assert!(target.shading.is_some());
        assert!(target.vertical_align.is_some());
        assert!(target.spacing.is_some());
        assert!(target.kerning.is_some());
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
        assert!(target.position.is_some());
        assert!(target.lang.is_some());
        assert!(target.border.is_some());
        assert!(target.fonts.ascii.is_some());
        assert!(target.fonts.high_ansi.is_some());
        assert!(target.fonts.east_asian.is_some());
        assert!(target.fonts.complex_script.is_some());
        assert!(target.font_size.is_some());
    }

    // ── ParagraphProperties merging ──────────────────────────────────────

    #[test]
    fn merge_para_empty_target_takes_from_base() {
        let mut target = ParagraphProperties::default();
        let base = ParagraphProperties {
            alignment: Some(Alignment::Center),
            keep_next: Some(true),
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert_eq!(target.alignment, Some(Alignment::Center));
        assert_eq!(target.keep_next, Some(true));
    }

    #[test]
    fn merge_para_target_values_not_overwritten() {
        let mut target = ParagraphProperties {
            alignment: Some(Alignment::End),
            ..Default::default()
        };
        let base = ParagraphProperties {
            alignment: Some(Alignment::Center),
            keep_next: Some(true),
            ..Default::default()
        };
        merge_paragraph_properties(&mut target, &base);

        assert_eq!(target.alignment, Some(Alignment::End), "target should win");
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
        assert_eq!(target.tabs[0].position, Dimension::<Twips>::new(720), "sorted: left@720 first");
        assert_eq!(target.tabs[1].position, Dimension::<Twips>::new(1440), "sorted: right@1440 second");
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

        assert!(target.tabs.is_empty(), "cleared tab should be gone, nothing added");
    }

    #[test]
    fn merge_para_all_fields_covered() {
        let base = ParagraphProperties {
            alignment: Some(Alignment::Start),
            indentation: Some(Indentation::default()),
            spacing: Some(ParagraphSpacing::default()),
            numbering: Some(NumberingReference { num_id: 1, level: 0 }),
            tabs: vec![TabStop {
                position: Dimension::new(720),
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }],
            borders: Some(ParagraphBorders {
                top: None, bottom: None, left: None, right: None, between: None,
            }),
            shading: Some(Shading {
                fill: Color::Rgb(0),
                pattern: ShadingPattern::Clear,
                color: Color::Rgb(0),
            }),
            keep_next: Some(true),
            keep_lines: Some(true),
            widow_control: Some(true),
            page_break_before: Some(true),
            suppress_auto_hyphens: Some(true),
            contextual_spacing: Some(true),
            bidi: Some(true),
            word_wrap: Some(true),
            outline_level: Some(OutlineLevel::new(1)),
            text_alignment: Some(TextAlignment::Center),
            cnf_style: Some(CnfStyle {
                val: None, first_row: Some(true), last_row: None,
                first_column: None, last_column: None, odd_v_band: None,
                even_v_band: None, odd_h_band: None, even_h_band: None,
                first_row_first_column: None, first_row_last_column: None,
                last_row_first_column: None, last_row_last_column: None,
            }),
            frame_properties: None,
            auto_space_de: Some(true),
            auto_space_dn: Some(true),
        };
        let mut target = ParagraphProperties::default();
        merge_paragraph_properties(&mut target, &base);

        assert!(target.alignment.is_some());
        assert!(target.indentation.is_some());
        assert!(target.spacing.is_some());
        assert!(target.numbering.is_some());
        assert!(!target.tabs.is_empty());
        assert!(target.borders.is_some());
        assert!(target.shading.is_some());
        assert!(target.keep_next.is_some());
        assert!(target.keep_lines.is_some());
        assert!(target.widow_control.is_some());
        assert!(target.page_break_before.is_some());
        assert!(target.suppress_auto_hyphens.is_some());
        assert!(target.contextual_spacing.is_some());
        assert!(target.bidi.is_some());
        assert!(target.word_wrap.is_some());
        assert!(target.outline_level.is_some());
        assert!(target.text_alignment.is_some());
        assert!(target.cnf_style.is_some());
        assert!(target.auto_space_de.is_some());
        assert!(target.auto_space_dn.is_some());
    }
}
