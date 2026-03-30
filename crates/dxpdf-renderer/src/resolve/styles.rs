//! Style inheritance resolution — walk `basedOn` chains and merge properties.

use std::collections::{HashMap, HashSet};

use dxpdf_docx_model::model::{
    ParagraphProperties, RunProperties, StyleId, StyleSheet, TableProperties, Theme,
};

use super::properties::{merge_paragraph_properties, merge_run_properties, merge_table_properties};

/// A fully resolved style — all `basedOn` inheritance has been applied.
#[derive(Clone, Debug)]
pub struct ResolvedStyle {
    pub paragraph: ParagraphProperties,
    pub run: RunProperties,
    pub table: Option<TableProperties>,
    /// §17.7.6.6: table style conditional formatting overrides.
    pub table_style_overrides: Vec<dxpdf_docx_model::model::TableStyleOverride>,
}

/// Resolve all styles in the stylesheet by walking `basedOn` chains.
/// The theme is used to resolve `asciiTheme` / `hAnsiTheme` font references
/// on style-level run properties (§17.3.2.26).
pub fn resolve_styles(sheet: &StyleSheet, theme: Option<&Theme>) -> HashMap<StyleId, ResolvedStyle> {
    let mut resolved: HashMap<StyleId, ResolvedStyle> = HashMap::new();

    for id in sheet.styles.keys() {
        if !resolved.contains_key(id) {
            resolve_one(id, sheet, theme, &mut resolved, &mut HashSet::new());
        }
    }

    resolved
}

/// Recursively resolve a single style, memoizing results.
/// `visiting` tracks the current chain for cycle detection.
fn resolve_one(
    id: &StyleId,
    sheet: &StyleSheet,
    theme: Option<&Theme>,
    resolved: &mut HashMap<StyleId, ResolvedStyle>,
    visiting: &mut HashSet<StyleId>,
) {
    if resolved.contains_key(id) {
        return;
    }

    let style = match sheet.styles.get(id) {
        Some(s) => s,
        None => return,
    };

    // Cycle detection: if we're already visiting this style, stop recursion.
    if !visiting.insert(id.clone()) {
        // Break the cycle — resolve with just own properties + doc defaults.
        let mut para = style
            .paragraph_properties
            .clone()
            .unwrap_or_default();
        let mut run = style.run_properties.clone().unwrap_or_default();
        let table = style.table_properties.clone();
        merge_paragraph_properties(&mut para, &sheet.doc_defaults_paragraph);
        merge_run_properties(&mut run, &sheet.doc_defaults_run);
        resolved.insert(
            id.clone(),
            ResolvedStyle {
                paragraph: para,
                run,
                table,
                table_style_overrides: style.table_style_overrides.clone(),
            },
        );
        return;
    }

    // Resolve parent first (if any).
    if let Some(ref parent_id) = style.based_on {
        if !resolved.contains_key(parent_id) {
            resolve_one(parent_id, sheet, theme, resolved, visiting);
        }
    }

    // Start with own properties.
    let mut para = style
        .paragraph_properties
        .clone()
        .unwrap_or_default();
    let mut run = style.run_properties.clone().unwrap_or_default();
    // §17.3.2.26: resolve theme font references on the style's own run properties.
    if let Some(th) = theme {
        super::fonts::resolve_font_set_themes(&mut run.fonts, th);
    }

    // §17.7.2: table property inheritance — cell margins from parent table styles.
    let mut table = style.table_properties.clone();
    // Merge from resolved parent (if it exists and was successfully resolved).
    if let Some(ref parent_id) = style.based_on {
        if let Some(parent_resolved) = resolved.get(parent_id) {
            merge_paragraph_properties(&mut para, &parent_resolved.paragraph);
            merge_run_properties(&mut run, &parent_resolved.run);
            merge_table_properties(&mut table, &parent_resolved.table);
        }
    }

    // §17.7.2: doc defaults are NOT merged here — they are merged by the
    // caller at the correct cascade level. For table cell paragraphs, the
    // table style must be able to override doc defaults, which requires
    // doc defaults to be the lowest priority in the merge chain.
    // Character styles inherit only from their basedOn chain.
    // Run defaults from docDefaults still apply for font resolution.
    if style.style_type != dxpdf_docx_model::model::StyleType::Character {
        merge_run_properties(&mut run, &sheet.doc_defaults_run);
    }

    visiting.remove(id);

    resolved.insert(
        id.clone(),
        ResolvedStyle {
            paragraph: para,
            run,
            table,
                table_style_overrides: style.table_style_overrides.clone(),
        },
    );
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::dimension::{Dimension, HalfPoints};
    use dxpdf_docx_model::model::*;

    fn make_sheet(styles: Vec<(StyleId, Style)>) -> StyleSheet {
        StyleSheet {
            styles: styles.into_iter().collect(),
            ..Default::default()
        }
    }

    fn style(
        based_on: Option<&str>,
        para: Option<ParagraphProperties>,
        run: Option<RunProperties>,
    ) -> Style {
        Style {
            name: None,
            style_type: StyleType::Paragraph,
            based_on: based_on.map(StyleId::new),
            is_default: false,
            paragraph_properties: para,
            run_properties: run,
            table_properties: None,
            table_style_overrides: vec![],
        }
    }

    #[test]
    fn single_style_no_inheritance() {
        let sheet = make_sheet(vec![(
            StyleId::new("Normal"),
            style(
                None,
                Some(ParagraphProperties {
                    alignment: Some(Alignment::Start),
                    ..Default::default()
                }),
                Some(RunProperties {
                    bold: Some(false),
                    font_size: Some(Dimension::<HalfPoints>::new(24)),
                    ..Default::default()
                }),
            ),
        )]);

        let resolved = resolve_styles(&sheet, None);
        let normal = resolved.get(&StyleId::new("Normal")).unwrap();

        assert_eq!(normal.paragraph.alignment, Some(Alignment::Start));
        assert_eq!(normal.run.bold, Some(false));
        assert_eq!(normal.run.font_size, Some(Dimension::<HalfPoints>::new(24)));
    }

    #[test]
    fn child_inherits_from_parent() {
        let sheet = make_sheet(vec![
            (
                StyleId::new("Normal"),
                style(
                    None,
                    Some(ParagraphProperties {
                        alignment: Some(Alignment::Start),
                        ..Default::default()
                    }),
                    Some(RunProperties {
                        font_size: Some(Dimension::<HalfPoints>::new(24)),
                        bold: Some(false),
                        ..Default::default()
                    }),
                ),
            ),
            (
                StyleId::new("Heading1"),
                style(
                    Some("Normal"),
                    Some(ParagraphProperties {
                        alignment: Some(Alignment::Center),
                        ..Default::default()
                    }),
                    Some(RunProperties {
                        bold: Some(true),
                        ..Default::default()
                    }),
                ),
            ),
        ]);

        let resolved = resolve_styles(&sheet, None);
        let h1 = resolved.get(&StyleId::new("Heading1")).unwrap();

        assert_eq!(h1.paragraph.alignment, Some(Alignment::Center), "child overrides parent");
        assert_eq!(h1.run.bold, Some(true), "child overrides parent");
        assert_eq!(
            h1.run.font_size,
            Some(Dimension::<HalfPoints>::new(24)),
            "inherited from Normal"
        );
    }

    #[test]
    fn three_level_chain() {
        let sheet = make_sheet(vec![
            (
                StyleId::new("Base"),
                style(
                    None,
                    None,
                    Some(RunProperties {
                        font_size: Some(Dimension::<HalfPoints>::new(20)),
                        bold: Some(false),
                        italic: Some(false),
                        ..Default::default()
                    }),
                ),
            ),
            (
                StyleId::new("Mid"),
                style(
                    Some("Base"),
                    None,
                    Some(RunProperties {
                        bold: Some(true),
                        ..Default::default()
                    }),
                ),
            ),
            (
                StyleId::new("Leaf"),
                style(
                    Some("Mid"),
                    None,
                    Some(RunProperties {
                        italic: Some(true),
                        ..Default::default()
                    }),
                ),
            ),
        ]);

        let resolved = resolve_styles(&sheet, None);
        let leaf = resolved.get(&StyleId::new("Leaf")).unwrap();

        assert_eq!(leaf.run.italic, Some(true), "own value");
        assert_eq!(leaf.run.bold, Some(true), "from Mid");
        assert_eq!(leaf.run.font_size, Some(Dimension::<HalfPoints>::new(20)), "from Base");
    }

    #[test]
    fn cycle_does_not_panic() {
        let sheet = make_sheet(vec![
            (
                StyleId::new("A"),
                style(Some("B"), None, Some(RunProperties {
                    bold: Some(true),
                    ..Default::default()
                })),
            ),
            (
                StyleId::new("B"),
                style(Some("A"), None, Some(RunProperties {
                    italic: Some(true),
                    ..Default::default()
                })),
            ),
        ]);

        // Should not infinite loop or panic
        let resolved = resolve_styles(&sheet, None);
        assert!(resolved.contains_key(&StyleId::new("A")));
        assert!(resolved.contains_key(&StyleId::new("B")));
    }

    #[test]
    fn missing_based_on_target_is_harmless() {
        let sheet = make_sheet(vec![(
            StyleId::new("Orphan"),
            style(
                Some("DoesNotExist"),
                None,
                Some(RunProperties {
                    bold: Some(true),
                    ..Default::default()
                }),
            ),
        )]);

        let resolved = resolve_styles(&sheet, None);
        let orphan = resolved.get(&StyleId::new("Orphan")).unwrap();
        assert_eq!(orphan.run.bold, Some(true));
    }

    #[test]
    fn doc_defaults_are_applied_as_base() {
        // §17.7.2: doc defaults paragraph properties are NOT merged during
        // style resolution — they are merged by the caller at the correct
        // cascade level. Only run defaults are merged into resolved styles.
        let sheet = StyleSheet {
            doc_defaults_paragraph: ParagraphProperties {
                alignment: Some(Alignment::Both),
                ..Default::default()
            },
            doc_defaults_run: RunProperties {
                font_size: Some(Dimension::<HalfPoints>::new(22)),
                ..Default::default()
            },
            styles: [(
                StyleId::new("Normal"),
                style(None, None, None),
            )]
            .into_iter()
            .collect(),
            latent_styles: None,
        };

        let resolved = resolve_styles(&sheet, None);
        let normal = resolved.get(&StyleId::new("Normal")).unwrap();

        // Paragraph doc defaults are deferred to the caller.
        assert_eq!(
            normal.paragraph.alignment,
            None,
            "paragraph doc defaults are not merged into resolved styles"
        );
        // Run doc defaults ARE merged during resolution.
        assert_eq!(
            normal.run.font_size,
            Some(Dimension::<HalfPoints>::new(22)),
            "should inherit from doc defaults"
        );
    }

    #[test]
    fn style_overrides_doc_defaults() {
        let sheet = StyleSheet {
            doc_defaults_run: RunProperties {
                font_size: Some(Dimension::<HalfPoints>::new(22)),
                bold: Some(false),
                ..Default::default()
            },
            styles: [(
                StyleId::new("Strong"),
                style(
                    None,
                    None,
                    Some(RunProperties {
                        bold: Some(true),
                        ..Default::default()
                    }),
                ),
            )]
            .into_iter()
            .collect(),
            ..Default::default()
        };

        let resolved = resolve_styles(&sheet, None);
        let strong = resolved.get(&StyleId::new("Strong")).unwrap();

        assert_eq!(strong.run.bold, Some(true), "style overrides doc default");
        assert_eq!(
            strong.run.font_size,
            Some(Dimension::<HalfPoints>::new(22)),
            "inherited from doc default"
        );
    }
}
