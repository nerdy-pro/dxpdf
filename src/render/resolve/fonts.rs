//! Font collection — extract all font families referenced in the document.

use std::collections::HashSet;

use crate::model::{Block, Document, FontSet, Inline};

/// Collect all unique font family names referenced in the document.
/// Sources: theme, style sheet, inline content, numbering levels, paragraph marks.
pub fn collect_font_families(doc: &Document) -> Vec<String> {
    let mut families = HashSet::new();

    // Theme fonts
    if let Some(ref theme) = doc.theme {
        add_nonempty(&mut families, &theme.major_font.latin);
        add_nonempty(&mut families, &theme.minor_font.latin);
        add_nonempty(&mut families, &theme.major_font.east_asian);
        add_nonempty(&mut families, &theme.minor_font.east_asian);
        add_nonempty(&mut families, &theme.major_font.complex_script);
        add_nonempty(&mut families, &theme.minor_font.complex_script);
        for sf in &theme.major_font.script_fonts {
            add_nonempty(&mut families, &sf.typeface);
        }
        for sf in &theme.minor_font.script_fonts {
            add_nonempty(&mut families, &sf.typeface);
        }
    }

    // Style sheet defaults
    collect_from_fontset(&mut families, &doc.styles.doc_defaults_run.fonts);

    // Style definitions
    for style in doc.styles.styles.values() {
        if let Some(ref rp) = style.run_properties {
            collect_from_fontset(&mut families, &rp.fonts);
        }
    }

    // Body content
    collect_from_blocks(&mut families, &doc.body);

    // Headers and footers
    for blocks in doc.headers.values() {
        collect_from_blocks(&mut families, blocks);
    }
    for blocks in doc.footers.values() {
        collect_from_blocks(&mut families, blocks);
    }

    // Footnotes and endnotes
    for blocks in doc.footnotes.values() {
        collect_from_blocks(&mut families, blocks);
    }
    for blocks in doc.endnotes.values() {
        collect_from_blocks(&mut families, blocks);
    }

    // Numbering levels
    for abs in doc.numbering.abstract_nums.values() {
        for level in &abs.levels {
            if let Some(ref rp) = level.run_properties {
                collect_from_fontset(&mut families, &rp.fonts);
            }
        }
    }

    families.into_iter().collect()
}

/// Extract the effective font family from a FontSet.
/// Priority: ascii > high_ansi > east_asian > complex_script.
pub fn effective_font(fonts: &FontSet) -> Option<&str> {
    fonts
        .ascii
        .as_deref()
        .or(fonts.high_ansi.as_deref())
        .or(fonts.east_asian.as_deref())
        .or(fonts.complex_script.as_deref())
}

/// §17.3.2.26: resolve theme font references in a FontSet.
///
/// For each slot where a theme reference is present, resolve it to the
/// actual font family name from the theme. The resolved name is written
/// into the corresponding name field (ascii, high_ansi, etc.), overwriting
/// any explicit name — theme references take precedence per §17.3.2.26.
pub fn resolve_font_set_themes(fonts: &mut FontSet, theme: &crate::model::Theme) {
    if let Some(ref tf) = fonts.ascii_theme {
        if let Some(name) = resolve_theme_font_ref(tf, theme) {
            fonts.ascii = Some(name);
        }
    }
    if let Some(ref tf) = fonts.high_ansi_theme {
        if let Some(name) = resolve_theme_font_ref(tf, theme) {
            fonts.high_ansi = Some(name);
        }
    }
    if let Some(ref tf) = fonts.east_asian_theme {
        if let Some(name) = resolve_theme_font_ref(tf, theme) {
            fonts.east_asian = Some(name);
        }
    }
    if let Some(ref tf) = fonts.complex_script_theme {
        if let Some(name) = resolve_theme_font_ref(tf, theme) {
            fonts.complex_script = Some(name);
        }
    }
}

fn resolve_theme_font_ref(
    tf: &crate::model::ThemeFontRef,
    theme: &crate::model::Theme,
) -> Option<String> {
    use crate::model::ThemeFontRef;
    let name = match tf {
        ThemeFontRef::MajorHAnsi => &theme.major_font.latin,
        ThemeFontRef::MajorEastAsia => &theme.major_font.east_asian,
        ThemeFontRef::MajorBidi => &theme.major_font.complex_script,
        ThemeFontRef::MinorHAnsi => &theme.minor_font.latin,
        ThemeFontRef::MinorEastAsia => &theme.minor_font.east_asian,
        ThemeFontRef::MinorBidi => &theme.minor_font.complex_script,
    };
    if name.is_empty() {
        None
    } else {
        Some(name.clone())
    }
}

fn add_nonempty(set: &mut HashSet<String>, s: &str) {
    if !s.is_empty() {
        set.insert(s.to_string());
    }
}

fn collect_from_fontset(set: &mut HashSet<String>, fonts: &FontSet) {
    if let Some(ref f) = fonts.ascii {
        add_nonempty(set, f);
    }
    if let Some(ref f) = fonts.high_ansi {
        add_nonempty(set, f);
    }
    if let Some(ref f) = fonts.east_asian {
        add_nonempty(set, f);
    }
    if let Some(ref f) = fonts.complex_script {
        add_nonempty(set, f);
    }
}

fn collect_from_blocks(set: &mut HashSet<String>, blocks: &[Block]) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                if let Some(ref mrp) = p.mark_run_properties {
                    collect_from_fontset(set, &mrp.fonts);
                }
                collect_from_inlines(set, &p.content);
            }
            Block::Table(t) => {
                for row in &t.rows {
                    for cell in &row.cells {
                        collect_from_blocks(set, &cell.content);
                    }
                }
            }
            Block::SectionBreak(_) => {}
        }
    }
}

fn collect_from_inlines(set: &mut HashSet<String>, inlines: &[Inline]) {
    for inline in inlines {
        match inline {
            Inline::TextRun(tr) => {
                collect_from_fontset(set, &tr.properties.fonts);
            }
            Inline::Hyperlink(h) => {
                collect_from_inlines(set, &h.content);
            }
            Inline::Field(f) => {
                collect_from_inlines(set, &f.content);
            }
            Inline::AlternateContent(ac) => {
                for choice in &ac.choices {
                    collect_from_inlines(set, &choice.content);
                }
                if let Some(ref fb) = ac.fallback {
                    collect_from_inlines(set, fb);
                }
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::*;
    use std::collections::HashMap;

    fn empty_doc() -> Document {
        Document {
            settings: DocumentSettings::default(),
            theme: None,
            styles: StyleSheet::default(),
            numbering: NumberingDefinitions::default(),
            body: vec![],
            final_section: SectionProperties::default(),
            headers: HashMap::new(),
            footers: HashMap::new(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            media: HashMap::new(),
            embedded_fonts: vec![],
        }
    }

    fn text_run(font: &str, text: &str) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties {
                fonts: FontSet {
                    ascii: Some(font.into()),
                    ..Default::default()
                },
                ..Default::default()
            },
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    fn para_with_run(font: &str) -> Block {
        Block::Paragraph(Box::new(Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: vec![text_run(font, "hello")],
            rsids: ParagraphRevisionIds::default(),
        }))
    }

    // ── effective_font ───────────────────────────────────────────────────

    #[test]
    fn effective_font_prefers_ascii() {
        let fs = FontSet {
            ascii: Some("Arial".into()),
            high_ansi: Some("Times".into()),
            ..Default::default()
        };
        assert_eq!(effective_font(&fs), Some("Arial"));
    }

    #[test]
    fn effective_font_falls_back_to_high_ansi() {
        let fs = FontSet {
            high_ansi: Some("Times".into()),
            ..Default::default()
        };
        assert_eq!(effective_font(&fs), Some("Times"));
    }

    #[test]
    fn effective_font_falls_back_to_east_asian() {
        let fs = FontSet {
            east_asian: Some("SimSun".into()),
            ..Default::default()
        };
        assert_eq!(effective_font(&fs), Some("SimSun"));
    }

    #[test]
    fn effective_font_empty_returns_none() {
        let fs = FontSet::default();
        assert_eq!(effective_font(&fs), None);
    }

    // ── collect_font_families ────────────────────────────────────────────

    #[test]
    fn collects_from_body_text_runs() {
        let mut doc = empty_doc();
        doc.body = vec![para_with_run("Calibri"), para_with_run("Arial")];

        let families = collect_font_families(&doc);
        assert!(families.contains(&"Calibri".to_string()));
        assert!(families.contains(&"Arial".to_string()));
    }

    #[test]
    fn collects_from_style_sheet_defaults() {
        let mut doc = empty_doc();
        doc.styles.doc_defaults_run = RunProperties {
            fonts: FontSet {
                ascii: Some("Cambria".into()),
                ..Default::default()
            },
            ..Default::default()
        };

        let families = collect_font_families(&doc);
        assert!(families.contains(&"Cambria".to_string()));
    }

    #[test]
    fn collects_from_theme_fonts() {
        let mut doc = empty_doc();
        doc.theme = Some(Theme {
            major_font: ThemeFontScheme {
                latin: "Calibri Light".into(),
                ..Default::default()
            },
            minor_font: ThemeFontScheme {
                latin: "Calibri".into(),
                ..Default::default()
            },
            ..Default::default()
        });

        let families = collect_font_families(&doc);
        assert!(families.contains(&"Calibri Light".to_string()));
        assert!(families.contains(&"Calibri".to_string()));
    }

    #[test]
    fn collects_from_headers() {
        let mut doc = empty_doc();
        let header_id = RelId::new("rId1");
        doc.headers
            .insert(header_id, vec![para_with_run("Georgia")]);

        let families = collect_font_families(&doc);
        assert!(families.contains(&"Georgia".to_string()));
    }

    #[test]
    fn deduplicates() {
        let mut doc = empty_doc();
        doc.body = vec![para_with_run("Arial"), para_with_run("Arial")];

        let families = collect_font_families(&doc);
        let count = families.iter().filter(|f| *f == "Arial").count();
        assert_eq!(count, 1);
    }

    #[test]
    fn empty_doc_returns_empty() {
        let doc = empty_doc();
        let families = collect_font_families(&doc);
        assert!(families.is_empty());
    }

    #[test]
    fn collects_from_style_definitions() {
        let mut doc = empty_doc();
        doc.styles.styles.insert(
            StyleId::new("Heading1"),
            Style {
                name: None,
                style_type: StyleType::Paragraph,
                based_on: None,
                is_default: false,
                paragraph_properties: None,
                run_properties: Some(RunProperties {
                    fonts: FontSet {
                        ascii: Some("Verdana".into()),
                        ..Default::default()
                    },
                    ..Default::default()
                }),
                table_properties: None,
                table_style_overrides: vec![],
            },
        );

        let families = collect_font_families(&doc);
        assert!(families.contains(&"Verdana".to_string()));
    }
}
