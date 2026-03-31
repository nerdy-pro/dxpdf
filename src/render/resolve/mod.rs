//! Resolve layer — transforms raw `Document` into layout-ready `ResolvedDocument`.

pub mod color;
pub mod conditional;
pub mod fonts;
pub mod images;
pub mod numbering;
pub mod properties;
pub mod sections;
pub mod styles;

use std::collections::HashMap;

use crate::model::{
    Block, Document, NoteId, NumId, NumPicBullet, NumPicBulletId, ParagraphProperties, RelId,
    RunProperties, StyleId, Theme,
};

use self::numbering::ResolvedNumberingLevel;
use self::sections::ResolvedSection;
use self::styles::ResolvedStyle;

/// A fully resolved document — ready for the layout pipeline.
/// All style inheritance resolved, sections split, headers/footers attached,
/// font families collected, image RelIds extracted.
#[derive(Debug)]
pub struct ResolvedDocument {
    /// Sections with their blocks, page geometry, and header/footer content.
    pub sections: Vec<ResolvedSection>,
    /// Fully resolved styles (basedOn chains walked, doc defaults applied).
    pub styles: HashMap<StyleId, ResolvedStyle>,
    /// Flattened numbering definitions.
    pub numbering: HashMap<NumId, Vec<ResolvedNumberingLevel>>,
    /// All unique font families referenced in the document.
    pub font_families: Vec<String>,
    /// Embedded media (images) — raw bytes keyed by relationship ID.
    pub media: HashMap<RelId, Vec<u8>>,
    /// §17.9.21: picture bullet definitions keyed by numPicBulletId.
    pub pic_bullets: HashMap<NumPicBulletId, NumPicBullet>,
    /// Theme (for color resolution during paint).
    pub theme: Option<Theme>,
    /// Document-level default paragraph properties (from docDefaults).
    pub doc_defaults_paragraph: ParagraphProperties,
    /// Document-level default run properties (from docDefaults).
    pub doc_defaults_run: RunProperties,
    /// §17.7.4.17: the default paragraph style (w:default="1", type="paragraph").
    /// Applied to paragraphs that don't specify a style explicitly.
    pub default_paragraph_style_id: Option<StyleId>,
    /// Footnote content keyed by note ID.
    pub footnotes: HashMap<NoteId, Vec<Block>>,
    /// Endnote content keyed by note ID.
    pub endnotes: HashMap<NoteId, Vec<Block>>,
}

/// Transform a raw parsed Document into a layout-ready ResolvedDocument.
pub fn resolve(doc: &Document) -> ResolvedDocument {
    use crate::model::StyleType;

    let styles = styles::resolve_styles(&doc.styles, doc.theme.as_ref());
    let numbering = numbering::resolve_numbering(&doc.numbering);
    let sections = sections::resolve_sections(doc);
    let font_families = fonts::collect_font_families(doc);

    // §17.7.4.17: find the default paragraph style.
    let default_paragraph_style_id = doc
        .styles
        .styles
        .iter()
        .find(|(_, s)| s.is_default && s.style_type == StyleType::Paragraph)
        .map(|(id, _)| id.clone());

    ResolvedDocument {
        sections,
        styles,
        numbering,
        font_families,
        media: doc.media.clone(),
        pic_bullets: doc.numbering.pic_bullets.clone(),
        doc_defaults_paragraph: doc.styles.doc_defaults_paragraph.clone(),
        doc_defaults_run: doc.styles.doc_defaults_run.clone(),
        default_paragraph_style_id,
        theme: doc.theme.clone(),
        footnotes: doc.footnotes.clone(),
        endnotes: doc.endnotes.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::{Dimension, HalfPoints};
    use crate::model::*;

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

    fn para(text: &str) -> Block {
        Block::Paragraph(Box::new(Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: vec![Inline::TextRun(Box::new(TextRun {
                style_id: None,
                properties: RunProperties {
                    fonts: FontSet {
                        ascii: Some("TestFont".into()),
                        ..Default::default()
                    },
                    ..Default::default()
                },
                content: vec![RunElement::Text(text.to_string())],
                rsids: RevisionIds::default(),
            }))],
            rsids: ParagraphRevisionIds::default(),
        }))
    }

    #[test]
    fn resolve_empty_doc() {
        let doc = empty_doc();
        let resolved = resolve(&doc);

        assert_eq!(resolved.sections.len(), 1);
        assert!(resolved.sections[0].blocks.is_empty());
        assert!(resolved.styles.is_empty());
        assert!(resolved.numbering.is_empty());
        assert!(resolved.font_families.is_empty());
        assert!(resolved.media.is_empty());
        assert!(resolved.theme.is_none());
    }

    #[test]
    fn resolve_preserves_body_content() {
        let mut doc = empty_doc();
        doc.body = vec![para("hello"), para("world")];

        let resolved = resolve(&doc);
        assert_eq!(resolved.sections.len(), 1);
        assert_eq!(resolved.sections[0].blocks.len(), 2);
    }

    #[test]
    fn resolve_splits_sections() {
        let mut doc = empty_doc();
        doc.body = vec![
            para("first"),
            Block::SectionBreak(Box::default()),
            para("second"),
        ];

        let resolved = resolve(&doc);
        assert_eq!(resolved.sections.len(), 2);
        assert_eq!(resolved.sections[0].blocks.len(), 1);
        assert_eq!(resolved.sections[1].blocks.len(), 1);
    }

    #[test]
    fn resolve_resolves_styles() {
        let mut doc = empty_doc();
        doc.styles.doc_defaults_run = RunProperties {
            font_size: Some(Dimension::<HalfPoints>::new(22)),
            ..Default::default()
        };
        doc.styles.styles.insert(
            StyleId::new("Normal"),
            Style {
                name: None,
                style_type: StyleType::Paragraph,
                based_on: None,
                is_default: true,
                paragraph_properties: Some(ParagraphProperties {
                    alignment: Some(Alignment::Start),
                    ..Default::default()
                }),
                run_properties: None,
                table_properties: None,
                table_style_overrides: vec![],
            },
        );

        let resolved = resolve(&doc);
        let normal = resolved.styles.get(&StyleId::new("Normal")).unwrap();
        assert_eq!(normal.paragraph.alignment, Some(Alignment::Start));
        assert_eq!(
            normal.run.font_size,
            Some(Dimension::<HalfPoints>::new(22)),
            "should inherit doc default"
        );
    }

    #[test]
    fn resolve_collects_fonts() {
        let mut doc = empty_doc();
        doc.body = vec![para("text")];

        let resolved = resolve(&doc);
        assert!(resolved.font_families.contains(&"TestFont".to_string()));
    }

    #[test]
    fn resolve_resolves_numbering() {
        let mut doc = empty_doc();
        doc.numbering.abstract_nums.insert(
            AbstractNumId::new(0),
            AbstractNumbering {
                levels: vec![NumberingLevelDefinition {
                    level: 0,
                    format: Some(NumberFormat::Decimal),
                    level_text: "%1.".into(),
                    start: Some(1),
                    justification: None,
                    indentation: None,
                    run_properties: None,
                    lvl_pic_bullet_id: None,
                }],
            },
        );
        doc.numbering.numbering_instances.insert(
            NumId::new(1),
            NumberingInstance {
                abstract_num_id: AbstractNumId::new(0),
                level_overrides: vec![],
            },
        );

        let resolved = resolve(&doc);
        let levels = resolved.numbering.get(&NumId::new(1)).unwrap();
        assert_eq!(levels.len(), 1);
        assert_eq!(levels[0].format, NumberFormat::Decimal);
    }

    #[test]
    fn resolve_preserves_media() {
        let mut doc = empty_doc();
        doc.media.insert(RelId::new("rId1"), vec![0xFF, 0xD8, 0xFF]);

        let resolved = resolve(&doc);
        assert!(resolved.media.contains_key(&RelId::new("rId1")));
        assert_eq!(resolved.media[&RelId::new("rId1")], vec![0xFF, 0xD8, 0xFF]);
    }

    #[test]
    fn resolve_preserves_theme() {
        let mut doc = empty_doc();
        doc.theme = Some(Theme {
            color_scheme: ThemeColorScheme {
                accent1: 0x4472C4,
                ..Default::default()
            },
            ..Default::default()
        });

        let resolved = resolve(&doc);
        assert!(resolved.theme.is_some());
        assert_eq!(resolved.theme.unwrap().color_scheme.accent1, 0x4472C4);
    }

    #[test]
    fn resolve_headers_attached_to_sections() {
        let mut doc = empty_doc();
        let hdr_id = RelId::new("rId1");
        doc.headers.insert(hdr_id.clone(), vec![para("header")]);
        doc.final_section = SectionProperties {
            header_refs: SectionHeaderFooterRefs {
                default: Some(hdr_id),
                ..Default::default()
            },
            ..Default::default()
        };
        doc.body = vec![para("body")];

        let resolved = resolve(&doc);
        assert!(resolved.sections[0].header.is_some());
    }
}
