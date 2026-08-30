//! Resolve layer — transforms raw `Document` into layout-ready `ResolvedDocument`.

pub mod color;
pub mod conditional;
pub mod drawing_color;
pub mod fonts;
pub mod header_footer;
pub mod images;
pub mod locale;
pub mod numbering;
pub mod properties;
pub mod revision;
pub mod sections;
pub mod shape_geometry;
pub mod shape_visuals;
pub mod spellout;
pub mod styles;

use std::collections::HashMap;

use crate::model::dimension::{Dimension, Twips};
use crate::model::{
    Block, Document, EmbeddedFont, NoteId, NumId, NumPicBullet, NumPicBulletId,
    ParagraphProperties, RelId, RunProperties, StyleId, Theme,
};

use self::images::MediaEntry;
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
    /// Embedded media (images) — shared bytes with detected format, keyed by relationship ID.
    pub media: HashMap<RelId, MediaEntry>,
    /// §17.8.3: embedded fonts, carried through so the font registry can be
    /// built from the resolved document alone. They belong here rather than
    /// being read back off the `Document` because [`resolve`] consumes it.
    pub embedded_fonts: Vec<EmbeddedFont>,
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
    /// §17.7.4.17: the default table style (`w:default="1"`, `type="table"`) —
    /// Word's `TableNormal`. Applied to tables that name no style; a table that
    /// names one reaches this through that style's `basedOn` chain instead,
    /// which is how Word's built-in table styles are written.
    pub default_table_style_id: Option<StyleId>,
    /// Footnote content keyed by note ID.
    pub footnotes: HashMap<NoteId, Vec<Block>>,
    /// Endnote content keyed by note ID.
    pub endnotes: HashMap<NoteId, Vec<Block>>,
    /// §17.10.1 — when true, even-numbered pages use the section's
    /// `even` header/footer slot (and the `default` slot is restricted
    /// to odd pages). Without this flag, the `even` slots are dead
    /// data even if the document supplies them.
    pub even_and_odd_headers: bool,
    /// §17.15.1.25: the document's default tab-stop interval (`w:defaultTabStop`,
    /// spec default 720 twips). Consumed by paragraph tab layout.
    pub default_tab_stop: Dimension<Twips>,
    /// Issue #154: whether tracked-change marks are drawn (`w:revisionView`).
    /// When false the final text renders — deletions suppressed, insertions
    /// plain.
    pub show_ins_del_marks: bool,
    /// Issue #154: whether comment anchors and balloons are drawn.
    pub show_comment_marks: bool,
    /// Issue #154: revision-mark color per author, in first-appearance order
    /// — see `resolve::revision`.
    pub revision_colors: std::collections::HashMap<String, crate::render::resolve::color::RgbColor>,
    /// Issue #154: comment content keyed by comment ID, for the balloon pass.
    pub comments: HashMap<crate::model::CommentId, crate::model::Comment>,
}

/// Transform a raw parsed Document into a layout-ready ResolvedDocument.
///
/// Takes the document **by value**. Resolve is the parse output's last reader —
/// nothing downstream can reach a `Document` — so every part it carries forward
/// unchanged (body blocks, note bodies, theme, picture bullets, embedded fonts,
/// media handles) is *moved*. Borrowing instead forced a deep clone of the
/// entire block tree plus a second copy of every image, which on a large or
/// image-heavy document was the largest allocation in the pipeline outside
/// paint.
///
/// The two derived maps — resolved styles and flattened numbering — still read
/// their sources by reference, because each produces a new value rather than
/// re-homing the old one.
pub fn resolve(doc: Document) -> ResolvedDocument {
    use crate::model::StyleType;

    let font_families = fonts::collect_font_families(&doc);

    // §17.7.4.17: the default style of each type that has a consumer — the one
    // applied to an object of that type which references no style by name.
    let default_style_of = |ty: StyleType| {
        doc.styles
            .styles
            .iter()
            .find(|(_, s)| s.is_default && s.style_type == ty)
            .map(|(id, _)| id.clone())
    };
    let default_paragraph_style_id = default_style_of(StyleType::Paragraph);
    let default_table_style_id = default_style_of(StyleType::Table);

    // Destructured rather than field-by-field so that adding a field to
    // `Document` breaks this build: an unhandled part is a silently dropped
    // one.
    let Document {
        settings,
        theme,
        styles,
        numbering,
        body,
        final_section,
        headers,
        footers,
        footnotes,
        endnotes,
        comments,
        media,
        embedded_fonts,
    } = doc;

    // Before `resolve_sections` consumes the body: the palette walk needs it
    // borrowed. Headers, footers and notes are not walked — a revision there
    // still gets a mark, in the fallback color (see `resolve::revision`).
    // Comment authors join the same numbering, in id order (deterministic
    // where map iteration is not), so a person who both edited and commented
    // keeps one color.
    let mut revision_colors = revision::collect_revision_colors(&body);
    // Issue #154: a range marker whose id the comments part never defined
    // must not wash anything — Word ignores such orphans. Cleared here, where
    // both sides are in hand, so layout can trust every stamp it sees.
    let mut body = body;
    revision::clear_orphan_comment_stamps(&mut body, &comments);
    {
        let mut ids: Vec<_> = comments.keys().copied().collect();
        ids.sort();
        for id in ids {
            revision::register(&comments[&id].author, &mut revision_colors);
        }
    }

    let resolved_styles = styles::resolve_styles(&styles, theme.as_ref());
    let resolved_numbering = numbering::resolve_numbering(&numbering);
    let sections = sections::resolve_sections(body, final_section, &headers, &footers);

    ResolvedDocument {
        sections,
        styles: resolved_styles,
        numbering: resolved_numbering,
        font_families,
        media: media
            .into_iter()
            .map(|(id, (data, format))| (id, MediaEntry { data, format }))
            .collect(),
        pic_bullets: numbering.pic_bullets,
        doc_defaults_paragraph: styles.doc_defaults_paragraph,
        doc_defaults_run: styles.doc_defaults_run,
        default_paragraph_style_id,
        default_table_style_id,
        theme,
        footnotes,
        endnotes,
        embedded_fonts,
        even_and_odd_headers: settings.even_and_odd_headers,
        default_tab_stop: settings.default_tab_stop,
        show_ins_del_marks: settings.show_ins_del_marks,
        show_comment_marks: settings.show_comment_marks,
        revision_colors,
        comments,
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
            comments: Default::default(),
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
                        ascii: FontSlot::from_name("TestFont"),
                        ..Default::default()
                    },
                    ..Default::default()
                },
                content: vec![RunElement::Text(text.to_string())],
                rsids: RevisionIds::default(),
                revision: None,
                comment: None,
            }))],
            rsids: ParagraphRevisionIds::default(),
            mark_deleted: false,
        }))
    }

    #[test]
    fn resolve_empty_doc() {
        let doc = empty_doc();
        let resolved = resolve(doc);

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

        let resolved = resolve(doc);
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

        let resolved = resolve(doc);
        assert_eq!(resolved.sections.len(), 2);
        assert_eq!(resolved.sections[0].blocks.len(), 1);
        assert_eq!(resolved.sections[1].blocks.len(), 1);
    }

    #[test]
    fn resolve_resolves_styles() {
        let mut doc = empty_doc();
        doc.styles.doc_defaults_run = RunProperties {
            font_size: Dup::from(Some(Dimension::<HalfPoints>::new(22))),
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
                    alignment: Dup::from(Some(Alignment::Start)),
                    ..Default::default()
                }),
                run_properties: None,
                table_properties: None,
                table_style_overrides: vec![],
            },
        );

        let resolved = resolve(doc);
        let normal = resolved.styles.get(&StyleId::new("Normal")).unwrap();
        assert_eq!(
            normal.paragraph.alignment,
            Dup::from(Some(Alignment::Start))
        );
        assert_eq!(
            normal.run.font_size,
            Dup::from(Some(Dimension::<HalfPoints>::new(22))),
            "should inherit doc default"
        );
    }

    #[test]
    fn resolve_collects_fonts() {
        let mut doc = empty_doc();
        doc.body = vec![para("text")];

        let resolved = resolve(doc);
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
                    suffix: crate::model::LevelSuffix::default(),
                    is_legal: false,
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

        let resolved = resolve(doc);
        let levels = resolved.numbering.get(&NumId::new(1)).unwrap();
        assert_eq!(levels.len(), 1);
        assert_eq!(levels[0].format, NumberFormat::Decimal);
    }

    #[test]
    fn resolve_preserves_media() {
        let mut doc = empty_doc();
        use crate::model::ImageFormat;
        doc.media.insert(
            RelId::new("rId1"),
            (
                std::sync::Arc::from(&[0xFF_u8, 0xD8, 0xFF][..]),
                ImageFormat::Jpeg,
            ),
        );

        let resolved = resolve(doc);
        assert!(resolved.media.contains_key(&RelId::new("rId1")));
        let entry = &resolved.media[&RelId::new("rId1")];
        assert_eq!(&*entry.data, &[0xFF_u8, 0xD8, 0xFF][..]);
        assert_eq!(entry.format, ImageFormat::Jpeg);
    }

    /// Media crosses the resolve boundary as a *handle*. Copying instead is
    /// invisible in the rendered output — every assertion above still holds —
    /// but it doubles the resident cost of every image, so nothing else can
    /// catch a regression here.
    #[test]
    fn resolve_shares_media_bytes_rather_than_copying_them() {
        use crate::model::ImageFormat;
        use std::sync::Arc;

        let bytes: Arc<[u8]> = Arc::from(&[0xFF_u8, 0xD8, 0xFF][..]);
        let mut doc = empty_doc();
        doc.media
            .insert(RelId::new("rId1"), (Arc::clone(&bytes), ImageFormat::Jpeg));

        let resolved = resolve(doc);
        assert!(
            Arc::ptr_eq(&resolved.media[&RelId::new("rId1")].data, &bytes),
            "resolve must pass the parser's allocation through, not a copy of it",
        );
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

        let resolved = resolve(doc);
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

        let resolved = resolve(doc);
        assert!(resolved.sections[0].headers.default.is_some());
    }
}
