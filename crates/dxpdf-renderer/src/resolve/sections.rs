//! Section splitting — split body blocks at SectionBreak nodes,
//! resolve header/footer content via Document.headers/footers.

use dxpdf_docx_model::model::{
    Block, Document, RelId, SectionProperties,
};

/// A resolved section with its blocks and properties.
#[derive(Clone, Debug)]
pub struct ResolvedSection {
    pub blocks: Vec<Block>,
    pub properties: SectionProperties,
    pub header: Option<Vec<Block>>,
    pub footer: Option<Vec<Block>>,
}

/// Split document body into sections at `Block::SectionBreak` boundaries.
/// The final section uses `Document.final_section`.
/// Header/footer content is resolved from `Document.headers`/`Document.footers`.
pub fn resolve_sections(doc: &Document) -> Vec<ResolvedSection> {
    let mut sections = Vec::new();
    let mut current_blocks = Vec::new();
    // §17.10.5: sections without explicit header/footer refs inherit from
    // the previous section.
    let mut prev_header: Option<Vec<Block>> = None;
    let mut prev_footer: Option<Vec<Block>> = None;

    for block in &doc.body {
        match block {
            Block::SectionBreak(props) => {
                let header = resolve_header(doc, &props.header_refs.default)
                    .or_else(|| prev_header.clone());
                let footer = resolve_footer(doc, &props.footer_refs.default)
                    .or_else(|| prev_footer.clone());
                prev_header.clone_from(&header);
                prev_footer.clone_from(&footer);
                sections.push(ResolvedSection {
                    blocks: std::mem::take(&mut current_blocks),
                    header,
                    footer,
                    properties: *props.clone(),
                });
            }
            other => {
                current_blocks.push(other.clone());
            }
        }
    }

    // Final section uses Document.final_section.
    let header = resolve_header(doc, &doc.final_section.header_refs.default)
        .or(prev_header);
    let footer = resolve_footer(doc, &doc.final_section.footer_refs.default)
        .or(prev_footer);
    sections.push(ResolvedSection {
        blocks: current_blocks,
        header,
        footer,
        properties: doc.final_section.clone(),
    });

    sections
}

/// Look up header content by RelId.
fn resolve_header(doc: &Document, rel_id: &Option<RelId>) -> Option<Vec<Block>> {
    rel_id
        .as_ref()
        .and_then(|id| doc.headers.get(id))
        .cloned()
}

/// Look up footer content by RelId.
fn resolve_footer(doc: &Document, rel_id: &Option<RelId>) -> Option<Vec<Block>> {
    rel_id
        .as_ref()
        .and_then(|id| doc.footers.get(id))
        .cloned()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;
    use dxpdf_docx_model::model::*;

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
                properties: RunProperties::default(),
                content: vec![RunElement::Text(text.to_string())],
                rsids: RevisionIds::default(),
            }))],
            rsids: ParagraphRevisionIds::default(),
        }))
    }

    #[test]
    fn empty_body_produces_one_section() {
        let doc = empty_doc();
        let sections = resolve_sections(&doc);
        assert_eq!(sections.len(), 1);
        assert!(sections[0].blocks.is_empty());
    }

    #[test]
    fn no_section_breaks_all_blocks_in_final_section() {
        let mut doc = empty_doc();
        doc.body = vec![para("a"), para("b"), para("c")];
        let sections = resolve_sections(&doc);

        assert_eq!(sections.len(), 1);
        assert_eq!(sections[0].blocks.len(), 3);
    }

    #[test]
    fn section_break_splits_into_two_sections() {
        let mut doc = empty_doc();
        let break_props = SectionProperties {
            section_type: Some(SectionType::NextPage),
            ..Default::default()
        };
        doc.body = vec![
            para("before"),
            Block::SectionBreak(Box::new(break_props)),
            para("after"),
        ];

        let sections = resolve_sections(&doc);
        assert_eq!(sections.len(), 2);
        assert_eq!(sections[0].blocks.len(), 1, "first section has 'before'");
        assert_eq!(sections[1].blocks.len(), 1, "second section has 'after'");
        assert_eq!(
            sections[0].properties.section_type,
            Some(SectionType::NextPage)
        );
    }

    #[test]
    fn header_resolved_from_document_headers() {
        let mut doc = empty_doc();
        let header_id = RelId::new("rId1");
        doc.headers
            .insert(header_id.clone(), vec![para("header text")]);
        doc.final_section = SectionProperties {
            header_refs: SectionHeaderFooterRefs {
                default: Some(header_id),
                ..Default::default()
            },
            ..Default::default()
        };
        doc.body = vec![para("body")];

        let sections = resolve_sections(&doc);
        assert_eq!(sections.len(), 1);
        assert!(sections[0].header.is_some());
        assert_eq!(sections[0].header.as_ref().unwrap().len(), 1);
    }

    #[test]
    fn footer_resolved_from_document_footers() {
        let mut doc = empty_doc();
        let footer_id = RelId::new("rId2");
        doc.footers
            .insert(footer_id.clone(), vec![para("footer text")]);
        doc.final_section = SectionProperties {
            footer_refs: SectionHeaderFooterRefs {
                default: Some(footer_id),
                ..Default::default()
            },
            ..Default::default()
        };
        doc.body = vec![para("body")];

        let sections = resolve_sections(&doc);
        assert!(sections[0].footer.is_some());
    }

    #[test]
    fn missing_header_ref_produces_none() {
        let mut doc = empty_doc();
        doc.final_section = SectionProperties {
            header_refs: SectionHeaderFooterRefs {
                default: Some(RelId::new("nonexistent")),
                ..Default::default()
            },
            ..Default::default()
        };
        doc.body = vec![para("body")];

        let sections = resolve_sections(&doc);
        assert!(sections[0].header.is_none());
    }
}
