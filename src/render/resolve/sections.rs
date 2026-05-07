//! Section splitting — split body blocks at SectionBreak nodes,
//! resolve header/footer content via Document.headers/footers.

use std::collections::HashMap;

use crate::model::{Block, Document, RelId, SectionHeaderFooterRefs, SectionProperties};

use super::header_footer::HeaderFooterSet;

/// A resolved section with its blocks, properties, and the three
/// `default` / `first` / `even` slots populated for both header and
/// footer per ECMA-376 §17.10.5. Slot selection per page lives in
/// `crate::render::layout::header_footer::select_slot`.
#[derive(Clone, Debug)]
pub struct ResolvedSection {
    pub blocks: Vec<Block>,
    pub properties: SectionProperties,
    pub headers: HeaderFooterSet<Vec<Block>>,
    pub footers: HeaderFooterSet<Vec<Block>>,
}

/// Split document body into sections at `Block::SectionBreak` boundaries.
/// The final section uses `Document.final_section`.
/// Header/footer content is resolved from `Document.headers`/`Document.footers`.
///
/// §17.10.5 inheritance is applied **per slot independently**: if a
/// section omits its `first` reference, it inherits the previous
/// section's `first`, regardless of whether `default` is overridden.
/// This mirrors Word's behavior — each header type is its own
/// inheritance chain.
pub fn resolve_sections(doc: &Document) -> Vec<ResolvedSection> {
    let mut sections = Vec::new();
    let mut current_blocks = Vec::new();
    let mut prev_headers: HeaderFooterSet<Vec<Block>> = HeaderFooterSet::default();
    let mut prev_footers: HeaderFooterSet<Vec<Block>> = HeaderFooterSet::default();

    for block in &doc.body {
        match block {
            Block::SectionBreak(props) => {
                let headers = resolve_set(&props.header_refs, &doc.headers, &prev_headers);
                let footers = resolve_set(&props.footer_refs, &doc.footers, &prev_footers);
                prev_headers = headers.clone();
                prev_footers = footers.clone();
                sections.push(ResolvedSection {
                    blocks: std::mem::take(&mut current_blocks),
                    headers,
                    footers,
                    properties: *props.clone(),
                });
            }
            other => {
                current_blocks.push(other.clone());
            }
        }
    }

    let headers = resolve_set(&doc.final_section.header_refs, &doc.headers, &prev_headers);
    let footers = resolve_set(&doc.final_section.footer_refs, &doc.footers, &prev_footers);
    sections.push(ResolvedSection {
        blocks: current_blocks,
        headers,
        footers,
        properties: doc.final_section.clone(),
    });

    sections
}

/// Resolve all three slots of a section's header (or footer) refs:
/// look each up in the document's loaded parts, falling back to the
/// previous section's same slot when this section omits the reference.
fn resolve_set(
    refs: &SectionHeaderFooterRefs,
    parts: &HashMap<RelId, Vec<Block>>,
    prev: &HeaderFooterSet<Vec<Block>>,
) -> HeaderFooterSet<Vec<Block>> {
    let resolve_one = |id: Option<&RelId>, fallback: &Option<Vec<Block>>| -> Option<Vec<Block>> {
        id.and_then(|i| parts.get(i))
            .cloned()
            .or_else(|| fallback.clone())
    };
    HeaderFooterSet {
        default: resolve_one(refs.default.as_ref(), &prev.default),
        first: resolve_one(refs.first.as_ref(), &prev.first),
        even: resolve_one(refs.even.as_ref(), &prev.even),
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
        assert!(sections[0].headers.default.is_some());
        assert_eq!(sections[0].headers.default.as_ref().unwrap().len(), 1);
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
        assert!(sections[0].footers.default.is_some());
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
        assert!(sections[0].headers.default.is_none());
    }

    /// Helper: produce a single-paragraph block list to use as
    /// distinct header/footer content per ref.
    fn block(text: &str) -> Vec<Block> {
        vec![para(text)]
    }

    #[test]
    fn all_three_slots_resolved_when_all_refs_present() {
        let mut doc = empty_doc();
        let (rd, rf, re) = (RelId::new("rD"), RelId::new("rF"), RelId::new("rE"));
        doc.headers.insert(rd.clone(), block("default"));
        doc.headers.insert(rf.clone(), block("first"));
        doc.headers.insert(re.clone(), block("even"));
        doc.final_section = SectionProperties {
            header_refs: SectionHeaderFooterRefs {
                default: Some(rd),
                first: Some(rf),
                even: Some(re),
            },
            ..Default::default()
        };

        let s = &resolve_sections(&doc)[0];
        assert!(s.headers.default.is_some());
        assert!(s.headers.first.is_some());
        assert!(s.headers.even.is_some());
    }

    #[test]
    fn missing_first_and_even_slots_remain_none_with_no_prior_section() {
        // §17.10.5: a slot that is neither set on this section nor
        // inheritable from a previous one stays `None`.
        let mut doc = empty_doc();
        let rd = RelId::new("rD");
        doc.headers.insert(rd.clone(), block("default"));
        doc.final_section = SectionProperties {
            header_refs: SectionHeaderFooterRefs {
                default: Some(rd),
                ..Default::default()
            },
            ..Default::default()
        };

        let s = &resolve_sections(&doc)[0];
        assert!(s.headers.default.is_some());
        assert!(s.headers.first.is_none());
        assert!(s.headers.even.is_none());
    }

    #[test]
    fn first_slot_inherits_independently_when_only_default_is_overridden() {
        // Section 1: full set. Section 2: overrides only `default` —
        // its `first` and `even` must inherit Section 1's values.
        let mut doc = empty_doc();
        let (s1d, s1f, s1e, s2d) = (
            RelId::new("s1D"),
            RelId::new("s1F"),
            RelId::new("s1E"),
            RelId::new("s2D"),
        );
        doc.headers.insert(s1d.clone(), block("S1 default"));
        doc.headers.insert(s1f.clone(), block("S1 first"));
        doc.headers.insert(s1e.clone(), block("S1 even"));
        doc.headers.insert(s2d.clone(), block("S2 default"));

        let s1_break = SectionProperties {
            section_type: Some(SectionType::NextPage),
            header_refs: SectionHeaderFooterRefs {
                default: Some(s1d),
                first: Some(s1f.clone()),
                even: Some(s1e.clone()),
            },
            ..Default::default()
        };
        doc.final_section = SectionProperties {
            header_refs: SectionHeaderFooterRefs {
                default: Some(s2d),
                first: None,
                even: None,
            },
            ..Default::default()
        };
        doc.body = vec![
            para("section 1 body"),
            Block::SectionBreak(Box::new(s1_break)),
            para("section 2 body"),
        ];

        let sections = resolve_sections(&doc);
        let s2 = &sections[1];
        assert_eq!(
            s2.headers.default.as_ref().map(|b| b.len()),
            Some(1),
            "section 2 default is its own override",
        );
        assert!(
            s2.headers.first.is_some(),
            "section 2 must inherit `first` from section 1",
        );
        assert!(
            s2.headers.even.is_some(),
            "section 2 must inherit `even` from section 1",
        );
    }

    #[test]
    fn slots_inherit_independently_across_three_sections() {
        // Section 1 sets `even` only. Section 2 sets `first` only.
        // Section 3 sets nothing — it should inherit `first` from S2,
        // `even` from S1, and have no `default`.
        let mut doc = empty_doc();
        let (s1e, s2f) = (RelId::new("s1E"), RelId::new("s2F"));
        doc.headers.insert(s1e.clone(), block("S1 even"));
        doc.headers.insert(s2f.clone(), block("S2 first"));

        let s1_break = SectionProperties {
            section_type: Some(SectionType::NextPage),
            header_refs: SectionHeaderFooterRefs {
                default: None,
                first: None,
                even: Some(s1e),
            },
            ..Default::default()
        };
        let s2_break = SectionProperties {
            section_type: Some(SectionType::NextPage),
            header_refs: SectionHeaderFooterRefs {
                default: None,
                first: Some(s2f),
                even: None,
            },
            ..Default::default()
        };
        doc.final_section = SectionProperties::default();
        doc.body = vec![
            para("S1"),
            Block::SectionBreak(Box::new(s1_break)),
            para("S2"),
            Block::SectionBreak(Box::new(s2_break)),
            para("S3"),
        ];

        let sections = resolve_sections(&doc);
        assert_eq!(sections.len(), 3);
        let s3 = &sections[2];
        assert!(s3.headers.default.is_none());
        assert!(s3.headers.first.is_some(), "S3 inherits `first` from S2",);
        assert!(
            s3.headers.even.is_some(),
            "S3 inherits `even` transitively from S1 (S2 didn't override it)",
        );
    }

    #[test]
    fn footer_slots_inherit_with_the_same_per_slot_rule() {
        let mut doc = empty_doc();
        let (s1d, s1f) = (RelId::new("s1D"), RelId::new("s1F"));
        doc.footers.insert(s1d.clone(), block("default"));
        doc.footers.insert(s1f.clone(), block("first"));

        let s1_break = SectionProperties {
            section_type: Some(SectionType::NextPage),
            footer_refs: SectionHeaderFooterRefs {
                default: Some(s1d),
                first: Some(s1f),
                even: None,
            },
            ..Default::default()
        };
        doc.final_section = SectionProperties::default();
        doc.body = vec![
            para("S1"),
            Block::SectionBreak(Box::new(s1_break)),
            para("S2"),
        ];

        let s2 = &resolve_sections(&doc)[1];
        assert!(s2.footers.default.is_some(), "S2 inherits default footer");
        assert!(s2.footers.first.is_some(), "S2 inherits first footer");
        assert!(s2.footers.even.is_none());
    }
}
