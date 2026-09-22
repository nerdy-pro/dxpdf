//! Parser for footnotes.xml and endnotes.xml.

use std::collections::HashMap;

use serde::Deserialize;

use crate::docx::error::Result;
use crate::docx::model::{Block, NoteId};
use crate::docx::parse::body;
use crate::docx::parse::serde_xml::from_xml;

/// Parse footnotes.xml or endnotes.xml into a map of note ID → blocks. The
/// schema accepts either `<w:footnote>` or `<w:endnote>` children under either
/// root, so one function serves both parts.
pub fn parse_notes(data: &[u8]) -> Result<HashMap<NoteId, Vec<Block>>> {
    if data.is_empty() {
        return Ok(HashMap::new());
    }
    let file: NotesFileXml = from_xml(data)?;
    let mut out = HashMap::new();
    for note in file.entries {
        let Some(id) = note.id else { continue };
        // A fresh ctx per note: notes are independent stories, and an
        // unbalanced revision or comment marker in one must not stamp the
        // ones after it (issue #154).
        let mut ctx = body::ConvertCtx::new();
        let (blocks, _) = body::convert_container(note.content, &mut ctx);
        out.insert(NoteId::new(id), blocks);
    }
    Ok(out)
}

/// Matches both `<w:footnotes>` and `<w:endnotes>`. Their children are
/// `<w:footnote>` or `<w:endnote>` respectively — we accept either tag.
#[derive(Deserialize)]
struct NotesFileXml {
    #[serde(alias = "footnote", alias = "endnote", default)]
    entries: Vec<NoteXml>,
}

#[derive(Deserialize)]
struct NoteXml {
    #[serde(rename = "@id", default)]
    id: Option<i64>,
    #[serde(rename = "$value", default)]
    content: Vec<crate::docx::parse::body_schema::BlockChildXml>,
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#;

    #[test]
    fn empty_data_yields_empty_map() {
        assert!(parse_notes(b"").unwrap().is_empty());
    }

    #[test]
    fn parses_footnotes_keyed_by_id() {
        let xml = format!(
            r#"<w:footnotes {NS}>
                <w:footnote w:id="1"><w:p><w:r><w:t>first</w:t></w:r></w:p></w:footnote>
                <w:footnote w:id="2"><w:p><w:r><w:t>second</w:t></w:r></w:p></w:footnote>
            </w:footnotes>"#
        );
        let notes = parse_notes(xml.as_bytes()).expect("parse footnotes");
        assert_eq!(notes.len(), 2);
        assert!(notes.contains_key(&NoteId::new(1)));
        assert!(notes.contains_key(&NoteId::new(2)));
    }

    #[test]
    fn note_without_id_is_dropped() {
        let xml = format!(
            r#"<w:footnotes {NS}>
                <w:footnote><w:p/></w:footnote>
                <w:footnote w:id="5"><w:p/></w:footnote>
            </w:footnotes>"#
        );
        let notes = parse_notes(xml.as_bytes()).expect("parse");
        assert_eq!(notes.len(), 1, "the id-less note is dropped");
        assert!(notes.contains_key(&NoteId::new(5)));
    }

    #[test]
    fn endnote_alias_is_accepted() {
        let xml = format!(
            r#"<w:endnotes {NS}>
                <w:endnote w:id="3"><w:p><w:r><w:t>e</w:t></w:r></w:p></w:endnote>
            </w:endnotes>"#
        );
        let notes = parse_notes(xml.as_bytes()).expect("parse endnotes");
        assert_eq!(notes.len(), 1);
        assert!(notes.contains_key(&NoteId::new(3)));
    }
}
