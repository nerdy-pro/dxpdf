//! Parser for footnotes.xml and endnotes.xml.

use std::collections::HashMap;

use serde::Deserialize;

use crate::docx::error::Result;
use crate::docx::model::{Block, NoteId};
use crate::docx::parse::body;
use crate::docx::parse::body_schema::BlockContainerXml;
use crate::docx::parse::serde_xml::from_xml;

/// Parse footnotes.xml or endnotes.xml into a map of note ID → blocks.
///
/// `note_tag` is informational — the schema accepts either `<w:footnote>` or
/// `<w:endnote>` children under either root.
pub fn parse_notes(data: &[u8], _note_tag: &str) -> Result<HashMap<NoteId, Vec<Block>>> {
    if data.is_empty() {
        return Ok(HashMap::new());
    }
    // Two-pass: extract drawings/picts from the whole file once, then serde.
    let embeds = body::extract_embeds(data, b"")?;
    let file: NotesFileXml = from_xml(data)?;
    let mut ctx = body::ConvertCtx::new(embeds);
    let mut out = HashMap::new();
    for note in file.entries {
        let Some(id) = note.id else { continue };
        let container = BlockContainerXml {
            children: note.content,
        };
        let (blocks, _) = body::convert_container(container.children, &mut ctx);
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
