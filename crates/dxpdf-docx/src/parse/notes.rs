//! Parser for footnotes.xml and endnotes.xml.

use std::collections::HashMap;

use log::warn;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::{Block, NoteId};
use crate::xml;

use super::body;

/// Parse footnotes.xml or endnotes.xml into a map of note ID → blocks.
pub fn parse_notes(data: &[u8], note_tag: &str) -> Result<HashMap<NoteId, Vec<Block>>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut notes = HashMap::new();

    let note_tag_bytes = note_tag.as_bytes();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == note_tag_bytes => {
                let id = xml::optional_attr(e, b"id")?.and_then(|s| s.parse::<i64>().ok());
                let note_type = xml::optional_attr(e, b"type")?;

                // Skip separator and continuation separator notes
                if let Some(ref t) = note_type {
                    if t == "separator" || t == "continuationSeparator" {
                        xml::skip_to_end(&mut reader, &mut buf, note_tag_bytes)?;
                        continue;
                    }
                }

                if let Some(note_id) = id {
                    let blocks = parse_note_content(&mut reader, &mut buf, note_tag_bytes)?;
                    notes.insert(NoteId(note_id), blocks);
                } else {
                    xml::skip_to_end(&mut reader, &mut buf, note_tag_bytes)?;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(notes)
}

fn parse_note_content(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<Vec<Block>> {
    let mut blocks = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"p" => {
                        let (para, sect) = body::parse_paragraph_public(e, reader, buf)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" => {
                        let table = body::parse_table_public(reader, buf)?;
                        blocks.push(Block::Table(Box::new(table)));
                    }
                    _ => {
                        warn!(
                            "note: unsupported block element <{}>",
                            String::from_utf8_lossy(&local)
                        );
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(blocks)
}
