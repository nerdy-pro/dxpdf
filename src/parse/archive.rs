use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;

/// Everything extracted from the DOCX ZIP archive needed for conversion.
pub struct DocxContents {
    pub document_xml: String,
    /// Relationship ID -> target path (relative to `word/`).
    pub relationships: HashMap<String, String>,
    /// Path (relative to `word/`, e.g. "media/image1.png") -> raw bytes.
    pub media_files: HashMap<String, Vec<u8>>,
}

/// Extract document XML, relationships, and media files from a DOCX archive.
pub fn extract_docx_contents(docx_bytes: &[u8]) -> Result<DocxContents, Error> {
    let cursor = std::io::Cursor::new(docx_bytes);
    let mut archive = zip::ZipArchive::new(cursor)?;

    // 1. Extract word/document.xml
    let document_xml = {
        let mut file = archive
            .by_name("word/document.xml")
            .map_err(|_| Error::MissingEntry("word/document.xml".into()))?;
        let mut xml = String::new();
        file.read_to_string(&mut xml)?;
        xml
    };

    // 2. Extract relationships from word/_rels/document.xml.rels
    let relationships = match archive.by_name("word/_rels/document.xml.rels") {
        Ok(mut file) => {
            let mut rels_xml = String::new();
            file.read_to_string(&mut rels_xml)?;
            parse_relationships(&rels_xml)?
        }
        Err(_) => HashMap::new(),
    };

    // 3. Extract all media files (word/media/*)
    let mut media_files = HashMap::new();
    let media_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let file = archive.by_index(i).ok()?;
            let name = file.name().to_string();
            if name.starts_with("word/media/") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for name in media_names {
        let mut file = archive.by_name(&name)?;
        let mut data = Vec::new();
        file.read_to_end(&mut data)?;
        // Store relative to word/ (e.g., "media/image1.png")
        let rel_path = name.strip_prefix("word/").unwrap_or(&name).to_string();
        media_files.insert(rel_path, data);
    }

    Ok(DocxContents {
        document_xml,
        relationships,
        media_files,
    })
}

/// Parse `word/_rels/document.xml.rels` into a map of relationship ID -> target path.
fn parse_relationships(xml: &str) -> Result<HashMap<String, String>, Error> {
    let mut reader = Reader::from_str(xml);
    let mut rels = HashMap::new();

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            Event::Empty(ref e) | Event::Start(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"Relationship" {
                    let mut id = None;
                    let mut target = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        match key {
                            b"Id" => {
                                id = Some(
                                    String::from_utf8_lossy(&attr.value).into_owned(),
                                );
                            }
                            b"Target" => {
                                target = Some(
                                    String::from_utf8_lossy(&attr.value).into_owned(),
                                );
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        rels.insert(id, target);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(rels)
}

/// Strip namespace prefix from an element/attribute name.
fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|&b| b == b':')
        .map(|i| &name[i + 1..])
        .unwrap_or(name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_rels_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpeg"/>
        </Relationships>"#;
        let rels = parse_relationships(xml).unwrap();
        assert_eq!(rels.get("rId1").unwrap(), "media/image1.png");
        assert_eq!(rels.get("rId2").unwrap(), "media/image2.jpeg");
    }

    #[test]
    fn parse_empty_rels() {
        let xml = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let rels = parse_relationships(xml).unwrap();
        assert!(rels.is_empty());
    }
}
