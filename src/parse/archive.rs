use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;

/// Document-wide default run properties from `word/styles.xml`.
pub struct DocDefaults {
    /// Default font size in half-points.
    pub font_size: Option<u32>,
    /// Default font family.
    pub font_family: Option<String>,
}

/// Everything extracted from the DOCX ZIP archive needed for conversion.
pub struct DocxContents {
    pub document_xml: String,
    /// Relationship ID -> target path (relative to `word/`).
    pub relationships: HashMap<String, String>,
    /// Path (relative to `word/`, e.g. "media/image1.png") -> raw bytes.
    pub media_files: HashMap<String, Vec<u8>>,
    /// Default tab stop interval in twips (from `word/settings.xml`).
    pub default_tab_stop: Option<u32>,
    /// Document-wide default run properties from `word/styles.xml`.
    pub doc_defaults: Option<DocDefaults>,
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

    // 4. Extract defaultTabStop from word/settings.xml
    let default_tab_stop = match archive.by_name("word/settings.xml") {
        Ok(mut file) => {
            let mut settings_xml = String::new();
            file.read_to_string(&mut settings_xml)?;
            parse_default_tab_stop(&settings_xml)
        }
        Err(_) => None,
    };

    // 5. Extract document defaults from word/styles.xml
    let doc_defaults = match archive.by_name("word/styles.xml") {
        Ok(mut file) => {
            let mut styles_xml = String::new();
            file.read_to_string(&mut styles_xml)?;
            parse_doc_defaults(&styles_xml)
        }
        Err(_) => None,
    };

    Ok(DocxContents {
        document_xml,
        relationships,
        media_files,
        default_tab_stop,
        doc_defaults,
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

/// Parse `word/styles.xml` to extract document default font size and family.
fn parse_doc_defaults(xml: &str) -> Option<DocDefaults> {
    let mut reader = Reader::from_str(xml);
    let mut in_doc_defaults = false;
    let mut in_rpr_default = false;
    let mut font_size = None;
    let mut font_family = None;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"docDefaults" => in_doc_defaults = true,
                    b"rPrDefault" if in_doc_defaults => in_rpr_default = true,
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"docDefaults" => break, // Done once we leave docDefaults
                    b"rPrDefault" => in_rpr_default = false,
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if in_rpr_default => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"sz" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == b"val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                font_size = val.parse().ok();
                            }
                        }
                    }
                    b"rFonts" => {
                        // Try ascii, then hAnsi
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"ascii" || key == b"hAnsi" {
                                let val =
                                    String::from_utf8_lossy(&attr.value).into_owned();
                                // Skip theme references like "minorHAnsi"
                                if font_family.is_none() && !val.is_empty() {
                                    font_family = Some(val);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    if font_size.is_some() || font_family.is_some() {
        Some(DocDefaults {
            font_size,
            font_family,
        })
    } else {
        None
    }
}

/// Parse `word/settings.xml` to find `w:defaultTabStop` value.
fn parse_default_tab_stop(xml: &str) -> Option<u32> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"defaultTabStop" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            return val.parse().ok();
                        }
                    }
                }
            }
            Err(_) => break,
            _ => {}
        }
    }
    None
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
