use std::io::Read;

use crate::error::Error;

/// Extract the `word/document.xml` content from a DOCX (ZIP) archive.
pub fn extract_document_xml(docx_bytes: &[u8]) -> Result<String, Error> {
    let cursor = std::io::Cursor::new(docx_bytes);
    let mut archive = zip::ZipArchive::new(cursor)?;
    let mut file = archive
        .by_name("word/document.xml")
        .map_err(|_| Error::MissingEntry("word/document.xml".into()))?;
    let mut xml = String::new();
    file.read_to_string(&mut xml)?;
    Ok(xml)
}
