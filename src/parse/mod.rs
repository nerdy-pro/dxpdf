mod archive;
pub mod xml;

use crate::error::Error;
use crate::model::Document;

/// Parse a DOCX file (as raw bytes) into a `Document`.
pub fn parse(docx_bytes: &[u8]) -> Result<Document, Error> {
    let xml = archive::extract_document_xml(docx_bytes)?;
    xml::parse_document_xml(&xml)
}
