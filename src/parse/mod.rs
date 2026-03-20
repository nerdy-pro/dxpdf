mod archive;
pub mod xml;

use std::collections::HashMap;
use std::path::Path;

use crate::error::Error;
use crate::model::{Block, Document, FormatHint, Inline, RelId};

fn resolve_image_data(
    rel_id: &RelId,
    data: &mut Vec<u8>,
    format_hint: &mut FormatHint,
    rels: &HashMap<String, String>,
    media: &HashMap<String, Vec<u8>>,
) {
    if let Some(target) = rels.get(rel_id.as_str()) {
        if let Some(bytes) = media.get(target) {
            *data = bytes.clone();
            *format_hint = FormatHint::new(
                Path::new(target)
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("")
                    .to_lowercase(),
            );
        }
    }
}

/// Parse a DOCX file (as raw bytes) into a `Document`.
pub fn parse(docx_bytes: &[u8]) -> Result<Document, Error> {
    let contents = archive::extract_docx_contents(docx_bytes)?;
    let mut document = xml::parse_document_xml(&contents.document_xml)?;
    if let Some(dts) = contents.default_tab_stop {
        document.default_tab_stop = dts;
    }
    resolve_images(
        &mut document,
        &contents.relationships,
        &contents.media_files,
    );
    Ok(document)
}

/// Walk the document tree and populate image data from the archive.
fn resolve_images(
    doc: &mut Document,
    rels: &HashMap<String, String>,
    media: &HashMap<String, Vec<u8>>,
) {
    for block in &mut doc.blocks {
        resolve_images_in_block(block, rels, media);
    }
}

fn resolve_images_in_block(
    block: &mut Block,
    rels: &HashMap<String, String>,
    media: &HashMap<String, Vec<u8>>,
) {
    match block {
        Block::Paragraph(p) => {
            for inline in &mut p.runs {
                if let Inline::Image(img) = inline {
                    resolve_image_data(&img.rel_id, &mut img.data, &mut img.format_hint, rels, media);
                }
            }
            for float in &mut p.floats {
                resolve_image_data(&float.rel_id, &mut float.data, &mut float.format_hint, rels, media);
            }
        }
        Block::Table(t) => {
            for row in &mut t.rows {
                for cell in &mut row.cells {
                    for block in &mut cell.blocks {
                        resolve_images_in_block(block, rels, media);
                    }
                }
            }
        }
    }
}
