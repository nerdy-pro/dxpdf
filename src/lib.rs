#![allow(
    clippy::too_many_arguments,
    clippy::collapsible_if,
    clippy::collapsible_match
)]

pub mod error;
pub mod model;
pub mod parse;
pub mod render;
pub mod units;

pub use error::Error;

/// Convert raw DOCX bytes into PDF bytes.
pub fn convert(docx_bytes: &[u8]) -> Result<Vec<u8>, Error> {
    let document = parse::parse(docx_bytes)?;
    convert_document(&document)
}

/// Convert a parsed `Document` into PDF bytes.
pub fn convert_document(document: &model::Document) -> Result<Vec<u8>, Error> {
    let config = render::layout::LayoutConfig::default();
    let pages = render::layout::layout(document, &config);
    render::painter::render_to_pdf(&pages)
}
