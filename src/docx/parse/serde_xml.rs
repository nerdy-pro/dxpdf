//! Shared helpers for serde-driven OOXML parsers.

use serde::de::DeserializeOwned;

use crate::docx::error::Result;

/// Deserialize an OOXML part into a schema type, mapping quick-xml's error
/// into the crate's `ParseError`.
pub fn from_xml<T: DeserializeOwned>(data: &[u8]) -> Result<T> {
    Ok(quick_xml::de::from_reader(data)?)
}
