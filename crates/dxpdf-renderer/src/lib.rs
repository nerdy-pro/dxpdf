//! PDF renderer for dxpdf — measure, layout, and paint pipeline.
//!
//! Takes a parsed `Document` from `dxpdf-docx` and produces PDF bytes.

pub mod dimension;
pub mod geometry;
pub mod layout;
pub mod resolve;
