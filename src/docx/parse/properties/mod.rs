//! OOXML property element schemas: pPr, rPr, tblPr, trPr, tcPr, sectPr.
//!
//! All parsing is serde-based; see [`schema`] for the per-element modules
//! and shared sub-schemas (border, shading, tabs, measure, etc.).

pub mod schema;
