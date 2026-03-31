//! DOCX document model — fully resolved algebraic data types.
//!
//! This crate contains the type-safe dimension system, geometry primitives,
//! and the complete DOCX document model. No parsing logic, no external dependencies.

pub mod dimension;
pub mod geometry;
pub mod types;
pub use types::*;
