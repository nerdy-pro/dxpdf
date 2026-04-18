//! Serde schema types for DrawingML (§20) — one submodule per primitive
//! family. Schemas are additive during Phase 5: they coexist with the
//! procedural parsers in `drawing/` and will progressively replace them
//! as higher-level container parsers are migrated in later sub-phases.

pub mod color;
pub mod fill;
pub mod stroke;
