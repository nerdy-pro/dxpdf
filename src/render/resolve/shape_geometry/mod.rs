//! Shape geometry resolution — evaluates OOXML geometry definitions
//! (`prstGeom`, `custGeom`) into drawable paths.
//!
//! Tier 0 lands only the guide-expression evaluator (§20.1.9.11). Preset
//! geometry generators and the path-verb resolver for `custGeom` land in
//! Phase 5.

pub mod guides;
