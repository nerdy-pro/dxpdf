//! Serde schema types for OOXML property elements.
//!
//! Schema types carry the `Xml` suffix and are `pub(crate)` — never exported
//! from `docx::`. They mirror OOXML grammar; `From<_Xml> for ModelType`
//! conversions live alongside each schema.
//!
//! During Phase 2 these coexist with the legacy event-driven parsers in
//! `parse/properties/*.rs`. Phase 4 (body) completes the migration by
//! consuming these schemas directly and retiring the legacy parsers. Until
//! then, the module is dead code to cargo; the allow attribute below
//! suppresses the warning — every struct is reachable via `From` impls and
//! tested in-module.

#![allow(dead_code)]

pub mod border;
pub mod cnf_style;
pub mod fonts;
pub mod insets;
pub mod lang;
pub mod measure;
pub mod paragraph;
pub mod run;
pub mod section;
pub mod shading;
pub mod table;
pub mod tabs;
