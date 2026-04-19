//! VML (Vector Markup Language) parsing — fully serde-based.
//!
//! Pure attribute-level sub-grammars live in sibling submodules
//! (`color`, `formulas`, `path_commands`, `style`) as `&str → value`
//! functions. XML structure lives in [`schema`].

mod color;
mod formulas;
mod path_commands;
pub mod schema;
mod style;
