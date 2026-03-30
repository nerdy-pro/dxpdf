//! OOXML field instruction parser and evaluator.
//!
//! Parses raw field instruction strings (from `w:instrText` or `w:fldSimple/@instr`)
//! into a typed AST, and evaluates them given a document context.
//!
//! # Examples
//!
//! ```
//! use dxpdf_field::{parse, evaluate, FieldContext};
//!
//! // Parse a PAGE field
//! let instr = parse(" PAGE ").unwrap();
//!
//! // Evaluate with context
//! let mut ctx = FieldContext::default();
//! ctx.page_number = Some(7);
//! let result = evaluate(&instr, &ctx);
//! ```

pub mod ast;
pub mod context;
pub mod error;
pub mod eval;
pub mod format;
pub(crate) mod parse;
pub(crate) mod switches;

pub use ast::{CommonSwitches, ComparisonOp, FieldInstruction};
pub use context::FieldContext;
pub use error::FieldParseError;
pub use eval::{evaluate, FieldValue};

/// Parse a field instruction string into a typed `FieldInstruction`.
///
/// The input is the raw instruction text as found in `w:instrText` elements
/// or the `instr` attribute of `w:fldSimple`, e.g. `" PAGE "`, `" TOC \o \"1-3\" \h "`.
///
/// Returns `Err(FieldParseError)` for malformed syntax (unterminated strings,
/// missing required arguments, invalid values). Unknown but syntactically valid
/// field types parse into `FieldInstruction::Unknown`.
pub fn parse(input: &str) -> Result<FieldInstruction, FieldParseError> {
    parse::parse(input)
}
