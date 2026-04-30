//! Font subsetting — collect glyph usage, subset typefaces, replace.
//!
//! The pass runs between layout and paint. Driven by the `subset-fonts` Cargo
//! feature (default-on). See `docs/font-subsetting.md` for the full design.
//!
//! Design invariants:
//! - **Single source of truth** for typeface bytes ([`crate::render::fonts::FontRegistry`])
//!   and for usage tracking ([`collect::GlyphUsage`]). Each piece of state lives
//!   in exactly one place.
//! - **Glyph-id preservation.** `subsetter` retains the original glyph ids in
//!   the subsetted font, so downstream paint's `text_to_glyphs` and `cmap`
//!   lookups remain valid against the subsetted typeface — no re-shaping
//!   needed.
//! - **Spec touchpoints.** ECMA-376 §17.8 (DOCX font embedding,
//!   deobfuscation) is enforced upstream by the parser. ISO 32000-1 §9.6.4
//!   subset prefixes (`AAAAAA+`-style) are emitted by Skia's PDF backend at
//!   write time — we only feed it smaller bytes.

pub mod apply;
pub mod collect;
pub mod extract;
pub mod format;
#[cfg(feature = "subset-fonts")]
pub mod name_splice;

pub use apply::{apply, SubsetOutcome, SubsetReport};
pub use collect::{collect, Codepoint, CodepointUsage};
pub use extract::{extract, ExtractedSfnt, ExtractionError};
pub use format::{FontFormat, FormatError, SfntFlavor, WoffVersion};
