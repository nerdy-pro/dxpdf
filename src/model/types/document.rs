//! Top-level Document struct.

use std::collections::HashMap;
use std::sync::Arc;

use super::content::{Block, Comment};
use super::drawing::ImageFormat;
use super::identifiers::{CommentId, NoteId, RelId};
use super::numbering::NumberingDefinitions;
use super::section::SectionProperties;
use super::settings::DocumentSettings;
use super::styles::StyleSheet;
use super::theme::Theme;

/// The fully parsed and resolved DOCX document.
#[derive(Clone, Debug)]
pub struct Document {
    pub settings: DocumentSettings,
    pub theme: Option<Theme>,
    /// Style definitions from `word/styles.xml`, with `basedOn` references intact.
    pub styles: StyleSheet,
    /// Numbering definitions from `word/numbering.xml`.
    pub numbering: NumberingDefinitions,
    pub body: Vec<Block>,
    /// Final section properties (from the last `w:sectPr` in `w:body`).
    pub final_section: SectionProperties,
    /// Header content keyed by relationship ID.
    pub headers: HashMap<RelId, Vec<Block>>,
    /// Footer content keyed by relationship ID.
    pub footers: HashMap<RelId, Vec<Block>>,
    pub footnotes: HashMap<NoteId, Vec<Block>>,
    pub endnotes: HashMap<NoteId, Vec<Block>>,
    /// Comment content keyed by comment ID (`word/comments.xml`, §11.3.1).
    pub comments: HashMap<CommentId, Comment>,
    /// Embedded media (images) — raw bytes and detected format, keyed by
    /// relationship ID.
    ///
    /// The bytes are shared rather than owned so that the resolve layer, which
    /// consumes this document, hands the painter a *handle* instead of copying
    /// every image a second time — on an image-heavy document that copy is the
    /// single largest allocation in the pipeline. `Arc` rather than `Rc` keeps
    /// `Document` `Send`, so a caller may still parse on one thread and render
    /// on another; the refcount is touched once per image placement, which is
    /// nowhere near hot enough for the atomic to matter.
    pub media: HashMap<RelId, (Arc<[u8]>, ImageFormat)>,
    /// §17.8.3: embedded fonts — de-obfuscated font data.
    pub embedded_fonts: Vec<EmbeddedFont>,
}

/// §17.8.3: an embedded font extracted from the DOCX package.
/// The font data is de-obfuscated (XOR with fontKey per §17.8.3.3)
/// and ready to be registered with a font manager.
#[derive(Clone, Debug)]
pub struct EmbeddedFont {
    /// Font family name from `w:font/@w:name`.
    pub family: String,
    /// Which variant of the font this is.
    pub variant: EmbeddedFontVariant,
    /// De-obfuscated TrueType/OpenType font data.
    pub data: Vec<u8>,
}

/// §17.8.3.3: which style variant an embedded font represents.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum EmbeddedFontVariant {
    /// §17.8.3.3 w:embedRegular
    Regular,
    /// §17.8.3.3 w:embedBold
    Bold,
    /// §17.8.3.3 w:embedItalic
    Italic,
    /// §17.8.3.3 w:embedBoldItalic
    BoldItalic,
}
