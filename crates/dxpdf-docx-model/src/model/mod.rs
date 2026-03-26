//! Complete DOCX document model — all types are fully resolved ADTs.
//! No unparsed strings, no style indirection, no invalid states.

mod color;
mod content;
mod document;
mod drawing;
mod formatting;
mod identifiers;
mod numbering;
mod paragraph;
mod run_properties;
mod section;
mod settings;
mod styles;
mod table;
mod theme;
mod vml;

pub use color::*;
pub use content::*;
pub use document::*;
pub use drawing::*;
pub use formatting::*;
pub use identifiers::*;
pub use numbering::*;
pub use paragraph::*;
pub use run_properties::*;
pub use section::*;
pub use settings::*;
pub use styles::*;
pub use table::*;
pub use theme::*;
pub use vml::*;
