//! Document-level settings.

use crate::model::dimension::{Dimension, Twips};

use super::identifiers::RevisionSaveId;

#[derive(Clone, Debug, Default)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// The rsid of the original editing session that created this document.
    pub rsid_root: Option<RevisionSaveId>,
    /// All revision save IDs recorded in this document's history.
    pub rsids: Vec<RevisionSaveId>,
}
