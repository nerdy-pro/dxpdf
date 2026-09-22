//! Document-level settings.

use crate::model::dimension::{Dimension, Twips};

use super::identifiers::RevisionSaveId;

#[derive(Clone, Debug)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// The rsid of the original editing session that created this document.
    pub rsid_root: Option<RevisionSaveId>,
    /// All revision save IDs recorded in this document's history.
    pub rsids: Vec<RevisionSaveId>,
    /// `w:revisionView` (issue #154): whether the document's saved view shows
    /// insertion/deletion marks. `true` — the spec default when the element
    /// or its attributes are absent — renders unaccepted deletions struck
    /// through and insertions underlined; `false` renders the final text:
    /// deletions suppressed, insertions plain.
    pub show_ins_del_marks: bool,
    /// `w:revisionView` likewise for comment anchors and balloons.
    pub show_comment_marks: bool,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            // §17.15.1.25: when `w:defaultTabStop` is omitted the default tab
            // stop interval is 720 twips (0.5"). A derived `Default` would give
            // 0 here — wrong per spec, and it would silently collapse default
            // tabs for any consumer that reads this field.
            default_tab_stop: Dimension::new(720),
            even_and_odd_headers: false,
            rsid_root: None,
            rsids: Vec::new(),
            // `w:revisionView` absent means nothing was hidden.
            show_ins_del_marks: true,
            show_comment_marks: true,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_tab_stop_is_720_twips_per_spec() {
        // §17.15.1.25: an omitted `w:defaultTabStop` is 720 twips (0.5"), not
        // the 0 a derived `Default` would produce. Guards against a "simplify
        // to #[derive(Default)]" regression.
        assert_eq!(DocumentSettings::default().default_tab_stop.raw(), 720);
    }

    #[test]
    fn remaining_defaults() {
        let s = DocumentSettings::default();
        assert!(!s.even_and_odd_headers);
        assert!(s.rsid_root.is_none());
        assert!(s.rsids.is_empty());
    }
}
