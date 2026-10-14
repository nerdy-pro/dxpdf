//! Document-level settings.

use crate::model::dimension::{Dimension, Twips};

use super::identifiers::RevisionSaveId;

#[derive(Clone, Debug)]
pub struct DocumentSettings {
    /// Default tab stop interval (OOXML default: 720 twips = 0.5 inch).
    pub default_tab_stop: Dimension<Twips>,
    /// Whether even/odd headers/footers are enabled.
    pub even_and_odd_headers: bool,
    /// §17.15.1.10 and its three companions — see [`HyphenationSettings`].
    pub hyphenation: HyphenationSettings,
    /// The rsid of the original editing session that created this document.
    pub rsid_root: Option<RevisionSaveId>,
    /// All revision save IDs recorded in this document's history.
    pub rsids: Vec<RevisionSaveId>,
}

/// The four `w:settings` children that govern automatic hyphenation
/// (issue #158), as the document declared them.
///
/// Document-wide only: §17.3.1.34 `w:suppressAutoHyphens` is the per-paragraph
/// opt-out and lives on `ParagraphProperties`, where the §17.7.2 cascade
/// already carries it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct HyphenationSettings {
    /// §17.15.1.10 `w:autoHyphenation`: whether the consumer hyphenates at
    /// all. "If this element is omitted, then hyphenation shall not
    /// automatically be performed."
    pub enabled: bool,
    /// §17.15.1.53 `w:hyphenationZone`: "the amount of whitespace which can
    /// be left at the end of a line … before hyphenation should be attempted
    /// on the next word". Omitted → 360 twips (0.25"). `0` is a legitimate
    /// value — hyphenate on any gap — not an absent one.
    pub zone: Dimension<Twips>,
    /// §17.15.1.22 `w:consecutiveHyphenLimit`: "the maximum number of
    /// consecutive lines of text that can end with a hyphen". Omitted *or* `0`
    /// → no limit; the two spellings are folded here so no consumer has to
    /// know they are the same.
    pub consecutive_limit: u32,
    /// §17.15.1.27 `w:doNotHyphenateCaps`: when set, "words comprised of all
    /// capital letters" are left whole. "If omitted, words in CAPITAL LETTERS
    /// shall be hyphenated."
    pub no_caps: bool,
}

impl Default for HyphenationSettings {
    fn default() -> Self {
        Self {
            // §17.15.1.10: omitted means off.
            enabled: false,
            // §17.15.1.53: "If this element is omitted, then the hyphenation
            // zone shall be 360 twentieths of a point". A derived `Default`
            // would give 0 — a real value with the opposite meaning.
            zone: Dimension::new(360),
            // §17.15.1.22: omitted means no limit, which this field spells 0.
            consecutive_limit: 0,
            // §17.15.1.27: omitted means capitals are hyphenated too.
            no_caps: false,
        }
    }
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
            hyphenation: HyphenationSettings::default(),
            rsid_root: None,
            rsids: Vec::new(),
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

    /// §17.15.1.10 / .53 / .22 / .27: the defaults when `w:settings` names
    /// none of the four. The zone is the one a derived `Default` would get
    /// wrong, and the one with the opposite meaning at 0.
    #[test]
    fn hyphenation_defaults_follow_the_four_clauses() {
        let h = HyphenationSettings::default();
        assert!(!h.enabled, "§17.15.1.10: omitted means off");
        assert_eq!(h.zone.raw(), 360, "§17.15.1.53: omitted means 360 twips");
        assert_eq!(
            h.consecutive_limit, 0,
            "§17.15.1.22: omitted means no limit"
        );
        assert!(!h.no_caps, "§17.15.1.27: omitted means capitals hyphenate");
        assert_eq!(DocumentSettings::default().hyphenation, h);
    }
}
