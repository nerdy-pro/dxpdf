//! Parser for `word/settings.xml`.

use crate::model::Dup;
use serde::Deserialize;

use crate::docx::dimension::{Dimension, Twips};
use crate::docx::error::Result;
use crate::docx::model::{DocumentSettings, RevisionSaveId};
use crate::docx::parse::primitives::units::deserialize_nonnegative_dimension;
use crate::docx::parse::primitives::OnOff;
use crate::docx::parse::serde_xml::from_xml;

/// Parse `word/settings.xml`. Entry point: deserializes into an intermediate
/// schema, then maps to the model type.
pub fn parse_settings(data: &[u8]) -> Result<DocumentSettings> {
    from_xml::<SettingsXml>(data).map(Into::into)
}

#[derive(Deserialize, Default)]
struct SettingsXml {
    #[serde(rename = "defaultTabStop", default)]
    default_tab_stop: Vec<DimensionVal<Twips>>,
    #[serde(rename = "evenAndOddHeaders", default)]
    even_and_odd_headers: Vec<OnOff>,
    /// §17.15.1.10 — a §17.17.4 boolean, so the bare element is on.
    #[serde(rename = "autoHyphenation", default)]
    auto_hyphenation: Vec<OnOff>,
    /// §17.15.1.22 `ST_DecimalNumber`.
    #[serde(rename = "consecutiveHyphenLimit", default)]
    consecutive_hyphen_limit: Vec<DecimalVal>,
    /// §17.15.1.53 `ST_TwipsMeasure`, so it shares `DimensionVal` with
    /// `w:defaultTabStop` and accepts the same universal-measure spellings.
    #[serde(rename = "hyphenationZone", default)]
    hyphenation_zone: Vec<DimensionVal<Twips>>,
    /// §17.15.1.27 — a §17.17.4 boolean.
    #[serde(rename = "doNotHyphenateCaps", default)]
    do_not_hyphenate_caps: Vec<OnOff>,
    #[serde(default)]
    rsids: Vec<RsidsXml>,
}

#[derive(Deserialize, Default)]
struct RsidsXml {
    #[serde(rename = "rsidRoot", default)]
    rsid_root: Vec<StringVal>,
    #[serde(rename = "rsid", default)]
    rsids: Vec<StringVal>,
}

#[derive(Deserialize)]
#[serde(bound(deserialize = "U: crate::docx::dimension::Unit"))]
struct DimensionVal<U: crate::docx::dimension::Unit> {
    #[serde(
        rename = "@val",
        deserialize_with = "deserialize_nonnegative_dimension"
    )]
    val: Dimension<U>,
}

#[derive(Deserialize)]
struct StringVal {
    #[serde(rename = "@val")]
    val: String,
}

/// `CT_DecimalNumber` — `xsd:integer`, so signed on the wire even where the
/// clause only makes sense for a count. The one consumer here
/// (`w:consecutiveHyphenLimit`) folds a negative into "no limit" at the seam
/// rather than failing the whole settings part over it.
#[derive(Deserialize)]
struct DecimalVal {
    #[serde(rename = "@val")]
    val: i64,
}

impl From<SettingsXml> for DocumentSettings {
    fn from(x: SettingsXml) -> Self {
        let mut s = DocumentSettings::default();
        if let Some(t) = Dup::from(x.default_tab_stop).into_value() {
            s.default_tab_stop = t.val;
        }
        if let Some(OnOff(on)) = Dup::from(x.even_and_odd_headers).into_value() {
            s.even_and_odd_headers = on;
        }
        // §17.15.1.10 / .53 / .22 / .27 (issue #158). Each starts from
        // `HyphenationSettings::default()`, which holds the spec's omitted
        // values, and is overwritten only by an element that is present.
        if let Some(OnOff(on)) = Dup::from(x.auto_hyphenation).into_value() {
            s.hyphenation.enabled = on;
        }
        if let Some(z) = Dup::from(x.hyphenation_zone).into_value() {
            s.hyphenation.zone = z.val;
        }
        if let Some(l) = Dup::from(x.consecutive_hyphen_limit).into_value() {
            // §17.15.1.22: 0 and omitted both mean "no limit"; a negative
            // count has no meaning the clause gives it, so it is read as the
            // same absence rather than rejected.
            s.hyphenation.consecutive_limit = u32::try_from(l.val).unwrap_or(0);
        }
        if let Some(OnOff(on)) = Dup::from(x.do_not_hyphenate_caps).into_value() {
            s.hyphenation.no_caps = on;
        }
        if let Some(r) = Dup::from(x.rsids).into_value() {
            if let Some(root) = Dup::from(r.rsid_root).into_value() {
                s.rsid_root = RevisionSaveId::from_hex(&root.val);
            }
            s.rsids = r
                .rsids
                .into_iter()
                .filter_map(|v| RevisionSaveId::from_hex(&v.val))
                .collect();
        }
        s
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::HyphenationSettings;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn settings(body: &str) -> DocumentSettings {
        let xml = format!(r#"<?xml version="1.0"?><w:settings xmlns:w="{W}">{body}</w:settings>"#);
        parse_settings(xml.as_bytes()).expect("settings parse")
    }

    /// §17.15.1.10 is a §17.17.4 boolean: the bare element is on, an explicit
    /// `w:val="0"` is off, and omission is off ("If this element is omitted,
    /// then hyphenation shall not automatically be performed").
    #[test]
    fn auto_hyphenation_parses_as_a_17_17_4_toggle() {
        assert!(settings("<w:autoHyphenation/>").hyphenation.enabled);
        assert!(
            settings(r#"<w:autoHyphenation w:val="true"/>"#)
                .hyphenation
                .enabled
        );
        assert!(
            !settings(r#"<w:autoHyphenation w:val="0"/>"#)
                .hyphenation
                .enabled
        );
        assert!(
            !settings(r#"<w:autoHyphenation w:val="false"/>"#)
                .hyphenation
                .enabled
        );
        assert!(!settings("").hyphenation.enabled);
    }

    /// §17.15.1.53: "If this element is omitted, then the hyphenation zone
    /// shall be 360 twentieths of a point" — and `0` is a value, not an
    /// absence (`test-files/document_outline.docx` declares exactly that).
    #[test]
    fn hyphenation_zone_defaults_to_360_twips_when_omitted() {
        assert_eq!(settings("<w:autoHyphenation/>").hyphenation.zone.raw(), 360);
        assert_eq!(
            settings(r#"<w:hyphenationZone w:val="720"/>"#)
                .hyphenation
                .zone
                .raw(),
            720
        );
        assert_eq!(
            settings(r#"<w:hyphenationZone w:val="0"/>"#)
                .hyphenation
                .zone
                .raw(),
            0
        );
    }

    /// §17.15.1.53 is `ST_TwipsMeasure`, so the §22.9.2.14 universal-measure
    /// spellings parse exactly as they do for `w:defaultTabStop`.
    #[test]
    fn hyphenation_zone_accepts_universal_measures() {
        // Neither value is the 360-twip default, so a parser that dropped the
        // element could not pass by accident.
        assert_eq!(
            settings(r#"<w:hyphenationZone w:val="36pt"/>"#)
                .hyphenation
                .zone
                .raw(),
            720
        );
        assert_eq!(
            settings(r#"<w:hyphenationZone w:val="0.5in"/>"#)
                .hyphenation
                .zone
                .raw(),
            720
        );
    }

    /// §17.15.1.22: "If this element is omitted, then there shall be no
    /// limit"; and a declared `0` means the same, so both arrive as 0.
    #[test]
    fn consecutive_hyphen_limit_zero_and_omitted_both_mean_no_limit() {
        assert_eq!(settings("").hyphenation.consecutive_limit, 0);
        assert_eq!(
            settings(r#"<w:consecutiveHyphenLimit w:val="0"/>"#)
                .hyphenation
                .consecutive_limit,
            0
        );
        assert_eq!(
            settings(r#"<w:consecutiveHyphenLimit w:val="2"/>"#)
                .hyphenation
                .consecutive_limit,
            2
        );
    }

    /// §17.15.1.27: "If omitted, words in CAPITAL LETTERS shall be
    /// hyphenated"; the bare element turns that off, like any §17.17.4 toggle.
    #[test]
    fn do_not_hyphenate_caps_defaults_off() {
        assert!(!settings("").hyphenation.no_caps);
        assert!(settings("<w:doNotHyphenateCaps/>").hyphenation.no_caps);
        assert!(
            !settings(r#"<w:doNotHyphenateCaps w:val="0"/>"#)
                .hyphenation
                .no_caps
        );
    }

    /// The settings part Word writes interleaves these four with dozens of
    /// elements this schema does not name; none of them may derail the parse,
    /// and the four are read wherever they fall in the sequence.
    #[test]
    fn the_four_elements_are_read_from_a_realistic_settings_part() {
        let s = settings(
            r#"<w:zoom w:percent="100"/><w:proofState w:spelling="clean"/>
               <w:defaultTabStop w:val="708"/><w:autoHyphenation/>
               <w:consecutiveHyphenLimit w:val="3"/><w:hyphenationZone w:val="425"/>
               <w:doNotHyphenateCaps/><w:characterSpacingControl w:val="doNotCompress"/>
               <w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat>"#,
        );
        assert_eq!(s.default_tab_stop.raw(), 708);
        assert_eq!(
            s.hyphenation,
            HyphenationSettings {
                enabled: true,
                zone: Dimension::new(425),
                consecutive_limit: 3,
                no_caps: true,
            }
        );
    }
}
