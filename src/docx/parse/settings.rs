//! Parser for `word/settings.xml`.

use crate::model::Dup;
use serde::Deserialize;

use crate::docx::dimension::{Dimension, Twips};
use crate::docx::error::Result;
use crate::docx::model::{DocumentSettings, RevisionSaveId};
use crate::docx::parse::primitives::units::deserialize_nonnegative_dimension;
use crate::docx::parse::primitives::{AttrBool, OnOff};
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
    #[serde(default)]
    rsids: Vec<RsidsXml>,
    #[serde(rename = "revisionView", default)]
    revision_view: Vec<RevisionViewXml>,
}

/// `w:revisionView` — which markup the document was saved displaying, and so
/// what its own print view shows (issue #154). Every attribute defaults to
/// "shown" when absent, per the spec; `@markup` is the master switch over the
/// per-kind toggles.
#[derive(Deserialize, Default)]
struct RevisionViewXml {
    #[serde(rename = "@markup")]
    markup: Option<AttrBool>,
    #[serde(rename = "@insDel")]
    ins_del: Option<AttrBool>,
    #[serde(rename = "@comments")]
    comments: Option<AttrBool>,
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

impl From<SettingsXml> for DocumentSettings {
    fn from(x: SettingsXml) -> Self {
        let mut s = DocumentSettings::default();
        if let Some(t) = Dup::from(x.default_tab_stop).into_value() {
            s.default_tab_stop = t.val;
        }
        if let Some(OnOff(on)) = Dup::from(x.even_and_odd_headers).into_value() {
            s.even_and_odd_headers = on;
        }
        if let Some(rv) = Dup::from(x.revision_view).into_value() {
            let on = |o: &Option<AttrBool>| o.map(|AttrBool(b)| b).unwrap_or(true);
            let markup = on(&rv.markup);
            s.show_ins_del_marks = markup && on(&rv.ins_del);
            s.show_comment_marks = markup && on(&rv.comments);
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

    fn parse(xml: &str) -> DocumentSettings {
        parse_settings(xml.as_bytes()).expect("settings must parse")
    }

    /// `w:revisionView` absent: nothing was hidden — the spec default.
    #[test]
    fn absent_revision_view_shows_all_marks() {
        let s = parse(r#"<w:settings xmlns:w="x"/>"#);
        assert!(s.show_ins_del_marks);
        assert!(s.show_comment_marks);
    }

    /// `@markup="0"` is the master switch: every markup kind is hidden,
    /// whatever the per-kind attributes say.
    #[test]
    fn markup_off_hides_everything() {
        let s = parse(
            r#"<w:settings xmlns:w="x"><w:revisionView w:markup="0" w:insDel="1"/></w:settings>"#,
        );
        assert!(!s.show_ins_del_marks);
        assert!(!s.show_comment_marks);
    }

    /// Per-kind toggles are independent under an (implicitly) on master.
    #[test]
    fn ins_del_off_leaves_comments_shown() {
        let s = parse(r#"<w:settings xmlns:w="x"><w:revisionView w:insDel="0"/></w:settings>"#);
        assert!(!s.show_ins_del_marks);
        assert!(s.show_comment_marks);
    }
}
