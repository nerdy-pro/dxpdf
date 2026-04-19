//! Language sub-schema (§17.3.2.20 w:lang) — tri-mode language tag.

use serde::Deserialize;

use crate::docx::model::Lang;

#[derive(Clone, Debug, Deserialize)]
pub(crate) struct LangXml {
    #[serde(rename = "@val", default)]
    val: Option<String>,
    #[serde(rename = "@eastAsia", default)]
    east_asia: Option<String>,
    #[serde(rename = "@bidi", default)]
    bidi: Option<String>,
}

impl From<LangXml> for Lang {
    fn from(x: LangXml) -> Self {
        Self {
            val: x.val,
            east_asia: x.east_asia,
            bidi: x.bidi,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> Lang {
        let x: LangXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn all_three_modes() {
        let l = parse(r#"<lang val="en-US" eastAsia="ja-JP" bidi="ar-SA"/>"#);
        assert_eq!(l.val.as_deref(), Some("en-US"));
        assert_eq!(l.east_asia.as_deref(), Some("ja-JP"));
        assert_eq!(l.bidi.as_deref(), Some("ar-SA"));
    }

    #[test]
    fn missing_attrs_become_none() {
        let l = parse(r#"<lang/>"#);
        assert!(l.val.is_none() && l.east_asia.is_none() && l.bidi.is_none());
    }
}
