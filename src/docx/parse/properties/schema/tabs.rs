//! Tabs sub-schema (§17.3.1.38 w:tabs).

use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, Twips};
use crate::docx::model::TabStop;
use crate::docx::parse::primitives::st_enums::{StTabJc, StTabTlc};

/// `<w:tabs>` — a list of `<w:tab>` stops.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TabsXml {
    #[serde(rename = "tab", default)]
    pub stops: Vec<TabXml>,
}

/// `<w:tab w:pos="..." w:val="..." w:leader="..."/>` inside a `<w:tabs>`.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TabXml {
    #[serde(rename = "@pos")]
    pos: Dimension<Twips>,
    #[serde(rename = "@val", default = "default_val")]
    val: StTabJc,
    #[serde(rename = "@leader", default = "default_leader")]
    leader: StTabTlc,
}

fn default_val() -> StTabJc {
    StTabJc::Left
}

fn default_leader() -> StTabTlc {
    StTabTlc::None
}

impl From<TabXml> for TabStop {
    fn from(x: TabXml) -> Self {
        Self {
            position: x.pos,
            alignment: x.val.into(),
            leader: x.leader.into(),
        }
    }
}

impl From<TabsXml> for Vec<TabStop> {
    fn from(x: TabsXml) -> Self {
        // `<w:tab val="clear"/>` clears an inherited stop — we drop them.
        // The clear semantics live in the style cascade (§17.7.2); no tab
        // with `Clear` alignment should propagate to the rendering layer.
        x.stops
            .into_iter()
            .filter(|t| !matches!(t.val, StTabJc::Clear))
            .map(Into::into)
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{TabAlignment, TabLeader};

    fn parse(xml: &str) -> Vec<TabStop> {
        let t: TabsXml = quick_xml::de::from_str(xml).unwrap();
        t.into()
    }

    #[test]
    fn single_tab_with_leader() {
        let ts = parse(r#"<tabs><tab pos="1440" val="center" leader="dot"/></tabs>"#);
        assert_eq!(ts.len(), 1);
        assert_eq!(ts[0].position.raw(), 1440);
        assert_eq!(ts[0].alignment, TabAlignment::Center);
        assert_eq!(ts[0].leader, TabLeader::Dot);
    }

    #[test]
    fn tab_defaults_left_and_no_leader() {
        let ts = parse(r#"<tabs><tab pos="720"/></tabs>"#);
        assert_eq!(ts.len(), 1);
        assert_eq!(ts[0].alignment, TabAlignment::Left);
        assert_eq!(ts[0].leader, TabLeader::None);
    }

    #[test]
    fn clear_tabs_filtered_out() {
        let ts = parse(
            r#"<tabs>
                <tab pos="1440" val="left"/>
                <tab pos="2880" val="clear"/>
                <tab pos="4320" val="right"/>
            </tabs>"#,
        );
        assert_eq!(ts.len(), 2);
        assert_eq!(ts[0].position.raw(), 1440);
        assert_eq!(ts[1].position.raw(), 4320);
    }

    #[test]
    fn legacy_num_becomes_left() {
        let ts = parse(r#"<tabs><tab pos="720" val="num"/></tabs>"#);
        assert_eq!(ts[0].alignment, TabAlignment::Left);
    }

    #[test]
    fn empty_tabs() {
        let ts = parse(r#"<tabs/>"#);
        assert!(ts.is_empty());
    }
}
