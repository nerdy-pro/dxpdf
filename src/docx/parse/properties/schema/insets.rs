//! Edge insets (§17.4.42 tcMar, §17.4.44 tblCellMar) — four-sided twips
//! padding shared by table-cell margins and table default cell margins.
//!
//! Each side is `<w:top w:w="N" w:type="dxa"/>` etc. Only `dxa` (twips) is
//! meaningful for cell padding; other `@type` values are ignored here.

use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, Twips};
use crate::docx::model::geometry::EdgeInsets;

#[derive(Clone, Copy, Debug, Default, Deserialize)]
pub(crate) struct EdgeInsetsTwipsXml {
    #[serde(default)]
    top: Option<SideXml>,
    #[serde(default)]
    bottom: Option<SideXml>,
    #[serde(default, alias = "start")]
    left: Option<SideXml>,
    #[serde(default, alias = "end")]
    right: Option<SideXml>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
struct SideXml {
    #[serde(rename = "@w", default)]
    w: Option<Dimension<Twips>>,
}

impl From<EdgeInsetsTwipsXml> for EdgeInsets<Twips> {
    fn from(x: EdgeInsetsTwipsXml) -> Self {
        Self::new(
            x.top.and_then(|s| s.w).unwrap_or_default(),
            x.right.and_then(|s| s.w).unwrap_or_default(),
            x.bottom.and_then(|s| s.w).unwrap_or_default(),
            x.left.and_then(|s| s.w).unwrap_or_default(),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> EdgeInsets<Twips> {
        let x: EdgeInsetsTwipsXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn all_four_sides_captured() {
        let e = parse(
            r#"<tcMar>
                <top w="100"/>
                <bottom w="200"/>
                <left w="50"/>
                <right w="75"/>
            </tcMar>"#,
        );
        assert_eq!(e.top.raw(), 100);
        assert_eq!(e.bottom.raw(), 200);
        assert_eq!(e.left.raw(), 50);
        assert_eq!(e.right.raw(), 75);
    }

    #[test]
    fn start_and_end_alias_left_right() {
        let e = parse(
            r#"<tcMar>
                <start w="80"/>
                <end w="120"/>
            </tcMar>"#,
        );
        assert_eq!(e.left.raw(), 80);
        assert_eq!(e.right.raw(), 120);
    }

    #[test]
    fn missing_sides_default_to_zero() {
        let e = parse(r#"<tcMar><top w="100"/></tcMar>"#);
        assert_eq!(e.top.raw(), 100);
        assert_eq!(e.right.raw(), 0);
        assert_eq!(e.bottom.raw(), 0);
        assert_eq!(e.left.raw(), 0);
    }
}
