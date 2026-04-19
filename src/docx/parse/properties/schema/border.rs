//! Border sub-schemas (§17.3.4 pBdr, §17.4.39 tblBorders, §17.4.66 tcBorders).
//!
//! `BorderXml` matches a single `<w:top>`/`<w:bottom>`/etc. element. The
//! container structs (paragraph / table / table-cell borders) share the same
//! inner `BorderXml` but differ in which sides are allowed. Each container
//! accepts both modern (`start`/`end`) and legacy (`left`/`right`) side
//! names per OOXML bidi handling.

use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, EighthPoints, Points};
use crate::docx::model::{Border, Color, ParagraphBorders, TableBorders, TableCellBorders};
use crate::docx::parse::primitives::st_enums::StBorderType;
use crate::docx::parse::primitives::HexColor;

/// A single `<w:top w:val="..." w:sz="..." w:space="..." w:color="..."/>` etc.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct BorderXml {
    #[serde(rename = "@val")]
    val: StBorderType,
    #[serde(rename = "@sz", default)]
    sz: Option<Dimension<EighthPoints>>,
    #[serde(rename = "@space", default)]
    space: Option<Dimension<Points>>,
    #[serde(rename = "@color", default)]
    color: Option<HexColor>,
}

impl From<BorderXml> for Border {
    fn from(x: BorderXml) -> Self {
        Self {
            style: x.val.into(),
            width: x.sz.unwrap_or_default(),
            space: x.space.unwrap_or_default(),
            color: x.color.map_or(Color::Auto, Into::into),
        }
    }
}

/// `<w:pBdr>` — five sides plus `between`.
#[derive(Clone, Copy, Debug, Default, Deserialize)]
pub(crate) struct ParagraphBordersXml {
    #[serde(default)]
    top: Option<BorderXml>,
    #[serde(default)]
    bottom: Option<BorderXml>,
    #[serde(default, alias = "start")]
    left: Option<BorderXml>,
    #[serde(default, alias = "end")]
    right: Option<BorderXml>,
    #[serde(default)]
    between: Option<BorderXml>,
}

impl From<ParagraphBordersXml> for ParagraphBorders {
    fn from(x: ParagraphBordersXml) -> Self {
        Self {
            top: x.top.map(Into::into),
            bottom: x.bottom.map(Into::into),
            left: x.left.map(Into::into),
            right: x.right.map(Into::into),
            between: x.between.map(Into::into),
        }
    }
}

/// `<w:tblBorders>` — six sides (adds `insideH`, `insideV`).
#[derive(Clone, Copy, Debug, Default, Deserialize)]
pub(crate) struct TableBordersXml {
    #[serde(default)]
    top: Option<BorderXml>,
    #[serde(default)]
    bottom: Option<BorderXml>,
    #[serde(default, alias = "start")]
    left: Option<BorderXml>,
    #[serde(default, alias = "end")]
    right: Option<BorderXml>,
    #[serde(rename = "insideH", default)]
    inside_h: Option<BorderXml>,
    #[serde(rename = "insideV", default)]
    inside_v: Option<BorderXml>,
}

impl From<TableBordersXml> for TableBorders {
    fn from(x: TableBordersXml) -> Self {
        Self {
            top: x.top.map(Into::into),
            bottom: x.bottom.map(Into::into),
            left: x.left.map(Into::into),
            right: x.right.map(Into::into),
            inside_h: x.inside_h.map(Into::into),
            inside_v: x.inside_v.map(Into::into),
        }
    }
}

/// `<w:tcBorders>` — eight sides (adds diagonal `tl2br`, `tr2bl`).
#[derive(Clone, Copy, Debug, Default, Deserialize)]
pub(crate) struct TableCellBordersXml {
    #[serde(default)]
    top: Option<BorderXml>,
    #[serde(default)]
    bottom: Option<BorderXml>,
    #[serde(default, alias = "start")]
    left: Option<BorderXml>,
    #[serde(default, alias = "end")]
    right: Option<BorderXml>,
    #[serde(rename = "insideH", default)]
    inside_h: Option<BorderXml>,
    #[serde(rename = "insideV", default)]
    inside_v: Option<BorderXml>,
    #[serde(rename = "tl2br", default)]
    tl2br: Option<BorderXml>,
    #[serde(rename = "tr2bl", default)]
    tr2bl: Option<BorderXml>,
}

impl From<TableCellBordersXml> for TableCellBorders {
    fn from(x: TableCellBordersXml) -> Self {
        Self {
            top: x.top.map(Into::into),
            bottom: x.bottom.map(Into::into),
            left: x.left.map(Into::into),
            right: x.right.map(Into::into),
            inside_h: x.inside_h.map(Into::into),
            inside_v: x.inside_v.map(Into::into),
            tl2br: x.tl2br.map(Into::into),
            tr2bl: x.tr2bl.map(Into::into),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn single_border_with_all_attrs() {
        let xml = r#"<top val="single" sz="4" space="0" color="FF0000"/>"#;
        let b: BorderXml = quick_xml::de::from_str(xml).unwrap();
        let m: Border = b.into();
        assert_eq!(m.style, crate::docx::model::BorderStyle::Single);
        assert_eq!(m.width.raw(), 4);
        assert_eq!(m.space.raw(), 0);
        assert_eq!(m.color, Color::Rgb(0xFF0000));
    }

    #[test]
    fn border_with_auto_color() {
        let xml = r#"<top val="single" color="auto"/>"#;
        let b: BorderXml = quick_xml::de::from_str(xml).unwrap();
        let m: Border = b.into();
        assert_eq!(m.color, Color::Auto);
    }

    #[test]
    fn paragraph_borders_aliases_start_left() {
        let xml = r#"<pBdr>
            <top val="single"/>
            <start val="double"/>
            <end val="thick"/>
        </pBdr>"#;
        let px: ParagraphBordersXml = quick_xml::de::from_str(xml).unwrap();
        let p: ParagraphBorders = px.into();
        assert!(p.top.is_some());
        assert_eq!(
            p.left.unwrap().style,
            crate::docx::model::BorderStyle::Double
        );
        assert_eq!(
            p.right.unwrap().style,
            crate::docx::model::BorderStyle::Thick
        );
    }

    #[test]
    fn table_borders_inside_pair() {
        let xml = r#"<tblBorders>
            <insideH val="single"/>
            <insideV val="dashed"/>
        </tblBorders>"#;
        let tx: TableBordersXml = quick_xml::de::from_str(xml).unwrap();
        let t: TableBorders = tx.into();
        assert!(t.inside_h.is_some());
        assert!(t.inside_v.is_some());
        assert!(t.top.is_none());
    }

    #[test]
    fn cell_borders_diagonals() {
        let xml = r#"<tcBorders>
            <tl2br val="single"/>
            <tr2bl val="dotted"/>
        </tcBorders>"#;
        let cx: TableCellBordersXml = quick_xml::de::from_str(xml).unwrap();
        let c: TableCellBorders = cx.into();
        assert!(c.tl2br.is_some());
        assert!(c.tr2bl.is_some());
    }
}
