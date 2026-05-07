//! Table property schemas: `<w:tblPr>`, `<w:trPr>`, `<w:tcPr>`.
//!
//! `TblPrXml::split` returns `(TableProperties, Option<StyleId>)` — the style
//! id travels separately for cascade reasons, matching the legacy parser's
//! signature.

use serde::{Deserialize, Deserializer};

use crate::docx::model::dimension::Twips;
use crate::docx::model::geometry::EdgeInsets;
use crate::docx::model::{
    Alignment, CnfStyle, StyleId, TableCellProperties, TableLook, TablePositioning,
    TableProperties, TableRowHeight, TableRowProperties, VerticalMerge,
};
use crate::docx::parse::primitives::st_enums::{
    StAnchor, StHeightRule, StJc, StTblLayoutType, StTblOverlap, StTextDirection, StVerticalJc,
    StXAlign, StYAlign,
};
use crate::docx::parse::primitives::OnOff;

use super::border::{TableBordersXml, TableCellBordersXml};
use super::cnf_style::CnfStyleXml;
use super::insets::EdgeInsetsTwipsXml;
use super::measure::TableMeasureXml;
use super::shading::ShdXml;

// ── tblPr ───────────────────────────────────────────────────────────────

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TblPrXml {
    #[serde(rename = "tblStyle", default)]
    tbl_style: Option<ValString>,
    #[serde(rename = "tblBorders", default)]
    tbl_borders: Option<TableBordersXml>,
    #[serde(rename = "tblCellMar", default)]
    tbl_cell_mar: Option<EdgeInsetsTwipsXml>,
    #[serde(rename = "jc", default)]
    jc: Option<ValAttr<StJc>>,
    #[serde(rename = "tblW", default)]
    tbl_w: Option<TableMeasureXml>,
    #[serde(rename = "tblLayout", default)]
    tbl_layout: Option<TblLayoutXml>,
    #[serde(rename = "tblInd", default)]
    tbl_ind: Option<TableMeasureXml>,
    #[serde(rename = "tblCellSpacing", default)]
    tbl_cell_spacing: Option<TableMeasureXml>,
    #[serde(rename = "tblLook", default)]
    tbl_look: Option<TblLookXml>,
    #[serde(rename = "tblStyleRowBandSize", default)]
    tbl_style_row_band_size: Option<ValAttr<u32>>,
    #[serde(rename = "tblStyleColBandSize", default)]
    tbl_style_col_band_size: Option<ValAttr<u32>>,
    #[serde(rename = "tblpPr", default)]
    tblp_pr: Option<TblpPrXml>,
    #[serde(rename = "tblOverlap", default)]
    tbl_overlap: Option<ValAttr<StTblOverlap>>,
}

/// `<w:tblLayout w:type="fixed"/>` — note `@type` (not `@val`).
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TblLayoutXml {
    #[serde(rename = "@type")]
    ty: StTblLayoutType,
}

/// `<w:tblLook>` — supports both the modern explicit attributes (firstRow,
/// lastRow, ...) and the legacy hex bitfield on `@val`. Per
/// [MS-OI29500] §2.1.1583, when both are present the explicit attribute
/// wins per-flag; otherwise the bitfield supplies the value.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TblLookXml {
    #[serde(rename = "@val", default)]
    val: Option<TblLookHex>,
    #[serde(rename = "@firstRow", default)]
    first_row: Option<AttrBool>,
    #[serde(rename = "@lastRow", default)]
    last_row: Option<AttrBool>,
    #[serde(rename = "@firstColumn", default)]
    first_column: Option<AttrBool>,
    #[serde(rename = "@lastColumn", default)]
    last_column: Option<AttrBool>,
    #[serde(rename = "@noHBand", default)]
    no_h_band: Option<AttrBool>,
    #[serde(rename = "@noVBand", default)]
    no_v_band: Option<AttrBool>,
}

/// Word's legacy `<w:tblLook val>` hex bitfield, per [MS-OI29500] §2.1.1583.
///
/// Bit positions:
/// | Mask     | Flag        |
/// |----------|-------------|
/// | `0x0020` | firstRow    |
/// | `0x0040` | lastRow     |
/// | `0x0080` | firstColumn |
/// | `0x0100` | lastColumn  |
/// | `0x0200` | noHBand     |
/// | `0x0400` | noVBand     |
///
/// Other bits are reserved/ignored. The Word default `04A0` =
/// firstRow + firstColumn + noVBand.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct TblLookHex(u16);

impl TblLookHex {
    const FIRST_ROW: u16 = 0x0020;
    const LAST_ROW: u16 = 0x0040;
    const FIRST_COLUMN: u16 = 0x0080;
    const LAST_COLUMN: u16 = 0x0100;
    const NO_H_BAND: u16 = 0x0200;
    const NO_V_BAND: u16 = 0x0400;

    fn first_row(self) -> bool {
        self.0 & Self::FIRST_ROW != 0
    }
    fn last_row(self) -> bool {
        self.0 & Self::LAST_ROW != 0
    }
    fn first_column(self) -> bool {
        self.0 & Self::FIRST_COLUMN != 0
    }
    fn last_column(self) -> bool {
        self.0 & Self::LAST_COLUMN != 0
    }
    fn no_h_band(self) -> bool {
        self.0 & Self::NO_H_BAND != 0
    }
    fn no_v_band(self) -> bool {
        self.0 & Self::NO_V_BAND != 0
    }
}

impl<'de> Deserialize<'de> for TblLookHex {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        u16::from_str_radix(s.trim_start_matches("0x"), 16)
            .map(TblLookHex)
            .map_err(serde::de::Error::custom)
    }
}

/// `<w:tblpPr>` — floating table positioning.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TblpPrXml {
    #[serde(rename = "@leftFromText", default)]
    left_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@rightFromText", default)]
    right_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@topFromText", default)]
    top_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@bottomFromText", default)]
    bottom_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@vertAnchor", default)]
    vert_anchor: Option<StAnchor>,
    #[serde(rename = "@horzAnchor", default)]
    horz_anchor: Option<StAnchor>,
    #[serde(rename = "@tblpXSpec", default)]
    x_spec: Option<StXAlign>,
    #[serde(rename = "@tblpYSpec", default)]
    y_spec: Option<StYAlign>,
    #[serde(rename = "@tblpX", default)]
    x: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@tblpY", default)]
    y: Option<crate::docx::model::dimension::Dimension<Twips>>,
}

/// §17.4.61 `<w:tblPrEx>` — table-level property exceptions scoped to
/// a single row. Per the spec it accepts the same vocabulary as
/// `<w:tblPr>` minus `tblStyle` and `tblpPr`. We model only the slice
/// the layout currently honors (table borders); other fields can be
/// added incrementally.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TblPrExXml {
    #[serde(rename = "tblBorders", default)]
    tbl_borders: Option<TableBordersXml>,
}

impl From<TblPrExXml> for crate::docx::model::TableRowPropertyExceptions {
    fn from(x: TblPrExXml) -> Self {
        Self {
            borders: x.tbl_borders.map(Into::into),
        }
    }
}

impl TblPrXml {
    pub(crate) fn split(self) -> (TableProperties, Option<StyleId>) {
        let style_id = self.tbl_style.map(|v| StyleId::new(v.val));
        let props = TableProperties {
            style_id: style_id.clone(),
            alignment: self.jc.map(|v| Alignment::from(v.val)),
            width: self.tbl_w.map(Into::into),
            layout: self
                .tbl_layout
                .map(|v| crate::docx::model::TableLayout::from(v.ty)),
            indent: self.tbl_ind.map(Into::into),
            borders: self.tbl_borders.map(Into::into),
            cell_margins: self.tbl_cell_mar.map(Into::into),
            cell_spacing: self.tbl_cell_spacing.map(Into::into),
            look: self.tbl_look.map(Into::into),
            style_row_band_size: self.tbl_style_row_band_size.map(|v| v.val),
            style_col_band_size: self.tbl_style_col_band_size.map(|v| v.val),
            positioning: self.tblp_pr.map(Into::into),
            overlap: self
                .tbl_overlap
                .map(|v| crate::docx::model::TableOverlap::from(v.val)),
        };
        (props, style_id)
    }
}

impl From<TblLookXml> for TableLook {
    fn from(x: TblLookXml) -> Self {
        // Per [MS-OI29500] §2.1.1583: explicit attribute wins per-flag,
        // legacy `val` supplies the fallback bit when the attribute is absent.
        let from_val = |bit: fn(TblLookHex) -> bool| x.val.map(bit);
        Self {
            first_row: x
                .first_row
                .map(|b| b.0)
                .or_else(|| from_val(TblLookHex::first_row)),
            last_row: x
                .last_row
                .map(|b| b.0)
                .or_else(|| from_val(TblLookHex::last_row)),
            first_column: x
                .first_column
                .map(|b| b.0)
                .or_else(|| from_val(TblLookHex::first_column)),
            last_column: x
                .last_column
                .map(|b| b.0)
                .or_else(|| from_val(TblLookHex::last_column)),
            no_h_band: x
                .no_h_band
                .map(|b| b.0)
                .or_else(|| from_val(TblLookHex::no_h_band)),
            no_v_band: x
                .no_v_band
                .map(|b| b.0)
                .or_else(|| from_val(TblLookHex::no_v_band)),
        }
    }
}

impl From<TblpPrXml> for TablePositioning {
    fn from(x: TblpPrXml) -> Self {
        Self {
            left_from_text: x.left_from_text,
            right_from_text: x.right_from_text,
            top_from_text: x.top_from_text,
            bottom_from_text: x.bottom_from_text,
            vert_anchor: x.vert_anchor.map(Into::into),
            horz_anchor: x.horz_anchor.map(Into::into),
            x_align: x.x_spec.map(Into::into),
            y_align: x.y_spec.map(Into::into),
            x: x.x,
            y: x.y,
        }
    }
}

// ── trPr ───────────────────────────────────────────────────────────────

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TrPrXml {
    #[serde(rename = "trHeight", default)]
    tr_height: Option<TrHeightXml>,
    #[serde(rename = "tblHeader", default)]
    tbl_header: Option<OnOff>,
    #[serde(rename = "cantSplit", default)]
    cant_split: Option<OnOff>,
    #[serde(rename = "jc", default)]
    jc: Option<ValAttr<StJc>>,
    #[serde(rename = "cnfStyle", default)]
    cnf_style: Option<CnfStyleXml>,
    #[serde(rename = "gridBefore", default)]
    grid_before: Option<ValAttr<u32>>,
    #[serde(rename = "wBefore", default)]
    w_before: Option<TableMeasureXml>,
    #[serde(rename = "gridAfter", default)]
    grid_after: Option<ValAttr<u32>>,
    #[serde(rename = "wAfter", default)]
    w_after: Option<TableMeasureXml>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TrHeightXml {
    #[serde(rename = "@val", default)]
    val: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@hRule", default)]
    rule: Option<StHeightRule>,
}

impl From<TrHeightXml> for TableRowHeight {
    fn from(x: TrHeightXml) -> Self {
        Self {
            value: x.val.unwrap_or_default(),
            rule: x
                .rule
                .map(Into::into)
                .unwrap_or(crate::docx::model::HeightRule::Auto),
        }
    }
}

impl From<TrPrXml> for TableRowProperties {
    fn from(x: TrPrXml) -> Self {
        Self {
            height: x.tr_height.map(Into::into),
            is_header: x.tbl_header.map(|OnOff(b)| b),
            cant_split: x.cant_split.map(|OnOff(b)| b),
            justification: x.jc.map(|v| Alignment::from(v.val)),
            cnf_style: x.cnf_style.map(CnfStyle::from),
            grid_before: x.grid_before.map(|v| v.val).unwrap_or(0),
            w_before: x.w_before.map(Into::into),
            grid_after: x.grid_after.map(|v| v.val).unwrap_or(0),
            w_after: x.w_after.map(Into::into),
        }
    }
}

// ── tcPr ───────────────────────────────────────────────────────────────

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TcPrXml {
    #[serde(rename = "tcBorders", default)]
    tc_borders: Option<TableCellBordersXml>,
    #[serde(rename = "tcMar", default)]
    tc_mar: Option<EdgeInsetsTwipsXml>,
    #[serde(rename = "tcW", default)]
    tc_w: Option<TableMeasureXml>,
    #[serde(rename = "shd", default)]
    shd: Option<ShdXml>,
    #[serde(rename = "vAlign", default)]
    v_align: Option<ValAttr<StVerticalJc>>,
    #[serde(rename = "vMerge", default)]
    v_merge: Option<VMergeXml>,
    #[serde(rename = "gridSpan", default)]
    grid_span: Option<ValAttr<u32>>,
    #[serde(rename = "textDirection", default)]
    text_direction: Option<ValAttr<StTextDirection>>,
    #[serde(rename = "noWrap", default)]
    no_wrap: Option<OnOff>,
    #[serde(rename = "cnfStyle", default)]
    cnf_style: Option<CnfStyleXml>,
}

/// `<w:vMerge/>` — absent `@val` means "continue"; `@val="restart"` starts a
/// new vertical merge group; `@val="continue"` is explicit continue.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct VMergeXml {
    #[serde(rename = "@val", default)]
    val: Option<VMergeKind>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum VMergeKind {
    Restart,
    Continue,
}

impl From<VMergeXml> for VerticalMerge {
    fn from(x: VMergeXml) -> Self {
        match x.val {
            Some(VMergeKind::Restart) => Self::Restart,
            Some(VMergeKind::Continue) | None => Self::Continue,
        }
    }
}

impl From<TcPrXml> for TableCellProperties {
    fn from(x: TcPrXml) -> Self {
        Self {
            width: x.tc_w.map(Into::into),
            borders: x.tc_borders.map(Into::into),
            shading: x.shd.map(Into::into),
            margins: x.tc_mar.map(EdgeInsets::from),
            vertical_align: x
                .v_align
                .map(|v| crate::docx::model::CellVerticalAlign::from(v.val)),
            vertical_merge: x.v_merge.map(Into::into),
            grid_span: x.grid_span.map(|v| v.val),
            text_direction: x
                .text_direction
                .map(|v| crate::docx::model::TextDirection::from(v.val)),
            no_wrap: x.no_wrap.map(|OnOff(b)| b),
            cnf_style: x.cnf_style.map(CnfStyle::from),
        }
    }
}

// ── shared helpers ──────────────────────────────────────────────────────

#[derive(Clone, Debug, Deserialize)]
struct ValString {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(bound(deserialize = "T: serde::Deserialize<'de>"))]
pub(crate) struct ValAttr<T> {
    #[serde(rename = "@val")]
    val: T,
}

use crate::docx::parse::primitives::AttrBool;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{
        BorderStyle, CellVerticalAlign, HeightRule, TableLayout, TableMeasure, TableOverlap,
        TextDirection,
    };

    // ── tblPr ──

    fn parse_tbl_pr(xml: &str) -> (TableProperties, Option<StyleId>) {
        let x: TblPrXml = quick_xml::de::from_str(xml).unwrap();
        x.split()
    }

    #[test]
    fn tbl_pr_style_and_width() {
        let (tp, sid) = parse_tbl_pr(
            r#"<tblPr><tblStyle val="TableGrid"/><tblW w="5000" type="pct"/></tblPr>"#,
        );
        assert_eq!(
            sid.map(|s| s.as_str().to_string()),
            Some("TableGrid".into())
        );
        match tp.width.unwrap() {
            TableMeasure::Pct(d) => assert_eq!(d.raw(), 5000),
            other => panic!("expected Pct, got {other:?}"),
        }
    }

    #[test]
    fn tbl_pr_layout_and_alignment() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><jc val="center"/><tblLayout type="fixed"/></tblPr>"#);
        assert_eq!(tp.layout, Some(TableLayout::Fixed));
        assert_eq!(tp.alignment, Some(Alignment::Center));
    }

    #[test]
    fn tbl_pr_borders_and_margins() {
        let (tp, _) = parse_tbl_pr(
            r#"<tblPr>
                <tblBorders><top val="single"/><left val="double"/></tblBorders>
                <tblCellMar><top w="100"/><left w="80"/></tblCellMar>
            </tblPr>"#,
        );
        let b = tp.borders.unwrap();
        assert_eq!(b.top.unwrap().style, BorderStyle::Single);
        assert_eq!(b.left.unwrap().style, BorderStyle::Double);
        assert_eq!(tp.cell_margins.unwrap().top.raw(), 100);
    }

    #[test]
    fn tbl_pr_tbl_look_attrs() {
        let (tp, _) =
            parse_tbl_pr(r#"<tblPr><tblLook firstRow="1" lastRow="0" noHBand="true"/></tblPr>"#);
        let l = tp.look.unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.last_row, Some(false));
        assert_eq!(l.no_h_band, Some(true));
    }

    /// [MS-OI29500] §2.1.1583: legacy `@val` hex bitfield must decode to the
    /// same flags as the modern explicit attributes. Bit positions:
    /// 0x0020 firstRow, 0x0040 lastRow, 0x0080 firstColumn, 0x0100 lastColumn,
    /// 0x0200 noHBand, 0x0400 noVBand. The Word-default `04A0` =
    /// firstRow + firstColumn + noVBand.
    #[test]
    fn tbl_pr_tbl_look_legacy_val_default() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="04A0"/></tblPr>"#);
        let l = tp.look.unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.last_row, Some(false));
        assert_eq!(l.first_column, Some(true));
        assert_eq!(l.last_column, Some(false));
        assert_eq!(l.no_h_band, Some(false));
        assert_eq!(l.no_v_band, Some(true));
    }

    /// Sample1's ITEM/NEEDED table uses `val="0620"` =
    /// firstRow + noHBand + noVBand. Without legacy decoding, banding is
    /// erroneously enabled and `band1Horz` CF paints inner borders.
    #[test]
    fn tbl_pr_tbl_look_legacy_val_suppresses_banding() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="0620"/></tblPr>"#);
        let l = tp.look.unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.no_h_band, Some(true));
        assert_eq!(l.no_v_band, Some(true));
    }

    #[test]
    fn tbl_pr_tbl_look_legacy_val_zero_clears_all() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="0000"/></tblPr>"#);
        let l = tp.look.unwrap();
        assert_eq!(l.first_row, Some(false));
        assert_eq!(l.last_row, Some(false));
        assert_eq!(l.first_column, Some(false));
        assert_eq!(l.last_column, Some(false));
        assert_eq!(l.no_h_band, Some(false));
        assert_eq!(l.no_v_band, Some(false));
    }

    /// Per [MS-OI29500]: when both legacy `val` and explicit attributes are
    /// specified, the explicit attribute wins per-flag. Sample3 has tables
    /// shaped like `<tblLook val="04A0" firstRow="1" noHBand="0" .../>`.
    #[test]
    fn tbl_pr_tbl_look_explicit_attrs_override_val() {
        let (tp, _) =
            parse_tbl_pr(r#"<tblPr><tblLook val="0000" firstRow="1" noVBand="1"/></tblPr>"#);
        let l = tp.look.unwrap();
        assert_eq!(l.first_row, Some(true), "explicit firstRow=1 wins");
        assert_eq!(l.last_row, Some(false), "from val=0000");
        assert_eq!(l.first_column, Some(false), "from val=0000");
        assert_eq!(l.no_v_band, Some(true), "explicit noVBand=1 wins");
    }

    #[test]
    fn tbl_pr_tbl_look_legacy_val_lowercase() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="04a0"/></tblPr>"#);
        let l = tp.look.unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.no_v_band, Some(true));
    }

    #[test]
    fn tbl_pr_overlap_and_positioning() {
        let (tp, _) = parse_tbl_pr(
            r#"<tblPr>
                <tblOverlap val="never"/>
                <tblpPr tblpX="100" tblpY="200" vertAnchor="page"
                        horzAnchor="margin" tblpXSpec="center"/>
            </tblPr>"#,
        );
        assert_eq!(tp.overlap, Some(TableOverlap::Never));
        let pos = tp.positioning.unwrap();
        assert_eq!(pos.x.unwrap().raw(), 100);
        assert_eq!(pos.y.unwrap().raw(), 200);
        assert_eq!(pos.vert_anchor, Some(crate::docx::model::TableAnchor::Page));
        assert_eq!(pos.x_align, Some(crate::docx::model::TableXAlign::Center));
    }

    // ── trPr ──

    fn parse_tr_pr(xml: &str) -> TableRowProperties {
        let x: TrPrXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn tr_pr_height_with_rule() {
        let tr = parse_tr_pr(r#"<trPr><trHeight val="440" hRule="atLeast"/></trPr>"#);
        let h = tr.height.unwrap();
        assert_eq!(h.value.raw(), 440);
        assert_eq!(h.rule, HeightRule::AtLeast);
    }

    #[test]
    fn tr_pr_is_header_and_cant_split() {
        let tr = parse_tr_pr(r#"<trPr><tblHeader/><cantSplit/></trPr>"#);
        assert_eq!(tr.is_header, Some(true));
        assert_eq!(tr.cant_split, Some(true));
    }

    #[test]
    fn tr_pr_grid_after_and_w_after() {
        let tr = parse_tr_pr(r#"<trPr><gridAfter val="2"/><wAfter w="500" type="dxa"/></trPr>"#);
        assert_eq!(tr.grid_after, 2);
        match tr.w_after.unwrap() {
            TableMeasure::Twips(d) => assert_eq!(d.raw(), 500),
            other => panic!("expected Twips, got {other:?}"),
        }
    }

    #[test]
    fn tr_pr_grid_before_and_w_before() {
        let tr = parse_tr_pr(r#"<trPr><gridBefore val="1"/><wBefore w="38" type="dxa"/></trPr>"#);
        assert_eq!(tr.grid_before, 1);
        match tr.w_before.unwrap() {
            TableMeasure::Twips(d) => assert_eq!(d.raw(), 38),
            other => panic!("expected Twips, got {other:?}"),
        }
    }

    #[test]
    fn tr_pr_grid_before_and_grid_after_default_zero() {
        let tr = parse_tr_pr(r#"<trPr/>"#);
        assert_eq!(tr.grid_before, 0);
        assert_eq!(tr.grid_after, 0);
        assert!(tr.w_before.is_none());
        assert!(tr.w_after.is_none());
    }

    // ── tcPr ──

    fn parse_tc_pr(xml: &str) -> TableCellProperties {
        let x: TcPrXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn tc_pr_width_and_borders() {
        let tc = parse_tc_pr(
            r#"<tcPr>
                <tcW w="2500" type="dxa"/>
                <tcBorders><top val="single"/><tl2br val="dotted"/></tcBorders>
            </tcPr>"#,
        );
        match tc.width.unwrap() {
            TableMeasure::Twips(d) => assert_eq!(d.raw(), 2500),
            other => panic!("expected Twips, got {other:?}"),
        }
        assert!(tc.borders.unwrap().tl2br.is_some());
    }

    #[test]
    fn tc_pr_vertical_align() {
        let tc = parse_tc_pr(r#"<tcPr><vAlign val="center"/></tcPr>"#);
        assert_eq!(tc.vertical_align, Some(CellVerticalAlign::Center));
    }

    #[test]
    fn tc_pr_v_merge_restart_and_continue() {
        let tc = parse_tc_pr(r#"<tcPr><vMerge val="restart"/></tcPr>"#);
        assert_eq!(tc.vertical_merge, Some(VerticalMerge::Restart));

        let tc = parse_tc_pr(r#"<tcPr><vMerge/></tcPr>"#);
        assert_eq!(tc.vertical_merge, Some(VerticalMerge::Continue));
    }

    #[test]
    fn tc_pr_grid_span_and_text_direction() {
        let tc = parse_tc_pr(r#"<tcPr><gridSpan val="3"/><textDirection val="tbRl"/></tcPr>"#);
        assert_eq!(tc.grid_span, Some(3));
        assert_eq!(
            tc.text_direction,
            Some(TextDirection::TopToBottomRightToLeft)
        );
    }

    #[test]
    fn tc_pr_no_wrap_and_cnf_style() {
        let tc = parse_tc_pr(r#"<tcPr><noWrap/><cnfStyle val="100000000000"/></tcPr>"#);
        assert_eq!(tc.no_wrap, Some(true));
        assert_eq!(tc.cnf_style, Some(CnfStyle::FIRST_ROW));
    }
}
