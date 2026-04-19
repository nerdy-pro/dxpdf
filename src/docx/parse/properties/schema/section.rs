//! `<w:sectPr>` schema (§17.6 section properties).
//!
//! RSIDs on the `<w:sectPr>` start tag are captured via `@rsidR`, `@rsidRPr`,
//! `@rsidSect` because serde reads them as ordinary attributes on the root.

use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, FractionPoints, Twips};
use crate::docx::model::{
    ChapterSeparator, ColumnDefinition, Columns, DocGrid, DocGridType, NumberFormat, PageMargins,
    PageNumberType, PageSize, RelId, RevisionSaveId, SectionHeaderFooterRefs, SectionProperties,
    SectionRevisionIds, SectionType,
};
use crate::docx::parse::primitives::st_enums::{
    StNumberFormat, StPageOrientation, StSectionMark,
};
use crate::docx::parse::primitives::OnOff;

/// `<w:sectPr>` root.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct SectPrXml {
    #[serde(rename = "@rsidR", default)]
    rsid_r: Option<String>,
    #[serde(rename = "@rsidRPr", default)]
    rsid_r_pr: Option<String>,
    #[serde(rename = "@rsidSect", default)]
    rsid_sect: Option<String>,

    #[serde(rename = "pgSz", default)]
    pg_sz: Option<PgSzXml>,
    #[serde(rename = "pgMar", default)]
    pg_mar: Option<PgMarXml>,
    #[serde(default)]
    cols: Option<ColsXml>,
    #[serde(rename = "docGrid", default)]
    doc_grid: Option<DocGridXml>,
    #[serde(rename = "headerReference", default)]
    header_refs: Vec<HfRefXml>,
    #[serde(rename = "footerReference", default)]
    footer_refs: Vec<HfRefXml>,
    #[serde(rename = "titlePg", default)]
    title_pg: Option<OnOff>,
    #[serde(rename = "type", default)]
    ty: Option<ValStSectionMark>,
    #[serde(rename = "pgNumType", default)]
    pg_num_type: Option<PgNumTypeXml>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct PgSzXml {
    #[serde(rename = "@w", default)]
    w: Option<Dimension<Twips>>,
    #[serde(rename = "@h", default)]
    h: Option<Dimension<Twips>>,
    #[serde(rename = "@orient", default)]
    orient: Option<StPageOrientation>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct PgMarXml {
    #[serde(rename = "@top", default)]
    top: Option<Dimension<Twips>>,
    #[serde(rename = "@right", default)]
    right: Option<Dimension<Twips>>,
    #[serde(rename = "@bottom", default)]
    bottom: Option<Dimension<Twips>>,
    #[serde(rename = "@left", default)]
    left: Option<Dimension<Twips>>,
    #[serde(rename = "@header", default)]
    header: Option<Dimension<Twips>>,
    #[serde(rename = "@footer", default)]
    footer: Option<Dimension<Twips>>,
    #[serde(rename = "@gutter", default)]
    gutter: Option<Dimension<Twips>>,
}

#[derive(Clone, Debug, Deserialize)]
pub(crate) struct ColsXml {
    #[serde(rename = "@num", default)]
    num: Option<u32>,
    #[serde(rename = "@space", default)]
    space: Option<Dimension<Twips>>,
    #[serde(rename = "@equalWidth", default)]
    equal_width: Option<OnOffFromAttr>,
    #[serde(rename = "col", default)]
    cols: Vec<ColXml>,
}

use crate::docx::parse::primitives::AttrBool as OnOffFromAttr;

#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct ColXml {
    #[serde(rename = "@w", default)]
    w: Option<Dimension<Twips>>,
    #[serde(rename = "@space", default)]
    space: Option<Dimension<Twips>>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct DocGridXml {
    #[serde(rename = "@type", default)]
    ty: Option<StDocGrid>,
    #[serde(rename = "@linePitch", default)]
    line_pitch: Option<Dimension<Twips>>,
    #[serde(rename = "@charSpace", default)]
    char_space: Option<Dimension<FractionPoints>>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum StDocGrid {
    Default,
    Lines,
    LinesAndChars,
    SnapToChars,
}

impl From<StDocGrid> for DocGridType {
    fn from(s: StDocGrid) -> Self {
        match s {
            StDocGrid::Default => Self::Default,
            StDocGrid::Lines => Self::Lines,
            StDocGrid::LinesAndChars => Self::LinesAndChars,
            StDocGrid::SnapToChars => Self::SnapToChars,
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
pub(crate) struct HfRefXml {
    #[serde(rename = "@id")]
    id: String,
    #[serde(rename = "@type", default)]
    ty: Option<StHdrFtr>,
}

#[derive(Clone, Copy, Debug, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
enum StHdrFtr {
    Default,
    First,
    Even,
}

#[derive(Clone, Debug, Deserialize)]
pub(crate) struct PgNumTypeXml {
    #[serde(rename = "@fmt", default)]
    fmt: Option<StNumberFormat>,
    #[serde(rename = "@start", default)]
    start: Option<u32>,
    #[serde(rename = "@chapStyle", default)]
    chap_style: Option<u32>,
    #[serde(rename = "@chapSep", default)]
    chap_sep: Option<StChapSep>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum StChapSep {
    Hyphen,
    Period,
    Colon,
    EmDash,
    EnDash,
}

impl From<StChapSep> for ChapterSeparator {
    fn from(s: StChapSep) -> Self {
        match s {
            StChapSep::Hyphen => Self::Hyphen,
            StChapSep::Period => Self::Period,
            StChapSep::Colon => Self::Colon,
            StChapSep::EmDash => Self::EmDash,
            StChapSep::EnDash => Self::EnDash,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
struct ValStSectionMark {
    #[serde(rename = "@val")]
    val: StSectionMark,
}

impl From<PgSzXml> for PageSize {
    fn from(x: PgSzXml) -> Self {
        Self {
            width: x.w,
            height: x.h,
            orientation: x.orient.map(Into::into),
        }
    }
}

impl From<PgMarXml> for PageMargins {
    fn from(x: PgMarXml) -> Self {
        Self {
            top: x.top,
            right: x.right,
            bottom: x.bottom,
            left: x.left,
            header: x.header,
            footer: x.footer,
            gutter: x.gutter,
        }
    }
}

impl From<ColsXml> for Columns {
    fn from(x: ColsXml) -> Self {
        Self {
            count: x.num,
            space: x.space,
            equal_width: x.equal_width.map(|v| v.0),
            columns: x.cols.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<ColXml> for ColumnDefinition {
    fn from(x: ColXml) -> Self {
        Self {
            width: x.w,
            space: x.space,
        }
    }
}

impl From<DocGridXml> for DocGrid {
    fn from(x: DocGridXml) -> Self {
        Self {
            grid_type: x.ty.map(Into::into),
            line_pitch: x.line_pitch,
            char_space: x.char_space,
        }
    }
}

impl From<PgNumTypeXml> for PageNumberType {
    fn from(x: PgNumTypeXml) -> Self {
        Self {
            format: x.fmt.map(NumberFormat::from),
            start: x.start,
            chap_style: x.chap_style,
            chap_sep: x.chap_sep.map(Into::into),
        }
    }
}

impl From<SectPrXml> for SectionProperties {
    fn from(x: SectPrXml) -> Self {
        let (header_refs, footer_refs) = split_hf_refs(&x.header_refs, &x.footer_refs);
        Self {
            page_size: x.pg_sz.map(Into::into),
            page_margins: x.pg_mar.map(Into::into),
            columns: x.cols.map(Into::into),
            doc_grid: x.doc_grid.map(Into::into),
            header_refs,
            footer_refs,
            title_page: x.title_pg.map(|OnOff(b)| b),
            section_type: x.ty.map(|v| SectionType::from(v.val)),
            page_number_type: x.pg_num_type.map(Into::into),
            rsids: SectionRevisionIds {
                r: x.rsid_r.as_deref().and_then(RevisionSaveId::from_hex),
                r_pr: x.rsid_r_pr.as_deref().and_then(RevisionSaveId::from_hex),
                sect: x.rsid_sect.as_deref().and_then(RevisionSaveId::from_hex),
            },
        }
    }
}

fn split_hf_refs(
    headers: &[HfRefXml],
    footers: &[HfRefXml],
) -> (SectionHeaderFooterRefs, SectionHeaderFooterRefs) {
    let collect = |refs: &[HfRefXml]| {
        let mut out = SectionHeaderFooterRefs::default();
        for r in refs {
            let rel = RelId::new(&r.id);
            match r.ty {
                Some(StHdrFtr::First) => out.first = Some(rel),
                Some(StHdrFtr::Even) => out.even = Some(rel),
                Some(StHdrFtr::Default) | None => out.default = Some(rel),
            }
        }
        out
    };
    (collect(headers), collect(footers))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::PageOrientation;

    fn parse(xml: &str) -> SectionProperties {
        let x: SectPrXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn page_size_with_orientation() {
        let s = parse(r#"<sectPr><pgSz w="12240" h="15840" orient="landscape"/></sectPr>"#);
        let ps = s.page_size.unwrap();
        assert_eq!(ps.width.unwrap().raw(), 12240);
        assert_eq!(ps.height.unwrap().raw(), 15840);
        assert_eq!(ps.orientation, Some(PageOrientation::Landscape));
    }

    #[test]
    fn page_margins_all_seven() {
        let s = parse(
            r#"<sectPr><pgMar top="1440" right="1800" bottom="1440" left="1800"
                 header="720" footer="720" gutter="0"/></sectPr>"#,
        );
        let pm = s.page_margins.unwrap();
        assert_eq!(pm.top.unwrap().raw(), 1440);
        assert_eq!(pm.header.unwrap().raw(), 720);
        assert_eq!(pm.gutter.unwrap().raw(), 0);
    }

    #[test]
    fn cols_with_child_definitions() {
        let s = parse(
            r#"<sectPr><cols num="2" space="720" equalWidth="false">
                <col w="3000" space="500"/>
                <col w="4000"/>
            </cols></sectPr>"#,
        );
        let c = s.columns.unwrap();
        assert_eq!(c.count, Some(2));
        assert_eq!(c.equal_width, Some(false));
        assert_eq!(c.columns.len(), 2);
        assert_eq!(c.columns[0].width.unwrap().raw(), 3000);
    }

    #[test]
    fn doc_grid_basic() {
        let s = parse(r#"<sectPr><docGrid type="lines" linePitch="360"/></sectPr>"#);
        let g = s.doc_grid.unwrap();
        assert_eq!(g.grid_type, Some(DocGridType::Lines));
        assert_eq!(g.line_pitch.unwrap().raw(), 360);
    }

    #[test]
    fn header_footer_refs_split_by_type() {
        let s = parse(
            r#"<sectPr>
                <headerReference id="rId1" type="default"/>
                <headerReference id="rId2" type="first"/>
                <footerReference id="rId3"/>
            </sectPr>"#,
        );
        assert_eq!(s.header_refs.default.as_ref().map(|r| r.as_str()), Some("rId1"));
        assert_eq!(s.header_refs.first.as_ref().map(|r| r.as_str()), Some("rId2"));
        assert_eq!(s.footer_refs.default.as_ref().map(|r| r.as_str()), Some("rId3"));
    }

    #[test]
    fn section_type_roundtrip() {
        let s = parse(r#"<sectPr><type val="evenPage"/></sectPr>"#);
        assert_eq!(s.section_type, Some(SectionType::EvenPage));
    }

    #[test]
    fn title_pg_toggle() {
        let s = parse(r#"<sectPr><titlePg/></sectPr>"#);
        assert_eq!(s.title_page, Some(true));
    }

    #[test]
    fn page_num_type_with_chap_sep() {
        let s = parse(r#"<sectPr><pgNumType fmt="decimal" start="1" chapStyle="1" chapSep="emDash"/></sectPr>"#);
        let p = s.page_number_type.unwrap();
        assert_eq!(p.format, Some(NumberFormat::Decimal));
        assert_eq!(p.start, Some(1));
        assert_eq!(p.chap_sep, Some(ChapterSeparator::EmDash));
    }

    #[test]
    fn rsids_captured_from_root_attrs() {
        let s = parse(r#"<sectPr rsidR="001A2B3C" rsidSect="00A1B2C3"/>"#);
        assert!(s.rsids.r.is_some());
        assert!(s.rsids.sect.is_some());
        assert!(s.rsids.r_pr.is_none());
    }

    #[test]
    fn empty_sect_pr_is_default() {
        let s = parse(r#"<sectPr/>"#);
        assert!(s.page_size.is_none());
        assert!(s.columns.is_none());
    }
}
