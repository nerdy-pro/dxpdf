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
use crate::docx::parse::primitives::st_enums::{StNumberFormat, StPageOrientation, StSectionMark};
use crate::docx::parse::primitives::OnOff;

/// `<w:sectPr>` root.
///
/// Children are collected into a single `$value` Vec rather than as named
/// fields because real DOCX files commonly interleave `<w:headerReference>`
/// and `<w:footerReference>` (one h+f pair per `type`), and quick-xml's
/// serde deserializer rejects non-contiguous repeats of a named field.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct SectPrXml {
    #[serde(rename = "@rsidR", default)]
    rsid_r: Option<String>,
    #[serde(rename = "@rsidRPr", default)]
    rsid_r_pr: Option<String>,
    #[serde(rename = "@rsidSect", default)]
    rsid_sect: Option<String>,

    #[serde(rename = "$value", default)]
    children: Vec<SectChildXml>,
}

#[derive(Clone, Debug, Deserialize)]
pub(crate) enum SectChildXml {
    #[serde(rename = "pgSz")]
    PgSz(PgSzXml),
    #[serde(rename = "pgMar")]
    PgMar(PgMarXml),
    #[serde(rename = "cols")]
    Cols(ColsXml),
    #[serde(rename = "docGrid")]
    DocGrid(DocGridXml),
    #[serde(rename = "headerReference")]
    HeaderRef(HfRefXml),
    #[serde(rename = "footerReference")]
    FooterRef(HfRefXml),
    #[serde(rename = "titlePg")]
    TitlePg(OnOff),
    #[serde(rename = "type")]
    Type(ValStSectionMark),
    #[serde(rename = "pgNumType")]
    PgNumType(PgNumTypeXml),
    #[serde(other)]
    Other,
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
pub(crate) struct ValStSectionMark {
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
        let mut page_size: Option<PgSzXml> = None;
        let mut page_margins: Option<PgMarXml> = None;
        let mut cols: Option<ColsXml> = None;
        let mut doc_grid: Option<DocGridXml> = None;
        let mut header_refs = SectionHeaderFooterRefs::default();
        let mut footer_refs = SectionHeaderFooterRefs::default();
        let mut title_page: Option<bool> = None;
        let mut section_type: Option<SectionType> = None;
        let mut page_number_type: Option<PgNumTypeXml> = None;

        for child in x.children {
            match child {
                SectChildXml::PgSz(v) => page_size = Some(v),
                SectChildXml::PgMar(v) => page_margins = Some(v),
                SectChildXml::Cols(v) => cols = Some(v),
                SectChildXml::DocGrid(v) => doc_grid = Some(v),
                SectChildXml::HeaderRef(r) => assign_hf_ref(&mut header_refs, r),
                SectChildXml::FooterRef(r) => assign_hf_ref(&mut footer_refs, r),
                SectChildXml::TitlePg(OnOff(b)) => title_page = Some(b),
                SectChildXml::Type(v) => section_type = Some(SectionType::from(v.val)),
                SectChildXml::PgNumType(v) => page_number_type = Some(v),
                SectChildXml::Other => {}
            }
        }

        Self {
            page_size: page_size.map(Into::into),
            page_margins: page_margins.map(Into::into),
            columns: cols.map(Into::into),
            doc_grid: doc_grid.map(Into::into),
            header_refs,
            footer_refs,
            title_page,
            section_type,
            page_number_type: page_number_type.map(Into::into),
            rsids: SectionRevisionIds {
                r: x.rsid_r.as_deref().and_then(RevisionSaveId::from_hex),
                r_pr: x.rsid_r_pr.as_deref().and_then(RevisionSaveId::from_hex),
                sect: x.rsid_sect.as_deref().and_then(RevisionSaveId::from_hex),
            },
        }
    }
}

fn assign_hf_ref(out: &mut SectionHeaderFooterRefs, r: HfRefXml) {
    let rel = RelId::new(&r.id);
    match r.ty {
        Some(StHdrFtr::First) => out.first = Some(rel),
        Some(StHdrFtr::Even) => out.even = Some(rel),
        Some(StHdrFtr::Default) | None => out.default = Some(rel),
    }
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
        assert_eq!(
            s.header_refs.default.as_ref().map(|r| r.as_str()),
            Some("rId1")
        );
        assert_eq!(
            s.header_refs.first.as_ref().map(|r| r.as_str()),
            Some("rId2")
        );
        assert_eq!(
            s.footer_refs.default.as_ref().map(|r| r.as_str()),
            Some("rId3")
        );
    }

    #[test]
    fn header_footer_refs_interleaved() {
        // Real docs commonly interleave header/footer refs by type
        // (even headers+footers, then default headers+footers, then first).
        let s = parse(
            r#"<sectPr>
                <headerReference id="rId6" type="even"/>
                <headerReference id="rId7" type="default"/>
                <footerReference id="rId8" type="even"/>
                <footerReference id="rId9" type="default"/>
                <headerReference id="rId10" type="first"/>
                <footerReference id="rId11" type="first"/>
            </sectPr>"#,
        );
        assert_eq!(
            s.header_refs.even.as_ref().map(|r| r.as_str()),
            Some("rId6")
        );
        assert_eq!(
            s.header_refs.default.as_ref().map(|r| r.as_str()),
            Some("rId7")
        );
        assert_eq!(
            s.header_refs.first.as_ref().map(|r| r.as_str()),
            Some("rId10")
        );
        assert_eq!(
            s.footer_refs.even.as_ref().map(|r| r.as_str()),
            Some("rId8")
        );
        assert_eq!(
            s.footer_refs.default.as_ref().map(|r| r.as_str()),
            Some("rId9")
        );
        assert_eq!(
            s.footer_refs.first.as_ref().map(|r| r.as_str()),
            Some("rId11")
        );
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
        let s = parse(
            r#"<sectPr><pgNumType fmt="decimal" start="1" chapStyle="1" chapSep="emDash"/></sectPr>"#,
        );
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
