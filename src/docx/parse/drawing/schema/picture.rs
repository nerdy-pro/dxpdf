//! DrawingML picture schema (§19.3.1.37 pic:pic).
//!
//! `<pic:pic>` has three children: non-visual picture properties
//! (`<pic:nvPicPr>`), blip fill (`<pic:blipFill>`), and shape properties
//! (`<pic:spPr>`). All three are fully modeled.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::model::{CnvPicProperties, DocProperties, NvPicProperties, PicLocks, Picture};

use super::fill::{AttrBool, BlipFillXml};
use super::shape::SpPrXml;

#[derive(Deserialize)]
pub(crate) struct PictureXml {
    #[serde(rename = "nvPicPr")]
    pub(crate) nv_pic_pr: NvPicPrXml,
    #[serde(rename = "blipFill")]
    pub(crate) blip_fill: BlipFillXml,
    #[serde(rename = "spPr", default)]
    pub(crate) sp_pr: Option<SpPrXml>,
}

#[derive(Debug, Deserialize)]
pub struct NvPicPrXml {
    #[serde(rename = "cNvPr")]
    pub cnv_pr: CNvPrXml,
    #[serde(rename = "cNvPicPr", default)]
    pub cnv_pic_pr: Option<CNvPicPrXml>,
}

#[derive(Debug, Deserialize)]
pub struct CNvPrXml {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@descr", default)]
    pub descr: Option<String>,
    #[serde(rename = "@hidden", default)]
    pub hidden: Option<AttrBool>,
    #[serde(rename = "@title", default)]
    pub title: Option<String>,
    // Children like hlinkClick, hlinkHover, extLst accepted but ignored.
}

#[derive(Debug, Deserialize)]
pub struct CNvPicPrXml {
    #[serde(rename = "@preferRelativeResize", default)]
    pub prefer_relative_resize: Option<AttrBool>,
    #[serde(rename = "picLocks", default)]
    pub pic_locks: Option<PicLocksXml>,
}

#[derive(Debug, Deserialize)]
pub struct PicLocksXml {
    #[serde(rename = "@noChangeAspect", default)]
    pub no_change_aspect: Option<AttrBool>,
    #[serde(rename = "@noCrop", default)]
    pub no_crop: Option<AttrBool>,
    #[serde(rename = "@noResize", default)]
    pub no_resize: Option<AttrBool>,
    #[serde(rename = "@noMove", default)]
    pub no_move: Option<AttrBool>,
    #[serde(rename = "@noRot", default)]
    pub no_rot: Option<AttrBool>,
    #[serde(rename = "@noSelect", default)]
    pub no_select: Option<AttrBool>,
    #[serde(rename = "@noEditPoints", default)]
    pub no_edit_points: Option<AttrBool>,
    #[serde(rename = "@noAdjustHandles", default)]
    pub no_adjust_handles: Option<AttrBool>,
    #[serde(rename = "@noChangeArrowheads", default)]
    pub no_change_arrowheads: Option<AttrBool>,
    #[serde(rename = "@noChangeShapeType", default)]
    pub no_change_shape_type: Option<AttrBool>,
    #[serde(rename = "@noGrp", default)]
    pub no_grp: Option<AttrBool>,
}

// ── Conversion to model ───────────────────────────────────────────────────

impl From<PictureXml> for Picture {
    fn from(x: PictureXml) -> Self {
        Self {
            nv_pic_pr: x.nv_pic_pr.into(),
            blip_fill: x.blip_fill.into(),
            shape_properties: x.sp_pr.map(Into::into),
        }
    }
}

impl From<NvPicPrXml> for NvPicProperties {
    fn from(x: NvPicPrXml) -> Self {
        Self {
            cnv_pr: x.cnv_pr.into(),
            cnv_pic_pr: x.cnv_pic_pr.map(Into::into),
        }
    }
}

impl From<CNvPrXml> for DocProperties {
    fn from(x: CNvPrXml) -> Self {
        Self {
            id: x.id,
            name: x.name,
            description: x.descr,
            hidden: x.hidden.map(|b| b.0),
            title: x.title,
        }
    }
}

impl From<CNvPicPrXml> for CnvPicProperties {
    fn from(x: CNvPicPrXml) -> Self {
        Self {
            prefer_relative_resize: x.prefer_relative_resize.map(|b| b.0),
            pic_locks: x.pic_locks.map(Into::into),
        }
    }
}

impl From<PicLocksXml> for PicLocks {
    fn from(x: PicLocksXml) -> Self {
        Self {
            no_change_aspect: x.no_change_aspect.map(|b| b.0),
            no_crop: x.no_crop.map(|b| b.0),
            no_resize: x.no_resize.map(|b| b.0),
            no_move: x.no_move.map(|b| b.0),
            no_rot: x.no_rot.map(|b| b.0),
            no_select: x.no_select.map(|b| b.0),
            no_edit_points: x.no_edit_points.map(|b| b.0),
            no_adjust_handles: x.no_adjust_handles.map(|b| b.0),
            no_change_arrowheads: x.no_change_arrowheads.map(|b| b.0),
            no_change_shape_type: x.no_change_shape_type.map(|b| b.0),
            no_grp: x.no_grp.map(|b| b.0),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::BlipFillKind;

    fn parse(xml: &str) -> Picture {
        let wrapped = format!(
            r#"<wrap xmlns:pic="urn:pic" xmlns:a="urn:a" xmlns:r="urn:r">{}</wrap>"#,
            xml
        );
        #[derive(Deserialize)]
        struct Wrap {
            pic: PictureXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.pic.into()
    }

    #[test]
    fn minimal_picture_with_blip_embed() {
        let p = parse(
            r#"<pic>
                <nvPicPr>
                    <cNvPr id="1" name="image1"/>
                </nvPicPr>
                <blipFill>
                    <blip r:embed="rId1"/>
                    <stretch/>
                </blipFill>
            </pic>"#,
        );
        assert_eq!(p.nv_pic_pr.cnv_pr.id, 1);
        assert_eq!(p.nv_pic_pr.cnv_pr.name, "image1");
        assert!(p.nv_pic_pr.cnv_pic_pr.is_none());
        assert_eq!(
            p.blip_fill
                .blip
                .as_ref()
                .and_then(|b| b.embed.as_ref())
                .map(|r| r.as_str()),
            Some("rId1")
        );
        assert!(matches!(p.blip_fill.fill_kind, BlipFillKind::Stretch(_)));
    }

    #[test]
    fn picture_with_description_and_title() {
        let p = parse(
            r#"<pic>
                <nvPicPr>
                    <cNvPr id="2" name="logo" descr="Company logo" title="Logo"/>
                </nvPicPr>
                <blipFill><blip r:embed="rId2"/></blipFill>
            </pic>"#,
        );
        assert_eq!(
            p.nv_pic_pr.cnv_pr.description.as_deref(),
            Some("Company logo")
        );
        assert_eq!(p.nv_pic_pr.cnv_pr.title.as_deref(), Some("Logo"));
    }

    #[test]
    fn picture_with_hidden_and_relative_resize() {
        let p = parse(
            r#"<pic>
                <nvPicPr>
                    <cNvPr id="3" name="hidden" hidden="1"/>
                    <cNvPicPr preferRelativeResize="0"/>
                </nvPicPr>
                <blipFill><blip r:embed="rId3"/></blipFill>
            </pic>"#,
        );
        assert_eq!(p.nv_pic_pr.cnv_pr.hidden, Some(true));
        assert_eq!(
            p.nv_pic_pr
                .cnv_pic_pr
                .as_ref()
                .unwrap()
                .prefer_relative_resize,
            Some(false)
        );
    }

    #[test]
    fn picture_with_pic_locks() {
        let p = parse(
            r#"<pic>
                <nvPicPr>
                    <cNvPr id="4" name="locked"/>
                    <cNvPicPr>
                        <picLocks noChangeAspect="1" noCrop="0" noResize="true"/>
                    </cNvPicPr>
                </nvPicPr>
                <blipFill><blip r:embed="rId4"/></blipFill>
            </pic>"#,
        );
        let locks = p
            .nv_pic_pr
            .cnv_pic_pr
            .as_ref()
            .and_then(|c| c.pic_locks.as_ref())
            .expect("picLocks");
        assert_eq!(locks.no_change_aspect, Some(true));
        assert_eq!(locks.no_crop, Some(false));
        assert_eq!(locks.no_resize, Some(true));
    }

    #[test]
    fn picture_sp_pr_fully_parsed() {
        use crate::docx::model::{PresetShapeType, ShapeGeometry};
        let p = parse(
            r#"<pic>
                <nvPicPr><cNvPr id="5" name="with-sp"/></nvPicPr>
                <blipFill><blip r:embed="rId5"/></blipFill>
                <spPr>
                    <xfrm rot="900000"><off x="10" y="20"/><ext cx="100" cy="200"/></xfrm>
                    <prstGeom prst="roundRect"/>
                </spPr>
            </pic>"#,
        );
        let sp = p.shape_properties.expect("spPr parsed");
        let t = sp.transform.unwrap();
        assert_eq!(t.rotation.unwrap().raw(), 900_000);
        assert_eq!(t.extent.unwrap().width.raw(), 100);
        match sp.geometry {
            Some(ShapeGeometry::Preset(g)) => assert_eq!(g.preset, PresetShapeType::RoundRect),
            other => panic!("expected Preset, got {other:?}"),
        }
    }

    #[test]
    fn blip_with_crop_source_rect() {
        let p = parse(
            r#"<pic>
                <nvPicPr><cNvPr id="6" name="cropped"/></nvPicPr>
                <blipFill>
                    <blip r:embed="rId6" cstate="hqprint"/>
                    <srcRect l="5000" t="0" r="5000" b="0"/>
                    <stretch><fillRect/></stretch>
                </blipFill>
            </pic>"#,
        );
        assert_eq!(
            p.blip_fill.src_rect.as_ref().unwrap().left.unwrap().raw(),
            5000
        );
        assert_eq!(
            p.blip_fill.blip.as_ref().unwrap().compression,
            Some(crate::docx::model::BlipCompression::Hqprint)
        );
    }

    #[test]
    fn blip_tile_fill() {
        let p = parse(
            r#"<pic>
                <nvPicPr><cNvPr id="7" name="tiled"/></nvPicPr>
                <blipFill>
                    <blip r:embed="rId7"/>
                    <tile tx="0" ty="0" sx="50000" sy="50000" algn="tl"/>
                </blipFill>
            </pic>"#,
        );
        match p.blip_fill.fill_kind {
            BlipFillKind::Tile(t) => {
                assert_eq!(t.sx.unwrap().raw(), 50000);
                assert_eq!(t.alignment, Some(crate::docx::model::RectAlignment::Tl));
            }
            other => panic!("expected Tile, got {other:?}"),
        }
    }
}
