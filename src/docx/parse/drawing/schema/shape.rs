//! DrawingML shape schema (§20.1.2.2.35 CT_ShapeProperties + §14.5 wps:wsp).
//!
//! Consumes color, fill, stroke, effect, and geometry schemas. Adds
//! `Transform2D`, `PresetGeometryDef` (with the ~150-variant
//! `PresetShapeType` enum), `BodyProperties`, and the `WordProcessingShape`
//! wrapper.
//!
//! Text box content inside `<wps:txbx>` uses `body_schema::BlockChildXml`
//! and is materialized into `Vec<Block>` via the body schema's
//! `convert_container` — same pipeline as body and notes.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::{Deserialize, Deserializer};

use crate::docx::dimension::{Dimension, Emu, SixtieThousandthDeg};
use crate::docx::geometry::Offset;
use crate::docx::model::{
    BlackWhiteMode, Block, BodyProperties, DrawingFill, GeomGuide, PresetGeometryDef,
    PresetShapeType, ShapeGeometry, ShapeProperties, TextAnchoringType, TextAutoFit,
    TextVerticalType, TextWrappingType, Transform2D, WordProcessingShape,
};

use super::effect::EffectListXml;
use super::fill::{AttrBool, DrawingFillXml};
use super::geometry::CustomGeometryXml;
use super::picture::CNvPrXml;
use super::stroke::OutlineXml;

// ── spPr (§20.1.2.2.35) ───────────────────────────────────────────────────

#[derive(Debug, Deserialize, Default)]
pub struct SpPrXml {
    #[serde(rename = "@bwMode", default)]
    pub bw_mode: Option<StBlackWhiteMode>,
    #[serde(rename = "xfrm", default)]
    pub xfrm: Option<XfrmXml>,
    #[serde(rename = "prstGeom", default)]
    pub prst_geom: Option<PrstGeomXml>,
    #[serde(rename = "custGeom", default)]
    pub cust_geom: Option<CustomGeometryXml>,
    // Six fill variants — flatten by routing through DrawingFillXml.
    #[serde(rename = "noFill", default)]
    pub no_fill: Option<super::fill::Empty>,
    #[serde(rename = "solidFill", default)]
    pub solid_fill: Option<super::fill::SolidFillXml>,
    #[serde(rename = "gradFill", default)]
    pub grad_fill: Option<super::fill::GradFillXml>,
    #[serde(rename = "blipFill", default)]
    pub blip_fill: Option<super::fill::BlipFillXml>,
    #[serde(rename = "pattFill", default)]
    pub patt_fill: Option<super::fill::PattFillXml>,
    #[serde(rename = "grpFill", default)]
    pub grp_fill: Option<super::fill::Empty>,
    #[serde(rename = "ln", default)]
    pub ln: Option<OutlineXml>,
    #[serde(rename = "effectLst", default)]
    pub effect_lst: Option<EffectListXml>,
}

impl From<SpPrXml> for ShapeProperties {
    fn from(x: SpPrXml) -> Self {
        let geometry = if let Some(p) = x.prst_geom {
            Some(ShapeGeometry::Preset(p.into()))
        } else {
            x.cust_geom.map(|c| ShapeGeometry::Custom(c.into()))
        };
        let fill = pick_fill(
            x.no_fill,
            x.grp_fill,
            x.solid_fill,
            x.grad_fill,
            x.blip_fill,
            x.patt_fill,
        );
        Self {
            bw_mode: x.bw_mode.map(Into::into),
            transform: x.xfrm.map(Into::into),
            geometry,
            fill,
            outline: x.ln.map(Into::into),
            effect_list: x.effect_lst.map(Into::into),
        }
    }
}

fn pick_fill(
    no: Option<super::fill::Empty>,
    grp: Option<super::fill::Empty>,
    solid: Option<super::fill::SolidFillXml>,
    grad: Option<super::fill::GradFillXml>,
    blip: Option<super::fill::BlipFillXml>,
    patt: Option<super::fill::PattFillXml>,
) -> Option<DrawingFill> {
    let f = if let Some(e) = no {
        Some(DrawingFillXml::NoFill(e))
    } else if let Some(e) = grp {
        Some(DrawingFillXml::GrpFill(e))
    } else if let Some(s) = solid {
        Some(DrawingFillXml::SolidFill(s))
    } else if let Some(g) = grad {
        Some(DrawingFillXml::GradFill(g))
    } else if let Some(b) = blip {
        Some(DrawingFillXml::BlipFill(b))
    } else {
        patt.map(DrawingFillXml::PattFill)
    };
    f.map(Into::into)
}

// ── xfrm (§20.1.7.6) ──────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct XfrmXml {
    #[serde(rename = "@rot", default)]
    pub rot: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@flipH", default)]
    pub flip_h: Option<AttrBool>,
    #[serde(rename = "@flipV", default)]
    pub flip_v: Option<AttrBool>,
    #[serde(rename = "off", default)]
    pub off: Option<OffXml>,
    #[serde(rename = "ext", default)]
    pub ext: Option<ExtXml>,
}

#[derive(Debug, Deserialize)]
pub struct OffXml {
    #[serde(rename = "@x")]
    pub x: Dimension<Emu>,
    #[serde(rename = "@y")]
    pub y: Dimension<Emu>,
}

#[derive(Debug, Deserialize)]
pub struct ExtXml {
    #[serde(rename = "@cx")]
    pub cx: Dimension<Emu>,
    #[serde(rename = "@cy")]
    pub cy: Dimension<Emu>,
}

impl From<XfrmXml> for Transform2D {
    fn from(x: XfrmXml) -> Self {
        use crate::docx::geometry::Size;
        Self {
            rotation: x.rot,
            flip_h: x.flip_h.map(|b| b.0),
            flip_v: x.flip_v.map(|b| b.0),
            offset: x.off.map(|o| Offset::new(o.x, o.y)),
            extent: x.ext.map(|e| Size::new(e.cx, e.cy)),
        }
    }
}

// ── prstGeom (§20.1.9.18) ─────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct PrstGeomXml {
    #[serde(rename = "@prst")]
    pub prst: PresetShapeTypeXml,
    #[serde(rename = "avLst", default)]
    pub av_lst: Option<super::geometry::GdListXml>,
}

impl From<PrstGeomXml> for PresetGeometryDef {
    fn from(x: PrstGeomXml) -> Self {
        Self {
            preset: x.prst.0,
            adjust_values: x
                .av_lst
                .map(|l| {
                    l.guides
                        .into_iter()
                        .map(|g| GeomGuide {
                            name: g.name,
                            formula: g.fmla,
                        })
                        .collect()
                })
                .unwrap_or_default(),
        }
    }
}

/// Wrapper that tolerates unknown preset names by falling back to
/// `PresetShapeType::Other`. Known values go through `rename_all="camelCase"`.
#[derive(Debug)]
pub struct PresetShapeTypeXml(pub PresetShapeType);

impl<'de> Deserialize<'de> for PresetShapeTypeXml {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        Ok(Self(map_preset_shape(&s)))
    }
}

fn map_preset_shape(val: &str) -> PresetShapeType {
    use PresetShapeType as P;
    match val {
        "rect" => P::Rect,
        "roundRect" => P::RoundRect,
        "ellipse" => P::Ellipse,
        "triangle" => P::Triangle,
        "rtTriangle" => P::RtTriangle,
        "diamond" => P::Diamond,
        "parallelogram" => P::Parallelogram,
        "trapezoid" => P::Trapezoid,
        "pentagon" => P::Pentagon,
        "hexagon" => P::Hexagon,
        "octagon" => P::Octagon,
        "star4" => P::Star4,
        "star5" => P::Star5,
        "star6" => P::Star6,
        "star8" => P::Star8,
        "star10" => P::Star10,
        "star12" => P::Star12,
        "star16" => P::Star16,
        "star24" => P::Star24,
        "star32" => P::Star32,
        "line" => P::Line,
        "plus" => P::Plus,
        "can" => P::Can,
        "cube" => P::Cube,
        "donut" => P::Donut,
        "noSmoking" => P::NoSmoking,
        "blockArc" => P::BlockArc,
        "heart" => P::Heart,
        "sun" => P::Sun,
        "moon" => P::Moon,
        "smileyFace" => P::SmileyFace,
        "lightningBolt" => P::LightningBolt,
        "cloud" => P::Cloud,
        "arc" => P::Arc,
        "plaque" => P::Plaque,
        "frame" => P::Frame,
        "bevel" => P::Bevel,
        "foldedCorner" => P::FoldedCorner,
        "chevron" => P::Chevron,
        "homePlate" => P::HomePlate,
        "ribbon" => P::Ribbon,
        "ribbon2" => P::Ribbon2,
        "pie" => P::Pie,
        "pieWedge" => P::PieWedge,
        "chord" => P::Chord,
        "teardrop" => P::Teardrop,
        "arrow" => P::Arrow,
        "leftArrow" => P::LeftArrow,
        "rightArrow" => P::RightArrow,
        "upArrow" => P::UpArrow,
        "downArrow" => P::DownArrow,
        "leftRightArrow" => P::LeftRightArrow,
        "upDownArrow" => P::UpDownArrow,
        "quadArrow" => P::QuadArrow,
        "bentArrow" => P::BentArrow,
        "uturnArrow" => P::UturnArrow,
        "circularArrow" => P::CircularArrow,
        "curvedRightArrow" => P::CurvedRightArrow,
        "curvedLeftArrow" => P::CurvedLeftArrow,
        "curvedUpArrow" => P::CurvedUpArrow,
        "curvedDownArrow" => P::CurvedDownArrow,
        "stripedRightArrow" => P::StripedRightArrow,
        "notchedRightArrow" => P::NotchedRightArrow,
        "bentUpArrow" => P::BentUpArrow,
        "leftUpArrow" => P::LeftUpArrow,
        "leftRightUpArrow" => P::LeftRightUpArrow,
        "leftArrowCallout" => P::LeftArrowCallout,
        "rightArrowCallout" => P::RightArrowCallout,
        "upArrowCallout" => P::UpArrowCallout,
        "downArrowCallout" => P::DownArrowCallout,
        "leftRightArrowCallout" => P::LeftRightArrowCallout,
        "upDownArrowCallout" => P::UpDownArrowCallout,
        "quadArrowCallout" => P::QuadArrowCallout,
        "swooshArrow" => P::SwooshArrow,
        "leftCircularArrow" => P::LeftCircularArrow,
        "leftRightCircularArrow" => P::LeftRightCircularArrow,
        "callout1" => P::Callout1,
        "callout2" => P::Callout2,
        "callout3" => P::Callout3,
        "accentCallout1" => P::AccentCallout1,
        "accentCallout2" => P::AccentCallout2,
        "accentCallout3" => P::AccentCallout3,
        "borderCallout1" => P::BorderCallout1,
        "borderCallout2" => P::BorderCallout2,
        "borderCallout3" => P::BorderCallout3,
        "accentBorderCallout1" => P::AccentBorderCallout1,
        "accentBorderCallout2" => P::AccentBorderCallout2,
        "accentBorderCallout3" => P::AccentBorderCallout3,
        "wedgeRectCallout" => P::WedgeRectCallout,
        "wedgeRoundRectCallout" => P::WedgeRoundRectCallout,
        "wedgeEllipseCallout" => P::WedgeEllipseCallout,
        "cloudCallout" => P::CloudCallout,
        "leftBracket" => P::LeftBracket,
        "rightBracket" => P::RightBracket,
        "leftBrace" => P::LeftBrace,
        "rightBrace" => P::RightBrace,
        "bracketPair" => P::BracketPair,
        "bracePair" => P::BracePair,
        "straightConnector1" => P::StraightConnector1,
        "bentConnector2" => P::BentConnector2,
        "bentConnector3" => P::BentConnector3,
        "bentConnector4" => P::BentConnector4,
        "bentConnector5" => P::BentConnector5,
        "curvedConnector2" => P::CurvedConnector2,
        "curvedConnector3" => P::CurvedConnector3,
        "curvedConnector4" => P::CurvedConnector4,
        "curvedConnector5" => P::CurvedConnector5,
        "flowChartProcess" => P::FlowChartProcess,
        "flowChartDecision" => P::FlowChartDecision,
        "flowChartInputOutput" => P::FlowChartInputOutput,
        "flowChartPredefinedProcess" => P::FlowChartPredefinedProcess,
        "flowChartInternalStorage" => P::FlowChartInternalStorage,
        "flowChartDocument" => P::FlowChartDocument,
        "flowChartMultidocument" => P::FlowChartMultidocument,
        "flowChartTerminator" => P::FlowChartTerminator,
        "flowChartPreparation" => P::FlowChartPreparation,
        "flowChartManualInput" => P::FlowChartManualInput,
        "flowChartManualOperation" => P::FlowChartManualOperation,
        "flowChartConnector" => P::FlowChartConnector,
        "flowChartPunchedCard" => P::FlowChartPunchedCard,
        "flowChartPunchedTape" => P::FlowChartPunchedTape,
        "flowChartSummingJunction" => P::FlowChartSummingJunction,
        "flowChartOr" => P::FlowChartOr,
        "flowChartCollate" => P::FlowChartCollate,
        "flowChartSort" => P::FlowChartSort,
        "flowChartExtract" => P::FlowChartExtract,
        "flowChartMerge" => P::FlowChartMerge,
        "flowChartOfflineStorage" => P::FlowChartOfflineStorage,
        "flowChartOnlineStorage" => P::FlowChartOnlineStorage,
        "flowChartMagneticTape" => P::FlowChartMagneticTape,
        "flowChartMagneticDisk" => P::FlowChartMagneticDisk,
        "flowChartMagneticDrum" => P::FlowChartMagneticDrum,
        "flowChartDisplay" => P::FlowChartDisplay,
        "flowChartDelay" => P::FlowChartDelay,
        "flowChartAlternateProcess" => P::FlowChartAlternateProcess,
        "flowChartOffpageConnector" => P::FlowChartOffpageConnector,
        "actionButtonBlank" => P::ActionButtonBlank,
        "actionButtonHome" => P::ActionButtonHome,
        "actionButtonHelp" => P::ActionButtonHelp,
        "actionButtonInformation" => P::ActionButtonInformation,
        "actionButtonForwardNext" => P::ActionButtonForwardNext,
        "actionButtonBackPrevious" => P::ActionButtonBackPrevious,
        "actionButtonEnd" => P::ActionButtonEnd,
        "actionButtonBeginning" => P::ActionButtonBeginning,
        "actionButtonReturn" => P::ActionButtonReturn,
        "actionButtonDocument" => P::ActionButtonDocument,
        "actionButtonSound" => P::ActionButtonSound,
        "actionButtonMovie" => P::ActionButtonMovie,
        "irregularSeal1" => P::IrregularSeal1,
        "irregularSeal2" => P::IrregularSeal2,
        "wave" => P::Wave,
        "doubleWave" => P::DoubleWave,
        "ellipseRibbon" => P::EllipseRibbon,
        "ellipseRibbon2" => P::EllipseRibbon2,
        "verticalScroll" => P::VerticalScroll,
        "horizontalScroll" => P::HorizontalScroll,
        "leftRightRibbon" => P::LeftRightRibbon,
        "gear6" => P::Gear6,
        "gear9" => P::Gear9,
        "funnel" => P::Funnel,
        "mathPlus" => P::MathPlus,
        "mathMinus" => P::MathMinus,
        "mathMultiply" => P::MathMultiply,
        "mathDivide" => P::MathDivide,
        "mathEqual" => P::MathEqual,
        "mathNotEqual" => P::MathNotEqual,
        "cornerTabs" => P::CornerTabs,
        "squareTabs" => P::SquareTabs,
        "plaqueTabs" => P::PlaqueTabs,
        "chartX" => P::ChartX,
        "chartStar" => P::ChartStar,
        "chartPlus" => P::ChartPlus,
        "halfFrame" => P::HalfFrame,
        "corner" => P::Corner,
        "diagStripe" => P::DiagStripe,
        "nonIsoscelesTrapezoid" => P::NonIsoscelesTrapezoid,
        "heptagon" => P::Heptagon,
        "decagon" => P::Decagon,
        "dodecagon" => P::Dodecagon,
        "round1Rect" => P::Round1Rect,
        "round2SameRect" => P::Round2SameRect,
        "round2DiagRect" => P::Round2DiagRect,
        "snipRoundRect" => P::SnipRoundRect,
        "snip1Rect" => P::Snip1Rect,
        "snip2SameRect" => P::Snip2SameRect,
        "snip2DiagRect" => P::Snip2DiagRect,
        other => P::Other(other.to_string()),
    }
}

// ── bodyPr (§20.1.2.1.1) ─────────────────────────────────────────────────

#[derive(Debug, Deserialize, Default)]
pub struct BodyPrXml {
    #[serde(rename = "@rot", default)]
    pub rot: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@vert", default)]
    pub vert: Option<StTextVerticalType>,
    #[serde(rename = "@wrap", default)]
    pub wrap: Option<StTextWrappingType>,
    #[serde(rename = "@lIns", default)]
    pub l_ins: Option<Dimension<Emu>>,
    #[serde(rename = "@tIns", default)]
    pub t_ins: Option<Dimension<Emu>>,
    #[serde(rename = "@rIns", default)]
    pub r_ins: Option<Dimension<Emu>>,
    #[serde(rename = "@bIns", default)]
    pub b_ins: Option<Dimension<Emu>>,
    #[serde(rename = "@anchor", default)]
    pub anchor: Option<StTextAnchoringType>,
    // Auto-fit choice — at most one present.
    #[serde(rename = "noAutofit", default)]
    pub no_autofit: Option<super::fill::Empty>,
    #[serde(rename = "normAutofit", default)]
    pub norm_autofit: Option<super::fill::Empty>,
    #[serde(rename = "spAutoFit", default)]
    pub sp_autofit: Option<super::fill::Empty>,
}

impl From<BodyPrXml> for BodyProperties {
    fn from(x: BodyPrXml) -> Self {
        let auto_fit = if x.no_autofit.is_some() {
            Some(TextAutoFit::NoAutoFit)
        } else if x.norm_autofit.is_some() {
            Some(TextAutoFit::NormalAutoFit)
        } else {
            x.sp_autofit.map(|_| TextAutoFit::SpAutoFit)
        };
        Self {
            rotation: x.rot,
            vert: x.vert.map(Into::into),
            wrap: x.wrap.map(Into::into),
            left_inset: x.l_ins,
            top_inset: x.t_ins,
            right_inset: x.r_ins,
            bottom_inset: x.b_ins,
            anchor: x.anchor.map(Into::into),
            auto_fit,
        }
    }
}

// ── wps:wsp (Word Processing Shape) ───────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct WspXml {
    #[serde(rename = "cNvPr", default)]
    pub(crate) cnv_pr: Option<CNvPrXml>,
    #[serde(rename = "cNvSpPr", default)]
    pub(crate) cnv_sp_pr: Option<super::fill::Empty>,
    #[serde(rename = "cNvCnPr", default)]
    pub(crate) cnv_cn_pr: Option<super::fill::Empty>,
    #[serde(rename = "spPr", default)]
    pub(crate) sp_pr: Option<SpPrXml>,
    #[serde(rename = "bodyPr", default)]
    pub(crate) body_pr: Option<BodyPrXml>,
    #[serde(rename = "txbx", default)]
    pub(crate) txbx: Option<TxbxXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TxbxXml {
    #[serde(rename = "txbxContent", default)]
    pub(crate) content: Option<TxbxContentXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TxbxContentXml {
    #[serde(rename = "$value", default)]
    pub(crate) children: Vec<crate::docx::parse::body_schema::BlockChildXml>,
}

impl WspXml {
    /// Convert to the model; `ctx` drives drawing/pict embed resolution
    /// inside the text box content.
    pub(crate) fn into_model(
        self,
        ctx: &mut crate::docx::parse::body::ConvertCtx,
    ) -> WordProcessingShape {
        let txbx_content: Vec<Block> = self
            .txbx
            .and_then(|t| t.content)
            .map(|c| {
                let (blocks, _) = crate::docx::parse::body::convert_container(c.children, ctx);
                blocks
            })
            .unwrap_or_default();
        WordProcessingShape {
            cnv_pr: self.cnv_pr.map(Into::into),
            shape_properties: self.sp_pr.map(Into::into),
            body_pr: self.body_pr.map(Into::into),
            txbx_content,
        }
    }
}

// ── ST enums ──────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBlackWhiteMode {
    Auto,
    Black,
    BlackGray,
    BlackWhite,
    Clr,
    Gray,
    GrayWhite,
    Hidden,
    InvGray,
    LtGray,
    White,
}

impl From<StBlackWhiteMode> for BlackWhiteMode {
    fn from(s: StBlackWhiteMode) -> Self {
        match s {
            StBlackWhiteMode::Auto => Self::Auto,
            StBlackWhiteMode::Black => Self::Black,
            StBlackWhiteMode::BlackGray => Self::BlackGray,
            StBlackWhiteMode::BlackWhite => Self::BlackWhite,
            StBlackWhiteMode::Clr => Self::Clr,
            StBlackWhiteMode::Gray => Self::Gray,
            StBlackWhiteMode::GrayWhite => Self::GrayWhite,
            StBlackWhiteMode::Hidden => Self::Hidden,
            StBlackWhiteMode::InvGray => Self::InvGray,
            StBlackWhiteMode::LtGray => Self::LtGray,
            StBlackWhiteMode::White => Self::White,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTextVerticalType {
    Horz,
    Vert,
    Vert270,
    WordArtVert,
    EaVert,
    MongolianVert,
    WordArtVertRtl,
}

impl From<StTextVerticalType> for TextVerticalType {
    fn from(s: StTextVerticalType) -> Self {
        use StTextVerticalType as X;
        match s {
            X::Horz => Self::Horz,
            X::Vert => Self::Vert,
            X::Vert270 => Self::Vert270,
            X::WordArtVert => Self::WordArtVert,
            X::EaVert => Self::EaVert,
            X::MongolianVert => Self::MongolianVert,
            X::WordArtVertRtl => Self::WordArtVertRtl,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTextWrappingType {
    None,
    Square,
}

impl From<StTextWrappingType> for TextWrappingType {
    fn from(s: StTextWrappingType) -> Self {
        match s {
            StTextWrappingType::None => Self::None,
            StTextWrappingType::Square => Self::Square,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub enum StTextAnchoringType {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "just")]
    Just,
    #[serde(rename = "dist")]
    Dist,
}

impl From<StTextAnchoringType> for TextAnchoringType {
    fn from(s: StTextAnchoringType) -> Self {
        match s {
            StTextAnchoringType::T => Self::Top,
            StTextAnchoringType::Ctr => Self::Center,
            StTextAnchoringType::B => Self::Bottom,
            StTextAnchoringType::Just => Self::Justified,
            StTextAnchoringType::Dist => Self::Distributed,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{PathFillMode, PresetShapeType, ShapeGeometry, TextWrappingType};

    fn parse_sp_pr(xml: &str) -> ShapeProperties {
        let wrapped = format!(
            r#"<wrap xmlns:a="urn:a" xmlns:r="urn:r" xmlns:w="urn:w" xmlns:wps="urn:wps">{}</wrap>"#,
            xml
        );
        #[derive(Deserialize)]
        struct Wrap {
            #[serde(rename = "spPr")]
            sp: SpPrXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.sp.into()
    }

    #[test]
    fn empty_sp_pr() {
        let sp = parse_sp_pr(r#"<spPr/>"#);
        assert!(sp.geometry.is_none());
        assert!(sp.fill.is_none());
    }

    #[test]
    fn transform_with_offset_and_ext() {
        let sp = parse_sp_pr(
            r#"<spPr>
                <xfrm rot="5400000" flipH="1">
                    <off x="100" y="200"/>
                    <ext cx="914400" cy="457200"/>
                </xfrm>
            </spPr>"#,
        );
        let t = sp.transform.unwrap();
        assert_eq!(t.rotation.unwrap().raw(), 5_400_000);
        assert_eq!(t.flip_h, Some(true));
        assert_eq!(t.offset.unwrap().x.raw(), 100);
        assert_eq!(t.extent.unwrap().width.raw(), 914_400);
    }

    #[test]
    fn preset_geometry_rect() {
        let sp = parse_sp_pr(r#"<spPr><prstGeom prst="rect"/></spPr>"#);
        match sp.geometry {
            Some(ShapeGeometry::Preset(p)) => assert_eq!(p.preset, PresetShapeType::Rect),
            other => panic!("expected Preset(Rect), got {other:?}"),
        }
    }

    #[test]
    fn preset_with_adjust_values() {
        let sp = parse_sp_pr(
            r#"<spPr>
                <prstGeom prst="roundRect">
                    <avLst><gd name="adj" fmla="val 20000"/></avLst>
                </prstGeom>
            </spPr>"#,
        );
        match sp.geometry {
            Some(ShapeGeometry::Preset(p)) => {
                assert_eq!(p.preset, PresetShapeType::RoundRect);
                assert_eq!(p.adjust_values.len(), 1);
                assert_eq!(p.adjust_values[0].formula, "val 20000");
            }
            other => panic!("got {other:?}"),
        }
    }

    #[test]
    fn preset_unknown_falls_back_to_other() {
        let sp = parse_sp_pr(r#"<spPr><prstGeom prst="newShapeType2030"/></spPr>"#);
        match sp.geometry {
            Some(ShapeGeometry::Preset(p)) => match p.preset {
                PresetShapeType::Other(s) => assert_eq!(s, "newShapeType2030"),
                other => panic!("expected Other, got {other:?}"),
            },
            other => panic!("got {other:?}"),
        }
    }

    #[test]
    fn custom_geometry_routed() {
        let sp = parse_sp_pr(
            r#"<spPr>
                <custGeom>
                    <pathLst>
                        <path w="100" h="100">
                            <moveTo><pt x="0" y="0"/></moveTo>
                            <close/>
                        </path>
                    </pathLst>
                </custGeom>
            </spPr>"#,
        );
        match sp.geometry {
            Some(ShapeGeometry::Custom(c)) => {
                assert_eq!(c.paths.len(), 1);
                assert_eq!(c.paths[0].fill, PathFillMode::Norm);
            }
            other => panic!("got {other:?}"),
        }
    }

    #[test]
    fn solid_fill_inside_sp_pr() {
        let sp = parse_sp_pr(r#"<spPr><solidFill><srgbClr val="FF0000"/></solidFill></spPr>"#);
        assert!(matches!(sp.fill, Some(DrawingFill::Solid(_))));
    }

    #[test]
    fn no_fill_inside_sp_pr() {
        let sp = parse_sp_pr(r#"<spPr><noFill/></spPr>"#);
        assert!(matches!(sp.fill, Some(DrawingFill::None)));
    }

    #[test]
    fn outline_routes_through_stroke() {
        let sp = parse_sp_pr(r#"<spPr><ln w="9525"><prstDash val="dashDot"/></ln></spPr>"#);
        let o = sp.outline.unwrap();
        assert_eq!(o.width.unwrap().raw(), 9525);
    }

    #[test]
    fn effect_list_routes_through_effect() {
        let sp = parse_sp_pr(
            r#"<spPr><effectLst><outerShdw blurRad="0" dist="0" dir="0">
                <srgbClr val="000000"/>
            </outerShdw></effectLst></spPr>"#,
        );
        assert_eq!(sp.effect_list.unwrap().effects.len(), 1);
    }

    // ── Body properties ──

    fn parse_body_pr(xml: &str) -> BodyProperties {
        let wrapped = format!(r#"<wrap xmlns:a="urn:a">{}</wrap>"#, xml);
        #[derive(Deserialize)]
        struct Wrap {
            #[serde(rename = "bodyPr")]
            bp: BodyPrXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.bp.into()
    }

    #[test]
    fn body_pr_attrs_and_autofit() {
        let bp = parse_body_pr(
            r#"<bodyPr rot="0" vert="horz" wrap="square" anchor="ctr"
                     lIns="91440" tIns="45720" rIns="91440" bIns="45720">
                <normAutofit/>
            </bodyPr>"#,
        );
        assert_eq!(bp.wrap, Some(TextWrappingType::Square));
        assert_eq!(bp.anchor, Some(TextAnchoringType::Center));
        assert_eq!(bp.auto_fit, Some(TextAutoFit::NormalAutoFit));
        assert_eq!(bp.left_inset.unwrap().raw(), 91440);
    }

    #[test]
    fn body_pr_no_autofit() {
        let bp = parse_body_pr(r#"<bodyPr><noAutofit/></bodyPr>"#);
        assert_eq!(bp.auto_fit, Some(TextAutoFit::NoAutoFit));
    }

    // ── wsp ──

    #[test]
    fn wsp_with_txbx_empty() {
        let xml = r#"<wrap xmlns:wps="urn:wps" xmlns:w="urn:w" xmlns:a="urn:a" xmlns:r="urn:r">
            <wsp>
                <cNvPr id="1" name="Shape1"/>
                <spPr><prstGeom prst="rect"/></spPr>
                <bodyPr/>
                <txbx><txbxContent/></txbx>
            </wsp>
        </wrap>"#;
        #[derive(Deserialize)]
        struct Wrap {
            wsp: WspXml,
        }
        let w: Wrap = quick_xml::de::from_str(xml).unwrap();
        let mut ctx = crate::docx::parse::body::ConvertCtx::new();
        let wsp = w.wsp.into_model(&mut ctx);
        assert_eq!(wsp.cnv_pr.unwrap().name, "Shape1");
        assert!(wsp.txbx_content.is_empty());
    }

    #[test]
    fn wsp_txbx_with_paragraph() {
        let xml = r#"<wrap xmlns:wps="urn:wps" xmlns:w="urn:w" xmlns:a="urn:a" xmlns:r="urn:r">
            <wsp>
                <cNvPr id="2" name="Shape2"/>
                <spPr><prstGeom prst="rect"/></spPr>
                <bodyPr/>
                <txbx>
                    <txbxContent>
                        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    </txbxContent>
                </txbx>
            </wsp>
        </wrap>"#;
        #[derive(Deserialize)]
        struct Wrap {
            wsp: WspXml,
        }
        let w: Wrap = quick_xml::de::from_str(xml).unwrap();
        let mut ctx = crate::docx::parse::body::ConvertCtx::new();
        let wsp = w.wsp.into_model(&mut ctx);
        assert_eq!(wsp.txbx_content.len(), 1);
        match &wsp.txbx_content[0] {
            Block::Paragraph(_) => (),
            other => panic!("expected Paragraph, got {other:?}"),
        }
    }

    // ── Picture spPr wiring (now that shape schema exists) ──
    // (We don't re-wire picture.rs here; that's a follow-up edit.)
}
