//! DrawingML outline schema (§20.1.2.2.24 CT_LineProperties).
//!
//! `<a:ln>` with optional attributes (width, cap, compound, alignment) and
//! optional children for fill, dash, join, and head/tail arrows. Consumes
//! `DrawingFillXml` from the fill schema.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::dimension::{Dimension, Emu, ThousandthPercent};
use crate::docx::model::{
    CompoundLine, DashStop, LineCap, LineDash, LineEnd, LineEndSize, LineEndType, LineJoin,
    Outline, PenAlignment, PresetLineDashVal,
};

use super::fill::DrawingFillXml;

/// `<a:ln …>` — outline (§20.1.2.2.24 CT_LineProperties).
#[derive(Debug, Deserialize)]
pub struct OutlineXml {
    #[serde(rename = "@w", default)]
    pub w: Option<Dimension<Emu>>,
    #[serde(rename = "@cap", default)]
    pub cap: Option<StLineCap>,
    #[serde(rename = "@cmpd", default)]
    pub cmpd: Option<StCompoundLine>,
    #[serde(rename = "@algn", default)]
    pub algn: Option<StPenAlignment>,
    #[serde(rename = "$value", default)]
    pub children: Vec<LnChildXml>,
}

/// Each child of `<a:ln>` is either a fill choice, a dash choice, a join
/// choice, or a head/tail end spec.
#[derive(Debug, Deserialize)]
pub enum LnChildXml {
    // Fill choices — forward to DrawingFillXml.
    #[serde(rename = "noFill")]
    NoFill(super::fill::Empty),
    #[serde(rename = "grpFill")]
    GrpFill(super::fill::Empty),
    #[serde(rename = "solidFill")]
    SolidFill(super::fill::SolidFillXml),
    #[serde(rename = "gradFill")]
    GradFill(super::fill::GradFillXml),
    #[serde(rename = "blipFill")]
    BlipFill(super::fill::BlipFillXml),
    #[serde(rename = "pattFill")]
    PattFill(super::fill::PattFillXml),

    // Dash choices.
    #[serde(rename = "prstDash")]
    PrstDash(PrstDashXml),
    #[serde(rename = "custDash")]
    CustDash(CustDashXml),

    // Join choices (parameterless unit variants for round/bevel/miter).
    #[serde(rename = "round")]
    Round(super::fill::Empty),
    #[serde(rename = "bevel")]
    Bevel(super::fill::Empty),
    #[serde(rename = "miter")]
    Miter(MiterXml),

    // Head/tail arrows.
    #[serde(rename = "headEnd")]
    HeadEnd(LineEndXml),
    #[serde(rename = "tailEnd")]
    TailEnd(LineEndXml),
}

#[derive(Debug, Deserialize)]
pub struct PrstDashXml {
    #[serde(rename = "@val")]
    pub val: StPresetLineDashVal,
}

#[derive(Debug, Deserialize, Default)]
pub struct CustDashXml {
    #[serde(rename = "ds", default)]
    pub stops: Vec<DsXml>,
}

#[derive(Debug, Deserialize)]
pub struct DsXml {
    #[serde(rename = "@d")]
    pub d: Dimension<ThousandthPercent>,
    #[serde(rename = "@sp")]
    pub sp: Dimension<ThousandthPercent>,
}

#[derive(Debug, Deserialize)]
pub struct MiterXml {
    #[serde(rename = "@lim", default)]
    pub lim: Option<Dimension<ThousandthPercent>>,
}

#[derive(Debug, Deserialize)]
pub struct LineEndXml {
    #[serde(rename = "@type", default)]
    pub ty: Option<StLineEndType>,
    #[serde(rename = "@w", default)]
    pub w: Option<StLineEndSize>,
    #[serde(rename = "@len", default)]
    pub len: Option<StLineEndSize>,
}

// ── ST enums ──────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, Deserialize)]
pub enum StLineCap {
    #[serde(rename = "flat")]
    Flat,
    #[serde(rename = "rnd")]
    Rnd,
    #[serde(rename = "sq")]
    Sq,
}

impl From<StLineCap> for LineCap {
    fn from(s: StLineCap) -> Self {
        match s {
            StLineCap::Flat => Self::Flat,
            StLineCap::Rnd => Self::Round,
            StLineCap::Sq => Self::Square,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub enum StCompoundLine {
    #[serde(rename = "sng")]
    Sng,
    #[serde(rename = "dbl")]
    Dbl,
    #[serde(rename = "thickThin")]
    ThickThin,
    #[serde(rename = "thinThick")]
    ThinThick,
    #[serde(rename = "tri")]
    Tri,
}

impl From<StCompoundLine> for CompoundLine {
    fn from(s: StCompoundLine) -> Self {
        match s {
            StCompoundLine::Sng => Self::Single,
            StCompoundLine::Dbl => Self::Double,
            StCompoundLine::ThickThin => Self::ThickThin,
            StCompoundLine::ThinThick => Self::ThinThick,
            StCompoundLine::Tri => Self::Triple,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub enum StPenAlignment {
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "in")]
    In,
}

impl From<StPenAlignment> for PenAlignment {
    fn from(s: StPenAlignment) -> Self {
        match s {
            StPenAlignment::Ctr => Self::Center,
            StPenAlignment::In => Self::Inset,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPresetLineDashVal {
    Solid,
    Dot,
    Dash,
    LgDash,
    DashDot,
    LgDashDot,
    LgDashDotDot,
    SysDash,
    SysDot,
    SysDashDot,
    SysDashDotDot,
}

impl From<StPresetLineDashVal> for PresetLineDashVal {
    fn from(s: StPresetLineDashVal) -> Self {
        use StPresetLineDashVal as X;
        match s {
            X::Solid => Self::Solid,
            X::Dot => Self::Dot,
            X::Dash => Self::Dash,
            X::LgDash => Self::LgDash,
            X::DashDot => Self::DashDot,
            X::LgDashDot => Self::LgDashDot,
            X::LgDashDotDot => Self::LgDashDotDot,
            X::SysDash => Self::SysDash,
            X::SysDot => Self::SysDot,
            X::SysDashDot => Self::SysDashDot,
            X::SysDashDotDot => Self::SysDashDotDot,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StLineEndType {
    None,
    Triangle,
    Stealth,
    Diamond,
    Oval,
    Arrow,
}

impl From<StLineEndType> for LineEndType {
    fn from(s: StLineEndType) -> Self {
        match s {
            StLineEndType::None => Self::None,
            StLineEndType::Triangle => Self::Triangle,
            StLineEndType::Stealth => Self::Stealth,
            StLineEndType::Diamond => Self::Diamond,
            StLineEndType::Oval => Self::Oval,
            StLineEndType::Arrow => Self::Arrow,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub enum StLineEndSize {
    #[serde(rename = "sm")]
    Sm,
    #[serde(rename = "med")]
    Med,
    #[serde(rename = "lg")]
    Lg,
}

impl From<StLineEndSize> for LineEndSize {
    fn from(s: StLineEndSize) -> Self {
        match s {
            StLineEndSize::Sm => Self::Sm,
            StLineEndSize::Med => Self::Med,
            StLineEndSize::Lg => Self::Lg,
        }
    }
}

// ── Conversion to model ───────────────────────────────────────────────────

impl From<OutlineXml> for Outline {
    fn from(x: OutlineXml) -> Self {
        let mut fill = None;
        let mut dash = None;
        let mut join = None;
        let mut head_end = None;
        let mut tail_end = None;

        for child in x.children {
            match child {
                LnChildXml::NoFill(_)
                | LnChildXml::GrpFill(_)
                | LnChildXml::SolidFill(_)
                | LnChildXml::GradFill(_)
                | LnChildXml::BlipFill(_)
                | LnChildXml::PattFill(_) => {
                    // Wrap the raw variant into a DrawingFillXml for conversion.
                    let f = match child {
                        LnChildXml::NoFill(e) => DrawingFillXml::NoFill(e),
                        LnChildXml::GrpFill(e) => DrawingFillXml::GrpFill(e),
                        LnChildXml::SolidFill(s) => DrawingFillXml::SolidFill(s),
                        LnChildXml::GradFill(g) => DrawingFillXml::GradFill(g),
                        LnChildXml::BlipFill(b) => DrawingFillXml::BlipFill(b),
                        LnChildXml::PattFill(p) => DrawingFillXml::PattFill(p),
                        _ => unreachable!(),
                    };
                    fill = Some(f.into());
                }
                LnChildXml::PrstDash(p) => {
                    dash = Some(LineDash::Preset(p.val.into()));
                }
                LnChildXml::CustDash(c) => {
                    dash = Some(LineDash::Custom(
                        c.stops
                            .into_iter()
                            .map(|d| DashStop {
                                dash: d.d,
                                space: d.sp,
                            })
                            .collect(),
                    ));
                }
                LnChildXml::Round(_) => join = Some(LineJoin::Round),
                LnChildXml::Bevel(_) => join = Some(LineJoin::Bevel),
                LnChildXml::Miter(m) => {
                    join = Some(LineJoin::Miter { limit: m.lim });
                }
                LnChildXml::HeadEnd(le) => {
                    head_end = Some(line_end(le));
                }
                LnChildXml::TailEnd(le) => {
                    tail_end = Some(line_end(le));
                }
            }
        }

        Self {
            width: x.w,
            cap: x.cap.map(Into::into),
            compound: x.cmpd.map(Into::into),
            alignment: x.algn.map(Into::into),
            fill,
            dash,
            join,
            head_end,
            tail_end,
        }
    }
}

fn line_end(le: LineEndXml) -> LineEnd {
    LineEnd {
        kind: le.ty.map(Into::into).unwrap_or(LineEndType::None),
        width: le.w.map(Into::into).unwrap_or(LineEndSize::Med),
        length: le.len.map(Into::into).unwrap_or(LineEndSize::Med),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::DrawingFill;

    fn parse(xml: &str) -> Outline {
        let wrapped = format!(r#"<wrap xmlns:a="urn:a" xmlns:r="urn:r">{}</wrap>"#, xml);
        #[derive(Deserialize)]
        struct Wrap {
            ln: OutlineXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.ln.into()
    }

    #[test]
    fn empty_outline() {
        let o = parse(r#"<ln/>"#);
        assert!(o.width.is_none());
        assert!(o.fill.is_none());
        assert!(o.dash.is_none());
    }

    #[test]
    fn width_and_cap() {
        let o = parse(r#"<ln w="12700" cap="flat" algn="ctr" cmpd="dbl"/>"#);
        assert_eq!(o.width.unwrap().raw(), 12700);
        assert_eq!(o.cap, Some(LineCap::Flat));
        assert_eq!(o.compound, Some(CompoundLine::Double));
        assert_eq!(o.alignment, Some(PenAlignment::Center));
    }

    #[test]
    fn solid_fill_child() {
        let o = parse(r#"<ln><solidFill><srgbClr val="FF0000"/></solidFill></ln>"#);
        match o.fill {
            Some(DrawingFill::Solid(_)) => (),
            other => panic!("expected Solid fill, got {other:?}"),
        }
    }

    #[test]
    fn preset_dash() {
        let o = parse(r#"<ln><prstDash val="dashDot"/></ln>"#);
        match o.dash {
            Some(LineDash::Preset(PresetLineDashVal::DashDot)) => (),
            other => panic!("expected Preset(DashDot), got {other:?}"),
        }
    }

    #[test]
    fn custom_dash_stops() {
        let o = parse(
            r#"<ln><custDash>
                <ds d="100000" sp="50000"/>
                <ds d="200000" sp="50000"/>
            </custDash></ln>"#,
        );
        match o.dash {
            Some(LineDash::Custom(stops)) => {
                assert_eq!(stops.len(), 2);
                assert_eq!(stops[0].dash.raw(), 100000);
                assert_eq!(stops[0].space.raw(), 50000);
            }
            other => panic!("expected Custom dash, got {other:?}"),
        }
    }

    #[test]
    fn join_round_bevel_miter() {
        assert!(matches!(
            parse(r#"<ln><round/></ln>"#).join,
            Some(LineJoin::Round)
        ));
        assert!(matches!(
            parse(r#"<ln><bevel/></ln>"#).join,
            Some(LineJoin::Bevel)
        ));
        match parse(r#"<ln><miter lim="800000"/></ln>"#).join {
            Some(LineJoin::Miter { limit: Some(d) }) => assert_eq!(d.raw(), 800000),
            other => panic!("expected Miter, got {other:?}"),
        }
    }

    #[test]
    fn head_end_arrow() {
        let o = parse(r#"<ln><headEnd type="arrow" w="lg" len="sm"/></ln>"#);
        let h = o.head_end.unwrap();
        assert_eq!(h.kind, LineEndType::Arrow);
        assert_eq!(h.width, LineEndSize::Lg);
        assert_eq!(h.length, LineEndSize::Sm);
    }

    #[test]
    fn tail_end_defaults_when_attrs_missing() {
        let o = parse(r#"<ln><tailEnd/></ln>"#);
        let t = o.tail_end.unwrap();
        assert_eq!(t.kind, LineEndType::None);
        assert_eq!(t.width, LineEndSize::Med);
        assert_eq!(t.length, LineEndSize::Med);
    }

    #[test]
    fn full_outline_end_to_end() {
        let o = parse(
            r#"<ln w="9525" cap="rnd" cmpd="sng" algn="in">
                <solidFill><schemeClr val="accent1"/></solidFill>
                <prstDash val="solid"/>
                <round/>
                <tailEnd type="triangle" w="med" len="med"/>
            </ln>"#,
        );
        assert_eq!(o.width.unwrap().raw(), 9525);
        assert_eq!(o.cap, Some(LineCap::Round));
        assert!(matches!(o.fill, Some(DrawingFill::Solid(_))));
        assert!(matches!(
            o.dash,
            Some(LineDash::Preset(PresetLineDashVal::Solid))
        ));
        assert_eq!(o.join, Some(LineJoin::Round));
        assert_eq!(o.tail_end.unwrap().kind, LineEndType::Triangle);
    }
}
