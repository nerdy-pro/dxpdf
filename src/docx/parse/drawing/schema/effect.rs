//! DrawingML effect schema (§20.1.8.24 CT_EffectList).
//!
//! Flat `<a:effectLst>` with eight child effect types, preserving document
//! order. Consumes `DrawingColorXml` and `DrawingFillXml`.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::dimension::{Dimension, Emu, SixtieThousandthDeg, ThousandthPercent};
use crate::docx::model::{
    BlendMode, BlurEffect, DrawingColor, DrawingFill, Effect, EffectList, FillOverlayEffect,
    GlowEffect, InnerShadowEffect, OuterShadowEffect, PresetShadowEffect, PresetShadowVal,
    RectAlignment, ReflectionEffect, SoftEdgeEffect,
};

use super::color::DrawingColorXml;
use super::fill::{AttrBool, DrawingFillXml, StRectAlignment};

// ── Top-level effect list ─────────────────────────────────────────────────

#[derive(Debug, Deserialize, Default)]
pub struct EffectListXml {
    #[serde(rename = "$value", default)]
    pub effects: Vec<EffectXml>,
}

#[derive(Debug, Deserialize)]
pub enum EffectXml {
    #[serde(rename = "blur")]
    Blur(BlurXml),
    #[serde(rename = "reflection")]
    Reflection(ReflectionXml),
    #[serde(rename = "softEdge")]
    SoftEdge(SoftEdgeXml),
    #[serde(rename = "fillOverlay")]
    FillOverlay(FillOverlayXml),
    #[serde(rename = "glow")]
    Glow(GlowXml),
    #[serde(rename = "innerShdw")]
    InnerShdw(InnerShdwXml),
    #[serde(rename = "outerShdw")]
    OuterShdw(OuterShdwXml),
    #[serde(rename = "prstShdw")]
    PrstShdw(PrstShdwXml),
}

// ── Parameter-only effects ────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct BlurXml {
    #[serde(rename = "@rad", default)]
    pub rad: Option<Dimension<Emu>>,
    #[serde(rename = "@grow", default)]
    pub grow: Option<AttrBool>,
}

#[derive(Debug, Deserialize)]
pub struct SoftEdgeXml {
    #[serde(rename = "@rad", default)]
    pub rad: Option<Dimension<Emu>>,
}

#[derive(Debug, Deserialize)]
pub struct ReflectionXml {
    #[serde(rename = "@blurRad", default)]
    pub blur_rad: Option<Dimension<Emu>>,
    #[serde(rename = "@stA", default)]
    pub start_alpha: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@stPos", default)]
    pub start_pos: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@endA", default)]
    pub end_alpha: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@endPos", default)]
    pub end_pos: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@dist", default)]
    pub distance: Option<Dimension<Emu>>,
    #[serde(rename = "@dir", default)]
    pub direction: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@fadeDir", default)]
    pub fade_direction: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@sx", default)]
    pub sx: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@sy", default)]
    pub sy: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@kx", default)]
    pub kx: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@ky", default)]
    pub ky: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@algn", default)]
    pub algn: Option<StRectAlignment>,
    #[serde(rename = "@rotWithShape", default)]
    pub rot_with_shape: Option<AttrBool>,
}

// ── Effects with fill/color children ──────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct FillOverlayXml {
    #[serde(rename = "@blend")]
    pub blend: StBlendMode,
    #[serde(rename = "$value", default)]
    pub fill: Option<DrawingFillXml>,
}

#[derive(Debug, Deserialize)]
pub struct GlowXml {
    #[serde(rename = "@rad", default)]
    pub rad: Option<Dimension<Emu>>,
    #[serde(rename = "$value", default)]
    pub color: Option<DrawingColorXml>,
}

#[derive(Debug, Deserialize)]
pub struct InnerShdwXml {
    #[serde(rename = "@blurRad", default)]
    pub blur_rad: Option<Dimension<Emu>>,
    #[serde(rename = "@dist", default)]
    pub distance: Option<Dimension<Emu>>,
    #[serde(rename = "@dir", default)]
    pub direction: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "$value", default)]
    pub color: Option<DrawingColorXml>,
}

#[derive(Debug, Deserialize)]
pub struct OuterShdwXml {
    #[serde(rename = "@blurRad", default)]
    pub blur_rad: Option<Dimension<Emu>>,
    #[serde(rename = "@dist", default)]
    pub distance: Option<Dimension<Emu>>,
    #[serde(rename = "@dir", default)]
    pub direction: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@sx", default)]
    pub sx: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@sy", default)]
    pub sy: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@kx", default)]
    pub kx: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@ky", default)]
    pub ky: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@algn", default)]
    pub algn: Option<StRectAlignment>,
    #[serde(rename = "@rotWithShape", default)]
    pub rot_with_shape: Option<AttrBool>,
    #[serde(rename = "$value", default)]
    pub color: Option<DrawingColorXml>,
}

#[derive(Debug, Deserialize)]
pub struct PrstShdwXml {
    #[serde(rename = "@prst")]
    pub prst: StPresetShadowVal,
    #[serde(rename = "@dist", default)]
    pub distance: Option<Dimension<Emu>>,
    #[serde(rename = "@dir", default)]
    pub direction: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "$value", default)]
    pub color: Option<DrawingColorXml>,
}

// ── ST enums ──────────────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBlendMode {
    Over,
    Mult,
    Screen,
    Darken,
    Lighten,
}

impl From<StBlendMode> for BlendMode {
    fn from(s: StBlendMode) -> Self {
        match s {
            StBlendMode::Over => Self::Over,
            StBlendMode::Mult => Self::Mult,
            StBlendMode::Screen => Self::Screen,
            StBlendMode::Darken => Self::Darken,
            StBlendMode::Lighten => Self::Lighten,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPresetShadowVal {
    Shdw1,
    Shdw2,
    Shdw3,
    Shdw4,
    Shdw5,
    Shdw6,
    Shdw7,
    Shdw8,
    Shdw9,
    Shdw10,
    Shdw11,
    Shdw12,
    Shdw13,
    Shdw14,
    Shdw15,
    Shdw16,
    Shdw17,
    Shdw18,
    Shdw19,
    Shdw20,
}

impl From<StPresetShadowVal> for PresetShadowVal {
    fn from(s: StPresetShadowVal) -> Self {
        use StPresetShadowVal as X;
        match s {
            X::Shdw1 => Self::Shdw1,
            X::Shdw2 => Self::Shdw2,
            X::Shdw3 => Self::Shdw3,
            X::Shdw4 => Self::Shdw4,
            X::Shdw5 => Self::Shdw5,
            X::Shdw6 => Self::Shdw6,
            X::Shdw7 => Self::Shdw7,
            X::Shdw8 => Self::Shdw8,
            X::Shdw9 => Self::Shdw9,
            X::Shdw10 => Self::Shdw10,
            X::Shdw11 => Self::Shdw11,
            X::Shdw12 => Self::Shdw12,
            X::Shdw13 => Self::Shdw13,
            X::Shdw14 => Self::Shdw14,
            X::Shdw15 => Self::Shdw15,
            X::Shdw16 => Self::Shdw16,
            X::Shdw17 => Self::Shdw17,
            X::Shdw18 => Self::Shdw18,
            X::Shdw19 => Self::Shdw19,
            X::Shdw20 => Self::Shdw20,
        }
    }
}

// ── Conversion to model ───────────────────────────────────────────────────

impl From<EffectListXml> for EffectList {
    fn from(x: EffectListXml) -> Self {
        Self {
            effects: x
                .effects
                .into_iter()
                .filter_map(|e| e.try_into().ok())
                .collect(),
        }
    }
}

impl TryFrom<EffectXml> for Effect {
    type Error = ();
    fn try_from(e: EffectXml) -> Result<Self, Self::Error> {
        Ok(match e {
            EffectXml::Blur(b) => Effect::Blur(BlurEffect {
                radius: b.rad.unwrap_or_default(),
                grow: b.grow.map(|o| o.0),
            }),
            EffectXml::SoftEdge(s) => Effect::SoftEdge(SoftEdgeEffect {
                radius: s.rad.unwrap_or_default(),
            }),
            EffectXml::Reflection(r) => Effect::Reflection(reflection(r)),
            EffectXml::FillOverlay(f) => {
                // Spec requires a fill child; if missing, skip this effect.
                let fill: DrawingFill = f.fill.ok_or(())?.into();
                Effect::FillOverlay(FillOverlayEffect {
                    fill,
                    blend: f.blend.into(),
                })
            }
            EffectXml::Glow(g) => Effect::Glow(GlowEffect {
                radius: g.rad.unwrap_or_default(),
                color: g.color.ok_or(())?.into(),
            }),
            EffectXml::InnerShdw(i) => {
                let color: DrawingColor = i.color.ok_or(())?.into();
                Effect::InnerShdw(InnerShadowEffect {
                    blur_radius: i.blur_rad.unwrap_or_default(),
                    distance: i.distance.unwrap_or_default(),
                    direction: i.direction.unwrap_or_default(),
                    color,
                })
            }
            EffectXml::OuterShdw(o) => {
                let color: DrawingColor = o.color.ok_or(())?.into();
                Effect::OuterShdw(OuterShadowEffect {
                    blur_radius: o.blur_rad.unwrap_or_default(),
                    distance: o.distance.unwrap_or_default(),
                    direction: o.direction.unwrap_or_default(),
                    sx: o.sx.unwrap_or(Dimension::new(100_000)),
                    sy: o.sy.unwrap_or(Dimension::new(100_000)),
                    kx: o.kx.unwrap_or_default(),
                    ky: o.ky.unwrap_or_default(),
                    alignment: o.algn.map(Into::into).unwrap_or(RectAlignment::B),
                    rot_with_shape: o.rot_with_shape.map(|b| b.0),
                    color,
                })
            }
            EffectXml::PrstShdw(p) => {
                let color: DrawingColor = p.color.ok_or(())?.into();
                Effect::PrstShdw(PresetShadowEffect {
                    preset: p.prst.into(),
                    distance: p.distance.unwrap_or_default(),
                    direction: p.direction.unwrap_or_default(),
                    color,
                })
            }
        })
    }
}

fn reflection(r: ReflectionXml) -> ReflectionEffect {
    ReflectionEffect {
        blur_radius: r.blur_rad.unwrap_or_default(),
        start_alpha: r.start_alpha.unwrap_or(Dimension::new(100_000)),
        start_pos: r.start_pos.unwrap_or_default(),
        end_alpha: r.end_alpha.unwrap_or_default(),
        end_pos: r.end_pos.unwrap_or(Dimension::new(100_000)),
        distance: r.distance.unwrap_or_default(),
        direction: r.direction.unwrap_or_default(),
        fade_direction: r.fade_direction.unwrap_or(Dimension::new(5_400_000)),
        sx: r.sx.unwrap_or(Dimension::new(100_000)),
        sy: r.sy.unwrap_or(Dimension::new(100_000)),
        kx: r.kx.unwrap_or_default(),
        ky: r.ky.unwrap_or_default(),
        alignment: r.algn.map(Into::into).unwrap_or(RectAlignment::B),
        rot_with_shape: r.rot_with_shape.map(|b| b.0),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> EffectList {
        let wrapped = format!(
            r#"<wrap xmlns:a="urn:a" xmlns:r="urn:r"><effectLst>{}</effectLst></wrap>"#,
            xml
        );
        #[derive(Deserialize)]
        struct Wrap {
            #[serde(rename = "effectLst")]
            el: EffectListXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.el.into()
    }

    #[test]
    fn empty_effect_list() {
        let el = parse("");
        assert!(el.effects.is_empty());
    }

    #[test]
    fn blur_effect() {
        let el = parse(r#"<blur rad="50000" grow="1"/>"#);
        assert_eq!(el.effects.len(), 1);
        match &el.effects[0] {
            Effect::Blur(b) => {
                assert_eq!(b.radius.raw(), 50000);
                assert_eq!(b.grow, Some(true));
            }
            other => panic!("expected Blur, got {other:?}"),
        }
    }

    #[test]
    fn soft_edge_effect() {
        let el = parse(r#"<softEdge rad="25000"/>"#);
        match &el.effects[0] {
            Effect::SoftEdge(s) => assert_eq!(s.radius.raw(), 25000),
            other => panic!("expected SoftEdge, got {other:?}"),
        }
    }

    #[test]
    fn glow_with_color() {
        let el = parse(r#"<glow rad="40000"><srgbClr val="FF0000"/></glow>"#);
        match &el.effects[0] {
            Effect::Glow(g) => {
                assert_eq!(g.radius.raw(), 40000);
                assert!(matches!(
                    g.color,
                    crate::docx::model::DrawingColor::Srgb { rgb: 0xFF0000, .. }
                ));
            }
            other => panic!("expected Glow, got {other:?}"),
        }
    }

    #[test]
    fn outer_shadow_full() {
        let el = parse(
            r#"<outerShdw blurRad="50800" dist="38100" dir="2700000" sx="100000" sy="100000"
                         algn="tl" rotWithShape="0">
                <srgbClr val="000000"/>
            </outerShdw>"#,
        );
        match &el.effects[0] {
            Effect::OuterShdw(o) => {
                assert_eq!(o.blur_radius.raw(), 50800);
                assert_eq!(o.direction.raw(), 2_700_000);
                assert_eq!(o.alignment, RectAlignment::Tl);
                assert_eq!(o.rot_with_shape, Some(false));
            }
            other => panic!("expected OuterShdw, got {other:?}"),
        }
    }

    #[test]
    fn inner_shadow_with_scheme_color() {
        let el = parse(
            r#"<innerShdw blurRad="12700" dist="6350" dir="13500000">
                <schemeClr val="accent1"/>
            </innerShdw>"#,
        );
        match &el.effects[0] {
            Effect::InnerShdw(i) => {
                assert_eq!(i.blur_radius.raw(), 12700);
                assert_eq!(i.direction.raw(), 13_500_000);
                assert!(matches!(
                    i.color,
                    crate::docx::model::DrawingColor::Scheme { .. }
                ));
            }
            other => panic!("expected InnerShdw, got {other:?}"),
        }
    }

    #[test]
    fn preset_shadow_with_color() {
        let el = parse(r#"<prstShdw prst="shdw5" dist="10000"><srgbClr val="808080"/></prstShdw>"#);
        match &el.effects[0] {
            Effect::PrstShdw(p) => {
                assert_eq!(p.preset, PresetShadowVal::Shdw5);
                assert_eq!(p.distance.raw(), 10000);
            }
            other => panic!("expected PrstShdw, got {other:?}"),
        }
    }

    #[test]
    fn fill_overlay_screen_blend() {
        let el = parse(
            r#"<fillOverlay blend="screen"><solidFill><srgbClr val="FFFFFF"/></solidFill></fillOverlay>"#,
        );
        match &el.effects[0] {
            Effect::FillOverlay(f) => {
                assert_eq!(f.blend, BlendMode::Screen);
                assert!(matches!(f.fill, DrawingFill::Solid(_)));
            }
            other => panic!("expected FillOverlay, got {other:?}"),
        }
    }

    #[test]
    fn reflection_preserves_order() {
        let el = parse(
            r#"<reflection blurRad="0" stA="50000" stPos="0" endA="10000" endPos="100000"
                           dist="20000" dir="5400000" fadeDir="5400000"
                           sx="100000" sy="-100000" kx="0" ky="0" algn="bl" rotWithShape="0"/>"#,
        );
        match &el.effects[0] {
            Effect::Reflection(r) => {
                assert_eq!(r.start_alpha.raw(), 50000);
                assert_eq!(r.sy.raw(), -100_000);
                assert_eq!(r.alignment, RectAlignment::Bl);
            }
            other => panic!("expected Reflection, got {other:?}"),
        }
    }

    #[test]
    fn multiple_effects_in_order() {
        let el = parse(
            r#"<outerShdw blurRad="0" dist="0" dir="0"><srgbClr val="000000"/></outerShdw>
               <blur rad="5000"/>
               <softEdge rad="2500"/>"#,
        );
        assert_eq!(el.effects.len(), 3);
        assert!(matches!(&el.effects[0], Effect::OuterShdw(_)));
        assert!(matches!(&el.effects[1], Effect::Blur(_)));
        assert!(matches!(&el.effects[2], Effect::SoftEdge(_)));
    }
}
