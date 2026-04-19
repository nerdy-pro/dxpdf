//! DrawingML fill schema (§20.1.8 EG_FillProperties).
//!
//! Six-way choice: `noFill`, `solidFill`, `gradFill`, `blipFill`, `pattFill`,
//! `grpFill`. Each variant has its own attribute/child grammar. Consumes
//! `DrawingColorXml` from the color schema.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::dimension::{Dimension, Emu, SixtieThousandthDeg, ThousandthPercent};
use crate::docx::model::{
    Blip, BlipCompression, BlipFill, BlipFillKind, DrawingFill, GradientFill,
    GradientShadeProperties, GradientStop, PathShadeType, PatternFill, PresetPatternVal,
    RectAlignment, RelativeRect, StretchFill, TileFill, TileFlipMode,
};
pub use crate::docx::parse::primitives::AttrBool;

use super::color::DrawingColorXml;

// ── Top-level choice ──────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub enum DrawingFillXml {
    #[serde(rename = "noFill")]
    NoFill(Empty),
    #[serde(rename = "grpFill")]
    GrpFill(Empty),
    #[serde(rename = "solidFill")]
    SolidFill(SolidFillXml),
    #[serde(rename = "gradFill")]
    GradFill(GradFillXml),
    #[serde(rename = "blipFill")]
    BlipFill(BlipFillXml),
    #[serde(rename = "pattFill")]
    PattFill(PattFillXml),
}

#[derive(Debug, Deserialize, Default)]
pub struct Empty {}

impl From<DrawingFillXml> for DrawingFill {
    fn from(f: DrawingFillXml) -> Self {
        match f {
            DrawingFillXml::NoFill(_) => Self::None,
            DrawingFillXml::GrpFill(_) => Self::Group,
            DrawingFillXml::SolidFill(s) => match s.color {
                Some(c) => Self::Solid(c.into()),
                // Ill-formed solidFill with no color — degrade to None per
                // legacy parser.
                None => Self::None,
            },
            DrawingFillXml::GradFill(g) => Self::Gradient(g.into()),
            DrawingFillXml::BlipFill(b) => Self::Blip(b.into()),
            DrawingFillXml::PattFill(p) => Self::Pattern(p.into()),
        }
    }
}

// ── solidFill ──────────────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct SolidFillXml {
    #[serde(rename = "$value", default)]
    pub color: Option<DrawingColorXml>,
}

// ── gradFill ───────────────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct GradFillXml {
    #[serde(rename = "@flip", default)]
    pub flip: Option<StTileFlipMode>,
    #[serde(rename = "@rotWithShape", default)]
    pub rot_with_shape: Option<AttrBool>,
    #[serde(rename = "gsLst", default)]
    pub gs_lst: Option<GsLstXml>,
    #[serde(rename = "lin", default)]
    pub lin: Option<LinShadeXml>,
    #[serde(rename = "path", default)]
    pub path: Option<PathShadeXml>,
    #[serde(rename = "tileRect", default)]
    pub tile_rect: Option<RelativeRectXml>,
}

#[derive(Debug, Deserialize, Default)]
pub struct GsLstXml {
    #[serde(rename = "gs", default)]
    pub stops: Vec<GsXml>,
}

#[derive(Debug, Deserialize)]
pub struct GsXml {
    #[serde(rename = "@pos")]
    pub pos: Dimension<ThousandthPercent>,
    #[serde(rename = "$value")]
    pub color: DrawingColorXml,
}

#[derive(Debug, Deserialize)]
pub struct LinShadeXml {
    #[serde(rename = "@ang", default)]
    pub angle: Option<Dimension<SixtieThousandthDeg>>,
    #[serde(rename = "@scaled", default)]
    pub scaled: Option<AttrBool>,
}

#[derive(Debug, Deserialize)]
pub struct PathShadeXml {
    #[serde(rename = "@path")]
    pub path_type: StPathShadeType,
    #[serde(rename = "fillToRect", default)]
    pub fill_to_rect: Option<RelativeRectXml>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPathShadeType {
    Shape,
    Circle,
    Rect,
}

impl From<StPathShadeType> for PathShadeType {
    fn from(s: StPathShadeType) -> Self {
        match s {
            StPathShadeType::Shape => Self::Shape,
            StPathShadeType::Circle => Self::Circle,
            StPathShadeType::Rect => Self::Rect,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTileFlipMode {
    None,
    X,
    Y,
    Xy,
}

impl From<StTileFlipMode> for TileFlipMode {
    fn from(s: StTileFlipMode) -> Self {
        match s {
            StTileFlipMode::None => Self::None,
            StTileFlipMode::X => Self::X,
            StTileFlipMode::Y => Self::Y,
            StTileFlipMode::Xy => Self::Xy,
        }
    }
}

#[derive(Debug, Deserialize)]
pub struct RelativeRectXml {
    #[serde(rename = "@l", default)]
    pub l: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@t", default)]
    pub t: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@r", default)]
    pub r: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@b", default)]
    pub b: Option<Dimension<ThousandthPercent>>,
}

impl From<RelativeRectXml> for RelativeRect {
    fn from(x: RelativeRectXml) -> Self {
        Self {
            left: x.l,
            top: x.t,
            right: x.r,
            bottom: x.b,
        }
    }
}

impl From<GradFillXml> for GradientFill {
    fn from(x: GradFillXml) -> Self {
        let stops = x
            .gs_lst
            .map(|l| l.stops.into_iter().map(Into::into).collect())
            .unwrap_or_default();
        let shade_properties = if let Some(lin) = x.lin {
            GradientShadeProperties::Linear {
                angle: lin.angle.unwrap_or_default(),
                scaled: lin.scaled.map(|o| o.0),
            }
        } else if let Some(path) = x.path {
            GradientShadeProperties::Path {
                path_type: path.path_type.into(),
                fill_to_rect: path.fill_to_rect.map(Into::into),
            }
        } else {
            GradientShadeProperties::Linear {
                angle: Dimension::new(0),
                scaled: None,
            }
        };
        Self {
            stops,
            shade_properties,
            flip: x.flip.map(Into::into),
            rot_with_shape: x.rot_with_shape.map(|o| o.0),
            tile_rect: x.tile_rect.map(Into::into),
        }
    }
}

impl From<GsXml> for GradientStop {
    fn from(x: GsXml) -> Self {
        Self {
            position: x.pos,
            color: x.color.into(),
        }
    }
}

// ── blipFill ───────────────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct BlipFillXml {
    #[serde(rename = "@rotWithShape", default)]
    pub rot_with_shape: Option<AttrBool>,
    #[serde(rename = "@dpi", default)]
    pub dpi: Option<u32>,
    #[serde(rename = "blip", default)]
    pub blip: Option<BlipXml>,
    #[serde(rename = "srcRect", default)]
    pub src_rect: Option<RelativeRectXml>,
    #[serde(rename = "stretch", default)]
    pub stretch: Option<StretchXml>,
    #[serde(rename = "tile", default)]
    pub tile: Option<TileXml>,
}

#[derive(Debug, Deserialize)]
pub struct BlipXml {
    #[serde(rename = "@r:embed", alias = "@embed", default)]
    pub embed: Option<String>,
    #[serde(rename = "@r:link", alias = "@link", default)]
    pub link: Option<String>,
    #[serde(rename = "@cstate", default)]
    pub cstate: Option<StBlipCompression>,
}

#[derive(Debug, Deserialize, Default)]
pub struct StretchXml {
    #[serde(rename = "fillRect", default)]
    pub fill_rect: Option<RelativeRectXml>,
}

#[derive(Debug, Deserialize)]
pub struct TileXml {
    #[serde(rename = "@tx", default)]
    pub tx: Option<Dimension<Emu>>,
    #[serde(rename = "@ty", default)]
    pub ty: Option<Dimension<Emu>>,
    #[serde(rename = "@sx", default)]
    pub sx: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@sy", default)]
    pub sy: Option<Dimension<ThousandthPercent>>,
    #[serde(rename = "@flip", default)]
    pub flip: Option<StTileFlipMode>,
    #[serde(rename = "@algn", default)]
    pub algn: Option<StRectAlignment>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBlipCompression {
    Email,
    Hqprint,
    None,
    Print,
    Screen,
}

impl From<StBlipCompression> for BlipCompression {
    fn from(s: StBlipCompression) -> Self {
        match s {
            StBlipCompression::Email => Self::Email,
            StBlipCompression::Hqprint => Self::Hqprint,
            StBlipCompression::None => Self::None,
            StBlipCompression::Print => Self::Print,
            StBlipCompression::Screen => Self::Screen,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub enum StRectAlignment {
    #[serde(rename = "tl")]
    Tl,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "tr")]
    Tr,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "bl")]
    Bl,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "br")]
    Br,
}

impl From<StRectAlignment> for RectAlignment {
    fn from(s: StRectAlignment) -> Self {
        match s {
            StRectAlignment::Tl => Self::Tl,
            StRectAlignment::T => Self::T,
            StRectAlignment::Tr => Self::Tr,
            StRectAlignment::L => Self::L,
            StRectAlignment::Ctr => Self::Ctr,
            StRectAlignment::R => Self::R,
            StRectAlignment::Bl => Self::Bl,
            StRectAlignment::B => Self::B,
            StRectAlignment::Br => Self::Br,
        }
    }
}

impl From<BlipFillXml> for BlipFill {
    fn from(x: BlipFillXml) -> Self {
        use crate::docx::model::RelId;
        let blip = x.blip.map(|b| Blip {
            embed: b.embed.map(RelId::new),
            link: b.link.map(RelId::new),
            compression: b.cstate.map(Into::into),
        });
        let fill_kind = match (x.stretch, x.tile) {
            (Some(s), _) => BlipFillKind::Stretch(StretchFill {
                fill_rect: s.fill_rect.map(Into::into),
            }),
            (None, Some(t)) => BlipFillKind::Tile(TileFill {
                tx: t.tx,
                ty: t.ty,
                sx: t.sx,
                sy: t.sy,
                flip: t.flip.map(Into::into),
                alignment: t.algn.map(Into::into),
            }),
            (None, None) => BlipFillKind::Unspecified,
        };
        Self {
            rotate_with_shape: x.rot_with_shape.map(|o| o.0),
            dpi: x.dpi,
            blip,
            src_rect: x.src_rect.map(Into::into),
            fill_kind,
        }
    }
}

// ── pattFill ───────────────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
pub struct PattFillXml {
    #[serde(rename = "@prst")]
    pub prst: StPresetPatternVal,
    #[serde(rename = "fgClr", default)]
    pub fg_clr: Option<ColorParent>,
    #[serde(rename = "bgClr", default)]
    pub bg_clr: Option<ColorParent>,
}

/// `<a:fgClr>` / `<a:bgClr>` wrap a single color-choice child.
#[derive(Debug, Deserialize)]
pub struct ColorParent {
    #[serde(rename = "$value")]
    pub color: DrawingColorXml,
}

impl From<PattFillXml> for PatternFill {
    fn from(x: PattFillXml) -> Self {
        Self {
            preset: x.prst.into(),
            fg_color: x.fg_clr.map(|c| c.color.into()),
            bg_color: x.bg_clr.map(|c| c.color.into()),
        }
    }
}

// ── ST_PresetPatternVal (§20.1.10.50) ─────────────────────────────────────

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPresetPatternVal {
    Pct5,
    Pct10,
    Pct20,
    Pct25,
    Pct30,
    Pct40,
    Pct50,
    Pct60,
    Pct70,
    Pct75,
    Pct80,
    Pct90,
    Horz,
    Vert,
    LtHorz,
    LtVert,
    DkHorz,
    DkVert,
    NarHorz,
    NarVert,
    DashHorz,
    DashVert,
    Cross,
    DnDiag,
    UpDiag,
    LtDnDiag,
    LtUpDiag,
    DkDnDiag,
    DkUpDiag,
    WdDnDiag,
    WdUpDiag,
    DashDnDiag,
    DashUpDiag,
    DiagCross,
    SmCheck,
    LgCheck,
    SmGrid,
    LgGrid,
    DotGrid,
    SmConfetti,
    LgConfetti,
    HorzBrick,
    DiagBrick,
    SolidDmnd,
    OpenDmnd,
    DotDmnd,
    Plaid,
    Sphere,
    Weave,
    DivotShingle,
    Trellis,
    ZigZag,
    Wave,
}

impl From<StPresetPatternVal> for PresetPatternVal {
    fn from(s: StPresetPatternVal) -> Self {
        use StPresetPatternVal as X;
        match s {
            X::Pct5 => Self::Pct5,
            X::Pct10 => Self::Pct10,
            X::Pct20 => Self::Pct20,
            X::Pct25 => Self::Pct25,
            X::Pct30 => Self::Pct30,
            X::Pct40 => Self::Pct40,
            X::Pct50 => Self::Pct50,
            X::Pct60 => Self::Pct60,
            X::Pct70 => Self::Pct70,
            X::Pct75 => Self::Pct75,
            X::Pct80 => Self::Pct80,
            X::Pct90 => Self::Pct90,
            X::Horz => Self::Horz,
            X::Vert => Self::Vert,
            X::LtHorz => Self::LtHorz,
            X::LtVert => Self::LtVert,
            X::DkHorz => Self::DkHorz,
            X::DkVert => Self::DkVert,
            X::NarHorz => Self::NarHorz,
            X::NarVert => Self::NarVert,
            X::DashHorz => Self::DashHorz,
            X::DashVert => Self::DashVert,
            X::Cross => Self::Cross,
            X::DnDiag => Self::DnDiag,
            X::UpDiag => Self::UpDiag,
            X::LtDnDiag => Self::LtDnDiag,
            X::LtUpDiag => Self::LtUpDiag,
            X::DkDnDiag => Self::DkDnDiag,
            X::DkUpDiag => Self::DkUpDiag,
            X::WdDnDiag => Self::WdDnDiag,
            X::WdUpDiag => Self::WdUpDiag,
            X::DashDnDiag => Self::DashDnDiag,
            X::DashUpDiag => Self::DashUpDiag,
            X::DiagCross => Self::DiagCross,
            X::SmCheck => Self::SmCheck,
            X::LgCheck => Self::LgCheck,
            X::SmGrid => Self::SmGrid,
            X::LgGrid => Self::LgGrid,
            X::DotGrid => Self::DotGrid,
            X::SmConfetti => Self::SmConfetti,
            X::LgConfetti => Self::LgConfetti,
            X::HorzBrick => Self::HorzBrick,
            X::DiagBrick => Self::DiagBrick,
            X::SolidDmnd => Self::SolidDmnd,
            X::OpenDmnd => Self::OpenDmnd,
            X::DotDmnd => Self::DotDmnd,
            X::Plaid => Self::Plaid,
            X::Sphere => Self::Sphere,
            X::Weave => Self::Weave,
            X::DivotShingle => Self::DivotShingle,
            X::Trellis => Self::Trellis,
            X::ZigZag => Self::ZigZag,
            X::Wave => Self::Wave,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> DrawingFill {
        let wrapped = format!(r#"<wrap xmlns:a="urn:a" xmlns:r="urn:r">{}</wrap>"#, xml);
        #[derive(Deserialize)]
        struct Wrap {
            #[serde(rename = "$value")]
            f: DrawingFillXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.f.into()
    }

    #[test]
    fn no_fill() {
        assert!(matches!(parse(r#"<noFill/>"#), DrawingFill::None));
    }

    #[test]
    fn grp_fill() {
        assert!(matches!(parse(r#"<grpFill/>"#), DrawingFill::Group));
    }

    #[test]
    fn solid_fill_srgb() {
        match parse(r#"<solidFill><srgbClr val="4F81BD"/></solidFill>"#) {
            DrawingFill::Solid(crate::docx::model::DrawingColor::Srgb { rgb, .. }) => {
                assert_eq!(rgb, 0x4F81BD);
            }
            other => panic!("expected Solid(Srgb), got {other:?}"),
        }
    }

    #[test]
    fn solid_fill_scheme_with_lum_mod() {
        match parse(
            r#"<solidFill><schemeClr val="accent1"><lumMod val="60000"/></schemeClr></solidFill>"#,
        ) {
            DrawingFill::Solid(crate::docx::model::DrawingColor::Scheme { name, transforms }) => {
                assert_eq!(name, crate::docx::model::SchemeColorVal::Accent1);
                assert_eq!(transforms.len(), 1);
            }
            other => panic!("expected Solid(Scheme), got {other:?}"),
        }
    }

    #[test]
    fn grad_fill_linear_two_stops() {
        match parse(
            r#"<gradFill rotWithShape="1">
                <gsLst>
                    <gs pos="0"><srgbClr val="FFFFFF"/></gs>
                    <gs pos="100000"><srgbClr val="000000"/></gs>
                </gsLst>
                <lin ang="5400000" scaled="1"/>
            </gradFill>"#,
        ) {
            DrawingFill::Gradient(g) => {
                assert_eq!(g.stops.len(), 2);
                assert_eq!(g.stops[0].position.raw(), 0);
                assert_eq!(g.stops[1].position.raw(), 100000);
                match g.shade_properties {
                    GradientShadeProperties::Linear { angle, scaled } => {
                        assert_eq!(angle.raw(), 5_400_000);
                        assert_eq!(scaled, Some(true));
                    }
                    other => panic!("expected Linear, got {other:?}"),
                }
                assert_eq!(g.rot_with_shape, Some(true));
            }
            other => panic!("expected Gradient, got {other:?}"),
        }
    }

    #[test]
    fn grad_fill_path_circle() {
        match parse(
            r#"<gradFill>
                <gsLst><gs pos="0"><srgbClr val="FF0000"/></gs></gsLst>
                <path path="circle"/>
            </gradFill>"#,
        ) {
            DrawingFill::Gradient(g) => match g.shade_properties {
                GradientShadeProperties::Path { path_type, .. } => {
                    assert_eq!(path_type, PathShadeType::Circle);
                }
                other => panic!("expected Path, got {other:?}"),
            },
            other => panic!("expected Gradient, got {other:?}"),
        }
    }

    #[test]
    fn blip_fill_embed_stretch() {
        match parse(
            r#"<blipFill rotWithShape="1">
                <blip r:embed="rId1"/>
                <stretch><fillRect l="0" t="0" r="0" b="0"/></stretch>
            </blipFill>"#,
        ) {
            DrawingFill::Blip(b) => {
                assert_eq!(b.rotate_with_shape, Some(true));
                assert_eq!(
                    b.blip
                        .as_ref()
                        .and_then(|bl| bl.embed.as_ref())
                        .map(|r| r.as_str()),
                    Some("rId1")
                );
                assert!(matches!(b.fill_kind, BlipFillKind::Stretch(_)));
            }
            other => panic!("expected Blip, got {other:?}"),
        }
    }

    #[test]
    fn blip_fill_tile() {
        match parse(
            r#"<blipFill>
                <blip r:embed="rId2"/>
                <tile tx="0" ty="0" sx="100000" sy="100000" flip="xy" algn="ctr"/>
            </blipFill>"#,
        ) {
            DrawingFill::Blip(b) => match b.fill_kind {
                BlipFillKind::Tile(t) => {
                    assert_eq!(t.sx.map(|d| d.raw()), Some(100000));
                    assert_eq!(t.flip, Some(TileFlipMode::Xy));
                    assert_eq!(t.alignment, Some(RectAlignment::Ctr));
                }
                other => panic!("expected Tile, got {other:?}"),
            },
            other => panic!("expected Blip, got {other:?}"),
        }
    }

    #[test]
    fn pattern_fill_with_fg_bg() {
        match parse(
            r#"<pattFill prst="cross">
                <fgClr><srgbClr val="FF0000"/></fgClr>
                <bgClr><srgbClr val="00FF00"/></bgClr>
            </pattFill>"#,
        ) {
            DrawingFill::Pattern(p) => {
                assert_eq!(p.preset, PresetPatternVal::Cross);
                assert!(p.fg_color.is_some());
                assert!(p.bg_color.is_some());
            }
            other => panic!("expected Pattern, got {other:?}"),
        }
    }

    #[test]
    fn pattern_fill_dk_horz() {
        match parse(r#"<pattFill prst="dkHorz"/>"#) {
            DrawingFill::Pattern(p) => assert_eq!(p.preset, PresetPatternVal::DkHorz),
            other => panic!("expected Pattern, got {other:?}"),
        }
    }

    #[test]
    fn blip_compression_alias() {
        // Per spec, `cstate="hqprint"` is valid.
        let xml = r#"<blipFill><blip r:embed="rId1" cstate="hqprint"/></blipFill>"#;
        match parse(xml) {
            DrawingFill::Blip(b) => {
                assert_eq!(b.blip.unwrap().compression, Some(BlipCompression::Hqprint));
            }
            other => panic!("expected Blip, got {other:?}"),
        }
    }
}
