//! DrawingML custom geometry schema (§20.1.9.8 CT_CustomGeometry2D).
//!
//! Path commands are mixed-content (moveTo / lnTo / cubicBezTo / quadBezTo /
//! arcTo / close) preserving document order via `$value` + untagged enum.
//! `AdjCoord` attributes are either decimal integers or guide-name
//! references, resolved by a custom `Deserialize` that tries integer first.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::{Deserialize, Deserializer};

use crate::docx::dimension::{Dimension, Emu};
use crate::docx::model::{
    AdjCoord, AdjPoint, AdjustHandle, ConnectionSite, CustomGeometry, GeomGuide, PathCommand,
    PathDef, PathFillMode, TextRect,
};

// ── CustomGeometry ────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, Default)]
pub struct CustomGeometryXml {
    #[serde(rename = "avLst", default)]
    pub av_lst: Option<GdListXml>,
    #[serde(rename = "gdLst", default)]
    pub gd_lst: Option<GdListXml>,
    #[serde(rename = "ahLst", default)]
    pub ah_lst: Option<AhListXml>,
    #[serde(rename = "cxnLst", default)]
    pub cxn_lst: Option<CxnListXml>,
    #[serde(rename = "rect", default)]
    pub rect: Option<TextRectXml>,
    #[serde(rename = "pathLst", default)]
    pub path_lst: Option<PathListXml>,
}

#[derive(Debug, Deserialize, Default)]
pub struct GdListXml {
    #[serde(rename = "gd", default)]
    pub guides: Vec<GeomGuideXml>,
}

#[derive(Debug, Deserialize)]
pub struct GeomGuideXml {
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@fmla")]
    pub fmla: String,
}

#[derive(Debug, Deserialize, Default)]
pub struct AhListXml {
    #[serde(rename = "$value", default)]
    pub handles: Vec<AhXml>,
}

#[derive(Debug, Deserialize)]
pub enum AhXml {
    #[serde(rename = "ahXY")]
    XY(AhXyXml),
    #[serde(rename = "ahPolar")]
    Polar(AhPolarXml),
}

#[derive(Debug, Deserialize)]
pub struct AhXyXml {
    #[serde(rename = "@gdRefX", default)]
    pub gd_ref_x: Option<String>,
    #[serde(rename = "@gdRefY", default)]
    pub gd_ref_y: Option<String>,
    #[serde(rename = "@minX", default)]
    pub min_x: Option<AdjCoordXml>,
    #[serde(rename = "@maxX", default)]
    pub max_x: Option<AdjCoordXml>,
    #[serde(rename = "@minY", default)]
    pub min_y: Option<AdjCoordXml>,
    #[serde(rename = "@maxY", default)]
    pub max_y: Option<AdjCoordXml>,
    #[serde(rename = "pos")]
    pub pos: AdjPointXml,
}

#[derive(Debug, Deserialize)]
pub struct AhPolarXml {
    #[serde(rename = "@gdRefR", default)]
    pub gd_ref_r: Option<String>,
    #[serde(rename = "@gdRefAng", default)]
    pub gd_ref_ang: Option<String>,
    #[serde(rename = "@minR", default)]
    pub min_r: Option<AdjCoordXml>,
    #[serde(rename = "@maxR", default)]
    pub max_r: Option<AdjCoordXml>,
    #[serde(rename = "@minAng", default)]
    pub min_ang: Option<AdjCoordXml>,
    #[serde(rename = "@maxAng", default)]
    pub max_ang: Option<AdjCoordXml>,
    #[serde(rename = "pos")]
    pub pos: AdjPointXml,
}

#[derive(Debug, Deserialize, Default)]
pub struct CxnListXml {
    #[serde(rename = "cxn", default)]
    pub sites: Vec<CxnXml>,
}

#[derive(Debug, Deserialize)]
pub struct CxnXml {
    #[serde(rename = "@ang")]
    pub ang: AdjCoordXml,
    #[serde(rename = "pos")]
    pub pos: AdjPointXml,
}

#[derive(Debug, Deserialize)]
pub struct TextRectXml {
    #[serde(rename = "@l")]
    pub l: AdjCoordXml,
    #[serde(rename = "@t")]
    pub t: AdjCoordXml,
    #[serde(rename = "@r")]
    pub r: AdjCoordXml,
    #[serde(rename = "@b")]
    pub b: AdjCoordXml,
}

#[derive(Debug, Deserialize, Default)]
pub struct PathListXml {
    #[serde(rename = "path", default)]
    pub paths: Vec<PathXml>,
}

#[derive(Debug, Deserialize, Default)]
pub struct PathXml {
    #[serde(rename = "@w", default)]
    pub w: Option<Dimension<Emu>>,
    #[serde(rename = "@h", default)]
    pub h: Option<Dimension<Emu>>,
    #[serde(rename = "@fill", default)]
    pub fill: Option<StPathFillMode>,
    /// Per §20.1.9.15, default is `true`.
    #[serde(rename = "@stroke", default)]
    pub stroke: Option<StrAttrBool>,
    /// Per §20.1.9.15, default is `false`.
    #[serde(rename = "@extrusionOk", default)]
    pub extrusion_ok: Option<StrAttrBool>,
    #[serde(rename = "$value", default)]
    pub commands: Vec<PathCommandXml>,
}

#[derive(Debug, Deserialize)]
pub enum PathCommandXml {
    #[serde(rename = "moveTo")]
    MoveTo(SinglePtXml),
    #[serde(rename = "lnTo")]
    LnTo(SinglePtXml),
    #[serde(rename = "cubicBezTo")]
    CubicBezTo(MultiPtXml),
    #[serde(rename = "quadBezTo")]
    QuadBezTo(MultiPtXml),
    #[serde(rename = "arcTo")]
    ArcTo(ArcToXml),
    #[serde(rename = "close")]
    Close(CloseXml),
}

#[derive(Debug, Deserialize)]
pub struct SinglePtXml {
    #[serde(rename = "pt")]
    pub pt: AdjPointXml,
}

#[derive(Debug, Deserialize)]
pub struct MultiPtXml {
    #[serde(rename = "pt", default)]
    pub pts: Vec<AdjPointXml>,
}

#[derive(Debug, Deserialize)]
pub struct ArcToXml {
    #[serde(rename = "@wR")]
    pub wr: AdjCoordXml,
    #[serde(rename = "@hR")]
    pub hr: AdjCoordXml,
    #[serde(rename = "@stAng")]
    pub st_ang: AdjCoordXml,
    #[serde(rename = "@swAng")]
    pub sw_ang: AdjCoordXml,
}

#[derive(Debug, Deserialize, Default)]
pub struct CloseXml {}

#[derive(Debug, Deserialize)]
pub struct AdjPointXml {
    #[serde(rename = "@x")]
    pub x: AdjCoordXml,
    #[serde(rename = "@y")]
    pub y: AdjCoordXml,
}

/// An `AdjCoord` attribute value: either a decimal integer literal or a
/// guide-name reference.
#[derive(Clone, Debug, PartialEq)]
pub struct AdjCoordXml(pub AdjCoord);

impl<'de> Deserialize<'de> for AdjCoordXml {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        if let Ok(n) = s.parse::<i64>() {
            return Ok(AdjCoordXml(AdjCoord::Lit(n)));
        }
        if s.is_empty() {
            return Err(serde::de::Error::custom(
                "empty AdjCoord: expected integer or guide name",
            ));
        }
        Ok(AdjCoordXml(AdjCoord::Guide(s)))
    }
}

/// Alias for the shared attribute-level boolean primitive.
pub use crate::docx::parse::primitives::AttrBool as StrAttrBool;

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPathFillMode {
    None,
    Norm,
    Lighten,
    LightenLess,
    Darken,
    DarkenLess,
}

impl From<StPathFillMode> for PathFillMode {
    fn from(s: StPathFillMode) -> Self {
        match s {
            StPathFillMode::None => Self::None,
            StPathFillMode::Norm => Self::Norm,
            StPathFillMode::Lighten => Self::Lighten,
            StPathFillMode::LightenLess => Self::LightenLess,
            StPathFillMode::Darken => Self::Darken,
            StPathFillMode::DarkenLess => Self::DarkenLess,
        }
    }
}

// ── Conversion to model ───────────────────────────────────────────────────

impl From<CustomGeometryXml> for CustomGeometry {
    fn from(x: CustomGeometryXml) -> Self {
        Self {
            av_list: x.av_lst.map(guides).unwrap_or_default(),
            gd_list: x.gd_lst.map(guides).unwrap_or_default(),
            ah_list: x
                .ah_lst
                .map(|a| a.handles.into_iter().map(Into::into).collect())
                .unwrap_or_default(),
            cxn_list: x
                .cxn_lst
                .map(|c| c.sites.into_iter().map(Into::into).collect())
                .unwrap_or_default(),
            rect: x.rect.map(Into::into),
            paths: x
                .path_lst
                .map(|p| p.paths.into_iter().map(Into::into).collect())
                .unwrap_or_default(),
        }
    }
}

fn guides(l: GdListXml) -> Vec<GeomGuide> {
    l.guides
        .into_iter()
        .map(|g| GeomGuide {
            name: g.name,
            formula: g.fmla,
        })
        .collect()
}

impl From<AdjPointXml> for AdjPoint {
    fn from(x: AdjPointXml) -> Self {
        Self { x: x.x.0, y: x.y.0 }
    }
}

impl From<AhXml> for AdjustHandle {
    fn from(h: AhXml) -> Self {
        match h {
            AhXml::XY(x) => Self::XY {
                guide_ref_x: x.gd_ref_x,
                guide_ref_y: x.gd_ref_y,
                min_x: x.min_x.map(|c| c.0),
                max_x: x.max_x.map(|c| c.0),
                min_y: x.min_y.map(|c| c.0),
                max_y: x.max_y.map(|c| c.0),
                position: x.pos.into(),
            },
            AhXml::Polar(p) => Self::Polar {
                guide_ref_r: p.gd_ref_r,
                guide_ref_ang: p.gd_ref_ang,
                min_r: p.min_r.map(|c| c.0),
                max_r: p.max_r.map(|c| c.0),
                min_ang: p.min_ang.map(|c| c.0),
                max_ang: p.max_ang.map(|c| c.0),
                position: p.pos.into(),
            },
        }
    }
}

impl From<CxnXml> for ConnectionSite {
    fn from(x: CxnXml) -> Self {
        Self {
            angle: x.ang.0,
            position: x.pos.into(),
        }
    }
}

impl From<TextRectXml> for TextRect {
    fn from(x: TextRectXml) -> Self {
        Self {
            left: x.l.0,
            top: x.t.0,
            right: x.r.0,
            bottom: x.b.0,
        }
    }
}

impl From<PathXml> for PathDef {
    fn from(x: PathXml) -> Self {
        Self {
            w: x.w.unwrap_or_default(),
            h: x.h.unwrap_or_default(),
            fill: x.fill.map(Into::into).unwrap_or(PathFillMode::Norm),
            stroke: x.stroke.map(|b| b.0).unwrap_or(true),
            extrusion_ok: x.extrusion_ok.map(|b| b.0).unwrap_or(false),
            commands: x.commands.into_iter().filter_map(to_command).collect(),
        }
    }
}

fn to_command(c: PathCommandXml) -> Option<PathCommand> {
    Some(match c {
        PathCommandXml::MoveTo(s) => PathCommand::MoveTo(s.pt.into()),
        PathCommandXml::LnTo(s) => PathCommand::LineTo(s.pt.into()),
        PathCommandXml::CubicBezTo(m) => {
            let mut it = m.pts.into_iter();
            let a = it.next()?.into();
            let b = it.next()?.into();
            let c = it.next()?.into();
            PathCommand::CubicBezTo(a, b, c)
        }
        PathCommandXml::QuadBezTo(m) => {
            let mut it = m.pts.into_iter();
            let a = it.next()?.into();
            let b = it.next()?.into();
            PathCommand::QuadBezTo(a, b)
        }
        PathCommandXml::ArcTo(a) => PathCommand::ArcTo {
            wr: a.wr.0,
            hr: a.hr.0,
            start_angle: a.st_ang.0,
            swing_angle: a.sw_ang.0,
        },
        PathCommandXml::Close(_) => PathCommand::Close,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> CustomGeometry {
        let wrapped = format!(r#"<wrap xmlns:a="urn:a">{}</wrap>"#, xml);
        #[derive(Deserialize)]
        struct Wrap {
            #[serde(rename = "custGeom")]
            cg: CustomGeometryXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        w.cg.into()
    }

    #[test]
    fn empty_custom_geometry() {
        let g = parse(r#"<custGeom/>"#);
        assert!(g.av_list.is_empty());
        assert!(g.paths.is_empty());
    }

    #[test]
    fn guides_av_and_gd() {
        let g = parse(
            r#"<custGeom>
                <avLst><gd name="adj1" fmla="val 25000"/></avLst>
                <gdLst>
                    <gd name="a" fmla="*/ w 1 2"/>
                    <gd name="b" fmla="+- h 0 0"/>
                </gdLst>
            </custGeom>"#,
        );
        assert_eq!(g.av_list.len(), 1);
        assert_eq!(g.av_list[0].name, "adj1");
        assert_eq!(g.gd_list.len(), 2);
        assert_eq!(g.gd_list[1].formula, "+- h 0 0");
    }

    #[test]
    fn adj_coord_literal_vs_guide() {
        let g = parse(
            r#"<custGeom>
                <pathLst>
                    <path w="100" h="100">
                        <moveTo><pt x="0" y="0"/></moveTo>
                        <lnTo><pt x="a" y="b"/></lnTo>
                    </path>
                </pathLst>
            </custGeom>"#,
        );
        assert_eq!(g.paths.len(), 1);
        match &g.paths[0].commands[0] {
            PathCommand::MoveTo(p) => {
                assert_eq!(p.x, AdjCoord::Lit(0));
                assert_eq!(p.y, AdjCoord::Lit(0));
            }
            other => panic!("expected MoveTo, got {other:?}"),
        }
        match &g.paths[0].commands[1] {
            PathCommand::LineTo(p) => {
                assert_eq!(p.x, AdjCoord::Guide("a".into()));
                assert_eq!(p.y, AdjCoord::Guide("b".into()));
            }
            other => panic!("expected LineTo, got {other:?}"),
        }
    }

    #[test]
    fn cubic_and_quad_beziers() {
        let g = parse(
            r#"<custGeom>
                <pathLst>
                    <path w="100" h="100">
                        <cubicBezTo>
                            <pt x="10" y="0"/>
                            <pt x="20" y="0"/>
                            <pt x="30" y="0"/>
                        </cubicBezTo>
                        <quadBezTo>
                            <pt x="40" y="5"/>
                            <pt x="50" y="0"/>
                        </quadBezTo>
                        <close/>
                    </path>
                </pathLst>
            </custGeom>"#,
        );
        let cmds = &g.paths[0].commands;
        assert_eq!(cmds.len(), 3);
        assert!(matches!(&cmds[0], PathCommand::CubicBezTo(_, _, _)));
        assert!(matches!(&cmds[1], PathCommand::QuadBezTo(_, _)));
        assert!(matches!(&cmds[2], PathCommand::Close));
    }

    #[test]
    fn arc_to_all_attrs() {
        let g = parse(
            r#"<custGeom>
                <pathLst>
                    <path w="100" h="100">
                        <arcTo wR="50" hR="50" stAng="0" swAng="21600000"/>
                    </path>
                </pathLst>
            </custGeom>"#,
        );
        match &g.paths[0].commands[0] {
            PathCommand::ArcTo {
                wr,
                hr,
                start_angle,
                swing_angle,
            } => {
                assert_eq!(*wr, AdjCoord::Lit(50));
                assert_eq!(*hr, AdjCoord::Lit(50));
                assert_eq!(*start_angle, AdjCoord::Lit(0));
                assert_eq!(*swing_angle, AdjCoord::Lit(21_600_000));
            }
            other => panic!("expected ArcTo, got {other:?}"),
        }
    }

    #[test]
    fn adjust_handle_xy() {
        let g = parse(
            r#"<custGeom>
                <ahLst>
                    <ahXY gdRefX="adj1" minX="0" maxX="100">
                        <pos x="adj1" y="0"/>
                    </ahXY>
                </ahLst>
            </custGeom>"#,
        );
        assert_eq!(g.ah_list.len(), 1);
        match &g.ah_list[0] {
            AdjustHandle::XY {
                guide_ref_x,
                min_x,
                max_x,
                position,
                ..
            } => {
                assert_eq!(guide_ref_x.as_deref(), Some("adj1"));
                assert_eq!(*min_x, Some(AdjCoord::Lit(0)));
                assert_eq!(*max_x, Some(AdjCoord::Lit(100)));
                assert_eq!(position.x, AdjCoord::Guide("adj1".into()));
            }
            other => panic!("expected XY, got {other:?}"),
        }
    }

    #[test]
    fn adjust_handle_polar() {
        let g = parse(
            r#"<custGeom>
                <ahLst>
                    <ahPolar gdRefR="adj2" minAng="0" maxAng="21600000">
                        <pos x="0" y="0"/>
                    </ahPolar>
                </ahLst>
            </custGeom>"#,
        );
        match &g.ah_list[0] {
            AdjustHandle::Polar {
                guide_ref_r,
                max_ang,
                ..
            } => {
                assert_eq!(guide_ref_r.as_deref(), Some("adj2"));
                assert_eq!(*max_ang, Some(AdjCoord::Lit(21_600_000)));
            }
            other => panic!("expected Polar, got {other:?}"),
        }
    }

    #[test]
    fn connection_site() {
        let g = parse(
            r#"<custGeom>
                <cxnLst>
                    <cxn ang="5400000"><pos x="50" y="0"/></cxn>
                </cxnLst>
            </custGeom>"#,
        );
        assert_eq!(g.cxn_list.len(), 1);
        assert_eq!(g.cxn_list[0].angle, AdjCoord::Lit(5_400_000));
        assert_eq!(g.cxn_list[0].position.x, AdjCoord::Lit(50));
    }

    #[test]
    fn text_rect() {
        let g = parse(
            r#"<custGeom>
                <rect l="0" t="0" r="w" b="h"/>
            </custGeom>"#,
        );
        let r = g.rect.unwrap();
        assert_eq!(r.left, AdjCoord::Lit(0));
        assert_eq!(r.right, AdjCoord::Guide("w".into()));
        assert_eq!(r.bottom, AdjCoord::Guide("h".into()));
    }

    #[test]
    fn path_attrs_with_fill_and_stroke() {
        let g = parse(
            r#"<custGeom>
                <pathLst>
                    <path w="100" h="100" fill="none" stroke="false" extrusionOk="1">
                        <moveTo><pt x="0" y="0"/></moveTo>
                    </path>
                </pathLst>
            </custGeom>"#,
        );
        let p = &g.paths[0];
        assert_eq!(p.w.raw(), 100);
        assert_eq!(p.h.raw(), 100);
        assert_eq!(p.fill, PathFillMode::None);
        assert!(!p.stroke);
        assert!(p.extrusion_ok);
    }

    #[test]
    fn path_defaults_fill_norm_stroke_true_extrusion_false() {
        let g = parse(
            r#"<custGeom>
                <pathLst>
                    <path w="100" h="100">
                        <close/>
                    </path>
                </pathLst>
            </custGeom>"#,
        );
        let p = &g.paths[0];
        assert_eq!(p.fill, PathFillMode::Norm);
        assert!(p.stroke);
        assert!(!p.extrusion_ok);
    }

    #[test]
    fn complex_path_preserves_order() {
        let g = parse(
            r#"<custGeom>
                <pathLst>
                    <path w="100" h="100">
                        <moveTo><pt x="0" y="0"/></moveTo>
                        <lnTo><pt x="50" y="0"/></lnTo>
                        <arcTo wR="10" hR="10" stAng="0" swAng="5400000"/>
                        <lnTo><pt x="0" y="100"/></lnTo>
                        <close/>
                    </path>
                </pathLst>
            </custGeom>"#,
        );
        let cmds = &g.paths[0].commands;
        assert_eq!(cmds.len(), 5);
        assert!(matches!(cmds[0], PathCommand::MoveTo(_)));
        assert!(matches!(cmds[1], PathCommand::LineTo(_)));
        assert!(matches!(cmds[2], PathCommand::ArcTo { .. }));
        assert!(matches!(cmds[3], PathCommand::LineTo(_)));
        assert!(matches!(cmds[4], PathCommand::Close));
    }
}
