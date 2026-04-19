//! Serde schema for VML (§14.1) — `<w:pict>` picture containers with shape
//! types, shapes, text boxes, and related child elements.
//!
//! XML structure goes through serde. Attribute-level string sub-grammars
//! (CSS-like `style`, `path` commands, `adj` / `coordsize` numeric lists,
//! color names/hex) are reused from `vml::{style, color, path_commands,
//! formulas}` as pure `&str` → value parsers.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::model::{
    Block, Pict, RelId, VmlConnectType, VmlDashStyle, VmlExtHandling, VmlFormula, VmlImageData,
    VmlJoinStyle, VmlLock, VmlPath, VmlShape, VmlShapeId, VmlShapeType, VmlStroke, VmlTextBox,
    VmlTextBoxInset, VmlVector2D, VmlWrap, VmlWrapSide, VmlWrapType,
};
use crate::docx::parse::body_schema::BlockChildXml;

use super::color::parse_color;
use super::formulas::parse_formula;
use super::path_commands::{parse_adj, parse_path_commands, parse_vector2d};
use super::style::{parse_length, parse_style};

// ── pict ──────────────────────────────────────────────────────────────────

/// `<w:pict>` — legacy VML picture container.
#[derive(Deserialize)]
pub(crate) struct PictXml {
    #[serde(rename = "shapetype", default)]
    pub shape_type: Option<ShapeTypeXml>,
    #[serde(rename = "shape", default)]
    pub shapes: Vec<ShapeXml>,
}

impl PictXml {
    /// Convert to the model; `ctx` resolves embedded drawings nested inside
    /// VML text box content (same iterator threading as shape).
    pub(crate) fn into_model(self, ctx: &mut crate::docx::parse::body::ConvertCtx) -> Pict {
        Pict {
            shape_type: self.shape_type.map(Into::into),
            shapes: self.shapes.into_iter().map(|s| s.into_model(ctx)).collect(),
        }
    }
}

// ── shapetype ─────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct ShapeTypeXml {
    #[serde(rename = "@id", default)]
    pub id: Option<String>,
    #[serde(rename = "@coordsize", default)]
    pub coordsize: Option<String>,
    #[serde(rename = "@spt", default)]
    pub spt: Option<String>,
    #[serde(rename = "@adj", default)]
    pub adj: Option<String>,
    #[serde(rename = "@path", default)]
    pub path: Option<String>,
    #[serde(rename = "@filled", default)]
    pub filled: Option<VmlBool>,
    #[serde(rename = "@stroked", default)]
    pub stroked: Option<VmlBool>,

    #[serde(rename = "stroke", default)]
    pub stroke: Option<StrokeXml>,
    #[serde(rename = "path", default)]
    pub vml_path: Option<PathXml>,
    #[serde(rename = "formulas", default)]
    pub formulas: Option<FormulasXml>,
    #[serde(rename = "lock", default)]
    pub lock: Option<LockXml>,
}

impl From<ShapeTypeXml> for VmlShapeType {
    fn from(x: ShapeTypeXml) -> Self {
        Self {
            id: x.id.map(VmlShapeId::new),
            coord_size: parse_vector2d(x.coordsize),
            spt: x.spt.and_then(|s| s.parse::<f32>().ok()),
            adj: parse_adj(x.adj),
            path: parse_path_commands(x.path),
            filled: x.filled.map(|b| b.0),
            stroked: x.stroked.map(|b| b.0),
            stroke: x.stroke.map(Into::into),
            vml_path: x.vml_path.map(Into::into),
            formulas: x
                .formulas
                .map(|f| {
                    f.entries
                        .into_iter()
                        .filter_map(|e| parse_formula(&e.eqn))
                        .collect()
                })
                .unwrap_or_default(),
            lock: x.lock.map(Into::into),
        }
    }
}

// ── shape ─────────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct ShapeXml {
    #[serde(rename = "@id", default)]
    pub id: Option<String>,
    /// Reference like `"#_x0000_t202"` — strip the leading `#`.
    #[serde(rename = "@type", default)]
    pub ty: Option<String>,
    #[serde(rename = "@style", default)]
    pub style: Option<String>,
    #[serde(rename = "@fillcolor", default)]
    pub fillcolor: Option<String>,
    #[serde(rename = "@stroked", default)]
    pub stroked: Option<VmlBool>,

    #[serde(rename = "stroke", default)]
    pub stroke: Option<StrokeXml>,
    #[serde(rename = "path", default)]
    pub vml_path: Option<PathXml>,
    #[serde(rename = "textbox", default)]
    pub textbox: Option<TextBoxXml>,
    #[serde(rename = "wrap", default)]
    pub wrap: Option<WrapXml>,
    #[serde(rename = "imagedata", default)]
    pub imagedata: Option<ImageDataXml>,
}

impl ShapeXml {
    fn into_model(self, ctx: &mut crate::docx::parse::body::ConvertCtx) -> VmlShape {
        VmlShape {
            id: self.id.map(VmlShapeId::new),
            shape_type_ref: self
                .ty
                .map(|s| VmlShapeId::new(s.strip_prefix('#').unwrap_or(&s))),
            style: parse_style(self.style),
            fill_color: self.fillcolor.as_deref().and_then(|s| parse_color(s).ok()),
            stroked: self.stroked.map(|b| b.0),
            stroke: self.stroke.map(Into::into),
            vml_path: self.vml_path.map(Into::into),
            text_box: self.textbox.map(|t| t.into_model(ctx)),
            wrap: self.wrap.map(Into::into),
            image_data: self.imagedata.map(Into::into),
        }
    }
}

// ── textbox ───────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct TextBoxXml {
    #[serde(rename = "@style", default)]
    pub style: Option<String>,
    #[serde(rename = "@inset", default)]
    pub inset: Option<String>,
    #[serde(rename = "txbxContent", default)]
    pub content: Option<TxbxContentXml>,
}

#[derive(Deserialize, Default)]
pub(crate) struct TxbxContentXml {
    #[serde(rename = "$value", default)]
    pub children: Vec<BlockChildXml>,
}

impl TextBoxXml {
    fn into_model(self, ctx: &mut crate::docx::parse::body::ConvertCtx) -> VmlTextBox {
        let content: Vec<Block> = self
            .content
            .map(|c| {
                let (blocks, _) = crate::docx::parse::body::convert_container(c.children, ctx);
                blocks
            })
            .unwrap_or_default();
        VmlTextBox {
            style: parse_style(self.style),
            inset: self.inset.and_then(parse_inset),
            content,
        }
    }
}

/// Parse comma-separated `left,top,right,bottom` CSS lengths.
fn parse_inset(s: String) -> Option<VmlTextBoxInset> {
    let parts: Vec<&str> = s.split(',').collect();
    Some(VmlTextBoxInset {
        left: parts.first().and_then(|v| parse_length(v)),
        top: parts.get(1).and_then(|v| parse_length(v)),
        right: parts.get(2).and_then(|v| parse_length(v)),
        bottom: parts.get(3).and_then(|v| parse_length(v)),
    })
}

// ── stroke ────────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct StrokeXml {
    #[serde(rename = "@dashstyle", default)]
    pub dashstyle: Option<String>,
    #[serde(rename = "@joinstyle", default)]
    pub joinstyle: Option<String>,
}

impl From<StrokeXml> for VmlStroke {
    fn from(x: StrokeXml) -> Self {
        Self {
            dash_style: x.dashstyle.as_deref().and_then(parse_dash_style),
            join_style: x.joinstyle.as_deref().and_then(parse_join_style),
        }
    }
}

fn parse_dash_style(s: &str) -> Option<VmlDashStyle> {
    match s {
        "solid" => Some(VmlDashStyle::Solid),
        "shortdash" => Some(VmlDashStyle::ShortDash),
        "shortdot" => Some(VmlDashStyle::ShortDot),
        "shortdashdot" => Some(VmlDashStyle::ShortDashDot),
        "shortdashdotdot" => Some(VmlDashStyle::ShortDashDotDot),
        "dot" => Some(VmlDashStyle::Dot),
        "dash" => Some(VmlDashStyle::Dash),
        "longdash" => Some(VmlDashStyle::LongDash),
        "dashdot" => Some(VmlDashStyle::DashDot),
        "longdashdot" => Some(VmlDashStyle::LongDashDot),
        "longdashdotdot" => Some(VmlDashStyle::LongDashDotDot),
        _ => None,
    }
}

fn parse_join_style(s: &str) -> Option<VmlJoinStyle> {
    match s {
        "round" => Some(VmlJoinStyle::Round),
        "bevel" => Some(VmlJoinStyle::Bevel),
        "miter" => Some(VmlJoinStyle::Miter),
        _ => None,
    }
}

// ── path ──────────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct PathXml {
    #[serde(rename = "@gradientshapeok", default)]
    pub gradient_shape_ok: Option<VmlBool>,
    #[serde(rename = "@connecttype", default)]
    pub connect_type: Option<String>,
    #[serde(rename = "@extrusionok", default)]
    pub extrusion_ok: Option<VmlBool>,
}

impl From<PathXml> for VmlPath {
    fn from(x: PathXml) -> Self {
        Self {
            gradient_shape_ok: x.gradient_shape_ok.map(|b| b.0),
            connect_type: x.connect_type.as_deref().and_then(|s| match s {
                "none" => Some(VmlConnectType::None),
                "rect" => Some(VmlConnectType::Rect),
                "segments" => Some(VmlConnectType::Segments),
                "custom" => Some(VmlConnectType::Custom),
                _ => None,
            }),
            extrusion_ok: x.extrusion_ok.map(|b| b.0),
        }
    }
}

// ── formulas ──────────────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
pub(crate) struct FormulasXml {
    #[serde(rename = "f", default)]
    pub entries: Vec<FormulaEntryXml>,
}

#[derive(Deserialize)]
pub(crate) struct FormulaEntryXml {
    #[serde(rename = "@eqn", default)]
    pub eqn: String,
}

// Re-exported so VmlFormula is reachable without silencing `unused` warnings.
#[doc(hidden)]
pub(crate) type _KeepFormulaAlive = VmlFormula;

// ── lock ──────────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct LockXml {
    #[serde(rename = "@aspectratio", default)]
    pub aspect_ratio: Option<VmlBool>,
    #[serde(rename = "@ext", default)]
    pub ext: Option<String>,
}

impl From<LockXml> for VmlLock {
    fn from(x: LockXml) -> Self {
        let ext = x.ext.as_deref().and_then(|s| match s {
            "edit" => Some(VmlExtHandling::Edit),
            "view" => Some(VmlExtHandling::View),
            "backwardCompatible" => Some(VmlExtHandling::BackwardCompatible),
            _ => None,
        });
        Self {
            aspect_ratio: x.aspect_ratio.map(|b| b.0),
            ext,
        }
    }
}

// ── imagedata ─────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct ImageDataXml {
    #[serde(rename = "@id", default)]
    pub id: Option<String>,
    #[serde(rename = "@title", default)]
    pub title: Option<String>,
}

impl From<ImageDataXml> for VmlImageData {
    fn from(x: ImageDataXml) -> Self {
        Self {
            rel_id: x.id.map(RelId::new),
            title: x.title,
        }
    }
}

// ── wrap ──────────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct WrapXml {
    #[serde(rename = "@type", default)]
    pub ty: Option<String>,
    #[serde(rename = "@side", default)]
    pub side: Option<String>,
}

impl From<WrapXml> for VmlWrap {
    fn from(x: WrapXml) -> Self {
        let wrap_type = x.ty.as_deref().and_then(|s| match s {
            "topAndBottom" => Some(VmlWrapType::TopAndBottom),
            "square" => Some(VmlWrapType::Square),
            "none" => Some(VmlWrapType::None),
            "tight" => Some(VmlWrapType::Tight),
            "through" => Some(VmlWrapType::Through),
            _ => None,
        });
        let side = x.side.as_deref().and_then(|s| match s {
            "both" => Some(VmlWrapSide::Both),
            "left" => Some(VmlWrapSide::Left),
            "right" => Some(VmlWrapSide::Right),
            "largest" => Some(VmlWrapSide::Largest),
            _ => None,
        });
        Self { wrap_type, side }
    }
}

// ── helpers ───────────────────────────────────────────────────────────────

/// VML boolean attribute — `"t"`/`"true"` → true, `"f"`/`"false"` → false,
/// anything else silently ignored.
#[derive(Clone, Copy, Debug)]
pub(crate) struct VmlBool(pub bool);

impl<'de> Deserialize<'de> for VmlBool {
    fn deserialize<D: serde::Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        Ok(Self(matches!(s.as_str(), "t" | "true")))
    }
}

// Keep VmlVector2D reachable so it doesn't trip dead-code lint when the
// only user is via parse_vector2d.
#[doc(hidden)]
pub(crate) fn _keep_vmlvector_alive(v: VmlVector2D) -> VmlVector2D {
    v
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::VmlColor;

    fn parse(xml: &str) -> Pict {
        let wrapped = format!(
            r#"<wrap xmlns:v="urn:v" xmlns:o="urn:o" xmlns:w="urn:w" xmlns:r="urn:r">{}</wrap>"#,
            xml
        );
        #[derive(Deserialize)]
        struct Wrap {
            pict: PictXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        let mut ctx = crate::docx::parse::body::ConvertCtx::new();
        w.pict.into_model(&mut ctx)
    }

    #[test]
    fn empty_pict() {
        let p = parse(r#"<pict/>"#);
        assert!(p.shape_type.is_none());
        assert!(p.shapes.is_empty());
    }

    #[test]
    fn minimal_shape_with_fill_and_type_ref() {
        let p = parse(
            r##"<pict>
                <shape id="s1" type="#_x0000_t202"
                       style="position:absolute;left:0;top:0;width:100pt;height:50pt"
                       fillcolor="#ff0000" stroked="f"/>
            </pict>"##,
        );
        assert_eq!(p.shapes.len(), 1);
        let s = &p.shapes[0];
        assert_eq!(s.id.as_ref().map(|v| v.as_str()), Some("s1"));
        assert_eq!(
            s.shape_type_ref.as_ref().map(|v| v.as_str()),
            Some("_x0000_t202")
        );
        assert_eq!(s.stroked, Some(false));
        assert!(matches!(s.fill_color, Some(VmlColor::Rgb(0xFF, 0, 0))));
    }

    #[test]
    fn shapetype_with_stroke_and_path() {
        let p = parse(
            r#"<pict>
                <shapetype id="t1" coordsize="21600,21600" adj="5400,5400"
                           filled="t" stroked="t" path="m 0,0 l 10,10 x e">
                    <stroke dashstyle="dash" joinstyle="miter"/>
                    <path gradientshapeok="t" connecttype="rect" extrusionok="f"/>
                    <lock aspectratio="t" ext="edit"/>
                </shapetype>
            </pict>"#,
        );
        let st = p.shape_type.expect("shapetype parsed");
        assert_eq!(st.filled, Some(true));
        assert_eq!(st.stroked, Some(true));
        assert_eq!(st.adj, vec![5400, 5400]);
        assert_eq!(st.coord_size, Some(VmlVector2D { x: 21600, y: 21600 }));
        assert!(!st.path.is_empty());
        let stroke = st.stroke.unwrap();
        assert_eq!(stroke.dash_style, Some(VmlDashStyle::Dash));
        assert_eq!(stroke.join_style, Some(VmlJoinStyle::Miter));
        assert_eq!(
            st.vml_path.unwrap().connect_type,
            Some(VmlConnectType::Rect)
        );
        let lock = st.lock.unwrap();
        assert_eq!(lock.aspect_ratio, Some(true));
        assert_eq!(lock.ext, Some(VmlExtHandling::Edit));
    }

    #[test]
    fn shape_with_imagedata() {
        let p = parse(
            r##"<pict>
                <shape id="img1" type="#_x0000_t75" style="width:100pt;height:50pt">
                    <imagedata r:id="rId7" o:title="Logo"/>
                </shape>
            </pict>"##,
        );
        let id = p.shapes[0].image_data.as_ref().unwrap();
        assert_eq!(id.rel_id.as_ref().map(|r| r.as_str()), Some("rId7"));
        assert_eq!(id.title.as_deref(), Some("Logo"));
    }

    #[test]
    fn shape_with_textbox_no_content() {
        let p = parse(
            r#"<pict>
                <shape id="tb" style="width:100pt;height:50pt">
                    <textbox style="mso-fit-shape-to-text:t" inset="1pt,2pt,3pt,4pt">
                        <txbxContent/>
                    </textbox>
                </shape>
            </pict>"#,
        );
        let tb = p.shapes[0].text_box.as_ref().unwrap();
        assert!(tb.content.is_empty());
        assert!(tb.inset.is_some());
    }

    #[test]
    fn shape_with_textbox_containing_paragraph() {
        let p = parse(
            r#"<pict>
                <shape id="tb" style="width:100pt">
                    <textbox>
                        <txbxContent>
                            <w:p><w:r><w:t>Hello VML</w:t></w:r></w:p>
                        </txbxContent>
                    </textbox>
                </shape>
            </pict>"#,
        );
        let tb = p.shapes[0].text_box.as_ref().unwrap();
        assert_eq!(tb.content.len(), 1);
        match &tb.content[0] {
            Block::Paragraph(_) => (),
            other => panic!("expected Paragraph, got {other:?}"),
        }
    }

    #[test]
    fn shape_with_wrap() {
        let p = parse(
            r#"<pict>
                <shape id="wr" style="width:100pt">
                    <wrap type="square" side="both"/>
                </shape>
            </pict>"#,
        );
        let w = p.shapes[0].wrap.unwrap();
        assert_eq!(w.wrap_type, Some(VmlWrapType::Square));
        assert_eq!(w.side, Some(VmlWrapSide::Both));
    }

    #[test]
    fn multiple_shapes_preserve_order() {
        let p = parse(
            r#"<pict>
                <shape id="a" style="width:10pt"/>
                <shape id="b" style="width:20pt"/>
                <shape id="c" style="width:30pt"/>
            </pict>"#,
        );
        assert_eq!(p.shapes.len(), 3);
        assert_eq!(p.shapes[0].id.as_ref().unwrap().as_str(), "a");
        assert_eq!(p.shapes[2].id.as_ref().unwrap().as_str(), "c");
    }

    #[test]
    fn style_is_parsed_via_existing_helper() {
        let p = parse(
            r#"<pict>
                <shape id="s" style="position:absolute;left:10pt;top:20pt;width:100pt;height:50pt;z-index:5"/>
            </pict>"#,
        );
        let st = &p.shapes[0].style;
        assert!(st.position.is_some());
        assert!(st.left.is_some());
        assert_eq!(st.z_index, Some(5));
    }
}
