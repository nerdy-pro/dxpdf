//! SmartArt parts (issue #155).
//!
//! A `dgm:relIds` reference names four parts, and rendering needs a fifth
//! that hangs off the first: the diagram **data** part's
//! `dgm:extLst/a:ext/dsp:dataModelExt/@relId` names the [MS-ODRAWXML]
//! `dsp:` **drawing** part — "the last successful layout of the diagram",
//! written by Word 2007 SP2 and later. Both relationship ids resolve against
//! the *host* part's rels (document.xml.rels, or a header's own), not the
//! data part's — the data part typically has no rels at all.
//!
//! The drawing part is the whole story for rendering: literal shapes with
//! baked fills and fitted text, in absolute EMU on a canvas the size of the
//! hosting drawing's `wp:extent`. The layout/quick-style/colors parts are
//! inputs to the layout *algorithm*, which this engine deliberately does not
//! implement — LibreOffice's own, after years of work, is still approximate,
//! and every Word-produced document carries the baked result. A diagram
//! whose drawing part is absent (Word 2007 RTM, some third-party producers)
//! renders nothing, with one warning per document.

use serde::Deserialize;

use crate::docx::error::Result;
use crate::docx::model::dimension::{Dimension, Emu};
use crate::docx::model::{DiagramDrawing, DiagramShape, ShapeProperties, Transform2D};
use crate::model::Dup;

use super::drawing::schema::shape::{ShapeStyleXml, SpPrXml, XfrmXml};
use super::drawing::schema::text_body::TextBodyXml;

// ── The data part: only the drawing-part pointer is read ──────────────────

/// Extract `dsp:dataModelExt/@relId` from a diagram data part, if present.
pub fn drawing_rel_id(data: &[u8]) -> Result<Option<String>> {
    let parsed: DataModelXml = super::serde_xml::from_xml(data)?;
    Ok(parsed
        .ext_lst
        .into_iter()
        .flat_map(|l| l.ext)
        .flat_map(|e| e.data_model_ext)
        .next()
        .map(|d| d.rel_id))
}

/// `dgm:dataModel` — everything but the extension list is ignored.
#[derive(Debug, Deserialize)]
struct DataModelXml {
    #[serde(rename = "extLst", default)]
    ext_lst: Vec<DataExtLstXml>,
}

#[derive(Debug, Deserialize)]
struct DataExtLstXml {
    #[serde(rename = "ext", default)]
    ext: Vec<DataExtXml>,
}

#[derive(Debug, Deserialize)]
struct DataExtXml {
    #[serde(rename = "dataModelExt", default)]
    data_model_ext: Vec<DataModelExtXml>,
}

/// [MS-ODRAWXML] §2.10.3.1 `dsp:dataModelExt`.
#[derive(Debug, Deserialize)]
struct DataModelExtXml {
    #[serde(rename = "@relId")]
    rel_id: String,
}

// ── The drawing part: dsp:spTree, flattened ───────────────────────────────

/// Parse a `dsp:` drawing part into leaf shapes in canvas coordinates.
pub fn parse_diagram_drawing(data: &[u8]) -> Result<DiagramDrawing> {
    let parsed: DspDrawingXml = super::serde_xml::from_xml(data)?;
    let mut shapes = Vec::new();
    if let Some(tree) = Dup::from(parsed.sp_tree).into_value() {
        flatten_group(tree, GroupMap::identity(), &mut shapes);
    }
    Ok(DiagramDrawing { shapes })
}

/// `dsp:drawing`.
#[derive(Debug, Deserialize)]
struct DspDrawingXml {
    #[serde(rename = "spTree", default)]
    sp_tree: Vec<GroupShapeXml>,
}

/// `dsp:spTree` / `dsp:grpSp` (CT_GroupShape). Everything comes through one
/// `$value` catch-all — quick-xml's serde rejects a struct mixing named
/// element fields with `$value` ("duplicate field") — and sibling
/// `sp`/`grpSp` keep their document order, which is z-order for overlapping
/// shapes.
#[derive(Debug, Deserialize)]
struct GroupShapeXml {
    #[serde(rename = "$value", default)]
    children: Vec<GroupChildXml>,
}

#[derive(Debug, Deserialize)]
enum GroupChildXml {
    #[serde(rename = "grpSpPr")]
    GrpSpPr(GroupShapePrXml),
    #[serde(rename = "sp")]
    Sp(Box<DspShapeXml>),
    #[serde(rename = "grpSp")]
    GrpSp(GroupShapeXml),
    #[serde(other)]
    Other,
}

/// a:CT_GroupShapeProperties, reduced to the transform.
#[derive(Debug, Deserialize)]
struct GroupShapePrXml {
    #[serde(rename = "xfrm", default)]
    xfrm: Vec<GroupXfrmXml>,
}

/// §20.1.7.5 CT_GroupTransform2D — `a:xfrm` with the child-space mapping.
#[derive(Debug, Deserialize)]
struct GroupXfrmXml {
    #[serde(rename = "off", default)]
    off: Vec<super::drawing::schema::shape::OffXml>,
    #[serde(rename = "ext", default)]
    ext: Vec<super::drawing::schema::shape::ExtXml>,
    #[serde(rename = "chOff", default)]
    ch_off: Vec<super::drawing::schema::shape::OffXml>,
    #[serde(rename = "chExt", default)]
    ch_ext: Vec<super::drawing::schema::shape::ExtXml>,
}

/// `dsp:sp` — wrappers are `dsp:`, contents plain `a:` DrawingML.
#[derive(Debug, Default, Deserialize)]
struct DspShapeXml {
    #[serde(rename = "spPr", default)]
    sp_pr: Vec<SpPrXml>,
    #[serde(rename = "style", default)]
    style: Vec<ShapeStyleXml>,
    #[serde(rename = "txBody", default)]
    tx_body: Vec<TextBodyXml>,
    #[serde(rename = "txXfrm", default)]
    tx_xfrm: Vec<XfrmXml>,
}

/// The affine a group applies to its children: child-space EMU → parent
/// EMU, `x' = off + (x − ch_off)·(ext/ch_ext)` per axis (§20.1.7.5).
#[derive(Clone, Copy)]
struct GroupMap {
    off: (f64, f64),
    ch_off: (f64, f64),
    scale: (f64, f64),
}

impl GroupMap {
    fn identity() -> Self {
        Self {
            off: (0.0, 0.0),
            ch_off: (0.0, 0.0),
            scale: (1.0, 1.0),
        }
    }

    fn point(&self, x: i64, y: i64) -> (i64, i64) {
        (
            (self.off.0 + (x as f64 - self.ch_off.0) * self.scale.0).round() as i64,
            (self.off.1 + (y as f64 - self.ch_off.1) * self.scale.1).round() as i64,
        )
    }

    fn size(&self, cx: i64, cy: i64) -> (i64, i64) {
        (
            (cx as f64 * self.scale.0).round() as i64,
            (cy as f64 * self.scale.1).round() as i64,
        )
    }

    /// Compose a child group's own mapping under this one.
    fn child(&self, x: GroupXfrmXml) -> Self {
        let off = Dup::from(x.off)
            .into_value()
            .map(|o| (o.x.raw() as f64, o.y.raw() as f64))
            .unwrap_or((0.0, 0.0));
        let ext = Dup::from(x.ext)
            .into_value()
            .map(|e| (e.cx.raw() as f64, e.cy.raw() as f64));
        let ch_off = Dup::from(x.ch_off)
            .into_value()
            .map(|o| (o.x.raw() as f64, o.y.raw() as f64))
            .unwrap_or((0.0, 0.0));
        let ch_ext = Dup::from(x.ch_ext)
            .into_value()
            .map(|e| (e.cx.raw() as f64, e.cy.raw() as f64));
        // Absent ext/chExt (or a degenerate chExt) means no rescale — Word's
        // own spTree writes no group xfrm at all.
        let scale = match (ext, ch_ext) {
            (Some((ex, ey)), Some((cx, cy))) if cx > 0.0 && cy > 0.0 => (ex / cx, ey / cy),
            _ => (1.0, 1.0),
        };
        // The child's parent-space output feeds this map's own transform:
        // compose by mapping the child's off through self.
        let (px, py) = self.point(off.0 as i64, off.1 as i64);
        Self {
            off: (px as f64, py as f64),
            ch_off,
            scale: (self.scale.0 * scale.0, self.scale.1 * scale.1),
        }
    }
}

fn flatten_group(group: GroupShapeXml, map: GroupMap, out: &mut Vec<DiagramShape>) {
    let mut shapes = Vec::new();
    let mut grp_sp_pr = Vec::new();
    for child in group.children {
        match child {
            GroupChildXml::GrpSpPr(p) => grp_sp_pr.push(p),
            other => shapes.push(other),
        }
    }
    let map = match Dup::from(grp_sp_pr)
        .into_value()
        .and_then(|p| Dup::from(p.xfrm).into_value())
    {
        Some(xfrm) => map.child(xfrm),
        None => map,
    };
    for child in shapes {
        match child {
            GroupChildXml::Sp(sp) => out.push(convert_shape(*sp, &map)),
            GroupChildXml::GrpSp(g) => flatten_group(g, map, out),
            GroupChildXml::GrpSpPr(_) | GroupChildXml::Other => {}
        }
    }
}

fn convert_shape(sp: DspShapeXml, map: &GroupMap) -> DiagramShape {
    let mut shape_properties: Option<ShapeProperties> =
        Dup::from(sp.sp_pr).into_value().map(Into::into);
    if let Some(ref mut props) = shape_properties {
        if let Some(ref mut t) = props.transform {
            apply_map(t, map);
        }
    }
    let mut text_transform: Option<Transform2D> =
        Dup::from(sp.tx_xfrm).into_value().map(Into::into);
    if let Some(ref mut t) = text_transform {
        apply_map(t, map);
    }
    let (style_line_ref, style_fill_ref, style_effect_ref, style_font_ref) =
        match Dup::from(sp.style).into_value() {
            Some(s) => (
                Dup::from(s.ln_ref).into_value().map(Into::into),
                Dup::from(s.fill_ref).into_value().map(Into::into),
                Dup::from(s.effect_ref).into_value().map(Into::into),
                Dup::from(s.font_ref).into_value().map(Into::into),
            ),
            None => (None, None, None, None),
        };
    DiagramShape {
        shape_properties,
        style_line_ref,
        style_fill_ref,
        style_effect_ref,
        style_font_ref,
        text_body: Dup::from(sp.tx_body).into_value().map(Into::into),
        text_transform,
    }
}

fn apply_map(t: &mut Transform2D, map: &GroupMap) {
    if let Some(ref mut off) = t.offset {
        let (x, y) = map.point(off.x.raw(), off.y.raw());
        *off = crate::docx::geometry::Offset::new(Dimension::<Emu>::new(x), Dimension::new(y));
    }
    if let Some(ref mut ext) = t.extent {
        let (cx, cy) = map.size(ext.width.raw(), ext.height.raw());
        *ext = crate::docx::geometry::Size::new(Dimension::<Emu>::new(cx), Dimension::new(cy));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::ShapeGeometry;

    #[test]
    fn data_part_yields_the_drawing_rel_id() {
        let xml = br#"<?xml version="1.0"?>
            <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <dgm:ptLst/>
              <dgm:extLst>
                <a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">
                  <dsp:dataModelExt relId="rId9" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
                </a:ext>
              </dgm:extLst>
            </dgm:dataModel>"#;
        assert_eq!(drawing_rel_id(xml).unwrap().as_deref(), Some("rId9"));
    }

    #[test]
    fn data_part_without_extension_yields_none() {
        let xml = br#"<?xml version="1.0"?>
            <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
              <dgm:ptLst/>
            </dgm:dataModel>"#;
        assert_eq!(drawing_rel_id(xml).unwrap(), None);
    }

    const DRAWING: &[u8] = br#"<?xml version="1.0"?>
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dsp:spTree>
            <dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>
            <dsp:grpSpPr/>
            <dsp:sp modelId="{A}">
              <dsp:nvSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm><a:off x="411480" y="0"/><a:ext cx="4663440" cy="3200400"/></a:xfrm>
                <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
              </dsp:spPr>
              <dsp:style>
                <a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>
                <a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>
                <a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>
                <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
              </dsp:style>
              <dsp:txBody>
                <a:bodyPr/><a:p><a:r><a:rPr lang="en-US" sz="2300"/><a:t>Step</a:t></a:r></a:p>
              </dsp:txBody>
              <dsp:txXfrm><a:off x="500000" y="900000"/><a:ext cx="1000000" cy="500000"/></dsp:txXfrm>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>"#;

    #[test]
    fn word_shaped_drawing_part_parses_to_positioned_shapes() {
        let d = parse_diagram_drawing(DRAWING).unwrap();
        assert_eq!(d.shapes.len(), 1);
        let s = &d.shapes[0];
        let t = s.shape_properties.as_ref().unwrap().transform.unwrap();
        assert_eq!(t.offset.unwrap().x.raw(), 411480, "absolute canvas EMU");
        assert_eq!(t.extent.unwrap().width.raw(), 4663440);
        assert!(matches!(
            s.shape_properties.as_ref().unwrap().geometry,
            Some(ShapeGeometry::Preset(_))
        ));
        assert!(s.style_font_ref.is_some());
        let body = s.text_body.as_ref().unwrap();
        assert_eq!(body.paragraphs.len(), 1);
        let tx = s.text_transform.unwrap();
        assert_eq!(tx.offset.unwrap().x.raw(), 500000);
    }

    /// A nested `dsp:grpSp` with its own child space maps leaves into the
    /// canvas: off (100, 100), children in a half-scale space.
    #[test]
    fn nested_groups_compose_into_canvas_coordinates() {
        let xml = br#"<?xml version="1.0"?>
            <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <dsp:spTree>
                <dsp:grpSpPr/>
                <dsp:grpSp>
                  <dsp:grpSpPr>
                    <a:xfrm>
                      <a:off x="1000" y="2000"/><a:ext cx="500" cy="500"/>
                      <a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/>
                    </a:xfrm>
                  </dsp:grpSpPr>
                  <dsp:sp modelId="{B}">
                    <dsp:spPr>
                      <a:xfrm><a:off x="200" y="400"/><a:ext cx="600" cy="600"/></a:xfrm>
                      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    </dsp:spPr>
                  </dsp:sp>
                </dsp:grpSp>
              </dsp:spTree>
            </dsp:drawing>"#;
        let d = parse_diagram_drawing(xml).unwrap();
        assert_eq!(d.shapes.len(), 1, "groups flatten to leaves");
        let t = d.shapes[0]
            .shape_properties
            .as_ref()
            .unwrap()
            .transform
            .unwrap();
        // x' = 1000 + (200 − 0)·(500/1000) = 1100; cx' = 600·0.5 = 300.
        assert_eq!(t.offset.unwrap().x.raw(), 1100);
        assert_eq!(t.offset.unwrap().y.raw(), 2200);
        assert_eq!(t.extent.unwrap().width.raw(), 300);
    }

    /// The Word 2007 RTM shape: no extension list, no drawing part — the
    /// caller declines. Also: an spTree with zero shapes parses to an empty
    /// drawing rather than an error.
    #[test]
    fn empty_sp_tree_is_an_empty_drawing() {
        let xml = br#"<?xml version="1.0"?>
            <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">
              <dsp:spTree><dsp:grpSpPr/></dsp:spTree>
            </dsp:drawing>"#;
        assert!(parse_diagram_drawing(xml).unwrap().shapes.is_empty());
    }
}
