//! DrawingML outer wrappers: `<wp:inline>` and `<wp:anchor>` → `Image`.
//!
//! These are the entry points from body content — `<w:drawing>` wraps exactly
//! one of these. They share most children (`extent`, `effectExtent`, `docPr`,
//! `cNvGraphicFramePr`, `graphic`) and differ in placement-specific children:
//! inline has just `distT/B/L/R`, while anchor adds position, wrap, and
//! z-ordering.

#![allow(dead_code, clippy::large_enum_variant)]

use serde::Deserialize;

use crate::docx::dimension::{Dimension, Emu};
use crate::docx::geometry::{EdgeInsets, Offset, Size};
use crate::docx::model::{
    AnchorAlignment, AnchorPosition, AnchorProperties, AnchorRelativeFrom, DocProperties,
    GraphicContent, GraphicFrameLocks, Image, ImagePlacement, TextWrap, WrapPolygon, WrapText,
};

use super::fill::AttrBool;
use super::picture::PictureXml;
use super::shape::WspXml;

// ── Shared bits ───────────────────────────────────────────────────────────

/// `<wp:extent cx=".." cy=".."/>` — positive size in EMU.
#[derive(Debug, Deserialize)]
pub struct ExtentXml {
    #[serde(rename = "@cx")]
    pub cx: Dimension<Emu>,
    #[serde(rename = "@cy")]
    pub cy: Dimension<Emu>,
}

/// `<wp:effectExtent l=".." t=".." r=".." b=".."/>` — drawing overflow.
#[derive(Debug, Deserialize)]
pub struct EffectExtentXml {
    #[serde(rename = "@l", default)]
    pub l: Option<Dimension<Emu>>,
    #[serde(rename = "@t", default)]
    pub t: Option<Dimension<Emu>>,
    #[serde(rename = "@r", default)]
    pub r: Option<Dimension<Emu>>,
    #[serde(rename = "@b", default)]
    pub b: Option<Dimension<Emu>>,
}

impl From<EffectExtentXml> for EdgeInsets<Emu> {
    fn from(x: EffectExtentXml) -> Self {
        Self::new(
            x.t.unwrap_or_default(),
            x.r.unwrap_or_default(),
            x.b.unwrap_or_default(),
            x.l.unwrap_or_default(),
        )
    }
}

/// `<wp:docPr>` — inline-image variant (distinct from `<pic:cNvPr>`).
#[derive(Debug, Deserialize)]
pub struct DocPrXml {
    #[serde(rename = "@id", default)]
    pub id: Option<u32>,
    #[serde(rename = "@name", default)]
    pub name: Option<String>,
    #[serde(rename = "@descr", default)]
    pub descr: Option<String>,
    #[serde(rename = "@hidden", default)]
    pub hidden: Option<AttrBool>,
    #[serde(rename = "@title", default)]
    pub title: Option<String>,
}

impl From<DocPrXml> for DocProperties {
    fn from(x: DocPrXml) -> Self {
        Self {
            id: x.id.unwrap_or(0),
            name: x.name.unwrap_or_default(),
            description: x.descr,
            hidden: x.hidden.map(|b| b.0),
            title: x.title,
        }
    }
}

/// `<wp:cNvGraphicFramePr>` — wraps `<a:graphicFrameLocks>`.
#[derive(Debug, Deserialize)]
pub struct CNvGraphicFramePrXml {
    #[serde(rename = "graphicFrameLocks", default)]
    pub locks: Option<GraphicFrameLocksXml>,
}

#[derive(Debug, Deserialize)]
pub struct GraphicFrameLocksXml {
    #[serde(rename = "@noChangeAspect", default)]
    pub no_change_aspect: Option<AttrBool>,
    #[serde(rename = "@noDrilldown", default)]
    pub no_drilldown: Option<AttrBool>,
    #[serde(rename = "@noGrp", default)]
    pub no_grp: Option<AttrBool>,
    #[serde(rename = "@noMove", default)]
    pub no_move: Option<AttrBool>,
    #[serde(rename = "@noResize", default)]
    pub no_resize: Option<AttrBool>,
    #[serde(rename = "@noSelect", default)]
    pub no_select: Option<AttrBool>,
}

impl From<CNvGraphicFramePrXml> for GraphicFrameLocks {
    fn from(x: CNvGraphicFramePrXml) -> Self {
        x.locks
            .map(|l| Self {
                no_change_aspect: l.no_change_aspect.map(|b| b.0),
                no_drilldown: l.no_drilldown.map(|b| b.0),
                no_grp: l.no_grp.map(|b| b.0),
                no_move: l.no_move.map(|b| b.0),
                no_resize: l.no_resize.map(|b| b.0),
                no_select: l.no_select.map(|b| b.0),
            })
            .unwrap_or(Self {
                no_change_aspect: None,
                no_drilldown: None,
                no_grp: None,
                no_move: None,
                no_resize: None,
                no_select: None,
            })
    }
}

/// `<a:graphic>` → `<a:graphicData>` → `<pic:pic>` or `<wps:wsp>`.
#[derive(Deserialize)]
pub(crate) struct GraphicXml {
    #[serde(rename = "graphicData", default)]
    pub(crate) data: Option<GraphicDataXml>,
}

#[derive(Deserialize)]
pub(crate) struct GraphicDataXml {
    #[serde(rename = "pic", default)]
    pub(crate) pic: Option<PictureXml>,
    #[serde(rename = "wsp", default)]
    pub(crate) wsp: Option<WspXml>,
}

impl GraphicXml {
    fn into_content(
        self,
        ctx: &mut crate::docx::parse::body::ConvertCtx,
    ) -> Option<GraphicContent> {
        let data = self.data?;
        if let Some(pic) = data.pic {
            Some(GraphicContent::Picture(pic.into()))
        } else {
            data.wsp
                .map(|w| GraphicContent::WordProcessingShape(w.into_model(ctx)))
        }
    }
}

// ── Inline ────────────────────────────────────────────────────────────────

/// `<wp:inline>` — inline drawing.
#[derive(Deserialize)]
pub(crate) struct InlineXml {
    #[serde(rename = "@distT", default)]
    pub dist_t: Option<Dimension<Emu>>,
    #[serde(rename = "@distB", default)]
    pub dist_b: Option<Dimension<Emu>>,
    #[serde(rename = "@distL", default)]
    pub dist_l: Option<Dimension<Emu>>,
    #[serde(rename = "@distR", default)]
    pub dist_r: Option<Dimension<Emu>>,

    #[serde(rename = "extent")]
    pub extent: ExtentXml,
    #[serde(rename = "effectExtent", default)]
    pub effect_extent: Option<EffectExtentXml>,
    #[serde(rename = "docPr")]
    pub doc_pr: DocPrXml,
    #[serde(rename = "cNvGraphicFramePr", default)]
    pub cnv_gfp: Option<CNvGraphicFramePrXml>,
    #[serde(rename = "graphic", default)]
    pub graphic: Option<GraphicXml>,
}

impl InlineXml {
    pub(crate) fn into_image(self, ctx: &mut crate::docx::parse::body::ConvertCtx) -> Image {
        let distance = EdgeInsets::new(
            self.dist_t.unwrap_or_default(),
            self.dist_r.unwrap_or_default(),
            self.dist_b.unwrap_or_default(),
            self.dist_l.unwrap_or_default(),
        );
        Image {
            extent: Size::new(self.extent.cx, self.extent.cy),
            effect_extent: self.effect_extent.map(Into::into),
            doc_properties: self.doc_pr.into(),
            graphic_frame_locks: self.cnv_gfp.map(Into::into),
            graphic: self.graphic.and_then(|g| g.into_content(ctx)),
            placement: ImagePlacement::Inline { distance },
        }
    }
}

// ── Anchor ────────────────────────────────────────────────────────────────

#[derive(Deserialize)]
pub(crate) struct AnchorXml {
    #[serde(rename = "@distT", default)]
    pub dist_t: Option<Dimension<Emu>>,
    #[serde(rename = "@distB", default)]
    pub dist_b: Option<Dimension<Emu>>,
    #[serde(rename = "@distL", default)]
    pub dist_l: Option<Dimension<Emu>>,
    #[serde(rename = "@distR", default)]
    pub dist_r: Option<Dimension<Emu>>,
    #[serde(rename = "@simplePos", default)]
    pub simple_pos_attr: Option<AttrBool>,
    #[serde(rename = "@relativeHeight")]
    pub relative_height: u32,
    #[serde(rename = "@behindDoc")]
    pub behind_doc: AttrBool,
    #[serde(rename = "@locked")]
    pub locked: AttrBool,
    #[serde(rename = "@allowOverlap")]
    pub allow_overlap: AttrBool,
    #[serde(rename = "@layoutInCell", default)]
    pub layout_in_cell: Option<AttrBool>,
    #[serde(rename = "@hidden", default)]
    pub hidden: Option<AttrBool>,

    #[serde(rename = "simplePos", default)]
    pub simple_pos: Option<SimplePosXml>,
    #[serde(rename = "positionH", default)]
    pub pos_h: Option<PositionXml>,
    #[serde(rename = "positionV", default)]
    pub pos_v: Option<PositionXml>,

    // Wrap variants — at most one present.
    #[serde(rename = "wrapNone", default)]
    pub wrap_none: Option<super::fill::Empty>,
    #[serde(rename = "wrapSquare", default)]
    pub wrap_square: Option<WrapSquareXml>,
    #[serde(rename = "wrapTight", default)]
    pub wrap_tight: Option<WrapTightThroughXml>,
    #[serde(rename = "wrapThrough", default)]
    pub wrap_through: Option<WrapTightThroughXml>,
    #[serde(rename = "wrapTopAndBottom", default)]
    pub wrap_top_and_bottom: Option<WrapTopAndBottomXml>,

    #[serde(rename = "extent")]
    pub extent: ExtentXml,
    #[serde(rename = "effectExtent", default)]
    pub effect_extent: Option<EffectExtentXml>,
    #[serde(rename = "docPr")]
    pub doc_pr: DocPrXml,
    #[serde(rename = "cNvGraphicFramePr", default)]
    pub cnv_gfp: Option<CNvGraphicFramePrXml>,
    #[serde(rename = "graphic", default)]
    pub graphic: Option<GraphicXml>,
}

#[derive(Debug, Deserialize)]
pub struct SimplePosXml {
    #[serde(rename = "@x")]
    pub x: Dimension<Emu>,
    #[serde(rename = "@y")]
    pub y: Dimension<Emu>,
}

/// `<wp:positionH>` or `<wp:positionV>` — one of `<wp:posOffset>` or
/// `<wp:align>`.
#[derive(Debug, Deserialize)]
pub struct PositionXml {
    #[serde(rename = "@relativeFrom")]
    pub relative_from: StRelFrom,
    #[serde(rename = "posOffset", default)]
    pub pos_offset: Option<PosOffsetXml>,
    #[serde(rename = "align", default)]
    pub align: Option<AlignXml>,
}

#[derive(Debug, Deserialize)]
pub struct PosOffsetXml {
    #[serde(rename = "$text", default)]
    pub value: String,
}

#[derive(Debug, Deserialize)]
pub struct AlignXml {
    #[serde(rename = "$text")]
    pub value: StAnchorAlignment,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StRelFrom {
    Page,
    Margin,
    Column,
    Character,
    Paragraph,
    Line,
    InsideMargin,
    OutsideMargin,
    TopMargin,
    BottomMargin,
    LeftMargin,
    RightMargin,
}

impl From<StRelFrom> for AnchorRelativeFrom {
    fn from(s: StRelFrom) -> Self {
        use StRelFrom as X;
        match s {
            X::Page => Self::Page,
            X::Margin => Self::Margin,
            X::Column => Self::Column,
            X::Character => Self::Character,
            X::Paragraph => Self::Paragraph,
            X::Line => Self::Line,
            X::InsideMargin => Self::InsideMargin,
            X::OutsideMargin => Self::OutsideMargin,
            X::TopMargin => Self::TopMargin,
            X::BottomMargin => Self::BottomMargin,
            X::LeftMargin => Self::LeftMargin,
            X::RightMargin => Self::RightMargin,
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StAnchorAlignment {
    Left,
    Center,
    Right,
    Inside,
    Outside,
    Top,
    Bottom,
}

impl From<StAnchorAlignment> for AnchorAlignment {
    fn from(s: StAnchorAlignment) -> Self {
        use StAnchorAlignment as X;
        match s {
            X::Left => Self::Left,
            X::Center => Self::Center,
            X::Right => Self::Right,
            X::Inside => Self::Inside,
            X::Outside => Self::Outside,
            X::Top => Self::Top,
            X::Bottom => Self::Bottom,
        }
    }
}

#[derive(Debug, Deserialize)]
pub struct WrapSquareXml {
    #[serde(rename = "@wrapText")]
    pub wrap_text: StWrapText,
    #[serde(rename = "@distT", default)]
    pub dist_t: Option<Dimension<Emu>>,
    #[serde(rename = "@distB", default)]
    pub dist_b: Option<Dimension<Emu>>,
    #[serde(rename = "@distL", default)]
    pub dist_l: Option<Dimension<Emu>>,
    #[serde(rename = "@distR", default)]
    pub dist_r: Option<Dimension<Emu>>,
}

#[derive(Debug, Deserialize)]
pub struct WrapTightThroughXml {
    #[serde(rename = "@wrapText")]
    pub wrap_text: StWrapText,
    #[serde(rename = "@distL", default)]
    pub dist_l: Option<Dimension<Emu>>,
    #[serde(rename = "@distR", default)]
    pub dist_r: Option<Dimension<Emu>>,
    #[serde(rename = "wrapPolygon", default)]
    pub polygon: Option<WrapPolygonXml>,
}

#[derive(Debug, Deserialize)]
pub struct WrapTopAndBottomXml {
    #[serde(rename = "@distT", default)]
    pub dist_t: Option<Dimension<Emu>>,
    #[serde(rename = "@distB", default)]
    pub dist_b: Option<Dimension<Emu>>,
}

#[derive(Debug, Deserialize)]
pub struct WrapPolygonXml {
    #[serde(rename = "@edited", default)]
    pub edited: Option<AttrBool>,
    #[serde(rename = "start", default)]
    pub start: Option<Point2DXml>,
    #[serde(rename = "lineTo", default)]
    pub line_to: Vec<Point2DXml>,
}

#[derive(Debug, Deserialize)]
pub struct Point2DXml {
    #[serde(rename = "@x")]
    pub x: Dimension<Emu>,
    #[serde(rename = "@y")]
    pub y: Dimension<Emu>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StWrapText {
    BothSides,
    Left,
    Right,
    Largest,
}

impl From<StWrapText> for WrapText {
    fn from(s: StWrapText) -> Self {
        match s {
            StWrapText::BothSides => Self::BothSides,
            StWrapText::Left => Self::Left,
            StWrapText::Right => Self::Right,
            StWrapText::Largest => Self::Largest,
        }
    }
}

impl AnchorXml {
    pub(crate) fn into_image(self, ctx: &mut crate::docx::parse::body::ConvertCtx) -> Image {
        let distance = EdgeInsets::new(
            self.dist_t.unwrap_or_default(),
            self.dist_r.unwrap_or_default(),
            self.dist_b.unwrap_or_default(),
            self.dist_l.unwrap_or_default(),
        );
        let simple_pos = self.simple_pos.map(|s| Offset::new(s.x, s.y));
        let use_simple_pos = self.simple_pos_attr.map(|b| b.0);
        let horizontal_position = position(self.pos_h);
        let vertical_position = position(self.pos_v);
        let wrap = pick_wrap(
            self.wrap_none,
            self.wrap_square,
            self.wrap_tight,
            self.wrap_through,
            self.wrap_top_and_bottom,
        );
        Image {
            extent: Size::new(self.extent.cx, self.extent.cy),
            effect_extent: self.effect_extent.map(Into::into),
            doc_properties: self.doc_pr.into(),
            graphic_frame_locks: self.cnv_gfp.map(Into::into),
            graphic: self.graphic.and_then(|g| g.into_content(ctx)),
            placement: ImagePlacement::Anchor(AnchorProperties {
                distance,
                simple_pos,
                use_simple_pos,
                horizontal_position,
                vertical_position,
                wrap,
                behind_text: self.behind_doc.0,
                lock_anchor: self.locked.0,
                allow_overlap: self.allow_overlap.0,
                relative_height: self.relative_height,
                layout_in_cell: self.layout_in_cell.map(|b| b.0),
                hidden: self.hidden.map(|b| b.0),
            }),
        }
    }
}

fn position(p: Option<PositionXml>) -> AnchorPosition {
    // Fallback to page+left when nothing is specified.
    let Some(p) = p else {
        return AnchorPosition::Align {
            relative_from: AnchorRelativeFrom::Page,
            alignment: AnchorAlignment::Left,
        };
    };
    let rel = p.relative_from.into();
    if let Some(a) = p.align {
        return AnchorPosition::Align {
            relative_from: rel,
            alignment: a.value.into(),
        };
    }
    if let Some(o) = p.pos_offset {
        let offset = o.value.trim().parse::<i64>().unwrap_or(0);
        return AnchorPosition::Offset {
            relative_from: rel,
            offset: Dimension::new(offset),
        };
    }
    AnchorPosition::Offset {
        relative_from: rel,
        offset: Dimension::new(0),
    }
}

fn pick_wrap(
    none: Option<super::fill::Empty>,
    square: Option<WrapSquareXml>,
    tight: Option<WrapTightThroughXml>,
    through: Option<WrapTightThroughXml>,
    top_and_bottom: Option<WrapTopAndBottomXml>,
) -> TextWrap {
    if none.is_some() {
        return TextWrap::None;
    }
    if let Some(s) = square {
        return TextWrap::Square {
            distance: EdgeInsets::new(
                s.dist_t.unwrap_or_default(),
                s.dist_r.unwrap_or_default(),
                s.dist_b.unwrap_or_default(),
                s.dist_l.unwrap_or_default(),
            ),
            wrap_text: s.wrap_text.into(),
        };
    }
    if let Some(t) = tight {
        return TextWrap::Tight {
            distance: EdgeInsets::new(
                Dimension::new(0),
                t.dist_r.unwrap_or_default(),
                Dimension::new(0),
                t.dist_l.unwrap_or_default(),
            ),
            wrap_text: t.wrap_text.into(),
            polygon: t.polygon.and_then(polygon),
        };
    }
    if let Some(th) = through {
        return TextWrap::Through {
            distance: EdgeInsets::new(
                Dimension::new(0),
                th.dist_r.unwrap_or_default(),
                Dimension::new(0),
                th.dist_l.unwrap_or_default(),
            ),
            wrap_text: th.wrap_text.into(),
            polygon: th.polygon.and_then(polygon),
        };
    }
    if let Some(tb) = top_and_bottom {
        return TextWrap::TopAndBottom {
            distance_top: tb.dist_t.unwrap_or_default(),
            distance_bottom: tb.dist_b.unwrap_or_default(),
        };
    }
    TextWrap::None
}

fn polygon(p: WrapPolygonXml) -> Option<WrapPolygon> {
    let start = p.start?;
    Some(WrapPolygon {
        edited: p.edited.map(|b| b.0),
        start: Offset::new(start.x, start.y),
        line_to: p
            .line_to
            .into_iter()
            .map(|pt| Offset::new(pt.x, pt.y))
            .collect(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_inline(xml: &str) -> Image {
        let wrapped = format!(
            r#"<wrap xmlns:wp="urn:wp" xmlns:a="urn:a" xmlns:r="urn:r" xmlns:pic="urn:pic" xmlns:wps="urn:wps" xmlns:w="urn:w">{}</wrap>"#,
            xml
        );
        #[derive(Deserialize)]
        struct Wrap {
            inline: InlineXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        let mut ctx = crate::docx::parse::body::ConvertCtx::new();
        w.inline.into_image(&mut ctx)
    }

    fn parse_anchor(xml: &str) -> Image {
        let wrapped = format!(
            r#"<wrap xmlns:wp="urn:wp" xmlns:a="urn:a" xmlns:r="urn:r" xmlns:pic="urn:pic" xmlns:wps="urn:wps" xmlns:w="urn:w">{}</wrap>"#,
            xml
        );
        #[derive(Deserialize)]
        struct Wrap {
            anchor: AnchorXml,
        }
        let w: Wrap = quick_xml::de::from_str(&wrapped).unwrap();
        let mut ctx = crate::docx::parse::body::ConvertCtx::new();
        w.anchor.into_image(&mut ctx)
    }

    #[test]
    fn inline_with_picture() {
        let img = parse_inline(
            r#"<inline distT="0" distB="0" distL="0" distR="0">
                <extent cx="914400" cy="457200"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="1" name="image1"/>
                <cNvGraphicFramePr><graphicFrameLocks noChangeAspect="1"/></cNvGraphicFramePr>
                <graphic>
                    <graphicData>
                        <pic>
                            <nvPicPr><cNvPr id="1" name="Pic1"/></nvPicPr>
                            <blipFill><blip r:embed="rId1"/></blipFill>
                        </pic>
                    </graphicData>
                </graphic>
            </inline>"#,
        );
        assert_eq!(img.extent.width.raw(), 914400);
        assert_eq!(img.extent.height.raw(), 457200);
        assert_eq!(img.doc_properties.name, "image1");
        assert!(matches!(img.placement, ImagePlacement::Inline { .. }));
        assert!(matches!(img.graphic, Some(GraphicContent::Picture(_))));
    }

    #[test]
    fn inline_without_effect_extent_defaults_to_none() {
        let img = parse_inline(
            r#"<inline>
                <extent cx="100" cy="100"/>
                <docPr id="2" name="bare"/>
                <graphic><graphicData/></graphic>
            </inline>"#,
        );
        assert!(img.effect_extent.is_none());
        assert!(img.graphic.is_none());
    }

    #[test]
    fn anchor_with_offset_positions_and_square_wrap() {
        let img = parse_anchor(
            r#"<anchor distT="0" distB="0" distL="114300" distR="114300"
                      simplePos="0" relativeHeight="251659264"
                      behindDoc="0" locked="0" allowOverlap="1">
                <simplePos x="0" y="0"/>
                <positionH relativeFrom="column">
                    <posOffset>1524000</posOffset>
                </positionH>
                <positionV relativeFrom="paragraph">
                    <posOffset>762000</posOffset>
                </positionV>
                <wrapSquare wrapText="bothSides"/>
                <extent cx="914400" cy="457200"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="3" name="float"/>
                <graphic><graphicData><pic>
                    <nvPicPr><cNvPr id="3" name="P3"/></nvPicPr>
                    <blipFill><blip r:embed="rId3"/></blipFill>
                </pic></graphicData></graphic>
            </anchor>"#,
        );
        let ImagePlacement::Anchor(a) = &img.placement else {
            panic!("expected Anchor");
        };
        assert_eq!(a.relative_height, 251_659_264);
        assert!(!a.behind_text);
        assert!(a.allow_overlap);
        match a.horizontal_position {
            AnchorPosition::Offset {
                relative_from,
                offset,
            } => {
                assert_eq!(relative_from, AnchorRelativeFrom::Column);
                assert_eq!(offset.raw(), 1_524_000);
            }
            other => panic!("expected Offset, got {other:?}"),
        }
        match a.vertical_position {
            AnchorPosition::Offset { offset, .. } => assert_eq!(offset.raw(), 762_000),
            other => panic!("expected Offset, got {other:?}"),
        }
        match &a.wrap {
            TextWrap::Square { wrap_text, .. } => assert_eq!(*wrap_text, WrapText::BothSides),
            other => panic!("expected Square, got {other:?}"),
        }
    }

    #[test]
    fn anchor_with_align_positions() {
        let img = parse_anchor(
            r#"<anchor distT="0" distB="0" distL="0" distR="0"
                      simplePos="0" relativeHeight="1"
                      behindDoc="0" locked="0" allowOverlap="1">
                <simplePos x="0" y="0"/>
                <positionH relativeFrom="page"><align>center</align></positionH>
                <positionV relativeFrom="page"><align>top</align></positionV>
                <wrapNone/>
                <extent cx="100" cy="100"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="4" name="centered"/>
                <graphic><graphicData/></graphic>
            </anchor>"#,
        );
        let ImagePlacement::Anchor(a) = &img.placement else {
            panic!("expected Anchor");
        };
        match a.horizontal_position {
            AnchorPosition::Align {
                relative_from,
                alignment,
            } => {
                assert_eq!(relative_from, AnchorRelativeFrom::Page);
                assert_eq!(alignment, AnchorAlignment::Center);
            }
            other => panic!("expected Align, got {other:?}"),
        }
        assert!(matches!(a.wrap, TextWrap::None));
    }

    #[test]
    fn anchor_with_tight_wrap_and_polygon() {
        let img = parse_anchor(
            r#"<anchor distT="0" distB="0" distL="0" distR="0"
                      simplePos="0" relativeHeight="2"
                      behindDoc="0" locked="0" allowOverlap="1">
                <simplePos x="0" y="0"/>
                <positionH relativeFrom="column"><posOffset>0</posOffset></positionH>
                <positionV relativeFrom="paragraph"><posOffset>0</posOffset></positionV>
                <wrapTight wrapText="bothSides">
                    <wrapPolygon edited="0">
                        <start x="0" y="0"/>
                        <lineTo x="100" y="0"/>
                        <lineTo x="100" y="100"/>
                        <lineTo x="0" y="100"/>
                    </wrapPolygon>
                </wrapTight>
                <extent cx="100" cy="100"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="5" name="tight"/>
                <graphic><graphicData/></graphic>
            </anchor>"#,
        );
        let ImagePlacement::Anchor(a) = &img.placement else {
            panic!("expected Anchor");
        };
        match &a.wrap {
            TextWrap::Tight {
                polygon: Some(p), ..
            } => {
                assert_eq!(p.line_to.len(), 3);
                assert_eq!(p.start.x.raw(), 0);
                assert_eq!(p.edited, Some(false));
            }
            other => panic!("expected Tight with polygon, got {other:?}"),
        }
    }

    #[test]
    fn anchor_with_top_and_bottom_wrap() {
        let img = parse_anchor(
            r#"<anchor distT="0" distB="0" distL="0" distR="0"
                      simplePos="0" relativeHeight="3"
                      behindDoc="0" locked="0" allowOverlap="1">
                <simplePos x="0" y="0"/>
                <positionH relativeFrom="margin"><posOffset>0</posOffset></positionH>
                <positionV relativeFrom="paragraph"><posOffset>0</posOffset></positionV>
                <wrapTopAndBottom distT="181610" distB="181610"/>
                <extent cx="100" cy="100"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="6" name="tb"/>
                <graphic><graphicData/></graphic>
            </anchor>"#,
        );
        let ImagePlacement::Anchor(a) = &img.placement else {
            panic!("expected Anchor");
        };
        match &a.wrap {
            TextWrap::TopAndBottom {
                distance_top,
                distance_bottom,
            } => {
                assert_eq!(distance_top.raw(), 181_610);
                assert_eq!(distance_bottom.raw(), 181_610);
            }
            other => panic!("expected TopAndBottom, got {other:?}"),
        }
    }

    #[test]
    fn anchor_behind_text_flag() {
        let img = parse_anchor(
            r#"<anchor distT="0" distB="0" distL="0" distR="0"
                      simplePos="0" relativeHeight="1"
                      behindDoc="1" locked="0" allowOverlap="0">
                <simplePos x="0" y="0"/>
                <positionH relativeFrom="page"><align>right</align></positionH>
                <positionV relativeFrom="page"><align>bottom</align></positionV>
                <wrapNone/>
                <extent cx="1" cy="1"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="7" name="watermark"/>
                <graphic><graphicData/></graphic>
            </anchor>"#,
        );
        let ImagePlacement::Anchor(a) = &img.placement else {
            panic!();
        };
        assert!(a.behind_text);
        assert!(!a.allow_overlap);
    }

    #[test]
    fn wsp_graphic_content() {
        let img = parse_inline(
            r#"<inline>
                <extent cx="100" cy="100"/>
                <effectExtent l="0" t="0" r="0" b="0"/>
                <docPr id="8" name="shape"/>
                <graphic><graphicData>
                    <wsp>
                        <cNvPr id="8" name="S8"/>
                        <spPr><prstGeom prst="rect"/></spPr>
                        <bodyPr/>
                    </wsp>
                </graphicData></graphic>
            </inline>"#,
        );
        assert!(matches!(
            img.graphic,
            Some(GraphicContent::WordProcessingShape(_))
        ));
    }
}
