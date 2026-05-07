//! Rewrite relationship IDs in a parsed block tree.
//!
//! Each XML part in an OOXML package (the main document, every header,
//! every footer, the footnotes/endnotes parts, …) has its **own**
//! relationships file (`_rels/<part>.xml.rels`) with its **own** `rId`
//! namespace. Two different parts may both contain `r:embed="rId1"`
//! pointing at entirely different image targets, or
//! `<w:hyperlink r:id="rId7"/>` pointing at entirely different URLs —
//! the rId is meaningful only with respect to the part it lives in.
//!
//! `Document.media` is a single document-wide `HashMap<RelId, _>`,
//! so a header part's `rId1` can't be allowed to collide with a
//! footer part's `rId1` (the second loader would silently overwrite
//! the first). For each subordinate part the caller (`mod.rs`'s
//! `load_part_rel_remap`) synthesizes a remap that contains:
//!
//! * **image rIds** → globally-unique synthesized ids, with the
//!   loaded image bytes registered in `media` under those ids;
//! * **hyperlink rIds** → the resolved external URL (carried as a
//!   `RelId`-shaped string, the same way the document-level
//!   hyperlink resolver represents a resolved external link).
//!
//! [`rewrite_part_rels_in_blocks`] walks the block tree once and
//! applies the remap to every relationship-bearing node it knows
//! about — pictures, blip fills on shapes, outlines, VML images,
//! VML text-box content, hyperlink targets, fields, and recursively
//! into shape `txbxContent`, table cells, and `AlternateContent`
//! fallbacks. After this pass every rId reference in the part's
//! blocks is unambiguous against `Document.media` (for images) or is
//! already a resolved URL (for hyperlinks).
//!
//! The body uses a separate code path: its image rels populate
//! `media` keyed by their original rIds (which is fine — there's
//! only one body), and `mod.rs::resolve_hyperlinks` does the
//! body-level URL resolution against `doc_rels`.

use std::collections::HashMap;

use crate::model::{
    Blip, BlipFill, Block, DrawingFill, GraphicContent, HyperlinkTarget, Image, Inline, Outline,
    Pict, RelId, ShapeProperties, VmlCommonAttrs, VmlImageData, VmlPrimitive, VmlTextBox,
    WordProcessingShape,
};

/// Apply `remap` to every relationship-bearing `RelId` inside
/// `blocks` — image embeds/links (DrawingML and VML), hyperlink
/// targets, and the same nested in shape `txbxContent`, VML
/// text-box content, table cells, fields, and `AlternateContent`
/// fallbacks. Identifiers absent from the remap are left unchanged,
/// so callers can run this pass over already-resolved content
/// safely.
pub fn rewrite_part_rels_in_blocks(blocks: &mut [Block], remap: &HashMap<RelId, RelId>) {
    if remap.is_empty() {
        return;
    }
    for block in blocks {
        match block {
            Block::Paragraph(p) => rewrite_in_inlines(&mut p.content, remap),
            Block::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        rewrite_part_rels_in_blocks(&mut cell.content, remap);
                    }
                }
            }
            Block::SectionBreak(_) => {}
        }
    }
}

fn rewrite_in_inlines(inlines: &mut [Inline], remap: &HashMap<RelId, RelId>) {
    for inline in inlines {
        match inline {
            Inline::Image(image) => rewrite_in_image(image, remap),
            Inline::Pict(pict) => rewrite_in_pict(pict, remap),
            Inline::Hyperlink(h) => {
                rewrite_in_hyperlink_target(&mut h.target, remap);
                rewrite_in_inlines(&mut h.content, remap);
            }
            Inline::Field(f) => rewrite_in_inlines(&mut f.content, remap),
            Inline::AlternateContent(ac) => {
                if let Some(ref mut fb) = ac.fallback {
                    rewrite_in_inlines(fb, remap);
                }
            }
            Inline::TextRun(_)
            | Inline::FootnoteRef(_)
            | Inline::EndnoteRef(_)
            | Inline::BookmarkStart { .. }
            | Inline::BookmarkEnd(_)
            | Inline::Symbol(_)
            | Inline::Separator
            | Inline::ContinuationSeparator
            | Inline::FieldChar(_)
            | Inline::InstrText(_)
            | Inline::FootnoteRefMark
            | Inline::EndnoteRefMark => {}
        }
    }
}

fn rewrite_in_image(image: &mut Image, remap: &HashMap<RelId, RelId>) {
    if let Some(ref mut graphic) = image.graphic {
        match graphic {
            GraphicContent::Picture(pic) => {
                if let Some(ref mut blip) = pic.blip_fill.blip {
                    rewrite_blip(blip, remap);
                }
                if let Some(ref mut props) = pic.shape_properties {
                    rewrite_in_shape_properties(props, remap);
                }
            }
            GraphicContent::WordProcessingShape(wsp) => {
                rewrite_in_word_processing_shape(wsp, remap);
            }
        }
    }
}

fn rewrite_in_word_processing_shape(wsp: &mut WordProcessingShape, remap: &HashMap<RelId, RelId>) {
    if let Some(ref mut props) = wsp.shape_properties {
        rewrite_in_shape_properties(props, remap);
    }
    // §17.17.1: text inside the shape can host its own paragraph tree
    // with images that share this part's rels namespace.
    rewrite_part_rels_in_blocks(&mut wsp.txbx_content, remap);
}

fn rewrite_in_shape_properties(props: &mut ShapeProperties, remap: &HashMap<RelId, RelId>) {
    if let Some(ref mut fill) = props.fill {
        rewrite_in_drawing_fill(fill, remap);
    }
    if let Some(ref mut outline) = props.outline {
        rewrite_in_outline(outline, remap);
    }
}

fn rewrite_in_outline(outline: &mut Outline, remap: &HashMap<RelId, RelId>) {
    if let Some(ref mut fill) = outline.fill {
        rewrite_in_drawing_fill(fill, remap);
    }
}

fn rewrite_in_drawing_fill(fill: &mut DrawingFill, remap: &HashMap<RelId, RelId>) {
    match fill {
        DrawingFill::Blip(blip_fill) => rewrite_in_blip_fill(blip_fill, remap),
        // None / Solid / Gradient / Pattern / Group: no rIds.
        DrawingFill::None
        | DrawingFill::Solid(_)
        | DrawingFill::Gradient(_)
        | DrawingFill::Pattern(_)
        | DrawingFill::Group => {}
    }
}

fn rewrite_in_blip_fill(blip_fill: &mut BlipFill, remap: &HashMap<RelId, RelId>) {
    if let Some(ref mut blip) = blip_fill.blip {
        rewrite_blip(blip, remap);
    }
}

fn rewrite_blip(blip: &mut Blip, remap: &HashMap<RelId, RelId>) {
    if let Some(new) = blip.embed.as_ref().and_then(|id| remap.get(id)).cloned() {
        blip.embed = Some(new);
    }
    if let Some(new) = blip.link.as_ref().and_then(|id| remap.get(id)).cloned() {
        blip.link = Some(new);
    }
}

fn rewrite_in_pict(pict: &mut Pict, remap: &HashMap<RelId, RelId>) {
    for primitive in &mut pict.primitives {
        rewrite_in_vml_primitive(primitive, remap);
    }
}

/// Walk a single VML primitive, rewriting every `RelId` it owns
/// directly (image data) or holds nested (text-box content, group
/// children). Adding new primitive variants in
/// `crate::model::VmlPrimitive` requires adding them here too — the
/// match arms are exhaustive deliberately.
fn rewrite_in_vml_primitive(p: &mut VmlPrimitive, remap: &HashMap<RelId, RelId>) {
    rewrite_in_vml_common(p.common_mut(), remap);
    match p {
        VmlPrimitive::Group(group) => {
            for child in &mut group.children {
                rewrite_in_vml_primitive(child, remap);
            }
        }
        // Other variants have no nested rels beyond the common attrs.
        VmlPrimitive::Shape(_)
        | VmlPrimitive::Rect(_)
        | VmlPrimitive::RoundRect(_)
        | VmlPrimitive::Oval(_)
        | VmlPrimitive::Line(_)
        | VmlPrimitive::PolyLine(_)
        | VmlPrimitive::Arc(_)
        | VmlPrimitive::Curve(_)
        | VmlPrimitive::Image(_) => {}
    }
}

fn rewrite_in_vml_common(c: &mut VmlCommonAttrs, remap: &HashMap<RelId, RelId>) {
    if let Some(ref mut data) = c.image_data {
        rewrite_vml_image_data(data, remap);
    }
    // §14.1.2.22: a VML shape can host a `<v:textbox>` with block
    // content; images inside it share the part's rels namespace.
    if let Some(ref mut tb) = c.text_box {
        rewrite_in_vml_text_box(tb, remap);
    }
}

fn rewrite_in_vml_text_box(tb: &mut VmlTextBox, remap: &HashMap<RelId, RelId>) {
    rewrite_part_rels_in_blocks(&mut tb.content, remap);
}

fn rewrite_vml_image_data(data: &mut VmlImageData, remap: &HashMap<RelId, RelId>) {
    if let Some(new) = data.rel_id.as_ref().and_then(|id| remap.get(id)).cloned() {
        data.rel_id = Some(new);
    }
}

/// Apply the remap to an `External(rId)` hyperlink target. For
/// header/footer/note parts the remap contains `rId → URL`
/// (resolved through the part's own rels), so this rewrite turns the
/// `External(rId)` placeholder into the final `External(URL)` form
/// the renderer consumes — exactly the job the body-level
/// `resolve_hyperlinks` pass does for the body. When the target is
/// already a URL (e.g. the same blocks were processed twice for some
/// reason) the URL won't match any rId in `remap` and this is a
/// no-op.
fn rewrite_in_hyperlink_target(target: &mut HyperlinkTarget, remap: &HashMap<RelId, RelId>) {
    if let HyperlinkTarget::External(ref id) = target {
        if let Some(new) = remap.get(id) {
            *target = HyperlinkTarget::External(new.clone());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::Dimension;
    use crate::model::geometry::{EdgeInsets, Size};
    use crate::model::*;

    fn remap_one(from: &str, to: &str) -> HashMap<RelId, RelId> {
        let mut m = HashMap::new();
        m.insert(RelId::new(from), RelId::new(to));
        m
    }

    fn picture_with_blip(rel_id: &str) -> Image {
        Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "img".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: Some(GraphicContent::Picture(Picture {
                nv_pic_pr: NvPicProperties {
                    cnv_pr: DocProperties {
                        id: 1,
                        name: "pic".into(),
                        description: None,
                        hidden: None,
                        title: None,
                    },
                    cnv_pic_pr: None,
                },
                blip_fill: BlipFill {
                    rotate_with_shape: None,
                    dpi: None,
                    blip: Some(Blip {
                        embed: Some(RelId::new(rel_id)),
                        link: None,
                        compression: None,
                    }),
                    src_rect: None,
                    fill_kind: BlipFillKind::Unspecified,
                },
                shape_properties: None,
            })),
            placement: ImagePlacement::Inline {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
            },
        }
    }

    fn paragraph_with_image(image: Image) -> Block {
        Block::Paragraph(Box::new(Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: vec![Inline::Image(Box::new(image))],
            rsids: ParagraphRevisionIds::default(),
        }))
    }

    #[test]
    fn rewrites_picture_blip_embed() {
        let mut blocks = vec![paragraph_with_image(picture_with_blip("rId1"))];
        let remap = remap_one("rId1", "header3.xml::rId1");

        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        let Block::Paragraph(p) = &blocks[0] else {
            panic!();
        };
        let Inline::Image(img) = &p.content[0] else {
            panic!();
        };
        let new = crate::render::resolve::images::extract_image_rel_id(img).unwrap();
        assert_eq!(new.as_str(), "header3.xml::rId1");
    }

    #[test]
    fn leaves_unmapped_rel_ids_unchanged() {
        let mut blocks = vec![paragraph_with_image(picture_with_blip("rId99"))];
        let remap = remap_one("rId1", "synth");

        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        let Block::Paragraph(p) = &blocks[0] else {
            panic!();
        };
        let Inline::Image(img) = &p.content[0] else {
            panic!();
        };
        let v = crate::render::resolve::images::extract_image_rel_id(img).unwrap();
        assert_eq!(v.as_str(), "rId99");
    }

    #[test]
    fn empty_remap_is_a_noop() {
        let mut blocks = vec![paragraph_with_image(picture_with_blip("rId1"))];
        let remap = HashMap::new();
        rewrite_part_rels_in_blocks(&mut blocks, &remap);
        // No panic, no change — the image still references rId1.
        let Block::Paragraph(p) = &blocks[0] else {
            panic!();
        };
        let Inline::Image(img) = &p.content[0] else {
            panic!();
        };
        let v = crate::render::resolve::images::extract_image_rel_id(img).unwrap();
        assert_eq!(v.as_str(), "rId1");
    }

    #[test]
    fn rewrites_blip_link_too() {
        let mut blip = Blip {
            embed: None,
            link: Some(RelId::new("rId7")),
            compression: None,
        };
        let remap = remap_one("rId7", "linked");
        rewrite_blip(&mut blip, &remap);
        assert_eq!(blip.link.unwrap().as_str(), "linked");
    }

    #[test]
    fn rewrites_vml_image_data_rel_id() {
        let mut data = VmlImageData {
            rel_id: Some(RelId::new("rId2")),
            title: None,
        };
        let remap = remap_one("rId2", "synth_vml");
        rewrite_vml_image_data(&mut data, &remap);
        assert_eq!(data.rel_id.unwrap().as_str(), "synth_vml");
    }

    #[test]
    fn recurses_into_table_cells() {
        let img = picture_with_blip("rId1");
        let cell = TableCell {
            properties: TableCellProperties::default(),
            content: vec![paragraph_with_image(img)],
        };
        let row = TableRow {
            properties: TableRowProperties::default(),
            cells: vec![cell],
            rsids: TableRowRevisionIds::default(),
            property_exceptions: None,
        };
        let mut blocks = vec![Block::Table(Box::new(Table {
            properties: TableProperties::default(),
            grid: vec![],
            rows: vec![row],
        }))];
        let remap = remap_one("rId1", "synth");
        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        let Block::Table(t) = &blocks[0] else {
            panic!();
        };
        let Block::Paragraph(p) = &t.rows[0].cells[0].content[0] else {
            panic!();
        };
        let Inline::Image(img) = &p.content[0] else {
            panic!();
        };
        let v = crate::render::resolve::images::extract_image_rel_id(img).unwrap();
        assert_eq!(v.as_str(), "synth");
    }

    // ── §17.17.1 wsp:txbxContent — text inside DrawingML shapes ─────────

    /// Build a `WordProcessingShape` with the given `txbx_content` and
    /// the given optional `shape_properties`. All other fields default
    /// to None / empty.
    fn shape_with_text(
        txbx_content: Vec<Block>,
        shape_properties: Option<ShapeProperties>,
    ) -> Image {
        Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "shape".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: Some(GraphicContent::WordProcessingShape(WordProcessingShape {
                cnv_pr: None,
                shape_properties,
                style_line_ref: None,
                style_effect_ref: None,
                body_pr: None,
                txbx_content,
            })),
            placement: ImagePlacement::Inline {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
            },
        }
    }

    #[test]
    fn recurses_into_word_processing_shape_txbx_content() {
        // §17.17.1: a WordProcessingShape may host its own paragraph
        // tree via `<wsp:txbxContent>`. Image rIds inside that subtree
        // belong to the same part as the outer shape and must be
        // rewritten alongside it.
        let inner_image = picture_with_blip("rId1");
        let shape = shape_with_text(vec![paragraph_with_image(inner_image)], None);
        let mut blocks = vec![paragraph_with_image(shape)];
        let remap = remap_one("rId1", "synth_inner");

        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        // Drill through: outer paragraph → outer shape → txbx_content
        // → inner paragraph → inner Picture → blip embed.
        let Block::Paragraph(outer_p) = &blocks[0] else {
            panic!();
        };
        let Inline::Image(outer_img) = &outer_p.content[0] else {
            panic!();
        };
        let Some(GraphicContent::WordProcessingShape(wsp)) = &outer_img.graphic else {
            panic!();
        };
        let Block::Paragraph(inner_p) = &wsp.txbx_content[0] else {
            panic!();
        };
        let Inline::Image(inner) = &inner_p.content[0] else {
            panic!();
        };
        let v = crate::render::resolve::images::extract_image_rel_id(inner).unwrap();
        assert_eq!(v.as_str(), "synth_inner");
    }

    // ── §20.1.8.14 blip image fill on shape body / outline ──────────────

    /// Build a `ShapeProperties` whose body fill is a Blip referencing
    /// `rel_id`. All other properties are empty.
    fn shape_props_with_blip_fill(rel_id: &str) -> ShapeProperties {
        ShapeProperties {
            bw_mode: None,
            transform: None,
            geometry: None,
            fill: Some(DrawingFill::Blip(BlipFill {
                rotate_with_shape: None,
                dpi: None,
                blip: Some(Blip {
                    embed: Some(RelId::new(rel_id)),
                    link: None,
                    compression: None,
                }),
                src_rect: None,
                fill_kind: BlipFillKind::Unspecified,
            })),
            outline: None,
            effect_list: None,
        }
    }

    #[test]
    fn rewrites_blip_image_fill_on_shape_body() {
        // §20.1.8.14: `DrawingFill::Blip` lets a shape use an image as
        // its fill. Whether or not the renderer paints it today, the
        // rId still belongs to the part's namespace and must be
        // remapped at parse time so future renderer wiring sees a
        // valid lookup key.
        let shape = shape_with_text(vec![], Some(shape_props_with_blip_fill("rId1")));
        let mut blocks = vec![paragraph_with_image(shape)];
        let remap = remap_one("rId1", "synth_shape_fill");

        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        let Block::Paragraph(p) = &blocks[0] else {
            panic!();
        };
        let Inline::Image(img) = &p.content[0] else {
            panic!();
        };
        let Some(GraphicContent::WordProcessingShape(wsp)) = &img.graphic else {
            panic!();
        };
        let Some(props) = &wsp.shape_properties else {
            panic!();
        };
        let Some(DrawingFill::Blip(blip_fill)) = &props.fill else {
            panic!();
        };
        let blip = blip_fill.blip.as_ref().unwrap();
        assert_eq!(blip.embed.as_ref().unwrap().as_str(), "synth_shape_fill");
    }

    #[test]
    fn rewrites_blip_image_fill_on_shape_outline() {
        // Outlines accept the same EG_FillProperties choice as bodies.
        let outline = Outline {
            width: None,
            cap: None,
            compound: None,
            alignment: None,
            fill: Some(DrawingFill::Blip(BlipFill {
                rotate_with_shape: None,
                dpi: None,
                blip: Some(Blip {
                    embed: Some(RelId::new("rId7")),
                    link: None,
                    compression: None,
                }),
                src_rect: None,
                fill_kind: BlipFillKind::Unspecified,
            })),
            dash: None,
            join: None,
            head_end: None,
            tail_end: None,
        };
        let props = ShapeProperties {
            bw_mode: None,
            transform: None,
            geometry: None,
            fill: None,
            outline: Some(outline),
            effect_list: None,
        };
        let shape = shape_with_text(vec![], Some(props));
        let mut blocks = vec![paragraph_with_image(shape)];
        let remap = remap_one("rId7", "synth_outline");

        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        let Block::Paragraph(p) = &blocks[0] else {
            panic!();
        };
        let Inline::Image(img) = &p.content[0] else {
            panic!();
        };
        let Some(GraphicContent::WordProcessingShape(wsp)) = &img.graphic else {
            panic!();
        };
        let outline = wsp
            .shape_properties
            .as_ref()
            .unwrap()
            .outline
            .as_ref()
            .unwrap();
        let Some(DrawingFill::Blip(blip_fill)) = &outline.fill else {
            panic!();
        };
        let blip = blip_fill.blip.as_ref().unwrap();
        assert_eq!(blip.embed.as_ref().unwrap().as_str(), "synth_outline");
    }

    // ── VML §14.1.2.22 v:textbox content recursion ──────────────────────

    #[test]
    fn recurses_into_vml_text_box_content() {
        let inner_image = picture_with_blip("rId1");
        let pict = Pict {
            shape_type: None,
            primitives: vec![VmlPrimitive::Shape(VmlShape {
                common: VmlCommonAttrs {
                    text_box: Some(VmlTextBox {
                        style: VmlStyle::default(),
                        inset: None,
                        content: vec![paragraph_with_image(inner_image)],
                    }),
                    ..VmlCommonAttrs::default()
                },
                shape_type_ref: None,
                vml_path: None,
            })],
        };
        let mut blocks = vec![Block::Paragraph(Box::new(Paragraph {
            style_id: None,
            properties: ParagraphProperties::default(),
            mark_run_properties: None,
            content: vec![Inline::Pict(pict)],
            rsids: ParagraphRevisionIds::default(),
        }))];
        let remap = remap_one("rId1", "synth_vml_box");

        rewrite_part_rels_in_blocks(&mut blocks, &remap);

        let Block::Paragraph(p) = &blocks[0] else {
            panic!();
        };
        let Inline::Pict(pict) = &p.content[0] else {
            panic!();
        };
        let shape = pict.shapes().next().unwrap();
        let tb = shape.common.text_box.as_ref().unwrap();
        let Block::Paragraph(inner_p) = &tb.content[0] else {
            panic!();
        };
        let Inline::Image(img) = &inner_p.content[0] else {
            panic!();
        };
        let v = crate::render::resolve::images::extract_image_rel_id(img).unwrap();
        assert_eq!(v.as_str(), "synth_vml_box");
    }
}
