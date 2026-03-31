//! Hierarchical parsing of DrawingML elements (inline images, anchor images).
//!
//! Follows the actual XML nesting per OOXML spec rather than flat event scanning.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::{Dimension, Emu};
use crate::docx::error::{ParseError, Result};
use crate::docx::geometry::{EdgeInsets, Offset, Size};
use crate::docx::model::*;
use crate::docx::xml;

// ── wp:inline (§20.4.2.8) ──────────────────────────────────────────────────

/// Parse `wp:inline` into an `Image` with `ImagePlacement::Inline`.
pub fn parse_inline_image(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Option<Image>> {
    let dist_t = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
    let dist_b = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
    let dist_l = xml::optional_attr_i64(e, b"distL")?.unwrap_or(0);
    let dist_r = xml::optional_attr_i64(e, b"distR")?.unwrap_or(0);
    let distance = EdgeInsets::new(
        Dimension::new(dist_t),
        Dimension::new(dist_r),
        Dimension::new(dist_b),
        Dimension::new(dist_l),
    );

    let mut extent = Size::ZERO;
    let mut effect_extent = None;
    let mut doc_properties = None;
    let mut graphic_frame_locks = None;
    let mut picture = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvGraphicFramePr" => {
                        graphic_frame_locks = Some(parse_cnv_graphic_frame_pr(reader, buf)?);
                    }
                    b"graphic" => {
                        picture = parse_graphic(reader, buf)?;
                    }
                    b"docPr" => {
                        doc_properties = Some(parse_doc_properties(e, Some(reader), Some(buf))?);
                    }
                    _ => {
                        xml::warn_unsupported_element("inline-image", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"extent" => {
                        extent = parse_positive_size_2d(e)?;
                    }
                    b"effectExtent" => {
                        effect_extent = Some(parse_effect_extent(e)?);
                    }
                    b"docPr" => {
                        doc_properties = Some(parse_doc_properties(e, None, None)?);
                    }
                    _ => xml::warn_unsupported_element("inline-image", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"inline" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"inline")),
            _ => {}
        }
    }

    let Some(doc_props) = doc_properties else {
        return Err(ParseError::MissingElement {
            parent: "wp:inline".into(),
            child: "wp:docPr".into(),
        });
    };

    Ok(Some(Image {
        extent,
        effect_extent,
        doc_properties: doc_props,
        graphic_frame_locks,
        graphic: picture,
        placement: ImagePlacement::Inline { distance },
    }))
}

// ── Shared element parsers ──────────────────────────────────────────────────

/// §20.4.2.7 / §20.1.7.3: parse `cx`, `cy` into `Size<Emu>`.
fn parse_positive_size_2d(e: &BytesStart<'_>) -> Result<Size<Emu>> {
    let cx = xml::optional_attr_i64(e, b"cx")?.ok_or_else(|| ParseError::MissingAttribute {
        element: "extent".into(),
        attr: "cx".into(),
    })?;
    let cy = xml::optional_attr_i64(e, b"cy")?.ok_or_else(|| ParseError::MissingAttribute {
        element: "extent".into(),
        attr: "cy".into(),
    })?;
    Ok(Size::new(Dimension::new(cx), Dimension::new(cy)))
}

/// §20.4.2.6: parse `wp:effectExtent` attributes (l, t, r, b).
fn parse_effect_extent(e: &BytesStart<'_>) -> Result<EdgeInsets<Emu>> {
    let l = xml::optional_attr_i64(e, b"l")?.unwrap_or(0);
    let t = xml::optional_attr_i64(e, b"t")?.unwrap_or(0);
    let r = xml::optional_attr_i64(e, b"r")?.unwrap_or(0);
    let b = xml::optional_attr_i64(e, b"b")?.unwrap_or(0);
    Ok(EdgeInsets::new(
        Dimension::new(t),
        Dimension::new(r),
        Dimension::new(b),
        Dimension::new(l),
    ))
}

/// §20.1.2.2.8 CT_NonVisualDrawingProps — parse `docPr` or `cNvPr` attributes.
/// If the element is `Start` (has children), skips to its end.
fn parse_doc_properties(
    e: &BytesStart<'_>,
    reader: Option<&mut Reader<&[u8]>>,
    buf: Option<&mut Vec<u8>>,
) -> Result<DocProperties> {
    let id = xml::optional_attr_u32(e, b"id")?.unwrap_or(0);
    let name = xml::optional_attr(e, b"name")?.unwrap_or_default();
    let description = xml::optional_attr(e, b"descr")?;
    let hidden = xml::optional_attr_bool(e, b"hidden")?;
    let title = xml::optional_attr(e, b"title")?;

    // If Start element, skip children (extensions etc.)
    if let (Some(reader), Some(buf)) = (reader, buf) {
        xml::skip_to_end(reader, buf, b"docPr")?;
    }

    Ok(DocProperties {
        id,
        name,
        description,
        hidden,
        title,
    })
}

// ── wp:cNvGraphicFramePr (§20.4.2.4) ───────────────────────────────────────

fn parse_cnv_graphic_frame_pr(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<GraphicFrameLocks> {
    let mut locks = GraphicFrameLocks {
        no_change_aspect: None,
        no_drilldown: None,
        no_grp: None,
        no_move: None,
        no_resize: None,
        no_select: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"graphicFrameLocks" => {
                        locks.no_change_aspect = xml::optional_attr_bool(e, b"noChangeAspect")?;
                        locks.no_drilldown = xml::optional_attr_bool(e, b"noDrilldown")?;
                        locks.no_grp = xml::optional_attr_bool(e, b"noGrp")?;
                        locks.no_move = xml::optional_attr_bool(e, b"noMove")?;
                        locks.no_resize = xml::optional_attr_bool(e, b"noResize")?;
                        locks.no_select = xml::optional_attr_bool(e, b"noSelect")?;
                    }
                    _ => xml::warn_unsupported_element("cNvGraphicFramePr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"cNvGraphicFramePr" => {
                break
            }
            Event::Eof => return Err(xml::unexpected_eof(b"cNvGraphicFramePr")),
            _ => {}
        }
    }

    Ok(locks)
}

// ── a:graphic / a:graphicData (§20.1.2.2.16, §20.1.2.2.17) ────────────────

/// Parse `a:graphic` → `a:graphicData` → content.
fn parse_graphic(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<GraphicContent>> {
    let mut content = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"graphicData" => {
                        content = parse_graphic_data(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("graphic", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"graphic" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"graphic")),
            _ => {}
        }
    }

    Ok(content)
}

/// Parse `a:graphicData` children. Dispatches based on content type.
fn parse_graphic_data(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Option<GraphicContent>> {
    let mut content = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"pic" => {
                        content = Some(GraphicContent::Picture(parse_picture(reader, buf)?));
                    }
                    b"wsp" => {
                        content = Some(GraphicContent::WordProcessingShape(
                            parse_word_processing_shape(reader, buf)?,
                        ));
                    }
                    _ => {
                        xml::warn_unsupported_element("graphicData", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"graphicData" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"graphicData")),
            _ => {}
        }
    }

    Ok(content)
}

// ── wps:wsp (§14.5) ────────────────────────────────────────────────────────

/// Parse `wps:wsp` — Word Processing Shape.
fn parse_word_processing_shape(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<WordProcessingShape> {
    let mut cnv_pr = None;
    let mut shape_properties = None;
    let mut body_pr = None;
    let mut txbx_content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvPr" => {
                        cnv_pr = Some(parse_cnv_pr(e, reader, buf)?);
                    }
                    b"spPr" => {
                        shape_properties = Some(parse_shape_properties(e, reader, buf)?);
                    }
                    b"bodyPr" => {
                        body_pr = Some(parse_body_properties(e, reader, buf)?);
                    }
                    b"txbx" => {
                        txbx_content = parse_wsp_txbx(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("wsp", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvPr" => {
                        cnv_pr = Some(parse_cnv_pr_attrs(e)?);
                    }
                    b"cNvSpPr" | b"cNvCnPr" => {} // non-visual shape props — attrs only
                    b"bodyPr" => {
                        body_pr = Some(parse_body_properties_attrs(e)?);
                    }
                    _ => xml::warn_unsupported_element("wsp", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"wsp" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"wsp")),
            _ => {}
        }
    }

    Ok(WordProcessingShape {
        cnv_pr,
        shape_properties,
        body_pr,
        txbx_content,
    })
}

/// Parse `wps:txbx` — contains `w:txbxContent`.
fn parse_wsp_txbx(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<Block>> {
    let mut blocks = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"txbxContent" => {
                        let (content, _) = crate::docx::parse::body::parse_block_content_public(
                            reader,
                            buf,
                            b"txbxContent",
                        )?;
                        blocks = content;
                    }
                    _ => {
                        xml::warn_unsupported_element("txbx", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"txbx" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"txbx")),
            _ => {}
        }
    }

    Ok(blocks)
}

/// §20.1.2.1.1: parse `a:bodyPr` attributes only.
fn parse_body_properties_attrs(e: &BytesStart<'_>) -> Result<BodyProperties> {
    Ok(BodyProperties {
        rotation: xml::optional_attr_i64(e, b"rot")?.map(Dimension::new),
        vert: xml::optional_attr(e, b"vert")?
            .map(|s| parse_text_vertical_type(&s))
            .transpose()?,
        wrap: xml::optional_attr(e, b"wrap")?
            .map(|s| parse_text_wrapping_type(&s))
            .transpose()?,
        left_inset: xml::optional_attr_i64(e, b"lIns")?.map(Dimension::new),
        top_inset: xml::optional_attr_i64(e, b"tIns")?.map(Dimension::new),
        right_inset: xml::optional_attr_i64(e, b"rIns")?.map(Dimension::new),
        bottom_inset: xml::optional_attr_i64(e, b"bIns")?.map(Dimension::new),
        anchor: xml::optional_attr(e, b"anchor")?
            .map(|s| parse_text_anchoring_type(&s))
            .transpose()?,
        auto_fit: None,
    })
}

/// §20.1.10.82 ST_TextVerticalType.
fn parse_text_vertical_type(val: &str) -> Result<TextVerticalType> {
    match val {
        "horz" => Ok(TextVerticalType::Horz),
        "vert" => Ok(TextVerticalType::Vert),
        "vert270" => Ok(TextVerticalType::Vert270),
        "wordArtVert" => Ok(TextVerticalType::WordArtVert),
        "eaVert" => Ok(TextVerticalType::EaVert),
        "mongolianVert" => Ok(TextVerticalType::MongolianVert),
        "wordArtVertRtl" => Ok(TextVerticalType::WordArtVertRtl),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "vert".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.82 ST_TextVerticalType".into(),
        }),
    }
}

/// §20.1.10.85 ST_TextWrappingType.
fn parse_text_wrapping_type(val: &str) -> Result<TextWrappingType> {
    match val {
        "none" => Ok(TextWrappingType::None),
        "square" => Ok(TextWrappingType::Square),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "wrap".into(),
            value: other.into(),
            reason: "expected none or square per §20.1.10.85 ST_TextWrappingType".into(),
        }),
    }
}

/// §20.1.10.59 ST_TextAnchoringType.
fn parse_text_anchoring_type(val: &str) -> Result<TextAnchoringType> {
    match val {
        "t" => Ok(TextAnchoringType::Top),
        "ctr" => Ok(TextAnchoringType::Center),
        "b" => Ok(TextAnchoringType::Bottom),
        "just" => Ok(TextAnchoringType::Justified),
        "dist" => Ok(TextAnchoringType::Distributed),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "anchor".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.59 ST_TextAnchoringType".into(),
        }),
    }
}

/// §20.1.2.1.1: parse `a:bodyPr` with children.
fn parse_body_properties(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<BodyProperties> {
    let mut bp = parse_body_properties_attrs(e)?;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"noAutofit" => bp.auto_fit = Some(TextAutoFit::NoAutoFit),
                    b"normAutofit" => bp.auto_fit = Some(TextAutoFit::NormalAutoFit),
                    b"spAutoFit" => bp.auto_fit = Some(TextAutoFit::SpAutoFit),
                    _ => {} // skip prstTxWarp, scene3d, etc.
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"bodyPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"bodyPr")),
            _ => {}
        }
    }

    Ok(bp)
}

// ── pic:pic (§19.3.1.37) ───────────────────────────────────────────────────

fn parse_picture(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Picture> {
    let mut nv_pic_pr = None;
    let mut blip_fill = None;
    let mut shape_properties = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"nvPicPr" => {
                        nv_pic_pr = Some(parse_nv_pic_pr(reader, buf)?);
                    }
                    b"blipFill" => {
                        blip_fill = Some(parse_blip_fill(e, reader, buf)?);
                    }
                    b"spPr" => {
                        shape_properties = Some(parse_shape_properties(e, reader, buf)?);
                    }
                    _ => {
                        xml::warn_unsupported_element("pic", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pic" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"pic")),
            _ => {}
        }
    }

    let nv_pic_pr = nv_pic_pr.ok_or_else(|| ParseError::MissingElement {
        parent: "pic:pic".into(),
        child: "pic:nvPicPr".into(),
    })?;
    let blip_fill = blip_fill.ok_or_else(|| ParseError::MissingElement {
        parent: "pic:pic".into(),
        child: "pic:blipFill".into(),
    })?;

    Ok(Picture {
        nv_pic_pr,
        blip_fill,
        shape_properties,
    })
}

// ── pic:nvPicPr (§19.3.1.32) ───────────────────────────────────────────────

fn parse_nv_pic_pr(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<NvPicProperties> {
    let mut cnv_pr = None;
    let mut cnv_pic_pr = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvPr" => {
                        cnv_pr = Some(parse_cnv_pr(e, reader, buf)?);
                    }
                    b"cNvPicPr" => {
                        cnv_pic_pr = Some(parse_cnv_pic_pr(e, reader, buf)?);
                    }
                    _ => {
                        xml::warn_unsupported_element("nvPicPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvPr" => {
                        cnv_pr = Some(parse_cnv_pr_attrs(e)?);
                    }
                    b"cNvPicPr" => {
                        cnv_pic_pr = Some(CnvPicProperties {
                            prefer_relative_resize: xml::optional_attr_bool(
                                e,
                                b"preferRelativeResize",
                            )?,
                            pic_locks: None,
                        });
                    }
                    _ => xml::warn_unsupported_element("nvPicPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"nvPicPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"nvPicPr")),
            _ => {}
        }
    }

    let cnv_pr = cnv_pr.ok_or_else(|| ParseError::MissingElement {
        parent: "pic:nvPicPr".into(),
        child: "pic:cNvPr".into(),
    })?;

    Ok(NvPicProperties { cnv_pr, cnv_pic_pr })
}

/// §20.1.2.2.8: parse `pic:cNvPr` attributes only.
fn parse_cnv_pr_attrs(e: &BytesStart<'_>) -> Result<DocProperties> {
    Ok(DocProperties {
        id: xml::optional_attr_u32(e, b"id")?.unwrap_or(0),
        name: xml::optional_attr(e, b"name")?.unwrap_or_default(),
        description: xml::optional_attr(e, b"descr")?,
        hidden: xml::optional_attr_bool(e, b"hidden")?,
        title: xml::optional_attr(e, b"title")?,
    })
}

/// §20.1.2.2.8: parse `pic:cNvPr` with children.
fn parse_cnv_pr(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<DocProperties> {
    let props = parse_cnv_pr_attrs(e)?;
    // cNvPr can have extension children (hlinkClick, hlinkHover, extLst) — skip.
    xml::skip_to_end(reader, buf, b"cNvPr")?;
    Ok(props)
}

// ── pic:cNvPicPr (§19.3.1.4) ───────────────────────────────────────────────

fn parse_cnv_pic_pr(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<CnvPicProperties> {
    let prefer_relative_resize = xml::optional_attr_bool(e, b"preferRelativeResize")?;
    let mut pic_locks = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"picLocks" => {
                        pic_locks = Some(parse_pic_locks(e)?);
                    }
                    _ => xml::warn_unsupported_element("cNvPicPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"cNvPicPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"cNvPicPr")),
            _ => {}
        }
    }

    Ok(CnvPicProperties {
        prefer_relative_resize,
        pic_locks,
    })
}

/// §20.1.2.2.31: parse `a:picLocks` attributes.
fn parse_pic_locks(e: &BytesStart<'_>) -> Result<PicLocks> {
    Ok(PicLocks {
        no_change_aspect: xml::optional_attr_bool(e, b"noChangeAspect")?,
        no_crop: xml::optional_attr_bool(e, b"noCrop")?,
        no_resize: xml::optional_attr_bool(e, b"noResize")?,
        no_move: xml::optional_attr_bool(e, b"noMove")?,
        no_rot: xml::optional_attr_bool(e, b"noRot")?,
        no_select: xml::optional_attr_bool(e, b"noSelect")?,
        no_edit_points: xml::optional_attr_bool(e, b"noEditPoints")?,
        no_adjust_handles: xml::optional_attr_bool(e, b"noAdjustHandles")?,
        no_change_arrowheads: xml::optional_attr_bool(e, b"noChangeArrowheads")?,
        no_change_shape_type: xml::optional_attr_bool(e, b"noChangeShapeType")?,
        no_grp: xml::optional_attr_bool(e, b"noGrp")?,
    })
}

// ── pic:blipFill (§20.1.8.14) ──────────────────────────────────────────────

fn parse_blip_fill(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<BlipFill> {
    let rotate_with_shape = xml::optional_attr_bool(start, b"rotWithShape")?;
    let dpi = xml::optional_attr_u32(start, b"dpi")?;
    let mut blip = None;
    let mut src_rect = None;
    let mut stretch = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"blip" => {
                        blip = Some(parse_blip(e, reader, buf)?);
                    }
                    b"stretch" => {
                        stretch = Some(parse_stretch(reader, buf)?);
                    }
                    _ => {
                        xml::warn_unsupported_element("blipFill", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"blip" => {
                        blip = Some(parse_blip_attrs(e)?);
                    }
                    b"srcRect" => {
                        src_rect = Some(parse_relative_rect(e)?);
                    }
                    _ => xml::warn_unsupported_element("blipFill", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"blipFill" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"blipFill")),
            _ => {}
        }
    }

    Ok(BlipFill {
        rotate_with_shape,
        dpi,
        blip,
        src_rect,
        stretch,
    })
}

// ── a:blip (§20.1.8.13) ────────────────────────────────────────────────────

fn parse_blip_attrs(e: &BytesStart<'_>) -> Result<Blip> {
    Ok(Blip {
        embed: xml::optional_attr(e, b"embed")?.map(RelId::new),
        link: xml::optional_attr(e, b"link")?.map(RelId::new),
        compression: xml::optional_attr(e, b"cstate")?
            .map(|s| parse_blip_compression(&s))
            .transpose()?,
    })
}

fn parse_blip(e: &BytesStart<'_>, reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Blip> {
    let blip = parse_blip_attrs(e)?;
    // Children are extension lists (extLst) — skip them.
    xml::skip_to_end(reader, buf, b"blip")?;
    Ok(blip)
}

fn parse_blip_compression(val: &str) -> Result<BlipCompression> {
    match val {
        "email" => Ok(BlipCompression::Email),
        "hqprint" => Ok(BlipCompression::Hqprint),
        "none" => Ok(BlipCompression::None),
        "print" => Ok(BlipCompression::Print),
        "screen" => Ok(BlipCompression::Screen),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "cstate".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.7 ST_BlipCompression".into(),
        }),
    }
}

// ── a:srcRect / a:fillRect (§20.1.10.48) ───────────────────────────────────

fn parse_relative_rect(e: &BytesStart<'_>) -> Result<RelativeRect> {
    Ok(RelativeRect {
        left: xml::optional_attr_i64(e, b"l")?.map(Dimension::new),
        top: xml::optional_attr_i64(e, b"t")?.map(Dimension::new),
        right: xml::optional_attr_i64(e, b"r")?.map(Dimension::new),
        bottom: xml::optional_attr_i64(e, b"b")?.map(Dimension::new),
    })
}

// ── a:stretch (§20.1.8.56) ─────────────────────────────────────────────────

fn parse_stretch(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<StretchFill> {
    let mut fill_rect = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"fillRect" => {
                        fill_rect = Some(parse_relative_rect(e)?);
                    }
                    _ => xml::warn_unsupported_element("stretch", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"stretch" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"stretch")),
            _ => {}
        }
    }

    Ok(StretchFill { fill_rect })
}

// ── pic:spPr (§20.1.2.2.35) ────────────────────────────────────────────────

fn parse_shape_properties(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ShapeProperties> {
    let bw_mode = xml::optional_attr(start, b"bwMode")?
        .map(|s| parse_black_white_mode(&s))
        .transpose()?;
    let mut transform = None;
    let mut preset_geometry = None;
    let mut fill = None;
    let mut outline = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"xfrm" => {
                        transform = Some(parse_transform_2d(e, reader, buf)?);
                    }
                    b"prstGeom" => {
                        preset_geometry = Some(parse_preset_geometry(e, reader, buf)?);
                    }
                    b"ln" => {
                        outline = Some(parse_outline(e, reader, buf)?);
                    }
                    b"custGeom" => {
                        xml::warn_unsupported_element("spPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    // Fill group — complex, skip children for now.
                    b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => {
                        xml::warn_unsupported_element("spPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    b"effectLst" | b"effectDag" | b"scene3d" | b"sp3d" | b"extLst" => {
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("spPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"noFill" => {
                        fill = Some(DrawingFill::NoFill);
                    }
                    _ => xml::warn_unsupported_element("spPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"spPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"spPr")),
            _ => {}
        }
    }

    Ok(ShapeProperties {
        bw_mode,
        transform,
        preset_geometry,
        fill,
        outline,
    })
}

fn parse_black_white_mode(val: &str) -> Result<BlackWhiteMode> {
    match val {
        "auto" => Ok(BlackWhiteMode::Auto),
        "black" => Ok(BlackWhiteMode::Black),
        "blackGray" => Ok(BlackWhiteMode::BlackGray),
        "blackWhite" => Ok(BlackWhiteMode::BlackWhite),
        "clr" => Ok(BlackWhiteMode::Clr),
        "gray" => Ok(BlackWhiteMode::Gray),
        "grayWhite" => Ok(BlackWhiteMode::GrayWhite),
        "hidden" => Ok(BlackWhiteMode::Hidden),
        "invGray" => Ok(BlackWhiteMode::InvGray),
        "ltGray" => Ok(BlackWhiteMode::LtGray),
        "white" => Ok(BlackWhiteMode::White),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "bwMode".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.10 ST_BlackWhiteMode".into(),
        }),
    }
}

// ── a:xfrm (§20.1.7.6) ────────────────────────────────────────────────────

fn parse_transform_2d(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Transform2D> {
    let rotation = xml::optional_attr_i64(start, b"rot")?.map(Dimension::new);
    let flip_h = xml::optional_attr_bool(start, b"flipH")?;
    let flip_v = xml::optional_attr_bool(start, b"flipV")?;
    let mut offset = None;
    let mut extent = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"off" => {
                        let x = xml::optional_attr_i64(e, b"x")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "a:off".into(),
                                attr: "x".into(),
                            }
                        })?;
                        let y = xml::optional_attr_i64(e, b"y")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "a:off".into(),
                                attr: "y".into(),
                            }
                        })?;
                        offset = Some(Offset::new(Dimension::new(x), Dimension::new(y)));
                    }
                    b"ext" => {
                        extent = Some(parse_positive_size_2d(e)?);
                    }
                    _ => xml::warn_unsupported_element("xfrm", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"xfrm" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"xfrm")),
            _ => {}
        }
    }

    Ok(Transform2D {
        rotation,
        flip_h,
        flip_v,
        offset,
        extent,
    })
}

// ── a:prstGeom (§20.1.9.18) ────────────────────────────────────────────────

fn parse_preset_geometry(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<PresetGeometryDef> {
    let prst = xml::optional_attr(start, b"prst")?.ok_or_else(|| ParseError::MissingAttribute {
        element: "a:prstGeom".into(),
        attr: "prst".into(),
    })?;
    let preset = parse_preset_shape_type(&prst);
    let mut adjust_values = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"avLst" => {
                        adjust_values = parse_geom_guide_list(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("prstGeom", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"avLst" => {} // empty adjustment list
                    _ => xml::warn_unsupported_element("prstGeom", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"prstGeom" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"prstGeom")),
            _ => {}
        }
    }

    Ok(PresetGeometryDef {
        preset,
        adjust_values,
    })
}

fn parse_geom_guide_list(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<GeomGuide>> {
    let mut guides = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"gd" => {
                let name = xml::optional_attr(e, b"name")?.unwrap_or_default();
                let formula = xml::optional_attr(e, b"fmla")?.unwrap_or_default();
                guides.push(GeomGuide { name, formula });
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"avLst" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"avLst")),
            _ => {}
        }
    }

    Ok(guides)
}

fn parse_preset_shape_type(val: &str) -> PresetShapeType {
    match val {
        "rect" => PresetShapeType::Rect,
        "roundRect" => PresetShapeType::RoundRect,
        "ellipse" => PresetShapeType::Ellipse,
        "triangle" => PresetShapeType::Triangle,
        "rtTriangle" => PresetShapeType::RtTriangle,
        "diamond" => PresetShapeType::Diamond,
        "parallelogram" => PresetShapeType::Parallelogram,
        "trapezoid" => PresetShapeType::Trapezoid,
        "pentagon" => PresetShapeType::Pentagon,
        "hexagon" => PresetShapeType::Hexagon,
        "octagon" => PresetShapeType::Octagon,
        "star4" => PresetShapeType::Star4,
        "star5" => PresetShapeType::Star5,
        "star6" => PresetShapeType::Star6,
        "star8" => PresetShapeType::Star8,
        "star10" => PresetShapeType::Star10,
        "star12" => PresetShapeType::Star12,
        "star16" => PresetShapeType::Star16,
        "star24" => PresetShapeType::Star24,
        "star32" => PresetShapeType::Star32,
        "line" => PresetShapeType::Line,
        "plus" => PresetShapeType::Plus,
        "can" => PresetShapeType::Can,
        "cube" => PresetShapeType::Cube,
        "donut" => PresetShapeType::Donut,
        "noSmoking" => PresetShapeType::NoSmoking,
        "blockArc" => PresetShapeType::BlockArc,
        "heart" => PresetShapeType::Heart,
        "sun" => PresetShapeType::Sun,
        "moon" => PresetShapeType::Moon,
        "smileyFace" => PresetShapeType::SmileyFace,
        "lightningBolt" => PresetShapeType::LightningBolt,
        "cloud" => PresetShapeType::Cloud,
        "arc" => PresetShapeType::Arc,
        "plaque" => PresetShapeType::Plaque,
        "frame" => PresetShapeType::Frame,
        "bevel" => PresetShapeType::Bevel,
        "foldedCorner" => PresetShapeType::FoldedCorner,
        "chevron" => PresetShapeType::Chevron,
        "homePlate" => PresetShapeType::HomePlate,
        "ribbon" => PresetShapeType::Ribbon,
        "ribbon2" => PresetShapeType::Ribbon2,
        "pie" => PresetShapeType::Pie,
        "pieWedge" => PresetShapeType::PieWedge,
        "chord" => PresetShapeType::Chord,
        "teardrop" => PresetShapeType::Teardrop,
        "arrow" => PresetShapeType::Arrow,
        "leftArrow" => PresetShapeType::LeftArrow,
        "rightArrow" => PresetShapeType::RightArrow,
        "upArrow" => PresetShapeType::UpArrow,
        "downArrow" => PresetShapeType::DownArrow,
        "leftRightArrow" => PresetShapeType::LeftRightArrow,
        "upDownArrow" => PresetShapeType::UpDownArrow,
        "quadArrow" => PresetShapeType::QuadArrow,
        "bentArrow" => PresetShapeType::BentArrow,
        "uturnArrow" => PresetShapeType::UturnArrow,
        "circularArrow" => PresetShapeType::CircularArrow,
        "curvedRightArrow" => PresetShapeType::CurvedRightArrow,
        "curvedLeftArrow" => PresetShapeType::CurvedLeftArrow,
        "curvedUpArrow" => PresetShapeType::CurvedUpArrow,
        "curvedDownArrow" => PresetShapeType::CurvedDownArrow,
        "stripedRightArrow" => PresetShapeType::StripedRightArrow,
        "notchedRightArrow" => PresetShapeType::NotchedRightArrow,
        "bentUpArrow" => PresetShapeType::BentUpArrow,
        "leftUpArrow" => PresetShapeType::LeftUpArrow,
        "leftRightUpArrow" => PresetShapeType::LeftRightUpArrow,
        "leftArrowCallout" => PresetShapeType::LeftArrowCallout,
        "rightArrowCallout" => PresetShapeType::RightArrowCallout,
        "upArrowCallout" => PresetShapeType::UpArrowCallout,
        "downArrowCallout" => PresetShapeType::DownArrowCallout,
        "leftRightArrowCallout" => PresetShapeType::LeftRightArrowCallout,
        "upDownArrowCallout" => PresetShapeType::UpDownArrowCallout,
        "quadArrowCallout" => PresetShapeType::QuadArrowCallout,
        "swooshArrow" => PresetShapeType::SwooshArrow,
        "leftCircularArrow" => PresetShapeType::LeftCircularArrow,
        "leftRightCircularArrow" => PresetShapeType::LeftRightCircularArrow,
        "callout1" => PresetShapeType::Callout1,
        "callout2" => PresetShapeType::Callout2,
        "callout3" => PresetShapeType::Callout3,
        "accentCallout1" => PresetShapeType::AccentCallout1,
        "accentCallout2" => PresetShapeType::AccentCallout2,
        "accentCallout3" => PresetShapeType::AccentCallout3,
        "borderCallout1" => PresetShapeType::BorderCallout1,
        "borderCallout2" => PresetShapeType::BorderCallout2,
        "borderCallout3" => PresetShapeType::BorderCallout3,
        "accentBorderCallout1" => PresetShapeType::AccentBorderCallout1,
        "accentBorderCallout2" => PresetShapeType::AccentBorderCallout2,
        "accentBorderCallout3" => PresetShapeType::AccentBorderCallout3,
        "wedgeRectCallout" => PresetShapeType::WedgeRectCallout,
        "wedgeRoundRectCallout" => PresetShapeType::WedgeRoundRectCallout,
        "wedgeEllipseCallout" => PresetShapeType::WedgeEllipseCallout,
        "cloudCallout" => PresetShapeType::CloudCallout,
        "leftBracket" => PresetShapeType::LeftBracket,
        "rightBracket" => PresetShapeType::RightBracket,
        "leftBrace" => PresetShapeType::LeftBrace,
        "rightBrace" => PresetShapeType::RightBrace,
        "bracketPair" => PresetShapeType::BracketPair,
        "bracePair" => PresetShapeType::BracePair,
        "straightConnector1" => PresetShapeType::StraightConnector1,
        "bentConnector2" => PresetShapeType::BentConnector2,
        "bentConnector3" => PresetShapeType::BentConnector3,
        "bentConnector4" => PresetShapeType::BentConnector4,
        "bentConnector5" => PresetShapeType::BentConnector5,
        "curvedConnector2" => PresetShapeType::CurvedConnector2,
        "curvedConnector3" => PresetShapeType::CurvedConnector3,
        "curvedConnector4" => PresetShapeType::CurvedConnector4,
        "curvedConnector5" => PresetShapeType::CurvedConnector5,
        "flowChartProcess" => PresetShapeType::FlowChartProcess,
        "flowChartDecision" => PresetShapeType::FlowChartDecision,
        "flowChartInputOutput" => PresetShapeType::FlowChartInputOutput,
        "flowChartPredefinedProcess" => PresetShapeType::FlowChartPredefinedProcess,
        "flowChartInternalStorage" => PresetShapeType::FlowChartInternalStorage,
        "flowChartDocument" => PresetShapeType::FlowChartDocument,
        "flowChartMultidocument" => PresetShapeType::FlowChartMultidocument,
        "flowChartTerminator" => PresetShapeType::FlowChartTerminator,
        "flowChartPreparation" => PresetShapeType::FlowChartPreparation,
        "flowChartManualInput" => PresetShapeType::FlowChartManualInput,
        "flowChartManualOperation" => PresetShapeType::FlowChartManualOperation,
        "flowChartConnector" => PresetShapeType::FlowChartConnector,
        "flowChartPunchedCard" => PresetShapeType::FlowChartPunchedCard,
        "flowChartPunchedTape" => PresetShapeType::FlowChartPunchedTape,
        "flowChartSummingJunction" => PresetShapeType::FlowChartSummingJunction,
        "flowChartOr" => PresetShapeType::FlowChartOr,
        "flowChartCollate" => PresetShapeType::FlowChartCollate,
        "flowChartSort" => PresetShapeType::FlowChartSort,
        "flowChartExtract" => PresetShapeType::FlowChartExtract,
        "flowChartMerge" => PresetShapeType::FlowChartMerge,
        "flowChartOfflineStorage" => PresetShapeType::FlowChartOfflineStorage,
        "flowChartOnlineStorage" => PresetShapeType::FlowChartOnlineStorage,
        "flowChartMagneticTape" => PresetShapeType::FlowChartMagneticTape,
        "flowChartMagneticDisk" => PresetShapeType::FlowChartMagneticDisk,
        "flowChartMagneticDrum" => PresetShapeType::FlowChartMagneticDrum,
        "flowChartDisplay" => PresetShapeType::FlowChartDisplay,
        "flowChartDelay" => PresetShapeType::FlowChartDelay,
        "flowChartAlternateProcess" => PresetShapeType::FlowChartAlternateProcess,
        "flowChartOffpageConnector" => PresetShapeType::FlowChartOffpageConnector,
        "actionButtonBlank" => PresetShapeType::ActionButtonBlank,
        "actionButtonHome" => PresetShapeType::ActionButtonHome,
        "actionButtonHelp" => PresetShapeType::ActionButtonHelp,
        "actionButtonInformation" => PresetShapeType::ActionButtonInformation,
        "actionButtonForwardNext" => PresetShapeType::ActionButtonForwardNext,
        "actionButtonBackPrevious" => PresetShapeType::ActionButtonBackPrevious,
        "actionButtonEnd" => PresetShapeType::ActionButtonEnd,
        "actionButtonBeginning" => PresetShapeType::ActionButtonBeginning,
        "actionButtonReturn" => PresetShapeType::ActionButtonReturn,
        "actionButtonDocument" => PresetShapeType::ActionButtonDocument,
        "actionButtonSound" => PresetShapeType::ActionButtonSound,
        "actionButtonMovie" => PresetShapeType::ActionButtonMovie,
        "irregularSeal1" => PresetShapeType::IrregularSeal1,
        "irregularSeal2" => PresetShapeType::IrregularSeal2,
        "wave" => PresetShapeType::Wave,
        "doubleWave" => PresetShapeType::DoubleWave,
        "ellipseRibbon" => PresetShapeType::EllipseRibbon,
        "ellipseRibbon2" => PresetShapeType::EllipseRibbon2,
        "verticalScroll" => PresetShapeType::VerticalScroll,
        "horizontalScroll" => PresetShapeType::HorizontalScroll,
        "leftRightRibbon" => PresetShapeType::LeftRightRibbon,
        "gear6" => PresetShapeType::Gear6,
        "gear9" => PresetShapeType::Gear9,
        "funnel" => PresetShapeType::Funnel,
        "mathPlus" => PresetShapeType::MathPlus,
        "mathMinus" => PresetShapeType::MathMinus,
        "mathMultiply" => PresetShapeType::MathMultiply,
        "mathDivide" => PresetShapeType::MathDivide,
        "mathEqual" => PresetShapeType::MathEqual,
        "mathNotEqual" => PresetShapeType::MathNotEqual,
        "cornerTabs" => PresetShapeType::CornerTabs,
        "squareTabs" => PresetShapeType::SquareTabs,
        "plaqueTabs" => PresetShapeType::PlaqueTabs,
        "chartX" => PresetShapeType::ChartX,
        "chartStar" => PresetShapeType::ChartStar,
        "chartPlus" => PresetShapeType::ChartPlus,
        "halfFrame" => PresetShapeType::HalfFrame,
        "corner" => PresetShapeType::Corner,
        "diagStripe" => PresetShapeType::DiagStripe,
        "nonIsoscelesTrapezoid" => PresetShapeType::NonIsoscelesTrapezoid,
        "heptagon" => PresetShapeType::Heptagon,
        "decagon" => PresetShapeType::Decagon,
        "dodecagon" => PresetShapeType::Dodecagon,
        "round1Rect" => PresetShapeType::Round1Rect,
        "round2SameRect" => PresetShapeType::Round2SameRect,
        "round2DiagRect" => PresetShapeType::Round2DiagRect,
        "snipRoundRect" => PresetShapeType::SnipRoundRect,
        "snip1Rect" => PresetShapeType::Snip1Rect,
        "snip2SameRect" => PresetShapeType::Snip2SameRect,
        "snip2DiagRect" => PresetShapeType::Snip2DiagRect,
        other => {
            log::warn!("prstGeom: unrecognized shape type {:?}", other);
            PresetShapeType::Other(other.to_owned())
        }
    }
}

// ── a:ln (§20.1.2.2.24) ────────────────────────────────────────────────────

fn parse_outline(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Outline> {
    let width = xml::optional_attr_i64(start, b"w")?.map(Dimension::new);
    let cap = xml::optional_attr(start, b"cap")?
        .map(|s| parse_line_cap(&s))
        .transpose()?;
    let compound = xml::optional_attr(start, b"cmpd")?
        .map(|s| parse_compound_line(&s))
        .transpose()?;
    let alignment = xml::optional_attr(start, b"algn")?
        .map(|s| parse_pen_alignment(&s))
        .transpose()?;
    let mut fill = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                // Skip complex children (dash, join, head/tail end, fills).
                xml::skip_to_end(reader, buf, local)?;
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                if local == b"noFill" {
                    fill = Some(DrawingFill::NoFill);
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"ln" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"ln")),
            _ => {}
        }
    }

    Ok(Outline {
        width,
        cap,
        compound,
        alignment,
        fill,
    })
}

fn parse_line_cap(val: &str) -> Result<LineCap> {
    match val {
        "flat" => Ok(LineCap::Flat),
        "rnd" => Ok(LineCap::Round),
        "sq" => Ok(LineCap::Square),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "cap".into(),
            value: other.into(),
            reason: "expected flat, rnd, or sq per §20.1.10.31 ST_LineCap".into(),
        }),
    }
}

fn parse_compound_line(val: &str) -> Result<CompoundLine> {
    match val {
        "sng" => Ok(CompoundLine::Single),
        "dbl" => Ok(CompoundLine::Double),
        "thickThin" => Ok(CompoundLine::ThickThin),
        "thinThick" => Ok(CompoundLine::ThinThick),
        "tri" => Ok(CompoundLine::Triple),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "cmpd".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.15 ST_CompoundLine".into(),
        }),
    }
}

fn parse_pen_alignment(val: &str) -> Result<PenAlignment> {
    match val {
        "ctr" => Ok(PenAlignment::Center),
        "in" => Ok(PenAlignment::Inset),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "algn".into(),
            value: other.into(),
            reason: "expected ctr or in per §20.1.10.39 ST_PenAlignment".into(),
        }),
    }
}

// ── wp:anchor (§20.4.2.3) ──────────────────────────────────────────────────

/// Parse `wp:anchor` into an `Image` with `ImagePlacement::Anchor`.
pub fn parse_anchor_image(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Option<Image>> {
    // §20.4.2.3 attributes
    let dist_t = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
    let dist_b = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
    let dist_l = xml::optional_attr_i64(e, b"distL")?.unwrap_or(0);
    let dist_r = xml::optional_attr_i64(e, b"distR")?.unwrap_or(0);
    let distance = EdgeInsets::new(
        Dimension::new(dist_t),
        Dimension::new(dist_r),
        Dimension::new(dist_b),
        Dimension::new(dist_l),
    );
    let use_simple_pos = xml::optional_attr_bool(e, b"simplePos")?;
    let relative_height = xml::optional_attr_u32(e, b"relativeHeight")?.ok_or_else(|| {
        ParseError::MissingAttribute {
            element: "wp:anchor".into(),
            attr: "relativeHeight".into(),
        }
    })?;
    let behind_text =
        xml::optional_attr_bool(e, b"behindDoc")?.ok_or_else(|| ParseError::MissingAttribute {
            element: "wp:anchor".into(),
            attr: "behindDoc".into(),
        })?;
    let lock_anchor =
        xml::optional_attr_bool(e, b"locked")?.ok_or_else(|| ParseError::MissingAttribute {
            element: "wp:anchor".into(),
            attr: "locked".into(),
        })?;
    let allow_overlap = xml::optional_attr_bool(e, b"allowOverlap")?.ok_or_else(|| {
        ParseError::MissingAttribute {
            element: "wp:anchor".into(),
            attr: "allowOverlap".into(),
        }
    })?;
    let layout_in_cell = xml::optional_attr_bool(e, b"layoutInCell")?;
    let hidden = xml::optional_attr_bool(e, b"hidden")?;

    let mut extent = Size::ZERO;
    let mut effect_extent = None;
    let mut doc_properties = None;
    let mut graphic_frame_locks = None;
    let mut picture = None;
    let mut simple_pos = None;
    let mut h_pos = None;
    let mut v_pos = None;
    let mut wrap = TextWrap::None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"simplePos" => {
                        let x = xml::optional_attr_i64(e, b"x")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "wp:simplePos".into(),
                                attr: "x".into(),
                            }
                        })?;
                        let y = xml::optional_attr_i64(e, b"y")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "wp:simplePos".into(),
                                attr: "y".into(),
                            }
                        })?;
                        simple_pos = Some(Offset::new(Dimension::new(x), Dimension::new(y)));
                        xml::skip_to_end(reader, buf, b"simplePos")?;
                    }
                    b"positionH" => {
                        let rel_from =
                            xml::optional_attr(e, b"relativeFrom")?.ok_or_else(|| {
                                ParseError::MissingAttribute {
                                    element: "wp:positionH".into(),
                                    attr: "relativeFrom".into(),
                                }
                            })?;
                        let rel = parse_anchor_relative_from(&rel_from)?;
                        h_pos = Some(parse_anchor_position(reader, buf, rel, b"positionH")?);
                    }
                    b"positionV" => {
                        let rel_from =
                            xml::optional_attr(e, b"relativeFrom")?.ok_or_else(|| {
                                ParseError::MissingAttribute {
                                    element: "wp:positionV".into(),
                                    attr: "relativeFrom".into(),
                                }
                            })?;
                        let rel = parse_anchor_relative_from(&rel_from)?;
                        v_pos = Some(parse_anchor_position(reader, buf, rel, b"positionV")?);
                    }
                    b"wrapSquare" => {
                        let wrap_text = parse_wrap_text_attr(e)?;
                        wrap = TextWrap::Square {
                            distance: parse_wrap_distance(e)?,
                            wrap_text,
                        };
                        xml::skip_to_end(reader, buf, b"wrapSquare")?;
                    }
                    b"wrapTight" => {
                        let wrap_text = parse_wrap_text_attr(e)?;
                        wrap = TextWrap::Tight {
                            distance: parse_wrap_distance(e)?,
                            wrap_text,
                        };
                        xml::skip_to_end(reader, buf, b"wrapTight")?;
                    }
                    b"wrapThrough" => {
                        let wrap_text = parse_wrap_text_attr(e)?;
                        wrap = TextWrap::Through {
                            distance: parse_wrap_distance(e)?,
                            wrap_text,
                        };
                        xml::skip_to_end(reader, buf, b"wrapThrough")?;
                    }
                    b"wrapTopAndBottom" => {
                        let dt = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
                        let db = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
                        wrap = TextWrap::TopAndBottom {
                            distance_top: Dimension::new(dt),
                            distance_bottom: Dimension::new(db),
                        };
                        xml::skip_to_end(reader, buf, b"wrapTopAndBottom")?;
                    }
                    b"docPr" => {
                        doc_properties = Some(parse_doc_properties(e, Some(reader), Some(buf))?);
                    }
                    b"cNvGraphicFramePr" => {
                        graphic_frame_locks = Some(parse_cnv_graphic_frame_pr(reader, buf)?);
                    }
                    b"graphic" => {
                        picture = parse_graphic(reader, buf)?;
                    }
                    // Office extensions — skip content but don't warn.
                    b"sizeRelH" | b"sizeRelV" => {
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("anchor-image", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"extent" => {
                        extent = parse_positive_size_2d(e)?;
                    }
                    b"effectExtent" => {
                        effect_extent = Some(parse_effect_extent(e)?);
                    }
                    b"docPr" => {
                        doc_properties = Some(parse_doc_properties(e, None, None)?);
                    }
                    b"simplePos" => {
                        let x = xml::optional_attr_i64(e, b"x")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "wp:simplePos".into(),
                                attr: "x".into(),
                            }
                        })?;
                        let y = xml::optional_attr_i64(e, b"y")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "wp:simplePos".into(),
                                attr: "y".into(),
                            }
                        })?;
                        simple_pos = Some(Offset::new(Dimension::new(x), Dimension::new(y)));
                    }
                    b"wrapNone" => {
                        wrap = TextWrap::None;
                    }
                    b"wrapSquare" => {
                        let wrap_text = parse_wrap_text_attr(e)?;
                        wrap = TextWrap::Square {
                            distance: parse_wrap_distance(e)?,
                            wrap_text,
                        };
                    }
                    b"wrapTight" => {
                        let wrap_text = parse_wrap_text_attr(e)?;
                        wrap = TextWrap::Tight {
                            distance: parse_wrap_distance(e)?,
                            wrap_text,
                        };
                    }
                    b"wrapThrough" => {
                        let wrap_text = parse_wrap_text_attr(e)?;
                        wrap = TextWrap::Through {
                            distance: parse_wrap_distance(e)?,
                            wrap_text,
                        };
                    }
                    b"wrapTopAndBottom" => {
                        let dt = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
                        let db = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
                        wrap = TextWrap::TopAndBottom {
                            distance_top: Dimension::new(dt),
                            distance_bottom: Dimension::new(db),
                        };
                    }
                    _ => xml::warn_unsupported_element("anchor-image", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"anchor" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"anchor")),
            _ => {}
        }
    }

    let doc_properties = doc_properties.ok_or_else(|| ParseError::MissingElement {
        parent: "wp:anchor".into(),
        child: "wp:docPr".into(),
    })?;

    let h_pos = h_pos.ok_or_else(|| ParseError::MissingElement {
        parent: "wp:anchor".into(),
        child: "wp:positionH".into(),
    })?;

    let v_pos = v_pos.ok_or_else(|| ParseError::MissingElement {
        parent: "wp:anchor".into(),
        child: "wp:positionV".into(),
    })?;

    Ok(Some(Image {
        extent,
        effect_extent,
        doc_properties,
        graphic_frame_locks,
        graphic: picture,
        placement: ImagePlacement::Anchor(AnchorProperties {
            distance,
            simple_pos,
            use_simple_pos,
            horizontal_position: h_pos,
            vertical_position: v_pos,
            wrap,
            behind_text,
            lock_anchor,
            allow_overlap,
            relative_height,
            layout_in_cell,
            hidden,
        }),
    }))
}

// ── Anchor helpers ──────────────────────────────────────────────────────────

/// §20.4.2.10 / §20.4.2.11: parse positionH or positionV children.
fn parse_anchor_position(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    relative_from: AnchorRelativeFrom,
    end_tag: &[u8],
) -> Result<AnchorPosition> {
    let mut result = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"posOffset" => {
                        let text = xml::read_text_content(reader, buf)?;
                        let val: i64 =
                            text.trim()
                                .parse()
                                .map_err(|_| ParseError::InvalidAttributeValue {
                                    attr: "posOffset".into(),
                                    value: text.clone(),
                                    reason: "expected integer EMU value".into(),
                                })?;
                        result = Some(AnchorPosition::Offset {
                            relative_from,
                            offset: Dimension::new(val),
                        });
                    }
                    b"align" => {
                        let text = xml::read_text_content(reader, buf)?;
                        result = Some(AnchorPosition::Align {
                            relative_from,
                            alignment: parse_anchor_alignment(text.trim())?,
                        });
                    }
                    _ => {
                        xml::warn_unsupported_element("anchor-position", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }

    result.ok_or_else(|| ParseError::MissingElement {
        parent: String::from_utf8_lossy(end_tag).into(),
        child: "wp:posOffset or wp:align".into(),
    })
}

/// §20.4.3.4 ST_RelFromH / §20.4.3.5 ST_RelFromV.
fn parse_anchor_relative_from(val: &str) -> Result<AnchorRelativeFrom> {
    match val {
        "page" => Ok(AnchorRelativeFrom::Page),
        "margin" => Ok(AnchorRelativeFrom::Margin),
        "column" => Ok(AnchorRelativeFrom::Column),
        "character" => Ok(AnchorRelativeFrom::Character),
        "paragraph" => Ok(AnchorRelativeFrom::Paragraph),
        "line" => Ok(AnchorRelativeFrom::Line),
        "insideMargin" => Ok(AnchorRelativeFrom::InsideMargin),
        "outsideMargin" => Ok(AnchorRelativeFrom::OutsideMargin),
        "topMargin" => Ok(AnchorRelativeFrom::TopMargin),
        "bottomMargin" => Ok(AnchorRelativeFrom::BottomMargin),
        "leftMargin" => Ok(AnchorRelativeFrom::LeftMargin),
        "rightMargin" => Ok(AnchorRelativeFrom::RightMargin),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "relativeFrom".into(),
            value: other.into(),
            reason: "expected value per §20.4.3.4/§20.4.3.5".into(),
        }),
    }
}

/// §20.4.3.1 ST_AlignH / §20.4.3.2 ST_AlignV.
fn parse_anchor_alignment(val: &str) -> Result<AnchorAlignment> {
    match val {
        "left" => Ok(AnchorAlignment::Left),
        "center" => Ok(AnchorAlignment::Center),
        "right" => Ok(AnchorAlignment::Right),
        "inside" => Ok(AnchorAlignment::Inside),
        "outside" => Ok(AnchorAlignment::Outside),
        "top" => Ok(AnchorAlignment::Top),
        "bottom" => Ok(AnchorAlignment::Bottom),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "align".into(),
            value: other.into(),
            reason: "expected value per §20.4.3.1/§20.4.3.2".into(),
        }),
    }
}

/// Parse wrap distance attributes (distT, distB, distL, distR) in EMU.
fn parse_wrap_distance(e: &BytesStart<'_>) -> Result<EdgeInsets<Emu>> {
    let t = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
    let b = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
    let l = xml::optional_attr_i64(e, b"distL")?.unwrap_or(0);
    let r = xml::optional_attr_i64(e, b"distR")?.unwrap_or(0);
    Ok(EdgeInsets::new(
        Dimension::new(t),
        Dimension::new(r),
        Dimension::new(b),
        Dimension::new(l),
    ))
}

/// §20.4.3.7 ST_WrapText — parse wrapText attribute.
fn parse_wrap_text_attr(e: &BytesStart<'_>) -> Result<WrapText> {
    let val = xml::optional_attr(e, b"wrapText")?.unwrap_or_default();
    match val.as_str() {
        "bothSides" | "" => Ok(WrapText::BothSides),
        "left" => Ok(WrapText::Left),
        "right" => Ok(WrapText::Right),
        "largest" => Ok(WrapText::Largest),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "wrapText".into(),
            value: other.into(),
            reason: "expected value per §20.4.3.7 ST_WrapText".into(),
        }),
    }
}
