use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::model::*;
use crate::docx::xml;

use super::{parse_cnv_pr, parse_cnv_pr_attrs, shape};

// ── pic:pic (§19.3.1.37) ───────────────────────────────────────────────────

pub(super) fn parse_picture(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Picture> {
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
                        shape_properties = Some(shape::parse_shape_properties(e, reader, buf)?);
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
