use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::{Dimension, Emu};
use crate::docx::error::{ParseError, Result};
use crate::docx::geometry::{EdgeInsets, Offset, Size};
use crate::docx::model::*;
use crate::docx::xml;

use super::{
    parse_cnv_graphic_frame_pr, parse_doc_properties, parse_effect_extent, parse_graphic,
    parse_positive_size_2d,
};

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
                    // MCE §M.2.1: mc:AlternateContent wraps elements like
                    // positionH/positionV with wp14 extensions in mc:Choice
                    // and standard fallbacks in mc:Fallback. Skip to Fallback
                    // and re-parse its children as anchor children.
                    b"AlternateContent" => {
                        parse_anchor_alternate_content(reader, buf, &mut h_pos, &mut v_pos)?;
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

/// MCE §M.2.1: parse mc:AlternateContent inside wp:anchor.
/// Skips mc:Choice, parses mc:Fallback children (positionH/positionV).
fn parse_anchor_alternate_content(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    h_pos: &mut Option<AnchorPosition>,
    v_pos: &mut Option<AnchorPosition>,
) -> Result<()> {
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"Choice" => xml::skip_to_end(reader, buf, b"Choice")?,
                    b"Fallback" => {
                        // Parse Fallback children as anchor-level elements.
                        loop {
                            match xml::next_event(reader, buf)? {
                                Event::Start(ref e2) => {
                                    let qn2 = e2.name();
                                    let local2 = xml::local_name(qn2.as_ref());
                                    match local2 {
                                        b"positionH" => {
                                            if let Some(rel) =
                                                xml::optional_attr(e2, b"relativeFrom")?
                                            {
                                                let rel = parse_anchor_relative_from(&rel)?;
                                                *h_pos = Some(parse_anchor_position(
                                                    reader,
                                                    buf,
                                                    rel,
                                                    b"positionH",
                                                )?);
                                            }
                                        }
                                        b"positionV" => {
                                            if let Some(rel) =
                                                xml::optional_attr(e2, b"relativeFrom")?
                                            {
                                                let rel = parse_anchor_relative_from(&rel)?;
                                                *v_pos = Some(parse_anchor_position(
                                                    reader,
                                                    buf,
                                                    rel,
                                                    b"positionV",
                                                )?);
                                            }
                                        }
                                        _ => xml::skip_to_end(reader, buf, local2)?,
                                    }
                                }
                                Event::End(ref e2)
                                    if xml::local_name(e2.name().as_ref()) == b"Fallback" =>
                                {
                                    break
                                }
                                Event::Eof => return Err(xml::unexpected_eof(b"Fallback")),
                                _ => {}
                            }
                        }
                    }
                    _ => xml::skip_to_end(reader, buf, local)?,
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"AlternateContent" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"AlternateContent")),
            _ => {}
        }
    }
    Ok(())
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
