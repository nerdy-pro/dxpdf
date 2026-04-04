//! Hierarchical parsing of DrawingML elements (inline images, anchor images).
//!
//! Follows the actual XML nesting per OOXML spec rather than flat event scanning.

mod anchor;
mod inline;
mod picture;
mod shape;

pub use anchor::parse_anchor_image;
pub use inline::parse_inline_image;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::{Dimension, Emu};
use crate::docx::error::{ParseError, Result};
use crate::docx::geometry::{EdgeInsets, Size};
use crate::docx::model::*;
use crate::docx::xml;

// ── Shared element parsers ──────────────────────────────────────────────────

/// §20.4.2.7 / §20.1.7.3: parse `cx`, `cy` into `Size<Emu>`.
pub(super) fn parse_positive_size_2d(e: &BytesStart<'_>) -> Result<Size<Emu>> {
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
pub(super) fn parse_effect_extent(e: &BytesStart<'_>) -> Result<EdgeInsets<Emu>> {
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
pub(super) fn parse_doc_properties(
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

/// §20.1.2.2.8: parse `pic:cNvPr` attributes only.
pub(super) fn parse_cnv_pr_attrs(e: &BytesStart<'_>) -> Result<DocProperties> {
    Ok(DocProperties {
        id: xml::optional_attr_u32(e, b"id")?.unwrap_or(0),
        name: xml::optional_attr(e, b"name")?.unwrap_or_default(),
        description: xml::optional_attr(e, b"descr")?,
        hidden: xml::optional_attr_bool(e, b"hidden")?,
        title: xml::optional_attr(e, b"title")?,
    })
}

/// §20.1.2.2.8: parse `pic:cNvPr` with children.
pub(super) fn parse_cnv_pr(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<DocProperties> {
    let props = parse_cnv_pr_attrs(e)?;
    // cNvPr can have extension children (hlinkClick, hlinkHover, extLst) — skip.
    xml::skip_to_end(reader, buf, b"cNvPr")?;
    Ok(props)
}

// ── wp:cNvGraphicFramePr (§20.4.2.4) ───────────────────────────────────────

pub(super) fn parse_cnv_graphic_frame_pr(
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
pub(super) fn parse_graphic(
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
                        content = Some(GraphicContent::Picture(picture::parse_picture(
                            reader, buf,
                        )?));
                    }
                    b"wsp" => {
                        content = Some(GraphicContent::WordProcessingShape(
                            shape::parse_word_processing_shape(reader, buf)?,
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
