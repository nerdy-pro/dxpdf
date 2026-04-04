use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::geometry::{EdgeInsets, Size};
use crate::docx::model::*;
use crate::docx::xml;

use super::{
    parse_cnv_graphic_frame_pr, parse_doc_properties, parse_effect_extent, parse_graphic,
    parse_positive_size_2d,
};

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
