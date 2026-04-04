//! Parser for VML (Vector Markup Language) elements: shapes, shape types,
//! path commands, formulas, styles, and related attributes.
//!
//! VML is the legacy drawing format used in `w:pict` containers (§14.1).

mod color;
mod formulas;
mod path_commands;
mod style;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

use self::color::parse_color;
use self::formulas::parse_formulas;
use self::path_commands::{parse_adj, parse_path_commands, parse_vector2d};
use self::style::{parse_length, parse_style};

// ── Pict container ─────────────────────────────────────────────────────────

/// §17.3.3.19: parse `w:pict` — VML picture container.
pub fn parse_pict(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Pict> {
    let mut shape_type = None;
    let mut shapes = Vec::new();

    loop {
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"shapetype" => {
                        shape_type = Some(if is_start {
                            parse_shapetype(e, reader, buf)?
                        } else {
                            parse_shapetype_from_attrs(e)?
                        });
                    }
                    b"shape" => {
                        shapes.push(if is_start {
                            parse_shape(e, reader, buf)?
                        } else {
                            parse_shape_from_attrs(e)?
                        });
                    }
                    _ => {
                        xml::warn_unsupported_element("pict", local);
                        if is_start {
                            xml::skip_to_end(reader, buf, local)?;
                        }
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pict" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"pict")),
            _ => {}
        }
    }

    Ok(Pict { shape_type, shapes })
}

// ── ShapeType ──────────────────────────────────────────────────────────────

/// VML §14.1.2.20: parse self-closing `v:shapetype` (attributes only, no children).
fn parse_shapetype_from_attrs(e: &BytesStart<'_>) -> Result<VmlShapeType> {
    Ok(VmlShapeType {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        coord_size: parse_vector2d(xml::optional_attr(e, b"coordsize")?),
        spt: xml::optional_attr(e, b"spt")?
            .map(|s| s.parse::<f32>())
            .transpose()
            .ok()
            .flatten(),
        adj: parse_adj(xml::optional_attr(e, b"adj")?),
        path: parse_path_commands(xml::optional_attr(e, b"path")?),
        filled: bool_attr(e, b"filled")?,
        stroked: bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        formulas: Vec::new(),
        lock: None,
    })
}

/// VML §14.1.2.20: parse `v:shapetype` with attributes and children.
fn parse_shapetype(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<VmlShapeType> {
    let mut st = VmlShapeType {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        coord_size: parse_vector2d(xml::optional_attr(e, b"coordsize")?),
        spt: xml::optional_attr(e, b"spt")?
            .map(|s| s.parse::<f32>())
            .transpose()
            .ok()
            .flatten(),
        adj: parse_adj(xml::optional_attr(e, b"adj")?),
        path: parse_path_commands(xml::optional_attr(e, b"path")?),
        filled: bool_attr(e, b"filled")?,
        stroked: bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        formulas: Vec::new(),
        lock: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"formulas" => {
                        st.formulas = parse_formulas(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("shapetype", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"stroke" => st.stroke = Some(parse_stroke_attrs(e)?),
                    b"path" => st.vml_path = Some(parse_path_attrs(e)?),
                    b"lock" => st.lock = Some(parse_lock(e)?),
                    _ => xml::warn_unsupported_element("shapetype", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"shapetype" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"shapetype")),
            _ => {}
        }
    }

    Ok(st)
}

// ── Shape ──────────────────────────────────────────────────────────────────

/// VML §14.1.2.19: parse self-closing `v:shape` (attributes only, no children).
fn parse_shape_from_attrs(e: &BytesStart<'_>) -> Result<VmlShape> {
    Ok(VmlShape {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        shape_type_ref: xml::optional_attr(e, b"type")?
            .map(|s| VmlShapeId::new(s.strip_prefix('#').unwrap_or(&s))),
        style: parse_style(xml::optional_attr(e, b"style")?),
        fill_color: xml::optional_attr(e, b"fillcolor")?
            .map(|s| parse_color(&s))
            .transpose()?,
        stroked: bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        text_box: None,
        wrap: None,
        image_data: None,
    })
}

/// VML §14.1.2.19: parse `v:shape` with attributes and children.
fn parse_shape(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<VmlShape> {
    let mut shape = VmlShape {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        shape_type_ref: xml::optional_attr(e, b"type")?
            .map(|s| VmlShapeId::new(s.strip_prefix('#').unwrap_or(&s))),
        style: parse_style(xml::optional_attr(e, b"style")?),
        fill_color: xml::optional_attr(e, b"fillcolor")?
            .map(|s| parse_color(&s))
            .transpose()?,
        stroked: bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        text_box: None,
        wrap: None,
        image_data: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"textbox" => {
                        shape.text_box = Some(parse_textbox(e, reader, buf)?);
                    }
                    b"wrap" => {
                        shape.wrap = Some(parse_wrap(e)?);
                        xml::skip_to_end(reader, buf, b"wrap")?;
                    }
                    _ => {
                        xml::warn_unsupported_element("shape", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"stroke" => shape.stroke = Some(parse_stroke_attrs(e)?),
                    b"path" => shape.vml_path = Some(parse_path_attrs(e)?),
                    b"imagedata" => shape.image_data = Some(parse_imagedata(e)?),
                    b"wrap" => shape.wrap = Some(parse_wrap(e)?),
                    _ => xml::warn_unsupported_element("shape", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"shape" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"shape")),
            _ => {}
        }
    }

    Ok(shape)
}

// ── TextBox ───────────────────────────────────────────────────────────────

/// VML §14.1.2.22: parse `v:textbox` with `w:txbxContent`.
fn parse_textbox(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<VmlTextBox> {
    let style = parse_style(xml::optional_attr(e, b"style")?);
    let inset = parse_textbox_inset(xml::optional_attr(e, b"inset")?);
    let mut content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"txbxContent" => {
                        let (blocks, _) =
                            super::body::parse_block_content(reader, buf, b"txbxContent")?;
                        content = blocks;
                    }
                    _ => {
                        xml::warn_unsupported_element("textbox", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"textbox" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"textbox")),
            _ => {}
        }
    }

    Ok(VmlTextBox {
        style,
        inset,
        content,
    })
}

// ── Stroke ─────────────────────────────────────────────────────────────────

/// VML §14.1.2.21: parse `v:stroke` attributes.
fn parse_stroke_attrs(e: &BytesStart<'_>) -> Result<VmlStroke> {
    let dash_style = match xml::optional_attr(e, b"dashstyle")?.as_deref() {
        Some("solid") => Some(VmlDashStyle::Solid),
        Some("shortdash") => Some(VmlDashStyle::ShortDash),
        Some("shortdot") => Some(VmlDashStyle::ShortDot),
        Some("shortdashdot") => Some(VmlDashStyle::ShortDashDot),
        Some("shortdashdotdot") => Some(VmlDashStyle::ShortDashDotDot),
        Some("dot") => Some(VmlDashStyle::Dot),
        Some("dash") => Some(VmlDashStyle::Dash),
        Some("longdash") => Some(VmlDashStyle::LongDash),
        Some("dashdot") => Some(VmlDashStyle::DashDot),
        Some("longdashdot") => Some(VmlDashStyle::LongDashDot),
        Some("longdashdotdot") => Some(VmlDashStyle::LongDashDotDot),
        Some(other) => {
            log::warn!("vml-stroke: unsupported dashstyle {:?}", other);
            None
        }
        None => None,
    };
    let join_style = match xml::optional_attr(e, b"joinstyle")?.as_deref() {
        Some("round") => Some(VmlJoinStyle::Round),
        Some("bevel") => Some(VmlJoinStyle::Bevel),
        Some("miter") => Some(VmlJoinStyle::Miter),
        Some(other) => {
            log::warn!("vml-stroke: unsupported joinstyle {:?}", other);
            None
        }
        None => None,
    };
    Ok(VmlStroke {
        dash_style,
        join_style,
    })
}

// ── Path element ───────────────────────────────────────────────────────────

/// VML §14.1.2.14: parse `v:path` attributes.
fn parse_path_attrs(e: &BytesStart<'_>) -> Result<VmlPath> {
    Ok(VmlPath {
        gradient_shape_ok: bool_attr(e, b"gradientshapeok")?,
        connect_type: match xml::optional_attr(e, b"connecttype")?.as_deref() {
            Some("none") => Some(VmlConnectType::None),
            Some("rect") => Some(VmlConnectType::Rect),
            Some("segments") => Some(VmlConnectType::Segments),
            Some("custom") => Some(VmlConnectType::Custom),
            Some(other) => {
                log::warn!("vml-path: unsupported connecttype {:?}", other);
                None
            }
            None => None,
        },
        extrusion_ok: bool_attr(e, b"extrusionok")?,
    })
}

// ── Misc attributes ────────────────────────────────────────────────────────

/// Parse VML textbox `inset` attribute — comma-separated CSS lengths (left,top,right,bottom).
fn parse_textbox_inset(s: Option<String>) -> Option<VmlTextBoxInset> {
    let s = s?;
    let parts: Vec<&str> = s.split(',').collect();
    Some(VmlTextBoxInset {
        left: parts.first().and_then(|v| parse_length(v)),
        top: parts.get(1).and_then(|v| parse_length(v)),
        right: parts.get(2).and_then(|v| parse_length(v)),
        bottom: parts.get(3).and_then(|v| parse_length(v)),
    })
}

/// Office VML extension: parse `o:lock` attributes.
fn parse_lock(e: &BytesStart<'_>) -> Result<VmlLock> {
    let aspect_ratio = bool_attr(e, b"aspectratio")?;
    let ext = match xml::optional_attr(e, b"ext")?.as_deref() {
        Some("edit") => Some(VmlExtHandling::Edit),
        Some("view") => Some(VmlExtHandling::View),
        Some("backwardCompatible") => Some(VmlExtHandling::BackwardCompatible),
        Some(other) => {
            return Err(crate::docx::error::ParseError::InvalidAttributeValue {
                attr: "v:ext".into(),
                value: other.into(),
                reason: "expected edit, view, or backwardCompatible".into(),
            });
        }
        None => None,
    };
    Ok(VmlLock { aspect_ratio, ext })
}

/// VML §14.1.2.11: parse `v:imagedata` attributes.
fn parse_imagedata(e: &BytesStart<'_>) -> Result<VmlImageData> {
    Ok(VmlImageData {
        rel_id: xml::optional_attr(e, b"id")?.map(RelId::new),
        title: xml::optional_attr(e, b"title")?,
    })
}

/// VML §14.1.2.23: parse `v:wrap` attributes.
fn parse_wrap(e: &BytesStart<'_>) -> Result<VmlWrap> {
    let wrap_type = match xml::optional_attr(e, b"type")?.as_deref() {
        Some("topAndBottom") => Some(VmlWrapType::TopAndBottom),
        Some("square") => Some(VmlWrapType::Square),
        Some("none") => Some(VmlWrapType::None),
        Some("tight") => Some(VmlWrapType::Tight),
        Some("through") => Some(VmlWrapType::Through),
        Some(other) => {
            return Err(crate::docx::error::ParseError::InvalidAttributeValue {
                attr: "type".into(),
                value: other.into(),
                reason: "expected value per VML §14.1.2.23".into(),
            });
        }
        None => None,
    };
    let side = match xml::optional_attr(e, b"side")?.as_deref() {
        Some("both") => Some(VmlWrapSide::Both),
        Some("left") => Some(VmlWrapSide::Left),
        Some("right") => Some(VmlWrapSide::Right),
        Some("largest") => Some(VmlWrapSide::Largest),
        Some(other) => {
            return Err(crate::docx::error::ParseError::InvalidAttributeValue {
                attr: "side".into(),
                value: other.into(),
                reason: "expected value per VML §14.1.2.23".into(),
            });
        }
        None => None,
    };
    Ok(VmlWrap { wrap_type, side })
}

/// Parse a VML boolean attribute ("t"/"f" or "true"/"false").
fn bool_attr(e: &BytesStart<'_>, name: &[u8]) -> Result<Option<bool>> {
    match xml::optional_attr(e, name)?.as_deref() {
        Some("t") | Some("true") => Ok(Some(true)),
        Some("f") | Some("false") => Ok(Some(false)),
        Some(other) => {
            xml::warn_unsupported_attr("vml", &String::from_utf8_lossy(name), other);
            Ok(None)
        }
        None => Ok(None),
    }
}
