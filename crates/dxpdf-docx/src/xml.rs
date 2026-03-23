//! XML parsing helper utilities.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::error::{ParseError, Result};

/// Extract the local name from a potentially namespaced XML tag.
/// e.g., b"w:p" → b"p", b"p" → b"p"
pub fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}

/// Extract the local name as an owned Vec for use across borrow boundaries.
pub fn local_name_owned(name: &[u8]) -> Vec<u8> {
    local_name(name).to_vec()
}

/// Read the next event from the reader, returning an owned event.
/// This avoids borrow conflicts between the event data and the buffer.
pub fn next_event(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Event<'static>> {
    buf.clear();
    Ok(reader.read_event_into(buf)?.into_owned())
}

/// Get a required attribute value as a string.
pub fn required_attr(elem: &BytesStart<'_>, attr_name: &[u8]) -> Result<String> {
    for attr in elem.attributes().with_checks(false) {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == attr_name {
            return Ok(String::from_utf8_lossy(&attr.value).into_owned());
        }
    }
    Err(ParseError::MissingAttribute {
        element: String::from_utf8_lossy(elem.name().as_ref()).into_owned(),
        attr: String::from_utf8_lossy(attr_name).into_owned(),
    })
}

/// Get an optional attribute value as a string.
pub fn optional_attr(elem: &BytesStart<'_>, attr_name: &[u8]) -> Result<Option<String>> {
    for attr in elem.attributes().with_checks(false) {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == attr_name {
            return Ok(Some(String::from_utf8_lossy(&attr.value).into_owned()));
        }
    }
    Ok(None)
}

/// Get an optional attribute parsed as i64.
pub fn optional_attr_i64(elem: &BytesStart<'_>, attr_name: &[u8]) -> Result<Option<i64>> {
    match optional_attr(elem, attr_name)? {
        Some(s) => Ok(Some(s.parse::<i64>()?)),
        None => Ok(None),
    }
}

/// Get an optional attribute parsed as u32.
pub fn optional_attr_u32(elem: &BytesStart<'_>, attr_name: &[u8]) -> Result<Option<u32>> {
    match optional_attr(elem, attr_name)? {
        Some(s) => Ok(Some(s.parse::<u32>()?)),
        None => Ok(None),
    }
}

/// Get an optional boolean attribute. Treats "1", "true", "on" as true.
pub fn optional_attr_bool(elem: &BytesStart<'_>, attr_name: &[u8]) -> Result<Option<bool>> {
    match optional_attr(elem, attr_name)? {
        Some(s) => Ok(Some(parse_bool(&s))),
        None => Ok(None),
    }
}

/// Parse an OOXML boolean value.
pub fn parse_bool(s: &str) -> bool {
    matches!(s, "1" | "true" | "on")
}

/// Parse an optional rsid attribute from an element.
pub fn optional_rsid(
    elem: &BytesStart<'_>,
    attr_name: &[u8],
) -> Result<Option<crate::model::RevisionSaveId>> {
    match optional_attr(elem, attr_name)? {
        Some(s) => Ok(crate::model::RevisionSaveId::from_hex(&s)),
        None => Ok(None),
    }
}

/// Parse a hex color string (6 hex digits) → u32 RGB.
pub fn parse_hex_color(s: &str) -> Option<u32> {
    if s.eq_ignore_ascii_case("auto") {
        return None; // Caller should handle "auto" separately
    }
    u32::from_str_radix(s, 16).ok()
}

/// Read all text content until the end of the current element.
pub fn read_text_content(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<String> {
    let mut text = String::new();
    loop {
        match next_event(reader, buf)? {
            Event::Text(e) => {
                text.push_str(&e.unescape()?);
            }
            Event::End(_) | Event::Eof => break,
            _ => {}
        }
    }
    Ok(text)
}

/// Log a warning for an unsupported element within a given context.
pub fn warn_unsupported_element(context: &str, local: &[u8]) {
    log::warn!(
        "{context}: unsupported element <{}>",
        String::from_utf8_lossy(local)
    );
}

/// Log a warning for an unsupported attribute value.
pub fn warn_unsupported_attr(context: &str, attr: &str, value: &str) {
    log::warn!("{context}: unsupported {attr} value \"{value}\"");
}

/// Skip elements until we reach the End event for `end_tag`.
pub fn skip_to_end(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>, end_tag: &[u8]) -> Result<()> {
    let mut depth = 1u32;
    loop {
        match next_event(reader, buf)? {
            Event::Start(_) => depth += 1,
            Event::End(ref e) => {
                depth -= 1;
                if depth == 0 && local_name(e.name().as_ref()) == end_tag {
                    return Ok(());
                }
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
    }
}
