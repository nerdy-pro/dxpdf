//! Parser for `word/styles.xml` — parses style definitions as-is.
//! No inheritance resolution — `basedOn` references are preserved.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::*;
use crate::xml;

use super::properties;

/// Parse `word/styles.xml` into a raw `StyleSheet`.
pub fn parse_styles(data: &[u8]) -> Result<StyleSheet> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut sheet = StyleSheet::default();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"docDefaults" => {
                        parse_doc_defaults(&mut reader, &mut buf, &mut sheet)?;
                    }
                    b"style" => {
                        if let Some((id, style)) = parse_style(e, &mut reader, &mut buf)? {
                            sheet.styles.insert(id, style);
                        }
                    }
                    _ => xml::warn_unsupported_element("styles", &local),
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(sheet)
}

fn parse_doc_defaults(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    sheet: &mut StyleSheet,
) -> Result<()> {
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"rPr" => {
                        let (parsed, _) = properties::parse_run_properties(reader, buf)?;
                        sheet.doc_defaults_run = parsed;
                    }
                    b"pPr" => {
                        let parsed = properties::parse_paragraph_properties(reader, buf)?;
                        sheet.doc_defaults_paragraph = parsed.properties;
                    }
                    _ => xml::warn_unsupported_element("docDefaults", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"docDefaults" => break,
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

fn parse_style(
    start: &quick_xml::events::BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Option<(StyleId, Style)>> {
    let style_type = match xml::optional_attr(start, b"type")?.as_deref() {
        Some("paragraph") | None => StyleType::Paragraph,
        Some("character") => StyleType::Character,
        Some("table") => StyleType::Table,
        Some("numbering") => StyleType::Numbering,
        Some(other) => {
            log::warn!("unknown style type: {other}");
            xml::skip_to_end(reader, buf, b"style")?;
            return Ok(None);
        }
    };

    let id = match xml::optional_attr(start, b"styleId")? {
        Some(id) => StyleId(id),
        None => {
            xml::skip_to_end(reader, buf, b"style")?;
            return Ok(None);
        }
    };

    let is_default = xml::optional_attr_bool(start, b"default")?.unwrap_or(false);

    let mut name: Option<String> = None;
    let mut based_on: Option<StyleId> = None;
    let mut ppr: Option<ParagraphProperties> = None;
    let mut rpr: Option<RunProperties> = None;
    let mut tbl_pr: Option<TableProperties> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pPr" => {
                        let parsed = properties::parse_paragraph_properties(reader, buf)?;
                        ppr = Some(parsed.properties);
                        // If rPr was inside pPr, it's the paragraph mark's run properties.
                        // For style definitions, we store it as the style's rPr if no
                        // standalone rPr follows.
                        if rpr.is_none() {
                            rpr = parsed.run_properties;
                        }
                    }
                    b"rPr" => {
                        let (parsed, _) = properties::parse_run_properties(reader, buf)?;
                        rpr = Some(parsed);
                    }
                    b"tblPr" => {
                        let (parsed, _) = properties::parse_table_properties(reader, buf)?;
                        tbl_pr = Some(parsed);
                    }
                    _ => xml::warn_unsupported_element("style", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                if local.as_slice() == b"basedOn" {
                    based_on = xml::optional_attr(e, b"val")?.map(StyleId);
                } else if local.as_slice() == b"name" {
                    name = xml::optional_attr(e, b"val")?;
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"style" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Some((
        id,
        Style {
            name,
            style_type,
            based_on,
            is_default,
            paragraph_properties: ppr,
            run_properties: rpr,
            table_properties: tbl_pr,
        },
    )))
}
