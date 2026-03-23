//! Parser for `word/numbering.xml` — parses definitions as-is, no resolution.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::*;
use crate::xml;

use super::properties;

/// Parse `word/numbering.xml` into raw `NumberingDefinitions`.
pub fn parse_numbering(data: &[u8]) -> Result<NumberingDefinitions> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut defs = NumberingDefinitions::default();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"abstractNum" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"abstractNumId")? {
                            let levels = parse_abstract_num(&mut reader, &mut buf)?;
                            defs.abstract_nums.insert(id, AbstractNumbering { levels });
                        }
                    }
                    b"num" => {
                        if let Some(num_id) = xml::optional_attr_i64(e, b"numId")? {
                            let instance = parse_num_instance(&mut reader, &mut buf)?;
                            defs.numbering_instances.insert(num_id, instance);
                        }
                    }
                    _ => xml::warn_unsupported_element("numbering", &local),
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(defs)
}

fn parse_abstract_num(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Vec<NumberingLevelDefinition>> {
    let mut levels = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"lvl" => {
                if let Some(ilvl) = xml::optional_attr_u32(e, b"ilvl")? {
                    levels.push(parse_level(reader, buf, ilvl as u8)?);
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"abstractNum" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(levels)
}

fn parse_level(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ilvl: u8,
) -> Result<NumberingLevelDefinition> {
    let mut format = NumberFormat::Decimal;
    let mut level_text = String::new();
    let mut start: Option<u32> = None;
    let mut indentation = Indentation::default();
    let mut run_properties: Option<RunProperties> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pPr" => {
                        let parsed = properties::parse_paragraph_properties(reader, buf)?;
                        indentation = parsed.properties.indentation;
                    }
                    b"rPr" => {
                        let (rpr, _) = properties::parse_run_properties(reader, buf)?;
                        run_properties = Some(rpr);
                    }
                    _ => xml::warn_unsupported_element("numbering-level", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"numFmt" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            format = parse_number_format(&val);
                        }
                    }
                    b"lvlText" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            level_text = val;
                        }
                    }
                    b"start" => {
                        start = xml::optional_attr_u32(e, b"val")?;
                    }
                    _ => xml::warn_unsupported_element("numbering-level", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"lvl" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(NumberingLevelDefinition {
        level: ilvl,
        format,
        level_text,
        start,
        indentation,
        run_properties,
    })
}

fn parse_num_instance(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<NumberingInstance> {
    let mut abstract_num_id: i64 = 0;
    let mut overrides = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                if local.as_slice() == b"lvlOverride" {
                    if let Some(ilvl) = xml::optional_attr_u32(e, b"ilvl")? {
                        if let Some(level) = parse_lvl_override(reader, buf, ilvl as u8)? {
                            overrides.push(level);
                        }
                    }
                }
            }
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"abstractNumId" => {
                if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                    abstract_num_id = val;
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"num" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(NumberingInstance {
        abstract_num_id,
        level_overrides: overrides,
    })
}

fn parse_lvl_override(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ilvl: u8,
) -> Result<Option<NumberingLevelDefinition>> {
    let mut result: Option<NumberingLevelDefinition> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"lvl" => {
                result = Some(parse_level(reader, buf, ilvl)?);
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"lvlOverride" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(result)
}

fn parse_number_format(val: &str) -> NumberFormat {
    match val {
        "decimal" => NumberFormat::Decimal,
        "upperRoman" => NumberFormat::UpperRoman,
        "lowerRoman" => NumberFormat::LowerRoman,
        "upperLetter" => NumberFormat::UpperLetter,
        "lowerLetter" => NumberFormat::LowerLetter,
        "bullet" => NumberFormat::Bullet,
        "ordinal" => NumberFormat::Ordinal,
        "cardinalText" => NumberFormat::CardinalText,
        "ordinalText" => NumberFormat::OrdinalText,
        "none" => NumberFormat::None,
        other => {
            log::warn!("unknown number format: {other}");
            NumberFormat::Decimal
        }
    }
}
