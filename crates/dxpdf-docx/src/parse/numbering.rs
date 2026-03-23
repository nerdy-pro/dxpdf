//! Parser for `word/numbering.xml` — abstract numbering definitions,
//! numbering instances, and resolution to concrete formats.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::*;
use crate::xml;

use super::properties;

/// A resolved numbering definition, keyed by (numId, ilvl).
pub type NumberingMap = HashMap<(i64, u8), NumberingProperties>;

#[derive(Clone, Debug)]
struct AbstractLevel {
    level: u8,
    format: NumberFormat,
    level_text: String,
    indentation: Indentation,
    run_properties: Option<RunProperties>,
}

pub fn parse_numbering(data: &[u8]) -> Result<NumberingMap> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut abstract_nums: HashMap<i64, Vec<AbstractLevel>> = HashMap::new();
    let mut num_instances: HashMap<i64, i64> = HashMap::new();
    let mut num_overrides: HashMap<i64, Vec<AbstractLevel>> = HashMap::new();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"abstractNum" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"abstractNumId")? {
                            let levels = parse_abstract_num(&mut reader, &mut buf)?;
                            abstract_nums.insert(id, levels);
                        }
                    }
                    b"num" => {
                        if let Some(num_id) = xml::optional_attr_i64(e, b"numId")? {
                            let (abstract_id, overrides) =
                                parse_num_instance(&mut reader, &mut buf)?;
                            if let Some(aid) = abstract_id {
                                num_instances.insert(num_id, aid);
                            }
                            if !overrides.is_empty() {
                                num_overrides.insert(num_id, overrides);
                            }
                        }
                    }
                    _ => xml::warn_unsupported_element("numbering", &local),
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    let mut result = NumberingMap::new();

    for (&num_id, &abstract_id) in &num_instances {
        if let Some(levels) = abstract_nums.get(&abstract_id) {
            for level in levels {
                let key = (num_id, level.level);
                result.insert(key, level_to_properties(level));
            }
        }
        if let Some(overrides) = num_overrides.get(&num_id) {
            for level in overrides {
                let key = (num_id, level.level);
                result.insert(key, level_to_properties(level));
            }
        }
    }

    Ok(result)
}

fn level_to_properties(level: &AbstractLevel) -> NumberingProperties {
    NumberingProperties {
        level: level.level,
        format: level.format,
        level_text: level.level_text.clone(),
        indent: level.indentation,
        run_properties: level.run_properties.clone(),
    }
}

fn parse_abstract_num(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<AbstractLevel>> {
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

fn parse_level(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>, ilvl: u8) -> Result<AbstractLevel> {
    let mut format = NumberFormat::Decimal;
    let mut level_text = String::new();
    let mut indentation = Indentation::default();
    let mut run_properties: Option<RunProperties> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pPr" => {
                        let (ppr, _, _) = properties::parse_paragraph_properties(reader, buf)?;
                        indentation = ppr.indentation;
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
                    b"lvlJc" => {
                        // We parse but don't store alignment on levels currently
                    }
                    _ => xml::warn_unsupported_element("numbering-level", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"lvl" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(AbstractLevel {
        level: ilvl,
        format,
        level_text,
        indentation,
        run_properties,
    })
}

fn parse_num_instance(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(Option<i64>, Vec<AbstractLevel>)> {
    let mut abstract_num_id: Option<i64> = None;
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
                abstract_num_id = xml::optional_attr_i64(e, b"val")?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"num" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((abstract_num_id, overrides))
}

fn parse_lvl_override(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ilvl: u8,
) -> Result<Option<AbstractLevel>> {
    let mut result: Option<AbstractLevel> = None;

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
