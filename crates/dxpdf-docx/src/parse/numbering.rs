//! Parser for `word/numbering.xml` — parses definitions as-is, no resolution.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::*;
use crate::xml;

use super::{body, properties};

/// Parse `word/numbering.xml`. Enters `<w:numbering>`, parses until `</w:numbering>`.
pub fn parse_numbering(data: &[u8]) -> Result<NumberingDefinitions> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut defs = NumberingDefinitions::default();

    // Find <w:numbering> root element.
    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"numbering" => break,
            Event::Eof => return Ok(defs),
            _ => {}
        }
    }

    // Parse content scoped to </w:numbering>.
    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"abstractNum" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"abstractNumId")? {
                            let levels = parse_abstract_num(&mut reader, &mut buf)?;
                            defs.abstract_nums
                                .insert(AbstractNumId::new(id), AbstractNumbering { levels });
                        }
                    }
                    b"num" => {
                        if let Some(num_id) = xml::optional_attr_i64(e, b"numId")? {
                            let instance = parse_num_instance(&mut reader, &mut buf)?;
                            defs.numbering_instances
                                .insert(NumId::new(num_id), instance);
                        }
                    }
                    b"numPicBullet" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"numPicBulletId")? {
                            let bullet_id = NumPicBulletId::new(id);
                            let pic_bullet =
                                parse_num_pic_bullet(bullet_id, &mut reader, &mut buf)?;
                            defs.pic_bullets.insert(bullet_id, pic_bullet);
                        }
                    }
                    _ => xml::warn_unsupported_element("numbering", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"numbering" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"numbering")),
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
            Event::Eof => return Err(xml::unexpected_eof(b"abstractNum")),
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
    let mut format: Option<NumberFormat> = None;
    let mut level_text = String::new();
    let mut start: Option<u32> = None;
    let mut justification: Option<Alignment> = None;
    let mut indentation: Option<Indentation> = None;
    let mut run_properties: Option<RunProperties> = None;
    let mut lvl_pic_bullet_id: Option<NumPicBulletId> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"pPr" => {
                        let parsed = properties::parse_paragraph_properties(reader, buf)?;
                        indentation = parsed.properties.indentation;
                    }
                    b"rPr" => {
                        let (rpr, _) = properties::parse_run_properties(reader, buf)?;
                        run_properties = Some(rpr);
                    }
                    _ => xml::warn_unsupported_element("numbering-level", local),
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"numFmt" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            format = Some(parse_number_format(&val)?);
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
                    b"lvlJc" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            justification = Some(properties::parse_alignment(&val)?);
                        }
                    }
                    b"lvlPicBulletId" => {
                        lvl_pic_bullet_id =
                            xml::optional_attr_i64(e, b"val")?.map(NumPicBulletId::new);
                    }
                    _ => xml::warn_unsupported_element("numbering-level", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"lvl" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"lvl")),
            _ => {}
        }
    }

    Ok(NumberingLevelDefinition {
        level: ilvl,
        format,
        level_text,
        start,
        justification,
        indentation,
        run_properties,
        lvl_pic_bullet_id,
    })
}

fn parse_num_instance(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<NumberingInstance> {
    let mut abstract_num_id = AbstractNumId::new(0);
    let mut overrides = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                if local == b"lvlOverride" {
                    if let Some(ilvl) = xml::optional_attr_u32(e, b"ilvl")? {
                        if let Some(level) = parse_lvl_override(reader, buf, ilvl as u8)? {
                            overrides.push(level);
                        }
                    }
                }
            }
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"abstractNumId" => {
                if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                    abstract_num_id = AbstractNumId::new(val);
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"num" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"num")),
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
            Event::Eof => return Err(xml::unexpected_eof(b"lvlOverride")),
            _ => {}
        }
    }

    Ok(result)
}

/// §17.9.21: parse `w:numPicBullet` — picture bullet definition.
fn parse_num_pic_bullet(
    id: NumPicBulletId,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<NumPicBullet> {
    let mut pict = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"pict" => {
                        pict = Some(body::parse_pict(reader, buf)?);
                    }
                    _ => {
                        xml::warn_unsupported_element("numPicBullet", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"numPicBullet" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"numPicBullet")),
            _ => {}
        }
    }

    Ok(NumPicBullet { id, pict })
}

/// §17.18.59 ST_NumberFormat
fn parse_number_format(val: &str) -> Result<NumberFormat> {
    match val {
        "decimal" => Ok(NumberFormat::Decimal),
        "upperRoman" => Ok(NumberFormat::UpperRoman),
        "lowerRoman" => Ok(NumberFormat::LowerRoman),
        "upperLetter" => Ok(NumberFormat::UpperLetter),
        "lowerLetter" => Ok(NumberFormat::LowerLetter),
        "bullet" => Ok(NumberFormat::Bullet),
        "ordinal" => Ok(NumberFormat::Ordinal),
        "cardinalText" => Ok(NumberFormat::CardinalText),
        "ordinalText" => Ok(NumberFormat::OrdinalText),
        "none" => Ok(NumberFormat::None),
        other => Err(crate::error::ParseError::InvalidAttributeValue {
            attr: "numFmt/val".into(),
            value: other.into(),
            reason: "unsupported value per OOXML spec".into(),
        }),
    }
}
