use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::Reader;

use super::archive::{attr_val, local_name};

pub(super) fn parse_numbering(xml: &str) -> crate::model::NumberingMap {
    use crate::model::{NumberFormat, NumberingDef, NumberingLevel};

    let mut reader = Reader::from_str(xml);
    let mut abstract_nums: HashMap<u32, Vec<NumberingLevel>> = HashMap::new();
    let mut num_to_abstract: HashMap<u32, u32> = HashMap::new();

    let mut in_abstract = false;
    let mut abstract_id: u32 = 0;
    let mut in_lvl = false;
    let mut lvl_ilvl: u32 = 0;
    let mut lvl_fmt = String::new();
    let mut lvl_text = String::new();
    let mut lvl_start: u32 = 1;
    let mut lvl_ind_left: crate::dimension::Twips = crate::dimension::Twips::new(0);
    let mut lvl_ind_hanging: crate::dimension::Twips = crate::dimension::Twips::new(0);
    let mut current_levels: Vec<NumberingLevel> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"abstractNum" => {
                        in_abstract = true;
                        current_levels.clear();
                        if let Some(v) = attr_val(e, b"abstractNumId") {
                            abstract_id = v.parse().unwrap_or(0);
                        }
                    }
                    b"lvl" if in_abstract => {
                        in_lvl = true;
                        lvl_fmt.clear();
                        lvl_text.clear();
                        lvl_start = 1;
                        lvl_ind_left = crate::dimension::Twips::new(0);
                        lvl_ind_hanging = crate::dimension::Twips::new(0);
                        if let Some(v) = attr_val(e, b"ilvl") {
                            lvl_ilvl = v.parse().unwrap_or(0);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"abstractNum" => {
                        abstract_nums.insert(abstract_id, current_levels.clone());
                        in_abstract = false;
                    }
                    b"lvl" => {
                        if in_lvl {
                            let format = match lvl_fmt.as_str() {
                                "bullet" => NumberFormat::Bullet(lvl_text.clone()),
                                "decimal" => NumberFormat::Decimal,
                                "lowerLetter" => NumberFormat::LowerLetter,
                                "upperLetter" => NumberFormat::UpperLetter,
                                "lowerRoman" => NumberFormat::LowerRoman,
                                "upperRoman" => NumberFormat::UpperRoman,
                                _ => NumberFormat::Decimal,
                            };
                            // Ensure levels vec is large enough
                            while current_levels.len() <= lvl_ilvl as usize {
                                current_levels.push(NumberingLevel {
                                    format: NumberFormat::Decimal,
                                    level_text: String::new(),
                                    start: 1,
                                    indent_left: crate::dimension::Twips::new(0),
                                    indent_hanging: crate::dimension::Twips::new(0),
                                });
                            }
                            current_levels[lvl_ilvl as usize] = NumberingLevel {
                                format,
                                level_text: lvl_text.clone(),
                                start: lvl_start,
                                indent_left: lvl_ind_left,
                                indent_hanging: lvl_ind_hanging,
                            };
                            in_lvl = false;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_lvl {
                    match local {
                        b"numFmt" => {
                            if let Some(v) = attr_val(e, b"val") {
                                lvl_fmt = v;
                            }
                        }
                        b"lvlText" => {
                            if let Some(v) = attr_val(e, b"val") {
                                lvl_text = v;
                            }
                        }
                        b"start" => {
                            if let Some(v) = attr_val(e, b"val") {
                                lvl_start = v.parse().unwrap_or(1);
                            }
                        }
                        b"ind" => {
                            if let Some(v) = attr_val(e, b"left") {
                                lvl_ind_left = v
                                    .parse::<i64>()
                                    .map(crate::dimension::Twips::new)
                                    .unwrap_or(crate::dimension::Twips::new(0));
                            }
                            if let Some(v) = attr_val(e, b"hanging") {
                                lvl_ind_hanging = v
                                    .parse::<i64>()
                                    .map(crate::dimension::Twips::new)
                                    .unwrap_or(crate::dimension::Twips::new(0));
                            }
                        }
                        _ => {}
                    }
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    // Second pass for num -> abstractNum mapping (simpler with regex)
    for m in regex_like_find_nums(xml) {
        num_to_abstract.insert(m.0, m.1);
    }

    // Build final map
    let mut result = crate::model::NumberingMap::new();
    for (num_id, abstract_id) in &num_to_abstract {
        if let Some(levels) = abstract_nums.get(abstract_id) {
            result.insert(
                *num_id,
                NumberingDef {
                    levels: levels.clone(),
                },
            );
        }
    }
    result
}

/// Simple extraction of num -> abstractNum mappings.
fn regex_like_find_nums(xml: &str) -> Vec<(u32, u32)> {
    let mut result = Vec::new();
    let bytes = xml.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        // Find <w:num w:numId="X">
        if let Some(pos) = xml[i..].find("<w:num ") {
            let abs_pos = i + pos;
            let end = xml[abs_pos..]
                .find('>')
                .map(|p| abs_pos + p)
                .unwrap_or(xml.len());
            let tag = &xml[abs_pos..end + 1];
            // Extract numId
            if let Some(nid_start) = tag.find("w:numId=\"") {
                let nid_start = nid_start + 9;
                if let Some(nid_end) = tag[nid_start..].find('"') {
                    if let Ok(num_id) = tag[nid_start..nid_start + nid_end].parse::<u32>() {
                        // Find abstractNumId inside this num element
                        let num_end = xml[abs_pos..]
                            .find("</w:num>")
                            .map(|p| abs_pos + p)
                            .unwrap_or(xml.len());
                        let num_body = &xml[abs_pos..num_end];
                        if let Some(aid_pos) = num_body.find("w:abstractNumId") {
                            if let Some(val_start) = num_body[aid_pos..].find("w:val=\"") {
                                let vs = aid_pos + val_start + 7;
                                if let Some(val_end) = num_body[vs..].find('"') {
                                    if let Ok(abstract_id) =
                                        num_body[vs..vs + val_end].parse::<u32>()
                                    {
                                        result.push((num_id, abstract_id));
                                    }
                                }
                            }
                        }
                    }
                }
            }
            i = end + 1;
        } else {
            break;
        }
    }
    result
}
