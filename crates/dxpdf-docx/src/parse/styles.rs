//! Parser for `word/styles.xml` — parses style definitions and resolves
//! `basedOn` inheritance chains to produce fully flattened styles.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::{ParseError, Result};
use crate::model::*;
use crate::xml;

use super::properties;

/// A style type as declared in the XML.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}

/// An unresolved style entry — may reference a parent via `based_on`.
#[derive(Clone, Debug)]
struct RawStyle {
    id: String,
    style_type: StyleType,
    based_on: Option<String>,
    paragraph_properties: ParagraphProperties,
    run_properties: RunProperties,
    table_properties: Option<TableProperties>,
}

/// Fully resolved style map — all `basedOn` chains flattened.
#[derive(Clone, Debug, Default)]
pub struct ResolvedStyles {
    pub paragraph: HashMap<String, ResolvedParagraphStyle>,
    pub character: HashMap<String, RunProperties>,
    pub table: HashMap<String, TableProperties>,
    pub default_paragraph: ParagraphProperties,
    pub default_run: RunProperties,
}

#[derive(Clone, Debug)]
pub struct ResolvedParagraphStyle {
    pub paragraph_properties: ParagraphProperties,
    pub run_properties: RunProperties,
}

/// Parse `word/styles.xml` and resolve all inheritance.
pub fn parse_styles(data: &[u8]) -> Result<ResolvedStyles> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut raw_styles: Vec<RawStyle> = Vec::new();
    let mut doc_defaults_ppr = ParagraphProperties::default();
    let mut doc_defaults_rpr = RunProperties::default();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"docDefaults" => {
                        parse_doc_defaults(
                            &mut reader,
                            &mut buf,
                            &mut doc_defaults_ppr,
                            &mut doc_defaults_rpr,
                        )?;
                    }
                    b"style" => {
                        if let Some(raw) = parse_raw_style(e, &mut reader, &mut buf)? {
                            raw_styles.push(raw);
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    resolve_styles(raw_styles, doc_defaults_ppr, doc_defaults_rpr)
}

fn parse_doc_defaults(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ppr: &mut ParagraphProperties,
    rpr: &mut RunProperties,
) -> Result<()> {
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"rPr" => {
                        let (parsed, _) = properties::parse_run_properties(reader, buf)?;
                        *rpr = parsed;
                    }
                    b"pPr" => {
                        let (parsed, _, _) = properties::parse_paragraph_properties(reader, buf)?;
                        *ppr = parsed;
                    }
                    _ => {}
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"docDefaults" => break,
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

fn parse_raw_style(
    start: &quick_xml::events::BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Option<RawStyle>> {
    let style_type = match xml::optional_attr(start, b"type")?.as_deref() {
        Some("paragraph") | None => StyleType::Paragraph,
        Some("character") => StyleType::Character,
        Some("table") => StyleType::Table,
        Some("numbering") => StyleType::Numbering,
        Some(_) => {
            xml::skip_to_end(reader, buf, b"style")?;
            return Ok(None);
        }
    };

    let id = match xml::optional_attr(start, b"styleId")? {
        Some(id) => id,
        None => {
            xml::skip_to_end(reader, buf, b"style")?;
            return Ok(None);
        }
    };

    let mut based_on: Option<String> = None;
    let mut ppr = ParagraphProperties::default();
    let mut rpr = RunProperties::default();
    let mut tbl_pr: Option<TableProperties> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pPr" => {
                        let (parsed, _, rp) = properties::parse_paragraph_properties(reader, buf)?;
                        ppr = parsed;
                        if let Some(rp) = rp {
                            rpr = rp;
                        }
                    }
                    b"rPr" => {
                        let (parsed, _) = properties::parse_run_properties(reader, buf)?;
                        rpr = parsed;
                    }
                    b"tblPr" => {
                        let (parsed, _) = properties::parse_table_properties(reader, buf)?;
                        tbl_pr = Some(parsed);
                    }
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                if local.as_slice() == b"basedOn" {
                    based_on = xml::optional_attr(e, b"val")?;
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"style" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Some(RawStyle {
        id,
        style_type,
        based_on,
        paragraph_properties: ppr,
        run_properties: rpr,
        table_properties: tbl_pr,
    }))
}

/// Resolve all `basedOn` inheritance chains.
fn resolve_styles(
    raw_styles: Vec<RawStyle>,
    doc_defaults_ppr: ParagraphProperties,
    doc_defaults_rpr: RunProperties,
) -> Result<ResolvedStyles> {
    let style_map: HashMap<&str, &RawStyle> =
        raw_styles.iter().map(|s| (s.id.as_str(), s)).collect();

    let mut resolved = ResolvedStyles {
        default_paragraph: doc_defaults_ppr.clone(),
        default_run: doc_defaults_rpr.clone(),
        ..Default::default()
    };

    for style in &raw_styles {
        match style.style_type {
            StyleType::Paragraph => {
                let mut chain = collect_chain(&style_map, &style.id)?;
                let mut final_ppr = doc_defaults_ppr.clone();
                let mut final_rpr = doc_defaults_rpr.clone();
                chain.reverse();
                for s in &chain {
                    merge_paragraph_properties(&mut final_ppr, &s.paragraph_properties);
                    merge_run_properties(&mut final_rpr, &s.run_properties);
                }
                resolved.paragraph.insert(
                    style.id.clone(),
                    ResolvedParagraphStyle {
                        paragraph_properties: final_ppr,
                        run_properties: final_rpr,
                    },
                );
            }
            StyleType::Character => {
                let chain = collect_chain(&style_map, &style.id)?;
                let mut final_rpr = doc_defaults_rpr.clone();
                for s in chain.iter().rev() {
                    merge_run_properties(&mut final_rpr, &s.run_properties);
                }
                resolved.character.insert(style.id.clone(), final_rpr);
            }
            StyleType::Table => {
                if let Some(ref tbl) = style.table_properties {
                    resolved.table.insert(style.id.clone(), tbl.clone());
                }
            }
            StyleType::Numbering => {}
        }
    }

    Ok(resolved)
}

fn collect_chain<'a>(
    map: &HashMap<&str, &'a RawStyle>,
    style_id: &str,
) -> Result<Vec<&'a RawStyle>> {
    let mut chain = Vec::new();
    let mut visited = std::collections::HashSet::new();
    let mut current_id = Some(style_id.to_string());

    while let Some(id) = current_id {
        if !visited.insert(id.clone()) {
            return Err(ParseError::CircularStyleInheritance(id));
        }
        if let Some(style) = map.get(id.as_str()) {
            chain.push(*style);
            current_id = style.based_on.clone();
        } else {
            break;
        }
    }

    Ok(chain)
}

fn merge_paragraph_properties(base: &mut ParagraphProperties, overlay: &ParagraphProperties) {
    let defaults = ParagraphProperties::default();
    if overlay.alignment != defaults.alignment {
        base.alignment = overlay.alignment;
    }
    if overlay.indentation != defaults.indentation {
        base.indentation = overlay.indentation;
    }
    if overlay.spacing != defaults.spacing {
        base.spacing = overlay.spacing;
    }
    if overlay.numbering.is_some() {
        base.numbering = overlay.numbering.clone();
    }
    if !overlay.tabs.is_empty() {
        base.tabs = overlay.tabs.clone();
    }
    if overlay.borders.is_some() {
        base.borders = overlay.borders.clone();
    }
    if overlay.shading.is_some() {
        base.shading = overlay.shading.clone();
    }
    if overlay.keep_next != defaults.keep_next {
        base.keep_next = overlay.keep_next;
    }
    if overlay.keep_lines != defaults.keep_lines {
        base.keep_lines = overlay.keep_lines;
    }
    if overlay.widow_control != defaults.widow_control {
        base.widow_control = overlay.widow_control;
    }
    if overlay.page_break_before != defaults.page_break_before {
        base.page_break_before = overlay.page_break_before;
    }
    if overlay.bidi != defaults.bidi {
        base.bidi = overlay.bidi;
    }
    if overlay.outline_level.is_some() {
        base.outline_level = overlay.outline_level;
    }
}

fn merge_run_properties(base: &mut RunProperties, overlay: &RunProperties) {
    let defaults = RunProperties::default();
    if overlay.fonts.ascii.is_some() {
        base.fonts.ascii = overlay.fonts.ascii.clone();
    }
    if overlay.fonts.high_ansi.is_some() {
        base.fonts.high_ansi = overlay.fonts.high_ansi.clone();
    }
    if overlay.fonts.east_asian.is_some() {
        base.fonts.east_asian = overlay.fonts.east_asian.clone();
    }
    if overlay.fonts.complex_script.is_some() {
        base.fonts.complex_script = overlay.fonts.complex_script.clone();
    }
    if overlay.font_size != defaults.font_size {
        base.font_size = overlay.font_size;
    }
    if overlay.bold != defaults.bold {
        base.bold = overlay.bold;
    }
    if overlay.italic != defaults.italic {
        base.italic = overlay.italic;
    }
    if overlay.underline != defaults.underline {
        base.underline = overlay.underline;
    }
    if overlay.strike != defaults.strike {
        base.strike = overlay.strike;
    }
    if overlay.color != defaults.color {
        base.color = overlay.color;
    }
    if overlay.highlight.is_some() {
        base.highlight = overlay.highlight;
    }
    if overlay.shading.is_some() {
        base.shading = overlay.shading.clone();
    }
    if overlay.vertical_align != defaults.vertical_align {
        base.vertical_align = overlay.vertical_align;
    }
    if overlay.spacing != defaults.spacing {
        base.spacing = overlay.spacing;
    }
    if overlay.kerning.is_some() {
        base.kerning = overlay.kerning;
    }
    if overlay.all_caps != defaults.all_caps {
        base.all_caps = overlay.all_caps;
    }
    if overlay.small_caps != defaults.small_caps {
        base.small_caps = overlay.small_caps;
    }
    if overlay.vanish != defaults.vanish {
        base.vanish = overlay.vanish;
    }
    if overlay.rtl != defaults.rtl {
        base.rtl = overlay.rtl;
    }
}
