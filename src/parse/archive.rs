use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;

/// Document-wide default properties from `word/styles.xml`.
#[derive(Debug)]
pub struct DocDefaults {
    /// Default font size in half-points.
    pub font_size: Option<u32>,
    /// Default font family.
    pub font_family: Option<String>,
    /// Default paragraph spacing after in twips.
    pub spacing_after: Option<u32>,
    /// Default paragraph spacing before in twips.
    pub spacing_before: Option<u32>,
    /// Default line spacing in twips.
    pub spacing_line: Option<u32>,
    /// Default table cell margins from the table grid style.
    pub cell_margins: Option<crate::model::CellMargins>,
    /// Default paragraph spacing inside table cells (from table grid style).
    pub table_cell_spacing: Option<crate::model::Spacing>,
    pub table_borders: Option<crate::model::TableBorders>,
    pub styles: crate::model::StyleMap,
}

/// Everything extracted from the DOCX ZIP archive needed for conversion.
#[derive(Debug)]
pub struct DocxContents {
    pub document_xml: String,
    /// Relationship ID -> target path (relative to `word/`).
    pub relationships: HashMap<String, String>,
    /// Path (relative to `word/`, e.g. "media/image1.png") -> raw bytes.
    pub media_files: HashMap<String, Vec<u8>>,
    /// Default tab stop interval in twips (from `word/settings.xml`).
    pub default_tab_stop: Option<u32>,
    /// Document-wide default run properties from `word/styles.xml`.
    pub doc_defaults: Option<DocDefaults>,
    /// Header/footer XML contents keyed by relationship ID.
    pub header_footer_xml: HashMap<String, String>,
    /// Header/footer relationships (rId -> media targets).
    pub header_footer_rels: HashMap<String, HashMap<String, String>>,
    /// Minor (body) font name from theme (e.g., "Calibri").
    pub theme_minor_font: Option<String>,
    /// Numbering definitions from `word/numbering.xml`.
    pub numbering: crate::model::NumberingMap,
}

/// Extract document XML, relationships, and media files from a DOCX archive.
pub fn extract_docx_contents(docx_bytes: &[u8]) -> Result<DocxContents, Error> {
    let cursor = std::io::Cursor::new(docx_bytes);
    let mut archive = zip::ZipArchive::new(cursor)?;

    // 1. Extract word/document.xml
    let document_xml = {
        let mut file = archive
            .by_name("word/document.xml")
            .map_err(|_| Error::MissingEntry("word/document.xml".into()))?;
        let mut xml = String::new();
        file.read_to_string(&mut xml)?;
        xml
    };

    // 2. Extract relationships from word/_rels/document.xml.rels
    let relationships = match archive.by_name("word/_rels/document.xml.rels") {
        Ok(mut file) => {
            let mut rels_xml = String::new();
            file.read_to_string(&mut rels_xml)?;
            parse_relationships(&rels_xml)?
        }
        Err(_) => HashMap::new(),
    };

    // 3. Extract all media files (word/media/*)
    let mut media_files = HashMap::new();
    let media_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let file = archive.by_index(i).ok()?;
            let name = file.name().to_string();
            if name.starts_with("word/media/") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for name in media_names {
        let mut file = archive.by_name(&name)?;
        let mut data = Vec::new();
        file.read_to_end(&mut data)?;
        // Store relative to word/ (e.g., "media/image1.png")
        let rel_path = name.strip_prefix("word/").unwrap_or(&name).to_string();
        media_files.insert(rel_path, data);
    }

    // 4. Extract defaultTabStop from word/settings.xml
    let default_tab_stop = match archive.by_name("word/settings.xml") {
        Ok(mut file) => {
            let mut settings_xml = String::new();
            file.read_to_string(&mut settings_xml)?;
            parse_default_tab_stop(&settings_xml)
        }
        Err(_) => None,
    };

    // 5. Extract document defaults from word/styles.xml
    let (mut doc_defaults, parsed_styles) = match archive.by_name("word/styles.xml") {
        Ok(mut file) => {
            let mut styles_xml = String::new();
            file.read_to_string(&mut styles_xml)?;
            (parse_doc_defaults(&styles_xml), parse_styles(&styles_xml))
        }
        Err(_) => (None, crate::model::StyleMap::new()),
    };
    if let Some(ref mut dd) = doc_defaults {
        dd.styles = parsed_styles;
    }

    // 6. Extract numbering definitions from word/numbering.xml
    let numbering = match archive.by_name("word/numbering.xml") {
        Ok(mut file) => {
            let mut xml = String::new();
            file.read_to_string(&mut xml)?;
            parse_numbering(&xml)
        }
        Err(_) => crate::model::NumberingMap::new(),
    };

    // 7. Extract theme fonts from word/theme/theme1.xml
    let (theme_minor_font, _theme_major_font) = extract_theme_fonts(&mut archive);

    // 7. Extract header/footer XML files
    let mut header_footer_xml = HashMap::new();
    let mut header_footer_rels = HashMap::new();
    let hf_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let file = archive.by_index(i).ok()?;
            let name = file.name().to_string();
            if (name.starts_with("word/header") || name.starts_with("word/footer"))
                && name.ends_with(".xml")
                && !name.contains("_rels")
            {
                Some(name)
            } else {
                None
            }
        })
        .collect();
    for name in &hf_names {
        if let Ok(mut file) = archive.by_name(name) {
            let mut xml = String::new();
            if file.read_to_string(&mut xml).is_ok() {
                let short = name.strip_prefix("word/").unwrap_or(name).to_string();
                header_footer_xml.insert(short, xml);
            }
        }
    }
    // Extract rels for each header/footer (for images inside headers)
    let hf_rel_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let file = archive.by_index(i).ok()?;
            let name = file.name().to_string();
            if name.starts_with("word/_rels/header") || name.starts_with("word/_rels/footer") {
                Some(name)
            } else {
                None
            }
        })
        .collect();
    for name in &hf_rel_names {
        if let Ok(mut file) = archive.by_name(name) {
            let mut xml = String::new();
            if file.read_to_string(&mut xml).is_ok() {
                if let Ok(rels) = parse_relationships(&xml) {
                    // Key by the header/footer filename (e.g., "header2.xml")
                    let hf_name = name
                        .strip_prefix("word/_rels/")
                        .unwrap_or(name)
                        .strip_suffix(".rels")
                        .unwrap_or(name)
                        .to_string();
                    header_footer_rels.insert(hf_name, rels);
                }
            }
        }
    }

    Ok(DocxContents {
        document_xml,
        relationships,
        media_files,
        header_footer_xml,
        header_footer_rels,
        default_tab_stop,
        doc_defaults,
        theme_minor_font,
        numbering,
    })
}

/// Extract minor and major font names from the theme file.
fn extract_theme_fonts(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> (Option<String>, Option<String>) {
    let mut minor = None;
    let mut major = None;

    let theme_xml = match archive.by_name("word/theme/theme1.xml") {
        Ok(mut file) => {
            let mut xml = String::new();
            if file.read_to_string(&mut xml).is_ok() {
                xml
            } else {
                return (None, None);
            }
        }
        Err(_) => return (None, None),
    };

    let mut reader = Reader::from_str(&theme_xml);
    let mut in_minor = false;
    let mut in_major = false;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"minorFont" => in_minor = true,
                    b"majorFont" => in_major = true,
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"minorFont" => in_minor = false,
                    b"majorFont" => in_major = false,
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"latin" {
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == b"typeface" {
                            let val =
                                String::from_utf8_lossy(&attr.value).into_owned();
                            if in_minor && minor.is_none() {
                                minor = Some(val);
                            } else if in_major && major.is_none() {
                                major = Some(val);
                            }
                        }
                    }
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    (minor, major)
}

/// Parse `word/_rels/document.xml.rels` into a map of relationship ID -> target path.
fn parse_relationships(xml: &str) -> Result<HashMap<String, String>, Error> {
    let mut reader = Reader::from_str(xml);
    let mut rels = HashMap::new();

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            Event::Empty(ref e) | Event::Start(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"Relationship" {
                    let mut id = None;
                    let mut target = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        match key {
                            b"Id" => {
                                id = Some(
                                    String::from_utf8_lossy(&attr.value).into_owned(),
                                );
                            }
                            b"Target" => {
                                target = Some(
                                    String::from_utf8_lossy(&attr.value).into_owned(),
                                );
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        rels.insert(id, target);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(rels)
}

/// Parse `word/styles.xml` to extract document default font size, family, and spacing.
fn parse_doc_defaults(xml: &str) -> Option<DocDefaults> {
    let mut reader = Reader::from_str(xml);
    let mut in_doc_defaults = false;
    let mut in_rpr_default = false;
    let mut in_ppr_default = false;
    let mut font_size = None;
    let mut font_family = None;
    let mut spacing_after = None;
    let mut spacing_before = None;
    let mut spacing_line = None;
    let mut cell_margins: Option<crate::model::CellMargins> = None;
    let mut in_tbl_cell_mar = false;
    let mut table_cell_spacing: Option<crate::model::Spacing> = None;
    let mut table_borders: Option<crate::model::TableBorders> = None;
    let mut in_table_style = false;
    let mut in_table_style_ppr = false;
    let mut in_table_style_borders = false;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"docDefaults" => in_doc_defaults = true,
                    b"rPrDefault" if in_doc_defaults => in_rpr_default = true,
                    b"pPrDefault" if in_doc_defaults => in_ppr_default = true,
                    b"tblCellMar" => in_tbl_cell_mar = true,
                    b"style" => {
                        // Check if this is a table style
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == b"type"
                                && attr.value.as_ref() == b"table"
                            {
                                in_table_style = true;
                            }
                        }
                    }
                    b"pPr" if in_table_style => {
                        in_table_style_ppr = true;
                    }
                    b"tblBorders" if in_table_style => {
                        in_table_style_borders = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"docDefaults" => {
                        in_doc_defaults = false;
                        in_rpr_default = false;
                        in_ppr_default = false;
                    }
                    b"rPrDefault" => in_rpr_default = false,
                    b"pPrDefault" => in_ppr_default = false,
                    b"tblCellMar" => in_tbl_cell_mar = false,
                    b"style" => {
                        in_table_style = false;
                        in_table_style_ppr = false;
                    }
                    b"pPr" if in_table_style => {
                        in_table_style_ppr = false;
                    }
                    b"tblBorders" => {
                        in_table_style_borders = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_rpr_default {
                    match local {
                        b"sz" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == b"val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    font_size = val.parse().ok();
                                }
                            }
                        }
                        b"rFonts" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == b"ascii" || key == b"hAnsi" {
                                    let val =
                                        String::from_utf8_lossy(&attr.value).into_owned();
                                    if font_family.is_none() && !val.is_empty() {
                                        font_family = Some(val);
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
                if in_ppr_default && local == b"spacing" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value);
                        match key {
                            b"after" => spacing_after = val.parse().ok(),
                            b"before" => spacing_before = val.parse().ok(),
                            b"line" => spacing_line = val.parse().ok(),
                            _ => {}
                        }
                    }
                }
                if in_tbl_cell_mar && cell_margins.is_none() {
                    // Parse margin children (top/bottom/left/right)
                    let w_type_attr = e.attributes().flatten().find(|a| {
                        local_name(a.key.as_ref()) == b"type"
                    });
                    let is_dxa = w_type_attr
                        .map(|a| a.value.as_ref() == b"dxa")
                        .unwrap_or(false);
                    if is_dxa {
                        let w_val: Option<u32> = e.attributes().flatten()
                            .find(|a| local_name(a.key.as_ref()) == b"w")
                            .and_then(|a| {
                                String::from_utf8_lossy(&a.value).parse().ok()
                            });
                        if let Some(val) = w_val {
                            let m = cell_margins
                                .get_or_insert(crate::model::CellMargins::default());
                            match local {
                                b"top" => m.top = val,
                                b"bottom" => m.bottom = val,
                                b"left" | b"start" => m.left = val,
                                b"right" | b"end" => m.right = val,
                                _ => {}
                            }
                        }
                    }
                }
                if in_table_style_borders && table_borders.is_none() {
                    // Parse border children (top/bottom/left/right/insideH/insideV)
                    let val = e.attributes().flatten()
                        .find(|a| local_name(a.key.as_ref()) == b"val")
                        .map(|a| String::from_utf8_lossy(&a.value).into_owned())
                        .unwrap_or_default();
                    let style = match val.as_str() {
                        "none" | "nil" => crate::model::BorderStyle::None,
                        "single" => crate::model::BorderStyle::Single,
                        "double" => crate::model::BorderStyle::Double,
                        "dashed" => crate::model::BorderStyle::Dashed,
                        "dotted" => crate::model::BorderStyle::Dotted,
                        _ => crate::model::BorderStyle::Single,
                    };
                    let size: u32 = e.attributes().flatten()
                        .find(|a| local_name(a.key.as_ref()) == b"sz")
                        .and_then(|a| String::from_utf8_lossy(&a.value).parse().ok())
                        .unwrap_or(4);
                    let color_str: String = e.attributes().flatten()
                        .find(|a| local_name(a.key.as_ref()) == b"color")
                        .map(|a| String::from_utf8_lossy(&a.value).into_owned())
                        .unwrap_or_default();
                    let color = if color_str == "auto" || color_str.is_empty() {
                        crate::model::Color { r: 0, g: 0, b: 0 }
                    } else {
                        crate::model::Color::from_hex(&color_str)
                            .unwrap_or(crate::model::Color { r: 0, g: 0, b: 0 })
                    };
                    let def = crate::model::BorderDef { style, size, color };
                    let b = table_borders.get_or_insert(crate::model::TableBorders::default());
                    match local {
                        b"top" => b.top = def,
                        b"bottom" => b.bottom = def,
                        b"left" | b"start" => b.left = def,
                        b"right" | b"end" => b.right = def,
                        b"insideH" => b.inside_h = def,
                        b"insideV" => b.inside_v = def,
                        _ => {}
                    }
                }
                if in_table_style_ppr && local == b"spacing"
                    && table_cell_spacing.is_none()
                {
                    let mut sp = crate::model::Spacing::default();
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value);
                        match key {
                            b"after" => sp.after = val.parse().ok(),
                            b"before" => sp.before = val.parse().ok(),
                            b"line" => sp.line = val.parse().ok(),
                            _ => {}
                        }
                    }
                    table_cell_spacing = Some(sp);
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    if font_size.is_some()
        || font_family.is_some()
        || spacing_after.is_some()
        || spacing_before.is_some()
        || spacing_line.is_some()
        || cell_margins.is_some()
        || table_cell_spacing.is_some()
        || table_borders.is_some()
    {
        Some(DocDefaults {
            font_size,
            font_family,
            spacing_after,
            spacing_before,
            spacing_line,
            cell_margins,
            table_cell_spacing,
            table_borders,
            styles: crate::model::StyleMap::new(), // filled by caller
        })
    } else {
        None
    }
}

/// Parse `word/settings.xml` to find `w:defaultTabStop` value.
fn parse_default_tab_stop(xml: &str) -> Option<u32> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"defaultTabStop" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"val" {
                            let val = String::from_utf8_lossy(&attr.value);
                            return val.parse().ok();
                        }
                    }
                }
            }
            Err(_) => break,
            _ => {}
        }
    }
    None
}

/// Parse numbering definitions from word/numbering.xml.
fn parse_numbering(xml: &str) -> crate::model::NumberingMap {
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
    let mut lvl_ind_left: u32 = 0;
    let mut lvl_ind_hanging: u32 = 0;
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
                        lvl_ind_left = 0;
                        lvl_ind_hanging = 0;
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
                                    indent_left: 0,
                                    indent_hanging: 0,
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
                                lvl_ind_left = v.parse().unwrap_or(0);
                            }
                            if let Some(v) = attr_val(e, b"hanging") {
                                lvl_ind_hanging = v.parse().unwrap_or(0);
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
            let end = xml[abs_pos..].find('>').map(|p| abs_pos + p).unwrap_or(xml.len());
            let tag = &xml[abs_pos..end + 1];
            // Extract numId
            if let Some(nid_start) = tag.find("w:numId=\"") {
                let nid_start = nid_start + 9;
                if let Some(nid_end) = tag[nid_start..].find('"') {
                    if let Ok(num_id) = tag[nid_start..nid_start + nid_end].parse::<u32>() {
                        // Find abstractNumId inside this num element
                        let num_end = xml[abs_pos..].find("</w:num>")
                            .map(|p| abs_pos + p)
                            .unwrap_or(xml.len());
                        let num_body = &xml[abs_pos..num_end];
                        if let Some(aid_pos) = num_body.find("w:abstractNumId") {
                            if let Some(val_start) = num_body[aid_pos..].find("w:val=\"") {
                                let vs = aid_pos + val_start + 7;
                                if let Some(val_end) = num_body[vs..].find('"') {
                                    if let Ok(abstract_id) = num_body[vs..vs + val_end].parse::<u32>() {
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

/// Parse named paragraph styles from word/styles.xml.
fn parse_styles(xml: &str) -> crate::model::StyleMap {
    use std::rc::Rc;

    let mut styles = crate::model::StyleMap::new();
    let mut reader = Reader::from_str(xml);
    let mut in_style = false;
    let mut style_id = String::new();
    let mut style_type = String::new();
    let mut in_ppr = false;
    let mut in_rpr = false;
    let mut based_on: Option<String> = None;

    // Current style properties being collected
    let mut alignment = None;
    let mut spacing = None;
    let mut indentation = None;
    let mut bold = None;
    let mut italic = None;
    let mut underline = None;
    let mut font_size = None;
    let mut font_family: Option<Rc<str>> = None;
    let mut color = None;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"style" => {
                        in_style = true;
                        style_id.clear();
                        style_type.clear();
                        based_on = None;
                        alignment = None;
                        spacing = None;
                        indentation = None;
                        bold = None;
                        italic = None;
                        underline = None;
                        font_size = None;
                        font_family = None;
                        color = None;
                        in_ppr = false;
                        in_rpr = false;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                b"styleId" => style_id = val.into_owned(),
                                b"type" => style_type = val.into_owned(),
                                _ => {}
                            }
                        }
                    }
                    b"pPr" if in_style => in_ppr = true,
                    b"rPr" if in_style => in_rpr = true,
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"style" => {
                        if in_style && (style_type == "paragraph" || style_type == "character") {
                            let mut resolved = crate::model::ResolvedParagraphStyle {
                                alignment,
                                spacing,
                                indentation,
                                run_props: crate::model::ResolvedRunStyle {
                                    bold,
                                    italic,
                                    underline,
                                    font_size,
                                    font_family: font_family.clone(),
                                    color,
                                },
                            };
                            // Inherit from basedOn style
                            if let Some(ref base_id) = based_on {
                                if let Some(base) = styles.get(base_id) {
                                    if resolved.alignment.is_none() {
                                        resolved.alignment = base.alignment;
                                    }
                                    if resolved.spacing.is_none() {
                                        resolved.spacing = base.spacing;
                                    }
                                    if resolved.indentation.is_none() {
                                        resolved.indentation = base.indentation;
                                    }
                                    if resolved.run_props.bold.is_none() {
                                        resolved.run_props.bold = base.run_props.bold;
                                    }
                                    if resolved.run_props.italic.is_none() {
                                        resolved.run_props.italic = base.run_props.italic;
                                    }
                                    if resolved.run_props.font_size.is_none() {
                                        resolved.run_props.font_size = base.run_props.font_size;
                                    }
                                    if resolved.run_props.font_family.is_none() {
                                        resolved.run_props.font_family =
                                            base.run_props.font_family.clone();
                                    }
                                    if resolved.run_props.color.is_none() {
                                        resolved.run_props.color = base.run_props.color;
                                    }
                                }
                            }
                            styles.insert(style_id.clone(), resolved);
                        }
                        in_style = false;
                    }
                    b"pPr" => in_ppr = false,
                    b"rPr" => in_rpr = false,
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if in_style => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_ppr {
                    match local {
                        b"jc" => {
                            if let Some(val) = attr_val(e, b"val") {
                                alignment = match val.as_str() {
                                    "left" | "start" => Some(crate::model::Alignment::Left),
                                    "center" => Some(crate::model::Alignment::Center),
                                    "right" | "end" => Some(crate::model::Alignment::Right),
                                    "both" | "justify" => Some(crate::model::Alignment::Justify),
                                    _ => None,
                                };
                            }
                        }
                        b"spacing" => {
                            let mut sp = crate::model::Spacing::default();
                            if let Some(v) = attr_val(e, b"before") { sp.before = v.parse().ok(); }
                            if let Some(v) = attr_val(e, b"after") { sp.after = v.parse().ok(); }
                            if let Some(v) = attr_val(e, b"line") { sp.line = v.parse().ok(); }
                            spacing = Some(sp);
                        }
                        b"ind" => {
                            let mut ind = crate::model::Indentation::default();
                            if let Some(v) = attr_val(e, b"left") { ind.left = v.parse().ok(); }
                            if let Some(v) = attr_val(e, b"right") { ind.right = v.parse().ok(); }
                            if let Some(v) = attr_val(e, b"firstLine") { ind.first_line = v.parse().ok(); }
                            if let Some(v) = attr_val(e, b"hanging") {
                                if let Ok(h) = v.parse::<i32>() { ind.first_line = Some(-h); }
                            }
                            indentation = Some(ind);
                        }
                        _ => {}
                    }
                }
                if in_rpr || (!in_ppr && !in_rpr && in_style) {
                    match local {
                        b"b" => bold = Some(true),
                        b"i" => italic = Some(true),
                        b"u" => underline = Some(true),
                        b"sz" => {
                            if let Some(v) = attr_val(e, b"val") {
                                font_size = v.parse().ok();
                            }
                        }
                        b"rFonts" => {
                            if let Some(v) = attr_val(e, b"ascii") {
                                font_family = Some(Rc::from(v.as_str()));
                            } else if let Some(v) = attr_val(e, b"hAnsi") {
                                font_family = Some(Rc::from(v.as_str()));
                            }
                        }
                        b"color" => {
                            if let Some(v) = attr_val(e, b"val") {
                                color = crate::model::Color::from_hex(&v);
                            }
                        }
                        b"basedOn" => {
                            based_on = attr_val(e, b"val");
                        }
                        _ => {}
                    }
                }
                // basedOn can be at style level (not inside pPr/rPr)
                if !in_ppr && !in_rpr && local == b"basedOn" {
                    based_on = attr_val(e, b"val");
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    styles
}

/// Get an attribute value by local name from a BytesStart element.
fn attr_val(e: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if local_name(attr.key.as_ref()) == name {
            return Some(String::from_utf8_lossy(&attr.value).into_owned());
        }
    }
    None
}

/// Strip namespace prefix from an element/attribute name.
fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|&b| b == b':')
        .map(|i| &name[i + 1..])
        .unwrap_or(name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_rels_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpeg"/>
        </Relationships>"#;
        let rels = parse_relationships(xml).unwrap();
        assert_eq!(rels.get("rId1").unwrap(), "media/image1.png");
        assert_eq!(rels.get("rId2").unwrap(), "media/image2.jpeg");
    }

    #[test]
    fn parse_empty_rels() {
        let xml = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let rels = parse_relationships(xml).unwrap();
        assert!(rels.is_empty());
    }
}
