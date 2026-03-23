use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;

/// Document-wide default properties from `word/styles.xml`.
#[derive(Debug)]
pub struct DocDefaults {
    /// Default font size in half-points.
    pub font_size: Option<crate::dimension::HalfPoints>,
    /// Default font family.
    pub font_family: Option<String>,
    /// Default paragraph spacing after in twips.
    pub spacing_after: Option<crate::dimension::Twips>,
    /// Default paragraph spacing before in twips.
    pub spacing_before: Option<crate::dimension::Twips>,
    /// Default line spacing in twips.
    pub spacing_line: Option<crate::dimension::Twips>,
    /// Default line spacing rule.
    pub spacing_line_rule: Option<crate::model::LineRule>,
    /// Default table cell margins from the table grid style.
    pub cell_margins: Option<crate::model::CellMargins>,
    /// Default paragraph spacing inside table cells (from table grid style).
    pub table_cell_spacing: Option<crate::model::Spacing>,
    pub table_borders: Option<crate::model::TableBorders>,
    pub styles: crate::model::StyleMap,
}

/// Header and footer XML content extracted from the archive.
#[derive(Debug, Default)]
pub struct HeaderFooterContents {
    /// Header/footer XML contents keyed by filename (e.g., "header1.xml").
    pub xml: HashMap<String, String>,
    /// Per-header/footer relationships (filename -> rId -> media target).
    pub rels: HashMap<String, HashMap<String, String>>,
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
    pub default_tab_stop: Option<crate::dimension::Twips>,
    /// Document-wide default run properties from `word/styles.xml`.
    pub doc_defaults: Option<DocDefaults>,
    /// Header/footer content and relationships.
    pub header_footer: HeaderFooterContents,
    /// Minor (body) font name from theme (e.g., "Calibri").
    pub theme_minor_font: Option<String>,
    /// Numbering definitions from `word/numbering.xml`.
    pub numbering: crate::model::NumberingMap,
}

/// Collect ZIP entry names that match a predicate.
fn collect_entry_names(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    predicate: impl Fn(&str) -> bool,
) -> Vec<String> {
    (0..archive.len())
        .filter_map(|i| {
            let file = archive.by_index(i).ok()?;
            let name = file.name().to_string();
            if predicate(&name) {
                Some(name)
            } else {
                None
            }
        })
        .collect()
}

/// Extract header/footer XML content and their relationships from the archive.
fn extract_headers_footers(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> HeaderFooterContents {
    let mut result = HeaderFooterContents::default();

    let hf_names = collect_entry_names(archive, |name| {
        (name.starts_with("word/header") || name.starts_with("word/footer"))
            && name.ends_with(".xml")
            && !name.contains("_rels")
    });
    for name in &hf_names {
        if let Ok(mut file) = archive.by_name(name) {
            let mut xml = String::new();
            if file.read_to_string(&mut xml).is_ok() {
                let short = name.strip_prefix("word/").unwrap_or(name).to_string();
                result.xml.insert(short, xml);
            }
        }
    }

    let hf_rel_names = collect_entry_names(archive, |name| {
        name.starts_with("word/_rels/header") || name.starts_with("word/_rels/footer")
    });
    for name in &hf_rel_names {
        if let Ok(mut file) = archive.by_name(name) {
            let mut xml = String::new();
            if file.read_to_string(&mut xml).is_ok() {
                if let Ok(rels) = parse_relationships(&xml) {
                    let hf_name = name
                        .strip_prefix("word/_rels/")
                        .unwrap_or(name)
                        .strip_suffix(".rels")
                        .unwrap_or(name)
                        .to_string();
                    result.rels.insert(hf_name, rels);
                }
            }
        }
    }

    result
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
    let media_names = collect_entry_names(&mut archive, |name| name.starts_with("word/media/"));

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
            (
                parse_doc_defaults(&styles_xml),
                super::styles::parse_styles(&styles_xml),
            )
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
            super::numbering::parse_numbering(&xml)
        }
        Err(_) => crate::model::NumberingMap::new(),
    };

    // 7. Extract theme fonts from word/theme/theme1.xml
    let (theme_minor_font, _theme_major_font) = extract_theme_fonts(&mut archive);

    // 8. Extract header/footer XML files and their relationships
    let header_footer = extract_headers_footers(&mut archive);

    Ok(DocxContents {
        document_xml,
        relationships,
        media_files,
        header_footer,
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
                            let val = String::from_utf8_lossy(&attr.value).into_owned();
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
                                id = Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                            b"Target" => {
                                target = Some(String::from_utf8_lossy(&attr.value).into_owned());
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

/// Accumulator for `parse_doc_defaults`.
#[derive(Default)]
struct DocDefaultsBuilder {
    // Run defaults (from docDefaults/rPrDefault)
    font_size: Option<crate::dimension::HalfPoints>,
    font_family: Option<String>,
    // Paragraph defaults (from docDefaults/pPrDefault)
    spacing_after: Option<crate::dimension::Twips>,
    spacing_before: Option<crate::dimension::Twips>,
    spacing_line: Option<crate::dimension::Twips>,
    spacing_line_rule: Option<crate::model::LineRule>,
    // Table style defaults
    cell_margins: Option<crate::model::CellMargins>,
    table_cell_spacing: Option<crate::model::Spacing>,
    table_borders: Option<crate::model::TableBorders>,
}

impl DocDefaultsBuilder {
    /// Handle an empty element inside `rPrDefault`.
    fn handle_rpr_element(&mut self, local: &[u8], e: &quick_xml::events::BytesStart<'_>) {
        match local {
            b"sz" => {
                for attr in e.attributes().flatten() {
                    if local_name(attr.key.as_ref()) == b"val" {
                        let val = String::from_utf8_lossy(&attr.value);
                        self.font_size = val
                            .parse::<i64>()
                            .ok()
                            .map(crate::dimension::HalfPoints::new);
                    }
                }
            }
            b"rFonts" => {
                for attr in e.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    if key == b"ascii" || key == b"hAnsi" {
                        let val = String::from_utf8_lossy(&attr.value).into_owned();
                        if self.font_family.is_none() && !val.is_empty() {
                            self.font_family = Some(val);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    /// Handle a spacing element inside `pPrDefault`.
    fn handle_ppr_spacing(&mut self, e: &quick_xml::events::BytesStart<'_>) {
        for attr in e.attributes().flatten() {
            let key = local_name(attr.key.as_ref());
            let val = String::from_utf8_lossy(&attr.value);
            match key {
                b"after" => {
                    self.spacing_after = val.parse::<i64>().ok().map(crate::dimension::Twips::new)
                }
                b"before" => {
                    self.spacing_before = val.parse::<i64>().ok().map(crate::dimension::Twips::new)
                }
                b"line" => {
                    self.spacing_line = val.parse::<i64>().ok().map(crate::dimension::Twips::new)
                }
                b"lineRule" => {
                    self.spacing_line_rule = match val.as_ref() {
                        "auto" => Some(crate::model::LineRule::Auto),
                        "exact" => Some(crate::model::LineRule::Exact),
                        "atLeast" => Some(crate::model::LineRule::AtLeast),
                        _ => None,
                    }
                }
                _ => {}
            }
        }
    }

    /// Handle a cell margin element inside `tblCellMar`.
    fn handle_cell_margin(&mut self, local: &[u8], e: &quick_xml::events::BytesStart<'_>) {
        let w_type_attr = e
            .attributes()
            .flatten()
            .find(|a| local_name(a.key.as_ref()) == b"type");
        let is_dxa = w_type_attr
            .map(|a| a.value.as_ref() == b"dxa")
            .unwrap_or(false);
        if is_dxa {
            let w_val: Option<crate::dimension::Twips> = e
                .attributes()
                .flatten()
                .find(|a| local_name(a.key.as_ref()) == b"w")
                .and_then(|a| String::from_utf8_lossy(&a.value).parse::<i64>().ok())
                .map(crate::dimension::Twips::new);
            if let Some(val) = w_val {
                let m = self
                    .cell_margins
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

    /// Handle a table style border element.
    fn handle_table_border(&mut self, local: &[u8], e: &quick_xml::events::BytesStart<'_>) {
        if let Ok(def) = super::xml::properties::parse_border_def(e) {
            let b = self
                .table_borders
                .get_or_insert(crate::model::TableBorders::default());
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
    }

    /// Handle a table style paragraph spacing element.
    fn handle_table_cell_spacing(&mut self, e: &quick_xml::events::BytesStart<'_>) {
        let mut sp = crate::model::Spacing::default();
        for attr in e.attributes().flatten() {
            let key = local_name(attr.key.as_ref());
            let val = String::from_utf8_lossy(&attr.value);
            match key {
                b"after" => sp.after = val.parse::<i64>().ok().map(crate::dimension::Twips::new),
                b"before" => sp.before = val.parse::<i64>().ok().map(crate::dimension::Twips::new),
                b"line" => sp.line = val.parse::<i64>().ok().map(crate::dimension::Twips::new),
                _ => {}
            }
        }
        self.table_cell_spacing = Some(sp);
    }

    fn finish(self) -> Option<DocDefaults> {
        if self.font_size.is_some()
            || self.font_family.is_some()
            || self.spacing_after.is_some()
            || self.spacing_before.is_some()
            || self.spacing_line.is_some()
            || self.cell_margins.is_some()
            || self.table_cell_spacing.is_some()
            || self.table_borders.is_some()
        {
            Some(DocDefaults {
                font_size: self.font_size,
                font_family: self.font_family,
                spacing_after: self.spacing_after,
                spacing_before: self.spacing_before,
                spacing_line: self.spacing_line,
                spacing_line_rule: self.spacing_line_rule,
                cell_margins: self.cell_margins,
                table_cell_spacing: self.table_cell_spacing,
                table_borders: self.table_borders,
                styles: crate::model::StyleMap::new(), // filled by caller
            })
        } else {
            None
        }
    }
}

/// Parse `word/styles.xml` to extract document default font size, family, and spacing.
fn parse_doc_defaults(xml: &str) -> Option<DocDefaults> {
    let mut reader = Reader::from_str(xml);
    let mut builder = DocDefaultsBuilder::default();
    let mut in_doc_defaults = false;
    let mut in_rpr_default = false;
    let mut in_ppr_default = false;
    let mut in_tbl_cell_mar = false;
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
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == b"type"
                                && attr.value.as_ref() == b"table"
                            {
                                in_table_style = true;
                            }
                        }
                    }
                    b"pPr" if in_table_style => in_table_style_ppr = true,
                    b"tblBorders" if in_table_style => in_table_style_borders = true,
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
                    b"pPr" if in_table_style => in_table_style_ppr = false,
                    b"tblBorders" => in_table_style_borders = false,
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_rpr_default {
                    builder.handle_rpr_element(local, e);
                }
                if in_ppr_default && local == b"spacing" {
                    builder.handle_ppr_spacing(e);
                }
                if in_tbl_cell_mar && builder.cell_margins.is_none() {
                    builder.handle_cell_margin(local, e);
                }
                if in_table_style_borders && builder.table_borders.is_none() {
                    builder.handle_table_border(local, e);
                }
                if in_table_style_ppr && local == b"spacing" && builder.table_cell_spacing.is_none()
                {
                    builder.handle_table_cell_spacing(e);
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    builder.finish()
}

/// Parse `word/settings.xml` to find `w:defaultTabStop` value.
fn parse_default_tab_stop(xml: &str) -> Option<crate::dimension::Twips> {
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
                            return val.parse::<i64>().ok().map(crate::dimension::Twips::new);
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

use super::xml::helpers::local_name;

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
