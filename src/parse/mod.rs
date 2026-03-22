mod archive;
mod numbering;
mod styles;
pub mod xml;

use std::collections::HashMap;

use std::rc::Rc;

use crate::error::Error;
use crate::model::{
    Block, Document, HeaderFooter, ImageStore, Inline, SectionProperties, Spacing, StyleMap,
};

/// Build an image store mapping relationship IDs to raw image bytes.
/// Body images use the document relationships; header/footer images use their own rels
/// (prefixed with the header/footer filename to avoid collisions).
fn build_image_store(
    rels: &HashMap<String, String>,
    media: &HashMap<String, Vec<u8>>,
    hf_rels: &HashMap<String, HashMap<String, String>>,
) -> ImageStore {
    let mut store = ImageStore::new();

    // Body images: rel_id -> bytes
    for (rel_id, target) in rels {
        if let Some(bytes) = media.get(target) {
            store.insert(rel_id.clone(), bytes.clone());
        }
    }

    // Header/footer images: prefix with hf filename to avoid collisions
    // e.g., "header1.xml::rId1"
    for (hf_name, hf_rel_map) in hf_rels {
        for (rel_id, target) in hf_rel_map {
            if let Some(bytes) = media.get(target) {
                let prefixed_key = format!("{hf_name}::{rel_id}");
                store.insert(prefixed_key, bytes.clone());
            }
        }
    }

    store
}

/// Parse a DOCX file (as raw bytes) into a `Document`.
pub fn parse(docx_bytes: &[u8]) -> Result<Document, Error> {
    let contents = archive::extract_docx_contents(docx_bytes)?;
    let mut document =
        xml::parse_document_xml_with_rels(&contents.document_xml, &contents.relationships)?;
    if let Some(dts) = contents.default_tab_stop {
        document.default_tab_stop = dts;
    }
    // Apply theme font as default (can be overridden by docDefaults)
    if let Some(ref tf) = contents.theme_minor_font {
        document.default_font_family = Rc::from(tf.as_str());
    }
    if let Some(dd) = &contents.doc_defaults {
        if let Some(fs) = dd.font_size {
            document.default_font_size = fs;
        }
        if let Some(ref ff) = dd.font_family {
            document.default_font_family = Rc::from(ff.as_str());
        }
        if let Some(sa) = dd.spacing_after {
            document.default_spacing.after = Some(sa);
        }
        if let Some(sb) = dd.spacing_before {
            document.default_spacing.before = Some(sb);
        }
        if let Some(sl) = dd.spacing_line {
            document.default_spacing.line = Some(sl);
        }

        if let Some(slr) = dd.spacing_line_rule {
            document.default_spacing.line_rule = slr;
        }
        if let Some(cm) = dd.cell_margins {
            document.default_cell_margins = cm;
        }
        if let Some(tcs) = dd.table_cell_spacing {
            document.table_cell_spacing = tcs;
        }
        if let Some(tb) = dd.table_borders {
            document.default_table_borders = tb;
        }
        if !dd.styles.is_empty() {
            document.styles = dd.styles.clone();
        }
    }
    if !contents.numbering.is_empty() {
        document.numbering = contents.numbering.clone();
    }

    // Build the image store from relationships + media files
    document.images = build_image_store(
        &contents.relationships,
        &contents.media_files,
        &contents.header_footer_rels,
    );

    // Apply named styles to paragraphs and runs
    let styles = document.styles.clone();
    apply_styles(&mut document.blocks, &styles);

    // Resolve headers/footers on sections
    resolve_headers_footers(
        &mut document,
        &contents.relationships,
        &contents.header_footer_xml,
        &contents.header_footer_rels,
    );

    Ok(document)
}

/// Resolve headers and footers from XML files, linking them to sections.
fn resolve_headers_footers(
    doc: &mut Document,
    rels: &HashMap<String, String>,
    hf_xml: &HashMap<String, String>,
    hf_rels: &HashMap<String, HashMap<String, String>>,
) {
    // Resolve headers/footers on section properties (both mid-document and final)
    let mut all_sections: Vec<&mut SectionProperties> = Vec::new();
    for block in &mut doc.blocks {
        if let Block::Paragraph(p) = block {
            if let Some(ref mut sect) = p.section_properties {
                all_sections.push(sect);
            }
        }
    }
    if let Some(ref mut sect) = doc.final_section {
        all_sections.push(sect);
    }

    for sect in &mut all_sections {
        if let Some(rid) = sect.header_rel_id.take() {
            if let Some(hf) = resolve_hf(&rid, rels, hf_xml, hf_rels) {
                sect.header = Some(hf);
            }
            sect.header_rel_id = Some(rid);
        }
        if let Some(rid) = sect.footer_rel_id.take() {
            if let Some(hf) = resolve_hf(&rid, rels, hf_xml, hf_rels) {
                sect.footer = Some(hf);
            }
            sect.footer_rel_id = Some(rid);
        }
    }

    // Set document defaults from the first section that has header/footer
    for block in &doc.blocks {
        if let Block::Paragraph(p) = block {
            if let Some(ref sect) = p.section_properties {
                if doc.default_header.is_none() {
                    doc.default_header = sect.header.clone();
                }
                if doc.default_footer.is_none() {
                    doc.default_footer = sect.footer.clone();
                }
            }
        }
    }
    if let Some(ref sect) = doc.final_section {
        if doc.default_header.is_none() {
            doc.default_header = sect.header.clone();
        }
        if doc.default_footer.is_none() {
            doc.default_footer = sect.footer.clone();
        }
    }
}

/// Resolve a single header or footer from its relationship ID.
fn resolve_hf(
    rid: &str,
    rels: &HashMap<String, String>,
    hf_xml: &HashMap<String, String>,
    hf_rels: &HashMap<String, HashMap<String, String>>,
) -> Option<HeaderFooter> {
    // Map rId -> filename (e.g., "header2.xml")
    let target = rels.get(rid)?;
    let xml_content = hf_xml.get(target)?;

    // Parse with the header/footer's own rels so hyperlinks resolve correctly.
    // Image data is resolved separately via the ImageStore.
    let empty_rels = HashMap::new();
    let own_rels = hf_rels.get(target).unwrap_or(&empty_rels);

    // Prefix rel_ids for header/footer images so they match the ImageStore keys
    let mut hf = xml::parse_header_footer_xml_with_rels(xml_content, own_rels).ok()?;
    prefix_image_rel_ids(&mut hf.blocks, target);

    Some(hf)
}

/// Prefix image rel_ids in header/footer blocks so they match the ImageStore keys.
/// E.g., "rId1" in "header1.xml" becomes "header1.xml::rId1".
fn prefix_image_rel_ids(blocks: &mut [Block], hf_name: &str) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                for inline in &mut p.runs {
                    if let Inline::Image(img) = inline {
                        let prefixed = format!("{hf_name}::{}", img.rel_id.as_str());
                        img.rel_id = prefixed.into();
                    }
                }
                for float in &mut p.floats {
                    let prefixed = format!("{hf_name}::{}", float.rel_id.as_str());
                    float.rel_id = prefixed.into();
                }
            }
            Block::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        prefix_image_rel_ids(&mut cell.blocks, hf_name);
                    }
                }
            }
        }
    }
}

/// Apply a character style's run properties as defaults (direct formatting wins).
fn apply_run_style(rp: &mut crate::model::RunProperties, style: &crate::model::ResolvedRunStyle) {
    if !rp.bold && style.bold == Some(true) {
        rp.bold = true;
    }
    if !rp.italic && style.italic == Some(true) {
        rp.italic = true;
    }
    if !rp.underline && style.underline == Some(true) {
        rp.underline = true;
    }
    if rp.font_size.is_none() {
        rp.font_size = style.font_size;
    }
    if rp.font_family.is_none() {
        rp.font_family = style.font_family.clone();
    }
    if rp.color.is_none() {
        rp.color = style.color;
    }
}

/// Apply named styles to paragraphs — style properties are defaults
/// that direct formatting overrides.
fn apply_styles(blocks: &mut [Block], styles: &StyleMap) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                if let Some(ref sid) = p.properties.style_id {
                    if let Some(style) = styles.get(sid) {
                        // Merge style properties as defaults (direct formatting wins)
                        let props = &mut p.properties;
                        if props.alignment.is_none() {
                            props.alignment = style.alignment;
                        }
                        // Merge spacing field-by-field so direct line=276 doesn't
                        // block style's after=0
                        if let Some(ref style_sp) = style.spacing {
                            let sp = props.spacing.get_or_insert(Spacing::default());
                            if sp.before.is_none() {
                                sp.before = style_sp.before;
                            }
                            if sp.after.is_none() {
                                sp.after = style_sp.after;
                            }
                            if sp.line.is_none() {
                                sp.line = style_sp.line;
                                sp.line_rule = style_sp.line_rule;
                            }
                        }
                        if props.indentation.is_none() {
                            props.indentation = style.indentation;
                        }

                        // Apply style's run properties to all runs that lack
                        // direct formatting
                        for run in &mut p.runs {
                            if let Inline::TextRun(tr) = run {
                                let rp = &mut tr.properties;
                                if !rp.bold && style.run_props.bold == Some(true) {
                                    rp.bold = true;
                                }
                                if !rp.italic && style.run_props.italic == Some(true) {
                                    rp.italic = true;
                                }
                                if rp.font_size.is_none() {
                                    rp.font_size = style.run_props.font_size;
                                }
                                if rp.font_family.is_none() {
                                    rp.font_family = style.run_props.font_family.clone();
                                }
                                if rp.color.is_none() {
                                    rp.color = style.run_props.color;
                                }
                            }
                        }
                    }
                }
                // Apply character styles (w:rStyle) to individual runs
                for run in &mut p.runs {
                    if let Inline::TextRun(tr) = run {
                        if let Some(ref sid) = tr.properties.style_id {
                            if let Some(style) = styles.get(sid) {
                                apply_run_style(&mut tr.properties, &style.run_props);
                            }
                        }
                    }
                }
            }
            Block::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        apply_styles(&mut cell.blocks, styles);
                    }
                }
            }
        }
    }
}
