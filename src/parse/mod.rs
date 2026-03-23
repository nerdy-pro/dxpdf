mod archive;
mod numbering;
mod styles;
pub mod xml;

use std::collections::HashMap;

use std::rc::Rc;

use crate::error::Error;
use crate::model::{Block, Document, HeaderFooter, ImageStore, Inline, Spacing, StyleMap};

/// Parse a DOCX file (as raw bytes) into a `Document`.
///
/// Returns the raw parsed model with no style resolution or property merging.
/// Call [`resolve`] on the result before rendering.
pub fn parse(docx_bytes: &[u8]) -> Result<Document, Error> {
    let contents = archive::extract_docx_contents(docx_bytes)?;
    let mut document =
        xml::parse_document_xml_with_rels(&contents.document_xml, &contents.relationships)?;

    // Populate document defaults from auxiliary XML files
    let defs = &mut document.defaults;
    if let Some(dts) = contents.default_tab_stop {
        defs.tab_stop = dts;
    }
    if let Some(ref tf) = contents.theme_minor_font {
        defs.font_family = Rc::from(tf.as_str());
    }
    if let Some(dd) = &contents.doc_defaults {
        if let Some(fs) = dd.font_size {
            defs.font_size = fs;
        }
        if let Some(ref ff) = dd.font_family {
            defs.font_family = Rc::from(ff.as_str());
        }
        if let Some(sa) = dd.spacing_after {
            defs.spacing.after = Some(sa);
        }
        if let Some(sb) = dd.spacing_before {
            defs.spacing.before = Some(sb);
        }
        if let Some(sl) = dd.spacing_line {
            defs.spacing.line = Some(sl);
        }
        if let Some(slr) = dd.spacing_line_rule {
            defs.spacing.line_rule = slr;
        }
        if let Some(cm) = dd.cell_margins {
            defs.cell_margins = cm;
        }
        if let Some(tcs) = dd.table_cell_spacing {
            defs.table_cell_spacing = tcs;
        }
        if let Some(tb) = dd.table_borders {
            defs.table_borders = tb;
        }
        if !dd.styles.is_empty() {
            defs.styles = dd.styles.clone();
        }
    }
    if !contents.numbering.is_empty() {
        defs.numbering = contents.numbering.clone();
    }

    // Build the image store from relationships + media files
    document.images = build_image_store(
        &contents.relationships,
        &contents.media_files,
        &contents.header_footer.rels,
    );

    // Resolve headers/footers from XML files
    resolve_headers_footers(
        &mut document,
        &contents.relationships,
        &contents.header_footer.xml,
        &contents.header_footer.rels,
    );

    Ok(document)
}

/// Resolve styles and inherited properties on a parsed document.
///
/// Applies named styles, character styles, and paragraph default run
/// properties so that every run carries its final formatting.
/// Must be called before rendering.
pub fn resolve(document: &mut Document) {
    let styles = document.defaults.styles.clone();
    for section in &mut document.sections {
        apply_styles(&mut section.blocks, &styles);
        apply_paragraph_run_defaults(&mut section.blocks);
    }
}

// ============================================================
// Internal helpers
// ============================================================

/// Build an image store mapping relationship IDs to raw image bytes.
fn build_image_store(
    rels: &HashMap<String, String>,
    media: &HashMap<String, Vec<u8>>,
    hf_rels: &HashMap<String, HashMap<String, String>>,
) -> ImageStore {
    let mut store = ImageStore::new();

    for (rel_id, target) in rels {
        if let Some(bytes) = media.get(target) {
            store.insert(rel_id.clone(), bytes.clone());
        }
    }

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

/// Resolve headers and footers from XML files, linking them to sections.
fn resolve_headers_footers(
    doc: &mut Document,
    rels: &HashMap<String, String>,
    hf_xml: &HashMap<String, String>,
    hf_rels: &HashMap<String, HashMap<String, String>>,
) {
    for section in &mut doc.sections {
        let sect = &mut section.properties;
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
}

fn resolve_hf(
    rid: &str,
    rels: &HashMap<String, String>,
    hf_xml: &HashMap<String, String>,
    hf_rels: &HashMap<String, HashMap<String, String>>,
) -> Option<HeaderFooter> {
    let target = rels.get(rid)?;
    let xml_content = hf_xml.get(target)?;

    let empty_rels = HashMap::new();
    let own_rels = hf_rels.get(target).unwrap_or(&empty_rels);

    let mut hf = xml::parse_header_footer_xml_with_rels(xml_content, own_rels).ok()?;
    prefix_image_rel_ids(&mut hf.blocks, target);

    Some(hf)
}

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
                        let props = &mut p.properties;
                        if props.alignment.is_none() {
                            props.alignment = style.alignment;
                        }
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

/// Apply paragraph-level default run properties (`pPr/rPr`) to runs that lack
/// explicit font size or family. Only metric-affecting properties are inherited —
/// `pPr/rPr` defines the paragraph mark's formatting, not a general default
/// for bold/italic/color/underline.
fn apply_paragraph_run_defaults(blocks: &mut [Block]) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                if let Some(ref drp) = p.properties.default_run_props {
                    for run in &mut p.runs {
                        if let Inline::TextRun(tr) = run {
                            if tr.properties.font_size.is_none() {
                                tr.properties.font_size = drp.font_size;
                            }
                            if tr.properties.font_family.is_none() {
                                tr.properties.font_family = drp.font_family.clone();
                            }
                        }
                    }
                }
            }
            Block::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        apply_paragraph_run_defaults(&mut cell.blocks);
                    }
                }
            }
        }
    }
}
