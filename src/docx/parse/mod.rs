//! DOCX parsing — all submodules that parse specific parts of the OOXML package.

pub mod body;
pub mod drawing;
pub mod fonts;
pub mod notes;
pub mod numbering;
pub mod properties;
pub mod settings;
pub mod styles;
pub mod theme;
pub mod vml;

use std::collections::HashMap;

use crate::docx::error::{ParseError, Result};
use crate::docx::model::*;
use crate::docx::zip::{self, PackageContents, RelationshipType, Relationships};

/// Parse a DOCX file from raw bytes into a `Document`.
pub fn parse(data: &[u8]) -> Result<Document> {
    // Phase 1: Unzip
    let mut package = PackageContents::from_bytes(data)?;

    // Phase 1b: Find main document part via package-level rels
    let pkg_rels_data = package.require_part("_rels/.rels")?;
    let pkg_rels = Relationships::parse(pkg_rels_data)?;
    let doc_rel = pkg_rels
        .find_by_type(&RelationshipType::OfficeDocument)
        .ok_or_else(|| ParseError::MissingPart("officeDocument relationship".into()))?;
    let doc_path = zip::resolve_target("", &doc_rel.target);
    let doc_dir = zip::part_directory(&doc_path);

    // Phase 1c: Parse document-level relationships
    let doc_rels_path = zip::rels_path_for(&doc_path);
    let doc_rels = if let Some(data) = package.get_part(&doc_rels_path) {
        Relationships::parse(data)?
    } else {
        Relationships::default()
    };

    // Phase 2: Parse auxiliary parts

    // Theme
    let theme = if let Some(theme_rel) = doc_rels.find_by_type(&RelationshipType::Theme) {
        let theme_path = zip::resolve_target(doc_dir, &theme_rel.target);
        if let Some(data) = package.get_part(&theme_path) {
            Some(theme::parse_theme(data)?)
        } else {
            None
        }
    } else {
        None
    };

    // Styles
    let style_sheet = if let Some(styles_rel) = doc_rels.find_by_type(&RelationshipType::Styles) {
        let styles_path = zip::resolve_target(doc_dir, &styles_rel.target);
        let data = package.require_part(&styles_path)?;
        styles::parse_styles(data)?
    } else {
        StyleSheet::default()
    };

    // Numbering
    let numbering_defs = if let Some(num_rel) = doc_rels.find_by_type(&RelationshipType::Numbering)
    {
        let num_path = zip::resolve_target(doc_dir, &num_rel.target);
        if let Some(data) = package.get_part(&num_path) {
            numbering::parse_numbering(data)?
        } else {
            NumberingDefinitions::default()
        }
    } else {
        NumberingDefinitions::default()
    };

    // Settings
    let doc_settings =
        if let Some(settings_rel) = doc_rels.find_by_type(&RelationshipType::Settings) {
            let settings_path = zip::resolve_target(doc_dir, &settings_rel.target);
            if let Some(data) = package.get_part(&settings_path) {
                settings::parse_settings(data)?
            } else {
                DocumentSettings::default()
            }
        } else {
            DocumentSettings::default()
        };

    // Phase 2b: Extract media
    let mut media = HashMap::new();
    for rel in doc_rels.filter_by_type(&RelationshipType::Image) {
        let media_path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.take_part(&media_path) {
            let fmt = ImageFormat::detect(&rel.target, &data);
            media.insert(rel.id.clone(), (data, fmt));
        }
    }

    // §17.9.21: extract picture bullet images from numbering.xml relationships.
    if let Some(num_rel) = doc_rels.find_by_type(&RelationshipType::Numbering) {
        let num_path = zip::resolve_target(doc_dir, &num_rel.target);
        let num_dir = zip::part_directory(&num_path);
        let num_rels_path = zip::rels_path_for(&num_path);
        if let Some(rels_data) = package.get_part(&num_rels_path) {
            let num_rels = Relationships::parse(rels_data)?;
            for rel in num_rels.filter_by_type(&RelationshipType::Image) {
                let img_path = zip::resolve_target(num_dir, &rel.target);
                if let Some(data) = package.take_part(&img_path) {
                    let fmt = ImageFormat::detect(&rel.target, &data);
                    media.insert(rel.id.clone(), (data, fmt));
                }
            }
        }
    }

    // Phase 2c: Parse embedded fonts from fontTable
    let embedded_fonts = if let Some(ft_rel) = doc_rels.find_by_type(&RelationshipType::FontTable) {
        let ft_path = zip::resolve_target(doc_dir, &ft_rel.target);
        let ft_dir = zip::part_directory(&ft_path);
        let ft_rels_path = zip::rels_path_for(&ft_path);
        // Take data out to avoid overlapping borrows.
        let ft_data = package.take_part(&ft_path);
        let ft_rels_data = package.take_part(&ft_rels_path);
        if let Some(ft_data) = ft_data {
            let ft_rels = if let Some(rd) = ft_rels_data {
                Relationships::parse(&rd)?
            } else {
                Relationships::default()
            };
            fonts::parse_embedded_fonts(&ft_data, &ft_rels, &mut package, ft_dir)?
        } else {
            Vec::new()
        }
    } else {
        Vec::new()
    };

    // Phase 3: Parse document body
    let doc_data = package.require_part(&doc_path)?;
    let (mut body_blocks, final_section) = body::parse_body(doc_data)?;

    // Phase 4: Parse headers and footers
    let mut headers = HashMap::new();
    let mut footers = HashMap::new();

    for rel in doc_rels.filter_by_type(&RelationshipType::Header) {
        let path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.get_part(&path) {
            let blocks = body::parse_blocks(data)?;
            headers.insert(rel.id.clone(), blocks);

            // Extract images from header rels
            let hf_rels_path = zip::rels_path_for(&path);
            if let Some(rd) = package.get_part(&hf_rels_path) {
                let hf_rels = Relationships::parse(rd)?;
                for img_rel in hf_rels.filter_by_type(&RelationshipType::Image) {
                    let img_path = zip::resolve_target(zip::part_directory(&path), &img_rel.target);
                    if let Some(img_data) = package.take_part(&img_path) {
                        let fmt = ImageFormat::detect(&img_rel.target, &img_data);
                        media.insert(img_rel.id.clone(), (img_data, fmt));
                    }
                }
            }
        }
    }

    for rel in doc_rels.filter_by_type(&RelationshipType::Footer) {
        let path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.get_part(&path) {
            let blocks = body::parse_blocks(data)?;
            footers.insert(rel.id.clone(), blocks);

            let hf_rels_path = zip::rels_path_for(&path);
            if let Some(rd) = package.get_part(&hf_rels_path) {
                let hf_rels = Relationships::parse(rd)?;
                for img_rel in hf_rels.filter_by_type(&RelationshipType::Image) {
                    let img_path = zip::resolve_target(zip::part_directory(&path), &img_rel.target);
                    if let Some(img_data) = package.take_part(&img_path) {
                        let fmt = ImageFormat::detect(&img_rel.target, &img_data);
                        media.insert(img_rel.id.clone(), (img_data, fmt));
                    }
                }
            }
        }
    }

    // Phase 4b: Parse footnotes and endnotes
    let mut footnotes = if let Some(fn_rel) = doc_rels.find_by_type(&RelationshipType::Footnotes) {
        let path = zip::resolve_target(doc_dir, &fn_rel.target);
        if let Some(data) = package.get_part(&path) {
            notes::parse_notes(data, "footnote")?
        } else {
            HashMap::new()
        }
    } else {
        HashMap::new()
    };

    let mut endnotes = if let Some(en_rel) = doc_rels.find_by_type(&RelationshipType::Endnotes) {
        let path = zip::resolve_target(doc_dir, &en_rel.target);
        if let Some(data) = package.get_part(&path) {
            notes::parse_notes(data, "endnote")?
        } else {
            HashMap::new()
        }
    } else {
        HashMap::new()
    };

    // Phase 5: Resolve hyperlink RelIds to actual URLs.
    resolve_hyperlinks(&mut body_blocks, &doc_rels);
    for blocks in headers.values_mut() {
        resolve_hyperlinks(blocks, &doc_rels);
    }
    for blocks in footers.values_mut() {
        resolve_hyperlinks(blocks, &doc_rels);
    }
    for blocks in footnotes.values_mut() {
        resolve_hyperlinks(blocks, &doc_rels);
    }
    for blocks in endnotes.values_mut() {
        resolve_hyperlinks(blocks, &doc_rels);
    }

    // Phase 6: Assemble
    Ok(Document {
        settings: doc_settings,
        theme,
        styles: style_sheet,
        numbering: numbering_defs,
        body: body_blocks,
        final_section,
        headers,
        footers,
        footnotes,
        endnotes,
        media,
        embedded_fonts,
    })
}

/// Walk blocks and resolve `HyperlinkTarget::External(RelId)` to actual URLs.
fn resolve_hyperlinks(blocks: &mut [Block], rels: &crate::docx::zip::Relationships) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => resolve_hyperlinks_in_inlines(&mut p.content, rels),
            Block::Table(t) => {
                for row in &mut t.rows {
                    for cell in &mut row.cells {
                        resolve_hyperlinks(&mut cell.content, rels);
                    }
                }
            }
            _ => {}
        }
    }
}

fn resolve_hyperlinks_in_inlines(
    inlines: &mut [crate::model::Inline],
    rels: &crate::docx::zip::Relationships,
) {
    use crate::model::{HyperlinkTarget, Inline, RelId};

    for inline in inlines {
        match inline {
            Inline::Hyperlink(link) => {
                // Resolve External(RelId) to the actual URL.
                if let HyperlinkTarget::External(ref rel_id) = link.target {
                    if let Some(rel) = rels.find_by_id(rel_id.as_str()) {
                        link.target = HyperlinkTarget::External(RelId::new(&rel.target));
                    }
                }
                // Recurse into hyperlink content.
                resolve_hyperlinks_in_inlines(&mut link.content, rels);
            }
            Inline::Field(field) => {
                resolve_hyperlinks_in_inlines(&mut field.content, rels);
            }
            Inline::AlternateContent(ac) => {
                if let Some(ref mut fallback) = ac.fallback {
                    resolve_hyperlinks_in_inlines(fallback, rels);
                }
            }
            _ => {}
        }
    }
}
