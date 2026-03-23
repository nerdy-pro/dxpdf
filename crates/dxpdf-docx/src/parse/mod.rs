//! DOCX parsing — all submodules that parse specific parts of the OOXML package.

pub mod body;
pub mod notes;
pub mod numbering;
pub mod properties;
pub mod settings;
pub mod styles;
pub mod theme;

use std::collections::HashMap;

use crate::error::{ParseError, Result};
use crate::model::*;
use crate::zip::{self, PackageContents, RelationshipType, Relationships};

/// Parse a DOCX file from raw bytes into a fully resolved `Document`.
pub fn parse(data: &[u8]) -> Result<Document> {
    // Phase 1: Unzip
    let package = PackageContents::from_bytes(data)?;

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
    let resolved_styles = if let Some(styles_rel) = doc_rels.find_by_type(&RelationshipType::Styles)
    {
        let styles_path = zip::resolve_target(doc_dir, &styles_rel.target);
        let data = package.require_part(&styles_path)?;
        styles::parse_styles(data)?
    } else {
        styles::ResolvedStyles::default()
    };

    // Numbering
    let numbering_map = if let Some(num_rel) = doc_rels.find_by_type(&RelationshipType::Numbering) {
        let num_path = zip::resolve_target(doc_dir, &num_rel.target);
        if let Some(data) = package.get_part(&num_path) {
            numbering::parse_numbering(data)?
        } else {
            numbering::NumberingMap::new()
        }
    } else {
        numbering::NumberingMap::new()
    };

    // Settings
    let mut doc_settings =
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

    // Apply docDefaults from styles to settings
    doc_settings.default_paragraph_properties = resolved_styles.default_paragraph.clone();
    doc_settings.default_run_properties = resolved_styles.default_run.clone();

    // Phase 2b: Extract media
    let mut media = HashMap::new();
    for rel in doc_rels.filter_by_type(&RelationshipType::Image) {
        let media_path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.get_part(&media_path) {
            media.insert(RelId(rel.id.clone()), data.to_vec());
        }
    }

    // Phase 3: Parse document body
    let ctx = body::ParseContext {
        styles: &resolved_styles,
        numbering: &numbering_map,
        rels: &doc_rels,
    };

    let doc_data = package.require_part(&doc_path)?;
    let (body_blocks, final_section) = body::parse_body(doc_data, &ctx)?;

    // Phase 4: Parse headers and footers
    let mut headers = HashMap::new();
    let mut footers = HashMap::new();

    for rel in doc_rels.filter_by_type(&RelationshipType::Header) {
        let path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.get_part(&path) {
            // Each header/footer may have its own rels
            let hf_rels_path = zip::rels_path_for(&path);
            let hf_rels = if let Some(rd) = package.get_part(&hf_rels_path) {
                Relationships::parse(rd)?
            } else {
                Relationships::default()
            };
            let hf_ctx = body::ParseContext {
                styles: &resolved_styles,
                numbering: &numbering_map,
                rels: &hf_rels,
            };
            let blocks = body::parse_blocks(data, &hf_ctx)?;
            headers.insert(RelId(rel.id.clone()), blocks);

            // Also extract any images from header rels
            for img_rel in hf_rels.filter_by_type(&RelationshipType::Image) {
                let img_path = zip::resolve_target(zip::part_directory(&path), &img_rel.target);
                if let Some(img_data) = package.get_part(&img_path) {
                    media.insert(RelId(img_rel.id.clone()), img_data.to_vec());
                }
            }
        }
    }

    for rel in doc_rels.filter_by_type(&RelationshipType::Footer) {
        let path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.get_part(&path) {
            let hf_rels_path = zip::rels_path_for(&path);
            let hf_rels = if let Some(rd) = package.get_part(&hf_rels_path) {
                Relationships::parse(rd)?
            } else {
                Relationships::default()
            };
            let hf_ctx = body::ParseContext {
                styles: &resolved_styles,
                numbering: &numbering_map,
                rels: &hf_rels,
            };
            let blocks = body::parse_blocks(data, &hf_ctx)?;
            footers.insert(RelId(rel.id.clone()), blocks);

            for img_rel in hf_rels.filter_by_type(&RelationshipType::Image) {
                let img_path = zip::resolve_target(zip::part_directory(&path), &img_rel.target);
                if let Some(img_data) = package.get_part(&img_path) {
                    media.insert(RelId(img_rel.id.clone()), img_data.to_vec());
                }
            }
        }
    }

    // Phase 4b: Parse footnotes and endnotes
    let footnotes = if let Some(fn_rel) = doc_rels.find_by_type(&RelationshipType::Footnotes) {
        let path = zip::resolve_target(doc_dir, &fn_rel.target);
        if let Some(data) = package.get_part(&path) {
            notes::parse_notes(data, "footnote", &ctx)?
        } else {
            HashMap::new()
        }
    } else {
        HashMap::new()
    };

    let endnotes = if let Some(en_rel) = doc_rels.find_by_type(&RelationshipType::Endnotes) {
        let path = zip::resolve_target(doc_dir, &en_rel.target);
        if let Some(data) = package.get_part(&path) {
            notes::parse_notes(data, "endnote", &ctx)?
        } else {
            HashMap::new()
        }
    } else {
        HashMap::new()
    };

    // Phase 5: Assemble
    Ok(Document {
        settings: doc_settings,
        theme,
        body: body_blocks,
        final_section,
        headers,
        footers,
        footnotes,
        endnotes,
        media,
    })
}
