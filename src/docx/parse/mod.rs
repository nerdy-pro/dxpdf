//! DOCX parsing — all submodules that parse specific parts of the OOXML package.

pub mod body;
pub mod body_schema;
pub mod drawing;
pub mod fonts;
pub mod notes;
pub mod numbering;
pub mod primitives;
pub mod properties;
pub mod rel_rewrite;
pub mod serde_xml;
pub mod settings;
pub mod styles;
pub mod theme;
pub mod vml;

use std::collections::HashMap;

use crate::docx::error::{ParseError, Result};
use crate::docx::model::*;
use crate::docx::relationships::{RelationshipType, Relationships};
use crate::docx::zip::{self, PackageContents};

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
            let mut blocks = body::parse_blocks(data)?;
            let remap = load_part_rel_remap(&path, &mut package, &mut media)?;
            rel_rewrite::rewrite_part_rels_in_blocks(&mut blocks, &remap);
            headers.insert(rel.id.clone(), blocks);
        }
    }

    for rel in doc_rels.filter_by_type(&RelationshipType::Footer) {
        let path = zip::resolve_target(doc_dir, &rel.target);
        if let Some(data) = package.get_part(&path) {
            let mut blocks = body::parse_blocks(data)?;
            let remap = load_part_rel_remap(&path, &mut package, &mut media)?;
            rel_rewrite::rewrite_part_rels_in_blocks(&mut blocks, &remap);
            footers.insert(rel.id.clone(), blocks);
        }
    }

    // Phase 4b: Parse footnotes and endnotes. All footnote bodies
    // share the same `footnotes.xml.rels` namespace (one part holds
    // all of them); same for endnotes.
    let mut footnotes = HashMap::new();
    if let Some(fn_rel) = doc_rels.find_by_type(&RelationshipType::Footnotes) {
        let path = zip::resolve_target(doc_dir, &fn_rel.target);
        if let Some(data) = package.get_part(&path) {
            footnotes = notes::parse_notes(data, "footnote")?;
            let remap = load_part_rel_remap(&path, &mut package, &mut media)?;
            for blocks in footnotes.values_mut() {
                rel_rewrite::rewrite_part_rels_in_blocks(blocks, &remap);
            }
        }
    }

    let mut endnotes = HashMap::new();
    if let Some(en_rel) = doc_rels.find_by_type(&RelationshipType::Endnotes) {
        let path = zip::resolve_target(doc_dir, &en_rel.target);
        if let Some(data) = package.get_part(&path) {
            endnotes = notes::parse_notes(data, "endnote")?;
            let remap = load_part_rel_remap(&path, &mut package, &mut media)?;
            for blocks in endnotes.values_mut() {
                rel_rewrite::rewrite_part_rels_in_blocks(blocks, &remap);
            }
        }
    }

    // Phase 5: Resolve body hyperlink RelIds to actual URLs.
    // Headers/footers/footnotes/endnotes already had both their image
    // and hyperlink rIds rewritten through `load_part_rel_remap` —
    // that pass is the per-part equivalent of this body-level resolve.
    resolve_hyperlinks(&mut body_blocks, &doc_rels);

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

/// Process the rels of a single subordinate XML part — a header,
/// footer, footnote, or endnote. §17.16 gives every part its own
/// `_rels/<part>.xml.rels` file with its own `rId` namespace, so an
/// `rId1` in a header is unrelated to an `rId1` in the document, in
/// a footer, or in another header.
///
/// To make the parsed `Document` collision-free we synthesize **two**
/// kinds of remap entries here and return them as one combined map
/// the caller applies via `rel_rewrite::rewrite_part_rels_in_blocks`:
///
/// * **Image rels** — each loaded image is inserted into the
///   document-wide `media` map under a synthesized unique key
///   (`<part_path>::<orig_rId>`); the part's blocks are rewritten to
///   reference that key.
/// * **Hyperlink rels** — `External(rId)` targets are rewritten to
///   `External(URL)` using *this part's* relationships, not the
///   document's. Pre-fix the hyperlink resolution pass used
///   `doc_rels` for every part, which silently re-targeted any
///   header/footer/note hyperlink whose rId happened to mean
///   something different in the document's own rels.
///
/// Other relationship types (font tables, OLE objects, custom XML, …)
/// are not modeled in the block tree yet and are skipped.
fn load_part_rel_remap(
    part_path: &str,
    package: &mut PackageContents,
    media: &mut HashMap<RelId, (Vec<u8>, ImageFormat)>,
) -> Result<HashMap<RelId, RelId>> {
    let mut remap: HashMap<RelId, RelId> = HashMap::new();
    let rels_path = zip::rels_path_for(part_path);
    let Some(rels_data) = package.get_part(&rels_path) else {
        return Ok(remap);
    };
    let rels = Relationships::parse(rels_data)?;

    // Image rels — synthesize unique ids and load bytes into `media`.
    for img_rel in rels.filter_by_type(&RelationshipType::Image) {
        let img_path = zip::resolve_target(zip::part_directory(part_path), &img_rel.target);
        // Clone, don't take: multiple parts (e.g. a header and a
        // footer) commonly reference the *same* image part; if we
        // moved the bytes out at the first reader the second reader
        // would silently drop the image.
        if let Some(img_data) = package.get_part(&img_path).map(<[u8]>::to_vec) {
            let fmt = ImageFormat::detect(&img_rel.target, &img_data);
            let unique_id = RelId::new(format!("{}::{}", part_path, img_rel.id.as_str()));
            media.insert(unique_id.clone(), (img_data, fmt));
            remap.insert(img_rel.id.clone(), unique_id);
        }
    }

    // Hyperlink rels — same mechanism the document-level hyperlink
    // resolver uses: rewrite `External(rId)` to `External(URL)` by
    // smuggling the URL through a `RelId` newtype. The model layer
    // already treats a URL-shaped RelId as an external link target.
    for link_rel in rels.filter_by_type(&RelationshipType::Hyperlink) {
        remap.insert(link_rel.id.clone(), RelId::new(link_rel.target.clone()));
    }

    Ok(remap)
}

/// Walk blocks and resolve `HyperlinkTarget::External(RelId)` to actual URLs.
fn resolve_hyperlinks(blocks: &mut [Block], rels: &crate::docx::relationships::Relationships) {
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
    rels: &crate::docx::relationships::Relationships,
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
