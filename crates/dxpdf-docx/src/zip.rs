//! ZIP extraction and OOXML relationship (.rels) parsing.

use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::Reader;

use log::warn;

use crate::error::{ParseError, Result};
use crate::model::RelId;
use crate::xml;

/// The contents of a DOCX package, extracted from the ZIP archive.
pub struct PackageContents {
    /// All files in the ZIP, keyed by normalized path (no leading slash).
    pub parts: HashMap<String, Vec<u8>>,
}

impl PackageContents {
    /// Extract all parts from a DOCX ZIP archive.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor)?;
        let mut parts = HashMap::with_capacity(archive.len());

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = normalize_path(file.name());
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)?;
            parts.insert(name, buf);
        }

        Ok(Self { parts })
    }

    /// Get the bytes for a part, case-insensitively.
    pub fn get_part(&self, path: &str) -> Option<&[u8]> {
        let normalized = normalize_path(path);
        self.parts.get(&normalized).map(|v| v.as_slice())
    }

    /// Get part bytes, or return MissingPart error.
    pub fn require_part(&self, path: &str) -> Result<&[u8]> {
        self.get_part(path)
            .ok_or_else(|| ParseError::MissingPart(path.to_string()))
    }

    /// Remove and return the owned bytes for a part. Avoids cloning.
    pub fn take_part(&mut self, path: &str) -> Option<Vec<u8>> {
        let normalized = normalize_path(path);
        self.parts.remove(&normalized)
    }
}

fn normalize_path(path: &str) -> String {
    path.trim_start_matches('/').to_lowercase()
}

// ── Relationships ────────────────────────────────────────────────────────────

/// A parsed relationship from a .rels file.
#[derive(Clone, Debug)]
pub struct Relationship {
    pub id: RelId,
    pub rel_type: RelationshipType,
    pub target: String,
    pub target_mode: TargetMode,
}

/// Known OOXML relationship types (§7.1, §11.3, §15.2).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RelationshipType {
    /// §11.3.10: main document part.
    OfficeDocument,
    /// §11.3.11: style definitions.
    Styles,
    /// §11.3.7: numbering definitions.
    Numbering,
    /// §11.3.9: document settings.
    Settings,
    /// §11.3.5: font table.
    FontTable,
    /// §14.2.7: theme.
    Theme,
    /// §11.3.6: header part.
    Header,
    /// §11.3.4: footer part.
    Footer,
    /// §11.3.3: footnotes part.
    Footnotes,
    /// §11.3.2: endnotes part.
    Endnotes,
    /// §15.2.13: font.
    Font,
    /// §15.2.14: image.
    Image,
    /// §15.3.6: hyperlink.
    Hyperlink,
    /// §11.3.1: comments part.
    Comments,
    /// §15.2.12.1: core (Dublin Core) properties.
    CoreProperties,
    /// §15.2.12.3: extended (application) properties.
    ExtendedProperties,
    /// §15.2.12.2: custom properties.
    CustomProperties,
    /// §15.2.1.1: custom XML data.
    CustomXml,
    /// §11.3.12: web settings.
    WebSettings,
    /// MS Office extension: styles with effects (Office 2007+).
    StylesWithEffects,
    /// §11.3.8: glossary/building blocks document.
    GlossaryDocument,
    /// Any relationship type not listed above.
    Unknown(String),
}

impl RelationshipType {
    fn from_uri(uri: &str) -> Self {
        // OOXML uses long URIs; match on the final segment.
        if uri.ends_with("/officeDocument") || uri.ends_with("/document") {
            Self::OfficeDocument
        } else if uri.ends_with("/styles") {
            Self::Styles
        } else if uri.ends_with("/numbering") {
            Self::Numbering
        } else if uri.ends_with("/settings") {
            Self::Settings
        } else if uri.ends_with("/fontTable") {
            Self::FontTable
        } else if uri.ends_with("/theme") {
            Self::Theme
        } else if uri.ends_with("/header") {
            Self::Header
        } else if uri.ends_with("/footer") {
            Self::Footer
        } else if uri.ends_with("/footnotes") {
            Self::Footnotes
        } else if uri.ends_with("/endnotes") {
            Self::Endnotes
        } else if uri.ends_with("/font") {
            Self::Font
        } else if uri.ends_with("/image") {
            Self::Image
        } else if uri.ends_with("/hyperlink") {
            Self::Hyperlink
        } else if uri.ends_with("/comments") {
            Self::Comments
        } else if uri.ends_with("/core-properties") || uri.ends_with("/metadata/core-properties")
        {
            Self::CoreProperties
        } else if uri.ends_with("/extended-properties") {
            Self::ExtendedProperties
        } else if uri.ends_with("/custom-properties") {
            Self::CustomProperties
        } else if uri.ends_with("/customXml") {
            Self::CustomXml
        } else if uri.ends_with("/webSettings") {
            Self::WebSettings
        } else if uri.ends_with("/stylesWithEffects") {
            Self::StylesWithEffects
        } else if uri.ends_with("/glossaryDocument") {
            Self::GlossaryDocument
        } else {
            warn!("unknown relationship type: {}", uri);
            Self::Unknown(uri.to_string())
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TargetMode {
    #[default]
    Internal,
    External,
}

/// A collection of relationships from a single .rels file.
#[derive(Clone, Debug, Default)]
pub struct Relationships {
    pub(crate) rels: Vec<Relationship>,
}

impl Relationships {
    /// Parse a .rels XML file.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut rels = Vec::new();

        loop {
            match xml::next_event(&mut reader, &mut buf)? {
                Event::Empty(ref e) | Event::Start(ref e)
                    if xml::local_name(e.name().as_ref()) == b"Relationship" =>
                {
                    let id = RelId::new(xml::required_attr(e, b"Id")?);
                    let rel_type_uri = xml::required_attr(e, b"Type")?;
                    let target = xml::required_attr(e, b"Target")?;
                    let target_mode = match xml::optional_attr(e, b"TargetMode")? {
                        Some(ref s) if s.eq_ignore_ascii_case("external") => TargetMode::External,
                        _ => TargetMode::Internal,
                    };

                    rels.push(Relationship {
                        id,
                        rel_type: RelationshipType::from_uri(&rel_type_uri),
                        target,
                        target_mode,
                    });
                }
                Event::Eof => break,
                _ => {}
            }
        }

        Ok(Self { rels })
    }

    /// Find the first relationship of a given type.
    pub fn find_by_type(&self, rel_type: &RelationshipType) -> Option<&Relationship> {
        self.rels.iter().find(|r| &r.rel_type == rel_type)
    }

    /// Find all relationships of a given type.
    pub fn filter_by_type(&self, rel_type: &RelationshipType) -> Vec<&Relationship> {
        self.rels
            .iter()
            .filter(|r| &r.rel_type == rel_type)
            .collect()
    }

    /// Look up a relationship by its ID.
    pub fn find_by_id(&self, id: &str) -> Option<&Relationship> {
        self.rels.iter().find(|r| r.id.as_str() == id)
    }

    /// Get all relationships.
    pub fn all(&self) -> &[Relationship] {
        &self.rels
    }
}

/// Resolve a relationship target to an absolute path within the package.
/// base_dir is the directory containing the source part (e.g., "word" for "word/document.xml").
pub fn resolve_target(base_dir: &str, target: &str) -> String {
    if target.starts_with('/') {
        // Absolute path within the package
        normalize_path(target)
    } else {
        // Relative to base_dir
        let mut path = if base_dir.is_empty() {
            target.to_string()
        } else {
            format!("{}/{}", base_dir, target)
        };
        // Simplify "../" sequences
        while let Some(pos) = path.find("/../") {
            if let Some(parent_start) = path[..pos].rfind('/') {
                path = format!("{}{}", &path[..parent_start], &path[pos + 3..]);
            } else {
                path = path[pos + 4..].to_string();
            }
        }
        normalize_path(&path)
    }
}

/// Get the .rels path for a given part path.
/// e.g., "word/document.xml" → "word/_rels/document.xml.rels"
pub fn rels_path_for(part_path: &str) -> String {
    let normalized = normalize_path(part_path);
    if let Some(slash_pos) = normalized.rfind('/') {
        format!(
            "{}/_rels/{}.rels",
            &normalized[..slash_pos],
            &normalized[slash_pos + 1..]
        )
    } else {
        format!("_rels/{}.rels", normalized)
    }
}

/// Get the directory portion of a part path.
pub fn part_directory(part_path: &str) -> &str {
    match part_path.rfind('/') {
        Some(pos) => &part_path[..pos],
        None => "",
    }
}
