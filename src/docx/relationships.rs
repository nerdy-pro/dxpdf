//! OOXML package relationships (`.rels` files).

use log::warn;
use serde::{Deserialize, Deserializer};

use crate::docx::error::Result;
use crate::docx::model::RelId;
use crate::docx::parse::serde_xml::from_xml;

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
        } else if uri.ends_with("/core-properties") {
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

impl<'de> Deserialize<'de> for TargetMode {
    fn deserialize<D: Deserializer<'de>>(d: D) -> std::result::Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        Ok(if s.eq_ignore_ascii_case("external") {
            TargetMode::External
        } else {
            TargetMode::Internal
        })
    }
}

/// A collection of relationships from a single .rels file.
#[derive(Clone, Debug, Default)]
pub struct Relationships {
    pub(crate) rels: Vec<Relationship>,
}

impl Relationships {
    /// Parse a .rels XML file.
    pub fn parse(data: &[u8]) -> Result<Self> {
        from_xml::<RelationshipsXml>(data).map(Into::into)
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

#[derive(Deserialize, Default)]
struct RelationshipsXml {
    #[serde(rename = "Relationship", default)]
    rels: Vec<RelationshipXml>,
}

#[derive(Deserialize)]
struct RelationshipXml {
    #[serde(rename = "@Id")]
    id: String,
    #[serde(rename = "@Type")]
    rel_type: String,
    #[serde(rename = "@Target")]
    target: String,
    #[serde(rename = "@TargetMode", default)]
    target_mode: TargetMode,
}

impl From<RelationshipsXml> for Relationships {
    fn from(x: RelationshipsXml) -> Self {
        Self {
            rels: x.rels.into_iter().map(Relationship::from).collect(),
        }
    }
}

impl From<RelationshipXml> for Relationship {
    fn from(r: RelationshipXml) -> Self {
        Self {
            id: RelId::new(r.id),
            rel_type: RelationshipType::from_uri(&r.rel_type),
            target: r.target,
            target_mode: r.target_mode,
        }
    }
}
