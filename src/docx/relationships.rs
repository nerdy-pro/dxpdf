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
    /// §14.2.1 (chart part, §21.2): a DrawingML chart referenced from
    /// `<c:chart r:id>`.
    Chart,
    /// §14.2.5 diagram data — the `dgm:relIds/@r:dm` target, whose extension
    /// list names the pre-laid-out drawing part.
    DiagramData,
    /// [MS-ODRAWXML] §2.1.3: the `dsp:` drawing part holding a diagram's
    /// last successful layout (`.../2007/relationships/diagramDrawing`).
    DiagramDrawing,
    /// §14.2.6/§14.2.8/§14.2.4: the three sibling diagram parts this engine
    /// reads nothing from (layout definition, quick style, colors) — Word
    /// bakes their effect into the drawing part. Classified so they do not
    /// warn as unknown on every SmartArt document.
    DiagramSupport,
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
        } else if uri.ends_with("/chart") {
            Self::Chart
        } else if uri.ends_with("/diagramData") {
            Self::DiagramData
        } else if uri.ends_with("/diagramDrawing") {
            Self::DiagramDrawing
        } else if uri.ends_with("/diagramLayout")
            || uri.ends_with("/diagramQuickStyle")
            || uri.ends_with("/diagramColors")
        {
            Self::DiagramSupport
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

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn classify(suffix: &str) -> RelationshipType {
        RelationshipType::from_uri(&format!("{NS}/{suffix}"))
    }

    #[test]
    fn classifies_common_types() {
        assert_eq!(classify("officeDocument"), RelationshipType::OfficeDocument);
        assert_eq!(classify("styles"), RelationshipType::Styles);
        assert_eq!(classify("numbering"), RelationshipType::Numbering);
        assert_eq!(classify("settings"), RelationshipType::Settings);
        assert_eq!(classify("theme"), RelationshipType::Theme);
        assert_eq!(classify("header"), RelationshipType::Header);
        assert_eq!(classify("footer"), RelationshipType::Footer);
        assert_eq!(classify("footnotes"), RelationshipType::Footnotes);
        assert_eq!(classify("endnotes"), RelationshipType::Endnotes);
        assert_eq!(classify("image"), RelationshipType::Image);
        assert_eq!(classify("hyperlink"), RelationshipType::Hyperlink);
    }

    // These pairs share a suffix prefix, so the order of the `ends_with`
    // chain in `from_uri` matters. A reorder that let the shorter suffix win
    // would silently misclassify — these lock the distinction in.
    #[test]
    fn font_table_is_not_font() {
        assert_eq!(classify("fontTable"), RelationshipType::FontTable);
        assert_eq!(classify("font"), RelationshipType::Font);
    }

    #[test]
    fn styles_with_effects_is_not_styles() {
        assert_eq!(
            classify("stylesWithEffects"),
            RelationshipType::StylesWithEffects
        );
        assert_eq!(classify("styles"), RelationshipType::Styles);
    }

    #[test]
    fn property_relationship_types_are_distinct() {
        // core-properties uses a different namespace in practice, but the
        // suffix match is what disambiguates it from the -properties siblings.
        assert_eq!(
            RelationshipType::from_uri(
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
            ),
            RelationshipType::CoreProperties
        );
        assert_eq!(
            classify("extended-properties"),
            RelationshipType::ExtendedProperties
        );
        assert_eq!(
            classify("custom-properties"),
            RelationshipType::CustomProperties
        );
    }

    #[test]
    fn unknown_type_is_preserved_not_dropped() {
        let uri = "http://example.com/relationships/somethingNovel";
        assert_eq!(
            RelationshipType::from_uri(uri),
            RelationshipType::Unknown(uri.to_string())
        );
    }

    #[test]
    fn target_mode_defaults_to_internal_and_is_case_insensitive() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://x/image" Target="media/i.png"/>
            <Relationship Id="rId2" Type="http://x/hyperlink" Target="http://e.com" TargetMode="External"/>
            <Relationship Id="rId3" Type="http://x/hyperlink" Target="http://e.com" TargetMode="eXtErNaL"/>
        </Relationships>"#;
        let rels = Relationships::parse(xml).expect("parse rels");
        assert_eq!(
            rels.find_by_id("rId1").unwrap().target_mode,
            TargetMode::Internal
        );
        assert_eq!(
            rels.find_by_id("rId2").unwrap().target_mode,
            TargetMode::External
        );
        assert_eq!(
            rels.find_by_id("rId3").unwrap().target_mode,
            TargetMode::External
        );
    }
}
