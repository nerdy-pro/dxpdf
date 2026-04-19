//! §17.8: Font table parsing — extract embedded font declarations and de-obfuscate font data.
//!
//! Parses `word/fontTable.xml` to discover embedded font references,
//! then de-obfuscates `.odttf` files per §17.8.3.3.

use serde::Deserialize;

use crate::docx::error::{ParseError, Result};
use crate::docx::parse::serde_xml::from_xml;
use crate::docx::relationships::Relationships;
use crate::docx::zip::PackageContents;

use crate::model::{EmbeddedFont, EmbeddedFontVariant, RelId};

/// A raw embedded font reference before de-obfuscation.
struct FontEmbedRef {
    family: String,
    variant: EmbeddedFontVariant,
    rel_id: String,
    font_key: String,
}

/// Parse `fontTable.xml` and extract all embedded fonts.
///
/// Returns de-obfuscated font data ready for registration with a font manager.
///
/// Per §17.8.3.3, embedded fonts are obfuscated by XOR-ing the first 32 bytes
/// with the fontKey GUID (16 bytes, applied twice).
pub fn parse_embedded_fonts(
    font_table_data: &[u8],
    font_table_rels: &Relationships,
    package: &mut PackageContents,
    font_table_dir: &str,
) -> Result<Vec<EmbeddedFont>> {
    let refs = parse_font_table(font_table_data)?;

    let mut fonts = Vec::new();
    for embed_ref in refs {
        // Look up the relationship to find the font file path.
        let target_id = RelId::new(&embed_ref.rel_id);
        let rel = font_table_rels
            .rels
            .iter()
            .find(|r| r.id == target_id)
            .ok_or_else(|| {
                ParseError::MissingPart(format!(
                    "font relationship '{}' for '{}'",
                    embed_ref.rel_id, embed_ref.family
                ))
            })?;

        let font_path = crate::docx::zip::resolve_target(font_table_dir, &rel.target);
        let font_data = package.take_part(&font_path).ok_or_else(|| {
            ParseError::MissingPart(format!(
                "font file '{font_path}' for '{}'",
                embed_ref.family
            ))
        })?;

        let key_bytes = parse_font_key(&embed_ref.font_key)?;
        let data = deobfuscate_font(font_data, &key_bytes);

        fonts.push(EmbeddedFont {
            family: embed_ref.family,
            variant: embed_ref.variant,
            data,
        });
    }

    Ok(fonts)
}

/// Parse fontTable.xml and extract embedded font references.
fn parse_font_table(data: &[u8]) -> Result<Vec<FontEmbedRef>> {
    let ft: FontTableXml = from_xml(data)?;
    let mut refs = Vec::new();
    for font in ft.fonts {
        for (variant, embed) in [
            (EmbeddedFontVariant::Regular, font.embed_regular),
            (EmbeddedFontVariant::Bold, font.embed_bold),
            (EmbeddedFontVariant::Italic, font.embed_italic),
            (EmbeddedFontVariant::BoldItalic, font.embed_bold_italic),
        ] {
            if let Some(e) = embed {
                refs.push(FontEmbedRef {
                    family: font.name.clone(),
                    variant,
                    rel_id: e.id,
                    font_key: e.font_key,
                });
            }
        }
    }
    Ok(refs)
}

#[derive(Deserialize)]
struct FontTableXml {
    #[serde(rename = "font", default)]
    fonts: Vec<FontEntryXml>,
}

#[derive(Deserialize)]
struct FontEntryXml {
    #[serde(rename = "@name")]
    name: String,
    #[serde(rename = "embedRegular", default)]
    embed_regular: Option<EmbedXml>,
    #[serde(rename = "embedBold", default)]
    embed_bold: Option<EmbedXml>,
    #[serde(rename = "embedItalic", default)]
    embed_italic: Option<EmbedXml>,
    #[serde(rename = "embedBoldItalic", default)]
    embed_bold_italic: Option<EmbedXml>,
}

#[derive(Deserialize)]
struct EmbedXml {
    /// Attribute `r:id` / `@id` depending on namespace handling. Try both.
    #[serde(rename = "@id", alias = "@r:id")]
    id: String,
    /// Attribute `fontKey` — `w:` prefix stripped by quick-xml if present.
    #[serde(rename = "@fontKey", alias = "@w:fontKey")]
    font_key: String,
}

/// §17.8.3.3: Parse a fontKey GUID string into 16 bytes.
///
/// The GUID format is `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.
/// Byte order per RFC 4122: first 4 bytes LE, next 2 LE, next 2 LE, remaining 8 BE.
fn parse_font_key(key: &str) -> Result<[u8; 16]> {
    let hex: String = key.chars().filter(|c| c.is_ascii_hexdigit()).collect();
    if hex.len() != 32 {
        return Err(ParseError::InvalidAttributeValue {
            attr: "w:fontKey".into(),
            value: key.into(),
            reason: "expected 32 hex digits in GUID".into(),
        });
    }

    let raw: Vec<u8> = (0..16)
        .map(|i| u8::from_str_radix(&hex[i * 2..i * 2 + 2], 16))
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|e| ParseError::InvalidAttributeValue {
            attr: "w:fontKey".into(),
            value: key.into(),
            reason: format!("hex parse: {e}"),
        })?;

    // §17.8.3.3: the key bytes are the GUID hex digits reversed as a whole.
    let mut key = [0u8; 16];
    for (i, b) in raw.iter().rev().enumerate() {
        key[i] = *b;
    }
    Ok(key)
}

/// §17.8.3.3: De-obfuscate embedded font data.
///
/// XOR the first 32 bytes with the 16-byte key (applied twice).
fn deobfuscate_font(mut data: Vec<u8>, key: &[u8; 16]) -> Vec<u8> {
    for i in 0..32.min(data.len()) {
        data[i] ^= key[i % 16];
    }
    data
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_font_key_valid() {
        let key = parse_font_key("{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}").unwrap();
        // GUID hex: 3EEE3167E5B84798AE48EA6B71E31D4D
        // Reversed: 4D1DE3716BEA48AE9847B8E56731EE3E
        assert_eq!(key[0], 0x4D);
        assert_eq!(key[1], 0x1D);
        assert_eq!(key[2], 0xE3);
        assert_eq!(key[3], 0x71);
        assert_eq!(key[15], 0x3E);
    }

    #[test]
    fn parse_font_key_invalid_length() {
        let result = parse_font_key("{SHORT}");
        assert!(result.is_err());
    }

    #[test]
    fn deobfuscate_round_trip() {
        let key = [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
            0x0F, 0x10,
        ];
        let original = vec![0xAA; 64];
        let obfuscated = deobfuscate_font(original.clone(), &key);
        // First 32 bytes should be XOR'd
        assert_ne!(&obfuscated[..32], &original[..32]);
        // Bytes after 32 should be unchanged
        assert_eq!(&obfuscated[32..], &original[32..]);
        // XOR again to recover
        let recovered = deobfuscate_font(obfuscated, &key);
        assert_eq!(recovered, original);
    }

    #[test]
    fn parse_font_table_extracts_embeds() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Ubuntu">
    <w:panose1 w:val="020B0504030602030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:embedRegular r:id="rId1" w:fontKey="{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}"/>
    <w:embedBold r:id="rId2" w:fontKey="{773052E6-96AB-44E5-B2F7-4EE7F6CC51B1}"/>
  </w:font>
  <w:font w:name="Times New Roman">
    <w:charset w:val="00"/>
  </w:font>
</w:fonts>"#;
        let refs = parse_font_table(xml).unwrap();
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[0].family, "Ubuntu");
        assert_eq!(refs[0].variant, EmbeddedFontVariant::Regular);
        assert_eq!(refs[0].rel_id, "rId1");
        assert_eq!(refs[1].family, "Ubuntu");
        assert_eq!(refs[1].variant, EmbeddedFontVariant::Bold);
        assert_eq!(refs[1].rel_id, "rId2");
    }

    #[test]
    fn parse_font_table_no_embeds() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Arial">
    <w:charset w:val="00"/>
  </w:font>
</w:fonts>"#;
        let refs = parse_font_table(xml).unwrap();
        assert!(refs.is_empty());
    }
}
