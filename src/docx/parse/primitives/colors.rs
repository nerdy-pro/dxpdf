//! OOXML color-attribute primitives.
//!
//! - [`HexColor`] — ST_HexColor (§17.3.4.1): either the literal `"auto"`
//!   sentinel or a 6-digit RGB hex. Used by `<w:color>` and DrawingML color
//!   choices where "auto" is spec-legal.
//! - [`RgbHexU32`] — ST_HexColorRGB (§20.1.10.41): strictly a 6-digit RGB hex.
//!   Used where the spec disallows "auto".
//!
//! Both fail deserialization on malformed input (strict per plan §Decisions).

use serde::{Deserialize, Deserializer};

use crate::docx::model::Color;

/// OOXML `ST_HexColor` (§17.3.4.1): `"auto"` or 6-digit RGB hex.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HexColor {
    Auto,
    Rgb(u32),
}

impl HexColor {
    /// RGB value if concrete; `None` if `Auto`.
    pub fn rgb(self) -> Option<u32> {
        match self {
            HexColor::Auto => None,
            HexColor::Rgb(v) => Some(v),
        }
    }
}

impl From<HexColor> for Color {
    fn from(h: HexColor) -> Self {
        match h {
            HexColor::Auto => Self::Auto,
            HexColor::Rgb(v) => Self::Rgb(v),
        }
    }
}

impl<'de> Deserialize<'de> for HexColor {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        if s.eq_ignore_ascii_case("auto") {
            return Ok(HexColor::Auto);
        }
        u32::from_str_radix(&s, 16)
            .map(HexColor::Rgb)
            .map_err(serde::de::Error::custom)
    }
}

/// OOXML `ST_HexColorRGB` (§20.1.10.41): strictly a 6-digit RGB hex.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct RgbHexU32(pub u32);

impl<'de> Deserialize<'de> for RgbHexU32 {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        u32::from_str_radix(&s, 16)
            .map(RgbHexU32)
            .map_err(serde::de::Error::custom)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Deserialize)]
    struct HexVal {
        #[serde(rename = "@val")]
        val: HexColor,
    }

    #[derive(Deserialize)]
    struct RgbVal {
        #[serde(rename = "@val")]
        val: RgbHexU32,
    }

    #[test]
    fn hex_color_rgb() {
        let v: HexVal = quick_xml::de::from_str(r#"<x val="4F81BD"/>"#).unwrap();
        assert_eq!(v.val, HexColor::Rgb(0x4F81BD));
        assert_eq!(v.val.rgb(), Some(0x4F81BD));
    }

    #[test]
    fn hex_color_auto_is_case_insensitive() {
        for raw in ["auto", "AUTO", "Auto"] {
            let xml = format!(r#"<x val="{raw}"/>"#);
            let v: HexVal = quick_xml::de::from_str(&xml).unwrap();
            assert_eq!(v.val, HexColor::Auto);
            assert!(v.val.rgb().is_none());
        }
    }

    #[test]
    fn hex_color_rejects_garbage() {
        let r: Result<HexVal, _> = quick_xml::de::from_str(r#"<x val="notahex"/>"#);
        assert!(r.is_err());
    }

    #[test]
    fn rgb_hex_accepts_six_digit() {
        let v: RgbVal = quick_xml::de::from_str(r#"<x val="DEADBE"/>"#).unwrap();
        assert_eq!(v.val.0, 0xDEADBE);
    }

    #[test]
    fn rgb_hex_rejects_auto() {
        let r: Result<RgbVal, _> = quick_xml::de::from_str(r#"<x val="auto"/>"#);
        assert!(
            r.is_err(),
            "RgbHexU32 must reject 'auto' per ST_HexColorRGB"
        );
    }

    #[test]
    fn rgb_hex_rejects_garbage() {
        let r: Result<RgbVal, _> = quick_xml::de::from_str(r#"<x val="xyz123"/>"#);
        assert!(r.is_err());
    }
}
