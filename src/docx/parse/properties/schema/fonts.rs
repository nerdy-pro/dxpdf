//! `<w:rFonts>` (§17.3.2.26) — four font slots, each with explicit-name and
//! theme-reference attributes. Splits into a `FontSet` model.

use serde::Deserialize;

use crate::docx::model::{FontSet, FontSlot};
use crate::docx::parse::primitives::st_enums::StTheme;

#[derive(Clone, Debug, Deserialize)]
pub(crate) struct RFontsXml {
    #[serde(rename = "@ascii", default)]
    ascii: Option<String>,
    #[serde(rename = "@asciiTheme", default)]
    ascii_theme: Option<StTheme>,
    #[serde(rename = "@hAnsi", default)]
    h_ansi: Option<String>,
    #[serde(rename = "@hAnsiTheme", default)]
    h_ansi_theme: Option<StTheme>,
    #[serde(rename = "@eastAsia", default)]
    east_asia: Option<String>,
    #[serde(rename = "@eastAsiaTheme", default)]
    east_asia_theme: Option<StTheme>,
    #[serde(rename = "@cs", default)]
    cs: Option<String>,
    #[serde(rename = "@cstheme", alias = "@csTheme", default)]
    cs_theme: Option<StTheme>,
}

impl From<RFontsXml> for FontSet {
    fn from(x: RFontsXml) -> Self {
        Self {
            ascii: FontSlot {
                explicit: x.ascii,
                theme: x.ascii_theme.map(Into::into),
            },
            high_ansi: FontSlot {
                explicit: x.h_ansi,
                theme: x.h_ansi_theme.map(Into::into),
            },
            east_asian: FontSlot {
                explicit: x.east_asia,
                theme: x.east_asia_theme.map(Into::into),
            },
            complex_script: FontSlot {
                explicit: x.cs,
                theme: x.cs_theme.map(Into::into),
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::ThemeFontRef;

    fn parse(xml: &str) -> FontSet {
        let x: RFontsXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn explicit_names() {
        let f = parse(r#"<rFonts ascii="Calibri" hAnsi="Calibri" eastAsia="SimSun" cs="Arial"/>"#);
        assert_eq!(f.ascii.explicit.as_deref(), Some("Calibri"));
        assert_eq!(f.high_ansi.explicit.as_deref(), Some("Calibri"));
        assert_eq!(f.east_asian.explicit.as_deref(), Some("SimSun"));
        assert_eq!(f.complex_script.explicit.as_deref(), Some("Arial"));
    }

    #[test]
    fn theme_refs() {
        let f = parse(
            r#"<rFonts asciiTheme="minorHAnsi" hAnsiTheme="minorHAnsi" eastAsiaTheme="minorEastAsia"/>"#,
        );
        assert_eq!(f.ascii.theme, Some(ThemeFontRef::MinorHAnsi));
        assert_eq!(f.east_asian.theme, Some(ThemeFontRef::MinorEastAsia));
    }

    #[test]
    fn cs_theme_accepts_legacy_and_modern_case() {
        let f = parse(r#"<rFonts cstheme="minorBidi"/>"#);
        assert_eq!(f.complex_script.theme, Some(ThemeFontRef::MinorBidi));

        let f = parse(r#"<rFonts csTheme="minorBidi"/>"#);
        assert_eq!(f.complex_script.theme, Some(ThemeFontRef::MinorBidi));
    }

    #[test]
    fn missing_attrs_become_empty_slots() {
        let f = parse(r#"<rFonts/>"#);
        assert!(f.ascii.explicit.is_none() && f.ascii.theme.is_none());
    }
}
