//! Shading sub-schema (§17.3.5 w:shd).

use serde::Deserialize;

use crate::docx::model::{Color, Shading};
use crate::docx::parse::primitives::st_enums::StShd;
use crate::docx::parse::primitives::HexColor;

/// `<w:shd w:val="..." w:fill="..." w:color="..."/>`.
///
/// `@fill` is the background color; `@color` the pattern foreground (for
/// striped/crossed patterns). Both accept `auto`. Defaults: fill=auto,
/// color=auto, pattern=clear.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct ShdXml {
    #[serde(rename = "@val", default = "default_pattern")]
    val: StShd,
    #[serde(rename = "@fill", default)]
    fill: Option<HexColor>,
    #[serde(rename = "@color", default)]
    color: Option<HexColor>,
}

fn default_pattern() -> StShd {
    StShd::Clear
}

impl From<ShdXml> for Shading {
    fn from(x: ShdXml) -> Self {
        Self {
            fill: x.fill.map_or(Color::Auto, Into::into),
            pattern: x.val.into(),
            color: x.color.map_or(Color::Auto, Into::into),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::ShadingPattern;

    fn parse(xml: &str) -> Shading {
        let x: ShdXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn full_solid() {
        let s = parse(r#"<shd val="solid" fill="FF0000" color="000000"/>"#);
        assert_eq!(s.pattern, ShadingPattern::Solid);
        assert_eq!(s.fill, Color::Rgb(0xFF0000));
        assert_eq!(s.color, Color::Rgb(0));
    }

    #[test]
    fn fill_auto_preserved() {
        let s = parse(r#"<shd val="clear" fill="auto" color="auto"/>"#);
        assert_eq!(s.pattern, ShadingPattern::Clear);
        assert_eq!(s.fill, Color::Auto);
        assert_eq!(s.color, Color::Auto);
    }

    #[test]
    fn missing_colors_default_to_auto() {
        let s = parse(r#"<shd val="pct25"/>"#);
        assert_eq!(s.pattern, ShadingPattern::Pct25);
        assert_eq!(s.fill, Color::Auto);
        assert_eq!(s.color, Color::Auto);
    }
}
