//! Parser for `word/theme/theme1.xml` (DrawingML theme).
//!
//! Hierarchical parsing per §20.1.6.9 `CT_OfficeStyleSheet`. Each sub-parser
//! is a pure function over the XML reader; the root composes the results into
//! a `Theme`.

mod color;
mod font;
mod script;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::docx::error::Result;
use crate::docx::model::{Theme, ThemeColorScheme};
use crate::docx::xml;

use self::font::FontSchemes;

pub fn parse_theme(data: &[u8]) -> Result<Theme> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut theme = Theme::default();

    // Find <a:theme> root element; absent root → default theme.
    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"theme" => break,
            Event::Eof => return Ok(theme),
            _ => {}
        }
    }

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"themeElements" => {
                        let elements = parse_theme_elements(&mut reader, &mut buf)?;
                        theme.color_scheme = elements.color_scheme;
                        theme.major_font = elements.fonts.major;
                        theme.minor_font = elements.fonts.minor;
                    }
                    // §20.1.6.7 / §20.1.6.8: siblings irrelevant for rendering — skip.
                    b"objectDefaults" | b"extraClrSchemeLst" | b"extLst" => {
                        xml::skip_to_end(&mut reader, &mut buf, local)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("theme", local);
                        xml::skip_to_end(&mut reader, &mut buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                if !matches!(local, b"objectDefaults" | b"extraClrSchemeLst") {
                    xml::warn_unsupported_element("theme", local);
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"theme" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"theme")),
            _ => {}
        }
    }

    Ok(theme)
}

struct ThemeElements {
    color_scheme: ThemeColorScheme,
    fonts: FontSchemes,
}

fn parse_theme_elements(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<ThemeElements> {
    let mut color_scheme = ThemeColorScheme::default();
    let mut fonts = FontSchemes::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"clrScheme" => color_scheme = color::parse_color_scheme(reader, buf)?,
                    b"fontScheme" => fonts = font::parse_font_scheme(reader, buf)?,
                    // §20.1.4.1.14: format scheme (fills, lines, effects) — skip.
                    b"fmtScheme" => xml::skip_to_end(reader, buf, b"fmtScheme")?,
                    _ => {
                        xml::warn_unsupported_element("themeElements", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"themeElements" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"themeElements")),
            _ => {}
        }
    }

    Ok(ThemeElements {
        color_scheme,
        fonts,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::ScriptTag;

    const MINIMAL_THEME: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="urn:a" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Hans" typeface="Noto Sans SC"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst/>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>"#;

    #[test]
    fn end_to_end_theme_parse() {
        let theme = parse_theme(MINIMAL_THEME.as_bytes()).unwrap();

        assert_eq!(theme.color_scheme.dark1, 0x000000);
        assert_eq!(theme.color_scheme.light1, 0xFFFFFF);
        assert_eq!(theme.color_scheme.accent1, 0x4F81BD);
        assert_eq!(theme.color_scheme.hyperlink, 0x0000FF);
        assert_eq!(theme.color_scheme.followed_hyperlink, 0x800080);

        assert_eq!(theme.major_font.latin, "Calibri Light");
        assert_eq!(theme.minor_font.latin, "Calibri");
        assert_eq!(theme.major_font.script_fonts.len(), 1);
        assert_eq!(theme.major_font.script_fonts[0].script, ScriptTag::Hans);
        assert_eq!(theme.major_font.script_fonts[0].typeface, "Noto Sans SC");
    }

    #[test]
    fn empty_input_returns_default_theme() {
        let theme = parse_theme(b"").unwrap();
        assert_eq!(theme.color_scheme.dark1, 0);
        assert!(theme.major_font.latin.is_empty());
    }

    #[test]
    fn missing_theme_elements_keeps_defaults() {
        let xml = r#"<a:theme xmlns:a="urn:a"><a:objectDefaults/></a:theme>"#;
        let theme = parse_theme(xml.as_bytes()).unwrap();
        assert_eq!(theme.color_scheme.dark1, 0);
        assert!(theme.minor_font.latin.is_empty());
    }
}
