//! Parser for `<a:fontScheme>` (§20.1.4.1.18) and its two font collections.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::error::Result;
use crate::docx::model::{ThemeFontScheme, ThemeScriptFont};
use crate::docx::parse::theme::script::parse_script_tag;
use crate::docx::xml;

/// Both collections extracted from a `<a:fontScheme>`.
#[derive(Debug, Default)]
pub struct FontSchemes {
    pub major: ThemeFontScheme,
    pub minor: ThemeFontScheme,
}

/// A child of `<a:majorFont>` / `<a:minorFont>` (§20.1.4.1.24 / §20.1.4.1.25).
#[derive(Clone, Copy, Debug)]
enum FontCollectionChild {
    Latin,
    EastAsian,
    ComplexScript,
    Font,
}

impl FontCollectionChild {
    fn from_local(name: &[u8]) -> Option<Self> {
        Some(match name {
            b"latin" => Self::Latin,
            b"ea" => Self::EastAsian,
            b"cs" => Self::ComplexScript,
            b"font" => Self::Font,
            _ => return None,
        })
    }
}

/// Parse `<a:fontScheme>…</a:fontScheme>`. Reader positioned after the Start.
pub fn parse_font_scheme(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<FontSchemes> {
    let mut schemes = FontSchemes::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"majorFont" => {
                        schemes.major = parse_font_collection(reader, buf, b"majorFont")?;
                    }
                    b"minorFont" => {
                        schemes.minor = parse_font_collection(reader, buf, b"minorFont")?;
                    }
                    _ => {
                        xml::warn_unsupported_element("fontScheme", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"fontScheme" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"fontScheme")),
            _ => {}
        }
    }

    Ok(schemes)
}

/// Parse a `<a:majorFont>` or `<a:minorFont>` collection. Reader positioned
/// after the Start event; `end_tag` names that same element.
fn parse_font_collection(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<ThemeFontScheme> {
    let mut scheme = ThemeFontScheme::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) => apply_child(e, &mut scheme)?,
            Event::Start(ref e) => {
                apply_child(e, &mut scheme)?;
                // Children of latin/ea/cs/font are not expected; skip any present.
                let local = xml::local_name_owned(e.name().as_ref());
                xml::skip_to_end(reader, buf, &local)?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }

    Ok(scheme)
}

fn apply_child(elem: &BytesStart<'_>, scheme: &mut ThemeFontScheme) -> Result<()> {
    let qn = elem.name();
    let local = xml::local_name(qn.as_ref());
    match FontCollectionChild::from_local(local) {
        Some(FontCollectionChild::Latin) => {
            scheme.latin = xml::required_attr(elem, b"typeface")?;
        }
        Some(FontCollectionChild::EastAsian) => {
            scheme.east_asian = xml::required_attr(elem, b"typeface")?;
        }
        Some(FontCollectionChild::ComplexScript) => {
            scheme.complex_script = xml::required_attr(elem, b"typeface")?;
        }
        Some(FontCollectionChild::Font) => {
            let script = parse_script_tag(&xml::required_attr(elem, b"script")?);
            let typeface = xml::required_attr(elem, b"typeface")?;
            scheme
                .script_fonts
                .push(ThemeScriptFont { script, typeface });
        }
        None => xml::warn_unsupported_element("fontCollection", local),
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::ScriptTag;

    fn parse_scheme(xml_src: &str) -> FontSchemes {
        let mut reader = Reader::from_reader(xml_src.as_bytes());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"fontScheme" => break,
                Event::Eof => panic!("no fontScheme"),
                _ => {}
            }
        }
        parse_font_scheme(&mut reader, &mut buf).unwrap()
    }

    #[test]
    fn major_and_minor_latin_are_captured() {
        let s = parse_scheme(
            r#"<a:fontScheme xmlns:a="urn:a">
                <a:majorFont>
                    <a:latin typeface="Calibri Light"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                </a:majorFont>
                <a:minorFont>
                    <a:latin typeface="Calibri"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                </a:minorFont>
            </a:fontScheme>"#,
        );
        assert_eq!(s.major.latin, "Calibri Light");
        assert_eq!(s.minor.latin, "Calibri");
    }

    #[test]
    fn per_script_font_entries_are_captured() {
        let s = parse_scheme(
            r#"<a:fontScheme xmlns:a="urn:a">
                <a:majorFont>
                    <a:latin typeface="Calibri Light"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                    <a:font script="Hans" typeface="Noto Sans SC"/>
                    <a:font script="Xxxx" typeface="FallbackFace"/>
                </a:majorFont>
                <a:minorFont>
                    <a:latin typeface="Calibri"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                </a:minorFont>
            </a:fontScheme>"#,
        );
        assert_eq!(s.major.script_fonts.len(), 2);
        assert_eq!(s.major.script_fonts[0].script, ScriptTag::Hans);
        assert_eq!(s.major.script_fonts[0].typeface, "Noto Sans SC");
        match &s.major.script_fonts[1].script {
            ScriptTag::Other(code) => assert_eq!(&**code, "Xxxx"),
            other => panic!("expected Other, got {other:?}"),
        }
    }

    #[test]
    fn east_asian_and_complex_script_fields() {
        let s = parse_scheme(
            r#"<a:fontScheme xmlns:a="urn:a">
                <a:majorFont>
                    <a:latin typeface="A"/>
                    <a:ea typeface="B"/>
                    <a:cs typeface="C"/>
                </a:majorFont>
                <a:minorFont>
                    <a:latin typeface=""/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                </a:minorFont>
            </a:fontScheme>"#,
        );
        assert_eq!(s.major.latin, "A");
        assert_eq!(s.major.east_asian, "B");
        assert_eq!(s.major.complex_script, "C");
    }
}
