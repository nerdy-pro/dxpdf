//! Parser for `<a:clrScheme>` (§20.1.6.2) and the nested color choice element.
//!
//! Color transforms (`satMod`, `shade`, `tint`, `alpha`, `lumMod`, `lumOff`)
//! are accepted syntactically but discarded — the resolved model stores only
//! flat RGB per slot. Extending the model to preserve transforms is a
//! downstream concern (see `docs/` for follow-up notes).

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::error::Result;
use crate::docx::model::ThemeColorScheme;
use crate::docx::xml;

/// A slot in the theme color scheme (§20.1.4.1.10 `ST_SchemeColorVal`).
#[derive(Clone, Copy, Debug)]
enum ThemeColorSlot {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

impl ThemeColorSlot {
    fn from_local(name: &[u8]) -> Option<Self> {
        Some(match name {
            b"dk1" => Self::Dark1,
            b"lt1" => Self::Light1,
            b"dk2" => Self::Dark2,
            b"lt2" => Self::Light2,
            b"accent1" => Self::Accent1,
            b"accent2" => Self::Accent2,
            b"accent3" => Self::Accent3,
            b"accent4" => Self::Accent4,
            b"accent5" => Self::Accent5,
            b"accent6" => Self::Accent6,
            b"hlink" => Self::Hyperlink,
            b"folHlink" => Self::FollowedHyperlink,
            _ => return None,
        })
    }

    fn apply(self, scheme: &mut ThemeColorScheme, rgb: u32) {
        match self {
            Self::Dark1 => scheme.dark1 = rgb,
            Self::Light1 => scheme.light1 = rgb,
            Self::Dark2 => scheme.dark2 = rgb,
            Self::Light2 => scheme.light2 = rgb,
            Self::Accent1 => scheme.accent1 = rgb,
            Self::Accent2 => scheme.accent2 = rgb,
            Self::Accent3 => scheme.accent3 = rgb,
            Self::Accent4 => scheme.accent4 = rgb,
            Self::Accent5 => scheme.accent5 = rgb,
            Self::Accent6 => scheme.accent6 = rgb,
            Self::Hyperlink => scheme.hyperlink = rgb,
            Self::FollowedHyperlink => scheme.followed_hyperlink = rgb,
        }
    }
}

/// Parse `<a:clrScheme>…</a:clrScheme>`. Reader positioned after the Start.
pub fn parse_color_scheme(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ThemeColorScheme> {
    let mut scheme = ThemeColorScheme::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                if let Some(slot) = ThemeColorSlot::from_local(local) {
                    if let Some(rgb) = parse_color_choice(reader, buf)? {
                        slot.apply(&mut scheme, rgb);
                    }
                } else {
                    xml::warn_unsupported_element("clrScheme", local);
                    xml::skip_to_end(reader, buf, local)?;
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"clrScheme" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"clrScheme")),
            _ => {}
        }
    }

    Ok(scheme)
}

/// Read the single color-choice element nested inside a slot.
///
/// Accepted choices (§20.1.2.3): `srgbClr`, `sysClr`. Other choices
/// (`scRgbClr`, `hslClr`, `schemeClr`, `prstClr`) are logged and skipped.
/// Transform children are consumed but ignored.
///
/// Reader is positioned after the slot's Start event. Consumes events through
/// the matching End using depth tracking.
fn parse_color_choice(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<u32>> {
    let mut rgb: Option<u32> = None;
    let mut depth: u32 = 1;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                if depth == 1 {
                    capture_rgb_or_warn(e, &mut rgb)?;
                }
                depth += 1;
            }
            Event::Empty(ref e) if depth == 1 => {
                capture_rgb_or_warn(e, &mut rgb)?;
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(rgb);
                }
            }
            Event::Eof => return Err(xml::unexpected_eof(b"color-choice")),
            _ => {}
        }
    }
}

fn capture_rgb_or_warn(elem: &BytesStart<'_>, rgb: &mut Option<u32>) -> Result<()> {
    let qn = elem.name();
    let local = xml::local_name(qn.as_ref());
    match local {
        b"srgbClr" => {
            if rgb.is_none() {
                *rgb = xml::optional_attr(elem, b"val")?.and_then(|v| xml::parse_hex_color(&v));
            }
        }
        b"sysClr" => {
            if rgb.is_none() {
                *rgb = xml::optional_attr(elem, b"lastClr")?.and_then(|v| xml::parse_hex_color(&v));
            }
        }
        _ => xml::warn_unsupported_element("themeColor", local),
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml_src: &str) -> ThemeColorScheme {
        let mut reader = Reader::from_reader(xml_src.as_bytes());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        // Advance past the outer <a:clrScheme> Start event.
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"clrScheme" => break,
                Event::Eof => panic!("no clrScheme"),
                _ => {}
            }
        }
        parse_color_scheme(&mut reader, &mut buf).unwrap()
    }

    #[test]
    fn srgb_empty_elements() {
        let scheme = parse(
            r#"<a:clrScheme xmlns:a="urn:a">
                <a:dk1><a:srgbClr val="112233"/></a:dk1>
                <a:accent3><a:srgbClr val="AABBCC"/></a:accent3>
            </a:clrScheme>"#,
        );
        assert_eq!(scheme.dark1, 0x112233);
        assert_eq!(scheme.accent3, 0xAABBCC);
    }

    #[test]
    fn sys_clr_uses_last_clr() {
        let scheme = parse(
            r#"<a:clrScheme xmlns:a="urn:a">
                <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            </a:clrScheme>"#,
        );
        assert_eq!(scheme.light1, 0xFFFFFF);
    }

    #[test]
    fn transforms_are_tolerated() {
        let scheme = parse(
            r#"<a:clrScheme xmlns:a="urn:a">
                <a:accent1>
                    <a:srgbClr val="DEADBE">
                        <a:shade val="75000"/>
                        <a:satMod val="200000"/>
                    </a:srgbClr>
                </a:accent1>
            </a:clrScheme>"#,
        );
        assert_eq!(scheme.accent1, 0xDEADBE);
    }

    #[test]
    fn followed_hyperlink_slot() {
        let scheme = parse(
            r#"<a:clrScheme xmlns:a="urn:a">
                <a:folHlink><a:srgbClr val="010203"/></a:folHlink>
            </a:clrScheme>"#,
        );
        assert_eq!(scheme.followed_hyperlink, 0x010203);
    }

    #[test]
    fn missing_slot_keeps_default_zero() {
        let scheme = parse(r#"<a:clrScheme xmlns:a="urn:a"></a:clrScheme>"#);
        assert_eq!(scheme.dark1, 0);
    }
}
