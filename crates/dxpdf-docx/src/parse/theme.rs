//! Parser for `word/theme/theme1.xml` (DrawingML theme).
//!
//! Hierarchical parsing per §20.1.6.9 CT_OfficeStyleSheet.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::{Theme, ThemeColorScheme, ThemeFontScheme, ThemeScriptFont};
use crate::xml;

pub fn parse_theme(data: &[u8]) -> Result<Theme> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut theme = Theme::default();

    // Find <a:theme> root element.
    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"theme" => break,
            Event::Eof => return Ok(theme),
            _ => {}
        }
    }

    // Parse children of <a:theme>.
    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"themeElements" => {
                        parse_theme_elements(&mut reader, &mut buf, &mut theme)?;
                    }
                    // §20.1.6.7: object defaults — skip.
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
                match local {
                    b"objectDefaults" | b"extraClrSchemeLst" => {}
                    _ => xml::warn_unsupported_element("theme", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"theme" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"theme")),
            _ => {}
        }
    }

    Ok(theme)
}

// ── a:themeElements (§20.1.6.10) ────────────────────────────────────────────

fn parse_theme_elements(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    theme: &mut Theme,
) -> Result<()> {
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"clrScheme" => {
                        theme.color_scheme = parse_color_scheme(reader, buf)?;
                    }
                    b"fontScheme" => {
                        parse_font_scheme(reader, buf, theme)?;
                    }
                    // §20.1.4.1.14: format scheme (fills, lines, effects) — skip.
                    b"fmtScheme" => {
                        xml::skip_to_end(reader, buf, b"fmtScheme")?;
                    }
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

    Ok(())
}

// ── a:clrScheme (§20.1.6.2) ────────────────────────────────────────────────

fn parse_color_scheme(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ThemeColorScheme> {
    let mut scheme = ThemeColorScheme::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2" | b"accent3"
                    | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink" => {
                        if let Some(rgb) = parse_theme_color(reader, buf, local)? {
                            set_theme_color(&mut scheme, local, rgb);
                        }
                    }
                    _ => {
                        xml::warn_unsupported_element("clrScheme", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"clrScheme" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"clrScheme")),
            _ => {}
        }
    }

    Ok(scheme)
}

/// Parse a single theme color element (e.g., <a:dk1>).
/// Contains either <a:srgbClr val="..."/> or <a:sysClr lastClr="..."/>
/// possibly with color transform children (satMod, shade, tint, alpha, lumMod, etc.).
fn parse_theme_color(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<Option<u32>> {
    let mut rgb = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"srgbClr" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            rgb = xml::parse_hex_color(&val);
                        }
                        // Skip color transform children (satMod, tint, shade, alpha, etc.)
                        xml::skip_to_end(reader, buf, b"srgbClr")?;
                    }
                    b"sysClr" => {
                        if let Some(val) = xml::optional_attr(e, b"lastClr")? {
                            rgb = xml::parse_hex_color(&val);
                        }
                        xml::skip_to_end(reader, buf, b"sysClr")?;
                    }
                    _ => {
                        xml::warn_unsupported_element("themeColor", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"srgbClr" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            rgb = xml::parse_hex_color(&val);
                        }
                    }
                    b"sysClr" => {
                        if let Some(val) = xml::optional_attr(e, b"lastClr")? {
                            rgb = xml::parse_hex_color(&val);
                        }
                    }
                    _ => xml::warn_unsupported_element("themeColor", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }

    Ok(rgb)
}

fn set_theme_color(scheme: &mut ThemeColorScheme, name: &[u8], rgb: u32) {
    match name {
        b"dk1" => scheme.dark1 = rgb,
        b"lt1" => scheme.light1 = rgb,
        b"dk2" => scheme.dark2 = rgb,
        b"lt2" => scheme.light2 = rgb,
        b"accent1" => scheme.accent1 = rgb,
        b"accent2" => scheme.accent2 = rgb,
        b"accent3" => scheme.accent3 = rgb,
        b"accent4" => scheme.accent4 = rgb,
        b"accent5" => scheme.accent5 = rgb,
        b"accent6" => scheme.accent6 = rgb,
        b"hlink" => scheme.hyperlink = rgb,
        b"folHlink" => scheme.followed_hyperlink = rgb,
        _ => {}
    }
}

// ── a:fontScheme (§20.1.4.1.18) ────────────────────────────────────────────

fn parse_font_scheme(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    theme: &mut Theme,
) -> Result<()> {
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"majorFont" => {
                        theme.major_font = parse_font_collection(reader, buf, b"majorFont")?;
                    }
                    b"minorFont" => {
                        theme.minor_font = parse_font_collection(reader, buf, b"minorFont")?;
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

    Ok(())
}

/// §20.1.4.1.24 / §20.1.4.1.25: parse majorFont or minorFont.
fn parse_font_collection(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<ThemeFontScheme> {
    let mut scheme = ThemeFontScheme::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"latin" => {
                        scheme.latin = xml::optional_attr(e, b"typeface")?
                            .ok_or_else(|| crate::error::ParseError::MissingAttribute {
                                element: "a:latin".into(),
                                attr: "typeface".into(),
                            })?;
                    }
                    b"ea" => {
                        scheme.east_asian = xml::optional_attr(e, b"typeface")?
                            .ok_or_else(|| crate::error::ParseError::MissingAttribute {
                                element: "a:ea".into(),
                                attr: "typeface".into(),
                            })?;
                    }
                    b"cs" => {
                        scheme.complex_script = xml::optional_attr(e, b"typeface")?
                            .ok_or_else(|| crate::error::ParseError::MissingAttribute {
                                element: "a:cs".into(),
                                attr: "typeface".into(),
                            })?;
                    }
                    b"font" => {
                        let script = xml::optional_attr(e, b"script")?
                            .ok_or_else(|| crate::error::ParseError::MissingAttribute {
                                element: "a:font".into(),
                                attr: "script".into(),
                            })?;
                        let typeface = xml::optional_attr(e, b"typeface")?
                            .ok_or_else(|| crate::error::ParseError::MissingAttribute {
                                element: "a:font".into(),
                                attr: "typeface".into(),
                            })?;
                        scheme.script_fonts.push(ThemeScriptFont { script, typeface });
                    }
                    _ => xml::warn_unsupported_element("fontCollection", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }

    Ok(scheme)
}
