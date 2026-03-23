//! Parser for `word/theme/theme1.xml` (DrawingML theme).

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Result;
use crate::model::{Theme, ThemeColorScheme};
use crate::xml;

pub fn parse_theme(data: &[u8]) -> Result<Theme> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut theme = Theme::default();
    let mut in_color_scheme = false;
    let mut in_major_font = false;
    let mut in_minor_font = false;
    let mut current_color_name: Option<String> = None;

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"clrScheme" => in_color_scheme = true,
                    b"majorFont" => in_major_font = true,
                    b"minorFont" => in_minor_font = true,
                    b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2" | b"accent3"
                    | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink"
                        if in_color_scheme =>
                    {
                        current_color_name = Some(String::from_utf8_lossy(&local).into_owned());
                    }
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"srgbClr" if current_color_name.is_some() => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            if let Some(rgb) = xml::parse_hex_color(&val) {
                                set_theme_color(
                                    &mut theme.color_scheme,
                                    current_color_name.as_deref().unwrap(),
                                    rgb,
                                );
                            }
                        }
                    }
                    b"sysClr" if current_color_name.is_some() => {
                        if let Some(val) = xml::optional_attr(e, b"lastClr")? {
                            if let Some(rgb) = xml::parse_hex_color(&val) {
                                set_theme_color(
                                    &mut theme.color_scheme,
                                    current_color_name.as_deref().unwrap(),
                                    rgb,
                                );
                            }
                        }
                    }
                    b"latin" if in_major_font || in_minor_font => {
                        if let Some(typeface) = xml::optional_attr(e, b"typeface")? {
                            let scheme = if in_major_font {
                                &mut theme.major_font
                            } else {
                                &mut theme.minor_font
                            };
                            scheme.latin = typeface;
                        }
                    }
                    b"ea" if in_major_font || in_minor_font => {
                        if let Some(typeface) = xml::optional_attr(e, b"typeface")? {
                            let scheme = if in_major_font {
                                &mut theme.major_font
                            } else {
                                &mut theme.minor_font
                            };
                            scheme.east_asian = typeface;
                        }
                    }
                    b"cs" if in_major_font || in_minor_font => {
                        if let Some(typeface) = xml::optional_attr(e, b"typeface")? {
                            let scheme = if in_major_font {
                                &mut theme.major_font
                            } else {
                                &mut theme.minor_font
                            };
                            scheme.complex_script = typeface;
                        }
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"clrScheme" => in_color_scheme = false,
                    b"majorFont" => in_major_font = false,
                    b"minorFont" => in_minor_font = false,
                    b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2" | b"accent3"
                    | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink" => {
                        current_color_name = None;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(theme)
}

fn set_theme_color(scheme: &mut ThemeColorScheme, name: &str, rgb: u32) {
    match name {
        "dk1" => scheme.dark1 = rgb,
        "lt1" => scheme.light1 = rgb,
        "dk2" => scheme.dark2 = rgb,
        "lt2" => scheme.light2 = rgb,
        "accent1" => scheme.accent1 = rgb,
        "accent2" => scheme.accent2 = rgb,
        "accent3" => scheme.accent3 = rgb,
        "accent4" => scheme.accent4 = rgb,
        "accent5" => scheme.accent5 = rgb,
        "accent6" => scheme.accent6 = rgb,
        "hlink" => scheme.hyperlink = rgb,
        "folHlink" => scheme.followed_hyperlink = rgb,
        _ => {}
    }
}
