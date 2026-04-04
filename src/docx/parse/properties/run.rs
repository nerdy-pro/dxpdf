use quick_xml::events::Event;
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

use super::{invalid_value, opt_val, parse_border, parse_color_attr, parse_shading, toggle_attr};

/// Parse `w:rPr` element. Returns (properties, optional style ID).
pub fn parse_run_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(RunProperties, Option<StyleId>)> {
    let mut props = RunProperties::default();
    let mut style_id: Option<StyleId> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"rStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId::new);
                    }
                    b"rFonts" => {
                        props.fonts = parse_font_set(e)?;
                    }
                    b"sz" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.font_size = Some(Dimension::new(val));
                        }
                    }
                    b"szCs" => {}
                    b"b" => {
                        props.bold = toggle_attr(e)?;
                    }
                    b"bCs" => {}
                    b"i" => {
                        props.italic = toggle_attr(e)?;
                    }
                    b"iCs" => {}
                    b"u" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.underline = Some(parse_underline_style(&val)?);
                        } else {
                            props.underline = Some(UnderlineStyle::Single);
                        }
                    }
                    b"strike" => {
                        let on = xml::optional_attr_bool(e, b"val")?.unwrap_or(true);
                        if on {
                            props.strike = Some(StrikeStyle::Single);
                        }
                    }
                    b"dstrike" => {
                        let on = xml::optional_attr_bool(e, b"val")?.unwrap_or(true);
                        if on {
                            props.strike = Some(StrikeStyle::Double);
                        }
                    }
                    b"color" => {
                        props.color = Some(parse_color_attr(e)?);
                    }
                    b"highlight" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.highlight = parse_highlight_color(&val)?;
                        }
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"vertAlign" => {
                        props.vertical_align = opt_val(e, parse_vertical_align)?;
                    }
                    b"spacing" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.spacing = Some(Dimension::new(val));
                        }
                    }
                    b"kern" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.kerning = Some(Dimension::new(val));
                        }
                    }
                    b"caps" => {
                        props.all_caps = toggle_attr(e)?;
                    }
                    b"smallCaps" => {
                        props.small_caps = toggle_attr(e)?;
                    }
                    b"vanish" => {
                        props.vanish = toggle_attr(e)?;
                    }
                    b"noProof" => {
                        props.no_proof = toggle_attr(e)?;
                    }
                    b"webHidden" => {
                        props.web_hidden = toggle_attr(e)?;
                    }
                    b"position" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.position = Some(Dimension::new(val));
                        }
                    }
                    b"lang" => {
                        props.lang = Some(Lang {
                            val: xml::optional_attr(e, b"val")?,
                            east_asia: xml::optional_attr(e, b"eastAsia")?,
                            bidi: xml::optional_attr(e, b"bidi")?,
                        });
                    }
                    b"rtl" => {
                        props.rtl = toggle_attr(e)?;
                    }
                    b"emboss" => {
                        props.emboss = toggle_attr(e)?;
                    }
                    b"imprint" => {
                        props.imprint = toggle_attr(e)?;
                    }
                    b"outline" => {
                        props.outline = toggle_attr(e)?;
                    }
                    b"shadow" => {
                        props.shadow = toggle_attr(e)?;
                    }
                    b"bdr" => {
                        props.border = Some(parse_border(e)?);
                    }
                    _ => xml::warn_unsupported_element("rPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"rPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"rPr")),
            _ => {}
        }
    }

    Ok((props, style_id))
}

fn parse_font_set(e: &quick_xml::events::BytesStart<'_>) -> Result<FontSet> {
    Ok(FontSet {
        ascii: FontSlot {
            explicit: xml::optional_attr(e, b"ascii")?,
            theme: xml::optional_attr(e, b"asciiTheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
        high_ansi: FontSlot {
            explicit: xml::optional_attr(e, b"hAnsi")?,
            theme: xml::optional_attr(e, b"hAnsiTheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
        east_asian: FontSlot {
            explicit: xml::optional_attr(e, b"eastAsia")?,
            theme: xml::optional_attr(e, b"eastAsiaTheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
        complex_script: FontSlot {
            explicit: xml::optional_attr(e, b"cs")?,
            theme: xml::optional_attr(e, b"cstheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
    })
}

/// §17.18.84 ST_Theme
fn parse_theme_font_ref(val: &str) -> Option<ThemeFontRef> {
    match val {
        "majorHAnsi" => Some(ThemeFontRef::MajorHAnsi),
        "majorEastAsia" => Some(ThemeFontRef::MajorEastAsia),
        "majorBidi" => Some(ThemeFontRef::MajorBidi),
        "minorHAnsi" => Some(ThemeFontRef::MinorHAnsi),
        "minorEastAsia" => Some(ThemeFontRef::MinorEastAsia),
        "minorBidi" => Some(ThemeFontRef::MinorBidi),
        _ => None,
    }
}

/// §17.18.99 ST_Underline
fn parse_underline_style(val: &str) -> Result<UnderlineStyle> {
    match val {
        "single" => Ok(UnderlineStyle::Single),
        "words" => Ok(UnderlineStyle::Words),
        "double" => Ok(UnderlineStyle::Double),
        "thick" => Ok(UnderlineStyle::Thick),
        "dotted" => Ok(UnderlineStyle::Dotted),
        "dottedHeavy" => Ok(UnderlineStyle::DottedHeavy),
        "dash" => Ok(UnderlineStyle::Dash),
        "dashedHeavy" => Ok(UnderlineStyle::DashedHeavy),
        "dashLong" => Ok(UnderlineStyle::DashLong),
        "dashLongHeavy" => Ok(UnderlineStyle::DashLongHeavy),
        "dotDash" => Ok(UnderlineStyle::DotDash),
        "dashDotHeavy" => Ok(UnderlineStyle::DashDotHeavy),
        "dotDotDash" => Ok(UnderlineStyle::DotDotDash),
        "dashDotDotHeavy" => Ok(UnderlineStyle::DashDotDotHeavy),
        "wave" => Ok(UnderlineStyle::Wave),
        "wavyHeavy" => Ok(UnderlineStyle::WavyHeavy),
        "wavyDouble" => Ok(UnderlineStyle::WavyDouble),
        "none" => Ok(UnderlineStyle::None),
        other => Err(invalid_value("u/val", other)),
    }
}

/// §17.18.100 ST_VerticalAlignRun
fn parse_vertical_align(val: &str) -> Result<VerticalAlign> {
    match val {
        "baseline" => Ok(VerticalAlign::Baseline),
        "superscript" => Ok(VerticalAlign::Superscript),
        "subscript" => Ok(VerticalAlign::Subscript),
        other => Err(invalid_value("vertAlign/val", other)),
    }
}

/// §17.18.40 ST_HighlightColor
fn parse_highlight_color(val: &str) -> Result<Option<HighlightColor>> {
    match val {
        "black" => Ok(Some(HighlightColor::Black)),
        "blue" => Ok(Some(HighlightColor::Blue)),
        "cyan" => Ok(Some(HighlightColor::Cyan)),
        "darkBlue" => Ok(Some(HighlightColor::DarkBlue)),
        "darkCyan" => Ok(Some(HighlightColor::DarkCyan)),
        "darkGray" => Ok(Some(HighlightColor::DarkGray)),
        "darkGreen" => Ok(Some(HighlightColor::DarkGreen)),
        "darkMagenta" => Ok(Some(HighlightColor::DarkMagenta)),
        "darkRed" => Ok(Some(HighlightColor::DarkRed)),
        "darkYellow" => Ok(Some(HighlightColor::DarkYellow)),
        "green" => Ok(Some(HighlightColor::Green)),
        "lightGray" => Ok(Some(HighlightColor::LightGray)),
        "magenta" => Ok(Some(HighlightColor::Magenta)),
        "red" => Ok(Some(HighlightColor::Red)),
        "white" => Ok(Some(HighlightColor::White)),
        "yellow" => Ok(Some(HighlightColor::Yellow)),
        "none" => Ok(None),
        other => Err(invalid_value("highlight/val", other)),
    }
}
