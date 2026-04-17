//! Parser for DrawingML colors (§20.1.2.3 EG_ColorChoice).
//!
//! Each color-choice element (`scRgbClr`, `srgbClr`, `hslClr`, `sysClr`,
//! `schemeClr`, `prstClr`) may carry an ordered list of color transforms as
//! children. Transforms preserve document order.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::model::{
    ColorTransform, DrawingColor, PresetColorVal, SchemeColorVal, SystemColorVal,
};
use crate::docx::xml;

/// Parse a color-choice element. The reader is positioned at the element's
/// Start or Empty event (caller supplies it via `start`).
///
/// Consumes events through the matching End (or returns immediately if the
/// caller's event was `Empty`, which must be signalled via `is_empty`).
///
/// Returns `Ok(None)` if the element is outside the recognized choice set;
/// the caller is responsible for logging and advancing past the element.
pub fn parse_color_choice(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    start: &BytesStart<'_>,
    is_empty: bool,
) -> Result<Option<DrawingColor>> {
    let qn = start.name();
    let local = xml::local_name(qn.as_ref());
    let transforms = |reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>, end: &[u8]| {
        if is_empty {
            Ok::<Vec<ColorTransform>, crate::docx::error::ParseError>(Vec::new())
        } else {
            parse_color_transforms(reader, buf, end)
        }
    };

    match local {
        b"srgbClr" => {
            let rgb = parse_srgb_val(start)?;
            let transforms = transforms(reader, buf, b"srgbClr")?;
            Ok(Some(DrawingColor::Srgb { rgb, transforms }))
        }
        b"scRgbClr" => {
            let (r, g, b) = parse_scrgb_attrs(start)?;
            let transforms = transforms(reader, buf, b"scRgbClr")?;
            Ok(Some(DrawingColor::ScRgb {
                r,
                g,
                b,
                transforms,
            }))
        }
        b"hslClr" => {
            let (hue, sat, lum) = parse_hsl_attrs(start)?;
            let transforms = transforms(reader, buf, b"hslClr")?;
            Ok(Some(DrawingColor::Hsl {
                hue,
                sat,
                lum,
                transforms,
            }))
        }
        b"sysClr" => {
            let (name, last_clr) = parse_sys_attrs(start)?;
            let transforms = transforms(reader, buf, b"sysClr")?;
            Ok(Some(DrawingColor::Sys {
                name,
                last_clr,
                transforms,
            }))
        }
        b"schemeClr" => {
            let name = parse_scheme_attrs(start)?;
            let transforms = transforms(reader, buf, b"schemeClr")?;
            Ok(Some(DrawingColor::Scheme { name, transforms }))
        }
        b"prstClr" => {
            let name = parse_prst_attrs(start)?;
            let transforms = transforms(reader, buf, b"prstClr")?;
            Ok(Some(DrawingColor::Prst { name, transforms }))
        }
        _ => Ok(None),
    }
}

// ── Base color attribute parsers ────────────────────────────────────────────

fn parse_srgb_val(e: &BytesStart<'_>) -> Result<u32> {
    let val = xml::required_attr(e, b"val")?;
    u32::from_str_radix(&val, 16).map_err(|_| ParseError::InvalidAttributeValue {
        attr: "val".into(),
        value: val,
        reason: "expected 6 hex digits per §20.1.2.3.32 ST_HexColorRGB".into(),
    })
}

fn parse_scrgb_attrs(
    e: &BytesStart<'_>,
) -> Result<(
    Dimension<crate::docx::dimension::ThousandthPercent>,
    Dimension<crate::docx::dimension::ThousandthPercent>,
    Dimension<crate::docx::dimension::ThousandthPercent>,
)> {
    let r = required_thousandth_percent(e, b"r")?;
    let g = required_thousandth_percent(e, b"g")?;
    let b = required_thousandth_percent(e, b"b")?;
    Ok((r, g, b))
}

fn parse_hsl_attrs(
    e: &BytesStart<'_>,
) -> Result<(
    Dimension<crate::docx::dimension::SixtieThousandthDeg>,
    Dimension<crate::docx::dimension::ThousandthPercent>,
    Dimension<crate::docx::dimension::ThousandthPercent>,
)> {
    let hue = required_sixtie_thousandth_deg(e, b"hue")?;
    let sat = required_thousandth_percent(e, b"sat")?;
    let lum = required_thousandth_percent(e, b"lum")?;
    Ok((hue, sat, lum))
}

fn parse_sys_attrs(e: &BytesStart<'_>) -> Result<(SystemColorVal, Option<u32>)> {
    let val = xml::required_attr(e, b"val")?;
    let name = parse_system_color_val(&val)?;
    let last_clr =
        match xml::optional_attr(e, b"lastClr")? {
            Some(s) => Some(u32::from_str_radix(&s, 16).map_err(|_| {
                ParseError::InvalidAttributeValue {
                    attr: "lastClr".into(),
                    value: s,
                    reason: "expected 6 hex digits per §20.1.2.3.32 ST_HexColorRGB".into(),
                }
            })?),
            None => None,
        };
    Ok((name, last_clr))
}

fn parse_scheme_attrs(e: &BytesStart<'_>) -> Result<SchemeColorVal> {
    let val = xml::required_attr(e, b"val")?;
    parse_scheme_color_val(&val)
}

fn parse_prst_attrs(e: &BytesStart<'_>) -> Result<PresetColorVal> {
    let val = xml::required_attr(e, b"val")?;
    parse_preset_color_val(&val)
}

// ── Transform list ──────────────────────────────────────────────────────────

/// Parse color transform children until the End event of `end_tag`.
/// Preserves document order. Unknown children log and skip.
fn parse_color_transforms(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<Vec<ColorTransform>> {
    let mut transforms = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) => {
                if let Some(t) = parse_color_transform(e)? {
                    transforms.push(t);
                } else {
                    let qn = e.name();
                    let local = xml::local_name(qn.as_ref());
                    xml::warn_unsupported_element("colorTransform", local);
                }
            }
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let local_owned = xml::local_name_owned(qn.as_ref());
                if let Some(t) = parse_color_transform(e)? {
                    transforms.push(t);
                } else {
                    xml::warn_unsupported_element("colorTransform", local);
                }
                xml::skip_to_end(reader, buf, &local_owned)?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }

    Ok(transforms)
}

/// Parse a single color-transform element into `ColorTransform`. Returns
/// `Ok(None)` for unrecognized elements; caller logs and skips.
fn parse_color_transform(e: &BytesStart<'_>) -> Result<Option<ColorTransform>> {
    let qn = e.name();
    let local = xml::local_name(qn.as_ref());

    // Parameterless transforms.
    match local {
        b"comp" => return Ok(Some(ColorTransform::Comp)),
        b"inv" => return Ok(Some(ColorTransform::Inv)),
        b"gray" => return Ok(Some(ColorTransform::Gray)),
        b"gamma" => return Ok(Some(ColorTransform::Gamma)),
        b"invGamma" => return Ok(Some(ColorTransform::InvGamma)),
        _ => {}
    }

    // Percentage-valued transforms (thousandth-percent).
    let pct = |ct: fn(Dimension<crate::docx::dimension::ThousandthPercent>) -> ColorTransform,
               elem: &BytesStart<'_>|
     -> Result<Option<ColorTransform>> {
        Ok(Some(ct(required_thousandth_percent(elem, b"val")?)))
    };

    match local {
        b"tint" => return pct(ColorTransform::Tint, e),
        b"shade" => return pct(ColorTransform::Shade, e),
        b"alpha" => return pct(ColorTransform::Alpha, e),
        b"alphaOff" => return pct(ColorTransform::AlphaOff, e),
        b"alphaMod" => return pct(ColorTransform::AlphaMod, e),
        b"hueMod" => return pct(ColorTransform::HueMod, e),
        b"sat" => return pct(ColorTransform::Sat, e),
        b"satOff" => return pct(ColorTransform::SatOff, e),
        b"satMod" => return pct(ColorTransform::SatMod, e),
        b"lum" => return pct(ColorTransform::Lum, e),
        b"lumOff" => return pct(ColorTransform::LumOff, e),
        b"lumMod" => return pct(ColorTransform::LumMod, e),
        b"red" => return pct(ColorTransform::Red, e),
        b"redOff" => return pct(ColorTransform::RedOff, e),
        b"redMod" => return pct(ColorTransform::RedMod, e),
        b"green" => return pct(ColorTransform::Green, e),
        b"greenOff" => return pct(ColorTransform::GreenOff, e),
        b"greenMod" => return pct(ColorTransform::GreenMod, e),
        b"blue" => return pct(ColorTransform::Blue, e),
        b"blueOff" => return pct(ColorTransform::BlueOff, e),
        b"blueMod" => return pct(ColorTransform::BlueMod, e),
        _ => {}
    }

    // Angle-valued transforms (60000ths of a degree).
    match local {
        b"hue" => Ok(Some(ColorTransform::Hue(required_sixtie_thousandth_deg(
            e, b"val",
        )?))),
        b"hueOff" => Ok(Some(ColorTransform::HueOff(
            required_sixtie_thousandth_deg(e, b"val")?,
        ))),
        _ => Ok(None),
    }
}

// ── Attribute typed-value helpers ───────────────────────────────────────────

fn required_thousandth_percent(
    e: &BytesStart<'_>,
    attr: &[u8],
) -> Result<Dimension<crate::docx::dimension::ThousandthPercent>> {
    let v = xml::optional_attr_i64(e, attr)?.ok_or_else(|| ParseError::MissingAttribute {
        element: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
        attr: String::from_utf8_lossy(attr).into_owned(),
    })?;
    Ok(Dimension::new(v))
}

fn required_sixtie_thousandth_deg(
    e: &BytesStart<'_>,
    attr: &[u8],
) -> Result<Dimension<crate::docx::dimension::SixtieThousandthDeg>> {
    let v = xml::optional_attr_i64(e, attr)?.ok_or_else(|| ParseError::MissingAttribute {
        element: String::from_utf8_lossy(e.name().as_ref()).into_owned(),
        attr: String::from_utf8_lossy(attr).into_owned(),
    })?;
    Ok(Dimension::new(v))
}

// ── ST_SchemeColorVal (§20.1.10.54) ────────────────────────────────────────

fn parse_scheme_color_val(val: &str) -> Result<SchemeColorVal> {
    Ok(match val {
        "bg1" => SchemeColorVal::Bg1,
        "tx1" => SchemeColorVal::Tx1,
        "bg2" => SchemeColorVal::Bg2,
        "tx2" => SchemeColorVal::Tx2,
        "accent1" => SchemeColorVal::Accent1,
        "accent2" => SchemeColorVal::Accent2,
        "accent3" => SchemeColorVal::Accent3,
        "accent4" => SchemeColorVal::Accent4,
        "accent5" => SchemeColorVal::Accent5,
        "accent6" => SchemeColorVal::Accent6,
        "hlink" => SchemeColorVal::Hlink,
        "folHlink" => SchemeColorVal::FolHlink,
        "phClr" => SchemeColorVal::PhClr,
        "dk1" => SchemeColorVal::Dk1,
        "lt1" => SchemeColorVal::Lt1,
        "dk2" => SchemeColorVal::Dk2,
        "lt2" => SchemeColorVal::Lt2,
        other => {
            return Err(ParseError::InvalidAttributeValue {
                attr: "val".into(),
                value: other.into(),
                reason: "expected value per §20.1.10.54 ST_SchemeColorVal".into(),
            })
        }
    })
}

// ── ST_SystemColorVal (§20.1.10.57) ────────────────────────────────────────

fn parse_system_color_val(val: &str) -> Result<SystemColorVal> {
    Ok(match val {
        "scrollBar" => SystemColorVal::ScrollBar,
        "background" => SystemColorVal::Background,
        "activeCaption" => SystemColorVal::ActiveCaption,
        "inactiveCaption" => SystemColorVal::InactiveCaption,
        "menu" => SystemColorVal::Menu,
        "window" => SystemColorVal::Window,
        "windowFrame" => SystemColorVal::WindowFrame,
        "menuText" => SystemColorVal::MenuText,
        "windowText" => SystemColorVal::WindowText,
        "captionText" => SystemColorVal::CaptionText,
        "activeBorder" => SystemColorVal::ActiveBorder,
        "inactiveBorder" => SystemColorVal::InactiveBorder,
        "appWorkspace" => SystemColorVal::AppWorkspace,
        "highlight" => SystemColorVal::Highlight,
        "highlightText" => SystemColorVal::HighlightText,
        "btnFace" => SystemColorVal::BtnFace,
        "btnShadow" => SystemColorVal::BtnShadow,
        "grayText" => SystemColorVal::GrayText,
        "btnText" => SystemColorVal::BtnText,
        "inactiveCaptionText" => SystemColorVal::InactiveCaptionText,
        "btnHighlight" => SystemColorVal::BtnHighlight,
        "3dDkShadow" => SystemColorVal::ThreeDDkShadow,
        "3dLight" => SystemColorVal::ThreeDLight,
        "infoText" => SystemColorVal::InfoText,
        "infoBk" => SystemColorVal::InfoBk,
        "hotLight" => SystemColorVal::HotLight,
        "gradientActiveCaption" => SystemColorVal::GradientActiveCaption,
        "gradientInactiveCaption" => SystemColorVal::GradientInactiveCaption,
        "menuHighlight" => SystemColorVal::MenuHighlight,
        "menuBar" => SystemColorVal::MenuBar,
        other => {
            return Err(ParseError::InvalidAttributeValue {
                attr: "val".into(),
                value: other.into(),
                reason: "expected value per §20.1.10.57 ST_SystemColorVal".into(),
            })
        }
    })
}

// ── ST_PresetColorVal (§20.1.10.47) ────────────────────────────────────────

fn parse_preset_color_val(val: &str) -> Result<PresetColorVal> {
    Ok(match val {
        "aliceBlue" => PresetColorVal::AliceBlue,
        "antiqueWhite" => PresetColorVal::AntiqueWhite,
        "aqua" => PresetColorVal::Aqua,
        "aquamarine" => PresetColorVal::Aquamarine,
        "azure" => PresetColorVal::Azure,
        "beige" => PresetColorVal::Beige,
        "bisque" => PresetColorVal::Bisque,
        "black" => PresetColorVal::Black,
        "blanchedAlmond" => PresetColorVal::BlanchedAlmond,
        "blue" => PresetColorVal::Blue,
        "blueViolet" => PresetColorVal::BlueViolet,
        "brown" => PresetColorVal::Brown,
        "burlyWood" => PresetColorVal::BurlyWood,
        "cadetBlue" => PresetColorVal::CadetBlue,
        "chartreuse" => PresetColorVal::Chartreuse,
        "chocolate" => PresetColorVal::Chocolate,
        "coral" => PresetColorVal::Coral,
        "cornflowerBlue" => PresetColorVal::CornflowerBlue,
        "cornsilk" => PresetColorVal::Cornsilk,
        "crimson" => PresetColorVal::Crimson,
        "cyan" => PresetColorVal::Cyan,
        "darkBlue" => PresetColorVal::DarkBlue,
        "darkCyan" => PresetColorVal::DarkCyan,
        "darkGoldenrod" => PresetColorVal::DarkGoldenrod,
        "darkGray" => PresetColorVal::DarkGray,
        "darkGreen" => PresetColorVal::DarkGreen,
        "darkGrey" => PresetColorVal::DarkGrey,
        "darkKhaki" => PresetColorVal::DarkKhaki,
        "darkMagenta" => PresetColorVal::DarkMagenta,
        "darkOliveGreen" => PresetColorVal::DarkOliveGreen,
        "darkOrange" => PresetColorVal::DarkOrange,
        "darkOrchid" => PresetColorVal::DarkOrchid,
        "darkRed" => PresetColorVal::DarkRed,
        "darkSalmon" => PresetColorVal::DarkSalmon,
        "darkSeaGreen" => PresetColorVal::DarkSeaGreen,
        "darkSlateBlue" => PresetColorVal::DarkSlateBlue,
        "darkSlateGray" => PresetColorVal::DarkSlateGray,
        "darkSlateGrey" => PresetColorVal::DarkSlateGrey,
        "darkTurquoise" => PresetColorVal::DarkTurquoise,
        "darkViolet" => PresetColorVal::DarkViolet,
        "deepPink" => PresetColorVal::DeepPink,
        "deepSkyBlue" => PresetColorVal::DeepSkyBlue,
        "dimGray" => PresetColorVal::DimGray,
        "dimGrey" => PresetColorVal::DimGrey,
        "dkBlue" => PresetColorVal::DkBlue,
        "dkCyan" => PresetColorVal::DkCyan,
        "dkGoldenrod" => PresetColorVal::DkGoldenrod,
        "dkGray" => PresetColorVal::DkGray,
        "dkGreen" => PresetColorVal::DkGreen,
        "dkGrey" => PresetColorVal::DkGrey,
        "dkKhaki" => PresetColorVal::DkKhaki,
        "dkMagenta" => PresetColorVal::DkMagenta,
        "dkOliveGreen" => PresetColorVal::DkOliveGreen,
        "dkOrange" => PresetColorVal::DkOrange,
        "dkOrchid" => PresetColorVal::DkOrchid,
        "dkRed" => PresetColorVal::DkRed,
        "dkSalmon" => PresetColorVal::DkSalmon,
        "dkSeaGreen" => PresetColorVal::DkSeaGreen,
        "dkSlateBlue" => PresetColorVal::DkSlateBlue,
        "dkSlateGray" => PresetColorVal::DkSlateGray,
        "dkSlateGrey" => PresetColorVal::DkSlateGrey,
        "dkTurquoise" => PresetColorVal::DkTurquoise,
        "dkViolet" => PresetColorVal::DkViolet,
        "dodgerBlue" => PresetColorVal::DodgerBlue,
        "firebrick" => PresetColorVal::Firebrick,
        "floralWhite" => PresetColorVal::FloralWhite,
        "forestGreen" => PresetColorVal::ForestGreen,
        "fuchsia" => PresetColorVal::Fuchsia,
        "gainsboro" => PresetColorVal::Gainsboro,
        "ghostWhite" => PresetColorVal::GhostWhite,
        "gold" => PresetColorVal::Gold,
        "goldenrod" => PresetColorVal::Goldenrod,
        "gray" => PresetColorVal::Gray,
        "green" => PresetColorVal::Green,
        "greenYellow" => PresetColorVal::GreenYellow,
        "grey" => PresetColorVal::Grey,
        "honeydew" => PresetColorVal::Honeydew,
        "hotPink" => PresetColorVal::HotPink,
        "indianRed" => PresetColorVal::IndianRed,
        "indigo" => PresetColorVal::Indigo,
        "ivory" => PresetColorVal::Ivory,
        "khaki" => PresetColorVal::Khaki,
        "lavender" => PresetColorVal::Lavender,
        "lavenderBlush" => PresetColorVal::LavenderBlush,
        "lawnGreen" => PresetColorVal::LawnGreen,
        "lemonChiffon" => PresetColorVal::LemonChiffon,
        "lightBlue" => PresetColorVal::LightBlue,
        "lightCoral" => PresetColorVal::LightCoral,
        "lightCyan" => PresetColorVal::LightCyan,
        "lightGoldenrodYellow" => PresetColorVal::LightGoldenrodYellow,
        "lightGray" => PresetColorVal::LightGray,
        "lightGreen" => PresetColorVal::LightGreen,
        "lightGrey" => PresetColorVal::LightGrey,
        "lightPink" => PresetColorVal::LightPink,
        "lightSalmon" => PresetColorVal::LightSalmon,
        "lightSeaGreen" => PresetColorVal::LightSeaGreen,
        "lightSkyBlue" => PresetColorVal::LightSkyBlue,
        "lightSlateGray" => PresetColorVal::LightSlateGray,
        "lightSlateGrey" => PresetColorVal::LightSlateGrey,
        "lightSteelBlue" => PresetColorVal::LightSteelBlue,
        "lightYellow" => PresetColorVal::LightYellow,
        "lime" => PresetColorVal::Lime,
        "limeGreen" => PresetColorVal::LimeGreen,
        "linen" => PresetColorVal::Linen,
        "ltBlue" => PresetColorVal::LtBlue,
        "ltCoral" => PresetColorVal::LtCoral,
        "ltCyan" => PresetColorVal::LtCyan,
        "ltGoldenrodYellow" => PresetColorVal::LtGoldenrodYellow,
        "ltGray" => PresetColorVal::LtGray,
        "ltGreen" => PresetColorVal::LtGreen,
        "ltGrey" => PresetColorVal::LtGrey,
        "ltPink" => PresetColorVal::LtPink,
        "ltSalmon" => PresetColorVal::LtSalmon,
        "ltSeaGreen" => PresetColorVal::LtSeaGreen,
        "ltSkyBlue" => PresetColorVal::LtSkyBlue,
        "ltSlateGray" => PresetColorVal::LtSlateGray,
        "ltSlateGrey" => PresetColorVal::LtSlateGrey,
        "ltSteelBlue" => PresetColorVal::LtSteelBlue,
        "ltYellow" => PresetColorVal::LtYellow,
        "magenta" => PresetColorVal::Magenta,
        "maroon" => PresetColorVal::Maroon,
        "medAquamarine" => PresetColorVal::MedAquamarine,
        "medBlue" => PresetColorVal::MedBlue,
        "medOrchid" => PresetColorVal::MedOrchid,
        "medPurple" => PresetColorVal::MedPurple,
        "medSeaGreen" => PresetColorVal::MedSeaGreen,
        "medSlateBlue" => PresetColorVal::MedSlateBlue,
        "medSpringGreen" => PresetColorVal::MedSpringGreen,
        "medTurquoise" => PresetColorVal::MedTurquoise,
        "medVioletRed" => PresetColorVal::MedVioletRed,
        "mediumAquamarine" => PresetColorVal::MediumAquamarine,
        "mediumBlue" => PresetColorVal::MediumBlue,
        "mediumOrchid" => PresetColorVal::MediumOrchid,
        "mediumPurple" => PresetColorVal::MediumPurple,
        "mediumSeaGreen" => PresetColorVal::MediumSeaGreen,
        "mediumSlateBlue" => PresetColorVal::MediumSlateBlue,
        "mediumSpringGreen" => PresetColorVal::MediumSpringGreen,
        "mediumTurquoise" => PresetColorVal::MediumTurquoise,
        "mediumVioletRed" => PresetColorVal::MediumVioletRed,
        "midnightBlue" => PresetColorVal::MidnightBlue,
        "mintCream" => PresetColorVal::MintCream,
        "mistyRose" => PresetColorVal::MistyRose,
        "moccasin" => PresetColorVal::Moccasin,
        "navajoWhite" => PresetColorVal::NavajoWhite,
        "navy" => PresetColorVal::Navy,
        "oldLace" => PresetColorVal::OldLace,
        "olive" => PresetColorVal::Olive,
        "oliveDrab" => PresetColorVal::OliveDrab,
        "orange" => PresetColorVal::Orange,
        "orangeRed" => PresetColorVal::OrangeRed,
        "orchid" => PresetColorVal::Orchid,
        "paleGoldenrod" => PresetColorVal::PaleGoldenrod,
        "paleGreen" => PresetColorVal::PaleGreen,
        "paleTurquoise" => PresetColorVal::PaleTurquoise,
        "paleVioletRed" => PresetColorVal::PaleVioletRed,
        "papayaWhip" => PresetColorVal::PapayaWhip,
        "peachPuff" => PresetColorVal::PeachPuff,
        "peru" => PresetColorVal::Peru,
        "pink" => PresetColorVal::Pink,
        "plum" => PresetColorVal::Plum,
        "powderBlue" => PresetColorVal::PowderBlue,
        "purple" => PresetColorVal::Purple,
        "red" => PresetColorVal::Red,
        "rosyBrown" => PresetColorVal::RosyBrown,
        "royalBlue" => PresetColorVal::RoyalBlue,
        "saddleBrown" => PresetColorVal::SaddleBrown,
        "salmon" => PresetColorVal::Salmon,
        "sandyBrown" => PresetColorVal::SandyBrown,
        "seaGreen" => PresetColorVal::SeaGreen,
        "seaShell" => PresetColorVal::SeaShell,
        "sienna" => PresetColorVal::Sienna,
        "silver" => PresetColorVal::Silver,
        "skyBlue" => PresetColorVal::SkyBlue,
        "slateBlue" => PresetColorVal::SlateBlue,
        "slateGray" => PresetColorVal::SlateGray,
        "slateGrey" => PresetColorVal::SlateGrey,
        "snow" => PresetColorVal::Snow,
        "springGreen" => PresetColorVal::SpringGreen,
        "steelBlue" => PresetColorVal::SteelBlue,
        "tan" => PresetColorVal::Tan,
        "teal" => PresetColorVal::Teal,
        "thistle" => PresetColorVal::Thistle,
        "tomato" => PresetColorVal::Tomato,
        "turquoise" => PresetColorVal::Turquoise,
        "violet" => PresetColorVal::Violet,
        "wheat" => PresetColorVal::Wheat,
        "white" => PresetColorVal::White,
        "whiteSmoke" => PresetColorVal::WhiteSmoke,
        "yellow" => PresetColorVal::Yellow,
        "yellowGreen" => PresetColorVal::YellowGreen,
        other => {
            return Err(ParseError::InvalidAttributeValue {
                attr: "val".into(),
                value: other.into(),
                reason: "expected value per §20.1.10.47 ST_PresetColorVal".into(),
            })
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::Reader;

    fn parse(xml_src: &str) -> DrawingColor {
        let mut reader = Reader::from_reader(xml_src.as_bytes());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Start(ref e) => {
                    return parse_color_choice(&mut reader, &mut buf, e, false)
                        .unwrap()
                        .unwrap();
                }
                Event::Empty(ref e) => {
                    return parse_color_choice(&mut reader, &mut buf, e, true)
                        .unwrap()
                        .unwrap();
                }
                Event::Eof => panic!("no color element"),
                _ => {}
            }
        }
    }

    #[test]
    fn srgb_empty_no_transforms() {
        let c = parse(r#"<a:srgbClr xmlns:a="urn:a" val="AABBCC"/>"#);
        assert!(matches!(c, DrawingColor::Srgb { rgb: 0xAABBCC, .. }));
        assert!(c.transforms().is_empty());
    }

    #[test]
    fn srgb_with_shade_tint() {
        let c = parse(
            r#"<a:srgbClr xmlns:a="urn:a" val="112233">
                <a:shade val="50000"/>
                <a:tint val="25000"/>
            </a:srgbClr>"#,
        );
        let DrawingColor::Srgb { rgb, transforms } = c else {
            panic!()
        };
        assert_eq!(rgb, 0x112233);
        assert_eq!(transforms.len(), 2);
        assert!(matches!(transforms[0], ColorTransform::Shade(_)));
        assert!(matches!(transforms[1], ColorTransform::Tint(_)));
    }

    #[test]
    fn scheme_accent_with_lum_mod_off() {
        let c = parse(
            r#"<a:schemeClr xmlns:a="urn:a" val="accent1">
                <a:lumMod val="75000"/>
                <a:lumOff val="25000"/>
            </a:schemeClr>"#,
        );
        let DrawingColor::Scheme { name, transforms } = c else {
            panic!()
        };
        assert_eq!(name, SchemeColorVal::Accent1);
        assert_eq!(transforms.len(), 2);
    }

    #[test]
    fn sys_color_with_last_clr() {
        let c = parse(r#"<a:sysClr xmlns:a="urn:a" val="window" lastClr="FFFFFF"/>"#);
        let DrawingColor::Sys { name, last_clr, .. } = c else {
            panic!()
        };
        assert_eq!(name, SystemColorVal::Window);
        assert_eq!(last_clr, Some(0xFFFFFF));
    }

    #[test]
    fn prst_color() {
        let c = parse(r#"<a:prstClr xmlns:a="urn:a" val="chocolate"/>"#);
        assert!(matches!(
            c,
            DrawingColor::Prst {
                name: PresetColorVal::Chocolate,
                ..
            }
        ));
    }

    #[test]
    fn hsl_color_with_sat_mod() {
        let c = parse(
            r#"<a:hslClr xmlns:a="urn:a" hue="10800000" sat="100000" lum="50000">
                <a:satMod val="80000"/>
            </a:hslClr>"#,
        );
        let DrawingColor::Hsl {
            hue,
            sat,
            lum,
            transforms,
        } = c
        else {
            panic!()
        };
        assert_eq!(hue.raw(), 10_800_000);
        assert_eq!(sat.raw(), 100_000);
        assert_eq!(lum.raw(), 50_000);
        assert_eq!(transforms.len(), 1);
    }

    #[test]
    fn scrgb_color() {
        let c = parse(r#"<a:scRgbClr xmlns:a="urn:a" r="50000" g="75000" b="100000"/>"#);
        let DrawingColor::ScRgb { r, g, b, .. } = c else {
            panic!()
        };
        assert_eq!(r.raw(), 50_000);
        assert_eq!(g.raw(), 75_000);
        assert_eq!(b.raw(), 100_000);
    }

    #[test]
    fn transform_order_preserved() {
        let c = parse(
            r#"<a:srgbClr xmlns:a="urn:a" val="FF0000">
                <a:alpha val="80000"/>
                <a:lumMod val="50000"/>
                <a:satMod val="90000"/>
                <a:comp/>
            </a:srgbClr>"#,
        );
        let DrawingColor::Srgb { transforms, .. } = c else {
            panic!()
        };
        assert!(matches!(transforms[0], ColorTransform::Alpha(_)));
        assert!(matches!(transforms[1], ColorTransform::LumMod(_)));
        assert!(matches!(transforms[2], ColorTransform::SatMod(_)));
        assert!(matches!(transforms[3], ColorTransform::Comp));
    }

    #[test]
    fn invalid_scheme_val_errors() {
        let mut reader =
            Reader::from_reader(r#"<a:schemeClr xmlns:a="urn:a" val="bogus"/>"#.as_bytes());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Empty(ref e) => {
                    let res = parse_color_choice(&mut reader, &mut buf, e, true);
                    assert!(res.is_err());
                    return;
                }
                Event::Start(ref e) => {
                    let res = parse_color_choice(&mut reader, &mut buf, e, false);
                    assert!(res.is_err());
                    return;
                }
                Event::Eof => panic!(),
                _ => {}
            }
        }
    }

    #[test]
    fn unknown_transform_element_warns_and_skips() {
        // <a:weird/> should be ignored; the <a:shade/> should still be captured.
        let c = parse(
            r#"<a:srgbClr xmlns:a="urn:a" val="112233">
                <a:weird val="1"/>
                <a:shade val="50000"/>
            </a:srgbClr>"#,
        );
        let DrawingColor::Srgb { transforms, .. } = c else {
            panic!()
        };
        assert_eq!(transforms.len(), 1);
        assert!(matches!(transforms[0], ColorTransform::Shade(_)));
    }
}
