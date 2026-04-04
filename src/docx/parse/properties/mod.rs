//! Parsers for OOXML property elements: pPr, rPr, tblPr, trPr, tcPr, sectPr.
//!
//! Each parser consumes events from the reader until the corresponding End event,
//! returning a fully-populated properties struct.

mod paragraph;
mod run;
mod section;
mod table;

pub use paragraph::{parse_paragraph_properties, ParsedParagraphProperties};
pub use run::parse_run_properties;
pub use section::{parse_section_properties, parse_section_rsids};
pub use table::{parse_table_cell_properties, parse_table_properties, parse_table_row_properties};

use quick_xml::events::BytesStart;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::model::*;
use crate::docx::xml;

pub(super) fn invalid_value(attr: &str, value: &str) -> ParseError {
    ParseError::InvalidAttributeValue {
        attr: attr.to_string(),
        value: value.to_string(),
        reason: "unsupported value per OOXML spec".to_string(),
    }
}

/// Read a `w:val` boolean attribute, defaulting to `true` when the attribute is
/// absent (OOXML §17.17.4 "toggle property" semantics: presence alone means on).
///
/// Returns `Ok(Some(bool))` always; `Ok(None)` is never produced because toggle
/// properties with no `val` mean `true`, and absence of the element itself is
/// handled at the call site (the arm simply won't match).
#[inline]
pub(super) fn toggle_attr(e: &BytesStart<'_>) -> Result<Option<bool>> {
    Ok(Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true)))
}

/// Read a `w:val` string attribute and map it through `f`, returning
/// `Ok(None)` when the attribute is absent.
///
/// Equivalent to `xml::optional_attr(e, b"val")?.map(|v| f(&v)).transpose()`.
#[inline]
pub(super) fn opt_val<T, F>(e: &BytesStart<'_>, f: F) -> Result<Option<T>>
where
    F: FnOnce(&str) -> Result<T>,
{
    xml::optional_attr(e, b"val")?.map(|v| f(&v)).transpose()
}

// ── Shared parsing helpers ──────────────────────────────────────────────────

/// §17.18.44 ST_Jc
pub fn parse_alignment(val: &str) -> Result<Alignment> {
    match val {
        "start" | "left" => Ok(Alignment::Start),
        "center" => Ok(Alignment::Center),
        "end" | "right" => Ok(Alignment::End),
        "both" | "justify" => Ok(Alignment::Both),
        "distribute" => Ok(Alignment::Distribute),
        "thaiDistribute" => Ok(Alignment::Thai),
        other => Err(invalid_value("jc", other)),
    }
}

pub(super) fn parse_border(e: &BytesStart<'_>) -> Result<Border> {
    let style = match xml::optional_attr(e, b"val")?.as_deref() {
        Some("single") => BorderStyle::Single,
        Some("thick") => BorderStyle::Thick,
        Some("double") => BorderStyle::Double,
        Some("dotted") => BorderStyle::Dotted,
        Some("dashed") => BorderStyle::Dashed,
        Some("dotDash") => BorderStyle::DotDash,
        Some("dotDotDash") => BorderStyle::DotDotDash,
        Some("triple") => BorderStyle::Triple,
        Some("thinThickSmallGap") => BorderStyle::ThinThickSmallGap,
        Some("thickThinSmallGap") => BorderStyle::ThickThinSmallGap,
        Some("thinThickThinSmallGap") => BorderStyle::ThinThickThinSmallGap,
        Some("thinThickMediumGap") => BorderStyle::ThinThickMediumGap,
        Some("thickThinMediumGap") => BorderStyle::ThickThinMediumGap,
        Some("thinThickThinMediumGap") => BorderStyle::ThinThickThinMediumGap,
        Some("thinThickLargeGap") => BorderStyle::ThinThickLargeGap,
        Some("thickThinLargeGap") => BorderStyle::ThickThinLargeGap,
        Some("thinThickThinLargeGap") => BorderStyle::ThinThickThinLargeGap,
        Some("wave") => BorderStyle::Wave,
        Some("doubleWave") => BorderStyle::DoubleWave,
        Some("dashSmallGap") => BorderStyle::DashSmallGap,
        Some("dashDotStroked") => BorderStyle::DashDotStroked,
        Some("threeDEmboss") => BorderStyle::ThreeDEmboss,
        Some("threeDEngrave") => BorderStyle::ThreeDEngrave,
        Some("outset") => BorderStyle::Outset,
        Some("inset") => BorderStyle::Inset,
        Some("none") | Some("nil") | None => BorderStyle::None,
        Some(other) => return Err(invalid_value("border/val", other)),
    };

    let sz = xml::optional_attr_i64(e, b"sz")?.unwrap_or(0);
    let space = xml::optional_attr_i64(e, b"space")?.unwrap_or(0);
    let color = parse_color_from_attr(e)?;

    Ok(Border {
        style,
        width: Dimension::new(sz),
        // §17.3.4: w:space is ST_PointMeasure (§17.18.68).
        space: Dimension::new(space),
        color,
    })
}

pub(super) fn parse_shading(e: &BytesStart<'_>) -> Result<Shading> {
    let fill = match xml::optional_attr(e, b"fill")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Color::Auto,
        Some(ref s) => xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?,
        None => Color::Auto,
    };

    let color = match xml::optional_attr(e, b"color")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Color::Auto,
        Some(ref s) => xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?,
        None => Color::Auto,
    };

    let pattern = match xml::optional_attr(e, b"val")?.as_deref() {
        Some("clear") | None => ShadingPattern::Clear,
        Some("solid") => ShadingPattern::Solid,
        Some("horzStripe") => ShadingPattern::HorzStripe,
        Some("vertStripe") => ShadingPattern::VertStripe,
        Some("reverseDiagStripe") => ShadingPattern::ReverseDiagStripe,
        Some("diagStripe") => ShadingPattern::DiagStripe,
        Some("horzCross") => ShadingPattern::HorzCross,
        Some("diagCross") => ShadingPattern::DiagCross,
        Some("thinHorzStripe") => ShadingPattern::ThinHorzStripe,
        Some("thinVertStripe") => ShadingPattern::ThinVertStripe,
        Some("thinReverseDiagStripe") => ShadingPattern::ThinReverseDiagStripe,
        Some("thinDiagStripe") => ShadingPattern::ThinDiagStripe,
        Some("thinHorzCross") => ShadingPattern::ThinHorzCross,
        Some("thinDiagCross") => ShadingPattern::ThinDiagCross,
        Some("pct5") => ShadingPattern::Pct5,
        Some("pct10") => ShadingPattern::Pct10,
        Some("pct12") => ShadingPattern::Pct12,
        Some("pct15") => ShadingPattern::Pct15,
        Some("pct20") => ShadingPattern::Pct20,
        Some("pct25") => ShadingPattern::Pct25,
        Some("pct30") => ShadingPattern::Pct30,
        Some("pct35") => ShadingPattern::Pct35,
        Some("pct37") => ShadingPattern::Pct37,
        Some("pct40") => ShadingPattern::Pct40,
        Some("pct45") => ShadingPattern::Pct45,
        Some("pct50") => ShadingPattern::Pct50,
        Some("pct55") => ShadingPattern::Pct55,
        Some("pct60") => ShadingPattern::Pct60,
        Some("pct62") => ShadingPattern::Pct62,
        Some("pct65") => ShadingPattern::Pct65,
        Some("pct70") => ShadingPattern::Pct70,
        Some("pct75") => ShadingPattern::Pct75,
        Some("pct80") => ShadingPattern::Pct80,
        Some("pct85") => ShadingPattern::Pct85,
        Some("pct87") => ShadingPattern::Pct87,
        Some("pct90") => ShadingPattern::Pct90,
        Some("pct95") => ShadingPattern::Pct95,
        Some(other) => return Err(invalid_value("shd/val", other)),
    };

    Ok(Shading {
        fill,
        pattern,
        color,
    })
}

pub(super) fn parse_color_attr(e: &BytesStart<'_>) -> Result<Color> {
    match xml::optional_attr(e, b"val")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Ok(Color::Auto),
        Some(ref s) => Ok(xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?),
        None => Ok(Color::Auto),
    }
}

pub(super) fn parse_color_from_attr(e: &BytesStart<'_>) -> Result<Color> {
    match xml::optional_attr(e, b"color")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Ok(Color::Auto),
        Some(ref s) => Ok(xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?),
        None => Ok(Color::Auto),
    }
}

pub(super) fn parse_edge_insets_twips(
    reader: &mut quick_xml::Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<crate::docx::geometry::EdgeInsets<crate::docx::dimension::Twips>> {
    use crate::docx::geometry::EdgeInsets;
    use quick_xml::events::Event;

    let mut insets = EdgeInsets::ZERO;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
                match local {
                    b"top" => insets.top = Dimension::new(w),
                    b"bottom" => insets.bottom = Dimension::new(w),
                    b"left" | b"start" => insets.left = Dimension::new(w),
                    b"right" | b"end" => insets.right = Dimension::new(w),
                    _ => xml::warn_unsupported_element("margins", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(b"container")),
            _ => {}
        }
    }

    Ok(insets)
}

/// §17.3.1.8: parse `w:cnfStyle` element attributes into a [`CnfStyle`] flag set.
///
/// The `val` binary string is parsed first (positions 0-11 -> flags), then any
/// explicit individual attributes override the corresponding bits, allowing
/// producers that omit `val` to be handled correctly.
pub(super) fn parse_cnf_style(e: &BytesStart<'_>) -> Result<CnfStyle> {
    // Seed from the legacy 12-char binary string if present.
    let mut flags = match xml::optional_attr(e, b"val")? {
        Some(s) => CnfStyle::from_val_str(&s),
        None => CnfStyle::empty(),
    };

    // Individual attributes take precedence over the `val` string.
    let pairs: &[(&[u8], CnfStyle)] = &[
        (b"firstRow", CnfStyle::FIRST_ROW),
        (b"lastRow", CnfStyle::LAST_ROW),
        (b"firstColumn", CnfStyle::FIRST_COLUMN),
        (b"lastColumn", CnfStyle::LAST_COLUMN),
        (b"oddVBand", CnfStyle::ODD_V_BAND),
        (b"evenVBand", CnfStyle::EVEN_V_BAND),
        (b"oddHBand", CnfStyle::ODD_H_BAND),
        (b"evenHBand", CnfStyle::EVEN_H_BAND),
        (b"firstRowFirstColumn", CnfStyle::FIRST_ROW_FIRST_COLUMN),
        (b"firstRowLastColumn", CnfStyle::FIRST_ROW_LAST_COLUMN),
        (b"lastRowFirstColumn", CnfStyle::LAST_ROW_FIRST_COLUMN),
        (b"lastRowLastColumn", CnfStyle::LAST_ROW_LAST_COLUMN),
    ];
    for &(attr, flag) in pairs {
        match xml::optional_attr_bool(e, attr)? {
            Some(true) => flags |= flag,
            Some(false) => flags &= !flag,
            None => {} // absent -- leave the val-seeded bit unchanged
        }
    }

    Ok(flags)
}
