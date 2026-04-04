use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

use super::{invalid_value, opt_val, toggle_attr};

/// Parse `w:sectPr`. Reader should have just read the Start event.
/// Extract rsid attributes from a `w:sectPr` Start event.
pub fn parse_section_rsids(e: &BytesStart<'_>) -> Result<SectionRevisionIds> {
    Ok(SectionRevisionIds {
        r: xml::optional_rsid(e, b"rsidR")?,
        r_pr: xml::optional_rsid(e, b"rsidRPr")?,
        sect: xml::optional_rsid(e, b"rsidSect")?,
    })
}

pub fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<SectionProperties> {
    let mut props = SectionProperties::default();

    loop {
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"pgSz" => {
                        let orientation = match xml::optional_attr(e, b"orient")?.as_deref() {
                            Some("landscape") => Some(PageOrientation::Landscape),
                            Some("portrait") => Some(PageOrientation::Portrait),
                            Some(other) => return Err(invalid_value("orient", other)),
                            None => None,
                        };
                        props.page_size = Some(PageSize {
                            width: xml::optional_attr_i64(e, b"w")?.map(Dimension::new),
                            height: xml::optional_attr_i64(e, b"h")?.map(Dimension::new),
                            orientation,
                        });
                    }
                    b"pgMar" => {
                        props.page_margins = Some(PageMargins {
                            top: xml::optional_attr_i64(e, b"top")?.map(Dimension::new),
                            right: xml::optional_attr_i64(e, b"right")?.map(Dimension::new),
                            bottom: xml::optional_attr_i64(e, b"bottom")?.map(Dimension::new),
                            left: xml::optional_attr_i64(e, b"left")?.map(Dimension::new),
                            header: xml::optional_attr_i64(e, b"header")?.map(Dimension::new),
                            footer: xml::optional_attr_i64(e, b"footer")?.map(Dimension::new),
                            gutter: xml::optional_attr_i64(e, b"gutter")?.map(Dimension::new),
                        });
                    }
                    b"cols" => {
                        let columns = if is_start {
                            parse_column_definitions(reader, buf)?
                        } else {
                            Vec::new()
                        };
                        props.columns = Some(Columns {
                            count: xml::optional_attr_u32(e, b"num")?,
                            space: xml::optional_attr_i64(e, b"space")?.map(Dimension::new),
                            equal_width: xml::optional_attr_bool(e, b"equalWidth")?,
                            columns,
                        });
                    }
                    b"headerReference" => {
                        if let Some(r_id) = xml::optional_attr(e, b"id")? {
                            let hf_type = xml::optional_attr(e, b"type")?;
                            let rel = RelId::new(r_id);
                            match hf_type.as_deref() {
                                Some("first") => props.header_refs.first = Some(rel),
                                Some("even") => props.header_refs.even = Some(rel),
                                _ => props.header_refs.default = Some(rel),
                            }
                        }
                    }
                    b"footerReference" => {
                        if let Some(r_id) = xml::optional_attr(e, b"id")? {
                            let hf_type = xml::optional_attr(e, b"type")?;
                            let rel = RelId::new(r_id);
                            match hf_type.as_deref() {
                                Some("first") => props.footer_refs.first = Some(rel),
                                Some("even") => props.footer_refs.even = Some(rel),
                                _ => props.footer_refs.default = Some(rel),
                            }
                        }
                    }
                    b"docGrid" => {
                        let grid_type = match xml::optional_attr(e, b"type")?.as_deref() {
                            Some("default") => Some(DocGridType::Default),
                            Some("lines") => Some(DocGridType::Lines),
                            Some("linesAndChars") => Some(DocGridType::LinesAndChars),
                            Some("snapToChars") => Some(DocGridType::SnapToChars),
                            Some(other) => return Err(invalid_value("docGrid/type", other)),
                            None => None,
                        };
                        props.doc_grid = Some(DocGrid {
                            grid_type,
                            line_pitch: xml::optional_attr_i64(e, b"linePitch")?
                                .map(Dimension::new),
                            char_space: xml::optional_attr_i64(e, b"charSpace")?
                                .map(Dimension::new),
                        });
                    }
                    b"titlePg" => {
                        props.title_page = toggle_attr(e)?;
                    }
                    b"pgNumType" => {
                        props.page_number_type = Some(parse_page_number_type(e)?);
                    }
                    b"type" => {
                        props.section_type = opt_val(e, parse_section_type)?;
                    }
                    _ => xml::warn_unsupported_element("sectPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"sectPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"sectPr")),
            _ => {}
        }
    }

    Ok(props)
}

/// §17.18.77 ST_SectionMark
fn parse_section_type(val: &str) -> Result<SectionType> {
    match val {
        "nextPage" => Ok(SectionType::NextPage),
        "continuous" => Ok(SectionType::Continuous),
        "evenPage" => Ok(SectionType::EvenPage),
        "oddPage" => Ok(SectionType::OddPage),
        "nextColumn" => Ok(SectionType::NextColumn),
        other => Err(invalid_value("type/val", other)),
    }
}

/// §17.6.3: parse `w:col` children inside `w:cols`.
fn parse_column_definitions(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Vec<ColumnDefinition>> {
    let mut cols = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"col" => {
                        cols.push(ColumnDefinition {
                            width: xml::optional_attr_i64(e, b"w")?.map(Dimension::new),
                            space: xml::optional_attr_i64(e, b"space")?.map(Dimension::new),
                        });
                    }
                    _ => xml::warn_unsupported_element("cols", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"cols" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"cols")),
            _ => {}
        }
    }

    Ok(cols)
}

/// §17.6.12: parse `w:pgNumType` attributes.
fn parse_page_number_type(e: &BytesStart<'_>) -> Result<PageNumberType> {
    Ok(PageNumberType {
        format: match xml::optional_attr(e, b"fmt")?.as_deref() {
            Some("decimal") => Some(NumberFormat::Decimal),
            Some("upperRoman") => Some(NumberFormat::UpperRoman),
            Some("lowerRoman") => Some(NumberFormat::LowerRoman),
            Some("upperLetter") => Some(NumberFormat::UpperLetter),
            Some("lowerLetter") => Some(NumberFormat::LowerLetter),
            Some("ordinal") => Some(NumberFormat::Ordinal),
            Some("cardinalText") => Some(NumberFormat::CardinalText),
            Some("ordinalText") => Some(NumberFormat::OrdinalText),
            Some("none") => Some(NumberFormat::None),
            Some(other) => return Err(invalid_value("pgNumType/fmt", other)),
            None => None,
        },
        start: xml::optional_attr_u32(e, b"start")?,
        chap_style: xml::optional_attr_u32(e, b"chapStyle")?,
        chap_sep: match xml::optional_attr(e, b"chapSep")?.as_deref() {
            Some("hyphen") => Some(ChapterSeparator::Hyphen),
            Some("period") => Some(ChapterSeparator::Period),
            Some("colon") => Some(ChapterSeparator::Colon),
            Some("emDash") => Some(ChapterSeparator::EmDash),
            Some("enDash") => Some(ChapterSeparator::EnDash),
            Some(other) => return Err(invalid_value("pgNumType/chapSep", other)),
            None => None,
        },
    })
}
