use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

use super::{
    invalid_value, opt_val, parse_alignment, parse_border, parse_cnf_style,
    parse_edge_insets_twips, parse_shading, toggle_attr,
};

/// Parse `w:tblPr`. Returns (properties, optional style ID).
pub fn parse_table_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(TableProperties, Option<StyleId>)> {
    let mut props = TableProperties::default();
    let mut style_id: Option<StyleId> = None;

    loop {
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tblStyle" => {
                        props.style_id = xml::optional_attr(e, b"val")?.map(StyleId::new);
                        style_id.clone_from(&props.style_id);
                    }
                    // Start-only: have child elements
                    b"tblBorders" if is_start => {
                        props.borders = Some(parse_table_borders(reader, buf)?);
                    }
                    b"tblCellMar" if is_start => {
                        props.cell_margins =
                            Some(parse_edge_insets_twips(reader, buf, b"tblCellMar")?);
                    }
                    // attrs-only — valid as Start or Empty
                    b"jc" => {
                        props.alignment = opt_val(e, parse_alignment)?;
                    }
                    b"tblW" => {
                        props.width = Some(parse_table_measure(e)?);
                    }
                    b"tblLayout" => {
                        props.layout = opt_val(e, |v| {
                            Ok(match v {
                                "fixed" => TableLayout::Fixed,
                                "autofit" | "auto" => TableLayout::Auto,
                                other => return Err(invalid_value("tblLayout/val", other)),
                            })
                        })?;
                    }
                    b"tblInd" => {
                        props.indent = Some(parse_table_measure(e)?);
                    }
                    b"tblCellSpacing" => {
                        props.cell_spacing = Some(parse_table_measure(e)?);
                    }
                    b"tblLook" => {
                        props.look = Some(parse_table_look(e)?);
                    }
                    b"tblStyleRowBandSize" => {
                        props.style_row_band_size = xml::optional_attr_u32(e, b"val")?;
                    }
                    b"tblStyleColBandSize" => {
                        props.style_col_band_size = xml::optional_attr_u32(e, b"val")?;
                    }
                    b"tblpPr" => {
                        props.positioning = Some(parse_table_positioning(e)?);
                    }
                    b"tblOverlap" => {
                        props.overlap = opt_val(e, parse_table_overlap)?;
                    }
                    _ => xml::warn_unsupported_element("tblPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tblPr")),
            _ => {}
        }
    }

    Ok((props, style_id))
}

/// Parse `w:trPr`.
pub fn parse_table_row_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<TableRowProperties> {
    let mut props = TableRowProperties::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"trHeight" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            let rule = match xml::optional_attr(e, b"hRule")?.as_deref() {
                                Some("exact") => HeightRule::Exact,
                                Some("atLeast") => HeightRule::AtLeast,
                                _ => HeightRule::Auto,
                            };
                            props.height = Some(TableRowHeight {
                                value: Dimension::new(val),
                                rule,
                            });
                        }
                    }
                    b"tblHeader" => {
                        props.is_header = toggle_attr(e)?;
                    }
                    b"cantSplit" => {
                        props.cant_split = toggle_attr(e)?;
                    }
                    b"jc" => {
                        props.justification = opt_val(e, parse_alignment)?;
                    }
                    b"cnfStyle" => {
                        props.cnf_style = Some(parse_cnf_style(e)?);
                    }
                    b"gridAfter" => {
                        props.grid_after = xml::optional_attr_u32(e, b"val")?;
                    }
                    b"wAfter" => {
                        props.w_after = Some(parse_table_measure(e)?);
                    }
                    _ => xml::warn_unsupported_element("trPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"trPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"trPr")),
            _ => {}
        }
    }

    Ok(props)
}

/// Parse `w:tcPr`.
pub fn parse_table_cell_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<TableCellProperties> {
    let mut props = TableCellProperties::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tcBorders" => {
                        props.borders = Some(parse_table_cell_borders(reader, buf)?);
                    }
                    b"tcMar" => {
                        props.margins = Some(parse_edge_insets_twips(reader, buf, b"tcMar")?);
                    }
                    _ => xml::warn_unsupported_element("tcPr", local),
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tcW" => {
                        props.width = Some(parse_table_measure(e)?);
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"vAlign" => {
                        props.vertical_align = opt_val(e, parse_cell_vertical_align)?;
                    }
                    b"vMerge" => {
                        let is_restart = xml::optional_attr(e, b"val")?
                            .map(|v| v == "restart")
                            .unwrap_or(false);
                        props.vertical_merge = Some(if is_restart {
                            VerticalMerge::Restart
                        } else {
                            VerticalMerge::Continue
                        });
                    }
                    b"gridSpan" => {
                        props.grid_span = xml::optional_attr_u32(e, b"val")?;
                    }
                    b"textDirection" => {
                        props.text_direction = opt_val(e, parse_text_direction)?;
                    }
                    b"noWrap" => {
                        props.no_wrap = toggle_attr(e)?;
                    }
                    b"cnfStyle" => {
                        props.cnf_style = Some(parse_cnf_style(e)?);
                    }
                    _ => xml::warn_unsupported_element("tcPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tcPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tcPr")),
            _ => {}
        }
    }

    Ok(props)
}

/// §17.18.96 ST_TblWidth
fn parse_table_measure(e: &BytesStart<'_>) -> Result<TableMeasure> {
    let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
    let measure_type = xml::optional_attr(e, b"type")?;

    match measure_type.as_deref() {
        Some("dxa") => Ok(TableMeasure::Twips(Dimension::new(w))),
        Some("pct") => Ok(TableMeasure::Pct(Dimension::new(w))),
        Some("nil") => Ok(TableMeasure::Nil),
        Some("auto") | None => Ok(TableMeasure::Auto),
        Some(other) => Err(invalid_value("type", other)),
    }
}

/// §17.4.56: parse tblLook. Supports both individual attributes (OOXML 2010+)
/// and the legacy `val` bit field format.
fn parse_table_look(e: &BytesStart<'_>) -> Result<TableLook> {
    let first_row = xml::optional_attr_bool(e, b"firstRow")?;
    let last_row = xml::optional_attr_bool(e, b"lastRow")?;
    let first_column = xml::optional_attr_bool(e, b"firstColumn")?;
    let last_column = xml::optional_attr_bool(e, b"lastColumn")?;
    let no_h_band = xml::optional_attr_bool(e, b"noHBand")?;
    let no_v_band = xml::optional_attr_bool(e, b"noVBand")?;

    // If individual attributes are present, use them directly.
    if first_row.is_some()
        || last_row.is_some()
        || first_column.is_some()
        || last_column.is_some()
        || no_h_band.is_some()
        || no_v_band.is_some()
    {
        return Ok(TableLook {
            first_row,
            last_row,
            first_column,
            last_column,
            no_h_band,
            no_v_band,
        });
    }

    // §17.4.56: legacy `val` attribute is a hex bit field.
    // Bit 0x0020: firstRow
    // Bit 0x0040: lastRow
    // Bit 0x0080: firstColumn
    // Bit 0x0100: lastColumn
    // Bit 0x0200: noHBand (horizontal banding disabled)
    // Bit 0x0400: noVBand (vertical banding disabled)
    if let Some(val_str) = xml::optional_attr(e, b"val")? {
        let val = u32::from_str_radix(&val_str, 16).unwrap_or(0);
        return Ok(TableLook {
            first_row: Some(val & 0x0020 != 0),
            last_row: Some(val & 0x0040 != 0),
            first_column: Some(val & 0x0080 != 0),
            last_column: Some(val & 0x0100 != 0),
            no_h_band: Some(val & 0x0200 != 0),
            no_v_band: Some(val & 0x0400 != 0),
        });
    }

    Ok(TableLook::default())
}

fn parse_table_borders(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<TableBorders> {
    let mut borders = TableBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
        inside_h: None,
        inside_v: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let border = parse_border(e)?;
                match local {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"insideH" => borders.inside_h = Some(border),
                    b"insideV" => borders.inside_v = Some(border),
                    _ => xml::warn_unsupported_element("tblBorders", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblBorders" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tblBorders")),
            _ => {}
        }
    }

    Ok(borders)
}

fn parse_table_cell_borders(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<TableCellBorders> {
    let mut borders = TableCellBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
        inside_h: None,
        inside_v: None,
        tl2br: None,
        tr2bl: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let border = parse_border(e)?;
                match local {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"insideH" => borders.inside_h = Some(border),
                    b"insideV" => borders.inside_v = Some(border),
                    b"tl2br" => borders.tl2br = Some(border),
                    b"tr2bl" => borders.tr2bl = Some(border),
                    _ => xml::warn_unsupported_element("tcBorders", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tcBorders" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tcBorders")),
            _ => {}
        }
    }

    Ok(borders)
}

/// §17.18.101 ST_VerticalJc
fn parse_cell_vertical_align(val: &str) -> Result<CellVerticalAlign> {
    match val {
        "top" => Ok(CellVerticalAlign::Top),
        "center" => Ok(CellVerticalAlign::Center),
        "bottom" => Ok(CellVerticalAlign::Bottom),
        "both" => Ok(CellVerticalAlign::Both),
        other => Err(invalid_value("vAlign/val", other)),
    }
}

/// §17.18.93 ST_TextDirection
fn parse_text_direction(val: &str) -> Result<TextDirection> {
    match val {
        "lrTb" => Ok(TextDirection::LeftToRightTopToBottom),
        "tbRl" => Ok(TextDirection::TopToBottomRightToLeft),
        "btLr" => Ok(TextDirection::BottomToTopLeftToRight),
        "lrTbV" => Ok(TextDirection::LeftToRightTopToBottomRotated),
        "tbRlV" => Ok(TextDirection::TopToBottomRightToLeftRotated),
        "tbLrV" => Ok(TextDirection::TopToBottomLeftToRightRotated),
        other => Err(invalid_value("textDirection/val", other)),
    }
}

/// §17.4.58: parse `w:tblpPr` attributes.
fn parse_table_positioning(e: &BytesStart<'_>) -> Result<TablePositioning> {
    Ok(TablePositioning {
        left_from_text: xml::optional_attr_i64(e, b"leftFromText")?.map(Dimension::new),
        right_from_text: xml::optional_attr_i64(e, b"rightFromText")?.map(Dimension::new),
        top_from_text: xml::optional_attr_i64(e, b"topFromText")?.map(Dimension::new),
        bottom_from_text: xml::optional_attr_i64(e, b"bottomFromText")?.map(Dimension::new),
        vert_anchor: match xml::optional_attr(e, b"vertAnchor")?.as_deref() {
            Some("text") => Some(TableAnchor::Text),
            Some("margin") => Some(TableAnchor::Margin),
            Some("page") => Some(TableAnchor::Page),
            Some(other) => return Err(invalid_value("tblpPr/vertAnchor", other)),
            None => None,
        },
        horz_anchor: match xml::optional_attr(e, b"horzAnchor")?.as_deref() {
            Some("text") => Some(TableAnchor::Text),
            Some("margin") => Some(TableAnchor::Margin),
            Some("page") => Some(TableAnchor::Page),
            Some(other) => return Err(invalid_value("tblpPr/horzAnchor", other)),
            None => None,
        },
        x_align: match xml::optional_attr(e, b"tblpXSpec")?.as_deref() {
            Some("left") => Some(TableXAlign::Left),
            Some("center") => Some(TableXAlign::Center),
            Some("right") => Some(TableXAlign::Right),
            Some("inside") => Some(TableXAlign::Inside),
            Some("outside") => Some(TableXAlign::Outside),
            Some(other) => return Err(invalid_value("tblpPr/tblpXSpec", other)),
            None => None,
        },
        y_align: match xml::optional_attr(e, b"tblpYSpec")?.as_deref() {
            Some("top") => Some(TableYAlign::Top),
            Some("center") => Some(TableYAlign::Center),
            Some("bottom") => Some(TableYAlign::Bottom),
            Some("inside") => Some(TableYAlign::Inside),
            Some("outside") => Some(TableYAlign::Outside),
            Some("inline") => Some(TableYAlign::Inline),
            Some(other) => return Err(invalid_value("tblpPr/tblpYSpec", other)),
            None => None,
        },
        x: xml::optional_attr_i64(e, b"tblpX")?.map(Dimension::new),
        y: xml::optional_attr_i64(e, b"tblpY")?.map(Dimension::new),
    })
}

/// §17.4.56 ST_TblOverlap
fn parse_table_overlap(val: &str) -> Result<TableOverlap> {
    match val {
        "overlap" => Ok(TableOverlap::Overlap),
        "never" => Ok(TableOverlap::Never),
        other => Err(invalid_value("tblOverlap/val", other)),
    }
}
