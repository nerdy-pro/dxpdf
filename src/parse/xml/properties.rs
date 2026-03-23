use log::warn;

use super::helpers::{get_attr, is_val_false};
use crate::dimension::EighthPoints;
use crate::error::Error;
use crate::model::*;

/// OOXML underline value for "no underline".
const UNDERLINE_NONE: &str = "none";
/// OOXML width type for twips.
const WIDTH_TYPE_DXA: &str = "dxa";
use super::ParseState;

/// Parse a border element like `<w:top w:val="single" w:sz="4" w:color="auto"/>`.
pub fn parse_border_def(e: &quick_xml::events::BytesStart<'_>) -> Result<BorderDef, Error> {
    let val = get_attr(e, b"val")?.unwrap_or_default();
    let style = match val.as_str() {
        "none" | "nil" => BorderStyle::None,
        "single" => BorderStyle::Single,
        "double" => BorderStyle::Double,
        "dashed" | "dashSmallGap" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        _ => BorderStyle::Single, // treat unknown styles as single
    };

    let size = get_attr(e, b"sz")?
        .and_then(|v| v.parse::<i64>().ok())
        .map(EighthPoints::new)
        .unwrap_or(BorderDef::DEFAULT_SIZE);

    let color_str = get_attr(e, b"color")?.unwrap_or_default();
    let color = if color_str == "auto" || color_str.is_empty() {
        Color::BLACK
    } else {
        Color::from_hex(&color_str).unwrap_or(Color::BLACK)
    };

    let space = get_attr(e, b"space")?
        .and_then(|v| v.parse::<f32>().ok())
        .map(crate::dimension::Pt::new)
        .unwrap_or(crate::dimension::Pt::ZERO);

    Ok(BorderDef {
        style,
        size,
        color,
        space,
    })
}

/// Apply a border element to the appropriate side of a TableBorders or CellBorders.
pub fn apply_table_border(
    borders: &mut TableBorders,
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), Error> {
    let def = parse_border_def(e)?;
    match local {
        b"top" => borders.top = def,
        b"bottom" => borders.bottom = def,
        b"left" | b"start" => borders.left = def,
        b"right" | b"end" => borders.right = def,
        b"insideH" => borders.inside_h = def,
        b"insideV" => borders.inside_v = def,
        _ => {}
    }
    Ok(())
}

pub fn apply_cell_border(
    borders: &mut CellBorders,
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), Error> {
    let def = parse_border_def(e)?;
    match local {
        b"top" => borders.top = Some(def),
        b"bottom" => borders.bottom = Some(def),
        b"left" | b"start" => borders.left = Some(def),
        b"right" | b"end" => borders.right = Some(def),
        _ => {}
    }
    Ok(())
}

/// Parse a margin value from an element like `<w:top w:w="0" w:type="dxa"/>`.
fn parse_margin_value(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<crate::dimension::Twips>, Error> {
    let w_type = get_attr(e, b"type")?.unwrap_or_default();
    if w_type == "dxa" {
        if let Some(val) = get_attr(e, b"w")? {
            return Ok(val.parse::<i64>().ok().map(crate::dimension::Twips::new));
        }
    }
    Ok(None)
}

/// Apply a margin element to a CellMargins struct, creating it if needed.
fn apply_margin(
    margins: &mut Option<CellMargins>,
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), Error> {
    if let Some(val) = parse_margin_value(e)? {
        let m = margins.get_or_insert(CellMargins::default());
        match local {
            b"top" => m.top = val,
            b"bottom" => m.bottom = val,
            b"left" | b"start" => m.left = val,
            b"right" | b"end" => m.right = val,
            _ => {}
        }
    }
    Ok(())
}

/// Handle empty elements that set properties on runs, paragraphs, tables, and cells.
pub(super) fn handle_empty_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
    warned: &mut std::collections::HashSet<&'static str>,
) -> Result<(), Error> {
    match state {
        ParseState::InRunProperties { ref mut props } => {
            match local {
                b"rStyle" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.style_id = Some(val);
                    }
                }
                b"b" => {
                    props.bold = !is_val_false(e)?;
                }
                b"i" => {
                    props.italic = !is_val_false(e)?;
                }
                b"u" => {
                    let val = get_attr(e, b"val")?;
                    props.underline = val.as_deref() != Some(UNDERLINE_NONE);
                }
                b"sz" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.font_size = val
                            .parse::<i64>()
                            .ok()
                            .map(crate::dimension::HalfPoints::new);
                    }
                }
                b"rFonts" => {
                    if let Some(val) = get_attr(e, b"ascii")? {
                        props.font_family = Some(std::rc::Rc::from(val.as_str()));
                    } else if let Some(val) = get_attr(e, b"hAnsi")? {
                        props.font_family = Some(std::rc::Rc::from(val.as_str()));
                    }
                }
                b"color" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.color = Color::from_hex(&val);
                    }
                }
                b"shd" => {
                    if let Some(fill) = get_attr(e, b"fill")? {
                        if fill != "auto" && !fill.is_empty() {
                            props.shading = Color::from_hex(&fill);
                        }
                    }
                }
                b"spacing" => {
                    // w:spacing in rPr = character spacing (in twips)
                    if let Some(val) = get_attr(e, b"val")? {
                        props.char_spacing =
                            val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                }
                b"strike" | b"dstrike" => {
                    if warned.insert("strike") {
                        warn!("Unsupported: strikethrough (w:strike/w:dstrike)");
                    }
                }
                b"vertAlign" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.vert_align = match val.as_str() {
                            "superscript" => Some(VertAlign::Superscript),
                            "subscript" => Some(VertAlign::Subscript),
                            _ => None,
                        };
                    }
                }
                b"caps" | b"smallCaps" => {
                    if warned.insert("caps") {
                        warn!("Unsupported: capitalization effect (w:caps/w:smallCaps)");
                    }
                }
                b"highlight" => {
                    if warned.insert("highlight") {
                        warn!("Unsupported: text highlighting (w:highlight)");
                    }
                }
                _ => {}
            }
        }
        ParseState::InParagraphProperties {
            ref mut props,
            ref in_pbdr,
            ..
        } => {
            if *in_pbdr {
                // Inside w:pBdr — parse border edge elements
                let border = parse_border_def(e)?;
                let borders = props
                    .paragraph_borders
                    .get_or_insert(ParagraphBorders::default());
                match local {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" => borders.left = Some(border),
                    b"right" => borders.right = Some(border),
                    _ => {}
                }
                return Ok(());
            }
            match local {
                b"jc" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.alignment = match val.as_str() {
                            "left" | "start" => Some(Alignment::Left),
                            "center" => Some(Alignment::Center),
                            "right" | "end" => Some(Alignment::Right),
                            "both" | "justify" => Some(Alignment::Justify),
                            _ => None,
                        };
                    }
                }
                b"spacing" => {
                    let mut spacing = Spacing::default();
                    if let Some(val) = get_attr(e, b"before")? {
                        spacing.before = val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                    if let Some(val) = get_attr(e, b"after")? {
                        spacing.after = val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                    if let Some(val) = get_attr(e, b"line")? {
                        spacing.line = val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                    if let Some(val) = get_attr(e, b"lineRule")? {
                        spacing.line_rule = match val.as_str() {
                            "auto" => LineRule::Auto,
                            "exact" => LineRule::Exact,
                            "atLeast" => LineRule::AtLeast,
                            _ => LineRule::Auto,
                        };
                    }
                    props.spacing = Some(spacing);
                }
                b"ind" => {
                    let mut indent = Indentation::default();
                    if let Some(val) = get_attr(e, b"left")? {
                        indent.left = val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                    if let Some(val) = get_attr(e, b"right")? {
                        indent.right = val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                    if let Some(val) = get_attr(e, b"firstLine")? {
                        indent.first_line =
                            val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                    }
                    if let Some(val) = get_attr(e, b"hanging")? {
                        if let Ok(v) = val.parse::<i64>() {
                            indent.first_line = Some(crate::dimension::Twips::new(-v));
                        }
                    }
                    props.indentation = Some(indent);
                }
                b"tab" => {
                    if let (Some(val), Some(pos)) = (get_attr(e, b"val")?, get_attr(e, b"pos")?) {
                        let stop_type = match val.as_str() {
                            "left" => Some(TabStopType::Left),
                            "center" => Some(TabStopType::Center),
                            "right" => Some(TabStopType::Right),
                            "decimal" => Some(TabStopType::Decimal),
                            "clear" => None,
                            _ => Some(TabStopType::Left),
                        };
                        if let (Some(st), Ok(p)) = (stop_type, pos.parse::<i64>()) {
                            props.tab_stops.push(TabStop {
                                position: crate::dimension::Twips::new(p),
                                stop_type: st,
                            });
                        }
                    }
                }
                b"shd" => {
                    if let Some(fill) = get_attr(e, b"fill")? {
                        if fill != "auto" && !fill.is_empty() {
                            props.shading = Color::from_hex(&fill);
                        }
                    }
                }
                b"pStyle" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.style_id = Some(val);
                    }
                }
                b"ilvl" => {
                    // Part of w:numPr — set level on existing list ref
                    if let Some(val) = get_attr(e, b"val")? {
                        if let Ok(lvl) = val.parse::<u32>() {
                            let lr = props.list_ref.get_or_insert(ListRef {
                                num_id: 0,
                                level: 0,
                            });
                            lr.level = lvl;
                        }
                    }
                }
                b"numId" => {
                    // Part of w:numPr — set numId on existing list ref
                    if let Some(val) = get_attr(e, b"val")? {
                        if let Ok(nid) = val.parse::<u32>() {
                            let lr = props.list_ref.get_or_insert(ListRef {
                                num_id: 0,
                                level: 0,
                            });
                            lr.num_id = nid;
                        }
                    }
                }
                _ => {}
            }
        }
        ParseState::InTableRow { ref mut height, .. } => {
            if local == b"trHeight" {
                if let Some(val) = get_attr(e, b"val")? {
                    *height = val.parse::<i64>().ok().map(crate::dimension::Twips::new);
                }
            }
        }
        ParseState::InTable {
            ref mut grid_cols,
            ref mut default_cell_margins,
            ref in_cell_mar,
            ref mut borders,
            ref in_borders,
            ..
        } => {
            if *in_borders {
                let b = borders.get_or_insert(TableBorders::default());
                apply_table_border(b, local, e)?;
                return Ok(());
            }
            if *in_cell_mar {
                apply_margin(default_cell_margins, local, e)?;
            } else if local == b"gridCol" {
                if let Some(val) = get_attr(e, b"w")? {
                    if let Ok(w) = val.parse::<i64>() {
                        grid_cols.push(crate::dimension::Twips::new(w));
                    }
                }
            }
        }
        ParseState::InTableCell {
            ref mut width,
            ref mut grid_span,
            ref mut vertical_merge,
            ref mut cell_margins,
            ref in_cell_mar,
            ref mut cell_borders,
            ref in_borders,
            ref mut shading,
            ..
        } => {
            if *in_borders {
                let b = cell_borders.get_or_insert(CellBorders::default());
                apply_cell_border(b, local, e)?;
                return Ok(());
            }
            if *in_cell_mar {
                apply_margin(cell_margins, local, e)?;
                return Ok(());
            }
            match local {
                b"tcW" => {
                    let w_type = get_attr(e, b"type")?.unwrap_or_default();
                    if w_type == WIDTH_TYPE_DXA {
                        if let Some(val) = get_attr(e, b"w")? {
                            if let Ok(w) = val.parse::<i64>() {
                                *width = Some(crate::dimension::Twips::new(w));
                            }
                        }
                    }
                }
                b"gridSpan" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        if let Ok(gs) = val.parse::<u32>() {
                            *grid_span = gs;
                        }
                    }
                }
                b"vMerge" => {
                    let val = get_attr(e, b"val")?.unwrap_or_default();
                    *vertical_merge = if val == "restart" {
                        Some(VerticalMerge::Restart)
                    } else {
                        Some(VerticalMerge::Continue)
                    };
                }
                b"shd" => {
                    if let Some(fill) = get_attr(e, b"fill")? {
                        if fill != "auto" && !fill.is_empty() {
                            *shading = Color::from_hex(&fill);
                        }
                    }
                }
                _ => {}
            }
        }
        _ => {}
    }

    Ok(())
}
