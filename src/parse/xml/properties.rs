use crate::error::Error;
use crate::model::*;
use crate::units::{UNDERLINE_NONE, WIDTH_TYPE_DXA};

use super::ParseState;
use super::helpers::{get_attr, is_val_false};

/// Parse a margin value from an element like `<w:top w:w="0" w:type="dxa"/>`.
fn parse_margin_value(e: &quick_xml::events::BytesStart<'_>) -> Result<Option<u32>, Error> {
    let w_type = get_attr(e, b"type")?.unwrap_or_default();
    if w_type == "dxa" {
        if let Some(val) = get_attr(e, b"w")? {
            return Ok(val.parse::<u32>().ok());
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
pub fn handle_empty_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
) -> Result<(), Error> {
    match state {
        ParseState::InRunProperties { ref mut props } => {
            match local {
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
                        props.font_size = val.parse::<u32>().ok();
                    }
                }
                b"rFonts" => {
                    if let Some(val) = get_attr(e, b"ascii")? {
                        props.font_family = Some(val);
                    } else if let Some(val) = get_attr(e, b"hAnsi")? {
                        props.font_family = Some(val);
                    }
                }
                b"color" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.color = Color::from_hex(&val);
                    }
                }
                _ => {}
            }
        }
        ParseState::InParagraphProperties { ref mut props, .. } => {
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
                        spacing.before = val.parse().ok();
                    }
                    if let Some(val) = get_attr(e, b"after")? {
                        spacing.after = val.parse().ok();
                    }
                    if let Some(val) = get_attr(e, b"line")? {
                        spacing.line = val.parse().ok();
                    }
                    props.spacing = Some(spacing);
                }
                b"ind" => {
                    let mut indent = Indentation::default();
                    if let Some(val) = get_attr(e, b"left")? {
                        indent.left = val.parse().ok();
                    }
                    if let Some(val) = get_attr(e, b"right")? {
                        indent.right = val.parse().ok();
                    }
                    if let Some(val) = get_attr(e, b"firstLine")? {
                        indent.first_line = val.parse().ok();
                    }
                    if let Some(val) = get_attr(e, b"hanging")? {
                        if let Ok(v) = val.parse::<i32>() {
                            indent.first_line = Some(-v);
                        }
                    }
                    props.indentation = Some(indent);
                }
                b"tab" => {
                    if let (Some(val), Some(pos)) =
                        (get_attr(e, b"val")?, get_attr(e, b"pos")?)
                    {
                        let stop_type = match val.as_str() {
                            "left" => Some(TabStopType::Left),
                            "center" => Some(TabStopType::Center),
                            "right" => Some(TabStopType::Right),
                            "decimal" => Some(TabStopType::Decimal),
                            "clear" => None,
                            _ => Some(TabStopType::Left),
                        };
                        if let (Some(st), Ok(p)) = (stop_type, pos.parse::<u32>()) {
                            props.tab_stops.push(TabStop {
                                position: p,
                                stop_type: st,
                            });
                        }
                    }
                }
                _ => {}
            }
        }
        ParseState::InTable {
            ref mut grid_cols,
            ref mut default_cell_margins,
            in_cell_mar,
            ..
        } => {
            if *in_cell_mar {
                apply_margin(default_cell_margins, local, e)?;
            } else if local == b"gridCol" {
                if let Some(val) = get_attr(e, b"w")? {
                    if let Ok(w) = val.parse::<u32>() {
                        grid_cols.push(w);
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
            ..
        } => {
            if *in_cell_mar {
                apply_margin(cell_margins, local, e)?;
                return Ok(());
            }
            match local {
                b"tcW" => {
                    let w_type = get_attr(e, b"type")?.unwrap_or_default();
                    if w_type == WIDTH_TYPE_DXA {
                        if let Some(val) = get_attr(e, b"w")? {
                            if let Ok(w) = val.parse::<u32>() {
                                *width = Some(w);
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
                _ => {}
            }
        }
        _ => {}
    }

    Ok(())
}
