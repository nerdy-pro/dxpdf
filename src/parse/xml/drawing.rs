use crate::error::Error;
use crate::model::*;

use super::helpers::get_attr;
use super::ParseState;

/// Handle elements inside a `w:drawing` subtree to extract image info.
pub fn handle_drawing_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
) -> Result<(), Error> {
    if let ParseState::InDrawing(ds) = state {
        match local {
            b"anchor" => {
                ds.is_anchor = true;
            }
            b"inline" => {}
            b"extent" => {
                if let Some(cx) = get_attr(e, b"cx")? {
                    ds.width_emu = cx.parse::<i64>().ok().map(crate::dimension::Emu::new);
                }
                if let Some(cy) = get_attr(e, b"cy")? {
                    ds.height_emu = cy.parse::<i64>().ok().map(crate::dimension::Emu::new);
                }
            }
            b"blip" => {
                if let Some(embed) = get_attr(e, b"embed")? {
                    ds.rel_id = Some(RelId::from(embed));
                }
            }
            b"positionH" => {
                ds.in_position_h = true;
            }
            b"positionV" => {
                ds.in_position_v = true;
            }
            b"posOffset" => {
                if ds.in_position_h {
                    ds.reading_pos_offset = Some('H');
                } else if ds.in_position_v {
                    ds.reading_pos_offset = Some('V');
                }
            }
            b"align" => {
                // wp:align contains text like "left", "right", "center"
                if ds.in_position_h {
                    ds.reading_align = Some('H');
                } else if ds.in_position_v {
                    ds.reading_align = Some('V');
                }
            }
            b"pctPosHOffset" => {
                if ds.in_position_h {
                    ds.reading_pct_pos = Some('H');
                }
            }
            b"pctPosVOffset" => {
                if ds.in_position_v {
                    ds.reading_pct_pos = Some('V');
                }
            }
            b"wrapTight" | b"wrapSquare" | b"wrapThrough" => {
                if let Some(val) = get_attr(e, b"wrapText")? {
                    ds.wrap_side = match val.as_str() {
                        "bothSides" => Some(WrapSide::BothSides),
                        "left" => Some(WrapSide::Left),
                        "right" => Some(WrapSide::Right),
                        _ => Some(WrapSide::BothSides),
                    };
                } else {
                    ds.wrap_side = Some(WrapSide::BothSides);
                }
            }
            b"wrapNone" => {}
            _ => {}
        }
    }
    Ok(())
}

/// Handle End events inside a drawing subtree for position tracking.
pub fn handle_drawing_end(local: &[u8], state: &mut ParseState) {
    if let ParseState::InDrawing(ds) = state {
        match local {
            b"positionH" => {
                ds.in_position_h = false;
                if ds.reading_pos_offset == Some('H') {
                    ds.reading_pos_offset = None;
                }
                if ds.reading_pct_pos == Some('H') {
                    ds.reading_pct_pos = None;
                }
            }
            b"positionV" => {
                ds.in_position_v = false;
                if ds.reading_pos_offset == Some('V') {
                    ds.reading_pos_offset = None;
                }
                if ds.reading_pct_pos == Some('V') {
                    ds.reading_pct_pos = None;
                }
            }
            b"posOffset" => {
                ds.reading_pos_offset = None;
            }
            b"pctPosHOffset" | b"pctPosVOffset" => {
                ds.reading_pct_pos = None;
            }
            b"align" => {
                ds.reading_align = None;
            }
            _ => {}
        }
    }
}
