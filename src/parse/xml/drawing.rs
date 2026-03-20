use crate::error::Error;
use crate::model::*;

use super::ParseState;
use super::helpers::get_attr;

/// Handle elements inside a `w:drawing` subtree to extract image info.
pub fn handle_drawing_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
) -> Result<(), Error> {
    if let ParseState::InDrawing {
        ref mut rel_id,
        ref mut width_emu,
        ref mut height_emu,
        ref mut is_anchor,
        ref mut wrap_side,
        ref mut in_position_h,
        ref mut in_position_v,
        ref mut reading_pos_offset,
        ref mut reading_align,
        ..
    } = state
    {
        match local {
            b"anchor" => {
                *is_anchor = true;
            }
            b"inline" => {}
            b"extent" => {
                if let Some(cx) = get_attr(e, b"cx")? {
                    *width_emu = cx.parse().ok();
                }
                if let Some(cy) = get_attr(e, b"cy")? {
                    *height_emu = cy.parse().ok();
                }
            }
            b"blip" => {
                if let Some(embed) = get_attr(e, b"embed")? {
                    *rel_id = Some(RelId::from(embed));
                }
            }
            b"positionH" => {
                *in_position_h = true;
            }
            b"positionV" => {
                *in_position_v = true;
            }
            b"posOffset" => {
                if *in_position_h {
                    *reading_pos_offset = Some('H');
                } else if *in_position_v {
                    *reading_pos_offset = Some('V');
                }
            }
            b"align" => {
                // wp:align contains text like "left", "right", "center"
                if *in_position_h {
                    *reading_align = Some('H');
                } else if *in_position_v {
                    *reading_align = Some('V');
                }
            }
            b"wrapTight" | b"wrapSquare" | b"wrapThrough" => {
                if let Some(val) = get_attr(e, b"wrapText")? {
                    *wrap_side = match val.as_str() {
                        "bothSides" => Some(WrapSide::BothSides),
                        "left" => Some(WrapSide::Left),
                        "right" => Some(WrapSide::Right),
                        _ => Some(WrapSide::BothSides),
                    };
                } else {
                    *wrap_side = Some(WrapSide::BothSides);
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
    if let ParseState::InDrawing {
        ref mut in_position_h,
        ref mut in_position_v,
        ref mut reading_pos_offset,
        ref mut reading_align,
        ..
    } = state
    {
        match local {
            b"positionH" => {
                *in_position_h = false;
                if *reading_pos_offset == Some('H') {
                    *reading_pos_offset = None;
                }
            }
            b"positionV" => {
                *in_position_v = false;
                if *reading_pos_offset == Some('V') {
                    *reading_pos_offset = None;
                }
            }
            b"posOffset" => {
                *reading_pos_offset = None;
            }
            b"align" => {
                *reading_align = None;
            }
            _ => {}
        }
    }
}
