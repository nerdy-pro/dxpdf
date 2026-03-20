use crate::error::Error;
use crate::model::*;
use crate::units::DEFAULT_PAGE_MARGIN_TWIPS;

use super::ParseState;
use super::helpers::get_attr;

/// Handle elements inside a `w:sectPr` to extract page size and margins.
pub fn handle_section_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
) -> Result<(), Error> {
    if let ParseState::InSectionProperties { ref mut section } = state {
        match local {
            b"pgSz" => {
                let w = get_attr(e, b"w")?.and_then(|v| v.parse::<u32>().ok());
                let h = get_attr(e, b"h")?.and_then(|v| v.parse::<u32>().ok());
                if let (Some(w), Some(h)) = (w, h) {
                    section.page_size = Some(PageSize {
                        width: w,
                        height: h,
                    });
                }
            }
            b"pgMar" => {
                let top = get_attr(e, b"top")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(DEFAULT_PAGE_MARGIN_TWIPS);
                let right = get_attr(e, b"right")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(DEFAULT_PAGE_MARGIN_TWIPS);
                let bottom = get_attr(e, b"bottom")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(DEFAULT_PAGE_MARGIN_TWIPS);
                let left = get_attr(e, b"left")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(DEFAULT_PAGE_MARGIN_TWIPS);
                section.page_margins = Some(PageMargins {
                    top,
                    right,
                    bottom,
                    left,
                });
            }
            _ => {}
        }
    }
    Ok(())
}
