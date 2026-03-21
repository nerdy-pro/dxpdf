use crate::error::Error;
use crate::model::*;
use crate::units::DEFAULT_PAGE_MARGIN_TWIPS;

use super::helpers::get_attr;
use super::ParseState;

/// Default header/footer distance from page edge in twips (0.5 inch).
const DEFAULT_HF_MARGIN_TWIPS: u32 = 720;

/// Handle elements inside a `w:sectPr` to extract page size, margins, and header/footer refs.
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
                let header = get_attr(e, b"header")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(DEFAULT_HF_MARGIN_TWIPS);
                let footer = get_attr(e, b"footer")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(DEFAULT_HF_MARGIN_TWIPS);
                section.page_margins = Some(PageMargins {
                    top,
                    right,
                    bottom,
                    left,
                    header,
                    footer,
                });
            }
            b"headerReference" => {
                // Store the relationship ID for the default header
                if let Some(hf_type) = get_attr(e, b"type")? {
                    if hf_type == "default" {
                        if let Some(rid) = get_attr(e, b"id")? {
                            section.header_rel_id = Some(rid);
                        }
                    }
                }
            }
            b"footerReference" => {
                if let Some(hf_type) = get_attr(e, b"type")? {
                    if hf_type == "default" {
                        if let Some(rid) = get_attr(e, b"id")? {
                            section.footer_rel_id = Some(rid);
                        }
                    }
                }
            }
            _ => {}
        }
    }
    Ok(())
}
