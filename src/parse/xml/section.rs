use crate::dimension::Twips;
use crate::error::Error;
use crate::model::*;

use super::helpers::get_attr;
use super::ParseState;

/// Default page margin in twips (1440 = 1 inch, OOXML default).
const DEFAULT_PAGE_MARGIN_TWIPS: Twips = Twips::new(1440);

/// Default header/footer distance from page edge in twips (0.5 inch).
const DEFAULT_HF_MARGIN_TWIPS: Twips = Twips::new(720);

/// Parse a twips attribute value, returning the given default if absent or unparseable.
fn parse_twips_attr(
    e: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    default: Twips,
) -> Result<Twips, Error> {
    Ok(get_attr(e, name)?
        .and_then(|v| v.parse::<i64>().ok())
        .map(Twips::new)
        .unwrap_or(default))
}

/// Handle elements inside a `w:sectPr` to extract page size, margins, and header/footer refs.
pub fn handle_section_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
) -> Result<(), Error> {
    if let ParseState::InSectionProperties { ref mut section } = state {
        match local {
            b"pgSz" => {
                let w = get_attr(e, b"w")?
                    .and_then(|v| v.parse::<i64>().ok())
                    .map(Twips::new);
                let h = get_attr(e, b"h")?
                    .and_then(|v| v.parse::<i64>().ok())
                    .map(Twips::new);
                if let (Some(w), Some(h)) = (w, h) {
                    section.page_size = Some(PageSize::new(w, h));
                }
            }
            b"pgMar" => {
                let top = parse_twips_attr(e, b"top", DEFAULT_PAGE_MARGIN_TWIPS)?;
                let right = parse_twips_attr(e, b"right", DEFAULT_PAGE_MARGIN_TWIPS)?;
                let bottom = parse_twips_attr(e, b"bottom", DEFAULT_PAGE_MARGIN_TWIPS)?;
                let left = parse_twips_attr(e, b"left", DEFAULT_PAGE_MARGIN_TWIPS)?;
                let header = parse_twips_attr(e, b"header", DEFAULT_HF_MARGIN_TWIPS)?;
                let footer = parse_twips_attr(e, b"footer", DEFAULT_HF_MARGIN_TWIPS)?;
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
