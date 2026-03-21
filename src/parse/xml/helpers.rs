use crate::error::Error;
use crate::model::*;

use super::ParseState;

/// Strip namespace prefix from an element name (e.g., `w:p` -> `p`).
pub fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|&b| b == b':')
        .map(|i| &name[i + 1..])
        .unwrap_or(name)
}

/// Check if a toggle element has w:val="false" or w:val="0".
pub fn is_val_false(e: &quick_xml::events::BytesStart<'_>) -> Result<bool, Error> {
    if let Some(val) = get_attr(e, b"val")? {
        Ok(val == "false" || val == "0")
    } else {
        Ok(false)
    }
}

/// Get an attribute value by local name (stripping namespace prefix).
pub fn get_attr(
    e: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, Error> {
    for attr in e.attributes() {
        let attr = attr?;
        let attr_local = local_name(attr.key.as_ref());
        if attr_local == name {
            return Ok(Some(
                String::from_utf8_lossy(&attr.value).into_owned(),
            ));
        }
    }
    Ok(None)
}

pub fn matches_body_or_cell(state: &ParseState) -> bool {
    matches!(
        state,
        ParseState::InBody | ParseState::InTableCell { .. }
    )
}

pub fn take_paragraph(
    state: &mut ParseState,
) -> (ParagraphProperties, Vec<Inline>, Vec<FloatingImage>, Option<SectionProperties>) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InParagraph { props, runs, floats, section_props, .. } => {
            (props, runs, floats, section_props)
        }
        _ => (ParagraphProperties::default(), Vec::new(), Vec::new(), None),
    }
}

pub fn take_run(state: &mut ParseState) -> (RunProperties, String) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InRun { props, text } => (props, text),
        _ => (RunProperties::default(), String::new()),
    }
}

pub fn push_block(state: &mut ParseState, top_blocks: &mut Vec<Block>, block: Block) {
    match state {
        ParseState::InBody => top_blocks.push(block),
        ParseState::InTableCell { ref mut blocks, .. } => blocks.push(block),
        _ => top_blocks.push(block),
    }
}

/// Push a floating image into the nearest paragraph context.
pub fn push_float(state: &mut ParseState, stack: &mut [ParseState], float: FloatingImage) {
    if let ParseState::InParagraph { ref mut floats, .. } = state {
        floats.push(float);
        return;
    }
    for s in stack.iter_mut().rev() {
        if let ParseState::InParagraph { ref mut floats, .. } = s {
            floats.push(float);
            return;
        }
    }
}

/// Push an inline element into the current paragraph or run context.
pub fn push_inline(state: &mut ParseState, inline: Inline) {
    match state {
        ParseState::InParagraph { ref mut runs, .. } => {
            runs.push(inline);
        }
        ParseState::InRun { .. } => {}
        _ => {}
    }
}
