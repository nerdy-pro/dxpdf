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
            return Ok(Some(String::from_utf8_lossy(&attr.value).into_owned()));
        }
    }
    Ok(None)
}

/// Get an attribute value by local name, ignoring malformed attributes.
/// Infallible alternative to [`get_attr`] for contexts that don't propagate errors.
pub fn get_attr_lossy(e: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if local_name(attr.key.as_ref()) == name {
            return Some(String::from_utf8_lossy(&attr.value).into_owned());
        }
    }
    None
}

pub(super) fn matches_body_or_cell(state: &ParseState) -> bool {
    matches!(state, ParseState::InBody | ParseState::InTableCell { .. })
}

pub(super) fn take_paragraph(
    state: &mut ParseState,
) -> (
    ParagraphProperties,
    Vec<Inline>,
    Vec<FloatingImage>,
    Option<SectionProperties>,
) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InParagraph {
            props,
            runs,
            floats,
            section_props,
            ..
        } => (props, runs, floats, section_props),
        _ => (ParagraphProperties::default(), Vec::new(), Vec::new(), None),
    }
}

pub(super) fn take_run(state: &mut ParseState) -> (RunProperties, String) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InRun { props, text } => (props, text),
        _ => (RunProperties::default(), String::new()),
    }
}

pub(super) fn push_block(state: &mut ParseState, top_blocks: &mut Vec<Block>, block: Block) {
    match state {
        ParseState::InBody => top_blocks.push(block),
        ParseState::InTableCell { ref mut blocks, .. } => blocks.push(block),
        _ => top_blocks.push(block),
    }
}

/// Push a floating image into the nearest paragraph context.
pub(super) fn push_float(state: &mut ParseState, stack: &mut [ParseState], float: FloatingImage) {
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

/// Push an inline element into the nearest paragraph — current state first, then stack.
pub(super) fn push_inline_to_paragraph(
    state: &mut ParseState,
    stack: &mut [ParseState],
    inline: Inline,
) {
    if let ParseState::InParagraph { ref mut runs, .. } = state {
        runs.push(inline);
        return;
    }
    for s in stack.iter_mut().rev() {
        if let ParseState::InParagraph { ref mut runs, .. } = s {
            runs.push(inline);
            return;
        }
    }
}

/// Set or clear the `in_cell_mar` flag on InTable or InTableCell state.
pub(super) fn set_in_cell_mar(state: &mut ParseState, value: bool) {
    match state {
        ParseState::InTable {
            ref mut in_cell_mar,
            ..
        } => *in_cell_mar = value,
        ParseState::InTableCell {
            ref mut in_cell_mar,
            ..
        } => *in_cell_mar = value,
        _ => {}
    }
}

/// Set or clear the `in_borders` flag on InTable or InTableCell state.
pub(super) fn set_in_borders(state: &mut ParseState, value: bool) {
    match state {
        ParseState::InTable {
            ref mut in_borders, ..
        } => *in_borders = value,
        ParseState::InTableCell {
            ref mut in_borders, ..
        } => *in_borders = value,
        _ => {}
    }
}

/// Handle a w:fldChar element (begin/separate/end field code state machine).
pub(super) fn handle_fld_char(
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
    stack: &mut [ParseState],
    field_instr: &mut Option<String>,
    field_suppressing: &mut bool,
    field_props: &mut RunProperties,
) -> Result<(), Error> {
    if let Some(val) = get_attr(e, b"fldCharType")? {
        match val.as_str() {
            "begin" => {
                *field_instr = Some(String::new());
                *field_suppressing = false;
                if let ParseState::InRun { ref props, .. } = state {
                    *field_props = props.clone();
                }
            }
            "separate" => {
                *field_suppressing = true;
            }
            "end" => {
                if let Some(instr) = field_instr.take() {
                    let trimmed = instr.trim().to_uppercase();
                    let field_type = if trimmed.starts_with("PAGE") {
                        Some(FieldType::Page)
                    } else if trimmed.starts_with("NUMPAGES") {
                        Some(FieldType::NumPages)
                    } else {
                        None
                    };
                    if let Some(ft) = field_type {
                        let fc = FieldCode {
                            field_type: ft,
                            properties: field_props.clone(),
                        };
                        push_inline_to_paragraph(state, stack, Inline::Field(fc));
                    }
                }
                *field_suppressing = false;
            }
            _ => {}
        }
    }
    Ok(())
}
