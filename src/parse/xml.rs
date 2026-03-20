use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;
use crate::model::*;

/// Parse the `word/document.xml` content into a `Document`.
pub fn parse_document_xml(xml: &str) -> Result<Document, Error> {
    let mut reader = Reader::from_str(xml);
    let mut blocks: Vec<Block> = Vec::new();
    let mut stack: Vec<ParseState> = Vec::new();
    let mut state = ParseState::Idle;
    let mut final_section: Option<SectionProperties> = None;

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            Event::Start(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"body" => {
                        state = ParseState::InBody;
                    }
                    b"p" if matches_body_or_cell(&state) => {
                        stack.push(state);
                        state = ParseState::InParagraph {
                            props: ParagraphProperties::default(),
                            runs: Vec::new(),
                            floats: Vec::new(),
                            section_props: None,
                        };
                    }
                    b"pPr" if matches!(state, ParseState::InParagraph { .. }) => {
                        let (props, runs, floats, section_props) =
                            take_paragraph(&mut state);
                        stack.push(ParseState::InParagraph {
                            props: ParagraphProperties::default(),
                            runs,
                            floats,
                            section_props,
                        });
                        state = ParseState::InParagraphProperties {
                            props,
                            section_props: None,
                        };
                    }
                    b"r" if matches!(state, ParseState::InParagraph { .. }) => {
                        stack.push(state);
                        state = ParseState::InRun {
                            props: RunProperties::default(),
                            text: String::new(),
                        };
                    }
                    b"rPr" if matches!(state, ParseState::InRun { .. }) => {
                        let (props, text) = take_run(&mut state);
                        stack.push(ParseState::InRun {
                            props: RunProperties::default(),
                            text,
                        });
                        state = ParseState::InRunProperties { props };
                    }
                    b"tbl" if matches_body_or_cell(&state) => {
                        stack.push(state);
                        state = ParseState::InTable {
                            rows: Vec::new(),
                        };
                    }
                    b"tr" if matches!(state, ParseState::InTable { .. }) => {
                        stack.push(state);
                        state = ParseState::InTableRow {
                            cells: Vec::new(),
                        };
                    }
                    b"tc" if matches!(state, ParseState::InTableRow { .. }) => {
                        stack.push(state);
                        state = ParseState::InTableCell {
                            blocks: Vec::new(),
                        };
                    }
                    b"t" if matches!(state, ParseState::InRun { .. }) => {
                        stack.push(state);
                        state = ParseState::InText {
                            text: String::new(),
                        };
                    }
                    b"drawing"
                        if matches!(
                            state,
                            ParseState::InRun { .. } | ParseState::InParagraph { .. }
                        ) =>
                    {
                        stack.push(state);
                        state = ParseState::InDrawing {
                            depth: 1,
                            rel_id: None,
                            width_emu: None,
                            height_emu: None,
                            is_anchor: false,
                            pos_h_emu: None,
                            pos_v_emu: None,
                            wrap_side: None,
                            reading_pos_offset: None,
                            in_position_h: false,
                            in_position_v: false,
                        };
                    }
                    _ if matches!(state, ParseState::InDrawing { .. }) => {
                        // Track depth of nested elements inside drawing
                        if let ParseState::InDrawing { ref mut depth, .. } = state {
                            *depth += 1;
                        }
                        // Check for extent and blip on Start elements too
                        handle_drawing_element(local, e, &mut state)?;
                    }
                    b"sectPr"
                        if matches!(
                            state,
                            ParseState::InParagraphProperties { .. }
                                | ParseState::InBody
                        ) =>
                    {
                        stack.push(state);
                        state = ParseState::InSectionProperties {
                            section: SectionProperties {
                                page_size: None,
                                page_margins: None,
                            },
                        };
                    }
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if matches!(state, ParseState::InDrawing { .. }) {
                    handle_drawing_element(local, e, &mut state)?;
                } else if matches!(state, ParseState::InSectionProperties { .. }) {
                    handle_section_element(local, e, &mut state)?;
                } else {
                    handle_empty_element(local, e, &mut state)?;
                }
            }
            Event::Text(ref e) => {
                if let ParseState::InText { ref mut text, .. } = state {
                    text.push_str(&e.unescape()?);
                } else if let ParseState::InDrawing {
                    ref reading_pos_offset,
                    ref mut pos_h_emu,
                    ref mut pos_v_emu,
                    ..
                } = state
                {
                    if let Some(axis) = reading_pos_offset {
                        let val_str = e.unescape().unwrap_or_default();
                        if let Ok(val) = val_str.trim().parse::<i64>() {
                            match axis {
                                'H' => *pos_h_emu = Some(val),
                                'V' => *pos_v_emu = Some(val),
                                _ => {}
                            }
                        }
                    }
                }
            }
            Event::End(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"body" => {
                        state = ParseState::Idle;
                    }
                    b"p" if matches!(state, ParseState::InParagraph { .. }) => {
                        let (props, runs, floats, section_props) =
                            take_paragraph(&mut state);
                        state = stack.pop().unwrap_or(ParseState::Idle);
                        let paragraph = Block::Paragraph(Paragraph {
                            properties: props,
                            runs,
                            floats,
                            section_properties: section_props,
                        });
                        push_block(&mut state, &mut blocks, paragraph);
                    }
                    b"pPr" if matches!(state, ParseState::InParagraphProperties { .. }) => {
                        if let ParseState::InParagraphProperties {
                            props,
                            section_props,
                        } = state
                        {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InParagraph {
                                props: ref mut p,
                                section_props: ref mut sp,
                                ..
                            } = state
                            {
                                *p = props;
                                *sp = section_props;
                            }
                        }
                    }
                    b"r" if matches!(state, ParseState::InRun { .. }) => {
                        let (props, text) = take_run(&mut state);
                        state = stack.pop().unwrap_or(ParseState::Idle);
                        if !text.is_empty() {
                            if let ParseState::InParagraph { ref mut runs, .. } = state {
                                runs.push(Inline::TextRun(TextRun {
                                    text,
                                    properties: props,
                                }));
                            }
                        }
                    }
                    b"rPr" if matches!(state, ParseState::InRunProperties { .. }) => {
                        if let ParseState::InRunProperties { props } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InRun {
                                props: ref mut p, ..
                            } = state
                            {
                                *p = props;
                            }
                        }
                    }
                    b"t" if matches!(state, ParseState::InText { .. }) => {
                        if let ParseState::InText { text, .. } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InRun {
                                text: ref mut t, ..
                            } = state
                            {
                                t.push_str(&text);
                            }
                        }
                    }
                    b"tbl" if matches!(state, ParseState::InTable { .. }) => {
                        if let ParseState::InTable { rows } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            let table = Block::Table(Table { rows });
                            push_block(&mut state, &mut blocks, table);
                        }
                    }
                    b"tr" if matches!(state, ParseState::InTableRow { .. }) => {
                        if let ParseState::InTableRow { cells } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InTable { ref mut rows } = state {
                                rows.push(TableRow { cells });
                            }
                        }
                    }
                    b"tc" if matches!(state, ParseState::InTableCell { .. }) => {
                        if let ParseState::InTableCell {
                            blocks: cell_blocks,
                        } = state
                        {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InTableRow { ref mut cells } = state {
                                cells.push(TableCell {
                                    blocks: cell_blocks,
                                });
                            }
                        }
                    }
                    b"drawing" if matches!(state, ParseState::InDrawing { .. }) => {
                        let old = std::mem::replace(&mut state, ParseState::Idle);
                        if let ParseState::InDrawing {
                            rel_id,
                            width_emu,
                            height_emu,
                            is_anchor,
                            pos_h_emu,
                            pos_v_emu,
                            wrap_side,
                            ..
                        } = old
                        {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let Some(rid) = rel_id {
                                let w = emu_to_pt(width_emu.unwrap_or(0));
                                let h = emu_to_pt(height_emu.unwrap_or(0));

                                if is_anchor {
                                    // Floating/anchored image
                                    let float = FloatingImage {
                                        rel_id: rid,
                                        width_pt: w,
                                        height_pt: h,
                                        data: Vec::new(),
                                        format_hint: FormatHint::default(),
                                        offset_x_pt: emu_to_pt(
                                            pos_h_emu.unwrap_or(0).unsigned_abs(),
                                        ),
                                        offset_y_pt: emu_to_pt(
                                            pos_v_emu.unwrap_or(0).unsigned_abs(),
                                        ),
                                        wrap_side: wrap_side
                                            .unwrap_or(WrapSide::BothSides),
                                    };
                                    push_float(&mut state, &mut stack, float);
                                } else {
                                    // Inline image
                                    let image = Inline::Image(InlineImage {
                                        rel_id: rid,
                                        width_pt: w,
                                        height_pt: h,
                                        data: Vec::new(),
                                        format_hint: FormatHint::default(),
                                    });
                                    if matches!(state, ParseState::InRun { .. }) {
                                        if let Some(para_state) =
                                            stack.iter_mut().rev().find(|s| {
                                                matches!(
                                                    s,
                                                    ParseState::InParagraph { .. }
                                                )
                                            })
                                        {
                                            if let ParseState::InParagraph {
                                                ref mut runs,
                                                ..
                                            } = para_state
                                            {
                                                runs.push(image);
                                            }
                                        }
                                    } else {
                                        push_inline(&mut state, image);
                                    }
                                }
                            }
                        }
                    }
                    b"sectPr"
                        if matches!(
                            state,
                            ParseState::InSectionProperties { .. }
                        ) =>
                    {
                        if let ParseState::InSectionProperties { section } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            match state {
                                ParseState::InParagraphProperties {
                                    ref mut section_props,
                                    ..
                                } => {
                                    *section_props = Some(section);
                                }
                                ParseState::InBody => {
                                    final_section = Some(section);
                                }
                                _ => {}
                            }
                        }
                    }
                    _ if matches!(state, ParseState::InDrawing { .. }) => {
                        // Handle sub-element End events (positionH, positionV, posOffset)
                        handle_drawing_end(local, &mut state);
                        // Track depth of nested elements inside drawing
                        if let ParseState::InDrawing { ref mut depth, .. } = state {
                            *depth -= 1;
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }

    Ok(Document { blocks, final_section })
}

enum ParseState {
    Idle,
    InBody,
    InParagraph {
        props: ParagraphProperties,
        runs: Vec<Inline>,
        floats: Vec<FloatingImage>,
        section_props: Option<SectionProperties>,
    },
    InParagraphProperties {
        props: ParagraphProperties,
        section_props: Option<SectionProperties>,
    },
    InSectionProperties {
        section: SectionProperties,
    },
    InRun {
        props: RunProperties,
        text: String,
    },
    InRunProperties {
        props: RunProperties,
    },
    InText {
        text: String,
    },
    InTable {
        rows: Vec<TableRow>,
    },
    InTableRow {
        cells: Vec<TableCell>,
    },
    InTableCell {
        blocks: Vec<Block>,
    },
    /// Inside a `w:drawing` subtree. We track nested depth and collect image info.
    InDrawing {
        depth: u32,
        rel_id: Option<RelId>,
        width_emu: Option<u64>,
        height_emu: Option<u64>,
        is_anchor: bool,
        pos_h_emu: Option<i64>,
        pos_v_emu: Option<i64>,
        wrap_side: Option<WrapSide>,
        /// Which position offset we're currently reading text for: 'H' or 'V'.
        reading_pos_offset: Option<char>,
        in_position_h: bool,
        in_position_v: bool,
    },
}

fn matches_body_or_cell(state: &ParseState) -> bool {
    matches!(
        state,
        ParseState::InBody | ParseState::InTableCell { .. }
    )
}

fn take_paragraph(
    state: &mut ParseState,
) -> (ParagraphProperties, Vec<Inline>, Vec<FloatingImage>, Option<SectionProperties>) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InParagraph { props, runs, floats, section_props } => {
            (props, runs, floats, section_props)
        }
        _ => (ParagraphProperties::default(), Vec::new(), Vec::new(), None),
    }
}

fn take_run(state: &mut ParseState) -> (RunProperties, String) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InRun { props, text } => (props, text),
        _ => (RunProperties::default(), String::new()),
    }
}

fn push_block(state: &mut ParseState, top_blocks: &mut Vec<Block>, block: Block) {
    match state {
        ParseState::InBody => top_blocks.push(block),
        ParseState::InTableCell { ref mut blocks } => blocks.push(block),
        _ => top_blocks.push(block),
    }
}

/// Strip namespace prefix from an element name (e.g., `w:p` -> `p`).
fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|&b| b == b':')
        .map(|i| &name[i + 1..])
        .unwrap_or(name)
}

fn handle_empty_element(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    state: &mut ParseState,
) -> Result<(), Error> {
    match state {
        ParseState::InRunProperties { ref mut props } => {
            match local {
                b"b" => {
                    // <w:b/> means bold on; <w:b w:val="false"/> means off
                    props.bold = !is_val_false(e)?;
                }
                b"i" => {
                    props.italic = !is_val_false(e)?;
                }
                b"u" => {
                    // <w:u w:val="single"/> etc. means underline on; val="none" means off
                    let val = get_attr(e, b"val")?;
                    props.underline = val.as_deref() != Some("none");
                }
                b"sz" => {
                    if let Some(val) = get_attr(e, b"val")? {
                        props.font_size = val.parse::<u32>().ok();
                    }
                }
                b"rFonts" => {
                    // Try w:ascii first, then w:hAnsi
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
                _ => {}
            }
        }
        _ => {}
    }

    Ok(())
}

/// Push a floating image into the nearest paragraph context.
fn push_float(state: &mut ParseState, stack: &mut [ParseState], float: FloatingImage) {
    // Try current state first
    if let ParseState::InParagraph { ref mut floats, .. } = state {
        floats.push(float);
        return;
    }
    // Search the stack for the nearest paragraph
    for s in stack.iter_mut().rev() {
        if let ParseState::InParagraph { ref mut floats, .. } = s {
            floats.push(float);
            return;
        }
    }
}

/// Push an inline element into the current paragraph or run context.
fn push_inline(state: &mut ParseState, inline: Inline) {
    match state {
        ParseState::InParagraph { ref mut runs, .. } => {
            runs.push(inline);
        }
        ParseState::InRun { .. } => {
            // If we're in a run, we can't push directly; this shouldn't normally happen
            // since drawing End pops back to paragraph level
        }
        _ => {}
    }
}

/// Handle elements inside a `w:sectPr` to extract page size and margins.
fn handle_section_element(
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
                    .unwrap_or(1440);
                let right = get_attr(e, b"right")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(1440);
                let bottom = get_attr(e, b"bottom")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(1440);
                let left = get_attr(e, b"left")?
                    .and_then(|v| v.parse::<u32>().ok())
                    .unwrap_or(1440);
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

/// Handle elements inside a `w:drawing` subtree to extract image info.
fn handle_drawing_element(
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
        ..
    } = state
    {
        match local {
            b"anchor" => {
                *is_anchor = true;
            }
            b"inline" => {
                // Explicit inline — already the default (is_anchor = false)
            }
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
                    *rel_id = Some(RelId::new(embed));
                }
            }
            b"positionH" => {
                *in_position_h = true;
            }
            b"positionV" => {
                *in_position_v = true;
            }
            b"posOffset" => {
                // This is a Start element — text content will follow
                if *in_position_h {
                    *reading_pos_offset = Some('H');
                } else if *in_position_v {
                    *reading_pos_offset = Some('V');
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
            b"wrapNone" => {
                // No text wrapping — still a float but text ignores it
            }
            _ => {}
        }
    }
    Ok(())
}

/// Handle End events inside a drawing subtree for position tracking.
fn handle_drawing_end(local: &[u8], state: &mut ParseState) {
    if let ParseState::InDrawing {
        ref mut in_position_h,
        ref mut in_position_v,
        ref mut reading_pos_offset,
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
            _ => {}
        }
    }
}

/// Check if a toggle element has w:val="false" or w:val="0".
fn is_val_false(e: &quick_xml::events::BytesStart<'_>) -> Result<bool, Error> {
    if let Some(val) = get_attr(e, b"val")? {
        Ok(val == "false" || val == "0")
    } else {
        Ok(false)
    }
}

/// Get an attribute value by local name (stripping namespace prefix).
fn get_attr(
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

#[cfg(test)]
mod tests {
    use super::*;

    fn wrap_body(content: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>{content}</w:body>
            </w:document>"#
        )
    }

    #[test]
    fn parse_empty_document() {
        let xml = wrap_body("");
        let doc = parse_document_xml(&xml).unwrap();
        assert!(doc.blocks.is_empty());
    }

    #[test]
    fn parse_single_paragraph() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        assert_eq!(doc.blocks.len(), 1);
        match &doc.blocks[0] {
            Block::Paragraph(p) => {
                assert_eq!(p.runs.len(), 1);
                match &p.runs[0] {
                    Inline::TextRun(tr) => assert_eq!(tr.text, "Hello World"),
                    _ => panic!("Expected TextRun"),
                }
            }
            _ => panic!("Expected Paragraph"),
        }
    }

    #[test]
    fn parse_bold_italic() {
        let xml = wrap_body(
            r#"<w:p><w:r>
                <w:rPr><w:b/><w:i/></w:rPr>
                <w:t>Bold Italic</w:t>
            </w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!("Expected Paragraph");
        };
        let Inline::TextRun(tr) = &p.runs[0] else {
            panic!("Expected TextRun");
        };
        assert!(tr.properties.bold);
        assert!(tr.properties.italic);
        assert!(!tr.properties.underline);
    }

    #[test]
    fn parse_font_size_and_color() {
        let xml = wrap_body(
            r#"<w:p><w:r>
                <w:rPr>
                    <w:sz w:val="28"/>
                    <w:color w:val="FF0000"/>
                </w:rPr>
                <w:t>Red 14pt</w:t>
            </w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!();
        };
        let Inline::TextRun(tr) = &p.runs[0] else {
            panic!();
        };
        assert_eq!(tr.properties.font_size, Some(28));
        assert_eq!(tr.properties.font_size_pt(), 14.0);
        assert_eq!(
            tr.properties.color,
            Some(Color { r: 255, g: 0, b: 0 })
        );
    }

    #[test]
    fn parse_alignment() {
        let xml = wrap_body(
            r#"<w:p>
                <w:pPr><w:jc w:val="center"/></w:pPr>
                <w:r><w:t>Centered</w:t></w:r>
            </w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!();
        };
        assert_eq!(p.properties.alignment, Some(Alignment::Center));
    }

    #[test]
    fn parse_spacing_and_indentation() {
        let xml = wrap_body(
            r#"<w:p>
                <w:pPr>
                    <w:spacing w:before="240" w:after="120" w:line="360"/>
                    <w:ind w:left="720" w:firstLine="360"/>
                </w:pPr>
                <w:r><w:t>Indented</w:t></w:r>
            </w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!();
        };
        let spacing = p.properties.spacing.unwrap();
        assert_eq!(spacing.before, Some(240));
        assert_eq!(spacing.after, Some(120));
        assert_eq!(spacing.line, Some(360));

        let indent = p.properties.indentation.unwrap();
        assert_eq!(indent.left, Some(720));
        assert_eq!(indent.first_line, Some(360));
    }

    #[test]
    fn parse_table() {
        let xml = wrap_body(
            r#"<w:tbl>
                <w:tr>
                    <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
                </w:tr>
                <w:tr>
                    <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
                </w:tr>
            </w:tbl>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        assert_eq!(doc.blocks.len(), 1);
        let Block::Table(table) = &doc.blocks[0] else {
            panic!("Expected Table");
        };
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0].cells.len(), 2);
        assert_eq!(table.rows[1].cells.len(), 2);

        // Check cell content
        let Block::Paragraph(p) = &table.rows[0].cells[0].blocks[0] else {
            panic!();
        };
        let Inline::TextRun(tr) = &p.runs[0] else {
            panic!();
        };
        assert_eq!(tr.text, "A1");
    }

    #[test]
    fn parse_multiple_runs() {
        let xml = wrap_body(
            r#"<w:p>
                <w:r><w:t>Hello </w:t></w:r>
                <w:r><w:rPr><w:b/></w:rPr><w:t>World</w:t></w:r>
            </w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!();
        };
        assert_eq!(p.runs.len(), 2);
        let Inline::TextRun(tr1) = &p.runs[0] else {
            panic!();
        };
        let Inline::TextRun(tr2) = &p.runs[1] else {
            panic!();
        };
        assert_eq!(tr1.text, "Hello ");
        assert!(!tr1.properties.bold);
        assert_eq!(tr2.text, "World");
        assert!(tr2.properties.bold);
    }

    #[test]
    fn parse_font_family() {
        let xml = wrap_body(
            r#"<w:p><w:r>
                <w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>
                <w:t>Arial text</w:t>
            </w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!();
        };
        let Inline::TextRun(tr) = &p.runs[0] else {
            panic!();
        };
        assert_eq!(tr.properties.font_family.as_deref(), Some("Arial"));
    }

    #[test]
    fn parse_inline_image() {
        let xml = wrap_body(
            r#"<w:p><w:r>
                <w:drawing>
                    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                        <wp:extent cx="914400" cy="457200"/>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData>
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:blipFill>
                                        <a:blip r:embed="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                    </pic:blipFill>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!("Expected Paragraph");
        };
        assert_eq!(p.runs.len(), 1);
        let Inline::Image(img) = &p.runs[0] else {
            panic!("Expected Image, got {:?}", p.runs[0]);
        };
        assert_eq!(img.rel_id, RelId::new("rId5"));
        // 914400 EMU = 72pt (1 inch)
        assert!((img.width_pt - 72.0).abs() < 0.1);
        // 457200 EMU = 36pt (0.5 inch)
        assert!((img.height_pt - 36.0).abs() < 0.1);
    }

    #[test]
    fn parse_image_with_text() {
        let xml = wrap_body(
            r#"<w:p>
                <w:r><w:t>Before </w:t></w:r>
                <w:r>
                    <w:drawing>
                        <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                            <wp:extent cx="914400" cy="914400"/>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                <a:graphicData>
                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:blipFill>
                                            <a:blip r:embed="rId3" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                        </pic:blipFill>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing>
                </w:r>
                <w:r><w:t> After</w:t></w:r>
            </w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else {
            panic!();
        };
        assert_eq!(p.runs.len(), 3);
        assert!(matches!(p.runs[0], Inline::TextRun(_)));
        assert!(matches!(p.runs[1], Inline::Image(_)));
        assert!(matches!(p.runs[2], Inline::TextRun(_)));
    }

    #[test]
    fn parse_section_properties() {
        let xml = wrap_body(
            r#"<w:p>
                <w:pPr>
                    <w:sectPr>
                        <w:pgSz w:w="11906" w:h="16838"/>
                        <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
                    </w:sectPr>
                </w:pPr>
                <w:r><w:t>Section 1</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Section 2</w:t></w:r></w:p>
            <w:sectPr>
                <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
                <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        assert_eq!(doc.blocks.len(), 2);

        // First paragraph should have section properties (mid-document break)
        let Block::Paragraph(p1) = &doc.blocks[0] else { panic!() };
        let sect1 = p1.section_properties.as_ref().expect("section_properties missing");
        let ps1 = sect1.page_size.unwrap();
        assert_eq!(ps1.width, 11906); // A4 portrait
        assert_eq!(ps1.height, 16838);
        let pm1 = sect1.page_margins.unwrap();
        assert_eq!(pm1.top, 720);

        // Second paragraph should NOT have section properties
        let Block::Paragraph(p2) = &doc.blocks[1] else { panic!() };
        assert!(p2.section_properties.is_none());

        // Final section should be on the document
        let final_sect = doc.final_section.as_ref().expect("final_section missing");
        let ps2 = final_sect.page_size.unwrap();
        assert_eq!(ps2.width, 16838); // A4 landscape
        assert_eq!(ps2.height, 11906);
        let pm2 = final_sect.page_margins.unwrap();
        assert_eq!(pm2.top, 1440);
    }
}
