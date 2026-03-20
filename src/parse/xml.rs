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
                        };
                    }
                    b"pPr" if matches!(state, ParseState::InParagraph { .. }) => {
                        let (props, runs) = take_paragraph(&mut state);
                        stack.push(ParseState::InParagraph {
                            props: ParagraphProperties::default(),
                            runs,
                        });
                        state = ParseState::InParagraphProperties { props };
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
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                handle_empty_element(local, e, &mut state)?;
            }
            Event::Text(ref e) => {
                if let ParseState::InText { ref mut text, .. } = state {
                    text.push_str(&e.unescape()?);
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
                        let (props, runs) = take_paragraph(&mut state);
                        state = stack.pop().unwrap_or(ParseState::Idle);
                        let paragraph = Block::Paragraph(Paragraph {
                            properties: props,
                            runs,
                        });
                        push_block(&mut state, &mut blocks, paragraph);
                    }
                    b"pPr" if matches!(state, ParseState::InParagraphProperties { .. }) => {
                        if let ParseState::InParagraphProperties { props } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InParagraph {
                                props: ref mut p, ..
                            } = state
                            {
                                *p = props;
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
                    _ => {}
                }
            }
            _ => {}
        }
    }

    Ok(Document { blocks })
}

enum ParseState {
    Idle,
    InBody,
    InParagraph {
        props: ParagraphProperties,
        runs: Vec<Inline>,
    },
    InParagraphProperties {
        props: ParagraphProperties,
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
}

fn matches_body_or_cell(state: &ParseState) -> bool {
    matches!(
        state,
        ParseState::InBody | ParseState::InTableCell { .. }
    )
}

fn take_paragraph(state: &mut ParseState) -> (ParagraphProperties, Vec<Inline>) {
    let old = std::mem::replace(state, ParseState::Idle);
    match old {
        ParseState::InParagraph { props, runs } => (props, runs),
        _ => (ParagraphProperties::default(), Vec::new()),
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
        ParseState::InParagraphProperties { ref mut props } => {
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
}
