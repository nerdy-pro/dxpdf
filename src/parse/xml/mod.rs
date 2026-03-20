mod drawing;
mod helpers;
mod properties;
mod section;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;
use crate::model::*;

use drawing::{handle_drawing_element, handle_drawing_end};
use helpers::*;
use properties::handle_empty_element;
use section::handle_section_element;

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
                            grid_cols: Vec::new(),
                            default_cell_margins: None,
                            in_cell_mar: false,
                            borders: None,
                            in_borders: false,
                        };
                    }
                    b"tr" if matches!(state, ParseState::InTable { .. }) => {
                        stack.push(state);
                        state = ParseState::InTableRow {
                            cells: Vec::new(),
                            height: None,
                        };
                    }
                    b"tc" if matches!(state, ParseState::InTableRow { .. }) => {
                        stack.push(state);
                        state = ParseState::InTableCell {
                            blocks: Vec::new(),
                            width: None,
                            grid_span: 1,
                            vertical_merge: None,
                            cell_margins: None,
                            in_cell_mar: false,
                            cell_borders: None,
                            in_borders: false,
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
                        if let ParseState::InDrawing { ref mut depth, .. } = state {
                            *depth += 1;
                        }
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
                    b"tblCellMar" if matches!(state, ParseState::InTable { .. }) => {
                        if let ParseState::InTable { ref mut in_cell_mar, .. } = state {
                            *in_cell_mar = true;
                        }
                    }
                    b"tblBorders" if matches!(state, ParseState::InTable { .. }) => {
                        if let ParseState::InTable { ref mut in_borders, .. } = state {
                            *in_borders = true;
                        }
                    }
                    b"tcMar" if matches!(state, ParseState::InTableCell { .. }) => {
                        if let ParseState::InTableCell { ref mut in_cell_mar, .. } = state {
                            *in_cell_mar = true;
                        }
                    }
                    b"tcBorders" if matches!(state, ParseState::InTableCell { .. }) => {
                        if let ParseState::InTableCell { ref mut in_borders, .. } = state {
                            *in_borders = true;
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"p" && matches_body_or_cell(&state) {
                    let paragraph = Block::Paragraph(Paragraph {
                        properties: ParagraphProperties::default(),
                        runs: Vec::new(),
                        floats: Vec::new(),
                        section_properties: None,
                    });
                    push_block(&mut state, &mut blocks, paragraph);
                } else if (local == b"br" || local == b"tab")
                    && matches!(state, ParseState::InRun { .. })
                {
                    if let ParseState::InRun {
                        ref props,
                        ref mut text,
                        ..
                    } = state
                    {
                        let flushed_text = std::mem::take(text);
                        let inline_to_push = if local == b"br" {
                            Inline::LineBreak
                        } else {
                            Inline::Tab
                        };
                        if let Some(para_state) = stack.iter_mut().rev().find(
                            |s| matches!(s, ParseState::InParagraph { .. }),
                        ) {
                            if let ParseState::InParagraph {
                                ref mut runs, ..
                            } = para_state
                            {
                                if !flushed_text.is_empty() {
                                    runs.push(Inline::TextRun(TextRun {
                                        text: flushed_text,
                                        properties: props.clone(),
                                    }));
                                }
                                runs.push(inline_to_push);
                            }
                        }
                    }
                } else if matches!(state, ParseState::InDrawing { .. }) {
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
                                let cleaned: String = text
                                    .chars()
                                    .filter(|c| *c != '\n' && *c != '\r')
                                    .collect();
                                t.push_str(&cleaned);
                            }
                        }
                    }
                    b"tbl" if matches!(state, ParseState::InTable { .. }) => {
                        if let ParseState::InTable { rows, grid_cols, default_cell_margins, borders, .. } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            let table = Block::Table(Table { rows, grid_cols, default_cell_margins, cell_spacing: None, borders });
                            push_block(&mut state, &mut blocks, table);
                        }
                    }
                    b"tr" if matches!(state, ParseState::InTableRow { .. }) => {
                        if let ParseState::InTableRow { cells, height } = state {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InTable { ref mut rows, .. } = state {
                                rows.push(TableRow { cells, height });
                            }
                        }
                    }
                    b"tc" if matches!(state, ParseState::InTableCell { .. }) => {
                        if let ParseState::InTableCell {
                            blocks: cell_blocks,
                            width: cell_width,
                            grid_span,
                            vertical_merge,
                            cell_margins,
                            cell_borders,
                            ..
                        } = state
                        {
                            state = stack.pop().unwrap_or(ParseState::Idle);
                            if let ParseState::InTableRow { ref mut cells, .. } = state {
                                cells.push(TableCell {
                                    blocks: cell_blocks,
                                    width: cell_width,
                                    grid_span,
                                    vertical_merge,
                                    cell_margins,
                                    cell_borders,
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
                        handle_drawing_end(local, &mut state);
                        if let ParseState::InDrawing { ref mut depth, .. } = state {
                            *depth -= 1;
                        }
                    }
                    b"tblCellMar" if matches!(state, ParseState::InTable { in_cell_mar: true, .. }) => {
                        if let ParseState::InTable { ref mut in_cell_mar, .. } = state {
                            *in_cell_mar = false;
                        }
                    }
                    b"tblBorders" if matches!(state, ParseState::InTable { in_borders: true, .. }) => {
                        if let ParseState::InTable { ref mut in_borders, .. } = state {
                            *in_borders = false;
                        }
                    }
                    b"tcMar" if matches!(state, ParseState::InTableCell { in_cell_mar: true, .. }) => {
                        if let ParseState::InTableCell { ref mut in_cell_mar, .. } = state {
                            *in_cell_mar = false;
                        }
                    }
                    b"tcBorders" if matches!(state, ParseState::InTableCell { in_borders: true, .. }) => {
                        if let ParseState::InTableCell { ref mut in_borders, .. } = state {
                            *in_borders = false;
                        }
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }

    Ok(Document {
        blocks,
        final_section,
        ..Document::default()
    })
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
        grid_cols: Vec<u32>,
        default_cell_margins: Option<CellMargins>,
        in_cell_mar: bool,
        borders: Option<TableBorders>,
        in_borders: bool,
    },
    InTableRow {
        cells: Vec<TableCell>,
        height: Option<u32>,
    },
    InTableCell {
        blocks: Vec<Block>,
        width: Option<u32>,
        grid_span: u32,
        vertical_merge: Option<VerticalMerge>,
        cell_margins: Option<CellMargins>,
        in_cell_mar: bool,
        cell_borders: Option<CellBorders>,
        in_borders: bool,
    },
    InDrawing {
        depth: u32,
        rel_id: Option<RelId>,
        width_emu: Option<u64>,
        height_emu: Option<u64>,
        is_anchor: bool,
        pos_h_emu: Option<i64>,
        pos_v_emu: Option<i64>,
        wrap_side: Option<WrapSide>,
        reading_pos_offset: Option<char>,
        in_position_h: bool,
        in_position_v: bool,
    },
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
        let xml = wrap_body(r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#);
        let doc = parse_document_xml(&xml).unwrap();
        assert_eq!(doc.blocks.len(), 1);
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        assert_eq!(p.runs.len(), 1);
        let Inline::TextRun(tr) = &p.runs[0] else { panic!() };
        assert_eq!(tr.text, "Hello World");
    }

    #[test]
    fn parse_bold_italic() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold Italic</w:t></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        let Inline::TextRun(tr) = &p.runs[0] else { panic!() };
        assert!(tr.properties.bold);
        assert!(tr.properties.italic);
        assert!(!tr.properties.underline);
    }

    #[test]
    fn parse_font_size_and_color() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Red 14pt</w:t></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        let Inline::TextRun(tr) = &p.runs[0] else { panic!() };
        assert_eq!(tr.properties.font_size, Some(28));
        assert_eq!(tr.properties.font_size_pt(), 14.0);
        assert_eq!(tr.properties.color, Some(Color { r: 255, g: 0, b: 0 }));
    }

    #[test]
    fn parse_alignment() {
        let xml = wrap_body(
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        assert_eq!(p.properties.alignment, Some(Alignment::Center));
    }

    #[test]
    fn parse_spacing_and_indentation() {
        let xml = wrap_body(
            r#"<w:p><w:pPr>
                <w:spacing w:before="240" w:after="120" w:line="360"/>
                <w:ind w:left="720" w:firstLine="360"/>
            </w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
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
        let Block::Table(table) = &doc.blocks[0] else { panic!() };
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0].cells.len(), 2);
        let Block::Paragraph(p) = &table.rows[0].cells[0].blocks[0] else { panic!() };
        let Inline::TextRun(tr) = &p.runs[0] else { panic!() };
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
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        assert_eq!(p.runs.len(), 2);
        let Inline::TextRun(tr1) = &p.runs[0] else { panic!() };
        let Inline::TextRun(tr2) = &p.runs[1] else { panic!() };
        assert_eq!(tr1.text, "Hello ");
        assert!(!tr1.properties.bold);
        assert_eq!(tr2.text, "World");
        assert!(tr2.properties.bold);
    }

    #[test]
    fn parse_font_family() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>Arial text</w:t></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        let Inline::TextRun(tr) = &p.runs[0] else { panic!() };
        assert_eq!(tr.properties.font_family.as_deref(), Some("Arial"));
    }

    #[test]
    fn parse_inline_image() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:drawing>
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
            </w:drawing></w:r></w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        let Inline::Image(img) = &p.runs[0] else { panic!() };
        assert_eq!(img.rel_id, RelId::from("rId5"));
        assert!((img.width_pt - 72.0).abs() < 0.1);
        assert!((img.height_pt - 36.0).abs() < 0.1);
    }

    #[test]
    fn parse_image_with_text() {
        let xml = wrap_body(
            r#"<w:p>
                <w:r><w:t>Before </w:t></w:r>
                <w:r><w:drawing>
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
                </w:drawing></w:r>
                <w:r><w:t> After</w:t></w:r>
            </w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        assert_eq!(p.runs.len(), 3);
        assert!(matches!(p.runs[0], Inline::TextRun(_)));
        assert!(matches!(p.runs[1], Inline::Image(_)));
        assert!(matches!(p.runs[2], Inline::TextRun(_)));
    }

    #[test]
    fn parse_section_properties() {
        let xml = wrap_body(
            r#"<w:p><w:pPr><w:sectPr>
                    <w:pgSz w:w="11906" w:h="16838"/>
                    <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
                </w:sectPr></w:pPr>
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
        let Block::Paragraph(p1) = &doc.blocks[0] else { panic!() };
        let sect1 = p1.section_properties.as_ref().unwrap();
        assert_eq!(sect1.page_size.unwrap().width, 11906);
        assert_eq!(sect1.page_size.unwrap().height, 16838);
        assert_eq!(sect1.page_margins.unwrap().top, 720);

        let Block::Paragraph(p2) = &doc.blocks[1] else { panic!() };
        assert!(p2.section_properties.is_none());

        let final_sect = doc.final_section.as_ref().unwrap();
        assert_eq!(final_sect.page_size.unwrap().width, 16838);
        assert_eq!(final_sect.page_margins.unwrap().top, 1440);
    }

    #[test]
    fn parse_tab_stops() {
        let xml = wrap_body(
            r#"<w:p><w:pPr><w:tabs>
                    <w:tab w:val="left" w:pos="2880"/>
                    <w:tab w:val="center" w:pos="4320"/>
                    <w:tab w:val="right" w:pos="9360"/>
                </w:tabs></w:pPr>
                <w:r><w:t>Col1</w:t></w:r>
            </w:p>"#,
        );
        let doc = parse_document_xml(&xml).unwrap();
        let Block::Paragraph(p) = &doc.blocks[0] else { panic!() };
        assert_eq!(p.properties.tab_stops.len(), 3);
        assert_eq!(p.properties.tab_stops[0].position, 2880);
        assert_eq!(p.properties.tab_stops[0].stop_type, TabStopType::Left);
        assert_eq!(p.properties.tab_stops[1].position, 4320);
        assert_eq!(p.properties.tab_stops[1].stop_type, TabStopType::Center);
        assert_eq!(p.properties.tab_stops[2].position, 9360);
        assert_eq!(p.properties.tab_stops[2].stop_type, TabStopType::Right);
    }
}
