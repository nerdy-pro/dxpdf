mod drawing;
mod helpers;
mod properties;
mod section;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::error::Error;
use crate::model::*;
use crate::units;

use drawing::{handle_drawing_element, handle_drawing_end};
use helpers::*;
use properties::handle_empty_element;
use section::handle_section_element;

/// Parse the `word/document.xml` content into a `Document`.
pub fn parse_document_xml(xml: &str) -> Result<Document, Error> {
    let empty = std::collections::HashMap::new();
    parse_document_xml_with_rels(xml, &empty)
}

/// Parse document XML with relationship data for hyperlink resolution.
pub fn parse_document_xml_with_rels(
    xml: &str,
    rels: &std::collections::HashMap<String, String>,
) -> Result<Document, Error> {
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
                    b"body" | b"hdr" | b"ftr" => {
                        state = ParseState::InBody;
                    }
                    b"p" if matches_body_or_cell(&state) => {
                        stack.push(state);
                        state = ParseState::InParagraph {
                            props: ParagraphProperties::default(),
                            runs: Vec::new(),
                            floats: Vec::new(),
                            section_props: None,
                            hyperlink_url: None,
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
                            hyperlink_url: None,
                        });
                        state = ParseState::InParagraphProperties {
                            props,
                            section_props: None,
                        };
                    }
                    b"hyperlink" if matches!(state, ParseState::InParagraph { .. }) => {
                        // Resolve hyperlink URL from r:id relationship
                        if let ParseState::InParagraph { ref mut hyperlink_url, .. } = state {
                            let url = get_attr(e, b"id")?
                                .and_then(|rid| rels.get(&rid).cloned());
                            *hyperlink_url = url;
                        }
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
                            shading: None,
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
                            align_h: None,
                            align_v: None,
                            reading_align: None,
                            wrap_side: None,
                            reading_pos_offset: None,
                            in_position_h: false,
                            in_position_v: false,
                            pct_pos_h: None,
                            pct_pos_v: None,
                            reading_pct_pos: None,
                        };
                    }
                    _ if matches!(state, ParseState::InSectionProperties { .. }) => {
                        // Handle headerReference/footerReference etc. as Start events
                        handle_section_element(local, e, &mut state)?;
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
                                header: None,
                                footer: None,
                                header_rel_id: None,
                                footer_rel_id: None,
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
                                        hyperlink_url: None,
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
                    ref reading_align,
                    ref reading_pct_pos,
                    ref mut pos_h_emu,
                    ref mut pos_v_emu,
                    ref mut align_h,
                    ref mut align_v,
                    ref mut pct_pos_h,
                    ref mut pct_pos_v,
                    ..
                } = state
                {
                    if let Some(axis) = reading_align {
                        let val_str = e.unescape().unwrap_or_default();
                        let val = val_str.trim().to_string();
                        match axis {
                            'H' => *align_h = Some(val),
                            'V' => *align_v = Some(val),
                            _ => {}
                        }
                    } else if let Some(axis) = reading_pct_pos {
                        let val_str = e.unescape().unwrap_or_default();
                        if let Ok(val) = val_str.trim().parse::<i32>() {
                            match axis {
                                'H' => *pct_pos_h = Some(val),
                                'V' => *pct_pos_v = Some(val),
                                _ => {}
                            }
                        }
                    } else if let Some(axis) = reading_pos_offset {
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
                    b"body" | b"hdr" | b"ftr" => {
                        state = ParseState::Idle;
                    }
                    b"hyperlink" if matches!(state, ParseState::InParagraph { .. }) => {
                        if let ParseState::InParagraph { ref mut hyperlink_url, .. } = state {
                            *hyperlink_url = None;
                        }
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
                            if let ParseState::InParagraph { ref mut runs, ref hyperlink_url, .. } = state {
                                runs.push(Inline::TextRun(TextRun {
                                    text,
                                    properties: props,
                                    hyperlink_url: hyperlink_url.clone(),
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
                            shading,
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
                                    shading,
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
                            align_h,
                            align_v,
                            wrap_side,
                            pct_pos_h,
                            pct_pos_v,
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
                                        data: std::rc::Rc::new(Vec::new()),
                                        format_hint: FormatHint::default(),
                                        offset_x_pt: units::emu_to_pt_signed(
                                            pos_h_emu.unwrap_or(0),
                                        ),
                                        offset_y_pt: units::emu_to_pt_signed(
                                            pos_v_emu.unwrap_or(0),
                                        ),
                                        align_h,
                                        align_v,
                                        wrap_side: wrap_side
                                            .unwrap_or(WrapSide::BothSides),
                                        pct_pos_h,
                                        pct_pos_v,
                                    };
                                    push_float(&mut state, &mut stack, float);
                                } else {
                                    let image = Inline::Image(InlineImage {
                                        rel_id: rid,
                                        width_pt: w,
                                        height_pt: h,
                                        data: std::rc::Rc::new(Vec::new()),
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

/// Parse a header or footer XML file into a list of blocks.
/// Header/footer XML has the same structure as document body
/// but with `w:hdr` or `w:ftr` as the root element.
pub fn parse_header_footer_xml(xml: &str) -> Result<HeaderFooter, Error> {
    let empty = std::collections::HashMap::new();
    parse_header_footer_xml_with_rels(xml, &empty)
}

pub fn parse_header_footer_xml_with_rels(
    xml: &str,
    rels: &std::collections::HashMap<String, String>,
) -> Result<HeaderFooter, Error> {
    let doc = parse_document_xml_with_rels(xml, rels)?;
    Ok(HeaderFooter { blocks: doc.blocks })
}

enum ParseState {
    Idle,
    InBody,
    InParagraph {
        props: ParagraphProperties,
        runs: Vec<Inline>,
        floats: Vec<FloatingImage>,
        section_props: Option<SectionProperties>,
        /// URL of the containing w:hyperlink element, if any.
        hyperlink_url: Option<String>,
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
        shading: Option<Color>,
    },
    InDrawing {
        depth: u32,
        rel_id: Option<RelId>,
        width_emu: Option<u64>,
        height_emu: Option<u64>,
        is_anchor: bool,
        pos_h_emu: Option<i64>,
        pos_v_emu: Option<i64>,
        align_h: Option<String>,
        align_v: Option<String>,
        /// Tracks whether we're reading text for posOffset ('H'/'V') or align ('h'/'v').
        reading_align: Option<char>,
        wrap_side: Option<WrapSide>,
        reading_pos_offset: Option<char>,
        in_position_h: bool,
        in_position_v: bool,
        /// wp14:pctPosHOffset value (percentage × 1000).
        pct_pos_h: Option<i32>,
        /// wp14:pctPosVOffset value (percentage × 1000).
        pct_pos_v: Option<i32>,
        /// Tracks whether reading text for pctPosHOffset ('H') or pctPosVOffset ('V').
        reading_pct_pos: Option<char>,
    },
}


#[cfg(test)]
mod tests;
