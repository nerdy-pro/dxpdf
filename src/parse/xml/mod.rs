mod drawing;
mod helpers;
mod properties;
mod section;

use std::collections::{HashMap, HashSet};

use log::warn;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader;

use crate::error::Error;
use crate::model::*;

use drawing::{handle_drawing_element, handle_drawing_end};
use helpers::*;
use properties::handle_empty_element;
use section::handle_section_element;

fn warn_once(warned: &mut HashSet<&'static str>, key: &'static str, msg: &str) {
    if warned.insert(key) {
        warn!("{msg}");
    }
}

/// Parse the `word/document.xml` content into a `Document`.
pub fn parse_document_xml(xml: &str) -> Result<Document, Error> {
    let empty = HashMap::new();
    parse_document_xml_with_rels(xml, &empty)
}

/// Parse document XML with relationship data for hyperlink resolution.
pub fn parse_document_xml_with_rels(
    xml: &str,
    rels: &HashMap<String, String>,
) -> Result<Document, Error> {
    let mut reader = Reader::from_str(xml);
    let mut ctx = ParserContext::new();

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            Event::Start(ref e) => ctx.handle_start(e, rels)?,
            Event::Empty(ref e) => ctx.handle_empty(e, rels)?,
            Event::Text(ref e) => ctx.handle_text(e)?,
            Event::End(ref e) => ctx.handle_end(e),
            _ => {}
        }
    }

    Ok(Document {
        blocks: ctx.blocks,
        final_section: ctx.final_section,
        ..Document::default()
    })
}

/// Parse a header or footer XML file into a list of blocks.
/// Header/footer XML has the same structure as document body
/// but with `w:hdr` or `w:ftr` as the root element.
pub fn parse_header_footer_xml(xml: &str) -> Result<HeaderFooter, Error> {
    let empty = HashMap::new();
    parse_header_footer_xml_with_rels(xml, &empty)
}

pub fn parse_header_footer_xml_with_rels(
    xml: &str,
    rels: &HashMap<String, String>,
) -> Result<HeaderFooter, Error> {
    let doc = parse_document_xml_with_rels(xml, rels)?;
    Ok(HeaderFooter { blocks: doc.blocks })
}

// ============================================================
// Parser context — groups all mutable state for the event loop
// ============================================================

struct ParserContext {
    blocks: Vec<Block>,
    stack: Vec<ParseState>,
    state: ParseState,
    final_section: Option<SectionProperties>,
    warned: HashSet<&'static str>,
    /// Field code state machine: tracks begin→instrText→separate→cached→end sequence.
    field_instr: Option<String>,
    field_suppressing: bool,
    field_props: RunProperties,
}

impl ParserContext {
    fn new() -> Self {
        Self {
            blocks: Vec::new(),
            stack: Vec::new(),
            state: ParseState::Idle,
            final_section: None,
            warned: HashSet::new(),
            field_instr: None,
            field_suppressing: false,
            field_props: RunProperties::default(),
        }
    }

    fn handle_start(&mut self, e: &BytesStart<'_>, rels: &HashMap<String, String>) -> Result<(), Error> {
        let name = e.name();
        let local = local_name(name.as_ref());
        match local {
            b"body" | b"hdr" | b"ftr" => {
                self.state = ParseState::InBody;
            }
            b"p" if matches_body_or_cell(&self.state) => {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InParagraph {
                    props: ParagraphProperties::default(),
                    runs: Vec::new(),
                    floats: Vec::new(),
                    section_props: None,
                    hyperlink_url: None,
                };
            }
            b"pPr" if matches!(self.state, ParseState::InParagraph { .. }) => {
                let (props, runs, floats, section_props) = take_paragraph(&mut self.state);
                self.stack.push(ParseState::InParagraph {
                    props: ParagraphProperties::default(),
                    runs,
                    floats,
                    section_props,
                    hyperlink_url: None,
                });
                self.state = ParseState::InParagraphProperties {
                    props,
                    section_props: None,
                    in_pbdr: false,
                };
            }
            b"hyperlink" if matches!(self.state, ParseState::InParagraph { .. }) => {
                if let ParseState::InParagraph { ref mut hyperlink_url, .. } = self.state {
                    let url = get_attr(e, b"id")?.and_then(|rid| rels.get(&rid).cloned());
                    *hyperlink_url = url;
                }
            }
            b"r" if matches!(self.state, ParseState::InParagraph { .. }) => {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InRun {
                    props: RunProperties::default(),
                    text: String::new(),
                };
            }
            b"rPr" if matches!(self.state, ParseState::InRun { .. }) => {
                let (props, text) = take_run(&mut self.state);
                self.stack.push(ParseState::InRun {
                    props: RunProperties::default(),
                    text,
                });
                self.state = ParseState::InRunProperties { props };
            }
            b"tbl" if matches_body_or_cell(&self.state) => {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InTable {
                    rows: Vec::new(),
                    grid_cols: Vec::new(),
                    default_cell_margins: None,
                    in_cell_mar: false,
                    borders: None,
                    in_borders: false,
                };
            }
            b"tr" if matches!(self.state, ParseState::InTable { .. }) => {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InTableRow {
                    cells: Vec::new(),
                    height: None,
                };
            }
            b"tc" if matches!(self.state, ParseState::InTableRow { .. }) => {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InTableCell {
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
            b"t" if matches!(self.state, ParseState::InRun { .. }) => {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InText {
                    text: String::new(),
                };
            }
            b"drawing"
                if matches!(
                    self.state,
                    ParseState::InRun { .. } | ParseState::InParagraph { .. }
                ) =>
            {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InDrawing {
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
            _ if matches!(self.state, ParseState::InSectionProperties { .. }) => {
                handle_section_element(local, e, &mut self.state)?;
            }
            _ if matches!(self.state, ParseState::InDrawing { .. }) => {
                if let ParseState::InDrawing { ref mut depth, .. } = self.state {
                    *depth += 1;
                }
                handle_drawing_element(local, e, &mut self.state)?;
            }
            b"pBdr" if matches!(self.state, ParseState::InParagraphProperties { .. }) => {
                if let ParseState::InParagraphProperties { ref mut in_pbdr, .. } = self.state {
                    *in_pbdr = true;
                }
            }
            b"sectPr"
                if matches!(
                    self.state,
                    ParseState::InParagraphProperties { .. } | ParseState::InBody
                ) =>
            {
                self.stack.push(std::mem::replace(&mut self.state, ParseState::Idle));
                self.state = ParseState::InSectionProperties {
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
            b"tblCellMar" | b"tcMar" => set_in_cell_mar(&mut self.state, true),
            b"tblBorders" | b"tcBorders" => set_in_borders(&mut self.state, true),
            b"pict" | b"object" => {
                warn_once(&mut self.warned, "pict", "Unsupported: VML image/object (w:pict/w:object) — use DrawingML (w:drawing) instead");
            }
            b"fldChar" => {
                handle_fld_char(e, &mut self.state, &mut self.stack, &mut self.field_instr, &mut self.field_suppressing, &mut self.field_props)?;
            }
            b"instrText" => {
                // instrText content will be captured in handle_text
            }
            b"footnoteReference" => {
                warn_once(&mut self.warned, "footnote", "Unsupported: footnote reference (w:footnoteReference)");
            }
            b"endnoteReference" => {
                warn_once(&mut self.warned, "endnote", "Unsupported: endnote reference (w:endnoteReference)");
            }
            b"ins" | b"moveTo" => {
                // Tracked insertions — content inside is valid, just not marked as tracked
            }
            b"del" | b"moveFrom" => {
                warn_once(&mut self.warned, "del", "Unsupported: tracked deletion (w:del/w:moveFrom) — deleted content may appear");
            }
            b"commentRangeStart" | b"commentRangeEnd" | b"commentReference" => {}
            b"bookmarkStart" | b"bookmarkEnd" | b"proofErr" | b"lastRenderedPageBreak" => {}
            _ => {}
        }
        Ok(())
    }

    fn handle_empty(&mut self, e: &BytesStart<'_>, rels: &HashMap<String, String>) -> Result<(), Error> {
        let _ = rels; // only used by handle_start for hyperlinks
        let name = e.name();
        let local = local_name(name.as_ref());
        if local == b"p" && matches_body_or_cell(&self.state) {
            let paragraph = Block::Paragraph(Box::new(Paragraph {
                properties: ParagraphProperties::default(),
                runs: Vec::new(),
                floats: Vec::new(),
                section_properties: None,
            }));
            push_block(&mut self.state, &mut self.blocks, paragraph);
        } else if (local == b"br" || local == b"tab")
            && matches!(self.state, ParseState::InRun { .. })
        {
            if let ParseState::InRun { ref props, ref mut text, .. } = self.state {
                let flushed_text = std::mem::take(text);
                let inline_to_push = if local == b"br" {
                    Inline::LineBreak
                } else {
                    Inline::Tab
                };
                if let Some(para_state) = self.stack
                    .iter_mut()
                    .rev()
                    .find(|s| matches!(s, ParseState::InParagraph { .. }))
                {
                    if let ParseState::InParagraph { ref mut runs, .. } = para_state {
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
        } else if local == b"fldChar" {
            handle_fld_char(e, &mut self.state, &mut self.stack, &mut self.field_instr, &mut self.field_suppressing, &mut self.field_props)?;
        } else if matches!(self.state, ParseState::InDrawing { .. }) {
            handle_drawing_element(local, e, &mut self.state)?;
        } else if matches!(self.state, ParseState::InSectionProperties { .. }) {
            handle_section_element(local, e, &mut self.state)?;
        } else {
            handle_empty_element(local, e, &mut self.state, &mut self.warned)?;
        }
        Ok(())
    }

    fn handle_text(&mut self, e: &BytesText<'_>) -> Result<(), Error> {
        // Field code handling: capture instruction text or suppress cached value
        if self.field_instr.is_some() && !self.field_suppressing {
            if let Some(ref mut instr) = self.field_instr {
                instr.push_str(&e.unescape().unwrap_or_default());
            }
            return Ok(());
        }
        if self.field_suppressing {
            if let Some(ref instr) = self.field_instr {
                let trimmed = instr.trim().to_uppercase();
                if trimmed.starts_with("PAGE") || trimmed.starts_with("NUMPAGES") {
                    return Ok(());
                }
            }
            // Unknown field: let cached text through as normal
        }

        if let ParseState::InText { ref mut text, .. } = self.state {
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
        } = self.state
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
                    let emu = crate::dimension::Emu::new(val);
                    match axis {
                        'H' => *pos_h_emu = Some(emu),
                        'V' => *pos_v_emu = Some(emu),
                        _ => {}
                    }
                }
            }
        }
        Ok(())
    }

    fn handle_end(&mut self, e: &BytesEnd<'_>) {
        let name = e.name();
        let local = local_name(name.as_ref());
        match local {
            b"body" | b"hdr" | b"ftr" => {
                self.state = ParseState::Idle;
            }
            b"hyperlink" if matches!(self.state, ParseState::InParagraph { .. }) => {
                if let ParseState::InParagraph { ref mut hyperlink_url, .. } = self.state {
                    *hyperlink_url = None;
                }
            }
            b"p" if matches!(self.state, ParseState::InParagraph { .. }) => {
                let (props, runs, floats, section_props) = take_paragraph(&mut self.state);
                self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                let paragraph = Block::Paragraph(Box::new(Paragraph {
                    properties: props,
                    runs,
                    floats,
                    section_properties: section_props,
                }));
                push_block(&mut self.state, &mut self.blocks, paragraph);
            }
            b"pBdr" if matches!(self.state, ParseState::InParagraphProperties { .. }) => {
                if let ParseState::InParagraphProperties { ref mut in_pbdr, .. } = self.state {
                    *in_pbdr = false;
                }
            }
            b"pPr" if matches!(self.state, ParseState::InParagraphProperties { .. }) => {
                if let ParseState::InParagraphProperties { props, section_props, .. } = std::mem::replace(&mut self.state, ParseState::Idle) {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    if let ParseState::InParagraph {
                        props: ref mut p,
                        section_props: ref mut sp,
                        ..
                    } = self.state
                    {
                        *p = props;
                        *sp = section_props;
                    }
                }
            }
            b"r" if matches!(self.state, ParseState::InRun { .. }) => {
                let (props, text) = take_run(&mut self.state);
                self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                if !text.is_empty() {
                    if let ParseState::InParagraph { ref mut runs, ref hyperlink_url, .. } = self.state {
                        runs.push(Inline::TextRun(TextRun {
                            text,
                            properties: props,
                            hyperlink_url: hyperlink_url.clone(),
                        }));
                    }
                }
            }
            b"rPr" if matches!(self.state, ParseState::InRunProperties { .. }) => {
                if let ParseState::InRunProperties { props } = std::mem::replace(&mut self.state, ParseState::Idle) {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    if let ParseState::InRun { props: ref mut p, .. } = self.state {
                        *p = props;
                    }
                }
            }
            b"t" if matches!(self.state, ParseState::InText { .. }) => {
                if let ParseState::InText { text, .. } = std::mem::replace(&mut self.state, ParseState::Idle) {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    if let ParseState::InRun { text: ref mut t, .. } = self.state {
                        let cleaned: String =
                            text.chars().filter(|c| *c != '\n' && *c != '\r').collect();
                        t.push_str(&cleaned);
                    }
                }
            }
            b"tbl" if matches!(self.state, ParseState::InTable { .. }) => {
                if let ParseState::InTable { rows, grid_cols, default_cell_margins, borders, .. } = std::mem::replace(&mut self.state, ParseState::Idle) {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    let table = Block::Table(Box::new(Table {
                        rows,
                        grid_cols,
                        default_cell_margins,
                        cell_spacing: None,
                        borders,
                    }));
                    push_block(&mut self.state, &mut self.blocks, table);
                }
            }
            b"tr" if matches!(self.state, ParseState::InTableRow { .. }) => {
                if let ParseState::InTableRow { cells, height } = std::mem::replace(&mut self.state, ParseState::Idle) {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    if let ParseState::InTable { ref mut rows, .. } = self.state {
                        rows.push(TableRow { cells, height });
                    }
                }
            }
            b"tc" if matches!(self.state, ParseState::InTableCell { .. }) => {
                if let ParseState::InTableCell {
                    blocks: cell_blocks,
                    width: cell_width,
                    grid_span,
                    vertical_merge,
                    cell_margins,
                    cell_borders,
                    shading,
                    ..
                } = std::mem::replace(&mut self.state, ParseState::Idle)
                {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    if let ParseState::InTableRow { ref mut cells, .. } = self.state {
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
            b"drawing" if matches!(self.state, ParseState::InDrawing { .. }) => {
                let old = std::mem::replace(&mut self.state, ParseState::Idle);
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
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    if let Some(rid) = rel_id {
                        use crate::dimension::{Emu, Pt};
                        let zero = Emu::new(0);
                        let w = Pt::from(width_emu.unwrap_or(zero));
                        let h = Pt::from(height_emu.unwrap_or(zero));

                        if is_anchor {
                            let float = FloatingImage {
                                rel_id: rid,
                                size: crate::geometry::PtSize::new(w, h),
                                offset: crate::geometry::PtOffset::new(
                                    Pt::from(pos_h_emu.unwrap_or(zero)),
                                    Pt::from(pos_v_emu.unwrap_or(zero)),
                                ),
                                align_h,
                                align_v,
                                wrap_side: wrap_side.unwrap_or(WrapSide::BothSides),
                                pct_pos_h,
                                pct_pos_v,
                            };
                            push_float(&mut self.state, &mut self.stack, float);
                        } else {
                            let image = Inline::Image(InlineImage {
                                rel_id: rid,
                                size: crate::geometry::PtSize::new(w, h),
                            });
                            push_inline_to_paragraph(&mut self.state, &mut self.stack, image);
                        }
                    }
                }
            }
            b"sectPr" if matches!(self.state, ParseState::InSectionProperties { .. }) => {
                if let ParseState::InSectionProperties { section } = std::mem::replace(&mut self.state, ParseState::Idle) {
                    self.state = self.stack.pop().unwrap_or(ParseState::Idle);
                    match self.state {
                        ParseState::InParagraphProperties { ref mut section_props, .. } => {
                            *section_props = Some(section);
                        }
                        ParseState::InBody => {
                            self.final_section = Some(section);
                        }
                        _ => {}
                    }
                }
            }
            _ if matches!(self.state, ParseState::InDrawing { .. }) => {
                handle_drawing_end(local, &mut self.state);
                if let ParseState::InDrawing { ref mut depth, .. } = self.state {
                    *depth -= 1;
                }
            }
            b"tblCellMar" | b"tcMar" => set_in_cell_mar(&mut self.state, false),
            b"tblBorders" | b"tcBorders" => set_in_borders(&mut self.state, false),
            _ => {}
        }
    }
}

// ============================================================
// ParseState enum
// ============================================================

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
        in_pbdr: bool,
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
        grid_cols: Vec<crate::dimension::Twips>,
        default_cell_margins: Option<CellMargins>,
        in_cell_mar: bool,
        borders: Option<TableBorders>,
        in_borders: bool,
    },
    InTableRow {
        cells: Vec<TableCell>,
        height: Option<crate::dimension::Twips>,
    },
    InTableCell {
        blocks: Vec<Block>,
        width: Option<crate::dimension::Twips>,
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
        width_emu: Option<crate::dimension::Emu>,
        height_emu: Option<crate::dimension::Emu>,
        is_anchor: bool,
        pos_h_emu: Option<crate::dimension::Emu>,
        pos_v_emu: Option<crate::dimension::Emu>,
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
