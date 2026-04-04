//! Parser for document body content: blocks (paragraphs, tables, section breaks)
//! and inline content (text runs, images, hyperlinks, fields, etc.).
//!
//! No style resolution or property merging — output is raw parsed data.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::field::FieldInstruction;

use crate::docx::dimension::Dimension;
use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

use super::{drawing, properties, vml};

/// Parse a body-level XML part (header, footer) into blocks.
/// Finds the root element (§17.10.1 `w:hdr`, §17.10.3 `w:ftr`, etc.),
/// then parses block content scoped to that element.
pub fn parse_blocks(data: &[u8]) -> Result<Vec<Block>> {
    let mut reader = Reader::from_reader(data);
    // Do NOT trim text — whitespace in <w:t> runs is significant.
    let mut buf = Vec::new();

    // Find the root element and remember its local name.
    let root_tag = loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                break xml::local_name(e.name().as_ref()).to_vec();
            }
            Event::Eof => return Ok(Vec::new()),
            _ => {}
        }
    };

    // Parse block content scoped to the root element.
    let (blocks, _) = parse_block_content(&mut reader, &mut buf, &root_tag)?;
    Ok(blocks)
}

/// Parse `w:document > w:body`, returning blocks and final section properties.
/// Scoped: enters `<w:document>`, then `<w:body>`, breaks on `</w:body>`.
pub fn parse_body(data: &[u8]) -> Result<(Vec<Block>, SectionProperties)> {
    let mut reader = Reader::from_reader(data);
    // Do NOT trim text — whitespace in <w:t> runs is significant.
    let mut buf = Vec::new();

    // Find <w:body> inside <w:document>.
    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"body" => break,
            Event::Eof => return Ok((Vec::new(), SectionProperties::default())),
            _ => {}
        }
    }

    // Parse block content scoped to <w:body>.
    parse_block_content(&mut reader, &mut buf, b"body")
}

/// Parse blocks until `</end_tag>`. Returns blocks and the final `w:sectPr`
/// (if one appears as a direct child — per §17.6.17 the last sectPr in body).
pub fn parse_block_content(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<(Vec<Block>, SectionProperties)> {
    let mut blocks = Vec::new();
    let mut final_section = SectionProperties::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"p" => {
                        let rsids = parse_paragraph_rsids(e)?;
                        let (para, sect) = parse_paragraph_inner(reader, buf, rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" => {
                        blocks.push(Block::Table(Box::new(parse_table(reader, buf)?)));
                    }
                    b"sectPr" => {
                        let sect_rsids = properties::parse_section_rsids(e)?;
                        final_section = properties::parse_section_properties(reader, buf)?;
                        final_section.rsids = sect_rsids;
                    }
                    _ => xml::warn_unsupported_element("body", local),
                }
            }
            // Self-closing elements: <w:p/> is an empty paragraph.
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                if local == b"p" {
                    let rsids = parse_paragraph_rsids(e)?;
                    blocks.push(Block::Paragraph(Box::new(Paragraph {
                        style_id: None,
                        properties: ParagraphProperties::default(),
                        mark_run_properties: None,
                        content: vec![],
                        rsids,
                    })));
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(b"container")),
            _ => {}
        }
    }

    Ok((blocks, final_section))
}

// ── Paragraph ────────────────────────────────────────────────────────────────

/// Returns the paragraph and optionally a section break that follows it
/// (per §17.6.18, sectPr inside pPr means this paragraph ends a section).
fn parse_paragraph_inner(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    rsids: ParagraphRevisionIds,
) -> Result<(Paragraph, Option<SectionProperties>)> {
    let mut para_props = ParagraphProperties::default();
    let mut mark_run_props: Option<RunProperties> = None;
    let mut style_id: Option<StyleId> = None;
    let mut section_props: Option<SectionProperties> = None;
    let mut content = Vec::new();

    loop {
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"bookmarkStart" => {
                        if let (Some(id), Some(name)) = (
                            xml::optional_attr_i64(e, b"id")?,
                            xml::optional_attr(e, b"name")?,
                        ) {
                            content.push(Inline::BookmarkStart {
                                id: BookmarkId::new(id),
                                name,
                            });
                        }
                    }
                    b"bookmarkEnd" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            content.push(Inline::BookmarkEnd(BookmarkId::new(id)));
                        }
                    }
                    // Start-only: have child elements
                    b"pPr" if is_start => {
                        let parsed = properties::parse_paragraph_properties(reader, buf)?;
                        para_props = parsed.properties;
                        style_id = parsed.style_id;
                        mark_run_props = parsed.run_properties;
                        section_props = parsed.section_properties;
                    }
                    b"r" if is_start => {
                        let run_rsids = parse_run_rsids(e)?;
                        parse_run(reader, buf, &mut content, run_rsids)?;
                    }
                    b"hyperlink" if is_start => {
                        let r_id = xml::optional_attr(e, b"id")?;
                        let anchor = xml::optional_attr(e, b"anchor")?;
                        let hyperlink = parse_hyperlink_content(r_id, anchor, reader, buf)?;
                        content.push(Inline::Hyperlink(hyperlink));
                    }
                    b"fldSimple" if is_start => {
                        let instr = xml::optional_attr(e, b"instr")?.unwrap_or_default();
                        let field = parse_simple_field_content(&instr, reader, buf)?;
                        content.push(Inline::Field(field));
                    }
                    // Empty self-closing <w:r/> carries only revision IDs
                    // (rsidR, rsidRPr) — no text or inline content to render.
                    b"r" => {}
                    _ => xml::warn_unsupported_element("paragraph", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"p" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"p")),
            _ => {}
        }
    }

    Ok((
        Paragraph {
            style_id,
            properties: para_props,
            mark_run_properties: mark_run_props,
            content,
            rsids,
        },
        section_props,
    ))
}

/// Parse a `<w:p>` paragraph element, extracting revision IDs from the start tag.
pub fn parse_paragraph(
    start_event: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(Paragraph, Option<SectionProperties>)> {
    let rsids = parse_paragraph_rsids(start_event)?;
    parse_paragraph_inner(reader, buf, rsids)
}

// ── Run ──────────────────────────────────────────────────────────────────────

/// Parse a `w:br` element into a `RunElement`.
///
/// §17.3.3.1: `w:type` selects page/column/textWrapping; `w:clear` selects
/// which float side to clear for textWrapping breaks.
fn parse_run_break(e: &BytesStart<'_>) -> Result<RunElement> {
    let br_type = xml::optional_attr(e, b"type")?;
    Ok(match br_type.as_deref() {
        Some("page") => RunElement::PageBreak,
        Some("column") => RunElement::ColumnBreak,
        _ => {
            let clear = match xml::optional_attr(e, b"clear")?.as_deref() {
                Some("left") => BreakClear::Left,
                Some("right") => BreakClear::Right,
                Some("all") => BreakClear::All,
                _ => BreakClear::None,
            };
            let kind = if clear != BreakClear::None {
                BreakKind::Clear(clear)
            } else {
                BreakKind::TextWrapping
            };
            RunElement::LineBreak(kind)
        }
    })
}

/// Parse a `w:fldChar` element into a `FieldChar`.
///
/// §17.16.18: `w:fldCharType` is required and must be `begin`, `separate`, or `end`.
fn parse_fld_char(e: &BytesStart<'_>) -> Result<FieldChar> {
    let field_char_type = match xml::optional_attr(e, b"fldCharType")?.as_deref() {
        Some("begin") => FieldCharType::Begin,
        Some("separate") => FieldCharType::Separate,
        Some("end") => FieldCharType::End,
        Some(other) => {
            return Err(crate::docx::error::ParseError::InvalidAttributeValue {
                attr: "fldChar/fldCharType".into(),
                value: other.into(),
                reason: "expected begin, separate, or end per §17.18.29".into(),
            });
        }
        None => {
            return Err(crate::docx::error::ParseError::MissingAttribute {
                element: "fldChar".into(),
                attr: "fldCharType".into(),
            });
        }
    };
    Ok(FieldChar {
        field_char_type,
        dirty: xml::optional_attr_bool(e, b"dirty")?,
        fld_lock: xml::optional_attr_bool(e, b"fldLock")?,
    })
}

/// Accumulates run elements and flushes them as a `TextRun` inline.
struct RunAccumulator<'a> {
    elements: Vec<RunElement>,
    style_id: Option<StyleId>,
    props: RunProperties,
    rsids: RevisionIds,
    content: &'a mut Vec<Inline>,
}

impl RunAccumulator<'_> {
    /// Flush accumulated run elements into a TextRun and push to content.
    fn flush(&mut self) {
        if !self.elements.is_empty() {
            self.content.push(Inline::TextRun(Box::new(TextRun {
                style_id: self.style_id.clone(),
                properties: self.props.clone(),
                content: std::mem::take(&mut self.elements),
                rsids: self.rsids,
            })));
        }
    }

    /// Flush, then push a non-run inline directly to content.
    fn flush_and_push(&mut self, inline: Inline) {
        self.flush();
        self.content.push(inline);
    }

    /// Consume the accumulator, flushing remaining elements by move (no clone).
    fn finish(self) {
        if !self.elements.is_empty() {
            self.content.push(Inline::TextRun(Box::new(TextRun {
                style_id: self.style_id,
                properties: self.props,
                content: self.elements,
                rsids: self.rsids,
            })));
        }
    }
}

fn parse_run(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    content: &mut Vec<Inline>,
    run_rsids: RevisionIds,
) -> Result<()> {
    let mut acc = RunAccumulator {
        elements: Vec::new(),
        style_id: None,
        props: RunProperties::default(),
        rsids: run_rsids,
        content,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"rPr" => {
                        let (rp, sid) = properties::parse_run_properties(reader, buf)?;
                        acc.props = rp;
                        acc.style_id = sid;
                    }
                    b"t" | b"delText" => {
                        let text = xml::read_text_content(reader, buf)?;
                        acc.elements.push(RunElement::Text(text));
                    }
                    // Non-run inlines: flush current run elements, then push separately.
                    b"instrText" => {
                        let text = xml::read_text_content(reader, buf)?;
                        acc.flush_and_push(Inline::InstrText(text));
                    }
                    b"drawing" => {
                        acc.flush();
                        if let Some(img) = parse_drawing(reader, buf)? {
                            acc.content.push(Inline::Image(Box::new(img)));
                        }
                    }
                    b"pict" => {
                        acc.flush_and_push(Inline::Pict(vml::parse_pict(reader, buf)?));
                    }
                    b"AlternateContent" => {
                        acc.flush_and_push(Inline::AlternateContent(parse_alternate_content(
                            reader, buf,
                        )?));
                    }
                    _ => xml::warn_unsupported_element("run", local),
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"rPr" => {}
                    // Run-level elements: accumulate in the current run.
                    b"tab" => acc.elements.push(RunElement::Tab),
                    b"br" => acc.elements.push(parse_run_break(e)?),
                    b"cr" => acc
                        .elements
                        .push(RunElement::LineBreak(BreakKind::TextWrapping)),
                    b"lastRenderedPageBreak" => {
                        acc.elements.push(RunElement::LastRenderedPageBreak)
                    }
                    // Non-run inlines: flush, then push separately.
                    b"sym" => {
                        acc.flush();
                        let font = xml::optional_attr(e, b"font")?.unwrap_or_default();
                        let char_str = xml::optional_attr(e, b"char")?.unwrap_or_default();
                        let char_code = u16::from_str_radix(&char_str, 16).unwrap_or_else(|e| {
                            log::warn!("w:sym: invalid char code {char_str:?}: {e}");
                            0
                        });
                        acc.content.push(Inline::Symbol(Symbol { font, char_code }));
                    }
                    b"footnoteReference" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            acc.flush_and_push(Inline::FootnoteRef(NoteId::new(id)));
                        }
                    }
                    b"endnoteReference" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            acc.flush_and_push(Inline::EndnoteRef(NoteId::new(id)));
                        }
                    }
                    b"separator" => acc.flush_and_push(Inline::Separator),
                    b"continuationSeparator" => acc.flush_and_push(Inline::ContinuationSeparator),
                    b"footnoteRef" => acc.flush_and_push(Inline::FootnoteRefMark),
                    b"endnoteRef" => acc.flush_and_push(Inline::EndnoteRefMark),
                    b"fldChar" => {
                        acc.flush_and_push(Inline::FieldChar(parse_fld_char(e)?));
                    }
                    _ => xml::warn_unsupported_element("run", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"r" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"r")),
            _ => {}
        }
    }

    acc.finish();
    Ok(())
}

// ── Hyperlink ────────────────────────────────────────────────────────────────

fn parse_hyperlink_content(
    r_id: Option<String>,
    anchor: Option<String>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Hyperlink> {
    let target = if let Some(id) = r_id {
        HyperlinkTarget::External(RelId::new(id))
    } else if let Some(anchor) = anchor {
        HyperlinkTarget::Internal { anchor }
    } else {
        HyperlinkTarget::Internal {
            anchor: String::new(),
        }
    };

    let mut inline_content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"r" => {
                let r_rsids = parse_run_rsids(e)?;
                parse_run(reader, buf, &mut inline_content, r_rsids)?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"hyperlink" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"hyperlink")),
            _ => {}
        }
    }

    Ok(Hyperlink {
        target,
        content: inline_content,
    })
}

// ── Field ────────────────────────────────────────────────────────────────────

fn parse_simple_field_content(
    instruction: &str,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Field> {
    let mut field_content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"r" => {
                let r_rsids = parse_run_rsids(e)?;
                parse_run(reader, buf, &mut field_content, r_rsids)?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"fldSimple" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"fldSimple")),
            _ => {}
        }
    }

    let parsed = parse_field_instruction(instruction);

    Ok(Field {
        instruction: parsed,
        content: field_content,
    })
}

/// Parse a raw field instruction string into a typed `FieldInstruction`.
/// Falls back to `FieldInstruction::Unknown` on parse errors.
fn parse_field_instruction(raw: &str) -> FieldInstruction {
    match crate::field::parse(raw) {
        Ok(instr) => instr,
        Err(e) => {
            log::warn!("failed to parse field instruction {:?}: {}", raw, e);
            FieldInstruction::Unknown {
                field_type: String::new(),
                raw: raw.to_owned(),
            }
        }
    }
}

// VML / Pict parsing has been extracted to `super::vml`.

// ── Alternate Content ───────────────────────────────────────────────────────

/// MCE §M.2.1: parse `mc:AlternateContent`.
fn parse_alternate_content(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<AlternateContent> {
    let mut choices = Vec::new();
    let mut fallback = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"Choice" => {
                        let requires_str = xml::optional_attr(e, b"Requires")?.unwrap_or_default();
                        let requires = match requires_str.as_str() {
                            "wps" => McRequires::Wps,
                            "wpg" => McRequires::Wpg,
                            "wpc" => McRequires::Wpc,
                            "wpi" => McRequires::Wpi,
                            "m" => McRequires::Math,
                            "a14" => McRequires::A14,
                            "w14" => McRequires::W14,
                            "w15" => McRequires::W15,
                            "w16" => McRequires::W16,
                            other => {
                                log::warn!("mc:Choice: unsupported Requires {:?}", other);
                                xml::skip_to_end(reader, buf, b"Choice")?;
                                continue;
                            }
                        };
                        let content = parse_run_inline_content(reader, buf, b"Choice")?;
                        choices.push(McChoice { requires, content });
                    }
                    b"Fallback" => {
                        fallback = Some(parse_run_inline_content(reader, buf, b"Fallback")?);
                    }
                    _ => {
                        xml::warn_unsupported_element("AlternateContent", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"AlternateContent" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"AlternateContent")),
            _ => {}
        }
    }

    Ok(AlternateContent { choices, fallback })
}

/// Parse run-level inline content inside a container (mc:Choice, mc:Fallback).
/// These containers hold the same elements as direct children of `w:r`:
/// `w:drawing`, `w:pict`, etc.
fn parse_run_inline_content(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<Vec<Inline>> {
    let mut content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"drawing" => {
                        if let Some(img) = parse_drawing(reader, buf)? {
                            content.push(Inline::Image(Box::new(img)));
                        }
                    }
                    b"pict" => {
                        content.push(Inline::Pict(vml::parse_pict(reader, buf)?));
                    }
                    _ => {
                        xml::warn_unsupported_element("mc-content", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }

    Ok(content)
}

// ── Drawing / Image ──────────────────────────────────────────────────────────

fn parse_drawing(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<Image>> {
    let mut image: Option<Image> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"inline" => {
                        image = drawing::parse_inline_image(e, reader, buf)?;
                    }
                    b"anchor" => {
                        image = drawing::parse_anchor_image(e, reader, buf)?;
                    }
                    _ => xml::warn_unsupported_element("drawing", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"drawing" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"drawing")),
            _ => {}
        }
    }

    Ok(image)
}

// ── Table ────────────────────────────────────────────────────────────────────

/// Parse a `<w:tbl>` table element.
pub fn parse_table(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Table> {
    let mut tbl_props = TableProperties::default();
    let mut grid = Vec::new();
    let mut rows = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tblPr" => {
                        let (tp, _) = properties::parse_table_properties(reader, buf)?;
                        tbl_props = tp;
                    }
                    b"tblGrid" => {
                        grid = parse_table_grid(reader, buf)?;
                    }
                    b"tr" => {
                        let tr_rsids = parse_table_row_rsids(e)?;
                        rows.push(parse_table_row(reader, buf, tr_rsids)?);
                    }
                    _ => xml::warn_unsupported_element("table", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tbl" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tbl")),
            _ => {}
        }
    }

    Ok(Table {
        properties: tbl_props,
        grid,
        rows,
    })
}

fn parse_table_grid(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<GridColumn>> {
    let mut cols = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e)
                if xml::local_name(e.name().as_ref()) == b"gridCol" =>
            {
                let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
                cols.push(GridColumn {
                    width: Dimension::new(w),
                });
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblGrid" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tblGrid")),
            _ => {}
        }
    }

    Ok(cols)
}

fn parse_table_row(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    row_rsids: TableRowRevisionIds,
) -> Result<TableRow> {
    let mut row_props = TableRowProperties::default();
    let mut cells = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"trPr" => {
                        row_props = properties::parse_table_row_properties(reader, buf)?;
                    }
                    b"tc" => {
                        cells.push(parse_table_cell(reader, buf)?);
                    }
                    _ => xml::warn_unsupported_element("table-row", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tr")),
            _ => {}
        }
    }

    Ok(TableRow {
        properties: row_props,
        cells,
        rsids: row_rsids,
    })
}

fn parse_table_cell(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<TableCell> {
    let mut cell_props = TableCellProperties::default();
    let mut blocks = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tcPr" => {
                        cell_props = properties::parse_table_cell_properties(reader, buf)?;
                    }
                    b"p" => {
                        let p_rsids = parse_paragraph_rsids(e)?;
                        let (para, sect) = parse_paragraph_inner(reader, buf, p_rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" => {
                        blocks.push(Block::Table(Box::new(parse_table(reader, buf)?)));
                    }
                    _ => xml::warn_unsupported_element("table-cell", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tc" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tc")),
            _ => {}
        }
    }

    Ok(TableCell {
        properties: cell_props,
        content: blocks,
    })
}

// ── Rsid extraction ──────────────────────────────────────────────────────────

fn parse_paragraph_rsids(e: &BytesStart<'_>) -> Result<ParagraphRevisionIds> {
    Ok(ParagraphRevisionIds {
        r: xml::optional_rsid(e, b"rsidR")?,
        r_default: xml::optional_rsid(e, b"rsidRDefault")?,
        p: xml::optional_rsid(e, b"rsidP")?,
        r_pr: xml::optional_rsid(e, b"rsidRPr")?,
        del: xml::optional_rsid(e, b"rsidDel")?,
    })
}

fn parse_run_rsids(e: &BytesStart<'_>) -> Result<RevisionIds> {
    Ok(RevisionIds {
        r: xml::optional_rsid(e, b"rsidR")?,
        r_pr: xml::optional_rsid(e, b"rsidRPr")?,
        del: xml::optional_rsid(e, b"rsidDel")?,
    })
}

fn parse_table_row_rsids(e: &BytesStart<'_>) -> Result<TableRowRevisionIds> {
    Ok(TableRowRevisionIds {
        r: xml::optional_rsid(e, b"rsidR")?,
        r_pr: xml::optional_rsid(e, b"rsidRPr")?,
        del: xml::optional_rsid(e, b"rsidDel")?,
        tr: xml::optional_rsid(e, b"rsidTr")?,
    })
}
