//! Parser for document body content: blocks (paragraphs, tables, section breaks)
//! and inline content (text runs, images, hyperlinks, fields, etc.).
//!
//! No style resolution or property merging — output is raw parsed data.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use dxpdf_field::FieldInstruction;

use crate::dimension::Dimension;
use crate::error::Result;
use crate::model::*;
use crate::xml;

use super::{drawing, properties};

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
fn parse_block_content(
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
                        let (para, sect) = parse_paragraph(reader, buf, rsids)?;
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
fn parse_paragraph(
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
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"pPr" => {
                        let parsed = properties::parse_paragraph_properties(reader, buf)?;
                        para_props = parsed.properties;
                        style_id = parsed.style_id;
                        mark_run_props = parsed.run_properties;
                        section_props = parsed.section_properties;
                    }
                    b"r" => {
                        let run_rsids = parse_run_rsids(e)?;
                        parse_run(reader, buf, &mut content, run_rsids)?;
                    }
                    b"hyperlink" => {
                        let r_id = xml::optional_attr(e, b"id")?;
                        let anchor = xml::optional_attr(e, b"anchor")?;
                        let hyperlink = parse_hyperlink_content(r_id, anchor, reader, buf)?;
                        content.push(Inline::Hyperlink(hyperlink));
                    }
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
                    b"fldSimple" => {
                        let instr = xml::optional_attr(e, b"instr")?.unwrap_or_default();
                        let field = parse_simple_field_content(&instr, reader, buf)?;
                        content.push(Inline::Field(field));
                    }
                    _ => xml::warn_unsupported_element("paragraph", local),
                }
            }
            Event::Empty(ref e) => {
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

/// Public wrapper for cross-module use (notes.rs).
pub fn parse_paragraph_public(
    start_event: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(Paragraph, Option<SectionProperties>)> {
    let rsids = parse_paragraph_rsids(start_event)?;
    parse_paragraph(reader, buf, rsids)
}

/// Public wrapper for cross-module use (notes.rs).
pub fn parse_table_public(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Table> {
    parse_table(reader, buf)
}

/// Public wrapper for cross-module use (drawing.rs, numbering.rs).
pub fn parse_block_content_public(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<(Vec<Block>, SectionProperties)> {
    parse_block_content(reader, buf, end_tag)
}

// ── Run ──────────────────────────────────────────────────────────────────────

fn parse_run(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    content: &mut Vec<Inline>,
    run_rsids: RevisionIds,
) -> Result<()> {
    let mut run_props = RunProperties::default();
    let mut char_style_id: Option<StyleId> = None;
    let mut texts = Vec::new();
    let mut pending_inlines: Vec<Inline> = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"rPr" => {
                        let (rp, sid) = properties::parse_run_properties(reader, buf)?;
                        run_props = rp;
                        char_style_id = sid;
                    }
                    b"t" | b"delText" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        let text = xml::read_text_content(reader, buf)?;
                        texts.push(text);
                    }
                    b"instrText" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        let text = xml::read_text_content(reader, buf)?;
                        pending_inlines.push(Inline::InstrText(text));
                    }
                    b"drawing" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        if let Some(img) = parse_drawing(reader, buf)? {
                            pending_inlines.push(Inline::Image(Box::new(img)));
                        }
                    }
                    b"pict" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::Pict(parse_pict(reader, buf)?));
                    }
                    b"AlternateContent" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::AlternateContent(
                            parse_alternate_content(reader, buf)?,
                        ));
                    }
                    _ => xml::warn_unsupported_element("run", local),
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"rPr" => {
                        // Empty rPr — no child elements, no properties to parse.
                    }
                    b"tab" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::Tab);
                    }
                    b"br" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        let br_type = xml::optional_attr(e, b"type")?;
                        match br_type.as_deref() {
                            Some("page") => pending_inlines.push(Inline::PageBreak),
                            Some("column") => pending_inlines.push(Inline::ColumnBreak),
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
                                pending_inlines.push(Inline::LineBreak(kind));
                            }
                        }
                    }
                    b"cr" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::LineBreak(BreakKind::TextWrapping));
                    }
                    b"lastRenderedPageBreak" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::LastRenderedPageBreak);
                    }
                    b"sym" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        let font = xml::optional_attr(e, b"font")?.unwrap_or_default();
                        let char_str = xml::optional_attr(e, b"char")?.unwrap_or_default();
                        let char_code = u16::from_str_radix(&char_str, 16).unwrap_or(0);
                        pending_inlines.push(Inline::Symbol(Symbol { font, char_code }));
                    }
                    b"footnoteReference" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            pending_inlines.push(Inline::FootnoteRef(NoteId::new(id)));
                        }
                    }
                    b"endnoteReference" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            pending_inlines.push(Inline::EndnoteRef(NoteId::new(id)));
                        }
                    }
                    b"separator" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::Separator);
                    }
                    b"continuationSeparator" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::ContinuationSeparator);
                    }
                    b"footnoteRef" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::FootnoteRefMark);
                    }
                    b"endnoteRef" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::EndnoteRefMark);
                    }
                    b"fldChar" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        let field_char_type = match xml::optional_attr(e, b"fldCharType")?
                            .as_deref()
                        {
                            Some("begin") => FieldCharType::Begin,
                            Some("separate") => FieldCharType::Separate,
                            Some("end") => FieldCharType::End,
                            Some(other) => {
                                return Err(crate::error::ParseError::InvalidAttributeValue {
                                    attr: "fldChar/fldCharType".into(),
                                    value: other.into(),
                                    reason: "expected begin, separate, or end per §17.18.29".into(),
                                });
                            }
                            None => {
                                return Err(crate::error::ParseError::MissingAttribute {
                                    element: "fldChar".into(),
                                    attr: "fldCharType".into(),
                                });
                            }
                        };
                        pending_inlines.push(Inline::FieldChar(FieldChar {
                            field_char_type,
                            dirty: xml::optional_attr_bool(e, b"dirty")?,
                            fld_lock: xml::optional_attr_bool(e, b"fldLock")?,
                        }));
                    }
                    _ => xml::warn_unsupported_element("run", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"r" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"r")),
            _ => {}
        }
    }

    flush_text_owned(&mut texts, char_style_id, run_props, &run_rsids, content);
    content.extend(pending_inlines);

    Ok(())
}

/// Flush accumulated text into a `TextRun`, cloning properties (mid-run flush).
fn flush_text(
    texts: &mut Vec<String>,
    style_id: &Option<StyleId>,
    props: &RunProperties,
    rsids: &RevisionIds,
    content: &mut Vec<Inline>,
) {
    if texts.is_empty() {
        return;
    }
    let combined: String = texts.drain(..).collect();
    if !combined.is_empty() {
        content.push(Inline::TextRun(Box::new(TextRun {
            style_id: style_id.clone(),
            properties: props.clone(),
            text: combined,
            rsids: *rsids,
        })));
    }
}

/// Flush accumulated text into a `TextRun`, taking ownership of properties (end-of-run flush).
fn flush_text_owned(
    texts: &mut Vec<String>,
    style_id: Option<StyleId>,
    props: RunProperties,
    rsids: &RevisionIds,
    content: &mut Vec<Inline>,
) {
    if texts.is_empty() {
        return;
    }
    let combined: String = texts.drain(..).collect();
    if !combined.is_empty() {
        content.push(Inline::TextRun(Box::new(TextRun {
            style_id,
            properties: props,
            text: combined,
            rsids: *rsids,
        })));
    }
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
    match dxpdf_field::parse(raw) {
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

// ── VML / Pict ──────────────────────────────────────────────────────────────

/// §17.3.3.19: parse `w:pict` — VML picture container.
pub fn parse_pict(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Pict> {
    let mut shape_type = None;
    let mut shapes = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"shapetype" => {
                        shape_type = Some(parse_vml_shapetype(e, reader, buf)?);
                    }
                    b"shape" => {
                        shapes.push(parse_vml_shape(e, reader, buf)?);
                    }
                    _ => {
                        xml::warn_unsupported_element("pict", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"shapetype" => {
                        shape_type = Some(parse_vml_shapetype_from_attrs(e)?);
                    }
                    b"shape" => {
                        shapes.push(parse_vml_shape_from_attrs(e)?);
                    }
                    _ => xml::warn_unsupported_element("pict", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pict" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"pict")),
            _ => {}
        }
    }

    Ok(Pict { shape_type, shapes })
}

/// VML §14.1.2.20: parse self-closing `v:shapetype` (attributes only, no children).
fn parse_vml_shapetype_from_attrs(e: &BytesStart<'_>) -> Result<VmlShapeType> {
    Ok(VmlShapeType {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        coord_size: parse_vml_vector2d(xml::optional_attr(e, b"coordsize")?),
        spt: xml::optional_attr(e, b"spt")?.map(|s| s.parse::<f32>()).transpose().ok().flatten(),
        adj: parse_vml_adj(xml::optional_attr(e, b"adj")?),
        path: parse_vml_path_commands(xml::optional_attr(e, b"path")?),
        filled: vml_bool_attr(e, b"filled")?,
        stroked: vml_bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        formulas: Vec::new(),
        lock: None,
    })
}

/// VML §14.1.2.20: parse `v:shapetype` with attributes and children.
fn parse_vml_shapetype(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<VmlShapeType> {
    let mut st = VmlShapeType {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        coord_size: parse_vml_vector2d(xml::optional_attr(e, b"coordsize")?),
        spt: xml::optional_attr(e, b"spt")?.map(|s| s.parse::<f32>()).transpose().ok().flatten(),
        adj: parse_vml_adj(xml::optional_attr(e, b"adj")?),
        path: parse_vml_path_commands(xml::optional_attr(e, b"path")?),
        filled: vml_bool_attr(e, b"filled")?,
        stroked: vml_bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        formulas: Vec::new(),
        lock: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"formulas" => {
                        st.formulas = parse_vml_formulas(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("shapetype", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"stroke" => st.stroke = Some(parse_vml_stroke_attrs(e)?),
                    b"path" => st.vml_path = Some(parse_vml_path_attrs(e)?),
                    b"lock" => st.lock = Some(parse_vml_lock(e)?),
                    _ => xml::warn_unsupported_element("shapetype", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"shapetype" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"shapetype")),
            _ => {}
        }
    }

    Ok(st)
}

/// VML §14.1.2.19: parse self-closing `v:shape` (attributes only, no children).
fn parse_vml_shape_from_attrs(e: &BytesStart<'_>) -> Result<VmlShape> {
    Ok(VmlShape {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        shape_type_ref: xml::optional_attr(e, b"type")?
            .map(|s| VmlShapeId::new(s.strip_prefix('#').unwrap_or(&s))),
        style: parse_vml_style(xml::optional_attr(e, b"style")?),
        fill_color: xml::optional_attr(e, b"fillcolor")?.map(|s| parse_vml_color(&s)).transpose()?,
        stroked: vml_bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        text_box: None,
        wrap: None,
        image_data: None,
    })
}

/// VML §14.1.2.19: parse `v:shape` with attributes and children.
fn parse_vml_shape(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<VmlShape> {
    let mut shape = VmlShape {
        id: xml::optional_attr(e, b"id")?.map(VmlShapeId::new),
        shape_type_ref: xml::optional_attr(e, b"type")?
            .map(|s| VmlShapeId::new(s.strip_prefix('#').unwrap_or(&s))),
        style: parse_vml_style(xml::optional_attr(e, b"style")?),
        fill_color: xml::optional_attr(e, b"fillcolor")?.map(|s| parse_vml_color(&s)).transpose()?,
        stroked: vml_bool_attr(e, b"stroked")?,
        stroke: None,
        vml_path: None,
        text_box: None,
        wrap: None,
        image_data: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"textbox" => {
                        shape.text_box = Some(parse_vml_textbox(e, reader, buf)?);
                    }
                    b"wrap" => {
                        shape.wrap = Some(parse_vml_wrap(e)?);
                        xml::skip_to_end(reader, buf, b"wrap")?;
                    }
                    _ => {
                        xml::warn_unsupported_element("shape", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"stroke" => shape.stroke = Some(parse_vml_stroke_attrs(e)?),
                    b"path" => shape.vml_path = Some(parse_vml_path_attrs(e)?),
                    b"imagedata" => shape.image_data = Some(parse_vml_imagedata(e)?),
                    b"wrap" => shape.wrap = Some(parse_vml_wrap(e)?),
                    _ => xml::warn_unsupported_element("shape", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"shape" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"shape")),
            _ => {}
        }
    }

    Ok(shape)
}

/// VML §14.1.2.21: parse `v:stroke` attributes.
fn parse_vml_stroke_attrs(e: &BytesStart<'_>) -> Result<VmlStroke> {
    let dash_style = match xml::optional_attr(e, b"dashstyle")?.as_deref() {
        Some("solid") => Some(VmlDashStyle::Solid),
        Some("shortdash") => Some(VmlDashStyle::ShortDash),
        Some("shortdot") => Some(VmlDashStyle::ShortDot),
        Some("shortdashdot") => Some(VmlDashStyle::ShortDashDot),
        Some("shortdashdotdot") => Some(VmlDashStyle::ShortDashDotDot),
        Some("dot") => Some(VmlDashStyle::Dot),
        Some("dash") => Some(VmlDashStyle::Dash),
        Some("longdash") => Some(VmlDashStyle::LongDash),
        Some("dashdot") => Some(VmlDashStyle::DashDot),
        Some("longdashdot") => Some(VmlDashStyle::LongDashDot),
        Some("longdashdotdot") => Some(VmlDashStyle::LongDashDotDot),
        Some(other) => {
            log::warn!("vml-stroke: unsupported dashstyle {:?}", other);
            None
        }
        None => None,
    };
    let join_style = match xml::optional_attr(e, b"joinstyle")?.as_deref() {
        Some("round") => Some(VmlJoinStyle::Round),
        Some("bevel") => Some(VmlJoinStyle::Bevel),
        Some("miter") => Some(VmlJoinStyle::Miter),
        Some(other) => {
            log::warn!("vml-stroke: unsupported joinstyle {:?}", other);
            None
        }
        None => None,
    };
    Ok(VmlStroke {
        dash_style,
        join_style,
    })
}

/// VML §14.1.2.14: parse `v:path` attributes.
fn parse_vml_path_attrs(e: &BytesStart<'_>) -> Result<VmlPath> {
    Ok(VmlPath {
        gradient_shape_ok: vml_bool_attr(e, b"gradientshapeok")?,
        connect_type: match xml::optional_attr(e, b"connecttype")?.as_deref() {
            Some("none") => Some(VmlConnectType::None),
            Some("rect") => Some(VmlConnectType::Rect),
            Some("segments") => Some(VmlConnectType::Segments),
            Some("custom") => Some(VmlConnectType::Custom),
            Some(other) => {
                log::warn!("vml-path: unsupported connecttype {:?}", other);
                None
            }
            None => None,
        },
        extrusion_ok: vml_bool_attr(e, b"extrusionok")?,
    })
}

/// VML §14.1.2.6: parse `v:formulas` — list of `v:f` formula equations.
fn parse_vml_formulas(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<VmlFormula>> {
    let mut formulas = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"f" => {
                if let Some(eqn) = xml::optional_attr(e, b"eqn")? {
                    if let Some(f) = parse_vml_formula(&eqn) {
                        formulas.push(f);
                    } else {
                        log::warn!("vml-formula: failed to parse {:?}", eqn);
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"formulas" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"formulas")),
            _ => {}
        }
    }

    Ok(formulas)
}

/// Parse a single VML formula equation string (e.g., "sum #0 0 10800").
fn parse_vml_formula(eqn: &str) -> Option<VmlFormula> {
    let parts: Vec<&str> = eqn.split_whitespace().collect();
    if parts.is_empty() {
        return None;
    }

    let operation = match parts[0] {
        "val" => VmlFormulaOp::Val,
        "sum" => VmlFormulaOp::Sum,
        "prod" => VmlFormulaOp::Product,
        "mid" => VmlFormulaOp::Mid,
        "abs" => VmlFormulaOp::Abs,
        "min" => VmlFormulaOp::Min,
        "max" => VmlFormulaOp::Max,
        "if" => VmlFormulaOp::If,
        "sqrt" => VmlFormulaOp::Sqrt,
        "mod" => VmlFormulaOp::Mod,
        "sin" => VmlFormulaOp::Sin,
        "cos" => VmlFormulaOp::Cos,
        "tan" => VmlFormulaOp::Tan,
        "atan2" => VmlFormulaOp::Atan2,
        "sinatan2" => VmlFormulaOp::SinAtan2,
        "cosatan2" => VmlFormulaOp::CosAtan2,
        "sumangle" => VmlFormulaOp::SumAngle,
        "ellipse" => VmlFormulaOp::Ellipse,
        other => {
            log::warn!("vml-formula: unsupported operation {:?}", other);
            return None;
        }
    };

    let arg = |i: usize| -> VmlFormulaArg {
        parts
            .get(i)
            .and_then(|s| parse_vml_formula_arg(s))
            .unwrap_or(VmlFormulaArg::Literal(0))
    };

    Some(VmlFormula {
        operation,
        args: [arg(1), arg(2), arg(3)],
    })
}

/// Parse a single VML formula argument.
fn parse_vml_formula_arg(s: &str) -> Option<VmlFormulaArg> {
    if let Some(rest) = s.strip_prefix('#') {
        return rest.parse::<u32>().ok().map(VmlFormulaArg::AdjRef);
    }
    if let Some(rest) = s.strip_prefix('@') {
        return rest.parse::<u32>().ok().map(VmlFormulaArg::FormulaRef);
    }
    let guide = match s {
        "width" => Some(VmlGuide::Width),
        "height" => Some(VmlGuide::Height),
        "xcenter" => Some(VmlGuide::XCenter),
        "ycenter" => Some(VmlGuide::YCenter),
        "xrange" => Some(VmlGuide::XRange),
        "yrange" => Some(VmlGuide::YRange),
        "pixelWidth" => Some(VmlGuide::PixelWidth),
        "pixelHeight" => Some(VmlGuide::PixelHeight),
        "pixelLineWidth" => Some(VmlGuide::PixelLineWidth),
        "emuWidth" => Some(VmlGuide::EmuWidth),
        "emuHeight" => Some(VmlGuide::EmuHeight),
        "emuWidth2" => Some(VmlGuide::EmuWidth2),
        "emuHeight2" => Some(VmlGuide::EmuHeight2),
        _ => None,
    };
    if let Some(g) = guide {
        return Some(VmlFormulaArg::Guide(g));
    }
    s.parse::<i64>().ok().map(VmlFormulaArg::Literal)
}

/// VML §14.1.2.22: parse `v:textbox` with `w:txbxContent`.
fn parse_vml_textbox(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<VmlTextBox> {
    let style = parse_vml_style(xml::optional_attr(e, b"style")?);
    let inset = parse_vml_textbox_inset(xml::optional_attr(e, b"inset")?);
    let mut content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"txbxContent" => {
                        content = parse_txbx_content(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("textbox", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"textbox" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"textbox")),
            _ => {}
        }
    }

    Ok(VmlTextBox {
        style,
        inset,
        content,
    })
}

/// §17.17.1: parse `w:txbxContent` — block-level content inside a text box.
fn parse_txbx_content(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<Block>> {
    let (blocks, _) = parse_block_content(reader, buf, b"txbxContent")?;
    Ok(blocks)
}

/// Parse a VML `style` attribute — semicolon-separated CSS2 properties.
fn parse_vml_style(s: Option<String>) -> VmlStyle {
    let s = match s {
        Some(s) => s,
        None => return VmlStyle::default(),
    };

    let mut style = VmlStyle::default();

    for decl in s.split(';') {
        let decl = decl.trim();
        if decl.is_empty() {
            continue;
        }
        let Some((key, val)) = decl.split_once(':') else {
            continue;
        };
        let key = key.trim();
        let val = val.trim();

        match key {
            "position" => {
                style.position = match val {
                    "static" => Some(CssPosition::Static),
                    "relative" => Some(CssPosition::Relative),
                    "absolute" => Some(CssPosition::Absolute),
                    _ => {
                        log::warn!("vml-style: unsupported position value {:?}", val);
                        None
                    }
                };
            }
            "left" => style.left = parse_vml_length(val),
            "top" => style.top = parse_vml_length(val),
            "width" => style.width = parse_vml_length(val),
            "height" => style.height = parse_vml_length(val),
            "margin-left" => style.margin_left = parse_vml_length(val),
            "margin-top" => style.margin_top = parse_vml_length(val),
            "margin-right" => style.margin_right = parse_vml_length(val),
            "margin-bottom" => style.margin_bottom = parse_vml_length(val),
            "z-index" => style.z_index = val.parse::<i64>().ok(),
            "rotation" => style.rotation = val.parse::<f64>().ok(),
            "flip" => {
                style.flip = match val {
                    "x" => Some(VmlFlip::X),
                    "y" => Some(VmlFlip::Y),
                    "xy" | "yx" => Some(VmlFlip::XY),
                    _ => {
                        log::warn!("vml-style: unsupported flip value {:?}", val);
                        None
                    }
                };
            }
            "visibility" => {
                style.visibility = match val {
                    "visible" => Some(CssVisibility::Visible),
                    "hidden" => Some(CssVisibility::Hidden),
                    "inherit" => Some(CssVisibility::Inherit),
                    _ => {
                        log::warn!("vml-style: unsupported visibility value {:?}", val);
                        None
                    }
                };
            }
            "mso-position-horizontal" => {
                style.mso_position_horizontal = match val {
                    "absolute" => Some(MsoPositionH::Absolute),
                    "left" => Some(MsoPositionH::Left),
                    "center" => Some(MsoPositionH::Center),
                    "right" => Some(MsoPositionH::Right),
                    "inside" => Some(MsoPositionH::Inside),
                    "outside" => Some(MsoPositionH::Outside),
                    _ => {
                        log::warn!("vml-style: unsupported mso-position-horizontal value {:?}", val);
                        None
                    }
                };
            }
            "mso-position-horizontal-relative" => {
                style.mso_position_horizontal_relative = match val {
                    "margin" => Some(MsoPositionHRelative::Margin),
                    "page" => Some(MsoPositionHRelative::Page),
                    "text" => Some(MsoPositionHRelative::Text),
                    "char" => Some(MsoPositionHRelative::Char),
                    "left-margin-area" => Some(MsoPositionHRelative::LeftMarginArea),
                    "right-margin-area" => Some(MsoPositionHRelative::RightMarginArea),
                    "inner-margin-area" => Some(MsoPositionHRelative::InnerMarginArea),
                    "outer-margin-area" => Some(MsoPositionHRelative::OuterMarginArea),
                    _ => {
                        log::warn!("vml-style: unsupported mso-position-horizontal-relative value {:?}", val);
                        None
                    }
                };
            }
            "mso-position-vertical" => {
                style.mso_position_vertical = match val {
                    "absolute" => Some(MsoPositionV::Absolute),
                    "top" => Some(MsoPositionV::Top),
                    "center" => Some(MsoPositionV::Center),
                    "bottom" => Some(MsoPositionV::Bottom),
                    "inside" => Some(MsoPositionV::Inside),
                    "outside" => Some(MsoPositionV::Outside),
                    _ => {
                        log::warn!("vml-style: unsupported mso-position-vertical value {:?}", val);
                        None
                    }
                };
            }
            "mso-position-vertical-relative" => {
                style.mso_position_vertical_relative = match val {
                    "margin" => Some(MsoPositionVRelative::Margin),
                    "page" => Some(MsoPositionVRelative::Page),
                    "text" => Some(MsoPositionVRelative::Text),
                    "line" => Some(MsoPositionVRelative::Line),
                    "top-margin-area" => Some(MsoPositionVRelative::TopMarginArea),
                    "bottom-margin-area" => Some(MsoPositionVRelative::BottomMarginArea),
                    "inner-margin-area" => Some(MsoPositionVRelative::InnerMarginArea),
                    "outer-margin-area" => Some(MsoPositionVRelative::OuterMarginArea),
                    _ => {
                        log::warn!("vml-style: unsupported mso-position-vertical-relative value {:?}", val);
                        None
                    }
                };
            }
            "mso-wrap-distance-left" => {
                style.mso_wrap_distance_left = parse_vml_length(val);
            }
            "mso-wrap-distance-right" => {
                style.mso_wrap_distance_right = parse_vml_length(val);
            }
            "mso-wrap-distance-top" => {
                style.mso_wrap_distance_top = parse_vml_length(val);
            }
            "mso-wrap-distance-bottom" => {
                style.mso_wrap_distance_bottom = parse_vml_length(val);
            }
            "mso-wrap-style" => {
                style.mso_wrap_style = match val {
                    "square" => Some(MsoWrapStyle::Square),
                    "none" => Some(MsoWrapStyle::None),
                    "tight" => Some(MsoWrapStyle::Tight),
                    "through" => Some(MsoWrapStyle::Through),
                    _ => {
                        log::warn!("vml-style: unsupported mso-wrap-style value {:?}", val);
                        None
                    }
                };
            }
            _ => log::warn!("vml-style: unsupported property {:?}: {:?}", key, val),
        }
    }

    style
}

/// Parse a CSS length value (e.g., "468pt", "0", "3.5in", "50%").
fn parse_vml_length(s: &str) -> Option<VmlLength> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }

    // Try known unit suffixes.
    let (num_str, unit) = if let Some(n) = s.strip_suffix("pt") {
        (n, VmlLengthUnit::Pt)
    } else if let Some(n) = s.strip_suffix("in") {
        (n, VmlLengthUnit::In)
    } else if let Some(n) = s.strip_suffix("cm") {
        (n, VmlLengthUnit::Cm)
    } else if let Some(n) = s.strip_suffix("mm") {
        (n, VmlLengthUnit::Mm)
    } else if let Some(n) = s.strip_suffix("px") {
        (n, VmlLengthUnit::Px)
    } else if let Some(n) = s.strip_suffix("em") {
        (n, VmlLengthUnit::Em)
    } else if let Some(n) = s.strip_suffix('%') {
        (n, VmlLengthUnit::Percent)
    } else {
        // Find where the numeric part ends and the suffix begins.
        let split = s
            .find(|c: char| !c.is_ascii_digit() && c != '.' && c != '-' && c != '+')
            .unwrap_or(s.len());
        if split < s.len() {
            log::warn!("vml-length: unsupported unit suffix {:?}", &s[split..]);
            return None;
        }
        (s, VmlLengthUnit::None)
    };

    let value = num_str.trim().parse::<f64>().ok()?;
    Some(VmlLength { value, unit })
}

/// Parse a VML color value (§14.1.2.1): `#RRGGBB`, `RRGGBB` hex, or named color.
fn parse_vml_color(s: &str) -> Result<VmlColor> {
    let hex = s.strip_prefix('#').unwrap_or(s);
    if hex.len() == 6 && hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
        let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
        let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);
        return Ok(VmlColor::Rgb(r, g, b));
    }
    match parse_vml_named_color(s) {
        Some(named) => Ok(VmlColor::Named(named)),
        None => Err(crate::error::ParseError::InvalidAttributeValue {
            attr: "fillcolor".into(),
            value: s.into(),
            reason: "unrecognized VML color name per §14.1.2.1".into(),
        }),
    }
}

fn parse_vml_named_color(s: &str) -> Option<VmlNamedColor> {
    // Case-insensitive match per CSS spec.
    Some(match s.to_ascii_lowercase().as_str() {
        // CSS2.1 standard colors.
        "black" => VmlNamedColor::Black,
        "silver" => VmlNamedColor::Silver,
        "gray" | "grey" => VmlNamedColor::Gray,
        "white" => VmlNamedColor::White,
        "maroon" => VmlNamedColor::Maroon,
        "red" => VmlNamedColor::Red,
        "purple" => VmlNamedColor::Purple,
        "fuchsia" => VmlNamedColor::Fuchsia,
        "green" => VmlNamedColor::Green,
        "lime" => VmlNamedColor::Lime,
        "olive" => VmlNamedColor::Olive,
        "yellow" => VmlNamedColor::Yellow,
        "navy" => VmlNamedColor::Navy,
        "blue" => VmlNamedColor::Blue,
        "teal" => VmlNamedColor::Teal,
        "aqua" => VmlNamedColor::Aqua,
        "orange" => VmlNamedColor::Orange,
        // SVG/CSS3 extended colors.
        "aliceblue" => VmlNamedColor::AliceBlue,
        "antiquewhite" => VmlNamedColor::AntiqueWhite,
        "beige" => VmlNamedColor::Beige,
        "bisque" => VmlNamedColor::Bisque,
        "blanchedalmond" => VmlNamedColor::BlanchedAlmond,
        "blueviolet" => VmlNamedColor::BlueViolet,
        "brown" => VmlNamedColor::Brown,
        "burlywood" => VmlNamedColor::BurlyWood,
        "cadetblue" => VmlNamedColor::CadetBlue,
        "chartreuse" => VmlNamedColor::Chartreuse,
        "chocolate" => VmlNamedColor::Chocolate,
        "coral" => VmlNamedColor::Coral,
        "cornflowerblue" => VmlNamedColor::CornflowerBlue,
        "cornsilk" => VmlNamedColor::Cornsilk,
        "crimson" => VmlNamedColor::Crimson,
        "cyan" => VmlNamedColor::Cyan,
        "darkblue" => VmlNamedColor::DarkBlue,
        "darkcyan" => VmlNamedColor::DarkCyan,
        "darkgoldenrod" => VmlNamedColor::DarkGoldenrod,
        "darkgray" | "darkgrey" => VmlNamedColor::DarkGray,
        "darkgreen" => VmlNamedColor::DarkGreen,
        "darkkhaki" => VmlNamedColor::DarkKhaki,
        "darkmagenta" => VmlNamedColor::DarkMagenta,
        "darkolivegreen" => VmlNamedColor::DarkOliveGreen,
        "darkorange" => VmlNamedColor::DarkOrange,
        "darkorchid" => VmlNamedColor::DarkOrchid,
        "darkred" => VmlNamedColor::DarkRed,
        "darksalmon" => VmlNamedColor::DarkSalmon,
        "darkseagreen" => VmlNamedColor::DarkSeaGreen,
        "darkslateblue" => VmlNamedColor::DarkSlateBlue,
        "darkslategray" | "darkslategrey" => VmlNamedColor::DarkSlateGray,
        "darkturquoise" => VmlNamedColor::DarkTurquoise,
        "darkviolet" => VmlNamedColor::DarkViolet,
        "deeppink" => VmlNamedColor::DeepPink,
        "deepskyblue" => VmlNamedColor::DeepSkyBlue,
        "dimgray" | "dimgrey" => VmlNamedColor::DimGray,
        "dodgerblue" => VmlNamedColor::DodgerBlue,
        "firebrick" => VmlNamedColor::Firebrick,
        "floralwhite" => VmlNamedColor::FloralWhite,
        "forestgreen" => VmlNamedColor::ForestGreen,
        "gainsboro" => VmlNamedColor::Gainsboro,
        "ghostwhite" => VmlNamedColor::GhostWhite,
        "gold" => VmlNamedColor::Gold,
        "goldenrod" => VmlNamedColor::Goldenrod,
        "greenyellow" => VmlNamedColor::GreenYellow,
        "honeydew" => VmlNamedColor::Honeydew,
        "hotpink" => VmlNamedColor::HotPink,
        "indianred" => VmlNamedColor::IndianRed,
        "indigo" => VmlNamedColor::Indigo,
        "ivory" => VmlNamedColor::Ivory,
        "khaki" => VmlNamedColor::Khaki,
        "lavender" => VmlNamedColor::Lavender,
        "lavenderblush" => VmlNamedColor::LavenderBlush,
        "lawngreen" => VmlNamedColor::LawnGreen,
        "lemonchiffon" => VmlNamedColor::LemonChiffon,
        "lightblue" => VmlNamedColor::LightBlue,
        "lightcoral" => VmlNamedColor::LightCoral,
        "lightcyan" => VmlNamedColor::LightCyan,
        "lightgoldenrodyellow" => VmlNamedColor::LightGoldenrodYellow,
        "lightgray" | "lightgrey" => VmlNamedColor::LightGray,
        "lightgreen" => VmlNamedColor::LightGreen,
        "lightpink" => VmlNamedColor::LightPink,
        "lightsalmon" => VmlNamedColor::LightSalmon,
        "lightseagreen" => VmlNamedColor::LightSeaGreen,
        "lightskyblue" => VmlNamedColor::LightSkyBlue,
        "lightslategray" | "lightslategrey" => VmlNamedColor::LightSlateGray,
        "lightsteelblue" => VmlNamedColor::LightSteelBlue,
        "lightyellow" => VmlNamedColor::LightYellow,
        "limegreen" => VmlNamedColor::LimeGreen,
        "linen" => VmlNamedColor::Linen,
        "magenta" => VmlNamedColor::Magenta,
        "mediumaquamarine" => VmlNamedColor::MediumAquamarine,
        "mediumblue" => VmlNamedColor::MediumBlue,
        "mediumorchid" => VmlNamedColor::MediumOrchid,
        "mediumpurple" => VmlNamedColor::MediumPurple,
        "mediumseagreen" => VmlNamedColor::MediumSeaGreen,
        "mediumslateblue" => VmlNamedColor::MediumSlateBlue,
        "mediumspringgreen" => VmlNamedColor::MediumSpringGreen,
        "mediumturquoise" => VmlNamedColor::MediumTurquoise,
        "mediumvioletred" => VmlNamedColor::MediumVioletRed,
        "midnightblue" => VmlNamedColor::MidnightBlue,
        "mintcream" => VmlNamedColor::MintCream,
        "mistyrose" => VmlNamedColor::MistyRose,
        "moccasin" => VmlNamedColor::Moccasin,
        "navajowhite" => VmlNamedColor::NavajoWhite,
        "oldlace" => VmlNamedColor::OldLace,
        "olivedrab" => VmlNamedColor::OliveDrab,
        "orangered" => VmlNamedColor::OrangeRed,
        "orchid" => VmlNamedColor::Orchid,
        "palegoldenrod" => VmlNamedColor::PaleGoldenrod,
        "palegreen" => VmlNamedColor::PaleGreen,
        "paleturquoise" => VmlNamedColor::PaleTurquoise,
        "palevioletred" => VmlNamedColor::PaleVioletRed,
        "papayawhip" => VmlNamedColor::PapayaWhip,
        "peachpuff" => VmlNamedColor::PeachPuff,
        "peru" => VmlNamedColor::Peru,
        "pink" => VmlNamedColor::Pink,
        "plum" => VmlNamedColor::Plum,
        "powderblue" => VmlNamedColor::PowderBlue,
        "rosybrown" => VmlNamedColor::RosyBrown,
        "royalblue" => VmlNamedColor::RoyalBlue,
        "saddlebrown" => VmlNamedColor::SaddleBrown,
        "salmon" => VmlNamedColor::Salmon,
        "sandybrown" => VmlNamedColor::SandyBrown,
        "seagreen" => VmlNamedColor::SeaGreen,
        "seashell" => VmlNamedColor::Seashell,
        "sienna" => VmlNamedColor::Sienna,
        "skyblue" => VmlNamedColor::SkyBlue,
        "slateblue" => VmlNamedColor::SlateBlue,
        "slategray" | "slategrey" => VmlNamedColor::SlateGray,
        "snow" => VmlNamedColor::Snow,
        "springgreen" => VmlNamedColor::SpringGreen,
        "steelblue" => VmlNamedColor::SteelBlue,
        "tan" => VmlNamedColor::Tan,
        "thistle" => VmlNamedColor::Thistle,
        "tomato" => VmlNamedColor::Tomato,
        "turquoise" => VmlNamedColor::Turquoise,
        "violet" => VmlNamedColor::Violet,
        "wheat" => VmlNamedColor::Wheat,
        "whitesmoke" => VmlNamedColor::WhiteSmoke,
        "yellowgreen" => VmlNamedColor::YellowGreen,
        // VML system colors.
        "buttonface" => VmlNamedColor::ButtonFace,
        "buttonhighlight" => VmlNamedColor::ButtonHighlight,
        "buttonshadow" => VmlNamedColor::ButtonShadow,
        "buttontext" => VmlNamedColor::ButtonText,
        "captiontext" => VmlNamedColor::CaptionText,
        "graytext" => VmlNamedColor::GrayText,
        "highlight" => VmlNamedColor::Highlight,
        "highlighttext" => VmlNamedColor::HighlightText,
        "inactiveborder" => VmlNamedColor::InactiveBorder,
        "inactivecaption" => VmlNamedColor::InactiveCaption,
        "inactivecaptiontext" => VmlNamedColor::InactiveCaptionText,
        "infobackground" => VmlNamedColor::InfoBackground,
        "infotext" => VmlNamedColor::InfoText,
        "menu" => VmlNamedColor::Menu,
        "menutext" => VmlNamedColor::MenuText,
        "scrollbar" => VmlNamedColor::Scrollbar,
        "threeddarkshadow" => VmlNamedColor::ThreeDDarkShadow,
        "threedface" => VmlNamedColor::ThreeDFace,
        "threedhighlight" => VmlNamedColor::ThreeDHighlight,
        "threedlightshadow" => VmlNamedColor::ThreeDLightShadow,
        "threedshadow" => VmlNamedColor::ThreeDShadow,
        "window" => VmlNamedColor::Window,
        "windowframe" => VmlNamedColor::WindowFrame,
        "windowtext" => VmlNamedColor::WindowText,
        _ => return None,
    })
}

/// Parse VML textbox `inset` attribute — comma-separated CSS lengths (left,top,right,bottom).
fn parse_vml_textbox_inset(s: Option<String>) -> Option<VmlTextBoxInset> {
    let s = s?;
    let parts: Vec<&str> = s.split(',').collect();
    Some(VmlTextBoxInset {
        left: parts.first().and_then(|v| parse_vml_length(v)),
        top: parts.get(1).and_then(|v| parse_vml_length(v)),
        right: parts.get(2).and_then(|v| parse_vml_length(v)),
        bottom: parts.get(3).and_then(|v| parse_vml_length(v)),
    })
}

/// Office VML extension: parse `o:lock` attributes.
fn parse_vml_lock(e: &BytesStart<'_>) -> Result<VmlLock> {
    let aspect_ratio = vml_bool_attr(e, b"aspectratio")?;
    let ext = match xml::optional_attr(e, b"ext")?.as_deref() {
        Some("edit") => Some(VmlExtHandling::Edit),
        Some("view") => Some(VmlExtHandling::View),
        Some("backwardCompatible") => Some(VmlExtHandling::BackwardCompatible),
        Some(other) => {
            return Err(crate::error::ParseError::InvalidAttributeValue {
                attr: "v:ext".into(),
                value: other.into(),
                reason: "expected edit, view, or backwardCompatible".into(),
            });
        }
        None => None,
    };
    Ok(VmlLock { aspect_ratio, ext })
}

/// VML §14.1.2.11: parse `v:imagedata` attributes.
fn parse_vml_imagedata(e: &BytesStart<'_>) -> Result<VmlImageData> {
    Ok(VmlImageData {
        rel_id: xml::optional_attr(e, b"id")?.map(RelId::new),
        title: xml::optional_attr(e, b"title")?,
    })
}

/// VML §14.1.2.23: parse `v:wrap` attributes.
fn parse_vml_wrap(e: &BytesStart<'_>) -> Result<VmlWrap> {
    let wrap_type = match xml::optional_attr(e, b"type")?.as_deref() {
        Some("topAndBottom") => Some(VmlWrapType::TopAndBottom),
        Some("square") => Some(VmlWrapType::Square),
        Some("none") => Some(VmlWrapType::None),
        Some("tight") => Some(VmlWrapType::Tight),
        Some("through") => Some(VmlWrapType::Through),
        Some(other) => {
            return Err(crate::error::ParseError::InvalidAttributeValue {
                attr: "type".into(),
                value: other.into(),
                reason: "expected value per VML §14.1.2.23".into(),
            });
        }
        None => None,
    };
    let side = match xml::optional_attr(e, b"side")?.as_deref() {
        Some("both") => Some(VmlWrapSide::Both),
        Some("left") => Some(VmlWrapSide::Left),
        Some("right") => Some(VmlWrapSide::Right),
        Some("largest") => Some(VmlWrapSide::Largest),
        Some(other) => {
            return Err(crate::error::ParseError::InvalidAttributeValue {
                attr: "side".into(),
                value: other.into(),
                reason: "expected value per VML §14.1.2.23".into(),
            });
        }
        None => None,
    };
    Ok(VmlWrap { wrap_type, side })
}

/// Parse a VML boolean attribute ("t"/"f" or "true"/"false").
fn vml_bool_attr(e: &BytesStart<'_>, name: &[u8]) -> Result<Option<bool>> {
    match xml::optional_attr(e, name)?.as_deref() {
        Some("t") | Some("true") => Ok(Some(true)),
        Some("f") | Some("false") => Ok(Some(false)),
        Some(other) => {
            xml::warn_unsupported_attr(
                "vml",
                &String::from_utf8_lossy(name),
                other,
            );
            Ok(None)
        }
        None => Ok(None),
    }
}

/// Parse VML path commands from the `path` attribute string (§14.2.1.6).
fn parse_vml_path_commands(s: Option<String>) -> Vec<VmlPathCommand> {
    let s = match s {
        Some(s) => s,
        None => return Vec::new(),
    };

    let mut cmds = Vec::new();
    // Tokenize: split on commas and whitespace, but keep command letters as separate tokens.
    let mut tokens: Vec<&str> = Vec::new();
    let mut rest = s.as_str();
    while !rest.is_empty() {
        // Skip whitespace and commas.
        rest = rest.trim_start_matches(|c: char| c == ',' || c.is_ascii_whitespace());
        if rest.is_empty() {
            break;
        }
        // Check for multi-char commands first.
        let cmd_len = if rest.starts_with("wa")
            || rest.starts_with("wr")
            || rest.starts_with("at")
            || rest.starts_with("ar")
            || rest.starts_with("qx")
            || rest.starts_with("qy")
            || rest.starts_with("nf")
            || rest.starts_with("ns")
            || rest.starts_with("hа") // ha..hh variants
        {
            2
        } else if rest.starts_with(|c: char| c.is_ascii_alphabetic()) {
            1
        } else {
            0
        };

        if cmd_len > 0 {
            tokens.push(&rest[..cmd_len]);
            rest = &rest[cmd_len..];
        } else if rest.starts_with('@') {
            // §14.2.1.6: @n formula reference — consume @digits.
            let end = rest[1..]
                .find(|c: char| !c.is_ascii_digit())
                .map(|p| p + 1)
                .unwrap_or(rest.len());
            tokens.push(&rest[..end]);
            rest = &rest[end..];
        } else {
            // Numeric token: consume until delimiter.
            let end = rest
                .find(|c: char| {
                    c == ',' || c == '@' || c.is_ascii_whitespace() || c.is_ascii_alphabetic()
                })
                .unwrap_or(rest.len());
            if end > 0 {
                tokens.push(&rest[..end]);
                rest = &rest[end..];
            } else {
                rest = &rest[1..]; // skip unrecognized char
            }
        }
    }

    let mut i = 0;
    while i < tokens.len() {
        let tok = tokens[i];
        i += 1;
        match tok {
            "m" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::MoveTo { x, y });
                }
            }
            "l" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::LineTo { x, y });
                }
            }
            "c" => {
                if let Some((x1, y1, x2, y2, x, y)) = take_6_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::CurveTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    });
                }
            }
            "r" => {
                if let Some((dx, dy)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::RLineTo { dx, dy });
                }
            }
            "v" => {
                if let Some((dx1, dy1, dx2, dy2, dx, dy)) = take_6_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::RCurveTo {
                        dx1,
                        dy1,
                        dx2,
                        dy2,
                        dx,
                        dy,
                    });
                }
            }
            "t" => {
                if let Some((dx, dy)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::RMoveTo { dx, dy });
                }
            }
            "x" => cmds.push(VmlPathCommand::Close),
            "e" => cmds.push(VmlPathCommand::End),
            "qx" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::QuadrantX { x, y });
                }
            }
            "qy" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::QuadrantY { x, y });
                }
            }
            "nf" => cmds.push(VmlPathCommand::NoFill),
            "ns" => cmds.push(VmlPathCommand::NoStroke),
            "wa" | "wr" | "at" | "ar" => {
                let kind = match tok {
                    "wa" => VmlArcKind::WA,
                    "wr" => VmlArcKind::WR,
                    "at" => VmlArcKind::AT,
                    _ => VmlArcKind::AR,
                };
                let args = (|| {
                    Some(VmlPathCommand::Arc {
                        kind,
                        bounding_x1: take_coord(&tokens, &mut i)?,
                        bounding_y1: take_coord(&tokens, &mut i)?,
                        bounding_x2: take_coord(&tokens, &mut i)?,
                        bounding_y2: take_coord(&tokens, &mut i)?,
                        start_x: take_coord(&tokens, &mut i)?,
                        start_y: take_coord(&tokens, &mut i)?,
                        end_x: take_coord(&tokens, &mut i)?,
                        end_y: take_coord(&tokens, &mut i)?,
                    })
                })();
                if let Some(cmd) = args {
                    cmds.push(cmd);
                }
            }
            _ => {
                // §14.2.1.6: bare coordinate in command position — implicit lineto.
                let x = if let Some(rest) = tok.strip_prefix('@') {
                    rest.parse::<u32>().ok().map(VmlPathCoord::FormulaRef)
                } else {
                    tok.parse::<i64>().ok().map(VmlPathCoord::Literal)
                };
                if let Some(x) = x {
                    if let Some(y) = take_coord(&tokens, &mut i) {
                        cmds.push(VmlPathCommand::LineTo { x, y });
                    }
                } else {
                    log::warn!("vml-path: unsupported command {:?}", tok);
                }
            }
        }
    }

    cmds
}

fn take_coord(tokens: &[&str], i: &mut usize) -> Option<VmlPathCoord> {
    if *i >= tokens.len() {
        return None;
    }
    let tok = tokens[*i];
    let coord = if let Some(rest) = tok.strip_prefix('@') {
        VmlPathCoord::FormulaRef(rest.parse::<u32>().ok()?)
    } else {
        VmlPathCoord::Literal(tok.parse::<i64>().ok()?)
    };
    *i += 1;
    Some(coord)
}

fn take_2_coord(
    tokens: &[&str],
    i: &mut usize,
) -> Option<(VmlPathCoord, VmlPathCoord)> {
    let a = take_coord(tokens, i)?;
    let b = take_coord(tokens, i)?;
    Some((a, b))
}

fn take_6_coord(
    tokens: &[&str],
    i: &mut usize,
) -> Option<(
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
)> {
    let a = take_coord(tokens, i)?;
    let b = take_coord(tokens, i)?;
    let c = take_coord(tokens, i)?;
    let d = take_coord(tokens, i)?;
    let e = take_coord(tokens, i)?;
    let f = take_coord(tokens, i)?;
    Some((a, b, c, d, e, f))
}



/// Parse a VML `adj` attribute — comma-separated integer adjustment values.
fn parse_vml_adj(s: Option<String>) -> Vec<i64> {
    match s {
        Some(s) => s
            .split(',')
            .filter_map(|v| v.trim().parse::<i64>().ok())
            .collect(),
        None => Vec::new(),
    }
}

/// Parse a VML Vector2D string ("x,y") into `VmlVector2D`.
fn parse_vml_vector2d(s: Option<String>) -> Option<VmlVector2D> {
    let s = s?;
    let (x_str, y_str) = s.split_once(',')?;
    let x = x_str.trim().parse::<i64>().ok()?;
    let y = y_str.trim().parse::<i64>().ok()?;
    Some(VmlVector2D { x, y })
}

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
                        let requires_str =
                            xml::optional_attr(e, b"Requires")?.unwrap_or_default();
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
                                log::warn!(
                                    "mc:Choice: unsupported Requires {:?}",
                                    other
                                );
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
            Event::End(ref e)
                if xml::local_name(e.name().as_ref()) == b"AlternateContent" =>
            {
                break
            }
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
                        content.push(Inline::Pict(parse_pict(reader, buf)?));
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

fn parse_table(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Table> {
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
                        let (para, sect) = parse_paragraph(reader, buf, p_rsids)?;
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
