//! Parser for document body content: blocks (paragraphs, tables, section breaks)
//! and inline content (text runs, images, hyperlinks, fields, etc.).
//!
//! No style resolution or property merging — output is raw parsed data.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::dimension::{Dimension, Emu};
use crate::error::Result;
use crate::geometry::{EdgeInsets, Size};
use crate::model::*;
use crate::xml;

use super::properties;

/// Parse a body-level XML part (header, footer, footnotes, etc.) into blocks.
pub fn parse_blocks(data: &[u8]) -> Result<Vec<Block>> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut blocks = Vec::new();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"p" => {
                        let rsids = parse_paragraph_rsids(e)?;
                        let (para, sect) = parse_paragraph(&mut reader, &mut buf, rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" => {
                        blocks.push(Block::Table(Box::new(parse_table(&mut reader, &mut buf)?)));
                    }
                    b"sectPr" => {
                        let sect_rsids = properties::parse_section_rsids(e)?;
                        let mut sect = properties::parse_section_properties(&mut reader, &mut buf)?;
                        sect.rsids = sect_rsids;
                        blocks.push(Block::SectionBreak(Box::new(sect)));
                    }
                    _ => xml::warn_unsupported_element("body", &local),
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(blocks)
}

/// Parse `w:body`, returning blocks and final section properties.
pub fn parse_body(data: &[u8]) -> Result<(Vec<Block>, SectionProperties)> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut blocks = Vec::new();
    let mut final_section = SectionProperties::default();
    let mut in_body = false;

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"body" => in_body = true,
                    b"p" if in_body => {
                        let rsids = parse_paragraph_rsids(e)?;
                        let (para, sect) = parse_paragraph(&mut reader, &mut buf, rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" if in_body => {
                        blocks.push(Block::Table(Box::new(parse_table(&mut reader, &mut buf)?)));
                    }
                    b"sectPr" if in_body => {
                        let sect_rsids = properties::parse_section_rsids(e)?;
                        final_section =
                            properties::parse_section_properties(&mut reader, &mut buf)?;
                        final_section.rsids = sect_rsids;
                    }
                    _ => xml::warn_unsupported_element("body", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"body" => break,
            Event::Eof => break,
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
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
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
                                id: BookmarkId(id),
                                name,
                            });
                        }
                    }
                    b"bookmarkEnd" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            content.push(Inline::BookmarkEnd(BookmarkId(id)));
                        }
                    }
                    b"fldSimple" => {
                        let instr = xml::optional_attr(e, b"instr")?.unwrap_or_default();
                        let field = parse_simple_field_content(instr, reader, buf)?;
                        content.push(Inline::Field(field));
                    }
                    _ => xml::warn_unsupported_element("paragraph", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"bookmarkStart" => {
                        if let (Some(id), Some(name)) = (
                            xml::optional_attr_i64(e, b"id")?,
                            xml::optional_attr(e, b"name")?,
                        ) {
                            content.push(Inline::BookmarkStart {
                                id: BookmarkId(id),
                                name,
                            });
                        }
                    }
                    b"bookmarkEnd" => {
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            content.push(Inline::BookmarkEnd(BookmarkId(id)));
                        }
                    }
                    _ => xml::warn_unsupported_element("paragraph", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"p" => break,
            Event::Eof => break,
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
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
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
                    b"drawing" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        if let Some(img) = parse_drawing(reader, buf)? {
                            pending_inlines.push(Inline::Image(img));
                        }
                    }
                    _ => xml::warn_unsupported_element("run", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
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
                            pending_inlines.push(Inline::FootnoteRef(NoteId(id)));
                        }
                    }
                    b"endnoteReference" => {
                        flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            pending_inlines.push(Inline::EndnoteRef(NoteId(id)));
                        }
                    }
                    _ => xml::warn_unsupported_element("run", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"r" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    flush_text(&mut texts, &char_style_id, &run_props, &run_rsids, content);
    content.extend(pending_inlines);

    Ok(())
}

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
        content.push(Inline::TextRun(TextRun {
            style_id: style_id.clone(),
            properties: props.clone(),
            text: combined,
            rsids: *rsids,
        }));
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
        HyperlinkTarget::External(RelId(id))
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
            Event::Eof => break,
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
    instruction: String,
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
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Field {
        instruction,
        content: field_content,
    })
}

// ── Drawing / Image ──────────────────────────────────────────────────────────

fn parse_drawing(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<Image>> {
    let mut image: Option<Image> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"inline" => {
                        image = parse_inline_image(reader, buf)?;
                    }
                    b"anchor" => {
                        let behind_text =
                            xml::optional_attr_bool(e, b"behindDoc")?.unwrap_or(false);
                        let lock_anchor = xml::optional_attr_bool(e, b"locked")?.unwrap_or(false);
                        let allow_overlap =
                            xml::optional_attr_bool(e, b"allowOverlap")?.unwrap_or(true);
                        let relative_height =
                            xml::optional_attr_u32(e, b"relativeHeight")?.unwrap_or(0);
                        image = parse_anchor_image(
                            behind_text,
                            lock_anchor,
                            allow_overlap,
                            relative_height,
                            reader,
                            buf,
                        )?;
                    }
                    _ => xml::warn_unsupported_element("drawing", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"drawing" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(image)
}

fn parse_inline_image(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<Image>> {
    let mut extent = Size::ZERO;
    let mut rel_id: Option<String> = None;
    let mut description: Option<String> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"extent" => {
                        let cx = xml::optional_attr_i64(e, b"cx")?.unwrap_or(0);
                        let cy = xml::optional_attr_i64(e, b"cy")?.unwrap_or(0);
                        extent = Size::new(Dimension::new(cx), Dimension::new(cy));
                    }
                    b"docPr" => {
                        description = xml::optional_attr(e, b"descr")?;
                    }
                    b"blip" => {
                        rel_id = xml::optional_attr(e, b"embed")?;
                    }
                    _ => xml::warn_unsupported_element("inline-image", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"inline" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(rel_id.map(|id| Image {
        rel_id: RelId(id),
        extent,
        placement: ImagePlacement::Inline,
        description,
    }))
}

fn parse_anchor_image(
    behind_text: bool,
    lock_anchor: bool,
    allow_overlap: bool,
    relative_height: u32,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Option<Image>> {
    let mut extent = Size::ZERO;
    let mut rel_id: Option<String> = None;
    let mut description: Option<String> = None;
    let mut h_pos = AnchorPosition::Offset {
        relative_from: None,
        offset: Dimension::ZERO,
    };
    let mut v_pos = AnchorPosition::Offset {
        relative_from: None,
        offset: Dimension::ZERO,
    };
    let mut wrap = TextWrap::None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"positionH" => {
                        let rel_from = xml::optional_attr(e, b"relativeFrom")?
                            .map(|s| parse_anchor_relative_from(&s))
                            .transpose()?;
                        h_pos = parse_anchor_position(reader, buf, rel_from, b"positionH")?;
                    }
                    b"positionV" => {
                        let rel_from = xml::optional_attr(e, b"relativeFrom")?
                            .map(|s| parse_anchor_relative_from(&s))
                            .transpose()?;
                        v_pos = parse_anchor_position(reader, buf, rel_from, b"positionV")?;
                    }
                    b"wrapSquare" => {
                        wrap = TextWrap::Square {
                            distance: parse_wrap_distance(e)?,
                        };
                    }
                    b"wrapTight" => {
                        wrap = TextWrap::Tight {
                            distance: parse_wrap_distance(e)?,
                        };
                    }
                    b"wrapThrough" => {
                        wrap = TextWrap::Through {
                            distance: parse_wrap_distance(e)?,
                        };
                    }
                    b"wrapTopAndBottom" => {
                        let dt = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
                        let db = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
                        wrap = TextWrap::TopAndBottom {
                            distance_top: Dimension::new(dt),
                            distance_bottom: Dimension::new(db),
                        };
                    }
                    b"blip" => {
                        rel_id = xml::optional_attr(e, b"embed")?;
                    }
                    _ => xml::warn_unsupported_element("anchor-image", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"extent" => {
                        let cx = xml::optional_attr_i64(e, b"cx")?.unwrap_or(0);
                        let cy = xml::optional_attr_i64(e, b"cy")?.unwrap_or(0);
                        extent = Size::new(Dimension::new(cx), Dimension::new(cy));
                    }
                    b"docPr" => {
                        description = xml::optional_attr(e, b"descr")?;
                    }
                    b"blip" => {
                        rel_id = xml::optional_attr(e, b"embed")?;
                    }
                    b"wrapNone" => {
                        wrap = TextWrap::None;
                    }
                    _ => xml::warn_unsupported_element("anchor-image", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"anchor" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(rel_id.map(|id| Image {
        rel_id: RelId(id),
        extent,
        placement: ImagePlacement::Anchor(AnchorProperties {
            horizontal_position: h_pos,
            vertical_position: v_pos,
            wrap,
            behind_text,
            lock_anchor,
            allow_overlap,
            relative_height,
        }),
        description,
    }))
}

fn parse_anchor_position(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    relative_from: Option<AnchorRelativeFrom>,
    end_tag: &[u8],
) -> Result<AnchorPosition> {
    let mut result = AnchorPosition::Offset {
        relative_from,
        offset: Dimension::ZERO,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"posOffset" => {
                        let text = xml::read_text_content(reader, buf)?;
                        let val: i64 = text.trim().parse().unwrap_or(0);
                        result = AnchorPosition::Offset {
                            relative_from,
                            offset: Dimension::new(val),
                        };
                    }
                    b"align" => {
                        let text = xml::read_text_content(reader, buf)?;
                        result = AnchorPosition::Align {
                            relative_from,
                            alignment: parse_anchor_alignment(text.trim())?,
                        };
                    }
                    _ => xml::warn_unsupported_element("anchor-position", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(result)
}

fn parse_anchor_relative_from(val: &str) -> Result<AnchorRelativeFrom> {
    use crate::error::ParseError;
    match val {
        "page" => Ok(AnchorRelativeFrom::Page),
        "margin" => Ok(AnchorRelativeFrom::Margin),
        "column" => Ok(AnchorRelativeFrom::Column),
        "character" => Ok(AnchorRelativeFrom::Character),
        "paragraph" => Ok(AnchorRelativeFrom::Paragraph),
        "line" => Ok(AnchorRelativeFrom::Line),
        "insideMargin" => Ok(AnchorRelativeFrom::InsideMargin),
        "outsideMargin" => Ok(AnchorRelativeFrom::OutsideMargin),
        "topMargin" => Ok(AnchorRelativeFrom::TopMargin),
        "bottomMargin" => Ok(AnchorRelativeFrom::BottomMargin),
        "leftMargin" => Ok(AnchorRelativeFrom::LeftMargin),
        "rightMargin" => Ok(AnchorRelativeFrom::RightMargin),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "relativeFrom".into(),
            value: other.into(),
            reason: "unsupported value per OOXML spec".into(),
        }),
    }
}

fn parse_anchor_alignment(val: &str) -> Result<AnchorAlignment> {
    use crate::error::ParseError;
    match val {
        "left" => Ok(AnchorAlignment::Left),
        "center" => Ok(AnchorAlignment::Center),
        "right" => Ok(AnchorAlignment::Right),
        "inside" => Ok(AnchorAlignment::Inside),
        "outside" => Ok(AnchorAlignment::Outside),
        "top" => Ok(AnchorAlignment::Top),
        "bottom" => Ok(AnchorAlignment::Bottom),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "align".into(),
            value: other.into(),
            reason: "unsupported value per OOXML spec".into(),
        }),
    }
}

fn parse_wrap_distance(e: &BytesStart<'_>) -> Result<EdgeInsets<Emu>> {
    let t = xml::optional_attr_i64(e, b"distT")?.unwrap_or(0);
    let b = xml::optional_attr_i64(e, b"distB")?.unwrap_or(0);
    let l = xml::optional_attr_i64(e, b"distL")?.unwrap_or(0);
    let r = xml::optional_attr_i64(e, b"distR")?.unwrap_or(0);
    Ok(EdgeInsets::new(
        Dimension::new(t),
        Dimension::new(r),
        Dimension::new(b),
        Dimension::new(l),
    ))
}

// ── Table ────────────────────────────────────────────────────────────────────

fn parse_table(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Table> {
    let mut tbl_props = TableProperties::default();
    let mut grid = Vec::new();
    let mut rows = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
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
                    _ => xml::warn_unsupported_element("table", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tbl" => break,
            Event::Eof => break,
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
            Event::Eof => break,
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
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"trPr" => {
                        row_props = properties::parse_table_row_properties(reader, buf)?;
                    }
                    b"tc" => {
                        cells.push(parse_table_cell(reader, buf)?);
                    }
                    _ => xml::warn_unsupported_element("table-row", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tr" => break,
            Event::Eof => break,
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
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
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
                    _ => xml::warn_unsupported_element("table-cell", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tc" => break,
            Event::Eof => break,
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
