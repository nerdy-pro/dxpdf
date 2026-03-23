//! Parser for document body content: blocks (paragraphs, tables, section breaks)
//! and inline content (text runs, images, hyperlinks, fields, etc.).

use quick_xml::events::Event;
use quick_xml::Reader;

use quick_xml::events::BytesStart;

use crate::dimension::{Dimension, Emu};
use crate::error::Result;
use crate::geometry::{EdgeInsets, Size};
use crate::model::*;
use crate::xml;
use crate::zip::Relationships;

use super::numbering::NumberingMap;
use super::properties;
use super::styles::ResolvedStyles;

/// Context needed during body parsing for resolving references.
pub struct ParseContext<'a> {
    pub styles: &'a ResolvedStyles,
    pub numbering: &'a NumberingMap,
    pub rels: &'a Relationships,
}

/// Parse a body-level XML part (header, footer, footnotes, etc.) into blocks.
pub fn parse_blocks(data: &[u8], ctx: &ParseContext<'_>) -> Result<Vec<Block>> {
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
                        let (para, sect) = parse_paragraph(&mut reader, &mut buf, ctx, rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" => {
                        blocks.push(Block::Table(Box::new(parse_table(
                            &mut reader,
                            &mut buf,
                            ctx,
                        )?)));
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
pub fn parse_body(data: &[u8], ctx: &ParseContext<'_>) -> Result<(Vec<Block>, SectionProperties)> {
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
                        let (para, sect) = parse_paragraph(&mut reader, &mut buf, ctx, rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" if in_body => {
                        blocks.push(Block::Table(Box::new(parse_table(
                            &mut reader,
                            &mut buf,
                            ctx,
                        )?)));
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
    ctx: &ParseContext<'_>,
    rsids: ParagraphRevisionIds,
) -> Result<(Paragraph, Option<SectionProperties>)> {
    let mut para_props = ParagraphProperties::default();
    let mut run_props_from_ppr: Option<RunProperties> = None;
    let mut style_id: Option<String> = None;
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
                        run_props_from_ppr = parsed.run_properties;
                        section_props = parsed.section_properties;
                    }
                    b"r" => {
                        let run_rsids = parse_run_rsids(e)?;
                        parse_run(reader, buf, ctx, &mut content, run_rsids)?;
                    }
                    b"hyperlink" => {
                        let r_id = xml::optional_attr(e, b"id")?;
                        let anchor = xml::optional_attr(e, b"anchor")?;
                        let hyperlink = parse_hyperlink_content(r_id, anchor, reader, buf, ctx)?;
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
                        let field = parse_simple_field_content(instr, reader, buf, ctx)?;
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

    // Resolve style
    if let Some(ref sid) = style_id {
        if let Some(resolved) = ctx.styles.paragraph.get(sid) {
            let mut base_ppr = resolved.paragraph_properties.clone();
            merge_direct_paragraph(&mut base_ppr, &para_props);
            if let Some(ref direct_rpr) = run_props_from_ppr {
                let mut base_rpr = resolved.run_properties.clone();
                merge_direct_run(&mut base_rpr, direct_rpr);
            }
            para_props = base_ppr;
        }
    }

    Ok((
        Paragraph {
            properties: para_props,
            content,
            rsids,
        },
        section_props,
    ))
}

/// Public wrapper for cross-module use (notes.rs).
/// Returns paragraph + optional trailing section break.
pub fn parse_paragraph_public(
    start_event: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
) -> Result<(Paragraph, Option<SectionProperties>)> {
    let rsids = parse_paragraph_rsids(start_event)?;
    parse_paragraph(reader, buf, ctx, rsids)
}

/// Public wrapper for cross-module use (notes.rs).
pub fn parse_table_public(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
) -> Result<Table> {
    parse_table(reader, buf, ctx)
}

// ── Run ──────────────────────────────────────────────────────────────────────

fn parse_run(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
    content: &mut Vec<Inline>,
    run_rsids: RevisionIds,
) -> Result<()> {
    let mut run_props = RunProperties::default();
    let mut char_style_id: Option<String> = None;
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
                        flush_text(&mut texts, &run_props, &run_rsids, content);
                        let text = xml::read_text_content(reader, buf)?;
                        texts.push(text);
                    }
                    b"drawing" => {
                        flush_text(&mut texts, &run_props, &run_rsids, content);
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
                        flush_text(&mut texts, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::Tab);
                    }
                    b"br" => {
                        flush_text(&mut texts, &run_props, &run_rsids, content);
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
                        flush_text(&mut texts, &run_props, &run_rsids, content);
                        pending_inlines.push(Inline::LineBreak(BreakKind::TextWrapping));
                    }
                    b"sym" => {
                        flush_text(&mut texts, &run_props, &run_rsids, content);
                        let font = xml::optional_attr(e, b"font")?.unwrap_or_default();
                        let char_str = xml::optional_attr(e, b"char")?.unwrap_or_default();
                        let char_code = u16::from_str_radix(&char_str, 16).unwrap_or(0);
                        pending_inlines.push(Inline::Symbol(Symbol { font, char_code }));
                    }
                    b"footnoteReference" => {
                        flush_text(&mut texts, &run_props, &run_rsids, content);
                        if let Some(id) = xml::optional_attr_i64(e, b"id")? {
                            pending_inlines.push(Inline::FootnoteRef(NoteId(id)));
                        }
                    }
                    b"endnoteReference" => {
                        flush_text(&mut texts, &run_props, &run_rsids, content);
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

    // Resolve character style
    if let Some(ref sid) = char_style_id {
        if let Some(resolved_rpr) = ctx.styles.character.get(sid) {
            let mut base = resolved_rpr.clone();
            merge_direct_run(&mut base, &run_props);
            run_props = base;
        }
    }

    flush_text(&mut texts, &run_props, &run_rsids, content);
    content.extend(pending_inlines);

    Ok(())
}

fn flush_text(
    texts: &mut Vec<String>,
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
            properties: props.clone(),
            text: combined,
            rsids: rsids.clone(),
        }));
    }
}

// ── Hyperlink ────────────────────────────────────────────────────────────────

fn parse_hyperlink_content(
    r_id: Option<String>,
    anchor: Option<String>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
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
                parse_run(reader, buf, ctx, &mut inline_content, r_rsids)?;
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
    instr: String,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
) -> Result<Field> {
    let kind = parse_field_kind(&instr);
    let mut field_content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"r" => {
                let r_rsids = parse_run_rsids(e)?;
                parse_run(reader, buf, ctx, &mut field_content, r_rsids)?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"fldSimple" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(Field {
        kind,
        content: field_content,
    })
}

fn parse_field_kind(instr: &str) -> FieldKind {
    let first_word = instr.split_whitespace().next().unwrap_or("").to_uppercase();

    match first_word.as_str() {
        "PAGE" => FieldKind::Page,
        "NUMPAGES" => FieldKind::NumPages,
        "DATE" => FieldKind::Date,
        "TIME" => FieldKind::Time,
        "FILENAME" => FieldKind::FileName,
        "AUTHOR" => FieldKind::Author,
        "TITLE" => FieldKind::Title,
        "TOC" => FieldKind::Toc,
        _ => FieldKind::Other(instr.trim().to_string()),
    }
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
        relative_from: AnchorRelativeFrom::Column,
        offset: Dimension::ZERO,
    };
    let mut v_pos = AnchorPosition::Offset {
        relative_from: AnchorRelativeFrom::Paragraph,
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
                            .unwrap_or(AnchorRelativeFrom::Column);
                        h_pos = parse_anchor_position(reader, buf, rel_from, b"positionH")?;
                    }
                    b"positionV" => {
                        let rel_from = xml::optional_attr(e, b"relativeFrom")?
                            .map(|s| parse_anchor_relative_from(&s))
                            .unwrap_or(AnchorRelativeFrom::Paragraph);
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
    relative_from: AnchorRelativeFrom,
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
                            alignment: parse_anchor_alignment(text.trim()),
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

fn parse_anchor_relative_from(val: &str) -> AnchorRelativeFrom {
    match val {
        "page" => AnchorRelativeFrom::Page,
        "margin" => AnchorRelativeFrom::Margin,
        "column" => AnchorRelativeFrom::Column,
        "character" => AnchorRelativeFrom::Character,
        "paragraph" => AnchorRelativeFrom::Paragraph,
        "line" => AnchorRelativeFrom::Line,
        "insideMargin" => AnchorRelativeFrom::InsideMargin,
        "outsideMargin" => AnchorRelativeFrom::OutsideMargin,
        "topMargin" => AnchorRelativeFrom::TopMargin,
        "bottomMargin" => AnchorRelativeFrom::BottomMargin,
        "leftMargin" => AnchorRelativeFrom::LeftMargin,
        "rightMargin" => AnchorRelativeFrom::RightMargin,
        other => {
            log::warn!("unknown anchor relative-from: {other}");
            AnchorRelativeFrom::Column
        }
    }
}

fn parse_anchor_alignment(val: &str) -> AnchorAlignment {
    match val {
        "left" => AnchorAlignment::Left,
        "center" => AnchorAlignment::Center,
        "right" => AnchorAlignment::Right,
        "inside" => AnchorAlignment::Inside,
        "outside" => AnchorAlignment::Outside,
        "top" => AnchorAlignment::Top,
        "bottom" => AnchorAlignment::Bottom,
        other => {
            log::warn!("unknown anchor alignment: {other}");
            AnchorAlignment::Left
        }
    }
}

fn parse_wrap_distance(e: &quick_xml::events::BytesStart<'_>) -> Result<EdgeInsets<Emu>> {
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

fn parse_table(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
) -> Result<Table> {
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
                        rows.push(parse_table_row(reader, buf, ctx, tr_rsids)?);
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
    ctx: &ParseContext<'_>,
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
                        cells.push(parse_table_cell(reader, buf, ctx)?);
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

fn parse_table_cell(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    ctx: &ParseContext<'_>,
) -> Result<TableCell> {
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
                        let (para, sect) = parse_paragraph(reader, buf, ctx, p_rsids)?;
                        blocks.push(Block::Paragraph(Box::new(para)));
                        if let Some(sp) = sect {
                            blocks.push(Block::SectionBreak(Box::new(sp)));
                        }
                    }
                    b"tbl" => {
                        blocks.push(Block::Table(Box::new(parse_table(reader, buf, ctx)?)));
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

// ── Property merging (direct formatting on top of resolved style) ────────────

fn merge_direct_paragraph(base: &mut ParagraphProperties, direct: &ParagraphProperties) {
    let defaults = ParagraphProperties::default();
    if direct.alignment != defaults.alignment {
        base.alignment = direct.alignment;
    }
    if direct.indentation != defaults.indentation {
        base.indentation = direct.indentation;
    }
    if direct.spacing != defaults.spacing {
        base.spacing = direct.spacing;
    }
    if direct.numbering.is_some() {
        base.numbering = direct.numbering.clone();
    }
    if !direct.tabs.is_empty() {
        base.tabs = direct.tabs.clone();
    }
    if direct.borders.is_some() {
        base.borders = direct.borders.clone();
    }
    if direct.shading.is_some() {
        base.shading = direct.shading.clone();
    }
    if direct.keep_next {
        base.keep_next = true;
    }
    if direct.keep_lines {
        base.keep_lines = true;
    }
    if !direct.widow_control {
        base.widow_control = false;
    }
    if direct.page_break_before {
        base.page_break_before = true;
    }
    if direct.bidi {
        base.bidi = true;
    }
    if direct.outline_level.is_some() {
        base.outline_level = direct.outline_level;
    }
}

fn merge_direct_run(base: &mut RunProperties, direct: &RunProperties) {
    let defaults = RunProperties::default();
    if direct.fonts.ascii.is_some() {
        base.fonts.ascii = direct.fonts.ascii.clone();
    }
    if direct.fonts.high_ansi.is_some() {
        base.fonts.high_ansi = direct.fonts.high_ansi.clone();
    }
    if direct.fonts.east_asian.is_some() {
        base.fonts.east_asian = direct.fonts.east_asian.clone();
    }
    if direct.fonts.complex_script.is_some() {
        base.fonts.complex_script = direct.fonts.complex_script.clone();
    }
    if direct.font_size != defaults.font_size {
        base.font_size = direct.font_size;
    }
    if direct.bold != defaults.bold {
        base.bold = direct.bold;
    }
    if direct.italic != defaults.italic {
        base.italic = direct.italic;
    }
    if direct.underline != defaults.underline {
        base.underline = direct.underline;
    }
    if direct.strike != defaults.strike {
        base.strike = direct.strike;
    }
    if direct.color != defaults.color {
        base.color = direct.color;
    }
    if direct.highlight.is_some() {
        base.highlight = direct.highlight;
    }
    if direct.shading.is_some() {
        base.shading = direct.shading.clone();
    }
    if direct.vertical_align != defaults.vertical_align {
        base.vertical_align = direct.vertical_align;
    }
    if direct.spacing != defaults.spacing {
        base.spacing = direct.spacing;
    }
    if direct.kerning.is_some() {
        base.kerning = direct.kerning;
    }
    if direct.all_caps {
        base.all_caps = true;
    }
    if direct.small_caps {
        base.small_caps = true;
    }
    if direct.vanish {
        base.vanish = true;
    }
    if direct.rtl {
        base.rtl = true;
    }
}
