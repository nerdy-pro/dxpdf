//! Parser for document body content: blocks (paragraphs, tables, section breaks)
//! and inline content (text runs, images, hyperlinks, fields, etc.).
//!
//! Two-pass approach: the first pass is a focused event-driven scan that
//! extracts `<w:drawing>` and `<w:pict>` sub-trees via the legacy DrawingML
//! and VML parsers (neither of which is serde-ified yet); the second pass
//! uses serde over the whole document and fills in the parsed
//! drawings/picts in document order via an iterator context.
//!
//! No style resolution or property merging — output is raw parsed data.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::parse::body_schema::*;
use crate::docx::parse::serde_xml::from_xml;
use crate::docx::xml;

use super::{drawing, vml};

/// Parse `w:document > w:body`, returning blocks and final section properties.
pub fn parse_body(data: &[u8]) -> Result<(Vec<Block>, SectionProperties)> {
    if data.is_empty() {
        return Ok((Vec::new(), SectionProperties::default()));
    }
    let embeds = extract_embeds(data, b"body")?;
    let doc: DocXml = from_xml(data)?;
    let mut ctx = ConvertCtx::new(embeds);
    let (blocks, final_section) = convert_container(doc.body.children, &mut ctx);
    Ok((blocks, final_section.unwrap_or_default()))
}

/// Parse a body-level XML part (header, footer, footnote body, etc.) into blocks.
pub fn parse_blocks(data: &[u8]) -> Result<Vec<Block>> {
    if data.is_empty() {
        return Ok(Vec::new());
    }
    // Root element varies (hdr/ftr/footnote/etc.) — extract embeds scoped to
    // whatever top-level element contains body content; drawings/picts live
    // only at inline positions so a straight scan works.
    let embeds = extract_embeds(data, b"")?;
    let container: BlockContainerXml = from_xml(data)?;
    let mut ctx = ConvertCtx::new(embeds);
    let (blocks, _) = convert_container(container.children, &mut ctx);
    Ok(blocks)
}

// ── Two-pass: pre-extract <w:drawing> / <w:pict> sub-trees ────────────────

#[derive(Debug)]
#[allow(clippy::large_enum_variant)]
pub(crate) enum Embed {
    Drawing(Option<Image>),
    Pict(Pict),
}

/// Iterate the raw XML in document order. For each top-level `<w:drawing>`
/// or `<w:pict>` found, hand off to the legacy parser and record the result.
/// The result vector is consumed in the same document order by the serde
/// post-pass.
///
/// `body_tag` narrows the scan to contents of that root; pass `b""` to scan
/// the whole document (used for headers/footers where the root varies).
pub(crate) fn extract_embeds(data: &[u8], body_tag: &[u8]) -> Result<Vec<Embed>> {
    let mut reader = Reader::from_reader(data);
    // For the pre-pass we want whitespace-preserving, like serde does.
    let mut buf = Vec::new();
    let mut out = Vec::new();

    // If a body_tag is specified, advance to it first.
    if !body_tag.is_empty() {
        loop {
            match xml::next_event(&mut reader, &mut buf)? {
                Event::Start(ref e) if xml::local_name(e.name().as_ref()) == body_tag => break,
                Event::Eof => return Ok(out),
                _ => {}
            }
        }
    }

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local: Vec<u8> = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"drawing" => {
                        out.push(Embed::Drawing(parse_drawing_sub_tree(
                            &mut reader,
                            &mut buf,
                        )?));
                    }
                    b"pict" => {
                        out.push(Embed::Pict(vml::parse_pict(&mut reader, &mut buf)?));
                    }
                    _ => {}
                }
            }
            Event::End(ref e) if !body_tag.is_empty() && xml::local_name(e.name().as_ref()) == body_tag => {
                break;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(out)
}

/// Mirror of legacy `parse_drawing` — finds inline/anchor child and produces
/// an `Image`.
fn parse_drawing_sub_tree(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<Image>> {
    let mut image: Option<Image> = None;
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local: Vec<u8> = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"inline" => image = drawing::parse_inline_image(e, reader, buf)?,
                    b"anchor" => image = drawing::parse_anchor_image(e, reader, buf)?,
                    _ => xml::skip_to_end(reader, buf, &local)?,
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"drawing" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"drawing")),
            _ => {}
        }
    }
    Ok(image)
}

// ── Top-level document schema wrapper ────────────────────────────────────

use serde::Deserialize;

/// Thin wrapper for `<w:document>` — just extracts `<w:body>`.
#[derive(Deserialize)]
struct DocXml {
    body: BlockContainerXml,
}

// ── Conversion ────────────────────────────────────────────────────────────

pub(crate) struct ConvertCtx {
    embeds: std::vec::IntoIter<Embed>,
}

impl ConvertCtx {
    pub(crate) fn new(embeds: Vec<Embed>) -> Self {
        Self {
            embeds: embeds.into_iter(),
        }
    }
    fn next_embed(&mut self) -> Option<Embed> {
        self.embeds.next()
    }
}

/// Convert a list of block-level children into `(Vec<Block>, Option<SectionProperties>)`.
/// The section properties, if returned, are for a trailing `<w:sectPr>` at
/// this level — the final section for `<w:body>`, or one that appears inside
/// a table cell (§17.6.17).
pub(crate) fn convert_container(
    children: Vec<BlockChildXml>,
    ctx: &mut ConvertCtx,
) -> (Vec<Block>, Option<SectionProperties>) {
    let mut blocks = Vec::new();
    let mut final_section = None;
    for child in children {
        match child {
            BlockChildXml::Paragraph(p) => {
                let (para, sect_after) = convert_paragraph(*p, ctx);
                blocks.push(Block::Paragraph(Box::new(para)));
                if let Some(sp) = sect_after {
                    blocks.push(Block::SectionBreak(Box::new(sp)));
                }
            }
            BlockChildXml::Table(t) => {
                blocks.push(Block::Table(Box::new(convert_table(*t, ctx))));
            }
            BlockChildXml::SectPr(sp) => {
                final_section = Some(SectionProperties::from(*sp));
            }
            BlockChildXml::Sdt(sdt) => {
                // Flatten SDT wrapper — treat its content as block-level.
                if let Some(content) = sdt.content {
                    let (nested_blocks, nested_sect) =
                        convert_container(content.children, ctx);
                    blocks.extend(nested_blocks);
                    if nested_sect.is_some() {
                        final_section = nested_sect;
                    }
                }
            }
            // Block-level markers and ignored elements — renderer doesn't use them.
            BlockChildXml::BookmarkStart(_)
            | BlockChildXml::BookmarkEnd(_)
            | BlockChildXml::CommentRangeStart(_)
            | BlockChildXml::CommentRangeEnd(_)
            | BlockChildXml::ProofErr(_)
            | BlockChildXml::Other => {}
        }
    }
    (blocks, final_section)
}

fn convert_paragraph(p: ParaXml, ctx: &mut ConvertCtx) -> (Paragraph, Option<SectionProperties>) {
    let rsids = ParagraphRevisionIds {
        r: hex_rsid(p.rsid_r.as_deref()),
        r_default: hex_rsid(p.rsid_r_default.as_deref()),
        p: hex_rsid(p.rsid_p.as_deref()),
        r_pr: hex_rsid(p.rsid_r_pr.as_deref()),
        del: hex_rsid(p.rsid_del.as_deref()),
    };

    // pPr may appear as either the dedicated field OR inside $value (serde
    // collects all matching children; since `pPr` is named on the struct
    // *and* in the enum, serde prefers the dedicated field — but just in
    // case, we merge from both sources).
    let p_pr = p.p_pr.or_else(|| {
        p.content
            .iter()
            .find_map(|c| if let ParaChildXml::PPr(pp) = c { Some((**pp).clone()) } else { None })
    });

    let parsed_p_pr = p_pr.map(|pp| pp.split());
    let (style_id, properties, mark_run_properties, section_properties) = match parsed_p_pr {
        Some(pp) => (
            pp.style_id,
            pp.properties,
            pp.run_properties,
            pp.section_properties,
        ),
        None => (None, ParagraphProperties::default(), None, None),
    };

    let mut content = Vec::new();
    for child in p.content {
        match child {
            ParaChildXml::Run(r) => extend_from_run(r, &mut content, ctx),
            ParaChildXml::Hyperlink(h) => {
                content.push(Inline::Hyperlink(convert_hyperlink(h, ctx)));
            }
            ParaChildXml::FldSimple(f) => {
                content.push(Inline::Field(convert_fld_simple(f, ctx)));
            }
            ParaChildXml::BookmarkStart(b) => content.push(Inline::BookmarkStart {
                id: BookmarkId::new(b.id),
                name: b.name,
            }),
            ParaChildXml::BookmarkEnd(b) => {
                content.push(Inline::BookmarkEnd(BookmarkId::new(b.id)));
            }
            ParaChildXml::PPr(_) => {} // already captured above
        }
    }

    (
        Paragraph {
            style_id,
            properties,
            mark_run_properties,
            content,
            rsids,
        },
        section_properties,
    )
}

/// Flatten a `RunXml`'s children into zero-or-more `Inline`s and append to
/// the parent content. Text / tab / br / cr / lastRenderedPageBreak are
/// accumulated into one `Inline::TextRun`; sibling inlines flush the accumulator
/// and append independently.
fn extend_from_run(r: RunXml, out: &mut Vec<Inline>, ctx: &mut ConvertCtx) {
    let rsids = RevisionIds {
        r: hex_rsid(r.rsid_r.as_deref()),
        r_pr: hex_rsid(r.rsid_r_pr.as_deref()),
        del: hex_rsid(r.rsid_del.as_deref()),
    };
    let (props, style_id) = r
        .r_pr
        .map(|rp| rp.split())
        .unwrap_or_default();

    let mut acc: Vec<RunElement> = Vec::new();
    let flush = |acc: &mut Vec<RunElement>, out: &mut Vec<Inline>| {
        if !acc.is_empty() {
            out.push(Inline::TextRun(Box::new(TextRun {
                style_id: style_id.clone(),
                properties: props.clone(),
                content: std::mem::take(acc),
                rsids,
            })));
        }
    };

    for child in r.content {
        match child {
            RunChildXml::Text(t) => acc.push(RunElement::Text(t.content)),
            RunChildXml::DelText(t) => acc.push(RunElement::Text(t.content)),
            RunChildXml::Tab => acc.push(RunElement::Tab),
            RunChildXml::Br(br) => acc.push(run_break(br)),
            RunChildXml::Cr => acc.push(RunElement::LineBreak(BreakKind::TextWrapping)),
            RunChildXml::LastRenderedPageBreak => acc.push(RunElement::LastRenderedPageBreak),
            RunChildXml::Drawing(_) => {
                flush(&mut acc, out);
                if let Some(Embed::Drawing(Some(img))) = ctx.next_embed() {
                    out.push(Inline::Image(Box::new(img)));
                }
            }
            RunChildXml::Pict(_) => {
                flush(&mut acc, out);
                if let Some(Embed::Pict(pict)) = ctx.next_embed() {
                    out.push(Inline::Pict(pict));
                }
            }
            RunChildXml::Sym(s) => {
                flush(&mut acc, out);
                let char_code = u16::from_str_radix(&s.char, 16).unwrap_or(0);
                out.push(Inline::Symbol(Symbol {
                    font: s.font,
                    char_code,
                }));
            }
            RunChildXml::InstrText(t) => {
                flush(&mut acc, out);
                out.push(Inline::InstrText(t.content));
            }
            RunChildXml::FldChar(fc) => {
                flush(&mut acc, out);
                out.push(Inline::FieldChar(FieldChar {
                    field_char_type: FieldCharType::from(fc.fld_char_type),
                    dirty: fc.dirty.map(|b| b.0),
                    fld_lock: fc.fld_lock.map(|b| b.0),
                }));
            }
            RunChildXml::FootnoteRef(n) => {
                flush(&mut acc, out);
                out.push(Inline::FootnoteRef(NoteId::new(n.id)));
            }
            RunChildXml::EndnoteRef(n) => {
                flush(&mut acc, out);
                out.push(Inline::EndnoteRef(NoteId::new(n.id)));
            }
            RunChildXml::FootnoteRefMark => {
                flush(&mut acc, out);
                out.push(Inline::FootnoteRefMark);
            }
            RunChildXml::EndnoteRefMark => {
                flush(&mut acc, out);
                out.push(Inline::EndnoteRefMark);
            }
            RunChildXml::Separator => {
                flush(&mut acc, out);
                out.push(Inline::Separator);
            }
            RunChildXml::ContinuationSeparator => {
                flush(&mut acc, out);
                out.push(Inline::ContinuationSeparator);
            }
            RunChildXml::AlternateContent(ac) => {
                flush(&mut acc, out);
                out.push(Inline::AlternateContent(convert_alt_content(ac, ctx)));
            }
            RunChildXml::RPr(_) => {} // already captured via r.r_pr
        }
    }
    flush(&mut acc, out);
}

fn run_break(br: BrXml) -> RunElement {
    use crate::docx::parse::body_schema::StBrType;
    match br.ty {
        Some(StBrType::Page) => RunElement::PageBreak,
        Some(StBrType::Column) => RunElement::ColumnBreak,
        _ => {
            let clear = br.clear.map(BreakClear::from).unwrap_or(BreakClear::None);
            if clear != BreakClear::None {
                RunElement::LineBreak(BreakKind::Clear(clear))
            } else {
                RunElement::LineBreak(BreakKind::TextWrapping)
            }
        }
    }
}

fn convert_hyperlink(h: HyperlinkXml, ctx: &mut ConvertCtx) -> Hyperlink {
    let target = if let Some(id) = h.r_id {
        HyperlinkTarget::External(RelId::new(id))
    } else {
        HyperlinkTarget::Internal {
            anchor: h.anchor.unwrap_or_default(),
        }
    };
    let mut content = Vec::new();
    for child in h.content {
        match child {
            ParaChildXml::Run(r) => extend_from_run(r, &mut content, ctx),
            ParaChildXml::Hyperlink(nested) => {
                content.push(Inline::Hyperlink(convert_hyperlink(nested, ctx)));
            }
            ParaChildXml::FldSimple(f) => {
                content.push(Inline::Field(convert_fld_simple(f, ctx)));
            }
            ParaChildXml::BookmarkStart(b) => content.push(Inline::BookmarkStart {
                id: BookmarkId::new(b.id),
                name: b.name,
            }),
            ParaChildXml::BookmarkEnd(b) => {
                content.push(Inline::BookmarkEnd(BookmarkId::new(b.id)));
            }
            ParaChildXml::PPr(_) => {}
        }
    }
    Hyperlink { target, content }
}

fn convert_fld_simple(f: FldSimpleXml, ctx: &mut ConvertCtx) -> Field {
    let instruction = match crate::field::parse(&f.instr) {
        Ok(i) => i,
        Err(e) => {
            log::warn!("failed to parse field instruction {:?}: {}", f.instr, e);
            crate::field::FieldInstruction::Unknown {
                field_type: String::new(),
                raw: f.instr.clone(),
            }
        }
    };
    let mut content = Vec::new();
    for child in f.content {
        match child {
            ParaChildXml::Run(r) => extend_from_run(r, &mut content, ctx),
            ParaChildXml::Hyperlink(h) => {
                content.push(Inline::Hyperlink(convert_hyperlink(h, ctx)));
            }
            ParaChildXml::FldSimple(inner) => {
                content.push(Inline::Field(convert_fld_simple(inner, ctx)));
            }
            ParaChildXml::BookmarkStart(b) => content.push(Inline::BookmarkStart {
                id: BookmarkId::new(b.id),
                name: b.name,
            }),
            ParaChildXml::BookmarkEnd(b) => {
                content.push(Inline::BookmarkEnd(BookmarkId::new(b.id)));
            }
            ParaChildXml::PPr(_) => {}
        }
    }
    Field {
        instruction,
        content,
    }
}

fn convert_alt_content(a: AltContentXml, ctx: &mut ConvertCtx) -> AlternateContent {
    let choices = a
        .choices
        .into_iter()
        .filter_map(|c| {
            let requires = mc_requires(&c.requires)?;
            let content = convert_mc_content(c.content, ctx);
            Some(McChoice { requires, content })
        })
        .collect();
    let fallback = a
        .fallback
        .map(|f| convert_mc_content(f.content, ctx));
    AlternateContent { choices, fallback }
}

fn mc_requires(s: &str) -> Option<McRequires> {
    match s {
        "wps" => Some(McRequires::Wps),
        "wpg" => Some(McRequires::Wpg),
        "wpc" => Some(McRequires::Wpc),
        "wpi" => Some(McRequires::Wpi),
        "m" => Some(McRequires::Math),
        "a14" => Some(McRequires::A14),
        "w14" => Some(McRequires::W14),
        "w15" => Some(McRequires::W15),
        "w16" => Some(McRequires::W16),
        other => {
            log::warn!("mc:Choice: unsupported Requires {:?}", other);
            None
        }
    }
}

fn convert_mc_content(items: Vec<McContentXml>, ctx: &mut ConvertCtx) -> Vec<Inline> {
    let mut out = Vec::new();
    for i in items {
        match i {
            McContentXml::Drawing(_) => {
                if let Some(Embed::Drawing(Some(img))) = ctx.next_embed() {
                    out.push(Inline::Image(Box::new(img)));
                }
            }
            McContentXml::Pict(_) => {
                if let Some(Embed::Pict(p)) = ctx.next_embed() {
                    out.push(Inline::Pict(p));
                }
            }
        }
    }
    out
}

fn convert_table(t: TableXml, ctx: &mut ConvertCtx) -> Table {
    let (properties, _style_id) = t.tbl_pr.map(|tp| tp.split()).unwrap_or_default();
    let grid = t
        .tbl_grid
        .map(|g| {
            g.cols
                .into_iter()
                .map(|c| GridColumn {
                    width: c.w.unwrap_or_default(),
                })
                .collect()
        })
        .unwrap_or_default();
    let rows = t
        .rows
        .into_iter()
        .map(|r| convert_table_row(r, ctx))
        .collect();
    Table {
        properties,
        grid,
        rows,
    }
}

fn convert_table_row(r: TableRowXml, ctx: &mut ConvertCtx) -> TableRow {
    let rsids = TableRowRevisionIds {
        r: hex_rsid(r.rsid_r.as_deref()),
        r_pr: hex_rsid(r.rsid_r_pr.as_deref()),
        del: hex_rsid(r.rsid_del.as_deref()),
        tr: hex_rsid(r.rsid_tr.as_deref()),
    };
    let properties = r.tr_pr.map(TableRowProperties::from).unwrap_or_default();
    let cells = r
        .cells
        .into_iter()
        .map(|c| convert_table_cell(c, ctx))
        .collect();
    TableRow {
        properties,
        cells,
        rsids,
    }
}

fn convert_table_cell(c: TableCellXml, ctx: &mut ConvertCtx) -> TableCell {
    let properties = c.tc_pr.map(TableCellProperties::from).unwrap_or_default();
    let (content, _final_sect) = convert_container(c.content, ctx);
    TableCell {
        properties,
        content,
    }
}

// ── helpers ──────────────────────────────────────────────────────────────

fn hex_rsid(s: Option<&str>) -> Option<RevisionSaveId> {
    s.and_then(RevisionSaveId::from_hex)
}
