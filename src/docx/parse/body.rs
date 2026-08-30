//! Parser for document body content: blocks (paragraphs, tables, section breaks)
//! and inline content (text runs, images, hyperlinks, fields, etc.).
//!
//! Single-pass serde over the full document. Drawings and VML picts are
//! serde-parsed inline via `DrawingXml` / `PictXml`; they produce their
//! model values (`Image` / `Pict`) during the `convert_container` walk via
//! the `ConvertCtx`.
//!
//! No style resolution or property merging — output is raw parsed data.

use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::parse::body_schema::*;
use crate::docx::parse::serde_xml::from_xml;
use crate::docx::whitespace_workaround::restore_whitespace_sentinels;
use crate::model::Dup;

/// Parse `w:document > w:body`, returning blocks and final section properties.
pub fn parse_body(data: &[u8]) -> Result<(Vec<Block>, SectionProperties)> {
    if data.is_empty() {
        return Ok((Vec::new(), SectionProperties::default()));
    }
    let doc: DocXml = from_xml(data)?;
    let mut ctx = ConvertCtx::new();
    let (blocks, final_section) = convert_container(doc.body.children, &mut ctx);
    Ok((blocks, final_section.unwrap_or_default()))
}

/// Parse a body-level XML part (header, footer, footnote body, etc.) into blocks.
pub fn parse_blocks(data: &[u8]) -> Result<Vec<Block>> {
    if data.is_empty() {
        return Ok(Vec::new());
    }
    let container: BlockContainerXml = from_xml(data)?;
    let mut ctx = ConvertCtx::new();
    let (blocks, _) = convert_container(container.children, &mut ctx);
    Ok(blocks)
}

// ── Top-level document schema wrapper ────────────────────────────────────

use serde::Deserialize;

/// Thin wrapper for `<w:document>` — just extracts `<w:body>`.
#[derive(Deserialize)]
struct DocXml {
    body: BlockContainerXml,
}

// ── Conversion ────────────────────────────────────────────────────────────

/// Conversion context. Previously carried a pre-pass iterator of parsed
/// drawings/picts; now empty, since drawings and picts are serde-parsed
/// inline. Kept as a type for future extensibility (e.g., if a later phase
/// needs cross-node state during conversion).
pub(crate) struct ConvertCtx {
    /// §17.13.5.14/.18: the innermost `<w:ins>`/`<w:del>` wrapper the walk is
    /// currently inside, stamped onto every run it reaches. Pushed and popped
    /// by `append_revision_children`; a nested wrapper shadows the outer one,
    /// so text inserted and then deleted carries the deletion.
    revision: Option<RunRevision>,
    /// Issue #154: comment ranges open at this point of the walk, in opening
    /// order — `commentRangeStart` pushes, `commentRangeEnd` removes. Runs
    /// between the markers are stamped with the first-opened id. Lives on the
    /// ctx because a range can span paragraphs (and, at block level, tables).
    open_comments: Vec<CommentId>,
}

impl ConvertCtx {
    pub(crate) fn new() -> Self {
        Self {
            revision: None,
            open_comments: Vec::new(),
        }
    }

    fn open_comment(&mut self, id: CommentId) {
        if !self.open_comments.contains(&id) {
            self.open_comments.push(id);
        }
    }

    fn close_comment(&mut self, id: CommentId) {
        self.open_comments.retain(|c| *c != id);
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
                if let Some(content) = Dup::from(sdt.content).into_value() {
                    let (nested_blocks, nested_sect) = convert_container(content.children, ctx);
                    blocks.extend(nested_blocks);
                    if nested_sect.is_some() {
                        final_section = nested_sect;
                    }
                }
            }
            // Issue #154: block-level comment range markers toggle the same
            // stamp the paragraph-level ones do — a range can open between
            // paragraphs and close inside one, or span a table.
            BlockChildXml::CommentRangeStart(c) => ctx.open_comment(CommentId::new(c.id)),
            BlockChildXml::CommentRangeEnd(c) => ctx.close_comment(CommentId::new(c.id)),
            // Block-level markers and ignored elements — renderer doesn't use them.
            BlockChildXml::BookmarkStart(_)
            | BlockChildXml::BookmarkEnd(_)
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
    let p_pr = Dup::from(p.p_pr).into_value().or_else(|| {
        p.content.iter().find_map(|c| {
            if let ParaChildXml::PPr(pp) = c {
                Some((**pp).clone())
            } else {
                None
            }
        })
    });

    let parsed_p_pr = p_pr.map(|pp| pp.split());
    let (style_id, properties, mark_run_properties, section_properties, mark_deleted) =
        match parsed_p_pr {
            Some(pp) => (
                pp.style_id,
                pp.properties,
                pp.run_properties,
                pp.section_properties,
                pp.mark_deleted,
            ),
            None => (None, ParagraphProperties::default(), None, None, false),
        };

    let content = convert_para_children(p.content, ctx);

    (
        Paragraph {
            style_id,
            properties,
            mark_run_properties,
            content,
            rsids,
            mark_deleted,
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
    let (props, style_id) = Dup::from(r.r_pr)
        .into_value()
        .map(|rp| rp.split())
        .unwrap_or_default();

    let mut acc: Vec<RunElement> = Vec::new();
    let revision = ctx.revision.clone();
    let comment = ctx.open_comments.first().copied();
    let flush = |acc: &mut Vec<RunElement>, out: &mut Vec<Inline>| {
        if !acc.is_empty() {
            out.push(Inline::TextRun(Box::new(TextRun {
                style_id: style_id.clone(),
                properties: props.clone(),
                content: std::mem::take(acc),
                rsids,
                revision: revision.clone(),
                comment,
            })));
        }
    };

    for child in r.content {
        match child {
            RunChildXml::Text(t) => {
                acc.push(RunElement::Text(restore_whitespace_sentinels(&t.content)))
            }
            RunChildXml::DelText(t) => {
                acc.push(RunElement::Text(restore_whitespace_sentinels(&t.content)))
            }
            RunChildXml::Tab => acc.push(RunElement::Tab),
            RunChildXml::PTab(p) => acc.push(RunElement::PositionTab(p.into())),
            RunChildXml::Br(br) => acc.push(run_break(br)),
            RunChildXml::Cr => acc.push(RunElement::LineBreak(BreakKind::TextWrapping)),
            RunChildXml::SoftHyphen => {} // optional hyphen — only visible if line breaks here; we don't hyphenate
            RunChildXml::NoBreakHyphen => acc.push(RunElement::Text("\u{2011}".to_string())),
            RunChildXml::LastRenderedPageBreak => acc.push(RunElement::LastRenderedPageBreak),
            RunChildXml::Drawing(d) => {
                flush(&mut acc, out);
                if let Some(img) = drawing_to_image(d, ctx) {
                    out.push(Inline::Image(Box::new(img)));
                }
            }
            RunChildXml::Pict(p) => {
                flush(&mut acc, out);
                out.push(Inline::Pict(p.into_model(ctx)));
            }
            RunChildXml::Sym(s) => {
                flush(&mut acc, out);
                let char_code = u16::from_str_radix(&s.char, 16).unwrap_or_else(|_| {
                    log::warn!("sym: invalid hex char code {:?}; using 0", s.char);
                    0
                });
                out.push(Inline::Symbol(Symbol {
                    font: s.font,
                    char_code,
                }));
            }
            RunChildXml::InstrText(t) | RunChildXml::DelInstrText(t) => {
                // §17.16.13: `delInstrText` is the deleted-field spelling of
                // the same payload; the enclosing `<w:del>`'s filter decides
                // whether the field survives, not this arm.
                flush(&mut acc, out);
                out.push(Inline::InstrText(restore_whitespace_sentinels(&t.content)));
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
            RunChildXml::AnnotationRef(_) => {}
            RunChildXml::CommentReference(c) => {
                // Issue #154: the balloon's anchor. Emitted between text
                // accumulations, like a bookmark. Word always writes the
                // range's end before the reference, so the reference also
                // closes the range — the repair for a document that lost its
                // `commentRangeEnd`, without which the wash would run to the
                // end of the part.
                flush(&mut acc, out);
                let id = CommentId::new(c.id);
                ctx.close_comment(id);
                out.push(Inline::CommentRef(id));
            }
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

/// Convert a sequence of paragraph-level children (`EG_PContent`) into inline
/// content. Shared by `<w:p>`, `<w:hyperlink>`, and `<w:fldSimple>`, and used
/// recursively to flatten revision/structural wrappers.
///
/// Tracked changes (issue #154): `<w:ins>`/`<w:del>` wrappers are flattened
/// with their runs stamped [`RunRevision`], keeping unaccepted deletions in
/// the model — the renderer's `w:revisionView` decision (marked, plain, or
/// suppressed) needs them there. Structural wrappers (`<w:smartTag>`,
/// `<w:customXml>`) and `<w:moveTo>` are flattened unmarked; `<w:moveFrom>`
/// (the source copy of moved text) is dropped so the move isn't duplicated.
/// Nested wrappers shadow: text inserted then deleted carries the deletion.
fn convert_para_children(children: Vec<ParaChildXml>, ctx: &mut ConvertCtx) -> Vec<Inline> {
    let mut content = Vec::new();
    append_para_children(children, &mut content, ctx);
    content
}

fn append_para_children(
    children: Vec<ParaChildXml>,
    content: &mut Vec<Inline>,
    ctx: &mut ConvertCtx,
) {
    for child in children {
        match child {
            ParaChildXml::Run(r) => extend_from_run(r, content, ctx),
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
            ParaChildXml::CommentRangeStart(c) => ctx.open_comment(CommentId::new(c.id)),
            ParaChildXml::CommentRangeEnd(c) => ctx.close_comment(CommentId::new(c.id)),
            // Tracked-change wrappers (issue #154): flattened, with each run
            // inside stamped as inserted or deleted — which way a stamped run
            // reaches the page (marked, plain, or suppressed) is the
            // renderer's `w:revisionView` decision, not the parser's.
            ParaChildXml::Ins(w) => {
                append_revision_children(w, RevisionKind::Inserted, content, ctx)
            }
            ParaChildXml::Del(w) => {
                append_revision_children(w, RevisionKind::Deleted, content, ctx)
            }
            // Structural wrappers: flatten and render. `<w:moveTo>` is the
            // destination copy of moved text — part of the final document and
            // not marked (Word's move marks are their own display family).
            ParaChildXml::MoveTo(w) | ParaChildXml::SmartTag(w) | ParaChildXml::CustomXml(w) => {
                append_para_children(w.content, content, ctx)
            }
            // The source copy of moved text: rendering both copies would
            // duplicate it, so the source stays dropped in every view.
            ParaChildXml::MoveFrom(_) => {}
            ParaChildXml::PPr(_) => {} // already captured on the parent
            ParaChildXml::Other => {}
        }
    }
}

/// Flatten one `<w:ins>`/`<w:del>` wrapper, stamping every run inside it.
///
/// Deleted content that cannot carry the stamp — an image or VML picture is
/// its own `Inline`, not a run — is dropped in every view, exactly as the
/// whole wrapper was before #154: a deleted picture never reaches the page,
/// while deleted *text* now survives to be struck through or suppressed by
/// the renderer. Hyperlinks and fields keep their frames (their runs carry
/// the stamp) but shed unstampable children by the same rule.
fn append_revision_children(
    w: RunTrackChangeXml,
    kind: RevisionKind,
    content: &mut Vec<Inline>,
    ctx: &mut ConvertCtx,
) {
    let outer = ctx.revision.take();
    // §17.13.5.14: deletion is sticky — text inserted and then deleted is a
    // deletion, whichever way the wrappers nest. An inner `<w:ins>` inside a
    // `<w:del>` therefore keeps the outer deletion stamp; every other nesting
    // lets the inner wrapper shadow the outer one.
    if matches!(&outer, Some(o) if o.kind == RevisionKind::Deleted) {
        ctx.revision = outer.clone();
    } else {
        ctx.revision = Some(RunRevision {
            kind,
            author: w.author.unwrap_or_default(),
        });
    }
    let effective_kind = ctx.revision.as_ref().map(|r| r.kind).unwrap_or(kind);
    let mut inner = Vec::new();
    append_para_children(w.content, &mut inner, ctx);
    if effective_kind == RevisionKind::Deleted {
        drop_unstampable_deleted(&mut inner);
    }
    content.extend(inner);
    ctx.revision = outer;
}

fn drop_unstampable_deleted(inlines: &mut Vec<Inline>) {
    inlines.retain_mut(|inline| match inline {
        // Stamped, or a frame whose runs are: the renderer's view decides.
        Inline::TextRun(_) => true,
        Inline::Hyperlink(h) => {
            drop_unstampable_deleted(&mut h.content);
            true
        }
        // Invisible markers: a bookmark or comment survives the deletion of
        // the text it spans until the deletion is accepted.
        Inline::BookmarkStart { .. } | Inline::BookmarkEnd(_) | Inline::CommentRef(_) => true,
        // Everything visible that cannot carry the run stamp is dropped in
        // every view — exactly what the whole wrapper got before #154, now
        // scoped to what a run cannot express. That includes fields: a
        // deleted dynamic field would otherwise re-evaluate and paint its
        // fresh value in the final view, resurrecting deleted content.
        Inline::Image(_)
        | Inline::Pict(_)
        | Inline::AlternateContent(_)
        | Inline::Symbol(_)
        | Inline::FootnoteRef(_)
        | Inline::EndnoteRef(_)
        | Inline::Field(_)
        | Inline::FieldChar(_)
        | Inline::InstrText(_)
        | Inline::Separator
        | Inline::ContinuationSeparator
        | Inline::FootnoteRefMark
        | Inline::EndnoteRefMark => false,
    });
}

fn convert_hyperlink(h: HyperlinkXml, ctx: &mut ConvertCtx) -> Hyperlink {
    let target = if let Some(id) = h.r_id {
        HyperlinkTarget::ExternalRel(RelId::new(id))
    } else {
        HyperlinkTarget::Internal {
            anchor: h.anchor.unwrap_or_default(),
        }
    };
    let content = convert_para_children(h.content, ctx);
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
    let content = convert_para_children(f.content, ctx);
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
            // §M.2.2: @Requires is a space-separated list of namespace prefixes;
            // the choice is usable only if we understand *all* of them. A single
            // unknown token drops the whole choice (falling through to the next
            // choice or the fallback).
            let requires: Vec<McRequires> = c
                .requires
                .split_whitespace()
                .map(mc_requires)
                .collect::<Option<_>>()?;
            let content = convert_mc_content(c.content, ctx);
            Some(McChoice { requires, content })
        })
        .collect();
    let fallback = Dup::from(a.fallback)
        .into_value()
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
            McContentXml::Drawing(d) => {
                if let Some(img) = drawing_to_image(d, ctx) {
                    out.push(Inline::Image(Box::new(img)));
                }
            }
            McContentXml::Pict(p) => {
                out.push(Inline::Pict(p.into_model(ctx)));
            }
        }
    }
    out
}

/// Convert a serde-parsed `<w:drawing>` into the model's `Image`. Returns
/// `None` when neither `<wp:inline>` nor `<wp:anchor>` is present.
fn drawing_to_image(
    d: crate::docx::parse::body_schema::DrawingXml,
    ctx: &mut ConvertCtx,
) -> Option<Image> {
    if let Some(inline) = Dup::from(d.inline).into_value() {
        return Some(inline.into_image(ctx));
    }
    Dup::from(d.anchor).into_value().map(|a| a.into_image(ctx))
}

fn convert_table(t: TableXml, ctx: &mut ConvertCtx) -> Table {
    let (properties, _style_id) = Dup::from(t.tbl_pr)
        .into_value()
        .map(|tp| tp.split())
        .unwrap_or_default();
    let grid = Dup::from(t.tbl_grid)
        .into_value()
        .map(|g| {
            g.cols
                .into_iter()
                .map(|c| GridColumn {
                    width: c.w.unwrap_or_default(),
                })
                .collect()
        })
        .unwrap_or_default();
    let mut rows = Vec::new();
    for item in collect_table_rows(t.children) {
        match item {
            // Issue #154: table-level range markers toggle the same comment
            // stamp the paragraph-level ones do, in document order — a range
            // may open before a row and close after it.
            TableItem::CommentStart(id) => ctx.open_comment(id),
            TableItem::CommentEnd(id) => ctx.close_comment(id),
            TableItem::Row(r) => rows.push(convert_table_row(*r, ctx)),
        }
    }
    Table {
        properties,
        grid,
        rows,
    }
}

/// Walk `<w:tbl>`'s direct children and flatten out every `<w:tr>`,
/// recursing through revision-tracking, custom-XML, and SDT row wrappers.
/// Range markers, proofreading errors, permission ranges, and the
/// `tblPr`/`tblGrid` duplicates produced by `$value` are dropped — they
/// have no rendered effect at table level. Document order is preserved.
///
/// Row-level revisions keep the **final** view in *both* renderer views:
/// insert-side wrappers (`<w:ins>`, `<w:moveTo>`) contribute their rows and
/// delete-side ones (`<w:del>`, `<w:moveFrom>`) do not. Unlike runs — which
/// carry a stamp so the markup view can strike them (issue #154) — a row has
/// no stamp to carry, so a tracked row deletion never renders struck; the
/// row is simply gone. A deliberate scope line, recorded in the README row.
///
/// A row is deleted in either of two spellings, and both are handled here: the
/// table-level wrapper above, and `<w:del>` inside the row's own `<w:trPr>` —
/// which is the one **Word** writes. The `trPr` form is the reason this cannot
/// be a wrapper-only rule: Word wraps the deleted row's runs in `<w:del>` too,
/// so ignoring the row marker left an empty row that belongs to neither view.
fn collect_table_rows(children: Vec<TableChildXml>) -> Vec<TableItem> {
    let mut items = Vec::with_capacity(children.len());
    for child in children {
        match child {
            TableChildXml::Row(r) => items.push(TableItem::Row(r)),
            TableChildXml::Sdt(s) => {
                if let Some(content) = Dup::from(s.content).into_value() {
                    items.extend(collect_table_rows(content.children));
                }
            }
            // Insert side: the rows are in the final document.
            TableChildXml::Ins(rt) | TableChildXml::MoveTo(rt) => {
                items.extend(rt.rows.into_iter().map(|r| TableItem::Row(Box::new(r))));
            }
            // Delete side: they are not.
            TableChildXml::Del(_) | TableChildXml::MoveFrom(_) => {}
            TableChildXml::CustomXml(cx) => {
                items.extend(collect_table_rows(cx.children));
            }
            // Issue #154: markers survive collection, in document order, so
            // `convert_table` can toggle the comment stamp where they stood.
            TableChildXml::CommentRangeStart(c) => {
                items.push(TableItem::CommentStart(CommentId::new(c.id)));
            }
            TableChildXml::CommentRangeEnd(c) => {
                items.push(TableItem::CommentEnd(CommentId::new(c.id)));
            }
            TableChildXml::BookmarkStart(_)
            | TableChildXml::BookmarkEnd(_)
            | TableChildXml::ProofErr(_)
            | TableChildXml::PermStart(_)
            | TableChildXml::PermEnd(_)
            | TableChildXml::TblPr(_)
            | TableChildXml::TblGrid(_)
            | TableChildXml::Other => {}
        }
    }
    // Applied once over everything gathered above, so a row marked deleted in
    // its own `<w:trPr>` goes whether it sat bare in the table, inside an
    // `<w:ins>`, or inside an `<w:sdt>`. Idempotent, so the recursive calls
    // having already run it costs nothing.
    // `last()` is [`Dup`]'s §17.7.2 last-wins rule spelled out on a borrow —
    // the same occurrence `convert_table_row`'s `Dup::from(r.tr_pr)` resolves
    // to, so a document that repeats `<w:trPr>` cannot have the deletion read
    // off one copy and every other property off another.
    items.retain(|item| match item {
        TableItem::Row(r) => !r.tr_pr.last().is_some_and(|pr| pr.marks_row_deleted()),
        _ => true,
    });
    items
}

/// One `<w:tbl>` child that survives collection: a row, or a comment range
/// marker in its document-order position (issue #154).
enum TableItem {
    Row(Box<TableRowXml>),
    CommentStart(CommentId),
    CommentEnd(CommentId),
}

/// The `<w:tr>`-level twin of [`TableItem`].
enum RowItem {
    Cell(Box<TableCellXml>),
    CommentStart(CommentId),
    CommentEnd(CommentId),
}

fn convert_table_row(r: TableRowXml, ctx: &mut ConvertCtx) -> TableRow {
    let rsids = TableRowRevisionIds {
        r: hex_rsid(r.rsid_r.as_deref()),
        r_pr: hex_rsid(r.rsid_r_pr.as_deref()),
        del: hex_rsid(r.rsid_del.as_deref()),
        tr: hex_rsid(r.rsid_tr.as_deref()),
    };
    let properties = Dup::from(r.tr_pr)
        .into_value()
        .map(TableRowProperties::from)
        .unwrap_or_default();
    let property_exceptions = Dup::from(r.tbl_pr_ex).into_value().map(Into::into);
    let mut cells = Vec::new();
    for item in collect_row_cells(r.children) {
        match item {
            RowItem::CommentStart(id) => ctx.open_comment(id),
            RowItem::CommentEnd(id) => ctx.close_comment(id),
            RowItem::Cell(c) => cells.push(convert_table_cell(*c, ctx)),
        }
    }
    TableRow {
        properties,
        cells,
        rsids,
        property_exceptions,
    }
}

/// Walk a `<w:tr>`'s direct children and flatten out every `<w:tc>`,
/// recursing through cell-level SDT (`<w:sdt>`) and custom-XML wrappers.
/// Range markers, proofreading errors, permission ranges, and the
/// `tblPrEx`/`trPr` duplicates produced by `$value` are dropped — they have
/// no rendered effect at cell level. Document order is preserved. Mirrors
/// `collect_table_rows` one level down.
fn collect_row_cells(children: Vec<RowChildXml>) -> Vec<RowItem> {
    let mut items = Vec::with_capacity(children.len());
    for child in children {
        match child {
            RowChildXml::Cell(c) => items.push(RowItem::Cell(c)),
            RowChildXml::Sdt(s) => {
                if let Some(content) = Dup::from(s.content).into_value() {
                    items.extend(collect_row_cells(content.children));
                }
            }
            RowChildXml::CustomXml(cx) => {
                items.extend(collect_row_cells(cx.children));
            }
            RowChildXml::CommentRangeStart(c) => {
                items.push(RowItem::CommentStart(CommentId::new(c.id)));
            }
            RowChildXml::CommentRangeEnd(c) => {
                items.push(RowItem::CommentEnd(CommentId::new(c.id)));
            }
            RowChildXml::BookmarkStart(_)
            | RowChildXml::BookmarkEnd(_)
            | RowChildXml::ProofErr(_)
            | RowChildXml::PermStart(_)
            | RowChildXml::PermEnd(_)
            | RowChildXml::TblPrEx(_)
            | RowChildXml::TrPr(_)
            | RowChildXml::Other => {}
        }
    }
    items
}

fn convert_table_cell(c: TableCellXml, ctx: &mut ConvertCtx) -> TableCell {
    let properties = Dup::from(c.tc_pr)
        .into_value()
        .map(TableCellProperties::from)
        .unwrap_or_default();
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

#[cfg(test)]
mod tests {
    use super::*;

    /// Collect all rendered text from a converted inline sequence, recursing
    /// into hyperlinks and fields.
    fn collect_text(inlines: &[Inline]) -> String {
        let mut s = String::new();
        for inl in inlines {
            match inl {
                Inline::TextRun(r) => {
                    for el in &r.content {
                        if let RunElement::Text(t) = el {
                            s.push_str(t);
                        }
                    }
                }
                Inline::Hyperlink(h) => s.push_str(&collect_text(&h.content)),
                Inline::Field(f) => s.push_str(&collect_text(&f.content)),
                _ => {}
            }
        }
        s
    }

    /// Parse a `<w:p>` and return the converted paragraph.
    fn parse_para(xml: &str) -> Paragraph {
        let p: ParaXml = quick_xml::de::from_str(xml).unwrap();
        let mut ctx = ConvertCtx::new();
        convert_paragraph(p, &mut ctx).0
    }

    /// Parse a `<w:p>` and return the concatenated rendered text.
    fn para_text(xml: &str) -> String {
        let p: ParaXml = quick_xml::de::from_str(xml).unwrap();
        let mut ctx = ConvertCtx::new();
        let (para, _) = convert_paragraph(p, &mut ctx);
        collect_text(&para.content)
    }

    /// Parse a `<w:p>` and return the run-elements of its first `TextRun`.
    fn first_run_elements(xml: &str) -> Vec<RunElement> {
        let p: ParaXml = quick_xml::de::from_str(xml).unwrap();
        let mut ctx = ConvertCtx::new();
        let (para, _) = convert_paragraph(p, &mut ctx);
        para.content
            .into_iter()
            .find_map(|i| match i {
                Inline::TextRun(r) => Some(r.content),
                _ => None,
            })
            .unwrap_or_default()
    }

    // ── Revision & structural wrappers (accept-all-changes / final view) ──

    /// The revision stamp of each `TextRun` in a parsed paragraph, in order,
    /// paired with its text — what the tracked-change tests below assert on.
    fn para_revisions(xml: &str) -> Vec<(String, Option<RunRevision>)> {
        parse_para(xml)
            .content
            .iter()
            .filter_map(|i| match i {
                Inline::TextRun(tr) => Some((
                    tr.content
                        .iter()
                        .filter_map(|e| match e {
                            RunElement::Text(t) => Some(t.as_str()),
                            _ => None,
                        })
                        .collect::<String>(),
                    tr.revision.clone(),
                )),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn ins_content_is_kept_and_stamped_inserted() {
        // Issue #154: the wrapper flattens, the runs carry the stamp — and
        // the author, which revision marks are colored by.
        let runs = para_revisions(
            r#"<w:p xmlns:w="x"><w:ins w:author="Ann"><w:r><w:t>kept</w:t></w:r></w:ins></w:p>"#,
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].0, "kept");
        assert_eq!(
            runs[0].1,
            Some(RunRevision {
                kind: RevisionKind::Inserted,
                author: "Ann".into(),
            })
        );
    }

    #[test]
    fn del_content_is_kept_and_stamped_deleted() {
        // Issue #154: an unaccepted deletion is document content — the
        // renderer decides whether it paints struck through or not at all,
        // so the parse must hand it over rather than pre-empt the decision.
        let runs = para_revisions(
            r#"<w:p xmlns:w="x"><w:del w:author="Ann"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>"#,
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].0, "gone");
        assert_eq!(
            runs[0].1.as_ref().map(|r| r.kind),
            Some(RevisionKind::Deleted)
        );
    }

    #[test]
    fn del_nested_in_ins_carries_the_deletion() {
        // Nested wrappers shadow: text inserted then deleted is a deletion;
        // the runs either side keep the insertion stamp.
        let runs = para_revisions(
            r#"<w:p xmlns:w="x"><w:ins w:author="Ann">
                 <w:r><w:t>A</w:t></w:r>
                 <w:del w:author="Ann"><w:r><w:delText>B</w:delText></w:r></w:del>
                 <w:r><w:t>C</w:t></w:r>
               </w:ins></w:p>"#,
        );
        let kinds: Vec<_> = runs
            .iter()
            .map(|(t, r)| (t.as_str(), r.as_ref().map(|r| r.kind)))
            .collect();
        assert_eq!(
            kinds,
            vec![
                ("A", Some(RevisionKind::Inserted)),
                ("B", Some(RevisionKind::Deleted)),
                ("C", Some(RevisionKind::Inserted)),
            ]
        );
    }

    /// §17.16.13: `delInstrText` is how Word spells the field code of a
    /// complex field deleted while tracking changes — it must parse (it used
    /// to fail the whole document), and the deleted field machinery must not
    /// survive into the model, where the final view would re-evaluate it.
    #[test]
    fn del_inst_text_parses_and_the_deleted_field_is_dropped() {
        let para = parse_para(
            r#"<w:p xmlns:w="x"><w:del w:author="A">
                 <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                 <w:r><w:delInstrText> PAGE </w:delInstrText></w:r>
                 <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                 <w:r><w:delText>7</w:delText></w:r>
                 <w:r><w:fldChar w:fldCharType="end"/></w:r>
               </w:del></w:p>"#,
        );
        assert!(
            !para
                .content
                .iter()
                .any(|i| matches!(i, Inline::FieldChar(_) | Inline::InstrText(_))),
            "no piece of the deleted field machinery may survive"
        );
    }

    /// Everything visible a run cannot stamp is dropped from a deletion in
    /// every view — a symbol, a note reference, a whole field, a drawing
    /// wrapped in `mc:AlternateContent` (Word's spelling of every wps shape).
    #[test]
    fn deleted_unstampables_are_dropped() {
        let para = parse_para(
            r#"<w:p xmlns:w="x" xmlns:mc="y"><w:del w:author="A">
                 <w:r><w:delText>kept-as-stamped</w:delText></w:r>
                 <w:r><w:sym w:font="F" w:char="F0FC"/></w:r>
                 <w:r><w:footnoteReference w:id="3"/></w:r>
                 <w:fldSimple w:instr=" DATE "><w:r><w:delText>CACHED</w:delText></w:r></w:fldSimple>
                 <w:r><mc:AlternateContent><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
               </w:del></w:p>"#,
        );
        for inline in &para.content {
            assert!(
                matches!(inline, Inline::TextRun(_)),
                "only stamped runs survive, found {inline:?}"
            );
        }
        assert_eq!(collect_text(&para.content), "kept-as-stamped");
    }

    /// §17.13.5.14: deletion is sticky — an `<w:ins>` inside a `<w:del>` is
    /// still deleted text, whoever inserted it first.
    #[test]
    fn ins_nested_in_del_stays_deleted() {
        let runs = para_revisions(
            r#"<w:p xmlns:w="x"><w:del w:author="B">
                 <w:ins w:author="A"><w:r><w:delText>gone</w:delText></w:r></w:ins>
               </w:del></w:p>"#,
        );
        assert_eq!(runs.len(), 1);
        assert_eq!(
            runs[0].1.as_ref().map(|r| r.kind),
            Some(RevisionKind::Deleted),
            "the deletion must win over the inner insertion"
        );
    }

    /// Issue #154: `commentReference` terminates its own range — Word writes
    /// the end marker first, so a document that lost it would otherwise wash
    /// everything to the end of the part.
    #[test]
    fn a_comment_reference_closes_an_unclosed_range() {
        let para = parse_para(
            r#"<w:p xmlns:w="x"><w:commentRangeStart w:id="9"/><w:r><w:t>in</w:t></w:r><w:r><w:commentReference w:id="9"/></w:r><w:r><w:t>out</w:t></w:r></w:p>"#,
        );
        let stamps: Vec<_> = para
            .content
            .iter()
            .filter_map(|i| match i {
                Inline::TextRun(tr) => Some((collect_text(std::slice::from_ref(i)), tr.comment)),
                _ => None,
            })
            .collect();
        assert_eq!(stamps[0].1.map(|c| c.value()), Some(9), "inside the range");
        assert_eq!(stamps[1].1, None, "after the reference the range is closed");
    }

    /// §17.13.5.15: the paragraph mark's own deletion is modelled — it is how
    /// Word deletes a whole paragraph, and the final view merges it away.
    #[test]
    fn a_deleted_paragraph_mark_is_parsed() {
        let para = parse_para(
            r#"<w:p xmlns:w="x"><w:pPr><w:rPr><w:del w:author="A"/></w:rPr></w:pPr><w:del w:author="A"><w:r><w:delText>x</w:delText></w:r></w:del></w:p>"#,
        );
        assert!(para.mark_deleted);
        let control = parse_para(r#"<w:p xmlns:w="x"><w:r><w:t>y</w:t></w:r></w:p>"#);
        assert!(!control.mark_deleted);
    }

    /// A deleted picture cannot carry the run stamp, so it stays out of the
    /// model in every view — exactly the pre-#154 behavior for all deleted
    /// content, now scoped to what a run cannot express.
    #[test]
    fn a_deleted_image_is_still_dropped() {
        let para = parse_para(
            r#"<w:p xmlns:w="x"><w:del w:author="Ann"><w:r><w:delText>t</w:delText></w:r><w:r><w:pict/></w:r></w:del></w:p>"#,
        );
        assert!(
            !para.content.iter().any(|i| matches!(i, Inline::Pict(_))),
            "no picture survives a deletion"
        );
    }

    #[test]
    fn move_to_rendered_move_from_dropped() {
        assert_eq!(
            para_text(r#"<w:p xmlns:w="x"><w:moveTo><w:r><w:t>here</w:t></w:r></w:moveTo></w:p>"#),
            "here"
        );
        assert_eq!(
            para_text(
                r#"<w:p xmlns:w="x"><w:moveFrom><w:r><w:t>away</w:t></w:r></w:moveFrom></w:p>"#
            ),
            ""
        );
    }

    #[test]
    fn smart_tag_and_inline_custom_xml_are_flattened() {
        assert_eq!(
            para_text(
                r#"<w:p xmlns:w="x"><w:smartTag><w:r><w:t>date</w:t></w:r></w:smartTag></w:p>"#
            ),
            "date"
        );
        assert_eq!(
            para_text(
                r#"<w:p xmlns:w="x"><w:customXml><w:r><w:t>tagged</w:t></w:r></w:customXml></w:p>"#
            ),
            "tagged"
        );
    }

    #[test]
    fn ins_wrapping_a_hyperlink_keeps_the_link_text() {
        assert_eq!(
            para_text(
                r#"<w:p xmlns:w="x"><w:ins>
                     <w:hyperlink w:anchor="a"><w:r><w:t>link</w:t></w:r></w:hyperlink>
                   </w:ins></w:p>"#
            ),
            "link"
        );
    }

    /// Issue #154: comment range markers at table level toggle the same
    /// stamp the paragraph-level ones do, in document order — a range that
    /// opens before a row and closes after it stamps the row's runs, and
    /// text after the close marker is clean.
    #[test]
    fn table_level_comment_markers_toggle_the_stamp() {
        let t: TableXml = quick_xml::de::from_str(
            r#"<w:tbl xmlns:w="x">
                 <w:commentRangeStart w:id="5"/>
                 <w:tr><w:tc><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:tc></w:tr>
                 <w:commentRangeEnd w:id="5"/>
                 <w:tr><w:tc><w:p><w:r><w:t>outside</w:t></w:r></w:p></w:tc></w:tr>
               </w:tbl>"#,
        )
        .unwrap();
        let mut ctx = ConvertCtx::new();
        let table = convert_table(t, &mut ctx);
        let stamp_of = |row: &crate::docx::model::TableRow| {
            row.cells
                .iter()
                .flat_map(|c| c.content.iter())
                .find_map(|b| match b {
                    Block::Paragraph(p) => p.content.iter().find_map(|i| match i {
                        Inline::TextRun(tr) => Some(tr.comment),
                        _ => None,
                    }),
                    _ => None,
                })
                .flatten()
        };
        assert_eq!(stamp_of(&table.rows[0]).map(|c| c.value()), Some(5));
        assert_eq!(stamp_of(&table.rows[1]), None);
    }

    // ── Revision-tracked *rows* (same final view as revision-tracked runs) ──

    /// The concatenated text of each surviving row of a `<w:tbl>`, in document
    /// order. Reads the converted model rather than the schema, so it measures
    /// what a renderer would be handed.
    fn table_row_texts(xml: &str) -> Vec<String> {
        let t: TableXml = quick_xml::de::from_str(xml).unwrap();
        let mut ctx = ConvertCtx::new();
        convert_table(t, &mut ctx)
            .rows
            .into_iter()
            .map(|r| {
                r.cells
                    .into_iter()
                    .flat_map(|c| c.content)
                    .filter_map(|b| match b {
                        Block::Paragraph(p) => Some(collect_text(&p.content)),
                        _ => None,
                    })
                    .collect::<String>()
            })
            .collect()
    }

    /// A `<w:tbl>` of three rows, the middle one wrapped in `wrapper`.
    fn table_with_wrapped_row(wrapper: &str) -> String {
        format!(
            r#"<w:tbl xmlns:w="x">
                 <w:tblPr/>
                 <w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>
                 <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
                 <w:{wrapper} w:id="1" w:author="a" w:date="d">
                   <w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
                 </w:{wrapper}>
                 <w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>
               </w:tbl>"#
        )
    }

    /// A `<w:tbl>` of three rows, the middle one carrying `marker` in its
    /// `<w:trPr>` — the form Word writes when a row is deleted or inserted with
    /// change tracking on.
    fn table_with_marked_row(marker: &str) -> String {
        format!(
            r#"<w:tbl xmlns:w="x">
                 <w:tblPr/>
                 <w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>
                 <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
                 <w:tr>
                   <w:trPr>{marker}</w:trPr>
                   <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                 </w:tr>
                 <w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr>
               </w:tbl>"#
        )
    }

    /// A table-level `<w:del>` is the delete side of a tracked change, exactly
    /// as `<w:del>` around a run is — and this parser renders the *final* view,
    /// so both go. Keeping deleted rows while dropping deleted runs left the
    /// engine rendering the final view for text and the original view for
    /// rows, in the same document.
    #[test]
    fn del_wrapped_rows_are_dropped() {
        assert_eq!(table_row_texts(&table_with_wrapped_row("del")), ["A", "C"]);
    }

    /// The move-away side goes with it (`move_to_rendered_move_from_dropped`
    /// is the run-level twin).
    #[test]
    fn move_from_wrapped_rows_are_dropped() {
        assert_eq!(
            table_row_texts(&table_with_wrapped_row("moveFrom")),
            ["A", "C"]
        );
    }

    /// The insert side is part of the final document and must survive — the
    /// control that a "drop tracked rows" fix would break.
    #[test]
    fn ins_and_move_to_wrapped_rows_are_kept() {
        assert_eq!(
            table_row_texts(&table_with_wrapped_row("ins")),
            ["A", "B", "C"]
        );
        assert_eq!(
            table_row_texts(&table_with_wrapped_row("moveTo")),
            ["A", "B", "C"]
        );
    }

    /// `<w:del>` inside `<w:trPr>` is how **Word** marks a deleted row — the
    /// table-level wrapper above is the other spelling, and the one Word does
    /// not use. Unmodelled, it left the row in place; its runs are separately
    /// wrapped in `<w:del>` and already dropped, so the row rendered as an
    /// empty ghost belonging to neither view of the document.
    #[test]
    fn a_row_deleted_in_its_tr_pr_is_dropped() {
        assert_eq!(
            table_row_texts(&table_with_marked_row(
                r#"<w:del w:id="1" w:author="a" w:date="d"/>"#
            )),
            ["A", "C"]
        );
    }

    /// `<w:ins>` in the same position marks an *inserted* row, which the final
    /// view keeps. Pins that the two markers are told apart rather than the
    /// row being dropped for carrying any revision mark at all.
    #[test]
    fn a_row_inserted_in_its_tr_pr_is_kept() {
        assert_eq!(
            table_row_texts(&table_with_marked_row(
                r#"<w:ins w:id="1" w:author="a" w:date="d"/>"#
            )),
            ["A", "B", "C"]
        );
    }

    /// A row whose `<w:trPr>` carries ordinary properties is untouched — the
    /// new field must not turn every `trPr` into a deletion.
    #[test]
    fn a_row_with_an_unrelated_tr_pr_is_kept() {
        assert_eq!(
            table_row_texts(&table_with_marked_row(r#"<w:cantSplit/>"#)),
            ["A", "B", "C"]
        );
    }

    // ── run_break ────────────────────────────────────────────────────────

    #[test]
    fn br_type_page_and_column() {
        assert!(matches!(
            first_run_elements(r#"<w:p xmlns:w="x"><w:r><w:br type="page"/></w:r></w:p>"#)
                .as_slice(),
            [RunElement::PageBreak]
        ));
        assert!(matches!(
            first_run_elements(r#"<w:p xmlns:w="x"><w:r><w:br type="column"/></w:r></w:p>"#)
                .as_slice(),
            [RunElement::ColumnBreak]
        ));
    }

    #[test]
    fn br_plain_is_text_wrapping_line_break() {
        assert!(matches!(
            first_run_elements(r#"<w:p xmlns:w="x"><w:r><w:br/></w:r></w:p>"#).as_slice(),
            [RunElement::LineBreak(BreakKind::TextWrapping)]
        ));
    }

    #[test]
    fn br_clear_all_is_clearing_line_break() {
        assert!(matches!(
            first_run_elements(r#"<w:p xmlns:w="x"><w:r><w:br clear="all"/></w:r></w:p>"#)
                .as_slice(),
            [RunElement::LineBreak(BreakKind::Clear(BreakClear::All))]
        ));
    }

    // ── mc:Choice Requires ───────────────────────────────────────────────

    #[test]
    fn mc_requires_known_and_unknown() {
        assert_eq!(mc_requires("wps"), Some(McRequires::Wps));
        assert_eq!(mc_requires("w14"), Some(McRequires::W14));
        // Unknown / unsupported token → None (choice is dropped, fallback used).
        assert_eq!(mc_requires("nope"), None);
        // `mc_requires` maps a *single* token; a space-separated list is split by
        // `convert_alt_content` (see `alt_content_requires_*` below), so a raw
        // multi-token string is not a valid single token.
        assert_eq!(mc_requires("wps w14"), None);
    }

    #[test]
    fn alt_content_multi_token_requires_is_kept_when_all_known() {
        // §M.2.2: a space-separated Requires with every token understood keeps
        // the choice (previously the whole choice was dropped).
        let xml = r#"<w:r xmlns:w="x" xmlns:mc="m">
              <mc:AlternateContent>
                <mc:Choice Requires="wps w14"><w:drawing/></mc:Choice>
                <mc:Fallback/>
              </mc:AlternateContent>
            </w:r>"#;
        let r: RunXml = quick_xml::de::from_str(xml).unwrap();
        let mut out = Vec::new();
        let mut ctx = ConvertCtx::new();
        extend_from_run(r, &mut out, &mut ctx);
        let Some(Inline::AlternateContent(ac)) = out.into_iter().next() else {
            panic!("expected AlternateContent");
        };
        assert_eq!(ac.choices.len(), 1, "multi-token choice kept");
        assert_eq!(
            ac.choices[0].requires,
            vec![McRequires::Wps, McRequires::W14]
        );
    }

    #[test]
    fn alt_content_choice_with_unknown_token_is_dropped() {
        let xml = r#"<w:r xmlns:w="x" xmlns:mc="m">
              <mc:AlternateContent>
                <mc:Choice Requires="wps nope"><w:drawing/></mc:Choice>
                <mc:Fallback/>
              </mc:AlternateContent>
            </w:r>"#;
        let r: RunXml = quick_xml::de::from_str(xml).unwrap();
        let mut out = Vec::new();
        let mut ctx = ConvertCtx::new();
        extend_from_run(r, &mut out, &mut ctx);
        let Some(Inline::AlternateContent(ac)) = out.into_iter().next() else {
            panic!("expected AlternateContent");
        };
        assert!(
            ac.choices.is_empty(),
            "choice with an unknown token dropped"
        );
    }

    /// Regression: a `<w:tr>` whose `<w:tc>` cells are interleaved with
    /// cell-level `<w:sdt>` (CT_SdtCell) wrappers must (a) parse — the old
    /// `Vec<TableCellXml>` field reported a duplicate `tc` when cells were
    /// non-contiguous — and (b) recover every cell, including the ones nested
    /// inside `<w:sdtContent>`.
    #[test]
    fn row_with_interspersed_sdt_cells_recovers_all_cells() {
        let xml = r#"
            <w:tr xmlns:w="x">
              <w:trPr/>
              <w:tc><w:p/></w:tc>
              <w:tc><w:p/></w:tc>
              <w:sdt>
                <w:sdtPr/>
                <w:sdtContent><w:tc><w:p/></w:tc></w:sdtContent>
              </w:sdt>
              <w:sdt>
                <w:sdtContent><w:tc><w:p/></w:tc></w:sdtContent>
              </w:sdt>
              <w:tc><w:p/></w:tc>
            </w:tr>
        "#;
        let row: TableRowXml = quick_xml::de::from_str(xml).expect("row must parse");
        assert!(!row.tr_pr.is_empty(), "trPr still captured on the parent");
        let cells = collect_row_cells(row.children);
        assert_eq!(cells.len(), 5, "3 bare + 2 sdt-wrapped cells recovered");
    }

    /// Nested `<w:customXml>` cell wrapper (CT_CustomXmlCell) also flattens.
    #[test]
    fn row_with_custom_xml_cell_wrapper_recovers_cells() {
        let xml = r#"
            <w:tr xmlns:w="x">
              <w:tc><w:p/></w:tc>
              <w:customXml>
                <w:tc><w:p/></w:tc>
                <w:tc><w:p/></w:tc>
              </w:customXml>
            </w:tr>
        "#;
        let row: TableRowXml = quick_xml::de::from_str(xml).expect("row must parse");
        assert_eq!(collect_row_cells(row.children).len(), 3);
    }

    /// A `<w:customXml>` row wrapper (CT_CustomXmlRow) inside `<w:tbl>` must
    /// have its nested `<w:tr>` rows recovered — the element is `customXml`,
    /// not `customXmlIns`/etc., so the wrong rename previously dropped them.
    #[test]
    fn table_with_custom_xml_row_wrapper_recovers_rows() {
        let xml = r#"
            <w:tbl xmlns:w="x">
              <w:tblPr/>
              <w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>
              <w:tr><w:tc><w:p/></w:tc></w:tr>
              <w:customXml>
                <w:tr><w:tc><w:p/></w:tc></w:tr>
              </w:customXml>
            </w:tbl>
        "#;
        let tbl: TableXml = quick_xml::de::from_str(xml).expect("table must parse");
        assert_eq!(collect_table_rows(tbl.children).len(), 2);
    }
}
