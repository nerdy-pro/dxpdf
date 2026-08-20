//! §17.6.22 `SectionType::Continuous` — a section break that starts the next
//! section on the page the previous one is still filling.
//!
//! Issue #83. The engine had no coverage for continuous breaks at all before
//! this file, so the cases here are written from the spec clause rather than
//! from the current behaviour.

use std::io::Write;

/// Minimal single-part DOCX around `body`.
fn docx(body: &str) -> Vec<u8> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/_rels/document.xml.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
  <Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdF2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
  <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"#).unwrap();

    // A single decimal list, so §17.9 label counters are observable in output.
    zip.start_file("word/numbering.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"#,
    )
    .unwrap();

    // A header, so sections can differ in header height for the clearance
    // cases. Header content is built by `build_non_story_content`, which calls
    // `build_fragments` directly and never `inject_list_label` — so measuring a
    // header does *not* touch §17.9 list counters. The document-order side
    // effect it does have (§17.11.12 footnote display numbers) is guarded by a
    // unit test in `render::tests`, where the peek itself is reachable.
    zip.start_file("word/header1.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>HDRSHORT</w:t></w:r></w:p>
</w:hdr>"#,
    )
    .unwrap();

    // §17.10.1: a deliberately *taller* header — four 20pt lines. Reserving it
    // pushes the body top far below where the short header would.
    zip.start_file("word/header2.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>HDRTALL1</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>HDRTALL2</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>HDRTALL3</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>HDRTALL4</w:t></w:r></w:p>
</w:hdr>"#,
    )
    .unwrap();

    // §17.10: the footer pair mirrors the header pair — one short slot, one
    // four-line 20pt slot. A footer raises the page *bottom* rather than
    // lowering the top, so the two are only symmetric if `for_page` treats them
    // so; that is the claim these fixtures exist to test rather than assume.
    zip.start_file("word/footer1.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>FTRSHORT</w:t></w:r></w:p>
</w:ftr>"#,
    )
    .unwrap();

    zip.start_file("word/footer2.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>FTRTALL1</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>FTRTALL2</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>FTRTALL3</w:t></w:r></w:p>
  <w:p><w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>FTRTALL4</w:t></w:r></w:p>
</w:ftr>"#,
    )
    .unwrap();

    // §17.11.23: two notes, one referenced on each side of the break.
    zip.start_file("word/footnotes.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>NOTEBEFORE</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="3"><w:p><w:r><w:t>NOTEAFTER</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="4"><w:p><w:r><w:t>NOTEMIDDLE</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>{body}</w:body>
</w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

const PG: &str = r#"<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567"/>"#;

/// A numbered list item — its rendered label is the §17.9 counter, so the
/// drawn text is a direct read-out of numbering state.
fn list_item(text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

/// Every text string drawn across every page, in page then emission order.
fn drawn_text(pages: &[dxpdf::render::layout::draw_command::LayoutedPage]) -> Vec<String> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, .. } => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

fn layout(body: &str) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    let doc = dxpdf::docx::parse(&docx(body)).expect("parse");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);
    pages
}

/// §17.9: list numbering runs in **document order**, so splitting a document
/// into two sections with a continuous break must not renumber it.
///
/// This guards the *relayout*, not the clearance peek: re-running a section's
/// layout must not re-run its block building, which is where list counters
/// advance (`inject_list_label`). Written first, while the relayout does not
/// exist yet, so it is green until that work lands and red the moment a
/// re-layout starts double-counting.
#[test]
fn continuous_break_does_not_renumber_lists() {
    let items = || format!("{}{}", list_item("Alpha"), list_item("Beta"));

    // One section, no break — the reference numbering.
    let single = layout(&format!(
        "{}{}<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdH1\" \
         xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>{PG}</w:sectPr>",
        items(),
        list_item("Gamma"),
    ));

    // Same content, split by a continuous break.
    let split = layout(&format!(
        "<w:p><w:pPr><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdH1\" \
         xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>{PG}</w:sectPr></w:pPr></w:p>\
         {}{}<w:sectPr><w:type w:val=\"continuous\"/>{PG}</w:sectPr>",
        items(),
        list_item("Gamma"),
    ));

    let labels = |pages: &[dxpdf::render::layout::draw_command::LayoutedPage]| {
        drawn_text(pages)
            .into_iter()
            .filter(|t| {
                t.trim_end().ends_with('.') && t.trim_end_matches('.').parse::<u32>().is_ok()
            })
            .collect::<Vec<_>>()
    };

    let a = labels(&single);
    let b = labels(&split);
    assert!(
        !a.is_empty(),
        "fixture must actually render numbered labels; got {:?}",
        drawn_text(&single)
    );
    assert_eq!(
        a, b,
        "a continuous section break must not change §17.9 list numbering — \
         block building ran more than once for the same blocks"
    );
}

/// §17.6.22: content order across the break is document order. The break moves
/// the *page* boundary, never the sequence.
#[test]
fn continuous_break_preserves_content_order() {
    let pages = layout(&format!(
        "<w:p><w:r><w:t>Before</w:t></w:r></w:p>\
         <w:p><w:pPr><w:sectPr>{PG}</w:sectPr></w:pPr></w:p>\
         <w:p><w:r><w:t>After</w:t></w:r></w:p>\
         <w:sectPr><w:type w:val=\"continuous\"/>{PG}</w:sectPr>"
    ));

    let texts = drawn_text(&pages);
    let before = texts.iter().position(|t| t.contains("Before"));
    let after = texts.iter().position(|t| t.contains("After"));
    assert!(
        before.is_some() && after.is_some(),
        "both sections must render: {texts:?}"
    );
    assert!(
        before < after,
        "the section after a continuous break follows it in document order: {texts:?}"
    );
}

/// §17.6.22: "continuous" means the next section shares the current page. It
/// must not silently become a page break.
#[test]
fn continuous_break_shares_one_physical_page() {
    let pages = layout(&format!(
        "<w:p><w:r><w:t>Before</w:t></w:r></w:p>\
         <w:p><w:pPr><w:sectPr>{PG}</w:sectPr></w:pPr></w:p>\
         <w:p><w:r><w:t>After</w:t></w:r></w:p>\
         <w:sectPr><w:type w:val=\"continuous\"/>{PG}</w:sectPr>"
    ));

    assert_eq!(
        pages.len(),
        1,
        "two short sections joined by a continuous break occupy one page"
    );
}

// ── §17.10 header/footer clearance on a shared page (issue #83, plan §2) ─────

const R_NS: &str =
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A `sectPr` referencing header `rel`, optionally as a continuous break.
fn sect_pr(rel: &str, continuous: bool) -> String {
    let ty = if continuous {
        r#"<w:type w:val="continuous"/>"#
    } else {
        ""
    };
    format!(
        r#"<w:sectPr>{ty}<w:headerReference w:type="default" r:id="{rel}" {R_NS}/>{PG}</w:sectPr>"#
    )
}

fn para(text: &str) -> String {
    format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>")
}

/// The y of the first draw command whose text is exactly `needle`.
fn y_of(pages: &[dxpdf::render::layout::draw_command::LayoutedPage], needle: &str) -> f32 {
    use dxpdf::render::layout::draw_command::DrawCommand;
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.contains(needle) => {
                Some(position.y.raw())
            }
            _ => None,
        })
        .unwrap_or_else(|| panic!("{needle:?} was never drawn"))
}

/// Two sections sharing one page: the first uses `first_hdr`, the second
/// continues on the same page under `second_hdr`.
fn shared_page(
    first_hdr: &str,
    second_hdr: &str,
) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr(first_hdr, false),
        para("BODYAFTER"),
        sect_pr(second_hdr, true),
    ))
}

/// One section, so `hdr` governs the whole page — the reference layout the
/// shared page must reproduce.
///
/// The empty paragraph is not padding: in [`shared_page`] the section break
/// rides on a `w:p`'s `w:pPr`, and that paragraph mark is real content that
/// renders as a blank line. Without it here the two documents differ by a line
/// and the comparison measures the fixture rather than the engine.
fn single_section(hdr: &str) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    layout(&format!(
        "{}<w:p/>{}{}",
        para("BODYBEFORE"),
        para("BODYAFTER"),
        sect_pr(hdr, false),
    ))
}

/// §17.6 / §17.10: ECMA-376 does not say which section owns the header of a
/// page holding a continuous break. This pins the engine's answer — the **last
/// section on the page** — so the clearance rule below has a stated premise
/// rather than an implicit one.
#[test]
fn shared_page_renders_the_succeeding_sections_header() {
    let pages = shared_page("rIdH1", "rIdH2");
    assert_eq!(pages.len(), 1, "the sections share one physical page");

    let texts = drawn_text(&pages);
    assert!(
        texts.iter().any(|t| t.contains("HDRTALL")),
        "the succeeding section's header is the one drawn: {texts:?}"
    );
    assert!(
        !texts.iter().any(|t| t.contains("HDRSHORT")),
        "the preceding section's header must not also be drawn: {texts:?}"
    );
}

/// §17.10.1: the shared page is drawn with the succeeding section's header, so
/// **all** body content on it — including the preceding section's — must clear
/// that header. Laying the preceding section out against its own shorter
/// header leaves its content underneath the taller one that is actually drawn.
#[test]
fn taller_succeeding_header_does_not_overlap_preceding_content() {
    let shared = shared_page("rIdH1", "rIdH2");
    let reference = single_section("rIdH2");

    assert_eq!(
        y_of(&shared, "BODYBEFORE"),
        y_of(&reference, "BODYBEFORE"),
        "content preceding the break must sit where the tall header puts it, \
         not where the short header it was laid out against did"
    );
}

/// The converse, and the reason this is not simply "take the larger of the
/// two": a *shorter* succeeding header must reclaim the band the preceding
/// section's taller one reserved, rather than leaving it blank.
#[test]
fn shorter_succeeding_header_reclaims_the_band() {
    let shared = shared_page("rIdH2", "rIdH1");
    let reference = single_section("rIdH1");

    assert_eq!(
        y_of(&shared, "BODYBEFORE"),
        y_of(&reference, "BODYBEFORE"),
        "a shorter succeeding header must not retain the preceding section's \
         clearance"
    );
}

/// Equal slots are the no-op case: the relayout must not perturb a document
/// whose sections agree.
#[test]
fn equal_headers_leave_the_shared_page_unchanged() {
    let shared = shared_page("rIdH1", "rIdH1");
    let reference = single_section("rIdH1");

    assert_eq!(
        y_of(&shared, "BODYBEFORE"),
        y_of(&reference, "BODYBEFORE"),
        "the shared page's own bounds are unchanged"
    );
    assert_eq!(
        y_of(&shared, "BODYAFTER"),
        y_of(&reference, "BODYAFTER"),
        "§17.6.22: a continuous break inserts no vertical space — content after \
         it flows exactly where it would have without the break. This fails on \
         the *continuation cursor*, not the bounds: reconstructing it from draw \
         commands as `baseline + font_size` overshoots the true line bottom."
    );
}

/// §17.6.11: the resolved body bounds are a *floor* on the configured margin,
/// never a reduction of it. A header shorter than `w:pgMar/@top` leaves the
/// margin standing.
#[test]
fn shared_page_bounds_never_reduce_the_configured_margin() {
    // 1134tw = 56.7pt top margin; HDRSHORT is one 10pt line from 567tw
    // (28.35pt), so it ends well above the margin.
    const TOP_MARGIN_PT: f32 = 56.7;
    let shared = shared_page("rIdH1", "rIdH1");
    assert!(
        y_of(&shared, "BODYBEFORE") >= TOP_MARGIN_PT,
        "body must start at or below the configured top margin, got {}",
        y_of(&shared, "BODYBEFORE")
    );
}

// ── §17.6.22 continuation fidelity (issue #83, plan §3) ─────────────────────

/// Content width for [`PG`]: 11906 − 1134 − 1134 twips.
const CONTENT_WIDTH_PT: f32 = (11906.0 - 1134.0 - 1134.0) / 20.0;

/// §17.11.23: the separator is drawn at one third of the content width, which
/// is distinctive enough to count in a fixture with no borders or tables.
fn footnote_separator_count(pages: &[dxpdf::render::layout::draw_command::LayoutedPage]) -> usize {
    use dxpdf::render::layout::draw_command::DrawCommand;
    let want = CONTENT_WIDTH_PT * 0.33;
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter(|c| match c {
            DrawCommand::Line { line, .. } => {
                ((line.end.x - line.start.x).raw() - want).abs() < 0.5
            }
            _ => false,
        })
        .count()
}

/// The y of the footnote separator on `page`, if one was drawn.
fn footnote_separator_y(page: &dxpdf::render::layout::draw_command::LayoutedPage) -> Option<f32> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    let want = CONTENT_WIDTH_PT * 0.33;
    page.commands.iter().find_map(|c| match c {
        DrawCommand::Line { line, .. }
            if ((line.end.x - line.start.x).raw() - want).abs() < 0.5 =>
        {
            Some(line.start.y.raw())
        }
        _ => None,
    })
}

/// A paragraph carrying a footnote reference.
fn para_with_note(text: &str, id: u32) -> String {
    format!(
        r#"<w:p><w:r><w:t>{text}</w:t></w:r><w:r><w:footnoteReference w:id="{id}"/></w:r></w:p>"#
    )
}

/// §17.11.23: a page carries **one** separator, above all the notes on it. Two
/// sections contributing notes to the same physical page is still one page.
///
/// Each section flushing its own notes draws a second separator, and positions
/// it from the page's unreduced bottom — so the second block is laid over the
/// first rather than above it.
#[test]
fn footnotes_from_both_sections_share_one_separator() {
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para_with_note("BODYBEFORE", 2),
        sect_pr("rIdH1", false),
        para_with_note("BODYAFTER", 3),
        sect_pr("rIdH1", true),
    ));

    assert_eq!(pages.len(), 1, "one shared physical page");
    let texts = drawn_text(&pages);
    assert!(
        texts.iter().any(|t| t.contains("NOTEBEFORE")),
        "the note referenced before the break must render: {texts:?}"
    );
    assert!(
        texts.iter().any(|t| t.contains("NOTEAFTER")),
        "the note referenced after the break must render: {texts:?}"
    );
    assert_eq!(
        footnote_separator_count(&pages),
        1,
        "§17.11.23: one separator per page, not one per section"
    );
}

/// §17.11.23: notes reserve space at the page bottom, and content after the
/// break must respect that reservation — the succeeding section inherits a
/// bottom already reduced by what the preceding one reserved, not the page's
/// unreduced bottom.
#[test]
fn content_after_the_break_clears_footnotes_reserved_before_it() {
    // The succeeding section must fill past the page bottom for the
    // reservation to bite: a couple of short paragraphs never reach it, and the
    // test would pass without anything being carried.
    let filler: String = (0..120).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para_with_note("BODYBEFORE", 2),
        sect_pr("rIdH1", false),
        filler,
        sect_pr("rIdH1", true),
    ));

    // The separator is the top of the reserved block — comparing against the
    // note's own baseline instead would call a line drawn *between* separator
    // and note a pass.
    let separator_y = footnote_separator_y(&pages[0]).expect("a separator on the shared page");
    let last_body_y = {
        use dxpdf::render::layout::draw_command::DrawCommand;
        pages[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, position, .. } if text.starts_with("FILL") => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .fold(f32::MIN, f32::max)
    };
    assert!(
        last_body_y > f32::MIN,
        "the succeeding section must actually reach the shared page"
    );
    assert!(
        last_body_y < separator_y,
        "body after the break (lowest line {last_body_y}) must stay above the \
         footnote separator ({separator_y}) reserved before it"
    );
}

/// §17.3.1.9: `space_before` and `space_after` collapse between adjacent
/// paragraphs. A continuous break is not a paragraph boundary of its own, so
/// the collapse must survive it — the succeeding section must know what the
/// preceding one's last paragraph left behind.
#[test]
fn paragraph_spacing_collapses_across_the_break() {
    let spaced = |text: &str, before: u32, after: u32| {
        format!(
            r#"<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };
    // 400tw = 20pt after, 400tw = 20pt before: collapsed to 20pt, not 40pt.
    //
    // The break rides on BODYBEFORE's own `w:pPr` rather than a following empty
    // paragraph: an empty paragraph is a real paragraph mark that renders a
    // blank line *and* takes part in the collapse, so it would make the two
    // documents incomparable.
    let split = layout(&format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="400"/>{}</w:pPr><w:r><w:t>BODYBEFORE</w:t></w:r></w:p>{}"#,
        sect_pr("rIdH1", false),
        spaced("BODYAFTER", 400, 0) + &sect_pr("rIdH1", true),
    ));
    let joined = layout(&format!(
        "{}{}{}",
        spaced("BODYBEFORE", 0, 400),
        spaced("BODYAFTER", 400, 0),
        sect_pr("rIdH1", false),
    ));

    let gap = |p: &[dxpdf::render::layout::draw_command::LayoutedPage]| {
        y_of(p, "BODYAFTER") - y_of(p, "BODYBEFORE")
    };
    assert!(
        (gap(&split) - gap(&joined)).abs() < 0.01,
        "collapse must survive the break: split gap {}, unbroken gap {}",
        gap(&split),
        gap(&joined),
    );
}

/// §17.3.1.33: `space_before` is suppressed at the top of a page. The first
/// paragraph *after* a continuous break is mid-page, so its `space_before`
/// applies — treating it as a page top swallows the space.
#[test]
fn space_before_is_not_suppressed_after_a_mid_page_break() {
    let with_space = |before: u32| {
        let after_para = format!(
            r#"<w:p><w:pPr><w:spacing w:before="{before}"/></w:pPr><w:r><w:t>BODYAFTER</w:t></w:r></w:p>"#
        );
        format!(
            "{}<w:p><w:pPr>{}</w:pPr></w:p>{after_para}{}",
            para("BODYBEFORE"),
            sect_pr("rIdH1", false),
            sect_pr("rIdH1", true),
        )
    };
    let none = layout(&with_space(0));
    let spaced = layout(&with_space(600)); // 30pt

    let gap = |p: &[dxpdf::render::layout::draw_command::LayoutedPage]| {
        y_of(p, "BODYAFTER") - y_of(p, "BODYBEFORE")
    };
    assert!(
        gap(&spaced) - gap(&none) > 29.0,
        "30pt of space_before must apply mid-page: {} vs {}",
        gap(&spaced),
        gap(&none),
    );
}

/// §17.3.3.1: an explicit page break in the paragraph *before* a continuous
/// break still moves the following content to a new page. The break is deferred
/// to the start of the next block, so it has to survive the section boundary.
#[test]
fn an_inline_page_break_before_the_continuous_break_is_honoured() {
    let pages = layout(&format!(
        r#"<w:p><w:r><w:t>BODYBEFORE</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>
           <w:p><w:pPr>{}</w:pPr></w:p>{}{}"#,
        sect_pr("rIdH1", false),
        para("BODYAFTER"),
        sect_pr("rIdH1", true),
    ));

    assert_eq!(
        pages.len(),
        2,
        "the page break before the section break must still start a new page"
    );
}

/// §17.3.1.26 `pageBreakBefore` on the **first** paragraph after the break.
///
/// A different path from the case above: there the break is a flag deferred
/// across the section boundary with the continuation, here it is a property of
/// the succeeding section's own first block — reached while that section thinks
/// it is starting, on a page it did not create.
#[test]
fn a_page_break_before_the_first_paragraph_after_the_break_is_honoured() {
    let pages = layout(&format!(
        r#"{}<w:p><w:pPr>{}</w:pPr></w:p>
           <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>BODYAFTER</w:t></w:r></w:p>{}"#,
        para("BODYBEFORE"),
        sect_pr("rIdH1", false),
        sect_pr("rIdH1", true),
    ));

    assert_eq!(pages.len(), 2, "pageBreakBefore must start a new page");
    let on = |i: usize| drawn_text(std::slice::from_ref(&pages[i]));
    assert!(
        on(0).iter().any(|t| t.contains("BODYBEFORE")),
        "the preceding section keeps the page it was already on: {:?}",
        on(0)
    );
    assert!(
        on(1).iter().any(|t| t.contains("BODYAFTER")),
        "and the succeeding section starts on the next one: {:?}",
        on(1)
    );
}

/// The same, as an inline `w:br` on the first run after the break rather than a
/// paragraph property — §17.3.3.1 rather than §17.3.1.26. The run precedes the
/// text, so the whole paragraph moves.
#[test]
fn an_inline_page_break_after_the_continuous_break_is_honoured() {
    let pages = layout(&format!(
        r#"{}<w:p><w:pPr>{}</w:pPr></w:p>
           <w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>BODYAFTER</w:t></w:r></w:p>{}"#,
        para("BODYBEFORE"),
        sect_pr("rIdH1", false),
        sect_pr("rIdH1", true),
    ));

    assert_eq!(pages.len(), 2, "the inline break must start a new page");
    assert!(
        drawn_text(std::slice::from_ref(&pages[1]))
            .iter()
            .any(|t| t.contains("BODYAFTER")),
        "text after the inline break belongs on the new page"
    );
}

/// §20.4.2.17 `wrapSquare` — a left-hand floating shape 200pt wide, anchored to
/// its paragraph, that text must flow to the right of.
const FLOAT_PARA: &str = r#"<w:p><w:r><w:drawing>
  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
             relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
    <wp:extent cx="2540000" cy="2540000"/>
    <wp:wrapSquare wrapText="bothSides"/>
    <wp:docPr id="1" name="Box"/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp>
        <wps:cNvSpPr/>
        <wps:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="2540000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
        </wps:spPr>
      </wps:wsp>
    </a:graphicData></a:graphic>
  </wp:anchor>
</w:drawing></w:r></w:p>"#;

/// §20.4.2: a float is a property of the *page*, not of the section that
/// registered it. Text after a continuous break shares the page, so it must
/// wrap around a float placed before the break — dropping the float set at the
/// boundary lets that text run straight through the shape.
#[test]
fn text_after_the_break_wraps_around_a_float_registered_before_it() {
    let x_of = |pages: &[dxpdf::render::layout::draw_command::LayoutedPage], needle: &str| {
        use dxpdf::render::layout::draw_command::DrawCommand;
        pages
            .iter()
            .flat_map(|p| &p.commands)
            .find_map(|c| match c {
                DrawCommand::Text { text, position, .. } if text.contains(needle) => {
                    Some(position.x.raw())
                }
                _ => None,
            })
            .unwrap_or_else(|| panic!("{needle:?} was never drawn"))
    };

    let pages = layout(&format!(
        "{FLOAT_PARA}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        sect_pr("rIdH1", false),
        para("BODYAFTER"),
        sect_pr("rIdH1", true),
    ));

    // 2540000 EMU = 200pt, anchored at the left margin (56.7pt).
    let left_margin = 1134.0 / 20.0;
    let x = x_of(&pages, "BODYAFTER");
    assert!(
        x > left_margin + 100.0,
        "text after the break must clear the 200pt float registered before it, \
         but starts at x={x} (left margin {left_margin})"
    );
}

// ── §17.6.4 column changes across a continuous break (issue #83, plan §4) ────

/// A `sectPr` with an explicit column count. `w:space` matches [`PG`]'s gutter
/// so the derived x positions below are predictable.
fn sect_pr_cols(rel: &str, continuous: bool, cols: u32) -> String {
    let ty = if continuous {
        r#"<w:type w:val="continuous"/>"#
    } else {
        ""
    };
    format!(
        r#"<w:sectPr>{ty}<w:headerReference w:type="default" r:id="{rel}" {R_NS}/>{PG}<w:cols w:num="{cols}" w:space="720"/></w:sectPr>"#
    )
}

/// Every x a `FILL*` line was drawn at, deduplicated to the nearest point.
fn fill_columns(pages: &[dxpdf::render::layout::draw_command::LayoutedPage]) -> Vec<i32> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    let mut xs: Vec<i32> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.starts_with("FILL") => {
                Some(position.x.raw().round() as i32)
            }
            _ => None,
        })
        .collect();
    xs.sort_unstable();
    xs.dedup();
    xs
}

/// §17.6.4: changing the column count is the main reason a continuous break
/// exists. Text after a 1 → 2 break must flow in two columns.
///
/// Left margin is 1134tw = 56.7pt and the content width 481.9pt, so with
/// `w:space="720"` (36pt) each column is 222.95pt and the second starts at
/// 56.7 + 222.95 + 36 ≈ 315.6.
#[test]
fn continuous_break_can_increase_the_column_count() {
    let filler: String = (0..90).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{filler}{}",
        para("BODYBEFORE"),
        sect_pr_cols("rIdH1", false, 1),
        sect_pr_cols("rIdH1", true, 2),
    ));

    let xs = fill_columns(&pages);
    assert!(
        xs.iter().any(|&x| (x - 57).abs() <= 1),
        "the first column starts at the left margin: {xs:?}"
    );
    assert!(
        xs.iter().any(|&x| (x - 316).abs() <= 2),
        "content must reach the second column at x≈316: {xs:?}"
    );
}

/// The converse, and the reason the incoming column index is **not** carried:
/// the succeeding section may have fewer columns than the index it would
/// inherit.
///
/// This is a panic guard, not a preference. Seeding the continuation with a
/// stale index — 3 columns before the break, 1 after — makes this fixture fail
/// with `index out of bounds: the len is 1 but the index is 2`, from the
/// `config.columns[state.current_col]` in `layout_section_with_clearance`.
/// Restarting at column 0 is what the change means *and* what keeps this a
/// layout question.
#[test]
fn continuous_break_can_decrease_the_column_count() {
    let filler: String = (0..90).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{filler}{}",
        para("BODYBEFORE"),
        sect_pr_cols("rIdH1", false, 3),
        sect_pr_cols("rIdH1", true, 1),
    ));

    let xs = fill_columns(&pages);
    assert_eq!(
        xs.len(),
        1,
        "single-column content occupies exactly one x: {xs:?}"
    );
    assert!(
        (xs[0] - 57).abs() <= 1,
        "and that x is the left margin: {xs:?}"
    );
}

/// §17.6.22: the shared-page handoff resumes below the **deepest** column of
/// the multi-column section, not below the last one. Column 0 fills to the
/// page bottom while column 1 holds only the overflow tail, so the raw cursor
/// at the end of the section sits high up in column 1 — resuming there paints
/// the next section's full-width content straight over column 0's tail (the
/// `013_columns` regression: "Back to one column" drawn over column
/// paragraphs).
#[test]
fn a_continuous_break_resumes_below_the_deepest_column() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let filler: String = (0..60).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{filler}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        sect_pr_cols("rIdH1", false, 2),
        para("AFTERBREAK"),
        sect_pr_cols("rIdH1", true, 1),
    ));

    let mut after: Option<(usize, f32)> = None;
    for (page_idx, page) in pages.iter().enumerate() {
        for command in &page.commands {
            if let DrawCommand::Text { text, position, .. } = command {
                if text.contains("AFTERBREAK") {
                    after = Some((page_idx, position.y.raw()));
                }
            }
        }
    }
    let (after_page, after_y) = after.expect("AFTERBREAK was drawn");

    // Every column paragraph on the shared page sits above the continued
    // section — nothing may remain below it, or the two are overlapping.
    for command in &pages[after_page].commands {
        if let DrawCommand::Text { text, position, .. } = command {
            if text.starts_with("FILL") {
                assert!(
                    position.y.raw() < after_y,
                    "column text {text:?} at y={} lies below the continued \
                     section at y={after_y} — the continuation resumed above \
                     the deepest column",
                    position.y.raw(),
                );
            }
        }
    }
}

// ── §17.6.22 the break changes nothing when the sections agree (issue #83) ──

/// Every drawn primitive as `page kind position [text]`.
///
/// Coarse enough to be readable in a failure and precise enough that a lost
/// table row, a moved border or a duplicated line shows up — where probing one
/// line's `y` would only catch a whole-page shift.
fn fingerprint(pages: &[dxpdf::render::layout::draw_command::LayoutedPage]) -> Vec<String> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    let seg = |l: &dxpdf::render::geometry::PtLineSegment| {
        format!(
            "{:.2},{:.2}-{:.2},{:.2}",
            l.start.x.raw(),
            l.start.y.raw(),
            l.end.x.raw(),
            l.end.y.raw()
        )
    };
    pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .map(|(i, c)| match c {
            DrawCommand::Text { position, text, .. } => format!(
                "p{i} text {:.2},{:.2} {text}",
                position.x.raw(),
                position.y.raw()
            ),
            DrawCommand::Line { line, .. } => format!("p{i} line {}", seg(line)),
            DrawCommand::Underline { line, .. } => format!("p{i} underline {}", seg(line)),
            other => format!("p{i} {other:?}"),
        })
        .collect()
}

/// §17.6.22 moves a *section* boundary, not a page boundary. Two sections whose
/// properties agree must therefore render exactly as the same content with no
/// break at all — every command, on every page, at the same position.
///
/// `<w:p/>` stands in for the break's own carrier paragraph on the unbroken
/// side: a mid-document `sectPr` lives on the `w:pPr` of the section's last
/// paragraph, and that paragraph mark is real content that renders a blank
/// line. Omitting it would make the two documents differ by a line and the
/// comparison would measure the fixture.
#[track_caller]
fn assert_break_changes_nothing(before: &str, after: &str, what: &str) {
    let split = layout(&format!(
        "{before}<w:p><w:pPr>{}</w:pPr></w:p>{after}{}",
        sect_pr("rIdH1", false),
        sect_pr("rIdH1", true),
    ));
    let joined = layout(&format!("{before}<w:p/>{after}{}", sect_pr("rIdH1", false)));
    assert_eq!(fingerprint(&split), fingerprint(&joined), "{what}");
}

/// A paragraph bound to the one after it.
fn keep_next(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

/// A table of `rows` single-cell rows.
fn table(rows: usize, prefix: &str) -> String {
    let body: String = (0..rows)
        .map(|i| {
            format!(
                r#"<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{prefix}{i}</w:t></w:r></w:p></w:tc></w:tr>"#
            )
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>{body}</w:tbl>"#
    )
}

/// §17.3.1.15: a `keepNext` chain in the succeeding section is fitted against
/// the bottom it *inherited*, and must still move as a unit when it does not
/// fit. The chain is positioned near the page bottom deliberately — a chain
/// with room to spare would pass whatever the inherited bounds were.
#[test]
fn a_keep_next_chain_after_the_break_still_moves_as_a_unit() {
    // Tuned so KEEPA alone *would* fit on the shared page and KEEPB would not:
    // the §17.3.1.15 move is then forced, rather than the whole chain simply
    // arriving after the page was already full. 60 fillers leave both on
    // page 1 and 63 push both off it; only this narrow window tests anything.
    let filler: String = (0..61).map(|i| para(&format!("FILL{i}"))).collect();
    let chain = format!("{}{}", keep_next("KEEPA"), para("KEEPB"));

    // The invariant first: the break must not perturb the chain's placement.
    assert_break_changes_nothing(
        &filler,
        &chain,
        "a keepNext chain must paginate the same either side of a continuous break",
    );

    // Then confirm the fixture actually exercises the chain rather than
    // leaving it comfortably mid-page.
    let split = layout(&format!(
        "{filler}<w:p><w:pPr>{}</w:pPr></w:p>{chain}{}",
        sect_pr("rIdH1", false),
        sect_pr("rIdH1", true),
    ));
    let page_of = |needle: &str| {
        split
            .iter()
            .position(|p| {
                drawn_text(std::slice::from_ref(p))
                    .iter()
                    .any(|t| t == needle)
            })
            .unwrap_or_else(|| panic!("{needle} was never drawn"))
    };
    assert_eq!(
        page_of("KEEPA"),
        page_of("KEEPB"),
        "§17.3.1.15: the chain must not be torn across the page boundary"
    );
    assert_eq!(
        page_of("KEEPA"),
        1,
        "the fixture must push the chain onto a second page, or it tests nothing"
    );
}

/// §17.4.1: a table starting from the continuation cursor splits across pages
/// against the inherited bounds, like any other block on a shared page.
#[test]
fn a_table_spanning_the_break_paginates_unchanged() {
    assert_break_changes_nothing(
        &para("BODYBEFORE"),
        &table(60, "ROW"),
        "a table that splits across pages must do so identically either side of \
         a continuous break",
    );
}

/// The preceding section ending in a table, rather than beginning with one:
/// the handover carries `last_para_start_y` and the table-chain state, and a
/// table is the block most likely to expose either.
#[test]
fn a_table_before_the_break_leaves_the_page_in_the_same_state() {
    assert_break_changes_nothing(
        &table(50, "ROW"),
        &para("BODYAFTER"),
        "content after the break must resume where a table before it left off",
    );
}

// ── §17.10 footer clearance on a shared page (issue #83) ────────────────────

/// A `sectPr` with the short header throughout and the given footer slot, so
/// only the footer varies between the documents compared below.
fn sect_pr_ftr(rel: &str, continuous: bool) -> String {
    let ty = if continuous {
        r#"<w:type w:val="continuous"/>"#
    } else {
        ""
    };
    format!(
        r#"<w:sectPr>{ty}<w:headerReference w:type="default" r:id="rIdH1" {R_NS}/><w:footerReference w:type="default" r:id="{rel}" {R_NS}/>{PG}</w:sectPr>"#
    )
}

/// Every `FILL*` line as `(page index, rounded y)`, in emission order — the
/// whole distribution, so a difference anywhere shows up rather than only at
/// whichever line a single probe happens to pick.
fn fill_positions(
    pages: &[dxpdf::render::layout::draw_command::LayoutedPage],
) -> Vec<(usize, i32)> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    pages
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.commands.iter().map(move |c| (i, c)))
        .filter_map(|(i, c)| match c {
            DrawCommand::Text { text, position, .. } if text.starts_with("FILL") => {
                Some((i, position.y.raw().round() as i32))
            }
            _ => None,
        })
        .collect()
}

const FOOTER_FILLER: usize = 60;

/// The filler lives in the **succeeding** section, unlike the header cases.
///
/// A footer only becomes observable where content reaches the page bottom, and
/// content that reaches it spills — which would make the preceding section's
/// last page a *later* page and take the shared page out of the comparison. Put
/// the filler after the break instead and the shared page stays page 1, while
/// the bottom the succeeding section inherits is exactly what the fit under
/// test produces.
fn shared_page_footer(
    first_ftr: &str,
    second_ftr: &str,
) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    let filler: String = (0..FOOTER_FILLER)
        .map(|i| para(&format!("FILL{i}")))
        .collect();
    layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{filler}{}",
        para("BODYBEFORE"),
        sect_pr_ftr(first_ftr, false),
        sect_pr_ftr(second_ftr, true),
    ))
}

/// The unbroken reference for [`shared_page_footer`].
fn single_section_footer(ftr: &str) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    let filler: String = (0..FOOTER_FILLER)
        .map(|i| para(&format!("FILL{i}")))
        .collect();
    layout(&format!(
        "{}<w:p/>{filler}{}",
        para("BODYBEFORE"),
        sect_pr_ftr(ftr, false),
    ))
}

/// §17.10: the shared page is drawn with the succeeding section's *footer*, so
/// the bottom the preceding section hands over must be that footer's — not the
/// one it laid itself out against.
#[test]
fn taller_succeeding_footer_pulls_the_shared_page_bottom_up() {
    // Guard the fixture before trusting the comparison: if the two footers
    // produced the same layout, every assertion below would pass vacuously.
    assert_ne!(
        fill_positions(&single_section_footer("rIdF1")),
        fill_positions(&single_section_footer("rIdF2")),
        "the two footer slots must differ enough to move content"
    );

    assert_eq!(
        fill_positions(&shared_page_footer("rIdF1", "rIdF2")),
        fill_positions(&single_section_footer("rIdF2")),
        "content on the shared page must break where the tall footer puts it, \
         not where the short footer it was laid out against did"
    );
}

/// The converse: a *shorter* succeeding footer must give the reserved band back
/// to the body rather than leaving it unused.
#[test]
fn shorter_succeeding_footer_reclaims_the_band() {
    assert_eq!(
        fill_positions(&shared_page_footer("rIdF2", "rIdF1")),
        fill_positions(&single_section_footer("rIdF1")),
        "a shorter succeeding footer must not retain the preceding section's \
         reservation"
    );
}

/// Equal slots are the no-op case, for footers as for headers.
#[test]
fn equal_footers_leave_the_shared_page_unchanged() {
    assert_eq!(
        fill_positions(&shared_page_footer("rIdF1", "rIdF1")),
        fill_positions(&single_section_footer("rIdF1")),
        "the shared page's own bottom is unchanged"
    );
}

// ── §17.6.22 chained breaks — three sections on one page (issue #83) ─────────

/// Three sections on one page: `a` closes the first, then two continuous breaks
/// carry the same physical page through `b` and `c`.
fn chained(a: &str, b: &str, c: &str) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr(a, false),
        para("BODYMIDDLE"),
        sect_pr(b, true),
        para("BODYAFTER"),
        sect_pr(c, true),
    ))
}

/// The unbroken reference for [`chained`] — same content, one section, so `hdr`
/// governs the whole page.
///
/// Two empty paragraphs, not none: in [`chained`] each break rides on a `w:p`'s
/// `w:pPr`, and those paragraph marks render blank lines. Omitting them here
/// would make the comparison measure the fixture instead of the engine.
fn single_section_of_three(hdr: &str) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    layout(&format!(
        "{}<w:p/>{}<w:p/>{}{}",
        para("BODYBEFORE"),
        para("BODYMIDDLE"),
        para("BODYAFTER"),
        sect_pr(hdr, false),
    ))
}

/// §17.6.22: a chain of continuous breaks is still one page. Nothing about the
/// second break makes it a page boundary that the first one was not.
#[test]
fn chained_continuous_breaks_share_one_page() {
    let pages = chained("rIdH1", "rIdH1", "rIdH1");
    assert_eq!(pages.len(), 1, "three sections, two breaks, one page");

    let texts = drawn_text(&pages);
    let at = |needle: &str| texts.iter().position(|t| t.contains(needle));
    assert!(
        at("BODYBEFORE") < at("BODYMIDDLE") && at("BODYMIDDLE") < at("BODYAFTER"),
        "document order survives both breaks: {texts:?}"
    );
}

/// §17.6 / §17.10: the ownership rule is "the last section on the page", not
/// "the section after this one". With three sections on a page the two readings
/// disagree, and only a chain can tell them apart.
#[test]
fn chained_breaks_render_only_the_last_sections_header() {
    // Middle section takes the tall header; only the *last* one's may be drawn.
    let pages = chained("rIdH1", "rIdH2", "rIdH1");
    let texts = drawn_text(&pages);
    assert_eq!(
        texts.iter().filter(|t| t.contains("HDRSHORT")).count(),
        1,
        "exactly one header is drawn, and it is the last section's: {texts:?}"
    );
    assert!(
        !texts.iter().any(|t| t.contains("HDRTALL")),
        "a middle section of the chain owns nothing: {texts:?}"
    );
}

/// §17.10.1, chained: **every** section's content on the shared page must clear
/// the header actually drawn on it — which is the last section's.
///
/// This is the case a single break cannot reach. Fitting each section against
/// its immediate successor is enough when there are two: the successor is the
/// owner. With three it is not — the first section gets pinned to the middle
/// section's clearance while the page is drawn with the last section's.
#[test]
fn chained_breaks_clear_the_last_sections_header() {
    // Only the final section carries the tall header, so a fit that stops at
    // the immediate successor never sees it.
    let pages = chained("rIdH1", "rIdH1", "rIdH2");
    let reference = single_section_of_three("rIdH2");

    for needle in ["BODYBEFORE", "BODYMIDDLE", "BODYAFTER"] {
        assert_eq!(
            y_of(&pages, needle),
            y_of(&reference, needle),
            "{needle} must sit where the header drawn on the page puts it"
        );
    }
}

/// The converse: a chain whose last section has the *shorter* header must
/// reclaim the band, rather than every section keeping the tall clearance it
/// was laid out against.
#[test]
fn chained_breaks_reclaim_the_band_when_the_last_header_is_shorter() {
    let pages = chained("rIdH2", "rIdH2", "rIdH1");
    let reference = single_section_of_three("rIdH1");

    for needle in ["BODYBEFORE", "BODYMIDDLE", "BODYAFTER"] {
        assert_eq!(
            y_of(&pages, needle),
            y_of(&reference, needle),
            "{needle} must not retain a preceding section's clearance"
        );
    }
}

/// A chain does **not** mean every section lands on the same page. A middle
/// section that overflows commits the shared page as its own, and the sections
/// after it start on a later one — so the search for the owner has to stop
/// there rather than run to the end of the chain.
///
/// Only the final section carries the tall header, and it never appears on
/// page 1. Taking the last section of the chain unconditionally would push
/// page 1's content down to clear a header drawn on page 2.
#[test]
fn a_spilling_middle_section_owns_the_page_it_overflows() {
    let filler: String = (0..80).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{filler}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr("rIdH1", false),
        sect_pr("rIdH1", true),
        para("BODYAFTER"),
        sect_pr("rIdH2", true),
    ));
    assert!(
        pages.len() >= 2,
        "the middle section must actually overflow for this to test anything"
    );

    let texts_on = |page: &dxpdf::render::layout::draw_command::LayoutedPage| {
        drawn_text(std::slice::from_ref(page))
    };
    let first = texts_on(&pages[0]);
    assert!(
        first.iter().any(|t| t.contains("HDRSHORT"))
            && !first.iter().any(|t| t.contains("HDRTALL")),
        "page 1 belongs to the section that overflows onto page 2: {first:?}"
    );

    assert_eq!(
        y_of(&pages, "BODYBEFORE"),
        y_of(&single_section("rIdH1"), "BODYBEFORE"),
        "content on page 1 clears the header drawn on page 1, not the one the \
         last section of the chain draws on page 2"
    );
}

/// §17.11.23: one separator per page holds however many sections contribute to
/// it. Carried notes have to chain through *two* handovers, each of which could
/// drop or re-flush them.
#[test]
fn footnotes_from_three_chained_sections_share_one_separator() {
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para_with_note("BODYBEFORE", 2),
        sect_pr("rIdH1", false),
        para_with_note("BODYMIDDLE", 4),
        sect_pr("rIdH1", true),
        para_with_note("BODYAFTER", 3),
        sect_pr("rIdH1", true),
    ));

    assert_eq!(pages.len(), 1, "one shared physical page");
    let texts = drawn_text(&pages);
    for note in ["NOTEBEFORE", "NOTEMIDDLE", "NOTEAFTER"] {
        assert!(
            texts.iter().any(|t| t.contains(note)),
            "{note} must survive being carried across two breaks: {texts:?}"
        );
    }
    assert_eq!(
        footnote_separator_count(&pages),
        1,
        "§17.11.23: one separator per page, however many sections share it"
    );
}

/// §17.6.4: a chain may change the column count at each break. The middle
/// section's count must not leak into the last one — the reset is per break,
/// not once per chain.
#[test]
fn chained_breaks_apply_each_sections_column_count() {
    let filler: String = (0..90).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}<w:p><w:pPr>{}</w:pPr></w:p>{filler}{}",
        para("BODYBEFORE"),
        sect_pr_cols("rIdH1", false, 1),
        para("BODYMIDDLE"),
        sect_pr_cols("rIdH1", true, 3),
        sect_pr_cols("rIdH1", true, 1),
    ));

    let xs = fill_columns(&pages);
    assert_eq!(
        xs.len(),
        1,
        "the last section is single-column; the middle section's three must not \
         survive its break: {xs:?}"
    );
    assert!(
        (xs[0] - 57).abs() <= 1,
        "and that column is at the left margin: {xs:?}"
    );
}

/// §17.6.4: the new column set starts at the shared cursor, not at the page
/// top — the preceding section's content is still above it on this page.
#[test]
fn a_new_column_set_starts_below_the_preceding_sections_content() {
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr_cols("rIdH1", false, 1),
        para("FILL0"),
        sect_pr_cols("rIdH1", true, 2),
    ));

    assert!(
        y_of(&pages, "FILL0") > y_of(&pages, "BODYBEFORE"),
        "the two-column area begins below the single-column content it follows"
    );
}

// ── §17.6.22 promotion when page setup differs (issue #112) ─────────────────

/// Same as [`PG`] but with a shorter bottom margin (720 vs 1134 twips) —
/// mirrors `sample-docx-files-sample3.docx`'s reproduction (720 vs 1440
/// there; the exact values don't matter, only that they differ).
const PG_SHORT_BOTTOM: &str = r#"<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="720" w:left="1134" w:header="567" w:footer="567"/>"#;

/// Same as [`PG`] but landscape — width and height swapped.
const PG_LANDSCAPE: &str = r#"<w:pgSz w:w="16838" w:h="11906"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567"/>"#;

/// A `sectPr` referencing header `rel` with a custom `w:pgSz`/`w:pgMar` block
/// in place of the fixed [`PG`], optionally as a continuous break.
fn sect_pr_pg(rel: &str, continuous: bool, pg: &str) -> String {
    let ty = if continuous {
        r#"<w:type w:val="continuous"/>"#
    } else {
        ""
    };
    format!(
        r#"<w:sectPr>{ty}<w:headerReference w:type="default" r:id="{rel}" {R_NS}/>{pg}</w:sectPr>"#
    )
}

/// Issue #112: one physical page cannot have two bottom margins, so a
/// continuous break into a section with a different one must become a page
/// break instead of a shared page — the direct fixture analog of
/// `sample-docx-files-sample3.docx`.
#[test]
fn a_margin_difference_promotes_the_break_to_a_page_break() {
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr("rIdH1", false),
        para("BODYAFTER"),
        sect_pr_pg("rIdH1", true, PG_SHORT_BOTTOM),
    ));

    assert_eq!(
        pages.len(),
        2,
        "a bottom-margin change makes one physical page impossible"
    );
    let texts_on = |page: &dxpdf::render::layout::draw_command::LayoutedPage| {
        drawn_text(std::slice::from_ref(page))
    };
    assert!(
        texts_on(&pages[0]).iter().any(|t| t.contains("BODYBEFORE")),
        "the preceding section keeps its own page"
    );
    assert!(
        texts_on(&pages[1]).iter().any(|t| t.contains("BODYAFTER")),
        "the succeeding section starts a fresh page"
    );
}

/// Issue #112, the other named candidate: a page-size change is just as
/// physically impossible to share as a margin change.
#[test]
fn a_page_size_difference_promotes_the_break_to_a_page_break() {
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr("rIdH1", false),
        para("BODYAFTER"),
        sect_pr_pg("rIdH1", true, PG_LANDSCAPE),
    ));

    assert_eq!(
        pages.len(),
        2,
        "a page-size change makes one physical page impossible"
    );
}

/// The converse, explicit rather than incidental: §17.6.4 columns are the main
/// legitimate reason a continuous break exists, and a column-only change must
/// not be swept up by the page-setup promotion rule.
#[test]
fn a_column_only_change_still_shares_the_page() {
    let filler: String = (0..90).map(|i| para(&format!("FILL{i}"))).collect();
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{filler}{}",
        para("BODYBEFORE"),
        sect_pr_cols("rIdH1", false, 1),
        sect_pr_cols("rIdH1", true, 2),
    ));

    assert_eq!(
        pages.len(),
        1,
        "a column-count-only change must not be promoted to a page break"
    );
}

/// §17.6.22, chained: a chain may promote at one link and share at another.
/// Only the second break (section 2 → section 3) changes margins, so the
/// first break still shares a page and the second starts a fresh one — the
/// case a single break cannot distinguish from "promote the whole chain".
#[test]
fn only_the_promoted_link_in_a_chain_breaks_the_page() {
    let pages = layout(&format!(
        "{}<w:p><w:pPr>{}</w:pPr></w:p>{}<w:p><w:pPr>{}</w:pPr></w:p>{}{}",
        para("BODYBEFORE"),
        sect_pr("rIdH1", false),
        para("BODYMIDDLE"),
        sect_pr("rIdH1", true),
        para("BODYAFTER"),
        sect_pr_pg("rIdH1", true, PG_SHORT_BOTTOM),
    ));

    assert_eq!(
        pages.len(),
        2,
        "only the second link promotes, so the chain lands on two pages"
    );
    let texts_on = |page: &dxpdf::render::layout::draw_command::LayoutedPage| {
        drawn_text(std::slice::from_ref(page))
    };
    let first = texts_on(&pages[0]);
    assert!(
        first.iter().any(|t| t.contains("BODYBEFORE"))
            && first.iter().any(|t| t.contains("BODYMIDDLE")),
        "the first, non-promoted break still shares page 1: {first:?}"
    );
    assert!(
        texts_on(&pages[1]).iter().any(|t| t.contains("BODYAFTER")),
        "the promoted break starts the third section on its own page: {:?}",
        texts_on(&pages[1])
    );
}
