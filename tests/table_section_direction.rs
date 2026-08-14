//! §17.6.6 `w:bidi` on `<w:sectPr>` — which margin a section's tables measure
//! from.
//!
//! # This is not `w:bidiVisual`
//!
//! §17.4.1 `w:bidiVisual` reverses the cells *within* a table and is asserted in
//! `tests/table_bidi_visual.rs`. It says nothing about where the table sits.
//! §17.6.6 is the other switch: "Right to Left Section Layout", which decides
//! which side of the content area is the **leading** one — and therefore what
//! §17.4.50 `tblInd` ("Table Indent from **Leading** Margin") measures from and
//! which edge §17.4.28 `w:jc`'s `start`/`end` name.
//!
//! A document may set either without the other, so conflating them would be
//! wrong in both directions. These tests set only `w:bidi`, and every table in
//! them keeps its columns in source order.
//!
//! # The `left` = `start` reading
//!
//! `<w:jc w:val="left"/>` on a table in a `w:bidi` section aligns it **right**,
//! because Transitional `left` *is* Strict `start` — the parse seam maps both to
//! `Alignment::Start` (`st_enums.rs`), and ISO/IEC 29500's own
//! Transitional→Strict migration maps `left`↔`start` losslessly, which only
//! holds if they are the same edge. `paragraph::line_emit::align_offset` already
//! records and applies that same reading for paragraphs; this extends it to the
//! one block-level thing that was still pinned to the left.
//!
//! Every assertion is a relation between the same document with and without
//! `<w:bidi/>`, so no page origin or glyph metric is pinned.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn make_docx(document_xml: &str) -> Vec<u8> {
    let mut buf = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
        let o = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", o).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

        zip.start_file("_rels/.rels", o).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/document.xml", o).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();
        zip.finish().unwrap();
    }
    buf
}

/// A single shaded 200pt table carrying `tbl_pr`, in a section that is RTL when
/// `bidi`. The page is 12240 twips wide with 1440-twip margins, so the content
/// area is 468pt starting at x = 72.
fn layout(bidi: bool, tbl_pr: &str) -> Vec<LayoutedPage> {
    let flag = if bidi { "<w:bidi/>" } else { "" };
    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="4000" w:type="dxa"/>
        <w:tblLayout w:type="fixed"/>
        {tbl_pr}
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr><w:tc>
        <w:tcPr>
          <w:tcW w:w="4000" w:type="dxa"/>
          <w:shd w:val="clear" w:color="auto" w:fill="C0FFEE"/>
        </w:tcPr>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:sectPr>
      {flag}
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"#
    );
    let doc = dxpdf::docx::parse(&make_docx(&document_xml)).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// The shaded table's left edge and width.
fn table_box(pages: &[LayoutedPage]) -> (f32, f32) {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Rect { rect, color }
                if (color.r, color.g, color.b) == (0xC0, 0xFF, 0xEE) =>
            {
                Some((rect.origin.x.raw(), rect.size.width.raw()))
            }
            _ => None,
        })
        .expect("the shaded table")
}

/// Content area: 12240 − 2×1440 twips = 468pt, from x = 72.
const CONTENT_LEFT: f32 = 72.0;
const CONTENT_RIGHT: f32 = 72.0 + 468.0;

// ── §17.4.28 `w:jc` ─────────────────────────────────────────────────────────

/// A table with no `w:jc` takes the section's leading margin: the left in an
/// ordinary section, the right in a `w:bidi` one.
///
/// Absent alignment is `Alignment::Start`, so this is the same claim as the
/// `jc="left"` case below — stated separately because "no alignment at all" is
/// the overwhelmingly common shape and a fix could easily reach one and not the
/// other.
#[test]
fn a_table_with_no_alignment_sits_at_the_sections_leading_margin() {
    let (ltr_x, w) = table_box(&layout(false, ""));
    let (rtl_x, rtl_w) = table_box(&layout(true, ""));

    assert_eq!(w, 200.0, "4000 twips");
    assert_eq!(rtl_w, w, "the table itself does not change size");
    assert_eq!(ltr_x, CONTENT_LEFT, "flush left without w:bidi");
    assert_eq!(rtl_x + w, CONTENT_RIGHT, "flush right with it");
}

/// `w:jc="left"` is `start`, so it follows the section too — the reading the
/// module doc argues from the Transitional→Strict migration.
#[test]
fn jc_left_is_the_start_edge_and_follows_the_section() {
    let (ltr_x, w) = table_box(&layout(false, r#"<w:jc w:val="left"/>"#));
    let (rtl_x, _) = table_box(&layout(true, r#"<w:jc w:val="left"/>"#));

    assert_eq!(ltr_x, CONTENT_LEFT);
    assert_eq!(
        rtl_x + w,
        CONTENT_RIGHT,
        "`left` means `start`, not the left"
    );
}

/// …and `w:jc="right"` is `end`, so it mirrors the other way.
///
/// The pair is what makes either meaningful: a renderer that ignored the
/// section would fail both, and one that reversed the sense of `Start`/`End`
/// without consulting it would pass this and fail the one above.
#[test]
fn jc_right_is_the_end_edge_and_follows_the_section() {
    let (ltr_x, w) = table_box(&layout(false, r#"<w:jc w:val="right"/>"#));
    let (rtl_x, _) = table_box(&layout(true, r#"<w:jc w:val="right"/>"#));

    assert_eq!(ltr_x + w, CONTENT_RIGHT);
    assert_eq!(rtl_x, CONTENT_LEFT, "`right` means `end`");
}

/// The control: `center` has no leading edge to follow, so the section's
/// direction must not move it.
///
/// Without this, a fix that mirrored every table about the content area would
/// satisfy all three assertions above.
#[test]
fn a_centred_table_does_not_move_with_the_section() {
    let (ltr_x, _) = table_box(&layout(false, r#"<w:jc w:val="center"/>"#));
    let (rtl_x, _) = table_box(&layout(true, r#"<w:jc w:val="center"/>"#));

    assert_eq!(rtl_x, ltr_x, "centre is centre in either direction");
}

// ── §17.4.50 `w:tblInd` ─────────────────────────────────────────────────────

/// `tblInd` is an indent from the **leading** margin, so it measures inward
/// from the right in a `w:bidi` section.
///
/// Asserted as the gap between the table and each margin, so the claim is which
/// side the indent is on rather than what x it produces.
#[test]
fn tbl_ind_measures_from_the_leading_margin() {
    let ind = r#"<w:tblInd w:w="1440" w:type="dxa"/>"#; // 72pt
    let (ltr_x, w) = table_box(&layout(false, ind));
    let (rtl_x, _) = table_box(&layout(true, ind));

    assert_eq!(ltr_x - CONTENT_LEFT, 72.0, "72pt in from the left");
    assert_eq!(
        CONTENT_RIGHT - (rtl_x + w),
        72.0,
        "and 72pt in from the right once the section is RTL"
    );
}
