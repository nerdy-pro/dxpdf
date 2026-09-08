//! Which margin a table measures from — §17.4.28 `w:jc` and §17.4.50 `tblInd`
//! read against a direction.
//!
//! # Two switches, and which one wins
//!
//! A table's placement is stated in **logical** terms: `ST_JcTable`'s Strict
//! spelling is `start`/`end` (Transitional `left`/`right` are aliases of those,
//! and `st_enums.rs` already parses them as one), and §17.4.50 is literally
//! "Table Indent from **Leading** Margin". Two elements can say which physical
//! edge is leading:
//!
//! * §17.4.1 `w:bidiVisual` — *this table* runs right to left;
//! * §17.6.6 `w:bidi` on `<w:sectPr>` — *this section* does.
//!
//! The table's own declaration wins where it is present, and the section's
//! answers when it is absent. That is ordinary OOXML property scoping, and it
//! keeps the two independent: a document may set either alone, and both
//! spellings appear in the wild.
//!
//! # The evidence for the `bidiVisual` half
//!
//! `test-files/bidi-visual-table.docx` holds two tables that differ in exactly
//! one element, `w:bidiVisual`. In Word the second is flush **right** and the
//! control is flush left. That fixture is what rules out the obvious rival
//! explanation: every cell paragraph in *both* tables carries `w:bidi`, so if
//! paragraph direction placed the table, both would move. Only one does.
//!
//! This engine got that wrong on the first pass — §17.4.1's text describes cell
//! order, and it was read as scoping the element to cell order. The element
//! makes the table a right-to-left table; where its cells sit is one
//! consequence and where the table sits is another.
//!
//! # What is still undecided
//!
//! `<w:bidiVisual w:val="0"/>` inside a `w:bidi` section is asserted below as
//! left-to-right placement, on the plain reading that a table saying it is not
//! right-to-left is not right-to-left. No Word render stands behind that one —
//! a document pairing the two would settle it.
//!
//! Every other assertion here is a relation between the same document with and
//! without one element, so no page origin or glyph metric is pinned.

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
/// `section_rtl`. The page is 12240 twips wide with 1440-twip margins, so the
/// content area is 468pt starting at x = 72.
///
/// The table has **one column**, so `<w:bidiVisual/>` in `tbl_pr` mirrors it
/// onto itself and every assertion below is about placement alone. Cell order is
/// `tests/table_bidi_visual.rs`'s subject.
fn layout(section_rtl: bool, tbl_pr: &str) -> Vec<LayoutedPage> {
    let flag = if section_rtl { "<w:bidi/>" } else { "" };
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

/// §17.4.1 on the table, as it is written in a `<w:tblPr>`.
const RTL_TABLE: &str = "<w:bidiVisual/>";

// ── §17.6.6: the section's direction ────────────────────────────────────────

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

/// `tblInd` is an indent from the **leading** margin, so it measures inward
/// from the right in a `w:bidi` section.
///
/// Asserted as the gap between the table and each margin, so the claim is which
/// side the indent is on rather than what x it produces.
#[test]
fn tbl_ind_measures_from_the_sections_leading_margin() {
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

// ── §17.4.1: the table's own direction ──────────────────────────────────────

/// `w:bidiVisual` makes the table right-to-left, so its leading margin is the
/// right one — in an ordinary left-to-right section, with no `w:jc` anywhere.
///
/// This is the shape `test-files/bidi-visual-table.docx` has and the one Word
/// renders flush right.
#[test]
fn a_bidi_visual_table_sits_at_the_right_margin() {
    let (plain_x, w) = table_box(&layout(false, ""));
    let (rtl_x, rtl_w) = table_box(&layout(false, RTL_TABLE));

    assert_eq!(rtl_w, w, "the table itself does not change size");
    assert_eq!(plain_x, CONTENT_LEFT, "the control stays where it was");
    assert_eq!(
        rtl_x + w,
        CONTENT_RIGHT,
        "a right-to-left table starts at the right margin"
    );
}

/// `w:jc="right"` is `end`, and a right-to-left table's trailing edge is the
/// left one. The pair with the test above is what distinguishes "consults the
/// table's direction" from "right-aligns anything with `bidiVisual`".
#[test]
fn jc_right_on_a_bidi_visual_table_is_the_left_margin() {
    let jc = r#"<w:jc w:val="right"/>"#;
    let (ltr_x, w) = table_box(&layout(false, jc));
    let (rtl_x, _) = table_box(&layout(false, &format!("{RTL_TABLE}{jc}")));

    assert_eq!(ltr_x + w, CONTENT_RIGHT);
    assert_eq!(
        rtl_x, CONTENT_LEFT,
        "`end` is the left edge of an RTL table"
    );
}

/// The control, again: `center` has no leading edge, so `bidiVisual` must not
/// move a centred table either.
#[test]
fn a_centred_bidi_visual_table_does_not_move() {
    let jc = r#"<w:jc w:val="center"/>"#;
    let (ltr_x, _) = table_box(&layout(false, jc));
    let (rtl_x, _) = table_box(&layout(false, &format!("{RTL_TABLE}{jc}")));

    assert_eq!(rtl_x, ltr_x, "centre is centre in either direction");
}

/// §17.4.50 measures from the leading margin, and `bidiVisual` is one of the two
/// things that decides which margin that is.
#[test]
fn tbl_ind_measures_from_the_tables_own_leading_margin() {
    let ind = r#"<w:tblInd w:w="1440" w:type="dxa"/>"#; // 72pt
    let (ltr_x, w) = table_box(&layout(false, ind));
    let (rtl_x, _) = table_box(&layout(false, &format!("{RTL_TABLE}{ind}")));

    assert_eq!(ltr_x - CONTENT_LEFT, 72.0);
    assert_eq!(
        CONTENT_RIGHT - (rtl_x + w),
        72.0,
        "72pt in from the right once the table itself is RTL"
    );
}

// ── the two switches together ───────────────────────────────────────────────

/// A table that declares nothing follows its section — the fallback half of the
/// precedence rule, stated here as the direct comparison the two other tests in
/// this block are read against.
#[test]
fn a_table_that_declares_no_direction_follows_its_section() {
    let (in_ltr, w) = table_box(&layout(false, ""));
    let (in_rtl, _) = table_box(&layout(true, ""));

    assert_eq!(in_ltr, CONTENT_LEFT);
    assert_eq!(in_rtl + w, CONTENT_RIGHT);
}

/// Both switches set the same way is still that way — the case that would pass
/// under a rule that consulted only the section, and so proves nothing alone.
/// It is here because the *pair* with the test below does: one rule satisfies
/// both.
#[test]
fn a_bidi_visual_table_in_a_bidi_section_stays_at_the_right_margin() {
    let (x, w) = table_box(&layout(true, RTL_TABLE));
    assert_eq!(x + w, CONTENT_RIGHT);
}

/// …and where the two disagree, the table's own declaration wins:
/// `<w:bidiVisual w:val="0"/>` is a table saying it is *not* right-to-left, so
/// it is placed left-to-right even inside a `w:bidi` section.
///
/// The one claim in this file with no Word render behind it — see the module
/// doc. A document pairing the two elements is what would settle it, and until
/// one exists this pins the plain reading rather than leaving the case to
/// whichever branch happens to run first.
#[test]
fn an_explicitly_left_to_right_table_overrides_a_right_to_left_section() {
    let (x, _) = table_box(&layout(true, r#"<w:bidiVisual w:val="0"/>"#));
    assert_eq!(
        x, CONTENT_LEFT,
        "the table's own `w:val=\"0\"` wins over the section"
    );
}
