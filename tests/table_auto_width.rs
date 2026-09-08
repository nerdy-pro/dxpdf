//! §17.4.63 / §17.4.52 — how wide an auto-width table is allowed to get.
//!
//! A `<w:tblW w:type="auto"/>` table states no preferred width, so this engine
//! keeps its `<w:tblGrid>` verbatim: Word's autofit algorithm is not
//! implemented (see `table::grid::compute_column_widths`), and scaling the
//! declared grid to the text column would be *implementing* one, badly.
//!
//! Verbatim has a floor, though. A grid summing to 20000 twips on a Letter page
//! drew the table from x=72 to x=1072 — 460 pt past the right edge of the paper
//! — and every glyph and border out there is simply gone from the PDF. This
//! file pins the guard, and pins how narrow it deliberately is.
//!
//! # Why the guard is drawn at the paper and not at the text column
//!
//! Clamping to the content area is the obvious fix and is refuted by the
//! corpus. Three real documents ship an auto table whose grid exceeds the text
//! column — `sample-docx-files-sample3.docx` by 101 twips, `ELH_2025-12-18` and
//! `KAB_2026-03-25` by ~172 twips across 39 tables — and all 39 of the latter
//! declare `<w:tblLayout w:type="fixed"/>`, which is §17.4.52's instruction to
//! use the declared widths rather than compute any. Word draws those tables a
//! few points into the right margin, on the paper and fully visible; nothing is
//! lost and nothing needs guarding. Normalising them to the text column would
//! move 40 tables of real Word output to a width no reading of §17.4.52 asks
//! for.
//!
//! So the guard is drawn at the one line every reading agrees on: content that
//! leaves the paper is lost, whatever the layout algorithm was supposed to be.
//! What the *right* width for an overflowing auto table is remains open, and a
//! **Word render** of one — at `tblLayout` both `fixed` and `autofit` — is what
//! would settle it.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn make_docx(document_xml: &str) -> Vec<u8> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .unwrap();

    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

/// A one-row table with the given `<w:gridCol>`s and `w:tblW`, bordered so its
/// extent is drawn rather than inferred. No `<w:sectPr>`, so the page is the
/// §17.6.13 default Letter with 1-inch margins: 612 pt wide, text column
/// 72…540, paper edge 612.
fn table(tbl_w: &str, grid: &str) -> Vec<u8> {
    let cells: String = grid
        .matches("<w:gridCol")
        .enumerate()
        .map(|(i, _)| format!(r#"<w:tc><w:p><w:r><w:t>c{i}</w:t></w:r></w:p></w:tc>"#))
        .collect();
    make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        {tbl_w}
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>{grid}</w:tblGrid>
      <w:tr>{cells}</w:tr>
    </w:tbl>
  </w:body>
</w:document>"#
    ))
}

fn layout(bytes: &[u8]) -> Vec<LayoutedPage> {
    let parsed = dxpdf::docx::parse(bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// `w:sz="4"` in [`table`], as points: the width of every border it declares.
///
/// §17.4.66: an outer border **straddles** its grid line and the table sits half
/// its leading border to the left of the indent, so that `w:tblInd` still
/// measures to the first cell's text edge (`borders::rasterize_border_grid` and
/// `build_table`, both measured against `test-files/border-outer-box.docx` —
/// Word draws that fixture's 12pt right border at 360–372 against a declared
/// right edge of 372).
///
/// The two halves cancel on the trailing edge: the box ends half a border short
/// of the grid and the border reaches half a border past the box, so **the ink
/// ends exactly on the declared grid line**, which is why no expectation below
/// adds anything to it. They do not cancel on the leading edge, where the shift
/// and the straddle both go left and the ink starts a whole border outside.
///
/// Kept as a named constant because the paper guard below is a *bound* rather
/// than an equality, and that bound is still in these units.
const OUTER: f32 = 0.5;

/// The rightmost x any border rect reaches — the table's drawn right edge.
fn drawn_right_edge(pages: &[LayoutedPage]) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => Some((rect.origin.x + rect.size.width).raw()),
            _ => None,
        })
        .fold(f32::NEG_INFINITY, f32::max)
}

/// The reported defect: `<w:gridCol w:w="20000"/>` (1000 pt) under
/// `<w:tblW w:type="auto"/>` drew the table to x=1072 on a 612 pt page, so
/// 460 pt of it — borders, cell, text — fell off the paper and out of the PDF.
///
/// Nothing in §17.4.63 or §17.4.52 licenses drawing off the page; the width to
/// use instead is genuinely open (see this file's header), and only the
/// containment is asserted here.
#[test]
fn an_auto_table_wider_than_the_page_is_kept_on_the_page() {
    let pages = layout(&table(
        r#"<w:tblW w:w="0" w:type="auto"/>"#,
        r#"<w:gridCol w:w="20000"/>"#,
    ));
    let page_width = pages[0].page_size.width.raw();

    // The **grid** is what the guard clamps, and it is clamped to the paper
    // edge exactly; the outer border may hang up to `OUTER` past it, which is the
    // third of the deliberate imprecisions `clamp_auto_grid_to_page` names —
    // the guard is given the grid and knows nothing of the borders that will be
    // drawn around it. Bounding the overflow rather than forbidding it keeps
    // this a regression test for the reported defect, which drew 460 pt of
    // table off the sheet.
    assert!(
        drawn_right_edge(&pages) <= page_width + OUTER + 0.01,
        "table drawn to x={:.1} on a {page_width:.0} pt page — {:.1} pt of it is \
         off the paper, more than the {OUTER} pt its own border may hang",
        drawn_right_edge(&pages),
        drawn_right_edge(&pages) - page_width,
    );
}

/// The deliberate narrowness, and the assertion that a clamp to the *content
/// area* would fail: a grid of 10000 twips (500 pt) overflows the 468 pt text
/// column by 32 pt but ends at x=572, well inside the 612 pt paper. It is left
/// exactly as declared.
///
/// This is the shape 40 tables of real Word/LibreOffice output in the corpus
/// have — `ELH_2025-12-18` and `KAB_2026-03-25` overflow the text column by
/// ~172 twips under an explicit `<w:tblLayout w:type="fixed"/>`.
#[test]
fn an_auto_table_that_overflows_the_margin_but_not_the_paper_is_untouched() {
    let pages = layout(&table(
        r#"<w:tblW w:w="0" w:type="auto"/>"#,
        r#"<w:gridCol w:w="10000"/>"#,
    ));

    // 10000 twips = 500 pt, drawn from the 72 pt left margin. The ink ends on
    // the grid line itself — see `OUTER` for why nothing is added here.
    assert!(
        (drawn_right_edge(&pages) - 572.0).abs() < 0.01,
        "declared grid was rescaled: right edge {:.2}, expected {:.2}",
        drawn_right_edge(&pages),
        572.0,
    );
}

/// And an auto table that fits keeps its grid to the point — the guard must be
/// invisible to every document that never trips it, which is all but three of
/// the 52 in the corpus.
#[test]
fn an_auto_table_that_fits_keeps_its_declared_grid() {
    let pages = layout(&table(
        r#"<w:tblW w:w="0" w:type="auto"/>"#,
        r#"<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>"#,
    ));

    // 6000 twips = 300 pt from x=72, and the ink ends on that grid line.
    assert!(
        (drawn_right_edge(&pages) - 372.0).abs() < 0.01,
        "a fitting grid was rescaled: right edge {:.2}, expected {:.2}",
        drawn_right_edge(&pages),
        372.0,
    );
}

/// §17.4.63: a table with a *declared* width is a different question, and this
/// guard must not answer it. `w:type="dxa"` states a preferred width, which
/// §17.4.63 says the columns are scaled to; an over-wide `dxa` table is the
/// author asking for exactly that width, not a producer forgetting to autofit.
/// Pinned so the auto-only scope of the guard is a decision with a test behind
/// it rather than an accident of where the branch was written.
#[test]
fn a_declared_width_wider_than_the_page_is_left_to_its_own_scaling() {
    let pages = layout(&table(
        r#"<w:tblW w:w="20000" w:type="dxa"/>"#,
        r#"<w:gridCol w:w="20000"/>"#,
    ));

    // 20000 twips = 1000 pt: §17.4.63 scaling takes the grid to the declared
    // width, which is what runs off the page. Unchanged by this guard.
    assert!(
        drawn_right_edge(&pages) > pages[0].page_size.width.raw(),
        "the dxa path was clamped too: right edge {:.2}",
        drawn_right_edge(&pages),
    );
}
