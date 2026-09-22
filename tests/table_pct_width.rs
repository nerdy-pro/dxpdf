//! §17.4.63 — a `pct` table width is measured against the text extents, and the
//! cell-margin extension belongs to full width alone (issue #231).
//!
//! Word autofits a table that fills the window by widening it by its cell
//! margins and pulling it left by one, so the cell *text* — not the cell box —
//! lines up with body text. `tests/table_geometry_sizing.rs` pins that at
//! `w:w="5000"`, and it stays pinned here.
//!
//! The gate was written `>=`, so the same extension fired for a table that asks
//! for *more* than the window. Word does not extend those: for a 104% table it
//! saves a `<w:tblGrid>` summing to 1.04 × the text extents, with no margin
//! term — which is all §17.4.63 describes ("relative to the text extents of the
//! page"). The 2.3% of extra width the extension added was enough to stop a
//! header cell wrapping where Word wraps it, and pushed the table's right edge
//! outside the right margin.

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

/// The §17.6.13 default page: Letter, one-inch margins, so the text extents are
/// 468 pt wide at x = 72.
const TEXT_EXTENTS: f32 = 468.0;
/// `w:tblCellMar` below declares 108 twips a side: 5.4 pt each, 10.8 pt total.
const CELL_MARGINS: f32 = 10.8;
/// §17.4.66 draws an outer border **straddling** its grid line, so the ink
/// reaches half a border past the grid at each edge and the drawn span is one
/// whole border wider than the declared width. `w:sz="4"` below is a 0.5 pt
/// border. `tests/table_auto_width.rs` measures the same straddle against
/// `test-files/border-outer-box.docx` and Word's own render of it.
const OUTER: f32 = 0.5;

/// A two-column table at the given `w:tblW`, with explicit 108-twip cell
/// margins so the extension has something to add.
fn pct_table(pct: u32) -> Vec<LayoutedPage> {
    let bytes = make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="{pct}" w:type="pct"/>
        <w:tblCellMar>
          <w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#
    ));
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// The table's drawn extent, from the leftmost to the rightmost border rect.
fn drawn_span(pages: &[LayoutedPage]) -> (f32, f32) {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => {
                Some((rect.origin.x.raw(), (rect.origin.x + rect.size.width).raw()))
            }
            _ => None,
        })
        .fold((f32::INFINITY, f32::NEG_INFINITY), |(lo, hi), (l, r)| {
            (lo.min(l), hi.max(r))
        })
}

/// The defect: at 104% the base was the text extents *plus* the cell margins,
/// so the table came out 2.3% wider than Word draws it.
#[test]
fn a_table_past_full_width_is_not_extended_by_its_cell_margins() {
    let (left, right) = drawn_span(&pct_table(5200));
    let width = right - left;
    let expected = TEXT_EXTENTS * 1.04 + OUTER;

    assert!(
        (width - expected).abs() < 0.05,
        "104% of the text extents is {expected:.2} pt; drawn {width:.2} pt \
         ({:.2} pt of cell margin was added to the base)",
        width - expected,
    );
}

/// The extension itself, unchanged: at exactly 100% the table is widened by its
/// cell margins and pulled left by one, so the first cell's text still sits on
/// the left margin. `tests/table_geometry_sizing.rs` owns this property; the
/// control here is what makes the gate above readable as an *above*-100% change.
#[test]
fn a_full_width_table_still_extends_by_its_cell_margins() {
    let (left, right) = drawn_span(&pct_table(5000));
    let width = right - left;

    assert!(
        (width - (TEXT_EXTENTS + CELL_MARGINS + OUTER)).abs() < 0.05,
        "a full-width table is the text extents plus its cell margins \
         ({:.2} pt); drawn {width:.2} pt",
        TEXT_EXTENTS + CELL_MARGINS,
    );
    assert!(
        (left - (72.0 - CELL_MARGINS / 2.0 - OUTER / 2.0)).abs() < 0.05,
        "and it starts half its margins left of the text edge; drawn at {left:.2}",
    );
}

/// Below full width nothing was ever added, and nothing is now: 50% is half the
/// text extents exactly.
#[test]
fn a_half_width_table_is_half_the_text_extents() {
    let (left, right) = drawn_span(&pct_table(2500));
    let width = right - left;

    assert!(
        (width - (TEXT_EXTENTS / 2.0 + OUTER)).abs() < 0.05,
        "50% of the text extents is {:.2} pt; drawn {width:.2} pt",
        TEXT_EXTENTS / 2.0,
    );
}
