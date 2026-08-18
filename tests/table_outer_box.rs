//! §17.4.50 / §17.4.66: where a table's **own** left and right edges sit
//! relative to the box it reports, and what `w:tblInd` therefore measures to.
//!
//! ECMA-376 specifies no stroke geometry for table borders at all, so this is
//! not a reading of the spec but a fact about Word, and it matters well beyond
//! the border: §17.4.50 puts the table at `w:tblInd` from the margin, and the
//! candidate readings put the grid up to a whole border apart.
//!
//! `test-files/border-outer-box.docx` is the probe. It puts a paragraph — whose
//! left edge is the page margin, and which is the reference — above two tables
//! of one `w:tblInd` and one `w:tblW`, differing **only** in outer border
//! weight: 0.5pt and 12pt. A 1pt `insideV` in each is the control.
//!
//! **Measured in Word**, and it took three readings to get right. The thick
//! table's frame is drawn at 60..72 and 360..372 against a 300pt grid at the
//! 72pt margin — so the left border begins a full 12pt to the *left* of the
//! margin and the right one ends on the same edge the thin table ends on. And
//! the control moves with it: the interior line is 6pt left of where the thin
//! table puts it.
//!
//! That last fact is what names the rule, because it rules out anything that
//! only moves the outer edges:
//!
//! > a table's own edge **straddles its grid line** like every other vertical,
//! > and the whole table sits half a leading border to the left of its indent —
//! > because §17.4.50 measures to the **first cell's text edge**, and the
//! > charged half of that border puts the text back on the indent.
//!
//! Both halves are asserted below, because a change that gets one without the
//! other is still wrong: the ink and the control move, the first cell's text
//! does not. Two readings died here first and are worth recording — drawing the
//! four edges *inside* the box left the interior line alone but pushed the first
//! column's text in by the whole border, and hanging them wholly *outside* put
//! the frame in the right place with the interior line still in the wrong one.
//!
//! Every assertion is a *difference between two renders of one document*, so no
//! page margin, glyph metric or absolute coordinate is pinned: whatever those
//! contribute, they contribute equally to both sides and cancel.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

/// `w:sz` in eighths of a point, and the point width it means.
const THIN: (&str, f32) = ("4", 0.5);
const THICK: (&str, f32) = ("96", 12.0);
const EPS: f32 = 1e-3;

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

/// One two-column table at `w:tblInd` zero with **zero cell margins**, so
/// nothing but the border can inset its text, carrying `w:sz` outer borders and
/// a 1pt `insideV` as the control.
fn outer_box_table(sz: &str) -> String {
    let edges: String = ["top", "left", "bottom", "right"]
        .iter()
        .map(|e| format!(r#"<w:{e} w:val="single" w:sz="{sz}" w:space="0" w:color="C00000"/>"#))
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="6000" w:type="dxa"/>
        <w:tblInd w:w="0" w:type="dxa"/>
        <w:tblBorders>{edges}<w:insideV w:val="single" w:sz="8" w:space="0" w:color="0070C0"/></w:tblBorders>
        <w:tblCellMar>
          <w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>
          <w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>
        </w:tblCellMar>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>LEFT</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>RIGHT</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>"#
    )
}

fn layout(body: &str) -> Vec<LayoutedPage> {
    let doc = dxpdf::docx::parse(&make_docx(body)).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// `(leftmost, rightmost)` x of every rect the table paints.
fn ink_span(pages: &[LayoutedPage]) -> (f32, f32) {
    let mut lo = f32::INFINITY;
    let mut hi = f32::NEG_INFINITY;
    for c in pages.iter().flat_map(|p| &p.commands) {
        if let DrawCommand::Rect { rect, .. } = c {
            lo = lo.min(rect.origin.x.raw());
            hi = hi.max(rect.origin.x.raw() + rect.size.width.raw());
        }
    }
    (lo, hi)
}

/// The x of the text command whose content is exactly `needle`.
fn glyph_x(pages: &[LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == needle => Some(position.x.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text {needle:?} on the page"))
}

/// The x of the 1pt blue `insideV` — the control. It stands on an edge two cells
/// **share**, so it straddles its grid line under either reading and must not
/// move between the two tables.
fn inside_v(pages: &[LayoutedPage]) -> f32 {
    let mut found: Vec<f32> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color }
                if (color.r, color.g, color.b) == (0x00, 0x70, 0xC0) =>
            {
                Some(rect.origin.x.raw())
            }
            _ => None,
        })
        .collect();
    found.sort_by(f32::total_cmp);
    found.dedup_by(|a, b| (*a - *b).abs() < EPS);
    assert_eq!(found.len(), 1, "one interior grid line: {found:?}");
    found[0]
}

/// A table's own edge straddles its grid line, and the table sits half a
/// leading border left of its indent: the ink reaches the whole difference
/// further left, and **does not move at all** on the right.
///
/// The right edge staying put is the load-bearing half. Straddling on its own
/// would push it outward by half the difference; the leftward shift takes
/// exactly that back, and only a rule that does both leaves it where it is.
#[test]
fn a_tables_own_edge_straddles_its_line_and_the_table_shifts_left_by_half_of_it() {
    let thin = layout(&outer_box_table(THIN.0));
    let thick = layout(&outer_box_table(THICK.0));
    let grew = THICK.1 - THIN.1;

    let (thin_lo, thin_hi) = ink_span(&thin);
    let (thick_lo, thick_hi) = ink_span(&thick);

    assert!(
        (thin_lo - thick_lo - grew).abs() < EPS,
        "the left edge should reach {grew}pt further out, not {}: {thin_lo} -> {thick_lo}",
        thin_lo - thick_lo
    );
    assert!(
        (thick_hi - thin_hi).abs() < EPS,
        "the right edge should not move at all: {thin_hi} -> {thick_hi}"
    );

    // And the grid went with it — the control is on a shared edge, so it
    // straddles under every reading and can only move because the table did.
    assert!(
        (inside_v(&thin) - inside_v(&thick) - grew * 0.5).abs() < EPS,
        "the interior line should be {}pt left of the thin table's, not {}: {} -> {}",
        grew * 0.5,
        inside_v(&thin) - inside_v(&thick),
        inside_v(&thin),
        inside_v(&thick)
    );
}

/// …and §17.4.50 measures to the **first cell's text edge**, so that text does
/// not move: the leftward shift and the half-border charged to the cell cancel.
///
/// The second cell is the counterpart and is asserted too — it is not against
/// the table's own edge, so it simply travels with the grid. Together the two
/// say the shift is a shift and not a change of column widths.
#[test]
fn the_indent_measures_to_the_first_cells_text_edge() {
    let thin = layout(&outer_box_table(THIN.0));
    let thick = layout(&outer_box_table(THICK.0));
    let grew = THICK.1 - THIN.1;

    assert_eq!(
        glyph_x(&thin, "LEFT"),
        glyph_x(&thick, "LEFT"),
        "the first cell's text is what the indent measures to, and must not move"
    );
    assert!(
        (glyph_x(&thin, "RIGHT") - glyph_x(&thick, "RIGHT") - grew * 0.5).abs() < EPS,
        "the second cell travels with the grid, {}pt left: {} -> {}",
        grew * 0.5,
        glyph_x(&thin, "RIGHT"),
        glyph_x(&thick, "RIGHT")
    );
}
