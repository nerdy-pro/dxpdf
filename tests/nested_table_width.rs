//! §17.4.63 / §17.4.42 — a table nested in a cell is measured against that
//! cell, not against the sheet (issue #229).
//!
//! `clamp_auto_grid_to_page` draws its limit at the paper edge and takes the
//! width the caller offered only as a *floor* under that limit. For a top-level
//! table the caller's width is the text column, and the deliberate slack
//! between the two is what lets the 40 corpus tables that reach a few points
//! into the right margin keep their declared grid — `tests/table_auto_width.rs`
//! pins that and must stay green.
//!
//! For a *nested* table the caller's width is the host cell's content width,
//! and there the slack is not slack: the cell is the container, and a grid
//! wider than it is drawn straight through the cell's border and over whatever
//! the neighbouring cell holds. The corpus fixture `019_nested_tables.docx`
//! shows it — Word keeps the nested table inside "Outer A"; dxpdf drew its rows
//! 12 pt past the border and into "Outer B"'s text.
//!
//! The pages here declare no `w:tblCellMar`, so cell margins are zero and the
//! arithmetic is exact: Letter with one-inch margins gives a 468 pt text
//! column at x = 72, the outer grid splits it 144 / 324, and the host cell's
//! content box is therefore x = 72…216.

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

/// The host cell is x = 72…216; its neighbour runs to the right margin at 540.
const CELL_LEFT: f32 = 72.0;
const CELL_RIGHT: f32 = 216.0;

/// An outer two-column table, 2880 + 6480 twips, whose first cell holds
/// `nested` and whose second holds the word `OUTER_B`.
fn outer_with(nested: &str) -> Vec<LayoutedPage> {
    let bytes = make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="9360" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="6480"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>{nested}</w:tc>
        <w:tc><w:tcPr><w:tcW w:w="6480" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>OUTER_B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#
    ));
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// A nested table of two equal columns, each `half` twips wide, holding `N11`
/// and `N12`. `tbl_w` says how it states its width.
fn nested(tbl_w: &str, half: u32, extra_pr: &str) -> String {
    format!(
        r#"<w:tbl>
      <w:tblPr>{tbl_w}{extra_pr}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="{half}"/><w:gridCol w:w="{half}"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>N11</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>N12</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>"#
    )
}

const AUTO: &str = r#"<w:tblW w:w="0" w:type="auto"/>"#;

/// The x of the first `Text` command whose text is `needle`.
fn text_x(pages: &[LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { position, text, .. } if &**text == needle => Some(position.x.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no Text command says {needle:?}"))
}

/// The reported defect: 7200 twips (360 pt) of declared grid inside a 144 pt
/// cell. Word scales it to the cell; dxpdf drew it at 360 pt, so the second
/// nested column opened at x = 252 — 36 pt into the neighbouring cell.
#[test]
fn a_nested_auto_table_is_scaled_to_its_host_cell() {
    let pages = outer_with(&nested(AUTO, 3600, ""));

    let n11 = text_x(&pages, "N11");
    let n12 = text_x(&pages, "N12");

    assert!(
        (n11 - CELL_LEFT).abs() < 0.01,
        "the nested table starts at the cell's content edge: {n11:.2} vs {CELL_LEFT:.2}",
    );
    assert!(
        n12 < CELL_RIGHT,
        "the second nested column opened at x={n12:.2}, past the host cell's \
         right edge at {CELL_RIGHT:.2} — the nested table is drawn over its neighbour",
    );
    // Scaled proportionally, so two equal declared columns stay equal: each is
    // half of the cell's 144 pt content box.
    assert!(
        (n12 - n11 - 72.0).abs() < 0.01,
        "columns are not half the cell: pitch {:.2}, expected 72.00",
        n12 - n11,
    );
}

/// Clamp down only. A nested grid that already fits keeps its declared widths —
/// Word does not stretch a table to its container, and neither does this.
#[test]
fn a_nested_auto_table_that_fits_keeps_its_declared_grid() {
    let pages = outer_with(&nested(AUTO, 1080, "")); // 2 × 54 pt = 108 pt in a 144 pt cell

    let n11 = text_x(&pages, "N11");
    let n12 = text_x(&pages, "N12");

    assert!(
        (n12 - n11 - 54.0).abs() < 0.01,
        "a nested table narrower than its cell was resized: pitch {:.2}, expected 54.00",
        n12 - n11,
    );
}

/// The ceiling is attached at the nested call site, so a top-level auto table
/// is untouched: it still keeps a grid that overflows the text column, which is
/// the shape 40 corpus tables have. `tests/table_auto_width.rs` owns this
/// property; the control here is what makes the two tests above readable as a
/// *nested*-only change.
#[test]
fn a_top_level_auto_table_still_keeps_a_grid_past_the_text_column() {
    let bytes = make_docx(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>T11</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>T12</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#,
    );
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    let pages = dxpdf::render::resolve_and_layout(parsed).1;

    // 5000 twips = 250 pt per column: 500 pt of table in a 468 pt text column,
    // still on the 612 pt sheet, so the guard leaves it alone.
    assert!(
        (text_x(&pages, "T12") - text_x(&pages, "T11") - 250.0).abs() < 0.01,
        "a top-level auto table must keep its declared grid",
    );
}

/// §17.4.57: a nested table that also declares `w:tblpPr` floats, and whether
/// Word confines a *floating* nested table to its host cell is not measured.
/// Until it is, such a table keeps the behaviour it has today — the page guard
/// — rather than being silently moved by this change. Flipping it is one line
/// once the Word render exists.
#[test]
fn a_nested_floating_table_still_measures_against_the_page() {
    let float = r#"<w:tblpPr w:vertAnchor="text" w:horzAnchor="text" w:tblpX="1" w:tblpY="1"/>"#;
    let pages = outer_with(&nested(AUTO, 3600, float));

    let pitch = text_x(&pages, "N12") - text_x(&pages, "N11");
    assert!(
        (pitch - 180.0).abs() < 0.01,
        "a floating nested table is not clamped to its cell yet: pitch {pitch:.2}, \
         expected the declared 180.00",
    );
}
