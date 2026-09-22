//! §17.4.52 / §17.4.71 — an autofit table resolves its columns from the cells'
//! `w:tcW` preferences, not from the saved `w:tblGrid` (issue #231).
//!
//! §17.4.63 and §17.4.71 carry the same paragraph verbatim: every width in a
//! table is "preferred", the table "shall satisfy the shared columns as
//! specified by the tblGrid", and "the table layout algorithm can require a
//! preference to be overridden". The spec states the conflict and leaves it
//! there. This engine used to make the grid the invariant; Word resolves it the
//! other way whenever §17.4.52's layout is autofit, which is the default when
//! `<w:tblLayout>` is absent.
//!
//! Measured on Word 16.0 (Office 2024 LTSC), on a page whose text column is
//! exactly 504 pt, with a grid of 2000/4000/4080 twips in every case:
//!
//! | `w:tcW` (twips) | sum | Word's columns (pt) |
//! |---|---|---|
//! | 1000 / 3000 / 6080 | = text column | 50 / 150 / 304 — the preferences, grid discarded |
//! | 1000 / 2000 / 3000 | < text column | 50 / 100 / 150 — no stretch to the column |
//! | 4000 / 4000 / 4000 | > text column | 168.2 / 168.15 / 168.15 — scaled down to it |
//!
//! The first row is what settles it: the grid, the preferences and an equal
//! division all predict different numbers there, and Word draws the
//! preferences.
//!
//! # What is deliberately left out
//!
//! Word also raises a column to its *content* minimum — the width of the
//! widest thing in it that cannot be broken — and takes the difference from the
//! others. The same probe measured it: `tcW` 500/3000/6580 with a long word in
//! column 2 came back as 39.65 / 149.25 / 315.6, where column 1's 25 pt
//! preference lost to its content. That needs a min-content measurement per
//! cell, which this change does not add; a table whose preferences are narrower
//! than its content therefore still draws them as asked.

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

/// The §17.6.13 default page: Letter, one-inch margins, text column 72…540.
const TEXT_LEFT: f32 = 72.0;
const TEXT_COLUMN: f32 = 468.0;

/// A three-column table: `w:tblGrid` always 2000/4000/4080 twips, each cell
/// declaring the matching `tcw`, one word per cell so the column starts can be
/// read off the drawn text.
fn table(tbl_pr_extra: &str, tcw: [u32; 3]) -> Vec<LayoutedPage> {
    let cells: String = (0..3)
        .map(|i| {
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{}" w:type="dxa"/></w:tcPr>
                     <w:p><w:r><w:t>c{i}</w:t></w:r></w:p></w:tc>"#,
                tcw[i]
            )
        })
        .collect();
    let bytes = make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/>{tbl_pr_extra}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/><w:gridCol w:w="4080"/></w:tblGrid>
      <w:tr>{cells}</w:tr>
    </w:tbl>
  </w:body>
</w:document>"#
    ));
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// Where each cell's word was drawn, which with zero cell margins is where its
/// column starts.
fn column_starts(pages: &[LayoutedPage]) -> Vec<f32> {
    ["c0", "c1", "c2"]
        .iter()
        .map(|needle| {
            pages
                .iter()
                .flat_map(|p| &p.commands)
                .find_map(|c| match c {
                    DrawCommand::Text { position, text, .. } if &**text == *needle => {
                        Some(position.x.raw())
                    }
                    _ => None,
                })
                .unwrap_or_else(|| panic!("no Text command says {needle:?}"))
        })
        .collect()
}

fn widths(pages: &[LayoutedPage]) -> Vec<f32> {
    let starts = column_starts(pages);
    vec![starts[1] - starts[0], starts[2] - starts[1]]
}

fn assert_close(got: &[f32], want: [f32; 2], what: &str) {
    assert!(
        (got[0] - want[0]).abs() < 0.05 && (got[1] - want[1]).abs() < 0.05,
        "{what}: columns 1 and 2 came out {:.2} / {:.2}, expected {:.2} / {:.2}",
        got[0],
        got[1],
        want[0],
        want[1],
    );
}

/// The discriminating case. The grid says 100/200/204 pt, the preferences say
/// 50/150/268, and an equal division would say 156 each. Word draws the
/// preferences, and so does this.
#[test]
fn an_autofit_table_takes_its_columns_from_tcw() {
    // 1000 + 3000 + 5360 = 9360 twips = 468 pt, the text column exactly.
    let pages = table("", [1000, 3000, 5360]);

    assert!(
        (column_starts(&pages)[0] - TEXT_LEFT).abs() < 0.05,
        "the table still starts at the text margin",
    );
    assert_close(&widths(&pages), [50.0, 150.0], "preferences honoured");
}

/// Preferences that do not fill the text column are drawn as asked: Word does
/// not stretch an autofit table to its container.
#[test]
fn preferences_narrower_than_the_text_column_are_not_stretched() {
    let pages = table("", [1000, 2000, 3000]); // 300 pt of table in a 468 pt column

    assert_close(&widths(&pages), [50.0, 100.0], "no stretch");
}

/// Preferences that overflow are scaled down proportionally until they fit —
/// 600 pt of equal preferences becomes three equal columns of the 468 pt text
/// column, as Word's 12000-twip probe did against its own 504 pt column.
#[test]
fn preferences_wider_than_the_text_column_are_scaled_down() {
    let pages = table("", [4000, 4000, 4000]);
    let third = TEXT_COLUMN / 3.0;

    assert_close(&widths(&pages), [third, third], "scaled to the text column");
}

/// §17.4.52: `<w:tblLayout w:type="fixed"/>` is the instruction *not* to
/// autofit, so the declared grid stands whatever the cells prefer. This is the
/// shape 40 corpus tables have — `ELH_2025-12-18` and `KAB_2026-03-25` declare
/// it on all 39 of theirs — and `tests/table_auto_width.rs` pins that they keep
/// a grid reaching past the text column.
#[test]
fn a_fixed_layout_table_keeps_its_declared_grid() {
    let pages = table(r#"<w:tblLayout w:type="fixed"/>"#, [1000, 3000, 5360]);

    // 2000 and 4000 twips: the grid, not the 1000/3000 preferences.
    assert_close(&widths(&pages), [100.0, 200.0], "grid kept");
}

/// A table that declares no `w:tcW` at all has no preferences to resolve, so it
/// keeps the declared grid — the path every fixture in
/// `tests/table_auto_width.rs` takes.
#[test]
fn a_table_without_preferences_keeps_its_declared_grid() {
    let bytes = make_docx(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/><w:gridCol w:w="4080"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>c0</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>c1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#,
    );
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    let pages = dxpdf::render::resolve_and_layout(parsed).1;

    assert_close(&widths(&pages), [100.0, 200.0], "grid kept");
}
