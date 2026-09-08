//! §17.4.80 / §17.4.84 / §17.4.6 — table row height, page-advance, and the two malformed
//! `w:vMerge` shapes.
//!
//! These defects are invisible at the `layout_table` level and only show up
//! once the section layer places the result: a zero-height row lets the *next*
//! block draw over the table, an abandoned empty leading slice becomes a blank
//! page, and an unpaired `w:vMerge` drops a cell's text out of the document
//! altogether. The unit tests in `table/mod.rs` and `build/table.rs` pin the
//! arithmetic; these pin what a reader of the PDF would actually see.

use std::io::Write;

use dxpdf::render::layout::draw_command::DrawCommand;

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

fn doc(body: &str) -> Vec<u8> {
    make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}</w:body>
</w:document>"#
    ))
}

fn layout(bytes: &[u8]) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    let parsed = dxpdf::docx::parse(bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// y of the first text command whose content equals `needle`.
fn y_of(pages: &[dxpdf::render::layout::draw_command::LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == needle => Some(position.y.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text command {needle:?}"))
}

/// A single-cell table whose only cell carries `vmerge` verbatim in `tcPr`.
fn lone_cell_table(vmerge: &str) -> Vec<u8> {
    doc(&format!(
        r#"<w:p><w:r><w:t>before</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/>{vmerge}</w:tcPr>
        <w:p><w:r><w:t>AAA</w:t></w:r></w:p>
        <w:p><w:r><w:t>BBB</w:t></w:r></w:p>
        <w:p><w:r><w:t>CCC</w:t></w:r></w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>after</w:t></w:r></w:p>"#
    ))
}

/// §17.4.84: a `vMerge="restart"` with no continuation is an ordinary cell.
///
/// The row used to collapse to zero height while still drawing its content,
/// so `after` was emitted *above* the table's own text instead of below it.
#[test]
fn lone_vmerge_restart_does_not_let_following_content_overlap_the_table() {
    let pages = layout(&lone_cell_table(r#"<w:vMerge w:val="restart"/>"#));

    let last_row_text = y_of(&pages, "CCC");
    let after = y_of(&pages, "after");

    assert!(
        after > last_row_text,
        "\"after\" must follow the table's last line (y={last_row_text:.1}), \
         but was drawn at y={after:.1} — on top of it"
    );
}

/// Calibrated against the unmerged control: a lone restart must lay out
/// identically to no merge at all, not merely "somewhere below".
#[test]
fn lone_vmerge_restart_matches_the_unmerged_layout() {
    let restart = layout(&lone_cell_table(r#"<w:vMerge w:val="restart"/>"#));
    let control = layout(&lone_cell_table(""));

    assert!(
        (y_of(&restart, "after") - y_of(&control, "after")).abs() < 0.01,
        "lone restart moved the following paragraph: {:.1} vs control {:.1}",
        y_of(&restart, "after"),
        y_of(&control, "after"),
    );
}

/// Every string this document draws, in page order.
fn all_text(pages: &[dxpdf::render::layout::draw_command::LayoutedPage]) -> Vec<String> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, .. } => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

/// A two-row, two-column table whose **first** row's left cell carries
/// `vmerge` verbatim. `w:val="continue"` there has no `restart` above it —
/// there is no row above it at all — which is the orphan case.
fn orphan_continue_table(vmerge: &str) -> Vec<u8> {
    doc(&format!(
        r#"<w:p><w:r><w:t>before</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>{vmerge}</w:tcPr>
          <w:p><w:r><w:t>ORPHAN</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>NEXT</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>NEXT2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>after</w:t></w:r></w:p>"#
    ))
}

/// §17.4.84: a `<w:vMerge w:val="continue"/>` with **no `restart` above it**
/// must not cost the cell its content.
///
/// §17.4.84 describes `continue` only as continuing the merge begun by a
/// `restart`, and says nothing about a `continue` that begins one — so the
/// engine is free to choose what an unpaired one means. It is *not* free to
/// choose that the text disappears: no reading of §17.4.84 deletes content,
/// and a merged cell shows the *restart* cell's content precisely because
/// there is one to show.
///
/// See `build::table::promote_orphan_vmerge_continues` for the choice made and
/// the corroborating LibreOffice behaviour.
#[test]
fn orphan_vmerge_continue_keeps_its_cell_content() {
    let pages = layout(&orphan_continue_table(r#"<w:vMerge w:val="continue"/>"#));
    let drawn = all_text(&pages);

    assert!(
        drawn.iter().any(|t| t == "ORPHAN"),
        "the orphaned cell's text was dropped from the document — drawn: {drawn:?}"
    );
}

/// …and the cell is an ordinary one, not merely a visible one: it lays out
/// exactly as the same table with no `w:vMerge` at all, which is what
/// "promote it to an ordinary cell" has to mean if it is to mean anything
/// measurable. Calibrated against that control rather than against a literal
/// y, so no glyph metric has to be known.
#[test]
fn orphan_vmerge_continue_matches_the_unmerged_layout() {
    let orphan = layout(&orphan_continue_table(r#"<w:vMerge w:val="continue"/>"#));
    let control = layout(&orphan_continue_table(""));

    for needle in ["ORPHAN", "PEER", "NEXT", "NEXT2", "after"] {
        assert!(
            (y_of(&orphan, needle) - y_of(&control, needle)).abs() < 0.01,
            "{needle} moved: y={:.2} vs control y={:.2}",
            y_of(&orphan, needle),
            y_of(&control, needle),
        );
    }
}

/// The control that must not move: a `continue` that *does* have a `restart`
/// above it is a real merge, and §17.4.84 makes the merged region show the
/// restart cell's content — so the continue cell's own content stays hidden.
/// A fix that simply stopped honouring `continue` would pass both tests above
/// and fail this one.
#[test]
fn a_paired_vmerge_continue_still_hides_its_own_content() {
    let pages = layout(&doc(r#"<w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>MERGED</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="continue"/></w:tcPr>
          <w:p><w:r><w:t>HIDDEN</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"#));
    let drawn = all_text(&pages);

    assert!(drawn.iter().any(|t| t == "MERGED"), "drawn: {drawn:?}");
    assert!(
        !drawn.iter().any(|t| t == "HIDDEN"),
        "a continued merge shows the restart cell's content, not the \
         continue cell's — drawn: {drawn:?}"
    );
}

// ── §17.4.84 × §17.4.17: the orphan rule where `gridSpan` makes a cell wider
//    than one grid column ────────────────────────────────────────────────────
//
// Every orphan case above uses span-1 cells, where a cell *is* a grid column
// and every candidate pairing rule agrees. The two below are the cases that
// separate them, and both were silently dropping content: the promotion pass
// paired by "any column the cell covers" while `expand_rows_for_vmerge` and
// `merged_span_height` anchor a merge on the restart's **first** grid column
// alone (`is_vmerge_continue(row_below, entry.grid_col)`). A `Continue` the
// promotion pass called paired and those passes did not merge is accounted
// for by nobody — `measure_table_rows` still hands it an empty `CellLayout`,
// so its text never reaches the page.
//
// Written with the **bare** `<w:vMerge/>` spelling (§17.4.84: an absent `@val`
// means `continue`), which is what real documents contain: all 98 `continue`
// elements across the 52-document corpus are bare, and not one is written
// `w:val="continue"`.

/// Row 0 is a single `gridSpan="2"` `restart`; row 1 is a plain cell at grid
/// column 0 beside a `continue` at grid column 1.
fn wide_restart_narrow_continue_table(vmerge: &str) -> Vec<u8> {
    doc(&format!(
        r#"<w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/><w:gridSpan w:val="2"/>
              <w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>WIDE</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PLAIN</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>{vmerge}</w:tcPr>
          <w:p><w:r><w:t>ORPHAN</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"#
    ))
}

/// The `continue` at column 1 is an orphan: `expand_rows_for_vmerge` anchors
/// the wide restart on column 0, finds `PLAIN` there, and never merges — so
/// nothing sizes or draws the `continue` cell. Its text must survive, exactly
/// as in the span-1 case; the width of the restart above it is not a reason to
/// delete a cell's content.
#[test]
fn a_continue_outside_a_wide_restarts_first_column_keeps_its_cell_content() {
    let pages = layout(&wide_restart_narrow_continue_table("<w:vMerge/>"));
    let drawn = all_text(&pages);

    assert!(
        drawn.iter().any(|t| t == "ORPHAN"),
        "a `continue` that the merge passes never join had its text dropped \
         from the document — drawn: {drawn:?}"
    );
}

/// …and it is an ordinary cell, not merely a visible one — the same calibration
/// the span-1 orphan uses: identical to the table with no `w:vMerge` on that
/// cell at all.
#[test]
fn a_continue_outside_a_wide_restarts_first_column_matches_the_unmerged_layout() {
    let orphan = layout(&wide_restart_narrow_continue_table("<w:vMerge/>"));
    let control = layout(&wide_restart_narrow_continue_table(""));

    for needle in ["WIDE", "PLAIN", "ORPHAN"] {
        assert!(
            (y_of(&orphan, needle) - y_of(&control, needle)).abs() < 0.01,
            "{needle} moved: y={:.2} vs control y={:.2}",
            y_of(&orphan, needle),
            y_of(&control, needle),
        );
    }
}

/// The control on the other side, and the one a blanket "promote everything
/// under a wide restart" would break: a `continue` under the restart's **first**
/// column *is* the merge those passes join, so it keeps hiding its own content.
#[test]
fn a_continue_under_a_wide_restarts_first_column_is_still_a_merge() {
    let pages = layout(&doc(r#"<w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/><w:gridSpan w:val="2"/>
              <w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>WIDE</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge/></w:tcPr>
          <w:p><w:r><w:t>HIDDEN</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"#));
    let drawn = all_text(&pages);

    assert!(drawn.iter().any(|t| t == "WIDE"), "drawn: {drawn:?}");
    assert!(
        !drawn.iter().any(|t| t == "HIDDEN"),
        "the `continue` sits under the restart's own grid column, so it \
         continues a real merge and shows the restart's content — drawn: \
         {drawn:?}"
    );
}

/// Row 1 is a `gridSpan="2"` `continue` over a span-1 `restart` at column 0.
/// It is paired — column 0 is open — but it must not thereby *open* column 1,
/// which no `restart` ever began.
fn wide_continue_over_narrow_restart_table(vmerge: &str) -> Vec<u8> {
    doc(&format!(
        r#"<w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>RESTART</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/><w:gridSpan w:val="2"/>
              <w:vMerge/></w:tcPr>
          <w:p><w:r><w:t>WIDECONT</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PLAIN</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>{vmerge}</w:tcPr>
          <w:p><w:r><w:t>ORPHAN</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"#
    ))
}

/// Column 1 was never opened by a `restart`, so the `continue` in the last row
/// is an orphan by every definition — including the one the merge passes use,
/// which only ever track the restart's column 0. Widening the open set to the
/// whole span of a paired `continue` made column 1 read as open and deleted
/// this cell's text.
#[test]
fn a_continue_below_a_wider_continue_keeps_its_cell_content() {
    let pages = layout(&wide_continue_over_narrow_restart_table("<w:vMerge/>"));
    let drawn = all_text(&pages);

    assert!(
        drawn.iter().any(|t| t == "ORPHAN"),
        "a `continue` in a column no `restart` ever opened had its text \
         dropped from the document — drawn: {drawn:?}"
    );
    assert!(
        !drawn.iter().any(|t| t == "WIDECONT"),
        "the row above it *is* paired (its span covers the restart's column), \
         so it stays a merge continuation and hides its own content — drawn: \
         {drawn:?}"
    );
}

/// The same calibration: promoted means ordinary, not merely drawn.
#[test]
fn a_continue_below_a_wider_continue_matches_the_unmerged_layout() {
    let orphan = layout(&wide_continue_over_narrow_restart_table("<w:vMerge/>"));
    let control = layout(&wide_continue_over_narrow_restart_table(""));

    for needle in ["RESTART", "PEER", "PLAIN", "ORPHAN"] {
        assert!(
            (y_of(&orphan, needle) - y_of(&control, needle)).abs() < 0.01,
            "{needle} moved: y={:.2} vs control y={:.2}",
            y_of(&orphan, needle),
            y_of(&control, needle),
        );
    }
}

/// §17.4.84: `<w:vMerge/>` with no `@val` **is** `continue`, and it is the
/// spelling documents actually use — every one of the 98 `continue` elements in
/// the 52-document corpus is bare. The span-1 orphan repair is asserted above
/// against `w:val="continue"`; this pins that the bare spelling reaches the
/// same code, so a parse-level regression that stopped reading it could not
/// hide behind the explicit form.
#[test]
fn a_bare_vmerge_orphan_keeps_its_cell_content() {
    let bare = layout(&orphan_continue_table("<w:vMerge/>"));
    let explicit = layout(&orphan_continue_table(r#"<w:vMerge w:val="continue"/>"#));

    assert!(
        all_text(&bare).iter().any(|t| t == "ORPHAN"),
        "drawn: {:?}",
        all_text(&bare)
    );
    assert_eq!(
        all_text(&bare),
        all_text(&explicit),
        "the two spellings of `continue` must render identically"
    );
}

/// The other half of the bare spelling: a bare `continue` that *is* paired must
/// still be a merge. Without this, "bare parses to nothing at all" would pass
/// every orphan assertion above.
#[test]
fn a_bare_vmerge_that_is_paired_still_hides_its_own_content() {
    let pages = layout(&doc(r#"<w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>
          <w:p><w:r><w:t>MERGED</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:vMerge/></w:tcPr>
          <w:p><w:r><w:t>HIDDEN</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>PEER2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"#));
    let drawn = all_text(&pages);

    assert!(drawn.iter().any(|t| t == "MERGED"), "drawn: {drawn:?}");
    assert!(
        !drawn.iter().any(|t| t == "HIDDEN"),
        "a bare `<w:vMerge/>` under a `restart` is a merge continuation — \
         drawn: {drawn:?}"
    );
}

/// §17.4.6: a `cantSplit` row taller than a whole page fits nowhere, so the
/// paginator must not advance to a fresh page it cannot use. It used to
/// abandon an empty leading slice, which the section layer turned into a
/// blank first page.
#[test]
fn oversized_cant_split_row_does_not_emit_a_blank_leading_page() {
    let paras: String = (0..70)
        .map(|i| format!("<w:p><w:r><w:t>line {i}</w:t></w:r></w:p>"))
        .collect();
    let pages = layout(&doc(&format!(
        r#"<w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
      <w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc>
        <w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>
        {paras}
      </w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>after</w:t></w:r></w:p>"#
    )));

    assert!(
        !pages[0].commands.is_empty(),
        "page 0 is blank — the table advanced to a page that gave it no more room"
    );
}

/// The converse: a table that genuinely doesn't fit in the *remaining* space
/// but does fit on a fresh page must still move. Guards against a fix that
/// simply stops advancing.
#[test]
fn table_after_body_text_still_moves_to_the_next_page_when_it_does_not_fit() {
    let filler: String = (0..60)
        .map(|i| format!("<w:p><w:r><w:t>filler {i}</w:t></w:r></w:p>"))
        .collect();
    let rows: String = (0..12)
        .map(|i| {
            format!(
                r#"<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc>
                     <w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>
                     <w:p><w:r><w:t>row {i}</w:t></w:r></w:p></w:tc></w:tr>"#
            )
        })
        .collect();
    let pages = layout(&doc(&format!(
        r#"{filler}
    <w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
      {rows}
    </w:tbl>"#
    )));

    assert!(
        pages.len() > 1,
        "filler plus a 12-row table should not fit on one page"
    );
    for (i, page) in pages.iter().enumerate() {
        assert!(!page.commands.is_empty(), "page {i} is blank");
    }
}

// ── §17.4.80: what a declared row height measures ───────────────────────────

/// A three-row, one-column table with a 3pt `insideH`, whose **middle** row
/// carries `tr_pr` verbatim.
fn banded_table(tr_pr: &str) -> Vec<u8> {
    let row = |extra: &str, label: &str| {
        format!(
            "<w:tr>{extra}<w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>\
               <w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:tc></w:tr>"
        )
    };
    doc(&format!(
        r#"<w:tbl>
      <w:tblPr>
        <w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="24" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="24" w:space="0" w:color="000000"/>
          <w:left w:val="nil"/><w:right w:val="nil"/>
          <w:insideH w:val="single" w:sz="24" w:space="0" w:color="000000"/>
          <w:insideV w:val="nil"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
      {upper}{middle}{lower}
    </w:tbl>"#,
        upper = row("", "upper"),
        middle = row(tr_pr, "middle"),
        lower = row("", "lower"),
    ))
}

/// The middle row's **content box**: the gap between the two rules bounding it.
fn middle_row_box(bytes: &[u8]) -> f32 {
    let pages = layout(bytes);
    let mut rules: Vec<(f32, f32)> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } if rect.size.height.raw() < rect.size.width.raw() => {
                Some((
                    rect.origin.y.raw(),
                    (rect.origin.y + rect.size.height).raw(),
                ))
            }
            _ => None,
        })
        .collect();
    rules.sort_by(|a, b| a.0.total_cmp(&b.0));
    assert_eq!(
        rules.len(),
        4,
        "table top, two interior rules, table bottom"
    );
    rules[2].0 - rules[1].1
}

/// §17.4.80 measures the row's **content box**, and an *interior* rule eats
/// into it from both sides — half the boundary above, half the one below — the
/// same way a shared vertical eats into a cell's width: a row declaring 40pt
/// between two 3pt `insideH` rules gets 40 − 1.5 − 1.5 = 37pt of cell.
///
/// **Measured** against a Word render of `test-files/issue-157-empty-row-edge.docx`,
/// table 4 — re-measured 2026-09-05, pixel-counted off a fresh render and
/// calibrated against the table's own fixed-layout width (200pt / 750px)
/// rather than eyeballed. 40pt is chosen to clear the 6pt of rule comfortably,
/// so the reading is not confused with a floor.
///
/// This has been two other ways, each wrong and each from a single render. The
/// first (shipped briefly) put the whole 6pt of rule inside the box, by
/// analogy with `border-content-charge.docx`'s shared-vertical finding taken
/// literally rather than halved. The second — recorded here from 2026-08-19 to
/// 2026-09-05 — read the same render as a clean 40pt with the rules wholly
/// outside, on the strength of tables 2 and 3, whose rows are small enough
/// that a 2pt cell and a hairline are hard to tell apart. Table 4 is 20 times
/// that size, and a second, calibrated measurement puts it at 37pt, not 40.
#[test]
fn an_exact_row_height_is_charged_half_of_each_interior_border() {
    let box_h = middle_row_box(&banded_table(
        r#"<w:trPr><w:trHeight w:val="800" w:hRule="exact"/></w:trPr>"#,
    ));
    assert!(
        (box_h - 37.0).abs() < 0.01,
        "800 twips is 40pt minus half of each of its two 3pt interior rules, \
         37pt of cell, not {box_h}pt"
    );
}

/// …and a row declaring less than the rules charged against it floors at zero
/// rather than going negative — the same rule as the 40pt case above, not a
/// separate collapse. Word's table 3 draws this row as a hairline rather than
/// a clean 2pt of cell, which is what a zero floor predicts and a pass-through
/// of the declared value does not.
///
/// A floored, zero-height row leaves its two interior rules with nothing
/// between them, so `middle_row_box`'s "table top, two interior rules, table
/// bottom" no longer holds — the pair coalesces into one 6pt rect the way any
/// two abutting same-colour rects do (`coalesce_abutting_rects`), leaving 3
/// rects rather than 4. That coalesced band is asserted directly instead.
#[test]
fn a_row_shorter_than_its_rules_floors_at_zero_rather_than_going_negative() {
    let pages = layout(&banded_table(
        r#"<w:trPr><w:trHeight w:val="40" w:hRule="exact"/></w:trPr>"#,
    ));
    let mut rules: Vec<(f32, f32)> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } if rect.size.height.raw() < rect.size.width.raw() => {
                Some((
                    rect.origin.y.raw(),
                    (rect.origin.y + rect.size.height).raw(),
                ))
            }
            _ => None,
        })
        .collect();
    rules.sort_by(|a, b| a.0.total_cmp(&b.0));
    assert_eq!(
        rules.len(),
        3,
        "table top, one coalesced interior band, table bottom — not 4 \
         separate rules, since the floored row leaves none of its own \
         between them"
    );
    let band_h = rules[1].1 - rules[1].0;
    assert!(
        (band_h - 6.0).abs() < 0.01,
        "40 twips (2pt) is less than the 3pt charged from each side, so the \
         content box floors at zero and its two 3pt interior rules meet with \
         nothing between them: one 6pt band, not {band_h}pt"
    );
}

/// §17.4.80 `atLeast` measures the same box, as a floor rather than a pin —
/// Word's table 5 is table 4's declaration with the rule changed, and draws the
/// same 40pt.
#[test]
fn an_at_least_row_height_is_measured_the_same_way() {
    let box_h = middle_row_box(&banded_table(
        r#"<w:trPr><w:trHeight w:val="800" w:hRule="atLeast"/></w:trPr>"#,
    ));
    assert!(
        (box_h - 40.0).abs() < 0.01,
        "800 twips is 40pt of cell under atLeast too, not {box_h}pt"
    );
}

/// §17.4.80 with **`w:val="0"`** is no constraint at all: the row is as tall as
/// its content, not flat.
///
/// [MS-OI29500] §2.4.77(c) records that Word requires `val="0"` whenever
/// `hRule="auto"`, which makes zero the marker for *unconstrained* rather than a
/// height of nothing — and Word renders it that way even when the rule says
/// `exact`. Its table 2 draws a full row of cell where a literal reading draws
/// none at all, which is how this was found.
///
/// Asserted as a relation to the same table with no `trHeight`, so the empty
/// paragraph's line height is never pinned — whatever it is, both sides get it.
#[test]
fn an_exact_height_of_zero_is_no_constraint_at_all() {
    let zero = middle_row_box(&banded_table(
        r#"<w:trPr><w:trHeight w:val="0" w:hRule="exact"/></w:trPr>"#,
    ));
    let unconstrained = middle_row_box(&banded_table(""));
    assert!(
        (zero - unconstrained).abs() < 0.01,
        "a zero exact height must lay out like no height at all: {zero} vs {unconstrained}"
    );
    assert!(
        zero > 1.0,
        "and that is a real row, not a flat one: {zero}pt"
    );
}
