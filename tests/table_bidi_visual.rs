//! §17.4.1 `w:bidiVisual` — a table whose columns run right to left.
//!
//! The element says the first cell of a row is the **rightmost** one. Everything
//! that follows from that is geometry: a `w:gridSpan` covers the mirrored run of
//! slots, a `w:gridBefore` leaves its gap at the visual right, a `w:vMerge`
//! spans the mirrored column, and a cell's logical start border paints on its
//! visual right.
//!
//! That last one is the only reading here ECMA-376 does not spell out in one
//! place, and it is settled from inside this repo rather than guessed: the
//! Transitional `w:left`/`w:right` this parser reads *are* Strict's
//! `w:start`/`w:end`, which is why `docx::parse::properties::schema::border` and
//! `::insets` already declare them as serde aliases of one another. They are
//! logical edges, so a `w:left` belongs at the cell's logical start — which
//! under `bidiVisual` is on the right.
//!
//! # How these tests are written
//!
//! Every assertion is a **relation between the same table with and without
//! `<w:bidiVisual/>`**, not a list of coordinates. A cell's box must satisfy
//!
//! ```text
//! x_rtl − table_left_rtl == table_right_ltr − (x_ltr + w_ltr)
//! ```
//!
//! with its width unchanged — a reflection about each table's own edges, stated
//! as a distance from those edges so it holds wherever the table sits. Written
//! that way, no page origin, cell margin or glyph metric has to be known, and
//! the assertions survive any later refinement of the geometry they reflect.
//!
//! **Where the table sits is a separate claim**, and it is asserted in
//! `tests/table_leading_margin.rs` rather than here: §17.4.1 also makes the
//! table's leading margin the right one, so the two tables in a pair do *not*
//! share a span. Reflecting about a shared span was this file's first reading
//! and is what made the defect invisible.
//!
//! The declared grid is deliberately **unequal** (1000/2000/3000 twips). With
//! three equal columns, reversing the cells while leaving the slot widths in
//! place produces exactly the same page as reversing both, so an equal grid
//! cannot tell a correct mirror from half of one.

use std::collections::HashMap;
use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn make_docx(document_xml: &str, styles_xml: Option<&str>) -> Vec<u8> {
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
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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

        if let Some(styles) = styles_xml {
            zip.start_file("word/_rels/document.xml.rels", o).unwrap();
            zip.write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
            )
            .unwrap();
            zip.start_file("word/styles.xml", o).unwrap();
            zip.write_all(styles.as_bytes()).unwrap();
        }

        zip.start_file("word/document.xml", o).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();
        zip.finish().unwrap();
    }
    buf
}

/// The fixture's grid: 1000 + 2000 + 3000 twips = 50 + 100 + 150 pt.
const GRID: &str = r#"<w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/>"#;

/// One table of `rows`, with `<w:bidiVisual/>` when `bidi`, plus any extra
/// `tblPr` children. `w:tblLayout="fixed"` and an explicit `w:tblW` keep the
/// declared grid from being rescaled, so the slot widths under test are the
/// ones written here.
fn table(bidi: bool, extra_tbl_pr: &str, rows: &str) -> String {
    let flag = if bidi { "<w:bidiVisual/>" } else { "" };
    format!(
        r#"<w:tbl>
  <w:tblPr>
    <w:tblW w:w="6000" w:type="dxa"/>
    <w:tblLayout w:type="fixed"/>
    {flag}{extra_tbl_pr}
  </w:tblPr>
  <w:tblGrid>{GRID}</w:tblGrid>
  {rows}
</w:tbl>"#
    )
}

fn layout(body: &str, styles: Option<&str>) -> Vec<LayoutedPage> {
    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"#
    );
    let doc = dxpdf::docx::parse(&make_docx(&document_xml, styles)).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// A shaded cell. The fill is how these tests read a cell's box off the page:
/// §17.4.33 shading is emitted at the cell box exactly, and a distinct colour
/// per cell keeps `coalesce_abutting_rects` from fusing neighbours.
fn cell(fill: &str, extra_tc_pr: &str) -> String {
    format!(
        r#"<w:tc>
  <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>{extra_tc_pr}</w:tcPr>
  <w:p><w:r><w:t>x</w:t></w:r></w:p>
</w:tc>"#
    )
}

type Rect = (f32, f32, f32, f32);

/// Every shaded box on the page, by fill colour.
fn boxes(pages: &[LayoutedPage]) -> HashMap<(u8, u8, u8), Rect> {
    let mut out = HashMap::new();
    for c in pages.iter().flat_map(|p| &p.commands) {
        if let DrawCommand::Rect { rect, color } = c {
            out.entry((color.r, color.g, color.b)).or_insert((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.size.width.raw(),
                rect.size.height.raw(),
            ));
        }
    }
    out
}

const RED: (u8, u8, u8) = (0xFF, 0x00, 0x00);
const GREEN: (u8, u8, u8) = (0x00, 0xFF, 0x00);
const BLUE: (u8, u8, u8) = (0x00, 0x00, 0xFF);

/// The `(left, right)` extent of a set of boxes.
fn extent(rects: impl Iterator<Item = Rect>) -> (f32, f32) {
    rects.fold(
        (f32::INFINITY, f32::NEG_INFINITY),
        |(l, r), (x, _, w, _)| (l.min(x), r.max(x + w)),
    )
}

/// The table's own span, as `(left, right)`, taken from the shaded boxes rather
/// than from the page margin — so the reflection is about the table's edges and
/// not about anything the section happens to set.
fn span(rects: &HashMap<(u8, u8, u8), Rect>) -> (f32, f32) {
    extent(rects.values().copied())
}

/// Assert that every box in `rtl` is the reflection of its `ltr` twin about its
/// own table's edges, with its width unchanged.
///
/// Measured as a distance from each table's own left edge, because the two
/// tables do **not** share a span: §17.4.1 puts the right-to-left one at the
/// right margin (`tests/table_leading_margin.rs`). Asserting the spans equal is
/// what this file used to do, and it pinned the placement defect in place.
fn assert_mirrored(ltr: &HashMap<(u8, u8, u8), Rect>, rtl: &HashMap<(u8, u8, u8), Rect>) {
    assert_eq!(ltr.len(), rtl.len(), "same cells either way");
    let (l, r) = span(ltr);
    let (rl, rr) = span(rtl);
    assert_eq!(rr - rl, r - l, "a mirrored table is the same width");
    for (colour, &(x, _, w, _)) in ltr {
        let &(rx, _, rw, _) = rtl
            .get(colour)
            .unwrap_or_else(|| panic!("no {colour:?} cell in the mirrored table"));
        assert_eq!(rw, w, "{colour:?}: a mirrored cell keeps its width");
        assert_eq!(
            rx - rl,
            r - (x + w),
            "{colour:?}: {x}..{} sits {} from the control's right edge, so its \
             mirror must sit that far from the mirrored table's left edge",
            x + w,
            r - (x + w)
        );
    }
}

// ── the mirror itself ───────────────────────────────────────────────────────

/// The first cell of a row is the rightmost one, and the slot widths travel with
/// the columns.
///
/// The non-vacuity assertions are what make the reflection mean something: the
/// three cells have three *different* widths and are in ascending order without
/// the flag, so a renderer that ignored `bidiVisual` — or that reversed the
/// cells while leaving the slot widths where they were — fails.
#[test]
fn the_first_cell_of_a_row_becomes_the_rightmost() {
    let row = format!(
        "<w:tr>{}{}{}</w:tr>",
        cell("FF0000", ""),
        cell("00FF00", ""),
        cell("0000FF", "")
    );
    let ltr = boxes(&layout(&table(false, "", &row), None));
    let rtl = boxes(&layout(&table(true, "", &row), None));

    // Non-vacuity: the declared grid really is unequal and really is in source
    // order without the flag.
    assert_eq!(ltr[&RED].2, 50.0, "1000 twips");
    assert_eq!(ltr[&GREEN].2, 100.0, "2000 twips");
    assert_eq!(ltr[&BLUE].2, 150.0, "3000 twips");
    assert!(
        ltr[&RED].0 < ltr[&GREEN].0 && ltr[&GREEN].0 < ltr[&BLUE].0,
        "without the flag the cells run left to right"
    );

    assert_mirrored(&ltr, &rtl);

    // …and the headline claim, stated directly rather than as a difference.
    assert!(
        rtl[&RED].0 > rtl[&GREEN].0 && rtl[&GREEN].0 > rtl[&BLUE].0,
        "with it they run right to left: {rtl:?}"
    );
}

// ── §17.4.17 `w:gridSpan` ───────────────────────────────────────────────────

/// A span covers the mirrored *run* of slots — not one slot, and not the run it
/// covered before the flip.
///
/// The grid is unequal, so this is a real constraint: the red cell spans
/// 1000 + 2000 twips on the left and must come out spanning the same two
/// columns once they are the rightmost two, keeping its 150 pt.
#[test]
fn a_grid_span_covers_the_mirrored_run_of_slots() {
    let row = format!(
        "<w:tr>{}{}</w:tr>",
        cell("FF0000", r#"<w:gridSpan w:val="2"/>"#),
        cell("00FF00", "")
    );
    let ltr = boxes(&layout(&table(false, "", &row), None));
    let rtl = boxes(&layout(&table(true, "", &row), None));

    assert_eq!(ltr[&RED].2, 150.0, "1000 + 2000 twips");
    assert_eq!(ltr[&GREEN].2, 150.0, "3000 twips");
    assert_mirrored(&ltr, &rtl);
}

// ── §17.4.15 `w:gridBefore` ─────────────────────────────────────────────────

/// A row's skipped columns are skipped at its *logical* start, so under the flag
/// the gap is at the visual right.
///
/// The second row spans the whole grid and is what makes the first row's gap
/// measurable: a row with a gap cannot say where the table's edges are, so on
/// its own every claim below would be about the cell relative to itself and
/// would hold however the gap fell.
#[test]
fn a_grid_before_gap_moves_to_the_visual_right() {
    let rows = format!(
        "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>{}</w:tr><w:tr>{}</w:tr>",
        cell("FF0000", r#"<w:gridSpan w:val="2"/>"#),
        cell("00FF00", r#"<w:gridSpan w:val="3"/>"#),
    );
    let ltr = boxes(&layout(&table(false, "", &rows), None));
    let rtl = boxes(&layout(&table(true, "", &rows), None));

    // 2000 + 3000 twips of cell, one 1000-twip column skipped ahead of it; the
    // full-width row below is the table's own 6000.
    assert_eq!(ltr[&RED].2, 250.0);
    assert_eq!(ltr[&GREEN].2, 300.0, "the reference row spans the grid");
    let (l, r) = span(&ltr);
    assert_eq!((l, r), (ltr[&GREEN].0, ltr[&GREEN].0 + 300.0));

    assert_eq!(
        ltr[&RED].0 - l,
        50.0,
        "without the flag the skipped 1000-twip column is on the left"
    );
    // Against the *mirrored* table's own right edge — §17.4.1 moves the table
    // to the right margin, so `r` above is the control's edge and not this
    // table's (`tests/table_leading_margin.rs`).
    let (_, rr) = span(&rtl);
    assert_eq!(
        rr - (rtl[&RED].0 + rtl[&RED].2),
        50.0,
        "with it the same column is skipped on the right"
    );
    assert_mirrored(&ltr, &rtl);
}

// ── §17.4.84 `w:vMerge` ─────────────────────────────────────────────────────

/// A merge spans the mirrored column, and still spans both rows.
///
/// The height is asserted as well as the box, because a merge that lost its
/// span would still mirror correctly as a one-row cell.
#[test]
fn a_vertical_merge_mirrors_with_its_column() {
    let rows = format!(
        "<w:tr>{}{}{}</w:tr><w:tr>{}{}{}</w:tr>",
        cell("FF0000", r#"<w:vMerge w:val="restart"/>"#),
        cell("00FF00", ""),
        cell("0000FF", ""),
        cell("FF0000", "<w:vMerge/>"),
        cell("00FFFF", ""),
        cell("FF00FF", ""),
    );
    let ltr = boxes(&layout(&table(false, "", &rows), None));
    let rtl = boxes(&layout(&table(true, "", &rows), None));

    assert_mirrored(&ltr, &rtl);
    assert_eq!(
        rtl[&RED].3, ltr[&RED].3,
        "the merged cell keeps the height of its span"
    );
    assert!(
        rtl[&RED].3 > rtl[&GREEN].3,
        "and that height really is more than one row: {rtl:?}"
    );
}

// ── §17.4.39 the logical start edge ─────────────────────────────────────────

/// A cell that declares only `w:left` paints that border on its **visual right**
/// under the flag, because `w:left` is Strict's `w:start` — the cell's logical
/// leading edge.
///
/// Read as the border's offset within its own cell box, so the assertion is
/// about which side of the cell the line is on and not about where the cell is.
#[test]
fn a_cells_start_border_paints_on_its_visual_right() {
    let row = format!(
        "<w:tr>{}</w:tr>",
        cell(
            "FF0000",
            r#"<w:tcBorders>
                 <w:left w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
                 <w:top w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
               </w:tcBorders>"#
        )
    );

    // `(gap left of the border, gap right of it)` within the cell's own box, so
    // the claim is which side of the cell the line sits on and needs neither the
    // cell's width nor its position spelled out.
    let gaps = |bidi: bool| -> (f32, f32) {
        let pages = layout(&table(bidi, "", &row), None);
        let rects = boxes(&pages);
        let (cx, _, cw, _) = rects[&RED];
        let (bx, _, bw, _) = rects[&BLUE];
        assert_eq!(bw, 3.0, "w:sz=24 is 3pt");
        (bx - cx, (cx + cw) - (bx + bw))
    };

    let (before, after) = gaps(false);
    // Negative because §17.4.66 straddles: half the border lies outside the cell
    // box on the edge it stands on. Measured against `border-outer-box.docx`,
    // where Word draws a 12pt outer border at 360–372 against a declared right
    // edge of 372. It read 0.0 under the older "a border is drawn inside its
    // box" convention, which that render refuted.
    assert_eq!(before, -1.5, "without the flag, straddling the left edge");
    assert!(after > 0.0, "…and the rest of the cell is to its right");
    // With the flag the two gaps swap: the border is flush with the *right*
    // edge, and the same amount of cell is now to its left.
    assert_eq!(
        gaps(true),
        (after, before),
        "the logical start border moves to the visual right of its cell"
    );
}

/// §17.4.41 `w:tcMar` mirrors for the same reason its borders do: `w:left` is
/// `w:start`, so a cell's leading inset is on its visual right.
///
/// Read as the content's offset inside its own cell rather than as a page x, so
/// the claim is about the inset and not about where the cell ended up. The two
/// margins are deliberately very different (0 and 40pt), because equal ones
/// would make swapping them a no-op and the test vacuous.
#[test]
fn a_cells_margins_mirror_with_it() {
    let row = format!(
        "<w:tr>{}</w:tr>",
        cell(
            "FF0000",
            r#"<w:tcMar>
                 <w:left w:w="0" w:type="dxa"/><w:right w:w="800" w:type="dxa"/>
                 <w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>
               </w:tcMar>"#
        )
    );

    let inset = |bidi: bool| -> f32 {
        let pages = layout(&table(bidi, "", &row), None);
        let (cx, _, _, _) = boxes(&pages)[&RED];
        pages
            .iter()
            .flat_map(|p| &p.commands)
            .find_map(|c| match c {
                DrawCommand::Text { text, position, .. } if &**text == "x" => {
                    Some(position.x.raw() - cx)
                }
                _ => None,
            })
            .expect("the cell's own text")
    };

    // 800 twips is 40pt, and the leading margin is the one the content sits
    // behind: 0 on the left without the flag, 40 with it.
    assert_eq!(inset(false), 0.0, "w:left = 0 leads without the flag");
    assert_eq!(inset(true), 40.0, "w:right = 800 twips leads with it");
}

/// §17.4.38 the *table's* own `w:left`, as opposed to a cell's — a separate
/// field on a separate struct (`TableBorderConfig`), reached by a separate
/// branch of the mirror, and so not covered by the cell-border case above.
#[test]
fn a_table_level_start_border_paints_on_the_visual_right() {
    let row = format!("<w:tr>{}</w:tr>", cell("FF0000", ""));
    let borders = r#"<w:tblBorders>
      <w:left w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
      <w:top w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
      <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
    </w:tblBorders>"#;

    let side = |bidi: bool| -> (f32, f32) {
        let rects = boxes(&layout(&table(bidi, borders, &row), None));
        let (cx, _, cw, _) = rects[&RED];
        let (bx, _, bw, _) = rects[&BLUE];
        (bx - cx, (cx + cw) - (bx + bw))
    };

    let (before, after) = side(false);
    // Straddling, so half the 3pt border is outside the cell box — see
    // `a_cells_start_border_paints_on_its_visual_right` for the render that
    // settled it.
    assert_eq!(before, -1.5, "the table's start edge is on the left");
    assert!(after > 0.0);
    assert_eq!(
        side(true),
        (after, before),
        "and on the right once the columns reverse"
    );
}

/// §17.4.60 `w:tblPrEx/w:tblBorders` — a *row's* override of those same edges,
/// which is a third struct on a third branch of the mirror.
///
/// The second row carries no override and is the control: it must not move,
/// which is what distinguishes a mirrored override from a mirrored table.
#[test]
fn a_row_level_border_override_mirrors_too() {
    let rows = format!(
        "<w:tr><w:tblPrEx><w:tblBorders>\
           <w:left w:val=\"single\" w:sz=\"24\" w:space=\"0\" w:color=\"0000FF\"/>\
           <w:top w:val=\"nil\"/><w:bottom w:val=\"nil\"/><w:right w:val=\"nil\"/>\
           <w:insideH w:val=\"nil\"/><w:insideV w:val=\"nil\"/>\
         </w:tblBorders></w:tblPrEx>{}</w:tr><w:tr>{}</w:tr>",
        cell("FF0000", ""),
        cell("00FF00", ""),
    );

    let side = |bidi: bool| -> (f32, f32) {
        let rects = boxes(&layout(&table(bidi, "", &rows), None));
        let (cx, _, cw, _) = rects[&RED];
        let (bx, _, bw, _) = rects[&BLUE];
        // The override belongs to the row that declares it, so it stands on one
        // of that row's own vertical edges — straddling it, half in and half
        // out, which is why this is a test of the border's *centre* and not of
        // containment as it was before `border-outer-box.docx` was measured.
        let centre = bx + bw / 2.0;
        assert!(
            (centre - cx).abs() < 0.01 || (centre - (cx + cw)).abs() < 0.01,
            "the override stands on an edge of the row that declares it"
        );
        (bx - cx, (cx + cw) - (bx + bw))
    };

    let (before, after) = side(false);
    assert_eq!(before, -1.5);
    assert!(after > 0.0);
    assert_eq!(side(true), (after, before));
}

/// A row addressing more grid columns than the table declares is malformed
/// input, and must not panic here — the arithmetic that renumbers `gridBefore`
/// subtracts two counts that a bad row can push past zero.
///
/// `seat_every_cell` repairs the grid ahead of this and `measure_table_rows`
/// clamps behind it, so the assertion is only that the content survives: a
/// document Word recovers from must not take the renderer down.
#[test]
fn a_row_that_overruns_its_grid_does_not_panic() {
    let rows = format!(
        "<w:tr><w:trPr><w:gridBefore w:val=\"9\"/></w:trPr>{}</w:tr>",
        cell("FF0000", r#"<w:gridSpan w:val="4"/>"#)
    );
    let pages = layout(&table(true, "", &rows), None);
    assert!(
        pages
            .iter()
            .flat_map(|p| &p.commands)
            .any(|c| matches!(c, DrawCommand::Text { text, .. } if &**text == "x")),
        "the overrun row's content still reaches the page"
    );
}

// ── §17.7.6 conditional formatting stays logical ────────────────────────────

/// `firstColumn` keeps meaning the **logical** first column, which under the
/// flag is the rightmost one.
///
/// This is the test that pins the ordering the whole design rests on: the
/// conditional region is resolved on logical grid columns before the mirror is
/// applied, so moving the mirror any earlier would silently shade the wrong
/// column. Without the flag the same style shades the leftmost cell, which is
/// the control.
#[test]
fn the_first_column_region_stays_the_logical_first_column() {
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="TestTbl">
    <w:name w:val="Test Table"/>
    <w:tblStylePr w:type="firstCol">
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:tcPr>
    </w:tblStylePr>
  </w:style>
</w:styles>"#;
    // No per-cell shading: the only fill on the page is the one the firstCol
    // layer paints, so its box *is* the answer.
    let plain = r#"<w:tc><w:tcPr/><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>"#;
    let row = format!("<w:tr>{plain}{plain}{plain}</w:tr>");
    let extra = r#"<w:tblStyle w:val="TestTbl"/>
                   <w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="1"
                              w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>"#;

    let ltr = boxes(&layout(&table(false, extra, &row), Some(styles)));
    let rtl = boxes(&layout(&table(true, extra, &row), Some(styles)));

    let (l, r) = (ltr[&GREEN].0, ltr[&GREEN].0 + ltr[&GREEN].2);
    assert_eq!(
        ltr[&GREEN].2, 50.0,
        "the logical first column is 1000 twips"
    );
    assert_eq!(rtl[&GREEN].2, 50.0, "…and still is, on the other side");
    assert!(
        rtl[&GREEN].0 > l,
        "the shaded column must move right, from {l}..{r} to {:?}",
        rtl[&GREEN]
    );
}

// ── the committed fixture, end to end ───────────────────────────────────────

/// `test-files/bidi-visual-table.docx` renders as its own control.
///
/// Everything above builds its document in memory, which skips the package: the
/// `.docx` is never opened, so a defect in ZIP handling, part discovery or the
/// `w:tblPr` seam inside a real file would not show. This is the only case that
/// runs the whole path, and it is also what makes the fixture worth committing
/// — a document a Word render can be compared against later has to be the same
/// document a test reads now.
///
/// The fixture holds the same table twice, differing **only** in
/// `<w:bidiVisual/>`, with a distinct fill per cell. So each colour appears
/// once per table, the upper box is the control and the lower is the mirror,
/// and no coordinate has to be written down here at all.
#[test]
fn the_committed_fixture_mirrors_its_own_control() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/test-files/bidi-visual-table.docx"
    );
    let bytes = std::fs::read(path).expect("read the fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse the fixture");
    let pages = dxpdf::render::resolve_and_layout(doc).1;

    // Every fill in the fixture, paired: the control's box and the mirrored
    // one. A `vMerge` continuation paints nothing, so a colour still appears
    // exactly twice — once per table.
    let mut by_colour: HashMap<(u8, u8, u8), Vec<Rect>> = HashMap::new();
    for c in pages.iter().flat_map(|p| &p.commands) {
        if let DrawCommand::Rect { rect, color } = c {
            by_colour
                .entry((color.r, color.g, color.b))
                .or_default()
                .push((
                    rect.origin.x.raw(),
                    rect.origin.y.raw(),
                    rect.size.width.raw(),
                    rect.size.height.raw(),
                ));
        }
    }
    // The eleven cell fills the generator writes, named rather than inferred:
    // "every colour that happens to appear twice" would also sweep in the
    // `C00000` start border, which pairs too and is asserted on its own above.
    const FILLS: [(u8, u8, u8); 11] = [
        (0xF8, 0xCB, 0xAD), // A
        (0xC6, 0xE0, 0xB4), // B
        (0xBD, 0xD7, 0xEE), // C
        (0xFF, 0xE6, 0x99), // D, gridSpan
        (0xD9, 0xD2, 0xE9), // E
        (0xF4, 0xCC, 0xCC), // F, gridBefore
        (0xD0, 0xE0, 0xE3), // G, vMerge restart
        (0xEA, 0xD1, 0xDC), // H
        (0xFF, 0xF2, 0xCC), // I
        (0xD9, 0xEA, 0xD3), // J
        (0xCF, 0xE2, 0xF3), // K
    ];
    let mut cells: Vec<((u8, u8, u8), Rect, Rect)> = Vec::new();
    for colour in FILLS {
        let mut rects = by_colour
            .get(&colour)
            .unwrap_or_else(|| panic!("fixture no longer paints {colour:?}"))
            .clone();
        // One box per table: a `vMerge` continuation paints nothing, so even
        // the merged column contributes exactly two.
        assert_eq!(rects.len(), 2, "{colour:?} must appear once per table");
        rects.sort_by(|a, b| a.1.total_cmp(&b.1));
        cells.push((colour, rects[0], rects[1]));
    }
    assert_eq!(cells.len(), 11);

    // Each table's own edges, taken from its own cells. The two are *not* the
    // same span — see the placement assertion below.
    let (ctrl_left, ctrl_right) = extent(cells.iter().map(|c| c.1));
    let (mirror_left, mirror_right) = extent(cells.iter().map(|c| c.2));

    // §17.4.28 / §17.4.50: the whole reason the two spans differ. Neither table
    // carries a `w:jc` and the section carries no `w:bidi`, so each sits at its
    // *own* leading margin — and `w:bidiVisual` is what makes the second one's
    // the right. This is what Word renders, and the fixture is a controlled
    // experiment for it: every cell paragraph in both tables is RTL, so
    // paragraph direction cannot be what moves one and not the other.
    //
    // The page is 12240 twips wide with 1440-twip margins.
    const CONTENT_LEFT: f32 = 72.0;
    const CONTENT_RIGHT: f32 = 72.0 + 468.0;
    assert_eq!(ctrl_left, CONTENT_LEFT, "the control is flush left");
    assert_eq!(
        mirror_right, CONTENT_RIGHT,
        "the w:bidiVisual table is flush right"
    );
    assert_eq!(
        mirror_right - mirror_left,
        ctrl_right - ctrl_left,
        "the two tables are the same width, so only their placement differs"
    );

    let mut moved = 0;
    for (colour, ctrl, mirror) in &cells {
        assert_eq!(mirror.2, ctrl.2, "{colour:?} keeps its width");
        assert_eq!(
            mirror.0 - mirror_left,
            ctrl_right - (ctrl.0 + ctrl.2),
            "{colour:?}: {ctrl:?} must reflect about its own table's edges"
        );
        if mirror.0 - mirror_left != ctrl.0 - ctrl_left {
            moved += 1;
        }
    }
    // Non-vacuity: a symmetric table reflects onto itself, so the fixture has
    // to contain cells that actually change place.
    assert!(
        moved >= 8,
        "only {moved} of {} cells moved — the fixture stopped discriminating",
        cells.len()
    );
}
