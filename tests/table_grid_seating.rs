//! §17.4.48 / §17.4.71 — the `<w:tblGrid>` must have a column for every cell a
//! row declares, and what happens when it does not.
//!
//! # The invariant, and the one place the spec leaves open
//!
//! §17.4.63 (`tblW`) and §17.4.71 (`tcW`) carry the same paragraph, word for
//! word: *"All widths in a table are considered preferred because: The table
//! **shall** satisfy the shared columns as specified by the `tblGrid` element …
//! Two or more widths can have conflicting values for the width of the same
//! grid column … The table layout algorithm can require a preference to be
//! overridden."* So the grid is the invariant, `tblW` and `tcW` are preferences
//! that may be overridden, and the spec names the conflict without resolving
//! it.
//!
//! What that leaves open is the **widths**, which is why nothing in this file
//! asserts that a `tcW` beats the grid slice it disagrees with. What it does
//! not leave open is the **seating**: a grid with fewer columns than a row has
//! cells cannot "satisfy the shared columns" for that row under any reading,
//! because there is no column for the last cell to sit in. Such a file is
//! self-contradictory and a renderer has to decide what gives.
//!
//! This file pins the half that must never move — a grid that *can* seat every
//! cell is scaled proportionally and nothing else happens to it. These are the
//! trap-detector for the repair in `build/table.rs::seat_every_cell`, which is
//! gated strictly on a grid too short to seat some row: if that gate ever
//! widens, these fail first. Seating counts exactly what the grid walk counts
//! (§17.4.15 `gridBefore`, §17.4.17 `gridSpan`, §17.4.14 `gridAfter`), so each
//! of those is pinned here as seating a grid rather than needing repair.

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

/// A bordered table with the given `<w:tblW>`, `<w:gridCol>` list and rows.
///
/// No `<w:sectPr>`, so the page is the §17.6.13 default Letter with 1-inch
/// margins: 612 pt wide, text column 72…540. No styles part either, so no
/// `TableNormal` cell margin — a cell's drawn extent is its column.
pub fn table_doc(tbl_w: &str, grid_cols: &str, rows: &str) -> Vec<u8> {
    make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        {tbl_w}
        <w:tblLayout w:type="fixed"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>{grid_cols}</w:tblGrid>
      {rows}
    </w:tbl>
  </w:body>
</w:document>"#
    ))
}

/// `<w:gridCol>`s from a list of twip widths.
pub fn grid(widths: &[i32]) -> String {
    widths
        .iter()
        .map(|w| format!(r#"<w:gridCol w:w="{w}"/>"#))
        .collect()
}

/// One `<w:tc>` labelled `text`, optionally carrying `w:tcW` and `w:gridSpan`.
pub fn cell(text: &str, tcw: Option<(i32, &str)>, span: Option<i32>) -> String {
    let w = match tcw {
        Some((v, t)) => format!(r#"<w:tcW w:w="{v}" w:type="{t}"/>"#),
        None => String::new(),
    };
    let s = match span {
        Some(n) => format!(r#"<w:gridSpan w:val="{n}"/>"#),
        None => String::new(),
    };
    format!(r#"<w:tc><w:tcPr>{w}{s}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"#)
}

pub fn row(cells: &str) -> String {
    format!("<w:tr>{cells}</w:tr>")
}

pub fn row_with(tr_pr: &str, cells: &str) -> String {
    format!("<w:tr><w:trPr>{tr_pr}</w:trPr>{cells}</w:tr>")
}

pub fn layout(bytes: &[u8]) -> Vec<LayoutedPage> {
    let parsed = dxpdf::docx::parse(bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// Every cell's drawn `(x, width)` in the table's first row, in cell order.
///
/// Read off the **verticals** standing on the first row, one per grid line its
/// cells reach. Not off the horizontal above them, which is what this used to
/// do: a junction is emitted among the horizontals now (`borders::junction_axes`
/// — Word gives the crossing to the horizontal), so `coalesce_abutting_rects`
/// fuses each one into the line it abuts and the whole boundary arrives as a
/// single unbroken rect with no gap in it to recover a grid line from. The
/// verticals are the family the junctions have been cut *out* of, so each is one
/// clean segment per row per grid line.
///
/// Recovering rather than counting rects is the point. A rect count measures the
/// decomposition; the geometry these tests are about is where the grid lines
/// fall, which is what a cell's column *is*.
pub fn first_row_cells(pages: &[LayoutedPage]) -> Vec<(f32, f32)> {
    // Every border rect as `(x0, x1, y0, y1)` — thin one way or the other. A
    // shading rect or an image would swamp the search.
    let rects: Vec<(f32, f32, f32, f32)> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => {
                let (x, y) = (rect.origin.x.raw(), rect.origin.y.raw());
                let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
                (w.min(h) <= 1.0).then_some((x, x + w, y, y + h))
            }
            _ => None,
        })
        .collect();

    // A vertical segment: thin in x, and taller than it is wide.
    let is_vertical =
        |&(x0, x1, y0, y1): &(f32, f32, f32, f32)| x1 - x0 <= 1.0 && y1 - y0 > x1 - x0;
    let first = rects
        .iter()
        .filter(|r| is_vertical(r))
        .map(|r| r.2)
        .fold(f32::INFINITY, f32::min);
    if !first.is_finite() {
        return Vec::new();
    }
    let mut bands: Vec<(f32, f32)> = rects
        .iter()
        .filter(|r| is_vertical(r) && (r.2 - first).abs() < 0.01)
        .map(|(x0, x1, ..)| (*x0, *x1))
        .collect();
    bands.sort_by(|p, q| p.0.total_cmp(&q.0));
    bands.dedup_by(|a, b| (a.0 - b.0).abs() < 0.01);
    if bands.len() < 2 {
        return Vec::new();
    }

    // §17.4.66: a border straddles a line two cells share and goes **inside** a
    // line shared with nothing. So an interior grid line is its band's centre,
    // while the table's own two edges are the outer faces of theirs.
    let last = bands.len() - 1;
    let lines: Vec<f32> = bands
        .iter()
        .enumerate()
        .map(|(i, &(x0, x1))| match i {
            0 => x0,
            i if i == last => x1,
            _ => (x0 + x1) * 0.5,
        })
        .collect();
    lines.windows(2).map(|p| (p[0], p[1] - p[0])).collect()
}

/// How far a recovered column may sit from its grid line: **half a border**.
///
/// The interior grid lines come back exactly — they are the middle of the gap
/// between two segments, which is the line whatever the borders do. The two
/// **outer** ends do not: §17.4.66 puts a border wholly inside the table's own
/// edge and straddling every other line, and a `gridBefore`/`gridAfter` row ends
/// on an *interior* line, so the ink's outer end is the grid line in one case
/// and half a border past it in the other. Nothing in the command stream says
/// which without knowing the table's own extent, and adding that to
/// `DrawCommand` for a test's benefit is not worth it.
///
/// So the tolerance is exactly that half-border, doubled for a width because a
/// column has two ends. It costs these tests nothing: they are about columns
/// tens of points wide, and where a border sits on its line is
/// `tests/table_grid_gap_borders.rs`' subject, not this file's.
const HALF_BORDER: f32 = 0.26;

/// Assert `got` matches `want` as `(x, width)` pairs, to [`HALF_BORDER`].
pub fn assert_cells(got: &[(f32, f32)], want: &[(f32, f32)], what: &str) {
    assert_eq!(
        got.len(),
        want.len(),
        "{what}: expected {} cells, drew {}: {got:?}",
        want.len(),
        got.len()
    );
    for (i, (g, w)) in got.iter().zip(want).enumerate() {
        assert!(
            (g.0 - w.0).abs() < HALF_BORDER && (g.1 - w.1).abs() < 2.0 * HALF_BORDER,
            "{what}: cell {i} drawn at x={:.2} w={:.2}, expected x={:.2} w={:.2} — all: {got:?}",
            g.0,
            g.1,
            w.0,
            w.1
        );
    }
}

// ── The half that must never move ────────────────────────────────────────────

/// §17.4.48: a grid with one column per cell is scaled to the declared `tblW`
/// and otherwise used as declared. 500/1000/1500 twips are one, two and three
/// sixths of the grid, and the grid sums to 3000 against a declared `tblW` of
/// 6000 — so the scale factor is 2 and the 300 pt table's columns come out at
/// 50, 100 and 150 pt from the 72 pt left margin.
///
/// The grid deliberately sums to *half* the declared width rather than to it.
/// With `sum(grid) == tblW` the scale factor is 1, and a mutation deleting the
/// scaling entirely still passes — which is exactly what the mutation check
/// caught when this test was first written that way.
#[test]
fn a_grid_that_seats_every_cell_is_scaled_proportionally() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="6000" w:type="dxa"/>"#,
        &grid(&[500, 1000, 1500]),
        &row(&format!(
            "{}{}{}",
            cell("a", None, None),
            cell("b", None, None),
            cell("c", None, None)
        )),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 50.0), (122.0, 100.0), (222.0, 150.0)],
        "declared grid scaled to tblW",
    );
}

/// §17.4.17: a `gridSpan` cell occupies that many grid columns, so two cells
/// spanning 1 and 2 seat a three-column grid exactly and nothing is appended.
#[test]
fn grid_span_counts_toward_seating() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="6000" w:type="dxa"/>"#,
        &grid(&[2000, 2000, 2000]),
        &row(&format!(
            "{}{}",
            cell("a", None, None),
            cell("b", None, Some(2))
        )),
    ));

    // 300 pt over three equal columns: 100 pt each, and the span-2 cell is 200.
    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 100.0), (172.0, 200.0)],
        "gridSpan seats the grid",
    );
}

/// §17.4.15 / §17.4.14: `gridBefore` and `gridAfter` occupy grid columns that
/// hold no cell, and both count toward seating — one cell plus one leading and
/// two trailing skips seat a four-column grid.
#[test]
fn grid_before_and_after_count_toward_seating() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="8000" w:type="dxa"/>"#,
        &grid(&[2000, 2000, 2000, 2000]),
        &row_with(
            r#"<w:gridBefore w:val="1"/><w:gridAfter w:val="2"/>"#,
            &cell("a", None, None),
        ),
    ));

    // 400 pt over four equal columns: 100 pt each, and gridBefore=1 puts the
    // only cell in column 1, so it is drawn from x = 72 + 100.
    assert_cells(
        &first_row_cells(&pages),
        &[(172.0, 100.0)],
        "gridBefore offsets the cell and seats the grid",
    );
}

/// A grid with *more* columns than a row uses still seats every cell in that
/// row — the row simply ends early, which is what `gridAfter` says explicitly
/// and what a producer omitting it leaves implicit.
///
/// Pinned because it is the shape most easily confused with the one the repair
/// exists for, and it is deliberately **not** repaired: every cell has a
/// column, so the seating invariant holds and there is nothing
/// self-contradictory to fix. Note the 4000-twip `tcW` on each cell disagrees
/// with its 2000-twip grid column and is ignored — that is §17.4.71's
/// unresolved conflict, and resolving it needs a **Word reference render**, not
/// this repair.
#[test]
fn a_row_shorter_than_the_grid_is_left_short() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="8000" w:type="dxa"/>"#,
        &grid(&[2000, 2000, 2000, 2000]),
        &row(&format!(
            "{}{}",
            cell("a", Some((4000, "dxa")), None),
            cell("b", Some((4000, "dxa")), None)
        )),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 100.0), (172.0, 100.0)],
        "a short row keeps its grid columns and stops",
    );
}

// ── The repair: a cell the grid cannot seat ──────────────────────────────────
//
// Every test below describes a file whose `<w:tblGrid>` has fewer columns than
// some row has cells. That file contradicts itself — §17.4.71 says the table
// "shall satisfy the shared columns as specified by the tblGrid element", and
// there is no column for the last cell to be satisfied in — so a renderer must
// decide what gives. What gave before was the cell: it was clamped to a
// zero-width slice at the table's right edge, drawn on top of the previous
// cell's border, its text unreadable and its column invisible. Content that
// renders at zero width is gone from the PDF as surely as content drawn off the
// paper, which is the same line `clamp_auto_grid_to_page` is drawn at.

/// The core of it: a cell that exists gets a column. Two declared columns
/// against four cells leaves cells 2 and 3 unseated; both must come out with a
/// width, at distinct positions, inside the table.
///
/// Asserted structurally first — no cell at zero width, no two cells at the
/// same x — because that is the property the repair exists for, and it holds
/// whatever widths the appended columns are given.
#[test]
fn a_cell_the_grid_cannot_seat_still_gets_a_column() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800, 4800]),
        &row(&format!(
            "{}{}{}{}",
            cell("a", Some((2400, "dxa")), None),
            cell("b", Some((2400, "dxa")), None),
            cell("c", Some((2400, "dxa")), None),
            cell("d", Some((2400, "dxa")), None)
        )),
    ));

    let cells = first_row_cells(&pages);
    assert_eq!(cells.len(), 4, "four cells were declared: {cells:?}");
    for (i, (x, w)) in cells.iter().enumerate() {
        assert!(
            *w > 0.0,
            "cell {i} drew at zero width — the grid could not seat it and it \
             was dropped rather than given a column: {cells:?}"
        );
        assert!(x.is_finite(), "cell {i} has no position: {cells:?}");
    }
    for i in 1..cells.len() {
        assert!(
            cells[i].0 > cells[i - 1].0 + 0.01,
            "cells {} and {i} collapsed onto the same column: {cells:?}",
            i - 1
        );
    }
}

/// And the width the appended column gets is the cell's own declared `tcW`
/// — the only width evidence the file offers for a column the grid never
/// mentioned.
///
/// The two unseated cells declare *different* widths (1200 and 3600 twips), so
/// an implementation that appended a uniform column would fail here while
/// passing the structural test above. The declared grid is kept as declared:
/// 4800 + 4800 + 1200 + 3600 = 14400 twips scaled to the 9600-twip `tblW` is a
/// factor of two thirds, giving 160, 160, 40 and 120 pt.
#[test]
fn an_unseated_cell_is_sized_from_its_declared_tcw() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800, 4800]),
        &row(&format!(
            "{}{}{}{}",
            cell("a", Some((4800, "dxa")), None),
            cell("b", Some((4800, "dxa")), None),
            cell("c", Some((1200, "dxa")), None),
            cell("d", Some((3600, "dxa")), None)
        )),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 160.0), (232.0, 160.0), (392.0, 40.0), (432.0, 120.0)],
        "appended columns take the unseated cells' declared tcW",
    );
}

/// §17.4.17: a `gridSpan` reaching past the end of the grid needs *every*
/// column it declares, not just one. Cell `b` spans three columns from column
/// 1, so columns 2 and 3 are missing; its 7200-twip `tcW` covers all three, of
/// which the declared column 1 already supplies 4800, so the remaining 2400 is
/// split between them.
///
/// 4800 + 4800 + 1200 + 1200 = 12000 twips scaled to 9600 is a factor of 0.8:
/// 192 pt for each declared column and 48 pt for each appended one, so the span
/// draws 192 + 48 + 48 = 288 pt from x = 72 + 192.
#[test]
fn a_span_the_grid_cannot_seat_gets_every_column_it_declares() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800, 4800]),
        &row(&format!(
            "{}{}",
            cell("a", Some((4800, "dxa")), None),
            cell("b", Some((7200, "dxa")), Some(3))
        )),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 192.0), (264.0, 288.0)],
        "the span's missing columns share what its tcW leaves over",
    );
}

/// A cell with no `tcW` at all still has to be seated — the repair is about
/// the cell existing, not about it declaring a width. With no evidence to
/// size the appended column, it takes the mean of the declared ones (4800),
/// so three equal columns scale to 160 pt each.
#[test]
fn a_cell_with_no_declared_width_still_gets_a_column() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800, 4800]),
        &row(&format!(
            "{}{}{}",
            cell("a", None, None),
            cell("b", None, None),
            cell("c", None, None)
        )),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 160.0), (232.0, 160.0), (392.0, 160.0)],
        "a cell with no tcW takes the mean of the declared columns",
    );
}

/// A `<w:tblGrid>` with no `<w:gridCol>` at all is the degenerate case of the
/// same rule: *every* cell is unseated, so every column comes from `tcW`.
///
/// This is where the old equal-distribution fallback was worst. 1600/4800/3200
/// twips is a 1:3:2 table, and it drew as three equal columns — the declared
/// widths were not merely overridden, they were never consulted. They sum to
/// the declared `tblW`, so the scale factor is 1: 80, 240 and 160 pt.
#[test]
fn a_table_with_no_grid_takes_its_columns_from_tcw() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        "",
        &row(&format!(
            "{}{}{}",
            cell("a", Some((1600, "dxa")), None),
            cell("b", Some((4800, "dxa")), None),
            cell("c", Some((3200, "dxa")), None)
        )),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 80.0), (152.0, 240.0), (392.0, 160.0)],
        "an absent grid is rebuilt from the declared cell widths",
    );
}

/// Rows disagree about how wide an appended column should be, and the widest
/// claim wins — a column narrower than that would put one of the two rows back
/// where it started, squeezed into less than it declared.
///
/// This is the same direction §17.4.63's own reconciliation runs in, and it is
/// the only choice here that is order-independent: taking the first or the last
/// row's claim would make the result depend on row order, which no reading of
/// §17.4.48 supports.
///
/// The **wider row comes first** deliberately. With the narrow row first, "the
/// last claim wins" and "the widest claim wins" agree, and a mutation replacing
/// the maximum with plain assignment passes — so the order is what gives this
/// test its teeth.
#[test]
fn the_widest_claim_on_an_appended_column_wins() {
    let rows = format!(
        "{}{}",
        row(&format!(
            "{}{}{}",
            cell("a", Some((4800, "dxa")), None),
            cell("b", Some((4800, "dxa")), None),
            cell("c", Some((4800, "dxa")), None)
        )),
        row(&format!(
            "{}{}{}",
            cell("d", Some((4800, "dxa")), None),
            cell("e", Some((4800, "dxa")), None),
            cell("f", Some((2400, "dxa")), None)
        ))
    );
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800, 4800]),
        &rows,
    ));

    // 4800 + 4800 + max(2400, 4800) = 14400 twips over a 9600-twip table.
    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 160.0), (232.0, 160.0), (392.0, 160.0)],
        "the appended column is as wide as the widest row claims",
    );
}

/// §17.4.15: `gridBefore` pushes a row's cells rightward, so it can be what
/// puts the last one past the end of the grid — the cells alone would fit.
///
/// Two cells in a two-column grid are seated; the same two behind a
/// `gridBefore` of 1 are not, and the second has nowhere to go. Counting only
/// the cells would miss it, which is exactly what the mutation check found:
/// every other test here leaves `gridBefore` at 0, so dropping it from the
/// demand broke nothing.
///
/// 4800 + 4800 + 2400 = 12000 twips scaled to 9600 is a factor of 0.8, and
/// `gridBefore` leaves the first column empty, so the cells draw at 192 and 96
/// pt from x = 72 + 192.
#[test]
fn grid_before_can_be_what_unseats_a_cell() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800, 4800]),
        &row_with(
            r#"<w:gridBefore w:val="1"/>"#,
            &format!(
                "{}{}",
                cell("a", Some((4800, "dxa")), None),
                cell("b", Some((2400, "dxa")), None)
            ),
        ),
    ));

    assert_cells(
        &first_row_cells(&pages),
        &[(264.0, 192.0), (456.0, 96.0)],
        "gridBefore counts toward the demand",
    );
}

/// The gate is about **cells**, not about grid columns in the abstract.
///
/// §17.4.14 `gridAfter` declares trailing columns that hold no cell, so a grid
/// too short to contain them loses nothing — there is no content in them to
/// lose, and the repair's whole justification is content that would otherwise
/// not be drawn. Repairing here would instead be a geometry change on
/// speculation: it would halve this cell, which today fills the table and is
/// perfectly legible.
///
/// Pinned so the narrowness is a decision with a test behind it. What Word does
/// with a `gridAfter` that overruns the grid is open, and a **Word reference
/// render** is what would settle it.
#[test]
fn a_grid_after_the_grid_cannot_hold_is_not_repaired() {
    let pages = layout(&table_doc(
        r#"<w:tblW w:w="9600" w:type="dxa"/>"#,
        &grid(&[4800]),
        &row_with(r#"<w:gridAfter w:val="1"/>"#, &cell("a", None, None)),
    ));

    // One declared column, scaled to the whole 9600-twip (480 pt) table.
    assert_cells(
        &first_row_cells(&pages),
        &[(72.0, 480.0)],
        "gridAfter must not append a column",
    );
}
