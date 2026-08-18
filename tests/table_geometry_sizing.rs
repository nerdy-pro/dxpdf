//! §17.4.63 / §17.4.48 / §17.4.80 — how wide a table comes out, and how tall
//! its rows are.
//!
//! `build_table` resolves `w:tblW` into a target width, scales the declared
//! `<w:tblGrid>` to it, and — for a full-width table that is *not* centred or
//! right-aligned — widens it by the table's own cell margins so cell text lines
//! up with body text. Each of those is a number on the page, and each was
//! reachable only through the unit tests of the function that computes it,
//! which cannot see whether the number survives to the paper.
//!
//! Row height is the same story from the other end: §17.4.80's `exact` had no
//! layout test at any level, only a parse test, so nothing said what an
//! author's declared row height does to the block below the table.
//!
//! The expectations here are arithmetic over the fixture. The page is the
//! default Letter sheet — 612 × 792 pt with 1-inch margins, so the text column
//! is 468 pt wide starting at x = 72 — and every width below is derived from
//! that and from the twips in the XML.

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

fn doc(body: &str) -> Vec<u8> {
    make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}</w:body>
</w:document>"#
    ))
}

fn layout(bytes: &[u8]) -> Vec<LayoutedPage> {
    let parsed = dxpdf::docx::parse(bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// The table's **box** — the span `w:tblW` sizes and `w:tblInd` places —
/// recovered from the rectangles actually drawn.
///
/// The two are not the same span. §17.4.66 straddles a table's own edges
/// (measured against `test-files/border-outer-box.docx`), so the ink runs half a
/// border wider than the box at each end, and the box is recovered by taking
/// that half back off both.
///
/// The half is taken from the render rather than written down: it is half the
/// width of the *narrowest* rectangle beginning at the leftmost x, which is the
/// left border itself. Narrowest and not simply leftmost, because the top and
/// bottom borders start at that same x and run the width of the table.
/// Self-deriving, so a case declaring a different `w:sz` needs nothing changed
/// here, and every assertion below stays about width and placement.
fn drawn_span(pages: &[LayoutedPage]) -> (f32, f32) {
    let rects: Vec<(f32, f32)> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => Some((rect.origin.x.raw(), rect.size.width.raw())),
            _ => None,
        })
        .collect();
    assert!(!rects.is_empty(), "no table rectangles were drawn");

    let left = rects.iter().map(|r| r.0).fold(f32::INFINITY, f32::min);
    let right = rects
        .iter()
        .map(|r| r.0 + r.1)
        .fold(f32::NEG_INFINITY, f32::max);
    let border = rects
        .iter()
        .filter(|r| (r.0 - left).abs() < 0.01)
        .map(|r| r.1)
        .fold(f32::INFINITY, f32::min);
    (left + border / 2.0, right - border / 2.0)
}

/// x of every text run, in draw order.
fn text_xs(pages: &[LayoutedPage]) -> Vec<f32> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { position, .. } => Some(position.x.raw()),
            _ => None,
        })
        .collect()
}

/// y of the first text command whose content equals `needle`.
fn y_of(pages: &[LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == needle => Some(position.y.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text command {needle:?}"))
}

const ALL_BORDERS: &str = r#"<w:tblBorders>
    <w:top w:val="single" w:sz="4" w:color="000000"/>
    <w:left w:val="single" w:sz="4" w:color="000000"/>
    <w:bottom w:val="single" w:sz="4" w:color="000000"/>
    <w:right w:val="single" w:sz="4" w:color="000000"/>
    <w:insideH w:val="single" w:sz="4" w:color="000000"/>
    <w:insideV w:val="single" w:sz="4" w:color="000000"/>
  </w:tblBorders>"#;

/// Half the 0.5 pt (`w:sz="4"`) border every table in this file declares.
///
/// A table's box begins this far to the **left** of its indent, and `build_table`
/// does that deliberately: §17.4.66 straddles the leading edge, so half the
/// border would otherwise be charged to the first cell and push its text right.
/// The shift puts the text back on the indent — which is what `w:tblInd`
/// measures to, per `test-files/border-outer-box.docx`. So a table "at the
/// margin" has its box at `72.0 - LEADING_HALF` and its first cell's text at 72,
/// and the assertions below say so rather than rounding it away.
///
/// Two cases below do not add it, and neither is an oversight. The full-width
/// left-aligned table is already pulled left by `max(cell margin, half the
/// leading border)`, and a 5.4 pt margin beats a 0.25 pt half outright. The
/// **centred** table does not carry an indent at all — `w:jc` places the box as
/// a unit within the width offered — so there is nothing for the shift to apply
/// to and its box lands on the margin exactly.
const LEADING_HALF: f32 = 0.25;

/// A one-row two-column table. `tbl_pr` carries whatever the case declares
/// beyond the borders; `grid` is the `<w:gridCol>` list verbatim.
fn two_col_table(tbl_pr: &str, grid: &str) -> Vec<u8> {
    doc(&format!(
        r#"<w:tbl>
      <w:tblPr>{tbl_pr}{ALL_BORDERS}</w:tblPr>
      <w:tblGrid>{grid}</w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>L</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>"#
    ))
}

// ── §17.4.63: `w:tblW` decides the width the grid is scaled to ──────────────

/// §17.4.63: a `dxa` `tblW` states a preferred width, and §17.4.48's grid
/// states the *proportions*. The two disagree here by construction — the grid
/// declares 8000 twips and the table 4000 — and the resolution is a scale, not
/// a truncation: the table comes out at its declared 200 pt with the columns
/// still in their declared 1:3 ratio.
///
/// The column boundary is read as the gap between the two cells' text rather
/// than as an absolute x, so it does not depend on the inset each cell's text
/// carries — as long as both carry the same one. That is what the 5 pt
/// `tblCellMar` is for: §17.4.66 collapsing gives the *left* cell the shared
/// vertical border and the right cell none, and a border wider than the margin
/// behind it displaces that cell's content by the excess. At 5 pt of margin
/// against a 0.5 pt border there is no excess on either side, so the difference
/// is the first column's width and nothing else.
#[test]
fn a_dxa_table_width_scales_the_declared_grid_rather_than_truncating_it() {
    let pages = layout(&two_col_table(
        r#"<w:tblW w:w="4000" w:type="dxa"/>
           <w:tblCellMar>
             <w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>
           </w:tblCellMar>"#,
        r#"<w:gridCol w:w="2000"/><w:gridCol w:w="6000"/>"#,
    ));

    let (left, right) = drawn_span(&pages);
    assert!(
        (left - (72.0 - LEADING_HALF)).abs() < 0.01
            && (right - (272.0 - LEADING_HALF)).abs() < 0.01,
        "4000 twips is 200 pt from the 72 pt margin, got {left:.2}..{right:.2}"
    );

    let xs = text_xs(&pages);
    assert_eq!(xs.len(), 2, "one text run per cell");
    assert!(
        (xs[1] - xs[0] - 50.0).abs() < 0.01,
        "the 2000:6000 grid must scale to 50:150 pt, but column 0 came out \
         {:.2} pt wide",
        xs[1] - xs[0]
    );
}

/// §17.4.63 / §17.18.90: a `pct` `tblW` is on the fiftieths-of-a-percent scale,
/// so `w="2500"` is 50% — of the 468 pt offered here, 234 pt.
///
/// Below `FULL` there is no cell-margin extension to apply (see the pair
/// below), so this is the plain fraction and the table starts at the margin.
#[test]
fn a_pct_table_width_is_that_fraction_of_the_offered_width() {
    let pages = layout(&two_col_table(
        r#"<w:tblW w:w="2500" w:type="pct"/>"#,
        r#"<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>"#,
    ));

    let (left, right) = drawn_span(&pages);
    assert!(
        (left - (72.0 - LEADING_HALF)).abs() < 0.01,
        "a table below full width starts at the margin, got {left:.2}"
    );
    assert!(
        (right - left - 234.0).abs() < 0.01,
        "2500/5000 of 468 pt is 234 pt, got {:.2}",
        right - left
    );
}

// ── §17.4.63 / §17.4.50: the full-width extension, and who gets it ──────────

/// A full-width **left-aligned** table is widened by its own cell margins and
/// pulled left by the left one, so that cell *text* — not the cell box — lines
/// up with the body text on both sides. With 108-twip (5.4 pt) margins that is
/// a 478.8 pt table starting at x = 66.6, whose text therefore starts at 72.
///
/// A **centred** one is not, and that is the whole of `extends_for_alignment`:
/// a centred table is placed as a unit, so widening it would only push it out
/// by half the margins on each side and leave two stacked tables with different
/// cell margins visibly different widths. It stays exactly the 468 pt offered.
///
/// The two cases are the same document but for `w:jc`, so nothing else can
/// account for the 10.8 pt between them.
#[test]
fn a_full_width_table_extends_by_its_cell_margins_only_when_it_is_not_centred() {
    let cell_mar = r#"<w:tblCellMar>
          <w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>"#;
    let full = |jc: &str| {
        two_col_table(
            &format!(r#"<w:tblW w:w="5000" w:type="pct"/>{jc}{cell_mar}"#),
            r#"<w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>"#,
        )
    };

    let (left_l, right_l) = drawn_span(&layout(&full("")));
    assert!(
        (right_l - left_l - 478.8).abs() < 0.02,
        "a left-aligned full-width table is 468 + 2 × 5.4 pt wide, got {:.2}",
        right_l - left_l
    );
    assert!(
        (left_l - 66.6).abs() < 0.02,
        "…and is pulled left by one cell margin, to 66.6, not {left_l:.2}"
    );
    assert!(
        (text_xs(&layout(&full(""))).first().copied().unwrap() - 72.0).abs() < 0.02,
        "which is the point: its first cell's text lands on the margin"
    );

    let (left_c, right_c) = drawn_span(&layout(&full(r#"<w:jc w:val="center"/>"#)));
    assert!(
        (right_c - left_c - 468.0).abs() < 0.02,
        "a centred full-width table takes the text column verbatim, got {:.2}",
        right_c - left_c
    );
    assert!(
        (left_c - 72.0).abs() < 0.02,
        "…and centring 468 pt in 468 pt leaves it at the margin, not {left_c:.2}"
    );
}

// ── §17.4.80: `w:trHeight` ──────────────────────────────────────────────────

/// A table of one row holding `lines` paragraphs, with `tr_pr` verbatim in the
/// row's `<w:trPr>`, between two paragraphs whose positions are the measurement.
fn row_height_doc(tr_pr: &str, lines: usize) -> Vec<u8> {
    let content: String = (0..lines)
        .map(|i| format!("<w:p><w:r><w:t>line{i}</w:t></w:r></w:p>"))
        .collect();
    doc(&format!(
        r#"<w:p><w:r><w:t>before</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w="8000" w:type="dxa"/>{ALL_BORDERS}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
      <w:tr><w:trPr>{tr_pr}</w:trPr><w:tc>{content}</w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>after</w:t></w:r></w:p>"#
    ))
}

/// y of the paragraph below the table — the only thing that shows how tall the
/// row came out, since the row itself draws nothing that names its height.
fn after_y(tr_pr: &str, lines: usize) -> f32 {
    y_of(&layout(&row_height_doc(tr_pr, lines)), "after")
}

/// §17.4.80 `hRule="exact"`: *"the height of the row shall be exactly the value
/// specified"*. So the block after the table sits in the same place whether the
/// row holds one line or six, and the `auto` control shows that six lines is
/// otherwise a very different table.
///
/// This asserts the row's **height** and deliberately nothing about the six
/// lines that do not fit in it. dxpdf does not clip them — they are painted
/// past the row's box and overprint whatever follows — and that is an open
/// question rather than a decision: §17.4.80 states the height and says nothing
/// about the overflow, LibreOffice clips, and no Word render is on record. A
/// green test asserting the current overflow would pin a guess.
#[test]
fn an_exact_row_height_is_the_declared_height_whatever_the_row_holds() {
    let exact = r#"<w:trHeight w:val="1000" w:hRule="exact"/>"#;
    assert!(
        (after_y(exact, 1) - after_y(exact, 6)).abs() < 0.01,
        "an exact row grew with its content: {:.2} vs {:.2}",
        after_y(exact, 1),
        after_y(exact, 6)
    );
    assert!(
        after_y("", 6) - after_y("", 1) > 1.0,
        "the control must show six lines is taller than one, or the equality \
         above is vacuous"
    );
}

/// …and the declared value is taken literally: 1000 twips more of row is
/// exactly 50 pt more of page. Both rows hold the same single line, so the
/// content cannot account for any of the difference.
#[test]
fn an_exact_row_height_moves_the_following_block_by_exactly_the_declared_amount() {
    let delta = after_y(r#"<w:trHeight w:val="2000" w:hRule="exact"/>"#, 1)
        - after_y(r#"<w:trHeight w:val="1000" w:hRule="exact"/>"#, 1);
    assert!(
        (delta - 50.0).abs() < 0.01,
        "1000 extra twips is 50 pt; the page moved by {delta:.2}"
    );
}

/// §17.4.80 `hRule="atLeast"` is the same value read as a **floor**, which is
/// the distinction `exact` exists against. Pinned as two equalities that
/// bracket it: with content shorter than the minimum an `atLeast` row is
/// indistinguishable from an `exact` one, and with content taller it is
/// indistinguishable from a row that declared no height at all.
///
/// Neither equality can hold if the two rules are conflated — the second fails
/// under `exact` semantics and the first under `auto`.
#[test]
fn an_at_least_row_height_is_a_floor_and_an_exact_one_is_not() {
    let at_least = r#"<w:trHeight w:val="1000" w:hRule="atLeast"/>"#;
    let exact = r#"<w:trHeight w:val="1000" w:hRule="exact"/>"#;

    assert!(
        (after_y(at_least, 1) - after_y(exact, 1)).abs() < 0.01,
        "under its minimum, atLeast must size exactly like exact: {:.2} vs {:.2}",
        after_y(at_least, 1),
        after_y(exact, 1)
    );
    assert!(
        (after_y(at_least, 6) - after_y("", 6)).abs() < 0.01,
        "over its minimum, atLeast must size like no rule at all: {:.2} vs {:.2}",
        after_y(at_least, 6),
        after_y("", 6)
    );
    assert!(
        (after_y(exact, 6) - after_y("", 6)).abs() > 1.0,
        "…and exact must not, or the pair above says nothing"
    );
}
