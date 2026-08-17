//! §17.4.36 `insideV` against §17.4.15 `gridBefore` and §17.4.14 `gridAfter` —
//! what is painted on the vertical edge where a row's cells stop short of the
//! table's grid.
//!
//! # The question
//!
//! `gridBefore` leaves the leftmost grid columns of a row blank: no `<w:tc>`
//! covers them. The row's first cell then has a grid column to its left but no
//! *cell* to its left, and §17.4.36 defines `insideV` as the border on the
//! table's "interior vertical edges" without defining "interior". Three readings
//! survive that wording:
//!
//! * **A** — interior is a property of the **grid**. There are columns to the
//!   left, so the edge takes `insideV`. This is what the engine does today:
//!   `table::borders::resolve_cell_effective_borders` keys `is_first_col` on the
//!   cell's absolute grid column, past `gridBefore`.
//! * **B** — interior is a property of the row's **cells**. Nothing is to the
//!   left, so this is the row's leading edge and takes the table's `w:left`.
//! * **C** — the edge is not painted at all.
//!
//! # What is settled, and what these tests therefore are
//!
//! A Word render of `test-files/bidi-visual-table.docx` **rules out A**: Word
//! paints nothing on that edge. It cannot separate B from C, because that
//! table's `w:left` is `nil` and both then predict nothing.
//!
//! So the assertions below split in two, and the names say which is which:
//!
//! * The `legend_*` and `reaches_*` tests assert properties that hold under
//!   every reading, and they are what keeps the fixture honest — a probe whose
//!   three borders became indistinguishable would still render, still pass a
//!   loose test, and quietly stop being able to answer anything.
//! * The `today_*` tests are **characterization**. They record reading A because
//!   that is what the engine does, not because it is right; one of them
//!   (`today_the_nil_table_paints_the_gap_edges`) records behaviour Word is
//!   already known to contradict. They exist so that a fix shows up as a
//!   deliberate edit here rather than as a silent change in output, and every
//!   one of them is expected to be inverted when `test-files/grid-gap-borders.docx`
//!   is measured in Word.
//!
//! Nothing here asserts a glyph metric, and the fixture states its own font, so
//! the two renders are directly comparable — see the fixture script.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

/// A painted rectangle, as `(x, y, w, h)`.
type Rect = (f32, f32, f32, f32);

const RED_LEFT: (u8, u8, u8) = (0xC0, 0x00, 0x00);
const GREEN_RIGHT: (u8, u8, u8) = (0x00, 0xB0, 0x50);
const BLUE_INSIDE_V: (u8, u8, u8) = (0x00, 0x70, 0xC0);

fn layout() -> Vec<LayoutedPage> {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/test-files/grid-gap-borders.docx"
    );
    let bytes = std::fs::read(path).expect("read the fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse the fixture");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Every rect of one colour, in paint order.
fn rects(pages: &[LayoutedPage], colour: (u8, u8, u8)) -> Vec<Rect> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color } if (color.r, color.g, color.b) == colour => Some((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.size.width.raw(),
                rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect()
}

/// The two boxes a cell fill paints — one per table — ordered by y, so `.0` is
/// the visible-border table and `.1` is the `nil` one.
///
/// Both tables use the same fills deliberately: it is what lets every assertion
/// below name a cell rather than an index, in a document that holds the same
/// four rows twice.
fn cell(pages: &[LayoutedPage], fill: (u8, u8, u8)) -> (Rect, Rect) {
    let mut found = rects(pages, fill);
    assert_eq!(
        found.len(),
        2,
        "{fill:?} must paint once per table — the fixture holds the same rows twice"
    );
    found.sort_by(|a, b| a.1.total_cmp(&b.1));
    (found[0], found[1])
}

/// Is a vertical border of `colour` painted across `x`, within `cell`'s rows?
/// Returns its width.
///
/// The match is **span-containing**, not origin-comparing: a border is drawn
/// inside the cell it belongs to, so a leading edge's rect starts at `x` while a
/// trailing edge's ends there (`x - w`). Comparing origins would need a
/// tolerance as wide as the thickest border, which would then also swallow the
/// neighbouring edge of a thin cell.
///
/// `rh > rw` excludes the junction squares that `table_border_corners.rs` fills
/// where a vertical border meets a horizontal one — a 3pt border crossing a 1pt
/// rule leaves a 3×1 rect of the vertical's own colour, which is wider than it
/// is tall and is not an edge.
fn vertical_at(pages: &[LayoutedPage], colour: (u8, u8, u8), x: f32, cell: Rect) -> Option<f32> {
    let (_, cy, _, ch) = cell;
    let mid = cy + ch * 0.5;
    rects(pages, colour)
        .into_iter()
        .find(|&(rx, ry, rw, rh)| {
            rh > rw && rx - 0.6 <= x && x <= rx + rw + 0.6 && ry <= mid && ry + rh >= mid
        })
        .map(|(_, _, rw, _)| rw)
}

// ── the fixture's own integrity ─────────────────────────────────────────────

/// Row 1 spans the whole grid, so it shows all three vertical borders at once —
/// and they must be tellable apart, or the probe cannot answer anything.
///
/// This is the test that fails if someone "tidies" the fixture into one border
/// colour or one weight. Everything else here would still pass.
#[test]
fn legend_row_shows_three_distinguishable_vertical_borders() {
    let pages = layout();
    let (a, _) = cell(&pages, (0xF8, 0xCB, 0xAD));
    let (b, _) = cell(&pages, (0xC6, 0xE0, 0xB4));
    let (c, _) = cell(&pages, (0xBD, 0xD7, 0xEE));

    let left = vertical_at(&pages, RED_LEFT, a.0, a).expect("w:left at the table's left edge");
    let right =
        vertical_at(&pages, GREEN_RIGHT, c.0 + c.2, c).expect("w:right at the table's right edge");
    let inside_ab =
        vertical_at(&pages, BLUE_INSIDE_V, a.0 + a.2, a).expect("insideV between A and B");
    let inside_bc =
        vertical_at(&pages, BLUE_INSIDE_V, b.0 + b.2, b).expect("insideV between B and C");

    assert_eq!(left, 3.0, "w:sz=24 is 3pt");
    assert_eq!(right, 3.0);
    assert_eq!(inside_ab, 1.0, "w:sz=8 is 1pt");
    assert_eq!(inside_bc, 1.0);
    assert_ne!(
        left, inside_ab,
        "outer and interior borders must differ in weight, not only in colour"
    );
}

/// The control that bounds any fix: a cell that genuinely *does* reach the
/// table's edge still takes the outer border there.
///
/// Cell `E` starts at grid column 0 (its row's gap is at the far end) and cell
/// `D` ends at the last column (its gap is at the near end), so between them
/// they pin both outer edges. A fix that replaced `insideV` with the outer
/// border everywhere, or that suppressed every edge of a gapped row, would fail
/// this — which is exactly what it is for.
#[test]
fn a_cell_that_reaches_the_table_edge_still_takes_the_outer_border() {
    let pages = layout();
    let (d, _) = cell(&pages, (0xFF, 0xE6, 0x99));
    let (e, _) = cell(&pages, (0xD9, 0xD2, 0xE9));

    assert_eq!(
        vertical_at(&pages, GREEN_RIGHT, d.0 + d.2, d),
        Some(3.0),
        "D reaches the last grid column, so its trailing edge is the table's"
    );
    assert_eq!(
        vertical_at(&pages, RED_LEFT, e.0, e),
        Some(3.0),
        "E starts at grid column 0, so its leading edge is the table's"
    );
}

// ── characterization: reading A, which Word contradicts ─────────────────────

/// §17.4.15: today the edge facing a `gridBefore` gap is painted `insideV`.
///
/// **Characterization, and known to be wrong.** Word paints nothing on this edge
/// in `bidi-visual-table.docx`. Whether the correct line here is the 3pt red
/// `w:left` (reading B) or nothing at all (reading C) is what a Word render of
/// this fixture settles — the two are distinguishable here precisely because
/// this table's `w:left` is visible.
#[test]
fn today_a_grid_before_gap_edge_takes_inside_v() {
    let pages = layout();
    let (d, _) = cell(&pages, (0xFF, 0xE6, 0x99));

    assert_eq!(
        vertical_at(&pages, BLUE_INSIDE_V, d.0, d),
        Some(1.0),
        "reading A: the gap-facing leading edge is treated as interior"
    );
    assert_eq!(
        vertical_at(&pages, RED_LEFT, d.0, d),
        None,
        "…and the table's own w:left is not painted there — reading B's answer"
    );
}

/// §17.4.14: the mirror image, at the trailing edge of a `gridAfter` row.
///
/// **Characterization.** `is_last_col` in `resolve_cell_effective_borders` has
/// the same shape as `is_first_col`, so this is the same defect at the other
/// end. No Word render has been taken of it — the reported case was
/// `gridBefore` — which is why the fixture asks both.
#[test]
fn today_a_grid_after_gap_edge_takes_inside_v() {
    let pages = layout();
    let (e, _) = cell(&pages, (0xD9, 0xD2, 0xE9));

    assert_eq!(
        vertical_at(&pages, BLUE_INSIDE_V, e.0 + e.2, e),
        Some(1.0),
        "reading A at the trailing edge"
    );
    assert_eq!(
        vertical_at(&pages, GREEN_RIGHT, e.0 + e.2, e),
        None,
        "…and not the table's w:right"
    );
}

/// A row gapped at **both** ends paints `insideV` on both.
///
/// **Characterization.** Its value is that a fix reaching one end and not the
/// other passes the two tests above in whichever order it was written and fails
/// here.
#[test]
fn today_a_row_gapped_at_both_ends_takes_inside_v_twice() {
    let pages = layout();
    let (f, _) = cell(&pages, (0xF4, 0xCC, 0xCC));

    assert_eq!(vertical_at(&pages, BLUE_INSIDE_V, f.0, f), Some(1.0));
    assert_eq!(vertical_at(&pages, BLUE_INSIDE_V, f.0 + f.2, f), Some(1.0));
}

/// The reported case, isolated: the same rows with `w:left`/`w:right` set to
/// `nil` still get a line at both gap edges.
///
/// **Characterization of a known defect.** This is `bidi-visual-table.docx`'s
/// symptom with no `w:bidiVisual` and no Hebrew involved — Word draws nothing at
/// either edge, and the engine draws 1pt blue. When that is fixed this
/// assertion becomes `None`, and it is the one to change first.
#[test]
fn today_the_nil_table_paints_the_gap_edges() {
    let pages = layout();
    let (_, d) = cell(&pages, (0xFF, 0xE6, 0x99));
    let (_, e) = cell(&pages, (0xD9, 0xD2, 0xE9));

    assert_eq!(
        vertical_at(&pages, BLUE_INSIDE_V, d.0, d),
        Some(1.0),
        "Word paints nothing here"
    );
    assert_eq!(
        vertical_at(&pages, BLUE_INSIDE_V, e.0 + e.2, e),
        Some(1.0),
        "…nor here"
    );
    // The control for the pair: `nil` really did reach the outer edges, so the
    // lines above are `insideV` reaching inward rather than `w:left` surviving.
    //
    // Stated as "no red or green anywhere in the second table" rather than as a
    // count, because a count also has to know about the junction squares
    // `table_border_corners.rs` fills — the red `w:left` crossing a 1pt `insideH`
    // leaves a third red rect that is not an edge.
    let (_, a2) = cell(&pages, (0xF8, 0xCB, 0xAD));
    for (name, colour) in [("w:left", RED_LEFT), ("w:right", GREEN_RIGHT)] {
        assert!(
            rects(&pages, colour).iter().all(|&(_, y, _, _)| y < a2.1),
            "{name} must not be painted in the nil table, which starts at y={}",
            a2.1
        );
    }
}
