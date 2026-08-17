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
//!   left, so the edge takes `insideV`.
//! * **B** — interior is a property of the row's **cells**. Nothing is to the
//!   left, so this is the row's leading edge and takes the table's `w:left`.
//! * **C** — no *table-level* border reaches the edge at all. It is not
//!   interior, because no cell faces it; and it is not the table's boundary,
//!   which is 50pt further left. A cell that wants a line there still says
//!   `w:tcBorders`.
//!
//! # Which one is right: measured, not argued
//!
//! **Word paints the table's own `w:left` and `w:right` there — reading B.**
//! Rendering `test-files/grid-gap-borders.docx` shows a 3pt red line on the
//! leading edge of `D`, `F` and `G`, and a 3pt green one on the trailing edge of
//! `E` and `F`: every gap-facing edge in the fixture, each in the table's outer
//! colour rather than `insideV`'s blue.
//!
//! That also explains the observation this fixture was built from. In
//! `test-files/bidi-visual-table.docx` the same edge is bare — but that table's
//! `w:left` is `nil`, which is exactly what B predicts. One rule now covers both
//! renders, where before they needed two.
//!
//! So a row's **first cell takes `w:left` and its last takes `w:right`**,
//! whether or not those cells reach the table's grid. §17.4.66 resolves a
//! cell edge against "cell borders and outer table borders", and a row's first
//! cell has no cell facing its leading edge, so the table's border is what is
//! left to face — `w:gridBefore` moves where that edge *is*, not what it is.
//!
//! Two earlier readings of this file were wrong and are recorded because the
//! way they failed is worth keeping. **A** — `insideV`, because grid columns
//! exist to the left — was refuted by the `nil` render. **C** — nothing at all,
//! because §17.4.35 places `w:left` "around the table" and the edge is 50pt
//! inside it — was argued from the spec's wording with no render behind it, fit
//! the one measurement then available, and is refuted by this one. The geometry
//! that made C plausible is real; it just is not what Word does.
//!
//! # How these tests are written
//!
//! Every assertion names a cell by its fill and asks what is painted at one of
//! its two vertical edges, so none of them knows a page origin or a glyph
//! metric. The fixture states its own font, so its render is directly comparable
//! to Word's — see the fixture script.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

/// A painted rectangle, as `(x, y, w, h)`.
type Rect = (f32, f32, f32, f32);

/// An RGB triple, as the draw commands carry it.
type Colour = (u8, u8, u8);

/// The horizontal extent a border occupies, as `(x0, x1)`.
type Band = (f32, f32);

const RED_LEFT: Colour = (0xC0, 0x00, 0x00);
const GREEN_RIGHT: Colour = (0x00, 0xB0, 0x50);
const BLUE_INSIDE_V: Colour = (0x00, 0x70, 0xC0);
const GREY_HORIZONTAL: Colour = (0x80, 0x80, 0x80);
const EVERY_BORDER: [Colour; 3] = [RED_LEFT, GREEN_RIGHT, BLUE_INSIDE_V];

// The fixture's cell fills. Named rather than inlined because most tests below
// name two of them and a transposed hex pair would otherwise read as a geometry
// failure.
const A: Colour = (0xF8, 0xCB, 0xAD);
const B: Colour = (0xC6, 0xE0, 0xB4);
const C: Colour = (0xBD, 0xD7, 0xEE);
const D: Colour = (0xFF, 0xE6, 0x99);
const E: Colour = (0xD9, 0xD2, 0xE9);
const F: Colour = (0xF4, 0xCC, 0xCC);
const G: Colour = (0xD9, 0xEA, 0xD3);
const H: Colour = (0xCF, 0xE2, 0xF3);

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
fn rects(pages: &[LayoutedPage], colour: Colour) -> Vec<Rect> {
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
/// five rows twice.
fn cell(pages: &[LayoutedPage], fill: Colour) -> (Rect, Rect) {
    let mut found = rects(pages, fill);
    assert_eq!(
        found.len(),
        2,
        "{fill:?} must paint once per table — the fixture holds the same rows twice"
    );
    found.sort_by(|a, b| a.1.total_cmp(&b.1));
    (found[0], found[1])
}

/// The band a vertical border of `colour` occupies at `x`, within `cell`'s rows.
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
fn band_at(pages: &[LayoutedPage], colour: Colour, x: f32, cell: Rect) -> Option<Band> {
    let (_, cy, _, ch) = cell;
    let mid = cy + ch * 0.5;
    rects(pages, colour)
        .into_iter()
        .find(|&(rx, ry, rw, rh)| {
            rh > rw && rx - 0.6 <= x && x <= rx + rw + 0.6 && ry <= mid && ry + rh >= mid
        })
        .map(|(rx, _, rw, _)| (rx, rx + rw))
}

/// The width of whatever vertical border is painted at `x`, of any colour, or
/// `None`. Used where the claim is "nothing at all is painted here", which has
/// to range over every border the fixture can produce.
fn anything_at(pages: &[LayoutedPage], x: f32, cell: Rect) -> Option<(Colour, Band)> {
    EVERY_BORDER
        .iter()
        .find_map(|&c| band_at(pages, c, x, cell).map(|band| (c, band)))
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
    let (a, _) = cell(&pages, A);
    let (b, _) = cell(&pages, B);
    let (c, _) = cell(&pages, C);

    let width = |band: Option<Band>| band.map(|(x0, x1)| x1 - x0);
    let left = width(band_at(&pages, RED_LEFT, a.0, a)).expect("w:left at the table's left edge");
    let right = width(band_at(&pages, GREEN_RIGHT, c.0 + c.2, c))
        .expect("w:right at the table's right edge");
    let inside_ab =
        width(band_at(&pages, BLUE_INSIDE_V, a.0 + a.2, a)).expect("insideV between A and B");
    let inside_bc =
        width(band_at(&pages, BLUE_INSIDE_V, b.0 + b.2, b)).expect("insideV between B and C");

    assert_eq!(left, 3.0, "w:sz=24 is 3pt");
    assert_eq!(right, 3.0);
    assert_eq!(inside_ab, 1.0, "w:sz=8 is 1pt");
    assert_eq!(inside_bc, 1.0);
    assert_ne!(
        left, inside_ab,
        "outer and interior borders must differ in weight, not only in colour"
    );
}

// ── what a gap-facing edge takes ────────────────────────────────────────────

/// §17.4.15: the edge facing a `gridBefore` gap takes the table's `w:left`.
///
/// Word paints a 3pt red line on cell `D`'s leading edge — the outer border, in
/// the outer colour, 50pt inside the table's own left side. Asserted by weight
/// *and* colour, which is what the fixture's palette is for: `insideV` here is
/// 1pt blue, so a renderer treating the edge as interior fails on both counts.
#[test]
fn a_grid_before_gap_edge_takes_the_tables_left_border() {
    let pages = layout();
    let (d, _) = cell(&pages, D);

    assert_eq!(
        band_at(&pages, RED_LEFT, d.0, d).map(|(x0, x1)| x1 - x0),
        Some(3.0),
        "D's leading edge takes w:left, as Word renders it"
    );
    assert_eq!(
        band_at(&pages, BLUE_INSIDE_V, d.0, d),
        None,
        "…and not insideV, the reading the nil render already refuted"
    );
}

/// §17.4.14: the mirror image — a `gridAfter` row's trailing edge takes
/// `w:right`, which Word paints green on cell `E`.
///
/// The pair is what stops a fix reaching only the end that was reported first.
#[test]
fn a_grid_after_gap_edge_takes_the_tables_right_border() {
    let pages = layout();
    let (e, _) = cell(&pages, E);

    assert_eq!(
        band_at(&pages, GREEN_RIGHT, e.0 + e.2, e).map(|(x0, x1)| x1 - x0),
        Some(3.0)
    );
    assert_eq!(band_at(&pages, BLUE_INSIDE_V, e.0 + e.2, e), None);
}

/// A row gapped at **both** ends takes the outer border at both — red on the
/// left of `F`, green on its right.
///
/// Its value is that a fix reaching one end and not the other passes the two
/// tests above in whichever order it was written and fails here.
#[test]
fn a_row_gapped_at_both_ends_takes_an_outer_border_at_each() {
    let pages = layout();
    let (f, _) = cell(&pages, F);

    assert_eq!(
        band_at(&pages, RED_LEFT, f.0, f).map(|(x0, x1)| x1 - x0),
        Some(3.0),
        "F's leading edge"
    );
    assert_eq!(
        band_at(&pages, GREEN_RIGHT, f.0 + f.2, f).map(|(x0, x1)| x1 - x0),
        Some(3.0),
        "F's trailing edge"
    );
}

/// The rule is about the row's cells, not the grid: `G` is its row's first cell
/// and takes `w:left` even though `H` follows it and the row reaches the grid's
/// end.
///
/// Without this, "the row's first cell" and "a row with exactly one cell" are
/// indistinguishable — every other gapped row in the fixture holds a single
/// cell, so a fix keyed on that would pass everything else here.
#[test]
fn the_first_cell_of_a_gapped_row_takes_the_left_border_even_with_cells_after_it() {
    let pages = layout();
    let (g, _) = cell(&pages, G);

    assert_eq!(
        band_at(&pages, RED_LEFT, g.0, g).map(|(x0, x1)| x1 - x0),
        Some(3.0),
        "G is its row's first cell, so its leading edge is the table's"
    );
}

/// The reported case, isolated: the same rows with `w:left`/`w:right` set to
/// `nil`. The rule is the same one — the gap-facing edge takes the table's outer
/// border — and here that border is `nil`, so nothing is painted.
///
/// This is `bidi-visual-table.docx`'s symptom with no `w:bidiVisual` and no
/// Hebrew involved, and the case that makes the pair with the tests above: one
/// rule has to produce a line in the first table and none in this one.
#[test]
fn the_nil_table_paints_nothing_at_either_gap_edge() {
    let pages = layout();
    let (_, d) = cell(&pages, D);
    let (_, e) = cell(&pages, E);

    assert_eq!(anything_at(&pages, d.0, d), None);
    assert_eq!(anything_at(&pages, e.0 + e.2, e), None);
}

// ── what the rule must not take away ────────────────────────────────────────

/// The control that bounds the fix at the table's edges: a cell that genuinely
/// *does* reach the table's boundary still takes the outer border there.
///
/// Cell `E` starts at grid column 0 (its row's gap is at the far end) and cell
/// `D` ends at the last column (its gap is at the near end), so between them
/// they pin both outer edges. A fix that suppressed every edge of a gapped row
/// would fail this — which is exactly what it is for.
#[test]
fn a_cell_that_reaches_the_table_edge_still_takes_the_outer_border() {
    let pages = layout();
    let (d, _) = cell(&pages, D);
    let (e, _) = cell(&pages, E);

    let width = |band: Option<Band>| band.map(|(x0, x1)| x1 - x0);
    assert_eq!(
        width(band_at(&pages, GREEN_RIGHT, d.0 + d.2, d)),
        Some(3.0),
        "D reaches the last grid column, so its trailing edge is the table's"
    );
    assert_eq!(
        width(band_at(&pages, RED_LEFT, e.0, e)),
        Some(3.0),
        "E starts at grid column 0, so its leading edge is the table's"
    );
}

/// The trap-detector: a gapped row's own **interior** boundary still paints.
///
/// Row 5 skips a column and then has two cells, so `G|H` has a cell on either
/// side and is interior by every reading. Every other gapped row in this fixture
/// holds a single cell, so without this a change that simply dropped all
/// vertical borders in any row carrying `gridBefore` would pass the whole file.
#[test]
fn an_interior_boundary_inside_a_gapped_row_is_still_painted() {
    let pages = layout();
    let (g, _) = cell(&pages, G);
    let (h, _) = cell(&pages, H);

    assert_eq!(g.0 + g.2, h.0, "G and H are adjacent — one grid line");
    let width = |band: Option<Band>| band.map(|(x0, x1)| x1 - x0);
    assert_eq!(
        width(band_at(&pages, BLUE_INSIDE_V, g.0 + g.2, g)),
        Some(1.0),
        "a cell faces this edge, so it is interior however the row starts"
    );
    // …while the same row's gap-facing edge takes the *outer* border, which is
    // what makes this a discrimination rather than a restatement: one edge of
    // one cell is interior and the other is the table's.
    assert_eq!(
        band_at(&pages, RED_LEFT, g.0, g).map(|(x0, x1)| x1 - x0),
        Some(3.0),
        "G's leading edge"
    );
}

// ── one grid line, one band ─────────────────────────────────────────────────

/// The y-extent of a horizontal border lying on `y` at `x`, as `(y0, y1)`.
fn horizontal_band_on(pages: &[LayoutedPage], y: f32, x: f32) -> Option<Band> {
    [GREY_HORIZONTAL, RED_LEFT, GREEN_RIGHT, BLUE_INSIDE_V]
        .into_iter()
        .find_map(|colour| {
            rects(pages, colour)
                .into_iter()
                .find(|&(rx, ry, rw, rh)| {
                    rw > rh && rx <= x && x <= rx + rw && ry - 0.6 <= y && y <= ry + rh + 0.6
                })
                .map(|(_, ry, _, rh)| (ry, ry + rh))
        })
}

/// §17.4.66: a collapsed border **straddles the line two cells share, and goes
/// inside a line shared with nothing**.
///
/// A cell edge is a line, not a strip. Word draws a collapsed border straddling
/// it — half the declared `w:sz` on each side — so a 1pt `insideV` on the grid
/// line at x = 122 occupies 121.5..122.5 and a 3pt `w:left` on the same line
/// occupies 120.5..123.5. This engine used to paint each border *inside* the
/// cell that owned it, which put those two on opposite sides of the line
/// instead: 121..122 against 122..125.
///
/// The table's own two vertical edges are the exception, and are asserted here
/// rather than skipped: nothing shares them, so there is nothing to straddle and
/// the border goes inside — which is also what keeps a table's ink within the
/// box it reports (`tests/table_auto_width.rs`).
///
/// Asserted as an audit over every vertical edge of every cell, because the
/// claim is about the model rather than about one border.
#[test]
fn every_vertical_border_straddles_a_shared_line_and_goes_inside_an_outer_one() {
    let pages = layout();
    // The tables' own left and right, taken from the cells rather than written
    // down: whatever the fixture's grid is, these are its two ends.
    let (mut left, mut right) = (f32::INFINITY, f32::NEG_INFINITY);
    for fill in [A, B, C, D, E, F, G, H] {
        let (upper, lower) = cell(&pages, fill);
        for b in [upper, lower] {
            left = left.min(b.0);
            right = right.max(b.0 + b.2);
        }
    }

    let mut straddled = 0;
    let mut inset = 0;
    for fill in [A, B, C, D, E, F, G, H] {
        let (upper, lower) = cell(&pages, fill);
        for cell_box in [upper, lower] {
            for x in [cell_box.0, cell_box.0 + cell_box.2] {
                let Some((_, (x0, x1))) = anything_at(&pages, x, cell_box) else {
                    continue;
                };
                let (want, what) = if (x - left).abs() < 0.05 {
                    (x0, "start at the table's own left")
                } else if (x - right).abs() < 0.05 {
                    (x1, "end at the table's own right")
                } else {
                    ((x0 + x1) * 0.5, "be centred on its shared grid line")
                };
                if what.starts_with("be centred") {
                    straddled += 1;
                } else {
                    inset += 1;
                }
                assert!(
                    (want - x).abs() < 0.05,
                    "{fill:?}: the border on the grid line at x={x} occupies \
                     {x0}..{x1}, which does not {what}"
                );
            }
        }
    }
    // Both halves of the rule must actually have been exercised, or the audit
    // proves only the half the fixture happened to contain.
    assert!(straddled >= 4, "only {straddled} shared lines seen");
    assert!(inset >= 2, "only {inset} outer lines seen");
}

/// The reported pair, stated directly: two borders of **different weights** on
/// one grid line come out concentric.
///
/// Row 1 puts a 1pt `insideV` on the line between `A` and `B`; row 2 puts the
/// 3pt `w:left` of its gapped row on the same line. Under the old
/// inside-the-cell model they sat on opposite sides of it and read as a step.
#[test]
fn borders_of_different_weights_on_one_grid_line_are_concentric() {
    let pages = layout();
    let (a, _) = cell(&pages, A);
    let (d, _) = cell(&pages, D);

    let grid_line = a.0 + a.2;
    assert_eq!(grid_line, d.0, "A ends where D begins — one grid line");

    let thin = band_at(&pages, BLUE_INSIDE_V, grid_line, a).expect("row 1's insideV");
    let thick = band_at(&pages, RED_LEFT, grid_line, d).expect("row 2's w:left");

    assert!(thick.1 - thick.0 > thin.1 - thin.0, "3pt against 1pt");
    assert!(
        ((thin.0 + thin.1) * 0.5 - (thick.0 + thick.1) * 0.5).abs() < 0.05,
        "concentric: {thin:?} and {thick:?} must share a centre"
    );
}

/// The same rule on a horizontal edge, both halves of it: the boundary `A` and
/// `D` share is straddled, and the table's own top — shared with nothing — goes
/// inside the first row.
#[test]
fn a_shared_row_boundary_is_straddled_and_the_tables_top_goes_inside() {
    let pages = layout();
    let (a, _) = cell(&pages, A);
    let (b, _) = cell(&pages, B);
    let (d, _) = cell(&pages, D);

    // The interior boundary is the middle of the §17.4.38 strip the two rows
    // reserve between their content boxes — the strip exists to hold that rule,
    // so its centre *is* the boundary. Row 2's `gridBefore` puts `D` under `B`
    // and not under `A`, which is the whole subject of this fixture, so the test
    // had better not assume otherwise.
    let (foot, head) = (b.1 + b.3, d.1);
    assert!(head >= foot - 0.05, "D's box begins at or below B's foot");
    let shared = (foot + head) * 0.5;
    let (y0, y1) = horizontal_band_on(&pages, shared, d.0 + d.2 * 0.5).expect("the shared rule");
    assert!(
        ((y0 + y1) * 0.5 - shared).abs() < 0.05,
        "the shared rule occupies {y0}..{y1}, whose centre is {} rather than {shared}",
        (y0 + y1) * 0.5,
    );

    // The table's own top: no row above it, so the whole border is below the
    // line rather than half of it.
    let (y0, y1) = horizontal_band_on(&pages, a.1, a.0 + a.2 * 0.5).expect("the table's top rule");
    assert!(
        (y0 - a.1).abs() < 0.05,
        "the top rule occupies {y0}..{y1}, which does not begin at the table's \
         own top edge {}",
        a.1
    );
}

// ── a row boundary is one line across ───────────────────────────────────────

/// Horizontal border runs lying on the line `y`, of any colour, clipped to
/// `x0..x1` and sorted.
///
/// "Lying on `y`" is *vertical extent contains `y`*, not *origin equals `y`*,
/// because a row boundary's line can be drawn either in the strip reserved
/// between the two rows (an upper cell's `bottom`) or inset into a cell's own
/// box (a `top`, where no strip was reserved). Both touch the boundary; only one
/// has its origin there.
fn horizontal_runs_on(pages: &[LayoutedPage], y: f32, x0: f32, x1: f32) -> Vec<Band> {
    let mut out: Vec<Band> = Vec::new();
    for colour in [GREY_HORIZONTAL, RED_LEFT, GREEN_RIGHT, BLUE_INSIDE_V] {
        for (rx, ry, rw, rh) in rects(pages, colour) {
            // Any rect crossing `y`, whatever its shape. Filtering to
            // wider-than-tall ones would ask a different question than the one
            // these audits mean: a junction square is as much a part of the line
            // through it as a segment is, and once
            // `coalesce_abutting_rects` fuses a junction with the *vertical*
            // below it — which is what keeps that pair from seaming — the ink at
            // the boundary's end arrives inside a tall rect. Reported as a hole,
            // it is the opposite of what this is looking for.
            //
            // The looseness costs little: a vertical that merely passes through
            // covers one border's width of the boundary, and a real hole in a
            // horizontal is a whole grid column wide.
            if ry - 0.05 <= y && y <= ry + rh + 0.05 {
                let (a, b) = (rx.max(x0), (rx + rw).min(x1));
                if b > a {
                    out.push((a, b));
                }
            }
        }
    }
    out.sort_by(|p, q| p.0.total_cmp(&q.0));
    out
}

/// Whether `runs` leave no hole in `x0..x1`.
fn leaves_no_hole(runs: &[Band], x0: f32, x1: f32) -> bool {
    let mut reached = x0;
    for &(a, b) in runs {
        if a > reached + 0.05 {
            return false;
        }
        reached = reached.max(b);
    }
    reached >= x1 - 0.05
}

/// §17.4.39: a cell wider than the row above it still gets its whole top border.
///
/// Cell `E` spans grid columns 0–1; the row above spans 1–2, so column 0 has no
/// cell above it at all. The boundary resolves to `insideH` over every column —
/// column 0 from `E`'s own top, with nothing facing it — so the line must run
/// across `E` from end to end.
///
/// The reported symptom is the left 50pt of it missing: §17.4.66's edge is owned
/// by one side, a cell paints one border across its whole width, and neither row
/// covers every bordered column here, so the part only `E` covers had no owner.
#[test]
fn a_cell_wider_than_the_row_above_still_gets_its_whole_top_border() {
    let pages = layout();
    let (e, _) = cell(&pages, E);
    let (x0, x1) = (e.0, e.0 + e.2);
    let runs = horizontal_runs_on(&pages, e.1, x0, x1);

    assert!(
        leaves_no_hole(&runs, x0, x1),
        "E spans {x0}..{x1} but its top edge is painted only over {runs:?}"
    );
}

/// The general property, over every cell of both tables: a cell's top and bottom
/// edges are each painted across the cell's whole width, or not at all — never
/// partly.
///
/// Stated as an audit rather than at the one boundary that was reported, because
/// the rule is about ownership of a shared edge and the next hole will be
/// somewhere else. It is satisfied vacuously by an edge that carries no border,
/// so the count assertion at the end keeps it honest.
#[test]
fn no_cell_edge_is_painted_across_only_part_of_its_width() {
    let pages = layout();
    let mut painted_edges = 0;
    for fill in [A, B, C, D, E, F, G, H] {
        let (upper, lower) = cell(&pages, fill);
        for cell_box in [upper, lower] {
            let (x0, x1) = (cell_box.0, cell_box.0 + cell_box.2);
            for (edge, y) in [("top", cell_box.1), ("bottom", cell_box.1 + cell_box.3)] {
                let runs = horizontal_runs_on(&pages, y, x0, x1);
                if runs.is_empty() {
                    continue;
                }
                painted_edges += 1;
                assert!(
                    leaves_no_hole(&runs, x0, x1),
                    "{fill:?}'s {edge} edge spans {x0}..{x1} but is painted only over {runs:?}"
                );
                // The other half of the rule — that no square is painted
                // *twice* — used to be asserted here and is not any more. It
                // cannot be, from these runs: `horizontal_runs_on` counts every
                // rect crossing the boundary, verticals included, and two
                // verticals of different widths on one grid line genuinely
                // overlap in the x-window a cell clips them to without anything
                // being painted twice. The property is real and still audited,
                // over whole pages and with the geometry to do it correctly, by
                // `tests/table_border_corners.rs`.
            }
        }
    }
    assert!(
        painted_edges >= 16,
        "the audit must actually see painted edges — only {painted_edges}"
    );
}
