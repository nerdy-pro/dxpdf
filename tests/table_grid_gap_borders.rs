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
//! # Which one this file asserts, and why
//!
//! A Word render of `test-files/bidi-visual-table.docx` **rules out A**: Word
//! paints nothing on that edge. It cannot separate B from C, because that
//! table's `w:left` is `nil` and both then predict nothing.
//!
//! The engine implements **C**, and the argument is §17.4.35/§17.4.37's own
//! wording: `w:left` and `w:right` are the borders displayed *around the
//! specified table*. In row 2 the table's left side is at x = 72 and the row's
//! first cell begins at x = 122, so reading B paints an outer table border 50pt
//! inside the table — a red 3pt line with the table's own edge visible to the
//! left of it. §17.4.66's conflict vocabulary is "cell borders and outer table
//! borders", and at this edge there is neither a facing cell border nor an outer
//! boundary, so nothing is seeded. C also removes a line rather than inventing
//! one, which is the conservative direction when the spec is silent.
//!
//! **What would settle it.** `test-files/grid-gap-borders.docx` is built so one
//! Word render decides: at cell `D`'s leading edge, a 3pt red line means B and
//! this file is wrong; nothing means C. The fixture's first table exists
//! precisely because the originally reported one could not tell them apart.
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

/// §17.4.15: the edge facing a `gridBefore` gap takes no table-level border.
///
/// Reading C. `insideV` is ruled out by Word (see the module doc); `w:left`
/// would put an outer table border 50pt inside the table, with the table's real
/// left edge still visible to the left of it.
#[test]
fn a_grid_before_gap_edge_takes_no_table_level_border() {
    let pages = layout();
    let (d, _) = cell(&pages, D);

    assert_eq!(
        anything_at(&pages, d.0, d),
        None,
        "nothing faces D's leading edge, so no table-level border reaches it"
    );
}

/// §17.4.14: the mirror image, at the trailing edge of a `gridAfter` row.
///
/// The two are one rule, and the pair is what stops a fix reaching only the end
/// that was reported.
#[test]
fn a_grid_after_gap_edge_takes_no_table_level_border() {
    let pages = layout();
    let (e, _) = cell(&pages, E);

    assert_eq!(anything_at(&pages, e.0 + e.2, e), None);
}

/// A row gapped at **both** ends takes a border at neither.
#[test]
fn a_row_gapped_at_both_ends_takes_a_border_at_neither() {
    let pages = layout();
    let (f, _) = cell(&pages, F);

    assert_eq!(anything_at(&pages, f.0, f), None, "F's leading edge");
    assert_eq!(anything_at(&pages, f.0 + f.2, f), None, "F's trailing edge");
}

/// The reported case, isolated: the same rows with `w:left`/`w:right` set to
/// `nil`. Word paints nothing at either gap edge, and now so does the engine.
///
/// This is `bidi-visual-table.docx`'s symptom with no `w:bidiVisual` and no
/// Hebrew involved — the one assertion in this file standing directly on a Word
/// measurement rather than on a reading of the spec.
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
    // …while the same row's gap-facing edge takes nothing, which is what makes
    // this a discrimination rather than a restatement.
    assert_eq!(anything_at(&pages, g.0, g), None, "G's leading edge");
}

// ── one grid line, one band ─────────────────────────────────────────────────

/// No grid line is painted in two different bands.
///
/// A grid line is one line, and which cell owns it decides which side of it the
/// band lands on: §17.4.66 hands a collapsed interior edge to the cell on the
/// **left**, and a border is painted inward from its owner's box, so an edge
/// owned by the cell on the *right* comes out one border-width over. That is
/// what a `gridBefore` leading edge used to be, and it showed as row 1 painting
/// x = 122 at 121..122 while row 2 painted it at 122..123.
///
/// Asserted as an audit over every vertical edge of every cell rather than at
/// the one boundary that was reported, because the property is general and the
/// next violation will not be in the same place. It knows nothing about gaps,
/// which is the point — a future reading-B implementation that painted `w:left`
/// inward at a gap-facing edge would reintroduce exactly this and be caught here
/// without anyone remembering to look.
#[test]
fn no_grid_line_is_painted_in_two_different_bands() {
    let pages = layout();
    let fills = [A, B, C, D, E, F, G, H];

    // Both tables at once: a band is keyed by its y as well as its x, so the two
    // tables' copies of a grid line never compare against each other.
    let mut seen: Vec<(i64, i64, i64, i64)> = Vec::new();
    for fill in fills {
        let (upper, lower) = cell(&pages, fill);
        for cell_box in [upper, lower] {
            for x in [cell_box.0, cell_box.0 + cell_box.2] {
                let Some((_, (x0, x1))) = anything_at(&pages, x, cell_box) else {
                    continue;
                };
                let key = (
                    (x * 1000.0).round() as i64,
                    (cell_box.1 * 1000.0).round() as i64,
                    (x0 * 1000.0).round() as i64,
                    (x1 * 1000.0).round() as i64,
                );
                if let Some(prev) = seen
                    .iter()
                    .find(|p| p.0 == key.0 && p.1 != key.1 && (p.2, p.3) != (key.2, key.3))
                {
                    panic!(
                        "the grid line at x={x} is painted in two bands: \
                         {:?} and {:?}",
                        (prev.2 as f32 / 1000.0, prev.3 as f32 / 1000.0),
                        (x0, x1)
                    );
                }
                seen.push(key);
            }
        }
    }
    assert!(
        seen.len() >= 8,
        "the audit must actually see borders — only {} found",
        seen.len()
    );
}
