//! §17.4.45/§17.4.43 `w:tblCellSpacing` geometry — issue #165.
//!
//! `test-files/issue-165-cellspacing-scale.docx` is four otherwise-identical
//! tables — same `w:tblW`, same three columns, same fixed layout — at spacings
//! 0, 200, 400 twips, and 400 with a row-level 800.
//!
//! # Measured, not reasoned
//!
//! Every claim below comes from a **Word render** of that fixture (2026-08-18),
//! and three of them replace a reading this file previously pinned the other
//! way. The readings are scale-free, which is what lets them be taken off a
//! screenshot: with `3w + 4g = 360pt`, table 3 shows 189px of cell against 120px
//! of gap — a ratio of 1.58, forcing `g ≈ 41pt`. Had the declared 20pt been the
//! gap, the cells would be 93.3pt and the ratio 4.67. That is a different
//! picture, not a different rounding.
//!
//! * **The declared value is a half-gap** — every gap is twice it.
//! * **The gap is the same at the table's own edge as between two cells.** This
//!   is the decisive reading and the one no earlier fixture could take: table 3's
//!   outline-to-first-cell gap is 120px and its first-to-second-cell gap 121px.
//! * **A row-level value supersedes the table-level one** (§17.4.43), and
//!   governs the table's edge inset too. Table 4 declares 400 on the table and
//!   800 on the row: supersede predicts 80pt gaps and 13.3pt cells, table-wins
//!   66.7pt cells, and summing the two 120pt gaps, which 360pt of table cannot
//!   hold. Word draws 13.5pt cells — narrow enough that the labels wrap to one
//!   glyph per line, which is why this file no longer finds a table by its cell
//!   text.
//! * **The spacing is carved out of `w:tblW`**, unchanged from before: all four
//!   tables span 360pt and the columns shrink to pay.
//!
//! # What that refuted
//!
//! This file used to argue from ONLYOFFICE — an independent implementation that
//! both renders and targets Word compatibility — that there is **no** factor:
//! `sdkjs`'s `TableRecalculate.js` insets a cell by `CellSpacing` on the table's
//! outer edges and `CellSpacing / 2` on every interior side, making every gap
//! the declared value. Word's doubled gaps were put down to the older probe,
//! `issue-165-cellspacing.docx`, declaring the spacing in `tblPr` *and* `trPr`,
//! so a Word that summed them would land on twice.
//!
//! Tables 2 and 3 here declare it at table level **only**, and Word doubles them
//! anyway. There is nothing to sum. A cross-implementation argument lost to a
//! render of the implementation being matched, which is the order those rank in.
//!
//! Geometry is read as **boxes, not ink**: a cell's box runs from the outer face
//! of its left border to the outer face of its right border. That is the level
//! the spacing acts at and the level the render resolves — the fixture's 1pt
//! borders are 2.9px on it, so where inside its box a line is drawn is not a
//! question that render answers, and nothing here asks it.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

const FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/test-files/issue-165-cellspacing-scale.docx"
);

/// One table's measurements, in page coordinates.
struct Table {
    /// `(left, right)` of every vertical border rect, in x order. A spaced
    /// table's read `outline | c1 c1 | c2 c2 | … | outline`.
    verticals: Vec<(f32, f32)>,
    /// x of the leftmost text drawn inside the table — the first cell's left
    /// edge plus a cell margin identical in all four tables.
    ///
    /// Taken as a minimum over every line, because a cell narrow enough forces
    /// its label to wrap and each line arrives as its own draw command. Table 4
    /// is exactly that case.
    first_cell_text_x: f32,
}

/// The four tables, in document order.
///
/// Found by grouping the vertical border rects into bands of overlapping y: a
/// table's outline spans its own cells' bands, nothing spans the paragraph
/// between two tables, and nothing else on these pages paints a rect. Which is
/// what makes this robust where identifying a table by its cell text is not —
/// see [`Table::first_cell_text_x`].
fn tables() -> Vec<Table> {
    let bytes = std::fs::read(FIXTURE).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("fixture parses");
    let pages: Vec<LayoutedPage> = dxpdf::render::resolve_and_layout(doc).1;

    let mut out: Vec<Table> = Vec::new();
    for page in &pages {
        // (x, y, w, h) of every vertical rect on this page, in y order.
        let mut v: Vec<(f32, f32, f32, f32)> = page
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, .. } => Some((
                    rect.origin.x.raw(),
                    rect.origin.y.raw(),
                    rect.size.width.raw(),
                    rect.size.height.raw(),
                )),
                _ => None,
            })
            .filter(|r| r.2 < r.3)
            .collect();
        v.sort_by(|a, b| a.1.partial_cmp(&b.1).unwrap());

        let mut bands: Vec<Vec<(f32, f32, f32, f32)>> = Vec::new();
        let mut bottom = f32::NEG_INFINITY;
        for r in v {
            if r.1 <= bottom + 0.01 {
                bands.last_mut().unwrap().push(r);
                bottom = bottom.max(r.1 + r.3);
            } else {
                bands.push(vec![r]);
                bottom = r.1 + r.3;
            }
        }

        for band in bands {
            let (top, bot) = band.iter().fold((f32::MAX, f32::MIN), |(t, b), r| {
                (t.min(r.1), b.max(r.1 + r.3))
            });
            let first_cell_text_x = page
                .commands
                .iter()
                .filter_map(|c| match c {
                    DrawCommand::Text { position, .. }
                        if top <= position.y.raw() && position.y.raw() <= bot =>
                    {
                        Some(position.x.raw())
                    }
                    _ => None,
                })
                .fold(f32::MAX, f32::min);
            let mut verticals: Vec<(f32, f32)> = band.iter().map(|r| (r.0, r.0 + r.2)).collect();
            verticals.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap());
            out.push(Table {
                verticals,
                first_cell_text_x,
            });
        }
    }
    assert_eq!(out.len(), 4, "four tables");
    out
}

impl Table {
    /// `(left, right)` of the table's own box: the outermost pair of verticals.
    fn table_box(&self) -> (f32, f32) {
        (
            self.verticals[0].0,
            self.verticals[self.verticals.len() - 1].1,
        )
    }

    /// `(left, right)` of each cell's box — the pairs inside the outline.
    ///
    /// Meaningful only for a spaced table; a collapsed one shares its verticals
    /// between the cells either side of them, and has no separate outline.
    fn cell_boxes(&self) -> Vec<(f32, f32)> {
        let n = self.verticals.len();
        assert!(n >= 4 && n.is_multiple_of(2), "{n} verticals");
        self.verticals[1..n - 1]
            .chunks(2)
            .map(|p| (p[0].0, p[1].1))
            .collect()
    }

    /// Every gap, leading edge first: table-to-first-cell, each cell-to-cell,
    /// last-cell-to-table.
    fn gaps(&self) -> Vec<f32> {
        let (left, right) = self.table_box();
        let cells = self.cell_boxes();
        let mut g = vec![cells[0].0 - left];
        g.extend(cells.windows(2).map(|w| w[1].0 - w[0].1));
        g.push(right - cells[cells.len() - 1].1);
        g
    }
}

fn close(a: f32, b: f32) -> bool {
    (a - b).abs() < 0.02
}

/// §17.4.45: the gap Word draws is **twice** the declared value.
///
/// Tables 2 and 3 declare the spacing at table level and nowhere else, which is
/// what makes them evidence about a factor rather than about summing two
/// declarations.
#[test]
fn the_rendered_gap_is_twice_the_declared_spacing() {
    let t = tables();
    for (idx, declared_pt, label) in [(1usize, 10.0_f32, "T2"), (2, 20.0, "T3")] {
        for (i, g) in t[idx].gaps().into_iter().enumerate() {
            assert!(
                close(g, 2.0 * declared_pt),
                "{label} gap {i}: {g}pt against a declared {declared_pt}pt"
            );
        }
    }
}

/// And it is the same gap at the table's own edge as between two cells — the
/// reading that separates a factor from a spacing applied only on the inside.
///
/// A relation between measurements of one table, so it holds whatever the
/// factor turns out to be and fails only if the two places disagree.
#[test]
fn the_gap_at_the_table_edge_equals_the_gap_between_cells() {
    let t = tables();
    for (idx, label) in [(1usize, "T2"), (2, "T3"), (3, "T4")] {
        let gaps = t[idx].gaps();
        for (i, g) in gaps.iter().enumerate() {
            assert!(
                close(*g, gaps[0]),
                "{label}: gap {i} is {g}pt where the leading edge is {}pt",
                gaps[0]
            );
        }
    }
}

/// And it scales linearly, which is what tells a factor from a constant offset.
/// §17.4.45 folds "the width of the table borders" into the spacing, so an
/// implementation that subtracted a border width would still look roughly
/// doubled at one value and wrong at another.
#[test]
fn doubling_the_declared_spacing_doubles_the_gap() {
    let t = tables();
    let (at_200, at_400) = (t[1].gaps()[0], t[2].gaps()[0]);
    assert!(
        close(at_400, at_200 * 2.0),
        "400 twips must give exactly twice the gap of 200: {at_200}pt vs {at_400}pt"
    );
}

/// The spacing is carved out of `w:tblW`, not added to it: all four tables
/// declare the same width and render the same width, with the cells shrinking as
/// the gaps grow. Confirmed against Word, which draws all four to one right
/// edge.
///
/// Table 1 is measured centre-to-centre and the others outer-face to outer-face,
/// because the collapsed path straddles its grid lines while `emit_table_outline`
/// still draws inside the box — a half-border inconsistency between the two
/// constructors that is real, open, and below what the render can resolve. Every
/// table is checked against the declared 360pt rather than against table 1, so
/// this test states the carve without taking a position on that.
#[test]
fn spacing_is_carved_out_of_the_declared_table_width() {
    let t = tables();
    let border = t[0].verticals[0].1 - t[0].verticals[0].0;

    let (left, right) = t[0].table_box();
    assert!(
        close(right - left - border, 360.0),
        "T1 spans {}pt centre-to-centre, not the declared tblW of 360",
        right - left - border
    );
    for (idx, label) in [(1usize, "T2"), (2, "T3"), (3, "T4")] {
        let (left, right) = t[idx].table_box();
        assert!(
            close(right - left, 360.0),
            "{label} spans {}pt, not the declared tblW of 360",
            right - left
        );
    }
}

/// §17.4.43: a row-level `w:tblCellSpacing` supersedes the table-level value.
///
/// The precedence was never in doubt in the text — §17.4.45's own value is
/// "superseded by a table-level exception (§17.4.44) or the row cell spacing
/// value (§17.4.43) in that order". What was in doubt was whether Word obeys it,
/// because the only render then on record declared the same value at both levels
/// and so could not tell supersede from sum. This one can, and Word supersedes.
#[test]
fn a_row_level_spacing_supersedes_the_table_level_one() {
    let t = tables();
    for (i, g) in t[3].gaps().into_iter().enumerate() {
        assert!(close(g, 80.0), "gap {i}: {g}pt against the row's 800 twips");
    }
    for (i, c) in t[3].cell_boxes().iter().enumerate() {
        assert!(
            close(c.1 - c.0, (360.0 - 4.0 * 80.0) / 3.0),
            "cell {i} is {}pt wide, not the 13.3 Word draws",
            c.1 - c.0
        );
    }
}

/// And the cells' *content* moves with the spacing, not only their borders.
///
/// Read as a difference between two **spaced** tables, which cancels the cell
/// margin and the border width — whatever they are, the first cell's text sits
/// the same distance inside its cell in every table here, so the shift between
/// two of them is the difference in their gaps.
///
/// The zero-spacing table would be the natural reference and is deliberately not
/// used: it is laid out by the collapsed constructor, which straddles its grid
/// lines where the spaced one draws inside the box, so measuring across the two
/// carries that half-border difference into a claim that is not about it. See
/// `spacing_is_carved_out_of_the_declared_table_width`.
#[test]
fn the_first_cells_text_moves_with_the_spacing() {
    let t = tables();
    // Declared 200 → 400 → 800 twips, so each step adds 10pt then 20pt to the
    // declaration and must add twice that to where the text starts.
    for (from, to, step_pt, label) in [(1usize, 2usize, 10.0_f32, "T2→T3"), (2, 3, 20.0, "T3→T4")]
    {
        let shift = t[to].first_cell_text_x - t[from].first_cell_text_x;
        assert!(
            close(shift, 2.0 * step_pt),
            "{label}: the text should start {}pt further in, got {shift}pt",
            2.0 * step_pt
        );
    }
}
