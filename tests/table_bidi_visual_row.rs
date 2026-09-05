//! §17.4.60 `<w:tblPrEx><w:bidiVisual/></w:tblPrEx>` — a single row flipped
//! against an otherwise left-to-right table.
//!
//! `test-files/issue-157-tblprex-bidi.docx` is the probe: three rows of cells
//! `A` (50pt), `B` (100pt), `C` (150pt) over an unequal 1000/2000/3000-twip
//! grid, with only the middle row (row 2) carrying the exception. The unequal
//! grid is what makes the width question answerable at all — with three equal
//! columns, "keeps its own width" and "takes the slot's" predict the same
//! page.
//!
//! Measured against a Word render (2026-09-05, pixel-counted off a fresh
//! render and calibrated against the page's own margins rather than
//! eyeballed): cell `A` comes out ~50pt, not ~150pt, so a flipped row keeps
//! each cell's own declared width rather than resizing to the grid slot it
//! visually lands in. The same render also showed two things the fixture's own
//! heading never asked about, which `RowBidiOverride`'s doc names as measured
//! from this one arrangement and not verified beyond it: the flipped row sits
//! with its own right edge on the *page's* content width, as if it were a mini
//! right-to-left table of its own; and — the surprise — **it paints third, not
//! second**. Word renders the page as row 1, then row 3 (both unflipped and
//! otherwise identical), then row 2 (flipped) last, rather than row 2 between
//! its neighbours as document order would suggest. Every test below reads
//! bands top to bottom accordingly: band 0 is row 1, band 1 is row 3, band 2
//! is row 2.

use std::path::Path;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

const FIXTURE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/test-files/issue-157-tblprex-bidi.docx"
);

fn layout() -> Vec<LayoutedPage> {
    let bytes = std::fs::read(Path::new(FIXTURE)).expect("fixture is committed");
    let doc = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

type Rect = (f32, f32, f32, f32);
type Colour = (u8, u8, u8);
type Shaded = (Colour, Rect);

const A: Colour = (0xF8, 0xCB, 0xAD);
const B: Colour = (0xC6, 0xE0, 0xB4);
const C: Colour = (0xBD, 0xD7, 0xEE);

/// Every shaded cell box on the page, grouped into rows by y-position and, in
/// document order within each row, by fill colour. All three rows share the
/// same three fills — `A`/`B`/`C` repeat down the table — so a colour alone
/// cannot name a row the way it can in `tests/table_bidi_visual.rs`'s
/// one-flip-per-document fixtures; distinct *y-bands*, top to bottom, are what
/// separates them here, and a colour is unique only once a band has isolated
/// one row.
fn rows_by_band(pages: &[LayoutedPage]) -> Vec<std::collections::HashMap<Colour, Rect>> {
    let mut shaded: Vec<Shaded> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color } => Some((
                (color.r, color.g, color.b),
                (
                    rect.origin.x.raw(),
                    rect.origin.y.raw(),
                    rect.size.width.raw(),
                    rect.size.height.raw(),
                ),
            )),
            _ => None,
        })
        .filter(|(colour, _)| [A, B, C].contains(colour))
        .collect();
    shaded.sort_by(|a, b| a.1 .1.total_cmp(&b.1 .1));

    let mut rows: Vec<Vec<Shaded>> = Vec::new();
    for (colour, rect) in shaded {
        match rows.last_mut() {
            // A new row starts once y has moved by more than a hairline from
            // the row this rect's predecessor belongs to.
            Some(last) if (rect.1 - last[0].1 .1).abs() < 1.0 => last.push((colour, rect)),
            _ => rows.push(vec![(colour, rect)]),
        }
    }
    rows.into_iter().map(|r| r.into_iter().collect()).collect()
}

/// §17.4.60: cell `A` in the flipped row (band 2 — see the module doc for why
/// it is the third band, not the second) keeps the 50pt it declares, rather
/// than resizing to the 150pt slot it visually lands in. The fixture's own
/// measurement question, answered directly.
#[test]
fn a_flipped_rows_cell_keeps_its_own_declared_width_not_the_slots() {
    let bands = rows_by_band(&layout());
    assert_eq!(bands.len(), 3, "three rows, three distinct y-bands");

    // Rows 1 and 3 (unflipped, bands 0 and 1) are the control: cell A is 50pt
    // wherever it appears without the exception.
    for (label, band) in [("row 1", &bands[0]), ("row 3", &bands[1])] {
        let (_, _, w, _) = band[&A];
        assert!(
            (w - 50.0).abs() < 0.5,
            "{label} (unflipped): cell A should be 50pt, got {w}pt"
        );
    }

    // Row 2 (flipped, band 2) is the question. 50pt confirms it keeps its own
    // declared width; 150pt would mean it took the slot's instead.
    let (_, _, flipped_w, _) = bands[2][&A];
    assert!(
        (flipped_w - 50.0).abs() < 0.5,
        "row 2 (flipped): cell A should still be 50pt (its own declared \
         width), not resized to the 150pt slot it visually lands in — got \
         {flipped_w}pt"
    );
}

/// The flipped row's three cells run right to left — `C` (150pt) leftmost,
/// `B` (100pt) in the middle, `A` (50pt) rightmost — each still its own
/// declared width, so the row's boundaries do not land on the same grid lines
/// as the unflipped rows above and below it.
#[test]
fn a_flipped_rows_cells_run_right_to_left_at_their_own_widths() {
    let bands = rows_by_band(&layout());
    let flipped = &bands[2];

    let (cx, _, cw, _) = flipped[&C];
    let (bx, _, bw, _) = flipped[&B];
    let (ax, _, aw, _) = flipped[&A];

    assert!(cx < bx && bx < ax, "left to right the row reads C, B, A");
    for (label, w, want) in [("C", cw, 150.0), ("B", bw, 100.0), ("A", aw, 50.0)] {
        assert!(
            (w - want).abs() < 0.5,
            "cell {label} should keep its own {want}pt, got {w}pt"
        );
    }
    // Adjacent, not merely ordered — no gap and no overlap between them.
    assert!((cx + cw - bx).abs() < 0.5, "C and B meet exactly");
    assert!((bx + bw - ax).abs() < 0.5, "B and A meet exactly");
}

/// §17.4.60 measured a second question the fixture's own heading never asked:
/// where the flipped row sits. It is not left-aligned like its neighbours —
/// its own right edge reaches the *page's* content width, the same basis a
/// whole `bidiVisual` table is placed against
/// (`tests/table_leading_margin.rs`), as if this one row were a mini
/// right-to-left table of its own.
///
/// The page is the fixture's own (Letter, 1in margins), so content nominally
/// runs 72..540pt — but the control's own right edge lands at 371.5, half a
/// point short of the nominal 72 + 300 = 372, because a table's own outer
/// border straddles the grid line it stands on and shifts the whole table half
/// a leading border left (`border-outer-box.docx`, `tests/table_outer_box.rs`).
/// The flipped row picks up the same half-point, landing at 539.5 rather than
/// a clean 540 — asserted against the control's own shift rather than a bare
/// literal, so the claim is "reaches the content edge this table actually
/// has", not a number this border geometry happens to produce today.
#[test]
fn a_flipped_row_reaches_the_pages_content_width_not_the_tables_own() {
    let bands = rows_by_band(&layout());
    let control_right = {
        let (x, _, w, _) = bands[0][&C];
        x + w
    };
    let flipped_right = {
        let (x, _, w, _) = bands[2][&A];
        x + w
    };
    assert!(
        (control_right - 371.5).abs() < 0.1,
        "control table's own right edge should be at 72 + 300, shifted half a \
         leading border left, got {control_right}"
    );
    let page_right_margin = 612.0 - 72.0; // Letter, 1in margins: content ends here.
    let expected_flipped_right = page_right_margin - 0.5; // same half-point shift.
    assert!(
        (flipped_right - expected_flipped_right).abs() < 0.1,
        "flipped row's right edge should reach the page's content width \
         (expected {expected_flipped_right}), not the table's own \
         {control_right} — got {flipped_right}"
    );
}

/// §17.4.60 measured a third question: paint order. The flipped row (row 2 in
/// document order) does not paint between its neighbours — it paints *after*
/// row 3, which is the control's own unflipped twin and therefore
/// indistinguishable from row 1 except by position.
///
/// Told apart from row 1 by y rather than by content — both are identical,
/// unflipped `A`/`B`/`C` rows — so band 1 is asserted to equal band 0's shape
/// exactly (same colours at the same x, only a different y), and band 2 is
/// asserted to be the flipped one instead.
#[test]
fn a_flipped_row_paints_after_the_row_that_would_otherwise_follow_it() {
    let bands = rows_by_band(&layout());

    // Band 1 (row 3) is an unflipped control row: same three x positions as
    // band 0 (row 1), in the same left-to-right order.
    for colour in [A, B, C] {
        let (x0, _, w0, _) = bands[0][&colour];
        let (x1, _, w1, _) = bands[1][&colour];
        assert!(
            (x0 - x1).abs() < 0.5 && (w0 - w1).abs() < 0.5,
            "{colour:?}: band 1 (row 3) should be an unflipped row in the \
             same columns as band 0 (row 1) at ({x0}, {w0}pt), got \
             ({x1}, {w1}pt)"
        );
    }

    // Band 2 (row 2) is the flipped row: reversed order, shifted right —
    // already pinned by the two tests above, restated here as "band 2, not
    // band 1, is the odd one out" to name what actually moved.
    let (ax, _, aw, _) = bands[2][&A];
    let (ax0, _, aw0, _) = bands[0][&A];
    assert!(
        (ax - ax0).abs() > 50.0 || (aw - aw0).abs() > 0.5,
        "band 2 should be the flipped row, visibly different from the \
         unflipped control"
    );
}
