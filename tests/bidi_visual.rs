//! §17.4.1 `w:bidiVisual` end to end (issue #157): a table carrying the flag
//! displays its columns right to left — the row's first cell is the rightmost
//! one — and `gridSpan`, `vMerge` and the column widths all mirror with it.
//!
//! The fixture (`scripts/make_issue157_fixture.py`) puts three tables over one
//! lopsided 72/144/216 pt grid, each cell holding a unique two-letter token in
//! a single run, so every token is one text draw command and the assertions
//! are pure geometry — no glyph metric and no face name is pinned. The
//! lopsided grid is what separates a genuine mirror from a mere reversal of
//! the text: the gap between mirrored neighbours must be the *mirrored*
//! column's width, which equal columns could not detect.

use std::path::Path;

use dxpdf::render::layout::draw_command::DrawCommand;

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

/// Every text draw command of the laid-out fixture as `(token, x)`, trimmed.
fn drawn_tokens() -> Vec<(String, f32)> {
    let path = Path::new(TEST_DIR).join("bidi-visual.docx");
    let bytes =
        std::fs::read(&path).unwrap_or_else(|e| panic!("failed to read {}: {e}", path.display()));
    let doc = dxpdf::docx::parse(&bytes)
        .unwrap_or_else(|e| panic!("failed to parse {}: {e}", path.display()));
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } => {
                Some((text.trim().to_string(), position.x.raw()))
            }
            _ => None,
        })
        .filter(|(t, _)| !t.is_empty())
        .collect()
}

/// The x of the one draw command holding `token`.
fn x_of(tokens: &[(String, f32)], token: &str) -> f32 {
    let hits: Vec<f32> = tokens
        .iter()
        .filter(|(t, _)| t == token)
        .map(|&(_, x)| x)
        .collect();
    assert_eq!(
        hits.len(),
        1,
        "token {token:?} must be drawn exactly once, found {hits:?} in {tokens:?}"
    );
    hits[0]
}

/// The row's first cell renders rightmost, and each gap between display
/// neighbours is the width of the column *between* them after the mirror —
/// CC's display column is the logical third (216 pt), BB's the logical
/// second (144 pt). Cell margins are uniform across the row, so they cancel
/// out of every difference.
#[test]
fn columns_render_right_to_left_with_mirrored_widths() {
    let tokens = drawn_tokens();
    let (aa, bb, cc) = (
        x_of(&tokens, "AA"),
        x_of(&tokens, "BB"),
        x_of(&tokens, "CC"),
    );
    assert!(
        cc < bb && bb < aa,
        "logical order AA|BB|CC must display as CC|BB|AA: AA={aa}, BB={bb}, CC={cc}"
    );
    assert!(
        (bb - cc - 216.0).abs() < 0.5,
        "CC sits in the mirrored 216 pt column, so BB starts 216 pt after it: {}",
        bb - cc
    );
    assert!(
        (aa - bb - 144.0).abs() < 0.5,
        "BB sits in the mirrored 144 pt column, so AA starts 144 pt after it: {}",
        aa - bb
    );
}

/// A `gridSpan=2` cell over the logical first two columns lands on the
/// display *right*, covering the mirrored last two, and the single cell
/// beside it takes the display-leftmost 216 pt column.
#[test]
fn a_grid_span_mirrors_to_the_right() {
    let tokens = drawn_tokens();
    let (dd, ee) = (x_of(&tokens, "DD"), x_of(&tokens, "EE"));
    assert!(
        ee < dd,
        "the spanning first cell must display right of its neighbour: DD={dd}, EE={ee}"
    );
    // 1 pt of slack, not 0.5: this table is bordered, and a cell against the
    // table's outer border is inset differently from one against an interior
    // `insideV` — a sub-line-width (0.5 pt) difference that is the border
    // model's business, pinned by `tests/table_cell_content_box.rs`. The
    // *exact* mirrored widths are pinned above on the borderless table.
    assert!(
        (dd - ee - 216.0).abs() < 1.0,
        "EE's display column is the mirrored 216 pt one: {}",
        dd - ee
    );
}

/// A `vMerge` pair in the logical first column displays rightmost: the
/// restart's content is drawn once, right of its row-mates, and the continue
/// row's remaining cells mirror like any others.
#[test]
fn a_v_merge_pair_mirrors_with_its_column() {
    let tokens = drawn_tokens();
    let (ff, gg, hh) = (
        x_of(&tokens, "FF"),
        x_of(&tokens, "GG"),
        x_of(&tokens, "HH"),
    );
    assert!(
        ff > gg && ff > hh,
        "the restart cell of logical column 1 must display rightmost: FF={ff}, GG={gg}, HH={hh}"
    );
    let (ii, jj) = (x_of(&tokens, "II"), x_of(&tokens, "JJ"));
    assert!(
        ii > jj,
        "under the continue, logical II|JJ must display as JJ|II: II={ii}, JJ={jj}"
    );
    assert!(
        (ff - ii - 144.0).abs() < 0.5,
        "FF's merge column sits one mirrored 144 pt column right of II — \
         the continue row leaves that gap under it: FF={ff}, II={ii}"
    );
}

/// The table without the flag keeps its columns left to right — `bidiVisual`
/// mirrors the table that carries it and nothing else.
#[test]
fn a_table_without_the_flag_stays_left_to_right() {
    let tokens = drawn_tokens();
    let (kk, ll, mm) = (
        x_of(&tokens, "KK"),
        x_of(&tokens, "LL"),
        x_of(&tokens, "MM"),
    );
    assert!(
        kk < ll && ll < mm,
        "the control table must stay LTR: KK={kk}, LL={ll}, MM={mm}"
    );
    assert!(
        (ll - kk - 72.0).abs() < 0.5 && (mm - ll - 144.0).abs() < 0.5,
        "…at the declared 72/144 pt column widths: {} and {}",
        ll - kk,
        mm - ll
    );
}
