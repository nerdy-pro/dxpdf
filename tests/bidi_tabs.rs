//! §17.3.1.37 tab stops under `w:bidi`, end to end (issue #156): a
//! right-to-left paragraph measures its stops from the *right* margin and
//! walks its tab-delimited segments right to left, and a numbering label
//! before its suffix tab lands against the right margin.
//!
//! The fixture (`scripts/make_issue156_fixture.py`) puts the text column at
//! 72..540 pt and pairs every RTL paragraph with an LTR control. The
//! metric-free assertions use *end* stops in the RTL paragraphs — an
//! end-anchored zone's left edge is exactly `540 − pos`, no glyph width
//! involved — mirroring the *start* stops of the controls at `72 + pos`.

use std::path::Path;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

/// The text column of the fixture's Letter page: 12240 − 2×1440 twips.
const LEFT: f32 = 72.0;
const RIGHT: f32 = 540.0;

fn pages() -> Vec<LayoutedPage> {
    let path = Path::new(TEST_DIR).join("bidi-tabs.docx");
    let bytes =
        std::fs::read(&path).unwrap_or_else(|e| panic!("failed to read {}: {e}", path.display()));
    let doc = dxpdf::docx::parse(&bytes)
        .unwrap_or_else(|e| panic!("failed to parse {}: {e}", path.display()));
    dxpdf::render::resolve_and_layout(doc).1
}

/// The x of the one text draw command holding exactly `token`.
#[track_caller]
fn x_of(pages: &[LayoutedPage], token: &str) -> f32 {
    let hits: Vec<f32> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.trim() == token => {
                Some(position.x.raw())
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        hits.len(),
        1,
        "token {token:?} must be drawn exactly once, found {hits:?}"
    );
    hits[0]
}

/// Custom stops mirror: the RTL paragraph's segments run right to left, its
/// end-stop zones' left edges land at `540 − pos` exactly, and the LTR
/// control keeps the unmirrored `72 + pos` positions.
#[test]
fn custom_stops_measure_from_the_right_margin() {
    let pages = pages();
    let (ra, rb, rc) = (x_of(&pages, "RA"), x_of(&pages, "RB"), x_of(&pages, "RC"));
    assert!(
        rc < rb && rb < ra,
        "logical RA→RB→RC must display right to left: RA={ra}, RB={rb}, RC={rc}"
    );
    assert!(
        (rb - (RIGHT - 100.0)).abs() < 0.5,
        "RB's end-stop at 100 pt from the right margin: {rb}"
    );
    assert!(
        (rc - (RIGHT - 200.0)).abs() < 0.5,
        "RC's end-stop at 200 pt from the right margin: {rc}"
    );

    let (la, lb, lc) = (x_of(&pages, "LA"), x_of(&pages, "LB"), x_of(&pages, "LC"));
    assert!(
        (la - LEFT).abs() < 0.5
            && (lb - (LEFT + 100.0)).abs() < 0.5
            && (lc - (LEFT + 200.0)).abs() < 0.5,
        "the LTR control keeps its start stops at 72/172/272: LA={la}, LB={lb}, LC={lc}"
    );
}

/// The numbering label of an RTL item sits against the right margin — right
/// of its own body text, in the hanging-indent area 18..36 pt in from the
/// right — while the LTR control pins the unmirrored geometry exactly:
/// hanging 18 pt inside a 36 pt indent puts its label at 72 + 18 = 90.
#[test]
fn a_numbering_label_lands_against_the_right_margin() {
    let pages = pages();
    // The RTL label "7." reaches the page as two commands: UAX #9 gives the
    // digit an LTR island level and the trailing period the paragraph level,
    // so the period draws to the *left* of the digit — ".7" read right to
    // left, which is Word's own rendering of an RTL list label. The digit is
    // the unique token to find.
    let (label, body) = (x_of(&pages, "7"), x_of(&pages, "RNUM"));
    assert!(
        label > body,
        "the RTL label must display right of its body text: label={label}, body={body}"
    );
    assert!(
        label > RIGHT - 40.0,
        "…in the hanging area against the right margin: {label}"
    );

    let (l_label, l_body) = (x_of(&pages, "3."), x_of(&pages, "LNUM"));
    assert!(
        (l_label - (LEFT + 18.0)).abs() < 0.5,
        "the LTR control's label starts at indent − hanging = 90: {l_label}"
    );
    assert!(l_label < l_body, "…left of its body text");
}

/// A `bar` stop names a column of the paragraph, so its position mirrors
/// like every other: the rule draws at 540 − 150.
#[test]
fn a_bar_stop_rule_mirrors_its_position() {
    let pages = pages();
    let bars: Vec<f32> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Line { line, .. }
                if line.start.x == line.end.x && line.start.y != line.end.y =>
            {
                Some(line.start.x.raw())
            }
            _ => None,
        })
        .collect();
    assert!(
        bars.iter().any(|&x| (x - (RIGHT - 150.0)).abs() < 0.5),
        "a vertical rule at 540 − 150 = 390: {bars:?}"
    );
    assert!(
        !bars.iter().any(|&x| (x - 150.0).abs() < 0.5),
        "…and none at the unmirrored 150: {bars:?}"
    );
}

/// With no custom stops the §17.15.1.25 default grid is walked from the
/// right: the segment after the tab lands left of the first one.
#[test]
fn the_default_grid_is_walked_from_the_right() {
    let pages = pages();
    let (re, rf) = (x_of(&pages, "RE"), x_of(&pages, "RF"));
    assert!(
        rf < re,
        "RF must land on a grid stop left of RE: RE={re}, RF={rf}"
    );
    assert!(re > RIGHT - 40.0, "RE hugs the right margin: {re}");
}
