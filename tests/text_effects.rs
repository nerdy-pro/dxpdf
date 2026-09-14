//! The four legacy run text effects, end to end (issue #148): `w:shadow`
//! draws a hard offset copy under the glyphs, `w:outline` strokes them
//! hollow, `w:emboss`/`w:imprint` draw a relief copy and lift the glyph
//! colour — where before this all four parsed, cascaded, and changed
//! nothing on the page.
//!
//! Geometry and colours follow the one citable Word-compatible
//! implementation (LibreOffice's VCL emulation, written for WW8/DOCX
//! compatibility) — the constants live with `TextEffects` in
//! `render::layout::fragment` and are pinned by its unit tests; here the
//! *shape* of each effect is asserted: copy counts, offset directions,
//! colour keying, draw order (copy under main), and the stroke flag.

use std::path::Path;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

fn pages() -> Vec<LayoutedPage> {
    let path = Path::new(TEST_DIR).join("text-effects.docx");
    let bytes =
        std::fs::read(&path).unwrap_or_else(|e| panic!("failed to read {}: {e}", path.display()));
    let doc = dxpdf::docx::parse(&bytes)
        .unwrap_or_else(|e| panic!("failed to parse {}: {e}", path.display()));
    dxpdf::render::resolve_and_layout(doc).1
}

/// One drawn command: `(x, y, (r, g, b), outline)`.
type Draw = (f32, f32, (u8, u8, u8), bool);

/// Every draw command holding `token`, in draw order.
fn draws_of(pages: &[LayoutedPage], token: &str) -> Vec<Draw> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text {
                text,
                position,
                color,
                outline,
                ..
            } if text.trim() == token => Some((
                position.x.raw(),
                position.y.raw(),
                (color.r, color.g, color.b),
                *outline,
            )),
            _ => None,
        })
        .collect()
}

const LIGHT_GRAY: (u8, u8, u8) = (192, 192, 192);
const BLACK: (u8, u8, u8) = (0, 0, 0);
const WHITE: (u8, u8, u8) = (255, 255, 255);

/// `[copy, main]` — the effect copy must be drawn first, so the main glyph
/// paints over it where they overlap.
#[track_caller]
fn copy_and_main(pages: &[LayoutedPage], token: &str) -> (Draw, Draw) {
    let draws = draws_of(pages, token);
    assert_eq!(
        draws.len(),
        2,
        "{token}: an effect copy plus the main glyphs; got {draws:?}"
    );
    (draws[0], draws[1])
}

/// The control: no effect, one command, not stroked.
#[test]
fn a_plain_run_draws_once_unstroked() {
    let pages = pages();
    let draws = draws_of(&pages, "PLAIN");
    assert_eq!(draws.len(), 1, "{draws:?}");
    assert!(!draws[0].3, "no outline flag on plain text");
}

/// §17.3.2.31: the shadow is a hard copy beneath the text and to its
/// right — down-right offset, equal in both axes — drawn under the main
/// glyphs. On dark text the copy is light gray.
#[test]
fn shadow_draws_a_light_gray_copy_down_right_of_dark_text() {
    let pages = pages();
    let (copy, main) = copy_and_main(&pages, "SHDW");
    let (dx, dy) = (copy.0 - main.0, copy.1 - main.1);
    assert!(
        dx > 0.0 && dy > 0.0 && (dx - dy).abs() < 1e-4,
        "the copy sits down-right of the main glyphs: dx={dx}, dy={dy}"
    );
    assert_eq!(copy.2, LIGHT_GRAY, "a dark run's shadow is light gray");
    assert_eq!(main.2, BLACK, "…and the text keeps its own colour");
}

/// The shadow colour is keyed on the text's luminance: a red run's shadow
/// is black, not gray.
#[test]
fn shadow_of_a_bright_run_is_black() {
    let pages = pages();
    let (copy, main) = copy_and_main(&pages, "SHDWRED");
    assert_eq!(copy.2, BLACK, "a bright run's shadow is black");
    assert_eq!(main.2, (255, 0, 0), "…under the red text");
}

/// §17.3.2.23: outline strokes the glyphs hollow — one command, flagged
/// for the painter's stroke pass.
#[test]
fn outline_strokes_a_single_command() {
    let pages = pages();
    let draws = draws_of(&pages, "OUTL");
    assert_eq!(draws.len(), 1, "{draws:?}");
    assert!(draws[0].3, "the command carries the outline flag");
}

/// Shadow + outline is the one §17.3.2.31-permitted combination: both
/// commands stroke, and the shadow offset grows by one pixel — so it must
/// exceed the plain shadow's.
#[test]
fn shadow_plus_outline_strokes_both_and_offsets_further() {
    let pages = pages();
    let (copy, main) = copy_and_main(&pages, "SHOUT");
    assert!(copy.3 && main.3, "both the copy and the main glyphs stroke");
    let combined_dx = copy.0 - main.0;

    let (plain_copy, plain_main) = copy_and_main(&pages, "SHDW");
    let plain_dx = plain_copy.0 - plain_main.0;
    assert!(
        combined_dx > plain_dx,
        "the outlined shadow offsets one pixel further: {combined_dx} vs {plain_dx}"
    );
}

/// §17.3.2.13: emboss draws a light-gray relief copy down-right and lifts
/// black text to white — raised off the page.
#[test]
fn emboss_draws_a_relief_copy_down_right_and_lifts_black_to_white() {
    let pages = pages();
    let (copy, main) = copy_and_main(&pages, "EMBS");
    let (dx, dy) = (copy.0 - main.0, copy.1 - main.1);
    assert!(
        dx > 0.0 && dy > 0.0 && (dx - dy).abs() < 1e-4,
        "the relief copy sits down-right: dx={dx}, dy={dy}"
    );
    assert_eq!(copy.2, LIGHT_GRAY);
    assert_eq!(main.2, WHITE, "black text is lifted to white");
}

/// §17.3.2.18: imprint is the mirror — the relief copy sits up-left, the
/// glyphs pressed into the page.
#[test]
fn imprint_draws_its_relief_copy_up_left() {
    let pages = pages();
    let (copy, main) = copy_and_main(&pages, "IMPR");
    let (dx, dy) = (copy.0 - main.0, copy.1 - main.1);
    assert!(
        dx < 0.0 && dy < 0.0 && (dx - dy).abs() < 1e-4,
        "the relief copy sits up-left: dx={dx}, dy={dy}"
    );
    assert_eq!(copy.2, LIGHT_GRAY);
    assert_eq!(main.2, WHITE);
}
