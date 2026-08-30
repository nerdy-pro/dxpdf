//! Issue #154 end-to-end — comment rendering over `test-files/comments.docx`
//! (two comments, two authors, a range spanning paragraphs), its
//! `comments-hidden.docx` twin (`w:revisionView w:comments="0"`), and the
//! corpus fixture `comment-reference.docx` (a real Word package, Cyrillic
//! author, narrow margin).
//!
//! Structural assertions over the draw-command stream: balloon text lands in
//! the right margin band, ranges get the wash, authors get their palette
//! colors, and the hidden view produces none of it.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};
use dxpdf::render::resolve::revision::{COMMENT_RANGE_SHADING, REVISION_PALETTE};

fn layout(name: &str) -> Vec<LayoutedPage> {
    let path = format!("{}/test-files/{name}", env!("CARGO_MANIFEST_DIR"));
    let data = std::fs::read(&path).unwrap_or_else(|e| panic!("failed to read {path}: {e}"));
    let doc = dxpdf::docx::parse(&data).expect("fixture must parse");
    dxpdf::render::resolve_and_layout(doc).1
}

fn text_cmd<'a>(pages: &'a [LayoutedPage], needle: &str) -> (&'a DrawCommand, f32) {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.contains(needle) => {
                Some((c, f32::from(position.x)))
            }
            _ => None,
        })
        .next()
        .unwrap_or_else(|| panic!("{needle:?} must reach paint"))
}

/// The balloon body reaches the page, inside the right margin band, and the
/// comment's whole content is there — both paragraphs of Ann's comment.
#[test]
fn balloons_land_in_the_margin_band() {
    let pages = layout("comments.docx");
    // Letter page 612pt wide, right margin 2160tw = 108pt → band starts at 504.
    let band_start = 612.0 - 108.0;

    for needle in ["balloon", "Second", "note."] {
        let (_, x) = text_cmd(&pages, needle);
        assert!(
            x >= band_start,
            "{needle:?} must render in the margin band (x={x}, band starts {band_start})"
        );
    }
    // Author labels, colored by the shared revision palette.
    let (ann, ann_x) = text_cmd(&pages, "Ann");
    assert!(ann_x >= band_start);
    if let DrawCommand::Text { color, .. } = ann {
        assert_eq!(*color, REVISION_PALETTE[0], "first author's color");
    }
    let (bob, _) = text_cmd(&pages, "Bob");
    if let DrawCommand::Text { color, .. } = bob {
        assert_eq!(*color, REVISION_PALETTE[1], "second author's color");
    }
}

/// The commented range gets the wash — including across the paragraph
/// boundary — and the control paragraph does not.
#[test]
fn commented_ranges_are_washed() {
    let pages = layout("comments.docx");
    let washed_behind = |needle: &str| -> bool {
        let (_, x) = text_cmd(&pages, needle);
        let (cmd, _) = text_cmd(&pages, needle);
        let y = match cmd {
            DrawCommand::Text { position, .. } => f32::from(position.y),
            _ => unreachable!(),
        };
        pages.iter().flat_map(|p| &p.commands).any(|c| match c {
            DrawCommand::Rect { rect, color } => {
                *color == COMMENT_RANGE_SHADING
                    && f32::from(rect.origin.x) <= x + 1.0
                    && f32::from(rect.origin.x) + f32::from(rect.size.width) >= x + 1.0
                    && f32::from(rect.origin.y) <= y
                    && f32::from(rect.origin.y) + f32::from(rect.size.height) >= y - 8.0
            }
            _ => false,
        })
    };
    assert!(washed_behind("range"), "in-paragraph range");
    assert!(washed_behind("spans"), "the run after commentRangeStart");
    assert!(
        washed_behind("closes"),
        "the stamp must survive the paragraph boundary"
    );
    assert!(!washed_behind("Control"), "control stays clean");
    assert!(
        !washed_behind("Before "),
        "text before the range stays clean"
    );
}

/// A connector line runs from the anchor into the band.
#[test]
fn a_connector_reaches_from_anchor_to_balloon() {
    let pages = layout("comments.docx");
    let band_start = 612.0 - 108.0;
    let connector = pages.iter().flat_map(|p| &p.commands).any(|c| match c {
        DrawCommand::Line { line, .. } => {
            f32::from(line.start.x) < band_start && f32::from(line.end.x) >= band_start
        }
        _ => false,
    });
    assert!(connector, "a line must bridge body and balloon");
}

/// `w:revisionView w:comments="0"`: same body, nothing comment-shaped on the
/// page — no balloon text, no wash, and the body text itself is unmoved.
#[test]
fn the_hidden_view_draws_no_comment_marks() {
    let pages = layout("comments-hidden.docx");
    assert!(
        !pages.iter().flat_map(|p| &p.commands).any(|c| match c {
            DrawCommand::Text { text, .. } => text.contains("balloon"),
            _ => false,
        }),
        "no balloon content"
    );
    assert!(
        !pages.iter().flat_map(|p| &p.commands).any(|c| match c {
            DrawCommand::Rect { color, .. } => *color == COMMENT_RANGE_SHADING,
            _ => false,
        }),
        "no range wash"
    );
    // The document's own text still renders.
    text_cmd(&pages, "range");
    text_cmd(&pages, "Control");
}

/// The corpus fixture — a real Word package with modern-comments sibling
/// parts and a Cyrillic author. Its margin is narrow (850tw ≈ 42.5pt), which
/// still clears the 24pt band floor, so the balloon renders.
#[test]
fn the_word_authored_fixture_renders_its_comment() {
    let pages = layout("comment-reference.docx");
    let (_, x) = text_cmd(&pages, "111");
    // A4 width 11906tw ≈ 595.3pt; right margin 850tw = 42.5pt.
    assert!(
        x > 595.3 - 42.5,
        "the comment body must render in the margin band (x={x})"
    );
    text_cmd(&pages, "Автор");
    // And the commented body text still paints where it was.
    let (_, body_x) = text_cmd(&pages, "Тест");
    assert!(body_x < 500.0);
}
