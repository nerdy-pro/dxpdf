//! Issue #154 end-to-end — tracked-change marks over
//! `test-files/tracked-changes.docx` (markup view) and
//! `tracked-changes-final.docx` (same body, `w:revisionView w:insDel="0"`).
//!
//! Assertions are structural over the draw-command stream: which text reaches
//! paint, in which color, and which decoration lines accompany it. The
//! revision palette is this engine's own (`resolve::revision`), so its RGB
//! literals are fair to pin; nothing else about the host is.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};
use dxpdf::render::resolve::color::RgbColor;
use dxpdf::render::resolve::revision::REVISION_PALETTE;

fn layout(name: &str) -> Vec<LayoutedPage> {
    let path = format!("{}/test-files/{name}", env!("CARGO_MANIFEST_DIR"));
    let data = std::fs::read(&path).unwrap_or_else(|e| {
        panic!("failed to read {path}: {e} — rebuild with scripts/make_tracked_changes_fixtures.py")
    });
    let doc = dxpdf::docx::parse(&data).expect("fixture must parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Every Text command as (text, color, x, y).
fn texts(pages: &[LayoutedPage]) -> Vec<(String, RgbColor, f32, f32)> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text {
                text,
                color,
                position,
                ..
            } => Some((
                text.to_string(),
                *color,
                f32::from(position.x),
                f32::from(position.y),
            )),
            _ => None,
        })
        .collect()
}

fn find<'a>(
    texts: &'a [(String, RgbColor, f32, f32)],
    needle: &str,
) -> &'a (String, RgbColor, f32, f32) {
    texts
        .iter()
        .find(|(t, ..)| t.contains(needle))
        .unwrap_or_else(|| panic!("{needle:?} must reach paint"))
}

/// Whether some Line command crosses the given x at a y *above* the given
/// baseline — a strike through that run.
fn line_above(pages: &[LayoutedPage], x: f32, baseline: f32) -> bool {
    pages.iter().flat_map(|p| &p.commands).any(|c| match c {
        DrawCommand::Line { line, .. } => {
            f32::from(line.start.x) <= x + 1.0
                && f32::from(line.end.x) >= x + 1.0
                && f32::from(line.start.y) < baseline
                && f32::from(line.start.y) > baseline - 12.0
        }
        _ => false,
    })
}

/// Whether some Underline command crosses the given x near the baseline.
fn underline_at(pages: &[LayoutedPage], x: f32, baseline: f32) -> bool {
    pages.iter().flat_map(|p| &p.commands).any(|c| match c {
        DrawCommand::Underline { line, .. } => {
            f32::from(line.start.x) <= x + 1.0
                && f32::from(line.end.x) >= x + 1.0
                && (f32::from(line.start.y) - baseline).abs() < 6.0
        }
        _ => false,
    })
}

/// Markup view (no `w:revisionView`): deletions paint struck through and
/// insertions underlined, each in its author's palette color; untouched text
/// stays black.
#[test]
fn markup_view_marks_insertions_and_deletions() {
    let pages = layout("tracked-changes.docx");
    let texts = texts(&pages);

    let ins = find(&texts, "beta");
    assert_eq!(ins.1, REVISION_PALETTE[0], "Ann is the first author seen");
    assert!(
        underline_at(&pages, ins.2, ins.3),
        "an insertion is underlined"
    );

    let del = find(&texts, "epsilon");
    assert_eq!(del.1, REVISION_PALETTE[0], "same author, same color");
    assert!(
        line_above(&pages, del.2, del.3),
        "a deletion is struck through"
    );

    let second = texts
        .iter()
        .find(|(t, ..)| t.trim() == "eta")
        .expect("Bob's insertion must reach paint");
    assert_eq!(second.1, REVISION_PALETTE[1], "Bob is the second author");

    let control = find(&texts, "Theta");
    assert_eq!(control.1, RgbColor::BLACK, "untouched text keeps its color");
    assert!(
        !underline_at(&pages, control.2, control.3),
        "and no decoration"
    );
}

/// Final view (`w:revisionView w:insDel="0"`): the deletion is not on the
/// page at all, the insertion paints plain — while ordinary §17.3.2.37
/// strike *formatting* still paints struck, which is what tells a revision
/// mark apart from formatting.
#[test]
fn final_view_suppresses_deletions_and_unmarks_insertions() {
    let pages = layout("tracked-changes-final.docx");
    let texts = texts(&pages);

    assert!(
        !texts.iter().any(|(t, ..)| t.contains("epsilon")),
        "an unaccepted deletion must not paint in the final view"
    );

    let ins = find(&texts, "beta");
    assert_eq!(ins.1, RgbColor::BLACK, "insertion renders plain");
    assert!(!underline_at(&pages, ins.2, ins.3), "and unmarked");

    // The neighbours of the suppressed run close up: both still paint.
    find(&texts, "Delta");
    find(&texts, "zeta");

    let struck = find(&texts, "struck");
    assert!(
        line_above(&pages, struck.2, struck.3),
        "w:strike is formatting, not a revision — it stays struck"
    );
}

/// §17.3.2.9 in both fixtures: dstrike draws two lines through its run.
#[test]
fn dstrike_draws_two_lines() {
    let pages = layout("tracked-changes.docx");
    let texts = texts(&pages);
    let d = find(&texts, "double");
    let count = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter(|c| match c {
            DrawCommand::Line { line, .. } => {
                f32::from(line.start.x) <= d.2 + 1.0
                    && f32::from(line.end.x) >= d.2 + 1.0
                    && f32::from(line.start.y) < d.3
                    && f32::from(line.start.y) > d.3 - 12.0
            }
            _ => false,
        })
        .count();
    assert_eq!(count, 2, "double strike is two lines");
}
