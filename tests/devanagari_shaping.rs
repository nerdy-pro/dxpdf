//! Issue #153 end-to-end — complex-script shaping with the shaped cluster as
//! the spacing unit, over `test-files/devanagari.docx`.
//!
//! Like `tests/font_fallback.rs`, every assertion is structural: a run is
//! *marked shaped*, spacing *reaches the command*, the Latin control *stays
//! off the shaping path* — never that any particular host font drew anything.
//! Hosts without Devanagari or Arabic coverage skip the affected tests (the
//! ubuntu CI runner has no Devanagari face), following the probe-and-eprintln
//! pattern this suite shares with `emoji_e2e.rs`. The glyph-order reorder
//! itself — the matra standing left of its consonant — is pinned against a
//! real face by the §20.1.9.3-style unit tests in `render::shape`, which skip
//! the same way.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};
use dxpdf::render::shape::RunDirection;
use skia_safe::{FontMgr, FontStyle};

fn fixture() -> dxpdf::model::Document {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files/devanagari.docx");
    let data = std::fs::read(path).unwrap_or_else(|e| {
        panic!("failed to read {path}: {e} — rebuild with scripts/make_devanagari_fixture.py")
    });
    dxpdf::docx::parse(&data).expect("fixture must parse")
}

/// Whether any face on this host can draw `ch` — the graceful-skip probe.
fn host_covers(ch: char) -> bool {
    FontMgr::new()
        .match_family_style_character("Times New Roman", FontStyle::normal(), &[], ch as i32)
        .is_some_and(|t| t.unichar_to_glyph(ch as i32) != 0)
}

/// Every `DrawCommand::Text` as (text, shaped, char_spacing in pt).
fn text_commands(pages: &[LayoutedPage]) -> Vec<(String, Option<RunDirection>, f32)> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text {
                text,
                shaped,
                char_spacing,
                ..
            } => Some((text.to_string(), *shaped, f32::from(*char_spacing))),
            _ => None,
        })
        .collect()
}

fn layout() -> Vec<LayoutedPage> {
    let (_, pages) = dxpdf::render::resolve_and_layout(fixture());
    pages
}

/// Paragraph 1: every Devanagari run is marked for left-to-right shaping —
/// the predicate half of issue #153, through the full pipeline (including the
/// per-glyph font fallback that must run first, since the document names no
/// font that covers Devanagari).
#[test]
fn devanagari_runs_are_shaped_left_to_right() {
    if !host_covers('\u{0915}') {
        eprintln!("skipping: no face on this host covers U+0915");
        return;
    }
    let commands = text_commands(&layout());
    let devanagari: Vec<_> = commands
        .iter()
        .filter(|(text, ..)| text.contains('\u{0915}') || text.contains('\u{0939}'))
        .collect();
    assert!(
        !devanagari.is_empty(),
        "the fixture's Devanagari must reach paint"
    );
    for (text, shaped, _) in devanagari {
        assert_eq!(
            *shaped,
            Some(RunDirection::LeftToRight),
            "{text:?} must be marked shaped, left to right"
        );
    }
}

/// Paragraph 4: the Latin control keeps the cmap path — shaping Devanagari
/// must not have widened the predicate — and keeps its authored 2pt spacing.
#[test]
fn the_latin_control_stays_on_the_cmap_path_with_its_spacing() {
    let commands = text_commands(&layout());
    let control = commands
        .iter()
        .find(|(text, ..)| text.contains("Latin"))
        .expect("the Latin control paragraph must reach paint");
    assert_eq!(control.1, None, "Latin must not be shaped");
    assert!(
        (control.2 - 2.0).abs() < 1e-4,
        "w:spacing w:val=\"40\" is 2pt, got {}",
        control.2
    );
}

/// Paragraph 2: authored §17.3.2.35 spacing survives onto a shaped command —
/// the painter opens it between shaped clusters, so it must arrive there.
#[test]
fn authored_spacing_reaches_a_shaped_command() {
    if !host_covers('\u{0915}') {
        eprintln!("skipping: no face on this host covers U+0915");
        return;
    }
    let commands = text_commands(&layout());
    assert!(
        commands
            .iter()
            .any(|(text, shaped, spacing)| text.contains('\u{0915}')
                && shaped.is_some()
                && (spacing - 2.0).abs() < 1e-4),
        "a shaped Devanagari command must carry the authored 2pt spacing"
    );
}

/// Paragraph 3: §17.3.1.13 `distribute` reaches a shaped Devanagari run as
/// distribution extra — the paragraph authors no `w:spacing`, so any spacing
/// on its commands is the line's slack, shared between shaped clusters.
#[test]
fn distribution_stretches_a_shaped_devanagari_line() {
    if !host_covers('\u{092E}') {
        eprintln!("skipping: no face on this host covers U+092E");
        return;
    }
    let commands = text_commands(&layout());
    let mat = commands
        .iter()
        .find(|(text, ..)| text.starts_with('\u{092E}'))
        .expect("the distribute paragraph's first word must reach paint");
    assert_eq!(mat.1, Some(RunDirection::LeftToRight));
    assert!(
        mat.2 > 0.0,
        "the line's slack must reach the shaped command as spacing, got {}",
        mat.2
    );
}

/// Paragraph 5: the same for Arabic — the regression pin for the repaired
/// defect where distribution widened a shaped run's decorations while the
/// painter ignored the spacing for its glyphs.
#[test]
fn distribution_stretches_a_shaped_arabic_line() {
    if !host_covers('\u{0645}') {
        eprintln!("skipping: no face on this host covers U+0645");
        return;
    }
    let commands = text_commands(&layout());
    let arabic: Vec<_> = commands
        .iter()
        .filter(|(text, ..)| text.contains('\u{0645}'))
        .collect();
    assert!(!arabic.is_empty(), "the Arabic paragraph must reach paint");
    for (text, shaped, spacing) in arabic {
        assert_eq!(
            *shaped,
            Some(RunDirection::RightToLeft),
            "{text:?} must be shaped right to left"
        );
        assert!(
            *spacing > 0.0,
            "{text:?} must carry the distributed slack, got {spacing}"
        );
    }
}
