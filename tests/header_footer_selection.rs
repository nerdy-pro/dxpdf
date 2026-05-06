//! End-to-end tests for ECMA-376 §17.10.1 (`evenAndOddHeaders`) and
//! §17.10.6 (`titlePg`) header/footer selection.
//!
//! Each test constructs a `Document` model directly with marker
//! header text per slot, runs the full resolve → layout pipeline,
//! and verifies which slot's text appears on which page.

use std::collections::HashMap;

use dxpdf::model::*;
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};
use dxpdf::render::resolve_and_layout;

// ── helpers ────────────────────────────────────────────────────────────────

fn empty_document() -> Document {
    Document {
        settings: DocumentSettings::default(),
        theme: None,
        styles: StyleSheet::default(),
        numbering: NumberingDefinitions::default(),
        body: vec![],
        final_section: SectionProperties::default(),
        headers: HashMap::new(),
        footers: HashMap::new(),
        footnotes: HashMap::new(),
        endnotes: HashMap::new(),
        media: HashMap::new(),
        embedded_fonts: vec![],
    }
}

fn run_with(elements: Vec<RunElement>) -> Inline {
    Inline::TextRun(Box::new(TextRun {
        style_id: None,
        properties: RunProperties::default(),
        content: elements,
        rsids: RevisionIds::default(),
    }))
}

fn para(text: &str) -> Block {
    Block::Paragraph(Box::new(Paragraph {
        style_id: None,
        properties: ParagraphProperties::default(),
        mark_run_properties: None,
        content: vec![run_with(vec![RunElement::Text(text.to_string())])],
        rsids: ParagraphRevisionIds::default(),
    }))
}

/// A paragraph that begins with a hard page break so it lands on a
/// fresh page regardless of upstream content height.
fn para_after_page_break(text: &str) -> Block {
    Block::Paragraph(Box::new(Paragraph {
        style_id: None,
        properties: ParagraphProperties::default(),
        mark_run_properties: None,
        content: vec![run_with(vec![
            RunElement::PageBreak,
            RunElement::Text(text.to_string()),
        ])],
        rsids: ParagraphRevisionIds::default(),
    }))
}

/// Concatenate every text fragment drawn on a page. Sufficient for
/// substring matching — we're not comparing layout coordinates here,
/// only which marker strings ended up on which page.
fn page_text(page: &LayoutedPage) -> String {
    let mut out = String::new();
    for cmd in &page.commands {
        if let DrawCommand::Text { text, .. } = cmd {
            out.push_str(text);
            out.push(' ');
        }
    }
    out
}

// ── §17.10.6 — titlePg ────────────────────────────────────────────────────

#[test]
fn title_page_uses_first_header_on_page_one_default_after() {
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_first = RelId::new("rF");
    doc.headers
        .insert(r_default.clone(), vec![para("DEFAULT_HEADER_TEXT")]);
    doc.headers
        .insert(r_first.clone(), vec![para("FIRST_HEADER_TEXT")]);
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: Some(r_first),
            even: None,
        },
        title_page: Some(true),
        ..Default::default()
    };
    doc.body = vec![para("body p1"), para_after_page_break("body p2")];

    let (_, pages) = resolve_and_layout(&doc);
    assert_eq!(pages.len(), 2, "expected 2 pages, got {}", pages.len());

    let p1 = page_text(&pages[0]);
    assert!(
        p1.contains("FIRST_HEADER_TEXT"),
        "page 1 must show first header; got: {p1:?}"
    );
    assert!(
        !p1.contains("DEFAULT_HEADER_TEXT"),
        "page 1 must NOT show default header; got: {p1:?}"
    );

    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("DEFAULT_HEADER_TEXT"),
        "page 2 must show default header; got: {p2:?}"
    );
    assert!(
        !p2.contains("FIRST_HEADER_TEXT"),
        "page 2 must NOT show first header; got: {p2:?}"
    );
}

#[test]
fn title_page_without_first_slot_blanks_page_one() {
    // §17.10.6 literal: when titlePg is set but the section has no
    // `first` reference, page 1 has no header. This is the rule we
    // need to render this is the test that covers the
    // vorlage_baustellenkoordinator_v12 case (`<w:titlePg/>` with a
    // `first` reference that points to an effectively-empty header).
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    doc.headers
        .insert(r_default.clone(), vec![para("DEFAULT_HEADER_TEXT")]);
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: None,
        },
        title_page: Some(true),
        ..Default::default()
    };
    doc.body = vec![para("body p1"), para_after_page_break("body p2")];

    let (_, pages) = resolve_and_layout(&doc);
    let p1 = page_text(&pages[0]);
    assert!(
        !p1.contains("DEFAULT_HEADER_TEXT"),
        "title page must be blank, not fall back to default; got: {p1:?}"
    );
    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("DEFAULT_HEADER_TEXT"),
        "page 2 still shows default; got: {p2:?}"
    );
}

#[test]
fn title_page_flag_off_keeps_default_on_page_one() {
    // Sanity: even with a `first` slot, missing `<w:titlePg/>` means
    // every page uses default.
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_first = RelId::new("rF");
    doc.headers
        .insert(r_default.clone(), vec![para("DEFAULT_HEADER_TEXT")]);
    doc.headers
        .insert(r_first.clone(), vec![para("FIRST_HEADER_TEXT")]);
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: Some(r_first),
            even: None,
        },
        title_page: None,
        ..Default::default()
    };
    doc.body = vec![para("body p1"), para_after_page_break("body p2")];

    let (_, pages) = resolve_and_layout(&doc);
    for (i, page) in pages.iter().enumerate() {
        let t = page_text(page);
        assert!(
            t.contains("DEFAULT_HEADER_TEXT"),
            "page {i} must show default header without titlePg; got: {t:?}"
        );
        assert!(
            !t.contains("FIRST_HEADER_TEXT"),
            "page {i} must NOT show first header without titlePg; got: {t:?}"
        );
    }
}

// ── §17.10.1 — evenAndOddHeaders ──────────────────────────────────────────

#[test]
fn even_and_odd_alternates_headers_across_three_pages() {
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_even = RelId::new("rE");
    doc.headers
        .insert(r_default.clone(), vec![para("ODD_HEADER")]);
    doc.headers.insert(r_even.clone(), vec![para("EVEN_HEADER")]);
    doc.settings.even_and_odd_headers = true;
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: Some(r_even),
        },
        title_page: None,
        ..Default::default()
    };
    doc.body = vec![
        para("p1"),
        para_after_page_break("p2"),
        para_after_page_break("p3"),
    ];

    let (_, pages) = resolve_and_layout(&doc);
    assert_eq!(pages.len(), 3);

    let p1 = page_text(&pages[0]);
    assert!(p1.contains("ODD_HEADER"), "page 1 (odd) → odd header; got {p1:?}");
    assert!(!p1.contains("EVEN_HEADER"));

    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("EVEN_HEADER"),
        "page 2 (even) → even header; got {p2:?}"
    );
    assert!(!p2.contains("ODD_HEADER"));

    let p3 = page_text(&pages[2]);
    assert!(p3.contains("ODD_HEADER"), "page 3 (odd) → odd header; got {p3:?}");
}

#[test]
fn even_and_odd_disabled_keeps_default_on_every_page() {
    // `even` slot is populated but the document setting is off →
    // every page uses default.
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_even = RelId::new("rE");
    doc.headers
        .insert(r_default.clone(), vec![para("ODD_HEADER")]);
    doc.headers.insert(r_even.clone(), vec![para("EVEN_HEADER")]);
    doc.settings.even_and_odd_headers = false;
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: Some(r_even),
        },
        ..Default::default()
    };
    doc.body = vec![para("p1"), para_after_page_break("p2")];

    let (_, pages) = resolve_and_layout(&doc);
    for (i, page) in pages.iter().enumerate() {
        let t = page_text(page);
        assert!(
            t.contains("ODD_HEADER"),
            "page {i} must use default; got {t:?}"
        );
        assert!(
            !t.contains("EVEN_HEADER"),
            "page {i} must not use even slot when flag is off; got {t:?}"
        );
    }
}

#[test]
fn even_and_odd_with_no_even_slot_blanks_even_pages() {
    // Spec: when evenAndOddHeaders is on but the section has no
    // `even` reference, the even page has a *blank* header — not
    // a fall-through to default.
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    doc.headers
        .insert(r_default.clone(), vec![para("ODD_HEADER")]);
    doc.settings.even_and_odd_headers = true;
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: None,
        },
        ..Default::default()
    };
    doc.body = vec![para("p1"), para_after_page_break("p2")];

    let (_, pages) = resolve_and_layout(&doc);
    let p1 = page_text(&pages[0]);
    assert!(p1.contains("ODD_HEADER"));
    let p2 = page_text(&pages[1]);
    assert!(
        !p2.contains("ODD_HEADER"),
        "page 2 must be blank when `even` slot is empty; got {p2:?}"
    );
}

// ── Combined: titlePg + evenAndOddHeaders ─────────────────────────────────

#[test]
fn title_page_takes_precedence_over_even_and_odd_on_page_one() {
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_first = RelId::new("rF");
    let r_even = RelId::new("rE");
    doc.headers
        .insert(r_default.clone(), vec![para("ODD_HEADER")]);
    doc.headers
        .insert(r_first.clone(), vec![para("FIRST_HEADER")]);
    doc.headers.insert(r_even.clone(), vec![para("EVEN_HEADER")]);
    doc.settings.even_and_odd_headers = true;
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: Some(r_first),
            even: Some(r_even),
        },
        title_page: Some(true),
        ..Default::default()
    };
    doc.body = vec![
        para("p1"),
        para_after_page_break("p2"),
        para_after_page_break("p3"),
    ];

    let (_, pages) = resolve_and_layout(&doc);
    assert_eq!(pages.len(), 3);

    let p1 = page_text(&pages[0]);
    assert!(
        p1.contains("FIRST_HEADER"),
        "page 1 (titlePg + even/odd both on) → first wins; got {p1:?}"
    );
    assert!(!p1.contains("ODD_HEADER"));
    assert!(!p1.contains("EVEN_HEADER"));

    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("EVEN_HEADER"),
        "page 2 → even (titlePg only fires on page 1); got {p2:?}"
    );

    let p3 = page_text(&pages[2]);
    assert!(p3.contains("ODD_HEADER"), "page 3 → odd default; got {p3:?}");
}

// ── Footers — same selection rules ────────────────────────────────────────

#[test]
fn footer_selection_follows_same_rules_as_header() {
    let mut doc = empty_document();
    let r_default = RelId::new("fD");
    let r_first = RelId::new("fF");
    doc.footers
        .insert(r_default.clone(), vec![para("DEFAULT_FOOTER")]);
    doc.footers
        .insert(r_first.clone(), vec![para("FIRST_FOOTER")]);
    doc.final_section = SectionProperties {
        footer_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: Some(r_first),
            even: None,
        },
        title_page: Some(true),
        ..Default::default()
    };
    doc.body = vec![para("p1"), para_after_page_break("p2")];

    let (_, pages) = resolve_and_layout(&doc);
    let p1 = page_text(&pages[0]);
    assert!(p1.contains("FIRST_FOOTER"), "page 1 footer; got {p1:?}");
    assert!(!p1.contains("DEFAULT_FOOTER"));
    let p2 = page_text(&pages[1]);
    assert!(p2.contains("DEFAULT_FOOTER"), "page 2 footer; got {p2:?}");
    assert!(!p2.contains("FIRST_FOOTER"));
}
