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
    doc.headers
        .insert(r_even.clone(), vec![para("EVEN_HEADER")]);
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
    assert!(
        p1.contains("ODD_HEADER"),
        "page 1 (odd) → odd header; got {p1:?}"
    );
    assert!(!p1.contains("EVEN_HEADER"));

    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("EVEN_HEADER"),
        "page 2 (even) → even header; got {p2:?}"
    );
    assert!(!p2.contains("ODD_HEADER"));

    let p3 = page_text(&pages[2]);
    assert!(
        p3.contains("ODD_HEADER"),
        "page 3 (odd) → odd header; got {p3:?}"
    );
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
    doc.headers
        .insert(r_even.clone(), vec![para("EVEN_HEADER")]);
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
    doc.headers
        .insert(r_even.clone(), vec![para("EVEN_HEADER")]);
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
    assert!(
        p3.contains("ODD_HEADER"),
        "page 3 → odd default; got {p3:?}"
    );
}

// ── Footers — same selection rules ────────────────────────────────────────

// ── §17.6.12 — pgNumType.start drives logical numbering ──────────────────

#[test]
fn pg_num_type_start_two_makes_first_page_even_for_selection() {
    // §17.6.12 + §17.10.1: when the section sets `pgNumType.start=2`,
    // the first physical page is logical page 2 — even — so the `even`
    // header applies on page 1, not the default. This is the rule we
    // need for documents that count from a non-1 start.
    use dxpdf::model::PageNumberType;
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_even = RelId::new("rE");
    doc.headers
        .insert(r_default.clone(), vec![para("ODD_HEADER")]);
    doc.headers
        .insert(r_even.clone(), vec![para("EVEN_HEADER")]);
    doc.settings.even_and_odd_headers = true;
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: Some(r_even),
        },
        page_number_type: Some(PageNumberType {
            format: None,
            start: Some(2),
            chap_style: None,
            chap_sep: None,
        }),
        ..Default::default()
    };
    doc.body = vec![para("p1"), para_after_page_break("p2")];

    let (_, pages) = resolve_and_layout(&doc);
    assert_eq!(pages.len(), 2);
    let p1 = page_text(&pages[0]);
    assert!(
        p1.contains("EVEN_HEADER"),
        "physical page 1 = logical page 2 (even) → even header; got {p1:?}"
    );
    assert!(!p1.contains("ODD_HEADER"));
    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("ODD_HEADER"),
        "physical page 2 = logical page 3 (odd) → default; got {p2:?}"
    );
}

#[test]
fn pg_num_type_start_renders_in_page_field_in_header() {
    // PAGE field in the header must render the **logical** page number
    // (matches Word). Section starts at 5 → page 1's PAGE field is "5".
    use dxpdf::field::{CommonSwitches, FieldInstruction};
    use dxpdf::model::{Field, PageNumberType};
    let mut doc = empty_document();
    let r_default = RelId::new("rD");

    // Header consists of a single PAGE field with cached value "999"
    // (so we'd notice if the renderer was using the cache instead).
    let page_field_para = Block::Paragraph(Box::new(Paragraph {
        style_id: None,
        properties: ParagraphProperties::default(),
        mark_run_properties: None,
        content: vec![Inline::Field(Field {
            instruction: FieldInstruction::Page {
                switches: CommonSwitches::default(),
            },
            content: vec![run_with(vec![RunElement::Text("999".into())])],
        })],
        rsids: ParagraphRevisionIds::default(),
    }));
    doc.headers.insert(r_default.clone(), vec![page_field_para]);
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: None,
        },
        page_number_type: Some(PageNumberType {
            format: None,
            start: Some(5),
            chap_style: None,
            chap_sep: None,
        }),
        ..Default::default()
    };
    doc.body = vec![para("p1"), para_after_page_break("p2")];

    let (_, pages) = resolve_and_layout(&doc);
    let p1 = page_text(&pages[0]);
    let p2 = page_text(&pages[1]);
    assert!(
        p1.contains("5"),
        "page 1 PAGE field renders logical 5 (start); got {p1:?}",
    );
    assert!(
        p2.contains("6"),
        "page 2 PAGE field renders logical 6; got {p2:?}",
    );
    assert!(
        !p1.contains("999"),
        "stale cached value must not appear; got {p1:?}",
    );
}

#[test]
fn pg_num_type_continues_across_sections_without_start() {
    // Two sections, both without `pgNumType.start`. Section 1 has 1
    // page (logical 1), section 2 has 2 pages (logical 2, 3). With
    // even/odd on, page 2 (logical 2) uses `even`.
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_even = RelId::new("rE");
    doc.headers.insert(r_default.clone(), vec![para("ODD")]);
    doc.headers.insert(r_even.clone(), vec![para("EVEN")]);
    doc.settings.even_and_odd_headers = true;

    let s1_break = SectionProperties {
        section_type: Some(SectionType::NextPage),
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default.clone()),
            first: None,
            even: Some(r_even.clone()),
        },
        ..Default::default()
    };
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: Some(r_even),
        },
        ..Default::default()
    };
    doc.body = vec![
        para("S1"),
        Block::SectionBreak(Box::new(s1_break)),
        para("S2 p1"),
        para_after_page_break("S2 p2"),
    ];

    let (_, pages) = resolve_and_layout(&doc);
    assert_eq!(pages.len(), 3, "expected 3 pages, got {}", pages.len());
    assert!(page_text(&pages[0]).contains("ODD"), "logical 1 → odd");
    assert!(
        page_text(&pages[1]).contains("EVEN"),
        "logical 2 → even (continued across section)",
    );
    assert!(page_text(&pages[2]).contains("ODD"), "logical 3 → odd");
}

#[test]
fn pg_num_type_start_resets_on_second_section() {
    // Section 1 produces 1 page (logical 1). Section 2 sets
    // `pgNumType.start=10` — its first page is logical 10 (even),
    // not the would-be continuation value 2.
    use dxpdf::model::PageNumberType;
    let mut doc = empty_document();
    let r_default = RelId::new("rD");
    let r_even = RelId::new("rE");
    doc.headers
        .insert(r_default.clone(), vec![para("ODD_HEADER")]);
    doc.headers
        .insert(r_even.clone(), vec![para("EVEN_HEADER")]);
    doc.settings.even_and_odd_headers = true;

    let s1_break = SectionProperties {
        section_type: Some(SectionType::NextPage),
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default.clone()),
            first: None,
            even: Some(r_even.clone()),
        },
        ..Default::default()
    };
    doc.final_section = SectionProperties {
        header_refs: SectionHeaderFooterRefs {
            default: Some(r_default),
            first: None,
            even: Some(r_even),
        },
        page_number_type: Some(PageNumberType {
            format: None,
            start: Some(10),
            chap_style: None,
            chap_sep: None,
        }),
        ..Default::default()
    };
    doc.body = vec![
        para("S1"),
        Block::SectionBreak(Box::new(s1_break)),
        para("S2 p1"),
    ];

    let (_, pages) = resolve_and_layout(&doc);
    assert_eq!(pages.len(), 2);
    let p1 = page_text(&pages[0]);
    assert!(
        p1.contains("ODD_HEADER"),
        "S1 page 1 = logical 1; got {p1:?}"
    );
    let p2 = page_text(&pages[1]);
    assert!(
        p2.contains("EVEN_HEADER"),
        "S2 page 1 = logical 10 (even due to start); got {p2:?}",
    );
}

// ── Real-document regression ─────────────────────────────────────────────

/// Regression for the `vorlage_baustellenkoordinator_v12.docx` issue:
/// the section uses `<w:titlePg/>` plus a `first` headerReference that
/// points to an empty header part (`header3.xml`). MS Word renders
/// page 1 with no header text and page 2+ with the default header
/// ("Begehungsprotokoll Baustellenkoordinator vom #datum_begehung#").
/// Pre-fix dxpdf rendered the default header on page 1 too.
///
/// `test-cases/` is untracked, so this test is gated on file
/// existence; CI without the fixture passes as a no-op while a local
/// run with the fixture exercises the full parse → resolve → layout
/// chain on a real customer document.
#[test]
fn vorlage_baustellenkoordinator_v12_page_one_header_is_blank() {
    let path = std::path::Path::new("test-cases/vorlage_baustellenkoordinator_v12.docx");
    if !path.exists() {
        eprintln!("SKIPPED: {} not present", path.display());
        return;
    }
    let bytes = std::fs::read(path).expect("read fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse fixture");
    let (_, pages) = resolve_and_layout(&doc);
    assert!(
        pages.len() >= 2,
        "expected at least 2 pages, got {}",
        pages.len()
    );

    // Body top margin in this section is 1417 twips ≈ 70.85pt; using
    // y < 70 isolates the strict header zone (between header margin
    // ~35pt and body top).
    let header_text_on = |page_idx: usize| -> String {
        let mut out = String::new();
        for cmd in &pages[page_idx].commands {
            if let DrawCommand::Text { text, position, .. } = cmd {
                if position.y.raw() < 70.0 {
                    out.push_str(text);
                    out.push(' ');
                }
            }
        }
        out
    };

    let header_p1 = header_text_on(0);
    let header_p2 = header_text_on(1);

    // Word/spec: the title page has a blank header — its `first` slot
    // is the (empty) header3.xml, so no body of `header2.xml`'s
    // "Begehungsprotokoll …" text should reach the strict header zone.
    assert!(
        !header_p1.contains("Begehungs"),
        "page 1 strict header zone must be blank (`first` slot is empty), got: {header_p1:?}"
    );

    // Page 2 falls back to default = header2.xml. Word splits the
    // word "Begehungsprotokoll" with whitespace ("Begehungs protokoll"),
    // so check the prefix.
    assert!(
        header_p2.contains("Begehungs"),
        "page 2 must show the default header, got: {header_p2:?}"
    );
}

/// Regression for per-part image-rId collisions — the same fixture
/// has overlapping `rId1` references in `header2.xml.rels`,
/// `header3.xml.rels`, and `footer3.xml.rels`. Pre-fix, the global
/// `media` map only ever held one entry under "rId1" (whichever rel
/// was loaded last), so the first-page header silently displayed the
/// footer's image. The fix synthesizes per-part unique rIds; this
/// test pins down the contract by inspecting the parsed document
/// directly.
#[test]
fn vorlage_baustellenkoordinator_v12_per_part_image_rels_do_not_collide() {
    let path = std::path::Path::new("test-cases/vorlage_baustellenkoordinator_v12.docx");
    if !path.exists() {
        eprintln!("SKIPPED: {} not present", path.display());
        return;
    }
    let bytes = std::fs::read(path).expect("read fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse fixture");

    // Every image rel from a header/footer part must have been keyed
    // under a synthesized unique id (containing the part path), not
    // a bare numeric rId. The bare rId would mean it shared a key
    // with another part — exactly the collision we're guarding
    // against.
    assert!(
        !doc.media.contains_key(&dxpdf::model::RelId::new("rId1")),
        "no media entry should be keyed by bare 'rId1' — that namespace \
         is shared across parts"
    );

    // Header part `word/header3.xml` is the title-page header; its
    // logo must be a distinct media entry from the footer that
    // happens to also use rId1 in its own rels file.
    let header3_logo = doc
        .media
        .get(&dxpdf::model::RelId::new("word/header3.xml::rId1"))
        .expect("header3's image must have its own synthesized media entry");
    let footer3_image = doc
        .media
        .get(&dxpdf::model::RelId::new("word/footer3.xml::rId1"))
        .expect("footer3's image must have its own synthesized media entry");
    assert_ne!(
        header3_logo.0.len(),
        footer3_image.0.len(),
        "header and footer images must point at different bytes — if \
         they're equal we may have re-introduced the rId collision"
    );
}

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
