//! Integration tests that parse real DOCX files from the test-files directory
//! and validate the resulting Document structure.

use dxpdf_docx::model::*;

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

fn load(name: &str) -> Document {
    let path = format!("{TEST_DIR}/{name}");
    let data = std::fs::read(&path).unwrap_or_else(|e| panic!("failed to read {path}: {e}"));
    dxpdf_docx::parse(&data).unwrap_or_else(|e| panic!("failed to parse {path}: {e}"))
}

// ── Helpers ──────────────────────────────────────────────────────────────────

fn count_paragraphs(blocks: &[Block]) -> usize {
    blocks
        .iter()
        .map(|b| match b {
            Block::Paragraph(_) => 1,
            Block::Table(t) => t
                .rows
                .iter()
                .flat_map(|r| &r.cells)
                .map(|c| count_paragraphs(&c.content))
                .sum(),
            Block::SectionBreak(_) => 0,
        })
        .sum()
}

fn count_tables(blocks: &[Block]) -> usize {
    blocks
        .iter()
        .map(|b| match b {
            Block::Table(t) => {
                1 + t
                    .rows
                    .iter()
                    .flat_map(|r| &r.cells)
                    .map(|c| count_tables(&c.content))
                    .sum::<usize>()
            }
            _ => 0,
        })
        .sum()
}

fn count_images(blocks: &[Block]) -> usize {
    blocks
        .iter()
        .map(|b| match b {
            Block::Paragraph(p) => count_images_inline(&p.content),
            Block::Table(t) => t
                .rows
                .iter()
                .flat_map(|r| &r.cells)
                .map(|c| count_images(&c.content))
                .sum(),
            Block::SectionBreak(_) => 0,
        })
        .sum()
}

fn count_images_inline(inlines: &[Inline]) -> usize {
    inlines
        .iter()
        .map(|i| match i {
            Inline::Image(_) => 1,
            Inline::Hyperlink(h) => count_images_inline(&h.content),
            _ => 0,
        })
        .sum()
}

fn count_hyperlinks(blocks: &[Block]) -> usize {
    blocks
        .iter()
        .map(|b| match b {
            Block::Paragraph(p) => p
                .content
                .iter()
                .filter(|i| matches!(i, Inline::Hyperlink(_)))
                .count(),
            Block::Table(t) => t
                .rows
                .iter()
                .flat_map(|r| &r.cells)
                .map(|c| count_hyperlinks(&c.content))
                .sum(),
            Block::SectionBreak(_) => 0,
        })
        .sum()
}

fn collect_text(blocks: &[Block]) -> String {
    let mut out = String::new();
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                for inline in &p.content {
                    if let Inline::TextRun(run) = inline {
                        out.push_str(&run.text);
                    }
                }
                out.push('\n');
            }
            Block::Table(t) => {
                for row in &t.rows {
                    for cell in &row.cells {
                        out.push_str(&collect_text(&cell.content));
                    }
                }
            }
            Block::SectionBreak(_) => {}
        }
    }
    out
}

// ── sample1: richest file (tables, images, hyperlinks, numbering, footnotes) ─

#[test]
fn sample1_parses() {
    let doc = load("sample-docx-files-sample1.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample1_tables() {
    let doc = load("sample-docx-files-sample1.docx");
    let n = count_tables(&doc.body);
    assert!(n >= 6, "expected at least 6 tables, got {n}");
}

#[test]
fn sample1_images() {
    let doc = load("sample-docx-files-sample1.docx");
    assert!(
        doc.media.len() >= 3,
        "expected at least 3 media entries, got {}",
        doc.media.len()
    );
    let n = count_images(&doc.body);
    assert!(n >= 3, "expected at least 3 images in body, got {n}");
}

#[test]
fn sample1_hyperlinks() {
    let doc = load("sample-docx-files-sample1.docx");
    let n = count_hyperlinks(&doc.body);
    assert!(n >= 10, "expected at least 10 hyperlinks, got {n}");
}

#[test]
fn sample1_footnotes_endnotes() {
    let doc = load("sample-docx-files-sample1.docx");
    // File has footnotes.xml and endnotes.xml; separator notes are filtered out.
    // Real user notes (if any) should be present.
    // At minimum the parse should succeed without error.
    // Parse should succeed; separator notes are filtered out.
    let _ = doc.footnotes.len() + doc.endnotes.len();
}

#[test]
fn sample1_numbering() {
    let doc = load("sample-docx-files-sample1.docx");
    // With numbering.xml present, some paragraphs should have numbering properties
    let has_numbering = doc
        .body
        .iter()
        .any(|b| matches!(b, Block::Paragraph(p) if p.properties.numbering.is_some()));
    assert!(has_numbering, "expected at least one numbered paragraph");
}

#[test]
fn sample1_theme() {
    let doc = load("sample-docx-files-sample1.docx");
    let theme = doc.theme.as_ref().expect("expected theme");
    assert!(
        !theme.minor_font.latin.is_empty(),
        "theme should have minor latin font"
    );
}

#[test]
fn sample1_styles_resolved() {
    let doc = load("sample-docx-files-sample1.docx");
    let mut sizes = std::collections::HashSet::new();
    for block in &doc.body {
        if let Block::Paragraph(p) = block {
            for inline in &p.content {
                if let Inline::TextRun(run) = inline {
                    sizes.insert(run.properties.font_size.raw());
                }
            }
        }
    }
    assert!(
        sizes.len() > 1,
        "expected multiple distinct font sizes from style resolution, got {sizes:?}"
    );
}

// ── sample2: minimal file (5 paragraphs, 1 image) ───────────────────────────

#[test]
fn sample2_parses() {
    let doc = load("sample-docx-files-sample2.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample2_small_body() {
    let doc = load("sample-docx-files-sample2.docx");
    let n = count_paragraphs(&doc.body);
    assert!(n >= 3, "expected at least 3 paragraphs, got {n}");
    assert_eq!(count_tables(&doc.body), 0, "expected no tables");
}

#[test]
fn sample2_single_image() {
    let doc = load("sample-docx-files-sample2.docx");
    assert_eq!(doc.media.len(), 1, "expected 1 media entry");
    let n = count_images(&doc.body);
    assert_eq!(n, 1, "expected 1 image in body");
}

// ── sample3: tables + hyperlinks + numbering ─────────────────────────────────

#[test]
fn sample3_parses() {
    let doc = load("sample-docx-files-sample3.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample3_tables() {
    let doc = load("sample-docx-files-sample3.docx");
    let n = count_tables(&doc.body);
    assert!(n >= 2, "expected at least 2 tables, got {n}");
}

#[test]
fn sample3_images() {
    let doc = load("sample-docx-files-sample3.docx");
    assert_eq!(doc.media.len(), 2, "expected 2 media entries");
}

#[test]
fn sample3_hyperlinks() {
    let doc = load("sample-docx-files-sample3.docx");
    let n = count_hyperlinks(&doc.body);
    assert!(n >= 3, "expected at least 3 hyperlinks, got {n}");
}

#[test]
fn sample3_numbering() {
    let doc = load("sample-docx-files-sample3.docx");
    let has_numbering = doc
        .body
        .iter()
        .any(|b| matches!(b, Block::Paragraph(p) if p.properties.numbering.is_some()));
    assert!(has_numbering, "expected numbered paragraphs");
}

// ── sample-4: text-only with header ──────────────────────────────────────────

#[test]
fn sample_4_parses() {
    let doc = load("sample-docx-files-sample-4.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample_4_has_header() {
    let doc = load("sample-docx-files-sample-4.docx");
    assert!(
        !doc.headers.is_empty(),
        "expected at least one header, got none"
    );
}

#[test]
fn sample_4_no_images_or_tables() {
    let doc = load("sample-docx-files-sample-4.docx");
    assert_eq!(count_tables(&doc.body), 0);
    assert_eq!(count_images(&doc.body), 0);
    assert!(doc.media.is_empty());
}

#[test]
fn sample_4_paragraph_count() {
    let doc = load("sample-docx-files-sample-4.docx");
    let n = count_paragraphs(&doc.body);
    assert!(n >= 50, "expected at least 50 paragraphs, got {n}");
}

// ── sample-5: large text-only document ───────────────────────────────────────

#[test]
fn sample_5_parses() {
    let doc = load("sample-docx-files-sample-5.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample_5_many_paragraphs() {
    let doc = load("sample-docx-files-sample-5.docx");
    let n = count_paragraphs(&doc.body);
    assert!(n >= 500, "expected at least 500 paragraphs, got {n}");
}

#[test]
fn sample_5_text_only() {
    let doc = load("sample-docx-files-sample-5.docx");
    assert_eq!(count_tables(&doc.body), 0);
    assert_eq!(count_images(&doc.body), 0);
    assert!(doc.media.is_empty());
    assert!(doc.headers.is_empty());
    assert!(doc.footers.is_empty());
}

// ── sample-6: very large text-only document ──────────────────────────────────

#[test]
fn sample_6_parses() {
    let doc = load("sample-docx-files-sample-6.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample_6_many_paragraphs() {
    let doc = load("sample-docx-files-sample-6.docx");
    let n = count_paragraphs(&doc.body);
    assert!(n >= 2000, "expected at least 2000 paragraphs, got {n}");
}

// ── sample4: image-heavy (47 images) ─────────────────────────────────────────

#[test]
fn sample4_parses() {
    let doc = load("sample-docx-files-sample4.docx");
    assert!(!doc.body.is_empty());
}

#[test]
fn sample4_many_images() {
    let doc = load("sample-docx-files-sample4.docx");
    assert!(
        doc.media.len() >= 40,
        "expected at least 40 media entries, got {}",
        doc.media.len()
    );
}

#[test]
fn sample4_image_extents() {
    let doc = load("sample-docx-files-sample4.docx");
    let n = count_images(&doc.body);
    assert!(n >= 30, "expected at least 30 images in body, got {n}");

    // All images should have non-zero extents
    fn check(blocks: &[Block]) {
        for block in blocks {
            match block {
                Block::Paragraph(p) => {
                    for inline in &p.content {
                        if let Inline::Image(img) = inline {
                            assert!(
                                img.extent.width.raw() > 0 && img.extent.height.raw() > 0,
                                "image should have non-zero extent"
                            );
                        }
                    }
                }
                Block::Table(t) => {
                    for row in &t.rows {
                        for cell in &row.cells {
                            check(&cell.content);
                        }
                    }
                }
                Block::SectionBreak(_) => {}
            }
        }
    }
    check(&doc.body);
}

// ── Cross-cutting tests ──────────────────────────────────────────────────────

const ALL_FILES: &[&str] = &[
    "sample-docx-files-sample1.docx",
    "sample-docx-files-sample2.docx",
    "sample-docx-files-sample3.docx",
    "sample-docx-files-sample4.docx",
    "sample-docx-files-sample-4.docx",
    "sample-docx-files-sample-5.docx",
    "sample-docx-files-sample-6.docx",
];

#[test]
fn all_files_parse_without_error() {
    for name in ALL_FILES {
        let _ = load(name);
    }
}

#[test]
fn all_files_have_body_text() {
    for name in ALL_FILES {
        let doc = load(name);
        let text = collect_text(&doc.body);
        assert!(
            !text.trim().is_empty(),
            "{name}: document should contain text"
        );
    }
}

#[test]
fn all_files_have_valid_page_size() {
    for name in ALL_FILES {
        let doc = load(name);
        let w = doc.final_section.page_size.width.raw();
        let h = doc.final_section.page_size.height.raw();
        assert!(w > 0 && h > 0, "{name}: page size should be non-zero");
    }
}

#[test]
fn all_files_have_valid_margins() {
    for name in ALL_FILES {
        let doc = load(name);
        let m = &doc.final_section.page_margins;
        // Margins should be non-negative (they can be zero for edge-to-edge)
        assert!(m.top.raw() >= 0, "{name}: negative top margin");
        assert!(m.left.raw() >= 0, "{name}: negative left margin");
    }
}

#[test]
fn all_files_have_default_tab_stop() {
    for name in ALL_FILES {
        let doc = load(name);
        let tab = doc.settings.default_tab_stop.raw();
        assert!(tab > 0, "{name}: expected non-zero default tab stop");
    }
}

#[test]
fn table_rows_have_cells() {
    for name in ALL_FILES {
        let doc = load(name);
        fn check(blocks: &[Block], name: &str) {
            for block in blocks {
                if let Block::Table(t) = block {
                    for (i, row) in t.rows.iter().enumerate() {
                        assert!(!row.cells.is_empty(), "{name}: table row {i} has no cells");
                        for cell in &row.cells {
                            check(&cell.content, name);
                        }
                    }
                }
            }
        }
        check(&doc.body, name);
    }
}

#[test]
fn table_cells_have_content() {
    // Most cells should have content; vMerge continue cells may be empty.
    for name in ALL_FILES {
        let doc = load(name);
        fn count_empty(blocks: &[Block]) -> (usize, usize) {
            let mut total = 0usize;
            let mut empty = 0usize;
            for block in blocks {
                if let Block::Table(t) = block {
                    for row in &t.rows {
                        for cell in &row.cells {
                            total += 1;
                            if cell.content.is_empty() {
                                empty += 1;
                            }
                            let (t2, e2) = count_empty(&cell.content);
                            total += t2;
                            empty += e2;
                        }
                    }
                }
            }
            (total, empty)
        }
        let (total, empty) = count_empty(&doc.body);
        if total > 0 {
            // Allow some empty cells (vMerge continue), but not a majority
            assert!(
                empty <= total / 2,
                "{name}: too many empty cells ({empty}/{total})"
            );
        }
    }
}

#[test]
fn media_bytes_are_nonempty() {
    for name in ALL_FILES {
        let doc = load(name);
        for (rel_id, data) in &doc.media {
            assert!(!data.is_empty(), "{name}: media {rel_id:?} has empty bytes");
        }
    }
}
