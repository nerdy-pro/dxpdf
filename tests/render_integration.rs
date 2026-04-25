//! Integration tests — parse real DOCX files, render with the renderer.

use std::path::Path;

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

fn test_docx_files() -> Vec<&'static str> {
    vec![
        "sample-docx-files-sample1.docx",
        "sample-docx-files-sample2.docx",
        "sample-docx-files-sample3.docx",
        "sample-docx-files-sample4.docx",
        "sample-docx-files-sample-4.docx",
        "sample-docx-files-sample-5.docx",
        "sample-docx-files-sample-6.docx",
    ]
}

fn parse_docx(filename: &str) -> dxpdf::model::Document {
    let path = Path::new(TEST_DIR).join(filename);
    let bytes = std::fs::read(&path).unwrap_or_else(|e| {
        panic!("Failed to read {}: {e}", path.display());
    });
    dxpdf::docx::parse(&bytes).unwrap_or_else(|e| {
        panic!("Failed to parse {}: {e}", path.display());
    })
}

#[test]
fn all_files_resolve_without_error() {
    for filename in test_docx_files() {
        let doc = parse_docx(filename);
        let resolved = dxpdf::render::resolve::resolve(&doc);
        assert!(
            !resolved.sections.is_empty(),
            "{filename}: should have at least one section"
        );
    }
}

#[test]
fn all_files_layout_without_error() {
    for filename in test_docx_files() {
        let doc = parse_docx(filename);
        let (_, pages) = dxpdf::render::resolve_and_layout(&doc);
        assert!(
            !pages.is_empty(),
            "{filename}: should produce at least one page"
        );
    }
}

#[test]
fn all_files_render_to_pdf() {
    let font_mgr = skia_safe::FontMgr::new();
    for filename in test_docx_files() {
        let doc = parse_docx(filename);
        let pdf_bytes = dxpdf::render::render_with_font_mgr(&doc, &font_mgr)
            .unwrap_or_else(|e| panic!("{filename}: render failed: {e}"));
        assert!(
            pdf_bytes.len() > 100,
            "{filename}: PDF output too small ({} bytes)",
            pdf_bytes.len()
        );
        assert!(
            pdf_bytes.starts_with(b"%PDF"),
            "{filename}: output doesn't start with %PDF header"
        );
    }
}

#[test]
fn resolve_collects_fonts_from_real_docs() {
    for filename in test_docx_files() {
        let doc = parse_docx(filename);
        let resolved = dxpdf::render::resolve::resolve(&doc);
        assert!(
            !resolved.font_families.is_empty(),
            "{filename}: should have at least one font family"
        );
    }
}

/// Subsetting effectiveness — sample1 embeds six TTF fonts and used to produce
/// a ~1.7 MB PDF; with the `subset-fonts` feature on (default), output should
/// shrink dramatically. Empirically observed at the time the feature shipped:
/// 1.73 MB → 274 KB, an 84% reduction. We assert a much looser bound (≤ 50%
/// of the no-subset baseline) so cross-platform variation in available fonts
/// can't make this flake.
#[test]
#[cfg(feature = "subset-fonts")]
fn font_subsetting_shrinks_pdf_with_embedded_fonts() {
    let font_mgr = skia_safe::FontMgr::new();
    let doc = parse_docx("sample-docx-files-sample1.docx");
    assert!(
        !doc.embedded_fonts.is_empty(),
        "test precondition: sample1 must contain embedded fonts"
    );
    let pdf_with_subset = dxpdf::render::render_with_font_mgr(&doc, &font_mgr)
        .expect("subset-on render must succeed");

    // Sanity: still a valid PDF, has actual content.
    assert!(pdf_with_subset.starts_with(b"%PDF"));
    assert!(pdf_with_subset.len() > 50_000);

    // The hard threshold — subsetting must produce at most 50% of the
    // no-subset baseline. Loose enough to absorb cross-platform font
    // availability differences while still catching regressions.
    const NO_SUBSET_BASELINE: usize = 1_771_367;
    assert!(
        pdf_with_subset.len() < NO_SUBSET_BASELINE / 2,
        "subset-on output ({} bytes) must be < 50% of no-subset baseline ({}), \
         observed shrinkage: {:.1}%",
        pdf_with_subset.len(),
        NO_SUBSET_BASELINE,
        100.0 * (1.0 - pdf_with_subset.len() as f64 / NO_SUBSET_BASELINE as f64)
    );
}

/// Validate that subsetted PDFs still parse cleanly via a real PDF parser
/// (`lopdf`). Catches the broken-output regression: any malformed cross-
/// reference table, bad stream length, or invalid object would fail here.
/// This is the integration-level equivalent of the unit-test invariant
/// `subset_output_is_skia_shapeable`.
#[test]
#[cfg(feature = "subset-fonts")]
fn subsetted_pdf_is_well_formed() {
    let font_mgr = skia_safe::FontMgr::new();
    let doc = parse_docx("sample-docx-files-sample1.docx");
    let pdf_bytes = dxpdf::render::render_with_font_mgr(&doc, &font_mgr).unwrap();

    let parsed =
        lopdf::Document::load_mem(&pdf_bytes).expect("subsetted PDF must parse cleanly with lopdf");
    assert!(
        !parsed.get_pages().is_empty(),
        "subsetted PDF must report at least one page"
    );

    // Walk every Font object and assert it has the structural fields a
    // PDF reader needs (Type=Font, Subtype, BaseFont). If subsetting had
    // damaged the font dictionaries, this would fail.
    let mut font_dict_count = 0;
    for obj in parsed.objects.values() {
        if let Ok(dict) = obj.as_dict() {
            if dict
                .get(b"Type")
                .ok()
                .and_then(|t| t.as_name().ok())
                .is_some_and(|n| n == b"Font")
            {
                font_dict_count += 1;
                assert!(
                    dict.get(b"Subtype").is_ok(),
                    "/Font object must have a /Subtype"
                );
                assert!(
                    dict.get(b"BaseFont").is_ok(),
                    "/Font object must have a /BaseFont"
                );
            }
        }
    }
    assert!(
        font_dict_count > 0,
        "subsetted PDF for a font-using DOCX must contain at least one /Font object"
    );
}

#[test]
fn layout_produces_text_commands() {
    for filename in test_docx_files() {
        let doc = parse_docx(filename);
        let (_, pages) = dxpdf::render::resolve_and_layout(&doc);
        let total_text_cmds: usize = pages
            .iter()
            .map(|p| {
                p.commands
                    .iter()
                    .filter(|c| {
                        matches!(
                            c,
                            dxpdf::render::layout::draw_command::DrawCommand::Text { .. }
                        )
                    })
                    .count()
            })
            .sum();
        assert!(
            total_text_cmds > 0,
            "{filename}: should produce at least one text command"
        );
    }
}
