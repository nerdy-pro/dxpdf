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

fn parse_docx(filename: &str) -> dxpdf::model::model::Document {
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
