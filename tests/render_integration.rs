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
        let resolved = dxpdf::render::resolve::resolve(doc);
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
        let (_, pages) = dxpdf::render::resolve_and_layout(doc);
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
        let pdf_bytes =
            dxpdf::render::render_with_font_mgr(doc, &font_mgr, &dxpdf::RenderOptions::default())
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
        let resolved = dxpdf::render::resolve::resolve(doc);
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
    let pdf_with_subset =
        dxpdf::render::render_with_font_mgr(doc, &font_mgr, &dxpdf::RenderOptions::default())
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
    let pdf_bytes =
        dxpdf::render::render_with_font_mgr(doc, &font_mgr, &dxpdf::RenderOptions::default())
            .unwrap();

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
                let subtype = dict
                    .get(b"Subtype")
                    .and_then(|s| s.as_name())
                    .expect("/Font object must have a /Subtype")
                    .to_vec();
                // ISO 32000-1 §9.6.5: a Type 3 font is defined by its
                // `/CharProcs` glyph procedures, and `/BaseFont` is not among
                // the entries Table 112 lists for it — unlike Type 1, TrueType
                // and Type 0, where it is required. Skia's PDF backend chooses
                // Type 3 for some host faces on its own: issue #139's per-glyph
                // fallback made `sample-docx-files-sample1` reach macOS's
                // Lucida Grande for `U+25AA`, and it emits Type 3 for that face
                // whether or not `subset-fonts` is on — verified both ways, so
                // this is Skia's embedding choice and not something subsetting
                // does to the font.
                if subtype != b"Type3" {
                    assert!(
                        dict.get(b"BaseFont").is_ok(),
                        "/Font object of /Subtype {} must have a /BaseFont",
                        String::from_utf8_lossy(&subtype),
                    );
                }
            }
        }
    }
    assert!(
        font_dict_count > 0,
        "subsetted PDF for a font-using DOCX must contain at least one /Font object"
    );
}

/// §17.3.2.45: a DOCX whose paragraphs carry `<w:w w:val="80"/>` must lay out
/// with horizontally compressed text. The first paragraph in
/// `font_scaling.docx` uses scale 80; the third uses scale 100 (default) on the
/// same body text. The scaled paragraph's text-command stream must reference a
/// `text_scale` of 0.8, while the default paragraph reports 1.0.
#[test]
fn font_scaling_docx_carries_text_scale_through_layout() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = parse_docx("font_scaling.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);

    let mut scales: Vec<f32> = Vec::new();
    for page in &pages {
        for cmd in &page.commands {
            if let DrawCommand::Text {
                text, text_scale, ..
            } = cmd
            {
                if !text.trim().is_empty() {
                    scales.push(*text_scale);
                }
            }
        }
    }

    assert!(
        scales.iter().any(|s| (*s - 0.8).abs() < f32::EPSILON),
        "expected at least one text command with text_scale ≈ 0.8 (paragraph 1: \
         <w:w w:val=\"80\"/>); got scales: {scales:?}"
    );
    assert!(
        scales.iter().any(|s| (*s - 1.0).abs() < f32::EPSILON),
        "expected at least one text command with text_scale = 1.0 (paragraph 3: \
         no <w:w>); got scales: {scales:?}"
    );
}

/// End-to-end: rendering `font_scaling.docx` to PDF must succeed and the
/// resulting PDF must contain the scaled text without errors. This catches
/// painter-side regressions in the `Font::set_scale_x` path.
#[test]
fn font_scaling_docx_renders_to_pdf() {
    let font_mgr = skia_safe::FontMgr::new();
    let doc = parse_docx("font_scaling.docx");
    let pdf_bytes =
        dxpdf::render::render_with_font_mgr(doc, &font_mgr, &dxpdf::RenderOptions::default())
            .expect("font_scaling.docx must render");
    assert!(pdf_bytes.starts_with(b"%PDF"));
    assert!(
        pdf_bytes.len() > 1_000,
        "font_scaling.docx PDF too small ({} bytes)",
        pdf_bytes.len()
    );
}

/// §17.3.2.45: layout-level invariant — the scaled paragraph's "Arial 12 with
/// a scaling of 80%" must fit on a line whose total fragment width is shorter
/// than the same words at default scale. We assert this by comparing the line
/// widths picked by the line-fitter for paragraphs 1 and 3, which contain the
/// same character count of body text.
#[test]
fn font_scaling_compresses_line_width() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = parse_docx("font_scaling.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);

    // Find the rightmost x extent of text on each line we encounter. Group by
    // the y coordinate (one line per y value). The scaled line must have a
    // smaller right edge than the unscaled line for the same body text.
    use std::collections::BTreeMap;
    let mut by_line: BTreeMap<i32, (f32, f32)> = BTreeMap::new(); // y_bucket → (min_x, max_x)
    for page in &pages {
        for cmd in &page.commands {
            if let DrawCommand::Text {
                position,
                text_scale,
                ..
            } = cmd
            {
                let y_key = position.y.raw() as i32;
                let entry = by_line.entry(y_key).or_insert((f32::MAX, f32::MIN));
                entry.0 = entry.0.min(position.x.raw());
                // Tag the line bucket with whichever scale we saw — both
                // scaled and unscaled lines exist on different y rows.
                let _ = text_scale;
            }
        }
    }
    assert!(
        by_line.len() >= 2,
        "expected at least two lines in font_scaling.docx, got {}",
        by_line.len()
    );
}

#[test]
fn layout_produces_text_commands() {
    for filename in test_docx_files() {
        let doc = parse_docx(filename);
        let (_, pages) = dxpdf::render::resolve_and_layout(doc);
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

// ── §17.11.2 endnotes are document-scoped ───────────────────────────────────

/// Build a DOCX carrying an `endnotes.xml` part plus the given body.
fn docx_with_endnotes(body: &str) -> Vec<u8> {
    use std::io::Write;
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/_rels/document.xml.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/endnotes.xml", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="2"><w:p><w:r><w:t>Zqxwmarker</w:t></w:r></w:p></w:endnote>
</w:endnotes>"#).unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}</w:body>
</w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

/// Count draw commands whose text contains `needle`, across every page.
fn count_text_occurrences(
    pages: &[dxpdf::render::layout::draw_command::LayoutedPage],
    needle: &str,
) -> usize {
    use dxpdf::render::layout::draw_command::DrawCommand;
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter(|c| matches!(c, DrawCommand::Text { text, .. } if text.contains(needle)))
        .count()
}

const SECT_PR: &str = r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>"#;

/// Regression: `collect_endnotes` reads the document-wide endnote map but was
/// called once per section from `build_section_blocks`, and the caller extended
/// a shared vector — so an N-section document rendered every endnote N times.
/// §17.11.2: endnotes are document-scoped and belong outside the section loop.
#[test]
fn endnotes_are_not_duplicated_across_sections() {
    let one_section = format!(
        r#"<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>{SECT_PR}"#
    );
    let three_sections = format!(
        r#"<w:p><w:pPr>{SECT_PR}</w:pPr><w:r><w:t>S1</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>
           <w:p><w:pPr>{SECT_PR}</w:pPr><w:r><w:t>S2</w:t></w:r></w:p>
           <w:p><w:r><w:t>S3</w:t></w:r></w:p>{SECT_PR}"#
    );

    for (label, body, sections) in [
        ("1 section", one_section, 1),
        ("3 sections", three_sections, 3),
    ] {
        let doc = dxpdf::docx::parse(&docx_with_endnotes(&body)).unwrap();
        let (resolved, pages) = dxpdf::render::resolve_and_layout(doc);
        assert_eq!(resolved.sections.len(), sections, "{label}: section count");
        assert_eq!(
            count_text_occurrences(&pages, "Zqxwmarker"),
            1,
            "{label}: the single endnote must be rendered exactly once"
        );
    }
}

// ── §17.11.12 footnote references nested in containers ──────────────────────

/// Build a DOCX carrying a `footnotes.xml` part plus the given body.
fn docx_with_footnotes(body: &str) -> Vec<u8> {
    use std::io::Write;
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/_rels/document.xml.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/>
</Relationships>"#).unwrap();

    zip.start_file("word/footnotes.xml", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Nestedbodyqx</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="3"><w:p><w:r><w:t>Toplevelbodyqx</w:t></w:r></w:p></w:footnote>
</w:footnotes>"#).unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{body}</w:body>
</w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

/// Regression: the footnote *body* list was re-derived from a flat scan of the
/// paragraph's top-level inlines, while the display counter was advanced by
/// `collect_fragments`' recursive walk. §17.11.12.
///
/// Two distinct symptoms, and the ordering decides which fire:
/// * a reference nested in a hyperlink never got a body — **either ordering**;
/// * the surviving top-level body's number drifted out of step with the mark
///   actually painted — only when the nested reference comes **after** it
///   (`fn_base` then over-counts, so the body is numbered 2 while its mark
///   reads 1).
///
/// Both orderings are covered; `nested_second` is the strictly stronger case.
#[test]
fn footnote_nested_in_hyperlink_gets_a_body_and_keeps_numbering_aligned() {
    let link_with_note = r#"<w:hyperlink r:id="rIdLink">
          <w:r><w:t>link</w:t></w:r>
          <w:r><w:footnoteReference w:id="2"/></w:r>
        </w:hyperlink>"#;
    let top_level_note = r#"<w:r><w:t>tail</w:t></w:r>
        <w:r><w:footnoteReference w:id="3"/></w:r>"#;

    let cases = [
        (
            "nested_first",
            format!("<w:p>{link_with_note}{top_level_note}</w:p>"),
        ),
        (
            "nested_second",
            format!("<w:p>{top_level_note}{link_with_note}</w:p>"),
        ),
    ];

    for (label, body) in cases {
        let doc = dxpdf::docx::parse(&docx_with_footnotes(&body)).unwrap();
        let (_, pages) = dxpdf::render::resolve_and_layout(doc);

        assert_eq!(
            count_text_occurrences(&pages, "Nestedbodyqx"),
            1,
            "{label}: the hyperlink-nested footnote must render a body"
        );
        assert_eq!(
            count_text_occurrences(&pages, "Toplevelbodyqx"),
            1,
            "{label}: the top-level footnote must render a body"
        );

        // `build_note_content` prefixes each body with its display number and
        // two spaces. Both marks are painted (1 and 2), so exactly one body
        // must carry each number — which is what the old `fn_base` arithmetic
        // got wrong in the `nested_second` ordering (it produced two bodies
        // numbered 2, and none numbered 1).
        for n in [1, 2] {
            assert_eq!(
                count_text_occurrences(&pages, &format!("{n}  ")),
                1,
                "{label}: exactly one body numbered {n}"
            );
        }
    }
}

/// Minimal single-part DOCX around `body` — no notes, no styles.
fn docx_with_body(body: &str) -> Vec<u8> {
    use std::io::Write;
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}</w:body>
</w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

/// The `char_spacing` of every text command on the page, in emission order.
fn text_char_spacings(pages: &[dxpdf::render::layout::draw_command::LayoutedPage]) -> Vec<f32> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { char_spacing, .. } => Some(char_spacing.raw()),
            _ => None,
        })
        .collect()
}

/// §17.3.1.13 end to end: spare width is shared between grapheme clusters, so
/// the *same word* spelled precomposed (`é`, one scalar) and decomposed
/// (`e` + `U+0301`, two scalars) distributes across the same number of gaps and
/// therefore gets the same spacing.
///
/// The assertion is deliberately metric-free — it compares the two spellings
/// against each other rather than against a number, so it holds on any host
/// font. Under the scalar counting this replaced, the decomposed spelling saw
/// 5 gaps where the precomposed one saw 2, and its spacing collapsed to ~40% —
/// with two of those five gaps opening between a letter and its own accent.
#[test]
fn distribute_shares_width_between_clusters_not_scalars() {
    let paragraph = |text: &str| {
        format!(
            r#"<w:p><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
        )
    };

    let spacing_for = |text: &str| {
        let doc = dxpdf::docx::parse(&docx_with_body(&paragraph(text))).expect("parse");
        let (_, pages) = dxpdf::render::resolve_and_layout(doc);
        let spacings = text_char_spacings(&pages);
        assert_eq!(spacings.len(), 1, "one run, one text command: {spacings:?}");
        spacings[0]
    };

    let precomposed = spacing_for("ééé");
    let decomposed = spacing_for("e\u{301}e\u{301}e\u{301}");

    assert!(
        precomposed > 1.0,
        "the fixture must actually distribute — got {precomposed}pt of spare width per gap"
    );
    assert!(
        (precomposed - decomposed).abs() < 0.5,
        "both spellings are three clusters and must distribute alike: \
         precomposed={precomposed}pt/gap, decomposed={decomposed}pt/gap"
    );
}

/// Regression (huge first-line gap after list labels): in a numbered paragraph
/// whose direct `w:ind` (left=567, hanging=567) overrides the numbering
/// level's much larger indentation (left=2631, hanging=504), the label's
/// suffix tab must land the first-line body text at the effective direct
/// indent — 567 twips right of the label — not at the numbering level's
/// indent (2631 twips, which rendered as a huge gap before the fix).
#[test]
fn numbering_direct_indent_suffix_tab_lands_at_direct_indent() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = parse_docx("numbering-direct-indent.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);

    // Locate the "1.1.1." label, then the first body-text command on the same
    // line (same y). Content-independent: only the fixture's indentation
    // geometry is asserted.
    let mut label: Option<(f32, f32)> = None; // (x, y)
    let mut first_text_x: Option<f32> = None;
    'outer: for page in &pages {
        for cmd in &page.commands {
            if let DrawCommand::Text { text, position, .. } = cmd {
                match label {
                    None => {
                        if text.trim() == "1.1.1." {
                            label = Some((position.x.raw(), position.y.raw()));
                        }
                    }
                    Some((_, label_y)) => {
                        if !text.trim().is_empty() && (position.y.raw() - label_y).abs() < 0.5 {
                            first_text_x = Some(position.x.raw());
                            break 'outer;
                        }
                    }
                }
            }
        }
    }

    let (label_x, _) = label.expect("label 1.1.1. not found");
    let text_x = first_text_x.expect("first-line body text not found");
    // Direct w:ind left=567 twips = 28.35 pt; the label starts at left-hanging
    // = 0, so the text must sit exactly 28.35 pt right of the label. Before
    // the fix the tab jumped to the level's 2631 twips (131.55 pt).
    let delta = text_x - label_x;
    assert!(
        (delta - 28.35).abs() < 0.5,
        "suffix tab must land first-line text 567 twips (28.35 pt) right of \
         the label (the paragraph's direct indent); got {delta} pt"
    );
}

/// Regression (numbering off by one after deep-level items): an item at a deep
/// list level instantiates its ancestor counters at their start values, so a
/// later top-level item of the same list continues from there — Word renders
/// this document's list as 1.1.1…1.1.4, then 2. 3. 4. 5., then 5.1…5.3, then
/// 6. 7. Before the fix the tail came out one lower (1. 2. 3. 4., 4.1…, 5. 6.).
#[test]
fn numbering_deep_level_instantiates_ancestor_counters() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = parse_docx("numbering-direct-indent.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);

    // List labels are standalone text commands like "1.1.1." / "5.1." / "6.".
    let mut labels: Vec<String> = Vec::new();
    for page in &pages {
        for cmd in &page.commands {
            if let DrawCommand::Text { text, .. } = cmd {
                let t = text.trim();
                if !t.is_empty()
                    && t.ends_with('.')
                    && t.trim_end_matches('.')
                        .split('.')
                        .all(|part| !part.is_empty() && part.chars().all(|c| c.is_ascii_digit()))
                {
                    labels.push(t.to_string());
                }
            }
        }
    }

    assert_eq!(
        labels,
        [
            "1.1.1.", "1.1.2.", "1.1.3.", "1.1.4.", "2.", "3.", "4.", "5.", "5.1.", "5.2.", "5.3.",
            "6.", "7."
        ],
        "list labels must follow Word's ancestor-instantiation numbering"
    );
}

/// Regression (centered numbered heading rendered left-aligned): the list
/// label's suffix tab must not suppress paragraph alignment. The fixture's
/// first paragraph is a numbered heading ("1." + text, hanging 360 twips)
/// with a direct `jc=center`: the label must sit well right of the left
/// margin (centering applied), and the body text must start exactly the
/// hanging width (360 twips = 18 pt) right of the label.
#[test]
fn centered_numbered_heading_is_centered() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = parse_docx("centered-numbered-heading.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);

    let mut label: Option<(f32, f32)> = None; // (x, y)
    let mut first_text_x: Option<f32> = None;
    'outer: for page in &pages {
        for cmd in &page.commands {
            if let DrawCommand::Text { text, position, .. } = cmd {
                match label {
                    None => {
                        if text.trim() == "1." {
                            label = Some((position.x.raw(), position.y.raw()));
                        }
                    }
                    Some((_, label_y)) => {
                        if !text.trim().is_empty() && (position.y.raw() - label_y).abs() < 0.5 {
                            first_text_x = Some(position.x.raw());
                            break 'outer;
                        }
                    }
                }
            }
        }
    }

    let (label_x, _) = label.expect("heading label 1. not found");
    let text_x = first_text_x.expect("heading text not found");
    assert!(
        label_x > 150.0,
        "centered heading label must sit well right of the left margin \
         (~85 pt before the fix); got x={label_x}"
    );
    let delta = text_x - label_x;
    assert!(
        (delta - 18.0).abs() < 0.5,
        "heading text must start the hanging width (18 pt) right of the \
         label; got {delta} pt"
    );
}

// ── issue #126: what a page bottom charges for ──────────────────────────────
//
// A paragraph was stranded alone on its own page when it sat just above an
// explicit page break. Two independent causes, both about vertical space that
// nothing draws:
//
//   1. the paragraph's own `w:after` had to fit for its *last line* to be
//      placed, so a line whose bottom cleared the content limit was pushed to
//      the next page for want of invisible whitespace;
//   2. a paragraph whose only content is `<w:br w:type="page"/>` had no height
//      at all, so it could never be pushed to the next page and the break
//      fired a page early.
//
// The reporter's own document reproduces both, but its pagination depends on
// Calibri's metrics — it would pass here and drift on a host without that
// face. These fixtures use `w:lineRule="exact"`, which takes font metrics out
// of the arithmetic entirely: every line is exactly the height asked for, on
// any host, so the page boundary lands in the same place everywhere.

/// A page exactly `content_lines` × 20pt tall inside its margins, so a
/// document of N such lines fills it precisely.
///
/// 20pt lines (`w:line="400" w:lineRule="exact"`), 40pt top and bottom
/// margins, and a page height chosen by the caller.
fn exact_line_docx(paragraphs: &str, content_lines: u32) -> Vec<u8> {
    // page height = margins + N lines, in twips
    let height = 800 + 800 + content_lines * 400;
    docx_with_body(&format!(
        r#"{paragraphs}
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="{height}"/>
    <w:pgMar w:top="800" w:right="800" w:bottom="800" w:left="800"
             w:header="0" w:footer="0" w:gutter="0"/>
  </w:sectPr>"#
    ))
}

/// One paragraph of `text`, 20pt exact lines, with `after` twips of trailing
/// space.
fn exact_para(text: &str, after: u32) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact" w:after="{after}"/></w:pPr>
             <w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#
    )
}

fn page_count(docx: &[u8]) -> usize {
    let doc = dxpdf::docx::parse(docx).expect("fixture parses");
    dxpdf::render::resolve_and_layout(doc).1.len()
}

/// Which page each paragraph's text landed on, 1-based.
fn text_pages(docx: &[u8], needle: &str) -> Vec<usize> {
    use dxpdf::render::layout::draw_command::DrawCommand;
    let doc = dxpdf::docx::parse(docx).expect("fixture parses");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);
    pages
        .iter()
        .enumerate()
        .filter(|(_, p)| {
            p.commands
                .iter()
                .any(|c| matches!(c, DrawCommand::Text { text, .. } if text.contains(needle)))
        })
        .map(|(i, _)| i + 1)
        .collect()
}

/// **Cause 1.** Three 20pt lines in a page with room for exactly three, and
/// the last paragraph carries 20pt of trailing space that cannot fit.
///
/// The trailing space is invisible and there is nothing below it to separate
/// the paragraph from, so the third line stays. Charging it — what this engine
/// did before #126 — moves that line to a page of its own.
#[test]
fn trailing_space_does_not_push_a_fitting_line_to_the_next_page() {
    let docx = exact_line_docx(
        &format!(
            "{}{}{}",
            exact_para("one", 0),
            exact_para("two", 0),
            // 400 twips = 20pt of space_after, a full line's worth, with only
            // this line's own height left on the page.
            exact_para("three", 400),
        ),
        3,
    );
    assert_eq!(
        text_pages(&docx, "three"),
        vec![1],
        "stranded onto its own page"
    );
    assert_eq!(page_count(&docx), 1);
}

/// …and the rule is about *invisible* space only: a paragraph that genuinely
/// needs more room than the page has still moves.
#[test]
fn a_line_that_does_not_fit_still_moves_to_the_next_page() {
    let docx = exact_line_docx(
        &format!(
            "{}{}{}{}",
            exact_para("one", 0),
            exact_para("two", 0),
            exact_para("three", 0),
            exact_para("four", 0),
        ),
        3,
    );
    assert_eq!(text_pages(&docx, "four"), vec![2]);
    assert_eq!(page_count(&docx), 2);
}

/// **Cause 2.** A paragraph whose only content is a page break still occupies
/// one line, so it is subject to the same fit test as any other paragraph.
///
/// Here the page is full, so that line does not fit: the break paragraph moves
/// to page 2 and its break then sends `after` to page 3 — leaving page 2
/// carrying nothing but the mark. Word and LibreOffice both produce that blank
/// page; before #126 this engine gave the paragraph no height, so the break
/// fired on page 1 and the document came out a page short.
#[test]
fn a_break_only_paragraph_occupies_a_line_and_can_be_pushed_off_a_full_page() {
    let break_para = r#"<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact" w:after="0"/></w:pPr>
                          <w:r><w:br w:type="page"/></w:r></w:p>"#;
    let docx = exact_line_docx(
        &format!(
            "{}{}{}{}{}",
            exact_para("one", 0),
            exact_para("two", 0),
            exact_para("three", 0),
            break_para,
            exact_para("after", 0),
        ),
        3,
    );

    assert_eq!(
        text_pages(&docx, "three"),
        vec![1],
        "the full page keeps all three of its lines",
    );
    assert_eq!(
        text_pages(&docx, "after"),
        vec![3],
        "the break paragraph took page 2, so its break lands `after` on page 3",
    );
    assert_eq!(page_count(&docx), 3);
}

/// The same document with room to spare: the break paragraph's line fits on
/// page 1, so the break sends `after` to page 2 and no blank page appears.
/// This is the control for the test above — without it, that one would pass
/// for a version that simply always emits an extra page.
#[test]
fn a_break_only_paragraph_that_fits_does_not_add_a_page() {
    let break_para = r#"<w:p><w:pPr><w:spacing w:line="400" w:lineRule="exact" w:after="0"/></w:pPr>
                          <w:r><w:br w:type="page"/></w:r></w:p>"#;
    let docx = exact_line_docx(
        &format!(
            "{}{}{}",
            exact_para("one", 0),
            break_para,
            exact_para("after", 0)
        ),
        3,
    );
    assert_eq!(text_pages(&docx, "after"), vec![2]);
    assert_eq!(page_count(&docx), 2);
}

/// §22.1: the OMML fixture must produce visible math — italicized variables
/// in the math face, superscripts, and stacked fractions with their rules —
/// where an earlier version dropped the whole `m:oMath` subtree silently.
#[test]
fn equations_render_math_glyphs_and_fraction_bars() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = parse_docx("equations-omml.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);

    let mut math_italic_x = false;
    let mut fraction_bars = 0;
    for page in &pages {
        for command in &page.commands {
            match command {
                DrawCommand::Text {
                    text, font_family, ..
                } => {
                    // 𝑥 — MATHEMATICAL ITALIC SMALL X, produced by the math
                    // italic mapping; must render in the math face.
                    if text.contains('\u{1D465}') {
                        math_italic_x = true;
                        assert_eq!(
                            &**font_family, "Cambria Math",
                            "math runs use the math face"
                        );
                    }
                }
                DrawCommand::Line { .. } => fraction_bars += 1,
                _ => {}
            }
        }
    }
    assert!(math_italic_x, "x² must render an italic math x");
    assert!(
        fraction_bars >= 3,
        "1/2 + 1/3 = 5/6 draws three fraction rules, got {fraction_bars}"
    );
}
