use std::io::Write;

/// Create a minimal DOCX (ZIP) file in memory with the given document.xml content.
fn make_docx(document_xml: &str) -> Vec<u8> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);

    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Content types
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .unwrap();

    // Relationships
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    // Main document
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

fn simple_docx(body_content: &str) -> Vec<u8> {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body_content}</w:body>
</w:document>"#
    );
    make_docx(&xml)
}

#[test]
fn convert_simple_docx_to_pdf() {
    let docx = simple_docx(r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#);
    let pdf = dxpdf::convert(&docx).unwrap();

    // PDF should start with the magic bytes
    assert!(pdf.len() > 4);
    assert_eq!(&pdf[..5], b"%PDF-");
}

#[test]
fn convert_formatted_docx_to_pdf() {
    let docx = simple_docx(
        r#"<w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:i/>
                    <w:sz w:val="36"/>
                    <w:color w:val="0000FF"/>
                    <w:rFonts w:ascii="Times New Roman"/>
                </w:rPr>
                <w:t>Formatted Title</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Normal paragraph text.</w:t></w:r>
        </w:p>"#,
    );
    let pdf = dxpdf::convert(&docx).unwrap();
    assert_eq!(&pdf[..5], b"%PDF-");
}

#[test]
fn convert_table_docx_to_pdf() {
    let docx = simple_docx(
        r#"<w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Cell A1</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Cell B1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Cell A2</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Cell B2</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>"#,
    );
    let pdf = dxpdf::convert(&docx).unwrap();
    assert_eq!(&pdf[..5], b"%PDF-");
}

#[test]
fn convert_empty_document() {
    let docx = simple_docx("");
    let pdf = dxpdf::convert(&docx).unwrap();
    assert_eq!(&pdf[..5], b"%PDF-");
}

/// Per ECMA-376 §17.4.38 (CT_Tbl), `<w:bookmarkStart>` and `<w:bookmarkEnd>`
/// may appear interleaved with `<w:tr>` elements (they're part of
/// EG_RangeMarkupElements in the choice group). The parser must not split
/// the row sequence across non-row siblings.
#[test]
fn convert_table_with_bookmarks_between_rows() {
    let docx = simple_docx(
        r#"<w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>R1</w:t></w:r></w:p></w:tc></w:tr>
            <w:bookmarkStart w:id="0" w:name="anchor"/>
            <w:tr><w:tc><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr><w:tc><w:p><w:r><w:t>R3</w:t></w:r></w:p></w:tc></w:tr>
            <w:bookmarkEnd w:id="0"/>
            <w:tr><w:tc><w:p><w:r><w:t>R4</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>"#,
    );
    let document = dxpdf::docx::parse(&docx).unwrap();
    let table = match &document.body[0] {
        dxpdf::model::Block::Table(t) => t,
        b => panic!("expected table, got {b:?}"),
    };
    assert_eq!(
        table.rows.len(),
        4,
        "all four rows must survive a bookmarked split",
    );
    let pdf = dxpdf::convert(&docx).unwrap();
    assert_eq!(&pdf[..5], b"%PDF-");
}

/// `<w:proofErr>`, `<w:permStart>`, `<w:permEnd>` are EG_RunLevelElts that
/// CT_Tbl admits between rows. Word emits these around proofreading marks
/// and document protection ranges; they have no rendered effect but must
/// not break row parsing.
#[test]
fn convert_table_with_proof_err_between_rows() {
    let docx = simple_docx(
        r#"<w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
            <w:proofErr w:type="spellStart"/>
            <w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
            <w:proofErr w:type="spellEnd"/>
        </w:tbl>"#,
    );
    let document = dxpdf::docx::parse(&docx).unwrap();
    let table = match &document.body[0] {
        dxpdf::model::Block::Table(t) => t,
        b => panic!("expected table, got {b:?}"),
    };
    assert_eq!(table.rows.len(), 2);
}

/// `<w:ins>` at table level wraps inserted rows (CT_RowTrackChange,
/// §17.13.5.16). The wrapped `<w:tr>` children must surface as ordinary
/// rows in the model — the renderer ignores the revision metadata, but
/// the rows themselves must render.
#[test]
fn convert_table_with_revision_tracked_inserted_rows() {
    let docx = simple_docx(
        r#"<w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>existing</w:t></w:r></w:p></w:tc></w:tr>
            <w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z">
                <w:tr><w:tc><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc></w:tr>
            </w:ins>
        </w:tbl>"#,
    );
    let document = dxpdf::docx::parse(&docx).unwrap();
    let table = match &document.body[0] {
        dxpdf::model::Block::Table(t) => t,
        b => panic!("expected table, got {b:?}"),
    };
    assert_eq!(
        table.rows.len(),
        2,
        "rows wrapped in <w:ins> must surface alongside ordinary rows",
    );
}

/// `<w:sdt>` at table level (CT_SdtRow, §17.5.2.30) wraps rows in a
/// content control. Rows live under `<w:sdtContent>`; they must still
/// render.
#[test]
fn convert_table_with_sdt_wrapping_rows() {
    let docx = simple_docx(
        r#"<w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:tc></w:tr>
            <w:sdt>
                <w:sdtContent>
                    <w:tr><w:tc><w:p><w:r><w:t>controlled</w:t></w:r></w:p></w:tc></w:tr>
                </w:sdtContent>
            </w:sdt>
        </w:tbl>"#,
    );
    let document = dxpdf::docx::parse(&docx).unwrap();
    let table = match &document.body[0] {
        dxpdf::model::Block::Table(t) => t,
        b => panic!("expected table, got {b:?}"),
    };
    assert_eq!(table.rows.len(), 2);
}

#[test]
fn convert_writes_to_file() {
    let docx = simple_docx(r#"<w:p><w:r><w:t>File test</w:t></w:r></w:p>"#);
    let pdf = dxpdf::convert(&docx).unwrap();

    let dir = tempfile::tempdir().unwrap();
    let out_path = dir.path().join("test.pdf");
    std::fs::write(&out_path, &pdf).unwrap();

    let read_back = std::fs::read(&out_path).unwrap();
    assert_eq!(read_back, pdf);
}

#[test]
fn parse_invalid_zip_returns_error() {
    let result = dxpdf::convert(b"not a zip file");
    assert!(result.is_err());
}

#[test]
fn parse_zip_without_document_xml_returns_error() {
    // Create a valid ZIP but without word/document.xml
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file("dummy.txt", options).unwrap();
    zip.write_all(b"hello").unwrap();
    let cursor = zip.finish().unwrap();
    let bytes = cursor.into_inner();

    let result = dxpdf::convert(&bytes);
    assert!(result.is_err());
    let err = result.unwrap_err().to_string();
    assert!(
        err.contains("missing required part"),
        "Error should mention a missing part: {err}"
    );
}

/// Regression: a whitespace-only `<w:t xml:space="preserve">` between two
/// other runs must round-trip the literal space into the parsed model.
/// quick-xml's serde Deserializer trims whitespace-only `$text` content,
/// which used to drop the separator and render `Label:Value` instead of
/// `Label: Value`. The whitespace-workaround module substitutes a Private
/// Use Area sentinel before parsing and reverses it during run conversion.
#[test]
fn whitespace_only_run_with_xml_space_preserve_roundtrips_to_space() {
    use dxpdf::model::{Block, Inline, RunElement};

    // Same shape as the failing cell in the Protokoll DIN VDE document:
    // bold label run, whitespace-only run with xml:space="preserve",
    // value run with different formatting.
    let docx = simple_docx(
        r#"<w:p>
            <w:r><w:rPr><w:b/></w:rPr><w:t>Label:</w:t></w:r>
            <w:r><w:t xml:space="preserve"> </w:t></w:r>
            <w:r><w:t>Value</w:t></w:r>
        </w:p>"#,
    );

    let document = dxpdf::docx::parse(&docx).expect("parse");
    let para = match document.body.first().expect("at least one block") {
        Block::Paragraph(p) => p,
        other => panic!("expected paragraph, got {other:?}"),
    };

    let texts: Vec<&str> = para
        .content
        .iter()
        .filter_map(|inline| match inline {
            Inline::TextRun(tr) => tr.content.iter().find_map(|el| match el {
                RunElement::Text(t) => Some(t.as_str()),
                _ => None,
            }),
            _ => None,
        })
        .collect();

    assert_eq!(
        texts,
        vec!["Label:", " ", "Value"],
        "whitespace-only run must survive parsing as a literal space"
    );
}

#[test]
fn convert_multi_paragraph_docx() {
    let mut body = String::new();
    for i in 0..50 {
        body.push_str(&format!(
            r#"<w:p>
                <w:pPr><w:spacing w:before="100" w:after="100"/></w:pPr>
                <w:r><w:t>Paragraph number {i} with some text content to make it wider.</w:t></w:r>
            </w:p>"#
        ));
    }
    let docx = simple_docx(&body);
    let pdf = dxpdf::convert(&docx).unwrap();
    assert_eq!(&pdf[..5], b"%PDF-");
    // Should produce a non-trivial PDF
    assert!(pdf.len() > 100);
}

/// §17.4.17 / §17.4.16: a row whose cells start with `<w:gridBefore>` (and/or
/// end with `<w:gridAfter>`) must offset its first cell to the right by that
/// many grid columns. Word emits this pattern for tables whose visible columns
/// are a subset of the `<w:tblGrid>` columns — typically a thin "phantom"
/// column at each edge to provide consistent table indentation.
///
/// Real-world repro: the "Sanitärräume" table in
/// `test-cases/ÜP - 27.02.2026-Rübenhofstr. 57, 22335 Hamburg, Nr. 7 OG2 Mitte.docx`
/// uses `<w:tblGrid>` `[38, 2905, 6872, 250]` with every row carrying
/// `<w:trPr><w:gridBefore val="1"/><w:gridAfter val="1"/>...</w:trPr>` and 2
/// cells. Without gridBefore handling, cells were placed in grid columns 0-1
/// instead of 1-2, leaving the wide "label" column rendered into the 38-twip
/// column and visually overlapping the value column.
#[test]
fn grid_before_offsets_each_row_first_cell() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let docx = simple_docx(
        r#"<w:tbl>
            <w:tblPr>
                <w:tblW w:w="10065" w:type="dxa"/>
                <w:tblLayout w:type="fixed"/>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="38"/>
                <w:gridCol w:w="2905"/>
                <w:gridCol w:w="6872"/>
                <w:gridCol w:w="250"/>
            </w:tblGrid>
            <w:tr>
                <w:trPr>
                    <w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>
                    <w:wBefore w:w="38" w:type="dxa"/>
                    <w:wAfter w:w="250" w:type="dxa"/>
                </w:trPr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2905" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>LeftA</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="6872" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>RightA</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
            <w:tr>
                <w:trPr>
                    <w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>
                    <w:wBefore w:w="38" w:type="dxa"/>
                    <w:wAfter w:w="250" w:type="dxa"/>
                </w:trPr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="2905" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>LeftB</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="6872" w:type="dxa"/></w:tcPr>
                    <w:p><w:r><w:t>RightB</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>"#,
    );

    let document = dxpdf::docx::parse(&docx).expect("parse");
    let (_, pages) = dxpdf::render::resolve_and_layout(&document);
    let cmds: Vec<&DrawCommand> = pages.iter().flat_map(|p| p.commands.iter()).collect();

    let position_of = |needle: &str| -> Option<(f32, f32)> {
        cmds.iter().find_map(|c| match c {
            DrawCommand::Text { position, text, .. } if text.as_ref() == needle => {
                Some((position.x.raw(), position.y.raw()))
            }
            _ => None,
        })
    };

    let (x_la, y_la) = position_of("LeftA").expect("LeftA present");
    let (x_ra, y_ra) = position_of("RightA").expect("RightA present");
    let (x_lb, y_lb) = position_of("LeftB").expect("LeftB present");
    let (x_rb, y_rb) = position_of("RightB").expect("RightB present");

    // Both rows align to the same columns — no per-row drift from gridBefore.
    assert!(
        (x_la - x_lb).abs() < 0.01,
        "LeftA ({x_la}) and LeftB ({x_lb}) must share the same x — both \
         rows declare gridBefore=1 so the first cell starts at the same \
         absolute grid column"
    );
    assert!(
        (x_ra - x_rb).abs() < 0.01,
        "RightA ({x_ra}) and RightB ({x_rb}) must share the same x"
    );
    assert!(y_la == y_ra, "row 0 cells share the same y baseline");
    assert!(y_lb == y_rb, "row 1 cells share the same y baseline");
    assert!(y_lb > y_la, "row 1 sits below row 0");

    // Right column must sit clearly to the right of the left column. The
    // 2905-twip left column is ~145pt wide once scaled down to fit page
    // width, so the right column's x should exceed the left column's x by
    // at least ~50pt — well beyond any rounding noise. Without gridBefore
    // handling, the right column would land just ~1.9pt past the left
    // column (the phantom 38-twip first grid column) and overlap.
    assert!(
        x_ra - x_la > 50.0,
        "RightA ({x_ra}) must be well to the right of LeftA ({x_la}) — \
         small separation indicates gridBefore was ignored and the right \
         column overlapped the left"
    );
}
