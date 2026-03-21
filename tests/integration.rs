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
fn parse_simple_docx() {
    let docx = simple_docx(r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#);
    let doc = dxpdf::parse::parse(&docx).unwrap();
    assert_eq!(doc.blocks.len(), 1);
    match &doc.blocks[0] {
        dxpdf::model::Block::Paragraph(p) => {
            assert_eq!(p.runs.len(), 1);
            match &p.runs[0] {
                dxpdf::model::Inline::TextRun(tr) => {
                    assert_eq!(tr.text, "Hello World");
                }
                _ => panic!("Expected TextRun"),
            }
        }
        _ => panic!("Expected Paragraph"),
    }
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
    let result = dxpdf::parse::parse(b"not a zip file");
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

    let result = dxpdf::parse::parse(&bytes);
    assert!(result.is_err());
    let err = result.unwrap_err().to_string();
    assert!(
        err.contains("document.xml"),
        "Error should mention document.xml: {err}"
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
