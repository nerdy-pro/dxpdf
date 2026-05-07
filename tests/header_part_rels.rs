//! Tests that header / footer / footnote / endnote parts get their
//! own per-part `_rels/<part>.xml.rels` namespace honored, both for
//! image rIds and hyperlink rIds. Pre-fix the parser merged every
//! image rel into one global `media` map (rIds collided across
//! parts) and resolved hyperlinks with the wrong rels file
//! (`document.xml.rels` instead of the part's own).

use std::io::Write;

use dxpdf::model::{Block, HyperlinkTarget, Inline};

/// Build a multi-part DOCX with overlapping `rId` values across the
/// document's rels and a header part's rels.
///
/// Layout:
/// ```text
/// [Content_Types].xml
/// _rels/.rels                     → word/document.xml
/// word/document.xml               header reference + body
/// word/_rels/document.xml.rels    rId-hdr → header1.xml; rId-link → WRONG_URL
/// word/header1.xml                contains a hyperlink with r:id="rId-link"
/// word/_rels/header1.xml.rels     rId-link → RIGHT_URL
/// ```
fn docx_with_header_hyperlink(right_url: &str, wrong_url: &str) -> Vec<u8> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let opts = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    let write = |zip: &mut zip::ZipWriter<_>, path: &str, body: &str| {
        zip.start_file(path, opts).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    };

    write(
        &mut zip,
        "[Content_Types].xml",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>"#,
    );

    write(
        &mut zip,
        "_rels/.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"#,
    );

    write(
        &mut zip,
        "word/document.xml",
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId-hdr"/>
    </w:sectPr>
  </w:body>
</w:document>"#,
    );

    // document.xml.rels — has a *colliding* `rId-link` pointing at a
    // different URL. Pre-fix that's what header hyperlinks resolved
    // against (wrongly).
    let doc_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId-hdr"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    Target="header1.xml"/>
  <Relationship Id="rId-link"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="{wrong_url}" TargetMode="External"/>
</Relationships>"#
    );
    write(&mut zip, "word/_rels/document.xml.rels", &doc_rels);

    write(
        &mut zip,
        "word/header1.xml",
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:hyperlink r:id="rId-link">
      <w:r><w:t>link</w:t></w:r>
    </w:hyperlink>
  </w:p>
</w:hdr>"#,
    );

    let header_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId-link"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="{right_url}" TargetMode="External"/>
</Relationships>"#
    );
    write(&mut zip, "word/_rels/header1.xml.rels", &header_rels);

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

/// Walk the parsed blocks of a header looking for the first
/// hyperlink and return its target URL.
fn find_first_hyperlink_url(blocks: &[Block]) -> Option<String> {
    for block in blocks {
        if let Block::Paragraph(p) = block {
            for inline in &p.content {
                if let Inline::Hyperlink(h) = inline {
                    if let HyperlinkTarget::External(url) = &h.target {
                        return Some(url.as_str().to_string());
                    }
                }
            }
        }
    }
    None
}

#[test]
fn header_hyperlink_resolves_against_header_rels_not_document_rels() {
    let docx = docx_with_header_hyperlink(
        "https://example.com/header-target",
        "https://example.com/document-wrong",
    );
    let doc = dxpdf::docx::parse(&docx).expect("parse");

    assert_eq!(doc.headers.len(), 1, "expected one header part");
    let (_, blocks) = doc.headers.iter().next().unwrap();
    let url = find_first_hyperlink_url(blocks).expect("header has a hyperlink");

    assert_eq!(
        url, "https://example.com/header-target",
        "header hyperlink must resolve through `header1.xml.rels`, \
         not the document's main rels"
    );
    assert_ne!(
        url, "https://example.com/document-wrong",
        "if this fires, hyperlink resolution is still using doc_rels"
    );
}
