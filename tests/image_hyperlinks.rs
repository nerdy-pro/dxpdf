//! §21.1.2.3.5 / §17.16.22 — a hyperlink on a picture survives conversion
//! (issue #250).
//!
//! Word stores a link added to a picture as `<a:hlinkClick r:id="…"/>` inside
//! `<wp:docPr>`, resolved through the part's relationships like any other
//! external hyperlink; LibreOffice wraps the run holding the drawing in
//! `<w:hyperlink>` instead. Neither produced a link annotation: the `docPr`
//! child was accepted and ignored, and the image fragment did not inherit the
//! enclosing hyperlink the way text fragments do.
//!
//! Each test draws a 1×1 PNG and asserts a link annotation covering exactly
//! the rectangle the image was drawn at — the hotspot is the picture.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

const PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0xf0,
    0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
];

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

fn make_docx(body: &str) -> Vec<u8> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("[Content_Types].xml", o).unwrap();
    zip.write_all(br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#).unwrap();
    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#).unwrap();
    zip.start_file("word/_rels/document.xml.rels", o).unwrap();
    zip.write_all(br##"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/picture" TargetMode="External"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#Top" TargetMode="External"/></Relationships>"##).unwrap();
    zip.start_file("word/media/image1.png", o).unwrap();
    zip.write_all(PNG).unwrap();
    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(format!(r#"<w:document {NS}><w:body>{body}</w:body></w:document>"#).as_bytes())
        .unwrap();
    zip.finish().unwrap().into_inner()
}

/// The `a:graphic` of a 1 in × 1 in picture, shared by every placement.
const GRAPHIC: &str = r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>"#;

fn inline_drawing(doc_pr_children: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1">{doc_pr_children}</wp:docPr>{GRAPHIC}</wp:inline></w:drawing>"#
    )
}

fn anchored_drawing(doc_pr_children: &str) -> String {
    format!(
        r#"<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Picture 1">{doc_pr_children}</wp:docPr>{GRAPHIC}</wp:anchor></w:drawing>"#
    )
}

fn layout(body: &str) -> Vec<LayoutedPage> {
    let parsed = dxpdf::docx::parse(&make_docx(body)).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

fn image_rects(pages: &[LayoutedPage]) -> Vec<(f32, f32, f32, f32)> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Image { rect, .. } => Some((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.size.width.raw(),
                rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect()
}

fn external_links(pages: &[LayoutedPage]) -> Vec<((f32, f32, f32, f32), String)> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::LinkAnnotation { rect, url } => Some((
                (
                    rect.origin.x.raw(),
                    rect.origin.y.raw(),
                    rect.size.width.raw(),
                    rect.size.height.raw(),
                ),
                url.to_string(),
            )),
            _ => None,
        })
        .collect()
}

fn internal_links(pages: &[LayoutedPage]) -> Vec<((f32, f32, f32, f32), String)> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::InternalLink { rect, destination } => Some((
                (
                    rect.origin.x.raw(),
                    rect.origin.y.raw(),
                    rect.size.width.raw(),
                    rect.size.height.raw(),
                ),
                destination.to_string(),
            )),
            _ => None,
        })
        .collect()
}

/// Word's way: `a:hlinkClick` in `wp:docPr`, pointing at a hyperlink
/// relationship. The annotation covers the drawn picture exactly.
#[test]
fn an_hlink_click_on_an_inline_picture_becomes_a_link_over_the_image() {
    let pages = layout(&format!(
        r#"<w:p><w:r>{}</w:r></w:p>"#,
        inline_drawing(r#"<a:hlinkClick r:id="rId3"/>"#)
    ));
    let images = image_rects(&pages);
    assert_eq!(images.len(), 1, "one picture drawn: {images:?}");
    let links = external_links(&pages);
    assert_eq!(
        links,
        vec![(images[0], "https://example.com/picture".to_string())]
    );
}

/// LibreOffice's way: the run holding the drawing sits inside `w:hyperlink`.
/// The picture inherits the link the way the text around it would.
#[test]
fn a_picture_inside_a_hyperlink_inherits_the_link() {
    let pages = layout(&format!(
        r#"<w:p><w:hyperlink r:id="rId3"><w:r>{}</w:r></w:hyperlink></w:p>"#,
        inline_drawing("")
    ));
    let images = image_rects(&pages);
    assert_eq!(images.len(), 1);
    assert_eq!(
        external_links(&pages),
        vec![(images[0], "https://example.com/picture".to_string())]
    );
}

/// A floating picture is placed by the anchor, not the line, and the link
/// must follow it there.
#[test]
fn an_hlink_click_on_an_anchored_picture_follows_the_float() {
    let pages = layout(&format!(
        r#"<w:p><w:r>{}</w:r><w:r><w:t>text beside the picture</w:t></w:r></w:p>"#,
        anchored_drawing(r#"<a:hlinkClick r:id="rId3"/>"#)
    ));
    let images = image_rects(&pages);
    assert_eq!(images.len(), 1, "one picture drawn: {images:?}");
    assert_eq!(
        external_links(&pages),
        vec![(images[0], "https://example.com/picture".to_string())]
    );
}

/// A relationship whose target is a bare fragment names a bookmark in this
/// document, so the hotspot is a GoTo, not a URI.
#[test]
fn an_hlink_click_to_a_bookmark_is_an_internal_link() {
    let pages = layout(&format!(
        r#"<w:p><w:r>{}</w:r></w:p>"#,
        inline_drawing(r#"<a:hlinkClick r:id="rId4"/>"#)
    ));
    let images = image_rects(&pages);
    assert_eq!(images.len(), 1);
    assert!(
        external_links(&pages).is_empty(),
        "no URI annotation for a bookmark"
    );
    assert_eq!(internal_links(&pages), vec![(images[0], "Top".to_string())]);
}

/// Control: a picture with no link gets no annotation.
#[test]
fn a_picture_without_a_link_gets_no_annotation() {
    let pages = layout(&format!(r#"<w:p><w:r>{}</w:r></w:p>"#, inline_drawing("")));
    assert_eq!(image_rects(&pages).len(), 1);
    assert!(external_links(&pages).is_empty() && internal_links(&pages).is_empty());
}
