//! §17.3.1.29 / §17.13.6.1 — a paragraph holding nothing but bookmarks is one
//! line, not two (issue #228).
//!
//! `w:bookmarkStart` / `w:bookmarkEnd` are range annotations: no glyphs, no
//! advance, no line box. A paragraph containing only them still shows its
//! paragraph mark, so it occupies exactly one line — the same line an empty
//! `<w:p/>` occupies.
//!
//! This is not a curiosity. Word writes the automatic `_GoBack` bookmark into
//! almost every document it saves, and it usually lands in the trailing empty
//! paragraph. Charging that paragraph two lines costs a whole page whenever the
//! last page is full to within one line, which is what a user's
//! image-plus-trailing-paragraph document did: two pages where Word gives one.
//!
//! The fitter is what miscounts. `MarkLine::of` sees no fragment whose
//! `occupies_line()` is true and injects a synthetic `Fragment::LineBreak` at
//! the *front* of the vector, so the paragraph becomes `[LineBreak, Bookmark]`;
//! `fit_lines` ends a line at a `LineBreak` and then emits the stranded
//! `Bookmark` as a second line, which `paragraph::line_height_for` charges a
//! full line box because it measures zero.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn make_docx(document_xml: &str) -> Vec<u8> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", o).unwrap();
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

    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

/// `TOP`, then `middle`, then `BOTTOM`, on the §17.6.13 default Letter page.
fn sandwich(middle: &str) -> Vec<LayoutedPage> {
    let bytes = make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>TOP</w:t></w:r></w:p>
    {middle}
    <w:p><w:r><w:t>BOTTOM</w:t></w:r></w:p>
  </w:body>
</w:document>"#
    ));
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// The y of the first `Text` command whose text is `needle`.
fn text_y(pages: &[LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { position, text, .. } if &**text == needle => Some(position.y.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no Text command says {needle:?}"))
}

const BOOKMARK_ONLY: &str =
    r#"<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>"#;

/// The property: a bookmark-only paragraph advances the cursor by exactly what
/// an empty paragraph advances it by. Stated as the distance the *following*
/// paragraph moves, because that is what a page break is decided on.
#[test]
fn a_bookmark_only_paragraph_advances_one_line() {
    let none = text_y(&sandwich(""), "BOTTOM");
    let empty = text_y(&sandwich("<w:p/>"), "BOTTOM");
    let bookmark = text_y(&sandwich(BOOKMARK_ONLY), "BOTTOM");

    let one_line = empty - none;
    assert!(
        one_line > 0.0,
        "the control is malformed: an empty paragraph must advance the page",
    );
    assert_eq!(
        bookmark,
        empty,
        "a bookmark-only paragraph advances one line ({one_line} pt), like `<w:p/>`; \
         it advanced {} pt",
        bookmark - none,
    );
}

/// Two bookmarks in one paragraph are still one line — the count of
/// annotations cannot change the paragraph's height.
#[test]
fn two_bookmarks_in_one_paragraph_are_still_one_line() {
    let empty = text_y(&sandwich("<w:p/>"), "BOTTOM");
    let two = text_y(
        &sandwich(
            r#"<w:p><w:bookmarkStart w:id="0" w:name="a"/><w:bookmarkEnd w:id="0"/>
               <w:bookmarkStart w:id="1" w:name="b"/><w:bookmarkEnd w:id="1"/></w:p>"#,
        ),
        "BOTTOM",
    );

    assert_eq!(two, empty, "two annotations are as tall as none");
}

/// The line the bookmark was stranded on is gone; the destination it anchors
/// must not be. Internal cross-references (`REF`, `PAGEREF`, a link to a
/// heading) resolve through this command.
#[test]
fn a_bookmark_only_paragraph_still_emits_its_destination() {
    let pages = sandwich(BOOKMARK_ONLY);
    let found = pages
        .iter()
        .flat_map(|p| &p.commands)
        .any(|c| matches!(c, DrawCommand::NamedDestination { name, .. } if name == "_GoBack"));

    assert!(found, "the bookmark still anchors a named destination");
}

/// The consequence that made this worth fixing: a page with room for one more
/// line must not gain a second page because the trailing paragraph carries
/// Word's `_GoBack`. The filler is sized against the empty-paragraph control,
/// so the test asserts a *difference* and cannot drift with font metrics.
#[test]
fn a_bookmark_only_paragraph_does_not_add_a_page() {
    // A Letter page with one-inch margins holds 648 pt of body, i.e. 56 lines
    // of 11.5 pt. 55 fillers plus one empty paragraph is 56 and fits; a second
    // line for that paragraph is 57 and does not. The control asserts the first
    // half before the bookmark case is read.
    let filler: String = (0..55)
        .map(|i| format!("<w:p><w:r><w:t>filler {i}</w:t></w:r></w:p>"))
        .collect();
    let body = |middle: &str| {
        let bytes = make_docx(&format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{filler}{middle}</w:body>
</w:document>"#
        ));
        let parsed = dxpdf::docx::parse(&bytes).expect("parse");
        dxpdf::render::resolve_and_layout(parsed).1.len()
    };

    let empty = body("<w:p/>");
    assert_eq!(
        empty, 1,
        "the control must fit one page for the test to mean anything"
    );
    assert_eq!(
        body(BOOKMARK_ONLY),
        empty,
        "a trailing `_GoBack` paragraph must not cost a page",
    );
}
