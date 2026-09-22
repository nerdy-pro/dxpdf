//! §17.9 — two `w:num` instances over one `w:abstractNum` share a sequence
//! (issue #230).
//!
//! `w:abstractNum` is the list definition and carries the counter; `w:num` is an
//! *instance* of it, and its reason to exist is §17.9.7's level overrides — a
//! document applies `w:startOverride` by declaring a second instance rather than
//! by duplicating the definition. Counting per instance therefore splits one
//! list into two whenever an author does that, which is what a user's
//! specification document showed: its last two headings drew 6 and 7 where Word
//! draws 9 and 10, because the engine's counter for the style's `numId` had been
//! frozen at 5 while a second `numId` ran its own sequence from the override.
//!
//! The counters here are keyed by the abstract definition, and `w:startOverride`
//! is a one-shot restart of that shared sequence rather than a permanent start
//! value.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

/// A document with `word/numbering.xml`: one abstract definition, decimal, and
/// the instances `body_numbering` declares.
fn make_docx(document_body: &str, numbering_instances: &str) -> Vec<u8> {
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
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
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

    zip.start_file("word/_rels/document.xml.rels", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Target="numbering.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/numbering.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="3">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/>
    </w:lvl>
  </w:abstractNum>
  {numbering_instances}
</w:numbering>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{document_body}</w:body>
</w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

/// Two instances of abstract 3; the second carries `overrides`.
fn instances(overrides: &str) -> String {
    format!(
        r#"<w:num w:numId="1"><w:abstractNumId w:val="3"/></w:num>
           <w:num w:numId="2"><w:abstractNumId w:val="3"/>{overrides}</w:num>"#
    )
}

/// One numbered paragraph on `num_id` at `ilvl`, whose body text is `text`.
fn item(num_id: u32, ilvl: u32, text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>
             <w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn layout(bytes: &[u8]) -> Vec<LayoutedPage> {
    let parsed = dxpdf::docx::parse(bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// Every drawn string, in page order — labels and body text interleaved, which
/// is what makes "which label went with which item" readable.
fn drawn_text(pages: &[LayoutedPage]) -> Vec<String> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { text, .. } => Some(text.to_string()),
            _ => None,
        })
        .collect()
}

/// The label drawn for the item whose body text is `text`: the drawn string
/// immediately before it.
fn label_of(pages: &[LayoutedPage], text: &str) -> String {
    let drawn = drawn_text(pages);
    let at = drawn
        .iter()
        .position(|s| s == text)
        .unwrap_or_else(|| panic!("nothing drew {text:?}; drawn: {drawn:?}"));
    assert!(at > 0, "{text:?} has nothing before it to be its label");
    drawn[at - 1].clone()
}

/// The reported defect, at its simplest: no overrides anywhere, so the two
/// instances are indistinguishable and the sequence must run straight through.
#[test]
fn two_instances_of_one_abstract_share_one_sequence() {
    let body = format!(
        "{}{}{}{}{}",
        item(1, 0, "one"),
        item(1, 0, "two"),
        item(2, 0, "three"),
        item(2, 0, "four"),
        item(1, 0, "five"),
    );
    let pages = layout(&make_docx(&body, &instances("")));

    let labels: Vec<String> = ["one", "two", "three", "four", "five"]
        .iter()
        .map(|t| label_of(&pages, t))
        .collect();
    assert_eq!(
        labels,
        vec!["1.", "2.", "3.", "4.", "5."],
        "one abstract definition is one list, however many instances point at it",
    );
}

/// §17.9.28: `w:startOverride` restarts the shared sequence where the instance
/// is first used, and everything after it — on either instance — continues from
/// there. This is the user's document in miniature: 1, 2, then an override to
/// 8, then back to the first instance for 9 and 10.
#[test]
fn a_start_override_restarts_the_shared_sequence() {
    let body = format!(
        "{}{}{}{}{}",
        item(1, 0, "one"),
        item(1, 0, "two"),
        item(2, 0, "eight"),
        item(1, 0, "nine"),
        item(1, 0, "ten"),
    );
    let pages = layout(&make_docx(
        &body,
        &instances(r#"<w:lvlOverride w:ilvl="0"><w:startOverride w:val="8"/></w:lvlOverride>"#),
    ));

    let labels: Vec<String> = ["one", "two", "eight", "nine", "ten"]
        .iter()
        .map(|t| label_of(&pages, t))
        .collect();
    assert_eq!(labels, vec!["1.", "2.", "8.", "9.", "10."]);
}

/// The restart is one-shot. Returning to the overriding instance later must
/// continue the sequence rather than jump back to the override's value — which
/// is the difference between `w:startOverride` and `w:start`, and the reason the
/// two cannot be collapsed into one field.
#[test]
fn a_start_override_does_not_fire_a_second_time() {
    let body = format!(
        "{}{}{}{}",
        item(1, 0, "one"),
        item(2, 0, "eight"),
        item(1, 0, "nine"),
        item(2, 0, "ten"),
    );
    let pages = layout(&make_docx(
        &body,
        &instances(r#"<w:lvlOverride w:ilvl="0"><w:startOverride w:val="8"/></w:lvlOverride>"#),
    ));

    let labels: Vec<String> = ["one", "eight", "nine", "ten"]
        .iter()
        .map(|t| label_of(&pages, t))
        .collect();
    assert_eq!(labels, vec!["1.", "8.", "9.", "10."]);
}

/// A second abstract definition is a second list: sharing is by abstract id, not
/// by "any two instances".
#[test]
fn instances_of_different_abstracts_keep_separate_counters() {
    let body = format!(
        "{}{}{}",
        item(1, 0, "a1"),
        item(2, 0, "b1"),
        item(1, 0, "a2"),
    );
    let two_abstracts = r#"<w:num w:numId="1"><w:abstractNumId w:val="3"/></w:num>
        <w:num w:numId="2"><w:abstractNumId w:val="4"/></w:num>"#;
    let bytes = make_docx(&body, two_abstracts);
    // Splice in the second abstract definition the instances refer to.
    let bytes = {
        let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut numbering = String::new();
        {
            use std::io::Read;
            z.by_name("word/numbering.xml")
                .unwrap()
                .read_to_string(&mut numbering)
                .unwrap();
        }
        let second = r#"<w:abstractNum w:abstractNumId="4">
            <w:multiLevelType w:val="singleLevel"/>
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
              <w:lvlText w:val="%1."/></w:lvl>
          </w:abstractNum>"#;
        let patched = numbering.replace(
            "<w:num w:numId=\"1\"",
            &format!("{second}<w:num w:numId=\"1\""),
        );
        let mut out = std::io::Cursor::new(Vec::new());
        {
            let mut w = zip::ZipWriter::new(&mut out);
            let o = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for i in 0..z.len() {
                let mut f = z.by_index(i).unwrap();
                let name = f.name().to_string();
                let mut buf = Vec::new();
                std::io::copy(&mut f, &mut buf).unwrap();
                w.start_file(&name, o).unwrap();
                if name == "word/numbering.xml" {
                    w.write_all(patched.as_bytes()).unwrap();
                } else {
                    w.write_all(&buf).unwrap();
                }
            }
            w.finish().unwrap();
        }
        out.into_inner()
    };
    let pages = layout(&bytes);

    assert_eq!(
        ["a1", "b1", "a2"].map(|t| label_of(&pages, t)),
        ["1.", "1.", "2."],
        "two abstract definitions are two lists",
    );
}

/// §17.9.9's multi-level template reads the same counters, so a `%1.%2` label
/// must see the shared level-0 value: an item on the second instance nests
/// under the first instance's current number.
#[test]
fn a_multilevel_label_reads_the_shared_ancestor_counter() {
    let body = format!(
        "{}{}{}",
        item(1, 0, "one"),
        item(1, 0, "two"),
        item(2, 1, "nested"),
    );
    let pages = layout(&make_docx(&body, &instances("")));

    assert_eq!(label_of(&pages, "nested"), "2.1.");
}

/// §17.9.28, the case a real Word-rendered fixture settles: instantiating a
/// level as an *ancestor* consumes that instance's restart for it.
///
/// `test-files/numbering-direct-indent.docx` opens with an `ilvl=2` item on an
/// instance that overrides level 0, and Word continues 2. 3. 4. at the top level
/// afterwards — it does not restart when the top level is finally used
/// directly. So the one-shot is spent the first time the instance touches the
/// level, whichever way it touches it.
#[test]
fn a_deep_item_consumes_its_instances_restart_for_the_ancestor() {
    let body = format!(
        "{}{}{}",
        item(2, 1, "deep"),  // instantiates level 0 of numId 2's abstract
        item(2, 0, "after"), // …so the override must not fire here
        item(2, 0, "next"),
    );
    let pages = layout(&make_docx(
        &body,
        &instances(r#"<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>"#),
    ));

    assert_eq!(
        ["deep", "after", "next"].map(|t| label_of(&pages, t)),
        ["1.1.", "2.", "3."],
    );
}
