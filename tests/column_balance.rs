//! §17.6.4 / §17.6.22 — balancing a column run at a continuous break, and the
//! `w:sep` rule down the gutter (issue #232).
//!
//! Both behaviours were measured on Word 16.0 (Office 2024 LTSC), on a Letter
//! page with 54 pt margins, two equal columns and a 25 pt gutter:
//!
//! | document | Word |
//! |---|---|
//! | column run closed by `<w:type w:val="continuous"/>` | 3 paragraphs in each column — balanced — and the following one-column section directly beneath them, on the same page |
//! | column run closed by `<w:type w:val="nextPage"/>` | all six paragraphs in column 1, tail section on page 2 |
//! | column run that ends the document | all six paragraphs in column 1 |
//!
//! So balancing is not a property of multi-column sections; it is what a
//! *continuous* break does to the run it closes. The greedy flow this engine
//! already performs is the right answer for the other two.
//!
//! The rule: Word drew a hairline at x = 305.76 on a gutter centred at 306,
//! running from the body top to the bottom of the deepest column — the used
//! height, not the page's. In the two documents where every paragraph stayed in
//! column 1 it drew no rule at all, so an unoccupied column gets no divider.

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

/// The §17.6.13 default page, two equal columns, 500-twip gutter: content
/// 72…540, so each column is 221.5 pt and the second starts at 318.5.
const COL1_X: f32 = 72.0;
const COL2_X: f32 = 318.5;
const GUTTER_CENTRE: f32 = 306.0;

const COLS: &str = r#"<w:cols w:num="2" w:space="500" w:sep="1"/>"#;
const COLS_NO_SEP: &str = r#"<w:cols w:num="2" w:space="500"/>"#;

/// Six short paragraphs in a two-column section, closed by `tail_type`, then a
/// one-column tail section.
fn two_sections(cols: &str, tail_type: &str) -> Vec<LayoutedPage> {
    let paras: String = (1..=6)
        .map(|i| format!("<w:p><w:r><w:t>P{i}</w:t></w:r></w:p>"))
        .collect();
    let bytes = make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {paras}
    <w:p><w:pPr><w:sectPr>{cols}</w:sectPr></w:pPr></w:p>
    <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
    <w:sectPr>{tail_type}<w:cols w:space="720"/></w:sectPr>
  </w:body>
</w:document>"#
    ));
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(parsed).1
}

/// Where a word was drawn, as (page index, x, y).
fn placed(pages: &[LayoutedPage], needle: &str) -> (usize, f32, f32) {
    pages
        .iter()
        .enumerate()
        .find_map(|(i, p)| {
            p.commands.iter().find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if &**text == needle => {
                    Some((i, position.x.raw(), position.y.raw()))
                }
                _ => None,
            })
        })
        .unwrap_or_else(|| panic!("no Text command says {needle:?}"))
}

/// Every vertical line on a page, as (x, top, bottom).
fn vertical_rules(page: &LayoutedPage) -> Vec<(f32, f32, f32)> {
    page.commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Line { line, .. }
                if (line.start.x.raw() - line.end.x.raw()).abs() < 0.01 =>
            {
                Some((
                    line.start.x.raw(),
                    line.start.y.raw().min(line.end.y.raw()),
                    line.start.y.raw().max(line.end.y.raw()),
                ))
            }
            _ => None,
        })
        .collect()
}

/// The reported defect: six paragraphs that fit a column twice over were all
/// left in column 1, and the tail section was pushed to page 2. Word splits
/// them 3 / 3 and keeps everything on page 1.
#[test]
fn a_continuous_break_balances_the_columns() {
    let pages = two_sections(COLS, r#"<w:type w:val="continuous"/>"#);

    assert_eq!(pages.len(), 1, "everything fits one page once balanced");
    for para in ["P1", "P2", "P3"] {
        assert!(
            (placed(&pages, para).1 - COL1_X).abs() < 0.5,
            "{para} should be in column 1, was drawn at x={:.2}",
            placed(&pages, para).1,
        );
    }
    for para in ["P4", "P5", "P6"] {
        assert!(
            (placed(&pages, para).1 - COL2_X).abs() < 0.5,
            "{para} should be in column 2, was drawn at x={:.2}",
            placed(&pages, para).1,
        );
    }
    let (_, tail_x, tail_y) = placed(&pages, "Tail");
    assert!(
        (tail_x - COL1_X).abs() < 0.5,
        "the tail section spans the full width",
    );
    assert!(
        tail_y > placed(&pages, "P3").2,
        "and sits below the balanced columns",
    );
}

/// §17.6.22: a `nextPage` break starts its own page, and Word does not balance
/// the run it closes — the greedy flow is the right answer there.
#[test]
fn a_next_page_break_does_not_balance() {
    let pages = two_sections(COLS, r#"<w:type w:val="nextPage"/>"#);

    for para in ["P1", "P6"] {
        assert!(
            (placed(&pages, para).1 - COL1_X).abs() < 0.5,
            "{para} stays in column 1 when the break is nextPage",
        );
    }
    assert_eq!(
        placed(&pages, "Tail").0,
        1,
        "the tail section starts a page"
    );
}

/// A column run that ends the document is not closed by a break at all, and
/// Word leaves it greedy too.
#[test]
fn a_section_that_ends_the_document_does_not_balance() {
    let paras: String = (1..=6)
        .map(|i| format!("<w:p><w:r><w:t>P{i}</w:t></w:r></w:p>"))
        .collect();
    let bytes = make_docx(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{paras}<w:sectPr>{COLS}</w:sectPr></w:body>
</w:document>"#
    ));
    let parsed = dxpdf::docx::parse(&bytes).expect("parse");
    let pages = dxpdf::render::resolve_and_layout(parsed).1;

    for para in ["P1", "P6"] {
        assert!(
            (placed(&pages, para).1 - COL1_X).abs() < 0.5,
            "{para} stays in column 1 at the end of the document",
        );
    }
}

/// §17.6.4 `w:sep`: one hairline down the gutter centre, spanning the height
/// the columns used.
#[test]
fn w_sep_draws_a_rule_down_the_gutter() {
    let pages = two_sections(COLS, r#"<w:type w:val="continuous"/>"#);
    let rules = vertical_rules(&pages[0]);

    assert_eq!(rules.len(), 1, "one gutter, one rule: {rules:?}");
    let (x, top, bottom) = rules[0];
    assert!(
        (x - GUTTER_CENTRE).abs() < 0.05,
        "the rule sits at the gutter centre {GUTTER_CENTRE}, drawn at {x:.2}",
    );
    assert!(
        (top - 72.0).abs() < 0.5,
        "it starts at the body top, drawn at {top:.2}",
    );
    let deepest = placed(&pages, "P3").2.max(placed(&pages, "P6").2);
    assert!(
        bottom > deepest && bottom < deepest + 30.0,
        "and ends just under the deepest column ({deepest:.2}), drawn at {bottom:.2}",
    );
}

/// An unoccupied column gets no divider — Word drew none in either of the two
/// probes where the content never left column 1.
#[test]
fn no_rule_is_drawn_when_the_second_column_is_empty() {
    let pages = two_sections(COLS, r#"<w:type w:val="nextPage"/>"#);

    assert!(
        vertical_rules(&pages[0]).is_empty(),
        "column 2 took no content, so there is nothing to separate",
    );
}

/// And none without `w:sep`, however full the columns are.
#[test]
fn no_rule_without_w_sep() {
    let pages = two_sections(COLS_NO_SEP, r#"<w:type w:val="continuous"/>"#);

    assert!(vertical_rules(&pages[0]).is_empty());
}
