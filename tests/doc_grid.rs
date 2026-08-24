//! §17.6.5 `<w:docGrid>` — the applied line grid, end to end (issue #151).
//!
//! Every assertion is font-free: either a *relation* between two renders of
//! identical content (the committed fixture's two sections differ only in
//! `@w:type`), or a baseline **delta** that the grid quantizes to an exact
//! pitch regardless of which face the host resolved. The one precondition —
//! that the natural line is shorter than the 18pt pitch — is asserted
//! explicitly from the ungridded control, never assumed.
//!
//! The grid's scope rules (§17.6.5): body text only. A table cell's lines are
//! excluded unless the legacy `w:adjustLineHeightInTable` compat flag is
//! present (unimplemented); an opted-out paragraph (`w:snapToGrid w:val="0"`,
//! §17.3.1.32) keeps its natural spacing; `lineRule="exact"` escapes by the
//! spec's own override.

use dxpdf::render::layout::draw_command::DrawCommand;

fn make_docx(body: &str) -> Vec<u8> {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{body}</w:document>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}

fn layout(docx: &[u8]) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    let doc = dxpdf::docx::parse(docx).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Baseline y of every text command on a page, in emission order, deduplicated
/// per line (fragments on one line share a baseline).
fn baselines(page: &dxpdf::render::layout::draw_command::LayoutedPage) -> Vec<f32> {
    let mut out: Vec<f32> = Vec::new();
    for cmd in &page.commands {
        if let DrawCommand::Text { position, .. } = cmd {
            let y = position.y.raw();
            if out.last().is_none_or(|last| (last - y).abs() > 0.01) {
                out.push(y);
            }
        }
    }
    out
}

fn deltas(ys: &[f32]) -> Vec<f32> {
    ys.windows(2).map(|w| w[1] - w[0]).collect()
}

const GRID_SECT: &str = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>\
    <w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" \
     w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>\
    <w:docGrid w:type=\"lines\" w:linePitch=\"360\"/>";

/// A paragraph guaranteed to wrap into several lines at 8pt (sz=16
/// half-points) — small enough that its natural line is far below the pitch
/// on any host.
const WRAPPING: &str = "<w:p><w:pPr><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:pPr>\
    <w:r><w:rPr><w:sz w:val=\"16\"/></w:rPr><w:t>The quick brown fox jumps over \
    the lazy dog and keeps running until this paragraph is comfortably longer \
    than one line, and then keeps going further still to be safely longer than \
    two lines of the available text column on this page, and after that it \
    adds one more clause so that even a compact host face wraps it into at \
    least four lines of the four hundred and sixty eight point column, with \
    room to spare for wider faces and their metrics.</w:t></w:r></w:p>";

fn body(paragraphs: &str, sect: &str) -> String {
    format!("<w:body>{paragraphs}<w:sectPr>{sect}</w:sectPr></w:body>")
}

// ── The committed fixture: gridded section vs its type-less control ─────────

#[test]
fn fixture_grids_section_one_and_leaves_the_control_alone() {
    let bytes = std::fs::read("test-files/doc-grid.docx").expect("fixture");
    let pages = layout(&bytes);
    assert!(pages.len() >= 2, "control section starts its own page");

    // Page 1 (type="lines"): the Latin and CJK paragraphs' in-paragraph
    // deltas are exactly the 18pt pitch.
    let gridded = baselines(&pages[0]);
    let control = baselines(&pages[1]);
    assert!(gridded.len() >= 3, "wrapping paragraphs give several lines");
    assert_eq!(
        gridded.len(),
        control.len(),
        "identical content, one section gridded"
    );

    // Precondition for everything below: the natural line is shorter than
    // the pitch (8–20pt text on an 18pt grid; the control's first delta is
    // the natural advance).
    let natural = control[1] - control[0];
    assert!(
        natural < 18.0,
        "natural line ({natural}pt) must sit under the 18pt pitch"
    );

    // The wrapping Latin paragraph: control deltas are natural, gridded are
    // the pitch.
    let d_gridded = deltas(&gridded);
    let d_control = deltas(&control);
    assert!((d_gridded[0] - 18.0).abs() < 1e-3, "gridded delta = pitch");
    assert!((d_control[0] - natural).abs() < 1e-3);

    // §17.6.5 centering: the first gridded baseline sits (pitch − natural)/2
    // below the first control baseline — both pages start at the same top
    // margin with the same content, so everything but the shift cancels.
    let shift = gridded[0] - control[0];
    assert!(
        (shift - (18.0 - natural) / 2.0).abs() < 1e-3,
        "line centered in its slot: shift {shift} vs natural {natural}"
    );
}

// ── In-memory probes ────────────────────────────────────────────────────────

/// type="lines" quantizes every in-paragraph advance to the pitch.
#[test]
fn gridded_lines_advance_by_whole_pitches() {
    let pages = layout(&make_docx(&body(WRAPPING, GRID_SECT)));
    let ys = baselines(&pages[0]);
    assert!(ys.len() >= 3, "the probe paragraph must wrap");
    for d in deltas(&ys) {
        assert!(
            (d - 18.0).abs() < 1e-3,
            "delta {d} should be the 18pt pitch"
        );
    }
}

/// The identical document minus `@w:type` — the docGrid every Word file
/// carries — must lay out at natural spacing. The regression this pins:
/// applying `linePitch="360"` to type-less grids would re-space every
/// Western document to 18pt.
#[test]
fn typeless_doc_grid_changes_nothing() {
    let gridless = GRID_SECT.replace(" w:type=\"lines\"", "");
    let pages = layout(&make_docx(&body(WRAPPING, &gridless)));
    let ys = baselines(&pages[0]);
    assert!(ys.len() >= 3);
    for d in deltas(&ys) {
        assert!(d < 17.0, "natural 8pt advance, not the pitch (got {d})");
    }
}

/// §17.3.1.32: `<w:snapToGrid w:val="0"/>` opts the paragraph out.
#[test]
fn snap_to_grid_off_escapes_the_grid() {
    let opted_out = WRAPPING.replace("<w:pPr>", "<w:pPr><w:snapToGrid w:val=\"0\"/>");
    let pages = layout(&make_docx(&body(&opted_out, GRID_SECT)));
    let ys = baselines(&pages[0]);
    assert!(ys.len() >= 3);
    for d in deltas(&ys) {
        assert!(
            d < 17.0,
            "opted-out paragraph keeps natural spacing (got {d})"
        );
    }
}

/// §17.6.5's own override: `lineRule="exact"` escapes the grid.
#[test]
fn exact_line_spacing_escapes_the_grid() {
    let exact = WRAPPING.replace(
        "<w:pPr>",
        "<w:pPr><w:spacing w:line=\"240\" w:lineRule=\"exact\"/>",
    );
    let pages = layout(&make_docx(&body(&exact, GRID_SECT)));
    let ys = baselines(&pages[0]);
    assert!(ys.len() >= 3);
    for d in deltas(&ys) {
        assert!(
            (d - 12.0).abs() < 1e-3,
            "exact 240tw = 12pt, not the pitch (got {d})"
        );
    }
}

/// A line taller than one pitch takes two — Word's signature spacing jump.
/// 20pt text has a natural line in (18, 36] on any host.
#[test]
fn tall_lines_take_two_slots() {
    let tall = WRAPPING.replace("w:val=\"16\"", "w:val=\"40\"");
    let pages = layout(&make_docx(&body(&tall, GRID_SECT)));
    let ys = baselines(&pages[0]);
    assert!(ys.len() >= 2, "20pt text still wraps");
    for d in deltas(&ys) {
        assert!((d - 36.0).abs() < 1e-3, "two 18pt slots (got {d})");
    }
}

/// An empty paragraph on a gridded page holds a whole slot open, like Word's
/// empty 行: text – empty – text puts the third paragraph's baseline exactly
/// two pitches under the first.
#[test]
fn empty_paragraph_consumes_one_slot() {
    let one_line = "<w:p><w:pPr><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:pPr>\
        <w:r><w:rPr><w:sz w:val=\"16\"/></w:rPr><w:t>line</w:t></w:r></w:p>";
    let paragraphs = format!("{one_line}<w:p/>{one_line}");
    let pages = layout(&make_docx(&body(&paragraphs, GRID_SECT)));
    let ys = baselines(&pages[0]);
    assert_eq!(ys.len(), 2, "two text lines around the empty paragraph");
    assert!(
        (ys[1] - ys[0] - 36.0).abs() < 1e-3,
        "text + empty slot = two pitches (got {})",
        ys[1] - ys[0]
    );
}

/// §17.3.1.15 × §17.6.5: the keepNext pre-measure must see gridded heights.
/// The page is filled so that the heading + follower group fits ungridded
/// (~30pt) but not gridded (3 slots = 54pt > 36pt left): a measurement that
/// ignored the grid would hold the heading at the page bottom and tear it
/// from its follower.
#[test]
fn keep_next_group_is_measured_at_gridded_heights() {
    let one_line = "<w:p><w:pPr><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:pPr>\
        <w:r><w:rPr><w:sz w:val=\"16\"/></w:rPr><w:t>filler</w:t></w:r></w:p>";
    // 648pt of body = 36 slots; 34 filler lines leave exactly 2 slots.
    let filler = one_line.repeat(34);
    let heading = "<w:p><w:pPr><w:keepNext/><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:pPr>\
        <w:r><w:rPr><w:sz w:val=\"16\"/></w:rPr><w:t>HEADING</w:t></w:r></w:p>";
    // The follower wraps to two lines at 8pt in a 468pt column.
    let follower = "<w:p><w:pPr><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:pPr>\
        <w:r><w:rPr><w:sz w:val=\"16\"/></w:rPr><w:t>FOLLOWER starts here and \
        keeps going with enough further words that even a compact host face has \
        to wrap this follower paragraph onto a second line of the four hundred \
        and sixty eight point text column with plenty of margin to spare for \
        wide faces.</w:t></w:r></w:p>";
    let paragraphs = format!("{filler}{heading}{follower}");
    let pages = layout(&make_docx(&body(&paragraphs, GRID_SECT)));

    let page_of = |needle: &str| {
        pages.iter().position(|p| {
            p.commands
                .iter()
                .any(|c| matches!(c, DrawCommand::Text { text, .. } if text.contains(needle)))
        })
    };
    let heading_page = page_of("HEADING").expect("heading rendered");
    let follower_page = page_of("FOLLOWER").expect("follower rendered");
    assert_eq!(
        heading_page, follower_page,
        "keepNext holds the heading with its follower under the grid"
    );
    assert_eq!(heading_page, 1, "the group moved off the full first page");
}

/// §17.6.5: "the line pitch shall not be added to any line which appears
/// within a table cell" (without the unimplemented legacy compat flag). The
/// cell's wrapped lines keep natural spacing while a body paragraph on the
/// same gridded page snaps.
#[test]
fn table_cell_lines_are_not_gridded() {
    let cell_text = "The quick brown fox jumps over the lazy dog and keeps \
        running until this cell paragraph is comfortably longer than one line \
        of its narrow cell.";
    let table = format!(
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>\
         <w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>\
         <w:p><w:pPr><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:pPr>\
         <w:r><w:rPr><w:sz w:val=\"16\"/></w:rPr><w:t>{cell_text}</w:t></w:r></w:p>\
         </w:tc></w:tr></w:tbl>"
    );
    let paragraphs = format!("{WRAPPING}{table}");
    let pages = layout(&make_docx(&body(&paragraphs, GRID_SECT)));
    let ys = baselines(&pages[0]);
    let ds = deltas(&ys);
    // The body paragraph's wrapped deltas lead; the cell's follow. The body
    // ones are the pitch, the cell's are natural — so the set must contain
    // both, and no cell delta may equal the pitch.
    let (snapped, natural): (Vec<f32>, Vec<f32>) =
        ds.iter().partition(|d| (**d - 18.0).abs() < 1e-3);
    assert!(!snapped.is_empty(), "body deltas snap to the pitch");
    assert!(
        natural.iter().any(|d| *d > 1.0 && *d < 17.0),
        "cell deltas stay natural: {ds:?}"
    );
}
