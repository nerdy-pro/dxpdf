//! Table geometry as it reaches the page — §17.18.2 `double` borders and
//! §17.4.83 `w:vAlign`, asserted end to end.
//!
//! Both are pinned in `render::layout::table`'s own unit tests, against the
//! layout inputs. What this file adds is the half of the path those cannot
//! reach: `w:sz` in eighths of a point becoming a `Pt` width, `w:val="double"`
//! surviving `convert_model_border`, `w:vAlign` surviving the §17.7.6 cascade,
//! and the whole thing arriving on a page rather than in a `TableSlice`.
//!
//! Every number below is a *difference between two measurements of the same
//! document*, so no glyph metric, cell margin or page origin has to be known:
//! whatever those contribute, they contribute equally to both sides and cancel.
//! That is what lets these assertions be exact without pinning the metrics of
//! whichever font the host happens to resolve.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn make_docx(document_xml: &str) -> Vec<u8> {
    let mut buf = Vec::new();
    {
        let mut zip = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
        let o = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", o).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

        zip.start_file("_rels/.rels", o).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/document.xml", o).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();
        zip.finish().unwrap();
    }
    buf
}

fn document(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>"#
    )
}

fn layout(body: &str) -> Vec<LayoutedPage> {
    let doc = dxpdf::docx::parse(&make_docx(&document(body))).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Every `Rect` on the page as `(x, y, w, h)`.
fn rects(pages: &[LayoutedPage]) -> Vec<(f32, f32, f32, f32)> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => Some((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.size.width.raw(),
                rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect()
}

/// The baseline y of the text command whose content is exactly `needle`.
fn baseline_of(pages: &[LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == needle => Some(position.y.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text {needle:?} on the page"))
}

// ── §17.18.2 `w:val="double"` ───────────────────────────────────────────────

/// A one-cell table whose every edge is `style` at `w:sz="24"`.
///
/// 24 eighths of a point is 3pt, chosen because it divides by three exactly —
/// `1 / 3` is not representable in binary, and a `w:sz` that made the sub-line
/// width inexact would put the test's own rounding between it and the defect it
/// is looking for.
fn one_cell_table(style: &str) -> String {
    let edges: String = ["top", "left", "bottom", "right", "insideH", "insideV"]
        .iter()
        .map(|e| format!(r#"<w:{e} w:val="{style}" w:sz="24" w:space="0" w:color="auto"/>"#))
        .collect();
    format!(
        r#"<w:tbl>
  <w:tblPr>
    <w:tblW w:w="4000" w:type="dxa"/>
    <w:tblBorders>{edges}</w:tblBorders>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
  <w:tr><w:tc>
    <w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
    <w:p><w:r><w:t>CELL</w:t></w:r></w:p>
  </w:tc></w:tr>
</w:tbl>"#
    )
}

/// One border line along its **thin** axis: where it starts, and how thick it
/// is. That is the only axis a `double` splits on, so it is the only one these
/// tests read.
type Run = (f32, f32);

/// The **distinct** horizontal border runs as `(y, height)`, sorted by y, and
/// the verticals as `(x, width)`, sorted by x.
///
/// Distinct, because a line arrives in pieces: the border rasterizer splits
/// every one at the junctions along it, so one 3pt edge is a junction square,
/// a segment and another junction square — three rects at one `(start,
/// thickness)`. How many pieces a line comes in is a property of that
/// decomposition; how many *lines* there are, and how thick each is, is the
/// question here, and collapsing equal runs asks exactly it.
fn edge_runs(pages: &[LayoutedPage]) -> (Vec<Run>, Vec<Run>) {
    let all = rects(pages);
    let distinct = |mut v: Vec<Run>| -> Vec<Run> {
        v.sort_by(|a: &Run, b: &Run| a.0.total_cmp(&b.0).then(a.1.total_cmp(&b.1)));
        v.dedup_by(|a, b| (a.0 - b.0).abs() < 1e-4 && (a.1 - b.1).abs() < 1e-4);
        v
    };
    let horizontal = all
        .iter()
        .filter(|(_, _, w, h)| w > h)
        .map(|&(_, y, _, h)| (y, h))
        .collect();
    let vertical = all
        .iter()
        .filter(|(_, _, w, h)| w <= h)
        .map(|&(x, _, w, _)| (x, w))
        .collect();
    (distinct(horizontal), distinct(vertical))
}

/// A `w:val="double"` edge reaches the page as **two** lines of a third the
/// declared width, a third of it apart — not one line, and not two lines of the
/// full width.
///
/// [MS-OI29500] §17.4.66 treats `w:sz` as the whole border for weight purposes
/// (a double outweighs a single of the same `sz` threefold), and §17.18.2's
/// `double` is a pair of rules with a gap. The two together fix the geometry:
/// the declared 3pt is the pair's total, so each rule is 1pt and the gap is 1pt.
///
/// Asserted against the same table drawn `single` as a control, so the claim is
/// a relation between the two — the pair occupies exactly the band the single
/// fills — rather than four numbers that happen to match the implementation.
/// The control is what makes the whole path load-bearing: it pins that `w:sz`
/// still reached layout as 3pt, so a failure here is about the *style* and not
/// about the unit.
#[test]
fn a_double_border_reaches_the_page_as_two_lines_a_third_of_the_width_apart() {
    let (single_h, single_v) = edge_runs(&layout(&one_cell_table("single")));
    let (double_h, double_v) = edge_runs(&layout(&one_cell_table("double")));

    assert_eq!(
        single_h.len(),
        2,
        "the control's top and bottom edges: {single_h:?}"
    );
    assert_eq!(single_v.len(), 2, "its left and right: {single_v:?}");
    assert!(
        single_h.iter().chain(&single_v).all(|&(_, t)| t == 3.0),
        "w:sz=24 is 3pt; control was {single_h:?} / {single_v:?}"
    );

    assert_eq!(
        double_h.len(),
        4,
        "two lines per horizontal edge: {double_h:?}"
    );
    assert_eq!(double_v.len(), 4, "and per vertical: {double_v:?}");
    assert!(
        double_h.iter().chain(&double_v).all(|&(_, t)| t == 1.0),
        "each line is w:sz / 3; got {double_h:?} / {double_v:?}"
    );

    // Each edge's pair sits at the two ends of the band the single filled: same
    // near side, same far side, a 1pt gap between them.
    for (pair, control, axis) in [
        (&double_h[..2], single_h[0], "top"),
        (&double_h[2..], single_h[1], "bottom"),
        (&double_v[..2], single_v[0], "left"),
        (&double_v[2..], single_v[1], "right"),
    ] {
        assert_eq!(
            pair[0].0, control.0,
            "{axis}: the pair starts where the single does"
        );
        assert_eq!(
            pair[1].0 + pair[1].1,
            control.0 + control.1,
            "{axis}: and ends where it ends"
        );
        assert_eq!(
            pair[1].0 - (pair[0].0 + pair[0].1),
            1.0,
            "{axis}: w:sz / 3 of clear space between the two lines"
        );
    }
}

// ── §17.4.83 `w:vAlign` ─────────────────────────────────────────────────────

/// A row of exactly `twips` height (§17.4.80 `hRule="exact"`), holding one
/// top-, one centre- and one bottom-aligned cell, each with one line of text
/// tagged `{tag}T`, `{tag}C`, `{tag}B`.
fn valign_row(tag: &str, twips: u32) -> String {
    let cells: String = [("T", "top"), ("C", "center"), ("B", "bottom")]
        .iter()
        .map(|(suffix, align)| {
            format!(
                r#"<w:tc>
  <w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="{align}"/></w:tcPr>
  <w:p><w:r><w:t>{tag}{suffix}</w:t></w:r></w:p>
</w:tc>"#
            )
        })
        .collect();
    format!(
        r#"<w:tr>
  <w:trPr><w:trHeight w:val="{twips}" w:hRule="exact"/></w:trPr>
  {cells}
</w:tr>"#
    )
}

/// §17.4.83: `top`, `center` and `bottom` place a cell's content 0,
/// `(row_h − content_h) / 2` and `row_h − content_h` below the row's top edge.
///
/// Neither `row_h` nor `content_h` needs to be known here, and that is the
/// point of the two rows. Within a row, `centre − top` must be exactly half of
/// `bottom − top`, whatever the content height is. Between the two rows, the
/// content is identical, so `content_h` cancels and `bottom − top` must grow by
/// exactly the difference in declared row heights — 2400 − 1200 twips, i.e.
/// 60pt.
///
/// The pair of assertions is what makes this more than a shape check: the first
/// alone would survive any renderer that placed `centre` midway between two
/// wrong extremes, and the second alone would survive one that ignored
/// `content_h` entirely.
#[test]
fn valign_offsets_scale_with_the_row_height_and_centre_is_half_of_bottom() {
    let table = format!(
        r#"<w:tbl>
  <w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
  {}
  {}
</w:tbl>"#,
        valign_row("A", 1200),
        valign_row("B", 2400)
    );
    let pages = layout(&table);

    let offsets = |tag: &str| -> (f32, f32) {
        let top = baseline_of(&pages, &format!("{tag}T"));
        (
            baseline_of(&pages, &format!("{tag}C")) - top,
            baseline_of(&pages, &format!("{tag}B")) - top,
        )
    };
    let (a_centre, a_bottom) = offsets("A");
    let (b_centre, b_bottom) = offsets("B");

    assert!(
        a_bottom > 0.0,
        "the 1200-twip row must have spare height for the alignment to \
         distribute, got {a_bottom}"
    );
    assert_eq!(
        a_centre,
        a_bottom / 2.0,
        "60pt row: centre is half of bottom"
    );
    assert_eq!(
        b_centre,
        b_bottom / 2.0,
        "120pt row: centre is half of bottom"
    );
    assert_eq!(
        b_bottom - a_bottom,
        60.0,
        "the same content in a row 60pt taller drops 60pt further; \
         a={a_bottom}, b={b_bottom}"
    );
}

/// The control, and the half that is easy to lose when fixing the other: a row
/// with no spare height distributes none of it, so all three alignments land in
/// the same place.
///
/// Without this, a renderer that always bottom-aligned would satisfy every
/// relation above — both are differences, and a constant offset cancels out of
/// each.
#[test]
fn valign_moves_nothing_in_a_row_that_is_exactly_its_content() {
    let table = r#"<w:tbl>
  <w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
      <w:p><w:r><w:t>NT</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
      <w:p><w:r><w:t>NC</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>
      <w:p><w:r><w:t>NB</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>"#;
    let pages = layout(table);

    let top = baseline_of(&pages, "NT");
    assert_eq!(
        baseline_of(&pages, "NC"),
        top,
        "centre has nothing to centre"
    );
    assert_eq!(baseline_of(&pages, "NB"), top, "bottom has nothing to drop");
}
