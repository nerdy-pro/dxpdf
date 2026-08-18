//! §17.4.66 / §17.18.2: the square where a vertical table border crosses a
//! horizontal one — who colours it, and what a `double` does to it.
//!
//! **ECMA-376 does not settle either question**, and neither does [MS-OI29500].
//! The standard specifies no stroke geometry for table borders at all;
//! §17.4.66's precedence list is about two *conflicting* declarations on one
//! edge, and a junction is not a conflict — all four segments meeting there are
//! correct and all four want the square. So the rule is this engine's, and until
//! now it was a guess: the `border_precedence` winner among the incident
//! segments, breaking a tie toward the vertical.
//!
//! `test-files/border-junction-colour.docx` is the probe that asks, authored so
//! each candidate reading predicts a *different picture*, and it has now been
//! **measured in Word**. It answers in two halves.
//!
//! Tables 1 and 2 carry 12pt `insideV` and `insideH` of equal weight and style
//! and swap which axis is darker:
//!
//! | reading | table 1 (V dark) | table 2 (H dark) |
//! |---|---|---|
//! | §17.4.66 precedence, darker wins | dark | dark |
//! | the vertical always wins | dark | pale |
//! | **the horizontal wins** | **pale** | **dark** |
//!
//! Word draws pale then dark — so **colour never decides**, and both of the
//! engine's rules died in one render.
//!
//! Table 5 pairs a 12pt vertical with a 3pt horizontal, which is the case those
//! two are silent about, since they tie the weights on purpose. Word draws the
//! vertical through it. So the whole order is **§17.4.66's weight step and
//! nothing after it**: the heavier line takes the square, the horizontal breaks
//! a tie, and style and colour are never consulted.
//!
//! Table 3 crosses two 12pt `double`s, and Word draws the crossing as a 2 × 2
//! lattice of ink with *both* gaps running through it — reported as "the borders
//! are negative space, so it looks like every cell has its own border", which is
//! what a lattice looks like and what a square drawn along one axis never can:
//! with the gaps interrupted at every crossing the picture is a continuous
//! double-ruled grid, not a set of separated per-cell rectangles. That fixes the
//! *geometry* of a square independently of who colours it — it is the product of
//! the two axes' §17.18.2 rules either way.
//!
//! Every assertion below is a *relation between two renders of one document* —
//! a table against the same table with its two colours exchanged, with its two
//! weights exchanged, or drawn `single` instead of `double`. No page origin,
//! glyph metric or absolute coordinate is pinned.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

/// 96 eighths of a point — §17.18.2 caps `w:sz` there, and it is 12pt, which
/// divides by three exactly. A `w:sz` whose third was inexact would put this
/// file's own rounding between it and the geometry it is measuring.
const SZ: &str = "96";
const WIDTH: f32 = 12.0;
const EPS: f32 = 1e-4;

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

/// `test-files/border-junction-colour.docx`'s table, as a document of its own.
///
/// Three columns by three rows with **only** `insideV` and `insideH` declared,
/// so the four crossings of the two interior grid lines with the two interior
/// row boundaries are the only junctions on the page — the table's own four
/// edges carry no border for them to meet.
///
/// `w:trHeight` fixes the rows at 60pt for one reason: it keeps every *segment*
/// off the aspect ratio the junctions are found by. A vertical segment is then
/// 48pt tall and a horizontal one 88pt wide, so nothing but a junction is
/// square, and [`squares`] needs to know no coordinate to find them.
fn junction_table(v_colour: &str, h_colour: &str, style: &str) -> String {
    weighted_table(v_colour, h_colour, style, SZ, SZ)
}

/// [`junction_table`] with a `w:sz` per axis, so the two lines meeting at a
/// crossing can differ in weight.
fn weighted_table(v_colour: &str, h_colour: &str, style: &str, v_sz: &str, h_sz: &str) -> String {
    let cells: String = (0..3)
        .map(|_| {
            r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc>"#.to_string()
        })
        .collect();
    let row = format!(
        r#"<w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="exact"/></w:trPr>{cells}</w:tr>"#
    );
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="6000" w:type="dxa"/>
        <w:tblBorders>
          <w:insideV w:val="{style}" w:sz="{v_sz}" w:space="0" w:color="{v_colour}"/>
          <w:insideH w:val="{style}" w:sz="{h_sz}" w:space="0" w:color="{h_colour}"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      {row}{row}{row}
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>"#
    )
}

fn layout(body: &str) -> Vec<LayoutedPage> {
    let doc = dxpdf::docx::parse(&make_docx(body)).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

type Colour = (u8, u8, u8);
const BLACK: Colour = (0x00, 0x00, 0x00);
const GREY: Colour = (0xBF, 0xBF, 0xBF);

/// The colour painted at `(px, py)`, or `None` for bare paper.
///
/// The **last** rect containing the point, because that is the one the page
/// shows. Containment is strict, and every point sampled below is the centre of
/// a band, so no answer here rests on which side of an edge a boundary falls.
fn ink_at(pages: &[LayoutedPage], (px, py): (f32, f32)) -> Option<Colour> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color } => {
                let (x, y) = (rect.origin.x.raw(), rect.origin.y.raw());
                let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
                (x < px && px < x + w && y < py && py < y + h)
                    .then_some((color.r, color.g, color.b))
            }
            _ => None,
        })
        .next_back()
}

/// The four crossing boxes of [`junction_table`], as `((x0, x1), (y0, y1))`.
///
/// Recovered from the render rather than computed: the vertical grid lines are
/// the x-bands of the tall rects and the row boundaries are the y-bands of the
/// wide ones, and a crossing is one of each. Nothing here knows the page margin,
/// the table's origin or how tall a row came out — only that the two families
/// exist and that a border grid is their product.
///
/// **Bands are merged when they are closer together than they are thick**, which
/// is what makes one grid line one node whatever its style: §17.18.2 draws a
/// `double` as two rules of `w:sz` with exactly `w:sz` between them
/// (`borders::drawn_width`), so its two rules merge and its neighbours 100pt
/// away never can.
///
/// Read off each render separately, and that matters: a `double` reserves three
/// times the strip a `single` does, so the two documents' row boundaries are not
/// at the same y and one render's nodes cannot be reused for the other.
fn crossings(pages: &[LayoutedPage]) -> Vec<((f32, f32), (f32, f32))> {
    let mut vertical: Vec<(f32, f32)> = Vec::new();
    let mut horizontal: Vec<(f32, f32)> = Vec::new();
    for c in pages.iter().flat_map(|p| &p.commands) {
        let DrawCommand::Rect { rect, .. } = c else {
            continue;
        };
        let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
        let (into, lo) = if h > w {
            (&mut vertical, rect.origin.x.raw())
        } else {
            (&mut horizontal, rect.origin.y.raw())
        };
        let span = (lo, lo + w.min(h));
        if !into.iter().any(|&b| (b.0 - span.0).abs() < EPS) {
            into.push(span);
        }
    }

    // One band per grid line: rules closer together than they are thick belong
    // to the same border.
    let merge = |mut v: Vec<(f32, f32)>| -> Vec<(f32, f32)> {
        v.sort_by(|a, b| a.0.total_cmp(&b.0));
        let mut out: Vec<(f32, f32)> = Vec::new();
        for (lo, hi) in v {
            match out.last_mut() {
                Some(last) if lo - last.1 <= (last.1 - last.0).max(hi - lo) + EPS => {
                    last.1 = last.1.max(hi)
                }
                _ => out.push((lo, hi)),
            }
        }
        out
    };
    let (vertical, horizontal) = (merge(vertical), merge(horizontal));
    vertical
        .iter()
        .flat_map(|&vx| horizontal.iter().map(move |&hy| (vx, hy)))
        .collect()
}

/// The centre of a crossing box.
fn centre(b: ((f32, f32), (f32, f32))) -> (f32, f32) {
    ((b.0 .0 + b.0 .1) * 0.5, (b.1 .0 + b.1 .1) * 0.5)
}

/// Word's answer to the probe: **the horizontal takes the crossing**, whichever
/// axis carries the darker line.
///
/// The two tables differ in nothing but which colour is written on which axis,
/// so the four crossings are in the same four places in both and only their
/// colour can move. Asserting *both* is what makes this a discrimination rather
/// than an observation: either one alone is consistent with two of the three
/// readings in the module doc, and only the pair names one.
#[test]
fn at_equal_weight_the_square_takes_the_horizontals_colour() {
    let v_dark = layout(&junction_table("000000", "BFBFBF", "single"));
    let h_dark = layout(&junction_table("BFBFBF", "000000", "single"));

    let nodes = crossings(&v_dark);
    assert_eq!(
        nodes.len(),
        4,
        "two interior grid lines crossing two interior row boundaries: {nodes:?}"
    );
    assert_eq!(
        nodes,
        crossings(&h_dark),
        "exchanging two colours must not move a grid line"
    );
    assert!(
        nodes
            .iter()
            .all(|b| (b.0 .1 - b.0 .0 - WIDTH).abs() < EPS
                && (b.1 .1 - b.1 .0 - WIDTH).abs() < EPS),
        "a crossing of two 12pt singles is 12pt square: {nodes:?}"
    );

    for node in nodes {
        let at = centre(node);
        assert_eq!(
            ink_at(&v_dark, at),
            Some(GREY),
            "the horizontal is the pale line here, so the crossing at {at:?} should be pale"
        );
        assert_eq!(
            ink_at(&h_dark, at),
            Some(BLACK),
            "and dark once the colours are exchanged, at {at:?}"
        );
    }
}

/// …but weight comes first: **the heavier of the two takes the crossing**,
/// and the horizontal only wins when they tie.
///
/// Word's answer to the probe's fifth table, which pairs a 12pt vertical with a
/// 3pt horizontal — the case tables 1 and 2 are silent about, since they tie the
/// two axes on weight on purpose. Word draws the vertical through it.
///
/// Asserted as a pair, one table each way round, and that is what makes it about
/// **weight** rather than about the axis: heavy-vertical gives the vertical's
/// colour and heavy-horizontal the horizontal's, so no single-axis rule and no
/// colour rule survives both. Together with the equal-weight case above, the
/// whole order is: heavier wins, ties go to the horizontal, colour never
/// decides — which is [MS-OI29500] §17.4.66's weight step with the horizontal
/// standing in for the style and colour steps that follow it there.
#[test]
fn the_heavier_of_the_two_borders_takes_the_square() {
    // 12pt against 3pt. Both black-vs-pale, so the crossing's colour names the
    // axis that won it whichever way round the weights are.
    let heavy_v = layout(&weighted_table("000000", "BFBFBF", "single", SZ, "24"));
    let heavy_h = layout(&weighted_table("BFBFBF", "000000", "single", "24", SZ));

    for (pages, want, which) in [
        (&heavy_v, BLACK, "the vertical is the heavy line here"),
        (
            &heavy_h,
            BLACK,
            "and the horizontal is, once they are exchanged",
        ),
    ] {
        let nodes = crossings(pages);
        assert_eq!(nodes.len(), 4, "{which}: crossings {nodes:?}");
        for node in nodes {
            let at = centre(node);
            assert_eq!(
                ink_at(pages, at),
                Some(want),
                "{which}, so the crossing at {at:?} should be its colour"
            );
        }
    }
}

/// §17.18.2: two `double` borders cross as a **2 × 2 lattice**, both gaps
/// running through — not as a square split along one axis.
///
/// Stated as a relation to the same table drawn `single`. Two things change at
/// once and both are measured: the crossing is three times the control's across
/// each axis, because a `double`'s `w:sz` is the width of one of its two rules
/// and not of the pair (`borders::drawn_width`, and
/// `tests/table_geometry_paint.rs` for that half on its own); and within that
/// box the ink is the four corners, so the centre and the four edge midpoints
/// are bare paper.
///
/// The midpoints are what distinguish a lattice from the two other shapes that
/// also leave a hole in the middle: a square split along one axis paints a whole
/// edge of the box, and the union of the two axes' own rules (a `#`) paints all
/// four midpoints.
#[test]
fn two_doubles_cross_as_a_lattice_with_both_gaps_running_through() {
    let single = layout(&junction_table("000000", "000000", "single"));
    let double = layout(&junction_table("000000", "000000", "double"));

    let control = crossings(&single);
    let nodes = crossings(&double);
    assert_eq!(control.len(), 4, "the control's crossings: {control:?}");
    assert_eq!(nodes.len(), 4, "and the double's: {nodes:?}");

    for (node, control) in nodes.iter().zip(&control) {
        let at = centre(*node);
        assert_eq!(
            ink_at(&single, centre(*control)),
            Some(BLACK),
            "the control fills its crossing whole"
        );

        // Same grid line, three times the box: two rules of the declared width
        // with one of it between them, on each axis.
        assert!(
            (at.0 - centre(*control).0).abs() < EPS,
            "the crossing did not move along the grid line: {node:?} vs {control:?}"
        );
        for (side, control_side) in [
            (node.0 .1 - node.0 .0, control.0 .1 - control.0 .0),
            (node.1 .1 - node.1 .0, control.1 .1 - control.1 .0),
        ] {
            assert!(
                (side - control_side * 3.0).abs() < EPS,
                "a double crossing is three times the control's: {side} vs {control_side}"
            );
        }

        // Nine thirds of it: the four corners are ink, the centre and the four
        // edge midpoints are the two gaps crossing.
        let third = (node.0 .1 - node.0 .0) / 3.0;
        for (i, j) in [0usize, 1, 2]
            .into_iter()
            .flat_map(|i| [0, 1, 2].map(|j| (i, j)))
        {
            let p = (
                node.0 .0 + third * (i as f32 + 0.5),
                node.1 .0 + third * (j as f32 + 0.5),
            );
            let corner = i != 1 && j != 1;
            assert_eq!(
                ink_at(&double, p),
                corner.then_some(BLACK),
                "at {p:?} — third ({i}, {j}) of the crossing {node:?}; the corners \
                 are ink and everything else is the two gaps crossing"
            );
        }
    }
}
