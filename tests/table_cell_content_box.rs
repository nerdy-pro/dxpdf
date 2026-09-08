//! §17.4.39/§17.4.66 border widths against §17.4.41/§17.4.42 cell margins —
//! where a cell's content box begins, and that the row is measured from *that*
//! box and not from the cell's outer edge.
//!
//! ECMA-376 gives the two halves separately and never states their interaction:
//! `w:tcMar`/`w:tblCellMar` inset the content from the cell edge, and
//! `w:tcBorders`/`w:tblBorders` give each edge a `w:sz` in eighths of a point,
//! but nothing says which wins when the border is thicker than the margin. This
//! engine draws a border *inside* the cell box (`table::borders::emit_cell_borders`)
//! and so must start the content after it: the inset is `max(border, margin)` on
//! each of the four sides, which is what `(border − margin).max(0)` computes at
//! `table::emit`. That choice is what the tests below pin — and, more to the
//! point, that the *measurement* uses the same rule as the *placement*.
//!
//! It is the disagreement between those two that the reported defect was. The
//! horizontal sides charged the border to the content width in
//! `table::measure`, while the vertical sides shifted the content down at
//! emission without ever growing the row, so a row whose top border was thicker
//! than its top cell margin overflowed its own box by exactly that difference —
//! and the overflow landed in the strip where the bottom border is painted. In
//! `sample-docx-files-sample1.docx` that is the `MediumShading2-Accent5` header:
//! `w:sz="18"` top border, `w:tblCellMar` top of 0, and "Graduating students"
//! wrapping to two lines with its second line drawn through the border.
//!
//! Every assertion is a *difference between two renders of the same document*,
//! so no glyph metric, cell margin or page origin has to be known: whatever
//! those contribute they contribute equally to both sides and cancel. The
//! declared border widths are the only absolute numbers, and they are chosen so
//! that eighths of a point divide exactly.

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

fn layout(body: &str) -> Vec<LayoutedPage> {
    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>"#
    );
    let doc = dxpdf::docx::parse(&make_docx(&document_xml)).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// A fill no border or theme colour can collide with, so the cell's shading
/// rect — which is emitted at the row box exactly, `(cell_x, row_top, cell_w,
/// row_height)` — can be picked out of the page by colour alone. That rect is
/// how these tests read the row box without a layout-internal API.
const FILL: (u8, u8, u8) = (0xC0, 0xFF, 0xEE);

/// The shaded cell's box as `(top, height)`.
fn row_box(pages: &[LayoutedPage]) -> (f32, f32) {
    let mut found: Vec<(f32, f32)> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color } if (color.r, color.g, color.b) == FILL => {
                Some((rect.origin.y.raw(), rect.size.height.raw()))
            }
            _ => None,
        })
        .collect();
    assert_eq!(found.len(), 1, "expected one shaded cell, got {found:?}");
    found.pop().unwrap()
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

/// A one-row, one-cell shaded table carrying a border on `edge` only, of `sz`
/// eighths of a point, with every cell margin set to `mar` twips.
///
/// Every other edge is `w:val="nil"` rather than omitted, so the table style
/// cascade cannot put a border back and change the answer: §17.4.66 `nil`
/// declines inheritance where an omitted edge accepts it.
fn one_bordered_edge(edge: &str, sz: u32, mar: u32) -> String {
    let edges: String = ["top", "left", "bottom", "right", "insideH", "insideV"]
        .iter()
        .map(|e| {
            if *e == edge {
                format!(r#"<w:{e} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>"#)
            } else {
                format!(r#"<w:{e} w:val="nil"/>"#)
            }
        })
        .collect();
    let margins: String = ["top", "left", "bottom", "right"]
        .iter()
        .map(|e| format!(r#"<w:{e} w:w="{mar}" w:type="dxa"/>"#))
        .collect();
    format!(
        r#"<w:tbl>
  <w:tblPr>
    <w:tblW w:w="4000" w:type="dxa"/>
    <w:tblBorders>{edges}</w:tblBorders>
    <w:tblCellMar>{margins}</w:tblCellMar>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
  <w:tr><w:tc>
    <w:tcPr>
      <w:tcW w:w="4000" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="C0FFEE"/>
    </w:tcPr>
    <w:p><w:r><w:t>CELL</w:t></w:r></w:p>
  </w:tc></w:tr>
</w:tbl>"#
    )
}

/// `w:sz="4"` is 0.5pt and `w:sz="48"` is 6pt, so the pair differs by 5.5pt —
/// the one absolute number these tests assert, and exact in binary.
const THIN: u32 = 4;
const THICK: u32 = 48;
const DIFFERENCE: f32 = 5.5;

// ── the top edge ────────────────────────────────────────────────────────────

/// A top border thicker than the top cell margin pushes the content down, so
/// the row must grow by the same amount.
///
/// This is the defect itself. `table::emit` shifted the content by
/// `(border − margin).max(0)` while `table::measure` sized the row as content
/// plus margins alone, so the row's height did not depend on its top border at
/// all and this difference was 0.
#[test]
fn a_thicker_top_border_makes_the_row_taller_by_the_difference() {
    let (_, thin) = row_box(&layout(&one_bordered_edge("top", THIN, 0)));
    let (_, thick) = row_box(&layout(&one_bordered_edge("top", THICK, 0)));

    assert_eq!(
        thick - thin,
        DIFFERENCE,
        "w:sz {THIN}→{THICK} is {DIFFERENCE}pt of extra top inset, so the row \
         box must be {DIFFERENCE}pt taller; got {thin} → {thick}"
    );
}

/// …and the content must still end where the box ends.
///
/// The companion to the test above, and the half that says *why* it matters:
/// the distance from the last baseline to the foot of the box is a property of
/// the content, so it cannot change when only the border does. When the row did
/// not grow, this shrank by the full 5.5pt and the content ran out through the
/// bottom of its own cell.
///
/// Asserted alongside the baseline having actually moved, because a renderer
/// that ignored the border on both sides at once would keep the first equality
/// while placing the content wrong.
#[test]
fn the_content_stays_inside_the_box_when_the_top_border_grows() {
    let thin_pages = layout(&one_bordered_edge("top", THIN, 0));
    let thick_pages = layout(&one_bordered_edge("top", THICK, 0));

    let (thin_top, thin_h) = row_box(&thin_pages);
    let (thick_top, thick_h) = row_box(&thick_pages);
    let thin_base = baseline_of(&thin_pages, "CELL");
    let thick_base = baseline_of(&thick_pages, "CELL");

    assert_eq!(
        thick_base - thin_base,
        DIFFERENCE,
        "the thicker border must push the content down by its extra {DIFFERENCE}pt"
    );
    assert_eq!(
        (thick_top + thick_h) - thick_base,
        (thin_top + thin_h) - thin_base,
        "the same content must sit the same distance above the foot of its box, \
         whatever the border above it is"
    );
}

/// The control, and the half that is easy to lose when fixing the other: a top
/// cell margin at least as thick as the border absorbs it entirely, so the row
/// does not grow at all.
///
/// §17.4.42's margin is measured from the cell edge, not from the inside of the
/// border, so a 6pt margin already clears a 6pt border. Without this, a fix that
/// simply added the border width to the row height would pass every assertion
/// above.
#[test]
fn a_top_margin_as_thick_as_the_border_absorbs_it() {
    // 120 twips is 6pt — exactly `THICK`.
    let (_, thin) = row_box(&layout(&one_bordered_edge("top", THIN, 120)));
    let (_, thick) = row_box(&layout(&one_bordered_edge("top", THICK, 120)));

    assert_eq!(
        thick, thin,
        "max(border, margin) is 6pt either way, so the row height cannot differ"
    );
}

// ── the bottom edge ─────────────────────────────────────────────────────────

/// The same rule on the other side. A table's last row has no strip reserved
/// below it, so its bottom border is drawn *inside* the box
/// (`emit_cell_borders`' `cell.y + cell.h - bot_w` branch) and the content must
/// stop above it.
#[test]
fn a_thicker_bottom_border_makes_the_row_taller_by_the_difference() {
    let (_, thin) = row_box(&layout(&one_bordered_edge("bottom", THIN, 0)));
    let (_, thick) = row_box(&layout(&one_bordered_edge("bottom", THICK, 0)));

    assert_eq!(thick - thin, DIFFERENCE, "got {thin} → {thick}");
}

/// …and, unlike the top edge, the content must *not* move: a bottom border
/// grows the box downward, under content that was already correctly placed.
///
/// The pair is what distinguishes the two edges. A fix that reserved the inset
/// on the wrong side would satisfy the height assertion above and fail here.
#[test]
fn the_content_does_not_move_when_the_bottom_border_grows() {
    let thin_pages = layout(&one_bordered_edge("bottom", THIN, 0));
    let thick_pages = layout(&one_bordered_edge("bottom", THICK, 0));

    let (thin_top, _) = row_box(&thin_pages);
    let (thick_top, _) = row_box(&thick_pages);

    assert_eq!(
        baseline_of(&thick_pages, "CELL") - thick_top,
        baseline_of(&thin_pages, "CELL") - thin_top,
        "a bottom border is below the content and must not displace it"
    );
}

// ── §17.4.83 `w:vAlign` against the same insets ─────────────────────────────

/// A bottom-aligned cell drops its content to the foot of the **content box**,
/// not of the cell — so a thicker bottom border pushes it back up.
///
/// The half of the fix that is invisible from the top-aligned tests above. Emit
/// distributes `row_h − content_h` as vAlign slack, and if the two insets were
/// not subtracted from that slack first, `bottom` would put the content right
/// back under the border the row had just made room for — cancelling the height
/// fix exactly where it matters most.
///
/// The row is `hRule="exact"` so its height is fixed by the author rather than
/// by the content, which is what leaves slack to distribute at all: without it
/// the row is its own content and every alignment lands in the same place.
#[test]
fn a_bottom_aligned_cell_sits_above_its_bottom_border() {
    let cell = |sz: u32| {
        format!(
            r#"<w:tbl>
  <w:tblPr>
    <w:tblW w:w="4000" w:type="dxa"/>
    <w:tblBorders>
      <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>
      <w:top w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/>
      <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
  <w:tr>
    <w:trPr><w:trHeight w:val="2400" w:hRule="exact"/></w:trPr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="4000" w:type="dxa"/>
        <w:vAlign w:val="bottom"/>
        <w:shd w:val="clear" w:color="auto" w:fill="C0FFEE"/>
      </w:tcPr>
      <w:p><w:r><w:t>CELL</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>"#
        )
    };

    let thin_pages = layout(&cell(THIN));
    let thick_pages = layout(&cell(THICK));

    // `hRule="exact"` pins the box, so both rows are the same height and only
    // the border inside them differs — the content must move by the difference.
    let (thin_top, thin_h) = row_box(&thin_pages);
    let (thick_top, thick_h) = row_box(&thick_pages);
    assert_eq!(
        thin_h, thick_h,
        "an exact row height is the author's, whatever the border"
    );

    assert_eq!(
        (thin_top + thin_h - baseline_of(&thin_pages, "CELL"))
            - (thick_top + thick_h - baseline_of(&thick_pages, "CELL")),
        -DIFFERENCE,
        "the thicker bottom border must lift the bottom-aligned content by its \
         extra {DIFFERENCE}pt, not leave it under the border"
    );
}

// ── the horizontal sides, which already agreed ──────────────────────────────

/// The regression guard for the side that was never broken: a left border
/// thicker than the left cell margin narrows the content, so a cell holding a
/// word that only just fits wraps.
///
/// `table::measure` subtracts `(border − margin).max(0)` from the layout width
/// before laying the cell out, which is exactly the rule the vertical sides were
/// missing. Asserted as a line count rather than a width, so it survives
/// whatever the host's font measures `WWWWWWWW` at.
#[test]
fn a_thicker_left_border_narrows_the_content_box() {
    let cell = |sz: u32| {
        format!(
            r#"<w:tbl>
  <w:tblPr>
    <w:tblW w:w="900" w:type="dxa"/>
    <w:tblBorders>
      <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>
      <w:top w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
      <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="900"/></w:tblGrid>
  <w:tr><w:tc>
    <w:tcPr>
      <w:tcW w:w="900" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="C0FFEE"/>
    </w:tcPr>
    <w:p><w:r><w:t>MMMMMM MMMMMM</w:t></w:r></w:p>
  </w:tc></w:tr>
</w:tbl>"#
        )
    };

    // 900 twips is 45pt; a 40pt left border leaves 5pt of content box, which no
    // host font fits `MMMMMM` into on one line.
    let (_, narrow) = row_box(&layout(&cell(320)));
    let (_, wide) = row_box(&layout(&cell(0)));

    assert!(
        narrow > wide,
        "a 40pt left border must narrow the content box enough to add lines; \
         got {wide} unbordered vs {narrow} bordered"
    );
}

// ── a shared vertical: how much of it is inside each cell ───────────────────

/// A one-row, two-cell table whose **only** border is the `insideV` the two
/// cells share, at `sz` eighths of a point, with zero cell margins.
///
/// Zero margins because the inset is `max(border, margin)` per side, and Word's
/// own `TableNormal` margin of 108 twips would mask any border under 5.4pt —
/// which is most of the range this measures over. Every other edge is
/// `w:val="nil"` for the reason `one_bordered_edge` gives: an omitted edge
/// inherits, a `nil` one does not.
fn one_shared_vertical(sz: u32) -> String {
    let edges: String = ["top", "left", "bottom", "right", "insideH"]
        .iter()
        .map(|e| format!(r#"<w:{e} w:val="nil"/>"#))
        .collect();
    let cell = |label: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
    <w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:tc>"#
        )
    };
    format!(
        r#"<w:tbl>
  <w:tblPr>
    <w:tblW w:w="4000" w:type="dxa"/>
    <w:tblBorders>{edges}<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/></w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>
    </w:tblCellMar>
    <w:tblLayout w:type="fixed"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
  <w:tr>{}{}</w:tr>
</w:tbl>"#,
        cell("LEFT"),
        cell("RIGHT")
    )
}

/// The x of the text command whose content is exactly `needle`.
fn glyph_x(pages: &[LayoutedPage], needle: &str) -> f32 {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if &**text == needle => Some(position.x.raw()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text {needle:?} on the page"))
}

/// §17.4.66: a cell is charged **half** of the border it shares with its
/// neighbour, not all of it and not none of it.
///
/// **Measured in Word** against `test-files/border-content-charge.docx`, which
/// steps a shared border 0.5 → 12pt across four rows with zero cell margins and
/// asks where the following cell's glyph lands. Word draws it flush against the
/// border's inner edge at every weight — so half the line lies in each cell,
/// which is the collapsed model and is what `rasterize_border_grid` already
/// paints.
///
/// Two readings are refuted, and the engine held both of them at once. The
/// second cell used to be charged **nothing**: §17.4.66 resolution handed the
/// shared edge to the cell on its left and cleared the facing `left`, so the
/// right-hand cell's content began on the grid line and a thick border was
/// painted straight through its first glyph. That is the picture the probe was
/// built to rule out, and Word does not draw it. Charging the full width to each
/// side is refuted by the same render, from the other direction: it would leave
/// a gap of half the border between the line and the glyph, and there is none.
///
/// Asserted as the difference between two renders of one document, so no glyph
/// metric or page origin is pinned. The two borders differ by 5.5pt, and each
/// cell should take half of that.
#[test]
fn a_shared_vertical_is_charged_half_to_each_of_the_cells_that_meet_on_it() {
    let thin = layout(&one_shared_vertical(THIN));
    let thick = layout(&one_shared_vertical(THICK));

    // The right-hand cell begins after half the border. It used to begin after
    // none of it — the defect.
    assert_eq!(
        glyph_x(&thick, "RIGHT") - glyph_x(&thin, "RIGHT"),
        DIFFERENCE / 2.0,
        "the following cell is charged half the shared border, not none of it"
    );

    // The control, and the half that says the charge is *shared*: the left-hand
    // cell's content starts at the table's own edge, which no `insideV` touches,
    // so its glyph must not move at all. A fix that charged the whole border to
    // the right-hand cell would satisfy the first assertion by moving this one
    // too, if the two were ever confused.
    assert_eq!(
        glyph_x(&thick, "LEFT"),
        glyph_x(&thin, "LEFT"),
        "the leading cell's content box does not start on the shared edge"
    );
}
