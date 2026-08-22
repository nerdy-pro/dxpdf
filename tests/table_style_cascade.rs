//! §17.7.2 / §17.7.6: what a **table style** declares, and what of it reaches
//! the table.
//!
//! The cascade a table sees has three levels — direct `<w:tblPr>` on the
//! `<w:tbl>`, the style it names, and that style's `basedOn` ancestors — and
//! every one of them speaks the same vocabulary (`CT_TblPrBase`). Most tests
//! here are written as *parity*: a property declared in the style must produce
//! the same page as the same property declared directly on the table, and both
//! must differ from a control that declares it nowhere. That formulation is the
//! spec statement itself ("the current style inherits all of the properties of
//! the base style", §17.7.4.3) rather than a transcription of current output,
//! and it stays true if the geometry these properties drive is ever refined.
//!
//! # The vocabulary is not shared in full
//!
//! `CT_TblPrBase` is one content model, but Word does not read all of it from a
//! style. [MS-OI29500] §2.1.250(a) (on §17.7.6.4, a style's own `tblPr`) and
//! §2.1.249(a) (on §17.7.6.3, a conditional one) each list the elements the
//! standard allows there and Word does not, so those cases are written as the
//! *inverse* parity — declaring it in the style must change nothing, while the
//! same element on the `<w:tbl>` still applies. `build_table` states which
//! elements and why; the split is:
//!
//! | reaches the table from a style | does not (§2.1.250(a))         |
//! |--------------------------------|--------------------------------|
//! | `jc`, `tblInd`, `tblBorders`   | `tblW`, `tblLook`, `tblpPr`    |
//! | `tblCellMar`, `tblCellSpacing` | `tblOverlap`, `tblLayout`      |
//! | the two band sizes             | `bidiVisual`, `tblStyle`       |
//!
//! No document in `test-files/` exercises a table style that declares anything
//! other than borders, cell margins, indent, the band sizes and `tblLayout`, so
//! the fixtures are built here: the XML *is* the point of each test, and a
//! `.docx` would hide it.

use std::io::Write;

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

fn make_docx(document_xml: &str, styles_xml: &str) -> Vec<u8> {
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
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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

        zip.start_file("word/_rels/document.xml.rels", o).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/styles.xml", o).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        zip.start_file("word/document.xml", o).unwrap();
        zip.write_all(document_xml.as_bytes()).unwrap();
        zip.finish().unwrap();
    }
    buf
}

/// A 4×2 table whose only *direct* table property is the style reference, plus
/// whatever `direct_tbl_pr` adds. Four rows and two columns is the smallest
/// grid that still distinguishes a row band size of 1 from 2 and a column band
/// size of 1 from 2.
///
/// The page geometry is stated explicitly rather than defaulted so the expected
/// positions are readable off the fixture: 12240 − 2×1440 twips of content is
/// 468 pt wide, starting at x = 72 pt, and the 2×2000-twip grid makes the table
/// 200 pt wide.
fn table_document(direct_tbl_pr: &str) -> String {
    let cell = |t: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                 <w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:tc>"#
        )
    };
    let row = |a: &str, b: &str| format!("<w:tr>{}{}</w:tr>", cell(a), cell(b));
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TestTbl"/>{direct_tbl_pr}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
      {}{}{}{}
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"#,
        row("A", "B"),
        row("C", "D"),
        row("E", "F"),
        row("G", "H"),
    )
}

/// Two copies of [`table_document`]'s table, back to back, so that `tblOverlap`
/// has a second float to collide with — §17.4.56 governs table-vs-table
/// overlap and says nothing about a lone float.
fn two_tables_document(direct_tbl_pr: &str) -> String {
    let one = table_document(direct_tbl_pr);
    // Splice a second `<w:tbl>` in ahead of the section properties.
    let (body, tail) = one.split_once("<w:sectPr>").expect("fixture has a sectPr");
    let table = {
        let start = body.find("<w:tbl>").expect("fixture has a table");
        &body[start..]
    };
    format!("{body}{table}<w:sectPr>{tail}")
}

/// A stylesheet with one table style, `TestTbl`, whose body is `style_body`.
fn styles_with(style_body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="TestTbl">
    <w:name w:val="Test Table"/>
    {style_body}
  </w:style>
</w:styles>"#
    )
}

/// A stylesheet where `TestTbl` — the style every fixture table names — is
/// `basedOn` a second style that carries `parent_body`.
fn styles_derived_from(parent_body: &str, child_body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="ParentTbl">
    <w:name w:val="Parent Table"/>
    {parent_body}
  </w:style>
  <w:style w:type="table" w:styleId="TestTbl">
    <w:name w:val="Test Table"/>
    <w:basedOn w:val="ParentTbl"/>
    {child_body}
  </w:style>
</w:styles>"#
    )
}

fn layout(document_xml: &str, styles_xml: &str) -> Vec<LayoutedPage> {
    let doc = dxpdf::docx::parse(&make_docx(document_xml, styles_xml)).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

/// Every drawn thing, as a stable string — positions to two decimals, colors in
/// hex. Comparing these compares the *page*, so a property that fails to reach
/// layout shows up whatever it would have moved.
fn page_geometry(pages: &[LayoutedPage]) -> Vec<String> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text {
                position,
                text,
                color,
                ..
            } => Some(format!(
                "text {text:?} @ {:.2},{:.2} #{:02X}{:02X}{:02X}",
                position.x.raw(),
                position.y.raw(),
                color.r,
                color.g,
                color.b
            )),
            DrawCommand::Rect { rect, color, .. } => Some(format!(
                "rect {:.2},{:.2} {:.2}x{:.2} #{:02X}{:02X}{:02X}",
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.size.width.raw(),
                rect.size.height.raw(),
                color.r,
                color.g,
                color.b
            )),
            _ => None,
        })
        .collect()
}

/// x of every text run, in draw order.
fn text_xs(pages: &[LayoutedPage]) -> Vec<f32> {
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Text { position, .. } => Some(position.x.raw()),
            _ => None,
        })
        .collect()
}

/// The three renders every case below is decided by: the property written
/// nowhere, written directly on the `<w:tbl>`, and written in the table style.
struct ThreeWays {
    control: Vec<String>,
    direct: Vec<String>,
    from_style: Vec<String>,
}

/// `style_extra` is appended to the style body after its `<w:tblPr>`, which is
/// where a `tblLook`/band-size case puts the `tblStylePr` layers it needs to be
/// observable at all. `constant_tbl_pr` is written directly on the table in
/// *all three* variants — the background a property needs in order to be
/// observable at all (`tblOverlap` says nothing about a table that does not
/// float).
fn three_ways(
    document: &dyn Fn(&str) -> String,
    constant_tbl_pr: &str,
    tbl_pr: &str,
    style_extra: &str,
) -> ThreeWays {
    ThreeWays {
        control: page_geometry(&layout(
            &document(constant_tbl_pr),
            &styles_with(&format!("<w:tblPr/>{style_extra}")),
        )),
        direct: page_geometry(&layout(
            &document(&format!("{constant_tbl_pr}{tbl_pr}")),
            &styles_with(&format!("<w:tblPr/>{style_extra}")),
        )),
        from_style: page_geometry(&layout(
            &document(constant_tbl_pr),
            &styles_with(&format!("<w:tblPr>{tbl_pr}</w:tblPr>{style_extra}")),
        )),
    }
}

/// §17.7.2: the whole point of a style. A property in the style must render
/// exactly as the same property written directly on the `<w:tbl>` — and the
/// control, which writes it in neither place, must differ, or the assertion
/// above it proves nothing.
#[track_caller]
fn assert_style_matches_direct(tbl_pr: &str, style_extra: &str, what: &str) {
    assert_style_matches_direct_with(&table_document, "", tbl_pr, style_extra, what);
}

/// [`assert_style_matches_direct`] over a different fixture.
#[track_caller]
fn assert_style_matches_direct_with(
    document: &dyn Fn(&str) -> String,
    constant_tbl_pr: &str,
    tbl_pr: &str,
    style_extra: &str,
    what: &str,
) {
    let w = three_ways(document, constant_tbl_pr, tbl_pr, style_extra);
    assert_ne!(
        w.control, w.direct,
        "{what}: the fixture does not discriminate — writing it directly on the \
         table changes nothing, so the style assertion below is vacuous"
    );
    assert_eq!(
        w.from_style, w.direct,
        "{what}: declared in the table style, it must render as if declared \
         directly on the table"
    );
}

/// The inverse, for the properties [MS-OI29500] §2.1.250(a) says Word does not
/// accept in a table style's `<w:tblPr>`: declaring one there must change
/// **nothing**, so the style render matches the control rather than the direct
/// one.
///
/// The same `control != direct` guard runs first, and does the same job in
/// the opposite direction: without it, "declaring it in the style changed
/// nothing" would also be satisfied by a fixture in which the property changes
/// nothing anywhere.
#[track_caller]
fn assert_style_does_not_reach_the_table(tbl_pr: &str, style_extra: &str, what: &str) {
    assert_style_does_not_reach_the_table_with(&table_document, "", tbl_pr, style_extra, what);
}

/// [`assert_style_does_not_reach_the_table`] over a different fixture.
#[track_caller]
fn assert_style_does_not_reach_the_table_with(
    document: &dyn Fn(&str) -> String,
    constant_tbl_pr: &str,
    tbl_pr: &str,
    style_extra: &str,
    what: &str,
) {
    let w = three_ways(document, constant_tbl_pr, tbl_pr, style_extra);
    assert_ne!(
        w.control, w.direct,
        "{what}: the fixture does not discriminate — writing it directly on the \
         table changes nothing, so the style assertion below is vacuous"
    );
    assert_eq!(
        w.from_style, w.control,
        "{what}: [MS-OI29500] §2.1.250(a) — Word does not accept this element \
         in a table style's tblPr, so declaring it there must change nothing"
    );
}

// ── §17.7.2: each `CT_TblPrBase` property, declared in the style ────────────

/// §17.4.28 `jc` — the table's horizontal alignment within the content area.
#[test]
fn a_table_style_can_align_the_table() {
    assert_style_matches_direct(r#"<w:jc w:val="center"/>"#, "", "jc");

    // The absolute placement, so a regression that moves *both* sides of the
    // parity together still fails: 72 pt left margin + half of (468 − 200).
    let centered = layout(
        &table_document(""),
        &styles_with(r#"<w:tblPr><w:jc w:val="center"/></w:tblPr>"#),
    );
    assert_eq!(
        text_xs(&centered).first().copied(),
        Some(206.0),
        "a 200 pt table centered in a 468 pt content area starts at 72 + 134"
    );
}

/// §17.4.63 `tblW` — and the first of the six the erratum removes.
///
/// [MS-OI29500] §2.1.250(a), on §17.7.6.4's `tblPr` — a table style's own —
/// says the standard permits `bidiVisual`, `tblLayout`, `tblLook`,
/// `tblOverlap`, `tblpPr`, `tblStyle` and `tblW` as its children and "Word
/// does not allow these elements to be child elements of the tblPr element".
/// §2.1.249(a) says the same of §17.7.6.3's conditional `tblPr`, adding the
/// two band sizes. A document that declares `tblW` in a style therefore either
/// fails to open in Word or renders without it, and matching Word is what this
/// engine is for.
#[test]
fn a_table_style_cannot_set_the_table_width() {
    assert_style_does_not_reach_the_table(r#"<w:tblW w:w="7200" w:type="dxa"/>"#, "", "tblW");

    // The absolute placement, so a regression that moves *both* sides of the
    // parity together still fails: the grid's own 2×2000 twips, unscaled.
    let from_style = layout(
        &table_document(""),
        &styles_with(r#"<w:tblPr><w:tblW w:w="7200" w:type="dxa"/></w:tblPr>"#),
    );
    assert_eq!(
        text_xs(&from_style).get(1).copied(),
        Some(172.0),
        "the style's 7200 twips is ignored, leaving the second column at its \
         declared 2000 twips = 100 pt past the 72 pt margin"
    );

    // …and the direct level is untouched: 7200 twips = 360 pt still scales the
    // two equal grid columns to 180 pt each.
    let direct = layout(
        &table_document(r#"<w:tblW w:w="7200" w:type="dxa"/>"#),
        &styles_with("<w:tblPr/>"),
    );
    assert_eq!(
        text_xs(&direct).get(1).copied(),
        Some(252.0),
        "written on the <w:tbl> it still applies — the removal is of the style \
         level only"
    );
}

/// §17.4.50 `tblInd` — indentation from the leading margin.
#[test]
fn a_table_style_can_indent_the_table() {
    assert_style_matches_direct(r#"<w:tblInd w:w="720" w:type="dxa"/>"#, "", "tblInd");

    let indented = layout(
        &table_document(""),
        &styles_with(r#"<w:tblPr><w:tblInd w:w="720" w:type="dxa"/></w:tblPr>"#),
    );
    assert_eq!(
        text_xs(&indented).first().copied(),
        Some(108.0),
        "720 twips = 36 pt past the 72 pt left margin"
    );
}

/// §17.4.45 `tblCellSpacing` — the gap carved out between and around cells.
#[test]
fn a_table_style_can_set_cell_spacing() {
    assert_style_matches_direct(
        r#"<w:tblCellSpacing w:w="144" w:type="dxa"/>"#,
        "",
        "tblCellSpacing",
    );
}

/// §17.4.55 `tblLook` — also on §2.1.250(a)'s list, and the one whose exclusion
/// the element's own title argues for independently: "Table Style Conditional
/// Formatting **Settings**" is the table's statement about which of the
/// referenced style's regions it wants, so it belongs to the reference and not
/// to the style. A style switching off its own regions could simply not define
/// them.
///
/// Only observable through a `tblStylePr` layer, so the style carries one.
#[test]
fn a_table_style_cannot_set_the_tbl_look() {
    let first_row_red = r#"<w:tblStylePr w:type="firstRow">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
           </w:tblStylePr>"#;
    let all_regions_off = r#"<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0"
                                        w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>"#;
    assert_style_does_not_reach_the_table(all_regions_off, first_row_red, "tblLook");

    let from_style = layout(
        &table_document(""),
        &styles_with(&format!(
            "<w:tblPr>{all_regions_off}</w:tblPr>{first_row_red}"
        )),
    );
    assert_eq!(
        red_cells(&from_style),
        2,
        "firstRow=0 in the style's own tblLook is ignored, so §17.4.55's \
         absent-element default (0x04A0, firstRow on) still shades the row"
    );

    let direct = layout(
        &table_document(all_regions_off),
        &styles_with(first_row_red),
    );
    assert_eq!(
        red_cells(&direct),
        0,
        "…while the same element on the <w:tbl> still switches the layer off"
    );
}

/// A `firstRow` layer painted red, and how many cells it painted — the probe
/// every `tblLook` case reads, since `tblLook` has no geometry of its own.
/// `firstRow` because §17.4.55 note (a)'s absent-element default is Word's
/// 0x04A0, which switches that region **on** and `lastRow` off, so the layer
/// paints unless something says otherwise.
fn red_cells(pages: &[LayoutedPage]) -> usize {
    cells_shaded(pages, (0xFF, 0x00, 0x00))
}

/// How many of the table's cells carry a `color` background.
///
/// Counted by asking, for each cell's *own* text, whether its baseline lies
/// inside a rect of that colour — not by counting the rects. §17.4.33 shading
/// is a cell property and reaches the page as one rect per cell, but only until
/// `coalesce_abutting_rects` fuses abutting same-colour neighbours into one
/// (`tests/table_shading_seams.rs` says why it must). After that a rect count
/// answers how many *runs* were painted rather than how many cells, and a pair
/// of shaded neighbours counts as one.
///
/// Every cell of `table_document` holds a distinct one-letter label, and a
/// label is inside its own cell by construction, so this counts cells however
/// the fills happened to be emitted — and asserts something a rect count never
/// did: that the shading actually covers the cell whose text it is behind.
fn cells_shaded(pages: &[LayoutedPage], color: (u8, u8, u8)) -> usize {
    let rects: Vec<_> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color: c } if (c.r, c.g, c.b) == color => Some((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.origin.x.raw() + rect.size.width.raw(),
                rect.origin.y.raw() + rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect();
    pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter(|c| {
            let DrawCommand::Text { text, position, .. } = c else {
                return false;
            };
            if text.trim().is_empty() {
                return false;
            }
            let (x, y) = (position.x.raw(), position.y.raw());
            rects
                .iter()
                .any(|&(x0, y0, x1, y1)| x >= x0 && x <= x1 && y >= y0 && y <= y1)
        })
        .count()
}

/// The two fills every conditional-layer fixture paints with.
const RED: (u8, u8, u8) = (0xFF, 0x00, 0x00);
const GREEN: (u8, u8, u8) = (0x00, 0xFF, 0x00);

const FIRST_ROW_RED: &str = r#"<w:tblStylePr w:type="firstRow">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
           </w:tblStylePr>"#;

/// §17.4.55: an **empty** `<w:tblLook/>` states nothing, so it resolves to the
/// absent element's default — not to "every region off".
///
/// `CT_TblLook`'s six attributes are optional `ST_OnOff` with no schema
/// default, and `@val` is optional too, so `<w:tblLook/>` carries exactly what
/// the omitted element carries: nothing. The parse seam is where that is
/// decided — `Dup::filter_map(tbl_look)` drops the occurrence rather than
/// producing an all-`None` value — and `tbl_pr_tbl_look_empty_element_states_
/// nothing` pins it there. This is the render end of the same rule, and the
/// arm that discriminates is the third: an element that *does* state
/// something is still a value, so the drop is narrowed to silence and is not
/// "a direct `tblLook` never counts".
#[test]
fn an_empty_tbl_look_is_the_absent_elements_default_not_every_region_off() {
    let styles = styles_with(FIRST_ROW_RED);

    let absent = layout(&table_document(""), &styles);
    assert_eq!(
        red_cells(&absent),
        2,
        "0x04A0 leaves firstRow active, so the layer shades both cells of the \
         first row — without which the assertions below are vacuous"
    );

    let empty = layout(&table_document("<w:tblLook/>"), &styles);
    assert_eq!(
        red_cells(&empty),
        2,
        "an empty <w:tblLook/> states nothing, so the same default applies"
    );

    let stated_off = layout(&table_document(r#"<w:tblLook w:val="0000"/>"#), &styles);
    assert_eq!(
        red_cells(&stated_off),
        0,
        "…while a tblLook that states every region off is a value, and is read"
    );
}

/// §17.7.6.7 `tblStyleRowBandSize` — how many rows one horizontal band spans.
///
/// One of the two properties that **stay** on the style level, and the evidence
/// runs the opposite way from `tblW`'s: §2.1.250(a)'s list omits the band
/// sizes, while [MS-OI29500] §2.1.164(a) — on §17.4.59, the `<w:tbl>`'s own
/// `tblPr` — says Word does *not* allow them there. So the style is the level
/// Word reads them at, which is also where ECMA documents them (§17.7.6.5 and
/// §17.7.6.7, under Table Styles, not §17.4) and where every one of the 693
/// band-size declarations in this repo's corpus sits. None is on a `<w:tbl>`.
#[test]
fn a_table_style_can_set_the_row_band_size() {
    let bands = r#"<w:tblStylePr w:type="band1Horz">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
           </w:tblStylePr>
           <w:tblStylePr w:type="band2Horz">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:tcPr>
           </w:tblStylePr>"#;
    // firstRow/lastRow off so all four rows band, noVBand so only rows band.
    // Written directly on the `<w:tbl>` in every variant, as the constant
    // background: `tblLook` is the switch that makes banding observable at all,
    // and it is one of the elements a style may not carry, so putting it in the
    // style would test its exclusion rather than the band size.
    let look = r#"<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0"
                             w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>"#;
    assert_style_matches_direct_with(
        &table_document,
        look,
        r#"<w:tblStyleRowBandSize w:val="2"/>"#,
        bands,
        "tblStyleRowBandSize",
    );

    let banded = layout(
        &table_document(look),
        &styles_with(&format!(
            r#"<w:tblPr><w:tblStyleRowBandSize w:val="2"/></w:tblPr>{bands}"#
        )),
    );
    let reds = cells_shaded(&banded, RED);
    assert_eq!(
        reds, 4,
        "a band size of 2 puts rows 0-1 (four cells) in band1, not rows 0 and 2"
    );
}

/// §17.7.6.5 `tblStyleColBandSize` — how many columns one vertical band spans.
/// The same evidence as the row band size above.
#[test]
fn a_table_style_can_set_the_column_band_size() {
    let bands = r#"<w:tblStylePr w:type="band1Vert">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
           </w:tblStylePr>
           <w:tblStylePr w:type="band2Vert">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:tcPr>
           </w:tblStylePr>"#;
    // noHBand so row banding cannot override column banding (§17.7.6 order);
    // direct, for the reason the row-band case states.
    let look = r#"<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0"
                             w:lastColumn="0" w:noHBand="1" w:noVBand="0"/>"#;
    assert_style_matches_direct_with(
        &table_document,
        look,
        r#"<w:tblStyleColBandSize w:val="2"/>"#,
        bands,
        "tblStyleColBandSize",
    );

    let banded = layout(
        &table_document(look),
        &styles_with(&format!(
            r#"<w:tblPr><w:tblStyleColBandSize w:val="2"/></w:tblPr>{bands}"#
        )),
    );
    assert_eq!(
        cells_shaded(&banded, RED),
        8,
        "a band size of 2 puts both columns in band1 — all eight cells"
    );
    assert_eq!(cells_shaded(&banded, GREEN), 0, "…and none in band2");
}

/// `tblpPr` — floating-table positioning, and on §2.1.250(a)'s list. A style
/// that could float the tables it is applied to would move them off the flow
/// wherever it was used, which is why Word's UI cannot write one there and why
/// the erratum says Word does not read one.
#[test]
fn a_table_style_cannot_float_the_table() {
    assert_style_does_not_reach_the_table(
        r#"<w:tblpPr w:vertAnchor="text" w:tblpY="360" w:tblpXSpec="center"/>"#,
        "",
        "tblpPr",
    );
}

/// `tblOverlap` — also on the list, and only meaningful once the table floats,
/// so `tblpPr` is the constant background here (written directly, where it is
/// read) and `tblOverlap` is the variable.
#[test]
fn a_table_style_cannot_forbid_float_overlap() {
    assert_style_does_not_reach_the_table_with(
        &two_tables_document,
        r#"<w:tblpPr w:vertAnchor="text" w:tblpY="0"/>"#,
        r#"<w:tblOverlap w:val="never"/>"#,
        "",
        "tblOverlap",
    );
}

/// §17.4.1 `bidiVisual` — also on the list. Written on the `<w:tbl>` it
/// mirrors the column order (issue #157), which the fixture's A|B row makes
/// observable as the two texts trading x positions; written in the style it
/// must mirror nothing. Until #157 this test would have been vacuous — an
/// element nothing read could not discriminate, which is why the file's
/// coverage table listed it while no test carried it.
#[test]
fn a_table_style_cannot_mirror_the_columns() {
    assert_style_does_not_reach_the_table(r#"<w:bidiVisual/>"#, "", "bidiVisual");
}

// ── §17.7.4.3: `basedOn` inheritance of the style's own `<w:tblPr>` ─────────

/// A `<w:tblPr>` carrying one of every table property that reaches layout
/// **from a style** — so no `tblW`, `tblLook`, `tblpPr` or `tblOverlap`, which
/// §2.1.250(a) excludes and which would therefore be inert here whether the
/// `basedOn` merge carried them or not. That is the point of leaving them out:
/// an inert element cannot fail this test, so including it would weaken the
/// claim rather than widen it.
const EVERY_TABLE_PROPERTY: &str = r#"<w:tblPr>
    <w:tblStyleRowBandSize w:val="2"/>
    <w:tblStyleColBandSize w:val="2"/>
    <w:jc w:val="center"/>
    <w:tblCellSpacing w:w="144" w:type="dxa"/>
    <w:tblInd w:w="720" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
      <w:bottom w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
      <w:left w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
      <w:right w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
      <w:insideH w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
      <w:insideV w:val="single" w:sz="24" w:space="0" w:color="0000FF"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
      <w:bottom w:w="60" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>"#;

/// §17.7.4.3: a style "inherits all of the properties of the base style". The
/// unit of inheritance is the **property**, not the `<w:tblPr>` element that
/// carries it — so a child that declares one child element must not thereby
/// discard every other property the parent declared.
///
/// The child here declares `<w:tblLayout>`, which is as unrelated as a
/// `tblPr` child gets and moves nothing on its own. Before the fix that lone
/// element was enough to erase the parent's borders, alignment, indentation,
/// cell margins, cell spacing and band sizes in one go.
#[test]
fn a_child_table_style_inherits_every_property_it_does_not_restate() {
    let bands = r#"<w:tblStylePr w:type="band1Horz">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
           </w:tblStylePr>"#;
    let expected = layout(
        &table_document(""),
        &styles_with(&format!("{EVERY_TABLE_PROPERTY}{bands}")),
    );
    let inherited = layout(
        &table_document(""),
        &styles_derived_from(
            EVERY_TABLE_PROPERTY,
            &format!(r#"<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>{bands}"#),
        ),
    );

    assert!(
        page_geometry(&expected)
            .iter()
            .any(|g| g.ends_with("#0000FF")),
        "the fixture must draw the parent's borders, or it discriminates nothing"
    );
    assert_eq!(
        page_geometry(&inherited),
        page_geometry(&expected),
        "declaring one unrelated `tblPr` child must not drop the inherited ones"
    );
}

/// …and the child's own value still wins where it states one, leaving the rest
/// of the parent's showing. Inheritance that replaced wholesale and inheritance
/// that never applied would both pass the test above if the child said nothing;
/// this is the half that pins "layered on top".
#[test]
fn a_child_table_style_overrides_only_the_properties_it_restates() {
    let derived = layout(
        &table_document(""),
        &styles_derived_from(
            EVERY_TABLE_PROPERTY,
            r#"<w:tblPr><w:jc w:val="start"/></w:tblPr>"#,
        ),
    );
    // The parent centers the 200 pt table in a 468 pt content area, which would
    // put its left edge at 206 pt; the child pulls it back to the left margin,
    // where the parent's `tblInd` then applies. First text: 72 (margin) + 36
    // (tblInd) + 7.2 (one cell spacing) + 6 (left cell margin).
    assert_eq!(
        text_xs(&derived).first().copied(),
        Some(121.20),
        "the child's own jc wins over the parent's"
    );
    assert!(
        page_geometry(&derived)
            .iter()
            .any(|g| g.ends_with("#0000FF")),
        "…while the parent's borders, which the child never mentions, survive"
    );
}

/// §17.4.50: an explicitly-zero `tblInd` is the same indent as none.
///
/// Not a tautology in this engine: a full-width left-aligned table at zero
/// indent is shifted left by its left cell margin, so cell content lines up
/// with body text, and that shift used to be conditioned on *no `tblInd`
/// element* rather than on the resolved value being zero. Since every built-in
/// Word table style declares `<w:tblInd w:w="0"/>`, reading the style level
/// turns the two spellings into different renders unless the condition is
/// stated on the value — which is what this pins.
#[test]
fn a_zero_tbl_ind_indents_exactly_as_an_absent_one() {
    let cell_mar = r#"<w:tblCellMar>
          <w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>"#;
    let full_width = r#"<w:tblW w:w="5000" w:type="pct"/>"#;
    let absent = layout(
        &table_document(full_width),
        &styles_with(&format!("<w:tblPr>{cell_mar}</w:tblPr>")),
    );
    let explicit_zero = layout(
        &table_document(full_width),
        &styles_with(&format!(
            r#"<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>{cell_mar}</w:tblPr>"#
        )),
    );
    assert_eq!(
        page_geometry(&explicit_zero),
        page_geometry(&absent),
        "a declared zero indent must render as an undeclared one"
    );
    assert_eq!(
        text_xs(&absent).first().copied(),
        Some(72.0),
        "and both must be the shifted placement, which puts the first cell's \
         text on the left margin rather than one cell margin past it"
    );
}

/// The same rule at the **direct** level, which the guard also changed and no
/// test reached: `<w:tblInd w:w="0"/>` written on the `<w:tbl>` used to mean
/// "an indent element is present, so do not shift" and now means "the resolved
/// indent is zero, so shift" — the rule the comment at the site states, applied
/// where the style-level test cannot see it. Word cannot tell a declared zero
/// from an undeclared one either.
///
/// No corpus table has this shape (a full-width left-aligned table with a
/// direct zero `tblInd`), so nothing pinned it; the style level is the only
/// place the corpus exercises, and a guard written for the style level alone
/// would pass that test while silently deciding this one.
#[test]
fn a_direct_zero_tbl_ind_indents_exactly_as_an_absent_one() {
    let cell_mar = r#"<w:tblCellMar>
          <w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>"#;
    let full_width = r#"<w:tblW w:w="5000" w:type="pct"/>"#;
    let styles = styles_with("<w:tblPr/>");

    let absent = layout(&table_document(&format!("{full_width}{cell_mar}")), &styles);
    let explicit_zero = layout(
        &table_document(&format!(
            r#"{full_width}<w:tblInd w:w="0" w:type="dxa"/>{cell_mar}"#
        )),
        &styles,
    );
    assert_eq!(
        page_geometry(&explicit_zero),
        page_geometry(&absent),
        "a zero tblInd written directly on the table must render as an \
         undeclared one"
    );
    assert_eq!(
        text_xs(&absent).first().copied(),
        Some(72.0),
        "and both must be the shifted placement"
    );

    // The discriminator: a *non-zero* direct `tblInd` is still taken literally,
    // so the parity above is not "the guard ignores tblInd".
    let non_zero = layout(
        &table_document(&format!(
            r#"{full_width}<w:tblInd w:w="720" w:type="dxa"/>{cell_mar}"#
        )),
        &styles,
    );
    assert_ne!(
        page_geometry(&non_zero),
        page_geometry(&absent),
        "720 twips of indent must move the table"
    );
}

// ── §17.7.6 + §17.7.4.3: `basedOn` inheritance of `<w:tblStylePr>` ──────────

/// A `tblStylePr` is part of the style definition, so `basedOn` carries it like
/// any other property. A user style derived from a banded built-in used to lose
/// every conditional layer the built-in defined.
#[test]
fn a_child_table_style_inherits_the_parents_conditional_layers() {
    let first_row_red = r#"<w:tblStylePr w:type="firstRow">
             <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
           </w:tblStylePr>"#;
    let derived = layout(
        &table_document(""),
        &styles_derived_from(
            &format!("<w:tblPr/>{first_row_red}"),
            r#"<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>"#,
        ),
    );
    assert_eq!(
        cells_shaded(&derived, RED),
        2,
        "the parent's firstRow layer must shade both cells of the first row"
    );
}

/// Two layers of the same `w:type` merge property by property — the child's own
/// values win and the parent's fill the gaps — rather than the child's layer
/// replacing the parent's whole layer. That is the granularity §17.7.2 uses at
/// every other level of the cascade.
#[test]
fn a_child_conditional_layer_merges_with_the_parents_of_the_same_type() {
    let derived = layout(
        &table_document(""),
        &styles_derived_from(
            r#"<w:tblPr/>
               <w:tblStylePr w:type="firstRow">
                 <w:rPr><w:color w:val="FF0000"/></w:rPr>
                 <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:tcPr>
               </w:tblStylePr>"#,
            // The child restates only the shading.
            r#"<w:tblPr/>
               <w:tblStylePr w:type="firstRow">
                 <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="0000FF"/></w:tcPr>
               </w:tblStylePr>"#,
        ),
    );
    let geometry = page_geometry(&derived);
    assert_eq!(
        cells_shaded(&derived, (0x00, 0x00, 0xFF)),
        2,
        "the child's own shading wins"
    );
    assert_eq!(
        geometry
            .iter()
            .filter(|g| g.starts_with("rect") && g.ends_with("#00FF00"))
            .count(),
        0,
        "…and fully replaces the parent's, rather than both being drawn"
    );
    assert_eq!(
        geometry
            .iter()
            .filter(|g| g.starts_with("text") && g.ends_with("#FF0000"))
            .count(),
        2,
        "…while the parent's run color, which the child never restates, survives"
    );
}

/// The table-level half: an inherited `wholeTable` layer carries a `tblPr`, and
/// that has to reach `build_table`'s cascade, which reads it off the style's
/// folded table properties rather than off the conditional chain.
///
/// Written so the child *restates* the property, which is the case that
/// discriminates. A child that stayed silent would inherit the parent's already
/// folded value through plain `tblPr` inheritance and prove nothing about the
/// layer. Here the inherited layer has to outrank the child's own `tblPr` —
/// see `resolve_one`'s comment for why that is the reading taken and what
/// would settle it.
#[test]
fn an_inherited_whole_table_layer_outranks_the_childs_own_tbl_pr() {
    let borders = |color: &str| {
        format!(
            r#"<w:tblBorders>
                 <w:top w:val="single" w:sz="24" w:space="0" w:color="{color}"/>
                 <w:bottom w:val="single" w:sz="24" w:space="0" w:color="{color}"/>
                 <w:left w:val="single" w:sz="24" w:space="0" w:color="{color}"/>
                 <w:right w:val="single" w:sz="24" w:space="0" w:color="{color}"/>
                 <w:insideH w:val="single" w:sz="24" w:space="0" w:color="{color}"/>
                 <w:insideV w:val="single" w:sz="24" w:space="0" w:color="{color}"/>
               </w:tblBorders>"#
        )
    };
    let derived = layout(
        &table_document(""),
        &styles_derived_from(
            &format!(
                r#"<w:tblPr/>
                   <w:tblStylePr w:type="wholeTable">
                     <w:tblPr>{}</w:tblPr>
                   </w:tblStylePr>"#,
                borders("0000FF")
            ),
            &format!("<w:tblPr>{}</w:tblPr>", borders("FF0000")),
        ),
    );
    let geometry = page_geometry(&derived);
    assert!(
        geometry.iter().any(|g| g.ends_with("#0000FF")),
        "an inherited wholeTable tblPr must reach the table-level cascade"
    );
    assert_eq!(
        geometry.iter().filter(|g| g.ends_with("#FF0000")).count(),
        0,
        "…and sit above the child's own tblPr, exactly as its own would"
    );
}

/// A layer type only the child defines is added, not dropped, and the parent's
/// other types stay — the merge is by `w:type`, not a replacement of the set.
///
/// The `lastRow` region has to be switched on explicitly: §17.4.55's absent
/// `tblLook` is Word's 0x04A0, which leaves lastRow clear. Without this the
/// child's layer would resolve to nothing and the assertion below would be
/// measuring the region flag rather than the layer merge.
#[test]
fn conditional_layers_the_child_alone_defines_are_added_to_the_parents() {
    let derived = layout(
        &table_document(
            r#"<w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="0"
                          w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>"#,
        ),
        &styles_derived_from(
            r#"<w:tblPr/>
               <w:tblStylePr w:type="firstRow">
                 <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
               </w:tblStylePr>"#,
            r#"<w:tblPr/>
               <w:tblStylePr w:type="lastRow">
                 <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:tcPr>
               </w:tblStylePr>"#,
        ),
    );
    assert_eq!(
        cells_shaded(&derived, RED),
        2,
        "the parent's firstRow survives"
    );
    assert_eq!(
        cells_shaded(&derived, GREEN),
        2,
        "…alongside the child's own lastRow"
    );
}

/// The same, one level down: a child layer's `<w:tcPr>` merges with the
/// parent's for that region rather than replacing it, so stating one cell
/// property does not discard the region's inherited shading.
#[test]
fn a_child_layers_cell_properties_merge_with_the_parents() {
    let derived = layout(
        &table_document(""),
        &styles_derived_from(
            r#"<w:tblPr/>
               <w:tblStylePr w:type="firstRow">
                 <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr>
               </w:tblStylePr>"#,
            r#"<w:tblPr/>
               <w:tblStylePr w:type="firstRow">
                 <w:tcPr><w:vAlign w:val="center"/></w:tcPr>
               </w:tblStylePr>"#,
        ),
    );
    assert_eq!(
        cells_shaded(&derived, RED),
        2,
        "a child tcPr that states only vAlign must not drop the parent's shading"
    );
}

// ── §17.7.4.17: the document default table style ───────────────────────────

/// [`table_document`] with no `<w:tblStyle>` at all — the case the default
/// table style exists for.
fn styleless_table_document(direct_tbl_pr: &str) -> String {
    let with_style = table_document(direct_tbl_pr);
    let out = with_style.replace(r#"<w:tblStyle w:val="TestTbl"/>"#, "");
    assert_ne!(
        out, with_style,
        "the fixture must have had a tblStyle to drop"
    );
    out
}

/// A stylesheet whose `TableNormal` carries `body`, marked default or not.
fn styles_with_table_normal(body: &str, is_default: bool) -> String {
    let default_attr = if is_default { r#" w:default="1""# } else { "" };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="TableNormal"{default_attr}>
    <w:name w:val="Normal Table"/>
    {body}
  </w:style>
  <w:style w:type="table" w:styleId="TestTbl">
    <w:name w:val="Test Table"/>
    <w:tblPr/>
  </w:style>
</w:styles>"#
    )
}

/// Word's own `TableNormal`, as Word writes it.
const TABLE_NORMAL: &str = r#"<w:tblPr>
    <w:tblInd w:w="0" w:type="dxa"/>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>"#;

/// §17.7.4.17: `w:default="1"` means the style applies to objects of that type
/// which reference no style. For tables that is `TableNormal`, and its 108-twip
/// left/right `tblCellMar` is why cell text in a Word table sits 5.4 pt in from
/// the cell edge rather than against it.
#[test]
fn a_table_naming_no_style_takes_the_document_default_table_style() {
    let applied = layout(
        &styleless_table_document(""),
        &styles_with_table_normal(TABLE_NORMAL, true),
    );
    assert_eq!(
        text_xs(&applied).first().copied(),
        Some(77.40),
        "the default style's 108-twip left cell inset applies: 72 + 5.4"
    );

    // Control: the very same style, not marked default, reaches nothing.
    let unmarked = layout(
        &styleless_table_document(""),
        &styles_with_table_normal(TABLE_NORMAL, false),
    );
    assert_eq!(
        text_xs(&unmarked).first().copied(),
        Some(72.00),
        "without w:default the style is just another unreferenced style"
    );
}

/// …and only when the table names none. A table that names a style resolves
/// through that style's `basedOn` chain instead, which is how Word's own
/// built-in table styles reach `TableNormal`.
#[test]
fn a_table_naming_a_style_does_not_also_take_the_default() {
    let named = layout(
        // `TestTbl` declares nothing and is not `basedOn` TableNormal.
        &table_document(""),
        &styles_with_table_normal(TABLE_NORMAL, true),
    );
    assert_eq!(
        text_xs(&named).first().copied(),
        Some(72.00),
        "naming a style opts out of the default, per §17.7.4.17"
    );
}

/// The default is the *base* of the cascade, not the top of it: a direct
/// `<w:tblCellMar>` on the table still wins, per edge.
#[test]
fn a_direct_property_still_beats_the_default_table_style() {
    let overridden = layout(
        &styleless_table_document(
            r#"<w:tblCellMar><w:left w:w="720" w:type="dxa"/></w:tblCellMar>"#,
        ),
        &styles_with_table_normal(TABLE_NORMAL, true),
    );
    assert_eq!(
        text_xs(&overridden).first().copied(),
        Some(108.00),
        "the table's own 720-twip left inset wins over the default's 108"
    );
}
