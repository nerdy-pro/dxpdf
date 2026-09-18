//! §17.6.10 `<w:pgBorders>` — page borders end to end (issue #151).
//!
//! The committed fixture `test-files/page-borders.docx` carries the same four
//! deliberately distinct edges through both offset modes — section 1
//! `offsetFrom="page"`, section 2 `offsetFrom="text"` — so every edge's band
//! is pinned by its own colour and the two pages differ only in the reference
//! frame. The in-memory documents cover the selection attributes (`display`,
//! `zOrder`) and the declines (`nil`, art names).
//!
//! Geometry pinned here, from §17.6.10 + [MS-OE376] §2.6.10 (US Letter,
//! 612×792pt, 72pt margins; `sz` in eighths, `space` in points):
//!
//! - `page` mode measures `space` from the page edge, line growing inward:
//!   top (6pt line, 24pt space) sits at y ∈ [24, 30].
//! - `text` mode measures from the text margin, line growing outward: the
//!   same top edge sits at y ∈ [72−24−6, 72−24] = [42, 48].

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

/// Every `Rect` on the page with this exact colour, as (x, y, w, h).
fn rects_of_color(
    page: &dxpdf::render::layout::draw_command::LayoutedPage,
    rgb: (u8, u8, u8),
) -> Vec<(f32, f32, f32, f32)> {
    page.commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color } if (color.r, color.g, color.b) == rgb => Some((
                rect.origin.x.raw(),
                rect.origin.y.raw(),
                rect.size.width.raw(),
                rect.size.height.raw(),
            )),
            _ => None,
        })
        .collect()
}

fn assert_rect_eq(got: (f32, f32, f32, f32), want: (f32, f32, f32, f32), what: &str) {
    let close = |a: f32, b: f32| (a - b).abs() < 1e-3;
    assert!(
        close(got.0, want.0)
            && close(got.1, want.1)
            && close(got.2, want.2)
            && close(got.3, want.3),
        "{what}: got {got:?}, want {want:?}"
    );
}

// ── The committed fixture: both offset modes, all four edges ────────────────

#[test]
fn fixture_page_mode_places_each_edge_from_the_page_edge() {
    let bytes = std::fs::read("test-files/page-borders.docx").expect("fixture");
    let pages = layout(&bytes);
    assert_eq!(pages.len(), 2, "one page per section");
    let page = &pages[0];

    // Top: red, sz=48 → 6pt, space=24pt from the page edge.
    let red = rects_of_color(page, (255, 0, 0));
    assert_eq!(red.len(), 1, "one top band");
    // Horizontal edges span between the vertical edges' outer x: left band
    // starts at 12 (its space), right band ends at 612 (space 0).
    assert_rect_eq(red[0], (12.0, 24.0, 600.0, 6.0), "top band");

    // Left: green, sz=24 → 3pt, space=12pt; runs between the horizontal
    // bands (top ends at 30, bottom starts at 783).
    let green = rects_of_color(page, (0, 255, 0));
    assert_eq!(green.len(), 1);
    assert_rect_eq(green[0], (12.0, 30.0, 3.0, 753.0), "left band");

    // Bottom: blue, double, sz=24 → 3pt band, space=6pt up from the page
    // edge → band y ∈ [783, 786]. §17.4.38: double = two lines of sz/3 at
    // the band's extremes.
    let blue = rects_of_color(page, (0, 0, 255));
    assert_eq!(blue.len(), 2, "a double edge is two sub-rects");
    assert_rect_eq(blue[0], (12.0, 783.0, 600.0, 1.0), "double outer line");
    assert_rect_eq(blue[1], (12.0, 785.0, 600.0, 1.0), "double inner line");

    // Right: magenta, sz=8 → 1pt, space=0 → hugs the page edge.
    let magenta = rects_of_color(page, (255, 0, 255));
    assert_eq!(magenta.len(), 1);
    assert_rect_eq(magenta[0], (611.0, 30.0, 1.0, 753.0), "right band");
}

#[test]
fn fixture_text_mode_places_each_edge_from_the_text_margin() {
    let bytes = std::fs::read("test-files/page-borders.docx").expect("fixture");
    let pages = layout(&bytes);
    let page = &pages[1];

    // Top: 24pt outward from the 72pt margin, growing toward the edge.
    let red = rects_of_color(page, (255, 0, 0));
    assert_eq!(red.len(), 1);
    assert_rect_eq(red[0], (57.0, 42.0, 484.0, 6.0), "top band (text mode)");

    // Left: inner edge 12pt outside the left margin.
    let green = rects_of_color(page, (0, 255, 0));
    assert_eq!(green.len(), 1);
    assert_rect_eq(green[0], (57.0, 48.0, 3.0, 678.0), "left band (text mode)");

    // Bottom: 6pt below the bottom text margin (y = 720), double.
    let blue = rects_of_color(page, (0, 0, 255));
    assert_eq!(blue.len(), 2);
    assert_rect_eq(blue[0], (57.0, 726.0, 484.0, 1.0), "double outer");
    assert_rect_eq(blue[1], (57.0, 728.0, 484.0, 1.0), "double inner");

    // Right: space=0 → flush against the right text margin (x = 540).
    let magenta = rects_of_color(page, (255, 0, 255));
    assert_eq!(magenta.len(), 1);
    assert_rect_eq(
        magenta[0],
        (540.0, 48.0, 1.0, 678.0),
        "right band (text mode)",
    );
}

// ── In-memory probes: display, zOrder, declines ─────────────────────────────

fn docx_with_sect_pr(paragraphs: &str, sect_pr_children: &str) -> Vec<u8> {
    make_docx(&format!(
        "<w:body>{paragraphs}<w:sectPr>\
         <w:pgSz w:w=\"12240\" w:h=\"15840\"/>\
         <w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" \
          w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>{sect_pr_children}</w:sectPr></w:body>"
    ))
}

const TWO_PAGES: &str = "<w:p><w:r><w:t>first page</w:t></w:r></w:p>\
     <w:p><w:r><w:br w:type=\"page\"/><w:t>second page</w:t></w:r></w:p>";

const RED_TOP: &str =
    "<w:pgBorders><w:top w:val=\"single\" w:sz=\"48\" w:space=\"24\" w:color=\"FF0000\"/></w:pgBorders>";

fn red_rect_count(page: &dxpdf::render::layout::draw_command::LayoutedPage) -> usize {
    rects_of_color(page, (255, 0, 0)).len()
}

/// §17.6.10: absent `display` = all pages of the section.
#[test]
fn display_defaults_to_all_pages() {
    let pages = layout(&docx_with_sect_pr(TWO_PAGES, RED_TOP));
    assert_eq!(pages.len(), 2);
    assert_eq!(red_rect_count(&pages[0]), 1);
    assert_eq!(red_rect_count(&pages[1]), 1);
}

#[test]
fn display_first_page_and_not_first_page_select_by_section_page_index() {
    let first = RED_TOP.replace("<w:pgBorders>", "<w:pgBorders w:display=\"firstPage\">");
    let pages = layout(&docx_with_sect_pr(TWO_PAGES, &first));
    assert_eq!(red_rect_count(&pages[0]), 1);
    assert_eq!(red_rect_count(&pages[1]), 0);

    let not_first = RED_TOP.replace("<w:pgBorders>", "<w:pgBorders w:display=\"notFirstPage\">");
    let pages = layout(&docx_with_sect_pr(TWO_PAGES, &not_first));
    assert_eq!(red_rect_count(&pages[0]), 0);
    assert_eq!(red_rect_count(&pages[1]), 1);
}

/// §17.18.67: `front` (the default) paints the border after everything else
/// on the page, `back` before everything — painting order is z-order.
#[test]
fn z_order_places_border_commands_around_the_page_content() {
    let text = "<w:p><w:r><w:t>body</w:t></w:r></w:p>";

    let border_pos = |pages: &[dxpdf::render::layout::draw_command::LayoutedPage]| {
        let page = &pages[0];
        let border = page
            .commands
            .iter()
            .position(|c| matches!(c, DrawCommand::Rect { color, .. } if color.r == 255))
            .expect("border rect present");
        let text = page
            .commands
            .iter()
            .position(|c| matches!(c, DrawCommand::Text { .. }))
            .expect("text present");
        (border, text)
    };

    let (border_at, text_at) = border_pos(&layout(&docx_with_sect_pr(text, RED_TOP)));
    assert!(
        border_at > text_at,
        "front (default): border paints over content"
    );

    let back = RED_TOP.replace("<w:pgBorders>", "<w:pgBorders w:zOrder=\"back\">");
    let (border_at, text_at) = border_pos(&layout(&docx_with_sect_pr(text, &back)));
    assert!(border_at < text_at, "back: border paints under content");
}

/// Art borders draw nothing (the LibreOffice reading); `nil`/`none` edges
/// draw nothing by definition. Neither may take down the render.
#[test]
fn art_and_nil_edges_draw_no_border() {
    let text = "<w:p><w:r><w:t>body</w:t></w:r></w:p>";
    let art = "<w:pgBorders w:offsetFrom=\"page\">\
        <w:top w:val=\"apples\" w:sz=\"31\" w:space=\"24\"/>\
        <w:left w:val=\"nil\" w:sz=\"24\" w:space=\"12\"/>\
        <w:bottom w:val=\"none\" w:sz=\"24\" w:space=\"12\"/>\
        </w:pgBorders>";
    let pages = layout(&docx_with_sect_pr(text, art));
    let rects = pages[0]
        .commands
        .iter()
        .filter(|c| matches!(c, DrawCommand::Rect { .. }))
        .count();
    assert_eq!(rects, 0, "no border band from art, nil, or none edges");
}

/// §17.11.2: the endnotes page continues the last section's flow, so it
/// wears that section's page borders — and at the page index *after* the
/// section's own pages, so `display="firstPage"` does not re-fire there.
#[test]
fn endnote_page_wears_the_last_sections_borders() {
    let pages = layout(&endnote_docx(RED_TOP));
    assert_eq!(pages.len(), 2, "body page + endnotes page");
    assert_eq!(red_rect_count(&pages[0]), 1);
    assert_eq!(
        red_rect_count(&pages[1]),
        1,
        "the endnotes page is framed too"
    );

    // `display="firstPage"`: the endnotes page is the section's *second*
    // page, so the border stays off it (and on the body page).
    let first = RED_TOP.replace("<w:pgBorders>", "<w:pgBorders w:display=\"firstPage\">");
    let pages = layout(&endnote_docx(&first));
    assert_eq!(red_rect_count(&pages[0]), 1);
    assert_eq!(red_rect_count(&pages[1]), 0, "firstPage does not re-fire");
}

fn endnote_docx(pg_borders: &str) -> Vec<u8> {
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
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"#).unwrap();
    zip.start_file("_rels/.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();
    zip.start_file("word/_rels/document.xml.rels", o).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"#).unwrap();
    zip.start_file("word/endnotes.xml", o).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:id="2"><w:p><w:r><w:t>the endnote</w:t></w:r></w:p></w:endnote>
</w:endnotes>"#,
    )
    .unwrap();
    zip.start_file("word/document.xml", o).unwrap();
    zip.write_all(format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>body</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
{pg_borders}</w:sectPr></w:body></w:document>"#
    ).as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

/// [MS-OE376] §2.6.10: an absent `offsetFrom` means `text`.
#[test]
fn offset_from_defaults_to_text() {
    let text = "<w:p><w:r><w:t>body</w:t></w:r></w:p>";
    let pages = layout(&docx_with_sect_pr(text, RED_TOP));
    let red = rects_of_color(&pages[0], (255, 0, 0));
    assert_eq!(red.len(), 1);
    // Text mode: y ∈ [72−24−6, 72−24]; span defaults to the text extent.
    assert_rect_eq(red[0], (72.0, 42.0, 468.0, 6.0), "text-mode default top");
}
