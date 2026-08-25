//! Issue #155 end to end: SmartArt diagrams and charts render as vector
//! scenes instead of vanishing.
//!
//! The committed fixtures carry the stock Office theme, so the scheme fills
//! resolve to the familiar accent RGBs and every color assertion is a
//! literal. Geometry is pinned in Pt straight from the fixtures' EMU
//! (÷12700), positioned relative to the drawing's own box so no font metric
//! enters the arithmetic.

use dxpdf::render::layout::draw_command::{DrawCommand, ResolvedFill};

/// A filled path's origin and its 8-bit fill color.
type FilledPath = ((f32, f32), (u8, u8, u8));

const ACCENT1: (u8, u8, u8) = (0x44, 0x72, 0xC4);
const ACCENT2: (u8, u8, u8) = (0xED, 0x7D, 0x31);
const ACCENT3: (u8, u8, u8) = (0xA5, 0xA5, 0xA5);

fn layout(path: &str) -> Vec<dxpdf::render::layout::draw_command::LayoutedPage> {
    let bytes = std::fs::read(path).expect("fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse");
    dxpdf::render::resolve_and_layout(doc).1
}

fn solid_paths(page: &dxpdf::render::layout::draw_command::LayoutedPage) -> Vec<FilledPath> {
    page.commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Path {
                origin,
                fill: ResolvedFill::Solid(f),
                ..
            } => Some((
                (origin.x.raw(), origin.y.raw()),
                (
                    (f.r * 255.0).round() as u8,
                    (f.g * 255.0).round() as u8,
                    (f.b * 255.0).round() as u8,
                ),
            )),
            _ => None,
        })
        .collect()
}

/// Whether the page's concatenated text contains `needle`. Concatenated
/// because line breaking splits a run into word-level commands.
fn has_text(page: &dxpdf::render::layout::draw_command::LayoutedPage, needle: &str) -> bool {
    let all: String = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Text { text, .. } => Some(text.as_ref()),
            _ => None,
        })
        .collect::<Vec<_>>()
        .join(" ");
    all.contains(needle) || all.replace("  ", " ").contains(needle)
}

// ── SmartArt ────────────────────────────────────────────────────────────────

/// The parse side: the diagram payload is keyed by the `r:dm` id, holds the
/// flattened `dsp:spTree`, and the body's drawing references it.
#[test]
fn smartart_parses_the_drawing_part() {
    let bytes = std::fs::read("test-files/smartart.docx").expect("fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse");
    assert_eq!(doc.diagrams.len(), 1);
    let drawing = doc.diagrams.values().next().unwrap();
    assert_eq!(drawing.shapes.len(), 4, "three nodes + one arrow");
    let with_text = drawing
        .shapes
        .iter()
        .filter(|s| s.text_body.is_some())
        .count();
    assert_eq!(with_text, 3, "the arrow carries no label");
}

/// The render side: four filled shapes at the fixture's EMU positions with
/// the Office accent colors, and the three labels drawn over them.
#[test]
fn smartart_renders_nodes_arrow_and_labels() {
    let pages = layout("test-files/smartart.docx");
    let page = &pages[0];

    let paths = solid_paths(page);
    let accent1 = paths.iter().filter(|(_, c)| *c == ACCENT1).count();
    let accent2 = paths.iter().filter(|(_, c)| *c == ACCENT2).count();
    let accent3 = paths.iter().filter(|(_, c)| *c == ACCENT3).count();
    assert_eq!(accent1, 1, "node One");
    assert_eq!(accent2, 2, "node Two and the arrow");
    assert_eq!(accent3, 1, "node Three");

    // Node spacing survives into page coordinates: the second roundRect
    // sits exactly (node + gap) = 2057400 EMU = 162pt right of the first.
    let mut accent_x: Vec<f32> = paths
        .iter()
        .filter(|(_, c)| *c == ACCENT1 || *c == ACCENT3)
        .map(|((x, _), _)| *x)
        .collect();
    accent_x.sort_by(f32::total_cmp);
    let dx = accent_x[1] - accent_x[0];
    assert!(
        (dx - 324.0).abs() < 0.5,
        "One→Three spacing is 2·162pt, got {dx}"
    );

    for label in ["One", "Two", "Three"] {
        assert!(has_text(page, label), "label {label} missing");
    }
    assert!(has_text(page, "diagram."));
    assert!(has_text(page, "After"));
}

/// The labels sit inside their nodes: each label's x lies between its
/// node's left and right edge (`dsp:txXfrm` honored), and the node fills
/// paint before the text so the text stays visible.
#[test]
fn smartart_labels_sit_inside_their_nodes() {
    let pages = layout("test-files/smartart.docx");
    let page = &pages[0];
    let node_one_x = solid_paths(page)
        .iter()
        .find(|(_, c)| *c == ACCENT1)
        .map(|((x, _), _)| *x)
        .expect("node One");
    let one = page
        .commands
        .iter()
        .find_map(|c| match c {
            DrawCommand::Text { text, position, .. } if text.contains("One") => {
                Some(position.x.raw())
            }
            _ => None,
        })
        .expect("label One");
    // Node width = 1371600 EMU = 108pt.
    assert!(
        one > node_one_x && one < node_one_x + 108.0,
        "label at {one}, node at {node_one_x}"
    );

    let path_idx = page
        .commands
        .iter()
        .position(|c| {
            matches!(
                c,
                DrawCommand::Path {
                    fill: ResolvedFill::Solid(_),
                    ..
                }
            )
        })
        .unwrap();
    let text_idx = page
        .commands
        .iter()
        .position(|c| matches!(c, DrawCommand::Text { text, .. } if text.contains("One")))
        .unwrap();
    assert!(path_idx < text_idx, "fill under label");
}

// ── Charts ──────────────────────────────────────────────────────────────────

/// The parse side: three chart parts with cached series.
#[test]
fn charts_parse_their_cached_series() {
    let bytes = std::fs::read("test-files/charts.docx").expect("fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse");
    assert_eq!(doc.charts.len(), 3);
    let bar = doc
        .charts
        .values()
        .find(|c| {
            c.plot_groups
                .first()
                .is_some_and(|g| matches!(g.kind, dxpdf::docx::model::PlotKind::Bar { .. }))
        })
        .expect("bar chart");
    assert_eq!(bar.plot_groups[0].series.len(), 2);
    assert_eq!(
        bar.plot_groups[0].series[0].values,
        vec![Some(4.0), Some(7.0), Some(5.0)]
    );
}

/// The bar chart: six bars (2 series × 3 categories) in the two accent
/// colors, with heights proportional to the cached values.
#[test]
fn bar_chart_draws_proportional_bars() {
    let pages = layout("test-files/charts.docx");
    let page = &pages[0];

    let bars: Vec<(f32, f32, (u8, u8, u8))> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, color }
                if (color.r, color.g, color.b) == ACCENT1
                    || (color.r, color.g, color.b) == ACCENT2 =>
            {
                // Legend swatches are 7pt squares; bars are taller.
                (rect.size.height.raw() > 8.0).then(|| {
                    (
                        rect.origin.x.raw(),
                        rect.size.height.raw(),
                        (color.r, color.g, color.b),
                    )
                })
            }
            _ => None,
        })
        .collect();
    assert_eq!(bars.len(), 6, "2 series × 3 categories");

    let north: Vec<&(f32, f32, (u8, u8, u8))> =
        bars.iter().filter(|(_, _, c)| *c == ACCENT1).collect();
    assert_eq!(north.len(), 3);
    // North = [4, 7, 5]: heights scale as the values (the axis runs 0..8).
    let h4 = north[0].1;
    let h7 = north[1].1;
    let h5 = north[2].1;
    assert!(
        (h7 / h4 - 7.0 / 4.0).abs() < 0.01,
        "7:4 ratio, got {}",
        h7 / h4
    );
    assert!((h5 / h4 - 5.0 / 4.0).abs() < 0.01);

    for label in ["Sales", "Q1", "Q2", "Q3", "North", "South", "0", "8"] {
        assert!(has_text(page, label), "chart text {label} missing");
    }
}

/// The pie: four slices in the per-point accent cycle (varyColors), share
/// legend on the right listing the categories.
#[test]
fn pie_chart_draws_a_slice_per_point() {
    let pages = layout("test-files/charts.docx");
    let page = &pages[0];
    let slices: Vec<(u8, u8, u8)> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Path {
                fill: ResolvedFill::Solid(f),
                stroke: Some(s),
                ..
            } if s.color.r == 1.0 && s.color.g == 1.0 && s.color.b == 1.0 => Some((
                (f.r * 255.0).round() as u8,
                (f.g * 255.0).round() as u8,
                (f.b * 255.0).round() as u8,
            )),
            _ => None,
        })
        .collect();
    assert_eq!(slices.len(), 4, "four white-separated slices");
    assert_eq!(slices[0], ACCENT1);
    assert_eq!(slices[1], ACCENT2);
    for cat in ["A", "B", "C", "D"] {
        assert!(has_text(page, cat), "pie legend {cat}");
    }
}

/// The line chart: a polyline broken at the blank cell — runs a–c and the
/// lone point e — with circle markers on present points only.
#[test]
fn line_chart_breaks_at_the_blank_cell() {
    let pages = layout("test-files/charts.docx");
    let page = &pages[0];
    // Stroked non-filled paths at 2.25pt = the series polylines.
    let polylines: Vec<usize> = page
        .commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Path {
                fill: ResolvedFill::None,
                stroke: Some(s),
                paths,
                ..
            } if (s.width.raw() - 2.25).abs() < 0.01 => Some(paths[0].verbs.len()),
            _ => None,
        })
        .collect();
    // One run of 3 points (a,b,c). The lone point e yields no polyline.
    assert_eq!(
        polylines,
        vec![3],
        "one 3-point run; the gap breaks the line"
    );

    // Circle markers: filled accent1 paths of 4 cubics each — 4 present
    // points (a, b, c, e).
    let markers = page
        .commands
        .iter()
        .filter(|c| match c {
            DrawCommand::Path {
                fill: ResolvedFill::Solid(f),
                stroke: None,
                paths,
                ..
            } => {
                (f.r * 255.0).round() as u8 == ACCENT1.0
                    && paths[0]
                        .verbs
                        .iter()
                        .filter(|v| {
                            matches!(
                                v,
                                dxpdf::render::resolve::shape_geometry::PathVerb::CubicTo(..)
                            )
                        })
                        .count()
                        == 4
            }
            _ => false,
        })
        .count();
    assert_eq!(markers, 4, "markers on the four present points");
}

// ── Anchored placement ─────────────────────────────────────────────────────

/// An *anchored* chart rides the floating-shape channel: same scene, placed
/// at the anchor offset, with text wrapping registered.
#[test]
fn anchored_chart_renders_at_its_anchor() {
    use std::io::Write;
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(buf);
    let o = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    let src = std::fs::read("test-files/charts.docx").unwrap();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(src)).unwrap();
    for i in 0..archive.len() {
        let mut f = archive.by_index(i).unwrap();
        let name = f.name().to_string();
        let mut data = Vec::new();
        std::io::Read::read_to_end(&mut f, &mut data).unwrap();
        if name == "word/document.xml" {
            // Re-body: one anchored bar chart at absolute (72pt, 72pt).
            let doc = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body><w:p><w:r><w:t>wrapped text</w:t></w:r><w:r><w:drawing>
<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1"
 behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>
<wp:extent cx="2743200" cy="1828800"/>
<wp:wrapSquare wrapText="bothSides"/>
<wp:docPr id="2" name="AnchoredChart"/>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId20"/>
</a:graphicData></a:graphic></wp:anchor>
</w:drawing></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>"#;
            zip.start_file(name, o).unwrap();
            zip.write_all(doc.as_bytes()).unwrap();
        } else {
            zip.start_file(name, o).unwrap();
            zip.write_all(&data).unwrap();
        }
    }
    let docx = zip.finish().unwrap().into_inner();

    let doc = dxpdf::docx::parse(&docx).expect("parse");
    let pages = dxpdf::render::resolve_and_layout(doc).1;
    let page = &pages[0];

    // The chart frame path lands at the anchor: origin (72, 72).
    let frame = page
        .commands
        .iter()
        .find_map(|c| match c {
            DrawCommand::Path {
                origin,
                fill: ResolvedFill::None,
                stroke: Some(_),
                ..
            } => Some((origin.x.raw(), origin.y.raw())),
            _ => None,
        })
        .expect("chart frame");
    assert_eq!(frame, (72.0, 72.0));
    assert!(has_text(page, "Sales"), "anchored chart draws its title");
    assert!(has_text(page, "wrapped"));
}
