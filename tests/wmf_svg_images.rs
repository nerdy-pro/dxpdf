//! WMF and SVG images, end to end (issue #150): both formats used to be
//! recognised as image parts and then dropped — the space reserved, nothing
//! drawn. Now a WMF's embedded bitmap decodes like the EMF path always did,
//! and an SVG rasterizes at the display target — with Word's `svgBlip`
//! extension read, so a picture carrying the SVG-plus-PNG-fallback pair
//! ([MS-ODRAWXML] "Pictures") renders the SVG rather than the fallback.
//!
//! The fixtures (`scripts/make_issue150_fixtures.py`) make each decision
//! observable: the fallback PNG is blue while the svgBlip's SVG is red, so
//! which part was picked *is* the drawn colour; the WMF's 2×2 DIB has four
//! distinct quadrant colours. Decode-level pixel assertions live with the
//! decoders' unit tests (`render::wmf`, `render::svg`, `render::dib`); this
//! file pins the pipeline: parse → media table → part selection → a PDF
//! that embeds an image XObject where it used to embed none.

use dxpdf::model::{Block, GraphicContent, ImageFormat, Inline, RelId};

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

fn load(name: &str) -> dxpdf::model::Document {
    let path = format!("{TEST_DIR}/{name}");
    let data = std::fs::read(&path).unwrap_or_else(|e| panic!("failed to read {path}: {e}"));
    dxpdf::docx::parse(&data).unwrap_or_else(|e| panic!("failed to parse {path}: {e}"))
}

/// Every picture blip in body order.
fn blips(doc: &dxpdf::model::Document) -> Vec<&dxpdf::model::Blip> {
    doc.body
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .flat_map(|p| &p.content)
        .filter_map(|i| match i {
            Inline::Image(img) => match img.graphic.as_ref()? {
                GraphicContent::Picture(pic) => pic.blip_fill.blip.as_ref(),
                _ => None,
            },
            _ => None,
        })
        .collect()
}

/// The `asvg:svgBlip` extension parses into the model — its relationship id
/// lands beside the main embed, and an extLst holding only *other*
/// extensions (the second picture has none at all) yields nothing.
#[test]
fn svg_blip_extension_parses_into_the_model() {
    let doc = load("svg-image.docx");
    let blips = blips(&doc);
    assert_eq!(blips.len(), 2, "two pictures in the fixture");

    assert_eq!(blips[0].embed, Some(RelId::new("rIdP")), "the PNG fallback");
    assert_eq!(
        blips[0].svg_embed,
        Some(RelId::new("rIdS")),
        "…and the svgBlip's SVG beside it"
    );
    assert_eq!(blips[1].embed, Some(RelId::new("rIdS2")));
    assert_eq!(blips[1].svg_embed, None, "no extension on the direct blip");
}

/// Both parts of the pair reach the media table with their formats — the
/// SVG part is not lost for being referenced only from an extension.
#[test]
fn media_parts_carry_their_formats() {
    let doc = load("svg-image.docx");
    assert_eq!(doc.media[&RelId::new("rIdP")].1, ImageFormat::Png);
    assert_eq!(doc.media[&RelId::new("rIdS")].1, ImageFormat::Svg);

    let doc = load("wmf-image.docx");
    assert_eq!(doc.media[&RelId::new("rIdW")].1, ImageFormat::Wmf);
}

/// The SVG is selected over the PNG fallback: [MS-ODRAWXML] says the
/// rasterized copy exists "for backward compatibility" with consumers that
/// cannot draw SVG — this one now can, so both pictures' draw commands must
/// carry the SVG parts, not the fallback.
#[test]
fn the_svg_part_is_selected_over_the_png_fallback() {
    use dxpdf::render::layout::draw_command::DrawCommand;

    let doc = load("svg-image.docx");
    let (_, pages) = dxpdf::render::resolve_and_layout(doc);
    let formats: Vec<ImageFormat> = pages
        .iter()
        .flat_map(|p| &p.commands)
        .filter_map(|c| match c {
            DrawCommand::Image { image_data, .. } => Some(image_data.format),
            _ => None,
        })
        .collect();
    assert_eq!(
        formats,
        [ImageFormat::Svg, ImageFormat::Svg],
        "both pictures resolve to their SVG parts"
    );
}

/// The end of the pipe: converting each fixture yields a PDF that embeds an
/// image XObject. Before #150 the WMF and the direct-SVG picture decoded to
/// nothing, and no `/Image` reached the file.
#[test]
fn both_formats_reach_the_pdf_as_images() {
    for fixture in ["wmf-image.docx", "svg-image.docx"] {
        let path = format!("{TEST_DIR}/{fixture}");
        let bytes = std::fs::read(&path).unwrap();
        let pdf = dxpdf::convert(&bytes).unwrap_or_else(|e| panic!("{fixture}: {e}"));
        let hits = pdf.windows(6).filter(|w| w == b"/Image").count();
        assert!(
            hits > 0,
            "{fixture}: the PDF must embed an image XObject; {} bytes, no /Image",
            pdf.len()
        );
    }
}
