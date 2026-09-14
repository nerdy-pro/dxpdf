//! SVG rasterization for embedded images (issue #150).
//!
//! Word embeds SVG through the `asvg:svgBlip` extension, and other producers
//! point the main blip straight at an `image/svg+xml` part; either way the
//! part reaches the painter as a [`crate::model::ImageFormat::Svg`] media
//! entry, which Skia's built-in decoder does not read.
//!
//! Rendered via `resvg` (Linebender), rasterizing at the **display target**
//! the painter already computes for downsampling — so an SVG comes out crisp
//! at exactly the `--image-dpi` every raster image is normalized to, and two
//! placements at different sizes rasterize separately through the painter's
//! size-keyed bitmap cache.
//!
//! # Why resvg, and why a raster
//!
//! Skia has its own SVG module, but enabling `skia-safe`'s `svg` feature
//! forfeits the prebuilt-binary cache on every platform this project builds
//! on (no published binary carries `svg` next to `embed-freetype`, and none
//! carries it without a GPU backend or `webp` in tow), turning every cold
//! build into a from-source Skia build. `resvg` with default features off is
//! pure Rust, ~40 crates, seconds to build, MIT/Apache-licensed. The price
//! is a raster in the PDF instead of vectors — the same deal every other
//! embedded image already gets — and no `<text>` support: the `text` feature
//! would pull a second font stack (`fontdb`/`harfrust`) beside Skia's, so
//! text inside an embedded SVG is dropped. Office-inserted SVGs are
//! overwhelmingly icons and diagrams; revisit if a corpus says otherwise.

use skia_safe::{images, AlphaType, ColorType, Data, Image, ImageInfo};

/// Rasters larger than this on a side are clamped, preserving aspect — a
/// guard against a malformed viewBox times a large display target.
const MAX_DIMENSION: u32 = 8192;

/// Render an SVG to a Skia image.
///
/// `target` is the display size in pixels — the painter's downsample target
/// for the rect the image is drawn into. The raster is stretched to exactly
/// that size, matching how `a:stretch` fills the picture extent; `None`
/// renders at the SVG's own intrinsic size.
pub fn decode_svg(data: &[u8], target: Option<(u32, u32)>) -> Option<Image> {
    let (width, height, rgba) = rasterize(data, target)?;
    let info = ImageInfo::new(
        (width as i32, height as i32),
        ColorType::RGBA8888,
        // resvg produces premultiplied alpha (tiny-skia's native format),
        // and unlike the DIB decoders an SVG genuinely carries transparency.
        AlphaType::Premul,
        None,
    );
    images::raster_from_data(&info, Data::new_copy(&rgba), width as usize * 4)
}

/// The resvg half, separated so the pixels are assertable without Skia:
/// premultiplied RGBA, top-down.
fn rasterize(data: &[u8], target: Option<(u32, u32)>) -> Option<(u32, u32, Vec<u8>)> {
    let options = resvg::usvg::Options::default();
    let tree = resvg::usvg::Tree::from_data(data, &options).ok()?;
    let size = tree.size();
    if !(size.width() > 0.0 && size.height() > 0.0) {
        return None;
    }

    let (width, height) = match target {
        Some((w, h)) => (w.max(1), h.max(1)),
        None => (
            size.width().ceil().max(1.0) as u32,
            size.height().ceil().max(1.0) as u32,
        ),
    };
    let width = width.min(MAX_DIMENSION);
    let height = height.min(MAX_DIMENSION);

    let mut pixmap = resvg::tiny_skia::Pixmap::new(width, height)?;
    resvg::render(
        &tree,
        resvg::tiny_skia::Transform::from_scale(
            width as f32 / size.width(),
            height as f32 / size.height(),
        ),
        &mut pixmap.as_mut(),
    );
    Some((width, height, pixmap.take()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const RED_4X4: &[u8] = br##"<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4" viewBox="0 0 4 4"><rect width="4" height="4" fill="#FF0000"/></svg>"##;

    /// A solid rect renders as its fill, stretched to the display target —
    /// every sampled pixel is opaque red.
    #[test]
    fn rasterizes_at_the_display_target() {
        let (w, h, rgba) = rasterize(RED_4X4, Some((8, 6))).expect("raster");
        assert_eq!((w, h), (8, 6), "stretched to the target, not the viewBox");
        assert!(
            rgba.chunks(4).all(|p| p == [0xFF, 0, 0, 0xFF]),
            "solid red fill"
        );
    }

    /// Without a target the SVG's own size wins.
    #[test]
    fn intrinsic_size_without_a_target() {
        let (w, h, _) = rasterize(RED_4X4, None).expect("raster");
        assert_eq!((w, h), (4, 4));
    }

    /// The uncovered part of the canvas is transparent — alpha survives the
    /// pipeline, unlike the opaque-by-construction DIB path.
    #[test]
    fn uncovered_canvas_is_transparent() {
        let svg = br##"<svg xmlns="http://www.w3.org/2000/svg" width="2" height="1"><rect width="1" height="1" fill="#00FF00"/></svg>"##;
        let (_, _, rgba) = rasterize(svg, None).expect("raster");
        assert_eq!(&rgba[0..4], &[0, 0xFF, 0, 0xFF], "covered pixel");
        assert_eq!(&rgba[4..8], &[0, 0, 0, 0], "uncovered pixel fully clear");
    }

    /// Garbage declines instead of panicking, and a degenerate target is
    /// clamped to one pixel.
    #[test]
    fn malformed_input_declines_and_zero_target_clamps() {
        assert!(rasterize(b"not svg at all", None).is_none());
        let (w, h, _) = rasterize(RED_4X4, Some((0, 0))).expect("raster");
        assert_eq!((w, h), (1, 1));
    }

    #[test]
    fn decode_svg_returns_a_premul_skia_image() {
        let image = decode_svg(RED_4X4, Some((8, 8))).expect("image");
        assert_eq!((image.width(), image.height()), (8, 8));
        assert_eq!(image.alpha_type(), AlphaType::Premul);
    }
}
