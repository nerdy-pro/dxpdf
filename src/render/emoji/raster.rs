//! Emoji cluster rasterization with per-render cache.
//!
//! Skia's PDF backend cannot emit color glyph tables (COLR/CPAL, CBDT/CBLC,
//! sbix, SVG-in-OT) — but its raster backend honors all four. We rasterize
//! emoji clusters onto an offscreen surface using the raster backend, snapshot
//! to an [`Image`], and let the painter embed it in the PDF at the run's
//! typographic position.
//!
//! The cache key includes the cluster text (NFC-normalized per UAX #15), the
//! typeface id, the requested point size, and the super-sample factor;
//! identical inputs yield a single rasterization shared across the document.

use std::collections::HashMap;
use std::rc::Rc;

use skia_safe::{surfaces, Color, Font, Image, Paint, PaintStyle, Point};
use unicode_normalization::UnicodeNormalization;

use crate::render::dimension::Pt;
use crate::render::emoji::cluster::EmojiCluster;
use crate::render::emoji::shape::shape_text;
use crate::render::fonts::{TypefaceEntry, TypefaceId};
use crate::render::geometry::PtSize;

// ─── Public ADTs ─────────────────────────────────────────────────────────────

/// Pixel density at which clusters are rasterized.
///
/// PDF viewers re-rasterize images at the user's chosen zoom; super-sampling
/// here trades larger PDF size for crispness when zoomed in. For sbix /
/// CBDT bitmap-emoji fonts (Apple Color Emoji, Noto Color Emoji), higher
/// super-sampling also lets Skia pick a higher-resolution source bitmap
/// from the font's strike table, sharpening the result before any paint-
/// time downsampling.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SuperSample {
    /// 1 pixel per Pt — minimum size, soft when zoomed.
    OnePerPt,
    /// 2 pixels per Pt — soft at print quality.
    TwoPerPt,
    /// 3 pixels per Pt — moderate.
    ThreePerPt,
    /// 4 pixels per Pt — default. Drives Skia to pick a 64–96px sbix
    /// strike at typical body-text sizes, which downsamples cleanly via
    /// Mitchell cubic at paint time.
    FourPerPt,
    /// 6 pixels per Pt — print-quality / display-zoom-friendly.
    SixPerPt,
}

impl SuperSample {
    pub const fn factor(self) -> f32 {
        match self {
            SuperSample::OnePerPt => 1.0,
            SuperSample::TwoPerPt => 2.0,
            SuperSample::ThreePerPt => 3.0,
            SuperSample::FourPerPt => 4.0,
            SuperSample::SixPerPt => 6.0,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RasterConfig {
    pub super_sample: SuperSample,
}

impl Default for RasterConfig {
    fn default() -> Self {
        Self {
            super_sample: SuperSample::FourPerPt,
        }
    }
}

/// Cache key for a rasterized cluster.
///
/// Cluster text is NFC-normalized (UAX #15) so canonically-equivalent inputs
/// share a slot. Size, scale, and target dimensions are stored as the bit
/// pattern of the f32 so the key is hashable and comparison is exact (no
/// rounding-induced misses).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct EmojiKey {
    pub cluster: String,
    pub typeface_id: TypefaceId,
    pub size_bits: u32,
    pub scale_bits: u32,
    /// Target image width in Pt as f32 bits. The rasterizer guarantees
    /// image_aspect == rect_aspect to prevent anisotropic stretching at
    /// paint time, so the cache key includes the target dimensions.
    pub target_w_bits: u32,
    pub target_h_bits: u32,
}

impl EmojiKey {
    pub fn new(
        text: &str,
        typeface: &TypefaceEntry,
        size: Pt,
        scale: SuperSample,
        target: PtSize,
    ) -> Self {
        Self {
            cluster: text.nfc().collect(),
            typeface_id: TypefaceId::from(&typeface.typeface),
            size_bits: f32::from(size).to_bits(),
            scale_bits: scale.factor().to_bits(),
            target_w_bits: target.width.raw().to_bits(),
            target_h_bits: target.height.raw().to_bits(),
        }
    }
}

/// A rasterized emoji image plus the metadata needed to place it at paint
/// time at the run's baseline.
#[derive(Clone, Debug)]
pub struct EmojiImage {
    /// Skia image snapshot. Cheap to clone (reference-counted internally).
    pub image: Image,
    /// Pixel dimensions of the underlying surface (width, height).
    pub pixels: (i32, i32),
    /// The size at which to draw the image in the PDF, in original Pt units
    /// (i.e. de-scaled from the super-sampled raster).
    pub draw_size: PtSize,
    /// Distance from the run's baseline to the top of `draw_size`, in Pt.
    /// Positive values mean the top sits above the baseline (the typical
    /// case for an emoji whose bounds lie above the baseline).
    pub baseline_offset: Pt,
}

// ─── Rasterizer ──────────────────────────────────────────────────────────────

/// Per-render rasterizer that owns the cache. Lifetime equals the painter's.
///
/// Maintains two caches:
/// 1. `cache` — the rasterized image keyed by [`EmojiKey`].
/// 2. `font_bytes` — typeface bytes by id, so a 190 MB Apple Color Emoji
///    typeface isn't re-extracted via `to_font_data` for every cluster.
pub struct EmojiRasterizer {
    config: RasterConfig,
    cache: HashMap<EmojiKey, EmojiImage>,
    font_bytes: HashMap<TypefaceId, Rc<Vec<u8>>>,
}

impl Default for EmojiRasterizer {
    fn default() -> Self {
        Self::new(RasterConfig::default())
    }
}

impl EmojiRasterizer {
    pub fn new(config: RasterConfig) -> Self {
        Self {
            config,
            cache: HashMap::new(),
            font_bytes: HashMap::new(),
        }
    }

    pub fn config(&self) -> RasterConfig {
        self.config
    }

    pub fn cached_count(&self) -> usize {
        self.cache.len()
    }

    /// Rasterize `cluster` at `size` using `typeface`, or return the cached
    /// image if previously seen.
    ///
    /// `target` is the layout's reserved rect (in Pt). The rasterizer
    /// allocates an image with **the same aspect ratio** as `target`,
    /// scaled by the super-sample factor — this is critical because
    /// `Canvas::draw_image_rect` does anisotropic scaling when image
    /// aspect ≠ rect aspect, distorting the emoji. By matching aspects
    /// here, the painter's image-to-rect scaling becomes uniform and
    /// the emoji's visual content is preserved.
    ///
    /// Internally shapes via `rustybuzz` (GSUB-aware) so multi-codepoint
    /// emoji sequences (keycap, modifier, ZWJ, RIS) render as their
    /// ligated single glyph — `canvas.draw_str` would have rendered each
    /// codepoint separately. See `shape.rs` for the shaper.
    ///
    /// `typeface` is guaranteed by the type system to be a real
    /// [`TypefaceEntry`] — callers that hold an [`EmojiTypeface::Unavailable`]
    /// cannot reach this method. (See plan test X8.)
    ///
    /// [`EmojiTypeface::Unavailable`]: super::resolve::EmojiTypeface::Unavailable
    pub fn rasterize(
        &mut self,
        cluster: &EmojiCluster,
        typeface: &TypefaceEntry,
        size: Pt,
        target: PtSize,
    ) -> &EmojiImage {
        let scale = self.config.super_sample;
        let key = EmojiKey::new(cluster.text, typeface, size, scale, target);
        if !self.cache.contains_key(&key) {
            let bytes = self.font_bytes_for(typeface);
            let image = rasterize_uncached(
                cluster.text,
                typeface,
                size,
                scale,
                target,
                bytes.as_deref().map(|v| v.as_slice()),
            );
            self.cache.insert(key.clone(), image);
        }
        self.cache.get(&key).expect("just inserted")
    }

    /// Per-render typeface byte cache. Apple Color Emoji is ~190 MB on
    /// macOS, so we extract once per typeface and reuse for every
    /// rasterization. Returns `None` (not `Err`) if the typeface refuses
    /// to expose bytes — the rasterizer falls back to the cmap-only
    /// `draw_str` path.
    fn font_bytes_for(&mut self, typeface: &TypefaceEntry) -> Option<Rc<Vec<u8>>> {
        let id = TypefaceId::from(&typeface.typeface);
        if let Some(bytes) = self.font_bytes.get(&id) {
            return Some(bytes.clone());
        }
        let bytes = typeface.typeface.to_font_data().map(|(b, _)| Rc::new(b))?;
        self.font_bytes.insert(id, bytes.clone());
        Some(bytes)
    }
}

fn rasterize_uncached(
    text: &str,
    typeface: &TypefaceEntry,
    size: Pt,
    scale: SuperSample,
    target: PtSize,
    font_bytes: Option<&[u8]>,
) -> EmojiImage {
    let factor = scale.factor();
    let scaled_size = f32::from(size) * factor;
    let font = Font::from_typeface(typeface.typeface.clone(), scaled_size);

    // Image dimensions are derived from the target rect (× scale). This
    // guarantees image_aspect == target_aspect → uniform scaling at paint
    // time. Anisotropic scaling would distort the emoji (a square keycap
    // squished to a rectangle). Apple Color Emoji's font.metrics() are
    // non-linear across point sizes (ascent+descent ratio = 1.64 at 11pt
    // but 1.37 at 22pt), so deriving image height from the rasterizer's
    // own metrics() would mismatch the layout's rect.
    let width_px = (target.width.raw() * factor).ceil().max(1.0) as i32;
    let height_px = (target.height.raw() * factor).ceil().max(1.0) as i32;

    // Baseline within the surface: the layout reserves space using the
    // *original*-size ascent (via `TextMeasurer::measure_with_typeface`,
    // which calls `font.metrics()` at the run's font size). The painter
    // then places the rect with `top_y = baseline_y - metrics.ascent`.
    // We must therefore position the glyph baseline within the image at
    // `original_ascent × factor`, NOT the scaled-size font's own ascent —
    // for fonts with non-linear metrics (Apple Color Emoji's ascent+descent
    // ratio is 1.64 at 11pt but 1.37 at 22pt), the two differ and the
    // emoji ends up floating above or below the line's baseline.
    let original_font = Font::from_typeface(typeface.typeface.clone(), f32::from(size));
    let (_, original_metrics) = original_font.metrics();
    let baseline_y_px = -original_metrics.ascent * factor;

    // Try the GSUB-aware shaping path. On any shaping failure (font bytes
    // unavailable, parse error, glyph id out of range), fall through to
    // the cmap-only `draw_str` path so the rasterizer still produces
    // output.
    let shaped = font_bytes.and_then(|b| shape_text(b, text, scaled_size).ok());

    let mut surface = surfaces::raster_n32_premul((width_px, height_px))
        .expect("raster_n32_premul returned None for non-degenerate dimensions");
    let canvas = surface.canvas();

    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_style(PaintStyle::Fill);
    // Color emoji fonts (COLR / CBDT / sbix) ignore the paint color and use
    // their internal palette. Monochrome emoji (e.g. Noto Emoji) honour it.
    // Black is the safest default for monochrome fallbacks.
    paint.set_color(Color::BLACK);

    match shaped {
        Some(run) => {
            // Walk shaped glyphs, accumulating positions. Baseline at
            // `baseline_y_px` (= layout's ascent at original size, scaled
            // by `factor`) from the top; each glyph's HarfBuzz y-offset
            // is positive-up, so we negate for Skia's y-down.
            let mut ids = Vec::with_capacity(run.glyphs.len());
            let mut positions = Vec::with_capacity(run.glyphs.len());
            let mut pen_x = 0.0f32;
            for g in &run.glyphs {
                ids.push(g.id);
                positions.push(Point::new(
                    pen_x + g.x_offset.raw(),
                    baseline_y_px - g.y_offset.raw(),
                ));
                pen_x += g.advance.raw();
            }
            canvas.draw_glyphs_at(&ids, &*positions, (0.0, 0.0), &font, &paint);
        }
        None => {
            // Fallback: cmap-level draw_str. The bounds-based translation
            // here is a best effort to land the glyph inside the surface.
            let (_, bounds) = font.measure_str(text, None);
            canvas.translate((-bounds.left(), -bounds.top()));
            canvas.draw_str(text, (0.0, 0.0), &font, &paint);
        }
    }

    let image = surface.image_snapshot();

    // The image dimensions exactly match `target × factor` (modulo ceil),
    // so draw_size returns to `target`. The painter draws `image` into a
    // rect of exactly these dimensions for uniform scaling.
    EmojiImage {
        image,
        pixels: (width_px, height_px),
        draw_size: target,
        baseline_offset: Pt::new(baseline_y_px / factor),
    }
}

// ─── Tests (X1–X6 from docs/emoji-rendering.md) ──────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::emoji::cluster::{EmojiPresentation, EmojiStructure};
    use crate::render::emoji::resolve::{resolve, EmojiTypeface, RegistryLookup};
    use crate::render::fonts::{FontRegistry, TypefaceOrigin};
    use skia_safe::{FontMgr, FontStyle};

    /// Construct a real `TypefaceEntry` from any host-default font. Used by
    /// the cache-shape tests that don't care about the actual glyphs.
    fn any_typeface() -> TypefaceEntry {
        let mgr = FontMgr::new();
        let tf = mgr
            .legacy_make_typeface(None::<&str>, FontStyle::normal())
            .expect("system has no default typeface — cannot run test");
        let id = TypefaceId::from(&tf);
        TypefaceEntry {
            typeface: tf,
            origin: TypefaceOrigin::System { typeface_id: id },
        }
    }

    fn single_emoji(text: &'static str) -> EmojiCluster<'static> {
        EmojiCluster {
            text,
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::Single,
        }
    }

    /// Default target rect for tests — non-degenerate, sized at the
    /// 12pt font size used throughout the test suite. Aspect 1:1.6 so
    /// distortion-style assertions can be checked.
    fn default_target() -> PtSize {
        PtSize::new(Pt::new(12.0), Pt::new(18.0))
    }

    /// X1 — same key twice → one cache entry.
    #[test]
    fn x1_same_input_dedupes_in_cache() {
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        let c = single_emoji("\u{1F4DE}");
        let _ = r.rasterize(&c, &tf, Pt::new(12.0), default_target());
        let _ = r.rasterize(&c, &tf, Pt::new(12.0), default_target());
        assert_eq!(r.cached_count(), 1, "identical key must reuse cache slot");
    }

    /// X2 — different cluster text → distinct entries.
    #[test]
    fn x2_distinct_clusters_cache_independently() {
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        let _ = r.rasterize(
            &single_emoji("\u{1F4DE}"),
            &tf,
            Pt::new(12.0),
            default_target(),
        );
        let _ = r.rasterize(
            &single_emoji("\u{1F4E7}"),
            &tf,
            Pt::new(12.0),
            default_target(),
        );
        assert_eq!(r.cached_count(), 2);
    }

    /// X3 — different size → distinct entries.
    #[test]
    fn x3_distinct_sizes_cache_independently() {
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        let c = single_emoji("\u{1F4DE}");
        let _ = r.rasterize(&c, &tf, Pt::new(12.0), default_target());
        let _ = r.rasterize(&c, &tf, Pt::new(24.0), default_target());
        assert_eq!(r.cached_count(), 2);
    }

    /// X4 — pixel dimensions are always at least 1×1.
    #[test]
    fn x4_pixel_dimensions_non_degenerate() {
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        let img = r
            .rasterize(
                &single_emoji("\u{1F4DE}"),
                &tf,
                Pt::new(12.0),
                default_target(),
            )
            .clone();
        assert!(
            img.pixels.0 >= 1,
            "width must be >= 1 px, got {}",
            img.pixels.0
        );
        assert!(
            img.pixels.1 >= 1,
            "height must be >= 1 px, got {}",
            img.pixels.1
        );
        assert!(img.draw_size.width.raw() > 0.0);
        assert!(img.draw_size.height.raw() > 0.0);
    }

    /// X4b — degenerate empty input must not crash. Image dimensions are
    /// governed by the target rect now (so the image aspect matches the
    /// painter's destination rect — see Y_aspect below), so we only
    /// assert non-degeneracy.
    #[test]
    fn x4b_zero_width_input_yields_non_degenerate_surface() {
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        let img = r
            .rasterize(&single_emoji(""), &tf, Pt::new(12.0), default_target())
            .clone();
        assert!(img.pixels.0 >= 1);
        assert!(img.pixels.1 >= 1);
    }

    /// Y_aspect — image surface aspect == target rect aspect. This is
    /// the property that prevents `Canvas::draw_image_rect` from
    /// stretching the emoji at paint time. Without it, fonts whose
    /// `ascent + descent` doesn't scale linearly (Apple Color Emoji's
    /// ratio is 1.64 at 11pt vs 1.37 at 22pt) produce images of one
    /// aspect that get drawn into rects of a different aspect →
    /// distortion.
    #[test]
    fn y_aspect_image_matches_target() {
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        // Pick an asymmetric target so a regression — using ascent+descent
        // for height — would obviously change the aspect.
        let target = PtSize::new(Pt::new(11.0), Pt::new(18.0));
        let img = r
            .rasterize(&single_emoji("A"), &tf, Pt::new(11.0), target)
            .clone();
        let img_aspect = img.pixels.0 as f32 / img.pixels.1 as f32;
        let target_aspect = target.width.raw() / target.height.raw();
        // Within rounding (ceil + integer pixels) — within 5% of the
        // target aspect.
        let rel_err = (img_aspect - target_aspect).abs() / target_aspect;
        assert!(
            rel_err < 0.05,
            "image aspect {img_aspect:.4} must match target aspect {target_aspect:.4} \
             within rounding (rel err {rel_err:.4})"
        );
    }

    /// X5 — rasterization of a renderable glyph produces non-trivial pixel
    /// data. Skipped on hosts where no color emoji typeface resolves; we
    /// don't bundle fonts, so CI without one passes via a clean skip.
    #[test]
    fn x5_rasterized_image_has_visible_pixels() {
        let registry = FontRegistry::new(FontMgr::new());
        let lookup = RegistryLookup {
            registry: &registry,
        };
        let resolved = resolve(&lookup, None);
        let entry = match resolved {
            EmojiTypeface::Resolved { entry, .. } => entry,
            EmojiTypeface::Unavailable { .. } => {
                eprintln!("skipping X5: no color emoji typeface on this host");
                return;
            }
        };
        let mut r = EmojiRasterizer::default();
        let img = r
            .rasterize(
                &single_emoji("\u{1F4DE}"),
                &entry,
                Pt::new(24.0),
                PtSize::new(Pt::new(24.0), Pt::new(36.0)),
            )
            .clone();
        let peek = img.image.peek_pixels();
        // peek_pixels can return None if the image is GPU-backed; raster
        // images always succeed.
        let pixels = peek.expect("raster image must expose pixel data");
        let bytes = pixels.bytes().expect("RGBA pixel data must be readable");
        assert!(
            bytes.iter().any(|&b| b != 0),
            "rendered emoji must contain at least one non-zero pixel"
        );
    }

    /// X6 — NFC-different but canonically-equivalent inputs share a cache
    /// slot. "é" can be either U+00E9 (precomposed) or U+0065 + U+0301
    /// (combining acute). Both NFC-normalize to U+00E9.
    #[test]
    fn x6_canonically_equivalent_inputs_share_cache() {
        // Note: "é" alone is not classified as emoji by `cluster::classify`
        // (no Emoji property), but the rasterizer doesn't care — the cache
        // key is built from raw text. We exercise the NFC path with two
        // canonically-equivalent representations.
        let mut r = EmojiRasterizer::default();
        let tf = any_typeface();
        let precomposed = EmojiCluster {
            text: "\u{00E9}",
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::Single,
        };
        let decomposed = EmojiCluster {
            text: "e\u{0301}",
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::Single,
        };
        let _ = r.rasterize(&precomposed, &tf, Pt::new(12.0), default_target());
        let _ = r.rasterize(&decomposed, &tf, Pt::new(12.0), default_target());
        assert_eq!(
            r.cached_count(),
            1,
            "NFC-equivalent inputs must share a cache slot"
        );
    }

    // ─── Shape invariants ─────────────────────────────────────────────────

    /// SuperSample.factor() is monotonically increasing. Sanity-check the
    /// enum values so a future refactor that swaps factors gets caught.
    #[test]
    fn super_sample_factors_monotonic() {
        assert!(SuperSample::OnePerPt.factor() < SuperSample::TwoPerPt.factor());
        assert!(SuperSample::TwoPerPt.factor() < SuperSample::ThreePerPt.factor());
    }

    /// Pixel surface scales with super-sample factor for outline glyphs.
    /// Use a plain ASCII letter so the test is independent of bitmap-only
    /// (sbix/CBDT) emoji glyph quantization, which selects different
    /// pre-rendered strike sizes at different requested point sizes.
    #[test]
    fn pixel_dimensions_scale_with_super_sample_for_outline_glyphs() {
        let tf = any_typeface();
        let c = single_emoji("A");
        let size = Pt::new(12.0);

        let mut r1 = EmojiRasterizer::new(RasterConfig {
            super_sample: SuperSample::OnePerPt,
        });
        let img1 = r1.rasterize(&c, &tf, size, default_target()).clone();

        let mut r3 = EmojiRasterizer::new(RasterConfig {
            super_sample: SuperSample::ThreePerPt,
        });
        let img3 = r3.rasterize(&c, &tf, size, default_target()).clone();

        // Outline glyphs scale linearly: at 3× super-sample the pixel
        // surface should be ~3× larger on each axis.
        assert!(
            img3.pixels.0 >= img1.pixels.0 * 2,
            "3× super-sample width must be at least 2× the 1× width"
        );
        assert!(
            img3.pixels.1 >= img1.pixels.1 * 2,
            "3× super-sample height must be at least 2× the 1× height"
        );

        // Draw size: should match within ceil-then-divide rounding (≤ 2pt).
        let dw1 = img1.draw_size.width.raw();
        let dw3 = img3.draw_size.width.raw();
        assert!(
            (dw1 - dw3).abs() <= 2.0,
            "outline glyph draw widths must match within rounding, got {dw1} vs {dw3}"
        );
    }
}
