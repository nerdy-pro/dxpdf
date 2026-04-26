//! Text measurer — resolves text widths through a `FontRegistry`.

use std::cell::RefCell;
use std::collections::HashSet;

use skia_safe::Font;

use crate::render::dimension::Pt;
use crate::render::emoji::resolve::{EmojiFamily, EmojiResolver, EmojiTypeface, RegistryLookup};
use crate::render::emoji::shape::shape_text;
use crate::render::fonts::{self, FontRegistry, TypefaceEntry};

use super::fragment::FontProps;

/// Measures text using Skia fonts resolved through a [`FontRegistry`].
/// Holds a per-instance `FontCache` for `Font` reuse across measurements.
///
/// Also owns the [`EmojiResolver`] for the same render — so emoji typeface
/// lookups dedupe through the same per-instance state, alongside the warn-
/// once dedup set for `Unavailable` clusters.
pub struct TextMeasurer<'r> {
    registry: &'r FontRegistry,
    font_cache: RefCell<fonts::FontCache>,
    emoji_resolver: EmojiResolver<RegistryLookup<'r>>,
    /// Per-render dedup set so we warn at most once per cluster about a
    /// missing color emoji typeface.
    warned_emoji: RefCell<HashSet<String>>,
}

impl<'r> TextMeasurer<'r> {
    pub fn new(registry: &'r FontRegistry) -> Self {
        Self {
            registry,
            font_cache: RefCell::new(fonts::FontCache::new()),
            emoji_resolver: EmojiResolver::new(RegistryLookup { registry }),
            warned_emoji: RefCell::new(HashSet::new()),
        }
    }

    pub fn registry(&self) -> &'r FontRegistry {
        self.registry
    }

    /// Measure a text string with the given font properties.
    /// Returns (width, TextMetrics).
    pub fn measure(
        &self,
        text: &str,
        font_props: &FontProps,
    ) -> (Pt, super::fragment::TextMetrics) {
        let mut cache = self.font_cache.borrow_mut();
        let font = cache.get(
            self.registry,
            &font_props.family,
            font_props.size,
            font_props.bold,
            font_props.italic,
        );

        let (width, _bounds) = font.measure_str(text, None);
        let (_, metrics) = font.metrics();
        let text_metrics = super::fragment::TextMetrics {
            ascent: Pt::new(-metrics.ascent),
            descent: Pt::new(metrics.descent),
            leading: Pt::new(metrics.leading.max(0.0)),
        };

        // §17.3.2.35: include character spacing in the measured width
        // so line fitting accounts for the extra inter-character space.
        let char_count = text.chars().count();
        let spacing_extra = if char_count > 0 {
            font_props.char_spacing * (char_count as f32)
        } else {
            Pt::ZERO
        };

        (Pt::new(width) + spacing_extra, text_metrics)
    }

    /// Query font metrics for underline positioning.
    /// Returns (underline_position, underline_thickness) in points.
    /// Position is positive below baseline per Skia convention.
    pub fn underline_metrics(&self, font_props: &FontProps) -> (Pt, Pt) {
        let mut cache = self.font_cache.borrow_mut();
        let font = cache.get(
            self.registry,
            &font_props.family,
            font_props.size,
            font_props.bold,
            font_props.italic,
        );
        let (_, metrics) = font.metrics();
        // Skia: underline_position() returns a negative value (below baseline).
        // We negate to get a positive offset below the baseline.
        // If the font doesn't provide metrics, log a warning and use descent as fallback.
        let raw_pos = metrics.underline_position();
        let raw_thick = metrics.underline_thickness();
        if raw_pos.is_none() || raw_thick.is_none() {
            log::warn!(
                "font '{}' ({:?}) missing underline metrics, using descent as fallback",
                font_props.family,
                font_props.size
            );
        }
        let position = Pt::new(-raw_pos.unwrap_or(metrics.descent));
        // Thickness fallback: 1pt (smallest visible line at 72dpi).
        let thickness = Pt::new(raw_thick.unwrap_or(1.0));
        (position, thickness)
    }

    /// Get line height for the default font (used for empty paragraphs).
    /// §17.3.1.33: includes leading so Auto line spacing scales the full
    /// font-recommended height.
    pub fn default_line_height(&self, family: &str, size: Pt) -> Pt {
        let mut cache = self.font_cache.borrow_mut();
        let font = cache.get(self.registry, family, size, false, false);
        let (_, metrics) = font.metrics();
        Pt::new(-metrics.ascent + metrics.descent + metrics.leading.max(0.0))
    }

    // ─── Emoji pipeline integration ────────────────────────────────────────

    /// Resolve a color emoji typeface via the per-render [`EmojiResolver`].
    /// Cached: repeat calls with the same `requested` family are O(1).
    pub fn resolve_emoji(&self, requested: Option<EmojiFamily>) -> EmojiTypeface {
        self.emoji_resolver.resolve(requested)
    }

    /// Measure a cluster directly against a resolved [`TypefaceEntry`],
    /// bypassing the family-name lookup path. Used by the emoji pipeline,
    /// which has already resolved the typeface and needs Skia raster metrics
    /// at the cluster's font size.
    ///
    /// **Shapes via rustybuzz** (GSUB-aware) so multi-codepoint emoji
    /// sequences measure to their *ligated* width, matching what the
    /// rasterizer produces at paint time. Without this, the layout would
    /// reserve `n × glyph_advance` for an `n`-codepoint sequence (cmap-
    /// only) but the rasterizer would draw a ligated single glyph that's
    /// narrower — the painter then stretches the image to fill the over-
    /// sized rect, distorting the emoji.
    ///
    /// Falls back to `font.measure_str` (cmap-only) if shaping fails — a
    /// best-effort policy mirroring the rasterizer's fallback path.
    pub fn measure_with_typeface(
        &self,
        text: &str,
        typeface: &TypefaceEntry,
        size: Pt,
    ) -> (Pt, super::fragment::TextMetrics) {
        let font = Font::from_typeface(typeface.typeface.clone(), f32::from(size));
        let (_, metrics) = font.metrics();
        let text_metrics = super::fragment::TextMetrics {
            ascent: Pt::new(-metrics.ascent),
            descent: Pt::new(metrics.descent),
            leading: Pt::new(metrics.leading.max(0.0)),
        };

        // Try the GSUB-aware advance first; fall back to the cmap-only
        // path if the typeface bytes can't be extracted or shaped.
        let advance = typeface
            .typeface
            .to_font_data()
            .and_then(|(bytes, _)| shape_text(&bytes, text, f32::from(size)).ok())
            .map(|run| run.total_advance)
            .unwrap_or_else(|| Pt::new(font.measure_str(text, None).0));

        (advance, text_metrics)
    }

    /// Log a warning once per cluster when no color emoji typeface is
    /// available on the host. The `attempted` list lets operators know
    /// which packages to install (e.g. `fonts-noto-color-emoji` on Debian).
    pub fn warn_emoji_unavailable_once(&self, cluster: &str, attempted: &[EmojiFamily]) {
        let inserted = self.warned_emoji.borrow_mut().insert(cluster.to_string());
        if inserted {
            log::warn!(
                "no color emoji typeface available for cluster {:?}; \
                 tried {:?}. Install a color emoji font on the host \
                 (e.g. fonts-noto-color-emoji) to render this cluster correctly.",
                cluster,
                attempted
            );
        }
    }
}
