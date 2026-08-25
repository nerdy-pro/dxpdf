//! Paint phase — iterate DrawCommands and emit Skia PDF canvas operations.

use std::collections::HashMap;

use rustc_hash::{FxHashMap, FxHashSet};
use skia_safe::{
    canvas::SrcRectConstraint, path_effect::PathEffect, pdf, BlurStyle, Color4f, Data, MaskFilter,
    Paint, Path, PathBuilder, PathFillType, TextBlob,
};

use crate::render::dimension::Pt;
use crate::render::emoji::raster::EmojiRasterizer;
use crate::render::error::RenderError;
use crate::render::fonts::{self, FontRegistry};
use crate::render::layout::draw_command::{
    DrawCommand, LayoutedPage, OutlineMark, ResolvedDashPattern, ResolvedEffect, ResolvedFill,
    ResolvedLineCap, ResolvedLineJoin, ResolvedStroke,
};
use crate::render::resolve::drawing_color::Rgba;
use crate::render::resolve::images::MediaEntry;
use crate::render::resolve::shape_geometry::{PathVerb, SubPath};
use crate::render::skia_conv::{to_color4f, to_line, to_point, to_rect, to_size};
use crate::render::spacing;
use crate::render::RenderOptions;

/// PDF user-space units per inch — the fixed 1 pt = 1/72 in conversion used to
/// turn a display size in points into a pixel target at a given image DPI.
///
/// The target image DPI itself is a render-time knob (`RenderOptions::image_dpi`,
/// default 220) threaded in from the caller. Higher values embed more source
/// pixels, which matters because Skia's PDF backend does not emit
/// `/Interpolate true` on image dicts, so viewers smooth-scale with
/// nearest-neighbor and need enough pixels to absorb the zoom.
const POINTS_PER_INCH: f32 = 72.0;

/// Quantized fractional `srcRect` crop, part of the cache key. Each of
/// (origin.x, origin.y, size.w, size.h) is scaled by the OOXML
/// thousandth-percent resolution (100000 = 100%) and rounded, so two crops that
/// are spec-identical hash alike. `None` marks an uncropped placement.
type CropKey = Option<(i32, i32, i32, i32)>;

/// Decoded-image cache, keyed by media data-pointer identity, the downsample
/// target size, **and** the crop baked into the cached bitmap. Two placements
/// of the same media get distinct entries when they draw at different sizes or
/// under different `srcRect` crops — each caches its own correctly-scaled,
/// pre-cropped bitmap, so neither reuses the other's resolution or crop.
type ImageCacheKey = (*const [u8], i32, i32, CropKey);
type ImageCache = HashMap<ImageCacheKey, skia_safe::Image>;

/// Downsample target dimensions (in pixels) for a display rect at `image_dpi`.
/// Shared by the image cache key and the downsample passes so both agree on the
/// resolution a placement needs. `image_dpi` is assumed sanitized (positive and
/// finite) by [`RenderOptions`](crate::render::RenderOptions).
fn downsample_target(rect: crate::render::geometry::PtRect, image_dpi: f32) -> (i32, i32) {
    let scale = image_dpi / POINTS_PER_INCH;
    (
        (rect.size.width.raw() * scale).ceil() as i32,
        (rect.size.height.raw() * scale).ceil() as i32,
    )
}

/// Quantize a fractional `srcRect` crop for the cache key, at the OOXML
/// thousandth-percent resolution so spec-identical crops map to one key.
fn quantize_crop(crop: &crate::render::geometry::PtRect) -> (i32, i32, i32, i32) {
    let q = |v: f32| (v * 100_000.0).round() as i32;
    (
        q(crop.origin.x.raw()),
        q(crop.origin.y.raw()),
        q(crop.size.width.raw()),
        q(crop.size.height.raw()),
    )
}

/// §17.3.1.19: build the PDF structure tree the document outline is derived
/// from, or `None` when the document has no headings.
///
/// **The tree is deliberately flat.** Skia derives the outline hierarchy from
/// the heading *level digit* in each node's type string, not from the shape of
/// this tree: nesting an `H2` node inside an `H1` node does not nest the
/// outline entry, it concatenates the two titles into one entry. Measured
/// against Skia **m145**, the pinned build: `SkPDFTag.cpp` on Skia's `main`
/// branch describes different behaviour (`main`'s public node has a `fTitle`
/// field this version doesn't, for one), so do not "simplify" toward it. The
/// obvious "mirror the heading hierarchy" design silently merges every
/// subheading into its parent.
///
/// The title comes from `set_alt`, and must: this Skia has no `fTitle` field on
/// the public node, so a node without an alt contributes **no outline entry at
/// all** rather than an untitled one.
fn build_outline_tree(pages: &[LayoutedPage]) -> Option<pdf::StructureElementNode<'static>> {
    let mut root = pdf::StructureElementNode::new("Document");
    let mut any = false;

    for page in pages {
        for cmd in &page.commands {
            let DrawCommand::Outline(OutlineMark::Begin(heading)) = cmd else {
                continue;
            };
            let mut node = pdf::StructureElementNode::new(heading_type_string(heading.level));
            node.set_node_id(heading.node_id);
            node.set_alt(&*heading.title);
            root.append_child(node);
            any = true;
        }
    }

    any.then_some(root)
}

/// The PDF structure type for a §17.3.1.19 outline level.
///
/// ISO 32000-1 (PDF 1.7) defines heading structure types `H1`–`H6` only, and
/// Skia enforces that range, while §17.3.1.19 allows nine levels. Levels 7–9
/// clamp to `H6`: the entry survives — dropping a heading outright is a worse
/// answer than a flattened one — and the hierarchy stops distinguishing depth
/// there. Warned once per render, because a document that relies on nine levels
/// should say so in the log rather than quietly lose its shape.
fn heading_type_string(level: crate::model::OutlineLevel) -> String {
    const DEEPEST_PDF_HEADING: u8 = 6;
    let level = level.value();
    if level > DEEPEST_PDF_HEADING {
        log::warn!(
            "outline: §17.3.1.19 level {level} has no PDF structure type \
             (ISO 32000-1 stops at H{DEEPEST_PDF_HEADING}) — clamping, so the \
             outline flattens below that depth"
        );
    }
    format!("H{}", level.min(DEEPEST_PDF_HEADING))
}

/// Render laid-out pages to PDF bytes via Skia.
///
/// `registry` owns the typeface universe for this render — paint resolves
/// every text run through it so any subsetted typefaces (swapped in by
/// `subset::apply` between layout and paint) are picked up correctly.
///
/// `options` carries the paint-phase knobs (currently the target image DPI; see
/// [`RenderOptions`]). Taking the struct rather than a raw `f32` keeps this
/// public entry point safe by construction — a `RenderOptions` can only hold a
/// sanitized DPI — and lets future knobs flow through without a signature change.
pub fn render_to_pdf(
    pages: &[LayoutedPage],
    registry: &FontRegistry,
    options: &RenderOptions,
) -> Result<Vec<u8>, RenderError> {
    let mut pdf_bytes: Vec<u8> = Vec::new();
    // §17.3.1.19: the outline is document-level, so its structure tree has to
    // exist before the first page — `pdf::new_document` takes it by reference
    // and holds it for the document's lifetime. Layout has already finished, so
    // the headings are all knowable here.
    let structure_tree = build_outline_tree(pages);
    let pdf_metadata = pdf::Metadata {
        encoding_quality: Some(85),
        // A document with no headings gets neither, so it keeps its current
        // output byte for byte: an empty `/Outlines`, or a `/StructTreeRoot`
        // for a document that has no structure to describe, is worse than none.
        //
        // `pdf::Outline` has a third value, `StructureElements`, which mirrors
        // the whole structure tree with titles taken from each node's *type
        // string* (`"Document"`, `"H1"`, …) rather than its heading text —
        // useless as a reader-facing outline, so `StructureElementHeaders`
        // (H1-H6 only, titled from `set_alt`) is the one used here.
        outline: if structure_tree.is_some() {
            pdf::Outline::StructureElementHeaders
        } else {
            pdf::Outline::None
        },
        structure_element_tree_root: structure_tree,
        ..Default::default()
    };
    let mut doc = pdf::new_document(&mut pdf_bytes, Some(&pdf_metadata));
    let mut state = PaintState::new();
    for page in pages {
        let mut on_page = doc.begin_page(to_size(page.page_size), None);
        {
            let canvas = on_page.canvas();
            render_page(canvas, page, registry, &mut state, options);
        }
        doc = on_page.end_page();
    }

    doc.close();
    Ok(pdf_bytes)
}

/// Mutable per-render caches, shared across every page and bundled into one
/// struct so `render_page` stays under clippy's argument-count limit and future
/// paint state can be added without re-threading each call site.
struct PaintState {
    /// Font slot / typeface cache (see [`fonts::FontCache`]).
    font_cache: fonts::FontCache,
    /// Decoded Skia images across pages — avoids re-decoding the same bytes on
    /// every page (e.g. a logo repeated in headers/footers). See [`ImageCache`].
    image_cache: ImageCache,
    /// Emoji rasterizer — clusters that recur across pages (footer 📞 etc.) are
    /// rasterized once and shared.
    emoji_rasterizer: EmojiRasterizer,
    /// Position-independent glyph runs, cached per (font slot, text). `draw_str`
    /// rebuilds this run on every call; reusing a `TextBlob` skips that remap for
    /// words repeated across the document.
    blob_cache: FxHashMap<usize, FxHashMap<Box<str>, TextBlob>>,
    /// GSUB-aware shaper for runs in a joining script (issue #131), built once
    /// per render like the caches around it. `None` if this Skia build exposes
    /// no HarfBuzz shaper, in which case such runs fall back to `draw_str` —
    /// unjoined, which is what every run looked like before #131.
    shaper: Option<crate::render::shape::Shaper>,
    /// Tier-0 features already reported, so an unsupported fill in a header
    /// warns once per render rather than once per page. Per-render, like
    /// `warned_emoji` in `layout::measurer` and `warned_border_styles` in
    /// `layout::build::convert` — never process-wide.
    warned: FxHashSet<&'static str>,
}

impl PaintState {
    fn new() -> Self {
        Self {
            font_cache: fonts::FontCache::new(),
            image_cache: HashMap::new(),
            emoji_rasterizer: EmojiRasterizer::default(),
            blob_cache: FxHashMap::default(),
            shaper: crate::render::shape::Shaper::new().ok(),
            warned: FxHashSet::default(),
        }
    }
}

/// Draw one run through HarfBuzz, returning whether it drew anything.
///
/// `false` means the caller must fall back to the cmap path: either this Skia
/// build has no shaper, or shaping produced no glyphs. Neither is a programming
/// error — the run is then painted the way every run was painted before issue
/// #131, which for a joining script is unjoined but present.
///
/// Positions come back as absolute offsets from the run origin, which is what
/// [`Canvas::draw_glyphs_at`] takes — the same convention `render::emoji`'s
/// rasterizer relies on, and the reason neither accumulates a pen.
fn draw_shaped(
    canvas: &skia_safe::Canvas,
    shaper: Option<&crate::render::shape::Shaper>,
    font: &skia_safe::Font,
    text: &str,
    direction: crate::render::shape::RunDirection,
    position: crate::render::geometry::PtOffset,
    paint: &Paint,
) -> bool {
    let Some(shaper) = shaper else {
        return false;
    };
    let Ok(run) = shaper.shape(&font.typeface(), text, font.size(), direction) else {
        return false;
    };
    if run.glyphs.is_empty() {
        return false;
    }
    let ids: Vec<skia_safe::GlyphId> = run.glyphs.iter().map(|g| g.id).collect();
    let offsets: Vec<skia_safe::Point> = run
        .glyphs
        .iter()
        .map(|g| skia_safe::Point::new(f32::from(g.x), f32::from(g.y)))
        .collect();
    canvas.draw_glyphs_at(&ids, &*offsets, to_point(position), font, paint);
    true
}

fn render_page(
    canvas: &skia_safe::Canvas,
    page: &LayoutedPage,
    registry: &FontRegistry,
    state: &mut PaintState,
    options: &RenderOptions,
) {
    // Destructure into disjoint `&mut` field bindings — same borrow semantics as
    // the individual caches, so the paint body below is unchanged.
    let PaintState {
        font_cache,
        image_cache,
        emoji_rasterizer,
        blob_cache,
        shaper,
        warned,
    } = state;
    let image_dpi = options.image_dpi();

    // Reusable paints — built once per page instead of per draw command (a
    // large doc emits 100k+). Each carries the fixed config for its purpose;
    // only the fields that vary (color, stroke width) are set per command, so
    // no state leaks between commands. `default_paint` is never mutated — it
    // backs the image/emoji `draw_image_rect` calls.
    let mut text_paint = Paint::default();
    text_paint.set_anti_alias(true);
    let mut stroke_paint = Paint::default();
    stroke_paint.set_anti_alias(true);
    stroke_paint.set_stroke(true);
    let mut rect_paint = Paint::default();
    rect_paint.set_anti_alias(false);
    let default_paint = Paint::default();

    for cmd in &page.commands {
        match cmd {
            // §17.3.1.19: marked-content boundaries — see `OutlineMark`'s doc
            // for why the pair matters even though it doesn't change any
            // observable destination.
            DrawCommand::Outline(OutlineMark::Begin(heading)) => {
                pdf::set_node_id(canvas, heading.node_id);
            }
            DrawCommand::Outline(OutlineMark::End) => {
                pdf::set_node_id(canvas, pdf::node_id::NOTHING);
            }
            DrawCommand::Text {
                position,
                text,
                font_family,
                char_spacing,
                font_size,
                bold,
                italic,
                color,
                text_scale,
                shaped,
            } => {
                let (slot, base_font) =
                    font_cache.get_indexed(registry, font_family, *font_size, *bold, *italic);
                // §17.3.2.45: a non-1.0 scale is applied via Skia's scale_x —
                // that scales glyph advances and horizontal glyph extent
                // without touching the cached, shared Font. Cloning and
                // mutating a fresh Font keeps the cache invariant intact. A
                // scaled font is a one-off clone whose glyphs differ, so it
                // bypasses the (slot-keyed) blob cache.
                let scaled_font;
                let (font, blob_slot): (&skia_safe::Font, Option<usize>) =
                    if (*text_scale - 1.0).abs() > f32::EPSILON {
                        let mut f = base_font.clone();
                        f.set_scale_x(*text_scale);
                        scaled_font = f;
                        (&scaled_font, None)
                    } else {
                        (base_font, Some(slot))
                    };
                log::trace!(
                    "[paint] '{}' → font='{}' size={:.1}pt bold={:?} italic={:?} scale={:.2}",
                    &text[..text.len().min(30)],
                    font.typeface().family_name(),
                    font_size.raw(),
                    bold,
                    italic,
                    text_scale,
                );
                text_paint.set_color4f(to_color4f(*color), None);

                if let Some(direction) = shaped {
                    // Issue #131: a joining script. `draw_str` and
                    // `TextBlob::from_str` map codepoints through the cmap
                    // alone, which for Arabic paints every letter in its
                    // isolated form — a different word, not a plainer one.
                    //
                    // Shaping here rather than carrying glyph ids on the
                    // command is deliberate: `render::subset` rewrites a
                    // typeface's bytes *between* layout and paint, and a glyph
                    // id recorded against the pre-subset font need not name the
                    // same glyph in the subsetted one. Shaping against whatever
                    // the registry holds now is correct by construction. The
                    // advance agrees with the width layout measured because
                    // both come from this same shaper over the same typeface,
                    // and `fontcull` keeps the GSUB closure of every kept
                    // codepoint (see `subset::collect`), so the substitutions
                    // survive subsetting.
                    //
                    // §17.3.2.35 `w:spacing` is not applied to a shaped run —
                    // `layout::fragment::shape` says why, and does not add it
                    // to the measured width either.
                    if !draw_shaped(
                        canvas,
                        shaper.as_ref(),
                        font,
                        text,
                        *direction,
                        *position,
                        &text_paint,
                    ) {
                        // No shaper in this build, or nothing came back. Fall
                        // through to the cmap path: wrong in the way this
                        // engine was wrong before #131, rather than blank.
                        canvas.draw_str(text, to_point(*position), font, &text_paint);
                    }
                } else if char_spacing.abs() > Pt::ZERO {
                    // §17.3.2.35 w:spacing / §17.3.1.13 distribute — draw one
                    // spacing unit at a time with explicit spacing between
                    // them, matching the measured fragment width. The unit is
                    // a grapheme cluster, and comes from the same module the
                    // measurer counts with, so the width laid out and the width
                    // painted are the same number — see
                    // [`crate::render::spacing`].
                    let unit_count = spacing::unit_count(text);
                    let glyphs = font.text_to_glyphs_vec(&**text);
                    // Batch path: use text_to_glyphs + get_widths when glyph
                    // count matches unit count (common Latin/CJK text).
                    // A multi-codepoint cluster maps to more glyphs than units,
                    // so it falls to per-unit measure_str below and is drawn
                    // whole — the mark composes over its base because its own
                    // advance is zero.
                    let batch_widths = if glyphs.len() == unit_count {
                        let mut widths = vec![0f32; glyphs.len()];
                        font.get_widths(&glyphs, &mut widths);
                        Some(widths)
                    } else {
                        None
                    };

                    let mut cursor = *position;
                    for (i, unit) in spacing::units(text).enumerate() {
                        // Per-glyph widths from `get_widths` already include
                        // scale_x — they're advances of the scaled font.
                        // measure_str on the scaled font likewise returns the
                        // scaled advance, so no further multiplication is
                        // needed. char_spacing stays unscaled (§17.3.2.45).
                        let w = if let Some(ref widths) = batch_widths {
                            widths[i]
                        } else {
                            font.measure_str(unit, None).0
                        };
                        canvas.draw_str(unit, to_point(cursor), font, &text_paint);
                        cursor.x += Pt::new(w) + *char_spacing;
                    }
                } else if let Some(slot) = blob_slot {
                    // Common path: reuse a cached, position-independent glyph
                    // run instead of remapping text→glyphs on every draw_str.
                    // Visually identical — a TextBlob built from the font uses
                    // the same default cmap+advance positioning draw_str does.
                    let inner = blob_cache.entry(slot).or_default();
                    if let Some(blob) = inner.get(&**text) {
                        canvas.draw_text_blob(blob, to_point(*position), &text_paint);
                    } else if let Some(blob) = TextBlob::from_str(&**text, font) {
                        canvas.draw_text_blob(&blob, to_point(*position), &text_paint);
                        inner.insert(Box::from(&**text), blob);
                    }
                } else {
                    canvas.draw_str(text, to_point(*position), font, &text_paint);
                }
            }
            DrawCommand::Underline { line, color, width }
            | DrawCommand::Line { line, color, width } => {
                stroke_paint.set_stroke_width(f32::from(*width));
                stroke_paint.set_color4f(to_color4f(*color), None);

                let (start, end) = to_line(*line);
                canvas.draw_line(start, end, &stroke_paint);
            }
            DrawCommand::Image {
                rect,
                image_data,
                src_rect,
            } => {
                // §20.1.10.48: resolve the srcRect against the image bounds. A
                // region reaching outside the image (negative insets) pads the
                // frame — it becomes an in-image crop plus a shrunken `dst`,
                // and the padding band of `dst` is left undrawn. No srcRect →
                // whole image to the whole dst; a region wholly outside the
                // image → nothing to draw.
                let resolved = match src_rect {
                    None => Some((*rect, None)),
                    Some(sr) => resolve_src_padding(*rect, sr).map(|(d, c)| (d, Some(c))),
                };
                if let Some((draw_rect, crop)) = resolved {
                    let dst = to_rect(draw_rect);
                    // Key on (media identity, downsample target for the display
                    // size, crop). The crop is baked into the cached bitmap
                    // below, so two placements that differ only in size or in
                    // `srcRect` each cache their own bitmap.
                    let (target_w, target_h) = downsample_target(draw_rect, image_dpi);
                    let key = (
                        std::sync::Arc::as_ptr(&image_data.data),
                        target_w,
                        target_h,
                        crop.as_ref().map(quantize_crop),
                    );
                    if let Some(image) = image_cache.get(&key) {
                        // The cached bitmap already has any crop baked in → the
                        // whole bitmap maps to `dst`.
                        canvas.draw_image_rect(image, None, dst, &default_paint);
                    } else if let Some(decoded) = decode_image(image_data) {
                        match &crop {
                            // Bake the visible sub-region into a right-sized
                            // bitmap so the PDF never carries the cropped-away
                            // pixels at full resolution (finding 3).
                            Some(crop) => {
                                match prepare_cropped(&decoded, crop, draw_rect, image_dpi) {
                                    Some(image) => {
                                        canvas.draw_image_rect(&image, None, dst, &default_paint);
                                        image_cache.insert(key, image);
                                    }
                                    // Near-OOM: the raster surface could not be
                                    // allocated. Fall back to a draw-time crop of
                                    // the full image — correct pixels, just
                                    // unoptimized — and leave it uncached (the key
                                    // implies a baked crop).
                                    None => {
                                        let src = src_pixel_rect(&decoded, crop);
                                        canvas.draw_image_rect(
                                            &decoded,
                                            Some((&src, SrcRectConstraint::Strict)),
                                            dst,
                                            &default_paint,
                                        );
                                    }
                                }
                            }
                            None => {
                                let image = downsample_if_oversize(decoded, draw_rect, image_dpi);
                                canvas.draw_image_rect(&image, None, dst, &default_paint);
                                image_cache.insert(key, image);
                            }
                        }
                    } else {
                        let magic = &image_data.data[..image_data.data.len().min(4)];
                        log::warn!(
                            "[paint] unsupported image format {:?} — could not decode {} bytes \
                             (magic: {:02x?}); image will be blank",
                            image_data.format,
                            image_data.data.len(),
                            magic,
                        );
                    }
                }
            }
            DrawCommand::EmojiCluster {
                rect,
                text,
                typeface,
                size,
                presentation,
                structure,
            } => {
                use crate::render::emoji::cluster::EmojiCluster;
                use skia_safe::{CubicResampler, SamplingOptions};
                let cluster = EmojiCluster {
                    text: text.as_str(),
                    presentation: *presentation,
                    structure: *structure,
                };
                // Pass `rect.size` so the rasterizer allocates an image
                // whose aspect matches the rect → uniform scaling at
                // `draw_image_rect`, no anisotropic distortion.
                // `None` means the offscreen surface could not be allocated;
                // the rasterizer has already logged why. Leave the cluster
                // blank, as an undecodable image is left blank above.
                if let Some(img) = emoji_rasterizer.rasterize(&cluster, typeface, *size, rect.size)
                {
                    // Mitchell cubic resampling — same filter we use for
                    // photographic image downsampling. Without explicit
                    // sampling, Skia defaults to nearest/bilinear which makes
                    // the emoji look blurry/pixelated at typical PDF zoom.
                    let sampling = SamplingOptions::from(CubicResampler::mitchell());
                    canvas.draw_image_rect_with_sampling_options(
                        &img.image,
                        None,
                        to_rect(*rect),
                        sampling,
                        &default_paint,
                    );
                }
            }
            DrawCommand::Rect { rect, color } => {
                rect_paint.set_color4f(to_color4f(*color), None);
                canvas.draw_rect(to_rect(*rect), &rect_paint);
            }
            DrawCommand::LinkAnnotation { rect, url } => {
                let mut url_bytes = url.as_bytes().to_vec();
                url_bytes.push(0);
                let url_data = Data::new_copy(&url_bytes);
                canvas.annotate_rect_with_url(to_rect(*rect), &url_data);
            }
            DrawCommand::InternalLink { rect, destination } => {
                let mut name_bytes = destination.as_bytes().to_vec();
                name_bytes.push(0);
                let name_data = Data::new_copy(&name_bytes);
                canvas.annotate_link_to_destination(to_rect(*rect), &name_data);
            }
            DrawCommand::NamedDestination { position, name } => {
                let mut name_bytes = name.as_bytes().to_vec();
                name_bytes.push(0);
                let name_data = Data::new_copy(&name_bytes);
                canvas.annotate_named_destination(to_point(*position), &name_data);
            }
            DrawCommand::Path {
                origin,
                rotation,
                flip_h,
                flip_v,
                extent,
                paths,
                fill,
                stroke,
                effects,
            } => {
                canvas.save();
                // Translate to the shape's origin.
                canvas.translate((origin.x.raw(), origin.y.raw()));
                // Apply flip / rotation around the shape's center.
                let cx = extent.width.raw() / 2.0;
                let cy = extent.height.raw() / 2.0;
                let rot_deg = rotation.raw() as f32 / 60_000.0;
                if *flip_h || *flip_v || rot_deg != 0.0 {
                    canvas.translate((cx, cy));
                    if rot_deg != 0.0 {
                        canvas.rotate(rot_deg, None);
                    }
                    let sx = if *flip_h { -1.0 } else { 1.0 };
                    let sy = if *flip_v { -1.0 } else { 1.0 };
                    if sx != 1.0 || sy != 1.0 {
                        canvas.scale((sx, sy));
                    }
                    canvas.translate((-cx, -cy));
                }
                let skia_path = build_skia_path(paths, warned);
                let strokable = build_skia_path_stroked_only(paths);
                // §20.1.8 effects render beneath the shape itself, in the
                // order they appear in the effect list.
                for effect in effects {
                    paint_effect(
                        canvas,
                        effect,
                        fill,
                        stroke.as_ref(),
                        &skia_path,
                        &strokable,
                    );
                }
                if let Some(paint) = fill_to_paint(fill, warned) {
                    canvas.draw_path(&skia_path, &paint);
                }
                if let Some(stroke) = stroke.as_ref() {
                    let paint = stroke_to_paint(stroke);
                    // Only stroke subpaths whose .stroked flag is set.
                    canvas.draw_path(&strokable, &paint);
                }
                canvas.restore();
            }
        }
    }
}

// ── Shape path helpers ──────────────────────────────────────────────────────

/// Build a Skia path from the subpaths OOXML actually fills, regardless of
/// stroke flag. Used both for fill painting and for the shadow silhouette, so
/// the two agree on what has an interior.
///
/// §20.1.10.45 `ST_PathFillMode`: a subpath with `fill="none"` contributes no
/// interior — it is outline-only, and filling it anyway paints the shape's
/// fill colour over whatever sits behind it. The `lighten`/`lightenLess`/
/// `darken`/`darkenLess` variants ask for a shaded version of the shape fill;
/// Tier-0 paints them as `norm` and warns once per render.
fn build_skia_path(paths: &[SubPath], warned: &mut FxHashSet<&'static str>) -> Path {
    use crate::model::PathFillMode;
    let mut builder = PathBuilder::new();
    builder.set_fill_type(PathFillType::Winding);
    for sub in paths {
        match sub.fill_mode {
            PathFillMode::None => continue,
            PathFillMode::Norm => {}
            shaded => {
                let name = match shaded {
                    PathFillMode::Lighten => "fillMode:lighten",
                    PathFillMode::LightenLess => "fillMode:lightenLess",
                    PathFillMode::Darken => "fillMode:darken",
                    _ => "fillMode:darkenLess",
                };
                if warned.insert(name) {
                    log::warn!(
                        "paint: §20.1.10.45 path fill mode {name:?} not shaded (Tier 0) — \
                         painting it as 'norm'"
                    );
                }
            }
        }
        emit_subpath(&mut builder, sub);
    }
    builder.snapshot()
}

/// Build a Skia path limited to subpaths with `.stroked == true`.
fn build_skia_path_stroked_only(paths: &[SubPath]) -> Path {
    let mut builder = PathBuilder::new();
    for sub in paths {
        if sub.stroked {
            emit_subpath(&mut builder, sub);
        }
    }
    builder.snapshot()
}

fn emit_subpath(builder: &mut PathBuilder, sub: &SubPath) {
    // Track the last point manually: `PathBuilder` has no `last_pt()` query,
    // and `arc_to` needs the current pen position to derive the bounding oval.
    let mut last_pt: (f32, f32) = (0.0, 0.0);
    for verb in &sub.verbs {
        match verb {
            PathVerb::MoveTo(p) => {
                let pt = (p.x.raw(), p.y.raw());
                builder.move_to(pt);
                last_pt = pt;
            }
            PathVerb::LineTo(p) => {
                let pt = (p.x.raw(), p.y.raw());
                builder.line_to(pt);
                last_pt = pt;
            }
            PathVerb::QuadTo(c, p) => {
                let pt = (p.x.raw(), p.y.raw());
                builder.quad_to((c.x.raw(), c.y.raw()), pt);
                last_pt = pt;
            }
            PathVerb::CubicTo(c1, c2, p) => {
                let pt = (p.x.raw(), p.y.raw());
                builder.cubic_to((c1.x.raw(), c1.y.raw()), (c2.x.raw(), c2.y.raw()), pt);
                last_pt = pt;
            }
            PathVerb::ArcTo {
                radii,
                start_angle,
                swing_angle,
            } => {
                // §20.1.9.3: the arc starts at the current pen position,
                // which lies ON the ellipse at `stAng` — the spec locks the
                // start angle "to the last known pen position", so the
                // center is the pen minus the start-point offset, not the
                // pen itself. `stAng`/`swAng` are ray angles measured at the
                // center (0° = 3 o'clock, clockwise positive); Skia's
                // `arc_to` angles are parametric (the point at angle t is
                // center + (wr·cos t, hr·sin t)), so each ray angle is
                // skewed onto the ellipse via t = atan2(wr·sinθ, hr·cosθ) —
                // the same conversion LibreOffice's ARCANGLETO and Apache
                // POI's ArcToCommand apply. On a circle the two conventions
                // coincide.
                let (wr, hr) = (radii.width.raw() as f64, radii.height.raw() as f64);
                let st = start_angle.raw() as f64 / 60_000.0;
                // Word clamps the swing to one full revolution
                // (LibreOffice matches it — tdf#122323).
                let sw = (swing_angle.raw() as f64 / 60_000.0).clamp(-360.0, 360.0);
                let t1 = ellipse_param_deg(st, wr, hr);
                let t2 = ellipse_param_deg(st + sw, wr, hr);
                let (cx, cy) = (
                    last_pt.0 as f64 - wr * t1.to_radians().cos(),
                    last_pt.1 as f64 - hr * t1.to_radians().sin(),
                );
                let oval = skia_safe::Rect::from_xywh(
                    (cx - wr) as f32,
                    (cy - hr) as f32,
                    (wr * 2.0) as f32,
                    (hr * 2.0) as f32,
                );
                // Skia treats a sweep modulo 360°, which would erase a
                // full-circle swing — emit it as two half turns.
                if sw.abs() >= 360.0 {
                    let tm = ellipse_param_deg(st + sw / 2.0, wr, hr);
                    builder.arc_to(oval, t1 as f32, (tm - t1) as f32, false);
                    builder.arc_to(oval, tm as f32, (t2 - tm) as f32, false);
                } else {
                    builder.arc_to(oval, t1 as f32, (t2 - t1) as f32, false);
                }
                let end = t2.to_radians();
                last_pt = ((cx + wr * end.cos()) as f32, (cy + hr * end.sin()) as f32);
            }
            PathVerb::Close => {
                builder.close();
            }
        }
    }
}

/// Skew a §20.1.9.3 ray angle (degrees, 0° = 3 o'clock, clockwise positive)
/// onto the ellipse's parametric angle (degrees), staying in the same
/// revolution as the input so sweeps keep their direction and multiplicity.
///
/// A point at ray angle θ from the center of an `(wr, hr)` ellipse sits at
/// parametric angle `t = atan2(wr·sinθ, hr·cosθ)`: the ray hits the ellipse
/// where `(wr·cos t, hr·sin t) ∝ (cosθ, sinθ)`. `atan2` collapses
/// revolutions, so the skew is re-applied as a delta around θ — the true
/// skew never reaches ±90°, which keeps the wrap-around unambiguous.
fn ellipse_param_deg(theta_deg: f64, wr: f64, hr: f64) -> f64 {
    let (s, c) = theta_deg.to_radians().sin_cos();
    let (y, x) = (wr * s, hr * c);
    if x == 0.0 && y == 0.0 {
        // Degenerate radii — keep the ray angle.
        return theta_deg;
    }
    let mut delta = (y.atan2(x).to_degrees() - theta_deg) % 360.0;
    if delta > 180.0 {
        delta -= 360.0;
    } else if delta < -180.0 {
        delta += 360.0;
    }
    theta_deg + delta
}

/// The paint for a shape's fill, or `None` when nothing should be filled.
///
/// Unsupported fill kinds warn **once per render** via `warned` — a
/// gradient-filled shape in a header would otherwise report on every page.
fn fill_to_paint(fill: &ResolvedFill, warned: &mut FxHashSet<&'static str>) -> Option<Paint> {
    let mut warn_once = |key: &'static str, tier: &str| {
        if warned.insert(key) {
            log::warn!("paint: {key} fill not yet rendered ({tier})");
        }
        None
    };
    match fill {
        ResolvedFill::None => None,
        ResolvedFill::Solid(color) => {
            let mut paint = Paint::default();
            paint.set_anti_alias(true);
            paint.set_style(skia_safe::PaintStyle::Fill);
            paint.set_color4f(rgba_to_color4f(*color), None);
            Some(paint)
        }
        ResolvedFill::Gradient(_) => warn_once("gradient", "Tier 2"),
        ResolvedFill::Blip(_) => warn_once("blip", "Tier 2"),
        ResolvedFill::Pattern(_) => warn_once("pattern", "Tier 3"),
    }
}

fn stroke_to_paint(stroke: &ResolvedStroke) -> Paint {
    let mut paint = Paint::default();
    paint.set_anti_alias(true);
    paint.set_style(skia_safe::PaintStyle::Stroke);
    paint.set_stroke_width(stroke.width.raw());
    paint.set_color4f(rgba_to_color4f(stroke.color), None);
    paint.set_stroke_cap(match stroke.cap {
        ResolvedLineCap::Butt => skia_safe::PaintCap::Butt,
        ResolvedLineCap::Round => skia_safe::PaintCap::Round,
        ResolvedLineCap::Square => skia_safe::PaintCap::Square,
    });
    paint.set_stroke_join(match stroke.join {
        ResolvedLineJoin::Round => skia_safe::PaintJoin::Round,
        ResolvedLineJoin::Bevel => skia_safe::PaintJoin::Bevel,
        ResolvedLineJoin::Miter => skia_safe::PaintJoin::Miter,
    });
    if let ResolvedDashPattern::Dashes(dashes) = &stroke.dash {
        if !dashes.is_empty() {
            let floats: Vec<f32> = dashes.iter().map(|p| p.raw()).collect();
            if let Some(effect) = PathEffect::dash(&floats, 0.0) {
                paint.set_path_effect(effect);
            }
        }
    }
    paint
}

fn rgba_to_color4f(c: Rgba) -> Color4f {
    Color4f::new(c.r, c.g, c.b, c.a)
}

/// Paint a shape effect beneath the shape itself. The effect color is the
/// one already resolved from the effect's `<a:srgbClr>` / `<a:schemeClr>`
/// plus color transforms; the fill/stroke silhouette drives the shadow's
/// shape.
fn paint_effect(
    canvas: &skia_safe::Canvas,
    effect: &ResolvedEffect,
    fill: &ResolvedFill,
    stroke: Option<&ResolvedStroke>,
    shape_path: &Path,
    strokable_path: &Path,
) {
    match effect {
        ResolvedEffect::OuterShadow {
            blur_radius,
            offset,
            color,
        } => {
            // §20.1.8.45: a Gaussian blur. Skia's mask-filter sigma ≈
            // radius / 2 — the conventional approximation used by other
            // renderers (LibreOffice, Chromium's CSS filter).
            let sigma = (blur_radius.raw() * 0.5).max(0.0);
            let mask = if sigma > 0.0 {
                MaskFilter::blur(BlurStyle::Normal, sigma, None)
            } else {
                None
            };
            canvas.save();
            canvas.translate((offset.x.raw(), offset.y.raw()));
            // Fill silhouette (when the shape has a fill).
            if !matches!(fill, ResolvedFill::None) {
                let mut paint = Paint::default();
                paint.set_anti_alias(true);
                paint.set_style(skia_safe::PaintStyle::Fill);
                paint.set_color4f(rgba_to_color4f(*color), None);
                if let Some(m) = mask.clone() {
                    paint.set_mask_filter(m);
                }
                canvas.draw_path(shape_path, &paint);
            }
            // Stroke silhouette — cast the shadow from the stroke's own
            // outline so line-preset shapes cast a visible shadow.
            if let Some(s) = stroke {
                let mut paint = Paint::default();
                paint.set_anti_alias(true);
                paint.set_style(skia_safe::PaintStyle::Stroke);
                paint.set_stroke_width(s.width.raw());
                paint.set_stroke_cap(match s.cap {
                    ResolvedLineCap::Butt => skia_safe::PaintCap::Butt,
                    ResolvedLineCap::Round => skia_safe::PaintCap::Round,
                    ResolvedLineCap::Square => skia_safe::PaintCap::Square,
                });
                paint.set_stroke_join(match s.join {
                    ResolvedLineJoin::Round => skia_safe::PaintJoin::Round,
                    ResolvedLineJoin::Bevel => skia_safe::PaintJoin::Bevel,
                    ResolvedLineJoin::Miter => skia_safe::PaintJoin::Miter,
                });
                paint.set_color4f(rgba_to_color4f(*color), None);
                if let Some(m) = mask.clone() {
                    paint.set_mask_filter(m);
                }
                canvas.draw_path(strokable_path, &paint);
            }
            canvas.restore();
        }
    }
}

/// Decode a `MediaEntry` to a Skia image, dispatching on format.
///
/// Returns `None` if the format is unsupported or the data is malformed.
fn decode_image(entry: &MediaEntry) -> Option<skia_safe::Image> {
    use crate::model::ImageFormat;
    match entry.format {
        ImageFormat::Emf => crate::render::emf::decode_emf_bitmap(&entry.data),
        // All other formats are handled by Skia's built-in decoder.
        _ => skia_safe::Image::from_encoded(Data::new_copy(&entry.data)),
    }
}

/// Downsample an image if its native pixel dimensions significantly exceed
/// the display dimensions at `image_dpi`. Uses Mitchell-Netravali cubic
/// filtering for high-quality results.
fn downsample_if_oversize(
    image: skia_safe::Image,
    rect: crate::render::geometry::PtRect,
    image_dpi: f32,
) -> skia_safe::Image {
    use skia_safe::CubicResampler;
    use skia_safe::{AlphaType, ColorType, ImageInfo, SamplingOptions};

    let (target_w, target_h) = downsample_target(rect, image_dpi);
    if image.width() > target_w && image.height() > target_h && target_w > 0 && target_h > 0 {
        log::debug!(
            "[paint] downsampling image {}×{} → {}×{} (display {:.0}×{:.0}pt @ {:.0} DPI)",
            image.width(),
            image.height(),
            target_w,
            target_h,
            rect.size.width.raw(),
            rect.size.height.raw(),
            image_dpi,
        );
        // Draw scaled image onto an opaque surface so Skia applies JPEG
        // encoding (encoding_quality) instead of lossless FlateDecode.
        let info = ImageInfo::new(
            (target_w, target_h),
            ColorType::RGBA8888,
            AlphaType::Opaque,
            None,
        );
        let sampling = SamplingOptions::from(CubicResampler::mitchell());
        if let Some(mut surface) = skia_safe::surfaces::raster(&info, None, None) {
            let dst = skia_safe::Rect::from_iwh(target_w, target_h);
            surface.canvas().draw_image_rect_with_sampling_options(
                &image,
                None,
                dst,
                sampling,
                &Paint::default(),
            );
            surface.image_snapshot()
        } else {
            image
        }
    } else {
        image
    }
}

/// The visible sub-rectangle in the image's own pixel space, from the
/// fractional `srcRect` crop (§20.1.10.48).
fn src_pixel_rect(
    image: &skia_safe::Image,
    crop: &crate::render::geometry::PtRect,
) -> skia_safe::Rect {
    let (iw, ih) = (image.width() as f32, image.height() as f32);
    skia_safe::Rect::from_xywh(
        crop.origin.x.raw() * iw,
        crop.origin.y.raw() * ih,
        crop.size.width.raw() * iw,
        crop.size.height.raw() * ih,
    )
}

/// Crop `image` to the fractional visible region (`src_rect`, §20.1.10.48) and
/// scale that region to the display target resolution in one Mitchell pass.
///
/// Baking the crop into the bitmap means only the visible pixels are embedded:
/// a heavy crop no longer carries its cropped-away pixels into the PDF at full
/// resolution (finding 3). The result maps whole → destination at draw time.
/// The surface is capped at the visible region's own pixel count, so a
/// magnifying crop is not upsampled, and at the display target, so a
/// high-resolution crop is downsampled.
///
/// Returns `None` only if the raster surface can't be allocated (near-OOM); the
/// caller then falls back to a draw-time crop of the full image.
fn prepare_cropped(
    image: &skia_safe::Image,
    src_rect: &crate::render::geometry::PtRect,
    display: crate::render::geometry::PtRect,
    image_dpi: f32,
) -> Option<skia_safe::Image> {
    use skia_safe::{AlphaType, ColorType, CubicResampler, ImageInfo, SamplingOptions};

    let src = src_pixel_rect(image, src_rect);
    // Never allocate more pixels than the visible region provides (no upsample)
    // nor more than the display needs at target DPI (no bloat).
    let (disp_w, disp_h) = downsample_target(display, image_dpi);
    let target_w = disp_w.min(src.width().round() as i32).max(1);
    let target_h = disp_h.min(src.height().round() as i32).max(1);

    // Preserve the source's opacity: an opaque image encodes as JPEG
    // (encoding_quality); a transparent one keeps its alpha (lossless).
    let alpha = if image.alpha_type() == AlphaType::Opaque {
        AlphaType::Opaque
    } else {
        AlphaType::Premul
    };
    let info = ImageInfo::new((target_w, target_h), ColorType::RGBA8888, alpha, None);
    let sampling = SamplingOptions::from(CubicResampler::mitchell());
    let mut surface = skia_safe::surfaces::raster(&info, None, None)?;
    let dst = skia_safe::Rect::from_iwh(target_w, target_h);
    surface.canvas().draw_image_rect_with_sampling_options(
        image,
        Some((&src, SrcRectConstraint::Strict)),
        dst,
        sampling,
        &Paint::default(),
    );
    Some(surface.image_snapshot())
}

/// Resolve an `a:srcRect` region against the image bounds for painting
/// (§20.1.10.48). The fractional region `[origin, origin + size]` maps linearly
/// onto `dst`. When it reaches outside the image's `[0, 1]` extent — negative
/// insets, i.e. letterbox/pillarbox padding — the outside area is *empty*, not
/// clipped: clamp the region to the image and shrink `dst` to the
/// sub-rectangle the clamped region maps onto, so the padding band of `dst` is
/// simply left undrawn (transparent) instead of `Strict` sampling stretching
/// the image over it.
///
/// Returns the in-bounds `(dst, crop)` to draw, or `None` when the region lies
/// wholly outside the image (nothing visible). An already in-bounds region
/// comes back with `dst` and `crop` unchanged.
fn resolve_src_padding(
    dst: crate::render::geometry::PtRect,
    src: &crate::render::geometry::PtRect,
) -> Option<(
    crate::render::geometry::PtRect,
    crate::render::geometry::PtRect,
)> {
    use crate::render::geometry::PtRect;

    let (sx0, sy0) = (src.origin.x.raw(), src.origin.y.raw());
    let (sw, sh) = (src.size.width.raw(), src.size.height.raw());
    if sw <= 0.0 || sh <= 0.0 {
        return None;
    }
    let (sx1, sy1) = (sx0 + sw, sy0 + sh);
    // Clamp the source region to the image extent [0, 1].
    let (cx0, cy0) = (sx0.max(0.0), sy0.max(0.0));
    let (cx1, cy1) = (sx1.min(1.0), sy1.min(1.0));
    if cx1 <= cx0 || cy1 <= cy0 {
        return None; // wholly outside the image
    }
    // Fraction of `dst` the clamped region covers (0 at the region's own edge).
    let (tx0, tx1) = ((cx0 - sx0) / sw, (cx1 - sx0) / sw);
    let (ty0, ty1) = ((cy0 - sy0) / sh, (cy1 - sy0) / sh);
    let (dx, dy) = (dst.origin.x.raw(), dst.origin.y.raw());
    let (dw, dh) = (dst.size.width.raw(), dst.size.height.raw());
    let inset_dst = PtRect::from_xywh(
        Pt::new(dx + tx0 * dw),
        Pt::new(dy + ty0 * dh),
        Pt::new((tx1 - tx0) * dw),
        Pt::new((ty1 - ty0) * dh),
    );
    let crop = PtRect::from_xywh(
        Pt::new(cx0),
        Pt::new(cy0),
        Pt::new(cx1 - cx0),
        Pt::new(cy1 - cy0),
    );
    Some((inset_dst, crop))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::geometry::{PtOffset, PtSize};
    use crate::render::resolve::color::RgbColor;
    use skia_safe::FontMgr;
    use std::rc::Rc;

    fn test_font_mgr() -> FontMgr {
        FontMgr::new()
    }

    fn test_registry() -> FontRegistry {
        FontRegistry::new(test_font_mgr())
    }

    // ── §20.1.9.3 arcTo ─────────────────────────────────────────────
    //
    // Expected values are computed from the spec's geometry, not from the
    // implementation: the pen lies ON the ellipse at `stAng` (the spec locks
    // the start angle "to the last known pen position"), so the center is
    // `pen − (wr·cos t₁, hr·sin t₁)` with t₁ the parametric angle of the
    // ray angle `stAng`; the end point sits at ray angle `stAng + swAng`.
    // Apache POI's `ArcToCommandIf` and LibreOffice's `ARCANGLETO` implement
    // the same construction.

    fn arc_path(move_to: (f32, f32), radii: (f32, f32), st: i64, sw: i64) -> Path {
        let sub = SubPath {
            verbs: vec![
                PathVerb::MoveTo(PtOffset::new(Pt::new(move_to.0), Pt::new(move_to.1))),
                PathVerb::ArcTo {
                    radii: PtSize::new(Pt::new(radii.0), Pt::new(radii.1)),
                    start_angle: crate::model::dimension::Dimension::new(st),
                    swing_angle: crate::model::dimension::Dimension::new(sw),
                },
            ],
            fill_mode: crate::model::PathFillMode::Norm,
            stroked: true,
        };
        let mut builder = PathBuilder::new();
        emit_subpath(&mut builder, &sub);
        builder.snapshot()
    }

    fn assert_close(actual: skia_safe::Point, expected: (f32, f32), what: &str) {
        let (dx, dy) = (actual.x - expected.0, actual.y - expected.1);
        assert!(
            dx.abs() < 0.02 && dy.abs() < 0.02,
            "{what}: got ({}, {}), expected ({}, {})",
            actual.x,
            actual.y,
            expected.0,
            expected.1
        );
    }

    #[test]
    fn arc_starts_at_pen_and_ends_on_ellipse() {
        // Pen (10,10) at stAng 0° on a circle of radius 5 → center (5,10).
        // A +90° swing ends at (5,15). Nothing may reach right of the pen:
        // the buggy center-at-pen reading put the arc start at (15,10) and
        // lined over to it.
        let path = arc_path((10.0, 10.0), (5.0, 5.0), 0, 5_400_000);
        assert_close(path.last_pt().unwrap(), (5.0, 15.0), "arc end");
        let b = path.compute_tight_bounds();
        assert!(
            (b.left - 5.0).abs() < 0.02
                && (b.top - 10.0).abs() < 0.02
                && (b.right - 10.0).abs() < 0.02
                && (b.bottom - 15.0).abs() < 0.02,
            "quarter-arc bounds: {b:?}, expected (5,10)-(10,15)"
        );
    }

    #[test]
    fn arc_angles_are_ray_angles_not_parametric() {
        // wr=20, hr=10, stAng=45°: the parametric angle of the pen is
        // t₁ = atan2(20·sin45°, 10·cos45°) = 63.435°, so the pen offset is
        // (20·cos t₁, 10·sin t₁) = (8.944, 8.944) and center = pen − that =
        // (111.056, 41.056). A +90° swing ends at ray angle 135°, i.e.
        // t₂ = 116.565° → end = center + (−8.944, 8.944) = (102.111, 50.0).
        // Reading the angles as parametric instead would end at
        // center + (20·cos135°, 10·sin135°) = (96.91, 48.13).
        let path = arc_path((120.0, 50.0), (20.0, 10.0), 2_700_000, 5_400_000);
        assert_close(path.last_pt().unwrap(), (102.1115, 50.0), "ellipse arc end");
    }

    #[test]
    fn arc_full_swing_draws_whole_ellipse() {
        // swAng = 21600000 (360°) draws the complete circle around
        // center (5,10) and returns to the pen. Skia treats a single
        // arc_to sweep modulo 360°, so an unsplit full swing vanishes.
        let path = arc_path((10.0, 10.0), (5.0, 5.0), 0, 21_600_000);
        assert_close(path.last_pt().unwrap(), (10.0, 10.0), "full-circle end");
        let b = path.compute_tight_bounds();
        assert!(
            (b.left - 0.0).abs() < 0.02
                && (b.top - 5.0).abs() < 0.02
                && (b.right - 10.0).abs() < 0.02
                && (b.bottom - 15.0).abs() < 0.02,
            "full-circle bounds: {b:?}, expected (0,5)-(10,15)"
        );
    }

    #[test]
    fn arc_negative_swing_sweeps_counterclockwise() {
        // Same circle as above; −90° sweeps the other way and ends at (5,5).
        let path = arc_path((10.0, 10.0), (5.0, 5.0), 0, -5_400_000);
        assert_close(path.last_pt().unwrap(), (5.0, 5.0), "ccw arc end");
        let b = path.compute_tight_bounds();
        assert!(
            (b.left - 5.0).abs() < 0.02
                && (b.top - 5.0).abs() < 0.02
                && (b.right - 10.0).abs() < 0.02
                && (b.bottom - 10.0).abs() < 0.02,
            "ccw quarter-arc bounds: {b:?}, expected (5,5)-(10,10)"
        );
    }

    // ── render_to_pdf integration ───────────────────────────────────

    #[test]
    fn render_text_command_produces_pdf() {
        let registry = test_registry();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                shaped: None,
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "Hello world".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::ZERO,
                font_size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                color: RgbColor::BLACK,
                text_scale: 1.0,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &registry, &RenderOptions::default())
            .expect("render_to_pdf must succeed");
        assert!(pdf_bytes.len() > 100, "PDF output must be non-trivial");
        assert_eq!(&pdf_bytes[..5], b"%PDF-", "output must be valid PDF");
    }

    #[test]
    fn render_text_with_char_spacing_produces_pdf() {
        let registry = test_registry();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                shaped: None,
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "Spaced".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::new(2.0),
                font_size: Pt::new(14.0),
                bold: Toggle::On,
                italic: Toggle::Absent,
                color: RgbColor::BLACK,
                text_scale: 1.0,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &registry, &RenderOptions::default())
            .expect("render_to_pdf must succeed");
        assert!(pdf_bytes.len() > 100);
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }

    /// §17.3.2.35 with a multi-scalar cluster. The spaced path draws one unit
    /// at a time, so a run whose cluster spans two scalars must still paint —
    /// and must reach the per-unit `measure_str` branch, not the batched
    /// per-glyph one (see the companion test below).
    #[test]
    fn render_combining_mark_with_char_spacing_produces_pdf() {
        let registry = test_registry();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                shaped: None,
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "e\u{301}cole\u{301}".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::new(3.0),
                font_size: Pt::new(14.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                color: RgbColor::BLACK,
                text_scale: 1.0,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &registry, &RenderOptions::default())
            .expect("render_to_pdf must succeed");
        assert!(pdf_bytes.len() > 100);
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }

    /// The batch fast path is gated on `glyphs.len() == spacing::unit_count`,
    /// and that equality is exactly what must fail for a cluster that spans
    /// more than one scalar: `e` + `U+0301` is two glyphs but one unit, so the
    /// painter falls back to measuring and drawing the cluster whole. Were the
    /// gate still a scalar count the two would agree, the batch path would be
    /// taken, and the accent would be drawn a spacing step away from its base.
    #[test]
    fn multi_scalar_cluster_falls_out_of_the_batched_width_path() {
        let registry = test_registry();
        let mut cache = fonts::FontCache::new();
        let (_, font) = cache.get_indexed(
            &registry,
            "Helvetica",
            Pt::new(14.0),
            Toggle::Absent,
            Toggle::Absent,
        );

        let accented = "e\u{301}";
        let glyphs = font.text_to_glyphs_vec(accented);
        if glyphs.is_empty() {
            return; // headless host with no usable typeface
        }
        assert_eq!(
            spacing::unit_count(accented),
            1,
            "one grapheme cluster is one spacing unit"
        );
        assert_ne!(
            glyphs.len(),
            spacing::unit_count(accented),
            "glyph count must not match unit count, or the cluster would be \
             split across the batched per-glyph width path"
        );

        // Plain ASCII still takes the batch path — the fallback is for clusters,
        // not a blanket slowdown.
        let plain = "abc";
        assert_eq!(
            font.text_to_glyphs_vec(plain).len(),
            spacing::unit_count(plain)
        );
    }

    #[test]
    fn render_empty_text_produces_pdf() {
        let registry = test_registry();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                shaped: None,
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: Rc::from(""),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::ZERO,
                font_size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                color: RgbColor::BLACK,
                text_scale: 1.0,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &registry, &RenderOptions::default())
            .expect("empty text must not panic");
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }

    // ── DrawCommand::Path ─────────────────────────────────────────────

    #[test]
    fn render_path_solid_filled_rect() {
        use crate::model::dimension::Dimension;
        use crate::model::PathFillMode;
        use crate::render::layout::draw_command::{
            ResolvedDashPattern, ResolvedFill, ResolvedLineCap, ResolvedLineJoin, ResolvedStroke,
        };
        use crate::render::resolve::drawing_color::Rgba;
        use crate::render::resolve::shape_geometry::{PathVerb, SubPath};

        let verbs = vec![
            PathVerb::MoveTo(PtOffset::new(Pt::ZERO, Pt::ZERO)),
            PathVerb::LineTo(PtOffset::new(Pt::new(100.0), Pt::ZERO)),
            PathVerb::LineTo(PtOffset::new(Pt::new(100.0), Pt::new(50.0))),
            PathVerb::LineTo(PtOffset::new(Pt::ZERO, Pt::new(50.0))),
            PathVerb::Close,
        ];
        let page = LayoutedPage {
            commands: vec![DrawCommand::Path {
                origin: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                rotation: Dimension::new(0),
                flip_h: false,
                flip_v: false,
                extent: PtSize::new(Pt::new(100.0), Pt::new(50.0)),
                paths: vec![SubPath {
                    verbs,
                    fill_mode: PathFillMode::Norm,
                    stroked: true,
                }],
                fill: ResolvedFill::Solid(Rgba {
                    r: 1.0,
                    g: 0.0,
                    b: 0.0,
                    a: 1.0,
                }),
                stroke: Some(ResolvedStroke {
                    width: Pt::new(1.0),
                    color: Rgba::BLACK,
                    dash: ResolvedDashPattern::Solid,
                    cap: ResolvedLineCap::Butt,
                    join: ResolvedLineJoin::Miter,
                }),
                effects: vec![],
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };
        let pdf = render_to_pdf(&[page], &test_registry(), &RenderOptions::default())
            .expect("render path");
        assert_eq!(&pdf[..5], b"%PDF-");
    }

    #[test]
    fn render_path_dashed_line() {
        use crate::model::dimension::Dimension;
        use crate::model::PathFillMode;
        use crate::render::layout::draw_command::{
            ResolvedDashPattern, ResolvedFill, ResolvedLineCap, ResolvedLineJoin, ResolvedStroke,
        };
        use crate::render::resolve::drawing_color::Rgba;
        use crate::render::resolve::shape_geometry::{PathVerb, SubPath};

        let page = LayoutedPage {
            commands: vec![DrawCommand::Path {
                origin: PtOffset::new(Pt::new(50.0), Pt::new(50.0)),
                rotation: Dimension::new(0),
                flip_h: false,
                flip_v: false,
                extent: PtSize::new(Pt::new(100.0), Pt::new(0.0)),
                paths: vec![SubPath {
                    verbs: vec![
                        PathVerb::MoveTo(PtOffset::new(Pt::ZERO, Pt::ZERO)),
                        PathVerb::LineTo(PtOffset::new(Pt::new(100.0), Pt::ZERO)),
                    ],
                    fill_mode: PathFillMode::None,
                    stroked: true,
                }],
                fill: ResolvedFill::None,
                stroke: Some(ResolvedStroke {
                    width: Pt::new(2.0),
                    color: Rgba {
                        r: 0.85,
                        g: 0.6,
                        b: 0.2,
                        a: 1.0,
                    },
                    dash: ResolvedDashPattern::Dashes(vec![Pt::new(6.0), Pt::new(3.0)]),
                    cap: ResolvedLineCap::Round,
                    join: ResolvedLineJoin::Round,
                }),
                effects: vec![],
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };
        let pdf = render_to_pdf(&[page], &test_registry(), &RenderOptions::default())
            .expect("render dashed line");
        assert_eq!(&pdf[..5], b"%PDF-");
    }

    #[test]
    fn render_unicode_text_produces_pdf() {
        let registry = test_registry();
        let page = LayoutedPage {
            commands: vec![DrawCommand::Text {
                shaped: None,
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: "Ärzte für Ökologie — 日本語".into(),
                font_family: Rc::from("Helvetica"),
                char_spacing: Pt::ZERO,
                font_size: Pt::new(11.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                color: RgbColor::BLACK,
                text_scale: 1.0,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        };

        let pdf_bytes = render_to_pdf(&[page], &registry, &RenderOptions::default())
            .expect("unicode text must not panic");
        assert_eq!(&pdf_bytes[..5], b"%PDF-");
    }

    // ── image cache keying ──────────────────────────────────────────

    fn rect(w: f32, h: f32) -> crate::render::geometry::PtRect {
        crate::render::geometry::PtRect::from_xywh(Pt::ZERO, Pt::ZERO, Pt::new(w), Pt::new(h))
    }

    /// A `size`×`size` PNG with a high-frequency checker pattern, so its
    /// encoded byte count scales clearly with the downsample resolution.
    fn textured_png(size: i32) -> Vec<u8> {
        use skia_safe::{AlphaType, ColorType, EncodedImageFormat, ImageInfo};
        let info = ImageInfo::new((size, size), ColorType::RGBA8888, AlphaType::Opaque, None);
        let mut surface = skia_safe::surfaces::raster(&info, None, None).expect("raster surface");
        let canvas = surface.canvas();
        canvas.clear(Color4f::new(1.0, 1.0, 1.0, 1.0));
        let cells = 32;
        let step = size as f32 / cells as f32;
        let mut paint = Paint::default();
        for gy in 0..cells {
            for gx in 0..cells {
                if (gx + gy) % 2 == 0 {
                    continue;
                }
                let shade = ((gx * 7 + gy * 13) % 255) as f32 / 255.0;
                paint.set_color4f(Color4f::new(shade, 1.0 - shade, shade * 0.5, 1.0), None);
                let r = skia_safe::Rect::from_xywh(gx as f32 * step, gy as f32 * step, step, step);
                canvas.draw_rect(r, &paint);
            }
        }
        surface
            .image_snapshot()
            .encode(None, EncodedImageFormat::PNG, None)
            .expect("encode png")
            .as_bytes()
            .to_vec()
    }

    #[test]
    fn render_to_pdf_image_dpi_controls_embedded_resolution() {
        // End-to-end: a large source image drawn small is embedded at the
        // target DPI. A higher DPI keeps more source pixels, so the PDF grows —
        // proving `image_dpi` threads through render_to_pdf into the downsample.
        use crate::model::ImageFormat;
        use crate::render::resolve::images::MediaEntry;

        let media = MediaEntry {
            data: std::sync::Arc::from(textured_png(1200).into_boxed_slice()),
            format: ImageFormat::Png,
        };
        let page = || LayoutedPage {
            commands: vec![DrawCommand::Image {
                rect: rect(72.0, 72.0),
                image_data: media.clone(),
                src_rect: None,
            }],
            page_size: PtSize::new(Pt::new(200.0), Pt::new(200.0)),
        };

        let registry = test_registry();
        let opts = |dpi| RenderOptions::default().with_image_dpi(dpi);
        let pdf_72 = render_to_pdf(&[page()], &registry, &opts(72.0)).expect("render at 72 dpi");
        let pdf_300 = render_to_pdf(&[page()], &registry, &opts(300.0)).expect("render at 300 dpi");

        assert!(
            pdf_300.len() > pdf_72.len(),
            "300 DPI PDF ({} bytes) should embed more image data than 72 DPI PDF ({} bytes)",
            pdf_300.len(),
            pdf_72.len()
        );
    }

    #[test]
    fn render_to_pdf_renders_clamped_low_dpi_without_degenerate_failure() {
        // `with_image_dpi` floors out-of-range requests to 1.0; the paint path
        // must then handle that floor — a 1×1 downsample target — without
        // panicking, and a clamped-from-bad request must render identically to
        // an explicit 1.0.
        use crate::model::ImageFormat;
        use crate::render::resolve::images::MediaEntry;

        let media = MediaEntry {
            data: std::sync::Arc::from(textured_png(1200).into_boxed_slice()),
            format: ImageFormat::Png,
        };
        let page = || LayoutedPage {
            commands: vec![DrawCommand::Image {
                rect: rect(72.0, 72.0),
                image_data: media.clone(),
                src_rect: None,
            }],
            page_size: PtSize::new(Pt::new(200.0), Pt::new(200.0)),
        };
        let registry = test_registry();

        let floor = render_to_pdf(
            &[page()],
            &registry,
            &RenderOptions::default().with_image_dpi(1.0),
        )
        .expect("render at floor");
        for bad in [-50.0f32, 0.0, f32::NAN, f32::NEG_INFINITY] {
            let opts = RenderOptions::default().with_image_dpi(bad);
            let got = render_to_pdf(&[page()], &registry, &opts).expect("must not panic");
            assert_eq!(
                &got[..5],
                b"%PDF-",
                "dpi request {bad} must produce a valid PDF"
            );
            assert_eq!(
                got.len(),
                floor.len(),
                "dpi request {bad} must clamp to the 1.0 floor"
            );
        }
    }

    #[test]
    fn downsample_target_scales_points_to_target_dpi() {
        // 72pt @ 72 DPI = 72px (1 px/pt); @ 300 DPI = 300px.
        assert_eq!(downsample_target(rect(72.0, 36.0), 72.0), (72, 36));
        assert_eq!(downsample_target(rect(72.0, 36.0), 300.0), (300, 150));
        assert_eq!(downsample_target(rect(72.0, 36.0), 150.0), (150, 75));
        // Sub-pixel targets round up so a placement never resolves to zero px.
        assert_eq!(downsample_target(rect(1.0, 1.0), 72.0), (1, 1));
    }

    #[test]
    fn quantize_crop_distinguishes_crops_but_matches_spec_identical_ones() {
        let a = crate::render::geometry::PtRect::from_xywh(
            Pt::new(0.25),
            Pt::new(0.25),
            Pt::new(0.5),
            Pt::new(0.5),
        );
        let b = crate::render::geometry::PtRect::from_xywh(
            Pt::new(0.1),
            Pt::new(0.1),
            Pt::new(0.8),
            Pt::new(0.8),
        );
        // Different crops key differently, so one placement can't reuse the
        // other's pre-cropped bitmap; an identical crop keys the same.
        assert_ne!(quantize_crop(&a), quantize_crop(&b));
        assert_eq!(quantize_crop(&a), quantize_crop(&a));
        // 50% == 50000 thousandth-percent.
        assert_eq!(quantize_crop(&a), (25000, 25000, 50000, 50000));
    }

    /// Opaque solid-color test image of the given pixel dimensions.
    fn solid_image(w: i32, h: i32) -> skia_safe::Image {
        use skia_safe::{AlphaType, ColorType, ImageInfo};
        let info = ImageInfo::new((w, h), ColorType::RGBA8888, AlphaType::Opaque, None);
        let mut surface = skia_safe::surfaces::raster(&info, None, None).expect("raster surface");
        surface.canvas().clear(Color4f::new(0.1, 0.2, 0.3, 1.0));
        surface.image_snapshot()
    }

    #[test]
    fn prepare_cropped_embeds_only_the_visible_region_not_the_whole_image() {
        // 1000×1000 source, cropped to the middle 10%×10% → a 100×100 visible
        // region. Displayed at 72pt (300px @ 300 DPI), but the visible region
        // only has 100px, so we don't upsample — the baked bitmap is 100×100,
        // NOT the full 1000×1000 (finding 3: no cropped-away pixels embedded).
        let src = solid_image(1000, 1000);
        let crop = crate::render::geometry::PtRect::from_xywh(
            Pt::new(0.45),
            Pt::new(0.45),
            Pt::new(0.1),
            Pt::new(0.1),
        );
        let out = prepare_cropped(&src, &crop, rect(72.0, 72.0), 300.0).expect("prepare");
        assert_eq!((out.width(), out.height()), (100, 100));
    }

    #[test]
    fn prepare_cropped_downsamples_a_high_res_crop_to_the_display_target() {
        // 4000×4000 source, cropped to the middle 50% → a 2000×2000 visible
        // region, displayed at 72pt. The region has more pixels than the
        // display needs (300px @ 300 DPI), so it downsamples to 300×300.
        let src = solid_image(4000, 4000);
        let crop = crate::render::geometry::PtRect::from_xywh(
            Pt::new(0.25),
            Pt::new(0.25),
            Pt::new(0.5),
            Pt::new(0.5),
        );
        let out = prepare_cropped(&src, &crop, rect(72.0, 72.0), 300.0).expect("prepare");
        assert_eq!((out.width(), out.height()), (300, 300));
    }

    // ── srcRect padding resolution ──────────────────────────────────

    fn xywh(x: f32, y: f32, w: f32, h: f32) -> crate::render::geometry::PtRect {
        crate::render::geometry::PtRect::from_xywh(Pt::new(x), Pt::new(y), Pt::new(w), Pt::new(h))
    }

    #[test]
    fn resolve_src_padding_is_a_no_op_for_an_in_bounds_crop() {
        // A crop fully inside [0,1] leaves dst unchanged and the crop as-is.
        let dst = rect(100.0, 50.0);
        let crop = xywh(0.1, 0.2, 0.5, 0.6);
        let (out_dst, out_crop) = resolve_src_padding(dst, &crop).expect("visible");
        assert_eq!(out_dst.origin.x.raw(), 0.0);
        assert_eq!(out_dst.origin.y.raw(), 0.0);
        assert_eq!(out_dst.size.width.raw(), 100.0);
        assert_eq!(out_dst.size.height.raw(), 50.0);
        assert!((out_crop.origin.x.raw() - 0.1).abs() < 1e-5);
        assert!((out_crop.size.width.raw() - 0.5).abs() < 1e-5);
    }

    #[test]
    fn resolve_src_padding_insets_dst_and_clamps_crop_for_negative_insets() {
        // srcRect region [-0.2, 1.2] horizontally (left = right = -0.2 → 20%
        // pillarbox padding each side), full height. The whole image is drawn
        // into the middle of dst; the side bands are left undrawn.
        let dst = xywh(0.0, 0.0, 100.0, 100.0);
        let src = xywh(-0.2, 0.0, 1.4, 1.0);
        let (out_dst, out_crop) = resolve_src_padding(dst, &src).expect("visible");
        // Clamped source region is the full image.
        assert!((out_crop.origin.x.raw() - 0.0).abs() < 1e-5);
        assert!((out_crop.size.width.raw() - 1.0).abs() < 1e-5);
        // dst inset by 0.2/1.4 ≈ 14.29% on the left, width 1.0/1.4 ≈ 71.43%.
        assert!((out_dst.origin.x.raw() - 100.0 * 0.2 / 1.4).abs() < 1e-3);
        assert!((out_dst.size.width.raw() - 100.0 * 1.0 / 1.4).abs() < 1e-3);
        // Vertical is untouched (region already in-bounds).
        assert_eq!(out_dst.origin.y.raw(), 0.0);
        assert_eq!(out_dst.size.height.raw(), 100.0);
    }

    #[test]
    fn resolve_src_padding_returns_none_when_region_is_wholly_outside() {
        // Region starts past the right edge of the image → nothing visible.
        let dst = rect(100.0, 100.0);
        let src = xywh(1.5, 0.0, 0.3, 1.0);
        assert!(resolve_src_padding(dst, &src).is_none());
    }

    // ─── Path fill mode (H1#3) and Tier-0 warn dedup (H1#5) ───────────────

    fn square_at(x: f32, mode: crate::model::PathFillMode) -> SubPath {
        SubPath {
            verbs: vec![
                PathVerb::MoveTo(PtOffset::new(Pt::new(x), Pt::ZERO)),
                PathVerb::LineTo(PtOffset::new(Pt::new(x + 10.0), Pt::ZERO)),
                PathVerb::LineTo(PtOffset::new(Pt::new(x + 10.0), Pt::new(10.0))),
                PathVerb::Close,
            ],
            fill_mode: mode,
            stroked: true,
        }
    }

    /// §20.1.10.45: `fill="none"` means the subpath has no interior. Filling
    /// it anyway paints the shape's fill colour over whatever sits behind the
    /// outline — the value was resolved end-to-end and then ignored here.
    #[test]
    fn build_skia_path_skips_fill_none_subpaths() {
        use crate::model::PathFillMode;
        let mut warned = FxHashSet::default();
        let only_none = build_skia_path(&[square_at(0.0, PathFillMode::None)], &mut warned);
        assert!(
            only_none.is_empty(),
            "a fill=none subpath must contribute no fillable geometry"
        );

        let mixed = build_skia_path(
            &[
                square_at(0.0, PathFillMode::Norm),
                square_at(100.0, PathFillMode::None),
            ],
            &mut warned,
        );
        assert!(
            mixed.bounds().right() <= 11.0,
            "only the fill=norm subpath may contribute; bounds reached {}              (the fill=none square sits at x=100)",
            mixed.bounds().right()
        );
    }

    /// The asymmetry is deliberate, not an oversight: `fill="none"` with
    /// `stroke="true"` is the ordinary outline-only subpath, so the stroke
    /// path must still carry it.
    #[test]
    fn stroked_path_still_includes_fill_none_subpaths() {
        use crate::model::PathFillMode;
        let stroked = build_skia_path_stroked_only(&[square_at(0.0, PathFillMode::None)]);
        assert!(
            !stroked.is_empty(),
            "fill=none is about the interior, not the outline"
        );
    }

    /// The shading variants are Tier-0: painted as `norm`, but reported —
    /// once per render, not once per shape.
    ///
    /// This pins the *report key*: that repeated shapes register one entry and
    /// that the key does not collide with another feature's. That the
    /// `log::warn!` itself sits inside the guard is structural — verifying it
    /// would need a process-global capturing logger, which parallel tests
    /// cannot share cleanly.
    #[test]
    fn shaded_fill_modes_are_filled_and_warned_once() {
        use crate::model::PathFillMode;
        let mut warned = FxHashSet::default();
        for _ in 0..3 {
            let p = build_skia_path(&[square_at(0.0, PathFillMode::Darken)], &mut warned);
            assert!(!p.is_empty(), "a shaded fill still fills at Tier 0");
        }
        assert_eq!(
            warned.iter().collect::<Vec<_>>(),
            vec![&"fillMode:darken"],
            "three shapes with the same mode must report once"
        );
    }

    /// A gradient-filled shape repeated in a header would otherwise warn on
    /// every page. Distinct kinds still each get their own report — a single
    /// shared key would silence two of the three. Same caveat as above: the
    /// key registry is what is pinned, not the macro's placement.
    #[test]
    fn unsupported_fills_warn_once_per_kind_per_render() {
        use crate::render::layout::draw_command::{
            ResolvedGradient, ResolvedGradientKind, ResolvedPattern,
        };
        let mut warned = FxHashSet::default();
        let gradient = ResolvedFill::Gradient(ResolvedGradient {
            stops: Vec::new(),
            kind: ResolvedGradientKind::Linear { angle_deg: 0.0 },
        });
        for _ in 0..4 {
            assert!(fill_to_paint(&gradient, &mut warned).is_none());
        }
        assert_eq!(warned.len(), 1, "one kind, one warning");
        let pattern = ResolvedFill::Pattern(ResolvedPattern {
            preset: crate::model::PresetPatternVal::Pct5,
            fg: Rgba {
                r: 0.0,
                g: 0.0,
                b: 0.0,
                a: 1.0,
            },
            bg: Rgba {
                r: 1.0,
                g: 1.0,
                b: 1.0,
                a: 1.0,
            },
        });
        assert!(fill_to_paint(&pattern, &mut warned).is_none());
        assert_eq!(warned.len(), 2, "a different kind is worth its own warning");
    }
}
