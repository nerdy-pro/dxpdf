//! Font resolution and caching with per-render ownership (`FontRegistry`).
//!
//! `FontRegistry` is the single source of truth for typeface data within
//! one render. It owns:
//!
//! - the document's embedded-font bytes (deobfuscated upstream by the
//!   parser per ECMA-376 §17.8.3.3),
//! - a cache of resolved Skia [`Typeface`]s keyed by (family, weight,
//!   slant), with embedded fonts taking priority over system resolution.
//!
//! A `FontRegistry` is constructed per render and passed by reference to
//! layout and paint. The previous `thread_local!` typeface cache leaked
//! typefaces across renders — once font subsetting mutates them, that
//! becomes a real correctness bug. With per-render ownership, no such
//! leakage is possible.

use std::cell::RefCell;
use std::collections::HashMap;

use skia_safe::{Data, Font, FontMgr, FontStyle, Typeface};

use crate::model::{EmbeddedFont, EmbeddedFontVariant};
use crate::render::dimension::Pt;

// ─── Public types ───────────────────────────────────────────────────────────

/// Stable id for an embedded font registered in the registry.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct EmbeddedFontId(u32);

impl EmbeddedFontId {
    pub fn raw(self) -> u32 {
        self.0
    }
}

/// Identity for a Skia [`Typeface`], wrapping `Typeface::unique_id`.
/// Used as the join key with [`crate::render::subset::GlyphUsage`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct TypefaceId(pub u32);

impl From<&Typeface> for TypefaceId {
    fn from(tf: &Typeface) -> Self {
        Self(tf.unique_id())
    }
}

/// Single source of truth for "where did this typeface come from?" — drives
/// byte extraction during subsetting (Embedded → registry's bytes, System →
/// `Typeface::to_font_data`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum TypefaceOrigin {
    /// Resolved from a font embedded in the DOCX (`word/fonts/*.odttf`).
    Embedded { id: EmbeddedFontId },
    /// Resolved through Skia's `FontMgr` — exact match, substitution, or
    /// system default fallback. The id is the original Skia typeface id
    /// at resolution time.
    System { typeface_id: TypefaceId },
}

#[derive(Clone, Debug)]
pub struct TypefaceEntry {
    pub typeface: Typeface,
    pub origin: TypefaceOrigin,
}

/// Cache key for resolved typefaces — case-insensitive family + weight + slant.
#[derive(Hash, Eq, PartialEq, Clone, Debug)]
pub struct TypefaceKey {
    pub family_lc: String,
    pub weight: i32,
    pub slant: skia_safe::font_style::Slant,
}

impl TypefaceKey {
    pub fn new(family: &str, style: FontStyle) -> Self {
        Self {
            family_lc: family.to_lowercase(),
            weight: *style.weight(),
            slant: style.slant(),
        }
    }
}

#[derive(thiserror::Error, Debug)]
pub enum RegisterError {
    #[error("invalid embedded font data for '{family}' ({variant:?})")]
    InvalidFontData {
        family: String,
        variant: EmbeddedFontVariant,
    },
}

/// Open-source metric-compatible substitutes for proprietary fonts. Tried
/// in order when `match_family_style` for the requested family fails.
const FONT_SUBSTITUTIONS: &[(&str, &[&str])] = &[
    ("Calibri", &["Carlito", "Liberation Sans", "Noto Sans"]),
    ("Cambria", &["Caladea", "Liberation Serif", "Noto Serif"]),
    ("Arial", &["Liberation Sans", "Noto Sans", "Helvetica"]),
    (
        "Times New Roman",
        &["Liberation Serif", "Noto Serif", "Times"],
    ),
    (
        "Courier New",
        &["Liberation Mono", "Noto Sans Mono", "Courier"],
    ),
    ("Verdana", &["DejaVu Sans", "Noto Sans"]),
    ("Georgia", &["DejaVu Serif", "Noto Serif"]),
    ("Trebuchet MS", &["Ubuntu", "Noto Sans"]),
    (
        "Consolas",
        &["Inconsolata", "Liberation Mono", "Noto Sans Mono"],
    ),
    ("Segoe UI", &["Noto Sans", "Liberation Sans"]),
];

#[derive(Debug, Clone)]
struct EmbeddedRecord {
    family: String,
    variant: EmbeddedFontVariant,
    bytes: Vec<u8>,
}

// ─── FontRegistry ────────────────────────────────────────────────────────────

pub struct FontRegistry {
    font_mgr: FontMgr,
    embedded: Vec<EmbeddedRecord>,
    embedded_index: HashMap<(String, EmbeddedFontVariant), EmbeddedFontId>,
    typefaces: RefCell<HashMap<TypefaceKey, TypefaceEntry>>,
}

impl FontRegistry {
    /// Empty registry without any embedded fonts.
    pub fn new(font_mgr: FontMgr) -> Self {
        Self {
            font_mgr,
            embedded: Vec::new(),
            embedded_index: HashMap::new(),
            typefaces: RefCell::new(HashMap::new()),
        }
    }

    /// Build a registry, registering all embedded fonts and preloading the
    /// requested family/style combinations.
    pub fn build(font_mgr: FontMgr, embedded: &[EmbeddedFont], families: &[String]) -> Self {
        let mut reg = Self::new(font_mgr);
        for ef in embedded {
            if let Err(err) = reg.register_embedded(&ef.family, ef.variant, ef.data.clone()) {
                log::warn!("{err}");
            }
        }
        reg.preload(families);
        reg
    }

    pub fn font_mgr(&self) -> &FontMgr {
        &self.font_mgr
    }

    pub fn embedded_font_count(&self) -> usize {
        self.embedded.len()
    }

    pub fn cached_typeface_count(&self) -> usize {
        self.typefaces.borrow().len()
    }

    /// Register an embedded font. Subsequent `resolve` calls for the same
    /// family + variant will return this typeface in preference to system
    /// resolution.
    pub fn register_embedded(
        &mut self,
        family: &str,
        variant: EmbeddedFontVariant,
        bytes: Vec<u8>,
    ) -> Result<EmbeddedFontId, RegisterError> {
        let data = Data::new_copy(&bytes);
        let typeface = self.font_mgr.new_from_data(&data, 0).ok_or_else(|| {
            RegisterError::InvalidFontData {
                family: family.to_string(),
                variant,
            }
        })?;
        let id = EmbeddedFontId(self.embedded.len() as u32);
        self.embedded.push(EmbeddedRecord {
            family: family.to_string(),
            variant,
            bytes,
        });
        self.embedded_index
            .insert((family.to_lowercase(), variant), id);
        let style = font_style_for_variant(variant);
        let key = TypefaceKey::new(family, style);
        self.typefaces.borrow_mut().insert(
            key,
            TypefaceEntry {
                typeface,
                origin: TypefaceOrigin::Embedded { id },
            },
        );
        log::debug!("registered embedded font '{}' {:?}", family, variant);
        Ok(id)
    }

    /// Bytes for a registered embedded font.
    pub fn embedded_bytes(&self, id: EmbeddedFontId) -> &[u8] {
        &self.embedded[id.0 as usize].bytes
    }

    /// Family + variant for a registered embedded font.
    pub fn embedded_meta(&self, id: EmbeddedFontId) -> (&str, EmbeddedFontVariant) {
        let r = &self.embedded[id.0 as usize];
        (&r.family, r.variant)
    }

    /// Resolve a typeface by family + style. Embedded fonts win over system.
    /// Cached after the first resolution; later calls are O(1).
    pub fn resolve(&self, family: &str, style: FontStyle) -> TypefaceEntry {
        let key = TypefaceKey::new(family, style);

        if let Some(entry) = self.typefaces.borrow().get(&key) {
            return entry.clone();
        }

        let entry = self.resolve_uncached(family, style);
        self.typefaces.borrow_mut().insert(key, entry.clone());
        entry
    }

    fn resolve_uncached(&self, family: &str, style: FontStyle) -> TypefaceEntry {
        let variant = variant_for_style(style);
        if let Some(id) = self
            .embedded_index
            .get(&(family.to_lowercase(), variant))
            .copied()
        {
            let bytes = &self.embedded[id.0 as usize].bytes;
            let data = Data::new_copy(bytes);
            if let Some(tf) = self.font_mgr.new_from_data(&data, 0) {
                log::debug!("[font] '{}' {:?} → embedded #{}", family, style, id.0);
                return TypefaceEntry {
                    typeface: tf,
                    origin: TypefaceOrigin::Embedded { id },
                };
            }
        }

        if let Some(tf) = match_exact(&self.font_mgr, family, style) {
            log::debug!("[font] '{}' {:?} → exact match", family, style);
            return system_entry(tf);
        }

        if let Some((_, subs)) = FONT_SUBSTITUTIONS
            .iter()
            .find(|(name, _)| name.eq_ignore_ascii_case(family))
        {
            for sub in *subs {
                if let Some(tf) = match_exact(&self.font_mgr, sub, style) {
                    log::debug!("[font] '{}' {:?} → substitute '{}'", family, style, sub);
                    return system_entry(tf);
                }
            }
        }

        let tf = self
            .font_mgr
            .legacy_make_typeface(None::<&str>, style)
            .expect("no fallback typeface available");
        log::debug!(
            "[font] '{}' {:?} → system default '{}'",
            family,
            style,
            tf.family_name()
        );
        system_entry(tf)
    }

    /// Resolve a typeface by exact family + style match, or `None` if the
    /// family is neither registered as an embedded font nor present in the
    /// host's font system.
    ///
    /// Unlike [`resolve`], this does not fall back to substitutes or to the
    /// system default — necessary for the emoji pipeline, where substituting
    /// a non-emoji typeface for a missing color emoji font is never correct.
    pub fn resolve_exact(&self, family: &str, style: FontStyle) -> Option<TypefaceEntry> {
        let variant = variant_for_style(style);
        if let Some(id) = self
            .embedded_index
            .get(&(family.to_lowercase(), variant))
            .copied()
        {
            let bytes = &self.embedded[id.0 as usize].bytes;
            let data = Data::new_copy(bytes);
            if let Some(tf) = self.font_mgr.new_from_data(&data, 0) {
                return Some(TypefaceEntry {
                    typeface: tf,
                    origin: TypefaceOrigin::Embedded { id },
                });
            }
        }
        match_exact(&self.font_mgr, family, style).map(system_entry)
    }

    /// Pre-resolve all four style variants for each family.
    pub fn preload(&self, families: &[String]) {
        let styles = [
            FontStyle::normal(),
            FontStyle::bold(),
            FontStyle::italic(),
            FontStyle::bold_italic(),
        ];
        for family in families {
            for &style in &styles {
                self.resolve(family, style);
            }
        }
    }

    /// Replace the typeface for every cached entry whose current id matches
    /// `old_id`. Returns the number of entries updated. Used by the font-
    /// subsetting pass to swap in subsetted bytes; multiple cache keys can
    /// share one underlying typeface (e.g. Calibri → Carlito substitution
    /// causes both keys to point at the same Skia typeface), so we update
    /// them all at once.
    pub fn replace_typeface_by_id(
        &mut self,
        old_id: TypefaceId,
        new_typeface: Typeface,
        new_origin: TypefaceOrigin,
    ) -> usize {
        let mut count = 0;
        for entry in self.typefaces.get_mut().values_mut() {
            if TypefaceId::from(&entry.typeface) == old_id {
                entry.typeface = new_typeface.clone();
                entry.origin = new_origin.clone();
                count += 1;
            }
        }
        count
    }

    /// Snapshot of all cached entries.
    pub fn cached_entries(&self) -> Vec<(TypefaceKey, TypefaceEntry)> {
        self.typefaces
            .borrow()
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect()
    }
}

// ─── Helpers ────────────────────────────────────────────────────────────────

fn font_style_for_variant(v: EmbeddedFontVariant) -> FontStyle {
    match v {
        EmbeddedFontVariant::Regular => FontStyle::normal(),
        EmbeddedFontVariant::Bold => FontStyle::bold(),
        EmbeddedFontVariant::Italic => FontStyle::italic(),
        EmbeddedFontVariant::BoldItalic => FontStyle::bold_italic(),
    }
}

fn variant_for_style(style: FontStyle) -> EmbeddedFontVariant {
    use skia_safe::font_style::{Slant, Weight};
    let bold = *style.weight() >= *Weight::SEMI_BOLD;
    let italic = matches!(style.slant(), Slant::Italic | Slant::Oblique);
    match (bold, italic) {
        (true, true) => EmbeddedFontVariant::BoldItalic,
        (true, false) => EmbeddedFontVariant::Bold,
        (false, true) => EmbeddedFontVariant::Italic,
        (false, false) => EmbeddedFontVariant::Regular,
    }
}

fn match_exact(font_mgr: &FontMgr, family: &str, style: FontStyle) -> Option<Typeface> {
    let tf = font_mgr.match_family_style(family, style)?;
    if tf.family_name().eq_ignore_ascii_case(family) {
        Some(tf)
    } else {
        None
    }
}

fn system_entry(tf: Typeface) -> TypefaceEntry {
    let id = TypefaceId::from(&tf);
    TypefaceEntry {
        typeface: tf,
        origin: TypefaceOrigin::System { typeface_id: id },
    }
}

// ─── FontCache (per-component, not per-render) ──────────────────────────────

#[derive(Hash, Eq, PartialEq)]
struct FontKey {
    family: String,
    /// Font size stored as bits for exact f32 hashing.
    size_bits: u32,
    weight: i32,
    slant: skia_safe::font_style::Slant,
}

/// Per-component cache of fully-configured `Font` objects, avoiding repeated
/// `FontRegistry::resolve` lookups and `Font::from_typeface` construction.
///
/// Must be discarded if the underlying `FontRegistry` is mutated (e.g. by
/// `replace_typeface_by_id`). The render pipeline already creates a fresh
/// `FontCache` for layout and another for paint, with the subset pass in
/// between; no stale Font objects can survive.
#[derive(Default)]
pub struct FontCache {
    cache: HashMap<FontKey, Font>,
}

impl FontCache {
    pub fn new() -> Self {
        Self::default()
    }

    /// Get or create a `Font` for the given properties.
    pub fn get(
        &mut self,
        registry: &FontRegistry,
        font_family: &str,
        font_size: Pt,
        bold: bool,
        italic: bool,
    ) -> &Font {
        let style = match (bold, italic) {
            (true, true) => FontStyle::bold_italic(),
            (true, false) => FontStyle::bold(),
            (false, true) => FontStyle::italic(),
            (false, false) => FontStyle::normal(),
        };
        let key = FontKey {
            family: font_family.to_lowercase(),
            size_bits: f32::from(font_size).to_bits(),
            weight: *style.weight(),
            slant: style.slant(),
        };
        self.cache.entry(key).or_insert_with_key(|k| {
            let style = FontStyle::new(
                skia_safe::font_style::Weight::from(k.weight),
                skia_safe::font_style::Width::NORMAL,
                k.slant,
            );
            let entry = registry.resolve(font_family, style);
            let mut font = Font::from_typeface(entry.typeface, f32::from(font_size));
            font.set_subpixel(true);
            font.set_linear_metrics(true);
            font.set_hinting(skia_safe::FontHinting::None);
            font
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fmgr() -> FontMgr {
        FontMgr::new()
    }

    /// Pull bytes from a guaranteed-available system typeface so tests don't
    /// need bundled font fixtures for the registry-level invariants.
    fn arbitrary_system_font_bytes() -> Vec<u8> {
        let mgr = fmgr();
        let tf = mgr
            .legacy_make_typeface(None::<&str>, FontStyle::normal())
            .expect("system has no default typeface — cannot run test");
        let (bytes, _ttc_index) = tf
            .to_font_data()
            .expect("legacy default typeface lacks raw font bytes — cannot run test");
        bytes
    }

    #[test]
    fn registry_empty_after_construction() {
        let r = FontRegistry::new(fmgr());
        assert_eq!(r.embedded_font_count(), 0);
        assert_eq!(r.cached_typeface_count(), 0);
    }

    #[test]
    fn registry_resolves_system_font_idempotently() {
        let r = FontRegistry::new(fmgr());
        let a = r.resolve("DefinitelyNotInstalledXYZ", FontStyle::normal());
        let b = r.resolve("DefinitelyNotInstalledXYZ", FontStyle::normal());
        assert_eq!(
            TypefaceId::from(&a.typeface),
            TypefaceId::from(&b.typeface),
            "second resolution must hit the cache and yield the same typeface"
        );
        assert!(matches!(a.origin, TypefaceOrigin::System { .. }));
    }

    #[test]
    fn registry_embedded_takes_precedence_over_system() {
        let bytes = arbitrary_system_font_bytes();

        let mut without = FontRegistry::new(fmgr());
        let baseline = without.resolve("NonexistentFamilyABC", FontStyle::normal());
        assert!(
            matches!(baseline.origin, TypefaceOrigin::System { .. }),
            "without embedding, an unknown family must fall back to the system path"
        );
        // Quiet the unused-mut warning while making the contrast explicit.
        let _ = &mut without;

        let mut with = FontRegistry::new(fmgr());
        let id = with
            .register_embedded("NonexistentFamilyABC", EmbeddedFontVariant::Regular, bytes)
            .expect("register_embedded should accept a valid system font's bytes");
        let resolved = with.resolve("NonexistentFamilyABC", FontStyle::normal());
        assert_eq!(
            resolved.origin,
            TypefaceOrigin::Embedded { id },
            "after registration, resolution must return the embedded origin"
        );
    }

    #[test]
    fn registry_stores_embedded_bytes_byte_identical() {
        // ECMA-376 §17.8.3.3 deobfuscation is enforced upstream by the parser
        // (see src/docx/parse/fonts.rs::deobfuscate_round_trip). The registry-
        // level invariant is that the bytes it stores must be byte-identical
        // to the bytes handed in — no re-encoding, no normalization.
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        let id = r
            .register_embedded(
                "ByteIdentityProbe",
                EmbeddedFontVariant::Regular,
                bytes.clone(),
            )
            .expect("registration should succeed for valid font bytes");
        assert_eq!(
            r.embedded_bytes(id),
            bytes.as_slice(),
            "stored bytes must match the originally passed-in bytes byte-for-byte"
        );
    }

    #[test]
    fn registry_drop_clears_all_typefaces() {
        // Two registries on the same thread must not share state — this is
        // the structural fix for the cross-render poisoning that the previous
        // thread_local!-backed cache caused.
        let r1 = FontRegistry::new(fmgr());
        let _ = r1.resolve("FamilyOne", FontStyle::normal());
        assert_eq!(r1.cached_typeface_count(), 1);
        drop(r1);

        let r2 = FontRegistry::new(fmgr());
        assert_eq!(
            r2.cached_typeface_count(),
            0,
            "a fresh registry must not see typefaces cached by an earlier one"
        );
    }

    #[test]
    fn registry_resolution_records_origin() {
        // Every resolved entry must report a non-default origin variant.
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("OriginEmbedded", EmbeddedFontVariant::Regular, bytes)
            .unwrap();
        let _ = r.resolve("OriginEmbedded", FontStyle::normal());
        let _ = r.resolve("OriginSystemFallbackXYZ", FontStyle::normal());

        for (_, entry) in r.cached_entries() {
            match entry.origin {
                TypefaceOrigin::Embedded { .. } | TypefaceOrigin::System { .. } => {}
            }
        }
    }

    #[test]
    fn replace_typeface_by_id_updates_all_keys_pointing_at_it() {
        // Two distinct cache keys can share one underlying typeface (e.g. via
        // FONT_SUBSTITUTIONS). The subsetting pass must update them all so
        // paint never sees a stale typeface.
        let mut r = FontRegistry::new(fmgr());

        // Two unknown families → both fall back to the system default → same
        // underlying Skia typeface.
        let a = r.resolve("UnknownFamilyA", FontStyle::normal());
        let b = r.resolve("UnknownFamilyB", FontStyle::normal());
        assert_eq!(
            TypefaceId::from(&a.typeface),
            TypefaceId::from(&b.typeface),
            "test setup precondition — both unknowns must resolve to the same default"
        );
        let shared_id = TypefaceId::from(&a.typeface);

        // Build a *different* typeface to swap in.
        let bytes = arbitrary_system_font_bytes();
        let data = Data::new_copy(&bytes);
        let replacement = r
            .font_mgr()
            .new_from_data(&data, 0)
            .expect("replacement typeface should construct from valid bytes");
        let replacement_origin = TypefaceOrigin::System {
            typeface_id: TypefaceId::from(&replacement),
        };

        let updated = r.replace_typeface_by_id(shared_id, replacement, replacement_origin);
        assert_eq!(
            updated, 2,
            "both shared-typeface entries must be updated in lockstep"
        );

        // Subsequent resolution returns the new typeface, not the old one.
        let after = r.resolve("UnknownFamilyA", FontStyle::normal());
        assert_ne!(
            TypefaceId::from(&after.typeface),
            shared_id,
            "post-replace resolution must yield the new typeface"
        );
    }

    #[test]
    fn font_cache_uses_registry_after_replacement() {
        // The cross-cutting invariant: if the registry is mutated and a
        // *new* FontCache is created afterwards, that cache must produce
        // Fonts backed by the new typeface. (Stale FontCaches must be
        // discarded — that contract is satisfied by the pipeline creating
        // fresh caches around the subset pass.)
        let mut r = FontRegistry::new(fmgr());
        let original = r.resolve("CacheReplaceProbe", FontStyle::normal());
        let original_id = TypefaceId::from(&original.typeface);

        let bytes = arbitrary_system_font_bytes();
        let data = Data::new_copy(&bytes);
        let replacement = r
            .font_mgr()
            .new_from_data(&data, 0)
            .expect("replacement typeface should construct");
        let new_id = TypefaceId::from(&replacement);
        assert_ne!(
            new_id, original_id,
            "test precondition — replacement must be a different typeface"
        );
        let replacement_origin = TypefaceOrigin::System {
            typeface_id: new_id,
        };
        r.replace_typeface_by_id(original_id, replacement, replacement_origin);

        let mut fresh_cache = FontCache::new();
        let font = fresh_cache.get(&r, "CacheReplaceProbe", Pt::new(12.0), false, false);
        assert_eq!(
            TypefaceId::from(&font.typeface()),
            new_id,
            "FontCache must observe the post-replacement typeface"
        );
    }

    #[test]
    fn resolve_exact_returns_none_for_unknown_family() {
        let r = FontRegistry::new(fmgr());
        assert!(
            r.resolve_exact("DefinitelyNotInstalledXYZ", FontStyle::normal())
                .is_none(),
            "unlike resolve(), resolve_exact must not invent a fallback"
        );
    }

    #[test]
    fn resolve_exact_returns_some_for_embedded_family() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("ExactProbe", EmbeddedFontVariant::Regular, bytes)
            .expect("register_embedded should accept valid font bytes");
        let entry = r
            .resolve_exact("ExactProbe", FontStyle::normal())
            .expect("embedded font must be resolvable via exact match");
        assert!(matches!(entry.origin, TypefaceOrigin::Embedded { .. }));
    }
}
