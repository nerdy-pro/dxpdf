//! Font resolution and caching with per-render ownership (`FontRegistry`).
//!
//! `FontRegistry` is the single source of truth for typeface data within one
//! render. It owns:
//!
//! - the document's embedded-font bytes (deobfuscated upstream by the parser
//!   per ECMA-376 §17.8.3.3),
//! - a [`FaceCatalog`] of everything selectable, host and embedded alike,
//! - a cache of resolved [`Typeface`]s keyed by the *request* that produced
//!   them ([`FaceRequestKey`]),
//! - a cache of opened [`Typeface`]s keyed by the *face* itself, so two
//!   different requests that land on the same face share one typeface — see
//!   the `faces` field.
//!
//! # How a name becomes a typeface
//!
//! 1. [`request`] turns a §17.7.2 property cascade into a [`FaceRequest`] —
//!    a name plus two tri-state toggles.
//! 2. [`catalog`] describes what the host and the document offer, reading each
//!    face's own OpenType tables through [`opentype`] when a request needs more
//!    than a family name.
//! 3. [`resolve`] runs the eight-step chain and names one [`FaceRecord`].
//! 4. This module opens it — positionally, and where the record is a named
//!    instance, baking its coordinates into a static SFNT (`instance_bake`)
//!    so embedding is correct too, not just measurement and paint. A named
//!    boundary (CFF2, WOFF/WOFF2, a `spec_next` cubic outline) falls back to
//!    `Typeface::clone_with_arguments` instead — see [`InstanceState`].
//!
//! Only step 4 touches Skia beyond reading tables, which is what lets the first
//! three be tested against fixture fonts identically on every platform.
//!
//! # Per-render ownership
//!
//! A `FontRegistry` is constructed per render and passed by reference to layout
//! and paint. The previous `thread_local!` typeface cache leaked typefaces
//! across renders — once font subsetting mutates them, that becomes a real
//! correctness bug. With per-render ownership, no such leakage is possible.

mod cache;
pub mod catalog;
pub mod face;
mod instance_bake;
pub mod opentype;
pub mod request;
pub mod resolve;
mod substitutes;

pub use cache::FontCache;
pub use catalog::{FaceCatalog, FaceLookup, FaceRef};
pub use face::{FaceIdentity, FaceName, FaceRecord, IntrinsicStyle, NameKind, VariationInstance};
pub use instance_bake::InstanceBakeFailure;
pub use request::{EffectiveStyle, FaceRequest, Synthesis, Toggle};
pub use resolve::{FaceResolution, FallbackReason, ResolutionStep, Selection};

use std::cell::{Cell, RefCell};
use std::collections::{HashMap, HashSet};

use skia_safe::font_style::{Slant, Weight};
use skia_safe::{FontMgr, FontStyle, Typeface};

use crate::model::{EmbeddedFont, EmbeddedFontVariant};
use opentype::FontFormat;

// ─── Public types ───────────────────────────────────────────────────────────

/// Stable id for an embedded font registered in the registry.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct EmbeddedFontId(u32);

impl EmbeddedFontId {
    pub fn raw(self) -> u32 {
        self.0
    }

    /// Rebuild from a raw index.
    ///
    /// Test-only: in production every id comes from
    /// [`FontRegistry::register_embedded`], and a face record only ever names
    /// one the registry minted. Tests that build a [`face::FaceRecord`] by hand
    /// still need to name an embedded font that does not exist.
    #[cfg(test)]
    pub(crate) fn from_raw(raw: u32) -> Self {
        Self(raw)
    }
}

/// Identity for a Skia [`Typeface`], wrapping `Typeface::unique_id`.
/// Used as the join key with [`crate::render::subset::CodepointUsage`].
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
    ///
    /// `collection_index` selects the face when those bytes are a TrueType
    /// Collection. OOXML cannot express one — `fontTable.xml` offers only four
    /// style slots per `w:font/@w:name` — so it is discovered by reading the
    /// bytes, and subsetting needs it to know which face to carve out.
    Embedded {
        id: EmbeddedFontId,
        collection_index: u32,
    },
    /// Resolved through Skia's `FontMgr` — exact match, substitution, or
    /// system default fallback. The id is the original Skia typeface id
    /// at resolution time.
    System { typeface_id: TypefaceId },
}

#[derive(Clone, Debug)]
pub struct TypefaceEntry {
    pub typeface: Typeface,
    pub origin: TypefaceOrigin,
    /// Whether `typeface` is a variable-font named instance, and if so,
    /// whether its coordinates are baked into `origin`'s bytes.
    pub instance: InstanceState,
    /// What resolution could not get from a real face for this typeface —
    /// issue #115. `FontCache` is what turns it into an actual
    /// `set_embolden`/`set_skew_x`. Defaults to
    /// [`Synthesis::default`] (nothing to synthesise) here: this is the
    /// **face**-keyed constructor, opened independent of any particular
    /// request's toggles — `resolve_uncached`/`host_default` are what
    /// overwrite it with the value a specific request's `Synthesis::compute`
    /// produces, since only they know what was asked.
    pub synthesis: Synthesis,
}

impl TypefaceEntry {
    fn system(typeface: Typeface) -> Self {
        let typeface_id = TypefaceId::from(&typeface);
        Self {
            typeface,
            origin: TypefaceOrigin::System { typeface_id },
            instance: InstanceState::NotInstanced,
            synthesis: Synthesis::default(),
        }
    }
}

/// The three states a resolved typeface can be in with respect to a variable
/// font's design space — replaces what used to be `Option<VariationInstance>`,
/// which could not distinguish "not an instance" from "an instance, unbaked."
///
/// Carried on [`TypefaceEntry`] rather than recovered from `typeface` itself
/// because the subsetting pass has to *know*, without re-deriving it, whether
/// the bytes it is about to embed already contain the pinned location — see
/// `SubsetOutcome::VariableInstanceNotBaked`.
#[derive(Clone, Debug, PartialEq)]
pub enum InstanceState {
    /// Not a variable-font instance — either an ordinary static face, or the
    /// font system declined to instance one at all (§113's own live-instance
    /// fallback, `apply_variation`, returning `None`; rare, and no worse than
    /// resolution failing to find the family in the first place).
    NotInstanced,
    /// `typeface`'s embedded bytes ARE this location — subsetting needs no
    /// special case for it.
    Baked(VariationInstance),
    /// `typeface` is still `Typeface::clone_with_arguments`-instanced (screen-
    /// and measurement-correct); its embedded bytes remain the font's default
    /// location, for the stated reason.
    Unbaked {
        instance: VariationInstance,
        reason: InstanceBakeFailure,
    },
}

/// Cache key for a resolved face: the *request*, not a Skia style.
///
/// This is the shape that keeps resolution deterministic across the two places
/// it happens. Layout and paint each ask the registry independently, and both
/// must arrive at the same face — including its collection index and any
/// variation coordinates, neither of which a `FontStyle` can express. Keying on
/// the request instead means resolution is a pure function of things the key
/// carries, so a cache hit is exactly as correct as a fresh resolution.
///
/// All three states of each toggle are kept even though [`Toggle::Off`] and
/// [`Toggle::Absent`] currently select the same face: normalising them here
/// would bake that equivalence into the cache, where a later change to
/// [`EffectiveStyle`] could not undo it. The cost is at most nine entries per
/// family.
#[derive(Hash, Eq, PartialEq, Clone, Debug)]
pub struct FaceRequestKey {
    pub name_lc: String,
    pub bold: Toggle,
    pub italic: Toggle,
}

impl FaceRequestKey {
    pub fn new(request: &FaceRequest<'_>) -> Self {
        Self {
            name_lc: face::fold_name(request.name),
            bold: request.bold,
            italic: request.italic,
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

/// Identity of an *opened* face: which face, and at which design-space
/// location.
///
/// The location is part of the key because instancing produces a genuinely
/// different Skia typeface. Coordinates are keyed on their bit patterns rather
/// than their values — `f32` is not `Hash`, and two coordinates that are bitwise
/// equal are the same location by any definition that matters here.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct OpenFaceKey {
    identity: FaceIdentity,
    location: Vec<([u8; 4], u32)>,
}

impl OpenFaceKey {
    fn new(record: &FaceRecord) -> Self {
        Self {
            identity: record.identity.clone(),
            location: record
                .instance
                .as_ref()
                .map(|instance| {
                    instance
                        .coords
                        .iter()
                        .map(|coord| (coord.axis.0, coord.value.to_bits()))
                        .collect()
                })
                .unwrap_or_default(),
        }
    }
}

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
    catalog: FaceCatalog,
    typefaces: RefCell<HashMap<FaceRequestKey, TypefaceEntry>>,
    /// One Skia typeface per catalogue face, so a face reached by two different
    /// requests is the *same* typeface.
    ///
    /// Not an optimisation. [`TypefaceId`] is how the subsetting pass
    /// deduplicates, so building a fresh typeface per request would embed one
    /// font several times over — observed on `sample-docx-files-sample1.docx`,
    /// where `<w:b/>` and `<w:b/><w:i w:val="0"/>` select the same embedded
    /// Ubuntu Bold and the PDF gained a second copy of it.
    faces: RefCell<HashMap<OpenFaceKey, TypefaceEntry>>,
    /// Families already reported as unresolvable, so the warning fires once per
    /// family rather than once per style variant. [`Self::preload`] alone would
    /// otherwise emit four identical lines for every missing font.
    warned: RefCell<HashSet<String>>,
    /// `legacy_make_typeface` answers, keyed by the effective style — the host
    /// default is the same face for every request with the same toggles.
    host_defaults: RefCell<HashMap<(i32, i32, i32), Typeface>>,
    /// How many times the host was actually asked; the memo's test reads it.
    host_default_fetches: Cell<usize>,
}

impl FontRegistry {
    /// Empty registry without any embedded fonts.
    pub fn new(font_mgr: FontMgr) -> Self {
        Self {
            catalog: FaceCatalog::build(font_mgr.clone()),
            font_mgr,
            embedded: Vec::new(),
            typefaces: RefCell::new(HashMap::new()),
            faces: RefCell::new(HashMap::new()),
            warned: RefCell::new(HashSet::new()),
            host_defaults: RefCell::new(HashMap::new()),
            host_default_fetches: Cell::new(0),
        }
    }

    /// Build a registry, registering all embedded fonts and preloading the
    /// requested family/style combinations.
    ///
    /// Fails with [`crate::render::error::RenderError::NoFontsAvailable`] when the host exposes no
    /// typeface at all. Checking here rather than at the point of use is what
    /// lets [`Self::resolve`] return a `TypefaceEntry` rather than an
    /// `Option`: the last-resort arm of `resolve_uncached` is unreachable for
    /// any registry this constructor returns.
    pub fn build(
        font_mgr: FontMgr,
        embedded: &[EmbeddedFont],
        families: &[String],
    ) -> Result<Self, crate::render::error::RenderError> {
        if font_mgr
            .legacy_make_typeface(None::<&str>, FontStyle::normal())
            .is_none()
        {
            return Err(crate::render::error::RenderError::NoFontsAvailable);
        }
        let mut reg = Self::new(font_mgr);
        for ef in embedded {
            if let Err(err) = reg.register_embedded(&ef.family, ef.variant, ef.data.clone()) {
                log::warn!("{err}");
            }
        }
        reg.preload(families);
        Ok(reg)
    }

    pub fn font_mgr(&self) -> &FontMgr {
        &self.font_mgr
    }

    pub fn catalog(&self) -> &FaceCatalog {
        &self.catalog
    }

    pub fn embedded_font_count(&self) -> usize {
        self.embedded.len()
    }

    pub fn cached_typeface_count(&self) -> usize {
        self.typefaces.borrow().len()
    }

    /// Register an embedded font.
    ///
    /// One call may contribute several catalogue faces: a TrueType Collection
    /// carries many, and OOXML has no way to say so, so the face count comes
    /// from the bytes. Every face is indexed under the document's declared
    /// `w:font/@w:name` as well as its own `name` table, because the declared
    /// name is what the document will ask for.
    ///
    /// The `variant` slot is recorded but not used to select: the four legacy
    /// slots cannot express a Semibold, and each face's own `OS/2` can, so
    /// ranking reads the font rather than the slot it was declared in. See
    /// [`Self::embedded_meta`].
    ///
    /// Every face of a collection is carved into a standalone SFNT
    /// (`open_embedded_typeface`) before the font system ever sees it, so no
    /// platform is asked to interpret a collection index directly — Skia's
    /// CoreText backend declines any index but 0. Only a font that yields no
    /// face at all — carving failed, or the platform declined even the
    /// carved bytes — is an error.
    pub fn register_embedded(
        &mut self,
        family: &str,
        variant: EmbeddedFontVariant,
        bytes: Vec<u8>,
    ) -> Result<EmbeddedFontId, RegisterError> {
        let invalid = || RegisterError::InvalidFontData {
            family: family.to_string(),
            variant,
        };

        // A collection's face count is in its header; anything else is one face.
        let face_count = match FontFormat::detect(&bytes) {
            Ok(FontFormat::Ttc { face_count }) => face_count,
            Ok(_) => 1,
            Err(err) => {
                log::debug!("[font] embedded '{family}' has unrecognised bytes ({err})");
                1
            }
        };

        let id = EmbeddedFontId(self.embedded.len() as u32);
        let mut registered = 0;
        for collection_index in 0..face_count {
            let Some(typeface) = self.open_embedded_typeface(&bytes, collection_index) else {
                continue;
            };
            self.catalog
                .add_embedded_face(id, collection_index, family, &typeface);
            registered += 1;
        }
        if registered == 0 {
            return Err(invalid());
        }

        self.embedded.push(EmbeddedRecord {
            family: family.to_string(),
            variant,
            bytes,
        });
        log::debug!(
            "[font] registered embedded '{family}' {variant:?} ({registered} face(s), id {})",
            id.0
        );
        Ok(id)
    }

    /// Open one face of embedded font bytes, carving a collection face into a
    /// standalone SFNT first (issue #116) — no platform is ever asked to
    /// interpret a non-zero collection index, which Skia's CoreText backend
    /// declines outright (`new_from_data(bytes, 1)` returns `None` on macOS).
    ///
    /// Used both at registration, to catalogue every face, and at
    /// [`Self::open_uncached`], to reopen a face resolution has already
    /// selected — carving is cheap enough to redo (collections in DOCX are
    /// rare and small) and keeps this the single place either call site
    /// needs to know the CoreText limitation exists at all.
    fn open_embedded_typeface(&self, bytes: &[u8], collection_index: u32) -> Option<Typeface> {
        let face_count = match FontFormat::detect(bytes) {
            Ok(FontFormat::Ttc { face_count }) => face_count,
            _ => 1,
        };
        if face_count <= 1 {
            return self.font_mgr.new_from_data(bytes, 0);
        }
        match opentype::carve_collection_face(bytes, collection_index, face_count) {
            Ok(carved) => self.font_mgr.new_from_data(&carved.bytes, 0),
            Err(err) => {
                log::debug!(
                    "[font] could not carve collection face {collection_index} of \
                     {face_count} ({err})"
                );
                None
            }
        }
    }

    /// Bytes for a registered embedded font.
    pub fn embedded_bytes(&self, id: EmbeddedFontId) -> &[u8] {
        &self.embedded[id.0 as usize].bytes
    }

    /// The `w:font/@w:name` and `w:embed*` slot a font was *declared* under.
    ///
    /// Not what resolution uses: an embedded font's faces are ranked by their
    /// own `OS/2` weight and slant like any other, because the four-slot model
    /// cannot express a Semibold and the font's own tables can. This is the
    /// document's assertion about the file, kept because it is the only record
    /// of what the producer thought it was embedding.
    pub fn embedded_meta(&self, id: EmbeddedFontId) -> (&str, EmbeddedFontVariant) {
        let r = &self.embedded[id.0 as usize];
        (&r.family, r.variant)
    }

    /// Resolve a face for a request. Cached after the first resolution; later
    /// calls are O(1).
    pub fn resolve(&self, request: &FaceRequest<'_>) -> TypefaceEntry {
        let key = FaceRequestKey::new(request);
        if let Some(entry) = self.typefaces.borrow().get(&key) {
            return entry.clone();
        }
        let entry = self.resolve_uncached(request);
        self.typefaces.borrow_mut().insert(key, entry.clone());
        entry
    }

    /// Convenience for callers that still think in `(family, bold, italic)`.
    pub fn resolve_flags(&self, family: &str, bold: bool, italic: bool) -> TypefaceEntry {
        self.resolve(&FaceRequest::new(
            family,
            if bold { Toggle::On } else { Toggle::Absent },
            if italic { Toggle::On } else { Toggle::Absent },
        ))
    }

    fn resolve_uncached(&self, request: &FaceRequest<'_>) -> TypefaceEntry {
        match resolve::resolve(request, &self.catalog) {
            resolve::FaceResolution::Selected(selection) => {
                if let Some(mut entry) = self.open(&selection.record) {
                    log::debug!(
                        "[font] '{}' {:?}/{:?} → {} '{}'{}",
                        request.name,
                        request.bold,
                        request.italic,
                        selection.step.label(),
                        selection.record.canonical_family,
                        selection
                            .record
                            .instance
                            .as_ref()
                            .map(|i| format!(" instance '{}'", i.subfamily))
                            .unwrap_or_default(),
                    );
                    // §115: `open` is cached by face identity, shared across
                    // every request that resolves to this face — a
                    // `Synthesis` computed *there* would be one request's
                    // decision leaking into another's cache hit. Computed
                    // here instead, against this request's own toggles, and
                    // set on the (already-cloned) entry this call owns.
                    entry.synthesis = Synthesis::compute(request, &selection.record.intrinsic);
                    return entry;
                }
                // The catalogue named a face the manager then declined to open.
                // Rare, and worth saying so plainly rather than silently
                // producing the default.
                self.warn_once(
                    request.name,
                    format_args!(
                        "matched '{}' but the font system would not open it",
                        selection.record.canonical_family
                    ),
                );
                self.host_default(request)
            }
            resolve::FaceResolution::Default { reason } => {
                match &reason {
                    resolve::FallbackReason::NoCandidate => self.warn_once(
                        request.name,
                        format_args!("no installed font and no substitute matched"),
                    ),
                    resolve::FallbackReason::AmbiguousName { step, candidates } => self.warn_once(
                        request.name,
                        format_args!(
                            "{candidates} different faces answer to this name at the {} step, \
                             so no face could be chosen",
                            step.label()
                        ),
                    ),
                }
                self.host_default(request)
            }
        }
    }

    /// Step 8: whatever the host offers when nothing matched.
    fn host_default(&self, request: &FaceRequest<'_>) -> TypefaceEntry {
        // Unreachable for a registry from `FontRegistry::build`, which rejects a
        // font-less `FontMgr` up front so this arm always has something to
        // return. A registry made with `FontRegistry::new` carries no such
        // guarantee — that constructor is for tests, which supply a real
        // `FontMgr`.
        let style = EffectiveStyle::resolve(request, IntrinsicStyle::NEUTRAL);
        // One host ask per effective style per render: `legacy_make_typeface`
        // is an `FcFontMatch` plus an `FT_Open_Face` on fontconfig, and every
        // unresolvable family in a document reached it four times (issue #188).
        let font_style = style.font_style();
        let key = (
            *font_style.weight(),
            *font_style.width(),
            font_style.slant() as i32,
        );
        let cached = self.host_defaults.borrow().get(&key).cloned();
        let tf = match cached {
            Some(tf) => tf,
            None => {
                self.host_default_fetches
                    .set(self.host_default_fetches.get() + 1);
                let tf = self
                    .font_mgr
                    .legacy_make_typeface(None::<&str>, font_style)
                    .expect("FontRegistry::build guarantees a last-resort typeface");
                self.host_defaults.borrow_mut().insert(key, tf.clone());
                tf
            }
        };
        log::debug!(
            "[font] '{}' → host default '{}'",
            request.name,
            tf.family_name()
        );
        let mut entry = TypefaceEntry::system(tf);
        // The *returned* typeface's own reported style, not the style
        // requested of `legacy_make_typeface` — the two can differ (the host
        // has no italic variant, hands back upright anyway), and it is what
        // the face actually *is* that decides whether synthesis is needed.
        let selected = IntrinsicStyle::from_font_style(entry.typeface.font_style());
        entry.synthesis = Synthesis::compute(request, &selected);
        entry
    }

    /// How many times the host was asked for its default face.
    pub fn host_default_fetches(&self) -> usize {
        self.host_default_fetches.get()
    }

    /// One warning per requested family, however many style variants ask.
    ///
    /// [`Self::preload`] resolves four combinations of every family the document
    /// mentions, so without the deduplication a single missing font would
    /// produce four identical lines.
    fn warn_once(&self, family: &str, detail: std::fmt::Arguments<'_>) {
        if self.warned.borrow_mut().insert(family.to_lowercase()) {
            log::warn!("[font] '{family}': {detail}; falling back to the host default");
        }
    }

    /// Open the Skia typeface a catalogue record names.
    ///
    /// Positional in both arms — a family plus a face index for a host font, a
    /// collection index for an embedded one. Reopening by `(family, style)`
    /// would re-run the host's matching and could hand back a different face
    /// than the one resolution chose, which is exactly the failure acceptance
    /// item 7 is about.
    fn open(&self, record: &FaceRecord) -> Option<TypefaceEntry> {
        let key = OpenFaceKey::new(record);
        if let Some(entry) = self.faces.borrow().get(&key) {
            return Some(entry.clone());
        }
        let entry = self.open_uncached(record)?;
        self.faces.borrow_mut().insert(key, entry.clone());
        Some(entry)
    }

    fn open_uncached(&self, record: &FaceRecord) -> Option<TypefaceEntry> {
        let (base, origin) = match &record.identity {
            FaceIdentity::Embedded {
                font,
                collection_index,
            } => {
                let bytes = self.embedded_bytes(*font);
                let typeface = self.open_embedded_typeface(bytes, *collection_index)?;
                (
                    typeface,
                    TypefaceOrigin::Embedded {
                        id: *font,
                        collection_index: *collection_index,
                    },
                )
            }
            FaceIdentity::System { family, face_index } => {
                let mut style_set = self.font_mgr.match_family(family);
                let typeface = style_set.new_typeface(*face_index)?;
                let typeface_id = TypefaceId::from(&typeface);
                (typeface, TypefaceOrigin::System { typeface_id })
            }
        };

        let Some(instance) = &record.instance else {
            return Some(TypefaceEntry {
                typeface: base,
                origin,
                instance: InstanceState::NotInstanced,
                synthesis: Synthesis::default(),
            });
        };

        // §113: bake the instance's coordinates into a static SFNT before
        // falling back to Skia's live instancing, which makes measurement and
        // painting correct but cannot make embedding correct on its own — see
        // `instance_bake`'s module doc. Tried here, at open time, rather than
        // only where subsetting is enabled, because the wrong-outline bug this
        // fixes is Skia's PDF backend embedding whatever typeface is used to
        // draw, independent of whether this render subsets anything at all.
        let bake_result = match &origin {
            TypefaceOrigin::Embedded {
                id,
                collection_index,
            } => {
                instance_bake::bake_instance(self.embedded_bytes(*id), *collection_index, instance)
            }
            TypefaceOrigin::System { .. } => match base.to_font_data() {
                Some((bytes, collection_index)) => {
                    instance_bake::bake_instance(&bytes, collection_index as u32, instance)
                }
                None => Err(InstanceBakeFailure::Failed(
                    "the font system returned no data for this typeface".to_string(),
                )),
            },
        };
        let baked = bake_result.and_then(|bytes| {
            let data = skia_safe::Data::new_copy(&bytes);
            self.font_mgr.new_from_data(&data, 0).ok_or_else(|| {
                InstanceBakeFailure::Failed(
                    "the font system declined to open the baked font".to_string(),
                )
            })
        });

        match baked {
            Ok(baked_tf) => {
                let typeface_id = TypefaceId::from(&baked_tf);
                Some(TypefaceEntry {
                    typeface: baked_tf,
                    origin: TypefaceOrigin::System { typeface_id },
                    instance: InstanceState::Baked(instance.clone()),
                    synthesis: Synthesis::default(),
                })
            }
            Err(reason) => {
                self.warn_once(
                    &record.canonical_family,
                    format_args!(
                        "'{}' is a variable instance ({reason}); the embedded \
                         font will carry the default location",
                        instance.subfamily
                    ),
                );
                // A named instance is a *different* Skia typeface with its own
                // `unique_id`, so measurement, paint and subset collection all
                // pick it up without further plumbing. If the platform
                // declines to instance it, the un-instanced face is still the
                // right family and weight class — better than falling through
                // to the host default.
                match apply_variation(&base, instance) {
                    Some(instanced) => {
                        let origin = match origin {
                            TypefaceOrigin::System { .. } => TypefaceOrigin::System {
                                typeface_id: TypefaceId::from(&instanced),
                            },
                            embedded => embedded,
                        };
                        Some(TypefaceEntry {
                            typeface: instanced,
                            origin,
                            instance: InstanceState::Unbaked {
                                instance: instance.clone(),
                                reason,
                            },
                            synthesis: Synthesis::default(),
                        })
                    }
                    None => {
                        log::debug!(
                            "[font] '{}': the font system declined to instance '{}'; \
                             using the default location",
                            record.canonical_family,
                            instance.subfamily
                        );
                        Some(TypefaceEntry {
                            typeface: base,
                            origin,
                            instance: InstanceState::NotInstanced,
                            synthesis: Synthesis::default(),
                        })
                    }
                }
            }
        }
    }

    /// Resolve by exact face or family match only, or `None`.
    ///
    /// Unlike [`Self::resolve`], this does not fall back to substitutes or to
    /// the host default — necessary for the emoji pipeline, where substituting a
    /// non-emoji typeface for a missing color emoji font is never correct.
    pub fn resolve_exact(&self, family: &str, style: FontStyle) -> Option<TypefaceEntry> {
        let request = request_from_style(family, style);
        match resolve::resolve(&request, &self.catalog) {
            // The step is the whole test. Every step up to the metadata aliases
            // matched `family` against a name the face actually carries, so a
            // further comparison against the face's *canonical* family would not
            // tighten anything — it would only reject a legitimate match whose
            // canonical name differs from the one asked for, which is exactly
            // the case for an embedded font indexed under the `w:font/@w:name`
            // the document declared.
            resolve::FaceResolution::Selected(selection) if is_exact_step(&selection.step) => {
                self.open(&selection.record)
            }
            _ => None,
        }
    }

    /// Resolve from the host font system only — bypasses embedded fonts.
    ///
    /// Word's font subsetter strips color glyph tables (sbix/CBDT/COLR/SVG) when
    /// embedding emoji fonts, so a docx-embedded "Segoe UI Emoji" carries the
    /// right family name but no color glyphs and must not satisfy emoji
    /// resolution.
    ///
    /// This is a deliberate portability boundary, not an oversight: a document
    /// that ships its own emoji font does not get it back, and the same
    /// document renders with different emoji artwork on macOS, Windows and
    /// Linux. It applies only to the color-emoji path — ordinary embedded text
    /// fonts (§17.8) are honoured normally. See `src/render/emoji/resolve.rs`.
    pub fn resolve_system_only(&self, family: &str, style: FontStyle) -> Option<TypefaceEntry> {
        let request = request_from_style(family, style);
        let faces = self.catalog.family_faces(family)?;
        let effective = EffectiveStyle::resolve(&request, IntrinsicStyle::NEUTRAL);
        let best = faces
            .iter()
            .min_by_key(|(_, record)| {
                (
                    u8::from(record.intrinsic.is_italic() != effective.is_italic()),
                    (record.intrinsic.weight - effective.weight).unsigned_abs(),
                )
            })
            .map(|(_, record)| record.clone())?;
        self.open(&best)
    }

    /// Make `name` resolve to `typeface`, so a face chosen by codepoint
    /// coverage can be carried as a **name** — issue #139.
    ///
    /// Per-glyph fallback (`layout::fragment::fallback`) picks a face by asking
    /// the host which one covers a character, but everything downstream of
    /// layout re-resolves from a family name: `DrawCommand::Text` carries one,
    /// and both the painter and `subset::collect` ask this registry again. A
    /// name is therefore the only handle that reaches both — and this is what
    /// makes it a handle rather than a hope.
    ///
    /// Needed because not every face the host will *return* is a face it will
    /// *find* by name. macOS answers an uncovered codepoint with `.LastResort`,
    /// whose dot-prefixed family plain resolution cannot reach: without a pin
    /// it resolves to some other face, and the codepoint silently goes back to
    /// drawing nothing.
    ///
    /// Pins nothing when the name already resolves to this same typeface, which
    /// is every ordinary family — so a document's own fonts stay authoritative
    /// and this is a no-op in the common case.
    pub fn pin_system_face(&self, name: &str, bold: Toggle, italic: Toggle, typeface: Typeface) {
        let request = FaceRequest::new(name, bold, italic);
        let key = FaceRequestKey::new(&request);
        let wanted = TypefaceId::from(&typeface);

        // `resolve` populates the cache as a side effect, so this both answers
        // "does the name find it anyway?" and leaves the natural entry in place
        // when it does.
        if TypefaceId::from(&self.resolve(&request).typeface) == wanted {
            return;
        }
        self.typefaces
            .borrow_mut()
            .insert(key, TypefaceEntry::system(typeface));
    }

    /// Pre-resolve the four toggle combinations for each family.
    ///
    /// Warms the cache so layout's measurement loop never takes the slow path.
    /// `Off` is not preloaded: it selects the same face as `Absent`, so a
    /// document using it hits the `Absent` work already done and pays only a
    /// second cache entry.
    ///
    /// Because this resolves every combination for every mentioned family
    /// regardless of whether any run actually uses it, most `RUST_LOG=debug`
    /// resolution lines describe a font nothing draws — confirm against the
    /// combinations actually emitted as draw commands before chasing a
    /// decision line here.
    pub fn preload(&self, families: &[String]) {
        for family in families {
            for (bold, italic) in [
                (Toggle::Absent, Toggle::Absent),
                (Toggle::On, Toggle::Absent),
                (Toggle::Absent, Toggle::On),
                (Toggle::On, Toggle::On),
            ] {
                self.resolve(&FaceRequest::new(family, bold, italic));
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
    pub fn cached_entries(&self) -> Vec<(FaceRequestKey, TypefaceEntry)> {
        self.typefaces
            .borrow()
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect()
    }
}

// ─── Helpers ────────────────────────────────────────────────────────────────

/// Rebuild a request from the `FontStyle` the two narrow resolvers still take.
///
/// Their callers — the emoji pipeline — genuinely have a Skia style rather than
/// an OOXML toggle, and always ask for `normal`. Mapping weight back onto a
/// toggle is lossy in principle and exact for every call that occurs.
fn request_from_style(family: &str, style: FontStyle) -> FaceRequest<'_> {
    FaceRequest::new(
        family,
        if *style.weight() >= *Weight::BOLD {
            Toggle::On
        } else {
            Toggle::Absent
        },
        if matches!(style.slant(), Slant::Italic | Slant::Oblique) {
            Toggle::On
        } else {
            Toggle::Absent
        },
    )
}

/// Steps that constitute an *exact* match for [`FontRegistry::resolve_exact`].
///
/// Everything up to and including the metadata aliases is the font asserting its
/// own identity. A parsed name or a substitution is a judgement call, and the
/// emoji pipeline must not act on one: substituting a non-emoji face for a
/// missing color-emoji font produces black blobs, and failing over to the
/// monochrome path is the better answer.
fn is_exact_step(step: &resolve::ResolutionStep) -> bool {
    use resolve::ResolutionStep::*;
    matches!(
        step,
        EmbeddedFace | EmbeddedFamily | SystemFamily | SystemFaceName | SystemMetadataAlias
    )
}

/// Instance `base` at a named location.
fn apply_variation(base: &Typeface, instance: &VariationInstance) -> Option<Typeface> {
    use skia_safe::font_arguments::{variation_position, VariationPosition};
    use skia_safe::{FontArguments, FourByteTag};

    let coordinates: Vec<variation_position::Coordinate> = instance
        .coords
        .iter()
        .map(|coord| variation_position::Coordinate {
            axis: FourByteTag::new(u32::from_be_bytes(coord.axis.0)),
            value: coord.value,
        })
        .collect();
    let arguments = FontArguments::new().set_variation_design_position(VariationPosition {
        coordinates: &coordinates,
    });
    base.clone_with_arguments(&arguments)
}

#[cfg(test)]
pub(crate) mod tests {
    /// Issue #188: the host default is one face per effective style, and
    /// asking the host for it is an `FcFontMatch` plus an `FT_Open_Face` on
    /// fontconfig — so two unknown names with the same toggles fetch it once.
    #[test]
    fn the_host_default_is_fetched_once_per_style() {
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let first = registry.resolve(&FaceRequest::plain("No Such Family 8f3a1c"));
        let second = registry.resolve(&FaceRequest::plain("Another Missing Family 2b7e"));
        assert_eq!(registry.host_default_fetches(), 1);
        assert_eq!(
            TypefaceId::from(&first.typeface),
            TypefaceId::from(&second.typeface),
            "both answers are the same host face"
        );
    }

    use super::*;
    use crate::render::fonts::opentype::NameId;
    use crate::render::fonts::{FaceRequest, Toggle};
    use skia_safe::font_style::Width;

    fn fmgr() -> FontMgr {
        FontMgr::new()
    }

    /// Pull bytes from a guaranteed-available system typeface so tests don't
    /// need bundled font fixtures for the registry-level invariants.
    pub(crate) fn arbitrary_system_font_bytes() -> Vec<u8> {
        let mgr = fmgr();
        let tf = mgr
            .legacy_make_typeface(None::<&str>, FontStyle::normal())
            .expect("system has no default typeface — cannot run test");
        let (bytes, _ttc_index) = tf
            .to_font_data()
            .expect("legacy default typeface lacks raw font bytes — cannot run test");
        bytes
    }

    // ─── Construction and ownership ─────────────────────────────────────

    #[test]
    fn registry_empty_after_construction() {
        let r = FontRegistry::new(fmgr());
        assert_eq!(r.embedded_font_count(), 0);
        assert_eq!(r.cached_typeface_count(), 0);
    }

    /// Two registries on the same thread must not share state — this is the
    /// structural fix for the cross-render poisoning that the previous
    /// `thread_local!`-backed cache caused.
    #[test]
    fn a_fresh_registry_shares_nothing_with_an_earlier_one() {
        let r1 = FontRegistry::new(fmgr());
        let _ = r1.resolve(&FaceRequest::plain("FamilyOne"));
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
    fn resolution_is_cached_and_idempotent() {
        let r = FontRegistry::new(fmgr());
        let request = FaceRequest::plain("DefinitelyNotInstalledXYZ");
        let a = r.resolve(&request);
        let b = r.resolve(&request);
        assert_eq!(
            TypefaceId::from(&a.typeface),
            TypefaceId::from(&b.typeface),
            "second resolution must hit the cache and yield the same typeface"
        );
        assert_eq!(r.cached_typeface_count(), 1);
        assert!(matches!(a.origin, TypefaceOrigin::System { .. }));
    }

    /// The cache key is the *request*, so the three toggle states are distinct
    /// entries even where two of them select the same face. Collapsing them here
    /// would bake that equivalence in where a change to `EffectiveStyle` could
    /// not undo it.
    #[test]
    fn each_toggle_state_gets_its_own_cache_entry() {
        let r = FontRegistry::new(fmgr());
        for bold in [Toggle::Absent, Toggle::Off, Toggle::On] {
            let _ = r.resolve(&FaceRequest::new("ToggleKeyProbe", bold, Toggle::Absent));
        }
        assert_eq!(r.cached_typeface_count(), 3);

        // …and `Off` still selects the same typeface as `Absent`.
        let absent = r.resolve(&FaceRequest::new(
            "ToggleKeyProbe",
            Toggle::Absent,
            Toggle::Absent,
        ));
        let off = r.resolve(&FaceRequest::new(
            "ToggleKeyProbe",
            Toggle::Off,
            Toggle::Absent,
        ));
        assert_eq!(
            TypefaceId::from(&absent.typeface),
            TypefaceId::from(&off.typeface)
        );
    }

    #[test]
    fn the_request_key_folds_case_and_whitespace() {
        let a = FaceRequestKey::new(&FaceRequest::plain("  Proxima   Nova "));
        let b = FaceRequestKey::new(&FaceRequest::plain("proxima nova"));
        assert_eq!(a, b);
        assert_ne!(
            a,
            FaceRequestKey::new(&FaceRequest::new(
                "Proxima Nova",
                Toggle::On,
                Toggle::Absent
            ))
        );
    }

    // ─── Embedded fonts ─────────────────────────────────────────────────

    #[test]
    fn an_embedded_font_outranks_the_host() {
        let bytes = arbitrary_system_font_bytes();

        let without = FontRegistry::new(fmgr());
        let baseline = without.resolve(&FaceRequest::plain("NonexistentFamilyABC"));
        assert!(
            matches!(baseline.origin, TypefaceOrigin::System { .. }),
            "without embedding, an unknown family must fall back to the system path"
        );

        let mut with = FontRegistry::new(fmgr());
        let id = with
            .register_embedded("NonexistentFamilyABC", EmbeddedFontVariant::Regular, bytes)
            .expect("register_embedded should accept a valid system font's bytes");
        let resolved = with.resolve(&FaceRequest::plain("NonexistentFamilyABC"));
        assert_eq!(
            resolved.origin,
            TypefaceOrigin::Embedded {
                id,
                collection_index: 0
            },
            "after registration, resolution must return the embedded origin"
        );
    }

    /// ECMA-376 §17.8.3.3 deobfuscation is enforced upstream by the parser. The
    /// registry-level invariant is that the bytes it stores are byte-identical
    /// to the bytes handed in — no re-encoding, no normalization.
    #[test]
    fn embedded_bytes_are_stored_verbatim() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        let id = r
            .register_embedded(
                "ByteIdentityProbe",
                EmbeddedFontVariant::Regular,
                bytes.clone(),
            )
            .expect("registration should succeed for valid font bytes");
        assert_eq!(r.embedded_bytes(id), bytes.as_slice());
        assert_eq!(
            r.embedded_meta(id),
            ("ByteIdentityProbe", EmbeddedFontVariant::Regular)
        );
    }

    /// issue #116: on Skia's CoreText backend, `FontMgr::new_from_data`
    /// declines any collection index but 0 — `register_embedded` must carve
    /// each face into its own standalone SFNT rather than asking the
    /// platform to open it by index, or the second face of a collection is
    /// silently unreachable. This exercises both the catalog side
    /// (`register_embedded` must catalog face 1, not just face 0) and the
    /// reopen side (`resolve` must actually reopen face 1 for rendering,
    /// through the crate-private `open`/`open_uncached`): a fix to only one
    /// of the two would still fail this test.
    #[test]
    fn every_face_of_an_embedded_collection_registers() {
        let mut r = FontRegistry::new(fmgr());
        let id = r
            .register_embedded(
                "DxCollection",
                EmbeddedFontVariant::Regular,
                fixture_bytes("DxCollection.ttc"),
            )
            .expect("a two-face collection must register");

        for (collection_index, family) in [(0u32, "Dx Collection One"), (1, "Dx Collection Two")] {
            let resolved = r.resolve(&FaceRequest::plain(family));
            assert_eq!(
                resolved.origin,
                TypefaceOrigin::Embedded {
                    id,
                    collection_index
                },
                "face {collection_index} ('{family}') must resolve to its own record"
            );
        }
    }

    /// An embedded font is indexed under the name the *document* declared as
    /// well as whatever its own `name` table says. The two routinely disagree,
    /// and the declared name is the one `w:rFonts` will use.
    #[test]
    fn an_embedded_face_answers_to_its_declared_family() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("DeclaredNameProbe", EmbeddedFontVariant::Regular, bytes)
            .unwrap();

        let faces = r.catalog().embedded_faces();
        assert!(!faces.is_empty());
        assert!(
            faces[0].names.iter().any(|n| n.text == "DeclaredNameProbe"),
            "the declared name must be indexed"
        );
        assert!(
            faces[0].names.iter().any(|n| n.kind.is_asserted()),
            "and so must the font's own, read from its tables"
        );
    }

    #[test]
    fn registering_unusable_bytes_reports_an_error() {
        let mut r = FontRegistry::new(fmgr());
        let err = r
            .register_embedded("Garbage", EmbeddedFontVariant::Regular, vec![0u8; 32])
            .expect_err("bytes that are not a font must be refused");
        assert!(matches!(err, RegisterError::InvalidFontData { .. }));
        assert_eq!(r.embedded_font_count(), 0, "and must not be recorded");
    }

    #[test]
    fn every_resolution_records_an_origin() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("OriginEmbedded", EmbeddedFontVariant::Regular, bytes)
            .unwrap();
        let _ = r.resolve(&FaceRequest::plain("OriginEmbedded"));
        let _ = r.resolve(&FaceRequest::plain("OriginSystemFallbackXYZ"));

        for (_, entry) in r.cached_entries() {
            match entry.origin {
                TypefaceOrigin::Embedded { .. } | TypefaceOrigin::System { .. } => {}
            }
        }
    }

    // ─── One face, one typeface ─────────────────────────────────────────

    /// Two requests that select the same face must yield the *same* Skia
    /// typeface.
    ///
    /// Found by the corpus check on `sample-docx-files-sample1.docx`. The
    /// subsetting pass deduplicates on [`TypefaceId`], so a fresh typeface per
    /// request meant one embedded font was carved and embedded twice — the PDF
    /// grew a second `Ubuntu-Bold` because `<w:b/>` and `<w:b/><w:i w:val="0"/>`
    /// both select it and each got its own `new_from_data`.
    #[test]
    fn one_face_reached_two_ways_is_one_typeface() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("OneFaceProbe", EmbeddedFontVariant::Regular, bytes)
            .unwrap();

        // `Absent` and `Off` select the same face by construction — see the
        // `request` module doc — so these are two requests, one face.
        let absent = r.resolve(&FaceRequest::new(
            "OneFaceProbe",
            Toggle::Absent,
            Toggle::Absent,
        ));
        let off = r.resolve(&FaceRequest::new("OneFaceProbe", Toggle::Off, Toggle::Off));

        assert_eq!(
            TypefaceId::from(&absent.typeface),
            TypefaceId::from(&off.typeface),
            "the same face must not be opened twice — subsetting deduplicates on this id"
        );
        assert_eq!(
            r.cached_typeface_count(),
            2,
            "…while still being two distinct cache entries"
        );
    }

    /// The same guarantee for host faces, which take the other arm of `open`.
    ///
    /// Whether this can fail is platform-dependent — CoreText hands back a
    /// shared typeface for a repeated `new_typeface(i)`, so on macOS the test
    /// passes with the cache removed. It is load-bearing wherever the manager
    /// mints a fresh object, which is exactly the case the embedded arm above
    /// proves is real.
    #[test]
    fn a_host_face_reached_two_ways_is_one_typeface() {
        let r = FontRegistry::new(fmgr());
        let names: Vec<String> = r.font_mgr().family_names().collect();
        let Some(family) = names.first() else { return };

        let a = r.resolve(&FaceRequest::new(family, Toggle::Absent, Toggle::Absent));
        let b = r.resolve(&FaceRequest::new(family, Toggle::Off, Toggle::Off));
        assert_eq!(TypefaceId::from(&a.typeface), TypefaceId::from(&b.typeface));
    }

    // ─── Subsetting contract ────────────────────────────────────────────

    /// Two distinct cache keys can share one underlying typeface (e.g. via
    /// `FONT_SUBSTITUTIONS`). The subsetting pass must update them all, or an
    /// un-updated key would keep painting with the pre-subset typeface.
    #[test]
    fn replace_typeface_by_id_updates_every_key_pointing_at_it() {
        let mut r = FontRegistry::new(fmgr());

        let a = r.resolve(&FaceRequest::plain("UnknownFamilyA"));
        let b = r.resolve(&FaceRequest::plain("UnknownFamilyB"));
        assert_eq!(
            TypefaceId::from(&a.typeface),
            TypefaceId::from(&b.typeface),
            "test setup precondition — both unknowns must resolve to the same default"
        );
        let shared_id = TypefaceId::from(&a.typeface);

        let bytes = arbitrary_system_font_bytes();
        let replacement = r
            .font_mgr()
            .new_from_data(&bytes[..], 0)
            .expect("replacement typeface should construct from valid bytes");
        let replacement_origin = TypefaceOrigin::System {
            typeface_id: TypefaceId::from(&replacement),
        };

        let updated = r.replace_typeface_by_id(shared_id, replacement, replacement_origin);
        assert_eq!(
            updated, 2,
            "both shared-typeface entries must be updated in lockstep"
        );

        let after = r.resolve(&FaceRequest::plain("UnknownFamilyA"));
        assert_ne!(
            TypefaceId::from(&after.typeface),
            shared_id,
            "post-replace resolution must yield the new typeface"
        );
    }

    // ─── The two narrow resolvers ───────────────────────────────────────

    #[test]
    fn resolve_exact_invents_no_fallback() {
        let r = FontRegistry::new(fmgr());
        assert!(
            r.resolve_exact("DefinitelyNotInstalledXYZ", FontStyle::normal())
                .is_none(),
            "unlike resolve(), resolve_exact must not invent a fallback"
        );
    }

    #[test]
    fn resolve_exact_finds_an_embedded_family() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("ExactProbe", EmbeddedFontVariant::Regular, bytes)
            .expect("register_embedded should accept valid font bytes");
        let entry = r
            .resolve_exact("ExactProbe", FontStyle::normal())
            .expect("embedded font must be resolvable via exact match");
        assert!(matches!(entry.origin, TypefaceOrigin::Embedded { .. }));
    }

    /// A substitution is a judgement call, and `resolve_exact` exists so the
    /// emoji pipeline never acts on one — a non-emoji face standing in for a
    /// missing color-emoji font rasterizes as black blobs.
    #[test]
    fn resolve_exact_refuses_a_substitution() {
        let r = FontRegistry::new(fmgr());
        // Calibri is in FONT_SUBSTITUTIONS, so `resolve` may well answer; the
        // exact resolver must not, unless the host genuinely has Calibri.
        let exact = r.resolve_exact("Calibri", FontStyle::normal());
        if let Some(entry) = exact {
            let family = entry.typeface.family_name();
            assert!(
                family.eq_ignore_ascii_case("Calibri"),
                "resolve_exact returned '{family}' for Calibri"
            );
        }
    }

    /// Word's font subsetter strips color glyph tables (sbix/CBDT/COLR/SVG)
    /// when embedding emoji fonts. A docx that embeds e.g. "Segoe UI Emoji"
    /// therefore registers a typeface with the right family name but only
    /// monochrome outlines — rasterizing it produces black blobs instead of the
    /// colored glyph. The emoji pipeline must bypass embedded fonts so
    /// resolution falls through to the host's real color emoji typeface.
    #[test]
    fn resolve_system_only_skips_embedded_fonts() {
        let bytes = arbitrary_system_font_bytes();
        let mut r = FontRegistry::new(fmgr());
        r.register_embedded("SystemOnlyProbe", EmbeddedFontVariant::Regular, bytes)
            .expect("register_embedded should accept valid font bytes");

        let exact = r
            .resolve_exact("SystemOnlyProbe", FontStyle::normal())
            .expect("embedded font must be reachable via resolve_exact");
        assert!(matches!(exact.origin, TypefaceOrigin::Embedded { .. }));

        assert!(
            r.resolve_system_only("SystemOnlyProbe", FontStyle::normal())
                .is_none(),
            "resolve_system_only must skip embedded fonts so emoji resolution \
             can fall through to the host's color emoji typeface"
        );
    }

    // ─── Preload ────────────────────────────────────────────────────────

    #[test]
    fn preload_warms_four_combinations_per_family() {
        let r = FontRegistry::new(fmgr());
        r.preload(&["PreloadProbeA".to_string(), "PreloadProbeB".to_string()]);
        assert_eq!(r.cached_typeface_count(), 8);
    }

    // ─── Face opening ───────────────────────────────────────────────────

    /// A record must reopen to the same face the catalogue described. Reopening
    /// by `(family, style)` instead of by index would re-run the host's matching
    /// and could hand back a different face — the failure acceptance item 7 is
    /// about.
    #[test]
    fn opening_a_manager_face_reproduces_its_family_and_style() {
        let r = FontRegistry::new(fmgr());
        let names: Vec<String> = r.font_mgr().family_names().collect();
        let Some(family) = names.first() else { return };
        let Some(faces) = r.catalog().family_faces(family) else {
            return;
        };
        for (_, record) in faces.iter().take(4) {
            let entry = r.open(record).expect("an indexed face must reopen");
            assert!(
                entry
                    .typeface
                    .family_name()
                    .eq_ignore_ascii_case(&record.canonical_family)
                    || entry.typeface.family_name().eq_ignore_ascii_case(family),
                "reopened '{}' for a record of '{}'",
                entry.typeface.family_name(),
                record.canonical_family
            );
        }
    }

    /// A face record with no instance must not be handed a variation position,
    /// and must come back `NotInstanced` — instanced/baked states are only for
    /// a face record that actually names a design-space location.
    #[test]
    fn an_opened_face_reports_whether_it_is_an_instance() {
        let r = FontRegistry::new(fmgr());
        let names: Vec<String> = r.font_mgr().family_names().collect();
        let Some(family) = names.first() else { return };
        let Some(faces) = r.catalog().family_faces(family) else {
            return;
        };
        let (_, record) = &faces[0];
        let entry = r.open(record).expect("reopen");
        assert_eq!(
            entry.instance,
            InstanceState::NotInstanced,
            "a manager face carries no design-space location"
        );
    }

    /// The embedding half of issue #113, end to end through the real
    /// `FontRegistry::open` path — private, which is why this lives here
    /// rather than in `tests/font_resolution.rs`. A variable instance
    /// registered as an embedded font bakes at open time, and its embedded
    /// bytes differ from the default location's: the same claim
    /// `instance_bake`'s own tests pin one layer lower, now through the exact
    /// path a real render takes.
    ///
    /// Opens the two records directly via `catalog().embedded_faces()`
    /// rather than through `resolve()` by name: `resolve()`'s step 1
    /// (`identifies_one_face`, the only embedded-scope name step) checks only
    /// `Full`/`CompatibleFull`/`PostScript` names — an embedded font's own
    /// `fvar` instance name is not `is_primary_identity()` and is therefore
    /// unreachable by name today, independent of this issue (step 5, the
    /// `Instance`-checking one, only ever runs for the host/`Deep` scope).
    /// This test is about what baking does once a record is opened, not about
    /// that separate, pre-existing reachability gap.
    ///
    /// Uses the committed `DxVariable.ttf` fixture rather than
    /// `arbitrary_system_font_bytes`: this needs a real variable font with
    /// real `gvar` deltas, which no host can be assumed to provide.
    #[test]
    fn a_named_instance_bakes_at_open_time() {
        let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("test-files/fonts/DxVariable.ttf");
        let bytes = std::fs::read(&fixture).unwrap_or_else(|e| {
            panic!("fixture 'DxVariable.ttf' is missing ({e}) — run scripts/make_font_fixtures.py")
        });

        let mut registry = FontRegistry::new(fmgr());
        registry
            .register_embedded(
                "Dx Variable",
                crate::model::EmbeddedFontVariant::Regular,
                bytes,
            )
            .expect("DxVariable.ttf must register");

        let records = registry.catalog().embedded_faces();
        let default_record = records
            .iter()
            .find(|r| r.instance.is_none())
            .expect("the default location has its own record");
        let semibold_record = records
            .iter()
            .find(|r| r.instance.as_ref().map(|i| i.subfamily.as_str()) == Some("SemiBold"))
            .expect("the SemiBold instance has its own record");

        let default_entry = registry.open(default_record).expect("open default");
        let semibold_entry = registry.open(semibold_record).expect("open SemiBold");

        match &semibold_entry.instance {
            InstanceState::Baked(instance) => assert_eq!(instance.subfamily, "SemiBold"),
            other => panic!("expected a baked instance, got {other:?}"),
        }
        assert_eq!(
            default_entry.instance,
            InstanceState::NotInstanced,
            "the default location is not itself an instance to bake"
        );

        let default_bytes = default_entry
            .typeface
            .to_font_data()
            .expect("default typeface must have data")
            .0
            .to_vec();
        let semibold_bytes = semibold_entry
            .typeface
            .to_font_data()
            .expect("baked typeface must have data")
            .0
            .to_vec();
        assert_ne!(
            default_bytes, semibold_bytes,
            "the baked instance's embedded bytes must differ from the default location's"
        );
    }

    fn fixture_bytes(name: &str) -> Vec<u8> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("test-files/fonts")
            .join(name);
        std::fs::read(&path).unwrap_or_else(|e| {
            panic!("fixture '{name}' is missing ({e}) — run scripts/make_font_fixtures.py")
        })
    }

    /// Issue #115, end to end through `FontRegistry::resolve` — the same
    /// path a real render takes, one layer above `Synthesis::compute`'s own
    /// pure tests in `request.rs`. `Dx-Regular.ttf` is a genuinely
    /// single-face family: nothing bolder exists for ranking to find.
    #[test]
    fn a_single_face_family_embolds_on_request() {
        let mut registry = FontRegistry::new(fmgr());
        registry
            .register_embedded(
                "Dx",
                crate::model::EmbeddedFontVariant::Regular,
                fixture_bytes("Dx-Regular.ttf"),
            )
            .expect("Dx-Regular.ttf must register");

        let entry = registry.resolve(&FaceRequest::new("Dx", Toggle::On, Toggle::Absent));
        assert!(entry.synthesis.embolden, "nothing bolder exists to select");
        assert!(!entry.synthesis.oblique);
    }

    /// Acceptance criterion 2, literally: a family with a real Bold face is
    /// unchanged by a bold request.
    #[test]
    fn a_family_with_a_real_bold_face_does_not_embold() {
        let mut registry = FontRegistry::new(fmgr());
        registry
            .register_embedded(
                "DxSans",
                crate::model::EmbeddedFontVariant::Regular,
                fixture_bytes("DxSans-Regular.ttf"),
            )
            .expect("DxSans-Regular.ttf must register");
        registry
            .register_embedded(
                "DxSans",
                crate::model::EmbeddedFontVariant::Bold,
                fixture_bytes("DxSans-Bold.ttf"),
            )
            .expect("DxSans-Bold.ttf must register");

        let entry = registry.resolve(&FaceRequest::new("DxSans", Toggle::On, Toggle::Absent));
        assert!(
            !entry.synthesis.embolden,
            "ranking must have selected the family's real Bold face"
        );
    }

    /// Acceptance criterion 3: a face whose `OS/2` already declares OBLIQUE
    /// (not ITALIC) must not be slanted again.
    #[test]
    fn an_already_oblique_face_does_not_double_slant() {
        let mut registry = FontRegistry::new(fmgr());
        registry
            .register_embedded(
                "Dx Oblique",
                crate::model::EmbeddedFontVariant::Italic,
                fixture_bytes("DxOblique.ttf"),
            )
            .expect("DxOblique.ttf must register");

        let entry = registry.resolve(&FaceRequest::new("Dx Oblique", Toggle::Absent, Toggle::On));
        assert!(
            !entry.synthesis.oblique,
            "the face is already Oblique, not Upright"
        );
    }

    // ─── Helpers ────────────────────────────────────────────────────────

    #[test]
    fn a_font_style_maps_back_onto_toggles_for_the_narrow_resolvers() {
        let bold = request_from_style("X", FontStyle::bold());
        assert_eq!(bold.bold, Toggle::On);
        assert_eq!(bold.italic, Toggle::Absent);

        let italic = request_from_style("X", FontStyle::italic());
        assert_eq!(italic.bold, Toggle::Absent);
        assert_eq!(italic.italic, Toggle::On);

        let plain = request_from_style("X", FontStyle::normal());
        assert_eq!(plain.bold, Toggle::Absent);
        assert_eq!(plain.italic, Toggle::Absent);

        // Semibold counts as bold: the threshold matches what the four-slot
        // embedded-variant model has always used.
        let semi = request_from_style(
            "X",
            FontStyle::new(Weight::SEMI_BOLD, Width::NORMAL, Slant::Upright),
        );
        assert_eq!(semi.bold, Toggle::Absent, "600 is below the BOLD threshold");
        let seven_hundred = request_from_style(
            "X",
            FontStyle::new(Weight::BOLD, Width::NORMAL, Slant::Upright),
        );
        assert_eq!(seven_hundred.bold, Toggle::On);
    }

    /// `resolve_exact` must accept a metadata match but never a guess.
    #[test]
    fn only_evidence_backed_steps_count_as_exact() {
        use resolve::ResolutionStep::*;
        assert!(is_exact_step(&EmbeddedFace));
        assert!(is_exact_step(&EmbeddedFamily));
        assert!(is_exact_step(&SystemFamily));
        assert!(is_exact_step(&SystemFaceName));
        assert!(is_exact_step(&SystemMetadataAlias));

        assert!(!is_exact_step(&ParsedFaceName {
            base_family: "X".into(),
            weight: 300
        }));
        assert!(!is_exact_step(&Substitution {
            via: "Calibri",
            substitute: "Carlito"
        }));
    }

    /// A missing font is worth one warning, not one per style variant.
    #[test]
    fn a_missing_family_is_reported_once() {
        let r = FontRegistry::new(fmgr());
        r.preload(&["NoSuchFamily8f3a1c".to_string()]);
        assert_eq!(
            r.warned.borrow().len(),
            1,
            "four preloaded combinations must not produce four warnings"
        );
        // A different family gets its own.
        let _ = r.resolve(&FaceRequest::plain("AnotherMissing8f3a1c"));
        assert_eq!(r.warned.borrow().len(), 2);
    }

    /// `NameId` round-trips are load-bearing for the catalogue's index, so pin
    /// the two the registry itself relies on.
    #[test]
    fn the_indexed_name_ids_are_the_ones_documents_use() {
        assert_eq!(NameId::from_raw(4), NameId::Full);
        assert_eq!(NameId::from_raw(6), NameId::PostScript);
    }
}
