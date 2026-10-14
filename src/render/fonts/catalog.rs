//! Everything that can be selected, in one place.
//!
//! The catalogue turns two very different things — a host font manager and a
//! handful of DOCX-embedded font files — into one uniform list of
//! [`FaceRecord`]s, so the resolver matches names against records and never has
//! to know where a face came from.
//!
//! # Three tiers, and what each costs
//!
//! Measured on a host with 210 families / 677 faces, warm, release build:
//!
//! | Operation | Cost |
//! |---|---|
//! | `FontMgr::family_names()` | 0.04 ms |
//! | `FontMgr::match_family()` × 210 | **28 ms** |
//! | `FontStyleSet::style(i)` × 677 | 0.4 ms |
//! | `FontStyleSet::new_typeface(i)` × 677, style sets retained | 7.6 ms |
//! | `copy_table_data` × 677 for `name`/`OS/2`/`fvar`/`STAT` | 1.7 ms |
//! | parsing 32 526 `name` records | 4.6 ms |
//!
//! One line dominates: `match_family` is the whole cost of enumeration, and
//! everything else is nearly free. That is what the tiering is shaped around,
//! and it is why the previous `FaceAliasIndex` cost 35 ms warm — it called
//! `match_family` for every family on the host to read one PostScript name per
//! face.
//!
//! - **Tier 0**, at construction: `family_names()` only. Effectively free, and
//!   enough to answer "does the host have a family by exactly this name?".
//! - **Tier 1**, per family on first touch: `match_family` for *that* family,
//!   then its faces from the style set. A document naming twenty families pays
//!   twenty of the 210 `match_family` calls — about 2.7 ms, not 28.
//! - **Tier 2** ([`lookup_deep`](FaceCatalog::lookup_deep)), once per render at
//!   most: every face instantiated and its own `name`, `OS/2`, `fvar` and
//!   `STAT` tables read. Reached only by a request that has already missed
//!   every cheaper step.
//!
//! Style sets are retained from tier 1, so tier 2 does not pay `match_family` a
//! second time — the difference between 7.6 ms and 37 ms for the same 677
//! `new_typeface` calls.
//!
//! End to end on the same host, against the 35 ms the previous unconditional
//! index cost every render:
//!
//! | Scenario | Cost |
//! |---|---|
//! | Construction | 0.05 ms |
//! | + twenty families touched (the ordinary document) | 2.6 ms |
//! | + deep tier reached (a document naming a font the host lacks) | 44 ms |
//!
//! So the common case is roughly thirteen times cheaper than before, and the
//! case that pays for the whole metadata model costs about 9 ms more.
//!
//! Those numbers are CoreText's, where instantiating a face is a descriptor
//! and reading a table is a call into the font server. They do not transfer
//! to fontconfig/FreeType (issue #188): there every `new_typeface` is an
//! `FT_Open_Face` with an mmap of the file, `enumerate` walks every font on
//! the host per family and opens each one that matches, and the manager
//! retains up to 1024 of the opened faces — so the deep tier costs seconds and
//! hundreds of megabytes on a Linux desktop, proportional to the host's font
//! count rather than to the document. That is why the resolver answers a
//! curated substitution *before* reaching this tier (`resolve.rs`, step 3b),
//! and why `family_faces` and `host_default` ask the host once per name.
//!
//! # Ownership
//!
//! Per render, owned by [`FontRegistry`](super::FontRegistry), never process
//! global. That is acceptance item 8 of issue #85, and it is also the rule that
//! made `FontRegistry` itself per-render: the subsetting pass mutates typefaces
//! in place, so anything cached across renders can leak a font subsetted for the
//! *previous* document. The catalogue holds no typefaces and no font bytes —
//! only names, weights and axis values — so it could not leak glyphs even if it
//! were shared. It is kept per-render anyway, because a catalogue built from one
//! `FontMgr` is simply wrong for another, and `render_with_font_mgr` accepts a
//! caller-supplied manager.
//!
//! The catalogue also holds no embedded font *bytes*: the registry already owns
//! those, and a second copy of a 14 MB embedded font would be a real cost.
//! Registration hands in a typeface built from those bytes, the catalogue reads
//! what it needs, and the registry reopens the face later from its own copy.
//!
//! # Why lookups return owned records
//!
//! Tier 1 is memoized behind a `RefCell`, so a borrowed `&FaceRecord` could not
//! outlive the lookup that produced it. Records are small — a handful of
//! `String`s — and the registry caches whole resolutions, so a clone happens
//! once per distinct request rather than once per word of text.

use std::cell::{Cell, OnceCell, RefCell};
use std::collections::HashMap;

use skia_safe::{FontMgr, FontStyle, FontStyleSet, Typeface};

use super::face::{
    fold_name, FaceIdentity, FaceName, FaceRecord, IntrinsicStyle, NameKind, VariationInstance,
};
use super::opentype::fvar::{Axis, VariationCoord, Variations};
use super::opentype::name::NameRecords;
use super::opentype::{fvar, name, os2, stat, table_tag, NameId};
use super::EmbeddedFontId;

/// Where a face lives, and therefore how much is known about it.
///
/// The distinction is not bookkeeping: a [`Manager`](Self::Manager) face is
/// known only by what the host font manager volunteered, so its names are
/// composed rather than asserted and it is not eligible for the metadata steps
/// of the resolution chain. A [`Deep`](Self::Deep) face has had its own tables
/// read.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FaceRef {
    /// A face in a DOCX-embedded font. Always deep — the bytes were in hand.
    Embedded(usize),
    /// A face known from the manager's own description: `family` indexes the
    /// enumerated family list, `face` the style set within it.
    Manager { family: usize, face: usize },
    /// A face whose own OpenType tables have been read.
    Deep(usize),
}

/// The result of asking the catalogue for a name.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FaceLookup {
    Found(FaceRef),
    /// Several genuinely different faces answer to this name. Acceptance
    /// criterion: *ambiguous aliases fail safely rather than choosing insertion
    /// order*, so this is a distinct outcome the resolver reports rather than an
    /// arbitrary pick.
    Ambiguous {
        candidates: usize,
    },
    Missing,
}

/// Which half of the catalogue a lookup searches.
///
/// Separate from [`FaceRef`] because it is a *question*, not an answer: asking
/// the deep scope is what builds tier 2, so the choice is a cost decision the
/// resolver makes deliberately at each step.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Scope {
    /// Faces from fonts the DOCX embedded.
    Embedded,
    /// Faces from the host font system, with their own tables read.
    Deep,
}

/// One host family: its style set, retained, plus its faces once enumerated.
struct FamilySlot {
    name: String,
    /// `None` until the family is first touched — retaining it is what lets
    /// tier 2 skip a second `match_family`.
    style_set: Option<FontStyleSet>,
    faces: Option<Vec<FaceRecord>>,
}

/// Name → candidate record indices. Candidates are filtered by name *kind* at
/// lookup rather than at insertion, so one index serves every resolution step.
#[derive(Default, Debug)]
struct NameIndex {
    by_name: HashMap<String, Vec<usize>>,
}

impl NameIndex {
    fn insert(&mut self, record: &FaceRecord, index: usize) {
        for name in &record.names {
            let key = fold_name(&name.text);
            if key.is_empty() {
                continue;
            }
            let slot = self.by_name.entry(key).or_default();
            if !slot.contains(&index) {
                slot.push(index);
            }
        }
    }

    fn candidates(&self, folded: &str) -> &[usize] {
        self.by_name.get(folded).map_or(&[], Vec::as_slice)
    }
}

/// Tier 2: records enriched from each face's own tables, plus their name index.
#[derive(Default, Debug)]
struct DeepTier {
    records: Vec<FaceRecord>,
    index: NameIndex,
}

pub struct FaceCatalog {
    font_mgr: FontMgr,
    /// Folded family name → index into `families`.
    family_index: HashMap<String, usize>,
    families: RefCell<Vec<FamilySlot>>,
    embedded: Vec<FaceRecord>,
    embedded_index: NameIndex,
    deep: OnceCell<DeepTier>,
    /// Folded name → the slot `adopt_unenumerated` answered with, both ways.
    /// One host probe per name per render; see `family_faces`.
    adopted: RefCell<HashMap<String, Option<usize>>>,
    /// How many times the manager was asked about a name it did not enumerate
    /// — the observable the memo's test reads.
    adoption_probes: Cell<usize>,
}

impl FaceCatalog {
    /// Tier 0: enumerate family *names* and nothing else.
    pub fn build(font_mgr: FontMgr) -> Self {
        let mut families = Vec::new();
        let mut family_index = HashMap::new();
        for name in font_mgr.family_names() {
            // A host may expose the same family name twice (per-user and
            // system-wide installs). Keep the first; they enumerate the same
            // faces, and a second slot would only ever be reachable by index.
            family_index
                .entry(fold_name(&name))
                .or_insert(families.len());
            families.push(FamilySlot {
                name,
                style_set: None,
                faces: None,
            });
        }
        log::debug!("[font] catalogue: {} host families", families.len());

        Self {
            font_mgr,
            family_index,
            families: RefCell::new(families),
            embedded: Vec::new(),
            embedded_index: NameIndex::default(),
            deep: OnceCell::new(),
            adopted: RefCell::new(HashMap::new()),
            adoption_probes: Cell::new(0),
        }
    }

    /// [`Self::from_faces`], but with tier 2 left to be built on first use —
    /// so a test can observe, through [`Self::deep_is_built`], whether a
    /// request reached the metadata tier at all. `from_faces` pre-builds it
    /// and so cannot answer that question.
    pub fn from_faces_without_deep_tier(font_mgr: FontMgr, faces: &[(String, Typeface)]) -> Self {
        let mut catalog = Self::from_faces(font_mgr, faces);
        catalog.deep = OnceCell::new();
        catalog
    }

    /// How many times the manager has been asked about a name it did not
    /// enumerate.
    pub fn adoption_probes(&self) -> usize {
        self.adoption_probes.get()
    }

    /// Build a catalogue over an explicit set of faces instead of the host's.
    ///
    /// Every tier is populated up front from `faces`, so nothing here consults
    /// the font manager for names or styles — `font_mgr` is retained only
    /// because opening a face later needs one.
    ///
    /// This is how issue #85's cross-platform criterion is met: the resolver's
    /// tests build a catalogue from the fonts committed under
    /// `test-files/fonts/` and get the same answers on macOS, Linux and
    /// Windows, where a catalogue over the host's fonts would get whatever that
    /// host happens to install.
    pub fn from_faces(font_mgr: FontMgr, faces: &[(String, Typeface)]) -> Self {
        let mut catalog = Self::build(FontMgr::empty());
        catalog.font_mgr = font_mgr;

        let mut families: Vec<FamilySlot> = Vec::new();
        let mut family_index: HashMap<String, usize> = HashMap::new();
        let mut deep = DeepTier::default();

        for (family, typeface) in faces {
            let slot = *family_index.entry(fold_name(family)).or_insert_with(|| {
                families.push(FamilySlot {
                    name: family.clone(),
                    style_set: None,
                    faces: Some(Vec::new()),
                });
                families.len() - 1
            });
            let face_index = families[slot].faces.as_ref().map_or(0, |faces| faces.len());
            let identity = FaceIdentity::System {
                family: family.clone(),
                face_index,
            };

            // The manager-tier record has to be built from the typeface rather
            // than from a style set, because there is no style set. Its *names*
            // are still manager-kind: they are composed here, not asserted by
            // the font, and the resolution steps must treat them accordingly.
            let mut manager = manager_face_record(
                family,
                face_index,
                typeface.font_style(),
                typeface_style_name(typeface),
            );
            manager.identity = identity.clone();

            let record = deep_face_record(typeface, identity, family, Some(&manager));
            let instances = instance_records(typeface, &record);

            if let Some(faces) = families[slot].faces.as_mut() {
                faces.push(manager);
            }
            deep.index.insert(&record, deep.records.len());
            deep.records.push(record);
            for instance in instances {
                deep.index.insert(&instance, deep.records.len());
                deep.records.push(instance);
            }
        }

        catalog.families = RefCell::new(families);
        catalog.family_index = family_index;
        let _ = catalog.deep.set(deep);
        catalog
    }

    /// Record one face of a DOCX-embedded font.
    ///
    /// `typeface` is built from the registry's own copy of the bytes and is
    /// borrowed only for the duration of the call — see the module doc on why
    /// the catalogue does not keep font bytes.
    ///
    /// `declared_family` is the `w:font/@w:name` the document wrote. It is kept
    /// alongside whatever the font's own `name` table says, because the two
    /// routinely disagree and the declared name is the one the document will ask
    /// for.
    pub fn add_embedded_face(
        &mut self,
        font: EmbeddedFontId,
        collection_index: u32,
        declared_family: &str,
        typeface: &Typeface,
    ) {
        let mut record = deep_face_record(
            typeface,
            FaceIdentity::Embedded {
                font,
                collection_index,
            },
            declared_family,
            None,
        );
        push_unique_name(
            &mut record.names,
            FaceName::composed(NameKind::ManagerFamily, declared_family),
        );

        let instances = instance_records(typeface, &record);
        self.embedded_index.insert(&record, self.embedded.len());
        self.embedded.push(record);
        for instance in instances {
            self.embedded_index.insert(&instance, self.embedded.len());
            self.embedded.push(instance);
        }
    }

    pub fn font_mgr(&self) -> &FontMgr {
        &self.font_mgr
    }

    pub fn embedded_faces(&self) -> &[FaceRecord] {
        &self.embedded
    }

    pub fn family_count(&self) -> usize {
        self.families.borrow().len()
    }

    /// Whether tier 2 has been built. Exposed so a test can assert that a
    /// request resolved *without* paying for it — that guarantee is the whole
    /// point of the tiering, and nothing else can observe it.
    pub fn deep_is_built(&self) -> bool {
        self.deep.get().is_some()
    }

    /// Whether a family's faces have been enumerated yet, for the same reason.
    pub fn family_is_enumerated(&self, name: &str) -> bool {
        self.family_index
            .get(&fold_name(name))
            .and_then(|&i| self.families.borrow().get(i).map(|f| f.faces.is_some()))
            .unwrap_or(false)
    }

    /// The face a [`FaceRef`] names, cloned — see the module doc.
    pub fn face(&self, face: FaceRef) -> Option<FaceRecord> {
        match face {
            FaceRef::Embedded(i) => self.embedded.get(i).cloned(),
            FaceRef::Manager { family, face } => {
                self.enumerate(family);
                self.families
                    .borrow()
                    .get(family)?
                    .faces
                    .as_ref()?
                    .get(face)
                    .cloned()
            }
            FaceRef::Deep(i) => self.deep.get().and_then(|d| d.records.get(i)).cloned(),
        }
    }

    /// Every face of a host family named exactly `name` (case- and
    /// whitespace-folded), enumerating that family if this is its first touch.
    ///
    /// `None` means no such family — after asking the manager directly, not
    /// merely after failing to find it in the enumeration; see
    /// `adopt_unenumerated` below.
    pub fn family_faces(&self, name: &str) -> Option<Vec<(FaceRef, FaceRecord)>> {
        let folded = fold_name(name);
        let family = match self.family_index.get(&folded) {
            Some(&family) => family,
            // Asked of the manager once per name per render, whichever way it
            // answered. `preload` asks for four style variants of every family,
            // and before this memo each repeat was another host probe — on
            // fontconfig an `FcFontMatch` plus an `FT_Open_Face` — and, for a
            // name the manager *does* match, another duplicate slot in
            // `families` that the deep tier then read as well (issue #188).
            None => {
                let adopted = self.adopted.borrow().get(&folded).copied();
                match adopted {
                    Some(slot) => slot?,
                    None => {
                        let slot = self.adopt_unenumerated(name);
                        self.adopted.borrow_mut().insert(folded, slot);
                        slot?
                    }
                }
            }
        };
        self.enumerate(family);
        let families = self.families.borrow();
        let faces = families.get(family)?.faces.as_ref()?;
        Some(
            faces
                .iter()
                .enumerate()
                .map(|(face, record)| (FaceRef::Manager { family, face }, record.clone()))
                .collect(),
        )
    }

    /// Take on a family the manager will match but did not enumerate, and
    /// return its new slot index.
    ///
    /// `family_names()` is not guaranteed to list every name `match_family_style`
    /// answers to. Before this existed, resolution went through
    /// `match_family_style` directly and filtered the answer by family name; that
    /// path is preserved here rather than dropped, because dropping it would
    /// silently stop resolving any family a host matches but does not advertise.
    ///
    /// The manager is *permitted to substitute* rather than decline — fontconfig
    /// routinely does, CoreText does not — so the answer is only accepted when
    /// the returned face genuinely carries the requested family name. Without
    /// that check every miss would be swallowed as a hit on Linux and the whole
    /// rest of the chain would be unreachable.
    fn adopt_unenumerated(&self, name: &str) -> Option<usize> {
        self.adoption_probes.set(self.adoption_probes.get() + 1);
        let matched = self
            .font_mgr
            .match_family_style(name, FontStyle::normal())
            .filter(|tf| carries_family(tf, name))?;
        log::debug!(
            "[font] '{name}' is matchable as '{}' but was not enumerated; adopting it",
            matched.family_name()
        );

        let mut families = self.families.borrow_mut();
        let index = families.len();
        families.push(FamilySlot {
            name: name.to_owned(),
            style_set: None,
            faces: None,
        });
        drop(families);
        // Not inserted into `family_index`: that map is built once from the
        // manager's own enumeration and is shared immutably. `family_faces`
        // remembers the answer instead, so the probe runs once per name.
        Some(index)
    }

    /// Faces answering to `name` under an accepted name kind, in one scope.
    ///
    /// Returns *all* candidates rather than insisting on one: whether several
    /// matches are an ambiguity or a family to rank within depends on the
    /// resolution step, and only the resolver knows which it is asking.
    pub fn candidates(
        &self,
        scope: Scope,
        name: &str,
        accept: impl Fn(NameKind) -> bool + Copy,
    ) -> Vec<(FaceRef, FaceRecord)> {
        let folded = fold_name(name);
        if folded.is_empty() {
            return Vec::new();
        }
        let (index, records, to_ref): (_, _, fn(usize) -> FaceRef) = match scope {
            Scope::Embedded => (
                &self.embedded_index,
                &self.embedded,
                FaceRef::Embedded as fn(usize) -> FaceRef,
            ),
            // Reaching the deep scope at all is what builds tier 2.
            Scope::Deep => {
                let deep = self.deep();
                (
                    &deep.index,
                    &deep.records,
                    FaceRef::Deep as fn(usize) -> FaceRef,
                )
            }
        };
        index
            .candidates(&folded)
            .iter()
            .filter_map(|&i| {
                let record = records.get(i)?;
                record
                    .answers_to(&folded, accept)
                    .then(|| (to_ref(i), record.clone()))
            })
            .collect()
    }

    /// Like [`candidates`](Self::candidates), but for a step that requires the
    /// name to identify exactly one face.
    pub fn lookup(
        &self,
        scope: Scope,
        name: &str,
        accept: impl Fn(NameKind) -> bool + Copy,
    ) -> FaceLookup {
        pick_unique(&self.candidates(scope, name, accept))
    }

    /// Look a system face up by name in tier 2, **building tier 2 if it has not
    /// been built yet**. Every caller is a resolution step that has already
    /// missed the cheap ones.
    pub fn lookup_deep(&self, name: &str, accept: impl Fn(NameKind) -> bool + Copy) -> FaceLookup {
        self.lookup(Scope::Deep, name, accept)
    }

    /// Tier 1: enumerate one family's faces, once.
    fn enumerate(&self, family: usize) {
        let mut families = self.families.borrow_mut();
        let Some(slot) = families.get_mut(family) else {
            return;
        };
        if slot.faces.is_some() {
            return;
        }
        let mut style_set = self.font_mgr.match_family(&slot.name);
        let count = style_set.count();
        let mut faces = Vec::with_capacity(count);
        for face_index in 0..count {
            let (style, style_name) = style_set.style(face_index);
            faces.push(manager_face_record(
                &slot.name, face_index, style, style_name,
            ));
        }
        slot.faces = Some(faces);
        slot.style_set = Some(style_set);
    }

    /// Tier 2, built on first use.
    fn deep(&self) -> &DeepTier {
        self.deep.get_or_init(|| {
            let started = std::time::Instant::now();
            let mut tier = DeepTier::default();
            // Re-read the length each round: `adopt_unenumerated` can append a
            // family, and a snapshot taken up front would leave it out of the
            // index for the rest of the render.
            let mut family = 0;
            while family < self.families.borrow().len() {
                // Enumerating first means the style set is retained, so the
                // `new_typeface` calls below cost 7.6 ms across the host rather
                // than 37 — the difference is a second `match_family` per
                // family. See the module doc's cost table.
                self.enumerate(family);
                for record in self.deep_records_for(family) {
                    tier.index.insert(&record, tier.records.len());
                    tier.records.push(record);
                }
                family += 1;
            }
            log::debug!(
                "[font] catalogue deep tier: {} faces, {} distinct names, {:?}",
                tier.records.len(),
                tier.index.by_name.len(),
                started.elapsed()
            );
            tier
        })
    }

    /// Deep records for one already-enumerated family: each face, followed
    /// immediately by its own named instances.
    fn deep_records_for(&self, family: usize) -> Vec<FaceRecord> {
        let mut families = self.families.borrow_mut();
        let Some(slot) = families.get_mut(family) else {
            return Vec::new();
        };
        let family_name = slot.name.clone();
        let Some(manager_faces) = slot.faces.clone() else {
            return Vec::new();
        };
        let Some(style_set) = slot.style_set.as_mut() else {
            return Vec::new();
        };

        let mut out = Vec::with_capacity(manager_faces.len());
        for (face_index, manager) in manager_faces.iter().enumerate() {
            let Some(typeface) = style_set.new_typeface(face_index) else {
                continue;
            };
            let record = deep_face_record(
                &typeface,
                FaceIdentity::System {
                    family: family_name.clone(),
                    face_index,
                },
                &family_name,
                Some(manager),
            );
            let instances = instance_records(&typeface, &record);
            out.push(record);
            out.extend(instances);
        }
        out
    }
}

/// Does `typeface` actually carry the family that was asked for?
///
/// `FontMgr::match_family_style` is permitted to answer with a *substitute*
/// rather than declining, and whether it does is platform-dependent: fontconfig
/// routinely falls back, while CoreText returns `None` for a family it does not
/// have (measured — every unknown name tried on macOS, including the CSS
/// generics `sans-serif`/`serif`/`monospace`, yields `None`). So this guard is
/// inert on macOS and load-bearing on Linux, where without it every miss would
/// be swallowed as a hit and the rest of the resolution chain would be
/// unreachable.
///
/// It is the last place a manager answer is taken on trust at all: every other
/// face in the catalogue is opened *positionally*, by family plus face index, so
/// there is nothing to validate.
///
/// Split out from [`FaceCatalog::adopt_unenumerated`] so it can be tested
/// directly — on a host whose matcher never substitutes there is no way to reach
/// the rejection path through the caller, and a test written that way would pass
/// without the guard existing.
pub(crate) fn carries_family(typeface: &Typeface, requested: &str) -> bool {
    typeface.family_name().eq_ignore_ascii_case(requested)
}

/// Narrow candidates to one face, or report that the name is ambiguous.
///
/// The same font installed twice, or reachable under two family names, is not
/// an ambiguity — every candidate would draw the same glyphs. Only genuinely
/// different faces are, and those are reported rather than silently resolved to
/// whichever happened to be indexed first.
pub(crate) fn pick_unique(candidates: &[(FaceRef, FaceRecord)]) -> FaceLookup {
    let Some((first_ref, first)) = candidates.first() else {
        return FaceLookup::Missing;
    };
    let distinct = candidates[1..]
        .iter()
        .filter(|(_, other)| !same_face(first, other))
        .count();
    if distinct == 0 {
        FaceLookup::Found(*first_ref)
    } else {
        FaceLookup::Ambiguous {
            candidates: distinct + 1,
        }
    }
}

/// Whether two records would draw the same thing.
///
/// Compared on what a reader would see, not on identity: two installs of the
/// same font file have different [`FaceIdentity`]s but are interchangeable, and
/// treating them as an ambiguity would make a perfectly resolvable name fail.
fn same_face(a: &FaceRecord, b: &FaceRecord) -> bool {
    a.canonical_family == b.canonical_family
        && a.intrinsic == b.intrinsic
        && instance_coords(a) == instance_coords(b)
}

fn instance_coords(record: &FaceRecord) -> Option<&[VariationCoord]> {
    record.instance.as_ref().map(|i| i.coords.as_slice())
}

/// The style name a typeface reports for itself, for a catalogue built without
/// a style set to ask.
///
/// `name` id 17 first: for a face like `"Dx Sans SemiBold"` the *typographic*
/// subfamily is `"SemiBold"` while the legacy one is `"Regular"`, and the
/// former is what a manager would have reported.
fn typeface_style_name(typeface: &Typeface) -> Option<String> {
    let bytes = read_table(typeface, b"name")?;
    let names = name::read(&bytes).ok()?;
    names
        .preferred(NameId::TypographicSubfamily)
        .or_else(|| names.preferred(NameId::Subfamily))
        .map(str::to_owned)
}

/// A tier-1 record: everything derivable from the manager without opening the
/// font.
fn manager_face_record(
    family: &str,
    face_index: usize,
    style: FontStyle,
    style_name: Option<String>,
) -> FaceRecord {
    let mut names = vec![FaceName::composed(NameKind::ManagerFamily, family)];
    if let Some(style_name) = style_name.filter(|s| !s.trim().is_empty()) {
        push_unique_name(
            &mut names,
            FaceName::composed(NameKind::ManagerFace, format!("{family} {style_name}")),
        );
    }
    FaceRecord {
        identity: FaceIdentity::System {
            family: family.to_owned(),
            face_index,
        },
        canonical_family: family.to_owned(),
        typographic_family: None,
        intrinsic: IntrinsicStyle::from_font_style(style),
        names,
        instance: None,
    }
}

/// A tier-2 record: the face's own tables, with the manager's view kept as a
/// fallback and as extra (composed) names.
fn deep_face_record(
    typeface: &Typeface,
    identity: FaceIdentity,
    fallback_family: &str,
    manager: Option<&FaceRecord>,
) -> FaceRecord {
    let names_table = read_table(typeface, b"name")
        .and_then(|bytes| log_malformed(name::read(&bytes), "name", fallback_family))
        .unwrap_or_default();

    let intrinsic = read_table(typeface, b"OS/2")
        .and_then(|bytes| log_malformed(os2::read(&bytes), "OS/2", fallback_family))
        .map(|m| IntrinsicStyle {
            weight: m.weight,
            width: m.width,
            slant: m.slant,
            bold_slot: m.bold_bit,
            regular_slot: m.regular_bit,
        })
        // No readable `OS/2` — fall back to what the platform says. Every
        // manager reports a `FontStyle` for every face, so this is never a dead
        // end, only a coarser answer.
        .unwrap_or_else(|| IntrinsicStyle::from_font_style(typeface.font_style()));

    let canonical_family = names_table
        .preferred(NameId::Family)
        .map(str::to_owned)
        .or_else(|| manager.map(|m| m.canonical_family.clone()))
        .unwrap_or_else(|| fallback_family.to_owned());
    let typographic_family = names_table
        .preferred(NameId::TypographicFamily)
        .map(str::to_owned);

    let mut names = Vec::new();
    for record in names_table.iter() {
        if !INDEXED_NAME_IDS.contains(&record.id) {
            continue;
        }
        push_unique_name(
            &mut names,
            FaceName {
                kind: NameKind::Table(record.id),
                text: record.text.clone(),
                localized: !record.language.is_english(),
            },
        );
    }
    if let Some(manager) = manager {
        for name in &manager.names {
            push_unique_name(&mut names, name.clone());
        }
    }
    push_unique_name(
        &mut names,
        FaceName::composed(NameKind::ManagerFamily, &canonical_family),
    );

    // §STAT: compose "family + style word" names the font itself sanctions,
    // e.g. "Fixture Sans Condensed SemiBold". This is the same *shape* of name
    // the old `strip_weight_suffix` guessed at, except every word comes from the
    // font's own axis-value records instead of a hard-coded English table.
    if let Some(stat) = read_table(typeface, b"STAT")
        .and_then(|bytes| log_malformed(stat::read(&bytes), "STAT", &canonical_family))
    {
        let family = typographic_family.as_deref().unwrap_or(&canonical_family);
        for (composed, coords) in composed_style_names(family, &stat, &names_table) {
            // Only the combinations that describe *this* face. A variable
            // font's other locations become their own records — see
            // `instance_records` — rather than extra names on the default one.
            if coords_describe(&coords, &intrinsic) {
                push_unique_name(
                    &mut names,
                    FaceName::composed(NameKind::ComposedStyle, composed),
                );
            }
        }
    }

    FaceRecord {
        identity,
        canonical_family,
        typographic_family,
        intrinsic,
        names,
        instance: None,
    }
}

/// The `name` ids worth indexing. Everything else in a `name` table is
/// copyright, licence text, designer URLs and sample strings — none of which a
/// document would ever put in `w:rFonts`, and all of which would add thousands
/// of index entries per host.
const INDEXED_NAME_IDS: &[NameId] = &[
    NameId::Family,
    NameId::UniqueId,
    NameId::Full,
    NameId::PostScript,
    NameId::TypographicFamily,
    NameId::CompatibleFull,
    NameId::WwsFamily,
];

/// One `STAT` axis-value name and the location it denotes.
///
/// A named type because the composition below builds a cross-product of these
/// and the nested tuple is unreadable inline.
type StyleWord<'a> = (&'a str, Vec<VariationCoord>);

/// A composed style name and the location the whole combination denotes.
type ComposedName = (String, Vec<VariationCoord>);

/// Cap on `STAT`-composed names per face.
///
/// The cross-product of axis values is combinatorial — a font with five axes and
/// six named values each would generate 7 776 names. Every real family is far
/// below this; the cap exists so a pathological font costs a truncated index
/// rather than a hung render, and truncation is logged rather than silent.
const MAX_COMPOSED_NAMES: usize = 128;

/// Compose `"family + style"` names from `STAT` axis values, in the font's own
/// axis order.
///
/// Elided values contribute nothing (that is how `"Regular"` and `"Normal"` stay
/// out of the name), so the cross-product runs over the nameable values of each
/// axis plus the option of omitting that axis entirely.
fn composed_style_names(
    family: &str,
    stat: &stat::StatAxisValues,
    names_table: &NameRecords,
) -> Vec<ComposedName> {
    // Group nameable values by axis, in declared axis order.
    let mut per_axis: Vec<(u16, Vec<StyleWord<'_>>)> = Vec::new();
    for axis in &stat.axes {
        let mut words: Vec<StyleWord<'_>> = Vec::new();
        for value in stat.nameable() {
            if !value.coords.iter().any(|c| c.axis == axis.tag) {
                continue;
            }
            let Some(text) = names_table.preferred(NameId::from_raw(value.name_id)) else {
                continue;
            };
            if !words.iter().any(|(existing, _)| *existing == text) {
                words.push((text, value.coords.clone()));
            }
        }
        if !words.is_empty() {
            per_axis.push((axis.ordering, words));
        }
    }
    per_axis.sort_by_key(|(ordering, _)| *ordering);

    let mut out: Vec<ComposedName> = vec![(family.to_owned(), Vec::new())];
    for (_, words) in &per_axis {
        let mut next = Vec::with_capacity(out.len() * (words.len() + 1));
        for (prefix, coords) in &out {
            // Omitting the axis is a legitimate combination: a font may name
            // both "Fixture Sans Bold" and "Fixture Sans Condensed Bold".
            next.push((prefix.clone(), coords.clone()));
            for (word, word_coords) in words {
                let mut combined = coords.clone();
                combined.extend(word_coords.iter().copied());
                next.push((format!("{prefix} {word}"), combined));
            }
        }
        if next.len() > MAX_COMPOSED_NAMES {
            log::debug!(
                "[font] '{family}': STAT cross-product exceeds {MAX_COMPOSED_NAMES}; \
                 composed names truncated"
            );
            next.truncate(MAX_COMPOSED_NAMES);
        }
        out = next;
    }
    // The bare family is not a *style* name; it is already indexed as such.
    out.retain(|(name, _)| name != family);
    out
}

/// Whether a `STAT` combination describes the face that carries it.
///
/// This is what stops a variable font's *default* record from answering to
/// `"Dx Variable SemiBold"`. A composed name names a point in the design space;
/// attaching it to whichever record happened to be built first made one face
/// answer to every style its family can take, and made the real named instance
/// look like an ambiguity.
///
/// An axis the combination does not mention is unconstrained — a font may name
/// `"Condensed"` without saying anything about weight.
fn coords_describe(coords: &[VariationCoord], intrinsic: &IntrinsicStyle) -> bool {
    use skia_safe::font_style::Slant;
    coords.iter().all(|coord| match coord.axis {
        fvar::AxisTag::WEIGHT => (coord.value.round() as i32) == intrinsic.weight,
        fvar::AxisTag::WIDTH => width_from_percent(coord.value) == intrinsic.width,
        fvar::AxisTag::ITALIC | fvar::AxisTag::SLANT => {
            (coord.value != 0.0) == (intrinsic.slant != Slant::Upright)
        }
        // An axis with no bearing on the style this engine models — optical
        // size, grade, a custom axis — constrains nothing.
        _ => true,
    })
}

/// One record per selectable location of a variable `base`.
///
/// Two sources, in priority order:
///
/// 1. every `fvar` **named instance**, which the designer named outright;
/// 2. every `STAT` **combination** that lands somewhere else — the axis-value
///    names composed in the font's own axis order, which is what lets
///    `"Dx Variable Condensed SemiBold"` be read even when `fvar` never named
///    that point.
///
/// Locations already named by `fvar` are not composed a second time: the
/// designer's own name wins, and two records at one location would look like an
/// ambiguity to [`pick_unique`].
///
/// Instances are separate records rather than a list hanging off the variable
/// face because a document selects one of them by name; see the
/// [`face`](super::face) module doc.
fn instance_records(typeface: &Typeface, base: &FaceRecord) -> Vec<FaceRecord> {
    let Some(bytes) = read_table(typeface, b"fvar") else {
        return Vec::new();
    };
    let Some(variations) = log_malformed(fvar::read(&bytes), "fvar", &base.canonical_family) else {
        return Vec::new();
    };
    if variations.instances.is_empty() {
        return Vec::new();
    }
    let names_table = read_table(typeface, b"name")
        .and_then(|bytes| name::read(&bytes).ok())
        .unwrap_or_default();

    let family = base
        .typographic_family
        .as_deref()
        .unwrap_or(&base.canonical_family);

    let mut out = Vec::new();
    for instance in &variations.instances {
        // The default location is the face itself — the bytes already draw it,
        // and adding a record for it would make every variable face resolvable
        // two ways with no difference between them.
        if instance.is_default(&variations.axes) {
            continue;
        }
        let Some(subfamily) = names_table.preferred(NameId::from_raw(instance.subfamily_name_id))
        else {
            continue;
        };
        let post_script_name = instance
            .post_script_name_id
            .and_then(|id| names_table.preferred(NameId::from_raw(id)))
            .map(str::to_owned);

        let mut names = vec![FaceName::composed(
            NameKind::Instance,
            format!("{family} {subfamily}"),
        )];
        if family != base.canonical_family {
            push_unique_name(
                &mut names,
                FaceName::composed(
                    NameKind::Instance,
                    format!("{} {subfamily}", base.canonical_family),
                ),
            );
        }
        if let Some(ps) = &post_script_name {
            push_unique_name(&mut names, FaceName::composed(NameKind::Instance, ps));
        }

        out.push(FaceRecord {
            identity: base.identity.clone(),
            canonical_family: base.canonical_family.clone(),
            typographic_family: base.typographic_family.clone(),
            intrinsic: intrinsic_at(&base.intrinsic, &instance.coords, &variations),
            names,
            instance: Some(VariationInstance {
                subfamily: subfamily.to_owned(),
                post_script_name,
                coords: instance.coords.clone(),
            }),
        });
    }

    out.extend(composed_instance_records(
        typeface,
        base,
        &variations,
        &names_table,
        &out,
    ));
    out
}

/// `STAT` combinations of a variable face that no `fvar` instance already names.
fn composed_instance_records(
    typeface: &Typeface,
    base: &FaceRecord,
    variations: &Variations,
    names_table: &NameRecords,
    named: &[FaceRecord],
) -> Vec<FaceRecord> {
    let Some(stat) = read_table(typeface, b"STAT")
        .and_then(|bytes| log_malformed(stat::read(&bytes), "STAT", &base.canonical_family))
    else {
        return Vec::new();
    };
    let family = base
        .typographic_family
        .as_deref()
        .unwrap_or(&base.canonical_family);

    let mut out = Vec::new();
    for (composed, coords) in composed_style_names(family, &stat, names_table) {
        // Every axis the combination mentions has to exist in the design space,
        // or the location is not one this font can actually be set to.
        if coords.iter().any(|c| variations.axis(c.axis).is_none()) {
            continue;
        }
        let full = fill_defaults(&coords, variations);
        // The default location is the base record; a named instance already has
        // its own record and its own name.
        if full.iter().all(|c| {
            variations
                .axis(c.axis)
                .is_some_and(|a| a.default == c.value)
        }) {
            continue;
        }
        if named.iter().any(|record| {
            record
                .instance
                .as_ref()
                .is_some_and(|instance| instance.coords == full)
        }) {
            continue;
        }
        out.push(FaceRecord {
            identity: base.identity.clone(),
            canonical_family: base.canonical_family.clone(),
            typographic_family: base.typographic_family.clone(),
            intrinsic: intrinsic_at(&base.intrinsic, &full, variations),
            names: vec![FaceName::composed(NameKind::ComposedStyle, &composed)],
            instance: Some(VariationInstance {
                subfamily: composed
                    .strip_prefix(family)
                    .map(|s| s.trim().to_owned())
                    .unwrap_or_else(|| composed.clone()),
                post_script_name: None,
                coords: full,
            }),
        });
    }
    out
}

/// Complete a partial location with each axis's default, in the design space's
/// own axis order — the order `fvar` instances use, so the two are comparable.
fn fill_defaults(coords: &[VariationCoord], variations: &Variations) -> Vec<VariationCoord> {
    variations
        .axes
        .iter()
        .map(|axis| VariationCoord {
            axis: axis.tag,
            value: coords
                .iter()
                .find(|c| c.axis == axis.tag)
                .map_or(axis.default, |c| c.value),
        })
        .collect()
}

/// The intrinsic style a face takes on at a design-space location.
///
/// `wght` and `wdth` are *user* coordinates on exactly the scales `OS/2` uses —
/// weight 1–1000, width as a percentage of normal — so they translate directly.
/// `ital` and `slnt` decide slant: any non-zero value on either makes the
/// instance italic.
fn intrinsic_at(
    base: &IntrinsicStyle,
    coords: &[VariationCoord],
    variations: &Variations,
) -> IntrinsicStyle {
    use skia_safe::font_style::Slant;

    let mut out = *base;
    for coord in coords {
        let clamped = variations
            .axis(coord.axis)
            .map_or(coord.value, |axis: &Axis| axis.clamp(coord.value));
        match coord.axis {
            fvar::AxisTag::WEIGHT => out.weight = clamped.round() as i32,
            fvar::AxisTag::WIDTH => out.width = width_from_percent(clamped),
            fvar::AxisTag::ITALIC | fvar::AxisTag::SLANT if clamped != 0.0 => {
                out.slant = Slant::Italic;
            }
            _ => {}
        }
    }
    // A named instance is a position, not a family slot: it claims neither the
    // Bold nor the Regular slot, whatever the variable face's own bits said.
    out.bold_slot = false;
    out.regular_slot = false;
    out
}

/// `wdth` is a percentage of normal width; `usWidthClass` is a 1–9 bucket. The
/// boundaries are the midpoints of the percentages the OpenType spec assigns to
/// each class (50, 62.5, 75, 87.5, 100, 112.5, 125, 150, 200).
fn width_from_percent(percent: f32) -> skia_safe::font_style::Width {
    use skia_safe::font_style::Width;
    match percent {
        p if p < 56.25 => Width::ULTRA_CONDENSED,
        p if p < 68.75 => Width::EXTRA_CONDENSED,
        p if p < 81.25 => Width::CONDENSED,
        p if p < 93.75 => Width::SEMI_CONDENSED,
        p if p < 106.25 => Width::NORMAL,
        p if p < 118.75 => Width::SEMI_EXPANDED,
        p if p < 137.5 => Width::EXPANDED,
        p if p < 175.0 => Width::EXTRA_EXPANDED,
        _ => Width::ULTRA_EXPANDED,
    }
}

fn push_unique_name(names: &mut Vec<FaceName>, candidate: FaceName) {
    if candidate.text.trim().is_empty() {
        return;
    }
    if names
        .iter()
        .any(|n| n.kind == candidate.kind && n.text == candidate.text)
    {
        return;
    }
    names.push(candidate);
}

/// Read one table's bytes.
///
/// `copy_table_data`, never `to_font_data`: see the [`opentype`](super::opentype)
/// module doc for the 549 MB reason.
fn read_table(typeface: &Typeface, tag: &[u8; 4]) -> Option<Vec<u8>> {
    typeface
        .copy_table_data(table_tag(tag))
        .map(|data| data.as_bytes().to_vec())
}

/// Log a malformed table at `debug` and carry on without it.
///
/// Not `warn`: shipping fonts are routinely a little malformed, and a host with
/// a few hundred faces would produce a wall of warnings that says nothing about
/// the document being converted. What a user *can* act on is a font that failed
/// to resolve, and that is reported by the resolver.
fn log_malformed<T, E: std::fmt::Display>(
    result: Result<T, E>,
    table: &str,
    family: &str,
) -> Option<T> {
    match result {
        Ok(value) => Some(value),
        Err(err) => {
            log::debug!("[font] '{family}': unreadable {table} table ({err})");
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::opentype::fvar::tests::build_fvar;
    use crate::render::fonts::opentype::name::tests::build_name_table;
    use crate::render::fonts::opentype::stat::tests::{build_stat, Value};
    use skia_safe::font_style::{Slant, Weight, Width};

    const WIN: u16 = 3;
    const EN: u16 = 0x0409;

    fn names_for(records: &[(u16, u16, u16, u16, &str)]) -> NameRecords {
        name::read(&build_name_table(0, records, &[])).unwrap()
    }

    // ── STAT-composed names ──────────────────────────────────────────────

    #[test]
    fn stat_composes_family_plus_style_in_the_font_s_axis_order() {
        let stat = stat::read(&build_stat(
            1,
            &[(b"wdth", 256, 0), (b"wght", 257, 1)],
            &[
                Value::Single(0, 0x0002, 300, 100.0), // Normal, elided
                Value::Single(0, 0, 301, 75.0),       // Condensed
                Value::Single(1, 0x0002, 302, 400.0), // Regular, elided
                Value::Single(1, 0, 303, 600.0),      // SemiBold
            ],
            302,
        ))
        .unwrap();
        let names = names_for(&[
            (WIN, 1, EN, 301, "Condensed"),
            (WIN, 1, EN, 303, "SemiBold"),
        ]);

        let composed: Vec<String> = composed_style_names("Fixture Sans", &stat, &names)
            .into_iter()
            .map(|(name, _)| name)
            .collect();
        assert!(composed.contains(&"Fixture Sans Condensed".to_string()));
        assert!(composed.contains(&"Fixture Sans SemiBold".to_string()));
        assert!(
            composed.contains(&"Fixture Sans Condensed SemiBold".to_string()),
            "width precedes weight because the font's axis ordering says so, got {composed:?}"
        );
        assert!(
            !composed.contains(&"Fixture Sans SemiBold Condensed".to_string()),
            "the reverse order is not a name this font sanctions"
        );

        // …and each name carries the location it denotes, which is what stops
        // the default face from answering to all of them.
        let with_coords = composed_style_names("Fixture Sans", &stat, &names);
        let (_, semibold) = with_coords
            .iter()
            .find(|(name, _)| name == "Fixture Sans SemiBold")
            .unwrap();
        assert_eq!(
            semibold
                .iter()
                .map(|c| (c.axis.to_string(), c.value))
                .collect::<Vec<_>>(),
            vec![("wght".to_string(), 600.0)]
        );
    }

    /// Elided values are how `"Regular"` and `"Normal"` stay out of a composed
    /// name — without this every face would also answer to
    /// `"Fixture Sans Regular Normal"`.
    #[test]
    fn elided_axis_values_contribute_no_words() {
        let stat = stat::read(&build_stat(
            1,
            &[(b"wght", 256, 0)],
            &[Value::Single(0, 0x0002, 300, 400.0)],
            300,
        ))
        .unwrap();
        let names = names_for(&[(WIN, 1, EN, 300, "Regular")]);
        assert!(composed_style_names("Fixture Sans", &stat, &names).is_empty());
    }

    /// The bare family is indexed as a family name, not as a composed style.
    #[test]
    fn the_family_alone_is_not_emitted_as_a_composed_style() {
        let stat = stat::read(&build_stat(
            1,
            &[(b"wght", 256, 0)],
            &[Value::Single(0, 0, 300, 600.0)],
            300,
        ))
        .unwrap();
        let names = names_for(&[(WIN, 1, EN, 300, "SemiBold")]);
        let composed: Vec<String> = composed_style_names("Fixture Sans", &stat, &names)
            .into_iter()
            .map(|(name, _)| name)
            .collect();
        assert_eq!(composed, vec!["Fixture Sans SemiBold"]);
    }

    /// A pathological axis count must truncate rather than hang.
    #[test]
    fn the_composed_cross_product_is_capped() {
        let axes: Vec<(&[u8; 4], u16, u16)> = vec![
            (b"wght", 256, 0),
            (b"wdth", 257, 1),
            (b"opsz", 258, 2),
            (b"slnt", 259, 3),
        ];
        let mut values = Vec::new();
        let mut records: Vec<(u16, u16, u16, u16, &str)> = Vec::new();
        const WORDS: [&str; 6] = ["Aa", "Bb", "Cc", "Dd", "Ee", "Ff"];
        for axis in 0..4u16 {
            for (w, word) in WORDS.iter().enumerate() {
                let name_id = 300 + axis * 10 + w as u16;
                values.push(Value::Single(axis, 0, name_id, 100.0 + w as f32));
                records.push((WIN, 1, EN, name_id, word));
            }
        }
        let stat = stat::read(&build_stat(1, &axes, &values, 300)).unwrap();
        let composed = composed_style_names("Big", &stat, &names_for(&records));
        assert!(
            composed.len() <= MAX_COMPOSED_NAMES,
            "got {} names",
            composed.len()
        );
    }

    // ── variation instances ──────────────────────────────────────────────

    #[test]
    fn width_percentages_bucket_onto_the_width_class_scale() {
        assert_eq!(width_from_percent(50.0), Width::ULTRA_CONDENSED);
        assert_eq!(width_from_percent(75.0), Width::CONDENSED);
        assert_eq!(width_from_percent(100.0), Width::NORMAL);
        assert_eq!(width_from_percent(125.0), Width::EXPANDED);
        assert_eq!(width_from_percent(200.0), Width::ULTRA_EXPANDED);
        // Boundaries land on the nearer class, not the lower one.
        assert_eq!(width_from_percent(94.0), Width::NORMAL);
        assert_eq!(width_from_percent(93.0), Width::SEMI_CONDENSED);
    }

    /// An instance's coordinates become its intrinsic style — that is what lets
    /// a `wght = 600` instance be ranked as Semibold rather than as whatever the
    /// variable face's default was.
    #[test]
    fn an_instance_takes_its_style_from_its_coordinates() {
        let variations = fvar::read(&build_fvar(
            &[
                (b"wght", 100.0, 400.0, 900.0, false),
                (b"wdth", 75.0, 100.0, 125.0, false),
            ],
            &[(300, &[600.0, 75.0], None)],
        ))
        .unwrap();
        let at = intrinsic_at(
            &IntrinsicStyle::NEUTRAL,
            &variations.instances[0].coords,
            &variations,
        );
        assert_eq!(at.weight, 600);
        assert_eq!(at.width, Width::CONDENSED);
        assert!(
            !at.bold_slot && !at.regular_slot,
            "an instance is a position, not a family slot"
        );
    }

    #[test]
    fn a_nonzero_slant_or_italic_axis_makes_an_instance_italic() {
        for tag in [b"ital", b"slnt"] {
            let variations = fvar::read(&build_fvar(
                &[(tag, -10.0, 0.0, 10.0, false)],
                &[(300, &[10.0], None)],
            ))
            .unwrap();
            let at = intrinsic_at(
                &IntrinsicStyle::NEUTRAL,
                &variations.instances[0].coords,
                &variations,
            );
            assert_eq!(at.slant, Slant::Italic, "{}", String::from_utf8_lossy(tag));
        }
    }

    /// A coordinate outside the axis range would ask Skia for an instance the
    /// font does not define.
    #[test]
    fn an_out_of_range_coordinate_is_clamped_to_the_axis() {
        let variations = fvar::read(&build_fvar(
            &[(b"wght", 100.0, 400.0, 900.0, false)],
            &[(300, &[400.0], None)],
        ))
        .unwrap();
        let at = intrinsic_at(
            &IntrinsicStyle::NEUTRAL,
            &[VariationCoord {
                axis: fvar::AxisTag::WEIGHT,
                value: 5000.0,
            }],
            &variations,
        );
        assert_eq!(at.weight, 900);
    }

    // ── ambiguity ────────────────────────────────────────────────────────

    fn record(family: &str, weight: i32, name: &str) -> FaceRecord {
        FaceRecord {
            identity: FaceIdentity::System {
                family: family.to_owned(),
                face_index: 0,
            },
            canonical_family: family.to_owned(),
            typographic_family: None,
            intrinsic: IntrinsicStyle {
                weight,
                ..IntrinsicStyle::NEUTRAL
            },
            names: vec![FaceName::asserted(NameId::Full, name)],
            instance: None,
        }
    }

    fn candidates(records: &[FaceRecord]) -> Vec<(FaceRef, FaceRecord)> {
        records
            .iter()
            .enumerate()
            .map(|(i, r)| (FaceRef::Deep(i), r.clone()))
            .collect()
    }

    #[test]
    fn a_single_candidate_resolves() {
        let records = [record("Alpha", 400, "Alpha Regular")];
        assert_eq!(
            pick_unique(&candidates(&records)),
            FaceLookup::Found(FaceRef::Deep(0))
        );
    }

    #[test]
    fn no_candidates_is_a_miss() {
        assert_eq!(pick_unique(&[]), FaceLookup::Missing);
    }

    /// Acceptance criterion: *ambiguous aliases fail safely rather than choosing
    /// insertion order*. Two different faces under one name must not silently
    /// resolve to whichever was indexed first.
    #[test]
    fn two_different_faces_under_one_name_are_ambiguous() {
        let records = [
            record("Alpha", 400, "Shared Name"),
            record("Beta", 700, "Shared Name"),
        ];
        assert_eq!(
            pick_unique(&candidates(&records)),
            FaceLookup::Ambiguous { candidates: 2 }
        );
    }

    /// …but the *same* font reachable twice is not an ambiguity. A host that
    /// installs a font both per-user and system-wide would otherwise make every
    /// one of its faces unresolvable.
    #[test]
    fn the_same_face_indexed_twice_is_not_ambiguous() {
        let records = [
            record("Alpha", 400, "Alpha Regular"),
            record("Alpha", 400, "Alpha Regular"),
        ];
        assert_eq!(
            pick_unique(&candidates(&records)),
            FaceLookup::Found(FaceRef::Deep(0))
        );
    }

    /// Two named instances of one variable face differ only in coordinates, and
    /// that difference has to count.
    #[test]
    fn instances_at_different_locations_are_different_faces() {
        let at = |weight: f32| {
            let mut r = record("Alpha", weight as i32, "Alpha SemiBold");
            r.instance = Some(VariationInstance {
                subfamily: "SemiBold".into(),
                post_script_name: None,
                coords: vec![VariationCoord {
                    axis: fvar::AxisTag::WEIGHT,
                    value: weight,
                }],
            });
            r
        };
        let records = [at(600.0), at(700.0)];
        assert_eq!(
            pick_unique(&candidates(&records)),
            FaceLookup::Ambiguous { candidates: 2 }
        );
    }

    // ── name-kind filtering, against a real host catalogue ───────────────

    #[test]
    fn lookups_are_case_and_whitespace_insensitive() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let first = family_names(&catalog)[0].clone();
        let padded = format!("  {}  ", first.to_uppercase());
        assert_eq!(
            catalog.family_faces(&padded).map(|f| f.len()),
            catalog.family_faces(&first).map(|f| f.len())
        );
    }

    #[test]
    fn an_empty_name_never_matches() {
        let catalog = FaceCatalog::build(FontMgr::new());
        assert!(catalog
            .candidates(Scope::Embedded, "   ", |_| true)
            .is_empty());
        assert!(catalog.family_faces("   ").is_none());
    }

    // ── the one manager answer taken on trust ───────────────────────────

    /// The guard that keeps the resolution chain reachable. A matcher that
    /// substitutes (fontconfig does; CoreText does not) would otherwise report
    /// every miss as a hit, and steps 4 onwards would be dead code.
    ///
    /// Tested on the predicate rather than through `adopt_unenumerated`, because
    /// on a non-substituting host the rejection path is unreachable from there —
    /// the test would pass without the guard existing.
    #[test]
    fn the_family_guard_rejects_a_substituted_face() {
        let mgr = FontMgr::new();
        let typeface = mgr
            .legacy_make_typeface(None::<&str>, FontStyle::normal())
            .expect("system default typeface");
        let family = typeface.family_name();

        assert!(carries_family(&typeface, &family), "its own family matches");
        assert!(
            carries_family(&typeface, &family.to_uppercase()),
            "the comparison is case-insensitive"
        );
        assert!(
            !carries_family(&typeface, "No Such Family 8f3a1c"),
            "a face carrying a different family must be refused — this is the \
             whole reason the rest of the chain is reachable"
        );
    }

    /// And a family the host genuinely has is still adopted, so the guard is not
    /// simply rejecting everything.
    #[test]
    fn a_matchable_family_is_still_adopted() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let Some(first) = family_names(&catalog).first().cloned() else {
            return;
        };
        assert!(catalog.family_faces(&first).is_some());
        assert!(
            catalog.family_faces("No Such Family 8f3a1c").is_none(),
            "a name the manager cannot match must not be adopted"
        );
    }

    // ── tiering ──────────────────────────────────────────────────────────

    fn family_names(catalog: &FaceCatalog) -> Vec<String> {
        catalog
            .families
            .borrow()
            .iter()
            .map(|f| f.name.clone())
            .collect()
    }

    /// Construction must not touch a single family. `match_family` is 28 ms
    /// across a 210-family host and `family_names` is 0.04 ms — paying only the
    /// second is the whole point of tier 0.
    #[test]
    fn construction_enumerates_names_only() {
        let catalog = FaceCatalog::build(FontMgr::new());
        assert!(
            catalog.family_count() > 0,
            "the test host must expose at least one family"
        );
        assert!(!catalog.deep_is_built());
        assert!(
            !catalog.family_is_enumerated(&family_names(&catalog)[0]),
            "constructing the catalogue must not enumerate any family's faces"
        );
    }

    #[test]
    fn touching_one_family_enumerates_only_that_family() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let names = family_names(&catalog);
        if names.len() < 2 {
            return;
        }

        assert!(catalog.family_faces(&names[0]).is_some());
        assert!(catalog.family_is_enumerated(&names[0]));
        assert!(
            !catalog.family_is_enumerated(&names[1]),
            "an unrelated family must not have been touched"
        );
        assert!(!catalog.deep_is_built(), "nor the deep tier");
    }

    #[test]
    fn family_lookup_is_case_and_whitespace_folded() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let first = family_names(&catalog)[0].clone();
        assert!(catalog.family_faces(&first).is_some());
        assert!(catalog.family_faces(&first.to_uppercase()).is_some());
        assert!(catalog.family_faces("No Such Family 8f3a1c").is_none());
    }

    #[test]
    fn a_manager_face_carries_manager_names_only() {
        let record = manager_face_record(
            "Inter",
            2,
            FontStyle::new(Weight::from(300), Width::NORMAL, Slant::Upright),
            Some("Light".to_owned()),
        );
        assert_eq!(record.canonical_family, "Inter");
        assert_eq!(record.intrinsic.weight, 300);
        assert_eq!(
            record.identity,
            FaceIdentity::System {
                family: "Inter".into(),
                face_index: 2
            }
        );
        assert!(
            record.names.iter().all(|n| !n.kind.is_asserted()),
            "nothing here was read from the font, so nothing may claim to be"
        );
        assert!(record
            .names
            .iter()
            .any(|n| n.kind == NameKind::ManagerFace && n.text == "Inter Light"));
    }

    #[test]
    fn a_face_with_no_style_name_gets_no_composed_face_name() {
        let record = manager_face_record("Inter", 0, FontStyle::normal(), Some("  ".to_owned()));
        assert_eq!(record.names.len(), 1);
        assert_eq!(record.names[0].kind, NameKind::ManagerFamily);
    }

    /// A `FaceRef` must round-trip back to the record it names, or the resolver
    /// would select one face and the registry open another.
    #[test]
    fn a_manager_face_ref_round_trips() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let Some(faces) = catalog.family_faces(&family_names(&catalog)[0]) else {
            return;
        };
        for (face_ref, record) in faces {
            assert_eq!(catalog.face(face_ref).as_ref(), Some(&record));
        }
    }

    /// Issue #188: a name the host does not enumerate is asked about once per
    /// render, whichever way the manager answers — `preload` asks for four
    /// style variants of every family, and on fontconfig every probe is an
    /// `FcFontMatch` plus an `FT_Open_Face`.
    #[test]
    fn an_unenumerable_name_is_probed_once() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let name = "No Such Family 8f3a1c";
        assert!(catalog.family_faces(name).is_none());
        assert!(catalog.family_faces(name).is_none());
        assert_eq!(
            catalog.adoption_probes(),
            1,
            "the second ask is answered from memory"
        );
    }

    // ── the deep tier, against whatever the host actually has ────────────

    #[test]
    fn the_deep_tier_builds_once() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let _ = catalog.lookup_deep("no such face 8f3a1c", |_| true);
        assert!(catalog.deep_is_built());
        let first = catalog.deep().records.len();

        let _ = catalog.lookup_deep("another miss 8f3a1c", |_| true);
        assert_eq!(catalog.deep().records.len(), first, "built exactly once");
    }

    /// Every deep record must be findable by the names it claims, and every
    /// `FaceRef` must round-trip. This is the invariant that keeps a catalogue
    /// entry and the typeface it yields from drifting apart.
    #[test]
    fn every_deep_record_is_findable_by_its_own_asserted_names() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let _ = catalog.lookup_deep("prime the tier 8f3a1c", |_| true);

        let mut checked = 0;
        let records = catalog.deep().records.clone();
        for record in &records {
            for name in record.names_matching(NameKind::is_primary_identity) {
                match catalog.lookup_deep(&name.text, NameKind::is_asserted) {
                    FaceLookup::Found(face_ref) => {
                        assert!(catalog.face(face_ref).is_some(), "ref must round-trip");
                        checked += 1;
                    }
                    FaceLookup::Ambiguous { .. } => checked += 1,
                    FaceLookup::Missing => {
                        panic!("'{}' is indexed but not findable", name.text)
                    }
                }
            }
            if checked > 200 {
                break;
            }
        }
        assert!(
            checked > 0,
            "the test host exposes no face with a full or PostScript name"
        );
    }

    /// The deep tier is a *full* pass: it must not silently cover a fraction of
    /// the host, which is what a borrow-juggling bug in its loop would look like.
    #[test]
    fn the_deep_tier_covers_every_manager_family() {
        let catalog = FaceCatalog::build(FontMgr::new());
        let _ = catalog.lookup_deep("prime the tier 8f3a1c", |_| true);

        let covered: std::collections::HashSet<String> = catalog
            .deep()
            .records
            .iter()
            .filter_map(|r| match &r.identity {
                FaceIdentity::System { family, .. } => Some(family.clone()),
                FaceIdentity::Embedded { .. } => None,
            })
            .collect();
        let enumerated = catalog.family_count();
        assert!(
            covered.len() >= enumerated.saturating_sub(enumerated / 10),
            "deep tier covers {} of {enumerated} families",
            covered.len()
        );
    }
}
