//! The eight-step face-resolution chain.
//!
//! A **pure function** of a [`FaceRequest`] and a [`FaceCatalog`]. Purity is the
//! design goal, not a side effect: it is what lets the whole chain be tested
//! against a catalogue built from committed fixture fonts and behave identically
//! on macOS, Linux and Windows.
//!
//! # The order, and why it is this order
//!
//! | Step | Matches | Kind of evidence |
//! |---|---|---|
//! | 1 | an embedded font's full, compatible-full, PostScript or instance name | the document's own font, by face |
//! | 2 | an embedded font's family | the document's own font, by family |
//! | 3 | a host family, exactly | the host, by family |
//! | 3b | a metric-compatible substitute, when the key's own family is absent | a curated table — placed here because step 4 is the expensive one |
//! | 4 | a host face's full, compatible-full or PostScript name | the font's `name` table |
//! | 5 | any other name the font sanctions — typographic, WWS, localized, a `STAT` style, an `fvar` instance | the font's own tables |
//! | 6 | a family plus a trailing weight word, parsed | a guess |
//! | 7 | a metric-compatible substitute | a curated table |
//! | 8 | the host default | nothing |
//!
//! Everything down to step 5 is something a font *asserts about itself*. Step 6
//! is the first guess, and it is the step this whole issue demoted: it used to
//! run before any metadata was read, so `"PT Sans Narrow"` and `"Foo Medium"`
//! were chopped into a base family plus a style word even when the host had a
//! family by exactly that name.
//!
//! # Ambiguity is an outcome, not a coin toss
//!
//! When a name matches several genuinely different faces the chain does not pick
//! one. It records the ambiguity and continues, and if nothing later matches
//! either, the ambiguity — not a bare "not found" — is what
//! [`FallbackReason`] reports. That is the difference between a diagnostic
//! someone can act on and one that sends them looking for a font that is, in
//! fact, installed twice.
//!
//! # What "selected" does not include
//!
//! Opening the face. The resolver returns a [`FaceRef`] and the effective style;
//! turning that into a Skia typeface — including applying variation coordinates
//! — is [`FontRegistry`](super::FontRegistry)'s job, because only the registry
//! owns embedded font bytes. Keeping the two apart is also what keeps this
//! module pure.

use skia_safe::font_style::Weight;

use super::catalog::{FaceCatalog, FaceLookup, Scope};
use super::face::{FaceRecord, IntrinsicStyle, NameKind};
use super::opentype::NameId;
use super::request::{EffectiveStyle, FaceRequest};
use super::substitutes::{canonical_weight_names, substitutes_for};
use super::FaceRef;

/// Which step of the chain produced a selection.
///
/// Carried through to the log so `RUST_LOG=debug` says not just *what* a name
/// resolved to but *on what evidence* — the fastest way to tell a correct
/// metadata match from a lucky substitution.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ResolutionStep {
    /// 1 — an embedded font, matched by a name that identifies one face.
    EmbeddedFace,
    /// 2 — an embedded font, matched by family.
    EmbeddedFamily,
    /// 3 — a host family by exactly the requested name.
    SystemFamily,
    /// 4 — a host face's full, compatible-full or PostScript name.
    SystemFaceName,
    /// 5 — any other name the font's own tables sanction.
    SystemMetadataAlias,
    /// 6 — the requested name parsed as `family + weight word`.
    ParsedFaceName { base_family: String, weight: i32 },
    /// 7 — a metric-compatible substitute.
    Substitution {
        /// The `FONT_SUBSTITUTIONS` key that matched,
        /// which may be the request's base family rather than the request.
        via: &'static str,
        substitute: &'static str,
    },
}

impl ResolutionStep {
    /// One-word label for logs.
    pub fn label(&self) -> &'static str {
        match self {
            Self::EmbeddedFace => "embedded face",
            Self::EmbeddedFamily => "embedded family",
            Self::SystemFamily => "system family",
            Self::SystemFaceName => "face name",
            Self::SystemMetadataAlias => "metadata alias",
            Self::ParsedFaceName { .. } => "parsed name",
            Self::Substitution { .. } => "substitute",
        }
    }
}

/// Why the chain fell through to the host default.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum FallbackReason {
    /// Nothing matched at any step.
    NoCandidate,
    /// A name matched several genuinely different faces, so selecting one would
    /// have been a guess, and no later step matched either.
    AmbiguousName {
        step: ResolutionStep,
        candidates: usize,
    },
}

/// A face the chain chose.
#[derive(Clone, Debug, PartialEq)]
pub struct Selection {
    pub face: FaceRef,
    pub record: FaceRecord,
    /// The style the request works out to. The chosen face is usually but not
    /// always at exactly this style — a family with no Bold answers a bold
    /// request with the heaviest it has.
    pub style: EffectiveStyle,
    pub step: ResolutionStep,
}

/// What the chain concluded.
#[derive(Clone, Debug, PartialEq)]
pub enum FaceResolution {
    Selected(Box<Selection>),
    /// Nothing in the catalogue answered. The caller falls back to the host
    /// default and reports `reason`.
    Default {
        reason: FallbackReason,
    },
}

/// Names that identify exactly one face — steps 1 and 4.
fn identifies_one_face(kind: NameKind) -> bool {
    kind.is_primary_identity()
}

/// Names that identify a family whose faces are then ranked — steps 2 and 3.
fn identifies_a_family(kind: NameKind) -> bool {
    matches!(
        kind,
        NameKind::Table(NameId::Family | NameId::TypographicFamily | NameId::WwsFamily)
            | NameKind::ManagerFamily
    )
}

/// Everything else a font sanctions about itself — step 5.
///
/// [`NameKind::ManagerFace`] is here rather than at step 6 because it *is*
/// metadata, just the platform's rather than the file's: the manager derived it
/// from the same `name` table, and it is the alias that used to carry the whole
/// face-name story before this module existed.
fn is_metadata_alias(kind: NameKind) -> bool {
    matches!(
        kind,
        NameKind::Instance
            | NameKind::ComposedStyle
            | NameKind::ManagerFace
            | NameKind::Table(NameId::UniqueId)
    )
}

/// Run the chain.
pub fn resolve(request: &FaceRequest<'_>, catalog: &FaceCatalog) -> FaceResolution {
    let mut ambiguity: Option<(ResolutionStep, usize)> = None;

    // 1 — the document's own font, by a name that names one face.
    if let Some(selection) = by_face_name(
        request,
        catalog,
        Scope::Embedded,
        identifies_one_face,
        ResolutionStep::EmbeddedFace,
        &mut ambiguity,
    ) {
        return selected(selection);
    }

    // 2 — the document's own font, by family.
    if let Some(selection) = by_family_names(
        request,
        catalog,
        Scope::Embedded,
        identifies_a_family,
        ResolutionStep::EmbeddedFamily,
    ) {
        return selected(selection);
    }

    // 3 — a host family by exactly this name. Cheap: one `match_family`.
    if let Some(faces) = catalog.family_faces(request.name) {
        let style = EffectiveStyle::resolve(request, IntrinsicStyle::NEUTRAL);
        if let Some(selection) = rank(&faces, &style, ResolutionStep::SystemFamily) {
            return selected(selection);
        }
    }

    // 3b — a curated substitute for a table key whose own family the host
    // lacks. Placed *before* the metadata tier on purpose: reaching step 4 is
    // what builds that tier, and on fontconfig building it instantiates every
    // face on the host — three `FT_Open_Face` per face, up to 1024 of them
    // retained — so a document naming Arial on a machine with Liberation Sans
    // paid seconds and hundreds of megabytes before the table that answers
    // "Arial → Liberation Sans" was consulted at all (issue #188).
    //
    // The guard keeps the metadata path for keys whose *base* family is
    // present: "Segoe UI Light" with Segoe UI installed still goes looking for
    // the real Light face. What the shortcut gives up is a face that carries
    // the key only as a typographic or localized alias (Arial Narrow's
    // typographic family is "Arial"); for a request by family name, the
    // metric-compatible substitute is the better answer than a face that
    // merely mentions it.
    if let Some((via, substitutes)) = substitutes_for(request.name) {
        // Step 3 just asked the host about `request.name`; a key equal to it
        // is known absent without a second probe.
        let key_family_present =
            !via.eq_ignore_ascii_case(request.name) && catalog.family_faces(via).is_some();
        if !key_family_present {
            if let Some(selection) = substitute(request, catalog, via, substitutes) {
                return selected(selection);
            }
        }
    }

    // 4 — a host face's own full/compatible-full/PostScript name. Reaching here
    // is what builds the catalogue's deep tier; everything above is cheap.
    log::debug!("[font] '{}' reached the metadata tier", request.name);
    if let Some(selection) = by_face_name(
        request,
        catalog,
        Scope::Deep,
        identifies_one_face,
        ResolutionStep::SystemFaceName,
        &mut ambiguity,
    ) {
        return selected(selection);
    }

    // 5 — any other name the font's tables sanction: typographic and WWS
    // families, localized spellings, `STAT` styles, `fvar` instances.
    if let Some(selection) = by_face_name(
        request,
        catalog,
        Scope::Deep,
        is_metadata_alias,
        ResolutionStep::SystemMetadataAlias,
        &mut ambiguity,
    ) {
        return selected(selection);
    }
    if let Some(selection) = by_family_names(
        request,
        catalog,
        Scope::Deep,
        identifies_a_family,
        ResolutionStep::SystemMetadataAlias,
    ) {
        return selected(selection);
    }

    // 6 — the first guess: `family + trailing weight word`.
    if let Some((base_family, weight)) = parse_face_name(request.name) {
        if let Some(faces) = catalog.family_faces(base_family) {
            let style = EffectiveStyle::resolve(
                request,
                IntrinsicStyle {
                    weight,
                    ..IntrinsicStyle::NEUTRAL
                },
            );
            let step = ResolutionStep::ParsedFaceName {
                base_family: base_family.to_owned(),
                weight,
            };
            if let Some(selection) = rank(&faces, &style, step) {
                return selected(selection);
            }
        }
    }

    // 7 — metric-compatible substitution for the keys step 3b passed over:
    // the base family was present, but neither metadata nor the parsed weight
    // produced a face for it.
    if let Some((via, substitutes)) = substitutes_for(request.name) {
        if let Some(selection) = substitute(request, catalog, via, substitutes) {
            return selected(selection);
        }
    }

    // 8 — nothing answered.
    FaceResolution::Default {
        reason: match ambiguity {
            Some((step, candidates)) => FallbackReason::AmbiguousName { step, candidates },
            None => FallbackReason::NoCandidate,
        },
    }
}

/// Steps 3b and 7: rank the substitutes the curated table lists for `via`.
/// The parsed weight carries over, so `"Segoe UI Light"` asks its substitute
/// for Light rather than Regular.
fn substitute(
    request: &FaceRequest<'_>,
    catalog: &FaceCatalog,
    via: &'static str,
    substitutes: &'static [&'static str],
) -> Option<Selection> {
    let base = IntrinsicStyle {
        weight: parse_face_name(request.name).map_or(400, |(_, weight)| weight),
        ..IntrinsicStyle::NEUTRAL
    };
    let style = EffectiveStyle::resolve(request, base);
    substitutes.iter().find_map(|substitute| {
        let faces = catalog.family_faces(substitute)?;
        rank(
            &faces,
            &style,
            ResolutionStep::Substitution { via, substitute },
        )
    })
}

fn selected(selection: Selection) -> FaceResolution {
    FaceResolution::Selected(Box::new(selection))
}

/// A step that matches a name identifying one face.
///
/// The matched face fixes the *base* style — that is the whole point of reading
/// metadata: `"Calibri Light"` starts from the Light face's own 342 rather than
/// from a content-free 400. The toggles are then applied on top, and if they
/// move the style the result is re-ranked within the face's own family, so
/// `"Calibri Light"` with `<w:b/>` lands on Calibri Bold exactly as it should.
fn by_face_name(
    request: &FaceRequest<'_>,
    catalog: &FaceCatalog,
    scope: Scope,
    accept: impl Fn(NameKind) -> bool + Copy,
    step: ResolutionStep,
    ambiguity: &mut Option<(ResolutionStep, usize)>,
) -> Option<Selection> {
    let candidates = catalog.candidates(scope, request.name, accept);
    match super::catalog::pick_unique(&candidates) {
        FaceLookup::Missing => None,
        FaceLookup::Ambiguous { candidates } => {
            log::debug!(
                "[font] '{}': {candidates} different faces answer to this name at the {} step; \
                 declining to guess",
                request.name,
                step.label()
            );
            ambiguity.get_or_insert((step, candidates));
            None
        }
        FaceLookup::Found(face) => {
            let record = catalog.face(face)?;
            let style = EffectiveStyle::resolve(request, record.intrinsic);
            Some(refine(catalog, scope, face, record, &style, step))
        }
    }
}

/// A step that matches a family name and ranks the family's faces.
///
/// A family name is carried by every face in the family, so several candidates
/// are the expected case, not an ambiguity. They are collected across every
/// family the name reaches and ranked together.
fn by_family_names(
    request: &FaceRequest<'_>,
    catalog: &FaceCatalog,
    scope: Scope,
    accept: impl Fn(NameKind) -> bool + Copy,
    step: ResolutionStep,
) -> Option<Selection> {
    let candidates = catalog.candidates(scope, request.name, accept);
    let style = EffectiveStyle::resolve(request, IntrinsicStyle::NEUTRAL);
    rank(&candidates, &style, step)
}

/// Pick the best of several candidate faces for a style.
fn rank(
    candidates: &[(FaceRef, FaceRecord)],
    style: &EffectiveStyle,
    step: ResolutionStep,
) -> Option<Selection> {
    let (face, record) = candidates
        .iter()
        .enumerate()
        .min_by_key(|(order, (_, record))| score(record, style, *order))
        .map(|(_, pair)| pair)?;
    Some(Selection {
        face: *face,
        record: record.clone(),
        style: *style,
        step,
    })
}

/// Re-rank within the matched face's own family when the request's toggles have
/// moved the style away from the face the name identified.
///
/// Without this, `"Calibri Light"` with `<w:b/>` would return the Light face and
/// silently ignore the bold. With it, the name selects the family and the
/// toggles select within it — which is what the old `merged_alias_weight` was
/// approximating by taking a `max`.
fn refine(
    catalog: &FaceCatalog,
    scope: Scope,
    face: FaceRef,
    record: FaceRecord,
    style: &EffectiveStyle,
    step: ResolutionStep,
) -> Selection {
    let unchanged = Selection {
        face,
        record: record.clone(),
        style: *style,
        step: step.clone(),
    };
    // The common case: nothing asked for a different style, so the named face is
    // the answer and no family lookup is needed.
    if !style.weight_is_requested() && !style.slant_is_requested() {
        return unchanged;
    }
    let family = record
        .typographic_family
        .as_deref()
        .unwrap_or(&record.canonical_family);
    let siblings = catalog.candidates(scope, family, identifies_a_family);
    if siblings.is_empty() {
        return unchanged;
    }
    match rank(&siblings, style, step) {
        // A sibling only wins if it is genuinely a better fit; an equal score
        // must keep the face the *name* chose, or a bold request against an
        // already-bold face could wander to a different member of the family.
        Some(better)
            if score(&better.record, style, usize::MAX)
                < score(&unchanged.record, style, usize::MAX) =>
        {
            better
        }
        _ => unchanged,
    }
}

/// How badly a face fits a request. Lower is better; ordering is lexicographic
/// over the fields, which is what makes each one a strict tie-break of the last.
#[derive(PartialEq, Eq, PartialOrd, Ord, Debug)]
struct FaceScore {
    /// Slant first: an upright face where italic was asked for is wrong in a way
    /// no amount of weight agreement makes up for.
    slant: u8,
    /// Then width, on the 1–9 `usWidthClass` scale.
    width: u16,
    /// Then weight, asymmetric when the request explicitly asked to be bold —
    /// see [`weight_distance`].
    weight: u32,
    /// Then the family slot the face declares. Separates two faces at the same
    /// weight, one of which the font marks as *the* Regular or *the* Bold.
    slot: u8,
    /// Then a static face over a variable instance. `FontRegistry::open_uncached`
    /// bakes most instances into a static SFNT at open time (issue #113), but a
    /// CFF2 or WOFF/WOFF2-sourced one still cannot round-trip — and this ranking
    /// step runs before any typeface is opened, so it cannot yet tell which case
    /// a given instance is. Kept as a blanket penalty until resolution can see
    /// font flavor cheaply enough to narrow it; see `InstanceState::Unbaked`.
    instance: u8,
    /// Finally the candidate's position, so an exact tie is resolved the same
    /// way on every run rather than by hash order.
    order: usize,
}

fn score(record: &FaceRecord, style: &EffectiveStyle, order: usize) -> FaceScore {
    FaceScore {
        slant: u8::from(record.intrinsic.is_italic() != style.is_italic()),
        width: (*record.intrinsic.width - *style.width).unsigned_abs() as u16,
        weight: weight_distance(record.intrinsic.weight, style),
        slot: slot_penalty(record, style),
        instance: u8::from(record.instance.is_some()),
        order,
    }
}

/// Distance from a face's weight to the requested one.
///
/// Symmetric when the request did not name a weight. When it explicitly asked to
/// be bold, being *lighter* than asked is penalised at double rate: a document
/// that says `<w:b/>` is better served by something heavier than the request
/// than by something lighter, and a family whose only options straddle the
/// request should land on the heavier one.
fn weight_distance(face: i32, style: &EffectiveStyle) -> u32 {
    let delta = face - style.weight;
    if delta >= 0 {
        delta.unsigned_abs()
    } else if style.weight_is_requested() {
        delta.unsigned_abs().saturating_mul(2)
    } else {
        delta.unsigned_abs()
    }
}

/// Prefer the face that declares itself the slot the request implies.
fn slot_penalty(record: &FaceRecord, style: &EffectiveStyle) -> u8 {
    let wants_bold = style.weight >= *Weight::BOLD;
    let declares = if wants_bold {
        record.intrinsic.bold_slot
    } else {
        record.intrinsic.regular_slot
    };
    u8::from(!declares)
}

/// Split a face-qualified name into its base family and the weight its trailing
/// word implies: `"Segoe UI Light"` → `("Segoe UI", 300)`.
///
/// This is step 6, and it is a *guess* — the same guess the whole of resolution
/// used to open with. It survives because a font with no readable metadata, or a
/// family the host does not install at all, leaves nothing better to go on.
fn parse_face_name(name: &str) -> Option<(&str, i32)> {
    let base = super::substitutes::strip_weight_suffix(name)?;
    let suffix = name[base.len()..].trim();
    let weight = (100..=900)
        .step_by(100)
        .find(|&w| {
            canonical_weight_names(w)
                .iter()
                .any(|word| word.eq_ignore_ascii_case(suffix))
        })
        .unwrap_or(*Weight::NORMAL);
    Some((base, weight))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::face::{FaceIdentity, FaceName, VariationInstance};
    use crate::render::fonts::opentype::fvar::{AxisTag, VariationCoord};
    use crate::render::fonts::request::Toggle;
    use crate::render::fonts::FaceRequest;
    use skia_safe::font_style::{Slant, Width};

    fn face(family: &str, weight: i32, slant: Slant) -> FaceRecord {
        FaceRecord {
            identity: FaceIdentity::System {
                family: family.to_owned(),
                face_index: 0,
            },
            canonical_family: family.to_owned(),
            typographic_family: None,
            intrinsic: IntrinsicStyle {
                weight,
                slant,
                bold_slot: weight >= *Weight::BOLD,
                regular_slot: weight == *Weight::NORMAL && slant == Slant::Upright,
                width: Width::NORMAL,
            },
            names: vec![FaceName::composed(NameKind::ManagerFamily, family)],
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

    fn want(bold: Toggle, italic: Toggle, base: IntrinsicStyle) -> EffectiveStyle {
        EffectiveStyle::resolve(&FaceRequest::new("probe", bold, italic), base)
    }

    fn pick(records: &[FaceRecord], style: &EffectiveStyle) -> usize {
        match rank(&candidates(records), style, ResolutionStep::SystemFamily)
            .expect("a non-empty candidate list always ranks")
            .face
        {
            FaceRef::Deep(i) => i,
            other => panic!("unexpected ref {other:?}"),
        }
    }

    // ── ranking ──────────────────────────────────────────────────────────

    #[test]
    fn a_plain_request_picks_the_regular_face() {
        let family = [
            face("F", 700, Slant::Upright),
            face("F", 400, Slant::Upright),
            face("F", 400, Slant::Italic),
        ];
        assert_eq!(
            pick(
                &family,
                &want(Toggle::Absent, Toggle::Absent, IntrinsicStyle::NEUTRAL)
            ),
            1
        );
    }

    #[test]
    fn bold_and_italic_each_select_their_face() {
        let family = [
            face("F", 400, Slant::Upright),
            face("F", 700, Slant::Upright),
            face("F", 400, Slant::Italic),
            face("F", 700, Slant::Italic),
        ];
        let n = IntrinsicStyle::NEUTRAL;
        assert_eq!(pick(&family, &want(Toggle::On, Toggle::Absent, n)), 1);
        assert_eq!(pick(&family, &want(Toggle::Absent, Toggle::On, n)), 2);
        assert_eq!(pick(&family, &want(Toggle::On, Toggle::On, n)), 3);
    }

    /// Slant outranks weight: an upright face is never the answer to an italic
    /// request while any italic face exists, however far off its weight.
    #[test]
    fn slant_outranks_weight() {
        let family = [
            face("F", 400, Slant::Upright),
            face("F", 900, Slant::Italic),
        ];
        assert_eq!(
            pick(
                &family,
                &want(Toggle::Absent, Toggle::On, IntrinsicStyle::NEUTRAL)
            ),
            1
        );
    }

    /// An explicit bold with nothing bold available should land on the heaviest
    /// face, not the nearest by absolute distance.
    #[test]
    fn an_explicit_bold_prefers_heavier_over_nearer() {
        // 600 is 100 away from the requested 700; 300 is 400 away. Even at
        // double rate the heavier face wins — and it should, because the
        // document asked to be bold.
        let family = [
            face("F", 300, Slant::Upright),
            face("F", 600, Slant::Upright),
        ];
        assert_eq!(
            pick(
                &family,
                &want(Toggle::On, Toggle::Absent, IntrinsicStyle::NEUTRAL)
            ),
            1
        );

        // And the asymmetry is real: equidistant either side, the heavier wins.
        let straddle = [
            face("F", 600, Slant::Upright),
            face("F", 800, Slant::Upright),
        ];
        assert_eq!(
            pick(
                &straddle,
                &want(Toggle::On, Toggle::Absent, IntrinsicStyle::NEUTRAL)
            ),
            1
        );
    }

    /// With no explicit bold, weight distance is symmetric — nothing is
    /// "asking" for a direction.
    #[test]
    fn an_unrequested_weight_is_symmetric() {
        let base = IntrinsicStyle {
            weight: 500,
            ..IntrinsicStyle::NEUTRAL
        };
        let straddle = [
            face("F", 400, Slant::Upright),
            face("F", 600, Slant::Upright),
        ];
        // Equidistant: the earlier candidate wins by the stable order tie-break.
        assert_eq!(
            pick(&straddle, &want(Toggle::Absent, Toggle::Absent, base)),
            0
        );
    }

    #[test]
    fn width_outranks_weight_but_not_slant() {
        let mut condensed = face("F", 400, Slant::Upright);
        condensed.intrinsic.width = Width::CONDENSED;
        let mut condensed_bold = face("F", 900, Slant::Upright);
        condensed_bold.intrinsic.width = Width::CONDENSED;

        let base = IntrinsicStyle {
            width: Width::CONDENSED,
            ..IntrinsicStyle::NEUTRAL
        };
        let family = [face("F", 400, Slant::Upright), condensed_bold, condensed];
        assert_eq!(
            pick(&family, &want(Toggle::Absent, Toggle::Absent, base)),
            2,
            "the condensed Regular beats both the normal-width Regular and the condensed Black"
        );
    }

    /// The §6 tie-break: where a static face and a variable instance fit
    /// equally, the static one is chosen — a blanket rule that costs nothing
    /// in the common case (baking now recovers it) but still matters for the
    /// CFF2/WOFF boundary baking cannot cross.
    #[test]
    fn a_static_face_beats_an_equally_good_instance() {
        let mut instance = face("F", 600, Slant::Upright);
        instance.instance = Some(VariationInstance {
            subfamily: "SemiBold".into(),
            post_script_name: None,
            coords: vec![VariationCoord {
                axis: AxisTag::WEIGHT,
                value: 600.0,
            }],
        });
        let statik = face("F", 600, Slant::Upright);

        let base = IntrinsicStyle {
            weight: 600,
            ..IntrinsicStyle::NEUTRAL
        };
        // Instance first in the list, so only the tie-break can put the static
        // face ahead.
        assert_eq!(
            pick(
                &[instance.clone(), statik.clone()],
                &want(Toggle::Absent, Toggle::Absent, base)
            ),
            1
        );
        // …but a *better-fitting* instance still wins, because the tie-break
        // only applies to an actual tie.
        let mut closer = instance;
        closer.intrinsic.weight = 600;
        let far_static = face("F", 400, Slant::Upright);
        assert_eq!(
            pick(
                &[closer, far_static],
                &want(Toggle::Absent, Toggle::Absent, base)
            ),
            0
        );
    }

    #[test]
    fn ranking_is_stable_for_identical_candidates() {
        let family = [
            face("F", 400, Slant::Upright),
            face("F", 400, Slant::Upright),
        ];
        for _ in 0..8 {
            assert_eq!(
                pick(
                    &family,
                    &want(Toggle::Absent, Toggle::Absent, IntrinsicStyle::NEUTRAL)
                ),
                0
            );
        }
    }

    #[test]
    fn an_empty_candidate_list_ranks_to_nothing() {
        assert!(rank(
            &[],
            &want(Toggle::Absent, Toggle::Absent, IntrinsicStyle::NEUTRAL),
            ResolutionStep::SystemFamily
        )
        .is_none());
    }

    // ── step 6's parse ───────────────────────────────────────────────────

    #[test]
    fn a_face_qualified_name_parses_to_a_family_and_a_weight() {
        assert_eq!(parse_face_name("Segoe UI Light"), Some(("Segoe UI", 300)));
        assert_eq!(parse_face_name("Calibri Light"), Some(("Calibri", 300)));
        assert_eq!(parse_face_name("Arial Black"), Some(("Arial", 900)));
        assert_eq!(parse_face_name("Foo Extra Bold"), Some(("Foo", 800)));
        assert_eq!(parse_face_name("Inter SemiBold"), Some(("Inter", 600)));
    }

    /// A name with no trailing weight word is not face-qualified, and must not
    /// be chopped.
    #[test]
    fn a_plain_family_name_does_not_parse() {
        assert_eq!(parse_face_name("Times New Roman"), None);
        assert_eq!(parse_face_name("Highlight"), None);
        assert_eq!(parse_face_name("Calibri"), None);
    }

    /// Case is irrelevant to the weight word, as it is everywhere else.
    #[test]
    fn the_parsed_weight_word_is_case_insensitive() {
        assert_eq!(parse_face_name("Calibri LIGHT"), Some(("Calibri", 300)));
        assert_eq!(parse_face_name("Calibri semibold"), Some(("Calibri", 600)));
    }

    // ── name-kind gating ─────────────────────────────────────────────────

    /// The three predicates must partition the kinds they care about, or a name
    /// would be matchable at two different steps and the order would stop
    /// meaning anything.
    #[test]
    fn the_step_predicates_do_not_overlap() {
        let kinds = [
            NameKind::Table(NameId::Family),
            NameKind::Table(NameId::Subfamily),
            NameKind::Table(NameId::UniqueId),
            NameKind::Table(NameId::Full),
            NameKind::Table(NameId::PostScript),
            NameKind::Table(NameId::TypographicFamily),
            NameKind::Table(NameId::TypographicSubfamily),
            NameKind::Table(NameId::CompatibleFull),
            NameKind::Table(NameId::WwsFamily),
            NameKind::Table(NameId::WwsSubfamily),
            NameKind::ManagerFamily,
            NameKind::ManagerFace,
            NameKind::Instance,
            NameKind::ComposedStyle,
        ];
        for kind in kinds {
            let hits = u8::from(identifies_one_face(kind))
                + u8::from(identifies_a_family(kind))
                + u8::from(is_metadata_alias(kind));
            assert!(hits <= 1, "{kind:?} is accepted by {hits} steps");
        }
    }

    /// And between them they must cover every kind a face can actually be
    /// indexed under, or that name would be dead weight in the index.
    #[test]
    fn every_indexable_kind_is_reachable_from_some_step() {
        for kind in [
            NameKind::Table(NameId::Family),
            NameKind::Table(NameId::UniqueId),
            NameKind::Table(NameId::Full),
            NameKind::Table(NameId::PostScript),
            NameKind::Table(NameId::TypographicFamily),
            NameKind::Table(NameId::CompatibleFull),
            NameKind::Table(NameId::WwsFamily),
            NameKind::ManagerFamily,
            NameKind::ManagerFace,
            NameKind::Instance,
            NameKind::ComposedStyle,
        ] {
            assert!(
                identifies_one_face(kind) || identifies_a_family(kind) || is_metadata_alias(kind),
                "{kind:?} is indexed but no step will ever match it"
            );
        }
    }
}
