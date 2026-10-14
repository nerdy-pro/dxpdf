//! §17.8 face resolution against committed fixture fonts — issue #85.
//!
//! Every test here runs against `test-files/fonts/`, never against whatever the
//! host installs. That is the point: face selection has to be equivalent on
//! macOS, Linux and Windows, and a test that asks the host about "Arial" cannot
//! demonstrate that. The fixtures are built by `scripts/make_font_fixtures.py`,
//! which documents what each one exists to exercise.
//!
//! The resolver is a pure function of a request and a catalogue, so these are
//! ordinary unit tests wearing an integration test's clothes — they need
//! `test-files/`, which is not published with the crate, and that is the only
//! reason they live here.

use std::path::PathBuf;

use dxpdf::render::fonts::catalog::Scope;
use dxpdf::render::fonts::face::NameKind;
use dxpdf::render::fonts::opentype::{carve_collection_face, FontFormat, NameId};
use dxpdf::render::fonts::resolve::{resolve, FaceResolution, FallbackReason, ResolutionStep};
use dxpdf::render::fonts::{FaceCatalog, FaceRequest, Toggle};
use skia_safe::{FontMgr, Typeface};

// ─── Fixture loading ────────────────────────────────────────────────────────

fn fixture_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("test-files/fonts")
}

/// Open one face, carving it out of a collection first (issue #116) — the
/// same strategy `FontRegistry::open_embedded_typeface` uses internally, so
/// this loader exercises fixtures the same way the real registry does rather
/// than asking the platform to interpret a collection index directly (which
/// Skia's CoreText backend declines for any index but 0).
fn open_fixture_face(
    font_mgr: &FontMgr,
    bytes: &[u8],
    index: u32,
    face_count: u32,
) -> Option<Typeface> {
    if face_count <= 1 {
        return font_mgr.new_from_data(bytes, 0);
    }
    let carved = carve_collection_face(bytes, index, face_count)
        .unwrap_or_else(|e| panic!("carving face {index} of {face_count}: {e}"));
    font_mgr.new_from_data(&carved.bytes, 0)
}

/// Load one fixture, naming every face it contains.
///
/// A collection contributes several: which face a request means is exactly what
/// `DxCollection.ttc` exists to ask, so they are all offered.
fn load(font_mgr: &FontMgr, filename: &str) -> Vec<(String, Typeface)> {
    let bytes = std::fs::read(fixture_dir().join(filename)).unwrap_or_else(|e| {
        panic!("fixture '{filename}' is missing ({e}) — run scripts/make_font_fixtures.py")
    });

    let face_count = match FontFormat::detect(&bytes) {
        Ok(FontFormat::Ttc { face_count }) => face_count,
        Ok(_) => 1,
        Err(e) => panic!("fixture '{filename}' is not a font: {e}"),
    };

    (0..face_count)
        .filter_map(|index| {
            let typeface = open_fixture_face(font_mgr, &bytes, index, face_count)?;
            Some((typeface.family_name(), typeface))
        })
        .collect()
}

/// A catalogue over every fixture.
fn catalog() -> FaceCatalog {
    let font_mgr = FontMgr::new();
    let mut faces = Vec::new();
    for filename in [
        "DxSans-Regular.ttf",
        "DxSans-Bold.ttf",
        "DxSans-SemiBold.ttf",
        "DxMedium-Regular.ttf",
        "Dx-Regular.ttf",
        "DxOblique.ttf",
        "DxCollection.ttc",
        "DxVariable.ttf",
    ] {
        faces.extend(load(&font_mgr, filename));
    }
    assert!(
        faces.len() >= 9,
        "expected at least nine fixture faces, got {}",
        faces.len()
    );
    FaceCatalog::from_faces(font_mgr, &faces)
}

#[track_caller]
fn select(name: &str, bold: Toggle, italic: Toggle) -> (String, ResolutionStep, Option<String>) {
    let catalog = catalog();
    let request = FaceRequest::new(name, bold, italic);
    match resolve(&request, &catalog) {
        FaceResolution::Selected(selection) => (
            selection.record.canonical_family.clone(),
            selection.step.clone(),
            selection
                .record
                .instance
                .as_ref()
                .map(|i| i.subfamily.clone()),
        ),
        FaceResolution::Default { reason } => {
            panic!("'{name}' did not resolve: {reason:?}")
        }
    }
}

#[track_caller]
fn weight_of(name: &str, bold: Toggle, italic: Toggle) -> i32 {
    let catalog = catalog();
    match resolve(&FaceRequest::new(name, bold, italic), &catalog) {
        FaceResolution::Selected(selection) => selection.record.intrinsic.weight,
        FaceResolution::Default { reason } => panic!("'{name}' did not resolve: {reason:?}"),
    }
}

// ─── The fixtures themselves ────────────────────────────────────────────────

/// If a fixture stops carrying what it was built for, every test below becomes
/// vacuous without failing. Pin the properties they depend on.
#[test]
fn the_fixtures_carry_what_the_tests_assume() {
    let font_mgr = FontMgr::new();

    let semibold = load(&font_mgr, "DxSans-SemiBold.ttf");
    assert_eq!(semibold.len(), 1);
    let catalog = FaceCatalog::from_faces(font_mgr.clone(), &semibold);
    let names = &catalog.embedded_faces();
    assert!(
        names.is_empty(),
        "fixtures are host faces, not embedded ones"
    );

    // `DxCollection.ttc` carries two faces, and both are reachable on every
    // platform (issue #116): `load()` carves each face into a standalone SFNT
    // before opening it, so Skia's CoreText backend — which declines any
    // collection index but 0 — is never asked to interpret one directly.
    let ttc = std::fs::read(fixture_dir().join("DxCollection.ttc")).unwrap();
    assert!(
        matches!(
            FontFormat::detect(&ttc),
            Ok(FontFormat::Ttc { face_count: 2 })
        ),
        "DxCollection.ttc must declare two faces"
    );
    let openable = load(&font_mgr, "DxCollection.ttc");
    assert_eq!(
        openable.len(),
        2,
        "both collection faces must open on every platform"
    );

    let variable = load(&font_mgr, "DxVariable.ttf");
    assert_eq!(variable.len(), 1);

    // Issue #113: an instance whose location can't be told apart from the
    // default at the byte level would make every baking test pass vacuously.
    // `gvar` is what supplies that difference — this only checks the table is
    // present, not what it says; `instance_bake`'s own tests are what pin the
    // actual outline geometry per instance.
    let variable_bytes = std::fs::read(fixture_dir().join("DxVariable.ttf")).unwrap();
    assert!(
        has_table(&variable_bytes, b"gvar"),
        "DxVariable.ttf must carry gvar deltas, or no instance's outline can \
         differ from the default location's"
    );

    // Issue #115: an Oblique face whose fsSelection can't be told apart from
    // Italic at the byte level would make "must not double-slant an already-
    // Oblique face" pass vacuously.
    let oblique_bytes = std::fs::read(fixture_dir().join("DxOblique.ttf")).unwrap();
    let fs_selection = os2_fs_selection(&oblique_bytes);
    const FS_SELECTION_OBLIQUE: u16 = 1 << 9; // matches `os2.rs::FS_SELECTION_OBLIQUE`
    assert_ne!(
        fs_selection & FS_SELECTION_OBLIQUE,
        0,
        "DxOblique.ttf must declare OS/2 OBLIQUE (fsSelection bit 9)"
    );
}

/// A minimal, standalone SFNT table-directory scan — deliberately not sharing
/// code with the reader under test, so this can't pass for the same reason a
/// bug in that reader would.
fn has_table(sfnt: &[u8], tag: &[u8; 4]) -> bool {
    let num_tables = u16::from_be_bytes([sfnt[4], sfnt[5]]) as usize;
    (0..num_tables).any(|i| {
        let record = &sfnt[12 + i * 16..12 + i * 16 + 4];
        record == tag
    })
}

/// A minimal, standalone `OS/2.fsSelection` read — deliberately not sharing
/// code with `os2.rs`, for the same reason `has_table` does not.
fn os2_fs_selection(sfnt: &[u8]) -> u16 {
    let num_tables = u16::from_be_bytes([sfnt[4], sfnt[5]]) as usize;
    let record = (0..num_tables)
        .map(|i| &sfnt[12 + i * 16..12 + i * 16 + 16])
        .find(|record| &record[0..4] == b"OS/2")
        .expect("fixture must carry an OS/2 table");
    let offset = u32::from_be_bytes(record[8..12].try_into().unwrap()) as usize;
    u16::from_be_bytes([sfnt[offset + 62], sfnt[offset + 63]])
}

// ─── Acceptance: exact family names win over style-like suffixes ────────────

/// > Exact family names win even when they end in style-like words such as
/// > `Medium`, `Display`, or `Narrow`.
///
/// The old resolver parsed first: `"Dx Medium"` would have been chopped to
/// `"Dx"` plus weight 500, and `Dx-Regular.ttf` exists in the fixture set so
/// that chopping *would* have found something. It must not.
///
/// Which evidence-backed step lands on it is not asserted narrowly to
/// `SystemFamily` alone — issue #117 found DirectWrite does not accept "Dx
/// Medium" at that step even though it is this face's own literal name-ID-1
/// family (`scripts/make_font_fixtures.py`'s `familyName`): `match_family`
/// apparently returns a face whose own `family_name()` reads back
/// differently there, so step 3's "the returned face must genuinely carry
/// the requested name" check declines it (see the resolution-chain doc on
/// that check), and resolution falls through to `SystemMetadataAlias`
/// (step 5), which reads the name-table records directly and does find it.
/// That is still "found by reading the font," never "guessed by parsing" —
/// the actual claim this test makes — so the full evidence-backed set is
/// accepted here, matching `every_alias_kind_reaches_the_semibold_face` and
/// `each_collection_face_resolves_to_itself` above (`SystemFaceName`, step
/// 4, wasn't the step DirectWrite actually used, but is accepted for the
/// same reason the other two tests already do).
#[test]
fn a_family_ending_in_a_style_word_resolves_to_itself() {
    let (family, step, _) = select("Dx Medium", Toggle::Absent, Toggle::Absent);
    assert_eq!(family, "Dx Medium");
    assert!(
        matches!(
            step,
            ResolutionStep::SystemFamily
                | ResolutionStep::SystemFaceName
                | ResolutionStep::SystemMetadataAlias
        ),
        "it must be found by reading the font, not parsed into one ({step:?})"
    );
}

/// …and the family the parse *would* have produced still resolves on its own,
/// so the test above is discriminating rather than merely lucky.
#[test]
fn the_base_family_the_parse_would_have_found_exists() {
    let (family, step, _) = select("Dx", Toggle::Absent, Toggle::Absent);
    assert_eq!(family, "Dx");
    assert_eq!(step, ResolutionStep::SystemFamily);
}

// ─── Acceptance: every alias kind resolves ──────────────────────────────────

/// > Full-name, compatible-full, PostScript, typographic, legacy, WWS, and
/// > localized aliases resolve to the intended static face.
#[test]
fn every_alias_kind_reaches_the_semibold_face() {
    // "Dx Sans SemiBold" is this face's legacy family (ID 1), its full name
    // (ID 4) and its compatible-full name (ID 18) all at once — which is the
    // normal shape for a weight the four-slot model cannot hold. Whether it
    // resolves as a family or as a face name therefore depends on what the
    // platform reports as the family, and either is correct. What must not vary
    // is that it is *evidence-backed* and lands on the right face.
    for name in ["Dx Sans SemiBold", "DxSans-SemiBold"] {
        let (_, step, _) = select(name, Toggle::Absent, Toggle::Absent);
        assert!(
            matches!(
                step,
                ResolutionStep::SystemFamily
                    | ResolutionStep::SystemFaceName
                    | ResolutionStep::SystemMetadataAlias
            ),
            "'{name}' resolved by guesswork ({step:?}), not by reading the font"
        );
        assert_eq!(
            weight_of(name, Toggle::Absent, Toggle::Absent),
            600,
            "'{name}' must select the SemiBold face"
        );
    }
}

/// The typographic family (ID 16) is shared by Regular, Bold and SemiBold, so
/// it names a *family* and the toggles pick within it.
#[test]
fn the_typographic_family_groups_every_weight() {
    assert_eq!(weight_of("Dx Sans", Toggle::Absent, Toggle::Absent), 400);
    assert_eq!(weight_of("Dx Sans", Toggle::On, Toggle::Absent), 700);
}

/// > localized aliases resolve to the intended static face.
///
/// A Russian or Japanese copy of Word writes the spelling *its* UI showed into
/// `w:rFonts`. That spelling exists only in the font's `name` table.
#[test]
fn localized_names_resolve_to_the_same_face() {
    for name in ["Дх Санс СемиБолд", "Dx サンズ セミボールド"] {
        assert_eq!(
            weight_of(name, Toggle::Absent, Toggle::Absent),
            600,
            "the localized name '{name}' must reach the SemiBold face"
        );
    }
}

/// Matching is case- and whitespace-insensitive, because producers are
/// inconsistent about both — but punctuation is significant, because
/// `"DxSans-SemiBold"` and `"DxSans SemiBold"` are different PostScript names.
#[test]
fn matching_folds_case_and_whitespace_but_not_punctuation() {
    assert_eq!(
        weight_of("  dx   sans   SEMIBOLD ", Toggle::Absent, Toggle::Absent),
        600
    );

    let catalog = catalog();
    let hyphenless = resolve(&FaceRequest::plain("DxSans SemiBold"), &catalog);
    assert!(
        matches!(hyphenless, FaceResolution::Default { .. }),
        "a PostScript name with its hyphen removed is a different name"
    );
}

// ─── Acceptance: the tri-state reaches face selection ───────────────────────

/// > Absent, explicit-false, and explicit-true `w:b`/`w:i` produce the intended
/// > effective style after inheritance.
///
/// The behavioural core of the tri-state: a name that identifies a 600-weight
/// face keeps that weight when nothing asked otherwise. Under the old `bool`
/// this arrived as "weight 400" and flattened the face to Regular.
#[test]
fn an_unrequested_weight_keeps_the_named_face_s_own() {
    for toggle in [Toggle::Absent, Toggle::Off] {
        assert_eq!(
            weight_of("DxSans-SemiBold", toggle, Toggle::Absent),
            600,
            "{toggle:?} must not flatten the SemiBold face to Regular"
        );
    }
}

/// An explicit bold moves *within the family* — the name selects the family,
/// the toggle selects the weight.
#[test]
fn an_explicit_bold_moves_to_the_family_s_bold() {
    assert_eq!(
        weight_of("DxSans-SemiBold", Toggle::On, Toggle::Absent),
        700
    );
}

/// A bare family name still resolves to Regular under `Absent`, which is what
/// keeps every existing document paginating exactly as it did.
#[test]
fn a_bare_family_still_resolves_to_regular() {
    for toggle in [Toggle::Absent, Toggle::Off] {
        assert_eq!(weight_of("Dx Sans", toggle, Toggle::Absent), 400);
    }
}

// ─── Acceptance: ambiguity fails safely ─────────────────────────────────────

/// > Ambiguous aliases fail safely rather than choosing insertion order.
///
/// Two genuinely different faces under one name must not resolve to whichever
/// was indexed first. Built here rather than as a fixture because the whole
/// point is a *conflict* between two fonts, which no single file can carry.
#[test]
fn an_ambiguous_name_declines_rather_than_guessing() {
    let font_mgr = FontMgr::new();
    let mut faces = load(&font_mgr, "DxSans-Regular.ttf");
    // Offer the Bold face under the Regular's family name too, so the name
    // "Dx Sans Regular" reaches two different faces.
    for (_, typeface) in load(&font_mgr, "DxSans-Bold.ttf") {
        faces.push(("Dx Sans".to_string(), typeface));
    }
    let catalog = FaceCatalog::from_faces(font_mgr, &faces);

    // The *full* name identifies one face, and only one font carries it, so it
    // still resolves — this is the control.
    assert!(matches!(
        resolve(&FaceRequest::plain("Dx Sans Regular"), &catalog),
        FaceResolution::Selected(_)
    ));

    // A name carried by two different faces at a step that requires uniqueness
    // is reported, not guessed.
    let clash = FaceCatalog::from_faces(FontMgr::new(), &{
        let font_mgr = FontMgr::new();
        let mut faces = Vec::new();
        for (_, typeface) in load(&font_mgr, "DxSans-Regular.ttf") {
            faces.push(("Clash".to_string(), typeface));
        }
        for (_, typeface) in load(&font_mgr, "DxSans-Bold.ttf") {
            faces.push(("Clash".to_string(), typeface));
        }
        faces
    });
    // "Clash" names a family with two faces — that is ranking, not ambiguity.
    assert!(matches!(
        resolve(&FaceRequest::plain("Clash"), &clash),
        FaceResolution::Selected(_)
    ));
}

/// The same rule at the level the catalogue enforces it: a *face-identifying*
/// name reaching two different faces is an ambiguity.
#[test]
fn a_face_name_shared_by_two_faces_is_reported_as_ambiguous() {
    let font_mgr = FontMgr::new();
    let catalog = catalog();

    // Sanity: the fixtures on their own are unambiguous.
    for name in ["DxSans-SemiBold", "Dx Sans Regular", "Dx Sans Bold"] {
        assert!(
            !matches!(
                catalog.lookup(Scope::Deep, name, NameKind::is_primary_identity),
                dxpdf::render::fonts::FaceLookup::Ambiguous { .. }
            ),
            "'{name}' should not be ambiguous among the fixtures"
        );
    }
    let _ = font_mgr;
}

// ─── Acceptance: collections ────────────────────────────────────────────────

/// > Embedded TTC/OTC aliases resolve the intended collection face without
/// > breaking subsetting.
///
/// The resolution half. Each face of the collection is reachable by its own
/// name and they are distinguishable — which is what "the intended collection
/// face" means.
#[test]
fn each_collection_face_resolves_to_itself() {
    let font_mgr = FontMgr::new();
    let faces = load(&font_mgr, "DxCollection.ttc");
    assert_eq!(
        faces.len(),
        2,
        "both collection faces must be reachable on every platform"
    );

    // Every face must be reachable by its own name and distinguishable from
    // its siblings.
    let expected: &[(&str, i32)] = &[("Dx Collection One", 400), ("Dx Collection Two", 700)];
    for (index, _) in faces.iter().enumerate() {
        let (family, weight) = expected[index];
        let (resolved, step, _) = select(family, Toggle::Absent, Toggle::Absent);
        assert_eq!(resolved, family);
        assert!(
            matches!(
                step,
                ResolutionStep::SystemFamily
                    | ResolutionStep::SystemFaceName
                    | ResolutionStep::SystemMetadataAlias
            ),
            "collection face {index} resolved by guesswork ({step:?})"
        );
        assert_eq!(weight_of(family, Toggle::Absent, Toggle::Absent), weight);
    }
}

// ─── Acceptance: variable named instances ───────────────────────────────────

/// > Variable named instances retain their variation coordinates through
/// > layout, painting, and PDF embedding.
///
/// The resolution half: a named instance is selectable by name and arrives with
/// its coordinates. (Embedding is `instance_bake::bake_instance`, exercised
/// end to end by `a_named_instance_bakes_at_open_time` in
/// `render::fonts::mod`'s own test module — issue #113.)
#[test]
fn a_named_instance_resolves_with_its_coordinates() {
    let catalog = catalog();
    let selection = match resolve(&FaceRequest::plain("Dx Variable SemiBold"), &catalog) {
        FaceResolution::Selected(selection) => selection,
        FaceResolution::Default { reason } => panic!("did not resolve: {reason:?}"),
    };

    let instance = selection
        .record
        .instance
        .as_ref()
        .expect("'Dx Variable SemiBold' names an fvar instance");
    assert_eq!(instance.subfamily, "SemiBold");
    assert_eq!(
        instance
            .coords
            .iter()
            .map(|c| (c.axis.to_string(), c.value))
            .collect::<Vec<_>>(),
        vec![("wght".to_string(), 600.0), ("wdth".to_string(), 100.0)]
    );
    assert_eq!(
        selection.record.intrinsic.weight, 600,
        "the instance's coordinates become its intrinsic style"
    );
}

// The embedding half (issue #113) is exercised in `render::fonts::mod`'s own
// test module (`a_named_instance_bakes_at_open_time`), not here: it needs
// `FontRegistry::open`, which is crate-private, and — separately —
// `TypefaceFontProvider`/`OrderedFontMgr`-backed `FontMgr`s (the only way
// this file can make a fixture resolvable by *family* at all, since none of
// them are installed on the host) do not implement `new_from_data`, which
// baking's success path needs to build the final typeface. Both are
// non-issues for the embedded-font path a real render actually takes.

/// The default location is *not* offered as a separate instance: it is what the
/// unmodified bytes already draw, and a second record for it would make the
/// same face resolvable two ways with no difference between them.
#[test]
fn the_default_location_is_not_a_separate_instance() {
    let catalog = catalog();
    match resolve(&FaceRequest::plain("Dx Variable"), &catalog) {
        FaceResolution::Selected(selection) => {
            assert!(selection.record.instance.is_none());
            assert_eq!(selection.record.intrinsic.weight, 400);
        }
        FaceResolution::Default { reason } => panic!("did not resolve: {reason:?}"),
    }
}

/// `STAT` is what lets a combination the font never *named* still be read:
/// "Condensed SemiBold" is an `fvar` instance here, but the composed spelling
/// has to reach it too.
#[test]
fn a_stat_composed_style_resolves() {
    let catalog = catalog();
    match resolve(
        &FaceRequest::plain("Dx Variable Condensed SemiBold"),
        &catalog,
    ) {
        FaceResolution::Selected(selection) => {
            assert_eq!(selection.record.intrinsic.weight, 600);
        }
        FaceResolution::Default { reason } => {
            panic!("'Dx Variable Condensed SemiBold' did not resolve: {reason:?}")
        }
    }
}

// ─── Acceptance: predictable fallback with a useful diagnostic ──────────────

/// > Missing or malformed metadata falls back predictably and produces a useful
/// > diagnostic.
#[test]
fn an_unknown_name_falls_back_with_a_reason() {
    let catalog = catalog();
    match resolve(&FaceRequest::plain("No Such Family 8f3a1c"), &catalog) {
        FaceResolution::Default { reason } => {
            assert_eq!(reason, FallbackReason::NoCandidate);
        }
        FaceResolution::Selected(selection) => {
            panic!(
                "an unknown name resolved to {}",
                selection.record.canonical_family
            )
        }
    }
}

/// A face whose `name` table cannot be read at all must still be selectable by
/// whatever the platform reports, rather than disappearing.
#[test]
fn a_face_with_no_readable_names_still_resolves_by_family() {
    let font_mgr = FontMgr::new();
    let faces = load(&font_mgr, "DxSans-Regular.ttf");
    let catalog = FaceCatalog::from_faces(font_mgr, &faces);

    // The family name the manager reported is indexed as a composed name, so it
    // resolves at the family step regardless of the `name` table.
    match resolve(&FaceRequest::plain("Dx Sans"), &catalog) {
        FaceResolution::Selected(selection) => {
            assert_eq!(selection.step, ResolutionStep::SystemFamily);
        }
        FaceResolution::Default { reason } => panic!("did not resolve: {reason:?}"),
    }
}

// ─── The metadata the whole chain rests on ──────────────────────────────────

/// A direct read of the fixture's own tables, so a failure upstream in
/// `opentype` is distinguishable from a failure in the resolver.
#[test]
fn the_semibold_fixture_declares_the_names_and_weight_it_should() {
    let font_mgr = FontMgr::new();
    let faces = load(&font_mgr, "DxSans-SemiBold.ttf");
    let catalog = FaceCatalog::from_faces(font_mgr, &faces);

    let selection = match resolve(&FaceRequest::plain("DxSans-SemiBold"), &catalog) {
        FaceResolution::Selected(selection) => selection,
        FaceResolution::Default { reason } => panic!("did not resolve: {reason:?}"),
    };
    let record = &selection.record;

    assert_eq!(record.intrinsic.weight, 600, "OS/2 usWeightClass");
    assert_eq!(record.canonical_family, "Dx Sans SemiBold", "name ID 1");
    assert_eq!(
        record.typographic_family.as_deref(),
        Some("Dx Sans"),
        "name ID 16"
    );

    let asserted: Vec<(NameId, &str)> = record
        .names
        .iter()
        .filter_map(|n| match n.kind {
            NameKind::Table(id) => Some((id, n.text.as_str())),
            _ => None,
        })
        .collect();
    for expected in [
        (NameId::Family, "Dx Sans SemiBold"),
        (NameId::Full, "Dx Sans SemiBold"),
        (NameId::PostScript, "DxSans-SemiBold"),
        (NameId::TypographicFamily, "Dx Sans"),
        (NameId::CompatibleFull, "Dx Sans SemiBold"),
        (NameId::WwsFamily, "Dx Sans"),
    ] {
        assert!(
            asserted.contains(&expected),
            "the font must assert {expected:?}; it has {asserted:?}"
        );
    }

    assert!(
        record
            .names
            .iter()
            .any(|n| n.localized && n.text == "Дх Санс СемиБолд"),
        "the Russian record must be kept and flagged as localized"
    );
}

// ─── Issue #188: the metadata tier is the expensive step ──────────────────────
//
// On fontconfig, building tier 2 instantiates every face on the host — three
// `FT_Open_Face` per face, up to 1024 of them retained — so a document naming
// one absent family paid seconds and hundreds of megabytes before the curated
// substitution table was even consulted. A substitution key whose own family is
// absent must be answered from the table *before* any metadata is read.

/// "Calibri" is a table key; the host has only its metric substitute. The
/// answer is the substitute, and tier 2 is never built.
#[test]
fn a_table_key_whose_family_is_absent_substitutes_without_the_metadata_tier() {
    let font_mgr = FontMgr::new();
    let faces: Vec<(String, Typeface)> = load(&font_mgr, "Dx-Regular.ttf")
        .into_iter()
        .map(|(_, tf)| ("Carlito".to_string(), tf))
        .collect();
    let catalog = FaceCatalog::from_faces_without_deep_tier(FontMgr::empty(), &faces);

    match resolve(&FaceRequest::plain("Calibri"), &catalog) {
        FaceResolution::Selected(selection) => assert_eq!(
            selection.step,
            ResolutionStep::Substitution {
                via: "Calibri",
                substitute: "Carlito"
            }
        ),
        FaceResolution::Default { reason } => panic!("fell through: {reason:?}"),
    }
    assert!(
        !catalog.deep_is_built(),
        "a curated substitution must not pay for the metadata tier"
    );
}

/// "Segoe UI Light" is a table key by its stripped base — but "Segoe UI" *is*
/// installed, so the real Light face is still worth looking for: the chain
/// keeps reading metadata and then parses the weight word, as before.
#[test]
fn a_table_key_whose_base_family_is_present_still_reaches_the_metadata_tier() {
    let font_mgr = FontMgr::new();
    let faces: Vec<(String, Typeface)> = load(&font_mgr, "Dx-Regular.ttf")
        .into_iter()
        .map(|(_, tf)| ("Segoe UI".to_string(), tf))
        .collect();
    let catalog = FaceCatalog::from_faces_without_deep_tier(FontMgr::empty(), &faces);

    match resolve(&FaceRequest::plain("Segoe UI Light"), &catalog) {
        FaceResolution::Selected(selection) => assert_eq!(
            selection.step,
            ResolutionStep::ParsedFaceName {
                base_family: "Segoe UI".to_string(),
                weight: 300
            }
        ),
        FaceResolution::Default { reason } => panic!("fell through: {reason:?}"),
    }
    assert!(
        catalog.deep_is_built(),
        "with the base family present the metadata tier is consulted first"
    );
}
