//! Host-OS color emoji typeface resolution.
//!
//! Given an optional requested emoji family (extracted from `<w:rFonts>` per
//! OOXML §17.3.2.26), find the best color emoji typeface available on the
//! host. The resolution chain is:
//!
//! 1. The requested family, if any.
//! 2. The host OS's native color emoji families, in priority order.
//!
//! When nothing matches we return [`EmojiTypeface::Unavailable`] with the
//! full attempted list so the caller can log loudly. We never bundle fonts
//! and never fall back to a non-emoji typeface — substituting (e.g.) Arial
//! for a missing color emoji font would produce visible junk, not a
//! degradation the user can act on.

use std::cell::RefCell;
use std::collections::HashMap;

use skia_safe::FontStyle;

use crate::render::fonts::{FontRegistry, TypefaceEntry};

// ─── ADTs ────────────────────────────────────────────────────────────────────

/// Closed enum of color emoji families we recognize. Adding a variant is a
/// code change — there is no free-form string path that lets arbitrary names
/// flow into the pipeline.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum EmojiFamily {
    /// Microsoft, ships with Windows. COLR/CPAL with outline base layer.
    SegoeUiEmoji,
    /// Apple, ships with macOS. sbix bitmap-only — no outline layer.
    AppleColorEmoji,
    /// Google, available on most Linux distros via `fonts-noto-color-emoji`.
    /// CBDT bitmap-only.
    NotoColorEmoji,
    /// Twitter (open-sourced as Twemoji), packaged on some Linux distros.
    /// COLR with outlines.
    TwitterColorEmoji,
    /// JoyPixels (formerly EmojiOne), proprietary set found on a few Linux
    /// hosts with the appropriate package installed.
    JoyPixels,
    /// Legacy EmojiOne Color, kept for backwards compatibility on hosts
    /// where the older package is still installed.
    EmojiOneColor,
}

impl EmojiFamily {
    /// The canonical OOXML / fontconfig family name for this emoji family.
    pub const fn family_name(self) -> &'static str {
        match self {
            EmojiFamily::SegoeUiEmoji => "Segoe UI Emoji",
            EmojiFamily::AppleColorEmoji => "Apple Color Emoji",
            EmojiFamily::NotoColorEmoji => "Noto Color Emoji",
            EmojiFamily::TwitterColorEmoji => "Twitter Color Emoji",
            EmojiFamily::JoyPixels => "JoyPixels",
            EmojiFamily::EmojiOneColor => "EmojiOne Color",
        }
    }

    /// Per-OS fallback chain consulted when the requested family is missing.
    /// Order is "most likely to be installed" descending.
    pub const fn host_default() -> &'static [EmojiFamily] {
        HOST_DEFAULT
    }

    /// Recognize a family name (case-insensitive) as a known emoji family.
    /// Returns `None` for any non-emoji name — the caller can then treat the
    /// run as ordinary text.
    pub fn from_name_ci(name: &str) -> Option<EmojiFamily> {
        match name.to_ascii_lowercase().as_str() {
            "segoe ui emoji" => Some(EmojiFamily::SegoeUiEmoji),
            "apple color emoji" => Some(EmojiFamily::AppleColorEmoji),
            "noto color emoji" => Some(EmojiFamily::NotoColorEmoji),
            "twitter color emoji" => Some(EmojiFamily::TwitterColorEmoji),
            "joypixels" => Some(EmojiFamily::JoyPixels),
            "emojione color" => Some(EmojiFamily::EmojiOneColor),
            _ => None,
        }
    }
}

#[cfg(target_os = "macos")]
const HOST_DEFAULT: &[EmojiFamily] = &[EmojiFamily::AppleColorEmoji];

#[cfg(target_os = "windows")]
const HOST_DEFAULT: &[EmojiFamily] = &[EmojiFamily::SegoeUiEmoji];

#[cfg(target_os = "linux")]
const HOST_DEFAULT: &[EmojiFamily] = &[
    EmojiFamily::NotoColorEmoji,
    EmojiFamily::TwitterColorEmoji,
    EmojiFamily::JoyPixels,
    EmojiFamily::EmojiOneColor,
];

#[cfg(not(any(target_os = "macos", target_os = "windows", target_os = "linux")))]
const HOST_DEFAULT: &[EmojiFamily] = &[];

/// Outcome of looking up a color emoji typeface on the host.
///
/// `Resolved` carries the actual [`TypefaceEntry`] from the registry — same
/// identity used by the rest of the renderer, so paint/cache lookups work.
/// `Unavailable` carries the full list of families we tried, so callers can
/// produce an actionable warning ("install fonts-noto-color-emoji").
#[derive(Clone, Debug)]
pub enum EmojiTypeface {
    Resolved {
        family: EmojiFamily,
        entry: TypefaceEntry,
    },
    Unavailable {
        attempted: Vec<EmojiFamily>,
    },
}

// ─── Test seam ───────────────────────────────────────────────────────────────

/// Lookup hook for `resolve`: given an [`EmojiFamily`], return the typeface
/// from the host (or `None` if absent). Production wires this to
/// [`FontRegistry::resolve_exact`] via [`RegistryLookup`]; tests can substitute
/// a deterministic mock.
pub trait EmojiTypefaceLookup {
    fn lookup(&self, family: EmojiFamily) -> Option<TypefaceEntry>;
}

/// Production lookup adapter — bridges [`EmojiTypefaceLookup`] to
/// [`FontRegistry::resolve_exact`]. Always queries `FontStyle::normal()`;
/// emoji typefaces are not weighted/italicized in OOXML.
pub struct RegistryLookup<'a> {
    pub registry: &'a FontRegistry,
}

impl EmojiTypefaceLookup for RegistryLookup<'_> {
    fn lookup(&self, family: EmojiFamily) -> Option<TypefaceEntry> {
        self.registry
            .resolve_exact(family.family_name(), FontStyle::normal())
    }
}

// ─── Resolution ──────────────────────────────────────────────────────────────

/// Stateless resolution. Returns either the first successful match or
/// `Unavailable` with the complete attempted list.
///
/// Order: requested family first (deduped against the host default chain),
/// then [`EmojiFamily::host_default`] in declaration order.
pub fn resolve(lookup: &impl EmojiTypefaceLookup, requested: Option<EmojiFamily>) -> EmojiTypeface {
    let mut attempted: Vec<EmojiFamily> = Vec::new();

    if let Some(family) = requested {
        attempted.push(family);
        if let Some(entry) = lookup.lookup(family) {
            return EmojiTypeface::Resolved { family, entry };
        }
    }

    for &family in EmojiFamily::host_default() {
        if attempted.contains(&family) {
            continue;
        }
        attempted.push(family);
        if let Some(entry) = lookup.lookup(family) {
            return EmojiTypeface::Resolved { family, entry };
        }
    }

    EmojiTypeface::Unavailable { attempted }
}

// ─── Cached resolver (per-render) ────────────────────────────────────────────

/// Per-render wrapper around an [`EmojiTypefaceLookup`] that dedupes
/// resolution by the requested family. Cheap to clone the resulting
/// [`EmojiTypeface`] (Skia typefaces are reference-counted internally).
///
/// Lifetime equals the host lookup's lifetime — typically the
/// [`FontRegistry`]'s, i.e. one render.
pub struct EmojiResolver<L: EmojiTypefaceLookup> {
    lookup: L,
    cache: RefCell<HashMap<Option<EmojiFamily>, EmojiTypeface>>,
}

impl<L: EmojiTypefaceLookup> EmojiResolver<L> {
    pub fn new(lookup: L) -> Self {
        Self {
            lookup,
            cache: RefCell::new(HashMap::new()),
        }
    }

    pub fn resolve(&self, requested: Option<EmojiFamily>) -> EmojiTypeface {
        if let Some(cached) = self.cache.borrow().get(&requested) {
            return cached.clone();
        }
        let result = resolve(&self.lookup, requested);
        self.cache.borrow_mut().insert(requested, result.clone());
        result
    }

    pub fn cached_count(&self) -> usize {
        self.cache.borrow().len()
    }
}

// ─── Tests (R1–R8 from docs/emoji-rendering.md) ──────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::{TypefaceId, TypefaceOrigin};
    use skia_safe::FontMgr;
    use std::cell::Cell;

    /// Build a real `TypefaceEntry` from any system default font. Used by
    /// tests that need the resolved variant to carry real Skia handles.
    fn any_typeface_entry() -> TypefaceEntry {
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

    /// Mock that succeeds only for the families in `available`, counting
    /// every lookup so caching tests can verify dedup.
    struct MockLookup {
        available: Vec<EmojiFamily>,
        calls: Cell<usize>,
    }

    impl MockLookup {
        fn new(available: Vec<EmojiFamily>) -> Self {
            Self {
                available,
                calls: Cell::new(0),
            }
        }
    }

    impl EmojiTypefaceLookup for MockLookup {
        fn lookup(&self, family: EmojiFamily) -> Option<TypefaceEntry> {
            self.calls.set(self.calls.get() + 1);
            if self.available.contains(&family) {
                Some(any_typeface_entry())
            } else {
                None
            }
        }
    }

    // ─── Real-host tests (platform-conditional) ──────────────────────────────

    /// R1 — macOS host with no requested family resolves to AppleColorEmoji.
    /// Apple Color Emoji is always present on macOS, so no runtime guard.
    #[cfg(target_os = "macos")]
    #[test]
    fn r1_macos_default_resolves_to_apple_color_emoji() {
        let registry = FontRegistry::new(FontMgr::new());
        let lookup = RegistryLookup {
            registry: &registry,
        };
        let result = resolve(&lookup, None);
        assert!(
            matches!(
                result,
                EmojiTypeface::Resolved {
                    family: EmojiFamily::AppleColorEmoji,
                    ..
                }
            ),
            "expected Resolved(AppleColorEmoji), got {result:?}"
        );
    }

    /// R2 — macOS host, requested SegoeUiEmoji (Windows family) is missing,
    /// falls through to AppleColorEmoji.
    #[cfg(target_os = "macos")]
    #[test]
    fn r2_macos_falls_back_when_requested_missing() {
        let registry = FontRegistry::new(FontMgr::new());
        let lookup = RegistryLookup {
            registry: &registry,
        };
        let result = resolve(&lookup, Some(EmojiFamily::SegoeUiEmoji));
        assert!(
            matches!(
                result,
                EmojiTypeface::Resolved {
                    family: EmojiFamily::AppleColorEmoji,
                    ..
                }
            ),
            "expected fallback to AppleColorEmoji, got {result:?}"
        );
    }

    /// R3 — Linux host with Noto Color Emoji installed resolves to
    /// NotoColorEmoji. Skipped if Noto is absent (no font bundling — we
    /// don't fail CI for hosts without the optional package).
    #[cfg(target_os = "linux")]
    #[test]
    fn r3_linux_noto_when_installed() {
        let registry = FontRegistry::new(FontMgr::new());
        let lookup = RegistryLookup {
            registry: &registry,
        };
        if lookup.lookup(EmojiFamily::NotoColorEmoji).is_none() {
            eprintln!("skipping R3: Noto Color Emoji not installed on this host");
            return;
        }
        let result = resolve(&lookup, None);
        assert!(
            matches!(
                result,
                EmojiTypeface::Resolved {
                    family: EmojiFamily::NotoColorEmoji,
                    ..
                }
            ),
            "expected Resolved(NotoColorEmoji), got {result:?}"
        );
    }

    // ─── Mock-host tests (deterministic, run on every platform) ──────────────

    /// R4 — Host with no emoji font: the result reports every family we
    /// attempted, in order, so callers can produce an actionable warning.
    #[test]
    fn r4_unavailable_reports_full_attempted_chain() {
        let mock = MockLookup::new(vec![]);
        let result = resolve(&mock, Some(EmojiFamily::SegoeUiEmoji));
        match result {
            EmojiTypeface::Unavailable { attempted } => {
                assert_eq!(attempted.first(), Some(&EmojiFamily::SegoeUiEmoji));
                let host_chain: &[EmojiFamily] = EmojiFamily::host_default();
                let host_set: std::collections::HashSet<_> = host_chain.iter().copied().collect();
                let attempted_set: std::collections::HashSet<_> =
                    attempted.iter().copied().collect();
                for fam in &host_set {
                    assert!(
                        attempted_set.contains(fam),
                        "host default chain member {fam:?} must appear in attempted list"
                    );
                }
            }
            other => panic!("expected Unavailable, got {other:?}"),
        }
    }

    /// R4b — Requested family in the host chain is not duplicated: order
    /// preserves "requested first, then host chain minus requested".
    #[test]
    fn r4b_attempted_list_dedupes_requested() {
        let mock = MockLookup::new(vec![]);
        let requested = EmojiFamily::host_default()
            .first()
            .copied()
            .unwrap_or(EmojiFamily::NotoColorEmoji);
        let result = resolve(&mock, Some(requested));
        match result {
            EmojiTypeface::Unavailable { attempted } => {
                let count = attempted.iter().filter(|&&f| f == requested).count();
                assert_eq!(count, 1, "requested family must appear exactly once");
            }
            _ => panic!("expected Unavailable"),
        }
    }

    /// R5 — Caching dedupes repeat lookups for the same requested family.
    #[test]
    fn r5_emoji_resolver_caches_repeated_calls() {
        let mock = MockLookup::new(vec![]);
        // Move ownership into the resolver so we can introspect call count
        // through a shared Cell.
        let calls_handle = std::rc::Rc::new(Cell::new(0usize));
        struct CountingLookup {
            calls: std::rc::Rc<Cell<usize>>,
        }
        impl EmojiTypefaceLookup for CountingLookup {
            fn lookup(&self, _family: EmojiFamily) -> Option<TypefaceEntry> {
                self.calls.set(self.calls.get() + 1);
                None
            }
        }
        drop(mock);

        let resolver = EmojiResolver::new(CountingLookup {
            calls: calls_handle.clone(),
        });
        // First call: every family in the resolution chain is queried.
        let _ = resolver.resolve(None);
        let after_first = calls_handle.get();
        assert!(
            after_first >= EmojiFamily::host_default().len(),
            "first call must query every host default family, got {after_first} calls"
        );
        // Second call: cache hit, no additional lookups.
        let _ = resolver.resolve(None);
        assert_eq!(
            calls_handle.get(),
            after_first,
            "EmojiResolver must dedupe repeat calls with the same `requested`"
        );
        assert_eq!(resolver.cached_count(), 1);
    }

    /// R5b — Different `requested` values are cached independently, each
    /// requiring its own underlying chain walk.
    #[test]
    fn r5b_distinct_requests_cache_independently() {
        let resolver = EmojiResolver::new(MockLookup::new(vec![]));
        let _ = resolver.resolve(None);
        let _ = resolver.resolve(Some(EmojiFamily::SegoeUiEmoji));
        let _ = resolver.resolve(Some(EmojiFamily::AppleColorEmoji));
        assert_eq!(resolver.cached_count(), 3);
    }

    /// R6 — Recognize the canonical Segoe UI Emoji name.
    #[test]
    fn r6_from_name_canonical() {
        assert_eq!(
            EmojiFamily::from_name_ci("Segoe UI Emoji"),
            Some(EmojiFamily::SegoeUiEmoji)
        );
    }

    /// R7 — Non-emoji names return None.
    #[test]
    fn r7_from_name_non_emoji_returns_none() {
        assert_eq!(EmojiFamily::from_name_ci("Calibri"), None);
        assert_eq!(EmojiFamily::from_name_ci(""), None);
        assert_eq!(EmojiFamily::from_name_ci("Arial Unicode MS"), None);
    }

    /// R8 — Recognition is case-insensitive (OOXML font names are CI).
    #[test]
    fn r8_from_name_case_insensitive() {
        assert_eq!(
            EmojiFamily::from_name_ci("segoe ui emoji"),
            Some(EmojiFamily::SegoeUiEmoji)
        );
        assert_eq!(
            EmojiFamily::from_name_ci("APPLE COLOR EMOJI"),
            Some(EmojiFamily::AppleColorEmoji)
        );
        assert_eq!(
            EmojiFamily::from_name_ci("NoTo CoLoR eMoJi"),
            Some(EmojiFamily::NotoColorEmoji)
        );
    }

    // ─── Round-trip and shape invariants ────────────────────────────────────

    /// Every family round-trips through `family_name` → `from_name_ci`.
    #[test]
    fn family_name_round_trip() {
        for fam in [
            EmojiFamily::SegoeUiEmoji,
            EmojiFamily::AppleColorEmoji,
            EmojiFamily::NotoColorEmoji,
            EmojiFamily::TwitterColorEmoji,
            EmojiFamily::JoyPixels,
            EmojiFamily::EmojiOneColor,
        ] {
            assert_eq!(EmojiFamily::from_name_ci(fam.family_name()), Some(fam));
        }
    }

    /// Resolution with `requested = Some(host_default[0])` and the host
    /// providing it must short-circuit on the first lookup.
    #[test]
    fn requested_match_short_circuits_chain() {
        let target = EmojiFamily::host_default()
            .first()
            .copied()
            .unwrap_or(EmojiFamily::NotoColorEmoji);
        let mock = MockLookup::new(vec![target]);
        let result = resolve(&mock, Some(target));
        assert!(matches!(
            result,
            EmojiTypeface::Resolved { family, .. } if family == target
        ));
        assert_eq!(
            mock.calls.get(),
            1,
            "must not query host chain after requested matches"
        );
    }
}
