//! §17.15.1.10 automatic hyphenation: *where* a word may be cut (issue #158).
//!
//! The companion to [`super::segment`], which owns the neighbouring question
//! — where may a line break — and deliberately not part of it. UAX #14 gives
//! a break opportunity that exists in the text; a hyphenation opportunity is
//! one the engine *creates*, by inserting a glyph the author never typed, and
//! whether to take it depends on how much of the line is left (§17.15.1.53's
//! zone) — which only the line fitter knows. So the segmenter stays unaware
//! of hyphenation, this module answers the language question for a single
//! word, and `layout::line` decides, at the one moment a word is known not to
//! fit, whether and where to cut it.
//!
//! # Patterns, not Word's dictionary
//!
//! Word hyphenates from its own per-language dictionaries, which are not
//! published. This engine uses Knuth–Liang pattern hyphenation over the TeX
//! hyph-utf8 pattern sets, through the `hyphenation` crate, with only the
//! dictionaries `scripts/make_hyphenation_data.sh` vendors into `data/hyph/`.
//! The two agree on the great majority of words — both encode the same
//! orthographic rules — but not on every one, so "breaks where Word does" is
//! a property of the word, not of the engine, and the fixtures under
//! `test-files/auto-hyphenation-*.docx` are built from compounds whose seams
//! both must find. Word's minimum word length and left/right minima are also
//! its own; the minima applied here are the pattern sets' (`Language::minima`,
//! e.g. 2/3 for English, 2/2 for German and Russian).
//!
//! # Which tag reaches which dictionary
//!
//! [`HyphenLanguage::from_tag`] reduces a §17.3.2.20 BCP-47 tag to a vendored
//! dictionary by its **primary subtag**, with one region exception: `en-GB`
//! takes the British patterns, every other `en-*` the American ones. Every
//! `de-*` tag — `de-DE`, `de-AT`, `de-CH` — takes `de-1996`: the only Swiss
//! set the crate has is `de-ch-1901`, the *traditional* orthography, which
//! modern Swiss German does not use (it differs from `de-1996` chiefly in
//! writing `ss` for `ß`, and the reformed patterns handle `ss` words). A tag
//! this engine has no dictionary for, and a run with no tag at all, is
//! `None`: an untagged run is never hyphenated. Falling back to English for
//! an untagged document was considered and rejected — Word hyphenates in the
//! run's proofing language, and guessing one would cut German text at English
//! syllables.
//!
//! # No Cargo feature gate
//!
//! Always on, for the reason [`super`]'s module doc gives for ICU data: a
//! document that declares `w:autoHyphenation` is laid out wrongly without it,
//! and a flag would fork correct output behind a switch.

use std::sync::OnceLock;

use hyphenation::{Hyphenator, Language, Load, Standard};

/// A language this engine has hyphenation patterns for.
///
/// Its own ADT rather than the crate's `Language` (per the spec-faithful
/// convention) because the crate has 80 variants and this engine ships four
/// dictionaries: a value of this type is a promise that `dictionary()` can be
/// answered, which a bare `Language` is not.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum HyphenLanguage {
    /// hyph-utf8 `en-us` — every `en-*` tag but `en-GB`.
    EnUs,
    /// hyph-utf8 `en-gb`.
    EnGb,
    /// hyph-utf8 `de-1996` (reformed orthography) — every `de-*` tag.
    De1996,
    /// hyph-utf8 `ru`.
    Ru,
}

/// Every variant, in the order `DICTIONARIES` is indexed.
#[cfg(test)]
const ALL: [HyphenLanguage; 4] = [
    HyphenLanguage::EnUs,
    HyphenLanguage::EnGb,
    HyphenLanguage::De1996,
    HyphenLanguage::Ru,
];

/// One lazily loaded dictionary per vendored language. A process-wide
/// `static` rather than the `thread_local!` the ICU providers need:
/// `Standard` is an `fst` automaton plus a `HashMap`, `Send + Sync`, with no
/// `Rc` anywhere in it — confirmed by the compiler accepting this `static`.
static DICTIONARIES: [OnceLock<Standard>; 4] = [const { OnceLock::new() }; 4];

impl HyphenLanguage {
    /// §17.3.2.20: the vendored dictionary for a BCP-47 tag, or `None` for a
    /// language this engine has no patterns for (an untagged run is never
    /// hyphenated — see the module doc for why not English).
    pub fn from_tag(tag: &str) -> Option<Self> {
        let mut subtags = tag.split('-');
        let primary = subtags.next()?;
        let region = subtags.next();
        if primary.eq_ignore_ascii_case("en") {
            Some(if region.is_some_and(|r| r.eq_ignore_ascii_case("GB")) {
                Self::EnGb
            } else {
                Self::EnUs
            })
        } else if primary.eq_ignore_ascii_case("de") {
            Some(Self::De1996)
        } else if primary.eq_ignore_ascii_case("ru") {
            Some(Self::Ru)
        } else {
            None
        }
    }

    /// The crate's own identifier for this dictionary, which is also the
    /// stem of its file under `data/hyph/`.
    fn language(self) -> Language {
        match self {
            Self::EnUs => Language::EnglishUS,
            Self::EnGb => Language::EnglishGB,
            Self::De1996 => Language::German1996,
            Self::Ru => Language::Russian,
        }
    }

    /// The vendored `.standard.bincode`, copied from the pinned crate by
    /// `scripts/make_hyphenation_data.sh`.
    fn bytes(self) -> &'static [u8] {
        match self {
            Self::EnUs => include_bytes!("data/hyph/en-us.standard.bincode"),
            Self::EnGb => include_bytes!("data/hyph/en-gb.standard.bincode"),
            Self::De1996 => include_bytes!("data/hyph/de-1996.standard.bincode"),
            Self::Ru => include_bytes!("data/hyph/ru.standard.bincode"),
        }
    }

    /// This language's dictionary, deserialized on first use and kept for
    /// the life of the process.
    fn dictionary(self) -> &'static Standard {
        DICTIONARIES[self as usize].get_or_init(|| {
            Standard::from_reader(self.language(), &mut self.bytes()).expect(
                "src/i18n/data/hyph/*.standard.bincode is committed and regenerated \
                 by scripts/make_hyphenation_data.sh; a load failure here means the \
                 vendored dictionaries and the pinned `hyphenation` crate version \
                 have drifted",
            )
        })
    }
}

/// The byte offsets inside `word` at which the language's patterns allow a
/// hyphen, ascending, each a `char` boundary strictly inside the word.
///
/// `word` is a single run of letters — the caller
/// (`layout::line::hyphenation_core`) has already stripped punctuation and
/// refused tokens with digits, because what counts as a word is a
/// segmentation decision the patterns do not make. Case-insensitive, as the
/// crate is; the offsets are into the original, uncased string. A word
/// shorter than the language's left + right minima has no opportunities.
pub fn opportunities(word: &str, lang: HyphenLanguage) -> Vec<usize> {
    let mut breaks = lang.dictionary().hyphenate(word).breaks;
    // The crate already bounds its answer by the minima and by `char`
    // boundaries of the folded word, and realigns across the one case
    // folding that changes byte length (`İ`). This is the contract the
    // fitter slices by, restated as a filter rather than trusted, because a
    // cut at a non-boundary would panic inside layout.
    breaks.retain(|&at| at > 0 && at < word.len() && word.is_char_boundary(at));
    breaks.sort_unstable();
    breaks.dedup();
    breaks
}

#[cfg(test)]
mod tests {
    use super::*;
    use HyphenLanguage::*;

    /// The primary subtag decides, case-insensitively, with `en-GB` the one
    /// region that picks a different dictionary; `de-CH` is *not* one (the
    /// crate's only Swiss set is the 1901 orthography). Unknown, empty and
    /// malformed tags are `None`, which is "never hyphenate".
    #[test]
    fn tags_map_by_primary_subtag_with_the_region_exceptions() {
        assert_eq!(HyphenLanguage::from_tag("de-DE"), Some(De1996));
        assert_eq!(HyphenLanguage::from_tag("de-AT"), Some(De1996));
        assert_eq!(HyphenLanguage::from_tag("de-CH"), Some(De1996));
        assert_eq!(HyphenLanguage::from_tag("de"), Some(De1996));
        assert_eq!(HyphenLanguage::from_tag("en"), Some(EnUs));
        assert_eq!(HyphenLanguage::from_tag("en-US"), Some(EnUs));
        assert_eq!(HyphenLanguage::from_tag("EN-us"), Some(EnUs));
        assert_eq!(HyphenLanguage::from_tag("en-CA"), Some(EnUs));
        assert_eq!(HyphenLanguage::from_tag("en-GB"), Some(EnGb));
        assert_eq!(HyphenLanguage::from_tag("en-gb"), Some(EnGb));
        assert_eq!(HyphenLanguage::from_tag("ru-RU"), Some(Ru));
        assert_eq!(HyphenLanguage::from_tag("ru"), Some(Ru));
        assert_eq!(HyphenLanguage::from_tag("fr-FR"), None);
        assert_eq!(HyphenLanguage::from_tag("zz"), None);
        assert_eq!(HyphenLanguage::from_tag(""), None);
        assert_eq!(HyphenLanguage::from_tag("-"), None);
    }

    /// The crate's own documented examples, which pin that the vendored
    /// `en-us` dictionary is the one its docs describe.
    #[test]
    fn the_crates_documented_examples_hold() {
        assert_eq!(opportunities("hyphenation", EnUs), [2, 6, 7]);
        assert_eq!(opportunities("anfractuous", EnUs), [2, 6, 8]);
    }

    /// Patterns are matched case-insensitively and the offsets come back
    /// into the original string, so a sentence-initial word hyphenates
    /// exactly like its lowercase form.
    #[test]
    fn a_capitalised_word_hyphenates_like_its_lowercase() {
        assert!(!opportunities("hyphenation", EnUs).is_empty());
        assert_eq!(
            opportunities("Hyphenation", EnUs),
            opportunities("hyphenation", EnUs)
        );
        assert_eq!(
            opportunities("HYPHENATION", EnUs),
            opportunities("hyphenation", EnUs)
        );
    }

    /// `en-gb` is a different dictionary, not a relabelled `en-us`:
    /// `university` is in its exception list as `uni-ver-sity`.
    #[test]
    fn british_patterns_are_their_own_dictionary() {
        assert_eq!(opportunities("university", EnGb), [3, 6]);
    }

    /// A German compound must at least break at its morpheme seams — the cuts
    /// a reader expects, and the ones Word's dictionary makes too, which is
    /// what lets the fixture be compared against a Word render. The patterns
    /// may offer more (inside the components); these three must be among them.
    #[test]
    fn a_german_compound_breaks_at_its_seams() {
        let breaks = opportunities("Donaudampfschifffahrt", De1996);
        for (seam, at) in [("Donau|", 5), ("dampf|", 10), ("schiff|", 16)] {
            assert!(
                breaks.contains(&at),
                "{seam} (byte {at}) is missing from {breaks:?}"
            );
        }
    }

    /// Russian: two bytes per letter, so this also pins that offsets are
    /// bytes, not chars. `переносить` is `пе-ре-но-сить`.
    #[test]
    fn a_russian_word_breaks_between_its_syllables() {
        let breaks = opportunities("переносить", Ru);
        for (seam, at) in [("пе|", 4), ("пере|", 8), ("перено|", 12)] {
            assert!(
                breaks.contains(&at),
                "{seam} (byte {at}) is missing from {breaks:?}"
            );
        }
    }

    /// The pattern set's left/right minima (`Language::minima`) bound every
    /// opportunity, and a word shorter than their sum has none.
    #[test]
    fn minima_leave_the_word_ends_uncut() {
        for (lang, word) in [
            (EnUs, "hyphenation"),
            (EnGb, "hyphenation"),
            (De1996, "Donaudampfschifffahrt"),
            (Ru, "переносить"),
        ] {
            let (left, right) = lang.language().minima();
            let chars = word.chars().count();
            let breaks = opportunities(word, lang);
            assert!(!breaks.is_empty(), "{lang:?}: {word} should hyphenate");
            for at in breaks {
                let before = word[..at].chars().count();
                assert!(
                    before >= left,
                    "{lang:?}: {word} cut after {before} chars, minimum {left}"
                );
                assert!(
                    chars - before >= right,
                    "{lang:?}: {word} cut leaving {} chars, minimum {right}",
                    chars - before
                );
            }
        }
        assert!(opportunities("abcd", EnUs).is_empty(), "shorter than 2 + 3");
        assert!(opportunities("ab", De1996).is_empty(), "shorter than 2 + 2");
    }

    /// The contract the fitter relies on when it slices the word.
    #[test]
    fn offsets_are_ascending_char_boundaries_strictly_inside_the_word() {
        for (lang, word) in [
            (EnUs, "Hyphenation"),
            (De1996, "Straßenbahnhaltestelle"),
            (Ru, "Переносить"),
        ] {
            let breaks = opportunities(word, lang);
            assert!(!breaks.is_empty());
            assert!(
                breaks.windows(2).all(|w| w[0] < w[1]),
                "{breaks:?} not ascending"
            );
            for at in breaks {
                assert!(at > 0 && at < word.len(), "{at} is outside {word}");
                assert!(word.is_char_boundary(at), "{at} splits a char of {word}");
            }
        }
    }

    /// The blob/`expect` contract: each vendored file deserializes and claims
    /// the language this engine files it under. Checked by loading rather
    /// than by `from_tag`, so a stale `.bincode` fails here, loudly, instead
    /// of at the first German paragraph in production.
    #[test]
    fn every_vendored_dictionary_loads() {
        for lang in ALL {
            assert_eq!(lang.dictionary().language(), lang.language(), "{lang:?}");
        }
    }
}
