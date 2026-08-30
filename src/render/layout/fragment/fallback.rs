//! Per-glyph font fallback — issue #139.
//!
//! The third pass over a paragraph's finished fragment vector, and the same
//! division `super::bidi` and `super::shape` keep: this module knows
//! *which* text its resolved face cannot draw and *what to do about it*, and
//! the [`FallbackLookup`] it is handed knows *how* to ask the host.
//!
//! # The problem
//!
//! [`crate::render::fonts::resolve`] is a pure function of a **name plus two
//! §17.7.2 toggles**. All eight of its steps answer "which face did the
//! document ask for", and none of them can consider which codepoints the run
//! actually needs — the request does not carry them. So `Times New Roman`
//! resolves, correctly, for a run containing `ア`, and the paint side then maps
//! that codepoint through the face's cmap, gets glyph 0, and Skia's PDF backend
//! emits nothing at all. Not a `.notdef` box — nothing. The run is measured and
//! positioned, and the space it was given sits empty.
//!
//! Real Word documents mostly dodge this, because §17.9.3 gives every numbering
//! level its own `w:rPr` and Word writes a covering font into it. The gap is
//! content that names no covering font, which is any hand-authored or
//! non-Word-generated DOCX.
//!
//! # Why a name, and not a resolved face
//!
//! A fallback has to survive two independent re-resolutions after this pass:
//! the painter's, and [`crate::render::subset::collect`]'s. Both re-resolve
//! from `DrawCommand::Text`'s `font_family` + toggles, and neither carries a
//! typeface id — so the only thing that reaches both is a **name**.
//!
//! That is sound because [`crate::render::fonts::FontRegistry::resolve`]
//! consults its request-keyed cache first, and the lookup below pins the chosen
//! face into that cache under the name it reports. The name is therefore
//! authoritative by construction rather than by hoping resolution agrees — it
//! has to be, because a host's last-resort face is typically hidden (macOS
//! offers `.LastResort`) and plain name resolution cannot reach it.
//!
//! # What this costs a document that needs nothing
//!
//! One [`FallbackLookup::covers_all`] call per text fragment, and no allocation
//! or vector rebuild: the scan finds nothing and returns. That early-out is
//! load-bearing, and so is what it is made of — printable-ASCII text is
//! answered from a per-face flag, and the family a run resolved to is
//! remembered from the previous fragment, so the common case reaches neither
//! Skia nor [`crate::render::fonts::FontRegistry::resolve`].
//!
//! Both of those were added because the naive version was measurable, not
//! because they looked prudent. Resolving per fragment — which case-folds the
//! family into a fresh `String` and clones a `TypefaceEntry` — cost
//! `sample-docx-files-sample4` about +15 ms of its ~420, on a document where no
//! fallback fires at all. With the flag and the memo the same measurement is
//! +0.2 ms on `sample-docx-files-sample-4` and within noise on `sample4`.

use std::cell::RefCell;
use std::rc::Rc;

use rustc_hash::FxHashMap;

use skia_safe::{GlyphId, Typeface, Unichar};
use unicode_segmentation::UnicodeSegmentation;

use crate::render::dimension::Pt;
use crate::render::fonts::{FaceRequest, FontRegistry, Toggle, TypefaceId};

use super::{BreakAfter, FontProps, Fragment, TextMetrics};

/// How this pass asks the host what it cannot answer itself.
///
/// A trait for the same reason [`crate::render::emoji::resolve::
/// EmojiTypefaceLookup`] is one: face coverage is a property of the machine the
/// render runs on, and the splitting logic below has to be testable without
/// depending on which fonts the test machine happens to have. CI runs on Linux;
/// this engine's authors do not all.
pub trait FallbackLookup {
    /// Whether the face `font` resolves to covers every scalar in `text`.
    ///
    /// Separate from [`Self::covers`] so the common answer — "yes, all of it" —
    /// can be one call over the whole string rather than one per character.
    fn covers_all(&self, font: &FontProps, text: &str) -> bool;

    /// Whether the face `font` resolves to maps `ch` to a real glyph.
    fn covers(&self, font: &FontProps, ch: char) -> bool;

    /// The family name of a face that covers `ch`, or `None` when the host
    /// offers nothing.
    ///
    /// The implementation is responsible for making the returned name resolve
    /// back to the face it chose — see this module's doc.
    fn fallback_family(&self, font: &FontProps, ch: char) -> Option<Rc<str>>;
}

/// The previous [`RegistryFallback::typeface`] question and its answer.
type LastFace = (String, Toggle, Toggle, Typeface);

/// Which family covers a given character for a run in a given face — `None`
/// where the host offered nothing, which is cached too so a document full of an
/// uncoverable script asks once rather than once per occurrence.
type ChosenFamilies = FxHashMap<(TypefaceId, u32), Option<Rc<str>>>;

/// The real [`FallbackLookup`]: coverage from the resolved typeface's cmap,
/// substitutes from the host font system.
///
/// Kept beside the trait for the same reason
/// [`crate::render::emoji::resolve::RegistryLookup`] is kept beside its own —
/// one real implementation and one seam for tests, in one place.
///
/// Both caches are per-render, like everything else keyed on a
/// [`FontRegistry`]: the registry is owned per render and the subset pass
/// mutates it in place, so a process-wide cache here would leak one document's
/// faces into the next.
pub struct RegistryFallback<'r> {
    registry: &'r FontRegistry,
    /// Memo of "which family covers this character, for a run in this face" —
    /// the host call is by far the expensive part, and a page of CJK asks the
    /// same question for every character in it.
    chosen: RefCell<ChosenFamilies>,
    /// Reused glyph buffer for [`Self::covers_all`], so the per-fragment scan
    /// that carries this pass allocates nothing.
    scratch: RefCell<Vec<GlyphId>>,
    /// Whether a face covers printable ASCII, tested once per face.
    ///
    /// This is what makes the early-out actually early. Without it every
    /// fragment of every document pays a Skia `str_to_glyphs`, and a document
    /// like `sample-docx-files-sample4` has upwards of a hundred thousand of
    /// them. Nearly all are printable ASCII, and nearly every text face covers
    /// all of it — so answering that from a flag instead of a Skia call is the
    /// difference between this pass being free and being measurable.
    ///
    /// A flag per face rather than an assumption about faces in general: a
    /// symbol or icon font genuinely may not cover ASCII, and one that does not
    /// must still reach the slow path.
    ascii: RefCell<FxHashMap<TypefaceId, bool>>,
    /// Raw inputs of the previous [`Self::typeface`] call and its answer.
    ///
    /// The same one-entry trick, and for the same reason, as
    /// `fonts::FontCache`'s `last`: every word of a run shares one
    /// [`FontProps`], so consecutive fragments ask this identical questions.
    /// `FontRegistry::resolve` is cheap but not free — it case-folds the family
    /// into a fresh `String` and clones a `TypefaceEntry` on every call — and
    /// at one call per fragment on a 171-page document that is the whole
    /// measurable cost of this pass.
    ///
    /// Compared by string equality rather than by `Rc` pointer: a pointer would
    /// be sharper, but this cache outlives any individual fragment and a freed
    /// allocation's address can be reused, which would hand back the wrong
    /// face. Comparing short family names costs no allocation.
    last: RefCell<Option<LastFace>>,
}

impl<'r> RegistryFallback<'r> {
    pub fn new(registry: &'r FontRegistry) -> Self {
        Self {
            registry,
            chosen: RefCell::new(FxHashMap::default()),
            scratch: RefCell::new(Vec::new()),
            ascii: RefCell::new(FxHashMap::default()),
            last: RefCell::new(None),
        }
    }

    /// Whether every printable ASCII character has a glyph in `typeface`.
    ///
    /// Printable only — `U+0020..=U+007E`. Tab and newline are ASCII too and
    /// routinely have no glyph, but they never reach a text fragment (a tab is
    /// its own [`Fragment`] variant), so restricting the question to printables
    /// avoids having to special-case them and keeps the answer a plain "yes".
    fn covers_printable_ascii(&self, typeface: &Typeface) -> bool {
        let id = TypefaceId::from(typeface);
        if let Some(hit) = self.ascii.borrow().get(&id) {
            return *hit;
        }
        const PRINTABLE: &str =
            " !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`\
             abcdefghijklmnopqrstuvwxyz{|}~";
        let mut glyphs = [0 as GlyphId; 95];
        let written = typeface.str_to_glyphs(PRINTABLE, &mut glyphs);
        let covered = written == PRINTABLE.len() && glyphs[..written].iter().all(|g| *g != 0);
        self.ascii.borrow_mut().insert(id, covered);
        covered
    }

    fn typeface(&self, font: &FontProps) -> Typeface {
        if let Some((family, bold, italic, typeface)) = &*self.last.borrow() {
            if *bold == font.bold && *italic == font.italic && family.as_str() == &*font.family {
                return typeface.clone();
            }
        }
        let typeface = self
            .registry
            .resolve(&FaceRequest::new(&font.family, font.bold, font.italic))
            .typeface;
        *self.last.borrow_mut() = Some((
            font.family.to_string(),
            font.bold,
            font.italic,
            typeface.clone(),
        ));
        typeface
    }

    /// Ask the host for a face covering `ch`, and make its name resolve to it.
    ///
    /// Returns `None` — meaning "keep the face you have, and draw nothing" —
    /// in the two cases where the host cannot actually help: it offers nothing
    /// at all, or what it offers does not cover `ch` after all. The second is
    /// not paranoia: the name is pinned only when it does not already resolve
    /// elsewhere, so a document that embeds its own font under the same family
    /// name keeps it, and that font need not cover the character.
    fn lookup_uncached(&self, font: &FontProps, base: &Typeface, ch: char) -> Option<Rc<str>> {
        // Matching on the base face's own style keeps the fallback at the
        // weight and slant the run asked for, rather than resetting to regular.
        let found = self.registry.font_mgr().match_family_style_character(
            &*font.family,
            base.font_style(),
            &[],
            ch as Unichar,
        )?;
        if found.unichar_to_glyph(ch as Unichar) == 0 {
            return None;
        }

        let name: Rc<str> = Rc::from(found.family_name().as_str());
        self.registry
            .pin_system_face(&name, font.bold, font.italic, found);

        // Verify against the registry rather than against `found`: what the
        // rest of the pipeline will draw with is whatever this name resolves
        // to now, which is not `found` when something else already held the
        // name. If that face cannot draw `ch` either, reporting the name would
        // move the text to a different face and still draw nothing.
        let settled = self
            .registry
            .resolve(&FaceRequest::new(&name, font.bold, font.italic))
            .typeface;
        (settled.unichar_to_glyph(ch as Unichar) != 0).then_some(name)
    }
}

impl FallbackLookup for RegistryFallback<'_> {
    fn covers_all(&self, font: &FontProps, text: &str) -> bool {
        let typeface = self.typeface(font);
        // The overwhelmingly common case, answered without touching Skia.
        if text.bytes().all(|b| (0x20..0x7f).contains(&b)) && self.covers_printable_ascii(&typeface)
        {
            return true;
        }
        let mut buf = self.scratch.borrow_mut();
        buf.clear();
        buf.resize(text.chars().count(), 0);
        let written = typeface.str_to_glyphs(text, &mut buf);
        // One glyph per scalar, in order. An ignorable that the face lacks is
        // not a miss — see `is_ignorable_for_coverage`; counting it as one
        // would push every document containing a soft hyphen onto the slow
        // path for nothing.
        buf[..written]
            .iter()
            .zip(text.chars())
            .all(|(glyph, ch)| *glyph != 0 || is_ignorable_for_coverage(ch))
    }

    fn covers(&self, font: &FontProps, ch: char) -> bool {
        self.typeface(font).unichar_to_glyph(ch as Unichar) != 0
    }

    fn fallback_family(&self, font: &FontProps, ch: char) -> Option<Rc<str>> {
        let typeface = self.typeface(font);
        let key = (TypefaceId::from(&typeface), ch as u32);
        if let Some(hit) = self.chosen.borrow().get(&key) {
            return hit.clone();
        }
        let chosen = self.lookup_uncached(font, &typeface, ch);
        if chosen.is_none() {
            // Once per (face, codepoint) per render, because the memo below
            // stops this arm being reached again for the same pair.
            log::warn!(
                "no host font covers U+{:04X}; it will not be drawn (font '{}')",
                ch as u32,
                font.family,
            );
        }
        self.chosen.borrow_mut().insert(key, chosen.clone());
        chosen
    }
}

/// Codepoints that say nothing about which face should draw a cluster.
///
/// A face is free not to have these and still render the cluster correctly:
/// they are joiners, bidi controls and variation selectors, none of which is
/// drawn. Asking the host for a face that "covers" U+200D would swap the face
/// for a whole cluster on the strength of a character that has no glyph in any
/// font — which is how a working emoji or Devanagari cluster gets moved to the
/// wrong face.
///
/// Written out rather than taken from a Unicode property table because the set
/// that matters here is small, closed, and clearer named than derived. It is
/// deliberately *not* the full `Default_Ignorable_Code_Point` set: this is a
/// list of things that must not *drive* a decision, not a list of things that
/// need no glyph.
fn is_ignorable_for_coverage(ch: char) -> bool {
    matches!(ch,
        '\u{00AD}'                      // SOFT HYPHEN
        | '\u{200B}'..='\u{200F}'       // ZWSP, ZWNJ, ZWJ, LRM, RLM
        | '\u{202A}'..='\u{202E}'       // bidi embedding/override controls
        | '\u{2066}'..='\u{2069}'       // bidi isolate controls
        | '\u{FEFF}'                    // BOM / ZWNBSP
        | '\u{FE00}'..='\u{FE0F}'       // variation selectors
        | '\u{E0100}'..='\u{E01EF}'     // variation selectors supplement
    )
}

/// The scalar that decides which face draws this cluster: its base.
///
/// Coverage is asked of the base alone, never of every scalar. A cluster is one
/// unit — [`crate::render::spacing`] owns why that is the unit this engine
/// measures in — so it must be drawn by one face, and the base is what that
/// face has to be able to draw. A combining mark the base's face happens to
/// lack is a worse reason to move the cluster elsewhere than to leave it: a
/// mark positions against the base's own metrics, so splitting the two across
/// faces misplaces it.
fn cluster_base(cluster: &str) -> Option<char> {
    cluster.chars().find(|c| !is_ignorable_for_coverage(*c))
}

/// Choose the family for one cluster: `None` keeps the fragment's own.
fn choose_family(cluster: &str, font: &FontProps, lookup: &impl FallbackLookup) -> Option<Rc<str>> {
    let base = cluster_base(cluster)?;
    if lookup.covers(font, base) {
        return None;
    }
    lookup.fallback_family(font, base)
}

/// Give every cluster a face that can draw it, splitting where that face
/// changes.
///
/// Call once per paragraph on the **finished** fragment vector, between
/// [`super::assign_bidi_levels`] and [`super::shape_complex_scripts`]. Both
/// edges of that ordering are forced:
///
/// - *after* bidi, because a coverage boundary cannot change an embedding
///   level — each piece inherits the level its fragment already carries, and
///   running first would only make bidi rebuild its analysis over more and
///   smaller fragments.
/// - *before* shaping, because shaping re-measures against the resolved
///   typeface. A fragment whose family changed here must be shaped against the
///   face it actually ends up with, or it is measured one way and painted
///   another — the failure `super::shape` documents as stranding an
///   underline.
pub fn apply_font_fallback<F>(
    fragments: &mut Vec<Fragment>,
    lookup: &impl FallbackLookup,
    measure_text: &F,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    // The early-out. See this module's doc for why it carries the pass.
    let needed = fragments.iter().any(|f| match f {
        Fragment::Text { text, font, .. } => !lookup.covers_all(font, text),
        _ => false,
    });
    if !needed {
        return;
    }

    let mut out = Vec::with_capacity(fragments.len());
    for fragment in fragments.drain(..) {
        split_at_coverage_boundaries(fragment, lookup, measure_text, &mut out);
    }
    *fragments = out;
}

fn split_at_coverage_boundaries<F>(
    fragment: Fragment,
    lookup: &impl FallbackLookup,
    measure_text: &F,
    out: &mut Vec<Fragment>,
) where
    F: Fn(&str, &FontProps) -> (Pt, TextMetrics),
{
    let Fragment::Text {
        ref text, ref font, ..
    } = fragment
    else {
        out.push(fragment);
        return;
    };

    if lookup.covers_all(font, text) {
        out.push(fragment);
        return;
    }

    // Byte ranges of maximal runs sharing one choice of family.
    let mut runs: Vec<(usize, usize, Option<Rc<str>>)> = Vec::new();
    for (offset, cluster) in text.grapheme_indices(true) {
        let choice = choose_family(cluster, font, lookup);
        let end = offset + cluster.len();
        match runs.last_mut() {
            Some(last) if last.2 == choice => last.1 = end,
            _ => runs.push((offset, end, choice)),
        }
    }

    // Nothing the host could improve on: one run, and it keeps its own family.
    // Reached whenever no face covers the missing codepoints, which is the
    // stated no-fallback behaviour — see `apply_font_fallback`'s callers.
    if runs.len() == 1 && runs[0].2.is_none() {
        out.push(fragment);
        return;
    }

    let Fragment::Text {
        text,
        font,
        color,
        shading,
        border,
        break_after,
        level,
        hyperlink_url,
        baseline_offset,
        text_offset,
        is_footnote_ref,
        ..
    } = fragment
    else {
        unreachable!("guarded by the `let ... else` above")
    };

    let last = runs.len() - 1;
    for (i, (start, end, family)) in runs.into_iter().enumerate() {
        let piece = &text[start..end];
        let piece_font = match family {
            Some(name) => Rc::new(FontProps {
                family: name,
                ..(*font).clone()
            }),
            None => Rc::clone(&font),
        };
        let (w, m) = measure_text(piece, &piece_font);
        let trimmed = piece.trim_end();
        let tw = if trimmed.len() < piece.len() {
            measure_text(trimmed, &piece_font).0
        } else {
            w
        };
        out.push(Fragment::Text {
            text: Rc::from(piece),
            font: piece_font,
            color,
            shading,
            border,
            // A coverage boundary is not a line-break opportunity — UAX #14
            // knew nothing about it, and the fragment was one word before this
            // pass ran. Only the last piece keeps whatever opportunity the
            // whole fragment had earned at its trailing edge.
            break_after: if i == last {
                break_after
            } else {
                BreakAfter::Prohibited
            },
            level,
            // Always `None` here: this pass runs before `shape_complex_scripts`,
            // which is what sets it, precisely so that it sees these pieces.
            shaped: None,
            width: w,
            trimmed_width: tw,
            metrics: m,
            hyperlink_url: hyperlink_url.clone(),
            baseline_offset,
            text_offset,
            is_footnote_ref,
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::i18n::bidi::BidiLevel;
    use crate::render::resolve::color::RgbColor;
    use std::collections::HashMap;

    /// A coverage oracle spelled out per test, so nothing here depends on the
    /// fonts the machine running it happens to have.
    struct FakeLookup {
        /// Characters the fragment's own face can draw.
        covered: Vec<char>,
        /// Which family to offer for a character it cannot.
        fallbacks: HashMap<char, &'static str>,
    }

    impl FakeLookup {
        fn new(covered: &str, fallbacks: &[(char, &'static str)]) -> Self {
            Self {
                covered: covered.chars().collect(),
                fallbacks: fallbacks.iter().copied().collect(),
            }
        }
    }

    impl FallbackLookup for FakeLookup {
        fn covers_all(&self, font: &FontProps, text: &str) -> bool {
            text.chars().all(|c| self.covers(font, c))
        }
        fn covers(&self, _font: &FontProps, ch: char) -> bool {
            self.covered.contains(&ch)
        }
        fn fallback_family(&self, _font: &FontProps, ch: char) -> Option<Rc<str>> {
            self.fallbacks.get(&ch).map(|n| Rc::from(*n))
        }
    }

    fn font(family: &str) -> Rc<FontProps> {
        Rc::new(FontProps {
            family: Rc::from(family),
            size: Pt::new(12.0),
            bold: Toggle::Absent,
            italic: Toggle::Absent,
            underline: false,
            rtl: Toggle::Absent,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
            strike_lines: 0,
            strike_position: Pt::ZERO,
            strike_thickness: Pt::ZERO,
        })
    }

    fn frag(text: &str) -> Fragment {
        Fragment::Text {
            text: Rc::from(text),
            font: font("Base"),
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            break_after: BreakAfter::Opportunity,
            level: BidiLevel::LTR,
            shaped: None,
            width: Pt::new(50.0),
            trimmed_width: Pt::new(50.0),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    /// One `Pt` per character, so a split's re-measurement is checkable.
    fn measure(text: &str, _f: &FontProps) -> (Pt, TextMetrics) {
        (
            Pt::new(text.chars().count() as f32),
            TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
        )
    }

    fn parts(fragments: &[Fragment]) -> Vec<(String, String)> {
        fragments
            .iter()
            .filter_map(|f| match f {
                Fragment::Text { text, font, .. } => {
                    Some((text.to_string(), font.family.to_string()))
                }
                _ => None,
            })
            .collect()
    }

    /// Acceptance criterion 3, asserted directly rather than inferred from
    /// pixels: a fragment whose face covers everything is not touched at all.
    #[test]
    fn a_covered_fragment_is_returned_untouched() {
        let lookup = FakeLookup::new("abc", &[]);
        let mut frags = vec![frag("abc")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(parts(&frags), [("abc".into(), "Base".into())]);
        // And its measurement is the one it arrived with, not a re-measure.
        match &frags[0] {
            Fragment::Text { width, .. } => assert_eq!(width.raw(), 50.0),
            _ => panic!("expected text"),
        }
    }

    #[test]
    fn a_missing_character_splits_the_fragment_and_changes_only_its_family() {
        let lookup = FakeLookup::new("ab", &[('ア', "CJK")]);
        let mut frags = vec![frag("aアb")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(
            parts(&frags),
            [
                ("a".into(), "Base".into()),
                ("ア".into(), "CJK".into()),
                ("b".into(), "Base".into()),
            ]
        );
    }

    /// Every piece is re-measured against the font it ends up with — the
    /// contract `bidi`'s splitter has, and the one that keeps a line's widths
    /// honest.
    #[test]
    fn split_pieces_are_remeasured() {
        let lookup = FakeLookup::new("ab", &[('ア', "CJK")]);
        let mut frags = vec![frag("aアb")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        for f in &frags {
            match f {
                Fragment::Text { width, .. } => assert_eq!(width.raw(), 1.0),
                _ => panic!("expected text"),
            }
        }
    }

    /// A coverage boundary is not a break opportunity: the word must not
    /// become breakable because part of it needed another face.
    #[test]
    fn inner_pieces_may_not_be_broken_after() {
        let lookup = FakeLookup::new("ab", &[('ア', "CJK")]);
        let mut frags = vec![frag("aアb")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        let breaks: Vec<BreakAfter> = frags
            .iter()
            .filter_map(|f| match f {
                Fragment::Text { break_after, .. } => Some(*break_after),
                _ => None,
            })
            .collect();
        assert_eq!(
            breaks,
            [
                BreakAfter::Prohibited,
                BreakAfter::Prohibited,
                BreakAfter::Opportunity
            ]
        );
    }

    /// Consecutive missing clusters that resolve to the same face are one
    /// fragment, not one per cluster — otherwise a line of CJK becomes a
    /// fragment per character and line fitting slows to a crawl.
    #[test]
    fn a_run_of_missing_characters_becomes_one_fragment() {
        let lookup = FakeLookup::new("ab", &[('ア', "CJK"), ('イ', "CJK"), ('ウ', "CJK")]);
        let mut frags = vec![frag("aアイウb")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(
            parts(&frags),
            [
                ("a".into(), "Base".into()),
                ("アイウ".into(), "CJK".into()),
                ("b".into(), "Base".into()),
            ]
        );
    }

    /// A base plus its combining mark is one cluster and one unit of drawing.
    /// The mark is not covered here, and that must not split the cluster — a
    /// mark positions against its base's metrics, so the two cannot be drawn
    /// by different faces.
    #[test]
    fn a_combining_mark_is_never_split_from_its_base() {
        // U+0301 COMBINING ACUTE ACCENT is deliberately absent from `covered`.
        let lookup = FakeLookup::new("ea", &[('\u{0301}', "Marks")]);
        let mut frags = vec![frag("e\u{0301}a")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(parts(&frags), [("e\u{0301}a".into(), "Base".into())]);
    }

    /// A joiner the face lacks says nothing about which face should draw the
    /// cluster. Falling back on it would move working text to another face.
    #[test]
    fn a_zero_width_joiner_the_face_lacks_does_not_trigger_fallback() {
        let lookup = FakeLookup::new("ab", &[('\u{200D}', "Wrong")]);
        let mut frags = vec![frag("a\u{200D}b")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(parts(&frags), [("a\u{200D}b".into(), "Base".into())]);
    }

    /// The stated behaviour when the host offers nothing: keep the original
    /// face. The codepoint still does not draw, but nothing else moves.
    #[test]
    fn a_host_with_no_covering_face_leaves_the_fragment_alone() {
        let lookup = FakeLookup::new("ab", &[]);
        let mut frags = vec![frag("aアb")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(parts(&frags), [("aアb".into(), "Base".into())]);
    }

    /// Two different missing scripts get two different faces, and the split
    /// tracks each boundary rather than lumping them together.
    #[test]
    fn different_scripts_get_different_faces() {
        let lookup = FakeLookup::new("ab", &[('ア', "CJK"), ('๑', "Thai")]);
        let mut frags = vec![frag("aア๑b")];
        apply_font_fallback(&mut frags, &lookup, &measure);
        assert_eq!(
            parts(&frags),
            [
                ("a".into(), "Base".into()),
                ("ア".into(), "CJK".into()),
                ("๑".into(), "Thai".into()),
                ("b".into(), "Base".into()),
            ]
        );
    }

    /// Non-text fragments pass through, and a document with nothing missing
    /// never rebuilds its vector.
    #[test]
    fn a_document_that_needs_nothing_is_not_rebuilt() {
        let lookup = FakeLookup::new("abc", &[('ア', "CJK")]);
        let mut frags = vec![frag("abc"), frag("cba")];
        let before: Vec<*const u8> = frags
            .iter()
            .map(|f| match f {
                Fragment::Text { text, .. } => text.as_ptr(),
                _ => std::ptr::null(),
            })
            .collect();
        apply_font_fallback(&mut frags, &lookup, &measure);
        let after: Vec<*const u8> = frags
            .iter()
            .map(|f| match f {
                Fragment::Text { text, .. } => text.as_ptr(),
                _ => std::ptr::null(),
            })
            .collect();
        assert_eq!(before, after, "the same Rc<str> allocations, untouched");
    }
}
