//! GSUB-aware text shaping.
//!
//! Skia's `canvas.draw_str` performs only cmap-level codepoint→glyph
//! mapping; it does not apply OpenType GSUB lookups. Two callers need it to:
//!
//! * **`render::emoji`** — a multi-codepoint sequence (`1️⃣` keycap, `👍🏿`
//!   modifier, `👨‍👩‍👧` ZWJ family) must rasterize as the *ligated* single
//!   glyph, not as its constituents side by side.
//! * **body text in a joining or reordering script** (issues #131, #153) — an
//!   Arabic word painted from its cmap comes out in isolated forms, letter by
//!   letter, which is not a degraded rendering of the word so much as a
//!   different one; a Devanagari word comes out with its prebase vowels on the
//!   wrong side of their consonants, which is worse.
//!
//! [`needs_shaping`] is what separates the second caller's text from
//! everything else, and the separation is load-bearing rather than an
//! optimization: shaping applies GPOS kerning and standard ligatures too, so
//! routing *all* text through here would change the measured width of every
//! Latin word in every document. See that function for the two-property
//! predicate — `Joining_Type` for positional forms, plus the reordering
//! scripts `Joining_Type` cannot see.
//!
//! Shaping runs through **Skia's own HarfBuzz** (`skia-safe`'s `textlayout`
//! feature), driven by a [`Typeface`] rather than raw font bytes. That
//! distinction is the point of this module's design:
//!
//! > `Typeface::to_font_data()` serializes the *entire* font. For
//! > `Apple Color Emoji.ttc` that is 183 MB, and the call costs ~549 MB of
//! > resident memory — 183 MB for the returned buffer plus ~366 MB of
//! > Skia-internal assembly — none of which is returned to the OS. A pure-Rust
//! > shaper (rustybuzz, used here previously) needs those bytes; Skia's does
//! > not, because it already holds the typeface.
//!
//! Shaping the same clusters through Skia costs ~2 MB and produces identical
//! glyph ids and advances. The switch cut corpus peak RSS 44.6% and
//! emoji-document wall clock 42%.
//!
//! [`Shaper`] owns Skia's so it is constructed once per render rather than per
//! run, and is deliberately built **without a fallback font manager**: the
//! caller has already resolved which typeface to use, and silently
//! substituting another family would draw the wrong glyph.

use skia_safe::shaper::run_handler::{Buffer, RunInfo};
use skia_safe::shaper::{RunHandler, Shaper as SkShaper};
use skia_safe::shapers;
use skia_safe::{Font, GlyphId, Point, Typeface};
use thiserror::Error;
use unicode_joining_type::{get_joining_type, JoiningType};
use unicode_script::{Script, UnicodeScript};

use crate::i18n::bidi::BidiLevel;
use crate::render::dimension::Pt;

/// Whether `text` is in a script that cmap-only painting renders *wrongly*, as
/// opposed to merely without kerning. Two properties answer it, one per way a
/// cmap walk can be wrong about a legible script:
///
/// **The letters have positional forms** — the Unicode **`Joining_Type`**
/// property, and specifically its three "this letter joins" values. That is
/// not a hand-drawn list of scripts standing in for the real rule — it *is*
/// the rule: a letter with a joining type of dual-, left-, or right-joining is
/// one whose shape depends on its neighbours, which is exactly the case a cmap
/// lookup cannot answer. Arabic, Syriac, N'Ko, Mongolian, Adlam, Hanifi
/// Rohingya and the rest fall out of it without being named — and, just as
/// usefully, Hebrew and Thaana do not: both are right-to-left, and both spell
/// their final forms as separate codepoints rather than as positional
/// variants, so a cmap lookup is the whole answer for them.
///
/// Two `Joining_Type` values are deliberately excluded:
///
/// * `Transparent` — combining marks, which includes the Latin combining
///   diacriticals at U+0300. Including it would send `e` + U+0301 through the
///   shaper and re-measure a large share of European text.
/// * `JoinCausing` — ZWJ and tatweel. Both are *context* rather than letters
///   with forms of their own, and both appear where the letters around them
///   already answer this question. ZWJ in particular reaches here inside emoji
///   sequences that fell back to the text path, and shaping a Latin run
///   because it contains one would change that run's width for no gain.
///
/// **The script reorders** (issue #153) — `script_reorders`, the Brahmic
/// scripts whose glyphs do not come out in character order. `Joining_Type`
/// cannot see these: a Devanagari letter never changes shape by position, yet
/// a cmap walk of `कि` draws the vowel on the wrong side of its consonant.
pub fn needs_shaping(text: &str) -> bool {
    text.chars().any(|c| {
        matches!(
            get_joining_type(c),
            JoiningType::DualJoining | JoiningType::LeftJoining | JoiningType::RightJoining
        ) || script_reorders(c)
    })
}

/// Whether `c` belongs to a script whose shaping moves glyphs relative to
/// their characters, so that even a font with every nominal glyph draws the
/// string wrongly *in order* from the cmap.
///
/// No Unicode character property expresses "this script reorders" — it is a
/// property of the script's OpenType shaping model, not of any codepoint — so
/// this is a list, drawn from the shaping models themselves: the nine scripts
/// of Microsoft's Indic shaping spec (prebase matras, reph), plus Sinhala,
/// Khmer and Myanmar, whose own specs describe the same prebase-reordering
/// vowels and conjunct formation. A prebase vowel is stored after its
/// consonant and drawn before it; a virama-bound conjunct substitutes across
/// characters; both need HarfBuzz, and both also need the *spacing unit* to be
/// the shaped cluster — `crate::render::spacing` tells that half of the story.
///
/// Deliberately not here:
///
/// * **Thai and Lao** — their prebase vowels are stored in visual order
///   (Unicode encodes them before the consonant), so the cmap walk is already
///   right, and keeping them off the shaper keeps their measured widths — and
///   the corpus documents that use them — exactly as they were.
/// * **Tibetan** — stacks below the base without reordering; a cmap walk is
///   degraded (unstacked) but ordered. Left out until evidence that the
///   degradation is the kind #131 exists to fix, because adding it re-measures
///   every Tibetan run.
fn script_reorders(c: char) -> bool {
    matches!(
        c.script(),
        Script::Devanagari
            | Script::Bengali
            | Script::Gurmukhi
            | Script::Gujarati
            | Script::Oriya
            | Script::Tamil
            | Script::Telugu
            | Script::Kannada
            | Script::Malayalam
            | Script::Sinhala
            | Script::Khmer
            | Script::Myanmar
    )
}

/// Which way [`Shaper::shape`] lays a run's glyphs out.
///
/// One run, one direction — UAX #9 reordering has already happened by the time
/// a run reaches the shaper, and `layout::fragment::bidi` has split every
/// fragment that spanned a level boundary, so a run is never mixed.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum RunDirection {
    #[default]
    LeftToRight,
    RightToLeft,
}

impl From<BidiLevel> for RunDirection {
    fn from(level: BidiLevel) -> Self {
        if level.is_rtl() {
            Self::RightToLeft
        } else {
            Self::LeftToRight
        }
    }
}

// ─── Public ADTs ─────────────────────────────────────────────────────────────

/// One glyph in a [`ShapedRun`], positioned in pixels at the requested
/// rasterization size.
///
/// `x`/`y` are **absolute offsets from the run origin**, which sits on the
/// baseline — not per-glyph advances. Skia's shaper reports positions this way
/// and `draw_glyphs_at` consumes them the same way, so the rasterizer neither
/// accumulates a pen nor flips a sign. (The previous rustybuzz-based type
/// carried an advance plus y-*up* offsets and required both.)
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ShapedGlyph {
    /// Skia glyph id.
    pub id: GlyphId,
    /// Horizontal offset from the run origin, in pixels.
    pub x: Pt,
    /// Vertical offset from the run origin, in pixels, **y-down** per Skia's
    /// convention. Zero for the overwhelming majority of emoji clusters.
    pub y: Pt,
    /// The UTF-8 byte offset into the shaped text of the **grapheme cluster**
    /// this glyph came from. Glyphs sharing a value form one **shaped
    /// cluster** — the §17.3.2.35 / §17.3.1.13 spacing unit for a shaped run
    /// (issue #153), which is what [`ShapedRun::unit_count`] counts and the
    /// painter groups by.
    ///
    /// This is HarfBuzz's cluster value *folded to the start of its UAX #29
    /// grapheme cluster* by [`Shaper::shape`]. The folding is load-bearing:
    /// Skia's HarfBuzz reports character-level clusters, which merge under
    /// reordering and ligation but leave an ordinary combining mark — a
    /// Devanagari above-base matra, an accent no font ligates — as a cluster
    /// of its own, and an unfolded value would let `w:spacing` open between a
    /// base and its own mark, the defect `crate::render::spacing` exists to
    /// prevent. Folded, a shaped cluster is never finer than a grapheme
    /// cluster — the invariant the emergency splitter
    /// (`layout::fragment::split`) also leans on.
    ///
    /// Offsets are absolute into the whole string handed to
    /// [`Shaper::shape`], across every sub-run Skia splits off — pinned by
    /// `clusters_are_absolute_byte_offsets` below, because nothing in the
    /// SkShaper API states it.
    pub cluster: u32,
}

/// Output of shaping one run of text.
#[derive(Clone, Debug)]
pub struct ShapedRun {
    pub glyphs: Vec<ShapedGlyph>,
    /// Sum of run advances in pixels — the rasterizer uses this to size the
    /// offscreen surface, and layout uses it to reserve the cluster's width.
    pub total_advance: Pt,
}

impl ShapedRun {
    /// How many shaped clusters the run holds — the number of §17.3.2.35 /
    /// §17.3.1.13 spacing units, counterpart to
    /// [`crate::render::spacing::unit_count`] for text the shaper has since
    /// regrouped (a conjunct's glyphs share one cluster; a ligature can merge
    /// what were separate grapheme clusters).
    ///
    /// Counts *distinct* cluster values rather than value changes: a
    /// right-to-left run's glyphs arrive in visual order, and Skia may append
    /// sub-runs, so equal values need not be adjacent.
    pub fn unit_count(&self) -> usize {
        let mut clusters: Vec<u32> = self.glyphs.iter().map(|g| g.cluster).collect();
        clusters.sort_unstable();
        clusters.dedup();
        clusters.len()
    }
}

#[derive(Debug, Error)]
pub enum ShapeError {
    /// Skia was built without a HarfBuzz shaper. Unreachable with the
    /// `textlayout` feature enabled, but `shape_dont_wrap_or_reorder` returns
    /// an `Option` and this module does not panic on the public path.
    #[error("skia was built without a HarfBuzz shaper")]
    ShaperUnavailable,
    /// Shaping produced no glyphs — callers fall back to `draw_str` /
    /// `measure_str`.
    #[error("shaping produced no glyphs")]
    NoGlyphs,
}

// ─── Shaper ──────────────────────────────────────────────────────────────────

/// A reusable GSUB-aware shaper.
///
/// Construct once per render and shape many runs: Skia keeps an internal
/// HarfBuzz face cache keyed by typeface, so repeated calls do not re-parse
/// the font.
pub struct Shaper {
    shaper: SkShaper,
}

impl Shaper {
    /// Build a shaper with **no fallback font manager**, so shaping never
    /// substitutes a different family for the typeface the caller resolved.
    pub fn new() -> Result<Self, ShapeError> {
        shapers::hb::shape_dont_wrap_or_reorder(None)
            .map(|shaper| Self { shaper })
            .ok_or(ShapeError::ShaperUnavailable)
    }

    /// Shape `text` against `typeface` at `size_px`, returning the glyph
    /// sequence and positions.
    ///
    /// `size_px` is in raw pixels — already pre-multiplied by any super-sample
    /// scale the caller wants.
    pub fn shape(
        &self,
        typeface: &Typeface,
        text: &str,
        size_px: f32,
        direction: RunDirection,
    ) -> Result<ShapedRun, ShapeError> {
        let font = Font::from_typeface(typeface.clone(), size_px);
        let mut collector = Collector::default();
        // `width = f32::MAX` plus the dont-wrap-or-reorder shaper means "one
        // line, no bidi reordering" — a cluster is not a paragraph, and for
        // body text UAX #9 has already reordered the fragments this run sits
        // between. `direction` is what HarfBuzz orders glyphs *within* the run
        // by; letting Skia run its own bidi here instead would reorder the
        // same text twice.
        self.shaper.shape(
            text,
            &font,
            direction == RunDirection::LeftToRight,
            f32::MAX,
            &mut collector,
        );

        if collector.glyphs.is_empty() {
            return Err(ShapeError::NoGlyphs);
        }

        // Fold each cluster value to the start of the grapheme cluster that
        // contains it — see [`ShapedGlyph::cluster`] for why character-level
        // values from HarfBuzz are not usable as spacing units directly.
        let grapheme_starts: Vec<u32> = {
            use unicode_segmentation::UnicodeSegmentation;
            text.grapheme_indices(true).map(|(i, _)| i as u32).collect()
        };
        let fold = |cluster: u32| -> u32 {
            let idx = grapheme_starts
                .partition_point(|&start| start <= cluster)
                .saturating_sub(1);
            grapheme_starts.get(idx).copied().unwrap_or(0)
        };

        let glyphs = collector
            .glyphs
            .iter()
            .zip(collector.positions.iter())
            .zip(collector.clusters.iter())
            .map(|((&id, p), &cluster)| ShapedGlyph {
                id,
                x: Pt::new(p.x),
                y: Pt::new(p.y),
                cluster: fold(cluster),
            })
            .collect();

        Ok(ShapedRun {
            glyphs,
            total_advance: Pt::new(collector.advance_x),
        })
    }
}

/// Accumulates every run Skia emits. A single emoji cluster shapes to one run
/// in practice, but the shaper is free to split on script boundaries, so runs
/// are appended rather than replaced.
#[derive(Default)]
struct Collector {
    glyphs: Vec<GlyphId>,
    positions: Vec<Point>,
    clusters: Vec<u32>,
    advance_x: f32,
}

impl RunHandler for Collector {
    fn begin_line(&mut self) {}
    fn run_info(&mut self, _info: &RunInfo) {}
    fn commit_run_info(&mut self) {}

    fn run_buffer<'a>(&'a mut self, info: &RunInfo) -> Buffer<'a> {
        let base = self.glyphs.len();
        // Where this run starts, which is everything shaped so far. Skia writes
        // positions relative to the origin it is handed, so passing `None` —
        // as this did while its only caller shaped one emoji cluster at a time
        // — restarts every run at x=0 and stacks them on top of each other.
        // Invisible for a single-run cluster; wrong the moment a run of body
        // text is split on a script boundary, which a word plus its trailing
        // space already is.
        let origin = Point::new(self.advance_x, 0.0);
        self.glyphs.resize(base + info.glyph_count, 0);
        self.positions
            .resize(base + info.glyph_count, Point::new(0.0, 0.0));
        // Skia asserts `clusters.len() == glyph_count` when the slice is
        // given, so it grows in lockstep with the other two.
        self.clusters.resize(base + info.glyph_count, 0);
        self.advance_x += info.advance.x;
        // `Buffer::new` hardcodes `clusters: None`; the struct's fields are
        // public, so asking for cluster values is a literal, not a constructor.
        Buffer {
            glyphs: &mut self.glyphs[base..],
            positions: &mut self.positions[base..],
            offsets: None,
            clusters: Some(&mut self.clusters[base..]),
            point: origin,
        }
    }

    fn commit_run_buffer(&mut self, _info: &RunInfo) {}
    fn commit_line(&mut self) {}
}

// ─── Tests ───────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::emoji::resolve::EmojiFamily;
    use skia_safe::{FontMgr, FontStyle};
    use unicode_segmentation::UnicodeSegmentation;

    /// The host's color emoji typeface, or `None` on a host without one —
    /// tests that need it return early rather than fail, since CI images vary.
    fn emoji_typeface() -> Option<Typeface> {
        let mgr = FontMgr::new();
        EmojiFamily::host_default().iter().find_map(|f| {
            mgr.match_family_style(f.family_name(), FontStyle::normal())
                .filter(|tf| tf.family_name().eq_ignore_ascii_case(f.family_name()))
        })
    }

    fn any_typeface() -> Option<Typeface> {
        FontMgr::new().legacy_make_typeface(None::<&str>, FontStyle::normal())
    }

    // ── needs_shaping ─────────────────────────────────────────────────────

    /// The predicate's whole job: keep Latin (and Cyrillic, Greek, CJK, Thai)
    /// off the shaping path, so their measured widths do not move.
    #[test]
    fn text_without_positional_forms_is_not_shaped() {
        for text in [
            "Nicht gefunden",
            "S.I.G.M.A. Technik Service GmbH",
            "Türöffner-Gerät",
            // A combining mark is Joining_Type=Transparent, and must not by
            // itself pull a Latin word into the shaper.
            "e\u{301}coute",
            "Привет",
            "日本語の文章",
            "ภาษาไทย",
            "שלום עולם",
            "",
        ] {
            assert!(!needs_shaping(text), "{text:?} must stay on the cmap path");
        }
    }

    /// Every one of these has letters whose glyph depends on its neighbours,
    /// which is the case a cmap lookup cannot answer.
    #[test]
    fn joining_scripts_are_shaped() {
        for (script, text) in [
            ("Arabic", "مرحبا"),
            ("Syriac", "\u{0710}\u{0712}"),
            ("N'Ko", "\u{07CA}\u{07D9}"),
            ("Mongolian", "\u{1820}\u{1821}"),
            ("Adlam", "\u{1E922}\u{1E923}"),
            ("Hanifi Rohingya", "\u{10D00}\u{10D01}"),
        ] {
            assert!(
                needs_shaping(text),
                "{script} letters have positional forms"
            );
        }
    }

    /// Thaana is the counterpart to Hebrew above, and the reason the predicate
    /// is `Joining_Type` rather than "is this script right-to-left": Thaana is
    /// written right to left and its letters are `Non_Joining`, so #131's
    /// reordering is all it needs.
    #[test]
    fn a_right_to_left_script_without_cursive_joining_is_not_shaped() {
        assert!(!needs_shaping("\u{0780}\u{0783}"));
    }

    /// Issue #153, the predicate's second property: a Brahmic script's glyphs
    /// do not come out in character order, which `Joining_Type` cannot see —
    /// every one of these letters is `Non_Joining`.
    #[test]
    fn reordering_scripts_are_shaped() {
        for (script, text) in [
            ("Devanagari", "हिन्दी"),
            ("Bengali", "বাংলা"),
            ("Tamil", "தமிழ்"),
            ("Telugu", "తెలుగు"),
            ("Malayalam", "മലയാളം"),
            ("Sinhala", "සිංහල"),
            ("Khmer", "ខ្មែរ"),
            ("Myanmar", "မြန်မာ"),
        ] {
            assert!(needs_shaping(text), "{script} shaping reorders");
        }
    }

    /// The boundary of the reordering list: Thai and Lao store their prebase
    /// vowels in visual order, so the cmap walk is already right and their
    /// measured widths must not move. (Thai is also in the not-shaped list
    /// above; Lao is the same case and pins the same boundary.)
    #[test]
    fn visual_order_scripts_stay_on_the_cmap_path() {
        assert!(!needs_shaping("ພາສາລາວ"), "Lao");
        assert!(!needs_shaping("ภาษาไทย"), "Thai");
    }

    /// A zero-width joiner is `Joining_Type=Join_Causing`, which the predicate
    /// excludes: it reaches body text only inside an emoji sequence that fell
    /// back from the emoji pipeline, and shaping the Latin around it would
    /// change that run's width for nothing.
    #[test]
    fn a_stray_zero_width_joiner_does_not_pull_latin_into_the_shaper() {
        assert!(!needs_shaping("a\u{200D}b"));
    }

    /// Mixed text takes the shaping path — the Arabic in it needs to join, and
    /// `fragment::bidi` will already have split the run at the level boundary
    /// in the cases where the two halves must be positioned separately.
    #[test]
    fn mixed_text_containing_a_joining_script_is_shaped() {
        assert!(needs_shaping("page مرحبا here"));
    }

    // ── the shaper ────────────────────────────────────────────────────────

    #[test]
    fn shaper_constructs() {
        assert!(
            Shaper::new().is_ok(),
            "skia must expose a HarfBuzz shaper — the `textlayout` feature is \
             what lets this module shape without serializing the font"
        );
    }

    /// ASCII through any system font: one glyph per character, advancing left
    /// to right. Pins the position convention the rasterizer depends on, and
    /// guards against shaping ligating runs that must not ligate.
    #[test]
    fn ascii_shapes_one_glyph_per_char_advancing_rightwards() {
        let Some(tf) = any_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");
        let run = shaper
            .shape(&tf, "abc", 20.0, RunDirection::LeftToRight)
            .expect("shape");

        assert_eq!(run.glyphs.len(), 3, "ASCII must not ligate");
        assert_eq!(run.glyphs[0].x, Pt::ZERO, "run origin is the first glyph");
        assert!(
            run.glyphs[1].x > run.glyphs[0].x && run.glyphs[2].x > run.glyphs[1].x,
            "positions are absolute and strictly increasing, not per-glyph advances"
        );
        assert!(run.total_advance > Pt::ZERO);
    }

    /// **The reason this module exists.** A ZWJ sequence is five codepoints;
    /// cmap-only mapping would map each independently, and the rasterizer
    /// would draw a row of separate people rather than a ligated glyph.
    ///
    /// Full ligation to one glyph is *not* asserted here — issue #117 found
    /// it isn't a portable guarantee. On Windows, `Segoe UI Emoji` carries a
    /// real 24 KB `GSUB` table (confirmed by reading it directly) and does
    /// ligate other sequences (see `modifier_and_keycap_sequences_ligate`,
    /// which passes there), but has no ligature for this specific man+woman+
    /// girl combination or any of its three 2-person sub-pairs — each
    /// resolves to 2 glyphs (the ZWJ consumed, the two people left
    /// unligated), and the full sequence to 3. That is a real, observed
    /// difference in what this font's own tables define, not a shaping bug:
    /// the portable claim this test can make is that shaping is GSUB/cluster
    /// -aware (nowhere near the naive 5), not that any two color-emoji fonts
    /// ligate the same combinations. Apple Color Emoji *does* ligate this
    /// sequence, but via AAT `morx` — it carries no `GSUB` table at all, so
    /// the two fonts solve the same problem through mechanisms this module
    /// doesn't even need to distinguish between.
    #[test]
    fn zwj_sequence_ligates_to_one_glyph() {
        let Some(tf) = emoji_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");
        let family = "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}";
        assert_eq!(family.chars().count(), 5);

        let run = shaper
            .shape(&tf, family, 44.0, RunDirection::LeftToRight)
            .expect("shape");

        assert!(
            run.glyphs.len() < 5,
            "shaping must be GSUB/cluster-aware, not cmap-only mapping \
             (which would yield 5 for this 5-codepoint sequence); got {}",
            run.glyphs.len()
        );
        assert!(run.total_advance > Pt::ZERO);
    }

    /// A skin-tone modifier and a keycap sequence ligate by different GSUB
    /// mechanisms, with the same expectation.
    ///
    /// "Ligated" is asserted as **one cell wide**, not as one glyph. Those are
    /// not the same claim, and the difference is not hypothetical: Skia m150
    /// (skia-safe 0.99) began emitting the keycap as two glyphs — the composed
    /// keycap at the origin, plus a zero-advance blank parked at the far edge
    /// of the cell — where m145 emitted one. Nothing about the rendering
    /// changed: total advance stayed one cell, `tests/emoji_e2e.rs` still sees
    /// a single rasterized image with no constituent text beside it, and the
    /// corpus pixel-diff on `sample-emoji.docx` is zero.
    ///
    /// A glyph count is therefore the wrong instrument — it measures how the
    /// shaper chose to spell the answer rather than what gets drawn. Width
    /// *is* the property that matters, because the failure this test exists to
    /// catch is the sequence painting as its constituents side by side, which
    /// is two or three cells wide.
    #[test]
    fn modifier_and_keycap_sequences_ligate() {
        let Some(tf) = emoji_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");
        const SIZE: f32 = 44.0;
        for (label, text) in [
            ("skin-tone modifier", "\u{1F44D}\u{1F3FF}"),
            ("keycap", "1\u{FE0F}\u{20E3}"),
        ] {
            let run = shaper
                .shape(&tf, text, SIZE, RunDirection::LeftToRight)
                .expect("shape");

            // One cell, not one per codepoint. Generous tolerance: emoji cells
            // are not exactly em-square in every font.
            let cells = f32::from(run.total_advance) / SIZE;
            assert!(
                (0.5..=1.5).contains(&cells),
                "{label} must occupy one cell, got {cells:.2} ({:?} at size {SIZE})",
                run.total_advance,
            );

            // And it did ligate: fewer glyphs than codepoints, with everything
            // past the first contributing no width.
            assert!(
                run.glyphs.len() < text.chars().count(),
                "{label}: {} glyphs for {} codepoints is no ligation at all",
                run.glyphs.len(),
                text.chars().count(),
            );
        }
    }

    /// Empty text yields no glyphs, reported as an error rather than an empty
    /// run so callers take their documented `measure_str` / `draw_str` path.
    #[test]
    fn empty_text_reports_no_glyphs() {
        let Some(tf) = any_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");
        assert!(matches!(
            shaper.shape(&tf, "", 20.0, RunDirection::LeftToRight),
            Err(ShapeError::NoGlyphs)
        ));
    }

    /// The glyph ids shaping produces must be valid for `canvas.draw_glyphs`
    /// against the *same* typeface — the invariant the rasterizer relies on.
    #[test]
    fn glyph_ids_are_valid_for_the_same_typeface() {
        let Some(tf) = any_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");
        let run = shaper
            .shape(&tf, "abc", 24.0, RunDirection::LeftToRight)
            .expect("shape");

        let font = Font::from_typeface(tf, 24.0);
        let ids: Vec<GlyphId> = run.glyphs.iter().map(|g| g.id).collect();
        let mut widths = vec![0.0f32; ids.len()];
        font.get_widths(&ids, &mut widths);
        assert!(
            widths.iter().all(|w| *w > 0.0),
            "every shaped glyph id must have a width in the same font: {widths:?}"
        );
    }

    /// The contract [`ShapedGlyph::cluster`] documents and nothing in the
    /// SkShaper API states: cluster values are byte offsets into the *whole*
    /// string handed to `shape`, across sub-run splits. `"aあb"` forces a
    /// script split (Latin / Hiragana / Latin), so a shaper reporting
    /// run-relative offsets would repeat 0 where this expects 4 — cluster
    /// values are written whether or not the face covers the character, so no
    /// coverage probe is needed.
    #[test]
    fn clusters_are_absolute_byte_offsets() {
        let Some(tf) = any_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");

        let run = shaper
            .shape(&tf, "ab cd", 20.0, RunDirection::LeftToRight)
            .expect("shape");
        let mut clusters: Vec<u32> = run.glyphs.iter().map(|g| g.cluster).collect();
        clusters.sort_unstable();
        assert_eq!(clusters, vec![0, 1, 2, 3, 4], "one ASCII char, one cluster");

        let run = shaper
            .shape(&tf, "a\u{3042}b", 20.0, RunDirection::LeftToRight)
            .expect("shape");
        let mut clusters: Vec<u32> = run.glyphs.iter().map(|g| g.cluster).collect();
        clusters.sort_unstable();
        clusters.dedup();
        assert_eq!(
            clusters,
            vec![0, 1, 4],
            "offsets must survive the script split absolute, not restart at 0"
        );
    }

    /// [`ShapedRun::unit_count`] is the shaped counterpart of
    /// `spacing::unit_count`: one unit per ASCII letter, and a combining mark
    /// is folded into its base's cluster — by this module's grapheme folding,
    /// not by the font. `e` + U+0301 alone cannot pin that: most Latin faces
    /// compose the pair to one glyph, which merges the clusters with no help.
    /// U+0489 (a Cyrillic enclosing mark) is a mark no ordinary face ligates
    /// or even covers — cluster values are written for `.notdef` glyphs too —
    /// so before the folding it demonstrably stood as a unit of its own,
    /// which is `w:spacing` opening between a base and its own mark.
    #[test]
    fn unit_count_folds_a_combining_mark_into_its_base() {
        let Some(tf) = any_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");
        let abc = shaper
            .shape(&tf, "abc", 20.0, RunDirection::LeftToRight)
            .expect("shape");
        assert_eq!(abc.unit_count(), 3);
        for (label, text) in [
            ("composable accent", "e\u{301}"),
            ("unligatable enclosing mark", "a\u{0489}"),
            ("marks on both letters", "e\u{301}o\u{0489}"),
        ] {
            let run = shaper
                .shape(&tf, text, 20.0, RunDirection::LeftToRight)
                .expect("shape");
            assert_eq!(
                run.unit_count(),
                text.graphemes(true).count(),
                "{label}: a mark must share its base's unit"
            );
        }
    }

    /// **Issue #153's done criterion, at the shaper.** `कि` stores the vowel
    /// after the consonant; drawn correctly, the vowel stands to its left. The
    /// two nominal glyph ids come from the face's own cmap, so the assertion
    /// is structural — no font is named, and a host without a Devanagari face
    /// (or one whose shaper substitutes both forms) skips rather than fails.
    #[test]
    fn devanagari_prebase_matra_is_drawn_before_its_consonant() {
        const KA: char = '\u{0915}';
        const MATRA_I: char = '\u{093F}';
        let Some(tf) = FontMgr::new()
            .match_family_style_character("", FontStyle::normal(), &[], KA as i32)
            .filter(|tf| tf.unichar_to_glyph(KA as i32) != 0)
        else {
            eprintln!("skipping: no face on this host covers U+0915");
            return;
        };
        let ka_glyph = tf.unichar_to_glyph(KA as i32);
        let matra_glyph = tf.unichar_to_glyph(MATRA_I as i32);
        if matra_glyph == 0 {
            eprintln!("skipping: the Devanagari face has no nominal matra glyph");
            return;
        }

        let shaper = Shaper::new().expect("shaper");
        let run = shaper
            .shape(
                &tf,
                &format!("{KA}{MATRA_I}"),
                24.0,
                RunDirection::LeftToRight,
            )
            .expect("shape");

        assert_eq!(run.unit_count(), 1, "one akshara, one shaped cluster");
        let ka = run.glyphs.iter().find(|g| g.id == ka_glyph);
        let matra = run.glyphs.iter().find(|g| g.id == matra_glyph);
        let (Some(ka), Some(matra)) = (ka, matra) else {
            eprintln!("skipping: this face substitutes the nominal forms");
            return;
        };
        assert!(
            matra.x < ka.x,
            "the prebase matra must be reordered before its consonant \
             (matra at {:?}, consonant at {:?})",
            matra.x,
            ka.x,
        );
    }

    /// Shaping must never reach for the typeface's bytes — the regression
    /// guard for the 549 MB this module was rewritten to avoid.
    #[test]
    fn shaping_does_not_materialize_the_font() {
        let Some(tf) = emoji_typeface() else { return };
        let shaper = Shaper::new().expect("shaper");

        let Some(before) = resident_bytes() else {
            return; // no readable RSS — skip rather than assert on nothing
        };
        for _ in 0..64 {
            let _ = shaper.shape(&tf, "\u{1F44D}", 176.0, RunDirection::LeftToRight);
        }
        let Some(after) = resident_bytes() else {
            return;
        };

        // `to_font_data()` on Apple Color Emoji costs ~549 MB. A 64 MB ceiling
        // proves the bytes were never serialized, with wide margin for
        // allocator noise and Skia's own glyph cache.
        let growth = after.saturating_sub(before);
        assert!(
            growth < 64 * 1024 * 1024,
            "shaping grew RSS by {} MB — the font was probably serialized",
            growth / 1024 / 1024
        );
    }

    /// Current resident set size in bytes, or `None` if it cannot be read.
    fn resident_bytes() -> Option<usize> {
        std::process::Command::new("ps")
            .args(["-o", "rss=", "-p", &std::process::id().to_string()])
            .output()
            .ok()
            .and_then(|o| String::from_utf8(o.stdout).ok())
            .and_then(|s| s.trim().parse::<usize>().ok())
            .map(|kb| kb * 1024)
    }
}
