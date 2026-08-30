//! Fragment conversion — transform Inline content into measured Fragments
//! for the line-fitting algorithm.

use std::rc::Rc;

use crate::model::{PTabAlignment, PTabRelativeTo, RunProperties, TabLeader, UnderlineStyle};

use crate::i18n::bidi::BidiLevel;
use crate::render::dimension::Pt;
use crate::render::emoji::cluster::{EmojiPresentation, EmojiStructure};
use crate::render::fonts::Toggle;
use crate::render::fonts::TypefaceEntry;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::resolve::color::RgbColor;
use crate::render::resolve::fonts::effective_font;
use crate::render::resolve::images::MediaEntry;
use crate::render::shape::RunDirection;

mod bidi;
mod collect;
mod fallback;
mod segment;
mod shape;
mod split;
mod text;

pub use bidi::assign_bidi_levels;
pub use collect::{
    collect_fragments, FieldContext, FootnoteTracker, FragmentCtx, RecordedFootnote,
};
pub use fallback::{apply_font_fallback, FallbackLookup, RegistryFallback};
pub use shape::shape_complex_scripts;
pub use split::split_oversized_fragments;

// ── Superscript / subscript rendering constants ───────────────────────────────
// §17.3.2.42: these ratios are "application-defined" per the spec; the values
// below match Word's rendering as documented in the OpenXML SDK reference.

/// Font size of super/subscript text as a fraction of the base font size.
/// Also the size of a note reference mark and its body number (§17.11.12).
pub(crate) const SUPERSCRIPT_FONT_SIZE_RATIO: f32 = 0.58;

/// Superscript baseline shift: fraction of base ascent to raise the text by.
pub(super) const SUPERSCRIPT_ASCENT_OFFSET_RATIO: f32 = 0.33;

/// Subscript baseline shift: fraction of base character height to lower the text by.
pub(super) const SUBSCRIPT_HEIGHT_OFFSET_RATIO: f32 = 0.08;

/// §17.11.12: baseline shift for a footnote/endnote reference mark, and for
/// the matching number prefixed to the note body — as a fraction of the base
/// **font size**.
///
/// Deliberately *not* [`SUPERSCRIPT_ASCENT_OFFSET_RATIO`]: that one is a
/// fraction of the measured *ascent* and carries a different value (0.33).
/// Note marks are raised relative to the font size so the mark and its body
/// number line up without a measurement round-trip.
pub(crate) const NOTE_REF_BASELINE_OFFSET_RATIO: f32 = 0.4;

/// Font properties needed for rendering a text fragment.
#[derive(Clone, Debug)]
pub struct FontProps {
    pub family: Rc<str>,
    pub size: Pt,
    /// §17.3.2.1 `w:b` as the §17.7.2 cascade left it — absent, explicitly off,
    /// or on. Carried as a tri-state rather than a `bool` all the way to face
    /// selection; see [`crate::render::fonts::request`] for why the difference
    /// between "absent" and "off" is what lets a face name keep its own weight.
    pub bold: Toggle,
    /// §17.3.2.16 `w:i`, likewise.
    pub italic: Toggle,
    pub underline: bool,
    /// §17.3.2.30 `w:rtl` as the §17.7.2 cascade left it.
    ///
    /// The *input* to UAX #9 level resolution, where [`Fragment::Text`]'s
    /// `level` is the output: `layout::fragment::bidi` turns a run whose toggle
    /// is [`Toggle::On`] into an RLI…PDI isolate around that run's text in the
    /// analysis string. Tri-state because §17.7.2 toggles are, and because the
    /// three states genuinely differ here — [`Toggle::Absent`] leaves UAX #9's
    /// own rules to decide the run's direction from its characters, while
    /// [`Toggle::Off`] states a left-to-right context that neighbouring
    /// right-to-left text must not leak into.
    ///
    /// On `FontProps` rather than on the fragment because it is per *run*, and
    /// this is the per-run resolved presentation every text fragment already
    /// shares by `Rc` — the same reason `underline` and `char_spacing` are
    /// here without being properties of a font either.
    pub rtl: Toggle,
    pub char_spacing: Pt,
    /// §17.3.2.45: horizontal character scale as a multiplier (1.0 = normal,
    /// 0.8 = 80%, 1.5 = 150%). Applied to glyph advances during measure and
    /// to the Skia font's `scale_x` during paint. Inter-character spacing
    /// (`char_spacing`) is **not** scaled by this — the spec keeps the two
    /// independent.
    pub text_scale: f32,
    /// Underline position from font metrics (positive = below baseline).
    pub underline_position: Pt,
    /// Underline thickness from font metrics.
    pub underline_thickness: Pt,
    /// §17.3.2.37 `w:strike` / §17.3.2.9 `w:dstrike`: how many lines to draw
    /// through the run — 0, 1, or 2. The model's tri-state (`None` inherit,
    /// `Some(StrikeStyle::None)` explicit off, `Some(Single | Double)`)
    /// collapses here into a count, exactly as `underline` collapses into a
    /// bool above: by this seam the cascade has already answered.
    pub strike_lines: u8,
    /// Strike position from font metrics (positive = above baseline — the
    /// opposite sign of `underline_position`, matching where each line goes).
    pub strike_position: Pt,
    /// Strike thickness from font metrics.
    pub strike_thickness: Pt,
}

/// Font metrics for a specific font at a specific size.
/// Evaluated once by the measurer and carried through the pipeline.
#[derive(Clone, Copy, Debug)]
pub struct TextMetrics {
    /// Distance from baseline to top of glyphs (positive upward).
    pub ascent: Pt,
    /// Distance from baseline to bottom of glyphs (positive downward).
    pub descent: Pt,
    /// §17.3.1.33: inter-line leading from the font's metrics.
    /// Included in Auto line spacing base but not in glyph height.
    pub leading: Pt,
}

impl TextMetrics {
    /// Glyph height (ascent + descent) — used for baseline positioning.
    pub fn height(&self) -> Pt {
        self.ascent + self.descent
    }

    /// §17.3.1.33: full line height including leading — the base unit
    /// that Auto line spacing multipliers scale.
    pub fn line_height(&self) -> Pt {
        self.ascent + self.descent + self.leading
    }
}

/// §17.3.2.4: run-level border for rendering.
#[derive(Clone, Copy, Debug)]
pub struct FragmentBorder {
    pub width: Pt,
    pub color: RgbColor,
    pub space: Pt,
}

/// The target of a hyperlink carried on a text fragment. Keeps the
/// §17.16.22 external-vs-internal distinction (from `HyperlinkTarget`) as a
/// closed ADT so the emitter routes each to the right PDF annotation
/// (external → URI action, internal → GoTo a named destination) instead of
/// re-deriving it from a URL-scheme string check.
///
/// The string is shared, for the same reason [`Fragment::Text`]'s `text` and
/// `font` are: a `w:hyperlink` fragments into one `Fragment::Text` per *word*,
/// each of which then emits its own annotation command. Owning the target
/// would copy the URL once per word and again per command; sharing it copies
/// once per link.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum LinkTarget {
    /// A resolved external URI (`http:`, `mailto:`, `file:`, …).
    External(Rc<str>),
    /// An internal bookmark name (`w:hyperlink/@w:anchor`).
    Internal(Rc<str>),
}

/// UAX #14 line-break status at the *end* of a [`Fragment::Text`] — whether
/// the line fitter may put the next fragment on a new line.
///
/// Carried as data rather than re-derived from the fragment's last character,
/// which is what [`fit_lines`](crate::render::layout::line::fit_lines) used to
/// do. Two copies of "may a line break here?" is exactly what issue #130 found
/// had drifted: the cutter broke after U+2012 FIGURE DASH and the fitter's
/// character list did not include it. [`crate::i18n::segment`] is the one
/// place that answers the question now; this enum is how the answer travels.
///
/// Two states, not three. UAX #14 also has *mandatory* breaks ([LB4]/[LB5]),
/// and none can reach a fragment: CR, LF and every other C0 control but TAB
/// are stripped while the text is collected, and an authored `<w:br/>`
/// (§17.3.3.1) becomes [`Fragment::LineBreak`] instead of text.
///
/// [LB4]: https://www.unicode.org/reports/tr14/#LB4
/// [LB5]: https://www.unicode.org/reports/tr14/#LB5
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BreakAfter {
    /// [LD3]: a line may break between this fragment and the next.
    ///
    /// [LD3]: https://www.unicode.org/reports/tr14/#LD3
    Opportunity,
    /// No opportunity — the next fragment continues the same unbreakable
    /// unit. Either UAX #14 prohibits the break (`ID-001` between `-` and
    /// `0`), or the fragment ends at a `<w:r>` boundary that falls inside a
    /// token and the run after it carries the rest.
    Prohibited,
}

/// A measured fragment — the atomic unit for line fitting.
#[derive(Clone, Debug)]
pub enum Fragment {
    Text {
        text: Rc<str>,
        /// Shared per run: all words of a run carry the same font properties,
        /// so an `Rc` keeps the `Fragment::Text` variant small (a pointer, not
        /// an embedded ~48-byte `FontProps`) and makes per-word clones a
        /// refcount bump.
        font: Rc<FontProps>,
        color: RgbColor,
        /// §17.3.2.32: run-level shading (background color behind text).
        shading: Option<RgbColor>,
        /// §17.3.2.4: run-level border (box around text).
        border: Option<FragmentBorder>,
        /// UAX #14: whether a line may break after this fragment. See
        /// [`BreakAfter`] — this is the only thing line fitting consults, so
        /// a fragment that must stay joined to the next one says so here
        /// rather than relying on what its text happens to end with.
        break_after: BreakAfter,
        /// UAX #9: the resolved embedding level of this fragment's text.
        ///
        /// The companion to `break_after`, and carried for the same reason:
        /// the answer is a paragraph-scope one — rules W1–W7 and N0–N2 resolve
        /// a neutral from strong characters that may be far away — while the
        /// question is asked per line, once the breaks are known. `layout::
        /// fragment::bidi` resolves it once and splits any fragment a level
        /// boundary falls inside, so that by the time
        /// [`line_emit`](crate::render::layout::paragraph) reorders a line
        /// every fragment on it has exactly one level and the reorder is a
        /// permutation.
        ///
        /// [`BidiLevel::LTR`] on every fragment of a document with no
        /// bidirectional text, which is what lets the reorder be skipped
        /// outright there.
        level: BidiLevel,
        /// Whether this run must be shaped to be legible, and which way round
        /// — `None` for every run in a script that a cmap lookup renders
        /// correctly, which is all of Latin, Cyrillic, Greek, CJK, Hebrew and
        /// Thai.
        ///
        /// Set by `layout::fragment::shape`, which is also what re-measures
        /// the fragment against the shaped advance, so `width` below and what
        /// the painter draws cannot disagree. Carried rather than re-derived
        /// at paint from the text, because a second copy of
        /// [`needs_shaping`](crate::render::shape::needs_shaping) is the kind
        /// of duplicate rule issue #130 was spent removing.
        shaped: Option<RunDirection>,
        /// Full width including trailing whitespace (used for positioning).
        width: Pt,
        /// Width excluding trailing whitespace (used for line-break overflow checking).
        /// Trailing whitespace is allowed to hang past the margin per Word behavior.
        trimmed_width: Pt,
        /// Font metrics (ascent + descent = text height).
        metrics: TextMetrics,
        /// Hyperlink target (external URI or internal bookmark), if this
        /// fragment is inside a `w:hyperlink`. Named `hyperlink_url` for
        /// historical reasons; carries the external/internal kind, not a bare
        /// URL, so the emitter never has to guess from the string.
        hyperlink_url: Option<LinkTarget>,
        baseline_offset: Pt,
        /// Horizontal offset for drawing text within the fragment width.
        /// Used for right/center-justified list labels where the text is
        /// positioned within a wider fragment. Default: Pt::ZERO.
        text_offset: Pt,
        /// §17.11.12: true if this is a footnote reference mark (the superscript
        /// number). Rendered as ordinary text, but tagged so across-page
        /// splitting can reserve each footnote on the page its mark lands on.
        is_footnote_ref: bool,
    },
    Image {
        size: PtSize,
        rel_id: String,
        image_data: Option<MediaEntry>,
        /// §20.1.10.48 `a:srcRect` — fractional source crop in `[0, 1]`.
        src_rect: Option<PtRect>,
    },
    /// One emoji grapheme cluster (UAX #29) classified as an emoji sequence
    /// (UTS #51), to be rasterized at paint time via Skia's raster backend
    /// and embedded as an inline PDF image, because Skia's PDF backend strips
    /// the color glyph tables its raster backend honours. Classified by
    /// `render::emoji::cluster`; becomes a `DrawCommand::EmojiCluster`.
    Emoji {
        /// Cluster text exactly as classified — one grapheme cluster, possibly
        /// multi-codepoint (ZWJ, modifier, RIS, tag, keycap sequences).
        text: String,
        /// Color emoji typeface resolved upstream by the emoji resolver.
        /// Frozen at fragment build so paint never re-resolves.
        typeface: TypefaceEntry,
        /// Font size at which to rasterize, in Pt.
        size: Pt,
        /// UTS #51 §2 presentation. `EmojiPresentation::Text` is preserved
        /// (the rasterizer can still render it via the same color path) but
        /// allows future paint-side decisions (e.g. monochrome over color).
        presentation: EmojiPresentation,
        /// UTS #51 §2 cluster structure. Carried for diagnostics + future
        /// painter behaviour (skin-tone modifier substitution, etc.).
        structure: EmojiStructure,
        /// Measured advance from Skia raster metrics at `size`.
        advance: Pt,
        /// Font metrics from the resolved emoji typeface. Drives the
        /// rasterized image's natural aspect ratio and the rect's vertical
        /// extent in `line_emit::emit_line_commands` — NOT the line-height
        /// contribution. Color emoji typefaces (Apple Color Emoji, Segoe UI
        /// Emoji) carry tall ascents (≈1.25× font size) so their glyph art
        /// fits, but bumping running-text line height by that amount makes
        /// emoji-mixed lines visibly taller than text-only lines.
        metrics: TextMetrics,
        /// Metrics for line-height contribution, derived from the run's
        /// font.size against the run-level typeface (not the emoji
        /// typeface). Keeps the inline emoji "1em-tall" semantics so a
        /// paragraph that mixes emoji and plain text lays out evenly.
        /// The rasterized image still draws at its natural extent and may
        /// overhang the line slightly.
        line_metrics: TextMetrics,
        /// Inherited from the run (super/subscript / `w:position`).
        baseline_offset: Pt,
    },
    Tab {
        line_height: Pt,
        /// §17.3.1.38: formatting of the run holding the `<w:tab/>`. A tab
        /// leader carries no formatting of its own — it is drawn in the
        /// formatting in effect at the tab — so the leader emitter reads its
        /// family and size from here rather than substituting a default.
        font: Rc<FontProps>,
        /// §17.3.1.38: text colour of the tab's run, for the same reason.
        color: RgbColor,
        /// Override minimum width for line fitting (default: MIN_TAB_WIDTH).
        fitting_width: Option<Pt>,
    },
    /// §17.3.1.30: absolute-position tab. Its resolved position depends on the
    /// line's geometry (margins / indents) and following content, so — like
    /// [`Fragment::Tab`] — it occupies only a nominal width during line
    /// fitting and is placed during line emission.
    PTab {
        align: PTabAlignment,
        relative_to: PTabRelativeTo,
        leader: TabLeader,
        line_height: Pt,
        /// §17.3.1.38: formatting of the run holding the `<w:ptab/>` — the
        /// leader is drawn in it. See [`Fragment::Tab`].
        font: Rc<FontProps>,
        /// §17.3.1.38: text colour of the ptab's run.
        color: RgbColor,
    },
    LineBreak {
        line_height: Pt,
    },
    /// §17.3.3.1: column break — forces content to the next column.
    ColumnBreak,
    /// §17.3.3.1: page break — forces content to the next page.
    PageBreak {
        line_height: Pt,
    },
    /// Named destination (bookmark target) — zero-width marker.
    Bookmark {
        name: String,
    },
}

impl Fragment {
    pub fn width(&self) -> Pt {
        match self {
            Fragment::Text { width, .. } => *width,
            Fragment::Image { size, .. } => size.width,
            Fragment::Emoji { advance, .. } => *advance,
            Fragment::Tab { fitting_width, .. } => fitting_width.unwrap_or(MIN_TAB_WIDTH),
            Fragment::PTab { .. } => MIN_TAB_WIDTH,
            Fragment::LineBreak { .. }
            | Fragment::ColumnBreak
            | Fragment::PageBreak { .. }
            | Fragment::Bookmark { .. } => Pt::ZERO,
        }
    }

    /// Width for overflow checking — excludes trailing whitespace on text fragments.
    pub fn trimmed_width(&self) -> Pt {
        match self {
            Fragment::Text { trimmed_width, .. } => *trimmed_width,
            other => other.width(),
        }
    }

    pub fn height(&self) -> Pt {
        match self {
            Fragment::Text { metrics, .. } => metrics.height(),
            Fragment::Image { size, .. } => size.height,
            Fragment::Emoji { line_metrics, .. } => line_metrics.height(),
            Fragment::Tab { line_height, .. }
            | Fragment::PTab { line_height, .. }
            | Fragment::LineBreak { line_height }
            | Fragment::PageBreak { line_height } => *line_height,
            Fragment::ColumnBreak | Fragment::Bookmark { .. } => Pt::ZERO,
        }
    }

    pub fn is_line_break(&self) -> bool {
        matches!(
            self,
            Fragment::LineBreak { .. } | Fragment::ColumnBreak | Fragment::PageBreak { .. }
        )
    }

    /// UAX #9: the embedding level to reorder this fragment at, in a paragraph
    /// whose base level is `base`.
    ///
    /// Only text carries a resolved level. Everything else takes the base, and
    /// the approximation is deliberate rather than pending: an image or emoji
    /// contributes U+FFFC — class ON, a neutral — to the analysis string, so
    /// the text *around* it resolves against it correctly, and the base level
    /// is where a neutral standing alone lands anyway. What it gets wrong is an
    /// inline image between two runs of the *other* direction, which would need
    /// the field on three more variants to fix. Tabs never reach a reorder at
    /// all — `line_emit::visual_order` segments the line at them.
    pub fn bidi_level(&self, base: BidiLevel) -> BidiLevel {
        match self {
            Fragment::Text { level, .. } => *level,
            _ => base,
        }
    }

    /// §17.3.3.1: true if this fragment is a page break that forces
    /// subsequent content to the next page.
    pub fn is_page_break(&self) -> bool {
        matches!(self, Fragment::PageBreak { .. })
    }

    /// True if this fragment puts something on a line.
    ///
    /// Written as an exhaustive match rather than a `matches!` with a
    /// wildcard: a new variant should have to state which side it is on, since
    /// getting it wrong silently adds or drops a blank line.
    ///
    /// `Bookmark` is on the false side and is the one that is easy to miss —
    /// it is a navigation anchor with no glyphs and no advance, so a paragraph
    /// holding nothing but bookmarks still shows nothing.
    pub fn occupies_line(&self) -> bool {
        match self {
            Fragment::PageBreak { .. } | Fragment::ColumnBreak | Fragment::Bookmark { .. } => false,
            Fragment::Text { .. }
            | Fragment::Image { .. }
            | Fragment::Emoji { .. }
            | Fragment::Tab { .. }
            | Fragment::PTab { .. }
            | Fragment::LineBreak { .. } => true,
        }
    }

    /// Get font properties if this is a text fragment.
    pub fn font_props(&self) -> Option<&FontProps> {
        match self {
            Fragment::Text { font, .. } => Some(font),
            _ => None,
        }
    }
}

/// Whether a paragraph's mark (¶) needs a line made for it.
///
/// §17.3.1.29: every paragraph occupies at least one line, because the mark
/// itself has a height, and it keeps that line when the paragraph has nothing
/// to show. What varies is whether some content already put a line there.
///
/// The subtlety, and the whole of issue #126, is that "has nothing to show" is
/// not "is empty". `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` has a run and
/// a fragment, so an `is_empty()` test passes it over; it still draws nothing,
/// and without a line it collapses to zero height and cannot be moved to the
/// next page at all. A bookmark is the same story with no run properties to
/// notice: a navigation anchor with no glyphs and no advance.
///
/// This asks about the paragraph as a whole rather than about the segment
/// after its last break, because the line `build::block` makes is offered to
/// the page the paragraph *starts* on. Two orderings were measured against
/// `ELH_2025-12-18.docx` and that one is right; `build::block` records the
/// numbers at the injection site.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkLine {
    /// Something in the paragraph draws, so the mark rides the last line that
    /// content produces and there is nothing to add.
    RidesContent,
    /// Nothing in the paragraph draws. The mark needs a line of its own.
    NeedsOwnLine,
}

impl MarkLine {
    /// Decide from a paragraph's fragments.
    pub fn of(fragments: &[Fragment]) -> Self {
        if fragments.iter().any(Fragment::occupies_line) {
            MarkLine::RidesContent
        } else {
            MarkLine::NeedsOwnLine
        }
    }
}

/// The [`BreakAfter`] a unit-test fixture means when it writes a word with a
/// trailing space.
///
/// Fixtures all over the layout tests build `Fragment::Text` literals by hand
/// and spell their break opportunities as trailing whitespace, because until
/// issue #130 that *was* the encoding — `fit_lines` read the last character.
/// Restating that reading once, here, keeps every one of those fixtures
/// meaning what it meant and every test outcome comparable across the change.
/// Test-only on purpose: this is the rule the production path no longer has,
/// and nothing outside `#[cfg(test)]` may call it.
#[cfg(test)]
pub(crate) fn fixture_break_after(text: &str) -> BreakAfter {
    let breaks = text.ends_with([' ', '\t', '-', '\u{2010}', '\u{2013}', '\u{2014}']);
    if breaks {
        BreakAfter::Opportunity
    } else {
        BreakAfter::Prohibited
    }
}

/// §17.3.1.37: minimum tab fragment width for line fitting.
/// Tabs resolve to tab stops defined on the paragraph; this constant is only
/// used as the fragment width during line breaking (actual tab position is
/// computed during paragraph layout).
pub const MIN_TAB_WIDTH: Pt = Pt::new(1.0);

/// Extract font properties from RunProperties with a default font family fallback.
///
/// `auto_fit` is the §20.1.2.1.18 `a:normAutofit` shrink of the enclosing shape
/// text body, applied to whichever size wins — the run's own or the inherited
/// default — because the scale is a property of the *body*, not of any run in
/// it. Every caller outside a shape text box passes
/// [`ShapeAutoFit::NONE`](crate::render::layout::ShapeAutoFit::NONE); it is a
/// parameter rather than a default so that a new call site has to say which it
/// is.
pub fn font_props_from_run(
    rp: &RunProperties,
    default_family: &str,
    default_size: Pt,
    auto_fit: crate::render::layout::ShapeAutoFit,
) -> FontProps {
    let family = effective_font(&rp.fonts).unwrap_or(default_family);

    let size = auto_fit.scale_font(rp.font_size.cloned().map(Pt::from).unwrap_or(default_size));

    let char_spacing = rp.spacing.cloned().map(Pt::from).unwrap_or(Pt::ZERO);

    let text_scale = rp.text_scale.cloned().map_or(1.0, |s| s.as_factor());

    FontProps {
        family: Rc::from(family),
        size,
        // The model already carries all three §17.7.2 states; this used to be
        // the single line that threw two of them away.
        bold: Toggle::from_option(rp.bold),
        italic: Toggle::from_option(rp.italic),
        // §17.3.2.40: an actual underline style sets the bool. The model's
        // tri-state — `None` (inherit), `Some(UnderlineStyle::None)`
        // (explicit "no underline" override), `Some(_actual_style_)` —
        // collapses here into "draw / don't draw"; only the third case
        // draws.
        underline: matches!(rp.underline.get(), Some(s) if *s != UnderlineStyle::None),
        // §17.3.2.30. Kept tri-state all the way to level resolution, for the
        // same reason `bold` and `italic` are kept tri-state all the way to
        // face selection: what the cascade left absent is not what it turned
        // off, and only `layout::fragment::bidi` is entitled to decide what
        // absent means.
        rtl: Toggle::from_option(rp.rtl),
        char_spacing,
        text_scale,
        // §17.3.2.37 / §17.3.2.9: `w:dstrike` already won over `w:strike` at
        // parse (`resolve_strike`), so the cascade hands one answer here;
        // `Some(StrikeStyle::None)` is the explicit "no strike" override and
        // draws nothing, mirroring `UnderlineStyle::None` above.
        strike_lines: match rp.strike {
            Some(crate::model::StrikeStyle::Single) => 1,
            Some(crate::model::StrikeStyle::Double) => 2,
            Some(crate::model::StrikeStyle::None) | None => 0,
        },
        // Populated by the measurer from Skia font metrics.
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
        strike_position: Pt::ZERO,
        strike_thickness: Pt::ZERO,
    }
}

/// Convert a number to lowercase Roman numerals.
pub fn to_roman_lower(mut n: u32) -> String {
    const VALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut s = String::new();
    for &(val, sym) in &VALS {
        while n >= val {
            s.push_str(sym);
            n -= val;
        }
    }
    s
}

#[cfg(test)]
mod mark_line_tests {
    use super::*;

    fn page_break() -> Fragment {
        Fragment::PageBreak {
            line_height: Pt::new(12.0),
        }
    }

    fn line_break() -> Fragment {
        Fragment::LineBreak {
            line_height: Pt::new(12.0),
        }
    }

    fn bookmark() -> Fragment {
        Fragment::Bookmark {
            name: "anchor".into(),
        }
    }

    /// Stands in for anything that draws. An image needs no font, so the test
    /// says "content is present" without dragging a `FontProps` in.
    fn drawn() -> Fragment {
        Fragment::Image {
            size: PtSize::new(Pt::new(10.0), Pt::new(10.0)),
            rel_id: "rId1".into(),
            image_data: None,
            src_rect: None,
        }
    }

    /// A paragraph with no content at all: the mark is all there is.
    #[test]
    fn an_empty_paragraph_needs_a_line() {
        assert_eq!(MarkLine::of(&[]), MarkLine::NeedsOwnLine);
    }

    /// Issue #126's shape, and the one this rule exists for. The mark is the
    /// only thing in the paragraph after the break.
    #[test]
    fn a_break_only_paragraph_needs_a_line() {
        assert_eq!(MarkLine::of(&[page_break()]), MarkLine::NeedsOwnLine);
        assert_eq!(
            MarkLine::of(&[Fragment::ColumnBreak]),
            MarkLine::NeedsOwnLine
        );
    }

    /// Content already puts a line on the page and the mark rides its last one.
    #[test]
    fn drawn_content_carries_the_mark() {
        assert_eq!(MarkLine::of(&[drawn()]), MarkLine::RidesContent);
    }

    /// Content anywhere in the paragraph carries the mark, on whichever side
    /// of a break it sits. Injecting a line for `[text, break]` would put a
    /// blank one *above* the text, since the injection goes at the front.
    #[test]
    fn content_carries_the_mark_on_either_side_of_a_break() {
        assert_eq!(
            MarkLine::of(&[drawn(), page_break()]),
            MarkLine::RidesContent
        );
        assert_eq!(
            MarkLine::of(&[page_break(), drawn()]),
            MarkLine::RidesContent
        );
        assert_eq!(
            MarkLine::of(&[page_break(), Fragment::ColumnBreak, drawn()]),
            MarkLine::RidesContent
        );
    }

    /// Several breaks and still nothing drawn: one line, not one per break.
    #[test]
    fn repeated_breaks_still_need_exactly_one_line() {
        assert_eq!(
            MarkLine::of(&[page_break(), page_break(), Fragment::ColumnBreak]),
            MarkLine::NeedsOwnLine
        );
    }

    /// A bookmark is a navigation anchor with no glyphs and no advance, so a
    /// paragraph carrying only bookmarks and a break still shows nothing. The
    /// old predicate asked "is every fragment a break?" and a bookmark made it
    /// false, silently losing the mark's line.
    #[test]
    fn bookmarks_do_not_carry_the_mark() {
        assert_eq!(MarkLine::of(&[bookmark()]), MarkLine::NeedsOwnLine);
        assert_eq!(
            MarkLine::of(&[bookmark(), page_break(), bookmark()]),
            MarkLine::NeedsOwnLine
        );
    }

    /// Idempotence, and it is load-bearing rather than tidy: the injected line
    /// is itself a `LineBreak`, so a rule that did not count it would inject a
    /// second one every time this ran again over its own output.
    #[test]
    fn an_already_injected_line_is_not_injected_twice() {
        assert_eq!(
            MarkLine::of(&[page_break(), line_break()]),
            MarkLine::RidesContent
        );
        assert_eq!(MarkLine::of(&[line_break()]), MarkLine::RidesContent);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Dup;
    use crate::model::UnderlineStyle;
    use crate::render::fonts::Toggle;

    #[test]
    fn font_props_default_fallback() {
        let rp = RunProperties::default();
        let fp = font_props_from_run(
            &rp,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert_eq!(&*fp.family, "Helvetica");
        assert_eq!(fp.size.raw(), 12.0);
        assert_eq!(fp.bold, Toggle::Absent);
        assert_eq!(fp.italic, Toggle::Absent);
    }

    // ── §17.3.2.1 / §17.3.2.16 bold and italic tri-state ───────────────────
    //
    // This is the seam issue #85 was about. `RunProperties::{bold, italic}`
    // carry all three §17.7.2 states out of the cascade, and this function used
    // to be the single line that threw two of them away with
    // `.unwrap_or(false)` — after which resolution could not tell a document
    // that wants Regular from one that said nothing at all.

    #[test]
    fn all_three_cascade_states_reach_font_props() {
        let at = |bold: Option<bool>, italic: Option<bool>| {
            font_props_from_run(
                &RunProperties {
                    bold,
                    italic,
                    ..RunProperties::default()
                },
                "Helvetica",
                Pt::new(12.0),
                crate::render::layout::ShapeAutoFit::NONE,
            )
        };

        let absent = at(None, None);
        assert_eq!(absent.bold, Toggle::Absent);
        assert_eq!(absent.italic, Toggle::Absent);

        // The state a `bool` could not express: the run explicitly declined the
        // toggle, which is not the same as never having been asked.
        let off = at(Some(false), Some(false));
        assert_eq!(off.bold, Toggle::Off);
        assert_eq!(off.italic, Toggle::Off);

        let on = at(Some(true), Some(true));
        assert_eq!(on.bold, Toggle::On);
        assert_eq!(on.italic, Toggle::On);

        // …and the two toggles are independent.
        let mixed = at(Some(true), Some(false));
        assert_eq!(mixed.bold, Toggle::On);
        assert_eq!(mixed.italic, Toggle::Off);
    }

    // ── §17.3.2.40 underline tri-state ─────────────────────────────────────
    //
    // `RunProperties::underline: Option<UnderlineStyle>` carries three states:
    //   * `None`                            — element absent; inherit (§17.7.2)
    //   * `Some(UnderlineStyle::None)`      — `<w:u w:val="none"/>` explicit override
    //   * `Some(UnderlineStyle::Single)` …  — actual underline style
    // `font_props.underline` is the rendering-decision boolean: it must be
    // `true` only when an actual underline style is in effect.

    fn rp_with_underline(style: Option<UnderlineStyle>) -> RunProperties {
        RunProperties {
            underline: Dup::from(style),
            ..RunProperties::default()
        }
    }

    #[test]
    fn font_props_underline_absent_is_false() {
        let fp = font_props_from_run(
            &rp_with_underline(None),
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(!fp.underline, "no <w:u> element → no underline");
    }

    #[test]
    fn font_props_underline_explicit_none_is_false() {
        let fp = font_props_from_run(
            &rp_with_underline(Some(UnderlineStyle::None)),
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(
            !fp.underline,
            "<w:u w:val=\"none\"/> is the spec's explicit \"no underline\" \
             override; font_props.underline must remain false"
        );
    }

    #[test]
    fn font_props_underline_single_is_true() {
        let fp = font_props_from_run(
            &rp_with_underline(Some(UnderlineStyle::Single)),
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(fp.underline, "<w:u w:val=\"single\"/> → underline drawn");
    }

    #[test]
    fn font_props_text_scale_default_is_one() {
        // §17.3.2.45: when <w:w> is absent the run renders at 100% width.
        let fp = font_props_from_run(
            &RunProperties::default(),
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert_eq!(fp.text_scale, 1.0);
    }

    #[test]
    fn font_props_text_scale_compressed() {
        // <w:w w:val="80"/> → 0.8× horizontal scale.
        let rp = RunProperties {
            text_scale: Dup::from(Some(crate::model::TextScale::new(80))),
            ..RunProperties::default()
        };
        let fp = font_props_from_run(
            &rp,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!((fp.text_scale - 0.8).abs() < f32::EPSILON);
    }

    #[test]
    fn font_props_text_scale_expanded() {
        // <w:w w:val="150"/> → 1.5× horizontal scale.
        let rp = RunProperties {
            text_scale: Dup::from(Some(crate::model::TextScale::new(150))),
            ..RunProperties::default()
        };
        let fp = font_props_from_run(
            &rp,
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!((fp.text_scale - 1.5).abs() < f32::EPSILON);
    }

    #[test]
    fn font_props_underline_double_is_true() {
        // Sanity: any non-`None` style sets the bool. A future renderer
        // change to support distinct styles will replace this bool with
        // an enum; for now, "any style other than None" → draw.
        let fp = font_props_from_run(
            &rp_with_underline(Some(UnderlineStyle::Double)),
            "Helvetica",
            Pt::new(12.0),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert!(fp.underline);
    }
}
