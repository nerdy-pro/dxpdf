//! Line fitting — break fragments into lines that fit within a max width.
//!
//! Also the one place automatic hyphenation (§17.15.1.10, issue #158) acts:
//! a word is cut only once it is known not to fit, which only fitting knows.
//! `crate::i18n::hyphenate` answers *where* a word may be cut; this module
//! decides *whether*, from the §17.15.1.53 zone, the §17.15.1.22 consecutive
//! limit and §17.15.1.27, and splices the two halves into the paragraph's
//! fragment vector so emission paints them as ordinary text.

use std::borrow::Cow;
use std::cell::Cell;

use crate::i18n::hyphenate::HyphenLanguage;
use crate::render::dimension::Pt;
use crate::render::layout::fragment::{BreakAfter, FontProps, Fragment, TextMetrics};

/// A fitted line — a slice of fragments that fit within the available width.
#[derive(Debug)]
pub struct FittedLine {
    /// Indices into the fragment list: [start, end).
    pub start: usize,
    pub end: usize,
    /// Total width of all fragments in this line.
    pub width: Pt,
    /// Maximum height of any fragment in this line.
    pub height: Pt,
    /// Maximum height of text-only fragments on this line.
    /// §17.3.1.33: Auto line spacing multiplier applies to text metrics,
    /// not to inline image heights.
    pub text_height: Pt,
    /// Maximum ascent of any text fragment in this line.
    pub ascent: Pt,
    /// Whether this line ends with an explicit line break.
    pub has_break: bool,
}

/// Where fitting starts in the fragment vector and how far it runs.
///
/// `start` is an index into the *whole* vector, so the lines come back in
/// the vector's own index space — the floats path in
/// `paragraph::line_emit::compute_line_placements` fits one line per call
/// and used to slice the vector and remap, which hyphenation forbids: a
/// cut word is spliced into the vector, and a slice cannot be spliced into.
/// `max_lines` is that same path's "one line, please": fitting past the
/// first line would hyphenate words the caller is about to re-fit at a
/// different width, leaving their inserted hyphens behind.
#[derive(Clone, Copy, Debug)]
pub struct FitWindow {
    /// Index of the first fragment to fit.
    pub start: usize,
    /// Stop once this many lines exist. `None` fits to the end.
    pub max_lines: Option<usize>,
}

impl FitWindow {
    /// The whole vector, every line.
    pub const ALL: FitWindow = FitWindow {
        start: 0,
        max_lines: None,
    };
}

/// What the fitter needs to hyphenate (§17.15.1.10, issue #158) — present
/// only for a paragraph the document and the §17.7.2 cascade allow it for,
/// and only when a real measurer exists; `None` is the path every document
/// without `w:autoHyphenation` takes, and it is bit-for-bit the old one.
///
/// `measure` and `opportunities` are hooks rather than direct calls so the
/// unit tests below can drive the zone and limit rules with a fixed-advance
/// measurer and a table of opportunities instead of pattern data.
pub struct HyphenationCtx<'a> {
    /// §17.15.1.53: the gap at the line end a word is allowed to leave before
    /// hyphenation is attempted on it.
    pub(crate) zone: Pt,
    /// §17.15.1.22: how many lines in a row may end in an inserted hyphen;
    /// `0` is no limit.
    pub(crate) consecutive_limit: u32,
    /// §17.15.1.27 `w:doNotHyphenateCaps`.
    pub(crate) no_caps: bool,
    /// Width of a candidate head (`"hyphen-"`) in the fragment's font — the
    /// same closure that measured the fragment, so the head's width is what
    /// the painter will draw.
    pub(crate) measure: &'a dyn Fn(&str, &FontProps) -> (Pt, TextMetrics),
    /// `crate::i18n::hyphenate::opportunities` in production.
    pub(crate) opportunities: &'a dyn Fn(&str, HyphenLanguage) -> Vec<usize>,
    /// Lines ended by an inserted hyphen so far, for `consecutive_limit`. A
    /// `Cell` because the floats path fits one line per call and the count
    /// has to survive between calls. Starts at 0 for every `place_paragraph`,
    /// so a paragraph split across pages restarts its count on the
    /// continuation — accepted: the clause counts lines, not pages, but a
    /// page break between two hyphenated lines is the case least worth a
    /// field on `PlacedParagraph`.
    pub(crate) consecutive: Cell<u32>,
}

/// Break fragments into lines that fit within `max_width`.
///
/// When a line overflows, it breaks at the last fragment that reported a UAX
/// #14 break opportunity ([`BreakAfter::Opportunity`]). A single fragment wider
/// than `max_width` gets its own line (no infinite loop); over-wide *text* is
/// cut into grapheme clusters upstream by
/// [`split_oversized_fragments`](crate::render::layout::fragment::split_oversized_fragments)
/// before it reaches here.
///
/// `first_line_width`: if provided, the first line uses this narrower width
/// (e.g., to account for first-line indent). Subsequent lines use `max_width`.
///
/// Never hyphenates: this is [`fit_lines_in`] with no [`HyphenationCtx`].
pub fn fit_lines(fragments: &[Fragment], max_width: Pt) -> Vec<FittedLine> {
    fit_lines_with_first(
        fragments,
        max_width,
        max_width,
        crate::render::layout::paragraph::PTabGeometry {
            max_width,
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: max_width,
            float_left: Pt::ZERO,
            float_right: Pt::ZERO,
        },
        false,
    )
}

/// Line fitting with separate first-line and remaining-line widths, over a
/// borrowed slice and without hyphenation — [`fit_lines_in`] with
/// [`FitWindow::ALL`] and no [`HyphenationCtx`]. The signature every test
/// and every caller had before issue #158.
pub fn fit_lines_with_first(
    fragments: &[Fragment],
    first_line_width: Pt,
    remaining_width: Pt,
    ptab_geometry: crate::render::layout::paragraph::PTabGeometry,
    bidi_rtl: bool,
) -> Vec<FittedLine> {
    fit_lines_in(
        &mut Cow::Borrowed(fragments),
        FitWindow::ALL,
        first_line_width,
        remaining_width,
        ptab_geometry,
        bidi_rtl,
        None,
    )
}

/// Line fitting with separate first-line and remaining-line widths.
///
/// `ptab_geometry` is the paragraph geometry §17.3.1.30 position tabs resolve
/// against. Fitting needs it because a tab whose alignment point lies behind
/// the pen advances to the next line, and only fitting can create one.
///
/// `bidi_rtl`: whether the paragraph's base direction is right-to-left.
/// Emission mirrors a tab-bearing RTL line (§17.3.1.37 under `w:bidi`) and, in
/// that mirrored frame, swaps which physical side each float narrows before
/// resolving a `<w:ptab>` — fitting has to swap the same way, or the two can
/// reach different Placed/AdvancesToNextLine verdicts for the same tab. A
/// line's tab-placement precondition for mirroring is already guaranteed
/// once fitting has *reached* a `PTab` fragment — whatever line ends up
/// holding it will satisfy `w:bidi`'s other half — so `bidi_rtl` alone is the
/// right test here, without also threading a per-line "has a tab" flag
/// fitting cannot know in advance.
///
/// `fragments` is a `Cow` because hyphenation (§17.15.1.10) is the one thing
/// in fitting that adds a glyph: a word that overflows is replaced in place by
/// its head (`"hyphen-"`, [`BreakAfter::Hyphenated`]) and its tail, and the
/// returned line indices are into the vector *after* that splice. With no
/// `hyphenation` context the borrow is never materialised and the returned
/// lines are bit-for-bit those of the pre-#158 fitter.
pub fn fit_lines_in(
    fragments: &mut Cow<'_, [Fragment]>,
    window: FitWindow,
    first_line_width: Pt,
    remaining_width: Pt,
    ptab_geometry: crate::render::layout::paragraph::PTabGeometry,
    bidi_rtl: bool,
    hyphenation: Option<&HyphenationCtx<'_>>,
) -> Vec<FittedLine> {
    if window.start >= fragments.len() {
        return Vec::new();
    }
    let reached = |lines: &[FittedLine]| window.max_lines.is_some_and(|n| lines.len() >= n);

    let mut lines = Vec::new();
    let mut line_start = window.start;
    let mut line_width = Pt::ZERO;
    // §17.3.1.30: where the pen actually *is*, as opposed to how much width
    // the line has accumulated. The two differ only across a position tab,
    // which jumps the pen to its anchor while contributing a nominal width.
    // Tracking it separately keeps `line_width` — and therefore every
    // non-ptab paragraph's fitting — bit-for-bit unchanged.
    let line_pen_start = |is_first_line: bool| {
        ptab_geometry.indent_left
            + if is_first_line {
                ptab_geometry.indent_first_line
            } else {
                Pt::ZERO
            }
    };
    let mut pen_x = line_pen_start(true);
    // §17.3.1.30: `relativeTo="margin"` measures against the full text area,
    // so once such a tab has placed content this line may legitimately use the
    // space a paragraph's own right indent excludes. Until then the ordinary
    // `content_width` bound applies.
    let mut margin_span_active = false;
    let mut line_height = Pt::ZERO;
    let mut line_text_height = Pt::ZERO;
    let mut line_ascent = Pt::ZERO;
    let mut last_break_point = None; // index after which we can break

    let mut i = window.start;
    while i < fragments.len() {
        let frag = &fragments[i];

        // Explicit line break — emit current line including the break fragment.
        // LineBreak height already includes leading (from default_line_height).
        if frag.is_line_break() {
            line_height = line_height.max(frag.height());
            line_text_height = line_text_height.max(frag.height());
            lines.push(FittedLine {
                start: line_start,
                end: i + 1,
                width: line_width,
                height: line_height,
                text_height: line_text_height,
                ascent: line_ascent,
                has_break: true,
            });
            if let Some(ctx) = hyphenation {
                ctx.consecutive.set(0);
            }
            if reached(&lines) {
                return lines;
            }
            line_start = i + 1;
            line_width = Pt::ZERO;
            pen_x = line_pen_start(lines.is_empty());
            margin_span_active = false;
            line_height = Pt::ZERO;
            line_text_height = Pt::ZERO;
            line_ascent = Pt::ZERO;
            last_break_point = None;
            i += 1;
            continue;
        }

        // §17.3.1.30: a position tab whose alignment point lies behind the
        // pen advances to that point on the *next* line. Decided here rather
        // than at emission because it is a line break, and emission cannot
        // create one.
        if let Fragment::PTab {
            align, relative_to, ..
        } = frag
        {
            // Fitting cannot bound the zone by the line end the way emission
            // does — the lines do not exist yet — so it scans to the end of the
            // fragment list. Where the two disagree (a zone fitting then splits
            // for width) emission is authoritative for the final x; this only
            // decides whether the tab can be honoured on this line.
            let end = crate::render::layout::paragraph::zone_end(fragments, i, fragments.len());
            // Emission mirrors a tab-bearing RTL line and swaps which
            // physical side each float narrows before resolving this same
            // tab (`line_emit.rs`'s `Fragment::PTab` arm) — matched here so
            // the two agree on whether the anchor is behind the pen. See
            // `bidi_rtl`'s doc for why that flag alone is the right test.
            let (geo_float_left, geo_float_right) = if bidi_rtl {
                (ptab_geometry.float_right, ptab_geometry.float_left)
            } else {
                (ptab_geometry.float_left, ptab_geometry.float_right)
            };
            let geometry = crate::render::layout::paragraph::PTabGeometry {
                float_left: geo_float_left,
                float_right: geo_float_right,
                ..ptab_geometry
            };
            let placement = crate::render::layout::paragraph::resolve_ptab(
                *align,
                *relative_to,
                geometry,
                pen_x,
                || crate::render::layout::paragraph::zone_width(fragments, i + 1, end),
            );
            match placement {
                crate::render::layout::paragraph::PTabPlacement::Placed(at) => {
                    pen_x = at;
                    margin_span_active |=
                        matches!(relative_to, crate::model::PTabRelativeTo::Margin);
                }
                crate::render::layout::paragraph::PTabPlacement::AdvancesToNextLine { .. } => {
                    // Only break when doing so can help. A tab already first
                    // on its line would find the same anchor behind the same
                    // pen on the next one — acting on a condition the action
                    // cannot change is how this engine's pagination loops have
                    // historically become infinite.
                    if line_start < i {
                        let m = measure_range(fragments, line_start, i);
                        lines.push(FittedLine {
                            start: line_start,
                            end: i,
                            width: m.width,
                            height: m.height,
                            text_height: m.text_height,
                            ascent: m.ascent,
                            has_break: false,
                        });
                        if let Some(ctx) = hyphenation {
                            ctx.consecutive.set(0);
                        }
                        if reached(&lines) {
                            return lines;
                        }
                        line_start = i;
                        line_width = Pt::ZERO;
                        pen_x = line_pen_start(lines.is_empty());
                        margin_span_active = false;
                        line_height = Pt::ZERO;
                        line_text_height = Pt::ZERO;
                        line_ascent = Pt::ZERO;
                        last_break_point = None;
                        // Re-evaluate this tab against the fresh line.
                        continue;
                    }
                }
            }
        }

        let frag_width = frag.width();
        let new_width = line_width + frag_width;

        // Use first-line width for line 0, remaining width for subsequent lines.
        let current_max = if lines.is_empty() {
            first_line_width
        } else {
            remaining_width
        };

        // For overflow checking, use trimmed width — trailing whitespace on the
        // last word is allowed to hang past the margin (standard Word behavior).
        // The check uses: previous fragments' full widths + this fragment's trimmed width.
        let check_width = line_width + frag.trimmed_width();

        // Check if adding this fragment overflows. Once a margin-relative tab
        // has placed content, the line's real right edge is the margin, and
        // the pen — not the accumulated width sum — says where we are.
        let overflows = if margin_span_active {
            pen_x + frag.trimmed_width() > ptab_geometry.max_width
        } else {
            check_width > current_max
        };
        if overflows && line_start < i {
            // §17.15.1.10: before falling back to the last opportunity, try
            // to cut the word that did not fit. Skipped while a
            // margin-relative position tab governs the line (§17.3.1.30):
            // there the bound is the pen against the margin, not
            // `current_max`, and a hyphenated line inside a tab zone is not a
            // case any fixture asks for.
            if let Some(ctx) = hyphenation.filter(|_| !margin_span_active) {
                if let Some((head, tail)) =
                    hyphenate_at(fragments, line_start, i, line_width, current_max, ctx)
                {
                    // The splice: the word becomes its head, and its tail is
                    // inserted after it. Everything already fitted sits
                    // before `i`, so no line's indices move; everything
                    // after `i` shifts by one and is fitted from the shifted
                    // vector, which is why `fragments` is a `Cow` and not a
                    // slice.
                    let owned = fragments.to_mut();
                    owned[i] = head;
                    owned.insert(i + 1, tail);
                    let m = measure_range(fragments, line_start, i + 1);
                    lines.push(FittedLine {
                        start: line_start,
                        end: i + 1,
                        width: m.width,
                        height: m.height,
                        text_height: m.text_height,
                        ascent: m.ascent,
                        has_break: false,
                    });
                    ctx.consecutive.set(ctx.consecutive.get() + 1);
                    if reached(&lines) {
                        return lines;
                    }
                    line_start = i + 1;
                    line_width = Pt::ZERO;
                    pen_x = line_pen_start(lines.is_empty());
                    margin_span_active = false;
                    line_height = Pt::ZERO;
                    line_text_height = Pt::ZERO;
                    line_ascent = Pt::ZERO;
                    last_break_point = None;
                    // The tail opens the next line.
                    i += 1;
                    continue;
                }
            }
            // Overflow — break at last break point, or before this fragment.
            let break_at = last_break_point.unwrap_or(i);
            let m = measure_range(fragments, line_start, break_at);
            lines.push(FittedLine {
                start: line_start,
                end: break_at,
                width: m.width,
                height: m.height,
                text_height: m.text_height,
                ascent: m.ascent,
                has_break: false,
            });
            // §17.15.1.22: a line that ends without an inserted hyphen
            // restarts the consecutive count.
            if let Some(ctx) = hyphenation {
                ctx.consecutive.set(0);
            }
            if reached(&lines) {
                return lines;
            }
            line_start = break_at;
            line_width = Pt::ZERO;
            pen_x = line_pen_start(lines.is_empty());
            margin_span_active = false;
            line_height = Pt::ZERO;
            line_text_height = Pt::ZERO;
            line_ascent = Pt::ZERO;
            last_break_point = None;
            // Resume *at the break*, not at the fragment that overflowed.
            //
            // `break_at` is usually `i` — no earlier opportunity, so the new
            // line starts with the fragment that did not fit and this rewinds
            // nothing. When there *was* an earlier opportunity it can be far
            // behind, and then everything between it and `i` belongs to the
            // new line and has to be measured onto it. Resuming at `i` left
            // that span inside the line's `[start, end)` range — so it painted
            // — while contributing nothing to the width, after which the line
            // could not overflow again and swallowed the rest of the
            // paragraph. That is how a Windows path in a 167.80 pt footer cell
            // became one 295 pt line running 91 pt past the edge of the page.
            //
            // Termination is unchanged: `line_start` still only ever moves
            // forward, and the next pass has `i == line_start`, where the
            // `line_start < i` guard sends an over-wide fragment down the
            // "first on the line, allow it" path instead of breaking again.
            i = break_at;
            continue;
        }

        // If this is the first fragment on the line and it overflows,
        // allow it (it will be the only fragment on this line). The
        // paragraph renderer will clip/overflow as needed.
        line_width = new_width;
        // A position tab has already jumped the pen to its anchor; every other
        // fragment advances it by its own width. Done here, after the overflow
        // check, so the check sees the pen *at* this fragment rather than past it.
        if !matches!(frag, Fragment::PTab { .. }) {
            pen_x += frag_width;
        }
        line_height = line_height.max(frag.height());
        // §17.3.1.33: text_height is the Auto line spacing base — use
        // line_height() (includes leading) for text, glyph height for tabs.
        match frag {
            Fragment::Text { metrics, .. } | Fragment::MathFraction { metrics, .. } => {
                line_text_height = line_text_height.max(metrics.line_height());
                line_ascent = line_ascent.max(metrics.ascent);
            }
            Fragment::Image { .. } => {} // images don't contribute to text_height
            _ => {
                line_text_height = line_text_height.max(frag.height());
            }
        }

        // Track break opportunity. A text fragment carries UAX #14's answer
        // (`crate::i18n::segment` computed it over the whole paragraph text,
        // across `<w:r>` boundaries); this used to re-derive it here by
        // sniffing the fragment's last character, against a list that had
        // already drifted from the one used to cut the fragments. A
        // [`BreakAfter::Hyphenated`] head is an opportunity this engine made
        // for itself; it only reaches here as *input* on a re-fit, and the
        // paragraph layer un-hyphenates before any re-fit, so the arm is for
        // completeness rather than a path.
        let is_break_point = match frag {
            Fragment::Text { break_after, .. } | Fragment::MathFraction { break_after, .. } => {
                matches!(
                    break_after,
                    BreakAfter::Opportunity | BreakAfter::Hyphenated
                )
            }
            _ => true, // tabs, images, line breaks are always break points
        };
        if is_break_point {
            last_break_point = Some(i + 1);
        }

        i += 1;
    }

    // Emit remaining fragments as the last line.
    if line_start < fragments.len() {
        lines.push(FittedLine {
            start: line_start,
            end: fragments.len(),
            width: line_width,
            height: line_height,
            text_height: line_text_height,
            ascent: line_ascent,
            has_break: false,
        });
    }

    lines
}

/// §17.15.1.10: try to cut the word at `fragments[i]`, which overflowed the
/// line `[line_start, i)` of width `line_width` against `current_max`.
///
/// Returns the head (`"hyphen-"`, [`BreakAfter::Hyphenated`]) and tail
/// fragments to splice in place of the word, or `None` when the word is not
/// one this engine hyphenates, the §17.15.1.53 zone says to leave the gap,
/// the §17.15.1.22 limit is reached, or no head plus hyphen fits.
///
/// What is never cut, and why:
/// * a fragment with no `FontProps::hyphenation` — an untagged run, or a
///   language with no vendored patterns;
/// * a shaped fragment (`shaped.is_some()`) — its width is HarfBuzz's over the
///   whole run, and a head re-measured through the cmap path would be fitted
///   at one width and painted at another;
/// * a footnote reference mark (§17.11.12) — text, but not a word;
/// * a fragment glued to its neighbour by [`BreakAfter::Prohibited`] on either
///   side — a token split across `<w:r>` boundaries, a token UAX #14 keeps
///   whole (`ID‑001`), the per-cluster pieces of an over-wide word
///   `split_oversized_fragments` has already cut, or the first body word
///   after a §17.9 list label whose suffix glues to it. Neither piece is the
///   whole word, so the patterns cannot be asked about it; v2 would rejoin
///   the chain for the lookup;
/// * a token whose letters are not one unbroken run — digits, an apostrophe
///   or a hyphen inside it (see [`hyphenation_core`]);
/// * under §17.15.1.27, a word whose letters are all capitals.
///
/// The zone is compared against the gap the *visible* line leaves — trailing
/// whitespace hanging as it does for the overflow test itself — and the
/// comparison is strict, so `w:hyphenationZone w:val="0"` hyphenates on any
/// gap. Among the pattern points, the longest head whose `head + "-"` fits is
/// taken; the head is re-measured through the same closure that measured the
/// fragment, so its width is what the painter will draw.
///
/// **Word reference render.** Two things here are this engine's reading, not
/// a measured match. §17.15.1.53 defines the zone as the gap a line may leave
/// *before* hyphenation is attempted, and this is that reading (gap > zone →
/// attempt); its own example reads the other way, and Microsoft's VBA
/// documentation says Word 2013+ in native layout (`compatibilityMode` 15,
/// which this engine does not parse) ignores the zone altogether — so Word
/// may hyphenate *more* than this does, never less. And Word hyphenates from
/// its own dictionaries, so a word's cut points can differ from the TeX
/// patterns'; the fixtures use compounds both agree on.
fn hyphenate_at(
    fragments: &[Fragment],
    line_start: usize,
    i: usize,
    line_width: Pt,
    current_max: Pt,
    ctx: &HyphenationCtx<'_>,
) -> Option<(Fragment, Fragment)> {
    let Fragment::Text {
        text,
        font,
        color,
        shading,
        border,
        break_after,
        level,
        shaped: None,
        hyperlink_url,
        baseline_offset,
        text_offset,
        is_footnote_ref: false,
        ..
    } = &fragments[i]
    else {
        return None;
    };
    let lang = font.hyphenation?;
    if *break_after == BreakAfter::Prohibited || glues_to_next(&fragments[i - 1]) {
        return None;
    }
    if ctx.consecutive_limit != 0 && ctx.consecutive.get() >= ctx.consecutive_limit {
        return None;
    }
    let core = hyphenation_core(text)?;
    let word = &text[core.clone()];
    if ctx.no_caps && word.chars().all(char::is_uppercase) {
        return None;
    }
    let gap = current_max - visible_width(fragments, line_start, i);
    if gap <= ctx.zone {
        return None;
    }

    let opportunities = (ctx.opportunities)(word, lang);
    for &at in opportunities.iter().rev() {
        let cut = core.start + at;
        let head_text = format!("{}-", &text[..cut]);
        let (head_width, head_metrics) = (ctx.measure)(&head_text, font);
        if line_width + head_width > current_max {
            continue;
        }
        let tail_text = &text[cut..];
        let (tail_width, tail_metrics) = (ctx.measure)(tail_text, font);
        let trimmed = tail_text.trim_end();
        let tail_trimmed_width = if trimmed.len() < tail_text.len() {
            (ctx.measure)(trimmed, font).0
        } else {
            tail_width
        };
        let piece = |text: &str, width, trimmed_width, metrics, break_after| Fragment::Text {
            text: text.into(),
            font: font.clone(),
            color: *color,
            shading: *shading,
            border: *border,
            break_after,
            // Cutting a word never crosses a level boundary: `fragment::bidi`
            // has already split at every one, and the shaped case is excluded
            // above, so both halves are at the word's level.
            level: *level,
            shaped: None,
            width,
            trimmed_width,
            metrics,
            hyperlink_url: hyperlink_url.clone(),
            baseline_offset: *baseline_offset,
            text_offset: *text_offset,
            is_footnote_ref: false,
        };
        return Some((
            // The head has no trailing whitespace, so its trimmed width is
            // its width.
            piece(
                &head_text,
                head_width,
                head_width,
                head_metrics,
                BreakAfter::Hyphenated,
            ),
            piece(
                tail_text,
                tail_width,
                tail_trimmed_width,
                tail_metrics,
                *break_after,
            ),
        ));
    }
    None
}

/// Undo what [`fit_lines_in`] hyphenated in `fragments`: every
/// [`BreakAfter::Hyphenated`] head is rejoined with the tail after it — the
/// inserted U+002D dropped, the text concatenated, the width re-measured (or
/// summed, when no measurer is at hand) and the tail's own `break_after`
/// restored — so the vector reads as the paragraph's words again.
///
/// For the §4 continuation re-fit (`PlacedParagraph::fragments_from_line`):
/// the remainder of a split paragraph is re-placed at the next page's or
/// column's width, around that page's floats, and a word cut for *this*
/// width may fit whole there. Re-fitting the cut fragments would put
/// `hyphen-` and `ation` side by side on one line and paint a hyphen
/// mid-word. The `Hyphenated` marker is what makes the undo possible, which
/// is why the head carries a variant of its own rather than a flag on the
/// line.
pub(crate) fn rejoin_hyphenated(
    fragments: &mut Vec<Fragment>,
    measure: crate::render::layout::paragraph::MeasureTextFn<'_>,
) {
    let mut i = 0;
    while i + 1 < fragments.len() {
        let joined = match (&fragments[i], &fragments[i + 1]) {
            (
                Fragment::Text {
                    text: head,
                    break_after: BreakAfter::Hyphenated,
                    font,
                    color,
                    shading,
                    border,
                    level,
                    hyperlink_url,
                    baseline_offset,
                    text_offset,
                    metrics,
                    width: head_width,
                    ..
                },
                Fragment::Text {
                    text: tail,
                    break_after,
                    width: tail_width,
                    trimmed_width: tail_trimmed,
                    ..
                },
            ) => {
                let text = format!("{}{tail}", head.strip_suffix('-').unwrap_or(head));
                let (width, trimmed_width, metrics) = match measure {
                    Some(measure) => {
                        let (width, metrics) = measure(&text, font);
                        let trimmed = text.trim_end();
                        let trimmed_width = if trimmed.len() < text.len() {
                            measure(trimmed, font).0
                        } else {
                            width
                        };
                        (width, trimmed_width, metrics)
                    }
                    // Only reachable if a head was made without a measurer,
                    // which `place_paragraph` never does; the sum keeps the
                    // line arithmetic sane rather than exact.
                    None => (
                        *head_width + *tail_width,
                        *head_width + *tail_trimmed,
                        *metrics,
                    ),
                };
                Fragment::Text {
                    text: text.into(),
                    font: font.clone(),
                    color: *color,
                    shading: *shading,
                    border: *border,
                    break_after: *break_after,
                    level: *level,
                    shaped: None,
                    width,
                    trimmed_width,
                    metrics,
                    hyperlink_url: hyperlink_url.clone(),
                    baseline_offset: *baseline_offset,
                    text_offset: *text_offset,
                    is_footnote_ref: false,
                }
            }
            _ => {
                i += 1;
                continue;
            }
        };
        fragments[i] = joined;
        fragments.remove(i + 1);
        // Not advanced: the joined fragment is re-examined, so a chain of
        // heads would fold completely.
    }
}

/// Whether `fragment` glues the fragment after it to itself — the
/// [`BreakAfter::Prohibited`] half of the "whole word" test in
/// [`hyphenate_at`]. Only text and math fragments carry the state; every
/// other kind (tab, image, bookmark) is a break point and glues nothing.
fn glues_to_next(fragment: &Fragment) -> bool {
    matches!(
        fragment,
        Fragment::Text {
            break_after: BreakAfter::Prohibited,
            ..
        } | Fragment::MathFraction {
            break_after: BreakAfter::Prohibited,
            ..
        }
    )
}

/// The width of `[start, end)` as it will be seen: full widths, except that
/// the last fragment that paints anything contributes its `trimmed_width`,
/// its trailing whitespace hanging past the edge. The same rule
/// `paragraph::line_emit::visible_line_width` applies to a finished line, so
/// the §17.15.1.53 gap and the justification slack agree on how wide a line
/// is.
fn visible_width(fragments: &[Fragment], start: usize, end: usize) -> Pt {
    let last_visible = (start..end)
        .rev()
        .find(|&idx| !matches!(fragments[idx], Fragment::Bookmark { .. }));
    (start..end)
        .map(|idx| {
            if Some(idx) == last_visible {
                fragments[idx].trimmed_width()
            } else {
                fragments[idx].width()
            }
        })
        .sum()
}

/// The run of letters inside a token that the patterns are asked about: the
/// span from its first alphabetic character to its last. Leading punctuation
/// (an opening quote or bracket) is left to the head, trailing punctuation
/// and whitespace to the tail.
///
/// `None` for a token that is not a word in the patterns' sense: one with no
/// letters, one containing a digit anywhere (`ID‑001`, `3rd`, `v2`), or one
/// whose letters are interrupted — an apostrophe (`don't`), a hyphen
/// (`hard-nosed`, which UAX #14 already lets a line break after), a soft
/// hyphen that reached the text. The `hyphenation` crate's own guidance is
/// that punctuation is a segmentation question; this is the segmentation.
pub(crate) fn hyphenation_core(text: &str) -> Option<std::ops::Range<usize>> {
    if text.chars().any(char::is_numeric) {
        return None;
    }
    let start = text.find(char::is_alphabetic)?;
    let last = text.rfind(char::is_alphabetic)?;
    let end = last + text[last..].chars().next()?.len_utf8();
    text[start..end]
        .chars()
        .all(char::is_alphabetic)
        .then_some(start..end)
}

/// Measurements for a range of fragments.
struct RangeMeasure {
    width: Pt,
    height: Pt,
    text_height: Pt,
    ascent: Pt,
}

/// Measure total width, max height, text height, and ascent for a range of fragments.
fn measure_range(fragments: &[Fragment], start: usize, end: usize) -> RangeMeasure {
    let mut m = RangeMeasure {
        width: Pt::ZERO,
        height: Pt::ZERO,
        text_height: Pt::ZERO,
        ascent: Pt::ZERO,
    };
    for frag in &fragments[start..end] {
        m.width += frag.width();
        m.height = m.height.max(frag.height());
        match frag {
            Fragment::Text { metrics, .. } | Fragment::MathFraction { metrics, .. } => {
                m.text_height = m.text_height.max(metrics.line_height());
                m.ascent = m.ascent.max(metrics.ascent);
            }
            Fragment::Image { .. } => {}
            _ => {
                m.text_height = m.text_height.max(frag.height());
            }
        }
    }
    m
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::fonts::Toggle;
    use crate::render::layout::fragment::{FontProps, TextMetrics};
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    /// A fragment the fitter may break *after* — what a word followed by a
    /// space becomes once `crate::i18n::segment` has looked at it.
    fn text_frag(text: &str, width: f32) -> Fragment {
        frag(text, width, BreakAfter::Opportunity)
    }

    /// A fragment glued to whatever follows it: a token cut by a `<w:r>`
    /// boundary, or one UAX #14 refuses to break inside.
    fn glued_frag(text: &str, width: f32) -> Fragment {
        frag(text, width, BreakAfter::Prohibited)
    }

    /// A `Fragment::MathFraction` with a caller-chosen `break_after` — e.g.
    /// `(1/2)²`, where the fraction is a superscript's base and must glue to
    /// the exponent that follows it (`fragment::math`'s `emit_elements` sets
    /// `Prohibited` on exactly this fragment for that case).
    fn math_fraction_frag(width: f32, break_after: BreakAfter) -> Fragment {
        let row = |text: &str| crate::render::layout::fragment::MathRow {
            text: text.into(),
            font: Rc::new(FontProps {
                effects: Default::default(),
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
                hyphenation: None,
            }),
            width: Pt::new(width * 0.5),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
        };
        Fragment::MathFraction {
            num: row("1"),
            den: row("2"),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            baseline_offset: Pt::ZERO,
            break_after,
            hyperlink_url: None,
        }
    }

    fn frag(text: &str, width: f32, break_after: BreakAfter) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after,
            font: Rc::new(FontProps {
                effects: Default::default(),
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
                hyphenation: None,
            }),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    /// A `<w:ptab>` fragment — `line_height`/`leader`/`color` are irrelevant
    /// to fitting, which only reads `align`/`relative_to` (plus the nominal
    /// `MIN_TAB_WIDTH` `.width()` contributes to `line_width`).
    fn ptab_frag(align: crate::model::PTabAlignment) -> Fragment {
        Fragment::PTab {
            align,
            relative_to: crate::model::PTabRelativeTo::Indent,
            leader: crate::model::TabLeader::None,
            line_height: Pt::ZERO,
            font: Rc::new(FontProps {
                effects: Default::default(),
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
                hyphenation: None,
            }),
            color: RgbColor::BLACK,
        }
    }

    #[test]
    fn empty_fragments_no_lines() {
        let lines = fit_lines(&[], Pt::new(100.0));
        assert!(lines.is_empty());
    }

    #[test]
    fn single_fragment_fits() {
        let frags = vec![text_frag("hello", 30.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].start, 0);
        assert_eq!(lines[0].end, 1);
        assert_eq!(lines[0].width.raw(), 30.0);
        assert_eq!(lines[0].height.raw(), 14.0);
    }

    #[test]
    fn two_fragments_fit_on_one_line() {
        let frags = vec![text_frag("hello ", 35.0), text_frag("world", 30.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].end, 2);
        assert_eq!(lines[0].width.raw(), 65.0);
    }

    #[test]
    fn overflow_breaks_at_boundary() {
        let frags = vec![
            text_frag("hello ", 60.0),
            text_frag("world ", 60.0),
            text_frag("end", 30.0),
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].start, 0);
        assert_eq!(lines[0].end, 1); // "hello " on first line
        assert_eq!(lines[1].start, 1);
        assert_eq!(lines[1].end, 3); // "world " + "end" on second line
    }

    /// Breaking back to an earlier opportunity must re-measure everything
    /// between it and the fragment that overflowed.
    ///
    /// The fitter accumulates forward, and on overflow it rewinds to the last
    /// opportunity. That opportunity can be many fragments behind — here the
    /// space after `"Z:`, with an unbreakable run of clusters after it. The
    /// rewind moved `line_start` back but left the cursor where the overflow
    /// was found, so every fragment in between stayed inside the new line's
    /// range — painted — while contributing nothing to its width. The line
    /// then never overflowed again and swallowed the rest of the paragraph.
    ///
    /// This is issue-shaped rather than theoretical: it is why a Windows path
    /// in a 167.80 pt footer cell laid out as one 295 pt line, 91 pt past the
    /// right edge of the page.
    #[test]
    fn a_backward_break_re_measures_the_fragments_it_rewinds_over() {
        // 10 glued clusters of 10pt each behind one opportunity.
        let mut frags = vec![text_frag("head ", 10.0)];
        frags.extend((0..10).map(|_| glued_frag("x", 10.0)));

        let lines = fit_lines(&frags, Pt::new(50.0));

        // `head ` alone, then the unbreakable run cut into 50pt lines.
        assert_eq!(lines[0].start, 0);
        assert_eq!(
            lines[0].end, 1,
            "the opportunity after `head ` is the first break"
        );
        for line in &lines[1..] {
            let measured: f32 = frags[line.start..line.end]
                .iter()
                .map(|f| f.width().raw())
                .sum();
            assert!(
                measured <= 50.0,
                "line [{}..{}) holds {measured}pt of fragments in a 50pt line",
                line.start,
                line.end,
            );
            assert!(
                (line.width.raw() - measured).abs() < 0.01,
                "line [{}..{}) records width {} but its fragments measure {measured}",
                line.start,
                line.end,
                line.width.raw(),
            );
        }
        let covered: usize = lines.iter().map(|l| l.end - l.start).sum();
        assert_eq!(
            covered,
            frags.len(),
            "every fragment lands on exactly one line"
        );
    }

    /// A token UAX #14 refuses to break inside stays whole even though it is
    /// split across two fragments: `ID‑001` carries a non-breaking hyphen, so
    /// the piece before it says [`BreakAfter::Prohibited`] and the fitter has
    /// to go back to the space before `prefix`.
    #[test]
    fn a_prohibited_boundary_is_not_a_break_point() {
        let frags = vec![
            text_frag("prefix ", 45.0),
            glued_frag("ID‑", 20.0),
            glued_frag("001", 35.0),
        ];
        let lines = fit_lines(&frags, Pt::new(70.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1);
        assert_eq!(lines[1].start, 1);
        assert_eq!(lines[1].end, 3);
    }

    /// `Fragment::MathFraction` carries its own `break_after` (a fraction
    /// used as a superscript's base, `(1/2)²`, is glued to the exponent that
    /// follows it) — the fitter must honor that field exactly like a
    /// `Fragment::Text`'s, not fall into the wildcard arm that used to treat
    /// every non-text fragment as an unconditional break point regardless of
    /// what it carried. Same shape as the test above, with the middle
    /// fragment swapped for a glued `MathFraction`.
    #[test]
    fn a_prohibited_math_fraction_is_not_a_break_point() {
        let frags = vec![
            text_frag("prefix ", 45.0),
            math_fraction_frag(20.0, BreakAfter::Prohibited),
            text_frag("sup", 35.0),
        ];
        let lines = fit_lines(&frags, Pt::new(70.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1);
        assert_eq!(lines[1].start, 1);
        assert_eq!(lines[1].end, 3);
    }

    /// The fitter reads [`BreakAfter`] and nothing else. This fragment ends in
    /// a hyphen and a trailing space — every character the old rule broke on —
    /// and still does not break, because UAX #14 said not to. If this fails,
    /// the character sniff has grown back and `crate::i18n::segment` is no
    /// longer the only answer to the question.
    #[test]
    fn the_text_itself_no_longer_decides_where_a_line_breaks() {
        let frags = vec![
            glued_frag("0100- ", 40.0),
            text_frag("600", 40.0),
            text_frag("rest", 40.0),
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2, "{lines:#?}");
        assert_eq!(
            lines[0].end, 2,
            "both halves of the token stay together despite the hyphen — the \
             old character rule would have broken after it and put line 0 at 1",
        );
    }

    #[test]
    fn oversized_fragment_gets_own_line() {
        let frags = vec![text_frag("verylongword", 200.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1, "oversized fragment still produces a line");
        assert_eq!(lines[0].end, 1);
    }

    #[test]
    fn line_break_forces_new_line() {
        let frags = vec![
            text_frag("before", 30.0),
            Fragment::LineBreak {
                line_height: Pt::new(14.0),
            },
            text_frag("after", 25.0),
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 2); // "before" + line break
        assert!(lines[0].has_break);
        assert_eq!(lines[1].start, 2);
        assert_eq!(lines[1].end, 3); // "after"
    }

    #[test]
    fn exact_fit_no_overflow() {
        let frags = vec![text_frag("a", 50.0), text_frag("b", 50.0)];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].width.raw(), 100.0);
    }

    #[test]
    fn tab_uses_min_width_for_fitting() {
        let frags = vec![
            text_frag("text", 80.0),
            Fragment::Tab {
                line_height: Pt::new(14.0),
                font: Rc::new(FontProps {
                    effects: Default::default(),
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: Rc::from("Test"),
                    size: Pt::new(12.0),
                    bold: Toggle::Absent,
                    italic: Toggle::Absent,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    text_scale: 1.0,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                    hyphenation: None,
                }),
                color: RgbColor::BLACK,
                fitting_width: None,
            },
            text_frag("more", 30.0),
        ];
        // 80 + 12 (MIN_TAB_WIDTH) = 92, still fits 100
        // But + 30 = 122, doesn't fit
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 2);
    }

    #[test]
    fn height_is_max_of_fragments() {
        let frags = vec![
            Fragment::Text {
                shaped: None,
                level: crate::i18n::bidi::BidiLevel::LTR,
                text: "small".into(),
                break_after: BreakAfter::Prohibited,
                font: Rc::new(FontProps {
                    effects: Default::default(),
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: Rc::from("Test"),
                    size: Pt::new(10.0),
                    bold: Toggle::Absent,
                    italic: Toggle::Absent,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    text_scale: 1.0,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                    hyphenation: None,
                }),
                color: RgbColor::BLACK,
                width: Pt::new(20.0),
                trimmed_width: Pt::new(20.0),
                metrics: TextMetrics {
                    ascent: Pt::new(9.0),
                    descent: Pt::new(3.0),
                    leading: Pt::ZERO,
                },
                hyperlink_url: None,
                shading: None,
                border: None,
                baseline_offset: Pt::ZERO,
                text_offset: Pt::ZERO,
                is_footnote_ref: false,
            },
            Fragment::Text {
                shaped: None,
                level: crate::i18n::bidi::BidiLevel::LTR,
                text: "big".into(),
                break_after: BreakAfter::Prohibited,
                font: Rc::new(FontProps {
                    effects: Default::default(),
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: Rc::from("Test"),
                    size: Pt::new(24.0),
                    bold: Toggle::Absent,
                    italic: Toggle::Absent,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    text_scale: 1.0,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                    hyphenation: None,
                }),
                color: RgbColor::BLACK,
                width: Pt::new(30.0),
                trimmed_width: Pt::new(30.0),
                metrics: TextMetrics {
                    ascent: Pt::new(22.0),
                    descent: Pt::new(6.0),
                    leading: Pt::ZERO,
                },
                hyperlink_url: None,
                shading: None,
                border: None,
                baseline_offset: Pt::ZERO,
                text_offset: Pt::ZERO,
                is_footnote_ref: false,
            },
        ];
        let lines = fit_lines(&frags, Pt::new(100.0));

        assert_eq!(lines.len(), 1);
        assert_eq!(lines[0].height.raw(), 28.0, "max of 12 and 28");
        assert_eq!(lines[0].ascent.raw(), 22.0, "max of 9 and 22");
    }

    #[test]
    fn multiple_overflows_produce_multiple_lines() {
        let frags = vec![
            text_frag("a ", 40.0),
            text_frag("b ", 40.0),
            text_frag("c ", 40.0),
            text_frag("d ", 40.0),
            text_frag("e", 40.0),
        ];
        // max_width=70: "a " fits (40), +"b " = 80 > 70 → break
        let lines = fit_lines(&frags, Pt::new(70.0));

        assert!(lines.len() >= 3, "should produce at least 3 lines");
        // Each line should have at most 1 fragment since 40+40=80 > 70
    }

    #[test]
    fn first_line_narrower_than_remaining() {
        // first_line_width=60, remaining_width=100.
        // "a " (40pt) fits the narrow first line alone; "b " (40pt) + "c" (40pt)
        // = 80pt fit together on the full 100pt remaining line.
        let frags = vec![
            text_frag("a ", 40.0),
            text_frag("b ", 40.0),
            text_frag("c", 40.0),
        ];
        let lines = fit_lines_with_first(
            &frags,
            Pt::new(60.0),
            Pt::new(100.0),
            crate::render::layout::paragraph::PTabGeometry {
                max_width: Pt::new(100.0),
                indent_left: Pt::ZERO,
                indent_first_line: Pt::ZERO,
                content_width: Pt::new(100.0),
                float_left: Pt::ZERO,
                float_right: Pt::ZERO,
            },
            false,
        );
        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1, "only 'a ' fits the narrow first line");
        assert_eq!(lines[1].start, 1);
        assert_eq!(
            lines[1].end, 3,
            "'b ' + 'c' both fit on the full second line"
        );
    }

    // ── PTab mirroring under `w:bidi` ────────────────────────────────────────

    #[test]
    fn bidi_rtl_ptab_swaps_floats_the_same_way_emission_does() {
        // A 60pt float narrows the line's physical left; under `w:bidi` that
        // is the line's *trailing* edge, so a `Left`/`Indent` ptab's anchor
        // must be measured from the physical *right* (an unnarrowed 0), not
        // the physical left (60) — exactly the swap emission's mirrored pen
        // already applies. Getting this wrong is invisible unless the swap
        // actually changes which line the tab lands on, which is why the
        // anchor is placed behind a 10pt fragment: 0 is behind it, 60 is not.
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(crate::model::PTabAlignment::Left),
            text_frag("B", 10.0),
        ];
        let geometry = crate::render::layout::paragraph::PTabGeometry {
            max_width: Pt::new(100.0),
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: Pt::new(100.0),
            float_left: Pt::new(60.0),
            float_right: Pt::ZERO,
        };
        let lines = fit_lines_with_first(&frags, Pt::new(100.0), Pt::new(100.0), geometry, true);
        assert_eq!(
            lines.len(),
            2,
            "the ptab's anchor (0, after the float swap) is behind the pen \
             (10, after 'A'), so fitting must break before it — matching \
             emission's own mirrored verdict for the same tab"
        );
        assert_eq!(lines[0].end, 1, "'A' alone, then the line breaks");
        assert_eq!(lines[1].start, 1, "the ptab restarts the next line");
        assert_eq!(lines[1].end, 3, "'B' joins the ptab on that line");
    }

    #[test]
    fn ltr_ptab_is_unaffected_by_bidi_rtl() {
        // The same fragments and geometry as the mirrored case above, but
        // `bidi_rtl: false` — the float swap must not fire, so the anchor
        // (60, unswapped) sits ahead of the pen (10) and the whole paragraph
        // fits on one line, as it always did before this fix.
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(crate::model::PTabAlignment::Left),
            text_frag("B", 10.0),
        ];
        let geometry = crate::render::layout::paragraph::PTabGeometry {
            max_width: Pt::new(100.0),
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: Pt::new(100.0),
            float_left: Pt::new(60.0),
            float_right: Pt::ZERO,
        };
        let lines = fit_lines_with_first(&frags, Pt::new(100.0), Pt::new(100.0), geometry, false);
        assert_eq!(lines.len(), 1, "no break: the anchor is ahead of the pen");
        assert_eq!(lines[0].end, 3);
    }

    // ── §17.15.1.10 automatic hyphenation (issue #158) ────────────────────────
    //
    // Every test here measures in a fixed 6 pt per character, hyphen included,
    // and takes its hyphenation points from a table rather than pattern data,
    // so each one isolates a rule of the fitter: the §17.15.1.53 zone, the
    // longest-fitting-head choice, the §17.15.1.22 consecutive limit,
    // §17.15.1.27, and the word shapes that are never cut.

    use crate::render::layout::paragraph::PTabGeometry;

    /// 6 pt per character, hyphen included.
    fn six_per_char(text: &str, _: &FontProps) -> (Pt, TextMetrics) {
        (
            Pt::new(6.0 * text.chars().count() as f32),
            TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
        )
    }

    /// Stands in for `crate::i18n::hyphenate::opportunities`: en-US
    /// `hyphenation` as the crate documents it, `sample` as `sam-ple`, and
    /// nothing for any other word. Case-insensitive, as the real one is.
    fn table(word: &str, _: HyphenLanguage) -> Vec<usize> {
        match word.to_lowercase().as_str() {
            "hyphenation" => vec![2, 6, 7],
            "sample" => vec![3],
            _ => Vec::new(),
        }
    }

    fn ctx(zone: f32, consecutive_limit: u32, no_caps: bool) -> HyphenationCtx<'static> {
        HyphenationCtx {
            zone: Pt::new(zone),
            consecutive_limit,
            no_caps,
            measure: &six_per_char,
            opportunities: &table,
            consecutive: Cell::new(0),
        }
    }

    /// A word of an `en-US` run, measured by `six_per_char` (trailing space
    /// counted in `width`, not in `trimmed_width`).
    fn word(text: &str, break_after: BreakAfter) -> Fragment {
        let mut f = frag(text, 6.0 * text.chars().count() as f32, break_after);
        if let Fragment::Text {
            font,
            trimmed_width,
            ..
        } = &mut f
        {
            *trimmed_width = Pt::new(6.0 * text.trim_end().chars().count() as f32);
            let mut props = (**font).clone();
            props.hyphenation = Some(HyphenLanguage::EnUs);
            *font = Rc::new(props);
        }
        f
    }

    fn text_of(f: &Fragment) -> String {
        match f {
            Fragment::Text { text, .. } => text.to_string(),
            other => panic!("expected a text fragment, got {other:?}"),
        }
    }

    fn break_after_of(f: &Fragment) -> BreakAfter {
        match f {
            Fragment::Text { break_after, .. } => *break_after,
            other => panic!("expected a text fragment, got {other:?}"),
        }
    }

    fn geometry(width: f32) -> PTabGeometry {
        PTabGeometry {
            max_width: Pt::new(width),
            indent_left: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            content_width: Pt::new(width),
            float_left: Pt::ZERO,
            float_right: Pt::ZERO,
        }
    }

    /// Fit `frags` at `width`, returning each line's texts and the vector
    /// after any splice.
    fn fit(
        frags: Vec<Fragment>,
        width: f32,
        hyphenation: Option<&HyphenationCtx<'_>>,
    ) -> (Vec<Vec<String>>, Vec<Fragment>) {
        let mut cow: Cow<[Fragment]> = Cow::Owned(frags);
        let lines = fit_lines_in(
            &mut cow,
            FitWindow::ALL,
            Pt::new(width),
            Pt::new(width),
            geometry(width),
            false,
            hyphenation,
        );
        let frags = cow.into_owned();
        let texts = lines
            .iter()
            .map(|l| frags[l.start..l.end].iter().map(text_of).collect())
            .collect();
        (texts, frags)
    }

    /// §17.15.1.53: "the amount of whitespace which can be left at the end of
    /// a line … before hyphenation should be attempted on the next word". The
    /// gap `hyphenation` would leave (66 − 18 = 48 pt) exceeds the 18 pt zone,
    /// so the word is cut at the longest pattern point whose head plus hyphen
    /// fits: `hyphen-` (42 pt) does, `hyphena-` (48 pt) does not.
    #[test]
    fn a_word_that_overflows_is_hyphenated_when_the_gap_exceeds_the_zone() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, frags) = fit(frags, 66.0, Some(&ctx(18.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "hyphen-"], vec!["ation"]]);
        assert_eq!(break_after_of(&frags[1]), BreakAfter::Hyphenated);
        assert_eq!(
            frags[1].width().raw(),
            42.0,
            "head re-measured with its hyphen"
        );
        assert_eq!(frags[2].width().raw(), 30.0, "tail re-measured");
    }

    /// §17.15.1.53, the other half: a gap the zone permits is left, and the
    /// word moves down whole. Same geometry as above with a 60 pt zone.
    #[test]
    fn a_gap_inside_the_zone_wraps_the_word_whole() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, frags) = fit(frags, 66.0, Some(&ctx(60.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphenation"]]);
        assert_eq!(frags.len(), 2, "nothing spliced");
    }

    /// The gap is measured from the line's *visible* width, with the trailing
    /// space hanging as it does for overflow: 66 − 18 = 48 pt, which clears a
    /// 45 pt zone. Measured from the full width (66 − 24 = 42) it would not,
    /// and the word would move down whole.
    #[test]
    fn the_zone_is_measured_from_the_visible_width() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, _) = fit(frags, 66.0, Some(&ctx(45.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "hyphen-"], vec!["ation"]]);
    }

    /// `w:hyphenationZone w:val="0"` (which `test-files/document_outline.docx`
    /// declares) means hyphenate on any gap at all; and the comparison is
    /// strict — a gap exactly the zone is still "whitespace which can be
    /// left". 30 pt of gap here: cut at zone 0, left at zone 30.
    #[test]
    fn a_zero_zone_hyphenates_on_any_gap_and_the_bound_is_strict() {
        let frags = || {
            vec![
                word("the ", BreakAfter::Opportunity),
                word("quick ", BreakAfter::Opportunity),
                word("sample", BreakAfter::Opportunity),
            ]
        };
        let (lines, _) = fit(frags(), 84.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "quick ", "sam-"], vec!["ple"]]);
        let (lines, _) = fit(frags(), 84.0, Some(&ctx(30.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "quick "], vec!["sample"]]);
    }

    /// Among the heads that fit, the longest wins: at 72 pt `hyphena-`
    /// (48 pt) fits exactly beside `the ` (24 pt).
    #[test]
    fn the_longest_head_that_fits_wins() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, _) = fit(frags, 72.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "hyphena-"], vec!["tion"]]);
    }

    /// When even the shortest head plus hyphen (`hy-`, 18 pt) will not fit
    /// in the 16 pt left, the word moves down whole — a hyphenation attempt
    /// that fails changes nothing.
    #[test]
    fn no_fitting_head_wraps_the_word_whole() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, frags) = fit(frags, 40.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphenation"]]);
        assert_eq!(frags.len(), 2);
    }

    /// §17.15.1.22: "Once this limit has been reached, the following line
    /// shall not be hyphenated regardless of whether or not it meets the
    /// criteria". Limit 1: the first line ends in a hyphen, so the second may
    /// not even though `hy-` would fit beside `ation `. Control: limit 0 is
    /// no limit and the second line is cut too.
    #[test]
    fn the_consecutive_limit_stops_the_next_line() {
        let frags = || {
            vec![
                word("the ", BreakAfter::Opportunity),
                word("hyphenation ", BreakAfter::Opportunity),
                word("hyphenation", BreakAfter::Opportunity),
            ]
        };
        let (lines, _) = fit(frags(), 66.0, Some(&ctx(0.0, 1, false)));
        assert_eq!(
            lines,
            [vec!["the ", "hyphen-"], vec!["ation "], vec!["hyphenation"]]
        );
        let (lines, _) = fit(frags(), 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(
            lines,
            [
                vec!["the ", "hyphen-"],
                vec!["ation ", "hy-"],
                vec!["phenation"]
            ]
        );
    }

    /// The limit counts *consecutive* lines: a line that ends without an
    /// inserted hyphen restarts the count. Limit 1: line 1 is cut, line 2
    /// may not be and breaks before `x `, and that resets the count so line 3
    /// is cut again. A counter that never reset would leave `hyphenation`
    /// whole on line 4.
    #[test]
    fn an_unhyphenated_line_resets_the_count() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation ", BreakAfter::Opportunity),
            word("word ", BreakAfter::Opportunity),
            word("x ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, _) = fit(frags, 66.0, Some(&ctx(0.0, 1, false)));
        assert_eq!(
            lines,
            [
                vec!["the ", "hyphen-"],
                vec!["ation ", "word "],
                vec!["x ", "hyphena-"],
                vec!["tion"],
            ]
        );
    }

    /// §17.15.1.27: with `w:doNotHyphenateCaps` a word "comprised of all
    /// capital letters" moves down whole; "If omitted, words in CAPITAL
    /// LETTERS shall be hyphenated" is the control.
    #[test]
    fn an_all_caps_word_is_left_whole_when_caps_are_excluded() {
        let frags = || {
            vec![
                word("the ", BreakAfter::Opportunity),
                word("HYPHENATION", BreakAfter::Opportunity),
            ]
        };
        let (lines, _) = fit(frags(), 66.0, Some(&ctx(0.0, 0, true)));
        assert_eq!(lines, [vec!["the "], vec!["HYPHENATION"]]);
        let (lines, _) = fit(frags(), 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "HYPHEN-"], vec!["ATION"]]);
    }

    /// A token split across a `<w:r>` boundary arrives as a
    /// [`BreakAfter::Prohibited`] chain. Neither half is the whole word, so
    /// neither is offered to the patterns: a fragment glued to the one
    /// before it, or glued to the one after it, is left alone.
    #[test]
    fn a_word_glued_across_a_run_boundary_is_left_alone() {
        // Glued to what precedes it: `hyphen` + `hyphenation` is one token.
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphen", BreakAfter::Prohibited),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, frags) = fit(frags, 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphen"], vec!["hyphenation"]]);
        assert!(frags
            .iter()
            .all(|f| break_after_of(f) != BreakAfter::Hyphenated));

        // Glued to what follows it: `hyphenation` + `s` is one token. The
        // chain is wider than the line, so the fitter cuts it at the overflow
        // point as it did before #158 — what matters is that no hyphen is
        // inserted into it.
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Prohibited),
            word("s", BreakAfter::Opportunity),
        ];
        let (lines, frags) = fit(frags, 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphenation"], vec!["s"]]);
        assert!(frags
            .iter()
            .all(|f| break_after_of(f) != BreakAfter::Hyphenated));
    }

    /// A run with no `w:lang` this engine has patterns for is never
    /// hyphenated, however good the gap.
    #[test]
    fn a_fragment_without_a_language_is_never_hyphenated() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            text_frag("hyphenation", 66.0),
        ];
        let (lines, _) = fit(frags, 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphenation"]]);
    }

    /// A shaped run's advance is HarfBuzz's over the whole run; cutting it and
    /// re-measuring a head with `measure` would paint a different width than
    /// the one fitted. Such a fragment is left whole.
    #[test]
    fn a_shaped_fragment_is_never_hyphenated() {
        let mut shaped = word("hyphenation", BreakAfter::Opportunity);
        if let Fragment::Text { shaped: s, .. } = &mut shaped {
            *s = Some(crate::render::shape::RunDirection::LeftToRight);
        }
        let frags = vec![word("the ", BreakAfter::Opportunity), shaped];
        let (lines, _) = fit(frags, 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphenation"]]);
    }

    /// §17.11.12: a footnote reference mark is text, but not a word.
    #[test]
    fn a_footnote_reference_is_never_hyphenated() {
        let mut mark = word("hyphenation", BreakAfter::Opportunity);
        if let Fragment::Text {
            is_footnote_ref, ..
        } = &mut mark
        {
            *is_footnote_ref = true;
        }
        let frags = vec![word("the ", BreakAfter::Opportunity), mark];
        let (lines, _) = fit(frags, 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the "], vec!["hyphenation"]]);
    }

    /// What is offered to the patterns is the run of letters inside the
    /// token; leading punctuation stays with the head, trailing punctuation
    /// and whitespace with the tail. A token with a digit, an apostrophe or
    /// a hyphen inside its letters is not a word the patterns are asked
    /// about.
    #[test]
    fn the_core_of_a_token_is_its_run_of_letters() {
        assert_eq!(hyphenation_core("hyphenation"), Some(0..11));
        assert_eq!(hyphenation_core("hyphenation "), Some(0..11));
        assert_eq!(hyphenation_core("(hyphenation),"), Some(1..12));
        assert_eq!(hyphenation_core("„Donau“ "), Some(3..8));
        assert_eq!(hyphenation_core("abc123"), None, "digits");
        assert_eq!(hyphenation_core("don't"), None, "punctuation inside");
        assert_eq!(hyphenation_core("hard-nosed"), None, "a hyphen inside");
        assert_eq!(hyphenation_core("..."), None, "no letters");
        assert_eq!(hyphenation_core(" "), None);
    }

    /// Leading and trailing punctuation travel with the halves they are
    /// attached to: `(hyphenation),` at 72 pt becomes `(hyphen-` + `ation),`
    /// (`(hyphena-` would be 54 pt, one character too wide beside `the `).
    #[test]
    fn punctuation_around_the_word_stays_with_its_half() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("(hyphenation),", BreakAfter::Opportunity),
        ];
        let (lines, _) = fit(frags, 72.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "(hyphen-"], vec!["ation),"]]);
    }

    /// The tail is the rest of the original fragment: its trailing space
    /// still hangs (`width` counts it, `trimmed_width` does not) and it keeps
    /// the fragment's own `break_after`, so what follows it fits as before.
    #[test]
    fn the_tail_keeps_its_trailing_space_and_original_break_status() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation ", BreakAfter::Opportunity),
            word("x", BreakAfter::Opportunity),
        ];
        let (lines, frags) = fit(frags, 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines, [vec!["the ", "hyphen-"], vec!["ation ", "x"]]);
        assert_eq!(text_of(&frags[2]), "ation ");
        assert_eq!(frags[2].width().raw(), 36.0);
        assert_eq!(frags[2].trimmed_width().raw(), 30.0);
        assert_eq!(break_after_of(&frags[2]), BreakAfter::Opportunity);
    }

    /// The guard for every document without `w:autoHyphenation`: with no
    /// context, `fit_lines_in` is the old fitter — same lines, field for
    /// field — and never materialises the borrowed vector.
    #[test]
    fn no_context_leaves_fitting_bit_for_bit_unchanged() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation ", BreakAfter::Opportunity),
            glued_frag("ID‑", 20.0),
            glued_frag("001 ", 35.0),
            Fragment::Tab {
                line_height: Pt::new(14.0),
                font: Rc::new(FontProps {
                    effects: Default::default(),
                    rtl: crate::render::fonts::Toggle::Absent,
                    family: Rc::from("Test"),
                    size: Pt::new(12.0),
                    bold: Toggle::Absent,
                    italic: Toggle::Absent,
                    underline: false,
                    char_spacing: Pt::ZERO,
                    text_scale: 1.0,
                    underline_position: Pt::ZERO,
                    underline_thickness: Pt::ZERO,
                    hyphenation: None,
                }),
                color: RgbColor::BLACK,
                fitting_width: None,
            },
            word("end", BreakAfter::Opportunity),
            Fragment::LineBreak {
                line_height: Pt::new(14.0),
            },
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let old = fit_lines_with_first(&frags, Pt::new(60.0), Pt::new(66.0), geometry(66.0), false);
        let mut cow: Cow<[Fragment]> = Cow::Borrowed(&frags);
        let new = fit_lines_in(
            &mut cow,
            FitWindow::ALL,
            Pt::new(60.0),
            Pt::new(66.0),
            geometry(66.0),
            false,
            None,
        );
        assert!(matches!(cow, Cow::Borrowed(_)), "no context, no copy");
        assert!(old.len() > 2, "the fixture produces several lines");
        assert_eq!(old.len(), new.len());
        for (a, b) in old.iter().zip(&new) {
            assert_eq!((a.start, a.end), (b.start, b.end));
            assert_eq!(a.width.raw(), b.width.raw());
            assert_eq!(a.height.raw(), b.height.raw());
            assert_eq!(a.text_height.raw(), b.text_height.raw());
            assert_eq!(a.ascent.raw(), b.ascent.raw());
            assert_eq!(a.has_break, b.has_break);
        }
    }

    /// The inverse of the splice: after fitting has cut two words, rejoining
    /// gives back the words that went in — text, break status, measured
    /// width and hanging space alike.
    #[test]
    fn rejoin_hyphenated_restores_the_words_the_fitter_cut() {
        let original = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let (lines, mut cut) = fit(original.clone(), 66.0, Some(&ctx(0.0, 0, false)));
        assert_eq!(lines.len(), 3, "both words were cut: {lines:?}");
        assert_eq!(cut.len(), 5);

        rejoin_hyphenated(&mut cut, Some(&six_per_char));

        assert_eq!(cut.len(), original.len());
        for (a, b) in cut.iter().zip(&original) {
            assert_eq!(text_of(a), text_of(b));
            assert_eq!(break_after_of(a), break_after_of(b));
            assert_eq!(a.width().raw(), b.width().raw());
            assert_eq!(a.trimmed_width().raw(), b.trimmed_width().raw());
        }
    }

    /// The floats path fits one line per call from wherever the last one
    /// ended: `start` is an index into the whole vector and the line comes
    /// back in that index space, and `max_lines: Some(1)` stops fitting there
    /// — so no word beyond the first line is hyphenated into a vector the
    /// caller is about to re-fit at another width.
    #[test]
    fn a_window_fits_from_its_start_and_stops_at_its_line_limit() {
        let frags = vec![
            word("the ", BreakAfter::Opportunity),
            word("hyphenation ", BreakAfter::Opportunity),
            word("hyphenation", BreakAfter::Opportunity),
        ];
        let hyph = ctx(0.0, 0, false);
        let mut cow: Cow<[Fragment]> = Cow::Owned(frags.clone());
        let lines = fit_lines_in(
            &mut cow,
            FitWindow {
                start: 0,
                max_lines: Some(1),
            },
            Pt::new(66.0),
            Pt::new(66.0),
            geometry(66.0),
            false,
            Some(&hyph),
        );
        assert_eq!(lines.len(), 1);
        assert_eq!((lines[0].start, lines[0].end), (0, 2));
        assert_eq!(cow.len(), 4, "only the first line's word was cut");
        assert_eq!(text_of(&cow[3]), "hyphenation", "the second is untouched");

        // Continue from where that line ended; indices stay absolute.
        let lines = fit_lines_in(
            &mut cow,
            FitWindow {
                start: 2,
                max_lines: Some(1),
            },
            Pt::new(66.0),
            Pt::new(66.0),
            geometry(66.0),
            false,
            Some(&hyph),
        );
        assert_eq!(lines.len(), 1);
        assert_eq!((lines[0].start, lines[0].end), (2, 4));
        assert_eq!(text_of(&cow[3]), "hy-");
        assert_eq!(text_of(&cow[4]), "phenation");
    }
}
