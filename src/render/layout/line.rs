//! Line fitting — break fragments into lines that fit within a max width.

use crate::render::dimension::Pt;
use crate::render::layout::fragment::{BreakAfter, Fragment};

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
    )
}

/// Line fitting with separate first-line and remaining-line widths.
///
/// `ptab_geometry` is the paragraph geometry §17.3.1.30 position tabs resolve
/// against. Fitting needs it because a tab whose alignment point lies behind
/// the pen advances to the next line, and only fitting can create one.
pub fn fit_lines_with_first(
    fragments: &[Fragment],
    first_line_width: Pt,
    remaining_width: Pt,
    ptab_geometry: crate::render::layout::paragraph::PTabGeometry,
) -> Vec<FittedLine> {
    if fragments.is_empty() {
        return Vec::new();
    }

    let mut lines = Vec::new();
    let mut line_start = 0;
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

    let mut i = 0;
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
            let placement = crate::render::layout::paragraph::resolve_ptab(
                *align,
                *relative_to,
                ptab_geometry,
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
        // already drifted from the one used to cut the fragments.
        let is_break_point = match frag {
            Fragment::Text { break_after, .. } => *break_after == BreakAfter::Opportunity,
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

    fn frag(text: &str, width: f32, break_after: BreakAfter) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after,
            font: Rc::new(FontProps {
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
        );
        assert_eq!(lines.len(), 2);
        assert_eq!(lines[0].end, 1, "only 'a ' fits the narrow first line");
        assert_eq!(lines[1].start, 1);
        assert_eq!(
            lines[1].end, 3,
            "'b ' + 'c' both fit on the full second line"
        );
    }
}
