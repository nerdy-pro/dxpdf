//! Line fitting — break fragments into lines that fit within a max width.

use crate::dimension::Pt;
use crate::layout::fragment::Fragment;

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
/// Breaks at the last whitespace/hyphen boundary when a line overflows.
/// A single fragment wider than `max_width` gets its own line (no infinite loop).
///
/// `first_line_width`: if provided, the first line uses this narrower width
/// (e.g., to account for first-line indent). Subsequent lines use `max_width`.
pub fn fit_lines(fragments: &[Fragment], max_width: Pt) -> Vec<FittedLine> {
    fit_lines_with_first(fragments, max_width, max_width)
}

/// Line fitting with separate first-line and remaining-line widths.
pub fn fit_lines_with_first(fragments: &[Fragment], first_line_width: Pt, remaining_width: Pt) -> Vec<FittedLine> {
    if fragments.is_empty() {
        return Vec::new();
    }

    let mut lines = Vec::new();
    let mut line_start = 0;
    let mut line_width = Pt::ZERO;
    let mut line_height = Pt::ZERO;
    let mut line_text_height = Pt::ZERO;
    let mut line_ascent = Pt::ZERO;
    let mut last_break_point = None; // index after which we can break

    let mut i = 0;
    while i < fragments.len() {
        let frag = &fragments[i];

        // Explicit line break — emit current line including the break fragment.
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
            line_height = Pt::ZERO;
            line_text_height = Pt::ZERO;
            line_ascent = Pt::ZERO;
            last_break_point = None;
            i += 1;
            continue;
        }

        let frag_width = frag.width();
        let new_width = line_width + frag_width;

        // Use first-line width for line 0, remaining width for subsequent lines.
        let current_max = if lines.is_empty() { first_line_width } else { remaining_width };

        // For overflow checking, use trimmed width — trailing whitespace on the
        // last word is allowed to hang past the margin (standard Word behavior).
        // The check uses: previous fragments' full widths + this fragment's trimmed width.
        let check_width = line_width + frag.trimmed_width();

        // Check if adding this fragment overflows.
        if check_width > current_max && line_start < i {
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
            line_height = Pt::ZERO;
            line_text_height = Pt::ZERO;
            line_ascent = Pt::ZERO;
            last_break_point = None;
            // Don't advance i — re-evaluate this fragment on the new line.
            continue;
        }

        // If this is the first fragment on the line and it overflows,
        // allow it (it will be the only fragment on this line). The
        // paragraph renderer will clip/overflow as needed.
        line_width = new_width;
        line_height = line_height.max(frag.height());
        if !matches!(frag, Fragment::Image { .. }) {
            line_text_height = line_text_height.max(frag.height());
        }
        if let Fragment::Text { metrics, .. } = frag {
            line_ascent = line_ascent.max(metrics.ascent);
        }

        // Track break opportunity: only after fragments that end with whitespace,
        // or non-text fragments (tabs, images). Text fragments without trailing
        // whitespace are mid-word continuations (e.g., a word split across runs)
        // and must not be broken.
        let is_break_point = match frag {
            Fragment::Text { text, .. } => text.ends_with(' ') || text.ends_with('\t'),
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
    let mut m = RangeMeasure { width: Pt::ZERO, height: Pt::ZERO, text_height: Pt::ZERO, ascent: Pt::ZERO };
    for frag in &fragments[start..end] {
        m.width += frag.width();
        m.height = m.height.max(frag.height());
        if !matches!(frag, Fragment::Image { .. }) {
            m.text_height = m.text_height.max(frag.height());
        }
        if let Fragment::Text { metrics, .. } = frag {
            m.ascent = m.ascent.max(metrics.ascent);
        }
    }
    m
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::layout::fragment::{FontProps, TextMetrics};
    use std::rc::Rc;
    use crate::resolve::color::RgbColor;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            text: text.into(),
            font: FontProps {
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: false,
                italic: false,
                underline: false,
                char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width), trimmed_width: Pt::new(width),
            metrics: TextMetrics { ascent: Pt::new(10.0), descent: Pt::new(4.0) },
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO, text_offset: Pt::ZERO,
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
                text: "small".into(),
                font: FontProps {
                    family: Rc::from("Test"),
                    size: Pt::new(10.0),
                    bold: false,
                    italic: false,
                    underline: false,
                    char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
                },
                color: RgbColor::BLACK,
                width: Pt::new(20.0), trimmed_width: Pt::new(20.0),
                metrics: TextMetrics { ascent: Pt::new(9.0), descent: Pt::new(3.0) },
                hyperlink_url: None,
                shading: None, border: None, baseline_offset: Pt::ZERO, text_offset: Pt::ZERO,
            },
            Fragment::Text {
                text: "big".into(),
                font: FontProps {
                    family: Rc::from("Test"),
                    size: Pt::new(24.0),
                    bold: false,
                    italic: false,
                    underline: false,
                    char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
                },
                color: RgbColor::BLACK,
                width: Pt::new(30.0), trimmed_width: Pt::new(30.0),
                metrics: TextMetrics { ascent: Pt::new(22.0), descent: Pt::new(6.0) },
                hyperlink_url: None,
                shading: None, border: None, baseline_offset: Pt::ZERO, text_offset: Pt::ZERO,
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
}
