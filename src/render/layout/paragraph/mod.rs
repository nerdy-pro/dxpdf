//! Paragraph layout — line fitting, alignment, spacing, borders, shading.
//!
//! Implements the LayoutBox protocol: receives BoxConstraints, returns PtSize,
//! emits DrawCommands at absolute offsets during paint.

mod borders;
mod line_emit;

// §17.3.1.30: line fitting needs the same position-tab resolution emission
// uses, so the two cannot disagree about where a tab puts content.
pub use line_emit::{resolve_ptab, PTabGeometry, PTabPlacement};
pub(crate) use line_emit::{zone_end, zone_width};
mod types;

pub use types::*;

use std::borrow::Cow;

use super::draw_command::DrawCommand;
use super::fragment::Fragment;
use super::BoxConstraints;
use crate::render::dimension::Pt;
use crate::render::geometry::{PtOffset, PtRect, PtSize};

use crate::render::layout::fragment::split_oversized_fragments;

use borders::{emit_paragraph_borders_and_shading, emit_segment_borders_and_shading, SegmentEdges};
#[cfg(test)]
use line_emit::resolve_line_height;
use line_emit::{compute_line_placements, emit_line_commands};

// ── Tab leader rendering constants ────────────────────────────────────────────

/// Fallback width for a single leader character when no measurer is available (pt).
const LEADER_CHAR_WIDTH_FALLBACK: Pt = Pt::new(4.0);

/// A fitted line together with the per-line float adjustments that were active when it was placed.
struct LinePlacement {
    line: super::line::FittedLine,
    /// Width stolen from the left by an active float.
    float_left: Pt,
    /// Width stolen from the right by an active float.
    float_right: Pt,
}

/// Shared layout parameters threaded through `compute_line_placements` and
/// `emit_line_commands`.
struct LineLayoutParams {
    content_width: Pt,
    /// The full constraint width (page/cell text area between margins), before
    /// paragraph indents are subtracted. §17.3.1.30 position tabs measured
    /// `relativeTo="margin"` reference this, not the indented `content_width`.
    max_width: Pt,
    first_line_adjustment: Pt,
    drop_cap_indent: Pt,
    drop_cap_lines: usize,
    default_line_height: Pt,
}

/// Clip inline images wider than `max_width` to the content box, cropping the
/// horizontal overflow rather than scaling.
///
/// Per ECMA-376 §20.4.2.7, a `wp:inline` object's `wp:extent` is its *final*
/// height and width — an inline object is **not** auto-scaled to fit its
/// container. When it is wider than the containing table cell (§17.4.53: an
/// autofit table at 100% width cannot grow its column), Word draws the image at
/// full size and clips the overflow to the cell boundary.
///
/// The retained slice is the **left** part of the object. An inline object that
/// is already wider than the line is placed at the line's start: justification
/// (§17.3.1.13 `w:jc`) distributes only *spare* width, and an over-wide object
/// leaves none — Word does not centre or right-shift it — so it overflows to
/// the right and the cell clips that side. (Our own line fitter agrees: it
/// clamps the alignment offset to `max(0)`.) We reproduce that by keeping the
/// image's height and scale, setting its displayed width to `max_width`, and
/// cropping the source to the left visible fraction, composed with any existing
/// `a:srcRect` crop — no symmetric/centre crop, which would shift the pixels.
///
/// Returns `None` when nothing is oversized, so the common case borrows the
/// input slice and allocates nothing.
fn clip_oversized_images(fragments: &[Fragment], max_width: Pt) -> Option<Vec<Fragment>> {
    let oversized =
        |f: &Fragment| matches!(f, Fragment::Image { size, .. } if size.width > max_width);
    if max_width <= Pt::ZERO || !fragments.iter().any(oversized) {
        return None;
    }
    Some(
        fragments
            .iter()
            .map(|f| match f {
                Fragment::Image {
                    size,
                    rel_id,
                    image_data,
                    src_rect,
                } if size.width > max_width => {
                    // Fraction of the object's width that survives the clip.
                    let visible = max_width.raw() / size.width.raw();
                    // Start from the existing crop (fraction of the source), or
                    // the whole image, then keep its left `visible` sub-fraction
                    // — same origin, narrower width (crop the right overflow).
                    let base = src_rect.unwrap_or_else(|| {
                        PtRect::from_xywh(Pt::ZERO, Pt::ZERO, Pt::new(1.0), Pt::new(1.0))
                    });
                    let cropped = PtRect::from_xywh(
                        base.origin.x,
                        base.origin.y,
                        Pt::new(base.size.width.raw() * visible),
                        base.size.height,
                    );
                    Fragment::Image {
                        // Height and scale are unchanged (§20.4.2.7); only the
                        // displayed width shrinks, showing the left, un-clipped
                        // part.
                        size: PtSize::new(max_width, size.height),
                        rel_id: rel_id.clone(),
                        image_data: image_data.clone(),
                        src_rect: Some(cropped),
                    }
                }
                other => other.clone(),
            })
            .collect(),
    )
}

/// A paragraph whose lines have been fitted and measured, ready to emit.
///
/// The layout of a paragraph splits into two halves: *place* (fit fragments into
/// lines and measure them — [`place_paragraph`]) and *emit* (produce
/// `DrawCommand`s). The common case emits the whole paragraph at once
/// ([`PlacedParagraph::emit_full`], wrapped by [`layout_paragraph`]). When a
/// paragraph must break across a page boundary (§17.3.1.14 `keepLines`,
/// §17.3.1.44 `widowControl`), the stacker emits sub-ranges of its lines at
/// different page origins via [`PlacedParagraph::emit_split_segment`].
///
/// `line_placements` index into `fragments`, which is the *effective* fragment
/// list after inline-image clipping (§20.4.2.7) and oversized-fragment
/// splitting — hence the `Cow`, which borrows on the common no-transform path.
pub(crate) struct PlacedParagraph<'a> {
    fragments: Cow<'a, [Fragment]>,
    line_placements: Vec<LinePlacement>,
    params: LineLayoutParams,
    style: &'a ParagraphStyle,
    constraints: &'a BoxConstraints,
    default_line_height: Pt,
    measure_text: MeasureTextFn<'a>,
    /// §17.3.1.11: precomputed baseline for the drop cap, if any.
    drop_cap_baseline_y: Option<Pt>,
    /// §17.3.1.33: resolved height of each fitted line, in line order. Lets the
    /// stacker choose split points without re-measuring.
    line_heights: Vec<Pt>,
}

/// Fit and measure a paragraph's lines without emitting draw commands — the
/// "place" half of paragraph layout (see [`PlacedParagraph`]).
pub(crate) fn place_paragraph<'a>(
    fragments: &'a [Fragment],
    constraints: &'a BoxConstraints,
    style: &'a ParagraphStyle,
    default_line_height: Pt,
    measure_text: MeasureTextFn<'a>,
) -> PlacedParagraph<'a> {
    // §17.3.1.11: drop cap text frame.
    // Drop mode: body text indented by drop cap position + width + hSpace.
    // Margin mode: drop cap is in the margin, body text is NOT indented.
    let drop_cap_indent = style
        .drop_cap
        .as_ref()
        .filter(|dc| !dc.margin_mode)
        .map(|dc| dc.indent + dc.width + dc.h_space)
        .unwrap_or(Pt::ZERO);
    let drop_cap_lines = style
        .drop_cap
        .as_ref()
        .map(|dc| dc.lines as usize)
        .unwrap_or(0);

    // §17.3.1.24: border space is the distance between the border line and the text.
    // Only the space reduces the text area, not the border line width.
    let border_space_left = style
        .borders
        .as_ref()
        .and_then(|b| b.left.as_ref())
        .map(|b| b.space)
        .unwrap_or(Pt::ZERO);
    let border_space_right = style
        .borders
        .as_ref()
        .and_then(|b| b.right.as_ref())
        .map(|b| b.space)
        .unwrap_or(Pt::ZERO);
    let content_width = constraints.max_width
        - style.indent_left
        - style.indent_right
        - border_space_left
        - border_space_right;
    // §17.3.1.12: first-line indent adjusts the first line's available width.
    // Positive = narrower (indent), negative = wider (hanging indent).
    // Drop cap indent also reduces width for the first N lines.
    let first_line_adjustment = style.indent_first_line + drop_cap_indent;

    // Effective fragments: clip inline images wider than the content box
    // (§20.4.2.7 — Word clips rather than scales), then split oversized text
    // fragments into per-character fragments for narrow-cell character-level
    // breaking. `Cow` so the common path (nothing over-wide) borrows the input
    // without cloning; only an actual transform materializes an owned `Vec`.
    let min_avail = (content_width - first_line_adjustment).max(Pt::ZERO);
    let effective: Cow<'a, [Fragment]> = {
        let clipped: Cow<'a, [Fragment]> = match clip_oversized_images(fragments, content_width) {
            Some(v) => Cow::Owned(v),
            None => Cow::Borrowed(fragments),
        };
        // `split_oversized_fragments` returns `None` when nothing needs
        // splitting, so the borrow survives the common case untouched — no
        // mirrored predicate, and no clone.
        match split_oversized_fragments(&clipped, min_avail, measure_text) {
            Some(split) => Cow::Owned(split),
            None => clipped,
        }
    };

    let params = LineLayoutParams {
        content_width,
        max_width: constraints.max_width,
        first_line_adjustment,
        drop_cap_indent,
        drop_cap_lines,
        default_line_height,
    };
    // Per-line float adjustment: fit one line at a time, computing the available
    // width for each line based on its absolute y position on the page.
    let line_placements = compute_line_placements(&effective, style, &params);

    // §17.3.1.11: compute the drop cap baseline.
    // When frame_height is set (lineRule="exact"):
    //   baseline = frame_top + frame_height - descent + position_offset
    // Otherwise fall back to aligning with the Nth body line's baseline.
    let drop_cap_baseline_y = if let Some(ref dc) = style.drop_cap {
        if let Some(fh) = dc.frame_height {
            Some(style.space_before + fh + dc.position_offset)
        } else {
            let n = dc.lines.max(1) as usize;
            let mut y = style.space_before;
            for (i, lp) in line_placements.iter().enumerate().take(n) {
                let natural = if lp.line.height > Pt::ZERO {
                    lp.line.height
                } else {
                    default_line_height
                };
                let text_h = if lp.line.text_height > Pt::ZERO {
                    lp.line.text_height
                } else {
                    default_line_height
                };
                let m = line_emit::resolve_line_metrics(natural, text_h, style);
                if i == n - 1 {
                    y += m.baseline_shift + lp.line.ascent;
                    break;
                }
                y += m.advance;
            }
            Some(y)
        }
    } else {
        None
    };

    // §17.3.1.33: resolve each line's height once, matching `emit_line_commands`.
    let line_heights = line_placements
        .iter()
        .map(|lp| {
            let natural = if lp.line.height > Pt::ZERO {
                lp.line.height
            } else {
                default_line_height
            };
            let text_h = if lp.line.text_height > Pt::ZERO {
                lp.line.text_height
            } else {
                default_line_height
            };
            line_emit::resolve_line_metrics(natural, text_h, style).advance
        })
        .collect();

    PlacedParagraph {
        fragments: effective,
        line_placements,
        params,
        style,
        constraints,
        default_line_height,
        measure_text,
        drop_cap_baseline_y,
        line_heights,
    }
}

impl PlacedParagraph<'_> {
    /// Number of fitted lines.
    pub(crate) fn line_count(&self) -> usize {
        self.line_placements.len()
    }

    /// §17.3.1.33: resolved height of line `i`.
    pub(crate) fn line_height(&self, i: usize) -> Pt {
        self.line_heights[i]
    }

    /// Number of leading lines that must stay together on the first segment
    /// when this paragraph splits — the larger of:
    ///
    /// - §17.3.1.11 the drop-cap span (a split inside it tears the glyph), and
    /// - §17.3.1.44 the float-wrapped run: every line up to and including the
    ///   last one narrowed by an active float. Keeping these on the first
    ///   segment guarantees the continuation is full-width, so the already-fitted
    ///   tail lines are reused correctly on the next page without re-wrapping.
    pub(crate) fn unbreakable_prefix_lines(&self) -> usize {
        let float_prefix = self
            .line_placements
            .iter()
            .rposition(|lp| lp.float_left > Pt::ZERO || lp.float_right > Pt::ZERO)
            .map_or(0, |i| i + 1);
        self.params.drop_cap_lines.max(float_prefix)
    }

    /// §4 continuation re-fit: the effective fragments from line `from` onward,
    /// cloned so the stacker can re-place the remainder against the next page's
    /// floats and width (instead of reusing this page's placement). `from` must
    /// be a valid line index (`< line_count()`).
    pub(crate) fn fragments_from_line(&self, from: usize) -> Vec<Fragment> {
        let start = self.line_placements[from].line.start;
        self.fragments[start..].to_vec()
    }

    /// §17.11.12: number of footnote reference marks in lines `[first, last)`.
    /// The stacker reserves this many footnotes (in document order) on the page
    /// that segment lands on, so a footnote sits with its reference.
    pub(crate) fn footnote_refs_in(&self, first: usize, last: usize) -> usize {
        self.line_placements[first..last]
            .iter()
            .map(|lp| {
                self.fragments[lp.line.start..lp.line.end]
                    .iter()
                    .filter(|f| {
                        matches!(
                            f,
                            Fragment::Text {
                                is_footnote_ref: true,
                                ..
                            }
                        )
                    })
                    .count()
            })
            .sum()
    }

    /// Render the drop cap glyphs at their precomputed baseline (§17.3.1.11).
    /// Only the first segment of a split paragraph carries the drop cap.
    fn emit_drop_cap(&self, commands: &mut Vec<DrawCommand>) {
        if let (Some(ref dc), Some(baseline_y)) = (&self.style.drop_cap, self.drop_cap_baseline_y) {
            // §17.3.1.11: position the drop cap using its own paragraph's indent.
            // Drop mode: at the drop cap paragraph's indent (inside text area).
            // Margin mode: in the page margin, to the left of text.
            let dc_x = if dc.margin_mode {
                dc.indent - dc.width - dc.h_space
            } else {
                dc.indent
            };
            for frag in &dc.fragments {
                if let Fragment::Text {
                    text, font, color, ..
                } = frag
                {
                    commands.push(DrawCommand::Text {
                        shaped: None,
                        position: PtOffset::new(dc_x, baseline_y),
                        text: text.clone(),
                        font_family: font.family.clone(),
                        char_spacing: font.char_spacing,
                        font_size: font.size,
                        bold: font.bold,
                        italic: font.italic,
                        color: *color,
                        text_scale: font.text_scale,
                    });
                }
            }
        }
    }

    /// Emit the whole paragraph — drop cap, every line, and border/shading with
    /// `space_before`/`space_after`. The common, non-split path.
    ///
    /// Total height is `space_before + Σ line heights + space_after`; the
    /// caller adds it to the page cursor. `space_after` can collapse with the
    /// next paragraph's `space_before` (§17.3.1.9).
    pub(crate) fn emit_full(&self) -> ParagraphLayout {
        let mut commands = Vec::new();
        let mut cursor_y = self.style.space_before;

        self.emit_drop_cap(&mut commands);

        emit_line_commands(
            &mut commands,
            &mut cursor_y,
            &self.line_placements,
            0..self.line_placements.len(),
            &self.fragments,
            self.style,
            &self.params,
            self.measure_text,
            true,
        );

        // §17.3.1.24: paragraph border and shading coordinate system.
        // Borders sit at the paragraph indent edges. The border `space` is the
        // distance between the border line and the text content. Top/bottom
        // border space expands the bordered area vertically.
        cursor_y = emit_paragraph_borders_and_shading(
            &mut commands,
            self.style,
            self.constraints,
            cursor_y,
            self.default_line_height,
            self.line_placements.is_empty(),
        );

        ParagraphLayout {
            commands,
            size: PtSize::new(self.constraints.max_width, cursor_y),
        }
    }

    /// §17.3.1.24: bottom border space this paragraph adds after its last line —
    /// charged, with `space_after`, only on the final segment.
    pub(crate) fn bottom_border_space(&self) -> Pt {
        self.style
            .borders
            .as_ref()
            .and_then(|b| b.bottom.as_ref())
            .map(|b| b.space)
            .unwrap_or(Pt::ZERO)
    }

    /// Emit lines `[first, last)` of a splittable paragraph as one page segment.
    ///
    /// The caller keeps the unbreakable prefix (see
    /// [`unbreakable_prefix_lines`](Self::unbreakable_prefix_lines)) on the first
    /// segment, so the drop cap and any float-wrapped lines stay intact.
    /// `space_before` (§17.3.1.33) and the top border belong to the first
    /// segment; `space_after` and the bottom border to the last; shading and side
    /// borders span every segment (§17.3.1.24/§17.3.1.31). Commands are
    /// positioned relative to the segment's own origin `(0, 0)`; the stacker
    /// shifts them to the page.
    pub(crate) fn emit_split_segment(
        &self,
        first: usize,
        last: usize,
        is_first_segment: bool,
        is_last_segment: bool,
    ) -> ParagraphLayout {
        debug_assert!(
            first < last && last <= self.line_placements.len(),
            "split range {first}..{last} out of bounds for {} lines",
            self.line_placements.len()
        );
        let mut commands = Vec::new();
        // §17.3.1.33: only the first segment carries the paragraph's space_before.
        let mut cursor_y = if is_first_segment {
            self.style.space_before
        } else {
            Pt::ZERO
        };

        // §17.3.1.11: the drop cap belongs to the first segment (the stacker's
        // guard keeps the whole drop-cap prefix there, so the glyph is intact).
        if is_first_segment {
            self.emit_drop_cap(&mut commands);
        }

        emit_line_commands(
            &mut commands,
            &mut cursor_y,
            &self.line_placements,
            first..last,
            &self.fragments,
            self.style,
            &self.params,
            self.measure_text,
            is_first_segment,
        );

        // §17.3.1.24/§17.3.1.31: per-segment border/shading box; the top edge and
        // space_before are on the first segment, the bottom edge and space_after
        // on the last, sides and shading on every segment.
        cursor_y = emit_segment_borders_and_shading(
            &mut commands,
            self.style,
            self.constraints,
            cursor_y,
            self.default_line_height,
            false,
            SegmentEdges::for_segment(is_first_segment, is_last_segment),
        );

        ParagraphLayout {
            commands,
            size: PtSize::new(self.constraints.max_width, cursor_y),
        }
    }
}

/// Lay out a paragraph: fit fragments into lines, apply alignment and spacing.
///
/// Returns draw commands positioned relative to (0, 0). The caller positions
/// the paragraph by adding its offset during the paint phase.
pub fn layout_paragraph(
    fragments: &[Fragment],
    constraints: &BoxConstraints,
    style: &ParagraphStyle,
    default_line_height: Pt,
    measure_text: MeasureTextFn<'_>,
) -> ParagraphLayout {
    place_paragraph(
        fragments,
        constraints,
        style,
        default_line_height,
        measure_text,
    )
    .emit_full()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Alignment, PTabAlignment, PTabRelativeTo};
    use crate::render::fonts::Toggle;
    use crate::render::layout::fragment::{FontProps, LinkTarget, TextMetrics};
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after: crate::render::layout::fragment::fixture_break_after(text),
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

    fn hyperlink_frag(text: &str, width: f32, url: &str) -> Fragment {
        let mut fragment = text_frag(text, width);
        if let Fragment::Text { hyperlink_url, .. } = &mut fragment {
            *hyperlink_url = Some(LinkTarget::External(std::rc::Rc::from(url)));
        }
        fragment
    }

    fn underlined_text_frag(text: &str, width: f32) -> Fragment {
        let mut fragment = text_frag(text, width);
        if let Fragment::Text { font, .. } = &mut fragment {
            Rc::make_mut(font).underline = true;
            Rc::make_mut(font).underline_thickness = Pt::new(0.5);
        }
        fragment
    }

    fn body_constraints(width: f32) -> BoxConstraints {
        BoxConstraints::new(Pt::ZERO, Pt::new(width), Pt::ZERO, Pt::new(1000.0))
    }

    fn image_frag(w: f32, h: f32) -> Fragment {
        Fragment::Image {
            size: PtSize::new(Pt::new(w), Pt::new(h)),
            rel_id: "rId1".to_string(),
            image_data: None,
            src_rect: None,
        }
    }

    fn clipped_src(frag: &Fragment) -> (PtSize, PtRect) {
        match frag {
            Fragment::Image { size, src_rect, .. } => (*size, src_rect.expect("crop applied")),
            other => panic!("expected image, got {other:?}"),
        }
    }

    #[test]
    fn clip_oversized_image_crops_left_keeps_height_and_scale() {
        // 200×100 image in a 100pt content box → keep the LEFT 50% of the
        // width, full height. Height is UNCHANGED (not scaled) — a crop, not a
        // fit — and the crop is left-anchored (origin.x stays 0), so the right
        // overflow is what gets clipped, matching Word.
        let frags = vec![image_frag(200.0, 100.0)];
        let out = clip_oversized_images(&frags, Pt::new(100.0)).expect("clip");
        let (size, crop) = clipped_src(&out[0]);
        assert_eq!(size.width.raw(), 100.0);
        assert_eq!(size.height.raw(), 100.0, "height must NOT change");
        assert!(
            (crop.origin.x.raw() - 0.0).abs() < 1e-4,
            "left-anchored, not centred"
        );
        assert!((crop.size.width.raw() - 0.5).abs() < 1e-4);
        assert!((crop.size.height.raw() - 1.0).abs() < 1e-4);
    }

    #[test]
    fn clip_oversized_image_composes_with_existing_src_rect() {
        // Image already cropped to [0.1, 0.9] (width 0.8); clipping to 50% keeps
        // the left 0.4 of the source, at the existing origin.
        let mut frag = image_frag(200.0, 100.0);
        if let Fragment::Image { src_rect, .. } = &mut frag {
            *src_rect = Some(PtRect::from_xywh(
                Pt::new(0.1),
                Pt::ZERO,
                Pt::new(0.8),
                Pt::new(1.0),
            ));
        }
        let out = clip_oversized_images(&[frag], Pt::new(100.0)).expect("clip");
        let crop = clipped_src(&out[0]).1;
        assert!((crop.origin.x.raw() - 0.1).abs() < 1e-4);
        assert!((crop.size.width.raw() - 0.4).abs() < 1e-4);
    }

    #[test]
    fn clip_oversized_images_leaves_a_fitting_image_untouched() {
        // Narrower than the content box → not oversized → borrow, no allocation.
        let frags = vec![image_frag(163.0, 290.0)];
        assert!(clip_oversized_images(&frags, Pt::new(499.0)).is_none());
    }

    #[test]
    fn clip_oversized_images_ignores_wide_non_image_fragments() {
        // Oversized text is handled by split_oversized_fragments, not here.
        let frags = vec![text_frag("very wide text", 999.0)];
        assert!(clip_oversized_images(&frags, Pt::new(100.0)).is_none());
    }

    #[test]
    fn clip_oversized_images_guards_nonpositive_width() {
        let frags = vec![image_frag(516.0, 290.0)];
        assert!(clip_oversized_images(&frags, Pt::ZERO).is_none());
    }

    #[test]
    fn empty_paragraph_has_default_height() {
        let result = layout_paragraph(
            &[],
            &body_constraints(400.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );
        assert_eq!(result.size.height.raw(), 14.0, "default line height");
        assert!(result.commands.is_empty());
    }

    #[test]
    fn single_line_produces_text_command() {
        let frags = vec![text_frag("hello", 30.0)];
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );

        assert_eq!(result.commands.len(), 1);
        if let DrawCommand::Text { text, position, .. } = &result.commands[0] {
            assert_eq!(&**text, "hello");
            assert_eq!(position.x.raw(), 0.0); // left aligned, no indent
        }
    }

    #[test]
    fn center_alignment_shifts_text() {
        let frags = vec![text_frag("hi", 20.0)];
        let style = ParagraphStyle {
            alignment: Alignment::Center,
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &style,
            Pt::new(14.0),
            None,
        );

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 40.0); // (100 - 20) / 2
        }
    }

    #[test]
    fn end_alignment_right_aligns() {
        let frags = vec![text_frag("hi", 20.0)];
        let style = ParagraphStyle {
            alignment: Alignment::End,
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &style,
            Pt::new(14.0),
            None,
        );

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 80.0); // 100 - 20
        }
    }

    #[test]
    fn literal_tab_text_preserves_center_and_end_alignment() {
        let fragments = vec![text_frag("alpha\tbeta", 20.0)];
        for (alignment, expected_x) in [(Alignment::Center, 20.0), (Alignment::End, 40.0)] {
            let style = ParagraphStyle {
                alignment,
                ..Default::default()
            };
            let result = layout_paragraph(
                &fragments,
                &body_constraints(60.0),
                &style,
                Pt::new(14.0),
                None,
            );
            assert!((text_xs(&result)[0] - expected_x).abs() < 0.01);
        }
    }

    #[test]
    fn structured_tab_zones_include_literal_tab_text() {
        for (alignment, expected_xs) in [
            (crate::model::TabAlignment::Right, [45.0, 55.0, 60.0]),
            (crate::model::TabAlignment::Center, [62.5, 72.5, 77.5]),
        ] {
            let style = ParagraphStyle {
                tabs: vec![TabStopDef {
                    position: Pt::new(80.0),
                    alignment,
                    leader: crate::model::TabLeader::None,
                }],
                ..Default::default()
            };
            let result = layout_paragraph(
                &[
                    Fragment::Tab {
                        line_height: Pt::new(14.0),
                        font: leader_font("Test", 12.0),
                        color: RgbColor::BLACK,
                        fitting_width: Some(Pt::new(5.0)),
                    },
                    text_frag("alpha", 10.0),
                    text_frag("\t", 5.0),
                    text_frag("beta", 20.0),
                ],
                &body_constraints(100.0),
                &style,
                Pt::new(14.0),
                None,
            );
            let xs = text_xs(&result);
            for (actual, expected) in xs.iter().zip(expected_xs) {
                assert!((actual - expected).abs() < 0.01, "xs: {xs:?}");
            }
        }
    }

    #[test]
    fn both_alignment_expands_inter_word_gaps_on_soft_wrapped_lines() {
        let mut fragments = vec![
            text_frag("alpha ", 25.0),
            hyperlink_frag("beta ", 25.0, "https://example.invalid"),
            text_frag("gamma", 30.0),
        ];
        if let Fragment::Text { font, .. } = &mut fragments[0] {
            Rc::make_mut(font).underline = true;
            Rc::make_mut(font).underline_thickness = Pt::new(0.5);
        }
        let style = ParagraphStyle {
            alignment: Alignment::Both,
            ..Default::default()
        };
        let result = layout_paragraph(
            &fragments,
            &body_constraints(60.0),
            &style,
            Pt::new(14.0),
            None,
        );
        let xs = text_xs(&result);
        assert!(
            (xs[1] - 35.0).abs() < 0.01,
            "beta must receive the 10pt gap: {xs:?}"
        );
        assert!(
            (xs[2] - 0.0).abs() < 0.01,
            "final line must remain unexpanded: {xs:?}"
        );

        let underline = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Underline { line, .. } => Some(*line),
                _ => None,
            })
            .expect("underlined first gap");
        assert!((underline.end.x.raw() - 35.0).abs() < 0.01);

        let link_rect = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::LinkAnnotation { rect, url }
                    if &**url == "https://example.invalid" =>
                {
                    Some(*rect)
                }
                _ => None,
            })
            .expect("beta hyperlink annotation");
        assert!((link_rect.origin.x.raw() - 35.0).abs() < 0.01);
        assert!((link_rect.size.width.raw() - 25.0).abs() < 0.01);
    }

    #[test]
    fn non_http_external_link_emits_uri_annotation() {
        // Regression: an external URL whose scheme isn't in the old
        // http/https/mailto/ftp allowlist (here `file://`) must still be a URI
        // LinkAnnotation, not misrouted to an internal GoTo. Routing is by the
        // fragment's LinkTarget kind now, not a URL-scheme guess.
        let mut frag = text_frag("doc", 20.0);
        if let Fragment::Text { hyperlink_url, .. } = &mut frag {
            *hyperlink_url = Some(LinkTarget::External("file://server/report.docx".into()));
        }
        let result = layout_paragraph(
            &[frag],
            &body_constraints(200.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );
        assert!(
            result.commands.iter().any(|c| matches!(
                c,
                DrawCommand::LinkAnnotation { url, .. } if &**url == "file://server/report.docx"
            )),
            "file:// external link must be a URI annotation"
        );
        assert!(
            !result
                .commands
                .iter()
                .any(|c| matches!(c, DrawCommand::InternalLink { .. })),
            "external link must not be misrouted to an internal GoTo"
        );
    }

    #[test]
    fn internal_link_emits_goto_destination() {
        let mut frag = text_frag("entry", 40.0);
        if let Fragment::Text { hyperlink_url, .. } = &mut frag {
            *hyperlink_url = Some(LinkTarget::Internal("_Toc123".into()));
        }
        let result = layout_paragraph(
            &[frag],
            &body_constraints(200.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );
        assert!(
            result.commands.iter().any(|c| matches!(
                c,
                DrawCommand::InternalLink { destination, .. } if &**destination == "_Toc123"
            )),
            "internal bookmark link must be a GoTo destination"
        );
        assert!(
            !result
                .commands
                .iter()
                .any(|c| matches!(c, DrawCommand::LinkAnnotation { .. })),
            "internal link must not be emitted as an external URI"
        );
    }

    #[test]
    fn both_alignment_justifies_from_trimmed_line_width() {
        let mut trailing = text_frag("beta ", 25.0);
        if let Fragment::Text { trimmed_width, .. } = &mut trailing {
            *trimmed_width = Pt::new(20.0);
        }
        let result = layout_paragraph(
            &[
                text_frag("alpha ", 25.0),
                trailing,
                text_frag("gamma", 30.0),
            ],
            &body_constraints(60.0),
            &ParagraphStyle {
                alignment: Alignment::Both,
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );

        let xs = text_xs(&result);
        assert!(
            (xs[1] - 40.0).abs() < 0.01,
            "the visible first line is 45pt wide, so its gap must grow by 15pt: {xs:?}",
        );
    }

    #[test]
    fn distribute_alignment_spreads_characters_across_single_line() {
        let result = layout_paragraph(
            &[
                text_frag("AB", 20.0),
                hyperlink_frag("CD", 20.0, "https://example.invalid/distributed"),
            ],
            &body_constraints(60.0),
            &ParagraphStyle {
                alignment: Alignment::Distribute,
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );

        let texts = result
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text {
                    position,
                    char_spacing,
                    ..
                } => Some((position.x.raw(), char_spacing.raw())),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!((texts[0].0 - 0.0).abs() < 0.01);
        assert!((texts[1].0 - 33.333).abs() < 0.01, "text runs: {texts:?}");
        assert!((texts[0].1 - 6.667).abs() < 0.01, "text runs: {texts:?}");
        assert!((texts[1].1 - 6.667).abs() < 0.01, "text runs: {texts:?}");

        let link = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::LinkAnnotation { rect, url }
                    if &**url == "https://example.invalid/distributed" =>
                {
                    Some(*rect)
                }
                _ => None,
            })
            .expect("distributed hyperlink");
        assert!(((link.origin.x + link.size.width).raw() - 60.0).abs() < 0.01);
    }

    /// §17.3.1.13: spare width is shared between *grapheme clusters*, never
    /// inside one. Two accented letters are four scalars but two units, so
    /// there is exactly one gap to fill — the scalar count would invent three
    /// and drop two of them between a letter and its own accent.
    #[test]
    fn distribute_alignment_counts_grapheme_clusters_not_scalars() {
        let accented = "e\u{301}e\u{301}";
        assert_eq!(accented.chars().count(), 4, "fixture must be four scalars");

        let result = layout_paragraph(
            &[text_frag(accented, 20.0)],
            &body_constraints(60.0),
            &ParagraphStyle {
                alignment: Alignment::Distribute,
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );

        let spacings = result
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { char_spacing, .. } => Some(char_spacing.raw()),
                _ => None,
            })
            .collect::<Vec<_>>();

        // 40pt spare over 1 inter-cluster gap. Over 3 inter-scalar gaps it
        // would have been 13.33.
        assert_eq!(spacings.len(), 1);
        assert!(
            (spacings[0] - 40.0).abs() < 0.01,
            "one gap between two clusters, not three between four scalars: {spacings:?}",
        );
    }

    #[test]
    fn distribute_alignment_adds_to_existing_character_spacing() {
        let mut first = text_frag("A", 10.0);
        let mut second = text_frag("B", 10.0);
        for fragment in [&mut first, &mut second] {
            if let Fragment::Text { font, .. } = fragment {
                Rc::make_mut(font).char_spacing = Pt::new(2.0);
            }
        }
        let result = layout_paragraph(
            &[first, second],
            &body_constraints(30.0),
            &ParagraphStyle {
                alignment: Alignment::Distribute,
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );
        let texts = result
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text {
                    position,
                    char_spacing,
                    ..
                } => Some((position.x.raw(), char_spacing.raw())),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert!((texts[0].1 - 12.0).abs() < 0.01, "text runs: {texts:?}");
        assert!((texts[1].0 - 20.0).abs() < 0.01, "text runs: {texts:?}");
    }

    #[test]
    fn distribute_alignment_keeps_tab_lines_at_natural_width() {
        let result = layout_paragraph(
            &[
                text_frag("A", 10.0),
                Fragment::Tab {
                    line_height: Pt::new(14.0),
                    font: leader_font("Test", 12.0),
                    color: RgbColor::BLACK,
                    fitting_width: None,
                },
                text_frag("B", 10.0),
            ],
            &body_constraints(60.0),
            &ParagraphStyle {
                alignment: Alignment::Distribute,
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );
        let spacings = result
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { char_spacing, .. } => Some(*char_spacing),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert!(spacings.iter().all(|spacing| *spacing == Pt::ZERO));
    }

    #[test]
    fn both_alignment_keeps_explicit_break_lines_at_natural_width() {
        let style = ParagraphStyle {
            alignment: Alignment::Both,
            ..Default::default()
        };
        let fragments = vec![
            underlined_text_frag("alpha ", 25.0),
            text_frag("beta", 25.0),
            Fragment::LineBreak {
                line_height: Pt::new(14.0),
            },
            text_frag("gamma", 30.0),
        ];
        let result = layout_paragraph(
            &fragments,
            &body_constraints(60.0),
            &style,
            Pt::new(14.0),
            None,
        );
        assert!((text_xs(&result)[1] - 25.0).abs() < 0.01);
        let underline = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Underline { line, .. } => Some(*line),
                _ => None,
            })
            .expect("underlined explicit-break gap");
        assert!((underline.end.x.raw() - 25.0).abs() < 0.01);
    }

    #[test]
    fn both_alignment_does_not_expand_nbsp_gaps() {
        let style = ParagraphStyle {
            alignment: Alignment::Both,
            ..Default::default()
        };
        let fragments = vec![
            underlined_text_frag("alpha\u{00a0}", 25.0),
            text_frag("beta", 25.0),
            text_frag("gamma", 30.0),
        ];
        let result = layout_paragraph(
            &fragments,
            &body_constraints(60.0),
            &style,
            Pt::new(14.0),
            None,
        );
        assert!((text_xs(&result)[1] - 25.0).abs() < 0.01);
        let underline = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Underline { line, .. } => Some(*line),
                _ => None,
            })
            .expect("underlined NBSP gap");
        assert!((underline.end.x.raw() - 25.0).abs() < 0.01);
    }

    #[test]
    fn both_alignment_keeps_structured_tab_lines_at_natural_width() {
        let style = ParagraphStyle {
            alignment: Alignment::Both,
            ..Default::default()
        };
        let fragments = vec![
            underlined_text_frag("alpha ", 25.0),
            Fragment::Tab {
                line_height: Pt::new(14.0),
                font: leader_font("Test", 12.0),
                color: RgbColor::BLACK,
                fitting_width: None,
            },
            text_frag("beta ", 25.0),
            text_frag("gamma", 30.0),
        ];
        let result = layout_paragraph(
            &fragments,
            &body_constraints(70.0),
            &style,
            Pt::new(14.0),
            None,
        );
        assert!((text_xs(&result)[1] - 36.0).abs() < 0.01);
        let underline = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Underline { line, .. } => Some(*line),
                _ => None,
            })
            .expect("underlined structured-tab gap");
        assert!((underline.end.x.raw() - 25.0).abs() < 0.01);
    }

    #[test]
    fn both_alignment_justifies_after_list_label_separator() {
        let result = layout_paragraph(
            &[
                text_frag("1.", 10.0),
                Fragment::Tab {
                    line_height: Pt::new(14.0),
                    font: leader_font("Test", 12.0),
                    color: RgbColor::BLACK,
                    fitting_width: Some(Pt::new(5.0)),
                },
                text_frag("alpha ", 20.0),
                text_frag("beta ", 20.0),
                text_frag("gamma", 30.0),
            ],
            &body_constraints(60.0),
            &ParagraphStyle {
                alignment: Alignment::Both,
                tabs: vec![TabStopDef {
                    position: Pt::new(15.0),
                    alignment: crate::model::TabAlignment::Left,
                    leader: crate::model::TabLeader::None,
                }],
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );

        let xs = text_xs(&result);
        assert!(
            (xs[2] - 40.0).abs() < 0.01,
            "the first list-item line must justify after its synthetic label tab: {xs:?}",
        );
    }

    #[test]
    fn both_alignment_justifies_after_picture_bullet_separator() {
        let result = layout_paragraph(
            &[
                image_frag(10.0, 10.0),
                Fragment::Tab {
                    line_height: Pt::new(14.0),
                    font: leader_font("Test", 12.0),
                    color: RgbColor::BLACK,
                    fitting_width: Some(Pt::new(5.0)),
                },
                text_frag("alpha ", 20.0),
                text_frag("beta ", 20.0),
                text_frag("gamma", 30.0),
            ],
            &body_constraints(60.0),
            &ParagraphStyle {
                alignment: Alignment::Both,
                tabs: vec![TabStopDef {
                    position: Pt::new(15.0),
                    alignment: crate::model::TabAlignment::Left,
                    leader: crate::model::TabLeader::None,
                }],
                ..Default::default()
            },
            Pt::new(14.0),
            None,
        );

        let xs = text_xs(&result);
        assert!(
            (xs[1] - 40.0).abs() < 0.01,
            "the first picture-bullet line must justify after its synthetic tab: {xs:?}",
        );
    }

    #[test]
    fn both_alignment_keeps_literal_tab_text_lines_at_natural_width() {
        let style = ParagraphStyle {
            alignment: Alignment::Both,
            ..Default::default()
        };
        let fragments = vec![
            underlined_text_frag("alpha ", 25.0),
            text_frag("\t", 5.0),
            text_frag("beta ", 25.0),
            text_frag("gamma", 30.0),
        ];
        let result = layout_paragraph(
            &fragments,
            &body_constraints(60.0),
            &style,
            Pt::new(14.0),
            None,
        );
        assert!((text_xs(&result)[2] - 30.0).abs() < 0.01);
        let underline = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Underline { line, .. } => Some(*line),
                _ => None,
            })
            .expect("underlined literal-tab gap");
        assert!((underline.end.x.raw() - 25.0).abs() < 0.01);
    }

    #[test]
    fn both_alignment_stretches_hyperlink_run_without_changing_text_order() {
        let style = ParagraphStyle {
            alignment: Alignment::Both,
            ..Default::default()
        };
        let result = layout_paragraph(
            &[
                text_frag("alpha", 10.0),
                hyperlink_frag("beta ", 25.0, "https://example.invalid/stretched"),
                text_frag("gamma ", 15.0),
                text_frag("delta", 30.0),
            ],
            &body_constraints(60.0),
            &style,
            Pt::new(14.0),
            None,
        );
        assert!((text_xs(&result)[2] - 45.0).abs() < 0.01);
        let text_runs: Vec<_> = result
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, .. } => Some(text.to_string()),
                _ => None,
            })
            .collect();
        assert_eq!(text_runs, ["alpha", "beta ", "gamma ", "delta"]);
        let link_rect = result
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::LinkAnnotation { rect, url }
                    if &**url == "https://example.invalid/stretched" =>
                {
                    Some(*rect)
                }
                _ => None,
            })
            .expect("stretched hyperlink annotation");
        assert!((link_rect.origin.x.raw() - 10.0).abs() < 0.01);
        assert!((link_rect.size.width.raw() - 35.0).abs() < 0.01);
    }

    #[test]
    fn start_center_and_end_alignment_positions_are_unchanged() {
        let fragments = vec![text_frag("alpha ", 25.0), text_frag("beta", 25.0)];
        for (alignment, expected_x) in [
            (Alignment::Start, 0.0),
            (Alignment::Center, 5.0),
            (Alignment::End, 10.0),
        ] {
            let style = ParagraphStyle {
                alignment,
                ..Default::default()
            };
            let result = layout_paragraph(
                &fragments,
                &body_constraints(60.0),
                &style,
                Pt::new(14.0),
                None,
            );
            assert!((text_xs(&result)[0] - expected_x).abs() < 0.01);
        }
    }

    #[test]
    fn indentation_shifts_text() {
        let frags = vec![text_frag("text", 40.0)];
        let style = ParagraphStyle {
            indent_left: Pt::new(36.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 36.0);
        }
    }

    #[test]
    fn first_line_indent() {
        let frags = vec![text_frag("first ", 40.0), text_frag("second", 40.0)];
        let style = ParagraphStyle {
            indent_first_line: Pt::new(24.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert_eq!(position.x.raw(), 24.0, "first line indented");
        }
    }

    #[test]
    fn space_before_and_after() {
        let frags = vec![text_frag("text", 30.0)];
        let style = ParagraphStyle {
            space_before: Pt::new(10.0),
            space_after: Pt::new(8.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        // Height should be: space_before(10) + line_height(14) + space_after(8) = 32
        assert_eq!(result.size.height.raw(), 32.0);

        // Text y should include space_before
        if let DrawCommand::Text { position, .. } = &result.commands[0] {
            assert!(
                position.y.raw() >= 10.0,
                "y should account for space_before"
            );
        }
    }

    #[test]
    fn line_spacing_exact() {
        let frags = vec![text_frag("line1 ", 60.0), text_frag("line2", 60.0)];
        let style = ParagraphStyle {
            line_spacing: LineSpacingRule::Exact(Pt::new(20.0)),
            ..Default::default()
        };
        // With max_width=80, they'll break into 2 lines
        let result = layout_paragraph(&frags, &body_constraints(80.0), &style, Pt::new(14.0), None);

        assert_eq!(result.size.height.raw(), 40.0, "2 lines * 20pt each");
    }

    #[test]
    fn line_spacing_at_least_with_larger_natural() {
        let frags = vec![text_frag("text", 30.0)];
        let style = ParagraphStyle {
            line_spacing: LineSpacingRule::AtLeast(Pt::new(10.0)),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        // Natural height is 14, at-least is 10 → should be 14
        assert_eq!(result.size.height.raw(), 14.0);
    }

    #[test]
    fn wrapping_produces_multiple_lines() {
        let frags = vec![
            text_frag("word1 ", 45.0),
            text_frag("word2 ", 45.0),
            text_frag("word3", 45.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(80.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );

        // Should have 3 text commands (one per word, each on its own line)
        let text_count = result
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 3);
        // Height: 3 lines * 14pt = 42pt
        assert_eq!(result.size.height.raw(), 42.0);
    }

    fn text_ys(commands: &[DrawCommand]) -> Vec<f32> {
        commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some(position.y.raw()),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn emit_split_segment_partitions_full_emit() {
        // Four words, each wrapping to its own 14pt line, with space_before and
        // space_after so we can check they land on the right segments.
        let frags = vec![
            text_frag("word1 ", 45.0),
            text_frag("word2 ", 45.0),
            text_frag("word3 ", 45.0),
            text_frag("word4", 45.0),
        ];
        let style = ParagraphStyle {
            space_before: Pt::new(10.0),
            space_after: Pt::new(6.0),
            ..Default::default()
        };
        let constraints = body_constraints(80.0);
        let placed = place_paragraph(&frags, &constraints, &style, Pt::new(14.0), None);
        assert_eq!(placed.line_count(), 4);
        assert_eq!(placed.unbreakable_prefix_lines(), 0, "no drop cap / floats");

        let full = placed.emit_full();
        let head = placed.emit_split_segment(0, 2, true, false);
        let tail = placed.emit_split_segment(2, 4, false, true);

        // Content preserved: 4 text commands total, 2 per segment.
        assert_eq!(text_ys(&full.commands).len(), 4);
        assert_eq!(text_ys(&head.commands).len(), 2);
        assert_eq!(text_ys(&tail.commands).len(), 2);

        // §17.3.1.33: the two segment heights reconstruct the whole height
        // (space_before on the head, space_after on the tail, no double count).
        assert_eq!(
            (head.size.height + tail.size.height).raw(),
            full.size.height.raw()
        );

        // The head reproduces the first lines verbatim; the tail reproduces the
        // rest, shifted up by the head's height.
        let full_ys = text_ys(&full.commands);
        assert_eq!(text_ys(&head.commands).as_slice(), &full_ys[..2]);
        for (i, y) in text_ys(&tail.commands).iter().enumerate() {
            assert_eq!(*y + head.size.height.raw(), full_ys[2 + i]);
        }
    }

    fn border(width: f32) -> BorderLine {
        BorderLine {
            width: Pt::new(width),
            color: RgbColor::BLACK,
            space: Pt::new(2.0),
        }
    }

    /// (horizontal border lines, vertical border lines) in a segment.
    fn classify_border_lines(commands: &[DrawCommand]) -> (usize, usize) {
        let mut horizontal = 0;
        let mut vertical = 0;
        for c in commands {
            if let DrawCommand::Line { line, .. } = c {
                if (line.start.y.raw() - line.end.y.raw()).abs() < 1e-3 {
                    horizontal += 1;
                } else if (line.start.x.raw() - line.end.x.raw()).abs() < 1e-3 {
                    vertical += 1;
                }
            }
        }
        (horizontal, vertical)
    }

    fn rect_count(commands: &[DrawCommand]) -> usize {
        commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Rect { .. }))
            .count()
    }

    fn horizontal_line_y(commands: &[DrawCommand]) -> Option<f32> {
        commands.iter().find_map(|c| match c {
            DrawCommand::Line { line, .. }
                if (line.start.y.raw() - line.end.y.raw()).abs() < 1e-3 =>
            {
                Some(line.start.y.raw())
            }
            _ => None,
        })
    }

    #[test]
    fn float_wrapped_lines_form_an_unbreakable_prefix() {
        // §17.3.1.44: an active float covering y[0,30) narrows the first three
        // 14pt lines (y 0, 14, 28). Those float-wrapped lines must stay on the
        // first segment so the full-width continuation reuses its fitted lines.
        use crate::render::layout::float::{ActiveFloat, FloatSource, WrapTextSide};
        let float = ActiveFloat {
            page_x: Pt::ZERO,
            page_y_start: Pt::ZERO,
            page_y_end: Pt::new(30.0),
            width: Pt::new(60.0),
            source: FloatSource::Image,
            wrap_text: WrapTextSide::BothSides,
        };
        let style = ParagraphStyle {
            page_floats: vec![float],
            page_x: Pt::ZERO,
            page_content_width: Pt::new(120.0),
            ..Default::default()
        };
        let frags: Vec<Fragment> = (0..20)
            .map(|i| text_frag(&format!("w{i} "), 20.0))
            .collect();
        let constraints = body_constraints(120.0);
        let placed = place_paragraph(&frags, &constraints, &style, Pt::new(14.0), None);

        assert!(placed.line_count() >= 4, "enough lines to clear the float");
        assert_eq!(
            placed.unbreakable_prefix_lines(),
            3,
            "the three float-narrowed lines are the unbreakable prefix"
        );
    }

    #[test]
    fn split_segment_draws_border_box_per_page() {
        // §17.3.1.24/§17.3.1.31: a bordered, shaded paragraph split across three
        // pages keeps side borders and shading on every segment, the top border
        // only on the first, and the bottom border only on the last.
        let frags: Vec<Fragment> = (0..6)
            .map(|i| text_frag(&format!("word{i} "), 100.0))
            .collect();
        let style = ParagraphStyle {
            borders: Some(ParagraphBorderStyle {
                top: Some(border(1.0)),
                bottom: Some(border(1.0)),
                left: Some(border(1.0)),
                right: Some(border(1.0)),
            }),
            shading: Some(RgbColor::BLACK),
            ..Default::default()
        };
        // Box wider than a word (with room for the 2pt side border spaces) so
        // each word is its own line without triggering per-character splitting.
        let constraints = body_constraints(120.0);
        let placed = place_paragraph(&frags, &constraints, &style, Pt::new(14.0), None);
        assert_eq!(
            placed.unbreakable_prefix_lines(),
            0,
            "bordered/shaded paragraphs split freely (no drop cap / floats)"
        );
        assert_eq!(placed.line_count(), 6);

        let head = placed.emit_split_segment(0, 2, true, false);
        let mid = placed.emit_split_segment(2, 4, false, false);
        let tail = placed.emit_split_segment(4, 6, false, true);

        // Shading fills every segment.
        assert_eq!(rect_count(&head.commands), 1);
        assert_eq!(rect_count(&mid.commands), 1);
        assert_eq!(rect_count(&tail.commands), 1);

        // Sides on every segment; top only first, bottom only last.
        assert_eq!(
            classify_border_lines(&head.commands),
            (1, 2),
            "first: top + sides"
        );
        assert_eq!(
            classify_border_lines(&mid.commands),
            (0, 2),
            "middle: sides only"
        );
        assert_eq!(
            classify_border_lines(&tail.commands),
            (1, 2),
            "last: bottom + sides"
        );

        // The first segment's horizontal edge sits above its text (the top
        // border, with border space above space_before); the last segment's sits
        // below its two lines (the bottom border).
        assert!(horizontal_line_y(&head.commands).unwrap() < 0.0);
        assert!(horizontal_line_y(&tail.commands).unwrap() > 28.0);
    }

    #[test]
    fn resolve_line_height_auto_text_only() {
        // Text-only line: multiplier applies to text_height.
        assert_eq!(
            resolve_line_height(
                Pt::new(14.0),
                Pt::new(14.0),
                &LineSpacingRule::Auto(1.0),
                crate::render::layout::ShapeAutoFit::NONE
            )
            .raw(),
            14.0
        );
        assert_eq!(
            resolve_line_height(
                Pt::new(14.0),
                Pt::new(14.0),
                &LineSpacingRule::Auto(1.5),
                crate::render::layout::ShapeAutoFit::NONE
            )
            .raw(),
            21.0
        );
    }

    #[test]
    fn resolve_line_height_auto_image_line() {
        // Image-only line: natural=325 (image), text_height=0 (no text).
        // The multiplier does NOT inflate the image height.
        let h = resolve_line_height(
            Pt::new(325.0),
            Pt::ZERO,
            &LineSpacingRule::Auto(1.08),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert_eq!(h.raw(), 325.0, "image height should not be multiplied");
    }

    #[test]
    fn resolve_line_height_auto_mixed_line() {
        // Line with text (14pt) and image (100pt): multiplier scales text only.
        // max(14*1.5=21, 100) = 100.
        let h = resolve_line_height(
            Pt::new(100.0),
            Pt::new(14.0),
            &LineSpacingRule::Auto(1.5),
            crate::render::layout::ShapeAutoFit::NONE,
        );
        assert_eq!(h.raw(), 100.0, "image dominates");
    }

    #[test]
    fn resolve_line_height_exact_overrides() {
        assert_eq!(
            resolve_line_height(
                Pt::new(14.0),
                Pt::new(14.0),
                &LineSpacingRule::Exact(Pt::new(20.0)),
                crate::render::layout::ShapeAutoFit::NONE,
            )
            .raw(),
            20.0
        );
    }

    #[test]
    fn resolve_line_height_at_least() {
        assert_eq!(
            resolve_line_height(
                Pt::new(14.0),
                Pt::new(14.0),
                &LineSpacingRule::AtLeast(Pt::new(10.0)),
                crate::render::layout::ShapeAutoFit::NONE,
            )
            .raw(),
            14.0,
            "natural > minimum"
        );
        assert_eq!(
            resolve_line_height(
                Pt::new(8.0),
                Pt::new(8.0),
                &LineSpacingRule::AtLeast(Pt::new(10.0)),
                crate::render::layout::ShapeAutoFit::NONE,
            )
            .raw(),
            10.0,
            "minimum > natural"
        );
    }

    // ── Absolute position tabs (§17.3.1.30) ──

    fn ptab_frag(align: PTabAlignment, relative_to: PTabRelativeTo) -> Fragment {
        Fragment::PTab {
            align,
            relative_to,
            leader: crate::model::TabLeader::None,
            line_height: Pt::new(12.0),
            font: leader_font("Test", 12.0),
            color: RgbColor::BLACK,
        }
    }

    /// Return the x positions of every emitted Text command, in order.
    fn text_xs(result: &ParagraphLayout) -> Vec<f32> {
        result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some(position.x.raw()),
                _ => None,
            })
            .collect()
    }

    /// Return `(text, x, y)` for every emitted Text command, in order.
    fn text_positions(result: &ParagraphLayout) -> Vec<(String, f32, f32)> {
        result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, position, .. } => {
                    Some((text.to_string(), position.x.raw(), position.y.raw()))
                }
                _ => None,
            })
            .collect()
    }

    #[test]
    fn ptab_whose_anchor_is_passed_advances_to_the_next_line() {
        // §17.3.1.30: when the alignment point has already been passed, the
        // tab advances to that location on the *next* line.
        //
        // Reproduction from the branch review: everything fits the 400pt line
        // by width (10 + 200 + 180 = 390), but the right ptab's anchor sits at
        // 400 - 180 = 220, behind the pen at 300. Clamping drew "C" at
        // 300..480 — 80pt off the page.
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(PTabAlignment::Center, PTabRelativeTo::Margin),
            text_frag("B", 200.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Margin),
            text_frag("C", 180.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 3, "no content dropped: {positions:?}");
        let (_, c_x, c_y) = &positions[2];
        let (_, _, a_y) = &positions[0];

        assert!(
            c_y > a_y,
            "C advances to the next line rather than overflowing: {positions:?}"
        );
        assert_eq!(*c_x, 220.0, "and lands on its anchor: 400 - 180");
        assert!(
            c_x + 180.0 <= 400.0,
            "C ends within the page: {}",
            c_x + 180.0
        );
    }

    #[test]
    fn left_margin_ptab_pulls_back_to_the_margin_on_the_next_line() {
        // §17.3.1.30: `left` + `relativeTo="margin"` anchors at the page
        // margin (x=0). Reached from an indented paragraph the anchor is
        // behind the pen, so — same rule — it advances to the next line
        // rather than becoming the silent no-op `0.max(x)` produced.
        let frags = vec![
            text_frag("indented", 60.0),
            ptab_frag(PTabAlignment::Left, PTabRelativeTo::Margin),
            text_frag("at-margin", 40.0),
        ];
        let style = ParagraphStyle {
            indent_left: Pt::new(100.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 2, "{positions:?}");
        let (_, first_x, first_y) = &positions[0];
        let (_, second_x, second_y) = &positions[1];

        assert!(second_y > first_y, "advances to the next line");
        assert_eq!(*second_x, 0.0, "and reaches the page margin");
        assert!(*first_x >= 100.0, "the indented run is unaffected");
    }

    #[test]
    fn a_ptab_at_the_line_start_places_rather_than_looping() {
        // The progress guard: when the tab is already the first thing on its
        // line, advancing to the next line cannot move the anchor any closer,
        // so it places (clamped) and overflows instead. Acting on a condition
        // the action cannot change is how the paginator bugs in this engine
        // have historically become infinite loops.
        // Three same-width runs, none individually oversized (so nothing is
        // split), but a zone far wider than the centre anchor — the tab's
        // desired x is negative and it is already at the line start.
        let frags = vec![
            ptab_frag(PTabAlignment::Center, PTabRelativeTo::Margin),
            text_frag("one", 90.0),
            text_frag("two", 90.0),
            text_frag("three", 90.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(
            positions.len(),
            3,
            "every run is emitted, none dropped: {positions:?}"
        );
        assert!(
            positions.iter().all(|(_, x, _)| *x >= 0.0),
            "nothing is placed left of the line start: {positions:?}"
        );
    }

    #[test]
    fn fitting_and_emission_agree_about_the_pen_on_an_indented_line() {
        // Fitting decides whether a ptab can be honoured on this line;
        // emission decides its x. Both must start the line at the same place,
        // or an indented paragraph gets classified one way and drawn another.
        //
        // The fitter seeded its pen at 0 while emission seeds it at the
        // indent, so the two could classify the same tab differently. Here the
        // right-margin anchor for a 150pt zone is 250: ahead of the fitter's
        // pen (200) but behind emission's (300). The fitter therefore saw no
        // reason to break, and emission — unable to break — clamped to 300,
        // running B out to 450, past the 400pt margin B1 was meant to protect.
        let frags = vec![
            text_frag("A", 200.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Margin),
            text_frag("B", 150.0),
        ];
        let style = ParagraphStyle {
            indent_left: Pt::new(100.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 2, "{positions:?}");
        assert!(
            positions[1].2 > positions[0].2,
            "B advances to the next line: {positions:?}"
        );
        assert_eq!(positions[1].1, 250.0, "and lands on its anchor: 400 - 150");
        assert!(positions[1].1 + 150.0 <= 400.0, "B stays within the margin");
    }

    #[test]
    fn a_margin_ptab_may_use_the_space_a_right_indent_excludes() {
        // §17.3.1.30: `relativeTo="margin"` measures against the full text
        // area, so its zone may occupy space the paragraph's own right indent
        // excludes. Fitting bounds lines by `content_width` (net of indents),
        // which wrapped such a line even though the content fits the margin
        // region — leaving a stray leader on the short line.
        //
        // content_width = 400 - 100 (right indent) = 300, but A(50) + B(320)
        // reaches only 400, exactly the margin.
        let frags = vec![
            text_frag("A", 50.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Margin),
            text_frag("B", 320.0),
        ];
        let style = ParagraphStyle {
            indent_right: Pt::new(100.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &style,
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 2, "{positions:?}");
        assert_eq!(
            positions[0].2, positions[1].2,
            "the line fits the margin region and must not wrap: {positions:?}"
        );
        assert_eq!(
            positions[1].1, 80.0,
            "B right-aligned at the margin: 400 - 320"
        );
    }

    #[test]
    fn an_indent_ptab_respects_the_lines_own_float_narrowing() {
        // §17.3.1.30 + §20.4.2: `relativeTo="indent"` measures against the
        // indented region, but that region is narrower on a line a float
        // overlaps. Using the paragraph-level width right-aligned the zone to
        // the paragraph's edge — on top of the float.
        //
        // A 60pt float occupies x 140..200 of a 200pt line, so the line's own
        // right edge is 140 and a 40pt zone must end there, at x=100.
        use crate::render::layout::float::{ActiveFloat, FloatSource, WrapTextSide};
        let float = ActiveFloat {
            page_x: Pt::new(140.0),
            page_y_start: Pt::ZERO,
            page_y_end: Pt::new(30.0),
            width: Pt::new(60.0),
            source: FloatSource::Image,
            wrap_text: WrapTextSide::BothSides,
        };
        let style = ParagraphStyle {
            page_floats: vec![float],
            page_x: Pt::ZERO,
            page_content_width: Pt::new(200.0),
            ..Default::default()
        };
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Indent),
            text_frag("B", 40.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(200.0),
            &style,
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 2, "{positions:?}");
        assert_eq!(
            positions[1].1, 100.0,
            "B right-aligns to the line's own edge (140), not the paragraph's (200)"
        );
        assert!(
            positions[1].1 + 40.0 <= 140.0,
            "and therefore clears the float at 140: {positions:?}"
        );
    }

    /// A float occupying `x0..x0+60` of a 200pt line for y `0..30`.
    fn float_at(x0: f32) -> crate::render::layout::float::ActiveFloat {
        use crate::render::layout::float::{ActiveFloat, FloatSource, WrapTextSide};
        ActiveFloat {
            page_x: Pt::new(x0),
            page_y_start: Pt::ZERO,
            page_y_end: Pt::new(30.0),
            width: Pt::new(60.0),
            source: FloatSource::Image,
            wrap_text: WrapTextSide::BothSides,
        }
    }

    fn style_with_float(x0: f32) -> ParagraphStyle {
        ParagraphStyle {
            page_floats: vec![float_at(x0)],
            page_x: Pt::ZERO,
            page_content_width: Pt::new(200.0),
            ..Default::default()
        }
    }

    #[test]
    fn an_indent_ptab_follows_a_left_float_too() {
        // The mirror of the right-float case: a float on the left moves the
        // indented region's *start*, so a left-aligned indent ptab must anchor
        // clear of it rather than at the paragraph's own left edge.
        let frags = vec![
            ptab_frag(PTabAlignment::Left, PTabRelativeTo::Indent),
            text_frag("B", 40.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(200.0),
            &style_with_float(0.0),
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 1, "{positions:?}");
        assert_eq!(
            positions[0].1, 60.0,
            "B anchors past the float, not at the paragraph's left edge"
        );
    }

    #[test]
    fn a_margin_ptab_is_deliberately_unaffected_by_floats() {
        // `margin` names the page text margins, which a float does not move.
        // Only `indent` follows the line's narrowed region — pinning the
        // difference so the two do not get "unified" later.
        let frags = vec![
            text_frag("A", 10.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Margin),
            text_frag("B", 40.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(200.0),
            &style_with_float(140.0),
            Pt::new(14.0),
            None,
        );

        let positions = text_positions(&result);
        assert_eq!(positions.len(), 2, "{positions:?}");
        assert_eq!(
            positions[1].1, 160.0,
            "right margin anchor stays at 200 — 200 - 40 — despite the float"
        );
    }

    #[test]
    fn ptab_three_region_header() {
        // The canonical Word header: left ⟶ ptab(center) ⟶ center ⟶
        // ptab(right) ⟶ right, all relative to the page margins.
        let frags = vec![
            text_frag("L", 10.0),
            ptab_frag(PTabAlignment::Center, PTabRelativeTo::Margin),
            text_frag("C", 20.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Margin),
            text_frag("R", 30.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );

        let xs = text_xs(&result);
        assert_eq!(xs.len(), 3, "three text runs, no leaders");
        assert_eq!(xs[0], 0.0, "left run at the left margin");
        assert_eq!(xs[1], 40.0, "center run centered on 50 → 50 - 20/2");
        assert_eq!(xs[2], 70.0, "right run ends at the right margin → 100 - 30");
    }

    #[test]
    fn ptab_right_relative_to_indent() {
        // relativeTo=indent uses [indent_left, content_width - indent_right].
        let frags = vec![
            text_frag("x", 10.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Indent),
            text_frag("yy", 20.0),
        ];
        let style = ParagraphStyle {
            indent_right: Pt::new(30.0),
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &style,
            Pt::new(14.0),
            None,
        );

        let xs = text_xs(&result);
        // span_end = 100 - 30 = 70; right run of width 20 ends there → 50.
        assert_eq!(xs[1], 50.0, "right run ends at the right indent");
    }

    #[test]
    fn ptab_suppresses_paragraph_alignment() {
        // A line with a ptab is positioned by the ptab, not by the paragraph's
        // alignment (§17.3.1.37 rationale) — the left run stays at the margin.
        let frags = vec![
            text_frag("L", 10.0),
            ptab_frag(PTabAlignment::Right, PTabRelativeTo::Margin),
            text_frag("R", 30.0),
        ];
        let style = ParagraphStyle {
            alignment: Alignment::End,
            ..Default::default()
        };
        let result = layout_paragraph(
            &frags,
            &body_constraints(100.0),
            &style,
            Pt::new(14.0),
            None,
        );

        let xs = text_xs(&result);
        assert_eq!(xs[0], 0.0, "left run not shifted by End alignment");
        assert_eq!(xs[1], 70.0, "right run ends at the right margin");
    }

    // ── Decimal tab alignment (§17.18.85 `ST_TabJc` = decimal) ──
    //
    // A decimal stop aligns the *separator* on the stop, not an edge of the
    // zone. These use a 10pt-per-character measurer so the expected x values
    // are exact arithmetic rather than font-dependent.

    /// Measure every character as 10pt wide, so a prefix's width is
    /// `10 * prefix.chars().count()`.
    fn ten_pt_per_char(text: &str, _f: &FontProps) -> (Pt, TextMetrics) {
        (
            Pt::new(10.0 * text.chars().count() as f32),
            TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
        )
    }

    fn decimal_stop_at(pos: f32) -> ParagraphStyle {
        ParagraphStyle {
            tabs: vec![TabStopDef {
                position: Pt::new(pos),
                alignment: crate::model::TabAlignment::Decimal,
                leader: crate::model::TabLeader::None,
            }],
            ..Default::default()
        }
    }

    /// Lay out `a<tab>zone` against a decimal stop at 100 and return the x of
    /// the zone's text.
    fn decimal_zone_x(zone: &str) -> f32 {
        let measurer: &dyn Fn(&str, &FontProps) -> (Pt, TextMetrics) = &ten_pt_per_char;
        let frags = vec![
            text_frag("a", 10.0),
            Fragment::Tab {
                line_height: Pt::new(12.0),
                font: leader_font("Test", 12.0),
                color: RgbColor::BLACK,
                fitting_width: None,
            },
            text_frag(zone, 10.0 * zone.chars().count() as f32),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(400.0),
            &decimal_stop_at(100.0),
            Pt::new(14.0),
            Some(measurer),
        );
        *text_xs(&result).last().expect("zone text is emitted")
    }

    #[test]
    fn decimal_tab_puts_the_separator_on_the_stop() {
        // §17.18.85: the decimal separator lands on the stop, whatever the
        // integer part's width. "1.5" has one char before the separator,
        // "22.75" has two — so they start 10pt apart and their separators
        // coincide at x=100.
        assert_eq!(decimal_zone_x("1.5"), 90.0, "one char before the separator");
        assert_eq!(
            decimal_zone_x("22.75"),
            80.0,
            "two chars before the separator"
        );
        assert_eq!(
            decimal_zone_x("333.125"),
            70.0,
            "three chars before the separator"
        );
    }

    #[test]
    fn decimal_tab_without_a_separator_right_aligns() {
        // Word right-aligns a decimal-tabbed zone that has no separator —
        // aligning its left edge would leave numeric columns ragged.
        assert_eq!(
            decimal_zone_x("1234"),
            60.0,
            "4 chars ending at the stop → 100 - 40"
        );
    }

    #[test]
    fn decimal_tab_aligns_on_the_first_separator() {
        // A version string has several; the first is the alignment point.
        assert_eq!(decimal_zone_x("1.2.3"), 90.0);
    }

    #[test]
    fn decimal_tab_never_moves_content_backwards() {
        // The zone is wider than the stop, so honouring the separator would
        // place text left of the current pen. Clamp, as right/center do.
        assert_eq!(
            decimal_zone_x("1234567890123.5"),
            10.0,
            "separator would land at -30; clamped to the pen after 'a'"
        );
    }

    // ── Tab leaders (§17.3.1.38) ──
    //
    // A leader carries no formatting of its own: §17.3.1.38 defines it as
    // filling the tab with a character, and the character is drawn in the
    // formatting in effect at the tab — i.e. the `<w:rPr>` of the run holding
    // the `<w:tab/>`. These assert that, not the paragraph default.

    /// Build a run font differing from every default, so a leader that
    /// inherits it cannot pass by coincidence.
    fn leader_font(family: &str, size: f32) -> Rc<FontProps> {
        Rc::new(FontProps {
            rtl: crate::render::fonts::Toggle::Absent,
            family: Rc::from(family),
            size: Pt::new(size),
            bold: Toggle::Absent,
            italic: Toggle::Absent,
            underline: false,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        })
    }

    fn dotted_stop_at(pos: f32) -> ParagraphStyle {
        ParagraphStyle {
            tabs: vec![TabStopDef {
                position: Pt::new(pos),
                alignment: crate::model::TabAlignment::Left,
                leader: crate::model::TabLeader::Dot,
            }],
            ..Default::default()
        }
    }

    /// The first Text command whose text is all leader dots.
    fn leader_command(result: &ParagraphLayout) -> &DrawCommand {
        result
            .commands
            .iter()
            .find(|c| match c {
                DrawCommand::Text { text, .. } => {
                    !text.is_empty() && text.chars().all(|c| c == '.')
                }
                _ => false,
            })
            .expect("a dotted tab stop must emit a leader command")
    }

    #[test]
    fn tab_leader_inherits_the_run_font_and_colour() {
        // §17.3.1.38: the leader is drawn in the tab run's own formatting.
        // Previously it was hardcoded to Times New Roman / black.
        let frags = vec![
            text_frag("a", 10.0),
            Fragment::Tab {
                line_height: Pt::new(12.0),
                font: leader_font("Georgia", 12.0),
                color: RgbColor { r: 255, g: 0, b: 0 },
                fitting_width: None,
            },
            text_frag("b", 10.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(200.0),
            &dotted_stop_at(100.0),
            Pt::new(14.0),
            None,
        );

        match leader_command(&result) {
            DrawCommand::Text {
                font_family, color, ..
            } => {
                assert_eq!(
                    &**font_family, "Georgia",
                    "leader takes the run's family, not a hardcoded one"
                );
                assert_eq!(
                    *color,
                    RgbColor { r: 255, g: 0, b: 0 },
                    "leader takes the run's colour, not black"
                );
            }
            _ => unreachable!("leader_command returns a Text command"),
        }
    }

    #[test]
    fn tab_leader_scales_with_the_run_font_size() {
        // The old code sized the leader from the paragraph line height and
        // capped it at 12pt, so a large run drew undersized leaders.
        let frags = vec![
            text_frag("a", 10.0),
            Fragment::Tab {
                line_height: Pt::new(24.0),
                font: leader_font("Georgia", 24.0),
                color: RgbColor::BLACK,
                fitting_width: None,
            },
            text_frag("b", 10.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(200.0),
            &dotted_stop_at(100.0),
            Pt::new(28.0),
            None,
        );

        match leader_command(&result) {
            DrawCommand::Text { font_size, .. } => assert_eq!(
                font_size.raw(),
                24.0,
                "a 24pt run gets 24pt leader dots, not a 12pt cap"
            ),
            _ => unreachable!(),
        }
    }

    #[test]
    fn ptab_leader_inherits_the_run_font_and_colour() {
        // Same rule on the §17.3.1.30 position-tab arm, which carries its
        // leader on the fragment rather than on a tab stop.
        let frags = vec![
            text_frag("L", 10.0),
            Fragment::PTab {
                align: PTabAlignment::Right,
                relative_to: PTabRelativeTo::Margin,
                leader: crate::model::TabLeader::Dot,
                line_height: Pt::new(12.0),
                font: leader_font("Georgia", 18.0),
                color: RgbColor { r: 0, g: 128, b: 0 },
            },
            text_frag("R", 30.0),
        ];
        let result = layout_paragraph(
            &frags,
            &body_constraints(200.0),
            &ParagraphStyle::default(),
            Pt::new(14.0),
            None,
        );

        match leader_command(&result) {
            DrawCommand::Text {
                font_family,
                font_size,
                color,
                ..
            } => {
                assert_eq!(&**font_family, "Georgia");
                assert_eq!(font_size.raw(), 18.0);
                assert_eq!(*color, RgbColor { r: 0, g: 128, b: 0 });
            }
            _ => unreachable!(),
        }
    }
}
