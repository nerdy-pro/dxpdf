//! Shared block stacker — used by both page-level and cell-level layout.

use super::super::draw_command::DrawCommand;
use super::super::float;
use super::super::paragraph::place_paragraph;
use super::super::table::layout_table;
use super::helpers::table_x_offset;
use super::types::{LayoutBlock, PageParity};
use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

/// One fitted line of stacked cell content, recorded so a §17.4.1 row split can
/// choose legal cut points from the paragraph structure rather than from raw
/// draw commands.
///
/// Cutting a cell "after this line" flows it and everything below to the
/// continuation page. Legality mirrors body across-page paragraph splitting:
/// §17.3.1.14 keepLines and borders/shading/drop-cap paragraphs are never cut
/// internally, §17.3.1.44 widow/orphan control forbids single-line segments of
/// a paragraph, and §17.3.1.15 keepNext forbids a cut at a paragraph's trailing
/// boundary. The [`stack_blocks`] producer leaves the line list empty for any
/// cell it cannot safely bisect (nested table, floating object).
#[derive(Debug, Clone)]
pub struct CellLine {
    /// Box top of this line, in cell-content coordinates (0 = content top,
    /// before the cell margin shift applied by `layout_cell`).
    pub top_y: Pt,
    /// Index of the paragraph (block) this line belongs to. Lines of one
    /// paragraph are contiguous and share widow/orphan and keepLines rules; a
    /// cut may fall between two paragraphs freely (unless the earlier one is
    /// keepNext).
    pub para: usize,
    /// §17.3.1.14 / §17.3.1.24 / §17.3.1.31 / §17.3.1.11: this line's paragraph
    /// forbids interior splits — keepLines, or a bordered / shaded / drop-cap
    /// paragraph whose box would be torn. Its lines may still be moved whole to
    /// the continuation, but never divided among themselves.
    pub interior_atomic: bool,
    /// §17.3.1.44: widow/orphan control is active for this paragraph, so an
    /// interior cut must leave `>= 2` of the paragraph's lines on each side.
    pub widow_control: bool,
    /// §17.3.1.15: this line's paragraph is kept with the following block, so a
    /// cut at the paragraph's trailing boundary is illegal.
    pub keep_next: bool,
}

/// Result of stacking blocks vertically.
pub struct StackResult {
    /// Draw commands positioned relative to the stacking origin (0,0).
    pub commands: Vec<DrawCommand>,
    /// Total height consumed by all blocks.
    pub height: Pt,
    /// Per-line cut model for §17.4.1 row splitting (cell-content coords).
    /// Empty when the content cannot be safely bisected (nested table or
    /// floating object present) — such cells move whole rather than split.
    pub lines: Vec<CellLine>,
}

/// Stack blocks vertically within a fixed-width area.
///
/// This is the shared core used by both page-level layout (`layout_section`)
/// and cell-level layout. It handles:
/// - Paragraph layout with spacing collapse and space_before suppression
/// - Table layout
/// - Floating image registration and text wrapping
///
/// It does NOT handle page breaks, column breaks, or footnote collection —
/// those are page-level concerns managed by `layout_section`.
///
/// That sharing is what makes a table cell behave like a miniature page: the
/// same code lays out cell content as body content, so spacing collapse and
/// float wrapping work identically in both.
///
/// `parity` resolves §20.4.3.1 `inside`/`outside` float positions, which mirror
/// on the page the object lands on. Callers that know their page pass its
/// parity; the table-cell path cannot (see [`layout_cell`]).
///
/// [`layout_cell`]: crate::render::layout::cell::layout_cell
pub fn stack_blocks(
    blocks: &[LayoutBlock],
    content_width: Pt,
    default_line_height: Pt,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
    parity: PageParity,
) -> StackResult {
    let constraints = super::super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
    let mut commands = Vec::new();
    let mut cursor_y = Pt::ZERO;
    let mut prev_space_after = Pt::ZERO;
    let mut prev_style_id: Option<crate::model::StyleId> = None;
    let mut page_floats: Vec<float::ActiveFloat> = Vec::new();
    // §17.4.1 row-split model: one entry per fitted line, plus a flag that a
    // block was encountered which makes the whole cell unsafe to bisect
    // (nested table or floating object).
    let mut cell_lines: Vec<CellLine> = Vec::new();
    let mut splittable = true;

    for (block_index, block) in blocks.iter().enumerate() {
        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                floating_images,
                floating_shapes,
                ..
            } => {
                let mut effective_style = style.clone_for_layout();

                // §17.3.1.9: spacing collapse.
                if effective_style.contextual_spacing
                    && effective_style.style_id.is_some()
                    && effective_style.style_id == prev_style_id
                {
                    cursor_y -= prev_space_after + effective_style.space_before;
                } else {
                    let collapse = prev_space_after.min(effective_style.space_before);
                    cursor_y -= collapse;
                }

                // Register floating images.
                let content_top = cursor_y + effective_style.space_before;
                // Where the paragraph would start with no `TopAndBottom` band,
                // so the band's effect on the cursor can be told apart from it.
                let pre_band_cursor = cursor_y;
                // §20.4.2.18: `wrapTopAndBottom` floats occupy a band of their
                // own above the paragraph and **stack** — each sits below the
                // one before it, and the paragraph starts below the whole band.
                // Tracks the band's running bottom so successive floats don't
                // all resolve to `content_top` and draw on top of each other.
                // Shared with the shape loop below, so a `TopAndBottom` image
                // and shape on the same paragraph stack against each other too.
                // `None` until the first one is placed, so a single float — and
                // any negative `wp:posOffset` on it — is positioned exactly as
                // before.
                let mut band_bottom: Option<Pt> = None;
                for fi in floating_images.iter() {
                    let y_start = fi.y.at(parity, content_top);
                    let y_end = y_start + fi.size.height;
                    if fi.is_wrap_top_and_bottom() {
                        let img_y = match fi.y.absolute(parity) {
                            // Page-absolute floats are positioned by the anchor,
                            // not by the band.
                            Some(y) => y,
                            None => match band_bottom {
                                Some(bottom) => y_start.max(bottom),
                                None => y_start,
                            },
                        };
                        commands.push(DrawCommand::Image {
                            rect: PtRect::from_xywh(
                                fi.x.resolve(parity),
                                img_y,
                                fi.size.width,
                                fi.size.height,
                            ),
                            image_data: fi.image_data.clone(),
                            src_rect: fi.src_rect,
                        });
                        let bottom = img_y + fi.size.height;
                        band_bottom = Some(match band_bottom {
                            Some(prev) => prev.max(bottom),
                            None => bottom,
                        });
                        if bottom > cursor_y {
                            cursor_y = bottom;
                        }
                    } else if fi.wrap_mode.registers_as_wrap_float() {
                        page_floats.push(float::ActiveFloat {
                            page_x: fi.x.resolve(parity) - fi.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fi.size.width + fi.dist_left + fi.dist_right,
                            source: float::FloatSource::Image,
                            wrap_text: fi.wrap_mode.wrap_text().into(),
                        });
                    }
                }

                // §20.4.2: register floating shapes (DrawingML). Mirrors the
                // image branch above: `TopAndBottom` emits now and advances
                // the cursor; wrap-enabled modes (Square/Tight/Through) are
                // registered as active floats so subsequent lines narrow
                // around them. `None` shapes emit after the paragraph.
                for fs in floating_shapes.iter() {
                    use crate::render::layout::section::WrapMode;
                    if matches!(fs.wrap_mode, WrapMode::None) {
                        continue;
                    }
                    let y_start = fs.y.at(parity, content_top);
                    let y_end = y_start + fs.size.height;
                    if fs.is_wrap_top_and_bottom() {
                        // Same band as the images above — see `band_bottom`.
                        let shape_y = match fs.y.absolute(parity) {
                            Some(y) => y,
                            None => match band_bottom {
                                Some(bottom) => y_start.max(bottom),
                                None => y_start,
                            },
                        };
                        commands.push(DrawCommand::Path {
                            origin: crate::render::geometry::PtOffset::new(
                                fs.x.resolve(parity),
                                shape_y,
                            ),
                            rotation: fs.rotation,
                            flip_h: fs.flip_h,
                            flip_v: fs.flip_v,
                            extent: fs.size,
                            paths: fs.paths.clone(),
                            fill: fs.fill.clone(),
                            stroke: fs.stroke.clone(),
                            effects: fs.effects.clone(),
                        });
                        let bottom = shape_y + fs.size.height;
                        band_bottom = Some(match band_bottom {
                            Some(prev) => prev.max(bottom),
                            None => bottom,
                        });
                        if bottom > cursor_y {
                            cursor_y = bottom;
                        }
                    } else {
                        page_floats.push(float::ActiveFloat {
                            page_x: fs.x.resolve(parity) - fs.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fs.size.width + fs.dist_left + fs.dist_right,
                            source: float::FloatSource::Shape,
                            wrap_text: fs.wrap_mode.wrap_text().into(),
                        });
                    }
                }

                // §20.4.2.18: the band was placed from `content_top`, which
                // already includes `space_before`, and the cursor now sits at
                // the band's bottom. `place_paragraph` applies `space_before`
                // again inside the paragraph box, so without this the text is
                // pushed a second `space_before` below the band it is supposed
                // to sit directly beneath — contradicting the band comment
                // above ("the paragraph starts below the whole band").
                //
                // Only when the band actually moved the cursor: a float with a
                // negative `wp:posOffset` that resolves above the paragraph
                // leaves the cursor alone and needs no correction.
                if cursor_y > pre_band_cursor {
                    cursor_y -= effective_style.space_before;
                }

                float::prune_floats(&mut page_floats, cursor_y);

                effective_style.page_floats = page_floats.clone();
                effective_style.page_y = cursor_y;
                effective_style.page_x = Pt::ZERO;
                effective_style.page_content_width = content_width;

                let placed = place_paragraph(
                    fragments,
                    &constraints,
                    &effective_style,
                    default_line_height,
                    measure_text,
                );

                // §17.4.1: record this paragraph's lines for the row-split cut
                // model. A floating object anchored here makes the whole cell
                // unsafe to bisect (per-line float offsets depend on absolute
                // y); a paragraph whose box would tear (keepLines, borders,
                // shading, drop cap) is kept internally atomic.
                if !floating_images.is_empty() || !floating_shapes.is_empty() {
                    splittable = false;
                } else if splittable {
                    let interior_atomic = effective_style.keep_lines
                        || effective_style.borders.is_some()
                        || effective_style.shading.is_some()
                        || effective_style.drop_cap.is_some();
                    let mut line_top = cursor_y + effective_style.space_before;
                    for i in 0..placed.line_count() {
                        cell_lines.push(CellLine {
                            top_y: line_top,
                            para: block_index,
                            interior_atomic,
                            widow_control: effective_style.widow_control,
                            keep_next: effective_style.keep_next,
                        });
                        line_top += placed.line_height(i);
                    }
                }

                let para = placed.emit_full();

                for mut cmd in para.commands {
                    cmd.shift_y(cursor_y);
                    commands.push(cmd);
                }

                cursor_y += para.size.height;

                // Emit non-wrapTopAndBottom floating images.
                let para_content_top = cursor_y - para.size.height + effective_style.space_before;
                for fi in floating_images {
                    if fi.is_wrap_top_and_bottom() {
                        continue;
                    }
                    let img_y = fi.y.at(parity, para_content_top);
                    commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(
                            fi.x.resolve(parity),
                            img_y,
                            fi.size.width,
                            fi.size.height,
                        ),
                        image_data: fi.image_data.clone(),
                        src_rect: fi.src_rect,
                    });
                    // Extend cursor to encompass the image so table cells
                    // expand to contain floating images.
                    let img_bottom = img_y + fi.size.height;
                    if img_bottom > cursor_y {
                        cursor_y = img_bottom;
                    }
                }

                // Emit floating shapes (DrawingML). `TopAndBottom` shapes
                // were emitted pre-layout along with the `cursor_y` advance;
                // skip them here. `None` + Square/Tight/Through emit now at
                // their resolved anchor position (the shape's bounding rect
                // is already registered as an active float for wrap modes).
                for fs in floating_shapes {
                    if fs.is_wrap_top_and_bottom() {
                        continue;
                    }
                    let shape_y = fs.y.at(parity, para_content_top);
                    commands.push(DrawCommand::Path {
                        origin: crate::render::geometry::PtOffset::new(
                            fs.x.resolve(parity),
                            shape_y,
                        ),
                        rotation: fs.rotation,
                        flip_h: fs.flip_h,
                        flip_v: fs.flip_v,
                        extent: fs.size,
                        paths: fs.paths.clone(),
                        fill: fs.fill.clone(),
                        stroke: fs.stroke.clone(),
                        effects: fs.effects.clone(),
                    });
                    // §17.17.1 / §20.1.2.1.1: shape's text-box content paints
                    // *over* the shape's fill — emit after the path. Each
                    // command is in shape-local coords; shift by the shape's
                    // resolved page origin.
                    for mut cmd in fs.text_commands.iter().cloned() {
                        cmd.shift(fs.x.resolve(parity), shape_y);
                        commands.push(cmd);
                    }
                }

                prev_space_after = effective_style.space_after;
                prev_style_id = effective_style.style_id.clone();
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                cell_spacing,
                border_config,
                indent,
                alignment,
                ..
            } => {
                // §17.4.1: a nested table can't be cleanly bisected, so the
                // enclosing cell must move whole (matches `build_row_groups`,
                // which marks rows with nested tables non-splittable).
                splittable = false;
                // stack_blocks is used for table cells and header/footer —
                // no adjacent table collapse in these contexts.
                let table = layout_table(
                    rows,
                    col_widths,
                    *cell_spacing,
                    default_line_height,
                    border_config.as_ref(),
                    measure_text,
                    false,
                );

                let table_x = table_x_offset(
                    *alignment,
                    *indent,
                    table.size.width,
                    content_width,
                    Pt::ZERO,
                );

                for mut cmd in table.commands {
                    cmd.shift_y(cursor_y);
                    cmd.shift_x(table_x);
                    commands.push(cmd);
                }

                cursor_y += table.size.height;
                prev_space_after = Pt::ZERO;
                prev_style_id = None;
            }
        }
    }

    StackResult {
        commands,
        height: cursor_y,
        lines: if splittable { cell_lines } else { Vec::new() },
    }
}

#[cfg(test)]
mod tests {
    use super::super::types::FloatingImageY;
    use super::*;
    use crate::model::ImageFormat;
    use crate::model::WrapText;
    use crate::render::fonts::Toggle;
    use crate::render::geometry::PtSize;
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::{FloatingImage, FloatingImageX, WrapMode};
    use crate::render::resolve::images::MediaEntry;
    use std::rc::Rc;

    const LINE: f32 = 14.0;

    fn image(height: f32, wrap: WrapMode, y: FloatingImageY) -> FloatingImage {
        FloatingImage {
            image_data: MediaEntry {
                data: std::sync::Arc::from(&b""[..]),
                format: ImageFormat::Png,
            },
            size: PtSize::new(Pt::new(50.0), Pt::new(height)),
            src_rect: None,
            x: FloatingImageX::Absolute(Pt::ZERO),
            y,
            wrap_mode: wrap,
            dist_left: Pt::ZERO,
            dist_right: Pt::ZERO,
            behind_doc: false,
        }
    }

    fn top_bottom(height: f32) -> FloatingImage {
        image(
            height,
            WrapMode::TopAndBottom,
            FloatingImageY::RelativeToParagraph(Pt::ZERO),
        )
    }

    /// A one-word text fragment, so the paragraph produces a fitted line.
    fn text_fragment() -> crate::render::layout::fragment::Fragment {
        use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: Rc::from("word"),
            break_after: crate::render::layout::fragment::fixture_break_after("word"),
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
            }),
            color: crate::render::resolve::color::RgbColor::BLACK,
            width: Pt::new(40.0),
            trimmed_width: Pt::new(40.0),
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

    fn text_para(style: ParagraphStyle) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_fragment()],
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    fn para(style: ParagraphStyle, images: Vec<FloatingImage>) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![],
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: images,
            floating_shapes: vec![],
        }
    }

    fn styled(space_before: f32, space_after: f32, style_id: Option<&str>) -> ParagraphStyle {
        ParagraphStyle {
            space_before: Pt::new(space_before),
            space_after: Pt::new(space_after),
            style_id: style_id.map(crate::model::StyleId::new),
            ..Default::default()
        }
    }

    fn stack(blocks: &[LayoutBlock]) -> StackResult {
        stack_blocks(blocks, Pt::new(400.0), Pt::new(LINE), None, PageParity::Odd)
    }

    fn image_ys(result: &StackResult) -> Vec<f32> {
        result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Image { rect, .. } => Some(rect.origin.y.raw()),
                _ => None,
            })
            .collect()
    }

    fn text_ys(result: &StackResult) -> Vec<f32> {
        result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some(position.y.raw()),
                _ => None,
            })
            .collect()
    }

    // ── §20.4.2.18 wrapTopAndBottom ──────────────────────────────────────

    /// Regression: the float anchor was computed once before the loop, so every
    /// `wrapTopAndBottom` float on a paragraph resolved to the same y — two of
    /// them drew on top of each other and the cursor advanced only to the
    /// tallest. They occupy a band and must stack.
    #[test]
    fn multiple_top_and_bottom_floats_stack_instead_of_overlapping() {
        let result = stack(&[para(
            ParagraphStyle::default(),
            vec![top_bottom(30.0), top_bottom(40.0)],
        )]);
        assert_eq!(
            image_ys(&result),
            vec![0.0, 30.0],
            "the second float sits below the first"
        );
        assert!(
            (result.height.raw() - (70.0 + LINE)).abs() < 1e-4,
            "the cursor clears the whole band (30+40) plus the paragraph line, got {}",
            result.height.raw()
        );
    }

    /// A single float is positioned exactly as before the stacking fix.
    #[test]
    fn a_single_top_and_bottom_float_is_unchanged() {
        let result = stack(&[para(ParagraphStyle::default(), vec![top_bottom(30.0)])]);
        assert_eq!(image_ys(&result), vec![0.0]);
        assert!((result.height.raw() - (30.0 + LINE)).abs() < 1e-4);
    }

    /// Page-absolute floats are placed by their anchor, not by the band.
    #[test]
    fn absolute_top_and_bottom_floats_are_not_stacked() {
        let result = stack(&[para(
            ParagraphStyle::default(),
            vec![
                image(
                    20.0,
                    WrapMode::TopAndBottom,
                    FloatingImageY::Absolute(Pt::new(100.0)),
                ),
                image(
                    20.0,
                    WrapMode::TopAndBottom,
                    FloatingImageY::Absolute(Pt::new(200.0)),
                ),
            ],
        )]);
        assert_eq!(image_ys(&result), vec![100.0, 200.0], "anchors are honored");
    }

    /// Wrap-mode floats don't occupy a band above the paragraph — they overlay
    /// it and the text narrows around them — but the stacked height still grows
    /// to contain the image, so a table cell expands rather than clipping it.
    #[test]
    fn wrapping_floats_extend_the_height_to_contain_the_image() {
        let result = stack(&[para(
            ParagraphStyle::default(),
            vec![image(
                60.0,
                WrapMode::Square(WrapText::BothSides),
                FloatingImageY::RelativeToParagraph(Pt::ZERO),
            )],
        )]);
        assert!(
            (result.height.raw() - 60.0).abs() < 1e-4,
            "height reaches the image bottom, got {}",
            result.height.raw()
        );
        assert_eq!(image_ys(&result), vec![0.0], "anchored at the content top");
    }

    // ── §17.3.1.9 spacing collapse ───────────────────────────────────────

    /// Adjacent paragraphs collapse to `min(prev_after, before)`, not the sum.
    #[test]
    fn adjacent_paragraph_spacing_collapses_to_the_minimum() {
        let blocks = vec![
            para(styled(0.0, 10.0, None), vec![]),
            para(styled(6.0, 0.0, None), vec![]),
        ];
        // Two lines + both paragraphs' spacing, less the min(10, 6) = 6 overlap.
        let expected = LINE + 10.0 + 6.0 + LINE - 6.0;
        assert!(
            (stack(&blocks).height.raw() - expected).abs() < 1e-4,
            "expected {expected}, got {}",
            stack(&blocks).height.raw()
        );
    }

    /// §17.3.1.9: with `contextualSpacing` and a matching style id, the whole
    /// `prev_after + before` gap is removed rather than collapsed.
    #[test]
    fn contextual_spacing_removes_the_gap_between_same_styled_paragraphs() {
        let mut first = styled(0.0, 10.0, Some("ListParagraph"));
        first.contextual_spacing = true;
        let mut second = styled(6.0, 0.0, Some("ListParagraph"));
        second.contextual_spacing = true;

        let blocks = vec![para(first, vec![]), para(second, vec![])];
        let expected = LINE + 10.0 + 6.0 + LINE - (10.0 + 6.0);
        assert!(
            (stack(&blocks).height.raw() - expected).abs() < 1e-4,
            "expected {expected}, got {}",
            stack(&blocks).height.raw()
        );
    }

    /// A differing style id makes `contextualSpacing` inapplicable, so the
    /// ordinary collapse applies instead.
    #[test]
    fn contextual_spacing_needs_a_matching_style_id() {
        let mut first = styled(0.0, 10.0, Some("ListParagraph"));
        first.contextual_spacing = true;
        let mut second = styled(6.0, 0.0, Some("Other"));
        second.contextual_spacing = true;

        let blocks = vec![para(first, vec![]), para(second, vec![])];
        let expected = LINE + 10.0 + 6.0 + LINE - 6.0;
        assert!(
            (stack(&blocks).height.raw() - expected).abs() < 1e-4,
            "ordinary collapse, expected {expected}, got {}",
            stack(&blocks).height.raw()
        );
    }

    // ── §17.4.1 CellLine cut model ───────────────────────────────────────

    /// One entry per fitted line, tagged with its owning block index.
    #[test]
    fn cell_lines_are_recorded_per_paragraph() {
        let result = stack(&[
            text_para(ParagraphStyle::default()),
            text_para(ParagraphStyle::default()),
        ]);
        assert_eq!(result.lines.len(), 2, "one fitted line per paragraph");
        assert_eq!(
            result.lines.iter().map(|l| l.para).collect::<Vec<_>>(),
            vec![0, 1],
            "tagged with the owning block index"
        );
    }

    /// A floating object anchored anywhere in the content makes the whole cell
    /// unsafe to bisect — the cut model is withheld so the row moves whole.
    #[test]
    fn a_floating_object_withholds_the_cut_model() {
        let result = stack(&[
            text_para(ParagraphStyle::default()),
            para(ParagraphStyle::default(), vec![top_bottom(20.0)]),
        ]);
        assert!(
            result.lines.is_empty(),
            "lines are withheld even though the first paragraph was safe"
        );
    }

    /// §17.3.1.14 / §17.3.1.24: keepLines and bordered paragraphs may move
    /// whole but never be divided internally.
    #[test]
    fn keep_lines_and_borders_mark_a_paragraph_interior_atomic() {
        let keep = ParagraphStyle {
            keep_lines: true,
            ..Default::default()
        };
        let result = stack(&[text_para(ParagraphStyle::default()), text_para(keep)]);
        assert_eq!(
            result
                .lines
                .iter()
                .map(|l| l.interior_atomic)
                .collect::<Vec<_>>(),
            vec![false, true]
        );
    }

    // ── §20.4.2.15 wrapNone ──────────────────────────────────────────────

    fn text_xs(result: &StackResult) -> Vec<f32> {
        result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some(position.x.raw()),
                _ => None,
            })
            .collect()
    }

    fn wide_para(images: Vec<FloatingImage>) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_fragment(), text_fragment(), text_fragment()],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: images,
            floating_shapes: vec![],
        }
    }

    fn sized_image(width: f32, wrap: WrapMode) -> FloatingImage {
        let mut i = image(60.0, wrap, FloatingImageY::RelativeToParagraph(Pt::ZERO));
        i.size = crate::render::geometry::PtSize::new(Pt::new(width), Pt::new(60.0));
        i
    }

    /// Regression: a `wrapNone` image was registered as an active wrap float,
    /// so text narrowed around a drawing that §20.4.2.15 says "shall not cause
    /// text to wrap — it shall be displayed either in front of or behind the
    /// text". A 200pt-wide float pushed the text from x=0 to x=200.
    #[test]
    fn wrap_none_image_does_not_reflow_text() {
        let baseline = stack(&[wide_para(vec![])]);
        let overlaid = stack(&[wide_para(vec![sized_image(200.0, WrapMode::None)])]);
        assert_eq!(
            text_xs(&overlaid),
            text_xs(&baseline),
            "a wrapNone image overlays the text and must not move it"
        );
    }

    /// ...but it is still *drawn*, and still expands the stacked height so a
    /// table cell contains it. Only the wrap registration was wrong.
    #[test]
    fn wrap_none_image_is_still_emitted() {
        let result = stack(&[wide_para(vec![sized_image(200.0, WrapMode::None)])]);
        assert_eq!(image_ys(&result), vec![0.0], "the image is painted");
        assert!(
            (result.height.raw() - 60.0).abs() < 1e-4,
            "height still reaches the image bottom, got {}",
            result.height.raw()
        );
    }

    /// The wrap-enabled modes are unaffected — they still narrow the text.
    #[test]
    fn wrapping_modes_still_reflow_text() {
        let baseline = stack(&[wide_para(vec![])]);
        for wrap in [
            WrapMode::Square(WrapText::BothSides),
            WrapMode::Tight(WrapText::BothSides),
            WrapMode::Through(WrapText::BothSides),
        ] {
            let wrapped = stack(&[wide_para(vec![sized_image(200.0, wrap)])]);
            assert_ne!(
                text_xs(&wrapped),
                text_xs(&baseline),
                "{wrap:?} must still narrow the line"
            );
        }
    }

    /// §20.4.2.18: the paragraph sits **directly** below the band.
    ///
    /// The band starts at `content_top` (`cursor_y + space_before`), so the
    /// spacing is already spent by the time the float is placed;
    /// `place_paragraph` then applies `space_before` again inside the paragraph
    /// box. Text landed a second `space_before` below the band — 60.0 for a
    /// 10pt space and a 30pt float, where the band ends at 40 and the baseline
    /// belongs at 50.
    #[test]
    fn space_before_is_not_applied_twice_around_a_top_and_bottom_band() {
        const SPACE: f32 = 10.0;
        const FLOAT_H: f32 = 30.0;
        const ASCENT: f32 = 10.0;

        let block = LayoutBlock::Paragraph {
            fragments: vec![text_fragment()],
            style: styled(SPACE, 0.0, None),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![top_bottom(FLOAT_H)],
            floating_shapes: vec![],
        };
        let result = stack(&[block]);

        assert_eq!(
            image_ys(&result),
            vec![SPACE],
            "the band still starts after space_before"
        );
        assert_eq!(
            text_ys(&result),
            vec![SPACE + FLOAT_H + ASCENT],
            "the first line sits directly below the band, not a second space below it"
        );
    }

    /// The control: with no band, `space_before` is applied exactly once, so
    /// the fix must not have removed it outright.
    #[test]
    fn space_before_still_applies_without_a_band() {
        const SPACE: f32 = 10.0;
        const ASCENT: f32 = 10.0;
        let result = stack(&[text_para(styled(SPACE, 0.0, None))]);
        assert_eq!(text_ys(&result), vec![SPACE + ASCENT]);
    }

    /// A non-`TopAndBottom` float leaves the cursor alone, so it must not
    /// trigger the correction either.
    #[test]
    fn a_square_wrap_float_does_not_change_paragraph_spacing() {
        const SPACE: f32 = 10.0;
        const ASCENT: f32 = 10.0;
        let block = LayoutBlock::Paragraph {
            fragments: vec![text_fragment()],
            style: styled(SPACE, 0.0, None),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![image(
                30.0,
                WrapMode::Square(WrapText::BothSides),
                FloatingImageY::RelativeToParagraph(Pt::ZERO),
            )],
            floating_shapes: vec![],
        };
        let result = stack(&[block]);
        assert_eq!(text_ys(&result), vec![SPACE + ASCENT]);
    }

    /// Two stacked floats consume the space once between them, not once per
    /// float — the correction is applied after the whole band, not per float.
    #[test]
    fn stacked_band_consumes_space_before_only_once() {
        const SPACE: f32 = 10.0;
        const ASCENT: f32 = 10.0;
        let block = LayoutBlock::Paragraph {
            fragments: vec![text_fragment()],
            style: styled(SPACE, 0.0, None),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![top_bottom(30.0), top_bottom(40.0)],
            floating_shapes: vec![],
        };
        let result = stack(&[block]);
        assert_eq!(image_ys(&result), vec![SPACE, SPACE + 30.0]);
        assert_eq!(
            text_ys(&result),
            vec![SPACE + 70.0 + ASCENT],
            "text sits below the 70pt band, with space_before counted once"
        );
    }
}
