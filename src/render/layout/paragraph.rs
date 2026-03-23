use std::rc::Rc;

use super::context::LayoutConstraints;
use super::fragment::*;
use super::measure::MeasuredParagraph;
use super::{offset_command, ActiveFloat, DrawCommand, Layouter};
use crate::dimension::Pt;
use crate::geometry::{PtLineSegment, PtOffset, PtRect};
use crate::model::*;

impl Layouter<'_> {
    pub(super) fn layout_paragraph(&mut self, para: &MeasuredParagraph) {
        let spacing = self.resolve_spacing(para.properties.spacing);
        let mut indent = para.properties.indentation.unwrap_or_default();

        // Use pre-resolved list label from the measure step
        let list_label = para.list_label.clone();

        // Capture page-level constraints (for floats, borders, shading)
        let page_ctx = *self.context.current();

        self.cursor_y += spacing.before.map(Pt::from).unwrap_or(Pt::ZERO);

        // Register floating images attached to this paragraph
        for float in &para.floats {
            if !self.image_cache.contains(&float.rel_id) {
                continue;
            }
            let content_w = page_ctx.available_width;
            let fw = float.size.width;
            let fh = float.size.height;
            let img_x = if let Some(pct) = float.pct_pos_h {
                page_ctx.page_size.width * (pct as f32 / 100_000.0)
            } else {
                match float.align_h.as_deref() {
                    Some("right") => page_ctx.x_origin + content_w - fw,
                    Some("center") => page_ctx.x_origin + (content_w - fw) / 2.0,
                    Some("left") => page_ctx.x_origin,
                    _ => page_ctx.x_origin + float.offset.x,
                }
            };
            let img_y = if let Some(pct) = float.pct_pos_v {
                page_ctx.page_size.height * (pct as f32 / 100_000.0)
            } else {
                self.cursor_y + float.offset.y
            };

            let image = self.image_cache.get(&float.rel_id);
            self.current_page.commands.push(DrawCommand::Image {
                rect: PtRect::from_xywh(img_x, img_y, fw, fh),
                image,
            });

            self.active_floats.push(ActiveFloat {
                page_x: img_x,
                page_y_start: img_y,
                page_y_end: img_y + fh,
                width: fw,
            });
        }

        if para.fragments.is_empty() && para.floats.is_empty() {
            let top = self.cursor_y;
            let natural_height = self.doc_defaults.default_line_height;
            let line_h = resolve_line_height(natural_height, spacing.line_spacing());
            // Paint borders before advancing — top border sits at paragraph start
            self.paint_paragraph_borders(
                &page_ctx,
                &para.properties.paragraph_borders,
                top,
                top + line_h,
            );
            self.cursor_y += line_h;
            self.cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
            return;
        }

        if para.fragments.is_empty() {
            self.prev_para_had_bottom_border = false;
            self.cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
            return;
        }

        // Apply list indentation (overrides paragraph indentation)
        let (list_label_text, list_label_x) = if let Some((ref label, left, hanging)) = list_label {
            if indent.left.is_none() {
                indent.left = Some(left);
            }
            if indent.first_line.is_none() {
                indent.first_line = Some(-hanging);
            }
            let left_pt = Pt::from(left);
            let hanging_pt = Pt::from(hanging);
            let label_x = page_ctx.x_origin + left_pt - hanging_pt;
            (Some(label.clone()), Some(label_x))
        } else {
            (None, None)
        };

        // Push paragraph-level constraints (narrowed by indentation)
        let para_ctx = page_ctx.for_paragraph(&indent);
        self.context.push(para_ctx);
        let fragments = &para.fragments;

        // ============================
        // PASS 1: MEASURE — produce lines with relative y-coordinates
        // ============================

        // Snapshot values needed by the float adjuster closure so we don't
        // borrow `self` while also passing other `self` fields to `measure_lines`.
        let abs_cursor_y = self.cursor_y;
        let default_tab = self.default_tab_stop_pt;
        let x_origin_page = page_ctx.x_origin;
        let float_data: Vec<(Pt, Pt, Pt, Pt)> = self
            .active_floats
            .iter()
            .map(|f| (f.page_x, f.page_y_start, f.page_y_end, f.width))
            .collect();

        let float_adj = |rel_line_top: Pt, rel_line_bottom: Pt| -> (Pt, Pt) {
            let abs_top = abs_cursor_y + rel_line_top;
            let abs_bottom = abs_cursor_y + rel_line_bottom;
            let gap = super::FLOAT_TEXT_GAP;
            let mut x_shift = Pt::ZERO;
            let mut width_reduction = Pt::ZERO;
            for &(page_x, page_y_start, page_y_end, width) in &float_data {
                if abs_top < page_y_end && abs_bottom > page_y_start {
                    let shift = (page_x - x_origin_page) + width + gap;
                    x_shift = x_shift.max(shift);
                    width_reduction = width_reduction.max(shift);
                }
            }
            (x_shift, width_reduction)
        };

        let ctx = self.context.current();
        let measured = measure_lines(
            fragments,
            ctx.x_origin,
            ctx.available_width,
            indent.first_line.map(Pt::from).unwrap_or(Pt::ZERO),
            para.properties.alignment,
            spacing.line_spacing(),
            &para.properties.tab_stops,
            default_tab,
            self.image_cache,
            if self.active_floats.is_empty() {
                None
            } else {
                Some(&float_adj)
            },
        );

        // ============================
        // PASS 2+3: LAYOUT & PAINT — assign to pages, emit commands
        // ============================
        // These passes are combined because page breaks require emitting
        // shading on the correct page before switching.

        let mut text_area_top: Option<Pt> = None;
        let mut text_area_bottom = self.cursor_y;
        let mut shading_insert_idx = self.current_page.commands.len();
        let mut first_line_painted = false;

        for (i, line) in measured.lines.iter().enumerate() {
            // Page break check
            if self.cursor_y + line.height > self.content_bottom() {
                // Paint paragraph shading for the current page before breaking
                self.paint_paragraph_shading(
                    &page_ctx,
                    &para.properties.shading,
                    text_area_top,
                    text_area_bottom,
                    shading_insert_idx,
                );

                self.new_page();

                // Reset shading tracking for new page
                text_area_top = None;
                shading_insert_idx = self.current_page.commands.len();
            }

            // List label on the first line
            if !first_line_painted {
                if let (Some(ref label), Some(lx)) = (&list_label_text, list_label_x) {
                    let font = &self.doc_defaults.font_family;
                    let fs = Pt::from(self.doc_defaults.font_size);
                    self.current_page.commands.push(DrawCommand::Text {
                        position: PtOffset::new(lx, self.cursor_y),
                        text: label.clone(),
                        font_family: Rc::clone(font),
                        char_spacing_pt: Pt::ZERO,
                        font_size: fs,
                        bold: false,
                        italic: false,
                        color: Color::BLACK,
                    });
                }
                first_line_painted = true;
            }

            // Track text area for shading
            if text_area_top.is_none() {
                text_area_top = Some(self.cursor_y);
            }

            // Paint line content at absolute position
            let rel_y_before: Pt = measured.lines[..i].iter().map(|l| l.height).sum();
            let y_offset = self.cursor_y - rel_y_before;

            for cmd in &line.commands {
                self.current_page
                    .commands
                    .push(offset_command(cmd, y_offset));
            }

            self.cursor_y += line.height;
            text_area_bottom = self.cursor_y;
        }

        // Paint paragraph shading for the last page
        self.paint_paragraph_shading(
            &page_ctx,
            &para.properties.shading,
            text_area_top,
            text_area_bottom,
            shading_insert_idx,
        );

        // Paint paragraph borders (uses page-level constraints for full width)
        if let Some(top) = text_area_top {
            self.paint_paragraph_borders(
                &page_ctx,
                &para.properties.paragraph_borders,
                top,
                text_area_bottom,
            );
        }

        // Pop paragraph constraints
        self.context.pop();

        self.cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
    }

    /// Paint paragraph border lines (w:pBdr).
    /// Word merges adjacent paragraph borders: when the previous paragraph had any
    /// border (top or bottom), the current paragraph's top border is suppressed to
    /// avoid duplicate lines.
    fn paint_paragraph_borders(
        &mut self,
        ctx: &LayoutConstraints,
        borders: &Option<ParagraphBorders>,
        top: Pt,
        bottom: Pt,
    ) {
        let has_any_border = borders.as_ref().is_some_and(|b| {
            b.top.as_ref().is_some_and(|d| d.is_visible())
                || b.bottom.as_ref().is_some_and(|d| d.is_visible())
        });

        if let Some(ref borders) = borders {
            let left = ctx.x_origin;
            let right = left + ctx.available_width;
            if let Some(ref b) = borders.top {
                // Skip top border if previous paragraph already drew a border
                if b.is_visible() && !self.prev_para_had_bottom_border {
                    let y = top - b.space;
                    self.current_page.commands.push(DrawCommand::Line {
                        line: PtLineSegment::new(PtOffset::new(left, y), PtOffset::new(right, y)),
                        color: b.color,
                        width: Pt::from(b.size),
                    });
                }
            }
            if let Some(ref b) = borders.bottom {
                if b.is_visible() {
                    let y = bottom + b.space;
                    self.current_page.commands.push(DrawCommand::Line {
                        line: PtLineSegment::new(PtOffset::new(left, y), PtOffset::new(right, y)),
                        color: b.color,
                        width: Pt::from(b.size),
                    });
                }
            }
            if let Some(ref b) = borders.left {
                if b.is_visible() {
                    let x = left - b.space;
                    self.current_page.commands.push(DrawCommand::Line {
                        line: PtLineSegment::new(PtOffset::new(x, top), PtOffset::new(x, bottom)),
                        color: b.color,
                        width: Pt::from(b.size),
                    });
                }
            }
            if let Some(ref b) = borders.right {
                if b.is_visible() {
                    let x = right + b.space;
                    self.current_page.commands.push(DrawCommand::Line {
                        line: PtLineSegment::new(PtOffset::new(x, top), PtOffset::new(x, bottom)),
                        color: b.color,
                        width: Pt::from(b.size),
                    });
                }
            }
        }

        self.prev_para_had_bottom_border = has_any_border;
    }

    /// Paint paragraph background shading covering the text area.
    fn paint_paragraph_shading(
        &mut self,
        ctx: &LayoutConstraints,
        shading: &Option<Color>,
        text_area_top: Option<Pt>,
        text_area_bottom: Pt,
        insert_idx: usize,
    ) {
        if let Some(bg) = shading {
            if let Some(top) = text_area_top {
                let height = text_area_bottom - top;
                if height > Pt::ZERO {
                    self.current_page.commands.insert(
                        insert_idx,
                        DrawCommand::Rect {
                            rect: PtRect::from_xywh(ctx.x_origin, top, ctx.available_width, height),
                            color: *bg,
                        },
                    );
                }
            }
        }
    }
}
