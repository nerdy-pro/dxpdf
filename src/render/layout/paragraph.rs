use std::rc::Rc;

use crate::dimension::Pt;
use crate::geometry::{PtLineSegment, PtOffset, PtRect};
use crate::model::*;
use crate::units::UNDERLINE_Y_OFFSET;

use super::fragment::*;
use super::{offset_command, ActiveFloat, DrawCommand, Layouter};

impl Layouter {
    pub(super) fn layout_paragraph(&mut self, para: &Paragraph) {
        let spacing = self.resolve_spacing(para.properties.spacing);
        let mut indent = para.properties.indentation.unwrap_or_default();

        // Resolve list label and override indentation
        let list_label = para
            .properties
            .list_ref
            .as_ref()
            .and_then(|lr| self.resolve_list_label(lr));

        self.cursor_y += spacing.before.map(Pt::from).unwrap_or(Pt::ZERO);

        // Register floating images attached to this paragraph
        for float in &para.floats {
            if !self.image_cache.contains(&float.rel_id) {
                continue;
            }
            let content_w = self.config.content_width();
            let fw = float.size.width;
            let fh = float.size.height;
            let img_x = if let Some(pct) = float.pct_pos_h {
                self.config.page_size.width * (pct as f32 / 100_000.0)
            } else {
                match float.align_h.as_deref() {
                    Some("right") => self.config.margins.left + content_w - fw,
                    Some("center") => self.config.margins.left + (content_w - fw) / 2.0,
                    Some("left") => self.config.margins.left,
                    _ => self.config.margins.left + float.offset.x,
                }
            };
            let img_y = if let Some(pct) = float.pct_pos_v {
                self.config.page_size.height * (pct as f32 / 100_000.0)
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

        if para.runs.is_empty() && para.floats.is_empty() {
            let top = self.cursor_y;
            let default_size = Pt::from(self.doc_defaults.font_size);
            let natural_height = self
                .measurer
                .font(&self.doc_defaults.font_family, default_size, false, false)
                .metrics()
                .line_height;
            let line_h = resolve_line_height(natural_height, spacing.line_spacing());
            // Paint borders before advancing — top border sits at paragraph start
            self.paint_paragraph_borders(&para.properties.paragraph_borders, top, top + line_h);
            self.cursor_y += line_h;
            self.cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
            return;
        }

        if para.runs.is_empty() {
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
            let label_x = self.config.margins.left + left_pt - hanging_pt;
            (Some(label.clone()), Some(label_x))
        } else {
            (None, None)
        };

        let base_content_width = self.config.content_width()
            - indent.left.map(Pt::from).unwrap_or(Pt::ZERO)
            - indent.right.map(Pt::from).unwrap_or(Pt::ZERO);
        let content_height = self.config.content_height();
        let fragments = collect_fragments(
            &para.runs,
            base_content_width,
            content_height,
            &self.doc_defaults,
            &self.measurer,
            &self.image_cache,
        );
        let base_x = self.config.margins.left + indent.left.map(Pt::from).unwrap_or(Pt::ZERO);

        // ============================
        // PASS 1: MEASURE — produce lines with relative y-coordinates
        // ============================
        let measured = self.measure_paragraph_lines(
            &fragments,
            base_x,
            base_content_width,
            indent.first_line.map(Pt::from).unwrap_or(Pt::ZERO),
            para.properties.alignment,
            spacing.line_spacing(),
            &para.properties.tab_stops,
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
            &para.properties.shading,
            text_area_top,
            text_area_bottom,
            shading_insert_idx,
        );

        // Paint paragraph borders
        if let Some(top) = text_area_top {
            self.paint_paragraph_borders(&para.properties.paragraph_borders, top, text_area_bottom);
        }

        self.cursor_y += spacing.after.map(Pt::from).unwrap_or(Pt::ZERO);
    }

    /// Paint paragraph border lines (w:pBdr).
    /// Word merges adjacent paragraph borders: when the previous paragraph had any
    /// border (top or bottom), the current paragraph's top border is suppressed to
    /// avoid duplicate lines.
    fn paint_paragraph_borders(&mut self, borders: &Option<ParagraphBorders>, top: Pt, bottom: Pt) {
        let has_any_border = borders.as_ref().is_some_and(|b| {
            b.top.as_ref().is_some_and(|d| d.is_visible())
                || b.bottom.as_ref().is_some_and(|d| d.is_visible())
        });

        if let Some(ref borders) = borders {
            let left = self.config.margins.left;
            let right = left + self.config.content_width();
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
                            rect: PtRect::from_xywh(
                                self.config.margins.left,
                                top,
                                self.config.content_width(),
                                height,
                            ),
                            color: *bg,
                        },
                    );
                }
            }
        }
    }

    /// MEASURE: Produce lines with relative y-coordinates from fragments.
    /// Handles float adjustment using current absolute cursor_y.
    fn measure_paragraph_lines(
        &mut self,
        fragments: &[Fragment],
        base_x: Pt,
        base_content_width: Pt,
        first_line_offset: Pt,
        alignment: Option<Alignment>,
        line_spacing: Option<LineSpacing>,
        tab_stops: &[TabStop],
    ) -> MeasuredLines {
        let mut lines = Vec::new();
        let mut rel_y = Pt::ZERO;
        let mut line_start = 0;
        let mut first_line = true;

        while line_start < fragments.len() {
            if !first_line {
                while line_start < fragments.len() {
                    if let Fragment::Text { ref text, .. } = fragments[line_start] {
                        if text.trim().is_empty() {
                            line_start += 1;
                            continue;
                        }
                    }
                    break;
                }
                if line_start >= fragments.len() {
                    break;
                }
            }

            let fl_offset = if first_line {
                first_line_offset
            } else {
                Pt::ZERO
            };

            // Float adjustment: use absolute position for overlap detection
            let abs_line_top = self.cursor_y + rel_y;
            let tentative_height = fragments[line_start..]
                .iter()
                .take(1)
                .map(|f| f.height())
                .fold(Pt::ZERO, Pt::max);
            let tentative_line_height = resolve_line_height(tentative_height, line_spacing);
            let (float_x_shift, float_width_reduction) =
                self.float_adjustment(abs_line_top, abs_line_top + tentative_line_height);

            let available_width =
                (base_content_width - fl_offset - float_width_reduction).max(Pt::ZERO);

            let (line_end, _) = fit_fragments(&fragments[line_start..], available_width);
            let actual_end = line_start + line_end.max(1);

            let frag_height = fragments[line_start..actual_end]
                .iter()
                .map(|f| f.height())
                .fold(Pt::ZERO, Pt::max);
            let line_height = resolve_line_height(frag_height, line_spacing);

            let used_width = if actual_end > line_start {
                measure_fragments(&fragments[line_start..actual_end])
            } else {
                Pt::ZERO
            };
            let x_offset = match alignment {
                Some(Alignment::Center) => (available_width - used_width) / 2.0,
                Some(Alignment::Right) => available_width - used_width,
                _ => Pt::ZERO,
            };

            rel_y += line_height;
            let mut commands = Vec::new();
            let mut x = base_x + fl_offset + float_x_shift + x_offset;

            for frag in &fragments[line_start..actual_end] {
                match frag {
                    Fragment::Text {
                        text,
                        font_family,
                        font_size,
                        bold,
                        italic,
                        underline,
                        color,
                        shading,
                        char_spacing_pt,
                        measured_width,
                        ascent,
                        hyperlink_url,
                        baseline_offset,
                        ..
                    } => {
                        let c = color.unwrap_or(Color::BLACK);
                        let line_top = rel_y - line_height;
                        let baseline_y = line_top + *ascent;
                        if let Some(bg) = shading {
                            commands.push(DrawCommand::Rect {
                                rect: PtRect::from_xywh(x, line_top, *measured_width, line_height),
                                color: *bg,
                            });
                        }
                        commands.push(DrawCommand::Text {
                            position: PtOffset::new(x, baseline_y + *baseline_offset),
                            text: text.clone(),
                            font_family: font_family.clone(),
                            char_spacing_pt: *char_spacing_pt,
                            font_size: *font_size,
                            bold: *bold,
                            italic: *italic,
                            color: c,
                        });
                        if *underline {
                            let uw = underline_width(*font_size, *bold);
                            let underline_y = baseline_y + UNDERLINE_Y_OFFSET;
                            commands.push(DrawCommand::Underline {
                                line: PtLineSegment::new(
                                    PtOffset::new(x, underline_y),
                                    PtOffset::new(x + *measured_width, underline_y),
                                ),
                                color: c,
                                width: uw,
                            });
                        }
                        if let Some(url) = hyperlink_url {
                            commands.push(DrawCommand::LinkAnnotation {
                                rect: PtRect::from_xywh(
                                    x,
                                    rel_y - line_height,
                                    *measured_width,
                                    line_height,
                                ),
                                url: url.clone(),
                            });
                        }
                        x += *measured_width;
                    }
                    Fragment::Image { size, rel_id } => {
                        let image = self.image_cache.get(rel_id);
                        commands.push(DrawCommand::Image {
                            rect: PtRect::from_xywh(
                                x,
                                rel_y - size.height,
                                size.width,
                                size.height,
                            ),
                            image,
                        });
                        x += size.width;
                    }
                    Fragment::Tab { .. } => {
                        let rel_x = x - base_x;
                        let next_stop =
                            find_next_tab_stop(rel_x, tab_stops, self.default_tab_stop_pt);
                        x = base_x + next_stop;
                    }
                    Fragment::LineBreak { .. } => {}
                }
            }

            lines.push(MeasuredLine {
                commands,
                height: line_height,
            });
            line_start = actual_end;
            first_line = false;
        }

        MeasuredLines {
            total_height: rel_y,
            lines,
        }
    }
}
