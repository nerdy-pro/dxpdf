use crate::model::*;
use crate::units::*;

use super::fragment::*;
use super::{ActiveFloat, DrawCommand, Layouter};

impl Layouter {
    pub(super) fn layout_paragraph(&mut self, para: &Paragraph) {
        let spacing = self.resolve_spacing(para.properties.spacing);
        let indent = para.properties.indentation.unwrap_or_default();

        self.cursor_y += spacing.before_pt();

        // Register floating images attached to this paragraph
        for float in &para.floats {
            if float.data.is_empty() {
                continue;
            }
            let content_w = self.config.content_width();
            let img_x = match float.align_h.as_deref() {
                Some("right") => self.config.margin_left + content_w - float.width_pt,
                Some("center") => self.config.margin_left + (content_w - float.width_pt) / 2.0,
                Some("left") => self.config.margin_left,
                _ => self.config.margin_left + float.offset_x_pt,
            };
            let img_y = self.cursor_y + float.offset_y_pt;

            self.current_page.commands.push(DrawCommand::Image {
                x: img_x,
                y: img_y,
                width: float.width_pt,
                height: float.height_pt,
                data: float.data.clone(),
            });

            self.active_floats.push(ActiveFloat {
                page_x: img_x,
                page_y_start: img_y,
                page_y_end: img_y + float.height_pt,
                width: float.width_pt,
            });
        }

        if para.runs.is_empty() && para.floats.is_empty() {
            self.cursor_y += spacing.line_pt();
            self.cursor_y += spacing.after_pt();
            return;
        }

        if para.runs.is_empty() {
            self.cursor_y += spacing.after_pt();
            return;
        }

        let base_content_width = self.config.content_width()
            - indent.left_pt()
            - indent.right_pt();
        let content_height = self.config.content_height();
        let fragments = collect_fragments(
            &para.runs,
            base_content_width,
            content_height,
            &self.doc_defaults,
            &self.measurer,
        );
        let base_x = self.config.margin_left + indent.left_pt();
        let mut para_cmd_start = self.current_page.commands.len();
        // Track the text area (excluding before/after spacing) for paragraph shading
        let mut text_area_top: Option<f32> = None;
        let mut text_area_bottom: f32 = self.cursor_y;

        let mut line_start = 0;
        let mut first_line = true;

        while line_start < fragments.len() {
            // Skip leading space fragments at the start of a new line
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

            let first_line_offset = if first_line {
                indent.first_line_pt()
            } else {
                0.0
            };

            let tentative_frag_height = fragments[line_start..]
                .iter()
                .take(1)
                .map(|f| f.height())
                .fold(0.0_f32, f32::max);
            let tentative_line_height = match spacing.line_pt_opt() {
                Some(lh) => tentative_frag_height.max(lh),
                None => tentative_frag_height,
            };
            let line_top = self.cursor_y;
            let line_bottom = self.cursor_y + tentative_line_height;
            let (float_x_shift, float_width_reduction) =
                self.float_adjustment(line_top, line_bottom);

            let available_width =
                base_content_width - first_line_offset - float_width_reduction;

            let (line_end, _) =
                fit_fragments(&fragments[line_start..], available_width.max(0.0));
            let actual_end = line_start + line_end.max(1);

            let frag_height = fragments[line_start..actual_end]
                .iter()
                .map(|f| f.height())
                .fold(0.0_f32, f32::max);
            let line_height = match spacing.line_pt_opt() {
                Some(lh) => frag_height.max(lh),
                None => frag_height,
            };

            if self.cursor_y + line_height > self.content_bottom() {
                self.new_page();
                // Reset shading tracking for the new page
                para_cmd_start = self.current_page.commands.len();
                text_area_top = None;
            }

            let used_width = if actual_end > line_start {
                measure_fragments(&fragments[line_start..actual_end])
            } else {
                0.0
            };
            let x_offset = match para.properties.alignment {
                Some(Alignment::Center) => (available_width - used_width) / 2.0,
                Some(Alignment::Right) => available_width - used_width,
                _ => 0.0,
            };

            let mut x = base_x + first_line_offset + float_x_shift + x_offset;
            // Track text area for paragraph shading
            if text_area_top.is_none() {
                text_area_top = Some(self.cursor_y);
            }
            self.cursor_y += line_height;
            text_area_bottom = self.cursor_y;

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
                        measured_width,
                        ..
                    } => {
                        let c = color.map(|c| (c.r, c.g, c.b)).unwrap_or((0, 0, 0));

                        if let Some(bg) = shading {
                            self.current_page.commands.push(DrawCommand::Rect {
                                x,
                                y: self.cursor_y - line_height,
                                width: *measured_width,
                                height: line_height,
                                color: (bg.r, bg.g, bg.b),
                            });
                        }

                        self.current_page.commands.push(DrawCommand::Text {
                            x,
                            y: self.cursor_y,
                            text: text.clone(),
                            font_family: font_family.clone(),
                            font_size: *font_size,
                            bold: *bold,
                            italic: *italic,
                            color: c,
                        });

                        if *underline {
                            self.current_page.commands.push(DrawCommand::Underline {
                                x1: x,
                                y1: self.cursor_y + UNDERLINE_Y_OFFSET,
                                x2: x + measured_width,
                                y2: self.cursor_y + UNDERLINE_Y_OFFSET,
                                color: c,
                                width: UNDERLINE_STROKE_WIDTH,
                            });
                        }

                        x += measured_width;
                    }
                    Fragment::Image { width, height, data } => {
                        self.current_page.commands.push(DrawCommand::Image {
                            x,
                            y: self.cursor_y - height,
                            width: *width,
                            height: *height,
                            data: data.clone(),
                        });
                        x += width;
                    }
                    Fragment::Tab { .. } => {
                        let rel_x = x - base_x;
                        let next_stop = find_next_tab_stop(
                            rel_x,
                            &para.properties.tab_stops,
                            self.default_tab_stop_pt,
                        );
                        x = base_x + next_stop;
                    }
                    Fragment::LineBreak { .. } => {}
                }
            }

            line_start = actual_end;
            first_line = false;
        }

        // Paragraph background shading — covers the text area only,
        // not before/after spacing
        if let Some(bg) = &para.properties.shading {
            if let Some(top) = text_area_top {
                let height = text_area_bottom - top;
                if height > 0.0 {
                    self.current_page.commands.insert(
                        para_cmd_start,
                        DrawCommand::Rect {
                            x: self.config.margin_left,
                            y: top,
                            width: self.config.content_width(),
                            height,
                            color: (bg.r, bg.g, bg.b),
                        },
                    );
                }
            }
        }

        self.cursor_y += spacing.after_pt();
    }
}
