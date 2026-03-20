use skia_safe::{Font, FontMgr, FontStyle};

use crate::model::*;

/// Measures text using Skia font metrics.
pub struct TextMeasurer {
    font_mgr: FontMgr,
}

impl TextMeasurer {
    pub fn new() -> Self {
        Self {
            font_mgr: FontMgr::new(),
        }
    }

    /// Get a Skia Font for the given properties.
    fn make_font(&self, font_family: &str, font_size: f32, bold: bool, italic: bool) -> Font {
        let style = match (bold, italic) {
            (true, true) => FontStyle::bold_italic(),
            (true, false) => FontStyle::bold(),
            (false, true) => FontStyle::italic(),
            (false, false) => FontStyle::normal(),
        };

        let typeface = self
            .font_mgr
            .match_family_style(font_family, style)
            .or_else(|| self.font_mgr.match_family_style("Helvetica", style))
            .or_else(|| self.font_mgr.legacy_make_typeface(None::<&str>, style))
            .expect("no fallback typeface available");

        Font::from_typeface(typeface, font_size)
    }

    /// Measure the width of a text string in points.
    fn measure_width(&self, text: &str, font_family: &str, font_size: f32, bold: bool, italic: bool) -> f32 {
        let font = self.make_font(font_family, font_size, bold, italic);
        let (width, _) = font.measure_str(text, None);
        width
    }

    /// Get the line height (ascent + descent + leading) for a font.
    fn line_height(&self, font_family: &str, font_size: f32, bold: bool, italic: bool) -> f32 {
        let font = self.make_font(font_family, font_size, bold, italic);
        let (_, metrics) = font.metrics();
        // ascent is negative in Skia (upward), descent is positive (downward)
        -metrics.ascent + metrics.descent + metrics.leading
    }
}

/// Page layout configuration in points (1 point = 1/72 inch).
#[derive(Debug, Clone, Copy)]
pub struct LayoutConfig {
    pub page_width: f32,
    pub page_height: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
}

impl Default for LayoutConfig {
    fn default() -> Self {
        Self {
            page_width: 612.0,   // 8.5 inches
            page_height: 792.0,  // 11 inches
            margin_top: 72.0,    // 1 inch
            margin_bottom: 72.0,
            margin_left: 72.0,
            margin_right: 72.0,
        }
    }
}

impl LayoutConfig {
    pub fn content_width(&self) -> f32 {
        self.page_width - self.margin_left - self.margin_right
    }

    pub fn content_height(&self) -> f32 {
        self.page_height - self.margin_top - self.margin_bottom
    }
}

/// A positioned drawing command.
#[derive(Debug, Clone)]
pub enum DrawCommand {
    Text {
        x: f32,
        y: f32,
        text: String,
        font_family: String,
        font_size: f32,
        bold: bool,
        italic: bool,
        color: (u8, u8, u8),
    },
    Underline {
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: (u8, u8, u8),
        width: f32,
    },
    Line {
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: (u8, u8, u8),
        width: f32,
    },
    Image {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        data: Vec<u8>,
    },
}

#[derive(Debug, Clone)]
pub struct LayoutedPage {
    pub commands: Vec<DrawCommand>,
    pub page_width: f32,
    pub page_height: f32,
}

/// Perform layout on a document, producing positioned draw commands per page.
pub fn layout(doc: &Document, config: &LayoutConfig) -> Vec<LayoutedPage> {
    // Collect all section configs in document order.
    // In DOCX, sectPr on a paragraph describes the section it ENDS.
    // The sequence is: [sect1_props, sect2_props, ..., final_section].
    // sect1 applies to content before the first break,
    // sect2 applies to content between break 1 and break 2, etc.
    let mut section_configs: Vec<LayoutConfig> = Vec::new();
    for block in &doc.blocks {
        if let Block::Paragraph(p) = block {
            if let Some(ref sect) = p.section_properties {
                let mut cfg = *config;
                apply_section_to_config(&mut cfg, sect);
                section_configs.push(cfg);
            }
        }
    }
    if let Some(ref sect) = doc.final_section {
        let mut cfg = *config;
        apply_section_to_config(&mut cfg, sect);
        section_configs.push(cfg);
    }

    // The first config in the list applies to the initial pages.
    // If no sections found, use the default config.
    let initial_config = section_configs.first().copied().unwrap_or(*config);
    // Remaining configs are applied at each section break in order.
    let mut next_configs = section_configs.into_iter().skip(1).collect::<Vec<_>>();
    next_configs.reverse(); // So we can pop from the end

    let default_tab_stop_pt = doc.default_tab_stop as f32 / 20.0;
    let doc_defaults = DocDefaultsLayout {
        font_size_half_pts: doc.default_font_size,
        font_family: doc.default_font_family.clone(),
        default_spacing: doc.default_spacing,
    };
    let mut layouter = Layouter::new(&initial_config, next_configs, default_tab_stop_pt, doc_defaults);

    for block in &doc.blocks {
        layouter.layout_block(block);
    }

    layouter.finish()
}

fn apply_section_to_config(config: &mut LayoutConfig, sect: &SectionProperties) {
    if let Some(ps) = &sect.page_size {
        config.page_width = ps.width_pt();
        config.page_height = ps.height_pt();
    }
    if let Some(pm) = &sect.page_margins {
        config.margin_top = pm.top_pt();
        config.margin_right = pm.right_pt();
        config.margin_bottom = pm.bottom_pt();
        config.margin_left = pm.left_pt();
    }
}

/// Document-level defaults for layout.
struct DocDefaultsLayout {
    font_size_half_pts: u32,
    font_family: String,
    default_spacing: Spacing,
}

/// A floating image that affects text layout on the current page.
struct ActiveFloat {
    page_x: f32,
    page_y_start: f32,
    page_y_end: f32,
    width: f32,
}

struct Layouter {
    config: LayoutConfig,
    pages: Vec<LayoutedPage>,
    current_page: LayoutedPage,
    cursor_y: f32,
    active_floats: Vec<ActiveFloat>,
    /// Queue of section configs to apply at each section break (reversed, pop from end).
    next_section_configs: Vec<LayoutConfig>,
    /// Default tab stop interval in points.
    default_tab_stop_pt: f32,
    /// Document-level default font settings.
    doc_defaults: DocDefaultsLayout,
    /// Text measurer using Skia font metrics.
    measurer: TextMeasurer,
}

impl Layouter {
    fn new(
        config: &LayoutConfig,
        next_section_configs: Vec<LayoutConfig>,
        default_tab_stop_pt: f32,
        doc_defaults: DocDefaultsLayout,
    ) -> Self {
        let measurer = TextMeasurer::new();
        Self {
            config: *config,
            pages: Vec::new(),
            current_page: LayoutedPage {
                commands: Vec::new(),
                page_width: config.page_width,
                page_height: config.page_height,
            },
            cursor_y: config.margin_top,
            active_floats: Vec::new(),
            next_section_configs,
            default_tab_stop_pt,
            doc_defaults,
            measurer,
        }
    }

    fn content_bottom(&self) -> f32 {
        self.config.page_height - self.config.margin_bottom
    }

    fn new_page(&mut self) {
        let page = std::mem::replace(
            &mut self.current_page,
            LayoutedPage {
                commands: Vec::new(),
                page_width: self.config.page_width,
                page_height: self.config.page_height,
            },
        );
        self.pages.push(page);
        self.cursor_y = self.config.margin_top;
        self.active_floats.clear();
    }

    /// Handle a section break: finish current page, then switch to
    /// the next section's config for subsequent pages.
    fn section_break(&mut self) {
        self.new_page();
        // Pop the next section config from the queue
        if let Some(next_config) = self.next_section_configs.pop() {
            self.config = next_config;
            // Update current (new) page dimensions
            self.current_page.page_width = self.config.page_width;
            self.current_page.page_height = self.config.page_height;
            self.cursor_y = self.config.margin_top;
        }
    }

    /// Compute how much the text start position shifts right and width reduces
    /// due to active floats overlapping the given vertical range.
    fn float_adjustment(&self, line_top: f32, line_bottom: f32) -> (f32, f32) {
        let gap = 4.0;
        let mut x_shift = 0.0_f32;
        let mut width_reduction = 0.0_f32;
        for f in &self.active_floats {
            if line_top < f.page_y_end && line_bottom > f.page_y_start {
                // Float overlaps this line; assume left-anchored
                let shift = (f.page_x - self.config.margin_left) + f.width + gap;
                x_shift = x_shift.max(shift);
                width_reduction = width_reduction.max(shift);
            }
        }
        (x_shift, width_reduction)
    }

    /// Resolve paragraph spacing: use the paragraph's explicit spacing,
    /// falling back to document defaults for unset fields.
    fn resolve_spacing(&self, para_spacing: Option<Spacing>) -> Spacing {
        let defaults = &self.doc_defaults.default_spacing;
        match para_spacing {
            Some(s) => Spacing {
                before: s.before.or(defaults.before),
                after: s.after.or(defaults.after),
                line: s.line.or(defaults.line),
            },
            None => *defaults,
        }
    }

    /// Get the width of a table cell: use tcW if available, otherwise grid_cols.
    fn cell_width(&self, col_idx: usize, cell: &TableCell, col_widths: &[f32]) -> f32 {
        cell.width_pt()
            .unwrap_or_else(|| col_widths.get(col_idx).copied().unwrap_or(72.0))
    }

    /// Prune floats that the cursor has moved past.
    fn prune_floats(&mut self) {
        self.active_floats
            .retain(|f| self.cursor_y < f.page_y_end);
    }

    fn layout_block(&mut self, block: &Block) {
        self.prune_floats();
        match block {
            Block::Paragraph(p) => {
                self.layout_paragraph(p);
                if p.section_properties.is_some() {
                    self.section_break();
                }
            }
            Block::Table(t) => self.layout_table(t),
        }
    }

    fn layout_paragraph(&mut self, para: &Paragraph) {
        let spacing = self.resolve_spacing(para.properties.spacing);
        let indent = para.properties.indentation.unwrap_or_default();

        // Space before
        self.cursor_y += spacing.before_pt();

        // Register floating images attached to this paragraph
        for float in &para.floats {
            if float.data.is_empty() {
                continue;
            }
            let img_x = self.config.margin_left + float.offset_x_pt;
            let img_y = self.cursor_y + float.offset_y_pt;

            // Emit the image draw command
            self.current_page.commands.push(DrawCommand::Image {
                x: img_x,
                y: img_y,
                width: float.width_pt,
                height: float.height_pt,
                data: float.data.clone(),
            });

            // Register as active float for text wrapping
            self.active_floats.push(ActiveFloat {
                page_x: img_x,
                page_y_start: img_y,
                page_y_end: img_y + float.height_pt,
                width: float.width_pt,
            });
        }

        if para.runs.is_empty() && para.floats.is_empty() {
            // Empty paragraph - just add line spacing
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
        let fragments = collect_fragments(&para.runs, base_content_width, content_height, &self.doc_defaults, &self.measurer);
        let base_x = self.config.margin_left + indent.left_pt();

        let mut line_start = 0;
        let mut first_line = true;

        while line_start < fragments.len() {
            let first_line_offset = if first_line {
                indent.first_line_pt()
            } else {
                0.0
            };

            // Compute float adjustment for this line's vertical position
            let tentative_line_height = fragments[line_start..]
                .iter()
                .take(1)
                .map(|f| f.height())
                .fold(spacing.line_pt(), f32::max);
            let line_top = self.cursor_y;
            let line_bottom = self.cursor_y + tentative_line_height;
            let (float_x_shift, float_width_reduction) =
                self.float_adjustment(line_top, line_bottom);

            let available_width =
                base_content_width - first_line_offset - float_width_reduction;

            let (line_end, _line_width) =
                fit_fragments(&fragments[line_start..], available_width.max(0.0));
            let actual_end = line_start + line_end.max(1);

            let line_height = fragments[line_start..actual_end]
                .iter()
                .map(|f| f.height())
                .fold(spacing.line_pt(), f32::max);

            if self.cursor_y + line_height > self.content_bottom() {
                self.new_page();
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
            self.cursor_y += line_height;

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
                        measured_width,
                        ..
                    } => {
                        let c = color
                            .map(|c| (c.r, c.g, c.b))
                            .unwrap_or((0, 0, 0));

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
                                y1: self.cursor_y + 2.0,
                                x2: x + measured_width,
                                y2: self.cursor_y + 2.0,
                                color: c,
                                width: 0.5,
                            });
                        }

                        x += measured_width;
                    }
                    Fragment::Image { width, height, data } => {
                        // Draw image positioned so its bottom aligns with the text baseline
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
                    Fragment::LineBreak { .. } => {
                        // Consumed by fit_fragments; no draw command needed
                    }
                }
            }

            line_start = actual_end;
            first_line = false;
        }

        // Space after
        self.cursor_y += spacing.after_pt();
    }

    fn layout_table(&mut self, table: &Table) {
        if table.rows.is_empty() {
            return;
        }

        let num_cols = table.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0);
        if num_cols == 0 {
            return;
        }

        let content_width = self.config.content_width();
        let cell_padding = 4.0;

        // Compute column widths in points
        let col_widths: Vec<f32> = if !table.grid_cols.is_empty() {
            // Use grid column definitions, scaled to fit content width
            let grid_total: f32 = table.grid_cols.iter().map(|w| *w as f32 / 20.0).sum();
            let scale = if grid_total > 0.0 {
                content_width / grid_total
            } else {
                1.0
            };
            table.grid_cols.iter().map(|w| *w as f32 / 20.0 * scale).collect()
        } else {
            // Fall back to even distribution
            vec![content_width / num_cols as f32; num_cols]
        };

        for row in &table.rows {
            // First pass: lay out each cell's content to compute row height
            let mut cell_layouts: Vec<Vec<DrawCommand>> = Vec::new();
            let mut row_height = cell_padding * 2.0 + 12.0; // minimum row height

            // Compute x positions for each cell
            let mut col_x_positions: Vec<f32> = Vec::with_capacity(row.cells.len());
            let mut x_acc = self.config.margin_left;
            for i in 0..row.cells.len() {
                col_x_positions.push(x_acc);
                let w = self.cell_width(i, &row.cells[i], &col_widths);
                x_acc += w;
            }

            for (col_idx, cell) in row.cells.iter().enumerate() {
                let cell_x = col_x_positions[col_idx];
                let col_width = self.cell_width(col_idx, cell, &col_widths);
                let cell_content_width = (col_width - cell_padding * 2.0).max(1.0);
                let mut commands = Vec::new();
                let mut cell_y = cell_padding;

                for block in &cell.blocks {
                    if let Block::Paragraph(p) = block {
                        let spacing = self.resolve_spacing(p.properties.spacing);
                        cell_y += spacing.before_pt();

                        // Render floating images within the cell.
                        // In table cells, use offset_x if it fits, otherwise center.
                        for float in &p.floats {
                            if float.data.is_empty() {
                                continue;
                            }
                            let scale = f32::min(
                                1.0,
                                f32::min(
                                    cell_content_width / float.width_pt.max(1.0),
                                    self.config.content_height() / float.height_pt.max(1.0),
                                ),
                            );
                            let img_w = float.width_pt * scale;
                            let img_h = float.height_pt * scale;
                            // Use the horizontal offset if it places the image
                            // within the cell; otherwise center it.
                            let img_x = if float.offset_x_pt > 0.0
                                && float.offset_x_pt + img_w <= cell_content_width
                            {
                                cell_x + cell_padding + float.offset_x_pt
                            } else {
                                cell_x + (col_width - img_w) / 2.0
                            };
                            commands.push(DrawCommand::Image {
                                x: img_x,
                                y: cell_y,
                                width: img_w,
                                height: img_h,
                                data: float.data.clone(),
                            });
                            cell_y += img_h;
                        }

                        let fragments = collect_fragments(
                            &p.runs,
                            cell_content_width,
                            self.config.content_height(),
                            &self.doc_defaults,
                            &self.measurer,
                        );

                        if fragments.is_empty() && p.floats.is_empty() {
                            cell_y += spacing.line_pt();
                            cell_y += spacing.after_pt();
                            continue;
                        }

                        if fragments.is_empty() {
                            cell_y += spacing.after_pt();
                            continue;
                        }

                        let mut line_start = 0;
                        while line_start < fragments.len() {
                            let (line_end, _) = fit_fragments(
                                &fragments[line_start..],
                                cell_content_width,
                            );
                            let actual_end = line_start + line_end.max(1);

                            let line_height = fragments[line_start..actual_end]
                                .iter()
                                .map(|f| f.height())
                                .fold(spacing.line_pt(), f32::max);

                            cell_y += line_height;

                            let used_width = measure_fragments(
                                &fragments[line_start..actual_end],
                            );
                            let align_offset = match p.properties.alignment {
                                Some(Alignment::Center) => {
                                    (cell_content_width - used_width) / 2.0
                                }
                                Some(Alignment::Right) => {
                                    cell_content_width - used_width
                                }
                                _ => 0.0,
                            };
                            let mut x = cell_x + cell_padding + align_offset;

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
                                        measured_width,
                                        ..
                                    } => {
                                        let c = color
                                            .map(|c| (c.r, c.g, c.b))
                                            .unwrap_or((0, 0, 0));

                                        commands.push(DrawCommand::Text {
                                            x,
                                            y: cell_y,
                                            text: text.clone(),
                                            font_family: font_family.clone(),
                                            font_size: *font_size,
                                            bold: *bold,
                                            italic: *italic,
                                            color: c,
                                        });

                                        if *underline {
                                            commands.push(DrawCommand::Underline {
                                                x1: x,
                                                y1: cell_y + 2.0,
                                                x2: x + measured_width,
                                                y2: cell_y + 2.0,
                                                color: c,
                                                width: 0.5,
                                            });
                                        }

                                        x += measured_width;
                                    }
                                    Fragment::Image { width, height, data } => {
                                        commands.push(DrawCommand::Image {
                                            x,
                                            y: cell_y - height,
                                            width: *width,
                                            height: *height,
                                            data: data.clone(),
                                        });
                                        x += width;
                                    }
                                    Fragment::Tab { .. } => {
                                        let rel_x = x - (cell_x + cell_padding);
                                        let next_stop = find_next_tab_stop(
                                            rel_x,
                                            &p.properties.tab_stops,
                                            self.default_tab_stop_pt,
                                        );
                                        x = cell_x + cell_padding + next_stop;
                                    }
                                    Fragment::LineBreak { .. } => {}
                                }
                            }

                            line_start = actual_end;
                        }

                        cell_y += spacing.after_pt();
                    }
                }

                let total_cell_height = cell_y + cell_padding;
                if total_cell_height > row_height {
                    row_height = total_cell_height;
                }
                cell_layouts.push(commands);
            }

            // Page break if needed
            if self.cursor_y + row_height > self.content_bottom() {
                self.new_page();
            }

            let row_top = self.cursor_y;

            // Second pass: emit cell content with correct row_top y offset
            for (col_idx, commands) in cell_layouts.iter().enumerate() {
                let cell_x = col_x_positions[col_idx];
                let cw = self.cell_width(col_idx, &row.cells[col_idx], &col_widths);

                // Cell borders: top and left
                self.current_page.commands.push(DrawCommand::Line {
                    x1: cell_x,
                    y1: row_top,
                    x2: cell_x + cw,
                    y2: row_top,
                    color: (0, 0, 0),
                    width: 0.5,
                });
                self.current_page.commands.push(DrawCommand::Line {
                    x1: cell_x,
                    y1: row_top,
                    x2: cell_x,
                    y2: row_top + row_height,
                    color: (0, 0, 0),
                    width: 0.5,
                });

                // Right border for last column
                if col_idx == row.cells.len() - 1 {
                    self.current_page.commands.push(DrawCommand::Line {
                        x1: cell_x + cw,
                        y1: row_top,
                        x2: cell_x + cw,
                        y2: row_top + row_height,
                        color: (0, 0, 0),
                        width: 0.5,
                    });
                }

                // Emit content commands with y offset by row_top
                for cmd in commands {
                    let adjusted = match cmd {
                        DrawCommand::Text {
                            x, y, text, font_family, font_size,
                            bold, italic, color,
                        } => DrawCommand::Text {
                            x: *x,
                            y: row_top + y,
                            text: text.clone(),
                            font_family: font_family.clone(),
                            font_size: *font_size,
                            bold: *bold,
                            italic: *italic,
                            color: *color,
                        },
                        DrawCommand::Underline {
                            x1, y1, x2, y2, color, width,
                        } => DrawCommand::Underline {
                            x1: *x1,
                            y1: row_top + y1,
                            x2: *x2,
                            y2: row_top + y2,
                            color: *color,
                            width: *width,
                        },
                        DrawCommand::Image {
                            x, y, width, height, data,
                        } => DrawCommand::Image {
                            x: *x,
                            y: row_top + y,
                            width: *width,
                            height: *height,
                            data: data.clone(),
                        },
                        DrawCommand::Line { .. } => cmd.clone(),
                    };
                    self.current_page.commands.push(adjusted);
                }
            }

            self.cursor_y += row_height;

            // Bottom border
            let table_width: f32 = (0..row.cells.len())
                .map(|i| self.cell_width(i, &row.cells[i], &col_widths))
                .sum();
            self.current_page.commands.push(DrawCommand::Line {
                x1: self.config.margin_left,
                y1: self.cursor_y,
                x2: self.config.margin_left + table_width,
                y2: self.cursor_y,
                color: (0, 0, 0),
                width: 0.5,
            });
        }

        self.cursor_y += 8.0; // Space after table
    }

    fn finish(mut self) -> Vec<LayoutedPage> {
        if !self.current_page.commands.is_empty() {
            self.pages.push(self.current_page);
        }
        // Always return at least one page
        if self.pages.is_empty() {
            self.pages.push(LayoutedPage {
                commands: Vec::new(),
                page_width: self.config.page_width,
                page_height: self.config.page_height,
            });
        }
        self.pages
    }
}

/// A flattened fragment for layout — either text, an image, or a tab.
enum Fragment {
    Text {
        text: String,
        font_family: String,
        font_size: f32,
        bold: bool,
        italic: bool,
        underline: bool,
        color: Option<Color>,
        /// Measured width from Skia.
        measured_width: f32,
        /// Measured line height from Skia font metrics.
        measured_height: f32,
    },
    Image {
        width: f32,
        height: f32,
        data: Vec<u8>,
    },
    /// A tab character. Its actual width depends on the current x position
    /// and is resolved during layout, not during fragment collection.
    Tab {
        line_height: f32,
    },
    /// A forced line break (`<w:br/>`).
    LineBreak {
        line_height: f32,
    },
}

impl Fragment {
    fn width(&self) -> f32 {
        match self {
            Fragment::Text { measured_width, .. } => *measured_width,
            Fragment::Image { width, .. } => *width,
            Fragment::Tab { .. } => 12.0,
            Fragment::LineBreak { .. } => 0.0,
        }
    }

    fn height(&self) -> f32 {
        match self {
            Fragment::Text { measured_height, .. } => *measured_height,
            Fragment::Image { height, .. } => *height,
            Fragment::Tab { line_height } | Fragment::LineBreak { line_height } => *line_height,
        }
    }

    fn is_line_break(&self) -> bool {
        matches!(self, Fragment::LineBreak { .. })
    }
}

fn collect_fragments(
    runs: &[Inline],
    content_width: f32,
    content_height: f32,
    defaults: &DocDefaultsLayout,
    measurer: &TextMeasurer,
) -> Vec<Fragment> {
    let mut fragments = Vec::new();
    for run in runs {
        match run {
            Inline::TextRun(tr) => {
                let collapsed = collapse_spaces(&tr.text);
                let font_family = tr
                    .properties
                    .font_family
                    .clone()
                    .unwrap_or_else(|| defaults.font_family.clone());
                let font_size = tr.properties.font_size_pt_with_default(
                    defaults.font_size_half_pts,
                );
                let bold = tr.properties.bold;
                let italic = tr.properties.italic;
                let line_height = measurer.line_height(&font_family, font_size, bold, italic);
                // Split into alternating word and space fragments.
                // This allows the line-break logic to not count trailing
                // spaces toward line width.
                for part in split_words_and_spaces(&collapsed) {
                    let measured_width = measurer.measure_width(part, &font_family, font_size, bold, italic);
                    let is_space = part.chars().all(|c| c == ' ');
                    fragments.push(Fragment::Text {
                        text: part.to_string(),
                        font_family: font_family.clone(),
                        font_size,
                        bold,
                        italic,
                        underline: !is_space && tr.properties.underline,
                        color: tr.properties.color,
                        measured_width,
                        measured_height: line_height,
                    });
                }
            }
            Inline::Image(img) if !img.data.is_empty() => {
                let scale = f32::min(
                    1.0,
                    f32::min(
                        content_width / img.width_pt.max(1.0),
                        content_height / img.height_pt.max(1.0),
                    ),
                );
                fragments.push(Fragment::Image {
                    width: img.width_pt * scale,
                    height: img.height_pt * scale,
                    data: img.data.clone(),
                });
            }
            Inline::Tab => {
                let default_size = defaults.font_size_half_pts as f32 / 2.0;
                let lh = measurer.line_height(&defaults.font_family, default_size, false, false);
                fragments.push(Fragment::Tab { line_height: lh });
            }
            Inline::LineBreak => {
                let default_size = defaults.font_size_half_pts as f32 / 2.0;
                let lh = measurer.line_height(&defaults.font_family, default_size, false, false);
                fragments.push(Fragment::LineBreak { line_height: lh });
            }
            Inline::Image(_) => {}
        }
    }
    fragments
}

/// Find the next tab stop position (in points, relative to paragraph left edge)
/// given the current x position relative to paragraph left edge.
fn find_next_tab_stop(current_x: f32, custom_stops: &[TabStop], default_interval: f32) -> f32 {
    // First, check custom tab stops (sorted by position)
    for stop in custom_stops {
        let pos = stop.position_pt();
        if pos > current_x + 1.0 {
            return pos;
        }
    }
    // Fall back to default tab interval
    if default_interval > 0.0 {
        let next_multiple = ((current_x / default_interval).floor() + 1.0) * default_interval;
        return next_multiple;
    }
    // Absolute fallback
    current_x + 36.0
}

/// Split text into alternating word and space segments.
/// E.g., "RCD Messung" -> ["RCD", " ", "Messung"]
fn split_words_and_spaces(text: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0;
    let bytes = text.as_bytes();
    let mut in_space = bytes.first() == Some(&b' ');

    for (i, &b) in bytes.iter().enumerate() {
        let is_space = b == b' ';
        if is_space != in_space {
            if start < i {
                parts.push(&text[start..i]);
            }
            start = i;
            in_space = is_space;
        }
    }
    if start < text.len() {
        parts.push(&text[start..]);
    }
    parts
}

/// Collapse runs of more than 2 consecutive spaces into a single space.
/// Word documents often use long runs of spaces for manual alignment.
fn collapse_spaces(text: &str) -> String {
    let mut result = String::with_capacity(text.len());
    let mut space_count = 0;
    for ch in text.chars() {
        if ch == ' ' {
            space_count += 1;
            if space_count <= 2 {
                result.push(ch);
            }
        } else {
            space_count = 0;
            result.push(ch);
        }
    }
    result
}

fn measure_fragments(fragments: &[Fragment]) -> f32 {
    fragments.iter().map(|f| f.width()).sum()
}

/// Find how many fragments fit within the available width.
/// Returns (count, total_width).
/// Trailing space fragments at a line break are included in the count
/// but not in the returned width, so they don't affect alignment.
fn fit_fragments(fragments: &[Fragment], available_width: f32) -> (usize, f32) {
    let mut total_width = 0.0;
    let mut last_break_point = 0;
    let mut width_at_break = 0.0;

    for (i, frag) in fragments.iter().enumerate() {
        // Forced line break: consume it and stop
        if frag.is_line_break() {
            return (i + 1, total_width);
        }

        let w = frag.width();
        let is_space = matches!(frag, Fragment::Text { ref text, .. } if text.trim().is_empty());

        if total_width + w > available_width && i > 0 {
            if last_break_point > 0 {
                return (last_break_point, width_at_break);
            }
            return (i, total_width);
        }

        total_width += w;

        if is_space {
            last_break_point = i + 1;
            width_at_break = total_width - w;
        }
    }
    (fragments.len(), total_width)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_doc(blocks: Vec<Block>) -> Document {
        Document {
            blocks,
            final_section: None,
            default_tab_stop: 720,
            default_font_size: 24,
            default_font_family: "Helvetica".to_string(),
            default_spacing: Spacing::default(),
        }
    }

    fn simple_paragraph(text: &str) -> Block {
        Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![Inline::TextRun(TextRun {
                text: text.to_string(),
                properties: RunProperties::default(),
            })],
            floats: Vec::new(),
            section_properties: None,
        })
    }

    #[test]
    fn layout_empty_document() {
        let doc = make_doc(vec![]);
        let pages = layout(&doc, &LayoutConfig::default());
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn layout_single_paragraph() {
        let doc = make_doc(vec![simple_paragraph("Hello World")]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        assert_eq!(pages.len(), 1);
        assert!(!pages[0].commands.is_empty());
        // Should have at least one Text command
        assert!(pages[0].commands.iter().any(|c| matches!(c, DrawCommand::Text { .. })));
    }

    #[test]
    fn layout_page_break() {
        // Create enough paragraphs to force a page break
        let mut blocks = Vec::new();
        for i in 0..100 {
            blocks.push(Block::Paragraph(Paragraph {
                properties: ParagraphProperties {
                    spacing: Some(Spacing {
                        before: Some(100),
                        after: Some(100),
                        line: Some(240),
                    }),
                    ..Default::default()
                },
                runs: vec![Inline::TextRun(TextRun {
                    text: format!("Paragraph {i}"),
                    properties: RunProperties::default(),
                })],
                floats: Vec::new(),
                section_properties: None,
            }));
        }
        let doc = make_doc(blocks);
        let pages = layout(&doc, &LayoutConfig::default());
        assert!(pages.len() > 1, "Expected multiple pages, got {}", pages.len());
    }

    #[test]
    fn layout_centered_text() {
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties {
                alignment: Some(Alignment::Center),
                ..Default::default()
            },
            runs: vec![Inline::TextRun(TextRun {
                text: "Center".to_string(),
                properties: RunProperties::default(),
            })],
            floats: Vec::new(),
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        if let Some(DrawCommand::Text { x, .. }) = pages[0].commands.first() {
            // Centered text should not start at the left margin
            assert!(*x > config.margin_left);
        }
    }

    #[test]
    fn tab_stop_default_interval() {
        // Default interval of 36pt (720 twips), current x at 10pt
        let pos = find_next_tab_stop(10.0, &[], 36.0);
        assert!((pos - 36.0).abs() < 0.1);

        // Current x at 37pt, should go to 72pt
        let pos = find_next_tab_stop(37.0, &[], 36.0);
        assert!((pos - 72.0).abs() < 0.1);
    }

    #[test]
    fn tab_stop_custom() {
        let stops = vec![
            TabStop { position: 2880, stop_type: TabStopType::Left },  // 144pt
            TabStop { position: 5760, stop_type: TabStopType::Left },  // 288pt
        ];
        // Current x at 10pt, should go to first custom stop (144pt)
        let pos = find_next_tab_stop(10.0, &stops, 36.0);
        assert!((pos - 144.0).abs() < 0.1);

        // Current x at 145pt, should go to second custom stop (288pt)
        let pos = find_next_tab_stop(145.0, &stops, 36.0);
        assert!((pos - 288.0).abs() < 0.1);

        // Current x past all custom stops, falls back to default interval
        let pos = find_next_tab_stop(300.0, &stops, 36.0);
        assert!((pos - 324.0).abs() < 0.1); // 9 * 36 = 324
    }
}
