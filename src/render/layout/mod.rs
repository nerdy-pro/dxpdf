mod fragment;
mod measurer;
mod paragraph;
mod table;

pub use measurer::TextMeasurer;

use crate::model::*;
use crate::units::*;
use fragment::DocDefaultsLayout;

/// Page layout configuration in points (1 point = 1/72 inch).
#[derive(Debug, Clone, Copy)]
pub struct LayoutConfig {
    pub page_width: f32,
    pub page_height: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
    /// Distance from page top to header content.
    pub header_margin: f32,
    /// Distance from page bottom to footer content.
    pub footer_margin: f32,
}

impl Default for LayoutConfig {
    fn default() -> Self {
        Self {
            page_width: US_LETTER_WIDTH_PT,
            page_height: US_LETTER_HEIGHT_PT,
            margin_top: DEFAULT_PAGE_MARGIN_PT,
            margin_bottom: DEFAULT_PAGE_MARGIN_PT,
            margin_left: DEFAULT_PAGE_MARGIN_PT,
            margin_right: DEFAULT_PAGE_MARGIN_PT,
            header_margin: DEFAULT_PAGE_MARGIN_PT / 2.0,
            footer_margin: DEFAULT_PAGE_MARGIN_PT / 2.0,
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
    Rect {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        color: (u8, u8, u8),
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

    let initial_config = section_configs.first().copied().unwrap_or(*config);
    let mut next_configs = section_configs.into_iter().skip(1).collect::<Vec<_>>();
    next_configs.reverse();

    let default_tab_stop_pt = doc.default_tab_stop as f32 / 20.0;
    let doc_defaults = DocDefaultsLayout {
        font_size_half_pts: doc.default_font_size,
        font_family: doc.default_font_family.clone(),
        default_spacing: doc.default_spacing,
        default_cell_margins: doc.default_cell_margins,
        table_cell_spacing: doc.table_cell_spacing,
        default_table_borders: doc.default_table_borders,
        default_header: doc.default_header.clone(),
        default_footer: doc.default_footer.clone(),
    };
    // Pre-compute header extent to adjust margin_top if header content
    // (including float images) extends past the declared top margin.
    let mut effective_config = initial_config;
    if let Some(ref header) = doc.default_header {
        let pre_measurer = measurer::TextMeasurer::new();
        let pre_defaults = DocDefaultsLayout {
            font_size_half_pts: doc.default_font_size,
            font_family: doc.default_font_family.clone(),
            default_spacing: doc.default_spacing,
            default_cell_margins: doc.default_cell_margins,
            table_cell_spacing: doc.table_cell_spacing,
            default_table_borders: doc.default_table_borders,
            default_header: None,
            default_footer: None,
        };
        let (_, header_bottom) = layout_header_footer_blocks(
            &header.blocks,
            initial_config.margin_left,
            initial_config.header_margin,
            initial_config.content_width(),
            initial_config.margin_top,
            initial_config.page_height,
            &pre_defaults,
            &pre_measurer,
            default_tab_stop_pt,
        );
        if header_bottom > effective_config.margin_top {
            effective_config.margin_top = header_bottom + 4.0; // small gap
        }
    }

    let mut layouter =
        Layouter::new(&effective_config, next_configs, default_tab_stop_pt, doc_defaults);

    let blocks = &doc.blocks;
    for (i, block) in blocks.iter().enumerate() {
        let next_is_table = blocks.get(i + 1).is_some_and(|b| matches!(b, Block::Table(_)));
        layouter.layout_block(block, next_is_table);
    }

    let mut pages = layouter.finish();

    // Render headers and footers on each page
    let hf_defaults = DocDefaultsLayout {
        font_size_half_pts: doc.default_font_size,
        font_family: doc.default_font_family.clone(),
        default_spacing: doc.default_spacing,
        default_cell_margins: doc.default_cell_margins,
        table_cell_spacing: doc.table_cell_spacing,
        default_table_borders: doc.default_table_borders,
        default_header: doc.default_header.clone(),
        default_footer: doc.default_footer.clone(),
    };
    render_headers_footers(&mut pages, &hf_defaults, default_tab_stop_pt, &initial_config);

    pages
}

/// Render header and footer content on each page.
fn render_headers_footers(
    pages: &mut [LayoutedPage],
    doc_defaults: &DocDefaultsLayout,
    default_tab_stop_pt: f32,
    config: &LayoutConfig,
) {
    let measurer = measurer::TextMeasurer::new();

    for page in pages.iter_mut() {
        let page_width = page.page_width;
        let page_height = page.page_height;
        let margin_left = config.margin_left;
        let margin_right = config.margin_right;
        let content_width = page_width - margin_left - margin_right;
        let header_y = config.header_margin;
        // Footer starts at the body's bottom margin boundary
        let footer_y = page_height - config.margin_bottom;

        let margin_top = config.margin_top;
        let margin_bottom = config.margin_bottom;

        // Render header
        if let Some(ref header) = doc_defaults.default_header {
            let (commands, _header_bottom) = layout_header_footer_blocks(
                &header.blocks,
                margin_left,
                header_y,
                content_width,
                margin_top,
                page_height, // don't clip header text — let it extend if needed
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
            );
            // Insert header commands at the beginning so they render behind body
            let body_commands = std::mem::take(&mut page.commands);
            page.commands = commands;
            page.commands.extend(body_commands);
        }

        // Render footer
        if let Some(ref footer) = doc_defaults.default_footer {
            let (commands, _) = layout_header_footer_blocks(
                &footer.blocks,
                margin_left,
                footer_y,
                content_width,
                margin_bottom,
                page_height,
                doc_defaults,
                &measurer,
                default_tab_stop_pt,
            );
            page.commands.extend(commands);
        }
    }
}

/// Layout header/footer blocks at a fixed position.
/// Returns (draw_commands, max_y_extent) where max_y_extent is the bottommost
/// y coordinate used by any content including float images.
fn layout_header_footer_blocks(
    blocks: &[Block],
    x_start: f32,
    y_start: f32,
    content_width: f32,
    margin_extent: f32,
    y_limit: f32,
    defaults: &DocDefaultsLayout,
    measurer: &measurer::TextMeasurer,
    default_tab_stop_pt: f32,
) -> (Vec<DrawCommand>, f32) {
    use fragment::*;

    let mut commands = Vec::new();
    let mut cursor_y = y_start;
    let mut max_y = y_start;

    for block in blocks {
        // Stop rendering if we've exceeded the allowed area
        if cursor_y >= y_limit {
            break;
        }
        if let Block::Paragraph(para) = block {
            let spacing = match para.properties.spacing {
                Some(s) => s,
                None => Spacing::default(),
            };
            cursor_y += spacing.before_pt();

            // Render floating images with alignment support
            for float in &para.floats {
                if float.data.is_empty() {
                    continue;
                }
                let scale = f32::min(
                    1.0,
                    content_width / float.width_pt.max(1.0),
                );
                let img_w = float.width_pt * scale;
                let img_h = float.height_pt * scale;
                // Use alignment if specified, otherwise offset
                let img_x = match float.align_h.as_deref() {
                    Some("right") => x_start + content_width - img_w,
                    Some("center") => x_start + (content_width - img_w) / 2.0,
                    Some("left") => x_start,
                    _ => x_start + float.offset_x_pt,
                };
                let img_y = match float.align_v.as_deref() {
                    Some("center") => (margin_extent - img_h) / 2.0,
                    Some("bottom") => margin_extent - img_h,
                    Some("top") => 0.0,
                    _ => cursor_y + float.offset_y_pt,
                };
                max_y = max_y.max(img_y + img_h);
                commands.push(DrawCommand::Image {
                    x: img_x,
                    y: img_y,
                    width: img_w,
                    height: img_h,
                    data: float.data.clone(),
                });
            }

            let fragments = collect_fragments(
                &para.runs,
                content_width,
                100.0, // max height for images
                defaults,
                measurer,
            );

            if fragments.is_empty() {
                cursor_y += spacing.line_pt();
                cursor_y += spacing.after_pt();
                continue;
            }

            let indent = para.properties.indentation.unwrap_or_default();
            let mut line_start = 0;
            let mut is_first_line = true;

            while line_start < fragments.len() {
                // Skip leading spaces
                if !is_first_line {
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

                let avail = content_width - indent.left_pt() - indent.right_pt();
                let (line_end, _) = fit_fragments(&fragments[line_start..], avail);
                let actual_end = line_start + line_end.max(1);

                let frag_height = fragments[line_start..actual_end]
                    .iter()
                    .map(|f| f.height())
                    .fold(0.0_f32, f32::max);
                let line_height = match spacing.line_pt_opt() {
                    Some(lh) => frag_height.max(lh),
                    None => frag_height,
                };
                cursor_y += line_height;

                let used_width = measure_fragments(&fragments[line_start..actual_end]);
                let x_offset = match para.properties.alignment {
                    Some(Alignment::Center) => (avail - used_width) / 2.0,
                    Some(Alignment::Right) => avail - used_width,
                    _ => 0.0,
                };
                let mut x = x_start + indent.left_pt() + x_offset;

                for frag in &fragments[line_start..actual_end] {
                    match frag {
                        Fragment::Text {
                            text, font_family, font_size, bold, italic,
                            underline, color, measured_width, ..
                        } => {
                            let c = color.map(|c| (c.r, c.g, c.b)).unwrap_or((0, 0, 0));
                            commands.push(DrawCommand::Text {
                                x, y: cursor_y, text: text.clone(),
                                font_family: font_family.clone(),
                                font_size: *font_size, bold: *bold, italic: *italic,
                                color: c,
                            });
                            if *underline {
                                commands.push(DrawCommand::Underline {
                                    x1: x, y1: cursor_y + crate::units::UNDERLINE_Y_OFFSET,
                                    x2: x + measured_width,
                                    y2: cursor_y + crate::units::UNDERLINE_Y_OFFSET,
                                    color: c, width: crate::units::UNDERLINE_STROKE_WIDTH,
                                });
                            }
                            x += measured_width;
                        }
                        Fragment::Image { width, height, data } => {
                            commands.push(DrawCommand::Image {
                                x, y: cursor_y - height,
                                width: *width, height: *height,
                                data: data.clone(),
                            });
                            x += width;
                        }
                        Fragment::Tab { .. } => {
                            let rel_x = x - x_start;
                            let next = find_next_tab_stop(
                                rel_x, &para.properties.tab_stops, default_tab_stop_pt,
                            );
                            x = x_start + next;
                        }
                        Fragment::LineBreak { .. } => {}
                    }
                }

                line_start = actual_end;
                is_first_line = false;
            }

            cursor_y += spacing.after_pt();
        }
    }

    max_y = max_y.max(cursor_y);
    (commands, max_y)
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
        config.header_margin = pm.header_pt();
        config.footer_margin = pm.footer_pt();
    }
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
    next_section_configs: Vec<LayoutConfig>,
    default_tab_stop_pt: f32,
    doc_defaults: DocDefaultsLayout,
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

    fn section_break(&mut self) {
        self.new_page();
        if let Some(next_config) = self.next_section_configs.pop() {
            self.config = next_config;
            self.current_page.page_width = self.config.page_width;
            self.current_page.page_height = self.config.page_height;
            self.cursor_y = self.config.margin_top;
        }
    }

    fn float_adjustment(&self, line_top: f32, line_bottom: f32) -> (f32, f32) {
        let gap = FLOAT_TEXT_GAP_PT;
        let mut x_shift = 0.0_f32;
        let mut width_reduction = 0.0_f32;
        for f in &self.active_floats {
            if line_top < f.page_y_end && line_bottom > f.page_y_start {
                let shift = (f.page_x - self.config.margin_left) + f.width + gap;
                x_shift = x_shift.max(shift);
                width_reduction = width_reduction.max(shift);
            }
        }
        (x_shift, width_reduction)
    }

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

    /// Resolve spacing for paragraphs inside table cells.
    /// Table style spacing (e.g., after=0) takes priority over document defaults.
    fn resolve_cell_spacing(
        &self,
        para_spacing: Option<Spacing>,
        table_spacing: Option<Spacing>,
    ) -> Spacing {
        let table_defaults = table_spacing
            .unwrap_or(self.doc_defaults.table_cell_spacing);
        match para_spacing {
            Some(s) => Spacing {
                before: s.before.or(table_defaults.before),
                after: s.after.or(table_defaults.after),
                line: s.line.or(table_defaults.line),
            },
            None => table_defaults,
        }
    }

    /// Get the width of a table cell, accounting for grid_span.
    fn cell_width(&self, grid_col_idx: usize, cell: &TableCell, col_widths: &[f32]) -> f32 {
        if !col_widths.is_empty() {
            let span = cell.grid_span.max(1) as usize;
            return (grid_col_idx..grid_col_idx + span)
                .map(|i| col_widths.get(i).copied().unwrap_or(72.0))
                .sum();
        }
        cell.width_pt().unwrap_or(72.0)
    }

    fn prune_floats(&mut self) {
        self.active_floats
            .retain(|f| self.cursor_y < f.page_y_end);
    }

    fn layout_block(&mut self, block: &Block, next_is_table: bool) {
        self.prune_floats();
        match block {
            Block::Paragraph(p) => {
                self.layout_paragraph(p);
                if p.section_properties.is_some() {
                    self.section_break();
                }
            }
            Block::Table(t) => self.layout_table(t, next_is_table),
        }
    }

    fn finish(mut self) -> Vec<LayoutedPage> {
        if !self.current_page.commands.is_empty() {
            self.pages.push(self.current_page);
        }
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

#[cfg(test)]
mod tests {
    use super::fragment::find_next_tab_stop;
    use super::*;

    fn make_doc(blocks: Vec<Block>) -> Document {
        Document { blocks, ..Document::default() }
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

    fn make_cell(text: &str) -> TableCell {
        TableCell {
            blocks: vec![Block::Paragraph(Paragraph {
                properties: ParagraphProperties::default(),
                runs: vec![Inline::TextRun(TextRun {
                    text: text.to_string(),
                    properties: RunProperties::default(),
                })],
                floats: Vec::new(),
                section_properties: None,
            })],
            width: None,
            grid_span: 1,
            vertical_merge: None,
            cell_margins: None,
            cell_borders: None,
            shading: None,
        }
    }

    fn make_spanned_cell(text: &str, span: u32) -> TableCell {
        let mut cell = make_cell(text);
        cell.grid_span = span;
        cell
    }

    fn extract_lines(pages: &[LayoutedPage]) -> Vec<(f32, f32, f32, f32)> {
        let mut lines = Vec::new();
        for page in pages {
            for cmd in &page.commands {
                if let DrawCommand::Line { x1, y1, x2, y2, .. } = cmd {
                    lines.push((*x1, *y1, *x2, *y2));
                }
            }
        }
        lines
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
        assert!(pages[0]
            .commands
            .iter()
            .any(|c| matches!(c, DrawCommand::Text { .. })));
    }

    #[test]
    fn layout_page_break() {
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
        assert!(
            pages.len() > 1,
            "Expected multiple pages, got {}",
            pages.len()
        );
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
            assert!(*x > config.margin_left);
        }
    }

    #[test]
    fn tab_stop_default_interval() {
        let pos = find_next_tab_stop(10.0, &[], 36.0);
        assert!((pos - 36.0).abs() < 0.1);
        let pos = find_next_tab_stop(37.0, &[], 36.0);
        assert!((pos - 72.0).abs() < 0.1);
    }

    #[test]
    fn tab_stop_custom() {
        let stops = vec![
            TabStop {
                position: 2880,
                stop_type: TabStopType::Left,
            },
            TabStop {
                position: 5760,
                stop_type: TabStopType::Left,
            },
        ];
        let pos = find_next_tab_stop(10.0, &stops, 36.0);
        assert!((pos - 144.0).abs() < 0.1);
        let pos = find_next_tab_stop(145.0, &stops, 36.0);
        assert!((pos - 288.0).abs() < 0.1);
        let pos = find_next_tab_stop(300.0, &stops, 36.0);
        assert!((pos - 324.0).abs() < 0.1);
    }

    #[test]
    fn table_borders_simple_2x2() {
        let table = Table {
            rows: vec![
                TableRow {
                    height: None,
                    cells: vec![make_cell("A1"), make_cell("B1")],
                },
                TableRow {
                    height: None,
                    cells: vec![make_cell("A2"), make_cell("B2")],
                },
            ],
            grid_cols: vec![2880, 2880],
            default_cell_margins: None,
            cell_spacing: None,
            borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let lines = extract_lines(&pages);

        let margin = config.margin_left;
        let scale = config.content_width() / 288.0;
        let col_w = 144.0 * scale;
        let x_left = margin;
        let x_mid = margin + col_w;
        let x_right = margin + 2.0 * col_w;

        let count_v_at = |x: f32| -> usize {
            lines
                .iter()
                .filter(|(x1, _, x2, _)| (x1 - x).abs() < 1.0 && (x2 - x).abs() < 1.0)
                .count()
        };
        assert!(count_v_at(x_left) >= 2);
        assert!(count_v_at(x_mid) >= 4);
        assert!(count_v_at(x_right) >= 2);
    }

    #[test]
    fn table_borders_with_gridspan() {
        let table = Table {
            rows: vec![
                TableRow {
                    height: None,
                    cells: vec![make_spanned_cell("AB", 2), make_cell("C")],
                },
                TableRow {
                    height: None,
                    cells: vec![make_cell("A"), make_cell("B"), make_cell("C")],
                },
            ],
            grid_cols: vec![2000, 2000, 2000],
            default_cell_margins: None,
            cell_spacing: None,
            borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let lines = extract_lines(&pages);

        let margin = config.margin_left;
        let scale = config.content_width() / 300.0;
        let col_w = 100.0 * scale;
        let x0 = margin;
        let x1 = margin + col_w;
        let x2 = margin + 2.0 * col_w;
        let x3 = margin + 3.0 * col_w;

        let h_lines_at_top: Vec<_> = lines
            .iter()
            .filter(|(_, y1, _, y2)| {
                let min_y = lines
                    .iter()
                    .filter(|(lx1, _, lx2, _)| (lx1 - lx2).abs() > 1.0)
                    .map(|(_, y, _, _)| *y)
                    .fold(f32::MAX, f32::min);
                (y1 - min_y).abs() < 0.5 && (y2 - min_y).abs() < 0.5
            })
            .collect();

        assert!(h_lines_at_top
            .iter()
            .any(|(lx1, _, lx2, _)| (lx1 - x0).abs() < 1.0 && (lx2 - x2).abs() < 1.0));
        assert!(h_lines_at_top
            .iter()
            .any(|(lx1, _, lx2, _)| (lx1 - x2).abs() < 1.0 && (lx2 - x3).abs() < 1.0));
        assert!(lines
            .iter()
            .any(|(lx1, _, lx2, _)| (lx1 - x1).abs() < 1.0 && (lx2 - x1).abs() < 1.0));
    }

    #[test]
    fn table_borders_alignment_across_rows() {
        let table = Table {
            rows: vec![
                TableRow {
                    height: None,
                    cells: vec![
                        make_spanned_cell("AB", 2),
                        make_cell("C"),
                        make_cell("D"),
                    ],
                },
                TableRow {
                    height: None,
                    cells: vec![
                        make_cell("A"),
                        make_cell("B"),
                        make_cell("C"),
                        make_cell("D"),
                    ],
                },
            ],
            grid_cols: vec![1000, 1000, 1000, 1000],
            default_cell_margins: None,
            cell_spacing: None,
            borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let lines = extract_lines(&pages);

        let scale = config.content_width() / 200.0;
        let cw = 50.0 * scale;
        let margin = config.margin_left;
        let x_after_2cols = margin + 2.0 * cw;

        let v_count = lines
            .iter()
            .filter(|(x1, _, x2, _)| {
                (x1 - x_after_2cols).abs() < 1.0 && (x2 - x_after_2cols).abs() < 1.0
            })
            .count();
        assert!(v_count >= 4);

        let right_edge = margin + 4.0 * cw;
        let v_right = lines
            .iter()
            .filter(|(x1, _, x2, _)| {
                (x1 - right_edge).abs() < 1.0 && (x2 - right_edge).abs() < 1.0
            })
            .count();
        assert!(v_right >= 2);
    }

    #[test]
    fn table_borders_tcw_vs_grid_alignment() {
        let table = Table {
            rows: vec![
                TableRow {
                    height: None,
                    cells: vec![
                        {
                            let mut c = make_spanned_cell("AB", 2);
                            c.width = Some(300);
                            c
                        },
                        {
                            let mut c = make_cell("C");
                            c.width = Some(300);
                            c
                        },
                    ],
                },
                TableRow {
                    height: None,
                    cells: vec![
                        {
                            let mut c = make_cell("A");
                            c.width = Some(100);
                            c
                        },
                        {
                            let mut c = make_cell("B");
                            c.width = Some(200);
                            c
                        },
                        {
                            let mut c = make_cell("C");
                            c.width = Some(300);
                            c
                        },
                    ],
                },
            ],
            grid_cols: vec![100, 200, 300],
            default_cell_margins: None,
            cell_spacing: None,
            borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let lines = extract_lines(&pages);

        let scale = config.content_width() / 30.0;
        let margin = config.margin_left;
        let boundary_12 = margin + 15.0 * scale;

        let v_at_boundary_row0 = lines
            .iter()
            .filter(|(x1, y1, x2, _)| {
                (x1 - boundary_12).abs() < 1.0
                    && (x2 - boundary_12).abs() < 1.0
                    && *y1 < (margin + 50.0)
            })
            .count();
        assert!(v_at_boundary_row0 >= 1);
    }
}
