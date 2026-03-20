mod fragment;
mod header_footer;
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
        font_family: std::rc::Rc<str>,
        char_spacing_pt: f32,
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
        data: ImageData,
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

    let default_tab_stop_pt = twips_to_pt(doc.default_tab_stop);
    let doc_defaults = DocDefaultsLayout::from_document(doc);

    // Pre-compute header extent to adjust margin_top if header content
    // (including float images) extends past the declared top margin.
    let mut effective_config = initial_config;
    if let Some(ref header) = doc.default_header {
        let pre_measurer = measurer::TextMeasurer::new();
        let (_, header_bottom) = header_footer::layout_header_footer_blocks(
            &header.blocks,
            initial_config.margin_left,
            initial_config.header_margin,
            initial_config.content_width(),
            initial_config.margin_top,
            initial_config.page_height,
            &doc_defaults,
            &pre_measurer,
            default_tab_stop_pt,
        );
        if header_bottom > effective_config.margin_top {
            effective_config.margin_top = header_bottom + HEADER_BODY_GAP_PT;
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
    let hf_defaults = DocDefaultsLayout::from_document(doc);
    header_footer::render_headers_footers(&mut pages, &hf_defaults, default_tab_stop_pt, &initial_config);

    pages
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
    /// Counters for numbered lists: (numId, level) -> current count.
    list_counters: std::collections::HashMap<(u32, u32), u32>,
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
            list_counters: std::collections::HashMap::new(),
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
                line_rule: if s.line.is_some() { s.line_rule } else { defaults.line_rule },
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
                line_rule: if s.line.is_some() { s.line_rule } else { table_defaults.line_rule },
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

    /// Resolve list label text and indentation for a paragraph with a list reference.
    fn resolve_list_label(&mut self, list_ref: &ListRef) -> Option<(String, f32, f32)> {
        let def = self.doc_defaults.numbering.get(&list_ref.num_id)?;
        let level = def.levels.get(list_ref.level as usize)?;

        let label = match &level.format {
            NumberFormat::Bullet(ch) => ch.clone(),
            NumberFormat::Decimal => {
                let counter = self.list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                level.level_text.replace("%1", &counter.to_string())
            }
            NumberFormat::LowerLetter => {
                let counter = self.list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                let ch = (b'a' + ((*counter - 1) % 26) as u8) as char;
                level.level_text.replace("%1", &ch.to_string())
            }
            NumberFormat::UpperLetter => {
                let counter = self.list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                let ch = (b'A' + ((*counter - 1) % 26) as u8) as char;
                level.level_text.replace("%1", &ch.to_string())
            }
            NumberFormat::LowerRoman | NumberFormat::UpperRoman => {
                let counter = self.list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                // Simple roman numeral conversion
                let roman = header_footer::to_roman(*counter);
                let roman = if matches!(level.format, NumberFormat::LowerRoman) {
                    roman.to_lowercase()
                } else {
                    roman
                };
                level.level_text.replace("%1", &roman)
            }
        };

        let indent_left = twips_to_pt(level.indent_left);
        let indent_hanging = twips_to_pt(level.indent_hanging);

        Some((label, indent_left, indent_hanging))
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
mod tests;
