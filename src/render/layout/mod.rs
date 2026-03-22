mod fragment;
mod header_footer;
mod measurer;
mod paragraph;
mod table;

pub use measurer::TextMeasurer;

use std::collections::HashMap;
use std::rc::Rc;

use crate::model::*;
use crate::units::*;
use fragment::DocDefaultsLayout;

/// US Letter page width in points (8.5 inches).
const US_LETTER_WIDTH_PT: f32 = 612.0;
/// US Letter page height in points (11 inches).
const US_LETTER_HEIGHT_PT: f32 = 792.0;
/// Default page margin in points (1 inch).
const DEFAULT_PAGE_MARGIN_PT: f32 = 72.0;

/// Pre-decoded Skia images keyed by relationship ID.
/// All images are decoded upfront in `new()` — lookups are free.
pub(crate) struct ImageCache {
    images: HashMap<String, Rc<skia_safe::Image>>,
}

impl ImageCache {
    /// Decode all images from the store upfront.
    pub fn new(store: &ImageStore) -> Self {
        let mut images = HashMap::with_capacity(store.len());
        for (rel_id, data) in store {
            let skia_data = skia_safe::Data::new_copy(data);
            if let Some(image) = skia_safe::Image::from_encoded(skia_data) {
                images.insert(rel_id.clone(), Rc::new(image));
            }
        }
        Self { images }
    }

    /// Check whether a given rel_id has a decoded image.
    pub fn contains(&self, rel_id: &str) -> bool {
        self.images.contains_key(rel_id)
    }

    /// Get a decoded image by rel_id. Panics if not found.
    pub fn get(&self, rel_id: &str) -> Rc<skia_safe::Image> {
        Rc::clone(self.images.get(rel_id).expect("image not in cache"))
    }
}

/// Page layout configuration in points (1 point = 1/72 inch).
/// Built internally from document `SectionProperties`.
#[derive(Debug, Clone, Copy)]
pub(crate) struct LayoutConfig {
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
    /// US Letter (8.5 × 11 in) with 1-inch margins.
    ///
    /// OOXML (ISO/IEC 29500-1, §17.6.2) does not mandate `w:sectPr` on the
    /// document body; when it is absent, consumers are expected to fall back
    /// to application-defined defaults. These match Microsoft Word's defaults.
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
    /// Build a layout config from section properties, using US Letter defaults
    /// for any values not specified.
    fn from_section(sect: &SectionProperties) -> Self {
        let mut cfg = Self::default();
        apply_section_to_config(&mut cfg, sect);
        cfg
    }

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
        image: Rc<skia_safe::Image>,
    },
    Rect {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        color: (u8, u8, u8),
    },
    LinkAnnotation {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        url: String,
    },
}

#[derive(Debug, Clone)]
pub struct LayoutedPage {
    pub commands: Vec<DrawCommand>,
    pub page_width: f32,
    pub page_height: f32,
}

/// Perform layout on a document, producing positioned draw commands per page.
///
/// Page dimensions and margins are derived from the document's section properties,
/// falling back to US Letter with 1-inch margins when not specified.
pub fn layout(doc: &Document, font_mgr: &skia_safe::FontMgr) -> Vec<LayoutedPage> {
    // Build per-section configs from inline section breaks + final section.
    let mut section_configs: Vec<LayoutConfig> = Vec::new();
    for block in &doc.blocks {
        if let Block::Paragraph(p) = block {
            if let Some(ref sect) = p.section_properties {
                section_configs.push(LayoutConfig::from_section(sect));
            }
        }
    }
    if let Some(ref sect) = doc.final_section {
        section_configs.push(LayoutConfig::from_section(sect));
    }

    let initial_config = section_configs.first().copied().unwrap_or_default();
    let mut next_configs = section_configs.into_iter().skip(1).collect::<Vec<_>>();
    next_configs.reverse();

    let default_tab_stop_pt = crate::dimension::Pt::from(doc.default_tab_stop).raw();
    let doc_defaults = DocDefaultsLayout::from_document(doc);
    let image_cache = ImageCache::new(&doc.images);
    let mut effective_config = initial_config;
    if let Some(ref header) = doc.default_header {
        let pre_measurer = measurer::TextMeasurer::with_font_mgr(font_mgr.clone());
        let (_, header_bottom) = header_footer::layout_header_footer_blocks(
            &header.blocks,
            initial_config.margin_left,
            initial_config.header_margin,
            initial_config.content_width(),
            initial_config.margin_top,
            initial_config.page_height,
            initial_config.page_width,
            initial_config.page_height,
            &doc_defaults,
            &pre_measurer,
            default_tab_stop_pt,
            None,
            &image_cache,
        );
        if header_bottom > effective_config.margin_top {
            effective_config.margin_top = header_bottom;
        }
    }

    let mut layouter = Layouter::new(
        &effective_config,
        next_configs,
        default_tab_stop_pt,
        doc_defaults,
        font_mgr.clone(),
        image_cache,
    );

    let blocks = &doc.blocks;
    for (i, block) in blocks.iter().enumerate() {
        let next_is_table = blocks
            .get(i + 1)
            .is_some_and(|b| matches!(b, Block::Table(_)));
        layouter.layout_block(block, next_is_table);
    }

    let (mut pages, image_cache) = layouter.finish();

    // Render headers and footers on each page
    let hf_defaults = DocDefaultsLayout::from_document(doc);
    header_footer::render_headers_footers(
        &mut pages,
        &hf_defaults,
        default_tab_stop_pt,
        &initial_config,
        font_mgr,
        &image_cache,
    );

    pages
}

fn apply_section_to_config(config: &mut LayoutConfig, sect: &SectionProperties) {
    use crate::dimension::Pt;
    if let Some(ps) = &sect.page_size {
        config.page_width = Pt::from(ps.width).raw();
        config.page_height = Pt::from(ps.height).raw();
    }
    if let Some(pm) = &sect.page_margins {
        config.margin_top = Pt::from(pm.top).raw();
        config.margin_right = Pt::from(pm.right).raw();
        config.margin_bottom = Pt::from(pm.bottom).raw();
        config.margin_left = Pt::from(pm.left).raw();
        config.header_margin = Pt::from(pm.header).raw();
        config.footer_margin = Pt::from(pm.footer).raw();
    }
}

/// Offset all y-coordinates in a draw command by a given amount.
/// Used to translate relative-positioned commands to absolute page positions.
pub(super) fn offset_command(cmd: &DrawCommand, y_offset: f32) -> DrawCommand {
    match cmd {
        DrawCommand::Text {
            x,
            y,
            text,
            font_family,
            char_spacing_pt,
            font_size,
            bold,
            italic,
            color,
        } => DrawCommand::Text {
            x: *x,
            y: y_offset + y,
            text: text.clone(),
            font_family: font_family.clone(),
            char_spacing_pt: *char_spacing_pt,
            font_size: *font_size,
            bold: *bold,
            italic: *italic,
            color: *color,
        },
        DrawCommand::Underline {
            x1,
            y1,
            x2,
            y2,
            color,
            width,
        } => DrawCommand::Underline {
            x1: *x1,
            y1: y_offset + y1,
            x2: *x2,
            y2: y_offset + y2,
            color: *color,
            width: *width,
        },
        DrawCommand::Image {
            x,
            y,
            width,
            height,
            image,
        } => DrawCommand::Image {
            x: *x,
            y: y_offset + y,
            width: *width,
            height: *height,
            image: image.clone(),
        },
        DrawCommand::Rect {
            x,
            y,
            width,
            height,
            color,
        } => DrawCommand::Rect {
            x: *x,
            y: y_offset + y,
            width: *width,
            height: *height,
            color: *color,
        },
        DrawCommand::Line { .. } => cmd.clone(),
        DrawCommand::LinkAnnotation {
            x,
            y,
            width,
            height,
            url,
        } => DrawCommand::LinkAnnotation {
            x: *x,
            y: y_offset + y,
            width: *width,
            height: *height,
            url: url.clone(),
        },
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
    /// Whether the previous paragraph had a bottom border (to suppress duplicate top borders).
    prev_para_had_bottom_border: bool,
    image_cache: ImageCache,
}

impl Layouter {
    fn new(
        config: &LayoutConfig,
        next_section_configs: Vec<LayoutConfig>,
        default_tab_stop_pt: f32,
        doc_defaults: DocDefaultsLayout,
        font_mgr: skia_safe::FontMgr,
        image_cache: ImageCache,
    ) -> Self {
        let measurer = TextMeasurer::with_font_mgr(font_mgr);
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
            prev_para_had_bottom_border: false,
            image_cache,
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
                line_rule: if s.line.is_some() {
                    s.line_rule
                } else {
                    defaults.line_rule
                },
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
        let table_defaults = table_spacing.unwrap_or(self.doc_defaults.table_cell_spacing);
        match para_spacing {
            Some(s) => Spacing {
                before: s.before.or(table_defaults.before),
                after: s.after.or(table_defaults.after),
                line: s.line.or(table_defaults.line),
                line_rule: if s.line.is_some() {
                    s.line_rule
                } else {
                    table_defaults.line_rule
                },
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
        cell.width
            .map(|w| crate::dimension::Pt::from(w).raw())
            .unwrap_or(72.0)
    }

    /// Resolve list label text and indentation for a paragraph with a list reference.
    fn resolve_list_label(&mut self, list_ref: &ListRef) -> Option<(String, f32, f32)> {
        let def = self.doc_defaults.numbering.get(&list_ref.num_id)?;
        let level = def.levels.get(list_ref.level as usize)?;

        let label = match &level.format {
            NumberFormat::Bullet(ch) => ch.clone(),
            NumberFormat::Decimal => {
                let counter = self
                    .list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                level.level_text.replace("%1", &counter.to_string())
            }
            NumberFormat::LowerLetter => {
                let counter = self
                    .list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                let ch = (b'a' + ((*counter - 1) % 26) as u8) as char;
                level.level_text.replace("%1", &ch.to_string())
            }
            NumberFormat::UpperLetter => {
                let counter = self
                    .list_counters
                    .entry((list_ref.num_id, list_ref.level))
                    .or_insert(level.start.saturating_sub(1));
                *counter += 1;
                let ch = (b'A' + ((*counter - 1) % 26) as u8) as char;
                level.level_text.replace("%1", &ch.to_string())
            }
            NumberFormat::LowerRoman | NumberFormat::UpperRoman => {
                let counter = self
                    .list_counters
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

        let indent_left = crate::dimension::Pt::from(level.indent_left).raw();
        let indent_hanging = crate::dimension::Pt::from(level.indent_hanging).raw();

        Some((label, indent_left, indent_hanging))
    }

    fn prune_floats(&mut self) {
        self.active_floats.retain(|f| self.cursor_y < f.page_y_end);
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
            Block::Table(t) => {
                self.prev_para_had_bottom_border = false;
                self.layout_table(t, next_is_table);
            }
        }
    }

    fn finish(mut self) -> (Vec<LayoutedPage>, ImageCache) {
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
        (self.pages, self.image_cache)
    }
}

#[cfg(test)]
mod tests;
