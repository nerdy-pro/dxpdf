use crate::model::*;

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
}

#[derive(Debug, Clone)]
pub struct LayoutedPage {
    pub commands: Vec<DrawCommand>,
}

/// Perform layout on a document, producing positioned draw commands per page.
pub fn layout(doc: &Document, config: &LayoutConfig) -> Vec<LayoutedPage> {
    let mut layouter = Layouter::new(config);

    for block in &doc.blocks {
        layouter.layout_block(block);
    }

    layouter.finish()
}

struct Layouter {
    config: LayoutConfig,
    pages: Vec<LayoutedPage>,
    current_page: LayoutedPage,
    cursor_y: f32,
}

impl Layouter {
    fn new(config: &LayoutConfig) -> Self {
        Self {
            config: *config,
            pages: Vec::new(),
            current_page: LayoutedPage {
                commands: Vec::new(),
            },
            cursor_y: config.margin_top,
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
            },
        );
        self.pages.push(page);
        self.cursor_y = self.config.margin_top;
    }

    fn layout_block(&mut self, block: &Block) {
        match block {
            Block::Paragraph(p) => self.layout_paragraph(p),
            Block::Table(t) => self.layout_table(t),
        }
    }

    fn layout_paragraph(&mut self, para: &Paragraph) {
        let spacing = para.properties.spacing.unwrap_or_default();
        let indent = para.properties.indentation.unwrap_or_default();

        // Space before
        self.cursor_y += spacing.before_pt();

        if para.runs.is_empty() {
            // Empty paragraph - just add line spacing
            self.cursor_y += spacing.line_pt();
            self.cursor_y += spacing.after_pt();
            return;
        }

        // Collect all text fragments with their properties for line wrapping
        let fragments = collect_fragments(&para.runs);
        let content_width = self.config.content_width()
            - indent.left_pt()
            - indent.right_pt();
        let base_x = self.config.margin_left + indent.left_pt();

        // Simple line wrapping: estimate character width from font size
        let mut line_start = 0;
        let mut first_line = true;

        while line_start < fragments.len() {
            let first_line_offset = if first_line {
                indent.first_line_pt()
            } else {
                0.0
            };
            let available_width = content_width - first_line_offset;

            // Find how many fragments fit on this line
            let (line_end, line_width) =
                fit_fragments(&fragments[line_start..], available_width);
            let actual_end = line_start + line_end.max(1); // Always consume at least one fragment

            // Determine line height from the tallest fragment on this line
            let line_height = fragments[line_start..actual_end]
                .iter()
                .map(|f| f.font_size * 1.2)
                .fold(spacing.line_pt(), f32::max);

            // Check for page break
            if self.cursor_y + line_height > self.content_bottom() {
                self.new_page();
            }

            // Calculate x offset for alignment
            let used_width = if actual_end > line_start {
                measure_fragments(&fragments[line_start..actual_end])
            } else {
                line_width
            };
            let x_offset = match para.properties.alignment {
                Some(Alignment::Center) => (available_width - used_width) / 2.0,
                Some(Alignment::Right) => available_width - used_width,
                _ => 0.0,
            };

            let mut x = base_x + first_line_offset + x_offset;
            self.cursor_y += line_height;

            for frag in &fragments[line_start..actual_end] {
                let color = frag
                    .color
                    .map(|c| (c.r, c.g, c.b))
                    .unwrap_or((0, 0, 0));

                self.current_page.commands.push(DrawCommand::Text {
                    x,
                    y: self.cursor_y,
                    text: frag.text.clone(),
                    font_family: frag.font_family.clone(),
                    font_size: frag.font_size,
                    bold: frag.bold,
                    italic: frag.italic,
                    color,
                });

                let frag_width = estimate_text_width(&frag.text, frag.font_size);

                if frag.underline {
                    self.current_page.commands.push(DrawCommand::Underline {
                        x1: x,
                        y1: self.cursor_y + 2.0,
                        x2: x + frag_width,
                        y2: self.cursor_y + 2.0,
                        color,
                        width: 0.5,
                    });
                }

                x += frag_width;
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
        let col_width = content_width / num_cols as f32;
        let cell_padding = 4.0;

        for row in &table.rows {
            // Estimate row height
            let row_height = 20.0_f32; // Simple fixed row height for now

            if self.cursor_y + row_height > self.content_bottom() {
                self.new_page();
            }

            let row_top = self.cursor_y;

            for (col_idx, cell) in row.cells.iter().enumerate() {
                let cell_x = self.config.margin_left + col_idx as f32 * col_width;

                // Draw cell border (top and left)
                self.current_page.commands.push(DrawCommand::Line {
                    x1: cell_x,
                    y1: row_top,
                    x2: cell_x + col_width,
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

                // Render cell text (first paragraph only for simplicity)
                if let Some(Block::Paragraph(p)) = cell.blocks.first() {
                    let text: String = p
                        .runs
                        .iter()
                        .filter_map(|r| match r {
                            Inline::TextRun(tr) => Some(tr.text.as_str()),
                            Inline::Tab => Some("\t"),
                            Inline::LineBreak => Some("\n"),
                        })
                        .collect();

                    if !text.is_empty() {
                        let font_size = p
                            .runs
                            .iter()
                            .find_map(|r| match r {
                                Inline::TextRun(tr) => Some(tr.properties.font_size_pt()),
                                _ => None,
                            })
                            .unwrap_or(12.0);

                        self.current_page.commands.push(DrawCommand::Text {
                            x: cell_x + cell_padding,
                            y: row_top + row_height - cell_padding,
                            text,
                            font_family: "Helvetica".to_string(),
                            font_size,
                            bold: false,
                            italic: false,
                            color: (0, 0, 0),
                        });
                    }
                }

                // Right border for last column
                if col_idx == row.cells.len() - 1 {
                    self.current_page.commands.push(DrawCommand::Line {
                        x1: cell_x + col_width,
                        y1: row_top,
                        x2: cell_x + col_width,
                        y2: row_top + row_height,
                        color: (0, 0, 0),
                        width: 0.5,
                    });
                }
            }

            self.cursor_y += row_height;

            // Bottom border
            let cols_in_row = row.cells.len();
            let table_width = cols_in_row as f32 * col_width;
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
            });
        }
        self.pages
    }
}

/// A flattened text fragment for layout.
struct Fragment {
    text: String,
    font_family: String,
    font_size: f32,
    bold: bool,
    italic: bool,
    underline: bool,
    color: Option<Color>,
}

fn collect_fragments(runs: &[Inline]) -> Vec<Fragment> {
    let mut fragments = Vec::new();
    for run in runs {
        match run {
            Inline::TextRun(tr) => {
                // Split by spaces to allow word wrapping
                let words: Vec<&str> = tr.text.split_inclusive(' ').collect();
                for word in words {
                    fragments.push(Fragment {
                        text: word.to_string(),
                        font_family: tr
                            .properties
                            .font_family
                            .clone()
                            .unwrap_or_else(|| "Helvetica".to_string()),
                        font_size: tr.properties.font_size_pt(),
                        bold: tr.properties.bold,
                        italic: tr.properties.italic,
                        underline: tr.properties.underline,
                        color: tr.properties.color,
                    });
                }
            }
            Inline::Tab => {
                fragments.push(Fragment {
                    text: "    ".to_string(),
                    font_family: "Helvetica".to_string(),
                    font_size: 12.0,
                    bold: false,
                    italic: false,
                    underline: false,
                    color: None,
                });
            }
            Inline::LineBreak => {
                // Force a new line by pushing a special fragment
                // For simplicity, we just add a newline marker
            }
        }
    }
    fragments
}

/// Estimate text width using average character width (0.5 * font_size).
fn estimate_text_width(text: &str, font_size: f32) -> f32 {
    text.len() as f32 * font_size * 0.5
}

fn measure_fragments(fragments: &[Fragment]) -> f32 {
    fragments
        .iter()
        .map(|f| estimate_text_width(&f.text, f.font_size))
        .sum()
}

/// Find how many fragments fit within the available width.
/// Returns (count, total_width).
fn fit_fragments(fragments: &[Fragment], available_width: f32) -> (usize, f32) {
    let mut total_width = 0.0;
    for (i, frag) in fragments.iter().enumerate() {
        let w = estimate_text_width(&frag.text, frag.font_size);
        if total_width + w > available_width && i > 0 {
            return (i, total_width);
        }
        total_width += w;
    }
    (fragments.len(), total_width)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_doc(blocks: Vec<Block>) -> Document {
        Document { blocks }
    }

    fn simple_paragraph(text: &str) -> Block {
        Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![Inline::TextRun(TextRun {
                text: text.to_string(),
                properties: RunProperties::default(),
            })],
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
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        if let Some(DrawCommand::Text { x, .. }) = pages[0].commands.first() {
            // Centered text should not start at the left margin
            assert!(*x > config.margin_left);
        }
    }
}
