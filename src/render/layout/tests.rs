use std::rc::Rc;

use super::fragment::{
    find_next_tab_stop, fit_fragments, measure_lines, resolve_line_height,
    Fragment,
};
use super::header_footer::to_roman;
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
                hyperlink_url: None,
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
                    hyperlink_url: None,
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
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                runs: vec![Inline::TextRun(TextRun {
                    text: format!("Paragraph {i}"),
                    properties: RunProperties::default(),
                    hyperlink_url: None,
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
                hyperlink_url: None,
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

    // ---- Helpers ----

    fn extract_texts(pages: &[LayoutedPage]) -> Vec<(f32, f32, String)> {
        let mut texts = Vec::new();
        for page in pages {
            for cmd in &page.commands {
                if let DrawCommand::Text { x, y, text, .. } = cmd {
                    texts.push((*x, *y, text.clone()));
                }
            }
        }
        texts
    }

    fn extract_rects(pages: &[LayoutedPage]) -> Vec<(f32, f32, f32, f32, (u8, u8, u8))> {
        let mut rects = Vec::new();
        for page in pages {
            for cmd in &page.commands {
                if let DrawCommand::Rect { x, y, width, height, color } = cmd {
                    rects.push((*x, *y, *width, *height, *color));
                }
            }
        }
        rects
    }

    // ---- Paragraph spacing ----

    #[test]
    fn spacing_before_after_affects_position() {
        let doc = Document {
            blocks: vec![
                Block::Paragraph(Paragraph {
                    properties: ParagraphProperties {
                        spacing: Some(Spacing {
                            before: Some(200),
                            after: Some(200),
                            line: None,
                            ..Default::default()
                        }),
                        ..Default::default()
                    },
                    runs: vec![Inline::TextRun(TextRun {
                        text: "First".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                }),
                Block::Paragraph(Paragraph {
                    properties: ParagraphProperties {
                        spacing: Some(Spacing {
                            before: Some(100),
                            after: None,
                            line: None,
                            ..Default::default()
                        }),
                        ..Default::default()
                    },
                    runs: vec![Inline::TextRun(TextRun {
                        text: "Second".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                }),
            ],
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        assert!(texts.len() >= 2);

        let first_y = texts.iter().find(|(_, _, t)| t == "First").unwrap().1;
        let second_y = texts.iter().find(|(_, _, t)| t == "Second").unwrap().1;

        assert!(second_y > first_y, "Second ({second_y}) should be below First ({first_y})");
        let gap = second_y - first_y;
        assert!(gap > 14.0, "Gap ({gap}) should include after=10 + before=5");
    }

    // ---- Indentation ----

    #[test]
    fn left_indentation_shifts_text_right() {
        let doc = Document {
            blocks: vec![
                simple_paragraph("NoIndent"),
                Block::Paragraph(Paragraph {
                    properties: ParagraphProperties {
                        indentation: Some(Indentation {
                            left: Some(720), // 36pt
                            right: None,
                            first_line: None,
                        }),
                        ..Default::default()
                    },
                    runs: vec![Inline::TextRun(TextRun {
                        text: "Indented".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                }),
            ],
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);

        let no_indent_x = texts.iter().find(|(_, _, t)| t == "NoIndent").unwrap().0;
        let indented_x = texts.iter().find(|(_, _, t)| t == "Indented").unwrap().0;

        assert!(
            (indented_x - no_indent_x - 36.0).abs() < 1.0,
            "Indented should be 36pt right of NoIndent: got {indented_x} vs {no_indent_x}"
        );
    }

    // ---- Section breaks ----

    #[test]
    fn section_break_changes_page_dimensions() {
        let doc = Document {
            blocks: vec![
                Block::Paragraph(Paragraph {
                    properties: ParagraphProperties::default(),
                    runs: vec![Inline::TextRun(TextRun {
                        text: "Portrait".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: Some(SectionProperties {
                        page_size: Some(PageSize { width: 12240, height: 15840 }),
                        page_margins: None,
                        header: None, footer: None,
                        header_rel_id: None, footer_rel_id: None,
                    }),
                }),
                Block::Paragraph(Paragraph {
                    properties: ParagraphProperties::default(),
                    runs: vec![Inline::TextRun(TextRun {
                        text: "Landscape".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                }),
            ],
            final_section: Some(SectionProperties {
                page_size: Some(PageSize { width: 15840, height: 12240 }),
                page_margins: None,
                header: None, footer: None,
                header_rel_id: None, footer_rel_id: None,
            }),
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        assert!(pages.len() >= 2);
        assert!((pages[0].page_width - 612.0).abs() < 1.0);
        assert!((pages[0].page_height - 792.0).abs() < 1.0);
        assert!((pages[1].page_width - 792.0).abs() < 1.0);
        assert!((pages[1].page_height - 612.0).abs() < 1.0);
    }

    // ---- Adjacent tables ----

    #[test]
    fn adjacent_tables_no_gap() {
        let mk_table = |text: &str| Table {
            rows: vec![TableRow { height: None, cells: vec![make_cell(text)] }],
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = Document {
            blocks: vec![Block::Table(mk_table("T1")), Block::Table(mk_table("T2"))],
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let lines = extract_lines(&pages);
        let mut h_ys: Vec<f32> = lines.iter()
            .filter(|(_, y1, _, y2)| (y1 - y2).abs() < 0.1)
            .map(|(_, y, _, _)| *y)
            .collect();
        h_ys.sort_by(|a, b| a.partial_cmp(b).unwrap());
        h_ys.dedup_by(|a, b| (*a - *b).abs() < 0.5);
        // T1 top, T1 bottom=T2 top, T2 bottom = 3 unique y positions
        assert!(h_ys.len() <= 4,
            "Expected <=4 unique y positions (no gap): {:?}", h_ys);
    }

    // ---- vMerge ----

    #[test]
    fn vmerge_skips_content_and_border() {
        let table = Table {
            rows: vec![
                TableRow {
                    height: None,
                    cells: vec![
                        { let mut c = make_cell("Merged"); c.vertical_merge = Some(VerticalMerge::Restart); c },
                        make_cell("R0C1"),
                    ],
                },
                TableRow {
                    height: None,
                    cells: vec![
                        { let mut c = make_cell("Hidden"); c.vertical_merge = Some(VerticalMerge::Continue); c },
                        make_cell("R1C1"),
                    ],
                },
            ],
            grid_cols: vec![3000, 3000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);

        assert!(texts.iter().any(|(_, _, t)| t == "Merged"));
        assert!(!texts.iter().any(|(_, _, t)| t == "Hidden"),
            "vMerge continue cell content should not render");
    }

    // ---- Line breaks ----

    #[test]
    fn line_break_forces_new_line() {
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![
                Inline::TextRun(TextRun { text: "Before".into(), properties: RunProperties::default(), hyperlink_url: None }),
                Inline::LineBreak,
                Inline::TextRun(TextRun { text: "After".into(), properties: RunProperties::default(), hyperlink_url: None }),
            ],
            floats: Vec::new(),
            section_properties: None,
        })]);
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        let before = texts.iter().find(|(_, _, t)| t == "Before").unwrap();
        let after = texts.iter().find(|(_, _, t)| t == "After").unwrap();
        assert!(after.1 > before.1, "After y ({}) > Before y ({})", after.1, before.1);
    }

    // ---- Paragraph shading ----

    #[test]
    fn paragraph_shading_excludes_spacing() {
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties {
                spacing: Some(Spacing { before: Some(200), after: Some(200), line: None, ..Default::default() }),
                shading: Some(Color { r: 200, g: 200, b: 200 }),
                ..Default::default()
            },
            runs: vec![Inline::TextRun(TextRun {
                text: "Shaded".into(),
                properties: RunProperties::default(),
                hyperlink_url: None,
            })],
            floats: Vec::new(),
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let rects = extract_rects(&pages);
        let texts = extract_texts(&pages);

        assert!(!rects.is_empty(), "Should have a shading rect");
        let rect_top = rects[0].1;
        let rect_bottom = rects[0].1 + rects[0].3;
        let text_y = texts.iter().find(|(_, _, t)| t == "Shaded").unwrap().1;

        // Rect should NOT start at margin_top (spacing.before pushes it down)
        assert!(rect_top > config.margin_top,
            "Rect top ({rect_top}) should be below margin ({}) due to spacing.before", config.margin_top);
        // Text should be within rect
        assert!(text_y >= rect_top && text_y <= rect_bottom + 1.0,
            "Text y ({text_y}) should be within rect [{rect_top}, {rect_bottom}]");
    }

    // ---- to_roman ----

    #[test]
    fn roman_numerals() {
        assert_eq!(to_roman(1), "I");
        assert_eq!(to_roman(4), "IV");
        assert_eq!(to_roman(9), "IX");
        assert_eq!(to_roman(14), "XIV");
        assert_eq!(to_roman(42), "XLII");
        assert_eq!(to_roman(99), "XCIX");
        assert_eq!(to_roman(2024), "MMXXIV");
    }

    // ---- Row height ----

    #[test]
    fn row_height_minimum_respected() {
        let table = Table {
            rows: vec![TableRow {
                height: Some(1000), // 50pt minimum
                cells: vec![make_cell("Short")],
            }],
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let lines = extract_lines(&pages);
        let h_ys: Vec<f32> = lines.iter()
            .filter(|(_, y1, _, y2)| (y1 - y2).abs() < 0.1)
            .map(|(_, y, _, _)| *y)
            .collect();
        let min_y = h_ys.iter().cloned().fold(f32::MAX, f32::min);
        let max_y = h_ys.iter().cloned().fold(f32::MIN, f32::max);
        assert!(max_y - min_y >= 49.0,
            "Row height ({}) should be >= 50pt", max_y - min_y);
    }

    // ---- Right alignment ----

    #[test]
    fn right_alignment() {
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties {
                alignment: Some(Alignment::Right),
                ..Default::default()
            },
            runs: vec![Inline::TextRun(TextRun {
                text: "Right".into(),
                properties: RunProperties::default(),
                hyperlink_url: None,
            })],
            floats: Vec::new(),
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let texts = extract_texts(&pages);
        let text_x = texts.iter().find(|(_, _, t)| t == "Right").unwrap().0;
        assert!(text_x > config.margin_left + config.content_width() / 2.0,
            "Right-aligned text x ({text_x}) should be in right half");
    }

    // ==============================================================
    // resolve_line_height
    // ==============================================================

    #[test]
    fn resolve_line_height_none_returns_frag_height() {
        assert!((resolve_line_height(14.0, None) - 14.0).abs() < 0.01);
    }

    #[test]
    fn resolve_line_height_multiplier() {
        let h = resolve_line_height(14.0, Some(LineSpacing::Multiplier(1.15)));
        assert!((h - 16.1).abs() < 0.1, "14 * 1.15 = {h}");
    }

    #[test]
    fn resolve_line_height_fixed() {
        let h = resolve_line_height(14.0, Some(LineSpacing::Fixed(20.0)));
        assert!((h - 20.0).abs() < 0.01, "Fixed should be exactly 20pt, got {h}");
    }

    #[test]
    fn resolve_line_height_at_least_larger() {
        let h = resolve_line_height(14.0, Some(LineSpacing::AtLeast(20.0)));
        assert!((h - 20.0).abs() < 0.01, "AtLeast(20) with frag=14 → {h}");
    }

    #[test]
    fn resolve_line_height_at_least_smaller() {
        let h = resolve_line_height(24.0, Some(LineSpacing::AtLeast(20.0)));
        assert!((h - 24.0).abs() < 0.01, "AtLeast(20) with frag=24 → {h}");
    }

    // ==============================================================
    // fit_fragments
    // ==============================================================

    fn make_text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            text: text.to_string(),
            font_family: Rc::from("Helvetica"),
            font_size: 12.0,
            bold: false,
            italic: false,
            underline: false,
            color: None,
            shading: None,
            char_spacing_pt: 0.0,
            measured_width: width,
            measured_height: 14.0,
            hyperlink_url: None,
        }
    }

    #[test]
    fn fit_fragments_empty() {
        let (count, width) = fit_fragments(&[], 100.0);
        assert_eq!(count, 0);
        assert!((width).abs() < 0.01);
    }

    #[test]
    fn fit_fragments_single_fits() {
        let frags = [make_text_frag("Hello", 40.0)];
        let (count, width) = fit_fragments(&frags, 100.0);
        assert_eq!(count, 1);
        assert!((width - 40.0).abs() < 0.01);
    }

    #[test]
    fn fit_fragments_exact_boundary() {
        let frags = [make_text_frag("Hello", 100.0)];
        let (count, _) = fit_fragments(&frags, 100.0);
        assert_eq!(count, 1);
    }

    #[test]
    fn fit_fragments_overflow_at_space() {
        let frags = [
            make_text_frag("Hello", 40.0),
            make_text_frag(" ", 5.0),
            make_text_frag("World", 40.0),
            make_text_frag(" ", 5.0),
            make_text_frag("Overflow", 50.0),
        ];
        let (count, _) = fit_fragments(&frags, 90.0);
        // "Hello"(40) + " "(5) + "World"(40) + " "(5) = 90pt fits exactly.
        // "Overflow"(50) would push to 140pt.
        // Break point is after the second space (index 4).
        assert_eq!(count, 4, "Should include up to the second space break point");
    }

    #[test]
    fn fit_fragments_line_break() {
        let frags = [
            make_text_frag("Before", 30.0),
            Fragment::LineBreak { line_height: 14.0 },
            make_text_frag("After", 30.0),
        ];
        let (count, _) = fit_fragments(&frags, 200.0);
        assert_eq!(count, 2, "Should stop at LineBreak (inclusive)");
    }

    #[test]
    fn fit_fragments_hyphen_break() {
        let frags = [
            make_text_frag("Funktions-", 60.0),
            make_text_frag("kleinspannungs-", 80.0),
            make_text_frag("Stromkreise", 70.0),
        ];
        let (count, width) = fit_fragments(&frags, 150.0);
        assert_eq!(count, 2, "Should break after 'kleinspannungs-'");
        assert!((width - 140.0).abs() < 0.01);
    }

    #[test]
    fn fit_fragments_single_oversized() {
        let frags = [make_text_frag("VeryLongWord", 200.0)];
        let (count, _) = fit_fragments(&frags, 100.0);
        assert_eq!(count, 1, "Must include at least 1 fragment even if oversized");
    }

    // ==============================================================
    // measure_lines
    // ==============================================================

    #[test]
    fn measure_lines_empty_fragments() {
        let measured = measure_lines(&[], 72.0, 468.0, 0.0, None, None, &[], 36.0);
        assert_eq!(measured.lines.len(), 0);
        assert!((measured.total_height).abs() < 0.01);
    }

    #[test]
    fn measure_lines_single_line() {
        let frags = [make_text_frag("Hello", 40.0)];
        let measured = measure_lines(&frags, 72.0, 468.0, 0.0, None, None, &[], 36.0);
        assert_eq!(measured.lines.len(), 1);
        assert!((measured.total_height - 14.0).abs() < 0.01);
        // Should have at least one Text command
        assert!(measured.lines[0].commands.iter().any(|c| matches!(c, DrawCommand::Text { .. })));
    }

    #[test]
    fn measure_lines_wraps_to_two_lines() {
        let frags = [
            make_text_frag("Hello", 250.0),
            make_text_frag(" ", 5.0),
            make_text_frag("World", 250.0),
        ];
        let measured = measure_lines(&frags, 72.0, 468.0, 0.0, None, None, &[], 36.0);
        assert_eq!(measured.lines.len(), 2, "Should wrap to 2 lines");
        assert!((measured.total_height - 28.0).abs() < 0.01, "2 lines × 14pt = {}", measured.total_height);
    }

    #[test]
    fn measure_lines_center_alignment() {
        let frags = [make_text_frag("Short", 40.0)];
        let measured = measure_lines(
            &frags, 72.0, 468.0, 0.0, Some(Alignment::Center), None, &[], 36.0,
        );
        // Text x should be offset to center
        if let Some(DrawCommand::Text { x, .. }) = measured.lines[0].commands.iter()
            .find(|c| matches!(c, DrawCommand::Text { .. })) {
            let expected_x = 72.0 + (468.0 - 40.0) / 2.0;
            assert!((*x - expected_x).abs() < 1.0, "Center x={x}, expected ~{expected_x}");
        }
    }

    #[test]
    fn measure_lines_right_alignment() {
        let frags = [make_text_frag("Short", 40.0)];
        let measured = measure_lines(
            &frags, 72.0, 468.0, 0.0, Some(Alignment::Right), None, &[], 36.0,
        );
        if let Some(DrawCommand::Text { x, .. }) = measured.lines[0].commands.iter()
            .find(|c| matches!(c, DrawCommand::Text { .. })) {
            let expected_x = 72.0 + 468.0 - 40.0;
            assert!((*x - expected_x).abs() < 1.0, "Right x={x}, expected ~{expected_x}");
        }
    }

    #[test]
    fn measure_lines_with_underline() {
        let frags = [Fragment::Text {
            text: "Underlined".to_string(),
            font_family: Rc::from("Helvetica"),
            font_size: 12.0,
            bold: false,
            italic: false,
            underline: true,
            color: None,
            shading: None,
            char_spacing_pt: 0.0,
            measured_width: 60.0,
            measured_height: 14.0,
            hyperlink_url: None,
        }];
        let measured = measure_lines(&frags, 72.0, 468.0, 0.0, None, None, &[], 36.0);
        assert!(measured.lines[0].commands.iter()
            .any(|c| matches!(c, DrawCommand::Underline { .. })),
            "Should have an underline command");
    }

    #[test]
    fn measure_lines_with_shading() {
        let frags = [Fragment::Text {
            text: "Shaded".to_string(),
            font_family: Rc::from("Helvetica"),
            font_size: 12.0,
            bold: false,
            italic: false,
            underline: false,
            color: None,
            shading: Some(Color { r: 255, g: 255, b: 0 }),
            char_spacing_pt: 0.0,
            measured_width: 40.0,
            measured_height: 14.0,
            hyperlink_url: None,
        }];
        let measured = measure_lines(&frags, 72.0, 468.0, 0.0, None, None, &[], 36.0);
        assert!(measured.lines[0].commands.iter()
            .any(|c| matches!(c, DrawCommand::Rect { color: (255, 255, 0), .. })),
            "Should have a shading rect");
    }

    #[test]
    fn measure_lines_first_line_offset() {
        let frags = [make_text_frag("Hello", 40.0)];
        let no_offset = measure_lines(&frags, 72.0, 468.0, 0.0, None, None, &[], 36.0);
        let with_offset = measure_lines(&frags, 72.0, 468.0, 20.0, None, None, &[], 36.0);
        let x_no = no_offset.lines[0].commands.iter()
            .find_map(|c| if let DrawCommand::Text { x, .. } = c { Some(*x) } else { None }).unwrap();
        let x_off = with_offset.lines[0].commands.iter()
            .find_map(|c| if let DrawCommand::Text { x, .. } = c { Some(*x) } else { None }).unwrap();
        assert!((x_off - x_no - 20.0).abs() < 0.01, "First-line offset should shift x by 20pt");
    }

    #[test]
    fn measure_lines_with_line_spacing() {
        let frags = [make_text_frag("Hello", 40.0)];
        let single = measure_lines(&frags, 72.0, 468.0, 0.0, None, None, &[], 36.0);
        let double = measure_lines(
            &frags, 72.0, 468.0, 0.0, None,
            Some(LineSpacing::Multiplier(2.0)), &[], 36.0,
        );
        assert!((single.total_height - 14.0).abs() < 0.01);
        assert!((double.total_height - 28.0).abs() < 0.01);
    }

    // ==============================================================
    // vMerge multi-row distribution
    // ==============================================================

    #[test]
    fn vmerge_three_rows_distributes_height() {
        // A cell spanning 3 rows with tall content should distribute evenly
        let mut restart_cell = make_cell("Tall content that spans three rows");
        restart_cell.vertical_merge = Some(VerticalMerge::Restart);
        let mut continue_cell1 = make_cell("");
        continue_cell1.vertical_merge = Some(VerticalMerge::Continue);
        let mut continue_cell2 = make_cell("");
        continue_cell2.vertical_merge = Some(VerticalMerge::Continue);

        let table = Table {
            rows: vec![
                TableRow { height: None, cells: vec![restart_cell, make_cell("R0C1")] },
                TableRow { height: None, cells: vec![continue_cell1, make_cell("R1C1")] },
                TableRow { height: None, cells: vec![continue_cell2, make_cell("R2C1")] },
            ],
            grid_cols: vec![3000, 3000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let lines = extract_lines(&pages);

        // Should have horizontal borders at 4 y-positions (top of each row + bottom)
        let mut h_ys: Vec<f32> = lines.iter()
            .filter(|(x1, y1, x2, y2)| (y1 - y2).abs() < 0.1 && (x2 - x1).abs() > 10.0)
            .map(|(_, y, _, _)| *y)
            .collect();
        h_ys.sort_by(|a, b| a.partial_cmp(b).unwrap());
        h_ys.dedup_by(|a, b| (*a - *b).abs() < 0.5);
        assert!(h_ys.len() >= 4, "Expected 4 unique y positions for 3 rows, got {:?}", h_ys);
    }

    #[test]
    fn vmerge_multiple_columns() {
        let mut restart_a = make_cell("ColA");
        restart_a.vertical_merge = Some(VerticalMerge::Restart);
        let mut continue_a = make_cell("HiddenA");
        continue_a.vertical_merge = Some(VerticalMerge::Continue);
        let mut restart_c = make_cell("ColC");
        restart_c.vertical_merge = Some(VerticalMerge::Restart);
        let mut continue_c = make_cell("HiddenC");
        continue_c.vertical_merge = Some(VerticalMerge::Continue);

        let table = Table {
            rows: vec![
                TableRow { height: None, cells: vec![restart_a, make_cell("B0"), restart_c] },
                TableRow { height: None, cells: vec![continue_a, make_cell("B1"), continue_c] },
            ],
            grid_cols: vec![2000, 2000, 2000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);

        // Restart cells should render
        assert!(texts.iter().any(|(_, _, t)| t == "ColA"), "ColA should render");
        assert!(texts.iter().any(|(_, _, t)| t == "ColC"), "ColC should render");
        // Continue cells should NOT render
        assert!(!texts.iter().any(|(_, _, t)| t == "HiddenA"), "HiddenA should not render");
        assert!(!texts.iter().any(|(_, _, t)| t == "HiddenC"), "HiddenC should not render");
        // Non-merged cells render normally
        assert!(texts.iter().any(|(_, _, t)| t == "B0"));
        assert!(texts.iter().any(|(_, _, t)| t == "B1"));
    }

    // ==============================================================
    // Spacing resolution
    // ==============================================================

    #[test]
    fn spacing_defaults_applied_when_paragraph_has_none() {
        let doc = Document {
            blocks: vec![simple_paragraph("Test")],
            default_spacing: Spacing {
                before: Some(100), // 5pt
                after: Some(200),  // 10pt
                line: None,
                ..Default::default()
            },
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        let text_y = texts.iter().find(|(_, _, t)| t == "Test").unwrap().1;
        // Text should be below margin_top + before_spacing(5pt) + line_height
        assert!(text_y > 72.0 + 5.0, "y={text_y} should include default before spacing");
    }

    #[test]
    fn direct_spacing_overrides_defaults() {
        let doc = Document {
            blocks: vec![Block::Paragraph(Paragraph {
                properties: ParagraphProperties {
                    spacing: Some(Spacing {
                        before: Some(400), // 20pt
                        after: None,
                        line: None,
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                runs: vec![Inline::TextRun(TextRun {
                    text: "Test".into(),
                    properties: RunProperties::default(),
                    hyperlink_url: None,
                })],
                floats: Vec::new(),
                section_properties: None,
            })],
            default_spacing: Spacing {
                before: Some(100), // 5pt — should be overridden
                after: Some(200),
                line: None,
                ..Default::default()
            },
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        let text_y = texts.iter().find(|(_, _, t)| t == "Test").unwrap().1;
        assert!(text_y > 72.0 + 20.0, "y={text_y} should include direct before=20pt");
    }

    // ==============================================================
    // Page breaks: paragraph across pages
    // ==============================================================

    #[test]
    fn paragraph_shading_split_across_pages() {
        // Fill page nearly to the bottom, then add a shaded paragraph
        let mut blocks = Vec::new();
        for _ in 0..40 {
            blocks.push(simple_paragraph("Filler line"));
        }
        blocks.push(Block::Paragraph(Paragraph {
            properties: ParagraphProperties {
                shading: Some(Color { r: 200, g: 200, b: 200 }),
                ..Default::default()
            },
            runs: vec![Inline::TextRun(TextRun {
                text: "Shaded text that may cross pages".into(),
                properties: RunProperties::default(),
                hyperlink_url: None,
            })],
            floats: Vec::new(),
            section_properties: None,
        }));
        let doc = make_doc(blocks);
        let pages = layout(&doc, &LayoutConfig::default());
        // Should not panic — verify at least one page has rect commands
        let total_rects: usize = pages.iter()
            .map(|p| p.commands.iter().filter(|c| matches!(c, DrawCommand::Rect { .. })).count())
            .sum();
        assert!(total_rects >= 1, "Should have at least one shading rect");
    }

    // ==============================================================
    // Page breaks: table across pages
    // ==============================================================

    #[test]
    fn table_splits_across_pages() {
        let mut rows = Vec::new();
        for i in 0..50 {
            rows.push(TableRow {
                height: Some(400), // 20pt each → 1000pt total > page height
                cells: vec![make_cell(&format!("Row {i}"))],
            });
        }
        let table = Table {
            rows,
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        assert!(pages.len() > 1, "Table should span multiple pages, got {} pages", pages.len());
        // First page should have text content
        assert!(pages[0].commands.iter().any(|c| matches!(c, DrawCommand::Text { .. })));
        // Last page too
        assert!(pages.last().unwrap().commands.iter().any(|c| matches!(c, DrawCommand::Text { .. })));
    }

    // ==============================================================
    // Cell margin resolution
    // ==============================================================

    #[test]
    fn cell_margins_from_table_default() {
        let table = Table {
            rows: vec![TableRow {
                height: None,
                cells: vec![make_cell("Content")],
            }],
            grid_cols: vec![5000],
            default_cell_margins: Some(CellMargins {
                top: 100,   // 5pt
                bottom: 100,
                left: 200,  // 10pt
                right: 200,
            }),
            cell_spacing: None,
            borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        let text = texts.iter().find(|(_, _, t)| t == "Content").unwrap();
        // Text should be offset from cell left edge by left margin (10pt)
        let expected_min_x = 72.0 + 10.0; // margin_left + cell_margin_left
        assert!(text.0 >= expected_min_x - 1.0,
            "Text x={} should be >= {} (cell left margin)", text.0, expected_min_x);
    }

    // ==============================================================
    // Cell shading
    // ==============================================================

    #[test]
    fn cell_shading_produces_rect() {
        let mut cell = make_cell("Shaded");
        cell.shading = Some(Color { r: 200, g: 100, b: 50 });
        let table = Table {
            rows: vec![TableRow { height: None, cells: vec![cell] }],
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let rects = extract_rects(&pages);
        assert!(rects.iter().any(|(_, _, _, _, c)| *c == (200, 100, 50)),
            "Should have a rect with color (200,100,50)");
    }

    // ==============================================================
    // Empty and edge-case tables
    // ==============================================================

    #[test]
    fn empty_table_no_crash() {
        let table = Table {
            rows: vec![],
            grid_cols: vec![],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        assert_eq!(pages.len(), 1);
    }

    #[test]
    fn single_cell_table() {
        let table = Table {
            rows: vec![TableRow {
                height: None,
                cells: vec![make_cell("Only")],
            }],
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        assert!(texts.iter().any(|(_, _, t)| t == "Only"));
    }

    #[test]
    fn table_with_empty_cell() {
        let table = Table {
            rows: vec![TableRow {
                height: None,
                cells: vec![
                    TableCell {
                        blocks: vec![],
                        width: None, grid_span: 1, vertical_merge: None,
                        cell_margins: None, cell_borders: None, shading: None,
                    },
                    make_cell("Filled"),
                ],
            }],
            grid_cols: vec![2500, 2500],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = make_doc(vec![Block::Table(table)]);
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        assert!(texts.iter().any(|(_, _, t)| t == "Filled"));
    }

    // ==============================================================
    // Empty paragraph height
    // ==============================================================

    #[test]
    fn empty_paragraph_still_has_height() {
        let doc = make_doc(vec![
            Block::Paragraph(Paragraph {
                properties: ParagraphProperties::default(),
                runs: vec![],
                floats: Vec::new(),
                section_properties: None,
            }),
            simple_paragraph("After"),
        ]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let texts = extract_texts(&pages);
        let after_y = texts.iter().find(|(_, _, t)| t == "After").unwrap().1;
        // Empty paragraph should push "After" down from margin_top
        assert!(after_y > config.margin_top + 10.0,
            "After y={after_y} should be pushed down by empty paragraph height");
    }

    // ==============================================================
    // Indentation: first-line and right
    // ==============================================================

    #[test]
    fn first_line_indent_shifts_first_line_only() {
        // Use a paragraph with text that wraps to 2 lines
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties {
                indentation: Some(Indentation {
                    left: None,
                    right: None,
                    first_line: Some(720), // 36pt
                }),
                ..Default::default()
            },
            runs: vec![
                Inline::TextRun(TextRun {
                    text: "First".into(),
                    properties: RunProperties::default(),
                    hyperlink_url: None,
                }),
                Inline::LineBreak,
                Inline::TextRun(TextRun {
                    text: "Second".into(),
                    properties: RunProperties::default(),
                    hyperlink_url: None,
                }),
            ],
            floats: Vec::new(),
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let texts = extract_texts(&pages);
        let first_x = texts.iter().find(|(_, _, t)| t == "First").unwrap().0;
        let second_x = texts.iter().find(|(_, _, t)| t == "Second").unwrap().0;
        assert!((first_x - second_x - 36.0).abs() < 1.0,
            "First line should be 36pt right of second: first={first_x}, second={second_x}");
    }

    // ==============================================================
    // List label generation
    // ==============================================================

    #[test]
    fn bullet_list_renders_label() {
        let doc = Document {
            blocks: vec![Block::Paragraph(Paragraph {
                properties: ParagraphProperties {
                    list_ref: Some(ListRef { num_id: 1, level: 0 }),
                    ..Default::default()
                },
                runs: vec![Inline::TextRun(TextRun {
                    text: "Item".into(),
                    properties: RunProperties::default(),
                    hyperlink_url: None,
                })],
                floats: Vec::new(),
                section_properties: None,
            })],
            numbering: {
                let mut map = std::collections::HashMap::new();
                map.insert(1, NumberingDef {
                    levels: vec![NumberingLevel {
                        format: NumberFormat::Bullet("•".to_string()),
                        level_text: "%1".to_string(),
                        start: 1,
                        indent_left: 720,
                        indent_hanging: 360,
                    }],
                });
                map
            },
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        assert!(texts.iter().any(|(_, _, t)| t == "•"),
            "Should render bullet label, got: {:?}", texts);
    }

    #[test]
    fn decimal_list_increments_counter() {
        let mut blocks = Vec::new();
        for i in 1..=3 {
            blocks.push(Block::Paragraph(Paragraph {
                properties: ParagraphProperties {
                    list_ref: Some(ListRef { num_id: 1, level: 0 }),
                    ..Default::default()
                },
                runs: vec![Inline::TextRun(TextRun {
                    text: format!("Item {i}"),
                    properties: RunProperties::default(),
                    hyperlink_url: None,
                })],
                floats: Vec::new(),
                section_properties: None,
            }));
        }
        let doc = Document {
            blocks,
            numbering: {
                let mut map = std::collections::HashMap::new();
                map.insert(1, NumberingDef {
                    levels: vec![NumberingLevel {
                        format: NumberFormat::Decimal,
                        level_text: "%1.".to_string(),
                        start: 1,
                        indent_left: 720,
                        indent_hanging: 360,
                    }],
                });
                map
            },
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        let texts = extract_texts(&pages);
        assert!(texts.iter().any(|(_, _, t)| t == "1."), "Should have label '1.'");
        assert!(texts.iter().any(|(_, _, t)| t == "2."), "Should have label '2.'");
        assert!(texts.iter().any(|(_, _, t)| t == "3."), "Should have label '3.'");
    }

    // ==============================================================
    // Float adjustment
    // ==============================================================

    #[test]
    fn float_adjustment_shifts_text() {
        // Create a paragraph with a float image on the left
        let float_img = FloatingImage {
            rel_id: RelId::from("rId1"),
            width_pt: 100.0,
            height_pt: 100.0,
            offset_x_pt: 0.0,
            offset_y_pt: 0.0,
            data: Rc::new(vec![0u8; 10]), // dummy data
            format_hint: FormatHint::from("png"),
            align_h: Some("left".to_string()),
            align_v: None,
            wrap_side: WrapSide::BothSides,
            pct_pos_h: None,
            pct_pos_v: None,
        };
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![Inline::TextRun(TextRun {
                text: "FloatTest".into(),
                properties: RunProperties::default(),
                hyperlink_url: None,
            })],
            floats: vec![float_img],
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let texts = extract_texts(&pages);
        let text = texts.iter().find(|(_, _, t)| t == "FloatTest").unwrap();
        // Text should be shifted right of the float (100pt + gap)
        assert!(text.0 > config.margin_left + 100.0,
            "Text x={} should be right of float (margin+100={})", text.0, config.margin_left + 100.0);
    }

    // ==============================================================
    // Header/footer rendering
    // ==============================================================

    #[test]
    fn header_renders_on_each_page() {
        let mut blocks = Vec::new();
        for i in 0..60 {
            blocks.push(simple_paragraph(&format!("Line {i}")));
        }
        let doc = Document {
            blocks,
            default_header: Some(HeaderFooter {
                blocks: vec![Block::Paragraph(Paragraph {
                    properties: ParagraphProperties::default(),
                    runs: vec![Inline::TextRun(TextRun {
                        text: "HEADER".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                })],
            }),
            ..Document::default()
        };
        let pages = layout(&doc, &LayoutConfig::default());
        assert!(pages.len() >= 2, "Need multiple pages");
        // Each page should have "HEADER" text
        for (i, page) in pages.iter().enumerate() {
            assert!(page.commands.iter().any(|c| matches!(c, DrawCommand::Text { text, .. } if text == "HEADER")),
                "Page {i} should have HEADER text");
        }
    }

    #[test]
    fn footer_renders_at_bottom() {
        let doc = Document {
            blocks: vec![simple_paragraph("Body")],
            default_footer: Some(HeaderFooter {
                blocks: vec![Block::Paragraph(Paragraph {
                    properties: ParagraphProperties::default(),
                    runs: vec![Inline::TextRun(TextRun {
                        text: "FOOTER".into(),
                        properties: RunProperties::default(),
                        hyperlink_url: None,
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                })],
            }),
            ..Document::default()
        };
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let texts = extract_texts(&pages);
        let footer = texts.iter().find(|(_, _, t)| t == "FOOTER").unwrap();
        let body = texts.iter().find(|(_, _, t)| t == "Body").unwrap();
        assert!(footer.1 > body.1 + 100.0,
            "Footer y={} should be well below body y={}", footer.1, body.1);
    }

    // ==============================================================
    // Unit conversion edge cases
    // ==============================================================

    #[test]
    fn twips_to_pt_zero() {
        assert!((crate::units::twips_to_pt(0) - 0.0).abs() < 0.001);
    }

    #[test]
    fn twips_to_pt_signed_negative() {
        let pt = crate::units::twips_to_pt_signed(-20);
        assert!((pt - (-1.0)).abs() < 0.001, "twips_to_pt_signed(-20) = {pt}");
    }

    #[test]
    fn emu_to_pt_signed_negative() {
        let pt = crate::units::emu_to_pt_signed(-914400);
        assert!((pt - (-72.0)).abs() < 0.1, "emu_to_pt_signed(-914400) = {pt}");
    }

    // ==============================================================
    // After-table spacing uses document defaults
    // ==============================================================

    #[test]
    fn after_table_spacing_uses_doc_default() {
        let table = Table {
            rows: vec![TableRow { height: None, cells: vec![make_cell("T")] }],
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc = Document {
            blocks: vec![Block::Table(table), simple_paragraph("After")],
            default_spacing: Spacing {
                after: Some(200), // 10pt
                ..Default::default()
            },
            ..Document::default()
        };
        let pages_with = layout(&doc, &LayoutConfig::default());

        let table2 = Table {
            rows: vec![TableRow { height: None, cells: vec![make_cell("T")] }],
            grid_cols: vec![5000],
            default_cell_margins: None, cell_spacing: None, borders: None,
        };
        let doc_no_sp = Document {
            blocks: vec![Block::Table(table2), simple_paragraph("After")],
            ..Document::default()
        };
        let pages_without = layout(&doc_no_sp, &LayoutConfig::default());

        let y_with = extract_texts(&pages_with).iter().find(|(_, _, t)| t == "After").unwrap().1;
        let y_without = extract_texts(&pages_without).iter().find(|(_, _, t)| t == "After").unwrap().1;
        assert!(y_with > y_without + 5.0,
            "With after-spacing y={y_with} should be further than without y={y_without}");
    }

    // ==============================================================
    // Percentage-based float positioning (wp14:pctPosVOffset)
    // ==============================================================

    fn extract_images(pages: &[LayoutedPage]) -> Vec<(f32, f32, f32, f32)> {
        let mut imgs = Vec::new();
        for page in pages {
            for cmd in &page.commands {
                if let DrawCommand::Image { x, y, width, height, .. } = cmd {
                    imgs.push((*x, *y, *width, *height));
                }
            }
        }
        imgs
    }

    #[test]
    fn pct_pos_offset_positions_float_by_page_percentage() {
        // pct_pos_v = 10000 → 10% of page height (792pt) = 79.2pt
        // pct_pos_h = 50000 → 50% of page width (612pt) = 306pt
        let float_img = FloatingImage {
            rel_id: RelId::from("rId1"),
            width_pt: 50.0,
            height_pt: 50.0,
            offset_x_pt: 0.0,
            offset_y_pt: 0.0,
            data: Rc::new(vec![0u8; 10]),
            format_hint: FormatHint::from("png"),
            align_h: None,
            align_v: None,
            wrap_side: WrapSide::BothSides,
            pct_pos_h: Some(50000),
            pct_pos_v: Some(10000),
        };
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![Inline::TextRun(TextRun {
                text: "Body".into(),
                properties: RunProperties::default(),
                hyperlink_url: None,
            })],
            floats: vec![float_img],
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let imgs = extract_images(&pages);
        assert!(!imgs.is_empty(), "Should have an image");
        let (ix, iy, _, _) = imgs[0];
        // 50% of 612 = 306
        assert!((ix - 306.0).abs() < 1.0, "Image x={ix}, expected ~306");
        // 10% of 792 = 79.2
        assert!((iy - 79.2).abs() < 1.0, "Image y={iy}, expected ~79.2");
    }

    #[test]
    fn pct_pos_none_uses_regular_offset() {
        let float_img = FloatingImage {
            rel_id: RelId::from("rId1"),
            width_pt: 50.0,
            height_pt: 50.0,
            offset_x_pt: 20.0,
            offset_y_pt: 10.0,
            data: Rc::new(vec![0u8; 10]),
            format_hint: FormatHint::from("png"),
            align_h: None,
            align_v: None,
            wrap_side: WrapSide::BothSides,
            pct_pos_h: None,
            pct_pos_v: None,
        };
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![Inline::TextRun(TextRun {
                text: "Body".into(),
                properties: RunProperties::default(),
                hyperlink_url: None,
            })],
            floats: vec![float_img],
            section_properties: None,
        })]);
        let config = LayoutConfig::default();
        let pages = layout(&doc, &config);
        let imgs = extract_images(&pages);
        let (ix, iy, _, _) = imgs[0];
        // Should use margin_left + offset_x_pt = 72 + 20 = 92
        assert!((ix - 92.0).abs() < 1.0, "Image x={ix}, expected ~92");
    }

    // ==============================================================
    // Hyperlinks produce LinkAnnotation commands
    // ==============================================================

    fn extract_link_annotations(pages: &[LayoutedPage]) -> Vec<(f32, f32, f32, f32, String)> {
        let mut links = Vec::new();
        for page in pages {
            for cmd in &page.commands {
                if let DrawCommand::LinkAnnotation { x, y, width, height, url } = cmd {
                    links.push((*x, *y, *width, *height, url.clone()));
                }
            }
        }
        links
    }

    #[test]
    fn hyperlink_produces_link_annotation() {
        let doc = make_doc(vec![Block::Paragraph(Paragraph {
            properties: ParagraphProperties::default(),
            runs: vec![Inline::TextRun(TextRun {
                text: "Click".into(),
                properties: RunProperties::default(),
                hyperlink_url: Some("https://example.com".to_string()),
            })],
            floats: Vec::new(),
            section_properties: None,
        })]);
        let pages = layout(&doc, &LayoutConfig::default());
        let links = extract_link_annotations(&pages);
        assert_eq!(links.len(), 1, "Should have one link annotation");
        assert_eq!(links[0].4, "https://example.com");
        assert!(links[0].2 > 0.0, "Link width should be > 0");
        assert!(links[0].3 > 0.0, "Link height should be > 0");
    }

    #[test]
    fn no_hyperlink_no_annotation() {
        let doc = make_doc(vec![simple_paragraph("Plain text")]);
        let pages = layout(&doc, &LayoutConfig::default());
        let links = extract_link_annotations(&pages);
        assert!(links.is_empty(), "No hyperlinks, no annotations");
    }
