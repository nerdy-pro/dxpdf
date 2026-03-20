use super::fragment::find_next_tab_stop;
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
                            before: Some(200), // 10pt
                            after: Some(200),  // 10pt
                            line: None,
                        }),
                        ..Default::default()
                    },
                    runs: vec![Inline::TextRun(TextRun {
                        text: "First".into(),
                        properties: RunProperties::default(),
                    })],
                    floats: Vec::new(),
                    section_properties: None,
                }),
                Block::Paragraph(Paragraph {
                    properties: ParagraphProperties {
                        spacing: Some(Spacing {
                            before: Some(100), // 5pt
                            after: None,
                            line: None,
                        }),
                        ..Default::default()
                    },
                    runs: vec![Inline::TextRun(TextRun {
                        text: "Second".into(),
                        properties: RunProperties::default(),
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
                Inline::TextRun(TextRun { text: "Before".into(), properties: RunProperties::default() }),
                Inline::LineBreak,
                Inline::TextRun(TextRun { text: "After".into(), properties: RunProperties::default() }),
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
                spacing: Some(Spacing { before: Some(200), after: Some(200), line: None }),
                shading: Some(Color { r: 200, g: 200, b: 200 }),
                ..Default::default()
            },
            runs: vec![Inline::TextRun(TextRun {
                text: "Shaded".into(),
                properties: RunProperties::default(),
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
