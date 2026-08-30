//! Section layout — sequence blocks vertically into pages.
//!
//! Takes measured blocks (paragraphs with fragments, tables with cells),
//! fits them into pages respecting page size and margins, handles page breaks.
//!
//! # §17.6.22 continuous breaks — one page, two sections
//!
//! Every other kind of section start is also a page start, so "section" and
//! "page" coincide and the two words can be used interchangeably. A
//! `Continuous` break is the case where they come apart: the succeeding section
//! begins partway down a page the preceding one is still filling. Most of the
//! machinery below exists for that one case, and this is the map into it.
//!
//! **When it does not come apart.** A `Continuous` break whose succeeding
//! section changes page size or margins cannot share a page — one physical
//! page cannot have two page sizes or two margin sets — so it is promoted to
//! an ordinary page break before any of the machinery below engages; none of
//! `SectionTail`, `ContinuationState` or the ownership rule below ever sees
//! it. The predicate is `continuous_break_promotes_to_page_break` in
//! `render::mod`, which also has the reasoning and its open edges (issue
//! #112).
//!
//! **Who owns the shared page.** Only one header and footer can be drawn on it,
//! and the *last* section on the page wins — so the page is committed by the
//! succeeding section and falls inside its page range. ECMA-376 §17.6 does not
//! settle this; the rule, and what evidence would revisit it, is stated where it
//! is decided, in `render::shared_page_owner_bounds`. The consequence to hold onto is
//! that the shared page is the succeeding section's *page 0*, which is why
//! `HeaderFooterClearance::for_page(0)` selects the right `first` / `even` /
//! `default` slot there without a special case.
//!
//! **How it is handed over.** `SectionLayout` separates the pages a section
//! commits from what became of its last one (`SectionTail`). A shared page is
//! deliberately **not** in `pages`: it travels inside
//! `SectionTail::SharedWithNext` as a [`ContinuationState`], so "committed by
//! both sections" and "committed by neither" are states a caller cannot
//! construct rather than mistakes it must remember not to make.
//!
//! **What crosses the break.** A page plus a cursor is not enough. The break is
//! not a page boundary, so the per-page state that survives *within* a page has
//! to survive across it too — footnotes reserved before the break and the bottom
//! they already reduced, registered floats, the §17.3.1.9 spacing-collapse
//! trail, a deferred page break. Each field of [`ContinuationState`] records the
//! defect a thinner version of it had.
//!
//! **What deliberately does not.** Columns reset to 0. A continuous break is the
//! usual way to *change* the column count, so an inherited index may not exist
//! in the succeeding section at all; Word additionally balances the preceding
//! columns at such a break, which is not implemented. Both are argued at the
//! reset site in `PageLayoutState::new`.
//!
//! **Why the page may be laid out twice.** The preceding section's content on
//! the shared page was placed against the preceding section's bounds, but the
//! page is drawn with the succeeding section's header and footer — a taller one
//! overlaps that content, a shorter one leaves the band it reserved blank. A
//! cursor adjustment cannot repair content already placed, so
//! `render::fit_shared_page` re-runs the section with the succeeding bounds
//! pinned to that page's index (`FinalPageBounds`). Pinning by index is the
//! point: content that spills off the shared page must not drag the override
//! onto a page that nothing shares.

mod floating_table;
mod helpers;
mod layout;
mod stacker;
mod types;

pub use layout::layout_section;
pub(crate) use layout::layout_section_with_clearance;
pub(crate) use layout::SectionStart;
pub(crate) use layout::{FinalPageBounds, LastPageOwner, SectionLayout, SectionTail};
pub use stacker::{stack_blocks, CellLine, StackResult};
pub use types::*;

// ── Footnote rendering constants ─────────────────────────────────────────────

/// §17.11.23: footnote separator width as a fraction of the content area.
/// Word renders the separator at one-third of the text column width.
const FOOTNOTE_SEPARATOR_RATIO: f32 = 0.33;

/// Thickness of the footnote separator line (pt).
const FOOTNOTE_SEPARATOR_LINE_WIDTH: crate::render::dimension::Pt =
    crate::render::dimension::Pt::new(0.5);

/// Vertical gap between the footnote separator and the first footnote paragraph.
/// Also used as the initial height budget for the separator region (pt).
const FOOTNOTE_SEPARATOR_GAP: crate::render::dimension::Pt = crate::render::dimension::Pt::new(4.0);

// ── Float deduplication ───────────────────────────────────────────────────────

/// Position tolerance for deduplicating floating images (pt).
/// Two float entries within this distance on every axis are treated as identical.
const FLOAT_DEDUP_EPSILON_PT: f32 = 0.1;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::dimension::Pt;
    use crate::render::fonts::Toggle;
    use crate::render::geometry::{PtEdgeInsets, PtSize};
    use crate::render::layout::draw_command::DrawCommand;
    use crate::render::layout::fragment::Fragment;
    use crate::render::layout::fragment::{FontProps, TextMetrics};
    use crate::render::layout::header_footer::HeaderFooterClearance;
    use crate::render::layout::page::PageConfig;
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::table::{
        TableBorderConfig, TableBorderLine, TableBorderStyle, TableCellInput, TableRowInput,
    };
    use crate::render::resolve::color::RgbColor;
    use crate::render::resolve::header_footer::HeaderFooterSet;
    use crate::render::resolve::images::MediaEntry;
    use std::cell::Cell;
    use std::rc::Rc;
    use std::sync::{Mutex, Once};

    struct TableWarningLogger;

    thread_local! {
        static CAPTURE_TABLE_WARNINGS: Cell<bool> = const { Cell::new(false) };
    }

    static TABLE_WARNING_LOGGER: TableWarningLogger = TableWarningLogger;
    static TABLE_WARNING_LOGGER_INIT: Once = Once::new();
    static TABLE_WARNINGS: Mutex<Vec<String>> = Mutex::new(Vec::new());

    impl log::Log for TableWarningLogger {
        fn enabled(&self, metadata: &log::Metadata<'_>) -> bool {
            metadata.level() <= log::Level::Warn
        }

        fn log(&self, record: &log::Record<'_>) {
            if self.enabled(record.metadata())
                && record.target().starts_with("dxpdf::render::layout::table")
                && CAPTURE_TABLE_WARNINGS.with(Cell::get)
            {
                TABLE_WARNINGS
                    .lock()
                    .unwrap()
                    .push(record.args().to_string());
            }
        }

        fn flush(&self) {}
    }

    fn capture_table_warnings<T>(run: impl FnOnce() -> T) -> (T, Vec<String>) {
        TABLE_WARNING_LOGGER_INIT.call_once(|| {
            log::set_logger(&TABLE_WARNING_LOGGER).unwrap();
            log::set_max_level(log::LevelFilter::Warn);
        });
        TABLE_WARNINGS.lock().unwrap().clear();
        CAPTURE_TABLE_WARNINGS.with(|capture| capture.set(true));
        let result = run();
        CAPTURE_TABLE_WARNINGS.with(|capture| capture.set(false));
        let warnings = std::mem::take(&mut *TABLE_WARNINGS.lock().unwrap());
        (result, warnings)
    }

    fn text_frag(text: &str, width: f32, height: f32) -> Fragment {
        Fragment::Text {
            shaped: None,
            level: crate::i18n::bidi::BidiLevel::LTR,
            text: text.into(),
            break_after: crate::render::layout::fragment::fixture_break_after(text),
            font: Rc::new(FontProps {
                rtl: crate::render::fonts::Toggle::Absent,
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: Toggle::Absent,
                italic: Toggle::Absent,
                underline: false,
                char_spacing: Pt::ZERO,
                text_scale: 1.0,
                underline_position: Pt::ZERO,
                underline_thickness: Pt::ZERO,
                strike_lines: 0,
                strike_position: Pt::ZERO,
                strike_thickness: Pt::ZERO,
            }),
            color: RgbColor::BLACK,
            width: Pt::new(width),
            trimmed_width: Pt::new(width),
            metrics: TextMetrics {
                ascent: Pt::new(height * 0.7),
                descent: Pt::new(height * 0.3),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            shading: None,
            border: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
            is_footnote_ref: false,
        }
    }

    fn para_block(text: &str, width: f32) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, width, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    fn floating_image(wrap_mode: WrapMode, height: f32) -> FloatingImage {
        FloatingImage {
            image_data: MediaEntry {
                data: std::sync::Arc::from([]),
                format: crate::model::ImageFormat::Png,
            },
            size: PtSize::new(Pt::new(80.0), Pt::new(height)),
            src_rect: None,
            x: FloatingImageX::Absolute(Pt::new(10.0)),
            y: FloatingImageY::RelativeToParagraph(Pt::ZERO),
            wrap_mode,
            dist_left: Pt::ZERO,
            dist_right: Pt::ZERO,
            behind_doc: false,
        }
    }

    fn absolute_floating_image(wrap_mode: WrapMode, y: f32, height: f32) -> FloatingImage {
        FloatingImage {
            y: FloatingImageY::Absolute(Pt::new(y)),
            ..floating_image(wrap_mode, height)
        }
    }

    fn styled_para_block(text: &str, keep_next: bool, page_break_before: bool) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, 30.0, 14.0)],
            style: ParagraphStyle {
                keep_next,
                ..Default::default()
            },
            page_break_before,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    fn one_column_table(rows: &[(&str, bool)]) -> LayoutBlock {
        LayoutBlock::Table {
            rows: rows
                .iter()
                .map(|(text, header)| TableRowInput {
                    cells: vec![TableCellInput {
                        blocks: vec![para_block(text, 30.0)],
                        margins: PtEdgeInsets::ZERO,
                        grid_span: 1,
                        shading: None,
                        cell_borders: None,
                        vertical_merge: None,
                        vertical_align: crate::render::layout::table::CellVAlign::Top,
                    }],
                    height_rule: None,
                    is_header: Some(*header),
                    cant_split: Some(true),
                    grid_before: 0,
                    border_overrides: None,
                })
                .collect(),
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        }
    }

    fn small_config() -> PageConfig {
        use crate::render::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
            columns: vec![ColumnGeometry {
                x_offset: Pt::ZERO,
                width: Pt::new(180.0),
            }],
        }
    }

    /// A `TopAndBottom` float carrying both parity readings, so it emits an
    /// `Image` command whose x is whatever the page resolved.
    fn mirrored_floating_image(odd: f32, even: f32) -> FloatingImage {
        FloatingImage {
            x: FloatingImageX::PageParity {
                odd: Pt::new(odd),
                even: Pt::new(even),
            },
            ..floating_image(WrapMode::TopAndBottom, 10.0)
        }
    }

    fn para_with_float(text: &str, float: FloatingImage, page_break_before: bool) -> LayoutBlock {
        let LayoutBlock::Paragraph {
            fragments,
            style,
            footnotes,
            floating_shapes,
            ..
        } = para_block(text, 30.0)
        else {
            unreachable!("para_block builds a paragraph")
        };
        LayoutBlock::Paragraph {
            fragments,
            style,
            page_break_before,
            footnotes,
            floating_shapes,
            floating_images: vec![float],
        }
    }

    fn image_xs(page: &crate::render::layout::draw_command::LayoutedPage) -> Vec<f32> {
        page.commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Image { rect, .. } => Some(rect.origin.x.raw()),
                _ => None,
            })
            .collect()
    }

    /// §20.4.3.1: the whole point of deferring x past build. The same anchor
    /// on two consecutive pages must resolve to the two different readings —
    /// which is only observable end to end, because `resolve_anchor_x` can be
    /// perfectly correct while the stacker still ignores what it produced.
    #[test]
    fn a_mirrored_float_resolves_per_page() {
        let blocks = vec![
            para_with_float("p1", mirrored_floating_image(11.0, 99.0), false),
            para_with_float("p2", mirrored_floating_image(11.0, 99.0), true),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2, "the page break splits them");
        assert_eq!(image_xs(&pages[0]), vec![11.0], "page 1 is odd");
        assert_eq!(image_xs(&pages[1]), vec![99.0], "page 2 is even");
    }

    /// The parity follows the *logical* page number, so a section that
    /// `w:pgNumType/@start` begins on an even page mirrors from its first
    /// page — the same key §17.10.6 uses to pick the even header.
    #[test]
    fn parity_follows_the_logical_page_number_not_the_position() {
        let blocks = vec![para_with_float(
            "p1",
            mirrored_floating_image(11.0, 99.0),
            false,
        )];
        let clearance =
            crate::render::layout::header_footer::HeaderFooterClearance::uniform(&small_config());
        let pages = super::layout::layout_section_with_clearance(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            super::layout::SectionStart {
                continuation: None,
                last_page: super::layout::LastPageOwner::Own,
                clearance: &clearance,
                logical_page_base: 2,
            },
        )
        .pages;

        assert_eq!(
            image_xs(&pages[0]),
            vec![99.0],
            "a section starting at logical page 2 starts even"
        );
    }

    #[test]
    fn empty_blocks_produces_one_empty_page() {
        let pages = layout_section(&[], &small_config(), None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages.len(), 1);
        assert!(pages[0].commands.is_empty());
    }

    #[test]
    fn single_paragraph_on_one_page() {
        let blocks = vec![para_block("hello", 30.0)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 1);
        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    // ── §17.3.1.14 keepLines / §17.3.1.44 widow-orphan: across-page splitting ──

    /// A paragraph of `n` words, each ~100pt wide so two never share a line in
    /// the 180pt content column — i.e. exactly `n` single-line rows.
    fn multiline_para(n: usize, keep_lines: bool, widow_control: bool) -> LayoutBlock {
        let fragments = (0..n)
            .map(|i| text_frag(&format!("word{i} "), 100.0, 14.0))
            .collect();
        LayoutBlock::Paragraph {
            fragments,
            style: ParagraphStyle {
                keep_lines,
                widow_control,
                ..Default::default()
            },
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    fn text_count(commands: &[DrawCommand]) -> usize {
        commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count()
    }

    fn two_column_config() -> PageConfig {
        use crate::render::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
            columns: vec![
                ColumnGeometry {
                    x_offset: Pt::ZERO,
                    width: Pt::new(85.0),
                },
                ColumnGeometry {
                    x_offset: Pt::new(95.0),
                    width: Pt::new(85.0),
                },
            ],
        }
    }

    #[test]
    fn paragraph_splits_across_columns() {
        // §17.6.4: two 85pt columns, 80pt tall → 5 lines each. A 12-line
        // paragraph fills column 0, then column 1, then spills to page 2.
        let fragments: Vec<Fragment> = (0..12)
            .map(|i| text_frag(&format!("word{i} "), 60.0, 14.0))
            .collect();
        let block = LayoutBlock::Paragraph {
            fragments,
            style: ParagraphStyle {
                widow_control: false,
                ..Default::default()
            },
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let pages = layout_section(
            &[block],
            &two_column_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert!(pages.len() >= 2, "12 lines exceed two 5-line columns");
        let xs: std::collections::BTreeSet<i32> = pages[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, .. } => Some(position.x.raw() as i32),
                _ => None,
            })
            .collect();
        assert!(xs.contains(&10), "column 0 text at x=10, got {xs:?}");
        assert!(xs.contains(&105), "column 1 text at x=105, got {xs:?}");
    }

    fn border_line() -> crate::render::layout::paragraph::BorderLine {
        crate::render::layout::paragraph::BorderLine {
            width: Pt::new(1.0),
            color: RgbColor::BLACK,
            space: Pt::new(2.0),
        }
    }

    #[test]
    fn tall_paragraph_splits_across_pages() {
        // 80pt content height / 14pt lines → 5 lines per page. Eight lines must
        // span two pages, and widow control keeps >= 2 lines on each.
        let blocks = vec![multiline_para(8, false, true)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2, "eight lines should not fit on one page");
        assert_eq!(
            text_count(&pages[0].commands) + text_count(&pages[1].commands),
            8
        );
        assert!(
            text_count(&pages[0].commands) >= 2 && text_count(&pages[1].commands) >= 2,
            "§17.3.1.44: each side of the split keeps >= 2 lines, got {} + {}",
            text_count(&pages[0].commands),
            text_count(&pages[1].commands)
        );
    }

    #[test]
    fn keep_lines_prevents_split() {
        // Same eight lines, but keepLines set: the paragraph is not split — it
        // stays whole on the first page (overflowing) rather than paginating.
        let blocks = vec![multiline_para(8, true, true)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 1, "§17.3.1.14: keepLines forbids the split");
        assert_eq!(text_count(&pages[0].commands), 8);
    }

    #[test]
    fn widow_control_moves_unsplittable_paragraph_whole() {
        // Page 1 holds three lines of para 1 (42pt), leaving ~38pt (< 3 lines).
        // A 3-line paragraph cannot split without a widow, so it moves whole.
        let blocks = vec![
            multiline_para(3, false, true),
            multiline_para(3, false, true),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2);
        assert_eq!(
            text_count(&pages[0].commands),
            3,
            "only para 1 fits on page 1"
        );
        assert_eq!(
            text_count(&pages[1].commands),
            3,
            "para 2 moved whole to page 2"
        );
    }

    #[test]
    fn keep_next_paragraph_taller_than_page_still_splits() {
        // §17.3.1.15: keepNext must not stop a paragraph taller than a page from
        // splitting (it cannot be "kept with next" on one page regardless). All
        // content is preserved and the following block trails it — no hang.
        let fragments: Vec<Fragment> = (0..8)
            .map(|i| text_frag(&format!("word{i} "), 100.0, 14.0))
            .collect();
        let keep = LayoutBlock::Paragraph {
            fragments,
            style: ParagraphStyle {
                keep_next: true,
                widow_control: false,
                ..Default::default()
            },
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let next = para_block("next", 40.0);
        let pages = layout_section(
            &[keep, next],
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert!(pages.len() >= 2, "the tall keepNext paragraph splits");
        let total: usize = pages.iter().map(|p| text_count(&p.commands)).sum();
        assert_eq!(total, 9, "all eight keepNext lines plus the next paragraph");
    }

    #[test]
    fn bordered_paragraph_splits_across_pages() {
        // §17.3.1.24: a bordered paragraph now splits (Phase C) instead of moving
        // whole — the eight lines span two pages, and each page carries border
        // draw commands (Line) for its segment of the box.
        let mut block = multiline_para(8, false, true);
        if let LayoutBlock::Paragraph { style, .. } = &mut block {
            style.borders = Some(crate::render::layout::paragraph::ParagraphBorderStyle {
                top: Some(border_line()),
                bottom: Some(border_line()),
                left: Some(border_line()),
                right: Some(border_line()),
            });
        }
        let pages = layout_section(
            &[block],
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(
            pages.len(),
            2,
            "bordered paragraph should split, not overflow"
        );
        for page in &pages {
            assert!(
                page.commands
                    .iter()
                    .any(|c| matches!(c, DrawCommand::Line { .. })),
                "each segment draws its part of the border box"
            );
        }
    }

    #[test]
    fn without_widow_control_paragraph_splits_leaving_single_line() {
        // Identical layout, but para 2 disables widow control: it now splits,
        // leaving two lines on page 1 and a single (widow) line on page 2.
        let blocks = vec![
            multiline_para(3, false, true),
            multiline_para(3, false, false),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2);
        assert_eq!(
            text_count(&pages[0].commands),
            5,
            "para 1 (3) + first 2 lines of para 2"
        );
        assert_eq!(text_count(&pages[1].commands), 1, "the widow line");
    }

    /// A paragraph whose first `drop_cap_lines` lines are spanned by a drop cap.
    fn dropcap_para(n: usize, drop_cap_lines: u32) -> LayoutBlock {
        use crate::render::layout::paragraph::DropCapInfo;
        let fragments = (0..n)
            .map(|i| text_frag(&format!("b{i} "), 100.0, 14.0))
            .collect();
        let style = ParagraphStyle {
            // Widow control off so the *natural* break would strand a single
            // line — which, inside the drop-cap prefix, the guard must prevent.
            widow_control: false,
            drop_cap: Some(DropCapInfo {
                fragments: vec![text_frag("§", 20.0, 30.0)],
                lines: drop_cap_lines,
                width: Pt::new(20.0),
                height: Pt::new(30.0),
                ascent: Pt::new(28.0),
                h_space: Pt::ZERO,
                margin_mode: false,
                indent: Pt::ZERO,
                frame_height: None,
                position_offset: Pt::ZERO,
            }),
            ..Default::default()
        };
        LayoutBlock::Paragraph {
            fragments,
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    fn page_has_dropcap_glyph(page: &crate::render::LayoutedPage) -> bool {
        page.commands
            .iter()
            .any(|c| matches!(c, DrawCommand::Text { text, .. } if text.as_ref() == "§"))
    }

    fn footnote_ref_frag() -> Fragment {
        let mut f = text_frag("1", 5.0, 8.0);
        if let Fragment::Text {
            is_footnote_ref, ..
        } = &mut f
        {
            *is_footnote_ref = true;
        }
        f
    }

    fn page_has_text(page: &crate::render::LayoutedPage, needle: &str) -> bool {
        page.commands
            .iter()
            .any(|c| matches!(c, DrawCommand::Text { text, .. } if text.as_ref() == needle))
    }

    #[test]
    fn footnotes_reserve_on_the_page_of_their_reference() {
        // §17.11.12: an eight-line paragraph with a footnote reference on line 0
        // and another on line 7 splits across two pages; each footnote must land
        // on the page carrying its reference, not both on one page.
        let mut fragments = vec![text_frag("word0 ", 100.0, 14.0), footnote_ref_frag()];
        for i in 1..8 {
            fragments.push(text_frag(&format!("word{i} "), 100.0, 14.0));
        }
        fragments.push(footnote_ref_frag()); // reference on the last line
        let footnotes = vec![
            (
                vec![text_frag("FNA", 30.0, 12.0)],
                ParagraphStyle::default(),
            ),
            (
                vec![text_frag("FNB", 30.0, 12.0)],
                ParagraphStyle::default(),
            ),
        ];
        let block = LayoutBlock::Paragraph {
            fragments,
            style: ParagraphStyle {
                widow_control: false,
                ..Default::default()
            },
            page_break_before: false,
            footnotes,
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let pages = layout_section(
            &[block],
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert!(pages.len() >= 2, "the paragraph should split");
        let last = pages.len() - 1;
        assert!(
            page_has_text(&pages[0], "FNA"),
            "footnote A on its reference's page"
        );
        assert!(
            !page_has_text(&pages[0], "FNB"),
            "footnote B is not on page 1"
        );
        assert!(
            page_has_text(&pages[last], "FNB"),
            "footnote B on the last page with its reference"
        );
    }

    #[test]
    fn drop_cap_prefix_is_not_split_across_pages() {
        // §17.3.1.11: page 1 leaves room for ~1 line after four lines of para 1.
        // A naive break would put a single line of the drop-cap paragraph on
        // page 1 and tear the 3-line drop cap. The guard moves the paragraph
        // whole, so page 1 holds only para 1 and the drop cap lands intact later.
        let blocks = vec![multiline_para(4, false, false), dropcap_para(6, 3)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert!(pages.len() >= 2);
        assert_eq!(
            text_count(&pages[0].commands),
            4,
            "only para 1 on page 1 — the drop-cap paragraph moved whole"
        );
        assert!(
            !page_has_dropcap_glyph(&pages[0]),
            "the drop cap must not appear on page 1"
        );
        // The drop cap lands on some later page, together with >= 3 body lines
        // (its prefix), never stranded.
        let cap_page = pages
            .iter()
            .find(|p| page_has_dropcap_glyph(p))
            .expect("drop cap is emitted on some page");
        // Body lines on the drop-cap page = text commands minus the glyph;
        // the 3-line drop-cap prefix stays with the glyph (so > 3 text commands).
        assert!(
            text_count(&cap_page.commands) > 3,
            "the drop-cap prefix (>= 3 lines) stays with the glyph"
        );
    }

    #[test]
    fn text_positioned_at_margins() {
        let blocks = vec![para_block("hello", 30.0)];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        if let Some(DrawCommand::Text { position, .. }) = pages[0].commands.first() {
            assert!(
                position.x.raw() >= config.margins.left.raw(),
                "x should be at least left margin"
            );
            assert!(
                position.y.raw() >= config.margins.top.raw(),
                "y should be at least top margin"
            );
        }
    }

    #[test]
    fn page_break_when_content_overflows() {
        // Page: 100pt tall, margins 10 each → 80pt content area
        // Each paragraph: 14pt tall
        // 6 paragraphs = 84pt > 80pt → should break to 2 pages
        let blocks: Vec<_> = (0..6).map(|i| para_block(&format!("p{i}"), 30.0)).collect();
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2, "should overflow to 2 pages");

        let page1_texts: Vec<_> = pages[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect();
        let page2_texts: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, .. } => Some(text.clone()),
                _ => None,
            })
            .collect();

        assert_eq!(page1_texts.len(), 5, "5 paras fit on page 1 (5*14=70 < 80)");
        assert_eq!(page2_texts.len(), 1, "1 para on page 2");
    }

    #[test]
    fn each_page_uses_its_selected_header_and_footer_clearance() {
        let config = small_config();
        let clearance = HeaderFooterClearance::new(
            &config,
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            true,
            false,
            1,
        );
        let blocks: Vec<_> = (0..5).map(|i| para_block(&format!("p{i}"), 30.0)).collect();

        let pages = layout_section_with_clearance(
            &blocks,
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            super::layout::SectionStart {
                continuation: None,
                last_page: super::layout::LastPageOwner::Own,
                clearance: &clearance,
                logical_page_base: 1,
            },
        )
        .pages;

        assert_eq!(pages.len(), 2, "page 2 must use the shorter default slots");
        let page_texts = pages
            .iter()
            .map(|page| {
                page.commands
                    .iter()
                    .filter_map(|command| match command {
                        DrawCommand::Text { text, position, .. } => {
                            Some((text.to_string(), position.y))
                        }
                        _ => None,
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();

        assert_eq!(
            page_texts[0].len(),
            2,
            "40pt first-page body fits two lines"
        );
        assert_eq!(
            page_texts[1].len(),
            3,
            "80pt default-page body fits the remainder"
        );
        assert!(
            page_texts[1][0].1 < page_texts[0][0].1,
            "page 2 body must start at the shorter default-header boundary",
        );
    }

    #[test]
    fn paragraph_moved_to_taller_page_is_relaid_before_following_block() {
        let config = small_config();
        let clearance = HeaderFooterClearance::new(
            &config,
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            true,
            false,
            1,
        );
        let moved = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("moved-{i} "), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let pages = layout_section_with_clearance(
            &[para_block("before", 30.0), moved, para_block("after", 30.0)],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            super::layout::SectionStart {
                continuation: None,
                last_page: super::layout::LastPageOwner::Own,
                clearance: &clearance,
                logical_page_base: 1,
            },
        )
        .pages;

        assert_eq!(pages.len(), 2);
        let moved_last_y = pages[1]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("moved-") => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .max_by(f32::total_cmp)
            .expect("moved paragraph text");
        let after_y = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.as_ref() == "after" => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .expect("following paragraph text");

        assert!(
            after_y > moved_last_y,
            "following paragraph at {after_y:?} overlaps moved paragraph ending at {moved_last_y:?}",
        );
    }

    #[test]
    fn moved_paragraph_keeps_its_wrap_float_during_destination_relayout() {
        let config = small_config();
        let clearance = HeaderFooterClearance::new(
            &config,
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            true,
            false,
            1,
        );
        let moved = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("moved-{i}"), 80.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section_with_clearance(
            &[para_block("before", 170.0), moved],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            super::layout::SectionStart {
                continuation: None,
                last_page: super::layout::LastPageOwner::Own,
                clearance: &clearance,
                logical_page_base: 1,
            },
        )
        .pages;

        assert_eq!(pages.len(), 2);
        let first_moved_x = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("moved-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("moved paragraph text on page 2");
        assert!(
            first_moved_x >= Pt::new(90.0),
            "moved text at {first_moved_x:?} overlaps its 10-90pt float",
        );
    }

    #[test]
    fn moved_absolute_float_relayouts_earlier_source_page_text() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2);
        assert!(
            !pages[0]
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. })),
            "the relocated float must not remain on the source page"
        );
        assert!(
            pages[1]
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. })),
            "the relocated float must be emitted on the destination page"
        );
        let source_positions: Vec<_> = pages[0]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);

        let owner_x = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("owner-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("owner text on destination page");
        assert!(
            owner_x >= Pt::new(90.0),
            "moved owner text at {owner_x:?} overlaps its 10-90pt float"
        );
    }

    #[test]
    fn leading_inline_page_break_relayouts_source_page_after_float_relocation() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: std::iter::once(Fragment::PageBreak {
                line_height: Pt::new(14.0),
            })
            .chain((0..2).map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0)))
            .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2);
        let source_positions: Vec<_> = pages[0]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);

        let owner_x = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("owner-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("owner text on destination page");
        assert!(
            owner_x >= Pt::new(90.0),
            "owner text at {owner_x:?} overlaps its 10-90pt float"
        );
    }

    #[test]
    fn column_break_page_transition_drops_forward_floats_from_previous_page() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: std::iter::once(Fragment::ColumnBreak)
                .chain((0..4).map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0)))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 3);
        let source_positions: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);
        assert!(
            pages[2]
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. })),
            "the relocated float must be emitted on page 3"
        );
    }

    #[test]
    fn nonempty_column_break_does_not_leak_future_float_to_source_page() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: std::iter::once(text_frag("before", 170.0, 14.0))
                .chain(std::iter::once(Fragment::ColumnBreak))
                .chain((0..4).map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0)))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 3);
        let before_x = pages[0]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.as_ref() == "before" => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("text before the column break on page 1");
        assert_eq!(before_x, Pt::new(10.0));
        let source_positions: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);
        assert!(pages[..2].iter().all(|page| {
            !page
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. }))
        }));
        assert_eq!(
            pages[2]
                .commands
                .iter()
                .filter(|command| matches!(command, DrawCommand::Image { .. }))
                .count(),
            1
        );
        let owner_x = pages[2]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("owner-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("owner text on page 3");
        assert!(owner_x >= Pt::new(90.0));
    }

    #[test]
    fn leading_inline_column_break_relayouts_source_page_after_float_relocation() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: std::iter::once(Fragment::ColumnBreak)
                .chain((0..2).map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0)))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 2);
        let source_positions: Vec<_> = pages[0]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);
        assert!(!pages[0]
            .commands
            .iter()
            .any(|command| matches!(command, DrawCommand::Image { .. })));
        assert_eq!(
            pages[1]
                .commands
                .iter()
                .filter(|command| matches!(command, DrawCommand::Image { .. }))
                .count(),
            1
        );
        let owner_x = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("owner-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("owner text on page 2");
        assert!(owner_x >= Pt::new(90.0));
    }

    #[test]
    fn nonempty_page_break_does_not_leak_future_float_to_source_page() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: std::iter::once(text_frag("before", 170.0, 14.0))
                .chain(std::iter::once(Fragment::PageBreak {
                    line_height: Pt::new(14.0),
                }))
                .chain((0..4).map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0)))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 3);
        let before_x = pages[0]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.as_ref() == "before" => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("text before the page break on page 1");
        assert_eq!(before_x, Pt::new(10.0));
        let source_positions: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);
        assert!(pages[..2].iter().all(|page| {
            !page
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. }))
        }));
        assert_eq!(
            pages[2]
                .commands
                .iter()
                .filter(|command| matches!(command, DrawCommand::Image { .. }))
                .count(),
            1
        );
        let owner_x = pages[2]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("owner-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("owner text on page 3");
        assert!(owner_x >= Pt::new(90.0));
    }

    #[test]
    fn same_page_column_break_preserves_page_float_state() {
        use crate::render::layout::page::ColumnGeometry;

        let mut config = small_config();
        config.columns = vec![
            ColumnGeometry {
                x_offset: Pt::ZERO,
                width: Pt::new(85.0),
            },
            ColumnGeometry {
                x_offset: Pt::new(95.0),
                width: Pt::new(85.0),
            },
        ];
        let source = LayoutBlock::Paragraph {
            fragments: vec![
                text_frag("before", 70.0, 14.0),
                Fragment::ColumnBreak,
                text_frag("after", 70.0, 14.0),
            ],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: vec![text_frag("owner", 70.0, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 1);
        let before_x = pages[0]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.as_ref() == "before" => {
                    Some(position.x)
                }
                _ => None,
            })
            .expect("text in the first column");
        assert!(before_x >= Pt::new(90.0));
        for expected in ["before", "after", "owner"] {
            assert_eq!(
                pages[0]
                    .commands
                    .iter()
                    .filter(|command| {
                        matches!(command, DrawCommand::Text { text, .. } if text.as_ref() == expected)
                    })
                    .count(),
                1
            );
        }
        assert_eq!(
            pages[0]
                .commands
                .iter()
                .filter(|command| matches!(command, DrawCommand::Image { .. }))
                .count(),
            1
        );
    }

    #[test]
    fn moved_absolute_float_replays_from_first_affected_page_two_paragraph() {
        let config = small_config();
        let source = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: true,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[para_block("before", 170.0), source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 3);
        let source_positions: Vec<_> = pages[1]
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.starts_with("source-") => {
                    Some(position.x)
                }
                _ => None,
            })
            .collect();
        assert_eq!(source_positions, vec![Pt::new(10.0); 4]);
        assert!(
            pages[2]
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. })),
            "the relocated float must be emitted on page 3"
        );
    }

    #[test]
    fn moved_absolute_float_does_not_replay_page_start_paragraph() {
        let config = small_config();
        let filler = LayoutBlock::Paragraph {
            fragments: (0..4)
                .map(|i| text_frag(&format!("filler-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let page_start = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("page-start-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let source = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("source-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };
        let owner = LayoutBlock::Paragraph {
            fragments: (0..2)
                .map(|i| text_frag(&format!("owner-{i}"), 170.0, 14.0))
                .collect(),
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[filler, page_start, source, owner],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 3);
        let page_start_commands = pages
            .iter()
            .flat_map(|page| &page.commands)
            .filter(|command| {
                matches!(
                    command,
                    DrawCommand::Text { text, .. } if text.starts_with("page-start-")
                )
            })
            .count();
        assert_eq!(page_start_commands, 2);
    }

    #[test]
    fn moved_paragraph_relocates_top_and_bottom_float_to_destination_page() {
        let config = small_config();
        let clearance = HeaderFooterClearance::new(
            &config,
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            HeaderFooterSet {
                default: None,
                first: Some(Pt::new(30.0)),
                even: None,
            },
            true,
            false,
            1,
        );
        let moved = LayoutBlock::Paragraph {
            fragments: vec![text_frag("moved", 170.0, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![floating_image(WrapMode::TopAndBottom, 30.0)],
            floating_shapes: vec![],
        };

        let pages = layout_section_with_clearance(
            &[para_block("before", 170.0), moved],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            super::layout::SectionStart {
                continuation: None,
                last_page: super::layout::LastPageOwner::Own,
                clearance: &clearance,
                logical_page_base: 1,
            },
        )
        .pages;

        assert_eq!(pages.len(), 2);
        assert!(
            !pages[0]
                .commands
                .iter()
                .any(|command| matches!(command, DrawCommand::Image { .. })),
            "the moved paragraph must not leave its image on page 1",
        );
        let image_bottom = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Image { rect, .. } => Some(rect.origin.y + rect.size.height),
                _ => None,
            })
            .expect("top-and-bottom image on page 2");
        let text_y = pages[1]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Text { text, position, .. } if text.as_ref() == "moved" => {
                    Some(position.y)
                }
                _ => None,
            })
            .expect("moved paragraph text on page 2");
        assert!(
            text_y >= image_bottom,
            "moved text at {text_y:?} must follow image ending at {image_bottom:?}",
        );
    }

    #[test]
    fn footnotes_render_from_the_selected_footer_boundary_once() {
        let config = small_config();
        let clearance = HeaderFooterClearance::new(
            &config,
            HeaderFooterSet::default(),
            HeaderFooterSet {
                default: Some(Pt::new(30.0)),
                first: None,
                even: None,
            },
            false,
            false,
            1,
        );
        let block = LayoutBlock::Paragraph {
            fragments: vec![text_frag("body", 30.0, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![(
                vec![text_frag("footnote", 30.0, 14.0)],
                ParagraphStyle::default(),
            )],
            floating_images: vec![],
            floating_shapes: vec![],
        };

        let pages = layout_section_with_clearance(
            &[block],
            &config,
            None,
            Pt::ZERO,
            Pt::new(14.0),
            super::layout::SectionStart {
                continuation: None,
                last_page: super::layout::LastPageOwner::Own,
                clearance: &clearance,
                logical_page_base: 1,
            },
        )
        .pages;
        let separator_y = pages[0]
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCommand::Line { line, .. } => Some(line.start.y),
                _ => None,
            })
            .expect("footnote separator");

        assert_eq!(separator_y, Pt::new(52.0));
    }

    #[test]
    fn page_size_set_on_layouted_page() {
        let config = small_config();
        let pages = layout_section(&[], &config, None, Pt::ZERO, Pt::new(14.0), None);
        assert_eq!(pages[0].page_size, config.page_size);
    }

    #[test]
    fn many_paragraphs_produce_multiple_pages() {
        // 20 paragraphs at 14pt each = 280pt
        // Content area = 80pt → need 4 pages (80/14 = 5.7 paras per page)
        let blocks: Vec<_> = (0..20)
            .map(|i| para_block(&format!("p{i}"), 30.0))
            .collect();
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 4);
    }

    #[test]
    fn table_on_page() {
        let blocks = vec![LayoutBlock::Table {
            rows: vec![TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: vec![text_frag("cell", 30.0, 14.0)],
                        style: ParagraphStyle::default(),
                        page_break_before: false,
                        footnotes: vec![],
                        floating_images: vec![],
                        floating_shapes: vec![],
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: None,
                    vertical_align: crate::render::layout::table::CellVAlign::Top,
                }],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
            }],
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        }];

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        assert_eq!(pages.len(), 1);

        let text_count = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Text { .. }))
            .count();
        assert_eq!(text_count, 1);
    }

    // ── §17.3.1.33 space_before suppression tests ──────────────────────

    #[test]
    fn space_before_suppressed_for_first_paragraph_of_section() {
        let style = ParagraphStyle {
            space_before: Pt::new(24.0),
            ..Default::default()
        };
        let blocks = vec![LayoutBlock::Paragraph {
            fragments: vec![text_frag("heading", 50.0, 14.0)],
            style,
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        // First paragraph on the section's initial page: space_before suppressed.
        if let Some(DrawCommand::Text { position, .. }) = pages[0].commands.first() {
            assert!(
                position.y.raw() < config.margins.top.raw() + 24.0,
                "space_before should be suppressed: y={}",
                position.y.raw()
            );
        }
    }

    #[test]
    fn space_before_preserved_for_page_break_before() {
        let heading_style = ParagraphStyle {
            space_before: Pt::new(24.0),
            ..Default::default()
        };

        let blocks = vec![
            para_block("first page", 30.0),
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("heading", 50.0, 14.0)],
                style: heading_style,
                page_break_before: true,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            },
        ];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);

        assert!(pages.len() >= 2, "should have at least 2 pages");
        let heading_y = pages[1]
            .commands
            .iter()
            .find_map(|c| match c {
                DrawCommand::Text { position, text, .. } if &**text == "heading" => {
                    Some(position.y)
                }
                _ => None,
            })
            .expect("heading should be on page 2");
        // §17.3.1.33: space_before is preserved — pageBreakBefore paragraphs
        // are not the structural first of the section.
        assert!(
            heading_y.raw() > config.margins.top.raw() + 20.0,
            "space_before should be preserved for pageBreakBefore: y={}",
            heading_y.raw(),
        );
    }

    // ── §17.4.58 — floating-table page overflow ────────────────────────
    //
    // A floating table (`<w:tbl>` with `<w:tblpPr>`) that is taller than
    // the available height on its anchor page must split at row
    // boundaries and continue on the next page. Word draws the
    // continuation at the top of the next page's content area; the
    // `tblpY` anchor applies only to the first slice.

    fn cell_with_text(text: &str) -> TableCellInput {
        TableCellInput {
            blocks: vec![LayoutBlock::Paragraph {
                fragments: vec![text_frag(text, 30.0, 14.0)],
                style: ParagraphStyle::default(),
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            }],
            margins: PtEdgeInsets::ZERO,
            grid_span: 1,
            shading: None,
            cell_borders: None,
            vertical_merge: None,
            vertical_align: crate::render::layout::table::CellVAlign::Top,
        }
    }

    fn row_with_label(label: &str) -> TableRowInput {
        TableRowInput {
            cells: vec![cell_with_text(label)],
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before: 0,
            border_overrides: None,
        }
    }

    /// Build a floating table with `n` rows, anchored to the page at
    /// `y_offset`. Used by overflow tests below.
    fn floating_table_with_rows(n: usize, y_offset: f32) -> LayoutBlock {
        LayoutBlock::Table {
            rows: (0..n).map(|i| row_with_label(&format!("r{i}"))).collect(),
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: Some(super::TableFloatInfo {
                right_gap: Pt::ZERO,
                bottom_gap: Pt::ZERO,
                x_align: None,
                y_offset: Pt::new(y_offset),
                vert_anchor: crate::model::TableAnchor::Page,
                overlap: None,
            }),
            style_id: None,
        }
    }

    /// Collect text strings emitted on each page, in order.
    fn texts_per_page(
        pages: &[crate::render::layout::draw_command::LayoutedPage],
    ) -> Vec<Vec<String>> {
        pages
            .iter()
            .map(|p| {
                p.commands
                    .iter()
                    .filter_map(|c| match c {
                        DrawCommand::Text { text, .. } => Some(text.to_string()),
                        _ => None,
                    })
                    .collect()
            })
            .collect()
    }

    #[test]
    fn keep_next_chain_moves_with_terminal_paragraph() {
        let mut blocks: Vec<_> = (0..3)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        blocks.extend([
            styled_para_block("heading", true, false),
            styled_para_block("subheading", true, false),
            styled_para_block("body", false, false),
        ]);
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        let last = per_page.last().unwrap();
        assert!(last.iter().any(|text| text == "heading"));
        assert!(last.iter().any(|text| text == "subheading"));
        assert!(last.iter().any(|text| text == "body"));
    }

    #[test]
    fn keep_next_fit_includes_group_footnotes() {
        let mut blocks: Vec<_> = (0..3)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        blocks.extend([
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("heading", 30.0, 14.0)],
                style: ParagraphStyle {
                    keep_next: true,
                    ..Default::default()
                },
                page_break_before: false,
                footnotes: vec![(
                    vec![text_frag("footnote", 30.0, 14.0)],
                    ParagraphStyle::default(),
                )],
                floating_images: vec![],
                floating_shapes: vec![],
            },
            styled_para_block("body", false, false),
        ]);

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        let heading_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "heading"))
            .expect("heading text");
        let body_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "body"))
            .expect("body text");

        assert_eq!(
            heading_page, body_page,
            "footnote height must participate in keep-next fit checks",
        );
    }

    #[test]
    fn contextual_keep_next_chain_stays_when_collapsed_spacing_fits() {
        let mut blocks: Vec<_> = (0..2)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        let style = ParagraphStyle {
            keep_next: true,
            contextual_spacing: true,
            style_id: Some(crate::model::StyleId::new("contextual")),
            space_before: Pt::new(10.0),
            space_after: Pt::new(10.0),
            ..Default::default()
        };
        blocks.extend([
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("heading", 30.0, 14.0)],
                style: style.clone(),
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            },
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("body", 30.0, 14.0)],
                style: ParagraphStyle {
                    keep_next: false,
                    ..style
                },
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            },
        ]);

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        assert_eq!(pages.len(), 1);
        let texts = &texts_per_page(&pages)[0];
        assert!(texts.iter().any(|text| text == "heading"));
        assert!(texts.iter().any(|text| text == "body"));
    }

    #[test]
    fn contextual_predecessor_allows_keep_next_table_chain_on_fresh_page() {
        let style = ParagraphStyle {
            keep_next: true,
            contextual_spacing: true,
            style_id: Some(crate::model::StyleId::new("contextual")),
            space_before: Pt::new(20.0),
            space_after: Pt::ZERO,
            ..Default::default()
        };
        let mut blocks = vec![
            para_block("fill", 30.0),
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("predecessor", 30.0, 14.0)],
                style: ParagraphStyle {
                    keep_next: false,
                    ..style.clone()
                },
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            },
        ];
        blocks.extend((0..4).map(|index| LayoutBlock::Paragraph {
            fragments: vec![text_frag(&format!("chain{index}"), 30.0, 14.0)],
            style: style.clone(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }));
        blocks.push(one_column_table(&[("row0", false), ("row1", false)]));

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        let heading_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "chain0"))
            .unwrap();
        let row0_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "row0"))
            .unwrap();

        assert_eq!(heading_page, row0_page);
    }

    #[test]
    fn keep_next_heading_moves_with_first_table_row() {
        for header in [false, true] {
            let mut blocks: Vec<_> = (0..4)
                .map(|i| para_block(&format!("fill{i}"), 30.0))
                .collect();
            blocks.push(styled_para_block("heading", true, false));
            blocks.push(one_column_table(&[
                ("row0", header),
                ("row1", false),
                ("row2", false),
            ]));
            let pages = layout_section(
                &blocks,
                &small_config(),
                None,
                Pt::ZERO,
                Pt::new(14.0),
                None,
            );
            let per_page = texts_per_page(&pages);
            let heading_page = per_page
                .iter()
                .position(|p| p.iter().any(|t| t == "heading"))
                .unwrap();
            let row_page = per_page
                .iter()
                .position(|p| p.iter().any(|t| t == "row0"))
                .unwrap();
            assert_eq!(heading_page, row_page);
        }
    }

    #[test]
    fn keep_next_does_not_move_heading_before_floating_table() {
        let mut blocks = (0..4)
            .map(|i| para_block(&format!("filler-{i}"), 30.0))
            .collect::<Vec<_>>();
        blocks.push(styled_para_block("heading", true, false));
        blocks.push(floating_table_with_rows(1, 15.0));

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let heading_page = pages
            .iter()
            .position(|page| {
                page.commands.iter().any(
                    |command| matches!(command, DrawCommand::Text { text, .. } if text.as_ref() == "heading"),
                )
            })
            .expect("heading text");

        assert_eq!(
            heading_page, 0,
            "a floating table must not act as the terminal block of a keep-next chain",
        );
    }

    #[test]
    fn keep_next_heading_moves_with_bordered_first_table_row() {
        let mut blocks: Vec<_> = (0..3)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        blocks.push(styled_para_block("heading", true, false));
        let mut table = one_column_table(&[
            ("row0", false),
            ("row1", false),
            ("row2", false),
            ("row3", false),
        ]);
        if let LayoutBlock::Table { border_config, .. } = &mut table {
            *border_config = Some(TableBorderConfig {
                top: None,
                bottom: None,
                left: None,
                right: None,
                inside_h: Some(TableBorderLine {
                    width: Pt::new(12.0),
                    color: RgbColor::BLACK,
                    style: TableBorderStyle::Single,
                }),
                inside_v: None,
            });
        }
        blocks.push(table);

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        let heading_page = per_page
            .iter()
            .position(|p| p.iter().any(|text| text == "heading"))
            .unwrap();
        let row0_page = per_page
            .iter()
            .position(|p| p.iter().any(|text| text == "row0"))
            .unwrap();

        assert_eq!(heading_page, row0_page);
        assert!(per_page.iter().skip(row0_page + 1).any(|page| page
            .iter()
            .any(|text| text == "row1" || text == "row2" || text == "row3")));
    }

    #[test]
    fn oversized_splittable_first_table_row_does_not_move_heading() {
        let mut blocks: Vec<_> = (0..4)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        blocks.push(styled_para_block("heading", true, false));
        blocks.push(LayoutBlock::Table {
            rows: vec![TableRowInput {
                cells: vec![TableCellInput {
                    blocks: vec![LayoutBlock::Paragraph {
                        fragments: (0..12)
                            .flat_map(|i| {
                                [
                                    text_frag(&format!("row0-{i}"), 30.0, 14.0),
                                    Fragment::LineBreak {
                                        line_height: Pt::new(14.0),
                                    },
                                ]
                            })
                            .collect(),
                        style: ParagraphStyle::default(),
                        page_break_before: false,
                        footnotes: vec![],
                        floating_images: vec![],
                        floating_shapes: vec![],
                    }],
                    margins: PtEdgeInsets::ZERO,
                    grid_span: 1,
                    shading: None,
                    cell_borders: None,
                    vertical_merge: None,
                    vertical_align: crate::render::layout::table::CellVAlign::Top,
                }],
                height_rule: None,
                is_header: None,
                cant_split: None,
                grid_before: 0,
                border_overrides: None,
            }],
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        });

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        let heading_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "heading"))
            .unwrap();
        let first_row_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "row0-0"))
            .unwrap();

        assert!(heading_page < first_row_page);
        let flattened: Vec<_> = per_page.into_iter().flatten().collect();
        for index in 0..12 {
            let needle = format!("row0-{index}");
            assert_eq!(flattened.iter().filter(|text| *text == &needle).count(), 1);
        }
    }

    #[test]
    fn keep_next_heading_stays_before_oversized_leading_vmerge_group() {
        let mut blocks: Vec<_> = (0..4)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        blocks.push(styled_para_block("heading", true, false));
        blocks.push(LayoutBlock::Table {
            rows: vec![
                TableRowInput {
                    cells: vec![
                        TableCellInput {
                            blocks: (0..6)
                                .map(|i| para_block(&format!("restart-{i}"), 30.0))
                                .collect(),
                            margins: PtEdgeInsets::ZERO,
                            grid_span: 1,
                            shading: None,
                            cell_borders: None,
                            vertical_merge: Some(
                                crate::render::layout::table::VerticalMergeState::Restart,
                            ),
                            vertical_align: crate::render::layout::table::CellVAlign::Top,
                        },
                        TableCellInput {
                            blocks: vec![para_block("row0-peer", 30.0)],
                            margins: PtEdgeInsets::ZERO,
                            grid_span: 1,
                            shading: None,
                            cell_borders: None,
                            vertical_merge: None,
                            vertical_align: crate::render::layout::table::CellVAlign::Top,
                        },
                    ],
                    height_rule: None,
                    is_header: None,
                    cant_split: None,
                    grid_before: 0,
                    border_overrides: None,
                },
                TableRowInput {
                    cells: vec![
                        TableCellInput {
                            blocks: vec![],
                            margins: PtEdgeInsets::ZERO,
                            grid_span: 1,
                            shading: None,
                            cell_borders: None,
                            vertical_merge: Some(
                                crate::render::layout::table::VerticalMergeState::Continue,
                            ),
                            vertical_align: crate::render::layout::table::CellVAlign::Top,
                        },
                        TableCellInput {
                            blocks: (0..2)
                                .map(|i| para_block(&format!("continuation-{i}"), 30.0))
                                .collect(),
                            margins: PtEdgeInsets::ZERO,
                            grid_span: 1,
                            shading: None,
                            cell_borders: None,
                            vertical_merge: None,
                            vertical_align: crate::render::layout::table::CellVAlign::Top,
                        },
                    ],
                    height_rule: None,
                    is_header: None,
                    cant_split: None,
                    grid_before: 0,
                    border_overrides: None,
                },
                TableRowInput {
                    cells: vec![
                        TableCellInput {
                            blocks: vec![para_block("after-merged", 30.0)],
                            margins: PtEdgeInsets::ZERO,
                            grid_span: 1,
                            shading: None,
                            cell_borders: None,
                            vertical_merge: None,
                            vertical_align: crate::render::layout::table::CellVAlign::Top,
                        },
                        TableCellInput {
                            blocks: vec![para_block("after-peer", 30.0)],
                            margins: PtEdgeInsets::ZERO,
                            grid_span: 1,
                            shading: None,
                            cell_borders: None,
                            vertical_merge: None,
                            vertical_align: crate::render::layout::table::CellVAlign::Top,
                        },
                    ],
                    height_rule: None,
                    is_header: None,
                    cant_split: None,
                    grid_before: 0,
                    border_overrides: None,
                },
            ],
            col_widths: vec![Pt::new(50.0), Pt::new(50.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        });

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        let heading_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "heading"))
            .unwrap();
        let restart_page = per_page
            .iter()
            .position(|page| page.iter().any(|text| text == "restart-0"))
            .unwrap();

        assert_eq!(heading_page, 0);
        assert!(heading_page < restart_page);
        let flattened: Vec<_> = per_page.into_iter().flatten().collect();
        for text in [
            "restart-0",
            "restart-1",
            "restart-2",
            "restart-3",
            "restart-4",
            "restart-5",
            "row0-peer",
            "continuation-0",
            "continuation-1",
            "after-merged",
            "after-peer",
        ] {
            assert_eq!(
                flattened
                    .iter()
                    .filter(|candidate| *candidate == text)
                    .count(),
                1
            );
        }
    }

    #[test]
    fn large_table_keep_next_preflight_does_not_probe_later_groups() {
        let mut blocks: Vec<_> = (0..4)
            .map(|i| para_block(&format!("fill{i}"), 30.0))
            .collect();
        blocks.push(styled_para_block("heading", true, false));
        blocks.push(LayoutBlock::Table {
            rows: (0..24)
                .map(|i| row_with_label(&format!("row{i}")))
                .collect(),
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: None,
            style_id: None,
        });

        let (pages, warnings) = capture_table_warnings(|| {
            layout_section(
                &blocks,
                &small_config(),
                None,
                Pt::ZERO,
                Pt::new(14.0),
                None,
            )
        });

        assert!(
            warnings.is_empty(),
            "unexpected table warnings: {warnings:?}"
        );
        let flattened: Vec<_> = texts_per_page(&pages).into_iter().flatten().collect();
        assert_eq!(
            flattened.iter().filter(|text| *text == "heading").count(),
            1
        );
        for index in 0..24 {
            let needle = format!("row{index}");
            assert_eq!(flattened.iter().filter(|text| *text == &needle).count(), 1);
        }
    }

    #[test]
    fn keep_next_does_not_force_an_entire_multi_page_table() {
        let blocks = vec![
            styled_para_block("heading", true, false),
            one_column_table(&[
                ("row0", true),
                ("row1", false),
                ("row2", false),
                ("row3", false),
                ("row4", false),
                ("row5", false),
            ]),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        assert!(pages.len() > 1);
        assert!(texts_per_page(&pages)[0]
            .iter()
            .any(|text| text == "heading"));
    }

    #[test]
    fn explicit_page_break_terminates_keep_next_group() {
        let blocks = vec![
            styled_para_block("heading", true, false),
            styled_para_block("forced", false, true),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        assert!(per_page[0].iter().any(|text| text == "heading"));
        assert!(per_page[1].iter().any(|text| text == "forced"));
    }

    #[test]
    fn oversized_keep_next_chain_makes_progress_without_duplicates() {
        let mut blocks = Vec::new();
        for index in 0..6 {
            blocks.push(styled_para_block(
                &format!("chain{index}"),
                index < 5,
                false,
            ));
        }
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let flattened: Vec<_> = texts_per_page(&pages).into_iter().flatten().collect();
        for index in 0..6 {
            let needle = format!("chain{index}");
            assert_eq!(flattened.iter().filter(|text| *text == &needle).count(), 1);
        }
    }

    // ── §17.3.1.15 keepNext tightening (plan §11.2) ──────────────────────

    /// A `keep_next` paragraph of `n` single-line rows (`{prefix}0..{prefix}{n}`),
    /// each word wide enough that only one fits per 180pt column line.
    fn multiline_keep_next(prefix: &str, n: usize) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: (0..n)
                .map(|i| text_frag(&format!("{prefix}{i}"), 100.0, 14.0))
                .collect(),
            style: ParagraphStyle {
                keep_next: true,
                ..Default::default()
            },
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    /// A 180×140pt content area (10 lines at 14pt), for multi-paragraph chains
    /// that must still fit on a fresh page.
    fn tall_config() -> PageConfig {
        use crate::render::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(160.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
            columns: vec![ColumnGeometry {
                x_offset: Pt::ZERO,
                width: Pt::new(180.0),
            }],
        }
    }

    #[test]
    fn splittable_leading_keep_next_paragraph_fills_current_page() {
        // Two fill lines, then a 4-line keepNext paragraph + 1-line terminal
        // (a 5-line group that fits a fresh page but not the 3 lines left here).
        // The old whole-group move left this page with only the fill; the
        // tightening splits the keepNext paragraph so its head fills the page,
        // and its tail travels to the fresh page WITH the terminal (§17.3.1.15).
        let mut blocks = vec![para_block("fill0", 100.0), para_block("fill1", 100.0)];
        blocks.push(multiline_keep_next("kn", 4));
        blocks.push(para_block("body", 100.0));

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        assert!(pages.len() >= 2);

        // Page 1 is filled with the keepNext paragraph's head, not left blank.
        assert!(
            per_page[0].iter().any(|t| t == "kn0"),
            "leading keepNext paragraph should fill the current page: {per_page:?}"
        );
        // The keepNext boundary holds: the paragraph's last line shares a page
        // with the terminal block.
        let kn_last = per_page
            .iter()
            .position(|p| p.iter().any(|t| t == "kn3"))
            .expect("kn3 present");
        let body = per_page
            .iter()
            .position(|p| p.iter().any(|t| t == "body"))
            .expect("body present");
        assert_eq!(
            kn_last, body,
            "keepNext binds the last line to the next block"
        );
    }

    #[test]
    fn keep_next_chain_leading_split_keeps_all_boundaries() {
        // A multi-paragraph chain (leading kn 6 lines → middle kn 2 lines →
        // terminal) that fits a fresh 10-line page but not the 5 lines left
        // after the fill. The splittable leading paragraph fills the current
        // page; the remainder is strictly smaller than the fresh-page-fitting
        // group, so it lands intact on the next page with every keepNext
        // boundary preserved.
        let mut blocks: Vec<LayoutBlock> = (0..5)
            .map(|i| para_block(&format!("fill{i}"), 100.0))
            .collect();
        blocks.push(multiline_keep_next("lead", 6));
        blocks.push(multiline_keep_next("mid", 2));
        blocks.push(para_block("term", 100.0));

        let pages = layout_section(&blocks, &tall_config(), None, Pt::ZERO, Pt::new(14.0), None);
        let per_page = texts_per_page(&pages);
        assert!(pages.len() >= 2);

        assert!(
            per_page[0].iter().any(|t| t == "lead0"),
            "leading paragraph head fills the current page: {per_page:?}"
        );
        // lead.last ↔ mid.first and mid.last ↔ term must each share a page.
        let find = |needle: &str| {
            per_page
                .iter()
                .position(|p| p.iter().any(|t| t == needle))
                .unwrap_or_else(|| panic!("{needle} present: {per_page:?}"))
        };
        assert_eq!(find("lead5"), find("mid0"), "lead→mid boundary intact");
        assert_eq!(find("mid1"), find("term"), "mid→term boundary intact");
    }

    #[test]
    fn three_line_keep_next_paragraph_moves_whole_under_widow_control() {
        // A 3-line keepNext paragraph cannot split (widow control needs >= 2
        // lines on each side), so the tightening must NOT peel it — the whole
        // group still moves to the fresh page.
        let mut blocks = vec![para_block("fill0", 100.0), para_block("fill1", 100.0)];
        blocks.push(multiline_keep_next("kn", 3));
        blocks.push(para_block("body", 100.0));

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        assert!(pages.len() >= 2);
        assert!(
            !per_page[0].iter().any(|t| t == "kn0"),
            "an unsplittable 3-line keepNext paragraph must move whole: {per_page:?}"
        );
        let kn_last = per_page
            .iter()
            .position(|p| p.iter().any(|t| t == "kn2"))
            .unwrap();
        let body = per_page
            .iter()
            .position(|p| p.iter().any(|t| t == "body"))
            .unwrap();
        assert_eq!(kn_last, body);
    }

    /// A floating table whose laid-out height exceeds the available
    /// space on its anchor page must split at row boundaries. With the
    /// 100pt-tall small_config and an anchor at y=15, 8 rows of ~14pt
    /// (~112pt total) cannot fit in the 85pt below the anchor — at
    /// least one row must spill to page 2.
    #[test]
    fn floating_table_splits_when_taller_than_anchor_page() {
        let blocks = vec![floating_table_with_rows(8, 15.0)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        assert!(
            pages.len() >= 2,
            "overflowing floating table must paginate: got {} pages",
            pages.len()
        );

        // The 8 row labels must appear collectively across pages.
        let per_page = texts_per_page(&pages);
        let all: Vec<&String> = per_page.iter().flatten().collect();
        for i in 0..8 {
            let needle = format!("r{i}");
            assert!(
                all.iter().any(|t| t.as_str() == needle),
                "row {needle} missing from output entirely",
            );
        }
    }

    /// Rows in a paginated floating table do not duplicate: each row
    /// appears on exactly one page. (This is what fails in the bug —
    /// the current code emits every row on the anchor page, then
    /// later content also lands on the anchor page, producing visual
    /// overdraw.)
    #[test]
    fn floating_table_split_rows_do_not_overlap_across_pages() {
        let blocks = vec![floating_table_with_rows(8, 15.0)];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        let per_page = texts_per_page(&pages);
        for i in 0..8 {
            let needle = format!("r{i}");
            let on_pages: Vec<usize> = per_page
                .iter()
                .enumerate()
                .filter_map(|(idx, page)| {
                    if page.iter().any(|t| t == &needle) {
                        Some(idx)
                    } else {
                        None
                    }
                })
                .collect();
            assert_eq!(
                on_pages.len(),
                1,
                "row {needle} appeared on pages {on_pages:?}; floating-table rows must not be duplicated",
            );
        }
    }

    /// Build a floating table with `n` rows, anchored to the page at
    /// `y_offset`, with an explicit `tblOverlap` value.
    fn floating_table_with_rows_and_overlap(
        n: usize,
        y_offset: f32,
        overlap: Option<crate::model::TableOverlap>,
    ) -> LayoutBlock {
        LayoutBlock::Table {
            rows: (0..n).map(|i| row_with_label(&format!("r{i}"))).collect(),
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::ZERO,
            border_config: None,
            indent: Pt::ZERO,
            alignment: None,
            float_info: Some(super::TableFloatInfo {
                right_gap: Pt::ZERO,
                bottom_gap: Pt::ZERO,
                x_align: None,
                y_offset: Pt::new(y_offset),
                vert_anchor: crate::model::TableAnchor::Page,
                overlap,
            }),
            style_id: None,
        }
    }

    /// §17.4.57: two floating tables both declaring `tblOverlap=Never`
    /// must not draw at overlapping y-positions on the same page. The
    /// second table either slides down past the first or spills to
    /// the next page.
    ///
    /// Setup: page 100pt tall, both tables anchored to the same y=15.
    /// First table (3 rows) fits at y=15; second table (3 rows) with
    /// Never must NOT draw on top of the first.
    #[test]
    fn floating_tables_never_overlap_when_both_set_never() {
        use crate::model::TableOverlap;
        let blocks = vec![
            floating_table_with_rows_and_overlap(3, 15.0, Some(TableOverlap::Never)),
            floating_table_with_rows_and_overlap(3, 15.0, Some(TableOverlap::Never)),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        // Collect (page_idx, x, y, text) for all rows.
        let mut row_positions: Vec<(usize, f32, f32, String)> = Vec::new();
        for (pi, page) in pages.iter().enumerate() {
            for cmd in &page.commands {
                if let DrawCommand::Text { text, position, .. } = cmd {
                    if text.starts_with('r') {
                        row_positions.push((
                            pi,
                            position.x.raw(),
                            position.y.raw(),
                            text.to_string(),
                        ));
                    }
                }
            }
        }
        // Group by page+x bucket; check no two row texts share the same y.
        // (The bug emits both tables at y=15, so r0/r1/r2 from each table
        // would land on identical y values.)
        for i in 0..row_positions.len() {
            for j in (i + 1)..row_positions.len() {
                let (pi, xi, yi, ti) = &row_positions[i];
                let (pj, xj, yj, tj) = &row_positions[j];
                if pi == pj && (xi - xj).abs() < 1.0 && (yi - yj).abs() < 0.5 {
                    panic!(
                        "two row labels at same position on page {pi}: {ti:?} and {tj:?} both at y={yi}",
                    );
                }
            }
        }
    }

    /// §17.4.57 governs table-vs-table overlap **only**. A floating
    /// *image* across the anchor must not push a `tblOverlap=Never`
    /// table down: the spec defines the setting as "whether the current
    /// table shall allow other floating **tables** to overlap its
    /// extents", and Word lets a table sit over an image.
    ///
    /// Self-calibrating: the same document is laid out twice, differing
    /// only in whether the paragraph carries the image float. The table
    /// must land in the identical place both times. Asserting equality
    /// rather than a hard-coded y keeps the test honest if page metrics
    /// or paragraph height ever change.
    #[test]
    fn never_overlap_table_is_not_shifted_by_a_floating_image() {
        use crate::model::TableOverlap;

        // The image spans y = 10..52, straight across the table's y=15
        // anchor — so a resolver that counted it would shift to y=52.
        let owner_with_image = LayoutBlock::Paragraph {
            fragments: vec![text_frag("owner", 30.0, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![absolute_floating_image(
                WrapMode::Square(crate::model::WrapText::BothSides),
                10.0,
                42.0,
            )],
            floating_shapes: vec![],
        };
        let owner_without_image = LayoutBlock::Paragraph {
            fragments: vec![text_frag("owner", 30.0, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };

        // (page index, y) of the first table row.
        let first_row_at = |owner: LayoutBlock| -> (usize, f32) {
            let blocks = vec![
                owner,
                floating_table_with_rows_and_overlap(3, 15.0, Some(TableOverlap::Never)),
            ];
            let pages = layout_section(
                &blocks,
                &small_config(),
                None,
                Pt::ZERO,
                Pt::new(14.0),
                None,
            );
            pages
                .iter()
                .enumerate()
                .find_map(|(pi, page)| {
                    page.commands.iter().find_map(|cmd| match cmd {
                        DrawCommand::Text { text, position, .. } if &**text == "r0" => {
                            Some((pi, position.y.raw()))
                        }
                        _ => None,
                    })
                })
                .expect("first table row is drawn")
        };

        let (page_with, y_with) = first_row_at(owner_with_image);
        let (page_without, y_without) = first_row_at(owner_without_image);

        assert_eq!(
            page_with, page_without,
            "the image must not push the table onto another page"
        );
        assert!(
            (y_with - y_without).abs() < 0.01,
            "table moved from y={y_without} to y={y_with} because of an image float; \
             tblOverlap does not govern images"
        );
    }

    /// §17.4.57 default behavior (overlap omitted) — overlap is
    /// permitted. Two tables at the same anchor on the same page
    /// DO draw at overlapping y-positions. This is intentional.
    #[test]
    fn floating_tables_overlap_by_default() {
        let blocks = vec![
            floating_table_with_rows_and_overlap(3, 15.0, None),
            floating_table_with_rows_and_overlap(3, 15.0, None),
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        // Both tables drawn on page 1; the test verifies that we
        // don't accidentally shift when overlap is permitted.
        assert!(!pages.is_empty());
        let page1_rows: Vec<f32> = pages[0]
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { text, position, .. } if text.starts_with('r') => {
                    Some(position.y.raw())
                }
                _ => None,
            })
            .collect();
        assert_eq!(
            page1_rows.len(),
            6,
            "with overlap permitted, both tables draw on page 1: got {} row entries",
            page1_rows.len()
        );
    }

    /// The continuation slice starts at the top content area on page 2
    /// (margin.top), not at the original `tblpY` anchor. §17.4.58 only
    /// anchors the first slice; the continuation flows at the top of
    /// each subsequent page. Behavioral check: the first text on page
    /// 2 must sit *higher* on the page than the first text on page 1
    /// (lower y), because page 1 starts at the anchor (>= margin.top)
    /// and page 2 starts at margin.top exactly.
    #[test]
    fn floating_table_continuation_anchors_at_top_margin() {
        let blocks = vec![floating_table_with_rows(8, 15.0)];
        let config = small_config();
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);
        assert!(pages.len() >= 2, "test precondition: expected overflow");

        let first_text_y =
            |page: &crate::render::layout::draw_command::LayoutedPage| -> Option<f32> {
                page.commands.iter().find_map(|c| match c {
                    DrawCommand::Text { position, .. } => Some(position.y.raw()),
                    _ => None,
                })
            };
        let y_p1 = first_text_y(&pages[0]).expect("page 1 must contain table content");
        let y_p2 = first_text_y(&pages[1]).expect("page 2 must contain table content");
        let anchor_y = 15.0;
        let margin_top = config.margins.top.raw();
        // Anchor sits at y=15 (> margin.top=10), so the page-1 baseline
        // is offset by 5pt versus page 2's continuation baseline.
        assert!(
            y_p2 < y_p1,
            "continuation (page 2 first y={y_p2}) must start above the anchor (page 1 first y={y_p1}); anchor={anchor_y}, margin.top={margin_top}",
        );
        // Sanity: page 2 first text is within the content area.
        assert!(y_p2 >= margin_top - 1.0);
    }

    // ── §17.4.28 + §17.4.44 — a table is aligned on its *outer* width ──
    //
    // `build/table.rs::reserve_cell_spacing` carves one `tblCellSpacing` out
    // of the grid, so `col_widths` is the table's own width *minus* the
    // spacing; `measure.rs` puts the spacing back to get the outer width the
    // table actually occupies. Aligning on the slot sum alone therefore lands
    // a centred table half a spacing off and a right-aligned one a whole
    // spacing off — invisible in the corpus, where no table sets both `w:jc`
    // and `w:tblCellSpacing`, and stated here so it stays fixed.
    //
    // Asserted through geometry rather than arithmetic: centred means equal
    // gaps to the two content edges, right-aligned means the table's right
    // edge meets the content edge.

    /// A one-cell table that draws its issue-#168 outer outline, so the
    /// table's outer edges are readable off the page.
    fn aligned_spaced_table(
        alignment: Option<crate::model::Alignment>,
        cell_spacing: f32,
        float_info: Option<super::TableFloatInfo>,
    ) -> LayoutBlock {
        let line = TableBorderLine {
            width: Pt::new(1.0),
            color: RgbColor::BLACK,
            style: TableBorderStyle::Single,
        };
        LayoutBlock::Table {
            rows: vec![row_with_label("x")],
            col_widths: vec![Pt::new(100.0)],
            cell_spacing: Pt::new(cell_spacing),
            border_config: Some(TableBorderConfig {
                top: Some(line),
                bottom: Some(line),
                left: Some(line),
                right: Some(line),
                inside_h: None,
                inside_v: None,
            }),
            indent: Pt::ZERO,
            alignment,
            float_info,
            style_id: None,
        }
    }

    /// The left and right edges of everything the table painted. The issue-#168
    /// outline spans the full outer width and every cell border sits inside it,
    /// so this bounding span *is* the table's outer span.
    fn painted_x_span(page: &crate::render::layout::draw_command::LayoutedPage) -> (f32, f32) {
        let rects: Vec<_> = page
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Rect { rect, .. } => {
                    Some((rect.origin.x.raw(), (rect.origin.x + rect.size.width).raw()))
                }
                _ => None,
            })
            .collect();
        assert!(!rects.is_empty(), "table must paint its borders");
        (
            rects.iter().map(|r| r.0).fold(f32::INFINITY, f32::min),
            rects.iter().map(|r| r.1).fold(f32::NEG_INFINITY, f32::max),
        )
    }

    #[test]
    fn a_centered_spaced_table_sits_equidistant_from_both_content_edges() {
        let config = small_config();
        let blocks = vec![aligned_spaced_table(
            Some(crate::model::Alignment::Center),
            20.0,
            None,
        )];
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);
        let (left, right) = painted_x_span(&pages[0]);
        let content_left = config.margins.left.raw();
        let content_right = content_left + config.content_width().raw();
        assert!(
            (left - content_left - (content_right - right)).abs() < 0.01,
            "centred: left gap {} must equal right gap {} (table spans {left}..{right})",
            left - content_left,
            content_right - right,
        );
    }

    #[test]
    fn a_right_aligned_spaced_table_ends_at_the_content_edge() {
        let config = small_config();
        let blocks = vec![aligned_spaced_table(
            Some(crate::model::Alignment::End),
            20.0,
            None,
        )];
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);
        let (_, right) = painted_x_span(&pages[0]);
        let content_right = (config.margins.left + config.content_width()).raw();
        assert!(
            (right - content_right).abs() < 0.01,
            "right-aligned: table right edge {right} must meet the content edge {content_right}",
        );
    }

    /// The floating branch (§17.4.58) reads the width off the laid-out table
    /// rather than re-deriving it, and is the control the body path had to be
    /// brought in line with — it must not move.
    #[test]
    fn a_centered_spaced_floating_table_sits_equidistant_too() {
        let config = small_config();
        let blocks = vec![aligned_spaced_table(
            Some(crate::model::Alignment::Center),
            20.0,
            Some(super::TableFloatInfo {
                right_gap: Pt::ZERO,
                bottom_gap: Pt::ZERO,
                x_align: None,
                y_offset: Pt::ZERO,
                vert_anchor: crate::model::TableAnchor::Page,
                overlap: None,
            }),
        )];
        let pages = layout_section(&blocks, &config, None, Pt::ZERO, Pt::new(14.0), None);
        let (left, right) = painted_x_span(&pages[0]);
        let content_left = config.margins.left.raw();
        let content_right = content_left + config.content_width().raw();
        assert!(
            (left - content_left - (content_right - right)).abs() < 0.01,
            "centred float: left gap {} must equal right gap {}",
            left - content_left,
            content_right - right,
        );
    }

    // ── §17.3.1.24 paragraph border grouping tests ─────────────────────

    #[test]
    fn identical_borders_suppress_second_top() {
        use crate::render::layout::paragraph::{BorderLine, ParagraphBorderStyle};
        let border = Some(ParagraphBorderStyle {
            top: Some(BorderLine {
                width: Pt::new(0.5),
                color: RgbColor::BLACK,
                space: Pt::new(1.0),
            }),
            bottom: None,
            left: None,
            right: None,
        });
        let style1 = ParagraphStyle {
            borders: border.clone(),
            ..Default::default()
        };
        let style2 = ParagraphStyle {
            borders: border,
            ..Default::default()
        };

        let blocks = vec![
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("para1", 30.0, 14.0)],
                style: style1,
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            },
            LayoutBlock::Paragraph {
                fragments: vec![text_frag("para2", 30.0, 14.0)],
                style: style2,
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            },
        ];
        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        // Count Line draw commands (border lines).
        // Only the first paragraph should draw its top border; the second's
        // top border is suppressed by §17.3.1.24 grouping.
        let line_cmds: Vec<_> = pages[0]
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCommand::Line { .. }))
            .collect();
        assert_eq!(
            line_cmds.len(),
            1,
            "only one top border line (grouped): got {}",
            line_cmds.len()
        );
    }

    // ── §4 continuation re-fit around a next-page float (plan §11.3) ─────

    /// A paragraph carrying one absolute-positioned floating image (left side),
    /// registered as a `Square` wrap float by the forward scan.
    fn para_with_abs_left_float(text: &str, fx: f32, fy: f32, fw: f32, fh: f32) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, 20.0, 14.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![FloatingImage {
                image_data: crate::render::resolve::images::MediaEntry {
                    data: std::sync::Arc::from(&[][..]),
                    format: crate::model::ImageFormat::Png,
                },
                size: PtSize::new(Pt::new(fw), Pt::new(fh)),
                src_rect: None,
                x: FloatingImageX::Absolute(Pt::new(fx)),
                y: FloatingImageY::Absolute(Pt::new(fy)),
                wrap_mode: WrapMode::Square(crate::model::WrapText::BothSides),
                dist_left: Pt::ZERO,
                dist_right: Pt::ZERO,
                behind_doc: false,
            }],
            floating_shapes: vec![],
        }
    }

    #[test]
    fn continuation_re_fits_around_a_float_on_its_own_page() {
        // §4/§11.3: a paragraph splits across a page boundary; a LATER block on
        // the continuation page owns an absolute float in the top-left. The
        // continuation's re-fit (not a reuse of the starting-page placement)
        // must wrap its first lines around that float — they start at the
        // float's right edge (x≈60), not the page margin (x=10).
        //
        // Layout (small_config: content x∈[10,190]? width 180, y∈[10,90]):
        //  - 3 one-line fills push the cursor to y=52 (below the float band).
        //  - P: 6 single-line words (100pt each → one per line at any width
        //    ≥100). Only ~2 fit below the fills on page 1; the rest continue.
        //  - Q: an absolute float x=10,y=10,w=50,h=40 → band y∈[10,50].
        let float_block = para_with_abs_left_float("Q", 10.0, 10.0, 50.0, 40.0);
        let mut blocks: Vec<LayoutBlock> = (0..3)
            .map(|i| para_block(&format!("fill{i}"), 100.0))
            .collect();
        blocks.push(multiline_para(6, false, true));
        blocks.push(float_block);

        let pages = layout_section(
            &blocks,
            &small_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );
        assert!(pages.len() >= 2, "paragraph must split across pages");

        // First continuation line on page 2 is "word2" (2 words stayed on
        // page 1). It sits in the float band, so it must be indented past the
        // float's right edge.
        let word2_x = pages[1]
            .commands
            .iter()
            .find_map(|c| match c {
                DrawCommand::Text { text, position, .. } if text.trim() == "word2" => {
                    Some(position.x.raw())
                }
                _ => None,
            })
            .expect("word2 on page 2");
        assert!(
            (word2_x - 60.0).abs() < 0.5,
            "continuation must wrap around the next-page float (x≈60), got {word2_x}"
        );
    }

    // ── §17.6.4 splitting across unequal-width columns (plan §11.4) ──────

    /// Two columns of *unequal* width sharing one 200×100 page: col 0 is 100pt
    /// wide at the left margin (x=10), col 1 is 70pt wide at x=120.
    fn unequal_two_column_config() -> PageConfig {
        use crate::render::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(200.0), Pt::new(100.0)),
            margins: PtEdgeInsets::new(Pt::new(10.0), Pt::new(10.0), Pt::new(10.0), Pt::new(10.0)),
            header_margin: Pt::new(5.0),
            footer_margin: Pt::new(5.0),
            columns: vec![
                ColumnGeometry {
                    x_offset: Pt::ZERO,
                    width: Pt::new(100.0),
                },
                ColumnGeometry {
                    x_offset: Pt::new(110.0),
                    width: Pt::new(70.0),
                },
            ],
        }
    }

    #[test]
    fn paragraph_splits_across_unequal_width_columns() {
        // §11.4: the per-column re-fit (§11.3) removed the equal-width gate, so
        // a paragraph now flows across columns of different widths. 8 single-
        // line words (65pt → one per line in both the 100pt and 70pt columns);
        // ~5 fit in col 0's 80pt height, the rest re-fit into col 1.
        //
        // Before the gate was lifted this paragraph was laid out atomically: at
        // the column top it overflowed col 0 and never reached col 1.
        let fragments: Vec<Fragment> = (0..8)
            .map(|i| text_frag(&format!("w{i} "), 65.0, 14.0))
            .collect();
        let block = LayoutBlock::Paragraph {
            fragments,
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
            floating_shapes: vec![],
        };

        let pages = layout_section(
            &[block],
            &unequal_two_column_config(),
            None,
            Pt::ZERO,
            Pt::new(14.0),
            None,
        );

        // Collect (x, text) across every page.
        let placed: Vec<(f32, String)> = pages
            .iter()
            .flat_map(|p| p.commands.iter())
            .filter_map(|c| match c {
                DrawCommand::Text { position, text, .. } => {
                    Some((position.x.raw(), text.trim().to_string()))
                }
                _ => None,
            })
            .collect();

        assert_eq!(
            placed.len(),
            8,
            "every word emitted exactly once: {placed:?}"
        );
        let in_col0 = placed
            .iter()
            .filter(|(x, _)| (*x - 10.0).abs() < 0.5)
            .count();
        let in_col1 = placed
            .iter()
            .filter(|(x, _)| (*x - 120.0).abs() < 0.5)
            .count();
        assert!(
            in_col0 >= 2,
            "col 0 (x≈10, 100pt wide) holds the head: {placed:?}"
        );
        assert!(
            in_col1 >= 1,
            "col 1 (x≈120, 70pt wide) holds the re-fit continuation: {placed:?}"
        );
        assert_eq!(
            in_col0 + in_col1,
            8,
            "all words land in one column or the other"
        );
    }
}
