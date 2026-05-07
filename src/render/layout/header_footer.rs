//! Header/footer layout — render headers and footers on each page.
//!
//! Headers and footers are laid out in a separate constraint frame
//! (between page edge and body margin), then their draw commands are
//! prepended to each page's command list.
//!
//! Content is built per-page so that PAGE / NUMPAGES fields (§17.16.4.1)
//! evaluate to the correct values on each page.

use crate::model::Block;

use crate::render::dimension::Pt;
use crate::render::resolve::header_footer::{HeaderFooterKind, HeaderFooterSet};

use super::build::{build_header_footer_content, BuildContext, BuildState, HeaderFooterContent};
use super::draw_command::{DrawCommand, LayoutedPage};
use super::page::PageConfig;
use super::section::stack_blocks;

/// Decide which slot of a `HeaderFooterSet` applies to a given page,
/// per ECMA-376 §17.10.6 (`titlePg`) and §17.10.1
/// (`evenAndOddHeaders`). Returns `None` when the spec says the page
/// should have *no* header/footer (e.g. `titlePg` is set but the
/// section has no `first` slot — Word leaves the title page blank
/// rather than reusing `default`).
///
/// Inputs:
/// * `set` — the section's resolved slots.
/// * `first_in_section` — true on the first physical page of the
///   section. The section that owns the page is the one this set
///   came from.
/// * `logical_page_number` — 1-based, with `pgNumType.start` applied;
///   matches the value the `PAGE` field would render. Even/odd
///   selection is on this number, not the physical doc-absolute page.
/// * `title_pg` — the section's `<w:titlePg/>` flag.
/// * `even_and_odd` — the document setting `<w:evenAndOddHeaders/>`.
pub fn select_slot<T>(
    set: &HeaderFooterSet<T>,
    first_in_section: bool,
    logical_page_number: usize,
    title_pg: bool,
    even_and_odd: bool,
) -> Option<&T> {
    if title_pg && first_in_section {
        // Spec: title-page header is *its own* slot. If absent, the
        // first page of the section has a blank header — we do not
        // fall through to `default`.
        return set.first.as_ref();
    }
    if even_and_odd && logical_page_number.is_multiple_of(2) {
        // Same rule for even pages: the `even` slot is authoritative
        // when the document opts in. Missing `even` → blank.
        return set.even.as_ref();
    }
    set.default.as_ref()
}

/// Variant of [`select_slot`] that also tells the caller which kind of
/// slot was chosen. Useful for diagnostics and for callers that need
/// to bookkeep per-kind layout state (e.g. header-area heights).
pub fn select_slot_with_kind<T>(
    set: &HeaderFooterSet<T>,
    first_in_section: bool,
    logical_page_number: usize,
    title_pg: bool,
    even_and_odd: bool,
) -> Option<(HeaderFooterKind, &T)> {
    if title_pg && first_in_section {
        return set.first.as_ref().map(|t| (HeaderFooterKind::First, t));
    }
    if even_and_odd && logical_page_number.is_multiple_of(2) {
        return set.even.as_ref().map(|t| (HeaderFooterKind::Even, t));
    }
    set.default.as_ref().map(|t| (HeaderFooterKind::Default, t))
}

/// Header and footer slots for a section, plus the spec flags that
/// drive per-page selection. Selection rules live in [`select_slot`];
/// the layout step calls it once per page in `render_headers_footers`.
pub struct HeaderFooterBlocks<'a> {
    /// Section's three header slots (`default` / `first` / `even`).
    pub headers: &'a HeaderFooterSet<Vec<Block>>,
    /// Section's three footer slots.
    pub footers: &'a HeaderFooterSet<Vec<Block>>,
    /// `<w:titlePg/>` set on the section.
    pub title_pg: bool,
    /// Document-level `<w:evenAndOddHeaders/>`.
    pub even_and_odd: bool,
}

/// Page numbering context for header/footer field evaluation.
///
/// `logical_page_base` is the value the `PAGE` field reports on the
/// first page of this section. Per §17.6.12 it is determined by the
/// section's `<w:pgNumType w:start="…"/>`, falling back to a continuation
/// of the previous section's logical numbering (see
/// [`next_logical_page_base`]). Even/odd header selection (§17.10.1)
/// also operates on this number.
pub struct PageRange {
    /// 0-based index of the first page in this section within the document.
    pub page_base: usize,
    /// Logical (PAGE-field) value of the first page in this section.
    pub logical_page_base: usize,
    /// Total physical page count in the document. NUMPAGES uses this
    /// value; it deliberately tracks physical, not logical, pages —
    /// matches Word.
    pub total_pages: usize,
}

/// §17.6.12 — compute the logical page number (PAGE-field value) of a
/// section's first page. If the section sets `<w:pgNumType w:start>`,
/// numbering resets to that value; otherwise it continues from the
/// previous section.
///
/// `prev_logical_end` is the value the next page *would* have if
/// numbering simply continued — i.e. one past the last logical page of
/// the previous section. For the first section, callers pass `1` so
/// the document starts at logical page 1 by default.
pub fn next_logical_page_base(
    prev_logical_end: usize,
    page_number_type: Option<&crate::model::PageNumberType>,
) -> usize {
    page_number_type
        .and_then(|pnt| pnt.start)
        .map(|s| s as usize)
        .unwrap_or(prev_logical_end)
}

/// Render headers and footers onto already-laid-out pages.
///
/// `hf_blocks` contains the raw DOCX header/footer blocks; content is
/// rebuilt per-page so that field values (PAGE, NUMPAGES) are correct.
pub fn render_headers_footers(
    pages: &mut [LayoutedPage],
    config: &PageConfig,
    hf_blocks: &HeaderFooterBlocks<'_>,
    ctx: &BuildContext,
    state: &mut BuildState,
    default_line_height: Pt,
    page_range: &PageRange,
) {
    let content_width = config.content_width();

    for (page_idx, page) in pages.iter_mut().enumerate() {
        // Logical page number drives both PAGE-field rendering
        // (§17.16.4.1) and even/odd header selection (§17.10.1).
        let logical_page_number = page_range.logical_page_base + page_idx;
        let first_in_section = page_idx == 0;

        // §17.10.6 + §17.10.1: pick the slot that applies to this page.
        // The selection is the same for header and footer; the spec
        // doesn't admit asymmetry between the two.
        let header_blocks = select_slot(
            hf_blocks.headers,
            first_in_section,
            logical_page_number,
            hf_blocks.title_pg,
            hf_blocks.even_and_odd,
        );
        let footer_blocks = select_slot(
            hf_blocks.footers,
            first_in_section,
            logical_page_number,
            hf_blocks.title_pg,
            hf_blocks.even_and_odd,
        );

        if let Some(blocks) = header_blocks {
            // Set per-page field context for PAGE/NUMPAGES evaluation.
            state.field_ctx = crate::render::layout::fragment::FieldContext {
                page_number: Some(logical_page_number),
                num_pages: Some(page_range.total_pages),
            };

            let hf = build_header_footer_content(blocks, ctx, state);
            render_header(page, config, &hf, content_width, default_line_height);
        }

        if let Some(blocks) = footer_blocks {
            state.field_ctx = crate::render::layout::fragment::FieldContext {
                page_number: Some(logical_page_number),
                num_pages: Some(page_range.total_pages),
            };

            let hf = build_header_footer_content(blocks, ctx, state);
            render_footer(page, config, &hf, content_width, default_line_height);
        }
    }

    // Reset field context after header/footer rendering.
    state.field_ctx = crate::render::layout::fragment::FieldContext::default();
}

/// Render a single header onto a page.
fn render_header(
    page: &mut LayoutedPage,
    config: &PageConfig,
    hf: &HeaderFooterContent,
    content_width: Pt,
    default_line_height: Pt,
) {
    if hf.blocks.is_empty() {
        return;
    }

    let (offset_x, offset_y) = if let Some((abs_x, abs_y)) = hf.absolute_position {
        (abs_x, abs_y)
    } else {
        (config.margins.left, config.header_margin)
    };

    let result = stack_blocks(&hf.blocks, content_width, default_line_height, None);

    let mut header_cmds: Vec<DrawCommand> = Vec::new();

    // §20.4.2.3 @behindDoc=true: paint behind text — emit before text commands.
    for fi in hf.floating_images.iter().filter(|fi| fi.behind_doc) {
        let img_y = match fi.y {
            super::section::FloatingImageY::Absolute(y) => y,
            super::section::FloatingImageY::RelativeToParagraph(off) => offset_y + off,
        };
        header_cmds.push(DrawCommand::Image {
            rect: crate::render::geometry::PtRect::from_xywh(
                fi.x,
                img_y,
                fi.size.width,
                fi.size.height,
            ),
            image_data: fi.image_data.clone(),
        });
    }

    // Text / table commands.
    for mut cmd in result.commands {
        cmd.shift(offset_x, offset_y);
        header_cmds.push(cmd);
    }

    // §20.4.2.3 @behindDoc=false: paint in front of text — emit after text commands.
    for fi in hf.floating_images.iter().filter(|fi| !fi.behind_doc) {
        let img_y = match fi.y {
            super::section::FloatingImageY::Absolute(y) => y,
            super::section::FloatingImageY::RelativeToParagraph(off) => offset_y + off,
        };
        header_cmds.push(DrawCommand::Image {
            rect: crate::render::geometry::PtRect::from_xywh(
                fi.x,
                img_y,
                fi.size.width,
                fi.size.height,
            ),
            image_data: fi.image_data.clone(),
        });
    }
    // Paragraph-anchored shapes ride through `result.commands` above (their
    // y depends on the host paragraph). Page-anchored shapes were resolved
    // in `Page` frame so their absolute y is authoritative — emit them here
    // without applying the header's offset shift.
    emit_page_anchored_shapes(&hf.floating_shapes, &mut header_cmds);

    // Prepend header commands before body content.
    header_cmds.append(&mut page.commands);
    page.commands = header_cmds;
}

/// Render a single footer onto a page.
fn render_footer(
    page: &mut LayoutedPage,
    config: &PageConfig,
    hf: &HeaderFooterContent,
    content_width: Pt,
    default_line_height: Pt,
) {
    if hf.blocks.is_empty() {
        return;
    }

    let result = stack_blocks(&hf.blocks, content_width, default_line_height, None);

    let footer_y = config.page_size.height - config.footer_margin - result.height;

    // §20.4.2.3 @behindDoc=true: paint behind text.
    for fi in hf.floating_images.iter().filter(|fi| fi.behind_doc) {
        let img_y = match fi.y {
            super::section::FloatingImageY::Absolute(y) => y,
            super::section::FloatingImageY::RelativeToParagraph(off) => footer_y + off,
        };
        page.commands.push(DrawCommand::Image {
            rect: crate::render::geometry::PtRect::from_xywh(
                fi.x,
                img_y,
                fi.size.width,
                fi.size.height,
            ),
            image_data: fi.image_data.clone(),
        });
    }

    for mut cmd in result.commands {
        cmd.shift(config.margins.left, footer_y);
        page.commands.push(cmd);
    }

    // §20.4.2.3 @behindDoc=false: paint in front of text.
    for fi in hf.floating_images.iter().filter(|fi| !fi.behind_doc) {
        let img_y = match fi.y {
            super::section::FloatingImageY::Absolute(y) => y,
            super::section::FloatingImageY::RelativeToParagraph(off) => footer_y + off,
        };
        page.commands.push(DrawCommand::Image {
            rect: crate::render::geometry::PtRect::from_xywh(
                fi.x,
                img_y,
                fi.size.width,
                fi.size.height,
            ),
            image_data: fi.image_data.clone(),
        });
    }
    // Paragraph-anchored shapes ride through `result.commands` above (their
    // y depends on the host paragraph). Page-anchored shapes were resolved
    // in `Page` frame so their absolute y is authoritative — emit them here
    // without applying the footer's stack shift.
    emit_page_anchored_shapes(&hf.floating_shapes, &mut page.commands);
}

/// Emit page-anchored floating shapes (§20.4.2.10 vertical anchor =
/// `page` / `margin` / etc.) as Path + text commands at their absolute
/// page coordinates. Mirrors the in-stacker emission in `stacker.rs`,
/// but skipped here for clarity:
/// no spacing collapse / paragraph anchoring, just absolute placement.
fn emit_page_anchored_shapes(shapes: &[super::section::FloatingShape], out: &mut Vec<DrawCommand>) {
    use super::section::FloatingImageY;
    for fs in shapes {
        let shape_y = match fs.y {
            FloatingImageY::Absolute(y) => y,
            // The extractor only routes Absolute-y shapes here; treat
            // any RelativeToParagraph as an unreachable misroute and
            // skip it rather than stacking on a non-existent paragraph.
            FloatingImageY::RelativeToParagraph(_) => continue,
        };
        out.push(DrawCommand::Path {
            origin: crate::render::geometry::PtOffset::new(fs.x, shape_y),
            rotation: fs.rotation,
            flip_h: fs.flip_h,
            flip_v: fs.flip_v,
            extent: fs.size,
            paths: fs.paths.clone(),
            fill: fs.fill.clone(),
            stroke: fs.stroke.clone(),
            effects: fs.effects.clone(),
        });
        for mut cmd in fs.text_commands.iter().cloned() {
            cmd.shift(fs.x, shape_y);
            out.push(cmd);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::geometry::{PtEdgeInsets, PtOffset, PtSize};
    use crate::render::layout::fragment::{FontProps, Fragment, TextMetrics};
    use crate::render::layout::paragraph::ParagraphStyle;
    use crate::render::layout::section::LayoutBlock;
    use crate::render::resolve::color::RgbColor;
    use std::rc::Rc;

    fn make_hf(frags: Vec<Fragment>) -> HeaderFooterContent {
        HeaderFooterContent {
            blocks: vec![LayoutBlock::Paragraph {
                fragments: frags,
                style: ParagraphStyle::default(),
                page_break_before: false,
                footnotes: vec![],
                floating_images: vec![],
                floating_shapes: vec![],
            }],
            absolute_position: None,
            floating_images: vec![],
            floating_shapes: vec![],
        }
    }

    fn text_frag(s: &str) -> Fragment {
        let font = FontProps {
            family: Rc::from("Test"),
            size: Pt::new(12.0),
            bold: false,
            italic: false,
            underline: false,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        };
        Fragment::Text {
            text: Rc::from(s),
            font,
            color: RgbColor::BLACK,
            shading: None,
            border: None,
            width: Pt::new(40.0),
            trimmed_width: Pt::new(40.0),
            metrics: TextMetrics {
                ascent: Pt::new(10.0),
                descent: Pt::new(4.0),
                leading: Pt::ZERO,
            },
            hyperlink_url: None,
            baseline_offset: Pt::ZERO,
            text_offset: Pt::ZERO,
        }
    }

    fn test_config() -> PageConfig {
        use crate::render::layout::page::ColumnGeometry;
        PageConfig {
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
            margins: PtEdgeInsets::new(Pt::new(72.0), Pt::new(72.0), Pt::new(72.0), Pt::new(72.0)),
            header_margin: Pt::new(36.0),
            footer_margin: Pt::new(36.0),
            columns: vec![ColumnGeometry {
                x_offset: Pt::ZERO,
                width: Pt::new(468.0),
            }],
        }
    }

    #[test]
    fn no_header_footer_leaves_page_unchanged() {
        let mut pages = [LayoutedPage::new(PtSize::new(
            Pt::new(612.0),
            Pt::new(792.0),
        ))];
        pages[0].commands.push(DrawCommand::Text {
            text: "body".into(),
            position: PtOffset::new(Pt::ZERO, Pt::ZERO),
            font_family: Rc::from("T"),
            font_size: Pt::new(12.0),
            char_spacing: Pt::ZERO,
            bold: false,
            italic: false,
            color: RgbColor::BLACK,
            text_scale: 1.0,
        });

        let config = test_config();
        // Direct call to render_header / render_footer with empty content.
        let hf = HeaderFooterContent {
            blocks: vec![],
            absolute_position: None,
            floating_images: vec![],
            floating_shapes: vec![],
        };
        render_header(
            &mut pages[0],
            &config,
            &hf,
            config.content_width(),
            Pt::new(14.0),
        );
        render_footer(
            &mut pages[0],
            &config,
            &hf,
            config.content_width(),
            Pt::new(14.0),
        );

        assert_eq!(pages[0].commands.len(), 1, "no changes");
    }

    #[test]
    fn header_prepended_to_page() {
        let mut pages = [LayoutedPage::new(PtSize::new(
            Pt::new(612.0),
            Pt::new(792.0),
        ))];
        pages[0].commands.push(DrawCommand::Text {
            text: "body".into(),
            position: PtOffset::new(Pt::ZERO, Pt::ZERO),
            font_family: Rc::from("T"),
            font_size: Pt::new(12.0),
            char_spacing: Pt::ZERO,
            bold: false,
            italic: false,
            color: RgbColor::BLACK,
            text_scale: 1.0,
        });

        let config = test_config();
        let header = make_hf(vec![text_frag("Header")]);
        render_header(
            &mut pages[0],
            &config,
            &header,
            config.content_width(),
            Pt::new(14.0),
        );

        assert!(pages[0].commands.len() > 1);
        // First command should be the header text
        if let DrawCommand::Text { text, .. } = &pages[0].commands[0] {
            assert_eq!(&**text, "Header");
        }
    }

    #[test]
    fn footer_appended_to_page() {
        let mut pages = [LayoutedPage::new(PtSize::new(
            Pt::new(612.0),
            Pt::new(792.0),
        ))];

        let config = test_config();
        let footer = make_hf(vec![text_frag("Footer")]);
        render_footer(
            &mut pages[0],
            &config,
            &footer,
            config.content_width(),
            Pt::new(14.0),
        );

        assert_eq!(pages[0].commands.len(), 1);
        if let DrawCommand::Text { text, position, .. } = &pages[0].commands[0] {
            assert_eq!(&**text, "Footer");
            // Footer y should be near the bottom of the page.
            assert!(position.y.raw() > 700.0, "footer y={}", position.y.raw());
        }
    }

    #[test]
    fn header_applied_to_all_pages() {
        let mut pages = vec![
            LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0))),
            LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0))),
        ];

        let config = test_config();
        let header = make_hf(vec![text_frag("H")]);
        for page in pages.iter_mut() {
            render_header(
                page,
                &config,
                &header,
                config.content_width(),
                Pt::new(14.0),
            );
        }

        // Both pages should have header
        for page in &pages {
            assert!(!page.commands.is_empty());
        }
    }

    #[test]
    fn header_y_position_uses_header_margin() {
        let mut pages = [LayoutedPage::new(PtSize::new(
            Pt::new(612.0),
            Pt::new(792.0),
        ))];
        let config = test_config();
        let header = make_hf(vec![text_frag("H")]);
        render_header(
            &mut pages[0],
            &config,
            &header,
            config.content_width(),
            Pt::new(14.0),
        );

        if let DrawCommand::Text { position, .. } = &pages[0].commands[0] {
            // Header y should be near header_margin (36) + ascent
            assert!(
                position.y.raw() > 36.0 && position.y.raw() < 72.0,
                "header y={} should be between header_margin and top margin",
                position.y.raw()
            );
        }
    }

    /// Truth-table for `select_slot` covering ECMA-376 §17.10.6
    /// (`titlePg`) and §17.10.1 (`evenAndOddHeaders`). Every test names
    /// the rule it pins down; failures should pinpoint a specific
    /// spec-rule violation.
    mod selection {
        use super::super::*;

        /// Helper: a fully-populated set with marker strings so tests
        /// can assert *which* slot was returned without reaching for
        /// equality against block content.
        fn full_set() -> HeaderFooterSet<&'static str> {
            HeaderFooterSet {
                default: Some("D"),
                first: Some("F"),
                even: Some("E"),
            }
        }

        // §17.10.6 — titlePg ----------------------------------------

        #[test]
        fn title_page_with_first_slot_returns_first_on_page_one() {
            let set = full_set();
            assert_eq!(select_slot(&set, true, 1, true, false), Some(&"F"));
        }

        #[test]
        fn title_page_without_first_slot_returns_none_not_default() {
            // Spec literal + Word behavior: a missing `first` leaves
            // the title page blank; it does NOT silently fall through
            // to `default`.
            let set = HeaderFooterSet::<&'static str> {
                default: Some("D"),
                first: None,
                even: None,
            };
            assert_eq!(select_slot(&set, true, 1, true, false), None);
        }

        #[test]
        fn title_page_flag_off_keeps_default_on_page_one() {
            let set = full_set();
            assert_eq!(select_slot(&set, true, 1, false, false), Some(&"D"));
        }

        #[test]
        fn title_page_only_applies_to_first_page_of_section() {
            // Even with `titlePg` on, page 2 of the section uses
            // `default` (or `even` if the parity flag would apply,
            // see combined-rule tests below).
            let set = full_set();
            assert_eq!(select_slot(&set, false, 2, true, false), Some(&"D"));
        }

        // §17.10.1 — evenAndOddHeaders ------------------------------

        #[test]
        fn even_and_odd_with_even_slot_returns_even_on_even_pages() {
            let set = full_set();
            assert_eq!(select_slot(&set, false, 2, false, true), Some(&"E"));
            assert_eq!(select_slot(&set, false, 4, false, true), Some(&"E"));
        }

        #[test]
        fn even_and_odd_without_even_slot_returns_none_not_default() {
            let set = HeaderFooterSet::<&'static str> {
                default: Some("D"),
                first: None,
                even: None,
            };
            assert_eq!(select_slot(&set, false, 2, false, true), None);
        }

        #[test]
        fn even_and_odd_uses_default_on_odd_pages() {
            let set = full_set();
            assert_eq!(select_slot(&set, false, 1, false, true), Some(&"D"));
            assert_eq!(select_slot(&set, false, 3, false, true), Some(&"D"));
        }

        #[test]
        fn even_and_odd_flag_off_keeps_default_on_even_pages() {
            // Without the document setting, the `even` slot is dead
            // weight even on page 2.
            let set = full_set();
            assert_eq!(select_slot(&set, false, 2, false, false), Some(&"D"));
        }

        // Combined rules --------------------------------------------

        #[test]
        fn title_page_takes_precedence_over_even_and_odd_on_first_page() {
            // titlePg with first_in_section=true wins outright, even
            // when evenAndOddHeaders is on and the page is even.
            let set = full_set();
            assert_eq!(select_slot(&set, true, 2, true, true), Some(&"F"));
            // Same rule with no `first` slot — still the title-page
            // rule applies, returning None (not `even`).
            let set2 = HeaderFooterSet::<&'static str> {
                default: Some("D"),
                first: None,
                even: Some("E"),
            };
            assert_eq!(select_slot(&set2, true, 2, true, true), None);
        }

        #[test]
        fn even_and_odd_governs_after_title_page_rule_is_satisfied() {
            // first_in_section=false on page 2 with both flags on
            // means the title-page rule no longer fires; even/odd
            // does and returns `even`.
            let set = full_set();
            assert_eq!(select_slot(&set, false, 2, true, true), Some(&"E"));
        }

        // Universal fall-throughs -----------------------------------

        #[test]
        fn empty_set_returns_none_in_every_mode() {
            let empty: HeaderFooterSet<&'static str> = HeaderFooterSet::default();
            for &title_pg in &[false, true] {
                for &even_and_odd in &[false, true] {
                    for &first in &[false, true] {
                        for page in 1..=3 {
                            assert_eq!(
                                select_slot(&empty, first, page, title_pg, even_and_odd),
                                None,
                                "empty set with title_pg={title_pg} even_and_odd={even_and_odd} \
                                 first_in_section={first} page={page} must be None",
                            );
                        }
                    }
                }
            }
        }

        #[test]
        fn default_used_when_neither_flag_applies() {
            let set = full_set();
            assert_eq!(select_slot(&set, false, 1, false, false), Some(&"D"));
            assert_eq!(select_slot(&set, false, 2, false, false), Some(&"D"));
            assert_eq!(select_slot(&set, true, 5, false, false), Some(&"D"));
        }

        // §17.6.12 — pgNumType.start --------------------------------

        #[test]
        fn next_logical_page_base_continues_when_no_pg_num_type() {
            // Section without `pgNumType` continues numbering from
            // `prev_logical_end` (one past the previous section's last
            // logical page).
            assert_eq!(next_logical_page_base(1, None), 1);
            assert_eq!(next_logical_page_base(7, None), 7);
        }

        #[test]
        fn next_logical_page_base_uses_start_when_set() {
            use crate::model::PageNumberType;
            let pnt = PageNumberType {
                format: None,
                start: Some(5),
                chap_style: None,
                chap_sep: None,
            };
            // The continuation hint is overridden by the explicit start.
            assert_eq!(next_logical_page_base(99, Some(&pnt)), 5);
        }

        #[test]
        fn next_logical_page_base_falls_through_when_start_is_none() {
            use crate::model::PageNumberType;
            // pgNumType present but `start` not set — selection should
            // ignore the (possibly non-default) format/chapStyle and
            // continue numbering normally.
            let pnt = PageNumberType {
                format: None,
                start: None,
                chap_style: None,
                chap_sep: None,
            };
            assert_eq!(next_logical_page_base(4, Some(&pnt)), 4);
        }

        // Variant returning the kind --------------------------------

        #[test]
        fn select_slot_with_kind_reports_the_chosen_slot() {
            let set = full_set();
            assert_eq!(
                select_slot_with_kind(&set, true, 1, true, false),
                Some((HeaderFooterKind::First, &"F")),
            );
            assert_eq!(
                select_slot_with_kind(&set, false, 2, false, true),
                Some((HeaderFooterKind::Even, &"E")),
            );
            assert_eq!(
                select_slot_with_kind(&set, false, 3, false, true),
                Some((HeaderFooterKind::Default, &"D")),
            );
        }
    }
}
