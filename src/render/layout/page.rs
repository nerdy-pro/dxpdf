//! Page configuration — convert section properties to layout-ready config.

use crate::model::{Columns, SectionProperties};

use crate::render::dimension::Pt;
use crate::render::geometry::{PtEdgeInsets, PtSize};

/// §17.6.13: default page width when w:pgSz is absent. Word uses US Letter (8.5 inches = 12240 twips).
const SPEC_DEFAULT_PAGE_WIDTH: Pt = Pt::new(612.0);
/// §17.6.13: default page height when w:pgSz is absent. Word uses US Letter (11 inches = 15840 twips).
const SPEC_DEFAULT_PAGE_HEIGHT: Pt = Pt::new(792.0);
/// §17.6.11: default page margin when w:pgMar is absent (1 inch = 1440 twips).
const SPEC_DEFAULT_MARGIN: Pt = Pt::new(72.0);

/// A single column's layout geometry.
#[derive(Debug, Clone, Copy)]
pub struct ColumnGeometry {
    /// X offset of this column relative to the left page margin.
    pub x_offset: Pt,
    /// Available text width within this column.
    pub width: Pt,
}

/// Page layout configuration in points.
#[derive(Debug, Clone)]
pub struct PageConfig {
    pub page_size: PtSize,
    pub margins: PtEdgeInsets,
    pub header_margin: Pt,
    pub footer_margin: Pt,
    /// §17.6.4: column layout. Single-element vec for normal single-column.
    pub columns: Vec<ColumnGeometry>,
    /// §17.6.6 `w:bidi`: which side of the content area is the section's
    /// **leading** margin.
    ///
    /// Carried as the same [`BaseDirection`](crate::i18n::bidi::BaseDirection) a
    /// paragraph resolves its own
    /// `w:bidi` to, rather than a second boolean, so the one block-level reader
    /// (`section::helpers::table_x_offset`) asks the question in the engine's
    /// existing vocabulary.
    ///
    /// **What reads it, and what deliberately does not.** §17.6.6 also makes
    /// right-to-left the default *paragraph* direction inside the section and
    /// reverses the order of a multi-column layout. Neither is wired: a
    /// paragraph already resolves its own `w:bidi` (`ParagraphStyle::base_direction`)
    /// and giving it a section-level default is its own unit with its own
    /// regression surface, and column order is untouched by anything here. Only
    /// table placement — §17.4.28 `jc` and §17.4.50 `tblInd` — consults this
    /// today.
    pub base_direction: crate::i18n::bidi::BaseDirection,
}

impl Default for PageConfig {
    fn default() -> Self {
        let content_width = SPEC_DEFAULT_PAGE_WIDTH - SPEC_DEFAULT_MARGIN - SPEC_DEFAULT_MARGIN;
        Self {
            page_size: PtSize::new(SPEC_DEFAULT_PAGE_WIDTH, SPEC_DEFAULT_PAGE_HEIGHT),
            margins: PtEdgeInsets::new(
                SPEC_DEFAULT_MARGIN,
                SPEC_DEFAULT_MARGIN,
                SPEC_DEFAULT_MARGIN,
                SPEC_DEFAULT_MARGIN,
            ),
            header_margin: SPEC_DEFAULT_MARGIN / 2.0,
            footer_margin: SPEC_DEFAULT_MARGIN / 2.0,
            columns: vec![ColumnGeometry {
                x_offset: Pt::ZERO,
                width: content_width,
            }],
            base_direction: crate::i18n::bidi::BaseDirection::Ltr,
        }
    }
}

impl PageConfig {
    /// Build from section properties, falling back to US Letter defaults.
    pub fn from_section(sect: &SectionProperties) -> Self {
        let mut cfg = Self::default();

        if let Some(ps) = sect.page_size.get() {
            if let Some(w) = ps.width {
                cfg.page_size.width = Pt::from(w);
            }
            if let Some(h) = ps.height {
                cfg.page_size.height = Pt::from(h);
            }
        }

        if let Some(pm) = sect.page_margins.get() {
            if let Some(t) = pm.top {
                cfg.margins.top = Pt::from(t);
            }
            if let Some(r) = pm.right {
                cfg.margins.right = Pt::from(r);
            }
            if let Some(b) = pm.bottom {
                cfg.margins.bottom = Pt::from(b);
            }
            if let Some(l) = pm.left {
                cfg.margins.left = Pt::from(l);
            }
            if let Some(h) = pm.header {
                cfg.header_margin = Pt::from(h);
            }
            if let Some(f) = pm.footer {
                cfg.footer_margin = Pt::from(f);
            }
        }

        // §17.6.4: compute column geometry from the *clamped* text width, so a
        // page whose margins exceed its width yields zero-width columns rather
        // than negative ones.
        cfg.columns = compute_columns(cfg.content_width(), sect.columns.get());

        // §17.6.6: absent or `w:val="0"` leaves the section left-to-right.
        if sect.bidi == Some(true) {
            cfg.base_direction = crate::i18n::bidi::BaseDirection::Rtl;
        }

        cfg
    }

    /// Available width for body content (page width minus left and right
    /// margins), never negative.
    ///
    /// The clamp is load-bearing, not defensive tidiness: margins wider than the
    /// page are expressible in the file format, and a negative width flows
    /// straight into `BoxConstraints::new`, whose `debug_assert!(min <= max)`
    /// then fires — a panic through the public `convert()` API in debug builds,
    /// and a negative-width layout in release. See also `compute_columns`,
    /// which clamps the same quantity per column.
    pub fn content_width(&self) -> Pt {
        (self.page_size.width - self.margins.left - self.margins.right).max(Pt::ZERO)
    }

    /// Number of columns in this section.
    pub fn num_columns(&self) -> usize {
        self.columns.len()
    }

    /// Available height for body content (page height minus top and bottom
    /// margins), never negative — same reasoning as [`PageConfig::content_width`].
    pub fn content_height(&self) -> Pt {
        (self.page_size.height - self.margins.top - self.margins.bottom).max(Pt::ZERO)
    }
}

/// §17.6.4: default spacing between columns when `w:space` is omitted
/// (720 twips = 0.5 inch).
const DEFAULT_COLUMN_SPACE: Pt = Pt::new(36.0);

/// Narrowest column worth laying out. Used only to bound a file-controlled
/// column count — see [`clamp_column_count`].
const MIN_COLUMN_WIDTH: Pt = Pt::new(1.0);

/// §17.6.4: bound the requested column count by what the page can physically
/// hold.
///
/// `w:num` is a file-controlled `u32` that both sizes an allocation and drives a
/// loop, so it needs a ceiling. The ceiling is derived from the page rather than
/// picked as a magic number: a column narrower than [`MIN_COLUMN_WIDTH`] cannot
/// show anything, so a count implying narrower columns describes no layout a
/// document could have meant.
///
/// **The gaps count.** `n` columns carry `n - 1` gaps of `space`, so the widest
/// usable count solves
/// `(content_width - space·(n-1)) / n >= MIN_COLUMN_WIDTH`, i.e.
/// `n <= (content_width + space) / (MIN_COLUMN_WIDTH + space)`. **This is the
/// only thing keeping the column width non-negative** — `compute_columns`
/// deliberately does not clamp afterwards. Bounding on
/// `content_width / MIN_COLUMN_WIDTH` alone ignores the gaps and still admits
/// counts whose column width comes out **negative** — which is what made a
/// 15-column section panic. Solving for the gaps is what makes the width
/// non-negative *by construction* rather than by a later clamp.
fn clamp_column_count(requested: u32, content_width: Pt, space: Pt) -> usize {
    let requested = (requested as usize).max(1);
    if content_width <= Pt::ZERO {
        return 1;
    }
    let space = space.raw().max(0.0);
    let max_by_page =
        (((content_width.raw() + space) / (MIN_COLUMN_WIDTH.raw() + space)) as usize).max(1);
    if requested > max_by_page {
        log::warn!(
            "§17.6.4: w:num={requested} does not fit a {content_width:?} text area at \
             {space}pt spacing; clamped to {max_by_page}"
        );
    }
    requested.min(max_by_page)
}

/// §17.6.4: compute column geometry from section column properties.
///
/// `w:num` is authoritative for the **count**: a row of `<w:col>` children
/// shorter than `w:num` leaves the remaining columns at the default width
/// rather than dropping them, and children beyond `w:num` are ignored.
///
/// Individual `<w:col>` widths apply only when the file **explicitly** opts out
/// of equal widths. §17.6.4's `equalWidth` defaults to `true` when omitted, so
/// an absent attribute means equal widths even if `<w:col>` children are
/// present — Word writes `equalWidth="0"` whenever it means the individual
/// definitions to win.
///
/// **Tier 0:** §17.6.4 `w:sep` — the vertical rule drawn between columns — is
/// parsed onto `model::Columns::separator` but never drawn. Position and width
/// of the columns themselves are unaffected; only the divider line is missing.
fn compute_columns(content_width: Pt, columns: Option<&Columns>) -> Vec<ColumnGeometry> {
    let single = || {
        vec![ColumnGeometry {
            x_offset: Pt::ZERO,
            width: content_width,
        }]
    };

    let cols = match columns {
        Some(c) if c.count.unwrap_or(1) > 1 => c,
        _ => return single(),
    };

    let default_space = cols.space.map(Pt::from).unwrap_or(DEFAULT_COLUMN_SPACE);
    let num = clamp_column_count(cols.count.unwrap_or(1), content_width, default_space);
    if num <= 1 {
        return single();
    }

    // Individual definitions, only on an explicit `equalWidth="0"`.
    if cols.equal_width == Some(false) && !cols.columns.is_empty() {
        let mut result = Vec::with_capacity(num);
        let mut x = Pt::ZERO;
        for i in 0..num {
            let def = cols.columns.get(i);
            let w = def
                .and_then(|d| d.width)
                .map(Pt::from)
                .unwrap_or(content_width / num as f32);
            result.push(ColumnGeometry {
                x_offset: x,
                width: w,
            });
            if i + 1 < num {
                let gap = def
                    .and_then(|d| d.space)
                    .map(Pt::from)
                    .unwrap_or(default_space);
                x += w + gap;
            }
        }
        return result;
    }

    // Equal-width columns.
    let total_gap = default_space * (num as f32 - 1.0);
    // No clamp here on purpose. `clamp_column_count` is the single guarantee
    // that this is positive, and a second `.max(Pt::ZERO)` was demonstrably
    // inert — removing it changed no test, because the bound already excludes
    // every count that could make it negative. Two mechanisms for one invariant
    // means the redundant one rots unnoticed; if the bound below ever changes,
    // this is the line that breaks, and `column_count_never_yields_a_negative_width`
    // is the test that catches it.
    let col_width = (content_width - total_gap) / num as f32;
    let mut result = Vec::with_capacity(num);
    for i in 0..num {
        let x = (col_width + default_space) * i as f32;
        result.push(ColumnGeometry {
            x_offset: x,
            width: col_width,
        });
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::{Dimension, Twips};
    use crate::model::{Dup, PageMargins, PageSize};

    #[test]
    fn default_is_us_letter() {
        let cfg = PageConfig::default();
        assert_eq!(cfg.page_size.width.raw(), 612.0);
        assert_eq!(cfg.page_size.height.raw(), 792.0);
        assert_eq!(cfg.margins.top.raw(), 72.0);
    }

    #[test]
    fn content_dimensions() {
        let cfg = PageConfig::default();
        assert_eq!(cfg.content_width().raw(), 468.0); // 612 - 72 - 72
        assert_eq!(cfg.content_height().raw(), 648.0); // 792 - 72 - 72
    }

    #[test]
    fn from_section_with_page_size() {
        let sect = SectionProperties {
            page_size: Dup::from(Some(PageSize {
                width: Some(Dimension::<Twips>::new(12240)), // 8.5in = 612pt
                height: Some(Dimension::<Twips>::new(15840)), // 11in = 792pt
                orientation: None,
            })),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.page_size.width.raw(), 612.0);
        assert_eq!(cfg.page_size.height.raw(), 792.0);
    }

    #[test]
    fn from_section_with_margins() {
        let sect = SectionProperties {
            page_margins: Dup::from(Some(PageMargins {
                top: Some(Dimension::<Twips>::new(1440)), // 1in = 72pt
                right: Some(Dimension::<Twips>::new(1440)),
                bottom: Some(Dimension::<Twips>::new(1440)),
                left: Some(Dimension::<Twips>::new(1440)),
                header: Some(Dimension::<Twips>::new(720)), // 0.5in = 36pt
                footer: Some(Dimension::<Twips>::new(720)),
                gutter: None,
            })),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.margins.top.raw(), 72.0);
        assert_eq!(cfg.header_margin.raw(), 36.0);
    }

    #[test]
    fn from_section_partial_uses_defaults() {
        let sect = SectionProperties {
            page_margins: Dup::from(Some(PageMargins {
                top: Some(Dimension::<Twips>::new(2880)), // 2in = 144pt
                right: None,
                bottom: None,
                left: None,
                header: None,
                footer: None,
                gutter: None,
            })),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.margins.top.raw(), 144.0, "custom top margin");
        assert_eq!(cfg.margins.right.raw(), 72.0, "default right margin");
    }

    // ── §17.6.4 compute_columns ──────────────────────────────────────────

    use crate::model::{ColumnDefinition, Columns};

    fn cols_of(c: Columns) -> Vec<ColumnGeometry> {
        compute_columns(Pt::new(468.0), Some(&c))
    }

    fn base() -> Columns {
        Columns {
            count: None,
            space: None,
            equal_width: None,
            separator: None,
            columns: vec![],
        }
    }

    #[test]
    fn no_columns_is_one_full_width_column() {
        let g = compute_columns(Pt::new(468.0), None);
        assert_eq!(g.len(), 1);
        assert_eq!(g[0].x_offset, Pt::ZERO);
        assert_eq!(g[0].width, Pt::new(468.0));
        // `w:num="1"` is the same thing stated explicitly.
        let one = cols_of(Columns {
            count: Some(1),
            ..base()
        });
        assert_eq!(one.len(), 1);
        assert_eq!(one[0].width, Pt::new(468.0));
    }

    /// Equal widths: gaps are `space × (n-1)`, the remainder splits evenly, and
    /// each column starts one `width + space` after the previous.
    #[test]
    fn equal_width_columns_split_the_remainder_after_gaps() {
        let g = cols_of(Columns {
            count: Some(3),
            ..base()
        });
        assert_eq!(g.len(), 3);
        // 468 - 36*2 = 396; 396/3 = 132.
        for c in &g {
            assert_eq!(c.width, Pt::new(132.0));
        }
        assert_eq!(g[0].x_offset, Pt::ZERO);
        assert_eq!(g[1].x_offset, Pt::new(168.0)); // 132 + 36
        assert_eq!(g[2].x_offset, Pt::new(336.0));
        // The last column ends exactly at the content edge.
        assert_eq!(g[2].x_offset + g[2].width, Pt::new(468.0));
    }

    #[test]
    fn explicit_space_overrides_the_default_gap() {
        let g = cols_of(Columns {
            count: Some(2),
            space: Some(Dimension::<Twips>::new(1440)), // 72pt
            ..base()
        });
        // 468 - 72 = 396; /2 = 198.
        assert_eq!(g[0].width, Pt::new(198.0));
        assert_eq!(g[1].x_offset, Pt::new(270.0)); // 198 + 72
    }

    /// §17.6.4: `equalWidth` defaults to **true**, so `<w:col>` children are
    /// ignored unless the file explicitly opts out. Word writes
    /// `equalWidth="0"` whenever it means them to win.
    #[test]
    fn col_children_are_ignored_unless_equal_width_is_explicitly_false() {
        let children = vec![
            ColumnDefinition {
                width: Some(Dimension::<Twips>::new(1440)), // 72pt
                space: None,
            },
            ColumnDefinition {
                width: Some(Dimension::<Twips>::new(7200)), // 360pt
                space: None,
            },
        ];
        // Absent `equalWidth` → equal widths, children ignored.
        let absent = cols_of(Columns {
            count: Some(2),
            columns: children.clone(),
            ..base()
        });
        assert_eq!(absent[0].width, absent[1].width, "absent means equal");
        assert_eq!(absent[0].width, Pt::new(216.0)); // (468-36)/2

        // Explicit opt-out → the individual widths apply.
        let opted_out = cols_of(Columns {
            count: Some(2),
            equal_width: Some(false),
            columns: children,
            ..base()
        });
        assert_eq!(opted_out[0].width, Pt::new(72.0));
        assert_eq!(opted_out[1].width, Pt::new(360.0));
        assert_eq!(opted_out[1].x_offset, Pt::new(108.0)); // 72 + 36 default gap
    }

    /// A per-column `w:space` overrides the default for that gap only, and the
    /// final column contributes no trailing gap.
    #[test]
    fn per_column_space_applies_to_its_own_gap_only() {
        let g = cols_of(Columns {
            count: Some(3),
            equal_width: Some(false),
            columns: vec![
                ColumnDefinition {
                    width: Some(Dimension::<Twips>::new(2000)), // 100pt
                    space: Some(Dimension::<Twips>::new(200)),  // 10pt
                },
                ColumnDefinition {
                    width: Some(Dimension::<Twips>::new(2000)),
                    space: None, // falls back to the 36pt default
                },
                ColumnDefinition {
                    width: Some(Dimension::<Twips>::new(2000)),
                    space: Some(Dimension::<Twips>::new(9999)), // trailing: unused
                },
            ],
            ..base()
        });
        assert_eq!(g[0].x_offset, Pt::ZERO);
        assert_eq!(g[1].x_offset, Pt::new(110.0)); // 100 + 10
        assert_eq!(g[2].x_offset, Pt::new(246.0)); // 110 + 100 + 36
    }

    /// §17.6.4: `w:num` is authoritative for the count. Fewer `<w:col>`
    /// children than `w:num` leaves the remainder at the default width instead
    /// of silently dropping columns.
    #[test]
    fn num_wins_when_col_children_are_missing_or_extra() {
        let short = cols_of(Columns {
            count: Some(3),
            equal_width: Some(false),
            columns: vec![ColumnDefinition {
                width: Some(Dimension::<Twips>::new(1440)), // 72pt
                space: None,
            }],
            ..base()
        });
        assert_eq!(short.len(), 3, "w:num=3 yields three columns");
        assert_eq!(short[0].width, Pt::new(72.0), "the declared one");
        assert_eq!(short[1].width, Pt::new(156.0), "defaulted: 468/3");
        assert_eq!(short[2].width, Pt::new(156.0));

        let extra = cols_of(Columns {
            count: Some(2),
            equal_width: Some(false),
            columns: vec![
                ColumnDefinition {
                    width: Some(Dimension::<Twips>::new(1440)),
                    space: None,
                },
                ColumnDefinition {
                    width: Some(Dimension::<Twips>::new(1440)),
                    space: None,
                },
                ColumnDefinition {
                    width: Some(Dimension::<Twips>::new(1440)),
                    space: None,
                },
            ],
            ..base()
        });
        assert_eq!(extra.len(), 2, "children beyond w:num are ignored");
    }

    /// `w:num` is a file-controlled `u32` that sizes an allocation. The ceiling
    /// comes from the page, so the reservation can never exceed what the text
    /// area could hold at `MIN_COLUMN_WIDTH`.
    #[test]
    fn absurd_column_counts_are_bounded_by_the_page() {
        let g = cols_of(Columns {
            count: Some(u32::MAX),
            ..base()
        });
        assert!(
            g.len() <= 468,
            "468pt of text area at 1pt minimum caps the count, got {}",
            g.len()
        );
        // A zero-width text area still yields a usable single column.
        let degenerate = compute_columns(
            Pt::ZERO,
            Some(&Columns {
                count: Some(u32::MAX),
                ..base()
            }),
        );
        assert_eq!(degenerate.len(), 1);
    }

    // ── E6#1: no route may produce a negative width ──────────────────────
    //
    // Three distinct routes reached the same defect, and only the first was in
    // the original finding. All three ended at `BoxConstraints::new`'s
    // `debug_assert!(min_width <= max_width)` — a panic through the public
    // `convert()` API in debug builds, a negative-width layout in release.

    /// Route 1: more columns than the text area can hold at the given spacing.
    /// The old bound (`content_width / MIN_COLUMN_WIDTH`) ignored the gaps, so
    /// 15 columns on Letter still produced `-25.9pt` each.
    #[test]
    fn column_count_never_yields_a_negative_width() {
        for n in [2u32, 13, 14, 15, 16, 50, 1000, u32::MAX] {
            let g = cols_of(Columns {
                count: Some(n),
                ..base()
            });
            assert!(!g.is_empty(), "n={n}");
            for (i, c) in g.iter().enumerate() {
                assert!(
                    c.width >= Pt::ZERO,
                    "n={n} col{i} width {:?} is negative",
                    c.width
                );
            }
        }
    }

    /// Route 2: margins wider than the page. This one never reaches the column
    /// arithmetic at all — a single-column section returns `content_width`
    /// verbatim — so clamping inside `compute_columns` alone would have missed
    /// it. `PageConfig::content_width` is where it has to be caught.
    #[test]
    fn margins_wider_than_the_page_clamp_to_zero_not_negative() {
        let sect = SectionProperties {
            page_size: Dup::from(Some(PageSize {
                width: Some(Dimension::<Twips>::new(2000)), // 100pt
                height: Some(Dimension::<Twips>::new(2000)),
                orientation: None,
            })),
            page_margins: Dup::from(Some(PageMargins {
                left: Some(Dimension::<Twips>::new(1440)),  // 72pt
                right: Some(Dimension::<Twips>::new(1440)), // 72pt — 144 > 100
                top: Some(Dimension::<Twips>::new(1440)),
                bottom: Some(Dimension::<Twips>::new(1440)),
                header: None,
                footer: None,
                gutter: None,
            })),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.content_width(), Pt::ZERO, "clamped, not -44pt");
        assert_eq!(cfg.content_height(), Pt::ZERO);
        assert!(cfg.columns.iter().all(|c| c.width >= Pt::ZERO));
    }

    /// Route 3: a spacing larger than the whole text area, at a column count
    /// that would otherwise be perfectly ordinary.
    #[test]
    fn spacing_wider_than_the_page_yields_one_column_not_negative_ones() {
        let g = cols_of(Columns {
            count: Some(2),
            space: Some(Dimension::<Twips>::new(20000)), // 1000pt on a 468pt area
            ..base()
        });
        assert_eq!(g.len(), 1, "two columns cannot fit a 1000pt gap");
        assert!(g[0].width >= Pt::ZERO);
    }

    /// The bound solves for the gaps, so every admitted count leaves each column
    /// at least `MIN_COLUMN_WIDTH` — not merely non-negative. On Letter with the
    /// default 36pt spacing that is 13 columns, one below the old sign-flip.
    #[test]
    fn admitted_counts_leave_every_column_usably_wide() {
        let g = cols_of(Columns {
            count: Some(13),
            ..base()
        });
        assert_eq!(g.len(), 13);
        for c in &g {
            assert!(c.width >= MIN_COLUMN_WIDTH, "got {:?}", c.width);
        }
        // 14 is where the gaps consume the area, so the count is reduced.
        let over = cols_of(Columns {
            count: Some(14),
            ..base()
        });
        assert_eq!(over.len(), 13, "clamped to what fits");
    }
}
