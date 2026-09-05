//! Pure data types for section layout.

use super::super::draw_command::{LayoutedPage, ResolvedEffect, ResolvedFill, ResolvedStroke};
use super::super::fragment::Fragment;
use super::super::paragraph::{ParagraphBorderStyle, ParagraphStyle};
use super::super::table::TableRowInput;
use crate::model::dimension::{Dimension, SixtieThousandthDeg};
use crate::model::{StyleId, WrapText};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::resolve::images::MediaEntry;
use crate::render::resolve::shape_geometry::SubPath;

/// Layout-resolved text wrap mode for a floating drawing.
///
/// Derived from the model's `TextWrap` (§20.4.2.14-18) — carries the
/// per-line side constraint (`wrap_text`) but strips distances, which are
/// baked into the float's resolved page-coordinate rectangle before
/// registration.
///
/// Tier 1 treats `Tight` and `Through` as spec-compliant placeholders: both
/// take `Square`'s path exactly — registered as an active float
/// ([`WrapMode::registers_as_wrap_float`]) and narrowing each line by the
/// float's bounding rect and the side constraint, never by the polygon the
/// two modes actually name. Full polygon-aware line fitting is deferred to
/// Tier 2 and not yet designed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WrapMode {
    /// §20.4.2.15 wrapNone — no reflow; drawing paints over or under text.
    None,
    /// §20.4.2.17 wrapSquare — text wraps around drawing's bounding box.
    Square(WrapText),
    /// §20.4.2.16 wrapTight — text follows drawing's polygon outline.
    Tight(WrapText),
    /// §20.4.2.14 wrapThrough — text flows through polygon interior.
    Through(WrapText),
    /// §20.4.2.18 wrapTopAndBottom — text stops at drawing top and
    /// resumes below drawing bottom.
    TopAndBottom,
}

impl WrapMode {
    /// Construct a layout wrap mode from a parsed `TextWrap`.
    pub fn from_model(wrap: &crate::model::TextWrap) -> Self {
        use crate::model::TextWrap as T;
        match wrap {
            T::None => Self::None,
            T::Square { wrap_text, .. } => Self::Square(*wrap_text),
            T::Tight { wrap_text, .. } => Self::Tight(*wrap_text),
            T::Through { wrap_text, .. } => Self::Through(*wrap_text),
            T::TopAndBottom { .. } => Self::TopAndBottom,
        }
    }

    /// Whether the drawing participates in wrap-around text flow (i.e.,
    /// should be registered as an active float). `None` and
    /// `TopAndBottom` don't — None is purely overlay; TopAndBottom is
    /// handled by advancing `cursor_y` past the drawing on emit.
    pub fn registers_as_wrap_float(self) -> bool {
        matches!(self, Self::Square(_) | Self::Tight(_) | Self::Through(_))
    }

    /// Side constraint for line narrowing. Returns `WrapText::BothSides`
    /// for modes that don't have one (None / TopAndBottom).
    pub fn wrap_text(self) -> WrapText {
        match self {
            Self::Square(w) | Self::Tight(w) | Self::Through(w) => w,
            Self::None | Self::TopAndBottom => WrapText::BothSides,
        }
    }
}

/// §17.4.58: positioning data for a floating table.
#[derive(Debug, Clone)]
pub struct TableFloatInfo {
    /// Gap between the table's right edge and surrounding text.
    pub right_gap: Pt,
    /// Gap between the table's bottom edge and surrounding text.
    pub bottom_gap: Pt,
    /// §17.4.58: horizontal alignment override (tblpXSpec).
    pub x_align: Option<crate::model::TableXAlign>,
    /// §17.4.58: absolute Y offset from the vertical anchor.
    pub y_offset: Pt,
    /// §17.4.58: vertical anchor reference (text / margin / page).
    pub vert_anchor: crate::model::TableAnchor,
    /// §17.4.57 `<w:tblOverlap>` — when present and set to `Never`,
    /// the layout shifts this table down past prior floating tables
    /// on the same page rather than letting them overlap. `None` /
    /// `Some(Overlap)` mean overlap is permitted (the §17.4.57
    /// default).
    pub overlap: Option<crate::model::TableOverlap>,
}

/// A floating (anchor) image to be positioned absolutely on the page.
#[derive(Clone)]
pub struct FloatingImage {
    pub image_data: MediaEntry,
    pub size: PtSize,
    /// §20.1.10.48 `a:srcRect` — fractional source crop in `[0, 1]`.
    pub src_rect: Option<crate::render::geometry::PtRect>,
    /// Horizontal position, possibly deferred to the page's parity.
    pub x: FloatingImageX,
    /// Resolved absolute y position on the page (may be relative to paragraph).
    pub y: FloatingImageY,
    /// §20.4.2.14-18: text wrap mode (drives float registration + cursor advance).
    pub wrap_mode: WrapMode,
    /// §20.4.2.3 distL/distR: horizontal distance from surrounding text.
    pub dist_left: Pt,
    pub dist_right: Pt,
    /// §20.4.2.3 @behindDoc: image is painted behind document text.
    pub behind_doc: bool,
}

impl FloatingImage {
    /// §20.4.2.18 convenience — kept for call-site ergonomics during the
    /// wrap-mode rollout. Equivalent to `matches!(self.wrap_mode, WrapMode::TopAndBottom)`.
    pub fn is_wrap_top_and_bottom(&self) -> bool {
        matches!(self.wrap_mode, WrapMode::TopAndBottom)
    }
}

/// Vertical position for a floating image.
///
/// Three cases, and the third is the one that carries the page: §20.4.3.2
/// `inside`/`outside` — both as an alignment and as the §20.4.3.5
/// `insideMargin` / `outsideMargin` references — mirror **top and bottom** with
/// the page's parity, the way the horizontal axis mirrors left and right (see
/// [`FloatingImageX`]). Floats are extracted before pagination, so the reading
/// cannot be chosen at build time: both are carried, and [`FloatingImageY::at`]
/// picks one once a page is assigned.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum FloatingImageY {
    /// Absolute page position, the same on every page.
    Absolute(Pt),
    /// Absolute page position, mirrored — which reading applies depends on the
    /// page's [`PageParity`].
    PageParity { odd: Pt, even: Pt },
    /// Relative to the paragraph's y position (offset added to cursor_y).
    RelativeToParagraph(Pt),
}

impl FloatingImageY {
    /// Build an absolute position from its two per-parity readings, collapsing
    /// to `Absolute` when they agree.
    ///
    /// They agree for every anchor that is not `inside`/`outside`, so a
    /// document with no mirrored anchor carries no deferral at all — the same
    /// property [`FloatingImageX::from_pages`] has on the other axis.
    pub fn absolute_from_pages(odd: Pt, even: Pt) -> Self {
        if odd == even {
            Self::Absolute(odd)
        } else {
            Self::PageParity { odd, even }
        }
    }

    /// The page-space y this object takes on a page of the given parity.
    ///
    /// `paragraph_top` is where the owning paragraph's content starts; only
    /// [`Self::RelativeToParagraph`] reads it, and the two absolute cases
    /// ignore it entirely.
    pub fn at(self, parity: PageParity, paragraph_top: Pt) -> Pt {
        match self {
            Self::Absolute(y) => y,
            Self::PageParity { odd, even } => match parity {
                PageParity::Odd => odd,
                PageParity::Even => even,
            },
            Self::RelativeToParagraph(offset) => paragraph_top + offset,
        }
    }

    /// The page-space y when the position is page-absolute, `None` when it is
    /// anchored to its paragraph.
    ///
    /// The distinction is behavioural, not merely arithmetic: a page-absolute
    /// float is placed by its anchor, so it ignores the §20.4.2.18
    /// `wrapTopAndBottom` band and the paragraph cursor that [`Self::at`]
    /// would otherwise fold in.
    pub fn absolute(self, parity: PageParity) -> Option<Pt> {
        match self {
            Self::Absolute(_) | Self::PageParity { .. } => Some(self.at(parity, Pt::ZERO)),
            Self::RelativeToParagraph(_) => None,
        }
    }

    /// Whether the position is page-absolute — the same question as
    /// [`Self::absolute`], for callers that have no page to ask about.
    pub fn is_page_absolute(self) -> bool {
        !matches!(self, Self::RelativeToParagraph(_))
    }
}

/// §20.4.3.1: which side of a two-sided document a page falls on.
///
/// Keyed on the **logical** page number — the one `w:pgNumType/@start` shifts,
/// and the one §17.10.6 `evenAndOddHeaders` already selects headers by. Sharing
/// the key keeps a document self-consistent: the page that gets the "even"
/// header mirrors the same way.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageParity {
    Odd,
    Even,
}

impl PageParity {
    /// From a 1-based logical page number.
    pub fn of_page(logical_page_number: usize) -> Self {
        if logical_page_number.is_multiple_of(2) {
            Self::Even
        } else {
            Self::Odd
        }
    }
}

/// Horizontal position for a floating object.
///
/// `Absolute` is settled during build. `PageParity` cannot be: §20.4.3.1
/// `inside`/`outside` — both as an alignment and as the `insideMargin` /
/// `outsideMargin` references — mirror according to the page the object lands
/// on, and floats are extracted *before* pagination. So the position is carried
/// as both readings and resolved when a page is assigned.
///
/// This mirrors [`FloatingImageY::RelativeToParagraph`], which exists for the
/// same reason on the other axis: a coordinate the build phase cannot know is
/// deferred, not guessed.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum FloatingImageX {
    /// The same x on every page.
    Absolute(Pt),
    /// Mirrored — which reading applies depends on the page's [`PageParity`].
    PageParity { odd: Pt, even: Pt },
}

impl FloatingImageX {
    /// Build from the two per-parity readings, collapsing to `Absolute` when
    /// they agree.
    ///
    /// They agree for every anchor that is not `inside`/`outside`, which is
    /// nearly all of them — so a document with no mirrored anchor carries no
    /// deferral at all, and the resolver below is a no-op on it.
    pub fn from_pages(odd: Pt, even: Pt) -> Self {
        if odd == even {
            Self::Absolute(odd)
        } else {
            Self::PageParity { odd, even }
        }
    }

    /// The x this object takes on a page of the given parity.
    pub fn resolve(self, parity: PageParity) -> Pt {
        match self {
            Self::Absolute(x) => x,
            Self::PageParity { odd, even } => match parity {
                PageParity::Odd => odd,
                PageParity::Even => even,
            },
        }
    }
}

/// A floating (anchor) DrawingML shape to be positioned absolutely on the
/// page. The geometry is already evaluated into path-local Pt subpaths; the
/// painter applies origin + rotation + flip.
///
/// Parallels `FloatingImage` — the stacker treats them the same way for
/// placement, differing only in the emitted `DrawCommand` variant.
#[derive(Clone)]
pub struct FloatingShape {
    /// Horizontal position of the shape's top-left, possibly deferred to the
    /// page's parity.
    pub x: FloatingImageX,
    /// Resolved absolute y position on the page.
    pub y: FloatingImageY,
    /// Shape bounding-box size in Pt.
    pub size: PtSize,
    /// §20.1.7.6 @rot — rotation around the shape's center, in 60000ths
    /// of a degree.
    pub rotation: Dimension<SixtieThousandthDeg>,
    /// §20.1.7.6 @flipH — mirror horizontally.
    pub flip_h: bool,
    /// §20.1.7.6 @flipV — mirror vertically.
    pub flip_v: bool,
    /// §20.4.2.14-18: text wrap mode.
    pub wrap_mode: WrapMode,
    /// §20.4.2.3 distL/distR — horizontal distance from surrounding text.
    pub dist_left: Pt,
    pub dist_right: Pt,
    /// §20.4.2.3 @behindDoc — painted behind document text.
    pub behind_doc: bool,
    /// Path subpaths in shape-local Pt (already scaled into `size`).
    pub paths: Vec<SubPath>,
    /// Resolved fill.
    pub fill: ResolvedFill,
    /// Optional resolved stroke.
    pub stroke: Option<ResolvedStroke>,
    /// Resolved post-processing effects (painter may defer all in Tier 0).
    pub effects: Vec<ResolvedEffect>,
    /// §17.17.1 / §20.1.2.1.1: pre-laid-out commands for the shape's
    /// text-box content (`wps:wsp/wps:txbx/w:txbxContent`). Each
    /// command is in shape-local coordinates; the stacker shifts
    /// them by the shape's resolved origin (plus body insets, pre-
    /// applied during sub-layout) and emits them *after* the shape's
    /// path so the text overlays the fill. Empty for shapes without
    /// text-box content or whose anchor model isn't supported by the
    /// sub-layout code yet (page-anchored vertical positions).
    pub text_commands: Vec<crate::render::layout::draw_command::DrawCommand>,
}

impl FloatingShape {
    /// §20.4.2.18 convenience — true when wrap_mode is `TopAndBottom`.
    pub fn is_wrap_top_and_bottom(&self) -> bool {
        matches!(self.wrap_mode, WrapMode::TopAndBottom)
    }
}

/// A block ready for layout — either a paragraph or a table.
///
/// The `Paragraph` variant is intentionally larger than `Table` (it carries
/// fragments, floats, and resolved style inline). Boxing would add an
/// allocation per block without a measurable benefit for this codebase,
/// where the Vec holding these is the dominant allocation.
#[allow(clippy::large_enum_variant)]
pub enum LayoutBlock {
    Paragraph {
        fragments: Vec<Fragment>,
        style: ParagraphStyle,
        /// §17.3.1.23: force a page break before this paragraph.
        page_break_before: bool,
        /// Footnotes referenced in this paragraph — rendered at page bottom.
        footnotes: Vec<(Vec<Fragment>, ParagraphStyle)>,
        /// §20.4.2.3: floating images anchored to this paragraph.
        floating_images: Vec<FloatingImage>,
        /// §14.5 / §20.1.2.2.35: floating DrawingML shapes anchored to this paragraph.
        floating_shapes: Vec<FloatingShape>,
    },
    Table {
        rows: Vec<TableRowInput>,
        /// Grid slot widths, already shrunk by `cell_spacing` so the slots plus
        /// one spacing sum to the table's own width.
        col_widths: Vec<Pt>,
        /// §17.4.44 `tblCellSpacing` resolved to points; zero when unset.
        cell_spacing: Pt,
        /// §17.4.38: resolved table border configuration.
        border_config: Option<super::super::table::TableBorderConfig>,
        /// §17.4.51: table indentation from left margin.
        indent: Pt,
        /// §17.4.28: table horizontal alignment.
        alignment: Option<crate::model::Alignment>,
        /// §17.4.1 `bidiVisual`: the table's own direction, which decides which
        /// margin `indent` and `alignment` are measured from. `None` when the
        /// table declares none, in which case the section's §17.6.6 direction
        /// answers instead — `section::helpers::table_x_offset` takes the one
        /// answer, and only its callers know both halves.
        direction: Option<crate::i18n::bidi::BaseDirection>,
        /// §17.4.58: floating table positioning — if present, text wraps around it.
        float_info: Option<TableFloatInfo>,
        /// §17.4.38: table style reference for adjacent table border collapse.
        style_id: Option<crate::model::StyleId>,
    },
}

/// §17.6.22: continuation state for `Continuous` section breaks.
/// Allows a new section to continue on the current page.
///
/// Everything the succeeding section needs in order to behave as though the
/// break were not there. A page plus a cursor is not enough: the break is not a
/// page boundary, so the per-page state that survives *within* a page has to
/// survive across it too. Each field below fixes a defect that a thinner
/// version had.
///
/// `Clone` because fitting a shared page may lay its section out more than
/// once (`fit_shared_page`), and each pass consumes the state it starts from.
#[derive(Clone)]
pub struct ContinuationState {
    pub page: LayoutedPage,
    /// Flow position, not the extent of the last drawn command.
    pub cursor_y: Pt,
    /// §17.11.23: the page bottom **already reduced** by footnotes reserved
    /// before the break. Taking the page's unreduced bottom lets the succeeding
    /// section lay text over the note block.
    pub bottom: Pt,
    /// §17.11.23: notes referenced before the break, deliberately **not yet
    /// flushed**. A page carries one separator above all its notes, so they are
    /// rendered together with whatever the succeeding section adds. Owned
    /// rather than borrowed because the blocks they came from belong to a
    /// section that is already finished.
    pub page_footnotes: Vec<(Vec<Fragment>, ParagraphStyle)>,
    /// §20.4.2: floats registered on this page. Text after the break wraps
    /// around them like any other text on the page.
    pub page_floats: Vec<super::super::float::ActiveFloat>,
    /// §17.3.1.9: what the last paragraph before the break left behind, so
    /// spacing collapses across it.
    pub prev_space_after: Pt,
    pub prev_style_id: Option<StyleId>,
    pub prev_borders: Option<ParagraphBorderStyle>,
    /// §17.4.38: for adjacent-table border collapse across the break.
    pub prev_table_style_id: Option<StyleId>,
    /// §17.4.58: anchor for a floating table in the succeeding section.
    pub last_para_start_y: Pt,
    /// §17.3.3.1: an inline page break at the end of the preceding section is
    /// deferred to the next block — which is in the succeeding section.
    pub pending_page_break: bool,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::Emu;
    use crate::model::geometry::EdgeInsets;
    use crate::model::TextWrap;

    fn insets() -> EdgeInsets<Emu> {
        EdgeInsets::new(
            Dimension::new(0),
            Dimension::new(0),
            Dimension::new(0),
            Dimension::new(0),
        )
    }

    /// The model → layout seam must be **total**: every §20.4.2.14-18 wrap
    /// type maps to a distinct `WrapMode`, and the per-line side constraint
    /// survives the conversion (it is the only part of `TextWrap` layout
    /// still needs — distances are baked into the resolved rect).
    #[test]
    fn from_model_maps_every_wrap_type() {
        assert_eq!(WrapMode::from_model(&TextWrap::None), WrapMode::None);
        assert_eq!(
            WrapMode::from_model(&TextWrap::Square {
                distance: insets(),
                wrap_text: WrapText::Left,
            }),
            WrapMode::Square(WrapText::Left),
        );
        assert_eq!(
            WrapMode::from_model(&TextWrap::Tight {
                distance: insets(),
                wrap_text: WrapText::Right,
                polygon: None,
            }),
            WrapMode::Tight(WrapText::Right),
        );
        assert_eq!(
            WrapMode::from_model(&TextWrap::Through {
                distance: insets(),
                wrap_text: WrapText::Largest,
                polygon: None,
            }),
            WrapMode::Through(WrapText::Largest),
        );
        assert_eq!(
            WrapMode::from_model(&TextWrap::TopAndBottom {
                distance_top: Dimension::new(0),
                distance_bottom: Dimension::new(0),
            }),
            WrapMode::TopAndBottom,
        );
    }

    /// The predicate that decides whether a drawing narrows surrounding text.
    /// §20.4.2.15 `wrapNone` is a pure overlay and §20.4.2.18 `wrapTopAndBottom`
    /// is a block spacer — neither participates. Callers that bypassed this and
    /// used a bare `else` instead let `wrapNone` images reflow text.
    #[test]
    fn only_wrap_enabled_modes_register_as_floats() {
        for mode in [
            WrapMode::Square(WrapText::BothSides),
            WrapMode::Tight(WrapText::BothSides),
            WrapMode::Through(WrapText::BothSides),
        ] {
            assert!(mode.registers_as_wrap_float(), "{mode:?} narrows text");
        }
        for mode in [WrapMode::None, WrapMode::TopAndBottom] {
            assert!(
                !mode.registers_as_wrap_float(),
                "{mode:?} must not narrow text"
            );
        }
    }

    /// The side constraint round-trips for wrap-enabled modes; the modes that
    /// have none report `BothSides`. That fallback is only meaningful because
    /// non-registering modes never reach a line-narrowing call site — see
    /// `only_wrap_enabled_modes_register_as_floats`.
    #[test]
    fn wrap_text_round_trips_and_defaults_to_both_sides() {
        for side in [
            WrapText::BothSides,
            WrapText::Left,
            WrapText::Right,
            WrapText::Largest,
        ] {
            assert_eq!(WrapMode::Square(side).wrap_text(), side);
            assert_eq!(WrapMode::Tight(side).wrap_text(), side);
            assert_eq!(WrapMode::Through(side).wrap_text(), side);
        }
        assert_eq!(WrapMode::None.wrap_text(), WrapText::BothSides);
        assert_eq!(WrapMode::TopAndBottom.wrap_text(), WrapText::BothSides);
    }
}
