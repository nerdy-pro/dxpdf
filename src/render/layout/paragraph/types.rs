//! Type definitions for paragraph layout.

use crate::i18n::bidi::BaseDirection;
use crate::model::Alignment;

use super::super::draw_command::DrawCommand;
use super::super::fragment::Fragment;
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::resolve::color::RgbColor;

/// §17.3.1.38: a resolved tab stop for layout.
#[derive(Clone, Debug)]
pub struct TabStopDef {
    /// Absolute position from paragraph left edge.
    pub position: Pt,
    /// §17.18.85 `ST_TabJc`: tab alignment (left, center, right, decimal).
    pub alignment: crate::model::TabAlignment,
    /// §17.18.86 `ST_TabTlc`: leader character (dot, hyphen, underscore, etc.).
    pub leader: crate::model::TabLeader,
}

/// Configuration for paragraph layout.
#[derive(Clone, Debug)]
pub struct ParagraphStyle {
    pub alignment: Alignment,
    /// §17.3.1.6 `w:bidi` — the paragraph's base embedding direction.
    ///
    /// Two things read it. UAX #9 level resolution takes it as the base level,
    /// which is what makes `Alignment::Start`/`End` and `indent_left`/
    /// `indent_right` below *logical* rather than physical: under
    /// [`BaseDirection::Rtl`] "start" is the right margin. See
    /// `line_emit::physical_edges` for the resolution and for why the model's
    /// already-logical naming (`Indentation::start`/`end`,
    /// `Alignment::Start`/`End`) made this a completion rather than a rewrite.
    pub base_direction: BaseDirection,
    pub space_before: Pt,
    pub space_after: Pt,
    /// §17.3.1.12 `w:ind/@w:start`. **Logical**: the inset at the edge text
    /// flows *from*, which is the right margin when `base_direction` is
    /// right-to-left. Named `left` because that is where it lands in every
    /// left-to-right paragraph, which is nearly all of them.
    pub indent_left: Pt,
    /// §17.3.1.12 `w:ind/@w:end` — logical, the mirror of `indent_left`.
    pub indent_right: Pt,
    pub indent_first_line: Pt,
    /// §17.3.1.33: see [`LineSpacingRule`].
    pub line_spacing: LineSpacingRule,
    /// §20.1.2.1.18: the enclosing shape text body's `a:normAutofit` shrink.
    /// [`ShapeAutoFit::NONE`](crate::render::layout::ShapeAutoFit::NONE) for
    /// every paragraph outside one.
    ///
    /// Its `@lnSpcReduction` half is applied to the *resolved* line height, not
    /// to `line_spacing` above; the reason is on
    /// [`ShapeAutoFit::scale_line_height`](crate::render::layout::ShapeAutoFit::scale_line_height).
    pub auto_fit: crate::render::layout::ShapeAutoFit,
    /// §17.3.1.38: custom tab stops.
    pub tabs: Vec<TabStopDef>,
    /// §17.15.1.25: the default tab-stop interval, used when no custom tab stop
    /// applies (from `w:settings/w:defaultTabStop`; spec default 720tw = 36pt).
    pub default_tab_stop: Pt,
    /// §17.3.2.20: the language in effect for this paragraph's own decisions,
    /// as a coarse bucket (English / writes-a-comma / writes-a-point).
    ///
    /// Resolved from the paragraph *mark*'s `w:rPr`, not from the runs inside
    /// the paragraph: a decimal stop is a `w:pPr` property, and a zone whose
    /// runs each declared their own `w:lang` would otherwise have no single
    /// answer. See `build::convert::paragraph_locale`. Region-blind by
    /// design — `decimal_separator` below is what actually answers §17.18.85
    /// for a `decimal` stop, since a region can diverge from its own
    /// language's bucket (`de-CH` vs `de-DE`).
    pub locale: crate::render::resolve::locale::Locale,
    /// §17.18.85 (issue #128): the character a `decimal` tab stop's zone
    /// aligns on, resolved from the same tag as `locale` above but through
    /// real CLDR data (`crate::i18n::decimal_separator_for_tag`) rather than
    /// `locale`'s primary-subtag-only bucket — the bucket alone can't tell
    /// `de-DE` from `de-CH`. Falls back to `locale.decimal_separator()` when
    /// ICU4X has nothing for the declared tag. See `build::convert::paragraph_locale`.
    pub decimal_separator: char,
    /// §17.3.1.19: this paragraph's PDF outline entry, when it is a heading.
    ///
    /// Carried on the style rather than derived at emission because the node ID
    /// is assigned once, in document order, while the paragraph is built — see
    /// [`OutlineHeading`](crate::render::layout::draw_command::OutlineHeading).
    /// `None` for body text, and for every heading outside the document body
    /// (headers, footers, notes), which are not positions in the document.
    pub outline: Option<crate::render::layout::draw_command::OutlineHeading>,
    /// Drop cap to render at the start of this paragraph.
    pub drop_cap: Option<DropCapInfo>,
    /// §17.3.1.24: paragraph borders.
    pub borders: Option<ParagraphBorderStyle>,
    /// §17.3.1.31: paragraph shading (background fill).
    pub shading: Option<RgbColor>,
    /// §17.3.1.15: keep this paragraph on the same page as the next.
    pub keep_next: bool,
    /// §17.3.1.14: keep all lines of this paragraph on one page where possible
    /// (suppresses across-page splitting of the paragraph).
    pub keep_lines: bool,
    /// §17.3.1.44: when the paragraph *does* split, forbid a single line stranded
    /// at the bottom (orphan) or top (widow) of a page. Word's default is on.
    pub widow_control: bool,
    /// §17.3.1.9: suppress spacing between paragraphs of the same style.
    pub contextual_spacing: bool,
    /// Style ID for contextual spacing comparison.
    pub style_id: Option<crate::model::StyleId>,
    /// Active floats for per-line width adjustment.
    pub page_floats: Vec<super::super::float::ActiveFloat>,
    /// Absolute y position of this paragraph on the page (for float overlap checks).
    pub page_y: Pt,
    /// Left margin x position (for float_adjustments computation).
    pub page_x: Pt,
    /// Total content width (for float_adjustments computation).
    pub page_content_width: Pt,
}

/// Resolved paragraph border style for rendering.
#[derive(Clone, Debug, PartialEq)]
pub struct ParagraphBorderStyle {
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
    pub right: Option<BorderLine>,
}

/// A single border line for rendering.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct BorderLine {
    pub width: Pt,
    pub color: RgbColor,
    /// §17.3.1.24: space between border and text in points.
    pub space: Pt,
}

/// Drop cap letter to float at the start of a paragraph.
#[derive(Clone, Debug)]
pub struct DropCapInfo {
    /// The drop cap fragments (usually a single large letter).
    pub fragments: Vec<Fragment>,
    /// §17.3.1.11 @w:lines: number of body text lines the drop cap spans.
    pub lines: u32,
    /// Total width of the drop cap (measured).
    pub width: Pt,
    /// Total height of the drop cap (measured).
    pub height: Pt,
    /// Ascent of the drop cap font (for baseline positioning).
    pub ascent: Pt,
    /// §17.3.1.11 @w:hSpace: horizontal distance from surrounding text.
    pub h_space: Pt,
    /// §17.3.1.11: true = Margin mode (drop cap in margin), false = Drop mode (in text area).
    pub margin_mode: bool,
    /// Left indent of the drop cap paragraph (from its own style cascade).
    pub indent: Pt,
    /// Frame height from the drop cap paragraph's spacing (lineRule="exact").
    pub frame_height: Option<Pt>,
    /// §17.3.2.19: vertical baseline offset in points (negative = down).
    pub position_offset: Pt,
}

impl ParagraphStyle {
    /// Clone the style for layout, leaving `page_floats` empty.
    ///
    /// The caller always replaces `page_floats` with computed effective floats,
    /// so cloning the (potentially large) source vec is wasted work.
    pub fn clone_for_layout(&self) -> Self {
        Self {
            alignment: self.alignment,
            base_direction: self.base_direction,
            space_before: self.space_before,
            space_after: self.space_after,
            indent_left: self.indent_left,
            indent_right: self.indent_right,
            indent_first_line: self.indent_first_line,
            line_spacing: self.line_spacing,
            auto_fit: self.auto_fit,
            tabs: self.tabs.clone(),
            default_tab_stop: self.default_tab_stop,
            locale: self.locale,
            decimal_separator: self.decimal_separator,
            outline: self.outline.clone(),
            drop_cap: self.drop_cap.clone(),
            borders: self.borders.clone(),
            shading: self.shading,
            keep_next: self.keep_next,
            keep_lines: self.keep_lines,
            widow_control: self.widow_control,
            contextual_spacing: self.contextual_spacing,
            style_id: self.style_id.clone(),
            // Caller will replace these — skip cloning the vec.
            page_floats: Vec::new(),
            page_y: Pt::ZERO,
            page_x: Pt::ZERO,
            page_content_width: Pt::ZERO,
        }
    }
}

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            alignment: Alignment::Start,
            base_direction: BaseDirection::Ltr,
            space_before: Pt::ZERO,
            space_after: Pt::ZERO,
            indent_left: Pt::ZERO,
            indent_right: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            auto_fit: crate::render::layout::ShapeAutoFit::NONE,
            line_spacing: LineSpacingRule::Auto(1.0),
            tabs: Vec::new(),
            // §17.15.1.25: 720 twips = 36pt is the OOXML default interval.
            default_tab_stop: Pt::new(36.0),
            locale: crate::render::resolve::locale::Locale::English,
            decimal_separator: crate::render::resolve::locale::Locale::English.decimal_separator(),
            outline: None,
            drop_cap: None,
            borders: None,
            shading: None,
            keep_next: false,
            keep_lines: false,
            // §17.3.1.44: Word enables widow/orphan control by default.
            widow_control: true,
            contextual_spacing: false,
            style_id: None,
            page_floats: Vec::new(),
            page_y: Pt::ZERO,
            page_x: Pt::ZERO,
            page_content_width: Pt::ZERO,
        }
    }
}

// Workaround for clippy: ParagraphStyle has many fields but they all map 1:1 to spec properties.
// A builder pattern would add complexity without value here.

/// §17.3.1.33: line spacing rules matching OOXML semantics.
#[derive(Clone, Copy, Debug)]
pub enum LineSpacingRule {
    /// Proportional: multiplier on natural line height (1.0 = single, 1.5 = 1.5x, etc.)
    Auto(f32),
    /// Exact line height in points. Taller content doesn't clip — it visually
    /// overruns into the neighboring line.
    Exact(Pt),
    /// Minimum line height in points.
    AtLeast(Pt),
}

/// Result of laying out a paragraph.
#[derive(Debug)]
pub struct ParagraphLayout {
    /// Draw commands positioned relative to the paragraph's top-left origin.
    pub commands: Vec<DrawCommand>,
    /// Total size consumed by this paragraph (including spacing).
    pub size: PtSize,
}

/// Optional measuring capability for accurate per-character splitting — the
/// same trait object as `fragment::split::MeasureFn`, aliased so the two
/// layers keep passing one value. See [`SplitMeasure`] for the two
/// operations and why a shaped piece needs the second.
pub type MeasureTextFn<'a> = Option<&'a dyn SplitMeasure>;

pub use super::super::fragment::SplitMeasure;
