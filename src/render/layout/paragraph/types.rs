//! Type definitions for paragraph layout.

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
    /// §17.18.81: tab alignment (left, center, right, decimal).
    pub alignment: crate::model::TabAlignment,
    /// §17.18.82: leader character (dot, hyphen, underscore, etc.).
    pub leader: crate::model::TabLeader,
}

/// Configuration for paragraph layout.
#[derive(Clone, Debug)]
pub struct ParagraphStyle {
    pub alignment: Alignment,
    pub space_before: Pt,
    pub space_after: Pt,
    pub indent_left: Pt,
    pub indent_right: Pt,
    pub indent_first_line: Pt,
    pub line_spacing: LineSpacingRule,
    /// §17.3.1.38: custom tab stops.
    pub tabs: Vec<TabStopDef>,
    /// Drop cap to render at the start of this paragraph.
    pub drop_cap: Option<DropCapInfo>,
    /// §17.3.1.24: paragraph borders.
    pub borders: Option<ParagraphBorderStyle>,
    /// §17.3.1.31: paragraph shading (background fill).
    pub shading: Option<RgbColor>,
    /// §17.3.1.14: keep this paragraph on the same page as the next.
    pub keep_next: bool,
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

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            alignment: Alignment::Start,
            space_before: Pt::ZERO,
            space_after: Pt::ZERO,
            indent_left: Pt::ZERO,
            indent_right: Pt::ZERO,
            indent_first_line: Pt::ZERO,
            line_spacing: LineSpacingRule::Auto(1.0),
            tabs: Vec::new(),
            drop_cap: None,
            borders: None,
            shading: None,
            keep_next: false,
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

/// Line spacing rules matching OOXML semantics.
#[derive(Clone, Copy, Debug)]
pub enum LineSpacingRule {
    /// Proportional: multiplier on natural line height (1.0 = single, 1.5 = 1.5x, etc.)
    Auto(f32),
    /// Exact line height in points.
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

/// Optional text measurement callback for accurate per-character splitting.
pub type MeasureTextFn<'a> = Option<
    &'a dyn Fn(
        &str,
        &super::super::fragment::FontProps,
    ) -> (Pt, super::super::fragment::TextMetrics),
>;
