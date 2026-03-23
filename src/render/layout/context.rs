//! Cascading layout constraint stack.
//!
//! `ConstraintFrame` holds optional overrides for each spatial property.
//! `LayoutConstraints` is a stack of frames — getters resolve by walking
//! top-to-bottom, returning the most recently set value for each field.
//!
//! The stack tracks nesting as we walk the document tree:
//! page → header/body → table → cell → paragraph.

use crate::dimension::Pt;
use crate::geometry::PtSize;
use crate::model::Indentation;

use super::LayoutConfig;

/// A single constraint layer. `None` fields inherit from the frame below.
#[derive(Debug, Clone, Copy, Default)]
pub struct ConstraintFrame {
    x_origin: Option<Pt>,
    available_width: Option<Pt>,
    available_height: Option<Pt>,
    page_size: Option<PtSize>,
}

impl ConstraintFrame {
    /// Create an empty frame (all fields inherited).
    pub fn new() -> Self {
        Self::default()
    }

    pub fn x_origin(mut self, v: Pt) -> Self {
        self.x_origin = Some(v);
        self
    }

    pub fn available_width(mut self, v: Pt) -> Self {
        self.available_width = Some(v);
        self
    }

    pub fn available_height(mut self, v: Pt) -> Self {
        self.available_height = Some(v);
        self
    }

    pub fn page_size(mut self, v: PtSize) -> Self {
        self.page_size = Some(v);
        self
    }
}

/// Cascading constraint stack.
///
/// Each nesting level pushes a frame that may override specific fields.
/// Getters walk the stack top-to-bottom and return the first defined value.
#[derive(Default)]
pub struct LayoutConstraints {
    stack: Vec<ConstraintFrame>,
}

impl LayoutConstraints {
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a constraint stack pre-loaded with page-level constraints.
    pub(crate) fn for_page(config: &LayoutConfig) -> Self {
        let mut c = Self::new();
        c.push_page(config);
        c
    }

    /// Create a constraint stack for header/footer layout.
    pub(crate) fn for_header_footer(config: &LayoutConfig, y_extent: Pt) -> Self {
        let mut c = Self::new();
        c.push(
            ConstraintFrame::new()
                .x_origin(config.margins.left)
                .available_width(config.content_width())
                .available_height(y_extent)
                .page_size(config.page_size),
        );
        c
    }

    // -- Stack operations --

    pub fn push(&mut self, frame: ConstraintFrame) {
        self.stack.push(frame);
    }

    pub fn pop(&mut self) {
        self.stack.pop();
    }

    // -- Convenience push methods --

    /// Push full page constraints (sets all fields).
    pub(crate) fn push_page(&mut self, config: &LayoutConfig) {
        self.push(
            ConstraintFrame::new()
                .x_origin(config.margins.left)
                .available_width(config.content_width())
                .available_height(config.content_height())
                .page_size(config.page_size),
        );
    }

    /// Push narrowed constraints for a table cell.
    pub fn push_cell(&mut self, x_origin: Pt, width: Pt, height: Pt) {
        self.push(
            ConstraintFrame::new()
                .x_origin(x_origin)
                .available_width(width)
                .available_height(height),
        );
    }

    /// Push narrowed constraints for a paragraph (indentation reduces width).
    pub fn push_paragraph(&mut self, indent: &Indentation) {
        let left = indent.left_pt();
        let right = indent.right_pt();
        self.push(
            ConstraintFrame::new()
                .x_origin(self.x_origin() + left)
                .available_width((self.available_width() - left - right).max(Pt::ZERO)),
        );
    }

    /// Clear the stack and push fresh page constraints (for section breaks).
    pub(crate) fn replace_page(&mut self, config: &LayoutConfig) {
        self.stack.clear();
        self.push_page(config);
    }

    // -- Getters (resolve from stack top to bottom) --

    /// Left edge of the content area (absolute page x-coordinate).
    pub fn x_origin(&self) -> Pt {
        self.resolve(|f| f.x_origin).unwrap_or(Pt::ZERO)
    }

    /// Available width for content at the current nesting level.
    pub fn available_width(&self) -> Pt {
        self.resolve(|f| f.available_width).unwrap_or(Pt::ZERO)
    }

    /// Available height for content.
    pub fn available_height(&self) -> Pt {
        self.resolve(|f| f.available_height).unwrap_or(Pt::ZERO)
    }

    /// Page size (carried for percentage-based positioning).
    pub fn page_size(&self) -> PtSize {
        self.resolve(|f| f.page_size).unwrap_or(PtSize::ZERO)
    }

    fn resolve<T: Copy>(&self, getter: impl Fn(&ConstraintFrame) -> Option<T>) -> Option<T> {
        self.stack.iter().rev().find_map(getter)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::dimension::{Pt, Twips};

    fn test_config() -> LayoutConfig {
        LayoutConfig::default()
    }

    #[test]
    fn page_constraints_set_all_fields() {
        let config = test_config();
        let c = LayoutConstraints::for_page(&config);
        assert_eq!(c.x_origin(), config.margins.left);
        assert_eq!(c.available_width(), config.content_width());
        assert_eq!(c.available_height(), config.content_height());
        assert_eq!(c.page_size(), config.page_size);
    }

    #[test]
    fn paragraph_narrows_by_indent() {
        let config = test_config();
        let mut c = LayoutConstraints::for_page(&config);
        let page_x = c.x_origin();
        let page_w = c.available_width();

        let indent = Indentation {
            left: Some(Twips::new(720)),  // 36pt
            right: Some(Twips::new(360)), // 18pt
            first_line: None,
        };
        c.push_paragraph(&indent);

        let left_pt = Pt::from(Twips::new(720));
        let right_pt = Pt::from(Twips::new(360));
        assert_eq!(c.x_origin(), page_x + left_pt);
        assert_eq!(c.available_width(), page_w - left_pt - right_pt);
        // Height and page_size inherited from page frame
        assert_eq!(c.available_height(), config.content_height());
        assert_eq!(c.page_size(), config.page_size);
    }

    #[test]
    fn cell_overrides_origin_and_dimensions() {
        let config = test_config();
        let mut c = LayoutConstraints::for_page(&config);
        c.push_cell(Pt::new(100.0), Pt::new(200.0), Pt::new(500.0));

        assert_eq!(c.x_origin(), Pt::new(100.0));
        assert_eq!(c.available_width(), Pt::new(200.0));
        assert_eq!(c.available_height(), Pt::new(500.0));
        // page_size inherited
        assert_eq!(c.page_size(), config.page_size);
    }

    #[test]
    fn push_pop_restores_parent() {
        let config = test_config();
        let mut c = LayoutConstraints::for_page(&config);
        let page_x = c.x_origin();

        c.push_cell(Pt::new(50.0), Pt::new(100.0), Pt::new(300.0));
        assert_eq!(c.x_origin(), Pt::new(50.0));

        c.pop();
        assert_eq!(c.x_origin(), page_x);
    }

    #[test]
    fn replace_page_resets_stack() {
        let config = test_config();
        let mut c = LayoutConstraints::for_page(&config);
        c.push_cell(Pt::new(50.0), Pt::new(100.0), Pt::new(300.0));

        let mut new_config = config;
        new_config.margins.left = Pt::new(100.0);
        c.replace_page(&new_config);
        assert_eq!(c.x_origin(), Pt::new(100.0));
    }

    #[test]
    fn partial_frame_inherits_from_below() {
        let config = test_config();
        let mut c = LayoutConstraints::for_page(&config);

        // Push a frame that only narrows height
        c.push(ConstraintFrame::new().available_height(Pt::new(300.0)));
        assert_eq!(c.available_height(), Pt::new(300.0));
        // Width and x_origin inherited from page
        assert_eq!(c.x_origin(), config.margins.left);
        assert_eq!(c.available_width(), config.content_width());
    }

    #[test]
    fn empty_stack_returns_defaults() {
        let c = LayoutConstraints::new();
        assert_eq!(c.x_origin(), Pt::ZERO);
        assert_eq!(c.available_width(), Pt::ZERO);
        assert_eq!(c.available_height(), Pt::ZERO);
    }

    #[test]
    fn nested_three_levels() {
        let config = test_config();
        let mut c = LayoutConstraints::for_page(&config);

        // Table cell
        c.push_cell(Pt::new(80.0), Pt::new(200.0), Pt::new(600.0));

        // Paragraph inside cell
        let indent = Indentation {
            left: Some(Twips::new(240)),
            right: None,
            first_line: None,
        };
        c.push_paragraph(&indent);
        assert!(c.available_width() < Pt::new(200.0));

        c.pop(); // leave paragraph
        c.pop(); // leave cell
    }

    #[test]
    fn header_footer_constraints() {
        let config = test_config();
        let c = LayoutConstraints::for_header_footer(&config, Pt::new(36.0));
        assert_eq!(c.x_origin(), config.margins.left);
        assert_eq!(c.available_width(), config.content_width());
        assert_eq!(c.available_height(), Pt::new(36.0));
        assert_eq!(c.page_size(), config.page_size);
    }
}
