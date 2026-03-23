//! Layout constraint propagation.
//!
//! `LayoutConstraints` describes the spatial bounds available to a child element.
//! `LayoutContext` is a stack of constraints that tracks nesting as we walk the
//! document tree (page → table → cell → paragraph).
//!
//! Both the measure step and the layout step use the same constraint type.

use crate::dimension::Pt;
use crate::geometry::PtSize;
use crate::model::Indentation;

use super::LayoutConfig;

/// Spatial constraints passed from parent to child during layout.
///
/// Constraints flow **down** the tree: page → block → table cell → paragraph.
/// Each level may narrow the available space.
#[derive(Debug, Clone, Copy)]
pub struct LayoutConstraints {
    /// Left edge of the content area (absolute page x-coordinate).
    pub x_origin: Pt,
    /// Available width for content at this nesting level.
    pub available_width: Pt,
    /// Available height for content (distance from current y to bottom bound).
    pub available_height: Pt,
    /// Page size (carried for percentage-based positioning).
    pub page_size: PtSize,
}

impl LayoutConstraints {
    /// Create constraints for a full page from a layout config.
    pub(crate) fn for_page(config: &LayoutConfig) -> Self {
        Self {
            x_origin: config.margins.left,
            available_width: config.content_width(),
            available_height: config.content_height(),
            page_size: config.page_size,
        }
    }

    /// Narrow constraints for a paragraph with indentation.
    pub fn for_paragraph(&self, indent: &Indentation) -> Self {
        let left = indent.left.map(Pt::from).unwrap_or(Pt::ZERO);
        let right = indent.right.map(Pt::from).unwrap_or(Pt::ZERO);
        Self {
            x_origin: self.x_origin + left,
            available_width: (self.available_width - left - right).max(Pt::ZERO),
            available_height: self.available_height,
            page_size: self.page_size,
        }
    }

    /// Narrow constraints for a table cell.
    pub fn for_cell(&self, cell_x: Pt, cell_content_width: Pt, cell_height_limit: Pt) -> Self {
        Self {
            x_origin: cell_x,
            available_width: cell_content_width,
            available_height: cell_height_limit,
            page_size: self.page_size,
        }
    }

    /// Create constraints for header/footer layout.
    pub(crate) fn for_header_footer(config: &LayoutConfig, y_extent: Pt) -> Self {
        Self {
            x_origin: config.margins.left,
            available_width: config.content_width(),
            available_height: y_extent,
            page_size: config.page_size,
        }
    }

    /// Constraints with no practical limit — used as a fallback when
    /// actual constraints are not yet known.
    pub fn unconstrained() -> Self {
        Self {
            x_origin: Pt::ZERO,
            available_width: Pt::new(10000.0),
            available_height: Pt::new(10000.0),
            page_size: PtSize::new(Pt::new(10000.0), Pt::new(10000.0)),
        }
    }
}

/// A stack of layout constraints tracking parent–child nesting.
///
/// Push a frame when entering a child (table cell, paragraph), pop when leaving.
/// The current (top) frame provides the active constraints.
pub struct LayoutContext {
    stack: Vec<LayoutConstraints>,
}

impl LayoutContext {
    /// Create a new context with initial (root) constraints.
    pub fn new(root: LayoutConstraints) -> Self {
        Self { stack: vec![root] }
    }

    /// Current (innermost) constraints.
    pub fn current(&self) -> &LayoutConstraints {
        self.stack.last().expect("context stack must not be empty")
    }

    /// Push child constraints onto the stack.
    pub fn push(&mut self, constraints: LayoutConstraints) {
        self.stack.push(constraints);
    }

    /// Pop the current constraints, returning to the parent.
    /// Panics if attempting to pop the root frame.
    pub fn pop(&mut self) -> LayoutConstraints {
        assert!(self.stack.len() > 1, "cannot pop root constraints");
        self.stack.pop().unwrap()
    }

    /// Replace the root constraints (e.g., after a section break changes page geometry).
    /// Only valid when the stack is at depth 1 (no child frames active).
    pub fn replace_root(&mut self, constraints: LayoutConstraints) {
        assert_eq!(
            self.stack.len(),
            1,
            "can only replace root when stack is at depth 1"
        );
        self.stack[0] = constraints;
    }

    /// Current nesting depth (1 = root only).
    pub fn depth(&self) -> usize {
        self.stack.len()
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
    fn for_page_derives_from_config() {
        let config = test_config();
        let c = LayoutConstraints::for_page(&config);
        assert_eq!(c.x_origin, config.margins.left);
        assert_eq!(c.available_width, config.content_width());
        assert_eq!(c.available_height, config.content_height());
    }

    #[test]
    fn for_paragraph_narrows_by_indent() {
        let page = LayoutConstraints::for_page(&test_config());
        let indent = Indentation {
            left: Some(Twips::new(720)),  // 36pt
            right: Some(Twips::new(360)), // 18pt
            first_line: None,
        };
        let para = page.for_paragraph(&indent);
        assert!(para.x_origin > page.x_origin);
        assert!(para.available_width < page.available_width);
        let left_pt = Pt::from(Twips::new(720));
        let right_pt = Pt::from(Twips::new(360));
        assert_eq!(para.x_origin, page.x_origin + left_pt);
        assert_eq!(
            para.available_width,
            page.available_width - left_pt - right_pt
        );
    }

    #[test]
    fn for_cell_sets_origin_and_width() {
        let page = LayoutConstraints::for_page(&test_config());
        let cell = page.for_cell(Pt::new(100.0), Pt::new(200.0), Pt::new(500.0));
        assert_eq!(cell.x_origin, Pt::new(100.0));
        assert_eq!(cell.available_width, Pt::new(200.0));
        assert_eq!(cell.available_height, Pt::new(500.0));
        // page_size is inherited
        assert_eq!(cell.page_size, page.page_size);
    }

    #[test]
    fn context_push_pop() {
        let root = LayoutConstraints::for_page(&test_config());
        let mut ctx = LayoutContext::new(root);
        assert_eq!(ctx.depth(), 1);

        let child = root.for_cell(Pt::new(50.0), Pt::new(100.0), Pt::new(300.0));
        ctx.push(child);
        assert_eq!(ctx.depth(), 2);
        assert_eq!(ctx.current().x_origin, Pt::new(50.0));

        ctx.pop();
        assert_eq!(ctx.depth(), 1);
        assert_eq!(ctx.current().x_origin, root.x_origin);
    }

    #[test]
    fn context_replace_root() {
        let root = LayoutConstraints::for_page(&test_config());
        let mut ctx = LayoutContext::new(root);

        let new_root = LayoutConstraints {
            x_origin: Pt::new(100.0),
            ..root
        };
        ctx.replace_root(new_root);
        assert_eq!(ctx.current().x_origin, Pt::new(100.0));
    }

    #[test]
    #[should_panic(expected = "cannot pop root")]
    fn pop_root_panics() {
        let root = LayoutConstraints::for_page(&test_config());
        let mut ctx = LayoutContext::new(root);
        ctx.pop();
    }

    #[test]
    #[should_panic(expected = "can only replace root")]
    fn replace_root_with_children_panics() {
        let root = LayoutConstraints::for_page(&test_config());
        let mut ctx = LayoutContext::new(root);
        ctx.push(root);
        ctx.replace_root(root);
    }

    #[test]
    fn unconstrained_is_large() {
        let c = LayoutConstraints::unconstrained();
        assert!(f32::from(c.available_width) > 9000.0);
        assert!(f32::from(c.available_height) > 9000.0);
    }

    #[test]
    fn nested_three_levels() {
        let root = LayoutConstraints::for_page(&test_config());
        let mut ctx = LayoutContext::new(root);

        // Table cell
        let cell = root.for_cell(Pt::new(80.0), Pt::new(200.0), Pt::new(600.0));
        ctx.push(cell);
        assert_eq!(ctx.depth(), 2);

        // Paragraph inside cell
        let indent = Indentation {
            left: Some(Twips::new(240)),
            right: None,
            first_line: None,
        };
        let para = ctx.current().for_paragraph(&indent);
        ctx.push(para);
        assert_eq!(ctx.depth(), 3);
        assert!(ctx.current().available_width < cell.available_width);

        ctx.pop(); // leave paragraph
        assert_eq!(ctx.depth(), 2);
        ctx.pop(); // leave cell
        assert_eq!(ctx.depth(), 1);
    }
}
