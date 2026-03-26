//! Flutter-inspired constraint-based layout engine.
//!
//! Core protocol: **constraints go down, sizes go up, parent sets position**.

pub mod draw_command;
pub mod fragment;

use crate::dimension::Pt;
use crate::geometry::{PtEdgeInsets, PtSize};

/// Box constraints passed from parent to child during layout.
///
/// Encodes the range of permissible widths and heights.
/// A child's `perform_layout` must return a size within these bounds.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct BoxConstraints {
    pub min_width: Pt,
    pub max_width: Pt,
    pub min_height: Pt,
    pub max_height: Pt,
}

impl BoxConstraints {
    /// Create constraints with explicit bounds.
    pub fn new(min_width: Pt, max_width: Pt, min_height: Pt, max_height: Pt) -> Self {
        debug_assert!(min_width.raw() <= max_width.raw());
        debug_assert!(min_height.raw() <= max_height.raw());
        Self {
            min_width,
            max_width,
            min_height,
            max_height,
        }
    }

    /// Tight constraints — child must be exactly this size.
    pub fn tight(size: PtSize) -> Self {
        Self {
            min_width: size.width,
            max_width: size.width,
            min_height: size.height,
            max_height: size.height,
        }
    }

    /// Tight width, loose height — child must be exactly this wide, any height up to max.
    pub fn tight_width(width: Pt, max_height: Pt) -> Self {
        Self {
            min_width: width,
            max_width: width,
            min_height: Pt::ZERO,
            max_height,
        }
    }

    /// Loose constraints — child can be 0..max_size.
    pub fn loose(max_size: PtSize) -> Self {
        Self {
            min_width: Pt::ZERO,
            max_width: max_size.width,
            min_height: Pt::ZERO,
            max_height: max_size.height,
        }
    }

    /// Unbounded constraints — child can be any size.
    pub fn unbounded() -> Self {
        Self {
            min_width: Pt::ZERO,
            max_width: Pt::INFINITY,
            min_height: Pt::ZERO,
            max_height: Pt::INFINITY,
        }
    }

    /// Whether width is tight (min == max).
    pub fn is_tight_width(&self) -> bool {
        self.min_width == self.max_width
    }

    /// Whether height is tight (min == max).
    pub fn is_tight_height(&self) -> bool {
        self.min_height == self.max_height
    }

    /// Whether both axes are tight.
    pub fn is_tight(&self) -> bool {
        self.is_tight_width() && self.is_tight_height()
    }

    /// Intersect with another set of constraints — the result satisfies both.
    /// Used when nesting containers that each impose their own limits.
    pub fn enforce(&self, other: &BoxConstraints) -> BoxConstraints {
        BoxConstraints {
            min_width: self.min_width.max(other.min_width).min(self.max_width),
            max_width: self.max_width.min(other.max_width).max(self.min_width),
            min_height: self.min_height.max(other.min_height).min(self.max_height),
            max_height: self.max_height.min(other.max_height).max(self.min_height),
        }
    }

    /// Subtract edge insets from the constraints — shrinks the available space.
    /// Used when adding padding, margins, or cell insets.
    pub fn deflate(&self, edges: &PtEdgeInsets) -> BoxConstraints {
        let h = edges.horizontal();
        let v = edges.vertical();
        BoxConstraints {
            min_width: (self.min_width - h).max(Pt::ZERO),
            max_width: (self.max_width - h).max(Pt::ZERO),
            min_height: (self.min_height - v).max(Pt::ZERO),
            max_height: (self.max_height - v).max(Pt::ZERO),
        }
    }

    /// Clamp a size to fit within these constraints.
    pub fn constrain(&self, size: PtSize) -> PtSize {
        PtSize {
            width: size.width.max(self.min_width).min(self.max_width),
            height: size.height.max(self.min_height).min(self.max_height),
        }
    }

    /// The maximum size allowed by these constraints.
    pub fn biggest(&self) -> PtSize {
        PtSize {
            width: self.max_width,
            height: self.max_height,
        }
    }

    /// The minimum size allowed by these constraints.
    pub fn smallest(&self) -> PtSize {
        PtSize {
            width: self.min_width,
            height: self.min_height,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tight_constraints() {
        let c = BoxConstraints::tight(PtSize::new(Pt::new(100.0), Pt::new(200.0)));
        assert!(c.is_tight());
        assert!(c.is_tight_width());
        assert!(c.is_tight_height());
        assert_eq!(c.min_width.raw(), 100.0);
        assert_eq!(c.max_width.raw(), 100.0);
        assert_eq!(c.min_height.raw(), 200.0);
        assert_eq!(c.max_height.raw(), 200.0);
    }

    #[test]
    fn tight_width_loose_height() {
        let c = BoxConstraints::tight_width(Pt::new(300.0), Pt::new(500.0));
        assert!(c.is_tight_width());
        assert!(!c.is_tight_height());
        assert_eq!(c.min_height.raw(), 0.0);
        assert_eq!(c.max_height.raw(), 500.0);
    }

    #[test]
    fn loose_constraints() {
        let c = BoxConstraints::loose(PtSize::new(Pt::new(400.0), Pt::new(600.0)));
        assert!(!c.is_tight());
        assert_eq!(c.min_width.raw(), 0.0);
        assert_eq!(c.max_width.raw(), 400.0);
        assert_eq!(c.min_height.raw(), 0.0);
        assert_eq!(c.max_height.raw(), 600.0);
    }

    #[test]
    fn unbounded_constraints() {
        let c = BoxConstraints::unbounded();
        assert!(c.max_width.raw().is_infinite());
        assert!(c.max_height.raw().is_infinite());
    }

    #[test]
    fn biggest_and_smallest() {
        let c = BoxConstraints::new(
            Pt::new(10.0), Pt::new(100.0),
            Pt::new(20.0), Pt::new(200.0),
        );
        assert_eq!(c.biggest(), PtSize::new(Pt::new(100.0), Pt::new(200.0)));
        assert_eq!(c.smallest(), PtSize::new(Pt::new(10.0), Pt::new(20.0)));
    }

    #[test]
    fn constrain_clamps_to_bounds() {
        let c = BoxConstraints::new(
            Pt::new(50.0), Pt::new(200.0),
            Pt::new(50.0), Pt::new(200.0),
        );
        // Too small
        let s1 = c.constrain(PtSize::new(Pt::new(10.0), Pt::new(10.0)));
        assert_eq!(s1, PtSize::new(Pt::new(50.0), Pt::new(50.0)));

        // Too big
        let s2 = c.constrain(PtSize::new(Pt::new(999.0), Pt::new(999.0)));
        assert_eq!(s2, PtSize::new(Pt::new(200.0), Pt::new(200.0)));

        // Within bounds
        let s3 = c.constrain(PtSize::new(Pt::new(100.0), Pt::new(100.0)));
        assert_eq!(s3, PtSize::new(Pt::new(100.0), Pt::new(100.0)));
    }

    #[test]
    fn deflate_shrinks_constraints() {
        let c = BoxConstraints::tight(PtSize::new(Pt::new(400.0), Pt::new(600.0)));
        let edges = PtEdgeInsets::new(
            Pt::new(10.0),  // top
            Pt::new(20.0),  // right
            Pt::new(30.0),  // bottom
            Pt::new(40.0),  // left
        );
        let d = c.deflate(&edges);

        // width shrinks by left+right = 60
        assert_eq!(d.max_width.raw(), 340.0);
        // height shrinks by top+bottom = 40
        assert_eq!(d.max_height.raw(), 560.0);
    }

    #[test]
    fn deflate_does_not_go_negative() {
        let c = BoxConstraints::tight(PtSize::new(Pt::new(10.0), Pt::new(10.0)));
        let edges = PtEdgeInsets::new(
            Pt::new(100.0), Pt::new(100.0), Pt::new(100.0), Pt::new(100.0),
        );
        let d = c.deflate(&edges);
        assert_eq!(d.max_width.raw(), 0.0);
        assert_eq!(d.max_height.raw(), 0.0);
    }

    #[test]
    fn enforce_intersects_constraints() {
        let parent = BoxConstraints::new(
            Pt::new(0.0), Pt::new(400.0),
            Pt::new(0.0), Pt::new(600.0),
        );
        let child = BoxConstraints::new(
            Pt::new(100.0), Pt::new(300.0),
            Pt::new(50.0), Pt::new(500.0),
        );
        let result = parent.enforce(&child);

        assert_eq!(result.min_width.raw(), 100.0);
        assert_eq!(result.max_width.raw(), 300.0);
        assert_eq!(result.min_height.raw(), 50.0);
        assert_eq!(result.max_height.raw(), 500.0);
    }

    #[test]
    fn enforce_tight_parent_wins() {
        let parent = BoxConstraints::tight(PtSize::new(Pt::new(200.0), Pt::new(300.0)));
        let child = BoxConstraints::loose(PtSize::new(Pt::new(400.0), Pt::new(600.0)));
        let result = parent.enforce(&child);

        // Parent is tight — result should also be tight at parent's size
        assert!(result.is_tight());
        assert_eq!(result.max_width.raw(), 200.0);
        assert_eq!(result.max_height.raw(), 300.0);
    }

    #[test]
    fn enforce_wider_child_gets_clamped() {
        let parent = BoxConstraints::new(
            Pt::new(0.0), Pt::new(200.0),
            Pt::new(0.0), Pt::new(200.0),
        );
        let child = BoxConstraints::new(
            Pt::new(0.0), Pt::new(999.0),
            Pt::new(0.0), Pt::new(999.0),
        );
        let result = parent.enforce(&child);

        assert_eq!(result.max_width.raw(), 200.0, "child can't exceed parent");
        assert_eq!(result.max_height.raw(), 200.0);
    }

    // ── Constraint flow simulation ───────────────────────────────────────

    #[test]
    fn page_to_body_to_cell_cascade() {
        // Simulate: Page (612x792) → margins (72 each) → table cell (200 wide, 20+20 margins)
        let page = BoxConstraints::tight(PtSize::new(Pt::new(612.0), Pt::new(792.0)));
        let margins = PtEdgeInsets::new(
            Pt::new(72.0), Pt::new(72.0), Pt::new(72.0), Pt::new(72.0),
        );
        let body = page.deflate(&margins);
        assert_eq!(body.max_width.raw(), 468.0);  // 612 - 144
        assert_eq!(body.max_height.raw(), 648.0);  // 792 - 144

        // Table cell: tight width 200, loose height
        let cell = BoxConstraints::tight_width(Pt::new(200.0), body.max_height);
        let cell_padding = PtEdgeInsets::new(
            Pt::new(5.0), Pt::new(10.0), Pt::new(5.0), Pt::new(10.0),
        );
        let cell_content = cell.deflate(&cell_padding);
        assert_eq!(cell_content.max_width.raw(), 180.0);  // 200 - 20
        assert!(cell_content.is_tight_width());
    }
}
