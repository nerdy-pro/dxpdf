//! Floating image layout — positioned outside the normal flow.

use crate::dimension::Pt;
use crate::geometry::PtRect;

/// A floating image that affects text layout on the current page.
#[derive(Debug, Clone)]
pub struct ActiveFloat {
    /// Absolute x position on page.
    pub page_x: Pt,
    /// Top of the float on the page.
    pub page_y_start: Pt,
    /// Bottom of the float on the page.
    pub page_y_end: Pt,
    /// Width of the float.
    pub width: Pt,
}

impl ActiveFloat {
    /// Whether a given y-position overlaps this float's vertical range.
    pub fn overlaps_y(&self, y: Pt) -> bool {
        y >= self.page_y_start && y < self.page_y_end
    }

    /// The rectangle occupied by this float.
    pub fn rect(&self) -> PtRect {
        PtRect::from_xywh(
            self.page_x,
            self.page_y_start,
            self.width,
            self.page_y_end - self.page_y_start,
        )
    }
}

/// Compute how much the available width should be reduced on a given line
/// due to active floating images.
///
/// Returns (indent_left, indent_right) — additional indentation to avoid floats.
/// `line_y` is the top of the line, `line_height` is the line's height.
/// A line overlaps a float if any part of the line's vertical range intersects
/// the float's vertical range.
pub fn float_adjustments(
    floats: &[ActiveFloat],
    line_y: Pt,
    page_x: Pt,
    content_width: Pt,
) -> (Pt, Pt) {
    float_adjustments_with_height(floats, line_y, Pt::ZERO, page_x, content_width)
}

/// Like `float_adjustments` but with explicit line height for overlap checking.
pub fn float_adjustments_with_height(
    floats: &[ActiveFloat],
    line_y: Pt,
    line_height: Pt,
    page_x: Pt,
    content_width: Pt,
) -> (Pt, Pt) {
    let mut indent_left = Pt::ZERO;
    let mut indent_right = Pt::ZERO;
    let line_bottom = line_y + line_height;

    for float in floats {
        // Check if any part of the line overlaps the float vertically.
        if line_bottom <= float.page_y_start || line_y >= float.page_y_end {
            continue;
        }

        let float_right_edge = float.page_x + float.width;
        let content_right = page_x + content_width;
        let content_center = page_x + content_width * 0.5;
        let float_center = float.page_x + float.width * 0.5;

        // If float is on the left side of content, push text right.
        // If on the right side, push text left.
        if float_center < content_center {
            let shift = float_right_edge - page_x;
            if shift > indent_left {
                indent_left = shift;
            }
        } else {
            let shift = content_right - float.page_x;
            if shift > indent_right {
                indent_right = shift;
            }
        }
    }

    (indent_left, indent_right)
}

/// Remove floats that the cursor has passed below.
pub fn prune_floats(floats: &mut Vec<ActiveFloat>, cursor_y: Pt) {
    floats.retain(|f| cursor_y < f.page_y_end);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn no_floats_no_adjustment() {
        let (l, r) = float_adjustments(&[], Pt::new(100.0), Pt::new(72.0), Pt::new(468.0));
        assert_eq!(l.raw(), 0.0);
        assert_eq!(r.raw(), 0.0);
    }

    #[test]
    fn float_on_left_pushes_text_right() {
        let floats = vec![ActiveFloat {
            page_x: Pt::new(72.0), // at left margin
            page_y_start: Pt::new(80.0),
            page_y_end: Pt::new(200.0),
            width: Pt::new(100.0),
        }];
        let (l, r) = float_adjustments(&floats, Pt::new(100.0), Pt::new(72.0), Pt::new(468.0));
        assert_eq!(l.raw(), 100.0, "push right by float width");
        assert_eq!(r.raw(), 0.0);
    }

    #[test]
    fn float_on_right_pushes_text_left() {
        let floats = vec![ActiveFloat {
            page_x: Pt::new(440.0), // near right margin
            page_y_start: Pt::new(80.0),
            page_y_end: Pt::new(200.0),
            width: Pt::new(100.0),
        }];
        let (l, r) = float_adjustments(&floats, Pt::new(100.0), Pt::new(72.0), Pt::new(468.0));
        assert_eq!(l.raw(), 0.0);
        assert!(r.raw() > 0.0, "should indent from right");
    }

    #[test]
    fn float_not_overlapping_line_no_adjustment() {
        let floats = vec![ActiveFloat {
            page_x: Pt::new(72.0),
            page_y_start: Pt::new(200.0),
            page_y_end: Pt::new(300.0),
            width: Pt::new(100.0),
        }];
        let (l, r) = float_adjustments(&floats, Pt::new(100.0), Pt::new(72.0), Pt::new(468.0));
        assert_eq!(l.raw(), 0.0, "line is above float");
        assert_eq!(r.raw(), 0.0);
    }

    #[test]
    fn prune_removes_passed_floats() {
        let mut floats = vec![
            ActiveFloat {
                page_x: Pt::ZERO,
                page_y_start: Pt::new(0.0),
                page_y_end: Pt::new(100.0),
                width: Pt::new(50.0),
            },
            ActiveFloat {
                page_x: Pt::ZERO,
                page_y_start: Pt::new(0.0),
                page_y_end: Pt::new(300.0),
                width: Pt::new(50.0),
            },
        ];
        prune_floats(&mut floats, Pt::new(150.0));
        assert_eq!(floats.len(), 1, "first float pruned, second still active");
    }

    #[test]
    fn overlaps_y_boundary() {
        let f = ActiveFloat {
            page_x: Pt::ZERO,
            page_y_start: Pt::new(100.0),
            page_y_end: Pt::new(200.0),
            width: Pt::new(50.0),
        };
        assert!(!f.overlaps_y(Pt::new(99.0)));
        assert!(f.overlaps_y(Pt::new(100.0)));
        assert!(f.overlaps_y(Pt::new(150.0)));
        assert!(!f.overlaps_y(Pt::new(200.0)));
    }
}
