//! Flutter-inspired constraint-based layout engine.
//!
//! Core protocol: **constraints go down, sizes go up, parent sets position**.

pub mod build;
pub mod cell;
pub mod draw_command;
pub mod float;
pub mod fragment;
pub mod header_footer;
pub mod line;
pub mod measurer;
pub mod page;
pub mod page_borders;
pub mod paragraph;
pub mod section;
pub mod table;

use crate::render::dimension::Pt;
use crate::render::geometry::{PtEdgeInsets, PtSize};

/// MCE §M.1.2: the branch of an `<mc:AlternateContent>` this renderer draws.
///
/// Exactly one branch is live. A consumer takes the first `<mc:Choice>` whose
/// requirements it can meet and otherwise the `<mc:Fallback>`; it never takes
/// both, because both describe the *same* object — the Choice in DrawingML,
/// the Fallback in VML for clients that predate it.
///
/// Every walker that meets the element consults [`live_mc_branch`], so the
/// answer is a property of the element rather than of who asked. Two
/// predicates over one element is what produced the original double render,
/// and a third walker that answered neither is what produced the *missing*
/// render after it.
///
/// The `Fallback` arm deliberately does **not** split "float" from "inline".
/// One VML fallback is routinely both, and by design: `extract_vml_primitive`
/// draws a `<v:rect>`'s geometry and leaves its `text_commands` empty,
/// `extract_vml_primitive_image` skips any shape that also hosts text, and the
/// inline collector picks the text box up at the host paragraph. That division
/// is by graphic role, not by branch, and it is exactly how a bare `<w:pict>`
/// already renders — so a live Fallback goes to all of them and each takes the
/// part it owns. Naming an owner per variant would look tidier and would drop
/// the text of every VML rect that has one.
pub(crate) enum McBranch<'a> {
    /// The `<mc:Choice>` elements, in document order.
    Choices(&'a [crate::model::McChoice]),
    /// The `<mc:Fallback>`, reached only when no Choice carries anything we
    /// can draw.
    Fallback(&'a [crate::model::Inline]),
    /// Choices we cannot draw, and no Fallback to fall back to.
    Neither,
}

/// Pick the live branch of `ac`.
///
/// The test is **content-based**, not a `Requires` namespace check: a Choice
/// may declare a namespace we nominally support and still hold nothing this
/// renderer turns into geometry, and the honest question is whether we will
/// actually draw it. What counts is an *anchor* — an anchored `wps:wsp` shape
/// and an anchored picture are both `Inline::Image` with
/// `ImagePlacement::Anchor`, so one question covers shapes and pictures alike
/// and no walker has to ask a narrower version of it for itself.
///
/// Recurses through `Hyperlink`/`Field` wrappers and nested elements. §M.1.2's
/// content model for a branch is `drawing | pict`, so a nested
/// `<mc:AlternateContent>` cannot come from a document — but it can be built,
/// and resolving it innermost-first is the only reading under which the outer
/// answer stays consistent with the inner one.
pub(crate) fn live_mc_branch(ac: &crate::model::AlternateContent) -> McBranch<'_> {
    use crate::model::{ImagePlacement, Inline};

    fn draws_an_anchor(inlines: &[Inline]) -> bool {
        inlines.iter().any(|inline| match inline {
            Inline::Image(img) => matches!(img.placement, ImagePlacement::Anchor(_)),
            Inline::Hyperlink(link) => draws_an_anchor(&link.content),
            Inline::Field(f) => draws_an_anchor(&f.content),
            Inline::AlternateContent(inner) => {
                matches!(live_mc_branch(inner), McBranch::Choices(_))
            }
            _ => false,
        })
    }

    if ac.choices.iter().any(|c| draws_an_anchor(&c.content)) {
        McBranch::Choices(&ac.choices)
    } else {
        match ac.fallback {
            Some(ref fallback) => McBranch::Fallback(fallback),
            None => McBranch::Neither,
        }
    }
}

/// §20.1.2.1.18: the uniform shrink `a:normAutofit` applies to one shape's text
/// body — a multiplier on every font size in it, and one on every paragraph's
/// line spacing.
///
/// Lives here rather than beside `BuildState` because both the block builder
/// and the fragment layer apply it, and `fragment` must not depend on `build`.
///
/// [`ShapeAutoFit::NONE`] is the identity and is what every call site outside a
/// shape text body passes. The functions that resolve a font size or a line
/// spacing take it as an explicit *parameter* rather than reaching for a
/// default, so that a new call site has to say which it is — a silently
/// inherited `1.0` is how these attributes came to be dropped in the first
/// place.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ShapeAutoFit {
    font_scale: f32,
    line_spacing_scale: f32,
}

impl Default for ShapeAutoFit {
    fn default() -> Self {
        Self::NONE
    }
}

impl ShapeAutoFit {
    /// No shrink: the identity for both factors.
    pub const NONE: Self = Self {
        font_scale: 1.0,
        line_spacing_scale: 1.0,
    };

    /// Read the shrink off a shape's `a:bodyPr`.
    ///
    /// Total over [`TextAutoFit`](crate::model::TextAutoFit) on purpose. Only
    /// `normAutofit` carries a shrink: §20.1.2.1.16 `noAutofit` is the explicit
    /// "do not", and §20.1.2.1.20 `spAutoFit` resizes the *shape* to its text,
    /// which this sub-layout cannot do — it lays a body out inside an extent
    /// the host already fixed. `spAutoFit` therefore degrades to no shrink,
    /// which draws the text at its authored size rather than inventing a fit.
    pub fn from_body(auto_fit: Option<crate::model::TextAutoFit>) -> Self {
        use crate::model::TextAutoFit;

        let Some(TextAutoFit::NormalAutoFit(na)) = auto_fit else {
            return Self::NONE;
        };
        // Absent `@fontScale` is 100%; absent `@lnSpcReduction` is 0% *off*,
        // hence a factor of 1. A negative or absurd value would invert or
        // explode the layout, so both are clamped to a sane band rather than
        // trusted — the file is untrusted input.
        let font_scale = na
            .font_scale
            .map_or(1.0, |p| p.to_fraction())
            .clamp(0.01, 10.0);
        let line_spacing_scale =
            (1.0 - na.line_spacing_reduction.map_or(0.0, |p| p.to_fraction())).clamp(0.01, 10.0);
        Self {
            font_scale,
            line_spacing_scale,
        }
    }

    /// Apply `@fontScale` to a resolved font size.
    pub fn scale_font(self, size: Pt) -> Pt {
        if self.font_scale == 1.0 {
            size
        } else {
            Pt::new(size.raw() * self.font_scale)
        }
    }

    /// Apply `@lnSpcReduction` to a line height that has already been resolved
    /// against the paragraph's own §17.3.1.33 spacing rule.
    ///
    /// It must land *after* that resolution, not inside it. Folding the
    /// reduction into an `Auto` multiplier looks equivalent and is not:
    /// `resolve_line_height` floors `Auto` at the line's natural height, so a
    /// multiplier below 1 is swallowed whole. That floor exists to stop a
    /// user-authored multiplier from colliding glyph boxes — but crossing it is
    /// precisely what `lnSpcReduction` is for. Word has already laid the body
    /// out and decided the tightened text fits; the reduction is most of how it
    /// made it fit.
    pub fn scale_line_height(self, height: Pt) -> Pt {
        if self.line_spacing_scale == 1.0 {
            height
        } else {
            Pt::new(height.raw() * self.line_spacing_scale)
        }
    }
}

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
        let c = BoxConstraints::new(Pt::new(10.0), Pt::new(100.0), Pt::new(20.0), Pt::new(200.0));
        assert_eq!(c.biggest(), PtSize::new(Pt::new(100.0), Pt::new(200.0)));
        assert_eq!(c.smallest(), PtSize::new(Pt::new(10.0), Pt::new(20.0)));
    }

    #[test]
    fn constrain_clamps_to_bounds() {
        let c = BoxConstraints::new(Pt::new(50.0), Pt::new(200.0), Pt::new(50.0), Pt::new(200.0));
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
            Pt::new(10.0), // top
            Pt::new(20.0), // right
            Pt::new(30.0), // bottom
            Pt::new(40.0), // left
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
            Pt::new(100.0),
            Pt::new(100.0),
            Pt::new(100.0),
            Pt::new(100.0),
        );
        let d = c.deflate(&edges);
        assert_eq!(d.max_width.raw(), 0.0);
        assert_eq!(d.max_height.raw(), 0.0);
    }

    #[test]
    fn enforce_intersects_constraints() {
        let parent =
            BoxConstraints::new(Pt::new(0.0), Pt::new(400.0), Pt::new(0.0), Pt::new(600.0));
        let child = BoxConstraints::new(
            Pt::new(100.0),
            Pt::new(300.0),
            Pt::new(50.0),
            Pt::new(500.0),
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
        let parent =
            BoxConstraints::new(Pt::new(0.0), Pt::new(200.0), Pt::new(0.0), Pt::new(200.0));
        let child = BoxConstraints::new(Pt::new(0.0), Pt::new(999.0), Pt::new(0.0), Pt::new(999.0));
        let result = parent.enforce(&child);

        assert_eq!(result.max_width.raw(), 200.0, "child can't exceed parent");
        assert_eq!(result.max_height.raw(), 200.0);
    }

    // ── Constraint flow simulation ───────────────────────────────────────

    #[test]
    fn page_to_body_to_cell_cascade() {
        // Simulate: Page (612x792) → margins (72 each) → table cell (200 wide, 20+20 margins)
        let page = BoxConstraints::tight(PtSize::new(Pt::new(612.0), Pt::new(792.0)));
        let margins = PtEdgeInsets::new(Pt::new(72.0), Pt::new(72.0), Pt::new(72.0), Pt::new(72.0));
        let body = page.deflate(&margins);
        assert_eq!(body.max_width.raw(), 468.0); // 612 - 144
        assert_eq!(body.max_height.raw(), 648.0); // 792 - 144

        // Table cell: tight width 200, loose height
        let cell = BoxConstraints::tight_width(Pt::new(200.0), body.max_height);
        let cell_padding =
            PtEdgeInsets::new(Pt::new(5.0), Pt::new(10.0), Pt::new(5.0), Pt::new(10.0));
        let cell_content = cell.deflate(&cell_padding);
        assert_eq!(cell_content.max_width.raw(), 180.0); // 200 - 20
        assert!(cell_content.is_tight_width());
    }
}
