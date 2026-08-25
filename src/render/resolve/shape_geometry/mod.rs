//! Shape geometry resolution — evaluates OOXML geometry definitions
//! (`prstGeom`, `custGeom`) into drawable paths in shape-local Pt
//! coordinates.
//!
//! # Pipeline
//!
//! 1. Parse produces a `ShapeGeometry` (Preset | Custom).
//! 2. `build_geometry(&geom, extent)` dispatches to the preset generator or
//!    the custom-geometry evaluator, yielding a `ShapePath`.
//! 3. The layout phase places the `ShapePath` within the document; the
//!    painter converts `PathVerb`s to Skia path operations.
//!
//! All coordinates in a `ShapePath` are in shape-local points (origin at the
//! top-left of the shape, positive x right, positive y down). Angles follow
//! the OOXML convention — 60000ths of a degree — and are kept in that unit
//! so the painter can apply them directly to Skia's clockwise-from-3-o'clock
//! arc operations without a unit conversion at the call site.

pub mod custom;
pub mod guides;
pub mod presets;

use crate::model::dimension::{Dimension, SixtieThousandthDeg};
use crate::model::ShapeGeometry;
use crate::render::dimension::Pt;
use crate::render::geometry::{PtOffset, PtRect, PtSize};

/// A fully evaluated shape geometry, ready for placement and painting.
#[derive(Clone, Debug)]
pub struct ShapePath {
    /// One or more subpaths. A preset typically produces one; `custGeom`
    /// preserves the source `<a:path>` grouping.
    pub paths: Vec<SubPath>,
    /// Optional text rectangle for `<wps:txbx>` body layout. Presets may
    /// leave this `None`; `custGeom` copies the `<a:rect>` child if present.
    pub text_rect: Option<PtRect>,
}

/// A single contiguous path in shape-local Pt.
#[derive(Clone, Debug)]
pub struct SubPath {
    pub verbs: Vec<PathVerb>,
    /// §20.1.10.45 path-level fill mode.
    pub fill_mode: crate::model::PathFillMode,
    /// Whether this path should be stroked with the shape's outline.
    pub stroked: bool,
}

/// Path verbs in shape-local Pt. Maps directly onto Skia path operations
/// with an angular-unit conversion for [`PathVerb::ArcTo`].
#[derive(Clone, Debug)]
pub enum PathVerb {
    MoveTo(PtOffset),
    LineTo(PtOffset),
    QuadTo(PtOffset, PtOffset),
    CubicTo(PtOffset, PtOffset, PtOffset),
    /// Elliptical arc. `radii` are horizontal/vertical radii; angles follow
    /// OOXML convention (60000ths of a degree; 0° points to the right,
    /// positive swing is clockwise). §20.1.9.3: they are *ray* angles
    /// measured at the ellipse center — on a non-circular ellipse the
    /// painter skews them onto Skia's parametric convention. The start
    /// point of the arc is implicit (the prior path cursor, which lies on
    /// the ellipse at `start_angle`); the arc does not implicitly line-to
    /// its start.
    ArcTo {
        radii: PtSize,
        start_angle: Dimension<SixtieThousandthDeg>,
        swing_angle: Dimension<SixtieThousandthDeg>,
    },
    Close,
}

/// Build a `ShapePath` for a parsed `ShapeGeometry` given the shape's
/// rendered extent.
///
/// Returns `None` if:
///  * a preset has no generator registered (Tier 0 supports `line` and
///    `rect` only; the call site logs once and skips the shape),
///  * the extent is zero in both dimensions (nothing to draw).
pub fn build_geometry(geometry: &ShapeGeometry, extent: PtSize) -> Option<ShapePath> {
    // Reject only fully zero-extent shapes. Lines are commonly authored as
    // `cx=0, cy=N` (vertical) or `cx=N, cy=0` (horizontal); both are valid
    // and must render.
    if extent.width <= Pt::ZERO && extent.height <= Pt::ZERO {
        return None;
    }
    match geometry {
        ShapeGeometry::Preset(def) => presets::build_preset(def, extent),
        ShapeGeometry::Custom(def) => custom::build_custom(def, extent),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{PresetGeometryDef, PresetShapeType, ShapeGeometry};

    fn preset(preset: PresetShapeType) -> ShapeGeometry {
        ShapeGeometry::Preset(PresetGeometryDef {
            preset,
            adjust_values: vec![],
        })
    }

    #[test]
    fn zero_extent_returns_none() {
        // Both dimensions zero → nothing to draw.
        let g = preset(PresetShapeType::Rect);
        assert!(build_geometry(&g, PtSize::new(Pt::ZERO, Pt::ZERO)).is_none());
    }

    #[test]
    fn single_zero_dimension_still_builds() {
        // A vertical line is authored as cx=0, cy=N — must not be rejected.
        let g = preset(PresetShapeType::Line);
        assert!(build_geometry(&g, PtSize::new(Pt::ZERO, Pt::new(40.0))).is_some());
        // …and a horizontal line as cx=N, cy=0.
        assert!(build_geometry(&g, PtSize::new(Pt::new(40.0), Pt::ZERO)).is_some());
    }

    #[test]
    fn preset_dispatch_builds_supported_shapes() {
        let extent = PtSize::new(Pt::new(30.0), Pt::new(20.0));
        assert!(build_geometry(&preset(PresetShapeType::Rect), extent).is_some());
        assert!(build_geometry(&preset(PresetShapeType::Line), extent).is_some());
    }

    #[test]
    fn unimplemented_preset_returns_none() {
        let extent = PtSize::new(Pt::new(30.0), Pt::new(20.0));
        assert!(build_geometry(&preset(PresetShapeType::Star12), extent).is_none());
    }
}
