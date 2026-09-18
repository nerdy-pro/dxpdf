//! Preset shape generators (§20.1.9.18 ST_ShapeType).
//!
//! Each generator is a pure function `PtSize → ShapePath`. Dispatch by
//! variant lives in [`build_preset`]. Unimplemented presets return `None`
//! and log once; callers should fall back to the shape's bounding box or
//! skip the shape.
//!
//! Tier 0 was `line` and `rect` — the minimum to validate the pipeline end
//! to end. Issue #155 added the shapes SmartArt's own layouts emit —
//! `roundRect` (the default node), `ellipse` (cycles), the four block
//! arrows, `chevron` and `homePlate` (processes), `diamond` and `triangle` —
//! which are also the most common hand-drawn `wps` shapes. The remaining
//! ~180 still return `None` and log once.

mod arrows;
mod common;
mod ellipse;
mod line;
mod polygons;
mod rect;
mod round_rect;

use crate::model::{PresetGeometryDef, PresetShapeType};
use crate::render::geometry::PtSize;

use super::ShapePath;

/// Dispatch a preset to its generator. Returns `None` for presets not yet
/// implemented; the call site is expected to log.
pub fn build_preset(def: &PresetGeometryDef, extent: PtSize) -> Option<ShapePath> {
    use arrows::Direction;
    match def.preset {
        PresetShapeType::Line => Some(line::build(extent)),
        PresetShapeType::Rect => Some(rect::build(extent)),
        PresetShapeType::RoundRect => Some(round_rect::build(def, extent)),
        PresetShapeType::Ellipse => Some(ellipse::build(def, extent)),
        PresetShapeType::RightArrow => Some(arrows::build(def, extent, Direction::Right)),
        PresetShapeType::LeftArrow => Some(arrows::build(def, extent, Direction::Left)),
        PresetShapeType::UpArrow => Some(arrows::build(def, extent, Direction::Up)),
        PresetShapeType::DownArrow => Some(arrows::build(def, extent, Direction::Down)),
        PresetShapeType::Chevron => Some(polygons::chevron(def, extent)),
        PresetShapeType::HomePlate => Some(polygons::home_plate(def, extent)),
        PresetShapeType::Diamond => Some(polygons::diamond(def, extent)),
        PresetShapeType::Triangle => Some(polygons::triangle(def, extent)),
        _ => {
            log::warn!(
                "shape_geometry: preset {:?} not yet implemented",
                def.preset
            );
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::dimension::Pt;

    fn def(preset: PresetShapeType) -> PresetGeometryDef {
        PresetGeometryDef {
            preset,
            adjust_values: vec![],
        }
    }

    #[test]
    fn line_dispatches() {
        let p = build_preset(
            &def(PresetShapeType::Line),
            PtSize::new(Pt::new(10.0), Pt::new(20.0)),
        );
        assert!(p.is_some());
    }

    #[test]
    fn rect_dispatches() {
        let p = build_preset(
            &def(PresetShapeType::Rect),
            PtSize::new(Pt::new(10.0), Pt::new(20.0)),
        );
        assert!(p.is_some());
    }

    #[test]
    fn unknown_preset_returns_none() {
        let p = build_preset(
            &def(PresetShapeType::Star12),
            PtSize::new(Pt::new(10.0), Pt::new(20.0)),
        );
        assert!(p.is_none());
    }
}
