//! Draw commands — the output of layout, consumed by the painter.

use std::rc::Rc;

use crate::model::dimension::{Dimension, SixtieThousandthDeg};
use crate::render::dimension::Pt;
use crate::render::emoji::cluster::{EmojiPresentation, EmojiStructure};
use crate::render::fonts::TypefaceEntry;
use crate::render::geometry::{PtLineSegment, PtOffset, PtRect, PtSize};
use crate::render::resolve::color::RgbColor;
use crate::render::resolve::drawing_color::Rgba;
use crate::render::resolve::images::MediaEntry;
use crate::render::resolve::shape_geometry::SubPath;

/// A positioned drawing command — absolute page coordinates.
#[derive(Debug, Clone)]
pub enum DrawCommand {
    Text {
        position: PtOffset,
        text: Rc<str>,
        font_family: Rc<str>,
        char_spacing: Pt,
        font_size: Pt,
        bold: bool,
        italic: bool,
        color: RgbColor,
        /// §17.3.2.45: horizontal scale factor (1.0 = normal, 0.8 = 80%,
        /// 1.5 = 150%). Painter applies via `Font::set_scale_x`.
        text_scale: f32,
    },
    Underline {
        line: PtLineSegment,
        color: RgbColor,
        width: Pt,
    },
    Line {
        line: PtLineSegment,
        color: RgbColor,
        width: Pt,
    },
    Image {
        rect: PtRect,
        image_data: MediaEntry,
    },
    /// One emoji grapheme cluster placed at `rect`. The painter rasterizes
    /// the cluster against `typeface` (Skia raster backend honours the color
    /// glyph tables that the PDF backend strips) and embeds the result as
    /// an inline image. See `docs/emoji-rendering.md`.
    EmojiCluster {
        /// Page-coordinate rectangle at which to draw the rasterized image.
        rect: PtRect,
        /// Cluster text — one grapheme cluster, possibly multi-codepoint.
        text: String,
        /// Color emoji typeface to rasterize against.
        typeface: TypefaceEntry,
        /// Font size in Pt at which to rasterize.
        size: Pt,
        /// UTS #51 cluster classification (carried for future paint-side
        /// behaviour and diagnostics; not needed by the raster pipeline).
        presentation: EmojiPresentation,
        structure: EmojiStructure,
    },
    Rect {
        rect: PtRect,
        color: RgbColor,
    },
    LinkAnnotation {
        rect: PtRect,
        url: String,
    },
    /// Internal link to a named destination (bookmark).
    InternalLink {
        rect: PtRect,
        destination: String,
    },
    /// Named destination marker (bookmark target).
    NamedDestination {
        position: PtOffset,
        name: String,
    },
    /// §20.1.8 shape draw command — a resolved geometry with fill, stroke,
    /// and optional effects. `paths` are in shape-local Pt; the painter
    /// applies the placement transform (origin/rotation/flip).
    Path {
        /// Top-left placement anchor in page coordinates.
        origin: PtOffset,
        /// Rotation in 60000ths of a degree, around the shape's center.
        rotation: Dimension<SixtieThousandthDeg>,
        /// Mirror the shape horizontally before placement.
        flip_h: bool,
        /// Mirror the shape vertically before placement.
        flip_v: bool,
        /// Shape-local size — the bounding box the geometry was built for.
        extent: PtSize,
        /// One or more subpaths; each may be stroked and/or filled.
        paths: Vec<SubPath>,
        /// Fill applied to path interiors.
        fill: ResolvedFill,
        /// Stroke applied to path outlines (if stroked at the path level).
        stroke: Option<ResolvedStroke>,
        /// Post-processing effects (shadow, glow, …). Painter applies in order.
        effects: Vec<ResolvedEffect>,
    },
}

// ── Resolved render-ready types ─────────────────────────────────────────────

/// Painter-ready fill specification. Tier 0 renders `None` and `Solid`
/// faithfully; the other variants fall through with a one-time log.
#[derive(Clone, Debug)]
pub enum ResolvedFill {
    /// No fill — path interior is transparent.
    None,
    /// §20.1.8.54 solidFill — a single RGBA color.
    Solid(Rgba),
    /// §20.1.8.33 gradFill — stops and a direction. Tier 2 renders.
    Gradient(ResolvedGradient),
    /// §20.1.8.14 blipFill — image fill. Tier 2 renders.
    Blip(ResolvedBlip),
    /// §20.1.8.47 pattFill — preset pattern with fg/bg. Tier 3 renders.
    Pattern(ResolvedPattern),
}

/// Painter-ready stroke specification. Color/width are always honoured;
/// dash pattern is applied as a Skia path effect; cap/join map to Skia's
/// stroke primitives.
#[derive(Clone, Debug)]
pub struct ResolvedStroke {
    pub width: Pt,
    pub color: Rgba,
    pub dash: ResolvedDashPattern,
    pub cap: ResolvedLineCap,
    pub join: ResolvedLineJoin,
}

/// §20.1.10.31 ST_LineCap, ready for Skia.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ResolvedLineCap {
    Butt,
    Round,
    Square,
}

/// §20.1.2.2.22 EG_LineJoinProperties, ready for Skia.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ResolvedLineJoin {
    Round,
    Bevel,
    Miter,
}

/// Stroke dash pattern. `Solid` renders without a dash effect; `Dashes`
/// alternates on/off lengths in Pt.
#[derive(Clone, Debug, PartialEq)]
pub enum ResolvedDashPattern {
    Solid,
    Dashes(Vec<Pt>),
}

/// Post-processing effects. Tier 0 carries the data only — the painter
/// renders none of these until later tiers.
#[derive(Clone, Debug)]
pub enum ResolvedEffect {
    /// §20.1.8.45 outerShdw.
    OuterShadow {
        blur_radius: Pt,
        offset: PtOffset,
        color: Rgba,
    },
}

/// Painter-ready gradient fill — included for typing completeness; Tier 0
/// painter logs and draws no fill. Tier 2 adds rendering.
#[derive(Clone, Debug)]
pub struct ResolvedGradient {
    pub stops: Vec<GradientStopRgba>,
    pub kind: ResolvedGradientKind,
}

#[derive(Clone, Copy, Debug)]
pub struct GradientStopRgba {
    /// Position in `[0, 1]`.
    pub position: f32,
    pub color: Rgba,
}

#[derive(Clone, Copy, Debug)]
pub enum ResolvedGradientKind {
    /// Linear gradient at angle (degrees, OOXML convention: 0° = horizontal,
    /// clockwise positive).
    Linear { angle_deg: f32 },
    /// Radial/path-based gradient.
    Radial,
}

/// Painter-ready image fill — pointer to decoded bytes + source crop.
#[derive(Clone, Debug)]
pub struct ResolvedBlip {
    pub data: Rc<[u8]>,
    pub format: crate::model::ImageFormat,
    /// Fraction of the source to crop: values in `[0, 1]` relative to the
    /// blip's natural extent.
    pub src_rect: Option<PtRect>,
}

/// Painter-ready pattern fill — foreground/background + preset id.
#[derive(Clone, Debug)]
pub struct ResolvedPattern {
    pub preset: crate::model::PresetPatternVal,
    pub fg: Rgba,
    pub bg: Rgba,
}

impl DrawCommand {
    /// Shift all coordinates by `(dx, dy)`.
    pub fn shift(&mut self, dx: Pt, dy: Pt) {
        match self {
            DrawCommand::Text { position, .. } => {
                position.x += dx;
                position.y += dy;
            }
            DrawCommand::Underline { line, .. } | DrawCommand::Line { line, .. } => {
                line.start.x += dx;
                line.start.y += dy;
                line.end.x += dx;
                line.end.y += dy;
            }
            DrawCommand::Image { rect, .. }
            | DrawCommand::EmojiCluster { rect, .. }
            | DrawCommand::Rect { rect, .. }
            | DrawCommand::LinkAnnotation { rect, .. }
            | DrawCommand::InternalLink { rect, .. } => {
                rect.origin.x += dx;
                rect.origin.y += dy;
            }
            DrawCommand::NamedDestination { position, .. } => {
                position.x += dx;
                position.y += dy;
            }
            DrawCommand::Path { origin, .. } => {
                origin.x += dx;
                origin.y += dy;
            }
        }
    }

    /// Shift all y-coordinates by `dy`.
    pub fn shift_y(&mut self, dy: Pt) {
        self.shift(Pt::ZERO, dy);
    }

    /// Shift all x-coordinates by `dx`.
    pub fn shift_x(&mut self, dx: Pt) {
        self.shift(dx, Pt::ZERO);
    }
}

/// A fully laid-out page — ready for painting.
#[derive(Debug, Clone)]
pub struct LayoutedPage {
    pub commands: Vec<DrawCommand>,
    pub page_size: PtSize,
}

impl LayoutedPage {
    pub fn new(page_size: PtSize) -> Self {
        Self {
            commands: Vec::new(),
            page_size,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shift_y_moves_text() {
        let mut cmd = DrawCommand::Text {
            position: PtOffset::new(Pt::new(10.0), Pt::new(20.0)),
            text: "hi".into(),
            font_family: Rc::from("Arial"),
            char_spacing: Pt::ZERO,
            font_size: Pt::new(12.0),
            bold: false,
            italic: false,
            color: RgbColor::BLACK,
            text_scale: 1.0,
        };
        cmd.shift_y(Pt::new(5.0));
        if let DrawCommand::Text { position, .. } = cmd {
            assert_eq!(position.y.raw(), 25.0);
            assert_eq!(position.x.raw(), 10.0); // x unchanged
        }
    }

    #[test]
    fn shift_y_moves_line() {
        let mut cmd = DrawCommand::Line {
            line: PtLineSegment::new(
                PtOffset::new(Pt::new(0.0), Pt::new(10.0)),
                PtOffset::new(Pt::new(100.0), Pt::new(10.0)),
            ),
            color: RgbColor::BLACK,
            width: Pt::new(1.0),
        };
        cmd.shift_y(Pt::new(50.0));
        if let DrawCommand::Line { line, .. } = cmd {
            assert_eq!(line.start.y.raw(), 60.0, "Line y shifted");
        }
    }

    #[test]
    fn shift_y_moves_rect() {
        let mut cmd = DrawCommand::Rect {
            rect: PtRect::from_xywh(Pt::new(0.0), Pt::new(10.0), Pt::new(50.0), Pt::new(20.0)),
            color: RgbColor::BLACK,
        };
        cmd.shift_y(Pt::new(100.0));
        if let DrawCommand::Rect { rect, .. } = cmd {
            assert_eq!(rect.origin.y.raw(), 110.0);
        }
    }

    #[test]
    fn layouted_page_new() {
        let page = LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)));
        assert!(page.commands.is_empty());
        assert_eq!(page.page_size.width.raw(), 612.0);
    }

    #[test]
    fn shift_y_moves_path_origin() {
        use crate::model::dimension::Dimension;

        let mut cmd = DrawCommand::Path {
            origin: PtOffset::new(Pt::new(10.0), Pt::new(20.0)),
            rotation: Dimension::new(0),
            flip_h: false,
            flip_v: false,
            extent: PtSize::new(Pt::new(100.0), Pt::new(50.0)),
            paths: vec![],
            fill: ResolvedFill::None,
            stroke: None,
            effects: vec![],
        };
        cmd.shift_y(Pt::new(7.5));
        if let DrawCommand::Path { origin, .. } = cmd {
            assert_eq!(origin.y.raw(), 27.5);
            assert_eq!(origin.x.raw(), 10.0);
        } else {
            panic!();
        }
    }
}
