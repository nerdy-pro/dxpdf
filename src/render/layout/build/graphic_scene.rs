//! SmartArt scenes (issue #155): a parsed `dsp:` drawing → draw commands.
//!
//! The drawing part is already laid out — every shape carries an absolute
//! `a:xfrm` in EMU on a canvas whose size is the hosting drawing's
//! `wp:extent` — so "rendering" is a per-shape walk through the pipeline
//! ordinary `wps` shapes already use: `build_geometry` for the outline,
//! `resolve_shape_visuals` for fill/stroke/effects, and the shape-text
//! sub-layout for the label, placed in the shape's own `dsp:txXfrm`
//! rectangle when it has one. Commands come out in drawing-local Pt
//! coordinates; the caller shifts them to the page (the same contract the
//! floating-shape text commands follow).
//!
//! One safety net: a non-Word producer may write shapes whose canvas
//! disagrees with `wp:extent` (Word rewrites the part on every resize, so
//! its coordinates always agree). When the shape union overflows the
//! extent, the whole scene is scaled uniformly to fit rather than drawn
//! cropped.

use crate::model::{
    self, DiagramDrawing, DiagramShape, DrawingRunProps, DrawingTextBody, DrawingTextRun,
};
use crate::render::dimension::Pt;
use crate::render::geometry::{PtRect, PtSize};
use crate::render::layout::draw_command::DrawCommand;
use crate::render::resolve::drawing_color::{resolve_drawing_color, DrawingColorContext};
use crate::render::resolve::shape_geometry::build_geometry;
use crate::render::resolve::shape_visuals::resolve_shape_visuals;

use super::floating::build_shape_text_commands;
use super::{BuildContext, BuildState};

/// Build a diagram's command scene in drawing-local Pt, `(0,0)`–`extent`.
pub(super) fn build_diagram_scene(
    drawing: &DiagramDrawing,
    extent: PtSize,
    ctx: &BuildContext,
    state: &BuildState,
) -> Vec<DrawCommand> {
    let scale = overflow_scale(drawing, extent);
    let mut commands = Vec::new();

    for shape in &drawing.shapes {
        let Some(props) = shape.shape_properties.as_ref() else {
            continue;
        };
        let Some(transform) = props.transform else {
            continue;
        };
        let (Some(off), Some(ext)) = (transform.offset, transform.extent) else {
            continue;
        };
        let shape_origin = crate::render::geometry::PtOffset::new(
            Pt::from(off.x) * scale,
            Pt::from(off.y) * scale,
        );
        let shape_extent = PtSize::new(Pt::from(ext.width) * scale, Pt::from(ext.height) * scale);

        let mut text_rect_from_geometry = None;
        if let Some(ref geometry) = props.geometry {
            if let Some(shape_path) = build_geometry(geometry, shape_extent) {
                let visuals = resolve_shape_visuals(
                    Some(props),
                    shape.style_line_ref.as_ref(),
                    shape.style_effect_ref.as_ref(),
                    shape.style_fill_ref.as_ref(),
                    ctx.resolved.theme.as_ref(),
                );
                text_rect_from_geometry = shape_path.text_rect;
                commands.push(DrawCommand::Path {
                    origin: shape_origin,
                    rotation: transform.rotation.unwrap_or_default(),
                    flip_h: transform.flip_h.unwrap_or(false),
                    flip_v: transform.flip_v.unwrap_or(false),
                    extent: shape_extent,
                    paths: shape_path.paths,
                    fill: visuals.fill,
                    stroke: visuals.stroke,
                    effects: visuals.effects,
                });
            }
        }

        let Some(ref body) = shape.text_body else {
            continue;
        };
        // §[MS-ODRAWXML]: `dsp:txXfrm` is the text rectangle, in the same
        // absolute canvas EMU as the shape — SmartArt's way of keeping a
        // label upright inside a rotated shape. Without it, the geometry's
        // own text box (offset to the shape's position) serves.
        let text_rect = match shape.text_transform.as_ref() {
            Some(t) => match (t.offset, t.extent) {
                (Some(o), Some(e)) => PtRect::from_xywh(
                    Pt::from(o.x) * scale,
                    Pt::from(o.y) * scale,
                    Pt::from(e.width) * scale,
                    Pt::from(e.height) * scale,
                ),
                _ => continue,
            },
            None => match text_rect_from_geometry {
                Some(r) => PtRect::from_xywh(
                    shape_origin.x + r.origin.x,
                    shape_origin.y + r.origin.y,
                    r.size.width,
                    r.size.height,
                ),
                None => PtRect::from_xywh(
                    shape_origin.x,
                    shape_origin.y,
                    shape_extent.width,
                    shape_extent.height,
                ),
            },
        };

        let wsp = text_carrier(shape, body, ctx);
        let mut text_commands = build_shape_text_commands(&wsp, text_rect.size, ctx, state);
        for cmd in &mut text_commands {
            cmd.shift(text_rect.origin.x, text_rect.origin.y);
        }
        commands.extend(text_commands);
    }
    commands
}

/// The uniform scale that fits the shape union into the extent — 1.0 for
/// every Word-produced part.
fn overflow_scale(drawing: &DiagramDrawing, extent: PtSize) -> f32 {
    let mut max_x = 0f32;
    let mut max_y = 0f32;
    for shape in &drawing.shapes {
        let Some(t) = shape.shape_properties.as_ref().and_then(|p| p.transform) else {
            continue;
        };
        if let (Some(off), Some(ext)) = (t.offset, t.extent) {
            max_x = max_x.max((Pt::from(off.x) + Pt::from(ext.width)).raw());
            max_y = max_y.max((Pt::from(off.y) + Pt::from(ext.height)).raw());
        }
    }
    let sx = if max_x > extent.width.raw() && max_x > 0.0 {
        extent.width.raw() / max_x
    } else {
        1.0
    };
    let sy = if max_y > extent.height.raw() && max_y > 0.0 {
        extent.height.raw() / max_y
    } else {
        1.0
    };
    sx.min(sy)
}

/// Dress a diagram shape's text as the `WordProcessingShape` the shared
/// shape-text sub-layout consumes: the DrawingML body becomes
/// WordprocessingML blocks (colors resolved to RGB here, where the theme
/// is in hand), and the `fontRef` rides along for the default face/color.
fn text_carrier(
    shape: &DiagramShape,
    body: &DrawingTextBody,
    ctx: &BuildContext,
) -> model::WordProcessingShape {
    model::WordProcessingShape {
        cnv_pr: None,
        shape_properties: None,
        style_line_ref: None,
        style_effect_ref: None,
        style_fill_ref: None,
        style_font_ref: shape.style_font_ref.clone(),
        body_pr: body.body_pr.clone(),
        txbx_content: drawing_text_to_blocks(body, ctx),
    }
}

/// §21.1.2 DrawingML text → WordprocessingML blocks, sizes converted
/// (hundredths of a point → half-points) and scheme colors resolved.
pub(super) fn drawing_text_to_blocks(
    body: &DrawingTextBody,
    ctx: &BuildContext,
) -> Vec<model::Block> {
    let color_ctx = DrawingColorContext::new(ctx.resolved.theme.as_ref());
    body.paragraphs
        .iter()
        .map(|p| {
            let mut para_props = model::ParagraphProperties::default();
            if let Some(a) = p.alignment {
                para_props.alignment = crate::model::Dup::from(Some(a));
            }
            // The label rides the shared shape-text sub-layout, whose block
            // cascade would otherwise apply the *document's* defaults (a
            // Normal style with spacing-after would air out a two-line
            // label). A DrawingML body is self-contained: explicit zero
            // spacing shields it.
            para_props.spacing = crate::model::Dup::from(Some(model::ParagraphSpacing {
                before: Some(crate::model::dimension::Dimension::new(0)),
                after: Some(crate::model::dimension::Dimension::new(0)),
                line: None,
                before_auto_spacing: Some(false),
                after_auto_spacing: Some(false),
            }));
            let content = p
                .runs
                .iter()
                .map(|run| match run {
                    DrawingTextRun::Text { text, props } => {
                        model::Inline::TextRun(Box::new(model::TextRun {
                            style_id: None,
                            properties: run_properties(props, p.default_run.as_ref(), &color_ctx),
                            content: vec![model::RunElement::Text(text.clone())],
                            rsids: Default::default(),
                        }))
                    }
                    DrawingTextRun::Break => model::Inline::TextRun(Box::new(model::TextRun {
                        style_id: None,
                        properties: run_properties(
                            &DrawingRunProps::default(),
                            p.default_run.as_ref(),
                            &color_ctx,
                        ),
                        content: vec![model::RunElement::LineBreak(model::BreakKind::TextWrapping)],
                        rsids: Default::default(),
                    })),
                })
                .collect();
            model::Block::Paragraph(Box::new(model::Paragraph {
                style_id: None,
                properties: para_props,
                mark_run_properties: None,
                content,
                rsids: Default::default(),
            }))
        })
        .collect()
}

/// Merge a run's own §21.1.2.3.9 properties over its paragraph's `defRPr`
/// and convert into the shared run-property model.
fn run_properties(
    props: &DrawingRunProps,
    default: Option<&DrawingRunProps>,
    color_ctx: &DrawingColorContext<'_>,
) -> model::RunProperties {
    let pick = |own: &Option<_>, def: fn(&DrawingRunProps) -> Option<_>| -> Option<_> {
        own.clone().or_else(|| default.and_then(def))
    };
    let size = props.size.or_else(|| default.and_then(|d| d.size));
    let bold = props.bold.or_else(|| default.and_then(|d| d.bold));
    let italic = props.italic.or_else(|| default.and_then(|d| d.italic));
    let color = props
        .color
        .clone()
        .or_else(|| default.and_then(|d| d.color.clone()));
    let family: Option<String> = pick(&props.family, |d| d.family.clone());

    let mut out = model::RunProperties {
        // §20.1.10.68 hundredths of a point → the model's half-points.
        font_size: crate::model::Dup::from(
            size.map(|s| crate::model::dimension::Dimension::new((s.raw() + 25) / 50)),
        ),
        bold,
        italic,
        color: crate::model::Dup::from(
            color.map(|c| model::Color::Rgb(resolve_drawing_color(&c, color_ctx).to_rgb24())),
        ),
        ..Default::default()
    };
    if let Some(f) = family {
        out.fonts.ascii.explicit = Some(f.clone());
        out.fonts.high_ansi.explicit = Some(f);
    }
    out
}
