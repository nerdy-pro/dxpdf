//! Chart rendering (issue #155): a parsed chart part → draw commands.
//!
//! Draws the 2D families — bar/column (clustered, stacked, percent-stacked),
//! line, pie/doughnut, area, scatter — from the part's cached data, inside
//! the hosting drawing's extent, using the primitives the pipeline already
//! has: `Rect` for bars, `Line` for axes and gridlines, `Path` for
//! polylines, areas, pie slices and markers, `Text` for every label.
//!
//! Where the spec is silent the numbers follow Word's defaults as recorded
//! in its own chart-style part and Excel's observed behavior: automatic
//! value-axis bounds with a 5% margin rounded to a 1/2/5×10ᵏ unit, series
//! colored on the theme accent1–6 cycle (darkened per completed cycle),
//! 9pt labels / 14pt title in the theme minor face at Word's dark-gray,
//! 2.25pt series lines, 0.75pt light-gray gridlines, `gapWidth` 150 and
//! stacked `overlap` 100.
//!
//! Deliberately not drawn: the kinds [`PlotKind::Unsupported`] names
//! (radar, surface, bubble, stock — each with a `RUST_LOG=warn` line), and
//! axis titles (rotated text is not in the pipeline; the title parses and
//! is left off silently). 3D variants were flattened at parse. Blank cells
//! follow `c:dispBlanksAs`'s default — bars/points skipped, polylines
//! broken at the gap; the explicit `zero`/`span` modes are not read.
//! Values outside a manually narrowed axis are clamped to the plot edge
//! rather than clipped away (no clipping primitive); numeric category
//! caches print raw values, so a date axis shows serials. A multi-series
//! pie plots its first series, as Word does.

use std::rc::Rc;

use crate::model::{
    ChartAxis, ChartAxisKind, ChartGrouping, ChartMarker, ChartSeries, ChartSpace, DrawingColor,
    DrawingTextBody, DrawingTextRun, LegendPosition, PlotGroup, PlotKind, SchemeColorVal,
};
use crate::render::dimension::Pt;
use crate::render::fonts::Toggle;
use crate::render::geometry::{PtLineSegment, PtOffset, PtRect, PtSize};
use crate::render::layout::draw_command::{
    DrawCommand, ResolvedDashPattern, ResolvedFill, ResolvedLineCap, ResolvedLineJoin,
    ResolvedStroke,
};
use crate::render::layout::fragment::{FontProps, TextMetrics};
use crate::render::resolve::color::RgbColor;
use crate::render::resolve::drawing_color::{resolve_drawing_color, DrawingColorContext, Rgba};
use crate::render::resolve::shape_geometry::{PathVerb, SubPath};
use crate::render::resolve::shape_visuals::resolve_shape_visuals;

use super::{BuildContext, BuildState};

/// Word's label/title gray (`tx1` at 65% luminance, per its style part).
const LABEL_COLOR: RgbColor = RgbColor {
    r: 0x59,
    g: 0x59,
    b: 0x59,
};
/// Gridline/axis gray (`tx1` lumMod 15% lumOff 85%).
const CHROME_COLOR: RgbColor = RgbColor {
    r: 0xD9,
    g: 0xD9,
    b: 0xD9,
};
const LABEL_SIZE: f32 = 9.0;
const TITLE_SIZE: f32 = 14.0;
/// §21.2: Word's default series line width, 28575 EMU.
const SERIES_LINE_WIDTH: f32 = 2.25;

/// Build a chart's command scene in drawing-local Pt, `(0,0)`–`extent`.
pub(super) fn build_chart_scene(
    chart: &ChartSpace,
    extent: PtSize,
    ctx: &BuildContext,
    _state: &BuildState,
) -> Vec<DrawCommand> {
    let painter = ChartPainter::new(chart, extent, ctx);
    painter.paint()
}

struct ChartPainter<'a> {
    chart: &'a ChartSpace,
    extent: PtSize,
    ctx: &'a BuildContext<'a>,
    family: Rc<str>,
    color_ctx: DrawingColorContext<'a>,
    commands: Vec<DrawCommand>,
}

impl<'a> ChartPainter<'a> {
    fn new(chart: &'a ChartSpace, extent: PtSize, ctx: &'a BuildContext<'a>) -> Self {
        let family: Rc<str> = ctx
            .resolved
            .theme
            .as_ref()
            .map(|t| t.minor_font.latin.as_str())
            .filter(|s| !s.is_empty())
            .unwrap_or("Calibri")
            .into();
        Self {
            chart,
            extent,
            ctx,
            family,
            color_ctx: DrawingColorContext::new(ctx.resolved.theme.as_ref()),
            commands: Vec::new(),
        }
    }

    fn paint(mut self) -> Vec<DrawCommand> {
        let (w, h) = (self.extent.width.raw(), self.extent.height.raw());
        // Chart frame: Word's default thin light border.
        self.stroke_rect(PtRect::from_xywh(
            Pt::ZERO,
            Pt::ZERO,
            self.extent.width,
            self.extent.height,
        ));

        // Frame padding ≈ 2.5% of the smaller side.
        let pad = (w.min(h) * 0.025).max(2.0);
        let mut top = pad;
        let bottom = h - pad;
        let (mut left, mut right) = (pad, w - pad);

        if let Some(title) = self.title_text() {
            top += self.draw_title(&title, top, w);
        }

        let groups: Vec<&PlotGroup> = self
            .chart
            .plot_groups
            .iter()
            .filter(|g| match &g.kind {
                PlotKind::Unsupported(name) => {
                    log::warn!("[chart] §21.2 {name} is not rendered (issue #155 draws the 2D bar/line/pie/area/scatter families)");
                    false
                }
                _ => true,
            })
            .collect();

        if let Some(legend) = &self.chart.legend {
            // `c:overlay`: the legend floats over the plot — still drawn,
            // but reserving no space.
            {
                match legend.position {
                    LegendPosition::Right | LegendPosition::TopRight => {
                        let used = self.draw_legend_column(&groups, right, top, bottom);
                        if !legend.overlay {
                            right -= used;
                        }
                    }
                    LegendPosition::Left => {
                        let used = self.draw_legend_column_at(&groups, left, top, bottom);
                        if !legend.overlay {
                            left += used;
                        }
                    }
                    LegendPosition::Bottom => {
                        self.draw_legend_row(&groups, left, right, bottom);
                    }
                    LegendPosition::Top => {
                        let used = self.draw_legend_row_at(&groups, left, right, top);
                        if !legend.overlay {
                            top += used;
                        }
                    }
                }
            }
        }
        let bottom = match self.chart.legend.as_ref() {
            Some(l) if !l.overlay && l.position == LegendPosition::Bottom => {
                bottom - (LABEL_SIZE * 1.6)
            }
            _ => bottom,
        };

        if groups.is_empty() {
            return self.commands;
        }

        // Pie-family groups own the whole plot box; everything else shares
        // one axis frame.
        if matches!(groups[0].kind, PlotKind::Pie | PlotKind::Doughnut) {
            self.draw_pie(groups[0], left, top, right, bottom);
        } else {
            self.draw_axis_chart(&groups, left, top, right, bottom);
        }
        self.commands
    }

    // ── shared text plumbing ─────────────────────────────────────────────

    fn font(&self, size: f32, bold: bool) -> FontProps {
        FontProps {
            family: self.family.clone(),
            size: Pt::new(size),
            bold: if bold { Toggle::On } else { Toggle::Absent },
            italic: Toggle::Absent,
            underline: false,
            rtl: Toggle::Absent,
            char_spacing: Pt::ZERO,
            text_scale: 1.0,
            underline_position: Pt::ZERO,
            underline_thickness: Pt::ZERO,
        }
    }

    fn measure(&self, text: &str, size: f32, bold: bool) -> (f32, TextMetrics) {
        let (w, m) = self.ctx.measurer.measure(text, &self.font(size, bold));
        (w.raw(), m)
    }

    /// Text with its top-left at `(x, top)`.
    fn text_at(&mut self, x: f32, top: f32, text: &str, size: f32, bold: bool, color: RgbColor) {
        let (_, m) = self.measure(text, size, bold);
        self.commands.push(DrawCommand::Text {
            position: PtOffset::new(Pt::new(x), Pt::new(top + m.ascent.raw())),
            text: text.into(),
            font_family: self.family.clone(),
            char_spacing: Pt::ZERO,
            font_size: Pt::new(size),
            bold: if bold { Toggle::On } else { Toggle::Absent },
            italic: Toggle::Absent,
            color,
            text_scale: 1.0,
            shaped: None,
        });
    }

    fn title_text(&self) -> Option<(String, f32, bool)> {
        if self.chart.auto_title_deleted {
            return None;
        }
        match &self.chart.title {
            Some(body) => {
                let text = body_text(body);
                if text.is_empty() {
                    return None;
                }
                let (size, bold) = first_run_style(body).unwrap_or((TITLE_SIZE, false));
                Some((text, size, bold))
            }
            // §21.2.2.6: no literal title + autoTitleDeleted false titles a
            // single-series chart with the series name.
            None => {
                let only: Vec<&ChartSeries> = self
                    .chart
                    .plot_groups
                    .iter()
                    .flat_map(|g| &g.series)
                    .collect();
                match only.as_slice() {
                    [s] => s.name.clone().map(|n| (n, TITLE_SIZE, false)),
                    _ => None,
                }
            }
        }
    }

    fn draw_title(&mut self, (text, size, bold): &(String, f32, bool), top: f32, w: f32) -> f32 {
        let (tw, m) = self.measure(text, *size, *bold);
        self.text_at((w - tw) / 2.0, top, text, *size, *bold, LABEL_COLOR);
        m.ascent.raw() + m.descent.raw() + size * 0.4
    }

    // ── colors ───────────────────────────────────────────────────────────

    /// The automatic color for format index `idx`: accent1–6, darkened /
    /// lightened per completed cycle (the colors1.xml lumMod ladder,
    /// approximated).
    fn auto_color(&self, idx: u32) -> Rgba {
        const ACCENTS: [SchemeColorVal; 6] = [
            SchemeColorVal::Accent1,
            SchemeColorVal::Accent2,
            SchemeColorVal::Accent3,
            SchemeColorVal::Accent4,
            SchemeColorVal::Accent5,
            SchemeColorVal::Accent6,
        ];
        let base = resolve_drawing_color(
            &DrawingColor::Scheme {
                name: ACCENTS[(idx as usize) % 6],
                transforms: vec![],
            },
            &self.color_ctx,
        );
        match (idx as usize / 6) % 4 {
            0 => base,
            1 => lum(base, 0.6, 0.0),
            2 => lum(base, 0.8, 0.2),
            _ => lum(base, 0.8, 0.0),
        }
    }

    /// A series' fill color: explicit `c:spPr` wins, else the cycle.
    fn series_color(&self, series: &ChartSeries) -> Rgba {
        self.explicit_fill(series.shape_properties.as_ref())
            .unwrap_or_else(|| self.auto_color(series.idx))
    }

    fn explicit_fill(&self, props: Option<&crate::model::ShapeProperties>) -> Option<Rgba> {
        let props = props?;
        let visuals = resolve_shape_visuals(
            Some(props),
            None,
            None,
            None,
            self.ctx.resolved.theme.as_ref(),
        );
        match visuals.fill {
            ResolvedFill::Solid(c) => Some(c),
            _ => None,
        }
    }

    /// A pie point's color: `c:dPt` override, else the per-point cycle.
    fn point_color(&self, series: &ChartSeries, point: usize) -> Rgba {
        series
            .point_properties
            .iter()
            .find(|(i, _)| *i as usize == point)
            .and_then(|(_, p)| self.explicit_fill(Some(p)))
            .unwrap_or_else(|| self.auto_color(point as u32))
    }

    // ── chrome ───────────────────────────────────────────────────────────

    fn stroke_rect(&mut self, rect: PtRect) {
        let verbs = vec![
            PathVerb::MoveTo(rect.origin),
            PathVerb::LineTo(PtOffset::new(
                rect.origin.x + rect.size.width,
                rect.origin.y,
            )),
            PathVerb::LineTo(PtOffset::new(
                rect.origin.x + rect.size.width,
                rect.origin.y + rect.size.height,
            )),
            PathVerb::LineTo(PtOffset::new(
                rect.origin.x,
                rect.origin.y + rect.size.height,
            )),
            PathVerb::Close,
        ];
        self.push_path(
            vec![SubPath {
                verbs,
                fill_mode: crate::model::PathFillMode::None,
                stroked: true,
            }],
            ResolvedFill::None,
            Some(chrome_stroke()),
        );
    }

    fn push_path(
        &mut self,
        paths: Vec<SubPath>,
        fill: ResolvedFill,
        stroke: Option<ResolvedStroke>,
    ) {
        self.commands.push(DrawCommand::Path {
            origin: PtOffset::new(Pt::ZERO, Pt::ZERO),
            rotation: Default::default(),
            flip_h: false,
            flip_v: false,
            extent: self.extent,
            paths,
            fill,
            stroke,
            effects: Vec::new(),
        });
    }

    fn line(&mut self, x0: f32, y0: f32, x1: f32, y1: f32, color: RgbColor, width: f32) {
        self.commands.push(DrawCommand::Line {
            line: PtLineSegment::new(
                PtOffset::new(Pt::new(x0), Pt::new(y0)),
                PtOffset::new(Pt::new(x1), Pt::new(y1)),
            ),
            color,
            width: Pt::new(width),
        });
    }

    // ── legend ───────────────────────────────────────────────────────────

    fn legend_entries(&self, groups: &[&PlotGroup]) -> Vec<(String, Rgba)> {
        let mut out = Vec::new();
        for g in groups {
            // A pie legend lists the categories, per Word — including for a
            // multi-series part, where only the first series is plotted.
            if matches!(g.kind, PlotKind::Pie | PlotKind::Doughnut) && !g.series.is_empty() {
                let s = &g.series[0];
                let vary = g.vary_colors.unwrap_or(true);
                for (i, cat) in s.categories.iter().enumerate() {
                    let name = cat.clone().unwrap_or_else(|| (i + 1).to_string());
                    let color = if vary {
                        self.point_color(s, i)
                    } else {
                        self.series_color(s)
                    };
                    out.push((name, color));
                }
                continue;
            }
            for (i, s) in g.series.iter().enumerate() {
                let name = s
                    .name
                    .clone()
                    .unwrap_or_else(|| format!("Series {}", i + 1));
                out.push((name, self.series_color(s)));
            }
        }
        out
    }

    /// Right-edge legend column; returns the width consumed.
    fn draw_legend_column(
        &mut self,
        groups: &[&PlotGroup],
        right: f32,
        top: f32,
        bottom: f32,
    ) -> f32 {
        let entries = self.legend_entries(groups);
        if entries.is_empty() {
            return 0.0;
        }
        let max_w = entries
            .iter()
            .map(|(n, _)| self.measure(n, LABEL_SIZE, false).0)
            .fold(0.0, f32::max);
        let width = max_w + 14.0;
        let row_h = LABEL_SIZE * 1.5;
        let total = row_h * entries.len() as f32;
        let mut y = (top + bottom - total) / 2.0;
        let x = right - width;
        for (name, color) in entries {
            self.swatch(x, y + (row_h - 7.0) / 2.0, color);
            self.text_at(
                x + 11.0,
                y + (row_h - LABEL_SIZE * 1.2) / 2.0,
                &name,
                LABEL_SIZE,
                false,
                LABEL_COLOR,
            );
            y += row_h;
        }
        width + 4.0
    }

    fn draw_legend_column_at(
        &mut self,
        groups: &[&PlotGroup],
        left: f32,
        top: f32,
        bottom: f32,
    ) -> f32 {
        let entries = self.legend_entries(groups);
        if entries.is_empty() {
            return 0.0;
        }
        let max_w = entries
            .iter()
            .map(|(n, _)| self.measure(n, LABEL_SIZE, false).0)
            .fold(0.0, f32::max);
        let width = max_w + 14.0;
        let row_h = LABEL_SIZE * 1.5;
        let total = row_h * entries.len() as f32;
        let mut y = (top + bottom - total) / 2.0;
        for (name, color) in entries {
            self.swatch(left, y + (row_h - 7.0) / 2.0, color);
            self.text_at(
                left + 11.0,
                y + (row_h - LABEL_SIZE * 1.2) / 2.0,
                &name,
                LABEL_SIZE,
                false,
                LABEL_COLOR,
            );
            y += row_h;
        }
        width + 4.0
    }

    /// Bottom legend row, centered; returns the height consumed.
    fn draw_legend_row(
        &mut self,
        groups: &[&PlotGroup],
        left: f32,
        right: f32,
        bottom: f32,
    ) -> f32 {
        self.legend_row_impl(groups, left, right, bottom - LABEL_SIZE * 1.4)
    }

    fn draw_legend_row_at(
        &mut self,
        groups: &[&PlotGroup],
        left: f32,
        right: f32,
        top: f32,
    ) -> f32 {
        self.legend_row_impl(groups, left, right, top)
    }

    fn legend_row_impl(&mut self, groups: &[&PlotGroup], left: f32, right: f32, y: f32) -> f32 {
        let entries = self.legend_entries(groups);
        if entries.is_empty() {
            return 0.0;
        }
        let widths: Vec<f32> = entries
            .iter()
            .map(|(n, _)| self.measure(n, LABEL_SIZE, false).0 + 16.0)
            .collect();
        let total: f32 = widths.iter().sum();
        let mut x = left + ((right - left) - total).max(0.0) / 2.0;
        for ((name, color), w) in entries.into_iter().zip(widths) {
            self.swatch(x, y + 2.0, color);
            self.text_at(x + 11.0, y, &name, LABEL_SIZE, false, LABEL_COLOR);
            x += w;
        }
        LABEL_SIZE * 1.6
    }

    fn swatch(&mut self, x: f32, y: f32, color: Rgba) {
        self.commands.push(DrawCommand::Rect {
            rect: PtRect::from_xywh(Pt::new(x), Pt::new(y), Pt::new(7.0), Pt::new(7.0)),
            color: to_rgb(color),
        });
    }

    // ── axis charts (bar / line / area / scatter) ────────────────────────

    fn draw_axis_chart(
        &mut self,
        groups: &[&PlotGroup],
        left: f32,
        top: f32,
        right: f32,
        bottom: f32,
    ) {
        let horizontal = matches!(groups[0].kind, PlotKind::Bar { horizontal: true });
        let scatter = matches!(groups[0].kind, PlotKind::Scatter);
        let percent = groups
            .iter()
            .any(|g| g.grouping == ChartGrouping::PercentStacked);

        // Value range across all groups (stacking sums within a category).
        // `self.chart` is copied out so the axis borrows live on the chart,
        // not on `self` — the emit calls below need `&mut self`.
        let chart = self.chart;
        let (vmin, vmax) = value_range(groups);
        // A scatter plot has *two* value axes; the y one sits left/right,
        // the x one bottom/top. For category charts the single value axis is
        // whichever exists.
        let val_axis = value_axis_for(chart, scatter, false);
        let scale = auto_scale(
            vmin,
            vmax,
            val_axis.and_then(|a| a.min),
            val_axis.and_then(|a| a.max),
            percent,
        );

        // Category slots (or the x scale for scatter).
        let n_cats = groups
            .iter()
            .flat_map(|g| &g.series)
            .map(|s| s.values.len().max(s.categories.len()))
            .max()
            .unwrap_or(0)
            .max(1);
        let x_scale = scatter.then(|| {
            let (xmin, xmax) = x_range(groups);
            let x_axis = value_axis_for(chart, true, true);
            auto_scale(
                xmin,
                xmax,
                x_axis.and_then(|a| a.min),
                x_axis.and_then(|a| a.max),
                false,
            )
        });

        // Reserve room for labels: value labels on the value side, category
        // labels under (or beside, for horizontal bars) the plot.
        let val_labels: Vec<String> = ticks(&scale)
            .map(|v| fmt_num(v, scale.unit, percent))
            .collect();
        let val_label_w = val_labels
            .iter()
            .map(|t| self.measure(t, LABEL_SIZE, false).0)
            .fold(0.0, f32::max);
        let cat_axis = axis_of(chart, ChartAxisKind::Category);
        let cat_labels_shown = cat_axis
            .map(|a| !a.deleted && !a.labels_hidden)
            .unwrap_or(true);
        let val_labels_shown = val_axis
            .map(|a| !a.deleted && !a.labels_hidden)
            .unwrap_or(true);

        let label_h = LABEL_SIZE * 1.3;
        let (plot_left, plot_right, plot_top, plot_bottom) = if horizontal {
            // Bars grow rightward: values along x, categories along y.
            let cat_w = if cat_labels_shown {
                self.max_cat_label_width(groups) + 4.0
            } else {
                0.0
            };
            (
                left + cat_w,
                right - 2.0,
                top + 2.0,
                bottom - if val_labels_shown { label_h } else { 0.0 },
            )
        } else {
            (
                left + if val_labels_shown {
                    val_label_w + 4.0
                } else {
                    0.0
                },
                right - 2.0,
                top + 2.0,
                bottom - if cat_labels_shown { label_h } else { 0.0 },
            )
        };
        if plot_right <= plot_left || plot_bottom <= plot_top {
            return;
        }

        // Gridlines + value labels along the value axis.
        let gridlines = val_axis.map(|a| a.major_gridlines).unwrap_or(true);
        let reversed = val_axis.map(|a| a.reversed).unwrap_or(false);
        let to_val = |v: f64| -> f32 {
            let f = ((v - scale.min) / (scale.max - scale.min)) as f32;
            let f = if reversed { 1.0 - f } else { f };
            if horizontal {
                plot_left + f * (plot_right - plot_left)
            } else {
                plot_bottom - f * (plot_bottom - plot_top)
            }
        };
        for (tick, label) in ticks(&scale).zip(&val_labels) {
            let pos = to_val(tick);
            if gridlines {
                if horizontal {
                    self.line(pos, plot_top, pos, plot_bottom, CHROME_COLOR, 0.75);
                } else {
                    self.line(plot_left, pos, plot_right, pos, CHROME_COLOR, 0.75);
                }
            }
            if val_labels_shown {
                let (tw, m) = self.measure(label, LABEL_SIZE, false);
                if horizontal {
                    self.text_at(
                        pos - tw / 2.0,
                        plot_bottom + 2.0,
                        label,
                        LABEL_SIZE,
                        false,
                        LABEL_COLOR,
                    );
                } else {
                    let th = m.ascent.raw() + m.descent.raw();
                    self.text_at(
                        plot_left - tw - 4.0,
                        pos - th / 2.0,
                        label,
                        LABEL_SIZE,
                        false,
                        LABEL_COLOR,
                    );
                }
            }
        }

        // Axis lines.
        let cat_axis_drawn = cat_axis.map(|a| !a.deleted).unwrap_or(true);
        if cat_axis_drawn {
            if horizontal {
                self.line(
                    plot_left,
                    plot_top,
                    plot_left,
                    plot_bottom,
                    CHROME_COLOR,
                    0.75,
                );
            } else {
                self.line(
                    plot_left,
                    plot_bottom,
                    plot_right,
                    plot_bottom,
                    CHROME_COLOR,
                    0.75,
                );
            }
        }

        // Category labels at slot centers.
        if cat_labels_shown && !scatter {
            let cats = first_categories(groups, n_cats);
            let span = if horizontal {
                plot_bottom - plot_top
            } else {
                plot_right - plot_left
            };
            let slot = span / n_cats as f32;
            // Skip labels that cannot fit side by side.
            let max_w = cats
                .iter()
                .map(|c| self.measure(c, LABEL_SIZE, false).0)
                .fold(0.0, f32::max);
            let step = if horizontal {
                1
            } else {
                ((max_w + 4.0) / slot).ceil().max(1.0) as usize
            };
            for (k, cat) in cats.iter().enumerate() {
                if k % step != 0 {
                    continue;
                }
                let slot_index = if horizontal { n_cats - 1 - k } else { k };
                let center = slot * (slot_index as f32 + 0.5);
                let (tw, m) = self.measure(cat, LABEL_SIZE, false);
                if horizontal {
                    let th = m.ascent.raw() + m.descent.raw();
                    self.text_at(
                        plot_left - tw - 4.0,
                        plot_top + center - th / 2.0,
                        cat,
                        LABEL_SIZE,
                        false,
                        LABEL_COLOR,
                    );
                } else {
                    self.text_at(
                        plot_left + center - tw / 2.0,
                        plot_bottom + 2.0,
                        cat,
                        LABEL_SIZE,
                        false,
                        LABEL_COLOR,
                    );
                }
            }
        }

        // Scatter x labels.
        if let Some(ref xs) = x_scale {
            let to_x = |v: f64| -> f32 {
                plot_left + ((v - xs.min) / (xs.max - xs.min)) as f32 * (plot_right - plot_left)
            };
            for tick in ticks(xs) {
                let label = fmt_num(tick, xs.unit, false);
                let (tw, _) = self.measure(&label, LABEL_SIZE, false);
                self.text_at(
                    to_x(tick) - tw / 2.0,
                    plot_bottom + 2.0,
                    &label,
                    LABEL_SIZE,
                    false,
                    LABEL_COLOR,
                );
            }
        }

        // The data.
        let plot = PtRect::from_xywh(
            Pt::new(plot_left),
            Pt::new(plot_top),
            Pt::new(plot_right - plot_left),
            Pt::new(plot_bottom - plot_top),
        );
        for group in groups {
            match &group.kind {
                PlotKind::Bar { horizontal } => {
                    self.draw_bars(group, plot, *horizontal, &scale, reversed, n_cats)
                }
                PlotKind::Line => self.draw_lines(group, plot, &scale, reversed, n_cats, false),
                PlotKind::Area => self.draw_lines(group, plot, &scale, reversed, n_cats, true),
                PlotKind::Scatter => {
                    if let Some(ref xs) = x_scale {
                        self.draw_scatter(group, plot, &scale, xs, reversed);
                    }
                }
                PlotKind::Pie | PlotKind::Doughnut | PlotKind::Unsupported(_) => {}
            }
        }
    }

    fn max_cat_label_width(&self, groups: &[&PlotGroup]) -> f32 {
        let n = groups
            .iter()
            .flat_map(|g| &g.series)
            .map(|s| s.categories.len())
            .max()
            .unwrap_or(0);
        first_categories(groups, n)
            .iter()
            .map(|c| self.measure(c, LABEL_SIZE, false).0)
            .fold(0.0, f32::max)
    }

    fn draw_bars(
        &mut self,
        group: &PlotGroup,
        plot: PtRect,
        horizontal: bool,
        scale: &Scale,
        reversed: bool,
        n_cats: usize,
    ) {
        let stacked = matches!(
            group.grouping,
            ChartGrouping::Stacked | ChartGrouping::PercentStacked
        );
        let n_series = group.series.len().max(1);
        let gap = group.gap_width.unwrap_or(150) as f32 / 100.0;
        // §21.2.2.75/§21.2.2.82: Word writes overlap 100 for stacked groups;
        // apply that default even when the element is absent.
        let overlap = group
            .overlap
            .map(|o| o as f32 / 100.0)
            .unwrap_or(if stacked { 1.0 } else { 0.0 });

        let span = if horizontal {
            plot.size.height.raw()
        } else {
            plot.size.width.raw()
        };
        let slot = span / n_cats as f32;
        let bar_w = if stacked {
            slot / (1.0 + gap)
        } else {
            slot / (n_series as f32 - (n_series as f32 - 1.0) * overlap + gap)
        };

        let to_val = |v: f64| -> f32 {
            let f = ((v - scale.min) / (scale.max - scale.min)).clamp(0.0, 1.0) as f32;
            let f = if reversed { 1.0 - f } else { f };
            if horizontal {
                plot.origin.x.raw() + f * plot.size.width.raw()
            } else {
                plot.origin.y.raw() + plot.size.height.raw() - f * plot.size.height.raw()
            }
        };

        // Per-category running totals for stacking, positive and negative.
        let mut pos_totals = vec![0.0f64; n_cats];
        let mut neg_totals = vec![0.0f64; n_cats];
        let cat_sums: Vec<f64> = (0..n_cats)
            .map(|k| {
                group
                    .series
                    .iter()
                    .filter_map(|s| s.values.get(k).copied().flatten())
                    .map(f64::abs)
                    .sum()
            })
            .collect();

        for (i, series) in group.series.iter().enumerate() {
            let color = to_rgb(self.series_color(series));
            for k in 0..n_cats {
                let Some(raw) = series.values.get(k).copied().flatten() else {
                    continue;
                };
                let v = if group.grouping == ChartGrouping::PercentStacked {
                    if cat_sums[k] == 0.0 {
                        0.0
                    } else {
                        raw / cat_sums[k] * 100.0
                    }
                } else {
                    raw
                };
                let (start_val, end_val) = if stacked {
                    let totals = if v >= 0.0 {
                        &mut pos_totals
                    } else {
                        &mut neg_totals
                    };
                    let start = totals[k];
                    totals[k] += v;
                    (start, totals[k])
                } else {
                    (0.0f64.clamp(scale.min, scale.max), v)
                };
                let a = to_val(start_val);
                let b = to_val(end_val);
                // §21.2: a `bar`-direction chart runs its category axis
                // bottom-up (Excel/Word's reading), so slot k counts from
                // the plot's bottom edge there.
                let slot_index = if horizontal { n_cats - 1 - k } else { k };
                let across = plot_offset(&plot, horizontal)
                    + slot * slot_index as f32
                    + gap * bar_w / 2.0
                    + if stacked {
                        0.0
                    } else {
                        i as f32 * (1.0 - overlap) * bar_w
                    };
                let rect = if horizontal {
                    PtRect::from_xywh(
                        Pt::new(a.min(b)),
                        Pt::new(across),
                        Pt::new((a - b).abs()),
                        Pt::new(bar_w),
                    )
                } else {
                    PtRect::from_xywh(
                        Pt::new(across),
                        Pt::new(a.min(b)),
                        Pt::new(bar_w),
                        Pt::new((a - b).abs()),
                    )
                };
                if rect.size.width > Pt::ZERO && rect.size.height > Pt::ZERO {
                    self.commands.push(DrawCommand::Rect { rect, color });
                }
            }
        }
    }

    fn draw_lines(
        &mut self,
        group: &PlotGroup,
        plot: PtRect,
        scale: &Scale,
        reversed: bool,
        n_cats: usize,
        area: bool,
    ) {
        let stacked = matches!(
            group.grouping,
            ChartGrouping::Stacked | ChartGrouping::PercentStacked
        );
        let slot = plot.size.width.raw() / n_cats as f32;
        let to_y = |v: f64| -> f32 {
            let f = ((v - scale.min) / (scale.max - scale.min)).clamp(0.0, 1.0) as f32;
            let f = if reversed { 1.0 - f } else { f };
            plot.origin.y.raw() + plot.size.height.raw() - f * plot.size.height.raw()
        };
        let x_of = |k: usize| plot.origin.x.raw() + slot * (k as f32 + 0.5);
        let baseline = to_y(0.0f64.clamp(scale.min, scale.max));

        let cat_sums: Vec<f64> = (0..n_cats)
            .map(|k| {
                group
                    .series
                    .iter()
                    .filter_map(|s| s.values.get(k).copied().flatten())
                    .map(f64::abs)
                    .sum()
            })
            .collect();
        let mut totals = vec![0.0f64; n_cats];

        for series in &group.series {
            let color = self.series_color(series);
            // Build point list; None breaks the polyline (dispBlanksAs gap).
            let mut points: Vec<Option<(f32, f32)>> = Vec::with_capacity(n_cats);
            // The stack level *under* this series — snapshotted before the
            // point loop below adds this series into `totals`.
            let prev_totals = totals.clone();
            for k in 0..n_cats {
                match series.values.get(k).copied().flatten() {
                    Some(raw) => {
                        let v = if group.grouping == ChartGrouping::PercentStacked {
                            if cat_sums[k] == 0.0 {
                                0.0
                            } else {
                                raw / cat_sums[k] * 100.0
                            }
                        } else {
                            raw
                        };
                        let y = if stacked {
                            totals[k] += v;
                            to_y(totals[k])
                        } else {
                            to_y(v)
                        };
                        points.push(Some((x_of(k), y)));
                    }
                    None => points.push(None),
                }
            }

            if area {
                // Filled polygon down to the previous stack level (or the
                // zero baseline), one polygon per contiguous run.
                for run in contiguous_runs(&points) {
                    let mut verbs = Vec::new();
                    let (first_k, pts) = run;
                    verbs.push(PathVerb::MoveTo(pt(pts[0].0, pts[0].1)));
                    for p in &pts[1..] {
                        verbs.push(PathVerb::LineTo(pt(p.0, p.1)));
                    }
                    // Back along the base.
                    for (j, p) in pts.iter().enumerate().rev() {
                        let k = first_k + j;
                        let base_y = if stacked {
                            to_y(prev_totals[k])
                        } else {
                            baseline
                        };
                        verbs.push(PathVerb::LineTo(pt(p.0, base_y)));
                    }
                    verbs.push(PathVerb::Close);
                    self.push_path(
                        vec![SubPath {
                            verbs,
                            fill_mode: crate::model::PathFillMode::Norm,
                            stroked: false,
                        }],
                        ResolvedFill::Solid(color),
                        None,
                    );
                }
            } else {
                let stroke = self
                    .explicit_stroke(series)
                    .unwrap_or_else(|| series_stroke(color));
                for (_, pts) in contiguous_runs(&points) {
                    if pts.len() < 2 {
                        continue;
                    }
                    let mut verbs = vec![PathVerb::MoveTo(pt(pts[0].0, pts[0].1))];
                    for p in &pts[1..] {
                        verbs.push(PathVerb::LineTo(pt(p.0, p.1)));
                    }
                    self.push_path(
                        vec![SubPath {
                            verbs,
                            fill_mode: crate::model::PathFillMode::None,
                            stroked: true,
                        }],
                        ResolvedFill::None,
                        Some(stroke.clone()),
                    );
                }
                // Markers.
                if let Some(marker) = series.marker.filter(|m| *m != ChartMarker::None) {
                    for p in points.iter().flatten() {
                        self.draw_marker(*p, marker, color);
                    }
                }
            }
        }
    }

    fn explicit_stroke(&self, series: &ChartSeries) -> Option<ResolvedStroke> {
        let props = series.shape_properties.as_ref()?;
        props.outline.as_ref()?;
        let visuals = resolve_shape_visuals(
            Some(props),
            None,
            None,
            None,
            self.ctx.resolved.theme.as_ref(),
        );
        visuals.stroke
    }

    fn draw_scatter(
        &mut self,
        group: &PlotGroup,
        plot: PtRect,
        y_scale: &Scale,
        x_scale: &Scale,
        reversed: bool,
    ) {
        let to_y = |v: f64| -> f32 {
            let f = ((v - y_scale.min) / (y_scale.max - y_scale.min)).clamp(0.0, 1.0) as f32;
            let f = if reversed { 1.0 - f } else { f };
            plot.origin.y.raw() + plot.size.height.raw() - f * plot.size.height.raw()
        };
        let to_x = |v: f64| -> f32 {
            let f = ((v - x_scale.min) / (x_scale.max - x_scale.min)).clamp(0.0, 1.0) as f32;
            plot.origin.x.raw() + f * plot.size.width.raw()
        };
        for series in &group.series {
            let color = self.series_color(series);
            // A scatter series without `c:xVal` plots against the point
            // numbers 1..n, Excel's own default.
            let fallback_x: Vec<Option<f64>>;
            let xs = if series.x_values.is_empty() {
                fallback_x = (1..=series.values.len()).map(|i| Some(i as f64)).collect();
                &fallback_x
            } else {
                &series.x_values
            };
            let points: Vec<Option<(f32, f32)>> = xs
                .iter()
                .zip(&series.values)
                .map(|(x, y)| match (x, y) {
                    (Some(x), Some(y)) => Some((to_x(*x), to_y(*y))),
                    _ => None,
                })
                .collect();
            // A line only when the series asks for one explicitly.
            if let Some(stroke) = self.explicit_stroke(series) {
                for (_, pts) in contiguous_runs(&points) {
                    if pts.len() < 2 {
                        continue;
                    }
                    let mut verbs = vec![PathVerb::MoveTo(pt(pts[0].0, pts[0].1))];
                    for p in &pts[1..] {
                        verbs.push(PathVerb::LineTo(pt(p.0, p.1)));
                    }
                    self.push_path(
                        vec![SubPath {
                            verbs,
                            fill_mode: crate::model::PathFillMode::None,
                            stroked: true,
                        }],
                        ResolvedFill::None,
                        Some(stroke.clone()),
                    );
                }
            }
            let marker = series.marker.unwrap_or(ChartMarker::Circle);
            if marker != ChartMarker::None {
                for p in points.iter().flatten() {
                    self.draw_marker(*p, marker, color);
                }
            }
        }
    }

    fn draw_marker(&mut self, (x, y): (f32, f32), marker: ChartMarker, color: Rgba) {
        let r = 2.5;
        let verbs = match marker {
            ChartMarker::Square => vec![
                PathVerb::MoveTo(pt(x - r, y - r)),
                PathVerb::LineTo(pt(x + r, y - r)),
                PathVerb::LineTo(pt(x + r, y + r)),
                PathVerb::LineTo(pt(x - r, y + r)),
                PathVerb::Close,
            ],
            ChartMarker::Diamond => vec![
                PathVerb::MoveTo(pt(x, y - r)),
                PathVerb::LineTo(pt(x + r, y)),
                PathVerb::LineTo(pt(x, y + r)),
                PathVerb::LineTo(pt(x - r, y)),
                PathVerb::Close,
            ],
            ChartMarker::Triangle => vec![
                PathVerb::MoveTo(pt(x, y - r)),
                PathVerb::LineTo(pt(x + r, y + r)),
                PathVerb::LineTo(pt(x - r, y + r)),
                PathVerb::Close,
            ],
            // Circle (and anything else): four quarter-arc cubics.
            _ => circle_verbs(x, y, r),
        };
        self.push_path(
            vec![SubPath {
                verbs,
                fill_mode: crate::model::PathFillMode::Norm,
                stroked: false,
            }],
            ResolvedFill::Solid(color),
            None,
        );
    }

    // ── pie / doughnut ───────────────────────────────────────────────────

    fn draw_pie(&mut self, group: &PlotGroup, left: f32, top: f32, right: f32, bottom: f32) {
        let Some(series) = group.series.first() else {
            return;
        };
        let values: Vec<(usize, f64)> = series
            .values
            .iter()
            .enumerate()
            .filter_map(|(i, v)| v.map(|v| (i, v.abs())))
            .filter(|(_, v)| *v > 0.0)
            .collect();
        let total: f64 = values.iter().map(|(_, v)| v).sum();
        if total <= 0.0 {
            return;
        }

        let cx = (left + right) / 2.0;
        let cy = (top + bottom) / 2.0;
        let radius = ((right - left).min(bottom - top) / 2.0 - 4.0).max(4.0);
        let hole = match group.kind {
            PlotKind::Doughnut => radius * group.hole_size.unwrap_or(75).min(90) as f32 / 100.0,
            _ => 0.0,
        };

        // `c:varyColors` defaults on for the pie family; off colors every
        // slice with the series color.
        let vary = group.vary_colors.unwrap_or(true);
        // `c:firstSliceAng`: degrees clockwise from 12 o'clock.
        let mut angle = -90.0 + group.first_slice_angle.unwrap_or(0) as f32;
        for (point_idx, v) in &values {
            let sweep = (v / total * 360.0) as f32;
            let color = if vary {
                self.point_color(series, *point_idx)
            } else {
                self.series_color(series)
            };
            let mut verbs = Vec::new();
            if hole > 0.0 {
                verbs.push(PathVerb::MoveTo(polar(cx, cy, radius, angle)));
                arc_verbs(&mut verbs, cx, cy, radius, angle, sweep);
                verbs.push(PathVerb::LineTo(polar(cx, cy, hole, angle + sweep)));
                arc_verbs(&mut verbs, cx, cy, hole, angle + sweep, -sweep);
            } else {
                verbs.push(PathVerb::MoveTo(pt(cx, cy)));
                verbs.push(PathVerb::LineTo(polar(cx, cy, radius, angle)));
                arc_verbs(&mut verbs, cx, cy, radius, angle, sweep);
            }
            verbs.push(PathVerb::Close);
            self.push_path(
                vec![SubPath {
                    verbs,
                    fill_mode: crate::model::PathFillMode::Norm,
                    stroked: true,
                }],
                ResolvedFill::Solid(color),
                // The thin white separator modern Word draws between slices.
                Some(ResolvedStroke {
                    width: Pt::new(0.75),
                    color: Rgba {
                        r: 1.0,
                        g: 1.0,
                        b: 1.0,
                        a: 1.0,
                    },
                    dash: ResolvedDashPattern::Solid,
                    cap: ResolvedLineCap::Butt,
                    join: ResolvedLineJoin::Round,
                }),
            );
            angle += sweep;
        }
    }
}

// ── free helpers ──────────────────────────────────────────────────────────

fn axis_of(chart: &ChartSpace, kind: ChartAxisKind) -> Option<&ChartAxis> {
    chart.axes.iter().find(|a| a.kind == kind)
}

/// The value axis that scales a given direction. Category charts have one;
/// a scatter has two, told apart by `c:axPos` — bottom/top is x, left/right
/// is y — with document order (x first, Word's own) as the tiebreak.
fn value_axis_for(chart: &ChartSpace, scatter: bool, want_x: bool) -> Option<&ChartAxis> {
    if !scatter {
        return axis_of(chart, ChartAxisKind::Value);
    }
    let vals: Vec<&ChartAxis> = chart
        .axes
        .iter()
        .filter(|a| a.kind == ChartAxisKind::Value)
        .collect();
    let horizontal = |a: &&ChartAxis| {
        matches!(
            a.position,
            crate::model::ChartAxisPosition::Bottom | crate::model::ChartAxisPosition::Top
        )
    };
    if want_x {
        vals.iter()
            .find(|a| horizontal(a))
            .copied()
            .or_else(|| vals.first().copied())
    } else {
        vals.iter()
            .find(|a| !horizontal(a))
            .copied()
            .or_else(|| vals.get(1).copied())
    }
}

fn pt(x: f32, y: f32) -> PtOffset {
    PtOffset::new(Pt::new(x), Pt::new(y))
}

fn plot_offset(plot: &PtRect, horizontal: bool) -> f32 {
    if horizontal {
        plot.origin.y.raw()
    } else {
        plot.origin.x.raw()
    }
}

fn to_rgb(c: Rgba) -> RgbColor {
    RgbColor {
        r: (c.r.clamp(0.0, 1.0) * 255.0).round() as u8,
        g: (c.g.clamp(0.0, 1.0) * 255.0).round() as u8,
        b: (c.b.clamp(0.0, 1.0) * 255.0).round() as u8,
    }
}

/// Luminance-modulate-and-offset, the colors1.xml cycle variation:
/// `c' = c·mod + off`.
fn lum(c: Rgba, m: f32, o: f32) -> Rgba {
    Rgba {
        r: (c.r * m + o).clamp(0.0, 1.0),
        g: (c.g * m + o).clamp(0.0, 1.0),
        b: (c.b * m + o).clamp(0.0, 1.0),
        a: c.a,
    }
}

fn series_stroke(color: Rgba) -> ResolvedStroke {
    ResolvedStroke {
        width: Pt::new(SERIES_LINE_WIDTH),
        color,
        dash: ResolvedDashPattern::Solid,
        cap: ResolvedLineCap::Round,
        join: ResolvedLineJoin::Round,
    }
}

fn chrome_stroke() -> ResolvedStroke {
    ResolvedStroke {
        width: Pt::new(0.75),
        color: Rgba {
            r: 0xD9 as f32 / 255.0,
            g: 0xD9 as f32 / 255.0,
            b: 0xD9 as f32 / 255.0,
            a: 1.0,
        },
        dash: ResolvedDashPattern::Solid,
        cap: ResolvedLineCap::Butt,
        join: ResolvedLineJoin::Miter,
    }
}

fn polar(cx: f32, cy: f32, r: f32, deg: f32) -> PtOffset {
    let rad = deg.to_radians();
    pt(cx + r * rad.cos(), cy + r * rad.sin())
}

/// Append cubic-Bézier arc segments approximating a circular arc (y-down,
/// positive sweep clockwise), ≤90° per cubic. The pen must sit at the arc's
/// start.
fn arc_verbs(verbs: &mut Vec<PathVerb>, cx: f32, cy: f32, r: f32, start_deg: f32, sweep_deg: f32) {
    let steps = (sweep_deg.abs() / 90.0).ceil().max(1.0) as usize;
    let step = sweep_deg / steps as f32;
    let mut a = start_deg;
    for _ in 0..steps {
        let a0 = a.to_radians();
        let a1 = (a + step).to_radians();
        let k = 4.0 / 3.0 * ((a1 - a0) / 4.0).tan();
        let p0 = (cx + r * a0.cos(), cy + r * a0.sin());
        let p1 = (cx + r * a1.cos(), cy + r * a1.sin());
        let c1 = pt(p0.0 - k * r * a0.sin(), p0.1 + k * r * a0.cos());
        let c2 = pt(p1.0 + k * r * a1.sin(), p1.1 - k * r * a1.cos());
        verbs.push(PathVerb::CubicTo(c1, c2, pt(p1.0, p1.1)));
        a += step;
    }
}

fn circle_verbs(cx: f32, cy: f32, r: f32) -> Vec<PathVerb> {
    let mut verbs = vec![PathVerb::MoveTo(pt(cx + r, cy))];
    arc_verbs(&mut verbs, cx, cy, r, 0.0, 360.0);
    verbs.push(PathVerb::Close);
    verbs
}

/// Contiguous `Some` runs of a broken point list, with their start index.
fn contiguous_runs(points: &[Option<(f32, f32)>]) -> Vec<(usize, Vec<(f32, f32)>)> {
    let mut out = Vec::new();
    let mut run: Vec<(f32, f32)> = Vec::new();
    let mut start = 0;
    for (i, p) in points.iter().enumerate() {
        match p {
            Some(p) => {
                if run.is_empty() {
                    start = i;
                }
                run.push(*p);
            }
            None => {
                if !run.is_empty() {
                    out.push((start, std::mem::take(&mut run)));
                }
            }
        }
    }
    if !run.is_empty() {
        out.push((start, run));
    }
    out
}

/// The value range the plot must cover, stacking within categories.
fn value_range(groups: &[&PlotGroup]) -> (f64, f64) {
    let mut min = f64::INFINITY;
    let mut max = f64::NEG_INFINITY;
    for g in groups {
        let stacked = matches!(
            g.grouping,
            ChartGrouping::Stacked | ChartGrouping::PercentStacked
        );
        let n = g.series.iter().map(|s| s.values.len()).max().unwrap_or(0);
        if stacked {
            for k in 0..n {
                let (mut pos, mut neg) = (0.0, 0.0);
                let sum: f64 = g
                    .series
                    .iter()
                    .filter_map(|s| s.values.get(k).copied().flatten())
                    .map(f64::abs)
                    .sum();
                for s in &g.series {
                    if let Some(v) = s.values.get(k).copied().flatten() {
                        let v = if g.grouping == ChartGrouping::PercentStacked {
                            if sum == 0.0 {
                                0.0
                            } else {
                                v / sum * 100.0
                            }
                        } else {
                            v
                        };
                        if v >= 0.0 {
                            pos += v;
                        } else {
                            neg += v;
                        }
                    }
                }
                min = min.min(neg).min(0.0);
                max = max.max(pos).max(0.0);
            }
        } else {
            for s in &g.series {
                for v in s.values.iter().flatten() {
                    min = min.min(*v);
                    max = max.max(*v);
                }
            }
        }
    }
    if !min.is_finite() || !max.is_finite() {
        (0.0, 1.0)
    } else {
        (min, max)
    }
}

fn x_range(groups: &[&PlotGroup]) -> (f64, f64) {
    let mut min = f64::INFINITY;
    let mut max = f64::NEG_INFINITY;
    for g in groups {
        for s in &g.series {
            for v in s.x_values.iter().flatten() {
                min = min.min(*v);
                max = max.max(*v);
            }
        }
    }
    if !min.is_finite() || !max.is_finite() {
        (0.0, 1.0)
    } else {
        (min, max)
    }
}

/// The categories of the first series that has any — the shared category
/// axis, padded with point numbers where labels are blank.
fn first_categories(groups: &[&PlotGroup], n: usize) -> Vec<String> {
    let cats = groups
        .iter()
        .flat_map(|g| &g.series)
        .map(|s| &s.categories)
        .find(|c| !c.is_empty());
    (0..n)
        .map(|k| {
            cats.and_then(|c| c.get(k).cloned().flatten())
                .unwrap_or_else(|| (k + 1).to_string())
        })
        .collect()
}

/// A resolved value-axis scale.
#[derive(Clone, Copy, Debug, PartialEq)]
struct Scale {
    min: f64,
    max: f64,
    unit: f64,
}

fn ticks(scale: &Scale) -> impl Iterator<Item = f64> + '_ {
    // Belt to `auto_scale`'s braces: never emit more than a sane number of
    // gridlines/labels whatever the inputs conspired to.
    let n = (((scale.max - scale.min) / scale.unit).round() as usize).min(64);
    (0..=n).map(move |i| scale.min + scale.unit * i as f64)
}

/// Excel's automatic value-axis bounds (as Peltier documents them and
/// LibreOffice's `ScaleAutomatism` implements them): a 5% margin beyond the
/// data, bounds snapped outward to a 1/2/5×10ᵏ major unit, and zero pinned
/// as the minimum for all-positive data unless the data floats high
/// (`min ≥ ⅚·max`).
fn auto_scale(
    data_min: f64,
    data_max: f64,
    manual_min: Option<f64>,
    manual_max: Option<f64>,
    percent: bool,
) -> Scale {
    if percent {
        return Scale {
            min: manual_min.unwrap_or(0.0),
            max: manual_max.unwrap_or(100.0),
            unit: 20.0,
        };
    }
    let (mut lo, mut hi) = (data_min.min(data_max), data_max.max(data_min));
    if lo == hi {
        if lo == 0.0 {
            hi = 1.0;
        } else if lo > 0.0 {
            lo = 0.0;
        } else {
            hi = 0.0;
        }
    }
    let span = hi - lo;
    let (target_lo, target_hi) = if lo >= 0.0 {
        let min = if lo >= hi * 5.0 / 6.0 {
            lo - span / 20.0
        } else {
            0.0
        };
        (min, hi + span / 20.0)
    } else if hi <= 0.0 {
        let max = if hi <= lo * 5.0 / 6.0 {
            hi + span / 20.0
        } else {
            0.0
        };
        (lo - span / 20.0, max)
    } else {
        (lo - span / 20.0, hi + span / 20.0)
    };

    // Excel lands 5–10 major ticks; dividing by 8 before snapping to the
    // 1/2/5 ladder reproduces its picks for common ranges (0–10 → 2,
    // 0–25 → 5, 0–100 → 20).
    let unit = nice_unit((target_hi - target_lo) / 8.0);
    // Manual bounds are attacker-adjacent input: non-finite values are
    // ignored, and when either bound is manual the unit is recomputed from
    // the *final* span — a data-derived unit against a huge manual bound
    // would otherwise tick the axis millions of times.
    let manual_min = manual_min.filter(|v| v.is_finite());
    let manual_max = manual_max.filter(|v| v.is_finite());
    let min = manual_min.unwrap_or_else(|| (target_lo / unit).floor() * unit);
    let max = manual_max.unwrap_or_else(|| (target_hi / unit).ceil() * unit);
    let (min, max) = if max > min {
        (min, max)
    } else {
        (min, min + unit)
    };
    let unit = if manual_min.is_some() || manual_max.is_some() {
        nice_unit((max - min) / 8.0)
    } else {
        unit
    };
    Scale { min, max, unit }
}

/// Round up to the nearest 1/2/5×10ᵏ.
fn nice_unit(raw: f64) -> f64 {
    if raw <= 0.0 || !raw.is_finite() {
        return 1.0;
    }
    let mag = 10f64.powf(raw.log10().floor());
    let norm = raw / mag;
    let nice = if norm <= 1.0 {
        1.0
    } else if norm <= 2.0 {
        2.0
    } else if norm <= 5.0 {
        5.0
    } else {
        10.0
    };
    nice * mag
}

/// Format an axis value: integers when the unit is whole, else just enough
/// decimals for the unit.
fn fmt_num(v: f64, unit: f64, percent: bool) -> String {
    let suffix = if percent { "%" } else { "" };
    if unit >= 1.0 || v == 0.0 {
        format!("{}{suffix}", v.round() as i64)
    } else {
        let decimals = (-unit.log10().floor()) as usize;
        format!("{v:.decimals$}{suffix}")
    }
}

/// All run text of a DrawingML body, breaks as spaces.
fn body_text(body: &DrawingTextBody) -> String {
    let mut out = String::new();
    for p in &body.paragraphs {
        for run in &p.runs {
            match run {
                DrawingTextRun::Text { text, .. } => out.push_str(text),
                DrawingTextRun::Break => out.push(' '),
            }
        }
    }
    out
}

fn first_run_style(body: &DrawingTextBody) -> Option<(f32, bool)> {
    for p in &body.paragraphs {
        let def = p.default_run.as_ref();
        for run in &p.runs {
            if let DrawingTextRun::Text { props, .. } = run {
                let size = props
                    .size
                    .or_else(|| def.and_then(|d| d.size))
                    .map(|s| s.raw() as f32 / 100.0)
                    .unwrap_or(TITLE_SIZE);
                let bold = props
                    .bold
                    .or_else(|| def.and_then(|d| d.bold))
                    .unwrap_or(false);
                return Some((size, bold));
            }
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn auto_scale_pins_zero_for_ordinary_positive_data() {
        let s = auto_scale(3.0, 17.0, None, None, false);
        assert_eq!(s.min, 0.0);
        assert!(s.max >= 17.0 + 0.7 - 1e-9, "5% margin before rounding");
        assert_eq!(s.max % s.unit, 0.0, "max lands on a unit multiple");
    }

    /// Excel's high-floating rule: data in a narrow high band does not force
    /// the axis to zero.
    #[test]
    fn auto_scale_releases_zero_when_data_floats_high() {
        let s = auto_scale(95.0, 100.0, None, None, false);
        assert!(s.min > 0.0, "min {} should float", s.min);
        assert!(s.min <= 95.0);
    }

    #[test]
    fn auto_scale_units_are_1_2_5_times_ten() {
        for (span, want) in [(10.0, 2.0), (25.0, 5.0), (100.0, 20.0), (2.4, 0.5)] {
            let s = auto_scale(0.0, span, None, None, false);
            assert_eq!(s.unit, want, "span {span}");
        }
    }

    #[test]
    fn auto_scale_manual_bounds_win() {
        let s = auto_scale(3.0, 17.0, Some(5.0), Some(40.0), false);
        assert_eq!((s.min, s.max), (5.0, 40.0));
    }

    #[test]
    fn percent_scale_is_fixed() {
        let s = auto_scale(-5.0, 300.0, None, None, true);
        assert_eq!((s.min, s.max, s.unit), (0.0, 100.0, 20.0));
    }

    #[test]
    fn nice_unit_rounds_up_the_ladder() {
        assert_eq!(nice_unit(0.9), 1.0);
        assert_eq!(nice_unit(1.1), 2.0);
        assert_eq!(nice_unit(3.0), 5.0);
        assert_eq!(nice_unit(7.0), 10.0);
        assert_eq!(nice_unit(30.0), 50.0);
        assert_eq!(nice_unit(0.03), 0.05);
    }

    #[test]
    fn contiguous_runs_break_at_gaps() {
        let pts = [Some((0.0, 0.0)), Some((1.0, 1.0)), None, Some((3.0, 3.0))];
        let runs = contiguous_runs(&pts);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].0, 0);
        assert_eq!(runs[0].1.len(), 2);
        assert_eq!(runs[1].0, 3);
    }

    #[test]
    fn value_range_sums_stacked_categories() {
        use crate::model::*;
        let g = PlotGroup {
            kind: PlotKind::Bar { horizontal: false },
            series: vec![
                ChartSeries {
                    values: vec![Some(2.0), Some(3.0)],
                    ..Default::default()
                },
                ChartSeries {
                    values: vec![Some(4.0), Some(-1.0)],
                    ..Default::default()
                },
            ],
            gap_width: None,
            overlap: None,
            grouping: ChartGrouping::Stacked,
            vary_colors: None,
            first_slice_angle: None,
            hole_size: None,
        };
        let (min, max) = value_range(&[&g]);
        assert_eq!(max, 6.0, "positive stack peaks at 2+4");
        assert_eq!(min, -1.0, "negatives stack separately");
    }

    #[test]
    fn fmt_num_matches_unit_precision() {
        assert_eq!(fmt_num(1500.0, 500.0, false), "1500");
        assert_eq!(fmt_num(0.5, 0.5, false), "0.5");
        assert_eq!(fmt_num(40.0, 20.0, true), "40%");
    }
}
