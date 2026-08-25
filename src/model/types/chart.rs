//! §21.2 DrawingML charts — the parsed subset of a chart part.
//!
//! A chart part describes series, axes and a plot area; crucially, every
//! producer that matters (Word, Excel, PowerPoint) writes the plotted values
//! into `c:numCache`/`c:strCache` on save — "the last data shown on the
//! chart" (§21.2.2.120/.191) — so the model carries those cached values and
//! never needs the embedded workbook. Data values are plain `f64`: they are
//! chart data, not document geometry, and no unit marker applies.
//!
//! What is modeled is what the renderer draws (issue #155): the 2D
//! bar/line/pie/doughnut/area/scatter families, two axes, legend and title.
//! The 3D variants parse into their 2D twins (the LibreOffice reading —
//! flattening beats refusing); everything else lands in
//! [`PlotKind::Unsupported`] with its element name, and the renderer
//! declines it visibly in the log rather than silently.

use super::drawing::{DrawingTextBody, ShapeProperties};

/// `c:chartSpace` → `c:chart`, reduced.
#[derive(Clone, Debug, Default)]
pub struct ChartSpace {
    /// `c:title` — literal rich text when present.
    pub title: Option<DrawingTextBody>,
    /// `c:autoTitleDeleted` — true means "no automatic title"; with a
    /// literal title absent and this false, Word titles a one-series chart
    /// with the series name.
    pub auto_title_deleted: bool,
    /// The plot-type groups inside `c:plotArea`, in document order; a combo
    /// chart has several sharing the axes.
    pub plot_groups: Vec<PlotGroup>,
    /// `c:catAx`/`c:valAx`/`c:dateAx`, in document order.
    pub axes: Vec<ChartAxis>,
    pub legend: Option<Legend>,
}

/// One plot-type element (`c:barChart`, `c:lineChart`, …) with its series.
#[derive(Clone, Debug)]
pub struct PlotGroup {
    pub kind: PlotKind,
    pub series: Vec<ChartSeries>,
    /// §21.2.2.75 `c:gapWidth` — gap between category clusters, percent of
    /// one bar width. Spec default 150.
    pub gap_width: Option<u32>,
    /// `c:overlap` — −100…100 percent; Word writes 100 for stacked.
    pub overlap: Option<i32>,
    pub grouping: ChartGrouping,
    /// `c:varyColors` — pie-family per-point coloring.
    pub vary_colors: Option<bool>,
    /// `c:firstSliceAng` — degrees clockwise from 12 o'clock.
    pub first_slice_angle: Option<u32>,
    /// `c:holeSize` — doughnut hole, percent of diameter.
    pub hole_size: Option<u32>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PlotKind {
    Bar {
        /// `c:barDir` — `col` grows up, `bar` grows right.
        horizontal: bool,
    },
    Line,
    Pie,
    Doughnut,
    Area,
    Scatter,
    /// Radar, surface, bubble, stock, ofPie, … — the element's local name,
    /// kept so the decline can say what it declined.
    Unsupported(String),
}

/// §21.2.2.76 ST_Grouping / ST_BarGrouping.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ChartGrouping {
    #[default]
    Clustered,
    Stacked,
    PercentStacked,
    /// Line/area `standard` — series drawn independently, like clustered.
    Standard,
}

/// One `c:ser`.
#[derive(Clone, Debug, Default)]
pub struct ChartSeries {
    /// `c:idx` — the format index that drives the automatic color cycle.
    pub idx: u32,
    /// `c:order` — plot order.
    pub order: u32,
    /// `c:tx` — the series name, from its string cache.
    pub name: Option<String>,
    /// `c:spPr` — explicit fill/line overriding the automatic color.
    pub shape_properties: Option<ShapeProperties>,
    /// `c:cat` (or `c:xVal` labels) — one entry per point; `None` for a
    /// blank cell.
    pub categories: Vec<Option<String>>,
    /// `c:val`/`c:yVal` from the number cache.
    pub values: Vec<Option<f64>>,
    /// `c:xVal` numeric cache (scatter); empty for category charts.
    pub x_values: Vec<Option<f64>>,
    /// `c:dPt` per-point property overrides, by point index.
    pub point_properties: Vec<(u32, ShapeProperties)>,
    /// Line/scatter `c:marker/c:symbol` — `None` = automatic.
    pub marker: Option<ChartMarker>,
}

/// §21.2.2.145 `c:symbol`, reduced to what draws.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ChartMarker {
    None,
    Circle,
    Square,
    Diamond,
    Triangle,
    /// Any other symbol — drawn as a circle.
    Other,
}

/// `c:catAx` / `c:valAx` / `c:dateAx` (dates read as categories).
#[derive(Clone, Debug, Default)]
pub struct ChartAxis {
    pub kind: ChartAxisKind,
    /// `c:delete` — a deleted axis still scales the plot; it just isn't
    /// drawn.
    pub deleted: bool,
    /// `c:axPos`.
    pub position: ChartAxisPosition,
    /// `c:scaling/c:orientation val="maxMin"`.
    pub reversed: bool,
    /// `c:scaling/c:min`/`c:max` manual bounds.
    pub min: Option<f64>,
    pub max: Option<f64>,
    /// `c:majorGridlines` present.
    pub major_gridlines: bool,
    /// `c:title` rich text.
    pub title: Option<DrawingTextBody>,
    /// `c:tickLblPos val="none"` hides labels while keeping the axis.
    pub labels_hidden: bool,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ChartAxisKind {
    #[default]
    Category,
    Value,
}

/// §21.2.2.5 ST_AxPos.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ChartAxisPosition {
    #[default]
    Bottom,
    Left,
    Top,
    Right,
}

/// `c:legend`.
#[derive(Clone, Debug, Default)]
pub struct Legend {
    pub position: LegendPosition,
    /// `c:overlay` — the legend floats over the plot instead of reserving
    /// space.
    pub overlay: bool,
}

/// §21.2.2.99 ST_LegendPos. Spec default `r`.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum LegendPosition {
    #[default]
    Right,
    Left,
    Top,
    Bottom,
    TopRight,
}
