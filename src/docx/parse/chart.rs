//! Chart part (§21.2) parser — the cached-data subset (issue #155).
//!
//! A chart part carries its own worksheet reference, but every producer that
//! matters writes the plotted values into `c:numCache`/`c:strCache` on save
//! — "the last data shown on the chart" — so this parser reads caches only
//! and never opens the embedded workbook. A hand-generated part with bare
//! `c:f` formulas and no caches parses to empty series, which render as an
//! empty plot rather than failing the document.
//!
//! §21.2's boolean elements (`c:delete`, `c:varyColors`, `c:overlay`,
//! `c:smooth`, `c:autoTitleDeleted`) share one trap: the element present
//! with **no** `val` attribute means *true* (CT_Boolean's default), while an
//! absent element means the property's own default. `CBoolXml` owns that
//! reading so no call site can get it wrong.
//!
//! The 3D chart kinds parse into their 2D twins — flattening beats refusing,
//! the reading LibreOffice takes — and the kinds this engine does not draw
//! (radar, surface, bubble, stock, ofPie) land in
//! [`PlotKind::Unsupported`] with their element name, so the renderer's
//! decline can say what it declined.

use serde::Deserialize;

use crate::docx::error::Result;
use crate::docx::model::{
    ChartAxis, ChartAxisKind, ChartAxisPosition, ChartGrouping, ChartMarker, ChartSeries,
    ChartSpace, DrawingTextBody, Legend, LegendPosition, PlotGroup, PlotKind,
};
use crate::model::Dup;

use super::drawing::schema::shape::SpPrXml;
use super::drawing::schema::text_body::TextBodyXml;

/// Parse a chart part.
pub fn parse_chart(data: &[u8]) -> Result<ChartSpace> {
    let parsed: ChartSpaceXml = super::serde_xml::from_xml(data)?;
    let Some(chart) = Dup::from(parsed.chart).into_value() else {
        return Ok(ChartSpace::default());
    };

    let mut plot_groups = Vec::new();
    let mut axes = Vec::new();
    if let Some(plot_area) = Dup::from(chart.plot_area).into_value() {
        for child in plot_area.children {
            match child {
                PlotAreaChildXml::BarChart(p) | PlotAreaChildXml::Bar3DChart(p) => {
                    let kind = bar_kind(&p_dir(&p));
                    plot_groups.push(p.into_group(kind));
                }
                PlotAreaChildXml::LineChart(p) | PlotAreaChildXml::Line3DChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Line));
                }
                PlotAreaChildXml::PieChart(p) | PlotAreaChildXml::Pie3DChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Pie));
                }
                PlotAreaChildXml::DoughnutChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Doughnut));
                }
                PlotAreaChildXml::AreaChart(p) | PlotAreaChildXml::Area3DChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Area));
                }
                PlotAreaChildXml::ScatterChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Scatter));
                }
                PlotAreaChildXml::OfPieChart(p) => {
                    // §21.2.2.126 pie-of-pie: the secondary plot is a
                    // presentation detail; the data reads as one pie.
                    plot_groups.push(p.into_group(PlotKind::Pie));
                }
                PlotAreaChildXml::RadarChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Unsupported("radarChart".into())));
                }
                PlotAreaChildXml::SurfaceChart(p) | PlotAreaChildXml::Surface3DChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Unsupported("surfaceChart".into())));
                }
                PlotAreaChildXml::BubbleChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Unsupported("bubbleChart".into())));
                }
                PlotAreaChildXml::StockChart(p) => {
                    plot_groups.push(p.into_group(PlotKind::Unsupported("stockChart".into())));
                }
                PlotAreaChildXml::CatAx(a) => axes.push(a.into_axis(ChartAxisKind::Category)),
                PlotAreaChildXml::DateAx(a) => axes.push(a.into_axis(ChartAxisKind::Category)),
                PlotAreaChildXml::ValAx(a) => axes.push(a.into_axis(ChartAxisKind::Value)),
                PlotAreaChildXml::Other => {}
            }
        }
    }

    Ok(ChartSpace {
        title: Dup::from(chart.title).into_value().and_then(TitleXml::text),
        auto_title_deleted: Dup::from(chart.auto_title_deleted)
            .into_value()
            .map(|b| b.value())
            .unwrap_or(false),
        plot_groups,
        axes,
        legend: Dup::from(chart.legend).into_value().map(Into::into),
    })
}

fn p_dir(p: &PlotTypeXml) -> String {
    Dup::from(p.bar_dir.clone())
        .into_value()
        .map(|v| v.val)
        .unwrap_or_else(|| "col".into())
}

fn bar_kind(dir: &str) -> PlotKind {
    PlotKind::Bar {
        horizontal: dir == "bar",
    }
}

// ── chartSpace skeleton ───────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
struct ChartSpaceXml {
    #[serde(rename = "chart", default)]
    chart: Vec<ChartXml>,
}

#[derive(Debug, Deserialize)]
struct ChartXml {
    #[serde(rename = "title", default)]
    title: Vec<TitleXml>,
    #[serde(rename = "autoTitleDeleted", default)]
    auto_title_deleted: Vec<CBoolXml>,
    #[serde(rename = "plotArea", default)]
    plot_area: Vec<PlotAreaXml>,
    #[serde(rename = "legend", default)]
    legend: Vec<LegendXml>,
}

#[derive(Debug, Deserialize)]
struct TitleXml {
    #[serde(rename = "tx", default)]
    tx: Vec<TitleTxXml>,
}

impl TitleXml {
    fn text(self) -> Option<DrawingTextBody> {
        Dup::from(self.tx)
            .into_value()
            .and_then(|t| Dup::from(t.rich).into_value())
            .map(Into::into)
    }
}

#[derive(Debug, Deserialize)]
struct TitleTxXml {
    #[serde(rename = "rich", default)]
    rich: Vec<TextBodyXml>,
}

/// `c:plotArea` — children in document order via `$value`, so combo charts
/// keep their group order and axes pair up as written.
#[derive(Debug, Deserialize)]
struct PlotAreaXml {
    #[serde(rename = "$value", default)]
    children: Vec<PlotAreaChildXml>,
}

#[derive(Debug, Deserialize)]
enum PlotAreaChildXml {
    #[serde(rename = "barChart")]
    BarChart(PlotTypeXml),
    #[serde(rename = "bar3DChart")]
    Bar3DChart(PlotTypeXml),
    #[serde(rename = "lineChart")]
    LineChart(PlotTypeXml),
    #[serde(rename = "line3DChart")]
    Line3DChart(PlotTypeXml),
    #[serde(rename = "pieChart")]
    PieChart(PlotTypeXml),
    #[serde(rename = "pie3DChart")]
    Pie3DChart(PlotTypeXml),
    #[serde(rename = "doughnutChart")]
    DoughnutChart(PlotTypeXml),
    #[serde(rename = "areaChart")]
    AreaChart(PlotTypeXml),
    #[serde(rename = "area3DChart")]
    Area3DChart(PlotTypeXml),
    #[serde(rename = "scatterChart")]
    ScatterChart(PlotTypeXml),
    #[serde(rename = "ofPieChart")]
    OfPieChart(PlotTypeXml),
    #[serde(rename = "radarChart")]
    RadarChart(PlotTypeXml),
    #[serde(rename = "surfaceChart")]
    SurfaceChart(PlotTypeXml),
    #[serde(rename = "surface3DChart")]
    Surface3DChart(PlotTypeXml),
    #[serde(rename = "bubbleChart")]
    BubbleChart(PlotTypeXml),
    #[serde(rename = "stockChart")]
    StockChart(PlotTypeXml),
    #[serde(rename = "catAx")]
    CatAx(AxisXml),
    #[serde(rename = "dateAx")]
    DateAx(AxisXml),
    #[serde(rename = "valAx")]
    ValAx(AxisXml),
    #[serde(other)]
    Other,
}

/// One plot-type group — the same fields serve every kind; serde ignores
/// the ones a given kind never writes.
#[derive(Debug, Deserialize)]
struct PlotTypeXml {
    #[serde(rename = "barDir", default)]
    bar_dir: Vec<ValAttrXml>,
    #[serde(rename = "grouping", default)]
    grouping: Vec<ValAttrXml>,
    #[serde(rename = "varyColors", default)]
    vary_colors: Vec<CBoolXml>,
    #[serde(rename = "ser", default)]
    ser: Vec<SeriesXml>,
    #[serde(rename = "gapWidth", default)]
    gap_width: Vec<U32AttrXml>,
    #[serde(rename = "overlap", default)]
    overlap: Vec<I32AttrXml>,
    #[serde(rename = "firstSliceAng", default)]
    first_slice_ang: Vec<U32AttrXml>,
    #[serde(rename = "holeSize", default)]
    hole_size: Vec<U32AttrXml>,
}

impl PlotTypeXml {
    fn into_group(self, kind: PlotKind) -> PlotGroup {
        let grouping = match Dup::from(self.grouping)
            .into_value()
            .map(|v| v.val)
            .as_deref()
        {
            Some("stacked") => ChartGrouping::Stacked,
            Some("percentStacked") => ChartGrouping::PercentStacked,
            Some("standard") => ChartGrouping::Standard,
            _ => ChartGrouping::Clustered,
        };
        let mut series: Vec<ChartSeries> = self.ser.into_iter().map(Into::into).collect();
        // §21.2.2.140 c:order governs plot order; the file may list series
        // in any order.
        series.sort_by_key(|s| s.order);
        PlotGroup {
            kind,
            series,
            gap_width: Dup::from(self.gap_width).into_value().map(|v| v.val),
            overlap: Dup::from(self.overlap).into_value().map(|v| v.val),
            grouping,
            vary_colors: Dup::from(self.vary_colors).into_value().map(|b| b.value()),
            first_slice_angle: Dup::from(self.first_slice_ang).into_value().map(|v| v.val),
            hole_size: Dup::from(self.hole_size).into_value().map(|v| v.val),
        }
    }
}

// ── series ────────────────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
struct SeriesXml {
    #[serde(rename = "idx", default)]
    idx: Vec<U32AttrXml>,
    #[serde(rename = "order", default)]
    order: Vec<U32AttrXml>,
    #[serde(rename = "tx", default)]
    tx: Vec<SeriesTxXml>,
    #[serde(rename = "spPr", default)]
    sp_pr: Vec<SpPrXml>,
    #[serde(rename = "marker", default)]
    marker: Vec<MarkerXml>,
    #[serde(rename = "dPt", default)]
    d_pt: Vec<DataPointXml>,
    #[serde(rename = "cat", default)]
    cat: Vec<DataSourceXml>,
    #[serde(rename = "val", default)]
    val: Vec<DataSourceXml>,
    #[serde(rename = "xVal", default)]
    x_val: Vec<DataSourceXml>,
    #[serde(rename = "yVal", default)]
    y_val: Vec<DataSourceXml>,
}

impl From<SeriesXml> for ChartSeries {
    fn from(x: SeriesXml) -> Self {
        let idx = Dup::from(x.idx).into_value().map(|v| v.val).unwrap_or(0);
        let cat = Dup::from(x.cat).into_value();
        let val = Dup::from(x.val).into_value();
        let x_val = Dup::from(x.x_val).into_value();
        let y_val = Dup::from(x.y_val).into_value();
        Self {
            idx,
            order: Dup::from(x.order)
                .into_value()
                .map(|v| v.val)
                .unwrap_or(idx),
            name: Dup::from(x.tx).into_value().and_then(SeriesTxXml::text),
            shape_properties: Dup::from(x.sp_pr).into_value().map(Into::into),
            categories: cat.map(DataSourceXml::strings).unwrap_or_default(),
            values: val
                .or(y_val)
                .map(DataSourceXml::numbers)
                .unwrap_or_default(),
            x_values: x_val.map(DataSourceXml::numbers).unwrap_or_default(),
            point_properties: x
                .d_pt
                .into_iter()
                .filter_map(|d| {
                    let idx = Dup::from(d.idx).into_value()?.val;
                    let props = Dup::from(d.sp_pr).into_value()?;
                    Some((idx, props.into()))
                })
                .collect(),
            marker: Dup::from(x.marker).into_value().and_then(MarkerXml::symbol),
        }
    }
}

#[derive(Debug, Deserialize)]
struct SeriesTxXml {
    #[serde(rename = "strRef", default)]
    str_ref: Vec<StrRefXml>,
    #[serde(rename = "v", default)]
    v: Vec<String>,
}

impl SeriesTxXml {
    fn text(self) -> Option<String> {
        if let Some(r) = Dup::from(self.str_ref).into_value() {
            return r.first_string();
        }
        Dup::from(self.v).into_value()
    }
}

#[derive(Debug, Deserialize)]
struct MarkerXml {
    #[serde(rename = "symbol", default)]
    symbol: Vec<ValAttrXml>,
}

impl MarkerXml {
    fn symbol(self) -> Option<ChartMarker> {
        Dup::from(self.symbol)
            .into_value()
            .map(|v| match v.val.as_str() {
                "none" => ChartMarker::None,
                "circle" | "auto" => ChartMarker::Circle,
                "square" => ChartMarker::Square,
                "diamond" => ChartMarker::Diamond,
                "triangle" => ChartMarker::Triangle,
                _ => ChartMarker::Other,
            })
    }
}

#[derive(Debug, Deserialize)]
struct DataPointXml {
    #[serde(rename = "idx", default)]
    idx: Vec<U32AttrXml>,
    #[serde(rename = "spPr", default)]
    sp_pr: Vec<SpPrXml>,
}

// ── data sources: refs with caches, or literals ───────────────────────────

#[derive(Debug, Deserialize)]
struct DataSourceXml {
    #[serde(rename = "strRef", default)]
    str_ref: Vec<StrRefXml>,
    #[serde(rename = "numRef", default)]
    num_ref: Vec<NumRefXml>,
    #[serde(rename = "strLit", default)]
    str_lit: Vec<CacheXml>,
    #[serde(rename = "numLit", default)]
    num_lit: Vec<CacheXml>,
}

impl DataSourceXml {
    fn cache(self) -> Option<CacheXml> {
        if let Some(r) = Dup::from(self.str_ref).into_value() {
            return Dup::from(r.str_cache).into_value();
        }
        if let Some(r) = Dup::from(self.num_ref).into_value() {
            return Dup::from(r.num_cache).into_value();
        }
        Dup::from(self.str_lit)
            .into_value()
            .or_else(|| Dup::from(self.num_lit).into_value())
    }

    /// Category labels: one slot per point, `None` for a blank cell. A
    /// numeric cache reads as its formatted-enough decimal text.
    fn strings(self) -> Vec<Option<String>> {
        self.cache().map(CacheXml::strings).unwrap_or_default()
    }

    fn numbers(self) -> Vec<Option<f64>> {
        self.cache().map(CacheXml::numbers).unwrap_or_default()
    }
}

#[derive(Debug, Deserialize)]
struct StrRefXml {
    #[serde(rename = "strCache", default)]
    str_cache: Vec<CacheXml>,
}

impl StrRefXml {
    fn first_string(self) -> Option<String> {
        Dup::from(self.str_cache)
            .into_value()
            .and_then(|c| c.strings().into_iter().flatten().next())
    }
}

#[derive(Debug, Deserialize)]
struct NumRefXml {
    #[serde(rename = "numCache", default)]
    num_cache: Vec<CacheXml>,
}

/// `c:strCache`/`c:numCache`/`c:strLit`/`c:numLit` — all the same shape:
/// a point count and sparse indexed points.
#[derive(Debug, Deserialize)]
struct CacheXml {
    #[serde(rename = "ptCount", default)]
    pt_count: Vec<U32AttrXml>,
    #[serde(rename = "pt", default)]
    pt: Vec<PointXml>,
}

#[derive(Debug, Deserialize)]
struct PointXml {
    #[serde(rename = "@idx")]
    idx: u32,
    #[serde(rename = "v", default)]
    v: Vec<String>,
}

impl CacheXml {
    fn len(&self) -> usize {
        let declared = Dup::from(self.pt_count.clone())
            .into_value()
            .map(|v| v.val as usize)
            .unwrap_or(0);
        let max_idx = self
            .pt
            .iter()
            .map(|p| p.idx as usize + 1)
            .max()
            .unwrap_or(0);
        declared.max(max_idx)
    }

    fn strings(self) -> Vec<Option<String>> {
        let mut out = vec![None; self.len()];
        for p in self.pt {
            if let Some(slot) = out.get_mut(p.idx as usize) {
                *slot = Dup::from(p.v).into_value();
            }
        }
        out
    }

    fn numbers(self) -> Vec<Option<f64>> {
        let mut out = vec![None; self.len()];
        for p in self.pt {
            if let Some(slot) = out.get_mut(p.idx as usize) {
                *slot = Dup::from(p.v)
                    .into_value()
                    .and_then(|v| v.trim().parse::<f64>().ok())
                    .filter(|v| v.is_finite());
            }
        }
        out
    }
}

// ── axes and legend ───────────────────────────────────────────────────────

#[derive(Debug, Deserialize)]
struct AxisXml {
    #[serde(rename = "delete", default)]
    delete: Vec<CBoolXml>,
    #[serde(rename = "axPos", default)]
    ax_pos: Vec<ValAttrXml>,
    #[serde(rename = "scaling", default)]
    scaling: Vec<ScalingXml>,
    #[serde(rename = "majorGridlines", default)]
    major_gridlines: Vec<IgnoredXml>,
    #[serde(rename = "title", default)]
    title: Vec<TitleXml>,
    #[serde(rename = "tickLblPos", default)]
    tick_lbl_pos: Vec<ValAttrXml>,
}

impl AxisXml {
    fn into_axis(self, kind: ChartAxisKind) -> ChartAxis {
        let scaling = Dup::from(self.scaling).into_value();
        ChartAxis {
            kind,
            deleted: Dup::from(self.delete)
                .into_value()
                .map(|b| b.value())
                .unwrap_or(false),
            position: match Dup::from(self.ax_pos)
                .into_value()
                .map(|v| v.val)
                .as_deref()
            {
                Some("l") => ChartAxisPosition::Left,
                Some("t") => ChartAxisPosition::Top,
                Some("r") => ChartAxisPosition::Right,
                _ => ChartAxisPosition::Bottom,
            },
            reversed: scaling
                .as_ref()
                .and_then(|s| Dup::from(s.orientation.clone()).into_value())
                .map(|v| v.val == "maxMin")
                .unwrap_or(false),
            min: scaling
                .as_ref()
                .and_then(|s| Dup::from(s.min.clone()).into_value())
                .map(|v| v.val),
            max: scaling
                .as_ref()
                .and_then(|s| Dup::from(s.max.clone()).into_value())
                .map(|v| v.val),
            major_gridlines: !self.major_gridlines.is_empty(),
            title: Dup::from(self.title).into_value().and_then(TitleXml::text),
            labels_hidden: Dup::from(self.tick_lbl_pos)
                .into_value()
                .map(|v| v.val == "none")
                .unwrap_or(false),
        }
    }
}

#[derive(Debug, Deserialize)]
struct ScalingXml {
    #[serde(rename = "orientation", default)]
    orientation: Vec<ValAttrXml>,
    #[serde(rename = "min", default)]
    min: Vec<F64AttrXml>,
    #[serde(rename = "max", default)]
    max: Vec<F64AttrXml>,
}

#[derive(Debug, Deserialize)]
struct LegendXml {
    #[serde(rename = "legendPos", default)]
    legend_pos: Vec<ValAttrXml>,
    #[serde(rename = "overlay", default)]
    overlay: Vec<CBoolXml>,
}

impl From<LegendXml> for Legend {
    fn from(x: LegendXml) -> Self {
        Self {
            position: match Dup::from(x.legend_pos)
                .into_value()
                .map(|v| v.val)
                .as_deref()
            {
                Some("l") => LegendPosition::Left,
                Some("t") => LegendPosition::Top,
                Some("b") => LegendPosition::Bottom,
                Some("tr") => LegendPosition::TopRight,
                _ => LegendPosition::Right,
            },
            overlay: Dup::from(x.overlay)
                .into_value()
                .map(|b| b.value())
                .unwrap_or(false),
        }
    }
}

// ── attribute atoms ───────────────────────────────────────────────────────

/// §21.2.2 CT_Boolean: `<c:x val="0"/>`. Present without `@val` is **true**.
#[derive(Clone, Debug, Deserialize)]
struct CBoolXml {
    #[serde(rename = "@val", default)]
    val: Option<crate::docx::parse::primitives::AttrBool>,
}

impl CBoolXml {
    fn value(&self) -> bool {
        self.val.map(|b| b.0).unwrap_or(true)
    }
}

#[derive(Clone, Debug, Deserialize)]
struct ValAttrXml {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Clone, Copy, Debug, Deserialize)]
struct U32AttrXml {
    #[serde(rename = "@val")]
    val: u32,
}

#[derive(Clone, Copy, Debug, Deserialize)]
struct I32AttrXml {
    #[serde(rename = "@val")]
    val: i32,
}

#[derive(Clone, Copy, Debug, Deserialize)]
struct F64AttrXml {
    #[serde(rename = "@val")]
    val: f64,
}

/// An element whose presence is the signal and whose content is ignored.
#[derive(Debug, Default, Deserialize)]
struct IgnoredXml {}

#[cfg(test)]
mod tests {
    use super::*;

    const BAR: &[u8] = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>S!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>2.5</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="150"/><c:overlap val="-27"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>
        <c:scaling><c:orientation val="minMax"/><c:max val="10"/></c:scaling><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>"#;

    #[test]
    fn bar_chart_parses_series_caches_and_axes() {
        let c = parse_chart(BAR).unwrap();
        assert_eq!(c.plot_groups.len(), 1);
        let g = &c.plot_groups[0];
        assert_eq!(g.kind, PlotKind::Bar { horizontal: false });
        assert_eq!(g.gap_width, Some(150));
        assert_eq!(g.overlap, Some(-27));
        let s = &g.series[0];
        assert_eq!(s.name.as_deref(), Some("North"));
        assert_eq!(
            s.categories,
            vec![Some("Q1".into()), Some("Q2".into()), Some("Q3".into())]
        );
        assert_eq!(
            s.values,
            vec![Some(4.0), None, Some(2.5)],
            "blank cell = None"
        );
        assert_eq!(c.axes.len(), 2);
        assert!(c.axes[1].major_gridlines);
        assert_eq!(c.axes[1].max, Some(10.0));
        assert_eq!(c.legend.as_ref().unwrap().position, LegendPosition::Bottom);
        let title = c.title.unwrap();
        assert_eq!(title.paragraphs.len(), 1);
    }

    /// CT_Boolean's trap: `<c:delete/>` with no `val` is TRUE; `val="0"`
    /// false; the absent element takes the property default.
    #[test]
    fn boolean_elements_default_to_true_when_bare() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea>
    <c:pieChart><c:varyColors/><c:ser><c:idx val="0"/><c:order val="0"/></c:ser></c:pieChart>
    <c:catAx><c:delete/></c:catAx>
  </c:plotArea></c:chart>
</c:chartSpace>"#;
        let c = parse_chart(xml).unwrap();
        assert_eq!(c.plot_groups[0].vary_colors, Some(true));
        assert!(c.axes[0].deleted);
    }

    /// 3D kinds flatten to their 2D twins; unsupported kinds carry their
    /// element name for the renderer's warn.
    #[test]
    fn threedee_flattens_and_radar_declines() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea>
    <c:bar3DChart><c:barDir val="bar"/><c:ser><c:idx val="0"/><c:order val="0"/></c:ser></c:bar3DChart>
    <c:radarChart><c:ser><c:idx val="1"/><c:order val="1"/></c:ser></c:radarChart>
  </c:plotArea></c:chart>
</c:chartSpace>"#;
        let c = parse_chart(xml).unwrap();
        assert_eq!(c.plot_groups[0].kind, PlotKind::Bar { horizontal: true });
        assert_eq!(
            c.plot_groups[1].kind,
            PlotKind::Unsupported("radarChart".into())
        );
    }

    /// Scatter series carry x and y caches.
    #[test]
    fn scatter_reads_both_value_caches() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:scatterChart>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:xVal><c:numRef><c:numCache><c:ptCount val="2"/>
        <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>
      <c:yVal><c:numRef><c:numCache><c:ptCount val="2"/>
        <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:yVal>
    </c:ser>
  </c:scatterChart></c:plotArea></c:chart>
</c:chartSpace>"#;
        let c = parse_chart(xml).unwrap();
        let s = &c.plot_groups[0].series[0];
        assert_eq!(s.x_values, vec![Some(1.0), Some(2.0)]);
        assert_eq!(s.values, vec![Some(10.0), Some(20.0)]);
    }
}
