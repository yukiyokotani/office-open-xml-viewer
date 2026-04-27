use wasm_bindgen::prelude::*;
use serde::Serialize;
use std::collections::HashMap;
use std::io::{Cursor, Read};
use base64::{engine::general_purpose::STANDARD as B64, Engine as _};

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Workbook {
    pub sheets: Vec<SheetMeta>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetMeta {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MergeCell {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Worksheet {
    pub name: String,
    pub rows: Vec<Row>,
    pub col_widths: HashMap<u32, f64>,
    pub row_heights: HashMap<u32, f64>,
    pub default_col_width: f64,
    pub default_row_height: f64,
    pub merge_cells: Vec<MergeCell>,
    pub freeze_rows: u32,
    pub freeze_cols: u32,
    pub conditional_formats: Vec<ConditionalFormat>,
    pub images: Vec<ImageAnchor>,
    pub charts: Vec<ChartAnchor>,
    /// Grouped shapes from `<xdr:grpSp>` inside a twoCellAnchor (ECMA-376
    /// §20.5.2.17). Each anchor flattens its shape tree to a list of leaf
    /// shapes with normalized geometry for the renderer.
    pub shape_groups: Vec<ShapeAnchor>,
    /// Whether to display zero values in cells (ECMA-376 §18.3.1.94)
    pub show_zeros: bool,
    /// Whether to draw default grid lines on this sheet. Mirrors the "View →
    /// Gridlines" checkbox in Excel; parsed from `<sheetView showGridLines>`
    /// (ECMA-376 §18.3.1.83). Defaults to true.
    pub show_gridlines: bool,
    /// Tab color for the sheet tab (ECMA-376 §18.3.1.79)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// AutoFilter range (ECMA-376 §18.3.1.2)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<CellRange>,
    /// Hyperlinks in this worksheet (ECMA-376 §18.3.1.47)
    pub hyperlinks: Vec<Hyperlink>,
    /// Cell refs (A1-style) that have an associated <comment> in xl/commentsN.xml.
    /// Excel shows a small red triangle in the top-right corner of each.
    pub comment_refs: Vec<String>,
    /// Defined names in scope for this sheet. Includes workbook-global names and
    /// any names whose `localSheetId` matches this sheet's position in the
    /// workbook. Used by conditional-formatting `expression` rules that call
    /// named ranges like `task_start`, `today`, etc. (ECMA-376 §18.2.5).
    pub defined_names: Vec<DefinedName>,
    /// Excel Tables defined for this sheet (ECMA-376 §18.5). Rendered with a
    /// built-in table style (bold header, banded rows, etc.) on top of the
    /// cells' own styles.
    pub tables: Vec<TableInfo>,
    /// Slicers anchored to the sheet's drawing (Office 2010+ extension —
    /// `http://schemas.microsoft.com/office/drawing/2010/slicer` inside
    /// `<mc:AlternateContent>/<mc:Choice>`). Each slicer resolves its cache
    /// and referenced pivotCacheDefinition so the renderer can draw a static
    /// button list with the saved selection state.
    pub slicers: Vec<SlicerAnchor>,
    /// Sparkline groups defined in the worksheet's `<extLst>` (Office 2010
    /// extension `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`,
    /// element `<x14:sparklineGroup>`, ECMA-376 §18.2 / Part 4).
    pub sparkline_groups: Vec<SparklineGroup>,
}

/// Single sparkline group (`<x14:sparklineGroup>`). Holds the shared formatting
/// for every individual sparkline cell that belongs to the group.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SparklineGroup {
    /// `line` (default) | `column` | `stem` (win/loss).
    pub kind: SparklineType,
    /// Show a marker dot at every data point (line type only).
    pub markers: bool,
    /// Highlight high / low / first / last / negative points.
    pub high: bool,
    pub low: bool,
    pub first: bool,
    pub last: bool,
    pub negative: bool,
    /// Show the horizontal axis line when data crosses zero.
    pub display_x_axis: bool,
    /// `gap` (default) | `zero` | `span` — how empty cells in the source
    /// range are treated. We only honor `gap` (default) at render time.
    pub display_empty_cells_as: String,
    /// Per-axis-bound type: `individual` (default) / `group` / `custom`.
    pub min_axis_type: String,
    pub max_axis_type: String,
    /// Used when *AxisType=`custom`. f64::NAN otherwise.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manual_max: Option<f64>,
    /// Stroke weight in pt (line type). ECMA-376 default 0.75.
    pub line_weight: f64,
    /// Resolved RGB hex strings (e.g. `#4472C4`) — theme + tint flattened
    /// at parse time so the renderer never sees a theme index. `None` means
    /// the property was not specified and the renderer should fall back.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_series: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_negative: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_axis: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_markers: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_first: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_last: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_high: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color_low: Option<String>,
    /// Individual sparklines (one per destination cell).
    pub sparklines: Vec<Sparkline>,
}

#[derive(Debug, Serialize, Clone, Copy)]
#[serde(rename_all = "camelCase")]
pub enum SparklineType {
    Line,
    Column,
    Stem,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Sparkline {
    /// 1-based row of the cell that displays this sparkline (`<xm:sqref>`).
    pub row: u32,
    /// 1-based column of the cell.
    pub col: u32,
    /// Numeric values resolved from the `<xm:f>` data range. `None` for
    /// empty / non-numeric / out-of-bounds cells.
    pub values: Vec<Option<f64>>,
}

/// Excel Table metadata (ECMA-376 §18.5 `<table>`). The renderer overlays a
/// built-in style on top of the cell styles inside `range`.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableInfo {
    /// Inclusive table area including the header row.
    pub range: CellRange,
    /// Built-in style name like "TableStyleLight18" (ECMA-376 §18.5.1.4).
    pub style_name: String,
    /// Number of header rows (default 1).
    pub header_row_count: u32,
    /// Number of totals rows at the bottom (default 0).
    pub totals_row_count: u32,
    /// `<tableStyleInfo showRowStripes>` — banded rows in the data region.
    pub show_row_stripes: bool,
    /// `<tableStyleInfo showColumnStripes>`.
    pub show_column_stripes: bool,
    /// `<tableStyleInfo showFirstColumn>`.
    pub show_first_column: bool,
    /// `<tableStyleInfo showLastColumn>`.
    pub show_last_column: bool,
    /// Accent color resolved from the built-in style name against this file's
    /// theme accents (e.g. `TableStyleLight18` → accent3 of theme1.xml). Used
    /// by the renderer to draw banding, header background, and rules.
    pub accent_color: String,
    /// Dxf index for the `wholeTable` element of a custom `<tableStyle>`
    /// (ECMA-376 §18.8.40). When set, its border/fill apply to every cell
    /// of the table as a base layer. Built-in style names use the renderer's
    /// accent-based fallback, not this field.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub whole_table_dxf: Option<u32>,
    /// Dxf index for the `headerRow` element of a custom `<tableStyle>`.
    /// Provides the header background fill, font color/weight, and any
    /// vertical separator borders shown between header cells.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub header_row_dxf: Option<u32>,
}

/// Workbook- or sheet-scoped defined name (ECMA-376 §18.2.5 `definedName`).
/// `formula` is the raw formula text (typically a cell/range reference, e.g.
/// `ProjectSchedule!$E1`). Relative references inside are shifted relative to
/// A1 when substituted into a formula.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
}

// ─── Chart types ────────────────────────────────────────────────────────────

/// A data series inside a chart.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartSeries {
    /// Display name of the series.
    pub name: String,
    /// Chart type for this series ("bar"|"line"|"area"|"pie"|"radar"|"scatter").
    /// Allows mixed charts (e.g. bar + line sharing the same axes).
    pub series_type: String,
    /// Category labels (X-axis for most charts).
    pub categories: Vec<String>,
    /// Numeric values; `None` = missing data point.
    pub values: Vec<Option<f64>>,
    /// Explicit fill color hex. When the series has a `<c:spPr>` fill we resolve
    /// it at parse time; otherwise we fall back to `theme.accent[idx+1]` (the
    /// default Excel palette, keyed by `<c:idx>` per ECMA-376 §21.2.2.27) so the
    /// renderer doesn't need theme access. None only when neither applies.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Whether to draw data-point markers on line/scatter series. Resolved at
    /// parse time from `<c:ser><c:marker><c:symbol val>` (ECMA-376 §21.2.2.32)
    /// falling back to the chart-type-level `<c:lineChart><c:marker val>`
    /// (§21.2.2.33). Absent markers default to hidden for line charts.
    pub show_marker: bool,
    /// `<c:val>/<c:numRef>/<c:numCache>/<c:formatCode>` — Excel number format
    /// code applied to the series' numeric values (e.g. `"¥"#,##0`). When the
    /// chart-level `<c:dLbls><c:numFmt>` is absent this governs how data
    /// labels are formatted per ECMA-376 §21.2.2.37.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_format_code: Option<String>,
    /// `<c:ser><c:order val>` — display order of this series (stacking
    /// ordering + legend ordering) per ECMA-376 §21.2.2.28. Lower `order`
    /// renders first/below; Excel's legend for horizontal bar charts reverses
    /// this so low-order series end up at the bottom of the plot.
    pub order: usize,
    /// `<c:marker><c:symbol val>` (ECMA-376 §21.2.2.32). One of
    /// "circle"|"square"|"diamond"|"triangle"|"x"|"plus"|"star"|"dot"|
    /// "dash"|"picture"|"none". Absent = renderer-default circle.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    /// `<c:marker><c:size val>` (ECMA-376 §21.2.2.34) — marker side length
    /// in points. Default is renderer-defined (~5 pt).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<u32>,
    /// `<c:marker><c:spPr><a:solidFill>` resolved hex (no `#`). Marker fill
    /// independent of series stroke color.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    /// `<c:marker><c:spPr><a:ln><a:solidFill>` resolved hex (no `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
    /// Per-data-point overrides (`<c:dPt idx>`, ECMA-376 §21.2.2.39). Each
    /// entry overrides marker / fill for a single point in the series.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub data_point_overrides: Vec<DataPointOverride>,
    /// Per-data-point custom labels (`<c:dLbl idx>` inside `<c:dLbls>`,
    /// ECMA-376 §21.2.2.45). Custom rich text — including
    /// `<a:fld type="CELLRANGE">` field references — is resolved to a plain
    /// string at parse time.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub data_label_overrides: Vec<DataLabelOverride>,
    /// Series-level `<c:dLbls>` defaults (showVal / showSerName / etc.). When
    /// neither this nor `data_label_overrides` is present, series labels are
    /// suppressed even when the chart-level `data_label_position` is set.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub series_data_labels: Option<SeriesDataLabels>,
    /// Error bars (`<c:errBars>`, ECMA-376 §21.2.2.20). Up to two per series
    /// (one for X, one for Y). `cust` / `fixedVal` / `stdErr` / `stdDev` /
    /// `percentage` are all resolved to absolute plus/minus arrays at parse
    /// time so the renderer just draws lines.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub err_bars: Vec<ErrBars>,
}

/// Per-point override pulled from `<c:dPt idx="N">` siblings of a series
/// (ECMA-376 §21.2.2.39). Any unset field falls back to the series-level
/// value.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataPointOverride {
    pub idx: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_symbol: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_size: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub marker_line: Option<String>,
}

/// Custom data label for one point (`<c:dLbl idx="N">`).
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DataLabelOverride {
    pub idx: u32,
    /// Resolved text. Empty when the label is intentionally blank
    /// (e.g. an empty `<a:p>`); the renderer should still skip drawing
    /// for that idx because Excel deletes it.
    pub text: String,
    /// `<c:dLblPos val>` — "l"|"r"|"t"|"b"|"ctr"|"outEnd"|"bestFit". When
    /// `None` the series-level / chart-level position applies.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    /// Resolved font color hex (no `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    /// Font size in OOXML hundredths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
    /// `<a:defRPr b="1">` inside the per-idx rich text. None = inherit.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
}

/// Series-level `<c:dLbls>` defaults applied to every point that doesn't
/// have its own `<c:dLbl>` override.
#[derive(Debug, Serialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub struct SeriesDataLabels {
    pub show_val: bool,
    pub show_cat_name: bool,
    pub show_ser_name: bool,
    pub show_percent: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format_code: Option<String>,
    /// `<c:dLbls><c:txPr>...defRPr@b>` series-level bold default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_bold: Option<bool>,
    /// `<c:dLbls><c:txPr>...defRPr@sz>` series-level font size in OOXML
    /// hundredths of a point (e.g. 1200 = 12 pt).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size_hpt: Option<i32>,
}

/// Error bars (`<c:errBars>`, ECMA-376 §21.2.2.20).
///
/// `errValType="cust"` — values come from `<c:plus>/<c:minus>/<c:numRef>`,
/// possibly cross-sheet. We resolve those at parse time into arrays. For
/// `fixedVal` / `stdErr` / `stdDev` / `percentage` the parser computes the
/// absolute per-point delta from the series values so the renderer just
/// draws line segments without needing the raw type info.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ErrBars {
    /// "x" | "y".
    pub dir: String,
    /// "plus" | "minus" | "both" — drives whether plus / minus is rendered.
    pub bar_type: String,
    /// Absolute positive deltas per point. `None` = no bar in that
    /// direction at that index.
    pub plus: Vec<Option<f64>>,
    /// Absolute negative deltas per point.
    pub minus: Vec<Option<f64>>,
    /// `<c:noEndCap val>` — when true the renderer skips the perpendicular
    /// cap line at the bar tip (ECMA-376 §21.2.2.21).
    pub no_end_cap: bool,
    /// Resolved stroke hex (no `#`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Stroke width in EMU (`<a:ln w>`). 12700 EMU = 1 pt.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_width_emu: Option<u32>,
    /// `<a:prstDash val>` — "solid"|"dash"|"dot"|"dashDot"|...
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
}

/// Parsed chart data extracted from `xl/charts/chartN.xml`.
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ChartData {
    /// Primary chart type: "bar"|"line"|"area"|"pie"|"doughnut"|"radar"|"scatter"
    pub chart_type: String,
    /// Bar direction: "col" (vertical) | "row" (horizontal). Only relevant for bar charts.
    pub bar_dir: String,
    /// Grouping mode: "clustered"|"stacked"|"standard"|"percentStacked"
    pub grouping: String,
    /// Optional chart title.
    pub title: Option<String>,
    /// Shared category list (from first series that has categories).
    pub categories: Vec<String>,
    /// All series across all chart-type elements in plotArea.
    pub series: Vec<ChartSeries>,
    /// Whether data labels are enabled (c:dLbls with showVal or showPercent).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub show_data_labels: bool,
    /// Category axis title (c:catAx/c:title).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_title: Option<String>,
    /// Value axis title (c:valAx/c:title).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_title: Option<String>,
    /// `c:valAx/c:numFmt@formatCode` — custom number format for the value axis
    /// tick labels (e.g. `"$"#,##0`). When unset, tick labels use a plain
    /// numeric format. `sourceLinked="1"` is treated as a non-override (i.e.
    /// the axis inherits the data's format code); we still capture it so the
    /// renderer can honor it.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_format_code: Option<String>,
    /// True when `<c:legend>` is present in the chart; false means no legend.
    pub show_legend: bool,
    /// `<c:legend><c:legendPos val>` — "r" (default) | "l" | "t" | "b" | "tr".
    /// None = default ("r").
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend_pos: Option<String>,
    /// Chart title font size in OOXML hundredths of a point (e.g. 1400 = 14pt).
    /// Taken from the first `defRPr@sz` or `rPr@sz` inside `c:title`. None =
    /// not specified; renderer falls back to a proportional default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_size_hpt: Option<i32>,
    /// Chart title font color as a hex string without '#'. Taken from the
    /// first `a:solidFill/a:srgbClr@val` inside `c:title`. None = default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_color: Option<String>,
    /// Chart title font family (ECMA-376 DrawingML §20.1.4.2.24 `a:latin@typeface`).
    /// Taken from the first `a:latin` element inside `c:title`. None = default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_face: Option<String>,
    /// Category axis tick-label font size in hundredths of a point
    /// (ECMA-376 §21.2.2.17 `c:txPr/a:defRPr@sz`). None = not specified.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_size_hpt: Option<i32>,
    /// Value axis tick-label font size in hundredths of a point.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_font_size_hpt: Option<i32>,
    /// ECMA-376 §21.2.2.40 — `<c:catAx><c:delete val="1"/>` hides the category
    /// axis (labels, ticks, and axis line).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub cat_axis_hidden: bool,
    /// ECMA-376 §21.2.2.40 — `<c:valAx><c:delete val="1"/>` hides the value
    /// axis (labels, ticks, and axis line).
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub val_axis_hidden: bool,
    /// Outer `<c:chartSpace><c:spPr>` fill resolution (ECMA-376 §21.2.2.5).
    /// `Some(hex)` for `<a:solidFill>`, `None` for `<a:noFill>` or when spPr is
    /// absent *and* no explicit fill is declared. An absent `chart_bg` on the
    /// JSON side (serde skips None) tells the renderer "no outer frame" —
    /// i.e. transparent, so the underlying cell panel shows through.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub chart_bg: Option<String>,
    /// True when the parser saw a `<c:chartSpace><c:spPr>` element at all.
    /// Lets the renderer distinguish "spec explicitly said noFill" (present
    /// but `chart_bg` is None) from "no spPr — use the default opaque white"
    /// (absent).
    pub has_chart_sp_pr: bool,
    /// `<c:legend><c:layout><c:manualLayout>` absolute placement (ECMA-376
    /// §21.2.2.31). All four fractions are relative to the chart space.
    /// None = use the default side-based layout from `legend_pos`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub legend_manual_layout: Option<LegendManualLayout>,
    /// `<c:barChart><c:gapWidth val>` — space between category groups as a
    /// percentage of bar width (ECMA-376 §21.2.2.13). Default per spec is 150.
    /// None = renderer default.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar_gap_width: Option<i32>,
    /// `<c:barChart><c:overlap val>` — overlap/gap between bars in a cluster as
    /// a signed percentage (ECMA-376 §21.2.2.25). Positive = bars overlap,
    /// negative = gap between bars, 0 = flush. Range [-100, 100].
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bar_overlap: Option<i32>,
    /// `<c:dLbls><c:dLblPos val>` (ECMA-376 §21.2.2.16). One of
    /// "ctr"|"inBase"|"inEnd"|"outEnd"|"l"|"r"|"t"|"b"|"bestFit" etc.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_position: Option<String>,
    /// Hex (no `#`) resolved from `<c:dLbls><c:txPr>` text fill. Used for data
    /// label text color (e.g. "FFFFFF" when labels sit inside filled bars).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_font_color: Option<String>,
    /// `<c:dLbls><c:numFmt@formatCode>` — optional chart-level override for
    /// data label number format (ECMA-376 §21.2.2.35). Takes precedence over
    /// per-series `val_format_code` at render time.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_label_format_code: Option<String>,
    /// `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr@b>` — bold flag for
    /// the chart title. None = inherit (treat as not bold).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_font_bold: Option<bool>,
    /// `<c:catAx><c:txPr>...defRPr@b>` — bold flag for X-axis tick labels.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_font_bold: Option<bool>,
    /// `<c:valAx><c:txPr>...defRPr@b>` — bold flag for Y-axis tick labels.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_font_bold: Option<bool>,
    /// `<c:catAx><c:majorTickMark val>` (ECMA-376 §21.2.2.49) — one of
    /// `none` / `out` / `in` / `cross`. Default `out`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_major_tick_mark: Option<String>,
    /// `<c:catAx><c:minorTickMark val>` — same vocabulary. Default `none`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_minor_tick_mark: Option<String>,
    /// `<c:valAx><c:majorTickMark val>` and `<c:minorTickMark val>`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_major_tick_mark: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_minor_tick_mark: Option<String>,
    /// `<c:catAx><c:spPr><a:ln>` resolved color (hex without `#`) and
    /// width in EMU. Renders the X-axis line itself; default light gray
    /// at 1 px when absent.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_line_width_emu: Option<u32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_line_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_line_width_emu: Option<u32>,
    /// `<c:catAx><c:crosses val>` — where the X axis sits along the Y axis
    /// (`autoZero` = at value 0, `min` = at the data min, `max` = at the
    /// data max). Default `autoZero`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses: Option<String>,
    /// `<c:catAx><c:crossesAt val>` — explicit numeric crossing point;
    /// takes precedence over `cat_axis_crosses`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_crosses_at: Option<f64>,
    /// `<c:valAx><c:crosses val>` and `<c:crossesAt val>` mirroring the
    /// catAx fields above.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_crosses_at: Option<f64>,
    /// `<c:catAx><c:numFmt@formatCode>` (or scatter's X-axis valAx). Used
    /// to format the bottom-axis tick labels, e.g. `"m/d/yyyy"`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_format_code: Option<String>,
    /// `<c:catAx><c:scaling><c:min/max>` — only meaningful for scatter
    /// charts (where the X axis is numeric). None = derive from data.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cat_axis_max: Option<f64>,
    /// `<c:valAx><c:scaling><c:min/max>` — explicit Y-axis range. None =
    /// derive from data.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_min: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub val_axis_max: Option<f64>,
    /// `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27) absolute
    /// placement for the chart title. When present, overrides the renderer's
    /// auto-positioning; absent = current default layout.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title_manual_layout: Option<ManualLayout>,
    /// `<c:plotArea><c:layout><c:manualLayout>` absolute placement for the
    /// plot area, with `layoutTarget=inner` (default) describing the inner
    /// plot rect (no axes / labels) or `outer` describing the outer rect
    /// (axes + labels included).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub plot_area_manual_layout: Option<ManualLayout>,
}

/// Generic `<c:manualLayout>` block (used for title, plotArea, legend).
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct ManualLayout {
    /// "edge" (fraction is from top-left of chart) | "factor" (fraction
    /// from default position). ECMA-376 §21.2.2.32 ST_LayoutMode.
    pub x_mode: String,
    pub y_mode: String,
    /// "inner" (excludes axes) | "outer" (includes axes). Only meaningful
    /// for plotArea; harmless for title.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub layout_target: Option<String>,
    pub x: f64,
    pub y: f64,
    /// Width / height fractions. `None` = let the renderer auto-fit
    /// (corresponds to ECMA-376 only-(x,y) layouts where the size keeps its
    /// auto value).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub w: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub h: Option<f64>,
}

/// `<c:manualLayout>` coordinates for a legend (ECMA-376 §21.2.2.31). Fractions
/// are of the chart space's width/height. `xMode`/`yMode` select between "edge"
/// (fraction from top-left) and "factor" (fraction from the default position).
#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct LegendManualLayout {
    pub x_mode: String,
    pub y_mode: String,
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

/// A chart anchored to a rectangular range of cells (ECMA-376 §20.5 twoCellAnchor).
/// Offsets are EMU (914400 EMU = 1 inch, 9525 EMU = 1 px @ 96 DPI).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    pub chart: ChartData,
}

/// A grouped-shape anchor (ECMA-376 §20.5.2.17, `<xdr:grpSp>` inside a
/// `<xdr:twoCellAnchor>`). Leaf shape elements (`<xdr:sp>`) from any nesting
/// level are flattened into `shapes` with normalized coordinates so the
/// renderer only needs to scale to the anchor rect.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    pub shapes: Vec<ShapeInfo>,
}

/// A leaf shape extracted from a grpSp/sp tree. Position/size are normalized
/// to [0,1] relative to the top-level grpSp extent (which itself maps to the
/// anchor rect).
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeInfo {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    /// Rotation in degrees (clockwise). DrawingML `a:xfrm/@rot` is in 60000ths
    /// of a degree; the parser converts to degrees here.
    pub rot: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stroke_color: Option<String>,
    /// Stroke width in EMU (914400 = 1 inch). 0 = no stroke.
    pub stroke_width: i64,
    pub geom: ShapeGeom,
}

#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum ShapeGeom {
    /// Preset geometry (rect, ellipse, roundRect, triangle, etc.).
    /// ECMA-376 §20.1.9.18 `a:prstGeom/@prst`.
    Preset { name: String },
    /// Freeform path geometry (ECMA-376 §20.1.9.2 `a:custGeom`).
    Custom { paths: Vec<PathInfo> },
    /// Bitmap image leaf inside a `<xdr:grpSp>` tree (ECMA-376 §20.5.2.17).
    /// `data_url` is a `data:<mime>;base64,…` URL produced from the drawing's
    /// relationship target (png/jpg/gif/…).
    #[serde(rename_all = "camelCase")]
    Image { data_url: String },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PathInfo {
    /// Path's own coordinate system width.
    pub w: f64,
    /// Path's own coordinate system height.
    pub h: f64,
    pub commands: Vec<PathCmd>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "op", rename_all = "camelCase")]
pub enum PathCmd {
    MoveTo { x: f64, y: f64 },
    LineTo { x: f64, y: f64 },
    CubicBezTo { x1: f64, y1: f64, x2: f64, y2: f64, x3: f64, y3: f64 },
    QuadBezTo { x1: f64, y1: f64, x2: f64, y2: f64 },
    /// ECMA-376 §20.1.9.3 `a:arcTo`. `stAng`/`swAng` are in 60000ths of a
    /// degree. The start point is the current pen position; the ellipse
    /// center is derived so the pen lies on the ellipse at `stAng`.
    ArcTo { wr: f64, hr: f64, st_ang: f64, sw_ang: f64 },
    Close,
}

/// Slicer anchor — a button bank that filters a connected pivot table or
/// Excel Table. Office stores slicers in a 2010 extension namespace
/// (`sle:slicer`) wrapped in `<mc:AlternateContent>`, with the cache data in
/// `xl/slicerCaches/*.xml` and the underlying item list in the linked
/// pivotCache's `sharedItems`.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// Slicer header text. Typically the `caption` of the slicer definition
    /// (`xl/slicers/slicerN.xml`), falling back to the drawing `cNvPr` name.
    pub caption: String,
    /// One row per cache item in display order. Items flagged "selected" are
    /// the ones currently active in the filter; non-selected items are drawn
    /// with the ghost style. When the cache selection state is unavailable,
    /// all items are emitted as selected (Excel's default).
    pub items: Vec<SlicerItem>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlicerItem {
    pub name: String,
    pub selected: bool,
}

/// An image anchored to a rectangular range of cells
/// (ECMA-376 §20.5, `<xdr:twoCellAnchor>`). Offsets are EMU (English
/// Metric Unit): 914400 EMU = 1 inch, and 9525 EMU = 1 pixel at 96 DPI.
#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImageAnchor {
    pub from_col: u32,
    pub from_col_off: i64,
    pub from_row: u32,
    pub from_row_off: i64,
    pub to_col: u32,
    pub to_col_off: i64,
    pub to_row: u32,
    pub to_row_off: i64,
    /// Data URL: "data:image/png;base64,..."
    pub data_url: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellRange {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormat {
    pub sqref: Vec<CellRange>,
    pub rules: Vec<CfRule>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CfRule {
    #[serde(rename_all = "camelCase")]
    CellIs { operator: String, formulas: Vec<String>, dxf_id: Option<u32>, priority: i32 },
    #[serde(rename_all = "camelCase")]
    Expression { formula: String, dxf_id: Option<u32>, priority: i32, stop_if_true: bool },
    #[serde(rename_all = "camelCase")]
    ColorScale { stops: Vec<CfStop>, priority: i32 },
    #[serde(rename_all = "camelCase")]
    DataBar { color: String, min: CfValue, max: CfValue, priority: i32, gradient: bool },
    #[serde(rename_all = "camelCase")]
    Top10 { top: bool, percent: bool, rank: u32, dxf_id: Option<u32>, priority: i32 },
    #[serde(rename_all = "camelCase")]
    AboveAverage { above_average: bool, dxf_id: Option<u32>, priority: i32 },
    #[serde(rename_all = "camelCase")]
    IconSet {
        icon_set: String,
        cfvos: Vec<CfValue>,
        reverse: bool,
        priority: i32,
        #[serde(skip_serializing_if = "Option::is_none")]
        custom_icons: Option<Vec<CfIcon>>,
    },
    #[serde(rename_all = "camelCase")]
    Other { kind: String, priority: i32 },
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfIcon {
    pub icon_set: String,
    pub icon_id: u32,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfStop {
    pub kind: String,
    pub value: Option<String>,
    pub color: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CfValue {
    pub kind: String,
    pub value: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Hyperlink {
    pub col: u32,
    pub row: u32,
    pub url: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    pub index: u32,
    pub height: Option<f64>,
    pub cells: Vec<Cell>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    pub col: u32,
    pub row: u32,
    pub col_ref: String,
    pub value: CellValue,
    pub style_index: u32,
    /// Raw `<f>` formula text (ECMA-376 §18.3.1.40), when present. The
    /// renderer uses this to recompute volatile functions like TODAY() /
    /// NOW() at display time so the cached `<v>` (frozen when the file was
    /// last saved) doesn't show a stale date.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase", tag = "type")]
pub enum CellValue {
    #[default]
    Empty,
    Text {
        text: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        runs: Option<Vec<Run>>,
    },
    Number { number: f64 },
    Bool { bool: bool },
    Error { error: String },
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<RunFont>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct RunFont {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct SharedString {
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub runs: Option<Vec<Run>>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Styles {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
    pub num_fmts: Vec<NumFmt>,
    pub dxfs: Vec<Dxf>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Dxf {
    pub font: Option<Font>,
    pub fill: Option<Fill>,
    pub border: Option<Border>,
    /// Number format override applied when the conditional-formatting rule
    /// matches. ECMA-376 §18.8.17 allows `<dxf>` to carry a `<numFmt>` that
    /// replaces the cell's own style numFmt (e.g. switching a calendar cell
    /// from `d` to `m"月"d"日"` on the first of each month).
    pub num_fmt: Option<NumFmt>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Font {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub size: f64,
    pub color: Option<String>,
    pub name: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
    /// When the fill element is a <gradientFill>, this carries the gradient
    /// stops + type + rotation. patternType stays "none" because xlsx does
    /// not mix gradient + pattern in the same fill.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gradient: Option<GradientFillSpec>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct GradientFillSpec {
    /// "linear" (default) or "path". Linear uses `degree`; path uses top/bottom/left/right.
    pub gradient_type: String,
    /// Linear-gradient rotation in degrees (0 = left→right).
    pub degree: f64,
    /// Path-gradient bounding box (0..1 within the cell). Unused for linear.
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub stops: Vec<GradientStopSpec>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct GradientStopSpec {
    pub position: f64,
    pub color: String,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
    /// Diagonal line from bottom-left to top-right (ECMA-376 §18.8.4 diagonalUp)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<BorderEdge>,
    /// Diagonal line from top-left to bottom-right (ECMA-376 §18.8.4 diagonalDown)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<BorderEdge>,
    /// Inner horizontal rule between rows inside a region (ECMA-376 §18.8.40
    /// `tableStyleElement/border/horizontal`). Only set on table-style dxfs;
    /// ignored on cell-level borders.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<BorderEdge>,
    /// Inner vertical rule between columns inside a region (same ECMA-376
    /// section). Only set on table-style dxfs.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub vertical: Option<BorderEdge>,
}

#[derive(Debug, Clone, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct BorderEdge {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct CellXf {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
    pub wrap_text: bool,
    /// Text indentation level (each level ≈ 3 characters wide, ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    /// Text rotation in degrees: 0–90 = counter-clockwise, 91–180 = (value−90)° clockwise, 255 = stacked (ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u32>,
    /// Shrink text to fit the cell width (ECMA-376 §18.8.44)
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub shrink_to_fit: bool,
}

#[derive(Debug, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct NumFmt {
    pub num_fmt_id: u32,
    pub format_code: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParsedWorkbook {
    pub workbook: Workbook,
    pub styles: Styles,
    pub shared_strings: Vec<SharedString>,
}

// Excel built-in indexed color palette (indices 0-63)
// Standard Excel 2003 color palette
const INDEXED_COLORS: &[&str] = &[
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", // 0-7
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", // 8-15
    "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080", // 16-23
    "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF", // 24-31
    "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF", // 32-39
    "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99", // 40-47
    "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696", // 48-55
    "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333", // 56-63
];

#[wasm_bindgen]
pub fn parse_xlsx(data: &[u8]) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    parse_xlsx_inner(data)
        .map(|wb| serde_json::to_string(&wb).unwrap())
        .map_err(|e| JsValue::from_str(&e))
}

#[wasm_bindgen]
pub fn parse_sheet(data: &[u8], sheet_index: u32, name: &str) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    // Resolve rId → target path from workbook.xml.rels
    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let sheet_xml = read_zip_entry(&mut archive, &format!("xl/{}", sheet_path))?;
    let (mut ws, hyperlink_rids) = parse_worksheet(&sheet_xml, &shared_strings, &theme_colors, name)
        .map_err(|e| e.to_string())?;

    // Attach any drawing-anchored images and charts for this sheet
    ws.images = load_sheet_images(&mut archive, &sheet_path);
    ws.charts = load_sheet_charts(&mut archive, &sheet_path, &theme_colors);
    ws.shape_groups = load_sheet_shape_groups(&mut archive, &sheet_path, &theme_colors);
    ws.hyperlinks = load_hyperlinks(&mut archive, &sheet_path, hyperlink_rids);
    ws.comment_refs = load_sheet_comment_refs(&mut archive, &sheet_path);
    ws.defined_names = parse_defined_names_for_sheet(&wb_doc, sheet_index);
    ws.tables = load_sheet_tables(&mut archive, &sheet_path, &theme_colors);
    ws.slicers = load_sheet_slicers(&mut archive, &sheet_path);
    ws.sparkline_groups = load_sheet_sparklines(&mut archive, &sheet_xml, &sheets, &rels_doc, &theme_colors);

    serde_json::to_string(&ws).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn parse_xlsx_inner(data: &[u8]) -> Result<ParsedWorkbook, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let styles = parse_styles(&mut archive, &theme_colors)?;

    Ok(ParsedWorkbook {
        workbook: Workbook { sheets },
        styles,
        shared_strings,
    })
}

/// Refuse to decompress individual ZIP entries larger than 512 MiB to prevent
/// zip-bomb DoS.
const MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;

fn read_zip_entry(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<String, String> {
    let mut file = archive
        .by_name(name)
        .map_err(|e| format!("entry '{}' not found: {}", name, e))?;
    if file.size() > MAX_ZIP_ENTRY_BYTES {
        return Err(format!("entry '{}' exceeds size limit", name));
    }
    let mut buf = String::new();
    file.by_ref().take(MAX_ZIP_ENTRY_BYTES).read_to_string(&mut buf).map_err(|e| e.to_string())?;
    Ok(buf)
}

fn parse_theme_colors(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> Vec<String> {
    let Ok(xml) = read_zip_entry(archive, "xl/theme/theme1.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // Find clrScheme node and collect child color elements in order
    // OOXML order: dk1, lt1, dk2, lt2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink
    let mut colors: Vec<String> = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "clrScheme" && node.tag_name().namespace() == Some(a_ns) {
            for child in node.children() {
                if !child.is_element() { continue; }
                // Each child is a color slot; its first child element holds the actual color
                for color_node in child.children() {
                    if !color_node.is_element() { continue; }
                    let hex = match color_node.tag_name().name() {
                        "srgbClr" => {
                            color_node.attribute("val").map(|v| format!("#{}", v.to_uppercase()))
                        }
                        "sysClr" => {
                            color_node.attribute("lastClr").map(|v| format!("#{}", v.to_uppercase()))
                        }
                        _ => None,
                    };
                    if let Some(h) = hex {
                        colors.push(h);
                        break;
                    }
                }
            }
            break;
        }
    }
    colors
}

/// Convert hex color + tint to resulting hex color using HLS model.
/// tint > 0: lighten; tint < 0: darken.
fn apply_tint(hex: &str, tint: f64) -> String {
    let hex = hex.trim_start_matches('#');
    if hex.len() < 6 { return format!("#{}", hex); }
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64 / 255.0;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64 / 255.0;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64 / 255.0;

    // RGB → HLS
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    let s = if max == min {
        0.0
    } else if l < 0.5 {
        (max - min) / (max + min)
    } else {
        (max - min) / (2.0 - max - min)
    };
    let h = if max == min {
        0.0
    } else if max == r {
        (g - b) / (max - min) / 6.0
    } else if max == g {
        ((b - r) / (max - min) + 2.0) / 6.0
    } else {
        ((r - g) / (max - min) + 4.0) / 6.0
    };
    let h = if h < 0.0 { h + 1.0 } else { h };

    // Apply tint to luminance
    let new_l = if tint > 0.0 {
        l * (1.0 - tint) + tint
    } else {
        l * (1.0 + tint)
    };

    // HLS → RGB
    let (nr, ng, nb) = hls_to_rgb(h, new_l, s);
    format!("#{:02X}{:02X}{:02X}", (nr * 255.0).round() as u8, (ng * 255.0).round() as u8, (nb * 255.0).round() as u8)
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);
    (r, g, b)
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 { t += 1.0; }
    if t > 1.0 { t -= 1.0; }
    if t < 1.0 / 6.0 { return p + (q - p) * 6.0 * t; }
    if t < 1.0 / 2.0 { return q; }
    if t < 2.0 / 3.0 { return p + (q - p) * (2.0 / 3.0 - t) * 6.0; }
    p
}

fn parse_color(node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    // rgb attribute (ARGB: 8 chars, drop alpha; or 6-char RGB)
    if let Some(rgb) = node.attribute("rgb") {
        if rgb.len() == 8 {
            return Some(format!("#{}", &rgb[2..].to_uppercase()));
        }
        return Some(format!("#{}", rgb.to_uppercase()));
    }

    // theme attribute → resolve from theme color array + optional tint
    //
    // ECMA-376 §18.8.3 stores the theme clrScheme in the order
    //   dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink
    // but cell style references (c:color/@theme, c:fgColor/@theme, etc.) use
    // the Excel-internal index where dk1↔lt1 and dk2↔lt2 are SWAPPED:
    //   0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..11 unchanged.
    // This is a well-known interoperability quirk (see Open-XML-SDK issue #46
    // and ECMA-376 §22.1.2.7 where "index values of 0 and 1 are swapped").
    if let Some(theme_str) = node.attribute("theme") {
        if let Ok(idx) = theme_str.parse::<usize>() {
            let mapped = match idx {
                0 => 1,
                1 => 0,
                2 => 3,
                3 => 2,
                n => n,
            };
            if let Some(base) = theme_colors.get(mapped) {
                let tint = node.attribute("tint").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                if tint == 0.0 {
                    return Some(base.clone());
                }
                return Some(apply_tint(base, tint));
            }
        }
    }

    // indexed attribute → Excel built-in palette
    if let Some(indexed_str) = node.attribute("indexed") {
        if let Ok(idx) = indexed_str.parse::<usize>() {
            // indices 64 (foreground) and 65 (background) are special: use black/white
            let color = match idx {
                64 => "#000000",
                65 => "#FFFFFF",
                _ => INDEXED_COLORS.get(idx).copied().unwrap_or("#000000"),
            };
            return Some(color.to_string());
        }
    }

    None
}

fn parse_workbook_sheets(doc: &roxmltree::Document) -> Vec<SheetMeta> {
    let mut sheets = Vec::new();
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "sheet" && node.tag_name().namespace() == Some(ns) {
            let name = node.attribute("name").unwrap_or("Sheet").to_string();
            let sheet_id = node
                .attribute("sheetId")
                .and_then(|v| v.parse().ok())
                .unwrap_or(1);
            let r_id = node
                .attribute((r_ns, "id"))
                .unwrap_or("")
                .to_string();
            sheets.push(SheetMeta { name, sheet_id, r_id });
        }
    }
    sheets
}

/// Collect `<definedName>` entries from `workbook.xml`. `sheet_index` selects
/// which names are in scope: workbook-global (no `localSheetId`) plus any
/// whose `localSheetId` matches the given sheet position.
fn parse_defined_names_for_sheet(doc: &roxmltree::Document, sheet_index: u32) -> Vec<DefinedName> {
    let mut names = Vec::new();
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for node in doc.descendants() {
        if node.tag_name().name() != "definedName" || node.tag_name().namespace() != Some(ns) {
            continue;
        }
        let local: Option<u32> = node.attribute("localSheetId").and_then(|s| s.parse().ok());
        if let Some(l) = local { if l != sheet_index { continue; } }
        let name = match node.attribute("name") { Some(n) => n.to_string(), None => continue };
        let formula = node.text().unwrap_or("").to_string();
        names.push(DefinedName { name, formula });
    }
    names
}

fn resolve_sheet_path(doc: &roxmltree::Document, r_id: &str) -> Option<String> {
    let ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    for node in doc.descendants() {
        if node.tag_name().name() == "Relationship" && node.tag_name().namespace() == Some(ns) {
            if node.attribute("Id") == Some(r_id) {
                return node.attribute("Target").map(|s| s.to_string());
            }
        }
    }
    None
}

fn read_shared_strings(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    theme_colors: &[String],
) -> Vec<SharedString> {
    let Ok(xml) = read_zip_entry(archive, "xl/sharedStrings.xml") else {
        return Vec::new();
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return Vec::new();
    };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut strings = Vec::new();
    for si in doc.descendants() {
        if si.tag_name().name() == "si" && si.tag_name().namespace() == Some(ns) {
            strings.push(parse_si_node(&si, ns, theme_colors));
        }
    }
    strings
}

/// Parse a `<si>` (shared) or `<is>` (inline) node into a SharedString.
/// The node may contain direct `<t>` text (plain) and/or multiple `<r>`
/// runs with per-run `<rPr>` font properties.
fn parse_si_node(
    node: &roxmltree::Node,
    ns: &str,
    theme_colors: &[String],
) -> SharedString {
    let mut text = String::new();
    let mut runs: Vec<Run> = Vec::new();
    let mut has_runs = false;
    for child in node.children() {
        if !child.is_element() { continue; }
        match child.tag_name().name() {
            "t" if child.tag_name().namespace() == Some(ns) => {
                if let Some(s) = child.text() {
                    text.push_str(s);
                }
            }
            "r" if child.tag_name().namespace() == Some(ns) => {
                has_runs = true;
                let mut run_text = String::new();
                let mut run_font: Option<RunFont> = None;
                for rc in child.children() {
                    match rc.tag_name().name() {
                        "t" => {
                            if let Some(s) = rc.text() {
                                run_text.push_str(s);
                            }
                        }
                        "rPr" => {
                            let mut f = RunFont::default();
                            for rp in rc.children() {
                                match rp.tag_name().name() {
                                    "b" => f.bold = true,
                                    "i" => f.italic = true,
                                    "u" => f.underline = true,
                                    "strike" => f.strike = true,
                                    "sz" => {
                                        f.size = rp.attribute("val").and_then(|s| s.parse().ok());
                                    }
                                    "color" => {
                                        f.color = parse_color(&rp, theme_colors);
                                    }
                                    "rFont" | "name" => {
                                        f.name = rp.attribute("val").map(|s| s.to_string());
                                    }
                                    _ => {}
                                }
                            }
                            run_font = Some(f);
                        }
                        _ => {}
                    }
                }
                text.push_str(&run_text);
                runs.push(Run { text: run_text, font: run_font });
            }
            _ => {}
        }
    }
    SharedString {
        text,
        runs: if has_runs { Some(runs) } else { None },
    }
}

fn parse_styles(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, theme_colors: &[String]) -> Result<Styles, String> {
    let xml = read_zip_entry(archive, "xl/styles.xml")?;
    let doc = roxmltree::Document::parse(&xml).map_err(|e| e.to_string())?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let num_fmts = parse_num_fmts(&doc, ns);
    let fonts = parse_fonts(&doc, ns, theme_colors);
    let fills = parse_fills(&doc, ns, theme_colors);
    let borders = parse_borders(&doc, ns, theme_colors);
    let cell_xfs = parse_cell_xfs(&doc, ns);
    let dxfs = parse_dxfs(&doc, ns, theme_colors);

    Ok(Styles { fonts, fills, borders, cell_xfs, num_fmts, dxfs })
}

fn parse_dxfs(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Dxf> {
    let mut dxfs = Vec::new();
    for dxfs_node in doc.descendants() {
        if dxfs_node.tag_name().name() != "dxfs" || dxfs_node.tag_name().namespace() != Some(ns) {
            continue;
        }
        for dxf_node in dxfs_node.children() {
            if dxf_node.tag_name().name() != "dxf" { continue; }
            let mut d = Dxf::default();
            for child in dxf_node.children() {
                match child.tag_name().name() {
                    "font" => {
                        let mut f = Font { size: 11.0, ..Default::default() };
                        for fc in child.children() {
                            match fc.tag_name().name() {
                                "b" => f.bold = true,
                                "i" => f.italic = true,
                                "u" => f.underline = true,
                                "strike" => f.strike = true,
                                "sz" => {
                                    if let Some(v) = fc.attribute("val").and_then(|s| s.parse().ok()) {
                                        f.size = v;
                                    }
                                }
                                "name" => {
                                    f.name = fc.attribute("val").map(|s| s.to_string());
                                }
                                "color" => {
                                    f.color = parse_color(&fc, theme_colors);
                                }
                                _ => {}
                            }
                        }
                        d.font = Some(f);
                    }
                    "fill" => {
                        let mut f = Fill::default();
                        for pf in child.children() {
                            if pf.tag_name().name() == "patternFill" {
                                f.pattern_type = pf.attribute("patternType").unwrap_or("solid").to_string();
                                for color_node in pf.children() {
                                    match color_node.tag_name().name() {
                                        "fgColor" => f.fg_color = parse_color(&color_node, theme_colors),
                                        "bgColor" => f.bg_color = parse_color(&color_node, theme_colors),
                                        _ => {}
                                    }
                                }
                            }
                        }
                        // In dxf, conditional format fills often only have bgColor; mirror into fgColor
                        if f.fg_color.is_none() && f.bg_color.is_some() {
                            f.fg_color = f.bg_color.clone();
                        }
                        d.fill = Some(f);
                    }
                    "border" => {
                        let mut b = Border::default();
                        for edge_node in child.children() {
                            let style = edge_node.attribute("style").unwrap_or("").to_string();
                            if style.is_empty() { continue; }
                            let color = edge_node.children().find(|c| c.is_element())
                                .and_then(|c| parse_color(&c, theme_colors));
                            let edge = Some(BorderEdge { style, color });
                            match edge_node.tag_name().name() {
                                "left" => b.left = edge,
                                "right" => b.right = edge,
                                "top" => b.top = edge,
                                "bottom" => b.bottom = edge,
                                "horizontal" => b.horizontal = edge,
                                "vertical"   => b.vertical   = edge,
                                _ => {}
                            }
                        }
                        d.border = Some(b);
                    }
                    "numFmt" => {
                        let num_fmt_id = child.attribute("numFmtId")
                            .and_then(|v| v.parse().ok()).unwrap_or(0);
                        let format_code = child.attribute("formatCode")
                            .unwrap_or("").to_string();
                        d.num_fmt = Some(NumFmt { num_fmt_id, format_code });
                    }
                    _ => {}
                }
            }
            dxfs.push(d);
        }
        break;
    }
    dxfs
}

fn parse_num_fmts(doc: &roxmltree::Document, ns: &str) -> Vec<NumFmt> {
    let mut fmts = Vec::new();
    for node in doc.descendants() {
        if node.tag_name().name() == "numFmts" && node.tag_name().namespace() == Some(ns) {
            for child in node.children() {
                if child.tag_name().name() != "numFmt" { continue; }
                let num_fmt_id = child.attribute("numFmtId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let format_code = child.attribute("formatCode").unwrap_or("").to_string();
                fmts.push(NumFmt { num_fmt_id, format_code });
            }
            break;
        }
    }
    fmts
}

fn parse_fonts(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Font> {
    let mut fonts = Vec::new();
    for fonts_node in doc.descendants() {
        if fonts_node.tag_name().name() == "fonts" && fonts_node.tag_name().namespace() == Some(ns) {
            for font_node in fonts_node.children() {
                if font_node.tag_name().name() != "font" { continue; }
                let mut f = Font { size: 11.0, ..Default::default() };
                for child in font_node.children() {
                    match child.tag_name().name() {
                        "b" => f.bold = true,
                        "i" => f.italic = true,
                        "u" => f.underline = true,
                        "strike" => f.strike = true,
                        "sz" => {
                            if let Some(v) = child.attribute("val").and_then(|s| s.parse().ok()) {
                                f.size = v;
                            }
                        }
                        "name" => {
                            f.name = child.attribute("val").map(|s| s.to_string());
                        }
                        "color" => {
                            f.color = parse_color(&child, theme_colors);
                        }
                        _ => {}
                    }
                }
                fonts.push(f);
            }
            break;
        }
    }
    fonts
}

fn parse_fills(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Fill> {
    let mut fills = Vec::new();
    for fills_node in doc.descendants() {
        if fills_node.tag_name().name() == "fills" && fills_node.tag_name().namespace() == Some(ns) {
            for fill_node in fills_node.children() {
                if fill_node.tag_name().name() != "fill" { continue; }
                let mut f = Fill::default();
                for pf in fill_node.children() {
                    match pf.tag_name().name() {
                        "patternFill" => {
                            f.pattern_type = pf.attribute("patternType").unwrap_or("none").to_string();
                            for color_node in pf.children() {
                                match color_node.tag_name().name() {
                                    "fgColor" => f.fg_color = parse_color(&color_node, theme_colors),
                                    "bgColor" => f.bg_color = parse_color(&color_node, theme_colors),
                                    _ => {}
                                }
                            }
                        }
                        "gradientFill" => {
                            // ECMA-376 §18.8.24 gradientFill — linear (default) uses
                            // `degree`, path uses top/bottom/left/right as a relative
                            // bounding box; children <stop position="n"><color/></stop>.
                            let gtype = pf.attribute("type").unwrap_or("linear").to_string();
                            let degree = pf.attribute("degree").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                            let left   = pf.attribute("left").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                            let right  = pf.attribute("right").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                            let top    = pf.attribute("top").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                            let bottom = pf.attribute("bottom").and_then(|s| s.parse::<f64>().ok()).unwrap_or(0.0);
                            let mut stops: Vec<GradientStopSpec> = pf.children()
                                .filter(|n| n.is_element() && n.tag_name().name() == "stop")
                                .filter_map(|stop| {
                                    let position = stop.attribute("position").and_then(|s| s.parse::<f64>().ok())?;
                                    let color_node = stop.children().find(|c| c.is_element() && c.tag_name().name() == "color")?;
                                    let color = parse_color(&color_node, theme_colors)?;
                                    Some(GradientStopSpec { position, color })
                                })
                                .collect();
                            stops.sort_by(|a, b| a.position.partial_cmp(&b.position).unwrap_or(std::cmp::Ordering::Equal));
                            if !stops.is_empty() {
                                f.gradient = Some(GradientFillSpec {
                                    gradient_type: gtype,
                                    degree,
                                    left, right, top, bottom,
                                    stops,
                                });
                            }
                        }
                        _ => {}
                    }
                }
                fills.push(f);
            }
            break;
        }
    }
    fills
}

fn parse_borders(doc: &roxmltree::Document, ns: &str, theme_colors: &[String]) -> Vec<Border> {
    let mut borders = Vec::new();
    for borders_node in doc.descendants() {
        if borders_node.tag_name().name() == "borders" && borders_node.tag_name().namespace() == Some(ns) {
            for border_node in borders_node.children() {
                if border_node.tag_name().name() != "border" { continue; }
                let has_diag_up = border_node.attribute("diagonalUp").map(|v| v == "1" || v == "true").unwrap_or(false);
                let has_diag_down = border_node.attribute("diagonalDown").map(|v| v == "1" || v == "true").unwrap_or(false);
                let mut b = Border::default();
                let mut diag_edge: Option<BorderEdge> = None;
                for edge_node in border_node.children() {
                    let style = edge_node.attribute("style").unwrap_or("").to_string();
                    let color = edge_node.children().find(|c| c.is_element()).and_then(|c| parse_color(&c, theme_colors));
                    match edge_node.tag_name().name() {
                        "left" if !style.is_empty() => b.left = Some(BorderEdge { style, color }),
                        "right" if !style.is_empty() => b.right = Some(BorderEdge { style, color }),
                        "top" if !style.is_empty() => b.top = Some(BorderEdge { style, color }),
                        "bottom" if !style.is_empty() => b.bottom = Some(BorderEdge { style, color }),
                        "diagonal" if !style.is_empty() => diag_edge = Some(BorderEdge { style, color }),
                        _ => {}
                    }
                }
                if has_diag_up { b.diagonal_up = diag_edge.clone(); }
                if has_diag_down { b.diagonal_down = diag_edge; }
                borders.push(b);
            }
            break;
        }
    }
    borders
}

fn parse_cell_xfs(doc: &roxmltree::Document, ns: &str) -> Vec<CellXf> {
    let mut xfs = Vec::new();
    for xfs_node in doc.descendants() {
        if xfs_node.tag_name().name() == "cellXfs" && xfs_node.tag_name().namespace() == Some(ns) {
            for xf_node in xfs_node.children() {
                if xf_node.tag_name().name() != "xf" { continue; }
                let font_id = xf_node.attribute("fontId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let fill_id = xf_node.attribute("fillId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let border_id = xf_node.attribute("borderId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let num_fmt_id = xf_node.attribute("numFmtId").and_then(|v| v.parse().ok()).unwrap_or(0);
                let mut align_h = None;
                let mut align_v = None;
                let mut wrap_text = false;
                let mut indent = None;
                let mut text_rotation = None;
                let mut shrink_to_fit = false;
                for child in xf_node.children() {
                    if child.tag_name().name() == "alignment" {
                        align_h = child.attribute("horizontal").map(|s| s.to_string());
                        align_v = child.attribute("vertical").map(|s| s.to_string());
                        wrap_text = child.attribute("wrapText").map(|v| v == "1" || v == "true").unwrap_or(false);
                        indent = child.attribute("indent").and_then(|s| s.parse::<u32>().ok()).filter(|&v| v > 0);
                        text_rotation = child.attribute("textRotation").and_then(|s| s.parse::<u32>().ok()).filter(|&v| v > 0);
                        shrink_to_fit = child.attribute("shrinkToFit").map(|v| v == "1" || v == "true").unwrap_or(false);
                    }
                }
                xfs.push(CellXf { font_id, fill_id, border_id, num_fmt_id, align_h, align_v, wrap_text, indent, text_rotation, shrink_to_fit });
            }
            break;
        }
    }
    xfs
}

fn parse_worksheet(
    xml: &str,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    name: &str,
) -> Result<(Worksheet, Vec<(u32, u32, String)>), String> {
    let doc = roxmltree::Document::parse(xml).map_err(|e| e.to_string())?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    let mut rows = Vec::new();
    let mut col_widths: HashMap<u32, f64> = HashMap::new();
    let mut row_heights: HashMap<u32, f64> = HashMap::new();
    let mut merge_cells: Vec<MergeCell> = Vec::new();
    let mut freeze_rows: u32 = 0;
    let mut freeze_cols: u32 = 0;
    let mut default_col_width = 8.43;
    let mut default_row_height = 15.0;
    let mut conditional_formats: Vec<ConditionalFormat> = Vec::new();
    let mut show_zeros = true;
    let mut show_gridlines = true;
    let mut tab_color: Option<String> = None;
    let mut auto_filter: Option<CellRange> = None;
    let mut hyperlink_rids: Vec<(u32, u32, String)> = Vec::new();

    // Pre-scan worksheet-level extLst for x14:dataBar extension attributes.
    // Excel 2010+ stores the `gradient` flag on `<x14:dataBar>` inside
    // `<extLst>/<ext>/<x14:conditionalFormattings>/<x14:conditionalFormatting>
    // /<x14:cfRule id="{GUID}">`, linked to the SpreadsheetML cfRule via a
    // matching `<x14:id>{GUID}</x14:id>` inside the cfRule's own extLst
    // (§2.6.3). Build a GUID → gradient map so cfRule parsing can look up
    // the override.
    let mut x14_databar_gradient: HashMap<String, bool> = HashMap::new();
    for x14_rule in doc.descendants().filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("dataBar")) {
        let Some(id) = x14_rule.attribute("id") else { continue };
        for bar in x14_rule.children().filter(|n| n.tag_name().name() == "dataBar") {
            if let Some(g) = bar.attribute("gradient") {
                x14_databar_gradient.insert(id.to_string(), !(g == "0" || g == "false"));
            }
        }
    }

    // Pre-scan worksheet-level extLst for x14:conditionalFormatting with
    // iconSet rules. Excel 2010+ stores custom icon sets (custom="1") here
    // with per-threshold `<x14:cfIcon iconSet="X" iconId="N"/>` overrides,
    // and cfvo values inside `<xm:f>` children instead of `val` attributes.
    // The sqref for x14 CF rules lives in a `<xm:sqref>` sibling.
    let mut x14_icon_formats: Vec<ConditionalFormat> = Vec::new();
    for x14_cf in doc.descendants().filter(|n| n.tag_name().name() == "conditionalFormatting" && n.tag_name().namespace().map(|u| u.contains("/spreadsheetml/2009/9")).unwrap_or(false)) {
        let sqref: Vec<CellRange> = x14_cf.children()
            .find(|n| n.tag_name().name() == "sqref")
            .and_then(|n| n.text())
            .map(parse_sqref)
            .unwrap_or_default();
        if sqref.is_empty() { continue; }
        let mut rules: Vec<CfRule> = Vec::new();
        for x14_rule in x14_cf.children().filter(|n| n.tag_name().name() == "cfRule" && n.attribute("type") == Some("iconSet")) {
            let priority: i32 = x14_rule.attribute("priority").and_then(|s| s.parse().ok()).unwrap_or(0);
            let Some(icon_node) = x14_rule.children().find(|n| n.tag_name().name() == "iconSet") else { continue };
            let custom = icon_node.attribute("custom").map(|v| v == "1" || v == "true").unwrap_or(false);
            let icon_set_name = icon_node.attribute("iconSet")
                .unwrap_or(if custom { "" } else { "3TrafficLights1" })
                .to_string();
            let reverse = icon_node.attribute("reverse").map(|v| v == "1" || v == "true").unwrap_or(false);
            let mut cfvos: Vec<CfValue> = Vec::new();
            let mut custom_icons: Vec<CfIcon> = Vec::new();
            for ch in icon_node.children().filter(|n| n.is_element()) {
                match ch.tag_name().name() {
                    "cfvo" => {
                        let kind = ch.attribute("type").unwrap_or("percent").to_string();
                        // x14:cfvo stores the value in `<xm:f>` child; attribute val fallback.
                        let value = ch.children()
                            .find(|n| n.tag_name().name() == "f")
                            .and_then(|n| n.text())
                            .map(|s| s.to_string())
                            .or_else(|| ch.attribute("val").map(|s| s.to_string()));
                        cfvos.push(CfValue { kind, value });
                    }
                    "cfIcon" => {
                        let set = ch.attribute("iconSet").unwrap_or("NoIcons").to_string();
                        let id = ch.attribute("iconId").and_then(|s| s.parse().ok()).unwrap_or(0);
                        custom_icons.push(CfIcon { icon_set: set, icon_id: id });
                    }
                    _ => {}
                }
            }
            rules.push(CfRule::IconSet {
                icon_set: icon_set_name,
                cfvos,
                reverse,
                priority,
                custom_icons: if custom { Some(custom_icons) } else { None },
            });
        }
        if !rules.is_empty() {
            x14_icon_formats.push(ConditionalFormat { sqref, rules });
        }
    }

    for node in doc.descendants() {
        match node.tag_name().name() {
            "sheetFormatPr" if node.tag_name().namespace() == Some(ns) => {
                if let Some(v) = node.attribute("defaultColWidth").and_then(|s| s.parse().ok()) {
                    default_col_width = v;
                }
                if let Some(v) = node.attribute("defaultRowHeight").and_then(|s| s.parse().ok()) {
                    default_row_height = v;
                }
            }
            "col" if node.tag_name().namespace() == Some(ns) => {
                let custom = node.attribute("customWidth").map(|v| v == "1").unwrap_or(false);
                let hidden = node.attribute("hidden").map(|v| v == "1").unwrap_or(false);
                // Only record widths for custom-widthed columns OR hidden columns
                if !custom && !hidden { continue; }
                let min: u32 = node.attribute("min").and_then(|s| s.parse().ok()).unwrap_or(1);
                let max: u32 = node.attribute("max").and_then(|s| s.parse().ok()).unwrap_or(1);
                // Cap range to avoid storing 16K entries for max=16384 ranges
                let max = max.min(min + 255);
                let width: f64 = if hidden {
                    0.0
                } else {
                    node.attribute("width").and_then(|s| s.parse().ok()).unwrap_or(default_col_width)
                };
                for c in min..=max {
                    col_widths.insert(c, width);
                }
            }
            "sheetView" if node.tag_name().namespace() == Some(ns) => {
                show_zeros = node.attribute("showZeros").map(|v| v != "0").unwrap_or(true);
                show_gridlines = node.attribute("showGridLines").map(|v| v != "0").unwrap_or(true);
            }
            "tabColor" if node.tag_name().namespace() == Some(ns) => {
                tab_color = parse_color(&node, theme_colors);
            }
            "autoFilter" if node.tag_name().namespace() == Some(ns) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    auto_filter = if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        Some(CellRange { top, left, bottom, right })
                    } else {
                        let (col, row) = parse_cell_ref(parts[0]);
                        Some(CellRange { top: row, left: col, bottom: row, right: col })
                    };
                }
            }
            "hyperlinks" if node.tag_name().namespace() == Some(ns) => {
                for hl in node.children() {
                    if !hl.is_element() || hl.tag_name().name() != "hyperlink" { continue; }
                    let Some(ref_str) = hl.attribute("ref") else { continue };
                    // Only first cell of ref range
                    let ref_single = ref_str.split(':').next().unwrap_or(ref_str);
                    let (col, row) = parse_cell_ref(ref_single);
                    if let Some(rid) = hl.attributes()
                        .find(|a| a.name() == "id" && a.namespace() == Some(r_ns))
                        .map(|a| a.value().to_string())
                    {
                        hyperlink_rids.push((col, row, rid));
                    }
                }
            }
            "pane" if node.tag_name().namespace() == Some(ns) => {
                let state = node.attribute("state").unwrap_or("");
                if state == "frozen" || state == "frozenSplit" {
                    freeze_rows = node.attribute("ySplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| v as u32)
                        .unwrap_or(0);
                    freeze_cols = node.attribute("xSplit")
                        .and_then(|s| s.parse::<f64>().ok())
                        .map(|v| v as u32)
                        .unwrap_or(0);
                }
            }
            "mergeCell" if node.tag_name().namespace() == Some(ns) => {
                if let Some(r) = node.attribute("ref") {
                    let parts: Vec<&str> = r.split(':').collect();
                    if parts.len() == 2 {
                        let (left, top) = parse_cell_ref(parts[0]);
                        let (right, bottom) = parse_cell_ref(parts[1]);
                        merge_cells.push(MergeCell { top, left, bottom, right });
                    }
                }
            }
            "row" if node.tag_name().namespace() == Some(ns) => {
                let row_idx: u32 = node.attribute("r").and_then(|s| s.parse().ok()).unwrap_or(0);
                let hidden = node.attribute("hidden").map(|v| v == "1").unwrap_or(false);
                let height: Option<f64> = if hidden {
                    Some(0.0)
                } else {
                    node.attribute("ht").and_then(|s| s.parse().ok())
                };
                if let Some(h) = height {
                    row_heights.insert(row_idx, h);
                }
                let cells = parse_row_cells(&node, shared_strings, theme_colors, ns);
                rows.push(Row { index: row_idx, height, cells });
            }
            "conditionalFormatting" if node.tag_name().namespace() == Some(ns) => {
                let sqref = node.attribute("sqref")
                    .map(|s| parse_sqref(s))
                    .unwrap_or_default();
                let mut rules: Vec<CfRule> = Vec::new();
                for cf in node.children() {
                    if cf.tag_name().name() != "cfRule" { continue; }
                    let kind = cf.attribute("type").unwrap_or("").to_string();
                    let priority: i32 = cf.attribute("priority").and_then(|s| s.parse().ok()).unwrap_or(0);
                    let dxf_id: Option<u32> = cf.attribute("dxfId").and_then(|s| s.parse().ok());
                    match kind.as_str() {
                        "cellIs" => {
                            let operator = cf.attribute("operator").unwrap_or("equal").to_string();
                            let formulas: Vec<String> = cf.children()
                                .filter(|n| n.tag_name().name() == "formula")
                                .filter_map(|n| n.text().map(|s| s.to_string()))
                                .collect();
                            rules.push(CfRule::CellIs { operator, formulas, dxf_id, priority });
                        }
                        "expression"
                        | "containsBlanks" | "notContainsBlanks"
                        | "containsText" | "notContainsText"
                        | "beginsWith" | "endsWith"
                        | "containsErrors" | "notContainsErrors" => {
                            // For `containsBlanks`/`notContainsBlanks`/`containsText` etc.,
                            // Excel serializes an equivalent boolean formula (e.g.
                            // `LEN(TRIM(C8))>0`) as the rule's `<formula>` child
                            // (ECMA-376 §18.3.1.10). Evaluate as an expression rule.
                            let formula = cf.children()
                                .find(|n| n.tag_name().name() == "formula")
                                .and_then(|n| n.text())
                                .unwrap_or("")
                                .to_string();
                            let stop_if_true = cf.attribute("stopIfTrue")
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            rules.push(CfRule::Expression { formula, dxf_id, priority, stop_if_true });
                        }
                        "colorScale" => {
                            let scale = cf.children().find(|n| n.tag_name().name() == "colorScale");
                            let mut stop_values: Vec<(String, Option<String>)> = Vec::new();
                            let mut stop_colors: Vec<String> = Vec::new();
                            if let Some(scale_node) = scale {
                                for child in scale_node.children() {
                                    match child.tag_name().name() {
                                        "cfvo" => {
                                            stop_values.push((
                                                child.attribute("type").unwrap_or("num").to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            stop_colors.push(parse_color(&child, theme_colors).unwrap_or_else(|| "#FFFFFF".to_string()));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            let stops: Vec<CfStop> = stop_values.into_iter().enumerate().map(|(i, (kind, value))| CfStop {
                                kind,
                                value,
                                color: stop_colors.get(i).cloned().unwrap_or_else(|| "#FFFFFF".to_string()),
                            }).collect();
                            rules.push(CfRule::ColorScale { stops, priority });
                        }
                        "dataBar" => {
                            let bar = cf.children().find(|n| n.tag_name().name() == "dataBar");
                            let mut cfvos: Vec<(String, Option<String>)> = Vec::new();
                            let mut color = "#638EC6".to_string();
                            if let Some(bar_node) = bar {
                                for child in bar_node.children() {
                                    match child.tag_name().name() {
                                        "cfvo" => {
                                            cfvos.push((
                                                child.attribute("type").unwrap_or("min").to_string(),
                                                child.attribute("val").map(|s| s.to_string()),
                                            ));
                                        }
                                        "color" => {
                                            if let Some(c) = parse_color(&child, theme_colors) { color = c; }
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            // Excel 2010+ x14:dataBar extension may override the
                            // gradient flag (§2.6.3, default="1"). "0" → solid
                            // fill. The override lives in a separate
                            // worksheet-level extLst and is linked via the
                            // `<x14:id>{GUID}</x14:id>` contained in this
                            // cfRule's own extLst.
                            let mut gradient = true;
                            'gradient_lookup: for ext_list in cf.children().filter(|n| n.tag_name().name() == "extLst") {
                                for ext in ext_list.children().filter(|n| n.tag_name().name() == "ext") {
                                    for id_node in ext.descendants().filter(|n| n.tag_name().name() == "id") {
                                        if let Some(guid) = id_node.text() {
                                            if let Some(&g) = x14_databar_gradient.get(guid) {
                                                gradient = g;
                                                break 'gradient_lookup;
                                            }
                                        }
                                    }
                                    // Fallback: some files embed <x14:dataBar>
                                    // directly in the cfRule's extLst.
                                    for x14_bar in ext.descendants().filter(|n| n.tag_name().name() == "dataBar") {
                                        if let Some(g) = x14_bar.attribute("gradient") {
                                            gradient = !(g == "0" || g == "false");
                                            break 'gradient_lookup;
                                        }
                                    }
                                }
                            }
                            let min = cfvos.first().map(|(k, v)| CfValue { kind: k.clone(), value: v.clone() })
                                .unwrap_or(CfValue { kind: "min".into(), value: None });
                            let max = cfvos.get(1).map(|(k, v)| CfValue { kind: k.clone(), value: v.clone() })
                                .unwrap_or(CfValue { kind: "max".into(), value: None });
                            rules.push(CfRule::DataBar { color, min, max, priority, gradient });
                        }
                        "top10" => {
                            let top = !cf.attribute("bottom").map(|v| v == "1" || v == "true").unwrap_or(false);
                            let percent = cf.attribute("percent").map(|v| v == "1" || v == "true").unwrap_or(false);
                            let rank = cf.attribute("rank").and_then(|s| s.parse().ok()).unwrap_or(10);
                            rules.push(CfRule::Top10 { top, percent, rank, dxf_id, priority });
                        }
                        "aboveAverage" => {
                            let above_average = cf.attribute("aboveAverage").map(|v| v != "0").unwrap_or(true);
                            rules.push(CfRule::AboveAverage { above_average, dxf_id, priority });
                        }
                        "iconSet" => {
                            let icon_set_node = cf.children().find(|n| n.tag_name().name() == "iconSet");
                            let icon_set = icon_set_node
                                .and_then(|n| n.attribute("iconSet"))
                                .unwrap_or("3TrafficLights1")
                                .to_string();
                            let reverse = icon_set_node
                                .and_then(|n| n.attribute("reverse"))
                                .map(|v| v == "1" || v == "true")
                                .unwrap_or(false);
                            let cfvos: Vec<CfValue> = icon_set_node
                                .map(|n| n.children()
                                    .filter(|c| c.is_element() && c.tag_name().name() == "cfvo")
                                    .map(|c| CfValue {
                                        kind: c.attribute("type").unwrap_or("percent").to_string(),
                                        value: c.attribute("val").map(|s| s.to_string()),
                                    })
                                    .collect()
                                )
                                .unwrap_or_default();
                            rules.push(CfRule::IconSet { icon_set, cfvos, reverse, priority, custom_icons: None });
                        }
                        other => {
                            rules.push(CfRule::Other { kind: other.to_string(), priority });
                        }
                    }
                }
                conditional_formats.push(ConditionalFormat { sqref, rules });
            }
            _ => {}
        }
    }

    conditional_formats.extend(x14_icon_formats);

    Ok((Worksheet {
        name: name.to_string(),
        rows,
        col_widths,
        row_heights,
        default_col_width,
        default_row_height,
        merge_cells,
        freeze_rows,
        freeze_cols,
        conditional_formats,
        images: Vec::new(),
        charts: Vec::new(),
        shape_groups: Vec::new(),
        show_zeros,
        show_gridlines,
        tab_color,
        auto_filter,
        hyperlinks: Vec::new(),
        comment_refs: Vec::new(),
        defined_names: Vec::new(),
        tables: Vec::new(),
        slicers: Vec::new(),
        sparkline_groups: Vec::new(),
    }, hyperlink_rids))
}

/// Parse a .rels file into rId → Target map.
fn parse_rels_map(xml: &str) -> HashMap<String, String> {
    let Ok(doc) = roxmltree::Document::parse(xml) else {
        return HashMap::new();
    };
    let mut map = HashMap::new();
    for rel in doc.root_element().children().filter(|n| n.is_element()) {
        if let (Some(id), Some(target)) = (rel.attribute("Id"), rel.attribute("Target")) {
            map.insert(id.to_string(), target.to_string());
        }
    }
    map
}

/// Parse xl/comments{N}.xml referenced from the sheet's rels and collect the
/// list of A1-style cell refs that have a `<comment>` associated. The
/// renderer draws a small red triangle in each cell's top-right corner to
/// indicate the presence of a comment (ECMA-376 §18.7.3 commentList).
fn load_sheet_comment_refs(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
) -> Vec<String> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&rels_xml) else { return Vec::new(); };

    // Accept both plain ("/comments") and threaded ("/threadedComment") relTypes
    // but prefer the classic comments file — threaded comments live in a
    // separate namespace and are an extension.
    let mut comments_target: Option<String> = None;
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/comments") {
            if let Some(t) = rel.attribute("Target") {
                comments_target = Some(t.to_string());
                break;
            }
        }
    }
    let Some(target) = comments_target else { return Vec::new(); };

    let comments_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
    let Ok(comments_xml) = read_zip_entry(archive, &comments_path) else { return Vec::new(); };
    let Ok(comments_doc) = roxmltree::Document::parse(&comments_xml) else { return Vec::new(); };

    let mut refs: Vec<String> = Vec::new();
    for node in comments_doc.descendants() {
        if node.tag_name().name() == "comment" && node.is_element() {
            if let Some(r) = node.attribute("ref") {
                refs.push(r.to_string());
            }
        }
    }
    refs
}

/// Parse `xl/tables/tableN.xml` files referenced from the sheet rels and
/// collect them for the renderer. Each table carries a ref range, style name
/// (e.g. "TableStyleLight18"), and the banded-rows / banded-cols flags from
/// `<tableStyleInfo>` (ECMA-376 §18.5).
/// Resolve a built-in table style's accent color from the theme.
///
/// Built-in style names follow the pattern `TableStyle{Light|Medium|Dark}{N}`
/// (ECMA-376 §18.5.1.4). Excel's UI lays the 21/28/11 built-ins out in a grid
/// of rows × 7 columns: column 0 is a "none" style (no accent), columns 1–6
/// map to accent1–accent6. So the accent index is `(N - 1) mod 7` where 0
/// means "no accent" and 1..=6 map to the theme's accent slots.
///
/// `theme_colors` is in OOXML natural order — accent1 lives at index 4, so
/// accent_n is at `theme_colors[3 + n]`. Falls back to a neutral gray when
/// the style name is unrecognised or the theme is missing accents.
/// dxf indices for the ECMA-376 §18.8.40 `<tableStyleElement>` roles we care
/// about. Built-in styles (`TableStyleLight18`, etc.) have no entry in the
/// file's `<tableStyles>` block and fall through to accent-based rendering;
/// custom styles (`"Gift Budget"`) reference dxfs from `<dxfs>`.
#[derive(Debug, Clone, Default)]
struct TableStyleElements {
    whole_table: Option<u32>,
    header_row: Option<u32>,
}

/// Parse `<tableStyles><tableStyle name="…"><tableStyleElement type="…" dxfId="…"/>`
/// into a lookup keyed by table-style name.
fn parse_table_styles_map(archive: &mut zip::ZipArchive<Cursor<&[u8]>>) -> std::collections::HashMap<String, TableStyleElements> {
    use std::collections::HashMap;
    let mut map: HashMap<String, TableStyleElements> = HashMap::new();
    let Ok(xml) = read_zip_entry(archive, "xl/styles.xml") else { return map; };
    let Ok(doc) = roxmltree::Document::parse(&xml) else { return map; };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for n in doc.descendants() {
        if n.tag_name().name() != "tableStyles" || n.tag_name().namespace() != Some(ns) { continue; }
        for ts in n.children().filter(|c| c.is_element() && c.tag_name().name() == "tableStyle") {
            let Some(name) = ts.attribute("name") else { continue; };
            let mut elems = TableStyleElements::default();
            for el in ts.children().filter(|c| c.is_element() && c.tag_name().name() == "tableStyleElement") {
                let t = el.attribute("type").unwrap_or("");
                let dxf: Option<u32> = el.attribute("dxfId").and_then(|s| s.parse().ok());
                match t {
                    "wholeTable" => elems.whole_table = dxf,
                    "headerRow"  => elems.header_row = dxf,
                    _ => {}
                }
            }
            map.insert(name.to_string(), elems);
        }
    }
    map
}

fn resolve_table_style_accent(style_name: &str, theme_colors: &[String]) -> String {
    let fallback = "#808080".to_string();
    let Some(rest) = style_name.strip_prefix("TableStyle") else { return fallback; };
    let digits_start = rest.find(|c: char| c.is_ascii_digit());
    let Some(start) = digits_start else { return fallback; };
    let Ok(n) = rest[start..].parse::<u32>() else { return fallback; };
    if n == 0 { return fallback; }
    let slot = ((n - 1) % 7) as usize;
    if slot == 0 { return fallback; }
    theme_colors.get(3 + slot).cloned().unwrap_or(fallback)
}

fn load_sheet_tables(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<TableInfo> {
    let custom_styles = parse_table_styles_map(archive);
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&rels_xml) else { return Vec::new(); };

    let mut table_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/table") {
            if let Some(t) = rel.attribute("Target") {
                table_targets.push(t.to_string());
            }
        }
    }

    let mut tables: Vec<TableInfo> = Vec::new();
    for target in table_targets {
        let table_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(xml) = read_zip_entry(archive, &table_path) else { continue; };
        let Ok(doc) = roxmltree::Document::parse(&xml) else { continue; };
        let root = doc.root_element();
        let Some(ref_attr) = root.attribute("ref") else { continue };
        let parts: Vec<&str> = ref_attr.split(':').collect();
        let range = if parts.len() == 2 {
            let (left, top) = parse_cell_ref(parts[0]);
            let (right, bottom) = parse_cell_ref(parts[1]);
            CellRange { top, left, bottom, right }
        } else {
            let (col, row) = parse_cell_ref(parts[0]);
            CellRange { top: row, left: col, bottom: row, right: col }
        };
        let header_row_count: u32 = root.attribute("headerRowCount")
            .and_then(|s| s.parse().ok())
            .unwrap_or(1);
        let totals_row_count: u32 = root.attribute("totalsRowCount")
            .and_then(|s| s.parse().ok())
            .unwrap_or(0);
        let style_info = root.children().find(|n| n.tag_name().name() == "tableStyleInfo");
        let style_name = style_info
            .and_then(|n| n.attribute("name"))
            .unwrap_or("TableStyleMedium2")
            .to_string();
        let bool_attr = |n: &roxmltree::Node, key: &str| n.attribute(key).map(|v| v == "1" || v == "true").unwrap_or(false);
        let (show_row_stripes, show_column_stripes, show_first_column, show_last_column) = match style_info {
            Some(n) => (
                bool_attr(&n, "showRowStripes"),
                bool_attr(&n, "showColumnStripes"),
                bool_attr(&n, "showFirstColumn"),
                bool_attr(&n, "showLastColumn"),
            ),
            None => (false, false, false, false),
        };
        let accent_color = resolve_table_style_accent(&style_name, theme_colors);
        let (whole_table_dxf, header_row_dxf) = match custom_styles.get(&style_name) {
            Some(e) => (e.whole_table, e.header_row),
            None => (None, None),
        };
        tables.push(TableInfo {
            range,
            style_name,
            header_row_count,
            totals_row_count,
            show_row_stripes,
            show_column_stripes,
            show_first_column,
            show_last_column,
            accent_color,
            whole_table_dxf,
            header_row_dxf,
        });
    }
    tables
}

/// Resolve hyperlink rIds to URLs from the sheet rels file.
fn load_hyperlinks(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    hyperlink_rids: Vec<(u32, u32, String)>,
) -> Vec<Hyperlink> {
    if hyperlink_rids.is_empty() { return Vec::new(); }
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let rels = read_zip_entry(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();
    hyperlink_rids.into_iter().map(|(col, row, rid)| Hyperlink {
        col, row, url: rels.get(&rid).cloned(),
    }).collect()
}

/// Read a binary file from the zip.
fn read_zip_bytes(archive: &mut zip::ZipArchive<Cursor<&[u8]>>, path: &str) -> Option<Vec<u8>> {
    let mut file = archive.by_name(path).ok()?;
    if file.size() > MAX_ZIP_ENTRY_BYTES {
        return None;
    }
    let mut buf = Vec::new();
    file.by_ref().take(MAX_ZIP_ENTRY_BYTES).read_to_end(&mut buf).ok()?;
    Some(buf)
}

/// Resolve a relative path ("../media/image1.png") against a base dir ("xl/drawings").
fn resolve_zip_path(base_dir: &str, target: &str) -> String {
    let mut parts: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for seg in target.split('/') {
        match seg {
            ".." => { parts.pop(); }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

fn mime_from_ext(path: &str) -> &'static str {
    match path.rsplit('.').next().unwrap_or("").to_ascii_lowercase().as_str() {
        "png"  => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif"  => "image/gif",
        "bmp"  => "image/bmp",
        "webp" => "image/webp",
        _      => "application/octet-stream",
    }
}

/// Parse `<xdr:twoCellAnchor>` elements from a drawing XML and resolve
/// embedded pictures into data URLs. `drawing_dir` is the folder that
/// contains `drawing_path` so relative `Target`s resolve correctly.
fn parse_drawing_anchors(
    drawing_xml: &str,
    drawing_rels: &HashMap<String, String>,
    drawing_dir: &str,
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> Vec<ImageAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let mut anchors: Vec<ImageAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns)
        {
            continue;
        }
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
        let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
        let mut pic_rid: Option<String> = None;

        for child in anchor.children() {
            if !child.is_element() { continue; }
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0;
                    let mut col_off: i64 = 0;
                    let mut row: u32 = 0;
                    let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from {
                        from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off;
                    } else {
                        to_col = col; to_col_off = col_off; to_row = row; to_row_off = row_off;
                    }
                }
                "pic" => {
                    // <xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>
                    let blip_fill = child.children()
                        .find(|n| n.tag_name().name() == "blipFill" && n.tag_name().namespace() == Some(xdr_ns));
                    if let Some(bf) = blip_fill {
                        let blip = bf.children()
                            .find(|n| n.tag_name().name() == "blip" && n.tag_name().namespace() == Some(a_ns));
                        if let Some(b) = blip {
                            // r:embed attribute
                            pic_rid = b.attributes()
                                .find(|a| a.name() == "embed" && a.namespace() == Some(r_ns))
                                .map(|a| a.value().to_string());
                        }
                    }
                }
                _ => {}
            }
        }

        let Some(rid) = pic_rid else { continue; };
        let Some(target) = drawing_rels.get(&rid) else { continue; };
        let media_path = resolve_zip_path(drawing_dir, target);
        let Some(bytes) = read_zip_bytes(archive, &media_path) else { continue; };
        let mime = mime_from_ext(&media_path);
        let data_url = format!("data:{mime};base64,{}", B64.encode(&bytes));

        anchors.push(ImageAnchor {
            from_col, from_col_off, from_row, from_row_off,
            to_col, to_col_off, to_row, to_row_off,
            data_url,
        });
    }
    anchors
}

// ─── Shape group parsing ────────────────────────────────────────────────────
//
// ECMA-376 §20.5.2.17 `<xdr:grpSp>` / §20.1.9 DrawingML shapes. Each
// top-level grpSp inside a twoCellAnchor has its own coordinate system:
//   - grpSpPr/xfrm/off,ext     : group's position/size in parent coords
//   - grpSpPr/xfrm/chOff,chExt : origin/extent of the group's child coords
//
// A child sp at child coord (cx, cy) maps to parent coord:
//   parent.x = off.x + (cx - chOff.x) / chExt.cx * ext.cx
//
// For rendering, we chain these transforms down to the top-level grpSp and
// then normalize each leaf shape's rect into [0,1] of the top-level ext.

#[derive(Clone, Copy)]
struct Xfrm {
    off_x: f64, off_y: f64,
    ext_x: f64, ext_y: f64,
    ch_off_x: f64, ch_off_y: f64,
    ch_ext_x: f64, ch_ext_y: f64,
    has_ch: bool,
}

fn parse_xfrm(xfrm_node: &roxmltree::Node) -> Option<Xfrm> {
    let mut off = (0.0_f64, 0.0_f64);
    let mut ext = (0.0_f64, 0.0_f64);
    let mut ch_off = (0.0_f64, 0.0_f64);
    let mut ch_ext = (0.0_f64, 0.0_f64);
    let mut has_ext = false;
    let mut has_ch = false;
    for c in xfrm_node.children() {
        match c.tag_name().name() {
            "off" => {
                off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
            }
            "ext" => {
                ext.0 = c.attribute("cx").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ext.1 = c.attribute("cy").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ext = true;
            }
            "chOff" => {
                ch_off.0 = c.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_off.1 = c.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            "chExt" => {
                ch_ext.0 = c.attribute("cx").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                ch_ext.1 = c.attribute("cy").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                has_ch = true;
            }
            _ => {}
        }
    }
    if !has_ext { return None; }
    Some(Xfrm {
        off_x: off.0, off_y: off.1,
        ext_x: ext.0, ext_y: ext.1,
        ch_off_x: ch_off.0, ch_off_y: ch_off.1,
        ch_ext_x: if ch_ext.0 == 0.0 { ext.0 } else { ch_ext.0 },
        ch_ext_y: if ch_ext.1 == 0.0 { ext.1 } else { ch_ext.1 },
        has_ch,
    })
}

fn parse_solid_fill(fill_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    for c in fill_node.children() {
        match c.tag_name().name() {
            "srgbClr" => {
                let v = c.attribute("val")?;
                return Some(format!("#{}", v.to_uppercase()));
            }
            "schemeClr" => {
                let v = c.attribute("val")?;
                // `theme_colors` is collected in OOXML clrScheme document
                // order: dk1, lt1, dk2, lt2, accent1..accent6, hlink,
                // folHlink. See `parse_theme_colors`. The earlier mapping
                // here had dk1/lt1 and dk2/lt2 swapped which darkened
                // shapes that painted "lt1" (the sheet paper colour).
                let idx = match v {
                    "dk1" | "tx1"    => Some(0),
                    "lt1" | "bg1"    => Some(1),
                    "dk2" | "tx2"    => Some(2),
                    "lt2" | "bg2"    => Some(3),
                    "accent1"        => Some(4),
                    "accent2"        => Some(5),
                    "accent3"        => Some(6),
                    "accent4"        => Some(7),
                    "accent5"        => Some(8),
                    "accent6"        => Some(9),
                    "hlink"          => Some(10),
                    "folHlink"       => Some(11),
                    _ => None,
                };
                return idx.and_then(|i| theme_colors.get(i).cloned());
            }
            _ => {}
        }
    }
    None
}

/// Parse a single custGeom path element. Each path has its own coordinate
/// system (`a:path/@w`, `@h`) that the renderer scales to the shape's rect.
fn parse_custom_path(path_node: &roxmltree::Node) -> PathInfo {
    let w: f64 = path_node.attribute("w").and_then(|s| s.parse().ok()).unwrap_or(0.0);
    let h: f64 = path_node.attribute("h").and_then(|s| s.parse().ok()).unwrap_or(0.0);
    let mut commands: Vec<PathCmd> = Vec::new();
    for cmd in path_node.children().filter(|n| n.is_element()) {
        let name = cmd.tag_name().name();
        // Collect `<a:pt x=.. y=..>` points in order.
        let pts: Vec<(f64, f64)> = cmd.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "pt")
            .map(|n| (
                n.attribute("x").and_then(|s| s.parse().ok()).unwrap_or(0.0),
                n.attribute("y").and_then(|s| s.parse().ok()).unwrap_or(0.0),
            ))
            .collect();
        match name {
            "moveTo"       => if let Some(p) = pts.first() { commands.push(PathCmd::MoveTo { x: p.0, y: p.1 }); },
            "lnTo"         => if let Some(p) = pts.first() { commands.push(PathCmd::LineTo { x: p.0, y: p.1 }); },
            "cubicBezTo"   => if pts.len() >= 3 {
                commands.push(PathCmd::CubicBezTo {
                    x1: pts[0].0, y1: pts[0].1,
                    x2: pts[1].0, y2: pts[1].1,
                    x3: pts[2].0, y3: pts[2].1,
                });
            },
            "quadBezTo"    => if pts.len() >= 2 {
                commands.push(PathCmd::QuadBezTo {
                    x1: pts[0].0, y1: pts[0].1,
                    x2: pts[1].0, y2: pts[1].1,
                });
            },
            "close"        => commands.push(PathCmd::Close),
            "arcTo" => {
                // ECMA-376 §20.1.9.3: `wR`/`hR` in path-coord units;
                // `stAng`/`swAng` in 60000ths of a degree.
                let wr:     f64 = cmd.attribute("wR").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let hr:     f64 = cmd.attribute("hR").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let st_ang: f64 = cmd.attribute("stAng").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                let sw_ang: f64 = cmd.attribute("swAng").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                commands.push(PathCmd::ArcTo { wr, hr, st_ang, sw_ang });
            }
            _ => {}
        }
    }
    PathInfo { w, h, commands }
}

fn parse_sp_geom(sp_pr: &roxmltree::Node) -> Option<ShapeGeom> {
    for c in sp_pr.children().filter(|n| n.is_element()) {
        match c.tag_name().name() {
            "prstGeom" => {
                return Some(ShapeGeom::Preset {
                    name: c.attribute("prst").unwrap_or("rect").to_string(),
                });
            }
            "custGeom" => {
                let mut paths: Vec<PathInfo> = Vec::new();
                for pl in c.children().filter(|n| n.is_element() && n.tag_name().name() == "pathLst") {
                    for p in pl.children().filter(|n| n.is_element() && n.tag_name().name() == "path") {
                        paths.push(parse_custom_path(&p));
                    }
                }
                return Some(ShapeGeom::Custom { paths });
            }
            _ => {}
        }
    }
    None
}

/// Recursively walk an `xdr:grpSp` / `xdr:sp` tree, chaining coordinate
/// transforms, and push leaf shapes (normalized to [0,1] of `root_ext`) into
/// `out`.
fn collect_shapes(
    node: &roxmltree::Node,
    root_off_x: f64, root_off_y: f64,
    root_ext_x: f64, root_ext_y: f64,
    // transform from current local coords into root (top-level grpSp) coords
    scale_x: f64, scale_y: f64,
    trans_x: f64, trans_y: f64,
    theme_colors: &[String],
    rid_urls: &HashMap<String, String>,
    out: &mut Vec<ShapeInfo>,
) {
    for child in node.children().filter(|n| n.is_element()) {
        let tag = child.tag_name().name();
        if tag == "grpSp" {
            // Nested grpSp: compose the transform by the group's own xfrm.
            let grp_sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
            let xfrm = grp_sp_pr
                .and_then(|n| n.children().find(|c| c.is_element() && c.tag_name().name() == "xfrm"))
                .as_ref()
                .and_then(parse_xfrm);
            let (sx, sy, tx, ty) = if let Some(x) = xfrm {
                if x.has_ch && x.ch_ext_x != 0.0 && x.ch_ext_y != 0.0 {
                    let csx = x.ext_x / x.ch_ext_x;
                    let csy = x.ext_y / x.ch_ext_y;
                    // Child point (cx, cy) → (x.off_x + (cx - x.ch_off_x)*csx) in parent coords,
                    // then apply outer (scale/trans) to reach root coords.
                    (
                        scale_x * csx,
                        scale_y * csy,
                        trans_x + scale_x * (x.off_x - x.ch_off_x * csx),
                        trans_y + scale_y * (x.off_y - x.ch_off_y * csy),
                    )
                } else {
                    // No child coord system: treat as identity mapping inside the group.
                    (scale_x, scale_y,
                     trans_x + scale_x * x.off_x,
                     trans_y + scale_y * x.off_y)
                }
            } else {
                (scale_x, scale_y, trans_x, trans_y)
            };
            collect_shapes(&child, root_off_x, root_off_y, root_ext_x, root_ext_y,
                           sx, sy, tx, ty, theme_colors, rid_urls, out);
        } else if tag == "sp" {
            let sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else { continue; };
            let xfrm_node = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            let rot_raw: f64 = xfrm_n.attribute("rot")
                .and_then(|s| s.parse().ok()).unwrap_or(0.0);

            // Shape rect in root coords
            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;

            // Normalize to [0,1] of root ext
            if root_ext_x == 0.0 || root_ext_y == 0.0 { continue; }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;

            let geom = parse_sp_geom(&sp_pr);
            let Some(geom) = geom else { continue; };

            // Fill
            let mut fill_color: Option<String> = None;
            let mut has_no_fill = false;
            for c in sp_pr.children().filter(|n| n.is_element()) {
                match c.tag_name().name() {
                    "solidFill" => { fill_color = parse_solid_fill(&c, theme_colors); }
                    "noFill"    => { has_no_fill = true; }
                    _ => {}
                }
            }
            if has_no_fill { fill_color = None; }

            // Stroke (line)
            let mut stroke_color: Option<String> = None;
            let mut stroke_width: i64 = 0;
            if let Some(ln) = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "ln") {
                stroke_width = ln.attribute("w").and_then(|s| s.parse().ok()).unwrap_or(0);
                for c in ln.children().filter(|n| n.is_element()) {
                    if c.tag_name().name() == "solidFill" {
                        stroke_color = parse_solid_fill(&c, theme_colors);
                    } else if c.tag_name().name() == "noFill" {
                        stroke_color = None;
                        stroke_width = 0;
                    }
                }
            }

            out.push(ShapeInfo {
                x: nx, y: ny, w: nw, h: nh,
                rot: rot_raw / 60000.0,
                fill_color,
                stroke_color,
                stroke_width,
                geom,
            });
        } else if tag == "pic" {
            // `<xdr:pic>` leaf inside a group (ECMA-376 §20.5.2.17). The image
            // binary is resolved via the drawing's .rels file; `rid_urls` maps
            // each r:id to its pre-encoded `data:<mime>;base64,…` URL.
            let sp_pr = child.children().find(|n| n.is_element() && n.tag_name().name() == "spPr");
            let Some(sp_pr) = sp_pr else { continue; };
            let xfrm_node = sp_pr.children().find(|n| n.is_element() && n.tag_name().name() == "xfrm");
            let Some(xfrm_n) = xfrm_node else { continue; };
            let Some(xfrm) = parse_xfrm(&xfrm_n) else { continue; };
            let rot_raw: f64 = xfrm_n.attribute("rot")
                .and_then(|s| s.parse().ok()).unwrap_or(0.0);

            // Resolve <a:blip r:embed="rIdN"/>. The r:embed attribute lives in
            // the relationships namespace, not the drawingml namespace.
            let r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            let pic_rid = child.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "blip")
                .and_then(|b| {
                    b.attributes()
                        .find(|a| a.name() == "embed" && a.namespace() == Some(r_ns))
                        .map(|a| a.value().to_string())
                });
            let Some(rid) = pic_rid else { continue; };
            let Some(data_url) = rid_urls.get(&rid) else { continue; };

            let root_x = trans_x + scale_x * xfrm.off_x;
            let root_y = trans_y + scale_y * xfrm.off_y;
            let root_w = scale_x * xfrm.ext_x;
            let root_h = scale_y * xfrm.ext_y;
            if root_ext_x == 0.0 || root_ext_y == 0.0 { continue; }
            let nx = (root_x - root_off_x) / root_ext_x;
            let ny = (root_y - root_off_y) / root_ext_y;
            let nw = root_w / root_ext_x;
            let nh = root_h / root_ext_y;
            if nw <= 0.0 || nh <= 0.0 { continue; }

            out.push(ShapeInfo {
                x: nx, y: ny, w: nw, h: nh,
                rot: rot_raw / 60000.0,
                fill_color: None,
                stroke_color: None,
                stroke_width: 0,
                geom: ShapeGeom::Image { data_url: data_url.clone() },
            });
        }
        // Ignore `xdr:cxnSp` / text-only elements for this minimal pass.
    }
}

fn parse_shape_anchors(
    drawing_xml: &str,
    theme_colors: &[String],
    rid_urls: &HashMap<String, String>,
) -> Vec<ShapeAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else { return Vec::new(); };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let mut anchors: Vec<ShapeAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns) { continue; }
        // Only anchors whose top-level child is `<xdr:grpSp>`.
        let grp = anchor.children().find(|n| n.is_element() && n.tag_name().name() == "grpSp");
        let Some(grp) = grp else { continue; };
        let grp_sp_pr = grp.children().find(|n| n.is_element() && n.tag_name().name() == "grpSpPr");
        let xfrm = grp_sp_pr
            .and_then(|n| n.children().find(|c| c.is_element() && c.tag_name().name() == "xfrm"))
            .as_ref()
            .and_then(parse_xfrm);
        let Some(root) = xfrm else { continue; };
        if !root.has_ch || root.ch_ext_x == 0.0 || root.ch_ext_y == 0.0 { continue; }

        // Parse from/to anchor rect
        let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
        let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
        for c in anchor.children() {
            if !c.is_element() { continue; }
            if c.tag_name().name() == "from" || c.tag_name().name() == "to" {
                let is_from = c.tag_name().name() == "from";
                let mut col: u32 = 0; let mut col_off: i64 = 0;
                let mut row: u32 = 0; let mut row_off: i64 = 0;
                for cc in c.children() {
                    match (cc.tag_name().name(), cc.text()) {
                        ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                        ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                        ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                        ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                        _ => {}
                    }
                }
                if is_from {
                    from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off;
                } else {
                    to_col = col; to_col_off = col_off; to_row = row; to_row_off = row_off;
                }
            }
        }

        // Map child coords → root coords with the grpSp's own chOff/chExt.
        let csx = root.ext_x / root.ch_ext_x;
        let csy = root.ext_y / root.ch_ext_y;
        let tx = root.off_x - root.ch_off_x * csx;
        let ty = root.off_y - root.ch_off_y * csy;

        let mut shapes: Vec<ShapeInfo> = Vec::new();
        collect_shapes(&grp, root.off_x, root.off_y, root.ext_x, root.ext_y,
                       csx, csy, tx, ty, theme_colors, rid_urls, &mut shapes);
        if shapes.is_empty() { continue; }

        anchors.push(ShapeAnchor {
            from_col, from_col_off, from_row, from_row_off,
            to_col, to_col_off, to_row, to_row_off,
            shapes,
        });
    }
    anchors
}

fn load_sheet_shape_groups(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<ShapeAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else { return Vec::new(); };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else { return Vec::new(); };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else { return Vec::new(); };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") { drawing_targets.push(t.to_string()); }
        }
    }
    let mut all: Vec<ShapeAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        let rid_urls = build_drawing_rid_urls(archive, &drawing_path);
        all.extend(parse_shape_anchors(&drawing_xml, theme_colors, &rid_urls));
    }
    all
}

/// Build a `HashMap<rId, data-URL>` for every image (png/jpg/…) target in
/// a drawing's `.rels` file. Used by `collect_shapes` to resolve `<xdr:pic>`
/// leaves inside a group. Mirrors the logic in `parse_drawing_anchors` but
/// eagerly encodes each referenced image so per-shape lookup is a single
/// HashMap hit.
fn build_drawing_rid_urls(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    drawing_path: &str,
) -> HashMap<String, String> {
    let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else {
        return HashMap::new();
    };
    let rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
    let rels = read_zip_entry(archive, &rels_path)
        .ok()
        .map(|xml| parse_rels_map(&xml))
        .unwrap_or_default();

    let mut result: HashMap<String, String> = HashMap::new();
    for (rid, target) in rels {
        let lower = target.to_lowercase();
        if !(lower.ends_with(".png") || lower.ends_with(".jpg")
            || lower.ends_with(".jpeg") || lower.ends_with(".gif")
            || lower.ends_with(".bmp")  || lower.ends_with(".webp"))
        {
            continue;
        }
        let media_path = resolve_zip_path(drawing_dir, &target);
        if let Some(bytes) = read_zip_bytes(archive, &media_path) {
            let mime = mime_from_ext(&media_path);
            result.insert(rid, format!("data:{mime};base64,{}", B64.encode(&bytes)));
        }
    }
    result
}

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s), and return all image anchors found.
fn load_sheet_images(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
) -> Vec<ImageAnchor> {
    // sheet rels path:  xl/worksheets/_rels/sheet1.xml.rels
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else {
        return Vec::new();
    };

    // Find all drawing relationships
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        if rel_type.ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() { return Vec::new(); }

    let mut all_anchors: Vec<ImageAnchor> = Vec::new();
    for target in drawing_targets {
        // sheet_dir is "worksheets", target typically "../drawings/drawing1.xml"
        // base dir for the drawing = "xl/worksheets" + "../drawings" → "xl/drawings"
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        // Drawing rels:  xl/drawings/_rels/drawing1.xml.rels
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else { continue; };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_entry(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        let mut anchors = parse_drawing_anchors(&drawing_xml, &drawing_rels, drawing_dir, archive);
        all_anchors.append(&mut anchors);
    }
    all_anchors
}

// ─── Slicer loading ─────────────────────────────────────────────────────────
//
// Office 2010+ extension (`sle:slicer` inside `<mc:AlternateContent>`).
// Resolving one slicer graphicFrame into a drawable anchor takes four
// XML files:
//   1. The sheet's drawing (for the anchor rect + graphicFrame name).
//   2. `xl/slicers/slicerN.xml` — slicer definition: graphicFrame name →
//      caption + cache name.
//   3. `xl/slicerCaches/slicerCacheN.xml` — cache definition: cache name →
//      source field + list of (item index, selected?).
//   4. `xl/pivotCache/pivotCacheDefinitionN.xml` — pivot cache: field name →
//      ordered string values.
// Excel also allows slicers bound to Excel Tables (`tableSlicerCache`), but
// the present sample is pivot-only; we only implement the pivot path.

#[derive(Default)]
struct SlicerCacheInfo {
    source_name: String,
    items: Vec<(u32, bool)>, // (index into pivot field, selected)
}

#[derive(Default)]
struct PivotCacheFields {
    by_name: HashMap<String, Vec<String>>, // field name → ordered string items
}

/// Parse every `xl/pivotCache/pivotCacheDefinition*.xml` and merge its
/// cacheFields (indexed by `@name`) into a single map. Sample workbooks
/// typically have one pivotCache but the loop keeps the code general.
fn load_all_pivot_cache_fields(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> PivotCacheFields {
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut out = PivotCacheFields::default();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/pivotCache/pivotCacheDefinition") && n.ends_with(".xml"))
        .collect();
    for name in names {
        let Ok(xml) = read_zip_entry(archive, &name) else { continue; };
        let Ok(doc) = roxmltree::Document::parse(&xml) else { continue; };
        for field in doc.descendants() {
            if field.tag_name().name() != "cacheField"
                || field.tag_name().namespace() != Some(ns)
            { continue; }
            let Some(field_name) = field.attribute("name") else { continue; };
            let mut items: Vec<String> = Vec::new();
            for shared in field.children().filter(|n| n.is_element() && n.tag_name().name() == "sharedItems") {
                for item in shared.children().filter(|n| n.is_element()) {
                    match item.tag_name().name() {
                        "s" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "n" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "d" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "b" => items.push(item.attribute("v").unwrap_or("").to_string()),
                        "m" => items.push(String::new()),
                        _ => {}
                    }
                }
            }
            if !items.is_empty() {
                out.by_name.insert(field_name.to_string(), items);
            }
        }
    }
    out
}

/// Parse every `xl/slicerCaches/slicerCache*.xml` and build a map keyed by
/// the slicerCache's `@name` attribute (e.g. `"スライサー_贈答相手1"`). That
/// name is what `<slicer cache="…"/>` in `xl/slicers/slicerN.xml` references.
fn load_all_slicer_caches(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> HashMap<String, SlicerCacheInfo> {
    let mut out: HashMap<String, SlicerCacheInfo> = HashMap::new();
    let names: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .filter(|n| n.starts_with("xl/slicerCaches/slicerCache") && n.ends_with(".xml"))
        .collect();
    for path in names {
        let Ok(xml) = read_zip_entry(archive, &path) else { continue; };
        let Ok(doc) = roxmltree::Document::parse(&xml) else { continue; };
        let root = doc.root_element();
        let cache_name = root.attribute("name").unwrap_or("").to_string();
        let source_name = root.attribute("sourceName").unwrap_or("").to_string();
        let mut items: Vec<(u32, bool)> = Vec::new();
        for tabular in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == "tabular") {
            for i_el in tabular.descendants().filter(|n| n.is_element() && n.tag_name().name() == "i") {
                let x: u32 = i_el.attribute("x").and_then(|v| v.parse().ok()).unwrap_or(0);
                // `s` defaults to "1" (selected) when absent — ECMA-376
                // extension schema for slicer caches.
                let selected = i_el.attribute("s").map(|v| v != "0").unwrap_or(true);
                items.push((x, selected));
            }
        }
        if !cache_name.is_empty() {
            out.insert(cache_name, SlicerCacheInfo { source_name, items });
        }
    }
    out
}

/// Slicer definition (`xl/slicers/slicerN.xml`): maps each graphicFrame name
/// on the sheet to its display caption and the slicerCache it's backed by.
#[derive(Default)]
struct SlicerDef {
    caption: String,
    cache: String,
}

fn parse_slicers_xml(xml: &str) -> HashMap<String, SlicerDef> {
    let mut out: HashMap<String, SlicerDef> = HashMap::new();
    let Ok(doc) = roxmltree::Document::parse(xml) else { return out; };
    for slicer in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == "slicer") {
        let name = slicer.attribute("name").unwrap_or("").to_string();
        let caption = slicer.attribute("caption").unwrap_or("").to_string();
        let cache = slicer.attribute("cache").unwrap_or("").to_string();
        if !name.is_empty() {
            out.insert(name, SlicerDef { caption, cache });
        }
    }
    out
}

fn load_sheet_slicers(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str, // e.g. "worksheets/sheet1.xml"
) -> Vec<SlicerAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };

    // 1. Collect slicer-definition and drawing targets from the sheet rels.
    let mut drawing_targets: Vec<String> = Vec::new();
    let mut slicer_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let rel_type = rel.attribute("Type").unwrap_or("");
        let Some(target) = rel.attribute("Target") else { continue; };
        if rel_type.ends_with("/drawing") {
            drawing_targets.push(target.to_string());
        } else if rel_type.ends_with("/slicer") {
            slicer_targets.push(target.to_string());
        }
    }
    if drawing_targets.is_empty() || slicer_targets.is_empty() {
        return Vec::new();
    }

    // 2. Parse all slicer definitions referenced by this sheet, keyed by
    //    graphicFrame name.
    let mut slicer_defs: HashMap<String, SlicerDef> = HashMap::new();
    for target in &slicer_targets {
        let slicer_path = resolve_zip_path(&format!("xl/{}", sheet_dir), target);
        let Ok(xml) = read_zip_entry(archive, &slicer_path) else { continue; };
        for (k, v) in parse_slicers_xml(&xml) {
            slicer_defs.insert(k, v);
        }
    }
    if slicer_defs.is_empty() { return Vec::new(); }

    // 3. Resolve caches (and their backing pivot fields) once.
    let slicer_caches = load_all_slicer_caches(archive);
    let pivot_fields = load_all_pivot_cache_fields(archive);

    // 4. Walk each drawing and pick up slicer graphicFrames.
    let mut out: Vec<SlicerAnchor> = Vec::new();
    for target in drawing_targets {
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        out.extend(parse_slicer_anchors(&drawing_xml, &slicer_defs, &slicer_caches, &pivot_fields));
    }
    out
}

fn parse_slicer_anchors(
    drawing_xml: &str,
    slicer_defs: &HashMap<String, SlicerDef>,
    slicer_caches: &HashMap<String, SlicerCacheInfo>,
    pivot_fields: &PivotCacheFields,
) -> Vec<SlicerAnchor> {
    let Ok(doc) = roxmltree::Document::parse(drawing_xml) else {
        return Vec::new();
    };
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let slicer_uri = "http://schemas.microsoft.com/office/drawing/2010/slicer";
    let mut out: Vec<SlicerAnchor> = Vec::new();

    for anchor in doc.descendants() {
        if anchor.tag_name().name() != "twoCellAnchor"
            || anchor.tag_name().namespace() != Some(xdr_ns)
        { continue; }

        // Anchor rect.
        let mut from = (0u32, 0i64, 0u32, 0i64);
        let mut to   = (0u32, 0i64, 0u32, 0i64);
        for child in anchor.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "from" | "to" => {
                    let is_from = child.tag_name().name() == "from";
                    let mut col: u32 = 0; let mut col_off: i64 = 0;
                    let mut row: u32 = 0; let mut row_off: i64 = 0;
                    for c in child.children() {
                        match (c.tag_name().name(), c.text()) {
                            ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                            ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                            ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                            ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    if is_from { from = (col, col_off, row, row_off); }
                    else       { to   = (col, col_off, row, row_off); }
                }
                _ => {}
            }
        }

        // Slicers live inside `<mc:AlternateContent><mc:Choice>` — descend
        // until we find a `<xdr:graphicFrame>` whose graphicData uri is the
        // 2010 slicer namespace, then harvest the graphicFrame's cNvPr name.
        let Some(frame_name) = anchor.descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "Choice" && n.tag_name().namespace() == Some(mc_ns))
            .flat_map(|choice| choice.descendants())
            .find_map(|n| {
                if n.is_element()
                    && n.tag_name().name() == "graphicData"
                    && n.attribute("uri") == Some(slicer_uri)
                {
                    // graphicData → ancestor graphicFrame → nvGraphicFramePr → cNvPr
                    let mut p = n.parent();
                    while let Some(pp) = p {
                        if pp.tag_name().name() == "graphicFrame" { break; }
                        p = pp.parent();
                    }
                    let frame = p?;
                    let cnvpr = frame.descendants()
                        .find(|d| d.is_element() && d.tag_name().name() == "cNvPr")?;
                    cnvpr.attribute("name").map(|s| s.to_string())
                } else { None }
            }) else { continue };

        let Some(slicer_def) = slicer_defs.get(&frame_name) else { continue; };

        // Resolve items via cache → pivot field; fall back to an empty list
        // if any link is broken (still renders the header and box).
        let items: Vec<SlicerItem> = slicer_caches.get(&slicer_def.cache)
            .map(|cache| {
                let field_items = pivot_fields.by_name.get(&cache.source_name);
                cache.items.iter().map(|(x, selected)| {
                    let name = field_items
                        .and_then(|list| list.get(*x as usize))
                        .cloned()
                        .unwrap_or_default();
                    SlicerItem { name, selected: *selected }
                }).collect()
            })
            .unwrap_or_default();

        let caption = if !slicer_def.caption.is_empty() {
            slicer_def.caption.clone()
        } else {
            frame_name.clone()
        };

        out.push(SlicerAnchor {
            from_col: from.0, from_col_off: from.1, from_row: from.2, from_row_off: from.3,
            to_col:   to.0,   to_col_off:   to.1,   to_row:   to.2,   to_row_off:   to.3,
            caption,
            items,
        });
    }
    out
}

// ─── Chart loading ──────────────────────────────────────────────────────────

/// Given a sheet path (e.g. "worksheets/sheet1.xml"), locate and parse
/// its drawing(s) for chart anchors (`<xdr:graphicFrame>` elements).
fn load_sheet_charts(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_path: &str,
    theme_colors: &[String],
) -> Vec<ChartAnchor> {
    let Some((sheet_dir, sheet_file)) = sheet_path.rsplit_once('/') else {
        return Vec::new();
    };
    let sheet_rels_path = format!("xl/{}/_rels/{}.rels", sheet_dir, sheet_file);
    let Ok(sheet_rels_xml) = read_zip_entry(archive, &sheet_rels_path) else {
        return Vec::new();
    };
    let Ok(rels_doc) = roxmltree::Document::parse(&sheet_rels_xml) else {
        return Vec::new();
    };

    // Collect all drawing relationship targets
    let mut drawing_targets: Vec<String> = Vec::new();
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        if rel.attribute("Type").unwrap_or("").ends_with("/drawing") {
            if let Some(t) = rel.attribute("Target") {
                drawing_targets.push(t.to_string());
            }
        }
    }
    if drawing_targets.is_empty() { return Vec::new(); }

    let mut all_charts: Vec<ChartAnchor> = Vec::new();
    let xdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let a_ns   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    let c_ns   = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    let r_ns   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    for target in drawing_targets {
        // Resolve drawing path relative to the sheet directory
        let drawing_path = resolve_zip_path(&format!("xl/{}", sheet_dir), &target);
        let Ok(drawing_xml) = read_zip_entry(archive, &drawing_path) else { continue; };
        let Ok(draw_doc) = roxmltree::Document::parse(&drawing_xml) else { continue; };

        // Load drawing rels (to resolve chart rIds)
        let Some((drawing_dir, drawing_file)) = drawing_path.rsplit_once('/') else { continue; };
        let drawing_rels_path = format!("{}/_rels/{}.rels", drawing_dir, drawing_file);
        let drawing_rels = read_zip_entry(archive, &drawing_rels_path)
            .ok()
            .map(|xml| parse_rels_map(&xml))
            .unwrap_or_default();

        // Iterate over twoCellAnchor elements
        for anchor in draw_doc.root_element().children().filter(|n| n.is_element()) {
            if anchor.tag_name().name() != "twoCellAnchor"
                || anchor.tag_name().namespace() != Some(xdr_ns)
            {
                continue;
            }

            let (mut from_col, mut from_col_off, mut from_row, mut from_row_off) = (0u32, 0i64, 0u32, 0i64);
            let (mut to_col,   mut to_col_off,   mut to_row,   mut to_row_off)   = (0u32, 0i64, 0u32, 0i64);
            let mut chart_rid: Option<String> = None;

            for child in anchor.children() {
                if !child.is_element() { continue; }
                match child.tag_name().name() {
                    "from" | "to" => {
                        let is_from = child.tag_name().name() == "from";
                        let mut col: u32 = 0; let mut col_off: i64 = 0;
                        let mut row: u32 = 0; let mut row_off: i64 = 0;
                        for c in child.children() {
                            match (c.tag_name().name(), c.text()) {
                                ("col",    Some(t)) => col     = t.trim().parse().unwrap_or(0),
                                ("colOff", Some(t)) => col_off = t.trim().parse().unwrap_or(0),
                                ("row",    Some(t)) => row     = t.trim().parse().unwrap_or(0),
                                ("rowOff", Some(t)) => row_off = t.trim().parse().unwrap_or(0),
                                _ => {}
                            }
                        }
                        if is_from { from_col = col; from_col_off = col_off; from_row = row; from_row_off = row_off; }
                        else       { to_col   = col; to_col_off   = col_off; to_row   = row; to_row_off   = row_off; }
                    }
                    "graphicFrame" => {
                        // Look for a:graphic/a:graphicData/c:chart[@r:id]
                        for gf_child in child.descendants() {
                            if gf_child.tag_name().name() == "chart"
                                && gf_child.tag_name().namespace() == Some(c_ns)
                            {
                                if let Some(rid) = gf_child.attributes()
                                    .find(|a| a.name() == "id" && a.namespace() == Some(r_ns))
                                    .map(|a| a.value().to_string())
                                {
                                    chart_rid = Some(rid);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }

            let Some(rid) = chart_rid else { continue; };
            let Some(chart_target) = drawing_rels.get(&rid) else { continue; };
            let chart_path = resolve_zip_path(drawing_dir, chart_target);
            let Ok(chart_xml) = read_zip_entry(archive, &chart_path) else { continue; };
            let Some(chart_data) = parse_chart_xml(&chart_xml, c_ns, a_ns, theme_colors) else { continue; };

            all_charts.push(ChartAnchor {
                from_col, from_col_off, from_row, from_row_off,
                to_col,   to_col_off,   to_row,   to_row_off,
                chart: chart_data,
            });
        }
    }
    all_charts
}

// ─── Chart XML parser ────────────────────────────────────────────────────────

/// Parse a `xl/charts/chartN.xml` file into a `ChartData`.
fn parse_chart_xml(xml: &str, c_ns: &str, a_ns: &str, theme_colors: &[String]) -> Option<ChartData> {
    let doc = roxmltree::Document::parse(xml).ok()?;

    // Find c:chart root element
    let chart_root = doc.descendants()
        .find(|n| n.tag_name().name() == "chart" && n.tag_name().namespace() == Some(c_ns))?;

    // Parse optional title
    let title = extract_chart_title(&chart_root, c_ns, a_ns);
    let title_font_size_hpt = extract_chart_title_size(&chart_root, c_ns, a_ns);
    let title_font_color = extract_chart_title_color(&chart_root, c_ns, a_ns);
    let title_font_face = extract_chart_title_face(&chart_root, c_ns, a_ns);
    let title_font_bold = extract_chart_title_bold(&chart_root, c_ns, a_ns);

    // Legend presence: <c:chart><c:legend> is the authoritative signal. Absence
    // means Excel hides the legend (default for a single-series chart with no
    // explicit legend element). When present, `<c:legendPos val>` picks a side
    // ("r"|"l"|"t"|"b"|"tr") — default "r" per ECMA-376 §21.2.2.10.
    let legend_node = chart_root.children()
        .find(|n| n.tag_name().name() == "legend" && n.tag_name().namespace() == Some(c_ns));
    let show_legend = legend_node.is_some();
    let legend_pos = legend_node.and_then(|ln| {
        ln.children()
            .find(|n| n.tag_name().name() == "legendPos" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|p| p.attribute("val").map(|s| s.to_string()))
    });
    // Legend <c:layout><c:manualLayout> (ECMA-376 §21.2.2.31) — when present,
    // gives explicit x/y/w/h fractions of the chart space. Used by the Excel
    // templates that position a top legend into a narrow band, e.g. over the
    // left half of the chart. We just collect the raw fractions here; the
    // renderer decides whether to honor `edge` vs `factor` placement.
    let legend_manual_layout = legend_node.and_then(|ln| {
        let layout = ln.children()
            .find(|n| n.tag_name().name() == "layout" && n.tag_name().namespace() == Some(c_ns))?;
        let manual = layout.children()
            .find(|n| n.tag_name().name() == "manualLayout" && n.tag_name().namespace() == Some(c_ns))?;
        let val = |tag: &str| manual.children()
            .find(|n| n.tag_name().name() == tag && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val").and_then(|v| v.parse::<f64>().ok()));
        let mode = |tag: &str| manual.children()
            .find(|n| n.tag_name().name() == tag && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val").map(|v| v.to_string()))
            .unwrap_or_else(|| "edge".to_string());
        Some(LegendManualLayout {
            x_mode: mode("xMode"),
            y_mode: mode("yMode"),
            x: val("x").unwrap_or(0.0),
            y: val("y").unwrap_or(0.0),
            w: val("w").unwrap_or(0.0),
            h: val("h").unwrap_or(0.0),
        })
    });

    // `<c:chartSpace><c:spPr>` outer fill (ECMA-376 §21.2.2.5). When the
    // element exists and carries `<a:noFill/>` the chart space is
    // transparent — this sample explicitly does that so the underlying
    // gray cell panel shows through. `<a:solidFill>` is resolved against
    // the theme just like series fills. When the element is absent we leave
    // `chart_bg` unset and tell the adapter to use the default opaque white
    // via `has_chart_sp_pr=false`.
    let chart_space_root = doc.descendants()
        .find(|n| n.tag_name().name() == "chartSpace" && n.tag_name().namespace() == Some(c_ns));
    let chart_sp_pr = chart_space_root.and_then(|cs| cs.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns)));
    let has_chart_sp_pr = chart_sp_pr.is_some();
    let chart_bg = chart_sp_pr.and_then(|sp| {
        // Walk direct children: noFill → None, solidFill → resolved color.
        let mut resolved: Option<String> = None;
        for ch in sp.children().filter(|n| n.is_element()) {
            match ch.tag_name().name() {
                "noFill"    => { return None; }
                "solidFill" => { resolved = resolve_fill_color(&ch, theme_colors); break; }
                _ => {}
            }
        }
        resolved
    });

    // `<c:title><c:layout><c:manualLayout>` (ECMA-376 §21.2.2.27).
    let title_manual_layout = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|t| t.children().find(|n| n.tag_name().name() == "layout" && n.tag_name().namespace() == Some(c_ns)))
        .and_then(|l| extract_manual_layout(&l, c_ns));

    // Find c:plotArea
    let plot_area = chart_root.children()
        .find(|n| n.tag_name().name() == "plotArea" && n.tag_name().namespace() == Some(c_ns))?;
    let plot_area_manual_layout = plot_area.children()
        .find(|n| n.tag_name().name() == "layout" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|l| extract_manual_layout(&l, c_ns));

    let mut primary_type = String::new();
    let mut bar_dir      = "col".to_string();
    let mut grouping     = "clustered".to_string();
    // `grouping` is recorded only from the first non-line chart-type element that
    // explicitly sets it. In combo charts (e.g. `<c:barChart grouping="stacked">`
    // followed by `<c:lineChart grouping="standard">`) the lineChart's grouping
    // must not overwrite the bar's, since stacking is a bar/area concept.
    let mut grouping_locked = false;
    let mut all_series: Vec<ChartSeries> = Vec::new();
    let mut shared_categories: Vec<String> = Vec::new();
    let mut show_data_labels = false;
    let mut cat_axis_title: Option<String> = None;
    let mut val_axis_title: Option<String> = None;
    let mut cat_axis_font_size_hpt: Option<i32> = None;
    let mut val_axis_font_size_hpt: Option<i32> = None;
    let mut val_axis_format_code: Option<String> = None;
    // ECMA-376 §21.2.2.40 — `<c:delete val="1"/>` on a `<c:catAx>`/`<c:valAx>`
    // hides the axis (labels, ticks, and lines). Default is "0" (visible).
    let mut cat_axis_hidden = false;
    let mut val_axis_hidden = false;
    let mut cat_axis_format_code: Option<String> = None;
    let mut cat_axis_min: Option<f64> = None;
    let mut cat_axis_max: Option<f64> = None;
    let mut val_axis_min: Option<f64> = None;
    let mut val_axis_max: Option<f64> = None;
    let mut cat_axis_font_bold: Option<bool> = None;
    let mut val_axis_font_bold: Option<bool> = None;
    let mut cat_axis_line_color: Option<String> = None;
    let mut cat_axis_line_width_emu: Option<u32> = None;
    let mut val_axis_line_color: Option<String> = None;
    let mut val_axis_line_width_emu: Option<u32> = None;
    let mut cat_axis_major_tick_mark: Option<String> = None;
    let mut cat_axis_minor_tick_mark: Option<String> = None;
    let mut val_axis_major_tick_mark: Option<String> = None;
    let mut val_axis_minor_tick_mark: Option<String> = None;
    let mut cat_axis_crosses: Option<String> = None;
    let mut cat_axis_crosses_at: Option<f64> = None;
    let mut val_axis_crosses: Option<String> = None;
    let mut val_axis_crosses_at: Option<f64> = None;
    let mut bar_gap_width: Option<i32> = None;
    let mut bar_overlap: Option<i32> = None;
    let mut data_label_position: Option<String> = None;
    let mut data_label_font_color: Option<String> = None;
    let mut data_label_format_code: Option<String> = None;

    // Recognised chart-type element names → our internal type strings
    let type_map: &[(&str, &str)] = &[
        ("barChart",      "bar"),
        ("lineChart",     "line"),
        ("areaChart",     "area"),
        ("pieChart",      "pie"),
        ("doughnutChart", "doughnut"),
        ("radarChart",    "radar"),
        ("scatterChart",  "scatter"),
        ("bubbleChart",   "scatter"), // treat bubble as scatter
    ];

    for child in plot_area.children() {
        if !child.is_element() { continue; }
        if child.tag_name().namespace() != Some(c_ns) { continue; }
        let elem_name = child.tag_name().name();

        // Axis title + tick label font size extraction (ECMA-376 §21.2.2.17
        // c:txPr/a:defRPr@sz gives tick labels their hpt size; absent = default).
        match elem_name {
            "catAx" => {
                if cat_axis_title.is_none() {
                    cat_axis_title = extract_chart_title(&child, c_ns, a_ns);
                }
                if cat_axis_font_size_hpt.is_none() {
                    cat_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                }
                if cat_axis_format_code.is_none() {
                    cat_axis_format_code = extract_axis_format_code(&child, c_ns);
                }
                if cat_axis_font_bold.is_none() {
                    cat_axis_font_bold = extract_axis_tick_label_bold(&child, c_ns);
                }
                if cat_axis_line_color.is_none() || cat_axis_line_width_emu.is_none() {
                    let (col, w) = extract_axis_line_style(&child, c_ns, theme_colors);
                    if cat_axis_line_color.is_none() { cat_axis_line_color = col; }
                    if cat_axis_line_width_emu.is_none() { cat_axis_line_width_emu = w; }
                }
                if cat_axis_major_tick_mark.is_none() {
                    cat_axis_major_tick_mark = extract_axis_tick_mark(&child, c_ns, "majorTickMark");
                }
                if cat_axis_minor_tick_mark.is_none() {
                    cat_axis_minor_tick_mark = extract_axis_tick_mark(&child, c_ns, "minorTickMark");
                }
                if let Some((mn, mx)) = extract_axis_scaling(&child, c_ns) {
                    if cat_axis_min.is_none() { cat_axis_min = mn; }
                    if cat_axis_max.is_none() { cat_axis_max = mx; }
                }
                let (cr, cra) = extract_axis_crosses(&child, c_ns);
                if cat_axis_crosses.is_none() { cat_axis_crosses = cr; }
                if cat_axis_crosses_at.is_none() { cat_axis_crosses_at = cra; }
                if axis_is_deleted(&child, c_ns) { cat_axis_hidden = true; }
                continue;
            }
            "valAx" => {
                // Scatter charts use two `<c:valAx>` (no catAx). Disambiguate
                // by `<c:axPos val>` — `b`(bottom)/`t`(top) → X axis, `l`/`r`
                // → Y axis. For non-scatter charts the first valAx hit is
                // always Y.
                let ax_pos = child.children()
                    .find(|n| n.tag_name().name() == "axPos" && n.tag_name().namespace() == Some(c_ns))
                    .and_then(|n| n.attribute("val"))
                    .unwrap_or("");
                let is_x_axis = matches!(ax_pos, "b" | "t");
                if is_x_axis {
                    if cat_axis_format_code.is_none() {
                        cat_axis_format_code = extract_axis_format_code(&child, c_ns);
                    }
                    if let Some((mn, mx)) = extract_axis_scaling(&child, c_ns) {
                        if cat_axis_min.is_none() { cat_axis_min = mn; }
                        if cat_axis_max.is_none() { cat_axis_max = mx; }
                    }
                    if cat_axis_font_size_hpt.is_none() {
                        cat_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                    }
                    if cat_axis_font_bold.is_none() {
                        cat_axis_font_bold = extract_axis_tick_label_bold(&child, c_ns);
                    }
                    if cat_axis_line_color.is_none() || cat_axis_line_width_emu.is_none() {
                        let (col, w) = extract_axis_line_style(&child, c_ns, theme_colors);
                        if cat_axis_line_color.is_none() { cat_axis_line_color = col; }
                        if cat_axis_line_width_emu.is_none() { cat_axis_line_width_emu = w; }
                    }
                    if cat_axis_major_tick_mark.is_none() {
                        cat_axis_major_tick_mark = extract_axis_tick_mark(&child, c_ns, "majorTickMark");
                    }
                    if cat_axis_minor_tick_mark.is_none() {
                        cat_axis_minor_tick_mark = extract_axis_tick_mark(&child, c_ns, "minorTickMark");
                    }
                    let (cr, cra) = extract_axis_crosses(&child, c_ns);
                    if cat_axis_crosses.is_none() { cat_axis_crosses = cr; }
                    if cat_axis_crosses_at.is_none() { cat_axis_crosses_at = cra; }
                    if axis_is_deleted(&child, c_ns) { cat_axis_hidden = true; }
                } else {
                    if val_axis_title.is_none() {
                        val_axis_title = extract_chart_title(&child, c_ns, a_ns);
                    }
                    if val_axis_font_size_hpt.is_none() {
                        val_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                    }
                    if val_axis_format_code.is_none() {
                        val_axis_format_code = extract_axis_format_code(&child, c_ns);
                    }
                    if val_axis_font_bold.is_none() {
                        val_axis_font_bold = extract_axis_tick_label_bold(&child, c_ns);
                    }
                    if let Some((mn, mx)) = extract_axis_scaling(&child, c_ns) {
                        if val_axis_min.is_none() { val_axis_min = mn; }
                        if val_axis_max.is_none() { val_axis_max = mx; }
                    }
                    let (cr, cra) = extract_axis_crosses(&child, c_ns);
                    if val_axis_crosses.is_none() { val_axis_crosses = cr; }
                    if val_axis_crosses_at.is_none() { val_axis_crosses_at = cra; }
                    if val_axis_line_color.is_none() || val_axis_line_width_emu.is_none() {
                        let (col, w) = extract_axis_line_style(&child, c_ns, theme_colors);
                        if val_axis_line_color.is_none() { val_axis_line_color = col; }
                        if val_axis_line_width_emu.is_none() { val_axis_line_width_emu = w; }
                    }
                    if val_axis_major_tick_mark.is_none() {
                        val_axis_major_tick_mark = extract_axis_tick_mark(&child, c_ns, "majorTickMark");
                    }
                    if val_axis_minor_tick_mark.is_none() {
                        val_axis_minor_tick_mark = extract_axis_tick_mark(&child, c_ns, "minorTickMark");
                    }
                    if axis_is_deleted(&child, c_ns) { val_axis_hidden = true; }
                }
                continue;
            }
            _ => {}
        }

        let ser_type = match type_map.iter().find(|(k, _)| *k == elem_name) {
            Some((_, v)) => *v,
            None => continue,
        };

        if primary_type.is_empty() {
            primary_type = ser_type.to_string();
        }

        // barDir / grouping / dLbls / marker (only meaningful for bar/line/area).
        // <c:marker val> at the chart-type element is the default for all line
        // series in that element (ECMA-376 §21.2.2.33). "1" = markers visible.
        let mut chart_marker_default = false;
        for attr_node in child.children().filter(|n| n.is_element()) {
            match attr_node.tag_name().name() {
                "barDir"   => { bar_dir  = attr_node.attribute("val").unwrap_or("col").to_string(); }
                "grouping" => {
                    let val = attr_node.attribute("val").unwrap_or("clustered").to_string();
                    if !grouping_locked && ser_type != "line" {
                        grouping = val;
                        grouping_locked = true;
                    }
                }
                "marker"   => {
                    chart_marker_default = attr_node.attribute("val").unwrap_or("0") != "0";
                }
                "gapWidth" => {
                    // ECMA-376 §21.2.2.13 — percent of bar width between category
                    // groups (default 150 per spec). Only meaningful on bar charts.
                    if bar_gap_width.is_none() {
                        bar_gap_width = attr_node.attribute("val").and_then(|v| v.parse().ok());
                    }
                }
                "overlap"  => {
                    // ECMA-376 §21.2.2.25 — signed percent of bar width for the
                    // overlap within a cluster (negative = gap).
                    if bar_overlap.is_none() {
                        bar_overlap = attr_node.attribute("val").and_then(|v| v.parse().ok());
                    }
                }
                "dLbls"    => {
                    for d in attr_node.children().filter(|n| n.is_element()) {
                        match d.tag_name().name() {
                            "showVal" | "showPercent" => {
                                if d.attribute("val").unwrap_or("1") != "0" {
                                    show_data_labels = true;
                                }
                            }
                            "dLblPos" => {
                                if data_label_position.is_none() {
                                    data_label_position = d.attribute("val").map(|s| s.to_string());
                                }
                            }
                            "numFmt" => {
                                if data_label_format_code.is_none() {
                                    data_label_format_code = d.attribute("formatCode")
                                        .map(|s| s.to_string())
                                        .filter(|s| !s.is_empty() && s != "General");
                                }
                            }
                            "txPr" => {
                                // Resolve first solidFill under defRPr/rPr for
                                // the data label text color. Common Excel
                                // pattern: <a:solidFill><a:schemeClr val="bg1"/>
                                // → white when bars are dark.
                                if data_label_font_color.is_none() {
                                    for desc in d.descendants().filter(|n| n.is_element()) {
                                        if desc.tag_name().namespace() != Some(a_ns) { continue; }
                                        if desc.tag_name().name() != "solidFill" { continue; }
                                        if let Some(c) = resolve_fill_color(&desc, theme_colors) {
                                            data_label_font_color = Some(c);
                                            break;
                                        }
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
                _ => {}
            }
        }

        // Parse series
        for ser_node in child.children()
            .filter(|n| n.is_element() && n.tag_name().name() == "ser" && n.tag_name().namespace() == Some(c_ns))
        {
            let s = parse_chart_series(&ser_node, c_ns, ser_type, chart_marker_default, theme_colors);
            if shared_categories.is_empty() && !s.categories.is_empty() {
                shared_categories = s.categories.clone();
            }
            all_series.push(s);

            // Per-series `<c:ser><c:dLbls>` fallback for chart-level
            // properties that Excel commonly writes on each series instead
            // of (or in addition to) the chart-level `<c:dLbls>`. Per
            // ECMA-376 §21.2.2.47 a series-level dLbls applies to that
            // series's data points; for the renderer we only need one
            // value, so the first series encountered "wins". This is how
            // the default-color/position/format travels when Excel emits
            // `<c:dLblPos val="inBase"/>` + `<c:txPr>` per series rather
            // than on the chart.
            if let Some(ser_dlbls) = ser_node.children()
                .find(|n| n.is_element()
                    && n.tag_name().name() == "dLbls"
                    && n.tag_name().namespace() == Some(c_ns))
            {
                for d in ser_dlbls.children().filter(|n| n.is_element()) {
                    match d.tag_name().name() {
                        "showVal" | "showPercent" => {
                            if d.attribute("val").unwrap_or("1") != "0" {
                                show_data_labels = true;
                            }
                        }
                        "dLblPos" => {
                            if data_label_position.is_none() {
                                data_label_position = d.attribute("val").map(|s| s.to_string());
                            }
                        }
                        "numFmt" => {
                            if data_label_format_code.is_none() {
                                data_label_format_code = d.attribute("formatCode")
                                    .map(|s| s.to_string())
                                    .filter(|s| !s.is_empty() && s != "General");
                            }
                        }
                        "txPr" => {
                            if data_label_font_color.is_none() {
                                for desc in d.descendants().filter(|n| n.is_element()) {
                                    if desc.tag_name().namespace() != Some(a_ns) { continue; }
                                    if desc.tag_name().name() != "solidFill" { continue; }
                                    if let Some(c) = resolve_fill_color(&desc, theme_colors) {
                                        data_label_font_color = Some(c);
                                        break;
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
    }

    if primary_type.is_empty() { return None; }

    // Fill in categories for series that have none (mixed charts share categories)
    for s in &mut all_series {
        if s.categories.is_empty() {
            s.categories = shared_categories.clone();
        }
    }

    // Stable-sort by `c:order` so the array is in Excel's display order.
    // ECMA-376 §21.2.2.28 — `<c:order>` is the authoritative stacking /
    // legend order, independent of document order.
    all_series.sort_by_key(|s| s.order);

    Some(ChartData {
        chart_type: primary_type,
        bar_dir,
        grouping,
        title,
        categories: shared_categories,
        series: all_series,
        show_data_labels,
        cat_axis_title,
        val_axis_title,
        show_legend,
        legend_pos,
        title_font_size_hpt,
        title_font_color,
        title_font_face,
        title_font_bold,
        cat_axis_font_bold,
        val_axis_font_bold,
        cat_axis_crosses,
        cat_axis_crosses_at,
        val_axis_crosses,
        val_axis_crosses_at,
        cat_axis_line_color,
        cat_axis_line_width_emu,
        val_axis_line_color,
        val_axis_line_width_emu,
        cat_axis_major_tick_mark,
        cat_axis_minor_tick_mark,
        val_axis_major_tick_mark,
        val_axis_minor_tick_mark,
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        val_axis_format_code,
        chart_bg,
        has_chart_sp_pr,
        legend_manual_layout,
        cat_axis_hidden,
        val_axis_hidden,
        bar_gap_width,
        bar_overlap,
        data_label_position,
        data_label_font_color,
        data_label_format_code,
        cat_axis_format_code,
        cat_axis_min,
        cat_axis_max,
        val_axis_min,
        val_axis_max,
        title_manual_layout,
        plot_area_manual_layout,
    })
}

/// `<c:catAx|valAx><c:numFmt@formatCode>` (ECMA-376 §21.2.2.21).
fn extract_axis_format_code(axis_node: &roxmltree::Node, c_ns: &str) -> Option<String> {
    axis_node.children()
        .find(|n| n.tag_name().name() == "numFmt" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("formatCode").map(|s| s.to_string()))
        .filter(|s| !s.is_empty() && s != "General")
}

/// `<c:catAx|valAx><c:majorTickMark val>` / `<c:minorTickMark val>` —
/// `none` / `out` / `in` / `cross` (ECMA-376 §21.2.2.49 ST_TickMark).
fn extract_axis_tick_mark(axis_node: &roxmltree::Node, c_ns: &str, name: &str) -> Option<String> {
    axis_node.children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string())
}

/// `<c:catAx|valAx><c:spPr><a:ln>` — resolved color (no `#`) and width
/// (EMU) for the axis line itself. None when not set.
fn extract_axis_line_style(
    axis_node: &roxmltree::Node,
    c_ns: &str,
    theme_colors: &[String],
) -> (Option<String>, Option<u32>) {
    let Some(sp_pr) = axis_node.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns))
    else { return (None, None); };
    let Some(ln) = sp_pr.children().find(|n| n.tag_name().name() == "ln") else { return (None, None); };
    let width = ln.attribute("w").and_then(|v| v.parse::<u32>().ok());
    let color = extract_solid_fill_in_drawingml(&ln, theme_colors);
    (color, width)
}

/// `<c:catAx|valAx><c:txPr>...defRPr@b>` — bold flag for axis tick labels.
fn extract_axis_tick_label_bold(axis_node: &roxmltree::Node, c_ns: &str) -> Option<bool> {
    let txpr = axis_node.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("b").map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// `<c:catAx|valAx><c:crosses>` and `<c:crossesAt>` — where the axis sits
/// along its perpendicular axis. `crosses` is a string ("autoZero" |
/// "min" | "max"); `crossesAt` is an explicit numeric override that
/// takes precedence at render time.
fn extract_axis_crosses(axis_node: &roxmltree::Node, c_ns: &str) -> (Option<String>, Option<f64>) {
    let crosses = axis_node.children()
        .find(|n| n.tag_name().name() == "crosses" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let crosses_at = axis_node.children()
        .find(|n| n.tag_name().name() == "crossesAt" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    (crosses, crosses_at)
}

/// Read explicit `<c:scaling><c:min>` / `<c:scaling><c:max>` values, returning
/// `(min, max)` where each is `None` if the axis didn't override that bound.
fn extract_axis_scaling(axis_node: &roxmltree::Node, c_ns: &str) -> Option<(Option<f64>, Option<f64>)> {
    let scaling = axis_node.children()
        .find(|n| n.tag_name().name() == "scaling" && n.tag_name().namespace() == Some(c_ns))?;
    let mn = scaling.children()
        .find(|n| n.tag_name().name() == "min" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    let mx = scaling.children()
        .find(|n| n.tag_name().name() == "max" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<f64>().ok());
    if mn.is_some() || mx.is_some() {
        Some((mn, mx))
    } else {
        None
    }
}

/// `<c:catAx|valAx><c:delete val="1"/>` — true when the axis is hidden.
fn axis_is_deleted(axis_node: &roxmltree::Node, c_ns: &str) -> bool {
    axis_node.children()
        .find(|n| n.tag_name().name() == "delete" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|v| v != "0" && !v.eq_ignore_ascii_case("false"))
        .unwrap_or(false)
}

/// Extract a `<c:layout><c:manualLayout>` block. The given `layout_node` is
/// `<c:layout>` (parent of `<c:manualLayout>`). Returns None when the layout
/// is auto (no `manualLayout` child).
fn extract_manual_layout(layout_node: &roxmltree::Node, c_ns: &str) -> Option<ManualLayout> {
    let ml = layout_node.children()
        .find(|n| n.tag_name().name() == "manualLayout" && n.tag_name().namespace() == Some(c_ns))?;
    let mut x_mode = "edge".to_string();
    let mut y_mode = "edge".to_string();
    let mut layout_target: Option<String> = None;
    let mut x = 0.0_f64;
    let mut y = 0.0_f64;
    let mut w: Option<f64> = None;
    let mut h: Option<f64> = None;
    for ch in ml.children().filter(|n| n.is_element() && n.tag_name().namespace() == Some(c_ns)) {
        let val = ch.attribute("val").map(|s| s.to_string());
        match ch.tag_name().name() {
            "xMode" => { if let Some(v) = val { x_mode = v; } }
            "yMode" => { if let Some(v) = val { y_mode = v; } }
            "layoutTarget" => { layout_target = val; }
            "x" => { if let Some(v) = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()) { x = v; } }
            "y" => { if let Some(v) = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()) { y = v; } }
            "w" => { w = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()); }
            "h" => { h = ch.attribute("val").and_then(|s| s.parse::<f64>().ok()); }
            _ => {}
        }
    }
    Some(ManualLayout { x_mode, y_mode, layout_target, x, y, w, h })
}

/// Extract a category/value axis tick-label font size (hundredths of a point)
/// from the first `a:defRPr@sz` (or `a:rPr@sz`) inside the axis' `c:txPr`.
/// ECMA-376 §21.2.2.17 — `<c:txPr>` controls tick label text properties.
fn extract_axis_tick_label_size(axis_node: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<i32> {
    let txpr = axis_node.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))?;
    txpr.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// Extract the chart title's font size (hundredths of a point) from the first
/// `a:defRPr@sz` or `a:rPr@sz` found under `c:title`. Returns None when absent.
fn extract_chart_title_size(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<i32> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("sz").and_then(|v| v.parse::<i32>().ok())
    })
}

/// Extract chart title bold flag from the first `a:defRPr@b` / `a:rPr@b`
/// inside `c:title`. Returns None when not specified (renderer treats as
/// not bold).
fn extract_chart_title_bold(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<bool> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        let tag = n.tag_name().name();
        if tag != "defRPr" && tag != "rPr" { return None; }
        n.attribute("b").map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    })
}

/// Extract the chart title's font color (hex without '#') from the first
/// `a:solidFill/a:srgbClr@val` inside `c:title`. Only srgb is resolved here —
/// scheme colors would require the workbook theme, which isn't wired through
/// to chart parsing yet.
fn extract_chart_title_color(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<String> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        if n.tag_name().name() != "srgbClr" { return None; }
        // Skip srgbClr nodes that aren't inside a solidFill (e.g. a gradient stop).
        let parent_is_solid = n.parent()
            .map(|p| p.tag_name().name() == "solidFill" && p.tag_name().namespace() == Some(a_ns))
            .unwrap_or(false);
        if !parent_is_solid { return None; }
        n.attribute("val").map(|s| s.to_string())
    })
}

/// Extract the chart title's font family from the first `a:latin@typeface`
/// descendant of `c:title` (ECMA-376 DrawingML §20.1.4.2.24).
fn extract_chart_title_face(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<String> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    title_node.descendants().find_map(|n| {
        if !n.is_element() { return None; }
        if n.tag_name().namespace() != Some(a_ns) { return None; }
        if n.tag_name().name() != "latin" { return None; }
        n.attribute("typeface").map(|s| s.to_string())
    })
}

/// Extract plain text from `c:chart/c:title`.
fn extract_chart_title(chart_root: &roxmltree::Node, c_ns: &str, a_ns: &str) -> Option<String> {
    let title_node = chart_root.children()
        .find(|n| n.tag_name().name() == "title" && n.tag_name().namespace() == Some(c_ns))?;
    // c:title/c:tx/c:rich/a:p/a:r/a:t  or  c:title/c:tx/c:strRef/c:strCache/c:pt/c:v
    let mut text = String::new();
    for node in title_node.descendants() {
        if node.tag_name().name() == "t" && node.tag_name().namespace() == Some(a_ns) {
            if let Some(t) = node.text() { text.push_str(t); }
        }
        if node.tag_name().name() == "v" && node.tag_name().namespace() == Some(c_ns) {
            if let Some(t) = node.text() { text.push_str(t); }
        }
    }
    if text.is_empty() { None } else { Some(text) }
}

/// Parse one `<c:ser>` element.
/// Resolve the fill color from a single DrawingML fill element. The caller
/// passes either a `<c:spPr>` (in which case we look for the first `<a:solidFill>`
/// **as a direct child** to avoid picking up text fills nested under `<c:dLbls>`
/// / `<c:txPr>`) or the `<a:solidFill>` directly. Supports `a:srgbClr` (explicit
/// hex) and `a:schemeClr` (theme accent/dark/light).
/// Theme colors use drawingML names (`accent1`..`accent6`, `dk1`/`dk2`/`lt1`/`lt2`)
/// which map to the parser's natural-order theme array (dk1@0, lt1@1, dk2@2,
/// lt2@3, accent1@4 … accent6@9).
fn resolve_fill_color(fill_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    // Accept either a `<a:solidFill>` directly or a `<c:spPr>` whose first
    // fill-ish child is `<a:solidFill>`. Looking at *direct* children (not
    // descendants) is intentional — chart series often carry label/axis text
    // colors under `c:dLbls`/`c:txPr` which must NOT be misread as fill.
    let solid = if fill_node.tag_name().name() == "solidFill" {
        Some(*fill_node)
    } else {
        fill_node.children().find(|n| n.is_element() && n.tag_name().name() == "solidFill")
    }?;
    for n in solid.children().filter(|n| n.is_element()) {
        let tag = n.tag_name().name();
        if tag == "srgbClr" {
            if let Some(v) = n.attribute("val") {
                return Some(v.to_lowercase());
            }
        }
        if tag == "schemeClr" {
            if let Some(v) = n.attribute("val") {
                let idx = match v {
                    "dk1"  | "tx1" => Some(0),
                    "lt1"  | "bg1" => Some(1),
                    "dk2"  | "tx2" => Some(2),
                    "lt2"  | "bg2" => Some(3),
                    "accent1" => Some(4),
                    "accent2" => Some(5),
                    "accent3" => Some(6),
                    "accent4" => Some(7),
                    "accent5" => Some(8),
                    "accent6" => Some(9),
                    "hlink"    => Some(10),
                    "folHlink" => Some(11),
                    _ => None,
                };
                if let Some(i) = idx {
                    if let Some(c) = theme_colors.get(i) {
                        return Some(c.trim_start_matches('#').to_lowercase());
                    }
                }
            }
        }
    }
    None
}

/// Series fill color from `<c:ser><c:spPr><a:solidFill>`. Returns None when
/// the series has no direct `<c:spPr>` or its fill isn't a recognised solid.
fn resolve_series_color(ser_node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    let sp_pr = ser_node.children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr")?;
    resolve_fill_color(&sp_pr, theme_colors)
}

fn parse_chart_series(
    node: &roxmltree::Node,
    c_ns: &str,
    ser_type: &str,
    chart_marker_default: bool,
    theme_colors: &[String],
) -> ChartSeries {
    let name = extract_series_name(node, c_ns);

    // `<c:idx val>` (ECMA-376 §21.2.2.27) — the canonical series index Excel
    // uses for default color selection. When absent, fall back to 0 so we
    // still produce a deterministic palette pick. `<c:order>` is the display
    // order (legend / stacking) and is intentionally ignored for coloring.
    let idx: usize = node.children()
        .find(|n| n.tag_name().name() == "idx" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    // `<c:order val>` (ECMA-376 §21.2.2.28) — series display order. Used for
    // stacking and legend ordering. Defaults to 0.
    let order: usize = node.children()
        .find(|n| n.tag_name().name() == "order" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0);

    // For scatter: xVal → categories (as strings), yVal → values
    // For others:  cat  → categories,             val  → values
    let (cat_tag, val_tag) = if ser_type == "scatter" { ("xVal", "yVal") } else { ("cat", "val") };

    let categories = collect_str_cache(node, c_ns, cat_tag);
    let values     = collect_num_cache(node, c_ns, val_tag);
    // `<c:val><c:numRef><c:numCache><c:formatCode>` (ECMA-376 §21.2.2.37)
    // preserves the Excel number format Excel stamped onto the cached values
    // at save time; absent "General" codes return None so the renderer can
    // fall back cleanly.
    let val_format_code = node.children()
        .find(|n| n.tag_name().name() == val_tag && n.tag_name().namespace() == Some(c_ns))
        .and_then(|val_node| val_node.descendants()
            .find(|n| n.is_element()
                && n.tag_name().name() == "formatCode"
                && n.tag_name().namespace() == Some(c_ns)))
        .and_then(|n| n.text().map(|s| s.to_string()))
        .filter(|s| !s.is_empty() && s != "General");

    // Series fill color from c:spPr/a:solidFill (supports a:srgbClr and a:schemeClr).
    // For schemeClr, resolves "accentN"/"dk1"/etc. against the workbook theme.
    //
    // When the series has no explicit fill, Excel's default palette assigns
    // `theme.accent[idx % 6 + 1]` — i.e. accent1, accent2, … cycling by
    // `<c:idx>`. That's the rule behind "first series = green, second = red"
    // when the theme's accent1/accent2 are green/red. We inline that
    // resolution here so the renderer doesn't need theme access.
    let color = resolve_series_color(node, theme_colors)
        .or_else(|| {
            // Theme order in `theme_colors`: dk1@0, lt1@1, dk2@2, lt2@3, accent1@4 … accent6@9.
            theme_colors.get(4 + (idx % 6)).map(|c| c.trim_start_matches('#').to_lowercase())
        });

    // Marker visibility (ECMA-376 §21.2.2.32 — c:marker/c:symbol default is
    // "none"). A per-series <c:marker><c:symbol> overrides; otherwise fall
    // back to the chart-type-level <c:lineChart><c:marker val> flag. Scatter
    // charts default to visible markers even without an explicit flag.
    let marker_node = node.children()
        .find(|n| n.tag_name().name() == "marker" && n.tag_name().namespace() == Some(c_ns));
    let (marker_symbol, marker_size, marker_fill, marker_line) = parse_marker_block(marker_node, c_ns, theme_colors);
    let show_marker = match (&marker_symbol, ser_type) {
        (Some(sym), _)   => sym != "none",
        (None, "scatter") => true,
        _                 => chart_marker_default,
    };

    let data_point_overrides = parse_data_point_overrides(node, c_ns, theme_colors);
    // `<c15:datalabelsRange>` lookup table for `<a:fld type="CELLRANGE">`
    // labels. Excel saves the actual cached label strings here; we resolve
    // CELLRANGE field placeholders against this at parse time so the
    // renderer just receives plain strings.
    let dlbl_range_cache = collect_dlbl_range_cache(node, c_ns);
    let (series_data_labels, data_label_overrides) = parse_data_labels(node, c_ns, theme_colors, &dlbl_range_cache);
    let err_bars = parse_error_bars(node, c_ns, &values, theme_colors);

    ChartSeries {
        name,
        series_type: ser_type.to_string(),
        categories,
        values,
        color,
        show_marker,
        val_format_code,
        order,
        marker_symbol,
        marker_size,
        marker_fill,
        marker_line,
        data_point_overrides,
        data_label_overrides,
        series_data_labels,
        err_bars,
    }
}

/// Parse `<c:marker>` into (symbol, size, fill, line) — all hex colors are
/// returned without `#`. ECMA-376 §21.2.2.32 / §21.2.2.34. The fill and
/// line colors come from `<c:spPr>` nested inside marker.
fn parse_marker_block(
    marker_node: Option<roxmltree::Node>,
    c_ns: &str,
    theme_colors: &[String],
) -> (Option<String>, Option<u32>, Option<String>, Option<String>) {
    let Some(mk) = marker_node else { return (None, None, None, None); };
    let symbol = mk.children()
        .find(|n| n.tag_name().name() == "symbol" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let size = mk.children()
        .find(|n| n.tag_name().name() == "size" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<u32>().ok());
    let sp_pr = mk.children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns));
    let fill = sp_pr.and_then(|p| extract_solid_fill_in_drawingml(&p, theme_colors));
    let line = sp_pr.and_then(|p| {
        let ln = p.children().find(|n| n.tag_name().name() == "ln");
        ln.and_then(|l| extract_solid_fill_in_drawingml(&l, theme_colors))
    });
    (symbol, size, fill, line)
}

/// Locate the first `a:solidFill > a:srgbClr@val` or `a:schemeClr@val` under
/// `node` (children only, not deep descendants — chart spPr is structured
/// shallowly). Returns the resolved hex without `#`. Handles theme refs and
/// `lumMod`/`lumOff`/`tint`/`shade`/`alpha` color transforms by delegating
/// to `apply_color_transforms`.
fn extract_solid_fill_in_drawingml(
    parent: &roxmltree::Node,
    theme_colors: &[String],
) -> Option<String> {
    for fill in parent.children().filter(|n| n.is_element() && n.tag_name().name() == "solidFill") {
        for clr in fill.children().filter(|n| n.is_element()) {
            match clr.tag_name().name() {
                "srgbClr" => {
                    if let Some(rgb) = clr.attribute("val") {
                        return Some(apply_color_transforms(rgb, &clr));
                    }
                }
                "schemeClr" => {
                    if let Some(scheme) = clr.attribute("val") {
                        let base = resolve_scheme_color(scheme, theme_colors);
                        if let Some(b) = base {
                            return Some(apply_color_transforms(&b, &clr));
                        }
                    }
                }
                _ => {}
            }
        }
    }
    None
}

/// Look up a scheme color name ("dk1"/"lt1"/"dk2"/"lt2"/"accent1"…"accent6"
/// /"hlink"/"folHlink") in the workbook theme color table. Returns hex
/// (no `#`) or None when unknown.
fn resolve_scheme_color(name: &str, theme_colors: &[String]) -> Option<String> {
    // Theme order (parse_theme_colors): dk1@0, lt1@1, dk2@2, lt2@3,
    // accent1@4..accent6@9, hlink@10, folHlink@11.
    let idx = match name {
        "dk1" | "tx1" | "bg2" => 0,
        "lt1" | "bg1" | "tx2" => 1,
        "dk2" => 2,
        "lt2" => 3,
        "accent1" => 4, "accent2" => 5, "accent3" => 6,
        "accent4" => 7, "accent5" => 8, "accent6" => 9,
        "hlink" => 10, "folHlink" => 11,
        _ => return None,
    };
    theme_colors.get(idx).map(|s| s.trim_start_matches('#').to_string())
}

/// Apply DrawingML color transforms (`lumMod`/`lumOff`/`tint`/`shade`/
/// `alpha` — drop alpha) found as children of a color element. Returns a
/// hex string without `#`. Already-existing `apply_tint` handles
/// lumMod-style brightness changes for the simpler `lumMod-only` case;
/// this widens it to combine multiple transforms.
fn apply_color_transforms(base_hex: &str, color_el: &roxmltree::Node) -> String {
    let cleaned = base_hex.trim_start_matches('#');
    let r = u8::from_str_radix(&cleaned.get(0..2).unwrap_or("00"), 16).unwrap_or(0);
    let g = u8::from_str_radix(&cleaned.get(2..4).unwrap_or("00"), 16).unwrap_or(0);
    let b = u8::from_str_radix(&cleaned.get(4..6).unwrap_or("00"), 16).unwrap_or(0);
    let mut rf = r as f64 / 255.0;
    let mut gf = g as f64 / 255.0;
    let mut bf = b as f64 / 255.0;
    for child in color_el.children().filter(|n| n.is_element()) {
        let pct = child.attribute("val")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100000.0);
        let Some(p) = pct else { continue };
        match child.tag_name().name() {
            "lumMod"  => { rf *= p; gf *= p; bf *= p; }
            "lumOff"  => { rf += p; gf += p; bf += p; }
            "tint"    => {
                // ECMA-376: lighten toward 1.0 by `p` (0..1).
                rf = rf + (1.0 - rf) * p;
                gf = gf + (1.0 - gf) * p;
                bf = bf + (1.0 - bf) * p;
            }
            "shade"   => {
                // Darken toward 0 by `1 - p`.
                rf *= p; gf *= p; bf *= p;
            }
            // alpha is dropped — we render opaque.
            _ => {}
        }
    }
    let clamp = |v: f64| -> u8 { (v.max(0.0).min(1.0) * 255.0).round() as u8 };
    format!("{:02X}{:02X}{:02X}", clamp(rf), clamp(gf), clamp(bf))
}

/// Walk every `<c:dPt>` direct child of the series and collect per-point
/// overrides. Multiple `<c:dPt>` per series is normal; each one targets a
/// single `<c:idx>` (ECMA-376 §21.2.2.39).
fn parse_data_point_overrides(
    ser_node: &roxmltree::Node,
    c_ns: &str,
    theme_colors: &[String],
) -> Vec<DataPointOverride> {
    let mut result = Vec::new();
    for dpt in ser_node.children().filter(|n| n.is_element() && n.tag_name().name() == "dPt" && n.tag_name().namespace() == Some(c_ns)) {
        let idx = dpt.children()
            .find(|n| n.tag_name().name() == "idx" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        let sp_pr = dpt.children()
            .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns));
        let color = sp_pr.and_then(|p| extract_solid_fill_in_drawingml(&p, theme_colors));
        let mk = dpt.children()
            .find(|n| n.tag_name().name() == "marker" && n.tag_name().namespace() == Some(c_ns));
        let (marker_symbol, marker_size, marker_fill, marker_line) = parse_marker_block(mk, c_ns, theme_colors);
        result.push(DataPointOverride {
            idx, color, marker_symbol, marker_size, marker_fill, marker_line,
        });
    }
    result
}

/// Resolve `<c:ser><c:extLst><c:ext><c15:datalabelsRange>` cache: index →
/// label text. Used to substitute `<a:fld type="CELLRANGE">` placeholders.
/// Returns indices in 0..ptCount; missing entries are empty strings.
fn collect_dlbl_range_cache(ser_node: &roxmltree::Node, c_ns: &str) -> HashMap<u32, String> {
    let mut map: HashMap<u32, String> = HashMap::new();
    let Some(ext_lst) = ser_node.children().find(|n| n.tag_name().name() == "extLst" && n.tag_name().namespace() == Some(c_ns)) else { return map; };
    for ext in ext_lst.children().filter(|n| n.is_element() && n.tag_name().name() == "ext" && n.tag_name().namespace() == Some(c_ns)) {
        for range in ext.descendants().filter(|n| n.is_element() && n.tag_name().name() == "datalabelsRange") {
            for cache in range.children().filter(|n| n.is_element() && n.tag_name().name() == "dlblRangeCache") {
                for pt in cache.children().filter(|n| n.is_element() && n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(c_ns)) {
                    let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<u32>().ok()) else { continue };
                    let v = pt.children()
                        .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(c_ns))
                        .and_then(|n| n.text())
                        .unwrap_or("")
                        .to_string();
                    map.insert(idx, v);
                }
            }
        }
    }
    map
}

/// Walk a `<c:tx><c:rich>` (or any DrawingML rich-text root) and reduce it
/// to plain text. `<a:fld type="CELLRANGE">` placeholders are substituted
/// from `cellrange_cache` keyed by `idx`. Other field types and runs are
/// concatenated. Newlines come from paragraph breaks.
fn flatten_rich_text(
    rich_root: &roxmltree::Node,
    cellrange_cache: Option<&str>,
) -> String {
    let mut out = String::new();
    let mut first_para = true;
    for p in rich_root.descendants().filter(|n| n.is_element() && n.tag_name().name() == "p") {
        if !first_para { out.push('\n'); }
        first_para = false;
        for child in p.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "r" => {
                    if let Some(t) = child.children().find(|n| n.tag_name().name() == "t") {
                        if let Some(s) = t.text() { out.push_str(s); }
                    }
                }
                "fld" => {
                    let typ = child.attribute("type").unwrap_or("");
                    if typ == "CELLRANGE" {
                        if let Some(s) = cellrange_cache { out.push_str(s); }
                    } else {
                        // VALUE/SERIESNAME/CATEGORYNAME field placeholders are
                        // resolved by the renderer using the series data, since
                        // they don't need cell-range expansion. We embed a marker
                        // so the renderer can recognise them.
                        if let Some(t) = child.children().find(|n| n.tag_name().name() == "t") {
                            if let Some(s) = t.text() { out.push_str(s); }
                        }
                    }
                }
                _ => {}
            }
        }
    }
    out
}

/// Parse `<c:dLbls>` (series-level defaults + per-idx overrides).
fn parse_data_labels(
    ser_node: &roxmltree::Node,
    c_ns: &str,
    theme_colors: &[String],
    cellrange_cache: &HashMap<u32, String>,
) -> (Option<SeriesDataLabels>, Vec<DataLabelOverride>) {
    let Some(d_lbls) = ser_node.children()
        .find(|n| n.tag_name().name() == "dLbls" && n.tag_name().namespace() == Some(c_ns))
    else { return (None, Vec::new()); };

    let bool_attr = |n: &roxmltree::Node, name: &str| {
        n.children()
            .find(|c| c.tag_name().name() == name && c.tag_name().namespace() == Some(c_ns))
            .and_then(|c| c.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false)
    };

    let position = d_lbls.children()
        .find(|n| n.tag_name().name() == "dLblPos" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .map(|s| s.to_string());
    let format_code = d_lbls.children()
        .find(|n| n.tag_name().name() == "numFmt" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("formatCode"))
        .map(|s| s.to_string());
    let font_color = d_lbls.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
                .and_then(|def| extract_solid_fill_in_drawingml(&def, theme_colors))
        });
    let font_bold_default = d_lbls.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr"))
                .and_then(|n| n.attribute("b"))
                .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
        });
    let font_size_default = d_lbls.children()
        .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|tx| {
            tx.descendants()
                .find(|n| n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr"))
                .and_then(|n| n.attribute("sz"))
                .and_then(|v| v.parse::<i32>().ok())
        });

    let series_defaults = SeriesDataLabels {
        show_val: bool_attr(&d_lbls, "showVal"),
        show_cat_name: bool_attr(&d_lbls, "showCatName"),
        show_ser_name: bool_attr(&d_lbls, "showSerName"),
        show_percent: bool_attr(&d_lbls, "showPercent"),
        position: position.clone(),
        font_color: font_color.clone(),
        format_code,
        font_bold: font_bold_default,
        font_size_hpt: font_size_default,
    };

    let mut overrides = Vec::new();
    for dl in d_lbls.children().filter(|n| n.is_element() && n.tag_name().name() == "dLbl" && n.tag_name().namespace() == Some(c_ns)) {
        let idx = dl.children()
            .find(|n| n.tag_name().name() == "idx" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        // <c:delete val="1"/> — the user explicitly removed this point's
        // label. Render as empty text so the renderer skips it.
        let deleted = dl.children()
            .find(|n| n.tag_name().name() == "delete" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);
        let pos = dl.children()
            .find(|n| n.tag_name().name() == "dLblPos" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());
        let cache_for_idx = cellrange_cache.get(&idx).map(|s| s.as_str());
        let text = if deleted {
            String::new()
        } else {
            // Custom text lives at `<c:tx><c:rich>` (ECMA-376 §21.2.2.46).
            // Without `<c:tx>` the override is metadata-only (e.g. only a
            // position change); show the cellrange cache value when
            // available, else empty.
            let tx = dl.children()
                .find(|n| n.tag_name().name() == "tx" && n.tag_name().namespace() == Some(c_ns));
            match tx {
                Some(tx_node) => flatten_rich_text(&tx_node, cache_for_idx),
                None => cache_for_idx.unwrap_or("").to_string(),
            }
        };
        let font_color = dl.children()
            .find(|n| n.tag_name().name() == "txPr" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|tx| {
                tx.descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
                    .and_then(|def| extract_solid_fill_in_drawingml(&def, theme_colors))
            });
        let font_size_hpt = dl.descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "defRPr")
            .and_then(|def| def.attribute("sz"))
            .and_then(|v| v.parse::<i32>().ok());
        let font_bold = dl.descendants()
            .find(|n| n.is_element() && (n.tag_name().name() == "defRPr" || n.tag_name().name() == "rPr"))
            .and_then(|def| def.attribute("b"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        overrides.push(DataLabelOverride { idx, text, position: pos, font_color, font_size_hpt, font_bold });
    }

    let any_default = series_defaults.show_val
        || series_defaults.show_cat_name
        || series_defaults.show_ser_name
        || series_defaults.show_percent
        || series_defaults.position.is_some()
        || series_defaults.font_color.is_some()
        || series_defaults.format_code.is_some()
        || series_defaults.font_bold.is_some()
        || series_defaults.font_size_hpt.is_some();
    let series_out = if any_default { Some(series_defaults) } else { None };
    (series_out, overrides)
}

/// Parse all `<c:errBars>` direct children of a series and resolve per-
/// point plus / minus deltas to absolute numbers. Each errBars block fixes
/// a direction (x|y); a series can have at most one of each direction.
fn parse_error_bars(
    ser_node: &roxmltree::Node,
    c_ns: &str,
    series_values: &[Option<f64>],
    theme_colors: &[String],
) -> Vec<ErrBars> {
    let mut result = Vec::new();
    for eb in ser_node.children().filter(|n| n.is_element() && n.tag_name().name() == "errBars" && n.tag_name().namespace() == Some(c_ns)) {
        let dir = eb.children()
            .find(|n| n.tag_name().name() == "errDir" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("y")
            .to_string();
        let bar_type = eb.children()
            .find(|n| n.tag_name().name() == "errBarType" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("both")
            .to_string();
        let val_type = eb.children()
            .find(|n| n.tag_name().name() == "errValType" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .unwrap_or("fixedVal")
            .to_string();
        let no_end_cap = eb.children()
            .find(|n| n.tag_name().name() == "noEndCap" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);

        let n_points = series_values.len();
        let mut plus: Vec<Option<f64>> = vec![None; n_points];
        let mut minus: Vec<Option<f64>> = vec![None; n_points];

        match val_type.as_str() {
            "cust" => {
                for (slot, target) in [("plus", &mut plus), ("minus", &mut minus)] {
                    let Some(side) = eb.children().find(|n| n.tag_name().name() == slot && n.tag_name().namespace() == Some(c_ns)) else { continue };
                    let vals = extract_num_block(&side, c_ns, n_points);
                    if !vals.is_empty() {
                        let len = vals.len().min(target.len());
                        for i in 0..len { target[i] = vals[i]; }
                    }
                }
            }
            "fixedVal" => {
                let v = eb.children()
                    .find(|n| n.tag_name().name() == "val" && n.tag_name().namespace() == Some(c_ns))
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for i in 0..n_points {
                    plus[i] = Some(v);
                    minus[i] = Some(v);
                }
            }
            "percentage" => {
                // Each point's bar = abs(value) * pct/100.
                let pct = eb.children()
                    .find(|n| n.tag_name().name() == "val" && n.tag_name().namespace() == Some(c_ns))
                    .and_then(|n| n.attribute("val"))
                    .and_then(|s| s.parse::<f64>().ok())
                    .unwrap_or(0.0);
                for (i, v) in series_values.iter().enumerate() {
                    if let Some(val) = v {
                        let d = val.abs() * pct / 100.0;
                        plus[i] = Some(d); minus[i] = Some(d);
                    }
                }
            }
            "stdErr" | "stdDev" => {
                let nums: Vec<f64> = series_values.iter().filter_map(|v| *v).collect();
                if !nums.is_empty() {
                    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
                    let var = nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / nums.len() as f64;
                    let std = var.sqrt();
                    let mult = eb.children()
                        .find(|n| n.tag_name().name() == "val" && n.tag_name().namespace() == Some(c_ns))
                        .and_then(|n| n.attribute("val"))
                        .and_then(|s| s.parse::<f64>().ok())
                        .unwrap_or(1.0);
                    let sample = if val_type == "stdErr" {
                        std / (nums.len() as f64).sqrt()
                    } else { std };
                    let delta = sample * mult;
                    for i in 0..n_points {
                        plus[i] = Some(delta); minus[i] = Some(delta);
                    }
                }
            }
            _ => {}
        }

        let sp_pr = eb.children()
            .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(c_ns));
        let color = sp_pr.and_then(|p| {
            let ln = p.children().find(|n| n.tag_name().name() == "ln");
            match ln {
                Some(l) => extract_solid_fill_in_drawingml(&l, theme_colors),
                None => extract_solid_fill_in_drawingml(&p, theme_colors),
            }
        });
        let line_width_emu = sp_pr
            .and_then(|p| p.children().find(|n| n.tag_name().name() == "ln"))
            .and_then(|ln| ln.attribute("w"))
            .and_then(|v| v.parse::<u32>().ok());
        let dash = sp_pr
            .and_then(|p| p.children().find(|n| n.tag_name().name() == "ln"))
            .and_then(|ln| ln.children().find(|n| n.tag_name().name() == "prstDash"))
            .and_then(|n| n.attribute("val"))
            .map(|s| s.to_string());

        result.push(ErrBars {
            dir, bar_type,
            plus, minus,
            no_end_cap,
            color, line_width_emu, dash,
        });
    }
    result
}

/// Read a `<c:numRef><c:numCache>` or `<c:numLit>` block under `parent` and
/// return per-point values keyed by `<c:pt idx>`. Length is at least
/// `expected_len` (padded with None).
fn extract_num_block(parent: &roxmltree::Node, c_ns: &str, expected_len: usize) -> Vec<Option<f64>> {
    let cache = parent.descendants()
        .find(|n| n.is_element()
            && (n.tag_name().name() == "numCache" || n.tag_name().name() == "numLit")
            && n.tag_name().namespace() == Some(c_ns));
    let Some(cache) = cache else { return Vec::new(); };
    let pt_count: usize = cache.children()
        .find(|n| n.tag_name().name() == "ptCount" && n.tag_name().namespace() == Some(c_ns))
        .and_then(|n| n.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(expected_len);
    let len = pt_count.max(expected_len);
    let mut values: Vec<Option<f64>> = vec![None; len];
    for pt in cache.children().filter(|n| n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(c_ns)) {
        let Some(idx) = pt.attribute("idx").and_then(|v| v.parse::<usize>().ok()) else { continue };
        let v = pt.children()
            .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if idx < values.len() { values[idx] = v; }
    }
    values
}

/// Extract series name from `c:tx`.
fn extract_series_name(node: &roxmltree::Node, c_ns: &str) -> String {
    // c:tx/c:strRef/c:strCache/c:pt[@idx=0]/c:v
    // or c:tx/c:v
    if let Some(tx) = node.children().find(|n| n.tag_name().name() == "tx" && n.tag_name().namespace() == Some(c_ns)) {
        for desc in tx.descendants() {
            if desc.tag_name().name() == "v" && desc.tag_name().namespace() == Some(c_ns) {
                if let Some(t) = desc.text() {
                    if !t.is_empty() { return t.to_string(); }
                }
            }
        }
    }
    String::new()
}

/// Collect string values from a cache child element (e.g. `<c:cat>` or `<c:xVal>`).
/// Reads `c:strRef/c:strCache`, `c:multiLvlStrRef/c:multiLvlStrCache`, or
/// `c:numRef/c:numCache` (formats numbers as strings).
fn collect_str_cache(ser_node: &roxmltree::Node, c_ns: &str, child_tag: &str) -> Vec<String> {
    let Some(child) = ser_node.children()
        .find(|n| n.tag_name().name() == child_tag && n.tag_name().namespace() == Some(c_ns))
    else { return Vec::new(); };

    // Multi-level categories: use only the first (innermost) lvl to get primary labels.
    if let Some(multi_cache) = child.descendants()
        .find(|n| n.tag_name().name() == "multiLvlStrCache" && n.tag_name().namespace() == Some(c_ns))
    {
        let pt_count: usize = multi_cache.children()
            .find(|n| n.tag_name().name() == "ptCount" && n.tag_name().namespace() == Some(c_ns))
            .and_then(|n| n.attribute("val"))
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        if let Some(first_lvl) = multi_cache.children()
            .find(|n| n.tag_name().name() == "lvl" && n.tag_name().namespace() == Some(c_ns))
        {
            let mut pts: Vec<(usize, String)> = Vec::new();
            for pt in first_lvl.children()
                .filter(|n| n.is_element() && n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(c_ns))
            {
                let idx: usize = pt.attribute("idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                let val = pt.children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            let len = pt_count.max(pts.iter().map(|(i, _)| i + 1).max().unwrap_or(0));
            let mut result = vec![String::new(); len];
            for (idx, val) in pts {
                if idx < result.len() { result[idx] = val; }
            }
            return result;
        }
    }

    // Standard strRef/strCache or numRef/numCache
    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, String)> = Vec::new();
    for desc in child.descendants() {
        match desc.tag_name().name() {
            "ptCount" if desc.tag_name().namespace() == Some(c_ns) => {
                pt_count = desc.attribute("val").and_then(|v| v.parse().ok()).unwrap_or(0);
            }
            "pt" if desc.tag_name().namespace() == Some(c_ns) => {
                let idx: usize = desc.attribute("idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                let val = desc.children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .unwrap_or("")
                    .to_string();
                pts.push((idx, val));
            }
            _ => {}
        }
    }
    if pt_count == 0 { pt_count = pts.len(); }
    let mut result = vec![String::new(); pt_count];
    for (idx, val) in pts {
        if idx < result.len() { result[idx] = val; }
    }
    result
}

/// Collect numeric values from a cache child element (e.g. `<c:val>` or `<c:yVal>`).
fn collect_num_cache(ser_node: &roxmltree::Node, c_ns: &str, child_tag: &str) -> Vec<Option<f64>> {
    let Some(child) = ser_node.children()
        .find(|n| n.tag_name().name() == child_tag && n.tag_name().namespace() == Some(c_ns))
    else { return Vec::new(); };

    let mut pt_count: usize = 0;
    let mut pts: Vec<(usize, f64)> = Vec::new();
    for desc in child.descendants() {
        match desc.tag_name().name() {
            "ptCount" if desc.tag_name().namespace() == Some(c_ns) => {
                pt_count = desc.attribute("val").and_then(|v| v.parse().ok()).unwrap_or(0);
            }
            "pt" if desc.tag_name().namespace() == Some(c_ns) => {
                let idx: usize = desc.attribute("idx").and_then(|v| v.parse().ok()).unwrap_or(0);
                if let Some(v) = desc.children()
                    .find(|n| n.tag_name().name() == "v")
                    .and_then(|n| n.text())
                    .and_then(|t| t.parse::<f64>().ok())
                {
                    pts.push((idx, v));
                }
            }
            _ => {}
        }
    }
    if pt_count == 0 { pt_count = pts.len(); }
    let mut result: Vec<Option<f64>> = vec![None; pt_count];
    for (idx, val) in pts {
        if idx < result.len() { result[idx] = Some(val); }
    }
    result
}

fn parse_sqref(s: &str) -> Vec<CellRange> {
    s.split_whitespace().map(|range_str| {
        if let Some((a, b)) = range_str.split_once(':') {
            let (left, top) = parse_cell_ref(a);
            let (right, bottom) = parse_cell_ref(b);
            CellRange { top, left, bottom, right }
        } else {
            let (col, row) = parse_cell_ref(range_str);
            CellRange { top: row, left: col, bottom: row, right: col }
        }
    }).collect()
}

/// Split an `<xm:f>` reference like `Sheet1!A1:A10` or `'My Sheet'!$B$3:$B$8`
/// into `(sheet_name, range)`. Returns `None` if the reference has no sheet
/// qualifier — sparkline data refs always do, so unqualified is treated as
/// "same sheet" by callers.
fn split_sheet_ref(s: &str) -> (Option<String>, String) {
    let s = s.trim();
    let Some(bang) = s.rfind('!') else { return (None, s.to_string()); };
    let mut sheet = s[..bang].to_string();
    // Strip absolute-ref dollars from the range part.
    let range = s[bang + 1..].replace('$', "");
    // Quoted sheet names ('foo''s sheet' uses doubled quotes for inner ').
    if sheet.starts_with('\'') && sheet.ends_with('\'') {
        sheet = sheet[1..sheet.len() - 1].replace("''", "'");
    }
    (Some(sheet), range)
}

/// Read a worksheet XML and extract numeric `<v>` values for the cells in
/// `range`. Returns one value per cell in row-major order across the range.
/// Empty cells, non-numeric values, and cells outside the range yield `None`.
///
/// This is intentionally lighter than `parse_row_cells`: sparklines only need
/// raw numbers, no styles, formulas, or shared strings.
fn extract_range_values(sheet_xml: &str, range: &CellRange) -> Vec<Option<f64>> {
    let total = ((range.bottom - range.top + 1) as usize)
        .saturating_mul((range.right - range.left + 1) as usize);
    let mut values: Vec<Option<f64>> = vec![None; total];
    let Ok(doc) = roxmltree::Document::parse(sheet_xml) else { return values; };
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let row_span = (range.right - range.left + 1) as usize;
    for c in doc.descendants().filter(|n| n.tag_name().name() == "c" && n.tag_name().namespace() == Some(ns)) {
        let Some(r_attr) = c.attribute("r") else { continue };
        let (col, row) = parse_cell_ref(r_attr);
        if row < range.top || row > range.bottom || col < range.left || col > range.right {
            continue;
        }
        // Only honor numeric / formula-numeric cells. `t` of "s" / "str" /
        // "inlineStr" / "b" / "e" all map to None for sparkline values.
        let t = c.attribute("t").unwrap_or("");
        if matches!(t, "s" | "str" | "inlineStr" | "b" | "e") { continue; }
        let v = c.children()
            .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(ns))
            .and_then(|n| n.text())
            .and_then(|s| s.trim().parse::<f64>().ok());
        if let Some(num) = v {
            let idx = (row - range.top) as usize * row_span + (col - range.left) as usize;
            if idx < values.len() {
                values[idx] = Some(num);
            }
        }
    }
    values
}

/// Walk the worksheet XML's `<extLst>` and produce one `SparklineGroup` per
/// `<x14:sparklineGroup>`. Resolves cross-sheet `<xm:f>` data references by
/// reading the referenced sheet from the archive (cached per call to avoid
/// re-reads). Theme colors are flattened to `#RRGGBB` via `parse_color`.
fn load_sheet_sparklines(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    sheet_xml: &str,
    sheets: &[SheetMeta],
    rels_doc: &roxmltree::Document,
    theme_colors: &[String],
) -> Vec<SparklineGroup> {
    let Ok(doc) = roxmltree::Document::parse(sheet_xml) else { return Vec::new(); };
    let mut groups: Vec<SparklineGroup> = Vec::new();
    // Cache: sheet name → loaded XML. Saves re-reading when many sparklines
    // reference the same source sheet (typical: one "data" sheet feeds many
    // dashboard sparklines).
    let mut xml_cache: HashMap<String, Option<String>> = HashMap::new();

    let parse_bool_attr = |n: &roxmltree::Node, key: &str, default: bool| -> bool {
        match n.attribute(key) {
            Some(v) => v == "1" || v.eq_ignore_ascii_case("true"),
            None => default,
        }
    };
    let parse_f64_attr = |n: &roxmltree::Node, key: &str| -> Option<f64> {
        n.attribute(key).and_then(|v| v.parse::<f64>().ok())
    };

    for group_node in doc.descendants().filter(|n| n.tag_name().name() == "sparklineGroup") {
        let kind = match group_node.attribute("type").unwrap_or("line") {
            "column" => SparklineType::Column,
            "stacked" => SparklineType::Stem,  // historical alias
            "stem" => SparklineType::Stem,
            // ECMA-376 lists `line` and a planned `stairStep`; treat unknown
            // types as line (closest visual fallback).
            _ => SparklineType::Line,
        };

        let resolve_color = |child_name: &str| -> Option<String> {
            group_node.children()
                .find(|n| n.is_element() && n.tag_name().name() == child_name)
                .and_then(|n| parse_color(&n, theme_colors))
        };

        let mut sparklines: Vec<Sparkline> = Vec::new();
        // <x14:sparklines> is the wrapper; <x14:sparkline> are the children.
        for sparklines_node in group_node.children().filter(|n| n.is_element() && n.tag_name().name() == "sparklines") {
            for sl in sparklines_node.children().filter(|n| n.is_element() && n.tag_name().name() == "sparkline") {
                let f_text = sl.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "f")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                let sqref_text = sl.children()
                    .find(|n| n.is_element() && n.tag_name().name() == "sqref")
                    .and_then(|n| n.text())
                    .unwrap_or("");
                if f_text.is_empty() || sqref_text.is_empty() { continue; }
                let (col, row) = parse_cell_ref(sqref_text.trim());
                let (source_sheet, range_str) = split_sheet_ref(f_text);
                let ranges = parse_sqref(&range_str);
                let Some(range) = ranges.into_iter().next() else { continue };

                // Look up source sheet XML (cross-sheet ref). When the ref
                // has no sheet qualifier, fall back to the *current* sheet
                // XML.
                let source_xml: Option<&str> = match source_sheet {
                    Some(name) => {
                        if !xml_cache.contains_key(&name) {
                            let path = sheets.iter()
                                .find(|s| s.name == name)
                                .and_then(|s| resolve_sheet_path(rels_doc, &s.r_id))
                                .map(|p| format!("xl/{}", p));
                            let xml = path.and_then(|p| read_zip_entry(archive, &p).ok());
                            xml_cache.insert(name.clone(), xml);
                        }
                        xml_cache.get(&name).and_then(|o| o.as_deref())
                    }
                    None => Some(sheet_xml),
                };
                let values = source_xml
                    .map(|xml| extract_range_values(xml, &range))
                    .unwrap_or_default();

                sparklines.push(Sparkline { row, col, values });
            }
        }

        groups.push(SparklineGroup {
            kind,
            markers: parse_bool_attr(&group_node, "markers", false),
            high: parse_bool_attr(&group_node, "high", false),
            low: parse_bool_attr(&group_node, "low", false),
            first: parse_bool_attr(&group_node, "first", false),
            last: parse_bool_attr(&group_node, "last", false),
            negative: parse_bool_attr(&group_node, "negative", false),
            display_x_axis: parse_bool_attr(&group_node, "displayXAxis", false),
            display_empty_cells_as: group_node.attribute("displayEmptyCellsAs").unwrap_or("gap").to_string(),
            min_axis_type: group_node.attribute("minAxisType").unwrap_or("individual").to_string(),
            max_axis_type: group_node.attribute("maxAxisType").unwrap_or("individual").to_string(),
            manual_min: parse_f64_attr(&group_node, "manualMin"),
            manual_max: parse_f64_attr(&group_node, "manualMax"),
            line_weight: parse_f64_attr(&group_node, "lineWeight").unwrap_or(0.75),
            color_series: resolve_color("colorSeries"),
            color_negative: resolve_color("colorNegative"),
            color_axis: resolve_color("colorAxis"),
            color_markers: resolve_color("colorMarkers"),
            color_first: resolve_color("colorFirst"),
            color_last: resolve_color("colorLast"),
            color_high: resolve_color("colorHigh"),
            color_low: resolve_color("colorLow"),
            sparklines,
        });
    }
    groups
}

fn parse_row_cells(
    row_node: &roxmltree::Node,
    shared_strings: &[SharedString],
    theme_colors: &[String],
    ns: &str,
) -> Vec<Cell> {
    let mut cells = Vec::new();
    for c_node in row_node.children() {
        if c_node.tag_name().name() != "c" || c_node.tag_name().namespace() != Some(ns) {
            continue;
        }
        let cell_ref = c_node.attribute("r").unwrap_or("A1").to_string();
        let (col, row) = parse_cell_ref(&cell_ref);
        let cell_type = c_node.attribute("t").unwrap_or("");
        let style_index: u32 = c_node.attribute("s").and_then(|s| s.parse().ok()).unwrap_or(0);

        // Inline string: <c t="inlineStr"><is>...</is></c>
        let is_node = c_node.children().find(|n| n.tag_name().name() == "is");

        // Formula text, if any (<f>…</f>). Kept so the renderer can
        // recompute volatile builtins (TODAY, NOW) at display time.
        let formula: Option<String> = c_node
            .children()
            .find(|n| n.tag_name().name() == "f")
            .and_then(|n| n.text())
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty());

        let v_text = c_node
            .children()
            .find(|n| n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .unwrap_or("")
            .to_string();

        let value = if cell_type == "inlineStr" {
            match is_node {
                Some(is) => {
                    let ss = parse_si_node(&is, ns, theme_colors);
                    CellValue::Text { text: ss.text, runs: ss.runs }
                }
                None => CellValue::Empty,
            }
        } else if v_text.is_empty() {
            CellValue::Empty
        } else {
            match cell_type {
                "s" => {
                    let idx: usize = v_text.parse().unwrap_or(0);
                    if let Some(ss) = shared_strings.get(idx) {
                        CellValue::Text { text: ss.text.clone(), runs: ss.runs.clone() }
                    } else {
                        CellValue::Text { text: String::new(), runs: None }
                    }
                }
                "str" => CellValue::Text { text: v_text, runs: None },
                "b" => CellValue::Bool { bool: v_text == "1" || v_text == "true" },
                "e" => CellValue::Error { error: v_text },
                _ => {
                    if let Ok(n) = v_text.parse::<f64>() {
                        CellValue::Number { number: n }
                    } else {
                        CellValue::Text { text: v_text, runs: None }
                    }
                }
            }
        };

        cells.push(Cell { col, row, col_ref: cell_ref, value, style_index, formula });
    }
    cells
}

fn parse_cell_ref(r: &str) -> (u32, u32) {
    let col_str: String = r.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = r.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    let col = col_str.chars().fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row = row_str.parse().unwrap_or(1);
    (col, row)
}

// ===========================
//  Native (non-WASM) API
// ===========================

/// Returns workbook overview (sheet names and metadata) as JSON.
/// Native equivalent of `parse_xlsx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_workbook_native(data: &[u8]) -> Result<String, String> {
    parse_xlsx_inner(data)
        .and_then(|wb| serde_json::to_string(&wb.workbook).map_err(|e| e.to_string()))
}

/// Parses a single worksheet by 0-based index and returns it as JSON.
/// Native equivalent of `parse_sheet` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_sheet_native(data: &[u8], sheet_index: u32, name: &str) -> Result<String, String> {
    let cursor = Cursor::new(data);
    let mut archive = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let wb_doc = roxmltree::Document::parse(&workbook_xml).map_err(|e| e.to_string())?;
    let sheets = parse_workbook_sheets(&wb_doc);

    let sheet_meta = sheets
        .get(sheet_index as usize)
        .ok_or_else(|| format!("sheet index {} out of range", sheet_index))?;

    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels_doc = roxmltree::Document::parse(&rels_xml).map_err(|e| e.to_string())?;
    let sheet_path = resolve_sheet_path(&rels_doc, &sheet_meta.r_id)
        .ok_or_else(|| format!("rId {} not found in rels", sheet_meta.r_id))?;

    let theme_colors = parse_theme_colors(&mut archive);
    let shared_strings = read_shared_strings(&mut archive, &theme_colors);
    let sheet_xml = read_zip_entry(&mut archive, &format!("xl/{}", sheet_path))?;
    let (mut ws, hyperlink_rids) =
        parse_worksheet(&sheet_xml, &shared_strings, &theme_colors, name)
            .map_err(|e| e.to_string())?;

    ws.images = load_sheet_images(&mut archive, &sheet_path);
    ws.charts = load_sheet_charts(&mut archive, &sheet_path, &theme_colors);
    ws.shape_groups = load_sheet_shape_groups(&mut archive, &sheet_path, &theme_colors);
    ws.hyperlinks = load_hyperlinks(&mut archive, &sheet_path, hyperlink_rids);
    ws.comment_refs = load_sheet_comment_refs(&mut archive, &sheet_path);
    ws.defined_names = parse_defined_names_for_sheet(&wb_doc, sheet_index);
    ws.tables = load_sheet_tables(&mut archive, &sheet_path, &theme_colors);
    ws.slicers = load_sheet_slicers(&mut archive, &sheet_path);
    ws.sparkline_groups = load_sheet_sparklines(&mut archive, &sheet_xml, &sheets, &rels_doc, &theme_colors);

    serde_json::to_string(&ws).map_err(|e| e.to_string())
}
