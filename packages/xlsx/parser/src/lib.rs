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
    /// Explicit fill color hex (from c:spPr/a:solidFill/a:srgbClr). None = use palette.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Whether to draw data-point markers on line/scatter series. Resolved at
    /// parse time from `<c:ser><c:marker><c:symbol val>` (ECMA-376 §21.2.2.32)
    /// falling back to the chart-type-level `<c:lineChart><c:marker val>`
    /// (§21.2.2.33). Absent markers default to hidden for line charts.
    pub show_marker: bool,
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

    // Find c:plotArea
    let plot_area = chart_root.children()
        .find(|n| n.tag_name().name() == "plotArea" && n.tag_name().namespace() == Some(c_ns))?;

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
                continue;
            }
            "valAx" => {
                if val_axis_title.is_none() {
                    val_axis_title = extract_chart_title(&child, c_ns, a_ns);
                }
                if val_axis_font_size_hpt.is_none() {
                    val_axis_font_size_hpt = extract_axis_tick_label_size(&child, c_ns, a_ns);
                }
                if val_axis_format_code.is_none() {
                    val_axis_format_code = child.children()
                        .find(|n| n.tag_name().name() == "numFmt" && n.tag_name().namespace() == Some(c_ns))
                        .and_then(|n| n.attribute("formatCode").map(|s| s.to_string()))
                        .filter(|s| !s.is_empty() && s != "General");
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
                "dLbls"    => {
                    for d in attr_node.children().filter(|n| n.is_element()) {
                        match d.tag_name().name() {
                            "showVal" | "showPercent" => {
                                if d.attribute("val").unwrap_or("1") != "0" {
                                    show_data_labels = true;
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
        }
    }

    if primary_type.is_empty() { return None; }

    // Fill in categories for series that have none (mixed charts share categories)
    for s in &mut all_series {
        if s.categories.is_empty() {
            s.categories = shared_categories.clone();
        }
    }

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
        cat_axis_font_size_hpt,
        val_axis_font_size_hpt,
        val_axis_format_code,
    })
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
/// Resolve the fill color under `c:spPr/a:solidFill` for a chart series.
/// Supports `a:srgbClr` (explicit hex) and `a:schemeClr` (theme accent/dark/light).
/// Theme colors use drawingML names (`accent1`..`accent6`, `dk1`/`dk2`/`lt1`/`lt2`)
/// which map to the parser's natural-order theme array (accent_n at index 3+n,
/// dk1@0, lt1@1, dk2@2, lt2@3).
fn resolve_series_color(node: &roxmltree::Node, theme_colors: &[String]) -> Option<String> {
    for n in node.descendants() {
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

fn parse_chart_series(
    node: &roxmltree::Node,
    c_ns: &str,
    ser_type: &str,
    chart_marker_default: bool,
    theme_colors: &[String],
) -> ChartSeries {
    let name = extract_series_name(node, c_ns);

    // For scatter: xVal → categories (as strings), yVal → values
    // For others:  cat  → categories,             val  → values
    let (cat_tag, val_tag) = if ser_type == "scatter" { ("xVal", "yVal") } else { ("cat", "val") };

    let categories = collect_str_cache(node, c_ns, cat_tag);
    let values     = collect_num_cache(node, c_ns, val_tag);

    // Series fill color from c:spPr/a:solidFill (supports a:srgbClr and a:schemeClr).
    // For schemeClr, resolves "accentN"/"dk1"/etc. against the workbook theme.
    let color = resolve_series_color(node, theme_colors);

    // Marker visibility (ECMA-376 §21.2.2.32 — c:marker/c:symbol default is
    // "none"). A per-series <c:marker><c:symbol> overrides; otherwise fall
    // back to the chart-type-level <c:lineChart><c:marker val> flag. Scatter
    // charts default to visible markers even without an explicit flag.
    let show_marker = if let Some(mk) = node.children()
        .find(|n| n.tag_name().name() == "marker" && n.tag_name().namespace() == Some(c_ns))
    {
        match mk.children().find(|n| n.tag_name().name() == "symbol" && n.tag_name().namespace() == Some(c_ns)) {
            Some(sym) => sym.attribute("val").map(|v| v != "none").unwrap_or(true),
            None => true,
        }
    } else if ser_type == "scatter" {
        true
    } else {
        chart_marker_default
    };

    ChartSeries {
        name,
        series_type: ser_type.to_string(),
        categories,
        values,
        color,
        show_marker,
    }
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
