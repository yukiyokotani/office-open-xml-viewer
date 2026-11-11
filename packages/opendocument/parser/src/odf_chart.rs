//! ODF charts (Part 3 §11 `chart:chart`, in an embedded object's
//! `content.xml`) as the shared OOXML chart model
//! (`ooxml_common::chart::ChartModel`) that the XLSX, PPTX and DOCX renderers
//! draw.
//!
//! Shared by the spreadsheet and presentation readers. A caller passes the
//! object's package directory (the `draw:object` `xlink:href` without `./`,
//! e.g. `Object 1`) and, for a spreadsheet, a [`RangeResolver`] over its
//! tables. Data comes from the cell ranges each series names
//! (`chart:values-cell-range-address`, `chart:categories`,
//! `chart:domain`); a range on the chart's own `local-table` (internal data,
//! as in presentations) and any range the resolver cannot read are taken
//! from that local table, which LibreOffice always writes as the chart's
//! data cache.
//!
//! ```ignore
//! let chart = odf_chart::read_chart_object(&package, "Object 1", None)?;
//! ```

use crate::odf_props::{
    self, get, get_element, hex_color, length_pt, percent, Style, StyleStore, NS_CHART,
};
use crate::ods_number::DataStyles;
use crate::package::OdfPackage;
use crate::xml::{self, NS_OFFICE, NS_TABLE, NS_TEXT};
use ooxml_common::chart::{
    canonical_chart_type, ChartModel, ChartSeries, ChartSeriesDataLabels, ChartThreeD,
};
use std::collections::HashMap;

/// Implementation bounds on one chart: series, points per series.
const MAX_SERIES: usize = 4096;
const MAX_POINTS: usize = 1_048_576;
/// Implementation bound on the local table's cells.
const MAX_LOCAL_CELLS: usize = 4_000_000;
/// Implementation bound on the data points one chart reads over all of its
/// ranges (series values, categories, labels).
const MAX_CHART_POINTS: usize = 4_000_000;
/// Implementation bound on the bytes one chart's reader holds: data-table
/// cells (each repeat), resolved range records, text, per-series copies of
/// categories and repeated point colours. Charged before each is kept.
const MAX_CHART_BYTES: usize = 128 * 1024 * 1024;

/// The per-chart budgets. Exhausting one stops reading further data (every
/// later charge fails) and makes the chart an error, so a reader never does
/// unbounded work for data it would not keep.
struct ChartBudget {
    points: std::cell::Cell<usize>,
    bytes: std::cell::Cell<usize>,
    exhausted: std::cell::Cell<bool>,
}

impl ChartBudget {
    fn new() -> Self {
        Self {
            points: std::cell::Cell::new(MAX_CHART_POINTS),
            bytes: std::cell::Cell::new(MAX_CHART_BYTES),
            exhausted: std::cell::Cell::new(false),
        }
    }
    /// Charge `bytes`; `false` (and exhausted) past the limit.
    fn charge(&self, bytes: usize) -> bool {
        match self.bytes.get().checked_sub(bytes) {
            Some(left) if !self.exhausted.get() => {
                self.bytes.set(left);
                true
            }
            _ => {
                self.exhausted.set(true);
                false
            }
        }
    }
    /// Charge one resolved data point holding `text_len` bytes of text.
    fn point(&self, text_len: usize) -> bool {
        match self.points.get().checked_sub(1) {
            Some(left) if self.charge(POINT_BYTES.saturating_add(text_len)) => {
                self.points.set(left);
                true
            }
            _ => {
                self.exhausted.set(true);
                false
            }
        }
    }
    fn check(&self) -> Result<(), String> {
        if self.exhausted.get() {
            Err("OpenDocument chart exceeds the data limit".into())
        } else {
            Ok(())
        }
    }
}

/// Charge for reading one axis (its style, title and model record).
const AXIS_BYTES: usize = 4096;

/// Implementation bound on the regression curves of one series.
const MAX_TRENDLINES: usize = 64;

/// Fixed size of one resolved data point record.
const POINT_BYTES: usize = std::mem::size_of::<(CellData, String)>();
/// Fixed size of one source-range map entry (key, value, table slot).
const SOURCE_ENTRY_BYTES: usize = std::mem::size_of::<(String, (usize, usize))>() * 2;

/// The space-separated cell range addresses of a list (§18.3.6), read
/// one at a time (spaces inside quoted sheet names are kept).
fn addresses(list: &str) -> impl Iterator<Item = &str> {
    let mut rest = list;
    std::iter::from_fn(move || {
        rest = rest.trim_start();
        if rest.is_empty() {
            return None;
        }
        let mut quoted = false;
        let mut end = rest.len();
        for (i, c) in rest.char_indices() {
            match c {
                '\'' => quoted = !quoted,
                c if c.is_whitespace() && !quoted => {
                    end = i;
                    break;
                }
                _ => {}
            }
        }
        let (part, tail) = rest.split_at(end);
        rest = tail;
        Some(part)
    })
}

/// LibreOffice's default series palette (used for series without a fill).
const DEFAULT_COLORS: [&str; 12] = [
    "004586", "FF420E", "FFD320", "579D1C", "7E0021", "83CAFF", "314004", "AECF00", "4B1F6F",
    "FF950E", "C5000B", "0084D1",
];

/// One cell of a range, as a chart reads it.
#[derive(Clone, Debug, PartialEq)]
pub enum CellData {
    Empty,
    Number(f64),
    Text(String),
}

/// Cell ranges outside the chart (a spreadsheet's tables).
pub trait RangeResolver {
    /// The cells of an ODF range address (`Sheet1.A1:Sheet1.A9`) in row-major
    /// order, or `None` when the address names no readable range.
    fn cells(&self, address: &str) -> Option<Vec<CellData>>;
    /// Hidden flags of the same cells, when known.
    fn hidden(&self, address: &str) -> Option<Vec<bool>>;
}

/// Read the chart object stored under `object_dir` (`Object 1`). `Ok(None)`
/// when the object is not a chart (another embedded document type).
pub fn read_chart_object(
    package: &OdfPackage,
    object_dir: &str,
    resolver: Option<&dyn RangeResolver>,
) -> Result<Option<ChartModel>, String> {
    let dir = object_dir.trim_start_matches("./").trim_end_matches('/');
    let content_path = format!("{dir}/content.xml");
    if !package.contains(&content_path) {
        return Ok(None);
    }
    let content = package.read_text(&content_path)?;
    let styles_path = format!("{dir}/styles.xml");
    let styles = if package.contains(&styles_path) {
        Some(package.read_text(&styles_path)?)
    } else {
        None
    };
    read_chart_document(&content, styles.as_deref(), resolver)
}

/// Read a chart from the object's `content.xml` (and `styles.xml`) text.
pub fn read_chart_document(
    content: &str,
    styles: Option<&str>,
    resolver: Option<&dyn RangeResolver>,
) -> Result<Option<ChartModel>, String> {
    let content_doc = xml::parse(content)?;
    let styles_doc = match styles {
        Some(text) => Some(xml::parse(text)?),
        None => None,
    };
    let mut store = StyleStore::default();
    let mut data = DataStyles::default();
    let mut faces = HashMap::new();
    if let Some(doc) = &styles_doc {
        odf_props::font_faces(doc.root_element(), &mut faces);
        for child in doc.root_element().children().filter(|n| n.is_element()) {
            if xml::is(child, NS_OFFICE, "styles") {
                store.add(child, false)?;
                data.add(child);
            }
        }
    }
    let root = content_doc.root_element();
    odf_props::font_faces(root, &mut faces);
    for child in root
        .children()
        .filter(|n| xml::is(*n, NS_OFFICE, "automatic-styles"))
    {
        store.add(child, true)?;
        data.add(child);
    }
    let Some(chart) = root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|b| b.children().find(|n| xml::is(*n, NS_OFFICE, "chart")))
        .and_then(|c| c.children().find(|n| xml::is(*n, NS_CHART, "chart")))
    else {
        return Ok(None);
    };
    let budget = ChartBudget::new();
    let local = LocalTable::read(chart, &budget)?;
    budget.check()?;
    let reader = Reader {
        store: &store,
        faces: &faces,
        local,
        resolver,
        budget,
        dashes: styles_doc
            .as_ref()
            .map(|doc| stroke_dashes(doc.root_element()))
            .unwrap_or_default(),
    };
    let model = reader.chart(chart, &mut data)?;
    reader.budget.check()?;
    Ok(Some(model))
}

struct Reader<'a> {
    store: &'a StyleStore,
    faces: &'a HashMap<String, String>,
    local: LocalTable,
    resolver: Option<&'a dyn RangeResolver>,
    budget: ChartBudget,
    /// `draw:stroke-dash` definitions (§16.41) as DrawingML presets.
    dashes: HashMap<String, &'static str>,
}

/// The chart's `local-table` (§11.1 data table): a grid with the header
/// row and column included, plus each cell's `svg:desc` source range.
#[derive(Default)]
struct LocalTable {
    cells: Vec<Vec<(CellData, String)>>,
    /// Source range address → (column, first data row).
    sources: HashMap<String, (usize, usize)>,
}

impl LocalTable {
    fn read(chart: roxmltree::Node<'_, '_>, budget: &ChartBudget) -> Result<Self, String> {
        let mut table = Self::default();
        let Some(node) = chart.children().find(|n| xml::is(*n, NS_TABLE, "table")) else {
            return Ok(table);
        };
        let mut count = 0usize;
        for row in node
            .descendants()
            .filter(|n| xml::is(*n, NS_TABLE, "table-row"))
        {
            let mut cells = Vec::new();
            for cell in row
                .children()
                .filter(|n| xml::is(*n, NS_TABLE, "table-cell"))
            {
                let repeat = cell
                    .attribute((NS_TABLE, "number-columns-repeated"))
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1)
                    .clamp(1, 1024);
                let text = chart_text(cell, budget)?;
                let value = match cell.attribute((NS_OFFICE, "value-type")) {
                    Some("float") | Some("percentage") | Some("currency") => cell
                        .attribute((NS_OFFICE, "value"))
                        .and_then(|v| v.parse::<f64>().ok())
                        .map_or(CellData::Empty, CellData::Number),
                    Some("string") => CellData::Text(text.clone()),
                    _ if !text.is_empty() => CellData::Text(text.clone()),
                    _ => CellData::Empty,
                };
                let desc = cell
                    .descendants()
                    .find(|n| xml::is(*n, crate::xml::NS_SVG, "desc"))
                    .and_then(|n| n.text());
                if let Some(desc) = desc {
                    for range in addresses(desc) {
                        // The owned key and its map entry.
                        if !budget.charge(range.len().saturating_add(SOURCE_ENTRY_BYTES)) {
                            return Err("OpenDocument chart exceeds the data limit".into());
                        }
                        table
                            .sources
                            .insert(range.to_string(), (cells.len(), table.cells.len()));
                    }
                }
                // Every repeat keeps its own copy of the value and the text.
                let copy = POINT_BYTES.saturating_add(text.len().saturating_mul(2));
                if !budget.charge(copy.saturating_mul(repeat)) {
                    return Err("OpenDocument chart exceeds the data limit".into());
                }
                for _ in 0..repeat {
                    cells.push((value.clone(), text.clone()));
                }
                count += repeat;
                if count > MAX_LOCAL_CELLS {
                    return Err("OpenDocument chart exceeds the data table limit".into());
                }
            }
            table.cells.push(cells);
        }
        Ok(table)
    }

    fn get(&self, row: usize, col: usize) -> Option<&(CellData, String)> {
        self.cells.get(row)?.get(col)
    }

    /// The local-table cells of a range on `local-table` (A1 = top-left),
    /// each charged to the chart's budget before it is copied.
    fn range(&self, address: &str, budget: &ChartBudget) -> Option<Vec<(CellData, String)>> {
        let (sheet, (r1, c1), (r2, c2)) = crate::odf_text::parse_range(address)?;
        if sheet.as_deref().is_some_and(|s| s != "local-table") {
            return None;
        }
        let mut out = Vec::new();
        for row in r1..=r2 {
            for col in c1..=c2 {
                if out.len() >= MAX_POINTS {
                    return None;
                }
                let cell = self.get(row as usize - 1, col as usize - 1);
                if !budget.point(cell.map_or(0, |(_, t)| t.len())) {
                    return Some(out);
                }
                out.push(cell.cloned().unwrap_or((CellData::Empty, String::new())));
            }
        }
        Some(out)
    }

    /// Cached cells of an external source range, via the `svg:desc`
    /// annotations LibreOffice writes on the first cell of each column.
    fn cached(&self, address: &str, budget: &ChartBudget) -> Option<Vec<(CellData, String)>> {
        let (col, row) = *self.sources.get(address)?;
        let (_, (r1, c1), (r2, c2)) = crate::odf_text::parse_range(address)?;
        // The address is file data (up to u32 rows): the count is bounded
        // before anything is collected, as a local range is.
        let span = |a: u32, b: u32| usize::try_from(b - a).unwrap_or(usize::MAX);
        let count = span(r1, r2)
            .max(span(c1, c2))
            .saturating_add(1)
            .min(MAX_POINTS);
        let horizontal = r1 == r2 && c2 > c1;
        let mut out = Vec::new();
        for i in 0..count {
            let (r, c) = if horizontal {
                (row, col.saturating_add(i))
            } else {
                (row.saturating_add(i), col)
            };
            let cell = self.get(r, c);
            if !budget.point(cell.map_or(0, |(_, t)| t.len())) {
                break;
            }
            out.push(cell.cloned().unwrap_or((CellData::Empty, String::new())));
        }
        Some(out)
    }
}

/// The text of a chart text element or data-table cell (paragraphs joined
/// by line breaks, `text:s` expanded), charged to the chart's budget as it
/// is built; `Err` past it.
fn chart_text(node: roxmltree::Node<'_, '_>, budget: &ChartBudget) -> Result<String, String> {
    let mut out = String::new();
    let mut first = true;
    for p in node
        .children()
        .filter(|n| xml::is(*n, crate::xml::NS_TEXT, "p") || xml::is(*n, crate::xml::NS_TEXT, "h"))
    {
        if !first {
            push_chart_text(&mut out, "\n", 1, budget)?;
        }
        first = false;
        chart_inline_text(p, &mut out, budget, 0)?;
    }
    Ok(out)
}

/// Bound on span nesting inside chart text.
const MAX_CHART_TEXT_DEPTH: usize = 64;

fn chart_inline_text(
    node: roxmltree::Node<'_, '_>,
    out: &mut String,
    budget: &ChartBudget,
    depth: usize,
) -> Result<(), String> {
    if depth > MAX_CHART_TEXT_DEPTH {
        return Err("OpenDocument chart text is nested too deeply".into());
    }
    for child in node.children() {
        if child.is_text() {
            push_chart_text(out, child.text().unwrap_or(""), 1, budget)?;
            continue;
        }
        if child.tag_name().namespace() != Some(crate::xml::NS_TEXT) {
            continue;
        }
        match child.tag_name().name() {
            "s" => {
                let count = child
                    .attribute((crate::xml::NS_TEXT, "c"))
                    .and_then(|c| c.parse::<usize>().ok())
                    .unwrap_or(1)
                    .min(4096);
                push_chart_text(out, " ", count, budget)?;
            }
            "tab" => push_chart_text(out, "\t", 1, budget)?,
            "line-break" => push_chart_text(out, "\n", 1, budget)?,
            "note" | "ruby-text" => {}
            _ => chart_inline_text(child, out, budget, depth + 1)?,
        }
    }
    Ok(())
}

fn push_chart_text(
    out: &mut String,
    piece: &str,
    times: usize,
    budget: &ChartBudget,
) -> Result<(), String> {
    if !budget.charge(piece.len().saturating_mul(times)) {
        return Err("OpenDocument chart exceeds the data limit".into());
    }
    for _ in 0..times {
        out.push_str(piece);
    }
    Ok(())
}

fn number_text(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        let text = format!("{value}");
        text
    }
}

fn hpt(pt: f64) -> i32 {
    (pt * 100.0).round() as i32
}

fn emu(pt: f64) -> u32 {
    (pt * 12_700.0).round().max(0.0) as u32
}

/// Text properties of a chart element style.
struct TextFont {
    size: Option<i32>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<String>,
    face: Option<String>,
}

/// Fill and stroke of a chart element style (§20.130, §20.164).
struct Paint {
    fill: Option<String>,
    fill_hidden: bool,
    line: Option<String>,
    line_hidden: bool,
    line_width: Option<u32>,
    dash: Option<&'static str>,
}

impl Reader<'_> {
    fn chain(&self, node: roxmltree::Node<'_, '_>) -> Vec<&Style> {
        self.store
            .chain("chart", node.attribute((NS_CHART, "style-name")))
    }

    /// A text style's font; the face it copies is charged first (an
    /// exhausted budget leaves the face unset and fails the chart).
    fn font(&self, chain: &[&Style]) -> TextFont {
        let face: Option<&str> = get(chain, "text:style:font-name")
            .map(|name| self.faces.get(name).map_or(name, String::as_str))
            .or_else(|| get(chain, "text:fo:font-family"));
        let face = face
            .filter(|f| {
                self.budget
                    .charge(std::mem::size_of::<TextFont>().saturating_add(f.len()))
            })
            .map(|f| {
                if get(chain, "text:style:font-name").is_some() {
                    f.to_string()
                } else {
                    odf_props::unquote_family(f)
                }
            });
        TextFont {
            size: get(chain, "text:fo:font-size").and_then(length_pt).map(hpt),
            bold: get(chain, "text:fo:font-weight")
                .map(|w| w == "bold" || w.parse::<u32>().is_ok_and(|v| v >= 600)),
            italic: get(chain, "text:fo:font-style").map(|s| s == "italic" || s == "oblique"),
            color: get(chain, "text:fo:color").and_then(hex_color),
            face,
        }
    }

    fn paint(&self, chain: &[&Style]) -> Paint {
        let fill_kind = get(chain, "graphic:draw:fill");
        let stroke = get(chain, "graphic:draw:stroke");
        Paint {
            fill: if fill_kind == Some("none") {
                None
            } else {
                get(chain, "graphic:draw:fill-color").and_then(hex_color)
            },
            fill_hidden: fill_kind == Some("none"),
            line: if stroke == Some("none") {
                None
            } else {
                get(chain, "graphic:svg:stroke-color").and_then(hex_color)
            },
            line_hidden: stroke == Some("none"),
            line_width: get(chain, "graphic:svg:stroke-width")
                .and_then(length_pt)
                .map(emu),
            dash: (stroke == Some("dash")).then(|| {
                get(chain, "graphic:draw:stroke-dash")
                    .and_then(|name| self.dashes.get(name).copied())
                    .unwrap_or("dash")
            }),
        }
    }

    /// The text of a title-like element (`text:p` children).
    /// A text element's text, charged to the chart's budget (empty once
    /// the budget is exhausted; the chart then fails).
    fn text(&self, node: roxmltree::Node<'_, '_>) -> String {
        chart_text(node, &self.budget).unwrap_or_default()
    }

    /// A copy of a range's texts (a series' own categories), charged first.
    fn copy_texts(&self, cells: &[(CellData, String)]) -> Option<Vec<String>> {
        let bytes = cells.iter().fold(0usize, |sum, (_, t)| {
            sum.saturating_add(std::mem::size_of::<String>() + t.len())
        });
        self.budget
            .charge(bytes)
            .then(|| cells.iter().map(|(_, t)| t.clone()).collect())
    }

    /// Resolve a range: the resolver, else the local table.
    fn range(&self, address: &str) -> Vec<(CellData, String)> {
        // Several space-separated ranges are concatenated (§18.3.5).
        let mut out = Vec::new();
        for part in addresses(address) {
            let local_first =
                part.starts_with("local-table.") || part.starts_with("'local-table'.");
            if self.budget.exhausted.get() {
                break;
            }
            let resolved = if local_first {
                self.local.range(part, &self.budget)
            } else {
                self.resolver
                    .and_then(|r| r.cells(part))
                    .map(|cells| {
                        let cached = self.local.cached(part, &self.budget);
                        cells
                            .into_iter()
                            .enumerate()
                            .take_while(|(_, cell)| {
                                let len = match cell {
                                    CellData::Text(t) => t.len(),
                                    _ => 0,
                                };
                                self.budget.point(len)
                            })
                            .map(|(i, cell)| {
                                // Display text from the cache when the source
                                // value is a number (formatted categories).
                                let text = match &cell {
                                    CellData::Text(t) => t.clone(),
                                    CellData::Number(n) => cached
                                        .as_ref()
                                        .and_then(|c| c.get(i))
                                        .map(|(_, t)| t.clone())
                                        .filter(|t| !t.is_empty())
                                        .unwrap_or_else(|| number_text(*n)),
                                    CellData::Empty => String::new(),
                                };
                                (cell, text)
                            })
                            .collect()
                    })
                    .or_else(|| self.local.cached(part, &self.budget))
                    .or_else(|| self.local.range(part, &self.budget))
            };
            out.extend(resolved.unwrap_or_default());
            if out.len() > MAX_POINTS {
                out.truncate(MAX_POINTS);
                break;
            }
        }
        out
    }

    fn chart(
        &self,
        chart: roxmltree::Node<'_, '_>,
        data: &mut DataStyles,
    ) -> Result<ChartModel, String> {
        let class = chart
            .attribute((NS_CHART, "class"))
            .unwrap_or("chart:bar")
            .trim_start_matches("chart:");
        let plot = chart
            .children()
            .find(|n| xml::is(*n, NS_CHART, "plot-area"));
        let plot_chain = plot.map(|p| self.chain(p)).unwrap_or_default();
        let flag = |key: &str| get(&plot_chain, &format!("chart:chart:{key}")) == Some("true");
        let horizontal = flag("vertical");
        let grouping = if flag("percentage") {
            "percentStacked"
        } else if flag("stacked") {
            "stacked"
        } else {
            "clustered"
        };
        let family = |class: &str| -> &'static str {
            match class {
                "bar" | "gantt" => "bar",
                "line" => "line",
                "area" => "area",
                "circle" => "pie",
                "ring" => "doughnut",
                "scatter" => "scatter",
                "bubble" => "bubble",
                "radar" | "filled-radar" => "radar",
                "stock" => "stock",
                "surface" => "surface",
                _ => "bar",
            }
        };
        let primary = family(class);
        let chart_type = match primary {
            "bar" | "line" | "area" => canonical_chart_type(
                primary,
                if horizontal { "bar" } else { "col" },
                if primary == "bar" {
                    grouping
                } else if grouping == "clustered" {
                    "standard"
                } else {
                    grouping
                },
            ),
            other => other.to_string(),
        };
        let mut model = ChartModel {
            chart_type,
            cat_axis_cross_between: if matches!(primary, "bar") {
                "between".into()
            } else {
                "midCat".into()
            },
            val_axis_major_tick_mark: "out".into(),
            cat_axis_major_tick_mark: "out".into(),
            plot_visible_only: Some(
                get(&plot_chain, "chart:chart:include-hidden-cells") != Some("true"),
            ),
            ..ChartModel::default()
        };
        if matches!(primary, "line" | "area") && grouping != "clustered" {
            // Category axis crossing for stacked line/area is unchanged.
        }
        if primary == "radar" {
            model.radar_style = Some(
                if class == "filled-radar" {
                    "filled"
                } else {
                    "marker"
                }
                .into(),
            );
        }
        if primary == "scatter" {
            model.scatter_style = Some(
                if get(&plot_chain, "chart:chart:lines") == Some("false") {
                    "marker"
                } else {
                    "lineMarker"
                }
                .into(),
            );
        }
        if flag("three-dimensional") {
            model.three_d = Some(ChartThreeD {
                view_3d_present: Some(true),
                ..ChartThreeD::default()
            });
        }
        model.disp_blanks_as = match get(&plot_chain, "chart:chart:treat-empty-cells") {
            Some("use-zero") => Some("zero".into()),
            Some("ignore") => Some("span".into()),
            Some("leave-gap") => Some("gap".into()),
            _ => None,
        };
        // Chart area (§11.1 chart:chart style).
        let area = self.paint(&self.chain(chart));
        model.chart_bg = area.fill;
        model.chart_fill_hidden = area.fill_hidden.then_some(true);
        model.chart_border_color = area.line;
        model.chart_border_width_emu = area.line_width;
        model.chart_border_hidden = area.line_hidden.then_some(true);
        // Title.
        if let Some(title) = chart.children().find(|n| xml::is(*n, NS_CHART, "title")) {
            let text = self.text(title);
            let font = self.font(&self.chain(title));
            let span_font = title
                .descendants()
                .find(|n| xml::is(*n, NS_TEXT, "span"))
                .and_then(|s| s.attribute((NS_TEXT, "style-name")))
                .map(|name| self.font(&self.store.chain("text", Some(name))));
            let font = merge_font(font, span_font);
            model.title_present = true;
            model.title = Some(text);
            model.title_font_size_hpt = font.size;
            model.title_font_bold = font.bold;
            model.title_font_italic = font.italic;
            model.title_font_color = font.color;
            model.title_font_face = font.face;
        }
        // Legend.
        if let Some(legend) = chart.children().find(|n| xml::is(*n, NS_CHART, "legend")) {
            model.show_legend = true;
            model.legend_pos = Some(
                match legend.attribute((NS_CHART, "legend-position")) {
                    Some("start") => "l",
                    Some("top") => "t",
                    Some("bottom") => "b",
                    Some("top-end") => "tr",
                    Some("top-start") => "t",
                    Some("bottom-start") | Some("bottom-end") => "b",
                    _ => "r",
                }
                .into(),
            );
            let chain = self.chain(legend);
            let font = self.font(&chain);
            model.legend_font_size_hpt = font.size;
            model.legend_font_bold = font.bold;
            model.legend_font_italic = font.italic;
            model.legend_font_color = font.color;
            model.legend_font_face = font.face;
            let paint = self.paint(&chain);
            model.legend_fill_color = paint.fill;
            model.legend_fill_hidden = paint.fill_hidden.then_some(true);
            model.legend_line_color = paint.line;
            model.legend_line_hidden = paint.line_hidden.then_some(true);
        }
        let Some(plot) = plot else {
            model.authored_without_series = true;
            return Ok(model);
        };
        // Plot area background is the wall (§11.6).
        if let Some(wall) = plot.children().find(|n| xml::is(*n, NS_CHART, "wall")) {
            let paint = self.paint(&self.chain(wall));
            model.plot_area_bg = paint.fill;
            model.plot_area_fill_hidden = paint.fill_hidden.then_some(true);
            model.plot_area_line_color = paint.line;
            model.plot_area_line_width_emu = paint.line_width;
            model.plot_area_line_hidden = paint.line_hidden.then_some(true);
        }
        // Pie geometry: angle offset (§20.17 chart:angle-offset, degrees
        // counter-clockwise from 3 o'clock) as the first-slice angle
        // (degrees clockwise from 12 o'clock).
        if matches!(primary, "pie" | "doughnut") {
            let offset = get(&plot_chain, "chart:chart:angle-offset")
                .and_then(|v| v.parse::<f64>().ok())
                .unwrap_or(90.0);
            model.first_slice_angle = Some(((90.0 - offset).rem_euclid(360.0)).round() as u32);
            if primary == "doughnut" {
                model.hole_size = get(&plot_chain, "chart:chart:hole-size")
                    .or_else(|| get(&plot_chain, "chart:loext:hole-size"))
                    .and_then(percent)
                    .map(|v| (v * 100.0).round() as u32)
                    .or(Some(50));
            }
        }
        self.axes(plot, &mut model, data);
        self.series(plot, &plot_chain, class, primary, &mut model, data)?;
        Ok(model)
    }

    fn axes(&self, plot: roxmltree::Node<'_, '_>, model: &mut ChartModel, data: &mut DataStyles) {
        for axis in plot.children().filter(|n| xml::is(*n, NS_CHART, "axis")) {
            // Each axis record read (later axes of a dimension replace it).
            if !self.budget.charge(AXIS_BYTES) {
                break;
            }
            let dimension = axis.attribute((NS_CHART, "dimension")).unwrap_or("");
            let name = axis.attribute((NS_CHART, "name")).unwrap_or("");
            let chain = self.chain(axis);
            let prop = |key: &str| get(&chain, &format!("chart:chart:{key}"));
            let number = |key: &str| prop(key).and_then(|v| v.parse::<f64>().ok());
            let font = self.font(&chain);
            let paint = self.paint(&chain);
            let visible = prop("visible") != Some("false");
            let labels = prop("display-label") != Some("false");
            let tick = |inner: &str, outer: &str| -> String {
                match (prop(inner) == Some("true"), prop(outer) == Some("true")) {
                    (true, true) => "cross",
                    (true, false) => "in",
                    (false, true) => "out",
                    (false, false) => "none",
                }
                .into()
            };
            let major_tick = tick("tick-marks-major-inner", "tick-marks-major-outer");
            let minor_tick = tick("tick-marks-minor-inner", "tick-marks-minor-outer");
            let title = axis
                .children()
                .find(|n| xml::is(*n, NS_CHART, "title"))
                .map(|t| (self.text(t), self.font(&self.chain(t))));
            let grid = |class: &str| {
                axis.children()
                    .find(|n| {
                        xml::is(*n, NS_CHART, "grid")
                            && n.attribute((NS_CHART, "class")).unwrap_or("major") == class
                    })
                    .map(|g| self.paint(&self.chain(g)))
            };
            let format_code = chain
                .iter()
                .find_map(|s| s.data_style.as_deref())
                .filter(|_| prop("link-data-style-to-source") != Some("true"))
                .and_then(|name| data.code(name))
                .map(|(code, _)| crate::ods_number::literal_code(&code))
                .filter(|code| code != "General");
            let secondary = name.starts_with("secondary");
            match (dimension, secondary) {
                ("x", false) => {
                    model.cat_axis_hidden = !visible;
                    model.cat_axis_line_hidden = paint.line_hidden;
                    model.cat_axis_line_color = paint.line;
                    model.cat_axis_line_width_emu = paint.line_width;
                    model.cat_axis_major_tick_mark = major_tick;
                    model.cat_axis_minor_tick_mark = Some(minor_tick);
                    if !labels {
                        model.cat_axis_tick_label_pos = Some("none".into());
                    }
                    model.cat_axis_font_size_hpt = font.size;
                    model.cat_axis_font_bold = font.bold;
                    model.cat_axis_font_italic = font.italic;
                    model.cat_axis_font_color = font.color;
                    model.cat_axis_font_face = font.face;
                    model.cat_axis_orientation =
                        (prop("reverse-direction") == Some("true")).then(|| "maxMin".into());
                    model.cat_axis_min = number("minimum");
                    model.cat_axis_max = number("maximum");
                    model.cat_axis_major_unit = number("interval-major");
                    model.cat_axis_log_base = (prop("logarithmic") == Some("true")).then_some(10.0);
                    model.cat_axis_format_code = format_code;
                    if let Some(rotation) = get(&chain, "chart:style:rotation-angle")
                        .and_then(|v| v.parse::<f64>().ok())
                        .filter(|r| *r != 0.0)
                    {
                        model.cat_axis_label_rotation = Some((-rotation * 60_000.0).round() as i32);
                    }
                    if let Some((text, font)) = title {
                        model.cat_axis_title = Some(text);
                        model.cat_axis_title_font_size_hpt = font.size;
                        model.cat_axis_title_font_bold = font.bold;
                        model.cat_axis_title_font_color = font.color;
                    }
                    if let Some(grid) = grid("major") {
                        model.cat_axis_major_gridlines = Some(true);
                        model.cat_axis_gridline_color = grid.line;
                        model.cat_axis_gridline_width_emu = grid.line_width;
                    }
                    if let Some(grid) = grid("minor") {
                        model.cat_axis_minor_gridlines = Some(true);
                        model.cat_axis_minor_gridline_color = grid.line;
                    }
                }
                ("y", false) => {
                    model.val_axis_hidden = !visible;
                    model.val_axis_line_hidden = paint.line_hidden;
                    model.val_axis_line_color = paint.line;
                    model.val_axis_line_width_emu = paint.line_width;
                    model.val_axis_major_tick_mark = major_tick;
                    model.val_axis_minor_tick_mark = Some(minor_tick);
                    if !labels {
                        model.val_axis_tick_label_pos = Some("none".into());
                    }
                    model.val_axis_font_size_hpt = font.size;
                    model.val_axis_font_bold = font.bold;
                    model.val_axis_font_italic = font.italic;
                    model.val_axis_font_color = font.color;
                    model.val_axis_font_face = font.face;
                    model.val_axis_orientation =
                        (prop("reverse-direction") == Some("true")).then(|| "maxMin".into());
                    model.val_min = number("minimum");
                    model.val_max = number("maximum");
                    model.val_axis_major_unit = number("interval-major");
                    if let (Some(major), Some(divisor)) =
                        (number("interval-major"), number("interval-minor-divisor"))
                    {
                        if divisor > 0.0 {
                            model.val_axis_minor_unit = Some(major / divisor);
                        }
                    }
                    model.val_axis_log_base = (prop("logarithmic") == Some("true")).then_some(10.0);
                    model.val_axis_format_code = format_code;
                    model.bar_gap_width = number("gap-width").map(|v| v.round() as i32);
                    model.bar_overlap = number("overlap").map(|v| v.round() as i32);
                    if let Some((text, font)) = title {
                        model.val_axis_title = Some(text);
                        model.val_axis_title_font_size_hpt = font.size;
                        model.val_axis_title_font_bold = font.bold;
                        model.val_axis_title_font_color = font.color;
                    }
                    if let Some(grid) = grid("major") {
                        model.val_axis_major_gridlines = Some(true);
                        model.val_axis_gridline_color = grid.line;
                        model.val_axis_gridline_width_emu = grid.line_width;
                    } else {
                        model.val_axis_major_gridlines = Some(false);
                    }
                    if let Some(grid) = grid("minor") {
                        model.val_axis_minor_gridlines = Some(true);
                        model.val_axis_minor_gridline_color = grid.line;
                    }
                }
                ("y", true) => {
                    let value = serde_json::json!({
                        "hidden": !visible,
                        "majorTickMark": major_tick,
                        "minorGridlines": grid("minor").is_some(),
                        "majorGridlines": grid("major").is_some(),
                        "lineHidden": paint.line_hidden,
                        "min": number("minimum"),
                        "max": number("maximum"),
                        "majorUnit": number("interval-major"),
                        "title": title.as_ref().map(|(t, _)| t.clone()),
                        "fontColor": font.color,
                        "fontSizeHpt": font.size,
                        "fontBold": font.bold,
                        "lineColor": paint.line,
                        "formatCode": format_code,
                        "orientation": (prop("reverse-direction") == Some("true")).then_some("maxMin"),
                        "logBase": (prop("logarithmic") == Some("true")).then_some(10.0),
                        "tickLabelPos": (!labels).then_some("none"),
                    });
                    model.secondary_val_axis = serde_json::from_value(value).ok();
                }
                _ => {}
            }
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn series(
        &self,
        plot: roxmltree::Node<'_, '_>,
        plot_chain: &[&Style],
        class: &str,
        primary: &str,
        model: &mut ChartModel,
        data: &mut DataStyles,
    ) -> Result<(), String> {
        let categories: Vec<String> = plot
            .descendants()
            .find(|n| xml::is(*n, NS_CHART, "categories"))
            .and_then(|c| c.attribute((NS_TABLE, "cell-range-address")))
            .map(|address| {
                self.range(address)
                    .into_iter()
                    .map(|(_, text)| text)
                    .collect()
            })
            .unwrap_or_default();
        // Each series keeps its own copy of the categories.
        let categories_bytes = categories.iter().fold(0usize, |sum, c| {
            sum.saturating_add(std::mem::size_of::<String>() + c.len())
        });
        let nodes: Vec<_> = plot
            .children()
            .filter(|n| xml::is(*n, NS_CHART, "series"))
            .take(MAX_SERIES)
            .collect();
        let mut series_models = Vec::new();
        for (index, node) in nodes.iter().enumerate() {
            let node = *node;
            let series_class = node
                .attribute((NS_CHART, "class"))
                .map(|c| c.trim_start_matches("chart:"))
                .unwrap_or(class);
            let kind = match series_class {
                "bar" | "gantt" => "bar",
                "line" => "line",
                "area" => "area",
                "circle" => "pie",
                "ring" => "doughnut",
                "scatter" => "scatter",
                "bubble" => "bubble",
                "radar" | "filled-radar" => "radar",
                "stock" => "stock",
                _ => primary,
            };
            let values: Vec<Option<f64>> = node
                .attribute((NS_CHART, "values-cell-range-address"))
                .map(|address| {
                    self.range(address)
                        .into_iter()
                        .map(|(cell, _)| match cell {
                            CellData::Number(n) => Some(n),
                            _ => None,
                        })
                        .collect()
                })
                .unwrap_or_default();
            let name = node
                .attribute((NS_CHART, "label-cell-address"))
                .map(|address| {
                    self.range(address)
                        .into_iter()
                        .map(|(_, text)| text)
                        .collect::<Vec<_>>()
                        .join(" ")
                })
                .unwrap_or_else(|| format!("Series{}", index + 1));
            let domains: Vec<Vec<(CellData, String)>> = node
                .children()
                .filter(|n| xml::is(*n, NS_CHART, "domain"))
                .filter_map(|d| d.attribute((NS_TABLE, "cell-range-address")))
                .map(|address| self.range(address))
                .collect();
            let chain = self.chain(node);
            let paint = self.paint(&chain);
            let line_drawn = matches!(kind, "line" | "scatter")
                || kind == "radar" && series_class != "filled-radar";
            let palette = DEFAULT_COLORS[index % DEFAULT_COLORS.len()].to_string();
            let mut series = ChartSeries {
                name,
                series_type: Some(
                    match kind {
                        "doughnut" => "doughnut",
                        other => other,
                    }
                    .to_string(),
                ),
                use_secondary_axis: (node.attribute((NS_CHART, "attached-axis"))
                    == Some("secondary-y"))
                .then_some(true),
                ..ChartSeries::default()
            };
            series.color = if line_drawn {
                paint.line.clone().or(paint.fill.clone())
            } else {
                paint.fill.clone()
            }
            .or(Some(palette));
            series.line_color = paint.line.clone();
            series.line_width_emu = paint.line_width;
            if paint.line_hidden && line_drawn {
                series.line_hidden = Some(true);
            }
            // A classic series' own fill and line travel in its element
            // style, as the DrawingML parser carries the series `c:spPr`:
            // `draw:fill="none"` / `draw:stroke="none"` paint nothing.
            let filled_series = !line_drawn || kind == "scatter";
            if paint.dash.is_some() || paint.line_hidden || paint.fill_hidden && filled_series {
                series.chartex_style = Some(ooxml_common::chart::ChartExElementStyle {
                    shape_properties_present: Some(true),
                    line_dash: paint.dash.map(str::to_string),
                    line_dash_authored: paint.dash.map(|_| true),
                    fill_hidden: (paint.fill_hidden && filled_series).then_some(true),
                    fill_paint_authored: paint.fill.as_ref().map(|_| true),
                    fill_colors: paint.fill.clone().map(|c| vec![Some(c)]),
                    line_hidden: paint.line_hidden.then_some(true),
                    line_paint_authored: paint.line.as_ref().map(|_| true),
                    line_colors: paint.line.clone().map(|c| vec![Some(c)]),
                    line_width_emu: paint.line_width,
                    ..Default::default()
                });
            }
            if paint.fill_hidden && !line_drawn {
                series.color = paint.line.clone();
            }
            if paint.line_hidden {
                series.line_hidden = Some(true);
            }
            // Scatter and bubble: domains are x (and for bubble y, x) values.
            match kind {
                "scatter" => {
                    if let Some(x) = domains.first() {
                        series.categories = self.copy_texts(x);
                    }
                    series.values = values;
                }
                "bubble" => {
                    // §19.12: values are the bubble sizes; the first domain
                    // the y values, the second the x values.
                    series.bubble_sizes = Some(values);
                    series.values = domains
                        .first()
                        .map(|d| {
                            d.iter()
                                .map(|(c, _)| {
                                    if let CellData::Number(n) = c {
                                        Some(*n)
                                    } else {
                                        None
                                    }
                                })
                                .collect()
                        })
                        .unwrap_or_default();
                    if let Some(x) = domains.get(1) {
                        series.categories = self.copy_texts(x);
                    }
                }
                _ => {
                    series.values = values;
                    if !categories.is_empty() && self.budget.charge(categories_bytes) {
                        series.categories = Some(categories.clone());
                    }
                }
            }
            // Markers (§20.53 chart:symbol-type).
            match get(&chain, "chart:chart:symbol-type") {
                Some("none") => series.show_marker = Some(false),
                Some("named-symbol") => {
                    series.show_marker = Some(true);
                    series.marker_symbol = get(&chain, "chart:chart:symbol-name").map(|s| {
                        match s {
                            "square" => "square",
                            "diamond" => "diamond",
                            "arrow-down" | "arrow-up" | "arrow-left" | "arrow-right" => "triangle",
                            "circle" => "circle",
                            "star" => "star",
                            "x" => "x",
                            "plus" => "plus",
                            "asterisk" => "star",
                            "horizontal-bar" => "dash",
                            "vertical-bar" => "dash",
                            _ => "square",
                        }
                        .to_string()
                    });
                    series.marker_fill = paint.fill.clone();
                    series.marker_line = paint.line.clone();
                }
                _ => {}
            }
            if let Some(size) = get(&chain, "chart:chart:symbol-width").and_then(length_pt) {
                series.marker_size = Some(size);
            }
            if get(&chain, "chart:chart:interpolation").is_some_and(|i| i != "none") {
                series.smooth = Some(true);
            }
            if let Some(offset) =
                get(&chain, "chart:chart:pie-offset").and_then(|v| v.parse::<f64>().ok())
            {
                if offset > 0.0 {
                    series.explosion = Some(offset.round() as u32);
                }
            }
            if let Some(labels) = self.data_labels(&chain, data) {
                series.series_data_labels = Some(labels);
            }
            // Data points (§11.13): per-point styles and repeats.
            let mut point_colors: Vec<Option<String>> = Vec::new();
            let mut any_point = false;
            for point in node
                .children()
                .filter(|n| xml::is(*n, NS_CHART, "data-point"))
            {
                let repeat = point
                    .attribute((NS_CHART, "repeated"))
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(1)
                    .clamp(1, MAX_POINTS);
                let color = point.attribute((NS_CHART, "style-name")).and_then(|_| {
                    let chain = self.chain(point);
                    let p = self.paint(&chain);
                    if line_drawn {
                        p.line.or(p.fill)
                    } else {
                        p.fill
                    }
                });
                any_point |= color.is_some();
                let copy =
                    std::mem::size_of::<Option<String>>() + color.as_ref().map_or(0, String::len);
                for _ in 0..repeat {
                    if point_colors.len() >= MAX_POINTS || !self.budget.charge(copy) {
                        break;
                    }
                    point_colors.push(color.clone());
                }
            }
            if any_point {
                series.data_point_colors = Some(point_colors);
            }
            // Regression curves (§11.16) as trendlines.
            let trendlines: Vec<ooxml_common::chart::ChartTrendline> = node
                .children()
                .filter(|n| xml::is(*n, NS_CHART, "regression-curve"))
                .take(MAX_TRENDLINES)
                .take_while(|_| {
                    self.budget
                        .charge(std::mem::size_of::<ooxml_common::chart::ChartTrendline>())
                })
                .filter_map(|curve| {
                    let chain = self.chain(curve);
                    let kind = match get(&chain, "chart:chart:regression-type")? {
                        "linear" => "linear",
                        "logarithmic" => "log",
                        "exponential" => "exp",
                        "power" => "power",
                        "polynomial" => "poly",
                        "moving-average" => "movingAvg",
                        _ => return None,
                    };
                    let paint = self.paint(&chain);
                    let mut value = serde_json::json!({ "trendlineType": kind });
                    if let Some(order) = get(&chain, "chart:chart:regression-max-degree")
                        .or_else(|| get(&chain, "chart:loext:regression-max-degree"))
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        value["order"] = order.into();
                    }
                    if let Some(period) = get(&chain, "chart:chart:regression-period")
                        .or_else(|| get(&chain, "chart:loext:regression-period"))
                        .and_then(|v| v.parse::<u32>().ok())
                    {
                        value["period"] = period.into();
                    }
                    let mut trend: ooxml_common::chart::ChartTrendline =
                        serde_json::from_value(value).ok()?;
                    let _ = paint;
                    trend.name = None;
                    Some(trend)
                })
                .collect();
            if !trendlines.is_empty() {
                series.trend_lines = Some(trendlines);
            }
            // The series record and its metadata (labels, trendlines,
            // styles); its point data was charged as it was read.
            let values = std::mem::take(&mut series.values);
            let categories = series.categories.take();
            let bubbles = series.bubble_sizes.take();
            let colors = series.data_point_colors.take();
            let metadata = ooxml_common::json_measurement::measure_json(&series)
                .ok()
                .and_then(|m| usize::try_from(m.json_bytes).ok())
                .unwrap_or(usize::MAX);
            if !self
                .budget
                .charge(metadata.saturating_add(std::mem::size_of::<ChartSeries>()))
            {
                return Err("OpenDocument chart exceeds the data limit".into());
            }
            series.values = values;
            series.categories = categories;
            series.bubble_sizes = bubbles;
            series.data_point_colors = colors;
            series_models.push(series);
        }
        if series_models.is_empty() {
            model.authored_without_series = true;
        }
        // Varying colors per point (§20.10 chart:varying-color… LibreOffice
        // writes the pie slices as styled data points).
        model.categories = if categories.is_empty() {
            series_models
                .first()
                .and_then(|s| s.categories.clone())
                .unwrap_or_else(|| {
                    let count = series_models
                        .iter()
                        .map(|s| s.values.len())
                        .max()
                        .unwrap_or(0);
                    (1..=count).map(|i| i.to_string()).collect()
                })
        } else {
            categories
        };
        // Combinations are carried per series (`series_type`,
        // `use_secondary_axis`), which the renderer's classic path draws on
        // the shared category axis.
        if primary == "stock" {
            model.stock_hi_low_lines = Some(true);
            if get(plot_chain, "chart:chart:japanese-candle-stick") == Some("true") {
                model.stock_up_down_bars = Some(true);
            }
        }
        model.show_data_labels = series_models.iter().any(|s| s.series_data_labels.is_some());
        model.series = series_models;
        Ok(())
    }

    /// Series data labels (§20.32-§20.36 chart:data-label-*).
    fn data_labels(
        &self,
        chain: &[&Style],
        data: &mut DataStyles,
    ) -> Option<ChartSeriesDataLabels> {
        let number = get(chain, "chart:chart:data-label-number");
        let show_val = matches!(number, Some("value") | Some("value-and-percentage"));
        let show_percent = matches!(number, Some("percentage") | Some("value-and-percentage"));
        let show_cat_name = get(chain, "chart:chart:data-label-text") == Some("true");
        let show_ser_name = get(chain, "chart:chart:data-label-series") == Some("true");
        if !(show_val || show_percent || show_cat_name || show_ser_name) {
            return None;
        }
        let font = self.font(chain);
        let separator = get_element(chain, "chart:chart:label-separator")
            .and_then(|e| e.get("#text"))
            .cloned();
        let format_code = chain
            .iter()
            .find_map(|s| s.data_style.as_deref())
            .filter(|_| get(chain, "chart:chart:link-data-style-to-source") != Some("true"))
            .and_then(|name| data.code(name))
            .map(|(code, _)| crate::ods_number::literal_code(&code))
            .filter(|code| code != "General");
        Some(ChartSeriesDataLabels {
            show_val,
            show_percent,
            show_cat_name,
            show_ser_name,
            position: get(chain, "chart:chart:label-position").map(|p| {
                match p {
                    "outside" => "outEnd",
                    "inside" => "inEnd",
                    "center" => "ctr",
                    "top" => "t",
                    "bottom" => "b",
                    "left" => "l",
                    "right" => "r",
                    "near-origin" => "inBase",
                    _ => "bestFit",
                }
                .to_string()
            }),
            font_color: font.color,
            font_size_hpt: font.size,
            font_bold: font.bold,
            font_italic: font.italic,
            font_face: font.face,
            separator,
            format_code,
            ..ChartSeriesDataLabels::default()
        })
    }
}

fn merge_font(base: TextFont, over: Option<TextFont>) -> TextFont {
    match over {
        None => base,
        Some(o) => TextFont {
            size: o.size.or(base.size),
            bold: o.bold.or(base.bold),
            italic: o.italic.or(base.italic),
            color: o.color.or(base.color),
            face: o.face.or(base.face),
        },
    }
}

/// `draw:stroke-dash` (§16.41): dot runs and distance relative to the line
/// width, matched to the DrawingML preset with the same pattern (§20.1.10.48).
fn stroke_dashes(root: roxmltree::Node<'_, '_>) -> HashMap<String, &'static str> {
    let draw = crate::xml::NS_DRAW;
    let mut out = HashMap::new();
    for node in root
        .descendants()
        .filter(|n| xml::is(*n, draw, "stroke-dash"))
    {
        let Some(name) = node.attribute((draw, "name")) else {
            continue;
        };
        let length = |key: &str| -> f64 {
            node.attribute((draw, key))
                .and_then(|v| percent(v).or_else(|| length_pt(v).map(|pt| pt / 0.75)))
                .unwrap_or(1.0)
        };
        let count = |key: &str| -> u32 {
            node.attribute((draw, key))
                .and_then(|v| v.parse().ok())
                .unwrap_or(0)
        };
        let (d1, l1, d2, l2) = (
            count("dots1"),
            length("dots1-length"),
            count("dots2"),
            length("dots2-length"),
        );
        let gap = length("distance");
        let preset = match (d1, d2) {
            (_, 0) | (0, _) => {
                let dash = if d1 > 0 { l1 } else { l2 };
                if dash <= 1.5 {
                    "sysDot"
                } else if dash >= 6.0 {
                    "lgDash"
                } else if gap <= 1.5 {
                    "sysDash"
                } else {
                    "dash"
                }
            }
            (_, n) => {
                let long = l1.max(l2);
                match (long >= 6.0, n >= 2 || d1 >= 2) {
                    (true, true) => "lgDashDotDot",
                    (true, false) => "lgDashDot",
                    (false, true) => "sysDashDotDot",
                    (false, false) => {
                        if gap <= 1.5 {
                            "sysDashDot"
                        } else {
                            "dashDot"
                        }
                    }
                }
            }
        };
        out.insert(name.to_string(), preset);
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cached_ranges_are_bounded_before_they_are_collected() {
        let mut table = LocalTable::default();
        let address = "Sheet1.A1:Sheet1.A4294967295";
        table.sources.insert(address.to_string(), (0, 0));
        assert_eq!(
            table.cached(address, &ChartBudget::new()).unwrap().len(),
            MAX_POINTS
        );
    }

    const NS: &str = concat!(
        r##"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "##,
        r##"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "##,
        r##"xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" "##,
        r##"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "##,
        r##"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "##,
        r##"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "##,
        r##"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "##,
        r##"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0""##
    );

    fn chart(body: &str, styles: &str) -> String {
        format!(
            r##"<office:document-content {NS} office:version="1.3"><office:automatic-styles>{styles}</office:automatic-styles><office:body><office:chart>{body}</office:chart></office:body></office:document-content>"##
        )
    }

    const LOCAL: &str = concat!(
        r##"<table:table table:name="local-table"><table:table-header-rows><table:table-row>"##,
        r##"<table:table-cell><text:p/></table:table-cell>"##,
        r##"<table:table-cell office:value-type="string"><text:p>Sales</text:p></table:table-cell>"##,
        r##"<table:table-cell office:value-type="string"><text:p>Cost</text:p></table:table-cell>"##,
        r##"</table:table-row></table:table-header-rows><table:table-rows>"##,
        r##"<table:table-row><table:table-cell office:value-type="string"><text:p>Q1</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell></table:table-row>"##,
        r##"<table:table-row><table:table-cell office:value-type="string"><text:p>Q2</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="12"><text:p>12</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell></table:table-row>"##,
        r##"</table:table-rows></table:table>"##
    );

    #[test]
    fn internal_data_bar_line_combination() {
        let body = format!(
            r##"<chart:chart chart:class="chart:bar"><chart:title><text:p>Revenue</text:p></chart:title><chart:legend chart:legend-position="bottom"/><chart:plot-area chart:style-name="pa"><chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="local-table.A2:.A3"/></chart:axis><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ay"><chart:grid chart:class="major"/></chart:axis><chart:series chart:style-name="s1" chart:values-cell-range-address="local-table.B2:.B3" chart:label-cell-address="local-table.B1"/><chart:series chart:class="chart:line" chart:style-name="s2" chart:values-cell-range-address="local-table.C2:.C3" chart:label-cell-address="local-table.C1"/></chart:plot-area>{LOCAL}</chart:chart>"##
        );
        let styles = concat!(
            r##"<style:style style:name="pa" style:family="chart"><style:chart-properties chart:stacked="true"/></style:style>"##,
            r##"<style:style style:name="ay" style:family="chart"><style:chart-properties chart:minimum="0" chart:maximum="20" chart:gap-width="50"/></style:style>"##,
            r##"<style:style style:name="s1" style:family="chart"><style:chart-properties chart:data-label-number="value" chart:label-position="outside"/><style:graphic-properties draw:fill-color="#4472c4"/></style:style>"##,
            r##"<style:style style:name="s2" style:family="chart"><style:graphic-properties svg:stroke-color="#ed7d31"/></style:style>"##,
        );
        let model = read_chart_document(&chart(&body, styles), None, None)
            .unwrap()
            .unwrap();
        assert_eq!(model.chart_type, "stackedBar");
        assert_eq!(model.title.as_deref(), Some("Revenue"));
        assert_eq!(model.legend_pos.as_deref(), Some("b"));
        assert_eq!(model.categories, ["Q1", "Q2"]);
        assert_eq!(model.series.len(), 2);
        assert_eq!(model.series[0].name, "Sales");
        assert_eq!(model.series[0].values, [Some(10.0), Some(12.0)]);
        assert_eq!(model.series[0].color.as_deref(), Some("4472C4"));
        assert_eq!(model.series[1].series_type.as_deref(), Some("line"));
        assert_eq!(model.series[1].color.as_deref(), Some("ED7D31"));
        assert_eq!((model.val_min, model.val_max), (Some(0.0), Some(20.0)));
        assert_eq!(model.bar_gap_width, Some(50));
        assert_eq!(model.val_axis_major_gridlines, Some(true));
        let labels = model.series[0].series_data_labels.as_ref().unwrap();
        assert!(labels.show_val);
        assert_eq!(labels.position.as_deref(), Some("outEnd"));
        assert!(model.plot_groups.is_none());
    }

    struct Fixed;
    impl RangeResolver for Fixed {
        fn cells(&self, address: &str) -> Option<Vec<CellData>> {
            match address {
                "Sheet1.A2:Sheet1.A3" => {
                    Some(vec![CellData::Number(45292.0), CellData::Number(45323.0)])
                }
                "Sheet1.B2:Sheet1.B3" => Some(vec![CellData::Number(1.0), CellData::Number(2.0)]),
                "Sheet1.B1" => Some(vec![CellData::Text("Units".into())]),
                _ => None,
            }
        }
        fn hidden(&self, _: &str) -> Option<Vec<bool>> {
            None
        }
    }

    #[test]
    fn sheet_ranges_with_cached_display_text_and_pie_geometry() {
        let local = concat!(
            r##"<table:table table:name="local-table"><table:table-rows>"##,
            r##"<table:table-row><table:table-cell office:value-type="string"><text:p>Jan 24</text:p><draw:g><svg:desc>Sheet1.A2:Sheet1.A3</svg:desc></draw:g></table:table-cell></table:table-row>"##,
            r##"<table:table-row><table:table-cell office:value-type="string"><text:p>Feb 24</text:p></table:table-cell></table:table-row>"##,
            r##"</table:table-rows></table:table>"##
        );
        let body = format!(
            r##"<chart:chart chart:class="chart:circle"><chart:plot-area chart:style-name="pa"><chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="Sheet1.A2:Sheet1.A3"/></chart:axis><chart:series chart:values-cell-range-address="Sheet1.B2:Sheet1.B3" chart:label-cell-address="Sheet1.B1"><chart:data-point chart:style-name="p1"/><chart:data-point chart:style-name="p2"/></chart:series></chart:plot-area>{local}</chart:chart>"##
        );
        let styles = concat!(
            r##"<style:style style:name="pa" style:family="chart"><style:chart-properties chart:angle-offset="0"/></style:style>"##,
            r##"<style:style style:name="p1" style:family="chart"><style:graphic-properties draw:fill-color="#ff0000"/></style:style>"##,
            r##"<style:style style:name="p2" style:family="chart"><style:graphic-properties draw:fill-color="#00ff00"/></style:style>"##,
        );
        let model = read_chart_document(&chart(&body, styles), None, Some(&Fixed))
            .unwrap()
            .unwrap();
        assert_eq!(model.chart_type, "pie");
        assert_eq!(model.categories, ["Jan 24", "Feb 24"]);
        assert_eq!(model.series[0].name, "Units");
        assert_eq!(model.first_slice_angle, Some(90));
        assert_eq!(
            model.series[0].data_point_colors,
            Some(vec![Some("FF0000".into()), Some("00FF00".into())])
        );
    }
}
