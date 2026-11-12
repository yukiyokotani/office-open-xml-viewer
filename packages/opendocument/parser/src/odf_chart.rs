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

use crate::budget::{self, Scope};
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

/// Implementation bounds on one chart: series, points per range or series.
/// A chart over either rejects the document. Memory and work are bounded
/// by the governor ([`crate::budget`]) the chart is read under.
const MAX_SERIES: usize = 4096;
const MAX_POINTS: usize = 1_048_576;

/// The rejection of a chart over a bound.
fn chart_limit(what: &str) -> String {
    crate::error::resource(what)
}

/// Reject a range of `count` points when it is over the bound on the points
/// of one range; resolvers call it before they walk a range.
pub fn check_points(count: u64) -> Result<(), String> {
    if count > MAX_POINTS as u64 {
        return Err(chart_limit("chart point"));
    }
    Ok(())
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

/// One cell of a range outside the chart, with whether its row or column is
/// hidden (§20.26 `chart:include-hidden-cells`).
#[derive(Clone, Debug, PartialEq)]
pub struct SourceCell {
    pub data: CellData,
    pub hidden: bool,
}

/// Cell ranges outside the chart (a spreadsheet's tables).
pub trait RangeResolver {
    /// The cells of an ODF range address (`Sheet1.A1:Sheet1.A9`) in row-major
    /// order, or `None` when the address names no readable range. The
    /// resolver rejects a range over the point bound ([`check_points`])
    /// and walks its cells through [`crate::budget::each`].
    fn cells(&self, address: &str) -> Result<Option<Vec<SourceCell>>, String>;
    /// Whether the resolver's dates are serials of the 1904 date system
    /// (§19.636 `table:null-date`); the chart's own dates follow it.
    fn date1904(&self) -> bool {
        false
    }
    /// The number format of the first cell of a range, as a format code and
    /// its section divisors; `None` for General or an unreadable range.
    fn format(&self, _address: &str) -> Option<(String, Option<Vec<f64>>)> {
        None
    }
}

/// One point of a resolved range: its value, its display text, and whether
/// its source cell is hidden.
#[derive(Clone)]
struct Point {
    cell: CellData,
    text: String,
    hidden: bool,
}

impl Point {
    /// A cell of the chart's own table (never hidden).
    fn local((cell, text): (CellData, String)) -> Self {
        Self {
            cell,
            text,
            hidden: false,
        }
    }

    fn number(&self) -> Option<f64> {
        match self.cell {
            CellData::Number(n) => Some(n),
            _ => None,
        }
    }
}

/// A chart's tabular data (the deprecated whole-range form, §19.597.6 /
/// §20.54): its category labels and, per series, a label and its points.
struct Tabular {
    categories: Vec<Point>,
    series: Vec<(Option<String>, Vec<Point>)>,
}

/// The hidden flags of `points` when any is hidden (the chart model keeps
/// none otherwise).
fn hidden_flags(points: &[Point]) -> Option<Vec<bool>> {
    points
        .iter()
        .any(|p| p.hidden)
        .then(|| points.iter().map(|p| p.hidden).collect())
}

/// The package directory of a chart object reference (`./Object 1/` →
/// `Object 1`): the key under which a reader may share one object's read.
pub fn object_dir(reference: &str) -> &str {
    reference.trim_start_matches("./").trim_end_matches('/')
}

/// Read the chart object stored under `object_dir` (`Object 1`). `Ok(None)`
/// when the object is not a chart (another embedded document type). Inside
/// a reader's governed scope the chart is part of that read; on its own it
/// is governed by a scope of its own.
pub fn read_chart_object(
    package: &OdfPackage,
    object_dir: &str,
    resolver: Option<&dyn RangeResolver>,
) -> Result<Option<ChartModel>, String> {
    let dir = self::object_dir(object_dir);
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
    let scope = Scope::enter(budget::MAX_MODEL_BYTES);
    let chart = read_chart(
        content_doc.root_element(),
        styles_doc.as_ref().map(|doc| doc.root_element()),
        resolver,
    )?;
    scope.finish()?;
    Ok(chart)
}

fn read_chart(
    root: roxmltree::Node<'_, '_>,
    styles_root: Option<roxmltree::Node<'_, '_>>,
    resolver: Option<&dyn RangeResolver>,
) -> Result<Option<ChartModel>, String> {
    crate::odf_text::check_style_values(root)?;
    if let Some(styles_root) = styles_root {
        crate::odf_text::check_style_values(styles_root)?;
    }
    let mut store = StyleStore::default();
    let mut data = DataStyles::default();
    let mut faces = HashMap::new();
    if let Some(styles_root) = styles_root {
        odf_props::font_faces(styles_root, &mut faces)?;
        for child in styles_root.children().filter(|n| n.is_element()) {
            if xml::is(child, NS_OFFICE, "styles") {
                store.add(child, false)?;
                data.add(child)?;
            }
        }
    }
    odf_props::font_faces(root, &mut faces)?;
    for child in root
        .children()
        .filter(|n| xml::is(*n, NS_OFFICE, "automatic-styles"))
    {
        store.add(child, true)?;
        data.add(child)?;
    }
    store.check_chains()?;
    data.check()?;
    let Some(chart) = root
        .children()
        .find(|n| xml::is(*n, NS_OFFICE, "body"))
        .and_then(|b| b.children().find(|n| xml::is(*n, NS_OFFICE, "chart")))
        .and_then(|c| c.children().find(|n| xml::is(*n, NS_CHART, "chart")))
    else {
        return Ok(None);
    };
    let date1904 = resolver.is_some_and(|resolver| resolver.date1904());
    let local = LocalTable::read(chart, date1904)?;
    let reader = Reader {
        store: &store,
        faces: &faces,
        local,
        date1904,
        resolver,
        dashes: styles_root.map(stroke_dashes).unwrap_or_default(),
    };
    reader.chart(chart, &mut data).map(Some)
}

struct Reader<'a> {
    store: &'a StyleStore,
    faces: &'a HashMap<String, String>,
    local: LocalTable,
    date1904: bool,
    resolver: Option<&'a dyn RangeResolver>,
    /// `draw:stroke-dash` definitions (§16.41) as DrawingML presets.
    dashes: HashMap<String, &'static str>,
}

/// The chart's `local-table` (§11.1 data table): a grid with the header
/// row and column included, plus each cell's `svg:desc` source range. Row
/// and column repeats (§19.675 / §19.680) are kept as runs, never expanded:
/// the table costs what its XML holds, whatever its repeat counts.
#[derive(Default)]
struct LocalTable {
    rows: Vec<RowRun>,
    /// Source range address → (column, first data row).
    sources: HashMap<String, (usize, usize)>,
    /// Whether the table marks its first row / column as labels
    /// (`table:table-header-rows` / `table:table-header-columns`).
    header_row: bool,
    header_column: bool,
}

/// Grid rows up to `end` (exclusive) that repeat one authored row.
struct RowRun {
    end: usize,
    cells: Vec<CellRun>,
}

/// Grid columns up to `end` (exclusive) that repeat one authored cell.
struct CellRun {
    end: usize,
    value: CellData,
    text: String,
}

/// A table repeat count (§19.675 / §19.680, a positive integer; absent or
/// unreadable is 1; one beyond the implementation bound rejects).
fn repeat_count(node: roxmltree::Node<'_, '_>, name: &str) -> Result<usize, String> {
    xml::count_attribute(node, NS_TABLE, name).map(|n| n as usize)
}

impl LocalTable {
    /// Read the grid; covered cells keep their grid positions (empty).
    fn read(chart: roxmltree::Node<'_, '_>, date1904: bool) -> Result<Self, String> {
        let mut table = Self::default();
        let Some(node) = chart.children().find(|n| xml::is(*n, NS_TABLE, "table")) else {
            return Ok(table);
        };
        table.header_row = node
            .children()
            .any(|n| xml::is(n, NS_TABLE, "table-header-rows"));
        table.header_column = node
            .children()
            .any(|n| xml::is(n, NS_TABLE, "table-header-columns"));
        let mut next_row = 0usize;
        for row in node
            .descendants()
            .filter(|n| xml::is(*n, NS_TABLE, "table-row"))
        {
            budget::tick(1)?;
            let row_repeat = repeat_count(row, "number-rows-repeated")?;
            let mut cells: Vec<CellRun> = Vec::new();
            let mut next_col = 0usize;
            for cell in row.children().filter(|n| {
                xml::is(*n, NS_TABLE, "table-cell") || xml::is(*n, NS_TABLE, "covered-table-cell")
            }) {
                budget::tick(1)?;
                let repeat = repeat_count(cell, "number-columns-repeated")?;
                let covered = xml::is(cell, NS_TABLE, "covered-table-cell");
                // §19.383 `office:string-value` holds a string cell's value
                // when it differs from (or stands in for) its paragraphs.
                let text = if covered {
                    String::new()
                } else if let Some(value) = cell
                    .attribute((NS_OFFICE, "string-value"))
                    .filter(|_| cell.attribute((NS_OFFICE, "value-type")) == Some("string"))
                {
                    value.to_string()
                } else {
                    crate::odf_text::paragraph_text(cell)?
                };
                let value = match cell.attribute((NS_OFFICE, "value-type")) {
                    _ if covered => CellData::Empty,
                    Some("float") | Some("percentage") | Some("currency") => cell
                        .attribute((NS_OFFICE, "value"))
                        .and_then(|v| v.parse::<f64>().ok())
                        .map_or(CellData::Empty, CellData::Number),
                    // Dates, times and booleans are numbers to a chart, as
                    // they are in a worksheet (§19.385).
                    Some("date") => cell
                        .attribute((NS_OFFICE, "date-value"))
                        .and_then(|v| crate::odf_text::date_serial(v, date1904))
                        .map_or(CellData::Empty, CellData::Number),
                    Some("time") => cell
                        .attribute((NS_OFFICE, "time-value"))
                        .and_then(crate::odf_text::duration_days)
                        .map_or(CellData::Empty, CellData::Number),
                    Some("boolean") => match cell.attribute((NS_OFFICE, "boolean-value")) {
                        Some("true") | Some("1") => CellData::Number(1.0),
                        Some(_) => CellData::Number(0.0),
                        None => CellData::Empty,
                    },
                    Some("string") => CellData::Text(text.clone()),
                    _ if !text.is_empty() => CellData::Text(text.clone()),
                    _ => CellData::Empty,
                };
                let desc = cell
                    .descendants()
                    .find(|n| xml::is(*n, crate::xml::NS_SVG, "desc"))
                    .and_then(|n| n.text());
                if let Some(desc) = desc {
                    for range in crate::odf_text::split_addresses(desc) {
                        budget::tick(1)?;
                        table
                            .sources
                            .insert(range.to_string(), (next_col, next_row));
                    }
                }
                next_col = next_col.saturating_add(repeat);
                cells.push(CellRun {
                    end: next_col,
                    value,
                    text,
                });
            }
            // An empty row still occupies its grid rows.
            next_row = next_row.saturating_add(row_repeat);
            table.rows.push(RowRun {
                end: next_row,
                cells,
            });
        }
        Ok(table)
    }

    fn get(&self, row: usize, col: usize) -> Option<&CellRun> {
        let row = self
            .rows
            .get(self.rows.partition_point(|run| run.end <= row))?;
        row.cells
            .get(row.cells.partition_point(|run| run.end <= col))
    }

    /// A cell copied out of the table (empty outside it).
    fn copy(&self, row: usize, col: usize) -> (CellData, String) {
        match self.get(row, col) {
            Some(cell) => (cell.value.clone(), cell.text.clone()),
            None => (CellData::Empty, String::new()),
        }
    }

    /// The grid's size (rows, columns).
    fn size(&self) -> (usize, usize) {
        (
            self.rows.last().map_or(0, |run| run.end),
            self.rows
                .iter()
                .filter_map(|run| run.cells.last())
                .map(|cell| cell.end)
                .max()
                .unwrap_or(0),
        )
    }

    /// The local-table cells of a range on `local-table` (A1 = top-left).
    fn range(&self, address: &str) -> Result<Option<Vec<(CellData, String)>>, String> {
        let Some((sheet, (r1, c1), (r2, c2))) = crate::odf_text::parse_range(address) else {
            return Ok(None);
        };
        if sheet.as_deref().is_some_and(|s| s != "local-table") {
            return Ok(None);
        }
        check_points((u64::from(r2 - r1) + 1).saturating_mul(u64::from(c2 - c1) + 1))?;
        let mut out = Vec::new();
        for row in budget::each(r1, r2) {
            let row = row?;
            for col in budget::each(c1, c2) {
                out.push(self.copy(row as usize - 1, col? as usize - 1));
            }
        }
        Ok(Some(out))
    }

    /// Cached cells of an external source range, via the `svg:desc`
    /// annotations LibreOffice writes on the first cell of each column.
    fn cached(&self, address: &str) -> Result<Option<Vec<(CellData, String)>>, String> {
        let Some(&(col, row)) = self.sources.get(address) else {
            return Ok(None);
        };
        let Some((_, (r1, c1), (r2, c2))) = crate::odf_text::parse_range(address) else {
            return Ok(None);
        };
        let count = (r2 - r1).max(c2 - c1);
        check_points(u64::from(count) + 1)?;
        let horizontal = r1 == r2 && c2 > c1;
        let mut out = Vec::new();
        for i in budget::each(0, count) {
            let i = i? as usize;
            let (r, c) = if horizontal {
                (row, col.saturating_add(i))
            } else {
                (row.saturating_add(i), col)
            };
            out.push(self.copy(r, c));
        }
        Ok(Some(out))
    }
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

    fn font(&self, chain: &[&Style]) -> TextFont {
        TextFont {
            size: get(chain, "text:fo:font-size").and_then(length_pt).map(hpt),
            bold: get(chain, "text:fo:font-weight")
                .map(|w| w == "bold" || w.parse::<u32>().is_ok_and(|v| v >= 600)),
            italic: get(chain, "text:fo:font-style").map(|s| s == "italic" || s == "oblique"),
            color: get(chain, "text:fo:color").and_then(hex_color),
            face: get(chain, "text:style:font-name")
                .map(|name| {
                    self.faces
                        .get(name)
                        .cloned()
                        .unwrap_or_else(|| name.to_string())
                })
                .or_else(|| get(chain, "text:fo:font-family").map(odf_props::unquote_family)),
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
    fn text(node: roxmltree::Node<'_, '_>) -> Result<String, String> {
        crate::odf_text::paragraph_text(node)
    }

    /// Resolve a range: the resolver, else the local table. A range over
    /// [`MAX_POINTS`] rejects the document.
    fn range(&self, address: &str) -> Result<Vec<Point>, String> {
        // Several space-separated ranges are concatenated (§18.3.5).
        let mut out: Vec<Point> = Vec::new();
        let local = |cells: Option<Vec<(CellData, String)>>| {
            cells.map(|cells| cells.into_iter().map(Point::local).collect::<Vec<_>>())
        };
        for part in crate::odf_text::split_addresses(address) {
            budget::tick(1)?;
            let local_first =
                part.starts_with("local-table.") || part.starts_with("'local-table'.");
            let resolved = if local_first {
                local(self.local.range(part)?)
            } else {
                let cells = match self.resolver {
                    Some(resolver) => resolver.cells(part)?,
                    None => None,
                };
                match cells {
                    Some(cells) => {
                        let cached = self.local.cached(part)?;
                        let mut resolved = Vec::new();
                        for (i, source) in cells.into_iter().enumerate() {
                            budget::tick(1)?;
                            // Display text from the cache when the source
                            // value is a number (formatted categories).
                            let text = match &source.data {
                                CellData::Text(t) => t.clone(),
                                CellData::Number(n) => cached
                                    .as_ref()
                                    .and_then(|c| c.get(i))
                                    .map(|(_, t)| t.clone())
                                    .filter(|t| !t.is_empty())
                                    .unwrap_or_else(|| number_text(*n)),
                                CellData::Empty => String::new(),
                            };
                            resolved.push(Point {
                                cell: source.data,
                                text,
                                hidden: source.hidden,
                            });
                        }
                        Some(resolved)
                    }
                    None => match self.local.cached(part)? {
                        Some(cached) => local(Some(cached)),
                        None => local(self.local.range(part)?),
                    },
                }
            };
            out.extend(resolved.unwrap_or_default());
            if out.len() > MAX_POINTS {
                return Err(chart_limit("chart point"));
            }
        }
        Ok(out)
    }

    /// The chart's tabular data, for series that name no range of their own:
    /// the plot area's `table:cell-range-address` (§19.597.6, deprecated)
    /// with its `chart:data-source-has-labels` (§19.17), else the chart's
    /// own table with its header row and column. Series are its columns,
    /// or its rows for `chart:series-source="rows"` (§20.54).
    fn tabular(
        &self,
        plot: roxmltree::Node<'_, '_>,
        plot_chain: &[&Style],
    ) -> Result<Option<Tabular>, String> {
        let (grid, rows, columns, label_row, label_column) =
            match plot.attribute((NS_TABLE, "cell-range-address")) {
                Some(address) => {
                    // One rectangle: a list of ranges has no single grid.
                    let parts = crate::odf_text::split_addresses(address);
                    let [part] = parts.as_slice() else {
                        return Ok(None);
                    };
                    let Some((_, (r1, c1), (r2, c2))) = crate::odf_text::parse_range(part) else {
                        return Ok(None);
                    };
                    let (rows, columns) = ((r2 - r1) as usize + 1, (c2 - c1) as usize + 1);
                    let grid = self.range(part)?;
                    if grid.len() != rows.saturating_mul(columns) {
                        return Ok(None);
                    }
                    let labels = plot
                        .attribute((NS_CHART, "data-source-has-labels"))
                        .unwrap_or("none");
                    (
                        grid,
                        rows,
                        columns,
                        matches!(labels, "row" | "both"),
                        matches!(labels, "column" | "both"),
                    )
                }
                None => {
                    let (rows, columns) = self.local.size();
                    check_points((rows as u64).saturating_mul(columns as u64))?;
                    let mut grid = Vec::new();
                    for row in budget::each(0, rows.saturating_sub(1) as u32) {
                        let row = row? as usize;
                        for column in budget::each(0, columns.saturating_sub(1) as u32) {
                            grid.push(Point::local(self.local.copy(row, column? as usize)));
                        }
                    }
                    if rows == 0 || columns == 0 {
                        return Ok(None);
                    }
                    (
                        grid,
                        rows,
                        columns,
                        self.local.header_row,
                        self.local.header_column,
                    )
                }
            };
        // Read as series in columns: for series in rows the grid is
        // transposed, and so are its label row and column.
        let by_rows = get(plot_chain, "chart:chart:series-source") == Some("rows");
        let (points, count, label_points, label_series) = if by_rows {
            (columns, rows, label_row, label_column)
        } else {
            (rows, columns, label_column, label_row)
        };
        let at = |point: usize, series: usize| -> &Point {
            if by_rows {
                &grid[series * columns + point]
            } else {
                &grid[point * columns + series]
            }
        };
        let first_point = usize::from(label_series);
        let first_series = usize::from(label_points);
        let mut tabular = Tabular {
            categories: Vec::new(),
            series: Vec::new(),
        };
        if label_points {
            for point in first_point..points {
                budget::tick(1)?;
                tabular.categories.push(at(point, 0).clone());
            }
        }
        for series in first_series..count {
            let mut values = Vec::new();
            for point in first_point..points {
                budget::tick(1)?;
                values.push(at(point, series).clone());
            }
            tabular
                .series
                .push((label_series.then(|| at(0, series).text.clone()), values));
        }
        Ok(Some(tabular))
    }

    /// The number format of the cells a range names (its first cell): what
    /// `chart:link-data-style-to-source` links an axis or a label to
    /// (§20.35).
    fn source_format(&self, address: Option<&str>) -> Option<(String, Option<Vec<f64>>)> {
        let first = crate::odf_text::split_addresses(address?)
            .into_iter()
            .next()?;
        self.resolver?.format(first)
    }

    /// The local name of a namespace-qualified chart class token
    /// (§19.15.1 `chart:class`, §19.12 series class): its prefix is
    /// resolved through the element's namespace bindings, and only a token
    /// in the chart namespace names a chart class.
    fn class_token<'n>(node: roxmltree::Node<'n, '_>, value: &'n str) -> Option<&'n str> {
        let (prefix, local) = value.trim().split_once(':')?;
        (node.lookup_namespace_uri(Some(prefix)) == Some(NS_CHART)).then_some(local)
    }

    fn chart(
        &self,
        chart: roxmltree::Node<'_, '_>,
        data: &mut DataStyles,
    ) -> Result<ChartModel, String> {
        let class = chart
            .attribute((NS_CHART, "class"))
            .and_then(|value| Self::class_token(chart, value))
            .unwrap_or("bar");
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
            date1904: self.date1904,
            cat_axis_cross_between: if matches!(primary, "bar") {
                "between".into()
            } else {
                "midCat".into()
            },
            val_axis_major_tick_mark: "out".into(),
            cat_axis_major_tick_mark: "out".into(),
            // §20.26 `chart:include-hidden-cells`: hidden cells are plotted
            // unless the plot area says otherwise.
            plot_visible_only: Some(
                get(&plot_chain, "chart:chart:include-hidden-cells") == Some("false"),
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
            let text = Self::text(title)?;
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
        let secondary_y = self.axes(plot, primary, &mut model, data)?;
        self.series(
            plot,
            &plot_chain,
            (class, primary),
            secondary_y,
            &mut model,
            data,
        )?;
        Ok(model)
    }

    /// The range whose cells `axis` plots. An x axis plots its categories,
    /// or the x values of a scatter or bubble chart; a y axis the values of
    /// the first series drawn against it, or the y values of a bubble chart.
    /// §11.13: the one domain of a scatter series is its x values; a bubble
    /// series has its y values then its x values, and its values range is
    /// the bubble sizes.
    fn axis_source<'n>(
        plot: roxmltree::Node<'n, '_>,
        axis: roxmltree::Node<'n, '_>,
        primary: &str,
    ) -> Result<Option<&'n str>, String> {
        let x = axis.attribute((NS_CHART, "dimension")) == Some("x");
        let name = axis.attribute((NS_CHART, "name")).unwrap_or("");
        if x {
            if let Some(categories) = axis
                .children()
                .find(|n| xml::is(*n, NS_CHART, "categories"))
                .and_then(|c| c.attribute((NS_TABLE, "cell-range-address")))
            {
                return Ok(Some(categories));
            }
        }
        // Which domain of a series the axis plots; `None` for its values.
        let domain = match (primary, x) {
            ("bubble", true) => Some(1),
            ("bubble", false) | (_, true) => Some(0),
            (_, false) => None,
        };
        for series in plot.children() {
            budget::tick(1)?;
            if !xml::is(series, NS_CHART, "series") {
                continue;
            }
            let attached = x
                || series
                    .attribute((NS_CHART, "attached-axis"))
                    .is_none_or(|attached| attached == name);
            if !attached {
                continue;
            }
            let Some(index) = domain else {
                return Ok(series.attribute((NS_CHART, "values-cell-range-address")));
            };
            let mut domains = series
                .children()
                .filter(|n| xml::is(*n, NS_CHART, "domain"))
                .filter_map(|d| d.attribute((NS_TABLE, "cell-range-address")));
            // Later series may leave their domains to the first that has
            // them, so a series without this one does not decide.
            if let Some(address) = domains.nth(index) {
                return Ok(Some(address));
            }
        }
        Ok(None)
    }

    /// Read the axes. The first axis of a dimension is its primary axis and
    /// the next y axis the secondary one, whatever they are named (§19.24
    /// `chart:name` is a free name; LibreOffice's `primary-y` /
    /// `secondary-y` are one choice). Returns the secondary y axis' name,
    /// which series attach to (§19.13).
    fn axes<'n>(
        &self,
        plot: roxmltree::Node<'n, '_>,
        primary: &str,
        model: &mut ChartModel,
        data: &mut DataStyles,
    ) -> Result<Option<&'n str>, String> {
        let mut seen: Vec<&str> = Vec::new();
        let mut secondary_y = None;
        for axis in plot.children().filter(|n| xml::is(*n, NS_CHART, "axis")) {
            budget::tick(1)?;
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
            let title = match axis.children().find(|n| xml::is(*n, NS_CHART, "title")) {
                Some(t) => Some((Self::text(t)?, self.font(&self.chain(t)))),
                None => None,
            };
            let grid = |class: &str| {
                axis.children()
                    .find(|n| {
                        xml::is(*n, NS_CHART, "grid")
                            && n.attribute((NS_CHART, "class")).unwrap_or("major") == class
                    })
                    .map(|g| self.paint(&self.chain(g)))
            };
            // §20.35: an axis linked to its source shows the number format
            // of the cells it plots.
            let format_code = if prop("link-data-style-to-source") == Some("true") {
                let source = Self::axis_source(plot, axis, primary)?;
                self.source_format(source)
            } else {
                chain
                    .iter()
                    .find_map(|s| s.data_style.as_deref())
                    .and_then(|name| data.code(name))
                    .filter(|(format, _)| !format.is_general())
                    .map(|(format, _)| format.chart_format())
            };
            let (format_code, format_divisors) = match format_code {
                Some((code, divisors)) => (Some(code), divisors),
                None => (None, None),
            };
            // 0 for the dimension's first axis, 1 for its second.
            let order = seen.iter().filter(|d| **d == dimension).count();
            if seen.len() < 16 {
                seen.push(dimension);
            }
            if dimension == "y" && order == 1 {
                secondary_y = Some(name);
            }
            match (dimension, order == 1) {
                // Further axes of a dimension have no place in the model.
                _ if order > 1 => {}
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
                    model.cat_axis_format_divisors = format_divisors;
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
                    model.val_axis_format_divisors = format_divisors;
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
                        "formatDivisors": format_divisors,
                        "orientation": (prop("reverse-direction") == Some("true")).then_some("maxMin"),
                        "logBase": (prop("logarithmic") == Some("true")).then_some(10.0),
                        "tickLabelPos": (!labels).then_some("none"),
                    });
                    model.secondary_val_axis = serde_json::from_value(value).ok();
                }
                _ => {}
            }
        }
        Ok(secondary_y)
    }

    fn series(
        &self,
        plot: roxmltree::Node<'_, '_>,
        plot_chain: &[&Style],
        (class, primary): (&str, &str),
        secondary_y: Option<&str>,
        model: &mut ChartModel,
        data: &mut DataStyles,
    ) -> Result<(), String> {
        let nodes: Vec<_> = plot
            .children()
            .filter(|n| xml::is(*n, NS_CHART, "series"))
            .collect();
        if nodes.len() > MAX_SERIES {
            return Err(chart_limit("chart series"));
        }
        // Series that name no values range take the chart's tabular data in
        // order (§19.597.6: the range of a series has priority).
        let mut tabular = if nodes.iter().all(|n| {
            n.attribute((NS_CHART, "values-cell-range-address"))
                .is_some()
        }) && !nodes.is_empty()
        {
            None
        } else {
            self.tabular(plot, plot_chain)?
        };
        if tabular
            .as_ref()
            .is_some_and(|t| t.series.len() > MAX_SERIES)
        {
            return Err(chart_limit("chart series"));
        }
        let category_points = match plot
            .descendants()
            .find(|n| xml::is(*n, NS_CHART, "categories"))
            .and_then(|c| c.attribute((NS_TABLE, "cell-range-address")))
        {
            Some(address) => self.range(address)?,
            None => tabular
                .as_mut()
                .map(|t| std::mem::take(&mut t.categories))
                .unwrap_or_default(),
        };
        model.category_source_hidden = hidden_flags(&category_points);
        let categories: Vec<String> = category_points.into_iter().map(|p| p.text).collect();
        let mut series_models = Vec::new();
        // §11.13: a series without the domains its class needs uses the
        // ones the previous series of the same class defined. Kept per
        // class as (y values, x values); shared, never copied.
        let mut domains_of: HashMap<&str, (Option<Domain>, Option<Domain>)> = HashMap::new();
        for (index, node) in nodes.iter().enumerate() {
            budget::tick(1)?;
            let node = *node;
            let series_class = node
                .attribute((NS_CHART, "class"))
                .and_then(|value| Self::class_token(node, value))
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
            let own_values = node.attribute((NS_CHART, "values-cell-range-address"));
            // Only a series without a range of its own is a column of the
            // tabular data (and takes that column's label).
            let from_table = tabular
                .as_mut()
                .filter(|_| own_values.is_none())
                .and_then(|t| t.series.get_mut(index))
                .map(std::mem::take);
            let (table_label, table_points) = from_table.unwrap_or_default();
            let value_points = match own_values {
                Some(address) => self.range(address)?,
                None => table_points,
            };
            let values: Vec<Option<f64>> = value_points.iter().map(Point::number).collect();
            let name = match node.attribute((NS_CHART, "label-cell-address")) {
                Some(address) => self
                    .range(address)?
                    .into_iter()
                    .map(|p| p.text)
                    .collect::<Vec<_>>()
                    .join(" "),
                None => table_label.unwrap_or_else(|| format!("Series{}", index + 1)),
            };
            let mut own_domains = Vec::new();
            for address in node
                .children()
                .filter(|n| xml::is(*n, NS_CHART, "domain"))
                .filter_map(|d| d.attribute((NS_TABLE, "cell-range-address")))
            {
                budget::tick(1)?;
                own_domains.push(Domain::new(self.range(address)?));
            }
            // §11.13: a scatter series has one domain, its x values; a
            // bubble series two, its y values then its x values.
            let (y_domain, x_domain) = match kind {
                "scatter" | "bubble" => {
                    let mut own = own_domains.into_iter();
                    let known = domains_of.entry(kind).or_default();
                    if kind == "bubble" {
                        if let Some(y) = own.next() {
                            known.0 = Some(y);
                        }
                    }
                    if let Some(x) = own.next() {
                        known.1 = Some(x);
                    }
                    known.clone()
                }
                _ => (None, None),
            };
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
                // §19.13: the series is drawn against the axis it names.
                use_secondary_axis: node
                    .attribute((NS_CHART, "attached-axis"))
                    .is_some_and(|axis| Some(axis) == secondary_y)
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
            // Scatter and bubble coordinates are the domains' numbers. The
            // model carries x values as text the renderer parses, so each is
            // written from the cell's value, never from its formatted text
            // (`1,000`, `50%`).
            // A point is hidden when any cell it is drawn from is (§20.26).
            let mut hidden: Vec<bool> = value_points.iter().map(|p| p.hidden).collect();
            let mut hide_with = |domain: &Domain| {
                for (at, point) in domain.0.iter().enumerate() {
                    match hidden.get_mut(at) {
                        Some(flag) => *flag |= point.hidden,
                        None => hidden.push(point.hidden),
                    }
                }
            };
            match kind {
                "scatter" => {
                    if let Some(x) = &x_domain {
                        series.categories = Some(x.coordinates()?);
                        hide_with(x);
                    }
                    series.values = values;
                }
                "bubble" => {
                    // §19.12: values are the bubble sizes.
                    series.bubble_sizes = Some(values);
                    if let Some(y) = &y_domain {
                        series.values = y.numbers()?;
                        hide_with(y);
                    }
                    if let Some(x) = &x_domain {
                        series.categories = Some(x.coordinates()?);
                        hide_with(x);
                    }
                }
                _ => {
                    series.values = values;
                    if !categories.is_empty() {
                        // Each series holds its own copy of the categories.
                        let mut copy = Vec::new();
                        for category in &categories {
                            budget::tick(1)?;
                            copy.push(category.clone());
                        }
                        series.categories = Some(copy);
                    }
                }
            }
            if hidden.iter().any(|flag| *flag) {
                series.source_hidden = Some(hidden);
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
            budget::tick(1)?;
            if let Some(labels) = self.data_labels(&chain, data, own_values) {
                series.series_data_labels = Some(labels);
            }
            // Data points (§11.13): per-point styles and repeats.
            let mut point_colors: Vec<Option<String>> = Vec::new();
            let mut any_point = false;
            for point in node
                .children()
                .filter(|n| xml::is(*n, NS_CHART, "data-point"))
            {
                // §19.21 chart:repeated: the point style covers that many
                // points; more points than a series holds reject.
                budget::tick(1)?;
                let repeat = xml::count_attribute(point, NS_CHART, "repeated")?;
                if point_colors.len().saturating_add(repeat as usize) > MAX_POINTS {
                    return Err(chart_limit("chart point"));
                }
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
                for item in budget::each(1, repeat) {
                    item?;
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
                        .and_then(|v| v.trim().parse::<u32>().ok())
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
            series_models.push(series);
        }
        // A chart without `chart:series` elements plots every series of its
        // tabular data, in the chart's class and the default colours.
        if nodes.is_empty() {
            for (index, (label, points)) in tabular
                .map(|t| t.series)
                .unwrap_or_default()
                .into_iter()
                .enumerate()
            {
                budget::tick(1)?;
                let mut copy = Vec::new();
                for category in &categories {
                    budget::tick(1)?;
                    copy.push(category.clone());
                }
                series_models.push(ChartSeries {
                    name: label.unwrap_or_else(|| format!("Series{}", index + 1)),
                    series_type: Some(primary.to_string()),
                    color: Some(DEFAULT_COLORS[index % DEFAULT_COLORS.len()].to_string()),
                    values: points.iter().map(Point::number).collect(),
                    source_hidden: hidden_flags(&points),
                    categories: (!copy.is_empty()).then_some(copy),
                    ..ChartSeries::default()
                });
            }
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
                    // At most `MAX_POINTS` values per series.
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
        values_address: Option<&str>,
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
        let format_code = if get(chain, "chart:chart:link-data-style-to-source") == Some("true") {
            self.source_format(values_address)
        } else {
            chain
                .iter()
                .find_map(|s| s.data_style.as_deref())
                .and_then(|name| data.code(name))
                .filter(|(format, _)| !format.is_general())
                .map(|(format, _)| format.chart_format())
        };
        let (format_code, format_divisors) = match format_code {
            Some((code, divisors)) => (Some(code), divisors),
            None => (None, None),
        };
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
            format_divisors,
            ..ChartSeriesDataLabels::default()
        })
    }
}

/// The points of one `chart:domain` range, shared by the series that use it.
#[derive(Clone)]
struct Domain(std::rc::Rc<Vec<Point>>);

impl Domain {
    fn new(points: Vec<Point>) -> Self {
        Self(std::rc::Rc::new(points))
    }

    /// The domain as values (a cell without a number is no value).
    fn numbers(&self) -> Result<Vec<Option<f64>>, String> {
        let mut out = Vec::new();
        for point in self.0.iter() {
            budget::tick(1)?;
            out.push(point.number());
        }
        Ok(out)
    }

    /// The domain as the coordinate texts the chart model carries: a number
    /// in the shortest form that parses back to it, a text cell as written
    /// (the renderer plots text-backed x values by position).
    fn coordinates(&self) -> Result<Vec<String>, String> {
        let mut out = Vec::new();
        for point in self.0.iter() {
            budget::tick(1)?;
            out.push(match &point.cell {
                CellData::Number(n) => number_text(*n),
                CellData::Text(text) => text.clone(),
                CellData::Empty => String::new(),
            });
        }
        Ok(out)
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
    // One pass over the styles document the scope already walked.
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

    fn line_series(values: &str, local: &str) -> String {
        format!(
            r##"<chart:chart chart:class="chart:line"><chart:plot-area><chart:series chart:values-cell-range-address="{values}"/></chart:plot-area>{local}</chart:chart>"##
        )
    }

    #[test]
    fn local_tables_keep_row_and_column_repeats_as_runs() {
        let local = concat!(
            r##"<table:table table:name="local-table"><table:table-rows>"##,
            r##"<table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="2" office:value-type="float" office:value="1"/></table:table-row>"##,
            r##"<table:table-row><table:covered-table-cell/><table:covered-table-cell/><table:table-cell office:value-type="float" office:value="2"/></table:table-row>"##,
            r##"</table:table-rows></table:table>"##
        );
        let model = read_chart_document(
            &chart(&line_series("local-table.B1:.B3", local), ""),
            None,
            None,
        )
        .unwrap()
        .unwrap();
        // The repeated row is two rows; a covered cell keeps its position.
        assert_eq!(model.series[0].values, [Some(1.0), Some(1.0), None]);
        let model = read_chart_document(
            &chart(&line_series("local-table.C3", local), ""),
            None,
            None,
        )
        .unwrap()
        .unwrap();
        assert_eq!(model.series[0].values, [Some(2.0)]);
        // Repeats are runs: a table of 2^64 cells costs its two elements.
        let wide = r##"<table:table table:name="local-table"><table:table-rows><table:table-row table:number-rows-repeated="4294967295"><table:table-cell table:number-columns-repeated="4294967295" office:value-type="float" office:value="7"/></table:table-row></table:table-rows></table:table>"##;
        let model = read_chart_document(
            &chart(&line_series("local-table.XFD1048575:.XFD1048576", wide), ""),
            None,
            None,
        )
        .unwrap()
        .unwrap();
        assert_eq!(model.series[0].values, [Some(7.0), Some(7.0)]);
    }

    #[test]
    fn cached_and_repeated_data_are_bounded() {
        // A cached source range far past the point bound.
        let cached = r##"<table:table table:name="local-table"><table:table-rows><table:table-row><table:table-cell office:value-type="float" office:value="1"><draw:g><svg:desc>Missing.A1:Missing.A1000000000</svg:desc></draw:g></table:table-cell></table:table-row></table:table-rows></table:table>"##;
        assert_eq!(
            read_chart_document(
                &chart(&line_series("Missing.A1:Missing.A1000000000", cached), ""),
                None,
                None
            )
            .err(),
            Some(chart_limit("chart point"))
        );
        // A 64 KiB string repeated over 1024 columns: reading the range
        // copies 64 MiB, past the limit of the read it is part of.
        let text = "x".repeat(64 * 1024);
        let strings = format!(
            r##"<table:table table:name="local-table"><table:table-rows><table:table-row><table:table-cell table:number-columns-repeated="1024" office:value-type="string"><text:p>{text}</text:p></table:table-cell></table:table-row></table:table-rows></table:table>"##
        );
        {
            let _scope = Scope::enter(16 << 20);
            assert_eq!(
                read_chart_document(
                    &chart(&line_series("local-table.A1:.AMJ1", &strings), ""),
                    None,
                    None
                )
                .err(),
                Some(chart_limit("model size"))
            );
        }
        // Too many points in a series' styles.
        let points = format!(
            r##"<chart:chart chart:class="chart:bar"><chart:plot-area><chart:series><chart:data-point chart:repeated="{}"/></chart:series></chart:plot-area></chart:chart>"##,
            MAX_POINTS + 1
        );
        assert_eq!(
            read_chart_document(&chart(&points, ""), None, None).err(),
            Some(chart_limit("chart point"))
        );
    }

    #[test]
    fn chart_classes_resolve_their_namespace_prefix() {
        let body = format!(
            r##"<chart:chart xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:x="urn:other" chart:class="c:line"><chart:plot-area><chart:series chart:class="x:area" chart:values-cell-range-address="local-table.B2:.B3"/><chart:series chart:class="c:area" chart:values-cell-range-address="local-table.C2:.C3"/></chart:plot-area>{LOCAL}</chart:chart>"##
        );
        let model = read_chart_document(&chart(&body, ""), None, None)
            .unwrap()
            .unwrap();
        assert!(model.chart_type.starts_with("line"), "{}", model.chart_type);
        // A token outside the chart namespace is no class: the chart's own.
        assert_eq!(model.series[0].series_type.as_deref(), Some("line"));
        assert_eq!(model.series[1].series_type.as_deref(), Some("area"));
    }

    struct Fixed;
    impl RangeResolver for Fixed {
        fn cells(&self, address: &str) -> Result<Option<Vec<SourceCell>>, String> {
            let visible = |cells: Vec<CellData>| {
                Some(
                    cells
                        .into_iter()
                        .map(|data| SourceCell {
                            data,
                            hidden: false,
                        })
                        .collect::<Vec<_>>(),
                )
            };
            // `H` holds two rows, the second hidden.
            if address == "Sheet1.H2:Sheet1.H3" {
                return Ok(Some(vec![
                    SourceCell {
                        data: CellData::Number(5.0),
                        hidden: false,
                    },
                    SourceCell {
                        data: CellData::Number(6.0),
                        hidden: true,
                    },
                ]));
            }
            Ok(match address {
                "Sheet1.A2:Sheet1.A3" => {
                    visible(vec![CellData::Number(45292.0), CellData::Number(45323.0)])
                }
                "Sheet1.B2:Sheet1.B3" => {
                    visible(vec![CellData::Number(1.0), CellData::Number(2.0)])
                }
                "Sheet1.B1" => visible(vec![CellData::Text("Units".into())]),
                "Sheet1.P2:Sheet1.P3" => {
                    visible(vec![CellData::Number(0.5), CellData::Number(0.25)])
                }
                "Sheet1.X2:Sheet1.X3" => {
                    visible(vec![CellData::Number(1000.0), CellData::Number(0.5)])
                }
                "Sheet1.Y2:Sheet1.Y3" => {
                    visible(vec![CellData::Number(3.0), CellData::Number(4.0)])
                }
                _ => None,
            })
        }
    }

    #[test]
    fn a_chart_without_series_ranges_plots_its_tabular_data() {
        // The deprecated source-range form: a plot-area range with labels
        // in its first row and column, and no `chart:series`.
        let body = |source: &str| {
            format!(
                r##"<chart:chart chart:class="chart:bar"><chart:plot-area table:cell-range-address="local-table.A1:.C3" chart:data-source-has-labels="both" {source}/>{LOCAL}</chart:chart>"##
            )
        };
        let model = read_chart_document(&chart(&body(""), ""), None, None)
            .unwrap()
            .unwrap();
        assert_eq!(model.categories, ["Q1", "Q2"]);
        let read = |model: &ChartModel| {
            model
                .series
                .iter()
                .map(|s| (s.name.clone(), s.values.clone()))
                .collect::<Vec<_>>()
        };
        assert_eq!(
            read(&model),
            [
                ("Sales".to_string(), vec![Some(10.0), Some(12.0)]),
                ("Cost".to_string(), vec![Some(4.0), Some(5.0)]),
            ]
        );
        assert!(!model.authored_without_series);
        // Series in rows (§20.54).
        let rows = r##"<style:style style:name="pa" style:family="chart"><style:chart-properties chart:series-source="rows"/></style:style>"##;
        let model =
            read_chart_document(&chart(&body(r#"chart:style-name="pa""#), rows), None, None)
                .unwrap()
                .unwrap();
        assert_eq!(model.categories, ["Sales", "Cost"]);
        assert_eq!(
            read(&model),
            [
                ("Q1".to_string(), vec![Some(10.0), Some(4.0)]),
                ("Q2".to_string(), vec![Some(12.0), Some(5.0)]),
            ]
        );
        // Series elements without ranges take the chart table's columns
        // (its header row and column are its labels).
        let local = LOCAL.replace(
            r##"<table:table table:name="local-table">"##,
            r##"<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns>"##,
        );
        let old = format!(
            r##"<chart:chart chart:class="chart:line"><chart:plot-area><chart:series/><chart:series chart:values-cell-range-address="local-table.B2:.B3"/></chart:plot-area>{local}</chart:chart>"##
        );
        let model = read_chart_document(&chart(&old, ""), None, None)
            .unwrap()
            .unwrap();
        assert_eq!(model.categories, ["Q1", "Q2"]);
        assert_eq!(
            read(&model),
            [
                ("Sales".to_string(), vec![Some(10.0), Some(12.0)]),
                ("Series2".to_string(), vec![Some(10.0), Some(12.0)]),
            ]
        );
    }

    #[test]
    fn linked_formats_come_from_the_source_cells() {
        struct Percent;
        impl RangeResolver for Percent {
            fn cells(&self, address: &str) -> Result<Option<Vec<SourceCell>>, String> {
                Fixed.cells(address)
            }
            fn format(&self, address: &str) -> Option<(String, Option<Vec<f64>>)> {
                (address == "Sheet1.P2:Sheet1.P3").then(|| ("0%".to_string(), None))
            }
        }
        let styles = concat!(
            r##"<number:number-style xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" style:name="N2"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style>"##,
            r##"<style:style style:name="linked" style:family="chart" style:data-style-name="N2"><style:chart-properties chart:link-data-style-to-source="true" chart:data-label-number="value"/></style:style>"##,
            r##"<style:style style:name="own" style:family="chart" style:data-style-name="N2"><style:chart-properties chart:link-data-style-to-source="false"/></style:style>"##,
        );
        let body = |axis: &str| {
            format!(
                r##"<chart:chart chart:class="chart:bar"><chart:plot-area><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="{axis}"/><chart:series chart:style-name="linked" chart:values-cell-range-address="Sheet1.P2:Sheet1.P3"/></chart:plot-area></chart:chart>"##
            )
        };
        let model = read_chart_document(&chart(&body("linked"), styles), None, Some(&Percent))
            .unwrap()
            .unwrap();
        // §20.35: the axis and the labels show the cells' percentage format,
        // not the style's own.
        assert_eq!(model.val_axis_format_code.as_deref(), Some("0%"));
        let labels = model.series[0].series_data_labels.as_ref().unwrap();
        assert_eq!(labels.format_code.as_deref(), Some("0%"));
        let model = read_chart_document(&chart(&body("own"), styles), None, Some(&Percent))
            .unwrap()
            .unwrap();
        assert_eq!(model.val_axis_format_code.as_deref(), Some("0.00"));
    }

    #[test]
    fn linked_axes_take_the_formats_of_the_ranges_they_plot() {
        // The format of a range is its address, so the test reads which
        // range each axis took.
        struct Named;
        impl RangeResolver for Named {
            fn cells(&self, address: &str) -> Result<Option<Vec<SourceCell>>, String> {
                Fixed.cells(address)
            }
            fn format(&self, address: &str) -> Option<(String, Option<Vec<f64>>)> {
                Some((address.to_string(), None))
            }
        }
        let styles = r##"<style:style style:name="linked" style:family="chart"><style:chart-properties chart:link-data-style-to-source="true"/></style:style>"##;
        let body = |class: &str| {
            format!(
                concat!(
                    r##"<chart:chart chart:class="chart:{class}"><chart:plot-area>"##,
                    r##"<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="linked"/>"##,
                    r##"<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="linked"/>"##,
                    r##"<chart:series chart:values-cell-range-address="S.C1:S.C2">"##,
                    r##"<chart:domain table:cell-range-address="S.A1:S.A2"/><chart:domain table:cell-range-address="S.B1:S.B2"/>"##,
                    r##"</chart:series></chart:plot-area></chart:chart>"##
                ),
                class = class
            )
        };
        let axes = |class: &str| {
            let model = read_chart_document(&chart(&body(class), styles), None, Some(&Named))
                .unwrap()
                .unwrap();
            (
                model.cat_axis_format_code.unwrap_or_default(),
                model.val_axis_format_code.unwrap_or_default(),
            )
        };
        // §11.13: a scatter series' domain is its x values, and its values
        // its y values; a bubble series has its y values then its x values
        // as domains, and its values are the sizes.
        assert_eq!(axes("scatter"), ("S.A1:S.A2".into(), "S.C1:S.C2".into()));
        assert_eq!(axes("bubble"), ("S.B1:S.B2".into(), "S.A1:S.A2".into()));
    }

    #[test]
    fn hidden_cells_axis_names_and_attribute_strings() {
        // Two y axes with free names: the second is the secondary axis and
        // the series attached to it uses it. The local label cell holds its
        // text in `office:string-value`.
        let local = concat!(
            r##"<table:table table:name="local-table"><table:table-rows>"##,
            r##"<table:table-row><table:table-cell office:value-type="string" office:string-value="North"/></table:table-row>"##,
            r##"<table:table-row><table:table-cell office:value-type="date" office:date-value="1900-03-01"/><table:table-cell office:value-type="time" office:time-value="PT12H"/><table:table-cell office:value-type="boolean" office:boolean-value="true"/></table:table-row>"##,
            r##"</table:table-rows></table:table>"##
        );
        // Local dates, times and booleans are numbers.
        let typed = format!(
            r##"<chart:chart chart:class="chart:line"><chart:plot-area><chart:series chart:values-cell-range-address="local-table.A2:.C2"/></chart:plot-area>{local}</chart:chart>"##
        );
        let model = read_chart_document(&chart(&typed, ""), None, None)
            .unwrap()
            .unwrap();
        assert_eq!(model.series[0].values, [Some(61.0), Some(0.5), Some(1.0)]);
        let body = |plot_style: &str| {
            format!(
                r##"<chart:chart chart:class="chart:bar"><chart:plot-area {plot_style}><chart:axis chart:dimension="x" chart:name="bottom"/><chart:axis chart:dimension="y" chart:name="left" chart:style-name="left"/><chart:axis chart:dimension="y" chart:name="right" chart:style-name="right"/><chart:series chart:values-cell-range-address="Sheet1.B2:Sheet1.B3" chart:label-cell-address="local-table.A1"/><chart:series chart:attached-axis="right" chart:values-cell-range-address="Sheet1.H2:Sheet1.H3"/></chart:plot-area>{local}</chart:chart>"##
            )
        };
        let styles = concat!(
            r##"<style:style style:name="left" style:family="chart"><style:chart-properties chart:minimum="1"/></style:style>"##,
            r##"<style:style style:name="right" style:family="chart"><style:chart-properties chart:minimum="2"/></style:style>"##,
            r##"<style:style style:name="hide" style:family="chart"><style:chart-properties chart:include-hidden-cells="false"/></style:style>"##,
        );
        let model = read_chart_document(&chart(&body(""), styles), None, Some(&Fixed))
            .unwrap()
            .unwrap();
        assert_eq!(model.val_min, Some(1.0));
        assert_eq!(
            model.secondary_val_axis.as_ref().and_then(|axis| axis.min),
            Some(2.0)
        );
        assert_eq!(model.series[0].name, "North");
        assert_eq!(model.series[0].use_secondary_axis, None);
        assert_eq!(model.series[1].use_secondary_axis, Some(true));
        // §20.26: hidden cells are plotted by default; the model still
        // says which points they are.
        assert_eq!(model.plot_visible_only, Some(false));
        assert_eq!(model.series[0].source_hidden, None);
        assert_eq!(model.series[1].source_hidden, Some(vec![false, true]));
        let model = read_chart_document(
            &chart(&body(r#"chart:style-name="hide""#), styles),
            None,
            Some(&Fixed),
        )
        .unwrap()
        .unwrap();
        assert_eq!(model.plot_visible_only, Some(true));
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

    #[test]
    fn per_series_label_styles_are_held_per_series() {
        // One shared style with a 64 KiB label separator, used by 300
        // series: every series holds its own copy.
        let separator = "s".repeat(64 * 1024);
        let styles = format!(
            r##"<style:style style:name="S" style:family="chart"><style:chart-properties chart:data-label-number="value"><chart:label-separator><text:p>{separator}</text:p></chart:label-separator></style:chart-properties></style:style>"##
        );
        let series = |count: usize| {
            let series: String = (0..count)
                .map(|_| r##"<chart:series chart:style-name="S"/>"##)
                .collect();
            chart(
                &format!(
                    r##"<chart:chart chart:class="chart:bar"><chart:plot-area>{series}</chart:plot-area></chart:chart>"##
                ),
                &styles,
            )
        };
        let model = read_chart_document(&series(2), None, None)
            .unwrap()
            .unwrap();
        assert_eq!(
            model.series[1]
                .series_data_labels
                .as_ref()
                .and_then(|l| l.separator.as_deref())
                .map(str::len),
            Some(64 * 1024)
        );
        let _scope = Scope::enter(8 << 20);
        assert_eq!(
            read_chart_document(&series(300), None, None).err(),
            Some(chart_limit("model size"))
        );
    }

    #[test]
    fn scatter_and_bubble_coordinates_are_numbers_and_domains_are_inherited() {
        // The x cells display as "1,000" and "50%"; the chart's cache holds
        // those texts for the source range.
        let local = concat!(
            r##"<table:table table:name="local-table"><table:table-rows>"##,
            r##"<table:table-row><table:table-cell office:value-type="string"><text:p>1,000</text:p><draw:g><svg:desc>Sheet1.X2:Sheet1.X3</svg:desc></draw:g></table:table-cell></table:table-row>"##,
            r##"<table:table-row><table:table-cell office:value-type="string"><text:p>50%</text:p></table:table-cell></table:table-row>"##,
            r##"</table:table-rows></table:table>"##
        );
        let scatter = format!(
            r##"<chart:chart chart:class="chart:scatter"><chart:plot-area><chart:series chart:values-cell-range-address="Sheet1.B2:Sheet1.B3"><chart:domain table:cell-range-address="Sheet1.X2:Sheet1.X3"/></chart:series><chart:series chart:values-cell-range-address="Sheet1.Y2:Sheet1.Y3"/></chart:plot-area>{local}</chart:chart>"##
        );
        let model = read_chart_document(&chart(&scatter, ""), None, Some(&Fixed))
            .unwrap()
            .unwrap();
        let x = Some(vec!["1000".to_string(), "0.5".to_string()]);
        assert_eq!(model.series[0].categories, x);
        // §11.13: the second series uses the first one's x values.
        assert_eq!(model.series[1].categories, x);
        assert_eq!(model.series[1].values, [Some(3.0), Some(4.0)]);
        // A bubble series' first domain is y, its second x; a later series
        // with only its y domain keeps the x values.
        let bubble = format!(
            r##"<chart:chart chart:class="chart:bubble"><chart:plot-area><chart:series chart:values-cell-range-address="Sheet1.B2:Sheet1.B3"><chart:domain table:cell-range-address="Sheet1.Y2:Sheet1.Y3"/><chart:domain table:cell-range-address="Sheet1.X2:Sheet1.X3"/></chart:series><chart:series chart:values-cell-range-address="Sheet1.Y2:Sheet1.Y3"><chart:domain table:cell-range-address="Sheet1.B2:Sheet1.B3"/></chart:series><chart:series chart:values-cell-range-address="Sheet1.B2:Sheet1.B3"/></chart:plot-area>{local}</chart:chart>"##
        );
        let model = read_chart_document(&chart(&bubble, ""), None, Some(&Fixed))
            .unwrap()
            .unwrap();
        assert_eq!(model.series[0].values, [Some(3.0), Some(4.0)]);
        assert_eq!(
            model.series[0].bubble_sizes,
            Some(vec![Some(1.0), Some(2.0)])
        );
        assert_eq!(model.series[1].values, [Some(1.0), Some(2.0)]);
        assert_eq!(model.series[1].categories, x);
        assert_eq!(model.series[2].values, [Some(1.0), Some(2.0)]);
        assert_eq!(model.series[2].categories, x);
    }

    #[test]
    fn chart_repeat_counts_past_the_bound_reject() {
        let rows = r##"<table:table table:name="local-table"><table:table-rows><table:table-row table:number-rows-repeated="4294967296"><table:table-cell/></table:table-row></table:table-rows></table:table>"##;
        assert_eq!(
            read_chart_document(&chart(&line_series("local-table.A1", rows), ""), None, None).err(),
            Some(crate::error::resource("repeat count"))
        );
        let points = r##"<chart:chart chart:class="chart:bar"><chart:plot-area><chart:series><chart:data-point chart:repeated="18446744073709551616"/></chart:series></chart:plot-area></chart:chart>"##;
        assert_eq!(
            read_chart_document(&chart(points, ""), None, None).err(),
            Some(crate::error::resource("repeat count"))
        );
    }

    #[test]
    fn display_factors_reach_chart_axes_and_labels() {
        let number = r##"xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0""##;
        let styles = format!(
            concat!(
                r##"<number:number-style {number} style:name="N10"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="10"/></number:number-style>"##,
                r##"<number:number-style {number} style:name="GK"><number:number number:display-factor="1000"/></number:number-style>"##,
                r##"<style:style style:name="ay" style:family="chart" style:data-style-name="N10"><style:chart-properties chart:link-data-style-to-source="false"/></style:style>"##,
                r##"<style:style style:name="s1" style:family="chart" style:data-style-name="GK"><style:chart-properties chart:data-label-number="value" chart:link-data-style-to-source="false"/></style:style>"##,
            ),
            number = number
        );
        let body = format!(
            r##"<chart:chart chart:class="chart:bar"><chart:plot-area><chart:axis chart:dimension="x" chart:name="primary-x"><chart:categories table:cell-range-address="local-table.A2:.A3"/></chart:axis><chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="ay"/><chart:series chart:style-name="s1" chart:values-cell-range-address="local-table.B2:.B3"/></chart:plot-area>{LOCAL}</chart:chart>"##
        );
        let model = read_chart_document(&chart(&body, &styles), None, None)
            .unwrap()
            .unwrap();
        // Any factor, not only powers of 1000, divides as in cells.
        assert_eq!(model.val_axis_format_code.as_deref(), Some("0"));
        assert_eq!(model.val_axis_format_divisors, Some(vec![10.0]));
        // A General code keeps its factor.
        let labels = model.series[0].series_data_labels.as_ref().unwrap();
        assert_eq!(labels.format_code.as_deref(), Some("General"));
        assert_eq!(labels.format_divisors, Some(vec![1000.0]));
    }
}
