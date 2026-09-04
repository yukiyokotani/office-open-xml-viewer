use rmcp::{handler::server::wrapper::Parameters, tool};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;
use std::fs;

// ─── Parameter types ─────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxPathParam {
    /// Absolute path to the XLSX file
    pub path: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxSearchParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index; omit to search all sheets
    pub sheet: Option<String>,
    /// Case-insensitive substring to search for in cell values and formulas
    pub query: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxSheetParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name (e.g. "Sheet1") or 0-based numeric index as a string (e.g. "0")
    pub sheet: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxCellRangeParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index
    pub sheet: String,
    /// Cell range in A1 notation, e.g. "A1:C10"
    pub range: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxOptSheetParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index; omit to scan all sheets
    pub sheet: Option<String>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct XlsxChartIndexParam {
    /// Absolute path to the XLSX file
    pub path: String,
    /// Sheet name or 0-based index
    pub sheet: String,
    /// 0-based chart index within the sheet (matches order in `xlsx_get_charts`)
    pub chart_index: usize,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn read_file(path: &str) -> Result<Vec<u8>, String> {
    fs::read(path).map_err(|e| format!("Cannot read '{}': {}", path, e))
}

/// Resolves a sheet identifier (name or 0-based index string) to (index, name).
fn resolve_sheet(workbook_json: &str, identifier: &str) -> Result<(u32, String), String> {
    let wb: Value = serde_json::from_str(workbook_json).map_err(|e| e.to_string())?;
    let sheets = wb["sheets"]
        .as_array()
        .ok_or("workbook has no 'sheets' array")?;

    if let Ok(idx) = identifier.parse::<usize>() {
        let sheet = sheets
            .get(idx)
            .ok_or_else(|| format!("sheet index {} out of range (total: {})", idx, sheets.len()))?;
        let name = sheet["name"].as_str().unwrap_or("").to_string();
        return Ok((idx as u32, name));
    }

    for (idx, sheet) in sheets.iter().enumerate() {
        if sheet["name"].as_str() == Some(identifier) {
            return Ok((idx as u32, identifier.to_string()));
        }
    }

    Err(format!(
        "sheet '{}' not found (available: {})",
        identifier,
        sheets
            .iter()
            .filter_map(|s| s["name"].as_str())
            .collect::<Vec<_>>()
            .join(", ")
    ))
}

/// Parses an A1-style cell reference to (col_1based, row_1based).
fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let col_str: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let row_str: String = s.chars().skip_while(|c| c.is_ascii_alphabetic()).collect();
    if col_str.is_empty() || row_str.is_empty() {
        return None;
    }
    let col = col_str
        .to_ascii_uppercase()
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row: u32 = row_str.parse().ok()?;
    Some((col, row))
}

/// Converts a 1-based column index to a letter reference (1→"A", 26→"Z", 27→"AA").
fn col_to_letter(mut col: u32) -> String {
    let mut s = String::new();
    while col > 0 {
        col -= 1;
        s.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    s
}

/// Resolves which sheets to operate on. When `identifier` is `Some`, returns a
/// single-element vec; when `None`, returns every sheet in workbook order.
fn target_sheets(
    workbook_json: &str,
    identifier: Option<&str>,
) -> Result<Vec<(u32, String)>, String> {
    if let Some(id) = identifier {
        return resolve_sheet(workbook_json, id).map(|r| vec![r]);
    }
    let wb: Value = serde_json::from_str(workbook_json).map_err(|e| e.to_string())?;
    let sheets = wb["sheets"]
        .as_array()
        .ok_or("workbook has no 'sheets' array")?;
    Ok(sheets
        .iter()
        .enumerate()
        .map(|(i, s)| (i as u32, s["name"].as_str().unwrap_or("").to_string()))
        .collect())
}

/// Formats a `MergeCell` JSON object as an A1 range string ("A1:B2").
fn merge_to_a1(merge: &Value) -> Option<String> {
    let top = merge["top"].as_u64()? as u32;
    let left = merge["left"].as_u64()? as u32;
    let bottom = merge["bottom"].as_u64()? as u32;
    let right = merge["right"].as_u64()? as u32;
    Some(format!(
        "{}{}:{}{}",
        col_to_letter(left),
        top,
        col_to_letter(right),
        bottom,
    ))
}

/// Formats a `CellRange` JSON object as an A1 range string. CellRange uses the
/// same `top`/`left`/`bottom`/`right` shape as MergeCell so this is a thin alias.
fn range_to_a1(range: &Value) -> Option<String> {
    merge_to_a1(range)
}

/// Returns the display string for a cell value from a `cell` JSON object.
/// CellValue uses `rename_all = "camelCase"` on its enum tag, so the JSON
/// type field is lower-case ("text"/"number"/"bool"/"error"/"empty") — *not*
/// PascalCase. An earlier revision matched "Text"/"Number", which never fired
/// and silently dropped every non-empty cell value through `xlsx_get_*`.
fn cell_display(cell: &Value) -> String {
    let val = &cell["value"];
    match val["type"].as_str().unwrap_or("empty") {
        "text" => val["text"].as_str().unwrap_or("").to_string(),
        "number" => val["number"]
            .as_f64()
            .map(|n| {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", n as i64)
                } else {
                    format!("{}", n)
                }
            })
            .unwrap_or_default(),
        // Bool variant carries the value under the `bool` field (renamed by
        // serde from the variant's inner field name).
        "bool" => val["bool"]
            .as_bool()
            .map(|b| {
                if b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            })
            .unwrap_or_default(),
        "error" => val["error"].as_str().unwrap_or("#ERR").to_string(),
        _ => String::new(),
    }
}

// ─── Tool implementations ─────────────────────────────────────────────────────

pub struct XlsxTools;

impl XlsxTools {
    #[tool(
        description = "Convert an XLSX file to GitHub-flavoured markdown — one `## SheetName` per sheet, followed by a pipe table of the cells' cached display values. Merged-cell continuation cells render empty. Formula cells show the cached result, not the formula text (use `xlsx_get_formulas` if you need the formulas). Designed for agents that need to *read* spreadsheet content efficiently. Lossy by design: drops styling, conditional formatting, charts, sparklines, drawings. For precise structure use the structured tools (`xlsx_get_cell_range`, `xlsx_get_sheet_layout`, etc.)"
    )]
    pub fn xlsx_to_markdown(Parameters(p): Parameters<XlsxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        match xlsx_parser::to_markdown_native(&data) {
            Ok(md) => md,
            Err(e) => format!("Error: {}", e),
        }
    }

    #[tool(
        description = "Parse an XLSX file and return workbook overview including sheet names and IDs"
    )]
    pub fn xlsx_parse(Parameters(p): Parameters<XlsxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        match xlsx_parser::parse_workbook_native(&data) {
            Ok(json) => json,
            Err(e) => format!("Error: {}", e),
        }
    }

    #[tool(description = "Return the dimensions (max row and column) of a worksheet")]
    pub fn xlsx_get_sheet_dimensions(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };

        let mut max_row = 0u32;
        let mut max_col = 0u32;
        if let Some(rows) = ws["rows"].as_array() {
            for row in rows {
                let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
                if row_idx > max_row {
                    max_row = row_idx;
                }
                if let Some(cells) = row["cells"].as_array() {
                    for cell in cells {
                        let col = cell["col"].as_u64().unwrap_or(0) as u32;
                        if col > max_col {
                            max_col = col;
                        }
                    }
                }
            }
        }

        serde_json::json!({
            "sheet": name,
            "maxRow": max_row,
            "maxCol": max_col,
            "maxColLetter": col_to_letter(max_col),
        })
        .to_string()
    }

    #[tool(
        description = "Return cell values and formulas for a given range (e.g. \"A1:C10\") in a worksheet"
    )]
    pub fn xlsx_get_cell_range(Parameters(p): Parameters<XlsxCellRangeParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };

        // Parse range "A1:C10"
        let parts: Vec<&str> = p.range.split(':').collect();
        if parts.len() != 2 {
            return format!("Error: range must be in 'A1:C10' format, got '{}'", p.range);
        }
        let (c1, r1) = match parse_cell_ref(parts[0]) {
            Some(v) => v,
            None => return format!("Error: invalid cell reference '{}'", parts[0]),
        };
        let (c2, r2) = match parse_cell_ref(parts[1]) {
            Some(v) => v,
            None => return format!("Error: invalid cell reference '{}'", parts[1]),
        };
        let (row_min, row_max) = (r1.min(r2), r1.max(r2));
        let (col_min, col_max) = (c1.min(c2), c1.max(c2));

        let mut result_rows: Vec<Value> = Vec::new();

        if let Some(rows) = ws["rows"].as_array() {
            for row in rows {
                let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
                if row_idx < row_min || row_idx > row_max {
                    continue;
                }
                let mut result_cells: Vec<Value> = Vec::new();
                if let Some(cells) = row["cells"].as_array() {
                    for cell in cells {
                        let col = cell["col"].as_u64().unwrap_or(0) as u32;
                        if col < col_min || col > col_max {
                            continue;
                        }
                        let mut entry = serde_json::json!({
                            "ref": format!("{}{}", col_to_letter(col), row_idx),
                            "value": cell_display(cell),
                        });
                        if let Some(formula) = cell["formula"].as_str() {
                            entry["formula"] = Value::String(formula.to_string());
                        }
                        result_cells.push(entry);
                    }
                }
                result_rows.push(serde_json::json!({
                    "row": row_idx,
                    "cells": result_cells,
                }));
            }
        }

        serde_json::json!({
            "sheet": name,
            "range": p.range,
            "rows": result_rows,
        })
        .to_string()
    }

    #[tool(description = "Return all cells that contain formulas in a worksheet")]
    pub fn xlsx_get_formulas(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };

        let mut formulas: Vec<Value> = Vec::new();
        if let Some(rows) = ws["rows"].as_array() {
            for row in rows {
                let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
                if let Some(cells) = row["cells"].as_array() {
                    for cell in cells {
                        if let Some(formula) = cell["formula"].as_str() {
                            let col = cell["col"].as_u64().unwrap_or(0) as u32;
                            formulas.push(serde_json::json!({
                                "ref": format!("{}{}", col_to_letter(col), row_idx),
                                "formula": formula,
                                "cachedValue": cell_display(cell),
                            }));
                        }
                    }
                }
            }
        }

        serde_json::json!({
            "sheet": name,
            "formulas": formulas,
        })
        .to_string()
    }

    #[tool(
        description = "Search for a substring in cell values and formulas across one or all sheets of an XLSX file"
    )]
    pub fn xlsx_search_cells(Parameters(p): Parameters<XlsxSearchParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let wb: Value = match serde_json::from_str(&wb_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let sheets = match wb["sheets"].as_array() {
            Some(s) => s.clone(),
            None => return "Error: no sheets found".to_string(),
        };

        // Collect which sheets to search
        let targets: Vec<(u32, String)> = if let Some(ref sheet_id) = p.sheet {
            match resolve_sheet(&wb_json, sheet_id) {
                Ok((idx, name)) => vec![(idx, name)],
                Err(e) => return format!("Error: {}", e),
            }
        } else {
            sheets
                .iter()
                .enumerate()
                .map(|(i, s)| (i as u32, s["name"].as_str().unwrap_or("").to_string()))
                .collect()
        };

        let query_lower = p.query.to_lowercase();
        let mut matches: Vec<Value> = Vec::new();

        for (idx, name) in targets {
            let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
                Ok(j) => j,
                Err(e) => return format!("Error parsing sheet '{}': {}", name, e),
            };
            let ws: Value = match serde_json::from_str(&ws_json) {
                Ok(v) => v,
                Err(e) => return format!("Error: {}", e),
            };

            if let Some(rows) = ws["rows"].as_array() {
                for row in rows {
                    let row_idx = row["index"].as_u64().unwrap_or(0) as u32;
                    if let Some(cells) = row["cells"].as_array() {
                        for cell in cells {
                            let value = cell_display(cell);
                            let formula = cell["formula"].as_str().unwrap_or("");
                            let hit_value = value.to_lowercase().contains(&query_lower);
                            let hit_formula = formula.to_lowercase().contains(&query_lower);
                            if hit_value || hit_formula {
                                let col = cell["col"].as_u64().unwrap_or(0) as u32;
                                let mut entry = serde_json::json!({
                                    "sheet": name,
                                    "ref": format!("{}{}", col_to_letter(col), row_idx),
                                    "value": value,
                                });
                                if !formula.is_empty() {
                                    entry["formula"] = Value::String(formula.to_string());
                                }
                                matches.push(entry);
                            }
                        }
                    }
                }
            }
        }

        serde_json::json!({
            "query": p.query,
            "matchCount": matches.len(),
            "matches": matches,
        })
        .to_string()
    }

    #[tool(
        description = "List charts on a worksheet (or all sheets if `sheet` is omitted). Returns a summary per chart: anchor cell range, chart type, title, axes, legend, and a series outline (without numeric values)"
    )]
    pub fn xlsx_get_charts(Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let targets = match target_sheets(&wb_json, p.sheet.as_deref()) {
            Ok(t) => t,
            Err(e) => return format!("Error: {}", e),
        };

        let mut all_charts: Vec<Value> = Vec::new();
        for (idx, name) in targets {
            let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
                Ok(j) => j,
                Err(e) => return format!("Error parsing sheet '{}': {}", name, e),
            };
            let ws: Value = match serde_json::from_str(&ws_json) {
                Ok(v) => v,
                Err(e) => return format!("Error: {}", e),
            };
            let Some(charts) = ws["charts"].as_array() else {
                continue;
            };
            for (chart_idx, anchor) in charts.iter().enumerate() {
                let chart = &anchor["chart"];
                let series_outline: Vec<Value> = chart["series"]
                    .as_array()
                    .map(|arr| {
                        arr.iter()
                            .map(|s| {
                                let value_count =
                                    s["values"].as_array().map(|a| a.len()).unwrap_or(0);
                                serde_json::json!({
                                    "name": s["name"],
                                    "type": s["seriesType"],
                                    "color": s["color"],
                                    "showMarker": s["showMarker"],
                                    "valueCount": value_count,
                                })
                            })
                            .collect()
                    })
                    .unwrap_or_default();
                all_charts.push(serde_json::json!({
                    "sheet": name,
                    "chartIndex": chart_idx,
                    "anchor": {
                        "from": { "col": anchor["fromCol"], "row": anchor["fromRow"] },
                        "to":   { "col": anchor["toCol"],   "row": anchor["toRow"] },
                    },
                    "type": chart["chartType"],
                    "barDir": chart["barDir"],
                    "grouping": chart["grouping"],
                    "title": chart["title"],
                    "legend": {
                        "show": chart["showLegend"],
                        "position": chart["legendPos"],
                    },
                    "axes": {
                        "cat": {
                            "title": chart["catAxisTitle"],
                            "formatCode": chart["catAxisFormatCode"],
                            "hidden": chart["catAxisHidden"],
                        },
                        "val": {
                            "title": chart["valAxisTitle"],
                            "formatCode": chart["valAxisFormatCode"],
                            "hidden": chart["valAxisHidden"],
                        },
                    },
                    "categories": chart["categories"],
                    "seriesCount": series_outline.len(),
                    "series": series_outline,
                }));
            }
        }

        serde_json::json!({ "charts": all_charts }).to_string()
    }

    #[tool(
        description = "Return one chart's full series data (categories and per-point values) for drill-down. `chart_index` matches the index from `xlsx_get_charts` for the same sheet"
    )]
    pub fn xlsx_get_chart_series(Parameters(p): Parameters<XlsxChartIndexParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let charts = match ws["charts"].as_array() {
            Some(c) => c,
            None => return format!("Error: sheet '{}' has no charts", name),
        };
        let anchor = match charts.get(p.chart_index) {
            Some(a) => a,
            None => {
                return format!(
                    "Error: chart index {} out of range (total: {})",
                    p.chart_index,
                    charts.len()
                )
            }
        };
        let chart = &anchor["chart"];
        let series: Vec<Value> = chart["series"]
            .as_array()
            .map(|arr| {
                arr.iter()
                    .map(|s| {
                        serde_json::json!({
                            "name": s["name"],
                            "type": s["seriesType"],
                            "color": s["color"],
                            "values": s["values"],
                            "categories": s["categories"],
                            "valFormatCode": s["valFormatCode"],
                        })
                    })
                    .collect()
            })
            .unwrap_or_default();

        serde_json::json!({
            "sheet": name,
            "chartIndex": p.chart_index,
            "type": chart["chartType"],
            "title": chart["title"],
            "categories": chart["categories"],
            "series": series,
        })
        .to_string()
    }

    #[tool(
        description = "Return all defined names (named ranges) visible in the workbook. Includes workbook-global names plus each sheet's local names; duplicates across sheets are merged"
    )]
    pub fn xlsx_get_named_ranges(Parameters(p): Parameters<XlsxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let targets = match target_sheets(&wb_json, None) {
            Ok(t) => t,
            Err(e) => return format!("Error: {}", e),
        };

        // (name, formula) → first sheet that exposed it
        let mut seen: Vec<(String, String, String)> = Vec::new();
        for (idx, name) in targets {
            let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
                Ok(j) => j,
                Err(e) => return format!("Error parsing sheet '{}': {}", name, e),
            };
            let ws: Value = match serde_json::from_str(&ws_json) {
                Ok(v) => v,
                Err(e) => return format!("Error: {}", e),
            };
            let Some(names) = ws["definedNames"].as_array() else {
                continue;
            };
            for dn in names {
                let n = dn["name"].as_str().unwrap_or("").to_string();
                let f = dn["formula"].as_str().unwrap_or("").to_string();
                if !seen.iter().any(|(nn, ff, _)| nn == &n && ff == &f) {
                    seen.push((n, f, name.clone()));
                }
            }
        }
        let defined_names: Vec<Value> = seen
            .into_iter()
            .map(|(n, f, sheet)| {
                serde_json::json!({
                    "name": n,
                    "refersTo": f,
                    "firstSeenSheet": sheet,
                })
            })
            .collect();
        serde_json::json!({ "definedNames": defined_names }).to_string()
    }

    #[tool(
        description = "List Excel Tables (Ctrl+T tables, ECMA-376 §18.5) on a sheet or across all sheets. Returns each table's range, style, header/totals row counts"
    )]
    pub fn xlsx_get_tables(Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let targets = match target_sheets(&wb_json, p.sheet.as_deref()) {
            Ok(t) => t,
            Err(e) => return format!("Error: {}", e),
        };

        let mut tables: Vec<Value> = Vec::new();
        for (idx, name) in targets {
            let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
                Ok(j) => j,
                Err(e) => return format!("Error parsing sheet '{}': {}", name, e),
            };
            let ws: Value = match serde_json::from_str(&ws_json) {
                Ok(v) => v,
                Err(e) => return format!("Error: {}", e),
            };
            let Some(arr) = ws["tables"].as_array() else {
                continue;
            };
            for t in arr {
                tables.push(serde_json::json!({
                    "sheet": name,
                    "range": range_to_a1(&t["range"]),
                    "styleName": t["styleName"],
                    "headerRowCount": t["headerRowCount"],
                    "totalsRowCount": t["totalsRowCount"],
                    "showRowStripes": t["showRowStripes"],
                    "showColumnStripes": t["showColumnStripes"],
                }));
            }
        }
        serde_json::json!({ "tables": tables }).to_string()
    }

    #[tool(
        description = "Return all merged cell ranges on a worksheet as A1 strings (e.g. \"A1:B2\")"
    )]
    pub fn xlsx_get_merged_cells(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let merges: Vec<String> = ws["mergeCells"]
            .as_array()
            .map(|arr| arr.iter().filter_map(merge_to_a1).collect())
            .unwrap_or_default();
        serde_json::json!({ "sheet": name, "merges": merges }).to_string()
    }

    #[tool(
        description = "Return conditional formatting rules on a worksheet. Each entry has the affected ranges (sqref) and the rule body (CellIs, Expression, ColorScale, DataBar, Top10, AboveAverage, IconSet, Other)"
    )]
    pub fn xlsx_get_conditional_formats(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };

        let formats: Vec<Value> = ws["conditionalFormats"]
            .as_array()
            .map(|arr| {
                arr.iter()
                    .map(|cf| {
                        let ranges: Vec<String> = cf["sqref"]
                            .as_array()
                            .map(|rs| rs.iter().filter_map(range_to_a1).collect())
                            .unwrap_or_default();
                        serde_json::json!({
                            "ranges": ranges,
                            "rules": cf["rules"],
                        })
                    })
                    .collect()
            })
            .unwrap_or_default();

        serde_json::json!({ "sheet": name, "formats": formats }).to_string()
    }

    #[tool(
        description = "Return all `<dataValidation>` rules on a worksheet: affected ranges, type, operator, formulas, and the optional prompt / error messages"
    )]
    pub fn xlsx_get_data_validations(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        serde_json::json!({
            "sheet": name,
            "validations": ws["dataValidations"].as_array().cloned().unwrap_or_default(),
        })
        .to_string()
    }

    #[tool(
        description = "Return all comments on a worksheet (or all sheets if `sheet` is omitted) with full text and resolved author. Each entry: { sheet, cellRef, author?, text }"
    )]
    pub fn xlsx_get_comments(Parameters(p): Parameters<XlsxOptSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let targets = match target_sheets(&wb_json, p.sheet.as_deref()) {
            Ok(t) => t,
            Err(e) => return format!("Error: {}", e),
        };

        let mut all_comments: Vec<Value> = Vec::new();
        for (idx, name) in targets {
            let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
                Ok(j) => j,
                Err(e) => return format!("Error parsing sheet '{}': {}", name, e),
            };
            let ws: Value = match serde_json::from_str(&ws_json) {
                Ok(v) => v,
                Err(e) => return format!("Error: {}", e),
            };
            let Some(comments) = ws["comments"].as_array() else {
                continue;
            };
            for c in comments {
                all_comments.push(serde_json::json!({
                    "sheet": name,
                    "cellRef": c["cellRef"],
                    "author": c["author"],
                    "text": c["text"],
                }));
            }
        }
        serde_json::json!({ "comments": all_comments }).to_string()
    }

    #[tool(
        description = "Return per-sheet layout: explicit column widths, row heights, freeze panes, gridline visibility, default sizes, and tab color"
    )]
    pub fn xlsx_get_sheet_layout(Parameters(p): Parameters<XlsxSheetParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let wb_json = match xlsx_parser::parse_workbook_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let (idx, name) = match resolve_sheet(&wb_json, &p.sheet) {
            Ok(r) => r,
            Err(e) => return format!("Error: {}", e),
        };
        let ws_json = match xlsx_parser::parse_sheet_native(&data, idx, &name) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let ws: Value = match serde_json::from_str(&ws_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };

        // colWidths / rowHeights are HashMap<u32, f64> — serialised as a JSON
        // object with numeric-string keys. Sort by numeric key for stable output.
        let mut cols: Vec<Value> = Vec::new();
        if let Some(map) = ws["colWidths"].as_object() {
            let mut entries: Vec<(u32, f64)> = map
                .iter()
                .filter_map(|(k, v)| Some((k.parse().ok()?, v.as_f64()?)))
                .collect();
            entries.sort_by_key(|(k, _)| *k);
            for (col, width) in entries {
                cols.push(
                    serde_json::json!({ "col": col, "width": width, "letter": col_to_letter(col) }),
                );
            }
        }
        let mut rows: Vec<Value> = Vec::new();
        if let Some(map) = ws["rowHeights"].as_object() {
            let mut entries: Vec<(u32, f64)> = map
                .iter()
                .filter_map(|(k, v)| Some((k.parse().ok()?, v.as_f64()?)))
                .collect();
            entries.sort_by_key(|(k, _)| *k);
            for (row, height) in entries {
                rows.push(serde_json::json!({ "row": row, "height": height }));
            }
        }

        serde_json::json!({
            "sheet": name,
            "defaultColWidth": ws["defaultColWidth"],
            "defaultRowHeight": ws["defaultRowHeight"],
            "freeze": { "rows": ws["freezeRows"], "cols": ws["freezeCols"] },
            "showGridlines": ws["showGridlines"],
            "showZeros": ws["showZeros"],
            "tabColor": ws["tabColor"],
            "colWidths": cols,
            "rowHeights": rows,
        })
        .to_string()
    }
}

#[cfg(test)]
mod sample_tests {
    use super::*;

    fn sample_path() -> String {
        format!(
            "{}/../xlsx/public/demo/sample-1.xlsx",
            env!("CARGO_MANIFEST_DIR")
        )
    }

    fn pp(path: &str) -> Parameters<XlsxPathParam> {
        Parameters(XlsxPathParam { path: path.into() })
    }

    #[test]
    fn xlsx_to_markdown_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_to_markdown(pp(&path));
        assert!(!out.starts_with("Error:"), "errored: {out}");
        assert!(
            out.contains("## "),
            "missing sheet heading: {}",
            &out[..200.min(out.len())]
        );
        assert!(out.contains("|"), "no pipe table emitted");
    }

    #[test]
    fn xlsx_parse_sample_returns_workbook() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return; // sample missing in this checkout — skip.
        }
        let out = XlsxTools::xlsx_parse(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("xlsx_parse must return JSON");
        assert!(
            v["sheets"].as_array().is_some(),
            "missing 'sheets' array: {out}"
        );
    }

    #[test]
    fn xlsx_get_charts_sample_returns_charts_field() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_charts(Parameters(XlsxOptSheetParam {
            path: path.clone(),
            sheet: None,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["charts"].as_array().is_some(),
            "missing 'charts' array: {out}"
        );
    }

    #[test]
    fn xlsx_get_named_ranges_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_named_ranges(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["definedNames"].as_array().is_some(),
            "missing 'definedNames'"
        );
    }

    #[test]
    fn xlsx_get_merged_cells_first_sheet() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_merged_cells(Parameters(XlsxSheetParam {
            path: path.clone(),
            sheet: "0".into(),
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["merges"].as_array().is_some(), "missing 'merges'");
    }

    #[test]
    fn xlsx_get_sheet_layout_first_sheet() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_sheet_layout(Parameters(XlsxSheetParam {
            path: path.clone(),
            sheet: "0".into(),
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["sheet"].is_string(), "missing 'sheet' name in {out}");
        assert!(v["colWidths"].as_array().is_some(), "missing 'colWidths'");
        assert!(v["rowHeights"].as_array().is_some(), "missing 'rowHeights'");
    }

    #[test]
    fn xlsx_get_data_validations_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_data_validations(Parameters(XlsxSheetParam {
            path,
            sheet: "0".into(),
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["validations"].as_array().is_some(),
            "missing 'validations'"
        );
    }

    #[test]
    fn xlsx_get_comments_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = XlsxTools::xlsx_get_comments(Parameters(XlsxOptSheetParam { path, sheet: None }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["comments"].as_array().is_some(), "missing 'comments'");
    }

    #[test]
    fn xlsx_invalid_path_returns_error_string() {
        let out = XlsxTools::xlsx_parse(pp("/nonexistent/does-not-exist.xlsx"));
        assert!(out.starts_with("Error:"), "expected error, got: {out}");
    }
}
