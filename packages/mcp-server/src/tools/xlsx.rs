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

/// Returns the display string for a cell value from a `cell` JSON object.
fn cell_display(cell: &Value) -> String {
    let val = &cell["value"];
    match val["type"].as_str().unwrap_or("Empty") {
        "Text" => val["text"].as_str().unwrap_or("").to_string(),
        "Number" => val["number"]
            .as_f64()
            .map(|n| {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", n as i64)
                } else {
                    format!("{}", n)
                }
            })
            .unwrap_or_default(),
        "Bool" => val["value"].as_bool().map(|b| b.to_string()).unwrap_or_default(),
        "Error" => val["error"].as_str().unwrap_or("#ERR").to_string(),
        _ => String::new(),
    }
}

// ─── Tool implementations ─────────────────────────────────────────────────────

pub struct XlsxTools;

impl XlsxTools {
    #[tool(description = "Parse an XLSX file and return workbook overview including sheet names and IDs")]
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

    #[tool(description = "Return only the list of sheet names from an XLSX file")]
    pub fn xlsx_get_sheet_names(Parameters(p): Parameters<XlsxPathParam>) -> String {
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
            Err(e) => return format!("Error parsing workbook JSON: {}", e),
        };
        let names: Vec<&str> = wb["sheets"]
            .as_array()
            .map(|sheets| {
                sheets
                    .iter()
                    .filter_map(|s| s["name"].as_str())
                    .collect()
            })
            .unwrap_or_default();
        serde_json::to_string(&names).unwrap_or_else(|e| format!("Error: {}", e))
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

    #[tool(description = "Return cell values and formulas for a given range (e.g. \"A1:C10\") in a worksheet")]
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

    #[tool(description = "Search for a substring in cell values and formulas across one or all sheets of an XLSX file")]
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
}
