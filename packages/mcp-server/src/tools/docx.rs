use rmcp::{handler::server::wrapper::Parameters, tool};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;
use std::fs;

// ─── Parameter types ─────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxPathParam {
    /// Absolute path to the DOCX file
    pub path: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxSearchParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// Case-insensitive substring to search for in paragraph and table cell text
    pub query: String,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn read_file(path: &str) -> Result<Vec<u8>, String> {
    fs::read(path).map_err(|e| format!("Cannot read '{}': {}", path, e))
}

fn collect_run_texts(runs: &Value, out: &mut String) {
    if let Some(arr) = runs.as_array() {
        for run in arr {
            if let Some(t) = run["text"].as_str() {
                out.push_str(t);
            }
        }
    }
}

fn collect_paragraph_text(para: &Value, out: &mut String) {
    collect_run_texts(&para["runs"], out);
    out.push('\n');
}

fn collect_table_text(table: &Value, out: &mut String) {
    if let Some(rows) = table["rows"].as_array() {
        for row in rows {
            if let Some(cells) = row["cells"].as_array() {
                let cell_texts: Vec<String> = cells
                    .iter()
                    .map(|cell| {
                        let mut t = String::new();
                        if let Some(paras) = cell["paragraphs"].as_array() {
                            for p in paras {
                                collect_run_texts(&p["runs"], &mut t);
                            }
                        }
                        t
                    })
                    .collect();
                out.push_str(&cell_texts.join("\t"));
                out.push('\n');
            }
        }
    }
}

fn extract_body_text(body: &[Value]) -> String {
    let mut out = String::new();
    for element in body {
        match element["type"].as_str().unwrap_or("") {
            "paragraph" => collect_paragraph_text(element, &mut out),
            "table" => collect_table_text(element, &mut out),
            _ => {}
        }
    }
    out
}

fn body_structure(body: &[Value]) -> Vec<Value> {
    body.iter()
        .map(|el| match el["type"].as_str().unwrap_or("") {
            "paragraph" => {
                let mut runs_text = String::new();
                collect_run_texts(&el["runs"], &mut runs_text);
                serde_json::json!({
                    "type": "paragraph",
                    "styleId": el["styleId"],
                    "text": runs_text.trim().to_string(),
                    "alignment": el["alignment"],
                })
            }
            "table" => {
                let rows = el["rows"]
                    .as_array()
                    .map(|r| r.len())
                    .unwrap_or(0);
                let cols = el["rows"]
                    .as_array()
                    .and_then(|r| r.first())
                    .and_then(|r| r["cells"].as_array())
                    .map(|c| c.len())
                    .unwrap_or(0);
                serde_json::json!({
                    "type": "table",
                    "rows": rows,
                    "cols": cols,
                })
            }
            _ => el.clone(),
        })
        .collect()
}

// ─── Tool implementations ─────────────────────────────────────────────────────

pub struct DocxTools;

impl DocxTools {
    #[tool(description = "Extract all plain text from a DOCX file")]
    pub fn docx_extract_text(Parameters(p): Parameters<DocxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let doc_json = match docx_parser::parse_docx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let doc: Value = match serde_json::from_str(&doc_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let body = doc["body"].as_array().map(|a| a.as_slice()).unwrap_or(&[]);
        extract_body_text(body)
    }

    #[tool(description = "Return the document structure (paragraphs and tables) of a DOCX file")]
    pub fn docx_get_structure(Parameters(p): Parameters<DocxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let doc_json = match docx_parser::parse_docx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let doc: Value = match serde_json::from_str(&doc_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let body = doc["body"].as_array().map(|a| a.as_slice()).unwrap_or(&[]);
        let structure = body_structure(body);
        serde_json::to_string(&structure).unwrap_or_else(|e| format!("Error: {}", e))
    }

    #[tool(description = "Return all tables from a DOCX file with their cell contents")]
    pub fn docx_get_tables(Parameters(p): Parameters<DocxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let doc_json = match docx_parser::parse_docx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let doc: Value = match serde_json::from_str(&doc_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let body = doc["body"].as_array().map(|a| a.as_slice()).unwrap_or(&[]);

        let tables: Vec<Value> = body
            .iter()
            .filter(|el| el["type"].as_str() == Some("table"))
            .enumerate()
            .map(|(table_idx, table)| {
                let rows = table["rows"]
                    .as_array()
                    .map(|rows| {
                        rows.iter()
                            .map(|row| {
                                row["cells"]
                                    .as_array()
                                    .map(|cells| {
                                        cells
                                            .iter()
                                            .map(|cell| {
                                                let mut text = String::new();
                                                if let Some(paras) =
                                                    cell["paragraphs"].as_array()
                                                {
                                                    for p in paras {
                                                        collect_run_texts(&p["runs"], &mut text);
                                                    }
                                                }
                                                Value::String(text)
                                            })
                                            .collect::<Vec<_>>()
                                    })
                                    .unwrap_or_default()
                            })
                            .collect::<Vec<_>>()
                    })
                    .unwrap_or_default();
                serde_json::json!({ "tableIndex": table_idx, "rows": rows })
            })
            .collect();

        serde_json::to_string(&tables).unwrap_or_else(|e| format!("Error: {}", e))
    }

    #[tool(description = "Search for a substring in all paragraph and table text of a DOCX file; returns matching excerpts with their position")]
    pub fn docx_search_text(Parameters(p): Parameters<DocxSearchParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let doc_json = match docx_parser::parse_docx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let doc: Value = match serde_json::from_str(&doc_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let body = doc["body"].as_array().map(|a| a.as_slice()).unwrap_or(&[]);
        let query_lower = p.query.to_lowercase();
        let mut matches: Vec<Value> = Vec::new();

        for (idx, element) in body.iter().enumerate() {
            match element["type"].as_str().unwrap_or("") {
                "paragraph" => {
                    let mut text = String::new();
                    collect_run_texts(&element["runs"], &mut text);
                    if text.to_lowercase().contains(&query_lower) {
                        matches.push(serde_json::json!({
                            "type": "paragraph",
                            "index": idx,
                            "styleId": element["styleId"],
                            "text": text.trim(),
                        }));
                    }
                }
                "table" => {
                    if let Some(rows) = element["rows"].as_array() {
                        for (row_idx, row) in rows.iter().enumerate() {
                            if let Some(cells) = row["cells"].as_array() {
                                for (col_idx, cell) in cells.iter().enumerate() {
                                    let mut text = String::new();
                                    if let Some(paras) = cell["paragraphs"].as_array() {
                                        for para in paras {
                                            collect_run_texts(&para["runs"], &mut text);
                                        }
                                    }
                                    if text.to_lowercase().contains(&query_lower) {
                                        matches.push(serde_json::json!({
                                            "type": "tableCell",
                                            "tableIndex": idx,
                                            "row": row_idx,
                                            "col": col_idx,
                                            "text": text.trim(),
                                        }));
                                    }
                                }
                            }
                        }
                    }
                }
                _ => {}
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
