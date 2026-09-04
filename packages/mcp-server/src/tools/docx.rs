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

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxIndexParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// 0-based index into the body element list (paragraphs and tables share indexing).
    pub index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxTableIndexParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// 0-based index of the table (counts tables only, in document order)
    pub table_index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DocxImagesParam {
    /// Absolute path to the DOCX file
    pub path: String,
    /// When true include the base64 `dataUrl` for each image. Defaults to false
    /// (just the metadata) since image bytes are large and rarely needed inline.
    #[serde(default)]
    pub include_data_url: bool,
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
                let rows = el["rows"].as_array().map(|r| r.len()).unwrap_or(0);
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
    #[tool(
        description = "Convert a DOCX file to GitHub-flavoured markdown. Preserves textual structure (headings from outlineLevel, paragraphs, bullet/numbered lists, tables, footnotes, comments) and rich-text formatting (bold/italic/strikethrough/hyperlinks). Discards positioning, section properties, font metrics, drawing shapes, and headers/footers. Designed for agents that need to *read* the document content efficiently — typical 10×+ token reduction vs. the structured JSON tools. Lossy by design: when you need precise layout or styling, fall back to `docx_get_structure` / `docx_get_body_element`"
    )]
    pub fn docx_to_markdown(Parameters(p): Parameters<DocxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        match docx_parser::to_markdown_native(&data) {
            Ok(md) => md,
            Err(e) => format!("Error: {}", e),
        }
    }

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
                                                if let Some(paras) = cell["paragraphs"].as_array() {
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

    #[tool(
        description = "Search for a substring in all paragraph and table text of a DOCX file; returns matching excerpts with their position"
    )]
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

    #[tool(
        description = "Return one body element's full detail (paragraph or table) including run-level formatting (bold/italic/color/font/hyperlink), indents, spacing, numbering, and tab stops. `index` is into the document body list (matches `docx_get_structure`)"
    )]
    pub fn docx_get_body_element(Parameters(p): Parameters<DocxIndexParam>) -> String {
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
        let element = match body.get(p.index) {
            Some(e) => e,
            None => {
                return format!(
                    "Error: body index {} out of range (total: {})",
                    p.index,
                    body.len()
                )
            }
        };
        let mut out = element.clone();
        if let Some(obj) = out.as_object_mut() {
            obj.insert("index".to_string(), Value::from(p.index));
        }
        out.to_string()
    }

    #[tool(
        description = "Return the document's section properties (page size/margins/docGrid) along with default/first/even header and footer body elements"
    )]
    pub fn docx_get_sections(Parameters(p): Parameters<DocxPathParam>) -> String {
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
        serde_json::json!({
            "section": doc["section"],
            "headers": doc["headers"],
            "footers": doc["footers"],
            "majorFont": doc["majorFont"],
            "minorFont": doc["minorFont"],
        })
        .to_string()
    }

    #[tool(
        description = "Return one table's full detail by index, including cell content, colSpan/vMerge, borders, shading, and row heights. Use this for deeper inspection than `docx_get_tables`"
    )]
    pub fn docx_get_table(Parameters(p): Parameters<DocxTableIndexParam>) -> String {
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
        let mut count = 0usize;
        let mut found: Option<(usize, &Value)> = None;
        for (idx, el) in body.iter().enumerate() {
            if el["type"].as_str() == Some("table") {
                if count == p.table_index {
                    found = Some((idx, el));
                    break;
                }
                count += 1;
            }
        }
        let (body_index, table) = match found {
            Some(t) => t,
            None => {
                let total = body
                    .iter()
                    .filter(|e| e["type"].as_str() == Some("table"))
                    .count();
                return format!(
                    "Error: table index {} out of range (total tables: {})",
                    p.table_index, total
                );
            }
        };
        serde_json::json!({
            "tableIndex": p.table_index,
            "bodyIndex": body_index,
            "table": table,
        })
        .to_string()
    }

    #[tool(
        description = "List all images in the document. Each entry carries the paragraph index, anchor mode, wrap settings, and dimensions. Set `include_data_url=true` to also receive the inline base64 image bytes (large)"
    )]
    pub fn docx_get_images(Parameters(p): Parameters<DocxImagesParam>) -> String {
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
        let mut images: Vec<Value> = Vec::new();
        for (para_idx, element) in body.iter().enumerate() {
            if element["type"].as_str() != Some("paragraph") {
                continue;
            }
            let Some(runs) = element["runs"].as_array() else {
                continue;
            };
            for (run_idx, run) in runs.iter().enumerate() {
                if run["type"].as_str() != Some("image") {
                    continue;
                }
                let mut entry = serde_json::json!({
                    "paragraphIndex": para_idx,
                    "runIndex": run_idx,
                    "widthPt": run["widthPt"],
                    "heightPt": run["heightPt"],
                    "anchor": run["anchor"],
                    "anchorXPt": run["anchorXPt"],
                    "anchorYPt": run["anchorYPt"],
                    "wrapMode": run["wrapMode"],
                });
                if p.include_data_url {
                    if let Some(obj) = entry.as_object_mut() {
                        obj.insert("dataUrl".into(), run["dataUrl"].clone());
                    }
                }
                images.push(entry);
            }
        }
        serde_json::json!({ "images": images }).to_string()
    }

    #[tool(
        description = "List all drawn shapes embedded in paragraphs (wps:wsp inside wp:anchor). Returns each shape's preset geometry, fill, stroke, dimensions, anchor offsets, rotation, and embedded text blocks"
    )]
    pub fn docx_get_shapes(Parameters(p): Parameters<DocxPathParam>) -> String {
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
        let mut shapes: Vec<Value> = Vec::new();
        for (para_idx, element) in body.iter().enumerate() {
            if element["type"].as_str() != Some("paragraph") {
                continue;
            }
            let Some(runs) = element["runs"].as_array() else {
                continue;
            };
            for (run_idx, run) in runs.iter().enumerate() {
                if run["type"].as_str() != Some("shape") {
                    continue;
                }
                shapes.push(serde_json::json!({
                    "paragraphIndex": para_idx,
                    "runIndex": run_idx,
                    "presetGeometry": run["presetGeometry"],
                    "widthPt": run["widthPt"],
                    "heightPt": run["heightPt"],
                    "anchorXPt": run["anchorXPt"],
                    "anchorYPt": run["anchorYPt"],
                    "rotation": run["rotation"],
                    "fill": run["fill"],
                    "stroke": run["stroke"],
                    "strokeWidth": run["strokeWidth"],
                    "textBlocks": run["textBlocks"],
                    "wrapMode": run["wrapMode"],
                    "behindDoc": run["behindDoc"],
                    "zOrder": run["zOrder"],
                }));
            }
        }
        serde_json::json!({ "shapes": shapes }).to_string()
    }

    #[tool(
        description = "Return the heading outline of the document. Each entry has the body index, outlineLevel (0-8), styleId, and visible text. Levels come from the parser's resolved `outlineLevel` (style chain + direct pPr) — useful for building TOCs without parsing styleId strings"
    )]
    pub fn docx_get_outline(Parameters(p): Parameters<DocxPathParam>) -> String {
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
        let mut outline: Vec<Value> = Vec::new();
        for (idx, el) in body.iter().enumerate() {
            if el["type"].as_str() != Some("paragraph") {
                continue;
            }
            let Some(level) = el["outlineLevel"].as_u64() else {
                continue;
            };
            let mut text = String::new();
            collect_run_texts(&el["runs"], &mut text);
            outline.push(serde_json::json!({
                "bodyIndex": idx,
                "level": level,
                "styleId": el["styleId"],
                "text": text.trim(),
            }));
        }
        serde_json::json!({ "outline": outline }).to_string()
    }

    #[tool(
        description = "List all `<w:comment>` entries from word/comments.xml: id, author, initials, date, plain text. Empty when the document has no comments part"
    )]
    pub fn docx_get_comments(Parameters(p): Parameters<DocxPathParam>) -> String {
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
        serde_json::json!({ "comments": doc["comments"].as_array().cloned().unwrap_or_default() })
            .to_string()
    }

    #[tool(
        description = "List footnote and endnote bodies from word/footnotes.xml and word/endnotes.xml. Each entry has the id (matches `<w:footnoteReference w:id>` in body) and concatenated plain text"
    )]
    pub fn docx_get_footnotes(Parameters(p): Parameters<DocxPathParam>) -> String {
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
        serde_json::json!({
            "footnotes": doc["footnotes"].as_array().cloned().unwrap_or_default(),
            "endnotes": doc["endnotes"].as_array().cloned().unwrap_or_default(),
        })
        .to_string()
    }

    #[tool(
        description = "Return all track-changes events found in the body: insertions and deletions with author, date, and the text. Empty when the document has no tracked changes"
    )]
    pub fn docx_get_revisions(Parameters(p): Parameters<DocxPathParam>) -> String {
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
        serde_json::json!({ "revisions": doc["revisions"].as_array().cloned().unwrap_or_default() })
            .to_string()
    }
}

#[cfg(test)]
mod sample_tests {
    use super::*;

    fn sample_path() -> String {
        format!(
            "{}/../docx/public/demo/sample-1.docx",
            env!("CARGO_MANIFEST_DIR")
        )
    }

    fn pp(path: &str) -> Parameters<DocxPathParam> {
        Parameters(DocxPathParam { path: path.into() })
    }

    #[test]
    fn docx_to_markdown_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_to_markdown(pp(&path));
        assert!(!out.starts_with("Error:"), "errored: {out}");
        assert!(!out.trim().is_empty(), "markdown should be non-empty");
        let plain = DocxTools::docx_extract_text(pp(&path));
        // Allow markdown to be larger than plain (markup overhead) but not
        // more than 3× — guards against accidentally serializing the rich
        // structure tree.
        assert!(
            out.len() < plain.len() * 3 + 1024,
            "markdown should stay within bounds — got {} vs plain {}",
            out.len(),
            plain.len()
        );
    }

    #[test]
    fn docx_extract_text_sample_non_empty() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_extract_text(pp(&path));
        assert!(!out.starts_with("Error:"), "got error: {out}");
        assert!(!out.trim().is_empty(), "extracted text should be non-empty");
    }

    #[test]
    fn docx_get_sections_sample() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_sections(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        // pageWidth is f64 in pt; should be > 0 for any real document.
        let pw = v["section"]["pageWidth"].as_f64().unwrap_or(0.0);
        assert!(pw > 0.0, "section.pageWidth should be > 0, got {pw}");
    }

    #[test]
    fn docx_get_body_element_first_element() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_body_element(Parameters(DocxIndexParam {
            path: path.clone(),
            index: 0,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert_eq!(v["index"].as_u64(), Some(0));
        assert!(v["type"].is_string(), "missing 'type' on body element");
    }

    #[test]
    fn docx_get_images_returns_array() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_images(Parameters(DocxImagesParam {
            path: path.clone(),
            include_data_url: false,
        }));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(
            v["images"].as_array().is_some(),
            "missing 'images' array: {out}"
        );
    }

    #[test]
    fn docx_invalid_path_returns_error_string() {
        let out = DocxTools::docx_extract_text(pp("/nonexistent/x.docx"));
        assert!(out.starts_with("Error:"), "expected error, got: {out}");
    }

    #[test]
    fn docx_get_outline_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_outline(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["outline"].as_array().is_some(), "missing 'outline'");
    }

    #[test]
    fn docx_get_comments_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_comments(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["comments"].as_array().is_some(), "missing 'comments'");
    }

    #[test]
    fn docx_get_revisions_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_revisions(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["revisions"].as_array().is_some(), "missing 'revisions'");
    }

    #[test]
    fn docx_get_footnotes_smoke() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_footnotes(pp(&path));
        let v: Value = serde_json::from_str(&out).expect("must return JSON");
        assert!(v["footnotes"].as_array().is_some(), "missing 'footnotes'");
        assert!(v["endnotes"].as_array().is_some(), "missing 'endnotes'");
    }

    #[test]
    fn docx_get_body_element_out_of_range_errors() {
        let path = sample_path();
        if !std::path::Path::new(&path).exists() {
            return;
        }
        let out = DocxTools::docx_get_body_element(Parameters(DocxIndexParam {
            path,
            index: 999_999,
        }));
        assert!(
            out.starts_with("Error:"),
            "expected out-of-range error, got: {out}"
        );
    }
}
