use rmcp::{handler::server::wrapper::Parameters, tool};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::Value;
use std::fs;

// ─── Parameter types ─────────────────────────────────────────────────────────

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxPathParam {
    /// Absolute path to the PPTX file
    pub path: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxSlideParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index
    pub slide_index: usize,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct PptxTextParam {
    /// Absolute path to the PPTX file
    pub path: String,
    /// 0-based slide index; omit to extract text from all slides
    pub slide_index: Option<usize>,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn read_file(path: &str) -> Result<Vec<u8>, String> {
    fs::read(path).map_err(|e| format!("Cannot read '{}': {}", path, e))
}

fn extract_text_runs(node: &Value, out: &mut String) {
    match node["type"].as_str().unwrap_or("") {
        "textRun" | "run" => {
            if let Some(t) = node["text"].as_str() {
                out.push_str(t);
            }
        }
        _ => {}
    }
    // Recurse into common container fields
    for key in &["runs", "paragraphs", "elements", "children"] {
        if let Some(arr) = node[key].as_array() {
            for child in arr {
                extract_text_runs(child, out);
            }
        }
    }
}

fn slide_title(slide: &Value) -> Option<String> {
    if let Some(elements) = slide["elements"].as_array() {
        for el in elements {
            if el["type"].as_str() == Some("shape") {
                if let Some(ph) = el["placeholderType"].as_str() {
                    if ph == "title" || ph == "centeredTitle" {
                        let mut text = String::new();
                        if let Some(tb) = el.get("textBody") {
                            if let Some(paras) = tb["paragraphs"].as_array() {
                                for para in paras {
                                    extract_text_runs(para, &mut text);
                                }
                            }
                        }
                        let trimmed = text.trim().to_string();
                        if !trimmed.is_empty() {
                            return Some(trimmed);
                        }
                    }
                }
            }
        }
    }
    None
}

fn extract_slide_text(slide: &Value) -> String {
    let mut out = String::new();
    if let Some(elements) = slide["elements"].as_array() {
        for el in elements {
            if let Some(tb) = el.get("textBody") {
                if let Some(paras) = tb["paragraphs"].as_array() {
                    for para in paras {
                        extract_text_runs(para, &mut out);
                        out.push('\n');
                    }
                }
            }
            // Table elements
            if el["type"].as_str() == Some("table") {
                if let Some(rows) = el["rows"].as_array() {
                    for row in rows {
                        if let Some(cells) = row["cells"].as_array() {
                            let cell_texts: Vec<String> = cells
                                .iter()
                                .map(|c| {
                                    let mut t = String::new();
                                    if let Some(paras) = c["paragraphs"].as_array() {
                                        for para in paras {
                                            extract_text_runs(para, &mut t);
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
        }
    }
    out
}

fn slide_structure(slide: &Value) -> Value {
    let elements: Vec<Value> = slide["elements"]
        .as_array()
        .map(|els| {
            els.iter()
                .map(|el| {
                    let mut text = String::new();
                    if let Some(tb) = el.get("textBody") {
                        if let Some(paras) = tb["paragraphs"].as_array() {
                            for para in paras {
                                extract_text_runs(para, &mut text);
                            }
                        }
                    }
                    serde_json::json!({
                        "type": el["type"],
                        "placeholderType": el["placeholderType"],
                        "x": el["x"], "y": el["y"],
                        "width": el["width"], "height": el["height"],
                        "text": text.trim().to_string(),
                    })
                })
                .collect()
        })
        .unwrap_or_default();

    serde_json::json!({
        "index": slide["index"],
        "slideNumber": slide["slideNumber"],
        "elements": elements,
    })
}

// ─── Tool implementations ─────────────────────────────────────────────────────

pub struct PptxTools;

impl PptxTools {
    #[tool(description = "Return the number of slides and each slide's title from a PPTX file")]
    pub fn pptx_get_slides(Parameters(p): Parameters<PptxPathParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"].as_array().map(|s| s.as_slice()).unwrap_or(&[]);
        let summary: Vec<Value> = slides
            .iter()
            .map(|s| {
                serde_json::json!({
                    "index": s["index"],
                    "slideNumber": s["slideNumber"],
                    "title": slide_title(s),
                })
            })
            .collect();
        serde_json::json!({
            "slideCount": slides.len(),
            "slides": summary,
        })
        .to_string()
    }

    #[tool(description = "Extract plain text from a PPTX file; optionally filter to a single slide by 0-based index")]
    pub fn pptx_extract_text(Parameters(p): Parameters<PptxTextParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"].as_array().map(|s| s.as_slice()).unwrap_or(&[]);

        if let Some(idx) = p.slide_index {
            let slide = match slides.get(idx) {
                Some(s) => s,
                None => {
                    return format!(
                        "Error: slide index {} out of range (total: {})",
                        idx,
                        slides.len()
                    )
                }
            };
            return extract_slide_text(slide);
        }

        let mut out = String::new();
        for (i, slide) in slides.iter().enumerate() {
            out.push_str(&format!("=== Slide {} ===\n", i + 1));
            out.push_str(&extract_slide_text(slide));
            out.push('\n');
        }
        out
    }

    #[tool(description = "Return the structure (elements with position, size, text) of a single slide")]
    pub fn pptx_get_slide_structure(Parameters(p): Parameters<PptxSlideParam>) -> String {
        let data = match read_file(&p.path) {
            Ok(d) => d,
            Err(e) => return format!("Error: {}", e),
        };
        let pres_json = match pptx_parser::parse_pptx_native(&data) {
            Ok(j) => j,
            Err(e) => return format!("Error: {}", e),
        };
        let pres: Value = match serde_json::from_str(&pres_json) {
            Ok(v) => v,
            Err(e) => return format!("Error: {}", e),
        };
        let slides = pres["slides"].as_array().map(|s| s.as_slice()).unwrap_or(&[]);
        let slide = match slides.get(p.slide_index) {
            Some(s) => s,
            None => {
                return format!(
                    "Error: slide index {} out of range (total: {})",
                    p.slide_index,
                    slides.len()
                )
            }
        };
        serde_json::to_string(&slide_structure(slide)).unwrap_or_else(|e| format!("Error: {}", e))
    }
}
