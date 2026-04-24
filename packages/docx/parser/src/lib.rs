use wasm_bindgen::prelude::*;

mod types;
mod xml_util;
mod styles;
mod numbering;
mod parser;

#[wasm_bindgen]
pub fn parse_docx(data: &[u8]) -> String {
    console_error_panic_hook::set_once();
    match parser::parse(data) {
        Ok(doc) => serde_json::to_string(&doc).unwrap_or_else(|e| {
            format!("{{\"error\":\"{}\"}}", e)
        }),
        Err(e) => format!("{{\"error\":\"{}\"}}", e),
    }
}

/// Native equivalent of `parse_docx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_docx_native(data: &[u8]) -> Result<String, String> {
    parser::parse(data)
        .and_then(|doc| serde_json::to_string(&doc).map_err(|e| e.to_string()))
}
