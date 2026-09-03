use wasm_bindgen::prelude::*;

/// Markdown-only browser entry point. Keeping this wrapper outside the viewer
/// parser WASM lets applications opt into semantic projection independently.
#[wasm_bindgen]
pub fn to_markdown(data: &[u8]) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    docx_parser::to_markdown_native(data).map_err(|error| JsValue::from_str(&error))
}
