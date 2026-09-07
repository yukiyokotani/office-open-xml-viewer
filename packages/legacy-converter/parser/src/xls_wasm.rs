//! Experimental direct `.xls` workbook-source boundary.
//!
//! This API projects the supported passive BIFF8 subset directly into the
//! existing XLSX renderer wire models. It does not create an OOXML package or
//! invoke the XLSX parser. The caller retains its original JavaScript bytes;
//! this constructor bounds and releases its temporary CFB copy after the owned
//! direct session has retained the admitted neutral data.

use crate::cfb::CompoundFile;
use crate::xls::{direct::DirectSession, direct_wire::DirectWire};
use wasm_bindgen::prelude::*;

/// Implementation resource policy, not an MS-XLS format limit.
const MAX_DIRECT_XLS_SOURCE_BYTES: usize = 256 * 1024 * 1024;

#[wasm_bindgen]
pub struct LegacyXlsWorkbook {
    wire: DirectWire,
}

#[wasm_bindgen]
impl LegacyXlsWorkbook {
    #[wasm_bindgen(constructor)]
    pub fn new(data: js_sys::Uint8Array) -> Result<LegacyXlsWorkbook, JsValue> {
        console_error_panic_hook::set_once();
        let source_len = data.length() as usize;
        if source_len > MAX_DIRECT_XLS_SOURCE_BYTES {
            return Err(js_error(
                "UNSUPPORTED:XLS direct source byte budget exceeded",
            ));
        }
        let mut source = Vec::new();
        source
            .try_reserve_exact(source_len)
            .map_err(|_| js_error("UNSUPPORTED:XLS direct source allocation failed"))?;
        source.resize(source_len, 0);
        data.copy_to(&mut source);
        let cfb = CompoundFile::open(&source)
            .map_err(|error| js_error(&format!("UNSUPPORTED:{error}")))?;
        let session = DirectSession::new(&cfb).map_err(string_error)?;
        Ok(Self {
            wire: DirectWire::new(session),
        })
    }

    pub fn measurement_request(&mut self) -> Result<Vec<u8>, JsValue> {
        self.wire.measurement_request().map_err(string_error)
    }

    pub fn configure_mdw(&mut self, maximum_digit_width: Option<f64>) -> Result<(), JsValue> {
        self.wire
            .configure_mdw(maximum_digit_width)
            .map_err(string_error)
    }

    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        self.wire.workbook_bootstrap().map_err(string_error)
    }

    pub fn open_sheet_cursor(&mut self, sheet_index: u32, name: &str) -> Result<(), JsValue> {
        self.wire
            .open_sheet_cursor(sheet_index as usize, name)
            .map_err(string_error)
    }

    pub fn pull_sheet_cursor(&mut self, row_credit: u32) -> Result<Vec<u8>, JsValue> {
        self.wire
            .pull_sheet_cursor(row_credit as usize)
            .map_err(string_error)
    }

    pub fn sheet_cursor_pull_finished(&self) -> bool {
        self.wire.sheet_cursor_pull_finished()
    }

    pub fn acknowledge_sheet_cursor_terminal(&mut self) -> Result<(), JsValue> {
        self.wire
            .acknowledge_sheet_cursor_terminal()
            .map_err(string_error)
    }

    pub fn cancel_sheet_cursor(&mut self) {
        self.wire.cancel_sheet_cursor();
    }

    pub fn close_sheet_cursor(&mut self) {
        self.wire.close_sheet_cursor();
    }

    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        self.wire.extract_image(path).map_err(string_error)
    }

    pub fn close_workbook_session(&mut self) -> Result<(), JsValue> {
        self.wire.close().map_err(string_error)
    }

    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.wire.assert_healthy().map_err(string_error)
    }

    pub fn resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        self.wire.assert_healthy().map_err(string_error)?;
        Err(js_error("xlsx resource usage is unavailable"))
    }

    pub fn sheet_cursor_resource_usage(&self) -> Result<Vec<u8>, JsValue> {
        self.wire.assert_healthy().map_err(string_error)?;
        Err(js_error("worksheet cursor usage is unavailable"))
    }

    pub fn to_markdown(&self) -> Result<String, JsValue> {
        self.wire.assert_healthy().map_err(string_error)?;
        Err(js_error(
            "UNSUPPORTED:legacy XLS direct markdown projection is not supported",
        ))
    }
}

fn string_error(error: String) -> JsValue {
    js_error(&error)
}

fn js_error(error: &str) -> JsValue {
    JsValue::from_str(error)
}
