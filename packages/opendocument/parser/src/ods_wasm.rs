//! WASM boundary of the ODS reader: the XLSX model-source archive contract
//! (workbook bootstrap, worksheet cursor, image reads and the host-layout
//! capability), backed by an [`OdsSession`].

use crate::ods::{self, wire::OdsSession};
use crate::package::{OdfPackage, PackageLimits};
use ooxml_common::resource::OoxmlFormat;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct OdsWorkbook {
    session: OdsSession,
}

#[wasm_bindgen]
impl OdsWorkbook {
    /// Open and fully read the package. Every rejection happens here. Limits
    /// are the caller's normalized ZIP policy in the shared governor's
    /// encoding (`undefined`: the standard limit; `0`: removed, hard
    /// ceilings still apply).
    #[wasm_bindgen(constructor)]
    pub fn new(
        data: js_sys::Uint8Array,
        max_archive_entry_bytes: Option<f64>,
        max_total_inflated_bytes: Option<f64>,
        max_archive_entries: Option<f64>,
    ) -> Result<OdsWorkbook, JsValue> {
        console_error_panic_hook::set_once();
        let limits = PackageLimits {
            max_archive_entry_bytes: limit(max_archive_entry_bytes)?,
            max_total_inflated_bytes: limit(max_total_inflated_bytes)?,
            max_archive_entries: limit(max_archive_entries)?,
        };
        let package = OdfPackage::open(data.to_vec(), OoxmlFormat::Xlsx, limits, &ods::MEDIA_TYPES)
            .map_err(error)?;
        let workbook = ods::read(&package).map_err(error)?;
        Ok(Self {
            session: OdsSession::new(package, workbook),
        })
    }

    pub fn host_layout_request(&mut self) -> Result<Vec<u8>, JsValue> {
        self.session.host_layout_request().map_err(error)
    }

    pub fn configure_host_layout(
        &mut self,
        maximum_digit_width: Option<f64>,
    ) -> Result<(), JsValue> {
        self.session
            .configure_host_layout(maximum_digit_width)
            .map_err(error)
    }

    pub fn parse(&mut self) -> Result<Vec<u8>, JsValue> {
        self.session.workbook_bootstrap().map_err(error)
    }

    pub fn open_sheet_cursor(&mut self, sheet_index: u32, name: &str) -> Result<(), JsValue> {
        self.session
            .open_sheet_cursor(sheet_index as usize, name)
            .map_err(error)
    }

    pub fn pull_sheet_cursor(&mut self, row_credit: u32) -> Result<Vec<u8>, JsValue> {
        self.session
            .pull_sheet_cursor(row_credit as usize)
            .map_err(error)
    }

    pub fn sheet_cursor_pull_finished(&self) -> bool {
        self.session.sheet_cursor_pull_finished()
    }

    pub fn acknowledge_sheet_cursor_terminal(&mut self) -> Result<(), JsValue> {
        self.session
            .acknowledge_sheet_cursor_terminal()
            .map_err(error)
    }

    pub fn cancel_sheet_cursor(&mut self) {
        self.session.cancel_sheet_cursor();
    }

    pub fn close_sheet_cursor(&mut self) {
        self.session.close_sheet_cursor();
    }

    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, JsValue> {
        self.session.extract_image(path).map_err(error)
    }

    pub fn close_workbook_session(&mut self) {
        self.session.close();
    }

    pub fn assert_healthy(&self) -> Result<(), JsValue> {
        self.session.assert_healthy().map_err(error)
    }
}

fn limit(value: Option<f64>) -> Result<Option<u64>, JsValue> {
    match value {
        None => Ok(None),
        Some(v)
            if v.is_finite() && v >= 0.0 && v.fract() == 0.0 && v <= 9_007_199_254_740_991.0 =>
        {
            Ok(Some(v as u64))
        }
        Some(_) => Err(JsValue::from_str("OpenDocument resource limit is invalid")),
    }
}

fn error(message: String) -> JsValue {
    JsValue::from_str(&message)
}
