//! The XLSX model-source archive contract over an admitted ODS workbook:
//! the host-layout decision, the workbook bootstrap, a worksheet cursor that
//! lends the projected rows in bounded JSON chunks, and image reads. It
//! mirrors the XLSX worksheet cursor envelope (`preview`, `rows`, `finished`).

use super::{bootstrap, OdsWorkbook};
use crate::package::OdfPackage;
use ooxml_common::resource::{
    HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES, HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
    STANDARD_MAX_ARCHIVE_ENTRY_BYTES,
};
use serde::Serialize;
use std::io::{self, Write};

/// Implementation bound on rows per cursor pull, and the soft byte target
/// of one rows chunk.
const MAX_ROW_CREDIT: usize = 128;
const TARGET_CHUNK_BYTES: usize = 1024 * 1024;

pub struct OdsSession {
    package: Option<OdfPackage>,
    workbook: Option<OdsWorkbook>,
    layout_decided: bool,
    bootstrapped: bool,
    cursor: Option<Cursor>,
    closed: bool,
}

struct Cursor {
    index: usize,
    position: usize,
    preview_sent: bool,
    terminal_awaiting_ack: bool,
    last_pull_terminal: bool,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HostLayoutRequest<'a> {
    family: &'a str,
    size_pt: f64,
    bold: bool,
    italic: bool,
}

#[derive(Serialize)]
struct Finished<'a> {
    kind: &'static str,
    worksheet: &'a xlsx_model::Worksheet,
}

impl OdsSession {
    pub fn new(package: OdfPackage, workbook: OdsWorkbook) -> Self {
        Self {
            package: Some(package),
            workbook: Some(workbook),
            layout_decided: false,
            bootstrapped: false,
            cursor: None,
            closed: false,
        }
    }

    fn workbook(&self) -> Result<&OdsWorkbook, String> {
        if self.closed {
            return Err("ODS session is closed".into());
        }
        self.workbook
            .as_ref()
            .ok_or_else(|| "ODS session is closed".into())
    }

    /// The default cell font, whose digit width encodes column widths; `null`
    /// once decided.
    pub fn host_layout_request(&mut self) -> Result<Vec<u8>, String> {
        let decided = self.layout_decided;
        let workbook = self.workbook()?;
        let (family, size, bold, italic) = &workbook.default_font;
        let request = (!decided).then_some(HostLayoutRequest {
            family,
            size_pt: *size,
            bold: *bold,
            italic: *italic,
        });
        serialize_bounded(
            &request,
            HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
            "host layout request",
        )
    }

    /// The host's single decision: the measured digit width, or `None` (the
    /// §18.3.1.13 fallback width).
    pub fn configure_host_layout(&mut self, mdw: Option<f64>) -> Result<(), String> {
        self.workbook()?;
        if self.layout_decided {
            return if mdw.is_none() {
                Ok(())
            } else {
                self.fail("ODS host layout decision is not pending")
            };
        }
        if mdw.is_some_and(|v| !v.is_finite() || !(1.0..=4096.0).contains(&v)) {
            return self.fail("invalid measured ODS maximum digit width");
        }
        self.workbook.as_mut().expect("checked").apply_layout(mdw);
        self.layout_decided = true;
        Ok(())
    }

    pub fn workbook_bootstrap(&mut self) -> Result<Vec<u8>, String> {
        self.workbook()?;
        if self.bootstrapped {
            return self.fail("ODS bootstrap already consumed");
        }
        if !self.layout_decided {
            self.configure_host_layout(None)?;
        }
        let parsed = bootstrap(self.workbook.as_mut().expect("checked"));
        self.bootstrapped = true;
        let result = serialize_bounded(
            &parsed,
            HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
            "workbook bootstrap",
        );
        self.terminal(result)
    }

    pub fn open_sheet_cursor(&mut self, index: usize, name: &str) -> Result<(), String> {
        let workbook = self.workbook()?;
        if !self.bootstrapped {
            return self.fail("ODS bootstrap must be consumed before worksheets");
        }
        if self.cursor.is_some() {
            return Err("worksheet cursor is already open".into());
        }
        let sheet = workbook
            .sheets
            .get(index)
            .ok_or_else(|| "ODS sheet index is out of range".to_string())?;
        if sheet.name != name {
            return Err("ODS sheet name does not match its index".into());
        }
        self.cursor = Some(Cursor {
            index,
            position: 0,
            preview_sent: false,
            terminal_awaiting_ack: false,
            last_pull_terminal: false,
        });
        Ok(())
    }

    pub fn pull_sheet_cursor(&mut self, row_credit: usize) -> Result<Vec<u8>, String> {
        if row_credit == 0 {
            return Err("worksheet row credit must be greater than zero".into());
        }
        let row_credit = row_credit.min(MAX_ROW_CREDIT);
        let (index, position, preview_sent, awaiting) = {
            let cursor = self
                .cursor
                .as_ref()
                .ok_or_else(|| "worksheet cursor is not open".to_string())?;
            (
                cursor.index,
                cursor.position,
                cursor.preview_sent,
                cursor.terminal_awaiting_ack,
            )
        };
        if awaiting {
            return Err("worksheet terminal product must be acknowledged before pulling".into());
        }
        if !preview_sent {
            // The workbook is materialized whole; the viewer takes the
            // complete-model path rather than a partial preview.
            let bytes = serialize_bounded(
                &serde_json::json!({"kind":"preview", "worksheet":null,
                    "reason":"metadata-unavailable", "maxRow":0, "maxCol":0}),
                HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
                "worksheet preview",
            );
            let bytes = self.terminal(bytes)?;
            self.cursor.as_mut().expect("open").preview_sent = true;
            return Ok(bytes);
        }
        let result = {
            let sheet = &self.workbook()?.sheets[index];
            if position < sheet.rows.len() {
                serialize_rows(&sheet.rows[position..], row_credit)
                    .map(|(bytes, count)| (bytes, position + count, false))
            } else {
                serialize_bounded(
                    &Finished {
                        kind: "finished",
                        worksheet: &sheet.worksheet,
                    },
                    HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
                    "worksheet terminal",
                )
                .map(|bytes| (bytes, position, true))
            }
        };
        let (bytes, next, terminal) = self.terminal(result)?;
        let cursor = self.cursor.as_mut().expect("open");
        cursor.position = next;
        cursor.last_pull_terminal = terminal;
        cursor.terminal_awaiting_ack = terminal;
        Ok(bytes)
    }

    pub fn sheet_cursor_pull_finished(&self) -> bool {
        self.cursor.as_ref().is_some_and(|c| c.last_pull_terminal)
    }

    pub fn acknowledge_sheet_cursor_terminal(&mut self) -> Result<(), String> {
        let cursor = self
            .cursor
            .as_ref()
            .ok_or_else(|| "worksheet cursor is not open".to_string())?;
        if !cursor.terminal_awaiting_ack {
            return Err("worksheet terminal product is not awaiting acknowledgement".into());
        }
        self.cursor = None;
        Ok(())
    }

    pub fn cancel_sheet_cursor(&mut self) {
        self.cursor = None;
    }

    pub fn close_sheet_cursor(&mut self) {
        self.cursor = None;
    }

    /// The bytes of an image a drawing references.
    pub fn extract_image(&mut self, path: &str) -> Result<Vec<u8>, String> {
        let workbook = self.workbook()?;
        if !workbook.images.contains(path) {
            return Err("unadmitted ODS image key".into());
        }
        let package = self
            .package
            .as_ref()
            .ok_or_else(|| "ODS session is closed".to_string())?;
        let bytes = super::draw::read_image_bytes(package, path);
        let bytes = self.terminal(bytes)?;
        if bytes.len() as u64 > STANDARD_MAX_ARCHIVE_ENTRY_BYTES {
            return self.fail("ODS image response byte budget exceeded");
        }
        Ok(bytes)
    }

    pub fn close(&mut self) {
        self.cursor = None;
        self.workbook = None;
        self.package = None;
        self.closed = true;
    }

    pub fn assert_healthy(&self) -> Result<(), String> {
        self.workbook().map(|_| ())
    }

    fn fail<T>(&mut self, message: &str) -> Result<T, String> {
        self.close();
        Err(message.to_string())
    }

    fn terminal<T>(&mut self, result: Result<T, String>) -> Result<T, String> {
        if result.is_err() {
            self.close();
        }
        result
    }
}

fn serialize_bounded<T: Serialize>(
    value: &T,
    ceiling: u64,
    label: &str,
) -> Result<Vec<u8>, String> {
    let ceiling = usize::try_from(ceiling).unwrap_or(usize::MAX);
    let mut output = BoundedJson::new(ceiling);
    serde_json::to_writer(&mut output, value).map_err(|_| {
        if output.exceeded {
            format!("OpenDocument spreadsheet exceeds the {label} JSON limit")
        } else {
            format!("ODS {label} serialization failed")
        }
    })?;
    Ok(output.bytes)
}

fn serialize_rows(rows: &[xlsx_model::Row], credit: usize) -> Result<(Vec<u8>, usize), String> {
    let ceiling = usize::try_from(HARD_MAX_XLSX_WORKSHEET_JSON_BYTES).unwrap_or(usize::MAX);
    let mut output = BoundedJson::new(ceiling);
    let too_large = || "OpenDocument spreadsheet exceeds the worksheet row JSON limit".to_string();
    output
        .write_all(br#"{"kind":"rows","rows":["#)
        .map_err(|_| too_large())?;
    let mut count = 0usize;
    for row in rows.iter().take(credit) {
        let checkpoint = output.bytes.len();
        let written = (|| {
            if count != 0 {
                output.write_all(b",")?;
            }
            serde_json::to_writer(&mut output, row).map_err(io::Error::other)
        })();
        if written.is_err() {
            if output.exceeded && count != 0 {
                output.bytes.truncate(checkpoint);
                output.exceeded = false;
                break;
            }
            return Err(too_large());
        }
        count += 1;
        if output.bytes.len() >= TARGET_CHUNK_BYTES {
            break;
        }
    }
    output.ceiling = output.ceiling.saturating_add(2);
    output.write_all(b"]}").map_err(|_| too_large())?;
    Ok((output.bytes, count))
}

struct BoundedJson {
    bytes: Vec<u8>,
    ceiling: usize,
    exceeded: bool,
}

impl BoundedJson {
    fn new(ceiling: usize) -> Self {
        Self {
            bytes: Vec::new(),
            ceiling: ceiling.saturating_sub(2),
            exceeded: false,
        }
    }
}

impl Write for BoundedJson {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len().saturating_add(bytes.len()) > self.ceiling {
            self.exceeded = true;
            return Err(io::Error::new(io::ErrorKind::StorageFull, "JSON ceiling"));
        }
        self.bytes
            .try_reserve(bytes.len())
            .map_err(|_| io::Error::other("JSON allocation"))?;
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
