//! Bounded JSON wire boundary for the owning native XLS model session.
//! It mirrors the XLSX worksheet cursor envelope without pretending BIFF decode
//! is streaming: projection is cached once, then rows are borrowed in chunks.

use super::direct::DirectSession;
use ooxml_common::resource::{
    HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES, HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
    STANDARD_MAX_ARCHIVE_ENTRY_BYTES,
};
use serde::Serialize;
use std::io::{self, Write};

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct MeasurementRequest<'a> {
    required: bool,
    font: Option<MeasurementFont<'a>>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct MeasurementFont<'a> {
    name: &'a str,
    size_points: f64,
    bold: bool,
    italic: bool,
}

pub(crate) struct DirectWire {
    session: Option<DirectSession>,
    cursor: Option<ActiveSheet>,
    closed: bool,
}

struct ActiveSheet {
    index: usize,
    name: String,
    row_position: usize,
    terminal_awaiting_ack: bool,
    last_pull_terminal: bool,
}

#[derive(Serialize)]
struct Finished<'a> {
    kind: &'static str,
    worksheet: &'a xlsx_model::Worksheet,
}

impl DirectWire {
    pub(crate) fn new(session: DirectSession) -> Self {
        Self {
            session: Some(session),
            cursor: None,
            closed: false,
        }
    }

    pub(crate) fn measurement_request(&mut self) -> Result<Vec<u8>, String> {
        let result = {
            let session = self.session()?;
            let request = MeasurementRequest {
                required: session.requires_measurement_decision(),
                font: session.measurement_font().map(|font| MeasurementFont {
                    name: &font.name,
                    size_points: font.size_points,
                    bold: font.bold,
                    italic: font.italic,
                }),
            };
            serialize_bounded(
                &request,
                HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
                "measurement request",
            )
        };
        self.terminal(result)
    }

    pub(crate) fn configure_mdw(&mut self, mdw: Option<f64>) -> Result<(), String> {
        self.with_session(|session| session.configure_mdw(mdw))
    }

    pub(crate) fn workbook_bootstrap(&mut self) -> Result<Vec<u8>, String> {
        let value = self.with_session(DirectSession::bootstrap)?;
        self.serialize(
            &value,
            HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
            "workbook bootstrap",
        )
    }

    pub(crate) fn open_sheet_cursor(&mut self, index: usize, name: &str) -> Result<(), String> {
        self.session()?;
        if self.cursor.is_some() {
            return Err("worksheet cursor is already open".into());
        }
        let result = self
            .session
            .as_mut()
            .expect("open session checked above")
            .projected_sheet(index, name)
            .map(|_| ());
        if let Err(error) = result {
            // Index/name mistakes are recoverable; projection failures poison
            // the owning session and therefore close the document wire.
            if self
                .session
                .as_ref()
                .expect("session retained after recoverable validation")
                .assert_healthy()
                .is_err()
            {
                let _ = self.close();
            }
            return Err(error);
        }
        let mut owned_name = String::new();
        let allocation = owned_name
            .try_reserve_exact(name.len())
            .map_err(|_| "UNSUPPORTED:XLS direct cursor name allocation failed".to_string());
        self.terminal(allocation)?;
        owned_name.push_str(name);
        self.cursor = Some(ActiveSheet {
            index,
            name: owned_name,
            row_position: 0,
            terminal_awaiting_ack: false,
            last_pull_terminal: false,
        });
        Ok(())
    }

    pub(crate) fn pull_sheet_cursor(&mut self, row_credit: usize) -> Result<Vec<u8>, String> {
        self.pull_sheet_cursor_with_limits(
            row_credit,
            HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
            1024 * 1024,
        )
    }

    fn pull_sheet_cursor_with_limits(
        &mut self,
        row_credit: usize,
        json_ceiling: u64,
        target_bytes: usize,
    ) -> Result<Vec<u8>, String> {
        if row_credit == 0 {
            return Err("worksheet row credit must be greater than zero".into());
        }
        let row_credit = row_credit.min(128);
        let cursor = self
            .cursor
            .as_ref()
            .ok_or_else(|| "worksheet cursor is not open".to_string())?;
        if cursor.terminal_awaiting_ack {
            return Err("worksheet terminal product must be acknowledged before pulling".into());
        }
        let result = (|| {
            let projected = self
                .session
                .as_mut()
                .ok_or_else(|| "XLS direct wire is closed".to_string())?
                .projected_sheet(cursor.index, &cursor.name)?;
            let position = cursor.row_position;
            if position < projected.rows.len() {
                serialize_rows(
                    &projected.rows[position..],
                    row_credit,
                    json_ceiling,
                    target_bytes,
                )
                .map(|(bytes, count)| (bytes, position + count, false))
            } else {
                serialize_bounded(
                    &Finished {
                        kind: "finished",
                        worksheet: projected.worksheet,
                    },
                    json_ceiling,
                    "worksheet terminal",
                )
                .map(|bytes| (bytes, position, true))
            }
        })();
        let (bytes, next_position, terminal) = self.terminal(result)?;
        let cursor = self.cursor.as_mut().expect("cursor retained on success");
        cursor.row_position = next_position;
        cursor.last_pull_terminal = terminal;
        cursor.terminal_awaiting_ack = terminal;
        Ok(bytes)
    }

    pub(crate) fn sheet_cursor_pull_finished(&self) -> bool {
        self.cursor
            .as_ref()
            .is_some_and(|cursor| cursor.last_pull_terminal)
    }

    pub(crate) fn acknowledge_sheet_cursor_terminal(&mut self) -> Result<(), String> {
        let cursor = self
            .cursor
            .as_ref()
            .ok_or_else(|| "worksheet cursor is not open".to_string())?;
        if !cursor.terminal_awaiting_ack {
            return Err("worksheet terminal product is not awaiting acknowledgement".into());
        }
        self.cursor.take();
        Ok(())
    }

    pub(crate) fn cancel_sheet_cursor(&mut self) {
        self.cursor.take();
    }

    pub(crate) fn close_sheet_cursor(&mut self) {
        self.cursor.take();
    }

    pub(crate) fn extract_image(&mut self, key: &str) -> Result<Vec<u8>, String> {
        let result = {
            let bytes = self.session()?.resource(key)?;
            if bytes.len() as u64 > STANDARD_MAX_ARCHIVE_ENTRY_BYTES {
                Err("UNSUPPORTED:XLS direct image response byte budget exceeded".into())
            } else {
                let mut owned = Vec::new();
                owned
                    .try_reserve_exact(bytes.len())
                    .map_err(|_| {
                        "UNSUPPORTED:XLS direct image response allocation failed".to_string()
                    })
                    .map(|_| {
                        owned.extend_from_slice(bytes);
                        owned
                    })
            }
        };
        self.terminal(result)
    }

    pub(crate) fn close(&mut self) -> Result<(), String> {
        self.cursor.take();
        self.session.take();
        self.closed = true;
        Ok(())
    }

    pub(crate) fn assert_healthy(&self) -> Result<(), String> {
        self.session().map(|_| ())
    }

    fn serialize<T: Serialize>(
        &mut self,
        value: &T,
        ceiling: u64,
        label: &str,
    ) -> Result<Vec<u8>, String> {
        let result = serialize_bounded(value, ceiling, label);
        self.terminal(result)
    }

    fn with_session<T>(
        &mut self,
        operation: impl FnOnce(&mut DirectSession) -> Result<T, String>,
    ) -> Result<T, String> {
        let result = match self.session.as_mut() {
            Some(session) if !self.closed => operation(session),
            _ => return Err("XLS direct wire is closed".into()),
        };
        self.terminal(result)
    }

    fn terminal<T>(&mut self, result: Result<T, String>) -> Result<T, String> {
        if result.is_err() {
            let _ = self.close();
        }
        result
    }

    fn session(&self) -> Result<&DirectSession, String> {
        if self.closed {
            return Err("XLS direct wire is closed".into());
        }
        self.session
            .as_ref()
            .ok_or_else(|| "XLS direct wire has no session".into())
    }
}

fn serialize_bounded<T: Serialize>(
    value: &T,
    ceiling: u64,
    label: &str,
) -> Result<Vec<u8>, String> {
    let ceiling = usize::try_from(ceiling)
        .map_err(|_| format!("UNSUPPORTED:XLS direct {label} JSON size exceeds this platform"))?;
    let mut output = BoundedJson::new(ceiling);
    if let Err(error) = serde_json::to_writer(&mut output, value) {
        return match output.failure {
            Some(WriteFailure::Ceiling) => Err(format!(
                "UNSUPPORTED:XLS direct {label} JSON byte budget exceeded"
            )),
            Some(WriteFailure::Allocation) => Err(format!(
                "UNSUPPORTED:XLS direct {label} JSON allocation failed"
            )),
            None => Err(format!("serialize error: {error}")),
        };
    }
    Ok(output.bytes)
}

fn serialize_rows(
    rows: &[xlsx_model::Row],
    row_credit: usize,
    ceiling: u64,
    target_bytes: usize,
) -> Result<(Vec<u8>, usize), String> {
    let ceiling = usize::try_from(ceiling).map_err(|_| {
        "UNSUPPORTED:XLS direct worksheet rows JSON size exceeds this platform".to_string()
    })?;
    let mut output = BoundedJson::new(ceiling);
    output
        .write_all(br#"{"kind":"rows","rows":["#)
        .map_err(|_| "UNSUPPORTED:XLS direct worksheet rows JSON byte budget exceeded")?;
    output.ceiling = ceiling.checked_sub(2).ok_or_else(|| {
        "UNSUPPORTED:XLS direct worksheet rows JSON byte budget exceeded".to_string()
    })?;
    let mut count = 0usize;
    for row in rows.iter().take(row_credit) {
        let checkpoint = output.bytes.len();
        let write_result = (|| {
            if count != 0 {
                output.write_all(b",")?;
            }
            serde_json::to_writer(&mut output, row).map_err(io::Error::other)
        })();
        if let Err(error) = write_result {
            if matches!(output.failure, Some(WriteFailure::Ceiling)) && count != 0 {
                output.bytes.truncate(checkpoint);
                output.failure = None;
                break;
            }
            return match output.failure {
                Some(WriteFailure::Ceiling) => {
                    Err("UNSUPPORTED:XLS direct worksheet row JSON byte budget exceeded".into())
                }
                Some(WriteFailure::Allocation) => {
                    Err("UNSUPPORTED:XLS direct worksheet rows JSON allocation failed".into())
                }
                None => Err(format!("serialize error: {error}")),
            };
        }
        count += 1;
        if output.bytes.len() >= target_bytes {
            break;
        }
    }
    debug_assert!(count > 0, "caller supplies at least one remaining row");
    output.ceiling = ceiling;
    output
        .write_all(b"]}")
        .map_err(|_| "UNSUPPORTED:XLS direct worksheet rows JSON byte budget exceeded")?;
    Ok((output.bytes, count))
}

#[derive(Clone, Copy)]
enum WriteFailure {
    Ceiling,
    Allocation,
}

struct BoundedJson {
    bytes: Vec<u8>,
    ceiling: usize,
    failure: Option<WriteFailure>,
    #[cfg(test)]
    growths: usize,
    #[cfg(test)]
    maximum_requested_capacity: usize,
}

impl BoundedJson {
    fn new(ceiling: usize) -> Self {
        Self {
            bytes: Vec::new(),
            ceiling,
            failure: None,
            #[cfg(test)]
            growths: 0,
            #[cfg(test)]
            maximum_requested_capacity: 0,
        }
    }
}

impl Write for BoundedJson {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let Some(next_len) = self.bytes.len().checked_add(bytes.len()) else {
            self.failure = Some(WriteFailure::Ceiling);
            return Err(io::Error::new(io::ErrorKind::StorageFull, "JSON ceiling"));
        };
        if next_len > self.ceiling {
            self.failure = Some(WriteFailure::Ceiling);
            return Err(io::Error::new(io::ErrorKind::StorageFull, "JSON ceiling"));
        }
        if next_len > self.bytes.capacity() {
            let doubled = self.bytes.capacity().checked_mul(2).unwrap_or(self.ceiling);
            let target = next_len.max(doubled.max(8)).min(self.ceiling);
            self.bytes
                .try_reserve_exact(target - self.bytes.len())
                .map_err(|_| {
                    self.failure = Some(WriteFailure::Allocation);
                    io::Error::other("JSON allocation")
                })?;
            #[cfg(test)]
            {
                self.growths += 1;
                self.maximum_requested_capacity = self.maximum_requested_capacity.max(target);
            }
            debug_assert!(self.bytes.capacity() >= next_len);
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn row(index: u32, text: &str) -> xlsx_model::Row {
        xlsx_model::Row {
            index,
            height: None,
            custom_height: false,
            cells: vec![xlsx_model::Cell {
                col: 1,
                row: index,
                value: xlsx_model::CellValue::Text {
                    text: text.into(),
                    runs: None,
                    phonetic_runs: Vec::new(),
                    phonetic_pr: None,
                },
                style_index: Some(0),
                formula: None,
                show_phonetic: false,
            }],
            outline_level: 0,
            collapsed: false,
            hidden: false,
        }
    }

    #[test]
    fn native_wire_cursor_chunks_rows_and_acknowledges_terminal() {
        let mut wire = DirectWire::new(super::super::direct::tests::wire_fixture());
        let request: serde_json::Value =
            serde_json::from_slice(&wire.measurement_request().unwrap()).unwrap();
        assert_eq!(request["required"], false);
        assert!(request["font"].is_null());
        let bootstrap: serde_json::Value =
            serde_json::from_slice(&wire.workbook_bootstrap().unwrap()).unwrap();
        assert_eq!(bootstrap["workbook"]["sheets"][0]["name"], "S");
        wire.open_sheet_cursor(0, "S").unwrap();
        assert!(!wire.sheet_cursor_pull_finished());
        assert!(wire.acknowledge_sheet_cursor_terminal().is_err());
        let rows: serde_json::Value =
            serde_json::from_slice(&wire.pull_sheet_cursor(1).unwrap()).unwrap();
        assert_eq!(rows["kind"], "rows");
        assert_eq!(rows["rows"].as_array().unwrap().len(), 1);
        assert!(!wire.sheet_cursor_pull_finished());
        let terminal: serde_json::Value =
            serde_json::from_slice(&wire.pull_sheet_cursor(1).unwrap()).unwrap();
        assert_eq!(terminal["kind"], "finished");
        assert_eq!(terminal["worksheet"]["name"], "S");
        assert_eq!(terminal["worksheet"]["rows"], serde_json::json!([]));
        assert!(wire.sheet_cursor_pull_finished());
        assert!(wire.pull_sheet_cursor(1).is_err());
        wire.acknowledge_sheet_cursor_terminal().unwrap();

        // ACK retains the projected slot, so reopening does not rebuild it.
        wire.open_sheet_cursor(0, "S").unwrap();
        assert_eq!(
            serde_json::from_slice::<serde_json::Value>(&wire.pull_sheet_cursor(128).unwrap())
                .unwrap()["kind"],
            "rows"
        );
        wire.cancel_sheet_cursor();
        wire.open_sheet_cursor(0, "S").unwrap();
        wire.close_sheet_cursor();
        assert!(wire.extract_image("legacy-xls/image/0").is_err());
        assert!(wire.assert_healthy().is_ok());
        wire.close().unwrap();
        assert!(wire.assert_healthy().is_err());
        assert!(wire.close().is_ok());
        assert!(wire.close().is_ok());
    }

    #[test]
    fn measurement_request_and_decision_use_the_inherited_session_state() {
        let mut wire = DirectWire::new(super::super::direct::tests::wire_picture_fixture());
        let request: serde_json::Value =
            serde_json::from_slice(&wire.measurement_request().unwrap()).unwrap();
        assert_eq!(request["required"], true);
        assert_eq!(request["font"]["name"], "Calibri");
        assert_eq!(request["font"]["sizePoints"], 11.0);
        wire.configure_mdw(Some(7.0)).unwrap();
        assert!(wire.workbook_bootstrap().is_ok());
        assert!(wire.configure_mdw(Some(8.0)).is_err());
        assert!(wire.assert_healthy().is_err());
    }

    #[test]
    fn serialization_ceiling_is_checked_before_output_allocation() {
        assert_eq!(serialize_bounded(&"a\nb", 6, "test").unwrap(), br#""a\nb""#);
        let error = serialize_bounded(&"a\nb", 5, "test").unwrap_err();
        assert!(error.contains("JSON byte budget exceeded"));
    }

    #[test]
    fn many_small_writes_grow_capacity_logarithmically_without_crossing_cap() {
        let mut output = BoundedJson::new(4096);
        for _ in 0..4096 {
            output.write_all(b"x").unwrap();
        }
        assert_eq!(output.bytes.len(), 4096);
        assert!(output.growths <= 10, "growths={}", output.growths);
        assert!(output.maximum_requested_capacity <= 4096);
        assert!(output.bytes.capacity() >= output.bytes.len());
        assert!(output.write_all(b"x").is_err());
        assert_eq!(output.bytes.len(), 4096);
    }

    #[test]
    fn growth_reserve_is_relative_to_length_not_existing_capacity() {
        let mut output = BoundedJson::new(23);
        output.write_all(b"abc").unwrap();
        assert!(output.bytes.capacity() >= 8);
        output.write_all(&[b'x'; 20]).unwrap();
        assert_eq!(output.bytes.len(), 23);
        assert!(output.bytes.capacity() >= 23);
        assert_eq!(output.maximum_requested_capacity, 23);
    }

    #[test]
    fn row_envelopes_split_at_soft_target_without_omission() {
        let rows = vec![row(1, "small"), row(2, &"x".repeat(300)), row(3, "tail")];
        let mut position = 0;
        let mut seen = Vec::new();
        while position < rows.len() {
            let (bytes, count) = serialize_rows(&rows[position..], 128, 1024, 180).unwrap();
            let value: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
            seen.extend(
                value["rows"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|row| row["index"].as_u64().unwrap()),
            );
            assert!(count > 0);
            position += count;
        }
        assert_eq!(seen, [1, 2, 3]);

        let error = serialize_rows(&[row(1, &"x".repeat(300))], 1, 64, 32).unwrap_err();
        assert!(error.contains("row JSON byte budget exceeded"));

        let many: Vec<_> = (1..=129).map(|index| row(index, "")).collect();
        let (_, count) = serialize_rows(&many, 128, 1024 * 1024, usize::MAX).unwrap();
        assert_eq!(count, 128);
    }

    #[test]
    fn inherited_state_error_closes_the_wire() {
        let mut wire = DirectWire::new(super::super::direct::tests::wire_fixture());
        assert!(wire.open_sheet_cursor(0, "S").is_err());
        assert!(wire.assert_healthy().is_err());
    }

    #[test]
    fn cursor_input_errors_are_recoverable_and_do_not_advance_rows() {
        let mut wire = DirectWire::new(super::super::direct::tests::wire_fixture());
        wire.workbook_bootstrap().unwrap();
        assert!(wire.open_sheet_cursor(9, "S").is_err());
        assert!(wire.open_sheet_cursor(0, "wrong").is_err());
        wire.open_sheet_cursor(0, "S").unwrap();
        assert!(wire.open_sheet_cursor(0, "S").is_err());
        assert!(wire.pull_sheet_cursor(0).is_err());
        let first: serde_json::Value =
            serde_json::from_slice(&wire.pull_sheet_cursor(1).unwrap()).unwrap();
        assert_eq!(first["rows"][0]["index"], 1);
        wire.cancel_sheet_cursor();
        assert!(wire.assert_healthy().is_ok());
    }

    #[test]
    fn row_serialization_failure_is_document_terminal() {
        let mut wire = DirectWire::new(super::super::direct::tests::wire_fixture());
        wire.workbook_bootstrap().unwrap();
        wire.open_sheet_cursor(0, "S").unwrap();
        assert!(wire.pull_sheet_cursor_with_limits(1, 1, 1).is_err());
        assert!(wire.assert_healthy().is_err());
    }
}
