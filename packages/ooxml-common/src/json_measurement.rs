//! Measurement and bounded serialization of JSON resources.
//!
//! The JSON byte count is the exact output size produced by `serde_json`. The
//! string-value count measures decoded UTF-8 content and deliberately excludes
//! object property names. This lets format parsers account for retained string
//! content separately from the serialized model without allocating that model.

use serde::Serialize;
use std::io::{self, Write};

use crate::package_session::PackageLimitReporter;
use crate::resource::HardResourceLimitKind;

/// A serialization failure distinct from crossing a parser's JSON byte ceiling.
#[derive(Debug)]
enum LimitedJsonError {
    LimitExceeded { observed: u64 },
    Serialize(String),
}

#[derive(Debug)]
struct JsonLimitExceeded {
    observed: u64,
    limit: u64,
}

impl std::fmt::Display for JsonLimitExceeded {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "JSON byte limit exceeded: {} > {}",
            self.observed, self.limit
        )
    }
}

impl std::error::Error for JsonLimitExceeded {}

struct LimitedJsonWriter {
    bytes: Vec<u8>,
    limit: u64,
    exceeded: Option<u64>,
}

impl Write for LimitedJsonWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let observed = (self.bytes.len() as u64).saturating_add(bytes.len() as u64);
        if observed > self.limit {
            self.exceeded = Some(observed);
            return Err(io::Error::other(JsonLimitExceeded {
                observed,
                limit: self.limit,
            }));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Serialize directly into a byte-capped buffer, aborting on the first crossing.
/// The reported observation is the first rejected write's cumulative size; it
/// can be less than the complete JSON size, while retaining the exact ceiling.
fn serialize_json_limited<T: Serialize>(
    value: &T,
    limit: u64,
) -> Result<Vec<u8>, LimitedJsonError> {
    let mut writer = LimitedJsonWriter {
        bytes: Vec::new(),
        limit,
        exceeded: None,
    };
    if let Err(error) = serde_json::to_writer(&mut writer, value) {
        if let Some(observed) = writer.exceeded {
            return Err(LimitedJsonError::LimitExceeded { observed });
        }
        return Err(LimitedJsonError::Serialize(error.to_string()));
    }
    Ok(writer.bytes)
}

/// Serialize one emitted JSON unit and observe its hard ceiling from the same
/// pass. The prefix is used only when no reporter rejects a limit crossing.
pub fn serialize_json_with_limit<T: Serialize>(
    value: &T,
    reporter: Option<&PackageLimitReporter>,
    kind: HardResourceLimitKind,
    part: Option<&str>,
    limit: u64,
    limit_error_prefix: &str,
) -> Result<Vec<u8>, String> {
    let bytes = match serialize_json_limited(value, limit) {
        Ok(bytes) => bytes,
        Err(LimitedJsonError::LimitExceeded { observed }) => {
            if let Some(reporter) = reporter {
                reporter.observe_hard_limit(kind, part, limit, observed)?;
            }
            return Err(format!("{limit_error_prefix}: {observed} > {limit}"));
        }
        Err(LimitedJsonError::Serialize(error)) => return Err(format!("serialize error: {error}")),
    };
    if let Some(reporter) = reporter {
        reporter.observe_hard_limit(kind, part, limit, bytes.len() as u64)?;
    }
    Ok(bytes)
}

/// Exact resource measurements for a serde JSON serialization.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct JsonMeasurement {
    pub json_bytes: u64,
    pub string_value_utf8_bytes: u64,
}

#[derive(Default)]
struct JsonMeasurementWriter {
    json_bytes: u64,
    string_value_utf8_bytes: u64,
    current_string_bytes: u64,
    pending_string_bytes: Option<u64>,
    in_string: bool,
    escaped: bool,
    unicode_digits: u8,
    unicode_value: u16,
    pending_high_surrogate: Option<u16>,
}

impl JsonMeasurementWriter {
    fn checked_add(target: &mut u64, amount: u64) -> io::Result<()> {
        *target = target.checked_add(amount).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "JSON resource measurement overflow",
            )
        })?;
        Ok(())
    }

    fn flush_pending_surrogate(&mut self) -> io::Result<()> {
        if self.pending_high_surrogate.take().is_some() {
            // A lone UTF-16 surrogate decodes as U+FFFD, which occupies three
            // UTF-8 bytes. serde strings cannot contain one, but accepting it
            // here keeps the streaming JSON scanner robust.
            Self::checked_add(&mut self.current_string_bytes, 3)?;
        }
        Ok(())
    }

    fn finish_unicode_escape(&mut self) -> io::Result<()> {
        let unit = std::mem::take(&mut self.unicode_value);
        match unit {
            0xD800..=0xDBFF => {
                self.flush_pending_surrogate()?;
                self.pending_high_surrogate = Some(unit);
            }
            0xDC00..=0xDFFF => {
                if self.pending_high_surrogate.take().is_some() {
                    Self::checked_add(&mut self.current_string_bytes, 4)?;
                } else {
                    Self::checked_add(&mut self.current_string_bytes, 3)?;
                }
            }
            _ => {
                self.flush_pending_surrogate()?;
                let width = if unit <= 0x7F {
                    1
                } else if unit <= 0x7FF {
                    2
                } else {
                    3
                };
                Self::checked_add(&mut self.current_string_bytes, width)?;
            }
        }
        Ok(())
    }

    fn finish(mut self) -> io::Result<JsonMeasurement> {
        if self.in_string || self.escaped || self.unicode_digits != 0 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "serializer ended inside a JSON string",
            ));
        }
        if let Some(bytes) = self.pending_string_bytes.take() {
            Self::checked_add(&mut self.string_value_utf8_bytes, bytes)?;
        }
        Ok(JsonMeasurement {
            json_bytes: self.json_bytes,
            string_value_utf8_bytes: self.string_value_utf8_bytes,
        })
    }
}

impl Write for JsonMeasurementWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let byte_count = u64::try_from(bytes.len()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "JSON write length does not fit u64",
            )
        })?;
        Self::checked_add(&mut self.json_bytes, byte_count)?;

        for &byte in bytes {
            if self.in_string {
                if self.unicode_digits != 0 {
                    let nibble = match byte {
                        b'0'..=b'9' => u16::from(byte - b'0'),
                        b'a'..=b'f' => u16::from(byte - b'a' + 10),
                        b'A'..=b'F' => u16::from(byte - b'A' + 10),
                        _ => {
                            return Err(io::Error::new(
                                io::ErrorKind::InvalidData,
                                "serializer emitted an invalid JSON unicode escape",
                            ));
                        }
                    };
                    self.unicode_value = self
                        .unicode_value
                        .checked_mul(16)
                        .and_then(|value| value.checked_add(nibble))
                        .ok_or_else(|| {
                            io::Error::new(
                                io::ErrorKind::InvalidData,
                                "JSON unicode escape overflow",
                            )
                        })?;
                    self.unicode_digits -= 1;
                    if self.unicode_digits == 0 {
                        self.escaped = false;
                        self.finish_unicode_escape()?;
                    }
                } else if self.escaped {
                    if byte == b'u' {
                        self.unicode_digits = 4;
                        self.unicode_value = 0;
                    } else {
                        self.flush_pending_surrogate()?;
                        Self::checked_add(&mut self.current_string_bytes, 1)?;
                        self.escaped = false;
                    }
                } else {
                    match byte {
                        b'\\' => self.escaped = true,
                        b'"' => {
                            self.flush_pending_surrogate()?;
                            self.in_string = false;
                            self.pending_string_bytes = Some(self.current_string_bytes);
                            self.current_string_bytes = 0;
                        }
                        _ => {
                            self.flush_pending_surrogate()?;
                            Self::checked_add(&mut self.current_string_bytes, 1)?;
                        }
                    }
                }
                continue;
            }

            if let Some(pending) = self.pending_string_bytes {
                if byte.is_ascii_whitespace() {
                    continue;
                }
                self.pending_string_bytes = None;
                if byte != b':' {
                    Self::checked_add(&mut self.string_value_utf8_bytes, pending)?;
                }
            }
            if byte == b'"' {
                self.in_string = true;
            }
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Measures exact serde JSON bytes and decoded UTF-8 bytes in string values.
pub fn measure_json<T: Serialize>(value: &T) -> Result<JsonMeasurement, String> {
    let mut counter = JsonMeasurementWriter::default();
    serde_json::to_writer(&mut counter, value)
        .map_err(|error| format!("serialize measurement error: {error}"))?;
    counter
        .finish()
        .map_err(|error| format!("serialize measurement error: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bounded_serialization_preserves_bytes_and_rejects_the_first_overflow() {
        let value = serde_json::json!({"text": "quote\" slash\\ é😀", "items": [1, null, true]});
        let expected = serde_json::to_vec(&value).unwrap();
        assert_eq!(
            serialize_json_limited(&value, expected.len() as u64).unwrap(),
            expected
        );
        let limit = expected.len() as u64 - 1;
        match serialize_json_limited(&value, limit).unwrap_err() {
            LimitedJsonError::LimitExceeded { observed } => {
                assert!(observed > limit);
            }
            other => panic!("expected a byte limit error, got {other:?}"),
        }
    }

    #[test]
    fn reported_serialization_keeps_limit_and_serde_error_formats() {
        let value = serde_json::json!({"text": "é😀"});
        let limit = serde_json::to_vec(&value).unwrap().len() as u64 - 1;
        let error = serialize_json_with_limit(
            &value,
            None,
            HardResourceLimitKind::WorksheetJsonBytes,
            Some("xl/worksheets/sheet1.xml"),
            limit,
            "worksheet JSON exceeds its hard ceiling",
        )
        .unwrap_err();
        let observed = error
            .strip_prefix("worksheet JSON exceeds its hard ceiling: ")
            .and_then(|rest| rest.strip_suffix(&format!(" > {limit}")))
            .and_then(|rest| rest.parse::<u64>().ok())
            .expect("limit error retains the observed/limit format");
        assert!(observed > limit);

        struct FailingValue;
        impl Serialize for FailingValue {
            fn serialize<S: serde::Serializer>(&self, _: S) -> Result<S::Ok, S::Error> {
                Err(serde::ser::Error::custom("failed value"))
            }
        }
        assert_eq!(
            serialize_json_with_limit(
                &FailingValue,
                None,
                HardResourceLimitKind::WorksheetJsonBytes,
                None,
                1,
                "worksheet JSON exceeds its hard ceiling",
            )
            .unwrap_err(),
            "serialize error: failed value"
        );
    }

    #[test]
    fn matches_exact_serde_json_bytes_and_excludes_property_names() {
        let value = serde_json::json!({
            "ignored-key-é": "quote\" slash\\ newline\n",
            "unicode-key-😀": "é😀",
            "array": [null, true, -0.0, "x"]
        });
        let measured = measure_json(&value).unwrap();
        assert_eq!(
            measured.json_bytes as usize,
            serde_json::to_vec(&value).unwrap().len()
        );
        assert_eq!(
            measured.string_value_utf8_bytes,
            "quote\" slash\\ newline\n".len() as u64 + "é😀".len() as u64 + 1
        );
    }

    #[test]
    fn decodes_controls_unicode_escapes_and_surrogates_across_writes() {
        let json = br#"{"ignored":"\b\t\n\f\r\u0000\u000b\u001f","pair":"\uD83D\uDE00","lone-high":"\uD800","lone-low":"\uDC00"}"#;
        let mut writer = JsonMeasurementWriter::default();
        for chunk in json.chunks(3) {
            writer.write_all(chunk).unwrap();
        }
        let measured = writer.finish().unwrap();
        assert_eq!(measured.json_bytes, json.len() as u64);
        assert_eq!(measured.string_value_utf8_bytes, 8 + 4 + 3 + 3);
    }

    #[test]
    fn checked_arithmetic_rejects_counter_overflow() {
        let mut writer = JsonMeasurementWriter {
            json_bytes: u64::MAX,
            ..Default::default()
        };
        assert_eq!(
            writer.write(b"x").unwrap_err().kind(),
            io::ErrorKind::InvalidData
        );

        let writer = JsonMeasurementWriter {
            string_value_utf8_bytes: u64::MAX,
            pending_string_bytes: Some(1),
            ..Default::default()
        };
        assert_eq!(
            writer.finish().unwrap_err().kind(),
            io::ErrorKind::InvalidData
        );
    }
}
