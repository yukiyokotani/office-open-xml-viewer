//! Owning pull/acknowledge boundary for a direct legacy Word document.

use super::{direct_model::DirectDocResult, pictures::DirectPictureResource};
use docx_model::{BodyElement, Document, StreamedDocumentUnit};
use ooxml_common::{
    json_measurement::measure_json,
    pull::insufficient_credit_error,
    resource::{
        HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES, HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
        HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES,
    },
};

struct Active {
    operation_id: u32,
    generation: u32,
    next_sequence: u32,
    accepted_json_bytes: u64,
    body: std::vec::IntoIter<BodyElement>,
    metadata: Option<Document>,
}

struct Prepared {
    operation_id: u32,
    generation: u32,
    sequence: u32,
    bytes: Option<Vec<u8>>,
    byte_length: usize,
    done: bool,
    accepted_json_bytes_after: u64,
}

pub(crate) struct DirectCursor {
    pending: Option<Document>,
    active: Option<Active>,
    prepared: Option<Prepared>,
    resources: Option<Vec<DirectPictureResource>>,
    opened: bool,
    limits: JsonLimits,
    failure: Option<String>,
}

#[derive(Clone, Copy)]
struct JsonLimits {
    body: u64,
    bootstrap: u64,
    retained: u64,
}

impl DirectCursor {
    pub(crate) fn new(mut result: DirectDocResult) -> Result<Self, String> {
        result.resources.sort_unstable_by(|a, b| a.key.cmp(&b.key));
        if result
            .resources
            .windows(2)
            .any(|pair| pair[0].key == pair[1].key)
        {
            return Err("duplicate DOC direct resource key".into());
        }
        Ok(Self {
            pending: Some(result.document),
            active: None,
            prepared: None,
            resources: Some(result.resources),
            opened: false,
            limits: JsonLimits {
                body: HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
                bootstrap: HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
                retained: HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES,
            },
            failure: None,
        })
    }

    pub(crate) fn open_document_cursor(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        self.healthy()?;
        if operation_id == 0 || generation == 0 {
            return Err("operation id and generation must be positive".into());
        }
        if self.opened {
            return Err("document cursor cannot be reopened".into());
        }
        let mut document = self.pending.take().ok_or("document model is unavailable")?;
        let body = std::mem::take(&mut document.body).into_iter();
        self.active = Some(Active {
            operation_id,
            generation,
            next_sequence: 0,
            accepted_json_bytes: 0,
            body,
            metadata: Some(document),
        });
        self.opened = true;
        Ok(())
    }

    pub(crate) fn pull_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: usize,
    ) -> Result<Vec<u8>, String> {
        self.healthy()?;
        if operation_id == 0 || generation == 0 || byte_credit == 0 {
            return Err("operation id, generation, and byte credit must be positive".into());
        }
        if self.prepared.is_none() {
            let identity = self
                .active
                .as_ref()
                .map(|active| (active.next_sequence, active.operation_id, active.generation))
                .ok_or("document cursor is not active")?;
            if identity != (sequence, operation_id, generation) {
                return Err("document cursor identity or sequence is stale".into());
            }
            match self.prepare(sequence, operation_id, generation) {
                Ok(prepared) => self.prepared = Some(prepared),
                Err(error) => {
                    self.active.take();
                    self.failure = Some(error.clone());
                    return Err(error);
                }
            }
        }
        let prepared = self.prepared.as_mut().expect("prepared above");
        if (
            prepared.sequence,
            prepared.operation_id,
            prepared.generation,
        ) != (sequence, operation_id, generation)
        {
            return Err("another document unit is awaiting acknowledgement".into());
        }
        if prepared.bytes.is_none() {
            return Err("document unit must be acknowledged before another pull".into());
        }
        if prepared.byte_length > byte_credit {
            return Err(insufficient_credit_error(prepared.byte_length, byte_credit));
        }
        Ok(prepared.bytes.take().expect("prepared bytes checked"))
    }

    pub(crate) fn document_chunk_done(&self) -> Result<bool, String> {
        self.prepared
            .as_ref()
            .map(|p| p.done)
            .ok_or_else(|| "no document unit is awaiting acknowledgement".into())
    }

    pub(crate) fn acknowledge_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        self.healthy()?;
        let p = self
            .prepared
            .as_ref()
            .ok_or("no document unit is awaiting acknowledgement")?;
        if (p.sequence, p.operation_id, p.generation) != (sequence, operation_id, generation) {
            return Err("document acknowledgement identity is stale or invalid".into());
        }
        if p.bytes.is_some() {
            return Err("document unit cannot be acknowledged before delivery".into());
        }
        let done = p.done;
        let accepted = p.accepted_json_bytes_after;
        self.prepared.take();
        if done {
            self.active.take();
        } else if let Some(active) = self.active.as_mut() {
            active.next_sequence = active
                .next_sequence
                .checked_add(1)
                .ok_or("document sequence overflow")?;
            active.accepted_json_bytes = accepted;
        }
        Ok(())
    }

    pub(crate) fn cancel_document_cursor(&mut self) -> Result<(), String> {
        self.healthy()?;
        self.prepared.take();
        self.active.take();
        self.pending.take();
        self.opened = true;
        Ok(())
    }

    pub(crate) fn close_document_session(&mut self) {
        // Matches DocxArchive: close the destructive document operation while
        // retaining admitted media for the viewer's later lazy image reads.
        self.prepared.take();
        self.active.take();
        self.pending.take();
        self.opened = true;
    }

    pub(crate) fn assert_healthy(&self) -> Result<(), String> {
        self.healthy()
    }

    pub(crate) fn extract_image(&self, key: &str) -> Result<Vec<u8>, String> {
        let source = &self.resource(key)?.bytes;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(source.len())
            .map_err(|_| "DOC direct resource allocation failed".to_string())?;
        bytes.extend_from_slice(source);
        Ok(bytes)
    }

    pub(crate) fn resource_mime_type(&self, key: &str) -> Result<&'static str, String> {
        Ok(self.resource(key)?.mime_type)
    }

    fn resource(&self, key: &str) -> Result<&DirectPictureResource, String> {
        self.healthy()?;
        let resources = self
            .resources
            .as_ref()
            .ok_or("DOC direct cursor is closed")?;
        resources
            .binary_search_by(|r| r.key.as_str().cmp(key))
            .ok()
            .map(|index| &resources[index])
            .ok_or_else(|| "unknown DOC direct resource key".into())
    }

    fn healthy(&self) -> Result<(), String> {
        self.failure.clone().map_or(Ok(()), Err)
    }

    fn prepare(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<Prepared, String> {
        let active = self
            .active
            .as_mut()
            .expect("active identity checked by caller");
        let unit = if let Some(element) = active.body.next() {
            StreamedDocumentUnit::Body {
                body: vec![element],
            }
        } else {
            StreamedDocumentUnit::Complete {
                document: Box::new(
                    active
                        .metadata
                        .take()
                        .ok_or("document terminal is unavailable")?,
                ),
            }
        };
        let done = matches!(unit, StreamedDocumentUnit::Complete { .. });
        let measured = measure_json(&unit)?.json_bytes;
        let unit_limit = if done {
            self.limits.bootstrap
        } else {
            self.limits.body
        };
        if measured > unit_limit {
            return Err("document cursor JSON exceeds its hard ceiling".into());
        }
        let accepted = active
            .accepted_json_bytes
            .checked_add(measured)
            .ok_or("document retained model JSON size overflow")?;
        if accepted > self.limits.retained {
            return Err("document retained model JSON exceeds its hard ceiling".into());
        }
        let capacity =
            usize::try_from(measured).map_err(|_| "document JSON size exceeds this platform")?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|_| "document JSON allocation failed")?;
        serde_json::to_writer(&mut bytes, &unit).map_err(|e| format!("serialize error: {e}"))?;
        debug_assert_eq!(bytes.len() as u64, measured);
        Ok(Prepared {
            operation_id,
            generation,
            sequence,
            byte_length: bytes.len(),
            bytes: Some(bytes),
            done,
            accepted_json_bytes_after: accepted,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_model::{DocParagraph, DocTable, HeaderFooter, HeadersFooters};

    fn result(body: Vec<BodyElement>) -> DirectDocResult {
        let mut document = Document::default();
        document.body = body;
        document.headers = HeadersFooters {
            default: Some(HeaderFooter {
                body: vec![BodyElement::Paragraph(Box::new(DocParagraph::default()))],
            }),
            ..HeadersFooters::default()
        };
        DirectDocResult {
            document,
            resources: vec![DirectPictureResource {
                key: "legacy-doc/image/7".into(),
                mime_type: "image/png",
                bytes: vec![1, 2, 3],
            }],
        }
    }

    fn pull(cursor: &mut DirectCursor, sequence: u32, credit: usize) -> Vec<u8> {
        cursor.pull_document_chunk(sequence, 7, 9, credit).unwrap()
    }

    #[test]
    fn streams_one_owned_body_element_then_bodyless_metadata() {
        let mut cursor = DirectCursor::new(result(vec![
            BodyElement::Table(Box::new(DocTable::default())),
            BodyElement::Paragraph(Box::new(DocParagraph::default())),
        ]))
        .unwrap();
        cursor.open_document_cursor(7, 9).unwrap();
        for sequence in 0..2 {
            let value: serde_json::Value =
                serde_json::from_slice(&pull(&mut cursor, sequence, usize::MAX)).unwrap();
            assert_eq!(value["kind"], "body");
            assert_eq!(value["body"].as_array().unwrap().len(), 1);
            cursor.acknowledge_document_chunk(sequence, 7, 9).unwrap();
        }
        let value: serde_json::Value =
            serde_json::from_slice(&pull(&mut cursor, 2, usize::MAX)).unwrap();
        assert_eq!(value["kind"], "complete");
        assert_eq!(value["document"]["body"], serde_json::json!([]));
        assert!(value["document"]["headers"]["default"].is_object());
        assert!(cursor.document_chunk_done().unwrap());
        cursor.acknowledge_document_chunk(2, 7, 9).unwrap();
    }

    #[test]
    fn validates_identity_credit_delivery_and_destructive_lifecycle() {
        let mut cursor = DirectCursor::new(result(vec![BodyElement::Paragraph(Box::new(
            DocParagraph::default(),
        ))]))
        .unwrap();
        assert!(cursor.open_document_cursor(0, 1).is_err());
        cursor.open_document_cursor(7, 9).unwrap();
        assert!(cursor.pull_document_chunk(1, 7, 9, 1).is_err());
        let error = cursor.pull_document_chunk(0, 7, 9, 1).unwrap_err();
        assert!(error.starts_with("OOXML_INSUFFICIENT_CREDIT:"));
        assert!(cursor.acknowledge_document_chunk(0, 7, 9).is_err());
        let envelope: serde_json::Value =
            serde_json::from_str(error.split_once(':').unwrap().1).unwrap();
        let exact = envelope["requiredBytes"].as_u64().unwrap() as usize;
        let bytes = pull(&mut cursor, 0, exact);
        assert_eq!(bytes.len(), exact);
        assert!(cursor.pull_document_chunk(0, 7, 9, exact).is_err());
        assert!(cursor.acknowledge_document_chunk(0, 7, 8).is_err());
        cursor.acknowledge_document_chunk(0, 7, 9).unwrap();
        cursor.cancel_document_cursor().unwrap();
        assert!(cursor.open_document_cursor(8, 10).is_err());
    }

    #[test]
    fn empty_document_and_retained_json_ceiling_are_enforced() {
        let mut empty = DirectCursor::new(result(vec![])).unwrap();
        empty.open_document_cursor(7, 9).unwrap();
        let value: serde_json::Value =
            serde_json::from_slice(&pull(&mut empty, 0, usize::MAX)).unwrap();
        assert_eq!(value["kind"], "complete");

        let mut limited = DirectCursor::new(result(vec![BodyElement::Paragraph(Box::new(
            DocParagraph::default(),
        ))]))
        .unwrap();
        limited.limits.retained = 1;
        limited.open_document_cursor(7, 9).unwrap();
        assert!(limited
            .pull_document_chunk(0, 7, 9, usize::MAX)
            .unwrap_err()
            .contains("retained model JSON"));
        let fatal = limited.assert_healthy().unwrap_err();
        assert_eq!(
            limited
                .pull_document_chunk(0, 7, 9, usize::MAX)
                .unwrap_err(),
            fatal
        );
    }

    #[test]
    fn body_and_terminal_ceiling_failures_are_fatal_without_skipping() {
        let mut body = DirectCursor::new(result(vec![
            BodyElement::Paragraph(Box::new(DocParagraph::default())),
            BodyElement::Table(Box::new(DocTable::default())),
        ]))
        .unwrap();
        body.limits.body = 1;
        body.open_document_cursor(7, 9).unwrap();
        let error = body.pull_document_chunk(0, 7, 9, usize::MAX).unwrap_err();
        assert!(error.contains("JSON exceeds"));
        assert_eq!(
            body.pull_document_chunk(0, 7, 9, usize::MAX).unwrap_err(),
            error
        );
        assert!(body.active.is_none());
        assert!(body.prepared.is_none());

        let mut terminal = DirectCursor::new(result(vec![])).unwrap();
        terminal.limits.bootstrap = 1;
        terminal.open_document_cursor(7, 9).unwrap();
        let error = terminal
            .pull_document_chunk(0, 7, 9, usize::MAX)
            .unwrap_err();
        assert!(error.contains("JSON exceeds"));
        assert_eq!(terminal.assert_healthy().unwrap_err(), error);
        terminal.close_document_session();
        assert_eq!(terminal.assert_healthy().unwrap_err(), error);
    }

    #[test]
    fn resources_survive_ack_cancel_and_document_close() {
        let mut cursor = DirectCursor::new(result(vec![])).unwrap();
        assert_eq!(
            cursor.resource_mime_type("legacy-doc/image/7").unwrap(),
            "image/png"
        );
        cursor.open_document_cursor(7, 9).unwrap();
        pull(&mut cursor, 0, usize::MAX);
        cursor.acknowledge_document_chunk(0, 7, 9).unwrap();
        cursor.close_document_session();
        cursor.close_document_session();
        assert_eq!(
            cursor.extract_image("legacy-doc/image/7").unwrap(),
            [1, 2, 3]
        );
        assert!(cursor.extract_image("legacy-doc/image/8").is_err());
        assert!(cursor.open_document_cursor(7, 9).is_err());
    }

    #[test]
    fn duplicate_resource_keys_are_rejected() {
        let mut value = result(vec![]);
        value.resources.push(DirectPictureResource {
            key: "legacy-doc/image/7".into(),
            mime_type: "image/jpeg",
            bytes: vec![],
        });
        assert!(DirectCursor::new(value)
            .err()
            .unwrap()
            .contains("duplicate"));
    }
}
