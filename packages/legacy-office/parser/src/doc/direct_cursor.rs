//! Owning pull/acknowledge boundary for a direct legacy Word document: the
//! shared DOCX owned cursor plus the DOC's own revision-markup settings.

use super::direct_model::DirectDocResult;
use docx_model::owned_cursor::{OwnedDocumentCursor, OwnedDocumentResource};

/// The shared cursor's resource errors in this reader's established wording.
fn legacy_message(error: String) -> String {
    match error.as_str() {
        "duplicate document source resource key" => "duplicate DOC direct resource key".into(),
        "document source resource allocation failed" => {
            "DOC direct resource allocation failed".into()
        }
        "document source cursor is closed" => "DOC direct cursor is closed".into(),
        "unknown document source resource key" => "unknown DOC direct resource key".into(),
        _ => error,
    }
}

pub(crate) struct DirectCursor {
    revision_markup: crate::doc::settings::RevisionMarkup,
    inner: OwnedDocumentCursor,
}

impl DirectCursor {
    pub(crate) fn new(result: DirectDocResult) -> Result<Self, String> {
        let resources = result
            .resources
            .into_iter()
            .map(|resource| OwnedDocumentResource {
                key: resource.key,
                mime_type: resource.mime_type.to_owned(),
                bytes: resource.bytes,
            })
            .collect();
        Ok(Self {
            revision_markup: result.revision_markup,
            inner: OwnedDocumentCursor::new(result.document, resources).map_err(legacy_message)?,
        })
    }

    /// MS-DOC 2.7.2 DopBase fRMPrint of the source document: whether Word
    /// prints (and exports to PDF) its revision markup.
    pub(crate) fn revision_markup_in_print(&self) -> bool {
        self.revision_markup.in_print
    }

    /// Whether the projected model carries any revision mark.
    pub(crate) fn has_revision_marks(&self) -> bool {
        self.revision_markup.has_marks
    }

    /// MS-DOC 2.7.2 DopBase fRMView of the source document.
    pub(crate) fn revision_markup_on_screen(&self) -> bool {
        self.revision_markup.on_screen
    }

    pub(crate) fn open_document_cursor(
        &mut self,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        self.inner.open_document_cursor(operation_id, generation)
    }

    pub(crate) fn pull_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
        byte_credit: usize,
    ) -> Result<Vec<u8>, String> {
        self.inner
            .pull_document_chunk(sequence, operation_id, generation, byte_credit)
    }

    pub(crate) fn document_chunk_done(&self) -> Result<bool, String> {
        self.inner.document_chunk_done()
    }

    pub(crate) fn acknowledge_document_chunk(
        &mut self,
        sequence: u32,
        operation_id: u32,
        generation: u32,
    ) -> Result<(), String> {
        self.inner
            .acknowledge_document_chunk(sequence, operation_id, generation)
    }

    pub(crate) fn cancel_document_cursor(&mut self) -> Result<(), String> {
        self.inner.cancel_document_cursor()
    }

    pub(crate) fn close_document_session(&mut self) {
        self.inner.close_document_session()
    }

    pub(crate) fn assert_healthy(&self) -> Result<(), String> {
        self.inner.assert_healthy()
    }

    pub(crate) fn extract_image(&self, key: &str) -> Result<Vec<u8>, String> {
        self.inner.extract_image(key).map_err(legacy_message)
    }

    pub(crate) fn resource_mime_type(&self, key: &str) -> Result<&str, String> {
        self.inner.resource_mime_type(key).map_err(legacy_message)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_model::{BodyElement, Document};

    fn result() -> DirectDocResult {
        DirectDocResult {
            document: Document {
                body: vec![BodyElement::Paragraph(Box::default())],
                ..Document::default()
            },
            resources: vec![super::super::pictures::DirectPictureResource {
                key: "legacy-doc/image/7".into(),
                mime_type: "image/png",
                bytes: vec![1, 2, 3],
            }],
            revision_markup: Default::default(),
        }
    }

    #[test]
    fn revision_markup_settings_are_exposed_beside_the_model() {
        let cursor = DirectCursor::new(result()).unwrap();
        assert!(!cursor.revision_markup_in_print());
        assert!(!cursor.revision_markup_on_screen());
        let mut value = result();
        value.revision_markup = crate::doc::settings::RevisionMarkup {
            on_screen: true,
            in_print: false,
            has_marks: true,
        };
        let cursor = DirectCursor::new(value).unwrap();
        assert!(cursor.revision_markup_on_screen());
        assert!(!cursor.revision_markup_in_print());
        assert!(cursor.has_revision_marks());
    }

    #[test]
    fn picture_resources_keep_their_keys_and_media_types() {
        let mut cursor = DirectCursor::new(result()).unwrap();
        assert_eq!(
            cursor.resource_mime_type("legacy-doc/image/7").unwrap(),
            "image/png"
        );
        cursor.close_document_session();
        assert_eq!(
            cursor.extract_image("legacy-doc/image/7").unwrap(),
            [1, 2, 3]
        );
        assert_eq!(
            cursor.extract_image("legacy-doc/image/8").unwrap_err(),
            "unknown DOC direct resource key"
        );
        assert_eq!(
            cursor.resource_mime_type("legacy-doc/image/8").unwrap_err(),
            "unknown DOC direct resource key"
        );
    }

    #[test]
    fn duplicate_resource_keys_keep_the_doc_message() {
        let mut value = result();
        value
            .resources
            .push(super::super::pictures::DirectPictureResource {
                key: "legacy-doc/image/7".into(),
                mime_type: "image/jpeg",
                bytes: vec![],
            });
        assert_eq!(
            DirectCursor::new(value).err().unwrap(),
            "duplicate DOC direct resource key"
        );
    }
}
