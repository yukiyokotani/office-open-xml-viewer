//! Native fuzz drivers (feature `fuzzing`), absent from production WASM.
//!
//! Each direct-source driver follows its WASM boundary's call sequence with
//! native types: open the compound file, build the owning session, then pull
//! every unit the session announces, as a host would. Errors are expected
//! outcomes; only a panic, a hang or unbounded allocation is a finding.
//! `wrapped_*` place arbitrary bytes into the streams of a root-linked
//! compound file so the format readers are reached behind the container.

#[cfg(any(feature = "direct-doc", feature = "direct-xls", feature = "direct-ppt"))]
use crate::cfb::{test_support::build_scoped_cfb, CompoundFile};

/// A per-unit credit large enough for any unit the budgets admit.
#[cfg(any(feature = "direct-doc", feature = "direct-ppt"))]
const BYTE_CREDIT: usize = 256 * 1024 * 1024;

/// `data` as the WordDocument and table streams of a root-linked DOC.
#[cfg(feature = "direct-doc")]
pub fn wrapped_doc(data: &[u8]) -> Vec<u8> {
    build_scoped_cfb(&[("WordDocument", data.to_vec()), ("0Table", data.to_vec())])
}

/// `data` as the Workbook stream of a root-linked XLS.
#[cfg(feature = "direct-xls")]
pub fn wrapped_xls(data: &[u8]) -> Vec<u8> {
    build_scoped_cfb(&[("Workbook", data.to_vec())])
}

/// `data` as the document and current-user streams of a root-linked PPT.
#[cfg(feature = "direct-ppt")]
pub fn wrapped_ppt(data: &[u8]) -> Vec<u8> {
    build_scoped_cfb(&[
        ("PowerPoint Document", data.to_vec()),
        ("Current User", data.to_vec()),
    ])
}

/// Direct DOC projection and its document cursor.
#[cfg(feature = "direct-doc")]
pub fn direct_doc(data: &[u8]) {
    use crate::doc::direct_cursor::DirectCursor;
    let Ok(cfb) = CompoundFile::open(data) else {
        return;
    };
    let Ok(result) = crate::doc::direct_model(&cfb, 64 * 1024 * 1024) else {
        return;
    };
    let Ok(mut cursor) = DirectCursor::new(result) else {
        return;
    };
    let _ = (
        cursor.revision_markup_in_print(),
        cursor.revision_markup_on_screen(),
        cursor.has_revision_marks(),
    );
    if cursor.open_document_cursor(1, 1).is_err() {
        return;
    }
    for sequence in 0u32.. {
        if cursor
            .pull_document_chunk(sequence, 1, 1, BYTE_CREDIT)
            .is_err()
        {
            break;
        }
        let done = cursor.document_chunk_done().unwrap_or(true);
        if cursor.acknowledge_document_chunk(sequence, 1, 1).is_err() || done {
            break;
        }
    }
    for index in 0..4 {
        let key = format!("legacy-doc/image/{index}");
        let _ = cursor.resource_mime_type(&key);
        let _ = cursor.extract_image(&key);
    }
    cursor.close_document_session();
}

/// Direct XLS session: host layout decision, bootstrap and every worksheet.
#[cfg(feature = "direct-xls")]
pub fn direct_xls(data: &[u8]) {
    use crate::xls::direct::DirectSession;
    let Ok(cfb) = CompoundFile::open(data) else {
        return;
    };
    let Ok(mut session) = DirectSession::new(&cfb) else {
        return;
    };
    let _ = session.measurement_font();
    let mdw = session.requires_measurement_decision().then_some(7.0);
    if session.configure_host_layout(mdw).is_err() {
        return;
    }
    let Ok(workbook) = session.bootstrap() else {
        return;
    };
    for (index, sheet) in workbook.workbook.sheets.iter().enumerate() {
        let _ = session.projected_sheet(index, &sheet.name);
    }
    for index in 0..4 {
        let _ = session.resource(&format!("xl/media/image{index}.png"));
    }
}

/// Direct PPT session through its slide cursor, including alternative shape
/// XML resolution on every shape that carries one.
#[cfg(feature = "direct-ppt")]
pub fn direct_ppt(data: &[u8]) {
    use crate::ppt::{direct_cursor::DirectCursor, direct_session::DirectSession};
    let Ok(cfb) = CompoundFile::open(data) else {
        return;
    };
    let Ok(session) = DirectSession::new(&cfb) else {
        return;
    };
    let slides = session.slide_count();
    let mut cursor = DirectCursor::new(session);
    if cursor.presentation_bootstrap().is_err() {
        return;
    }
    for index in 0..slides {
        let operation = u32::try_from(index).map_or(u32::MAX, |n| n.saturating_add(1));
        if cursor.pull_slide(index, operation, 1, BYTE_CREDIT).is_err()
            || cursor.acknowledge_slide(operation, 1).is_err()
        {
            break;
        }
    }
    for index in 0..4 {
        let _ = cursor.extract_image(&format!("legacy-ppt/image/{index}"));
    }
    let _ = cursor.close_presentation_session();
}

/// Alternative shape XML resolution against a fixed binary shape, with
/// `data` both as the whole metroBlob package and as the shape part of an
/// otherwise well-formed package.
#[cfg(feature = "direct-ppt")]
pub fn ppt_alternative(data: &[u8]) {
    crate::ppt::fuzz_alternative(data);
}
