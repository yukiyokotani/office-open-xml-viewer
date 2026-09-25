//! Purpose-built passive Office 97-2003 readers.
//!
//! Each optional feature builds one experimental direct source that projects a
//! documented passive subset of an untrusted CFB container straight into the
//! shared renderer models: `direct-doc` (Word 97-2003 into the DOCX model),
//! `direct-xls` (BIFF8 into the XLSX model) and `direct-ppt` (PowerPoint
//! 97-2003 into the PPTX model). No path executes document code or creates an
//! OOXML package; unsupported versions and encryption fail closed.

#[cfg(all(
    target_arch = "wasm32",
    not(any(feature = "direct-doc", feature = "direct-xls", feature = "direct-ppt"))
))]
compile_error!(
    "legacy-office-converter builds no WASM source without one of the \
     `direct-doc`, `direct-xls` or `direct-ppt` features"
);

#[cfg(any(feature = "direct-doc", feature = "direct-xls", feature = "direct-ppt"))]
mod cfb;
#[cfg(feature = "direct-doc")]
mod doc;
#[cfg(feature = "direct-doc")]
mod doc_wasm;
#[cfg(feature = "fuzzing")]
pub mod fuzzing;
#[cfg(feature = "direct-doc")]
mod lcid;
#[cfg(any(feature = "direct-doc", feature = "direct-xls", feature = "direct-ppt"))]
mod officeart;
#[cfg(feature = "direct-ppt")]
mod ppt;
#[cfg(feature = "direct-ppt")]
mod ppt_wasm;
#[cfg(feature = "direct-xls")]
mod xls;
#[cfg(feature = "direct-xls")]
mod xls_wasm;

/// A passive catalog image: BLIP store index, file extension and bytes.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub type XlsImage = (u32, &'static str, Vec<u8>);

/// Native development helper, not part of the production WASM contract.
/// Returns passive catalog images, not necessarily visible sheet objects.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_xls_images(data: &[u8]) -> Result<Vec<XlsImage>, String> {
    xls::inspect_images(&inspection_source(data)?)
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub use xls::drawing_anchors::{CellCorner, DrawingAnchor, PictureReference};

/// Native-only raw sheet anchor evidence. Not a visibility or image-admission API.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_xls_anchors(data: &[u8]) -> Result<Vec<DrawingAnchor>, String> {
    xls::inspect_anchors(&inspection_source(data)?)
}

/// Native development bindings. Each retained anchor references one supported
/// image here; groups, inherited visibility and layout remain uninterpreted.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub struct XlsPictureInspection {
    pub anchors: Vec<DrawingAnchor>,
    pub images: Vec<XlsImage>,
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_xls_pictures(data: &[u8]) -> Result<XlsPictureInspection, String> {
    xls::inspect_pictures(&inspection_source(data)?)
}

/// Native-only direct PPT renderer-model inspection session, not part of the
/// WASM contract.
#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub struct PptDirectInspection {
    inner: ppt::direct_session::DirectSession,
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
impl PptDirectInspection {
    pub fn slide_count(&self) -> usize {
        self.inner.slide_count()
    }

    pub fn size(&self) -> (u32, u32) {
        self.inner.size()
    }

    pub fn slide(&mut self, index: usize) -> Result<pptx_model::Slide, String> {
        self.inner.slide(index)
    }

    pub fn resource(&self, key: &str) -> Result<(&'static str, &[u8]), String> {
        let resource = self.inner.resource(key)?;
        Ok((resource.extension, resource.bytes))
    }
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
pub fn inspect_ppt_direct(data: &[u8]) -> Result<PptDirectInspection, String> {
    if data.len() > 256 * 1024 * 1024 {
        return Err("UNSUPPORTED:PPT inspection source byte budget exceeded".into());
    }
    let cfb = cfb::CompoundFile::open(data).map_err(|e| format!("UNSUPPORTED:{e}"))?;
    Ok(PptDirectInspection {
        inner: ppt::direct_session::DirectSession::new(&cfb)?,
    })
}

#[cfg(all(feature = "inspection", not(target_arch = "wasm32")))]
fn inspection_source(data: &[u8]) -> Result<cfb::CompoundFile<'_>, String> {
    if data.len() > 256 * 1024 * 1024 {
        return Err("UNSUPPORTED:XLS inspection source byte budget exceeded".into());
    }
    cfb::CompoundFile::open(data).map_err(|e| format!("UNSUPPORTED:{e}"))
}
