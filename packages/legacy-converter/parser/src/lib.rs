//! Purpose-built passive Office 97-2003 readers.
//!
//! Each optional feature builds one experimental direct source that projects a
//! documented passive subset of an untrusted CFB container straight into the
//! shared renderer models: `direct-ppt` (PowerPoint 97-2003 into the PPTX
//! model). No path executes document code or creates an OOXML package;
//! unsupported versions and encryption fail closed.

#[cfg(all(target_arch = "wasm32", not(feature = "direct-ppt")))]
compile_error!("legacy-office-converter builds no WASM source without the `direct-ppt` feature");

#[cfg(feature = "direct-ppt")]
mod cfb;
#[cfg(feature = "fuzzing")]
pub mod fuzzing;
#[cfg(feature = "direct-ppt")]
mod officeart;
#[cfg(feature = "direct-ppt")]
mod ppt;
#[cfg(feature = "direct-ppt")]
mod ppt_wasm;

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
