//! Bounded OfficeArt / PowerPoint record headers (MS-ODRAW 2.2.1).
pub(crate) mod geometry;
mod metafile;
pub(crate) mod properties;
#[cfg(test)]
pub(crate) use metafile::tests::emf_test_blip;
pub(crate) mod raster;
pub(crate) mod stroke;

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct Record<'a> {
    pub version: u8,
    pub instance: u16,
    pub kind: u16,
    pub payload: &'a [u8],
}

pub(crate) fn record_with_end<'a>(
    bytes: &'a [u8],
    offset: usize,
    budget: &mut usize,
    context: &str,
) -> Result<(Record<'a>, usize), String> {
    if *budget == 0 {
        return Err(unsupported(format!("too many {context} records")));
    }
    *budget -= 1;
    let remaining = bytes
        .get(offset..)
        .filter(|tail| tail.len() >= 8)
        .ok_or_else(|| unsupported(format!("truncated {context} record header")))?;
    let options = u16::from_le_bytes(remaining[0..2].try_into().unwrap());
    let kind = u16::from_le_bytes(remaining[2..4].try_into().unwrap());
    let size = usize::try_from(u32::from_le_bytes(remaining[4..8].try_into().unwrap()))
        .map_err(|_| unsupported(format!("{context} record is too large")))?;
    let end = offset
        .checked_add(8)
        .and_then(|start| start.checked_add(size))
        .ok_or_else(|| unsupported(format!("{context} record range overflow")))?;
    let payload = bytes
        .get(offset + 8..end)
        .ok_or_else(|| {
            unsupported(format!(
                "truncated {context} record at offset {offset}: declared {size} bytes with {} available",
                bytes.len().saturating_sub(offset + 8),
            ))
        })?;
    Ok((
        Record {
            version: (options & 0x000f) as u8,
            instance: options >> 4,
            kind,
            payload,
        },
        end,
    ))
}
