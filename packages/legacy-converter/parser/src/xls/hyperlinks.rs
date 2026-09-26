//! MS-XLS 2.4.140 HLink records (MS-OSHARED 2.3.7.1 Hyperlink Object) as
//! the XLSX model's hyperlinks, as the XLSX parser reads ECMA-376 18.3.1.47
//! `hyperlink` (first cell of `ref`, external target, `location`,
//! `display`).
//!
//! URL monikers become the external target and file monikers their path;
//! composite, item and anti monikers are not projected and reject the
//! workbook. HLinkTooltip (2.4.141) ToolTips are validated but not carried:
//! the XLSX model has no ToolTip, and its parser ignores `tooltip` too.
//! The display name is kept as stored; Excel's .xlsx counterparts omit
//! `display` for some links, and the model does not render it either way.
//! The viewer follows `url`/`location`, which match the counterparts.

use super::{u16_at, u32_at, unsupported};

fn truncated() -> String {
    unsupported("truncated XLS hyperlink")
}

/// Standard link CLSID {79EAC9D0-BAF9-11CE-8C82-00AA004BA90B}.
const STD_LINK: [u8; 16] = [
    0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
/// URL moniker CLSID {79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}.
const URL_MONIKER: [u8; 16] = [
    0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
/// File moniker CLSID {00000303-0000-0000-C000-000000000046}.
const FILE_MONIKER: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

fn utf16(data: &[u8], offset: usize, units: usize) -> Result<String, String> {
    let bytes = data
        .get(offset..offset.checked_add(units * 2).ok_or_else(truncated)?)
        .ok_or_else(truncated)?;
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    String::from_utf16(&units).map_err(|_| truncated())
}

/// HyperlinkString (2.3.7.9): character count including the terminating
/// NULL, then the characters.
fn hyperlink_string(data: &[u8], offset: usize) -> Result<(String, usize), String> {
    let count = usize::try_from(u32_at(data, offset)?).map_err(|_| truncated())?;
    if count == 0 || count > 32767 {
        return Err(unsupported("invalid XLS hyperlink string"));
    }
    let text = utf16(data, offset + 4, count - 1)?;
    if u16_at(data, offset + 4 + (count - 1) * 2)? != 0 {
        return Err(unsupported("unterminated XLS hyperlink string"));
    }
    Ok((text, 4 + count * 2))
}

/// HyperlinkMoniker (2.3.7.2) of a URL or file moniker.
fn moniker(data: &[u8], offset: usize) -> Result<(String, usize), String> {
    let clsid = data.get(offset..offset + 16).ok_or_else(truncated)?;
    let at = offset + 16;
    if clsid == URL_MONIKER {
        // URLMoniker (2.3.7.6): length, NULL-terminated url, optional
        // 24-byte serialization tail.
        let length = usize::try_from(u32_at(data, at)?).map_err(|_| truncated())?;
        // `length` is an input u32: 32-bit `usize` (wasm32) cannot add it
        // unchecked.
        let body = (at + 4)
            .checked_add(length)
            .and_then(|end| data.get(at + 4..end))
            .ok_or_else(truncated)?;
        let end = body
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .ok_or_else(|| unsupported("unterminated XLS hyperlink URL"))?;
        let url = utf16(body, 0, end)?;
        let url_bytes = (end + 1) * 2;
        if length != url_bytes && length != url_bytes + 24 {
            return Err(unsupported("invalid XLS hyperlink URL moniker"));
        }
        return Ok((url, 16 + 4 + length));
    }
    if clsid == FILE_MONIKER {
        // FileMoniker (2.3.7.8).
        let anti = usize::from(u16_at(data, at)?);
        let ansi_length = usize::try_from(u32_at(data, at + 2)?).map_err(|_| truncated())?;
        if ansi_length == 0 || ansi_length > 32767 {
            return Err(unsupported("invalid XLS hyperlink file moniker"));
        }
        let ansi = data
            .get(at + 6..at + 6 + ansi_length)
            .ok_or_else(truncated)?;
        if ansi.last() != Some(&0) {
            return Err(unsupported("unterminated XLS hyperlink file path"));
        }
        let mut next = at + 6 + ansi_length;
        if u16_at(data, next + 2)? != 0xdead {
            return Err(unsupported("invalid XLS hyperlink file moniker"));
        }
        next += 4 + 16 + 4;
        let unicode_size = usize::try_from(u32_at(data, next)?).map_err(|_| truncated())?;
        next += 4;
        let path = if unicode_size > 0 {
            let bytes = usize::try_from(u32_at(data, next)?).map_err(|_| truncated())?;
            if u16_at(data, next + 4)? != 3
                || bytes.checked_add(6) != Some(unicode_size)
                || bytes % 2 != 0
            {
                return Err(unsupported("invalid XLS hyperlink file moniker"));
            }
            let path = utf16(data, next + 6, bytes / 2)?;
            next += unicode_size;
            path
        } else {
            // ANSI path in the workbook code page; only ASCII is decoded
            // without the code page.
            let bytes = &ansi[..ansi.len() - 1];
            if !bytes.is_ascii() {
                return Err(unsupported("non-ASCII XLS hyperlink file path"));
            }
            bytes.iter().map(|&byte| char::from(byte)).collect()
        };
        let path = format!("{}{}", "..\\".repeat(anti), path);
        return Ok((path, next - offset));
    }
    Err(unsupported("unsupported XLS hyperlink moniker"))
}

/// One HLink record as a model hyperlink (first cell of its range).
pub(super) fn hlink(data: &[u8]) -> Result<xlsx_model::Hyperlink, String> {
    let (row, col) = (u16_at(data, 0)?, u16_at(data, 4)?);
    if row > u16_at(data, 2)? || col > u16_at(data, 6)? || col > 0x00ff {
        return Err(unsupported("invalid XLS hyperlink range"));
    }
    if data.get(8..24) != Some(&STD_LINK[..]) || u32_at(data, 24)? != 2 {
        return Err(unsupported("invalid XLS hyperlink header"));
    }
    let flags = u32_at(data, 28)?;
    let bit = |index: u32| flags & (1 << index) != 0;
    let mut at = 32;
    let mut display = None;
    let mut url = None;
    let mut location = None;
    if bit(4) {
        let (text, size) = hyperlink_string(data, at)?;
        display = Some(text);
        at += size;
    }
    if bit(7) {
        // Target frame names select a browser frame; not display state.
        at += hyperlink_string(data, at)?.1;
    }
    if bit(0) {
        let (target, size) = if bit(8) {
            hyperlink_string(data, at)?
        } else {
            moniker(data, at)?
        };
        url = Some(target);
        at += size;
    } else if bit(8) {
        return Err(unsupported("invalid XLS hyperlink flags"));
    }
    if bit(3) {
        let (text, size) = hyperlink_string(data, at)?;
        location = Some(text);
        at += size;
    }
    if bit(5) {
        at += 16;
    }
    if bit(6) {
        at += 8;
    }
    if at != data.len() {
        return Err(unsupported("unexpected XLS hyperlink tail"));
    }
    if url.is_none() && location.is_none() {
        return Err(unsupported("XLS hyperlink without a target"));
    }
    Ok(xlsx_model::Hyperlink {
        col: u32::from(col) + 1,
        row: u32::from(row) + 1,
        url,
        location,
        display,
    })
}

/// HLinkTooltip (2.4.141): validated only (see the module note).
pub(super) fn tooltip(data: &[u8]) -> Result<(), String> {
    if u16_at(data, 0)? != 0x0800 || data.len() < 14 || !data.len().is_multiple_of(2) {
        return Err(unsupported("invalid XLS hyperlink ToolTip"));
    }
    let units = (data.len() - 10) / 2;
    if !(2..=256).contains(&units) || u16_at(data, data.len() - 2)? != 0 {
        return Err(unsupported("invalid XLS hyperlink ToolTip"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn string(text: &str) -> Vec<u8> {
        let units: Vec<u16> = text.encode_utf16().chain([0]).collect();
        let mut data = (units.len() as u32).to_le_bytes().to_vec();
        data.extend(units.iter().flat_map(|unit| unit.to_le_bytes()));
        data
    }

    fn header(flags: u32) -> Vec<u8> {
        let mut data = Vec::new();
        for value in [1u16, 1, 2, 2] {
            data.extend(value.to_le_bytes());
        }
        data.extend(STD_LINK);
        data.extend(2u32.to_le_bytes());
        data.extend(flags.to_le_bytes());
        data
    }

    #[test]
    fn internal_and_url_links_project_their_targets() {
        let mut data = header(0x1c);
        data.extend(string("Diet"));
        data.extend(string("Sheet!A1"));
        let link = hlink(&data).unwrap();
        assert_eq!((link.row, link.col), (2, 3));
        assert_eq!(link.location.as_deref(), Some("Sheet!A1"));
        assert_eq!(link.display.as_deref(), Some("Diet"));
        assert!(link.url.is_none());

        let mut data = header(0x03);
        data.extend(URL_MONIKER);
        let url: Vec<u8> = "https://e.x/"
            .encode_utf16()
            .chain([0])
            .flat_map(|u| u.to_le_bytes())
            .collect();
        data.extend((url.len() as u32).to_le_bytes());
        data.extend(url);
        let link = hlink(&data).unwrap();
        assert_eq!(link.url.as_deref(), Some("https://e.x/"));

        // A maximal URL moniker length is truncated input, not an overflow.
        let mut huge = header(0x03);
        huge.extend(URL_MONIKER);
        huge.extend(u32::MAX.to_le_bytes());
        huge.extend([0; 8]);
        assert_eq!(hlink(&huge).unwrap_err(), truncated());

        // Item monikers are not projected; trailing bytes reject.
        let mut item = header(0x01);
        item.extend([0x04, 0x03, 0, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46]);
        assert!(hlink(&item).is_err());
        let mut tail = header(0x08);
        tail.extend(string("A1"));
        tail.push(0);
        assert!(hlink(&tail).is_err());
    }

    #[test]
    fn file_moniker_unicode_length_rejects_maximal_sizes() {
        let mut data = header(0x01);
        data.extend(FILE_MONIKER);
        data.extend(0u16.to_le_bytes());
        data.extend(2u32.to_le_bytes());
        data.extend(*b"a\0");
        data.extend([0xff, 0xff, 0xad, 0xde]);
        data.extend([0; 20]);
        // cbUnicodePathSize = 5, cbUnicodePathBytes = u32::MAX: the
        // declared sizes must not wrap into agreement.
        data.extend(5u32.to_le_bytes());
        data.extend(u32::MAX.to_le_bytes());
        data.extend(3u16.to_le_bytes());
        assert_eq!(
            moniker(&data, 32).unwrap_err(),
            unsupported("invalid XLS hyperlink file moniker")
        );
    }
}
