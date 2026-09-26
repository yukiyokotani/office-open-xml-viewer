//! Version acquisition only; [MS-DOC] 2.5.14-15. FibBase.nFib can be
//! superseded by FibRgCswNew.nFibNew. Skip counted arrays without allocating.
use super::{u16_at, unsupported};

const CSW_OFFSET: usize = 32;
const EXPECTED_CSW: u16 = 14;
const CSLW_OFFSET: usize = CSW_OFFSET + 2 + EXPECTED_CSW as usize * 2;
const EXPECTED_CSLW: u16 = 22;
const CB_RG_FC_LCB_OFFSET: usize = CSLW_OFFSET + 2 + EXPECTED_CSLW as usize * 4;
const FIB_RG_FC_LCB_97_VALUES: u16 = 0x005d;
const FIXED_PREFIX_END: usize = CB_RG_FC_LCB_OFFSET + 2 + FIB_RG_FC_LCB_97_VALUES as usize * 8;

/// Validate the canonical Word 97 prefix assumed by the DOC readers' fixed
/// FibRgLw97 and FibRgFcLcb97 field offsets.
///
/// [MS-DOC] 2.5.1 requires the first two counts to be 0x000E and 0x0016.
/// Later FibRgFcLcb structures extend the 93-value Word 97 prefix, so their
/// additional values remain admissible.
pub(super) fn validate_fixed_prefix(word: &[u8]) -> Result<(), String> {
    if u16_at(word, CSW_OFFSET)? != EXPECTED_CSW || u16_at(word, CSLW_OFFSET)? != EXPECTED_CSLW {
        return Err(unsupported("noncanonical Word FIB fixed prefix"));
    }
    if u16_at(word, CB_RG_FC_LCB_OFFSET)? < FIB_RG_FC_LCB_97_VALUES {
        return Err(unsupported("short Word FIB fc/lcb prefix"));
    }
    word.get(..FIXED_PREFIX_END)
        .ok_or_else(|| unsupported("truncated Word FIB fc/lcb prefix"))?;
    Ok(())
}

#[cfg(test)]
pub(super) fn write_minimal_word97_header(word: &mut [u8]) {
    assert!(word.len() >= FIXED_PREFIX_END + 2);
    word[0..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    word[2..4].copy_from_slice(&0x00c1u16.to_le_bytes());
    word[CSW_OFFSET..CSW_OFFSET + 2].copy_from_slice(&EXPECTED_CSW.to_le_bytes());
    word[CSLW_OFFSET..CSLW_OFFSET + 2].copy_from_slice(&EXPECTED_CSLW.to_le_bytes());
    word[CB_RG_FC_LCB_OFFSET..CB_RG_FC_LCB_OFFSET + 2]
        .copy_from_slice(&FIB_RG_FC_LCB_97_VALUES.to_le_bytes());
    word[FIXED_PREFIX_END..FIXED_PREFIX_END + 2].copy_from_slice(&0u16.to_le_bytes());
}

pub(super) fn effective_version(word: &[u8]) -> Result<u16, String> {
    word.get(..32)
        .ok_or_else(|| unsupported("truncated Word FIB base"))?;
    let mut offset = 32usize;
    for unit in [2usize, 4, 8] {
        let count = usize::from(u16_at(word, offset)?);
        offset = offset
            .checked_add(2)
            .and_then(|start| {
                count
                    .checked_mul(unit)
                    .and_then(|len| start.checked_add(len))
            })
            .filter(|end| *end <= word.len())
            .ok_or_else(|| unsupported("truncated Word FIB counted array"))?;
    }
    let count = usize::from(u16_at(word, offset)?);
    let start = offset
        .checked_add(2)
        .ok_or_else(|| unsupported("Word FIB extension range overflow"))?;
    let end = count
        .checked_mul(2)
        .and_then(|len| start.checked_add(len))
        .filter(|end| *end <= word.len())
        .ok_or_else(|| unsupported("truncated Word FIB extension"))?;
    let version = if count == 0 {
        u16_at(word, 2)?
    } else {
        u16_at(&word[start..end], 0)?
    };
    if !matches!(version, 0x00c1 | 0x00d9 | 0x0101 | 0x010c | 0x0112)
        || (count != 0 && version == 0x00c1)
    {
        return Err(unsupported("unsupported Word FIB version"));
    }
    Ok(version)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fib(base: u16, extension: Option<u16>, counts: [u16; 3]) -> Vec<u8> {
        let mut bytes = vec![0; 32];
        bytes[2..4].copy_from_slice(&base.to_le_bytes());
        for (count, unit) in counts.into_iter().zip([2, 4, 8]) {
            bytes.extend(count.to_le_bytes());
            bytes.resize(bytes.len() + usize::from(count) * unit, 0xa5);
        }
        let count = extension.map_or(0u16, |version| if version == 0x0112 { 5 } else { 2 });
        bytes.extend(count.to_le_bytes());
        if let Some(version) = extension {
            bytes.extend(version.to_le_bytes());
            bytes.resize(bytes.len() + usize::from(count - 1) * 2, 0);
        }
        bytes
    }

    #[test]
    fn extension_supersedes_base_and_respects_counted_array_locations() {
        for (version, count) in [(0xd9, 0x6c), (0x101, 0x88), (0x10c, 0xa4), (0x112, 0xb7)] {
            let bytes = fib(0xc1, Some(version), [14, 22, count]);
            assert_eq!(effective_version(&bytes).unwrap(), version);
        }
        assert_eq!(
            effective_version(&fib(0xc1, None, [14, 22, 0x5d])).unwrap(),
            0xc1
        );
        // 2.5.15 requires skipping unknown counted-array tails, not using a
        // fixed offset inferred from FibBase.nFib.
        assert_eq!(
            effective_version(&fib(0xc1, Some(0x101), [15, 23, 0x89])).unwrap(),
            0x101
        );
        assert_eq!(
            effective_version(&fib(0xd9, None, [14, 22, 0x6c])).unwrap(),
            0xd9
        );
    }

    #[test]
    fn fixed_prefix_requires_canonical_word97_offsets_and_all_fc_lcb_values() {
        for count in [0x005d, 0x006c, 0x0088, 0x00a4, 0x00b7] {
            assert!(validate_fixed_prefix(&fib(0xc1, None, [14, 22, count])).is_ok());
        }
        for counts in [
            [13, 22, 0x005d],
            [15, 22, 0x005d],
            [14, 21, 0x005d],
            [14, 23, 0x005d],
            [14, 22, 0x005c],
        ] {
            assert!(validate_fixed_prefix(&fib(0xc1, None, counts)).is_err());
        }

        let canonical = fib(0xc1, None, [14, 22, 0x005d]);
        assert_eq!(FIXED_PREFIX_END, 0x0382);
        for end in 0..FIXED_PREFIX_END {
            assert!(
                validate_fixed_prefix(&canonical[..end.min(canonical.len())]).is_err(),
                "prefix {end}"
            );
        }
        assert!(validate_fixed_prefix(&canonical[..FIXED_PREFIX_END]).is_ok());
    }

    #[test]
    fn rejects_every_truncated_prefix_and_unknown_effective_version() {
        for bytes in [
            fib(0xc1, None, [14, 22, 0x5d]),
            fib(0xc1, Some(0x112), [14, 22, 0xb7]),
        ] {
            for end in 0..bytes.len() {
                assert!(effective_version(&bytes[..end]).is_err(), "prefix {end}");
            }
        }
        for version in [0, 0xa5, 0xc0, 0xc2, 0x113, 0xffff] {
            assert!(effective_version(&fib(version, None, [14, 22, 0x5d])).is_err());
            assert!(effective_version(&fib(0xc1, Some(version), [14, 22, 0x5d])).is_err());
        }
        assert!(effective_version(&fib(0xc1, Some(0xc1), [14, 22, 0x5d])).is_err());
    }
}
