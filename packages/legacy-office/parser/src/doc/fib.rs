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

/// FIB selector and independently acquired DOP ownership from counted bytes.
/// Numeric nFib alone is insufficient: MS-DOC 2.7.1 selects Dop97 when
/// cswNew is zero, including a tolerated higher base selector. Physical
/// extension bounds and the documented legacy reading family are checked
/// before creation.
#[derive(Clone, Copy)]
pub(super) struct KnownFibSchema {
    version: u16,
    extension_words: usize,
}

impl KnownFibSchema {
    pub(super) fn effective_version(self) -> u16 {
        self.version
    }

    /// The legacy reading family needs the whole documented Dop97 prefix.
    pub(super) fn requires_legacy_dop97_prefix(self) -> bool {
        (0x00c3..0x00d9).contains(&self.version)
    }

    /// Modern enumerated extension selectors embed Dop2000. The legacy
    /// pre-2000 reading family owns Dop97 even with a counted extension.
    /// Physical DOP length alone never grants later field ownership.
    pub(super) fn has_dop2000_fields(self) -> bool {
        self.extension_words != 0 && self.version >= 0x00d9
    }
}

pub(super) fn known_schema(word: &[u8]) -> Result<KnownFibSchema, String> {
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
        // MS-DOC 2.5.2 product behavior note 11 assigns these legacy
        // FibBase selectors Word 97 semantics. A present nFibNew remains
        // authoritative (2.5.14); this alias never applies to the extension.
        match u16_at(word, 2)? {
            0x00c0 | 0x00c2 => 0x00c1,
            base => base,
        }
    } else {
        u16_at(&word[start..end], 0)?
    };
    // Microsoft Office Word 97-2007 Binary File Format Specification (2007),
    // "File Information Block": readers skip future counted-array entries;
    // rgswNew begins with the actual nFib. "Document Properties" defines the
    // pre-2000 DOP family for 105 < nFib < 217. This is a bounded reading
    // profile, not a new MS-DOC schema or an alias to C1/D9. Retain the actual
    // selector and all extra bytes in the original stream; never interpret
    // this family's extension/FC-LCB tails as modern fields. Only the canonical
    // Word97 prefix with a compatible base/back-version is supported here.
    let word97_prefix = count != 0 && u16_at(word, 2)? == 0x00c1 && u16_at(word, 12)? <= 0x00c1;
    let legacy_forward = word97_prefix && (0x00c3..0x00d9).contains(&version);
    // An nFibNew that names one of the Word 97 selectors of MS-DOC 2.5.2
    // product behavior note 11 is outside the nFibNew values of 2.5.14, but
    // the base, back-version and counted prefix are Word 97's and Word opens
    // such a file. Read it with the Word 97 profile; its extension carries no
    // later field this profile would interpret.
    if word97_prefix && matches!(version, 0x00c0..=0x00c2) {
        validate_fixed_prefix(word)?;
        return Ok(KnownFibSchema {
            version: crate::rejection::approximated("unsupported Word FIB version", 0x00c1),
            extension_words: count,
        });
    }
    if legacy_forward {
        validate_fixed_prefix(word)?;
    }
    if (!matches!(version, 0x00c1 | 0x00d9 | 0x0101 | 0x010c | 0x0112) && !legacy_forward)
        || (count != 0 && version == 0x00c1)
    {
        return Err(crate::rejection::out_of_scope_version(
            "unsupported Word FIB version",
        ));
    }
    Ok(KnownFibSchema {
        version,
        extension_words: count,
    })
}

#[cfg(test)]
fn effective_version(word: &[u8]) -> Result<u16, String> {
    known_schema(word).map(KnownFibSchema::effective_version)
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
    fn documented_word97_base_aliases_have_word97_semantics() {
        // MS-DOC 2.5.2 product behavior note 11 explicitly assigns these
        // FibBase selectors the semantics of Word 97; it is not a numeric range.
        for base in [0x00c0, 0x00c2] {
            let bytes = fib(base, None, [14, 22, 0x5d]);
            validate_fixed_prefix(&bytes).unwrap();
            assert_eq!(effective_version(&bytes).unwrap(), 0x00c1);
            for end in 0..bytes.len() {
                assert!(
                    effective_version(&bytes[..end]).is_err(),
                    "base {base:x}, prefix {end}"
                );
            }
        }
    }

    #[test]
    fn base_aliases_cannot_override_or_become_extension_selectors() {
        for base in [0x00c0, 0x00c2] {
            let known = fib(base, Some(0x0101), [14, 22, 0x88]);
            assert_eq!(effective_version(&known).unwrap(), 0x0101);
            for extension in [0x00c0, 0x00c1, 0x00c2, 0x0120] {
                assert!(effective_version(&fib(base, Some(extension), [14, 22, 0x88])).is_err());
            }
            assert!(effective_version(&known[..known.len() - 1]).is_err());
            assert!(validate_fixed_prefix(&fib(base, None, [15, 22, 0x5d])).is_err());
        }
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
        for version in [0, 0xa5, 0x113, 0xffff] {
            assert!(effective_version(&fib(version, None, [14, 22, 0x5d])).is_err());
            assert!(effective_version(&fib(0xc1, Some(version), [14, 22, 0x5d])).is_err());
        }
        // A Word 97 selector in the extension keeps the Word 97 profile.
        assert_eq!(
            effective_version(&fib(0xc1, Some(0xc1), [14, 22, 0x5d])).unwrap(),
            0xc1
        );
    }

    #[test]
    fn legacy_forward_selector_retains_prefix_without_modern_dop_ownership() {
        for selector in [0x00c3, 0x00d8] {
            let mut word = fib(0x00c1, Some(selector), [14, 22, 0x006c]);
            let offset = word.len() - 6;
            word[offset..offset + 2].copy_from_slice(&4u16.to_le_bytes());
            word.extend_from_slice(&[0xe5, 0x00, 0xff, 0xff]);
            let schema = known_schema(&word).unwrap();
            assert_eq!(schema.effective_version(), selector);
            assert!(!schema.has_dop2000_fields());
            for end in 0..word.len() {
                assert!(known_schema(&word[..end]).is_err());
            }
            word[12..14].copy_from_slice(&0x00d9u16.to_le_bytes());
            assert!(known_schema(&word).is_err());
            word[12..14].fill(0);
            word[2..4].copy_from_slice(&0x00c3u16.to_le_bytes());
            assert!(known_schema(&word).is_err());
        }
        for selector in [0x00da, 0x0100] {
            assert!(known_schema(&fib(0x00c1, Some(selector), [14, 22, 0x006c])).is_err());
        }
    }

    #[test]
    fn word97_selector_in_the_extension_reads_with_the_word97_profile() {
        crate::rejection::take_best_effort();
        for selector in [0x00c0, 0x00c1, 0x00c2] {
            let mut word = fib(0x00c1, Some(selector), [14, 22, 0x006c]);
            let offset = word.len() - 6;
            word[offset..offset + 2].copy_from_slice(&4u16.to_le_bytes());
            word.extend_from_slice(&[0xe5, 0x00, 0xff, 0xff]);
            word[12..14].copy_from_slice(&0x00bfu16.to_le_bytes());
            let schema = known_schema(&word).unwrap();
            assert_eq!(schema.effective_version(), 0x00c1);
            assert!(!schema.has_dop2000_fields());
            assert!(!schema.requires_legacy_dop97_prefix());
            assert_eq!(
                crate::rejection::take_best_effort(),
                [(
                    crate::rejection::BestEffort::Approximated,
                    "unsupported Word FIB version"
                )]
            );
            for end in 0..word.len() {
                assert!(known_schema(&word[..end]).is_err());
            }
            // A later back-version or a non-Word 97 base is not that profile.
            word[12..14].copy_from_slice(&0x00d9u16.to_le_bytes());
            assert!(known_schema(&word).is_err());
            word[12..14].fill(0);
            word[2..4].copy_from_slice(&0x00d9u16.to_le_bytes());
            assert!(known_schema(&word).is_err());
        }
        crate::rejection::take_best_effort();
    }

    #[test]
    fn fib_dop_owner_known_extension_retains_minimum_extra_and_ignored_tail_acquisition() {
        for (version, words) in [(0x00d9u16, 1u16), (0x00d9, 5), (0x0112, 1), (0x0112, 5)] {
            let mut word = fib(0x00c1, None, [14, 22, 0x00b7]);
            let offset = word.len() - 2;
            word[offset..].copy_from_slice(&words.to_le_bytes());
            word.extend_from_slice(&version.to_le_bytes());
            word.resize(word.len() + usize::from(words - 1) * 2, 0xff);
            // Keep quick-save in its published domain; only ignored tails vary.
            if words > 1 {
                word[offset + 4..offset + 6].fill(0);
            }
            assert_eq!(effective_version(&word).unwrap(), version);
            assert!(effective_version(&word[..word.len() - 1]).is_err());
        }
    }
}
