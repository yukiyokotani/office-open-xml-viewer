//! Passive, shape-local PP9 numbering. MS-PPT 2.7.14–18, 2.9.26–27/67–68.
//! Outline/document/master PP9 inheritance is deliberately not inferred here.
use super::*;

/// Follow only the exact ClientData -> ProgTags -> ProgBinaryTag -> ___PPT9
/// ownership chain. Other tags, actions, links and nested decoys are opaque.
pub(in crate::ppt) fn local_atom<'a>(
    tags: Record<'a>,
    budget: &mut usize,
) -> Result<Option<&'a [u8]>, String> {
    if tags.version != 15 {
        return Err(unsupported("invalid PowerPoint shape tags"));
    }
    let mut found = None;
    for tag in parse_records(tags.payload, budget)? {
        if tag.kind != 5002 {
            continue;
        }
        if tag.version != 15 || tag.instance != 0 {
            return Err(unsupported("invalid PowerPoint binary tag container"));
        }
        let pair = parse_records(tag.payload, budget)?;
        let Some(name) = pair.first() else { continue };
        if name.kind != 4026 || name.payload != b"_\0_\0_\0P\0P\0T\09\0" {
            continue;
        }
        if name.version != 0 || name.instance != 0 || pair.len() != 2 || found.is_some() {
            return Err(unsupported("ambiguous PowerPoint PP9 shape tag"));
        }
        let blob = pair[1];
        if blob.kind != 5003 || blob.version != 0 || blob.instance != 0 {
            return Err(unsupported("invalid PowerPoint PP9 shape blob"));
        }
        let atoms = parse_records(blob.payload, budget)?;
        if atoms.len() != 1
            || atoms[0].kind != 4012
            || atoms[0].version != 0
            || atoms[0].instance != 0
        {
            return Err(unsupported("invalid PowerPoint local PP9 text style"));
        }
        found = Some(atoms[0].payload);
    }
    Ok(found)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::ppt) struct Number {
    pub scheme: &'static str,
    pub start: u16,
}

// MS-PPT 2.13.28 -> ECMA-376 ST_TextAutonumberScheme. Enumeration values
// are explicit identities, not conversions from a font's fallback glyph.
const SCHEMES: [&str; 41] = [
    "alphaLcPeriod",
    "alphaUcPeriod",
    "arabicParenR",
    "arabicPeriod",
    "romanLcParenBoth",
    "romanLcParenR",
    "romanLcPeriod",
    "romanUcPeriod",
    "alphaLcParenBoth",
    "alphaLcParenR",
    "alphaUcParenBoth",
    "alphaUcParenR",
    "arabicParenBoth",
    "arabicPlain",
    "romanUcParenBoth",
    "romanUcParenR",
    "ea1ChsPlain",
    "ea1ChsPeriod",
    "circleNumDbPlain",
    "circleNumWdWhitePlain",
    "circleNumWdBlackPlain",
    "ea1ChtPlain",
    "ea1ChtPeriod",
    "arabic1Minus",
    "arabic2Minus",
    "hebrew2Minus",
    "ea1JpnKorPlain",
    "ea1JpnKorPeriod",
    "arabicDbPlain",
    "arabicDbPeriod",
    "thaiAlphaPeriod",
    "thaiAlphaParenR",
    "thaiAlphaParenBoth",
    "thaiNumPeriod",
    "thaiNumParenR",
    "thaiNumParenBoth",
    "hindiAlphaPeriod",
    "hindiNumPeriod",
    "ea1JpnChsDbPeriod",
    "hindiNumParenR",
    "hindiAlpha1Period",
];

fn read(r: &mut Reader<'_, '_>) -> Result<Option<Number>, String> {
    let pf = r.u32()?;
    // PFMasks reserved/unused bits have no payload; ignore them, but reject
    // base-exception fields in this structurally different extension.
    if pf & 0x003f_fdff != 0 {
        return Err(unsupported("base PowerPoint paragraph fields in PP9 style"));
    }
    let picture = r.optional16(pf, 0x00800000)?;
    let enabled = r.optional16(pf, 0x02000000)?;
    if enabled.is_some_and(|n| n > 1) {
        return Err(unsupported("invalid PowerPoint automatic numbering flag"));
    }
    let number = if pf & 0x01000000 != 0 {
        let scheme = r.u16()?;
        let start = r.u16()?;
        if !(1..=32767).contains(&start) {
            return Err(unsupported("invalid PowerPoint automatic numbering start"));
        }
        Some(Number {
            scheme: *SCHEMES
                .get(usize::from(scheme))
                .ok_or_else(|| unsupported("invalid PowerPoint automatic numbering scheme"))?,
            start,
        })
    } else {
        None
    };
    // TextCFException9 has only the optional four-byte pp10runid word.
    let cf = r.u32()?;
    if cf & 0x07ef_3eb7 != 0 {
        return Err(unsupported("invalid PowerPoint PP9 character mask"));
    }
    if cf & 0x00100000 != 0 {
        r.u32()?;
    }
    // TextSIException, constrained by StyleTextProp9: no spell/lang/smartTags.
    let si = r.u32()?;
    if si & 0x207 != 0 {
        return Err(unsupported("invalid PowerPoint PP9 text information mask"));
    }
    if si & 0x40 != 0 && r.u16()? > 1 {
        return Err(unsupported("invalid PowerPoint PP9 bidirectional flag"));
    }
    if si & 0x20 != 0 {
        r.u32()?;
    }
    // No guessed default scheme/start, no precedence guess for a real picture
    // bullet (BlipRef -1 means no picture, MS-PPT 2.2.1).
    Ok(number.filter(|_| enabled == Some(1) && picture.is_none_or(|v| v == 65535)))
}

/// Linear merge of extension entries with consecutive character-run groups.
/// MS-PPT 2.9.67 requires skipping an entry whose index modulo 16 does not
/// match the *next* group's pp9rt. The raw value defaults to zero, never an
/// inherited master's run ID. Returned ends are original UTF-16 positions.
pub(super) fn bind(
    bytes: &[u8],
    characters: &[(usize, Character)],
    budget: &mut usize,
) -> Result<Vec<(usize, Option<Number>)>, String> {
    let mut reader = Reader {
        bytes,
        pos: 0,
        budget,
    };
    let mut groups = Vec::new();
    let (mut ci, mut index) = (0, 0usize);
    while reader.pos < bytes.len() {
        *reader.budget = reader
            .budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("PowerPoint PP9 text work budget exceeded"))?;
        let value = read(&mut reader)?;
        if ci < characters.len() && index % 16 == usize::from((characters[ci].1.style >> 10) & 15) {
            let id = (characters[ci].1.style >> 10) & 15;
            let mut end = characters[ci].0;
            ci += 1;
            while ci < characters.len() && (characters[ci].1.style >> 10) & 15 == id {
                end = characters[ci].0;
                ci += 1;
            }
            groups.push((end, value));
        }
        index += 1;
    }
    if let Some((end, _)) = characters.last() {
        if ci < characters.len() {
            groups.push((*end, None));
        }
    }
    Ok(groups)
}

/// Only project a paragraph-wide, uniform explicit choice. Conflicting PP9
/// character sequences inside a paragraph are outside this supported subset.
pub(super) fn paragraph(
    groups: &[(usize, Option<Number>)],
    cursor: &mut usize,
    start: usize,
    end: usize,
) -> Option<Number> {
    while *cursor < groups.len() && groups[*cursor].0 <= start {
        *cursor += 1;
    }
    let value = groups.get(*cursor)?.1;
    let mut uniform = true;
    while groups[*cursor].0 < end {
        *cursor += 1;
        uniform &= groups.get(*cursor)?.1 == value;
    }
    if uniform {
        value
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn entry(scheme: u16, start: u16) -> Vec<u8> {
        [
            0x03800000u32.to_le_bytes().to_vec(),
            65535u16.to_le_bytes().to_vec(),
            1u16.to_le_bytes().to_vec(),
            scheme.to_le_bytes().to_vec(),
            start.to_le_bytes().to_vec(),
            vec![0; 8],
        ]
        .concat()
    }
    fn chars(runs: &[(usize, u16)]) -> Vec<(usize, Character)> {
        runs.iter()
            .map(|&(end, id)| {
                let mut c = Level::empty(0).character;
                c.style = id << 10;
                (end, c)
            })
            .collect()
    }
    #[test]
    fn local_tags_enforce_ownership_and_ignore_foreign_payloads() {
        fn record(kind: u16, version: u16, payload: &[u8]) -> Vec<u8> {
            [
                version.to_le_bytes().to_vec(),
                kind.to_le_bytes().to_vec(),
                (payload.len() as u32).to_le_bytes().to_vec(),
                payload.to_vec(),
            ]
            .concat()
        }
        let name = record(4026, 0, b"_\0_\0_\0P\0P\0T\09\0");
        let data = record(5003, 0, &record(4012, 0, &entry(3, 1)));
        let tag = record(5002, 15, &[name.clone(), data.clone()].concat());
        let parse = |bytes: &[u8], budget: &mut usize| {
            local_atom(
                Record {
                    kind: 5000,
                    version: 15,
                    instance: 0,
                    payload: bytes,
                },
                budget,
            )
            .map(|v| v.map(<[u8]>::to_vec))
        };
        assert_eq!(parse(&tag, &mut 100).unwrap(), Some(entry(3, 1)));
        assert!(parse(&tag, &mut 1).is_err());
        for bytes in [
            [tag.clone(), tag.clone()].concat(),
            record(5002, 15, &name),
            record(
                5002,
                15,
                &[
                    name.clone(),
                    record(5003, 0, &record(4012, 16, &entry(3, 1))),
                ]
                .concat(),
            ),
            record(
                5002,
                15,
                &[
                    name.clone(),
                    record(
                        5003,
                        0,
                        &[record(4012, 0, &entry(3, 1)), record(4012, 0, &entry(3, 1))].concat(),
                    ),
                ]
                .concat(),
            ),
        ] {
            assert!(parse(&bytes, &mut 100).is_err());
        }
        // A foreign tag cannot masquerade as a local extension. Nor can an
        // action container or extra nesting make embedded marker data visible.
        for bytes in [
            record(
                5002,
                15,
                &[record(4026, 0, b"other"), record(5003, 0, &[255])].concat(),
            ),
            record(4082, 15, &tag),
            record(5000, 15, &tag),
        ] {
            assert_eq!(parse(&bytes, &mut 100).unwrap(), None);
        }
    }
    #[test]
    fn boundaries_and_all_explicit_numbering_schemes() {
        for scheme in 0..41 {
            for start in [1, 32767] {
                let result = bind(&entry(scheme, start), &chars(&[(2, 0)]), &mut 1).unwrap();
                assert_eq!(result[0].1.unwrap().scheme, SCHEMES[usize::from(scheme)]);
                assert_eq!(result[0].1.unwrap().start, start);
            }
        }
        for (scheme, start) in [(41, 1), (65535, 1), (3, 0), (3, 32768), (3, 65535)] {
            assert!(bind(&entry(scheme, start), &chars(&[(2, 0)]), &mut 1).is_err());
        }
    }
    #[test]
    fn extension_run_ids_wrap_and_contiguous_equal_ids_share_one_entry() {
        let records = (0..19).flat_map(|i| entry(3, i + 1)).collect::<Vec<_>>();
        let mut runs = (0..18)
            .map(|i| ((i + 1) * 2, (i % 16) as u16))
            .collect::<Vec<_>>();
        runs.insert(1, (3, 0));
        let output = bind(&records, &chars(&runs), &mut 19).unwrap();
        assert_eq!(output.len(), 18);
        assert_eq!(
            output[0],
            (
                3,
                Some(Number {
                    scheme: "arabicPeriod",
                    start: 1
                })
            )
        );
        assert_eq!(output[17].1.unwrap().start, 18);
        // Matching is to the NEXT group, not random indexing by the current ID.
        let output = bind(&records, &chars(&[(2, 3), (4, 1)]), &mut 19).unwrap();
        assert_eq!(output[0].1.unwrap().start, 4);
        assert_eq!(output[1].1.unwrap().start, 18);
        assert!(bind(&records, &chars(&runs), &mut 18)
            .unwrap_err()
            .contains("budget"));
    }
    #[test]
    fn paragraph_requires_uniform_choice_including_its_terminator() {
        let one = Some(Number {
            scheme: "arabicPeriod",
            start: 1,
        });
        let two = Some(Number {
            scheme: "arabicPeriod",
            start: 2,
        });
        for differing in [two, None] {
            let groups = [(1, one), (2, differing), (4, two)];
            let mut cursor = 0;
            assert_eq!(paragraph(&groups, &mut cursor, 0, 2), None);
            assert_eq!(paragraph(&groups, &mut cursor, 2, 4), two);
        }
        assert_eq!(paragraph(&[(1, one), (4, one)], &mut 0, 0, 2), one);
        assert_eq!(paragraph(&[], &mut 0, 0, 2), None);
    }
    #[test]
    fn no_default_or_picture_precedence_is_invented() {
        let characters = chars(&[(2, 0)]);
        for picture in [0, 1, 32767] {
            let mut data = entry(3, 1);
            data[4..6].copy_from_slice(&(picture as u16).to_le_bytes());
            assert_eq!(bind(&data, &characters, &mut 1).unwrap()[0].1, None);
        }
        for enabled in [0, 2, 65535] {
            let mut data = entry(3, 1);
            data[6..8].copy_from_slice(&(enabled as u16).to_le_bytes());
            let result = bind(&data, &characters, &mut 1);
            if enabled == 0 {
                assert_eq!(result.unwrap()[0].1, None)
            } else {
                assert!(result.is_err())
            }
        }
        let flag_only = [
            0x02000000u32.to_le_bytes().to_vec(),
            1u16.to_le_bytes().to_vec(),
            vec![0; 8],
        ]
        .concat();
        let scheme_only = [
            0x01000000u32.to_le_bytes().to_vec(),
            vec![3, 0, 1, 0],
            vec![0; 8],
        ]
        .concat();
        for data in [flag_only, scheme_only, vec![0; 12]] {
            assert_eq!(bind(&data, &characters, &mut 1).unwrap()[0].1, None);
        }
    }
    #[test]
    fn validates_all_entries_including_unmatched_and_handles_optional_words() {
        let mut data = entry(3, 1);
        data.truncate(12);
        data.extend(0x00100000u32.to_le_bytes());
        data.extend(u32::MAX.to_le_bytes()); // pp10runid plus ignored unused bits
        data.extend(0x60u32.to_le_bytes()); // bidi plus pp11 group word
        data.extend(1u16.to_le_bytes());
        data.extend(0xfedcba98u32.to_le_bytes());
        assert!(bind(&data, &chars(&[(2, 0)]), &mut 1).unwrap()[0]
            .1
            .is_some());
        for cut in 1..data.len() {
            assert!(
                bind(&data[..cut], &chars(&[(2, 15)]), &mut 1).is_err(),
                "cut {cut}"
            );
        }
        for (offset, bad_mask) in [(0, 1u32), (12, 1u32), (20, 0x200u32)] {
            let mut bad = data.clone();
            bad[offset..offset + 4].copy_from_slice(&bad_mask.to_le_bytes());
            assert!(bind(&bad, &chars(&[(2, 15)]), &mut 1).is_err());
        }
    }
}
