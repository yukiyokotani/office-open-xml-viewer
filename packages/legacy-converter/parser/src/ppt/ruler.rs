//! Local text-body custom stops: MS-PPT 2.9.23-24, 2.9.29-30, 2.13.32.
//! Keep the ruler origin; do not apply TextPFException's conditional origin.
use super::*;

#[derive(Clone, Copy)]
pub(super) struct Tabs<'a> {
    entries: &'a [u8],
}

pub(super) fn read<'a>(atom: Record<'a>, budget: &mut usize) -> Result<Option<Tabs<'a>>, String> {
    if atom.kind != 4006 || atom.version != 0 || atom.instance != 0 {
        return Err(unsupported("invalid PowerPoint text ruler header"));
    }
    let data = atom.payload;
    let flags = u32_at(data, 0)?;
    // Presence bits differ from serialization order. Reserved bits are ignored.
    let mut pos = 4;
    for bit in [2, 1] {
        if flags & bit != 0 {
            u16_at(data, pos)?;
            pos += 2;
        }
    }
    let tabs = if flags & 4 != 0 {
        let count = usize::from(u16_at(data, pos)?);
        pos += 2;
        let end = pos + count * 4; // u16 count: arithmetic is bounded.
        let entries = data
            .get(pos..end)
            .ok_or_else(|| unsupported("truncated PowerPoint ruler tabs"))?;
        *budget = budget
            .checked_sub(count)
            .ok_or_else(|| unsupported("PowerPoint ruler tab work budget exceeded"))?;
        for entry in entries.chunks_exact(4) {
            if u16_at(entry, 2)? > 3 {
                return Err(unsupported("invalid PowerPoint ruler tab alignment"));
            }
        }
        pos = end;
        Some(Tabs { entries })
    } else {
        None
    };
    // Consume all five interleaved margin/indent pairs even though this path
    // emits only explicit local custom tabs. Other ruler properties and linked
    // master/default-ruler inheritance are separate, still-unsupported work.
    for level in 0..5 {
        for bit in [8 << level, 256 << level] {
            if flags & bit != 0 {
                u16_at(data, pos)?;
                pos += 2;
            }
        }
    }
    if pos != data.len() {
        return Err(unsupported("unexpected PowerPoint text ruler tail"));
    }
    Ok(tabs)
}

impl Tabs<'_> {
    pub(super) fn write(
        self,
        output: &mut String,
        xml: &mut usize,
        work: &mut usize,
    ) -> Result<(), String> {
        *work = work
            .checked_sub(self.entries.len() / 4)
            .ok_or_else(|| unsupported("PowerPoint ruler tab work budget exceeded"))?;
        drawing::append(output, xml, "<a:tabLst>")?;
        for entry in self.entries.chunks_exact(4) {
            let pos = master_to_emu(i64::from(i16::from_le_bytes([entry[0], entry[1]])));
            let alignment = ["l", "ctr", "r", "dec"][usize::from(entry[2])];
            // ECMA-376 21.1.2.2.13-14. Office roundtrip of the controlled
            // baseline maps a local ruler's 1152 units to pos=1828800, and
            // keeps that stop fixed when paragraph marL changes. No offset
            // correction, sorting or duplicate-position policy is invented.
            drawing::append(
                output,
                xml,
                &format!("<a:tab pos=\"{pos}\" algn=\"{alignment}\"/>"),
            )?;
        }
        drawing::append(output, xml, "</a:tabLst>")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn atom(payload: &[u8]) -> Record<'_> {
        Record {
            kind: 4006,
            version: 0,
            instance: 0,
            payload,
        }
    }
    #[test]
    fn absent_empty_signed_and_all_alignment_values_are_distinct() {
        assert!(read(atom(&0u32.to_le_bytes()), &mut 100).unwrap().is_none());
        let mut data = [4u32.to_le_bytes().as_slice(), &4u16.to_le_bytes()].concat();
        for (position, alignment) in [(-576i16, 0u16), (0, 1), (576, 2), (1152, 3)] {
            data.extend(position.to_le_bytes());
            data.extend(alignment.to_le_bytes());
        }
        let tabs = read(atom(&data), &mut 100).unwrap().unwrap();
        let mut result = String::new();
        tabs.write(&mut result, &mut 1000, &mut 100).unwrap();
        assert_eq!(result, "<a:tabLst><a:tab pos=\"-914400\" algn=\"l\"/><a:tab pos=\"0\" algn=\"ctr\"/><a:tab pos=\"914400\" algn=\"r\"/><a:tab pos=\"1828800\" algn=\"dec\"/></a:tabLst>");
        let empty = [4u32.to_le_bytes().as_slice(), &0u16.to_le_bytes()].concat();
        let mut result = String::new();
        read(atom(&empty), &mut 0)
            .unwrap()
            .unwrap()
            .write(&mut result, &mut 100, &mut 0)
            .unwrap();
        assert_eq!(result, "<a:tabLst></a:tabLst>");
    }
    #[test]
    fn every_presence_mask_consumes_fields_in_order_and_ignores_reserved_bits() {
        for flags in 0u32..8192 {
            let mut data = (flags | 0xffffe000).to_le_bytes().to_vec();
            for bit in [2, 1, 4, 8, 256, 16, 512, 32, 1024, 64, 2048, 128, 4096] {
                if flags & bit != 0 {
                    data.extend(0u16.to_le_bytes());
                }
            }
            assert_eq!(
                read(atom(&data), &mut 100).unwrap().is_some(),
                flags & 4 != 0
            );
            for end in 0..data.len() {
                assert!(read(atom(&data[..end]), &mut 100).is_err());
            }
            data.push(0);
            assert!(read(atom(&data), &mut 100).is_err());
        }
    }
    #[test]
    fn maximum_count_and_repeated_emission_are_budgeted() {
        let count = usize::from(u16::MAX);
        let mut data = [4u32.to_le_bytes().as_slice(), &u16::MAX.to_le_bytes()].concat();
        data.resize(6 + count * 4, 0);
        assert!(read(atom(&data), &mut (count - 1)).is_err());
        let mut read_work = count;
        let tabs = read(atom(&data), &mut read_work).unwrap().unwrap();
        assert_eq!(read_work, 0);
        let mut work = count;
        let mut xml = count * 40;
        let mut output = String::new();
        tabs.write(&mut output, &mut xml, &mut work).unwrap();
        assert_eq!(work, 0);
        assert_eq!(output.matches("<a:tab pos=").count(), count);
        let length = output.len();
        assert!(tabs.write(&mut output, &mut xml, &mut work).is_err());
        assert_eq!(output.len(), length);
    }
    #[test]
    fn malformed_headers_alignment_and_budgets_fail_closed() {
        let data = [
            4u32.to_le_bytes().as_slice(),
            &1u16.to_le_bytes(),
            &576i16.to_le_bytes(),
            &0u16.to_le_bytes(),
        ]
        .concat();
        for bad in [
            Record {
                version: 1,
                ..atom(&data)
            },
            Record {
                instance: 1,
                ..atom(&data)
            },
            Record {
                kind: 4011,
                ..atom(&data)
            },
        ] {
            assert!(read(bad, &mut 100).is_err());
        }
        assert!(read(atom(&data), &mut 0).is_err());
        let tabs = read(atom(&data), &mut 1).unwrap().unwrap();
        assert!(tabs.write(&mut String::new(), &mut 1000, &mut 0).is_err());
        assert!(tabs.write(&mut String::new(), &mut 10, &mut 1).is_err());
        for value in [4u16, 255, 256, 65535] {
            let mut invalid = data.clone();
            invalid[8..10].copy_from_slice(&value.to_le_bytes());
            assert!(read(atom(&invalid), &mut 100).is_err());
        }
    }
}
