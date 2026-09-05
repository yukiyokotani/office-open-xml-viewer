//! Word list-table decoder. [MS-DOC] 2.5.1 (FibRgFcLcb97), 2.9.131–133,
//! 2.9.147, 2.9.149–150 and 2.9.200–201. List identity and override identity
//! must stay separate: 2.4.6.4 counts by LSID, not by the paragraph's iLfo.

use super::{fkp, u16_at, u32_at, unsupported};
use std::collections::BTreeMap;

const FC_PLF_LST: usize = 0x2e2;
const FC_PLF_LFO: usize = 0x2ea;
// Resource policies, not MS-DOC limits. Payloads are borrowed, not copied per
// paragraph/reference. Charge before allocating record arrays or scanning text.
const MAX_OVERRIDES: usize = 65_536;
const MAX_LEVELS: usize = 65_536;
const MAX_METADATA_BYTES: usize = 16 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Reference {
    pub index: usize,
    pub level: u8,
    pub preserve_indent: bool,
}

impl Reference {
    /// MS-DOC 2.6.2 sprmPIlfo / sprmPIlvl. A removed/skipped paragraph must
    /// not accidentally advance a list counter. Validate before taking abs().
    pub fn new(ilfo: i16, level: u8) -> Result<Option<Self>, String> {
        if ilfo == 0 || ilfo == -2047 {
            return Ok(None);
        }
        if !(-2046..=2046).contains(&ilfo) {
            return Err(unsupported("invalid Word list reference"));
        }
        if level == 12 {
            return Ok(None);
        }
        if level > 8 {
            return Err(unsupported("invalid Word list level reference"));
        }
        Ok(Some(Self {
            index: ilfo.unsigned_abs() as usize - 1,
            level,
            preserve_indent: ilfo < 0,
        }))
    }
}

#[derive(Debug)]
pub struct Level<'a> {
    pub start: Option<u16>,
    pub format: u8,
    pub justification: u8,
    pub legal: bool,
    /// Zero-based first level that does NOT trigger a restart. None means no
    /// number sequence. For OOXML lvlRestart, the same integer is one-based
    /// last level that DOES trigger a restart; zero means never restart.
    pub restart: Option<u8>,
    pub follow: u8,
    pub tentative: bool,
    pub papx: &'a [u8],
    pub chpx: &'a [u8],
    /// Raw UTF-16LE, excluding Xst.cch. Symbol-font code points and literal
    /// percent characters must survive until the OOXML/text mapping stage.
    pub text: &'a [u8],
    /// (one-based UTF-16 offset, zero-based referenced level).
    pub placeholders: [Option<(u8, u8)>; 9],
}

#[derive(Debug)]
pub struct List<'a> {
    pub id: i32,
    pub styles: [u16; 9],
    pub simple: bool,
    pub hybrid: bool,
    pub auto_number: bool,
    pub levels: Vec<Level<'a>>,
}

#[derive(Debug)]
pub struct LevelOverride<'a> {
    pub index: u8,
    pub start: Option<u16>,
    pub formatting: Option<Level<'a>>,
}

#[derive(Debug)]
pub struct Override<'a> {
    pub list_index: usize,
    pub first_cp: Option<u32>,
    pub auto_number_field: Option<u8>,
    // Most LFOs have no overrides. Do not allocate nine potentially large LVL
    // slots for every header; retain only the records actually present.
    pub levels: Vec<LevelOverride<'a>>,
}

#[derive(Debug, Default)]
pub struct Tables<'a> {
    pub lists: Vec<List<'a>>,
    pub overrides: Vec<Override<'a>>,
}

pub struct Selection<'t, 'a> {
    pub list: &'t List<'a>,
    pub instance: &'t Override<'a>,
    pub level: &'t Level<'a>,
    pub start_override: Option<u16>,
}

impl<'a> Tables<'a> {
    pub fn read(word: &[u8], table: &'a [u8]) -> Result<Self, String> {
        let mut budget = Budget {
            levels: MAX_LEVELS,
            bytes: MAX_METADATA_BYTES,
        };
        let mut lists = Vec::new();
        let mut by_id = BTreeMap::new();
        let headers = fkp::table_part(word, table, FC_PLF_LST)?;
        if !headers.is_empty() {
            let count = u16_at(headers, 0)? as i16;
            if count < 0 || headers.len() != 2 + count as usize * 28 {
                return Err(unsupported("invalid Word list definition array"));
            }
            budget.bytes(headers.len())?;
            // lcbPlfLst covers only cLst + LSTFs. The variable-length LVLs
            // follow it and are deliberately OUTSIDE that declared byte range.
            // table_part already checked start + size against table.len().
            let levels_start = u32_at(word, FC_PLF_LST)? as usize + headers.len();
            let mut levels = Reader::new(&table[levels_start..]);
            for header in headers[2..].chunks_exact(28) {
                let id = u32_at(header, 0)? as i32;
                if id == -1 || by_id.insert(id, lists.len()).is_some() {
                    return Err(unsupported("invalid or duplicate Word list identifier"));
                }
                let mut styles = [0xfff; 9];
                for (i, style) in styles.iter_mut().enumerate() {
                    *style = u16_at(header, 8 + 2 * i)?;
                    if *style > 0xfff {
                        return Err(unsupported("invalid Word list paragraph style index"));
                    }
                }
                let simple = header[26] & 1 != 0;
                let hybrid = header[26] & 0x10 != 0;
                let mut definition = List {
                    id,
                    styles,
                    simple,
                    hybrid,
                    auto_number: header[26] & 4 != 0,
                    levels: Vec::new(),
                };
                for index in 0..if simple { 1 } else { 9 } {
                    definition
                        .levels
                        .push(read_level(&mut levels, index, hybrid, &mut budget)?);
                }
                lists.push(definition);
            }
        }
        let bytes = fkp::table_part(word, table, FC_PLF_LFO)?;
        let mut overrides = Vec::new();
        if !bytes.is_empty() {
            let mut data = Reader::new(bytes);
            let count = u32_at(data.take(4)?, 0)? as usize;
            // Every LFO has a 16-byte header and at least a 4-byte LFOData CP.
            // Test physical extent before reserving or iterating by its count.
            if count > MAX_OVERRIDES || count > data.bytes.len() / 20 {
                return Err(unsupported("invalid or excessive Word list override count"));
            }
            budget.bytes(4 + count * 20)?;
            let headers = data.take(count * 16)?;
            for header in headers.chunks_exact(16) {
                let id = u32_at(header, 0)? as i32;
                let list_index = *by_id.get(&id).ok_or_else(|| {
                    unsupported("Word list override references a missing definition")
                })?;
                let auto_number_field = match header[13] {
                    0 | 0xff => None,
                    0xfc..=0xfe => Some(header[13]),
                    _ => return Err(unsupported("invalid Word automatic-number field type")),
                };
                if auto_number_field.is_some() != lists[list_index].auto_number {
                    return Err(unsupported("inconsistent Word automatic-number list flags"));
                }
                if header[12] > 9 {
                    return Err(unsupported("excessive Word list level override count"));
                }
                let cp = u32_at(data.take(4)?, 0)?;
                let mut instance = Override {
                    list_index,
                    first_cp: if cp == u32::MAX { None } else { Some(cp) },
                    auto_number_field,
                    levels: Vec::new(),
                };
                let mut seen = 0_u16;
                for _ in 0..header[12] {
                    budget.bytes(8)?;
                    let header = data.take(8)?;
                    let index = header[4] & 15;
                    if index > 8 || seen & (1 << index) != 0 {
                        // A repeated level has no unambiguous selection under
                        // MS-DOC 2.4.6.3. Reject it; do not choose first/last.
                        return Err(unsupported("invalid or duplicate Word level override"));
                    }
                    seen |= 1 << index;
                    let formatting = if header[4] & 0x20 != 0 {
                        Some(read_level(
                            &mut data,
                            index,
                            lists[list_index].hybrid,
                            &mut budget,
                        )?)
                    } else {
                        None
                    };
                    let start = if header[4] & 0x10 == 0 {
                        None
                    } else if let Some(level) = &formatting {
                        // Both fStartAt and fFormatting: the 8-byte prefix's
                        // iStartAt is undefined; use the nested LVL instead.
                        level.start
                    } else {
                        Some(start_value(u32_at(header, 0)? as i32)?)
                    };
                    instance.levels.push(LevelOverride {
                        index,
                        start,
                        formatting,
                    });
                }
                overrides.push(instance);
            }
            if !data.bytes.is_empty() {
                return Err(unsupported("trailing Word list override data"));
            }
        }
        Ok(Self { lists, overrides })
    }

    pub fn resolve(&self, reference: Reference) -> Result<Selection<'_, 'a>, String> {
        if reference.level > 8 {
            return Err(unsupported("invalid Word list level reference"));
        }
        let instance = self
            .overrides
            .get(reference.index)
            .ok_or_else(|| unsupported("Word paragraph references a missing list override"))?;
        let list = &self.lists[instance.list_index];
        let level_override = instance.levels.iter().find(|o| o.index == reference.level);
        let level = level_override
            .and_then(|o| o.formatting.as_ref())
            .or_else(|| list.levels.get(usize::from(reference.level)))
            .ok_or_else(|| unsupported("Word paragraph references a missing list level"))?;
        Ok(Selection {
            list,
            instance,
            level,
            start_override: level_override.and_then(|o| o.start),
        })
    }
}

struct Reader<'a> {
    bytes: &'a [u8],
}
impl<'a> Reader<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes }
    }
    fn take(&mut self, size: usize) -> Result<&'a [u8], String> {
        let result = self
            .bytes
            .get(..size)
            .ok_or_else(|| unsupported("truncated Word list data"))?;
        self.bytes = &self.bytes[size..];
        Ok(result)
    }
}

struct Budget {
    levels: usize,
    bytes: usize,
}
impl Budget {
    fn bytes(&mut self, count: usize) -> Result<(), String> {
        self.bytes = self
            .bytes
            .checked_sub(count)
            .ok_or_else(|| unsupported("Word list metadata byte budget exceeded"))?;
        Ok(())
    }
}

fn start_value(value: i32) -> Result<u16, String> {
    if !(0..=32767).contains(&value) {
        return Err(unsupported("invalid Word list start value"));
    }
    Ok(value as u16)
}

fn read_level<'a>(
    data: &mut Reader<'a>,
    index: u8,
    hybrid: bool,
    budget: &mut Budget,
) -> Result<Level<'a>, String> {
    budget.levels = budget
        .levels
        .checked_sub(1)
        .ok_or_else(|| unsupported("Word list level budget exceeded"))?;
    budget.bytes(30)?; // LVLF + Xst.cch, before touching any variable payload.
    let header = data.take(28)?;
    let format = header[4];
    if !matches!(format, 0..=59 | 0xff) || matches!(format, 8 | 9 | 15 | 19) {
        return Err(unsupported("invalid Word list number format"));
    }
    let numbered = !matches!(format, 0x17 | 0xff);
    let start = if numbered {
        Some(start_value(u32_at(header, 0)? as i32)?)
    } else {
        None
    };
    let justification = header[5] & 3;
    let follow = header[15];
    if justification > 2 || follow > 2 {
        return Err(unsupported("invalid Word list justification or suffix"));
    }
    let restart = if !numbered {
        None
    } else if header[5] & 8 == 0 {
        Some(index)
    } else {
        if header[26] > index {
            return Err(unsupported("invalid Word list restart limit"));
        }
        Some(header[26])
    };
    budget.bytes(usize::from(header[24]) + usize::from(header[25]))?;
    let papx = data.take(usize::from(header[25]))?;
    let chpx = data.take(usize::from(header[24]))?;
    let length = usize::from(u16_at(data.take(2)?, 0)?) * 2;
    budget.bytes(length)?;
    let text = data.take(length)?;
    let mut placeholders = [None; 9];
    let mut previous = 0;
    for (i, offset) in header[6..15]
        .iter()
        .copied()
        .take_while(|v| *v != 0)
        .enumerate()
    {
        if i > usize::from(index) || offset <= previous {
            return Err(unsupported("invalid Word list placeholder offsets"));
        }
        let value = u16_at(text, 2 * (usize::from(offset) - 1))?;
        if value > u16::from(index) {
            return Err(unsupported(
                "Word list placeholder references a deeper level",
            ));
        }
        placeholders[i] = Some((offset, value as u8));
        previous = offset;
    }
    if format == 0x17 && (length != 2 || placeholders[0].is_some()) {
        return Err(unsupported("invalid Word bullet text"));
    }
    Ok(Level {
        start,
        format,
        justification,
        legal: header[5] & 4 != 0,
        restart,
        follow,
        tentative: hybrid && header[5] & 0x80 != 0,
        papx,
        chpx,
        text,
        placeholders,
    })
}

#[cfg(test)]
mod tests;
