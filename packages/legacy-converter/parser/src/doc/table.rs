//! Binary table properties, [MS-DOC] 2.4.3, 2.6.3, TDefTableOperand/TC80.
//! A row's definition belongs to its TTP mark, not its first text paragraph.
use super::{border::Border, u16_at, u32_at, unsupported};

#[derive(Clone, Default)]
pub struct Cell {
    pub width: i32,
    pub flags: u16,
    pub preferred: u16,
    pub margins: [Option<u16>; 4],
    pub borders: [Option<Border>; 6],
}

#[derive(Default)]
pub struct Properties {
    pub in_table: bool,
    depth: Option<i32>,
    pub row_end: bool,
    pub inner_cell: bool,
    pub inner_row: bool,
    pub row: Row,
}

#[derive(Clone)]
pub struct Row {
    pub identity: std::collections::BTreeMap<u16, Vec<u8>>,
    pub cells: Vec<Cell>,
    pub left: i32,
    pub gap: i32,
    left_is_edge: bool,
    pub height: i32,
    pub margins: [u16; 4],
    pub autofit: bool,
    pub header: bool,
    pub cant_split: bool,
    pub bidi: bool,
    pub alignment: (u16, bool),
    pub borders: [Option<Border>; 6],
}

impl Default for Row {
    fn default() -> Self {
        Self {
            identity: Default::default(),
            cells: vec![],
            left: 0,
            gap: 0,
            left_is_edge: false,
            height: 0,
            margins: [0, 108, 0, 108],
            autofit: false,
            header: false,
            cant_split: false,
            bidi: false,
            alignment: (0, false),
            borders: Default::default(),
        }
    }
}

fn signed(b: &[u8]) -> Result<i32, String> {
    Ok(u16_at(b, 0)? as i16 as i32)
}
fn nonnegative(b: &[u8]) -> Result<i32, String> {
    let value = signed(b)?;
    if !(0..=31680).contains(&value) {
        return Err(unsupported("invalid Word table width"));
    }
    Ok(value)
}
fn boolean(b: u8) -> Result<bool, String> {
    match b {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(unsupported("invalid Word table boolean")),
    }
}
fn range(bytes: &[u8], len: usize) -> Result<std::ops::Range<usize>, String> {
    let first = *bytes
        .first()
        .ok_or_else(|| unsupported("short Word cell range"))? as usize;
    let last = *bytes
        .get(1)
        .ok_or_else(|| unsupported("short Word cell range"))? as usize;
    if first > last || last > len {
        return Err(unsupported("Word cell range outside row"));
    }
    Ok(first..last)
}

pub fn prm0(prm: u16) -> Option<[u8; 3]> {
    let code: u16 = match (prm >> 1) & 127 {
        0x18 => 0x2416,
        0x19 => 0x2417,
        _ => return None,
    };
    let [a, b] = code.to_le_bytes();
    Some([a, b, (prm >> 8) as u8])
}

impl Properties {
    pub fn depth(&self) -> Result<usize, String> {
        let n = self.depth.unwrap_or(i32::from(self.in_table));
        // Resource policy independent of the file's representable table depth.
        if !(0..=32).contains(&n) {
            return Err(unsupported("Word table nesting budget exceeded"));
        }
        Ok(n as usize)
    }

    pub fn apply(&mut self, code: u16, b: &[u8]) -> Result<bool, String> {
        match code {
            0x2416 => self.in_table = boolean(b[0])?,
            0x2417 => self.row_end = boolean(b[0])?,
            0x244b => self.inner_cell = boolean(b[0])?,
            0x244c => self.inner_row = boolean(b[0])?,
            0x6649 => self.depth = Some(u32_at(b, 0)? as i32),
            0x664a => {
                self.depth = Some(
                    self.depth
                        .unwrap_or(0)
                        .checked_add(u32_at(b, 0)? as i32)
                        .ok_or_else(|| unsupported("Word table depth overflow"))?,
                )
            }
            _ => return self.row.apply(code, b),
        }
        Ok(true)
    }
}

impl Row {
    pub fn origin(&self) -> i32 {
        // TDefTable boundaries already include all outer cell spacing. TDxaLeft
        // instead defines the origin before TDxaGapHalf is subtracted.
        if self.left_is_edge {
            self.left
        } else {
            self.left - self.gap
        }
    }
    pub fn apply(&mut self, code: u16, b: &[u8]) -> Result<bool, String> {
        if matches!(
            code,
            0x7469 | 0x563a | 0x360d | 0x3465 | 0x940e | 0x940f | 0x9410 | 0x9411 | 0x941e | 0x941f
        ) {
            self.identity.insert(code, b.to_vec());
        }
        match code {
            0xd608 => {
                let n = *b
                    .get(2)
                    .ok_or_else(|| unsupported("short Word table definition"))?
                    as usize;
                if n > 63 {
                    return Err(unsupported("too many Word table cells"));
                }
                let end = 3 + (n + 1) * 2;
                let boundaries = b
                    .get(3..end)
                    .ok_or_else(|| unsupported("short Word table boundaries"))?;
                self.left = signed(boundaries)?;
                self.left_is_edge = true;
                let mut cells = Vec::with_capacity(n);
                for i in 0..n {
                    let width = signed(&boundaries[(i + 1) * 2..])? - signed(&boundaries[i * 2..])?;
                    if width < 0 {
                        return Err(unsupported("unordered Word table boundaries"));
                    }
                    let mut cell = Cell {
                        width,
                        ..Cell::default()
                    };
                    if let Some(tc) = b.get(end + i * 20..end + (i + 1) * 20) {
                        cell.flags = u16_at(tc, 0)?;
                        cell.preferred = u16_at(tc, 2)?;
                        for s in 0..4 {
                            cell.borders[s] = Some(Border::read(&tc[4 + s * 4..], true)?);
                        }
                    }
                    cells.push(cell);
                }
                self.cells = cells;
            }
            0x7621 => {
                let first = b[0] as usize;
                let count = b[1] as usize;
                let width = nonnegative(&b[2..])?;
                if count == 0
                    || first > self.cells.len()
                    || self.cells.len() + count > 63
                    || self.cells.iter().map(|c| c.width).sum::<i32>() + width * count as i32
                        > 31680
                {
                    return Err(unsupported("invalid Word cell insertion"));
                }
                self.cells.splice(
                    first..first,
                    (0..count).map(|_| Cell {
                        width,
                        ..Cell::default()
                    }),
                );
            }
            0x5622 => {
                let r = range(b, self.cells.len())?;
                if r.len() == self.cells.len() {
                    return Err(unsupported("Word row cannot delete every cell"));
                }
                self.cells.drain(r);
            }
            0x7623 => {
                let r = range(b, self.cells.len())?;
                let width = nonnegative(&b[2..])?;
                for c in &mut self.cells[r] {
                    c.width = width;
                }
            }
            0x5624 | 0x5625 => {
                let r = range(b, self.cells.len())?;
                for (i, c) in self.cells[r].iter_mut().enumerate() {
                    let merge = if code == 0x5625 {
                        0
                    } else if i == 0 {
                        2
                    } else {
                        1
                    };
                    c.flags = (c.flags & !3) | merge;
                }
            }
            0xd62b if b[0] == 2 => {
                if ![0, 1, 3].contains(&b[2]) {
                    return Err(unsupported("invalid Word vertical merge"));
                }
                let cell = self
                    .cells
                    .get_mut(b[1] as usize)
                    .ok_or_else(|| unsupported("Word vertical merge outside row"))?;
                cell.flags = (cell.flags & !(3 << 5)) | ((b[2] as u16) << 5);
            }
            0xd62c if b[0] == 3 => {
                if b[3] > 2 {
                    return Err(unsupported("invalid Word cell alignment"));
                }
                let r = range(&b[1..], self.cells.len())?;
                for c in &mut self.cells[r] {
                    c.flags = (c.flags & !(3 << 7)) | ((b[3] as u16) << 7);
                }
            }
            0x9601 => {
                self.left = signed(b)?;
                self.left_is_edge = false;
            }
            0xd605 | 0xd613 => {
                let old = code == 0xd605;
                let size = if old { 4 } else { 8 };
                if b[0] as usize != 6 * size {
                    return Err(unsupported("invalid Word table border array"));
                }
                for s in 0..6 {
                    self.borders[s] = Some(Border::read(&b[1 + s * size..], old)?);
                }
            }
            0xd620 | 0xd62f => {
                let old = code == 0xd620;
                if b[0] != if old { 7 } else { 11 } {
                    return Err(unsupported("invalid Word cell border operand length"));
                }
                // An unknown optional border-side flag has no inferred meaning.
                // Omit this property with the caller's unsupported-property warning.
                if b[3] & !if old { 15 } else { 63 } != 0 {
                    return Ok(false);
                }
                let r = range(&b[1..], self.cells.len())?;
                let border = Border::read(&b[4..], old)?;
                for c in &mut self.cells[r] {
                    for s in 0..6 {
                        if b[3] & (1 << s) != 0 {
                            c.borders[s] = Some(border.clone());
                        }
                    }
                }
            }
            0x9602 => self.gap = nonnegative(b)?,
            0x9407 => self.height = signed(b)?,
            0x3404 => self.header = boolean(b[0])?,
            0x3403 | 0x3466 => self.cant_split = boolean(b[0])?,
            0x3615 => self.autofit = boolean(b[0])?,
            0x560b | 0x5664 => {
                let value = u16_at(b, 0)?;
                if value > 1 {
                    return Err(unsupported("invalid Word table direction"));
                }
                self.bidi |= value != 0;
            }
            0x5400 | 0x548a => {
                let value = u16_at(b, 0)?;
                if value > 2 {
                    return Err(unsupported("invalid Word table alignment"));
                }
                self.alignment = (value, code == 0x5400);
            }
            0xd632 | 0xd634 if b[0] == 6 => {
                if b[3] & !15 != 0 || ![0, 3].contains(&b[4]) {
                    return Err(unsupported("invalid Word cell margin"));
                }
                let width = nonnegative(&b[5..])? as u16;
                if b[4] == 0 && width != 0 {
                    return Err(unsupported("nonzero Word nil cell margin"));
                }
                if code == 0xd634 {
                    if b[1..3] != [0, 1] {
                        return Err(unsupported("invalid Word default margin range"));
                    }
                    for i in 0..4 {
                        if b[3] & (1 << i) != 0 {
                            self.margins[i] = width;
                        }
                    }
                } else {
                    let r = range(&b[1..], self.cells.len())?;
                    for c in &mut self.cells[r] {
                        for i in 0..4 {
                            if b[3] & (1 << i) != 0 {
                                c.margins[i] = Some(width);
                            }
                        }
                    }
                }
            }
            _ => return Ok(false),
        }
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn definition_edges_do_not_subtract_the_origin_gap_twice() {
        let mut row = Row::default();
        // One column with outer edges at 100 and 1100, no explicit TC80.
        row.apply(0xd608, &[6, 0, 1, 100, 0, 0x4c, 4]).unwrap();
        row.apply(0x9602, &[108, 0]).unwrap();
        assert_eq!(row.origin(), 100);
        assert_eq!(row.cells[0].width, 1000);
        row.apply(0x9601, &[208, 0]).unwrap();
        assert_eq!(row.origin(), 100);
    }
    #[test]
    fn unknown_old_border_sides_are_warned_not_reinterpreted_as_modern_diagonals() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 1, 100, 0]).unwrap();
        assert!(!row.apply(0xd620, &[7, 0, 1, 16, 8, 1, 0, 0]).unwrap());
        assert!(row.cells[0].borders.iter().all(Option::is_none));
        assert!(row
            .apply(0xd62f, &[11, 0, 1, 16, 0, 0, 0, 0, 8, 1, 0, 0])
            .unwrap());
        assert!(row.cells[0].borders[4].is_some());
    }
    #[test]
    fn cell_edits_preserve_widths_and_merge_primary() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 3, 0xa0, 5]).unwrap();
        row.apply(0x7623, &[1, 2, 0xd0, 2]).unwrap();
        row.apply(0x5624, &[0, 2]).unwrap();
        assert_eq!(
            row.cells.iter().map(|c| c.width).collect::<Vec<_>>(),
            [1440, 720, 1440]
        );
        assert_eq!(
            row.cells.iter().map(|c| c.flags & 3).collect::<Vec<_>>(),
            [2, 1, 0]
        );
        assert!(row.apply(0x5622, &[0, 3]).is_err());
        assert!(row.apply(0x7621, &[0, 64, 1, 0]).is_err());
    }
    #[test]
    fn depth_is_direct_and_bounded() {
        let mut p = Properties::default();
        p.apply(0x2416, &[1]).unwrap();
        assert_eq!(p.depth().unwrap(), 1);
        p.apply(0x6649, &2u32.to_le_bytes()).unwrap();
        p.apply(0x664a, &(-1i32).to_le_bytes()).unwrap();
        assert_eq!(p.depth().unwrap(), 1);
        p.apply(0x6649, &u32::MAX.to_le_bytes()).unwrap();
        assert!(p.depth().is_err());
    }
}
