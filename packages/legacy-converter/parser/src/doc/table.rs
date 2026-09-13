//! Binary table properties, [MS-DOC] 2.4.3, 2.6.3, TDefTableOperand/TC80.
//! A row's definition belongs to its TTP mark, not its first text paragraph.
use super::{border::Border, u16_at, u32_at, unsupported};
mod shading;
pub(in crate::doc) use shading::Shading;
#[cfg(feature = "direct-doc")]
pub(in crate::doc) use shading::{Color, DirectShadingFacts};
mod position;
pub(in crate::doc) use position::Position;
mod width;
pub(crate) use width::PreferredWidth;

/// Cell shading prepared for the table-style-aware cascade. This stays
/// separate from the compatibility projection in `Cell::shading` until the
/// table style has been resolved.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(in crate::doc) enum PreparedCellShading {
    StyleDeferred,
    Explicit(Shading),
    Unsupported,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) struct TableShadingPolicy {
    pub effective_nfib: u16,
    pub interpret_table_styles: bool,
}

impl TableShadingPolicy {
    fn enabled(self) -> bool {
        self.effective_nfib > 0x00d9 && self.interpret_table_styles
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum StyleAwareShadingApply {
    Unhandled,
    Handled,
    HandledUnsupported,
}

#[derive(Clone, Default)]
pub struct Cell {
    pub shading: Option<Shading>,
    pub(in crate::doc) prepared_shading: Option<PreparedCellShading>,
    pub width: i32,
    pub flags: u16,
    pub preferred: Option<PreferredWidth>,
    pub margins: [Option<u16>; 4],
    pub borders: [Option<Border>; 6],
}

pub struct Properties<R = Row> {
    pub in_table: bool,
    depth: Option<i32>,
    pub row_end: bool,
    pub inner_cell: bool,
    pub inner_row: bool,
    pub row: R,
}

impl Default for Properties {
    fn default() -> Self {
        Self {
            in_table: false,
            depth: None,
            row_end: false,
            inner_cell: false,
            inner_row: false,
            row: Row::default(),
        }
    }
}

#[derive(Clone)]
pub struct Row {
    pub position: Position,
    pub shading: Option<Shading>,
    pub identity: std::collections::BTreeMap<u16, Vec<u8>>,
    /// Last directly selected table style. Resolution against STSH is deferred
    /// until the row-owning TTP is known ([MS-DOC] 2.4.6.6 Part 1, step 6).
    pub table_style: Option<usize>,
    /// Live optional-style flags from the last sprmTTlp on this row. The TLP
    /// itl field is historical auto-format metadata and is not live formatting.
    pub table_style_options: Option<u16>,
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
    pub preferred_width: Option<PreferredWidth>,
}

impl Default for Row {
    fn default() -> Self {
        Self {
            position: Position::default(),
            shading: None,
            identity: Default::default(),
            table_style: None,
            table_style_options: None,
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
            preferred_width: None,
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

impl<R> Properties<R> {
    pub fn depth(&self) -> Result<usize, String> {
        let n = self.depth.unwrap_or(i32::from(self.in_table));
        // Resource policy independent of the file's representable table depth.
        if !(0..=32).contains(&n) {
            return Err(unsupported("Word table nesting budget exceeded"));
        }
        Ok(n as usize)
    }
}

impl Properties {
    pub(super) fn borrowed(&self) -> Properties<&Row> {
        Properties {
            in_table: self.in_table,
            depth: self.depth,
            row_end: self.row_end,
            inner_cell: self.inner_cell,
            inner_row: self.inner_row,
            row: &self.row,
        }
    }

    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn retained_heap_bytes(&self) -> Result<usize, String> {
        let mut bytes = self
            .row
            .cells
            .capacity()
            .checked_mul(std::mem::size_of::<Cell>())
            .ok_or("OUTPUT_TOO_LARGE")?;
        // BTreeMap exposes neither node layout nor allocation capacity. Charge
        // the retained key/value payload and each owned operand's capacity.
        // Allocator-internal node overhead is not exact byte-accounted; the
        // entry count is bounded by the fixed identity-code set in Row::apply.
        let identity_entry = std::mem::size_of::<(u16, Vec<u8>)>();
        bytes = bytes
            .checked_add(
                self.row
                    .identity
                    .len()
                    .checked_mul(identity_entry)
                    .ok_or("OUTPUT_TOO_LARGE")?,
            )
            .ok_or("OUTPUT_TOO_LARGE")?;
        for operand in self.row.identity.values() {
            bytes = bytes
                .checked_add(operand.capacity())
                .ok_or("OUTPUT_TOO_LARGE")?;
        }
        for border in self
            .row
            .borders
            .iter()
            .chain(self.row.cells.iter().flat_map(|cell| cell.borders.iter()))
            .flatten()
        {
            bytes = bytes
                .checked_add(border.retained_bytes()?)
                .ok_or("OUTPUT_TOO_LARGE")?;
        }
        Ok(bytes)
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

    /// Applies the shading Sprms whose meaning changes when table styles are
    /// interpreted. [MS-DOC] 2.6.3 and 2.9.53 require nFib > 0x00D9 readers
    /// that interpret table styles to ignore the compatibility arrays and use
    /// the Raw arrays. Raw ShdNil and omitted entries defer to the style;
    /// ShdAuto is an explicit no-fill value.
    pub(in crate::doc) fn apply_style_aware_shading(
        &mut self,
        code: u16,
        b: &[u8],
        policy: TableShadingPolicy,
    ) -> Result<StyleAwareShadingApply, String> {
        if !policy.enabled() {
            return Ok(StyleAwareShadingApply::Unhandled);
        }
        match code {
            // Compatibility arrays and range shading are ignored under the
            // enabled policy ([MS-DOC] 2.6.3 sprmTDefTableShd*, sprmTSetShd,
            // and sprmTSetShdOdd).
            0xd609 | 0xd612 | 0xd616 | 0xd60c | 0xd62d | 0xd62e => {
                Ok(StyleAwareShadingApply::Handled)
            }
            0xd670 | 0xd671 | 0xd672 => {
                let start = match code {
                    0xd671 => 22,
                    0xd672 => 44,
                    _ => 0,
                };
                let max = if code == 0xd672 { 19 } else { 22 };
                let cb = usize::from(
                    *b.first()
                        .ok_or_else(|| unsupported("short Word raw shading array"))?,
                );
                if b.len() != cb + 1
                    || cb % 10 != 0
                    || cb / 10 > max
                    || (cb > 0 && start + cb / 10 > self.cells.len())
                {
                    return Err(unsupported("invalid Word raw cell shading array"));
                }

                let mut values = Vec::with_capacity(cb / 10);
                let mut supported = true;
                for bytes in b[1..].chunks_exact(10) {
                    let shd_nil = Shading::is_shd_nil(bytes);
                    let value = match Shading::read(bytes, false)? {
                        Some(_) if shd_nil => PreparedCellShading::StyleDeferred,
                        Some(value) => PreparedCellShading::Explicit(value),
                        None => {
                            supported = false;
                            PreparedCellShading::Unsupported
                        }
                    };
                    values.push(value);
                }

                // The omitted tail has the Raw-array default: table-style
                // shading. This is a replacement, not a sparse update.
                let end = (start + max).min(self.cells.len());
                for cell in &mut self.cells[start.min(end)..end] {
                    cell.prepared_shading = Some(PreparedCellShading::StyleDeferred);
                }
                if !values.is_empty() {
                    for (cell, value) in self.cells[start..start + values.len()]
                        .iter_mut()
                        .zip(values)
                    {
                        cell.prepared_shading = Some(value);
                    }
                }
                Ok(if supported {
                    StyleAwareShadingApply::Handled
                } else {
                    StyleAwareShadingApply::HandledUnsupported
                })
            }
            _ => Ok(StyleAwareShadingApply::Unhandled),
        }
    }

    pub fn apply(&mut self, code: u16, b: &[u8]) -> Result<bool, String> {
        if matches!(
            code,
            0x7469 | 0x563a | 0x360d | 0x3465 | 0x940e | 0x940f | 0x9410 | 0x9411 | 0x941e | 0x941f
        ) {
            self.identity.insert(
                code,
                if code == 0x360d {
                    vec![b[0] & 0xf0]
                } else {
                    b.to_vec()
                },
            );
        }
        match code {
            0x563a => {
                // [MS-DOC] 2.6.3 sprmTIstd: each application selects a table
                // style and resets the previous selection. Retain last-wins
                // state here. Returning false keeps the existing unsupported
                // output gate until table TAPX/PAPX/CHPX are projected.
                self.table_style = Some(usize::from(u16_at(b, 0)?));
                return Ok(false);
            }
            0x740a => {
                // MS-DOC 2.9.326 TLP: ignore the historical itl and retain
                // only the live grfatl options. Returning false keeps TTlp
                // itself behind the existing table-format admission gate.
                if b.len() != 4 {
                    return Err(unsupported("invalid Word table style options"));
                }
                let _ = u16_at(b, 0)? as i16;
                self.table_style_options = Some(u16_at(b, 2)?);
                return Ok(false);
            }
            0x360d | 0x940e | 0x940f | 0x9410 | 0x9411 | 0x941e | 0x941f | 0x3465 => {
                return self.position.apply(code, b);
            }
            // MS-DOC 2.6.3: use the compatibility shading arrays while table
            // style interpretation is unsupported. Do not apply ShdRaw as a
            // second layer: its ShdNil requires the unresolved table style.
            0xd609 | 0xd612 | 0xd616 | 0xd60c => {
                let old = code == 0xd609;
                let size = if old { 2 } else { 10 };
                let start = match code {
                    0xd616 => 22,
                    0xd60c => 44,
                    _ => 0,
                };
                let max = match code {
                    0xd609 => 63,
                    0xd60c => 19,
                    _ => 22,
                };
                let cb = usize::from(
                    *b.first()
                        .ok_or_else(|| unsupported("short Word shading array"))?,
                );
                if b.len() != cb + 1
                    || cb % size != 0
                    || cb / size > max
                    || (cb > 0 && start + cb / size > self.cells.len())
                {
                    return Err(unsupported("invalid Word cell shading array"));
                }
                // MS-DOC 2.9.53: omitted trailing entries are non-shaded, not
                // a sparse update retaining colors from an earlier array.
                if !old {
                    let end = (start + max).min(self.cells.len());
                    for cell in &mut self.cells[start.min(end)..end] {
                        cell.shading = None;
                    }
                }
                if cb == 0 {
                    return Ok(true);
                }
                let mut supported = true;
                for (cell, bytes) in self.cells[start..start + cb / size]
                    .iter_mut()
                    .zip(b[1..].chunks_exact(size))
                {
                    cell.shading = Shading::read(bytes, old)?;
                    supported &= cell.shading.is_some();
                }
                return Ok(supported);
            }
            0xd62d | 0xd62e => {
                if b.len() != 13 || b[0] != 12 {
                    return Err(unsupported("invalid Word cell shading range length"));
                }
                let cells = range(&b[1..], self.cells.len())?;
                let value = Shading::read(&b[3..], false)?;
                let supported = value.is_some();
                for index in cells.step_by(if code == 0xd62e { 2 } else { 1 }) {
                    self.cells[index].shading = value.clone();
                }
                return Ok(supported);
            }
            0xd660 => {
                if b.len() != 11 || b[0] != 10 {
                    return Err(unsupported("invalid Word table shading length"));
                }
                self.shading = Shading::read(&b[1..], false)?;
                return Ok(self.shading.is_some());
            }
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
                // [MS-DOC] 2.9.321 defines rgTc80 as an array of complete
                // 20-byte TC80 structures. Whole entries may be omitted or
                // exceed NumberOfColumns, but a partial entry is malformed.
                if b[end..].len() % 20 != 0 {
                    return Err(unsupported("partial Word TC80 table definition"));
                }
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
                        // [MS-DOC] 2.9.317/2.9.342: TCGRF.vertMerge is a
                        // VerticalMergeFlag; its 2-bit value 2 is reserved.
                        if (cell.flags >> 5) & 3 == 2 {
                            return Err(unsupported("reserved Word TC80 vertical merge"));
                        }
                        cell.preferred = PreferredWidth::tc80(cell.flags, u16_at(tc, 2)?)?;
                        for s in 0..4 {
                            cell.borders[s] = Some(Border::read(&tc[4 + s * 4..], true)?);
                        }
                    }
                    cells.push(cell);
                }
                self.cells = cells;
            }
            0xf614 => self.preferred_width = PreferredWidth::table(b)?,
            0xd635 => {
                if b.len() != 6 || b[0] != 5 {
                    return Err(unsupported("invalid Word cell width operand length"));
                }
                let cells = range(&b[1..], self.cells.len())?;
                let preferred = PreferredWidth::part(&b[3..])?;
                for cell in &mut self.cells[cells] {
                    cell.preferred = preferred;
                }
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
    fn table_identity_ignores_position_padding_but_retains_anchor_changes() {
        let mut a = Row::default();
        let mut b = Row::default();
        a.apply(0x360d, &[0x20]).unwrap();
        b.apply(0x360d, &[0x2f]).unwrap();
        assert_eq!(a.identity, b.identity);
        assert_eq!(a.position.xml(), b.position.xml());
        b.apply(0x360d, &[0x60]).unwrap();
        assert_ne!(a.identity, b.identity);
    }

    #[test]
    fn tistd_retains_the_last_selected_table_style_without_claiming_projection() {
        let mut row = Row::default();
        assert_eq!(row.table_style, None);
        assert!(!row.apply(0x563a, &7u16.to_le_bytes()).unwrap());
        assert_eq!(row.table_style, Some(7));
        assert!(!row.apply(0x563a, &8u16.to_le_bytes()).unwrap());
        assert_eq!(row.table_style, Some(8));
        assert_eq!(row.identity[&0x563a], 8u16.to_le_bytes());
    }

    #[test]
    fn ttlp_retains_only_live_options_and_stays_behind_the_table_gate() {
        let mut row = Row::default();
        assert_eq!(row.table_style_options, None);
        assert!(!row.apply(0x740a, &[7, 0, 0x20, 0x03]).unwrap());
        assert_eq!(row.table_style_options, Some(0x0320));

        // TTlp is row-local and absent from the exhaustive adjacent-row table
        // identity in MS-DOC 2.4.3. Neither its historical itl nor its live
        // grfatl splits the logical table used for row/column conditions.
        let mut other = Row::default();
        other.apply(0x740a, &[0xff, 0xff, 0x40, 0]).unwrap();
        assert_eq!(row.identity, other.identity);
        assert!(!row.apply(0x740a, &[0xff, 0xff, 0x20, 0x03]).unwrap());
        assert!(row.apply(0x740a, &[0, 0, 0]).is_err());
        assert!(row.apply(0x740a, &[0, 0, 0, 0, 0]).is_err());
    }

    fn shade(pattern: u16) -> Vec<u8> {
        let mut b = vec![10, 0, 0, 0, 255, 0x12, 0x34, 0x56, 0];
        b.extend(pattern.to_le_bytes());
        b
    }
    fn raw_shade(pattern: u16) -> [u8; 10] {
        shade(pattern)[1..].try_into().unwrap()
    }
    fn raw_array(value: [u8; 10], count: usize) -> Vec<u8> {
        let mut operand = Vec::with_capacity(1 + count * 10);
        operand.push((count * 10) as u8);
        for _ in 0..count {
            operand.extend(value);
        }
        operand
    }
    fn style_aware_policy() -> TableShadingPolicy {
        TableShadingPolicy {
            effective_nfib: 0x00da,
            interpret_table_styles: true,
        }
    }

    #[test]
    fn empty_raw_segments_are_safe_on_short_rows_and_reset_only_their_cells() {
        for size in [0, 1, 21, 22, 23, 43, 44, 45, 63] {
            for (code, start, max) in [(0xd670, 0, 22), (0xd671, 22, 22), (0xd672, 44, 19)] {
                let mut row = Row {
                    cells: vec![
                        Cell {
                            prepared_shading: Some(PreparedCellShading::Unsupported),
                            ..Cell::default()
                        };
                        size
                    ],
                    ..Row::default()
                };
                assert_eq!(
                    row.apply_style_aware_shading(code, &[0], style_aware_policy())
                        .unwrap(),
                    StyleAwareShadingApply::Handled
                );
                for (index, cell) in row.cells.iter().enumerate() {
                    assert_eq!(
                        cell.prepared_shading,
                        Some(if (start..start + max).contains(&index) {
                            PreparedCellShading::StyleDeferred
                        } else {
                            PreparedCellShading::Unsupported
                        }),
                        "size {size}, code {code:#x}, cell {index}"
                    );
                }
            }
        }
    }

    #[test]
    fn prepared_shading_follows_source_cells_across_insert_delete_and_redefinition() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 3, 100, 0]).unwrap();
        let mut operand = vec![30];
        for pattern in [0, 1, 2] {
            operand.extend(raw_shade(pattern));
        }
        row.apply_style_aware_shading(0xd670, &operand, style_aware_policy())
            .unwrap();
        let original: Vec<_> = row
            .cells
            .iter()
            .map(|cell| cell.prepared_shading.clone())
            .collect();
        row.apply(0x7621, &[1, 1, 100, 0]).unwrap();
        assert_eq!(row.cells.len(), 4);
        assert_eq!(row.cells[0].prepared_shading, original[0]);
        assert_eq!(row.cells[1].prepared_shading, None);
        assert_eq!(row.cells[2].prepared_shading, original[1]);
        assert_eq!(row.cells[3].prepared_shading, original[2]);
        row.apply(0x5622, &[0, 2]).unwrap();
        assert_eq!(
            row.cells
                .iter()
                .map(|cell| cell.prepared_shading.clone())
                .collect::<Vec<_>>(),
            original[1..]
        );
        row.apply(0xd608, &[6, 0, 1, 0, 0, 100, 0]).unwrap();
        assert_eq!(row.cells.len(), 1);
        assert_eq!(row.cells[0].prepared_shading, None);
    }

    #[test]
    fn style_aware_raw_arrays_cover_all_three_bounded_segments() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 63, 1, 0]).unwrap();
        for (code, start, count) in [(0xd670, 0, 22), (0xd671, 22, 22), (0xd672, 44, 19)] {
            assert_eq!(
                row.apply_style_aware_shading(
                    code,
                    &raw_array(raw_shade(0), count),
                    style_aware_policy(),
                )
                .unwrap(),
                StyleAwareShadingApply::Handled
            );
            assert!(row.cells[start..start + count].iter().all(|cell| {
                matches!(
                    &cell.prepared_shading,
                    Some(PreparedCellShading::Explicit(shading))
                        if shading.xml().contains("w:fill=\"123456\"")
                )
            }));
        }
    }

    #[test]
    fn style_aware_raw_nil_and_omitted_entries_defer_but_auto_clears() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 24, 1, 0]).unwrap();
        row.apply_style_aware_shading(0xd670, &raw_array(raw_shade(0), 22), style_aware_policy())
            .unwrap();
        row.apply_style_aware_shading(0xd671, &raw_array(raw_shade(0), 2), style_aware_policy())
            .unwrap();

        let nil = [255, 255, 255, 255, 255, 255, 255, 255, 0, 0];
        assert_eq!(
            row.apply_style_aware_shading(0xd670, &raw_array(nil, 1), style_aware_policy(),)
                .unwrap(),
            StyleAwareShadingApply::Handled
        );
        assert!(row.cells[..22].iter().all(|cell| matches!(
            cell.prepared_shading,
            Some(PreparedCellShading::StyleDeferred)
        )));
        assert!(matches!(
            row.cells[22].prepared_shading,
            Some(PreparedCellShading::Explicit(_))
        ));

        let auto = [0, 0, 0, 255, 0, 0, 0, 255, 0, 0];
        row.apply_style_aware_shading(0xd671, &raw_array(auto, 1), style_aware_policy())
            .unwrap();
        let Some(PreparedCellShading::Explicit(auto)) = &row.cells[22].prepared_shading else {
            panic!("ShdAuto must remain an explicit cell value");
        };
        assert_eq!(
            auto.xml(),
            "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"auto\"/>"
        );
        assert!(matches!(
            row.cells[23].prepared_shading,
            Some(PreparedCellShading::StyleDeferred)
        ));
    }

    #[test]
    fn style_aware_shading_ignores_compatibility_ranges_and_preserves_raw_order() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 2, 1, 0]).unwrap();
        let mut direct = vec![12, 0, 1];
        direct.extend(raw_shade(0));

        row.apply_style_aware_shading(0xd670, &raw_array(raw_shade(0), 1), style_aware_policy())
            .unwrap();
        assert_eq!(
            row.apply_style_aware_shading(0xd62d, &direct, style_aware_policy())
                .unwrap(),
            StyleAwareShadingApply::Handled
        );
        assert!(matches!(
            &row.cells[0].prepared_shading,
            Some(PreparedCellShading::Explicit(shading))
                if shading.xml().contains("w:fill=\"123456\"")
        ));

        row.apply_style_aware_shading(
            0xd670,
            &raw_array([255, 255, 255, 255, 255, 255, 255, 255, 0, 0], 1),
            style_aware_policy(),
        )
        .unwrap();
        assert!(row.cells.iter().all(|cell| matches!(
            cell.prepared_shading,
            Some(PreparedCellShading::StyleDeferred)
        )));
    }

    #[test]
    fn style_aware_policy_is_explicit_and_leaves_legacy_application_unchanged() {
        for policy in [
            TableShadingPolicy {
                effective_nfib: 0x00d9,
                interpret_table_styles: true,
            },
            TableShadingPolicy {
                effective_nfib: 0x00da,
                interpret_table_styles: false,
            },
        ] {
            let mut row = Row::default();
            row.apply(0x7621, &[0, 1, 1, 0]).unwrap();
            assert_eq!(
                row.apply_style_aware_shading(0xd612, &shade(0), policy)
                    .unwrap(),
                StyleAwareShadingApply::Unhandled
            );
            assert!(row.apply(0xd612, &shade(0)).unwrap());
            assert!(row.cells[0].shading.is_some());
            assert!(row.cells[0].prepared_shading.is_none());
        }

        let mut enabled = Row::default();
        enabled.apply(0x7621, &[0, 1, 1, 0]).unwrap();
        assert_eq!(
            enabled
                .apply_style_aware_shading(0xd612, &[1, 0], style_aware_policy())
                .unwrap(),
            StyleAwareShadingApply::Handled
        );
        assert!(enabled.cells[0].shading.is_none());
    }

    #[test]
    fn style_aware_raw_arrays_reject_malformed_and_bound_unsupported_patterns() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 63, 1, 0]).unwrap();
        for code in [0xd670, 0xd671, 0xd672] {
            assert!(row
                .apply_style_aware_shading(code, &[], style_aware_policy())
                .is_err());
            assert!(row
                .apply_style_aware_shading(code, &[1, 0], style_aware_policy())
                .is_err());
        }
        assert!(row
            .apply_style_aware_shading(0xd670, &raw_array(raw_shade(0), 23), style_aware_policy(),)
            .is_err());
        assert!(row
            .apply_style_aware_shading(0xd672, &raw_array(raw_shade(0), 20), style_aware_policy(),)
            .is_err());

        assert_eq!(
            row.apply_style_aware_shading(
                0xd670,
                &raw_array(raw_shade(0x23), 1),
                style_aware_policy(),
            )
            .unwrap(),
            StyleAwareShadingApply::HandledUnsupported
        );
        assert!(matches!(
            row.cells[0].prepared_shading,
            Some(PreparedCellShading::Unsupported)
        ));
        assert_eq!(
            row.apply_style_aware_shading(
                0xd670,
                &raw_array(raw_shade(0xffff), 1),
                style_aware_policy(),
            )
            .unwrap(),
            StyleAwareShadingApply::Handled
        );
        assert!(matches!(
            row.cells[0].prepared_shading,
            Some(PreparedCellShading::Explicit(_))
        ));
        assert!(
            row.apply_style_aware_shading(
                0xd670,
                &raw_array(raw_shade(0x3e), 1),
                style_aware_policy(),
            )
            .is_err()
        );

        let mut one = Row::default();
        one.apply(0x7621, &[0, 1, 1, 0]).unwrap();
        assert!(one
            .apply_style_aware_shading(0xd671, &raw_array(raw_shade(0), 1), style_aware_policy(),)
            .is_err());
    }
    #[test]
    fn shading_arrays_cover_three_segments_and_clear_omitted_segment_tails() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 63, 1, 0]).unwrap();
        for (code, index) in [(0xd612, 0), (0xd616, 22), (0xd60c, 44)] {
            assert!(row.apply(code, &shade(0)).unwrap());
            assert!(row.cells[index]
                .shading
                .as_ref()
                .unwrap()
                .xml()
                .contains("123456"));
            assert!(row.cells[index + 1].shading.is_none());
        }
        assert!(row.apply(0xd609, &[2, 0xe0, 0]).unwrap()); // Ico yellow background.
        assert!(row.cells[0]
            .shading
            .as_ref()
            .unwrap()
            .xml()
            .contains("FFFF00"));
        assert!(row.cells[22].shading.is_some());
        assert!(row.cells[44].shading.is_some());
        assert!(row.apply(0xd612, &[0]).unwrap());
        assert!(row.cells[0].shading.is_none());
        assert!(row.cells[22].shading.is_some());
        assert!(row.cells[44].shading.is_some());
        let mut longer = vec![20];
        longer.extend(&shade(0)[1..]);
        longer.extend(&shade(0)[1..]);
        assert!(row.apply(0xd612, &longer).unwrap());
        assert!(row.cells[1].shading.is_some());
        assert!(row.apply(0xd612, &shade(0)).unwrap());
        assert!(row.cells[0].shading.is_some());
        assert!(row.cells[1].shading.is_none());
        let mut empty = Row::default();
        assert!(empty.apply(0xd616, &[0]).unwrap());
        assert!(empty.apply(0xd60c, &[0]).unwrap());
    }
    #[test]
    fn shading_ranges_start_alternation_at_the_first_selected_cell() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 6, 1, 0]).unwrap();
        let mut b = vec![12, 1, 6];
        b.extend(&shade(0)[1..]);
        assert!(row.apply(0xd62e, &b).unwrap());
        assert_eq!(
            row.cells
                .iter()
                .map(|c| c.shading.is_some())
                .collect::<Vec<_>>(),
            [false, true, false, true, false, true]
        );
        b[1] = 0;
        b[2] = 6;
        assert!(row.apply(0xd62d, &b).unwrap());
        assert!(row.cells.iter().all(|c| c.shading.is_some()));
        // Defined but unmappable patterns clear the stale value and warn.
        b[11] = 0x23;
        assert!(!row.apply(0xd62d, &b).unwrap());
        assert!(row.cells.iter().all(|c| c.shading.is_none()));
    }
    #[test]
    fn shading_bad_lengths_and_out_of_range_cells_reject() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 63, 1, 0]).unwrap();
        for code in [0xd609, 0xd612, 0xd616, 0xd60c, 0xd660, 0xd62d, 0xd62e] {
            assert!(row.apply(code, &[]).is_err());
            assert!(row.apply(code, &[1, 0]).is_err());
        }
        let mut b = vec![200];
        b.resize(201, 0);
        assert!(row.apply(0xd60c, &b).is_err());
        let mut few = Row::default();
        few.apply(0x7621, &[0, 1, 1, 0]).unwrap();
        assert!(few.apply(0xd616, &shade(0)).is_err());
        let mut b = vec![12, 1, 2];
        b.extend(&shade(0)[1..]);
        assert!(few.apply(0xd62d, &b).is_err());
        assert!(!few.apply(0xd670, &shade(0)).unwrap()); // No guessed raw/style cascade.
    }
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
    fn preferred_widths_follow_prl_order_without_changing_physical_edges() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 3, 0xa0, 5]).unwrap();
        let physical = row.cells.iter().map(|c| c.width).collect::<Vec<_>>();
        row.apply(0xf614, &[2, 0xc4, 9]).unwrap(); // 50% table width.
        row.apply(0xd635, &[5, 0, 2, 3, 0xd0, 2]).unwrap();
        assert_eq!(row.preferred_width, Some(PreferredWidth::Percent(2500)));
        assert_eq!(
            row.cells.iter().map(|c| c.preferred).collect::<Vec<_>>(),
            [
                Some(PreferredWidth::Dxa(720)),
                Some(PreferredWidth::Dxa(720)),
                None
            ]
        );
        assert_eq!(
            row.cells.iter().map(|c| c.width).collect::<Vec<_>>(),
            physical
        );

        // A later insertion carries no preference; delete/range operations address
        // the then-current cells, and a later TDefTable replaces earlier facts.
        row.apply(0x7621, &[1, 1, 100, 0]).unwrap();
        assert_eq!(row.cells[1].preferred, None);
        row.apply(0xd635, &[5, 1, 3, 2, 0x88, 0x13]).unwrap();
        assert_eq!(row.cells[1].preferred, Some(PreferredWidth::Percent(5000)));
        assert_eq!(row.cells[2].preferred, Some(PreferredWidth::Percent(5000)));
        row.apply(0x5622, &[1, 3]).unwrap();
        assert_eq!(row.cells.len(), 2);
        row.apply(0xd608, &[6, 0, 1, 0, 0, 0xe8, 3]).unwrap();
        assert_eq!(row.cells[0].preferred, None);
    }

    #[test]
    fn cell_width_operand_rejects_bad_lengths_ranges_and_units_atomically() {
        let mut row = Row::default();
        row.apply(0x7621, &[0, 2, 100, 0]).unwrap();
        for operand in [
            &[4, 0, 1, 3, 100][..],
            &[5, 0, 3, 3, 100, 0],
            &[5, 1, 0, 3, 100, 0],
            &[5, 0, 1, 2, 0x89, 0x13],
            &[5, 0, 1, 1, 1, 0],
        ] {
            assert!(row.apply(0xd635, operand).is_err());
            assert!(row.cells.iter().all(|c| c.preferred.is_none()));
        }
        // TablePart nil ignores its payload and clears a prior preference.
        row.apply(0xd635, &[5, 0, 2, 3, 100, 0]).unwrap();
        row.apply(0xd635, &[5, 0, 2, 0, 0xff, 0xff]).unwrap();
        assert!(row.cells.iter().all(|c| c.preferred.is_none()));
    }

    #[test]
    fn tc80_width_unit_is_typed_independently_of_grid_width() {
        // One physical 1000-twip cell and a TC80 dxa preference above the
        // tighter 31,680-twip FtsWWidth_TablePart maximum.
        let mut operand = vec![26, 0, 1, 0, 0, 0xe8, 3];
        operand.extend((3u16 << 9).to_le_bytes());
        operand.extend((i16::MAX as u16).to_le_bytes());
        operand.resize(27, 0);
        let mut row = Row::default();
        row.apply(0xd608, &operand).unwrap();
        assert_eq!(row.cells[0].width, 1000);
        assert_eq!(
            row.cells[0].preferred,
            Some(PreferredWidth::Dxa(i16::MAX as u16))
        );
    }
    #[test]
    fn tc80_rejects_reserved_vertical_merge_only_for_used_descriptors() {
        fn definition(flags: &[u16]) -> Vec<u8> {
            let mut operand = vec![0, 0, 1, 0, 0, 0xe8, 3];
            for flags in flags {
                operand.extend(flags.to_le_bytes());
                operand.resize(operand.len() + 18, 0);
            }
            let cb = (operand.len() - 1) as u16;
            operand[..2].copy_from_slice(&cb.to_le_bytes());
            operand
        }

        for vertical_merge in [0, 1, 3] {
            let mut row = Row::default();
            row.apply(0xd608, &definition(&[vertical_merge << 5]))
                .unwrap();
            assert_eq!((row.cells[0].flags >> 5) & 3, vertical_merge);
        }
        assert!(Row::default()
            .apply(0xd608, &definition(&[2 << 5]))
            .is_err());

        // [MS-DOC] 2.9.321 ignores TC80 entries beyond NumberOfColumns.
        let mut excess = Row::default();
        excess.apply(0xd608, &definition(&[0, 2 << 5])).unwrap();
        assert_eq!(excess.cells.len(), 1);

        // A wholly omitted rgTc80 uses the default TC80 formatting.
        let mut omitted = Row::default();
        omitted.apply(0xd608, &definition(&[])).unwrap();
        assert_eq!(omitted.cells[0].flags, 0);
        assert_eq!(omitted.cells[0].preferred, None);

        // rgTc80 is an array of complete 20-byte TC80 structures. A partial
        // descriptor for a used column is malformed rather than omitted.
        for partial in 1..20 {
            let mut definition = definition(&[]);
            definition.resize(definition.len() + partial, 0);
            let cb = (definition.len() - 1) as u16;
            definition[..2].copy_from_slice(&cb.to_le_bytes());
            assert!(Row::default().apply(0xd608, &definition).is_err());
        }
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
