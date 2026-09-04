//! BIFF8 (`.xls`) compatibility subset.
//!
//! The converter preserves worksheet names, scalar/string/boolean/error values,
//! cached formula results, and merged-cell ranges. Formula token programs,
//! formatting, drawings, charts, external links, and macros are never evaluated
//! or copied. See [MS-XLS] 2.4 for record structures and 2.5.293 for BIFF8
//! Unicode strings. FILEPASS and pre-BIFF8 workbooks fail closed.

use std::collections::{BTreeMap, HashSet};

use crate::cfb::CompoundFile;
use crate::ooxml::{write_package, xml_attr, xml_text, ROOT_RELS_XLSX};

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const FILEPASS: u16 = 0x002f;
const BOUNDSHEET8: u16 = 0x0085;
const SST: u16 = 0x00fc;
const CONTINUE: u16 = 0x003c;
const NUMBER: u16 = 0x0203;
const RK: u16 = 0x027e;
const MULRK: u16 = 0x00bd;
const LABELSST: u16 = 0x00fd;
const LABEL: u16 = 0x0204;
const BOOLERR: u16 = 0x0205;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const MERGEDCELLS: u16 = 0x00e5;
const BIFF8: u16 = 0x0600;
const WORKBOOK_GLOBALS: u16 = 0x0005;
const WORKSHEET: u16 = 0x0010;
const MAX_RECORDS: usize = 2_000_000;
const MAX_SHEETS: usize = 65_536;
const MAX_CELLS: usize = 10_000_000;

pub struct XlsConversion {
    pub bytes: Vec<u8>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone)]
struct BoundSheet {
    offset: usize,
    name: String,
    sheet_type: u8,
}

#[derive(Debug, Clone)]
enum CellValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

#[derive(Default)]
struct SheetData {
    rows: BTreeMap<u16, BTreeMap<u16, CellValue>>,
    merged: Vec<(u16, u16, u16, u16)>,
    formula_results: bool,
}

pub fn convert(cfb: &CompoundFile<'_>, max_output_bytes: usize) -> Result<XlsConversion, String> {
    let workbook = cfb
        .stream("Workbook")
        .or_else(|_| cfb.stream("Book"))
        .map_err(unsupported)?;
    let records = records(&workbook)?;
    let first = records
        .first()
        .ok_or_else(|| unsupported("empty BIFF workbook"))?;
    if first.kind != BOF
        || u16_at(first.data, 0)? != BIFF8
        || u16_at(first.data, 2)? != WORKBOOK_GLOBALS
    {
        return Err(unsupported("only BIFF8 workbook globals are supported"));
    }
    if records.iter().any(|record| record.kind == FILEPASS) {
        return Err(unsupported("encrypted BIFF workbooks are not supported"));
    }

    let mut sheets = Vec::new();
    let mut shared_strings = Vec::new();
    let mut saw_sst = false;
    for (index, record) in records.iter().enumerate() {
        match record.kind {
            BOUNDSHEET8 => {
                if sheets.len() >= MAX_SHEETS {
                    return Err(unsupported("too many BIFF worksheets"));
                }
                sheets.push(parse_bound_sheet(record.data)?);
            }
            SST => {
                if saw_sst {
                    return Err(unsupported("multiple BIFF shared string tables"));
                }
                let mut fragments = vec![record.data];
                let mut continued = index + 1;
                while let Some(next) = records.get(continued) {
                    if next.kind != CONTINUE {
                        break;
                    }
                    fragments.push(next.data);
                    continued += 1;
                }
                shared_strings = parse_sst(&fragments)?;
                saw_sst = true;
            }
            _ => {}
        }
    }
    if sheets.is_empty() {
        return Err(unsupported("BIFF workbook contains no sheets"));
    }

    let mut sheet_offsets = HashSet::with_capacity(sheets.len());
    if sheets
        .iter()
        .any(|sheet| !sheet_offsets.insert(sheet.offset))
    {
        return Err(unsupported("duplicate BIFF worksheet offsets"));
    }

    let mut converted = Vec::new();
    let mut skipped_non_worksheets = false;
    let mut formula_results = false;
    for sheet in sheets {
        if sheet.sheet_type != 0 {
            skipped_non_worksheets = true;
            continue;
        }
        let data = parse_sheet(&records, &sheet, &shared_strings)?;
        formula_results |= data.formula_results;
        converted.push((sheet.name, data));
    }
    if converted.is_empty() {
        return Err(unsupported(
            "BIFF workbook contains no supported worksheets",
        ));
    }
    let bytes = build_xlsx(&converted, max_output_bytes)?;
    let mut warnings = vec![
        "legacy-xls:cell-values-and-merges-only".into(),
        "legacy-xls:formatting-drawings-and-external-links-omitted".into(),
    ];
    if formula_results {
        warnings.push("legacy-xls:formulas-replaced-with-cached-results".into());
    }
    if skipped_non_worksheets {
        warnings.push("legacy-xls:non-worksheet-tabs-omitted".into());
    }
    Ok(XlsConversion { bytes, warnings })
}

#[derive(Clone, Copy)]
struct Record<'a> {
    kind: u16,
    offset: usize,
    data: &'a [u8],
}

fn records(bytes: &[u8]) -> Result<Vec<Record<'_>>, String> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        if output.len() >= MAX_RECORDS {
            return Err(unsupported("too many BIFF records"));
        }
        if bytes.len() - offset < 4 {
            return if bytes[offset..].iter().all(|byte| *byte == 0) {
                Ok(output)
            } else {
                Err(unsupported("truncated BIFF record header"))
            };
        }
        let kind = u16_at(bytes, offset)?;
        let size = u16_at(bytes, offset + 2)? as usize;
        if kind == 0 && size == 0 {
            return if bytes[offset..].iter().all(|byte| *byte == 0) {
                Ok(output)
            } else {
                Err(unsupported("unexpected zero BIFF record"))
            };
        }
        let end = offset
            .checked_add(4 + size)
            .ok_or_else(|| unsupported("BIFF record range overflow"))?;
        let data = bytes
            .get(offset + 4..end)
            .ok_or_else(|| unsupported("truncated BIFF record"))?;
        output.push(Record { kind, offset, data });
        offset = end;
    }
    Ok(output)
}

fn parse_bound_sheet(data: &[u8]) -> Result<BoundSheet, String> {
    if data.len() < 8 {
        return Err(unsupported("truncated BOUNDSHEET8 record"));
    }
    let offset = usize::try_from(u32_at(data, 0)?)
        .map_err(|_| unsupported("BIFF sheet offset is too large"))?;
    let sheet_type = data[5];
    let chars = data[6] as usize;
    let high_byte = (data[7] & 0x01) != 0;
    let name = decode_biff_chars(data, 8, chars, high_byte)?.0;
    if name.is_empty() {
        return Err(unsupported("empty BIFF sheet name"));
    }
    Ok(BoundSheet {
        offset,
        name,
        sheet_type,
    })
}

fn parse_sst(fragments: &[&[u8]]) -> Result<Vec<String>, String> {
    let total_bytes = fragments.iter().try_fold(0usize, |total, fragment| {
        total
            .checked_add(fragment.len())
            .ok_or_else(|| unsupported("BIFF shared string table size overflow"))
    })?;
    let mut cursor = SstCursor::new(fragments);
    let counts = cursor.read_fixed(8, "truncated SST record")?;
    let unique = usize::try_from(u32_at(counts, 4)?)
        .map_err(|_| unsupported("BIFF shared string count is too large"))?;
    if unique > MAX_CELLS || unique > total_bytes.saturating_sub(8) / 3 {
        return Err(unsupported("too many BIFF shared strings"));
    }
    let mut strings = Vec::with_capacity(unique);
    for _ in 0..unique {
        let header = cursor.read_fixed(3, "split or truncated BIFF string header")?;
        let chars = u16_at(header, 0)? as usize;
        let flags = header[2];
        let rich_runs = if (flags & 0x08) != 0 {
            let count = cursor.read_fixed(2, "split or truncated BIFF rich string header")?;
            u16_at(count, 0)? as usize
        } else {
            0
        };
        let ext_size = if (flags & 0x04) != 0 {
            let count = cursor.read_fixed(4, "split or truncated BIFF extended string header")?;
            usize::try_from(u32_at(count, 0)?)
                .map_err(|_| unsupported("BIFF extended string data is too large"))?
        } else {
            0
        };
        strings.push(cursor.read_characters(chars, (flags & 0x01) != 0)?);
        cursor.skip_variable(
            rich_runs
                .checked_mul(4)
                .and_then(|size| size.checked_add(ext_size))
                .ok_or_else(|| unsupported("BIFF string size overflow"))?,
        )?;
    }
    Ok(strings)
}

struct SstCursor<'a, 'b> {
    fragments: &'a [&'b [u8]],
    fragment: usize,
    offset: usize,
}

impl<'a, 'b> SstCursor<'a, 'b> {
    fn new(fragments: &'a [&'b [u8]]) -> Self {
        Self {
            fragments,
            fragment: 0,
            offset: 0,
        }
    }

    /// SST non-variable fields cannot straddle BIFF record boundaries.
    fn read_fixed(&mut self, count: usize, message: &str) -> Result<&'b [u8], String> {
        self.advance_empty_fragments();
        let fragment = self
            .fragments
            .get(self.fragment)
            .ok_or_else(|| unsupported(message))?;
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| unsupported("BIFF shared string offset overflow"))?;
        let bytes = fragment
            .get(self.offset..end)
            .ok_or_else(|| unsupported(message))?;
        self.offset = end;
        Ok(bytes)
    }

    /// [MS-XLS] 2.5.293 permits the character array to cross CONTINUE records.
    /// Each continued character fragment starts with its own compression flag.
    fn read_characters(
        &mut self,
        mut remaining: usize,
        mut high_byte: bool,
    ) -> Result<String, String> {
        let mut output = String::new();
        while remaining > 0 {
            if self.current_fragment_exhausted() {
                self.fragment += 1;
                self.offset = 0;
                let option = *self
                    .fragments
                    .get(self.fragment)
                    .and_then(|fragment| fragment.first())
                    .ok_or_else(|| unsupported("truncated continued BIFF string"))?;
                if option > 1 {
                    return Err(unsupported(
                        "invalid continued BIFF string compression flag",
                    ));
                }
                high_byte = option == 1;
                self.offset = 1;
            }

            let fragment = self
                .fragments
                .get(self.fragment)
                .ok_or_else(|| unsupported("truncated continued BIFF string"))?;
            let width = if high_byte { 2 } else { 1 };
            let available_bytes = fragment.len() - self.offset;
            if high_byte && remaining > available_bytes / 2 && !available_bytes.is_multiple_of(2) {
                return Err(unsupported(
                    "continued BIFF Unicode string splits a double-byte character",
                ));
            }
            let take = remaining.min(available_bytes / width);
            if take == 0 {
                return Err(unsupported("empty continued BIFF string fragment"));
            }
            let byte_count = take * width;
            let bytes = &fragment[self.offset..self.offset + byte_count];
            if high_byte {
                let units = bytes
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
                output.extend(char::decode_utf16(units).map(|value| value.unwrap_or('\u{fffd}')));
            } else {
                output.extend(bytes.iter().map(|byte| char::from(*byte)));
            }
            self.offset += byte_count;
            remaining -= take;
        }
        Ok(output)
    }

    fn skip_variable(&mut self, mut remaining: usize) -> Result<(), String> {
        while remaining > 0 {
            self.advance_empty_fragments();
            let fragment = self
                .fragments
                .get(self.fragment)
                .ok_or_else(|| unsupported("truncated BIFF rich or extended string data"))?;
            let take = remaining.min(fragment.len() - self.offset);
            self.offset += take;
            remaining -= take;
        }
        Ok(())
    }

    fn advance_empty_fragments(&mut self) {
        while self.current_fragment_exhausted() && self.fragment < self.fragments.len() {
            self.fragment += 1;
            self.offset = 0;
        }
    }

    fn current_fragment_exhausted(&self) -> bool {
        self.fragments
            .get(self.fragment)
            .is_none_or(|fragment| self.offset == fragment.len())
    }
}

fn parse_biff_string(data: &[u8]) -> Result<(String, usize), String> {
    let chars = u16_at(data, 0)? as usize;
    let flags = *data
        .get(2)
        .ok_or_else(|| unsupported("truncated BIFF string flags"))?;
    let high_byte = (flags & 0x01) != 0;
    let rich_runs = if (flags & 0x08) != 0 {
        u16_at(data, 3)? as usize
    } else {
        0
    };
    let ext_size_offset = 3 + usize::from((flags & 0x08) != 0) * 2;
    let ext_size = if (flags & 0x04) != 0 {
        usize::try_from(u32_at(data, ext_size_offset)?)
            .map_err(|_| unsupported("BIFF extended string data is too large"))?
    } else {
        0
    };
    let chars_offset = ext_size_offset + usize::from((flags & 0x04) != 0) * 4;
    let (value, char_bytes) = decode_biff_chars(data, chars_offset, chars, high_byte)?;
    let consumed = chars_offset
        .checked_add(char_bytes)
        .and_then(|value| value.checked_add(rich_runs * 4))
        .and_then(|value| value.checked_add(ext_size))
        .ok_or_else(|| unsupported("BIFF string size overflow"))?;
    if consumed > data.len() {
        return Err(unsupported("truncated BIFF rich or extended string data"));
    }
    Ok((value, consumed))
}

fn decode_biff_chars(
    data: &[u8],
    offset: usize,
    chars: usize,
    high_byte: bool,
) -> Result<(String, usize), String> {
    let byte_count = chars
        .checked_mul(if high_byte { 2 } else { 1 })
        .ok_or_else(|| unsupported("BIFF string size overflow"))?;
    let bytes = data
        .get(offset..offset + byte_count)
        .ok_or_else(|| unsupported("truncated BIFF string"))?;
    let value = if high_byte {
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        char::decode_utf16(units)
            .map(|value| value.unwrap_or('\u{fffd}'))
            .collect()
    } else {
        bytes.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok((value, byte_count))
}

fn parse_sheet(
    all_records: &[Record<'_>],
    sheet: &BoundSheet,
    shared_strings: &[String],
) -> Result<SheetData, String> {
    let start_index = all_records
        .binary_search_by_key(&sheet.offset, |record| record.offset)
        .map_err(|_| unsupported("BOUNDSHEET8 points outside the BIFF record stream"))?;
    let bof = all_records[start_index];
    if bof.kind != BOF || u16_at(bof.data, 0)? != BIFF8 || u16_at(bof.data, 2)? != WORKSHEET {
        return Err(unsupported(
            "BOUNDSHEET8 does not point to a BIFF8 worksheet",
        ));
    }
    let mut output = SheetData::default();
    let mut cell_count = 0usize;
    let mut pending_formula_string = None;
    for record in &all_records[start_index + 1..] {
        if record.kind == EOF {
            break;
        }
        if record.kind != STRING {
            pending_formula_string = None;
        }
        match record.kind {
            NUMBER => {
                let (row, column) = cell_position(record.data)?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Number(f64_at(record.data, 6)?),
                    &mut cell_count,
                )?;
            }
            RK => {
                let (row, column) = cell_position(record.data)?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Number(decode_rk(u32_at(record.data, 6)?)),
                    &mut cell_count,
                )?;
            }
            MULRK => parse_mul_rk(record.data, &mut output, &mut cell_count)?,
            LABELSST => {
                let (row, column) = cell_position(record.data)?;
                let index = usize::try_from(u32_at(record.data, 6)?)
                    .map_err(|_| unsupported("BIFF shared string index is too large"))?;
                let value = shared_strings
                    .get(index)
                    .ok_or_else(|| unsupported("BIFF shared string index is out of range"))?
                    .clone();
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Text(value),
                    &mut cell_count,
                )?;
            }
            LABEL => {
                let (row, column) = cell_position(record.data)?;
                let (value, _) = parse_biff_string(
                    record
                        .data
                        .get(6..)
                        .ok_or_else(|| unsupported("truncated LABEL record"))?,
                )?;
                insert_cell(
                    &mut output,
                    row,
                    column,
                    CellValue::Text(value),
                    &mut cell_count,
                )?;
            }
            BOOLERR => {
                let (row, column) = cell_position(record.data)?;
                let value = *record
                    .data
                    .get(6)
                    .ok_or_else(|| unsupported("truncated BOOLERR record"))?;
                let is_error = *record
                    .data
                    .get(7)
                    .ok_or_else(|| unsupported("truncated BOOLERR record"))?
                    != 0;
                let value = if is_error {
                    CellValue::Error(error_text(value).into())
                } else {
                    CellValue::Bool(value != 0)
                };
                insert_cell(&mut output, row, column, value, &mut cell_count)?;
            }
            FORMULA => {
                let (row, column) = cell_position(record.data)?;
                match formula_cached_value(record.data)? {
                    FormulaResult::Value(value) => {
                        insert_cell(&mut output, row, column, value, &mut cell_count)?;
                    }
                    FormulaResult::String => pending_formula_string = Some((row, column)),
                    FormulaResult::Empty => {}
                }
                output.formula_results = true;
            }
            STRING => {
                if let Some((row, column)) = pending_formula_string.take() {
                    let (value, _) = parse_biff_string(record.data)?;
                    insert_cell(
                        &mut output,
                        row,
                        column,
                        CellValue::Text(value),
                        &mut cell_count,
                    )?;
                }
            }
            MERGEDCELLS => parse_merged_cells(record.data, &mut output.merged)?,
            FILEPASS => return Err(unsupported("encrypted BIFF worksheet")),
            _ => {}
        }
    }
    Ok(output)
}

fn insert_cell(
    sheet: &mut SheetData,
    row: u16,
    column: u16,
    value: CellValue,
    count: &mut usize,
) -> Result<(), String> {
    let row_values = sheet.rows.entry(row).or_default();
    if row_values.insert(column, value).is_none() {
        *count += 1;
        if *count > MAX_CELLS {
            return Err(unsupported("too many BIFF worksheet cells"));
        }
    }
    Ok(())
}

fn parse_mul_rk(data: &[u8], sheet: &mut SheetData, count: &mut usize) -> Result<(), String> {
    if data.len() < 12 || !(data.len() - 6).is_multiple_of(6) {
        return Err(unsupported("invalid MULRK record"));
    }
    let row = u16_at(data, 0)?;
    let first_column = u16_at(data, 2)?;
    let last_column = u16_at(data, data.len() - 2)?;
    let values = (data.len() - 6) / 6;
    if last_column < first_column || usize::from(last_column - first_column) + 1 != values {
        return Err(unsupported("inconsistent MULRK column range"));
    }
    for index in 0..values {
        let column = first_column
            .checked_add(index as u16)
            .ok_or_else(|| unsupported("MULRK column overflow"))?;
        let raw = u32_at(data, 6 + index * 6)?;
        insert_cell(sheet, row, column, CellValue::Number(decode_rk(raw)), count)?;
    }
    Ok(())
}

fn parse_merged_cells(data: &[u8], output: &mut Vec<(u16, u16, u16, u16)>) -> Result<(), String> {
    let count = u16_at(data, 0)? as usize;
    let required = 2usize
        .checked_add(count * 8)
        .ok_or_else(|| unsupported("MERGEDCELLS size overflow"))?;
    if required > data.len() || output.len() + count > MAX_CELLS {
        return Err(unsupported("invalid or excessive MERGEDCELLS record"));
    }
    for index in 0..count {
        let offset = 2 + index * 8;
        let first_row = u16_at(data, offset)?;
        let last_row = u16_at(data, offset + 2)?;
        let first_column = u16_at(data, offset + 4)?;
        let last_column = u16_at(data, offset + 6)?;
        if first_row > last_row || first_column > last_column || last_column > 255 {
            return Err(unsupported("invalid BIFF merged-cell range"));
        }
        output.push((first_row, last_row, first_column, last_column));
    }
    Ok(())
}

enum FormulaResult {
    Value(CellValue),
    String,
    Empty,
}

fn formula_cached_value(data: &[u8]) -> Result<FormulaResult, String> {
    let raw = data
        .get(6..14)
        .ok_or_else(|| unsupported("truncated FORMULA result"))?;
    if raw[6] == 0xff && raw[7] == 0xff {
        return Ok(match raw[0] {
            0 => FormulaResult::String,
            1 => FormulaResult::Value(CellValue::Bool(raw[2] != 0)),
            2 => FormulaResult::Value(CellValue::Error(error_text(raw[2]).into())),
            3 => FormulaResult::Empty,
            _ => return Err(unsupported("invalid FORMULA cached result type")),
        });
    }
    Ok(FormulaResult::Value(CellValue::Number(f64::from_le_bytes(
        raw.try_into().expect("eight-byte slice"),
    ))))
}

fn cell_position(data: &[u8]) -> Result<(u16, u16), String> {
    let row = u16_at(data, 0)?;
    let column = u16_at(data, 2)?;
    if column > 255 {
        return Err(unsupported("BIFF cell column exceeds the BIFF8 limit"));
    }
    Ok((row, column))
}

fn decode_rk(raw: u32) -> f64 {
    let mut value = if (raw & 0x02) != 0 {
        ((raw as i32) >> 2) as f64
    } else {
        f64::from_bits(u64::from(raw & 0xffff_fffc) << 32)
    };
    if (raw & 0x01) != 0 {
        value /= 100.0;
    }
    value
}

fn build_xlsx(sheets: &[(String, SheetData)], max_output_bytes: usize) -> Result<Vec<u8>, String> {
    let mut workbook = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#,
    );
    let mut workbook_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut content_types = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    );
    let mut parts = vec![
        ("_rels/.rels".into(), ROOT_RELS_XLSX.to_string()),
        ("xl/styles.xml".into(), minimal_styles()),
    ];
    for (index, (name, sheet)) in sheets.iter().enumerate() {
        let id = index + 1;
        workbook.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
            xml_attr(name),
            id,
            id
        ));
        workbook_rels.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>",
            id, id
        ));
        content_types.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
            id
        ));
        parts.push((
            format!("xl/worksheets/sheet{id}.xml"),
            build_sheet_xml(sheet),
        ));
    }
    workbook.push_str("</sheets></workbook>");
    workbook_rels.push_str(&format!(
        "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>",
        sheets.len() + 1
    ));
    content_types.push_str("</Types>");
    parts.push(("xl/workbook.xml".into(), workbook));
    parts.push(("xl/_rels/workbook.xml.rels".into(), workbook_rels));
    parts.push(("[Content_Types].xml".into(), content_types));
    write_package(&parts, max_output_bytes)
}

fn build_sheet_xml(sheet: &SheetData) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
    );
    for (row, cells) in &sheet.rows {
        xml.push_str(&format!("<row r=\"{}\">", u32::from(*row) + 1));
        for (column, value) in cells {
            let reference = cell_reference(*row, *column);
            match value {
                CellValue::Number(value) if value.is_finite() => {
                    xml.push_str(&format!("<c r=\"{reference}\"><v>{value}</v></c>"));
                }
                CellValue::Number(_) => {
                    xml.push_str(&format!("<c r=\"{reference}\" t=\"e\"><v>#NUM!</v></c>"));
                }
                CellValue::Text(value) => {
                    xml.push_str(&format!(
                        "<c r=\"{reference}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
                        xml_text(value)
                    ));
                }
                CellValue::Bool(value) => {
                    xml.push_str(&format!(
                        "<c r=\"{reference}\" t=\"b\"><v>{}</v></c>",
                        u8::from(*value)
                    ));
                }
                CellValue::Error(value) => {
                    xml.push_str(&format!(
                        "<c r=\"{reference}\" t=\"e\"><v>{}</v></c>",
                        xml_text(value)
                    ));
                }
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData>");
    if !sheet.merged.is_empty() {
        xml.push_str(&format!("<mergeCells count=\"{}\">", sheet.merged.len()));
        for (first_row, last_row, first_column, last_column) in &sheet.merged {
            xml.push_str(&format!(
                "<mergeCell ref=\"{}:{}\"/>",
                cell_reference(*first_row, *first_column),
                cell_reference(*last_row, *last_column)
            ));
        }
        xml.push_str("</mergeCells>");
    }
    xml.push_str("</worksheet>");
    xml
}

fn minimal_styles() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.into()
}

fn cell_reference(row: u16, column: u16) -> String {
    let mut value = u32::from(column) + 1;
    let mut letters = Vec::new();
    while value > 0 {
        let remainder = ((value - 1) % 26) as u8;
        letters.push(char::from(b'A' + remainder));
        value = (value - 1) / 26;
    }
    letters.reverse();
    format!(
        "{}{}",
        letters.into_iter().collect::<String>(),
        u32::from(row) + 1
    )
}

fn error_text(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0f => "#VALUE!",
        0x17 => "#REF!",
        0x1d => "#NAME?",
        0x24 => "#NUM!",
        0x2a => "#N/A",
        _ => "#VALUE!",
    }
}

fn unsupported(message: impl Into<String>) -> String {
    format!("UNSUPPORTED:{}", message.into())
}

fn u16_at(bytes: &[u8], offset: usize) -> Result<u16, String> {
    let raw = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| unsupported("truncated BIFF integer"))?;
    Ok(u16::from_le_bytes([raw[0], raw[1]]))
}

fn u32_at(bytes: &[u8], offset: usize) -> Result<u32, String> {
    let raw = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| unsupported("truncated BIFF integer"))?;
    Ok(u32::from_le_bytes(raw.try_into().expect("four-byte slice")))
}

fn f64_at(bytes: &[u8], offset: usize) -> Result<f64, String> {
    let raw = bytes
        .get(offset..offset + 8)
        .ok_or_else(|| unsupported("truncated BIFF number"))?;
    Ok(f64::from_le_bytes(
        raw.try_into().expect("eight-byte slice"),
    ))
}

#[cfg(test)]
mod tests {
    use super::{cell_reference, decode_rk, parse_biff_string, parse_sst, records};

    #[test]
    fn decodes_biff8_unicode_strings() {
        let mut raw = vec![3, 0, 1];
        for unit in "日本語".encode_utf16() {
            raw.extend_from_slice(&unit.to_le_bytes());
        }
        assert_eq!(parse_biff_string(&raw).unwrap().0, "日本語");
    }

    #[test]
    fn decodes_integer_and_scaled_rk_values() {
        assert_eq!(decode_rk((42u32 << 2) | 2), 42.0);
        assert_eq!(decode_rk((1234u32 << 2) | 3), 12.34);
    }

    #[test]
    fn formats_biff8_cell_references() {
        assert_eq!(cell_reference(0, 0), "A1");
        assert_eq!(cell_reference(65535, 255), "IV65536");
    }

    #[test]
    fn rejects_impossible_shared_string_counts_before_allocation() {
        let mut raw = vec![0; 8];
        raw[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_sst(&[&raw]).is_err());
    }

    #[test]
    fn decodes_sst_text_across_continue_records() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&4u16.to_le_bytes());
        sst.push(0);
        sst.extend_from_slice(b"ab");
        let continued = [1, 0x2d, 0x4e, 0x87, 0x65];

        assert_eq!(
            parse_sst(&[&sst, &continued]).unwrap(),
            ["ab\u{4e2d}\u{6587}"]
        );
    }

    #[test]
    fn decodes_unicode_sst_text_after_switching_to_compressed_continue() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&4u16.to_le_bytes());
        sst.push(1);
        for value in ['日', '本'] {
            sst.extend_from_slice(&(value as u16).to_le_bytes());
        }
        let continued = [0, b'a', b'b'];

        assert_eq!(parse_sst(&[&sst, &continued]).unwrap(), ["日本ab"]);
    }

    #[test]
    fn rejects_sst_string_headers_split_across_continue_records() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u16.to_le_bytes());
        let continued = [0, b'a'];

        assert!(parse_sst(&[&sst, &continued]).is_err());
    }

    #[test]
    fn stops_unicode_text_before_the_next_sst_string_header() {
        let mut sst = Vec::new();
        sst.extend_from_slice(&2u32.to_le_bytes());
        sst.extend_from_slice(&2u32.to_le_bytes());
        for value in ['\u{4e2d}', '\u{6587}'] {
            sst.extend_from_slice(&1u16.to_le_bytes());
            sst.push(1);
            sst.extend_from_slice(&(value as u16).to_le_bytes());
        }

        assert_eq!(parse_sst(&[&sst]).unwrap(), ["\u{4e2d}", "\u{6587}"]);
    }

    #[test]
    fn rejects_zero_records_that_hide_nonzero_trailing_data() {
        assert!(records(&[0, 0, 0, 0, 1]).is_err());
        assert!(records(&[0; 5]).unwrap().is_empty());
    }
}
