//! Cell addresses, range addresses (ODF Part 3 §18.3.1-§18.3.5) and the
//! plain text of paragraphs, shared by the spreadsheet and chart readers.

use crate::xml;

/// A cell address `A1` (1-based row, column).
pub fn parse_cell(reference: &str) -> Option<(u32, u32)> {
    let reference = reference.trim().trim_start_matches('.').replace('$', "");
    let split = reference.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = reference.split_at(split);
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let mut col = 0u32;
    for c in letters.chars() {
        col = col
            .checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)?;
    }
    let row: u32 = digits.parse().ok()?;
    (row >= 1 && col >= 1).then_some((row, col))
}

pub fn column_letters(mut col: u32) -> String {
    let mut out = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    out.iter().rev().collect()
}

/// An ODF cell range address `Sheet1.A1:Sheet1.B2` (§18.3.3/§18.3.4):
/// (sheet name, top-left, bottom-right), 1-based.
/// A parsed range: sheet name, top-left and bottom-right (row, column).
pub type RangeAddress = (Option<String>, (u32, u32), (u32, u32));

pub fn parse_range(address: &str) -> Option<RangeAddress> {
    let (start, end) = split_range(address);
    let (sheet, first) = split_sheet(start);
    let first = parse_cell(first)?;
    let last = match end {
        Some(end) => parse_cell(split_sheet(end).1)?,
        None => first,
    };
    Some((
        sheet,
        (first.0.min(last.0), first.1.min(last.1)),
        (first.0.max(last.0), first.1.max(last.1)),
    ))
}

/// Split at the `:` outside a quoted sheet name.
fn split_range(address: &str) -> (&str, Option<&str>) {
    let mut quoted = false;
    for (i, c) in address.char_indices() {
        match c {
            '\'' => quoted = !quoted,
            ':' if !quoted => return (&address[..i], Some(&address[i + 1..])),
            _ => {}
        }
    }
    (address, None)
}

/// `$'My sheet'.A1` → (Some("My sheet"), "A1").
pub fn split_sheet(reference: &str) -> (Option<String>, &str) {
    let reference = reference.trim();
    let mut quoted = false;
    let mut split = None;
    for (i, c) in reference.char_indices() {
        match c {
            '\'' => quoted = !quoted,
            '.' if !quoted => split = Some(i),
            _ => {}
        }
    }
    match split {
        Some(0) | None => (None, reference.trim_start_matches('.')),
        Some(i) => {
            let sheet = reference[..i].trim_start_matches('$');
            let sheet = if sheet.starts_with('\'') && sheet.ends_with('\'') && sheet.len() >= 2 {
                sheet[1..sheet.len() - 1].replace("''", "'")
            } else {
                sheet.to_string()
            };
            (Some(sheet), &reference[i + 1..])
        }
    }
}

/// A sheet name as a SpreadsheetML formula prefix (`'My sheet'!`).
pub fn sheet_prefix(name: &str) -> String {
    let plain = name
        .chars()
        .all(|c| c.is_alphanumeric() || c == '_' || c == '.')
        && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
    if plain {
        format!("{name}!")
    } else {
        format!("'{}'!", name.replace('\'', "''"))
    }
}

/// Collected text of `text:p` / `text:h` children (§6.1), joined by line
/// feeds, honouring `text:s`, `text:tab` and `text:line-break`.
pub fn paragraph_text(node: roxmltree::Node<'_, '_>) -> Result<String, String> {
    let mut out = String::new();
    let mut first = true;
    for p in node
        .children()
        .filter(|n| xml::is(*n, crate::xml::NS_TEXT, "p") || xml::is(*n, crate::xml::NS_TEXT, "h"))
    {
        if !first {
            out.push('\n');
        }
        first = false;
        inline_text(p, &mut Collapse::default(), &mut out)?;
    }
    Ok(out)
}

/// §6.1.2 white-space collapsing over the character data of one paragraph,
/// across its spans: tab, carriage return and line feed are spaces, a run
/// of spaces is one space, and the spaces that begin or end the paragraph
/// are removed. `text:s`, `text:tab` and `text:line-break` stand for the
/// characters they name and are not character data: a space beside one is
/// kept, and only consecutive character-data spaces merge.
///
/// The state is fed in document order: [`Collapse::text`] for character
/// data, [`Collapse::element`] before the characters of an element. A space
/// is written with the text that follows it, so the space that ends a
/// paragraph is never written.
#[derive(Default)]
pub struct Collapse {
    /// Something of the paragraph has been written.
    started: bool,
    /// A space is owed before the next character.
    pending: bool,
}

impl Collapse {
    /// Append the collapsed form of character data to `out`.
    pub fn text(&mut self, text: &str, out: &mut String) {
        for c in text.chars() {
            if matches!(c, ' ' | '\t' | '\r' | '\n') {
                self.pending = self.started;
            } else {
                self.element(out);
                out.push(c);
            }
        }
    }

    /// An element that stands for characters follows: write the space owed
    /// before it.
    pub fn element(&mut self, out: &mut String) {
        if std::mem::take(&mut self.pending) {
            out.push(' ');
        }
        self.started = true;
    }
}

/// Split a space-separated list of cell range addresses (§18.3.6), leaving
/// spaces inside quoted sheet names alone.
pub fn split_addresses(list: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut quoted = false;
    let mut start = 0;
    for (i, c) in list.char_indices() {
        match c {
            '\'' => quoted = !quoted,
            c if c.is_whitespace() && !quoted => {
                if start < i {
                    out.push(&list[start..i]);
                }
                start = i + c.len_utf8();
            }
            _ => {}
        }
    }
    if start < list.len() {
        out.push(&list[start..]);
    }
    out
}

/// A namespace-qualified attribute value (`of:=SUM(…)`, §18.3.12
/// namespacedToken, §19.595 `table:condition`) without its prefix when the
/// prefix is a namespace bound where the attribute is written; any other
/// text is returned whole.
pub fn without_bound_prefix<'a>(node: roxmltree::Node<'_, '_>, value: &'a str) -> &'a str {
    match value.split_once(':') {
        Some((prefix, rest))
            if !prefix.is_empty()
                && prefix
                    .chars()
                    .all(|c| c.is_alphanumeric() || matches!(c, '-' | '_' | '.'))
                && node.lookup_namespace_uri(Some(prefix)).is_some() =>
        {
            rest
        }
        _ => value,
    }
}

/// The spaces one `text:s` stands for (§19.768 `text:c`, a
/// nonNegativeInteger; one when absent or not a number). The count is the
/// file's: spaces are written through [`crate::budget::push_spaces`], never
/// allocated at once, and a count beyond the platform's integers rejects
/// instead of becoming one space.
pub fn space_count(node: roxmltree::Node<'_, '_>) -> Result<usize, String> {
    let count = crate::xml::bounded_count(
        node.attribute((crate::xml::NS_TEXT, "c")),
        usize::MAX as u64,
        "text space",
    )?;
    Ok(count.map_or(1, |count| count as usize))
}

/// Days since 1970-01-01 of a proleptic Gregorian date.
pub fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let y = if month <= 2 { year - 1 } else { year };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let m = month;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

pub fn days_from_civil_text(value: &str) -> Option<i64> {
    let date = value.split('T').next()?;
    let mut parts = date.splitn(3, '-');
    let year: i64 = parts.next()?.parse().ok()?;
    let month: i64 = parts.next()?.parse().ok()?;
    let day: i64 = parts.next()?.get(..2)?.parse().ok()?;
    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }
    Some(days_from_civil(year, month, day))
}

/// `office:date-value` (§19.372, an xsd:date or xsd:dateTime) as the serial
/// the workbook renderer displays as that date.
///
/// In the 1904 system that is the day count from 1904-01-01. In the 1900
/// system it is the day count from 1899-12-30 from 1900-03-01 on
/// (§18.17.4.1). Before that the renderer, like Excel, counts the
/// nonexistent 1900-02-29 and shows serial `n` as the day `n + 1` after
/// 1899-12-30 (`core/src/excel-date.ts`), so an earlier date is one less:
/// 1900-01-01 is serial 1. The 1900 system has no serial for 1899-12-30 or
/// before (Excel shows serial 0 as "1900-01-00"); such a value keeps its
/// plain day count, which shows its time of day and, for a negative count,
/// the renderer's invalid-date marker.
pub fn date_serial(value: &str, date1904: bool) -> Option<f64> {
    let days = days_from_civil_text(value)?;
    let mut fraction = 0.0;
    if let Some(time) = value.split('T').nth(1) {
        let time = time.trim_end_matches('Z');
        // A zone offset (§19.372 xsd:dateTime, either sign) is dropped: the
        // authored local time is shown.
        let time = time.split(['+', '-']).next().unwrap_or(time);
        let mut parts = time.splitn(3, ':');
        let h: f64 = parts.next()?.parse().ok()?;
        let m: f64 = parts.next().unwrap_or("0").parse().ok()?;
        let s: f64 = parts.next().unwrap_or("0").parse().ok()?;
        fraction = (h * 3600.0 + m * 60.0 + s) / 86_400.0;
    }
    let serial = if date1904 {
        days - days_from_civil(1904, 1, 1)
    } else {
        let days = days - days_from_civil(1899, 12, 30);
        if (1..=60).contains(&days) {
            days - 1
        } else {
            days
        }
    };
    Some(serial as f64 + fraction)
}

/// `office:time-value` (§19.382, an xsd:duration such as `PT12H30M05S`) as a
/// fraction of a day.
pub fn duration_days(value: &str) -> Option<f64> {
    let value = value.trim();
    let (negative, rest) = match value.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, value),
    };
    let rest = rest.strip_prefix('P')?;
    let (date_part, time_part) = match rest.split_once('T') {
        Some((d, t)) => (d, t),
        None => (rest, ""),
    };
    let mut seconds = 0.0;
    let mut number = String::new();
    for c in date_part.chars() {
        match c {
            'D' => {
                seconds += number.parse::<f64>().ok()? * 86_400.0;
                number.clear();
            }
            'Y' | 'M' | 'W' => return None,
            _ => number.push(c),
        }
    }
    for c in time_part.chars() {
        let unit = match c {
            'H' => 3600.0,
            'M' => 60.0,
            'S' => 1.0,
            _ => {
                number.push(c);
                continue;
            }
        };
        seconds += number.parse::<f64>().ok()? * unit;
        number.clear();
    }
    let days = seconds / 86_400.0;
    Some(if negative { -days } else { days })
}

/// Implementation bound on one style value: any attribute of a style,
/// font-face, property, data-style or picture-fill definition (`style:` /
/// `number:` namespaces, `*-properties`, `draw:fill-image`: names, font
/// families, lengths, colours, fill paths), and any style reference
/// (`*style-name`) anywhere. Readers compare, hash and copy style values
/// wherever a style is used (a family into each font that inherits it), so
/// the bound makes each such use a small constant of work; a longer value
/// rejects the part. (LibreOffice's values are a few dozen bytes.)
pub const MAX_STYLE_VALUE_BYTES: usize = 1024;
/// The bound on a `style:map` condition (§19.472), a formula.
pub const MAX_STYLE_CONDITION_BYTES: usize = 4096;

const NS_DATASTYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

/// Reject a part with a style value over [`MAX_STYLE_VALUE_BYTES`] (a map
/// condition over [`MAX_STYLE_CONDITION_BYTES`]).
pub fn check_style_values(root: roxmltree::Node<'_, '_>) -> Result<(), String> {
    for node in root.descendants().filter(|n| n.is_element()) {
        crate::budget::tick(1)?;
        // Style, data-style and property elements, whatever the namespace
        // of an extension property set (`loext:graphic-properties`), and
        // picture-fill definitions, whose package path every shape that
        // uses the fill resolves.
        let namespace = node.tag_name().namespace();
        let style_element = namespace == Some(crate::xml::NS_STYLE)
            || namespace == Some(NS_DATASTYLE)
            || node.tag_name().name().ends_with("-properties")
            || xml::is(node, crate::xml::NS_DRAW, "fill-image");
        for attribute in node.attributes() {
            let name = attribute.name();
            let bound = if style_element && name == "condition" {
                MAX_STYLE_CONDITION_BYTES
            } else if style_element || name.ends_with("style-name") {
                MAX_STYLE_VALUE_BYTES
            } else {
                continue;
            };
            if attribute.value().len() > bound {
                return Err(crate::error::resource("style value"));
            }
        }
    }
    Ok(())
}

pub fn inline_text(
    node: roxmltree::Node<'_, '_>,
    collapse: &mut Collapse,
    out: &mut String,
) -> Result<(), String> {
    for child in node.children() {
        crate::budget::tick(1)?;
        if child.is_text() {
            collapse.text(child.text().unwrap_or(""), out);
            continue;
        }
        if child.tag_name().namespace() != Some(crate::xml::NS_TEXT) {
            continue;
        }
        match child.tag_name().name() {
            "s" => {
                collapse.element(out);
                crate::budget::push_spaces(out, space_count(child)?)?
            }
            "tab" => {
                collapse.element(out);
                out.push('\t')
            }
            "line-break" => {
                collapse.element(out);
                out.push('\n')
            }
            "note" | "ruby-text" => {}
            _ => inline_text(child, collapse, out)?,
        }
    }
    Ok(())
}
