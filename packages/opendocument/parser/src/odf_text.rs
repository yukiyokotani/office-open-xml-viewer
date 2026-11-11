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
pub fn paragraph_text(node: roxmltree::Node<'_, '_>) -> String {
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
        inline_text(p, &mut out);
    }
    out
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

/// Implementation bound on the spaces all `text:s` elements of one part
/// expand to (§6.1.3 `text:c`); a part over it rejects the document, so a
/// reader never allocates an unbounded run of spaces.
pub const MAX_PART_SPACES: usize = 16 * 1024 * 1024;

/// The spaces one `text:s` stands for (§19.763 `text:c`, default 1).
pub fn space_count(node: roxmltree::Node<'_, '_>) -> usize {
    node.attribute((crate::xml::NS_TEXT, "c"))
        .and_then(|c| c.trim().parse::<usize>().ok())
        .unwrap_or(1)
}

/// Reject a part whose `text:s` elements expand past [`MAX_PART_SPACES`].
pub fn check_space_runs(root: roxmltree::Node<'_, '_>) -> Result<(), String> {
    let mut total = 0usize;
    for node in root
        .descendants()
        .filter(|n| xml::is(*n, crate::xml::NS_TEXT, "s"))
    {
        total = total.saturating_add(space_count(node));
        if total > MAX_PART_SPACES {
            return Err(crate::error::resource("text space"));
        }
    }
    Ok(())
}

/// Implementation bound on one font family or font-face name (`svg` /
/// `fo:font-family`, `style:font-name*`, a face's `style:name`). Readers copy
/// a resolved family into every font, text run and axis that inherits it,
/// so the bound keeps each copy (and each nesting level's) small; a longer
/// name rejects the part.
pub const MAX_FONT_NAME_BYTES: usize = 1024;

/// Reject a part naming a font longer than [`MAX_FONT_NAME_BYTES`].
pub fn check_font_names(root: roxmltree::Node<'_, '_>) -> Result<(), String> {
    for node in root.descendants().filter(|n| n.is_element()) {
        let face = xml::is(node, crate::xml::NS_STYLE, "font-face");
        for attribute in node.attributes() {
            let name = attribute.name();
            let font = name.starts_with("font-family")
                || name.starts_with("font-name")
                || face && name == "name";
            if font && attribute.value().len() > MAX_FONT_NAME_BYTES {
                return Err(crate::error::resource("font name"));
            }
        }
    }
    Ok(())
}

pub fn inline_text(node: roxmltree::Node<'_, '_>, out: &mut String) {
    for child in node.children() {
        if child.is_text() {
            out.push_str(child.text().unwrap_or(""));
            continue;
        }
        if child.tag_name().namespace() != Some(crate::xml::NS_TEXT) {
            continue;
        }
        match child.tag_name().name() {
            // Readers bound every part with `check_space_runs` first.
            "s" => out.extend(std::iter::repeat_n(' ', space_count(child))),
            "tab" => out.push('\t'),
            "line-break" => out.push('\n'),
            "note" | "ruby-text" => {}
            _ => inline_text(child, out),
        }
    }
}
