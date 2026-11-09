//! ODF data styles (Part 3 §16.29) as SpreadsheetML number format codes
//! (ECMA-376 Part 1 §18.8.30), the formatting language the XLSX renderer
//! evaluates.
//!
//! Every data-style element has a direct format-code token: digits, literal
//! text, fill characters, date and time fields, the text placeholder `@`.
//! `style:map` conditions (§16.3) become the code's sections: LibreOffice
//! writes a positive/negative pair as `value()>=0` on the negative style, and
//! a three-way split as `value()>0` / `value()<0` on the zero style, which
//! are exactly SpreadsheetML's implicit `pos;neg` and `pos;neg;zero`
//! sections. Other conditions become explicit `[op value]` section
//! conditions. A text style mapped from a number style (`value-type` text)
//! is the fourth (text) section.
//!
//! A field with no SpreadsheetML token is written as the literal text
//! LibreOffice displays for it where that text is fixed (`number:quarter`,
//! `number:week-of-year` have none: they are written as their field names'
//! closest token, see `date_field`).

use crate::xml::{self, NS_STYLE};
use std::collections::HashMap;

pub const NS_NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const NS_LOEXT: &str = crate::xml::NS_LOEXT;
const NS_FO: &str = crate::xml::NS_FO;

/// The code standing for SpreadsheetML's built-in short date (id 14).
pub const BUILTIN_SHORT_DATE: &str = "#builtin:14";

/// A format code where a built-in id cannot be named (differential formats,
/// chart axes): the built-in short date as its §18.8.30 definition.
pub fn literal_code(code: &str) -> String {
    if code == BUILTIN_SHORT_DATE {
        "mm-dd-yy".into()
    } else {
        code.to_string()
    }
}

/// Implementation bound on one generated format code (SpreadsheetML limits
/// a code to 255 characters; LibreOffice codes are far shorter).
const MAX_CODE_CHARS: usize = 1024;
/// Implementation bound on `style:map` sections composed into one code.
const MAX_MAP_SECTIONS: usize = 3;

/// What a data style formats, which decides how `office:value-type` cells
/// are shown when the style does not match (§19.389).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DataKind {
    Number,
    Percentage,
    Currency,
    Date,
    Time,
    Boolean,
    Text,
}

#[derive(Clone, Debug)]
struct Section {
    kind: DataKind,
    code: String,
    /// `style:map` children: (condition, applied style name).
    maps: Vec<(String, String)>,
    /// `number:volatile` and other attributes are irrelevant here.
    general: bool,
}

/// Every named data style of the document, resolved to a format code.
#[derive(Default)]
pub struct DataStyles {
    sections: HashMap<String, Section>,
    resolved: HashMap<String, (String, DataKind)>,
}

impl DataStyles {
    /// Read the data styles below `office:styles` or `office:automatic-styles`.
    pub fn add(&mut self, parent: roxmltree::Node<'_, '_>) {
        for node in parent.children().filter(|n| n.is_element()) {
            if node.tag_name().namespace() != Some(NS_NUMBER) {
                continue;
            }
            let kind = match node.tag_name().name() {
                "number-style" => DataKind::Number,
                "percentage-style" => DataKind::Percentage,
                "currency-style" => DataKind::Currency,
                "date-style" => DataKind::Date,
                "time-style" => DataKind::Time,
                "boolean-style" => DataKind::Boolean,
                "text-style" => DataKind::Text,
                _ => continue,
            };
            let Some(name) = node.attribute((NS_STYLE, "name")) else {
                continue;
            };
            let (code, general) = section_code(node, kind);
            let maps = node
                .children()
                .filter(|n| xml::is(*n, NS_STYLE, "map"))
                .filter_map(|map| {
                    Some((
                        map.attribute((NS_STYLE, "condition"))?.to_string(),
                        map.attribute((NS_STYLE, "apply-style-name"))?.to_string(),
                    ))
                })
                .collect();
            self.sections.insert(
                name.to_string(),
                Section {
                    kind,
                    code,
                    maps,
                    general,
                },
            );
        }
    }

    /// The format code (and kind) of the data style `name`; `None` when the
    /// document does not define it.
    pub fn code(&mut self, name: &str) -> Option<(String, DataKind)> {
        if let Some(found) = self.resolved.get(name) {
            return Some(found.clone());
        }
        let section = self.sections.get(name)?.clone();
        let code = if section.maps.is_empty() {
            section.code.clone()
        } else {
            self.compose(&section)
        };
        let mut code: String = code.chars().take(MAX_CODE_CHARS).collect();
        if code.is_empty() {
            // Nothing is shown for the values the style formats: empty
            // numeric sections, or an empty text section (§18.8.30).
            code = if section.kind == DataKind::Text {
                "General;General;General;".into()
            } else {
                ";;".into()
            };
        }
        self.resolved
            .insert(name.to_string(), (code.clone(), section.kind));
        Some((code, section.kind))
    }

    /// Whether `name` is the plain "General" number style.
    pub fn is_general(&self, name: &str) -> bool {
        self.sections
            .get(name)
            .is_some_and(|s| s.general && s.maps.is_empty())
    }

    fn compose(&self, base: &Section) -> String {
        let mapped = |name: &str| -> Option<&Section> { self.sections.get(name) };
        let mut numeric: Vec<(Option<String>, String)> = Vec::new();
        let mut text: Option<String> = None;
        for (condition, target) in base.maps.iter().take(MAX_MAP_SECTIONS) {
            let Some(section) = mapped(target) else {
                continue;
            };
            if section.kind == DataKind::Text {
                text = Some(section.code.clone());
                continue;
            }
            numeric.push((Some(condition_code(condition)), section.code.clone()));
        }
        let base_code = if base.kind == DataKind::Text {
            text.get_or_insert_with(|| base.code.clone());
            // Numbers no map matches show in the standard format.
            Some("General".to_string())
        } else {
            Some(base.code.clone())
        };
        let conditions: Vec<&str> = numeric.iter().filter_map(|(c, _)| c.as_deref()).collect();
        let mut sections: Vec<String> = Vec::new();
        match (conditions.as_slice(), &base_code) {
            // §16.3 conditions LibreOffice writes for sign-dependent formats.
            (["[>0]", "[<0]", "[=0]"], _) => {
                for (_, code) in &numeric {
                    sections.push(code.clone());
                }
            }
            (["[>=0]"], Some(base)) => {
                sections.push(numeric[0].1.clone());
                sections.push(base.clone());
            }
            (["[>0]", "[<0]"], Some(base)) => {
                sections.push(numeric[0].1.clone());
                sections.push(numeric[1].1.clone());
                sections.push(base.clone());
            }
            _ => {
                for (condition, code) in &numeric {
                    sections.push(format!("{}{}", condition.as_deref().unwrap_or(""), code));
                }
                if let Some(base) = &base_code {
                    sections.push(base.clone());
                }
            }
        }
        if let Some(text) = text {
            while sections.len() < 3 {
                sections.push(
                    sections
                        .first()
                        .cloned()
                        .unwrap_or_else(|| "General".into()),
                );
            }
            sections.truncate(3);
            sections.push(text);
        }
        sections.join(";")
    }
}

/// `value()>=0` → `[>=0]` (§16.3: `value()` compared with a number).
fn condition_code(condition: &str) -> String {
    let condition = condition.trim();
    let rest = condition
        .strip_prefix("value()")
        .unwrap_or(condition)
        .trim();
    for op in [">=", "<=", "!=", "<>", ">", "<", "="] {
        if let Some(value) = rest.strip_prefix(op) {
            let op = if op == "!=" { "<>" } else { op };
            return format!("[{op}{}]", value.trim());
        }
    }
    String::new()
}

/// Quote a literal for a format code: `"text"`, with `"` written as `\"`.
fn literal(text: &str) -> String {
    if text.is_empty() {
        return String::new();
    }
    if text.chars().count() == 1 && matches!(text, " " | "-" | "(" | ")" | ":" | "/" | "$" | "+") {
        return text.to_string();
    }
    // A quote cannot occur inside a quoted literal: it is written as `\"`
    // between the quoted runs.
    text.split('"')
        .map(|part| {
            if part.is_empty() {
                String::new()
            } else {
                format!("\"{part}\"")
            }
        })
        .collect::<Vec<_>>()
        .join("\\\"")
}

fn attr<'a>(node: roxmltree::Node<'a, '_>, name: &str) -> Option<&'a str> {
    node.attribute((NS_NUMBER, name))
}

fn attr_u32(node: roxmltree::Node<'_, '_>, name: &str) -> Option<u32> {
    attr(node, name)
        .or_else(|| node.attribute((NS_LOEXT, name)))
        .and_then(|v| v.parse().ok())
        .map(|v: u32| v.min(30))
}

fn long(node: roxmltree::Node<'_, '_>) -> bool {
    attr(node, "style") == Some("long")
}

/// SpreadsheetML color names for the eight base colors (§18.8.30).
fn color_token(node: roxmltree::Node<'_, '_>) -> Option<&'static str> {
    let color = node
        .children()
        .find(|n| xml::is(*n, NS_STYLE, "text-properties"))?
        .attribute((NS_FO, "color"))?
        .to_ascii_lowercase();
    Some(match color.as_str() {
        "#000000" => "[Black]",
        "#ffffff" => "[White]",
        "#ff0000" => "[Red]",
        "#00ff00" => "[Green]",
        "#0000ff" => "[Blue]",
        "#ffff00" => "[Yellow]",
        "#ff00ff" => "[Magenta]",
        "#00ffff" => "[Cyan]",
        _ => return None,
    })
}

/// The integer digits of `number:number` / `number:scientific-number`
/// (§19.355 min-integer-digits, §19.350 grouping), with embedded text
/// (§16.29.4 `number:embedded-text`, position counted leftward from the
/// decimal separator).
fn integer_digits(node: roxmltree::Node<'_, '_>, min: u32, grouping: bool) -> String {
    let embedded: Vec<(u32, String)> = node
        .children()
        .filter(|n| xml::is(*n, NS_NUMBER, "embedded-text"))
        .filter_map(|n| {
            let position = attr(n, "position")?.parse::<i64>().ok()?;
            let text: String = n.text().unwrap_or("").to_string();
            Some((position.clamp(0, 30) as u32, text))
        })
        .collect();
    let widest = embedded.iter().map(|(p, _)| *p).max().unwrap_or(0);
    let count = (min as usize)
        .max(if grouping { 4 } else { 1 })
        .max(widest as usize);
    // LibreOffice's `loext:max-blank-integer-digits`: that many of the
    // mandatory digits show a blank instead of a zero (SpreadsheetML `?`).
    let blank = node
        .attribute((NS_LOEXT, "max-blank-integer-digits"))
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(0)
        .min(min);
    // Digit slots from the right: '0' for mandatory digits, '#' otherwise.
    let mut slots: Vec<String> = (0..count)
        .map(|i| {
            let i = i as u32;
            if i < min.saturating_sub(blank) {
                "0".into()
            } else if i < min {
                "?".into()
            } else {
                "#".into()
            }
        })
        .collect();
    slots.reverse();
    if grouping && count >= 2 {
        slots[0].push(',');
    }
    // Insert literals: position p sits left of the p-th digit from the right.
    let mut out = String::new();
    for (index, slot) in slots.iter().enumerate() {
        let from_right = (count - index) as u32;
        for (position, text) in &embedded {
            if *position == from_right && from_right != 0 {
                out.push_str(&literal(text));
            }
        }
        out.push_str(slot);
    }
    for (position, text) in &embedded {
        if *position == 0 {
            out.push_str(&literal(text));
        }
    }
    if min == 0 && !grouping && embedded.is_empty() {
        return "#".into();
    }
    out
}

fn decimals(node: roxmltree::Node<'_, '_>) -> String {
    let places = attr_u32(node, "decimal-places").unwrap_or(0);
    let min = attr_u32(node, "min-decimal-places")
        .unwrap_or(places)
        .min(places);
    if places == 0 {
        return String::new();
    }
    let mut out = String::from(".");
    for i in 0..places {
        out.push(if i < min { '0' } else { '#' });
    }
    out
}

/// The code of one data-style element's children, and whether it is the
/// plain "General" number.
fn section_code(style: roxmltree::Node<'_, '_>, kind: DataKind) -> (String, bool) {
    let mut code = String::new();
    if let Some(token) = color_token(style) {
        code.push_str(token);
    }
    // Japanese era fields read as the ja-JP calendar (§18.8.30 `[$-411]`).
    let gengou = style.descendants().any(|n| {
        n.is_element()
            && n.tag_name().namespace() == Some(NS_NUMBER)
            && attr(n, "calendar") == Some("gengou")
    });
    if gengou {
        code.push_str("[$-411]");
    }
    let duration = attr(style, "truncate-on-overflow") == Some("false");
    let children: Vec<_> = style
        .children()
        .filter(|n| n.is_element() && n.tag_name().namespace() == Some(NS_NUMBER))
        .collect();
    // §19.340 number:automatic-order: a numeric day-month-year date whose
    // field order follows the locale is the locale's short date, which
    // SpreadsheetML spells as the built-in format 14 (§18.8.30), rendered
    // in the viewer's locale order.
    if kind == DataKind::Date
        && (attr(style, "automatic-order") == Some("true")
            || attr(style, "format-source") == Some("language"))
    {
        let fields: Vec<&str> = children
            .iter()
            .map(|n| n.tag_name().name())
            .filter(|n| *n != "text")
            .collect();
        let numeric = children
            .iter()
            .all(|n| attr(*n, "textual") != Some("true") && attr(*n, "calendar").is_none());
        let mut sorted = fields.clone();
        sorted.sort_unstable();
        if numeric && sorted == ["day", "month", "year"] {
            return (BUILTIN_SHORT_DATE.into(), false);
        }
    }
    let mut general = false;
    let mut first_time_field = true;
    for (index, child) in children.iter().enumerate() {
        let child = *child;
        match child.tag_name().name() {
            "number" => {
                let grouping = attr(child, "grouping") == Some("true");
                let places = attr_u32(child, "decimal-places");
                let min_int = attr_u32(child, "min-integer-digits").unwrap_or(0);
                let has_embedded = child
                    .children()
                    .any(|n| xml::is(n, NS_NUMBER, "embedded-text"));
                if places.is_none() && !grouping && !has_embedded && min_int <= 1 {
                    // §19.343: without decimal-places the number has the
                    // application's standard (General) precision.
                    code.push_str("General");
                    general = children.len() == 1;
                } else {
                    code.push_str(&integer_digits(child, min_int, grouping));
                    code.push_str(&decimals(child));
                }
                if let Some(factor) =
                    attr(child, "display-factor").and_then(|v| v.parse::<f64>().ok())
                {
                    let mut f = factor;
                    while f >= 999.5 {
                        code.push(',');
                        f /= 1000.0;
                    }
                }
            }
            "scientific-number" => {
                let grouping = attr(child, "grouping") == Some("true");
                let min_int = attr_u32(child, "min-integer-digits").unwrap_or(1);
                code.push_str(&integer_digits(child, min_int, grouping));
                code.push_str(&decimals(child));
                let sign = match attr(child, "forced-exponent-sign")
                    .or_else(|| child.attribute((NS_LOEXT, "forced-exponent-sign")))
                {
                    Some("false") => "-",
                    _ => "+",
                };
                let exponent = attr_u32(child, "min-exponent-digits").unwrap_or(2).max(1);
                code.push('E');
                code.push_str(sign);
                code.push_str(&"0".repeat(exponent as usize));
            }
            "fraction" => {
                if let Some(min_int) = attr_u32(child, "min-integer-digits") {
                    let grouping = attr(child, "grouping") == Some("true");
                    code.push_str(&integer_digits(child, min_int, grouping));
                    code.push(' ');
                }
                let numerator = attr_u32(child, "min-numerator-digits").unwrap_or(1).max(1);
                code.push_str(&"?".repeat(numerator as usize));
                code.push('/');
                match attr(child, "denominator-value") {
                    Some(value) if value.parse::<u64>().is_ok() => code.push_str(value),
                    _ => {
                        let digits = attr_u32(child, "min-denominator-digits")
                            .or_else(|| {
                                attr_u32(child, "max-denominator-value")
                                    .map(|v| v.to_string().len() as u32)
                            })
                            .unwrap_or(1)
                            .max(1);
                        code.push_str(&"?".repeat(digits as usize));
                    }
                }
            }
            "text" => {
                let text = child.text().unwrap_or("");
                // LibreOffice's `loext:blank-width-char`: blanks as wide as
                // the named characters (SpreadsheetML `_c`).
                if let Some(chars) = child.attribute((NS_LOEXT, "blank-width-char")) {
                    let chars = chars.split(',').next().unwrap_or("");
                    if !chars.is_empty() && text.chars().all(|c| c == ' ') {
                        for c in chars.chars() {
                            code.push('_');
                            code.push(c);
                        }
                        continue;
                    }
                }
                if kind == DataKind::Percentage && text.contains('%') {
                    // The percent sign scales the value (§18.8.30), as in ODF.
                    for part in text.split_inclusive('%') {
                        if let Some(before) = part.strip_suffix('%') {
                            code.push_str(&literal(before));
                            code.push('%');
                        } else {
                            code.push_str(&literal(part));
                        }
                    }
                } else {
                    code.push_str(&literal(text));
                }
            }
            "fill-character" => {
                if let Some(c) = child.text().and_then(|t| t.chars().next()) {
                    code.push('*');
                    code.push(c);
                }
            }
            "currency-symbol" => {
                let symbol = child.text().unwrap_or("");
                code.push_str(&literal(symbol));
            }
            "text-content" => code.push('@'),
            "boolean" => code.push_str("General"),
            "day" => code.push_str(if long(child) { "dd" } else { "d" }),
            "month" => {
                let textual = attr(child, "textual") == Some("true");
                code.push_str(match (textual, long(child)) {
                    (true, true) => "mmmm",
                    (true, false) => "mmm",
                    (false, true) => "mm",
                    (false, false) => "m",
                });
            }
            "year" => {
                if attr(child, "calendar") == Some("gengou") {
                    code.push_str(if long(child) { "ee" } else { "e" });
                } else {
                    code.push_str(if long(child) { "yyyy" } else { "yy" });
                }
            }
            // LibreOffice's short era is the abbreviated era name (明), which
            // SpreadsheetML spells `gg`; its long era `ggg` (明治).
            "era" => code.push_str(if long(child) { "ggg" } else { "gg" }),
            "day-of-week" => code.push_str(if long(child) { "dddd" } else { "ddd" }),
            // No SpreadsheetML token: LibreOffice shows "Q1" / "1st quarter"
            // and the ISO week number; the closest fixed reading is kept as
            // literal field text so the cell is never blank.
            "quarter" => code.push_str(&literal("Q")),
            "week-of-year" => code.push_str(&literal("W")),
            "hours" => {
                let token = if long(child) { "hh" } else { "h" };
                if duration && first_time_field {
                    code.push('[');
                    code.push_str(token);
                    code.push(']');
                } else {
                    code.push_str(token);
                }
                first_time_field = false;
            }
            "minutes" => {
                let token = if long(child) { "mm" } else { "m" };
                // `m` reads as minutes only next to an hour or second field.
                let next_to_time = index > 0
                    && children[..index]
                        .iter()
                        .rev()
                        .find(|n| n.tag_name().name() != "text")
                        .is_some_and(|n| n.tag_name().name() == "hours")
                    || children[index + 1..]
                        .iter()
                        .find(|n| n.tag_name().name() != "text")
                        .is_some_and(|n| n.tag_name().name() == "seconds");
                if duration && first_time_field {
                    code.push('[');
                    code.push_str(token);
                    code.push(']');
                } else if next_to_time {
                    code.push_str(token);
                } else {
                    code.push('[');
                    code.push_str(token);
                    code.push(']');
                }
                first_time_field = false;
            }
            "seconds" => {
                let token = if long(child) { "ss" } else { "s" };
                if duration && first_time_field {
                    code.push('[');
                    code.push_str(token);
                    code.push(']');
                } else {
                    code.push_str(token);
                }
                if let Some(places) = attr_u32(child, "decimal-places").filter(|p| *p > 0) {
                    code.push('.');
                    code.push_str(&"0".repeat(places.min(3) as usize));
                }
                first_time_field = false;
            }
            "am-pm" => code.push_str("AM/PM"),
            _ => {}
        }
    }
    // A style whose only fields are empty texts shows nothing (the `;;;`
    // idiom); a style without any field is the standard format.
    if children.is_empty() && (code.is_empty() || code == "[$-411]") {
        return (
            match kind {
                DataKind::Text => "@".into(),
                _ => "General".into(),
            },
            true,
        );
    }
    (code, general)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn codes(styles: &str) -> DataStyles {
        let xml = format!(
            r##"<office:automatic-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="{NS_STYLE}" xmlns:number="{NS_NUMBER}" xmlns:fo="{NS_FO}" xmlns:loext="{NS_LOEXT}">{styles}</office:automatic-styles>"##
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        let mut data = DataStyles::default();
        data.add(document.root_element());
        data
    }

    #[test]
    fn numbers_currency_percent_and_maps() {
        let mut data = codes(concat!(
            r##"<number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>"##,
            r##"<number:number-style style:name="N1"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:number-style>"##,
            r##"<number:percentage-style style:name="P"><number:number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>"##,
            r##"<number:currency-style style:name="C0P0"><number:currency-symbol>€</number:currency-symbol><number:text> </number:text><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>"##,
            r##"<number:currency-style style:name="C0"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text><number:currency-symbol>€</number:currency-symbol><number:text> </number:text><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/><style:map style:condition="value()&gt;=0" style:apply-style-name="C0P0"/></number:currency-style>"##,
            r##"<number:number-style style:name="S"><number:scientific-number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:min-exponent-digits="2"/></number:number-style>"##,
            r##"<number:number-style style:name="F"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="2"/></number:number-style>"##,
            r##"<number:text-style style:name="T"><number:text-content/></number:text-style>"##,
        ));
        assert_eq!(data.code("N0").unwrap().0, "General");
        assert!(data.is_general("N0"));
        assert_eq!(data.code("N1").unwrap().0, "#,##0.00");
        assert_eq!(data.code("P").unwrap().0, "0.0%");
        assert_eq!(
            data.code("C0").unwrap().0,
            "\"€\" #,##0.00;[Red]-\"€\" #,##0.00"
        );
        assert_eq!(data.code("S").unwrap().0, "0.00E+00");
        assert_eq!(data.code("F").unwrap().0, "# ?/??");
        assert_eq!(data.code("T").unwrap().0, "@");
    }

    #[test]
    fn dates_times_and_durations() {
        let mut data = codes(concat!(
            r##"<number:date-style style:name="D"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>"##,
            r##"<number:date-style style:name="D2"><number:day/><number:text>-</number:text><number:month number:textual="true"/><number:text> </number:text><number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text> </number:text><number:am-pm/></number:date-style>"##,
            r##"<number:time-style style:name="T" number:truncate-on-overflow="false"><number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long" number:decimal-places="2"/></number:time-style>"##,
            r##"<number:date-style style:name="J"><number:era number:style="long" number:calendar="gengou"/><number:year number:calendar="gengou"/><number:text>年</number:text></number:date-style>"##,
        ));
        assert_eq!(data.code("D").unwrap().0, "yyyy-mm-dd");
        assert_eq!(data.code("D2").unwrap().0, "d-mmm h:mm AM/PM");
        assert_eq!(data.code("T").unwrap().0, "[h]:mm:ss.00");
        assert_eq!(data.code("J").unwrap().0, "[$-411]ggge\"年\"");
    }

    #[test]
    fn three_way_maps_and_text_sections() {
        let mut data = codes(concat!(
            r##"<number:number-style style:name="P"><number:number number:decimal-places="0" number:min-integer-digits="1"/></number:number-style>"##,
            r##"<number:number-style style:name="N"><number:text>(</number:text><number:number number:decimal-places="0" number:min-integer-digits="1"/><number:text>)</number:text></number:number-style>"##,
            r##"<number:text-style style:name="TX"><number:text>"</number:text><number:text-content/></number:text-style>"##,
            r##"<number:number-style style:name="Z"><number:text>-</number:text><style:map style:condition="value()&gt;0" style:apply-style-name="P"/><style:map style:condition="value()&lt;0" style:apply-style-name="N"/></number:number-style>"##,
            r##"<number:text-style style:name="ALL"><number:text-content/><style:map style:condition="value()&gt;0" style:apply-style-name="P"/><style:map style:condition="value()&lt;0" style:apply-style-name="N"/></number:text-style>"##,
        ));
        assert_eq!(data.code("Z").unwrap().0, "0;(0);-");
        assert_eq!(data.code("TX").unwrap().0, "\\\"@");
        assert_eq!(data.code("ALL").unwrap().0, "0;(0);General;@");
    }
}
