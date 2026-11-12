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
//! `number:quarter` and `number:week-of-year` have no SpreadsheetML token:
//! they are written as the calendar-field brackets the workbook formatter
//! defines for them (`[Q]`, `[QQ]`, `[WW]`; `core/src/excel-number-format.ts`).

use crate::xml::{self, NS_STYLE};
use std::collections::HashMap;

pub const NS_NUMBER: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const NS_LOEXT: &str = crate::xml::NS_LOEXT;
const NS_FO: &str = crate::xml::NS_FO;

/// The code standing for SpreadsheetML's built-in short date (id 14).
pub const BUILTIN_SHORT_DATE: &str = "#builtin:14";
/// The code standing for the built-in short date and time (id 22).
pub const BUILTIN_SHORT_DATE_TIME: &str = "#builtin:22";

/// A format code where a built-in id cannot be named (differential formats,
/// chart axes): the built-in short date as its §18.8.30 definition.
pub fn literal_code(code: &str) -> String {
    if code == BUILTIN_SHORT_DATE {
        "mm-dd-yy".into()
    } else if code == BUILTIN_SHORT_DATE_TIME {
        "m/d/yy h:mm".into()
    } else {
        code.to_string()
    }
}

/// The week numbering LibreOffice applies to `number:week-of-year` (ODF
/// leaves it to the application, §16.29.17): the first day of the week (0
/// Sunday … 6 Saturday) and the days of the year the first week must hold,
/// by the format's locale (`number:language` / `number:country`).
///
/// Measured with LibreOffice 26.8 over every locale of its locale data
/// (285), each language with an unknown country, and each language alone,
/// on the first and last days of 2015-2026. Each entry is a locale
/// `language-COUNTRY`, `language-*` (the language with any other country)
/// or `language` (the language without a country). Anything not listed
/// numbers weeks from Sunday with 1 January in week 1, which is also what
/// LibreOffice does for a locale it does not know.
const WEEK_RULES: &[((u8, u8), &str)] = &[
    ((0, 1), "bn-IN"),
    ((0, 4), "en-CA fr-CA la la-*"),
    ((0, 7), "zh-HK zh-MO zh-TW"),
    (
        (1, 1),
        "ak-* an-* ast-* az-* be-* bg-* bm-* br-* bs bs-* ca-* cv-* cy-* ee-* \
         el-* en-BZ en-IN eo eo-* es-* eu eu-* ext-* fur-* fy-* gd-* gl-* \
         ha-* haw-* hu-* hy-* id-* jbo jbo-* ka-* kk-* km-* kmr-* ky-* lg-* \
         ltg-* lv-* mfe-* mk-* mos-* ms-* my-* myv-* oc-ES om-* pa-* pap-BQ \
         pap-CW pl-* plt-* pt-AO ro-* ru-* rue-SK rw-* sc-* sk-* sl-* sq-* sr \
         sr-* sun-ID sw-TZ szl-* tg tg-* tr-* tt-* uz uz-* vec-* vi-*",
    ),
    (
        (1, 4),
        "cs-* da-* de-* dsb dsb-* en-BW en-DK en-GB en-GM en-GY en-KE en-LK \
         en-MU en-MW en-MY en-NG en-ZM et-* fi-* fo-* fr-* gos-* hr-* hsb-* \
         ia ia-* it-* kl-* lb-* lij-* ln-* lt-* mn-MN mt-* nb nb-* nds-* nl-* \
         nn nn-* no-* pt-* sv-* vro-* wa-* zh zh-*",
    ),
    (
        (1, 7),
        "axk-CG beq-* bkw-* bvx-* dde-* ebo-* gsc-FR iyx-* kkw-* kng-CG ldi-* \
         mdw-* mkw-* ngz-* njx-* njy-CG oc-* puu-CG sdj-* sg-* tek-* tk-* \
         tsa-* tyx-* vif-* xku-* yom-CG",
    ),
    ((6, 1), "ar ar-* bn-* bo-* dz dz-* fa-* so-*"),
];

/// The [`WEEK_RULES`] entry of a format's locale. A data style without
/// `number:language` takes the locale LibreOffice itself runs in (measured:
/// not the document's default language), which a file does not record; it
/// gets the unlisted rule.
fn week_rule(language: Option<&str>, country: Option<&str>) -> (u8, u8) {
    let Some(language) = language else {
        return (0, 1);
    };
    let find = |key: &str| {
        WEEK_RULES
            .iter()
            .find(|(_, tags)| tags.split_whitespace().any(|tag| tag == key))
            .map(|(rule, _)| *rule)
    };
    let language = language.to_ascii_lowercase();
    match country {
        Some(country) => find(&format!("{language}-{}", country.to_ascii_uppercase()))
            .or_else(|| find(&format!("{language}-*"))),
        None => find(&language),
    }
    .unwrap_or((0, 1))
}

/// Implementation bound on one generated format code (SpreadsheetML limits
/// a code to 255 characters; LibreOffice codes are far shorter). A longer
/// code rejects the document.
const MAX_CODE_CHARS: usize = 1024;
/// The numeric sections a SpreadsheetML code holds before its text section
/// (§18.8.30: at most four sections, the fourth formats text). A data style
/// whose `style:map` conditions need more rejects the document.
const MAX_NUMERIC_SECTIONS: usize = 3;

/// A format code with the §19.346 `number:display-factor` of each numeric
/// section: the displayed number is the value divided by the factor. Empty
/// `divisors` means no section scales.
#[derive(Clone, Debug, PartialEq)]
pub struct FormatCode {
    pub code: String,
    /// Per numeric section, in code order (the text section excluded).
    pub divisors: Vec<f64>,
}

impl FormatCode {
    pub fn plain(code: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            divisors: Vec::new(),
        }
    }

    /// Whether this is the plain standard (General) format.
    pub fn is_general(&self) -> bool {
        self.code == "General" && self.divisors.is_empty()
    }

    /// The code and section divisors for a chart axis or data label
    /// (`ChartModel::*_format_code` / `*_format_divisors`): the chart
    /// formatter divides each section exactly as the cell formatter does.
    pub fn chart_format(&self) -> (String, Option<Vec<f64>>) {
        let divisors = self
            .divisors
            .iter()
            .any(|d| *d != 1.0)
            .then(|| self.divisors.clone());
        (literal_code(&self.code), divisors)
    }
}

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
    /// §19.346 `number:display-factor` of the style's number (1: none).
    divisor: f64,
    /// `style:map` children: (condition, applied style name).
    maps: Vec<(String, String)>,
    /// `number:volatile` and other attributes are irrelevant here.
    general: bool,
}

/// Every named data style of the document, resolved to a format code.
#[derive(Default)]
pub struct DataStyles {
    sections: HashMap<String, Section>,
    resolved: HashMap<String, (FormatCode, DataKind)>,
}

impl DataStyles {
    /// Read the data styles below `office:styles` or `office:automatic-styles`.
    pub fn add(&mut self, parent: roxmltree::Node<'_, '_>) -> Result<(), String> {
        for node in parent.children().filter(|n| n.is_element()) {
            crate::budget::tick(1)?;
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
            let (code, general, divisor) = section_code(node, kind)?;
            let maps = node
                .children()
                .filter(|n| xml::is(*n, NS_STYLE, "map"))
                .filter_map(|map| {
                    let condition = map.attribute((NS_STYLE, "condition"))?;
                    Some((
                        crate::odf_props::of_condition(map, condition)?.to_string(),
                        map.attribute((NS_STYLE, "apply-style-name"))?.to_string(),
                    ))
                })
                .collect();
            self.sections.insert(
                name.to_string(),
                Section {
                    kind,
                    code,
                    divisor,
                    maps,
                    general,
                },
            );
        }
        Ok(())
    }

    /// Resolve every data style once, rejecting the document when one needs
    /// more than a SpreadsheetML code holds (more numeric sections than
    /// [`MAX_NUMERIC_SECTIONS`], or a code over [`MAX_CODE_CHARS`]). Readers
    /// call it after the last [`add`](Self::add), so [`code`](Self::code)
    /// never meets an over-limit style.
    pub fn check(&mut self) -> Result<(), String> {
        let mut names: Vec<String> = self.sections.keys().cloned().collect();
        names.sort_unstable();
        for name in names {
            crate::budget::tick(1)?;
            self.resolve(&name)?;
        }
        Ok(())
    }

    /// The format code (and kind) of the data style `name`; `None` when the
    /// document does not define it.
    pub fn code(&mut self, name: &str) -> Option<(FormatCode, DataKind)> {
        self.resolve(name).ok().flatten()
    }

    fn resolve(&mut self, name: &str) -> Result<Option<(FormatCode, DataKind)>, String> {
        if let Some(found) = self.resolved.get(name) {
            return Ok(Some(found.clone()));
        }
        let Some(section) = self.sections.get(name).cloned() else {
            return Ok(None);
        };
        let mut format = if section.maps.is_empty() {
            FormatCode {
                // A text style without `number:text-content` shows its
                // literal text for any string (§16.29.26): the code's text
                // section, which the formatter knows by its place when it
                // has no `@`.
                code: if section.kind == DataKind::Text
                    && !section.code.is_empty()
                    && !has_text_placeholder(&section.code)
                {
                    format!("General;General;General;{}", section.code)
                } else {
                    section.code.clone()
                },
                divisors: if section.kind != DataKind::Text && section.divisor != 1.0 {
                    vec![section.divisor]
                } else {
                    Vec::new()
                },
            }
        } else {
            self.compose(&section)?
        };
        if format.code.chars().count() > MAX_CODE_CHARS {
            return Err(crate::error::resource("number format code"));
        }
        if format.code.is_empty() {
            // Nothing is shown for the values the style formats: empty
            // numeric sections, or an empty text section (§18.8.30).
            format = FormatCode::plain(if section.kind == DataKind::Text {
                "General;General;General;"
            } else {
                ";;"
            });
        }
        if format.divisors.iter().all(|d| *d == 1.0) {
            format.divisors.clear();
        }
        self.resolved
            .insert(name.to_string(), (format.clone(), section.kind));
        Ok(Some((format, section.kind)))
    }

    /// Whether `name` is the plain "General" number style.
    pub fn is_general(&self, name: &str) -> bool {
        self.sections
            .get(name)
            .is_some_and(|s| s.general && s.maps.is_empty() && s.divisor == 1.0)
    }

    fn compose(&self, base: &Section) -> Result<FormatCode, String> {
        let mapped = |name: &str| -> Option<&Section> { self.sections.get(name) };
        // (condition, code, divisor) of each numeric map target.
        let mut numeric: Vec<(String, String, f64)> = Vec::new();
        // The text section: the last text map's code, borrowed so that many
        // maps to one large text style copy nothing.
        let mut text: Option<&str> = None;
        for (condition, target) in &base.maps {
            let Some(section) = mapped(target) else {
                continue;
            };
            if section.kind == DataKind::Text {
                text = Some(section.code.as_str());
                continue;
            }
            if numeric.len() >= MAX_NUMERIC_SECTIONS {
                return Err(crate::error::resource("number format section"));
            }
            // A condition that is not `value() op number` is not
            // recognized; its map is ignored (§19.472).
            let Some(condition) = condition_code(condition) else {
                continue;
            };
            // A section is spelled out: the built-in markers stand for a
            // whole code only.
            numeric.push((condition, literal_code(&section.code), section.divisor));
        }
        let base_code = if base.kind == DataKind::Text {
            text.get_or_insert(base.code.as_str());
            // Numbers no map matches show in the standard format.
            ("General".to_string(), 1.0)
        } else {
            (literal_code(&base.code), base.divisor)
        };
        let conditions: Vec<&str> = numeric.iter().map(|(c, _, _)| c.as_str()).collect();
        let mut sections: Vec<(String, f64)> = Vec::new();
        match conditions.as_slice() {
            // §16.3 conditions LibreOffice writes for sign-dependent formats.
            ["[>0]", "[<0]", "[=0]"] => {
                for (_, code, divisor) in &numeric {
                    sections.push((code.clone(), *divisor));
                }
            }
            ["[>=0]"] => {
                sections.push((numeric[0].1.clone(), numeric[0].2));
                sections.push(base_code);
            }
            ["[>0]", "[<0]"] => {
                sections.push((numeric[0].1.clone(), numeric[0].2));
                sections.push((numeric[1].1.clone(), numeric[1].2));
                sections.push(base_code);
            }
            _ => {
                for (condition, code, divisor) in &numeric {
                    sections.push((format!("{condition}{code}"), *divisor));
                }
                sections.push(base_code);
            }
        }
        if sections.len() > MAX_NUMERIC_SECTIONS {
            return Err(crate::error::resource("number format section"));
        }
        if let Some(text) = text {
            while sections.len() < MAX_NUMERIC_SECTIONS {
                let first = sections
                    .first()
                    .cloned()
                    .unwrap_or_else(|| ("General".into(), 1.0));
                sections.push(first);
            }
            let divisors = sections.iter().map(|(_, d)| *d).collect();
            let mut codes: Vec<String> = sections.into_iter().map(|(c, _)| c).collect();
            codes.push(text.to_string());
            return Ok(FormatCode {
                code: codes.join(";"),
                divisors,
            });
        }
        Ok(FormatCode {
            code: sections
                .iter()
                .map(|(c, _)| c.as_str())
                .collect::<Vec<_>>()
                .join(";"),
            divisors: sections.iter().map(|(_, d)| *d).collect(),
        })
    }
}

/// Whether a text code holds the `@` placeholder outside quoted literals.
fn has_text_placeholder(code: &str) -> bool {
    let mut quoted = false;
    let mut escaped = false;
    for c in code.chars() {
        if escaped {
            escaped = false;
        } else if c == '\\' && !quoted {
            escaped = true;
        } else if c == '"' {
            quoted = !quoted;
        } else if c == '@' && !quoted {
            return true;
        }
    }
    false
}

/// `value()>=0` → `[>=0]` (§16.3: `value()` compared with a number);
/// `None` for any other condition.
fn condition_code(condition: &str) -> Option<String> {
    let rest = condition.trim().strip_prefix("value()")?.trim();
    for op in [">=", "<=", "!=", "<>", ">", "<", "="] {
        if let Some(value) = rest.strip_prefix(op) {
            let value = value.trim();
            value.parse::<f64>().ok()?;
            let op = if op == "!=" { "<>" } else { op };
            return Some(format!("[{op}{value}]"));
        }
    }
    None
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

/// A digit count of a data-style element (`number:decimal-places`,
/// `number:min-integer-digits`, …). A count no code of
/// [`MAX_CODE_CHARS`] could spell rejects the document; nothing is clipped.
fn attr_u32(node: roxmltree::Node<'_, '_>, name: &str) -> Result<Option<u32>, String> {
    let Some(count) = attr(node, name)
        .or_else(|| node.attribute((NS_LOEXT, name)))
        .and_then(|v| v.trim().parse::<u64>().ok())
    else {
        return Ok(None);
    };
    if count > MAX_CODE_CHARS as u64 {
        return Err(crate::error::resource("number format code"));
    }
    Ok(Some(count as u32))
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
fn integer_digits(
    node: roxmltree::Node<'_, '_>,
    min: u32,
    grouping: bool,
) -> Result<String, String> {
    let mut embedded: Vec<(u32, String)> = Vec::new();
    for n in node
        .children()
        .filter(|n| xml::is(*n, NS_NUMBER, "embedded-text"))
    {
        crate::budget::tick(1)?;
        let Some(position) = attr(n, "position").and_then(|p| p.trim().parse::<i64>().ok()) else {
            continue;
        };
        if position > MAX_CODE_CHARS as i64 {
            return Err(crate::error::resource("number format code"));
        }
        embedded.push((position.max(0) as u32, n.text().unwrap_or("").to_string()));
    }
    let widest = embedded.iter().map(|(p, _)| *p).max().unwrap_or(0);
    // Text at position p needs a digit slot on its left to hold the higher
    // digits (`#-#0` for a hyphen at position 2, as LibreOffice writes it).
    let count = (min as usize)
        .max(if grouping { 4 } else { 1 })
        .max(if widest > 0 { widest as usize + 1 } else { 0 });
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
        return Ok("#".into());
    }
    Ok(out)
}

fn decimals(node: roxmltree::Node<'_, '_>) -> Result<String, String> {
    let places = attr_u32(node, "decimal-places")?.unwrap_or(0);
    // §19.344 `number:decimal-replacement=""`: an integer shows no decimal
    // places, which LibreOffice reads as optional decimals (`0.##`).
    let optional = attr(node, "decimal-replacement") == Some("");
    let min = attr_u32(node, "min-decimal-places")?
        .unwrap_or(if optional { 0 } else { places })
        .min(places);
    if places == 0 {
        return Ok(String::new());
    }
    // At most `MAX_CODE_CHARS` places (`attr_u32`).
    let mut out = String::from(".");
    for i in 0..places {
        out.push(if i < min { '0' } else { '#' });
    }
    Ok(out)
}

/// The code of one data-style element's children, whether it is the plain
/// "General" number, and its `number:display-factor` (1: none).
fn section_code(
    style: roxmltree::Node<'_, '_>,
    kind: DataKind,
) -> Result<(String, bool, f64), String> {
    let mut code = String::new();
    let mut divisor = 1.0f64;
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
            return Ok((BUILTIN_SHORT_DATE.into(), false, 1.0));
        }
        // With hours and minutes it is the short date-time, built-in 22.
        if numeric && sorted == ["day", "hours", "minutes", "month", "year"] {
            return Ok((BUILTIN_SHORT_DATE_TIME.into(), false, 1.0));
        }
    }
    let mut general = false;
    let mut first_time_field = true;
    for child in children.iter() {
        let child = *child;
        match child.tag_name().name() {
            "number" => {
                let grouping = attr(child, "grouping") == Some("true");
                let places = attr_u32(child, "decimal-places")?;
                let min_int = attr_u32(child, "min-integer-digits")?.unwrap_or(0);
                if places.is_none() {
                    // §19.343.2: without `number:decimal-places` the number
                    // takes the default precision. LibreOffice shows such a
                    // number in its standard format, whatever its grouping
                    // and minimum integer digits (measured: `General` for
                    // each of them): the General keyword, which the
                    // formatter also reads beside literal text
                    // (`"$"General`).
                    code.push_str("General");
                    general = children.len() == 1;
                } else {
                    code.push_str(&integer_digits(child, min_int, grouping)?);
                    code.push_str(&decimals(child)?);
                }
                // §19.346: the displayed number is the value divided by the
                // factor (an xsd:double, sign unrestricted). Only a finite
                // positive factor scales, as LibreOffice reads it
                // (xmloff `SvXMLNumFormatContext::AddNumber` scales only for
                // a factor above 0, so a negative factor shows the value
                // unscaled); `0`, `INF` and `NaN` are no factor either.
                if let Some(factor) = attr(child, "display-factor")
                    .and_then(|v| v.trim().parse::<f64>().ok())
                    .filter(|f| f.is_finite() && *f > 0.0)
                {
                    divisor = factor;
                }
            }
            "scientific-number" => {
                let grouping = attr(child, "grouping") == Some("true");
                let min_int = attr_u32(child, "min-integer-digits")?.unwrap_or(1);
                // §19.351 `number:exponent-interval`: the exponent is a
                // multiple of the interval, which SpreadsheetML expresses
                // by that many integer digit slots (`##0.00E+0`).
                let interval = attr_u32(child, "exponent-interval")?.unwrap_or(1);
                if interval > min_int && !grouping {
                    for slot in 0..interval {
                        code.push(if slot < interval - min_int { '#' } else { '0' });
                    }
                } else {
                    code.push_str(&integer_digits(child, min_int, grouping)?);
                }
                code.push_str(&decimals(child)?);
                let sign = match attr(child, "forced-exponent-sign")
                    .or_else(|| child.attribute((NS_LOEXT, "forced-exponent-sign")))
                {
                    Some("false") => "-",
                    _ => "+",
                };
                let exponent = attr_u32(child, "min-exponent-digits")?.unwrap_or(2).max(1);
                code.push('E');
                code.push_str(sign);
                code.push_str(&"0".repeat(exponent as usize));
            }
            "fraction" => {
                // §19.352 / §19.353: the denominator has as many digit
                // slots as the largest permitted denominator
                // (`number:max-denominator-value`), of which the minimum
                // digits are always shown. As LibreOffice writes it: `##?`
                // for a maximum of 999 with one minimum digit; without a
                // maximum the minimum digits are the slots.
                let fixed = attr(child, "denominator-value").filter(|v| v.parse::<u64>().is_ok());
                let min_digits = attr_u32(child, "min-denominator-digits")?.unwrap_or(0);
                let slots = attr(child, "max-denominator-value")
                    .or_else(|| child.attribute((NS_LOEXT, "max-denominator-value")))
                    .and_then(|v| v.trim().parse::<u64>().ok())
                    .filter(|max| *max > 0)
                    .map_or(min_digits, |max| max.to_string().len() as u32)
                    .min(30);
                // LibreOffice shows a fraction that names no numerator
                // digits, or no denominator at all, in its standard format
                // (measured over the combinations of the four attributes).
                let numerator = attr_u32(child, "min-numerator-digits")?;
                let (Some(numerator), true) = (numerator, fixed.is_some() || slots > 0) else {
                    code.push_str("General");
                    general = children.len() == 1;
                    continue;
                };
                if let Some(min_int) = attr_u32(child, "min-integer-digits")? {
                    let grouping = attr(child, "grouping") == Some("true");
                    code.push_str(&integer_digits(child, min_int, grouping)?);
                    code.push(' ');
                }
                code.push_str(&"?".repeat(numerator.max(1) as usize));
                code.push('/');
                match fixed {
                    Some(value) => code.push_str(value),
                    None => {
                        let shown = min_digits.min(slots);
                        code.push_str(&"#".repeat((slots - shown) as usize));
                        code.push_str(&"?".repeat(shown as usize));
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
            // §16.29.18 / §16.29.17: the formatter's calendar-field brackets.
            // LibreOffice shows the short quarter as `Q1` and the long one
            // as the locale's name (`1st quarter`).
            "quarter" => code.push_str(if long(child) { "[QQ]" } else { "[Q]" }),
            "week-of-year" => match week_rule(attr(style, "language"), attr(style, "country")) {
                (0, 1) => code.push_str("[WW]"),
                (first_day, min_days) => {
                    code.push_str(&format!("[WW:{first_day},{min_days}]"));
                }
            },
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
                // A minutes field that is the style's largest component and
                // does not truncate (§19.369) counts all minutes: `[mm]`.
                // Any other is written `m` / `mm`, which reads as minutes
                // beside hours or seconds and as the month elsewhere. That
                // is how LibreOffice shows a minutes field standing alone:
                // its format code for it is `MM`, and PT1H30M displays as
                // 12, the month of the null date (measured with
                // LibreOffice 26.8; §16.29.21 would give 30).
                if duration && first_time_field {
                    code.push('[');
                    code.push_str(token);
                    code.push(']');
                } else {
                    code.push_str(token);
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
                // The workbook formatter shows seconds to the millisecond
                // (three places); further places have no token.
                if let Some(places) = attr_u32(child, "decimal-places")?.filter(|p| *p > 0) {
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
        return Ok((
            match kind {
                DataKind::Text => "@".into(),
                _ => "General".into(),
            },
            true,
            1.0,
        ));
    }
    Ok((code, general, divisor))
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
        data.add(document.root_element()).unwrap();
        data
    }

    #[test]
    fn numbers_without_places_fractions_and_engineering_follow_libreoffice() {
        let mut data = codes(concat!(
            // No decimal places: the standard format, whatever else is set.
            r##"<number:number-style style:name="G"><number:number number:min-integer-digits="3" number:grouping="true"/></number:number-style>"##,
            r##"<number:currency-style style:name="CG"><number:currency-symbol>$</number:currency-symbol><number:number number:min-integer-digits="1"/></number:currency-style>"##,
            // `decimal-replacement=""`: optional decimals.
            r##"<number:number-style style:name="R"><number:number number:decimal-places="2" number:min-integer-digits="1" number:decimal-replacement=""/></number:number-style>"##,
            // Embedded text keeps a digit slot on its left.
            r##"<number:number-style style:name="E"><number:number number:decimal-places="2" number:min-integer-digits="1"><number:embedded-text number:position="2">-</number:embedded-text></number:number></number:number-style>"##,
            r##"<number:number-style style:name="X"><number:scientific-number number:decimal-places="2" number:min-decimal-places="0" number:min-integer-digits="1" number:min-exponent-digits="1" number:exponent-interval="3"/></number:number-style>"##,
            r##"<number:number-style style:name="F1"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="1" number:max-denominator-value="999"/></number:number-style>"##,
            r##"<number:number-style style:name="F2"><number:fraction number:min-numerator-digits="2" number:min-denominator-digits="3" number:max-denominator-value="9"/></number:number-style>"##,
            r##"<number:number-style style:name="F3"><number:fraction number:min-integer-digits="1" number:min-numerator-digits="1" number:max-denominator-value="99"/></number:number-style>"##,
            r##"<number:number-style style:name="F4"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:denominator-value="16"/></number:number-style>"##,
            r##"<number:number-style style:name="F5"><number:fraction number:min-integer-digits="0" number:min-numerator-digits="1"/></number:number-style>"##,
            r##"<number:number-style style:name="F6"><number:fraction number:min-integer-digits="0" number:min-denominator-digits="2"/></number:number-style>"##,
        ));
        let mut code = |name: &str| data.code(name).unwrap().0.code;
        assert_eq!(code("G"), "General");
        assert_eq!(code("CG"), "$General");
        assert_eq!(code("R"), "0.##");
        assert_eq!(code("E"), "#-#0.00");
        assert_eq!(code("X"), "##0.##E+0");
        assert_eq!(code("F1"), "# ?/##?");
        assert_eq!(code("F2"), "??/?");
        assert_eq!(code("F3"), "0 ?/##");
        assert_eq!(code("F4"), "# ?/16");
        assert_eq!(code("F5"), "General");
        assert_eq!(code("F6"), "General");
        assert!(data.is_general("G") && data.is_general("F5"));
    }

    #[test]
    fn counts_literal_text_styles_and_composed_markers() {
        let mut data = codes(concat!(
            // 31 integer digits are 31 digits, not 30.
            r##"<number:number-style style:name="N31"><number:number number:decimal-places="0" number:min-integer-digits="31"/></number:number-style>"##,
            // A text style of literal text only.
            r##"<number:text-style style:name="LIT"><number:text>MASKED</number:text></number:text-style>"##,
            r##"<number:text-style style:name="AT"><number:text>a@b </number:text><number:text-content/></number:text-style>"##,
            // A minutes field alone: the month, as LibreOffice shows it;
            // all minutes when the style does not truncate.
            r##"<number:time-style style:name="MIN"><number:minutes number:style="long"/></number:time-style>"##,
            r##"<number:time-style style:name="MINX" number:truncate-on-overflow="false"><number:minutes number:style="long"/></number:time-style>"##,
            // A locale-ordered short date as the target of a condition.
            r##"<number:date-style style:name="SHORT" number:automatic-order="true"><number:month/><number:text>/</number:text><number:day/><number:text>/</number:text><number:year/></number:date-style>"##,
            r##"<number:date-style style:name="COND"><number:year number:style="long"/><style:map style:condition="value()&gt;=0" style:apply-style-name="SHORT"/></number:date-style>"##,
        ));
        let mut code = |name: &str| data.code(name).unwrap().0.code;
        assert_eq!(code("N31"), "0".repeat(31));
        assert_eq!(code("LIT"), "General;General;General;\"MASKED\"");
        assert_eq!(code("AT"), "\"a@b \"@");
        assert_eq!(code("MIN"), "mm");
        assert_eq!(code("MINX"), "[mm]");
        assert_eq!(code("COND"), "mm-dd-yy;yyyy");
        // A count no code could spell rejects the document.
        let xml = format!(
            r##"<office:automatic-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="{NS_STYLE}" xmlns:number="{NS_NUMBER}"><number:number-style style:name="BIG"><number:number number:decimal-places="4000000000" number:min-integer-digits="1"/></number:number-style></office:automatic-styles>"##
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        assert_eq!(
            DataStyles::default().add(document.root_element()),
            Err(crate::error::resource("number format code"))
        );
    }

    #[test]
    fn quarters_and_weeks_are_calendar_fields() {
        let mut data = codes(concat!(
            r##"<number:date-style style:name="Q"><number:text>Q</number:text><number:quarter/><number:text> </number:text><number:quarter number:style="long"/></number:date-style>"##,
            r##"<number:date-style style:name="W"><number:week-of-year/></number:date-style>"##,
            r##"<number:date-style style:name="WDE" number:language="de" number:country="DE"><number:week-of-year/></number:date-style>"##,
            r##"<number:date-style style:name="WCH" number:language="de" number:country="XX"><number:week-of-year/></number:date-style>"##,
            r##"<number:date-style style:name="WIN" number:language="bn" number:country="IN"><number:week-of-year/></number:date-style>"##,
            r##"<number:date-style style:name="WBD" number:language="bn" number:country="BD"><number:week-of-year/></number:date-style>"##,
            r##"<number:date-style style:name="WDEX" number:language="de"><number:week-of-year/></number:date-style>"##,
            r##"<number:date-style style:name="WZH" number:language="zh"><number:week-of-year/></number:date-style>"##,
        ));
        assert_eq!(data.code("Q").unwrap().0.code, "\"Q\"[Q] [QQ]");
        let mut week = |name: &str| data.code(name).unwrap().0.code;
        assert_eq!(week("W"), "[WW]");
        assert_eq!(week("WDE"), "[WW:1,4]");
        // An unknown country falls back within the language.
        assert_eq!(week("WCH"), "[WW:1,4]");
        // A listed locale overrides its language's fallback.
        assert_eq!(week("WIN"), "[WW]");
        assert_eq!(week("WBD"), "[WW:6,1]");
        // A language without a country is LibreOffice's default unless
        // listed alone.
        assert_eq!(week("WDEX"), "[WW]");
        assert_eq!(week("WZH"), "[WW:1,4]");
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
        assert_eq!(data.code("N0").unwrap().0.code, "General");
        assert!(data.is_general("N0"));
        assert_eq!(data.code("N1").unwrap().0.code, "#,##0.00");
        assert_eq!(data.code("P").unwrap().0.code, "0.0%");
        assert_eq!(
            data.code("C0").unwrap().0.code,
            "\"€\" #,##0.00;[Red]-\"€\" #,##0.00"
        );
        assert_eq!(data.code("S").unwrap().0.code, "0.00E+00");
        assert_eq!(data.code("F").unwrap().0.code, "# ?/??");
        assert_eq!(data.code("T").unwrap().0.code, "@");
    }

    #[test]
    fn dates_times_and_durations() {
        let mut data = codes(concat!(
            r##"<number:date-style style:name="D"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>"##,
            r##"<number:date-style style:name="D2"><number:day/><number:text>-</number:text><number:month number:textual="true"/><number:text> </number:text><number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text> </number:text><number:am-pm/></number:date-style>"##,
            r##"<number:time-style style:name="T" number:truncate-on-overflow="false"><number:hours/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long" number:decimal-places="2"/></number:time-style>"##,
            r##"<number:date-style style:name="J"><number:era number:style="long" number:calendar="gengou"/><number:year number:calendar="gengou"/><number:text>年</number:text></number:date-style>"##,
        ));
        assert_eq!(data.code("D").unwrap().0.code, "yyyy-mm-dd");
        assert_eq!(data.code("D2").unwrap().0.code, "d-mmm h:mm AM/PM");
        assert_eq!(data.code("T").unwrap().0.code, "[h]:mm:ss.00");
        assert_eq!(data.code("J").unwrap().0.code, "[$-411]ggge\"年\"");
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
        assert_eq!(data.code("Z").unwrap().0.code, "0;(0);-");
        assert_eq!(data.code("TX").unwrap().0.code, "\\\"@");
        assert_eq!(data.code("ALL").unwrap().0.code, "0;(0);General;@");
    }

    #[test]
    fn display_factors_divide_each_section() {
        let mut data = codes(concat!(
            r##"<number:number-style style:name="K"><number:number number:decimal-places="1" number:min-integer-digits="1" number:display-factor="1000"/></number:number-style>"##,
            r##"<number:number-style style:name="TEN"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="10"/></number:number-style>"##,
            r##"<number:number-style style:name="G"><number:number number:display-factor="2000"/></number:number-style>"##,
            r##"<number:number-style style:name="NEG"><number:text>(</number:text><number:number number:min-integer-digits="1" number:decimal-places="0"/><number:text>)</number:text></number:number-style>"##,
            r##"<number:number-style style:name="MIX"><number:number number:min-integer-digits="1" number:decimal-places="0"/><style:map style:condition="value()&gt;=0" style:apply-style-name="TEN"/></number:number-style>"##,
        ));
        data.check().unwrap();
        let (k, _) = data.code("K").unwrap();
        assert_eq!(
            (k.code.as_str(), k.divisors.as_slice()),
            ("0.0", &[1000.0][..])
        );
        assert_eq!(k.chart_format(), ("0.0".into(), Some(vec![1000.0])));
        let (ten, _) = data.code("TEN").unwrap();
        assert_eq!(
            (ten.code.as_str(), ten.divisors.as_slice()),
            ("0", &[10.0][..])
        );
        // Charts divide by any factor, as cells do.
        assert_eq!(ten.chart_format(), ("0".into(), Some(vec![10.0])));
        let (general, _) = data.code("G").unwrap();
        assert_eq!(general.divisors, [2000.0]);
        // A General code keeps its factor on chart axes and labels.
        assert_eq!(
            general.chart_format(),
            ("General".into(), Some(vec![2000.0]))
        );
        assert!(!data.is_general("G"));
        let (mix, _) = data.code("MIX").unwrap();
        assert_eq!(
            (mix.code.as_str(), mix.divisors.as_slice()),
            ("0;0", &[10.0, 1.0][..])
        );
    }

    #[test]
    fn non_finite_display_factors_terminate_and_do_not_scale() {
        let mut data = codes(concat!(
            r##"<number:number-style style:name="INF"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="INF"/></number:number-style>"##,
            r##"<number:number-style style:name="NAN"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="NaN"/></number:number-style>"##,
            r##"<number:number-style style:name="ZERO"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="0"/></number:number-style>"##,
            // Negative: LibreOffice shows the value unscaled.
            r##"<number:number-style style:name="NEG"><number:number number:decimal-places="0" number:min-integer-digits="1" number:display-factor="-10"/></number:number-style>"##,
        ));
        data.check().unwrap();
        for name in ["INF", "NAN", "ZERO", "NEG"] {
            let (format, _) = data.code(name).unwrap();
            assert_eq!(format, FormatCode::plain("0"), "{name}");
        }
    }

    #[test]
    fn styles_beyond_a_spreadsheet_code_reject() {
        let maps: String = (0..4)
            .map(|i| {
                format!(
                    r#"<style:map style:condition="value()&gt;{i}" style:apply-style-name="P"/>"#
                )
            })
            .collect();
        let mut data = codes(&format!(
            r##"<number:number-style style:name="P"><number:number number:min-integer-digits="1"/></number:number-style><number:number-style style:name="M"><number:number number:min-integer-digits="1"/>{maps}</number:number-style>"##
        ));
        assert_eq!(
            data.check(),
            Err(crate::error::resource("number format section"))
        );
        let long = "x".repeat(MAX_CODE_CHARS + 1);
        let mut data = codes(&format!(
            r##"<number:number-style style:name="L"><number:number number:min-integer-digits="1"/><number:text>{long}</number:text></number:number-style>"##
        ));
        assert_eq!(
            data.check(),
            Err(crate::error::resource("number format code"))
        );
    }

    #[test]
    fn map_conditions_resolve_their_namespace_prefix() {
        let mut data = codes(concat!(
            r##"<number:number-style style:name="P"><number:number number:decimal-places="0" number:min-integer-digits="1"/></number:number-style>"##,
            r##"<number:number-style style:name="OF" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"><number:text>-</number:text><number:number number:decimal-places="0" number:min-integer-digits="1"/><style:map style:condition="of:value()&gt;=0" style:apply-style-name="P"/></number:number-style>"##,
            r##"<number:number-style style:name="X" xmlns:x="urn:other"><number:text>-</number:text><number:number number:decimal-places="0" number:min-integer-digits="1"/><style:map style:condition="x:value()&gt;=0" style:apply-style-name="P"/></number:number-style>"##,
            r##"<number:number-style style:name="BAD"><number:text>-</number:text><number:number number:decimal-places="0" number:min-integer-digits="1"/><style:map style:condition="value()" style:apply-style-name="P"/></number:number-style>"##,
        ));
        assert_eq!(data.code("OF").unwrap().0.code, "0;-0");
        // A condition in another namespace, or none §16.3 defines, is
        // ignored: it never becomes an unconditional section.
        assert_eq!(data.code("X").unwrap().0.code, "-0");
        assert_eq!(data.code("BAD").unwrap().0.code, "-0");
    }
}
