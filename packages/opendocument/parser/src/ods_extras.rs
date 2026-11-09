//! Workbook features beside the cell grid: named ranges and expressions
//! (ODF Part 3 §9.4), database ranges with filter buttons (§9.4.1), content
//! validations (§9.4.4), LibreOffice's conditional formats and sparklines
//! (`calcext`, an extension namespace of ODF 1.3 producers), and the
//! conversion of OpenFormula text (ODF Part 4) to SpreadsheetML formula text.

use super::{parse_range, sheet_prefix, split_sheet, Context, OdsSheet, NS_TABLE};
use crate::odf_props::NS_CALCEXT;
use crate::xml;
use xlsx_model::{CellRange, CfRule, CfStop, CfValue, ConditionalFormat, DataValidation};

/// OpenFormula (`of:=SUM([.A1:.B2];[$Sheet2.C3])`) as SpreadsheetML formula
/// text (`SUM(A1:B2,Sheet2!C3)`): references in brackets become A1 references
/// with a `Sheet!` prefix, `;` parameter separators become `,`.
pub fn convert_formula(formula: &str) -> String {
    let body = formula
        .split_once(':')
        .filter(|(ns, _)| !ns.contains(['(', '"', '[']) && ns.len() <= 8)
        .map_or(formula, |(_, rest)| rest);
    let body = body.strip_prefix('=').unwrap_or(body);
    let mut out = String::with_capacity(body.len());
    let mut chars = body.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                out.push('"');
                for d in chars.by_ref() {
                    out.push(d);
                    if d == '"' {
                        break;
                    }
                }
            }
            '[' => {
                let mut reference = String::new();
                for d in chars.by_ref() {
                    if d == ']' {
                        break;
                    }
                    reference.push(d);
                }
                out.push_str(&convert_reference(&reference));
            }
            ';' => out.push(','),
            _ => out.push(c),
        }
    }
    out
}

/// One bracketed OpenFormula reference (`$Sheet1.A1:.B2`).
fn convert_reference(reference: &str) -> String {
    let mut parts = Vec::new();
    let mut sheet_seen: Option<String> = None;
    for (i, part) in split_colon(reference).into_iter().enumerate() {
        let (sheet, cell) = split_sheet(part);
        let prefix = match (&sheet, i) {
            (Some(name), 0) => {
                sheet_seen = Some(name.clone());
                sheet_prefix(name)
            }
            (Some(name), _) if Some(name) != sheet_seen.as_ref() => sheet_prefix(name),
            _ => String::new(),
        };
        parts.push(format!("{prefix}{cell}"));
    }
    parts.join(":")
}

fn split_colon(reference: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut quoted = false;
    let mut start = 0;
    for (i, c) in reference.char_indices() {
        match c {
            '\'' => quoted = !quoted,
            ':' if !quoted => {
                out.push(&reference[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    out.push(&reference[start..]);
    out
}

/// A range address list (`Sheet1.A1:Sheet1.B2 Sheet1.D4`) as a SpreadsheetML
/// reference list with absolute references (`Sheet1!$A$1:$B$2,Sheet1!$D$4`).
pub fn convert_range_list(addresses: &str) -> Option<String> {
    let mut out = Vec::new();
    for address in crate::odf_text::split_addresses(addresses) {
        let (sheet, (r1, c1), (r2, c2)) = parse_range(address)?;
        let prefix = sheet.as_deref().map(sheet_prefix).unwrap_or_default();
        let a = format!("${}${}", super::column_letters(c1), r1);
        let b = format!("${}${}", super::column_letters(c2), r2);
        out.push(if (r1, c1) == (r2, c2) {
            format!("{prefix}{a}")
        } else {
            format!("{prefix}{a}:{b}")
        });
    }
    (!out.is_empty()).then(|| out.join(","))
}

fn range_of(address: &str) -> Option<(Option<String>, CellRange)> {
    let (sheet, (top, left), (bottom, right)) = parse_range(address)?;
    Some((
        sheet,
        CellRange {
            top,
            left,
            bottom,
            right,
        },
    ))
}

/// `table:named-expressions` (§9.4.11): global ones under the spreadsheet,
/// sheet-local ones under a table.
pub fn named_expressions(
    spreadsheet: roxmltree::Node<'_, '_>,
    sheet_names: &[String],
) -> Vec<(String, Option<usize>, String)> {
    let mut names = Vec::new();
    let mut read = |container: roxmltree::Node<'_, '_>, scope: Option<usize>| {
        for node in container.children().filter(|n| n.is_element()) {
            let Some(name) = node.attribute((NS_TABLE, "name")) else {
                continue;
            };
            let formula = if xml::is(node, NS_TABLE, "named-range") {
                node.attribute((NS_TABLE, "cell-range-address"))
                    .and_then(convert_range_list)
            } else if xml::is(node, NS_TABLE, "named-expression") {
                node.attribute((NS_TABLE, "expression"))
                    .map(convert_formula)
            } else {
                None
            };
            if let Some(formula) = formula {
                names.push((name.to_string(), scope, formula));
            }
        }
    };
    for child in spreadsheet.children().filter(|n| n.is_element()) {
        if xml::is(child, NS_TABLE, "named-expressions") {
            read(child, None);
        } else if xml::is(child, NS_TABLE, "table") {
            let index = sheet_names
                .iter()
                .position(|n| Some(n.as_str()) == child.attribute((NS_TABLE, "name")));
            for local in child
                .children()
                .filter(|n| xml::is(*n, NS_TABLE, "named-expressions"))
            {
                read(local, index);
            }
        }
    }
    names
}

/// `table:database-range` (§9.4.1): named ranges, and the filter buttons of
/// a range with `table:display-filter-buttons` as the sheet's auto filter.
pub fn database_ranges(
    spreadsheet: roxmltree::Node<'_, '_>,
    sheets: &mut [OdsSheet],
    sheet_names: &[String],
    names: &mut Vec<(String, Option<usize>, String)>,
) {
    for ranges in spreadsheet
        .children()
        .filter(|n| xml::is(*n, NS_TABLE, "database-ranges"))
    {
        for range in ranges
            .children()
            .filter(|n| xml::is(*n, NS_TABLE, "database-range"))
        {
            let Some(address) = range.attribute((NS_TABLE, "target-range-address")) else {
                continue;
            };
            let Some((sheet, cells)) = range_of(address) else {
                continue;
            };
            let index = sheet
                .as_deref()
                .and_then(|name| sheet_names.iter().position(|n| n == name));
            if let Some(name) = range.attribute((NS_TABLE, "name")) {
                if !name.starts_with("__Anonymous_Sheet_DB__") {
                    if let Some(formula) = convert_range_list(address) {
                        names.push((name.to_string(), None, formula));
                    }
                }
            }
            if range.attribute((NS_TABLE, "display-filter-buttons")) == Some("true") {
                if let Some(sheet) = index.and_then(|i| sheets.get_mut(i)) {
                    sheet.worksheet.auto_filter = Some(cells);
                }
            }
        }
    }
}

/// `table:content-validation` (§19.595 table:condition) as SpreadsheetML
/// data validations over the cells that name them.
pub fn content_validations(spreadsheet: roxmltree::Node<'_, '_>, sheets: &mut [OdsSheet]) {
    let Some(container) = spreadsheet
        .children()
        .find(|n| xml::is(*n, NS_TABLE, "content-validations"))
    else {
        return;
    };
    for sheet in sheets.iter_mut() {
        let refs = std::mem::take(&mut sheet.validation_refs);
        if refs.is_empty() {
            continue;
        }
        for node in container
            .children()
            .filter(|n| xml::is(*n, NS_TABLE, "content-validation"))
        {
            let Some(name) = node.attribute((NS_TABLE, "name")) else {
                continue;
            };
            let sqref: Vec<String> = refs
                .iter()
                .filter(|(n, _)| n == name)
                .map(|(_, r)| range_text(r))
                .collect();
            if sqref.is_empty() {
                continue;
            }
            let mut validation = DataValidation {
                sqref: sqref.join(" "),
                validation_type: None,
                operator: None,
                formula1: None,
                formula2: None,
                allow_blank: node.attribute((NS_TABLE, "allow-empty-cell")) != Some("false"),
                prompt_title: None,
                prompt: None,
                error_title: None,
                error_message: None,
            };
            if let Some(condition) = node.attribute((NS_TABLE, "condition")) {
                read_validation_condition(&convert_namespace(condition), &mut validation);
            }
            if let Some(help) = node
                .children()
                .find(|n| xml::is(*n, NS_TABLE, "help-message"))
            {
                validation.prompt_title = help.attribute((NS_TABLE, "title")).map(str::to_string);
                validation.prompt = Some(super::paragraph_text(help));
            }
            if let Some(error) = node
                .children()
                .find(|n| xml::is(*n, NS_TABLE, "error-message"))
            {
                validation.error_title = error.attribute((NS_TABLE, "title")).map(str::to_string);
                validation.error_message = Some(super::paragraph_text(error));
            }
            sheet.worksheet.data_validations.push(validation);
        }
    }
}

fn convert_namespace(condition: &str) -> String {
    condition
        .strip_prefix("of:")
        .or_else(|| condition.strip_prefix("ooo:"))
        .unwrap_or(condition)
        .to_string()
}

fn range_text(range: &CellRange) -> String {
    let a = format!("{}{}", super::column_letters(range.left), range.top);
    if range.top == range.bottom && range.left == range.right {
        a
    } else {
        format!("{a}:{}{}", super::column_letters(range.right), range.bottom)
    }
}

/// Arguments of `name(...)` within `text` (top-level `;` separated).
fn call_args<'a>(text: &'a str, name: &str) -> Option<Vec<&'a str>> {
    let start = text.find(&format!("{name}("))? + name.len() + 1;
    let mut depth = 0i32;
    let mut quoted = false;
    let mut args = Vec::new();
    let mut from = start;
    for (offset, c) in text[start..].char_indices() {
        let i = start + offset;
        match c {
            '"' => quoted = !quoted,
            '(' if !quoted => depth += 1,
            ')' if !quoted => {
                if depth == 0 {
                    args.push(text[from..i].trim());
                    return Some(args);
                }
                depth -= 1;
            }
            ';' | ',' if !quoted && depth == 0 => {
                args.push(text[from..i].trim());
                from = i + 1;
            }
            _ => {}
        }
    }
    None
}

fn read_validation_condition(condition: &str, validation: &mut DataValidation) {
    if let Some(args) = call_args(condition, "cell-content-is-in-list") {
        validation.validation_type = Some("list".into());
        let joined = args.join(";");
        validation.formula1 = Some(if joined.starts_with('"') {
            let items: Vec<String> = args
                .iter()
                .map(|a| a.trim_matches('"').replace("\"\"", "\""))
                .collect();
            format!("\"{}\"", items.join(","))
        } else {
            convert_formula(&joined)
        });
        return;
    }
    validation.validation_type = Some(
        if condition.contains("cell-content-is-whole-number") {
            "whole"
        } else if condition.contains("cell-content-is-decimal-number") {
            "decimal"
        } else if condition.contains("cell-content-is-date") {
            "date"
        } else if condition.contains("cell-content-is-time") {
            "time"
        } else if condition.contains("cell-content-text-length") {
            "textLength"
        } else if condition.contains("is-true-formula") {
            "custom"
        } else {
            "decimal"
        }
        .into(),
    );
    if let Some(args) = call_args(condition, "is-true-formula") {
        validation.formula1 = args.first().map(|a| convert_formula(a));
        return;
    }
    for (name, operator) in [
        ("cell-content-is-not-between", "notBetween"),
        ("cell-content-is-between", "between"),
        ("cell-content-text-length-is-not-between", "notBetween"),
        ("cell-content-text-length-is-between", "between"),
    ] {
        if let Some(args) = call_args(condition, name) {
            validation.operator = Some(operator.into());
            validation.formula1 = args.first().map(|a| convert_formula(a));
            validation.formula2 = args.get(1).map(|a| convert_formula(a));
            return;
        }
    }
    let tail = condition
        .rsplit_once("cell-content()")
        .or_else(|| condition.rsplit_once("cell-content-text-length()"))
        .map(|(_, rest)| rest.trim());
    if let Some(rest) = tail {
        for (op, name) in [
            (">=", "greaterThanOrEqual"),
            ("<=", "lessThanOrEqual"),
            ("!=", "notEqual"),
            ("<>", "notEqual"),
            (">", "greaterThan"),
            ("<", "lessThan"),
            ("=", "equal"),
        ] {
            if let Some(value) = rest.strip_prefix(op) {
                validation.operator = Some(name.into());
                validation.formula1 = Some(convert_formula(value.trim()));
                return;
            }
        }
    }
}

/// LibreOffice conditional formats (`calcext:conditional-format`) and
/// sparkline groups (`calcext:sparkline-group`) of each table.
pub(crate) fn conditional_formats(
    ctx: &mut Context,
    tables: &[roxmltree::Node<'_, '_>],
    sheets: &mut [OdsSheet],
) -> Result<(), String> {
    for (index, table) in tables.iter().enumerate() {
        let mut formats = Vec::new();
        let mut priority = 1i32;
        for container in table
            .children()
            .filter(|n| xml::is(*n, NS_CALCEXT, "conditional-formats"))
        {
            for format in container
                .children()
                .filter(|n| xml::is(*n, NS_CALCEXT, "conditional-format"))
            {
                let Some(target) = format.attribute((NS_CALCEXT, "target-range-address")) else {
                    continue;
                };
                let sqref: Vec<CellRange> = crate::odf_text::split_addresses(target)
                    .into_iter()
                    .filter_map(range_of)
                    .map(|(_, r)| r)
                    .collect();
                if sqref.is_empty() {
                    continue;
                }
                let anchor = format!("{}{}", super::column_letters(sqref[0].left), sqref[0].top);
                let mut rules = Vec::new();
                for rule in format.children().filter(|n| n.is_element()) {
                    if let Some(rule) = cf_rule(ctx, rule, priority, &anchor) {
                        rules.push(rule);
                        priority += 1;
                    }
                }
                if !rules.is_empty() {
                    formats.push(ConditionalFormat { sqref, rules });
                }
            }
        }
        if let Some(sheet) = sheets.get_mut(index) {
            sheet.worksheet.conditional_formats = formats;
        }
    }
    // Sparklines read cell values from every table.
    for (index, table) in tables.iter().enumerate() {
        let groups = sparkline_groups(*table, sheets as &[OdsSheet]);
        if let Some(sheet) = sheets.get_mut(index) {
            sheet.worksheet.sparkline_groups = groups;
        }
    }
    Ok(())
}

fn cf_value(node: roxmltree::Node<'_, '_>) -> CfValue {
    let kind = match node.attribute((NS_CALCEXT, "type")).unwrap_or("number") {
        "minimum" => "min",
        "maximum" => "max",
        "percentile" => "percentile",
        "percent" => "percent",
        "formula" => "formula",
        "auto-minimum" => "autoMin",
        "auto-maximum" => "autoMax",
        _ => "num",
    };
    CfValue {
        kind: kind.into(),
        value: node
            .attribute((NS_CALCEXT, "value"))
            .filter(|_| !matches!(kind, "min" | "max" | "autoMin" | "autoMax"))
            .map(convert_formula),
    }
}

fn hex(value: Option<&str>) -> String {
    value
        .and_then(crate::odf_props::hex_color)
        .unwrap_or_else(|| "000000".into())
}

fn cf_rule(
    ctx: &mut Context,
    node: roxmltree::Node<'_, '_>,
    priority: i32,
    anchor: &str,
) -> Option<CfRule> {
    match node.tag_name().name() {
        "color-scale" => {
            let stops = node
                .children()
                .filter(|n| xml::is(*n, NS_CALCEXT, "color-scale-entry"))
                .map(|entry| {
                    let value = cf_value(entry);
                    CfStop {
                        kind: value.kind,
                        value: value.value,
                        color: hex(entry.attribute((NS_CALCEXT, "color"))),
                    }
                })
                .collect();
            Some(CfRule::ColorScale {
                stops,
                priority,
                active_formula: None,
                stop_if_true: false,
            })
        }
        "data-bar" => {
            let entries: Vec<CfValue> = node
                .children()
                .filter(|n| xml::is(*n, NS_CALCEXT, "formatting-entry"))
                .map(cf_value)
                .collect();
            let mut entries = entries.into_iter();
            Some(CfRule::DataBar {
                color: hex(node.attribute((NS_CALCEXT, "positive-color"))),
                min: entries.next().unwrap_or(CfValue {
                    kind: "min".into(),
                    value: None,
                }),
                max: entries.next().unwrap_or(CfValue {
                    kind: "max".into(),
                    value: None,
                }),
                priority,
                gradient: node.attribute((NS_CALCEXT, "gradient")) != Some("false"),
                active_formula: None,
                ext_threshold_formula: false,
                stop_if_true: false,
            })
        }
        "icon-set" => {
            let cfvos = node
                .children()
                .filter(|n| xml::is(*n, NS_CALCEXT, "formatting-entry"))
                .map(cf_value)
                .collect();
            Some(CfRule::IconSet {
                icon_set: node
                    .attribute((NS_CALCEXT, "icon-set-type"))
                    .unwrap_or("3TrafficLights1")
                    .to_string(),
                cfvos,
                reverse: false,
                priority,
                custom_icons: None,
                active_formula: None,
                stop_if_true: false,
            })
        }
        "condition" => {
            let style = node.attribute((NS_CALCEXT, "apply-style-name"))?;
            let dxf = Some(
                ctx.xfs
                    .dxf_for(&ctx.store, &mut ctx.data, &ctx.faces, style),
            );
            let value = node.attribute((NS_CALCEXT, "value"))?.trim();
            condition_rule(value, dxf, priority, anchor)
        }
        "date-is" => {
            let style = node.attribute((NS_CALCEXT, "style"))?;
            let _ = ctx
                .xfs
                .dxf_for(&ctx.store, &mut ctx.data, &ctx.faces, style);
            Some(CfRule::Other {
                kind: "timePeriod".into(),
                priority,
                unsupported_formula_phases: 0,
                stop_if_true: false,
            })
        }
        _ => None,
    }
}

/// One `calcext:condition` value (LibreOffice's condition grammar, which
/// follows ODF 1.2 `style:map` conditions).
fn condition_rule(value: &str, dxf_id: Option<u32>, priority: i32, anchor: &str) -> Option<CfRule> {
    let cell_is = |operator: &str, formulas: Vec<String>| CfRule::CellIs {
        operator: operator.into(),
        formulas,
        dxf_id,
        priority,
        stop_if_true: false,
    };
    let expression = |formula: String| CfRule::Expression {
        formula,
        dxf_id,
        priority,
        stop_if_true: false,
    };
    for (name, operator) in [("not-between", "notBetween"), ("between", "between")] {
        if value.starts_with(&format!("{name}(")) {
            let args = call_args(value, name)?;
            return Some(cell_is(
                operator,
                args.iter().map(|a| convert_formula(a)).collect(),
            ));
        }
    }
    for (op, operator) in [
        (">=", "greaterThanOrEqual"),
        ("<=", "lessThanOrEqual"),
        ("!=", "notEqual"),
        ("<>", "notEqual"),
        (">", "greaterThan"),
        ("<", "lessThan"),
        ("=", "equal"),
    ] {
        if let Some(rest) = value.strip_prefix(op) {
            return Some(cell_is(operator, vec![convert_formula(rest.trim())]));
        }
    }
    if let Some(args) = call_args(value, "formula-is") {
        return Some(expression(convert_formula(args.first()?)));
    }
    let text_rule = |name: &str, build: &dyn Fn(&str) -> String| -> Option<CfRule> {
        let args = call_args(value, name)?;
        let text = convert_formula(args.first()?);
        Some(expression(build(&text)))
    };
    if value.starts_with("begins-with(") {
        return text_rule("begins-with", &|t| format!("LEFT({anchor},LEN({t}))={t}"));
    }
    if value.starts_with("ends-with(") {
        return text_rule("ends-with", &|t| format!("RIGHT({anchor},LEN({t}))={t}"));
    }
    if value.starts_with("contains-text(") {
        return text_rule("contains-text", &|t| {
            format!("NOT(ISERROR(SEARCH({t},{anchor})))")
        });
    }
    if value.starts_with("not-contains-text(") {
        return text_rule("not-contains-text", &|t| {
            format!("ISERROR(SEARCH({t},{anchor}))")
        });
    }
    if value == "error" {
        return Some(expression(format!("ISERROR({anchor})")));
    }
    if value == "no-error" {
        return Some(expression(format!("NOT(ISERROR({anchor}))")));
    }
    let top = |name: &str, top: bool, percent: bool| -> Option<CfRule> {
        let args = call_args(value, name)?;
        Some(CfRule::Top10 {
            top,
            percent,
            rank: args.first()?.trim().parse().ok()?,
            dxf_id,
            priority,
            stop_if_true: false,
        })
    };
    if value.starts_with("top-elements(") {
        return top("top-elements", true, false);
    }
    if value.starts_with("bottom-elements(") {
        return top("bottom-elements", false, false);
    }
    if value.starts_with("top-percent(") {
        return top("top-percent", true, true);
    }
    if value.starts_with("bottom-percent(") {
        return top("bottom-percent", false, true);
    }
    let average = |above: bool, equal: bool| CfRule::AboveAverage {
        above_average: above,
        equal_average: equal,
        std_dev: None,
        dxf_id,
        priority,
        stop_if_true: false,
    };
    match value {
        "above-average" => return Some(average(true, false)),
        "below-average" => return Some(average(false, false)),
        "above-equal-average" => return Some(average(true, true)),
        "below-equal-average" => return Some(average(false, true)),
        "duplicate" | "unique" => {
            return Some(CfRule::Other {
                kind: if value == "duplicate" {
                    "duplicateValues"
                } else {
                    "uniqueValues"
                }
                .into(),
                priority,
                unsupported_formula_phases: 0,
                stop_if_true: false,
            })
        }
        _ => {}
    }
    Some(expression(convert_formula(value)))
}

/// `calcext:sparkline-group` (LibreOffice's sparkline model, which mirrors
/// the Office 2010 `x14:sparklineGroup`).
fn sparkline_groups(
    table: roxmltree::Node<'_, '_>,
    sheets: &[OdsSheet],
) -> Vec<xlsx_model::SparklineGroup> {
    let mut groups = Vec::new();
    for container in table
        .children()
        .filter(|n| xml::is(*n, NS_CALCEXT, "sparkline-groups"))
    {
        for group in container
            .children()
            .filter(|n| xml::is(*n, NS_CALCEXT, "sparkline-group"))
        {
            let a = |name: &str| group.attribute((NS_CALCEXT, name));
            let flag = |name: &str| a(name) == Some("true");
            let color = |name: &str| a(name).and_then(crate::odf_props::hex_color);
            let mut sparklines = Vec::new();
            for spark in group
                .descendants()
                .filter(|n| xml::is(*n, NS_CALCEXT, "sparkline"))
            {
                let Some((_, (row, col), _)) = spark
                    .attribute((NS_CALCEXT, "cell-address"))
                    .and_then(parse_range)
                else {
                    continue;
                };
                let values = spark
                    .attribute((NS_CALCEXT, "data-range"))
                    .and_then(|r| range_values(sheets, r))
                    .unwrap_or_default();
                sparklines.push(xlsx_model::Sparkline { row, col, values });
            }
            groups.push(xlsx_model::SparklineGroup {
                kind: match a("type") {
                    Some("column") => xlsx_model::SparklineType::Column,
                    Some("stacked") => xlsx_model::SparklineType::Stem,
                    _ => xlsx_model::SparklineType::Line,
                },
                markers: flag("markers"),
                high: flag("high"),
                low: flag("low"),
                first: flag("first"),
                last: flag("last"),
                negative: flag("negative"),
                display_x_axis: flag("display-x-axis"),
                display_empty_cells_as: match a("display-empty-cells-as") {
                    Some("zero") => "zero",
                    Some("span") => "span",
                    _ => "gap",
                }
                .into(),
                min_axis_type: a("min-axis-type").unwrap_or("individual").into(),
                max_axis_type: a("max-axis-type").unwrap_or("individual").into(),
                manual_min: a("manual-min").and_then(|v| v.parse().ok()),
                manual_max: a("manual-max").and_then(|v| v.parse().ok()),
                line_weight: a("line-width")
                    .and_then(crate::odf_props::length_pt)
                    .unwrap_or(0.75),
                color_series: color("color-series"),
                color_negative: color("color-negative"),
                color_axis: color("color-axis"),
                color_markers: color("color-markers"),
                color_first: color("color-first"),
                color_last: color("color-last"),
                color_high: color("color-high"),
                color_low: color("color-low"),
                sparklines,
            });
        }
    }
    groups
}

fn range_values(sheets: &[OdsSheet], address: &str) -> Option<Vec<Option<f64>>> {
    let (sheet, (r1, c1), (r2, c2)) = parse_range(address)?;
    let sheet = sheets
        .iter()
        .find(|s| sheet.as_deref().is_none_or(|n| n == s.name))?;
    if (u64::from(r2 - r1) + 1) * (u64::from(c2 - c1) + 1) > 1_048_576 {
        return None;
    }
    let mut out = Vec::new();
    for row in r1..=r2 {
        for col in c1..=c2 {
            out.push(
                match super::find_cell(&sheet.rows, row, col).map(|c| &c.value) {
                    Some(xlsx_model::CellValue::Number { number }) => Some(*number),
                    _ => None,
                },
            );
        }
    }
    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn openformula_references_and_separators() {
        assert_eq!(
            convert_formula("of:=SUM([.A1:.B2];[$Sheet2.C3])"),
            "SUM(A1:B2,Sheet2!C3)"
        );
        assert_eq!(
            convert_formula("of:=IF([.A1]=\"a;b\";1;0)"),
            "IF(A1=\"a;b\",1,0)"
        );
        assert_eq!(
            convert_formula("of:=['My Sheet'.$A$1:.$A$3]"),
            "'My Sheet'!$A$1:$A$3"
        );
        assert_eq!(
            convert_range_list("Sheet1.A1:Sheet1.L23 Sheet1.B5").as_deref(),
            Some("Sheet1!$A$1:$L$23,Sheet1!$B$5")
        );
    }

    #[test]
    fn conditions_and_validations() {
        match condition_rule("between(1;5)", Some(0), 1, "A1").unwrap() {
            CfRule::CellIs {
                operator, formulas, ..
            } => {
                assert_eq!(operator, "between");
                assert_eq!(formulas, ["1", "5"]);
            }
            other => panic!("{other:?}"),
        }
        match condition_rule("contains-text(\"x\")", Some(0), 1, "B2").unwrap() {
            CfRule::Expression { formula, .. } => {
                assert_eq!(formula, "NOT(ISERROR(SEARCH(\"x\",B2)))")
            }
            other => panic!("{other:?}"),
        }
        let mut v = DataValidation {
            sqref: String::new(),
            validation_type: None,
            operator: None,
            formula1: None,
            formula2: None,
            allow_blank: true,
            prompt_title: None,
            prompt: None,
            error_title: None,
            error_message: None,
        };
        read_validation_condition("cell-content-is-in-list(\"a\";\"b\")", &mut v);
        assert_eq!(v.validation_type.as_deref(), Some("list"));
        assert_eq!(v.formula1.as_deref(), Some("\"a,b\""));
        read_validation_condition(
            "cell-content-is-whole-number() and cell-content-is-between(1;10)",
            &mut v,
        );
        assert_eq!(
            (v.validation_type.as_deref(), v.operator.as_deref()),
            (Some("whole"), Some("between"))
        );
    }
}
