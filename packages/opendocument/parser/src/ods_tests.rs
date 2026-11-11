//! Workbook-level tests of the spreadsheet reader over synthetic packages.

use super::*;
use crate::package::tests::{manifest, package_bytes};
use crate::package::PackageLimits;
use ooxml_common::resource::OoxmlFormat;

const ODS: &str = "application/vnd.oasis.opendocument.spreadsheet";
const NS: &str = concat!(
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" "#,
    r#"xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0""#,
);

fn content(automatic: &str, body: &str) -> String {
    format!(
        r#"<office:document-content {NS} office:version="1.3"><office:automatic-styles>{automatic}</office:automatic-styles><office:body><office:spreadsheet>{body}</office:spreadsheet></office:body></office:document-content>"#
    )
}

fn styles(common: &str) -> String {
    format!(
        r#"<office:document-styles {NS} office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"/>{common}</office:styles></office:document-styles>"#
    )
}

/// Read a package of `files` (manifest and mimetype added).
fn read_files(files: &[(&str, &[u8])]) -> Result<OdsWorkbook, String> {
    let manifest = manifest(ODS, "");
    let mut all: Vec<(&str, &[u8])> = vec![("META-INF/manifest.xml", manifest.as_bytes())];
    all.extend_from_slice(files);
    let package = OdfPackage::open(
        package_bytes(ODS, &all),
        OoxmlFormat::Xlsx,
        PackageLimits::default(),
        &MEDIA_TYPES,
    )?;
    read(&package)
}

/// Read `content` and `styles`; the workbook (or its rejection) and the
/// bytes charged to its budget.
fn read_charging(content: &str, styles: &str) -> (Result<OdsWorkbook, String>, usize) {
    let manifest = manifest(ODS, "");
    let files: Vec<(&str, &[u8])> = vec![
        ("META-INF/manifest.xml", manifest.as_bytes()),
        ("content.xml", content.as_bytes()),
        ("styles.xml", styles.as_bytes()),
    ];
    let package = OdfPackage::open(
        package_bytes(ODS, &files),
        OoxmlFormat::Xlsx,
        PackageLimits::default(),
        &MEDIA_TYPES,
    )
    .unwrap();
    let budget = Budget::new();
    let workbook = read_within(&package, budget.clone());
    (workbook, budget.used())
}

/// Bytes charged reading `content` and `styles` (which must read).
fn charged(content: &str, styles: &str) -> usize {
    let (workbook, used) = read_charging(content, styles);
    workbook.unwrap();
    used
}

/// A styles part whose `Default` cell style uses font face `F`, declared
/// with a `family_len`-byte family name.
fn styles_with_default_family(family_len: usize, common: &str) -> String {
    format!(
        r#"<office:document-styles {NS} office:version="1.3"><office:font-face-decls><style:font-face style:name="F" svg:font-family="{}"/></office:font-face-decls><office:styles><style:style style:name="Default" style:family="table-cell"><style:text-properties style:font-name="F"/></style:style>{common}</office:styles></office:document-styles>"#,
        "f".repeat(family_len)
    )
}

fn read_text(content: &str, styles: Option<&str>) -> Result<OdsWorkbook, String> {
    let mut files: Vec<(&str, &[u8])> = vec![("content.xml", content.as_bytes())];
    if let Some(styles) = styles {
        files.push(("styles.xml", styles.as_bytes()));
    }
    read_files(&files)
}

fn table(rows: &str) -> String {
    format!(r#"<table:table table:name="Sheet1">{rows}</table:table>"#)
}

#[test]
fn a_package_without_styles_xml_reads() {
    let workbook = read_text(
        &content(
            "",
            &table(r#"<table:table-row><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row>"#),
        ),
        None,
    )
    .unwrap();
    let cell = &workbook.sheets[0].rows[0].cells[0];
    assert!(matches!(cell.value, xlsx_model::CellValue::Number { number } if number == 7.0));
}

#[test]
fn xf_zero_is_the_default_cell_style() {
    let workbook = read_text(
        &content(
            r##"<style:style style:name="red" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ff0000"/></style:style>"##,
            &table(r#"<table:table-row><table:table-cell table:style-name="red" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell></table:table-row>"#),
        ),
        Some(&styles("")),
    )
    .unwrap();
    let styles = workbook.styles.as_ref().unwrap();
    // xf 0 paints no fill; the red cell has its own xf.
    assert_eq!(
        styles.fills[styles.cell_xfs[0].fill_id as usize].pattern_type,
        "none"
    );
    let cells = &workbook.sheets[0].rows[0].cells;
    assert_ne!(cells[0].style_index, Some(0));
    assert_eq!(cells[1].style_index, Some(0));
}

#[test]
fn repeated_links_cover_every_cell_and_are_budgeted() {
    let linked = |url: &str, cols: u32, rows: u32| {
        content(
            "",
            &table(&format!(
                r#"<table:table-row table:number-rows-repeated="{rows}"><table:table-cell table:number-columns-repeated="{cols}" office:value-type="string"><text:p><text:a xlink:href="{url}">x</text:a></text:p></table:table-cell></table:table-row>"#
            )),
        )
    };
    let workbook = read_text(&linked("https://example.test/", 300, 2), Some(&styles(""))).unwrap();
    // Every repeated cell carries the link: nothing is cut at 256.
    assert_eq!(workbook.sheets[0].worksheet.hyperlinks.len(), 600);
    // 256 × 256 copies of a 100 kB URL exceed the model budget.
    let url = format!("https://example.test/{}", "a".repeat(100_000));
    let error = read_text(&linked(&url, 256, 256), Some(&styles("")))
        .err()
        .unwrap();
    assert_eq!(error, resource("model size"));
}

#[test]
fn package_references_resolve_as_iris() {
    let png = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR\0\0\0\x01\0\0\0\x01\x08\x06\0\0\0\x1f\x15\xc4\x89";
    let chart_content = format!(
        r#"<office:document-content {NS} office:version="1.3"><office:body><office:chart><chart:chart chart:class="chart:line"/></office:chart></office:body></office:document-content>"#
    );
    let body = table(concat!(
        r#"<table:table-row><table:table-cell>"#,
        r#"<draw:frame svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/a%20b.png"/></draw:frame>"#,
        r#"<draw:frame svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object%201"/></draw:frame>"#,
        r#"</table:table-cell></table:table-row>"#
    ));
    let content = content("", &body);
    let styles = styles("");
    let workbook = read_files(&[
        ("content.xml", content.as_bytes()),
        ("styles.xml", styles.as_bytes()),
        ("Pictures/a b.png", png),
        ("Object 1/content.xml", chart_content.as_bytes()),
    ])
    .unwrap();
    let worksheet = &workbook.sheets[0].worksheet;
    assert_eq!(worksheet.images.len(), 1);
    assert_eq!(worksheet.images[0].image_path, "Pictures/a b.png");
    assert_eq!(worksheet.charts.len(), 1);
    assert!(worksheet.charts[0].chart.chart_type.starts_with("line"));
}

#[test]
fn standard_cell_style_conditions_apply() {
    let automatic = concat!(
        r#"<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default">"#,
        r#"<style:map style:condition="cell-content()&gt;0" style:apply-style-name="Good" style:base-cell-address="Sheet1.B2"/>"#,
        r#"<style:map style:condition="is-true-formula(of:[.A1]&gt;5)" style:apply-style-name="Good" style:base-cell-address="Sheet1.A1"/>"#,
        r#"</style:style>"#
    );
    let common = r##"<style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#00ff00"/></style:style>"##;
    let body = table(concat!(
        r#"<table:table-row><table:table-cell/></table:table-row>"#,
        r#"<table:table-row table:number-rows-repeated="2"><table:table-cell/><table:table-cell table:style-name="ce1" office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row>"#
    ));
    let workbook = read_text(&content(automatic, &body), Some(&styles(common))).unwrap();
    let formats = &workbook.sheets[0].worksheet.conditional_formats;
    assert_eq!(formats.len(), 1);
    let format = &formats[0];
    assert_eq!(
        (
            format.sqref[0].top,
            format.sqref[0].left,
            format.sqref[0].bottom,
            format.sqref[0].right
        ),
        (2, 2, 3, 2)
    );
    match &format.rules[0] {
        xlsx_model::CfRule::CellIs {
            operator,
            formulas,
            dxf_id,
            ..
        } => {
            assert_eq!(
                (operator.as_str(), formulas.as_slice()),
                ("greaterThan", &["0".to_string()][..])
            );
            let styles = workbook.styles.as_ref().unwrap();
            let dxf = &styles.dxfs[dxf_id.unwrap() as usize];
            assert_eq!(
                dxf.fill.as_ref().unwrap().fg_color.as_deref(),
                Some("00FF00")
            );
        }
        other => panic!("{other:?}"),
    }
    // The formula is restated from its base cell A1 to the range's B2.
    match &format.rules[1] {
        xlsx_model::CfRule::Expression { formula, .. } => assert_eq!(formula, "B2>5"),
        other => panic!("{other:?}"),
    }
}

#[test]
fn calcext_conditions_supersede_style_maps() {
    let automatic = r#"<style:style style:name="ce1" style:family="table-cell"><style:map style:condition="cell-content()&gt;0" style:apply-style-name="Good"/></style:style>"#;
    let common = r##"<style:style style:name="Good" style:family="table-cell"><style:table-cell-properties fo:background-color="#00ff00"/></style:style>"##;
    let body = r#"<table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="3"/></table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.A1"><calcext:condition calcext:apply-style-name="Good" calcext:value="&lt;0" calcext:base-cell-address="Sheet1.A1"/></calcext:conditional-format></calcext:conditional-formats></table:table>"#;
    let workbook = read_text(&content(automatic, body), Some(&styles(common))).unwrap();
    let formats = &workbook.sheets[0].worksheet.conditional_formats;
    assert_eq!(formats.len(), 1);
    assert!(
        matches!(&formats[0].rules[0], xlsx_model::CfRule::CellIs { operator, .. } if operator == "lessThan")
    );
}

#[test]
fn draw_path_reads_its_svg_path_data() {
    let body = r#"<table:table table:name="Sheet1"><table:shapes><draw:path svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm" svg:viewBox="0 0 100 100" svg:d="M0 0 L100 0 L50 100 Z"/><draw:polygon svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm" svg:viewBox="0 0 10 10" draw:points="0,0 10,0 5,10"/></table:shapes><table:table-row><table:table-cell/></table:table-row></table:table>"#;
    let workbook = read_text(&content("", body), Some(&styles(""))).unwrap();
    let groups = &workbook.sheets[0].worksheet.shape_groups;
    assert_eq!(groups.len(), 2);
    for group in groups {
        match &group.shapes[0].geom {
            xlsx_model::ShapeGeom::Custom { paths } => {
                assert!(matches!(
                    paths[0].commands.last(),
                    Some(xlsx_model::PathCmd::Close)
                ));
            }
            other => panic!("{other:?}"),
        }
    }
}

#[test]
fn unbounded_inputs_reject() {
    // Spaces past the part bound.
    let spaces = content(
        "",
        &table(&format!(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p><text:s text:c="{}"/></text:p></table:table-cell></table:table-row>"#,
            crate::odf_text::MAX_PART_SPACES + 1
        )),
    );
    assert_eq!(
        read_text(&spaces, Some(&styles(""))).err(),
        Some(crate::error::resource("text space"))
    );
    // A parent chain past the style bound, and a cycle.
    let chain: String = (0..80)
        .map(|i| {
            let parent = if i == 0 { "Default".to_string() } else { format!("s{}", i - 1) };
            format!(r#"<style:style style:name="s{i}" style:family="table-cell" style:parent-style-name="{parent}"/>"#)
        })
        .collect();
    let plain = content(
        "",
        &table("<table:table-row><table:table-cell/></table:table-row>"),
    );
    assert_eq!(
        read_text(&plain, Some(&styles(&chain))).err(),
        Some(crate::error::resource("style inheritance depth"))
    );
    let cycle = r#"<style:style style:name="a" style:family="table-cell" style:parent-style-name="b"/><style:style style:name="b" style:family="table-cell" style:parent-style-name="a"/>"#;
    assert!(read_text(&plain, Some(&styles(cycle)))
        .err()
        .is_some_and(|e| crate::error::is_invalid(&e)));
    // A sparkline over a range past its bound.
    let sparkline = content(
        "",
        r#"<table:table table:name="Sheet1"><table:table-row><table:table-cell/></table:table-row><calcext:sparkline-groups><calcext:sparkline-group><calcext:sparklines><calcext:sparkline calcext:cell-address="Sheet1.A1" calcext:data-range="Sheet1.A1:Sheet1.Z100000"/></calcext:sparklines></calcext:sparkline-group></calcext:sparkline-groups></table:table>"#,
    );
    assert_eq!(
        read_text(&sparkline, Some(&styles(""))).err(),
        Some(resource("sparkline range"))
    );
}

#[test]
fn repeated_notes_are_charged_their_expanded_text() {
    // A note of 1 MiB of `text:s` spaces repeated across 1024 columns holds
    // 1 GiB of text: each copy is admitted at its expanded size.
    let note = |cols: u32| {
        content(
            "",
            &table(&format!(
                r#"<table:table-row><table:table-cell table:number-columns-repeated="{cols}"><office:annotation><text:p><text:s text:c="1048576"/></text:p></office:annotation></table:table-cell></table:table-row>"#
            )),
        )
    };
    let workbook = read_text(&note(2), Some(&styles(""))).unwrap();
    let comments = &workbook.sheets[0].worksheet.comments;
    assert_eq!(comments.len(), 2);
    assert_eq!(comments[1].cell_ref, "B1");
    assert_eq!(comments[1].text.len(), 1024 * 1024);
    assert_eq!(
        read_text(&note(1024), Some(&styles(""))).err(),
        Some(resource("model size"))
    );
}

#[test]
fn repeated_rich_text_is_charged_its_run_fonts() {
    // A one-character span whose font family is 1 KiB (the bound), repeated
    // over 16 × 16 cells: every copy holds the family name.
    let rich_with = |family: &str, rows: u32, cols: u32| {
        content(
            &format!(
                r#"<style:style style:name="T1" style:family="text"><style:text-properties fo:font-family="{family}"/></style:style>"#
            ),
            &table(&format!(
                r#"<table:table-row table:number-rows-repeated="{rows}"><table:table-cell table:number-columns-repeated="{cols}" office:value-type="string"><text:p><text:span text:style-name="T1">x</text:span></text:p></table:table-cell></table:table-row>"#
            )),
        )
    };
    let family = "F".repeat(FAMILY);
    let automatic = format!(
        r#"<style:style style:name="T1" style:family="text"><style:text-properties fo:font-family="{family}"/></style:style>"#
    );
    let rich = |rows: u32, cols: u32| {
        content(
            &automatic,
            &table(&format!(
                r#"<table:table-row table:number-rows-repeated="{rows}"><table:table-cell table:number-columns-repeated="{cols}" office:value-type="string"><text:p><text:span text:style-name="T1">x</text:span></text:p></table:table-cell></table:table-row>"#
            )),
        )
    };
    let workbook = read_text(&rich(2, 2), Some(&styles(""))).unwrap();
    let cell = &workbook.sheets[0].rows[1].cells[1];
    match &cell.value {
        xlsx_model::CellValue::Text {
            runs: Some(runs), ..
        } => {
            assert_eq!(
                runs[0].font.as_ref().and_then(|f| f.name.as_deref()),
                Some(family.as_str())
            );
        }
        other => panic!("{other:?}"),
    }
    let long = charged(&rich_with(&family, 16, 16), &styles(""));
    let short = charged(&rich_with("F", 16, 16), &styles(""));
    assert!(long - short >= 256 * (FAMILY - 1), "{long} - {short}");
}

#[test]
fn global_names_are_charged_per_sheet() {
    // One 1 MiB global expression copied into each of 1024 sheets.
    let expression = format!("of:=&quot;{}&quot;", "e".repeat(1024 * 1024));
    let names = |sheets: usize| {
        let tables: String = (0..sheets)
            .map(|i| format!(r#"<table:table table:name="S{i}"><table:table-row><table:table-cell/></table:table-row></table:table>"#))
            .collect();
        content(
            "",
            &format!(
                r#"{tables}<table:named-expressions><table:named-expression table:name="Big" table:expression="{expression}"/></table:named-expressions>"#
            ),
        )
    };
    let workbook = read_text(&names(2), Some(&styles(""))).unwrap();
    assert_eq!(workbook.sheets[1].worksheet.defined_names.len(), 1);
    assert_eq!(
        read_text(&names(1024), Some(&styles(""))).err(),
        Some(resource("model size"))
    );
}

#[test]
fn repeat_counts_past_the_bound_reject() {
    for row in [
        r#"<table:table-row table:number-rows-repeated="4294967296"><table:table-cell/></table:table-row>"#,
        r#"<table:table-row><table:table-cell table:number-columns-repeated="99999999999"/></table:table-row>"#,
        r#"<table:table-row><table:table-cell table:number-columns-spanned="4294967296"/></table:table-row>"#,
    ] {
        let body = content("", &table(row));
        assert_eq!(
            read_text(&body, Some(&styles(""))).err(),
            Some(crate::error::resource("repeat count")),
            "{row}"
        );
    }
    let columns = content(
        "",
        r#"<table:table table:name="Sheet1"><table:table-column table:number-columns-repeated="4294967296"/><table:table-row><table:table-cell/></table:table-row></table:table>"#,
    );
    assert_eq!(
        read_text(&columns, Some(&styles(""))).err(),
        Some(crate::error::resource("repeat count"))
    );
}

#[test]
fn style_map_conditions_resolve_their_namespace_prefix() {
    // §19.472: `of:` bound to OpenFormula is the default namespace written
    // out; a condition in another namespace is ignored.
    let automatic = concat!(
        r#"<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:x="urn:other">"#,
        r#"<style:map style:condition="of:cell-content()&gt;0" style:apply-style-name="Good" style:base-cell-address="Sheet1.A1"/>"#,
        r#"<style:map style:condition="x:cell-content()&lt;0" style:apply-style-name="Good" style:base-cell-address="Sheet1.A1"/>"#,
        r#"</style:style>"#
    );
    let common = r##"<style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#00ff00"/></style:style>"##;
    let body = table(
        r#"<table:table-row><table:table-cell table:style-name="ce1" office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell></table:table-row>"#,
    );
    let workbook = read_text(&content(automatic, &body), Some(&styles(common))).unwrap();
    let formats = &workbook.sheets[0].worksheet.conditional_formats;
    assert_eq!(formats.len(), 1);
    assert_eq!(formats[0].rules.len(), 1);
    assert!(
        matches!(&formats[0].rules[0], xlsx_model::CfRule::CellIs { operator, .. } if operator == "greaterThan")
    );
}

#[test]
fn validations_are_charged_per_sheet() {
    // One validation with a 1 MiB message, used on each of 1024 sheets.
    let message = "m".repeat(1024 * 1024);
    let validations = |sheets: usize| {
        let tables: String = (0..sheets)
            .map(|i| format!(r#"<table:table table:name="S{i}"><table:table-row><table:table-cell table:content-validation-name="V"/></table:table-row></table:table>"#))
            .collect();
        content(
            "",
            &format!(
                r#"<table:content-validations><table:content-validation table:name="V" table:condition="of:cell-content-is-whole-number()"><table:error-message table:title="t"><text:p>{message}</text:p></table:error-message></table:content-validation></table:content-validations>{tables}"#
            ),
        )
    };
    let workbook = read_text(&validations(2), Some(&styles(""))).unwrap();
    let validation = &workbook.sheets[1].worksheet.data_validations[0];
    assert_eq!(validation.sqref, "A1");
    assert_eq!(
        validation.error_message.as_deref().map(str::len),
        Some(1024 * 1024)
    );
    assert_eq!(
        read_text(&validations(1024), Some(&styles(""))).err(),
        Some(resource("model size"))
    );
}

fn empty_tables(sheets: usize, columns: &str) -> String {
    (0..sheets)
        .map(|i| format!(r#"<table:table table:name="S{i}">{columns}<table:table-row><table:table-cell/></table:table-row></table:table>"#))
        .collect()
}

const FAMILY: usize = crate::odf_text::MAX_STYLE_VALUE_BYTES;

#[test]
fn the_default_family_is_charged_per_sheet() {
    // Every worksheet keeps its own copy of the default family.
    let body = content("", &empty_tables(64, ""));
    let long = charged(&body, &styles_with_default_family(FAMILY, ""));
    let short = charged(&body, &styles_with_default_family(1, ""));
    assert!(long - short >= 64 * (FAMILY - 1), "{long} - {short}");
}

#[test]
fn interned_fonts_are_charged_with_their_keys() {
    // Styles that differ only in size each intern a font holding the
    // family, and a key that spells it.
    let sized: String = (0..32)
        .map(|i| format!(r#"<style:style style:name="s{i}" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-size="{}pt"/></style:style>"#, i + 1))
        .collect();
    let cells: String = (0..32)
        .map(|i| format!(r#"<table:table-cell table:style-name="s{i}" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>"#))
        .collect();
    let body = content(
        "",
        &table(&format!("<table:table-row>{cells}</table:table-row>")),
    );
    let (workbook, long) = read_charging(&body, &styles_with_default_family(FAMILY, &sized));
    // 32 sizes, one of them the default's 10 pt.
    assert_eq!(workbook.unwrap().styles.unwrap().fonts.len(), 32);
    let short = charged(&body, &styles_with_default_family(1, &sized));
    assert!(long - short >= 2 * 32 * (FAMILY - 1), "{long} - {short}");
}

#[test]
fn repeated_column_metadata_is_charged_per_sheet() {
    // A hidden, outlined column repeated across the whole grid: one map
    // entry per column and map on every sheet.
    let columns = r#"<table:table-column-group><table:table-column table:visibility="collapse" table:number-columns-repeated="16384"/></table:table-column-group>"#;
    let (workbook, hidden) = read_charging(&content("", &empty_tables(16, columns)), &styles(""));
    let workbook = workbook.unwrap();
    assert_eq!(workbook.sheets[15].worksheet.col_hidden.len(), 16_384);
    assert_eq!(
        workbook.sheets[15].worksheet.col_outline_levels.len(),
        16_384
    );
    let plain = charged(&content("", &empty_tables(16, "")), &styles(""));
    assert!(hidden - plain >= 16 * 16_384 * 2 * 10, "{hidden} - {plain}");
}

#[test]
fn shape_text_runs_are_charged_their_font_copies() {
    // Short spans in one text box; every run holds its own copy of the
    // graphic style's family, admitted before it is made.
    let spans: String = (0..2048).map(|_| r#"<text:span>x</text:span>y"#).collect();
    let body = content(
        r#"<style:style style:name="gr1" style:family="graphic"><style:text-properties style:font-name="F"/></style:style>"#,
        &table(&format!(
            r#"<table:table-row><table:table-cell><draw:frame draw:style-name="gr1" svg:width="2cm" svg:height="1cm"><draw:text-box><text:p>{spans}</text:p></draw:text-box></draw:frame></table:table-cell></table:table-row>"#
        )),
    );
    let (workbook, long) = read_charging(&body, &styles_with_default_family(FAMILY, ""));
    let workbook = workbook.unwrap();
    let shape = &workbook.sheets[0].worksheet.shape_groups[0].shapes[0];
    let runs = &shape.text.as_ref().unwrap().paragraphs[0].runs;
    assert_eq!(runs.len(), 4096);
    assert!(matches!(
        &runs[0],
        xlsx_model::ShapeTextRun::Text { font_face: Some(face), .. } if face.len() == FAMILY
    ));
    let short = charged(&body, &styles_with_default_family(1, ""));
    assert!(long - short >= 4096 * (FAMILY - 1), "{long} - {short}");
}

#[test]
fn style_values_past_the_bound_reject() {
    let body = content("", &empty_tables(1, ""));
    let (workbook, _) = read_charging(&body, &styles_with_default_family(FAMILY + 1, ""));
    assert_eq!(workbook.err(), Some(crate::error::resource("style value")));
}

#[test]
fn blank_cells_over_many_column_runs_walk_nothing() {
    // 16,384 single-column runs and 4,096 rows of one blank cell repeated
    // over all of them: no cell is materialized and no run is walked.
    let columns: String = (0..16_384).map(|_| "<table:table-column/>").collect();
    let rows: String = (0..4_096)
        .map(|_| r#"<table:table-row><table:table-cell table:number-columns-repeated="16384"/></table:table-row>"#)
        .collect();
    let body = content(
        "",
        &format!(r#"<table:table table:name="S">{columns}{rows}</table:table>"#),
    );
    let manifest = manifest(ODS, "");
    let styles = styles("");
    let files: Vec<(&str, &[u8])> = vec![
        ("META-INF/manifest.xml", manifest.as_bytes()),
        ("content.xml", body.as_bytes()),
        ("styles.xml", styles.as_bytes()),
    ];
    let package = OdfPackage::open(
        package_bytes(ODS, &files),
        OoxmlFormat::Xlsx,
        PackageLimits::default(),
        &MEDIA_TYPES,
    )
    .unwrap();
    let budget = Budget::new();
    let workbook = read_within(&package, budget.clone()).unwrap();
    assert!(workbook.sheets[0].rows.is_empty());
    // Only the formatted-column scan's first pass; no row walks a run.
    assert!(budget.steps_used() < 4_096, "{}", budget.steps_used());
}

#[test]
fn default_style_conditions_reach_columns_past_the_declared_runs() {
    // Only column A is declared; the `Default` style carries a condition,
    // so both unstyled blank cells are covered by it.
    let styles = format!(
        r##"<office:document-styles {NS} office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"><style:map style:condition="is-true-formula(1)" style:apply-style-name="Hit"/></style:style><style:style style:name="Hit" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ff0000"/></style:style></office:styles></office:document-styles>"##
    );
    let body = content(
        "",
        r#"<table:table table:name="S"><table:table-column/><table:table-row><table:table-cell/><table:table-cell/></table:table-row></table:table>"#,
    );
    let workbook = read_text(&body, Some(&styles)).unwrap();
    let formats = &workbook.sheets[0].worksheet.conditional_formats;
    let covered: Vec<(u32, u32)> = formats
        .iter()
        .flat_map(|f| f.sqref.iter().map(|r| (r.left, r.right)))
        .collect();
    assert!(
        covered.iter().any(|(l, r)| *l <= 1 && 1 <= *r),
        "{covered:?}"
    );
    assert!(
        covered.iter().any(|(l, r)| *l <= 2 && 2 <= *r),
        "{covered:?}"
    );
}
