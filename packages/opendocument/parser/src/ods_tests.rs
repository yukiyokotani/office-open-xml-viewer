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

/// Read `content` and `styles` holding at most `limit` bytes: the workbook
/// (or its rejection), the bytes held at the peak and the steps taken.
fn read_measured_text(
    content: &str,
    styles: &str,
    limit: usize,
) -> (Result<OdsWorkbook, String>, usize, u64) {
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
    read_measured(&package, limit)
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
fn repeated_links_cover_every_cell_and_are_bounded() {
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
    // A `text:s` run larger than the model: rejected while it is written,
    // in a cell, a note and a text box alike.
    for holder in [
        r#"<table:table-cell office:value-type="string"><text:p><text:s text:c="18446744073709551615"/></text:p></table:table-cell>"#,
        r#"<table:table-cell><office:annotation><text:p><text:s text:c="4000000000"/></text:p></office:annotation></table:table-cell>"#,
        r#"<table:table-cell><draw:frame svg:width="2cm" svg:height="1cm"><draw:text-box><text:p><text:s text:c="4000000000"/></text:p></draw:text-box></draw:frame></table:table-cell>"#,
    ] {
        let spaces = content(
            "",
            &table(&format!("<table:table-row>{holder}</table:table-row>")),
        );
        assert_eq!(
            read_measured_text(&spaces, &styles(""), 8 << 20).0.err(),
            Some(resource("model size")),
            "{holder}"
        );
    }
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
fn repeated_notes_hold_their_expanded_text() {
    // A note of 1 MiB of `text:s` spaces repeated across 1024 columns holds
    // 1 GiB of text: the copies are counted as they are made.
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
fn repeated_rich_text_holds_its_run_fonts() {
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
    // Nothing estimates sizes: the allocator sees every copy of the family
    // the 4096 cells hold.
    let held = |family: &str| {
        let (workbook, peak, _) =
            read_measured_text(&rich_with(family, 64, 64), &styles(""), MAX_MODEL_BYTES);
        workbook.unwrap();
        peak
    };
    let (long, short) = (held(&family), held("F"));
    assert!(long - short >= 4096 * (FAMILY - 1), "{long} - {short}");
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

fn empty_tables(sheets: usize, columns: &str) -> String {
    (0..sheets)
        .map(|i| format!(r#"<table:table table:name="S{i}">{columns}<table:table-row><table:table-cell/></table:table-row></table:table>"#))
        .collect()
}

const FAMILY: usize = crate::odf_text::MAX_STYLE_VALUE_BYTES;

#[test]
fn style_values_past_the_bound_reject() {
    let body = content("", &empty_tables(1, ""));
    let workbook = read_text(&body, Some(&styles_with_default_family(FAMILY + 1, "")));
    assert_eq!(workbook.err(), Some(crate::error::resource("style value")));
    // A shared picture-fill path, a style reference and an extension
    // property set are bounded alike.
    let long = "p".repeat(FAMILY + 1);
    for common in [
        format!(r#"<draw:fill-image draw:name="I" xlink:href="Pictures/{long}.png"/>"#),
        format!(
            r#"<style:style style:name="g" style:family="graphic"><loext:graphic-properties xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" svg:stroke-width="{long}"/></style:style>"#
        ),
    ] {
        let workbook = read_text(&body, Some(&styles_with_default_family(1, &common)));
        assert_eq!(workbook.err(), Some(crate::error::resource("style value")));
    }
    let reference = content(
        "",
        &table(&format!(
            r#"<table:table-row><table:table-cell table:style-name="{long}"/></table:table-row>"#
        )),
    );
    let workbook = read_text(&reference, Some(&styles("")));
    assert_eq!(workbook.err(), Some(crate::error::resource("style value")));
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

#[test]
fn a_styled_cell_does_not_take_its_column_default_properties() {
    // A1's own style sets only bold: as LibreOffice reads it (and the
    // files it writes rely on), the column's default cell style does not
    // fill in what the cell's style leaves unset.
    let automatic = r##"<style:style style:name="bold" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="yellow" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffff00"/></style:style>"##;
    let body = content(
        automatic,
        r#"<table:table table:name="S"><table:table-column table:default-cell-style-name="yellow"/><table:table-row><table:table-cell table:style-name="bold" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row></table:table>"#,
    );
    let workbook = read_text(&body, Some(&styles(""))).unwrap();
    let styles = workbook.styles.as_ref().unwrap();
    let xf = &styles.cell_xfs[workbook.sheets[0].rows[0].cells[0].style_index.unwrap() as usize];
    assert!(styles.fonts[xf.font_id as usize].bold);
    assert_eq!(styles.fills[xf.fill_id as usize].fg_color, None);
}

#[test]
fn nested_spans_inherit_the_outer_span_formatting() {
    // §16.2: the inner span sets only italic; bold comes from the outer.
    let automatic = r#"<style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>"#;
    let body = content(
        automatic,
        &table(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p>a<text:span text:style-name="B">b<text:span text:style-name="I">c</text:span></text:span></text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(""))).unwrap();
    let xlsx_model::CellValue::Text {
        runs: Some(runs), ..
    } = &workbook.sheets[0].rows[0].cells[0].value
    else {
        panic!("no runs");
    };
    let fonts: Vec<(&str, bool, bool)> = runs
        .iter()
        .map(|r| {
            let font = r.font.as_ref().unwrap();
            (r.text.as_str(), font.bold, font.italic)
        })
        .collect();
    assert_eq!(
        fonts,
        [("a", false, false), ("b", true, false), ("c", true, true)]
    );
}

#[test]
fn paragraph_styles_format_cell_text_under_the_text_default() {
    // The paragraph's bold applies; the text family default (normal
    // weight) does not override what the cell and paragraph set.
    let automatic = r#"<style:style style:name="P" style:family="paragraph"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>"#;
    let common = r#"<style:default-style style:family="text"><style:text-properties fo:font-weight="normal"/></style:default-style>"#;
    let body = content(
        automatic,
        &table(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p text:style-name="P">a<text:span text:style-name="I">b</text:span></text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let xlsx_model::CellValue::Text {
        runs: Some(runs), ..
    } = &workbook.sheets[0].rows[0].cells[0].value
    else {
        panic!("no runs");
    };
    let fonts: Vec<(&str, bool, bool)> = runs
        .iter()
        .map(|r| {
            let font = r.font.as_ref().unwrap();
            (r.text.as_str(), font.bold, font.italic)
        })
        .collect();
    assert_eq!(fonts, [("a", true, false), ("b", true, true)]);
}

#[test]
fn calcext_formats_supersede_style_maps_only_on_their_own_table() {
    let automatic = r#"<style:style style:name="ce1" style:family="table-cell"><style:map style:condition="is-true-formula(1)" style:apply-style-name="Good"/></style:style>"#;
    let common = r##"<style:style style:name="Good" style:family="table-cell"><style:table-cell-properties fo:background-color="#00ff00"/></style:style>"##;
    let body = r#"<table:table table:name="S"><table:table-row><table:table-cell table:style-name="ce1"/></table:table-row></table:table><table:table table:name="T"><table:table-row><table:table-cell/></table:table-row><calcext:conditional-formats/></table:table>"#;
    let workbook = read_text(&content(automatic, body), Some(&styles(common))).unwrap();
    assert_eq!(workbook.sheets[0].worksheet.conditional_formats.len(), 1);
}

/// (text, bold, italic, size, colour) of one run.
type RunFacts = (String, bool, bool, Option<f64>, Option<String>);

/// The facts of each run of a cell.
fn run_fonts(cell: &xlsx_model::Cell) -> Vec<RunFacts> {
    let xlsx_model::CellValue::Text {
        runs: Some(runs), ..
    } = &cell.value
    else {
        panic!("no runs: {:?}", cell.value);
    };
    runs.iter()
        .map(|r| {
            let font = r.font.as_ref().unwrap();
            (
                r.text.clone(),
                font.bold,
                font.italic,
                font.size,
                font.color.clone(),
            )
        })
        .collect()
}

#[test]
fn the_text_family_default_is_the_last_fallback_of_a_span() {
    let common = r##"<style:default-style style:family="text"><style:text-properties fo:color="#ff0000"/></style:default-style>"##;
    let automatic = r#"<style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>"#;
    let body = content(
        automatic,
        &table(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p><text:span text:style-name="I">x</text:span></text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let runs = run_fonts(&workbook.sheets[0].rows[0].cells[0]);
    assert_eq!(runs[0].4.as_deref(), Some("FF0000"));
    assert!(runs[0].2);
}

#[test]
fn repeated_rich_text_takes_each_column_default() {
    let automatic = r#"<style:style style:name="bold" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="plain" style:family="table-cell"><style:text-properties fo:font-weight="normal"/></style:style><style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>"#;
    let body = content(
        automatic,
        r#"<table:table table:name="S"><table:table-column table:default-cell-style-name="bold"/><table:table-column table:default-cell-style-name="plain"/><table:table-row><table:table-cell table:number-columns-repeated="2" office:value-type="string"><text:p><text:span text:style-name="I">x</text:span></text:p></table:table-cell></table:table-row></table:table>"#,
    );
    let workbook = read_text(&body, Some(&styles(""))).unwrap();
    let cells = &workbook.sheets[0].rows[0].cells;
    let a = run_fonts(&cells[0]);
    let b = run_fonts(&cells[1]);
    assert!(a[0].1 && a[0].2, "{a:?}");
    assert!(!b[0].1 && b[0].2, "{b:?}");
}

#[test]
fn percentage_font_sizes_scale_the_parent_style() {
    // §20.190: 150% of the parent style's 20pt is 30pt, whatever the cell.
    let common = r#"<style:style style:name="Base" style:family="paragraph"><style:text-properties fo:font-size="20pt"/></style:style><style:style style:name="Relative" style:family="paragraph" style:parent-style-name="Base"><style:text-properties fo:font-size="150%"/></style:style>"#;
    let body = content(
        "",
        &table(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p text:style-name="Relative">x</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let runs = run_fonts(&workbook.sheets[0].rows[0].cells[0]);
    assert_eq!(runs[0].3, Some(30.0));
}

#[test]
fn a_cell_paragraph_falls_back_to_the_paragraph_family_default() {
    // §16.2: the initial family of `<text:p text:style-name="P">` is
    // paragraph, so its default (red) applies, not the text default.
    let common = r##"<style:default-style style:family="paragraph"><style:text-properties fo:color="#ff0000"/></style:default-style><style:default-style style:family="text"><style:text-properties fo:color="#0000ff"/></style:default-style>"##;
    let automatic = r#"<style:style style:name="P" style:family="paragraph"><style:text-properties fo:font-style="italic"/></style:style>"#;
    let body = content(
        automatic,
        &table(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p text:style-name="P">x</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let runs = run_fonts(&workbook.sheets[0].rows[0].cells[0]);
    assert_eq!(runs[0].4.as_deref(), Some("FF0000"));
}

#[test]
fn shape_text_defaults_come_after_the_graphic_and_paragraph_styles() {
    // The graphic style's green beats the paragraph default's red; a span
    // setting only italic keeps the paragraph's bold over the text
    // default's normal weight, and an unset colour still reaches the
    // text default where nothing else sets one.
    let common = r##"<style:default-style style:family="paragraph"><style:text-properties fo:color="#ff0000"/></style:default-style><style:default-style style:family="text"><style:text-properties fo:font-weight="normal"/></style:default-style>"##;
    let automatic = r##"<style:style style:name="gr1" style:family="graphic"><style:text-properties fo:color="#00ff00"/></style:style><style:style style:name="P" style:family="paragraph"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>"##;
    let body = content(
        automatic,
        &table(
            r#"<table:table-row><table:table-cell><draw:frame draw:style-name="gr1" svg:width="2cm" svg:height="1cm"><draw:text-box><text:p text:style-name="P">a<text:span text:style-name="I">b</text:span></text:p></draw:text-box></draw:frame></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let shape = &workbook.sheets[0].worksheet.shape_groups[0].shapes[0];
    let runs: Vec<(String, bool, bool, Option<String>)> = shape.text.as_ref().unwrap().paragraphs
        [0]
    .runs
    .iter()
    .filter_map(|run| match run {
        xlsx_model::ShapeTextRun::Text {
            text,
            bold,
            italic,
            color,
            ..
        } => Some((text.clone(), *bold, *italic, color.clone())),
        _ => None,
    })
    .collect();
    assert_eq!(
        runs,
        [
            ("a".to_string(), true, false, Some("#00FF00".to_string())),
            ("b".to_string(), true, true, Some("#00FF00".to_string())),
        ]
    );
}

#[test]
fn a_cell_percentage_keeps_its_basis_inside_a_paragraph_context() {
    // 150% of the table-cell default's 10pt is 15pt in the cell's font and
    // in its paragraph's runs alike, whatever the paragraph default says.
    let common = r#"<style:default-style style:family="table-cell"><style:text-properties fo:font-size="10pt"/></style:default-style><style:default-style style:family="paragraph"><style:text-properties fo:font-size="20pt"/></style:default-style><style:style style:name="S" style:family="table-cell"><style:text-properties fo:font-size="150%"/></style:style><style:style style:name="P" style:family="paragraph"><style:text-properties fo:font-style="italic"/></style:style>"#;
    let body = content(
        "",
        &table(
            r#"<table:table-row><table:table-cell table:style-name="S" office:value-type="string"><text:p text:style-name="P">x</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let cell = &workbook.sheets[0].rows[0].cells[0];
    let styles = workbook.styles.as_ref().unwrap();
    let xf = &styles.cell_xfs[cell.style_index.unwrap() as usize];
    assert_eq!(styles.fonts[xf.font_id as usize].size, 15.0);
    assert_eq!(run_fonts(cell)[0].3, Some(15.0));
}

#[test]
fn window_colour_off_inherits_the_parent_colour() {
    // §20.395: `false` selects fo:color, inherited from the parent here.
    let common = r##"<style:style style:name="Red" style:family="table-cell"><style:text-properties fo:color="#ff0000"/></style:style><style:style style:name="Child" style:family="table-cell" style:parent-style-name="Red"><style:text-properties style:use-window-font-color="false"/></style:style>"##;
    let body = content(
        "",
        &table(
            r#"<table:table-row><table:table-cell table:style-name="Child" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(common))).unwrap();
    let styles = workbook.styles.as_ref().unwrap();
    let xf = &styles.cell_xfs[workbook.sheets[0].rows[0].cells[0].style_index.unwrap() as usize];
    assert_eq!(
        styles.fonts[xf.font_id as usize].color.as_deref(),
        Some("FF0000")
    );
}

/// The format code of a cell's xf (`General` for the built-in id 0).
fn format_code(workbook: &OdsWorkbook, cell: &xlsx_model::Cell) -> String {
    let styles = workbook.styles.as_ref().unwrap();
    let id = styles.cell_xfs[cell.style_index.unwrap() as usize].num_fmt_id;
    styles
        .num_fmts
        .iter()
        .find(|f| f.num_fmt_id == id)
        .map_or_else(|| "General".to_string(), |f| f.format_code.clone())
}

#[test]
fn blank_repeats_walk_their_xml_not_their_grid() {
    // 16,384 single-column runs and 4,096 rows of one blank cell repeated
    // over all of them (67 million grid cells): nothing is materialized and
    // the work follows the XML, whether or not the cell could span.
    let columns: String = (0..16_384).map(|_| "<table:table-column/>").collect();
    let rows: String = (0..4_096)
        .map(|_| r#"<table:table-row><table:table-cell table:number-columns-repeated="16384"/></table:table-row>"#)
        .collect();
    let body = content(
        "",
        &format!(r#"<table:table table:name="S">{columns}{rows}</table:table>"#),
    );
    let (workbook, _, work) = read_measured_text(&body, &styles(""), MAX_MODEL_BYTES);
    assert!(workbook.unwrap().sheets[0].rows.is_empty());
    assert!(work < 400_000, "{work}");
}

#[test]
fn repeated_spanned_cells_merge_once_per_cell() {
    let body = content(
        "",
        &table(
            r#"<table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="3" table:number-columns-spanned="2" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let workbook = read_text(&body, Some(&styles(""))).unwrap();
    let merges = &workbook.sheets[0].worksheet.merge_cells;
    assert_eq!(merges.len(), 6);
    assert_eq!(
        (
            merges[5].top,
            merges[5].left,
            merges[5].bottom,
            merges[5].right
        ),
        (2, 3, 2, 4)
    );
}

#[test]
fn rich_text_is_bounded_while_its_spans_are_read() {
    // Sibling spans that each resolve a 1 KiB family: the read stops at
    // its limit while it walks them, not after collecting them all.
    let family = "F".repeat(FAMILY);
    let spans: String = (0..40_000)
        .map(|_| r#"<text:span text:style-name="T1">x</text:span>y"#)
        .collect();
    let body = content(
        &format!(
            r#"<style:style style:name="T1" style:family="text"><style:text-properties fo:font-family="{family}"/></style:style>"#
        ),
        &table(&format!(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p>{spans}</text:p></table:table-cell></table:table-row>"#
        )),
    );
    let (workbook, peak, _) = read_measured_text(&body, &styles(""), 4 << 20);
    assert_eq!(workbook.err(), Some(resource("model size")));
    // 40,000 contexts would hold 40 MiB of family copies.
    assert!(peak < 6 << 20, "{peak}");
}

#[test]
fn a_grid_of_repeated_content_rejects() {
    // One value repeated over the whole grid: 17 billion cells.
    let body = content(
        "",
        &table(
            r#"<table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="16384" office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let (workbook, peak, _) = read_measured_text(&body, &styles(""), 16 << 20);
    assert_eq!(workbook.err(), Some(resource("model size")));
    assert!(peak < 24 << 20, "{peak}");
}

#[test]
fn early_dates_use_the_serials_the_renderer_displays() {
    let body = content(
        "",
        &table(concat!(
            r#"<table:table-row>"#,
            r#"<table:table-cell office:value-type="date" office:date-value="1900-01-01"><text:p>x</text:p></table:table-cell>"#,
            r#"<table:table-cell office:value-type="date" office:date-value="1900-02-28"><text:p>x</text:p></table:table-cell>"#,
            r#"<table:table-cell office:value-type="date" office:date-value="1900-03-01"><text:p>x</text:p></table:table-cell>"#,
            r#"</table:table-row>"#
        )),
    );
    let workbook = read_text(&body, Some(&styles(""))).unwrap();
    let serials: Vec<f64> = workbook.sheets[0].rows[0]
        .cells
        .iter()
        .map(|cell| match cell.value {
            xlsx_model::CellValue::Number { number } => number,
            _ => f64::NAN,
        })
        .collect();
    // The renderer shows serial 1 as 1900-01-01 and 59 as 1900-02-28.
    assert_eq!(serials, [1.0, 59.0, 61.0]);
}

#[test]
fn a_formula_zero_shown_as_nothing_stays_a_number() {
    let body = content(
        "",
        &table(concat!(
            r#"<table:table-row>"#,
            // An empty-cell result: LibreOffice displays nothing.
            r#"<table:table-cell table:formula="of:=VLOOKUP(1;[.X1:.Y2];2)" office:value-type="float" office:value="0"><text:p/></table:table-cell>"#,
            // A zero LibreOffice displays.
            r#"<table:table-cell table:formula="of:=1-1" office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell>"#,
            // A constant zero keeps its own format.
            r#"<table:table-cell office:value-type="float" office:value="0"><text:p/></table:table-cell>"#,
            r#"</table:table-row>"#
        )),
    );
    let workbook = read_text(&body, Some(&styles(""))).unwrap();
    let cells = &workbook.sheets[0].rows[0].cells;
    for cell in cells {
        assert!(
            matches!(cell.value, xlsx_model::CellValue::Number { number } if number == 0.0),
            "{:?}",
            cell.value
        );
    }
    assert_eq!(format_code(&workbook, &cells[0]), "0;-0;;@");
    assert_eq!(format_code(&workbook, &cells[1]), "General");
    assert_eq!(format_code(&workbook, &cells[2]), "General");
}

#[test]
fn charts_take_the_workbook_date_system() {
    let chart_content = format!(
        r#"<office:document-content {NS} office:version="1.3"><office:body><office:chart><chart:chart chart:class="chart:line"/></office:chart></office:body></office:document-content>"#
    );
    let body = content(
        "",
        &format!(
            r#"<table:calculation-settings><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>{}"#,
            table(
                r#"<table:table-row><table:table-cell><draw:frame svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1"/></draw:frame></table:table-cell></table:table-row>"#
            )
        ),
    );
    let styles = styles("");
    let workbook = read_files(&[
        ("content.xml", body.as_bytes()),
        ("styles.xml", styles.as_bytes()),
        ("Object 1/content.xml", chart_content.as_bytes()),
    ])
    .unwrap();
    assert!(workbook.date1904);
    assert!(workbook.sheets[0].worksheet.charts[0].chart.date1904);
}
