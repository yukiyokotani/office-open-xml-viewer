use super::*;
use crate::package::tests::{manifest, package_bytes, ODT};
use crate::package::PackageLimits;
use docx_model::{CellElement, DocTable, TextBoxBlockWire};
use ooxml_common::resource::OoxmlFormat;

const NS: &str = r##"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:officeooo="http://openoffice.org/2009/office" xmlns:xlink="http://www.w3.org/1999/xlink""##;
const FACES: &str = r##"<office:font-face-decls><style:font-face style:name="Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman"/><style:font-face style:name="Sans" svg:font-family="'Liberation Sans', Arial" style:font-family-generic="swiss"/><style:font-face style:name="Mincho" svg:font-family="'Hiragino Mincho ProN'" style:font-family-generic="system"/></office:font-face-decls>"##;

#[derive(Default)]
struct Fixture<'a> {
    styles: &'a str,
    automatic: &'a str,
    /// Automatic styles of styles.xml (page layouts and master-page styles).
    styles_automatic: &'a str,
    /// Extra master pages after "Standard".
    masters: &'a str,
    /// Content of the "Standard" master page.
    master: &'a str,
    /// Extra attributes of the "Standard" page layout properties.
    page: &'a str,
    /// Children of the "Standard" page layout (header-style, ...).
    page_children: &'a str,
    body: &'a str,
    files: &'a [(&'a str, &'a [u8])],
}

fn read_fixture(fixture: Fixture<'_>) -> Document {
    read_fixture_resources(fixture).0
}

fn read_fixture_resources(fixture: Fixture<'_>) -> (Document, Vec<OwnedDocumentResource>) {
    try_read_fixture(fixture).unwrap()
}

fn try_read_fixture(
    fixture: Fixture<'_>,
) -> Result<(Document, Vec<OwnedDocumentResource>), String> {
    let styles = format!(
        r##"<office:document-styles {NS} office:version="1.3">{FACES}<office:styles><style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="1.25cm"/><style:text-properties style:font-name="Serif" fo:font-size="12pt" style:font-name-asian="Mincho" style:font-size-asian="10.5pt" style:font-name-complex="Serif" style:font-size-complex="12pt"/></style:default-style>{}</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="210mm" fo:page-height="297mm" fo:margin-top="20mm" fo:margin-bottom="20mm" fo:margin-left="25mm" fo:margin-right="25mm"{}/>{}</style:page-layout>{}</office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1">{}</style:master-page>{}</office:master-styles></office:document-styles>"##,
        fixture.styles,
        fixture.page,
        fixture.page_children,
        fixture.styles_automatic,
        fixture.master,
        fixture.masters,
    );
    let content = format!(
        r##"<office:document-content {NS} office:version="1.3">{FACES}<office:automatic-styles>{}</office:automatic-styles><office:body><office:text>{}</office:text></office:body></office:document-content>"##,
        fixture.automatic, fixture.body
    );
    let files = [
        ("META-INF/manifest.xml", manifest(ODT, "")),
        ("styles.xml", styles),
        ("content.xml", content),
    ];
    let mut files: Vec<(&str, &[u8])> = files
        .iter()
        .map(|(name, body)| (*name, body.as_bytes()))
        .collect();
    files.extend(fixture.files.iter().copied());
    let bytes = package_bytes(ODT, &files);
    let package = OdfPackage::open(
        bytes,
        OoxmlFormat::Docx,
        PackageLimits::default(),
        &MEDIA_TYPES,
    )
    .unwrap();
    read(&package)
}

fn paragraphs(elements: &[BodyElement]) -> Vec<&DocParagraph> {
    elements
        .iter()
        .filter_map(|element| match element {
            BodyElement::Paragraph(paragraph) => Some(paragraph.as_ref()),
            _ => None,
        })
        .collect()
}

fn texts(paragraph: &DocParagraph) -> Vec<String> {
    paragraph
        .runs
        .iter()
        .map(|run| match run {
            DocRun::Text(text) => text.text.clone(),
            DocRun::Break { .. } => "<br>".into(),
            DocRun::Field(field) => format!("<{}>", field.field_type),
            DocRun::Image(_) => "<image>".into(),
            DocRun::Shape(_) => "<shape>".into(),
            DocRun::AnchorHost(_) => "<host>".into(),
            _ => "?".into(),
        })
        .collect()
}

fn text_runs(paragraph: &DocParagraph) -> Vec<&TextRun> {
    paragraph
        .runs
        .iter()
        .filter_map(|run| match run {
            DocRun::Text(text) => Some(text.as_ref()),
            _ => None,
        })
        .collect()
}

#[test]
fn collapses_white_space_and_keeps_spaces_tabs_and_breaks() {
    let document = read_fixture(Fixture {
        body:
            "<text:p>\n  a \t b<text:s text:c=\"2\"/>cd<text:tab/>x<text:line-break/>  e  </text:p>",
        ..Fixture::default()
    });
    assert_eq!(
        texts(paragraphs(&document.body)[0]),
        ["a b  cd\tx", "<br>", " e"]
    );
    assert!((document.section.margin_left - 70.866).abs() < 0.01);
}

#[test]
fn adds_adjacent_margins_and_removes_them_between_contextual_paragraphs() {
    let styles = r##"<style:style style:name="A" style:family="paragraph"><style:paragraph-properties fo:margin-top="6pt" fo:margin-bottom="10pt"/></style:style><style:style style:name="C" style:family="paragraph"><style:paragraph-properties fo:margin-top="12pt" fo:margin-bottom="12pt" style:contextual-spacing="true"/></style:style>"##;
    let document = read_fixture(Fixture {
        styles,
        body: r##"<text:p text:style-name="A">a</text:p><text:p text:style-name="A">b</text:p><text:p text:style-name="C">c</text:p><text:p text:style-name="C">d</text:p>"##,
        ..Fixture::default()
    });
    let spacing: Vec<(f64, f64)> = paragraphs(&document.body)
        .iter()
        .map(|p| (p.space_before, p.space_after))
        .collect();
    assert_eq!(spacing, [(6.0, 0.0), (16.0, 0.0), (22.0, 0.0), (0.0, 0.0)]);
}

#[test]
fn maps_line_heights_borders_and_padding() {
    let automatic = r##"<style:style style:name="F" style:family="paragraph"><style:paragraph-properties fo:line-height="20pt"/></style:style><style:style style:name="L" style:family="paragraph"><style:paragraph-properties style:line-height-at-least="30pt"/></style:style><style:style style:name="P" style:family="paragraph"><style:paragraph-properties fo:line-height="150%" fo:margin-left="10pt" fo:border="1pt solid #ff0000" fo:padding="2pt"/></style:style><style:style style:name="Q" style:family="paragraph"><style:paragraph-properties fo:padding-left="5pt"/></style:style>"##;
    let document = read_fixture(Fixture {
        automatic,
        body: r##"<text:p text:style-name="F">a</text:p><text:p text:style-name="L">b</text:p><text:p text:style-name="P">c</text:p><text:p text:style-name="Q">d</text:p>"##,
        ..Fixture::default()
    });
    let p = paragraphs(&document.body);
    let rule = |i: usize| {
        let spacing = p[i].line_spacing.as_ref().unwrap();
        (spacing.rule.as_str(), spacing.value)
    };
    assert_eq!(rule(0), ("exact", 20.0));
    assert_eq!(rule(1), ("atLeast", 30.0));
    assert_eq!(rule(2), ("auto", 1.5));
    // The border's outside sits at the margin; the text after the border and
    // its padding.
    assert_eq!(p[2].indent_left, 13.0);
    let left = p[2].borders.as_ref().unwrap().left.as_ref().unwrap();
    assert_eq!(
        (left.width, left.space, left.color.as_deref()),
        (1.0, 2.0, Some("ff0000"))
    );
    // Padding without a border insets the text.
    assert_eq!(p[3].indent_left, 5.0);
}

#[test]
fn resolves_character_properties() {
    let automatic = r##"<style:style style:name="T" style:family="text"><style:text-properties fo:letter-spacing="1pt" style:text-scale="50%" style:text-position="super 58%" fo:font-weight="bold" style:text-underline-style="dotted" style:text-underline-type="single" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style><style:style style:name="U" style:family="text"><style:text-properties fo:text-transform="uppercase" fo:font-variant="small-caps"/></style:style><style:style style:name="W" style:family="text"><style:text-properties fo:text-transform="lowercase" fo:font-size="150%"/></style:style><style:style style:name="H" style:family="text"><style:text-properties text:display="none"/></style:style>"##;
    let document = read_fixture(Fixture {
        automatic,
        body: r##"<text:p><text:span text:style-name="T">a</text:span><text:span text:style-name="U">b</text:span><text:span text:style-name="W">CD</text:span><text:span text:style-name="H">hidden</text:span></text:p>"##,
        ..Fixture::default()
    });
    let p = paragraphs(&document.body)[0];
    let runs = text_runs(p);
    assert_eq!(runs.len(), 3);
    let a = runs[0];
    assert_eq!(a.char_spacing, Some(1.0));
    assert_eq!(a.char_scale, Some(0.5));
    assert!(a.bold);
    assert_eq!(a.underline_style.as_deref(), Some("dotted"));
    // Raised so the 58% glyphs' ascent ends at the full-size ascent.
    assert!((a.position.unwrap() - 0.8 * 0.42 * 12.0).abs() < 1e-9);
    assert!((a.font_size - 0.58 * 12.0).abs() < 1e-9);
    assert!(runs[1].all_caps && runs[1].small_caps);
    assert_eq!(runs[2].text, "cd");
    assert_eq!(runs[2].font_size, 18.0);
}

#[test]
fn splits_script_classes_into_their_own_faces() {
    let document = read_fixture(Fixture {
        body: "<text:p>Ab 日本 שלום</text:p>",
        ..Fixture::default()
    });
    let runs = text_runs(paragraphs(&document.body)[0]);
    let faces: Vec<(&str, Option<&str>, f64)> = runs
        .iter()
        .map(|r| (r.text.as_str(), r.font_family.as_deref(), r.font_size))
        .collect();
    assert_eq!(
        faces,
        [
            ("Ab ", Some("Liberation Serif"), 12.0),
            ("日本 ", Some("Hiragino Mincho ProN"), 10.5),
            ("שלום", Some("Liberation Serif"), 12.0),
        ]
    );
    assert_eq!(runs[2].rtl, Some(true));
}

#[test]
fn maps_tables_with_spans_and_covered_cells() {
    let automatic = r##"<style:style style:name="Tb" style:family="table"><style:table-properties style:width="200pt" table:align="left"/></style:style><style:style style:name="Cl" style:family="table-column"><style:table-column-properties style:column-width="100pt"/></style:style><style:style style:name="Cell" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000" fo:padding="2pt" fo:background-color="#00ff00" style:vertical-align="middle"/></style:style><style:style style:name="R" style:family="table-row"><style:table-row-properties style:min-row-height="30pt"/></style:style>"##;
    let body = r##"<table:table table:style-name="Tb"><table:table-column table:style-name="Cl" table:number-columns-repeated="2"/><table:table-header-rows><table:table-row table:style-name="R"><table:table-cell table:style-name="Cell" table:number-columns-spanned="2"><text:p>head</text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table-header-rows><table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell/><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table><text:p>after</text:p>"##;
    let document = read_fixture(Fixture {
        automatic,
        body,
        ..Fixture::default()
    });
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("a table")
    };
    let table: &DocTable = table;
    assert_eq!(table.col_widths, [100.0, 100.0]);
    assert_eq!(table.rows.len(), 3);
    let head = &table.rows[0];
    assert!(head.is_header);
    assert_eq!(
        (head.row_height, head.row_height_rule.as_str()),
        (Some(30.0), "atLeast")
    );
    assert_eq!(head.cells.len(), 1);
    assert_eq!(head.cells[0].col_span, 2);
    assert_eq!(head.cells[0].background.as_deref(), Some("00ff00"));
    assert_eq!(head.cells[0].v_align, "center");
    assert_eq!(table.rows[1].cells[0].v_merge, Some(true));
    assert_eq!(table.rows[2].cells[0].v_merge, Some(false));
    assert_eq!(table.rows[2].cells.len(), 2);
    let CellElement::Paragraph(c) = &table.rows[2].cells[1].content[0] else {
        panic!()
    };
    assert_eq!(texts(c), ["c"]);
}

#[test]
fn page_headers_footers_and_first_page_regions() {
    let page_children = r##"<style:header-style><style:header-footer-properties fo:min-height="10pt" fo:margin-bottom="5pt"/></style:header-style><style:footer-style><style:header-footer-properties fo:min-height="8pt" fo:margin-top="4pt"/></style:footer-style>"##;
    let master = r##"<style:header><text:p>Head <text:page-number>1</text:page-number> of <text:page-count>2</text:page-count></text:p></style:header><style:header-first><text:p>First</text:p></style:header-first><style:footer><text:p>Foot</text:p></style:footer>"##;
    let document = read_fixture(Fixture {
        page_children,
        master,
        ..Fixture::default()
    });
    let mm = 72.0 / 25.4;
    assert!((document.section.header_distance - 20.0 * mm).abs() < 1e-6);
    assert!((document.section.margin_top - (20.0 * mm + 10.0)).abs() < 1e-6);
    assert!((document.section.margin_bottom - (20.0 * mm + 8.0)).abs() < 1e-6);
    assert!(document.section.title_page);
    let header = document.headers.default.as_ref().unwrap();
    assert_eq!(
        texts(paragraphs(&header.body)[0]),
        ["Head ", "<page>", " of ", "<numPages>"]
    );
    let first = document.headers.first.as_ref().unwrap();
    assert_eq!(texts(paragraphs(&first.body)[0]), ["First"]);
    // The footer has no first-page variant: its ordinary story.
    let footer_first = document.footers.first.as_ref().unwrap();
    assert_eq!(texts(paragraphs(&footer_first.body)[0]), ["Foot"]);
}

#[test]
fn page_style_changes_and_column_sections_become_docx_sections() {
    let styles_automatic = r##"<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="297mm" fo:page-height="210mm" fo:margin-top="10mm" fo:margin-bottom="10mm" fo:margin-left="10mm" fo:margin-right="10mm" style:num-format="i"/></style:page-layout>"##;
    let masters = r##"<style:master-page style:name="Wide" style:page-layout-name="pm2"/>"##;
    let automatic = r##"<style:style style:name="PW" style:family="paragraph" style:master-page-name="Wide"><style:paragraph-properties style:page-number="5"/></style:style><style:style style:name="S2" style:family="section"><style:section-properties><style:columns fo:column-count="2" fo:column-gap="12pt"/></style:section-properties></style:style>"##;
    let body = r##"<text:p>one</text:p><text:section text:name="s" text:style-name="S2"><text:p>two</text:p></text:section><text:p>three</text:p><text:p text:style-name="PW">four</text:p>"##;
    let document = read_fixture(Fixture {
        styles_automatic,
        masters,
        automatic,
        body,
        ..Fixture::default()
    });
    let outline: Vec<String> = document
        .body
        .iter()
        .map(|element| match element {
            BodyElement::Paragraph(p) => texts(p).join(""),
            BodyElement::SectionBreak { kind, columns, .. } => {
                format!("§{kind}:{}", columns.as_ref().map_or(1, |c| c.count))
            }
            _ => "?".into(),
        })
        .collect();
    assert_eq!(
        outline,
        [
            "one",
            "§nextPage:1",
            "two",
            "§continuous:2",
            "three",
            "§continuous:1",
            "four"
        ]
    );
    let section = &document.section;
    assert_eq!(section.section_start.as_deref(), Some("nextPage"));
    assert!((section.page_width - 297.0 * 72.0 / 25.4).abs() < 1e-6);
    let numbering = section.page_num_type.as_ref().unwrap();
    assert_eq!(
        (numbering.start, numbering.fmt.as_deref()),
        (Some(5), Some("lowerRoman"))
    );
}

#[test]
fn notes_become_docx_notes_with_their_marks() {
    let document = read_fixture(Fixture {
        body: r##"<text:p>Text<text:note text:id="ftn1" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>Note text</text:p></text:note-body></text:note>, more<text:note text:note-class="endnote"><text:note-citation text:label="*">*</text:note-citation><text:note-body><text:p>End</text:p></text:note-body></text:note></text:p>"##,
        ..Fixture::default()
    });
    let runs = text_runs(paragraphs(&document.body)[0]);
    let reference = runs.iter().find(|r| r.note_ref.is_some()).unwrap();
    let note_ref = reference.note_ref.as_ref().unwrap();
    assert_eq!(
        (note_ref.kind.as_str(), note_ref.id.as_str()),
        ("footnote", "1")
    );
    assert_eq!(document.footnotes.len(), 1);
    assert_eq!(document.endnotes.len(), 1);
    let note = paragraphs(&document.footnotes[0].content)[0];
    // The note's own number placeholder, then its text.
    let note_runs = text_runs(note);
    assert!(note_runs[0]
        .note_ref
        .as_ref()
        .is_some_and(|r| r.id.is_empty()));
    assert_eq!(note_runs.last().unwrap().text, "Note text");
    let custom = runs
        .iter()
        .filter(|r| r.note_ref.as_ref().is_some_and(|n| n.kind == "endnote"))
        .collect::<Vec<_>>();
    assert!(custom[0].note_ref.as_ref().unwrap().custom_mark_follows);
}

#[test]
fn frames_become_anchored_pictures_and_text_boxes() {
    let png = image::tests_png();
    let styles = r##"<style:style style:name="Graphics" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="center" style:horizontal-rel="paragraph" style:vertical-pos="from-top" style:vertical-rel="paragraph"/></style:style>"##;
    let automatic = r##"<style:style style:name="fr1" style:family="graphic" style:parent-style-name="Graphics"/><style:style style:name="fr2" style:family="graphic"><style:graphic-properties fo:background-color="#ffff00" fo:padding="3pt" fo:border="1pt solid #0000ff" style:wrap="none"/></style:style>"##;
    let body = r##"<text:p>before<draw:frame draw:style-name="fr1" text:anchor-type="paragraph" svg:y="12pt" svg:width="72pt" svg:height="36pt" draw:z-index="2"><draw:image xlink:href="Pictures/a.png"/></draw:frame><draw:frame text:anchor-type="as-char" svg:width="10pt" svg:height="10pt"><draw:image xlink:href="Pictures/a.png"/></draw:frame><draw:frame draw:style-name="fr2" text:anchor-type="char" svg:x="5pt" svg:y="6pt" svg:width="100pt" fo:min-height="20pt"><draw:text-box><text:p>boxed</text:p></draw:text-box></draw:frame></text:p>"##;
    let (document, resources) = read_fixture_resources(Fixture {
        styles,
        automatic,
        body,
        files: &[("Pictures/a.png", &png)],
        ..Fixture::default()
    });
    assert_eq!(resources.len(), 1);
    assert_eq!(resources[0].mime_type, "image/png");
    let p = paragraphs(&document.body)[0];
    assert_eq!(
        texts(p),
        ["before", "<host>", "<image>", "<image>", "<host>", "<shape>"]
    );
    let DocRun::Image(floating) = &p.runs[2] else {
        panic!()
    };
    assert!(floating.anchor);
    assert_eq!(floating.anchor_x_align.as_deref(), Some("center"));
    assert_eq!(floating.anchor_y_pt, 12.0);
    assert_eq!(floating.wrap_mode.as_deref(), Some("square"));
    let DocRun::Image(inline) = &p.runs[3] else {
        panic!()
    };
    assert!(!inline.anchor);
    let DocRun::Shape(shape) = &p.runs[5] else {
        panic!()
    };
    assert_eq!((shape.anchor_x_pt, shape.anchor_y_pt), (5.0, 6.0));
    assert_eq!(shape.wrap_mode.as_deref(), Some("topAndBottom"));
    assert_eq!(shape.text_inset_l, 4.0);
    assert_eq!(shape.text_autofit.as_deref(), Some("sp"));
    let TextBoxBlockWire::Body(BodyElement::Paragraph(boxed)) = &shape.text_box_content[0] else {
        panic!()
    };
    assert_eq!(texts(boxed), ["boxed"]);
}

#[test]
fn numbers_lists_and_headings() {
    let styles = r##"<text:list-style style:name="L"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="36pt" fo:text-indent="-18pt" fo:margin-left="36pt"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")" text:display-levels="2"><style:list-level-properties text:space-before="36pt" text:min-label-width="18pt"/></text:list-level-style-number></text:list-style><text:outline-style style:name="Outline"><text:outline-level-style text:level="1" style:num-format="I" style:num-suffix=" "><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="nothing"/></style:list-level-properties></text:outline-level-style></text:outline-style>"##;
    let body = r##"<text:h text:outline-level="1">Intro</text:h><text:list text:style-name="L"><text:list-item><text:p>one</text:p><text:list><text:list-item><text:p>sub</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>two</text:p></text:list-item></text:list><text:h text:outline-level="1">Next</text:h>"##;
    let document = read_fixture(Fixture {
        styles,
        body,
        ..Fixture::default()
    });
    let p = paragraphs(&document.body);
    let labels: Vec<Vec<String>> = p.iter().map(|p| texts(p)).collect();
    assert_eq!(labels[0], ["I ", "Intro"]);
    assert_eq!(labels[1], ["1.", "\t", "one"]);
    assert_eq!(labels[2], ["1.a)", "\t", "sub"]);
    assert_eq!(labels[3], ["2.", "\t", "two"]);
    assert_eq!(labels[4], ["II ", "Next"]);
    assert_eq!((p[1].indent_left, p[1].indent_first), (36.0, -18.0));
    // The label width and position mode: the text at space-before plus the
    // minimum label width, the label hanging before it.
    assert_eq!((p[2].indent_left, p[2].indent_first), (54.0, -18.0));
    assert_eq!(p[1].tab_stops[0].pos, 36.0);
}

#[test]
fn hyperlinks_tracked_deletions_and_fields() {
    let body = r##"<text:tracked-changes><text:changed-region text:id="c1"><text:deletion><office:change-info/><text:p>gone</text:p></text:deletion></text:changed-region></text:tracked-changes><text:p><text:a xlink:href="https://example.com/">link</text:a> <text:a xlink:href="#mark">inner</text:a><text:change-start text:change-id="c1"/>dropped<text:change-end text:change-id="c1"/> <text:date>2026-01-01</text:date><text:bookmark text:name="mark"/></text:p>"##;
    let document = read_fixture(Fixture {
        body,
        ..Fixture::default()
    });
    let p = paragraphs(&document.body)[0];
    let runs = text_runs(p);
    assert_eq!(runs[0].hyperlink.as_deref(), Some("https://example.com/"));
    assert_eq!(runs[2].hyperlink_anchor.as_deref(), Some("mark"));
    let all: String = runs.iter().map(|r| r.text.as_str()).collect();
    assert_eq!(all, "link inner 2026-01-01");
    assert_eq!(p.bookmarks, ["mark"]);
}

#[test]
fn custom_shapes_evaluate_their_geometry() {
    let automatic = r##"<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="none" style:wrap="run-through" style:run-through="background"/></style:style>"##;
    let body = r##"<text:p><draw:custom-shape draw:style-name="gr1" text:anchor-type="paragraph" svg:x="1pt" svg:y="2pt" svg:width="50pt" svg:height="50pt"><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="isosceles-triangle" draw:modifiers="10800" draw:enhanced-path="M $0 0 L 21600 21600 0 21600 Z N"/></draw:custom-shape>x</text:p>"##;
    let document = read_fixture(Fixture {
        automatic,
        body,
        ..Fixture::default()
    });
    let p = paragraphs(&document.body)[0];
    let shape = p
        .runs
        .iter()
        .find_map(|run| match run {
            DocRun::Shape(shape) => Some(shape),
            _ => None,
        })
        .unwrap();
    assert!(shape.behind_doc);
    assert_eq!(shape.stroke, None);
    assert!(
        matches!(&shape.fill, Some(docx_model::ShapeFill::Solid { color }) if color == "ff0000")
    );
    assert_eq!(shape.subpaths.len(), 1);
    assert!(
        matches!(shape.subpaths[0][0], docx_model::PathCmd::MoveTo { x, y } if x == 0.5 && y == 0.0)
    );
}

#[test]
fn deeply_nested_content_stays_bounded() {
    let groups = format!(
        "<text:p>{}<draw:rect text:anchor-type=\"paragraph\" svg:width=\"1pt\" svg:height=\"1pt\"/>{}x</text:p>",
        "<draw:g text:anchor-type=\"paragraph\">".repeat(60),
        "</draw:g>".repeat(60)
    );
    let sections = format!(
        "{}<text:p>deep</text:p>{}",
        "<text:section text:name=\"s\">".repeat(60),
        "</text:section>".repeat(60)
    );
    for body in [groups, sections] {
        let styles = format!(
            r#"<office:document-styles {NS} office:version="1.3"><office:styles/><office:master-styles><style:master-page style:name="Standard"/></office:master-styles></office:document-styles>"#
        );
        let content = format!(
            r#"<office:document-content {NS} office:version="1.3"><office:body><office:text>{body}</office:text></office:body></office:document-content>"#
        );
        let files = [
            ("META-INF/manifest.xml", manifest(ODT, "")),
            ("styles.xml", styles),
            ("content.xml", content),
        ];
        let files: Vec<(&str, &[u8])> = files.iter().map(|(n, b)| (*n, b.as_bytes())).collect();
        let package = OdfPackage::open(
            package_bytes(ODT, &files),
            OoxmlFormat::Docx,
            PackageLimits::default(),
            &MEDIA_TYPES,
        )
        .unwrap();
        assert!(read(&package).err().unwrap().contains("nests beyond"));
    }
}

#[test]
fn heading_styles_naming_a_list_style_override_outline_numbering() {
    let styles = r##"<style:style style:name="Plain" style:family="paragraph" style:list-style-name=""/><text:outline-style style:name="Outline"><text:outline-level-style text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="space"/></style:list-level-properties></text:outline-level-style></text:outline-style>"##;
    let document = read_fixture(Fixture {
        styles,
        body: r#"<text:h text:outline-level="1">Numbered</text:h><text:h text:style-name="Plain" text:outline-level="1">Plain</text:h>"#,
        ..Fixture::default()
    });
    let p = paragraphs(&document.body);
    assert_eq!(texts(p[0]), ["1.", " ", "Numbered"]);
    assert_eq!(texts(p[1]), ["Plain"]);
}

#[test]
fn page_backgrounds_draw_behind_every_page_and_paragraphs_indent_from_floats() {
    let document = read_fixture(Fixture {
        page: r##" fo:background-color="#ffeedd""##,
        body: "<text:p>x</text:p>",
        ..Fixture::default()
    });
    // No header region: the story carries only the background, and the
    // body margin is unchanged.
    assert!((document.section.margin_top - 20.0 * 72.0 / 25.4).abs() < 1e-6);
    let header = document.headers.default.as_ref().unwrap();
    let p = paragraphs(&header.body)[0];
    let DocRun::Shape(shape) = &p.runs[1] else {
        panic!()
    };
    assert!(shape.behind_doc);
    assert!((shape.width_pt - 210.0 * 72.0 / 25.4).abs() < 1e-6);
    assert!(
        matches!(&shape.fill, Some(docx_model::ShapeFill::Solid { color }) if color == "ffeedd")
    );
    assert!(paragraphs(&document.body)[0].indents_from_floats);
}

#[test]
fn an_unbalanced_section_ending_the_document_stays_the_last_section() {
    let automatic = r##"<style:style style:name="S2" style:family="section"><style:section-properties text:dont-balance-text-columns="true"><style:columns fo:column-count="2" fo:column-gap="12pt"/></style:section-properties></style:style>"##;
    let body = r##"<text:p>one</text:p><text:section text:name="s" text:style-name="S2"><text:p>two</text:p></text:section>"##;
    let document = read_fixture(Fixture {
        automatic,
        body,
        ..Fixture::default()
    });
    // No continuous section follows (it would balance the columns), and
    // LibreOffice's per-column footnote areas are requested.
    assert_eq!(document.section.columns.as_ref().map(|c| c.count), Some(2));
    assert!(document
        .settings
        .as_ref()
        .is_some_and(|s| s.footnote_column_bands));
}

#[test]
fn column_counts_beyond_the_limit_are_refused() {
    let automatic = r##"<style:style style:name="S" style:family="section"><style:section-properties><style:columns fo:column-count="4000000000"/></style:section-properties></style:style>"##;
    let body =
        r##"<text:section text:name="s" text:style-name="S"><text:p>a</text:p></text:section>"##;
    let error = try_read_fixture(Fixture {
        automatic,
        body,
        ..Fixture::default()
    })
    .err()
    .expect("refused");
    assert!(error.contains("column limit"), "{error}");
}

#[test]
fn the_initial_page_style_keeps_its_columns() {
    // §17.12 on the initial page style, with no master page named by the
    // body: its first section is laid out in the layout's columns.
    let page = r##"><style:columns fo:column-count="2" fo:column-gap="12pt"/></style:page-layout-properties><style:header-style"##;
    let document = read_fixture(Fixture {
        page,
        body: "<text:p>a</text:p>",
        ..Fixture::default()
    });
    assert_eq!(document.section.columns.as_ref().map(|c| c.count), Some(2));
}

#[test]
fn repeated_cells_stream_up_to_the_last_column() {
    let body = r##"<table:table><table:table-column table:number-columns-repeated="2"/><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="1000000"><text:p>x</text:p></table:table-cell><table:table-cell table:number-columns-repeated="1000000"/></table:table-row></table:table>"##;
    let document = read_fixture(Fixture {
        body,
        ..Fixture::default()
    });
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("a table")
    };
    assert_eq!(table.rows.len(), 3);
    // The row is wider than the declared columns: the grid grows to the
    // column bound and the repetitions stop there.
    assert_eq!(table.col_widths.len(), 1024);
    assert!(table.rows.iter().all(|row| row.cells.len() == 1024));
}

#[test]
fn cells_take_row_then_column_default_cell_styles() {
    // §16.2: the cell's style, then its row's default cell style, then its
    // column's, then the family default, property by property.
    let automatic = r##"<style:style style:name="Own" style:family="table-cell"><style:table-cell-properties fo:padding="1pt"/></style:style><style:style style:name="RowCell" style:family="table-cell"><style:table-cell-properties fo:background-color="#ff0000" fo:padding="3pt"/></style:style><style:style style:name="ColCell" style:family="table-cell"><style:table-cell-properties fo:background-color="#0000ff" style:vertical-align="middle" fo:border="0.5pt solid #000000"/></style:style>"##;
    let body = r##"<table:table><table:table-column table:default-cell-style-name="ColCell" table:number-columns-repeated="2"/><table:table-row table:default-cell-style-name="RowCell"><table:table-cell table:style-name="Own"><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><text:p>c</text:p></table:table-cell><table:table-cell table:style-name="Own"><text:p>d</text:p></table:table-cell></table:table-row></table:table>"##;
    let document = read_fixture(Fixture {
        automatic,
        body,
        ..Fixture::default()
    });
    let BodyElement::Table(table) = &document.body[0] else {
        panic!("a table")
    };
    let cell = |row: usize, column: usize| &table.rows[row].cells[column];
    // Own padding over the row default's; background from the row default
    // over the column default's; alignment and border from the column.
    assert_eq!(cell(0, 0).margin_top, Some(1.0));
    assert_eq!(cell(0, 0).background.as_deref(), Some("ff0000"));
    assert_eq!(cell(0, 0).v_align, "center");
    assert!(cell(0, 0).borders.top.is_some());
    assert_eq!(cell(0, 1).margin_top, Some(3.0));
    assert_eq!(cell(0, 1).background.as_deref(), Some("ff0000"));
    // Without a row default: the column default.
    assert_eq!(cell(1, 0).background.as_deref(), Some("0000ff"));
    assert_eq!(cell(1, 0).margin_top, Some(0.0));
    assert_eq!(cell(1, 1).margin_top, Some(1.0));
    assert_eq!(cell(1, 1).v_align, "center");
}

#[test]
fn hidden_paragraph_fields_hide_their_paragraph() {
    let body = r##"<text:p>shown</text:p><text:p><text:hidden-paragraph text:condition="ooow:1" text:is-hidden="true"/>hidden</text:p><text:p><text:span><text:hidden-paragraph text:condition="ooow:0" text:is-hidden="false"/>kept</text:span></text:p>"##;
    let document = read_fixture(Fixture {
        body,
        ..Fixture::default()
    });
    let shown: Vec<Vec<String>> = paragraphs(&document.body)
        .iter()
        .map(|p| texts(p))
        .collect();
    assert_eq!(shown, [vec!["shown".to_string()], vec!["kept".to_string()]]);
}

#[test]
fn phantom_list_ancestors_keep_the_counter_bound() {
    let styles = r##"<text:list-style style:name="L"><text:list-level-style-number text:level="1" style:num-format="a" style:num-letter-sync="true" text:start-value="4294967295"/><text:list-level-style-number text:level="2" style:num-format="a" style:num-letter-sync="true" text:start-value="4294967295"/><text:list-level-style-number text:level="3" style:num-format="a" style:num-letter-sync="true" text:start-value="4294967295" text:display-levels="3"/></text:list-style>"##;
    let body = r##"<text:list text:style-name="L"><text:list-item><text:list><text:list-item><text:list><text:list-item><text:p>deep</text:p></text:list-item></text:list></text:list-item></text:list></text:list-item></text:list>"##;
    let document = read_fixture(Fixture {
        styles,
        body,
        ..Fixture::default()
    });
    let label = texts(paragraphs(&document.body)[0])[0].clone();
    // Three components of at most MAX_COUNTED / 26 + 1 letters each.
    assert!(label.len() <= 3 * (1_000_000 / 26 + 2), "{}", label.len());
}

#[test]
fn hyperlink_targets_are_charged_per_split_run() {
    // One 4 MiB target around alternating Latin and Asian characters: each
    // emitted run owns a copy of the target, so the copies reach the
    // construction limit instead of retaining gigabytes.
    let href = "x".repeat(4 << 20);
    let text = "aあ".repeat(200);
    let body =
        format!(r##"<text:p><text:a xlink:href="https://e/{href}">{text}</text:a></text:p>"##);
    let error = try_read_fixture(Fixture {
        body: &body,
        ..Fixture::default()
    })
    .err()
    .expect("refused");
    assert!(error.contains("construction limit"), "{error}");
}

#[test]
fn repeated_rows_are_bounded_by_the_source_they_reread() {
    // A row whose cell holds about 100 KB of ignored markup, repeated 10,000
    // times: each pass rereads the markup, so the repetitions reach the
    // repetition bound instead of a billion node visits.
    let filler = "<text:soft-page-break/>".repeat(4_000);
    let body = format!(
        r##"<table:table><table:table-column/><table:table-row table:number-rows-repeated="10000"><table:table-cell>{filler}</table:table-cell></table:table-row></table:table>"##
    );
    let error = try_read_fixture(Fixture {
        body: &body,
        ..Fixture::default()
    })
    .err()
    .expect("refused");
    assert!(error.contains("repetition limit"), "{error}");
}

#[test]
fn column_redistribution_does_not_overflow() {
    // Authored lengths are unbounded: products past i64 saturate.
    let out = table::redistribute(&[i64::MAX / 2, i64::MAX / 2], i64::MAX);
    assert_eq!(out.len(), 2);
    assert_eq!(table::redistribute(&[100, 300], 800), [200, 600]);
}

#[test]
fn iri_dot_segments_normalize_linearly() {
    let href = format!("{}Pictures/a.png", "./".repeat(200_000));
    assert!(matches!(
        iri::resolve("content.xml", &href),
        Some(iri::Target::Package { path, .. }) if path == "Pictures/a.png"
    ));
}

#[test]
fn page_numbered_anchors_become_page_objects() {
    // §19.758: a page-anchored frame with a physical page belongs to that
    // page, from the body or from a master page's header (LibreOffice moves
    // both to the page); a page anchor without a page number stays with the
    // next paragraph.
    let frame = |page: &str| {
        format!(
            r##"<draw:frame text:anchor-type="page" {page} svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="1cm"><draw:text-box><text:p>box</text:p></draw:text-box></draw:frame>"##
        )
    };
    let master = format!(
        r##"<style:header><text:p>Head{}</text:p></style:header>"##,
        frame(r#"text:anchor-page-number="2""#)
    );
    let body = format!(
        "{}{}{}<text:p>text</text:p>",
        frame(r#"text:anchor-page-number="3""#),
        frame(r#"text:anchor-page-number="3""#),
        frame("")
    );
    let document = read_fixture(Fixture {
        master: &master,
        body: &body,
        ..Fixture::default()
    });
    let pages: Vec<(u32, usize)> = document
        .page_objects
        .iter()
        .map(|object| {
            let shapes = paragraphs(&object.body)
                .iter()
                .map(|host| {
                    host.runs
                        .iter()
                        .filter(|run| matches!(run, DocRun::Shape(_)))
                        .count()
                })
                .sum::<usize>();
            (object.page, shapes)
        })
        .collect();
    assert_eq!(pages, [(2, 1), (3, 2)]);
    // The unnumbered page anchor is carried by the body paragraph.
    let shapes_in_body = paragraphs(&document.body)[0]
        .runs
        .iter()
        .filter(|run| matches!(run, DocRun::Shape(_)))
        .count();
    assert_eq!(shapes_in_body, 1);
    // The header keeps its text but not the page object.
    let header = document.headers.default.as_ref().unwrap();
    let BodyElement::Paragraph(head) = &header.body[0] else {
        panic!()
    };
    assert!(!head.runs.iter().any(|run| matches!(run, DocRun::Shape(_))));
}

#[test]
fn style_anchors_and_groups_keep_their_page() {
    // §20.417-§20.418: the anchor type and page may come from the graphic
    // style; a page-anchored group keeps every member on its page.
    let automatic = r##"<style:style style:name="gp" style:family="graphic"><style:graphic-properties text:anchor-type="page"/></style:style>"##;
    let body = r##"<draw:frame draw:style-name="gp" text:anchor-page-number="3" svg:width="3cm" svg:height="1cm"><draw:text-box><text:p>x</text:p></draw:text-box></draw:frame><draw:g text:anchor-type="page" text:anchor-page-number="2"><draw:rect svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/><draw:rect svg:x="2cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/></draw:g><text:p>a</text:p>"##;
    let document = read_fixture(Fixture {
        automatic,
        body,
        ..Fixture::default()
    });
    let pages: Vec<(u32, usize)> = document
        .page_objects
        .iter()
        .map(|object| (object.page, object.body.len()))
        .collect();
    assert_eq!(pages, [(2, 2), (3, 1)]);
    // The group members take the group's page anchoring.
    for host in paragraphs(&document.page_objects[0].body) {
        let shape = host
            .runs
            .iter()
            .find_map(|run| match run {
                DocRun::Shape(shape) => Some(shape),
                _ => None,
            })
            .unwrap();
        assert_eq!(shape.anchor_y_relative_from.as_deref(), Some("page"));
    }
    assert!(!paragraphs(&document.body)[0]
        .runs
        .iter()
        .any(|run| matches!(run, DocRun::Shape(_))));
}

#[test]
fn page_anchors_past_the_page_bound_are_refused() {
    let body = r##"<draw:frame text:anchor-type="page" text:anchor-page-number="4000000000" svg:width="3cm" svg:height="1cm"><draw:text-box><text:p>x</text:p></draw:text-box></draw:frame><text:p>a</text:p>"##;
    let error = try_read_fixture(Fixture {
        body,
        ..Fixture::default()
    })
    .err()
    .expect("refused");
    assert!(error.contains("page limit"), "{error}");
}
