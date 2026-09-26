//! ECMA-376 Part 2 §6.4 / §6.5.2.3 relationship-target resolution for every
//! part the DOCX parser loads, through the native, streaming and markdown
//! entry points.

use super::*;
use std::io::{Cursor, Write};
use zip::{write::SimpleFileOptions, ZipWriter};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

/// How every relationship in the fixture spells its target.
#[derive(Clone, Copy, Debug)]
enum Form {
    /// `footnotes.xml`
    Plain,
    /// `./footnotes.xml`
    Dot,
    /// `../word/footnotes.xml`
    DotDot,
    /// `/word/footnotes.xml`
    Absolute,
    /// `../WORD/FOOTNOTES.XML` (ASCII case-insensitive equivalence, §6.2.2.3)
    MixedCase,
    /// `%66%6F%6f...` (every letter percent-encoded: RFC 3986 §6.2.2.2)
    PercentEncoded,
    /// `x%2Ffootnotes.xml` (a percent-encoded `/` is not a part name)
    EncodedSlash,
    /// `%2E%2E/../footnotes.xml`: RFC 3986 §5 resolves the literal `..`
    /// against the `%2E%2E` segment before §6.2.2 normalization.
    EncodedDotDotThenDotDot,
    /// `a b/../footnotes.xml`: not an IRI reference (RFC 3987 §2.2), even
    /// though dot removal would drop the offending segment.
    InvalidCharacter,
    /// `https://example.com/word/footnotes.xml` (Internal, names no part)
    Scheme,
    /// `//example.com/word/footnotes.xml` (Internal, names no part)
    Authority,
    /// `footnotes.xml` with `TargetMode="External"`
    External,
    /// `missing-footnotes.xml` (a part name absent from the package)
    Missing,
}

/// Spell `part` (a ZIP part name) as a relationship target from a source
/// part living in `source_dir`. Returns the Target and the TargetMode.
fn spell(form: Form, source_dir: &str, part: &str) -> (String, &'static str) {
    let relative = part
        .strip_prefix(&format!("{source_dir}/"))
        .expect("fixture parts live below their source part's directory");
    let leaf_dir = source_dir.rsplit('/').next().unwrap_or(source_dir);
    match form {
        Form::Plain => (relative.to_string(), "Internal"),
        Form::Dot => (format!("./{relative}"), "Internal"),
        Form::DotDot => (format!("../{leaf_dir}/{relative}"), "Internal"),
        Form::Absolute => (format!("/{part}"), "Internal"),
        Form::MixedCase => (
            format!("../{}/{}", leaf_dir.to_uppercase(), relative.to_uppercase()),
            "Internal",
        ),
        Form::PercentEncoded => (
            relative
                .bytes()
                .enumerate()
                .map(
                    |(index, byte)| match (byte.is_ascii_alphabetic(), index % 2) {
                        (true, 0) => format!("%{byte:02x}"),
                        (true, _) => format!("%{byte:02X}"),
                        (false, _) => char::from(byte).to_string(),
                    },
                )
                .collect(),
            "Internal",
        ),
        Form::EncodedSlash => (format!("x%2F{relative}"), "Internal"),
        Form::EncodedDotDotThenDotDot => (format!("%2E%2E/../{relative}"), "Internal"),
        Form::InvalidCharacter => (format!("a b/../{relative}"), "Internal"),
        Form::Scheme => (format!("https://example.com/{part}"), "Internal"),
        Form::Authority => (format!("//example.com/{part}"), "Internal"),
        Form::External => (relative.to_string(), "External"),
        Form::Missing => {
            let (dir, file) = relative.rsplit_once('/').unwrap_or(("", relative));
            let missing = if dir.is_empty() {
                format!("missing-{file}")
            } else {
                format!("{dir}/missing-{file}")
            };
            (missing, "Internal")
        }
    }
}

fn rels(form: Form, source_dir: &str, entries: &[(&str, &str, &str)]) -> String {
    let body: String = entries
        .iter()
        .map(|(id, kind, part)| {
            let (target, mode) = spell(form, source_dir, part);
            let kind = if kind.starts_with("http") {
                (*kind).to_string()
            } else {
                format!("{R}/{kind}")
            };
            format!(
                r#"<Relationship Id="{id}" Type="{kind}" Target="{target}" TargetMode="{mode}"/>"#
            )
        })
        .collect();
    format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#
    )
}

fn inline_picture(rid: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="91440" cy="91440"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="91440" cy="91440"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

/// A package whose main document references one part of every kind the
/// parser loads, and whose headers, notes, fontTable and chart reference
/// their own related parts, every target spelled with `form`. The parts
/// themselves always exist at their conventional names.
fn package(form: Form) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>
          <w:p><w:pPr><w:pStyle w:val="Probe"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
            <w:commentRangeStart w:id="1"/><w:r><w:rPr><w:rFonts w:ascii="ProbeFont"/></w:rPr><w:t>BODY-TEXT</w:t></w:r><w:commentRangeEnd w:id="1"/>
            <w:r><w:commentReference w:id="1"/></w:r>
            <w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>
          <w:p>{picture}</w:p>
          <w:p>{svg_picture}</w:p>
          <w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="C"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
          <w:p><w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440"><v:shape id="ole" style="width:72pt;height:72pt"><v:imagedata r:id="rOlePreview"/></v:shape><o:OLEObject Type="Embed" ProgID="Package" ShapeID="ole" r:id="rOle"/></w:object></w:r></w:p>
          <w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:footerReference w:type="default" r:id="rFooter"/></w:sectPr>
        </w:body></w:document>"#,
        picture = inline_picture("rImage"),
        svg_picture = inline_picture("rSvg"),
    );
    let document_rels = rels(
        form,
        "word",
        &[
            ("rStyles", "styles", "word/styles.xml"),
            ("rNumbering", "numbering", "word/numbering.xml"),
            ("rSettings", "settings", "word/settings.xml"),
            ("rFontTable", "fontTable", "word/fontTable.xml"),
            ("rTheme", "theme", "word/theme/theme1.xml"),
            ("rHeader", "header", "word/header1.xml"),
            ("rFooter", "footer", "word/footer1.xml"),
            ("rFootnotes", "footnotes", "word/footnotes.xml"),
            ("rEndnotes", "endnotes", "word/endnotes.xml"),
            ("rComments", "comments", "word/comments.xml"),
            ("rImage", "image", "word/media/image1.png"),
            ("rOlePreview", "image", "word/media/image2.png"),
            // Stored as `word/media/image5.%73vg`, an equivalent item name.
            ("rSvg", "image", "word/media/image5.svg"),
            ("rOle", "oleObject", "word/embeddings/oleObject1.bin"),
            ("rChart", "chart", "word/charts/chart1.xml"),
        ],
    );
    let header = format!(
        r#"<w:hdr xmlns:w="{W}" xmlns:r="{R}"><w:p><w:r><w:t>HEADER-TEXT</w:t></w:r>{}</w:p></w:hdr>"#,
        inline_picture("rHeaderImage")
    );
    let footnotes = format!(
        r#"<w:footnotes xmlns:w="{W}" xmlns:r="{R}"><w:footnote w:id="1"><w:p><w:r><w:t>FOOTNOTE-TEXT</w:t></w:r>{}</w:p></w:footnote></w:footnotes>"#,
        inline_picture("rNoteImage")
    );
    let chart = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="2"/><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
    let chart_style = r#"<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cs:chartArea><cs:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></cs:spPr></cs:chartArea></cs:chartStyle>"#;
    let text_parts: Vec<(&str, String)> = vec![
        ("word/document.xml", document),
        ("word/_rels/document.xml.rels", document_rels),
        (
            "word/styles.xml",
            format!(
                r#"<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Probe"><w:rPr><w:sz w:val="46"/></w:rPr></w:style></w:styles>"#
            ),
        ),
        (
            "word/numbering.xml",
            format!(
                r#"<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="7"/><w:numFmt w:val="decimal"/><w:lvlText w:val="NUM-%1:"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
            ),
        ),
        (
            "word/settings.xml",
            format!(r#"<w:settings xmlns:w="{W}"><w:evenAndOddHeaders/></w:settings>"#),
        ),
        (
            "word/fontTable.xml",
            format!(
                r#"<w:fonts xmlns:w="{W}" xmlns:r="{R}"><w:font w:name="ProbeFont"><w:family w:val="roman"/><w:embedRegular r:id="rFont" w:fontKey="{{00000000-0000-0000-0000-000000000000}}"/></w:font></w:fonts>"#
            ),
        ),
        (
            "word/_rels/fontTable.xml.rels",
            rels(form, "word", &[("rFont", "font", "word/fonts/font1.odttf")]),
        ),
        (
            "word/theme/theme1.xml",
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="C"/><a:fontScheme name="F"><a:majorFont><a:latin typeface="ProbeMajor"/></a:majorFont><a:minorFont><a:latin typeface="ProbeMinor"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
                .to_string(),
        ),
        ("word/header1.xml", header),
        (
            "word/_rels/header1.xml.rels",
            rels(form, "word", &[("rHeaderImage", "image", "word/media/image3.png")]),
        ),
        (
            "word/footer1.xml",
            format!(r#"<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>FOOTER-TEXT</w:t></w:r></w:p></w:ftr>"#),
        ),
        ("word/footnotes.xml", footnotes),
        (
            "word/_rels/footnotes.xml.rels",
            rels(form, "word", &[("rNoteImage", "image", "word/media/image4.png")]),
        ),
        (
            "word/endnotes.xml",
            format!(
                r#"<w:endnotes xmlns:w="{W}"><w:endnote w:id="1"><w:p><w:r><w:t>ENDNOTE-TEXT</w:t></w:r></w:p></w:endnote></w:endnotes>"#
            ),
        ),
        (
            "word/comments.xml",
            format!(
                r#"<w:comments xmlns:w="{W}"><w:comment w:id="1" w:author="A"><w:p><w:r><w:t>COMMENT-TEXT</w:t></w:r></w:p></w:comment></w:comments>"#
            ),
        ),
        ("word/charts/chart1.xml", chart.to_string()),
        (
            "word/charts/_rels/chart1.xml.rels",
            rels(
                form,
                "word/charts",
                &[(
                    "rStyle",
                    "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
                    "word/charts/style1.xml",
                )],
            ),
        ),
        ("word/charts/style1.xml", chart_style.to_string()),
    ];
    let mut bytes = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut bytes));
        write_test_content_types(&mut zip);
        let options = SimpleFileOptions::default();
        for (name, text) in &text_parts {
            zip.start_file(*name, options).expect("in-memory zip");
            zip.write_all(text.as_bytes()).expect("in-memory zip");
        }
        for name in [
            "word/media/image1.png",
            "word/media/image2.png",
            "word/media/image3.png",
            "word/media/image4.png",
        ] {
            zip.start_file(name, options).expect("in-memory zip");
            zip.write_all(PNG_1X1).expect("in-memory zip");
        }
        zip.start_file("word/media/image5.%73vg", options)
            .expect("in-memory zip");
        zip.write_all(br#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>"#)
            .expect("in-memory zip");
        for name in ["word/fonts/font1.odttf", "word/embeddings/oleObject1.bin"] {
            zip.start_file(name, options).expect("in-memory zip");
            zip.write_all(b"\0").expect("in-memory zip");
        }
        zip.finish().expect("in-memory zip");
    }
    bytes
}

fn native(form: Form) -> serde_json::Value {
    serde_json::to_value(parse_from_bytes(&package(form)).expect("native parse"))
        .expect("document serializes")
}

fn streamed(form: Form) -> serde_json::Value {
    serde_json::to_value(
        parse_from_bytes_streamed_with_limits(&package(form), None, None, "parse")
            .expect("streamed parse"),
    )
    .expect("document serializes")
}

fn markdown(form: Form) -> String {
    crate::to_markdown_native(&package(form)).expect("markdown")
}

/// Observable evidence that each referenced part was found at its resolved
/// name. Paths are the stored ZIP item names, whatever the target spelling.
const LOADED_PART_EVIDENCE: &[(&str, &str)] = &[
    ("styles", r#""fontSize":23.0"#),
    ("numbering", "NUM-7:"),
    ("settings", r#""evenAndOddHeaders":true"#),
    ("fontTable", r#""ProbeFont":"roman""#),
    (
        "fontTable embedded font",
        r#""partPath":"word/fonts/font1.odttf""#,
    ),
    ("theme", r#""majorFont":"ProbeMajor""#),
    ("header", "HEADER-TEXT"),
    ("header image", r#""word/media/image3.png""#),
    ("footer", "FOOTER-TEXT"),
    ("footnotes", "FOOTNOTE-TEXT"),
    ("footnote image", r#""word/media/image4.png""#),
    ("endnotes", "ENDNOTE-TEXT"),
    ("comments", "COMMENT-TEXT"),
    ("image", r#""word/media/image1.png""#),
    ("embedded object preview", r#""word/media/image2.png""#),
    (
        "image stored under an equivalent item name",
        r#""imagePath":"word/media/image5.%73vg","mimeType":"image/svg+xml""#,
    ),
    ("chart", r#""lineWidthEmu":25400"#),
];

#[test]
fn every_part_kind_resolves_relative_absolute_and_equivalent_targets() {
    for form in [
        Form::Plain,
        Form::Dot,
        Form::DotDot,
        Form::Absolute,
        Form::MixedCase,
        Form::PercentEncoded,
        Form::EncodedDotDotThenDotDot,
    ] {
        for (api, json) in [("native", native(form)), ("streamed", streamed(form))] {
            let json = json.to_string();
            for (part, evidence) in LOADED_PART_EVIDENCE {
                assert!(
                    json.contains(evidence),
                    "{api} {form:?}: {part} was not loaded ({evidence})"
                );
            }
        }
        let markdown = markdown(form);
        for text in ["BODY-TEXT", "FOOTNOTE-TEXT", "ENDNOTE-TEXT", "COMMENT-TEXT"] {
            assert!(markdown.contains(text), "markdown {form:?}: {text}");
        }
    }
}

/// A target that names no part of this package (absolute IRI, network-path
/// reference, or External mode even when the spelling would resolve) reads
/// exactly like a relationship whose part is missing: the same omissions,
/// the same defaults, and never the conventional-name fallback.
#[test]
fn targets_naming_no_package_part_behave_like_missing_parts() {
    let missing_native = native(Form::Missing);
    let missing_streamed = streamed(Form::Missing);
    let missing_markdown = markdown(Form::Missing);
    let missing_text = missing_native.to_string();
    for (part, evidence) in LOADED_PART_EVIDENCE {
        assert!(
            !missing_text.contains(evidence),
            "missing {part} must not be loaded ({evidence})"
        );
    }
    for form in [
        Form::Scheme,
        Form::Authority,
        Form::External,
        Form::EncodedSlash,
        Form::InvalidCharacter,
    ] {
        assert_eq!(native(form), missing_native, "native {form:?}");
        assert_eq!(streamed(form), missing_streamed, "streamed {form:?}");
        assert_eq!(markdown(form), missing_markdown, "markdown {form:?}");
    }
}
