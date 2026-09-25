//! Alternative shape XML (`metroBlob`, MS-ODRAW 2.3.4.41, opid 0x3A9).
//!
//! PowerPoint 2007 and later store, per shape, an OPC package whose
//! `drs/shapexml.xml` is the shape's DrawingML (`p:sp`). MS-ODRAW says the
//! property SHOULD be ignored; implementation note 32 says Office 2007/2010
//! do not ignore it. Its text characters are masked (letters replaced by
//! spaces or underscores); the characters themselves come from the binary.
//!
//! Evidence from PowerPoint 16 PDF exports:
//! - controls whose binary is byte-identical and whose alternative XML alone
//!   was edited (package rebuilt at the same length) render the edit: a card
//!   fill changed to red and heading character spacing raised to 20 pt;
//! - Office-saved decks render character spacing, shrink-to-fit scales and
//!   per-paragraph indents that exist only in the alternative XML;
//! - a control whose binary shape properties were edited (preset adjust and
//!   fill color) while the alternative XML was kept renders the binary, and
//!   a deck whose alternative states 10.5 pt text over a binary 10 pt run
//!   renders without the alternative's run properties: the alternative is
//!   not used once it disagrees with the binary.
//!
//! The direct model therefore adopts the alternative only when it agrees with
//! the binary shape on everything the two record in common and the direct
//! model can compare: the geometry (same preset name, same adjust values
//! within a master-unit rounding tolerance, an omitted value standing for
//! the preset default, or custom geometry on both
//! sides), the untransformed position and size, rotation and flips, a solid
//! fill color when both sides have one, run font size, bold and italic where
//! both state them, and the text structure (the same paragraphs, runs and
//! line breaks at the same UTF-16 lengths). The
//! downrev checksums that PowerPoint stores beside the XML cannot be
//! recomputed (they are not a checksum of the binary records), so this
//! structural agreement stands in for them. Placeholders (no layout context),
//! shapes with relationship references (the package carries no media), math
//! runs and any parse or resource failure keep the binary projection. The
//! adopted shape keeps the binary's transform, identifier and characters.
use super::*;
use pptx_model::{ShapeElement, TextRun};

/// Implementation resource policy, not a format limit.
const MAX_BLOB_BYTES: usize = 4 * 1024 * 1024;
const MAX_PART_BYTES: u64 = 1024 * 1024;
const MAX_PACKAGE_BYTES: u64 = 4 * 1024 * 1024;
const MAX_THEME_BYTES: u64 = 2 * 1024 * 1024;
const MAX_THEME_ENTRIES: usize = 64;
/// One master unit (1/576 inch) rounds to at most 1588 EMU per coordinate.
const XFRM_TOLERANCE_EMU: i64 = 1588;
/// A 21600-based adjust value converts to 1/100000 with at most 4.63 units of
/// rounding per step; allow two steps.
const ADJUST_TOLERANCE: f64 = 10.0;

/// The main master's round-trip theme (MS-PPT RoundTripTheme12Atom, 0x040E)
/// and color map (RoundTripColorMapping12Atom, 0x040F).
pub(in crate::ppt) struct Theme {
    theme_xml: String,
    clr_map: Option<String>,
}

/// Read the round-trip theme of a main master's child records. Any malformed,
/// oversized or ambiguous package yields `None`: the alternative XML is then
/// never adopted for that master's slides.
pub(in crate::ppt) fn master_theme(records: &[Record<'_>]) -> Option<Theme> {
    let mut themes = records.iter().filter(|r| r.kind == 0x040e);
    let theme = themes.next()?;
    if themes.next().is_some() {
        return None;
    }
    let mut maps = records.iter().filter(|r| r.kind == 0x040f);
    let clr_map = match (maps.next(), maps.next()) {
        (Some(map), None) => Some(std::str::from_utf8(map.payload).ok()?.to_owned()),
        (None, None) => None,
        _ => return None,
    };
    Some(Theme {
        theme_xml: theme_part(theme.payload)?,
        clr_map,
    })
}

fn theme_part(package: &[u8]) -> Option<String> {
    use std::io::Read;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(package)).ok()?;
    if archive.len() > MAX_THEME_ENTRIES {
        return None;
    }
    let mut read = |name: &str| -> Option<String> {
        let entry = archive.by_name(name).ok()?;
        if entry.size() > MAX_THEME_BYTES || entry.encrypted() {
            return None;
        }
        let mut text = String::new();
        entry
            .take(MAX_THEME_BYTES + 1)
            .read_to_string(&mut text)
            .ok()?;
        (text.len() as u64 <= MAX_THEME_BYTES).then_some(text)
    };
    // Follow the package relationships: root -> theme manager -> theme.
    let target = |rels: &str, kind: &str| -> Option<String> {
        let doc = roxmltree::Document::parse(rels).ok()?;
        let mut found = doc.root_element().children().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type").is_some_and(|t| t.ends_with(kind))
        });
        let first = found.next()?.attribute("Target")?.to_owned();
        found.next().is_none().then_some(first)
    };
    let manager = target(&read("_rels/.rels")?, "/officeDocument")?;
    let manager = manager.trim_start_matches('/').to_owned();
    let (dir, name) = manager.rsplit_once('/')?;
    let theme = target(&read(&format!("{dir}/_rels/{name}.rels"))?, "/theme")?;
    if theme.contains("..") || theme.starts_with('/') {
        return None;
    }
    read(&format!("{dir}/{theme}"))
}

/// Adopt the alternative shape XML for `binary` when it agrees with it.
/// `leaf` is the binary shape's untransformed (group-local) transform,
/// `nested` whether it sits in a group's child coordinate space, and `text`
/// its binary characters, if any.
#[cfg(feature = "direct-ppt")]
pub(in crate::ppt) fn adopt(
    binary: &ShapeElement,
    leaf: &pptx_model::Transform,
    nested: bool,
    text: Option<&str>,
    blob: &[u8],
    theme: &Theme,
    work_budget: &mut usize,
    text_budget: &mut usize,
) -> Option<ShapeElement> {
    if blob.len() > MAX_BLOB_BYTES {
        return None;
    }
    // One parse attempt of work, and the package's declared inflated size
    // against the session's decoded-text byte budget. An exhausted budget
    // keeps the binary projection.
    *work_budget = work_budget.checked_sub(1)?;
    *text_budget = text_budget.checked_sub(inflated_size(blob)?)?;
    let parsed = pptx_parser::parse_standalone_shape_part(
        blob,
        "drs/shapexml.xml",
        &theme.theme_xml,
        theme.clr_map.as_deref(),
        MAX_PART_BYTES,
        MAX_PACKAGE_BYTES,
    )
    .ok()??;
    if parsed.placeholder || parsed.relationship_references {
        return None;
    }
    let mut shape = parsed.element;
    if !same_geometry(binary, &shape)
        || !same_transform(leaf, nested, &shape)
        || !same_solid_fill(binary, &shape)
    {
        return None;
    }
    if !same_run_formatting(binary, &shape) {
        return None;
    }
    substitute_text(&mut shape, text)?;
    shape.x = binary.x;
    shape.y = binary.y;
    shape.width = binary.width;
    shape.height = binary.height;
    shape.rotation = binary.rotation;
    shape.flip_h = binary.flip_h;
    shape.flip_v = binary.flip_v;
    shape.id = binary.id.clone();
    Some(shape)
}

/// Without the direct presentation model feature nothing is adopted.
#[cfg(not(feature = "direct-ppt"))]
pub(in crate::ppt) fn adopt(
    _binary: &ShapeElement,
    _leaf: &pptx_model::Transform,
    _nested: bool,
    _text: Option<&str>,
    _blob: &[u8],
    _theme: &Theme,
    _work_budget: &mut usize,
    _text_budget: &mut usize,
) -> Option<ShapeElement> {
    None
}

/// Sum of the package entries' declared uncompressed sizes, bounded by the
/// package policy. The PPTX package reader enforces the same limits while
/// inflating.
#[cfg(feature = "direct-ppt")]
fn inflated_size(blob: &[u8]) -> Option<usize> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).ok()?;
    let mut total = 0u64;
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index).ok()?;
        if entry.size() > MAX_PART_BYTES {
            return None;
        }
        total = total.checked_add(entry.size())?;
    }
    (total <= MAX_PACKAGE_BYTES).then(|| total as usize)
}

fn adjusts(shape: &ShapeElement) -> [Option<f64>; 8] {
    [
        shape.adj, shape.adj2, shape.adj3, shape.adj4, shape.adj5, shape.adj6, shape.adj7,
        shape.adj8,
    ]
}

fn same_geometry(binary: &ShapeElement, alternative: &ShapeElement) -> bool {
    if binary.geometry != alternative.geometry {
        return false;
    }
    if binary.geometry == "custGeom" {
        return binary.cust_geom.is_some() && alternative.cust_geom.is_some();
    }
    // An omitted adjust means the preset default, so an explicit value equal
    // to that default is the same geometry (ECMA-376 20.1.9.5).
    let defaults = crate::officeart::preset_defaults::defaults(&binary.geometry).unwrap_or(&[]);
    adjusts(binary)
        .iter()
        .zip(adjusts(alternative))
        .enumerate()
        .all(|(index, (a, b))| {
            let default = defaults.get(index).map(|&v| f64::from(v));
            match (a.or(default), b.or(default)) {
                (None, None) => true,
                (Some(a), Some(b)) => (a - b).abs() <= ADJUST_TOLERANCE,
                _ => false,
            }
        })
}

/// A group's child anchors (MS-ODRAW OfficeArtChildAnchor) are read with the
/// master-unit scale; the alternative XML stores the same child coordinate
/// values unscaled (`a:chOff`/`a:chExt` space), so compare them unscaled.
fn same_transform(leaf: &pptx_model::Transform, nested: bool, alternative: &ShapeElement) -> bool {
    let scale = if nested { 1587.5 } else { 1.0 };
    let close = |a: i64, b: i64| (a as f64 / scale - b as f64).abs() <= XFRM_TOLERANCE_EMU as f64;
    let turn = (leaf.rot - alternative.rotation).rem_euclid(360.0);
    close(leaf.x, alternative.x)
        && close(leaf.y, alternative.y)
        && close(leaf.cx, alternative.width)
        && close(leaf.cy, alternative.height)
        && turn.min(360.0 - turn) < 0.01
        && leaf.flip_h == alternative.flip_h
        && leaf.flip_v == alternative.flip_v
}

fn same_solid_fill(binary: &ShapeElement, alternative: &ShapeElement) -> bool {
    use pptx_model::Fill;
    match (&binary.fill, &alternative.fill) {
        (Some(Fill::Solid { color: a, .. }), Some(Fill::Solid { color: b, .. })) => {
            a.eq_ignore_ascii_case(b)
        }
        _ => true,
    }
}

/// Run formatting both forms state must agree character by character: font
/// size, bold and italic, where each side specifies them. The binary stores
/// whole points only; a corpus deck whose alternative says 10.5 pt over a
/// binary 10 pt renders in PowerPoint without the alternative's other run
/// properties (its character spacing), while a deck whose sizes agree
/// renders them. Paragraphs holding fields are compared structurally only,
/// because the binary projection substitutes field text.
fn same_run_formatting(binary: &ShapeElement, alternative: &ShapeElement) -> bool {
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return true;
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return false;
    }
    type Format = (Option<f64>, Option<bool>, Option<bool>);
    let spans = |runs: &[TextRun]| -> Option<Vec<(usize, Format)>> {
        runs.iter()
            .map(|run| match run {
                TextRun::Text(d) if d.field_type.is_none() => Some((
                    d.text.encode_utf16().count(),
                    (d.font_size, d.bold, d.italic),
                )),
                TextRun::Break => Some((1, (None, None, None))),
                _ => None,
            })
            .collect()
    };
    let agree = |x: Option<f64>, y: Option<f64>| match (x, y) {
        (Some(x), Some(y)) => (x - y).abs() < 1e-6,
        _ => true,
    };
    let agree_bool = |x: Option<bool>, y: Option<bool>| x.zip(y).is_none_or(|(x, y)| x == y);
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        let (Some(bs), Some(as_)) = (spans(&bp.runs), spans(&ap.runs)) else {
            continue;
        };
        let expand = |spans: &[(usize, Format)]| -> Vec<Format> {
            spans
                .iter()
                .flat_map(|&(n, f)| std::iter::repeat_n(f, n))
                .collect()
        };
        let (bx, ax) = (expand(&bs), expand(&as_));
        if bx.len() != ax.len() {
            return false;
        }
        if !bx
            .iter()
            .zip(&ax)
            .all(|(x, y)| agree(x.0, y.0) && agree_bool(x.1, y.1) && agree_bool(x.2, y.2))
        {
            return false;
        }
    }
    true
}

/// Replace the masked characters of the alternative text with the binary
/// characters. Every paragraph, run and line break must line up exactly at
/// UTF-16 lengths; a run boundary inside a surrogate pair does not.
fn substitute_text(shape: &mut ShapeElement, text: Option<&str>) -> Option<()> {
    let Some(body) = shape.text_body.as_mut() else {
        return text.is_none_or(str::is_empty).then_some(());
    };
    let text = text.unwrap_or("");
    let paragraphs: Vec<&str> = text.split('\r').collect();
    if paragraphs.len() != body.paragraphs.len() {
        // A body without text still has one empty paragraph in both forms.
        return None;
    }
    for (paragraph, source) in body.paragraphs.iter_mut().zip(paragraphs) {
        let units: Vec<u16> = source.encode_utf16().collect();
        let mut at = 0;
        for run in paragraph.runs.iter_mut() {
            match run {
                TextRun::Text(data) => {
                    let len = data.text.encode_utf16().count();
                    let slice = units.get(at..at + len)?;
                    if slice.iter().any(|&u| matches!(u, 0x0b | 0x0a | 0x2028)) {
                        return None;
                    }
                    data.text = String::from_utf16(slice).ok()?;
                    at += len;
                }
                TextRun::Break => {
                    if !matches!(units.get(at), Some(0x0b | 0x0a | 0x2028)) {
                        return None;
                    }
                    at += 1;
                }
                _ => return None,
            }
        }
        if at != units.len() {
            return None;
        }
    }
    Some(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn shape(geometry: &str) -> ShapeElement {
        serde_json::from_value(serde_json::json!({
            "x": 0, "y": 0, "width": 100, "height": 50, "rotation": 0.0,
            "flipH": false, "flipV": false, "geometry": geometry,
            "fill": null, "stroke": null, "textBody": null,
            "defaultTextColor": null, "custGeom": null,
            "adj": null, "adj2": null, "adj3": null, "adj4": null,
            "adj5": null, "adj6": null, "adj7": null, "adj8": null,
            "id": null, "name": null, "hyperlink": null,
            "placeholderType": null, "placeholderIdx": null
        }))
        .unwrap()
    }

    #[test]
    fn geometry_requires_same_preset_and_close_adjusts() {
        let mut binary = shape("roundRect");
        let mut alternative = shape("roundRect");
        assert!(same_geometry(&binary, &alternative));
        // roundRect's default adjust is 16667: an explicit default matches.
        binary.adj = Some(16667.0);
        assert!(same_geometry(&binary, &alternative));
        binary.adj = Some(30000.0);
        assert!(!same_geometry(&binary, &alternative));
        alternative.adj = Some(29990.0);
        assert!(same_geometry(&binary, &alternative));
        alternative.adj = Some(50000.0);
        assert!(!same_geometry(&binary, &alternative));
        assert!(!same_geometry(&shape("rect"), &shape("ellipse")));
        // An explicit default adjust equals an omitted one.
        let binary = shape("wedgeRoundRectCallout");
        let mut alternative = shape("wedgeRoundRectCallout");
        alternative.adj3 = Some(16667.0);
        assert!(same_geometry(&binary, &alternative));
        alternative.adj3 = Some(20000.0);
        assert!(!same_geometry(&binary, &alternative));
    }

    #[test]
    fn transform_allows_one_master_unit_of_rounding() {
        let leaf = pptx_model::Transform {
            x: 1000,
            y: 2000,
            cx: 3000,
            cy: 4000,
            ..Default::default()
        };
        let mut alternative = shape("rect");
        alternative.x = 1000 + 1588;
        alternative.y = 2000;
        alternative.width = 3000;
        alternative.height = 4000 - 1500;
        assert!(same_transform(&leaf, false, &alternative));
        alternative.x = 1000 + 1589;
        assert!(!same_transform(&leaf, false, &alternative));
        alternative.x = 1000;
        alternative.rotation = 90.0;
        assert!(!same_transform(&leaf, false, &alternative));
        // Angles compare modulo a full turn.
        let turned = pptx_model::Transform {
            rot: -90.0,
            ..leaf.clone()
        };
        alternative.rotation = 270.0;
        assert!(same_transform(&turned, false, &alternative));
        // Child anchors compare in the unscaled child space.
        let child = pptx_model::Transform {
            x: 1587 * 1000,
            y: 1587 * 2000,
            cx: 1587 * 3000,
            cy: 1587 * 2500,
            rot: 270.0,
            ..leaf
        };
        assert!(same_transform(&child, true, &alternative));
        assert!(!same_transform(&child, false, &alternative));
    }

    #[test]
    fn masked_text_takes_binary_characters_run_by_run() {
        let mut alternative = shape("rect");
        alternative.text_body = Some(
            serde_json::from_value(serde_json::json!({
                "verticalAnchor": "t", "defaultFontSize": null,
                "defaultBold": null, "defaultItalic": null,
                "lIns": 0, "rIns": 0, "tIns": 0, "bIns": 0,
                "wrap": "square", "vert": "horz", "autoFit": "none",
                "paragraphs": [
                    para(&[run("__"), serde_json::json!({"type": "break"}), run(" _ ")]),
                    para(&[run("____")])
                ]
            }))
            .unwrap(),
        );
        let mut adopted = alternative.clone();
        assert!(substitute_text(&mut adopted, Some("AB\u{b}C😀\rDEFG")).is_some());
        let texts: Vec<String> = adopted.text_body.unwrap().paragraphs[0]
            .runs
            .iter()
            .filter_map(|r| match r {
                TextRun::Text(d) => Some(d.text.clone()),
                _ => None,
            })
            .collect();
        assert_eq!(texts, ["AB", "C😀"]);
        // Length, paragraph and break mismatches keep the binary projection.
        for text in [
            "AB\u{b}C😀\rDEF",
            "AB\u{b}C😀",
            "ABXC😀\rDEFG",
            "AB\u{b}CD\u{b}\rDEFG",
        ] {
            assert!(
                substitute_text(&mut alternative.clone(), Some(text)).is_none(),
                "{text:?}"
            );
        }
        // A run boundary inside a surrogate pair is rejected.
        let mut split = alternative.clone();
        assert!(substitute_text(&mut split, Some("AB\u{b}😀C\rDEFG")).is_some());
        let mut split = alternative;
        if let Some(body) = split.text_body.as_mut() {
            body.paragraphs[0].runs =
                vec![TextRun::Text(text_data(" ")), TextRun::Text(text_data(" "))];
            body.paragraphs.truncate(1);
        }
        assert!(substitute_text(&mut split, Some("😀")).is_none());
    }

    #[test]
    fn run_formatting_must_agree_where_both_sides_state_it() {
        let body = |size: serde_json::Value| -> pptx_model::TextBody {
            let mut r = run("ABCD");
            r["fontSize"] = size;
            serde_json::from_value(serde_json::json!({
                "verticalAnchor": "t", "defaultFontSize": null,
                "defaultBold": null, "defaultItalic": null,
                "lIns": 0, "rIns": 0, "tIns": 0, "bIns": 0,
                "wrap": "square", "vert": "horz", "autoFit": "none",
                "paragraphs": [para(&[r])]
            }))
            .unwrap()
        };
        let mut binary = shape("rect");
        let mut alternative = shape("rect");
        binary.text_body = Some(body(serde_json::json!(10.0)));
        alternative.text_body = Some(body(serde_json::json!(10.0)));
        assert!(same_run_formatting(&binary, &alternative));
        // Whole-point binary sizes cannot state 10.5 pt: not the same shape.
        alternative.text_body = Some(body(serde_json::json!(10.5)));
        assert!(!same_run_formatting(&binary, &alternative));
        // An unstated size on either side is not a disagreement.
        binary.text_body = Some(body(serde_json::Value::Null));
        assert!(same_run_formatting(&binary, &alternative));
    }

    fn text_data(text: &str) -> pptx_model::TextRunData {
        serde_json::from_value(run(text)).unwrap_or_else(|_| unreachable!())
    }

    fn run(text: &str) -> serde_json::Value {
        serde_json::json!({
            "type": "text", "text": text, "bold": null, "italic": null,
            "underline": false, "strikethrough": false, "strikeDouble": false,
            "fontSize": null, "color": null, "fontFamily": null, "fieldType": null
        })
    }

    fn para(runs: &[serde_json::Value]) -> serde_json::Value {
        serde_json::json!({
            "alignment": "l", "marL": 0, "marR": 0, "indent": 0,
            "spaceBefore": null, "spaceAfter": null, "spaceLine": null, "lvl": 0,
            "bullet": {"type": "none"}, "defFontSize": null, "defColor": null,
            "defBold": null, "defItalic": null, "defFontFamily": null,
            "tabStops": [], "rtl": false, "eaLnBrk": true, "runs": runs
        })
    }
}
