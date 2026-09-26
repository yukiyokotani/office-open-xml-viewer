//! Alternative shape XML (`metroBlob`, MS-ODRAW 2.3.4.41, opid 0x3A9).
//!
//! PowerPoint 2007 and later store, per shape, an OPC package holding the
//! shape's DrawingML. MS-ODRAW says the property SHOULD be ignored and that
//! Office deletes it when the shape is modified; implementation note 32 says
//! Office 2007/2010 do not ignore it. Its text characters are masked
//! (letters replaced by spaces or underscores); the characters themselves
//! come from the binary.
//!
//! Evidence from PowerPoint 16 PDF exports:
//! - controls whose binary is byte-identical and whose alternative XML alone
//!   was edited (package rebuilt at the same length) render the edit: a fill
//!   color changed in the XML only and character spacing raised in the XML
//!   only are drawn;
//! - Office-saved decks render character spacing, shrink-to-fit scales and
//!   per-paragraph indents that exist only in the alternative XML;
//! - a control whose binary shape properties were edited (preset adjust and
//!   fill color) while the alternative XML was kept renders the binary, and
//!   a deck whose alternative states 10.5 pt text over a binary 10 pt run
//!   renders without the alternative's run properties: the alternative is
//!   not used once it disagrees with the binary.
//!
//! The alternative therefore carries display information the binary lacks,
//! and the binary projection alone does not reproduce a shape whose
//! alternative PowerPoint uses. Each shape resolves to one of three
//! outcomes (`adopt`):
//! - no alternative part (a package with only the `downRev` checksums), or
//!   an ordinary alternative that verifiably disagrees with the binary on a
//!   compared attribute: the binary projection;
//! - an alternative that agrees on every applicable compared attribute: the
//!   alternative, with the binary's transform, identifier and characters;
//!   a binary freeform keeps its outline, and a placeholder inherits locally
//!   omitted shape properties from the binary projection;
//! - anything else fails closed as unsupported: an oversized, over-budget,
//!   ambiguous or unreadable package or theme, a part that is not a shape or
//!   connector, relationship references, and a compared attribute that the
//!   two forms state in ways this reader cannot equate. Placeholder shape
//!   properties omitted locally inherit from the binary projection; their
//!   transform always follows the binary anchor (PowerPoint 16 controls).
//!   A disagreement in a locally stated placeholder fill or geometry is also
//!   unsupported because its edited side cannot be determined.
//!
//! The compared attributes are those the evidence covers: geometry (preset
//! name and adjust values, or custom paths, except a binary freeform with an
//! XML preset), the untransformed position, size, rotation and flips for
//! ordinary shapes, the recorded fill when locally stated, run font size,
//! bold and italic where both state them, and the text structure (the same
//! paragraphs, runs and line breaks at the same UTF-16 lengths). Stroke,
//! effects and other text properties are not compared, so a binary edit that
//! preserves every compared attribute would still adopt a stale alternative.
//! The downrev checksums that PowerPoint stores beside the XML cannot be
//! recomputed (they are not a checksum of the binary records), so this
//! structural agreement stands in for them.
use super::*;
#[cfg(any(test, feature = "direct-ppt"))]
use pptx_model::{ShapeElement, TextRun};

/// Implementation resource policy, not a format limit.
#[cfg(feature = "direct-ppt")]
const MAX_BLOB_BYTES: usize = 4 * 1024 * 1024;
#[cfg(feature = "direct-ppt")]
const MAX_PART_BYTES: u64 = 1024 * 1024;
#[cfg(feature = "direct-ppt")]
const MAX_PACKAGE_BYTES: u64 = 4 * 1024 * 1024;
const MAX_THEME_BYTES: u64 = 2 * 1024 * 1024;
const MAX_THEME_ENTRIES: usize = 64;
/// One master unit is 1/576 inch.
#[cfg(any(test, feature = "direct-ppt"))]
const EMU_PER_MASTER_UNIT: f64 = 914_400.0 / 576.0;
/// One unit of the binary anchor's resolution, plus the EMU rounding of the
/// binary's master-unit conversion on each side.
#[cfg(any(test, feature = "direct-ppt"))]
const XFRM_TOLERANCE_UNITS: f64 = 1.0 + 1.0 / EMU_PER_MASTER_UNIT;
/// Normalized path coordinates are ratios of the same integer vertices and
/// extents on both sides; allow only floating-point rounding.
#[cfg(any(test, feature = "direct-ppt"))]
const PATH_TOLERANCE: f64 = 1e-9;

/// The main master's round-trip theme (MS-PPT RoundTripTheme12Atom, 0x040E)
/// and color map (RoundTripColorMapping12Atom, 0x040F), against which the
/// alternative XML's theme references resolve.
pub(in crate::ppt) enum Theme {
    // Read only by `adopt`.
    #[cfg_attr(not(feature = "direct-ppt"), allow(dead_code))]
    Readable {
        theme_xml: String,
        clr_map: Option<String>,
    },
    /// Malformed, oversized or ambiguous: no alternative XML on this
    /// master's slides can be resolved.
    Unreadable,
}

/// Read the round-trip theme of a main master's child records; `None` when
/// the master has none.
pub(in crate::ppt) fn master_theme(records: &[Record<'_>]) -> Option<Theme> {
    let mut themes = records.iter().filter(|r| r.kind == 0x040e);
    let theme = themes.next()?;
    let mut readable = || -> Option<Theme> {
        if themes.next().is_some() {
            return None;
        }
        let mut maps = records.iter().filter(|r| r.kind == 0x040f);
        let clr_map = match (maps.next(), maps.next()) {
            (Some(map), None) => Some(std::str::from_utf8(map.payload).ok()?.to_owned()),
            (None, None) => None,
            _ => return None,
        };
        Some(Theme::Readable {
            theme_xml: theme_part(theme.payload)?,
            clr_map,
        })
    };
    Some(readable().unwrap_or(Theme::Unreadable))
}

fn theme_part(package: &[u8]) -> Option<String> {
    use std::io::Read;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(package)).ok()?;
    if archive.len() > MAX_THEME_ENTRIES {
        return None;
    }
    let mut read = |name: &str| -> Option<String> {
        let index = crate::opc_part::entry_index(&archive, name)?;
        let entry = archive.by_index(index).ok()?;
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
    let target = |rels: &str, source: &str, kind: &str| -> Option<String> {
        let doc = roxmltree::Document::parse(rels).ok()?;
        let mut found = doc.root_element().children().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type").is_some_and(|t| t.ends_with(kind))
        });
        let first = found.next()?;
        if first
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
            || found.next().is_some()
        {
            return None;
        }
        ooxml_common::rels::resolve_part_name(source, first.attribute("Target")?)
    };
    let manager = target(&read("_rels/.rels")?, "", "/officeDocument")?;
    let (dir, name) = manager.rsplit_once('/').unwrap_or(("", &manager));
    let rels = if dir.is_empty() {
        format!("_rels/{name}.rels")
    } else {
        format!("{dir}/_rels/{name}.rels")
    };
    let theme = target(&read(&rels)?, &manager, "/theme")?;
    read(&theme)
}

/// What the binary shape records, for comparison with its alternative.
#[cfg(any(test, feature = "direct-ppt"))]
#[cfg_attr(not(feature = "direct-ppt"), allow(dead_code))]
pub(in crate::ppt) struct BinaryShape<'a> {
    /// The binary projection: geometry, adjust values and text runs.
    pub element: &'a ShapeElement,
    /// The untransformed (group-local) transform.
    pub leaf: &'a pptx_model::Transform,
    /// Whether the shape sits in a group's child coordinate space.
    pub nested: bool,
    /// The binary text characters, if any.
    pub text: Option<&'a str>,
    pub fill: RecordedFill,
    /// Authored per-path (fill, stroke) flags of custom geometry, before
    /// PowerPoint's open-path display rule (`officeart::geometry`).
    pub path_paint: Option<Vec<(bool, bool)>>,
    /// The range of each converted preset adjust value under the rounding
    /// of the binary anchor (`officeart::preset::adjustment_bounds`).
    pub adjust_bounds: [Option<(f64, f64)>; 8],
}

/// The fill a binary shape records (MS-ODRAW 2.3.7), as a model fill.
#[cfg(any(test, feature = "direct-ppt"))]
pub(in crate::ppt) enum RecordedFill {
    /// The geometry has no fill area (lines and connectors).
    NotDisplayed,
    /// No fill property is stated.
    Unstated,
    /// The stated fill; `None` when it is off.
    Stated(Option<Box<pptx_model::Fill>>),
    /// A stated fill this reader cannot restate as a model fill.
    Unknown,
}

/// Agreement of one attribute that both forms record.
#[cfg(any(test, feature = "direct-ppt"))]
#[derive(Clone, Copy, Debug, PartialEq)]
enum Verdict {
    Same,
    /// Both forms state the attribute, with different values.
    Differs(&'static str),
    /// Both forms state the attribute, but not in a form this reader can
    /// compare: agreement is unknown.
    Unverifiable(&'static str),
}

/// Adopt only when every compared attribute is the same. Any verified
/// difference keeps the binary projection: PowerPoint ignores an alternative
/// that disagrees with the binary shape. Otherwise, when some attribute
/// cannot be compared, the display is unknown and the shape fails closed.
#[cfg(any(test, feature = "direct-ppt"))]
fn decide(verdicts: &[Verdict]) -> Result<bool, String> {
    if verdicts.iter().any(|v| matches!(v, Verdict::Differs(_))) {
        return Ok(false);
    }
    match verdicts.iter().find_map(|v| match v {
        Verdict::Unverifiable(what) => Some(what),
        _ => None,
    }) {
        Some(what) => Err(unverifiable(what)),
        None => Ok(true),
    }
}

#[cfg(any(test, feature = "direct-ppt"))]
pub(in crate::ppt) fn unverifiable(what: &str) -> String {
    unsupported(format!(
        "PowerPoint alternative shape XML {what} cannot be compared with the binary shape"
    ))
}

/// The alternative part named by the package's root relationships. Office
/// writes one `downRev` relationship (the checksums) and at most one
/// relationship to the alternative DrawingML part; its type names the part's
/// kind (observed: shapeXml, connectorXml, groupShapeXml, pictureXml,
/// graphicFrameDoc, inkXml). A package with only `downRev` carries no
/// alternative.
#[cfg(feature = "direct-ppt")]
fn alternative_part(blob: &[u8]) -> Result<Option<(String, String)>, String> {
    use std::io::Read;
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let index = crate::opc_part::entry_index(&archive, "_rels/.rels").ok_or_else(unreadable)?;
    let entry = archive.by_index(index).map_err(|_| unreadable())?;
    if entry.size() > MAX_PART_BYTES || entry.encrypted() {
        return Err(unreadable());
    }
    let mut rels = String::new();
    entry
        .take(MAX_PART_BYTES + 1)
        .read_to_string(&mut rels)
        .map_err(|_| unreadable())?;
    let doc = roxmltree::Document::parse(&rels).map_err(|_| unreadable())?;
    let mut found = None;
    for node in doc.root_element().children().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Relationship" {
            return Err(unreadable());
        }
        let kind = node.attribute("Type").ok_or_else(unreadable)?;
        let kind = kind.rsplit_once('/').map_or(kind, |(_, kind)| kind);
        if kind == "downRev" {
            continue;
        }
        let target = ooxml_common::rels::resolve_part_name(
            "",
            node.attribute("Target").ok_or_else(unreadable)?,
        )
        .ok_or_else(|| unverifiable("package relationship"))?;
        if node
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
            || found.replace((kind.to_owned(), target)).is_some()
        {
            return Err(unverifiable("package relationship"));
        }
    }
    Ok(found)
}

/// Resolve the alternative shape XML of a binary shape: `Ok(Some)` adopts
/// it; `Ok(None)` keeps the binary projection, because the package carries
/// no alternative part or the alternative verifiably disagrees with the
/// binary; an error means agreement cannot be established (oversized, over
/// budget, unreadable, or not comparable), and the shape fails closed. The
/// adopted shape is charged to `model_budget` at its serialized size.
#[cfg(feature = "direct-ppt")]
pub(in crate::ppt) fn adopt(
    binary: &BinaryShape<'_>,
    blob: &[u8],
    theme: &Theme,
    work_budget: &mut usize,
    text_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    if blob.len() > MAX_BLOB_BYTES {
        return Err(unsupported(
            "PowerPoint alternative shape XML exceeds its size limit",
        ));
    }
    // One parse attempt of work, and the package's declared inflated size
    // against the session's decoded-text byte budget.
    *work_budget = work_budget
        .checked_sub(1)
        .ok_or_else(|| unsupported("PowerPoint alternative shape XML work budget exceeded"))?;
    super::charge_text(text_budget, inflated_size(blob)?)?;
    let Some((kind, part)) = alternative_part(blob)? else {
        return Ok(None);
    };
    // Only shapes and connectors are parts this projection can compare
    // with a binary shape.
    if !matches!(kind.as_str(), "shapeXml" | "connectorXml") {
        return Err(unverifiable("part"));
    }
    let Theme::Readable { theme_xml, clr_map } = theme else {
        return Err(unverifiable("theme"));
    };
    let parsed = pptx_parser::parse_standalone_shape_part(
        blob,
        &part,
        theme_xml,
        clr_map.as_deref(),
        MAX_PART_BYTES,
        MAX_PACKAGE_BYTES,
    );
    let parsed = parsed
        .map_err(|error| {
            unsupported(format!(
                "unreadable PowerPoint alternative shape XML: {error}"
            ))
        })?
        .ok_or_else(|| unverifiable("part"))?;
    // Relationship references (pictures, links) name parts this projection
    // does not compare. A PowerPoint 16 control changes the rendered crop
    // through an XML-only picture-fill edit, so keeping the binary here would
    // silently discard visible information.
    if parsed.relationship_references {
        return Err(unverifiable("relationship"));
    }
    let local = if parsed.placeholder {
        Some(placeholder_locals(blob, &part)?)
    } else {
        None
    };
    let mut shape = parsed.element;
    // PowerPoint 16 controls with a binary OfficeArt freeform and an XML
    // preset (frame and trapezoid, including visibly matching outlines)
    // retain the binary outline when the XML preset alone changes. An XML
    // fill edit on the same freeform is visible. This is geometry-specific
    // precedence, not rejection of the whole alternative. The converse
    // (binary preset, XML freeform) has not been established and stays closed.
    let binary_freeform_preset =
        binary.element.geometry == "custGeom" && shape.geometry != "custGeom";
    let substituted = {
        let mut candidate = shape.clone();
        substitute_text(&mut candidate, binary.text).map(|()| candidate)
    };
    // A local placeholder override is ambiguous when it disagrees: XML-only
    // geometry/fill edits render from the alternative, whereas binary-only
    // edits on the same controls invalidate it and render from the binary.
    // The proprietary checksum cannot be recomputed here to identify which
    // side was edited. Reject that pair rather than guessing its provenance.
    let local_override = |verdict, stated, attribute| {
        if stated && matches!(verdict, Verdict::Differs(_)) {
            Verdict::Unverifiable(attribute)
        } else {
            verdict
        }
    };
    let verdicts = [
        if binary_freeform_preset || local.is_some_and(|p| !p.geometry) {
            Verdict::Same
        } else {
            local_override(
                same_geometry(binary, &shape),
                parsed.placeholder,
                "placeholder geometry precedence",
            )
        },
        if parsed.placeholder {
            Verdict::Same
        } else {
            same_transform(binary.leaf, binary.nested, &shape)
        },
        if local.is_some_and(|p| !p.fill) {
            Verdict::Same
        } else {
            local_override(
                same_fill(&binary.fill, &shape, binary.element.rotation),
                parsed.placeholder,
                "placeholder fill precedence",
            )
        },
        same_run_formatting(binary.element, &shape),
        match &substituted {
            Ok(_) => Verdict::Same,
            Err(verdict) => *verdict,
        },
    ];
    if !decide(&verdicts)? {
        return Ok(None);
    }
    shape = substituted.unwrap_or_else(|_| unreachable!("text agreement was decided"));
    let element = binary.element;
    if binary_freeform_preset || local.is_some_and(|p| !p.geometry) {
        shape.geometry = element.geometry.clone();
        shape.cust_geom = element.cust_geom.clone();
        shape.cust_geom_paint = element.cust_geom_paint.clone();
        shape.adj = element.adj;
        shape.adj2 = element.adj2;
        shape.adj3 = element.adj3;
        shape.adj4 = element.adj4;
        shape.adj5 = element.adj5;
        shape.adj6 = element.adj6;
        shape.adj7 = element.adj7;
        shape.adj8 = element.adj8;
    }
    if let Some(local) = local {
        // ECMA-376 Part 1 Annex L.3.2.3: a placeholder takes absent shape
        // properties from its layout/master. The standalone DrawingML part
        // has no layout; the binary projection already holds those values.
        // PowerPoint 16 controls with title/body/subtitle placeholders show
        // XML-only font and bold edits but binary-anchor position/size. An
        // XML-only local fill or geometry edit is visible; a binary edit can
        // invalidate the blob, so explicit values still pass the ordinary
        // agreement gate above. Do not infer missing local values from the
        // standalone parser's defaults.
        if !local.fill {
            shape.fill = element.fill.clone();
        }
        if !local.stroke {
            shape.stroke = element.stroke.clone();
        }
        if !local.effects {
            shape.shadow = element.shadow.clone();
            shape.inner_shadow = element.inner_shadow.clone();
            shape.glow = element.glow.clone();
            shape.soft_edge = element.soft_edge.clone();
            shape.reflection = element.reflection.clone();
        }
    }
    shape.x = element.x;
    shape.y = element.y;
    shape.width = element.width;
    shape.height = element.height;
    shape.rotation = element.rotation;
    shape.flip_h = element.flip_h;
    shape.flip_v = element.flip_v;
    shape.id = element.id.clone();
    let retained = usize::try_from(
        ooxml_common::json_measurement::measure_json(&shape)
            .map_err(unsupported)?
            .json_bytes,
    )
    .map_err(|_| unsupported("PowerPoint direct slide model budget exceeded"))?;
    *model_budget = model_budget
        .checked_sub(retained)
        .ok_or_else(|| unsupported("PowerPoint direct slide model budget exceeded"))?;
    Ok(Some(shape))
}

#[cfg(feature = "direct-ppt")]
#[derive(Clone, Copy)]
struct PlaceholderLocals {
    geometry: bool,
    fill: bool,
    stroke: bool,
    effects: bool,
}

/// A standalone placeholder parser supplies schema defaults where the actual
/// shape relies on a missing layout. Inspect only the direct `p:spPr` children
/// to distinguish a local override from inherited geometry and fill. The
/// archive and part sizes have already passed `inflated_size` and the PPTX
/// standalone parser's bounds; this second read is limited to placeholders.
#[cfg(feature = "direct-ppt")]
fn placeholder_locals(blob: &[u8], part: &str) -> Result<PlaceholderLocals, String> {
    use std::io::Read;
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML placeholder");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let index = crate::opc_part::entry_index(&archive, part).ok_or_else(unreadable)?;
    let entry = archive.by_index(index).map_err(|_| unreadable())?;
    if entry.size() > MAX_PART_BYTES || entry.encrypted() {
        return Err(unreadable());
    }
    let mut xml = String::new();
    entry
        .take(MAX_PART_BYTES + 1)
        .read_to_string(&mut xml)
        .map_err(|_| unreadable())?;
    if xml.len() as u64 > MAX_PART_BYTES {
        return Err(unreadable());
    }
    let doc = roxmltree::Document::parse(&xml).map_err(|_| unreadable())?;
    let sp_pr = doc
        .root_element()
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "spPr");
    let has = |names: &[&str]| {
        sp_pr.is_some_and(|sp_pr| {
            sp_pr
                .children()
                .any(|n| n.is_element() && names.contains(&n.tag_name().name()))
        })
    };
    Ok(PlaceholderLocals {
        geometry: has(&["prstGeom", "custGeom"]),
        fill: has(&[
            "noFill",
            "solidFill",
            "gradFill",
            "pattFill",
            "blipFill",
            "grpFill",
        ]),
        stroke: has(&["ln"]),
        effects: has(&["effectLst", "effectDag"]),
    })
}

/// Without the direct presentation model feature there is no DrawingML
/// parser, so no alternative can be verified (only the feature-less unit
/// tests compile the direct model).
#[cfg(all(test, not(feature = "direct-ppt")))]
pub(in crate::ppt) fn adopt(
    _binary: &BinaryShape<'_>,
    _blob: &[u8],
    _theme: &Theme,
    _work_budget: &mut usize,
    _text_budget: &mut usize,
    _model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    Err(unverifiable("part"))
}

/// Sum of the package entries' declared uncompressed sizes, bounded by the
/// package policy. The PPTX package reader enforces the same limits while
/// inflating.
#[cfg(feature = "direct-ppt")]
fn inflated_size(blob: &[u8]) -> Result<usize, String> {
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let oversized = || unsupported("PowerPoint alternative shape XML exceeds its size limit");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(blob)).map_err(|_| unreadable())?;
    let mut total = 0u64;
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index).map_err(|_| unreadable())?;
        if entry.size() > MAX_PART_BYTES {
            return Err(oversized());
        }
        total = total.checked_add(entry.size()).ok_or_else(oversized)?;
    }
    if total > MAX_PACKAGE_BYTES {
        return Err(oversized());
    }
    usize::try_from(total).map_err(|_| oversized())
}

#[cfg(any(test, feature = "direct-ppt"))]
fn adjusts(shape: &ShapeElement) -> [Option<f64>; 8] {
    [
        shape.adj, shape.adj2, shape.adj3, shape.adj4, shape.adj5, shape.adj6, shape.adj7,
        shape.adj8,
    ]
}

/// Presets compare by name and adjust values (both in ECMA-376 1/100000
/// units), custom geometry path by path. A binary shape type maps to one
/// preset (`officeart::preset`), so two different preset names are
/// different shapes. PowerPoint saves a preset that has no MS-ODRAW shape
/// type as a binary freeform: a preset on one side and custom geometry on
/// the other would need the preset's formulas evaluated to compare. `adopt`
/// handles the observed binary-freeform/XML-preset precedence before calling
/// this comparator; the converse still reaches its unverifiable result.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_geometry(binary: &BinaryShape<'_>, alternative: &ShapeElement) -> Verdict {
    let element = binary.element;
    match (
        element.geometry == "custGeom",
        alternative.geometry == "custGeom",
    ) {
        (true, true) => return same_paths(binary, alternative),
        (false, false) if element.geometry != alternative.geometry => {
            return Verdict::Differs("preset geometry");
        }
        (false, false) => {}
        _ => {
            return Verdict::Unverifiable("preset geometry");
        }
    }
    // An omitted adjust means the preset default, so an explicit value equal
    // to that default is the same geometry (ECMA-376 20.1.9.5). The binary
    // value converts from a whole 21600-based unit on the binary anchor,
    // which rounds PowerPoint's own extent to a master unit; the
    // alternative's value, written in whole 1/100000 units from that exact
    // extent, lies within the range this rounding allows.
    let defaults = crate::officeart::preset_defaults::defaults(&element.geometry);
    let mut verdict = Verdict::Same;
    for (index, (bounds, value)) in binary
        .adjust_bounds
        .iter()
        .zip(adjusts(alternative))
        .enumerate()
    {
        let default = defaults.and_then(|d| d.get(index)).map(|&v| f64::from(v));
        match (bounds.or(default.map(|d| (d, d))), value.or(default)) {
            (None, None) => {}
            (Some((low, high)), Some(value)) => {
                if value < low - 1.0 || high + 1.0 < value {
                    return Verdict::Differs("preset adjust values");
                }
            }
            // One side omits a value whose default the preset definitions
            // do not state.
            _ => verdict = Verdict::Unverifiable("preset adjust values"),
        }
    }
    verdict
}

/// Custom geometry: both sides hold ECMA-376 20.1.9.14 path commands
/// normalized by their path extents (the binary's MS-ODRAW geoLeft..geoRight
/// and geoTop..geoBottom, the alternative's `a:path` w and h). A closing line
/// back to a subpath's start point immediately before its close draws
/// nothing that the close does not (20.1.9.3), and MS-ODRAW freeforms store
/// it explicitly while the alternative omits it, so it is dropped before
/// comparing. Paths, their paint and their command kinds must then match
/// one to one; a different vertex is a different shape, while a different
/// command structure may be another encoding of the same outline.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_paths(binary: &BinaryShape<'_>, alternative: &ShapeElement) -> Verdict {
    use pptx_model::PathCmd;
    let (Some(b), Some(a), Some(b_paint)) = (
        &binary.element.cust_geom,
        &alternative.cust_geom,
        &binary.path_paint,
    ) else {
        return Verdict::Unverifiable("custom geometry");
    };
    if b.len() != a.len() || b_paint.len() != b.len() {
        return Verdict::Unverifiable("custom geometry");
    }
    let mut verdict = Verdict::Same;
    for (index, (bp, ap)) in b.iter().zip(a).enumerate() {
        let a_paint = match alternative.cust_geom_paint.as_ref().map(|p| p.get(index)) {
            None => (true, true),
            Some(Some(p)) => match p.fill.as_deref() {
                None => (true, p.stroke),
                Some("none") => (false, p.stroke),
                // A lightened or darkened fill has no MS-ODRAW path flag.
                Some(_) => return Verdict::Unverifiable("custom geometry path fill"),
            },
            Some(None) => return Verdict::Unverifiable("custom geometry"),
        };
        if b_paint[index] != a_paint {
            verdict = Verdict::Differs("custom geometry path paint");
        }
        let (bp, ap) = (without_closing_lines(bp), without_closing_lines(ap));
        if bp.len() != ap.len() {
            return Verdict::Unverifiable("custom geometry");
        }
        for (bc, ac) in bp.iter().zip(&ap) {
            let pairs = match (bc, ac) {
                (PathCmd::MoveTo { x, y }, PathCmd::MoveTo { x: ax, y: ay })
                | (PathCmd::LineTo { x, y }, PathCmd::LineTo { x: ax, y: ay }) => {
                    vec![(x, ax), (y, ay)]
                }
                (
                    PathCmd::CubicBezTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    },
                    PathCmd::CubicBezTo {
                        x1: ax1,
                        y1: ay1,
                        x2: ax2,
                        y2: ay2,
                        x: ax,
                        y: ay,
                    },
                ) => vec![(x1, ax1), (y1, ay1), (x2, ax2), (y2, ay2), (x, ax), (y, ay)],
                (PathCmd::Close, PathCmd::Close) => Vec::new(),
                _ => return Verdict::Unverifiable("custom geometry"),
            };
            if pairs.iter().any(|(p, q)| (*p - *q).abs() > PATH_TOLERANCE) {
                verdict = Verdict::Differs("custom geometry vertices");
            }
        }
    }
    verdict
}

/// The path without a line to its subpath's start point directly before a
/// close.
#[cfg(any(test, feature = "direct-ppt"))]
fn without_closing_lines(path: &[pptx_model::PathCmd]) -> Vec<&pptx_model::PathCmd> {
    use pptx_model::PathCmd;
    let mut result: Vec<&PathCmd> = Vec::with_capacity(path.len());
    let mut start = None;
    for command in path {
        match command {
            PathCmd::MoveTo { x, y } => start = Some((*x, *y)),
            PathCmd::Close => {
                if let (Some(PathCmd::LineTo { x, y }), Some((sx, sy))) = (result.last(), start) {
                    if (x - sx).abs() <= PATH_TOLERANCE && (y - sy).abs() <= PATH_TOLERANCE {
                        result.pop();
                    }
                }
            }
            _ => {}
        }
        result.push(command);
    }
    result
}

/// Both sides are compared in the binary anchor's own unit, where one unit
/// is its resolution: master units (1/576 inch; the alternative states EMU)
/// for a top-level shape, and the group's child coordinate units for a group
/// child (MS-ODRAW OfficeArtChildAnchor, which the alternative states as the
/// same unscaled `a:chOff`/`a:chExt` space values). The binary transform was
/// converted to EMU with the master-unit scale either way.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_transform(
    leaf: &pptx_model::Transform,
    nested: bool,
    alternative: &ShapeElement,
) -> Verdict {
    let alternative_unit = if nested { 1.0 } else { EMU_PER_MASTER_UNIT };
    let close = |a: i64, b: i64| {
        (a as f64 / EMU_PER_MASTER_UNIT - b as f64 / alternative_unit).abs() <= XFRM_TOLERANCE_UNITS
    };
    let turn = (leaf.rot - alternative.rotation).rem_euclid(360.0);
    let same = close(leaf.x, alternative.x)
        && close(leaf.y, alternative.y)
        && close(leaf.cx, alternative.width)
        && close(leaf.cy, alternative.height)
        && turn.min(360.0 - turn) < 0.01
        && leaf.flip_h == alternative.flip_h
        && leaf.flip_v == alternative.flip_v;
    if same {
        Verdict::Same
    } else {
        Verdict::Differs("transform")
    }
}

/// Solid colors compare exactly (RGB with any alpha suffix). A fill that is
/// off on one side and solid on the other is a different paint. Other
/// fill kinds on either side have no common model representation this
/// reader can equate (PowerPoint may restate a binary fill as another
/// DrawingML kind), except identical gradient and pattern projections. In a
/// PowerPoint 16 control, an XML pattern over a binary duotone image changes
/// when only its XML foreground changes; editing the binary uses its image.
/// That pair stays unverifiable. So does a rotated gradient whose binary says
/// `rotWithShape=false` while XML omits it: the schema has no default.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_fill(binary: &RecordedFill, alternative: &ShapeElement, rotation: f64) -> Verdict {
    use pptx_model::Fill;
    let paint = |fill: Option<&Fill>| match fill {
        None | Some(Fill::None) => None,
        Some(fill) => Some(fill.clone()),
    };
    let alternative = paint(alternative.fill.as_ref());
    let binary = match binary {
        RecordedFill::NotDisplayed => return Verdict::Same,
        RecordedFill::Unstated if alternative.is_none() => return Verdict::Same,
        RecordedFill::Unstated | RecordedFill::Unknown => return Verdict::Unverifiable("fill"),
        RecordedFill::Stated(fill) => paint(fill.as_deref()),
    };
    match (&binary, &alternative) {
        (None, None) => Verdict::Same,
        (Some(Fill::Solid { color: a }), Some(Fill::Solid { color: b })) => {
            match same_color(a, b) {
                Some(true) => Verdict::Same,
                Some(false) => Verdict::Differs("solid fill color"),
                None => Verdict::Unverifiable("fill"),
            }
        }
        (None, Some(Fill::Solid { .. })) | (Some(Fill::Solid { .. }), None) => {
            Verdict::Differs("fill")
        }
        (Some(a), Some(b)) if same_paint(a, b, rotation) => Verdict::Same,
        _ => Verdict::Unverifiable("fill"),
    }
}

/// Model colors are hex RGB with an optional alpha byte. The binary stores
/// the displayed RGB, while an alternative color may be a theme color with
/// transforms (ECMA-376 20.1.2.3) that this reader resolves itself; two
/// roundings of one computed channel differ by at most one unit (a corpus
/// alternative's 50% gray resolves to 0x80 where the binary stores 0x7F).
/// Alpha converts from 16.16 (MS-ODRAW) and 1/100000 (DrawingML) alike.
/// `None`: not a model color.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_color(binary: &str, alternative: &str) -> Option<bool> {
    let channels = |color: &str| -> Option<Vec<i16>> {
        if !matches!(color.len(), 6 | 8) || !color.is_ascii() {
            return None;
        }
        (0..color.len())
            .step_by(2)
            .map(|at| i16::from_str_radix(&color[at..at + 2], 16).ok())
            .collect()
    };
    let (mut b, mut a) = (channels(binary)?, channels(alternative)?);
    // An absent alpha byte is opaque.
    b.resize(4, 0xff);
    a.resize(4, 0xff);
    Some(b.iter().zip(&a).all(|(x, y)| (x - y).abs() <= 1))
}

/// Identical gradient or pattern paint. Gradient stop positions are 16.16
/// fractions in the binary (MS-ODRAW 2.3.7.17 fillShadeColors) and
/// 1/100000 in DrawingML (ECMA-376 20.1.8.36), so each may round by one
/// step of either unit; angles are whole degrees on both sides. PowerPoint 16
/// uses the XML in an Office-saved multi-stop gradient whose binary duplicates
/// a terminal color and whose XML has fewer, shifted stops; changing only the
/// binary fill color makes it use the binary instead. That restatement has no
/// established general mapping, so unequal stop arrays remain unverifiable.
/// ECMA-376 CT_GradientFillProperties gives `flip` the default `none`; a
/// `tileRect` with all zero offsets covers the whole shape, like no tileRect.
/// `rotWithShape` has no visible effect when the projected shape has no
/// rotation; keep its absent/present distinction for rotated shapes because
/// the schema declares no default.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_paint(binary: &pptx_model::Fill, alternative: &pptx_model::Fill, rotation: f64) -> bool {
    use pptx_model::Fill;
    const POSITION: f64 = 1.0 / 65536.0 + 1.0 / 100_000.0;
    fn sorted(stops: &[pptx_model::GradStop]) -> Vec<(f64, &str)> {
        let mut stops: Vec<(f64, &str)> = stops
            .iter()
            .map(|s| (s.position, s.color.as_str()))
            .collect();
        stops.sort_by(|a, b| a.0.total_cmp(&b.0));
        stops
    }
    let rect = |r: &Option<ooxml_common::fill::FillRect>| r.as_ref().map(|r| [r.l, r.t, r.r, r.b]);
    let tile_rect =
        |r: &Option<ooxml_common::fill::FillRect>| rect(r).unwrap_or([0.0, 0.0, 0.0, 0.0]);
    let rotation_irrelevant = rotation.rem_euclid(360.0) == 0.0;
    match (binary, alternative) {
        (
            Fill::Gradient {
                stops: bs,
                angle: ba,
                grad_type: bt,
                scaled: bsc,
                path: bp,
                fill_to_rect: bf,
                tile_rect: btr,
                flip: bfl,
                rot_with_shape: br,
            },
            Fill::Gradient {
                stops: as_,
                angle: aa,
                grad_type: at,
                scaled: asc,
                path: ap,
                fill_to_rect: af,
                tile_rect: atr,
                flip: afl,
                rot_with_shape: ar,
            },
        ) => {
            let (bs, as_) = (sorted(bs), sorted(as_));
            let turn = (ba - aa).rem_euclid(360.0);
            bs.len() == as_.len()
                && bs.iter().zip(&as_).all(|(b, a)| {
                    (b.0 - a.0).abs() <= POSITION && same_color(b.1, a.1) == Some(true)
                })
                && turn.min(360.0 - turn) < 0.01
                && bt == at
                && bsc == asc
                && bp == ap
                && rect(bf) == rect(af)
                && tile_rect(btr) == tile_rect(atr)
                && bfl.as_deref().unwrap_or("none") == afl.as_deref().unwrap_or("none")
                && (br == ar || rotation_irrelevant)
        }
        (
            Fill::Pattern {
                fg: bf,
                bg: bb,
                preset: bp,
            },
            Fill::Pattern {
                fg: af,
                bg: ab,
                preset: ap,
            },
        ) => same_color(bf, af) == Some(true) && same_color(bb, ab) == Some(true) && bp == ap,
        _ => false,
    }
}

/// Run formatting both forms state must agree character by character: font
/// size, bold and italic, where each side specifies them. The binary stores
/// whole points only; a corpus deck whose alternative says 10.5 pt over a
/// binary 10 pt renders in PowerPoint without the alternative's other run
/// properties (its character spacing), while a deck whose sizes agree
/// renders them. A field compares as one unit, because the binary
/// projection substitutes its displayed text.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_run_formatting(binary: &ShapeElement, alternative: &ShapeElement) -> Verdict {
    #[derive(Clone, Copy, PartialEq)]
    enum Unit {
        Character,
        Field,
        Break,
    }
    type Format = (Option<f64>, Option<bool>, Option<bool>);
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return Verdict::Same;
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return Verdict::Differs("paragraph structure");
    }
    let units = |runs: &[TextRun]| -> Option<Vec<(Unit, Format)>> {
        let mut units = Vec::new();
        for run in runs {
            match run {
                TextRun::Text(d) => {
                    let format = (d.font_size, d.bold, d.italic);
                    if d.field_type.is_some() {
                        units.push((Unit::Field, format));
                    } else {
                        units.extend(std::iter::repeat_n(
                            (Unit::Character, format),
                            d.text.encode_utf16().count(),
                        ));
                    }
                }
                TextRun::Break => units.push((Unit::Break, (None, None, None))),
                TextRun::Math { .. } => return None,
            }
        }
        Some(units)
    };
    let agree = |x: Option<f64>, y: Option<f64>| match (x, y) {
        (Some(x), Some(y)) => (x - y).abs() < 1e-6,
        _ => true,
    };
    let agree_bool = |x: Option<bool>, y: Option<bool>| x.zip(y).is_none_or(|(x, y)| x == y);
    let mut verdict = Verdict::Same;
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        let (Some(bx), Some(ax)) = (units(&bp.runs), units(&ap.runs)) else {
            return Verdict::Unverifiable("equation");
        };
        // A masked alternative states a line break as one text character
        // (see `substitute_text`); its formatting still aligns.
        let kind = |unit: Unit| {
            if unit == Unit::Break {
                Unit::Character
            } else {
                unit
            }
        };
        if bx.len() != ax.len() || bx.iter().zip(&ax).any(|(x, y)| kind(x.0) != kind(y.0)) {
            // `substitute_text` decides the text structure.
            continue;
        }
        if !bx.iter().zip(&ax).all(|(x, y)| {
            agree(x.1 .0, y.1 .0) && agree_bool(x.1 .1, y.1 .1) && agree_bool(x.1 .2, y.1 .2)
        }) {
            verdict = Verdict::Differs("run formatting");
        }
    }
    verdict
}

/// Replace the masked characters of the alternative text with the binary
/// characters. Every paragraph, run and line break must line up exactly at
/// UTF-16 lengths; a run boundary inside a surrogate pair does not.
///
/// XML 1.0 cannot carry the vertical tab (U+000B) that breaks a line inside
/// a binary paragraph (MS-PPT 2.9.43), and the masked alternative states each
/// one as a masked character of the enclosing run, not as `a:br` (every
/// corpus alternative; both runs of one character and breaks leading a
/// longer run). Each such character becomes a line break that splits its run
/// into runs of the same formatting. Other break characters inside a run
/// have no observed form and are not compared. An equation has no binary
/// characters to compare.
#[cfg(any(test, feature = "direct-ppt"))]
fn substitute_text(shape: &mut ShapeElement, text: Option<&str>) -> Result<(), Verdict> {
    const DIFFERS: Verdict = Verdict::Differs("text structure");
    let Some(body) = shape.text_body.as_mut() else {
        return if text.is_none_or(str::is_empty) {
            Ok(())
        } else {
            Err(DIFFERS)
        };
    };
    let text = text.unwrap_or("");
    let paragraphs: Vec<&str> = text.split('\r').collect();
    if paragraphs.len() != body.paragraphs.len() {
        // A body without text still has one empty paragraph in both forms.
        return Err(DIFFERS);
    }
    for (paragraph, source) in body.paragraphs.iter_mut().zip(paragraphs) {
        let units: Vec<u16> = source.encode_utf16().collect();
        let mut at = 0usize;
        let mut runs = Vec::with_capacity(paragraph.runs.len());
        for run in std::mem::take(&mut paragraph.runs) {
            match run {
                TextRun::Text(mut data) => {
                    let len = data.text.encode_utf16().count();
                    let slice = at
                        .checked_add(len)
                        .and_then(|end| units.get(at..end))
                        .ok_or(DIFFERS)?;
                    at += len;
                    if slice.iter().any(|&u| matches!(u, 0x0a | 0x2028))
                        || (data.field_type.is_some() && slice.contains(&0x0b))
                    {
                        return Err(Verdict::Unverifiable("line break"));
                    }
                    let mut pieces = slice.split(|&u| u == 0x0b).peekable();
                    while let Some(piece) = pieces.next() {
                        if !piece.is_empty() {
                            let mut part = data.clone();
                            part.text = String::from_utf16(piece).map_err(|_| DIFFERS)?;
                            runs.push(TextRun::Text(part));
                        }
                        if pieces.peek().is_some() {
                            runs.push(TextRun::Break);
                        }
                    }
                    if slice.is_empty() {
                        data.text.clear();
                        runs.push(TextRun::Text(data));
                    }
                }
                TextRun::Break => {
                    if !matches!(units.get(at), Some(0x0b | 0x0a | 0x2028)) {
                        return Err(DIFFERS);
                    }
                    at += 1;
                    runs.push(TextRun::Break);
                }
                TextRun::Math { .. } => return Err(Verdict::Unverifiable("equation")),
            }
        }
        if at != units.len() {
            return Err(DIFFERS);
        }
        paragraph.runs = runs;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pptx_model::{Fill, PathCmd};

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

    fn binary<'a>(element: &'a ShapeElement, leaf: &'a pptx_model::Transform) -> BinaryShape<'a> {
        BinaryShape {
            element,
            leaf,
            nested: false,
            text: None,
            fill: RecordedFill::Stated(element.fill.clone().map(Box::new)),
            path_paint: element
                .cust_geom
                .as_ref()
                .map(|paths| vec![(true, true); paths.len()]),
            adjust_bounds: adjusts(element).map(|v| v.map(|v| (v, v))),
        }
    }

    fn solid(color: &str) -> Option<Fill> {
        Some(Fill::Solid {
            color: color.to_owned(),
        })
    }

    #[test]
    fn a_verified_difference_keeps_the_binary_and_an_unknown_one_fails_closed() {
        use Verdict::*;
        assert_eq!(decide(&[Same, Same]), Ok(true));
        assert_eq!(
            decide(&[Same, Differs("fill"), Unverifiable("geometry")]),
            Ok(false)
        );
        let error = decide(&[Same, Unverifiable("fill")]).unwrap_err();
        assert!(
            error.starts_with("UNSUPPORTED:") && error.contains("fill"),
            "{error}"
        );
    }

    #[test]
    fn preset_adjusts_compare_within_the_rounding_of_their_binary_inputs() {
        let leaf = pptx_model::Transform::default();
        let element = shape("roundRect");
        let mut alternative = shape("roundRect");
        let mut shape = binary(&element, &leaf);
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        // An omitted adjust is the preset default (roundRect: 16667).
        shape.adjust_bounds[0] = Some((16660.0, 16670.0));
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        shape.adjust_bounds[0] = Some((29000.0, 31000.0));
        assert_eq!(
            same_geometry(&shape, &alternative),
            Verdict::Differs("preset adjust values")
        );
        alternative.adj = Some(31001.0);
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        alternative.adj = Some(31002.0);
        assert_ne!(same_geometry(&shape, &alternative), Verdict::Same);
        // Different presets are different shapes; a preset against a
        // freeform would need the preset formulas evaluated.
        let other = self::shape("ellipse");
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &other),
            Verdict::Differs("preset geometry")
        );
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &self::shape("custGeom")),
            Verdict::Unverifiable("preset geometry")
        );
        // The preset definitions state no upArrow defaults (ECMA-376 lists
        // upDownArrow twice instead), so an omitted value is unknown.
        let up = self::shape("upArrow");
        let mut shape = binary(&up, &leaf);
        shape.adjust_bounds[1] = Some((50000.0, 50000.0));
        assert_eq!(
            same_geometry(&shape, &self::shape("upArrow")),
            Verdict::Unverifiable("preset adjust values")
        );
    }

    #[test]
    fn custom_paths_compare_vertices_paint_and_structure() {
        let leaf = pptx_model::Transform::default();
        let square = |closing_line: bool, corner: f64| {
            let mut commands = vec![
                PathCmd::MoveTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: corner, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
            ];
            if closing_line {
                commands.push(PathCmd::LineTo { x: 0.0, y: 0.0 });
            }
            commands.push(PathCmd::Close);
            let mut element = shape("custGeom");
            element.cust_geom = Some(vec![commands]);
            element
        };
        // MS-ODRAW stores the line back to the start point explicitly.
        let element = square(true, 1.0);
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &square(false, 1.0)),
            Verdict::Same
        );
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &square(false, 0.9)),
            Verdict::Differs("custom geometry vertices")
        );
        let mut unfilled = binary(&element, &leaf);
        unfilled.path_paint = Some(vec![(false, true)]);
        assert_eq!(
            same_geometry(&unfilled, &square(false, 1.0)),
            Verdict::Differs("custom geometry path paint")
        );
        let mut curved = square(false, 1.0);
        if let Some(paths) = curved.cust_geom.as_mut() {
            paths[0][1] = PathCmd::QuadBezTo {
                x1: 0.5,
                y1: 0.0,
                x: 1.0,
                y: 0.0,
            };
        }
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &curved),
            Verdict::Unverifiable("custom geometry")
        );
    }

    #[test]
    fn transforms_compare_in_the_binary_anchor_unit() {
        let master = |units: f64| (units * EMU_PER_MASTER_UNIT).round() as i64;
        let leaf = pptx_model::Transform {
            x: master(100.0),
            y: master(200.0),
            cx: master(300.0),
            cy: master(400.0),
            ..Default::default()
        };
        let mut alternative = shape("rect");
        alternative.x = master(101.0);
        alternative.y = master(200.0);
        alternative.width = master(300.0);
        alternative.height = master(399.5);
        assert_eq!(same_transform(&leaf, false, &alternative), Verdict::Same);
        alternative.x = master(101.0) + 2;
        assert_eq!(
            same_transform(&leaf, false, &alternative),
            Verdict::Differs("transform")
        );
        alternative.x = master(100.0);
        alternative.rotation = 90.0;
        assert_ne!(same_transform(&leaf, false, &alternative), Verdict::Same);
        // Angles compare modulo a full turn.
        let turned = pptx_model::Transform {
            rot: -90.0,
            ..leaf.clone()
        };
        alternative.rotation = 270.0;
        assert_eq!(same_transform(&turned, false, &alternative), Verdict::Same);
        // A group child's alternative states the unscaled child units: one
        // child unit apart is within rounding, two are not.
        let mut child = shape("rect");
        (child.x, child.y, child.width, child.height) = (101, 200, 300, 400);
        assert_eq!(same_transform(&leaf, true, &child), Verdict::Same);
        child.x = 102;
        assert_eq!(
            same_transform(&leaf, true, &child),
            Verdict::Differs("transform")
        );
        // The same child values read as EMU are far off.
        child.x = 100;
        assert_eq!(
            same_transform(&leaf, false, &child),
            Verdict::Differs("transform")
        );
    }

    #[test]
    fn fills_compare_the_recorded_paint() {
        let mut alternative = shape("rect");
        assert_eq!(
            same_fill(&RecordedFill::Unstated, &alternative, 90.0),
            Verdict::Same
        );
        alternative.fill = solid("FFFFFF");
        assert_eq!(
            same_fill(&RecordedFill::Unstated, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        assert_eq!(
            same_fill(&RecordedFill::NotDisplayed, &alternative, 90.0),
            Verdict::Same
        );
        assert_eq!(
            same_fill(&RecordedFill::Stated(None), &alternative, 90.0),
            Verdict::Differs("fill")
        );
        // A color transform resolved independently may round one unit away.
        alternative.fill = solid("808080");
        assert_eq!(
            same_fill(
                &RecordedFill::Stated(solid("7f7f7f").map(Box::new)),
                &alternative,
                90.0
            ),
            Verdict::Same
        );
        assert_eq!(
            same_fill(
                &RecordedFill::Stated(solid("7E7F7F").map(Box::new)),
                &alternative,
                90.0
            ),
            Verdict::Differs("solid fill color")
        );
        assert_eq!(
            same_fill(
                &RecordedFill::Stated(solid("80808080").map(Box::new)),
                &alternative,
                90.0
            ),
            Verdict::Differs("solid fill color")
        );
        let gradient = |stops: &[(f64, &str)], rotate: Option<bool>| {
            Some(Fill::Gradient {
                stops: stops
                    .iter()
                    .map(|&(position, color)| pptx_model::GradStop {
                        position,
                        color: color.to_owned(),
                    })
                    .collect(),
                angle: 90.0,
                grad_type: "linear".to_owned(),
                scaled: Some(false),
                path: None,
                fill_to_rect: None,
                tile_rect: None,
                flip: None,
                rot_with_shape: rotate,
            })
        };
        let recorded = RecordedFill::Stated(
            gradient(&[(0.0, "000000"), (0.3099975, "FFFFFF")], Some(true)).map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], Some(true));
        assert_eq!(same_fill(&recorded, &alternative, 90.0), Verdict::Same);
        // ECMA-376 CT_GradientFillProperties defaults flip to "none";
        // tileRect with all-zero edges covers the same whole shape as none.
        if let Some(Fill::Gradient {
            tile_rect, flip, ..
        }) = alternative.fill.as_mut()
        {
            *tile_rect = Some(ooxml_common::fill::FillRect::default());
            *flip = Some("none".into());
        }
        assert_eq!(same_fill(&recorded, &alternative, 90.0), Verdict::Same);
        let no_rotation = RecordedFill::Stated(
            gradient(&[(0.0, "000000"), (0.3099975, "FFFFFF")], Some(false)).map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], None);
        assert_eq!(same_fill(&no_rotation, &alternative, 0.0), Verdict::Same);
        assert_eq!(
            same_fill(&no_rotation, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        // Another stop layout or an unstated rotation may restate the same
        // paint: not comparable.
        alternative.fill = gradient(&[(0.0, "000000"), (0.31, "FFFFFF")], None);
        assert_eq!(
            same_fill(&recorded, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        assert_eq!(
            same_fill(&RecordedFill::Unknown, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
    }

    fn body(paragraphs: &[&[serde_json::Value]]) -> Option<pptx_model::TextBody> {
        Some(
            serde_json::from_value(serde_json::json!({
                "verticalAnchor": "t", "defaultFontSize": null,
                "defaultBold": null, "defaultItalic": null,
                "lIns": 0, "rIns": 0, "tIns": 0, "bIns": 0,
                "wrap": "square", "vert": "horz", "autoFit": "none",
                "paragraphs": paragraphs.iter().map(|runs| para(runs)).collect::<Vec<_>>()
            }))
            .unwrap(),
        )
    }

    fn texts(shape: &ShapeElement) -> Vec<Vec<String>> {
        shape
            .text_body
            .iter()
            .flat_map(|body| &body.paragraphs)
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| match run {
                        TextRun::Text(d) => d.text.clone(),
                        TextRun::Break => "<br>".to_owned(),
                        TextRun::Math { .. } => "<math>".to_owned(),
                    })
                    .collect()
            })
            .collect()
    }

    #[test]
    fn masked_text_takes_binary_characters_run_by_run() {
        let mut alternative = shape("rect");
        alternative.text_body = body(&[
            &[run("__"), serde_json::json!({"type": "break"}), run(" _ ")],
            &[run("____")],
        ]);
        let mut adopted = alternative.clone();
        assert_eq!(
            substitute_text(&mut adopted, Some("AB\u{b}C😀\rDEFG")),
            Ok(())
        );
        assert_eq!(texts(&adopted), [vec!["AB", "<br>", "C😀"], vec!["DEFG"]]);
        // Length, paragraph and break mismatches are different text.
        for text in [
            "AB\u{b}C😀\rDEF",
            "AB\u{b}C😀",
            "ABXC😀\rDEFG",
            "AB\u{b}C😀\u{b}\rDEFG",
        ] {
            assert_eq!(
                substitute_text(&mut alternative.clone(), Some(text)),
                Err(Verdict::Differs("text structure")),
                "{text:?}"
            );
        }
        // A run boundary inside a surrogate pair is rejected.
        let mut split = shape("rect");
        split.text_body = body(&[&[run(" "), run(" ")]]);
        assert!(substitute_text(&mut split, Some("😀")).is_err());
    }

    #[test]
    fn a_masked_vertical_tab_becomes_a_line_break_inside_its_run() {
        let mut alternative = shape("rect");
        alternative.text_body = body(&[&[run("__"), run("_"), run("___"), run("__")]]);
        let mut adopted = alternative.clone();
        assert_eq!(
            substitute_text(&mut adopted, Some("AB\u{b}\u{b}CDEF")),
            Ok(())
        );
        assert_eq!(texts(&adopted), [vec!["AB", "<br>", "<br>", "CD", "EF"]]);
        // Other break characters inside a run have no observed form.
        assert_eq!(
            substitute_text(&mut alternative.clone(), Some("AB\u{2028}\u{b}CDEF")),
            Err(Verdict::Unverifiable("line break"))
        );
    }

    #[test]
    fn run_formatting_must_agree_where_both_sides_state_it() {
        let sized = |size: serde_json::Value| {
            let mut r = run("ABCD");
            r["fontSize"] = size;
            body(&[&[r]])
        };
        let mut binary = shape("rect");
        let mut alternative = shape("rect");
        binary.text_body = sized(serde_json::json!(10.0));
        alternative.text_body = sized(serde_json::json!(10.0));
        assert_eq!(same_run_formatting(&binary, &alternative), Verdict::Same);
        // Whole-point binary sizes cannot state 10.5 pt: not the same shape.
        alternative.text_body = sized(serde_json::json!(10.5));
        assert_eq!(
            same_run_formatting(&binary, &alternative),
            Verdict::Differs("run formatting")
        );
        // An unstated size on either side is not a disagreement.
        binary.text_body = sized(serde_json::Value::Null);
        assert_eq!(same_run_formatting(&binary, &alternative), Verdict::Same);
        // A field compares as one unit whatever its displayed text.
        let mut field = run("12");
        field["fieldType"] = serde_json::json!("slidenum");
        field["fontSize"] = serde_json::json!(20.0);
        binary.text_body = body(&[&[run("A"), field.clone()]]);
        field["text"] = serde_json::json!("<#>");
        field["fontSize"] = serde_json::json!(18.0);
        alternative.text_body = body(&[&[run("A"), field]]);
        assert_eq!(
            same_run_formatting(&binary, &alternative),
            Verdict::Differs("run formatting")
        );
    }

    #[cfg(feature = "direct-ppt")]
    mod adoption {
        use super::*;
        use std::io::Write;

        const SHAPE_XML: &str = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="9" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1587500" cy="793750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>"#;

        fn package(parts: &[(&str, &str)], relationships: &[(&str, &str)]) -> Vec<u8> {
            package_named_rels("_rels/.rels", parts, relationships)
        }

        fn package_named_rels(
            rels_name: &str,
            parts: &[(&str, &str)],
            relationships: &[(&str, &str)],
        ) -> Vec<u8> {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            let rels: String = relationships
                .iter()
                .enumerate()
                .map(|(i, (kind, target))| {
                    format!(r#"<Relationship Id="rId{i}" Type="http://schemas.microsoft.com/office/2006/relationships/{kind}" Target="{target}"/>"#)
                })
                .collect();
            let rels = format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
            );
            for (name, body) in [(rels_name, rels.as_str())].iter().chain(parts) {
                writer.start_file(*name, options).unwrap();
                writer.write_all(body.as_bytes()).unwrap();
            }
            writer.finish().unwrap().into_inner()
        }

        fn blob(shape_xml: &str) -> Vec<u8> {
            package(
                &[("drs/shapexml.xml", shape_xml), ("drs/downrev.xml", "<x/>")],
                &[
                    ("downRev", "drs/downrev.xml"),
                    ("shapeXml", "drs/shapexml.xml"),
                ],
            )
        }

        #[test]
        fn alternative_relationship_uses_opc_resolution_and_equivalent_part_lookup() {
            let package = package_named_rels(
                "_RELS/%2Erels",
                &[("drs/ShapeXML.xml", SHAPE_XML)],
                &[("shapeXml", "DRS/%73hapeXML.xml#shape")],
            );
            assert_eq!(
                alternative_part(&package).unwrap(),
                Some(("shapeXml".into(), "DRS/shapeXML.xml".into()))
            );
            assert!(placeholder_locals(&package, "DRS/shapeXML.xml").is_ok());
        }

        fn theme() -> Theme {
            Theme::Readable {
                theme_xml: crate::ppt::theme(),
                clr_map: None,
            }
        }

        fn element(color: &str) -> ShapeElement {
            let mut element = shape("rect");
            (element.width, element.height) = (1587500, 793750);
            element.fill = solid(color);
            element.id = Some("3".to_owned());
            element
        }

        fn run_adopt(
            element: &ShapeElement,
            blob: &[u8],
            theme: &Theme,
            budgets: (usize, usize, usize),
        ) -> (Result<Option<ShapeElement>, String>, usize) {
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let shape = binary(element, &leaf);
            let (mut work, mut text, mut model) = budgets;
            let result = adopt(&shape, blob, theme, &mut work, &mut text, &mut model);
            (result, model)
        }

        const AMPLE: (usize, usize, usize) = (usize::MAX, usize::MAX, usize::MAX);

        #[test]
        fn a_consistent_alternative_is_adopted_and_charged() {
            let (result, model) = run_adopt(&element("FF0000"), &blob(SHAPE_XML), &theme(), AMPLE);
            let adopted = result.unwrap().expect("consistent alternative");
            assert_eq!(adopted.id.as_deref(), Some("3"));
            let charged = ooxml_common::json_measurement::measure_json(&adopted).unwrap();
            assert_eq!(usize::MAX - model, charged.json_bytes as usize);
            // A model budget below the adopted shape fails closed.
            let short = (usize::MAX, usize::MAX, charged.json_bytes as usize - 1);
            assert!(
                run_adopt(&element("FF0000"), &blob(SHAPE_XML), &theme(), short)
                    .0
                    .is_err()
            );
        }

        #[test]
        fn an_absent_or_disagreeing_alternative_keeps_the_binary() {
            let downrev_only = package(
                &[("drs/downrev.xml", "<x/>")],
                &[("downRev", "drs/downrev.xml")],
            );
            assert!(matches!(
                run_adopt(&element("FF0000"), &downrev_only, &theme(), AMPLE).0,
                Ok(None)
            ));
            assert!(matches!(
                run_adopt(&element("00FF00"), &blob(SHAPE_XML), &theme(), AMPLE).0,
                Ok(None)
            ));
        }

        #[test]
        fn placeholder_inherits_missing_shape_properties_from_the_binary() {
            let placeholder = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="title"/></p:nvPr>"#)
                .replace(
                    r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1587500" cy="793750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>"#,
                    "<p:spPr/>",
                );
            let adopted = run_adopt(&element("FF0000"), &blob(&placeholder), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("placeholder alternative");
            assert_eq!(adopted.geometry, "rect");
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
            assert_eq!((adopted.width, adopted.height), (1587500, 793750));
        }

        #[test]
        fn placeholder_uses_binary_transform_even_with_a_local_xml_transform() {
            let placeholder = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="body"/></p:nvPr>"#)
                .replace(r#"<a:off x="0" y="0"/>"#, r#"<a:off x="635000" y="0"/>"#);
            let adopted = run_adopt(&element("FF0000"), &blob(&placeholder), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("binary anchor wins");
            assert_eq!(adopted.x, 0);
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
        }

        #[test]
        fn disagreeing_local_placeholder_overrides_have_unknown_provenance() {
            let placeholder =
                SHAPE_XML.replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="body"/></p:nvPr>"#);
            let geometry = placeholder.replace("prst=\"rect\"", "prst=\"ellipse\"");
            let fill = placeholder.replace("val=\"FF0000\"", "val=\"00FF00\"");
            for (xml, attribute) in [
                (geometry.as_str(), "placeholder geometry precedence"),
                (fill.as_str(), "placeholder fill precedence"),
            ] {
                let error = run_adopt(&element("FF0000"), &blob(xml), &theme(), AMPLE)
                    .0
                    .unwrap_err();
                assert!(error.starts_with("UNSUPPORTED:") && error.contains(attribute));
            }
        }

        #[test]
        fn placeholder_keeps_alternative_text_formatting_with_binary_characters() {
            let xml = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="title"/></p:nvPr>"#)
                .replace(
                    "</p:sp>",
                    "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"4400\" b=\"1\"/><a:t>_____</a:t></a:r></a:p></p:txBody></p:sp>",
                );
            let mut element = element("FF0000");
            element.text_body = body(&[&[run("HELLO")]]);
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let mut binary = binary(&element, &leaf);
            binary.text = Some("HELLO");
            let (mut work, mut text, mut model) = AMPLE;
            let adopted = adopt(
                &binary,
                &blob(&xml),
                &theme(),
                &mut work,
                &mut text,
                &mut model,
            )
            .unwrap()
            .expect("placeholder alternative");
            let run = &adopted.text_body.unwrap().paragraphs[0].runs[0];
            assert!(
                matches!(run, TextRun::Text(data) if data.text == "HELLO" && data.font_size == Some(44.0) && data.bold == Some(true))
            );
        }

        #[test]
        fn freeform_binary_keeps_its_outline_with_a_preset_alternative() {
            let mut element = element("FF0000");
            element.geometry = "custGeom".into();
            element.cust_geom = Some(vec![vec![
                pptx_model::PathCmd::MoveTo { x: 0.0, y: 0.0 },
                pptx_model::PathCmd::LineTo { x: 1.0, y: 0.0 },
                pptx_model::PathCmd::LineTo { x: 1.0, y: 1.0 },
                pptx_model::PathCmd::Close,
            ]]);
            let adopted = run_adopt(&element, &blob(SHAPE_XML), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("alternative formatting with binary outline");
            assert_eq!(adopted.geometry, "custGeom");
            assert!(adopted.cust_geom.is_some());
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
        }

        #[test]
        fn unverifiable_alternatives_fail_closed() {
            let element = element("FF0000");
            let unsupported = |result: (Result<Option<ShapeElement>, String>, usize)| {
                let error = result.0.expect_err("must fail closed");
                assert!(error.starts_with("UNSUPPORTED:"), "{error}");
            };
            // Oversized, over budget, unreadable.
            unsupported(run_adopt(
                &element,
                &vec![0; MAX_BLOB_BYTES + 1],
                &theme(),
                AMPLE,
            ));
            unsupported(run_adopt(
                &element,
                &blob(SHAPE_XML),
                &theme(),
                (0, usize::MAX, usize::MAX),
            ));
            unsupported(run_adopt(
                &element,
                &blob(SHAPE_XML),
                &theme(),
                (usize::MAX, 10, usize::MAX),
            ));
            unsupported(run_adopt(&element, b"not a package", &theme(), AMPLE));
            unsupported(run_adopt(&element, &blob("<p:sp"), &theme(), AMPLE));
            unsupported(run_adopt(
                &element,
                &blob(SHAPE_XML),
                &Theme::Unreadable,
                AMPLE,
            ));
            // Non-shape alternatives.
            let group = package(
                &[("drs/groupshapexml.xml", "<x/>")],
                &[
                    ("downRev", "drs/downrev.xml"),
                    ("groupShapeXml", "drs/groupshapexml.xml"),
                ],
            );
            unsupported(run_adopt(&element, &group, &theme(), AMPLE));
            let two = package(
                &[("drs/shapexml.xml", SHAPE_XML)],
                &[
                    ("shapeXml", "drs/shapexml.xml"),
                    ("connectorXml", "drs/shapexml.xml"),
                ],
            );
            unsupported(run_adopt(&element, &two, &theme(), AMPLE));
            // A common attribute that cannot be compared.
            let gradient = SHAPE_XML.replace(
                r#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
                r#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>"#,
            );
            let mut patterned = element.clone();
            patterned.fill = Some(Fill::Pattern {
                fg: "FF0000".into(),
                bg: "0000FF".into(),
                preset: "pct5".into(),
            });
            unsupported(run_adopt(&patterned, &blob(&gradient), &theme(), AMPLE));
        }
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
