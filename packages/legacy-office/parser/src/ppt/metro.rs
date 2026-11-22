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
//!   not used once it disagrees with the binary;
//! - controls whose XML alone changes a gradient stop or states
//!   `rotWithShape="1"` render the XML, while a control whose binary alone
//!   turns the gradient's rotation flag on renders the binary; a control
//!   whose binary freeform alone moves its start vertex along the first edge
//!   (the traced outline is unchanged) prints like its source.
//!
//! The alternative therefore carries display information the binary lacks,
//! and the binary projection alone does not reproduce a shape whose
//! alternative PowerPoint uses. Each shape resolves to one of three
//! outcomes (`adopt`):
//! - no alternative part (a package with only the `downRev` checksums), or
//!   an ordinary alternative that verifiably disagrees with the binary on a
//!   compared attribute, or whose binary overrides it (`Verdict::Overridden`:
//!   a binary gradient rotation flag the XML does not state as `1`, a stroked
//!   custom path whose outline matches only by folding a line that turns
//!   back on itself): the binary projection;
//! - an alternative that agrees on every applicable compared attribute: the
//!   alternative, with the binary's transform, identifier and characters;
//!   a binary freeform keeps its outline, and a placeholder inherits locally
//!   omitted shape properties from the binary projection;
//! - an alternative whose only disagreements are restatements (`Verdict::
//!   Restated`: gradient stops, an XML gradient rotation flag `1` over a
//!   binary that states none): the alternative, as above. This precedence
//!   is inferred from the measured XML-only edit controls; no control
//!   measures a binary-only edit of these values that the comparison reads;
//! - alternative resource exhaustion keeps the binary projection;
//! - encryption failures keep their own classification;
//! - representational failures reject as unsupported: an ambiguous or
//!   unreadable package or theme, a part that is not a shape or
//!   connector, unverified relationship references, and a compared attribute that the
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
const MAX_PART_BYTES: u64 = 1024 * 1024;
#[cfg(feature = "direct-ppt")]
const MAX_PACKAGE_BYTES: u64 = 4 * 1024 * 1024;
/// Optional verification gets the existing standard OPC package byte allowance
/// for actual decoder input (including legacy projection rereads) and index
/// work. The ordinary OOXML policy has no input/output ratio limit, so use its
/// package ceiling rather than inventing one or reusing a binary record count.
/// Calibration: all 34 local PPTs (357 slides) completed without verification
/// fallback; maxima were 48,012 consumed input bytes and 48,049 total
/// work units per alternative. These measurements validate headroom, not the
/// policy's derivation. Preflight and legacy rereads are metered; modern parser
/// replay is bounded by the same immutable preflighted package.
/// This does not admit/reject by blob length or compressed-size declarations.
const MAX_INPUT_WORK: usize = ooxml_common::resource::STANDARD_MAX_TOTAL_INFLATED_BYTES as usize;
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
        /// A master theme is shared by many placeholders; resolve its style
        /// matrix once when checking explicit line/effect references.
        format_scheme: std::cell::OnceCell<ooxml_common::theme::ThemeFormatScheme>,
    },
    /// Malformed or ambiguous: no alternative XML on this
    /// master's slides can be resolved.
    Unreadable,
    /// Optional theme verification ran out of resources. Shapes using it keep
    /// their binary projections; this is not a malformed-theme rejection.
    BudgetLimited,
}

impl Theme {
    #[cfg(feature = "direct-ppt")]
    fn format_scheme(&self) -> Option<&ooxml_common::theme::ThemeFormatScheme> {
        let Theme::Readable {
            theme_xml,
            format_scheme,
            ..
        } = self
        else {
            return None;
        };
        Some(format_scheme.get_or_init(|| ooxml_common::theme::ThemeFormatScheme::parse(theme_xml)))
    }
    /// ECMA-376 §20.1.6.2 supplies hlink/folHlink in the theme color scheme.
    /// PowerPoint-saved PPTX controls retain those slots for legacy slides;
    /// an inherited binary text color must not hide them on linked text.
    pub(in crate::ppt) fn hyperlink_colors(&self) -> Option<(String, String)> {
        let Theme::Readable { theme_xml, .. } = self else {
            return None;
        };
        let colors = ooxml_common::theme::ThemeColorScheme::parse(theme_xml);
        let color = |slot| {
            colors
                .get(slot)
                .filter(|value| value.len() == 6 && value.bytes().all(|b| b.is_ascii_hexdigit()))
                .map(str::to_owned)
        };
        Some((color("hlink")?, color("folHlink")?))
    }
}

/// Read the round-trip theme of a main master's child records; `None` when
/// the master has none.
pub(in crate::ppt) fn master_theme(
    records: &[Record<'_>],
    decoded_budget: &mut usize,
) -> Result<Option<Theme>, String> {
    // Round-trip themes serve optional alternative verification. Input work
    // cannot consume the binary record ledger. Retained theme output is still
    // shared across masters, but a failed attempt commits none of its charges.
    let mut decoded = *decoded_budget;
    let mut work = MAX_INPUT_WORK;
    match read_master_theme(records, &mut decoded, &mut work) {
        Ok(theme) => {
            *decoded_budget = decoded;
            Ok(theme)
        }
        Err(error) if resource_rejection(&error) => Ok(Some(Theme::BudgetLimited)),
        Err(error) => Err(error),
    }
}

fn read_master_theme(
    records: &[Record<'_>],
    decoded_budget: &mut usize,
    work_budget: &mut usize,
) -> Result<Option<Theme>, String> {
    let mut themes = records.iter().filter(|r| r.kind == 0x040e);
    let Some(theme) = themes.next() else {
        return Ok(None);
    };
    let unreadable = || unsupported("unreadable PowerPoint round-trip theme");
    let mut readable = || -> Result<Theme, String> {
        if themes.next().is_some() {
            return Err(unreadable());
        }
        let mut maps = records.iter().filter(|r| r.kind == 0x040f);
        let clr_map = match (maps.next(), maps.next()) {
            (Some(map), None) => {
                super::charge_text(decoded_budget, map.payload.len())?;
                Some(
                    std::str::from_utf8(map.payload)
                        .map_err(|_| unreadable())?
                        .to_owned(),
                )
            }
            (None, None) => None,
            _ => return Err(unreadable()),
        };
        Ok(Theme::Readable {
            theme_xml: theme_part(theme.payload, decoded_budget, work_budget)?,
            clr_map,
            format_scheme: std::cell::OnceCell::new(),
        })
    };
    // Keep the measured failure distinct for the owner: optional resource
    // exhaustion becomes BudgetLimited, while encryption stays fatal. Ordinary
    // unreadability still reaches the shape's existing theme agreement gate.
    match readable() {
        Ok(theme) => Ok(Some(theme)),
        Err(error) if error.starts_with("UNSUPPORTED:") => Ok(Some(Theme::Unreadable)),
        Err(error) => Err(error),
    }
}

/// Resolve the OPC name before opening a raw, work-metered entry. Encryption
/// retains its classification; output and input work are bounded during reads.
fn open_part<'a, R: std::io::Read + std::io::Seek>(
    archive: &'a mut zip::ZipArchive<R>,
    name: &str,
) -> Result<crate::opc_part::Entry<'a, R>, String> {
    let index = crate::opc_part::entry_index(archive, name)
        .ok_or_else(|| unsupported("unreadable PowerPoint embedded part"))?;
    crate::opc_part::Entry::open(archive, index)
}

fn theme_part(
    package: &[u8],
    decoded_budget: &mut usize,
    work_budget: &mut usize,
) -> Result<String, String> {
    let unreadable = || unsupported("unreadable PowerPoint round-trip theme");
    // The budgeted footer/index preflight opens at the validated archive
    // offset, so the ZIP reader never scans the payload for signatures.
    let mut archive = zip_index::open(package, work_budget).map_err(|error| {
        if error.starts_with("UNSUPPORTED:") {
            unreadable()
        } else {
            error
        }
    })?;
    if archive.len() > MAX_THEME_ENTRIES {
        return Err(crate::rejection::resource(
            "PowerPoint theme ZIP entry budget exceeded",
        ));
    }
    // Even an entry not selected by the relationships must retain its
    // encryption classification, matching the alternative-shape preflight.
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index).map_err(|_| unreadable())?;
        if entry.encrypted() {
            return Err(crate::rejection::encryption(
                "encrypted PowerPoint theme ZIP entry",
            ));
        }
    }
    let mut read = |name: &str| -> Result<String, String> {
        let entry = open_part(&mut archive, name)?;
        // Charge actual output across all masters, never ZIP declarations.
        let mut entry = entry;
        let mut bytes = Vec::new();
        let mut chunk = [0; 8192];
        loop {
            let allowance = (MAX_THEME_BYTES as usize - bytes.len())
                .min(*decoded_budget)
                .saturating_add(1)
                .min(chunk.len());
            let count = entry.read(&mut chunk[..allowance], work_budget)?;
            if count == 0 {
                break;
            }
            super::charge_text(decoded_budget, count)?;
            if bytes.len() + count > MAX_THEME_BYTES as usize {
                return Err(crate::rejection::resource(
                    "PowerPoint theme part byte budget exceeded",
                ));
            }
            bytes
                .try_reserve(count)
                .map_err(|_| crate::rejection::resource("PowerPoint theme allocation failed"))?;
            bytes.extend_from_slice(&chunk[..count]);
        }
        // Header consistency is integrity validation after measured work.
        if bytes.len() as u64 != entry.size() {
            return Err(unreadable());
        }
        let text = String::from_utf8(bytes).map_err(|_| unreadable())?;
        Ok(text)
    };
    // Follow the package relationships: root -> theme manager -> theme.
    let target = |rels: &str, source: &str, kind: &str| -> Result<String, String> {
        let doc = parse_xml(rels)?;
        let mut found = doc.root_element().children().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type").is_some_and(|t| t.ends_with(kind))
        });
        let first = found.next().ok_or_else(unreadable)?;
        if first
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
            || found.next().is_some()
        {
            return Err(unreadable());
        }
        ooxml_common::rels::resolve_part_name(
            source,
            first.attribute("Target").ok_or_else(unreadable)?,
        )
        .ok_or_else(unreadable)
    };
    let manager = target(&read("_rels/.rels")?, "", "/officeDocument")?;
    let (dir, name) = manager.rsplit_once('/').unwrap_or(("", &manager));
    let rels = if dir.is_empty() {
        format!("_rels/{name}.rels")
    } else {
        format!("{dir}/_rels/{name}.rels")
    };
    let theme = target(&read(&rels)?, &manager, "/theme")?;
    let xml = read(&theme)?;
    let doc = parse_xml(&xml)?;
    let root = doc.root_element();
    if root.tag_name().name() != "theme" || !ooxml_common::ns::is_a_ns(root.tag_name().namespace())
    {
        return Err(unreadable());
    }
    Ok(xml)
}

fn parse_xml(xml: &str) -> Result<roxmltree::Document<'_>, String> {
    ooxml_common::depth::parse_guarded(xml).map_err(|error| match error {
        ooxml_common::depth::GuardedParseError::TooDeep => {
            crate::rejection::resource("PowerPoint embedded XML depth budget exceeded")
        }
        _ => unsupported("unreadable PowerPoint embedded XML"),
    })
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
    /// Direct CFStyle font-size bits, before master inheritance. Only used
    /// when an XML picture bullet needs inherited-size normalization.
    pub direct_size_authored: Option<&'a [bool]>,
    pub fill: RecordedFill,
    /// Authored per-path (fill, stroke) flags of custom geometry, before
    /// PowerPoint's open-path display rule (`officeart::geometry`).
    pub path_paint: Option<Vec<(bool, bool)>>,
    /// The range of each converted preset adjust value under the rounding
    /// of the binary anchor (`officeart::preset::adjustment_bounds`).
    pub adjust_bounds: [Option<(f64, f64)>; 8],
    /// Already admitted OfficeArt image bytes for a picture fill. The image
    /// path in `element.fill` identifies the same retained resource.
    pub image_bytes: Option<Vec<u8>>,
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
    /// Both forms state the attribute in encodings that do not correspond
    /// one to one, in a pairing whose PowerPoint 16 PDF controls draw the
    /// alternative (gradient stops, and an XML rotation flag `1` over a
    /// binary that states none; `same_paint`). The value names the rejection
    /// this outcome replaces.
    Restated(&'static str),
    /// Both forms state the attribute in encodings that do not correspond
    /// one to one, in a pairing whose PowerPoint 16 PDF controls draw the
    /// binary (a binary gradient rotation flag that the XML does not state
    /// as `1`, and a concentric shade whose XML adds a tile rectangle or
    /// restates a stop; `same_paint`), or draw both forms alike (a stroked
    /// custom path whose outline matches only by folding a line that turns
    /// back on itself; `same_outline`), or in a pairing this reader cannot
    /// compare while the binary form is complete by itself (a pattern
    /// bitmap against a preset whose cells were not measured;
    /// `same_binary_pattern`). The shape keeps its binary projection, as for
    /// [`Verdict::Differs`]; the value names the rejection this outcome
    /// replaces.
    Overridden(&'static str),
}

#[cfg(any(test, feature = "direct-ppt"))]
const RESTATED_FILL: &str =
    "PowerPoint alternative shape XML fill cannot be compared with the binary shape";
#[cfg(any(test, feature = "direct-ppt"))]
const RESTATED_GEOMETRY: &str =
    "PowerPoint alternative shape XML custom geometry cannot be compared with the binary shape";

/// How a shape resolves before its alternative is projected.
#[cfg(any(test, feature = "direct-ppt"))]
#[derive(Debug, PartialEq)]
enum Decision {
    /// The binary projection.
    Binary,
    /// The alternative; once it is adopted, the restated attributes it draws
    /// are recorded with these replaced rejections.
    Alternative(Vec<&'static str>),
}

/// Any verified difference keeps the binary projection: PowerPoint ignores an
/// alternative that disagrees with the binary shape. Without one, an
/// overriding binary encoding keeps it too, and is recorded as the decision
/// that replaced its rejection. Otherwise, when some attribute cannot be
/// compared, the display is unknown and the shape fails closed. Otherwise the
/// alternative is adopted, with its restated attributes.
#[cfg(any(test, feature = "direct-ppt"))]
fn decide(verdicts: &[Verdict]) -> Result<Decision, String> {
    if verdicts.iter().any(|v| matches!(v, Verdict::Differs(_))) {
        return Ok(Decision::Binary);
    }
    let overridden: Vec<_> = verdicts
        .iter()
        .filter_map(|v| match v {
            Verdict::Overridden(site) => Some(*site),
            _ => None,
        })
        .collect();
    if !overridden.is_empty() {
        for site in overridden {
            crate::rejection::approximated(site, ());
        }
        return Ok(Decision::Binary);
    }
    if let Some(what) = verdicts.iter().find_map(|v| match v {
        Verdict::Unverifiable(what) => Some(what),
        _ => None,
    }) {
        return Err(unverifiable(what));
    }
    Ok(Decision::Alternative(
        verdicts
            .iter()
            .filter_map(|v| match v {
                Verdict::Restated(site) => Some(*site),
                _ => None,
            })
            .collect(),
    ))
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
fn alternative_part(
    blob: &[u8],
    work_budget: &mut usize,
) -> Result<Option<(String, String)>, String> {
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let mut archive = zip_index::open(blob, work_budget)?;
    let rels =
        open_part(&mut archive, "_rels/.rels")?.read_to_string(work_budget, MAX_PART_BYTES)?;
    let doc = parse_xml(&rels)?;
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

/// Check the OPC references used by the alternative without fetching an
/// external target or treating a resolved target as evidence that the binary
/// and XML agree. ECMA-376 Part 2 §6.5.2.3 locates the relationship part
/// beside its source part; §6.5.3.4 keeps external targets outside the ZIP.
/// PowerPoint precedence for these references still needs its control PDFs.
#[cfg(feature = "direct-ppt")]
fn read_blob_reference_part<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
    work_budget: &mut usize,
) -> Result<String, String> {
    let text = open_part(archive, name)?.read_to_string(work_budget, MAX_PART_BYTES)?;
    Ok(text)
}

#[cfg(feature = "direct-ppt")]
fn validate_blob_references(
    blob: &[u8],
    part: &str,
    work_budget: &mut usize,
) -> Result<(), String> {
    use std::collections::BTreeSet;

    let unreadable = || unverifiable("relationship");
    let mut archive = zip_index::open(blob, work_budget)?;
    // Resource policy: avoid a reference-by-entry quadratic scan of a ZIP
    // with a huge number of tiny entries. Ordinary Office shape blobs are
    // much smaller; this limit makes the passive check bounded.
    if archive.len() > 128 {
        return Err(crate::rejection::resource(
            "PowerPoint alternative ZIP entry budget exceeded",
        ));
    }
    let xml = read_blob_reference_part(&mut archive, part, work_budget)?;
    let doc = parse_xml(&xml)?;
    let ids: BTreeSet<&str> = doc
        .descendants()
        .flat_map(|node| node.attributes())
        .filter(|attribute| ooxml_common::ns::is_r_ns(attribute.namespace()))
        .map(|attribute| attribute.value())
        // An empty identifier names no relationship: PowerPoint writes
        // `r:id=""` on a hyperlink that only states an action verb
        // (ECMA-376 21.1.2.3.5).
        .filter(|id| !id.is_empty())
        .collect();
    if ids.is_empty() {
        // The selected theme style may own the reference instead. It is not
        // in this blob and remains unverifiable at the adoption gate.
        return Ok(());
    }
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path, work_budget)?;
    let rels = parse_xml(&rels)?;
    let root = rels.root_element();
    if root.tag_name().name() != "Relationships"
        || root.tag_name().namespace()
            != Some("http://schemas.openxmlformats.org/package/2006/relationships")
    {
        return Err(unreadable());
    }
    let mut seen = BTreeSet::new();
    let mut matched = BTreeSet::new();
    for relationship in root.children().filter(|node| node.is_element()) {
        if relationship.tag_name().name() != "Relationship"
            || relationship.tag_name().namespace() != root.tag_name().namespace()
        {
            return Err(unreadable());
        }
        let id = relationship.attribute("Id").ok_or_else(unreadable)?;
        if !seen.insert(id) {
            return Err(unreadable());
        }
        if !ids.contains(id) {
            continue;
        }
        let target = relationship.attribute("Target").ok_or_else(unreadable)?;
        if target.is_empty() || relationship.attribute("Type").is_none_or(str::is_empty) {
            return Err(unreadable());
        }
        match relationship.attribute("TargetMode") {
            None | Some("Internal") => {
                let name =
                    ooxml_common::rels::resolve_part_name(part, target).ok_or_else(unreadable)?;
                if crate::opc_part::entry_index(&archive, &name).is_none() {
                    return Err(unreadable());
                }
            }
            Some("External") => {} // Opaque URI; never read or fetched.
            _ => return Err(unreadable()),
        }
        matched.insert(id);
    }
    if matched != ids {
        return Err(unreadable());
    }
    Ok(())
}

/// Read one image that is the only reference in the alternative shape part.
/// The ZIP and part size limits were checked at the adoption boundary. A
/// non-image or second reference is left to the typed relationship gate.
#[cfg(feature = "direct-ppt")]
fn sole_blob_image(
    blob: &[u8],
    part: &str,
    image_path: &str,
    work_budget: &mut usize,
) -> Result<Vec<u8>, String> {
    let unreadable = || unverifiable("relationship");
    let mut archive = zip_index::open(blob, work_budget)?;
    let xml = read_blob_reference_part(&mut archive, part, work_budget)?;
    let doc = parse_xml(&xml)?;
    let refs: Vec<_> = doc
        .descendants()
        .flat_map(|node| node.attributes().map(move |attribute| (node, attribute)))
        .filter(|(_, attribute)| ooxml_common::ns::is_r_ns(attribute.namespace()))
        .collect();
    if refs.len() != 1 || refs[0].0.tag_name().name() != "blip" || refs[0].1.name() != "embed" {
        return Err(unreadable());
    }
    let id = refs[0].1.value();
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path, work_budget)?;
    let rels = parse_xml(&rels)?;
    let mut matching = rels.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().name() == "Relationship"
            && node.attribute("Id") == Some(id)
    });
    let relationship = matching.next().ok_or_else(unreadable)?;
    if matching.next().is_some()
        || !relationship
            .attribute("Type")
            .is_some_and(|kind| kind.ends_with("/image"))
        || relationship
            .attribute("TargetMode")
            .is_some_and(|mode| mode != "Internal")
    {
        return Err(unreadable());
    }
    let target = ooxml_common::rels::resolve_part_name(
        part,
        relationship.attribute("Target").ok_or_else(unreadable)?,
    )
    .ok_or_else(unreadable)?;
    if target != image_path {
        return Err(unreadable());
    }
    let bytes = open_part(&mut archive, &target)?.read_to_end(work_budget, MAX_PART_BYTES)?;
    Ok(bytes)
}

/// A shape whose only OPC dependencies are external click or mouse-over
/// URLs can be compared to [MS-PPT] InteractiveInfoAtom/ExHyperlinkContainer
/// targets. A hyperlink with an empty identifier depends on no relationship.
/// Other relationship classes still have no binary counterpart here.
#[cfg(feature = "direct-ppt")]
fn blob_hyperlinks_only(blob: &[u8], part: &str, work_budget: &mut usize) -> Result<bool, String> {
    use std::collections::BTreeSet;
    let mut archive = zip_index::open(blob, work_budget)?;
    let xml = read_blob_reference_part(&mut archive, part, work_budget)?;
    let doc = parse_xml(&xml)?;
    let mut ids = BTreeSet::new();
    let mut hyperlinks = false;
    for node in doc.descendants() {
        for attr in node
            .attributes()
            .filter(|a| ooxml_common::ns::is_r_ns(a.namespace()))
        {
            if !matches!(node.tag_name().name(), "hlinkClick" | "hlinkMouseOver")
                || attr.name() != "id"
            {
                return Ok(false);
            }
            hyperlinks = true;
            if !attr.value().is_empty() {
                ids.insert(attr.value().to_owned());
            }
        }
    }
    if ids.is_empty() {
        return Ok(hyperlinks);
    }
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path, work_budget)?;
    let rels = parse_xml(&rels)?;
    Ok(ids.iter().all(|id| {
        rels.descendants()
            .filter(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Id") == Some(id)
            })
            .count()
            == 1
            && rels.descendants().any(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Id") == Some(id)
                    && n.attribute("Type")
                        .is_some_and(|t| t.ends_with("/hyperlink"))
                    && n.attribute("TargetMode") == Some("External")
            })
    }))
}

/// Return only image references used by DrawingML picture bullets. The
/// presentation model already has a generic buBlip marker; this extracts its
/// passive OPC targets for the legacy resource provider. Unrelated links are
/// left at the typed gate.
#[cfg(any(test, feature = "direct-ppt"))]
type BlobBulletAsset = (String, &'static str, Vec<u8>);

#[cfg(feature = "direct-ppt")]
fn blob_bullet_images(
    blob: &[u8],
    part: &str,
    work_budget: &mut usize,
) -> Result<Option<Vec<BlobBulletAsset>>, String> {
    use std::collections::BTreeMap;
    let unreadable = || unverifiable("relationship");
    let mut archive = zip_index::open(blob, work_budget)?;
    let xml = read_blob_reference_part(&mut archive, part, work_budget)?;
    let doc = parse_xml(&xml)?;
    let mut ids = Vec::new();
    for node in doc.descendants() {
        for attr in node
            .attributes()
            .filter(|a| ooxml_common::ns::is_r_ns(a.namespace()))
        {
            if node.tag_name().name() != "blip"
                || node
                    .parent()
                    .is_none_or(|p| p.tag_name().name() != "buBlip")
                || attr.name() != "embed"
            {
                return Ok(None);
            }
            ids.push(attr.value().to_owned());
        }
    }
    if ids.is_empty() {
        return Ok(None);
    }
    let rels_path = ooxml_common::rels::relationship_part_path(part);
    let rels = read_blob_reference_part(&mut archive, &rels_path, work_budget)?;
    let rels = parse_xml(&rels)?;
    let mut images = BTreeMap::new();
    for id in ids {
        let mut matches = rels.descendants().filter(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Id") == Some(id.as_str())
        });
        let relation = matches.next().ok_or_else(unreadable)?;
        if matches.next().is_some()
            || !relation
                .attribute("Type")
                .is_some_and(|t| t.ends_with("/image"))
            || relation
                .attribute("TargetMode")
                .is_some_and(|m| m != "Internal")
        {
            return Err(unreadable());
        }
        let path = ooxml_common::rels::resolve_part_name(
            part,
            relation.attribute("Target").ok_or_else(unreadable)?,
        )
        .ok_or_else(unreadable)?;
        if images.contains_key(&path) {
            continue;
        }
        let extension = match path
            .rsplit_once('.')
            .map(|(_, ext)| ext.to_ascii_lowercase())
            .as_deref()
        {
            Some("gif") => "gif",
            Some("png") => "png",
            Some("jpg" | "jpeg") => "jpg",
            _ => return Err(unreadable()),
        };
        let bytes = open_part(&mut archive, &path)?.read_to_end(work_budget, MAX_PART_BYTES)?;
        let signature = match extension {
            "gif" => bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a"),
            "png" => bytes.starts_with(b"\x89PNG\r\n\x1a\n"),
            _ => bytes.starts_with(&[0xff, 0xd8, 0xff]),
        };
        if !signature {
            return Err(unreadable());
        }
        images.insert(path, (extension, bytes));
    }
    Ok(Some(
        images
            .into_iter()
            .map(|(path, (ext, bytes))| (path, ext, bytes))
            .collect(),
    ))
}

/// OfficeArt msofillPattern stores an 8x8 GIF, or a 10x10 one of which only
/// the upper-left 8x8 cells tile. PowerPoint 16 controls of pct30 and
/// ltUpDiag show that white pixels map to the XML foreground and black to
/// the background. Decode only a single bounded GIF image of either size.
#[cfg(any(test, feature = "direct-ppt"))]
fn pattern_image(bytes: &[u8]) -> Option<PatternImage> {
    if bytes.len() < 26 || !matches!(&bytes[..6], b"GIF87a" | b"GIF89a") {
        return None;
    }
    let word = |at: usize| -> Option<u16> {
        Some(u16::from_le_bytes(bytes.get(at..at + 2)?.try_into().ok()?))
    };
    let size = word(6)?;
    if !matches!(size, 8 | 10) || word(8)? != size {
        return None;
    }
    let mut at = 13usize;
    let table_len = |flag: u8| 3usize.checked_mul(1usize << (usize::from(flag & 7) + 1));
    let mut palette = if bytes[10] & 0x80 != 0 {
        let len = table_len(bytes[10])?;
        let table = bytes.get(at..at + len)?.to_vec();
        at += len;
        table
    } else {
        return None;
    };
    let mut transparent = None;
    let mut pixels = None;
    while at < bytes.len() {
        let kind = *bytes.get(at)?;
        at += 1;
        match kind {
            0x21 => {
                let label = *bytes.get(at)?;
                at += 1;
                if label == 0xf9 {
                    if *bytes.get(at)? != 4 {
                        return None;
                    }
                    let packed = *bytes.get(at + 1)?;
                    transparent = (packed & 1 != 0).then_some(*bytes.get(at + 4)?);
                }
                loop {
                    let size = usize::from(*bytes.get(at)?);
                    at += 1;
                    if size == 0 {
                        break;
                    }
                    at = at.checked_add(size)?;
                    bytes.get(..at)?;
                }
            }
            0x2c if pixels.is_none() => {
                if word(at)? != 0
                    || word(at + 2)? != 0
                    || word(at + 4)? != size
                    || word(at + 6)? != size
                {
                    return None;
                }
                let flags = *bytes.get(at + 8)?;
                if flags & 0x40 != 0 {
                    return None; // interlaced cells need a different row order
                }
                at += 9;
                if flags & 0x80 != 0 {
                    let len = table_len(flags)?;
                    palette = bytes.get(at..at + len)?.to_vec();
                    at += len;
                }
                let min_code = *bytes.get(at)?;
                at += 1;
                if !(2..=8).contains(&min_code) {
                    return None;
                }
                let mut data = Vec::new();
                loop {
                    let size = usize::from(*bytes.get(at)?);
                    at += 1;
                    if size == 0 {
                        break;
                    }
                    data.extend_from_slice(bytes.get(at..at + size)?);
                    at += size;
                }
                pixels = Some((data, min_code));
            }
            0x3b => break,
            _ => return None,
        }
    }
    let (data, min_code) = pixels?;
    Some(PatternImage {
        size: usize::from(size),
        data,
        min_code,
        palette,
        transparent,
    })
}

#[cfg(any(test, feature = "direct-ppt"))]
struct PatternImage {
    /// Width and height in pixels: 8 or 10.
    size: usize,
    data: Vec<u8>,
    min_code: u8,
    palette: Vec<u8>,
    transparent: Option<u8>,
}

#[cfg(any(test, feature = "direct-ppt"))]
fn binary_pattern_cells(bytes: &[u8]) -> Result<Option<[u8; 8]>, String> {
    // The comparison with a DrawingML preset was measured on 10x10 bitmaps.
    Ok(gif_pattern_cells(bytes)?
        .filter(|(size, _)| *size == 10)
        .map(|(_, cells)| cells))
}

/// The size (8 or 10) and the upper-left 8x8 cells of a black-and-white GIF
/// pattern bitmap: one byte per row, a set bit for a white pixel.
#[cfg(any(test, feature = "direct-ppt"))]
pub(in crate::ppt) fn gif_pattern_cells(bytes: &[u8]) -> Result<Option<(usize, [u8; 8])>, String> {
    // Decoder admission policy, independent of the 10x10 pattern subset.
    if bytes.len() > 4096 {
        return Err(crate::rejection::resource(
            "PowerPoint pattern GIF byte budget exceeded",
        ));
    }
    let Some(PatternImage {
        size,
        data,
        min_code,
        palette,
        transparent,
    }) = pattern_image(bytes)
    else {
        return Ok(None);
    };
    let Some(pixels) = decode_pattern_lzw(&data, min_code, size * size)? else {
        return Ok(None);
    };
    Ok((|| {
        let mut rows = [0u8; 8];
        for y in 0..8 {
            for x in 0..8 {
                let index = pixels[y * size + x];
                if transparent == Some(index) {
                    return None;
                }
                let color = palette.get(usize::from(index) * 3..usize::from(index) * 3 + 3)?;
                match color {
                    [255, 255, 255] => rows[y] |= 0x80 >> x,
                    [0, 0, 0] => {}
                    _ => return None,
                }
            }
        }
        Some((size, rows))
    })())
}

/// GIF LZW streams are little-endian bit packed. The admission bound of
/// `pixels` (at most 10x10) caps both output and dictionary work,
/// independent of compressed input.
#[cfg(any(test, feature = "direct-ppt"))]
fn decode_pattern_lzw(data: &[u8], min_code: u8, pixels: usize) -> Result<Option<Vec<u8>>, String> {
    (|| -> Option<Result<Vec<u8>, String>> {
        let clear = 1usize << min_code;
        let end = clear + 1;
        let initial = || {
            let mut table: Vec<Vec<u8>> = (0..clear).map(|value| vec![value as u8]).collect();
            table.extend([Vec::new(), Vec::new()]);
            table
        };
        let mut table = initial();
        let mut width = usize::from(min_code) + 1;
        let mut bit = 0usize;
        let mut previous: Option<Vec<u8>> = None;
        let mut output = Vec::with_capacity(pixels);
        for _ in 0..512 {
            if bit + width > data.len() * 8 {
                return None;
            }
            let mut code = 0usize;
            for shift in 0..width {
                let position = bit + shift;
                code |= usize::from((data[position / 8] >> (position % 8)) & 1) << shift;
            }
            bit += width;
            if code == clear {
                table = initial();
                width = usize::from(min_code) + 1;
                previous = None;
                continue;
            }
            if code == end {
                return (output.len() == pixels).then_some(Ok(output));
            }
            let entry = if code < table.len() && !table[code].is_empty() {
                table[code].clone()
            } else if code == table.len() {
                let mut entry = previous.clone()?;
                entry.push(*entry.first()?);
                entry
            } else {
                return None;
            };
            if output.len().checked_add(entry.len())? > pixels {
                return None;
            }
            output.extend_from_slice(&entry);
            if let Some(mut prev) = previous {
                if table.len() < 4096 {
                    prev.push(*entry.first()?);
                    table.push(prev);
                    if table.len() == 1 << width && width < 12 {
                        width += 1;
                    }
                }
            }
            previous = Some(entry);
        }
        Some(Err(crate::rejection::resource(
            "PowerPoint pattern LZW work budget exceeded",
        )))
    })()
    .transpose()
}

#[cfg(any(test, feature = "direct-ppt"))]
fn same_binary_pattern(
    binary: &BinaryShape<'_>,
    alternative: &ShapeElement,
) -> Result<Option<Verdict>, String> {
    use pptx_model::Fill;
    let (
        Some(Fill::Image {
            duotone: Some(colors),
            ..
        }),
        Some(Fill::Pattern { fg, bg, preset }),
    ) = (&binary.element.fill, &alternative.fill)
    else {
        return Ok(None);
    };
    // A preset whose cells were not measured, or a bitmap this reader does
    // not decode, cannot be compared: the binary pattern is drawn as stored.
    let expected = match preset.as_str() {
        "pct30" => [0xaa, 0x44, 0xaa, 0x11, 0xaa, 0x44, 0xaa, 0x11],
        "ltUpDiag" => [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
        _ => return Ok(Some(Verdict::Overridden(RESTATED_FILL))),
    };
    let cells = binary
        .image_bytes
        .as_deref()
        .map(binary_pattern_cells)
        .transpose()?
        .flatten();
    let Some(cells) = cells else {
        return Ok(Some(Verdict::Overridden(RESTATED_FILL)));
    };
    // MS-ODRAW binary pattern BLIP: black is clr1/background; white is
    // clr2/foreground. The XML pattern preset names the same 8x8 cells.
    Ok(Some(
        if cells == expected
            && same_color(&colors.clr1, bg) == Some(true)
            && same_color(&colors.clr2, fg) == Some(true)
        {
            Verdict::Same
        } else {
            Verdict::Differs("pattern fill")
        },
    ))
}

/// MS-ODRAW §2.3.4.41 permits ignoring this optional XML alternative. Library
/// policy: a verification quota rejects only this alternative, keeping the
/// already projected binary shape. Scratch inflation/text/model and input work
/// belong to one attempt, never to the binary record/text session ledgers.
/// Commit adopted-model charges only on success; rejection must leave enough
/// session budget to retain the binary content. Encryption and representational
/// failures retain their existing classifications.
#[cfg(feature = "direct-ppt")]
pub(in crate::ppt) fn adopt(
    binary: &BinaryShape<'_>,
    blob: &[u8],
    theme: &Theme,
    media: &mut super::media::SpanStore,
    model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    adopt_with_work(binary, blob, theme, media, model_budget, MAX_INPUT_WORK)
}

#[cfg(feature = "direct-ppt")]
fn adopt_with_work(
    binary: &BinaryShape<'_>,
    blob: &[u8],
    theme: &Theme,
    media: &mut super::media::SpanStore,
    model_budget: &mut usize,
    mut work: usize,
) -> Result<Option<ShapeElement>, String> {
    if matches!(theme, Theme::BudgetLimited) {
        return Ok(None);
    }
    let mut text = MAX_PACKAGE_BYTES as usize;
    let mut model = *model_budget;
    let checkpoint = media.blob_checkpoint();
    let result = verify(binary, blob, theme, media, &mut work, &mut text, &mut model);
    if !matches!(result, Ok(Some(_))) {
        media.rollback_blobs(checkpoint);
    }
    match result {
        Ok(Some(shape)) => {
            *model_budget = model;
            Ok(Some(shape))
        }
        Ok(None) => Ok(None),
        Err(error) if resource_rejection(&error) => Ok(None),
        Err(error) => Err(error),
    }
}

fn resource_rejection(error: &str) -> bool {
    error.starts_with("LEGACY_RESOURCE_LIMIT:") || error.starts_with("OOXML_RESOURCE_LIMIT:")
}

/// Verify one alternative using attempt-local ledgers. The metered preflight
/// covers actual input/output, including unselected entries; legacy projection
/// rereads spend input work without double-charging inflated text.
#[cfg(feature = "direct-ppt")]
fn verify(
    binary: &BinaryShape<'_>,
    blob: &[u8],
    theme: &Theme,
    media: &mut super::media::SpanStore,
    work_budget: &mut usize,
    text_budget: &mut usize,
    model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    // One parse attempt of work. Preflight charges actual decoder input and
    // output for every entry, including unselected entries. Legacy projection
    // re-reads share the input-work ledger without charging decoded text twice.
    *work_budget = work_budget.checked_sub(1).ok_or_else(|| {
        crate::rejection::resource("PowerPoint alternative shape XML work budget exceeded")
    })?;
    measure_inflation(blob, work_budget, text_budget)?;
    let Some((kind, part)) = alternative_part(blob, work_budget)? else {
        return Ok(None);
    };
    // Only shapes and connectors are parts this projection can compare
    // with a binary shape.
    if !matches!(kind.as_str(), "shapeXml" | "connectorXml") {
        return Err(unverifiable("part"));
    }
    let Theme::Readable {
        theme_xml, clr_map, ..
    } = theme
    else {
        return Err(unverifiable("theme"));
    };
    // The modern parser can tolerate an absent/malformed style matrix. An
    // optional binary alternative has a stronger contract: its theme facts
    // must be readable before claiming equivalence with the binary shape.
    let theme_doc = parse_xml(theme_xml)?;
    let theme_root = theme_doc.root_element();
    if theme_root.tag_name().name() != "theme"
        || !ooxml_common::ns::is_a_ns(theme_root.tag_name().namespace())
    {
        return Err(unverifiable("theme"));
    }
    let parsed = pptx_parser::parse_standalone_shape_part(
        blob,
        &part,
        theme_xml,
        clr_map.as_deref(),
        MAX_PART_BYTES,
        MAX_PACKAGE_BYTES,
    );
    let parsed = parsed
        // Keep measured modern-parser resource envelopes byte-for-byte.
        .map_err(unsupported)?
        .ok_or_else(|| unverifiable("part"))?;
    // PowerPoint 16 controls retain XML-only picture crop edits when the
    // relationship-backed image bytes equal the binary BLIP, and switch to
    // binary after a BLIP-index edit changes those bytes. Compare the passive
    // package target with the already admitted binary resource. On adoption
    // the XML crop uses that resource's path; no blob media is retained.
    // External click hyperlinks use the binary text-action catalog. Other
    // references, including theme styles, remain at the typed gate.
    let (image_comparison, hyperlink_comparison, bullet_assets) = if parsed.relationship_references
    {
        validate_blob_references(blob, &part, work_budget)?;
        let bullet_assets = blob_bullet_images(blob, &part, work_budget)?;
        match (
            &parsed.element.fill,
            &binary.element.fill,
            binary.image_bytes.as_deref(),
        ) {
            (
                Some(pptx_model::Fill::Image { image_path, .. }),
                Some(pptx_model::Fill::Image { .. }),
                Some(binary_bytes),
            ) => {
                let blob_bytes = sole_blob_image(blob, &part, image_path, work_budget)?;
                (
                    Some(if blob_bytes == binary_bytes {
                        Verdict::Same
                    } else {
                        Verdict::Differs("image content")
                    }),
                    None,
                    None,
                )
            }
            _ if blob_hyperlinks_only(blob, &part, work_budget)? => (
                None,
                Some(same_hyperlinks(binary.element, &parsed.element)),
                None,
            ),
            _ if bullet_assets.is_some() => (None, None, bullet_assets),
            _ => return Err(unverifiable("relationship")),
        }
    } else {
        (None, None, None)
    };
    // Validate all selected style references, including ordinary shapes.
    // Missing matrix entries must not adopt the modern renderer's fallback.
    let locals = placeholder_locals(blob, &part, theme, work_budget)?;
    let local = if parsed.placeholder {
        Some(locals)
    } else {
        None
    };
    let mut shape = parsed.element;
    let pattern_comparison = same_binary_pattern(binary, &shape)?;
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
        match substitute_text_bounded(&mut candidate, binary.text, model_budget) {
            Ok(()) => Ok(candidate),
            Err(SubstitutionError::Comparison(verdict)) => Err(verdict),
            Err(SubstitutionError::Resource(error)) => return Err(error),
        }
    };
    // A local placeholder override is ambiguous when it disagrees: XML-only
    // geometry/fill edits render from the alternative, whereas binary-only
    // edits on the same controls invalidate it and render from the binary.
    // The proprietary checksum cannot be recomputed here to identify which
    // side was edited. Reject that pair rather than guessing its provenance.
    // The controls that draw a restated attribute from the alternative are
    // ordinary shapes; a placeholder keeps the same gate for it.
    let local_override = |verdict, stated, attribute| {
        if stated
            && matches!(
                verdict,
                Verdict::Differs(_) | Verdict::Restated(_) | Verdict::Overridden(_)
            )
        {
            Verdict::Unverifiable(attribute)
        } else {
            verdict
        }
    };
    let verdicts = [
        hyperlink_comparison.unwrap_or(Verdict::Same),
        bullet_assets.as_ref().map_or(Verdict::Same, |assets| {
            same_bullet_fallback(binary.element, &shape, assets)
        }),
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
                image_comparison
                    .or(pattern_comparison)
                    .unwrap_or_else(|| same_fill(&binary.fill, &shape, binary.element.rotation)),
                parsed.placeholder,
                "placeholder fill precedence",
            )
        },
        same_run_formatting(
            binary.element,
            &shape,
            bullet_assets.as_ref().and(binary.direct_size_authored),
        ),
        match &substituted {
            Ok(_) => Verdict::Same,
            Err(verdict) => *verdict,
        },
    ];
    let Decision::Alternative(restated) = decide(&verdicts)? else {
        return Ok(None);
    };
    shape = substituted.unwrap_or_else(|_| unreachable!("text agreement was decided"));
    if let Some(assets) = bullet_assets {
        let mut retained = std::collections::BTreeMap::new();
        for (path, extension, bytes) in assets {
            let id = media.admit_blob_image(extension, bytes)?;
            retained.insert(path, format!("legacy-ppt/image/{id}"));
        }
        let body = shape
            .text_body
            .as_mut()
            .ok_or_else(|| unverifiable("picture bullet"))?;
        for paragraph in &mut body.paragraphs {
            if let pptx_model::Bullet::Blip { image_path, .. } = &mut paragraph.bullet {
                *image_path = retained
                    .get(image_path)
                    .ok_or_else(|| unverifiable("picture bullet"))?
                    .clone();
            }
        }
    }
    let element = binary.element;
    if image_comparison == Some(Verdict::Same) {
        if let (
            Some(pptx_model::Fill::Image {
                image_path,
                mime_type,
                ..
            }),
            Some(pptx_model::Fill::Image {
                image_path: binary_path,
                mime_type: binary_mime,
                ..
            }),
        ) = (&mut shape.fill, &element.fill)
        {
            *image_path = binary_path.clone();
            *mime_type = binary_mime.clone();
        }
    }
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
    } else if shape.geometry != "custGeom" {
        // PowerPoint 16 binary-only controls for an adjusted upArrow (the
        // alternative has an empty avLst) write the binary guide on save;
        // changing that binary value changes the PDF. An omitted XML guide
        // leaves that guide to the binary shape, even when the preset's
        // DrawingML default is unknown or the binary guide is non-default.
        for (xml, binary) in [
            (&mut shape.adj, element.adj),
            (&mut shape.adj2, element.adj2),
            (&mut shape.adj3, element.adj3),
            (&mut shape.adj4, element.adj4),
            (&mut shape.adj5, element.adj5),
            (&mut shape.adj6, element.adj6),
            (&mut shape.adj7, element.adj7),
            (&mut shape.adj8, element.adj8),
        ] {
            if xml.is_none() {
                *xml = binary;
            }
        }
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
    .map_err(|_| crate::rejection::resource("PowerPoint direct slide model budget exceeded"))?;
    *model_budget = model_budget.checked_sub(retained).ok_or_else(|| {
        crate::rejection::resource("PowerPoint direct slide model budget exceeded")
    })?;
    for site in restated {
        crate::rejection::approximated(site, ());
    }
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
/// shape relies on a missing layout. Inspect direct `p:spPr` components and
/// authored `p:style` references to distinguish local style from inherited
/// layout properties (ECMA-376 Annex L.3.2.3). The archive and part sizes
/// have already passed `measure_inflation` and the PPTX
/// standalone parser's bounds; this second read is limited to placeholders.
#[cfg(feature = "direct-ppt")]
fn placeholder_locals(
    blob: &[u8],
    part: &str,
    theme: &Theme,
    work_budget: &mut usize,
) -> Result<PlaceholderLocals, String> {
    use ooxml_common::ns::{is_a_ns, is_p_ns};
    let mut archive = zip_index::open(blob, work_budget)?;
    let xml = open_part(&mut archive, part)?.read_to_string(work_budget, MAX_PART_BYTES)?;
    let doc = parse_xml(&xml)?;
    let sp_pr = doc.root_element().children().find(|n| {
        n.is_element() && n.tag_name().name() == "spPr" && is_p_ns(n.tag_name().namespace())
    });
    let style = doc.root_element().children().find(|n| {
        n.is_element() && n.tag_name().name() == "style" && is_p_ns(n.tag_name().namespace())
    });
    let has = |names: &[&str]| {
        sp_pr.is_some_and(|sp_pr| {
            sp_pr.children().any(|n| {
                n.is_element()
                    && is_a_ns(n.tag_name().namespace())
                    && names.contains(&n.tag_name().name())
            })
        })
    };
    let style_has = |name| {
        style.is_some_and(|style| {
            style.children().any(|n| {
                n.is_element() && is_a_ns(n.tag_name().namespace()) && n.tag_name().name() == name
            })
        })
    };
    if let Some(style) = style {
        let format = theme
            .format_scheme()
            .ok_or_else(|| unverifiable("placeholder style theme"))?;
        for reference in style
            .children()
            .filter(|node| node.is_element() && is_a_ns(node.tag_name().namespace()))
        {
            let selected = match reference.tag_name().name() {
                "fillRef" => format.lookup_fill_ref(
                    reference
                        .attribute("idx")
                        .and_then(|idx| idx.parse().ok())
                        .ok_or_else(|| unverifiable("fill style"))?,
                ),
                "lnRef" => format.lookup_line_ref(
                    reference
                        .attribute("idx")
                        .and_then(|idx| idx.parse().ok())
                        .ok_or_else(|| unverifiable("placeholder line style"))?,
                ),
                "effectRef" => format.lookup_effect_ref(
                    reference
                        .attribute("idx")
                        .and_then(|idx| idx.parse().ok())
                        .ok_or_else(|| unverifiable("placeholder effect style"))?,
                ),
                _ => continue,
            };
            if matches!(selected, ooxml_common::theme::StyleMatrixLookup::Missing) {
                return Err(unverifiable("placeholder style reference"));
            }
        }
    }
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
        stroke: has(&["ln"]) || style_has("lnRef"),
        // PowerPoint-saved placeholders with only an effectRef retain their
        // theme-resolved shadow. The standalone parser resolves that style;
        // replacing it with a shadowless binary fallback drops that shadow.
        effects: has(&["effectLst", "effectDag"]) || style_has("effectRef"),
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
    _media: &mut super::media::SpanStore,
    _model_budget: &mut usize,
) -> Result<Option<ShapeElement>, String> {
    Err(unverifiable("part"))
}

#[cfg(feature = "direct-ppt")]
mod zip_index;

/// Library resource policy for actual decompressor work, not ZIP metadata.
/// Drain with a fixed buffer so a crossing fails before retaining an oversized
/// part. The text ledger covers each entry once per adoption attempt. The
/// input-work ledger charges decoder consumption, including zero-output blocks,
/// and is shared by later legacy projection reads within this attempt.
/// The modern parser replays this same immutable, fully measured package under
/// its own output limits; preflight bounds that replay's zero-output work too.
/// Validate size consistency only after inflation: forged declarations are
/// malformed ZIP data, not evidence of a resource violation. This also keeps
/// the ordinary PPTX parser's declared-size checks from becoming admission
/// policy for an alternative that has already passed this actual-byte gate.
#[cfg(feature = "direct-ppt")]
fn measure_inflation(
    blob: &[u8],
    work_budget: &mut usize,
    text_budget: &mut usize,
) -> Result<(), String> {
    let unreadable = || unsupported("unreadable PowerPoint alternative shape XML package");
    let oversized =
        || crate::rejection::resource("PowerPoint alternative shape XML exceeds its size limit");
    let mut archive = zip_index::open(blob, work_budget)?;
    let mut total = 0u64;
    let mut buffer = [0; 8192];
    for index in 0..archive.len() {
        if archive
            .by_index_raw(index)
            .map_err(|_| unreadable())?
            .encrypted()
        {
            return Err(crate::rejection::encryption(
                "encrypted PowerPoint alternative ZIP entry",
            ));
        }
        let mut entry = crate::opc_part::Entry::open(&mut archive, index)?;
        let mut actual = 0u64;
        loop {
            let allowance = (MAX_PART_BYTES - actual)
                .min(MAX_PACKAGE_BYTES - total)
                .min(*text_budget as u64)
                .saturating_add(1)
                .min(buffer.len() as u64) as usize;
            let count = entry.read(&mut buffer[..allowance], work_budget)?;
            if count == 0 {
                break;
            }
            super::charge_text(text_budget, count)?;
            actual += count as u64;
            total += count as u64;
            if actual > MAX_PART_BYTES || total > MAX_PACKAGE_BYTES {
                return Err(oversized());
            }
        }
    }
    Ok(())
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
    // PowerPoint 16 saves/PDFs of an upArrow whose XML avLst is empty retain
    // its adjusted binary guide, including after a binary-only edit. Omission
    // therefore does not assert the DrawingML default against a stated binary
    // guide. Two stated guides are compared within the rounding of the
    // binary anchor; when only the XML states a guide, the normative preset
    // default is the binary side of the comparison when defined.
    let defaults = crate::officeart::preset_defaults::defaults(&element.geometry);
    let mut verdict = Verdict::Same;
    for (index, (bounds, value)) in binary
        .adjust_bounds
        .iter()
        .zip(adjusts(alternative))
        .enumerate()
    {
        let default = defaults.and_then(|d| d.get(index)).map(|&v| f64::from(v));
        match (bounds, value) {
            (_, None) => {}
            (Some((low, high)), Some(value)) => {
                if value < low - 1.0 || high + 1.0 < value {
                    return Verdict::Differs("preset adjust values");
                }
            }
            (None, Some(value)) => match default {
                Some(default) if (value - default).abs() <= 1.0 => {}
                Some(_) => return Verdict::Differs("preset adjust values"),
                None => verdict = Verdict::Unverifiable("preset adjust values"),
            },
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
///
/// Commands that do not pair (another count, or another kind at one
/// position) are compared as outlines instead (`same_outline`): the same
/// traced contours, with another start point or with collinear lines split
/// differently, are the same shape. The PowerPoint 16 control whose binary
/// start vertex was moved along the first edge (adding a binary line the
/// alternative does not state) prints like its source and like the control
/// without an alternative, so neither side's edit is visible there. Other
/// unpaired structures stay unverifiable: no control establishes which side
/// PowerPoint draws for them. So does a stroked path with dashes or
/// arrowheads, which follow the traversal that the outline comparison
/// normalizes away. A solid stroke draws collinear splits and another start
/// point alike, but draws joins where a line turns back on itself (that
/// control's binary path does, at its moved start); such a stroked path keeps
/// the binary, which draws its own turns as authored and which the control
/// without an alternative shows PowerPoint printing alike.
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
        let (raw_b, raw_a) = (bp, ap);
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
        if bp.len() != ap.len()
            || bp
                .iter()
                .zip(&ap)
                .any(|(bc, ac)| std::mem::discriminant(*bc) != std::mem::discriminant(*ac))
        {
            // Dashes and arrowheads follow the traversal (start point and
            // direction of each segment), which the outline comparison
            // normalizes away. Solid, unterminated strokes draw alike up to
            // the joins of lines that turn back on themselves, which
            // `same_outline` reports.
            let traversal = |stroke: &Option<pptx_model::Stroke>| {
                stroke.as_ref().is_some_and(|stroke| {
                    !stroke.custom_dash.is_empty()
                        || stroke.dash_style.as_deref().is_some_and(|d| d != "solid")
                        || [&stroke.head_end, &stroke.tail_end]
                            .into_iter()
                            .flatten()
                            .any(|end| end.kind != "none")
                })
            };
            if (b_paint[index].1 || a_paint.1)
                && (traversal(&binary.element.stroke) || traversal(&alternative.stroke))
            {
                return Verdict::Unverifiable("custom geometry");
            }
            match same_outline(raw_b, raw_a, b_paint[index].1 || a_paint.1) {
                Verdict::Same => {}
                Verdict::Differs(what) => verdict = Verdict::Differs(what),
                Verdict::Overridden(site) if verdict == Verdict::Same => {
                    verdict = Verdict::Overridden(site)
                }
                Verdict::Overridden(_) => {}
                other => return other,
            }
            continue;
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
                (
                    PathCmd::ArcTo {
                        wr,
                        hr,
                        st_ang,
                        sw_ang,
                    },
                    PathCmd::ArcTo {
                        wr: awr,
                        hr: ahr,
                        st_ang: ast_ang,
                        sw_ang: asw_ang,
                    },
                ) => vec![(wr, awr), (hr, ahr), (st_ang, ast_ang), (sw_ang, asw_ang)],
                (PathCmd::Close, PathCmd::Close) => Vec::new(),
                // Equal kinds this comparator has no operands for.
                _ => return Verdict::Unverifiable("custom geometry"),
            };
            if pairs.iter().any(|(p, q)| (*p - *q).abs() > PATH_TOLERANCE) {
                verdict = Verdict::Differs("custom geometry vertices");
            }
        }
    }
    verdict
}

/// One drawn segment of a contour, with its end point; `Line` also keeps its
/// start for the collinear merge.
#[cfg(any(test, feature = "direct-ppt"))]
#[derive(Clone, Copy, Debug)]
enum Segment {
    Line([f64; 2], [f64; 2]),
    Arc([f64; 4], [f64; 2]),
    Cubic([f64; 4], [f64; 2]),
}

#[cfg(any(test, feature = "direct-ppt"))]
impl Segment {
    fn end(&self) -> [f64; 2] {
        match *self {
            Segment::Line(_, end) | Segment::Arc(_, end) | Segment::Cubic(_, end) => end,
        }
    }
}

/// A contour: its start point, whether it is closed, and its segments.
#[cfg(any(test, feature = "direct-ppt"))]
type Contour = ([f64; 2], bool, Vec<Segment>);

/// The contours of one path. `None` when a
/// command has no end point this reader can place exactly: an arc whose
/// angles are not multiples of 90 degrees (where parametric and visual
/// ECMA-376 angle readings differ), a quadratic Bézier, or drawing without
/// a start point.
#[cfg(any(test, feature = "direct-ppt"))]
fn contours(path: &[pptx_model::PathCmd]) -> Option<Vec<Contour>> {
    use pptx_model::PathCmd;
    fn quarter(degrees: f64) -> Option<(f64, f64)> {
        let turns = degrees / 90.0;
        if (turns - turns.round()).abs() > PATH_TOLERANCE {
            return None;
        }
        Some(match (turns.round() as i64).rem_euclid(4) {
            0 => (1.0, 0.0),
            1 => (0.0, 1.0),
            2 => (-1.0, 0.0),
            _ => (0.0, -1.0),
        })
    }
    let mut result: Vec<Contour> = Vec::new();
    let mut current: Option<[f64; 2]> = None;
    for command in path {
        let segment = match *command {
            PathCmd::MoveTo { x, y } => {
                result.push(([x, y], false, Vec::new()));
                current = Some([x, y]);
                continue;
            }
            PathCmd::Close => {
                let (start, closed, segments) = result.last_mut()?;
                if *closed {
                    return None;
                }
                let at = current?;
                segments.push(Segment::Line(at, *start));
                *closed = true;
                current = Some(*start);
                continue;
            }
            PathCmd::LineTo { x, y } => Segment::Line(current?, [x, y]),
            PathCmd::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x,
                y,
            } => Segment::Cubic([x1, y1, x2, y2], [x, y]),
            PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } => {
                let at = current?;
                let (sc, ss) = quarter(st_ang)?;
                let (ec, es) = quarter(st_ang + sw_ang)?;
                let end = [at[0] - wr * sc + wr * ec, at[1] - hr * ss + hr * es];
                Segment::Arc([wr, hr, st_ang, sw_ang], end)
            }
            _ => return None,
        };
        let (_, closed, segments) = result.last_mut()?;
        if *closed {
            return None;
        }
        current = Some(segment.end());
        segments.push(segment);
    }
    Some(result)
}

/// Two ECMA-376 20.1.9 paths compared as traced contours, for command lists
/// that do not pair one to one. Zero-length lines draw nothing, and a run of
/// collinear lines whose points all lie between the run's first start and
/// last end traces exactly one line between them, so both are normalized
/// away; a closed contour may start at any of its segment boundaries. The
/// normalized contours must then pair segment by segment, kind by kind,
/// like paired commands: equal points are the same outline, and other
/// points are a different vertex (`Differs`, as for paired commands). A
/// contour that does not normalize to the same kinds stays unverifiable.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_outline(
    binary: &[pptx_model::PathCmd],
    alternative: &[pptx_model::PathCmd],
    stroked: bool,
) -> Verdict {
    let near = |p: [f64; 2], q: [f64; 2]| {
        (p[0] - q[0]).abs() <= PATH_TOLERANCE && (p[1] - q[1]).abs() <= PATH_TOLERANCE
    };
    // The distance of `p` from the line through `a` and `b` (`a` != `b`).
    let off_line = |a: [f64; 2], b: [f64; 2], p: [f64; 2]| {
        let d = [b[0] - a[0], b[1] - a[1]];
        ((d[0] * (p[1] - a[1]) - d[1] * (p[0] - a[0])) / d[0].hypot(d[1])).abs()
    };
    // Merge a cyclic (closed) or linear run of segments; `None` when the
    // contour degenerates to collinear lines only.
    let reversed = std::cell::Cell::new(false);
    let normalize = |closed: bool, segments: Vec<Segment>| -> Option<Vec<Segment>> {
        let mut segments: Vec<Segment> = segments
            .into_iter()
            .filter(|s| !matches!(s, Segment::Line(a, b) if near(*a, *b)))
            .collect();
        let continues = |s: &Segment, t: &Segment| match (s, t) {
            (Segment::Line(a, b), Segment::Line(_, c)) => off_line(*a, *b, *c) <= PATH_TOLERANCE,
            _ => false,
        };
        if closed {
            // Rotate so the list starts where a run of collinear lines
            // cannot continue across the boundary.
            let n = segments.len();
            let boundary =
                (0..n).find(|&i| !continues(&segments[(i + n - 1) % n], &segments[i]))?;
            segments.rotate_left(boundary);
        }
        // A run of lines traces exactly the line from its first start to
        // its last end when every point of the run lies on that line,
        // between those two ends.
        let mut merged: Vec<Segment> = Vec::with_capacity(segments.len());
        let mut index = 0;
        while index < segments.len() {
            let Segment::Line(first, _) = segments[index] else {
                merged.push(segments[index]);
                index += 1;
                continue;
            };
            let mut end = index + 1;
            while end < segments.len() && continues(&segments[end - 1], &segments[end]) {
                end += 1;
            }
            let last = segments[end - 1].end();
            let length = (last[0] - first[0]).hypot(last[1] - first[1]);
            let along = |p: [f64; 2]| {
                ((p[0] - first[0]) * (last[0] - first[0])
                    + (p[1] - first[1]) * (last[1] - first[1]))
                    / length
            };
            let traced = length > PATH_TOLERANCE
                && segments[index..end].iter().all(|s| {
                    let p = s.end();
                    off_line(first, last, p) <= PATH_TOLERANCE
                        && (-PATH_TOLERANCE..=length + PATH_TOLERANCE).contains(&along(p))
                });
            if end - index > 1 && traced {
                // A run that turns back on itself traces the same points,
                // but a stroke draws joins at its turns.
                let mut previous = 0.0;
                for s in &segments[index..end] {
                    let at = along(s.end());
                    if at < previous - PATH_TOLERANCE {
                        reversed.set(true);
                    }
                    previous = at;
                }
                merged.push(Segment::Line(first, last));
            } else {
                merged.extend_from_slice(&segments[index..end]);
            }
            index = end;
        }
        Some(merged)
    };
    let (Some(b), Some(a)) = (contours(binary), contours(alternative)) else {
        return Verdict::Unverifiable("custom geometry");
    };
    if b.len() != a.len() {
        return Verdict::Unverifiable("custom geometry");
    }
    let mut differs = false;
    for ((bs, bc, bseg), (as_, ac, aseg)) in b.into_iter().zip(a) {
        let (Some(bseg), Some(aseg)) = (normalize(bc, bseg), normalize(ac, aseg)) else {
            return Verdict::Unverifiable("custom geometry");
        };
        if bc != ac || bseg.len() != aseg.len() || bseg.is_empty() {
            return Verdict::Unverifiable("custom geometry");
        }
        // An open contour pairs from its start point; a closed one at any
        // rotation. Rotations are found in linear time on exact keys: the
        // segment kinds, and the points rounded to the comparison
        // tolerance; a rounded match is then checked within it.
        let kinds = |segments: &[Segment]| -> Vec<u8> {
            segments
                .iter()
                .map(|s| match s {
                    Segment::Line(..) => 0,
                    Segment::Arc(..) => 1,
                    Segment::Cubic(..) => 2,
                })
                .collect()
        };
        let values = |s: &Segment| -> Vec<f64> {
            match s {
                Segment::Line(_, e) => e.to_vec(),
                Segment::Arc(p, e) | Segment::Cubic(p, e) => p.iter().chain(e).copied().collect(),
            }
        };
        let rounded = |segments: &[Segment]| -> Vec<(u8, Vec<i64>)> {
            kinds(segments)
                .into_iter()
                .zip(segments)
                .map(|(kind, s)| {
                    let key = values(s)
                        .into_iter()
                        .map(|v| (v / PATH_TOLERANCE).round() as i64)
                        .collect();
                    (kind, key)
                })
                .collect()
        };
        let same_at = |offset: usize| {
            let n = bseg.len();
            (bc || near(bs, as_))
                && (0..n).all(|i| {
                    values(&bseg[(i + offset) % n])
                        .iter()
                        .zip(values(&aseg[i]))
                        .all(|(p, q)| (p - q).abs() <= PATH_TOLERANCE)
                })
        };
        fn offset<T: PartialEq>(closed: bool, b: &[T], a: &[T]) -> Option<usize> {
            if closed {
                cyclic_offset(b, a)
            } else {
                (b == a).then_some(0)
            }
        }
        if offset(bc, &rounded(&bseg), &rounded(&aseg)).is_some_and(same_at) {
            continue;
        }
        if offset(bc, &kinds(&bseg), &kinds(&aseg)).is_none() {
            return Verdict::Unverifiable("custom geometry");
        }
        differs = true;
    }
    if differs {
        Verdict::Differs("custom geometry vertices")
    } else if stroked && reversed.get() {
        // The stroke of the alternative may draw other joins: keep the
        // binary, which draws its own turns as authored.
        Verdict::Overridden(RESTATED_GEOMETRY)
    } else {
        Verdict::Same
    }
}

/// A rotation `offset` with `text[(i + offset) % n] == pattern[i]` for all
/// `i`, found with the Knuth-Morris-Pratt failure function over the doubled
/// text in O(n).
#[cfg(any(test, feature = "direct-ppt"))]
fn cyclic_offset<T: PartialEq>(text: &[T], pattern: &[T]) -> Option<usize> {
    let n = pattern.len();
    if n == 0 || text.len() != n {
        return None;
    }
    let mut failure = vec![0usize; n];
    let mut k = 0;
    for i in 1..n {
        while k > 0 && pattern[i] != pattern[k] {
            k = failure[k - 1];
        }
        if pattern[i] == pattern[k] {
            k += 1;
        }
        failure[i] = k;
    }
    let mut k = 0;
    for i in 0..2 * n - 1 {
        let c = &text[i % n];
        while k > 0 && *c != pattern[k] {
            k = failure[k - 1];
        }
        if *c == pattern[k] {
            k += 1;
        }
        if k == n {
            return Some(i + 1 - n);
        }
    }
    None
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
/// DrawingML kind), except comparable gradient and pattern projections.
/// PowerPoint 16 controls show XML pattern edits and binary image edits both
/// affect PDF output. `same_binary_pattern` compares the image/pattern pair
/// using its decoded 8x8 bitmap and duotone endpoints; where the bitmap does
/// not decode or the preset's cells were not measured, the binary pattern is
/// drawn (its PDFs show the stored bitmap's cells at the binary scale). Two gradients, with
/// their stops and rotation flag, are compared by `same_paint` below.
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
        (Some(a), Some(b)) => same_paint(a, b, rotation),
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

/// Comparable gradient or pattern paint. Gradient stop positions are 16.16
/// fractions in the binary (MS-ODRAW 2.3.7.17 fillShadeColors) and
/// 1/100000 in DrawingML (ECMA-376 20.1.8.36), so each may round by one
/// step of either unit; angles are whole degrees on both sides. PowerPoint 16
/// uses the XML in Office-saved multi-stop gradients whose binary duplicates
/// the first colour at position zero and synthesizes a final stop from the
/// scalar back colour. The XML omits the redundant first stop and retains the
/// terminal colour at a position that the binary does not encode. Controls
/// with three and five binary stops, two and four XML stops, and changed
/// XML/binary colours establish this restatement; compare the authored
/// interior positions and colours, and the terminal colour, without inventing
/// a terminal position for the binary.
///
/// Other unequal stops are restated and drawn from the alternative. This is
/// the precedence the measured controls show, not an encoding limit: the
/// PowerPoint 16 PDFs of controls whose XML alone changes a stop colour (five and three binary
/// stops) or moves an interior stop by 19% print the XML stops. Moving it by
/// 1% prints no visible change and decides nothing. No control edits a binary
/// stop that this comparison reads (the binary-only colour controls edit
/// `fillColor`, which the binary projection of these gradients does not
/// draw), so PowerPoint's choice for that pairing is not measured; it is
/// drawn like the measured XML-only edits.
/// ECMA-376 CT_GradientFillProperties gives `flip` the default `none`; a
/// `tileRect` with all zero offsets covers the whole shape, like no tileRect.
/// A difference in the angle, type, scaling, path or rectangles has no
/// control and stays unverifiable.
///
/// Rotation with the shape, on 90- and 180-degree rotated shapes
/// (PowerPoint 16 PDF controls; without rotation the flag draws nothing):
///
/// | binary | XML | PDF | verdict |
/// |---|---|---|---|
/// | false, stated or not | omitted | XML, rotating like `1` | same |
/// | false | `0` | XML | same |
/// | false | `1` | XML | restated |
/// | true | omitted | binary | overridden |
/// | true | `1` | not measured; equal | same |
/// | true | `0` | not measured; a stated disagreement like the row above | overridden |
///
/// PowerPoint writes no binary flag for either XML value, so a binary
/// `false` does not contradict the XML, while a binary `true` is a later
/// edit of the binary unless the XML states the same.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_paint(binary: &pptx_model::Fill, alternative: &pptx_model::Fill, rotation: f64) -> Verdict {
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
    fn comparable_stops(binary: &[(f64, &str)], alternative: &[(f64, &str)]) -> bool {
        let matching = |b: &(f64, &str), a: &(f64, &str)| {
            (b.0 - a.0).abs() <= POSITION && same_color(b.1, a.1) == Some(true)
        };
        if binary.len() == alternative.len() {
            return binary.iter().zip(alternative).all(|(b, a)| matching(b, a));
        }
        // The 0 stop duplicates the first authored colour; the final binary
        // stop is synthesized from fillBackColor at 1, not authored in
        // fillShadeColors. PowerPoint's XML retains its independent terminal
        // position, which must follow the last binary-authored stop.
        if binary.len() != alternative.len() + 1 || binary.len() < 3 {
            return false;
        }
        let redundant_start =
            binary[0].0.abs() <= POSITION && same_color(binary[0].1, binary[1].1) == Some(true);
        let authored = binary[1..binary.len() - 1]
            .iter()
            .zip(&alternative[..alternative.len() - 1])
            .all(|(b, a)| matching(b, a));
        let end = (binary.last().unwrap(), alternative.last().unwrap());
        redundant_start
            && authored
            && (end.0 .0 - 1.0).abs() <= POSITION
            && end.1 .0 + POSITION >= binary[binary.len() - 2].0
            && end.1 .0 <= 1.0 + POSITION
            && same_color(end.0 .1, end.1 .1) == Some(true)
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
            // The centre rectangle is a 16.16 fraction in the binary
            // (MS-ODRAW 2.3.7.16-19) and 1/100000 in DrawingML (ECMA-376
            // 20.1.10.48), so each edge may round by one step of either unit.
            let same_focus = match (rect(bf), rect(af)) {
                (None, None) => true,
                (Some(b), Some(a)) => b.iter().zip(&a).all(|(b, a)| (b - a).abs() <= POSITION),
                _ => false,
            };
            if !(turn.min(360.0 - turn) < 0.01
                && bt == at
                && bsc == asc
                && bp == ap
                && same_focus
                && bfl.as_deref().unwrap_or("none") == afl.as_deref().unwrap_or("none"))
            {
                return Verdict::Unverifiable("fill");
            }
            if tile_rect(btr) != tile_rect(atr) {
                // The binary has no tile rectangle (MS-ODRAW 2.3.7), so the
                // XML one restates nothing the binary could contradict.
                // PowerPoint 16 writes `<a:tileRect l="-100000" t="-100000"/>`
                // for its corner shade, whose PDF has the centre at the
                // shape's own corner, as the binary centre rectangle says.
                return if bp.is_some() {
                    Verdict::Overridden(RESTATED_FILL)
                } else {
                    Verdict::Unverifiable("fill")
                };
            }
            let rotation = match (br, ar) {
                _ if rotation_irrelevant || br == ar => Verdict::Same,
                (Some(false), None) => Verdict::Same,
                (Some(false), Some(_)) => Verdict::Restated(RESTATED_FILL),
                (Some(true), _) => Verdict::Overridden(RESTATED_FILL),
                (None, _) => Verdict::Unverifiable("fill"),
            };
            // A concentric shade keeps its binary stops, whatever its
            // rotation flags say: the PowerPoint 16 PDF of a control whose
            // XML alone gives an interior stop 50% opacity (the binary shade
            // colours have none) prints that stop opaque.
            if bp.is_some() && !comparable_stops(&bs, &as_) {
                return Verdict::Overridden(RESTATED_FILL);
            }
            match rotation {
                Verdict::Same if !comparable_stops(&bs, &as_) => Verdict::Restated(RESTATED_FILL),
                rotation => rotation,
            }
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
        ) if same_color(bf, af) == Some(true) && same_color(bb, ab) == Some(true) && bp == ap => {
            Verdict::Same
        }
        _ => Verdict::Unverifiable("fill"),
    }
}

/// PowerPoint 16 picture-bullet controls with two distinct embedded GIFs
/// saved a plain U+2022 OfficeArt bullet as their down-revision fallback.
/// XML-only image replacement changed the PDF, while changing that binary
/// fallback to U+25A0 produced exactly the binary-only PDF. Thus only the
/// observed U+2022 fallback is compatible with an XML buBlip; a different
/// binary marker invalidates the alternative. This applies to picture bullet
/// references only, with every target resolved and retained above.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_bullet_fallback(
    binary: &ShapeElement,
    alternative: &ShapeElement,
    assets: &[BlobBulletAsset],
) -> Verdict {
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return Verdict::Unverifiable("picture bullet");
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return Verdict::Unverifiable("picture bullet");
    }
    let mut found = false;
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        if let pptx_model::Bullet::Blip { image_path, .. } = &ap.bullet {
            found = true;
            if !assets.iter().any(|(path, _, _)| path == image_path) {
                return Verdict::Unverifiable("picture bullet resource");
            }
            match &bp.bullet {
                pptx_model::Bullet::Char { ch, .. } if ch == "•" => {}
                _ => return Verdict::Differs("picture bullet fallback"),
            }
        }
    }
    if found {
        Verdict::Same
    } else {
        Verdict::Unverifiable("picture bullet")
    }
}

/// [MS-PPT] 2.6.10/2.9.57 text click ranges refer to the document's
/// ExHyperlinkContainer targets. PowerPoint 16 ignores an XML-only r:id URL
/// edit in a saved legacy deck, so a differing target selects binary. Compare
/// per UTF-16 position because the encodings may split style runs differently.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_hyperlinks(binary: &ShapeElement, alternative: &ShapeElement) -> Verdict {
    if binary.hyperlink != alternative.hyperlink
        || binary.hyperlink_action != alternative.hyperlink_action
    {
        return Verdict::Differs("shape hyperlink");
    }
    type Link<'a> = [Option<&'a str>; 4];
    let (Some(b), Some(a)) = (&binary.text_body, &alternative.text_body) else {
        return if binary.text_body.is_none() && alternative.text_body.is_none() {
            Verdict::Same
        } else {
            Verdict::Unverifiable("text hyperlink")
        };
    };
    if b.paragraphs.len() != a.paragraphs.len() {
        return Verdict::Differs("text hyperlink");
    }
    fn spans(runs: &[TextRun]) -> Option<impl Iterator<Item = (usize, Link<'_>)>> {
        if runs.iter().any(|run| matches!(run, TextRun::Math { .. })) {
            return None;
        }
        Some(
            runs.iter()
                .map(|run| match run {
                    TextRun::Text(data) => (
                        data.text.encode_utf16().count(),
                        [
                            data.hyperlink.as_deref(),
                            data.hyperlink_action.as_deref(),
                            data.hyperlink_mouse_over.as_deref(),
                            data.hyperlink_mouse_over_action.as_deref(),
                        ],
                    ),
                    _ => (1, [None; 4]),
                })
                .filter(|(len, _)| *len > 0),
        )
    }
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        let (Some(mut bx), Some(mut ax)) = (spans(&bp.runs), spans(&ap.runs)) else {
            return Verdict::Unverifiable("text hyperlink");
        };
        let (mut bs, mut aspan) = (bx.next(), ax.next());
        // Compare each run overlap once. Expanding per UTF-16 position would
        // allocate large scratch arrays and repeatedly compare long URLs.
        while let (Some((bn, bl)), Some((an, al))) = (bs, aspan) {
            if bl != al {
                return Verdict::Differs("text hyperlink");
            }
            let common = bn.min(an);
            bs = if bn == common {
                bx.next()
            } else {
                Some((bn - common, bl))
            };
            aspan = if an == common {
                ax.next()
            } else {
                Some((an - common, al))
            };
        }
        if bs.is_some() || aspan.is_some() {
            return Verdict::Differs("text hyperlink");
        }
    }
    Verdict::Same
}

/// Run formatting both forms state must agree character by character: font
/// size, bold and italic, where each side specifies them. The binary stores
/// whole points only; a corpus deck whose alternative says 10.5 pt over a
/// binary 10 pt renders in PowerPoint without the alternative's other run
/// properties (its character spacing), while a deck whose sizes agree
/// renders them. A field compares as one unit, because the binary
/// projection substitutes its displayed text.
#[cfg(any(test, feature = "direct-ppt"))]
fn same_run_formatting(
    binary: &ShapeElement,
    alternative: &ShapeElement,
    direct_size_authored: Option<&[bool]>,
) -> Verdict {
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
    fn units(runs: &[TextRun]) -> Option<impl Iterator<Item = (Unit, Format)> + Clone + '_> {
        if runs.iter().any(|run| matches!(run, TextRun::Math { .. })) {
            return None;
        }
        // Stream the logical positions rather than allocating a multi-word
        // tuple for every character. The text budget bounds traversal work.
        Some(runs.iter().flat_map(|run| {
            let (unit, format, count) = match run {
                TextRun::Text(d) => (
                    (if d.field_type.is_some() {
                        Unit::Field
                    } else {
                        Unit::Character
                    }),
                    (d.font_size, d.bold, d.italic),
                    if d.field_type.is_some() {
                        1
                    } else {
                        d.text.encode_utf16().count()
                    },
                ),
                _ => (Unit::Break, (None, None, None), 1),
            };
            std::iter::repeat_n((unit, format), count)
        }))
    }
    let agree = |x: Option<f64>, y: Option<f64>| match (x, y) {
        (Some(x), Some(y)) => (x - y).abs() < 1e-6,
        _ => true,
    };
    let agree_bool = |x: Option<bool>, y: Option<bool>| x.zip(y).is_none_or(|(x, y)| x == y);
    let mut verdict = Verdict::Same;
    let mut position = 0;
    for (bp, ap) in b.paragraphs.iter().zip(&a.paragraphs) {
        let (Some(bx), Some(ax)) = (units(&bp.runs), units(&ap.runs)) else {
            return Verdict::Unverifiable("equation");
        };
        let binary_count = bx.clone().count();
        if direct_size_authored.is_some_and(|mask| position + binary_count > mask.len()) {
            return Verdict::Unverifiable("run size origin");
        }
        // A masked alternative states a line break as one text character
        // (see `substitute_text`); its formatting still aligns.
        let kind = |unit: Unit| {
            if unit == Unit::Break {
                Unit::Character
            } else {
                unit
            }
        };
        if binary_count != ax.clone().count()
            || bx
                .clone()
                .zip(ax.clone())
                .any(|(x, y)| kind(x.0) != kind(y.0))
        {
            // `substitute_text` decides the text structure.
            position += binary_count + 1;
            continue;
        }
        if !bx.zip(ax).enumerate().all(|(index, (x, y))| {
            let size_stated = direct_size_authored.is_none_or(|mask| mask[position + index]);
            (!size_stated || agree(x.1 .0, y.1 .0))
                && agree_bool(x.1 .1, y.1 .1)
                && agree_bool(x.1 .2, y.1 .2)
        }) {
            verdict = Verdict::Differs("run formatting");
        }
        position += binary_count + 1;
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
/// longer run). Restore it as the receiving model's in-run LF, keeping the
/// alternative's run boundaries and formatting. Authored `a:br` stays a
/// distinct break with its own rPr. Other break characters inside a run
/// have no observed form and are not compared. An equation has no binary
/// characters to compare.
#[cfg(any(test, feature = "direct-ppt"))]
#[cfg(test)]
fn substitute_text(shape: &mut ShapeElement, text: Option<&str>) -> Result<(), SubstitutionError> {
    let mut budget = usize::MAX;
    substitute_text_bounded(shape, text, &mut budget)
}

// A resource failure is not an attribute-comparison verdict: it must abort
// even when another attribute differs and would keep the binary projection.
#[derive(Debug, PartialEq)]
enum SubstitutionError {
    Comparison(Verdict),
    Resource(String),
}
impl From<Verdict> for SubstitutionError {
    fn from(verdict: Verdict) -> Self {
        Self::Comparison(verdict)
    }
}
impl From<String> for SubstitutionError {
    fn from(error: String) -> Self {
        Self::Resource(error)
    }
}

fn substitute_text_bounded(
    shape: &mut ShapeElement,
    text: Option<&str>,
    model_budget: &mut usize,
) -> Result<(), SubstitutionError> {
    const DIFFERS: Verdict = Verdict::Differs("text structure");
    let Some(body) = shape.text_body.as_mut() else {
        return if text.is_none_or(str::is_empty) {
            Ok(())
        } else {
            Err(DIFFERS.into())
        };
    };
    let text = text.unwrap_or("");
    let paragraphs: Vec<&str> = text.split('\r').collect();
    if paragraphs.len() != body.paragraphs.len() {
        // A body without text still has one empty paragraph in both forms.
        return Err(DIFFERS.into());
    }
    for (paragraph, source) in body.paragraphs.iter_mut().zip(paragraphs) {
        // Reserve scratch against the same cumulative model budget before
        // expanding text into UTF-16, independently of final JSON charging.
        let unit_count = source.encode_utf16().count();
        let scratch = unit_count
            .checked_mul(std::mem::size_of::<u16>())
            .ok_or_else(|| {
                crate::rejection::resource("PowerPoint text substitution model budget exceeded")
            })?;
        *model_budget = model_budget.checked_sub(scratch).ok_or_else(|| {
            crate::rejection::resource("PowerPoint text substitution model budget exceeded")
        })?;
        let mut units: Vec<u16> = source.encode_utf16().collect();
        let mut at = 0usize;
        let mut runs = Vec::new();
        for run in std::mem::take(&mut paragraph.runs) {
            match run {
                TextRun::Text(mut data) => {
                    let len = data.text.encode_utf16().count();
                    let slice = at
                        .checked_add(len)
                        .and_then(|end| units.get_mut(at..end))
                        .ok_or(DIFFERS)?;
                    at += len;
                    if slice.iter().any(|&u| matches!(u, 0x0a | 0x2028))
                        || (data.field_type.is_some() && slice.contains(&0x0b))
                    {
                        return Err(Verdict::Unverifiable("line break").into());
                    }
                    // This character came from inside the alternative's a:t,
                    // not from an authored a:br. Keep its enclosing run and
                    // normalize binary VT to the model's in-run LF carrier.
                    // The receiver already sizes empty lines opened by an LF
                    // from that run; inventing a:br nodes loses this
                    // ownership and changes final insertion-line formatting.
                    // No endParaRPr or binary/XML style precedence is inferred.
                    let owned = slice.len().checked_mul(3).ok_or_else(|| {
                        crate::rejection::resource(
                            "PowerPoint text substitution model budget exceeded",
                        )
                    })?;
                    *model_budget = model_budget.checked_sub(owned).ok_or_else(|| {
                        crate::rejection::resource(
                            "PowerPoint text substitution model budget exceeded",
                        )
                    })?;
                    for unit in slice.iter_mut() {
                        if *unit == 0x0b {
                            *unit = 0x0a;
                        }
                    }
                    data.text = String::from_utf16(slice).map_err(|_| DIFFERS)?;
                    text_style::direct_model::reserve_run_slot(&mut runs, model_budget)?;
                    runs.push(TextRun::Text(data));
                }
                line_break @ TextRun::Break { .. } => {
                    if !matches!(units.get(at), Some(0x0b | 0x0a | 0x2028)) {
                        return Err(DIFFERS.into());
                    }
                    at += 1;
                    text_style::direct_model::reserve_run_slot(&mut runs, model_budget)?;
                    runs.push(line_break);
                }
                TextRun::Math { .. } => return Err(Verdict::Unverifiable("equation").into()),
            }
        }
        if at != units.len() {
            return Err(DIFFERS.into());
        }
        paragraph.runs = runs;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pptx_model::{Fill, PathCmd};

    #[test]
    fn round_trip_theme_exposes_authored_hyperlink_colors() {
        let theme = Theme::Readable {
            theme_xml: "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements><a:clrScheme name=\"Office\"><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>".into(),
            clr_map: None,
            format_scheme: std::cell::OnceCell::new(),
        };
        assert_eq!(
            theme.hyperlink_colors(),
            Some(("0000FF".into(), "800080".into()))
        );
    }

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
            direct_size_authored: None,
            fill: RecordedFill::Stated(element.fill.clone().map(Box::new)),
            path_paint: element
                .cust_geom
                .as_ref()
                .map(|paths| vec![(true, true); paths.len()]),
            adjust_bounds: adjusts(element).map(|v| v.map(|v| (v, v))),
            image_bytes: None,
        }
    }

    fn solid(color: &str) -> Option<Fill> {
        Some(Fill::Solid {
            color: color.to_owned(),
        })
    }

    #[test]
    fn a_verified_difference_keeps_the_binary_and_an_unknown_one_fails_closed() {
        use crate::rejection::take_best_effort;
        use Verdict::*;
        let alternative = |sites: &[&'static str]| Ok(Decision::Alternative(sites.to_vec()));
        take_best_effort();
        assert_eq!(decide(&[Same, Same]), alternative(&[]));
        // A verified difference decides alone: an overriding encoding beside
        // it replaces no rejection.
        assert_eq!(
            decide(&[
                Same,
                Differs("fill"),
                Overridden(RESTATED_FILL),
                Unverifiable("geometry"),
                Restated(RESTATED_FILL)
            ]),
            Ok(Decision::Binary)
        );
        let error = decide(&[Same, Restated(RESTATED_FILL), Unverifiable("fill")]).unwrap_err();
        assert!(
            error.starts_with("UNSUPPORTED:") && error.contains("fill"),
            "{error}"
        );
        // Restated attributes are returned to the caller, which records them
        // only once the alternative is adopted.
        assert_eq!(
            decide(&[Restated(RESTATED_FILL), Same]),
            alternative(&[RESTATED_FILL])
        );
        assert!(take_best_effort().is_empty());
        // An overriding binary encoding keeps the binary and outranks
        // restatements and unknowns; it is recorded as the decision.
        assert_eq!(
            decide(&[
                Restated(RESTATED_FILL),
                Overridden(RESTATED_FILL),
                Unverifiable("fill")
            ]),
            Ok(Decision::Binary)
        );
        assert_eq!(
            take_best_effort(),
            [(crate::rejection::BestEffort::Approximated, RESTATED_FILL)]
        );
        // The site is the rejection these outcomes replace.
        assert_eq!(unverifiable("fill"), format!("UNSUPPORTED:{RESTATED_FILL}"));
    }

    #[test]
    fn preset_adjusts_compare_within_the_rounding_of_their_binary_inputs() {
        let leaf = pptx_model::Transform::default();
        let element = shape("roundRect");
        let mut alternative = shape("roundRect");
        let mut shape = binary(&element, &leaf);
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        // An omitted XML guide leaves the binary adjusted guide in force.
        shape.adjust_bounds[0] = Some((16660.0, 16670.0));
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
        shape.adjust_bounds[0] = Some((29000.0, 31000.0));
        assert_eq!(same_geometry(&shape, &alternative), Verdict::Same);
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
        // upDownArrow twice instead), but XML omission is still verifiable
        // because PowerPoint's binary-only control supplies the guide.
        let up = self::shape("upArrow");
        let mut shape = binary(&up, &leaf);
        shape.adjust_bounds[1] = Some((50000.0, 50000.0));
        assert_eq!(
            same_geometry(&shape, &self::shape("upArrow")),
            Verdict::Same
        );
    }

    #[test]
    fn cyclic_offsets_are_found_in_linear_time() {
        assert_eq!(cyclic_offset(&[3, 1, 2], &[1, 2, 3]), Some(1));
        assert_eq!(cyclic_offset(&[1, 1, 2, 1], &[1, 2, 1, 1]), Some(1));
        assert_eq!(cyclic_offset(&[1, 2, 3], &[1, 3, 2]), None);
        assert_eq!(cyclic_offset::<u8>(&[], &[]), None);
        assert_eq!(cyclic_offset(&[1, 2], &[1, 2, 1]), None);
        // A long periodic contour that differs in one place: the search stays
        // linear (the former all-rotations scan was quadratic).
        let n = 200_000;
        let text: Vec<u8> = (0..n).map(|i| u8::from(i == n - 1)).collect();
        let mut pattern = vec![0u8; n];
        pattern[n / 2] = 1;
        assert_eq!(cyclic_offset(&text, &pattern), Some(n / 2 - 1));
        pattern[n / 2] = 2;
        assert_eq!(cyclic_offset(&text, &pattern), None);
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
        // Commands that do not pair are compared as outlines. A quadratic
        // Bézier has no exact comparison here.
        assert_eq!(
            same_geometry(&binary(&element, &leaf), &curved),
            Verdict::Unverifiable("custom geometry")
        );
        // A binary that starts inside the first edge and closes through the
        // corner traces the same outline in the same order: the same shape.
        let forward = |corner: f64| {
            let mut element = shape("custGeom");
            element.cust_geom = Some(vec![vec![
                PathCmd::MoveTo { x: 0.1, y: 0.0 },
                PathCmd::LineTo { x: corner, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
                PathCmd::LineTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: 0.1, y: 0.0 },
                PathCmd::Close,
            ]]);
            element
        };
        assert_eq!(
            same_geometry(&binary(&forward(1.0), &leaf), &square(false, 1.0)),
            Verdict::Same
        );
        // Once normalized, the outlines pair like paired commands: another
        // corner is a different vertex.
        assert_eq!(
            same_geometry(&binary(&forward(0.9), &leaf), &square(false, 1.0)),
            Verdict::Differs("custom geometry vertices")
        );
        // A binary start vertex moved along the first edge with the line
        // back to the corner turns back on itself there: the same filled
        // outline, but a stroke draws joins at the turns. A stroked path
        // keeps the binary, which draws them as authored; an unstroked one
        // is the same shape.
        let mut moved = square(true, 1.0);
        if let Some(paths) = moved.cust_geom.as_mut() {
            paths[0][0] = PathCmd::MoveTo { x: 0.1, y: 0.0 };
            paths[0].insert(1, PathCmd::LineTo { x: 0.0, y: 0.0 });
        }
        assert_eq!(
            same_geometry(&binary(&moved, &leaf), &square(false, 1.0)),
            Verdict::Overridden(RESTATED_GEOMETRY)
        );
        // A different vertex beside the fold is a verified difference.
        assert_eq!(
            same_geometry(&binary(&moved, &leaf), &square(false, 0.9)),
            Verdict::Differs("custom geometry vertices")
        );
        assert_eq!(
            same_outline(
                moved.cust_geom.as_ref().unwrap()[0].as_slice(),
                square(false, 1.0).cust_geom.as_ref().unwrap()[0].as_slice(),
                false
            ),
            Verdict::Same
        );
        // Unpaired commands that trace another outline stay unknown: a
        // vertex off the edge, a spike past the edge's end (a stroke would
        // draw it), or a bend behind a tiny first segment, whose direction
        // must not stand for the run's line.
        for extra in [
            vec![PathCmd::LineTo { x: 0.5, y: 0.1 }],
            vec![PathCmd::LineTo { x: 1.2, y: 0.0 }],
            vec![
                PathCmd::LineTo { x: 1e-8, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 0.09 },
            ],
        ] {
            let mut other = square(true, 1.0);
            if let Some(paths) = other.cust_geom.as_mut() {
                for (offset, command) in extra.into_iter().enumerate() {
                    paths[0].insert(1 + offset, command);
                }
            }
            assert_eq!(
                same_geometry(&binary(&other, &leaf), &square(false, 1.0)),
                Verdict::Unverifiable("custom geometry")
            );
        }
        // A long closed contour that starts elsewhere compares in linear
        // time.
        let n = 10_000;
        let vertex = |i: usize| {
            let angle = std::f64::consts::TAU * (i % n) as f64 / n as f64;
            (0.5 + 0.5 * angle.cos(), 0.5 + 0.5 * angle.sin())
        };
        // The binary also splits its first edge at the midpoint, so the
        // command lists do not pair.
        let polygon = |from: usize, closing: bool| {
            let (x, y) = vertex(from);
            let mut commands = vec![PathCmd::MoveTo { x, y }];
            if closing {
                let (nx, ny) = vertex(from + 1);
                commands.push(PathCmd::LineTo {
                    x: (x + nx) / 2.0,
                    y: (y + ny) / 2.0,
                });
            }
            for i in 1..n + usize::from(closing) {
                let (x, y) = vertex(from + i);
                commands.push(PathCmd::LineTo { x, y });
            }
            commands.push(PathCmd::Close);
            let mut element = shape("custGeom");
            element.cust_geom = Some(vec![commands]);
            element
        };
        let started = std::time::Instant::now();
        assert_eq!(
            same_geometry(&binary(&polygon(n / 3, true), &leaf), &polygon(0, false)),
            Verdict::Same
        );
        assert!(started.elapsed() < std::time::Duration::from_secs(5));
        // A verified difference in the paint of the unpaired path is
        // reported as such and keeps the binary.
        let mut unfilled = binary(&moved, &leaf);
        unfilled.path_paint = Some(vec![(false, true)]);
        assert_eq!(
            same_geometry(&unfilled, &square(false, 1.0)),
            Verdict::Differs("custom geometry path paint")
        );
        let two = |first: &ShapeElement, corner: f64| {
            let mut element = first.clone();
            if let Some(paths) = element.cust_geom.as_mut() {
                paths.extend(square(false, corner).cust_geom.unwrap());
            }
            element
        };
        let binary_two = two(&forward(1.0), 1.0);
        assert_eq!(
            same_geometry(&binary(&binary_two, &leaf), &two(&square(false, 1.0), 1.0)),
            Verdict::Same
        );
        assert_eq!(
            same_geometry(&binary(&binary_two, &leaf), &two(&square(false, 1.0), 0.9)),
            Verdict::Differs("custom geometry vertices")
        );
        // Another number of paths is not a pairing of commands.
        assert_eq!(
            same_geometry(&binary(&binary_two, &leaf), &square(false, 1.0)),
            Verdict::Unverifiable("custom geometry")
        );
        // Equal kinds without comparable operands stay unverifiable.
        assert_eq!(
            same_geometry(&binary(&curved, &leaf), &curved),
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
        // A concentric shade keeps its binary form when the alternative
        // adds a tile rectangle or restates a stop.
        let concentric = |color: &str, tile: Option<[f64; 4]>| {
            let mut fill = gradient(&[(0.0, "FF0000"), (1.0, color)], Some(true));
            if let Some(Fill::Gradient {
                grad_type,
                path,
                fill_to_rect,
                tile_rect,
                ..
            }) = fill.as_mut()
            {
                *grad_type = "radial".to_owned();
                *path = Some("rect".to_owned());
                let rect = |[l, t, r, b]: [f64; 4]| ooxml_common::fill::FillRect { l, t, r, b };
                *fill_to_rect = Some(rect([0.0, 0.0, 1.0, 1.0]));
                *tile_rect = tile.map(rect);
            }
            fill
        };
        // The centre rectangle compares within the rounding of 16.16 and
        // 1/100000 fractions: 0x4ccd / 65536 is DrawingML 30000.
        let focused = |edge: f64| {
            let mut fill = concentric("0000FF", None);
            if let Some(Fill::Gradient { fill_to_rect, .. }) = fill.as_mut() {
                *fill_to_rect = Some(ooxml_common::fill::FillRect {
                    l: edge,
                    t: edge,
                    r: 1.0 - edge,
                    b: 1.0 - edge,
                });
            }
            fill
        };
        let binary_focus = RecordedFill::Stated(focused(f64::from(0x4ccd) / 65536.0).map(Box::new));
        alternative.fill = focused(0.3);
        assert_eq!(same_fill(&binary_focus, &alternative, 0.0), Verdict::Same);
        alternative.fill = focused(0.3001);
        assert_eq!(
            same_fill(&binary_focus, &alternative, 0.0),
            Verdict::Unverifiable("fill")
        );
        // On a rotated shape the binary stops win before the XML's
        // rotation flag is weighed.
        let unrotating = |color: &str, rotate: Option<bool>| {
            let mut fill = concentric(color, None);
            if let Some(Fill::Gradient { rot_with_shape, .. }) = fill.as_mut() {
                *rot_with_shape = rotate;
            }
            fill
        };
        let fixed = RecordedFill::Stated(unrotating("0000FF", Some(false)).map(Box::new));
        alternative.fill = unrotating("0000FF80", Some(true));
        assert_eq!(
            same_fill(&fixed, &alternative, 90.0),
            Verdict::Overridden(RESTATED_FILL)
        );
        alternative.fill = unrotating("0000FF", Some(true));
        assert_eq!(
            same_fill(&fixed, &alternative, 90.0),
            Verdict::Restated(RESTATED_FILL)
        );
        let corner = RecordedFill::Stated(concentric("0000FF", None).map(Box::new));
        for (fill, expected) in [
            (concentric("0000FF", None), Verdict::Same),
            (concentric("0000FF", Some([0.0; 4])), Verdict::Same),
            (
                concentric("0000FF", Some([-1.0, -1.0, 0.0, 0.0])),
                Verdict::Overridden(RESTATED_FILL),
            ),
            (
                concentric("0000FF80", None),
                Verdict::Overridden(RESTATED_FILL),
            ),
        ] {
            alternative.fill = fill;
            assert_eq!(same_fill(&corner, &alternative, 0.0), expected);
        }
        let recorded = RecordedFill::Stated(
            gradient(&[(0.0, "000000"), (0.3099975, "FFFFFF")], Some(true)).map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], Some(true));
        assert_eq!(same_fill(&recorded, &alternative, 90.0), Verdict::Same);
        // PowerPoint's binary 0 stop may duplicate the first authored shade;
        // its scalar back colour has no authored stop position. The XML keeps
        // the latter position, including values short of 100%.
        let resampled = RecordedFill::Stated(
            gradient(
                &[
                    (0.0, "000000"),
                    (0.18, "000000"),
                    (0.39, "172DA6"),
                    (0.61, "00B0F0"),
                    (1.0, "20A472"),
                ],
                Some(false),
            )
            .map(Box::new),
        );
        alternative.fill = gradient(
            &[
                (0.18, "000000"),
                (0.39, "172DA6"),
                (0.61, "00B0F0"),
                (0.92, "20A472"),
            ],
            None,
        );
        assert_eq!(same_fill(&resampled, &alternative, 180.0), Verdict::Same);
        // The same restatement also occurs with three binary stops and two
        // XML stops. A distinct colour at zero is not the redundant stop.
        let short = RecordedFill::Stated(
            gradient(
                &[(0.0, "000000"), (0.31, "000000"), (1.0, "FFFFFF")],
                Some(false),
            )
            .map(Box::new),
        );
        alternative.fill = gradient(&[(0.31, "000000"), (0.88, "FFFFFF")], None);
        assert_eq!(same_fill(&short, &alternative, 90.0), Verdict::Same);
        let nonredundant = RecordedFill::Stated(
            gradient(
                &[(0.0, "172DA6"), (0.31, "000000"), (1.0, "FFFFFF")],
                Some(false),
            )
            .map(Box::new),
        );
        assert_eq!(
            same_fill(&nonredundant, &alternative, 90.0),
            Verdict::Restated(RESTATED_FILL)
        );
        alternative.fill = gradient(
            &[
                (0.18, "000000"),
                (0.39, "172DA6"),
                (0.61, "00B0F0"),
                (0.92, "20A472"),
            ],
            None,
        );
        // Stops that are not this restatement are drawn from the
        // alternative: a moved interior stop, or another stop colour.
        if let Some(Fill::Gradient { stops, .. }) = alternative.fill.as_mut() {
            stops[1].position = 0.42;
        }
        assert_eq!(
            same_fill(&resampled, &alternative, 180.0),
            Verdict::Restated(RESTATED_FILL)
        );
        if let Some(Fill::Gradient { stops, .. }) = alternative.fill.as_mut() {
            stops[1].position = 0.39;
            stops[3].color = "FFFFFF".into();
        }
        assert_eq!(
            same_fill(&resampled, &alternative, 180.0),
            Verdict::Restated(RESTATED_FILL)
        );
        alternative.fill = gradient(&[(0.31, "FFFFFF"), (0.0, "000000")], Some(true));
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
        assert_eq!(same_fill(&no_rotation, &alternative, 90.0), Verdict::Same);
        // Rotation with a rotated shape. A binary `false` states nothing
        // against the XML; a binary `true` is the binary's own statement.
        let stops = [(0.31, "FFFFFF"), (0.0, "000000")];
        for (binary, xml, verdict) in [
            (false, Some(false), Verdict::Same),
            (false, Some(true), Verdict::Restated(RESTATED_FILL)),
            (true, None, Verdict::Overridden(RESTATED_FILL)),
            (true, Some(false), Verdict::Overridden(RESTATED_FILL)),
            (true, Some(true), Verdict::Same),
        ] {
            let recorded = if binary { &recorded } else { &no_rotation };
            alternative.fill = gradient(&stops, xml);
            assert_eq!(
                same_fill(recorded, &alternative, 90.0),
                verdict,
                "{binary} {xml:?}"
            );
            // Without rotation the flag draws nothing.
            assert_eq!(same_fill(recorded, &alternative, 360.0), Verdict::Same);
        }
        // A binary rotation flag outranks restated stops.
        alternative.fill = gradient(&[(0.0, "000000"), (0.5, "FFFFFF")], None);
        assert_eq!(
            same_fill(&recorded, &alternative, 90.0),
            Verdict::Overridden(RESTATED_FILL)
        );
        assert_eq!(
            same_fill(&no_rotation, &alternative, 90.0),
            Verdict::Restated(RESTATED_FILL)
        );
        // The remaining gradient properties have no control.
        if let Some(Fill::Gradient { angle, .. }) = alternative.fill.as_mut() {
            *angle = 45.0;
        }
        assert_eq!(
            same_fill(&no_rotation, &alternative, 90.0),
            Verdict::Unverifiable("fill")
        );
        // Neither do two fills of different kinds.
        alternative.fill = Some(Fill::Pattern {
            fg: "000000".into(),
            bg: "FFFFFF".into(),
            preset: "pct30".into(),
        });
        assert_eq!(
            same_fill(&no_rotation, &alternative, 90.0),
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
                        TextRun::Break { .. } => "<br>".to_owned(),
                        TextRun::Math { .. } => "<math>".to_owned(),
                    })
                    .collect()
            })
            .collect()
    }

    #[test]
    fn masked_text_scratch_and_restored_characters_are_charged_before_projection() {
        let mut alternative = shape("rect");
        alternative.text_body = body(&[&[run("__________")]]);
        // 10 UTF-16 units of scratch, bounded UTF-8 output, one initial
        // two-slot run allocation. Deleting either string charge must fail
        // the successful exact-boundary assertion, not just an error case.
        let required = 10 * 2 + 10 * 3 + 2 * std::mem::size_of::<TextRun>();
        let mut budget = required;
        let mut admitted = alternative.clone();
        assert_eq!(
            substitute_text_bounded(
                &mut admitted,
                Some("A\u{b}A\u{b}A\u{b}A\u{b}A\u{b}"),
                &mut budget
            ),
            Ok(())
        );
        assert_eq!(budget, 0);
        assert_eq!(texts(&admitted), [vec!["A\nA\nA\nA\nA\n"]]);
        for available in [0, 10 * 2, required - 1] {
            let mut rejected = alternative.clone();
            let Err(SubstitutionError::Resource(error)) = substitute_text_bounded(
                &mut rejected,
                Some("A\u{b}A\u{b}A\u{b}A\u{b}A\u{b}"),
                &mut available.clone(),
            ) else {
                panic!("substitution quota must remain a resource failure")
            };
            assert!(error.starts_with("LEGACY_RESOURCE_LIMIT:"), "{error}");
        }
    }

    #[test]
    fn masked_text_retains_line_break_formatting() {
        let mut alternative = shape("rect");
        let mut styled = run("___");
        styled["fontSize"] = serde_json::json!(24.0);
        styled["bold"] = serde_json::json!(true);
        alternative.text_body = body(&[&[
            styled,
            serde_json::json!({"type":"break", "fontSize":36.0, "fontFamily":"Arial", "italic":true}),
        ]]);
        substitute_text(&mut alternative, Some("A\u{b}B\u{b}")).unwrap();
        let runs = &alternative.text_body.as_ref().unwrap().paragraphs[0].runs;
        assert!(matches!(
            &runs[0], TextRun::Text(v)
                if v.text == "A\nB" && v.font_size == Some(24.0) && v.bold == Some(true)
        ));
        assert!(
            matches!(&runs[1], TextRun::Break { font_size: Some(36.0), font_family: Some(face), italic: Some(true), .. } if face == "Arial")
        );
    }

    #[test]
    fn masked_in_run_breaks_preserve_the_run_and_do_not_invent_insertion_properties() {
        let mut alternative = shape("rect");
        let mut styled = run("_____");
        styled["fontSize"] = serde_json::json!(40.0);
        styled["fontFamily"] = serde_json::json!("Arial");
        alternative.text_body = body(&[&[styled]]);
        substitute_text(&mut alternative, Some("\u{b}A\u{b}B\u{b}")).unwrap();
        let p = &alternative.text_body.as_ref().unwrap().paragraphs[0];
        assert!(matches!(p.runs.as_slice(), [TextRun::Text(v)]
            if v.text == "\nA\nB\n" && v.font_size == Some(40.0)
                && v.font_family.as_deref() == Some("Arial")));
        assert!(p.end_run_properties.is_none());
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
                Err(SubstitutionError::Comparison(Verdict::Differs(
                    "text structure"
                ))),
                "{text:?}"
            );
        }
        // A run boundary inside a surrogate pair is rejected.
        let mut split = shape("rect");
        split.text_body = body(&[&[run(" "), run(" ")]]);
        assert!(substitute_text(&mut split, Some("😀")).is_err());
    }

    #[test]
    fn picture_bullet_requires_the_office_downrevision_marker_and_live_target() {
        let mut binary = shape("rect");
        binary.text_body = body(&[&[run("A")]]);
        binary.text_body.as_mut().unwrap().paragraphs[0].bullet = pptx_model::Bullet::Char {
            ch: "•".to_owned(),
            color: None,
            size_pct: None,
            size_pts: None,
            font_family: None,
        };
        let mut alternative = binary.clone();
        alternative.text_body.as_mut().unwrap().paragraphs[0].bullet = pptx_model::Bullet::Blip {
            image_path: "media/bullet.gif".to_owned(),
            mime_type: "image/gif".to_owned(),
            size_pct: None,
            size_pts: None,
        };
        let assets = vec![("media/bullet.gif".to_owned(), "gif", b"GIF89a".to_vec())];
        assert_eq!(
            same_bullet_fallback(&binary, &alternative, &assets),
            Verdict::Same
        );
        if let pptx_model::Bullet::Char { ch, .. } =
            &mut binary.text_body.as_mut().unwrap().paragraphs[0].bullet
        {
            *ch = "■".to_owned();
        }
        assert_eq!(
            same_bullet_fallback(&binary, &alternative, &assets),
            Verdict::Differs("picture bullet fallback")
        );
        assert_eq!(
            same_bullet_fallback(&binary, &alternative, &[]),
            Verdict::Unverifiable("picture bullet resource")
        );
    }

    #[test]
    fn pattern_decoder_ceilings_keep_resource_classification() {
        assert!(binary_pattern_cells(&vec![0; 4097])
            .unwrap_err()
            .starts_with("LEGACY_RESOURCE_LIMIT:"));
        // 513 clear codes, width fixed at 3 bits; resets produce no pixels.
        let mut data = vec![0; (513 * 3_usize).div_ceil(8)];
        for code in 0..513 {
            let bit = code * 3 + 2;
            data[bit / 8] |= 1 << (bit % 8);
        }
        assert!(decode_pattern_lzw(&data, 2, 100)
            .unwrap_err()
            .starts_with("LEGACY_RESOURCE_LIMIT:"));
    }

    #[test]
    fn binary_pattern_cells_match_measured_drawingml_presets() {
        // A synthetic single-frame GIF: clear before every literal keeps its
        // LZW code width at three bits, independent of dictionary growth.
        fn gif(rows: [u8; 8]) -> Vec<u8> {
            let mut pixels = [0u8; 100];
            for y in 0..8 {
                for x in 0..8 {
                    pixels[y * 10 + x] = u8::from(rows[y] & (0x80 >> x) != 0);
                }
            }
            let mut codes = Vec::new();
            for pixel in pixels {
                codes.extend([4u8, pixel]);
            }
            codes.push(5);
            let mut compressed = vec![0u8; (codes.len() * 3).div_ceil(8)];
            for (i, code) in codes.into_iter().enumerate() {
                for bit in 0..3 {
                    let at = i * 3 + bit;
                    compressed[at / 8] |= ((code >> bit) & 1) << (at % 8);
                }
            }
            let mut gif = b"GIF87a\x0a\x00\x0a\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff\x2c\x00\x00\x00\x00\x0a\x00\x0a\x00\x00\x02".to_vec();
            gif.push(compressed.len() as u8);
            gif.extend(compressed);
            gif.extend([0, 0x3b]);
            gif
        }
        let pct30 = [0xaa, 0x44, 0xaa, 0x11, 0xaa, 0x44, 0xaa, 0x11];
        let diagonal = [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88];
        assert_eq!(binary_pattern_cells(&gif(pct30)).unwrap(), Some(pct30));
        assert_eq!(
            binary_pattern_cells(&gif(diagonal)).unwrap(),
            Some(diagonal)
        );
        let mut changed = pct30;
        changed[0] ^= 1;
        assert_ne!(binary_pattern_cells(&gif(changed)).unwrap(), Some(pct30));
        let mut truncated = gif(pct30);
        truncated.truncate(32);
        assert!(binary_pattern_cells(&truncated).unwrap().is_none());

        // The pattern pair compares where the preset's cells were measured
        // and the bitmap decodes; otherwise the binary pattern is drawn.
        let pattern = |image: Option<Vec<u8>>, preset: &str| {
            let mut element = shape("rect");
            element.fill = serde_json::from_value(serde_json::json!({
                "fillType": "image", "imagePath": "legacy-ppt/image/1",
                "mimeType": "image/gif", "stretch": false,
                "duotone": {"clr1": "FFFFFF", "clr2": "000000"}
            }))
            .unwrap();
            let mut alternative = shape("rect");
            alternative.fill = Some(pptx_model::Fill::Pattern {
                fg: "000000".into(),
                bg: "FFFFFF".into(),
                preset: preset.into(),
            });
            let leaf = pptx_model::Transform::default();
            let mut binary = binary(&element, &leaf);
            binary.image_bytes = image;
            same_binary_pattern(&binary, &alternative).unwrap()
        };
        assert_eq!(pattern(Some(gif(pct30)), "pct30"), Some(Verdict::Same));
        assert_eq!(
            pattern(Some(gif(diagonal)), "pct30"),
            Some(Verdict::Differs("pattern fill"))
        );
        for (image, preset) in [
            (Some(gif(pct30)), "dkDnDiag"),
            (Some(truncated), "pct30"),
            (None, "pct30"),
        ] {
            assert_eq!(
                pattern(image, preset),
                Some(Verdict::Overridden(RESTATED_FILL))
            );
        }
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
        assert_eq!(texts(&adopted), [vec!["AB", "\n", "\nCD", "EF"]]);
        // Other break characters inside a run have no observed form.
        assert_eq!(
            substitute_text(&mut alternative.clone(), Some("AB\u{2028}\u{b}CDEF")),
            Err(SubstitutionError::Comparison(Verdict::Unverifiable(
                "line break"
            )))
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
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Same
        );
        // Whole-point binary sizes cannot state 10.5 pt: not the same shape.
        alternative.text_body = sized(serde_json::json!(10.5));
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Differs("run formatting")
        );
        // A picture-bullet alternative may restate an inherited master size;
        // a directly authored CFStyle size still has to agree.
        assert_eq!(
            same_run_formatting(&binary, &alternative, Some(&[false; 5])),
            Verdict::Same
        );
        assert_eq!(
            same_run_formatting(&binary, &alternative, Some(&[true; 5])),
            Verdict::Differs("run formatting")
        );
        // An unstated size on either side is not a disagreement.
        binary.text_body = sized(serde_json::Value::Null);
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
            Verdict::Same
        );
        // A field compares as one unit whatever its displayed text.
        let mut field = run("12");
        field["fieldType"] = serde_json::json!("slidenum");
        field["fontSize"] = serde_json::json!(20.0);
        binary.text_body = body(&[&[run("A"), field.clone()]]);
        field["text"] = serde_json::json!("<#>");
        field["fontSize"] = serde_json::json!(18.0);
        alternative.text_body = body(&[&[run("A"), field]]);
        assert_eq!(
            same_run_formatting(&binary, &alternative, None),
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
                alternative_part(&package, &mut usize::MAX.clone()).unwrap(),
                Some(("shapeXml".into(), "DRS/shapeXML.xml".into()))
            );
            assert!(placeholder_locals(
                &package,
                "DRS/shapeXML.xml",
                &theme(),
                &mut usize::MAX.clone()
            )
            .is_ok());
        }

        #[test]
        fn placeholder_style_references_are_local_effects_and_lines() {
            let shape = SHAPE_XML.replace(
                "</p:spPr>",
                "</p:spPr><p:style><a:lnRef idx=\"1\"/><a:effectRef idx=\"1\"/></p:style>",
            );
            let local = placeholder_locals(
                &blob(&shape),
                "drs/shapexml.xml",
                &theme(),
                &mut usize::MAX.clone(),
            )
            .unwrap();
            assert!(local.stroke);
            assert!(local.effects);
            assert!(local.fill);
            let foreign = SHAPE_XML.replace(
                "</p:spPr>",
                "</p:spPr><x:style xmlns:x=\"urn:foreign\"><x:lnRef/><x:effectRef/></x:style>",
            );
            let foreign = placeholder_locals(
                &blob(&foreign),
                "drs/shapexml.xml",
                &theme(),
                &mut usize::MAX.clone(),
            )
            .unwrap();
            assert!(!foreign.stroke);
            assert!(!foreign.effects);
            let missing = shape.replace("idx=\"1\"", "idx=\"99\"");
            assert!(placeholder_locals(
                &blob(&missing),
                "drs/shapexml.xml",
                &theme(),
                &mut usize::MAX.clone()
            )
            .err()
            .unwrap()
            .contains("placeholder style reference"));
        }

        #[test]
        fn hyperlinks_without_a_relationship_and_mouse_over_targets_compare() {
            let run = |link: &str| {
                SHAPE_XML.replace(
                    "</p:spPr>",
                    &format!(
                        "</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{link}</a:rPr><a:t>x</a:t></a:r></a:p></p:txBody>"
                    ),
                )
            };
            let rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid/\" TargetMode=\"External\"/></Relationships>";
            let part = "drs/shapexml.xml";
            // An action-only click states `r:id=""`: no relationship part.
            let jump =
                run("<a:hlinkClick r:id=\"\" action=\"ppaction://hlinkshowjump?jump=nextslide\"/>");
            let jump = package(&[(part, jump.as_str())], &[]);
            assert!(validate_blob_references(&jump, part, &mut usize::MAX.clone()).is_ok());
            assert!(blob_hyperlinks_only(&jump, part, &mut usize::MAX.clone()).unwrap());
            // A mouse-over target is an external hyperlink like a click one.
            let over = run("<a:hlinkMouseOver r:id=\"rId1\"/>");
            let over = package(
                &[(part, over.as_str()), ("drs/_rels/shapexml.xml.rels", rels)],
                &[],
            );
            assert!(validate_blob_references(&over, part, &mut usize::MAX.clone()).is_ok());
            assert!(blob_hyperlinks_only(&over, part, &mut usize::MAX.clone()).unwrap());
            // Another relationship-backed element is not a hyperlink-only blob.
            let image = run("<a:hlinkClick r:id=\"rId1\"/><a:blip r:embed=\"rId1\"/>");
            let image = package(
                &[
                    (part, image.as_str()),
                    ("drs/_rels/shapexml.xml.rels", rels),
                ],
                &[],
            );
            assert!(!blob_hyperlinks_only(&image, part, &mut usize::MAX.clone()).unwrap());

            // The mouse-over target and the action verb are compared like
            // the click target.
            let text = |over: Option<&str>, action: Option<&str>| {
                let mut element = shape("rect");
                element.text_body = body(&[&[serde_json::json!({
                    "type": "text", "text": "x", "bold": null, "italic": null,
                    "underline": false, "strikethrough": false, "strikeDouble": false,
                    "fontSize": null, "color": null, "fontFamily": null, "fieldType": null,
                    "hyperlinkUsesTextFill": false,
                    "hyperlinkAction": action, "hyperlinkMouseOver": over
                })]]);
                element
            };
            let binary = text(Some("https://example.invalid/"), Some("ppaction://a"));
            assert_eq!(same_hyperlinks(&binary, &binary.clone()), Verdict::Same);
            for other in [
                text(None, Some("ppaction://a")),
                text(Some("https://example.invalid/other"), Some("ppaction://a")),
                text(Some("https://example.invalid/"), None),
            ] {
                assert_eq!(
                    same_hyperlinks(&binary, &other),
                    Verdict::Differs("text hyperlink")
                );
            }
        }

        #[test]
        fn blob_relationship_targets_are_checked_passively() {
            let shape = SHAPE_XML.replace(
                "</p:spPr>",
                "<a:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rIdImage\"/></a:blipFill></p:spPr>",
            );
            let rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image.jpeg\"/></Relationships>";
            let parts = &[
                ("drs/shapexml.xml", shape.as_str()),
                ("drs/_rels/shapexml.xml.rels", rels),
                ("drs/media/image.jpeg", "passive: never decode"),
            ];
            let present = package(parts, &[]);
            assert!(validate_blob_references(
                &present,
                "drs/shapexml.xml",
                &mut usize::MAX.clone()
            )
            .is_ok());
            assert_eq!(
                sole_blob_image(
                    &present,
                    "drs/shapexml.xml",
                    "drs/media/image.jpeg",
                    &mut usize::MAX.clone()
                )
                .unwrap(),
                b"passive: never decode"
            );
            assert!(sole_blob_image(
                &present,
                "drs/shapexml.xml",
                "drs/media/other.jpeg",
                &mut usize::MAX.clone()
            )
            .is_err());
            let missing = package(&parts[..2], &[]);
            assert!(validate_blob_references(
                &missing,
                "drs/shapexml.xml",
                &mut usize::MAX.clone()
            )
            .is_err());

            let external = rels.replace(
                "Target=\"media/image.jpeg\"",
                "Target=\"https://example.test/image\" TargetMode=\"External\"",
            );
            let external = package(
                &[
                    ("drs/shapexml.xml", &shape),
                    ("drs/_rels/shapexml.xml.rels", &external),
                ],
                &[],
            );
            assert!(validate_blob_references(
                &external,
                "drs/shapexml.xml",
                &mut usize::MAX.clone()
            )
            .is_ok());
            assert!(sole_blob_image(
                &external,
                "drs/shapexml.xml",
                "drs/media/image.jpeg",
                &mut usize::MAX.clone()
            )
            .is_err());

            let duplicate = rels.replacen(
                "<Relationship ",
                "<Relationship Id=\"rIdImage\" Type=\"image\" Target=\"media/image.jpeg\"/><Relationship ",
                1,
            );
            let duplicate = package(
                &[
                    ("drs/shapexml.xml", &shape),
                    ("drs/_rels/shapexml.xml.rels", &duplicate),
                    ("drs/media/image.jpeg", "passive"),
                ],
                &[],
            );
            assert!(validate_blob_references(
                &duplicate,
                "drs/shapexml.xml",
                &mut usize::MAX.clone()
            )
            .is_err());
        }

        fn theme() -> Theme {
            Theme::Readable {
                theme_xml: crate::ppt::theme(),
                clr_map: None,
                format_scheme: std::cell::OnceCell::new(),
            }
        }

        fn element(color: &str) -> ShapeElement {
            let mut element = shape("rect");
            (element.width, element.height) = (1587500, 793750);
            element.fill = solid(color);
            element.id = Some("3".to_owned());
            element
        }

        fn run_verification(
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
            let result = verify(
                &shape,
                blob,
                theme,
                &mut super::media::SpanStore::new(Vec::new()),
                &mut work,
                &mut text,
                &mut model,
            );
            (result, model)
        }

        const AMPLE: (usize, usize, usize) = (usize::MAX, usize::MAX, usize::MAX);

        fn error<T>(result: Result<T, String>, prefix: &str) {
            let error = result.err().expect("expected rejection");
            assert!(error.starts_with(prefix), "{error}");
        }

        // Flag both ZIP headers: passwordless by_index must fail, but the raw
        // metadata still identifies encryption before any content is opened.
        fn encrypted_entry(mut bytes: Vec<u8>, name: &str) -> Vec<u8> {
            for (signature, flag, length, header) in
                [(b"PK\x03\x04", 6, 26, 30), (b"PK\x01\x02", 8, 28, 46)]
            {
                let offset = bytes
                    .windows(4)
                    .enumerate()
                    .find_map(|(at, sig)| {
                        (sig == signature
                            && bytes.get(at + header..at + header + name.len())
                                == Some(name.as_bytes())
                            && bytes.get(at + length..at + length + 2)
                                == Some(&(name.len() as u16).to_le_bytes()))
                        .then_some(at)
                    })
                    .unwrap();
                bytes[offset + flag] |= 1;
            }
            bytes
        }

        #[test]
        fn standalone_resource_envelope_reaches_verification_owner() {
            // 200,001 DOM units, below the embedded byte ceilings.
            let xml = SHAPE_XML.replace("</p:sp>", &format!("{}</p:sp>", "<x/>".repeat(200_000)));
            error(
                run_verification(&element("FF0000"), &blob(&xml), &theme(), AMPLE).0,
                "OOXML_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn standalone_depth_and_context_complexity_remain_resource_failures() {
            let deep = format!("{}{}", "<x>".repeat(257), "</x>".repeat(257));
            error(
                run_verification(&element("FF0000"), &blob(&deep), &theme(), AMPLE).0,
                "OOXML_RESOURCE_LIMIT:",
            );
            for color_map in [false, true] {
                let mut context = theme();
                let Theme::Readable {
                    theme_xml, clr_map, ..
                } = &mut context
                else {
                    unreachable!()
                };
                if color_map {
                    *clr_map = Some(format!("<a:clrMap xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">{}</a:clrMap>", "<x/>".repeat(200_000)));
                } else {
                    *theme_xml = theme_xml.replace(
                        "</a:theme>",
                        &format!("{}</a:theme>", "<x/>".repeat(200_000)),
                    );
                }
                error(
                    run_verification(&element("FF0000"), &blob(SHAPE_XML), &context, AMPLE).0,
                    "OOXML_RESOURCE_LIMIT:",
                );
            }
        }

        #[test]
        fn substitution_quota_is_not_overruled_by_an_attribute_disagreement() {
            let xml = SHAPE_XML.replace("</p:sp>", "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>_</a:t></a:r></a:p></p:txBody></p:sp>");
            let mut binary_element = element("0000FF"); // differs from XML red
            binary_element.text_body = body(&[&[run("A")]]);
            let leaf = pptx_model::Transform {
                cx: binary_element.width,
                cy: binary_element.height,
                ..Default::default()
            };
            let mut binary = binary(&binary_element, &leaf);
            binary.text = Some("A");
            error(
                verify(
                    &binary,
                    &blob(&xml),
                    &theme(),
                    &mut super::media::SpanStore::new(vec![]),
                    &mut AMPLE.0.clone(),
                    &mut AMPLE.1.clone(),
                    &mut 0,
                ),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn encrypted_shape_and_relationship_entries_keep_encryption_classification() {
            for name in ["_rels/.rels", "drs/shapexml.xml", "drs/downrev.xml"] {
                let bytes = encrypted_entry(blob(SHAPE_XML), name);
                error(
                    run_verification(&element("FF0000"), &bytes, &theme(), AMPLE).0,
                    "LEGACY_UNSUPPORTED_ENCRYPTION:",
                );
            }
        }

        #[test]
        fn passive_reference_readers_preserve_depth_bytes_entries_and_encryption() {
            let too_deep = format!("{}{}", "<x>".repeat(257), "</x>".repeat(257));
            error(
                validate_blob_references(
                    &package(&[("shape.xml", &too_deep)], &[]),
                    "shape.xml",
                    &mut usize::MAX.clone(),
                ),
                "LEGACY_RESOURCE_LIMIT:",
            );
            let oversized = "x".repeat(MAX_PART_BYTES as usize + 1);
            let bytes = package(&[("shape.xml", &oversized)], &[]);
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&bytes)).unwrap();
            error(
                read_blob_reference_part(&mut archive, "shape.xml", &mut usize::MAX.clone()),
                "LEGACY_RESOURCE_LIMIT:",
            );
            let encrypted = encrypted_entry(package(&[("shape.xml", "<x/>")], &[]), "shape.xml");
            let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&encrypted)).unwrap();
            error(
                read_blob_reference_part(&mut archive, "shape.xml", &mut usize::MAX.clone()),
                "LEGACY_UNSUPPORTED_ENCRYPTION:",
            );
            let parts: Vec<_> = (0..128)
                .map(|i| (format!("{i}.xml"), "<x/>".to_owned()))
                .collect();
            let refs: Vec<_> = parts
                .iter()
                .map(|(name, xml)| (name.as_str(), xml.as_str()))
                .collect();
            error(
                validate_blob_references(&package(&refs, &[]), "0.xml", &mut usize::MAX.clone()),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn theme_limits_and_encryption_are_not_unreadable_subset_verdicts() {
            let xml = crate::ppt::theme();
            let rels = r#"<Relationships><Relationship Type="x/theme" Target="theme.xml"/></Relationships>"#;
            let make = |xml: &str| {
                package(
                    &[("_rels/manager.xml.rels", rels), ("theme.xml", xml)],
                    &[("officeDocument", "manager.xml")],
                )
            };
            let read = |bytes: &[u8]| {
                let mut budget = usize::MAX;
                read_master_theme(
                    &[Record {
                        version: 0,
                        instance: 0,
                        kind: 0x040e,
                        payload: bytes,
                    }],
                    &mut budget,
                    &mut usize::MAX.clone(),
                )
            };
            for name in ["_rels/.rels", "_rels/manager.xml.rels", "theme.xml"] {
                error(
                    read(&encrypted_entry(make(&xml), name)),
                    "LEGACY_UNSUPPORTED_ENCRYPTION:",
                );
            }
            error(
                read(&make(&"x".repeat(MAX_THEME_BYTES as usize + 1))),
                "LEGACY_RESOURCE_LIMIT:",
            );
            let parts: Vec<_> = (0..MAX_THEME_ENTRIES)
                .map(|i| (format!("{i}.xml"), "<x/>".to_owned()))
                .collect();
            let refs: Vec<_> = parts
                .iter()
                .map(|(name, xml)| (name.as_str(), xml.as_str()))
                .collect();
            error(read(&package(&refs, &[])), "LEGACY_RESOURCE_LIMIT:");
        }

        #[test]
        fn theme_decoder_reports_input_work_quota_to_optional_owner() {
            let package = crate::opc_part::test_empty_deflate("_rels/.rels", 1000, b"");
            let record = Record {
                version: 0,
                instance: 0,
                kind: 0x040e,
                payload: &package,
            };
            let mut work = 100;
            error(
                read_master_theme(&[record], &mut 0, &mut work),
                "LEGACY_RESOURCE_LIMIT:",
            );
            assert_eq!(work, 0);
        }

        #[test]
        fn ordinary_alternative_requires_readable_and_resolved_theme_styles() {
            let styled = SHAPE_XML.replace("</p:sp>", r#"<p:style><a:lnRef idx="999"><a:srgbClr val="FF0000"/></a:lnRef></p:style></p:sp>"#);
            assert!(
                run_verification(&element("FF0000"), &blob(&styled), &theme(), AMPLE)
                    .0
                    .is_err()
            );
            let unreadable = Theme::Readable {
                theme_xml: "<broken".to_owned(),
                clr_map: None,
                format_scheme: std::cell::OnceCell::new(),
            };
            assert!(
                run_verification(&element("FF0000"), &blob(SHAPE_XML), &unreadable, AMPLE)
                    .0
                    .is_err()
            );
        }

        #[test]
        fn forged_theme_expanded_size_is_not_retained() {
            let theme_xml = format!("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><!--{}--></a:theme>", "x".repeat(65536));
            let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for (name, bytes) in [
                (
                    "_rels/.rels",
                    r#"<Relationships><Relationship Id="m" Type="http://schemas.microsoft.com/office/2006/relationships/officeDocument" Target="manager.xml"/></Relationships>"#,
                ),
                (
                    "_rels/manager.xml.rels",
                    r#"<Relationships><Relationship Id="t" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme.xml"/></Relationships>"#,
                ),
                ("theme.xml", theme_xml.as_str()),
            ] {
                zip.start_file(name, options).unwrap();
                zip.write_all(bytes.as_bytes()).unwrap();
            }
            let mut blob = zip.finish().unwrap().into_inner();
            fn record(payload: &[u8]) -> Record<'_> {
                Record {
                    version: 0,
                    instance: 0,
                    kind: 0x040e,
                    payload,
                }
            }
            assert!(matches!(
                read_master_theme(&[record(&blob)], &mut 100_000, &mut usize::MAX.clone()).unwrap(),
                Some(Theme::Readable { .. })
            ));
            // Central-directory uncompressed size, with real deflate data and
            // CRC intact. Charge actual reads, then reject the inconsistent header.
            let offset = blob
                .windows(4)
                .enumerate()
                .find_map(|(offset, signature)| {
                    (signature == b"PK\x01\x02"
                        && blob.get(offset + 46..offset + 55) == Some(b"theme.xml"))
                    .then_some(offset)
                })
                .unwrap();
            blob[offset + 24..offset + 28].copy_from_slice(&0u32.to_le_bytes());
            assert!(matches!(
                read_master_theme(&[record(&blob)], &mut 100_000, &mut usize::MAX.clone()).unwrap(),
                Some(Theme::Unreadable)
            ));
        }

        #[test]
        fn master_themes_share_a_decoded_byte_budget() {
            let manager_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme.xml"/></Relationships>"#;
            let blob = package(
                &[
                    ("manager.xml", "<x/>"),
                    ("_rels/manager.xml.rels", manager_rels),
                    ("theme.xml", &crate::ppt::theme()),
                ],
                &[("officeDocument", "manager.xml")],
            );
            let record = Record {
                version: 0,
                instance: 0,
                kind: 0x040e,
                payload: &blob,
            };
            let mut budget = 4096;
            assert!(matches!(
                read_master_theme(&[record], &mut budget, &mut usize::MAX.clone()).unwrap(),
                Some(Theme::Readable { .. })
            ));
            let decoded_bytes = 4096 - budget;
            assert!(decoded_bytes > 0);
            let mut budget = decoded_bytes;
            assert!(matches!(
                read_master_theme(&[record], &mut budget, &mut usize::MAX.clone()).unwrap(),
                Some(Theme::Readable { .. })
            ));
            error(
                read_master_theme(&[record], &mut budget, &mut usize::MAX.clone()),
                "LEGACY_RESOURCE_LIMIT:",
            );
            let mut remaining = 1;
            let limited = master_theme(&[record], &mut remaining).unwrap().unwrap();
            assert!(matches!(limited, Theme::BudgetLimited));
            assert_eq!(remaining, 1);
            let element = element("FF0000");
            let leaf = pptx_model::Transform::default();
            assert!(adopt(
                &binary(&element, &leaf),
                &blob,
                &limited,
                &mut super::media::SpanStore::new(Vec::new()),
                &mut usize::MAX.clone()
            )
            .unwrap()
            .is_none());
        }

        #[test]
        fn alternative_quotas_fall_back_without_spending_session_model() {
            let element = element("FF0000");
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let binary = binary(&element, &leaf);
            let theme = theme();
            let mut media = super::media::SpanStore::new(Vec::new());
            let valid = blob(SHAPE_XML);
            // Input (including zero-output blocks), metadata/index, inflated
            // output, modern DOM complexity, and model/substitution quotas all
            // reject only the optional alternative. Later attempts stay usable.
            let cases = [
                (
                    crate::opc_part::test_empty_deflate("unused", 1000, b""),
                    100,
                    usize::MAX,
                ),
                (empty_package(70_000), MAX_INPUT_WORK, usize::MAX),
                (
                    blob(&"x".repeat(MAX_PART_BYTES as usize + 1)),
                    MAX_INPUT_WORK,
                    usize::MAX,
                ),
                (
                    blob(
                        &SHAPE_XML
                            .replace("</p:sp>", &format!("{}</p:sp>", "<x/>".repeat(200_000))),
                    ),
                    MAX_INPUT_WORK,
                    usize::MAX,
                ),
                (valid.clone(), MAX_INPUT_WORK, 0),
            ];
            for (package, work, mut model) in cases {
                let before = model;
                assert!(
                    adopt_with_work(&binary, &package, &theme, &mut media, &mut model, work)
                        .unwrap()
                        .is_none()
                );
                assert_eq!(model, before);
                assert!(
                    adopt(&binary, &valid, &theme, &mut media, &mut usize::MAX.clone())
                        .unwrap()
                        .is_some()
                );
            }
        }

        #[test]
        fn model_fallback_releases_new_bullet_assets_and_keeps_prior_media() {
            let mut element = element("FF0000");
            element.text_body = body(&[&[run("A")]]);
            element.text_body.as_mut().unwrap().paragraphs[0].bullet = pptx_model::Bullet::Char {
                ch: "•".into(),
                color: None,
                size_pct: None,
                size_pts: None,
                font_family: None,
            };
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let mut binary = binary(&element, &leaf);
            binary.text = Some("A");
            let xml = SHAPE_XML.replace("</p:sp>", r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buBlip><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="image"/></a:buBlip></a:pPr><a:r><a:t>_</a:t></a:r></a:p></p:txBody></p:sp>"#);
            let package = package(
                &[
                    ("drs/shapexml.xml", &xml),
                    (
                        "drs/_rels/shapexml.xml.rels",
                        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bullet.gif"/></Relationships>"#,
                    ),
                    ("drs/media/bullet.gif", "GIF89a"),
                ],
                &[("shapeXml", "drs/shapexml.xml")],
            );
            let theme = theme();
            let make_media = || {
                let mut media = super::media::SpanStore::new(Vec::new());
                media.admit_blob_image("png", b"prior".to_vec()).unwrap();
                media
            };
            let mut ample = usize::MAX;
            assert!(
                adopt(&binary, &package, &theme, &mut make_media(), &mut ample)
                    .unwrap()
                    .is_some()
            );
            // Enough for substitution and image admission, one byte short of
            // retaining the final adopted shape. Partial image retention must
            // be undone so later shapes can use the same resource allowance.
            let mut model = usize::MAX - ample - 1;
            let before = model;
            let mut media = make_media();
            for _ in 0..2 {
                assert!(adopt(&binary, &package, &theme, &mut media, &mut model)
                    .unwrap()
                    .is_none());
                assert_eq!(model, before);
            }
            assert_eq!(media.image(1, &[], None).unwrap().unwrap().1, b"prior");
            assert_eq!(media.admit_blob_image("png", b"next".to_vec()).unwrap(), 2);
        }

        #[test]
        fn repeated_alternatives_do_not_share_binary_record_or_text_work() {
            let element = element("FF0000");
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let binary = binary(&element, &leaf);
            let theme = theme();
            let mut media = super::media::SpanStore::new(Vec::new());
            // One alternative alone exceeds the former 2M binary-record work
            // allowance but fits the existing OPC package input-work policy.
            let empty = empty_deflate_package(500_000);
            let mut source = zip::ZipArchive::new(std::io::Cursor::new(&empty)).unwrap();
            let valid = blob(SHAPE_XML);
            let mut shapes = zip::ZipArchive::new(std::io::Cursor::new(&valid)).unwrap();
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            writer.raw_copy_file(source.by_index(0).unwrap()).unwrap();
            for index in 0..shapes.len() {
                writer
                    .raw_copy_file(shapes.by_index(index).unwrap())
                    .unwrap();
            }
            let package = writer.finish().unwrap().into_inner();
            let mut model = usize::MAX;
            for _ in 0..2 {
                assert!(adopt(&binary, &package, &theme, &mut media, &mut model)
                    .unwrap()
                    .is_some());
            }
        }

        #[test]
        fn a_consistent_alternative_is_adopted_and_charged() {
            let (result, model) =
                run_verification(&element("FF0000"), &blob(SHAPE_XML), &theme(), AMPLE);
            let adopted = result.unwrap().expect("consistent alternative");
            assert_eq!(adopted.id.as_deref(), Some("3"));
            let charged = ooxml_common::json_measurement::measure_json(&adopted).unwrap();
            assert_eq!(usize::MAX - model, charged.json_bytes as usize);
            // A model budget below the adopted shape fails closed.
            let short = (usize::MAX, usize::MAX, charged.json_bytes as usize - 1);
            assert!(
                run_verification(&element("FF0000"), &blob(SHAPE_XML), &theme(), short)
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
                run_verification(&element("FF0000"), &downrev_only, &theme(), AMPLE).0,
                Ok(None)
            ));
            assert!(matches!(
                run_verification(&element("00FF00"), &blob(SHAPE_XML), &theme(), AMPLE).0,
                Ok(None)
            ));
        }

        const GRADIENT_XML: &str = r#"<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>"#;

        fn gradient(stops: &[(f64, &str)], rotate: bool) -> Option<Fill> {
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
                rot_with_shape: Some(rotate),
            })
        }

        fn gradient_stops(shape: &ShapeElement) -> Vec<(f64, String)> {
            match &shape.fill {
                Some(Fill::Gradient { stops, .. }) => stops
                    .iter()
                    .map(|stop| (stop.position, stop.color.clone()))
                    .collect(),
                other => panic!("{other:?}"),
            }
        }

        #[test]
        fn restated_gradient_stops_are_drawn_from_the_alternative() {
            use crate::rejection::{take_best_effort, BestEffort};
            let xml = SHAPE_XML.replace(
                r#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
                GRADIENT_XML,
            );
            let mut element = element("FF0000");
            // The binary holds another middle stop: not the restatement of
            // these XML stops that `same_paint` equates.
            element.fill = gradient(&[(0.0, "000000"), (0.4, "FF0000"), (1.0, "FFFFFF")], false);
            take_best_effort();
            let adopted = run_verification(&element, &blob(&xml), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("alternative with its own stops");
            assert_eq!(
                gradient_stops(&adopted),
                [
                    (0.0, "000000".to_owned()),
                    (0.5, "00FF00".to_owned()),
                    (1.0, "FFFFFF".to_owned())
                ]
            );
            assert!(matches!(
                adopted.fill,
                Some(Fill::Gradient {
                    rot_with_shape: None,
                    ..
                })
            ));
            // The transform and identifier stay the binary's.
            assert_eq!(adopted.id.as_deref(), Some("3"));
            assert_eq!(
                (adopted.width, adopted.height),
                (element.width, element.height)
            );
            assert_eq!(
                take_best_effort(),
                [(BestEffort::Approximated, RESTATED_FILL)]
            );
            // A placeholder that states the fill locally keeps its gate.
            let placeholder = xml.replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="body"/></p:nvPr>"#);
            let error = run_verification(&element, &blob(&placeholder), &theme(), AMPLE)
                .0
                .unwrap_err();
            assert!(
                error.starts_with("UNSUPPORTED:") && error.contains("placeholder fill precedence"),
                "{error}"
            );
            // Another angle is not a restatement of the stops.
            let turned = xml.replace("5400000", "2700000");
            let error = run_verification(&element, &blob(&turned), &theme(), AMPLE)
                .0
                .unwrap_err();
            assert_eq!(error, format!("UNSUPPORTED:{RESTATED_FILL}"));
            assert!(take_best_effort().is_empty());
            // An adoption that fails afterwards (here the model budget)
            // records no restatement.
            let error =
                run_verification(&element, &blob(&xml), &theme(), (usize::MAX, usize::MAX, 1))
                    .0
                    .unwrap_err();
            assert!(error.starts_with("LEGACY_RESOURCE_LIMIT:"), "{error}");
            assert!(take_best_effort().is_empty());
        }

        #[test]
        fn gradient_rotation_follows_the_xml_unless_the_binary_states_true() {
            use crate::rejection::{take_best_effort, BestEffort};
            let xml = SHAPE_XML
                .replace(
                    r#"<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>"#,
                    GRADIENT_XML,
                )
                .replace("<a:xfrm>", r#"<a:xfrm rot="5400000">"#);
            let stops = [(0.0, "000000"), (0.5, "00FF00"), (1.0, "FFFFFF")];
            let rotated = |rotate| {
                let mut element = element("FF0000");
                element.rotation = 90.0;
                element.fill = gradient(&stops, rotate);
                element
            };
            let verify = |element: &ShapeElement, xml: &str| {
                let leaf = pptx_model::Transform {
                    cx: element.width,
                    cy: element.height,
                    rot: element.rotation,
                    ..Default::default()
                };
                let (mut work, mut text, mut model) = AMPLE;
                verify(
                    &binary(element, &leaf),
                    &blob(xml),
                    &theme(),
                    &mut super::media::SpanStore::new(Vec::new()),
                    &mut work,
                    &mut text,
                    &mut model,
                )
            };
            let flag = |shape: &ShapeElement| match shape.fill {
                Some(Fill::Gradient { rot_with_shape, .. }) => rot_with_shape,
                _ => panic!("gradient"),
            };
            take_best_effort();
            // An omitted XML flag over a binary without the flag: the XML.
            let adopted = verify(&rotated(false), &xml).unwrap().expect("XML");
            assert_eq!(flag(&adopted), None);
            assert!(take_best_effort().is_empty());
            // An XML `1` over the same binary: the XML.
            let stated = xml.replace("<a:gradFill>", r#"<a:gradFill rotWithShape="1">"#);
            let adopted = verify(&rotated(false), &stated).unwrap().expect("XML");
            assert_eq!(flag(&adopted), Some(true));
            assert_eq!(adopted.rotation, 90.0);
            assert_eq!(
                take_best_effort(),
                [(BestEffort::Approximated, RESTATED_FILL)]
            );
            // A binary `true` the XML does not restate: the binary.
            for xml in [
                xml.clone(),
                xml.replace("<a:gradFill>", r#"<a:gradFill rotWithShape="0">"#),
            ] {
                assert!(matches!(verify(&rotated(true), &xml), Ok(None)));
                assert_eq!(
                    take_best_effort(),
                    [(BestEffort::Approximated, RESTATED_FILL)]
                );
            }
            let adopted = verify(&rotated(true), &stated).unwrap().expect("XML");
            assert_eq!(flag(&adopted), Some(true));
            assert!(take_best_effort().is_empty());
        }

        #[test]
        fn unpaired_custom_path_commands_compare_their_outline() {
            use crate::rejection::{take_best_effort, BestEffort};
            use pptx_model::PathCmd;
            let xml = SHAPE_XML.replace(
                r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
                r#"<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>"#,
            );
            let outline = vec![
                PathCmd::MoveTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
                PathCmd::Close,
            ];
            let mut element = element("00FF00");
            element.geometry = "custGeom".into();
            // The binary starts inside the first edge and closes through the
            // corner: the same outline, traced in the same order.
            element.cust_geom = Some(vec![vec![
                PathCmd::MoveTo { x: 0.1, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
                PathCmd::LineTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: 0.1, y: 0.0 },
                PathCmd::Close,
            ]]);
            element.fill = Some(Fill::Solid {
                color: "FF0000".into(),
            });
            take_best_effort();
            let adopted = run_verification(&element, &blob(&xml), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("the same outline adopts the alternative");
            assert_eq!(
                serde_json::to_value(&adopted.cust_geom).unwrap(),
                serde_json::to_value(vec![outline.clone()]).unwrap()
            );
            assert!(matches!(&adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
            assert_eq!(adopted.id.as_deref(), Some("3"));
            assert!(take_best_effort().is_empty());
            // Dashes and arrowheads follow the traversal that the outline
            // comparison normalizes away: such a stroked path stays unknown,
            // while a solid line adopts as above.
            for line in [
                r#"<a:prstDash val="dash"/>"#,
                r#"<a:custDash><a:ds d="300000" sp="100000"/></a:custDash>"#,
                r#"<a:tailEnd type="triangle"/>"#,
                r#"<a:prstDash val="solid"/>"#,
            ] {
                let stroked = xml.replace(
                    "</p:spPr>",
                    &format!(
                        r#"<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>{line}</a:ln></p:spPr>"#
                    ),
                );
                let result = run_verification(&element, &blob(&stroked), &theme(), AMPLE).0;
                if line.contains("solid") {
                    assert!(result.unwrap().is_some(), "{line}");
                } else {
                    let error = result.unwrap_err();
                    assert!(
                        error.starts_with("UNSUPPORTED:") && error.contains("custom geometry"),
                        "{line}: {error}"
                    );
                }
                assert!(take_best_effort().is_empty());
            }
            // The start moved along the first edge with a line back to the
            // corner (the PowerPoint 16 control): the path turns back on
            // itself, and its stroke may draw joins the alternative does not.
            // The binary is kept, as the control without an alternative
            // prints like the control.
            element.cust_geom = Some(vec![vec![
                PathCmd::MoveTo { x: 0.1, y: 0.0 },
                PathCmd::LineTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
                PathCmd::LineTo { x: 0.0, y: 0.0 },
                PathCmd::Close,
            ]]);
            assert!(matches!(
                run_verification(&element, &blob(&xml), &theme(), AMPLE).0,
                Ok(None)
            ));
            assert_eq!(
                take_best_effort(),
                [(BestEffort::Approximated, RESTATED_GEOMETRY)]
            );
            // With another corner as well, the verified difference keeps
            // the binary and replaces no rejection.
            if let Some(paths) = element.cust_geom.as_mut() {
                paths[0][3] = PathCmd::LineTo { x: 1.0, y: 0.9 };
            }
            assert!(matches!(
                run_verification(&element, &blob(&xml), &theme(), AMPLE).0,
                Ok(None)
            ));
            assert!(take_best_effort().is_empty());
            if let Some(paths) = element.cust_geom.as_mut() {
                paths[0][3] = PathCmd::LineTo { x: 1.0, y: 1.0 };
            }
            // A closing line to the moved start instead cuts another
            // triangle: no control says which side PowerPoint draws.
            if let Some(paths) = element.cust_geom.as_mut() {
                paths[0][4] = PathCmd::LineTo { x: 0.1, y: 0.0 };
            }
            let error = run_verification(&element, &blob(&xml), &theme(), AMPLE)
                .0
                .unwrap_err();
            assert!(
                error.starts_with("UNSUPPORTED:") && error.contains("custom geometry"),
                "{error}"
            );
            assert!(take_best_effort().is_empty());
            // A bend behind a tiny first segment is another outline too.
            element.cust_geom = Some(vec![vec![
                PathCmd::MoveTo { x: 0.0, y: 0.0 },
                PathCmd::LineTo { x: 1e-8, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 0.09 },
                PathCmd::LineTo { x: 1.0, y: 0.0 },
                PathCmd::LineTo { x: 1.0, y: 1.0 },
                PathCmd::Close,
            ]]);
            let error = run_verification(&element, &blob(&xml), &theme(), AMPLE)
                .0
                .unwrap_err();
            assert!(
                error.starts_with("UNSUPPORTED:") && error.contains("custom geometry"),
                "{error}"
            );
            assert!(take_best_effort().is_empty());
            // The same commands with another vertex are a verified
            // difference: the binary.
            let mut vertex = outline;
            vertex[1] = PathCmd::LineTo { x: 0.9, y: 0.0 };
            element.cust_geom = Some(vec![vertex]);
            assert!(matches!(
                run_verification(&element, &blob(&xml), &theme(), AMPLE).0,
                Ok(None)
            ));
            assert!(take_best_effort().is_empty());
        }

        #[test]
        fn placeholder_inherits_missing_shape_properties_from_the_binary() {
            let placeholder = SHAPE_XML
                .replace("<p:nvPr/>", r#"<p:nvPr><p:ph type="title"/></p:nvPr>"#)
                .replace(
                    r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1587500" cy="793750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>"#,
                    "<p:spPr/>",
                );
            let adopted =
                run_verification(&element("FF0000"), &blob(&placeholder), &theme(), AMPLE)
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
            let adopted =
                run_verification(&element("FF0000"), &blob(&placeholder), &theme(), AMPLE)
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
                let error = run_verification(&element("FF0000"), &blob(xml), &theme(), AMPLE)
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
            let adopted = verify(
                &binary,
                &blob(&xml),
                &theme(),
                &mut super::media::SpanStore::new(Vec::new()),
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
            let adopted = run_verification(&element, &blob(SHAPE_XML), &theme(), AMPLE)
                .0
                .unwrap()
                .expect("alternative formatting with binary outline");
            assert_eq!(adopted.geometry, "custGeom");
            assert!(adopted.cust_geom.is_some());
            assert!(matches!(adopted.fill, Some(Fill::Solid { color }) if color == "FF0000"));
        }

        fn inflation_package(sizes: &[usize]) -> Vec<u8> {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            for (index, size) in sizes.iter().enumerate() {
                writer
                    .start_file(
                        format!("{index}.xml"),
                        zip::write::SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Deflated),
                    )
                    .unwrap();
                writer.write_all(&vec![b'x'; *size]).unwrap();
            }
            writer.finish().unwrap().into_inner()
        }

        fn declared_size(mut bytes: Vec<u8>, size: u32) -> Vec<u8> {
            for (signature, offset) in [(b"PK\x03\x04", 22), (b"PK\x01\x02", 24)] {
                let at = bytes.windows(4).position(|sig| sig == signature).unwrap();
                bytes[at + offset..at + offset + 4].copy_from_slice(&size.to_le_bytes());
            }
            bytes
        }

        fn empty_package(count: usize) -> Vec<u8> {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            for index in 0..count {
                writer
                    .start_file(
                        format!("{index}.xml"),
                        zip::write::SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Stored),
                    )
                    .unwrap();
            }
            writer.finish().unwrap().into_inner()
        }

        fn empty_deflate_package(blocks: usize) -> Vec<u8> {
            crate::opc_part::test_empty_deflate("0.xml", blocks, b"")
        }

        #[test]
        fn zero_output_deflate_consumes_running_work_before_eof() {
            let package = empty_deflate_package(20_000);
            let mut work = 101; // one index entry, then 100 decoder input bytes
            error(
                measure_inflation(&package, &mut work, &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
            assert_eq!(work, 0);

            let small = empty_deflate_package(10);
            let mut work = 102; // exactly two index walks + two 50-byte streams
            measure_inflation(&small, &mut work, &mut 0).unwrap();
            measure_inflation(&small, &mut work, &mut 0).unwrap();
            assert_eq!(work, 0);
            error(
                measure_inflation(&small, &mut work, &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn zero_output_input_is_metered_across_verification_rereads() {
            fn with_empty_entry(blocks: usize) -> Vec<u8> {
                let bytes = empty_deflate_package(blocks);
                let mut empty = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
                let mut shape =
                    zip::ZipArchive::new(std::io::Cursor::new(blob(SHAPE_XML))).unwrap();
                let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
                writer.raw_copy_file(empty.by_index(0).unwrap()).unwrap();
                for index in 0..shape.len() {
                    writer
                        .raw_copy_file(shape.by_index(index).unwrap())
                        .unwrap();
                }
                writer.finish().unwrap().into_inner()
            }
            let element = element("FF0000");
            let leaf = pptx_model::Transform {
                cx: element.width,
                cy: element.height,
                ..Default::default()
            };
            let binary = binary(&element, &leaf);
            let theme = theme();
            let mut media = super::media::SpanStore::new(Vec::new());
            let (mut text, mut model) = (usize::MAX, usize::MAX);
            let large = with_empty_entry(500_000); // input exceeds this attempt's supplied work allowance
            let mut work = MAX_RECORDS;
            error(
                verify(
                    &binary, &large, &theme, &mut media, &mut work, &mut text, &mut model,
                ),
                "LEGACY_RESOURCE_LIMIT:",
            );
            assert_eq!(work, 0);

            let small = with_empty_entry(10);
            let mut work = MAX_RECORDS;
            assert!(
                verify(&binary, &small, &theme, &mut media, &mut work, &mut text, &mut model)
                    .unwrap()
                    .is_some()
            );
            let used = MAX_RECORDS - work;
            let mut work = 2 * used - 1;
            assert!(
                verify(&binary, &small, &theme, &mut media, &mut work, &mut text, &mut model)
                    .unwrap()
                    .is_some()
            );
            error(
                verify(
                    &binary, &small, &theme, &mut media, &mut work, &mut text, &mut model,
                ),
                "LEGACY_RESOURCE_LIMIT:",
            );
            assert!(work < used);
        }

        #[test]
        fn seventy_thousand_empty_entries_fail_closed_without_inflated_output() {
            let package = empty_package(70_000);
            error(
                measure_inflation(&package, &mut MAX_RECORDS.clone(), &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
            error(
                run_verification(&element("FF0000"), &package, &theme(), AMPLE).0,
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn empty_entries_consume_session_work_including_repeated_adoptions() {
            let package = empty_package(2);
            let mut work = 2;
            measure_inflation(&package, &mut work, &mut 0).unwrap();
            assert_eq!(work, 0);
            error(
                measure_inflation(&package, &mut work, &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn forged_capacities_fail_before_index_construction() {
            // No central directory exists. Ordinary ZIP opening would reject
            // the container, but only after trusting its index capacity. The
            // metadata-work boundary must reject first, with no allocation
            // proportional to this attacker-selected count.
            let count = 65_534u16;
            let mut classic = vec![0; count as usize];
            let central = classic.len();
            classic.extend_from_slice(&[0; 46]);
            classic[central..central + 4].copy_from_slice(b"PK\x01\x02");
            let footer = classic.len();
            classic.extend_from_slice(&[0; 22]);
            classic[footer..footer + 4].copy_from_slice(b"PK\x05\x06");
            classic[footer + 8..footer + 10].copy_from_slice(&count.to_le_bytes());
            classic[footer + 10..footer + 12].copy_from_slice(&count.to_le_bytes());
            classic[footer + 12..footer + 16].copy_from_slice(&46u32.to_le_bytes());
            classic[footer + 16..footer + 20].copy_from_slice(&(central as u32).to_le_bytes());
            // count <= directory offset: zip 2.x would reserve this large index
            // before finding that the directory has only one header.
            error(
                measure_inflation(&classic, &mut usize::MAX.clone(), &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
            let mut bytes = vec![0u8; 56 + 20 + 22];
            bytes[..4].copy_from_slice(b"PK\x06\x06");
            bytes[4..12].copy_from_slice(&44u64.to_le_bytes());
            bytes[24..32].copy_from_slice(&u64::MAX.to_le_bytes());
            bytes[32..40].copy_from_slice(&u64::MAX.to_le_bytes());
            bytes[56..60].copy_from_slice(b"PK\x06\x07");
            bytes[72..76].copy_from_slice(&1u32.to_le_bytes());
            bytes[76..80].copy_from_slice(b"PK\x05\x06");
            bytes[84..88].fill(0xff);
            error(
                measure_inflation(&bytes, &mut usize::MAX.clone(), &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn zip64_index_work_uses_physical_metadata_with_a_prepended_stub() {
            let package = empty_package(2);
            let footer = package.len() - 22;
            let mut bytes = vec![0; 4096];
            bytes[..4].copy_from_slice(b"PK\x05\x06"); // inert signature in the stub
            bytes.extend_from_slice(&package[..footer]);
            let mut tail = vec![0u8; 56 + 20];
            tail[..4].copy_from_slice(b"PK\x06\x06");
            tail[4..12].copy_from_slice(&44u64.to_le_bytes());
            tail[24..32].copy_from_slice(&2u64.to_le_bytes());
            tail[32..40].copy_from_slice(&2u64.to_le_bytes());
            let u32_at = |at| u32::from_le_bytes(package[at..at + 4].try_into().unwrap()) as u64;
            tail[40..48].copy_from_slice(&u32_at(footer + 12).to_le_bytes());
            tail[48..56].copy_from_slice(&u32_at(footer + 16).to_le_bytes());
            tail[56..60].copy_from_slice(b"PK\x06\x07");
            tail[64..72].copy_from_slice(&(footer as u64).to_le_bytes());
            tail[72..76].copy_from_slice(&1u32.to_le_bytes());
            bytes.extend_from_slice(&tail);
            bytes.extend_from_slice(&package[footer..]);
            let end = bytes.len();
            bytes[end - 14..end - 10].fill(0xff);
            let mut work = 55; // 53 ZIP64 search positions, two index entries
            measure_inflation(&bytes, &mut work, &mut 0).unwrap();
            assert_eq!(work, 0);
            // A physically present extensible-data sector makes the search
            // longer without increasing entry count or inflated output.
            let zip64 = 4096 + footer;
            bytes[zip64 + 4..zip64 + 12].copy_from_slice(&1044u64.to_le_bytes());
            bytes.splice(zip64 + 56..zip64 + 56, [0; 1000]);
            let mut work = 55;
            error(
                measure_inflation(&bytes, &mut work, &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
            assert_eq!(work, 0);
        }

        #[test]
        fn ignored_zip64_locator_cannot_hide_the_classic_index_count() {
            let mut bytes = vec![0u8; 56 + 20 + 22];
            bytes[..4].copy_from_slice(b"PK\x06\x06");
            bytes[4..12].copy_from_slice(&44u64.to_le_bytes());
            bytes[24..32].copy_from_slice(&1u64.to_le_bytes());
            bytes[32..40].copy_from_slice(&1u64.to_le_bytes());
            bytes[56..60].copy_from_slice(b"PK\x06\x07");
            bytes[72..76].copy_from_slice(&1u32.to_le_bytes());
            bytes[76..80].copy_from_slice(b"PK\x05\x06");
            bytes[84..86].copy_from_slice(&65_534u16.to_le_bytes());
            bytes[86..88].copy_from_slice(&1u16.to_le_bytes());
            // zip 2.x consults ZIP64 only when the classic total count or
            // directory offset is saturated. It would index 65,534 here.
            error(
                measure_inflation(&bytes, &mut usize::MAX.clone(), &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn retained_zip_names_exhaust_metadata_before_index_construction() {
            let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
            for index in 0..150 {
                writer
                    .start_file(
                        format!("{index}{}", "x".repeat(60_000)),
                        zip::write::SimpleFileOptions::default()
                            .compression_method(zip::CompressionMethod::Stored),
                    )
                    .unwrap();
            }
            let bytes = writer.finish().unwrap().into_inner();
            error(
                measure_inflation(&bytes, &mut usize::MAX.clone(), &mut 0),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn a_footer_in_the_comment_is_never_seen_by_the_index() {
            let mut bytes = empty_package(2);
            let footer = bytes.len() - 22;
            bytes[footer + 20..footer + 22].copy_from_slice(&23u16.to_le_bytes());
            let mut fake = [0u8; 23];
            fake[..4].copy_from_slice(b"PK\x05\x06");
            fake[8..12].copy_from_slice(&[0xfe, 0xff, 0xfe, 0xff]);
            bytes.extend_from_slice(&fake);
            // The dependency would try this later footer first, but while it
            // indexes, the comment reads as zero (shared IndexView): only the
            // reserved footer reaches its allocator, and the inert comment
            // bytes do not reject the package (OPC Annex B.2).
            measure_inflation(&bytes, &mut usize::MAX.clone(), &mut 0)
                .expect("only the reserved footer is indexed");
        }

        #[test]
        fn inflation_enforces_actual_part_package_and_session_boundaries() {
            let part = MAX_PART_BYTES as usize;
            let mut unlimited = usize::MAX;
            let exact = inflation_package(&[part; 4]);
            let mut budget = MAX_PACKAGE_BYTES as usize;
            measure_inflation(&exact, &mut usize::MAX.clone(), &mut budget).unwrap();
            assert_eq!(budget, 0);
            error(
                measure_inflation(
                    &inflation_package(&[part, part, part, part, 1]),
                    &mut usize::MAX.clone(),
                    &mut unlimited,
                ),
                "LEGACY_RESOURCE_LIMIT:",
            );
            // A forged small declaration cannot hide actual decompressor work.
            let too_large = declared_size(inflation_package(&[part + 1]), 1);
            error(
                measure_inflation(&too_large, &mut usize::MAX.clone(), &mut unlimited),
                "LEGACY_RESOURCE_LIMIT:",
            );
            let small = inflation_package(&[32]);
            let mut exact = 32;
            measure_inflation(&small, &mut usize::MAX.clone(), &mut exact).unwrap();
            assert_eq!(exact, 0);
            error(
                measure_inflation(&small, &mut usize::MAX.clone(), &mut 31),
                "LEGACY_RESOURCE_LIMIT:",
            );
        }

        #[test]
        fn oversized_declarations_are_integrity_errors_after_actual_work() {
            let forged = declared_size(inflation_package(&[32]), u32::MAX);
            let mut budget = 32;
            error(
                measure_inflation(&forged, &mut usize::MAX.clone(), &mut budget),
                "UNSUPPORTED:",
            );
            assert_eq!(budget, 0); // actual output, not the untrusted declaration
        }

        #[test]
        fn compressed_container_length_does_not_reject_a_small_alternative() {
            // ZIP permits a prepended self-extracting stub. Its bytes perform
            // no inflation and must not consume the alternative's budget.
            let mut bytes = vec![0; 4 * 1024 * 1024];
            bytes.extend(blob(SHAPE_XML));
            assert!(
                run_verification(&element("FF0000"), &bytes, &theme(), AMPLE)
                    .0
                    .unwrap()
                    .is_some()
            );
        }

        #[test]
        fn verification_preserves_rejection_classifications() {
            let element = element("FF0000");
            let unsupported = |result: (Result<Option<ShapeElement>, String>, usize)| {
                let error = result.0.expect_err("must fail closed");
                assert!(error.starts_with("UNSUPPORTED:"), "{error}");
            };
            let resource_limited = |result: (Result<Option<ShapeElement>, String>, usize)| {
                let error = result.0.expect_err("must fail closed");
                assert!(error.starts_with("LEGACY_RESOURCE_LIMIT:"), "{error}");
            };
            // Oversized actual output and exhausted budgets are resource policy.
            resource_limited(run_verification(
                &element,
                &blob(&"x".repeat(MAX_PART_BYTES as usize + 1)),
                &theme(),
                AMPLE,
            ));
            resource_limited(run_verification(
                &element,
                &blob(SHAPE_XML),
                &theme(),
                (0, usize::MAX, usize::MAX),
            ));
            resource_limited(run_verification(
                &element,
                &blob(SHAPE_XML),
                &theme(),
                (usize::MAX, 10, usize::MAX),
            ));
            unsupported(run_verification(
                &element,
                b"not a package",
                &theme(),
                AMPLE,
            ));
            unsupported(run_verification(&element, &blob("<p:sp"), &theme(), AMPLE));
            unsupported(run_verification(
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
            unsupported(run_verification(&element, &group, &theme(), AMPLE));
            let two = package(
                &[("drs/shapexml.xml", SHAPE_XML)],
                &[
                    ("shapeXml", "drs/shapexml.xml"),
                    ("connectorXml", "drs/shapexml.xml"),
                ],
            );
            unsupported(run_verification(&element, &two, &theme(), AMPLE));
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
            unsupported(run_verification(
                &patterned,
                &blob(&gradient),
                &theme(),
                AMPLE,
            ));
        }

        #[test]
        fn non_presentation_root_and_theme_image_dependency_fail_closed() {
            let element = element("FF0000");
            for root in [
                SHAPE_XML.replace(
                    "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "urn:foreign",
                ),
                SHAPE_XML.replace("p:sp", "sp"),
            ] {
                let error = run_verification(&element, &blob(&root), &theme(), AMPLE)
                    .0
                    .expect_err("foreign root cannot be adopted");
                assert!(error.starts_with("UNSUPPORTED:"), "{error}");
            }

            let xml = SHAPE_XML
                .replace("<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>", "")
                .replace(
                    "</p:sp>",
                    "<p:style><a:fillRef idx=\"1\"/></p:style></p:sp>",
                );
            for prefix in ["r", "関係"] {
                let theme = Theme::Readable {
                    theme_xml: format!(
                        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:{prefix}="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:themeElements><a:fmtScheme name="x"><a:fillStyleLst><a:blipFill><a:blip {prefix}:embed="rIdImage"/></a:blipFill></a:fillStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
                    ),
                    clr_map: None,
                    format_scheme: std::cell::OnceCell::new(),
                };
                let error = run_verification(&element, &blob(&xml), &theme, AMPLE)
                    .0
                    .expect_err("theme image relationship cannot be verified");
                assert!(
                    error.starts_with("UNSUPPORTED:") && error.contains("relationship"),
                    "{error}"
                );
            }
        }
    }

    fn run(text: &str) -> serde_json::Value {
        serde_json::json!({
            "type": "text", "text": text, "bold": null, "italic": null,
            "underline": false, "strikethrough": false, "strikeDouble": false,
            "fontSize": null, "color": null, "fontFamily": null, "fieldType": null, "hyperlinkUsesTextFill": false
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
