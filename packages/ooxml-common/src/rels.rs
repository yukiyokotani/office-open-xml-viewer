//! OPC relationship (`.rels`) parsing and part-name resolution shared by the
//! docx, pptx and xlsx parsers.
//!
//! Every OOXML package resolves references through the Open Packaging
//! Conventions relationship grammar (ECMA-376 Part 2 §9.3, ISO/IEC 29500-2):
//! a `_rels/<part>.rels` file lists `<Relationship Id Type Target
//! TargetMode?>` entries, and a part references another by relationship id
//! (`r:id` / `r:embed`). The three parsers had three near-identical private
//! copies of "parse the rels map" and "resolve a Target against the source
//! part's directory". Sharing them keeps path resolution byte-identical across
//! formats — notably the `../` normalization that docx's private copy was
//! missing (it concatenated `base_dir + target` verbatim, leaving
//! `word/charts/../media/image.png` unresolved for chart / footnote media).
//!
//! Scope is deliberately "types + parse + pure predicate": this module computes
//! the *zip part name* a relationship points at. It does not read bytes, does
//! not extract media, and does not know about any host schema — each parser
//! keeps its own media pipeline and only borrows the resolution logic.

use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

/// ECMA-376 Part 2 §9.3.2 `TargetMode`: whether a relationship's `Target`
/// names a part *inside* the package (`Internal`, the default) or an external
/// resource such as a hyperlink URL (`External`). External targets are opaque
/// URLs and must never be run through part-name resolution.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TargetMode {
    /// `TargetMode="Internal"` (or omitted). `Target` is a package part name,
    /// relative to the source part's directory (or root-absolute with a leading
    /// `/`); resolve it with [`resolve_target`].
    Internal,
    /// `TargetMode="External"`. `Target` is an absolute URI (hyperlink, linked
    /// image, etc.) that is used verbatim — never resolved as a zip path.
    External,
}

/// A single parsed relationship: its `Target` string exactly as authored, plus
/// its [`TargetMode`]. The target is kept raw (unresolved) so callers can decide
/// whether to resolve it as a part name (Internal) or use it as a URL
/// (External); resolution against a base directory is a separate step
/// ([`resolve_target`]) because the base differs per source part.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RelTarget {
    /// The `Target` attribute verbatim (e.g. `../media/image1.png`,
    /// `/word/media/image1.png`, or `https://example.com/`).
    pub target: String,
    /// OPC relationship type URI. Retained so consumers such as DrawingML
    /// blip resolvers can reject an `rId` whose target is not an image part.
    pub relationship_type: Option<String>,
    /// Internal (package part) vs External (opaque URL).
    pub mode: TargetMode,
}

/// Parse a `.rels` XML document into a `rId → `[`RelTarget`] map.
///
/// Reads every `<Relationship>` child of the root `<Relationships>` element
/// that carries both an `Id` and a `Target`, recording its `TargetMode`
/// (`External` when the attribute equals `"External"` case-insensitively, else
/// `Internal` — the spec default). Malformed XML or an empty string yields an
/// empty map. The returned [`BTreeMap`] keeps ids in sorted order so any
/// serialized form is deterministic.
///
/// This does not resolve targets to part names — Targets are stored verbatim;
/// call [`resolve_target`] per source part for Internal entries.
pub fn parse_rels(xml: &str) -> BTreeMap<String, RelTarget> {
    let mut map = BTreeMap::new();
    if xml.is_empty() {
        return map;
    }
    let doc = match crate::depth::parse_guarded(xml) {
        Ok(d) => d,
        Err(_) => return map,
    };
    for rel in doc
        .root_element()
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        let (Some(id), Some(target)) = (rel.attribute("Id"), rel.attribute("Target")) else {
            continue;
        };
        // TargetMode is optional; only "External" (case-insensitive per the OPC
        // schema's xsd:string enumeration usage in the wild) diverts a target
        // away from part-name resolution.
        let mode = match rel.attribute("TargetMode") {
            Some(m) if m.eq_ignore_ascii_case("External") => TargetMode::External,
            _ => TargetMode::Internal,
        };
        map.insert(
            id.to_string(),
            RelTarget {
                target: target.to_string(),
                relationship_type: rel.attribute("Type").map(str::to_owned),
                mode,
            },
        );
    }
    map
}

/// Resolve an OPC relationship `Target` to a normalized zip part name.
///
/// Two cases, both anchored at the package root (ECMA-376 Part 2 §9.3 — part
/// names are root-relative, `/`-separated, with no `.`/`..` segments in their
/// canonical form):
///
/// - **Root-absolute** (`Target` starts with `/`, e.g. openpyxl's
///   `/xl/drawings/drawing1.xml` or `/word/media/image1.png`): resolved from the
///   package root, ignoring `base_dir`. The leading slash is dropped so the
///   result is a bare part name that matches a zip entry.
/// - **Relative** (`../media/image1.png`, `slide1.xml`): resolved against
///   `base_dir` — the *directory* of the source part (e.g. `xl/drawings` for
///   `xl/drawings/_rels/drawing1.xml.rels`). `base_dir` may carry a trailing
///   slash; empty segments are dropped so both `word/` and `xl/worksheets`
///   forms work.
///
/// `..` pops one segment and `.` / empty segments are skipped, yielding a
/// normalized name with no relative components — so `word/charts` +
/// `../media/x.png` becomes `word/media/x.png`, never the unresolved
/// `word/charts/../media/x.png`.
///
/// This directory-based form does not reject targets that name no package
/// part (a scheme, an authority, a query, or an invalid part name). Prefer
/// [`resolve_part_name`], which applies the complete OPC procedure against the
/// source part; DOCX and the shared chart-image index use it.
pub fn resolve_target(base_dir: &str, target: &str) -> String {
    let mut parts: Vec<&str> = if target.starts_with('/') {
        // Root-absolute part name: ignore base_dir entirely.
        Vec::new()
    } else {
        base_dir.split('/').filter(|s| !s.is_empty()).collect()
    };
    for seg in target.split('/') {
        match seg {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// Resolve an Internal relationship `Target` against its source part, returning
/// the ZIP part name it identifies, or `None` when the reference does not
/// identify a part of this package.
///
/// This is the normative OPC procedure rather than directory concatenation:
///
/// - ECMA-376 Part 2 §6.5.2.3 (and §6.5.2.2 for `/_rels/.rels`): for an
///   Internal relationship the base IRI is the pack IRI of the *source part*
///   (`source_part`, a ZIP-style name without the leading `/`; pass `""` for
///   package relationships, whose base is the package root).
/// - §6.4.1: the reference is resolved with RFC 3986 §5 unchanged. So
///   §5.2.2/§5.2.3 merge a relative path with the base path minus its last
///   segment, an absolute-path reference (`/word/x.xml`) replaces it, an empty
///   reference is the base itself, and §5.2.4 `remove_dot_segments` removes
///   `.`/`..` segments. A `..` above the root is dropped by that algorithm
///   (`/a/../../b` becomes `/b`), so a path reference always stays inside the
///   package.
/// - A reference with a scheme (RFC 3986 §3.1) or an authority (`//host`)
///   replaces the pack IRI's scheme/authority and therefore names a resource
///   outside this package. §6.5.3.4 requires an Internal target to be a
///   relative reference *to a part*, so such a target identifies no part.
/// - §6.3.3(f): the resolved path must be a valid part name (§6.2.2.2:
///   `1*( "/" isegment-nz )`, no segment ending in `.`). A query component is
///   not part of that grammar, and an empty or trailing-dot segment is not a
///   part name; each yields `None`. A fragment identifies a location inside
///   the resource and is removed before the part is named.
/// - RFC 3986 §6.2.2.1/§6.2.2.2 (the equivalence RFC 3987 §5.3.2.3 extends
///   to IRIs): a percent-encoded unreserved ASCII character (`ALPHA / DIGIT /
///   "-" / "." / "_" / "~"`) is equivalent to the character itself, so it is
///   decoded; any other percent-encoding keeps its octet with uppercase hex.
///   Part 2 §6.2.2.2 forbids *producers* from writing such encodings, but a
///   consumer resolving a reference applies the RFC equivalence. As in the
///   RFC 3986 §6.2.2 normalization ladder, percent-encoding normalization
///   precedes path-segment normalization, so `%2E%2E` is the dot segment `..`
///   (it still cannot leave the package root). A `%` not followed by two hex
///   digits is not an IRI and yields `None`, as does a percent-encoded `/` or
///   `\` (forbidden by §6.2.2.2; it is never decoded into a separator).
/// - §7.3.4: the ZIP item name is the part name without its leading `/`, with
///   every non-ASCII character percent-encoded (as UTF-8, uppercase hex).
///
/// ASCII case is kept as authored: part names are equivalent under ASCII
/// case-insensitive matching (§6.2.2.3), which the package lookup applies
/// through [`part_name_equivalence_key`]. The returned name may therefore
/// differ in case from the stored ZIP item name while naming the same part.
///
/// Library policy: `None` means "this relationship names no readable part".
/// Callers treat it exactly as a relationship whose target part is absent.
pub fn resolve_part_name(source_part: &str, target: &str) -> Option<String> {
    let reference = target.split_once('#').map_or(target, |(before, _)| before);
    if has_uri_scheme(reference) || reference.starts_with("//") || reference.contains('?') {
        return None;
    }
    let base = format!("/{}", source_part.trim_start_matches('/'));
    let merged = if reference.starts_with('/') {
        reference.to_owned()
    } else if reference.is_empty() {
        base
    } else {
        let directory_end = base.rfind('/').map_or(0, |index| index + 1);
        format!("{}{}", &base[..directory_end], reference)
    };
    let path = remove_dot_segments(&normalize_percent_encoding(&merged)?);
    let name = path.strip_prefix('/')?.to_owned();
    if name.is_empty()
        || name.split('/').any(|segment| {
            segment.is_empty()
                || segment.ends_with('.')
                || segment.contains("%2F")
                || segment.contains("%5C")
        })
    {
        return None;
    }
    Some(name)
}

/// ASCII unreserved characters of RFC 3986 §2.3.
fn is_unreserved_ascii(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~')
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

/// Percent-encoding normalization of one part name (see
/// [`resolve_part_name`]): decode percent-encoded unreserved ASCII, uppercase
/// the hex of every other triplet, percent-encode non-ASCII characters
/// (§7.3.4). `None` for a malformed `%` triplet.
fn normalize_percent_encoding(name: &str) -> Option<String> {
    let bytes = name.as_bytes();
    let mut out = String::with_capacity(name.len());
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'%' {
            let high = hex_value(*bytes.get(index + 1)?)?;
            let low = hex_value(*bytes.get(index + 2)?)?;
            let decoded = high * 16 + low;
            if is_unreserved_ascii(decoded) {
                out.push(char::from(decoded));
            } else {
                out.push_str(&format!("%{decoded:02X}"));
            }
            index += 3;
        } else if byte.is_ascii() {
            out.push(char::from(byte));
            index += 1;
        } else {
            out.push_str(&format!("%{byte:02X}"));
            index += 1;
        }
    }
    Some(out)
}

/// Equivalence key of a part name or ZIP item name (ECMA-376 Part 2 §6.2.2.3
/// with the RFC 3986 §6.2.2 percent-encoding equivalence): two names name the
/// same part exactly when their keys are equal.
///
/// The key decodes percent-encoded unreserved ASCII, percent-encodes non-ASCII
/// (§7.3.4 ZIP mapping), and folds ASCII case — which also makes the hex of
/// the remaining triplets case-insensitive, as RFC 3986 §6.2.2.1 requires. A
/// malformed `%` is kept literally, since a stored ZIP item name must still
/// have a key. Package validation rejects two items with the same key, so a
/// key lookup is unambiguous.
pub fn part_name_equivalence_key(name: &str) -> String {
    let bytes = name.as_bytes();
    let mut key = String::with_capacity(name.len());
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        let triplet = (byte == b'%')
            .then(|| {
                Some(hex_value(*bytes.get(index + 1)?)? * 16 + hex_value(*bytes.get(index + 2)?)?)
            })
            .flatten();
        match triplet {
            Some(decoded) if is_unreserved_ascii(decoded) => {
                key.push(char::from(decoded.to_ascii_lowercase()));
                index += 3;
            }
            Some(decoded) => {
                key.push_str(&format!("%{decoded:02x}"));
                index += 3;
            }
            None if byte.is_ascii() => {
                key.push(char::from(byte.to_ascii_lowercase()));
                index += 1;
            }
            None => {
                key.push_str(&format!("%{byte:02x}"));
                index += 1;
            }
        }
    }
    key
}

/// RFC 3986 §3.1 `scheme ":"` prefix: `ALPHA *( ALPHA / DIGIT / "+" / "-" /
/// "." )` before the first `:`, with no `/`, `?` or `#` in front of it. A
/// relative-path reference cannot carry a colon in its first segment (§4.2),
/// so a Windows drive path such as `C:\x` also parses as a scheme.
fn has_uri_scheme(reference: &str) -> bool {
    let Some(colon) = reference.find(':') else {
        return false;
    };
    let scheme = &reference[..colon];
    let mut bytes = scheme.bytes();
    bytes
        .next()
        .is_some_and(|first| first.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
}

/// RFC 3986 §5.2.4 `remove_dot_segments` for an absolute path.
fn remove_dot_segments(path: &str) -> String {
    let mut output: Vec<&str> = Vec::new();
    let segments: Vec<&str> = path.split('/').skip(1).collect();
    let last = segments.len().saturating_sub(1);
    for (index, segment) in segments.iter().enumerate() {
        match *segment {
            "." | ".." => {
                if *segment == ".." {
                    output.pop();
                }
                // A final `.`/`..` leaves the path ending in `/`.
                if index == last {
                    output.push("");
                }
            }
            other => output.push(other),
        }
    }
    format!("/{}", output.join("/"))
}

/// Derive the relationship part name belonging to an OPC source part.
///
/// ECMA-376 Part 2 §6.5.2.3 places the relationship part in an adjacent
/// `_rels` directory and appends `.rels` to the complete source filename:
/// `ppt/slides/slide1.xml` becomes
/// `ppt/slides/_rels/slide1.xml.rels`. A package-root source follows the same
/// rule (`document.xml` becomes `_rels/document.xml.rels`).
pub fn relationship_part_path(source_part_path: &str) -> String {
    let (dir, file) = source_part_path
        .rsplit_once('/')
        .map_or(("", source_part_path), |(dir, file)| (dir, file));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

impl RelTarget {
    /// Resolve this relationship to the package part it names, per
    /// [`resolve_part_name`]. External relationships never name a part
    /// (ECMA-376 Part 2 §6.5.3.4) and yield `None`.
    pub fn resolve_part(&self, source_part: &str) -> Option<String> {
        match self.mode {
            TargetMode::Internal => resolve_part_name(source_part, &self.target),
            TargetMode::External => None,
        }
    }

    /// Resolve this relationship's target against `base_dir`, honoring
    /// [`TargetMode`]: Internal targets are normalized to a part name via
    /// [`resolve_target`]; External targets (URLs) are returned verbatim.
    pub fn resolve(&self, base_dir: &str) -> String {
        match self.mode {
            TargetMode::Internal => resolve_target(base_dir, &self.target),
            TargetMode::External => self.target.clone(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_reads_id_target_and_mode() {
        let xml = r#"<?xml version="1.0"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="…/image" Target="../media/image1.png"/>
          <Relationship Id="rId2" Type="…/hyperlink" Target="https://example.com/" TargetMode="External"/>
        </Relationships>"#;
        let map = parse_rels(xml);
        assert_eq!(map.len(), 2);
        assert_eq!(
            map.get("rId1"),
            Some(&RelTarget {
                target: "../media/image1.png".to_string(),
                relationship_type: Some("…/image".to_string()),
                mode: TargetMode::Internal,
            })
        );
        assert_eq!(
            map.get("rId2"),
            Some(&RelTarget {
                target: "https://example.com/".to_string(),
                relationship_type: Some("…/hyperlink".to_string()),
                mode: TargetMode::External,
            })
        );
    }

    #[test]
    fn parse_defaults_missing_target_mode_to_internal() {
        let xml = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="slide1.xml"/>
        </Relationships>"#;
        assert_eq!(
            parse_rels(xml).get("rId1").unwrap().mode,
            TargetMode::Internal
        );
    }

    #[test]
    fn parse_empty_or_malformed_is_empty() {
        assert!(parse_rels("").is_empty());
        assert!(parse_rels("<not xml").is_empty());
    }

    #[test]
    fn resolve_relative_target_against_base_dir() {
        // The everyday case: a part references a sibling directory's media.
        assert_eq!(
            resolve_target("ppt/slides", "../media/image1.png"),
            "ppt/media/image1.png"
        );
        assert_eq!(
            resolve_target("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
    }

    #[test]
    fn relationship_part_path_for_normal_nested_source() {
        assert_eq!(
            relationship_part_path("ppt/slides/slide1.xml"),
            "ppt/slides/_rels/slide1.xml.rels"
        );
    }

    #[test]
    fn relationship_part_path_uses_the_actual_source_directory() {
        assert_eq!(
            relationship_part_path("custom/deck/slides/a.xml"),
            "custom/deck/slides/_rels/a.xml.rels"
        );
    }

    #[test]
    fn relationship_part_path_supports_package_root_sources() {
        assert_eq!(
            relationship_part_path("document.xml"),
            "_rels/document.xml.rels"
        );
    }

    #[test]
    fn resolve_absolute_leading_slash_ignores_base_dir() {
        // Root-absolute Targets (openpyxl style) resolve from the package root.
        assert_eq!(
            resolve_target("ppt/slides", "/ppt/charts/chart5.xml"),
            "ppt/charts/chart5.xml"
        );
        assert_eq!(
            resolve_target("xl/drawings", "/xl/charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
        // A trailing-slash base (docx's "word/" convention) is irrelevant for
        // absolute targets and must not leak in.
        assert_eq!(
            resolve_target("word/", "/word/media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn resolve_multi_level_dotdot_normalizes() {
        // Deeply nested relative targets fully normalize — this is exactly the
        // case docx's old `format!("{}{}", base_dir, target)` left unresolved.
        assert_eq!(
            resolve_target("word/charts", "../../media/deep.png"),
            "media/deep.png"
        );
        assert_eq!(
            resolve_target("word/charts", "../media/chart_img.png"),
            "word/media/chart_img.png"
        );
        // Trailing-slash base with a single `..` (docx "word/" convention).
        assert_eq!(
            resolve_target("word/", "../media/footnote.png"),
            "media/footnote.png"
        );
    }

    #[test]
    fn resolve_part_name_follows_rfc3986_against_the_source_part() {
        let cases = [
            // Plain, `./`, `../` and absolute-path references from the main part.
            (
                "word/document.xml",
                "footnotes.xml",
                Some("word/footnotes.xml"),
            ),
            (
                "word/document.xml",
                "./footnotes.xml",
                Some("word/footnotes.xml"),
            ),
            (
                "word/document.xml",
                "../word/footnotes.xml",
                Some("word/footnotes.xml"),
            ),
            (
                "word/document.xml",
                "/word/footnotes.xml",
                Some("word/footnotes.xml"),
            ),
            (
                "word/document.xml",
                "notes/./a/../footnotes.xml",
                Some("word/notes/footnotes.xml"),
            ),
            // The base is the source PART: its last segment is replaced.
            (
                "word/charts/chart1.xml",
                "../media/i.png",
                Some("word/media/i.png"),
            ),
            ("document.xml", "./media/i.png", Some("media/i.png")),
            ("", "word/document.xml", Some("word/document.xml")),
            // §5.2.4 drops `..` above the root; it cannot leave the package.
            ("word/document.xml", "../../../x.xml", Some("x.xml")),
            // Empty reference is the base; a fragment is not part of the name.
            ("word/document.xml", "", Some("word/document.xml")),
            (
                "word/document.xml",
                "footnotes.xml#n1",
                Some("word/footnotes.xml"),
            ),
            // Not a part of this package, or not a valid part name.
            ("word/document.xml", "https://example.com/x.xml", None),
            ("word/document.xml", "pack://x/word/a.xml", None),
            ("word/document.xml", "C:\\x.xml", None),
            ("word/document.xml", "//host/word/a.xml", None),
            ("word/document.xml", "a.xml?x=1", None),
            ("word/document.xml", "./", None),
            ("word/document.xml", "..", None),
            ("word/document.xml", "a//b.xml", None),
            ("word/document.xml", "name.", None),
            // RFC 3986 §6.2.2 percent-encoding normalization and §7.3.4.
            (
                "word/document.xml",
                "%66ootnotes.xml",
                Some("word/footnotes.xml"),
            ),
            ("word/document.xml", "%2e/%41.xml", Some("word/A.xml")),
            ("word/document.xml", "a%20b.xml", Some("word/a%20b.xml")),
            ("word/document.xml", "a%3cb.xml", Some("word/a%3Cb.xml")),
            (
                "word/document.xml",
                "caf\u{e9}.xml",
                Some("word/caf%C3%A9.xml"),
            ),
            ("word/document.xml", "%2E%2E/x.xml", Some("x.xml")),
            ("word/document.xml", "a%2Fb.xml", None),
            ("word/document.xml", "a%5cb.xml", None),
            ("word/document.xml", "a%zzb.xml", None),
            ("word/document.xml", "a%4", None),
            // Case is kept; the package lookup folds it (§6.2.2.3).
            (
                "word/document.xml",
                "../WORD/Footnotes.XML",
                Some("WORD/Footnotes.XML"),
            ),
        ];
        for (source, target, expected) in cases {
            assert_eq!(
                resolve_part_name(source, target).as_deref(),
                expected,
                "{source} + {target}"
            );
        }
    }

    #[test]
    fn equivalence_key_folds_case_and_unreserved_percent_encoding() {
        let key = part_name_equivalence_key;
        assert_eq!(key("word/Footnotes.XML"), key("WORD/footnotes.xml"));
        assert_eq!(key("word/%41.xml"), key("word/a.xml"));
        assert_eq!(key("word/%7e%5F.xml"), key("word/~_.xml"));
        assert_eq!(key("word/a%3Cb.xml"), key("word/a%3cb.xml"));
        assert_eq!(key("word/caf\u{e9}.xml"), key("word/caf%C3%A9.xml"));
        assert_ne!(key("word/a%20b.xml"), key("word/a b.xml"));
        assert_ne!(key("word/a%2Fb.xml"), key("word/a/b.xml"));
        assert_ne!(key("word/a.xml"), key("word/b.xml"));
        // A malformed triplet still has a (literal) key.
        assert_eq!(key("word/%zz"), "word/%zz");
    }

    #[test]
    fn resolve_part_never_resolves_external_targets() {
        let external = RelTarget {
            target: "footnotes.xml".to_string(),
            relationship_type: None,
            mode: TargetMode::External,
        };
        assert_eq!(external.resolve_part("word/document.xml"), None);
    }

    #[test]
    fn resolve_via_rel_target_honors_external() {
        let internal = RelTarget {
            target: "../media/image1.png".to_string(),
            relationship_type: Some("…/image".to_string()),
            mode: TargetMode::Internal,
        };
        assert_eq!(internal.resolve("ppt/slides"), "ppt/media/image1.png");
        let external = RelTarget {
            target: "https://example.com/x".to_string(),
            relationship_type: Some("…/hyperlink".to_string()),
            mode: TargetMode::External,
        };
        // External targets pass through untouched regardless of base_dir.
        assert_eq!(external.resolve("ppt/slides"), "https://example.com/x");
    }
}
