//! A flat ODF style store for the spreadsheet and chart readers (ODF Part 3
//! §16.2): every `style:style` and `style:default-style` of any family, its
//! parent, its data style and its property groups, attribute by attribute.
//! Lookups cascade through the parent chain and then the family's default
//! style, per attribute (§16.2).
//!
//! The text reader keeps its own admission-checking store (`style.rs`);
//! this one admits everything, because the spreadsheet and chart readers
//! render every property they understand and ignore the rest.

use crate::xml::{self, NS_STYLE};
use std::collections::HashMap;

/// Implementation bound on a parent chain (the style itself included): a
/// longer chain rejects the document and a cycle is a structural error
/// (`StyleStore::check_chains`).
const MAX_PARENT_DEPTH: usize = 64;
/// The canonical prefix of a namespace, for property keys.
pub fn prefix(namespace: Option<&str>) -> &'static str {
    match namespace {
        Some(crate::xml::NS_FO) => "fo",
        Some(crate::xml::NS_STYLE) => "style",
        Some(crate::xml::NS_TABLE) => "table",
        Some(crate::xml::NS_DRAW) => "draw",
        Some(crate::xml::NS_SVG) => "svg",
        Some(crate::xml::NS_LOEXT) => "loext",
        Some(crate::xml::NS_TEXT) => "text",
        Some(NS_CHART) => "chart",
        Some(NS_CHARTOOO) => "chartooo",
        Some(NS_CALCEXT) => "calcext",
        Some(crate::xml::NS_XLINK) => "xlink",
        Some("urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0") => "number",
        _ => "other",
    }
}

pub const NS_CHART: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
/// OpenFormula (ODF 1.3 Part 4), the default namespace of `style:map`
/// conditions (§19.472).
pub const NS_OF: &str = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
pub const NS_CHARTOOO: &str = "http://openoffice.org/2010/chart";
pub const NS_CALCEXT: &str = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

/// A property group's short name in keys: `cell`, `par`, `text`, ….
fn group_name(node: roxmltree::Node<'_, '_>) -> Option<&'static str> {
    let ns = node.tag_name().namespace();
    Some(match (ns, node.tag_name().name()) {
        (Some(NS_STYLE), "table-cell-properties") => "cell",
        (Some(NS_STYLE), "paragraph-properties") => "par",
        (Some(NS_STYLE), "text-properties") => "text",
        (Some(NS_STYLE), "table-column-properties") => "col",
        (Some(NS_STYLE), "table-row-properties") => "row",
        (Some(NS_STYLE), "table-properties") => "table",
        (Some(NS_STYLE), "graphic-properties") => "graphic",
        (Some(NS_STYLE), "chart-properties") => "chart",
        (Some(NS_STYLE), "drawing-page-properties") => "page",
        (Some(crate::xml::NS_LOEXT), "graphic-properties") => "graphic",
        _ => return None,
    })
}

/// A `style:map` (§16.3).
#[derive(Clone, Debug)]
pub struct StyleMap {
    pub condition: String,
    pub apply_style_name: String,
    pub base_cell_address: Option<String>,
}

#[derive(Clone, Debug, Default)]
pub struct Style {
    pub name: String,
    pub family: String,
    pub parent: Option<String>,
    pub data_style: Option<String>,
    /// `style:percentage-data-style-name` (§19.511), of chart styles.
    pub percentage_data_style: Option<String>,
    pub display_name: Option<String>,
    /// `style:master-page-name` and the other style-element attributes.
    pub attributes: HashMap<String, String>,
    /// `group:prefix:local` → value.
    pub properties: HashMap<String, String>,
    /// Element-valued properties by `group:prefix:local`: the element's
    /// attributes (`style:background-image`, `chart:label-separator` text).
    pub elements: HashMap<String, HashMap<String, String>>,
    pub maps: Vec<StyleMap>,
}

/// A style of the store, by position: a cheap identity that a reader
/// resolves once from a `*:style-name` and keeps (a column's or row's
/// default cell style), so nothing hashes a style name per cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct StyleId(u32);

/// Styles of one family by name.
type ByName = HashMap<String, StyleId>;

#[derive(Default)]
pub struct StyleStore {
    all: Vec<Style>,
    /// Family → name → style; lookups borrow both keys.
    common: HashMap<String, ByName>,
    automatic: HashMap<String, ByName>,
    defaults: HashMap<String, StyleId>,
}

impl StyleStore {
    /// Read `style:style` / `style:default-style` children of `parent`
    /// (`office:styles`, or `office:automatic-styles` with `automatic`).
    pub fn add(&mut self, parent: roxmltree::Node<'_, '_>, automatic: bool) -> Result<(), String> {
        for node in parent.children().filter(|n| n.is_element()) {
            let default = xml::is(node, NS_STYLE, "default-style");
            if !default && !xml::is(node, NS_STYLE, "style") {
                continue;
            }
            crate::budget::tick(1)?;
            let style = read_style(node)?;
            let id = StyleId(
                u32::try_from(self.all.len()).map_err(|_| crate::error::resource("style count"))?,
            );
            if default {
                self.defaults.insert(style.family.clone(), id);
            } else {
                let families = if automatic {
                    &mut self.automatic
                } else {
                    &mut self.common
                };
                match families.get_mut(style.family.as_str()) {
                    Some(names) => {
                        names.insert(style.name.clone(), id);
                    }
                    None => {
                        families.insert(
                            style.family.clone(),
                            HashMap::from([(style.name.clone(), id)]),
                        );
                    }
                }
            }
            self.all.push(style);
        }
        Ok(())
    }

    fn common(&self, family: &str, name: &str) -> Option<&Style> {
        let id = self.common.get(family)?.get(name)?;
        self.all.get(id.0 as usize)
    }

    fn parent_of(&self, style: &Style) -> Option<&Style> {
        self.common(&style.family, style.parent.as_deref()?)
    }

    /// Check every parent chain once all styles are read: a chain longer
    /// than [`MAX_PARENT_DEPTH`] rejects the document (a resource limit), a
    /// cycle is a structural error. [`chain`](Self::chain) then never stops
    /// early on a valid document.
    pub fn check_chains(&self) -> Result<(), String> {
        // Chain length (the style itself included) of each common ancestor.
        let mut depth_of: HashMap<*const Style, usize> = HashMap::new();
        for start in &self.all {
            crate::budget::tick(1)?;
            let mut path: Vec<&Style> = Vec::new();
            let mut base = 0usize;
            let mut next = self.parent_of(start);
            while let Some(parent) = next {
                if let Some(depth) = depth_of.get(&(parent as *const Style)) {
                    base = *depth;
                    break;
                }
                if path.iter().any(|s| std::ptr::eq(*s, parent)) {
                    return Err(crate::error::invalid("a style inherits from itself"));
                }
                path.push(parent);
                if path.len() + 1 > MAX_PARENT_DEPTH {
                    return Err(crate::error::resource("style inheritance depth"));
                }
                next = self.parent_of(parent);
            }
            if base + path.len() + 1 > MAX_PARENT_DEPTH {
                return Err(crate::error::resource("style inheritance depth"));
            }
            for (index, member) in path.iter().enumerate() {
                depth_of.insert(*member as *const Style, base + path.len() - index);
            }
        }
        Ok(())
    }

    /// The identity of the style a `*:style-name` names: automatic first,
    /// then common.
    pub fn id(&self, family: &str, name: &str) -> Option<StyleId> {
        self.automatic
            .get(family)
            .and_then(|names| names.get(name))
            .or_else(|| self.common.get(family).and_then(|names| names.get(name)))
            .copied()
    }

    pub fn style(&self, id: StyleId) -> &Style {
        &self.all[id.0 as usize]
    }

    /// The style a `*:style-name` names: automatic first, then common.
    pub fn find(&self, family: &str, name: &str) -> Option<&Style> {
        self.id(family, name).map(|id| self.style(id))
    }

    pub fn default_style(&self, family: &str) -> Option<&Style> {
        self.defaults.get(family).map(|id| self.style(*id))
    }

    /// `style` with its parents (common styles), nearest first, then the
    /// family's default style: the search order of a property (§16.2).
    pub fn chain_of<'s>(&'s self, family: &str, style: Option<&'s Style>) -> Vec<&'s Style> {
        let mut chain: Vec<&Style> = Vec::new();
        let mut current = style;
        while let Some(style) = current {
            if chain.len() >= MAX_PARENT_DEPTH || chain.iter().any(|s| std::ptr::eq(*s, style)) {
                break;
            }
            chain.push(style);
            current = self.parent_of(style);
        }
        if let Some(default) = self.default_style(family) {
            chain.push(default);
        }
        chain
    }

    /// The style and its parents (common styles), nearest first, without
    /// the family's default style: a context nested in another (a span in a
    /// paragraph in a cell) inherits what its own chain leaves unset.
    pub fn chain_own<'s>(&'s self, family: &str, name: &str) -> Vec<&'s Style> {
        self.chain_own_of(family, self.find(family, name))
    }

    /// [`chain_own`](Self::chain_own) of a resolved style.
    pub fn chain_own_of<'s>(&'s self, family: &str, style: Option<&'s Style>) -> Vec<&'s Style> {
        let mut chain = self.chain_of(family, style);
        if self.default_style(family).is_some() {
            chain.pop();
        }
        chain
    }

    /// The style and its parents (common styles), nearest first, then the
    /// family's default style.
    pub fn chain<'s>(&'s self, family: &str, name: Option<&str>) -> Vec<&'s Style> {
        self.chain_of(family, name.and_then(|n| self.find(family, n)))
    }
}

/// Nearest value of `key` in a resolved chain.
pub fn get<'a>(chain: &[&'a Style], key: &str) -> Option<&'a str> {
    chain
        .iter()
        .find_map(|style| style.properties.get(key).map(String::as_str))
}

pub fn get_element<'a>(chain: &[&'a Style], key: &str) -> Option<&'a HashMap<String, String>> {
    chain.iter().find_map(|style| style.elements.get(key))
}

fn qualified(node: &roxmltree::Attribute<'_, '_>) -> String {
    format!("{}:{}", prefix(node.namespace()), node.name())
}

fn read_style(node: roxmltree::Node<'_, '_>) -> Result<Style, String> {
    let mut style = Style {
        name: node.attribute((NS_STYLE, "name")).unwrap_or("").to_string(),
        family: node
            .attribute((NS_STYLE, "family"))
            .unwrap_or("")
            .to_string(),
        parent: node
            .attribute((NS_STYLE, "parent-style-name"))
            .map(str::to_string),
        data_style: node
            .attribute((NS_STYLE, "data-style-name"))
            .map(str::to_string),
        percentage_data_style: node
            .attribute((NS_STYLE, "percentage-data-style-name"))
            .map(str::to_string),
        display_name: node
            .attribute((NS_STYLE, "display-name"))
            .map(str::to_string),
        ..Style::default()
    };
    for attribute in node.attributes() {
        style
            .attributes
            .insert(qualified(&attribute), attribute.value().to_string());
    }
    for child in node.children().filter(|n| n.is_element()) {
        crate::budget::tick(1)?;
        if xml::is(child, NS_STYLE, "map") {
            if let (Some(condition), Some(apply)) = (
                child
                    .attribute((NS_STYLE, "condition"))
                    .and_then(|c| of_condition(child, c)),
                child.attribute((NS_STYLE, "apply-style-name")),
            ) {
                style.maps.push(StyleMap {
                    condition: condition.to_string(),
                    apply_style_name: apply.to_string(),
                    base_cell_address: child
                        .attribute((NS_STYLE, "base-cell-address"))
                        .map(str::to_string),
                });
            }
            continue;
        }
        let Some(group) = group_name(child) else {
            continue;
        };
        for attribute in child.attributes() {
            style.properties.insert(
                format!("{group}:{}", qualified(&attribute)),
                attribute.value().to_string(),
            );
        }
        for element in child.children().filter(|n| n.is_element()) {
            crate::budget::tick(1)?;
            let mut attributes: HashMap<String, String> = element
                .attributes()
                .map(|a| (qualified(&a), a.value().to_string()))
                .collect();
            let text: String = element
                .descendants()
                .filter(|n| n.is_text())
                .filter_map(|n| n.text())
                .collect();
            if !text.is_empty() {
                attributes.insert("#text".into(), text);
            }
            style.elements.insert(
                format!(
                    "{group}:{}:{}",
                    prefix(element.tag_name().namespace()),
                    element.tag_name().name()
                ),
                attributes,
            );
        }
    }
    Ok(style)
}

/// A length in points (§18.3.18): `cm`, `mm`, `in`, `pt`, `pc`, `px`.
pub fn length_pt(value: &str) -> Option<f64> {
    let value = value.trim();
    let split = value
        .find(|c: char| c.is_ascii_alphabetic() || c == '%')
        .unwrap_or(value.len());
    let number: f64 = value[..split].trim().parse().ok()?;
    if !number.is_finite() {
        return None;
    }
    let factor = match &value[split..] {
        "pt" => 1.0,
        "cm" => 72.0 / 2.54,
        "mm" => 72.0 / 25.4,
        "in" | "inch" => 72.0,
        "pc" => 12.0,
        "px" => 0.75,
        "" if number == 0.0 => 1.0,
        _ => return None,
    };
    Some(number * factor)
}

/// `#rrggbb` → `RRGGBB`; `None` for `transparent` or a malformed value.
pub fn hex_color(value: &str) -> Option<String> {
    let hex = value.trim().strip_prefix('#')?;
    if hex.len() == 6 && hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        Some(hex.to_ascii_uppercase())
    } else if hex.len() == 3 && hex.bytes().all(|b| b.is_ascii_hexdigit()) {
        Some(
            hex.chars()
                .flat_map(|c| [c, c])
                .collect::<String>()
                .to_ascii_uppercase(),
        )
    } else {
        None
    }
}

/// A `percent` (§18.3.23) as a fraction: `"50%"` → 0.5; `None` unless
/// finite.
pub fn percent(value: &str) -> Option<f64> {
    let number: f64 = value.trim().strip_suffix('%')?.trim().parse().ok()?;
    number.is_finite().then_some(number / 100.0)
}

/// `svg:font-family` without the CSS quotes (§19.528).
pub fn unquote_family(value: &str) -> String {
    let value = value.trim();
    let first = value.split(',').next().unwrap_or(value).trim();
    first.trim_matches(|c| c == '\'' || c == '"').to_string()
}

/// `style:font-face` names → family (§16.23).
pub fn font_faces(
    root: roxmltree::Node<'_, '_>,
    faces: &mut HashMap<String, String>,
) -> Result<(), String> {
    for decls in root
        .children()
        .filter(|n| xml::is(*n, crate::xml::NS_OFFICE, "font-face-decls"))
    {
        for face in decls
            .children()
            .filter(|n| xml::is(*n, NS_STYLE, "font-face"))
        {
            crate::budget::tick(1)?;
            if let Some(name) = face.attribute((NS_STYLE, "name")) {
                let family = face
                    .attribute((crate::xml::NS_SVG, "font-family"))
                    .map(unquote_family)
                    .unwrap_or_else(|| name.to_string());
                faces.insert(name.to_string(), family);
            }
        }
    }
    Ok(())
}

/// A `style:map` condition (§19.472) in the OpenFormula namespace, without
/// its namespace prefix. The prefix is resolved through the bindings in
/// scope where the map is written; a missing prefix defaults to
/// OpenFormula. A condition in any other (or an unbound) namespace is not
/// recognized: `None`, and its map is ignored, as §19.472 requires.
pub fn of_condition<'a>(node: roxmltree::Node<'_, '_>, value: &'a str) -> Option<&'a str> {
    let value = value.trim();
    match value.split_once(':') {
        Some((prefix, rest)) if is_ncname(prefix) => {
            (node.lookup_namespace_uri(Some(prefix)) == Some(NS_OF)).then_some(rest.trim())
        }
        _ => Some(value),
    }
}

/// An XML NCName (a namespace prefix): a letter or `_`, then letters,
/// digits, `-`, `_` or `.`.
fn is_ncname(text: &str) -> bool {
    let mut chars = text.chars();
    chars.next().is_some_and(|c| c.is_alphabetic() || c == '_')
        && chars.all(|c| c.is_alphanumeric() || matches!(c, '-' | '_' | '.'))
}

// ---------------------------------------------------------------------------
// Values the text, spreadsheet and presentation readers read alike. Each was
// settled once, against LibreOffice 26.8's PDF export of synthetic documents
// that vary one parameter at a time (Writer, Calc, Impress and Draw), so the
// three readers cannot drift apart again.
// ---------------------------------------------------------------------------

/// The width in points of a border-width keyword in `fo:border` (§20.183,
/// XSL 1.1 §7.8.21 `border-*-width`), or `None` for a token that is not one.
///
/// XSL leaves the keyword widths to the implementation (only thin ≤ medium ≤
/// thick). These are LibreOffice's named widths, measured in a Writer
/// paragraph, a Calc cell and an Impress table cell alike: a keyword border
/// exports the same stroke as the explicit width (`thin` as `0.05pt`, its
/// hairline; `middle` as `1pt`; `thick` as `2.5pt`), and LibreOffice saves
/// them back as `0.06pt`, `0.99pt` and `2.49pt` (its internal units,
/// rounded). So `thin` is 0.05 pt (one twip), `middle` 1 pt and `thick`
/// 2.5 pt. LibreOffice does not know XSL's `medium` (it spells the
/// middle width `middle`) and drops a border that names it; `medium` is a valid
/// XSL keyword, so it is read as LibreOffice's middle width rather than
/// dropped. (An Impress table then strokes 1 pt and 2.5 pt borders at 1.75 pt
/// and 4.4 pt, explicit widths and keywords alike: that is its table
/// painter, not the keyword.)
pub fn border_width_keyword(token: &str) -> Option<f64> {
    match token {
        "thin" => Some(0.05),
        "medium" | "middle" => Some(1.0),
        "thick" => Some(2.5),
        _ => None,
    }
}

/// The shear factor of `skewX (a)` / `skewY (a)` in `draw:transform`
/// (§19.228): `skewX (a)` maps (x, y) to (x + k·y, y) and `skewY (a)` maps it
/// to (x, y + k·x), in page coordinates (y down), with `a` in radians.
///
/// The specification only names the transforms. LibreOffice shears against
/// the SVG sense: a 4 cm × 2 cm rectangle written `skewX (0.5)` keeps its top
/// edge and moves its bottom edge 2 cm · tan 0.5 to the left, `skewX (-0.5)`
/// to the right, and `skewY (0.5)` moves its right edge up (Draw and Calc,
/// measured identically). So k = −tan a.
pub fn skew_shear(angle: f64) -> f64 {
    -angle.tan()
}

/// One unsigned number of the enhanced-geometry grammar (§19.171
/// `draw:formula`: `integer`, `fractional-constant exponent?` or `integer
/// exponent`, the exponent `('e' | 'E') sign? integer`), read from `at`;
/// `None` (with `at` unmoved) when no number starts there. An `e` not followed
/// by an exponent's digits is not consumed.
///
/// The same grammar is read in `draw:enhanced-path` parameters. LibreOffice
/// reads exponents of either case in path parameters, and an upper-case `E`
/// in formulas, but treats a formula spelling a lower-case `e` (`4e2`) as
/// invalid (the formula evaluates to 0). The readers follow the grammar,
/// which allows both.
pub fn scan_number(bytes: &[u8], at: &mut usize) -> Option<f64> {
    let digits = |from: usize| {
        bytes[from.min(bytes.len())..]
            .iter()
            .take_while(|b| b.is_ascii_digit())
            .count()
    };
    let start = *at;
    let mut end = start + digits(start);
    let mut any = end > start;
    if bytes.get(end) == Some(&b'.') {
        let fraction = digits(end + 1);
        if any || fraction > 0 {
            end += 1 + fraction;
            any = true;
        }
    }
    if !any {
        return None;
    }
    if matches!(bytes.get(end), Some(b'e' | b'E')) {
        let mut exponent = end + 1;
        if matches!(bytes.get(exponent), Some(b'+' | b'-')) {
            exponent += 1;
        }
        let count = digits(exponent);
        if count > 0 {
            end = exponent + count;
        }
    }
    let value = std::str::from_utf8(&bytes[start..end])
        .ok()?
        .parse::<f64>()
        .ok()
        .filter(|v| v.is_finite())?;
    *at = end;
    Some(value)
}

/// The parametric start angle and sweep (radians, clockwise on the page,
/// y down) of an enhanced-path `G` arc (ODF 1.4 Part 3 Table 10
/// `arcangleto`: radii `wR` `hR`, start angle `st` and swing `sw` in degrees;
/// the ellipse is placed so that the current point is the arc's start).
///
/// `st` and `sw` are angles at the ellipse's centre (`sw` is "the central
/// angle"), not parameters: the start is where the ray at `st` meets the
/// ellipse, parameter atan2(wR·sin st, hR·cos st). Measured in LibreOffice:
/// `G 400 200 -45 90` from (100, 500) ends 357.8 units below its start, the
/// ray intersections (the parameter reading gives 282.8); `sw = -90` turns
/// counter-clockwise, and `sw = 450` draws the whole ellipse and then the
/// arc of `sw = 90` again. A positive `sw` turns clockwise. The angles are
/// central in the path's own units: `G 400 400 -45 90` in a square view box
/// stretched over a 2:1 shape ends where the circle's arc ends, so `rx` and
/// `ry` are the radii before the view box is scaled to the shape.
///
/// The parametric sweep is the one within half a turn of `sw`: the parameter
/// and the central angle of a point differ by less than a quarter turn, so
/// that choice is unique and keeps whole extra turns: every turn is drawn,
/// so a dashed outline keeps its dash phase. Only past [`MAX_ARC_TURNS`]
/// whole turns (an implementation bound on the curves one arc draws) are
/// pairs of turns taken off; that keeps the outline and the even-odd fill
/// (`N`) parity, but not the dash phase of such a stroke.
pub fn arc_angle_to(rx: f64, ry: f64, start_deg: f64, swing_deg: f64) -> (f64, f64) {
    // Whole turns past the bound come off in pairs (two turns, 720°, keep
    // the parity), in degrees and before any trigonometry: the remainder of
    // a float is exact, so the end angle survives any swing. The start is
    // taken modulo a turn likewise.
    let limit = (MAX_ARC_TURNS + 1.0) * 360.0;
    let swing_deg = if swing_deg.abs() > limit {
        let rest = swing_deg % 720.0;
        rest + swing_deg.signum() * ((limit - rest.abs()) / 720.0).floor() * 720.0
    } else {
        swing_deg
    };
    let start_deg = start_deg % 360.0;
    let parameter = |angle: f64| (rx * angle.sin()).atan2(ry * angle.cos());
    let (start, swing) = (start_deg.to_radians(), swing_deg.to_radians());
    let t0 = parameter(start);
    let delta = parameter(start + swing) - t0;
    let tau = std::f64::consts::TAU;
    (t0, delta + ((swing - delta) / tau).round() * tau)
}

/// Implementation bound: the whole turns one `G` arc draws (its sweep stays
/// within `MAX_ARC_TURNS + 1` turns and half a turn). The readers size their
/// arc flattening for it.
pub const MAX_ARC_TURNS: f64 = 16.0;

/// The natural size in points of a Windows metafile, which `fo:clip`
/// (§20.172) lengths are measured on: an EMF's header frame (`rclFrame`,
/// hundredths of a millimetre; [MS-EMF] 2.2.9) or a placeable WMF's bounding
/// box in its units per inch ([MS-WMF] 2.3.2.3 META_PLACEABLE). `None` for a
/// WMF without the placeable header (it states no size) or anything else.
///
/// Measured in LibreOffice (Writer): a picture whose frame is 1 in × 0.5 in
/// (EMF `rclFrame`, with device bounds of a different size) or whose
/// placeable bounding box is 1440 × 720 at 1440 per inch, clipped by
/// `rect(0in, 0.25in, 0in, 0in)` in a 2 in frame, shows its left three
/// quarters: the clip is a quarter of a 1 in natural width.
pub fn metafile_size_pt(bytes: &[u8]) -> Option<(f64, f64)> {
    let i32_at = |at: usize| -> Option<f64> {
        Some(f64::from(i32::from_le_bytes(
            bytes.get(at..at + 4)?.try_into().ok()?,
        )))
    };
    let i16_at = |at: usize| -> Option<f64> {
        Some(f64::from(i16::from_le_bytes(
            bytes.get(at..at + 2)?.try_into().ok()?,
        )))
    };
    let size = if bytes.get(0..4) == Some(&[1, 0, 0, 0]) && bytes.get(40..44) == Some(b" EMF") {
        let (l, t, r, b) = (i32_at(24)?, i32_at(28)?, i32_at(32)?, i32_at(36)?);
        ((r - l) * 72.0 / 2540.0, (b - t) * 72.0 / 2540.0)
    } else if bytes.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
        let (l, t, r, b) = (i16_at(6)?, i16_at(8)?, i16_at(10)?, i16_at(12)?);
        let inch = f64::from(u16::from_le_bytes(bytes.get(14..16)?.try_into().ok()?));
        if inch <= 0.0 {
            return None;
        }
        ((r - l).abs() * 72.0 / inch, (b - t).abs() * 72.0 / inch)
    } else {
        return None;
    };
    (size.0 > 0.0 && size.1 > 0.0).then_some(size)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lengths_and_colors() {
        assert_eq!(length_pt("1in"), Some(72.0));
        assert!((length_pt("2.54cm").unwrap() - 72.0).abs() < 1e-9);
        assert_eq!(length_pt("0"), Some(0.0));
        assert_eq!(length_pt("1em"), None);
        assert_eq!(hex_color("#a1b2c3").as_deref(), Some("A1B2C3"));
        assert_eq!(hex_color("transparent"), None);
        assert_eq!(
            unquote_family("'Liberation Sans', Arial"),
            "Liberation Sans"
        );
    }

    #[test]
    fn cascade_follows_parents_then_the_default_style() {
        let xml = format!(
            r##"<r xmlns:style="{NS_STYLE}" xmlns:fo="{}"><style:default-style style:family="table-cell"><style:text-properties fo:font-size="10pt" fo:color="#000000"/></style:default-style><style:style style:name="Default" style:family="table-cell"><style:text-properties fo:font-size="11pt"/></style:style><style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style></r>"##,
            crate::xml::NS_FO
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        let mut store = StyleStore::default();
        store.add(document.root_element(), false).unwrap();
        store.check_chains().unwrap();
        let chain = store.chain("table-cell", Some("ce1"));
        assert_eq!(get(&chain, "text:fo:font-weight"), Some("bold"));
        assert_eq!(get(&chain, "text:fo:font-size"), Some("11pt"));
        assert_eq!(get(&chain, "text:fo:color"), Some("#000000"));
    }

    #[test]
    fn numbers_follow_the_formula_grammar() {
        let read = |text: &str| {
            let mut at = 0;
            scan_number(text.as_bytes(), &mut at).map(|v| (v, at))
        };
        assert_eq!(read("4e2"), Some((400.0, 3)));
        assert_eq!(read("4E-1*"), Some((0.4, 4)));
        assert_eq!(read("1.5e+1"), Some((15.0, 6)));
        assert_eq!(read(".5"), Some((0.5, 2)));
        assert_eq!(read("3."), Some((3.0, 2)));
        // An `e` without exponent digits ends the number.
        assert_eq!(read("2e"), Some((2.0, 1)));
        assert_eq!(read("2e-x"), Some((2.0, 1)));
        assert_eq!(read("."), None);
        assert_eq!(read("-1"), None);
        assert_eq!(read("1e999"), None);
    }

    #[test]
    fn arc_angle_to_reads_central_angles() {
        let degrees = |(t0, sweep): (f64, f64)| (t0.to_degrees(), sweep.to_degrees());
        let close = |(a, b): (f64, f64), (x, y): (f64, f64)| {
            assert!(
                (a - x).abs() < 1e-9 && (b - y).abs() < 1e-9,
                "{a} {b} vs {x} {y}"
            );
        };
        let t = 63.434_948_822_922_01; // atan(2): the ray at 45° on a 2:1 ellipse.
        close(
            degrees(arc_angle_to(400.0, 200.0, -45.0, 90.0)),
            (-t, 2.0 * t),
        );
        close(
            degrees(arc_angle_to(400.0, 200.0, -45.0, -90.0)),
            (-t, -(180.0 - 2.0 * t)),
        );
        close(
            degrees(arc_angle_to(400.0, 200.0, -45.0, 450.0)),
            (-t, 360.0 + 2.0 * t),
        );
        close(
            degrees(arc_angle_to(400.0, 200.0, -45.0, 360.0)),
            (-t, 360.0),
        );
        // Whole turns are kept up to the bound; past it, pairs come off.
        close(
            degrees(arc_angle_to(400.0, 200.0, -45.0, 2610.0)),
            (-t, 2520.0 + 2.0 * t),
        );
        close(
            degrees(arc_angle_to(400.0, 200.0, -45.0, -1440.0)),
            (-t, -1440.0),
        );
        close(
            degrees(arc_angle_to(5.0, 5.0, 0.0, 16.0 * 360.0 + 90.0)),
            (0.0, 16.0 * 360.0 + 90.0),
        );
        close(
            degrees(arc_angle_to(5.0, 5.0, 0.0, 17.0 * 360.0 + 90.0)),
            (0.0, 15.0 * 360.0 + 90.0),
        );
        close(
            degrees(arc_angle_to(5.0, 5.0, 0.0, -(18.0 * 360.0 + 90.0))),
            (0.0, -(16.0 * 360.0 + 90.0)),
        );
        // A huge swing keeps its end angle (1e20 ≡ 280° modulo 360°, and an
        // odd number of turns modulo 720°) and its turn parity.
        let (t0, sweep) = degrees(arc_angle_to(5.0, 5.0, 0.0, 1e20));
        assert_eq!((t0, 1e20f64 % 720.0), (0.0, 640.0));
        assert!((sweep - (15.0 * 360.0 + 280.0)).abs() < 1e-6, "{sweep}");
        // A huge start is taken modulo a turn.
        close(degrees(arc_angle_to(5.0, 5.0, 1e20, 90.0)), (-80.0, 90.0));
        assert_eq!(arc_angle_to(400.0, 200.0, 30.0, 0.0).1, 0.0);
        // Across the atan2 branch cut (170° + 800° = 970° ≡ -110°): two
        // whole turns and 80° stay, in either direction.
        for (start, swing, kept) in [
            (170.0, 800.0, 800.0),
            (-170.0, -800.0, -800.0),
            (170.0, 1160.0, 1160.0),
            (170.0, -1160.0, -1160.0),
        ] {
            close(
                degrees(arc_angle_to(400.0, 400.0, start, swing)),
                (start, kept),
            );
        }
        // A circle's parameter is its central angle.
        close(
            degrees(arc_angle_to(5.0, 5.0, 10.0, -200.0)),
            (10.0, -200.0),
        );
    }

    #[test]
    fn metafile_natural_sizes() {
        let mut emf = vec![0u8; 88];
        emf[0] = 1;
        emf[40..44].copy_from_slice(b" EMF");
        for (at, value) in [(24, 100i32), (28, 0), (32, 2640), (36, 1270)] {
            emf[at..at + 4].copy_from_slice(&value.to_le_bytes());
        }
        assert_eq!(metafile_size_pt(&emf), Some((72.0, 36.0)));
        let mut wmf = vec![0xD7, 0xCD, 0xC6, 0x9A, 0, 0];
        for value in [0i16, 0, 1440, 720] {
            wmf.extend_from_slice(&value.to_le_bytes());
        }
        wmf.extend_from_slice(&1440u16.to_le_bytes());
        assert_eq!(metafile_size_pt(&wmf), Some((72.0, 36.0)));
        // A standard WMF header states no size.
        assert_eq!(metafile_size_pt(&[1, 0, 9, 0, 0, 3]), None);
    }

    #[test]
    fn shared_keyword_widths_and_skew() {
        assert_eq!(border_width_keyword("thin"), Some(0.05));
        assert_eq!(border_width_keyword("medium"), Some(1.0));
        assert_eq!(border_width_keyword("middle"), Some(1.0));
        assert_eq!(border_width_keyword("thick"), Some(2.5));
        assert_eq!(border_width_keyword("solid"), None);
        assert!((skew_shear(0.5) + 0.5f64.tan()).abs() < 1e-12);
    }
}
