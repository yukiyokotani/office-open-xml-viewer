//! DrawingML theme parsing (`<a:theme>`) shared by the docx, pptx and xlsx
//! parsers.
//!
//! Every OOXML host embeds the same theme grammar (ECMA-376 §20.1.6 /
//! §14.2.7 / §20.1.4.1): a `<a:clrScheme>` of twelve named color slots, a
//! `<a:fontScheme>` with major/minor typefaces per script, and the four style
//! lists in `<a:fmtScheme>`. The three parsers had three partial, drifting
//! copies — pptx resolved `<a:prstClr>` preset names
//! while docx and xlsx silently dropped them; xlsx never read the font scheme;
//! docx never read line styles. Consolidating the *parse* here fixes the prstClr
//! gap uniformly (a preset color now resolves in all three formats) while each
//! parser keeps its own thin key-format adapter and color casing.
//!
//! Scope is "types + parse + pure predicate". This module reads the theme XML
//! into owned structs and stores each color slot's hex **exactly as authored**
//! (the `srgbClr@val` / `sysClr@lastClr` string verbatim, or a preset's
//! canonical hex). Case-folding, `#` prefixing, logical-name (`clrMap`)
//! resolution and the runtime color transforms (lumMod/tint/…) are NOT here —
//! they diverge per host and stay in each parser / the renderer.

use crate::ns::is_a_ns;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::sync::Arc;

/// One style-matrix entry retained for on-demand self-contained XML.
///
/// A theme entry may use namespace prefixes declared only on `<a:theme>`
/// (including extension namespaces such as `a14`). Keeping only the source
/// range of `<a:ln>`/`<a:solidFill>` therefore produces an invalid fragment.
/// This descriptor keeps the entry bytes and only the namespace prefixes the
/// fragment actually references; namespace URIs are shared across entries.
/// Consumers can therefore build a valid temporary wrapper without multiplying
/// every root namespace by every style-matrix entry.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StandaloneThemeStyleXml {
    fragment: String,
    namespaces: Vec<(Option<String>, Arc<str>)>,
}

impl StandaloneThemeStyleXml {
    fn from_node(
        node: roxmltree::Node<'_, '_>,
        source: &str,
        namespace_pool: &mut HashMap<String, Arc<str>>,
    ) -> Self {
        let fragment = source[node.range()].to_owned();
        let mut prefixes = BTreeSet::new();
        let bytes = fragment.as_bytes();
        for colon in bytes
            .iter()
            .enumerate()
            .filter_map(|(index, byte)| (*byte == b':').then_some(index))
        {
            let mut start = colon;
            while start > 0 && is_xml_name_byte(bytes[start - 1]) {
                start -= 1;
            }
            if start < colon {
                prefixes.insert(fragment[start..colon].to_owned());
            }
        }

        let intern = |uri: &str, pool: &mut HashMap<String, Arc<str>>| {
            pool.entry(uri.to_owned())
                .or_insert_with(|| Arc::<str>::from(uri))
                .clone()
        };
        let mut namespaces = Vec::with_capacity(prefixes.len() + 1);
        if let Some(uri) = node.lookup_namespace_uri(None) {
            namespaces.push((None, intern(uri, namespace_pool)));
        }
        for prefix in prefixes {
            if prefix == "xml" {
                continue;
            }
            if let Some(uri) = node.lookup_namespace_uri(Some(&prefix)) {
                namespaces.push((Some(prefix), intern(uri, namespace_pool)));
            }
        }
        Self {
            fragment,
            namespaces,
        }
    }

    /// Build a self-contained wrapper on demand. Namespace URIs are interned
    /// across all entries and only prefixes referenced by this fragment are
    /// retained, so parsing a theme cannot expand O(namespaces × entries) in
    /// memory. The temporary wrapper lives only for the caller's DOM parse.
    pub fn to_xml(&self) -> String {
        let mut xml = String::from("<themeStyleRoot");
        for (prefix, uri) in &self.namespaces {
            match prefix {
                Some(prefix) => {
                    xml.push_str(" xmlns:");
                    xml.push_str(prefix);
                }
                None => xml.push_str(" xmlns"),
            }
            xml.push_str("=\"");
            xml.push_str(&quick_xml::escape::escape(uri.as_ref()));
            xml.push('"');
        }
        xml.push('>');
        xml.push_str(&self.fragment);
        xml.push_str("</themeStyleRoot>");
        xml
    }
}

fn is_xml_name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')
}

/// Result of resolving a style-matrix index. `NoStyle` is an authored sentinel
/// (`0`, and `1000` for fill references), while `Missing` means the index asked
/// for a list entry that the theme does not contain. Keeping these distinct
/// prevents a corrupt/missing recipe from being mistaken for explicit no-fill.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StyleMatrixLookup<'a> {
    NoStyle,
    Missing,
    Entry(&'a StandaloneThemeStyleXml),
}

/// The four lists in DrawingML `CT_StyleMatrix` (`a:fmtScheme`).
///
/// Entries are stored as namespace-complete XML rather than partially decoded
/// fields. That preserves extension markup and lets the shared fill/line/effect
/// parsers evolve without reparsing or retaining the full theme document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ThemeFormatScheme {
    fill_styles: Vec<StandaloneThemeStyleXml>,
    line_styles: Vec<StandaloneThemeStyleXml>,
    effect_styles: Vec<StandaloneThemeStyleXml>,
    background_fill_styles: Vec<StandaloneThemeStyleXml>,
}

impl ThemeFormatScheme {
    /// Parse `<a:fmtScheme>` from either Transitional or Strict DrawingML.
    /// Malformed XML or a missing format scheme yields an empty value.
    pub fn parse(xml: &str) -> Self {
        let Ok(doc) = crate::depth::parse_guarded(xml) else {
            return Self::default();
        };
        let Some(format_scheme) = doc.descendants().find(|node| {
            node.is_element()
                && node.tag_name().name() == "fmtScheme"
                && is_a_ns(node.tag_name().namespace())
        }) else {
            return Self::default();
        };

        let mut namespace_pool = HashMap::new();
        let mut collect = |list_name: &str| -> Vec<StandaloneThemeStyleXml> {
            format_scheme
                .children()
                .find(|node| {
                    node.is_element()
                        && node.tag_name().name() == list_name
                        && is_a_ns(node.tag_name().namespace())
                })
                .into_iter()
                .flat_map(|list| list.children())
                .filter(|node| node.is_element() && is_a_ns(node.tag_name().namespace()))
                .map(|node| StandaloneThemeStyleXml::from_node(node, xml, &mut namespace_pool))
                .collect()
        };

        Self {
            fill_styles: collect("fillStyleLst"),
            line_styles: collect("lnStyleLst"),
            effect_styles: collect("effectStyleLst"),
            background_fill_styles: collect("bgFillStyleLst"),
        }
    }

    /// Resolve `CT_StyleMatrixReference@idx` for `fillRef`.
    pub fn lookup_fill_ref(&self, index: usize) -> StyleMatrixLookup<'_> {
        match index {
            0 | 1000 => StyleMatrixLookup::NoStyle,
            1..=999 => one_based_lookup(&self.fill_styles, index),
            _ => one_based_lookup(&self.background_fill_styles, index - 1000),
        }
    }

    /// Resolve `CT_StyleMatrixReference@idx` for `lnRef`.
    pub fn lookup_line_ref(&self, index: usize) -> StyleMatrixLookup<'_> {
        if index == 0 {
            StyleMatrixLookup::NoStyle
        } else {
            one_based_lookup(&self.line_styles, index)
        }
    }

    /// Resolve `CT_StyleMatrixReference@idx` for `effectRef`.
    pub fn lookup_effect_ref(&self, index: usize) -> StyleMatrixLookup<'_> {
        if index == 0 {
            StyleMatrixLookup::NoStyle
        } else {
            one_based_lookup(&self.effect_styles, index)
        }
    }
}

fn one_based_lookup(entries: &[StandaloneThemeStyleXml], index: usize) -> StyleMatrixLookup<'_> {
    entries
        .get(index.saturating_sub(1))
        .map(StyleMatrixLookup::Entry)
        .unwrap_or(StyleMatrixLookup::Missing)
}

/// The twelve `<a:clrScheme>` slot names in ECMA-376 §20.1.6.2 declaration
/// order: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`..`accent6`, `hlink`,
/// `folHlink`. Exposed so a consumer that needs positional order (xlsx indexes
/// its palette by slot ordinal) can iterate without hard-coding the list.
pub const CLR_SCHEME_SLOTS: [&str; 12] = [
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// The parsed `<a:clrScheme>`: slot name → hex string, stored **raw** (as
/// authored — no `#`, no case-folding). Keys are the twelve slot names present
/// in the theme; a slot whose color could not be read (unsupported child, or a
/// `prstClr` name outside [`preset_color`]) is simply absent.
///
/// A [`BTreeMap`] backs it so any serialized form is deterministic; lookups by
/// slot name are the common case ([`ThemeColorScheme::get`]).
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeColorScheme {
    colors: BTreeMap<String, String>,
}

impl ThemeColorScheme {
    /// Parse the `<a:clrScheme>` from a theme XML document. Each slot child holds
    /// one color element; `srgbClr` contributes its `val`, `sysClr` its
    /// `lastClr`, and `prstClr` its [`preset_color`] hex. The stored string is
    /// verbatim (the caller applies any casing / `#` prefix). Missing scheme or
    /// malformed XML yields an empty scheme.
    pub fn parse(xml: &str) -> Self {
        let mut colors = BTreeMap::new();
        let Ok(doc) = crate::depth::parse_guarded(xml) else {
            return Self { colors };
        };
        let Some(scheme) = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "clrScheme")
        else {
            return Self { colors };
        };
        for slot in scheme.children().filter(|n| n.is_element()) {
            let name = slot.tag_name().name().to_owned();
            for c in slot.children().filter(|n| n.is_element()) {
                let hex = color_node_hex(c);
                if let Some(h) = hex {
                    colors.insert(name, h);
                    break;
                }
            }
        }
        Self { colors }
    }

    /// Look up a slot's raw hex by name (`dk1`, `accent1`, `hlink`, …).
    pub fn get(&self, slot: &str) -> Option<&str> {
        self.colors.get(slot).map(String::as_str)
    }

    /// The twelve slots in spec order (`CLR_SCHEME_SLOTS`), each `Some(raw hex)`
    /// when present. Lets a positional consumer (xlsx builds a `Vec` indexed by
    /// ordinal) reconstruct its ordered palette while dropping absent slots as it
    /// sees fit.
    pub fn slots_in_order(&self) -> [Option<&str>; 12] {
        CLR_SCHEME_SLOTS.map(|slot| self.get(slot))
    }

    /// Iterate slot-name → raw-hex pairs (sorted by slot name). For a host that
    /// stores the palette in its own string-keyed map.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.colors.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// True when no slot color was parsed.
    pub fn is_empty(&self) -> bool {
        self.colors.is_empty()
    }
}

/// Resolve a single DrawingML color element to its raw hex, covering the three
/// forms a `<a:clrScheme>` slot uses: `srgbClr@val`, `sysClr@lastClr`, and
/// `prstClr@val` (via [`preset_color`]). Other elements (e.g. `scheme`-relative
/// colors, which never appear inside a theme's own scheme) yield `None`. The
/// returned string is verbatim — no casing or `#` is applied.
fn color_node_hex(node: roxmltree::Node<'_, '_>) -> Option<String> {
    match node.tag_name().name() {
        "srgbClr" => node.attribute("val").map(str::to_owned),
        "sysClr" => node.attribute("lastClr").map(str::to_owned),
        "prstClr" => preset_color(node.attribute("val").unwrap_or_default()),
        _ => None,
    }
}

/// The parsed `<a:fontScheme>`: major (heading) and minor (body) typefaces.
/// ECMA-376 `CT_FontCollection` has the latin/ea/cs axes followed by zero or
/// more `CT_SupplementalFont` entries keyed by script. Each host maps these
/// authored facts onto its own font-selection policy.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFonts {
    /// `<a:majorFont>` typefaces (latin / ea / cs).
    pub major: ThemeFontGroup,
    /// `<a:minorFont>` typefaces (latin / ea / cs).
    pub minor: ThemeFontGroup,
}

/// One `<a:majorFont>` or `<a:minorFont>` collection. Empty `typeface=""`
/// (common for `ea`/`cs`) is normalized to `None` or omitted from the
/// supplemental list. Supplemental entries retain source order so malformed
/// duplicate script mappings are not silently resolved by parse order.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeFontGroup {
    pub latin: Option<String>,
    pub ea: Option<String>,
    pub cs: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub supplemental: Vec<ThemeSupplementalFont>,
}

/// ECMA-376 `CT_SupplementalFont`: an authored script-to-typeface association.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ThemeSupplementalFont {
    pub script: String,
    pub typeface: String,
}

impl ThemeFontGroup {
    /// Return a unique supplemental face for `script`. A conflicting duplicate
    /// is ambiguous input and must not be converted into a font-selection rule.
    pub fn typeface_for_script(&self, script: &str) -> Option<&str> {
        let mut selected = None;
        for font in self
            .supplemental
            .iter()
            .filter(|font| font.script == script)
        {
            match selected {
                Some(previous) if previous != font.typeface => return None,
                Some(_) => {}
                None => selected = Some(font.typeface.as_str()),
            }
        }
        selected
    }
}

impl ThemeFonts {
    /// Parse the `<a:fontScheme>` collections. Missing scheme
    /// or malformed XML yields all-`None`.
    pub fn parse(xml: &str) -> Self {
        let Ok(doc) = crate::depth::parse_guarded(xml) else {
            return Self::default();
        };
        let Some(scheme) = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "fontScheme")
        else {
            return Self::default();
        };
        Self {
            major: parse_font_group(scheme, "majorFont"),
            minor: parse_font_group(scheme, "minorFont"),
        }
    }
}

/// Read one `<a:majorFont>` / `<a:minorFont>` child's latin/ea/cs typefaces.
fn parse_font_group(scheme: roxmltree::Node<'_, '_>, group_name: &str) -> ThemeFontGroup {
    let Some(group) = scheme
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == group_name)
    else {
        return ThemeFontGroup::default();
    };
    let read = |axis: &str| -> Option<String> {
        group
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == axis)
            .and_then(|n| n.attribute("typeface"))
            .filter(|t| !t.is_empty())
            .map(str::to_owned)
    };
    ThemeFontGroup {
        latin: read("latin"),
        ea: read("ea"),
        cs: read("cs"),
        supplemental: group
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "font")
            .filter_map(|n| {
                let script = n.attribute("script").filter(|s| !s.is_empty())?;
                let typeface = n.attribute("typeface").filter(|s| !s.is_empty())?;
                Some(ThemeSupplementalFont {
                    script: script.to_owned(),
                    typeface: typeface.to_owned(),
                })
            })
            .collect(),
    }
}

/// Resolve a DrawingML `<a:prstClr>` preset color name (ECMA-376 §20.1.10.48
/// ST_PresetColorVal) to its canonical uppercase hex. Covers the **complete**
/// 190-value enumeration (each `<xsd:enumeration>` of `ST_PresetColorVal` in
/// `dml-main.xsd`), with the RGB values taken verbatim from the §20.1.10.48
/// value table. Unrecognized names yield `None` (the caller leaves the slot /
/// fill unset). Single source of truth for the three parsers, which previously
/// handled only a subset (pptx) or none (docx/xlsx).
///
/// The 190 enum values collapse to 140 distinct hexes: the `dk*`/`lt*` Office
/// shorthands are *usually* exact aliases of the corresponding `dark*`/`light*`
/// CSS names, so they share an arm — but the spec table lists two deliberate
/// exceptions that must NOT be merged:
///
/// - `dkSeaGreen` = 8FBC8B vs `darkSeaGreen` = 8FBC8F
/// - `ltGoldenrodYellow` = FAFA78 vs `lightGoldenrodYellow` = FAFAD2
///
/// These are spec-authored differences (verified against the §20.1.10.48 table),
/// not typos, so they get their own arms.
pub fn preset_color(name: &str) -> Option<String> {
    let hex = match name {
        "aliceBlue" => "F0F8FF",
        "antiqueWhite" => "FAEBD7",
        "aqua" | "cyan" => "00FFFF",
        "aquamarine" => "7FFFD4",
        "azure" => "F0FFFF",
        "beige" => "F5F5DC",
        "bisque" => "FFE4C4",
        "black" => "000000",
        "blanchedAlmond" => "FFEBCD",
        "blue" => "0000FF",
        "blueViolet" => "8A2BE2",
        "brown" => "A52A2A",
        "burlyWood" => "DEB887",
        "cadetBlue" => "5F9EA0",
        "chartreuse" => "7FFF00",
        "chocolate" => "D2691E",
        "coral" => "FF7F50",
        "cornflowerBlue" => "6495ED",
        "cornsilk" => "FFF8DC",
        "crimson" => "DC143C",
        "darkBlue" | "dkBlue" => "00008B",
        "darkCyan" | "dkCyan" => "008B8B",
        "darkGoldenrod" | "dkGoldenrod" => "B8860B",
        "darkGray" | "darkGrey" | "dkGray" | "dkGrey" => "A9A9A9",
        "darkGreen" | "dkGreen" => "006400",
        "darkKhaki" | "dkKhaki" => "BDB76B",
        "darkMagenta" | "dkMagenta" => "8B008B",
        "darkOliveGreen" | "dkOliveGreen" => "556B2F",
        "darkOrange" | "dkOrange" => "FF8C00",
        "darkOrchid" | "dkOrchid" => "9932CC",
        "darkRed" | "dkRed" => "8B0000",
        "darkSalmon" | "dkSalmon" => "E9967A",
        "darkSeaGreen" => "8FBC8F",
        "dkSeaGreen" => "8FBC8B",
        "darkSlateBlue" | "dkSlateBlue" => "483D8B",
        "darkSlateGray" | "darkSlateGrey" | "dkSlateGray" | "dkSlateGrey" => "2F4F4F",
        "darkTurquoise" | "dkTurquoise" => "00CED1",
        "darkViolet" | "dkViolet" => "9400D3",
        "deepPink" => "FF1493",
        "deepSkyBlue" => "00BFFF",
        "dimGray" | "dimGrey" => "696969",
        "dodgerBlue" => "1E90FF",
        "firebrick" => "B22222",
        "floralWhite" => "FFFAF0",
        "forestGreen" => "228B22",
        "fuchsia" | "magenta" => "FF00FF",
        "gainsboro" => "DCDCDC",
        "ghostWhite" => "F8F8FF",
        "gold" => "FFD700",
        "goldenrod" => "DAA520",
        "gray" | "grey" => "808080",
        "green" => "008000",
        "greenYellow" => "ADFF2F",
        "honeydew" => "F0FFF0",
        "hotPink" => "FF69B4",
        "indianRed" => "CD5C5C",
        "indigo" => "4B0082",
        "ivory" => "FFFFF0",
        "khaki" => "F0E68C",
        "lavender" => "E6E6FA",
        "lavenderBlush" => "FFF0F5",
        "lawnGreen" => "7CFC00",
        "lemonChiffon" => "FFFACD",
        "lightBlue" | "ltBlue" => "ADD8E6",
        "lightCoral" | "ltCoral" => "F08080",
        "lightCyan" | "ltCyan" => "E0FFFF",
        "lightGoldenrodYellow" => "FAFAD2",
        "ltGoldenrodYellow" => "FAFA78",
        "lightGray" | "lightGrey" | "ltGray" | "ltGrey" => "D3D3D3",
        "lightGreen" | "ltGreen" => "90EE90",
        "lightPink" | "ltPink" => "FFB6C1",
        "lightSalmon" | "ltSalmon" => "FFA07A",
        "lightSeaGreen" | "ltSeaGreen" => "20B2AA",
        "lightSkyBlue" | "ltSkyBlue" => "87CEFA",
        "lightSlateGray" | "lightSlateGrey" | "ltSlateGray" | "ltSlateGrey" => "778899",
        "lightSteelBlue" | "ltSteelBlue" => "B0C4DE",
        "lightYellow" | "ltYellow" => "FFFFE0",
        "lime" => "00FF00",
        "limeGreen" => "32CD32",
        "linen" => "FAF0E6",
        "maroon" => "800000",
        "medAquamarine" | "mediumAquamarine" => "66CDAA",
        "medBlue" | "mediumBlue" => "0000CD",
        "medOrchid" | "mediumOrchid" => "BA55D3",
        "medPurple" | "mediumPurple" => "9370DB",
        "medSeaGreen" | "mediumSeaGreen" => "3CB371",
        "medSlateBlue" | "mediumSlateBlue" => "7B68EE",
        "medSpringGreen" | "mediumSpringGreen" => "00FA9A",
        "medTurquoise" | "mediumTurquoise" => "48D1CC",
        "medVioletRed" | "mediumVioletRed" => "C71585",
        "midnightBlue" => "191970",
        "mintCream" => "F5FFFA",
        "mistyRose" => "FFE4E1",
        "moccasin" => "FFE4B5",
        "navajoWhite" => "FFDEAD",
        "navy" => "000080",
        "oldLace" => "FDF5E6",
        "olive" => "808000",
        "oliveDrab" => "6B8E23",
        "orange" => "FFA500",
        "orangeRed" => "FF4500",
        "orchid" => "DA70D6",
        "paleGoldenrod" => "EEE8AA",
        "paleGreen" => "98FB98",
        "paleTurquoise" => "AFEEEE",
        "paleVioletRed" => "DB7093",
        "papayaWhip" => "FFEFD5",
        "peachPuff" => "FFDAB9",
        "peru" => "CD853F",
        "pink" => "FFC0CB",
        "plum" => "DDA0DD",
        "powderBlue" => "B0E0E6",
        "purple" => "800080",
        "red" => "FF0000",
        "rosyBrown" => "BC8F8F",
        "royalBlue" => "4169E1",
        "saddleBrown" => "8B4513",
        "salmon" => "FA8072",
        "sandyBrown" => "F4A460",
        "seaGreen" => "2E8B57",
        "seaShell" => "FFF5EE",
        "sienna" => "A0522D",
        "silver" => "C0C0C0",
        "skyBlue" => "87CEEB",
        "slateBlue" => "6A5ACD",
        "slateGray" | "slateGrey" => "708090",
        "snow" => "FFFAFA",
        "springGreen" => "00FF7F",
        "steelBlue" => "4682B4",
        "tan" => "D2B48C",
        "teal" => "008080",
        "thistle" => "D8BFD8",
        "tomato" => "FF6347",
        "turquoise" => "40E0D0",
        "violet" => "EE82EE",
        "wheat" => "F5DEB3",
        "white" => "FFFFFF",
        "whiteSmoke" => "F5F5F5",
        "yellow" => "FFFF00",
        "yellowGreen" => "9ACD32",
        _ => return None,
    };
    Some(hex.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// All 190 `ST_PresetColorVal` enum values (in `dml-main.xsd` order). Pins
    /// that `preset_color` resolves the complete enumeration — a drift guard
    /// against the XSD.
    const PRESET_ENUM_NAMES: [&str; 190] = [
        "aliceBlue",
        "antiqueWhite",
        "aqua",
        "aquamarine",
        "azure",
        "beige",
        "bisque",
        "black",
        "blanchedAlmond",
        "blue",
        "blueViolet",
        "brown",
        "burlyWood",
        "cadetBlue",
        "chartreuse",
        "chocolate",
        "coral",
        "cornflowerBlue",
        "cornsilk",
        "crimson",
        "cyan",
        "darkBlue",
        "darkCyan",
        "darkGoldenrod",
        "darkGray",
        "darkGrey",
        "darkGreen",
        "darkKhaki",
        "darkMagenta",
        "darkOliveGreen",
        "darkOrange",
        "darkOrchid",
        "darkRed",
        "darkSalmon",
        "darkSeaGreen",
        "darkSlateBlue",
        "darkSlateGray",
        "darkSlateGrey",
        "darkTurquoise",
        "darkViolet",
        "dkBlue",
        "dkCyan",
        "dkGoldenrod",
        "dkGray",
        "dkGrey",
        "dkGreen",
        "dkKhaki",
        "dkMagenta",
        "dkOliveGreen",
        "dkOrange",
        "dkOrchid",
        "dkRed",
        "dkSalmon",
        "dkSeaGreen",
        "dkSlateBlue",
        "dkSlateGray",
        "dkSlateGrey",
        "dkTurquoise",
        "dkViolet",
        "deepPink",
        "deepSkyBlue",
        "dimGray",
        "dimGrey",
        "dodgerBlue",
        "firebrick",
        "floralWhite",
        "forestGreen",
        "fuchsia",
        "gainsboro",
        "ghostWhite",
        "gold",
        "goldenrod",
        "gray",
        "grey",
        "green",
        "greenYellow",
        "honeydew",
        "hotPink",
        "indianRed",
        "indigo",
        "ivory",
        "khaki",
        "lavender",
        "lavenderBlush",
        "lawnGreen",
        "lemonChiffon",
        "lightBlue",
        "lightCoral",
        "lightCyan",
        "lightGoldenrodYellow",
        "lightGray",
        "lightGrey",
        "lightGreen",
        "lightPink",
        "lightSalmon",
        "lightSeaGreen",
        "lightSkyBlue",
        "lightSlateGray",
        "lightSlateGrey",
        "lightSteelBlue",
        "lightYellow",
        "ltBlue",
        "ltCoral",
        "ltCyan",
        "ltGoldenrodYellow",
        "ltGray",
        "ltGrey",
        "ltGreen",
        "ltPink",
        "ltSalmon",
        "ltSeaGreen",
        "ltSkyBlue",
        "ltSlateGray",
        "ltSlateGrey",
        "ltSteelBlue",
        "ltYellow",
        "lime",
        "limeGreen",
        "linen",
        "magenta",
        "maroon",
        "medAquamarine",
        "medBlue",
        "medOrchid",
        "medPurple",
        "medSeaGreen",
        "medSlateBlue",
        "medSpringGreen",
        "medTurquoise",
        "medVioletRed",
        "mediumAquamarine",
        "mediumBlue",
        "mediumOrchid",
        "mediumPurple",
        "mediumSeaGreen",
        "mediumSlateBlue",
        "mediumSpringGreen",
        "mediumTurquoise",
        "mediumVioletRed",
        "midnightBlue",
        "mintCream",
        "mistyRose",
        "moccasin",
        "navajoWhite",
        "navy",
        "oldLace",
        "olive",
        "oliveDrab",
        "orange",
        "orangeRed",
        "orchid",
        "paleGoldenrod",
        "paleGreen",
        "paleTurquoise",
        "paleVioletRed",
        "papayaWhip",
        "peachPuff",
        "peru",
        "pink",
        "plum",
        "powderBlue",
        "purple",
        "red",
        "rosyBrown",
        "royalBlue",
        "saddleBrown",
        "salmon",
        "sandyBrown",
        "seaGreen",
        "seaShell",
        "sienna",
        "silver",
        "skyBlue",
        "slateBlue",
        "slateGray",
        "slateGrey",
        "snow",
        "springGreen",
        "steelBlue",
        "tan",
        "teal",
        "thistle",
        "tomato",
        "turquoise",
        "violet",
        "wheat",
        "white",
        "whiteSmoke",
        "yellow",
        "yellowGreen",
    ];

    const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
          <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
          <a:dk2><a:srgbClr val="44546A"/></a:dk2>
          <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
          <a:accent2><a:prstClr val="orange"/></a:accent2>
          <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
          <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
          <a:majorFont>
            <a:latin typeface="Aptos Display"/>
            <a:ea typeface="Yu Gothic"/>
            <a:cs typeface=""/>
          </a:majorFont>
          <a:minorFont>
            <a:latin typeface="Aptos"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
          </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
          <a:lnStyleLst>
            <a:ln w="6350"/>
            <a:ln w="12700"/>
            <a:ln/>
          </a:lnStyleLst>
        </a:fmtScheme>
      </a:themeElements>
    </a:theme>"#;

    #[test]
    fn clr_scheme_reads_srgb_sys_and_preset_raw() {
        let s = ThemeColorScheme::parse(THEME);
        // sysClr uses lastClr, srgbClr uses val — both raw (as authored).
        assert_eq!(s.get("dk1"), Some("000000"));
        assert_eq!(s.get("lt1"), Some("FFFFFF"));
        assert_eq!(s.get("dk2"), Some("44546A"));
        assert_eq!(s.get("accent1"), Some("4472C4"));
        // prstClr resolves through preset_color (uniform across formats).
        assert_eq!(s.get("accent2"), Some("FFA500"));
        assert_eq!(s.get("hlink"), Some("0563C1"));
        // Slots not present in the XML are absent (not empty strings).
        assert_eq!(s.get("accent3"), None);
    }

    #[test]
    fn clr_scheme_slots_in_order_matches_spec_positions() {
        let s = ThemeColorScheme::parse(THEME);
        let ordered = s.slots_in_order();
        assert_eq!(ordered[0], Some("000000")); // dk1
        assert_eq!(ordered[1], Some("FFFFFF")); // lt1
        assert_eq!(ordered[4], Some("4472C4")); // accent1
        assert_eq!(ordered[5], Some("FFA500")); // accent2 (preset)
        assert_eq!(ordered[6], None); // accent3 absent
        assert_eq!(ordered[10], Some("0563C1")); // hlink
    }

    #[test]
    fn font_scheme_reads_axes_and_drops_empty() {
        let f = ThemeFonts::parse(THEME);
        assert_eq!(f.major.latin.as_deref(), Some("Aptos Display"));
        assert_eq!(f.major.ea.as_deref(), Some("Yu Gothic"));
        // Empty typeface="" normalizes to None (not Some("")).
        assert_eq!(f.major.cs, None);
        assert_eq!(f.minor.latin.as_deref(), Some("Aptos"));
        assert_eq!(f.minor.ea, None);
        assert_eq!(f.minor.cs, None);
    }

    #[test]
    fn font_scheme_preserves_supplemental_scripts_and_rejects_ambiguous_duplicates() {
        // ECMA-376 CT_FontCollection permits repeated CT_SupplementalFont
        // children. Preserve their order while avoiding an arbitrary winner.
        let xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements><a:fontScheme name="Test">
            <a:majorFont><a:latin typeface="Heading"/><a:ea typeface=""/><a:cs typeface=""/>
              <a:font script="Jpan" typeface="Yu Gothic"/>
              <a:font script="Arab" typeface="Arabic Face"/>
              <a:font script="Jpan" typeface="Yu Gothic"/>
              <a:font script="" typeface="Ignored"/>
              <a:font script="Hani" typeface=""/>
            </a:majorFont>
            <a:minorFont><a:latin typeface="Body"/><a:ea typeface=""/><a:cs typeface=""/>
              <a:font script="Jpan" typeface="Yu Gothic"/>
              <a:font script="Jpan" typeface="Meiryo"/>
            </a:minorFont>
          </a:fontScheme></a:themeElements>
        </a:theme>"#;
        let fonts = ThemeFonts::parse(xml);
        assert_eq!(fonts.major.supplemental.len(), 3);
        assert_eq!(fonts.major.supplemental[0].script, "Jpan");
        assert_eq!(fonts.major.typeface_for_script("Jpan"), Some("Yu Gothic"));
        assert_eq!(fonts.major.typeface_for_script("Arab"), Some("Arabic Face"));
        assert_eq!(fonts.major.typeface_for_script("Hani"), None);
        assert_eq!(fonts.minor.supplemental.len(), 2);
        assert_eq!(fonts.minor.typeface_for_script("Jpan"), None);
        assert_eq!(fonts.minor.typeface_for_script("Arab"), None);
        let round_trip: ThemeFonts =
            serde_json::from_str(&serde_json::to_string(&fonts).unwrap()).unwrap();
        assert_eq!(round_trip, fonts);
    }

    #[test]
    fn format_scheme_preserves_in_scope_namespaces_and_lookup_semantics() {
        const XML: &str = r#"<a:theme
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
          <a:themeElements>
            <a:fmtScheme name="Office">
              <a:fillStyleLst><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:fillStyleLst>
              <a:lnStyleLst>
                <a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:extLst><a:ext uri="x"><a14:hiddenEffects/></a:ext></a:extLst></a:ln>
              </a:lnStyleLst>
              <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
              <a:bgFillStyleLst><a:noFill/></a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>"#;

        let scheme = ThemeFormatScheme::parse(XML);
        assert_eq!(scheme.lookup_fill_ref(0), StyleMatrixLookup::NoStyle);
        assert!(matches!(
            scheme.lookup_fill_ref(1),
            StyleMatrixLookup::Entry(_)
        ));
        assert_eq!(scheme.lookup_fill_ref(2), StyleMatrixLookup::Missing);
        assert_eq!(scheme.lookup_fill_ref(1000), StyleMatrixLookup::NoStyle);
        assert!(matches!(
            scheme.lookup_fill_ref(1001),
            StyleMatrixLookup::Entry(_)
        ));
        assert_eq!(scheme.lookup_line_ref(0), StyleMatrixLookup::NoStyle);
        assert_eq!(scheme.lookup_line_ref(2), StyleMatrixLookup::Missing);
        let StyleMatrixLookup::Entry(line) = scheme.lookup_line_ref(1) else {
            panic!("line style 1 should exist");
        };
        let line_xml = line.to_xml();
        let parsed =
            roxmltree::Document::parse(&line_xml).expect("standalone line style must parse");
        assert!(parsed
            .descendants()
            .any(|node| node.tag_name().name() == "hiddenEffects"
                && node.tag_name().namespace()
                    == Some("http://schemas.microsoft.com/office/drawing/2010/main")));
    }

    #[test]
    fn format_scheme_accepts_strict_drawingml_namespace() {
        const XML: &str = r#"<a:theme xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
          <a:themeElements><a:fmtScheme name="Strict">
            <a:lnStyleLst><a:ln w="25400"><a:noFill/></a:ln></a:lnStyleLst>
          </a:fmtScheme></a:themeElements>
        </a:theme>"#;
        let scheme = ThemeFormatScheme::parse(XML);
        let StyleMatrixLookup::Entry(line) = scheme.lookup_line_ref(1) else {
            panic!("strict line style should exist");
        };
        assert!(line
            .to_xml()
            .contains("http://purl.oclc.org/ooxml/drawingml/main"));
    }

    #[test]
    fn format_scheme_does_not_copy_unused_root_namespaces_into_every_entry() {
        let unused = (0..128)
            .map(|index| {
                format!(
                    r#" xmlns:u{index}="urn:unused:{index}:{}""#,
                    "x".repeat(128)
                )
            })
            .collect::<String>();
        let xml = format!(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"{unused}><a:themeElements><a:fmtScheme name="bounded"><a:lnStyleLst><a:ln/><a:ln/></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>"#
        );
        let scheme = ThemeFormatScheme::parse(&xml);
        for index in [1, 2] {
            let StyleMatrixLookup::Entry(entry) = scheme.lookup_line_ref(index) else {
                panic!("line style should exist");
            };
            let standalone = entry.to_xml();
            assert!(standalone.contains("xmlns:a="));
            assert!(!standalone.contains("urn:unused"));
            roxmltree::Document::parse(&standalone).expect("bounded wrapper parses");
        }
    }

    #[test]
    fn preset_color_covers_named_presets() {
        assert_eq!(preset_color("black").as_deref(), Some("000000"));
        assert_eq!(preset_color("white").as_deref(), Some("FFFFFF"));
        assert_eq!(preset_color("gray").as_deref(), Some("808080"));
        assert_eq!(preset_color("grey").as_deref(), Some("808080"));
        assert_eq!(preset_color("orange").as_deref(), Some("FFA500"));
        // Previously unrecognized names now resolve (full 190-value table).
        assert_eq!(preset_color("chartreuse").as_deref(), Some("7FFF00"));
        assert_eq!(preset_color("cornflowerBlue").as_deref(), Some("6495ED"));
        assert_eq!(preset_color("rebeccaPurple"), None); // truly absent name
    }

    /// The full §20.1.10.48 table resolves all 190 enum values, and the
    /// dk*/lt* Office shorthands alias their dark*/light* CSS names *except*
    /// for two spec-authored exceptions.
    #[test]
    fn preset_color_full_table_and_alias_exceptions() {
        // dk*/lt* usual aliases.
        assert_eq!(preset_color("dkBlue"), preset_color("darkBlue"));
        assert_eq!(preset_color("ltGray"), preset_color("lightGray"));
        assert_eq!(preset_color("medBlue"), preset_color("mediumBlue"));
        // Two deliberate spec exceptions — must NOT be equal.
        assert_eq!(preset_color("darkSeaGreen").as_deref(), Some("8FBC8F"));
        assert_eq!(preset_color("dkSeaGreen").as_deref(), Some("8FBC8B"));
        assert_ne!(preset_color("dkSeaGreen"), preset_color("darkSeaGreen"));
        assert_eq!(
            preset_color("lightGoldenrodYellow").as_deref(),
            Some("FAFAD2")
        );
        assert_eq!(preset_color("ltGoldenrodYellow").as_deref(), Some("FAFA78"));
        assert_ne!(
            preset_color("ltGoldenrodYellow"),
            preset_color("lightGoldenrodYellow")
        );
        // Every one of the 190 enum values resolves to some hex.
        for name in PRESET_ENUM_NAMES {
            assert!(
                preset_color(name).is_some(),
                "preset_color({name}) returned None"
            );
        }
        assert_eq!(PRESET_ENUM_NAMES.len(), 190);
    }

    #[test]
    fn empty_and_malformed_yield_empty() {
        assert!(ThemeColorScheme::parse("").is_empty());
        assert!(ThemeColorScheme::parse("<not xml").is_empty());
        assert_eq!(ThemeFonts::parse(""), ThemeFonts::default());
        assert_eq!(ThemeFormatScheme::parse(""), ThemeFormatScheme::default());
    }
}
