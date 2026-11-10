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
/// Implementation bound on stored style data (owned strings plus a fixed
/// per-entry overhead), charged before a style is stored.
const MAX_STYLE_BYTES: usize = 64 * 1024 * 1024;
const STORED_ENTRY_OVERHEAD: usize = 48;

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

#[derive(Default)]
pub struct StyleStore {
    stored_bytes: usize,
    styles: HashMap<(String, String), Style>,
    automatic: HashMap<(String, String), Style>,
    defaults: HashMap<String, Style>,
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
            self.charge(node)?;
            let style = read_style(node);
            if default {
                self.defaults.insert(style.family.clone(), style);
            } else if automatic {
                self.automatic
                    .insert((style.family.clone(), style.name.clone()), style);
            } else {
                self.styles
                    .insert((style.family.clone(), style.name.clone()), style);
            }
        }
        Ok(())
    }

    fn charge(&mut self, node: roxmltree::Node<'_, '_>) -> Result<(), String> {
        let bytes = node
            .descendants()
            .filter(|n| n.is_element())
            .fold(0usize, |sum, e| {
                e.attributes()
                    .fold(sum.saturating_add(STORED_ENTRY_OVERHEAD), |sum, a| {
                        sum.saturating_add(a.name().len() + a.value().len() + STORED_ENTRY_OVERHEAD)
                    })
            });
        self.stored_bytes = self
            .stored_bytes
            .checked_add(bytes)
            .filter(|total| *total <= MAX_STYLE_BYTES)
            .ok_or_else(|| "OpenDocument styles exceed the style storage limit".to_string())?;
        Ok(())
    }

    /// Check every parent chain once all styles are read: a chain longer
    /// than [`MAX_PARENT_DEPTH`] rejects the document (a resource limit), a
    /// cycle is a structural error. [`chain`](Self::chain) then never stops
    /// early on a valid document.
    pub fn check_chains(&self) -> Result<(), String> {
        let parent_of = |style: &Style| {
            style
                .parent
                .as_deref()
                .and_then(|parent| self.styles.get(&(style.family.clone(), parent.to_string())))
        };
        // Chain length (the style itself included) of each common ancestor.
        let mut depth_of: HashMap<(&str, &str), usize> = HashMap::new();
        for start in self.automatic.values().chain(self.styles.values()) {
            let mut path: Vec<&Style> = Vec::new();
            let mut base = 0usize;
            let mut next = parent_of(start);
            while let Some(parent) = next {
                if let Some(depth) = depth_of.get(&(parent.family.as_str(), parent.name.as_str())) {
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
                next = parent_of(parent);
            }
            if base + path.len() + 1 > MAX_PARENT_DEPTH {
                return Err(crate::error::resource("style inheritance depth"));
            }
            for (index, member) in path.iter().enumerate() {
                depth_of.insert(
                    (member.family.as_str(), member.name.as_str()),
                    base + path.len() - index,
                );
            }
        }
        Ok(())
    }

    /// The style a `*:style-name` names: automatic first, then common.
    pub fn find(&self, family: &str, name: &str) -> Option<&Style> {
        let key = (family.to_string(), name.to_string());
        self.automatic.get(&key).or_else(|| self.styles.get(&key))
    }

    pub fn default_style(&self, family: &str) -> Option<&Style> {
        self.defaults.get(family)
    }

    /// The style and its parents (common styles), nearest first, then the
    /// family's default style.
    pub fn chain<'s>(&'s self, family: &str, name: Option<&str>) -> Vec<&'s Style> {
        let mut chain: Vec<&Style> = Vec::new();
        let mut current = name.and_then(|n| self.find(family, n));
        while let Some(style) = current {
            if chain.len() >= MAX_PARENT_DEPTH || chain.iter().any(|s| std::ptr::eq(*s, style)) {
                break;
            }
            chain.push(style);
            current = style
                .parent
                .as_deref()
                .and_then(|parent| self.styles.get(&(family.to_string(), parent.to_string())));
        }
        if let Some(default) = self.defaults.get(family) {
            chain.push(default);
        }
        chain
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

fn read_style(node: roxmltree::Node<'_, '_>) -> Style {
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
        if xml::is(child, NS_STYLE, "map") {
            if let (Some(condition), Some(apply)) = (
                child.attribute((NS_STYLE, "condition")),
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
    style
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

/// A percentage `"50%"` as a fraction.
pub fn percent(value: &str) -> Option<f64> {
    value
        .trim()
        .strip_suffix('%')?
        .trim()
        .parse::<f64>()
        .ok()
        .map(|v| v / 100.0)
}

/// `svg:font-family` without the CSS quotes (§19.528).
pub fn unquote_family(value: &str) -> String {
    let value = value.trim();
    let first = value.split(',').next().unwrap_or(value).trim();
    first.trim_matches(|c| c == '\'' || c == '"').to_string()
}

/// `style:font-face` names → family (§16.23).
pub fn font_faces(root: roxmltree::Node<'_, '_>, faces: &mut HashMap<String, String>) {
    for decls in root
        .children()
        .filter(|n| xml::is(*n, crate::xml::NS_OFFICE, "font-face-decls"))
    {
        for face in decls
            .children()
            .filter(|n| xml::is(*n, NS_STYLE, "font-face"))
        {
            if let Some(name) = face.attribute((NS_STYLE, "name")) {
                let family = face
                    .attribute((crate::xml::NS_SVG, "font-family"))
                    .map(unquote_family)
                    .unwrap_or_else(|| name.to_string());
                faces.insert(name.to_string(), family);
            }
        }
    }
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
        let chain = store.chain("table-cell", Some("ce1"));
        assert_eq!(get(&chain, "text:fo:font-weight"), Some("bold"));
        assert_eq!(get(&chain, "text:fo:font-size"), Some("11pt"));
        assert_eq!(get(&chain, "text:fo:color"), Some("#000000"));
    }
}
