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
}
