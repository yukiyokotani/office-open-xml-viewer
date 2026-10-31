//! Styles and their cascade (ODF Part 3 §16.2, §16.4).
//!
//! Every property group (`style:paragraph-properties`, `style:text-properties`,
//! …) is stored attribute by attribute; inheritance is per attribute (§16.2),
//! except element-valued properties such as `style:tab-stops`, which inherit
//! as a whole element (§17.7) and are stored with their parent group.

use crate::error::invalid;
use crate::xml::{self, NS_STYLE, NS_TEXT};
use std::collections::{HashMap, HashSet};

/// Implementation bound on a parent chain; ODF requires none, a cycle is a
/// container error.
const MAX_PARENT_DEPTH: usize = 64;
/// Implementation bound on stored style data: every owned string (expanded
/// attribute and element names included) plus a fixed per-entry overhead,
/// charged before the style is read. Expanded names repeat a namespace URI
/// declared once, so storage is not linear in the XML size without it.
const MAX_STYLE_BYTES: usize = 64 * 1024 * 1024;
const STORED_ENTRY_OVERHEAD: usize = 64;

/// The style families readers look up (§19.480 `style:family`). Styles of
/// every family are stored under their declared family name.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Family {
    Paragraph,
    Text,
}

impl Family {
    fn name(self) -> &'static str {
        match self {
            Family::Paragraph => "paragraph",
            Family::Text => "text",
        }
    }
}

/// A property group's local name (`paragraph-properties`, …).
pub type Group = &'static str;
pub const PARAGRAPH: Group = "paragraph-properties";
pub const TEXT: Group = "text-properties";
/// LibreOffice's `loext:graphic-properties` (area fill of paragraphs).
pub const LOEXT_GRAPHIC: Group = "loext:graphic-properties";

/// An attribute key: (namespace URI, local name).
pub type Key = (String, String);

#[derive(Debug, Default, Clone)]
pub struct Properties {
    pub attributes: HashMap<Key, String>,
    /// Child elements of the group (e.g. `style:tab-stops`), by local name,
    /// kept as their serialized children for the readers that admit them.
    pub elements: HashMap<String, Vec<Element>>,
    /// Captured element nodes and attributes, descendants included: the
    /// cost of walking `elements`.
    pub element_nodes: usize,
}

/// A captured child element: name and attributes, plus its own children.
#[derive(Debug, Clone)]
pub struct Element {
    pub namespace: String,
    pub name: String,
    pub attributes: HashMap<Key, String>,
    pub children: Vec<Element>,
}

impl Element {
    /// Element nodes and their attributes, descendants included.
    fn nodes(&self) -> usize {
        1 + self.attributes.len() + self.children.iter().map(Element::nodes).sum::<usize>()
    }
}

#[derive(Debug, Clone)]
pub struct Style {
    pub name: String,
    pub family: String,
    pub parent: Option<String>,
    /// Declared in `office:automatic-styles` (never a parent, §16.2).
    pub automatic: bool,
    /// `style:master-page-name` and other style-element attributes.
    pub attributes: HashMap<Key, String>,
    pub groups: HashMap<String, Properties>,
}

/// Which `office:automatic-styles` a reference belongs to (§3.15): content
/// automatic styles for body content, styles.xml automatic styles for
/// master-page content. Common styles are shared by both.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Scope {
    Content,
    Styles,
}

#[derive(Debug, Default)]
pub struct StyleSheet {
    stored_bytes: usize,
    common: HashMap<(String, String), Style>,
    defaults: HashMap<String, Style>,
    automatic: HashMap<(Scope, String, String), Style>,
    list_common: HashMap<String, ListStyle>,
    list_automatic: HashMap<(Scope, String), ListStyle>,
}

/// A `text:list-style` (§16.30): its captured level definitions, and each
/// level's own `style:text-properties` as an automatic text style, so that
/// the label's formatting cascades like a span's.
#[derive(Debug)]
pub struct ListStyle {
    pub name: String,
    pub element: Element,
    pub level_text: HashMap<String, Style>,
}

/// A resolved reference to one stored style.
#[derive(Clone, Copy, Debug)]
pub struct StyleRef<'a> {
    style: &'a Style,
}

impl StyleSheet {
    pub fn add_common(&mut self, parent: roxmltree::Node<'_, '_>) -> Result<(), String> {
        for node in parent
            .children()
            .filter(|n| xml::is(*n, NS_TEXT, "list-style"))
        {
            self.charge(node)?;
            let list = read_list_style(node)?;
            if self.list_common.insert(list.name.clone(), list).is_some() {
                return Err(invalid("two list styles share a name"));
            }
        }
        for node in parent.children().filter(|n| xml::is(*n, NS_STYLE, "style")) {
            self.charge(node)?;
            let style = read_style(node)?;
            let key = (style.family.clone(), style.name.clone());
            if self.common.insert(key, style).is_some() {
                return Err(invalid("two common styles share a name and family"));
            }
        }
        for node in parent
            .children()
            .filter(|n| xml::is(*n, NS_STYLE, "default-style"))
        {
            self.charge(node)?;
            let mut style = read_style_body(node, String::new())?;
            style.parent = None;
            if self.defaults.insert(style.family.clone(), style).is_some() {
                return Err(invalid("two default styles share a family"));
            }
        }
        Ok(())
    }

    pub fn add_automatic(
        &mut self,
        scope: Scope,
        parent: roxmltree::Node<'_, '_>,
    ) -> Result<(), String> {
        for node in parent
            .children()
            .filter(|n| xml::is(*n, NS_TEXT, "list-style"))
        {
            self.charge(node)?;
            let list = read_list_style(node)?;
            if self
                .list_automatic
                .insert((scope, list.name.clone()), list)
                .is_some()
            {
                return Err(invalid("two automatic list styles share a name"));
            }
        }
        for node in parent.children().filter(|n| xml::is(*n, NS_STYLE, "style")) {
            self.charge(node)?;
            let mut style = read_style(node)?;
            style.automatic = true;
            let key = (scope, style.family.clone(), style.name.clone());
            if self.automatic.insert(key, style).is_some() {
                return Err(invalid("two automatic styles share a name and family"));
            }
        }
        Ok(())
    }

    /// Charge the storage a style element's subtree will take before it is
    /// read (borrowed XML only; nothing is allocated here).
    fn charge(&mut self, node: roxmltree::Node<'_, '_>) -> Result<(), String> {
        let bytes = node
            .descendants()
            .filter(|n| n.is_element())
            .fold(0usize, |bytes, element| {
                let name = element.tag_name();
                let attributes = element.attributes().fold(0usize, |sum, attribute| {
                    sum.saturating_add(
                        attribute.namespace().map_or(0, str::len)
                            + attribute.name().len()
                            + attribute.value().len()
                            + STORED_ENTRY_OVERHEAD,
                    )
                });
                bytes
                    .saturating_add(name.namespace().map_or(0, str::len) + name.name().len())
                    .saturating_add(STORED_ENTRY_OVERHEAD)
                    .saturating_add(attributes)
            });
        self.stored_bytes = self
            .stored_bytes
            .checked_add(bytes)
            .filter(|total| *total <= MAX_STYLE_BYTES)
            .ok_or_else(|| "OpenDocument styles exceed the style storage limit".to_string())?;
        Ok(())
    }

    /// The style a `*:style-name` attribute names: an automatic style of the
    /// scope first, else a common style (§19.471 et al.).
    pub fn find(&self, scope: Scope, family: Family, name: &str) -> Option<StyleRef<'_>> {
        let family = family.name().to_string();
        self.automatic
            .get(&(scope, family.clone(), name.to_string()))
            .or_else(|| self.common.get(&(family, name.to_string())))
            .map(|style| StyleRef { style })
    }

    /// The list style a `text:style-name` names: an automatic list style of
    /// the scope first, else a common one.
    pub fn find_list(&self, scope: Scope, name: &str) -> Option<&ListStyle> {
        self.list_automatic
            .get(&(scope, name.to_string()))
            .or_else(|| self.list_common.get(name))
    }

    pub fn default_style(&self, family: Family) -> Option<StyleRef<'_>> {
        self.defaults
            .get(family.name())
            .map(|style| StyleRef { style })
    }

    /// The style and its parent chain (common styles), nearest first.
    /// `charge` is called with each lookup's identifier bytes before the
    /// lookup.
    pub fn chain<'s>(
        &'s self,
        start: StyleRef<'s>,
        mut charge: impl FnMut(usize) -> Result<(), String>,
    ) -> Result<Vec<StyleRef<'s>>, String> {
        let mut chain = vec![start];
        let mut current = start.style;
        while let Some(parent) = current.parent.as_deref() {
            if chain.len() > MAX_PARENT_DEPTH {
                return Err(invalid("a style parent chain is cyclic or too deep"));
            }
            charge(1 + current.family.len() + parent.len())?;
            let next = self
                .common
                .get(&(current.family.clone(), parent.to_string()))
                .ok_or_else(|| invalid("a style names a missing parent style"))?;
            chain.push(StyleRef { style: next });
            current = next;
        }
        Ok(chain)
    }
}

impl<'a> StyleRef<'a> {
    /// A reference to a style the sheet owns outside the named families (a
    /// list level's text properties).
    pub fn new(style: &'a Style) -> Self {
        Self { style }
    }

    pub fn style(&self) -> &'a Style {
        self.style
    }
    pub fn group(&self, group: Group) -> Option<&'a Properties> {
        self.style.groups.get(group)
    }
}

/// An ordered cascade for one element and property group: the §16.2 search
/// order flattened into the styles to consult, nearest first. Attribute
/// values are resolved in one pass over the search order.
pub struct Cascade<'a> {
    pub styles: Vec<StyleRef<'a>>,
    pub group: Group,
    /// Nearest value of every attribute set in the search order.
    values: HashMap<&'a Key, &'a str>,
    /// The same keys in first-seen order (nearest style first).
    order: Vec<&'a Key>,
}

impl<'a> Cascade<'a> {
    pub fn new(styles: Vec<StyleRef<'a>>, group: Group) -> Self {
        let mut values: HashMap<&'a Key, &'a str> = HashMap::new();
        let mut order = Vec::new();
        for style in &styles {
            if let Some(properties) = style.group(group) {
                for (key, value) in &properties.attributes {
                    if let std::collections::hash_map::Entry::Vacant(entry) = values.entry(key) {
                        entry.insert(value.as_str());
                        order.push(key);
                    }
                }
            }
        }
        Self {
            styles,
            group,
            values,
            order,
        }
    }

    pub fn get(&self, namespace: &str, name: &str) -> Option<&'a str> {
        self.values
            .get(&(namespace.to_string(), name.to_string()))
            .copied()
    }

    pub fn element(&self, name: &str) -> Option<&'a [Element]> {
        self.styles.iter().find_map(|style| {
            style
                .group(self.group)
                .and_then(|properties| properties.elements.get(name))
                .map(Vec::as_slice)
        })
    }

    /// Every attribute key set anywhere in the cascade (for admission).
    pub fn keys(&self) -> Vec<Key> {
        self.order.iter().map(|key| (*key).clone()).collect()
    }

    /// Every element-valued property name set anywhere in the cascade.
    pub fn element_names(&self) -> Vec<String> {
        let mut seen: HashSet<&String> = HashSet::new();
        let mut names: Vec<String> = Vec::new();
        for style in &self.styles {
            if let Some(properties) = style.group(self.group) {
                for name in properties.elements.keys() {
                    if seen.insert(name) {
                        names.push(name.clone());
                    }
                }
            }
        }
        names
    }
}

fn read_list_style(node: roxmltree::Node<'_, '_>) -> Result<ListStyle, String> {
    let name = node
        .attribute((NS_STYLE, "name"))
        .ok_or_else(|| invalid("a list style has no name"))?
        .to_string();
    let element = capture(node, 0)?;
    // LibreOffice applies the last of repeated level definitions; the reader
    // takes none.
    let mut levels = HashSet::new();
    for level in node.children().filter(|n| n.is_element()) {
        if !levels.insert(list_level_number(level)?) {
            return Err(crate::error::unsupported("list-level-duplicate"));
        }
    }
    let mut level_text = HashMap::new();
    for level in node.children().filter(|n| n.is_element()) {
        let Some(text) = level
            .children()
            .find(|n| xml::is(*n, NS_STYLE, "text-properties"))
        else {
            continue;
        };
        let key = list_level_number(level)?.to_string();
        let style = Style {
            name: String::new(),
            family: "text".to_string(),
            parent: None,
            automatic: true,
            attributes: HashMap::new(),
            groups: HashMap::from([(
                TEXT.to_string(),
                Properties {
                    attributes: attributes(text),
                    ..Properties::default()
                },
            )]),
        };
        if level_text.insert(key, style).is_some() {
            return Err(invalid("a list style repeats a level"));
        }
    }
    Ok(ListStyle {
        name,
        element,
        level_text,
    })
}

/// A list level's `text:level` (§19.840, a positive integer), compared by
/// value: "01" is level 1.
pub fn list_level_number(level: roxmltree::Node<'_, '_>) -> Result<u32, String> {
    level
        .attribute((NS_TEXT, "level"))
        .and_then(|value| value.parse::<u32>().ok())
        .filter(|value| (1..=10).contains(value))
        .ok_or_else(|| crate::error::unsupported("list-level"))
}

fn read_style(node: roxmltree::Node<'_, '_>) -> Result<Style, String> {
    let name = node
        .attribute((NS_STYLE, "name"))
        .ok_or_else(|| invalid("a style has no name"))?
        .to_string();
    read_style_body(node, name)
}

fn read_style_body(node: roxmltree::Node<'_, '_>, name: String) -> Result<Style, String> {
    let family = node
        .attribute((NS_STYLE, "family"))
        .ok_or_else(|| invalid("a style has no family"))?
        .to_string();
    let parent = node
        .attribute((NS_STYLE, "parent-style-name"))
        .map(str::to_string);
    let mut groups = HashMap::new();
    for child in node.children().filter(|n| n.is_element()) {
        // LibreOffice also writes `loext:graphic-properties` (its area fill
        // model) beside the paragraph properties; it is kept as its own group
        // and checked by the reader.
        let group = match (child.tag_name().namespace(), child.tag_name().name()) {
            (Some(NS_STYLE), name) => name.to_string(),
            (Some(crate::xml::NS_LOEXT), "graphic-properties") => LOEXT_GRAPHIC.to_string(),
            _ => return Err(crate::error::unsupported("style-child")),
        };
        let mut properties = Properties {
            attributes: attributes(child),
            ..Properties::default()
        };
        for element in child.children().filter(|n| n.is_element()) {
            // Each element-valued property occurs at most once per group.
            let captured = capture(element, 0)?;
            properties.element_nodes += captured.nodes();
            match properties
                .elements
                .entry(element.tag_name().name().to_string())
            {
                std::collections::hash_map::Entry::Occupied(_) => {
                    return Err(invalid("a style repeats an element-valued property"));
                }
                std::collections::hash_map::Entry::Vacant(entry) => {
                    entry.insert(vec![captured]);
                }
            }
        }
        if groups.insert(group, properties).is_some() {
            return Err(invalid("a style repeats a property group"));
        }
    }
    Ok(Style {
        name,
        family,
        parent,
        automatic: false,
        attributes: attributes(node),
        groups,
    })
}

pub fn attributes(node: roxmltree::Node<'_, '_>) -> HashMap<Key, String> {
    node.attributes()
        .map(|attribute| {
            (
                (
                    attribute.namespace().unwrap_or("").to_string(),
                    attribute.name().to_string(),
                ),
                attribute.value().to_string(),
            )
        })
        .collect()
}

/// Implementation bound: property child elements (tab stops, columns, list
/// levels) nest a few levels at most.
const MAX_CAPTURE_DEPTH: usize = 8;

fn capture(node: roxmltree::Node<'_, '_>, depth: usize) -> Result<Element, String> {
    if depth > MAX_CAPTURE_DEPTH {
        return Err(crate::error::unsupported("style-nesting"));
    }
    Ok(Element {
        namespace: node.tag_name().namespace().unwrap_or("").to_string(),
        name: node.tag_name().name().to_string(),
        attributes: attributes(node),
        children: node
            .children()
            .filter(|n| n.is_element())
            .map(|child| capture(child, depth + 1))
            .collect::<Result<_, _>>()?,
    })
}
