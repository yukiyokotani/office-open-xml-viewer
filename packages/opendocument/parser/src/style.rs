//! Styles and their cascade (ODF Part 3 §16.2, §16.4).
//!
//! Every property group (`style:paragraph-properties`, `style:text-properties`,
//! …) is stored attribute by attribute; inheritance is per attribute (§16.2),
//! except element-valued properties such as `style:tab-stops`, which inherit
//! as a whole element (§17.7) and are stored with their parent group.

use crate::budget;
use crate::error::invalid;
use crate::xml::{self, Stepped, NS_STYLE, NS_TEXT};
use std::collections::{HashMap, HashSet};

/// Implementation bound on a parent chain; ODF requires none, a cycle is a
/// container error.
const MAX_PARENT_DEPTH: usize = 64;
/// Implementation bound on one style value (a name, a family, a length, a
/// colour): [`check_values`] rejects a part holding a longer one before any
/// style is stored, so a reader may treat a style value as a small constant
/// however often it resolves it.
pub const MAX_STYLE_VALUE_BYTES: usize = 1024;
/// Implementation bound on a style condition (§19.468 `style:condition`).
pub const MAX_STYLE_CONDITION_BYTES: usize = 4096;
const NS_DATASTYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

/// Reject a part with a style value over [`MAX_STYLE_VALUE_BYTES`] (a
/// condition: [`MAX_STYLE_CONDITION_BYTES`]): the attributes of style,
/// data-style and property elements (whatever the namespace of an extension
/// property set), of list and outline level styles and the notes and line
/// numbering configurations, and every `*style-name` reference.
pub fn check_values(root: roxmltree::Node<'_, '_>) -> Result<(), String> {
    for node in root.subtree().filter(|n| n.is_element()) {
        budget::tick(1)?;
        let namespace = node.tag_name().namespace();
        let name = node.tag_name().name();
        let style_element = namespace == Some(NS_STYLE)
            || namespace == Some(NS_DATASTYLE)
            || name.ends_with("-properties")
            || (namespace == Some(NS_TEXT)
                && (name.starts_with("list-level-style-")
                    || matches!(
                        name,
                        "list-style"
                            | "outline-style"
                            | "outline-level-style"
                            | "notes-configuration"
                            | "linenumbering-configuration"
                    )));
        for attribute in node.attributes() {
            let name = attribute.name();
            let bound = if style_element && name == "condition" {
                MAX_STYLE_CONDITION_BYTES
            } else if style_element || name.ends_with("style-name") {
                MAX_STYLE_VALUE_BYTES
            } else {
                continue;
            };
            if attribute.value().len() > bound {
                return Err(crate::error::resource("style value"));
            }
        }
    }
    Ok(())
}

/// The style families readers look up (§19.480 `style:family`). Styles of
/// every family are stored under their declared family name.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Family {
    Paragraph,
    Text,
    Table,
    TableColumn,
    TableRow,
    TableCell,
    Graphic,
    Section,
    Ruby,
}

impl Family {
    fn name(self) -> &'static str {
        match self {
            Family::Paragraph => "paragraph",
            Family::Text => "text",
            Family::Table => "table",
            Family::TableColumn => "table-column",
            Family::TableRow => "table-row",
            Family::TableCell => "table-cell",
            Family::Graphic => "graphic",
            Family::Section => "section",
            Family::Ruby => "ruby",
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
    /// The steps of reading the element: one per node and attribute,
    /// descendants included (values are bounded by [`check_values`]).
    pub fn nodes(&self) -> usize {
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
    common: HashMap<(String, String), Style>,
    defaults: HashMap<String, Style>,
    automatic: HashMap<(Scope, String, String), Style>,
    list_common: HashMap<String, ListStyle>,
    list_automatic: HashMap<(Scope, String), ListStyle>,
    /// `text:outline-style` (§16.36): heading numbering, captured like a list
    /// style.
    outline: Option<ListStyle>,
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
    /// Common styles of `office:styles` (§3.15.2). Style names are unique
    /// per family (§19.498.2); a document that repeats one is read anyway:
    /// the later definition replaces the earlier one, as in the ODS and ODP
    /// readers' style maps. A style without the name or family
    /// that every reference needs cannot be referenced and is skipped.
    pub fn add_common(&mut self, parent: roxmltree::Node<'_, '_>) -> Result<(), String> {
        for node in parent.kids().filter(|n| xml::is(*n, NS_TEXT, "list-style")) {
            budget::tick(1)?;
            if let Some(list) = read_list_style(node)? {
                self.list_common.insert(list.name.clone(), list);
            }
        }
        for node in parent.kids().filter(|n| xml::is(*n, NS_STYLE, "style")) {
            budget::tick(1)?;
            if let Some(style) = read_style(node)? {
                let key = (style.family.clone(), style.name.clone());
                self.common.insert(key, style);
            }
        }
        for node in parent
            .kids()
            .filter(|n| xml::is(*n, NS_STYLE, "default-style"))
        {
            budget::tick(1)?;
            if let Some(mut style) = read_style_body(node, String::new())? {
                style.parent = None;
                self.defaults.insert(style.family.clone(), style);
            }
        }
        Ok(())
    }

    /// Capture `text:outline-style` children of `office:styles` (§16.36).
    pub fn add_outline(&mut self, parent: roxmltree::Node<'_, '_>) -> Result<(), String> {
        for node in parent
            .kids()
            .filter(|n| xml::is(*n, NS_TEXT, "outline-style"))
        {
            budget::tick(1)?;
            let mut list = read_list_style_named(node, String::new())?;
            list.name = node.attribute((NS_STYLE, "name")).unwrap_or("").to_string();
            self.outline = Some(list);
        }
        Ok(())
    }

    /// The document's outline style, if any.
    pub fn outline(&self) -> Option<&ListStyle> {
        self.outline.as_ref()
    }

    pub fn add_automatic(
        &mut self,
        scope: Scope,
        parent: roxmltree::Node<'_, '_>,
    ) -> Result<(), String> {
        // Repeated, nameless and familyless styles as in `add_common`.
        for node in parent.kids().filter(|n| xml::is(*n, NS_TEXT, "list-style")) {
            budget::tick(1)?;
            if let Some(list) = read_list_style(node)? {
                self.list_automatic.insert((scope, list.name.clone()), list);
            }
        }
        for node in parent.kids().filter(|n| xml::is(*n, NS_STYLE, "style")) {
            budget::tick(1)?;
            if let Some(mut style) = read_style(node)? {
                style.automatic = true;
                let key = (scope, style.family.clone(), style.name.clone());
                self.automatic.insert(key, style);
            }
        }
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

    /// Whether `name` names a common list style (not an automatic one of
    /// `scope`): its references resolve among common styles only.
    pub fn list_is_common(&self, scope: Scope, name: &str) -> bool {
        !self.list_automatic.contains_key(&(scope, name.to_string()))
            && self.list_common.contains_key(name)
    }

    /// A common style of the family (§16.2: common styles name common
    /// styles).
    pub fn find_common(&self, family: Family, name: &str) -> Option<StyleRef<'_>> {
        self.common
            .get(&(family.name().to_string(), name.to_string()))
            .map(|style| StyleRef { style })
    }

    pub fn default_style(&self, family: Family) -> Option<StyleRef<'_>> {
        self.defaults
            .get(family.name())
            .map(|style| StyleRef { style })
    }

    /// The style and its parent chain (common styles), nearest first.
    pub fn chain<'s>(&'s self, start: StyleRef<'s>) -> Result<Vec<StyleRef<'s>>, String> {
        let mut chain = vec![start];
        let mut current = start.style;
        while let Some(parent) = current.parent.as_deref() {
            if chain.len() > MAX_PARENT_DEPTH {
                return Err(invalid("a style parent chain is cyclic or too deep"));
            }
            budget::tick(1)?;
            // A missing parent ends the chain (LibreOffice ignores it).
            let Some(next) = self
                .common
                .get(&(current.family.clone(), parent.to_string()))
            else {
                break;
            };
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
    /// Nearest value of every attribute set in the search order, by
    /// (namespace, name): the index of its key and value in `order`. The
    /// keys borrow the stored names, so a lookup allocates nothing.
    values: HashMap<(&'a str, &'a str), usize>,
    /// The keys and values in first-seen order (nearest style first).
    order: Vec<(&'a Key, &'a str)>,
}

impl<'a> Cascade<'a> {
    /// The cascade of `group` over `styles`. Building one is the only way
    /// to resolve properties, and it counts its own steps: one per style
    /// of the search order, per attribute and per captured element node.
    pub fn build(styles: Vec<StyleRef<'a>>, group: Group) -> Result<Self, String> {
        budget::tick(styles.iter().fold(0usize, |steps, style| {
            steps.saturating_add(1).saturating_add(
                style
                    .group(group)
                    .map_or(0, |p| p.attributes.len() + p.element_nodes),
            )
        }))?;
        let mut values: HashMap<(&'a str, &'a str), usize> = HashMap::new();
        let mut order = Vec::new();
        for style in &styles {
            if let Some(properties) = style.group(group) {
                for (key, value) in &properties.attributes {
                    if let std::collections::hash_map::Entry::Vacant(entry) =
                        values.entry((key.0.as_str(), key.1.as_str()))
                    {
                        entry.insert(order.len());
                        order.push((key, value.as_str()));
                    }
                }
            }
        }
        Ok(Self {
            styles,
            group,
            values,
            order,
        })
    }

    pub fn get(&self, namespace: &str, name: &str) -> Option<&'a str> {
        // The map's key lifetime shortens to the arguments' (it is
        // covariant); the value comes from `order`, at its full lifetime.
        let values: &HashMap<(&str, &str), usize> = &self.values;
        values
            .get(&(namespace, name))
            .map(|index| self.order[*index].1)
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
        self.order.iter().map(|(key, _)| (*key).clone()).collect()
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

/// A named list style; `None` for one without a name, which nothing can
/// reference.
fn read_list_style(node: roxmltree::Node<'_, '_>) -> Result<Option<ListStyle>, String> {
    let Some(name) = node.attribute((NS_STYLE, "name")) else {
        return Ok(None);
    };
    read_list_style_named(node, name.to_string()).map(Some)
}

fn read_list_style_named(node: roxmltree::Node<'_, '_>, name: String) -> Result<ListStyle, String> {
    let element = capture(node, 0)?;
    // Repeated level definitions: LibreOffice applies the last one, and so
    // does the map below (readers look levels up through `level_text` and the
    // captured element, last match first).
    let mut level_text = HashMap::new();
    for level in node.kids().filter(|n| n.is_element()) {
        let Some(text) = level
            .kids()
            .rfind(|n| xml::is(*n, NS_STYLE, "text-properties"))
        else {
            continue;
        };
        let Some(number) = list_level_number(level) else {
            continue;
        };
        let key = number.to_string();
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
        level_text.insert(key, style);
    }
    Ok(ListStyle {
        name,
        element,
        level_text,
    })
}

/// A list level's `text:level` (§19.840, a positive integer), compared by
/// value: "01" is level 1. ODF lists have ten levels (LibreOffice's
/// `MAXLEVEL`); a level definition outside 1–10, or without a level, is
/// never used and is ignored rather than refusing the document.
pub fn list_level_number(level: roxmltree::Node<'_, '_>) -> Option<u32> {
    list_level_value(level.attribute((NS_TEXT, "level"))?)
}

/// [`list_level_number`] of an attribute value.
pub fn list_level_value(value: &str) -> Option<u32> {
    value
        .parse::<u32>()
        .ok()
        .filter(|value| (1..=MAX_LIST_LEVEL).contains(value))
}

/// The deepest list level ODF defines (§19.840; LibreOffice's ten levels).
pub const MAX_LIST_LEVEL: u32 = 10;

/// A named style; `None` for one without a name, which nothing can
/// reference.
fn read_style(node: roxmltree::Node<'_, '_>) -> Result<Option<Style>, String> {
    let Some(name) = node.attribute((NS_STYLE, "name")) else {
        return Ok(None);
    };
    read_style_body(node, name.to_string())
}

/// A style's body; `None` without `style:family` (§19.480, required), which
/// no reference or default can select.
fn read_style_body(node: roxmltree::Node<'_, '_>, name: String) -> Result<Option<Style>, String> {
    let Some(family) = node.attribute((NS_STYLE, "family")) else {
        return Ok(None);
    };
    let family = family.to_string();
    let parent = node
        .attribute((NS_STYLE, "parent-style-name"))
        .map(str::to_string);
    let mut groups = HashMap::new();
    for child in node.kids().filter(|n| n.is_element()) {
        // LibreOffice also writes `loext:graphic-properties` (its area fill
        // model) beside the paragraph properties; it is kept as its own group
        // and checked by the reader.
        let group = match (child.tag_name().namespace(), child.tag_name().name()) {
            (Some(NS_STYLE), name) => name.to_string(),
            (Some(crate::xml::NS_LOEXT), "graphic-properties") => LOEXT_GRAPHIC.to_string(),
            // Other producers' extension groups carry nothing ODF defines.
            _ => continue,
        };
        let mut properties = Properties {
            attributes: attributes(child),
            ..Properties::default()
        };
        for element in child.kids().filter(|n| n.is_element()) {
            // Each element-valued property occurs at most once per group.
            let captured = capture(element, 0)?;
            properties.element_nodes += captured.nodes();
            // A repeated element-valued property: the first is kept.
            properties
                .elements
                .entry(element.tag_name().name().to_string())
                .or_insert_with(|| vec![captured]);
        }
        // A repeated property group: the first is kept.
        groups.entry(group).or_insert(properties);
    }
    Ok(Some(Style {
        name,
        family,
        parent,
        automatic: false,
        attributes: attributes(node),
        groups,
    }))
}

/// The shorthand properties that set four sides at once (§20.205
/// fo:margin, §20.217 fo:padding, §20.183 fo:border, §20.185
/// style:border-line-width), by namespace and name.
pub const SIDE_SHORTHANDS: [(&str, &str); 4] = [
    (crate::xml::NS_FO, "margin"),
    (crate::xml::NS_FO, "padding"),
    (crate::xml::NS_FO, "border"),
    (NS_STYLE, "border-line-width"),
];
pub const SIDES: [&str; 4] = ["top", "bottom", "left", "right"];

/// An element's attributes, with each side shorthand also written out as
/// its four sides (a side the element sets itself wins). Sides are then
/// ordinary properties: a style's shorthand replaces the sides of the
/// styles it inherits from, and a reader of a side sees the shorthand.
pub fn attributes(node: roxmltree::Node<'_, '_>) -> HashMap<Key, String> {
    let mut out: HashMap<Key, String> = node
        .attributes()
        .map(|attribute| {
            (
                (
                    attribute.namespace().unwrap_or("").to_string(),
                    attribute.name().to_string(),
                ),
                attribute.value().to_string(),
            )
        })
        .collect();
    for (namespace, name) in SIDE_SHORTHANDS {
        let Some(value) = node.attribute((namespace, name)) else {
            continue;
        };
        for side in SIDES {
            out.entry((namespace.to_string(), format!("{name}-{side}")))
                .or_insert_with(|| value.to_string());
        }
    }
    out
}

/// Implementation bound: property child elements (tab stops, columns, list
/// levels) nest a few levels at most.
const MAX_CAPTURE_DEPTH: usize = 8;

fn capture(node: roxmltree::Node<'_, '_>, depth: usize) -> Result<Element, String> {
    budget::tick(1)?;
    // Deeper descendants are not captured (no property nests that far).
    let children = if depth >= MAX_CAPTURE_DEPTH {
        Vec::new()
    } else {
        node.kids()
            .filter(|n| n.is_element())
            .map(|child| capture(child, depth + 1))
            .collect::<Result<_, _>>()?
    };
    Ok(Element {
        namespace: node.tag_name().namespace().unwrap_or("").to_string(),
        name: node.tag_name().name().to_string(),
        attributes: attributes(node),
        children,
    })
}
