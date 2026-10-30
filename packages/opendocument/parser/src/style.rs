//! Styles and their cascade (ODF Part 3 §16.2, §16.4).
//!
//! Every property group (`style:paragraph-properties`, `style:text-properties`,
//! …) is stored attribute by attribute; inheritance is per attribute (§16.2),
//! except element-valued properties such as `style:tab-stops`, which inherit
//! as a whole element (§17.7) and are stored with their parent group.

use crate::error::invalid;
use crate::xml::{self, NS_STYLE};
use std::collections::HashMap;

/// Implementation bound on a parent chain; ODF requires none, a cycle is a
/// container error.
const MAX_PARENT_DEPTH: usize = 64;

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

/// An attribute key: (namespace URI, local name).
pub type Key = (String, String);

#[derive(Debug, Default, Clone)]
pub struct Properties {
    pub attributes: HashMap<Key, String>,
    /// Child elements of the group (e.g. `style:tab-stops`), by local name,
    /// kept as their serialized children for the readers that admit them.
    pub elements: HashMap<String, Vec<Element>>,
}

/// A captured child element: name and attributes, plus its own children.
#[derive(Debug, Clone)]
pub struct Element {
    pub namespace: String,
    pub name: String,
    pub attributes: HashMap<Key, String>,
    pub children: Vec<Element>,
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
}

/// A resolved reference to one stored style.
#[derive(Clone, Copy, Debug)]
pub struct StyleRef<'a> {
    style: &'a Style,
}

impl StyleSheet {
    pub fn add_common(&mut self, parent: roxmltree::Node<'_, '_>) -> Result<(), String> {
        for node in parent.children().filter(|n| xml::is(*n, NS_STYLE, "style")) {
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
        for node in parent.children().filter(|n| xml::is(*n, NS_STYLE, "style")) {
            let mut style = read_style(node)?;
            style.automatic = true;
            let key = (scope, style.family.clone(), style.name.clone());
            if self.automatic.insert(key, style).is_some() {
                return Err(invalid("two automatic styles share a name and family"));
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
    pub fn style(&self) -> &'a Style {
        self.style
    }
    pub fn group(&self, group: Group) -> Option<&'a Properties> {
        self.style.groups.get(group)
    }
}

/// An ordered cascade for one element and property group: the §16.2 search
/// order flattened into the styles to consult, nearest first.
pub struct Cascade<'a> {
    pub styles: Vec<StyleRef<'a>>,
    pub group: Group,
}

impl<'a> Cascade<'a> {
    pub fn get(&self, namespace: &str, name: &str) -> Option<&'a str> {
        self.styles.iter().find_map(|style| {
            style
                .group(self.group)
                .and_then(|properties| {
                    properties
                        .attributes
                        .get(&(namespace.to_string(), name.to_string()))
                })
                .map(String::as_str)
        })
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
        let mut keys: Vec<Key> = Vec::new();
        for style in &self.styles {
            if let Some(properties) = style.group(self.group) {
                for key in properties.attributes.keys() {
                    if !keys.contains(key) {
                        keys.push(key.clone());
                    }
                }
            }
        }
        keys
    }

    /// Every element-valued property name set anywhere in the cascade.
    pub fn element_names(&self) -> Vec<String> {
        let mut names: Vec<String> = Vec::new();
        for style in &self.styles {
            if let Some(properties) = style.group(self.group) {
                for name in properties.elements.keys() {
                    if !names.contains(name) {
                        names.push(name.clone());
                    }
                }
            }
        }
        names
    }
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
        if child.tag_name().namespace() != Some(NS_STYLE) {
            return Err(crate::error::unsupported("style-child"));
        }
        let mut properties = Properties {
            attributes: attributes(child),
            ..Properties::default()
        };
        for element in child.children().filter(|n| n.is_element()) {
            properties
                .elements
                .entry(element.tag_name().name().to_string())
                .or_default()
                .push(capture(element, 0)?);
        }
        if groups
            .insert(child.tag_name().name().to_string(), properties)
            .is_some()
        {
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
