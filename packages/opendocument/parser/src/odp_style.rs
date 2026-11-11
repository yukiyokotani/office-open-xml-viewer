//! Style lookup for the ODP reader (ODF 1.3 Part 3 §16).
//!
//! The presentation reader resolves styles directly over the parsed
//! `content.xml` / `styles.xml` trees: a style is a borrowed `style:style`
//! node, a cascade is the ordered list of property elements (nearest first)
//! that §16.2 consults for one attribute. Property groups other than the text
//! families (graphic, presentation, drawing-page, table-cell, …) are read
//! here; the shared `style.rs` sheet stays specific to text documents.

use crate::error::invalid;
use crate::xml::{self, NS_DRAW, NS_LOEXT, NS_OFFICE, NS_STYLE, NS_SVG, NS_TEXT};
use std::collections::HashMap;

pub type Nd<'a> = roxmltree::Node<'a, 'a>;

/// Implementation bound on a parent chain; a cycle is a container error.
const MAX_PARENT_DEPTH: usize = 64;
/// Implementation bounds on one style record, so every property lookup
/// through a style costs a bounded amount: element children of a style
/// (property elements, maps) or list style (levels), and the attributes and
/// element children of one property element.
const MAX_STYLE_CHILDREN: usize = 64;
const MAX_PROPERTY_ATTRIBUTES: usize = 256;
const MAX_PROPERTY_CHILDREN: usize = 64;
const MAX_PROPERTY_VALUE_BYTES: usize = 1024;
/// All nodes of one style record (elements, text, comments; nested
/// records such as tab stops included).
const MAX_STYLE_NODES: usize = 4096;
/// All direct children of one style record (any node type), so finding a
/// property element in it scans at most this many.
const MAX_STYLE_CHILD_NODES: usize = 64;
/// The work charged for searching one style's children (in work units of
/// 16 nodes), see [`MAX_STYLE_CHILD_NODES`].
pub const STYLE_SCAN_COST: usize = MAX_STYLE_CHILD_NODES / 16;

/// The work of reading a property cascade: each property element, each of
/// its attributes, and their value bytes (in 64-byte units), which bounds
/// the cost of the lookups and value parsing done through it.
pub fn props_cost(props: &Props<'_>) -> usize {
    props.nodes.iter().fold(0usize, |sum, node| {
        let attributes = node
            .attributes()
            .fold(0usize, |s, a| s.saturating_add(1 + a.value().len() / 64));
        sum.saturating_add(1).saturating_add(attributes)
    })
}

/// Reject a named drawing definition (gradient, hatch, marker, dash,
/// opacity, fill image, table template) with more nodes than a style
/// record may have: every use reads it again.
fn check_definition(node: Nd<'_>) -> Result<(), String> {
    let error = || invalid("a drawing definition exceeds the style record limits");
    if node.descendants().nth(MAX_STYLE_NODES).is_some() {
        return Err(error());
    }
    let marker = xml::is(node, NS_DRAW, "marker");
    for n in node.descendants().filter(|n| n.is_element()) {
        let long = n.attributes().any(|a| {
            a.value().len() > MAX_PROPERTY_VALUE_BYTES
                && !(marker && a.namespace() == Some(NS_SVG) && a.name() == "d")
        });
        if long {
            return Err(error());
        }
    }
    Ok(())
}

/// Reject a style record past the bounds above.
fn check_style_record(style: Nd<'_>) -> Result<(), String> {
    let error = || invalid("a style exceeds the style record limits");
    if style.descendants().nth(MAX_STYLE_NODES).is_some()
        || style.children().nth(MAX_STYLE_CHILD_NODES).is_some()
    {
        return Err(error());
    }
    // Every attribute of the record (its own and its nested ones).
    if style.descendants().filter(|n| n.is_element()).any(|n| {
        n.attributes()
            .any(|a| a.value().len() > MAX_PROPERTY_VALUE_BYTES)
    }) {
        return Err(error());
    }
    let mut children = 0usize;
    for child in style.children().filter(|n| n.is_element()) {
        children += 1;
        if children > MAX_STYLE_CHILDREN {
            return Err(error());
        }
        if child.attributes().len() > MAX_PROPERTY_ATTRIBUTES
            || child
                .attributes()
                .any(|a| a.value().len() > MAX_PROPERTY_VALUE_BYTES)
        {
            return Err(error());
        }
        if child
            .children()
            .filter(|n| n.is_element())
            .nth(MAX_PROPERTY_CHILDREN)
            .is_some()
        {
            return Err(error());
        }
    }
    Ok(())
}

/// Which `office:automatic-styles` a reference belongs to (§3.15).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Scope {
    Content,
    Styles,
}

#[derive(Default)]
pub struct Sheet<'a> {
    automatic: HashMap<(Scope, &'a str, &'a str), Nd<'a>>,
    common: HashMap<(&'a str, &'a str), Nd<'a>>,
    defaults: HashMap<&'a str, Nd<'a>>,
    lists_automatic: HashMap<(Scope, &'a str), Nd<'a>>,
    lists_common: HashMap<&'a str, Nd<'a>>,
    /// `draw:gradient`, `draw:hatch`, `draw:fill-image`, `draw:marker`,
    /// `draw:stroke-dash`, `draw:opacity` by (local name, draw:name).
    named: HashMap<(&'a str, &'a str), Nd<'a>>,
    pub master_pages: HashMap<&'a str, Nd<'a>>,
    pub master_order: Vec<&'a str>,
    page_layouts: HashMap<(Scope, &'a str), Nd<'a>>,
    font_faces: HashMap<&'a str, String>,
    /// `text:list` elements with an `xml:id` (§19.785 continuation targets),
    /// with the scope of their document.
    pub list_ids: HashMap<&'a str, (Scope, Nd<'a>)>,
    /// Resolved list numbers per text, for continuations across texts
    /// (`odp_text`); `None` while a text is being resolved.
    pub list_numbers: ListNumbers,
}

/// Cache of the last numbers of the identified lists of each text, by the
/// text's (document, node) identity.
pub type ListNumbers = std::cell::RefCell<
    HashMap<(usize, roxmltree::NodeId), Option<std::rc::Rc<HashMap<String, u32>>>>,
>;

impl<'a> Sheet<'a> {
    pub fn add_document(&mut self, root: Nd<'a>, scope: Scope) -> Result<(), String> {
        for list in root.descendants().filter(|n| xml::is(*n, NS_TEXT, "list")) {
            // List identities are hashed for every numbered item and
            // continuation: bounded like any short name.
            let long = |v: Option<&str>| v.is_some_and(|v| v.len() > MAX_PROPERTY_VALUE_BYTES);
            if long(list.attribute((xml::NS_XML, "id")))
                || long(list.attribute((NS_TEXT, "continue-list")))
                || long(list.attribute((NS_TEXT, "style-name")))
            {
                return Err(invalid("a list identifier exceeds the name limit"));
            }
            if let Some(id) = list.attribute((xml::NS_XML, "id")) {
                self.list_ids.entry(id).or_insert((scope, list));
            }
        }
        for child in root.children().filter(|n| n.is_element()) {
            if xml::is(child, NS_OFFICE, "font-face-decls") {
                for face in child
                    .children()
                    .filter(|n| xml::is(*n, NS_STYLE, "font-face"))
                {
                    // Font names are copied into every run that uses them.
                    if face
                        .attributes()
                        .any(|a| a.value().len() > MAX_PROPERTY_VALUE_BYTES)
                    {
                        return Err(invalid("a font face exceeds the name limit"));
                    }
                    if let Some(name) = face.attribute((NS_STYLE, "name")) {
                        let family = face
                            .attribute((NS_SVG, "font-family"))
                            .map(unquote_family)
                            .unwrap_or_else(|| name.to_string());
                        self.font_faces.insert(name, family);
                    }
                }
            } else if xml::is(child, NS_OFFICE, "styles") {
                self.add_styles(child, None)?;
            } else if xml::is(child, NS_OFFICE, "automatic-styles") {
                self.add_styles(child, Some(scope))?;
            } else if xml::is(child, NS_OFFICE, "master-styles") {
                for master in child
                    .children()
                    .filter(|n| xml::is(*n, NS_STYLE, "master-page"))
                {
                    let name = master
                        .attribute((NS_STYLE, "name"))
                        .ok_or_else(|| invalid("a master page has no name"))?;
                    // Its name selects its outline styles for every paragraph.
                    if name.len() > MAX_PROPERTY_VALUE_BYTES {
                        return Err(invalid("a master page name exceeds the name limit"));
                    }
                    if self.master_pages.insert(name, master).is_none() {
                        self.master_order.push(name);
                    }
                }
            }
        }
        Ok(())
    }

    fn add_styles(&mut self, parent: Nd<'a>, scope: Option<Scope>) -> Result<(), String> {
        for node in parent.children().filter(|n| n.is_element()) {
            let ns = node.tag_name().namespace();
            let local = node.tag_name().name();
            match (ns, local) {
                (Some(NS_STYLE), "style") => {
                    check_style_record(node)?;
                    let name = node
                        .attribute((NS_STYLE, "name"))
                        .ok_or_else(|| invalid("a style has no name"))?;
                    let family = node.attribute((NS_STYLE, "family")).unwrap_or("");
                    match scope {
                        Some(scope) => {
                            self.automatic.insert((scope, family, name), node);
                        }
                        None => {
                            self.common.insert((family, name), node);
                        }
                    }
                }
                (Some(NS_STYLE), "default-style") => {
                    check_style_record(node)?;
                    let family = node.attribute((NS_STYLE, "family")).unwrap_or("");
                    self.defaults.insert(family, node);
                }
                (Some(NS_TEXT), "list-style") => {
                    check_style_record(node)?;
                    if let Some(name) = node.attribute((NS_STYLE, "name")) {
                        match scope {
                            Some(scope) => {
                                self.lists_automatic.insert((scope, name), node);
                            }
                            None => {
                                self.lists_common.insert(name, node);
                            }
                        }
                    }
                }
                (Some(NS_STYLE), "page-layout") => {
                    if let Some(name) = node.attribute((NS_STYLE, "name")) {
                        self.page_layouts
                            .insert((scope.unwrap_or(Scope::Styles), name), node);
                    }
                }
                (Some(crate::xml::NS_TABLE), "table-template") => {
                    check_definition(node)?;
                    if let Some(name) = node.attribute((crate::xml::NS_TABLE, "name")) {
                        self.named.insert(("table-template", name), node);
                    }
                }
                (Some(NS_DRAW), kind) => {
                    check_definition(node)?;
                    if let Some(name) = node.attribute((NS_DRAW, "name")) {
                        self.named.insert((kind, name), node);
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    /// The named drawing element (`draw:gradient`, …) of `kind`.
    pub fn named(&self, kind: &str, name: &str) -> Option<Nd<'a>> {
        self.named.get(&(kind, name)).copied()
    }

    pub fn page_layout(&self, name: &str) -> Option<Nd<'a>> {
        self.page_layouts
            .get(&(Scope::Styles, name))
            .or_else(|| self.page_layouts.get(&(Scope::Content, name)))
            .copied()
    }

    pub fn font_family(&self, face: &str) -> Option<&str> {
        self.font_faces.get(face).map(String::as_str)
    }

    pub fn find(&self, scope: Scope, family: &str, name: &str) -> Option<Nd<'a>> {
        self.automatic
            .get(&(scope, family, name))
            .or_else(|| self.common.get(&(family, name)))
            .copied()
    }

    pub fn common(&self, family: &str, name: &str) -> Option<Nd<'a>> {
        self.common.get(&(family, name)).copied()
    }

    pub fn default_style(&self, family: &str) -> Option<Nd<'a>> {
        self.defaults.get(family).copied()
    }

    pub fn list(&self, scope: Scope, name: &str) -> Option<Nd<'a>> {
        self.lists_automatic
            .get(&(scope, name))
            .or_else(|| self.lists_common.get(name))
            .copied()
    }

    /// The style named `name` and its parents (common styles of the same
    /// family), nearest first. A missing parent ends the chain (§16.2: the
    /// default style then applies); a cycle is rejected.
    pub fn chain(&self, scope: Scope, family: &str, name: &str) -> Result<Vec<Nd<'a>>, String> {
        let mut out = Vec::new();
        let Some(mut current) = self.find(scope, family, name) else {
            return Ok(out);
        };
        out.push(current);
        while let Some(parent) = current.attribute((NS_STYLE, "parent-style-name")) {
            if out.len() > MAX_PARENT_DEPTH {
                return Err(invalid("a style parent chain is cyclic or too deep"));
            }
            let Some(next) = self.common(family, parent) else {
                break;
            };
            out.push(next);
            current = next;
        }
        Ok(out)
    }
}

/// `svg:font-family` is a CSS font-family list; the first family is used
/// (quotes removed).
pub fn unquote_family(value: &str) -> String {
    let first = value.split(',').next().unwrap_or("").trim();
    first.trim_matches(|c| c == '\'' || c == '"').to_string()
}

/// An ordered list of property elements (nearest first) for one property
/// group: attribute lookups take the first element that sets the attribute.
#[derive(Clone, Default)]
pub struct Props<'a> {
    pub nodes: Vec<Nd<'a>>,
}

impl<'a> Props<'a> {
    pub fn new() -> Self {
        Self { nodes: Vec::new() }
    }

    /// Append the `group` property element of each style (nearest first).
    pub fn extend_styles(&mut self, styles: &[Nd<'a>], group: (&str, &str)) {
        for style in styles {
            if let Some(element) = property_element(*style, group) {
                self.nodes.push(element);
            }
        }
    }

    pub fn get(&self, namespace: &str, name: &str) -> Option<&'a str> {
        self.nodes
            .iter()
            .find_map(|node| node.attribute((namespace, name)))
    }

    /// The nearest element-valued property `namespace:name` (e.g.
    /// `style:columns`), inherited as a whole (§17.7).
    pub fn child(&self, namespace: &str, name: &str) -> Option<Nd<'a>> {
        self.nodes
            .iter()
            .find_map(|node| node.children().find(|c| xml::is(*c, namespace, name)))
    }
}

pub const GRAPHIC: (&str, &str) = (NS_STYLE, "graphic-properties");
pub const PARAGRAPH: (&str, &str) = (NS_STYLE, "paragraph-properties");
pub const TEXT: (&str, &str) = (NS_STYLE, "text-properties");
pub const DRAWING_PAGE: (&str, &str) = (NS_STYLE, "drawing-page-properties");
pub const TABLE_CELL: (&str, &str) = (NS_STYLE, "table-cell-properties");
pub const TABLE_COLUMN: (&str, &str) = (NS_STYLE, "table-column-properties");
pub const TABLE_ROW: (&str, &str) = (NS_STYLE, "table-row-properties");
/// LibreOffice's paragraph area fill (`loext:graphic-properties`).
pub const LOEXT_GRAPHIC: (&str, &str) = (NS_LOEXT, "graphic-properties");

pub fn property_element<'a>(style: Nd<'a>, group: (&str, &str)) -> Option<Nd<'a>> {
    style.children().find(|c| xml::is(*c, group.0, group.1))
}
