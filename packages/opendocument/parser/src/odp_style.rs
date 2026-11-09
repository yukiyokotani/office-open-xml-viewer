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
}

impl<'a> Sheet<'a> {
    pub fn add_document(&mut self, root: Nd<'a>, scope: Scope) -> Result<(), String> {
        for child in root.children().filter(|n| n.is_element()) {
            if xml::is(child, NS_OFFICE, "font-face-decls") {
                for face in child
                    .children()
                    .filter(|n| xml::is(*n, NS_STYLE, "font-face"))
                {
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
                    let family = node.attribute((NS_STYLE, "family")).unwrap_or("");
                    self.defaults.insert(family, node);
                }
                (Some(NS_TEXT), "list-style") => {
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
                    if let Some(name) = node.attribute((crate::xml::NS_TABLE, "name")) {
                        self.named.insert(("table-template", name), node);
                    }
                }
                (Some(NS_DRAW), kind) => {
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
