//! Explicit OfficeArt master-shape links (MS-ODRAW 2.2.40 / 2.3.2.1).
use super::*;
use std::{collections::BTreeMap, rc::Rc};

// Resource policy for retained master-shape metadata, independent of slide count.
const MAX_MASTER_SHAPES: usize = 100_000;
pub(super) struct Node {
    pub id: u32,
    pub parent: Option<u32>,
    pub text_type: Option<u16>,
    pub direct: Vec<Option<text_style::Level>>,
    pub base: Option<Rc<text_style::Master>>,
    pub paint: paint::Paint,
}
#[derive(Default)]
pub(super) struct Resolver {
    nodes: BTreeMap<u32, Node>,
    resolved: BTreeMap<u32, Resolved>,
}
struct Resolved {
    levels: Rc<Vec<text_style::Level>>,
    paint: paint::Paint,
    depth: usize,
}
impl Resolver {
    pub fn insert(&mut self, node: Node) -> Result<(), String> {
        if self.nodes.len() + self.resolved.len() >= MAX_MASTER_SHAPES {
            return Err(unsupported("PowerPoint master shape limit exceeded"));
        }
        if node.id == 0 || self.nodes.contains_key(&node.id) || self.resolved.contains_key(&node.id)
        {
            return Err(unsupported("duplicate or zero PowerPoint master shape ID"));
        }
        self.nodes.insert(node.id, node);
        Ok(())
    }
    pub fn finish(&mut self, budget: &mut usize) -> Result<(), String> {
        for id in self.nodes.keys().copied().collect::<Vec<_>>() {
            self.resolve(id, &mut Vec::new(), budget)?;
        }
        // Parsing metadata is no longer needed once immutable levels/paint exist.
        self.nodes.clear();
        Ok(())
    }
    pub fn levels(&self, id: u32) -> Result<&[text_style::Level], String> {
        self.resolved
            .get(&id)
            .map(|v| v.levels.as_slice())
            .ok_or_else(|| unsupported("unresolved PowerPoint master shape"))
    }
    pub fn paint(&self, id: u32) -> Result<&paint::Paint, String> {
        self.resolved
            .get(&id)
            .map(|v| &v.paint)
            .ok_or_else(|| unsupported("unresolved PowerPoint master shape"))
    }
    fn resolve(
        &mut self,
        id: u32,
        path: &mut Vec<u32>,
        budget: &mut usize,
    ) -> Result<Rc<Vec<text_style::Level>>, String> {
        *budget = budget
            .checked_sub(1)
            .ok_or_else(|| unsupported("PowerPoint master shape work budget exceeded"))?;
        if let Some(resolved) = self.resolved.get(&id) {
            // A cached suffix still counts toward the complete chain depth.
            if path.len() + resolved.depth > MAX_DEPTH {
                return Err(unsupported("excessive PowerPoint master shape inheritance"));
            }
            return Ok(resolved.levels.clone());
        }
        if path.len() >= MAX_DEPTH || path.contains(&id) {
            return Err(unsupported(
                "cyclic or excessive PowerPoint master shape inheritance",
            ));
        }
        let node = self
            .nodes
            .get(&id)
            .ok_or_else(|| unsupported("unresolved PowerPoint master shape"))?;
        let parent = node.parent;
        path.push(id);
        let inherited = parent
            .map(|parent| self.resolve(parent, path, budget))
            .transpose()?;
        let depth = parent.map_or(1, |parent| self.resolved[&parent].depth + 1);
        let node = &self.nodes[&id];
        let paint = match parent {
            Some(parent) => node.paint.inherit(&self.resolved[&parent].paint),
            None => node.paint,
        };
        let base = inherited.as_ref().map(|v| v.as_slice()).or_else(|| {
            node.base
                .as_ref()
                .and_then(|b| node.text_type.and_then(|t| b.levels(t)))
        });
        let mut levels = Vec::with_capacity(5);
        for index in 0..5 {
            let local = node.direct.get(index).and_then(Option::as_ref);
            let inherited = base.and_then(|b| b.get(index));
            levels.push(match local {
                Some(v) => v.inherit(inherited),
                None => inherited
                    .cloned()
                    .unwrap_or_else(|| text_style::Level::empty(index as u16)),
            });
        }
        path.pop();
        let levels = Rc::new(levels);
        self.resolved.insert(
            id,
            Resolved {
                levels: levels.clone(),
                paint,
                depth,
            },
        );
        Ok(levels)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn node(id: u32, parent: Option<u32>) -> Node {
        Node {
            id,
            parent,
            text_type: None,
            direct: Vec::new(),
            base: None,
            paint: paint::Paint::default(),
        }
    }
    #[test]
    fn paint_resolves_through_cached_chains_without_losing_explicit_false() {
        let mut r = Resolver::default();
        let mut root = node(1, None);
        root.paint.property(0x181, 255).unwrap();
        root.paint.property(0x1c0, 0xff0000).unwrap();
        let mut middle = node(2, Some(1));
        middle.paint.property(0x1bf, 0x00100000).unwrap();
        let mut leaf = node(3, Some(2));
        leaf.paint.property(0x1bf, 0x00100010).unwrap();
        r.insert(root).unwrap();
        r.insert(middle).unwrap();
        r.insert(leaf).unwrap();
        r.finish(&mut 100).unwrap();
        let middle = r.paint(2).unwrap().xml_with_scheme(1, None);
        assert!(!middle.contains("FF0000"));
        assert!(middle.contains("0000FF"));
        let leaf = r.paint(3).unwrap().xml_with_scheme(1, None);
        assert!(leaf.contains("FF0000"));
        assert!(leaf.contains("0000FF"));
        assert!(r.paint(4).is_err());
    }
    #[test]
    fn resolves_chains_reuses_cache_and_releases_parse_metadata() {
        let mut r = Resolver::default();
        r.insert(node(1, None)).unwrap();
        r.insert(node(2, Some(1))).unwrap();
        r.finish(&mut 100).unwrap();
        assert!(r.nodes.is_empty());
        assert_eq!(r.levels(2).unwrap().len(), 5);
        assert!(r.levels(3).is_err());
        let a = r.resolve(2, &mut Vec::new(), &mut 1).unwrap();
        let b = r.resolve(2, &mut Vec::new(), &mut 1).unwrap();
        assert!(Rc::ptr_eq(&a, &b));
        assert!(r.resolve(2, &mut Vec::new(), &mut 0).is_err());
        assert!(r.insert(node(2, None)).is_err());
    }
    #[test]
    fn rejects_cycles_missing_parents_and_excessive_depth() {
        let mut r = Resolver::default();
        r.insert(node(1, Some(2))).unwrap();
        r.insert(node(2, Some(1))).unwrap();
        assert!(r.finish(&mut 100).unwrap_err().contains("cyclic"));
        let mut r = Resolver::default();
        r.insert(node(1, Some(2))).unwrap();
        assert!(r.finish(&mut 100).unwrap_err().contains("unresolved"));
        let mut r = Resolver::default();
        for id in 1..=MAX_DEPTH as u32 + 1 {
            r.insert(node(id, Some(id + 1))).unwrap();
        }
        assert!(r.finish(&mut 1000).unwrap_err().contains("excessive"));
        // Ascending resolution order caches every parent before its child.
        // The cap must not depend on whether ancestors were already resolved.
        let mut r = Resolver::default();
        r.insert(node(1, None)).unwrap();
        for id in 2..=MAX_DEPTH as u32 {
            r.insert(node(id, Some(id - 1))).unwrap();
        }
        r.finish(&mut 1000).unwrap();
        r.insert(node(MAX_DEPTH as u32 + 1, Some(MAX_DEPTH as u32)))
            .unwrap();
        assert!(r.finish(&mut 1000).unwrap_err().contains("excessive"));
    }
    #[test]
    fn bounds_retained_nodes_before_inserting_and_rejects_ambiguous_ids() {
        let mut r = Resolver::default();
        assert!(r.insert(node(0, None)).is_err());
        for id in 1..=MAX_MASTER_SHAPES as u32 {
            r.insert(node(id, None)).unwrap();
        }
        assert!(r
            .insert(node(MAX_MASTER_SHAPES as u32 + 1, None))
            .unwrap_err()
            .contains("limit"));
        let mut r = Resolver::default();
        r.insert(node(1, None)).unwrap();
        assert!(r.insert(node(1, None)).unwrap_err().contains("duplicate"));
    }
}
