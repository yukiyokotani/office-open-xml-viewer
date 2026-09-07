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
    pub geometry: crate::officeart::geometry::SpannedGeometry,
    pub gradient: crate::officeart::gradient::Spanned,
}
#[derive(Default)]
pub(super) struct Resolver {
    nodes: BTreeMap<u32, Node>,
    resolved: BTreeMap<u32, Resolved>,
}
struct Resolved {
    levels: Rc<Vec<text_style::Level>>,
    text_base: Option<TextBase>,
    paint: paint::Paint,
    geometry: crate::officeart::geometry::SpannedGeometry,
    gradient: crate::officeart::gradient::Spanned,
    depth: usize,
}
#[derive(Clone)]
struct TextBase {
    authored_font_sizes: Rc<text_style::AuthoredFontSizeTable>,
    text_type: u16,
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
    pub fn authored_base_font_sizes(
        &self,
        id: u32,
        authored_type: u16,
    ) -> Result<Option<(u16, text_style::AuthoredFontSizes)>, String> {
        let resolved = self
            .resolved
            .get(&id)
            .ok_or_else(|| unsupported("unresolved PowerPoint master shape"))?;
        Ok(resolved.text_base.as_ref().and_then(|base| {
            base.authored_font_sizes
                .get(usize::from(authored_type))
                .copied()
                .flatten()
                .map(|sizes| (base.text_type, sizes))
        }))
    }
    pub fn paint(&self, id: u32) -> Result<&paint::Paint, String> {
        self.resolved
            .get(&id)
            .map(|v| &v.paint)
            .ok_or_else(|| unsupported("unresolved PowerPoint master shape"))
    }
    pub fn geometry(
        &self,
        id: u32,
    ) -> Result<&crate::officeart::geometry::SpannedGeometry, String> {
        self.resolved
            .get(&id)
            .map(|v| &v.geometry)
            .ok_or_else(|| unsupported("unresolved PowerPoint master geometry"))
    }
    pub fn gradient(
        &self,
        id: u32,
    ) -> Result<&crate::officeart::gradient::Spanned, String> {
        self.resolved
            .get(&id)
            .map(|value| &value.gradient)
            .ok_or_else(|| unsupported("unresolved PowerPoint master gradient"))
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
        let text_base = match parent {
            Some(parent) => self.resolved[&parent].text_base.clone(),
            None => node
                .base
                .as_ref()
                .zip(node.text_type)
                .map(|(master, text_type)| TextBase {
                    authored_font_sizes: master.authored_font_size_table(),
                    text_type,
                }),
        };
        let paint = match parent {
            Some(parent) => node.paint.inherit(&self.resolved[&parent].paint),
            None => node.paint,
        };
        let geometry = match parent {
            Some(parent) => node.geometry.inherit(&self.resolved[&parent].geometry),
            None => node.geometry.clone(),
        };
        let gradient = match parent {
            Some(parent) => node.gradient.inherit(&self.resolved[&parent].gradient),
            None => node.gradient.clone(),
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
                text_base,
                paint,
                geometry,
                gradient,
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
            geometry: crate::officeart::geometry::SpannedGeometry::default(),
            gradient: crate::officeart::gradient::Spanned::default(),
        }
    }
    fn master_size(size: u16) -> Rc<text_style::Master> {
        let bytes = [
            1u16.to_le_bytes().to_vec(),
            0u16.to_le_bytes().to_vec(),
            0u32.to_le_bytes().to_vec(),
            0x20000u32.to_le_bytes().to_vec(),
            size.to_le_bytes().to_vec(),
        ]
        .concat();
        Rc::new(
            text_style::Master::parse(
                &[Record {
                    version: 0,
                    instance: 8,
                    kind: 4003,
                    payload: &bytes,
                }],
                &[],
                &mut 100,
            )
            .unwrap(),
        )
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
    fn resolved_parent_chain_retains_authored_master_base_provenance_after_finish() {
        let master = master_size(20);
        let weak = Rc::downgrade(&master);
        let mut root = node(1, None);
        root.text_type = Some(1);
        root.base = Some(master.clone());
        let mut r = Resolver::default();
        r.insert(root).unwrap();
        r.insert(node(2, Some(1))).unwrap();
        drop(master);
        r.finish(&mut 100).unwrap();
        assert!(r.nodes.is_empty());
        assert!(weak.upgrade().is_none());
        let (base_type, authored) = r.authored_base_font_sizes(2, 8).unwrap().unwrap();
        assert_eq!(base_type, 1);
        assert_eq!(authored.level_count(), 1);
        assert_eq!(authored.get(0), Some(20));
        assert!(r.authored_base_font_sizes(2, 7).unwrap().is_none());
        assert!(r.authored_base_font_sizes(3, 8).is_err());
    }
    #[test]
    fn parent_text_base_provenance_wins_over_an_ignored_child_base() {
        let mut root = node(1, None);
        root.text_type = Some(1);
        root.base = Some(master_size(20));
        let mut child = node(2, Some(1));
        child.text_type = Some(2);
        child.base = Some(master_size(44));
        let mut absent = node(3, None);
        absent.text_type = Some(1);
        let mut r = Resolver::default();
        r.insert(root).unwrap();
        r.insert(child).unwrap();
        r.insert(absent).unwrap();
        r.finish(&mut 100).unwrap();
        let (base_type, authored) = r.authored_base_font_sizes(2, 8).unwrap().unwrap();
        assert_eq!(base_type, 1);
        assert_eq!(authored.get(0), Some(20));
        assert!(r.authored_base_font_sizes(3, 8).unwrap().is_none());
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
        assert!(
            r.insert(node(MAX_MASTER_SHAPES as u32 + 1, None))
                .unwrap_err()
                .contains("limit")
        );
        let mut r = Resolver::default();
        r.insert(node(1, None)).unwrap();
        assert!(r.insert(node(1, None)).unwrap_err().contains("duplicate"));
    }

    #[test]
    fn resolved_master_geometry_survives_backing_move_and_empty_child_reset() {
        let mut backing = [2u16.to_le_bytes(), 2u16.to_le_bytes(), 8u16.to_le_bytes()].concat();
        backing.extend([0i32, 0, 10, 10].into_iter().flat_map(i32::to_le_bytes));
        let span = crate::officeart::ByteSpan::new(0..backing.len(), backing.len(), "geometry").unwrap();
        let mut root = node(1, None);
        root.geometry.complex(0x145, span);
        let mut child = node(2, Some(1));
        child.geometry.scalar(0x145, 0).unwrap();
        let mut resolver = Resolver::default();
        resolver.insert(root).unwrap();
        resolver.insert(child).unwrap();
        resolver.finish(&mut 20).unwrap();
        let moved = backing;
        assert!(resolver.geometry(1).unwrap().view(&moved).unwrap().decode(&mut 10).unwrap().is_some());
        assert!(resolver.geometry(2).unwrap().view(&moved).unwrap().decode(&mut 10).unwrap().is_none());
        assert!(resolver.geometry(1).unwrap().view(&moved[..moved.len() - 1]).is_err());
    }
    #[test]
    fn resolved_master_gradient_inherits_and_explicit_reset_vetoes_parent() {
        let backing = vec![1, 0, 1, 0, 8, 0, 7, 0, 0, 0, 0, 0, 0, 0];
        let span = crate::officeart::ByteSpan::new(
            0..backing.len(),
            backing.len(),
            "gradient",
        )
        .unwrap();
        let mut root = node(1, None);
        root.gradient.set(span);
        let inherited = node(2, Some(1));
        let mut reset = node(3, Some(2));
        reset.gradient.scalar(0);
        let mut resolver = Resolver::default();
        resolver.insert(root).unwrap();
        resolver.insert(inherited).unwrap();
        resolver.insert(reset).unwrap();
        resolver.finish(&mut 20).unwrap();
        assert_eq!(
            resolver
                .gradient(2)
                .unwrap()
                .view(&backing)
                .unwrap()
                .decode(&mut 1, &mut 8)
                .unwrap()
                .unwrap()[0]
                .color,
            7
        );
        assert!(resolver
            .gradient(3)
            .unwrap()
            .view(&backing)
            .unwrap()
            .decode(&mut 1, &mut 8)
            .unwrap()
            .is_none());
    }
}
