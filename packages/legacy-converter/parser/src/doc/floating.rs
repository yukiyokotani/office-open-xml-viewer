//! Floating drawing anchors of the main and header documents, MS-DOC 2.8.27
//! and 2.9.253.
use super::pictures::Options as PictureOptions;
use super::{u16_at, u32_at, unsupported};
use crate::officeart::{
    raster::{read_store_entry_as, Image, Raster},
    record_with_end, Record,
};
use std::collections::{BTreeMap, BTreeSet};

mod direct;
mod group;
mod reading_picture;
mod shape;
pub(in crate::doc) mod textbox;
pub(in crate::doc) use direct::DirectRun;

/// The drawing part that owns a PlcfSpa, its OfficeArtDgContainer and its
/// textbox story (MS-DOC 2.8.27, 2.9.171; MS-ODRAW 2.2.13).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::doc) enum Part {
    Main,
    Header,
}

impl Part {
    /// FibRgFcLcb97 fcPlcSpaMom / fcPlcSpaHdr.
    fn anchor_table(self) -> usize {
        match self {
            Self::Main => 0x1da,
            Self::Header => 0x1e2,
        }
    }
    /// OfficeArtWordDrawing.dgglbl: 0 main document, 1 header document.
    fn label(self) -> u8 {
        match self {
            Self::Main => 0,
            Self::Header => 1,
        }
    }
}

/// The anchors and top-level shape containers of one drawing part.
#[derive(Default)]
struct Drawings<'a> {
    anchors: Vec<Anchor>,
    /// spid -> registered shape.
    shapes: BTreeMap<u32, ContainerShape<'a>>,
    container: Option<Record<'a>>,
    nested_solver: bool,
    line_rules: Option<BTreeMap<u32, [u32; 2]>>,
    acquired_ids: BTreeMap<u32, Option<u32>>,
    new_line_acquired: bool,
    ambiguous_ids: bool,
}

pub(super) struct Store<'a> {
    /// Indexed by `Part as usize`: main document, header document.
    parts: [Drawings<'a>; 2],
    entries: Vec<Record<'a>>,
    word: &'a [u8],
    table: &'a [u8],
    /// The CLX that maps textbox stories; empty when none was supplied.
    clx: &'a [u8],
    group_read: bool,
    /// Default-strict library policy. Only the direct DOC constructor may opt
    /// in; full native formatting/story/resource/omission gates remain intact.
    reading_relocation: bool,
    reading_picture_defaults: reading_picture::Defaults<'a>,
    header_container: Option<Record<'a>>,
    textboxes: [Option<textbox::Textboxes<'a>>; 2],
    images: BTreeMap<usize, Option<Image<'a>>>,
    /// Whether PNG BLIPs holding TIFF data are admitted (direct model only).
    pub raster: Raster,
    budget: usize,
    remaining_bytes: usize,
    occurrences: u32,
    selected_images: std::collections::BTreeSet<usize>,
    pub omitted: bool,
}

impl<'a> Store<'a> {
    pub(in crate::doc) fn set_reading_relocation(&mut self, enabled: bool) {
        self.reading_relocation = enabled;
    }

    fn allows_reading_contour(&self, wrapping: u8) -> bool {
        self.reading_relocation && matches!(wrapping, 4 | 5)
    }

    /// Read the main document's drawing part. Header drawings are loaded only
    /// on request, so this subset never assigns them to the main story.
    #[cfg(test)]
    pub fn read(word: &'a [u8], table: &'a [u8], main_units: usize) -> Result<Self, String> {
        Self::read_stories(word, table, &[], main_units)
    }

    /// Like `read`, retaining the CLX from which the direct model can later
    /// decode textbox stories on request.
    pub fn read_stories(
        word: &'a [u8],
        table: &'a [u8],
        clx: &'a [u8],
        main_units: usize,
    ) -> Result<Self, String> {
        let anchors = anchors_in(word, table, Part::Main, main_units)?;
        let mut result = Self {
            parts: [
                Drawings {
                    anchors,
                    ..Drawings::default()
                },
                Drawings::default(),
            ],
            entries: Vec::new(),
            word,
            table,
            clx,
            group_read: false,
            reading_relocation: false,
            reading_picture_defaults: reading_picture::Defaults::default(),
            header_container: None,
            textboxes: [None, None],
            images: BTreeMap::new(),
            raster: Raster::Advertised,
            budget: 1_000_000,
            remaining_bytes: 128 * 1024 * 1024,
            occurrences: 0,
            selected_images: std::collections::BTreeSet::new(),
            omitted: false,
        };
        if result.parts[0].anchors.is_empty() {
            return Ok(result);
        }
        result.read_group()?;
        Ok(result)
    }

    /// Load the header document's anchors (PlcSpaHdr, CPs relative to the
    /// header document) and its own drawing container (dgglbl 1).
    fn load_header(&mut self, header_units: usize) -> Result<(), String> {
        let anchors = anchors_in(self.word, self.table, Part::Header, header_units)?;
        if anchors.is_empty() {
            return Ok(());
        }
        self.parts[1].anchors = anchors;
        self.read_group()?;
        if let Some(container) = self.header_container {
            self.parts[1].shapes = container_shapes(
                container,
                &mut self.budget,
                &mut self.parts[1].nested_solver,
            )?;
            self.parts[1].container = Some(container);
        } else if !self.omitted {
            return Err(unsupported("Word header anchors lack their drawing"));
        }
        Ok(())
    }

    /// Load everything the direct model resolves beyond main-story anchors:
    /// header drawings and both textbox stories (MS-DOC 2.3.6-2.3.7).
    pub(in crate::doc) fn load_direct_parts(&mut self) -> Result<(), String> {
        let header_units = u32_at(self.word, 0x54)? as usize; // FibRgLw97.ccpHdd
        self.load_header(header_units)?;
        for part in [Part::Main, Part::Header] {
            self.textboxes[part as usize] =
                textbox::Textboxes::read(self.word, self.table, self.clx, part)?;
        }
        Ok(())
    }

    pub(in crate::doc) fn textbox(&self, part: Part) -> Option<&textbox::Textboxes<'a>> {
        self.textboxes[part as usize].as_ref()
    }

    fn read_group(&mut self) -> Result<(), String> {
        if self.group_read {
            return Ok(());
        }
        self.group_read = true;
        // MS-DOC 2.9.171: the OfficeArt delay stream is WordDocument, NOT Data.
        let offset = u32_at(self.word, 0x22a)? as usize;
        let length = u32_at(self.word, 0x22e)? as usize;
        if length == 0 {
            self.omitted = true;
            return Ok(());
        }
        let bytes = self
            .table
            .get(offset..)
            .and_then(|b| b.get(..length))
            .ok_or_else(|| unsupported("Word drawing group out of bounds"))?;
        let (group, mut position) =
            record_with_end(bytes, 0, &mut self.budget, "Word drawing group")?;
        if group.kind != 0xf000 || group.version != 15 {
            return Err(unsupported("invalid Word drawing group"));
        }
        let mut store_seen = false;
        for child in records(group.payload, &mut self.budget)? {
            self.reading_picture_defaults.register(child);
            if child.kind != 0xf001 {
                continue;
            }
            if store_seen || child.version != 15 {
                return Err(unsupported("invalid Word floating image store"));
            }
            store_seen = true;
            self.entries = records(child.payload, &mut self.budget)?;
            if self.entries.len() != usize::from(child.instance) {
                return Err(unsupported("Word floating image store count mismatch"));
            }
        }
        let mut seen = [false; 2];
        while position < bytes.len() {
            let label = bytes[position];
            let (drawing, end) =
                record_with_end(bytes, position + 1, &mut self.budget, "Word drawing")?;
            position = end;
            if label > 1 || drawing.kind != 0xf002 || drawing.version != 15 {
                return Err(unsupported("invalid Word drawing container"));
            }
            if seen[usize::from(label)] {
                return Err(unsupported("duplicate Word drawing container"));
            }
            seen[usize::from(label)] = true;
            if label == Part::Main.label() {
                self.parts[0].shapes =
                    container_shapes(drawing, &mut self.budget, &mut self.parts[0].nested_solver)?;
                self.parts[0].container = Some(drawing);
            } else {
                // Never leak header drawings into the body; they are
                // resolved only against header anchors.
                self.header_container = Some(drawing);
            }
        }
        Ok(())
    }

    /// Associate IDs encountered during ordinary group acquisition with their
    /// owning anchor. No discovery traversal is added. Historical-only parts
    /// retain their permissive treatment of unused identity collisions; once
    /// the new line class uses solver targets, its whole acquired part must
    /// have unambiguous identities, including top-level anchor IDs.
    fn bind_group_ids(
        &mut self,
        part: Part,
        owner: u32,
        ids: impl Iterator<Item = u32>,
        new_line: bool,
    ) -> Result<(), String> {
        let drawing = &mut self.parts[part as usize];
        let mut local_ids = BTreeSet::new();
        for id in ids {
            let binding = drawing.acquired_ids.entry(id).or_insert(Some(owner));
            if !local_ids.insert(id)
                || *binding != Some(owner)
                || (id != owner && drawing.shapes.contains_key(&id))
            {
                *binding = None;
                drawing.ambiguous_ids = true;
            }
            drawing.ambiguous_ids |= id == 0;
        }
        drawing.new_line_acquired |= new_line;
        if drawing.new_line_acquired && drawing.ambiguous_ids {
            return Err(unsupported("ambiguous Word line connector shape ownership"));
        }
        Ok(())
    }

    /// Retain the authored static line; do not reconstruct endpoint routing.
    /// MS-ODRAW 2.2.36 permits consumers to ignore connector rules and ignores
    /// cptiA/B when spidA/B is zero. Our narrower policy admits new msosptLine
    /// connectors only when no endpoint binding is discarded. Established
    /// connector presets keep their existing contract. Parse the solver once,
    /// lazily, so scenes without an acquired new line retain their old budget
    /// and treatment of unregistered shapes. Solvers found in the patriarch or
    /// acquired line/group headers are rejected outside their Dg solver scope;
    /// this does not recursively validate unacquired shape contents.
    fn check_line_connector(
        &mut self,
        part: Part,
        kind: u16,
        flags: u32,
        spid: u32,
        nested_solver: bool,
    ) -> Result<(), String> {
        if kind != 20 || flags & 0x100 == 0 {
            return Ok(());
        }
        let drawing = &mut self.parts[part as usize];
        if drawing
            .acquired_ids
            .get(&spid)
            .is_some_and(|owner| owner.is_none())
        {
            return Err(unsupported("ambiguous Word line connector shape ownership"));
        }
        if spid == 0 || nested_solver || drawing.nested_solver {
            return Err(unsupported(
                "invalid Word line connector ownership or solver scope",
            ));
        }
        if drawing.line_rules.is_none() {
            let container = drawing
                .container
                .ok_or_else(|| unsupported("line connector lacks its drawing"))?;
            let mut rules = BTreeMap::new();
            let mut ids = BTreeSet::new();
            // MS-ODRAW 2.2.13 provides solvers1/solvers2; 2.2.18 declares
            // recInstance as the contained file-block count, not a zero instance.
            let mut solver_count = 0;
            for solver in records(container.payload, &mut self.budget)? {
                if solver.kind != 0xf005 {
                    continue;
                }
                solver_count += 1;
                if solver_count > 2 || solver.version != 15 {
                    return Err(unsupported("invalid Word line connector solver"));
                }
                let children = records(solver.payload, &mut self.budget)?;
                if usize::from(solver.instance) != children.len() {
                    return Err(unsupported("Word line connector solver count mismatch"));
                }
                for rule in children {
                    if rule.kind != 0xf012
                        || rule.version != 1
                        || rule.instance != 0
                        || rule.payload.len() != 24
                    {
                        return Err(unsupported("unsupported Word line connector solver rule"));
                    }
                    let id = u32_at(rule.payload, 0)?;
                    let target = u32_at(rule.payload, 12)?;
                    if target == 0
                        || !ids.insert(id)
                        || rules
                            .insert(target, [u32_at(rule.payload, 4)?, u32_at(rule.payload, 8)?])
                            .is_some()
                    {
                        return Err(unsupported(
                            "duplicate line connector rule or invalid target",
                        ));
                    }
                }
            }
            drawing.line_rules = Some(rules);
        }
        if drawing
            .line_rules
            .as_ref()
            .and_then(|rules| rules.get(&spid))
            .is_some_and(|endpoints| *endpoints != [0, 0])
        {
            return Err(unsupported("line connector has bound endpoints"));
        }
        Ok(())
    }

    fn resolve(&mut self, part: Part, cp: usize) -> Result<Option<ResolvedDrawing>, String> {
        let drawings = &self.parts[part as usize];
        let Ok(index) = drawings.anchors.binary_search_by_key(&cp, |a| a.cp) else {
            self.omitted = true;
            return Ok(None);
        };
        let anchor = drawings.anchors[index].clone();
        let anchor = &anchor;
        let Some(&(anchor_index, order, shape)) = drawings.shapes.get(&anchor.shape_id) else {
            self.omitted = true;
            return Ok(None);
        };
        if anchor_index != index {
            return Err(unsupported("Word shape/anchor index mismatch"));
        }
        if shape.kind == 0xf003 {
            // A top-level OfficeArt group, resolved in `group`.
            return self.resolve_group(part, anchor, order, shape);
        }
        let mut picture = PictureOptions::default();
        let mut placement = Placement::default();
        let mut flags = None;
        let mut kind = None;
        let mut nested_solver = false;
        for property in records(shape.payload, &mut self.budget)? {
            match property.kind {
                0xf00a => {
                    if flags.is_some() || property.version != 2 || property.payload.len() != 8 {
                        return Err(unsupported("invalid Word floating shape properties"));
                    }
                    flags = Some(u32_at(property.payload, 4)?);
                    kind = Some(property.instance);
                }
                0xf00b | 0xf122 => {
                    picture.apply_indexed(property, &mut self.budget)?;
                    placement.apply(property, &mut self.budget)?;
                }
                0xf005 => nested_solver = true,
                _ => {}
            }
        }
        let flags = flags.unwrap_or(0);
        if placement.hidden && !placement.script {
            // MS-ODRAW 2.3.4.44 fHidden: the shape is prevented from
            // displaying, so the direct model projects nothing for it. Word's
            // PDF of a corpus document with hidden header lines agrees.
            return Ok(None);
        }
        let [left, top, right, bottom] = anchor.rect.map(i64::from);
        let extent = [(right - left) * 635, (bottom - top) * 635];
        if kind != Some(75) {
            // Non-picture drawing shapes: every property is classified by
            // `shape`; unsupported sub-cases are rejected with their reason.
            if placement.hidden || placement.script {
                self.omitted = true;
                return Ok(None);
            }
            let facts = shape::Facts::read(
                kind.ok_or_else(|| unsupported("Word drawing shape lacks its type"))?,
                flags,
                false,
                shape,
                extent,
                &mut self.budget,
            )?;
            self.check_line_connector(part, kind.unwrap(), flags, anchor.shape_id, nested_solver)?;
            if kind == Some(20) && flags & 0x100 != 0 {
                self.bind_group_ids(
                    part,
                    anchor.shape_id,
                    std::iter::once(anchor.shape_id),
                    true,
                )?;
            }
            // No Word evidence yet shows whether a rotated top-level shape's
            // SPA rectangle holds its rotated bounds (see `group`).
            if facts.rotation.rem_euclid(360.0) != 0.0 {
                return Err(unsupported(
                    "rotated Word drawing shapes outside groups are not supported",
                ));
            }
            if !self.load_fill_picture(&facts)? {
                self.omitted = true;
                return Ok(None);
            }
            if facts.pseudo_inline {
                // A pseudo-inline shape is the drawing of a SHAPE field: an
                // absolutely positioned shape at the anchor character's
                // origin (posrelh/posrelv 3, character and line, with a zero
                // SPA offset). Word's own DOCX of the corpus documents writes
                // each as an ordinary `wp:inline` WPS shape of the SPA size
                // in the field's place. Any offset, alignment or wrapping
                // other than none has no such evidence.
                let [left, top, _, _] = anchor.rect;
                if placement.relative_horizontal != Some(3)
                    || placement.relative_vertical != Some(3)
                    || placement.horizontal != 0
                    || placement.vertical != 0
                    || left != 0
                    || top != 0
                    || anchor.wrapping != 3
                {
                    return Err(unsupported(
                        "Word pseudo-inline drawing is not at its character origin",
                    ));
                }
                let mut drawing = self.finish(
                    anchor,
                    &placement,
                    order,
                    extent,
                    flags,
                    [None; 2],
                    Content::Shape(Box::new(facts)),
                )?;
                drawing.inline = true;
                return Ok(Some(drawing));
            }
            let align = direct_alignment(anchor, &placement)?;
            if matches!(anchor.wrapping, 0 | 4 | 5) {
                return Err(unsupported(
                    "Word drawing shape uses an unsupported wrap contour",
                ));
            }
            return self
                .finish(
                    anchor,
                    &placement,
                    order,
                    extent,
                    flags,
                    align,
                    Content::Shape(Box::new(facts)),
                )
                .map(Some);
        }
        let align = if placement.horizontal == 0 && placement.vertical == 0 {
            None
        } else {
            match direct_alignment(anchor, &placement) {
                Ok(align) => Some(align),
                Err(_) => {
                    self.omitted = true;
                    return Ok(None);
                }
            }
        };
        if kind != Some(75)
            // MS-ODRAW 2.2.40 identifies fOleShape (0x10) as shape
            // metadata, not permission to execute or inspect an OLE payload.
            // Its indexed passive BLIP still passes the ordinary validation
            // and resource budgets below. Retain every other structural gate.
            || flags & 0x10f != 0
            || placement.hidden
            || placement.script
            || picture.rotation != 0
            // SPA provides an explicit, host-defined coordinate origin; aligned
            // positions are accepted only through `direct_alignment`.
            || (align.is_none() && (placement.horizontal != 0 || placement.vertical != 0))
            || (matches!(anchor.wrapping, 0 | 4 | 5)
                && !self.allows_reading_contour(anchor.wrapping))
        {
            self.omitted = true;
            return Ok(None);
        }
        if self.allows_reading_contour(anchor.wrapping) {
            if anchor.textbox_count != 0 {
                return Err(unsupported(
                    "reading picture has unacquired owned textbox content",
                ));
            }
            reading_picture::acquire(shape, &self.reading_picture_defaults, &mut self.budget)?;
        }
        let Some(image_index) = picture.pib else {
            self.omitted = true;
            return Ok(None);
        };
        let Some(extension) = self.load_image(image_index)? else {
            self.omitted = true;
            return Ok(None);
        };
        if extent.iter().any(|v| *v <= 0) {
            return Err(unsupported("invalid Word floating picture extent"));
        }
        if picture.crop[0] + picture.crop[1] >= 100000
            || picture.crop[2] + picture.crop[3] >= 100000
        {
            return Err(unsupported("empty Word floating picture crop"));
        }
        let content = Content::Picture {
            image_index,
            extension,
            crop: picture.crop,
        };
        self.finish(
            anchor,
            &placement,
            order,
            extent,
            flags,
            align.unwrap_or([None; 2]),
            content,
        )
        .map(|mut drawing| {
            if self.allows_reading_contour(anchor.wrapping) {
                drawing.reading_picture_source = Some((part, anchor.shape_id));
            }
            Some(drawing)
        })
    }

    /// Load a shape's picture fill BLIP; `false` when it is not a supported
    /// passive image, which the caller reports as omitted content.
    fn load_fill_picture(&mut self, facts: &shape::Facts) -> Result<bool, String> {
        match facts.fill_picture {
            Some((index, _)) => Ok(self.load_image(index)?.is_some()),
            None => Ok(true),
        }
    }

    /// Decode an indexed delayed BLIP once (MS-DOC 2.9.171) under the store's
    /// media budget. `None` means the entry is not a supported passive image.
    fn load_image(&mut self, image_index: usize) -> Result<Option<&'static str>, String> {
        if !self.images.contains_key(&image_index) {
            let entry = *self
                .entries
                .get(image_index)
                .ok_or_else(|| unsupported("Word floating image index out of bounds"))?;
            let image = read_store_entry_as(
                entry,
                Some(self.word),
                &mut self.budget,
                self.remaining_bytes,
                self.raster,
            )?;
            if let Some(image) = &image {
                self.remaining_bytes = self
                    .remaining_bytes
                    .checked_sub(image.bytes.len())
                    .ok_or_else(|| unsupported("Word floating media budget exceeded"))?;
            }
            self.images.insert(image_index, image);
        }
        Ok(self.images[&image_index]
            .as_ref()
            .map(|image| image.extension))
    }

    #[allow(clippy::too_many_arguments)]
    fn finish(
        &mut self,
        anchor: &Anchor,
        placement: &Placement,
        order: u32,
        extent: [i64; 2],
        flags: u32,
        align: [Option<&'static str>; 2],
        content: Content,
    ) -> Result<ResolvedDrawing, String> {
        let [left, top, _, _] = anchor.rect.map(i64::from);
        i32::try_from(left * 635)
            .and_then(|_| i32::try_from(top * 635))
            .map_err(|_| unsupported("Word floating position exceeds DrawingML range"))?;
        if self.occurrences >= 100_000 {
            return Err(unsupported("Word floating occurrence budget exceeded"));
        }
        self.occurrences += 1;
        Ok(ResolvedDrawing {
            content,
            reading_picture_source: None,
            inline: false,
            shape_id: anchor.shape_id,
            extent,
            flip: [flags & 0x40 != 0, flags & 0x80 != 0],
            x_emu: left * 635,
            y_emu: top * 635,
            horizontal: anchor.horizontal,
            vertical: anchor.vertical,
            align,
            wrapping: anchor.wrapping,
            side: anchor.side,
            behind: anchor.behind,
            locked: anchor.locked,
            distances: placement.distances,
            in_cell: placement.in_cell,
            overlap: placement.overlap,
            z_order: placement.z_order.unwrap_or(order),
            occurrence: self.occurrences,
        })
    }
}

enum Content {
    Picture {
        image_index: usize,
        extension: &'static str,
        crop: [i64; 4],
    },
    Shape(Box<shape::Facts>),
    /// Members of an OfficeArt group in source (paint) order.
    Group(Vec<group::Member>),
}

struct ResolvedDrawing {
    content: Content,
    reading_picture_source: Option<(Part, u32)>,
    /// Projected in paragraph flow instead of anchored (pseudo-inline).
    inline: bool,
    shape_id: u32,
    extent: [i64; 2],
    flip: [bool; 2],
    x_emu: i64,
    y_emu: i64,
    horizontal: &'static str,
    vertical: &'static str,
    /// Horizontal/vertical `wp:align` values replacing the offsets when set.
    align: [Option<&'static str>; 2],
    wrapping: u8,
    side: &'static str,
    behind: bool,
    locked: bool,
    distances: [u32; 4],
    in_cell: bool,
    overlap: bool,
    z_order: u32,
    occurrence: u32,
}

fn records<'a>(bytes: &'a [u8], budget: &mut usize) -> Result<Vec<Record<'a>>, String> {
    let mut position = 0;
    let mut result = Vec::new();
    while position < bytes.len() {
        let (record, end) = record_with_end(bytes, position, budget, "Word OfficeArt")?;
        position = end;
        result.push(record);
    }
    Ok(result)
}

/// A registered shape: (anchor index, [package order, document order],
/// shape or group container).
/// Anchor index, document order among top-level shapes and groups, and the
/// shape container.
type ContainerShape<'a> = (usize, u32, Record<'a>);

/// The anchored shapes and groups of one OfficeArtDgContainer's patriarch.
/// Groups are kept whole: their members use the group coordinate space.
fn container_shapes<'a>(
    drawing: Record<'a>,
    budget: &mut usize,
    nested_solver: &mut bool,
) -> Result<BTreeMap<u32, ContainerShape<'a>>, String> {
    let mut shapes = BTreeMap::new();
    for child in records(drawing.payload, budget)? {
        if child.kind != 0xf003 {
            continue;
        }
        if child.version != 15 {
            return Err(unsupported("invalid Word shape group"));
        }
        for shape in records(child.payload, budget)? {
            if shape.kind == 0xf005 {
                *nested_solver = true;
            }
            // A nested OfficeArtSpgrContainer is an anchored group: its first
            // OfficeArtSpContainer carries the group's FSP and client anchor
            // (MS-ODRAW 2.2.16). It is registered as a whole and resolved
            // only by the direct model.
            let (entry, header) = match shape.kind {
                0xf004 => (shape, shape),
                0xf003 if shape.version == 15 => {
                    // A malformed group stays unregistered: its anchor is
                    // then reported as omitted content, as before.
                    match records(shape.payload, budget)?.into_iter().next() {
                        Some(first) if first.kind == 0xf004 => (shape, first),
                        _ => continue,
                    }
                }
                _ => continue,
            };
            if header.version != 15 {
                return Err(unsupported("invalid Word floating shape container"));
            }
            let shape = entry;
            let mut id = None;
            let mut anchor_index = None;
            for property in records(header.payload, budget)? {
                match property.kind {
                    0xf00a if property.payload.len() == 8 => {
                        id = Some(u32_at(property.payload, 0)?);
                    }
                    0xf010 if property.payload.len() == 4 => {
                        anchor_index = usize::try_from(u32_at(property.payload, 0)? as i32).ok();
                    }
                    _ => {}
                }
            }
            if let (Some(id), Some(anchor_index)) = (id, anchor_index) {
                if shapes.len() >= 100_000 {
                    return Err(unsupported("Word floating shape budget exceeded"));
                }
                let order = shapes.len() as u32 + 1;
                if shapes.insert(id, (anchor_index, order, shape)).is_some() {
                    return Err(unsupported("duplicate Word floating shape identifier"));
                }
            }
        }
    }
    Ok(shapes)
}

/// Narrow automatic-picture class for explicit block relocation. Keep all
/// existing picture parsing/budget gates; unlike the historical passive image
/// projection, this route additionally classifies every property occurrence.
/// [MS-ODRAW] picture crop/pib and group-shape positioning are represented by
/// the retained image/anchor model. Stored contours and unrepresented effects
/// cannot be relabelled automatic or silently lost.
#[cfg(test)]
fn verify_reading_picture_properties(shape: Record<'_>, budget: &mut usize) -> Result<(), String> {
    reading_picture::acquire(shape, &reading_picture::Defaults::default(), budget).map(|_| ())
}

struct Placement {
    horizontal: u32,
    vertical: u32,
    relative_horizontal: Option<u32>,
    relative_vertical: Option<u32>,
    distances: [u32; 4],
    in_cell: bool,
    overlap: bool,
    hidden: bool,
    script: bool,
    z_order: Option<u32>,
}
impl Default for Placement {
    fn default() -> Self {
        Self {
            horizontal: 0,
            vertical: 0,
            relative_horizontal: None,
            relative_vertical: None,
            distances: [114300, 0, 114300, 0],
            in_cell: true,
            overlap: true,
            hidden: false,
            script: false,
            z_order: None,
        }
    }
}
impl Placement {
    fn apply(&mut self, property: Record<'_>, budget: &mut usize) -> Result<(), String> {
        let count = usize::from(property.instance);
        *budget = budget
            .checked_sub(count)
            .ok_or_else(|| unsupported("Word placement property budget exceeded"))?;
        let entries = property
            .payload
            .get(..count * 6)
            .ok_or_else(|| unsupported("truncated Word placement properties"))?;
        for p in entries.as_chunks::<6>().0 {
            let key = u16_at(p, 0)?;
            let value = u32_at(p, 2)?;
            match key {
                0x384..=0x387 => {
                    if value > i32::MAX as u32 {
                        return Err(unsupported("negative Word picture wrap distance"));
                    }
                    self.distances[usize::from(key - 0x384)] = value;
                }
                0x38f => self.horizontal = value,
                0x390 => self.relative_horizontal = Some(value),
                0x391 => self.vertical = value,
                0x392 => self.relative_vertical = Some(value),
                0x3aa if value != 0 => self.z_order = Some(value),
                0x3bf => {
                    for (bit, target) in [
                        (15, &mut self.in_cell),
                        (9, &mut self.overlap),
                        (1, &mut self.hidden),
                        (7, &mut self.script),
                    ] {
                        if value & (1 << (bit + 16)) != 0 {
                            *target = value & (1 << bit) != 0;
                        }
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }
}

/// Resolve MS-ODRAW posh/posv alignment (2.3.4.19/21) for the direct model.
///
/// The MS-DOC SPA origin (2.9.253 bx/by: margin, page, column/paragraph) is
/// the normative container of the rca rectangle, so an aligned position is
/// placed within that container. posrelh/posrelv only corroborate it: Word
/// writes them as 0 margin, 1 page, 2 text, 3 character/line, one less than
/// the enumeration published in MS-ODRAW 2.3.4.20/22 (1-4); every aligned and
/// absolute shape in the local private corpus satisfies bx = posrelh and
/// by = posrelv under that numbering, and an absent value is the documented
/// msoprhText/msoprvText default (column/paragraph). Any disagreement,
/// character/line-relative alignment, inside/outside page-parity alignment
/// and vertical alignment within a paragraph stay unsupported: each needs an
/// Office control before its container can be asserted.
fn direct_alignment(
    anchor: &Anchor,
    placement: &Placement,
) -> Result<[Option<&'static str>; 2], String> {
    let horizontal = match placement.horizontal {
        0 => None,
        1 => Some("left"),
        2 => Some("center"),
        3 => Some("right"),
        _ => {
            return Err(unsupported(
                "Word drawing uses page-parity or invalid horizontal alignment",
            ))
        }
    };
    if horizontal.is_some() {
        let origin = match anchor.horizontal {
            "margin" => 0,
            "page" => 1,
            _ => 2,
        };
        if placement.relative_horizontal.unwrap_or(2) != origin {
            return Err(unsupported(
                "Word aligned drawing disagrees with its horizontal SPA origin",
            ));
        }
    }
    let vertical = match placement.vertical {
        0 => None,
        1 => Some("top"),
        2 => Some("center"),
        3 => Some("bottom"),
        _ => {
            return Err(unsupported(
                "Word drawing uses page-parity or invalid vertical alignment",
            ))
        }
    };
    if vertical.is_some() {
        let origin = match anchor.vertical {
            "margin" => 0,
            "page" => 1,
            _ => {
                return Err(unsupported(
                    "Word drawing aligned within its paragraph is not supported",
                ))
            }
        };
        if placement.relative_vertical.unwrap_or(2) != origin {
            return Err(unsupported(
                "Word aligned drawing disagrees with its vertical SPA origin",
            ));
        }
    }
    Ok([horizontal, vertical])
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Anchor {
    pub cp: usize,
    pub shape_id: u32,
    /// [MS-DOC] Spa cTxbx: group-owned textbox content must not disappear.
    pub textbox_count: i32,
    pub rect: [i32; 4],
    pub horizontal: &'static str,
    pub vertical: &'static str,
    pub wrapping: u8,
    pub side: &'static str,
    pub behind: bool,
    pub locked: bool,
}

#[cfg(test)]
pub(super) fn anchors(word: &[u8], table: &[u8], main_units: usize) -> Result<Vec<Anchor>, String> {
    anchors_in(word, table, Part::Main, main_units)
}

fn anchors_in(
    word: &[u8],
    table: &[u8],
    part: Part,
    main_units: usize,
) -> Result<Vec<Anchor>, String> {
    // FibRgFcLcb97 fields 40/41: main and header shape PLCs. Each part's
    // anchors address only its own story; they are never reassigned.
    let fib = part.anchor_table();
    if word.len() < fib + 8 {
        return Ok(Vec::new());
    }
    let size = u32_at(word, fib + 4)? as usize;
    if size == 0 {
        return Ok(Vec::new());
    }
    if size < 4 || !(size - 4).is_multiple_of(30) {
        return Err(unsupported("invalid Word floating-anchor PLC size"));
    }
    let count = (size - 4) / 30;
    if count > 100_000 {
        return Err(unsupported("Word floating-anchor budget exceeded"));
    }
    let offset = u32_at(word, fib)? as usize;
    let bytes = table
        .get(offset..)
        .and_then(|b| b.get(..size))
        .ok_or_else(|| unsupported("Word floating-anchor PLC out of bounds"))?;
    let mut result = Vec::with_capacity(count);
    let mut ids = std::collections::BTreeSet::new();
    for i in 0..count {
        let cp = u32_at(bytes, i * 4)? as usize;
        let next = u32_at(bytes, (i + 1) * 4)? as usize;
        if cp > main_units || cp >= next {
            return Err(unsupported("invalid Word floating-anchor CP order"));
        }
        let record = &bytes[(count + 1) * 4 + i * 26..][..26];
        let shape_id = u32_at(record, 0)?;
        if !ids.insert(shape_id) {
            return Err(unsupported(
                "duplicate Word floating-anchor shape identifier",
            ));
        }
        let flags = u16_at(record, 20)?;
        let horizontal = match (flags >> 1) & 3 {
            0 => "margin",
            1 => "page",
            2 => "column",
            _ => return Err(unsupported("invalid Word floating horizontal origin")),
        };
        let vertical = match (flags >> 3) & 3 {
            0 => "margin",
            1 => "page",
            2 => "paragraph",
            _ => return Err(unsupported("invalid Word floating vertical origin")),
        };
        let wrapping = ((flags >> 5) & 15) as u8;
        if wrapping > 5 {
            return Err(unsupported("invalid Word floating wrap mode"));
        }
        let side = if matches!(wrapping, 1 | 3) {
            "bothSides"
        } else {
            match (flags >> 9) & 15 {
                0 => "bothSides",
                1 => "left",
                2 => "right",
                3 => "largest",
                _ => return Err(unsupported("invalid Word floating wrap side")),
            }
        };
        result.push(Anchor {
            cp,
            shape_id,
            textbox_count: u32_at(record, 22)? as i32,
            rect: [
                u32_at(record, 4)? as i32,
                u32_at(record, 8)? as i32,
                u32_at(record, 12)? as i32,
                u32_at(record, 16)? as i32,
            ],
            horizontal,
            vertical,
            wrapping,
            side,
            behind: wrapping == 3 && flags & 0x4000 != 0,
            locked: flags & 0x8000 != 0,
        });
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u16, options: u16, body: &[u8]) -> Vec<u8> {
        [
            options.to_le_bytes().as_slice(),
            &kind.to_le_bytes(),
            &(body.len() as u32).to_le_bytes(),
            body,
        ]
        .concat()
    }
    #[test]
    fn line_connectors_require_unbound_unique_solver_rules() {
        fn drawing(
            kind: u16,
            endpoints: [u32; 2],
            duplicate: bool,
            nested: bool,
            ignored_shape: bool,
        ) -> Vec<u8> {
            let line = record(
                0xf004,
                15,
                &[
                    record(
                        0xf00a,
                        (kind << 4) | 2,
                        &[7u32.to_le_bytes(), 0xb00u32.to_le_bytes()].concat(),
                    ),
                    record(0xf010, 0, &0u32.to_le_bytes()),
                ]
                .concat(),
            );
            // Unavailable endpoint shapes have ignored site indexes
            // (MS-ODRAW 2.2.36); a sentinel is not required.
            let rule = record(
                0xf012,
                1,
                &[
                    3u32.to_le_bytes(),
                    endpoints[0].to_le_bytes(),
                    endpoints[1].to_le_bytes(),
                    7u32.to_le_bytes(),
                    4u32.to_le_bytes(),
                    9u32.to_le_bytes(),
                ]
                .concat(),
            );
            let mut other = rule.clone();
            other[8..12].copy_from_slice(&4u32.to_le_bytes()); // ruid
            other[12..20].fill(0); // unavailable endpoints
            other[20..24].copy_from_slice(&8u32.to_le_bytes()); // spidC
                                                                // MS-ODRAW 2.2.18 recInstance counts file blocks, not zero.
            let solver = record(
                0xf005,
                (2 << 4) | 15,
                &[rule.clone(), if duplicate { rule } else { other }].concat(),
            );
            let other_shape = record(
                0xf004,
                15,
                &record(
                    0xf00a,
                    (32 << 4) | 2,
                    &[8u32.to_le_bytes(), 0xb00u32.to_le_bytes()].concat(),
                ),
            );
            let ignored = if ignored_shape {
                record(0xf004, 15, &record(0xf00a, (1 << 4) | 2, &[0; 4]))
            } else {
                Vec::new()
            };
            record(
                0xf002,
                15,
                &[
                    record(
                        0xf003,
                        15,
                        &[
                            line,
                            other_shape,
                            ignored,
                            if nested { solver.clone() } else { Vec::new() },
                        ]
                        .concat(),
                    ),
                    if nested { Vec::new() } else { solver },
                ]
                .concat(),
            )
        }
        fn resolve(bytes: Vec<u8>) -> Result<ResolvedDrawing, String> {
            let (mut word, mut table) = input((1 << 1) | (2 << 3) | (3 << 5));
            table[8..12].copy_from_slice(&7u32.to_le_bytes());
            // A horizontal authored line: a zero height must remain valid.
            table[24..28].copy_from_slice(&200i32.to_le_bytes());
            let art = [record(0xf000, 15, &[]), vec![0], bytes].concat();
            word[0x22a..0x22e].copy_from_slice(&(table.len() as u32).to_le_bytes());
            word[0x22e..0x232].copy_from_slice(&(art.len() as u32).to_le_bytes());
            table.extend(art);
            Store::read(&word, &table, 20)?
                .resolve(Part::Main, 12)?
                .ok_or_else(|| "line was omitted".into())
        }
        fn add_empty_solvers(bytes: Vec<u8>, count: usize) -> Vec<u8> {
            let (drawing, _) = record_with_end(&bytes, 0, &mut 1000, "test").unwrap();
            let mut payload = drawing.payload.to_vec();
            for _ in 0..count {
                payload.extend(record(0xf005, 15, &[]));
            }
            record(0xf002, 15, &payload)
        }
        let line = resolve(drawing(20, [0, 0], false, false, false)).unwrap();
        assert_eq!(line.extent, [254000, 0]);
        let Content::Shape(facts) = line.content else {
            panic!("line shape was not retained")
        };
        assert_eq!(facts.preset, Some("line"));
        assert!(facts.fill.is_none());
        assert!(facts.line.is_some());
        let mut wrong_count = drawing(20, [0, 0], false, false, false);
        let position = wrong_count
            .windows(8)
            .position(|header| header == [0x2f, 0, 0x05, 0xf0, 64, 0, 0, 0])
            .unwrap();
        wrong_count[position..position + 2].copy_from_slice(&15u16.to_le_bytes());
        assert!(resolve(wrong_count)
            .err()
            .unwrap()
            .contains("solver count mismatch"));
        // DgContainer provides solvers1 and solvers2 (MS-ODRAW 2.2.13).
        assert!(resolve(add_empty_solvers(
            drawing(20, [0, 0], false, false, false),
            1
        ))
        .is_ok());
        assert!(resolve(add_empty_solvers(
            drawing(20, [0, 0], false, false, false),
            2
        ))
        .err()
        .unwrap()
        .contains("invalid Word line connector solver"));

        assert!(resolve(drawing(20, [8, 0], false, false, false))
            .err()
            .unwrap()
            .contains("line connector has bound endpoints"));
        assert!(resolve(drawing(20, [0, 0], true, false, false))
            .err()
            .unwrap()
            .contains("duplicate line connector rule"));
        assert!(resolve(drawing(20, [0, 0], false, true, false))
            .err()
            .unwrap()
            .contains("solver scope"));
        // Established kind32 never invokes the new line solver policy or
        // reparses formerly ignored malformed, unanchored FSP contents.
        assert!(resolve(drawing(32, [8, 0], true, true, true)).is_ok());

        // Distinct anchored groups must not acquire the same child identity:
        // otherwise a cached solver target could describe two different lines.
        fn group(id: u32, anchor_index: u32) -> Vec<u8> {
            let head = record(
                0xf004,
                15,
                &[
                    record(
                        0xf00a,
                        2,
                        &[id.to_le_bytes(), 0x201u32.to_le_bytes()].concat(),
                    ),
                    record(
                        0xf009,
                        1,
                        &[0i32, 0, 100, 100]
                            .into_iter()
                            .flat_map(i32::to_le_bytes)
                            .collect::<Vec<_>>(),
                    ),
                    record(0xf010, 0, &anchor_index.to_le_bytes()),
                ]
                .concat(),
            );
            let line = record(
                0xf004,
                15,
                &[
                    record(
                        0xf00a,
                        (20 << 4) | 2,
                        &[7u32.to_le_bytes(), 0xb02u32.to_le_bytes()].concat(),
                    ),
                    record(
                        0xf00f,
                        0,
                        &[0i32, 0, 100, 0]
                            .into_iter()
                            .flat_map(i32::to_le_bytes)
                            .collect::<Vec<_>>(),
                    ),
                ]
                .concat(),
            );
            record(0xf003, 15, &[head, line].concat())
        }
        // Every acquired member contributes to a new line group's scope,
        // including a valid picture whose header is outside the Dg solver.
        fn mixed_picture_group(nested_solver: bool) -> Result<ResolvedDrawing, String> {
            let (mut word, mut table) = drawing_input(0xa00, 0);
            let start = u32_at(&word, 0x22a)? as usize;
            let (_, end) = record_with_end(&table[start..], 0, &mut 1000, "test")?;
            let dgg = table[start..start + end].to_vec();
            let picture = record(
                0xf004,
                15,
                &[
                    record(
                        0xf00a,
                        (75 << 4) | 2,
                        &[9u32.to_le_bytes(), 0xa02u32.to_le_bytes()].concat(),
                    ),
                    record(
                        0xf00b,
                        (1 << 4) | 3,
                        &[0x4104u16.to_le_bytes().as_slice(), &1u32.to_le_bytes()].concat(),
                    ),
                    record(
                        0xf00f,
                        0,
                        &[0i32, 0, 20, 20]
                            .into_iter()
                            .flat_map(i32::to_le_bytes)
                            .collect::<Vec<_>>(),
                    ),
                    if nested_solver {
                        record(
                            0xf005,
                            (1 << 4) | 15,
                            &record(
                                0xf012,
                                1,
                                &[77u32, 8, 0, 7, 0, 0]
                                    .into_iter()
                                    .flat_map(u32::to_le_bytes)
                                    .collect::<Vec<_>>(),
                            ),
                        )
                    } else {
                        Vec::new()
                    },
                ]
                .concat(),
            );
            let source_group = group(40, 0);
            let (source_group, _) = record_with_end(&source_group, 0, &mut 1000, "test")?;
            let mixed = record(0xf003, 15, &[source_group.payload, &picture].concat());
            let art = [
                dgg,
                vec![0],
                record(0xf002, 15, &record(0xf003, 15, &mixed)),
            ]
            .concat();
            table.truncate(start);
            table[8..12].copy_from_slice(&40u32.to_le_bytes());
            word[0x22e..0x232].copy_from_slice(&(art.len() as u32).to_le_bytes());
            table.extend(art);
            Store::read(&word, &table, 20)?
                .resolve(Part::Main, 12)?
                .ok_or_else(|| "mixed group was omitted".into())
        }
        let positive = mixed_picture_group(false).unwrap();
        let Content::Group(members) = positive.content else {
            panic!("mixed group was not retained")
        };
        assert_eq!(
            members.iter().map(|member| member.spid).collect::<Vec<_>>(),
            [7, 9]
        );
        assert!(
            matches!(&members[0].content, Content::Shape(facts) if facts.preset == Some("line"))
        );
        assert!(matches!(
            &members[1].content,
            Content::Picture { image_index: 0, .. }
        ));
        assert!(mixed_picture_group(true)
            .err()
            .expect("picture member solver scope must reject the complete group")
            .contains("solver scope"));

        let (mut word, one_anchor) = input((1 << 1) | (2 << 3) | (3 << 5));
        let mut a = one_anchor[8..].to_vec();
        a[..4].copy_from_slice(&40u32.to_le_bytes());
        let mut b = a.clone();
        b[..4].copy_from_slice(&41u32.to_le_bytes());
        let mut table = [
            12u32.to_le_bytes(),
            14u32.to_le_bytes(),
            30u32.to_le_bytes(),
        ]
        .concat();
        table.extend(a);
        table.extend(b);
        word[0x1de..0x1e2].copy_from_slice(&(table.len() as u32).to_le_bytes());
        let art = [
            record(0xf000, 15, &[]),
            vec![0],
            record(
                0xf002,
                15,
                &record(0xf003, 15, &[group(40, 0), group(41, 1)].concat()),
            ),
        ]
        .concat();
        word[0x22a..0x22e].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0x22e..0x232].copy_from_slice(&(art.len() as u32).to_le_bytes());
        table.extend(art);
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.resolve(Part::Main, 12).unwrap().is_some());
        assert!(store
            .resolve(Part::Main, 14)
            .err()
            .unwrap()
            .contains("ambiguous Word line connector shape ownership"));
    }

    fn drawing_input(shape_flags: u32, group_flags: u32) -> (Vec<u8>, Vec<u8>) {
        drawing_with_options(shape_flags, group_flags, &[])
    }
    fn drawing_with_options(
        shape_flags: u32,
        group_flags: u32,
        options: &[(u16, u32)],
    ) -> (Vec<u8>, Vec<u8>) {
        let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
        png.extend(2u32.to_be_bytes());
        png.extend(3u32.to_be_bytes());
        png.extend([8, 2, 0, 0, 0, 0, 0, 0, 0]);
        let blip = record(0xf01e, 0x6e0 << 4, &[vec![0; 17], png].concat());
        drawing_with_blip(shape_flags, group_flags, options, blip, 6)
    }
    fn drawing_with_blip(
        shape_flags: u32,
        group_flags: u32,
        options: &[(u16, u32)],
        blip: Vec<u8>,
        kind: u8,
    ) -> (Vec<u8>, Vec<u8>) {
        let (mut word, mut table) = input((1 << 1) | (1 << 3) | (2 << 5));
        word.resize(1024, 0);
        word.extend(&blip);
        let mut bse = vec![0; 36];
        bse[0] = kind;
        bse[1] = kind;
        bse[24] = 1;
        bse[20..24].copy_from_slice(&(blip.len() as u32).to_le_bytes());
        bse[28..32].copy_from_slice(&1024u32.to_le_bytes());
        let group = record(
            0xf000,
            15,
            &record(
                0xf001,
                31,
                &record(0xf007, (u16::from(kind) << 4) | 2, &bse),
            ),
        );
        let mut props = [
            0x4104u16.to_le_bytes().as_slice(),
            &1u32.to_le_bytes(),
            &0x3bfu16.to_le_bytes(),
            &group_flags.to_le_bytes(),
        ]
        .concat();
        for (key, value) in options {
            props.extend(key.to_le_bytes());
            props.extend(value.to_le_bytes());
        }
        let shape = record(
            0xf004,
            15,
            &[
                record(
                    0xf00a,
                    (75 << 4) | 2,
                    &[1027u32.to_le_bytes(), shape_flags.to_le_bytes()].concat(),
                ),
                record(0xf00b, (((2 + options.len()) as u16) << 4) | 3, &props),
                record(0xf010, 0, &0u32.to_le_bytes()),
            ]
            .concat(),
        );
        let art = [
            group,
            vec![0],
            record(0xf002, 15, &record(0xf003, 15, &shape)),
        ]
        .concat();
        word[0x22a..0x22e].copy_from_slice(&(table.len() as u32).to_le_bytes());
        word[0x22e..0x232].copy_from_slice(&(art.len() as u32).to_le_bytes());
        table.extend(art);
        (word, table)
    }

    /// The direct picture at CP 12, or None when nothing is projected.
    fn picture(store: &mut Store<'_>) -> Result<Option<direct::DirectFloatingPicture>, String> {
        store.direct_picture(12, &mut usize::MAX.clone())
    }

    fn resources(store: Store<'_>) -> Vec<super::super::pictures::DirectPictureResource> {
        let mut resources = Vec::new();
        store
            .append_direct_resources(&mut resources, &mut usize::MAX.clone())
            .unwrap();
        resources
    }

    fn reading_picture_with_options(options: &[(u16, u32)]) -> (Vec<u8>, Vec<u8>) {
        let (word, mut table) = drawing_with_options(0xa00, 0, options);
        // The fixture's PlcfSpa begins at table offset zero: CP array (8),
        // spid/rectangle (20), then SPA flags. Select wrapTight so the strict
        // reading-picture property gate is exercised before BLIP selection.
        let flags = (1u16 << 1) | (1 << 3) | (4 << 5);
        table[28..30].copy_from_slice(&flags.to_le_bytes());
        (word, table)
    }

    #[test]
    fn reading_picture_edit_locks_preserve_projection_and_resources() {
        let project = |options: &[(u16, u32)]| {
            let (word, table) = reading_picture_with_options(options);
            let mut store = Store::read(&word, &table, 20).unwrap();
            store.set_reading_relocation(true);
            let picture = picture(&mut store).unwrap().unwrap();
            assert_eq!(
                picture
                    .image
                    .anchor_acquisition
                    .as_ref()
                    .unwrap()
                    .wrap
                    .authored_kinds,
                ["wrapTight"]
            );
            assert!(!store.omitted);
            (
                serde_json::to_value(picture.image).unwrap(),
                picture.occurrence_id,
                resources(store),
            )
        };
        let control = project(&[]);
        // MS-ODRAW 2.3.20.1: zero, explicit editing locks, and all bits set
        // (including ignored unused bits) cannot change the painted picture.
        for value in [0, 0x01ff_01ff, u32::MAX] {
            assert_eq!(project(&[(0x007f, value)]), control);
        }
    }

    #[test]
    fn reading_picture_edit_locks_do_not_admit_nonordinary_or_painting_properties() {
        for property in [
            (0x407f, 0), // fBid is forbidden for Protection Boolean Properties.
            (0x807f, 0), // Even an empty complex payload is not this structure.
            (0x007e, 0), // No neighboring protection-property range admission.
            (0x0200, 1), // A shadow/effect is still outside the reading subset.
            (0x01bf, 1), // Non-default fill painting remains unsupported.
        ] {
            let (word, table) = reading_picture_with_options(&[property]);
            let mut store = Store::read(&word, &table, 20).unwrap();
            store.set_reading_relocation(true);
            assert!(picture(&mut store).is_err(), "property {property:?}");
            assert!(
                store.images.is_empty(),
                "refusal must precede BLIP acquisition"
            );
            assert!(resources(store).is_empty());
        }
    }

    #[test]
    fn reading_picture_edit_locks_keep_table_and_budget_validation() {
        let entry = [0x007fu16.to_le_bytes().as_slice(), &0u32.to_le_bytes()].concat();
        for kind in [0xf00b, 0xf122] {
            let verify = |options, body: &[u8], budget: &mut usize| {
                let property_table = record(kind, options, body);
                verify_reading_picture_properties(
                    Record {
                        version: 15,
                        instance: 0,
                        kind: 0xf004,
                        payload: &property_table,
                    },
                    budget,
                )
            };
            assert!(verify(0x13, &entry[..5], &mut 10)
                .unwrap_err()
                .contains("truncated OfficeArt properties"));
            // Zero first exhausts the enclosing record walker. One permits
            // the property table record and then exhausts its property visit.
            assert!(verify(0x13, &entry, &mut 0)
                .unwrap_err()
                .contains("too many Word OfficeArt records"));
            assert!(verify(0x13, &entry, &mut 1)
                .unwrap_err()
                .contains("OfficeArt property work budget exceeded"));
            assert!(verify(0x12, &entry, &mut 10)
                .unwrap_err()
                .contains("invalid OfficeArt property table"));
            let complex = [0x807fu16.to_le_bytes().as_slice(), &1u32.to_le_bytes()].concat();
            assert!(verify(0x13, &complex, &mut 10)
                .unwrap_err()
                .contains("truncated OfficeArt complex shape property"));
            let extra_tail = [entry.as_slice(), &[0]].concat();
            assert!(verify(0x13, &extra_tail, &mut 10)
                .unwrap_err()
                .contains("unexpected OfficeArt property data"));
        }
        let (word, table) = reading_picture_with_options(&[(0x007f, 0)]);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.set_reading_relocation(true);
        assert_eq!(
            store.direct_picture(12, &mut 0).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
    }
    // Independently authored FOPT entries. Values remain explicit so malformed
    // length controls exercise the production property-table walker.
    fn reading_metadata_fopt(entries: &[(u16, u32, &[u8])], kind: u16) -> Vec<u8> {
        let mut entries = entries.to_vec();
        entries.sort_by_key(|entry| entry.0 & 0x3fff);
        let mut body = Vec::new();
        for (opid, value, _) in &entries {
            body.extend(opid.to_le_bytes());
            body.extend(value.to_le_bytes());
        }
        for (opid, _, bytes) in &entries {
            if opid & 0x8000 != 0 {
                body.extend(*bytes);
            }
        }
        record(kind, ((entries.len() as u16) << 4) | 3, &body)
    }

    fn reading_picture_with_metadata(
        local: &[(u16, u32, &[u8])],
        document: &[(u16, u32, &[u8])],
        tertiary: &[(u16, u32, &[u8])],
    ) -> (Vec<u8>, Vec<u8>) {
        let (mut word, mut table) = reading_picture_with_options(&[]);
        let art_start = u32::from_le_bytes(word[0x22a..0x22e].try_into().unwrap()) as usize;
        let art = &table[art_start..];
        let group_size = 8 + u32::from_le_bytes(art[4..8].try_into().unwrap()) as usize;
        let group = records(&art[..group_size], &mut 1000).unwrap()[0];
        let drawing = records(&art[group_size + 1..], &mut 1000).unwrap()[0];
        let shape_group = records(drawing.payload, &mut 1000).unwrap()[0];
        let shape = records(shape_group.payload, &mut 1000).unwrap()[0];
        let mut entries = vec![(0x4104, 1, &[][..]), (0x03bf, 0, &[][..])];
        entries.extend_from_slice(local);
        let shape_body = records(shape.payload, &mut 1000)
            .unwrap()
            .into_iter()
            .map(|child| {
                if child.kind == 0xf00b {
                    reading_metadata_fopt(&entries, 0xf00b)
                } else {
                    record(
                        child.kind,
                        (child.instance << 4) | u16::from(child.version),
                        child.payload,
                    )
                }
            })
            .collect::<Vec<_>>()
            .concat();
        let mut group_body = group.payload.to_vec();
        if !document.is_empty() {
            group_body.extend(reading_metadata_fopt(document, 0xf00b));
        }
        if !tertiary.is_empty() {
            group_body.extend(reading_metadata_fopt(tertiary, 0xf122));
        }
        let next_art = [
            record(0xf000, 15, &group_body),
            vec![0],
            record(
                0xf002,
                15,
                &record(0xf003, 15, &record(0xf004, 15, &shape_body)),
            ),
        ]
        .concat();
        word[0x22e..0x232].copy_from_slice(&(next_art.len() as u32).to_le_bytes());
        table.truncate(art_start);
        table.extend(next_art);
        (word, table)
    }

    fn project_reading_metadata(
        local: &[(u16, u32, &[u8])],
        document: &[(u16, u32, &[u8])],
        tertiary: &[(u16, u32, &[u8])],
    ) -> (
        serde_json::Value,
        String,
        Vec<super::super::pictures::DirectPictureResource>,
    ) {
        let (word, table) = reading_picture_with_metadata(local, document, tertiary);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.set_reading_relocation(true);
        let projected = picture(&mut store).unwrap().unwrap();
        assert!(!store.omitted);
        (
            serde_json::to_value(projected.image).unwrap(),
            projected.occurrence_id,
            resources(store),
        )
    }

    fn assert_reading_metadata_refused(
        local: &[(u16, u32, &[u8])],
        document: &[(u16, u32, &[u8])],
        tertiary: &[(u16, u32, &[u8])],
    ) {
        let (word, table) = reading_picture_with_metadata(local, document, tertiary);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.set_reading_relocation(true);
        assert!(picture(&mut store).is_err());
        assert!(
            store.images.is_empty(),
            "property refusal precedes required BLIP acquisition"
        );
        assert!(resources(store).is_empty());
    }

    fn without_reading_metadata(mut image: serde_json::Value) -> serde_json::Value {
        image["__anchorAcquisition"]
            .as_object_mut()
            .unwrap()
            .remove("nativePictureMetadata");
        image
    }

    #[test]
    fn reading_picture_metadata_retains_passive_names_without_changing_image_or_resources() {
        let label = [b'L', 0, 0, 0];
        let description = [b'D', 0, 0, 0];
        let old = [b'O', 0, 0, 0];
        let control = project_reading_metadata(&[], &[], &[]);
        let projected = project_reading_metadata(
            &[
                (0xc105, 4, &label),
                (0x0106, 0, &[]),
                (0x8380, 4, &label),
                (0xc381, 4, &description),
            ],
            &[(0x8380, 4, &old)],
            &[],
        );
        let metadata = &projected.0["__anchorAcquisition"]["nativePictureMetadata"];
        assert_eq!(metadata["blipName"]["text"], "L");
        assert_eq!(metadata["shapeName"]["text"], "L");
        assert_eq!(metadata["description"]["text"], "D");
        assert_eq!(metadata["shapeName"]["scope"], "shape");
        assert_eq!(
            metadata["description"]["rawBytes"],
            serde_json::json!([68, 0, 0, 0])
        );
        assert_eq!(without_reading_metadata(projected.0), control.0);
        assert_eq!(projected.1, control.1);
        assert_eq!(projected.2, control.2);
        let empty_names = project_reading_metadata(
            &[(0x0105, 0, &[]), (0x4380, 0, &[]), (0x0381, 0, &[])],
            &[],
            &[],
        );
        assert_eq!(
            empty_names.0["__anchorAcquisition"]["nativePictureMetadata"]["shapeName"]["text"],
            ""
        );
        assert_eq!(without_reading_metadata(empty_names.0), control.0);
        assert_eq!(empty_names.2, control.2);
    }

    #[test]
    fn reading_picture_metadata_resolves_document_boolean_members_before_accepting_picture() {
        let control = project_reading_metadata(&[], &[], &[]);
        // Raw unused local gray=true does not defeat the default-false contract.
        assert_eq!(
            project_reading_metadata(&[(0x013f, 4, &[])], &[], &[]),
            control
        );
        // An authored local gray=false overrides an authored document gray=true.
        assert_eq!(
            project_reading_metadata(
                &[(0x013f, 0x0004_0000, &[])],
                &[(0x013f, 0x0004_0004, &[])],
                &[]
            ),
            control
        );
        // Unused local bits do not override an active document owner.
        assert_reading_metadata_refused(&[(0x013f, 0, &[])], &[(0x013f, 0x0004_0004, &[])], &[]);
        // Primary and tertiary document properties are not guessed chronological owners.
        assert_reading_metadata_refused(
            &[],
            &[(0x013f, 0x0004_0004, &[])],
            &[(0x013f, 0x0004_0000, &[])],
        );
        // A local whole-property false cannot erase a different active member.
        assert_reading_metadata_refused(
            &[(0x013f, 0x0004_0000, &[])],
            &[(0x013f, 0x0006_0006, &[])],
            &[],
        );
        assert_reading_metadata_refused(&[(0x413f, 0, &[])], &[], &[]);
        assert_reading_metadata_refused(&[(0x813f, 0, &[])], &[], &[]);
    }

    #[test]
    fn reading_picture_metadata_retains_only_proved_inactive_opaque_or_index_carriers() {
        let control = project_reading_metadata(&[], &[], &[]);
        // Complete empty complex tails remain complex carriers, not zero indices.
        for document in [false, true] {
            let carriers = [(0x8186, 0, &[][..]), (0x81c5, 2, &[0x41, 0x42][..])];
            let owners = [
                (0x01bf, 0x0010_0000, &[][..]),
                (0x01ff, 0x0008_0000, &[][..]),
            ];
            let all = [carriers.as_slice(), owners.as_slice()].concat();
            let projected = if document {
                project_reading_metadata(&owners, &carriers, &[])
            } else {
                project_reading_metadata(&all, &[], &[])
            };
            let metadata = &projected.0["__anchorAcquisition"]["nativePictureMetadata"];
            assert_eq!(
                metadata["inactiveFillCarrier"]["retention"],
                "inactiveOpaqueNotDecoded"
            );
            assert_eq!(
                metadata["inactiveLineCarrier"]["rawBytes"],
                serde_json::json!([65, 66])
            );
            assert_eq!(
                metadata["inactiveLineCarrier"]["scope"],
                if document { "documentDefault" } else { "shape" }
            );
            assert_eq!(without_reading_metadata(projected.0), control.0);
            assert_eq!(projected.1, control.1);
            assert_eq!(projected.2, control.2);
        }
        let ignored = project_reading_metadata(&[(0x4186, 0, &[]), (0x41c5, 0, &[])], &[], &[]);
        assert_eq!(
            ignored.0["__anchorAcquisition"]["nativePictureMetadata"]["inactiveFillCarrier"]
                ["retention"],
            "ignoredZeroIndex"
        );
        let unresolved =
            project_reading_metadata(&[(0x4186, 17, &[]), (0x01bf, 0x0010_0000, &[])], &[], &[]);
        assert_eq!(
            unresolved.0["__anchorAcquisition"]["nativePictureMetadata"]["inactiveFillCarrier"]
                ["retention"],
            "inactiveIndexNotResolved"
        );
        assert_eq!(without_reading_metadata(unresolved.0), control.0);
        assert_eq!(unresolved.2, control.2);
    }

    #[test]
    fn reading_picture_metadata_does_not_guess_inactive_paint_or_skip_other_effects() {
        let empty = [(0x8186, 0, &[][..]), (0x81c5, 0, &[][..])];
        assert_reading_metadata_refused(&[], &empty, &[]);
        assert_reading_metadata_refused(&[(0x01bf, 0, &[]), (0x01ff, 0, &[])], &empty, &[]);
        let inactive = [
            (0x01bf, 0x0010_0000, &[][..]),
            (0x01ff, 0x0008_0000, &[][..]),
        ];
        // Active document defaults are overridden only by the corresponding authored bits.
        let active = [
            (0x01bf, 0x0010_0010, &[][..]),
            (0x01ff, 0x0008_0008, &[][..]),
        ];
        let defaults = [empty.as_slice(), active.as_slice()].concat();
        assert!(
            project_reading_metadata(&inactive, &defaults, &[]).0["__anchorAcquisition"]
                ["nativePictureMetadata"]
                .is_object()
        );
        assert_reading_metadata_refused(&[(0x01bf, 0x0010_0000, &[])], &defaults, &[]);
        let dashed = [
            empty.as_slice(),
            &[
                (0x01bf, 0x0010_0000, &[][..]),
                (0x01ff, 0x0009_0001, &[][..]),
            ],
        ]
        .concat();
        assert_reading_metadata_refused(&dashed, &[], &[]);
        let inherited_dash = [empty.as_slice(), &[(0x01ff, 0x0001_0001, &[][..])]].concat();
        assert_reading_metadata_refused(&inactive, &inherited_dash, &[]);
        let explicit_no_dash = [
            (0x01bf, 0x0010_0000, &[][..]),
            (0x01ff, 0x0009_0000, &[][..]),
        ];
        assert!(
            project_reading_metadata(&explicit_no_dash, &inherited_dash, &[]).0
                ["__anchorAcquisition"]["nativePictureMetadata"]
                .is_object()
        );
        let reserved = [
            empty.as_slice(),
            &[
                (0x01bf, 0x0010_0000, &[][..]),
                (0x01ff, 0x0008_0080, &[][..]),
            ],
        ]
        .concat();
        assert_reading_metadata_refused(&reserved, &[], &[]);
        for effect in [
            (0x0180, 3, &[][..]),
            (0x01c4, 1, &[][..]),
            (0x0200, 1, &[][..]),
            (0x0107, 1, &[][..]),
            (0xc186, 1, &[0x41][..]),
            (0x0186, 0, &[][..]),
        ] {
            let props = [empty.as_slice(), inactive.as_slice(), &[effect]].concat();
            assert_reading_metadata_refused(&props, &[], &[]);
        }
    }

    #[test]
    fn reading_picture_metadata_rejects_links_invalid_names_and_late_tail_errors_atomically() {
        for (opid, value, bytes) in [
            (0x0106, 1, &[][..]),
            (0x0106, 2, &[][..]),
            (0x0106, 3, &[][..]),
            (0x0106, 4, &[][..]),
            (0x0106, 8, &[][..]),
            (0x0106, 9, &[][..]),
            (0x0105, 1, &[][..]),
            (0x8380, 1, &[0][..]),
            (0x8381, 2, &[0x41, 0][..]),
            (0x8381, 4, &[0x00, 0xd8, 0, 0][..]),
            (0x8381, 4, &[0, 0, 0, 0][..]),
            (0x8381, 4, &[0, 0][..]),
        ] {
            assert_reading_metadata_refused(&[(opid, value, bytes)], &[], &[]);
        }
        assert_reading_metadata_refused(
            &[(0x8380, 4, &[b'A', 0, 0, 0]), (0x8380, 4, &[b'B', 0, 0, 0])],
            &[],
            &[],
        );
        // A valid local picture cannot hide a malformed document-default tail.
        assert_reading_metadata_refused(&[], &[(0x8186, 2, &[])], &[]);
        assert_reading_metadata_refused(&[], &[(0x8380, 4, &[0, 0])], &[]);
        assert_reading_metadata_refused(&[], &[(0x0200, 0, &[])], &[]);
        // Duplicate Dgg primary tables are ambiguous even when their scalar
        // values agree; exercise actual registration, not an injected flag.
        let (mut word, mut table) = reading_picture_with_metadata(&[], &[], &[]);
        let start = u32::from_le_bytes(word[0x22a..0x22e].try_into().unwrap()) as usize;
        let art = &table[start..];
        let group_size = 8 + u32::from_le_bytes(art[4..8].try_into().unwrap()) as usize;
        let group = records(&art[..group_size], &mut 100).unwrap()[0];
        let option = reading_metadata_fopt(&[(0x013f, 0, &[])], 0xf00b);
        let next_art = [
            record(0xf000, 15, &[group.payload, &option, &option].concat()),
            art[group_size..].to_vec(),
        ]
        .concat();
        word[0x22e..0x232].copy_from_slice(&(next_art.len() as u32).to_le_bytes());
        table.truncate(start);
        table.extend(next_art);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.set_reading_relocation(true);
        assert!(picture(&mut store).is_err());
        assert!(store.images.is_empty());
        assert!(resources(store).is_empty());
    }

    #[test]
    fn reading_picture_metadata_owned_projection_respects_output_budget_and_strict_policy() {
        let name = [b'L', 0, 0, 0];
        let props = [
            (0x8380, 4, &name[..]),
            (0x8186, 0, &[][..]),
            (0x01bf, 0x0010_0000, &[][..]),
        ];
        let (word, table) = reading_picture_with_metadata(&props, &[], &[]);
        let mut reading = Store::read(&word, &table, 20).unwrap();
        reading.set_reading_relocation(true);
        assert_eq!(
            reading.direct_picture(12, &mut 0).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        let mut strict = Store::read(&word, &table, 20).unwrap();
        assert!(
            picture(&mut strict).unwrap().is_none(),
            "metadata does not authorize missing-contour strict projection"
        );
        assert!(strict.omitted);
    }

    #[test]
    fn reading_picture_complex_bid_retains_raw_flag_without_changing_visible_projection() {
        let control = project_reading_metadata(&[], &[], &[]);
        let owners = [
            (0x01bf, 0x0010_0000, &[][..]),
            (0x01ff, 0x0008_0000, &[][..]),
        ];
        // MS-ODRAW 2.2.8: fBid is ignored when fComplex is set; op remains a length.
        // Keep the authored flag in the sidecar, rather than rewriting the input.
        for (id, field) in [
            (0x0186, "inactiveFillCarrier"),
            (0x01c5, "inactiveLineCarrier"),
        ] {
            for bytes in [&[][..], &[0x41, 0x42, 0x43][..]] {
                for document in [false, true] {
                    for bid in [0, 0x4000] {
                        let opid = id | 0x8000 | bid;
                        let carrier = [(opid, bytes.len() as u32, bytes)];
                        let local = [owners.as_slice(), carrier.as_slice()].concat();
                        let projected = if document {
                            project_reading_metadata(&owners, &carrier, &[])
                        } else {
                            project_reading_metadata(&local, &[], &[])
                        };
                        let metadata =
                            &projected.0["__anchorAcquisition"]["nativePictureMetadata"][field];
                        assert_eq!(metadata["opid"], opid);
                        assert_eq!(metadata["value"], bytes.len() as u32);
                        assert_eq!(metadata["rawBytes"], serde_json::json!(bytes));
                        assert_eq!(metadata["retention"], "inactiveOpaqueNotDecoded");
                        assert_eq!(
                            metadata["scope"],
                            if document { "documentDefault" } else { "shape" }
                        );
                        assert_eq!(without_reading_metadata(projected.0), control.0);
                        assert_eq!(projected.1, control.1);
                        assert_eq!(projected.2, control.2);
                    }
                }
            }
        }
        // Equivalent duplicates do not conflict merely because the ignored bit
        // differs. Metadata preserves the last agreeing encoded representative.
        for (id, field) in [
            (0x0186, "inactiveFillCarrier"),
            (0x01c5, "inactiveLineCarrier"),
        ] {
            for bytes in [&[][..], &[0x41, 0x42][..]] {
                for flags in [[0, 0x4000], [0x4000, 0]] {
                    let repeated = [
                        (id | 0x8000 | flags[0], bytes.len() as u32, bytes),
                        (id | 0x8000 | flags[1], bytes.len() as u32, bytes),
                    ];
                    let local = [owners.as_slice(), repeated.as_slice()].concat();
                    let projected = project_reading_metadata(&local, &[], &[]);
                    let metadata =
                        &projected.0["__anchorAcquisition"]["nativePictureMetadata"][field];
                    assert_eq!(metadata["opid"], id | 0x8000 | flags[1]);
                    assert_eq!(metadata["rawBytes"], serde_json::json!(bytes));
                    assert_eq!(metadata["retention"], "inactiveOpaqueNotDecoded");
                    assert_eq!(without_reading_metadata(projected.0), control.0);
                    assert_eq!(projected.1, control.1);
                    assert_eq!(projected.2, control.2);
                    let conflict = [(id | 0x8000, 1, &[0x58][..]), (id | 0xc000, 1, &[0x59][..])];
                    assert_reading_metadata_refused(
                        &[owners.as_slice(), conflict.as_slice()].concat(),
                        &[],
                        &[],
                    );
                }
            }
        }
    }

    #[test]
    fn reading_picture_complex_bid_keeps_inactive_owner_and_scalar_index_guards() {
        let control = project_reading_metadata(&[], &[], &[]);
        for (id, owner, disabled, enabled, field) in [
            (
                0x0186,
                0x01bf,
                0x0010_0000,
                0x0010_0010,
                "inactiveFillCarrier",
            ),
            (
                0x01c5,
                0x01ff,
                0x0008_0000,
                0x0008_0008,
                "inactiveLineCarrier",
            ),
        ] {
            for bid in [0, 0x4000] {
                let carrier = (id | 0x8000 | bid, 2, &[0x41, 0x42][..]);
                assert_reading_metadata_refused(&[carrier], &[], &[]);
                assert_reading_metadata_refused(&[carrier, (owner, enabled, &[])], &[], &[]);
                assert_reading_metadata_refused(
                    &[carrier, (owner, 0, &[])],
                    &[(owner, enabled, &[])],
                    &[],
                );
            }
            // A scalar BLIP index still requires fBid; zero does not waive that rule.
            for index in [0, 17] {
                assert_reading_metadata_refused(
                    &[(id, index, &[]), (owner, disabled, &[])],
                    &[],
                    &[],
                );
                let projected = project_reading_metadata(
                    &[(id | 0x4000, index, &[]), (owner, disabled, &[])],
                    &[],
                    &[],
                );
                let metadata = &projected.0["__anchorAcquisition"]["nativePictureMetadata"][field];
                assert_eq!(metadata["opid"], id | 0x4000);
                assert_eq!(
                    metadata["retention"],
                    if index == 0 {
                        "ignoredZeroIndex"
                    } else {
                        "inactiveIndexNotResolved"
                    }
                );
                assert_eq!(without_reading_metadata(projected.0), control.0);
                assert_eq!(projected.2, control.2);
            }
        }
    }

    #[test]
    fn reading_picture_complex_bid_keeps_complete_tail_and_owned_output_budget() {
        for (id, owner, disabled) in [(0x0186, 0x01bf, 0x0010_0000), (0x01c5, 0x01ff, 0x0008_0000)]
        {
            for bid in [0, 0x4000] {
                let opid = id | 0x8000 | bid;
                // Declared complex length is authoritative even though fBid is ignored.
                assert_reading_metadata_refused(
                    &[(opid, 3, &[0x41, 0x42]), (owner, disabled, &[])],
                    &[],
                    &[],
                );
                for bytes in [&[][..], &[0x41, 0x42, 0x43][..]] {
                    let props = [(opid, bytes.len() as u32, bytes), (owner, disabled, &[])];
                    let (word, table) = reading_picture_with_metadata(&props, &[], &[]);
                    let mut store = Store::read(&word, &table, 20).unwrap();
                    store.set_reading_relocation(true);
                    assert_eq!(
                        store.direct_picture(12, &mut 0).unwrap_err(),
                        "OUTPUT_TOO_LARGE"
                    );
                }
            }
        }
    }

    #[test]
    fn delayed_metafiles_keep_owned_bytes_cached_across_floating_occurrences() {
        for (source, blip, mime) in [
            {
                let (s, b) = crate::officeart::emf_test_blip();
                (s, b, "image/emf")
            },
            {
                let (s, b) = crate::officeart::wmf_test_blip();
                (s, b, "image/wmf")
            },
        ] {
            let (word, table) = drawing_with_blip(0xa00, 0, &[], blip, 2);
            let mut store = Store::read(&word, &table, 20).unwrap();
            store.remaining_bytes = source.len();
            let first = picture(&mut store).unwrap().unwrap();
            assert_eq!(store.remaining_bytes, 0);
            // Shape records are still parsed per occurrence; image bytes are not.
            let second = picture(&mut store).unwrap().unwrap();
            assert_eq!(first.image.image_path, second.image.image_path);
            assert_eq!(first.image.mime_type, mime);
            let resources = resources(store);
            assert_eq!(resources.len(), 1);
            assert_eq!(resources[0].bytes, source);
        }
    }
    #[test]
    fn post_eof_wmf_retains_the_floating_picture_and_bounded_source() {
        let (mut source, _) = crate::officeart::wmf_test_blip();
        source.extend_from_slice(&[0, 0]);
        let words = (source.len() / 2) as u32;
        source[6..10].copy_from_slice(&words.to_le_bytes());
        let mut payload = vec![0; 50];
        payload[16..20].copy_from_slice(&(source.len() as u32).to_le_bytes());
        payload[44..48].copy_from_slice(&(source.len() as u32).to_le_bytes());
        payload[48] = 0xfe;
        payload[49] = 0xfe;
        payload.extend_from_slice(&source);
        let blip = record(0xf01b, 0x2160, &payload);
        for flags in [0xa00, 0xa10] {
            let (word, table) = drawing_with_blip(flags, 0, &[], blip.clone(), 2);
            let mut store = Store::read(&word, &table, 20).unwrap();
            // Shared WMF admission validates records through META_EOF and
            // retains the bounded original BLIP payload, including its tail.
            let projected = picture(&mut store).unwrap().unwrap();
            assert_eq!(projected.image.mime_type, "image/wmf");
            assert!(!store.omitted);
            let retained = resources(store);
            assert_eq!(retained.len(), 1);
            assert_eq!(retained[0].bytes, source);
        }
    }
    #[test]
    fn ole_shape_without_pib_is_omitted_without_residue() {
        let (word, mut table) = drawing_input(0xa10, 0);
        let key = 0x4104u16.to_le_bytes();
        let offset = table
            .windows(key.len())
            .position(|bytes| bytes == key)
            .expect("fixture contains pib property");
        table[offset..offset + 2].copy_from_slice(&0x4105u16.to_le_bytes());
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(picture(&mut store).unwrap().is_none());
        assert!(store.omitted);
        assert!(resources(store).is_empty());
    }
    #[test]
    fn delayed_pictures_use_word_stream_and_share_resources_across_occurrences() {
        let (word, table) = drawing_input(0xa00, 0);
        let mut store = Store::read(&word, &table, 20).unwrap();
        let first = picture(&mut store).unwrap().unwrap();
        let second = picture(&mut store).unwrap().unwrap();
        assert!(first.image.anchor);
        assert_eq!(first.image.anchor_x_pt, -5.0);
        assert_eq!((first.image.width_pt, first.image.height_pt), (20.0, 15.0));
        assert_ne!(first.occurrence_id, second.occurrence_id);
        assert_eq!(first.image.image_path, second.image.image_path);
        let resources = resources(store);
        assert_eq!(resources.len(), 1);
        assert!(resources[0].bytes.starts_with(b"\x89PNG"));
        let mut truncated = Store::read(&word[..1024], &table, 20).unwrap();
        assert!(picture(&mut truncated).is_err());
    }
    #[test]
    fn direct_model_projects_nothing_for_hidden_drawings() {
        // fHidden (use bit 17, value bit 1) prevents display; the BLIP is
        // never dereferenced. Script anchors remain an omission.
        let (word, table) = drawing_input(0xa00, 0x0002_0002);
        let mut store = Store::read(&word[..1024], &table, 20).unwrap();
        assert!(store
            .direct_picture(12, &mut usize::MAX.clone())
            .unwrap()
            .is_none());
        assert!(!store.omitted && store.images.is_empty());
        let (word, table) = drawing_input(0xa00, 0x0082_0082);
        let mut store = Store::read(&word[..1024], &table, 20).unwrap();
        assert!(store
            .direct_picture(12, &mut usize::MAX.clone())
            .unwrap()
            .is_none());
        assert!(store.omitted);
    }

    #[test]
    fn direct_alignment_uses_the_spa_origin_and_rejects_disagreement() {
        // The fixture's SPA origin is page/page.
        let aligned = |options: &[(u16, u32)]| {
            let (word, table) = drawing_with_options(0xa00, 0, options);
            let mut store = Store::read(&word, &table, 20).unwrap();
            let picture = store.direct_picture(12, &mut usize::MAX.clone());
            (picture, store.omitted)
        };
        for (posh, align) in [(1, "left"), (2, "center"), (3, "right")] {
            let (picture, omitted) = aligned(&[(0x38f, posh), (0x390, 1)]);
            let image = picture.unwrap().unwrap().image;
            assert!(!omitted);
            assert_eq!(image.anchor_x_align.as_deref(), Some(align));
            assert_eq!(image.anchor_x_relative_from.as_deref(), Some("page"));
            let facts = image.anchor_acquisition.unwrap();
            assert!(matches!(
                facts.horizontal.choice,
                docx_model::AnchorAxisChoiceWire::Align { ref value } if value == align
            ));
        }
        for (posv, align) in [(1, "top"), (2, "center"), (3, "bottom")] {
            let (picture, _) = aligned(&[(0x391, posv), (0x392, 1)]);
            let image = picture.unwrap().unwrap().image;
            assert_eq!(image.anchor_y_align.as_deref(), Some(align));
            assert_eq!(image.anchor_x_align, None);
        }
        // Disagreeing, defaulted (text) or character-relative containers and
        // page-parity alignment keep the drawing out of the document.
        for options in [
            &[(0x38f, 2), (0x390, 0)][..],
            &[(0x38f, 2)][..],
            &[(0x38f, 2), (0x390, 3)][..],
            &[(0x38f, 4), (0x390, 1)][..],
            &[(0x391, 1), (0x392, 0)][..],
            &[(0x391, 5), (0x392, 1)][..],
        ] {
            let (picture, omitted) = aligned(options);
            assert!(picture.unwrap().is_none() && omitted, "{options:x?}");
        }
    }

    #[test]
    fn direct_floating_passes_validated_metafiles_with_docx_media_types() {
        for ((source, blip), mime) in [
            (crate::officeart::emf_test_blip(), "image/emf"),
            (crate::officeart::wmf_test_blip(), "image/wmf"),
        ] {
            let (word, table) = drawing_with_blip(0xa00, 0, &[], blip, 2);
            let mut store = Store::read(&word, &table, 20).unwrap();
            let mut budget = usize::MAX;
            let picture = store.direct_picture(12, &mut budget).unwrap().unwrap();
            assert_eq!(picture.image.mime_type, mime);
            let mut resources = Vec::new();
            store
                .append_referenced_direct_resources(
                    &mut resources,
                    &[picture.image.image_path.as_str()],
                    &mut budget,
                )
                .unwrap();
            assert_eq!(resources[0].mime_type, mime);
            assert_eq!(resources[0].bytes, source);
        }
    }

    #[test]
    fn direct_floating_deduplicates_resources_and_admits_before_reserving() {
        let (word, table) = drawing_with_options(
            0xac0,
            0x8200_0000,
            &[
                (0x384, 12_700),
                (0x385, 25_400),
                (0x386, 38_100),
                (0x387, 50_800),
            ],
        );
        let mut store = Store::read(&word, &table, 20).unwrap();
        let mut budget = usize::MAX;
        let first = store.direct_picture(12, &mut budget).unwrap().unwrap();
        let second = store.direct_picture(12, &mut budget).unwrap().unwrap();
        assert_eq!(first.image.image_path, second.image.image_path);
        assert_ne!(first.occurrence_id, second.occurrence_id);
        assert_eq!(first.image.wrap_mode.as_deref(), Some("square"));
        assert_eq!(
            first
                .image
                .anchor_acquisition
                .as_ref()
                .unwrap()
                .wrap
                .authored_kinds,
            ["wrapSquare"]
        );
        let acquisition = first.image.anchor_acquisition.as_ref().unwrap();
        let expected_payload = std::mem::size_of::<docx_model::ImageRun>()
            + first.image.image_path.capacity()
            + first.image.mime_type.capacity()
            + first.image.wrap_mode.as_ref().unwrap().capacity()
            + first.image.wrap_side.as_ref().unwrap().capacity()
            + first
                .image
                .anchor_x_relative_from
                .as_ref()
                .unwrap()
                .capacity()
            + first
                .image
                .anchor_y_relative_from
                .as_ref()
                .unwrap()
                .capacity()
            + first.occurrence_id.capacity()
            + acquisition.occurrence_id.capacity()
            + acquisition
                .horizontal
                .relative_from
                .as_ref()
                .unwrap()
                .capacity()
            + acquisition
                .vertical
                .relative_from
                .as_ref()
                .unwrap()
                .capacity()
            + acquisition.wrap.side.as_ref().unwrap().capacity()
            + acquisition.wrap.authored_kinds.capacity() * std::mem::size_of::<String>()
            + acquisition.wrap.authored_kinds[0].capacity();
        let mut exact_store = Store::read(&word, &table, 20).unwrap();
        let mut short = expected_payload - 1;
        assert_eq!(
            exact_store.direct_picture(12, &mut short).unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );

        let mut resources = Vec::new();
        let mut none = 0;
        assert_eq!(
            store
                .append_direct_resources(&mut resources, &mut none)
                .unwrap_err(),
            "OUTPUT_TOO_LARGE"
        );
        assert_eq!(resources.capacity(), 0);

        let mut store = Store::read(&word, &table, 20).unwrap();
        let mut budget = usize::MAX;
        store.direct_picture(12, &mut budget).unwrap().unwrap();
        store.direct_picture(12, &mut budget).unwrap().unwrap();
        let mut resources = Vec::new();
        store
            .append_direct_resources(&mut resources, &mut budget)
            .unwrap();
        assert_eq!(resources.len(), 1);
        assert!(resources[0].bytes.starts_with(b"\x89PNG"));

        for horizontal in 0u16..=2 {
            for vertical in 0u16..=2 {
                for wrapping in 1u16..=3 {
                    let (word, mut table) = drawing_input(0xa00, 0);
                    let flags = (horizontal << 1) | (vertical << 3) | (wrapping << 5);
                    table[28..30].copy_from_slice(&flags.to_le_bytes());
                    let mut store = Store::read(&word, &table, 20).unwrap();
                    let mut budget = usize::MAX;
                    let image = store
                        .direct_picture(12, &mut budget)
                        .unwrap()
                        .unwrap()
                        .image;
                    assert_eq!(
                        image.anchor_x_relative_from.as_deref(),
                        Some(["margin", "page", "column"][horizontal as usize])
                    );
                    assert_eq!(
                        image.anchor_y_relative_from.as_deref(),
                        Some(["margin", "page", "paragraph"][vertical as usize])
                    );
                    assert_eq!(image.anchor_x_from_margin, horizontal != 1);
                    assert_eq!(image.anchor_y_from_para, vertical == 2);
                    assert_eq!(
                        image.wrap_mode.as_deref(),
                        Some(["", "topAndBottom", "square", "none"][wrapping as usize])
                    );
                }
            }
        }
    }

    #[test]
    fn shared_resolution_rejects_position_overflow_before_occurrence() {
        let (word, table) = drawing_input(0xa00, 0);
        let anchor_offset = u32_at(&word, 0x1da).unwrap() as usize;
        let record = anchor_offset + 8;
        let mut table = table;
        table[record + 4..record + 8].copy_from_slice(&4_000_000i32.to_le_bytes());
        table[record + 12..record + 16].copy_from_slice(&4_000_400i32.to_le_bytes());
        let mut store = Store::read(&word, &table, 20).unwrap();
        let mut budget = usize::MAX;
        assert!(store
            .direct_picture(12, &mut budget)
            .unwrap_err()
            .contains("position"));
        assert_eq!(store.occurrences, 0);
    }
    #[test]
    fn ole_shapes_admit_only_their_validated_passive_blip() {
        let (png_word, png_table) = drawing_input(0xa10, 0);
        let (wmf_source, wmf_blip) = crate::officeart::wmf_test_blip();
        let (wmf_word, wmf_table) = drawing_with_blip(0xa10, 0, &[], wmf_blip, 2);
        for (word, table, expected) in [
            (png_word, png_table, None),
            (wmf_word, wmf_table, Some(wmf_source)),
        ] {
            let mut store = Store::read(&word, &table, 20).unwrap();
            let image = picture(&mut store).unwrap().unwrap().image;
            assert!(image.anchor);
            assert_eq!(image.anchor_x_relative_from.as_deref(), Some("page"));
            assert_eq!(image.anchor_y_relative_from.as_deref(), Some("page"));
            assert_eq!((image.width_pt, image.height_pt), (20.0, 15.0));
            assert_eq!(image.wrap_mode.as_deref(), Some("square"));
            assert_eq!(image.wrap_side.as_deref(), Some("bothSides"));
            assert!(!store.omitted);
            let resources = resources(store);
            assert_eq!(resources.len(), 1);
            if let Some(expected) = expected {
                assert_eq!(resources[0].bytes, expected);
            } else {
                assert!(resources[0].bytes.starts_with(b"\x89PNG"));
            }
        }
    }
    #[test]
    fn structural_hidden_and_script_flags_still_prevent_blip_access() {
        for flag in [0x01, 0x02, 0x04, 0x08, 0x100] {
            for ole in [0, 0x10] {
                let (word, table) = drawing_input(0xa00 | flag | ole, 0);
                let mut store = Store::read(&word[..1024], &table, 20).unwrap();
                assert!(picture(&mut store).unwrap().is_none());
                assert!(store.images.is_empty());
                assert!(store.omitted);
            }
        }
        // A hidden drawing projects nothing without an omission; a script
        // anchor stays an omission. Neither dereferences its BLIP.
        for (group, omitted) in [(0x00020002, false), (0x00800080, true)] {
            let (word, table) = drawing_input(0xa10, group);
            let mut store = Store::read(&word[..1024], &table, 20).unwrap();
            assert!(picture(&mut store).unwrap().is_none());
            assert!(store.images.is_empty());
            assert_eq!(store.omitted, omitted);
        }
    }
    #[test]
    fn ambiguous_alignment_and_rotated_shapes_do_not_dereference_blips() {
        for key in [0x38f, 0x391, 4] {
            for value in 1..=5 {
                let operand = if key == 4 { value << 16 } else { value };
                let (word, table) = drawing_with_options(0xa00, 0, &[(key, operand)]);
                let mut store = Store::read(&word[..1024], &table, 20).unwrap();
                assert!(picture(&mut store).unwrap().is_none());
                assert!(store.images.is_empty());
                assert!(store.omitted);
            }
        }
    }
    #[test]
    fn enforces_media_occurrence_and_operation_budgets() {
        let (word, table) = drawing_input(0xa00, 0);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.remaining_bytes = 0;
        assert!(picture(&mut store).unwrap_err().contains("media budget"));
        assert!(resources(store).is_empty());

        let mut store = Store::read(&word, &table, 20).unwrap();
        store.occurrences = 100_000;
        assert!(picture(&mut store)
            .unwrap_err()
            .contains("occurrence budget"));

        let mut store = Store::read(&word, &table, 20).unwrap();
        store.budget = 0;
        assert!(picture(&mut store).is_err());
    }
    #[test]
    fn does_not_reassign_header_or_nested_group_drawings_to_main_story() {
        let (word, mut table) = drawing_input(0xa00, 0);
        let start = u32_at(&word, 0x22a).unwrap() as usize;
        let (_, group_end) = record_with_end(&table, start, &mut 100, "test").unwrap();
        table[group_end] = 1;
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(picture(&mut store).unwrap().is_none());
        assert!(store.images.is_empty());

        table[group_end] = 0;
        // Replace the independent SpContainer tag with a nested SpgrContainer
        // lacking its group shape; it is never flattened into the story.
        let shape_start = group_end + 1 + 8 + 8;
        table[shape_start + 2..shape_start + 4].copy_from_slice(&0xf003u16.to_le_bytes());
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(picture(&mut store).unwrap().is_none());
        assert!(store.images.is_empty());
    }
    #[test]
    fn placement_masks_honor_explicit_false_and_ignore_unused_bits() {
        let mut placement = Placement::default();
        let apply = |p: &mut Placement, value: u32| {
            let body = [0x3bfu16.to_le_bytes().as_slice(), &value.to_le_bytes()].concat();
            p.apply(
                Record {
                    version: 3,
                    instance: 1,
                    kind: 0xf00b,
                    payload: &body,
                },
                &mut 10,
            )
            .unwrap();
        };
        apply(&mut placement, 0x00020002);
        assert!(placement.hidden);
        apply(&mut placement, 0);
        assert!(placement.hidden);
        apply(&mut placement, 0x00020000);
        assert!(!placement.hidden);
        apply(&mut placement, 0x82008000);
        assert!(placement.in_cell);
        assert!(!placement.overlap);
    }
    fn input(flags: u16) -> (Vec<u8>, Vec<u8>) {
        let mut word = vec![0u8; 0x232];
        word[0x1de..0x1e2].copy_from_slice(&34u32.to_le_bytes());
        let mut table = [
            12u32.to_le_bytes(),
            30u32.to_le_bytes(),
            1027u32.to_le_bytes(),
            (-100i32).to_le_bytes(),
            200i32.to_le_bytes(),
            300i32.to_le_bytes(),
            500i32.to_le_bytes(),
        ]
        .concat();
        table.extend(flags.to_le_bytes());
        table.extend([0; 4]);
        (word, table)
    }
    #[test]
    fn preserves_signed_rectangle_origin_wrap_and_layer() {
        let (word, table) = input((1 << 1) | (2 << 3) | (3 << 5) | (1 << 14) | (1 << 15));
        let a = anchors(&word, &table, 20).unwrap();
        assert_eq!(a.len(), 1);
        assert_eq!(a[0].cp, 12);
        assert_eq!(a[0].shape_id, 1027);
        assert_eq!(a[0].rect, [-100, 200, 300, 500]);
        assert_eq!(a[0].horizontal, "page");
        assert_eq!(a[0].vertical, "paragraph");
        assert_eq!(a[0].wrapping, 3);
        assert!(a[0].behind);
        assert!(a[0].locked);
        // The final sentinel CP is undefined, apart from monotonicity. It may
        // exceed ccpText and must not be treated as a live anchor.
    }
    #[test]
    fn ignored_wrapping_fields_do_not_reject_top_bottom_anchors() {
        let (word, table) = input((1 << 5) | (15 << 9) | (1 << 14));
        let a = anchors(&word, &table, 20).unwrap();
        assert_eq!(a[0].side, "bothSides");
        assert!(!a[0].behind);
    }
    #[test]
    fn rejects_invalid_plc_size_origin_and_live_cp_but_allows_absent_table() {
        let (word, table) = input(3 << 1);
        assert!(anchors(&word, &table, 20).is_err());
        let (word, table) = input(0);
        assert!(anchors(&word, &table, 11).is_err());
        assert!(anchors(&word, &table[..table.len() - 1], 20).is_err());
        let mut word = word;
        word[0x1de..0x1e2].fill(0);
        assert!(anchors(&word, &[], 20).unwrap().is_empty());
    }
}
