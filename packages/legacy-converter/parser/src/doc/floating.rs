//! Floating drawing anchors of the main and header documents, MS-DOC 2.8.27
//! and 2.9.253.
use super::pictures::{Options as PictureOptions, Picture};
use super::{u16_at, u32_at, unsupported};
use crate::officeart::{
    raster::{read_store_entry, Image},
    record_with_end, Record,
};
use std::collections::BTreeMap;

#[cfg(feature = "direct-doc")]
mod direct;
#[cfg(feature = "direct-doc")]
mod group;
#[cfg(feature = "direct-doc")]
mod shape;
#[cfg(feature = "direct-doc")]
pub(in crate::doc) mod textbox;
#[cfg(feature = "direct-doc")]
pub(in crate::doc) use direct::DirectRun;

/// The drawing part that owns a PlcfSpa, its OfficeArtDgContainer and its
/// textbox story (MS-DOC 2.8.27, 2.9.171; MS-ODRAW 2.2.13).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
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

/// Which projection consumes the resolved facts. The package writer keeps
/// its original picture-only subset; the direct model admits the drawing
/// shapes and aligned positions implemented in `shape` and `direct`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Mode {
    Package,
    #[cfg(feature = "direct-doc")]
    Direct,
}

/// The anchors and top-level shape containers of one drawing part.
#[derive(Default)]
struct Drawings<'a> {
    anchors: Vec<Anchor>,
    /// spid -> (anchor index, [package order, document order], container).
    shapes: BTreeMap<u32, (usize, [u32; 2], Record<'a>)>,
}

pub(super) struct Store<'a> {
    /// Indexed by `Part as usize`: main document, header document.
    parts: [Drawings<'a>; 2],
    entries: Vec<Record<'a>>,
    word: &'a [u8],
    table: &'a [u8],
    /// The CLX that maps textbox stories; empty when none was supplied.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    clx: &'a [u8],
    group_read: bool,
    header_container: Option<Record<'a>>,
    #[cfg(feature = "direct-doc")]
    textboxes: [Option<textbox::Textboxes<'a>>; 2],
    images: BTreeMap<usize, Option<Image<'a>>>,
    budget: usize,
    remaining_bytes: usize,
    occurrences: u32,
    #[cfg(feature = "direct-doc")]
    selected_images: std::collections::BTreeSet<usize>,
    pub omitted: bool,
}

impl<'a> Store<'a> {
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
                    shapes: BTreeMap::new(),
                },
                Drawings::default(),
            ],
            entries: Vec::new(),
            word,
            table,
            clx,
            group_read: false,
            header_container: None,
            #[cfg(feature = "direct-doc")]
            textboxes: [None, None],
            images: BTreeMap::new(),
            budget: 1_000_000,
            remaining_bytes: 128 * 1024 * 1024,
            occurrences: 0,
            #[cfg(feature = "direct-doc")]
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
    #[cfg(feature = "direct-doc")]
    fn load_header(&mut self, header_units: usize) -> Result<(), String> {
        let anchors = anchors_in(self.word, self.table, Part::Header, header_units)?;
        if anchors.is_empty() {
            return Ok(());
        }
        self.parts[1].anchors = anchors;
        self.read_group()?;
        if let Some(container) = self.header_container {
            self.parts[1].shapes = container_shapes(container, &mut self.budget)?;
        } else if !self.omitted {
            return Err(unsupported("Word header anchors lack their drawing"));
        }
        Ok(())
    }

    /// Load everything the direct model resolves beyond main-story anchors:
    /// header drawings and both textbox stories (MS-DOC 2.3.6-2.3.7).
    #[cfg(feature = "direct-doc")]
    pub(in crate::doc) fn load_direct_parts(&mut self) -> Result<(), String> {
        let header_units = u32_at(self.word, 0x54)? as usize; // FibRgLw97.ccpHdd
        self.load_header(header_units)?;
        for part in [Part::Main, Part::Header] {
            self.textboxes[part as usize] =
                textbox::Textboxes::read(self.word, self.table, self.clx, part)?;
        }
        Ok(())
    }

    #[cfg(feature = "direct-doc")]
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
                self.parts[0].shapes = container_shapes(drawing, &mut self.budget)?;
            } else {
                // Never leak header drawings into the body; they are
                // resolved only against header anchors.
                self.header_container = Some(drawing);
            }
        }
        Ok(())
    }

    pub fn drawing(&mut self, cp: usize) -> Result<String, String> {
        let Some(resolved) = self.resolve(Part::Main, cp, Mode::Package)? else {
            return Ok(String::new());
        };
        let (image_index, crop) = match resolved.content {
            Content::Picture {
                image_index, crop, ..
            } => (image_index, crop),
            #[cfg(feature = "direct-doc")]
            Content::Shape(_) | Content::Group(_) => {
                unreachable!("package projection resolves pictures only")
            }
        };
        let image = self.images[&image_index]
            .as_ref()
            .expect("resolved floating image");
        let [dist_l, dist_t, dist_r, dist_b] = resolved.distances;
        let x = position(resolved.horizontal, resolved.x_emu, false)?;
        let y = position(resolved.vertical, resolved.y_emu, true)?;
        let wrap = match resolved.wrapping {
            1 => "<wp:wrapTopAndBottom/>".to_string(),
            2 => format!("<wp:wrapSquare wrapText=\"{}\"/>", resolved.side),
            3 => "<wp:wrapNone/>".into(),
            _ => unreachable!(),
        };
        let opening = format!(
            r#"<wp:anchor distL="{dist_l}" distT="{dist_t}" distR="{dist_r}" distB="{dist_b}" simplePos="0" relativeHeight="{}" behindDoc="{}" locked="{}" layoutInCell="{}" allowOverlap="{}"><wp:simplePos x="0" y="0"/>{x}{y}<wp:extent cx="{}" cy="{}"/>{wrap}"#,
            resolved.z_order,
            u8::from(resolved.behind),
            u8::from(resolved.locked),
            u8::from(resolved.in_cell),
            u8::from(resolved.overlap),
            resolved.extent[0],
            resolved.extent[1]
        );
        let image = Picture {
            image: Image {
                bytes: std::borrow::Cow::Borrowed(image.bytes.as_ref()),
                extension: image.extension,
            },
            extent: resolved.extent,
            crop,
            flip: resolved.flip,
            rotation: 0,
        };
        Ok(image.xml(
            1_000_000 + resolved.occurrence,
            &format!("rFloatImg{image_index}"),
            &opening,
            "</wp:anchor>",
        ))
    }

    fn resolve(
        &mut self,
        part: Part,
        cp: usize,
        mode: Mode,
    ) -> Result<Option<ResolvedDrawing>, String> {
        let drawings = &self.parts[part as usize];
        let Ok(index) = drawings.anchors.binary_search_by_key(&cp, |a| a.cp) else {
            self.omitted = true;
            return Ok(None);
        };
        let anchor = drawings.anchors[index].clone();
        let anchor = &anchor;
        let Some(&(anchor_index, orders, shape)) = drawings.shapes.get(&anchor.shape_id) else {
            self.omitted = true;
            return Ok(None);
        };
        if anchor_index != index {
            return Err(unsupported("Word shape/anchor index mismatch"));
        }
        // The package writer numbers only independent shapes, as it always
        // has; the direct model counts groups in their document order too.
        let order = match mode {
            Mode::Package => orders[0],
            #[cfg(feature = "direct-doc")]
            Mode::Direct => orders[1],
        };
        if shape.kind == 0xf003 {
            // A top-level OfficeArt group. The package writer never flattens
            // groups; the direct model resolves them in `group`.
            #[cfg(feature = "direct-doc")]
            if mode == Mode::Direct {
                return self.resolve_group(anchor, order, shape);
            }
            self.omitted = true;
            return Ok(None);
        }
        let mut picture = PictureOptions::default();
        let mut placement = Placement::default();
        let mut flags = None;
        let mut kind = None;
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
                _ => {}
            }
        }
        let flags = flags.unwrap_or(0);
        #[cfg(feature = "direct-doc")]
        if mode == Mode::Direct && placement.hidden && !placement.script {
            // MS-ODRAW 2.3.4.44 fHidden: the shape is prevented from
            // displaying, so the direct model projects nothing for it. Word's
            // PDF of a corpus document with hidden header lines agrees. The
            // package writer keeps its original omission.
            return Ok(None);
        }
        let [left, top, right, bottom] = anchor.rect.map(i64::from);
        let extent = [(right - left) * 635, (bottom - top) * 635];
        #[cfg(feature = "direct-doc")]
        if mode == Mode::Direct && kind != Some(75) {
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
        let align = match mode {
            Mode::Package => None,
            #[cfg(feature = "direct-doc")]
            Mode::Direct => {
                if placement.horizontal == 0 && placement.vertical == 0 {
                    None
                } else {
                    match direct_alignment(anchor, &placement) {
                        Ok(align) => Some(align),
                        Err(_) => {
                            self.omitted = true;
                            return Ok(None);
                        }
                    }
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
            // SPA provides an explicit, host-defined coordinate origin. The
            // package writer does not resolve aligned positions; the direct
            // model accepts them only through `direct_alignment`.
            || (align.is_none() && (placement.horizontal != 0 || placement.vertical != 0))
            || matches!(anchor.wrapping, 0 | 4 | 5)
        {
            self.omitted = true;
            return Ok(None);
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
        .map(Some)
    }

    /// Load a shape's picture fill BLIP; `false` when it is not a supported
    /// passive image, which the caller reports as omitted content.
    #[cfg(feature = "direct-doc")]
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
            let image = read_store_entry(
                entry,
                Some(self.word),
                &mut self.budget,
                self.remaining_bytes,
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
            inline: false,
            #[cfg(feature = "direct-doc")]
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
    pub fn relationships(&self) -> String {
        self.images.iter().filter_map(|(id,image)|image.as_ref().map(|p|format!(r#"<Relationship Id="rFloatImg{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/float{id}.{}"/>"#,p.extension))).collect()
    }
    pub fn parts(&self) -> Vec<(String, &[u8])> {
        self.images
            .iter()
            .filter_map(|(id, image)| {
                image.as_ref().map(|p| {
                    (
                        format!("word/media/float{id}.{}", p.extension),
                        p.bytes.as_ref(),
                    )
                })
            })
            .collect()
    }
}

enum Content {
    Picture {
        image_index: usize,
        #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
        extension: &'static str,
        crop: [i64; 4],
    },
    #[cfg(feature = "direct-doc")]
    Shape(Box<shape::Facts>),
    /// Members of an OfficeArt group in source (paint) order.
    #[cfg(feature = "direct-doc")]
    Group(Vec<group::Member>),
}

struct ResolvedDrawing {
    content: Content,
    /// Projected in paragraph flow instead of anchored (pseudo-inline).
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
    inline: bool,
    #[cfg(feature = "direct-doc")]
    shape_id: u32,
    extent: [i64; 2],
    flip: [bool; 2],
    x_emu: i64,
    y_emu: i64,
    horizontal: &'static str,
    vertical: &'static str,
    /// Horizontal/vertical `wp:align` values replacing the offsets when set.
    #[cfg_attr(not(feature = "direct-doc"), allow(dead_code))]
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

/// The anchored shapes and groups of one OfficeArtDgContainer's patriarch.
/// Groups are kept whole: their members use the group coordinate space.
fn container_shapes<'a>(
    drawing: Record<'a>,
    budget: &mut usize,
) -> Result<BTreeMap<u32, (usize, [u32; 2], Record<'a>)>, String> {
    let mut shapes = BTreeMap::new();
    let mut independent = 0u32;
    for child in records(drawing.payload, budget)? {
        if child.kind != 0xf003 {
            continue;
        }
        if child.version != 15 {
            return Err(unsupported("invalid Word shape group"));
        }
        for shape in records(child.payload, budget)? {
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
                if entry.kind == 0xf004 {
                    independent += 1;
                }
                let orders = [independent, shapes.len() as u32 + 1];
                if shapes.insert(id, (anchor_index, orders, shape)).is_some() {
                    return Err(unsupported("duplicate Word floating shape identifier"));
                }
            }
        }
    }
    Ok(shapes)
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
        for p in entries.chunks_exact(6) {
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
#[cfg(feature = "direct-doc")]
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

fn position(origin: &str, offset: i64, vertical: bool) -> Result<String, String> {
    let axis = if vertical { "V" } else { "H" };
    i32::try_from(offset)
        .map_err(|_| unsupported("Word floating position exceeds DrawingML range"))?;
    let value = format!("<wp:posOffset>{offset}</wp:posOffset>");
    Ok(format!(
        "<wp:position{axis} relativeFrom=\"{origin}\">{value}</wp:position{axis}>"
    ))
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Anchor {
    pub cp: usize,
    pub shape_id: u32,
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
    #[test]
    fn delayed_metafiles_keep_owned_bytes_cached_across_floating_occurrences() {
        for (source, blip, extension) in [
            {
                let (s, b) = crate::officeart::emf_test_blip();
                (s, b, ".emf")
            },
            {
                let (s, b) = crate::officeart::wmf_test_blip();
                (s, b, ".wmf")
            },
        ] {
            let (word, table) = drawing_with_blip(0xa00, 0, &[], blip, 2);
            let mut store = Store::read(&word, &table, 20).unwrap();
            store.remaining_bytes = source.len();
            assert!(store.drawing(12).unwrap().contains("<wp:anchor"));
            let pointer = store.parts()[0].1.as_ptr();
            assert_eq!(store.parts()[0].1, source);
            assert_eq!(store.remaining_bytes, 0);
            // Shape records are still parsed per occurrence; image bytes are not.
            assert!(store.drawing(12).unwrap().contains("<wp:anchor"));
            assert_eq!(store.parts()[0].1.as_ptr(), pointer);
            assert!(store.relationships().contains(extension));
        }
    }
    #[test]
    fn post_eof_wmf_omits_floating_xml_relationship_and_media() {
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
            assert!(store.drawing(12).unwrap().is_empty());
            assert!(store.omitted);
            assert!(store.relationships().is_empty());
            assert!(store.parts().is_empty());
        }
    }
    #[test]
    fn ole_shape_without_pib_is_omitted_without_package_residue() {
        let (word, mut table) = drawing_input(0xa10, 0);
        let key = 0x4104u16.to_le_bytes();
        let offset = table
            .windows(key.len())
            .position(|bytes| bytes == key)
            .expect("fixture contains pib property");
        table[offset..offset + 2].copy_from_slice(&0x4105u16.to_le_bytes());
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.drawing(12).unwrap().is_empty());
        assert!(store.relationships().is_empty());
        assert!(store.parts().is_empty());
        assert!(store.omitted);
    }
    #[test]
    fn delayed_pictures_use_word_stream_and_share_parts_without_sharing_drawing_ids() {
        let (word, table) = drawing_input(0xa00, 0);
        let mut store = Store::read(&word, &table, 20).unwrap();
        let first = store.drawing(12).unwrap();
        let second = store.drawing(12).unwrap();
        assert!(first.contains("<wp:anchor"));
        assert!(first.contains("<wp:posOffset>-63500</wp:posOffset>"));
        assert!(first.contains("cx=\"254000\" cy=\"190500\""));
        assert!(first.contains("id=\"1000001\""));
        assert!(second.contains("id=\"1000002\""));
        assert_eq!(store.parts().len(), 1);
        assert_eq!(store.relationships().matches("<Relationship ").count(), 1);
        assert!(store.parts()[0].1.starts_with(b"\x89PNG"));
        let mut truncated = Store::read(&word[..1024], &table, 20).unwrap();
        assert!(truncated.drawing(12).is_err());
    }
    #[cfg(feature = "direct-doc")]
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

    #[cfg(feature = "direct-doc")]
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
        // Package conversion still omits every aligned position.
        let (word, table) = drawing_with_options(0xa00, 0, &[(0x38f, 2), (0x390, 1)]);
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.drawing(12).unwrap().is_empty() && store.omitted);
    }

    #[cfg(feature = "direct-doc")]
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

    #[cfg(feature = "direct-doc")]
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

    #[cfg(feature = "direct-doc")]
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
            let xml = store.drawing(12).unwrap();
            assert!(xml.contains("<wp:anchor"));
            assert!(xml.contains("<wp:positionH relativeFrom=\"page\">"));
            assert!(xml.contains("<wp:positionV relativeFrom=\"page\">"));
            assert!(xml.contains("cx=\"254000\" cy=\"190500\""));
            assert!(xml.contains("<wp:wrapSquare wrapText=\"bothSides\"/>"));
            assert_eq!(store.parts().len(), 1);
            if let Some(expected) = expected {
                assert_eq!(store.parts()[0].1, expected);
            } else {
                assert!(store.parts()[0].1.starts_with(b"\x89PNG"));
            }
            assert_eq!(store.relationships().matches("<Relationship ").count(), 1);
            assert!(!store.omitted);
        }
    }
    #[test]
    fn structural_hidden_and_script_flags_still_prevent_blip_access() {
        for flag in [0x01, 0x02, 0x04, 0x08, 0x100] {
            for ole in [0, 0x10] {
                let (word, table) = drawing_input(0xa00 | flag | ole, 0);
                let mut store = Store::read(&word[..1024], &table, 20).unwrap();
                assert!(store.drawing(12).unwrap().is_empty());
                assert!(store.parts().is_empty());
                assert!(store.omitted);
            }
        }
        for group in [0x00020002, 0x00800080] {
            let (word, table) = drawing_input(0xa10, group);
            let mut store = Store::read(&word[..1024], &table, 20).unwrap();
            assert!(store.drawing(12).unwrap().is_empty());
            assert!(store.parts().is_empty());
            assert!(store.omitted);
        }
    }
    #[test]
    fn ambiguous_alignment_and_rotated_shapes_do_not_dereference_blips() {
        for key in [0x38f, 0x391, 4] {
            for value in 1..=5 {
                let operand = if key == 4 { value << 16 } else { value };
                let (word, table) = drawing_with_options(0xa00, 0, &[(key, operand)]);
                let mut store = Store::read(&word[..1024], &table, 20).unwrap();
                assert!(store.drawing(12).unwrap().is_empty());
                assert!(store.parts().is_empty());
                assert!(store.omitted);
            }
        }
    }
    #[test]
    fn enforces_media_occurrence_and_operation_budgets() {
        let (word, table) = drawing_input(0xa00, 0);
        let mut store = Store::read(&word, &table, 20).unwrap();
        store.remaining_bytes = 0;
        assert!(store.drawing(12).unwrap_err().contains("media budget"));
        assert!(store.parts().is_empty());

        let mut store = Store::read(&word, &table, 20).unwrap();
        store.occurrences = 100_000;
        assert!(store.drawing(12).unwrap_err().contains("occurrence budget"));

        let mut store = Store::read(&word, &table, 20).unwrap();
        store.budget = 0;
        assert!(store.drawing(12).is_err());
    }
    #[test]
    fn does_not_reassign_header_or_nested_group_drawings_to_main_story() {
        let (word, mut table) = drawing_input(0xa00, 0);
        let start = u32_at(&word, 0x22a).unwrap() as usize;
        let (_, group_end) = record_with_end(&table, start, &mut 100, "test").unwrap();
        table[group_end] = 1;
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.drawing(12).unwrap().is_empty());
        assert!(store.parts().is_empty());

        table[group_end] = 0;
        // Replace the independent SpContainer tag with a nested SpgrContainer
        // lacking its group shape; it is never flattened into the story.
        let shape_start = group_end + 1 + 8 + 8;
        table[shape_start + 2..shape_start + 4].copy_from_slice(&0xf003u16.to_le_bytes());
        let mut store = Store::read(&word, &table, 20).unwrap();
        assert!(store.drawing(12).unwrap().is_empty());
        assert!(store.parts().is_empty());
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
        assert_eq!(
            position("margin", 999, false).unwrap(),
            "<wp:positionH relativeFrom=\"margin\"><wp:posOffset>999</wp:posOffset></wp:positionH>"
        );
        assert!(position("page", i64::MAX, false).is_err());
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
