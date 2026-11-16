//! BIFF8 drawing shapes and sheet-anchored groups -> XLSX-model shape
//! anchors, without generating SpreadsheetML. MS-XLS 2.4.181 (Obj ot 0-4, 6,
//! 9 and 30), 2.4.329 (TxO), 2.5.129 (FontIndex); MS-ODRAW 2.2.12 (drawing
//! defaults), 2.2.40 (FSP), 2.3 (properties) and 2.4.24 (shape types).
//!
//! Excel's own XLSX of the corpus workbooks these XLS files were saved from is
//! the mapping evidence: rectangles (msosptRectangle) and text boxes
//! (msosptTextBox) are `prstGeom rect` shapes with the same solid fill and
//! line, text box runs carry the BIFF Font's face, size, weight and palette
//! color, TxO alignment is the paragraph `algn` and body `anchor`, and a group
//! keeps its members' child-anchor placement. Other shape types are the
//! DrawingML presets Office converts them to (`officeart::preset`), a shape
//! with its own path is that path, and fills are the model's fills (`fill`).
//!
//! What the shape model cannot state is drawn as its nearest equivalent and
//! recorded with the best-effort helpers (`crate::rejection`): a shadow and a
//! property without a counterpart are not drawn, underline, strikethrough
//! and the other font effects of a text run are left out, a shape type
//! without a preset is a rectangle. Form controls are not drawn.
use super::drawing_anchors::{self, CellCorner, GroupFrame, ShapeSource};
use super::{styles, u16_at, unsupported, Record, SheetData};
use crate::officeart::{paint::Paint, record_with_end};
use ooxml_common::drawing::{DrawingGroupSpec, DrawingGroupTransform, DrawingRect};
use std::collections::BTreeMap;

mod fill;
mod pattern;

/// MS-ODRAW 2.3.21.2-5 default text margins (0.1 in, 0.05 in). Excel writes
/// fAutoTextMargin (2.3.21.15, <57>) for text whose XLSX counterpart has a
/// `bodyPr` without insets, i.e. the DrawingML defaults (ECMA-376
/// 21.1.2.1.1), which are these same values.
const DEFAULT_MARGINS: [u32; 4] = [91_440, 45_720, 91_440, 45_720];

#[derive(Default)]
pub(super) struct Shapes {
    sheets: BTreeMap<usize, Vec<Prepared>>,
    /// Store entries of grouped pictures (their media is loaded with the
    /// sheet pictures', see `attach_images`).
    images: std::collections::BTreeSet<u32>,
}

struct Prepared {
    from: CellCorner,
    to: CellCorner,
    behavior: u16,
    leaves: Vec<PreparedLeaf>,
}

struct PreparedLeaf {
    info: xlsx_model::ShapeInfo,
    placement: Placement,
    /// A preset shape's type and legacy adjust values, converted once the
    /// shape's extent is known (`officeart::preset::adjustments`).
    preset: Option<(u16, [Option<i32>; 10])>,
    /// The picture its fill waits for.
    blip: Option<fill::Blip>,
}

/// One projected shape: placement is resolved with its anchor.
struct Projected {
    info: xlsx_model::ShapeInfo,
    rotation: i32,
    preset: Option<(u16, [Option<i32>; 10])>,
    blip: Option<fill::Blip>,
}

/// The object types drawn as shapes (MS-XLS 2.5.FtCmo ot): line, rectangle,
/// oval, arc, text box, polygon and OfficeArt shape. Each is drawn from its
/// OfficeArt shape type and properties.
pub(super) fn shape_object(kind: u16) -> bool {
    matches!(kind, 1 | 2 | 3 | 4 | 6 | 9 | 30)
}

/// Form controls (button, check box, option button, edit box, label, dialog
/// frame, spinner, scroll bar, list box, group box, drop-down). The model
/// has no control, so none is drawn.
pub(super) fn form_control(kind: u16) -> bool {
    matches!(kind, 7 | 11..=20)
}

/// Where a leaf sits: the enclosing groups (outermost first, the first one
/// placed by the cell anchor) and its own child anchor, rotation and flips,
/// all as stored. Placement is resolved in EMU once the anchor size is known,
/// so rotation composes in true proportions.
struct Placement {
    groups: Vec<GroupFrame>,
    /// Child anchor in the innermost group; `None` for a shape that fills
    /// the cell anchor itself.
    anchor: Option<[i32; 4]>,
    rotation: i32,
    flip_h: bool,
    flip_v: bool,
}

/// OfficeArt placement -> DrawingML (MS-ODRAW 2.3.18.5 defines only a
/// clockwise angle about the centre). Two Office rules, established by
/// PowerPoint's 120-case rotation control for the PPT reader
/// (`ppt::drawing::direct_transform`) and matched by Excel's own XLSX of
/// sample workbooks (a -45 degree freeform keeps its unswapped child anchor
/// and a 15.76 degree group keeps its frame): the angle is negated when
/// exactly one flip is set, and an angle normalised into [45, 135) or
/// [225, 315) means the stored rectangle holds the rotated bounds, so width
/// and height swap about the same centre. A group's child coordinate space
/// is never swapped.
fn drawing_rect(rect: [f64; 4], rotation: i32, flip_h: bool, flip_v: bool) -> DrawingRect {
    let [x, y, w, h] = rect;
    let normalized = i64::from(rotation).rem_euclid(360 * 65_536);
    let swapped = (45 * 65_536..135 * 65_536).contains(&normalized)
        || (225 * 65_536..315 * 65_536).contains(&normalized);
    let (x, y, w, h) = if swapped {
        (x + (w - h) / 2.0, y + (h - w) / 2.0, h, w)
    } else {
        (x, y, w, h)
    };
    let degrees = f64::from(rotation) / 65_536.0;
    DrawingRect {
        x,
        y,
        width: w,
        height: h,
        rotation_degrees: if flip_h != flip_v { -degrees } else { degrees },
        flip_h,
        flip_v,
    }
}

fn corners(rect: [i32; 4]) -> [f64; 4] {
    let [left, top, right, bottom] = rect.map(f64::from);
    [left, top, right - left, bottom - top]
}

impl Placement {
    /// The leaf rectangle in anchor EMU (anchor at the origin, `cx`×`cy`).
    fn resolve(&self, cx: f64, cy: f64) -> DrawingRect {
        let mut transform = DrawingGroupTransform::IDENTITY;
        let mut parent = [0.0, 0.0, cx, cy];
        for group in &self.groups {
            let placed = drawing_rect(
                group.anchor.map_or(parent, corners),
                group.rotation,
                group.flip_h,
                group.flip_v,
            );
            let child = corners(group.rect);
            let spec = DrawingGroupSpec {
                off_x: placed.x,
                off_y: placed.y,
                ext_x: placed.width,
                ext_y: placed.height,
                child_off_x: child[0],
                child_off_y: child[1],
                child_ext_x: child[2],
                child_ext_y: child[3],
                rotation_degrees: placed.rotation_degrees,
                flip_h: placed.flip_h,
                flip_v: placed.flip_v,
            };
            transform = if transform == DrawingGroupTransform::IDENTITY {
                DrawingGroupTransform::from_group(spec)
            } else {
                transform.compose_group(spec)
            };
            parent = child;
        }
        let leaf = drawing_rect(
            self.anchor.map_or(parent, corners),
            self.rotation,
            self.flip_h,
            self.flip_v,
        );
        transform.apply_rect(leaf)
    }
}

impl Shapes {
    pub(super) fn prepare(
        records: &[Record<'_>],
        tabs: &[usize],
        styles: &styles::Styles<'_>,
    ) -> Result<Self, String> {
        let mut budget = super::rich::MAX_MODEL_BYTES;
        let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
        let mut prepared = Self::default();
        let mut defaults = None;
        for anchor in drawing_anchors::strict(records)? {
            let Some(&sheet) = sheet_ids.get(&anchor.sheet) else {
                continue;
            };
            if anchor.object_type != 0 && !shape_object(anchor.object_type) {
                continue;
            }
            let defaults = match &defaults {
                Some(table) => table,
                None => defaults.insert(drawing_defaults(records)?),
            };
            let source = anchor
                .shape
                .as_ref()
                .ok_or_else(|| unsupported("BIFF drawing shape without its properties"))?;
            let mut leaves = Vec::new();
            if anchor.object_type == 0 {
                if group_hidden(defaults, source)? {
                    continue;
                }
                for member in &anchor.members {
                    if let (8, Some(picture)) = (member.object_type, member.picture) {
                        prepared.images.insert(picture.store_index);
                        leaves.push(PreparedLeaf {
                            info: picture_leaf(member.order, &picture),
                            placement: Placement {
                                groups: member.groups.clone(),
                                anchor: Some(member.anchor),
                                rotation: picture.rotation,
                                flip_h: member.shape_flags & 0x40 != 0,
                                flip_v: member.shape_flags & 0x80 != 0,
                            },
                            preset: None,
                            blip: None,
                        });
                        continue;
                    }
                    // A grouped form control is not drawn (see
                    // `admit_direct_object`, which records it).
                    if form_control(member.object_type) {
                        continue;
                    }
                    if !shape_object(member.object_type) {
                        return Err(unsupported(format!(
                            "grouped BIFF drawing object type {} is not projected",
                            member.object_type
                        )));
                    }
                    let source = member
                        .shape
                        .as_ref()
                        .ok_or_else(|| unsupported("BIFF drawing shape without its properties"))?;
                    let leaf = Leaf {
                        flags: member.shape_flags,
                        child: true,
                        order: member.order,
                    };
                    if let Some(projected) =
                        leaf.project(records, styles, defaults, source, &mut budget)?
                    {
                        prepared
                            .images
                            .extend(projected.blip.as_ref().map(|blip| blip.store));
                        leaves.push(PreparedLeaf {
                            info: projected.info,
                            placement: Placement {
                                groups: member.groups.clone(),
                                anchor: Some(member.anchor),
                                rotation: projected.rotation,
                                flip_h: member.shape_flags & 0x40 != 0,
                                flip_v: member.shape_flags & 0x80 != 0,
                            },
                            preset: projected.preset,
                            blip: projected.blip,
                        });
                    }
                }
            } else {
                let leaf = Leaf {
                    flags: anchor.shape_flags,
                    child: false,
                    order: anchor.order,
                };
                if let Some(projected) =
                    leaf.project(records, styles, defaults, source, &mut budget)?
                {
                    prepared
                        .images
                        .extend(projected.blip.as_ref().map(|blip| blip.store));
                    leaves.push(PreparedLeaf {
                        info: projected.info,
                        placement: Placement {
                            groups: Vec::new(),
                            anchor: None,
                            rotation: projected.rotation,
                            flip_h: anchor.shape_flags & 0x40 != 0,
                            flip_v: anchor.shape_flags & 0x80 != 0,
                        },
                        preset: projected.preset,
                        blip: projected.blip,
                    });
                }
            }
            if !leaves.is_empty() {
                prepared.sheets.entry(sheet).or_default().push(Prepared {
                    from: anchor.from,
                    to: anchor.to,
                    behavior: anchor.behavior,
                    leaves,
                });
            }
        }
        Ok(prepared)
    }

    pub(super) fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }

    pub(super) fn picture_indices(&self) -> std::collections::BTreeSet<u32> {
        self.images.clone()
    }

    /// Give each grouped picture and each picture fill the MIME type of its
    /// decoded media, and each pattern fill the preset its bitmap draws.
    pub(super) fn attach_images(
        &mut self,
        pictures: &super::pictures::Pictures,
    ) -> Result<(), String> {
        for prepared in self.sheets.values_mut().flatten() {
            for leaf in &mut prepared.leaves {
                if let xlsx_model::ShapeGeom::Image {
                    image_path,
                    mime_type,
                    ..
                } = &mut leaf.info.geom
                {
                    let id = image_path
                        .rsplit('/')
                        .next()
                        .and_then(|id| id.parse().ok())
                        .ok_or_else(|| unsupported("invalid grouped XLS picture"))?;
                    let extension = pictures.extension(id).ok_or_else(|| {
                        unsupported("grouped BIFF picture BLIP is not a supported image")
                    })?;
                    *mime_type = ooxml_common::blip::mime_from_ext(extension).into();
                }
                let Some(blip) = leaf.blip.take() else {
                    continue;
                };
                let (extension, bytes) = pictures.media(blip.store).ok_or_else(|| {
                    unsupported("grouped BIFF picture BLIP is not a supported image")
                })?;
                match &mut leaf.info.fill {
                    Some(xlsx_model::ShapeFill::Pattern { preset, .. }) => {
                        *preset = fill::pattern_preset(bytes, extension);
                    }
                    Some(xlsx_model::ShapeFill::Image { mime_type, .. }) => {
                        *mime_type = ooxml_common::blip::mime_from_ext(extension).into();
                    }
                    _ => {}
                }
            }
        }
        Ok(())
    }

    /// Resolve MS-XLS 2.5.193 cell fractions to DrawingML cell offsets, as
    /// pictures and charts do.
    pub(super) fn resolve(
        self,
        sheets: &[(String, SheetData)],
        mdw: f64,
        warnings: &mut Vec<String>,
    ) -> BTreeMap<usize, Vec<xlsx_model::ShapeAnchor>> {
        let mut output = BTreeMap::new();
        let mut omitted = false;
        for (index, prepared) in self.sheets {
            let sheet = &sheets[index].1;
            if !sheet.geometry.has_sheet_defaults() || sheet.views.displays_formulas() {
                omitted = true;
                continue;
            }
            let max_row = prepared
                .iter()
                .map(|p| p.from.row.max(p.to.row))
                .max()
                .unwrap_or(0);
            let max_col = prepared
                .iter()
                .map(|p| p.from.column.max(p.to.column))
                .max()
                .unwrap_or(0);
            let columns = super::pictures::prefix(max_col, |c| sheet.geometry.column_emu(c, mdw));
            let rows = super::pictures::prefix(max_row, |r| sheet.geometry.row_emu(r));
            let locate = |corner: CellCorner| -> Option<(f64, f64, i64, i64)> {
                let column = usize::from(corner.column);
                let row = usize::from(corner.row);
                let x = columns[column]?;
                let y = rows[row]?;
                let dx = (columns[column + 1]? - x) * f64::from(corner.dx) / 1024.0;
                let dy = (rows[row + 1]? - y) * f64::from(corner.dy) / 256.0;
                Some((x, y, dx.round() as i64, dy.round() as i64))
            };
            let mut anchors = Vec::new();
            for shape in prepared {
                let (Some(from), Some(to)) = (locate(shape.from), locate(shape.to)) else {
                    omitted = true;
                    continue;
                };
                let edit_as = match shape.behavior {
                    0 => "twoCell",
                    2 => "oneCell",
                    3 => "absolute",
                    _ => {
                        omitted = true;
                        continue;
                    }
                };
                let cx = (to.0.round() as i64 + to.2) - (from.0.round() as i64 + from.2);
                let cy = (to.1.round() as i64 + to.3) - (from.1.round() as i64 + from.3);
                if cx <= 0 || cy <= 0 {
                    omitted = true;
                    continue;
                }
                let (width, height) = (cx as f64, cy as f64);
                let shapes = shape
                    .leaves
                    .into_iter()
                    .map(|leaf| {
                        let PreparedLeaf {
                            mut info,
                            placement,
                            preset,
                            ..
                        } = leaf;
                        let rect = placement.resolve(width, height);
                        if let (Some((kind, legacy)), xlsx_model::ShapeGeom::Preset { adj, .. }) =
                            (preset, &mut info.geom)
                        {
                            *adj = preset_adjustments(kind, &legacy, rect.width, rect.height);
                        }
                        info.x = rect.x / width;
                        info.y = rect.y / height;
                        info.w = rect.width / width;
                        info.h = rect.height / height;
                        info.rot = rect.rotation_degrees;
                        info.flip_h = rect.flip_h;
                        info.flip_v = rect.flip_v;
                        info
                    })
                    .collect();
                anchors.push(xlsx_model::ShapeAnchor {
                    from_col: u32::from(shape.from.column),
                    from_col_off: from.2,
                    from_row: u32::from(shape.from.row),
                    from_row_off: from.3,
                    to_col: u32::from(shape.to.column),
                    to_col_off: to.2,
                    to_row: u32::from(shape.to.row),
                    to_row_off: to.3,
                    edit_as: Some(edit_as.into()),
                    native_ext_cx: cx,
                    native_ext_cy: cy,
                    // Binary OfficeArt anchor, not an XML anchor element.
                    anchor_tag: None,
                    anchor_ext_cx: None,
                    anchor_ext_cy: None,
                    shapes,
                });
            }
            if !anchors.is_empty() {
                output.insert(index, anchors);
            }
        }
        if omitted {
            warnings.push("legacy-xls:unresolved-shape-geometry-omitted".into());
        }
        output
    }
}

/// Primary and tertiary FOPT values of one shape over the drawing defaults.
/// Scalars must agree when repeated; Boolean property sets merge by their use
/// bits (MS-ODRAW 2.3.1).
#[derive(Default, Clone)]
struct Table {
    values: BTreeMap<u16, u32>,
}

impl Table {
    fn add(&mut self, opid: u16, value: u32) -> Result<(), String> {
        let id = opid & 0x3fff;
        if opid & 0x8000 != 0 {
            // Name, description, hyperlink, tooltip and the alternate
            // DrawingML blob identify, link or duplicate the shape; they
            // never change its binary rendering.
            return match id {
                0x380 | 0x381 | 0x382 | 0x38d | 0x3a9 => Ok(()),
                // Geometry arrays and fillShadeColors: the value is the
                // array size and the data travels in `ShapeSource::complex`.
                0x145 | 0x146 | 0x151 | 0x152 | 0x155..=0x157 | 0x197 => {
                    self.values.insert(id, value);
                    Ok(())
                }
                // A property is framed by its own size, so one without a
                // counterpart in the shape model is passed over.
                _ => {
                    crate::rejection::skipped(
                        "XLS drawing shape complex property is not projected",
                    );
                    Ok(())
                }
            };
        }
        if opid & 0x4000 != 0 {
            // fBid: the value is a one-based BLIP store index. fillBlip is
            // the picture of a pattern, texture or picture fill.
            if id == 0x186 {
                self.values.insert(id, value);
            } else {
                crate::rejection::skipped("XLS drawing shape BLIP property is not projected");
            }
            return Ok(());
        }
        match self.values.get_mut(&id) {
            None => {
                self.values.insert(id, value);
            }
            Some(current) if id & 0x3f == 0x3f => {
                let (old_use, new_use) = (*current >> 16, value >> 16);
                if (*current ^ value) & old_use & new_use & 0xffff != 0 {
                    return Err(unsupported("conflicting XLS drawing Boolean properties"));
                }
                *current = ((old_use | new_use) << 16) | (*current & old_use) | (value & new_use);
            }
            Some(current) if *current != value => {
                return Err(unsupported("conflicting XLS drawing shape properties"));
            }
            Some(_) => {}
        }
        Ok(())
    }

    fn with_defaults(mut self, defaults: &Self) -> Self {
        for (&id, &value) in &defaults.values {
            match self.values.get_mut(&id) {
                None => {
                    self.values.insert(id, value);
                }
                Some(current) if id & 0x3f == 0x3f => {
                    // Members the shape does not set come from the defaults.
                    let shape_use = *current >> 16;
                    let inherited = (value >> 16) & !shape_use;
                    *current |= (inherited << 16) | (value & inherited);
                }
                Some(_) => {}
            }
        }
        self
    }

    fn boolean(&self, id: u16, bit: u32) -> Option<bool> {
        let value = *self.values.get(&id)?;
        (value & (1 << (bit + 16)) != 0).then_some(value & (1 << bit) != 0)
    }
}

/// The drawing group's drawingPrimaryOptions (MS-ODRAW 2.2.12): the default
/// properties of every shape in the workbook (MsoDrawingGroup, MS-XLS 2.4.170).
fn drawing_defaults(records: &[Record<'_>]) -> Result<Table, String> {
    let mut bytes = Vec::new();
    let mut active = false;
    for record in records.iter().take_while(|r| r.kind != super::EOF) {
        if record.kind == 0x00eb || (active && record.kind == 0x003c) {
            if bytes.len() + record.data.len() > 64 * 1024 * 1024 {
                return Err(crate::rejection::resource("oversized BIFF drawing group"));
            }
            bytes.extend_from_slice(record.data);
            active = true;
        } else {
            active = false;
        }
    }
    let mut table = Table::default();
    if bytes.is_empty() {
        return Ok(table);
    }
    let mut work = 1_000_000usize;
    let (root, end) = record_with_end(&bytes, 0, &mut work, "XLS drawing group")?;
    if root.kind != 0xf000 || root.version != 15 {
        return Err(unsupported("invalid BIFF drawing group"));
    }
    let mut at = 8;
    while at < end {
        let (record, next) = record_with_end(&bytes[..end], at, &mut work, "XLS drawing group")?;
        if record.kind == 0xf00b {
            crate::officeart::properties::visit(record, &mut work, |p| table.add(p.opid, p.value))?;
        }
        at = next;
    }
    Ok(table)
}

fn group_hidden(defaults: &Table, source: &ShapeSource) -> Result<bool, String> {
    let mut table = Table::default();
    for &(opid, value) in &source.properties {
        table.add(opid, value)?;
    }
    let table = table.with_defaults(defaults);
    for (&id, &value) in &table.values {
        match id {
            // Rotation belongs to the group frame (see `Placement`).
            0x0004 | 0x0040..=0x007f | 0x03bf => {}
            // Shape defaults that only matter to the group's members.
            0x0080..=0x00bf | 0x0180..=0x01ff => {}
            _ => {
                return Err(unsupported(format!(
                    "XLS drawing group property {id:#06x}={value:#x} is not projected"
                )))
            }
        }
    }
    Ok(table.boolean(0x3bf, 1) == Some(true))
}

struct Leaf {
    flags: u32,
    child: bool,
    order: u64,
}

impl Leaf {
    fn project(
        &self,
        records: &[Record<'_>],
        styles: &styles::Styles<'_>,
        defaults: &Table,
        source: &ShapeSource,
        budget: &mut usize,
    ) -> Result<Option<Projected>, String> {
        // MS-ODRAW 2.2.40: groups, patriarchs, deleted, OLE, master-linked
        // and background shapes need facts this projection lacks. A
        // connector (fConnector) is drawn as its shape type.
        let membership = if self.child { 0x2 } else { 0 };
        if self.flags & 0x43f != membership || self.flags & 0xa00 != 0xa00 {
            return Err(unsupported("XLS drawing shape has unsupported shape flags"));
        }
        // A shape with its own path (pVertices) is drawn from it whatever
        // its type; msosptNotPrimitive (0) is always its own path.
        let freeform = source.kind == 0;
        let custom = freeform || source.complex.iter().any(|(id, _)| *id == 0x145);
        let mut geometry = crate::officeart::geometry::Geometry::default();
        let mut table = Table::default();
        for &(opid, value) in &source.properties {
            table.add(opid, value)?;
        }
        let table = table.with_defaults(defaults);
        let mut paint = Paint::default();
        let mut margins = DEFAULT_MARGINS;
        let mut wrap = "square";
        let mut anchor_text = None;
        let mut direction = Direction::LeftToRight;
        let mut rotation = 0i32;
        let mut inscribed = false;
        let mut adjust = [None; 10];
        let mut compound = None;
        for (&id, &value) in &table.values {
            match id {
                // Transform (2.3.18.5): resolved with the placement.
                0x0004 => rotation = value as i32,
                // Protection (2.3.1-2.3.2) and the text identifier, whose
                // text Excel stores in the TxO record.
                0x0040..=0x0080 => {}
                0x0081..=0x0084 => {
                    if value > 0x132f540 {
                        return Err(unsupported("invalid XLS text margin"));
                    }
                    margins[usize::from(id - 0x81)] = value;
                }
                // WrapText (2.3.21.6): square or none.
                0x0085 => {
                    wrap = match value {
                        0 => "square",
                        2 => "none",
                        _ => return Err(unsupported("XLS text wrapping mode is not projected")),
                    }
                }
                // anchorText (2.3.21.8, used by Excel <52>): top, middle or
                // bottom; it must agree with the TxO alignment below.
                0x0087 if value <= 2 => anchor_text = Some(value),
                // hspNext (2.3.21.11) names a linked text box; zero is none.
                0x008a if value == 0 => {}
                // unused134, unused140 and unused141 (2.3.21.7/13/14) MUST
                // be ignored; 0x008E is not an MS-ODRAW property, which a
                // reader can ignore (MS-ODRAW 2.2: "If a property has an
                // unrecognized opid.opid field, a vendor can ignore the
                // property").
                0x0086 | 0x008c | 0x008d | 0x008e => {}
                // Geometry (2.3.6): the explicit path. A path that
                // references a guide or an escape is not decoded (below).
                0x0140..=0x0146 if custom => match source.complex.iter().find(|(c, _)| *c == id) {
                    Some((_, data)) => geometry.complex(id, data.as_slice()),
                    None => geometry.scalar(id, value)?,
                },
                // The coordinate space of a preset shape, which its preset
                // geometry does not use.
                0x0140..=0x0144 => {}
                // adjustValue..adjust10Value (2.3.6.10-19) of a preset
                // shape, converted with its extent.
                0x0147..=0x0150 => adjust[usize::from(id - 0x147)] = Some(value as i32),
                // Connection sites, their directions, adjust handles, guides
                // (which feed formulas) and the connection-site type
                // (2.3.6.18-30) only serve editing.
                0x0151 | 0x0152 | 0x0155 | 0x0156 | 0x0158 => {}
                // pInscribe (2.3.6.28): the text rectangle of the path.
                0x0157 => inscribed = value != 0,
                // Geometry Boolean Properties (2.3.6.31): fill/line vetoes.
                0x017f => paint.property(id, value)?,
                // txflTextFlow and cdirFont (2.3.21.9-10): horizontal text.
                // Vertical or rotated text flow is not projected.
                0x0088 | 0x0089 => {
                    if value != 0 {
                        return Err(unsupported(format!(
                            "XLS drawing shape property {id:#06x}={value:#x} is not projected"
                        )));
                    }
                }
                // txdir (2.3.21.13): left-to-right, right-to-left or from
                // the context (the first strong character).
                0x008b => {
                    direction = match value {
                        0 => Direction::LeftToRight,
                        1 => Direction::RightToLeft,
                        2 => Direction::Context,
                        _ => crate::rejection::spec_default(
                            "XLS drawing shape property is not projected",
                            Direction::LeftToRight,
                        ),
                    }
                }
                0x00bf => {}
                // Fill (2.3.7): type, colors, opacities, the gradient angle,
                // focus and shade type, and the texture tile.
                0x0180..=0x0184
                | 0x0189..=0x018c
                | 0x0195
                | 0x0198..=0x019c
                | 0x01bf
                | 0x01c0
                | 0x01c1
                | 0x01c4
                | 0x01cb
                | 0x01cc
                | 0x01ce
                | 0x01d0..=0x01d7
                | 0x01ff => paint.property(id, value)?,
                // fillBlip (stored by `Table::add` without its fBid flag).
                0x0186 => paint.fill_blip = Some(value),
                // fillToLeft..fillToBottom and fillShadeColors: read with
                // the gradient (`fill`). fillCrMod, the blip name and flags,
                // the shade preset (whose colors are stored) and the
                // extended colors (2.3.7.34-41, the scheme colors behind the
                // stored RGB values) change nothing that is drawn.
                0x0185 | 0x0187 | 0x0188 | 0x018d..=0x0190 | 0x0196 | 0x0197 | 0x019d..=0x01a8 => {}
                // lineStyle (2.3.8.16, MSOLINESTYLE): the compound line.
                0x01cd => {
                    compound = match value {
                        0 => None,
                        1 => Some("dbl"),
                        2 => Some("thickThin"),
                        3 => Some("thinThick"),
                        4 => Some("tri"),
                        _ => crate::rejection::spec_default(
                            "XLS drawing shape property is not projected",
                            None,
                        ),
                    }
                }
                // lineCrMod and the extended line colors (2.3.8.29-33).
                0x01c3 | 0x01d9..=0x01de => {}
                // Shadow (2.3.13): its color, offsets and softness matter
                // only when fShadow is set (below).
                0x0200..=0x023e => {}
                // cxstyle (2.3.2.8): how a connector is rerouted.
                0x0303 => {}
                // Shadow, Shape and Group Shape Boolean Properties
                // (MS-ODRAW 2.3.13.23, 2.3.2.12 and 2.3.4.44): checked below.
                0x023f | 0x033f | 0x03bf => {}
                // No counterpart in the shape model.
                _ => crate::rejection::skipped("XLS drawing shape property is not projected"),
            }
        }
        if table.boolean(0x3bf, 1) == Some(true) {
            return Ok(None);
        }
        // fShadow: the model's shapes have no shadow.
        if table.boolean(0x23f, 1) == Some(true) {
            crate::rejection::skipped("XLS drawing shadows are not projected");
        }
        for (id, bit, reason) in [
            (0x1ff, 9, "XLS opaque line background is not projected"),
            (0x1ff, 6, "XLS inset line pens are not projected"),
            (0x1ff, 0, "XLS no-line dash rendering is not projected"),
            // Shape Boolean Properties: background shapes and OLE icons. The
            // shape-type lock, relative-resize preference, rules initiator
            // and policy labels only affect editing.
            (0x33f, 0, "XLS background drawing shapes are not projected"),
            (0x33f, 5, "XLS OLE icon shapes are not projected"),
        ] {
            if table.boolean(id, bit) == Some(true) {
                return Err(unsupported(reason));
            }
        }
        if table.boolean(0x33f, 6).is_some() || table.boolean(0x33f, 7).is_some() {
            return Err(unsupported("XLS drawing flip overrides are not projected"));
        }
        if table.boolean(0xbf, 3) == Some(true) {
            margins = DEFAULT_MARGINS;
        }

        let (fill, blip) = if paint.filled.unwrap_or(true) && paint.fill_ok.unwrap_or(true) {
            let (fill, blip) = fill::project(&paint, &table, source, styles)?;
            (Some(fill), blip)
        } else {
            (None, None)
        };
        let line = if paint.lined.unwrap_or(true) && paint.line_ok.unwrap_or(true) {
            if paint.line_type.unwrap_or(0) != 0 {
                return Err(unsupported("non-solid XLS shape lines are not projected"));
            }
            Some(rgb(
                color(paint.line.unwrap_or(0), [0, 0, 0], styles),
                paint.line_alpha.unwrap_or(65_536),
            ))
        } else {
            None
        };
        let text = match source.text {
            Some(offset) => text(
                records,
                styles,
                offset,
                TextLayout {
                    anchor_text,
                    margins,
                    wrap,
                    direction,
                    table: &table,
                },
                budget,
            )?,
            None => None,
        };
        // pInscribe: the text rectangle inside the path. The model lays
        // text out in the shape's bounds.
        if custom && inscribed && text.is_some() {
            crate::rejection::approximated("XLS freeform text rectangles are not projected", ());
        }
        let decoded = if custom {
            let decoded = geometry.decode(&mut 1_000_000usize)?;
            if decoded.is_none() && freeform {
                return Err(unsupported(
                    "XLS freeform uses guide formulas or path escapes",
                ));
            }
            decoded
        } else {
            None
        };
        let (geom, preset) = match decoded {
            Some(decoded) => {
                // Each path carries its authored noFill/noLine escape
                // (ECMA-376 path@fill/@stroke, see `custom_geometry`). Whether
                // Excel fills an open subpath of a filled shape is not
                // established (PowerPoint does not), so such a path is rejected.
                if fill.is_some()
                    && decoded
                        .paths()
                        .any(|path| path.authored_fill() && !path.fill())
                {
                    return Err(unsupported(
                        "XLS filled freeforms with open paths are not projected",
                    ));
                }
                (custom_geometry(&decoded), None)
            }
            None => {
                // MS-ODRAW 2.4.24 shape type as the preset Office converts
                // it to. A type without one (or a typed shape whose own
                // path is not decoded) is drawn as its bounding rectangle.
                let name = match crate::officeart::preset::name(source.kind).filter(|_| !custom) {
                    Some(name) => name,
                    None => crate::rejection::approximated(
                        "XLS drawing shape type is not projected",
                        "rect",
                    ),
                };
                (
                    xlsx_model::ShapeGeom::Preset {
                        name: name.into(),
                        adj: Vec::new(),
                    },
                    (!custom && adjust.iter().any(Option::is_some))
                        .then_some((source.kind, adjust)),
                )
            }
        };
        let (join, miter) = paint.details.join();
        let end = |index: usize| {
            line.as_ref()
                .and(paint.details.line_end(index))
                .map(|end| xlsx_model::ShapeLineEnd {
                    r#type: end.kind.to_owned(),
                    w: end.width.to_owned(),
                    len: end.length.to_owned(),
                })
        };
        let info = xlsx_model::ShapeInfo {
            z_order: self.order,
            // Placement is resolved with the anchor size.
            x: 0.0,
            y: 0.0,
            w: 1.0,
            h: 1.0,
            rot: 0.0,
            flip_h: false,
            flip_v: false,
            fill_color: fill.as_ref().and_then(fill::fallback_color),
            fill,
            stroke_width: line
                .as_ref()
                .map_or(0, |_| i64::from(paint.width.unwrap_or(9525))),
            stroke_dash_style: line.as_ref().and_then(|_| {
                paint
                    .dash
                    .and_then(crate::officeart::stroke::preset_dash)
                    .filter(|dash| *dash != "solid")
                    .map(str::to_owned)
            }),
            stroke_line_cap: line.as_ref().map(|_| paint.details.canvas_cap().to_owned()),
            stroke_line_join: line.as_ref().map(|_| join.to_owned()),
            stroke_miter_limit: line.as_ref().and(miter),
            stroke_fill: None,
            stroke_custom_dash: Vec::new(),
            stroke_alignment: None,
            stroke_cmpd: line.as_ref().and(compound).map(str::to_owned),
            // lineStartArrowhead is the DrawingML headEnd, lineEndArrowhead
            // its tailEnd (ECMA-376 20.1.8.38 and 20.1.8.57).
            stroke_head_end: end(0),
            stroke_tail_end: end(1),
            stroke_color: line,
            geom,
            text,
        };
        use ooxml_common::chart::RetainedBytes;
        charge_shape(
            budget,
            std::mem::size_of::<xlsx_model::ShapeInfo>()
                .saturating_add(usize::try_from(info.heap_bytes()).unwrap_or(usize::MAX)),
        )?;
        Ok(Some(Projected {
            info,
            rotation,
            preset,
            blip,
        }))
    }
}

/// MS-ODRAW 2.3.21.13 txdir.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Direction {
    LeftToRight,
    RightToLeft,
    Context,
}

/// The `a:avLst` values of a preset shape from its legacy adjust values and
/// its extent in EMU. A shape type whose adjust values have no established
/// DrawingML counterpart keeps the preset's default proportions.
fn preset_adjustments(
    kind: u16,
    legacy: &[Option<i32>; 10],
    width: f64,
    height: f64,
) -> Vec<Option<f64>> {
    let converted = crate::officeart::preset::adjustments(
        kind,
        legacy,
        width.round() as i64,
        height.round() as i64,
    );
    match converted {
        Ok(Some(values)) => {
            let used = values
                .iter()
                .rposition(Option::is_some)
                .map_or(0, |i| i + 1);
            values[..used].to_vec()
        }
        Ok(None) => Vec::new(),
        Err(_) => {
            crate::rejection::approximated("XLS drawing shape type is not projected", Vec::new())
        }
    }
}

/// Decoded OfficeArt paths as the shape model's custom geometry, in the
/// path's own coordinate space (MS-ODRAW 2.3.6.1-4 geoLeft..geoBottom).
fn custom_geometry(decoded: &crate::officeart::geometry::Decoded) -> xlsx_model::ShapeGeom {
    use crate::officeart::geometry::DecodedCommand;
    let (w, h) = (decoded.width() as f64, decoded.height() as f64);
    let paths = decoded
        .paths()
        .map(|path| xlsx_model::PathInfo {
            w,
            h,
            // MS-ODRAW segment escapes (noFill/noLine) per path, as
            // ECMA-376 §20.1.9.15 path@fill="none" / path@stroke="0". The
            // authored flag is used: PowerPoint's open-path fill veto has no
            // Excel evidence.
            fill: (!path.authored_fill()).then(|| "none".to_owned()),
            stroke: path.stroke(),
            commands: path
                .commands()
                .map(|command| match command {
                    DecodedCommand::Move([x, y]) => xlsx_model::PathCmd::MoveTo {
                        x: x as f64,
                        y: y as f64,
                    },
                    DecodedCommand::Line([x, y]) => xlsx_model::PathCmd::LineTo {
                        x: x as f64,
                        y: y as f64,
                    },
                    DecodedCommand::Cubic([[x1, y1], [x2, y2], [x3, y3]]) => {
                        xlsx_model::PathCmd::CubicBezTo {
                            x1: x1 as f64,
                            y1: y1 as f64,
                            x2: x2 as f64,
                            y2: y2 as f64,
                            x3: x3 as f64,
                            y3: y3 as f64,
                        }
                    }
                    // MS-ODRAW quarter ellipse decoded in degree units; XLSX
                    // arcTo uses 60000ths of a degree (ECMA-376 20.1.9.3).
                    DecodedCommand::Arc { wr, hr, start, .. } => xlsx_model::PathCmd::ArcTo {
                        wr: wr as f64,
                        hr: hr as f64,
                        st_ang: f64::from(start) * 60_000.0,
                        sw_ang: -90.0 * 60_000.0,
                    },
                    DecodedCommand::Close => xlsx_model::PathCmd::Close,
                })
                .collect(),
        })
        .collect();
    xlsx_model::ShapeGeom::Custom { paths }
}

/// A grouped picture: the store entry's media, cropped as the sheet pictures
/// are (MS-ODRAW 2.3.23.1-4 cropFrom*, 16.16 fractions of the picture).
fn picture_leaf(order: u64, picture: &drawing_anchors::PictureReference) -> xlsx_model::ShapeInfo {
    let [top, bottom, left, right] = picture.crop.map(|value| f64::from(value) / 65_536.0);
    xlsx_model::ShapeInfo {
        z_order: order,
        x: 0.0,
        y: 0.0,
        w: 1.0,
        h: 1.0,
        rot: 0.0,
        flip_h: false,
        flip_v: false,
        fill_color: None,
        fill: None,
        stroke_color: None,
        stroke_width: 0,
        stroke_fill: None,
        stroke_dash_style: None,
        stroke_custom_dash: Vec::new(),
        stroke_line_cap: None,
        stroke_line_join: None,
        stroke_miter_limit: None,
        stroke_alignment: None,
        stroke_cmpd: None,
        stroke_head_end: None,
        stroke_tail_end: None,
        geom: xlsx_model::ShapeGeom::Image {
            image_path: format!("legacy-xls/image/{}", picture.store_index),
            mime_type: String::new(),
            svg_image_path: None,
            src_rect: picture.crop.iter().any(|value| *value != 0).then_some(
                ooxml_common::blip::SrcRect {
                    t: top,
                    b: bottom,
                    l: left,
                    r: right,
                },
            ),
            alpha: None,
            duotone: None,
            blip_effects: Vec::new(),
        },
        text: None,
    }
}

fn right_to_left(c: char) -> bool {
    matches!(u32::from(c), 0x0590..=0x08ff | 0xfb1d..=0xfdff | 0xfe70..=0xfeff)
}

/// OfficeArtCOLORREF (MS-ODRAW 2.2.2) as red, green and blue. A plain,
/// palette-matched (fPaletteRGB) or system-matched (fSystemRGB) color is its
/// own RGB value. With fSchemeIndex the red byte is an index into the
/// workbook's color table (MS-XLS 2.5.161 Icv: the Palette record and the
/// display colors, see `styles::Styles::drawing_color`); Excel writes the
/// default fill and line of every shape that way. A reference this reader
/// cannot resolve (a palette or system index, an Icv outside the table)
/// states no color, and the property takes `default`, the color MS-ODRAW
/// gives it.
fn color(value: u32, default: [u8; 3], styles: &styles::Styles<'_>) -> [u8; 3] {
    let [red, green, blue, flags] = value.to_le_bytes();
    let resolved = match flags {
        0x00 | 0x02 | 0x04 => Some([red, green, blue]),
        0x08 if green == 0 && blue == 0 => styles.drawing_color(u16::from(red)),
        _ => None,
    };
    resolved.unwrap_or_else(|| {
        crate::rejection::spec_default(
            "XLS drawing palette or scheme colors are not projected",
            default,
        )
    })
}

/// A color with a 16.16 opacity (fillOpacity 2.3.7.5, lineOpacity 2.3.8.2)
/// as the model's `#RRGGBB`, or `#RRGGBBAA` when translucent, with the alpha
/// byte and opaque threshold of the shared DrawingML color parser. Excel's
/// XLSX of a 50% line (`a:alpha 50000`) saves lineOpacity 0x8080, i.e. the
/// same alpha byte 0x80.
fn rgb([red, green, blue]: [u8; 3], alpha: u32) -> String {
    let mut hex = format!("#{red:02X}{green:02X}{blue:02X}");
    let alpha = f64::from(alpha) / 65_536.0;
    if (alpha - 1.0).abs() >= 0.004 {
        hex.push_str(&format!(
            "{:02X}",
            (alpha.clamp(0.0, 1.0) * 255.0).round() as u8
        ));
    }
    hex
}

/// Resource policy admits amplified text work and complete retained shapes
/// against one workbook quota; OfficeArt record sizes alone do not bound it.
fn charge_shape(budget: &mut usize, bytes: usize) -> Result<(), String> {
    *budget = budget
        .checked_sub(bytes)
        .ok_or_else(|| crate::rejection::resource("XLS shape model retention budget exceeded"))?;
    Ok(())
}

fn charge_run(budget: &mut usize, font: &RunFont, text_bytes: usize) -> Result<(), String> {
    charge_shape(
        budget,
        std::mem::size_of::<xlsx_model::ShapeTextRun>()
            .saturating_add(font.name.len().saturating_mul(2))
            .saturating_add(font.color.len())
            .saturating_add(text_bytes),
    )
}

/// The `characters` UTF-16 code units of the TxO at `records[index]`, from
/// the Continue records that follow it (each an XLUnicodeStringNoCch
/// fragment with its own fHighByte), and the index of the record after them.
fn txo_characters(
    records: &[Record<'_>],
    index: usize,
    characters: usize,
) -> Result<(Vec<u16>, usize), String> {
    let mut next = index + 1;
    let mut units: Vec<u16> = Vec::with_capacity(characters);
    while units.len() < characters {
        let record = records
            .get(next)
            .filter(|record| record.kind == 0x003c && !record.data.is_empty())
            .ok_or_else(|| unsupported("truncated BIFF TxO text"))?;
        let (flag, chars) = (record.data[0], &record.data[1..]);
        match flag {
            0 => units.extend(chars.iter().map(|&byte| u16::from(byte))),
            1 if chars.len() % 2 == 0 => units.extend(
                chars
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
            ),
            _ => return Err(unsupported("invalid BIFF TxO text fragment")),
        }
        if units.len() > characters {
            return Err(unsupported("BIFF TxO text longer than cchText"));
        }
        next += 1;
    }
    Ok((units, next))
}

/// The plain text of the TxO record at stream offset `offset`.
pub(super) fn txo_text(records: &[Record<'_>], offset: usize) -> Result<String, String> {
    let index = records
        .binary_search_by_key(&offset, |record| record.offset)
        .map_err(|_| unsupported("BIFF TxO record is not a record boundary"))?;
    let txo = records[index].data;
    if txo.len() < 18 {
        return Err(unsupported("truncated BIFF TxO record"));
    }
    let characters = usize::from(u16_at(txo, 10)?);
    let run_bytes = usize::from(u16_at(txo, 12)?);
    // cbFmla: the ObjFmla (2.5.187) that follows the fixed fields.
    let formula = usize::from(u16_at(txo, 16)?);
    if txo.len() != 18 + formula {
        return Err(unsupported("truncated BIFF TxO record"));
    }
    if !drawing_anchors::object_formula_framed(&txo[18..]) {
        return Err(unsupported("truncated BIFF TxO record"));
    }
    let (units, mut next) = txo_characters(records, index, characters)?;
    // The formatting runs are not projected, but they are part of the TxO:
    // FormatRun (2.5.132) entries from character 0 in ascending order, then
    // TxOLastRun (2.5.271) at cchText; none for an empty text.
    let invalid = || unsupported("invalid BIFF TxO formatting runs");
    if characters == 0 {
        return if run_bytes == 0 {
            Ok(String::new())
        } else {
            Err(invalid())
        };
    }
    if run_bytes < 16 || run_bytes % 8 != 0 {
        return Err(invalid());
    }
    let mut runs = Vec::with_capacity(run_bytes);
    while runs.len() < run_bytes {
        let record = records
            .get(next)
            // An empty Continue carries no run and would not advance.
            .filter(|record| record.kind == 0x003c && !record.data.is_empty())
            .ok_or_else(|| unsupported("truncated BIFF TxO formatting runs"))?;
        runs.extend_from_slice(record.data);
        next += 1;
    }
    if runs.len() != run_bytes {
        return Err(invalid());
    }
    let positions: Vec<usize> = runs
        .as_chunks::<8>()
        .0
        .iter()
        .map(|run| usize::from(u16::from_le_bytes([run[0], run[1]])))
        .collect();
    let (last, starts) = positions.split_last().expect("at least two runs");
    if *last != characters
        || starts.first() != Some(&0)
        || starts.windows(2).any(|pair| pair[0] >= pair[1])
        || starts.last().is_some_and(|start| *start >= characters)
    {
        return Err(invalid());
    }
    Ok(String::from_utf16_lossy(&units))
}

struct TextLayout<'a> {
    anchor_text: Option<u32>,
    margins: [u32; 4],
    wrap: &'a str,
    direction: Direction,
    table: &'a Table,
}

/// TxO (MS-XLS 2.4.329): the text string in Continue records of
/// XLUnicodeStringNoCch fragments, then its TxORuns formatting runs.
fn text(
    records: &[Record<'_>],
    styles: &styles::Styles<'_>,
    offset: usize,
    layout: TextLayout<'_>,
    budget: &mut usize,
) -> Result<Option<xlsx_model::ShapeText>, String> {
    let TextLayout {
        anchor_text,
        margins,
        wrap,
        direction,
        table,
    } = layout;
    let index = records
        .binary_search_by_key(&offset, |record| record.offset)
        .map_err(|_| unsupported("BIFF TxO record is not a record boundary"))?;
    let txo = records[index].data;
    if txo.len() < 18 {
        return Err(unsupported("truncated BIFF TxO record"));
    }
    let grbit = u16_at(txo, 0)?;
    let rotation = u16_at(txo, 2)?;
    let characters = usize::from(u16_at(txo, 10)?);
    let run_bytes = usize::from(u16_at(txo, 12)?);
    if u16_at(txo, 16)? != 0 {
        return Err(unsupported(
            "XLS shape text linked to a formula is not projected",
        ));
    }
    if rotation != 0 {
        return Err(unsupported("rotated XLS shape text is not projected"));
    }
    let align = match (grbit >> 1) & 7 {
        1 => "l",
        2 => "ctr",
        3 => "r",
        4 => "just",
        7 => "dist",
        _ => return Err(unsupported("invalid XLS text horizontal alignment")),
    };
    let (anchor, vertical) = match (grbit >> 4) & 7 {
        1 => ("t", 0),
        2 => ("ctr", 1),
        3 => ("b", 2),
        4 | 7 => {
            return Err(unsupported(
                "justified XLS vertical text alignment is not projected",
            ))
        }
        _ => return Err(unsupported("invalid XLS text vertical alignment")),
    };
    if anchor_text.is_some_and(|value| value != vertical) {
        return Err(unsupported(
            "XLS text anchor disagrees with its TxO alignment",
        ));
    }
    if characters == 0 {
        if run_bytes != 0 {
            return Err(unsupported("invalid BIFF TxO formatting runs"));
        }
        return Ok(None);
    }
    if run_bytes < 16 || run_bytes % 8 != 0 {
        return Err(unsupported("invalid BIFF TxO formatting runs"));
    }
    let (units, mut next) = txo_characters(records, index, characters)?;
    let mut runs_data = Vec::with_capacity(run_bytes);
    while runs_data.len() < run_bytes {
        let record = records
            .get(next)
            .filter(|record| record.kind == 0x003c)
            .ok_or_else(|| unsupported("truncated BIFF TxO formatting runs"))?;
        runs_data.extend_from_slice(record.data);
        next += 1;
    }
    if runs_data.len() != run_bytes {
        return Err(unsupported("invalid BIFF TxO formatting runs"));
    }
    // FormatRun (2.5.132) entries, then TxOLastRun (2.5.271) at cchText.
    let mut runs = Vec::new();
    for entry in runs_data.as_chunks::<8>().0 {
        runs.push((usize::from(u16_at(entry, 0)?), u16_at(entry, 2)?));
    }
    let (last, _) = runs.pop().expect("at least two runs");
    if last != characters
        || runs.first().map(|run| run.0) != Some(0)
        || runs.windows(2).any(|pair| pair[0].0 >= pair[1].0)
        || runs.last().is_some_and(|run| run.0 >= characters)
    {
        return Err(unsupported("invalid BIFF TxO formatting runs"));
    }
    let mut fonts = BTreeMap::new();
    for &(_, font) in &runs {
        if let std::collections::btree_map::Entry::Vacant(entry) = fonts.entry(font) {
            entry.insert(run_font(styles, font)?);
        }
    }
    let font_at = |position: usize| {
        let run = runs.partition_point(|run| run.0 <= position).max(1) - 1;
        &fonts[&runs[run].1]
    };
    let make_run = |text: String, font: &RunFont| xlsx_model::ShapeTextRun::Text {
        text,
        bold: font.bold,
        italic: font.italic,
        size: font.size,
        color: Some(font.color.clone()),
        // One BIFF face serves every script of the run.
        font_face: Some(font.name.clone()),
        font_face_ea: Some(font.name.clone()),
        font_face_cs: None,
        font_face_sym: None,
        spacing: None,
    };
    let mut paragraphs = Vec::new();
    let mut start = 0usize;
    loop {
        let end = units[start..]
            .iter()
            .position(|&unit| unit == 0x000a)
            .map_or(characters, |at| start + at);
        // TxO cchText bounds encoded text, not the paragraph/run expansion.
        // Admit each generated slot and repeated font string before allocating.
        charge_shape(
            budget,
            std::mem::size_of::<xlsx_model::ShapeParagraph>() + align.len(),
        )?;
        let mut paragraph_runs = Vec::new();
        if start == end {
            // An empty line keeps the height of the font at its position.
            let font = font_at(start);
            charge_run(budget, font, 0)?;
            paragraph_runs.push(make_run(String::new(), font));
        }
        let mut at = start;
        while at < end {
            let run = runs.partition_point(|run| run.0 <= at) - 1;
            let run_end = runs.get(run + 1).map_or(characters, |next| next.0).min(end);
            // UTF-8 occupies at most three bytes per UTF-16 code unit.
            charge_run(budget, &fonts[&runs[run].1], (run_end - at) * 3)?;
            let text = String::from_utf16(&units[at..run_end])
                .map_err(|_| unsupported("invalid UTF-16 in BIFF TxO text"))?;
            paragraph_runs.push(make_run(text, &fonts[&runs[run].1]));
            at = run_end;
        }
        // ECMA-376 21.1.2.2.7 `a:pPr@rtl`. With txdir from the context the
        // first strong character of the paragraph decides, as the Unicode
        // bidirectional algorithm's paragraph level does.
        let rtl = match direction {
            Direction::LeftToRight => false,
            Direction::RightToLeft => true,
            Direction::Context => char::decode_utf16(units[start..end].iter().copied())
                .filter_map(Result::ok)
                .find_map(|c| {
                    if right_to_left(c) {
                        Some(true)
                    } else if c.is_alphabetic() {
                        Some(false)
                    } else {
                        None
                    }
                })
                .unwrap_or(false),
        };
        paragraphs.push(xlsx_model::ShapeParagraph {
            align: align.into(),
            rtl,
            mar_l: None,
            mar_r: None,
            indent: None,
            space_line: None,
            space_before: None,
            space_after: None,
            def_tab_sz: None,
            tab_stops: Vec::new(),
            runs: paragraph_runs,
        });
        if end == characters {
            break;
        }
        start = end + 1;
    }
    Ok(Some(xlsx_model::ShapeText {
        anchor: anchor.into(),
        wrap: wrap.into(),
        auto_fit: if table.boolean(0xbf, 1) == Some(true) {
            "sp".into()
        } else {
            "none".into()
        },
        font_scale: None,
        ln_spc_reduction: None,
        l_ins: i64::from(margins[0]),
        t_ins: i64::from(margins[1]),
        r_ins: i64::from(margins[2]),
        b_ins: i64::from(margins[3]),
        // TxO/OfficeArt facts already admitted above do not carry these
        // optional DrawingML body properties; use the model defaults.
        anchor_ctr: false,
        spc_first_last_para: false,
        vert: "horz".into(),
        paragraphs,
    }))
}

struct RunFont {
    name: String,
    size: f64,
    bold: bool,
    italic: bool,
    color: String,
}

fn run_font(styles: &styles::Styles<'_>, index: u16) -> Result<RunFont, String> {
    let font = styles.shape_font(index)?;
    // The model's text runs carry weight, slant, size, color and face.
    if font.underline || font.strike || font.other_effects {
        crate::rejection::approximated(
            "XLS shape text underline, strikethrough or font effects are not projected",
            (),
        );
    }
    if font.size_twips == 0 || font.name.is_empty() {
        return Err(unsupported("XLS shape text font weight is not projected"));
    }
    // FONT bls (MS-XLS 2.4.122): 400 is normal and 700 bold. The model has
    // one bold flag, so another weight is the nearer of the two.
    let bold = match font.weight {
        400 => false,
        700 => true,
        weight => crate::rejection::approximated(
            "XLS shape text font weight is not projected",
            weight >= 550,
        ),
    };
    // An Icv of the Palette, or a display color: the automatic font color
    // (0x7FFF) is the window text color.
    let [red, green, blue] = match styles.drawing_color(font.color_index) {
        Some(color) => color,
        None => crate::rejection::spec_default(
            "XLS shape text system colors are not projected",
            [0, 0, 0],
        ),
    };
    Ok(RunFont {
        name: font.name,
        size: f64::from(font.size_twips) / 20.0,
        bold,
        italic: font.italic,
        color: format!("#{red:02X}{green:02X}{blue:02X}"),
    })
}

#[cfg(test)]
mod tests;
