//! BIFF8 rectangles, text boxes and sheet-anchored groups -> XLSX-model shape
//! anchors, without generating SpreadsheetML. MS-XLS 2.4.181 (Obj ot 0, 2 and
//! 6), 2.4.329 (TxO), 2.5.129 (FontIndex); MS-ODRAW 2.2.12 (drawing defaults),
//! 2.2.40 (FSP) and 2.3 (properties).
//!
//! Excel's own XLSX of the corpus workbooks these XLS files were saved from is
//! the mapping evidence: rectangles (msosptRectangle) and text boxes
//! (msosptTextBox) are `prstGeom rect` shapes with the same solid fill and
//! line, text box runs carry the BIFF Font's face, size, weight and palette
//! color, TxO alignment is the paragraph `algn` and body `anchor`, and a group
//! keeps its members' child-anchor placement. Every other drawn fact is
//! rejected, never approximated.
use super::drawing_anchors::{self, CellCorner, GroupFrame, ShapeSource};
use super::{styles, u16_at, unsupported, Record, SheetData};
use crate::officeart::{paint::Paint, record_with_end};
use ooxml_common::drawing::{DrawingGroupSpec, DrawingGroupTransform, DrawingRect};
use std::collections::BTreeMap;

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
    leaves: Vec<(xlsx_model::ShapeInfo, Placement)>,
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
        let sheet_ids: BTreeMap<_, _> = tabs.iter().enumerate().map(|(i, &tab)| (tab, i)).collect();
        let mut prepared = Self::default();
        let mut defaults = None;
        for anchor in drawing_anchors::strict(records)? {
            let Some(&sheet) = sheet_ids.get(&anchor.sheet) else {
                continue;
            };
            if !matches!(anchor.object_type, 0 | 2 | 6 | 9) {
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
                        leaves.push((
                            picture_leaf(member.order, &picture),
                            Placement {
                                groups: member.groups.clone(),
                                anchor: Some(member.anchor),
                                rotation: picture.rotation,
                                flip_h: member.shape_flags & 0x40 != 0,
                                flip_v: member.shape_flags & 0x80 != 0,
                            },
                        ));
                        continue;
                    }
                    if !matches!(member.object_type, 2 | 6 | 9) {
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
                    if let Some((info, rotation)) =
                        leaf.project(records, styles, defaults, source)?
                    {
                        leaves.push((
                            info,
                            Placement {
                                groups: member.groups.clone(),
                                anchor: Some(member.anchor),
                                rotation,
                                flip_h: member.shape_flags & 0x40 != 0,
                                flip_v: member.shape_flags & 0x80 != 0,
                            },
                        ));
                    }
                }
            } else {
                let leaf = Leaf {
                    flags: anchor.shape_flags,
                    child: false,
                    order: anchor.order,
                };
                if let Some((info, rotation)) = leaf.project(records, styles, defaults, source)? {
                    leaves.push((
                        info,
                        Placement {
                            groups: Vec::new(),
                            anchor: None,
                            rotation,
                            flip_h: anchor.shape_flags & 0x40 != 0,
                            flip_v: anchor.shape_flags & 0x80 != 0,
                        },
                    ));
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

    /// Give each grouped picture the MIME type of its decoded media.
    pub(super) fn attach_images(
        &mut self,
        pictures: &super::pictures::Pictures,
    ) -> Result<(), String> {
        for prepared in self.sheets.values_mut().flatten() {
            for (info, _) in &mut prepared.leaves {
                if let xlsx_model::ShapeGeom::Image {
                    image_path,
                    mime_type,
                    ..
                } = &mut info.geom
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
                    .map(|(mut info, placement)| {
                        let rect = placement.resolve(width, height);
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
                // Geometry arrays: the value is the array size and the data
                // travels in `ShapeSource::complex`.
                0x145 | 0x146 | 0x151 | 0x152 | 0x155..=0x157 => {
                    self.values.insert(id, value);
                    Ok(())
                }
                _ => Err(unsupported(format!(
                    "XLS drawing shape complex property {id:#06x} is not projected"
                ))),
            };
        }
        if opid & 0x4000 != 0 {
            return Err(unsupported(format!(
                "XLS drawing shape BLIP property {id:#06x} is not projected"
            )));
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
                return Err(unsupported("oversized BIFF drawing group"));
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
    ) -> Result<Option<(xlsx_model::ShapeInfo, i32)>, String> {
        // MS-ODRAW 2.2.40: groups, patriarchs, deleted, OLE, master-linked,
        // connector and background shapes need facts this projection lacks.
        let membership = if self.child { 0x2 } else { 0 };
        if self.flags & 0x53f != membership || self.flags & 0xa00 != 0xa00 {
            return Err(unsupported("XLS drawing shape has unsupported shape flags"));
        }
        // msosptNotPrimitive (0) is a freeform drawn from its own path.
        if !matches!(source.kind, 0 | 1 | 202) {
            return Err(unsupported(format!(
                "XLS drawing shape type {} is not projected",
                source.kind
            )));
        }
        let freeform = source.kind == 0;
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
        let mut context_direction = false;
        let mut rotation = 0i32;
        let mut inscribed = false;
        for (&id, &value) in &table.values {
            match id {
                // Transform (2.3.18.5): resolved with the placement.
                0x0004 => rotation = value as i32,
                // Protection (2.3.1-2.3.2) and the text identifier, whose
                // text Excel stores in the TxO record.
                0x0040..=0x007f | 0x0080 => {}
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
                // Freeform geometry (2.3.6): the explicit path. A path that
                // references a guide or an escape is rejected below.
                0x0140..=0x0146 if freeform => {
                    match source.complex.iter().find(|(c, _)| *c == id) {
                        Some((_, data)) => geometry.complex(id, data.as_slice()),
                        None => geometry.scalar(id, value)?,
                    }
                }
                // Connection sites, their directions, adjust handles and the
                // connection-site type (2.3.6.18-30) only serve editing.
                0x0151 | 0x0152 | 0x0155 | 0x0158 if freeform => {}
                // Guides feed formulas; a path using them is not decoded.
                0x0156 if freeform => {}
                // pInscribe (2.3.6.28): the text rectangle of the path.
                0x0157 if freeform => inscribed = value != 0,
                // Geometry Boolean Properties (2.3.6.31): fill/line vetoes.
                0x017f => paint.property(id, value)?,
                0x0088 | 0x0089 if value == 0 => {}
                // txdir (2.3.21.13): left-to-right or from the context.
                0x008b if value == 0 => {}
                0x008b if value == 2 => context_direction = true,
                0x00bf => {}
                0x0180..=0x0183
                | 0x01bf
                | 0x01c0
                | 0x01c1
                | 0x01c4
                | 0x01cb
                | 0x01ce
                | 0x01d0..=0x01d7
                | 0x01ff => paint.property(id, value)?,
                0x01cd if value == 0 => {}
                // Shadow, Shape and Group Shape Boolean Properties
                // (MS-ODRAW 2.3.13.23, 2.3.2.12 and 2.3.4.44): checked below.
                0x023f | 0x033f | 0x03bf => {}
                _ => {
                    return Err(unsupported(format!(
                        "XLS drawing shape property {id:#06x}={value:#x} is not projected"
                    )))
                }
            }
        }
        if table.boolean(0x3bf, 1) == Some(true) {
            return Ok(None);
        }
        for (id, bit, reason) in [
            (0x23f, 1, "XLS drawing shadows are not projected"),
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

        let fill = if paint.filled.unwrap_or(true) && paint.fill_ok.unwrap_or(true) {
            if paint.fill_type.unwrap_or(0) != 0 || paint.fill_rect == Some(true) {
                return Err(unsupported("non-solid XLS shape fills are not projected"));
            }
            Some(rgb(
                paint.fill.unwrap_or(0x00ff_ffff),
                paint.fill_alpha.unwrap_or(65_536),
            )?)
        } else {
            None
        };
        let line = if paint.lined.unwrap_or(true) && paint.line_ok.unwrap_or(true) {
            if paint.line_type.unwrap_or(0) != 0 {
                return Err(unsupported("non-solid XLS shape lines are not projected"));
            }
            if paint.details.line_end(0).is_some() || paint.details.line_end(1).is_some() {
                return Err(unsupported("XLS shape line ends are not projected"));
            }
            Some(rgb(
                paint.line.unwrap_or(0),
                paint.line_alpha.unwrap_or(65_536),
            )?)
        } else {
            None
        };
        let text = match source.text {
            Some(offset) => text(records, styles, offset, anchor_text, margins, wrap, &table)?,
            None => None,
        };
        if freeform && inscribed && text.is_some() {
            return Err(unsupported(
                "XLS freeform text rectangles are not projected",
            ));
        }
        let geom = if freeform {
            let decoded = geometry
                .decode(&mut 1_000_000usize)?
                .ok_or_else(|| unsupported("XLS freeform uses guide formulas or path escapes"))?;
            // The shape model carries one fill and one line for all paths.
            // Paths flagged without fill or line are representable only
            // when the shape itself has none, and whether Excel fills an
            // open path is not established, so a filled shape needs every
            // path closed and filled.
            let (path_fill, path_stroke) = decoded
                .uniform_paint()
                .ok_or_else(|| unsupported("XLS freeform paths with differing paint"))?;
            let authored_fill = decoded.uniform_authored_paint().map(|paint| paint.0);
            if fill.is_some() && !(path_fill && authored_fill == Some(true)) {
                return Err(unsupported(
                    "XLS filled freeforms with open or unfilled paths are not projected",
                ));
            }
            if line.is_some() && !path_stroke {
                return Err(unsupported(
                    "XLS freeform paths without their line are not projected",
                ));
            }
            custom_geometry(&decoded)
        } else {
            xlsx_model::ShapeGeom::Preset {
                name: "rect".into(),
                adj: Vec::new(),
            }
        };
        if context_direction
            && text.as_ref().is_some_and(|text| {
                text.paragraphs.iter().flat_map(|p| &p.runs).any(|run| {
                    matches!(run, xlsx_model::ShapeTextRun::Text { text, .. }
                        if text.chars().any(right_to_left))
                })
            })
        {
            return Err(unsupported("right-to-left XLS shape text is not projected"));
        }
        let (join, miter) = paint.details.join();
        Ok(Some((
            xlsx_model::ShapeInfo {
                z_order: self.order,
                // Placement is resolved with the anchor size.
                x: 0.0,
                y: 0.0,
                w: 1.0,
                h: 1.0,
                rot: 0.0,
                flip_h: false,
                flip_v: false,
                fill_color: fill.clone(),
                fill: fill.map(|color| xlsx_model::ShapeFill::Solid { color }),
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
                stroke_color: line,
                stroke_fill: None,
                stroke_custom_dash: Vec::new(),
                stroke_alignment: None,
                stroke_cmpd: None,
                stroke_head_end: None,
                stroke_tail_end: None,
                geom,
                text,
            },
            rotation,
        )))
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
        },
        text: None,
    }
}

fn right_to_left(c: char) -> bool {
    matches!(u32::from(c), 0x0590..=0x08ff | 0xfb1d..=0xfdff | 0xfe70..=0xfeff)
}

/// OfficeArtCOLORREF (MS-ODRAW 2.2.2) with a 16.16 opacity (fillOpacity
/// 2.3.7.5, lineOpacity 2.3.8.2) as the model's `#RRGGBB`, or `#RRGGBBAA`
/// when translucent, with the alpha byte and opaque threshold of the shared
/// DrawingML color parser. Excel's XLSX of a 50% line (`a:alpha 50000`)
/// saves lineOpacity 0x8080, i.e. the same alpha byte 0x80. Palette, scheme
/// and system indexes need Excel's color tables and are rejected.
fn rgb(color: u32, alpha: u32) -> Result<String, String> {
    if !matches!(color & 0xff00_0000, 0 | 0x0400_0000) {
        return Err(unsupported(
            "XLS drawing palette or scheme colors are not projected",
        ));
    }
    let mut hex = format!(
        "#{:02X}{:02X}{:02X}",
        color & 0xff,
        (color >> 8) & 0xff,
        (color >> 16) & 0xff
    );
    let alpha = f64::from(alpha) / 65_536.0;
    if (alpha - 1.0).abs() >= 0.004 {
        hex.push_str(&format!(
            "{:02X}",
            (alpha.clamp(0.0, 1.0) * 255.0).round() as u8
        ));
    }
    Ok(hex)
}

/// TxO (MS-XLS 2.4.329): the text string in Continue records of
/// XLUnicodeStringNoCch fragments, then its TxORuns formatting runs.
fn text(
    records: &[Record<'_>],
    styles: &styles::Styles<'_>,
    offset: usize,
    anchor_text: Option<u32>,
    margins: [u32; 4],
    wrap: &str,
    table: &Table,
) -> Result<Option<xlsx_model::ShapeText>, String> {
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
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
            ),
            _ => return Err(unsupported("invalid BIFF TxO text fragment")),
        }
        if units.len() > characters {
            return Err(unsupported("BIFF TxO text longer than cchText"));
        }
        next += 1;
    }
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
    for entry in runs_data.chunks_exact(8) {
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
    };
    let mut paragraphs = Vec::new();
    let mut start = 0usize;
    loop {
        let end = units[start..]
            .iter()
            .position(|&unit| unit == 0x000a)
            .map_or(characters, |at| start + at);
        let mut paragraph_runs = Vec::new();
        if start == end {
            // An empty line keeps the height of the font at its position.
            paragraph_runs.push(make_run(String::new(), font_at(start)));
        }
        let mut at = start;
        while at < end {
            let run = runs.partition_point(|run| run.0 <= at) - 1;
            let run_end = runs.get(run + 1).map_or(characters, |next| next.0).min(end);
            let text = String::from_utf16(&units[at..run_end])
                .map_err(|_| unsupported("invalid UTF-16 in BIFF TxO text"))?;
            paragraph_runs.push(make_run(text, &fonts[&runs[run].1]));
            at = run_end;
        }
        paragraphs.push(xlsx_model::ShapeParagraph {
            align: align.into(),
            rtl: false,
            mar_l: None,
            mar_r: None,
            indent: None,
            space_line: None,
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
    if font.underline || font.strike || font.other_effects {
        return Err(unsupported(
            "XLS shape text underline, strikethrough or font effects are not projected",
        ));
    }
    if !matches!(font.weight, 400 | 700) || font.size_twips == 0 || font.name.is_empty() {
        return Err(unsupported("XLS shape text font weight is not projected"));
    }
    let color = match (font.color, font.automatic_color) {
        (Some(color), _) => color,
        (None, true) => {
            return Err(unsupported(
                "automatic XLS shape text color is not projected",
            ))
        }
        (None, false) => {
            return Err(unsupported(
                "XLS shape text system colors are not projected",
            ))
        }
    };
    Ok(RunFont {
        name: font.name,
        size: f64::from(font.size_twips) / 20.0,
        bold: font.weight == 700,
        italic: font.italic,
        color,
    })
}

#[cfg(test)]
mod tests;
