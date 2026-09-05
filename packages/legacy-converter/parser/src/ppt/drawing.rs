//! Positioned text and basic preset reconstruction of the live slide's OfficeArt tree.
//! [MS-PPT] 2.5.13, 2.7.1, 2.9.76 and [MS-ODRAW] 2.2.14/16/38/39/40.
//! Output uses ECMA-376 DrawingML CT_(Group)Transform2D, never binary-aware paint.
use super::*;
use crate::officeart::geometry;

#[derive(Clone, Copy)]
pub(super) struct TextContext<'a> {
    pub fonts: &'a [String],
    pub styles: &'a [Option<&'a [u8]>],
    pub scheme: Option<&'a scheme::Scheme>,
    pub types: &'a [u16],
    pub master: Option<&'a text_style::Master>,
    pub shapes: Option<&'a shape_master::Resolver<'a>>,
    pub outline_slide_numbers: &'a [Vec<u32>],
    pub slide_number: u32,
}

#[cfg(test)]
pub(super) fn render<'a>(
    slide: &[u8],
    outline: &'a [String],
    records: &mut usize,
    text: &mut usize,
    xml: &mut usize,
    context: Option<TextContext<'a>>,
    media: Option<&mut media::Store<'a>>,
) -> Result<Option<String>, String> {
    let result = render_with_masters(slide, &[], outline, records, text, xml, context, media)?;
    Ok((!result.fallback).then_some(result.tree))
}

pub(super) struct Rendered {
    pub tree: String,
    pub fallback: bool,
}

#[allow(clippy::too_many_arguments)]
pub(super) fn render_with_masters<'a>(
    slide: &[u8],
    masters: &[Record<'_>],
    outline: &'a [String],
    records: &mut usize,
    text: &mut usize,
    xml: &mut usize,
    context: Option<TextContext<'a>>,
    media: Option<&mut media::Store<'a>>,
) -> Result<Rendered, String> {
    let mut writer = Writer {
        outline,
        records,
        text,
        remaining: xml,
        output: String::new(),
        id: 1,
        context,
        media,
        inherited: true,
    };
    // Flatten passive master objects below slide-local objects. One writer owns
    // the IDs and budgets across all layers (ECMA-376 19.3.1.45 lexical z-order).
    for master in masters {
        writer.drawing(master.payload)?;
    }
    writer.inherited = false;
    let local = writer.drawing(slide)?;
    if !local {
        // Preserve the prior missing-drawing fallback even if a master exists.
        let mut blocks = Vec::new();
        collect_text(slide, 0, writer.records, &mut blocks, outline, writer.text)?;
        let id = writer.next_id()?;
        let fallback = fallback_text(&blocks, id, writer.remaining)?;
        writer.output.push_str(&fallback);
    }
    Ok(Rendered {
        tree: writer.output,
        fallback: !local,
    })
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct Rect {
    x: i64,
    y: i64,
    w: i64,
    h: i64,
}

impl Rect {
    fn read(record: Record<'_>) -> Result<Self, String> {
        let b = record.payload;
        let values: Vec<i64> = match (record.kind, record.version, b.len()) {
            (0xf010, 0, 8) => b
                .chunks_exact(2)
                .map(|v| i16::from_le_bytes([v[0], v[1]]) as i64)
                .collect(),
            (0xf010 | 0xf00f, 0, 16) | (0xf009, 1, 16) => b
                .chunks_exact(4)
                .map(|v| i32::from_le_bytes(v.try_into().expect("four bytes")) as i64)
                .collect(),
            _ => return Err(unsupported("invalid PowerPoint shape anchor")),
        };
        // PPT RectStruct is top/left/right/bottom; OfficeArt child/group bounds
        // are left/top/right/bottom. Convert all coordinate spaces consistently.
        let (left, top) = if record.kind == 0xf010 {
            (values[1], values[0])
        } else {
            (values[0], values[1])
        };
        if values[2] < left || values[3] < top {
            return Err(unsupported("inverted PowerPoint shape anchor"));
        }
        // One master unit = 1/576 inch = 1587.5 EMU. Signed rounding to nearest
        // keeps negative off-slide positions instead of clamping them to zero.
        let emu = master_to_emu;
        Ok(Self {
            x: emu(left),
            y: emu(top),
            w: emu(values[2]) - emu(left),
            h: emu(values[3]) - emu(top),
        })
    }
}

struct Shape<'a> {
    id: u32,
    kind: u16,
    flags: u32,
    anchor: Option<Rect>,
    child_space: Option<Rect>,
    textbox: Option<Record<'a>>,
    placeholder: bool,
    props: Properties<'a>,
}

impl<'a> Shape<'a> {
    fn read(record: Record<'a>, nested: bool, budget: &mut usize) -> Result<Self, String> {
        if record.kind != 0xf004 || record.version != 15 {
            return Err(unsupported("invalid PowerPoint shape container"));
        }
        let mut flags = None;
        let mut id = 0;
        let mut kind = 0;
        let mut anchor = None;
        let mut child_space = None;
        let mut textbox = None;
        let mut placeholder = None;
        let mut props = Properties::default();
        for child in parse_records(record.payload, budget)? {
            match child.kind {
                0xf00a => {
                    if flags.is_some() || child.version != 2 || child.payload.len() != 8 {
                        return Err(unsupported("invalid PowerPoint shape flags"));
                    }
                    flags = Some(u32_at(child.payload, 4)?);
                    id = u32_at(child.payload, 0)?;
                    kind = child.instance;
                }
                0xf010 | 0xf00f => {
                    if anchor.is_some() || (child.kind == 0xf00f) != nested {
                        return Err(unsupported("ambiguous PowerPoint shape coordinate space"));
                    }
                    anchor = Some(Rect::read(child)?);
                }
                0xf009 => {
                    if child_space.is_some() {
                        return Err(unsupported("duplicate PowerPoint group bounds"));
                    }
                    child_space = Some(Rect::read(child)?);
                }
                0xf00d => {
                    if textbox.is_some() || child.version != 15 {
                        return Err(unsupported("invalid PowerPoint shape text container"));
                    }
                    textbox = Some(child);
                }
                0xf00b => props.read(child, budget)?,
                0xf011 => {
                    if child.version != 15 {
                        return Err(unsupported("invalid PowerPoint client data"));
                    }
                    // Inspect only direct placeholder metadata. Never descend
                    // into action/link containers or collect their text.
                    for atom in parse_records(child.payload, budget)? {
                        if atom.kind != 3011 {
                            continue;
                        }
                        if placeholder.is_some() || atom.version != 0 || atom.payload.len() != 8 {
                            return Err(unsupported("invalid PowerPoint placeholder metadata"));
                        }
                        placeholder = Some(u32_at(atom.payload, 0)? != u32::MAX);
                    }
                }
                _ => {}
            }
        }
        Ok(Self {
            id,
            kind,
            flags: flags.ok_or_else(|| unsupported("missing PowerPoint shape flags"))?,
            anchor,
            child_space,
            textbox,
            placeholder: placeholder.unwrap_or(false),
            props,
        })
    }
    fn omitted(&self) -> bool {
        self.flags & (8 | 16 | 1024) != 0 || self.props.script
    }
    fn master(&self) -> Option<u32> {
        (self.flags & 0x20 != 0).then_some(self.props.master.unwrap_or(0))
    }
    fn transform(&self, anchor: Rect, group: Option<Rect>) -> String {
        let mut xml = format!("<a:xfrm rot=\"{}\" flipH=\"{}\" flipV=\"{}\"><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/>", self.props.rotation, (self.flags >> 6) & 1, (self.flags >> 7) & 1, anchor.x, anchor.y, anchor.w, anchor.h);
        if let Some(ch) = group {
            xml.push_str(&format!(
                "<a:chOff x=\"{}\" y=\"{}\"/><a:chExt cx=\"{}\" cy=\"{}\"/>",
                ch.x, ch.y, ch.w, ch.h
            ));
        }
        xml.push_str("</a:xfrm>");
        xml
    }
}

struct Properties<'a> {
    geometry: geometry::Geometry<'a>,
    hidden: bool,
    script: bool,
    master: Option<u32>,
    picture: u32,
    crop: [i64; 4],
    paint: paint::Paint,
    rotation: i64,
    margins: [u32; 4],
    wrap: &'static str,
    anchor: &'static str,
    center: bool,
}
impl Default for Properties<'_> {
    fn default() -> Self {
        Self {
            geometry: geometry::Geometry::default(),
            hidden: false,
            script: false,
            master: None,
            picture: 0,
            crop: [0; 4],
            paint: paint::Paint::default(),
            rotation: 0,
            margins: [91440, 45720, 91440, 45720],
            wrap: "square",
            anchor: "t",
            center: false,
        }
    }
}
impl<'a> Properties<'a> {
    fn read(&mut self, record: Record<'a>, budget: &mut usize) -> Result<(), String> {
        // [MS-ODRAW] 2.2.7–9: six-byte property entries precede complex data.
        // Even ignored complex properties must fit; never use their payload as XML.
        if record.version != 3 {
            return Err(unsupported("invalid PowerPoint shape property table"));
        }
        *budget = budget
            .checked_sub(usize::from(record.instance))
            .ok_or_else(|| unsupported("PowerPoint shape property work budget exceeded"))?;
        let len = usize::from(record.instance) * 6;
        let entries = record
            .payload
            .get(..len)
            .ok_or_else(|| unsupported("truncated PowerPoint shape properties"))?;
        let mut end = len;
        for entry in entries.chunks_exact(6) {
            let opid = u16_at(entry, 0)?;
            let value = u32_at(entry, 2)?;
            if opid & 0x8000 != 0 {
                if matches!(opid & 0x3fff, 0x145..=0x150) {
                    self.paint.custom_geometry = true;
                }
                let start = end;
                end = end
                    .checked_add(value as usize)
                    .filter(|n| *n <= record.payload.len())
                    .ok_or_else(|| unsupported("truncated PowerPoint complex shape property"))?;
                self.geometry
                    .complex(opid & 0x3fff, &record.payload[start..end]);
                continue;
            }
            if matches!(opid & 0x3fff, 0x145 | 0x146) {
                self.geometry.scalar(opid & 0x3fff, value)?;
            }
            if opid & 0x4000 != 0 {
                if opid == 0x4104 {
                    self.picture = value;
                } else if opid == 0x4186 {
                    self.paint.property(opid, value)?;
                }
                continue;
            }
            match opid {
                // MS-ODRAW 2.3.4.44. Hidden shapes and active script anchors
                // are omitted before any text/image references are followed.
                0x3bf => {
                    for (bit, target) in [(1, &mut self.hidden), (7, &mut self.script)] {
                        if value & (1 << (bit + 16)) != 0 {
                            *target = value & (1 << bit) != 0;
                        }
                    }
                }
                // MS-ODRAW hspMaster is a scalar MSOSPID, not a BLIP index.
                0x301 => self.master = Some(value),
                // MS-ODRAW crop order: top, bottom, left, right. Signed 16.16
                // fractions become DrawingML 1/1000 percentages without clamping.
                0x100..=0x103 => {
                    let value = i64::from(value as i32) * 100000;
                    let value = (value + value.signum() * 32768) / 65536;
                    i32::try_from(value).map_err(|_| {
                        unsupported("PowerPoint crop exceeds DrawingML percentage range")
                    })?;
                    self.crop[usize::from(opid - 0x100)] = value;
                }
                // [MS-ODRAW] 2.3.18.5: signed 16.16 degrees -> 1/60000 degree.
                4 => self.rotation = i64::from(value as i32) * 60000 / 65536,
                0x81..=0x84 => {
                    if value > 0x132f540 {
                        return Err(unsupported("invalid PowerPoint text margin"));
                    }
                    self.margins[usize::from(opid - 0x81)] = value;
                }
                0x85 => self.wrap = if value == 2 { "none" } else { "square" },
                0x87 if value <= 5 => {
                    self.anchor = ["t", "ctr", "b"][(value % 3) as usize];
                    self.center = value >= 3;
                }
                _ => {
                    self.geometry.scalar(opid, value)?;
                    self.paint.property(opid, value)?;
                }
            }
        }
        if end != record.payload.len() {
            return Err(unsupported("unexpected PowerPoint property data"));
        }
        Ok(())
    }
}

/// A slide background is the ungrouped OfficeArt background shape, not an
/// arbitrary full-slide rectangle. Never inspect nested client/action data.
pub(super) fn background(slide: &[u8], budget: &mut usize) -> Result<Option<paint::Paint>, String> {
    let children = parse_records(slide, budget)?;
    let mut drawings = children.iter().filter(|r| r.kind == 1036);
    let Some(drawing) = drawings.next() else {
        return Ok(None);
    };
    if drawings.next().is_some() || drawing.version != 15 {
        return Err(unsupported("invalid PowerPoint background drawing"));
    }
    let groups = parse_records(drawing.payload, budget)?;
    if groups.len() != 1 || groups[0].kind != 0xf002 || groups[0].version != 15 {
        return Err(unsupported(
            "invalid PowerPoint background OfficeArt drawing",
        ));
    }
    let mut result = None;
    for record in parse_records(groups[0].payload, budget)? {
        if record.kind != 0xf004 {
            continue;
        }
        // Only the flag record is needed to decide whether this is a background.
        let flags: Vec<_> = parse_records(record.payload, budget)?
            .into_iter()
            .filter(|r| r.kind == 0xf00a)
            .collect();
        if flags.len() != 1 || flags[0].version != 2 || flags[0].payload.len() != 8 {
            return Err(unsupported("invalid PowerPoint background shape flags"));
        }
        let value = u32_at(flags[0].payload, 4)?;
        if value & 1024 == 0 || value & (8 | 16) != 0 {
            continue;
        }
        if result.is_some() {
            return Err(unsupported("duplicate PowerPoint background shapes"));
        }
        result = Some(Shape::read(record, false, budget)?.props.paint);
    }
    Ok(result)
}

pub(super) fn master_shapes<'a>(
    slide: &'a [u8],
    base: Option<std::rc::Rc<text_style::Master>>,
    output: &mut shape_master::Resolver<'a>,
    budget: &mut usize,
    text_budget: &mut usize,
) -> Result<(), String> {
    fn visit<'a>(
        record: Record<'a>,
        nested: bool,
        depth: usize,
        base: &Option<std::rc::Rc<text_style::Master>>,
        output: &mut shape_master::Resolver<'a>,
        budget: &mut usize,
        text_budget: &mut usize,
    ) -> Result<(), String> {
        if depth >= MAX_DEPTH {
            return Err(unsupported("PowerPoint master drawing depth exceeded"));
        }
        if record.kind == 0xf003 {
            if record.version != 15 {
                return Err(unsupported("invalid PowerPoint master group"));
            }
            let children = parse_records(record.payload, budget)?;
            let first = children
                .first()
                .ok_or_else(|| unsupported("empty PowerPoint master group"))?;
            let group = Shape::read(*first, nested, budget)?;
            if group.flags & 1 == 0 || (nested && group.flags & 4 != 0) {
                return Err(unsupported("invalid PowerPoint master group flags"));
            }
            if group.omitted() {
                return Ok(());
            }
            let child_nested = group.flags & 4 == 0;
            visit(*first, nested, depth + 1, base, output, budget, text_budget)?;
            for child in &children[1..] {
                visit(
                    *child,
                    child_nested,
                    depth + 1,
                    base,
                    output,
                    budget,
                    text_budget,
                )?;
            }
        } else if record.kind == 0xf004 {
            let shape = Shape::read(record, nested, budget)?;
            if shape.omitted() {
                return Ok(());
            }
            let (mut kind, mut text, mut style) = (None, None, None);
            if let Some(textbox) = shape.textbox {
                for atom in parse_records(textbox.payload, budget)? {
                    match atom.kind {
                        3999 => {
                            if kind.is_some() {
                                return Err(unsupported("duplicate master text header"));
                            }
                            kind = Some(text_style::text_type(atom)?);
                        }
                        TEXT_CHARS_ATOM | TEXT_BYTES_ATOM => {
                            if text.is_some() {
                                return Err(unsupported("duplicate master text body"));
                            }
                            let decoded = decode_text(atom)?;
                            charge_text(text_budget, decoded.len())?;
                            text = Some(decoded);
                        }
                        4001 => {
                            if style.is_some() || atom.version != 0 {
                                return Err(unsupported("invalid master text style"));
                            }
                            style = Some(atom.payload);
                        }
                        _ => {} // Actions, links and metacharacter evaluation remain absent.
                    }
                }
            }
            let direct = match (text.as_deref(), style) {
                (Some(text), Some(style)) => text_style::shape_levels(text, style, budget)?,
                _ => Vec::new(),
            };
            output.insert(shape_master::Node {
                id: shape.id,
                parent: shape.master(),
                text_type: kind,
                direct,
                base: base.clone(),
                paint: shape.props.paint,
                geometry: shape.props.geometry,
            })?;
        }
        Ok(())
    }
    for drawing in parse_records(slide, budget)?
        .into_iter()
        .filter(|r| r.kind == 1036)
    {
        for dg in parse_records(drawing.payload, budget)? {
            if dg.kind != 0xf002 || dg.version != 15 {
                return Err(unsupported("invalid master OfficeArt drawing"));
            }
            for child in parse_records(dg.payload, budget)? {
                visit(child, false, 0, &base, output, budget, text_budget)?;
            }
        }
    }
    Ok(())
}

struct Writer<'a, 'b> {
    inherited: bool,
    media: Option<&'b mut media::Store<'a>>,
    outline: &'a [String],
    records: &'b mut usize,
    text: &'b mut usize,
    remaining: &'b mut usize,
    output: String,
    id: u32,
    context: Option<TextContext<'a>>,
}
impl Writer<'_, '_> {
    fn drawing(&mut self, slide: &[u8]) -> Result<bool, String> {
        let children = parse_records(slide, self.records)?;
        let mut drawings = children.iter().filter(|r| r.kind == 1036);
        let Some(drawing) = drawings.next() else {
            return Ok(false);
        };
        if drawings.next().is_some() || drawing.version != 15 {
            return Err(unsupported("invalid PowerPoint drawing container"));
        }
        for dg in parse_records(drawing.payload, self.records)? {
            if dg.kind != 0xf002 || dg.version != 15 {
                return Err(unsupported("invalid PowerPoint OfficeArt drawing"));
            }
            for child in parse_records(dg.payload, self.records)? {
                self.node(child, false, 0)?;
            }
        }
        Ok(true)
    }
    fn push(&mut self, value: &str) -> Result<(), String> {
        append(&mut self.output, self.remaining, value)
    }
    fn next_id(&mut self) -> Result<u32, String> {
        // Independent resource policy; IDs are output-local, never source actions.
        if self.id >= 100_001 {
            return Err(unsupported("too many PowerPoint drawing shapes"));
        }
        self.id += 1;
        Ok(self.id)
    }
    fn node(&mut self, record: Record<'_>, nested: bool, depth: usize) -> Result<(), String> {
        if depth > MAX_DEPTH {
            return Err(unsupported("PowerPoint drawing nesting is too deep"));
        }
        match record.kind {
            0xf003 => {
                if record.version != 15 {
                    return Err(unsupported("invalid PowerPoint group container"));
                }
                let children = parse_records(record.payload, self.records)?;
                let first = children
                    .first()
                    .ok_or_else(|| unsupported("empty PowerPoint group"))?;
                let group = Shape::read(*first, nested, self.records)?;
                if group.omitted() || group.props.hidden || (self.inherited && group.placeholder) {
                    return Ok(());
                }
                if group.flags & 1 == 0 {
                    return Err(unsupported("missing PowerPoint group flag"));
                }
                let patriarch = group.flags & 4 != 0;
                if patriarch && nested {
                    return Err(unsupported("nested PowerPoint patriarch group"));
                }
                if !patriarch {
                    let anchor = group
                        .anchor
                        .ok_or_else(|| unsupported("missing PowerPoint group anchor"))?;
                    let child_space = group
                        .child_space
                        .filter(|r| r.w > 0 && r.h > 0)
                        .ok_or_else(|| unsupported("invalid PowerPoint group coordinate space"))?;
                    let id = self.next_id()?;
                    self.push(&format!("<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"{id}\" name=\"Legacy group {id}\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>{}</p:grpSpPr>", group.transform(anchor, Some(child_space))))?;
                }
                for child in &children[1..] {
                    self.node(*child, !patriarch, depth + 1)?;
                }
                if !patriarch {
                    self.push("</p:grpSp>")?;
                }
            }
            0xf004 => {
                let shape = Shape::read(record, nested, self.records)?;
                if shape.omitted() || shape.props.hidden || (self.inherited && shape.placeholder) {
                    return Ok(());
                }
                let paint = match (shape.master(), self.context.and_then(|c| c.shapes)) {
                    (Some(id), Some(shapes)) => shape.props.paint.inherit(shapes.paint(id)?),
                    _ => shape.props.paint,
                };
                if shape.kind == 75 && shape.props.picture != 0 {
                    let index = shape.props.picture;
                    if self
                        .media
                        .as_deref_mut()
                        .map(|m| m.reference(index, self.records))
                        .transpose()?
                        .unwrap_or(false)
                    {
                        let anchor = shape
                            .anchor
                            .ok_or_else(|| unsupported("missing PowerPoint picture anchor"))?;
                        let id = self.next_id()?;
                        let [top, bottom, left, right] = shape.props.crop;
                        if left + right >= 100000 || top + bottom >= 100000 {
                            return Err(unsupported("empty PowerPoint picture crop"));
                        }
                        self.push(&format!("<p:pic><p:nvPicPr><p:cNvPr id=\"{id}\" name=\"Legacy picture {id}\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rImg{index}\"/><a:srcRect l=\"{left}\" t=\"{top}\" r=\"{right}\" b=\"{bottom}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{}</p:spPr></p:pic>", shape.transform(anchor, None), paint.xml_with_scheme(1, self.context.and_then(|c| c.scheme))))?;
                    }
                }
                let mut text = Vec::new();
                let mut style = None;
                let mut body_seen = false;
                let mut outline_body = false;
                let mut text_type = None;
                let mut slide_numbers = Vec::new();
                // The only source of visible text is this shape's ClientTextbox.
                let atoms = shape
                    .textbox
                    .map(|t| parse_records(t.payload, self.records))
                    .transpose()?
                    .unwrap_or_default();
                for atom in atoms {
                    match atom.kind {
                        3999 => {
                            if text_type.is_some() || body_seen {
                                return Err(unsupported("ambiguous PowerPoint text header"));
                            }
                            text_type = Some(text_style::text_type(atom)?);
                        }
                        TEXT_CHARS_ATOM | TEXT_BYTES_ATOM => {
                            if body_seen {
                                return Err(unsupported("duplicate PowerPoint text body"));
                            }
                            body_seen = true;
                            let value = decode_text(atom)?;
                            charge_text(self.text, value.len())?;
                            push_text(&mut text, value)?;
                        }
                        3998 => {
                            // Master placeholder exemplars must never acquire
                            // slide-local outline text. Master non-placeholder
                            // text should be stored directly in ClientTextbox.
                            if self.inherited {
                                return Err(unsupported(
                                    "outline reference in inherited master object",
                                ));
                            }
                            if body_seen || style.is_some() {
                                return Err(unsupported("ambiguous PowerPoint outline text body"));
                            }
                            body_seen = true;
                            outline_body = true;
                            let index = u32_at(atom.payload, 0)? as usize;
                            if text_type.is_some() {
                                return Err(unsupported(
                                    "inline header on PowerPoint outline reference",
                                ));
                            }
                            text_type = self.context.and_then(|c| c.types.get(index).copied());
                            style = self
                                .context
                                .and_then(|context| context.styles.get(index).copied().flatten());
                            if let Some(positions) = self
                                .context
                                .and_then(|c| c.outline_slide_numbers.get(index))
                            {
                                *self.records =
                                    self.records.checked_sub(positions.len()).ok_or_else(|| {
                                        unsupported("PowerPoint slide-number work budget exceeded")
                                    })?;
                                slide_numbers.extend_from_slice(positions);
                            }
                            let value = self.outline.get(index).ok_or_else(|| {
                                unsupported("PowerPoint outline text index out of range")
                            })?;
                            charge_text(self.text, value.len())?;
                            push_text(&mut text, value.clone())?;
                        }
                        4001 => {
                            if atom.version != 0 || style.is_some() || outline_body {
                                return Err(unsupported("invalid PowerPoint text style record"));
                            }
                            style = Some(atom.payload);
                        }
                        4056 => {
                            if !body_seen || outline_body {
                                return Err(unsupported("orphan PowerPoint slide-number atom"));
                            }
                            slide_numbers.push(text_style::slide_number_position(atom)?);
                        }
                        // Do not descend into interactive/action containers.
                        _ => {}
                    }
                }
                if !slide_numbers.is_empty() && text.len() != 1 {
                    return Err(unsupported("ambiguous PowerPoint slide-number text body"));
                }
                let preset = paint.geometry(shape.kind);
                // Picture clipping is a distinct path; this restores foreground
                // vector shapes, including master objects, without painting an
                // extra vector shape over a picture frame.
                let geometry = match (shape.master(), self.context.and_then(|c| c.shapes)) {
                    (Some(id), Some(shapes)) => shape.props.geometry.inherit(shapes.geometry(id)?),
                    _ => shape.props.geometry,
                };
                let custom = if shape.kind == 75 {
                    None
                } else {
                    // The existing PPTX model has one paint per shape. Keep
                    // mixed per-path paint unsupported rather than discard its
                    // flags or introduce legacy-specific rendering behavior.
                    geometry
                        .decode(self.records)?
                        .filter(|g| g.uniform_paint().is_some())
                };
                if text.is_empty() && preset.is_none() && custom.is_none() {
                    return Ok(());
                }
                let anchor = shape
                    .anchor
                    .ok_or_else(|| unsupported("missing PowerPoint shape anchor"))?;
                let id = self.next_id()?;
                // Unsupported geometry may still carry text. Preserve its text
                // frame, but never paint an invented rectangle in its place.
                let text_box =
                    u8::from(shape.kind == 202 || (preset.is_none() && custom.is_none()));
                self.push(&format!("<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Legacy shape {id}\"/><p:cNvSpPr txBox=\"{text_box}\"/><p:nvPr/></p:nvSpPr><p:spPr>{}", shape.transform(anchor, None)))?;
                let scheme = self.context.and_then(|c| c.scheme);
                if let Some(custom) = custom {
                    let (fill, stroke) = custom
                        .uniform_paint()
                        .expect("uniform paths filtered above");
                    custom.write_xml(&mut self.output, self.remaining)?;
                    self.push(&paint.xml_with_custom_geometry(scheme, fill, stroke))?;
                } else {
                    self.push(&format!(
                        "<a:prstGeom prst=\"{}\"><a:avLst/></a:prstGeom>{}",
                        preset.unwrap_or("rect"),
                        paint.xml_with_scheme(shape.kind, scheme)
                    ))?;
                }
                self.push("</p:spPr>")?;
                if text.is_empty() {
                    self.push("</p:sp>")?;
                    return Ok(());
                }
                self.push("<p:txBody>")?;
                let p = &shape.props;
                self.push(&format!("<a:bodyPr wrap=\"{}\" anchor=\"{}\" anchorCtr=\"{}\" lIns=\"{}\" tIns=\"{}\" rIns=\"{}\" bIns=\"{}\"/><a:lstStyle/>", p.wrap, p.anchor, u8::from(p.center), p.margins[0], p.margins[1], p.margins[2], p.margins[3]))?;
                let linked = match (shape.master(), self.context.and_then(|c| c.shapes)) {
                    (Some(id), Some(shapes)) => Some(shapes.levels(id)?),
                    _ => None,
                };
                let levels = if linked.is_some() {
                    linked
                } else if shape.placeholder {
                    self.context
                        .and_then(|c| c.master)
                        .and_then(|m| text_type.and_then(|t| m.levels(t)))
                } else {
                    None
                };
                slide_numbers.sort_unstable();
                let default_style = if style.is_none()
                    && (levels.is_some() || !slide_numbers.is_empty())
                    && text.len() == 1
                {
                    Some(text_style::default_style(&text[0]))
                } else {
                    None
                };
                if let Some(style) = style.or(default_style.as_deref()) {
                    if text.len() != 1 {
                        return Err(unsupported("ambiguous PowerPoint styled text body"));
                    }
                    text_style::write(
                        &text[0],
                        style,
                        text_style::Context {
                            fonts: self.context.map_or(&[], |c| c.fonts),
                            scheme: self.context.and_then(|c| c.scheme),
                            levels,
                            slide_numbers: &slide_numbers,
                            slide_number: self.context.map_or(0, |c| c.slide_number),
                        },
                        &mut self.output,
                        self.remaining,
                        self.records,
                    )?;
                } else {
                    paragraphs(&text, &mut self.output, self.remaining)?;
                }
                self.push("</p:txBody></p:sp>")?;
            }
            _ => {}
        }
        Ok(())
    }
}

pub(super) fn append(output: &mut String, budget: &mut usize, value: &str) -> Result<(), String> {
    *budget = budget
        .checked_sub(value.len())
        .ok_or_else(|| "OUTPUT_TOO_LARGE".to_string())?;
    output.push_str(value);
    Ok(())
}

pub(super) fn paragraphs(
    blocks: &[String],
    output: &mut String,
    budget: &mut usize,
) -> Result<(), String> {
    for block in blocks {
        for line in block.split('\r') {
            append(output, budget, "<a:p>")?;
            // Bounded chunks prevent XML escaping of one large atom from allocating
            // a multiple of the entire input before the XML budget is checked.
            text_style::write_run(line, "<a:rPr sz=\"1800\"/>", output, budget)?;
            append(output, budget, "<a:endParaRPr sz=\"1800\"/></a:p>")?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::persist::tests::record;
    use super::*;

    fn ints(values: &[i32]) -> Vec<u8> {
        values.iter().flat_map(|n| n.to_le_bytes()).collect()
    }
    fn sp(flags: i32, children: Vec<Vec<u8>>) -> Vec<u8> {
        record(
            15,
            0xf004,
            &[
                vec![record((202 << 4) | 2, 0xf00a, &ints(&[42, flags]))],
                children,
            ]
            .concat()
            .concat(),
        )
    }
    fn text(value: &str) -> Vec<u8> {
        record(15, 0xf00d, &record(0, 4008, value.as_bytes()))
    }
    fn drawing(shapes: Vec<Vec<u8>>) -> Vec<u8> {
        record(15, 1036, &record(15, 0xf002, &shapes.concat()))
    }
    fn xml(bytes: &[u8]) -> Result<String, String> {
        render(
            bytes,
            &["Outline".into()],
            &mut MAX_RECORDS.clone(),
            &mut MAX_TEXT_BYTES.clone(),
            &mut (256 * 1024 * 1024),
            None,
            None,
        )
        .map(|x| x.unwrap_or_default())
    }

    fn properties(values: &[(u16, u32)]) -> Vec<u8> {
        let payload: Vec<u8> = values
            .iter()
            .flat_map(|(id, value)| {
                [id.to_le_bytes().to_vec(), value.to_le_bytes().to_vec()].concat()
            })
            .collect();
        record(((values.len() as u16) << 4) | 3, 0xf00b, &payload)
    }

    #[test]
    fn master_layers_and_missing_local_drawing_share_ids_and_xml_budget() {
        let master = drawing(vec![sp(
            0xa00,
            vec![record(0, 0xf010, &ints(&[0, 0, 576, 576])), text("Master")],
        )]);
        let layer = Record {
            version: 15,
            instance: 0,
            kind: 1016,
            payload: &master,
        };
        let local = record(0, 4008, b"Local fallback");
        let rendered = render_with_masters(
            &local,
            &[layer],
            &[],
            &mut 1000,
            &mut 1000,
            &mut 10000,
            None,
            None,
        )
        .unwrap();
        assert!(rendered.fallback);
        assert!(rendered.tree.contains("id=\"2\" name=\"Legacy shape 2\""));
        assert!(rendered
            .tree
            .contains("id=\"3\" name=\"Legacy slide text\""));
        assert!(
            rendered.tree.find(">Master<").unwrap()
                < rendered.tree.find(">Local fallback<").unwrap()
        );
        assert!(render_with_masters(
            &local,
            &[layer],
            &[],
            &mut 1000,
            &mut 1000,
            &mut 100,
            None,
            None
        )
        .is_err());
    }

    #[test]
    fn hidden_and_script_shapes_do_not_follow_outline_references_or_require_anchors() {
        let bad_outline = record(15, 0xf00d, &record(0, 3998, &ints(&[99])));
        for (value, omitted) in [
            (0x00020002, true),
            (0x00800080, true),
            (2, false),
            (0x00020000, false),
        ] {
            let shape = sp(
                0xa00,
                vec![properties(&[(0x3bf, value)]), bad_outline.clone()],
            );
            let result = xml(&drawing(vec![shape]));
            if omitted {
                assert_eq!(result.unwrap(), "");
            } else {
                assert!(result.is_err());
            }
        }
        let hidden_group = record(
            15,
            0xf003,
            &[
                sp(1, vec![properties(&[(0x3bf, 0x00020002)])]),
                sp(0xa00, vec![bad_outline]),
            ]
            .concat(),
        );
        assert_eq!(xml(&drawing(vec![hidden_group])).unwrap(), "");
    }

    #[test]
    fn backgrounds_are_explicit_ungrouped_live_shapes_without_anchor_requirements() {
        let bg = sp(0xc00, vec![properties(&[(0x181, 0x123456)])]);
        let input = drawing(vec![bg.clone()]);
        assert!(background(&input, &mut 100)
            .unwrap()
            .unwrap()
            .background_fill(None)
            .unwrap()
            .contains("563412"));
        assert!(xml(&input).unwrap().is_empty()); // Never emit a foreground rectangle.
        for flag in [0x800, 0xc08, 0xc10] {
            assert!(background(&drawing(vec![sp(flag, vec![])]), &mut 100)
                .unwrap()
                .is_none());
        }
        assert!(
            background(&drawing(vec![record(15, 0xf003, &bg)]), &mut 100)
                .unwrap()
                .is_none()
        );
        assert!(background(&drawing(vec![bg.clone(), bg]), &mut 100)
            .err()
            .unwrap()
            .contains("duplicate"));
        assert!(background(&input, &mut 1).is_err());
    }

    #[test]
    fn picture_properties_require_blip_reference_bit_and_preserve_signed_crop() {
        let bytes = properties(&[
            (0x104, 9),
            (0x4104, 2),
            (0x100, 16384),
            (0x101, (-8192i32) as u32),
            (0x102, 32768),
            (0x103, 0),
        ]);
        let mut props = Properties::default();
        props
            .read(parse_record_at(&bytes, 0, &mut 100).unwrap(), &mut 100)
            .unwrap();
        assert_eq!(props.picture, 2);
        assert_eq!(props.crop, [25000, -12500, 50000, 0]);
        let bytes = properties(&[(0x104, 7)]);
        let mut props = Properties::default();
        props
            .read(parse_record_at(&bytes, 0, &mut 100).unwrap(), &mut 100)
            .unwrap();
        assert_eq!(props.picture, 0);
        let bytes = properties(&[(0x100, i32::MAX as u32)]);
        assert!(props
            .read(parse_record_at(&bytes, 0, &mut 100).unwrap(), &mut 100)
            .is_err());
    }

    #[test]
    fn master_text_only_applies_to_verified_placeholders_not_ordinary_text_boxes() {
        let master_bytes = [
            1u16.to_le_bytes().to_vec(),
            0x800u32.to_le_bytes().to_vec(),
            1u16.to_le_bytes().to_vec(),
            0x20000u32.to_le_bytes().to_vec(),
            48u16.to_le_bytes().to_vec(),
        ]
        .concat();
        let master = text_style::Master::parse(
            &[Record {
                version: 0,
                instance: 0,
                kind: 4003,
                payload: &master_bytes,
            }],
            &[],
            &mut 100,
        )
        .unwrap();
        let shape = |position: u32, label: &str| {
            sp(
                0xa00,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 1152, 576])),
                    record(
                        15,
                        0xf00d,
                        &[record(0, 3999, &[0; 4]), record(0, 4008, label.as_bytes())].concat(),
                    ),
                    record(
                        15,
                        0xf011,
                        &record(0, 3011, &[position.to_le_bytes(), [1, 0, 0, 0]].concat()),
                    ),
                ],
            )
        };
        let input = drawing(vec![shape(0, "Title"), shape(u32::MAX, "Ordinary")]);
        let out = render(
            &input,
            &[],
            &mut MAX_RECORDS.clone(),
            &mut MAX_TEXT_BYTES.clone(),
            &mut 8192,
            Some(TextContext {
                fonts: &[],
                styles: &[],
                types: &[],
                scheme: None,
                master: Some(&master),
                shapes: None,
                outline_slide_numbers: &[],
                slide_number: 0,
            }),
            None,
        )
        .unwrap()
        .unwrap();
        assert_eq!(out.matches("algn=\"ctr\"").count(), 1);
        assert_eq!(out.matches("sz=\"4800\"").count(), 2);
        assert_eq!(out.matches("sz=\"1800\"").count(), 2);
        assert!(out.contains("Title") && out.contains("Ordinary"));
    }

    #[test]
    fn preserves_nontext_geometry_paint_and_stacking_order() {
        let shape = |kind: u16, label: Option<&str>| {
            let mut atoms = vec![
                record((kind << 4) | 2, 0xf00a, &ints(&[42, 0xa00])),
                record(0, 0xf010, &ints(&[144, 288, 864, 720])),
                properties(&[(0x181, 0x00563412), (0x1c0, 255), (0x1cb, 25400)]),
            ];
            if let Some(label) = label {
                atoms.push(text(label));
            }
            record(15, 0xf004, &atoms.concat())
        };
        let out = xml(&drawing(vec![shape(3, None), shape(1, Some("Label"))])).unwrap();
        assert_eq!(out.matches("<p:sp>").count(), 2);
        assert_eq!(out.matches("<p:txBody>").count(), 1);
        assert!(out.contains("prst=\"ellipse\""));
        assert!(out.contains("<a:ln w=\"25400\" cap=\"flat\">"));
        assert!(out.contains("val=\"123456\""));
        assert!(out.find("ellipse").unwrap() < out.find("Label").unwrap());
        assert!(!out.contains("txBox=\"1\""));
    }

    #[test]
    fn adjusted_shapes_keep_text_without_painting_a_fake_preset() {
        let anchor = record(0, 0xf010, &ints(&[0, 0, 576, 576]));
        let options = properties(&[(0x147, 100), (0x181, 255)]);
        let bytes = drawing(vec![
            sp(0xa00, vec![anchor.clone(), options.clone()]),
            sp(0xa00, vec![anchor, options, text("Custom outline")]),
        ]);
        let out = xml(&bytes).unwrap();
        assert_eq!(out.matches("<p:sp>").count(), 1);
        assert!(out.contains("Custom outline"));
        assert!(!out.contains("solidFill"));
        let complex = record(0x13, 0xf00b, &[0x45, 0x81, 0, 0, 0, 0]);
        let out = xml(&drawing(vec![sp(0xa00, vec![complex])])).unwrap();
        assert!(out.is_empty());
    }

    #[test]
    fn nontext_shapes_share_shape_count_and_xml_budgets() {
        let bytes = sp(
            0xa00,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                properties(&[(0x181, 255)]),
            ],
        );
        for (id, budget, message) in [(100_001, 4096, "too many"), (1, 1, "OUTPUT_TOO_LARGE")] {
            let mut xml_budget = budget;
            let mut writer = Writer {
                inherited: false,
                outline: &[],
                records: &mut MAX_RECORDS.clone(),
                text: &mut MAX_TEXT_BYTES.clone(),
                remaining: &mut xml_budget,
                output: String::new(),
                id,
                context: None,
                media: None,
            };
            let record = parse_records(&bytes, &mut MAX_RECORDS.clone()).unwrap()[0];
            assert!(writer.node(record, false, 0).unwrap_err().contains(message));
        }
    }

    #[test]
    fn does_not_apply_inline_style_to_an_outline_reference() {
        for atoms in [
            [record(0, 4001, &[]), record(0, 3998, &[0; 4])].concat(),
            [record(0, 3998, &[0; 4]), record(0, 4001, &[])].concat(),
        ] {
            let bytes = drawing(vec![sp(
                0x200,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                    record(15, 0xf00d, &atoms),
                ],
            )]);
            assert!(xml(&bytes).is_err());
        }
    }

    #[test]
    fn preserves_signed_rotation_flips_and_emu_text_margins() {
        let bytes = drawing(vec![sp(
            0x2c0,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                text("Rotated"),
                properties(&[
                    (4, (-45i32 * 65536) as u32),
                    (0x81, 12700),
                    (0x82, 25400),
                    (0x83, 38100),
                    (0x84, 50800),
                    (0x85, 2),
                    (0x87, 4),
                ]),
            ],
        )]);
        let out = xml(&bytes).unwrap();
        assert!(out.contains("rot=\"-2700000\" flipH=\"1\" flipV=\"1\""));
        assert!(out.contains("wrap=\"none\" anchor=\"ctr\" anchorCtr=\"1\" lIns=\"12700\" tIns=\"25400\" rIns=\"38100\" bIns=\"50800\""));
    }

    #[test]
    fn validates_complex_property_tails_and_charges_property_work() {
        let malformed = drawing(vec![sp(0x200, vec![properties(&[(0x8380, 100)])])]);
        assert!(xml(&malformed)
            .unwrap_err()
            .contains("complex shape property"));
        let opts = properties(&[(4, 0), (0x85, 0)]);
        let entry = parse_records(&opts, &mut 1).unwrap()[0];
        assert!(Properties::default()
            .read(entry, &mut 1)
            .unwrap_err()
            .contains("work budget"));
    }

    #[test]
    fn escapes_multibyte_text_across_chunks_without_losing_unicode() {
        let text = format!("{}日本語<&", "x".repeat(1023));
        let mut output = String::new();
        paragraphs(&[text], &mut output, &mut 4096).unwrap();
        assert!(output.contains("日本語&lt;&amp;"));
    }

    #[test]
    fn rejects_zero_group_scale_missing_anchors_and_deep_nesting() {
        let group_header = sp(
            0x201,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                record(1, 0xf009, &ints(&[0, 0, 0, 576])),
            ],
        );
        assert!(xml(&drawing(vec![record(15, 0xf003, &group_header)])).is_err());
        assert!(xml(&drawing(vec![sp(0, vec![text("missing")])])).is_err());
        let mut group = sp(
            0x202,
            vec![record(0, 0xf00f, &ints(&[0, 0, 576, 576])), text("deep")],
        );
        for _ in 0..=MAX_DEPTH {
            group = record(
                15,
                0xf003,
                &[
                    sp(
                        0x203,
                        vec![
                            record(0, 0xf00f, &ints(&[0, 0, 576, 576])),
                            record(1, 0xf009, &ints(&[0, 0, 576, 576])),
                        ],
                    ),
                    group,
                ]
                .concat(),
            );
        }
        let mut writer = Writer {
            inherited: false,
            outline: &[],
            records: &mut MAX_RECORDS.clone(),
            text: &mut MAX_TEXT_BYTES.clone(),
            remaining: &mut (1024 * 1024),
            output: String::new(),
            id: 1,
            context: None,
            media: None,
        };
        let record = parse_records(&group, &mut MAX_RECORDS.clone()).unwrap()[0];
        assert!(writer
            .node(record, true, 0)
            .unwrap_err()
            .contains("nesting"));
    }

    #[test]
    fn separate_frames_use_ppt_top_left_order_and_both_anchor_widths() {
        let small: Vec<u8> = [144i16, -288, 576, 432]
            .iter()
            .flat_map(|x| x.to_le_bytes())
            .collect();
        let bytes = drawing(vec![
            sp(0x200, vec![record(0, 0xf010, &small), text("First")]),
            sp(
                0x200,
                vec![
                    record(0, 0xf010, &ints(&[576, 1152, 1728, 864])),
                    text("Second"),
                ],
            ),
        ]);
        let out = xml(&bytes).unwrap();
        assert_eq!(out.matches("<p:sp>").count(), 2);
        assert!(out.contains("<a:off x=\"-457200\" y=\"228600\"/>"));
        assert!(out.contains("<a:ext cx=\"1371600\" cy=\"457200\"/>"));
        assert!(out.contains("<a:off x=\"1828800\" y=\"914400\"/>"));
        assert!(out.find("First").unwrap() < out.find("Second").unwrap());
    }

    #[test]
    fn preserves_group_coordinates_and_resolves_only_owned_outline_text() {
        let group = record(
            15,
            0xf003,
            &[
                sp(
                    0x201,
                    vec![
                        record(0, 0xf010, &ints(&[288, 576, 1728, 864])),
                        record(1, 0xf009, &ints(&[100, 200, 500, 600])),
                    ],
                ),
                sp(
                    0x202,
                    vec![
                        record(0, 0xf00f, &ints(&[100, 300, 300, 400])),
                        record(15, 0xf00d, &record(0, 3998, &ints(&[0]))),
                    ],
                ),
            ]
            .concat(),
        );
        let out = xml(&drawing(vec![group])).unwrap();
        assert!(out.contains("<p:grpSp>"));
        assert!(out.contains("<a:chOff x=\"158750\" y=\"317500\"/>"));
        assert!(out.contains("<a:chExt cx=\"635000\" cy=\"635000\"/>"));
        assert!(out.contains("<a:off x=\"158750\" y=\"476250\"/>"));
        assert_eq!(out.matches("Outline").count(), 1);
        assert!(out.contains("id=\"2\""));
        assert!(out.contains("id=\"3\""));
    }

    #[test]
    fn skips_deleted_shapes_and_does_not_collect_client_data_text() {
        let anchor = record(0, 0xf010, &ints(&[0, 0, 576, 576]));
        let out = xml(&drawing(vec![
            sp(0x208, vec![anchor.clone(), text("Deleted")]),
            sp(0x210, vec![anchor.clone(), text("OLE")]),
            sp(
                0x200,
                vec![
                    anchor,
                    text("Visible"),
                    record(15, 0xf011, &record(0, 4008, b"Action")),
                ],
            ),
        ]))
        .unwrap();
        assert!(out.contains("Visible"));
        for omitted in ["Deleted", "OLE", "Action"] {
            assert!(!out.contains(omitted));
        }
    }

    #[test]
    fn rejects_truncated_anchors_and_unbounded_xml() {
        assert!(xml(&drawing(vec![sp(
            0x200,
            vec![record(0, 0xf010, &[0; 7]), text("x")]
        )]))
        .is_err());
        let bytes = drawing(vec![sp(
            0x200,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                text("\r".repeat(100).as_str()),
            ],
        )]);
        assert!(render(
            &bytes,
            &[],
            &mut MAX_RECORDS.clone(),
            &mut MAX_TEXT_BYTES.clone(),
            &mut 1024,
            None,
            None,
        )
        .is_err());
    }
}
