//! Validated reading of the live slide's and masters' OfficeArt shape trees:
//! shape records, properties, placeholders, backgrounds and master shapes.
//! [MS-PPT] 2.5.13, 2.7.1, 2.9.76 and [MS-ODRAW] 2.2.14/16/38/39/40.
//! `direct_model` projects them into the presentation model, with placement in
//! ECMA-376 DrawingML CT_(Group)Transform2D terms (`direct_transform`).
use super::*;
use crate::officeart::geometry;

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

mod direct_geometry;
pub(super) mod direct_model;
mod direct_transform;
#[cfg(test)]
mod gradient_integration_tests;

trait ShapeSource {
    type Record: Clone;
    type Complex: Default + Clone;
    type Style;
    fn with_record<T>(
        &self,
        record: &Self::Record,
        f: impl FnOnce(Record<'_>) -> Result<T, String>,
    ) -> Result<T, String>;
    fn children(
        &self,
        record: &Self::Record,
        budget: &mut usize,
    ) -> Result<Vec<Self::Record>, String>;
    fn primary(
        &self,
        record: &Self::Record,
        props: &mut PropertiesStorage<Self::Complex>,
        budget: &mut usize,
    ) -> Result<(), String>;
    fn tertiary(
        &self,
        record: &Self::Record,
        props: &mut PropertiesStorage<Self::Complex>,
        budget: &mut usize,
    ) -> Result<(), String>;
    fn style(
        &self,
        record: &Self::Record,
        budget: &mut usize,
    ) -> Result<Option<Self::Style>, String>;
}

struct SpannedSlideSource<'a> {
    backing: &'a [u8],
}
/// Shape records as validated ranges of the retained document stream; the
/// local PP9 text style is retained as a range too.
impl ShapeSource for SpannedSlideSource<'_> {
    type Record = RecordSpan;
    type Complex = ByteSpan;
    type Style = ByteSpan;
    fn with_record<T>(
        &self,
        record: &Self::Record,
        f: impl FnOnce(Record<'_>) -> Result<T, String>,
    ) -> Result<T, String> {
        f(record.view(self.backing)?)
    }
    fn children(
        &self,
        record: &Self::Record,
        budget: &mut usize,
    ) -> Result<Vec<Self::Record>, String> {
        parse_record_spans(self.backing, record.payload_span(), budget)
    }
    fn primary(
        &self,
        record: &Self::Record,
        props: &mut PropertiesStorage<Self::Complex>,
        budget: &mut usize,
    ) -> Result<(), String> {
        props.read_span(record, self.backing, budget)
    }
    fn tertiary(
        &self,
        record: &Self::Record,
        props: &mut PropertiesStorage<Self::Complex>,
        budget: &mut usize,
    ) -> Result<(), String> {
        props.read_tertiary_span(record, self.backing, budget)
    }
    fn style(
        &self,
        record: &Self::Record,
        budget: &mut usize,
    ) -> Result<Option<Self::Style>, String> {
        text_style::auto_number::local_atom_span(record, self.backing, budget)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PlaceholderSize {
    Full,
    Half,
    Quarter,
    /// Retained without interpretation as a defined value. MS-PPT 2.13.22 defines
    /// only values 0..=2, but the pre-existing reader accepted every byte.
    Unknown(u8),
}

impl From<u8> for PlaceholderSize {
    fn from(value: u8) -> Self {
        match value {
            0 => Self::Full,
            1 => Self::Half,
            2 => Self::Quarter,
            value => Self::Unknown(value),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct PlaceholderMetadata {
    position: i32,
    placement_id: u8,
    preferred_size: PlaceholderSize,
}

impl PlaceholderMetadata {
    fn is_placeholder(self) -> bool {
        // MS-PPT 2.7.8: -1 marks a present PlaceholderAtom whose shape is not
        // a placeholder. Atom presence and placeholder status are distinct.
        self.position != -1
    }
}

struct ShapeStorage<R, C, S> {
    id: u32,
    kind: u16,
    flags: u32,
    anchor: Option<Rect>,
    child_space: Option<Rect>,
    textbox: Option<R>,
    style9: Option<S>,
    placeholder: Option<PlaceholderMetadata>,
    props: PropertiesStorage<C>,
}
type SpannedShape = ShapeStorage<RecordSpan, ByteSpan, ByteSpan>;

impl<R: Clone, C: Default + Clone, S> ShapeStorage<R, C, S> {
    fn is_placeholder(&self) -> bool {
        self.placeholder
            .is_some_and(PlaceholderMetadata::is_placeholder)
    }

    fn read_from(
        source: &impl ShapeSource<Record = R, Complex = C, Style = S>,
        record: R,
        nested: bool,
        budget: &mut usize,
    ) -> Result<Self, String> {
        source.with_record(&record, |view| {
            if view.kind == 0xf004 && view.version == 15 {
                Ok(())
            } else {
                Err(unsupported("invalid PowerPoint shape container"))
            }
        })?;
        let mut flags = None;
        let mut id = 0;
        let mut kind = 0;
        let mut anchor = None;
        let mut child_space = None;
        let mut textbox = None;
        let mut style9 = None;
        let mut tags_seen = false;
        let mut placeholder = None;
        let mut ole_ref = None;
        let mut recolor = None;
        let mut tertiary_seen = false;
        let mut props = PropertiesStorage::<C>::default();
        for child in source.children(&record, budget)? {
            let (child_kind, child_version, child_instance, child_len) = source
                .with_record(&child, |view| {
                    Ok((view.kind, view.version, view.instance, view.payload.len()))
                })?;
            match child_kind {
                0xf00a => {
                    if flags.is_some() || child_version != 2 || child_len != 8 {
                        return Err(unsupported("invalid PowerPoint shape flags"));
                    }
                    (id, flags) = source.with_record(&child, |view| {
                        Ok((u32_at(view.payload, 0)?, Some(u32_at(view.payload, 4)?)))
                    })?;
                    kind = child_instance;
                }
                0xf010 | 0xf00f => {
                    if anchor.is_some() || (child_kind == 0xf00f) != nested {
                        return Err(unsupported("ambiguous PowerPoint shape coordinate space"));
                    }
                    anchor = Some(source.with_record(&child, Rect::read)?);
                }
                0xf009 => {
                    if child_space.is_some() {
                        return Err(unsupported("duplicate PowerPoint group bounds"));
                    }
                    child_space = Some(source.with_record(&child, Rect::read)?);
                }
                0xf00d => {
                    if textbox.is_some() || child_version != 15 {
                        return Err(unsupported("invalid PowerPoint shape text container"));
                    }
                    textbox = Some(child);
                }
                0xf00b => source.primary(&child, &mut props, budget)?,
                0xf122 => {
                    if tertiary_seen {
                        return Err(unsupported("duplicate PowerPoint tertiary properties"));
                    }
                    tertiary_seen = true;
                    source.tertiary(&child, &mut props, budget)?;
                }
                0xf011 => {
                    if child_version != 15 {
                        return Err(unsupported("invalid PowerPoint client data"));
                    }
                    // Inspect direct placeholder metadata and exact passive PP9
                    // tags only. Never descend into action/link containers.
                    for atom in source.children(&child, budget)? {
                        let (atom_kind, atom_version, atom_len) = source
                            .with_record(&atom, |view| {
                                Ok((view.kind, view.version, view.payload.len()))
                            })?;
                        if atom_kind == 5000 {
                            if tags_seen {
                                return Err(unsupported("duplicate PowerPoint shape tags"));
                            }
                            tags_seen = true;
                            style9 = source.style(&atom, budget)?;
                        }
                        if atom_kind == 0x0bc1 {
                            // MS-PPT 2.7.7: recVer 0, recLen 4. The reference is
                            // only an identifier; it is resolved, never followed
                            // into object storage.
                            if ole_ref.is_some() || atom_version != 0 || atom_len != 4 {
                                return Err(unsupported("invalid PowerPoint ExObjRefAtom"));
                            }
                            ole_ref =
                                Some(source.with_record(&atom, |view| u32_at(view.payload, 0))?);
                            continue;
                        }
                        if atom_kind == 0x0fe7 {
                            // MS-PPT 2.7.9: a 12-byte fixed part precedes the
                            // entries; bit 0 of the first field is fShouldRecolor.
                            if recolor.is_some() || atom_version != 0 || atom_len < 12 {
                                return Err(unsupported("invalid PowerPoint RecolorInfoAtom"));
                            }
                            recolor = Some(
                                source.with_record(&atom, |view| Ok(view.payload[0] & 1 != 0))?,
                            );
                            continue;
                        }
                        if atom_kind != 3011 {
                            continue;
                        }
                        if placeholder.is_some() || atom_version != 0 || atom_len != 8 {
                            return Err(unsupported("invalid PowerPoint placeholder metadata"));
                        }
                        placeholder = Some(source.with_record(&atom, |view| {
                            Ok(PlaceholderMetadata {
                                position: u32_at(view.payload, 0)? as i32,
                                placement_id: view.payload[4],
                                preferred_size: view.payload[5].into(),
                                // MS-PPT 2.7.8: payload bytes 6..8 are undefined
                                // and MUST be ignored.
                            })
                        })?);
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
            style9,
            placeholder,
            props: PropertiesStorage {
                ole_ref,
                recolor: recolor.unwrap_or(false),
                ..props
            },
        })
    }
}

impl<R, C, S> ShapeStorage<R, C, S> {
    /// Master-shape metadata omission: deleted (fDeleted), OLE (fOleShape) and
    /// background (fBackground) shapes and active script anchors contribute no
    /// inherited master-shape properties (MS-ODRAW 2.2.40, 2.3.4.44).
    fn omitted(&self) -> bool {
        self.flags & (8 | 16 | 1024) != 0 || self.props.script
    }
    /// Direct-model omission. Unlike master-shape metadata, an OLE shape
    /// (fOleShape, MS-ODRAW 2.2.40) is not omitted: it is a picture frame whose
    /// pib names the BLIP to display (MS-ODRAW 2.3.23.5), and the direct model
    /// either shows that stored presentation picture or rejects the shape.
    fn direct_omitted(&self) -> bool {
        self.flags & (8 | 1024) != 0 || self.props.script
    }
    fn is_ole(&self) -> bool {
        self.flags & 16 != 0
    }
    fn master(&self) -> Option<u32> {
        (self.flags & 0x20 != 0).then_some(self.props.master.unwrap_or(0))
    }
}

struct PropertiesStorage<T> {
    geometry: geometry::GeometryStorage<T>,
    gradient: crate::officeart::gradient::Storage<T>,
    hidden: bool,
    script: bool,
    master: Option<u32>,
    picture: u32,
    /// pib_complex (MS-ODRAW 2.3.23.6): a picture named by file, not a BLIP.
    picture_linked: bool,
    /// First non-default MS-ODRAW 2.3.23 display adjustment without a
    /// projection (recolor and color modifiers), if any; rejected.
    picture_adjustment: Option<&'static str>,
    /// MS-ODRAW 2.3.23.10 pictureTransparent (OfficeArtCOLORREF), when set.
    picture_transparent: Option<u32>,
    /// MS-ODRAW 2.3.23.11-12 pictureContrast / pictureBrightness, when set.
    picture_contrast: Option<u32>,
    picture_brightness: Option<u32>,
    /// MS-ODRAW 2.3.23.35 fPictureGray / fPictureBiLevel, when their use bits
    /// are set.
    picture_gray: Option<bool>,
    picture_bilevel: Option<bool>,
    /// MS-PPT 2.7.7 ExObjRefAtom from the shape's client data: the external
    /// object behind an OLE shape.
    ole_ref: Option<u32>,
    /// MS-PPT 2.7.9 RecolorInfoAtom.fShouldRecolor from the shape's client
    /// data: metafile color remapping of the displayed picture.
    recolor: bool,
    crop: [i64; 4],
    paint: paint::Paint,
    rotation: i64,
    margins: [u32; 4],
    wrap: &'static str,
    anchor: &'static str,
    center: bool,
    text_flow: Option<u32>,
    font_direction: Option<u32>,
    /// MS-ODRAW 2.3.21.15 fFitShapeToText (honored only with its use bit).
    fit_shape_to_text: bool,
    /// MS-ODRAW 2.3.4.41 metroBlob: the alternative shape XML package, adopted
    /// only when it agrees with the binary shape (see `ppt::metro`). A second
    /// copy makes the alternative ambiguous and it is never adopted.
    metro: Option<T>,
    metro_ambiguous: bool,
}
type SpannedProperties = PropertiesStorage<ByteSpan>;

impl<T> Default for PropertiesStorage<T> {
    fn default() -> Self {
        Self {
            geometry: geometry::GeometryStorage::default(),
            gradient: crate::officeart::gradient::Storage::default(),
            hidden: false,
            script: false,
            master: None,
            picture: 0,
            picture_linked: false,
            picture_adjustment: None,
            picture_transparent: None,
            picture_contrast: None,
            picture_brightness: None,
            picture_gray: None,
            picture_bilevel: None,
            ole_ref: None,
            recolor: false,
            crop: [0; 4],
            paint: paint::Paint::default(),
            rotation: 0,
            margins: [91440, 45720, 91440, 45720],
            wrap: "square",
            anchor: "t",
            center: false,
            text_flow: None,
            font_direction: None,
            fit_shape_to_text: false,
            metro: None,
            metro_ambiguous: false,
        }
    }
}
impl<T: Default + Clone> PropertiesStorage<T> {
    fn apply_tertiary(&mut self, opid: u16, value: u32, complex: Option<T>) -> Result<(), String> {
        if opid == 0x01bf && complex.is_none() {
            self.paint.tertiary_fill_boolean_property(value)?;
        }
        if opid == 0x013f && complex.is_none() {
            self.blip_booleans(value)?;
        }
        if opid & 0x3fff == 0x03a9 {
            if let Some(complex) = complex {
                if self.metro.replace(complex).is_some() {
                    self.metro_ambiguous = true;
                }
            }
        }
        Ok(())
    }

    /// MS-ODRAW 2.3.23.35: fPictureGray is bit 2 and fPictureBiLevel bit 1,
    /// each honored only with its use bit (18 and 17). Other Blip Booleans
    /// (hit testing, looping, active OLE server) do not change the display.
    /// Primary and tertiary tables have no documented precedence, so
    /// contradictory active bits are rejected.
    fn blip_booleans(&mut self, value: u32) -> Result<(), String> {
        for (use_bit, bit, target) in [
            (18, 2, &mut self.picture_gray),
            (17, 1, &mut self.picture_bilevel),
        ] {
            if value & (1 << use_bit) == 0 {
                continue;
            }
            let active = value & (1 << bit) != 0;
            if target.is_some_and(|current| current != active) {
                return Err(unsupported("ambiguous PowerPoint picture color mode"));
            }
            *target = Some(active);
        }
        Ok(())
    }

    fn apply_primary(&mut self, opid: u16, value: u32, complex: Option<T>) -> Result<(), String> {
        if matches!(opid & 0x3fff, 0x88 | 0x89) {
            if opid & 0xc000 != 0 || complex.is_some() {
                return Err(unsupported("invalid PowerPoint text direction property"));
            }
            let target = if opid == 0x88 {
                if value > 5 {
                    return Err(unsupported("invalid PowerPoint text flow"));
                }
                &mut self.text_flow
            } else {
                if value > 3 {
                    return Err(unsupported("invalid PowerPoint font direction"));
                }
                &mut self.font_direction
            };
            if target.replace(value).is_some() {
                return Err(unsupported("duplicate PowerPoint text direction property"));
            }
            return Ok(());
        }
        if let Some(complex) = complex {
            if opid & 0x3fff == 0x197 {
                self.gradient.set(complex);
                return Ok(());
            }
            if opid & 0x3fff == 0x104 {
                self.picture_linked = true;
                return Ok(());
            }
            if matches!(opid & 0x3fff, 0x145..=0x150) {
                self.paint.custom_geometry = true;
            }
            self.geometry.complex(opid & 0x3fff, complex);
            return Ok(());
        }
        if matches!(opid & 0x3fff, 0x145 | 0x146) {
            self.geometry.scalar(opid & 0x3fff, value)?;
        }
        if opid & 0x3fff == 0x197 {
            self.gradient.scalar(value);
            return Ok(());
        }
        if opid & 0x4000 != 0 {
            if opid == 0x4104 {
                self.picture = value;
            } else if opid == 0x4186 {
                self.paint.property(opid, value)?;
            }
            return Ok(());
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
            // MS-ODRAW 2.3.23.10-12 and 24-32: defaults are no transparent
            // color, contrast 0x10000, brightness 0, no recolor color and an
            // MSOTINTSHADE of 0x20000000 for the Ext modifiers.
            0x107 => self.picture_transparent = (value != 0xffff_ffff).then_some(value),
            0x115 | 0x11a | 0x11b if value != 0xffff_ffff => {
                self.picture_adjustment.get_or_insert(if opid == 0x115 {
                    "extended transparent color"
                } else {
                    "recolor"
                });
            }
            0x117 | 0x11d if value != 0x2000_0000 => {
                self.picture_adjustment.get_or_insert("color modification");
            }
            0x108 => self.picture_contrast = Some(value),
            0x109 => self.picture_brightness = Some(value),
            0x13f => self.blip_booleans(value)?,
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
            // [MS-ODRAW] 2.3.18.5: signed 16.16 degrees -> nearest 1/60000
            // degree (45.01 is stored as 2949775/65536 and reads as 2700600).
            4 => {
                let value = i64::from(value as i32) * 60000;
                self.rotation = (value + value.signum() * 32768) / 65536;
            }
            0x81..=0x84 => {
                if value > 0x132f540 {
                    return Err(unsupported("invalid PowerPoint text margin"));
                }
                self.margins[usize::from(opid - 0x81)] = value;
            }
            0x85 => self.wrap = if value == 2 { "none" } else { "square" },
            // MS-ODRAW 2.3.21.15 Text Boolean Properties: fFitShapeToText is
            // bit 1 and fUsefFitShapeToText bit 17; the default is false.
            0xbf => {
                if value & (1 << 17) != 0 {
                    self.fit_shape_to_text = value & (1 << 1) != 0;
                }
            }
            0x87 if value <= 5 => {
                self.anchor = ["t", "ctr", "b"][(value % 3) as usize];
                self.center = value >= 3;
            }
            _ => {
                self.geometry.scalar(opid, value)?;
                self.paint.property(opid, value)?;
            }
        }
        Ok(())
    }
}

impl SpannedProperties {
    fn read_span(
        &mut self,
        record: &RecordSpan,
        backing: &[u8],
        budget: &mut usize,
    ) -> Result<(), String> {
        crate::officeart::properties::visit_span(record, backing, budget, |property| {
            self.apply_primary(property.opid, property.value, property.complex)
        })
    }

    fn read_tertiary_span(
        &mut self,
        record: &RecordSpan,
        backing: &[u8],
        budget: &mut usize,
    ) -> Result<(), String> {
        crate::officeart::properties::visit_tertiary_span(record, backing, budget, |property| {
            self.apply_tertiary(property.opid, property.value, property.complex)
        })
    }
}

/// A slide background is the ungrouped OfficeArt background shape, not an
/// arbitrary full-slide rectangle. Never inspect nested client/action data.
#[derive(Clone)]
pub(super) struct BackgroundStorage<T> {
    pub paint: paint::Paint,
    pub gradient: crate::officeart::gradient::Storage<T>,
}
pub(super) type SpannedBackground = BackgroundStorage<ByteSpan>;

pub(super) fn spanned_background(
    backing: &[u8],
    slide: &RecordSpan,
    budget: &mut usize,
) -> Result<Option<SpannedBackground>, String> {
    background_from(&SpannedSlideSource { backing }, slide, budget)
}

fn background_from<S: ShapeSource>(
    source: &S,
    slide: &S::Record,
    budget: &mut usize,
) -> Result<Option<BackgroundStorage<S::Complex>>, String> {
    let mut drawings = Vec::new();
    for record in source.children(slide, budget)? {
        if source.with_record(&record, |view| Ok(view.kind == 1036))? {
            drawings.push(record);
        }
    }
    let mut drawings = drawings.into_iter();
    let Some(drawing) = drawings.next() else {
        return Ok(None);
    };
    if drawings.next().is_some() || source.with_record(&drawing, |view| Ok(view.version != 15))? {
        return Err(unsupported("invalid PowerPoint background drawing"));
    }
    let groups = source.children(&drawing, budget)?;
    if groups.len() != 1 {
        return Err(unsupported(
            "invalid PowerPoint background OfficeArt drawing",
        ));
    }
    if source.with_record(&groups[0], |view| {
        Ok(view.kind != 0xf002 || view.version != 15)
    })? {
        return Err(unsupported(
            "invalid PowerPoint background OfficeArt drawing",
        ));
    }
    let mut result = None;
    for record in source.children(&groups[0], budget)? {
        if source.with_record(&record, |view| Ok(view.kind != 0xf004))? {
            continue;
        }
        let mut flags = Vec::new();
        for child in source.children(&record, budget)? {
            if source.with_record(&child, |view| Ok(view.kind == 0xf00a))? {
                flags.push(child);
            }
        }
        if flags.len() != 1 {
            return Err(unsupported("invalid PowerPoint background shape flags"));
        }
        let value = source.with_record(&flags[0], |flag| {
            if flag.version != 2 || flag.payload.len() != 8 {
                return Err(unsupported("invalid PowerPoint background shape flags"));
            }
            u32_at(flag.payload, 4)
        })?;
        if value & 1024 == 0 || value & (8 | 16) != 0 {
            continue;
        }
        if result.is_some() {
            return Err(unsupported("duplicate PowerPoint background shapes"));
        }
        let props = ShapeStorage::read_from(source, record, false, budget)?.props;
        result = Some(BackgroundStorage {
            paint: props.paint,
            gradient: props.gradient,
        });
    }
    Ok(result)
}

pub(super) fn master_shapes(
    backing: &[u8],
    slide: &RecordSpan,
    base: Option<std::rc::Rc<text_style::Master>>,
    output: &mut shape_master::Resolver,
    budget: &mut usize,
    text_budget: &mut usize,
) -> Result<(), String> {
    #[allow(clippy::too_many_arguments)]
    fn visit(
        backing: &[u8],
        record: RecordSpan,
        nested: bool,
        depth: usize,
        base: &Option<std::rc::Rc<text_style::Master>>,
        output: &mut shape_master::Resolver,
        budget: &mut usize,
        text_budget: &mut usize,
    ) -> Result<(), String> {
        if depth >= MAX_DEPTH {
            return Err(unsupported("PowerPoint master drawing depth exceeded"));
        }
        let viewed = record.view(backing)?;
        if viewed.kind == 0xf003 {
            if viewed.version != 15 {
                return Err(unsupported("invalid PowerPoint master group"));
            }
            let children = parse_record_spans(backing, record.payload_span(), budget)?;
            let first = children
                .first()
                .ok_or_else(|| unsupported("empty PowerPoint master group"))?;
            let group = SpannedShape::read_from(
                &SpannedSlideSource { backing },
                first.clone(),
                nested,
                budget,
            )?;
            if group.flags & 1 == 0 || (nested && group.flags & 4 != 0) {
                return Err(unsupported("invalid PowerPoint master group flags"));
            }
            if group.omitted() {
                return Ok(());
            }
            let child_nested = group.flags & 4 == 0;
            visit(
                backing,
                first.clone(),
                nested,
                depth + 1,
                base,
                output,
                budget,
                text_budget,
            )?;
            for child in &children[1..] {
                visit(
                    backing,
                    child.clone(),
                    child_nested,
                    depth + 1,
                    base,
                    output,
                    budget,
                    text_budget,
                )?;
            }
        } else if viewed.kind == 0xf004 {
            let shape = SpannedShape::read_from(
                &SpannedSlideSource { backing },
                record.clone(),
                nested,
                budget,
            )?;
            if shape.omitted() {
                return Ok(());
            }
            let (mut kind, mut text, mut style) = (None, None, None);
            if let Some(ref textbox) = shape.textbox {
                for atom in parse_record_spans(backing, textbox.payload_span(), budget)? {
                    let atom_view = atom.view(backing)?;
                    match atom_view.kind {
                        3999 => {
                            if kind.is_some() {
                                return Err(unsupported("duplicate master text header"));
                            }
                            kind = Some(text_style::text_type(atom_view)?);
                        }
                        TEXT_CHARS_ATOM | TEXT_BYTES_ATOM => {
                            if text.is_some() {
                                return Err(unsupported("duplicate master text body"));
                            }
                            let decoded = decode_text(atom_view)?;
                            charge_text(text_budget, decoded.len())?;
                            text = Some(decoded);
                        }
                        4001 => {
                            if style.is_some() || atom_view.version != 0 {
                                return Err(unsupported("invalid master text style"));
                            }
                            style = Some(atom.payload_span().clone());
                        }
                        _ => {} // Actions, links and metacharacter evaluation remain absent.
                    }
                }
            }
            let direct = match (text.as_deref(), style.as_ref()) {
                (Some(text), Some(style)) => {
                    text_style::shape_levels(text, style.view(backing)?, budget)?
                }
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
                gradient: shape.props.gradient,
            })?;
        }
        Ok(())
    }
    for drawing in parse_record_spans(backing, slide.payload_span(), budget)? {
        if drawing.view(backing)?.kind != 1036 {
            continue;
        }
        for dg in parse_record_spans(backing, drawing.payload_span(), budget)? {
            let dg_view = dg.view(backing)?;
            if dg_view.kind != 0xf002 || dg_view.version != 15 {
                return Err(unsupported("invalid master OfficeArt drawing"));
            }
            for child in parse_record_spans(backing, dg.payload_span(), budget)? {
                visit(backing, child, false, 0, &base, output, budget, text_budget)?;
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::persist::tests::record;
    use super::*;
    use pptx_model::{Fill, ShapeElement, Slide, SlideElement, TextRun};

    fn ints(values: &[i32]) -> Vec<u8> {
        values.iter().flat_map(|n| n.to_le_bytes()).collect()
    }
    fn sp(flags: i32, children: Vec<Vec<u8>>) -> Vec<u8> {
        sp_kind(202, flags, children)
    }
    fn sp_kind(kind: u16, flags: i32, children: Vec<Vec<u8>>) -> Vec<u8> {
        record(
            15,
            0xf004,
            &[
                vec![record((kind << 4) | 2, 0xf00a, &ints(&[42, flags]))],
                children,
            ]
            .concat()
            .concat(),
        )
    }
    fn text(value: &str) -> Vec<u8> {
        record(15, 0xf00d, &record(0, 4008, value.as_bytes()))
    }
    /// A TextRulerAtom giving level 0 explicit zero origins (MS-PPT 2.9.30),
    /// so plain text boxes project without a text master.
    fn origin_ruler() -> Vec<u8> {
        record(
            0,
            4006,
            &[
                (8u32 | 256).to_le_bytes().as_slice(),
                &0i16.to_le_bytes(),
                &0i16.to_le_bytes(),
            ]
            .concat(),
        )
    }
    fn ruled_text(value: &str) -> Vec<u8> {
        record(
            15,
            0xf00d,
            &[record(0, 4008, value.as_bytes()), origin_ruler()].concat(),
        )
    }
    fn drawing(shapes: Vec<Vec<u8>>) -> Vec<u8> {
        record(15, 1036, &record(15, 0xf002, &shapes.concat()))
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
    fn tertiary_properties(values: &[(u16, u32)]) -> Vec<u8> {
        let mut bytes = properties(values);
        bytes[2..4].copy_from_slice(&0xf122u16.to_le_bytes());
        bytes
    }
    fn png_blip() -> Vec<u8> {
        // Complete 2x1 RGBA PNG, including valid zlib data and CRCs.
        let png = vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 1,
            8, 6, 0, 0, 0, 244, 34, 127, 138, 0, 0, 0, 14, 73, 68, 65, 84, 120, 156, 99, 248, 207,
            192, 0, 66, 13, 0, 15, 122, 3, 126, 119, 233, 127, 151, 0, 0, 0, 0, 73, 69, 78, 68,
            174, 66, 96, 130,
        ];
        record(0x6e00, 0xf01e, &[vec![0; 17], png].concat())
    }

    fn read_shape(bytes: &[u8], budget: &mut usize) -> Result<SpannedShape, String> {
        let (span, _) = record_span_with_end(bytes, 0, &mut 1, "shape")?;
        SpannedShape::read_from(&SpannedSlideSource { backing: bytes }, span, false, budget)
    }

    fn read_properties(bytes: &[u8], budget: &mut usize) -> Result<SpannedProperties, String> {
        let (span, _) = record_span_with_end(bytes, 0, &mut 1, "properties")?;
        let mut props = SpannedProperties::default();
        props.read_span(&span, bytes, budget)?;
        Ok(props)
    }

    fn slide_span(tree: &[u8]) -> (Vec<u8>, RecordSpan) {
        let slide = record(15, SLIDE_CONTAINER, tree);
        let span = record_span_with_end(&slide, 0, &mut 1, "slide").unwrap().0;
        (slide, span)
    }

    fn background_of(tree: &[u8], budget: &mut usize) -> Result<Option<SpannedBackground>, String> {
        let (slide, span) = slide_span(tree);
        spanned_background(&slide, &span, budget)
    }

    /// Project the slide whose drawing is `tree` through the direct model, with
    /// one outline block and `blips` as the image store. Returns the slide and
    /// the image store indexes it admitted.
    fn project_slide(
        tree: &[u8],
        blips: &[Vec<u8>],
        mut text_budget: usize,
    ) -> Result<(Slide, Vec<u32>), String> {
        let mut document = record(15, SLIDE_CONTAINER, tree);
        let mut offsets = Vec::new();
        for blip in blips {
            offsets.push(document.len());
            document.extend_from_slice(blip);
        }
        let slide = record_span_with_end(&document, 0, &mut 1, "slide")?.0;
        let entries = offsets
            .into_iter()
            .map(|offset| record_span_with_end(&document, offset, &mut 1, "blip").map(|r| r.0))
            .collect::<Result<Vec<_>, _>>()?;
        let presentation = persist::PresentationStorage {
            shape_masters: shape_master::Resolver::default(),
            slides: vec![(slide, vec!["Outline".to_owned()])],
            outline_styles: vec![Vec::new()],
            outline_types: vec![Vec::new()],
            outline_slide_numbers: vec![Vec::new()],
            first_slide_number: 1,
            text_masters: vec![None],
            metro_themes: vec![None],
            document_text_axes: None,
            fonts: Vec::new(),
            schemes: vec![None],
            image_entries: entries.clone(),
            ole_objects: media::OleCatalog::default(),
            backgrounds: vec![None],
            object_masters: vec![std::rc::Rc::from([])],
            size: (720, 540),
        };
        let mut media = media::SpanStore::new(entries);
        let slide = direct_model::slide(
            0,
            &presentation,
            &document,
            None,
            &mut media,
            &mut MAX_RECORDS.clone(),
            &mut text_budget,
            &mut (256 * 1024 * 1024),
        )?;
        let used = media.used_images().map(|(id, _)| id).collect();
        Ok((slide, used))
    }

    fn project(tree: &[u8]) -> Result<Slide, String> {
        project_slide(tree, &[], MAX_TEXT_BYTES).map(|(slide, _)| slide)
    }

    fn shapes(slide: &Slide) -> Vec<&ShapeElement> {
        slide
            .elements
            .iter()
            .map(|element| match element {
                SlideElement::Shape(shape) => shape,
                _ => panic!("expected only shapes"),
            })
            .collect()
    }

    fn shape_text(shape: &ShapeElement) -> String {
        shape
            .text_body
            .iter()
            .flat_map(|body| &body.paragraphs)
            .flat_map(|paragraph| &paragraph.runs)
            .filter_map(|run| match run {
                TextRun::Text(run) => Some(run.text.as_str()),
                _ => None,
            })
            .collect()
    }

    fn shape_with_placeholder(
        position: i32,
        placement_id: u8,
        preferred_size: u8,
        unused: [u8; 2],
    ) -> Vec<u8> {
        sp(
            0,
            vec![record(
                15,
                0xf011,
                &record(
                    0,
                    3011,
                    &[
                        position.to_le_bytes().as_slice(),
                        &[placement_id, preferred_size],
                        &unused,
                    ]
                    .concat(),
                ),
            )],
        )
    }

    #[test]
    fn retains_placeholder_presence_identity_position_and_preferred_size() {
        let absent = read_shape(&sp(0, vec![]), &mut 10).unwrap();
        assert_eq!(absent.placeholder, None);
        assert!(!absent.is_placeholder());

        for (raw, expected) in [
            (0, PlaceholderSize::Full),
            (1, PlaceholderSize::Half),
            (2, PlaceholderSize::Quarter),
            (3, PlaceholderSize::Unknown(3)),
            (u8::MAX, PlaceholderSize::Unknown(u8::MAX)),
        ] {
            let bytes = shape_with_placeholder(-17, 0x1a, raw, [0x55, 0xaa]);
            let shape = read_shape(&bytes, &mut 20).unwrap();
            assert_eq!(
                shape.placeholder,
                Some(PlaceholderMetadata {
                    position: -17,
                    placement_id: 0x1a,
                    preferred_size: expected,
                })
            );
            assert!(shape.is_placeholder());
            assert!(shape.textbox.is_none());
        }

        let shape = read_shape(&shape_with_placeholder(-1, 7, 0, [1, 2]), &mut 20).unwrap();
        assert!(shape.placeholder.is_some());
        assert!(!shape.is_placeholder());

        // Retained metadata keeps every position once the source is gone.
        for position in [i32::MIN, -1, 0, i32::MAX] {
            let bytes = shape_with_placeholder(position, u8::MAX, 2, [0xde, 0xad]);
            let shape = read_shape(&bytes, &mut 20).unwrap();
            drop(bytes);
            assert_eq!(shape.placeholder.unwrap().position, position);
        }
    }

    #[test]
    fn placeholder_metadata_keeps_existing_duplicate_version_and_length_rejections() {
        let payload = [0u8; 8];
        let client = |atoms: Vec<Vec<u8>>| sp(0, vec![record(15, 0xf011, &atoms.concat())]);
        for bytes in [
            client(vec![record(1, 3011, &payload)]),
            client(vec![record(0, 3011, &[0; 7])]),
            client(vec![record(0, 3011, &payload), record(0, 3011, &payload)]),
        ] {
            let error = read_shape(&bytes, &mut 20)
                .err()
                .expect("invalid placeholder metadata must fail");
            assert!(error.contains("invalid PowerPoint placeholder metadata"));
        }
    }

    #[test]
    fn shape_source_reads_structure_and_rejects_truncation_and_exhausted_work() {
        let bytes = sp(
            0x20,
            vec![
                record(0, 0xf010, &ints(&[-2, 3, 574, 291])),
                properties(&[(0x301, 77), (0x145, 0), (0x146, 0)]),
            ],
        );
        let shape = read_shape(&bytes, &mut 20).unwrap();
        assert_eq!((shape.id, shape.flags), (42, 0x20));
        // PPT RectStruct order is top/left/right/bottom.
        assert_eq!(
            shape.anchor,
            Some(Rect {
                x: master_to_emu(3),
                y: master_to_emu(-2),
                w: master_to_emu(574) - master_to_emu(3),
                h: master_to_emu(291) - master_to_emu(-2),
            })
        );
        assert_eq!(shape.master(), Some(77));
        let moved = bytes.clone();
        assert!(shape
            .props
            .geometry
            .view(&moved)
            .unwrap()
            .decode(&mut 10)
            .unwrap()
            .is_none());
        for length in 0..bytes.len() {
            assert!(read_shape(&bytes[..length], &mut 100).is_err());
        }
        let duplicate_flags = sp(0, vec![record(2, 0xf00a, &[0; 8])]);
        let malformed_child = record(
            15,
            0xf004,
            &[record(2, 0xf00a, &[0; 8]), vec![1, 2, 3]].concat(),
        );
        for (invalid, work) in [(&duplicate_flags, 20), (&malformed_child, 20), (&bytes, 0)] {
            assert!(read_shape(invalid, &mut work.clone()).is_err());
        }
    }

    #[test]
    fn stretched_shape_picture_fill_references_its_blip_and_keeps_the_line() {
        let tree = |booleans| {
            drawing(vec![sp(
                0x200,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 288])),
                    properties(&[
                        (0x180, 3),
                        (0x182, 32768),
                        (0x4186, 1),
                        (0x1bf, booleans),
                        (0x1c0, 0xff),
                    ]),
                ],
            )])
        };
        let (slide, used) = project_slide(&tree(0x00200020), &[png_blip()], 100).unwrap();
        let shape = shapes(&slide)[0];
        assert_eq!(shape.geometry, "rect");
        assert!(matches!(
            &shape.fill,
            Some(Fill::Image { image_path, stretch: true, rot_with_shape: Some(true), alpha: Some(alpha), .. })
                if image_path == "legacy-ppt/image/1" && *alpha == 0.5
        ));
        assert_eq!(shape.stroke.as_ref().unwrap().color, "FF0000");
        assert_eq!(used, [1]);
        // A vetoed fill neither paints nor admits its BLIP.
        let (slide, used) = project_slide(&tree(0x00100000), &[png_blip()], 100).unwrap();
        assert!(matches!(shapes(&slide)[0].fill, Some(Fill::None)));
        assert!(used.is_empty());
    }

    #[test]
    fn tertiary_fill_rotation_is_scoped_and_supports_explicit_true_and_false() {
        for (tertiary, expected) in [(0x00600060, true), (0x00600040, false)] {
            let tree = drawing(vec![sp(
                0x200,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 288])),
                    properties(&[(0x180, 3), (0x4186, 1), (0x1bf, 0x00100010)]),
                    tertiary_properties(&[(0x1bf, tertiary), (0x180, 99)]),
                ],
            )]);
            let (slide, _) = project_slide(&tree, &[png_blip()], 100).unwrap();
            assert!(matches!(
                shapes(&slide)[0].fill,
                Some(Fill::Image { rot_with_shape: Some(rotate), .. }) if rotate == expected
            ));
        }
    }

    #[test]
    fn tertiary_fill_properties_reject_duplicate_tables_and_boolean_conflicts() {
        let base = vec![
            record(0, 0xf010, &ints(&[0, 0, 576, 288])),
            properties(&[(0x1bf, 0x00200020)]),
        ];
        let mut conflicting = base.clone();
        conflicting.push(tertiary_properties(&[(0x1bf, 0x00200000)]));
        assert!(project(&drawing(vec![sp(0x200, conflicting)])).is_err());

        let mut duplicate = base;
        duplicate.push(tertiary_properties(&[(0x1bf, 0x00200020)]));
        duplicate.push(tertiary_properties(&[(0x1bf, 0x00200020)]));
        assert!(project(&drawing(vec![sp(0x200, duplicate)])).is_err());
    }

    #[test]
    fn shape_picture_fill_validates_its_store_index() {
        let tree = drawing(vec![sp(
            0x200,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 288])),
                properties(&[(0x180, 3), (0x4186, 2)]),
            ],
        )]);
        assert!(project_slide(&tree, &[png_blip()], 100)
            .unwrap_err()
            .contains("index out of range"));
    }

    #[test]
    fn straight_connector_keeps_its_preset_line_and_arrow_without_a_fill() {
        let tree = drawing(vec![sp_kind(
            32,
            0xa00,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                properties(&[(0x181, 255), (0x1c0, 0xff0000), (0x1d1, 5)]),
            ],
        )]);
        let slide = project(&tree).unwrap();
        let shape = shapes(&slide)[0];
        assert_eq!(shape.geometry, "straightConnector1");
        assert!(matches!(shape.fill, Some(Fill::None)));
        let stroke = shape.stroke.as_ref().unwrap();
        assert_eq!(stroke.color, "0000FF");
        assert_eq!(stroke.tail_end.as_ref().unwrap().kind, "arrow");
    }

    #[test]
    fn local_ruler_tabs_reach_each_paragraph_without_a_style_atom() {
        // MS-PPT 2.9.29-30: explicit ruler tabs, signed positions and enum
        // types, with level 0 origins.
        let ruler = record(
            0,
            4006,
            &[
                (4u32 | 8 | 256).to_le_bytes().as_slice(),
                &2u16.to_le_bytes(),
                &576i16.to_le_bytes(),
                &0u16.to_le_bytes(),
                &1152i16.to_le_bytes(),
                &2u16.to_le_bytes(),
                &0i16.to_le_bytes(),
                &0i16.to_le_bytes(),
            ]
            .concat(),
        );
        let textbox = record(
            15,
            0xf00d,
            &[
                record(0, 3999, &4u32.to_le_bytes()),
                record(0, 4008, b"A\tB\rC\tD"),
                ruler,
            ]
            .concat(),
        );
        let shape = sp(
            0xa00,
            vec![record(0, 0xf010, &ints(&[0, 0, 5760, 4320])), textbox],
        );
        let slide = project(&drawing(vec![shape])).unwrap();
        let paragraphs = &shapes(&slide)[0].text_body.as_ref().unwrap().paragraphs;
        assert_eq!(paragraphs.len(), 2);
        for paragraph in paragraphs {
            let tabs: Vec<_> = paragraph
                .tab_stops
                .iter()
                .map(|tab| (tab.pos, tab.algn.as_str()))
                .collect();
            assert_eq!(tabs, [(914400, "l"), (1828800, "r")]);
        }
    }

    #[test]
    fn rejects_ambiguous_local_ruler_ownership_without_guessing_precedence() {
        let ruler = record(
            0,
            4006,
            &[4u32.to_le_bytes().as_slice(), &0u16.to_le_bytes()].concat(),
        );
        for atoms in [
            vec![record(0, 4008, b"A"), ruler.clone(), ruler.clone()],
            vec![record(0, 4008, b"A"), record(0, 4008, b"B"), ruler],
        ] {
            let textbox = record(
                15,
                0xf00d,
                &[record(0, 3999, &4u32.to_le_bytes()), atoms.concat()].concat(),
            );
            let shape = sp(
                0xa00,
                vec![record(0, 0xf010, &ints(&[0, 0, 5760, 4320])), textbox],
            );
            assert!(project(&drawing(vec![shape])).is_err());
        }
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
            let result = project(&drawing(vec![shape]));
            if omitted {
                assert!(result.unwrap().elements.is_empty());
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
        assert!(project(&drawing(vec![hidden_group]))
            .unwrap()
            .elements
            .is_empty());
    }

    #[test]
    fn backgrounds_are_explicit_ungrouped_live_shapes_without_anchor_requirements() {
        let bg = sp(0xc00, vec![properties(&[(0x181, 0x123456)])]);
        let input = drawing(vec![bg.clone()]);
        assert!(matches!(
            background_of(&input, &mut 100).unwrap().unwrap().paint.background_model(None, None),
            Some(Fill::Solid { ref color }) if color == "563412"
        ));
        // Never project the background shape as a foreground rectangle.
        assert!(project(&input).unwrap().elements.is_empty());
        for flag in [0x800, 0xc08, 0xc10] {
            assert!(background_of(&drawing(vec![sp(flag, vec![])]), &mut 100)
                .unwrap()
                .is_none());
        }
        assert!(
            background_of(&drawing(vec![record(15, 0xf003, &bg)]), &mut 100)
                .unwrap()
                .is_none()
        );
        assert!(background_of(&drawing(vec![bg.clone(), bg]), &mut 100)
            .err()
            .unwrap()
            .contains("duplicate"));
        assert!(background_of(&input, &mut 1).is_err());
    }

    #[test]
    fn background_gradient_is_retained_as_a_span_without_eager_decode() {
        let shade = [1, 0, 1, 0, 8, 0, 7, 0, 0, 0, 0, 0, 0, 0];
        let fopt = record(
            (1 << 4) | 3,
            0xf00b,
            &[
                0x8197u16.to_le_bytes().as_slice(),
                (shade.len() as u32).to_le_bytes().as_slice(),
                shade.as_slice(),
            ]
            .concat(),
        );
        let slide = record(15, 1006, &drawing(vec![sp(0xc00, vec![fopt])]));
        let span = record_span_with_end(&slide, 0, &mut 100, "test").unwrap().0;
        let retained = spanned_background(&slide, &span, &mut 100)
            .unwrap()
            .unwrap();
        let moved = slide;
        assert_eq!(
            retained
                .gradient
                .view(&moved)
                .unwrap()
                .decode(&mut 1, &mut 8)
                .unwrap()
                .unwrap()[0]
                .color,
            7
        );
        assert!(retained.gradient.view(&moved[..moved.len() - 1]).is_err());
    }

    #[test]
    // MS-ODRAW 2.3.7.26 leaves fBid undefined for fillShadeColors: ignore
    // that flag, retain scalar zero as reset, and defer invalid values.
    fn scalar_shade_reset_ignores_fbid_and_defers_invalid_value_rejection() {
        let shade = [1, 0, 1, 0, 8, 0, 7, 0, 0, 0, 0, 0, 0, 0];
        for opid in [0x197, 0x4197] {
            let mut props = PropertiesStorage::<&[u8]>::default();
            props
                .apply_primary(0x8197, shade.len() as u32, Some(&shade))
                .unwrap();
            props.apply_primary(opid, 0, None).unwrap();
            assert!(props.gradient.decode(&mut 1, &mut 8).unwrap().is_none());
        }
        let mut invalid = PropertiesStorage::<&[u8]>::default();
        invalid.apply_primary(0x197, 1, None).unwrap();
        assert!(invalid.gradient.decode(&mut 1, &mut 8).is_err());
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
        let props = read_properties(&bytes, &mut 100).unwrap();
        assert_eq!(props.picture, 2);
        assert_eq!(props.crop, [25000, -12500, 50000, 0]);
        let props = read_properties(&properties(&[(0x104, 7)]), &mut 100).unwrap();
        assert_eq!(props.picture, 0);
        assert!(read_properties(&properties(&[(0x100, i32::MAX as u32)]), &mut 100).is_err());
    }

    #[test]
    fn fit_shape_to_text_requires_its_use_bit() {
        let fit = |values: &[(u16, u32)]| {
            read_properties(&properties(values), &mut 100)
                .unwrap()
                .fit_shape_to_text
        };
        assert!(!fit(&[]));
        // Office-saved values: use bits 17-18 with and without the fit bit.
        assert!(fit(&[(0xbf, 0x60002)]));
        assert!(!fit(&[(0xbf, 0x60000)]));
        // Without fUsefFitShapeToText the fit bit is ignored.
        assert!(!fit(&[(0xbf, 0x2)]));
    }

    #[test]
    fn preserves_nontext_geometry_paint_and_stacking_order() {
        let shape = |kind: u16| {
            record(
                15,
                0xf004,
                &[
                    record((kind << 4) | 2, 0xf00a, &ints(&[42, 0xa00])),
                    record(0, 0xf010, &ints(&[144, 288, 864, 720])),
                    properties(&[(0x181, 0x00563412), (0x1c0, 255), (0x1cb, 25400)]),
                ]
                .concat(),
            )
        };
        let slide = project(&drawing(vec![shape(3), shape(1)])).unwrap();
        let shapes = shapes(&slide);
        assert_eq!(
            shapes
                .iter()
                .map(|shape| shape.geometry.as_str())
                .collect::<Vec<_>>(),
            ["ellipse", "rect"]
        );
        for shape in shapes {
            assert!(matches!(shape.fill, Some(Fill::Solid { ref color }) if color == "123456"));
            let stroke = shape.stroke.as_ref().unwrap();
            assert_eq!((stroke.color.as_str(), stroke.width), ("FF0000", 25400));
            assert!(shape.text_body.is_none());
        }
    }

    #[test]
    fn repeated_outline_references_share_the_decoded_text_budget() {
        let referencing = || {
            sp(
                0xa00,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                    record(
                        15,
                        0xf00d,
                        &[record(0, 3998, &ints(&[0])), origin_ruler()].concat(),
                    ),
                ],
            )
        };
        let tree = drawing(vec![referencing(), referencing()]);
        // "Outline" is seven bytes; each reference is charged before copying.
        let (slide, _) = project_slide(&tree, &[], 14).unwrap();
        assert!(shapes(&slide)
            .iter()
            .all(|shape| shape_text(shape) == "Outline"));
        assert!(project_slide(&tree, &[], 13)
            .unwrap_err()
            .contains("decoded text budget"));
    }

    #[test]
    fn does_not_apply_inline_style_to_an_outline_reference() {
        for atoms in [
            [record(0, 4001, &[]), record(0, 3998, &[0; 4])].concat(),
            [record(0, 3998, &[0; 4]), record(0, 4001, &[])].concat(),
        ] {
            let tree = drawing(vec![sp(
                0x200,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                    record(15, 0xf00d, &atoms),
                ],
            )]);
            assert!(project(&tree).is_err());
        }
    }

    #[test]
    fn preserves_signed_rotation_flips_and_emu_text_margins() {
        let tree = drawing(vec![sp(
            0x2c0,
            vec![
                record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                ruled_text("Rotated"),
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
        let slide = project(&tree).unwrap();
        let shape = shapes(&slide)[0];
        assert_eq!(
            (shape.rotation, shape.flip_h, shape.flip_v),
            (-45.0, true, true)
        );
        let body = shape.text_body.as_ref().unwrap();
        assert_eq!(
            (body.l_ins, body.t_ins, body.r_ins, body.b_ins),
            (12700, 25400, 38100, 50800)
        );
        assert_eq!(
            (body.wrap.as_str(), body.vertical_anchor.as_str()),
            ("none", "ctr")
        );
    }

    #[test]
    fn maps_only_owned_text_flow_one_with_default_font_direction_to_ea_vert() {
        let vert = |flags: i32, values: &[(u16, u32)]| {
            let tree = drawing(vec![sp(
                flags,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                    ruled_text("Same text 123"),
                    properties(values),
                ],
            )]);
            let slide = project(&tree).unwrap();
            let vert = shapes(&slide)[0].text_body.as_ref().unwrap().vert.clone();
            vert
        };
        for values in [
            &[][..],
            &[(0x88, 0)],
            &[(0x88, 2)],
            &[(0x88, 3)],
            &[(0x88, 4)],
            &[(0x88, 5)],
        ] {
            assert_eq!(vert(0xa00, values), "horz");
        }
        for flags in [0xa00, 0xa40, 0xa80, 0xac0] {
            assert_eq!(
                vert(flags, &[(4, (-45i32 * 65536) as u32), (0x88, 1)]),
                "eaVert"
            );
        }
        assert_eq!(vert(0xa00, &[(0x88, 1), (0x89, 0)]), "eaVert");
        for direction in 1..=3 {
            assert_eq!(vert(0xa00, &[(0x88, 1), (0x89, direction)]), "horz");
        }
    }

    #[test]
    fn validates_text_direction_scalars_without_guessing_other_values() {
        let tree = |options: Vec<u8>| {
            drawing(vec![sp(
                0xa00,
                vec![
                    record(0, 0xf010, &ints(&[0, 0, 576, 576])),
                    ruled_text("Direction"),
                    options,
                ],
            )])
        };
        for values in [
            vec![(0x88, 1), (0x88, 1)],
            vec![(0x89, 0), (0x89, 0)],
            vec![(0x88, 6)],
            vec![(0x89, 4)],
        ] {
            assert!(project(&tree(properties(&values))).is_err());
        }
        let complex = record((1 << 4) | 3, 0xf00b, &[0x88, 0x80, 0, 0, 0, 0]);
        let blip = properties(&[(0x4088, 1)]);
        assert!(project(&tree(complex)).is_err());
        assert!(project(&tree(blip)).is_err());

        // Direction fields in a tertiary table are outside this measured,
        // primary-owned mapping and remain horizontal rather than reinterpreted.
        let slide = project(&tree(tertiary_properties(&[(0x88, 1)]))).unwrap();
        assert_eq!(shapes(&slide)[0].text_body.as_ref().unwrap().vert, "horz");
    }

    #[test]
    fn validates_complex_property_tails_and_charges_property_work() {
        let malformed = drawing(vec![sp(0x200, vec![properties(&[(0x8380, 100)])])]);
        assert!(project(&malformed)
            .unwrap_err()
            .contains("complex shape property"));
        let opts = properties(&[(4, 0), (0x85, 0)]);
        assert!(read_properties(&opts, &mut 1)
            .err()
            .expect("property work must be charged")
            .contains("work budget"));
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
        assert!(project(&drawing(vec![record(15, 0xf003, &group_header)]))
            .unwrap_err()
            .contains("group coordinate space"));
        assert!(project(&drawing(vec![sp(0, vec![text("missing")])]))
            .unwrap_err()
            .contains("missing PowerPoint shape anchor"));
        let child_space = record(1, 0xf009, &ints(&[0, 0, 576, 576]));
        let mut group = sp(0x202, vec![record(0, 0xf00f, &ints(&[0, 0, 576, 576]))]);
        for _ in 0..=MAX_DEPTH {
            let head = sp(
                0x203,
                vec![
                    record(0, 0xf00f, &ints(&[0, 0, 576, 576])),
                    child_space.clone(),
                ],
            );
            group = record(15, 0xf003, &[head, group].concat());
        }
        let head = sp(
            0x201,
            vec![record(0, 0xf010, &ints(&[0, 0, 576, 576])), child_space],
        );
        let outer = record(15, 0xf003, &[head, group].concat());
        assert!(project(&drawing(vec![outer]))
            .unwrap_err()
            .contains("nesting"));
    }

    #[test]
    fn separate_frames_use_ppt_top_left_order_and_both_anchor_widths() {
        let small: Vec<u8> = [144i16, -288, 576, 432]
            .iter()
            .flat_map(|x| x.to_le_bytes())
            .collect();
        let tree = drawing(vec![
            sp(0x200, vec![record(0, 0xf010, &small)]),
            sp(
                0x200,
                vec![record(0, 0xf010, &ints(&[576, 1152, 1728, 864]))],
            ),
        ]);
        let slide = project(&tree).unwrap();
        let frames: Vec<_> = shapes(&slide)
            .iter()
            .map(|shape| (shape.x, shape.y, shape.width, shape.height))
            .collect();
        assert_eq!(
            frames,
            [
                (-457200, 228600, 1371600, 457200),
                (1828800, 914400, 914400, 457200)
            ]
        );
    }

    #[test]
    fn flattens_group_coordinates_onto_the_slide() {
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
                sp(0x202, vec![record(0, 0xf00f, &ints(&[100, 300, 300, 400]))]),
            ]
            .concat(),
        );
        let slide = project(&drawing(vec![group])).unwrap();
        let shape = shapes(&slide)[0];
        // Group anchor (914400, 457200, 1828800 x 914400) over child space
        // (158750, 317500, 635000 x 635000).
        assert_eq!(
            (shape.x, shape.y, shape.width, shape.height),
            (914400, 685800, 914400, 228600)
        );
    }

    #[test]
    fn skips_deleted_shapes_and_does_not_collect_client_data_text() {
        let anchor = record(0, 0xf010, &ints(&[0, 0, 576, 576]));
        let slide = project(&drawing(vec![
            sp(0x208, vec![anchor.clone(), text("Deleted")]),
            sp(
                0x200,
                vec![
                    anchor,
                    ruled_text("Visible"),
                    record(15, 0xf011, &record(0, 4008, b"Action")),
                ],
            ),
        ]))
        .unwrap();
        let shapes = shapes(&slide);
        assert_eq!(shapes.len(), 1);
        assert_eq!(shape_text(shapes[0]), "Visible");
    }

    #[test]
    fn rejects_truncated_anchors() {
        assert!(project(&drawing(vec![sp(
            0x200,
            vec![record(0, 0xf010, &[0; 7]), text("x")]
        )]))
        .unwrap_err()
        .contains("invalid PowerPoint shape anchor"));
    }
}
