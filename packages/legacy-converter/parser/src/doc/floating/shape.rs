//! Direct-model facts of a Word floating drawing shape that is not a picture
//! frame: preset geometry, solid paint and textbox text (MS-ODRAW 2.2.40,
//! 2.3 and 2.4.24; MS-DOC 2.9.106).
//!
//! Every OfficeArt property of the shape is classified. A property that is
//! display-relevant but not represented here rejects the shape with a precise
//! reason instead of being dropped; only editing, identification and
//! black-and-white print-mode properties are ignored.
//!
//! Absent properties take their normative MS-ODRAW defaults. The document's
//! OfficeArtDggContainer.drawingPrimaryOptions are not applied as an
//! inherited layer: the section 2.3.21.15 text description states that absent
//! internal margins use the property defaults, and Word's PDF of a local
//! corpus document places text box text at the 0.1 inch default margin (plus
//! half the line width) although its drawingPrimaryOptions store a smaller
//! margin.

use super::super::{u32_at, unsupported};
use super::records;
use crate::officeart::{
    geometry::{Decoded, DecodedCommand, Geometry},
    paint::Paint,
    properties::{self, Property},
    stroke::LineEnd,
    Record,
};
use docx_model::PathCmd;
use std::collections::BTreeMap;

pub(in crate::doc) struct Facts {
    /// ECMA-376 preset for a preset shape type; `None` for a freeform whose
    /// outline is `subpaths`.
    pub preset: Option<&'static str>,
    /// Freeform outline normalized to the shape box (ECMA-376 20.1.9.8
    /// custGeom), one entry per OfficeArt path.
    pub subpaths: Vec<Vec<PathCmd>>,
    /// Solid fill color, `RRGGBB` or `RRGGBBAA`.
    pub fill: Option<String>,
    pub line: Option<Line>,
    pub text: Option<Text>,
    /// A stretched msofillPicture fill (MS-ODRAW 2.3.7.1-2.3.7.3): the
    /// zero-based drawing-store BLIP index and whether the fill rotates with
    /// the shape (fUseShapeAnchor, 2.3.7.43).
    pub fill_picture: Option<(usize, bool)>,
    /// MS-ODRAW 2.3.18.5 rotation, clockwise about the centre, in degrees.
    /// Only group members may carry a nonzero rotation (see `group`).
    pub rotation: f64,
    /// Relative width and height (MS-ODRAW 2.3.5.1-2.3.5.2, 2.3.5.5-2.3.5.6):
    /// fraction of the named page element; see `relative_size`.
    pub relative_size: [Option<(f64, &'static str)>; 2],
    /// fPseudoInline (MS-ODRAW 2.3.17.11): the shape stands in for an inline
    /// object; see `Store::resolve`.
    pub pseudo_inline: bool,
}

pub(in crate::doc) struct Line {
    /// `RRGGBB` or `RRGGBBAA`.
    pub color: String,
    pub width_emu: u32,
    pub dash: Option<&'static str>,
    /// Canvas cap name (`round`, `square`, `butt`).
    pub cap: &'static str,
    pub join: &'static str,
    pub miter: Option<f64>,
    pub ends: [Option<LineEnd<'static>>; 2],
}

pub(in crate::doc) struct Text {
    /// One-based FTXBXS index from the high word of MSOPSText_lTxid.
    pub index: usize,
    /// dxTextLeft, dyTextTop, dxTextRight, dyTextBottom in EMUs.
    pub insets: [u32; 4],
    /// fFitShapeToText (MS-ODRAW 2.3.21.15), used by Word 2002 and later.
    pub fit_shape: bool,
}

/// MS-ODRAW 2.4.24 shape types whose unadjusted outline is the named
/// ECMA-376 ST_ShapeType preset. Adjusted and custom geometries are rejected
/// before this mapping is consulted.
fn preset(kind: u16) -> Option<&'static str> {
    Paint::default().geometry(kind)
}

fn is_line(kind: u16) -> bool {
    matches!(kind, 20 | 32)
}

impl Facts {
    /// `child` selects a member of an OfficeArt group: it must carry fChild
    /// and its OfficeArtChildAnchor, which the caller has already mapped.
    pub fn read(
        kind: u16,
        flags: u32,
        child: bool,
        shape: Record<'_>,
        extent: [i64; 2],
        budget: &mut usize,
    ) -> Result<Self, String> {
        // msosptNotPrimitive (0) is a freeform: its outline is the explicit
        // OfficeArt path, decoded in `facts`.
        let preset = if kind == 0 {
            None
        } else {
            Some(preset(kind).ok_or_else(|| {
                unsupported(format!("Word drawing shape type {kind} is not supported"))
            })?)
        };
        // MS-ODRAW 2.2.40 FSP flags: group members, patriarchs, deleted,
        // OLE and master-linked shapes need facts this projection lacks.
        // fConnector is accepted only for the straight connector preset,
        // whose static path is kept without endpoint rerouting.
        let membership = if child { 0x2 } else { 0 };
        if flags & 0x43f != membership || (flags & 0x100 != 0 && kind != 32) {
            return Err(unsupported(
                "Word drawing shape has unsupported shape flags",
            ));
        }
        let [width, height] = extent;
        if width < 0 || height < 0 || (width == 0 && height == 0) {
            return Err(unsupported("invalid Word drawing shape extent"));
        }
        if !is_line(kind) && (width == 0 || height == 0) {
            return Err(unsupported("invalid Word drawing shape extent"));
        }

        let mut table = Table::default();
        let mut client_text = None;
        for record in records(shape.payload, budget)? {
            match record.kind {
                // FSP was decoded by the caller; anchors are the SPA's.
                0xf00a | 0xf010 | 0xf011 => {}
                0xf00f if child => {}
                0xf00b => properties::visit(record, budget, |p| table.add(p))?,
                0xf122 => properties::visit_tertiary(record, budget, |p| table.add(p))?,
                0xf00d => {
                    if record.payload.len() != 4 || client_text.is_some() {
                        return Err(unsupported("invalid Word shape textbox reference"));
                    }
                    client_text = Some(u32_at(record.payload, 0)?);
                }
                _ => {
                    return Err(unsupported(format!(
                        "Word drawing shape record {:#06x} is not supported",
                        record.kind
                    )))
                }
            }
        }
        table.facts(kind, preset, client_text, budget)
    }
}

/// Primary and tertiary FOPT values. Scalars must agree when repeated;
/// Boolean property sets merge by their use bits (MS-ODRAW 2.3.1).
#[derive(Default)]
struct Table<'a> {
    values: BTreeMap<u16, u32>,
    /// Geometry arrays (MS-ODRAW 2.3.6.6-2.3.6.9, 2.3.6.18-2.3.6.20).
    complex: BTreeMap<u16, &'a [u8]>,
}

impl<'a> Table<'a> {
    fn add(&mut self, property: Property<'a>) -> Result<(), String> {
        let id = property.opid & 0x3fff;
        if let Some(data) = property.complex {
            // Name, description and the alternate metro XML blob identify or
            // duplicate the shape; they never change its binary rendering.
            // Word sets fBid on these complex strings as well.
            return match id {
                0x380 | 0x381 | 0x3a9 => Ok(()),
                0x145 | 0x146 | 0x151 | 0x152 | 0x155..=0x157 => {
                    if self.complex.insert(id, data).is_some_and(|old| old != data) {
                        return Err(unsupported("conflicting Word drawing geometry"));
                    }
                    self.values.insert(id, property.value);
                    Ok(())
                }
                _ => Err(unsupported(format!(
                    "Word drawing shape complex property {id:#06x} is not supported"
                ))),
            };
        }
        // fillBlip (MS-ODRAW 2.3.7.3) is the only BLIP reference a drawing
        // shape may carry here; the caller resolves it in the drawing store.
        if property.opid & 0x4000 != 0 && id != 0x186 {
            return Err(unsupported(format!(
                "Word drawing shape BLIP property {id:#06x} is not supported"
            )));
        }
        let value = property.value;
        match self.values.get_mut(&id) {
            None => {
                self.values.insert(id, value);
            }
            Some(current) if id & 0x3f == 0x3f => {
                // A Boolean property set: high-word use bits select the valid
                // low-word values. Tables may specify disjoint members; the
                // same member with different values has no precedence rule.
                let (old_use, new_use) = (*current >> 16, value >> 16);
                if (*current ^ value) & old_use & new_use & 0xffff != 0 {
                    return Err(unsupported("conflicting Word drawing Boolean properties"));
                }
                *current = ((old_use | new_use) << 16) | (*current & old_use) | (value & new_use);
            }
            Some(current) if *current != value => {
                return Err(unsupported("conflicting Word drawing shape properties"));
            }
            Some(_) => {}
        }
        Ok(())
    }

    fn boolean(&self, id: u16, bit: u32) -> Option<bool> {
        let value = *self.values.get(&id)?;
        (value & (1 << (bit + 16)) != 0).then_some(value & (1 << bit) != 0)
    }

    fn facts(
        &self,
        kind: u16,
        preset: Option<&'static str>,
        client_text: Option<u32>,
        budget: &mut usize,
    ) -> Result<Facts, String> {
        let mut paint = Paint::default();
        let mut geometry = Geometry::default();
        let mut rotation = 0.0;
        let freeform = kind == 0;
        let mut insets = [0x16530, 0xb298, 0x16530, 0xb298];
        let mut text_id = None;
        for (&id, &value) in &self.values {
            match id {
                // Transform (MS-ODRAW 2.3.18.5): a FixedPoint angle. Callers
                // decide where a rotated shape can be projected.
                0x4 => rotation = f64::from(value as i32) / 65536.0,
                // Protection (2.3.1-2.3.2): editing locks.
                0x40..=0x7f => {}
                0x80 => text_id = Some(value),
                0x81..=0x84 => {
                    if value > 0x132f540 {
                        return Err(unsupported("invalid Word textbox margin"));
                    }
                    insets[usize::from(id - 0x81)] = value;
                }
                // WrapText: square and byPoints wrap; topBottom/through are
                // undefined and MUST be ignored. No-wrap is not represented.
                0x85 if matches!(value, 0 | 1 | 3 | 4) => {}
                // unused134, unused140, unused141.
                0x86 | 0x8c | 0x8d => {}
                // anchorText (2.3.21.8 and <52>), cdirFont (<53>) and txdir
                // (<55>) are used by PowerPoint/Excel only; Word ignores them.
                0x87 | 0x89 | 0x8b => {}
                0x88 if value == 0 => {}
                0x8a if value == 0 => {}
                // hspNext names the next shape of a linked textbox chain.
                0x8a => return Err(unsupported("linked Word textbox chains are not supported")),
                0xbf => {}
                // Picture-only Boolean properties have no effect on shapes.
                0x13f => {}
                // Freeform geometry (MS-ODRAW 2.3.6): the explicit path. Adjust
                // values, guides, handles and connection sites only feed
                // formulas and editing; a path that references a guide or an
                // escape is not decoded and is rejected below.
                0x140..=0x144 if freeform => geometry.scalar(id, value)?,
                0x145 | 0x146 if freeform => match self.complex.get(&id) {
                    Some(data) => geometry.complex(id, data),
                    None => geometry.scalar(id, value)?,
                },
                0x147..=0x150 | 0x151 | 0x155..=0x158 if freeform => {}
                0x152 if freeform && (client_text.is_none() || value == 0) => {}
                0x186 => paint.property(0x4186, value)?,
                0x17f | 0x180..=0x1bf | 0x1c0..=0x1d7 | 0x1ff => paint.property(id, value)?,
                0x23f => {}
                // Black-and-white display modes only affect B/W output.
                0x304..=0x306 => {}
                0x303 if value == 0 => {}
                0x33f => {}
                0x384..=0x387 | 0x388 | 0x38f..=0x392 | 0x3aa | 0x3bf => {}
                0x53f => {}
                0x7c2 | 0x7c3 if value == 0 => {}
                // pctHorizPos/pctVertPos (2.3.5.3-2.3.5.4) have no Word
                // evidence in the corpus yet.
                0x7c2 | 0x7c3 => {
                    return Err(unsupported(
                        "Word relative drawing positions are not supported",
                    ))
                }
                0x7c0 | 0x7c1 | 0x7c4 | 0x7c5 => {}
                _ => {
                    return Err(unsupported(format!(
                        "Word drawing shape property {id:#06x}={value:#x} is not supported"
                    )))
                }
            }
        }
        // Boolean properties whose effects are not represented.
        for (id, bit, reason) in [
            (0xbf, 3, "Word textbox automatic margins are not supported"),
            (0x1ff, 9, "Word opaque line background is not supported"),
            (0x1ff, 6, "Word inset line pens are not supported"),
            (0x1ff, 0, "Word no-line dash rendering is not supported"),
            (0x23f, 1, "Word drawing shadows are not supported"),
            (0x33f, 0, "Word background drawing shapes are not supported"),
            (0x33f, 5, "Word OLE icon shapes are not supported"),
            (0x3bf, 11, "Word horizontal rules are not supported"),
            (
                0x3bf,
                8,
                "Word really-hidden drawing shapes are not supported",
            ),
        ] {
            if self.boolean(id, bit) == Some(true) {
                return Err(unsupported(reason));
            }
        }
        if self.boolean(0x33f, 7).is_some() || self.boolean(0x33f, 6).is_some() {
            return Err(unsupported("Word drawing flip overrides are not supported"));
        }
        if self.values.get(&0x1cd).is_some_and(|value| *value != 0) {
            return Err(unsupported("Word compound drawing lines are not supported"));
        }

        let mut subpaths = Vec::new();
        let mut path_paint = (true, true);
        if freeform {
            let decoded = geometry
                .decode(budget)?
                .ok_or_else(|| unsupported("Word freeform uses guide formulas or path escapes"))?;
            // The DOCX shape model carries one fill and one line for all of
            // its paths, so per-path fill/stroke flags must agree.
            path_paint = decoded.uniform_paint().ok_or_else(|| {
                unsupported("Word freeform paths use differing fill or line flags")
            })?;
            subpaths = normalized(&decoded, budget)?;
        }
        let line_shape = is_line(kind);
        // msofillPicture stretches the picture over the shape (2.3.7.1). A
        // fill rectangle, tiling origin, view-relative sizing or opacity has
        // no representation here and stays rejected with the non-solid fills.
        let fill_picture = if !line_shape && path_paint.0 && paint.fill_type == Some(3) {
            paint
                .foreground_image()
                .filter(|(_, alpha, _)| *alpha == 65536)
                .map(|(blip, _, rotates)| ((blip - 1) as usize, rotates))
        } else {
            None
        };
        let fill = if !path_paint.0 || fill_picture.is_some() {
            None
        } else if let Some((color, alpha)) = paint.solid_fill_or_default(!line_shape) {
            Some(rgb(color, alpha)?)
        } else {
            None
        };
        // An enabled but non-solid fill (gradient, pattern, texture,
        // picture or custom fill rectangle) must not become "no fill".
        if !line_shape
            && path_paint.0
            && fill.is_none()
            && fill_picture.is_none()
            && paint.filled.unwrap_or(true)
            && paint.fill_ok.unwrap_or(true)
        {
            return Err(unsupported(
                "Word non-solid drawing fills are not supported",
            ));
        }
        let line = match paint.solid_line_or_default(path_paint.1) {
            Some((color, alpha)) => {
                let (join, miter) = paint.details.join();
                Some(Line {
                    color: rgb(color, alpha)?,
                    width_emu: paint.width.unwrap_or(9525),
                    dash: paint
                        .dash
                        .and_then(crate::officeart::stroke::preset_dash)
                        .filter(|dash| *dash != "solid"),
                    cap: paint.details.canvas_cap(),
                    join,
                    miter,
                    ends: [paint.details.line_end(0), paint.details.line_end(1)],
                })
            }
            None => {
                if path_paint.1 && paint.lined.unwrap_or(true) && paint.line_ok.unwrap_or(true) {
                    return Err(unsupported(
                        "Word non-solid drawing lines are not supported",
                    ));
                }
                None
            }
        };
        if line_shape && line.is_none() {
            return Err(unsupported("Word line shape has no line"));
        }

        let text = match (text_id.filter(|id| *id != 0), client_text) {
            (None, None) => None,
            (Some(id), client) if client.is_none_or(|client| client == id) => {
                if line_shape {
                    return Err(unsupported("Word line shapes cannot carry text"));
                }
                // MS-DOC 2.9.106: the high word is the one-based FTXBXS index
                // and the low word is zero. A nonzero low word is a later box
                // of a linked chain.
                if id & 0xffff != 0 || id >> 16 == 0 {
                    return Err(unsupported("linked Word textbox chains are not supported"));
                }
                Some(Text {
                    index: (id >> 16) as usize,
                    insets,
                    fit_shape: self.boolean(0xbf, 1) == Some(true),
                })
            }
            _ => return Err(unsupported("inconsistent Word shape textbox reference")),
        };
        Ok(Facts {
            preset,
            subpaths,
            fill_picture,
            rotation,
            relative_size: relative_size(&self.values)?,
            pseudo_inline: self.boolean(0x53f, 0) == Some(true),
            fill,
            line,
            text,
        })
    }
}

/// pctHoriz/pctVert with sizerelh/sizerelv (MS-ODRAW 2.3.5; Word 2007 and
/// later honour them, <34>-<39>). They are the DrawingML `wp14:sizeRelH/V`
/// relative sizes: Word's own DOCX of the local corpus writes pctVert 200
/// (0.1% units) with sizerelv 0 as `sizeRelV relativeFrom="margin"` with
/// `pctHeight 20000` (1/1000 %). The enumeration names match ST_SizeRelFromH
/// and ST_SizeRelFromV; absent sizerelh/sizerelv default to msosrhPage and
/// msosrvPage. A zero percentage leaves the size to the anchor extent.
fn relative_size(values: &BTreeMap<u16, u32>) -> Result<[Option<(f64, &'static str)>; 2], String> {
    let mut result = [None; 2];
    for (axis, (pct, relative, names)) in [
        (
            0x7c0u16,
            0x7c4u16,
            [
                "margin",
                "page",
                "leftMargin",
                "rightMargin",
                "insideMargin",
                "outsideMargin",
            ],
        ),
        (
            0x7c1,
            0x7c5,
            [
                "margin",
                "page",
                "topMargin",
                "bottomMargin",
                "insideMargin",
                "outsideMargin",
            ],
        ),
    ]
    .into_iter()
    .enumerate()
    {
        let value = values.get(&pct).copied().unwrap_or(0);
        if value == 0 {
            continue;
        }
        if value > 10_000 {
            return Err(unsupported("invalid Word relative drawing size"));
        }
        let from = *names
            .get(values.get(&relative).copied().unwrap_or(1) as usize)
            .ok_or_else(|| unsupported("invalid Word relative drawing size origin"))?;
        result[axis] = Some((f64::from(value) / 1000.0, from));
    }
    Ok(result)
}

/// Normalize decoded OfficeArt path coordinates to the unit shape box, the
/// representation of DOCX custom-geometry subpaths.
fn normalized(decoded: &Decoded, budget: &mut usize) -> Result<Vec<Vec<PathCmd>>, String> {
    let width = decoded.width() as f64;
    let height = decoded.height() as f64;
    let mut result = Vec::with_capacity(decoded.paths().len());
    for path in decoded.paths() {
        *budget = budget
            .checked_sub(path.commands().len())
            .ok_or_else(|| unsupported("OfficeArt geometry work budget exceeded"))?;
        result.push(
            path.commands()
                .map(|command| match command {
                    DecodedCommand::Move([x, y]) => PathCmd::MoveTo {
                        x: x as f64 / width,
                        y: y as f64 / height,
                    },
                    DecodedCommand::Line([x, y]) => PathCmd::LineTo {
                        x: x as f64 / width,
                        y: y as f64 / height,
                    },
                    DecodedCommand::Cubic([[x1, y1], [x2, y2], [x, y]]) => PathCmd::CubicBezTo {
                        x1: x1 as f64 / width,
                        y1: y1 as f64 / height,
                        x2: x2 as f64 / width,
                        y2: y2 as f64 / height,
                        x: x as f64 / width,
                        y: y as f64 / height,
                    },
                    DecodedCommand::Close => PathCmd::Close,
                })
                .collect(),
        );
    }
    Ok(result)
}

/// OfficeArtCOLORREF (MS-ODRAW 2.2.2). fSystemRGB is an ordinary solid RGB
/// color. Of the fSysIndex system colors only the two with Word evidence are
/// resolved: Word's own DOCX of the local corpus drawings (from which Word
/// saved the DOC files) writes every index 0x0001 as `sysClr windowText`
/// with lastClr 000000 and every index 0x0011 as `sysClr window` with
/// lastClr FFFFFF (27 and 16 properties), and Word's PDFs draw them black
/// and white. These indices are not GetSysColor numbers, so no other index
/// is inferred; procedural modifiers in the blue byte, palette and scheme
/// colors stay rejected.
fn rgb(color: u32, alpha: u32) -> Result<String, String> {
    let color = match color {
        0x1000_0001 => 0x0000_0000,
        0x1000_0011 => 0x00ff_ffff,
        _ => color,
    };
    if !matches!(color & 0xff00_0000, 0 | 0x0400_0000) {
        return Err(unsupported(
            "Word drawing colors other than literal RGB are not supported",
        ));
    }
    let mut value = format!(
        "{:02X}{:02X}{:02X}",
        color & 255,
        (color >> 8) & 255,
        (color >> 16) & 255
    );
    if alpha != 65536 {
        // FixedPoint 16.16 opacity, validated to [0, 1] by `Paint`.
        value.push_str(&format!("{:02X}", (u64::from(alpha) * 255 + 32768) / 65536));
    }
    Ok(value)
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

    fn table(kind: u16, properties: &[(u16, u32)], complex: &[u8]) -> Vec<u8> {
        let mut body = Vec::new();
        for (key, value) in properties {
            body.extend(key.to_le_bytes());
            body.extend(value.to_le_bytes());
        }
        body.extend(complex);
        record(kind, ((properties.len() as u16) << 4) | 3, &body)
    }

    fn container(primary: &[(u16, u32)], tertiary: &[(u16, u32)], extra: &[u8]) -> Vec<u8> {
        record(
            0xf004,
            15,
            &[
                table(0xf00b, primary, &[]),
                table(0xf122, tertiary, &[]),
                extra.to_vec(),
            ]
            .concat(),
        )
    }

    fn read(kind: u16, flags: u32, bytes: &[u8], extent: [i64; 2]) -> Result<Facts, String> {
        let (shape, _) = crate::officeart::record_with_end(bytes, 0, &mut 1000, "test").unwrap();
        Facts::read(kind, flags, false, shape, extent, &mut 1000)
    }

    #[test]
    fn absent_paint_and_margins_use_the_normative_property_defaults() {
        // A default Word textbox: fFilled/fLine use bits only, lTxid 1.
        let bytes = container(
            &[
                (0x80, 0x10000),
                (0xbf, 0x60000),
                (0x1bf, 0x100010),
                (0x1ff, 0x80008),
            ],
            &[(0x1ff, 0x400000), (0x3bf, 0x8200_8200)],
            &record(0xf00d, 0, &0x10000u32.to_le_bytes()),
        );
        let facts = read(202, 0xa00, &bytes, [9000, 3000]).unwrap();
        assert_eq!(facts.preset, Some("rect"));
        assert_eq!(facts.fill.as_deref(), Some("FFFFFF"));
        let line = facts.line.unwrap();
        assert_eq!(line.color, "000000");
        assert_eq!(line.width_emu, 9525);
        assert_eq!((line.cap, line.join, line.dash), ("butt", "round", None));
        let text = facts.text.unwrap();
        assert_eq!(text.index, 1);
        assert_eq!(text.insets, [91_440, 45_720, 91_440, 45_720]);
        assert!(!text.fit_shape);
    }

    #[test]
    fn explicit_solid_paint_margins_and_fit_shape_are_preserved() {
        let bytes = container(
            &[
                (0x80, 0x20000),
                (0x81, 0),
                (0x84, 12_700),
                (0xbf, 0x60002),
                (0x181, 0x00bd_814f),
                (0x182, 0x8000),
                (0x1c0, 0x0400_1020),
                (0x1cb, 25_400),
                (0x1ce, 6),
                (0x1d6, 1),
            ],
            &[],
            &[],
        );
        let facts = read(1, 0xa00, &bytes, [9000, 3000]).unwrap();
        assert_eq!(facts.rotation, 0.0);
        assert_eq!(facts.fill.as_deref(), Some("4F81BD80"));
        // The FixedPoint angle is reported; callers decide where it is valid.
        let rotated = container(&[(0x4, 0xffd6_4e6d)], &[], &[]);
        let turned = read(1, 0xa00, &rotated, [9, 9]).unwrap();
        assert!((turned.rotation + 41.694).abs() < 1e-3);
        let line = facts.line.unwrap();
        assert_eq!(line.color, "201000");
        assert_eq!(line.width_emu, 25_400);
        assert_eq!(line.dash, Some("dash"));
        assert_eq!(line.join, "miter");
        let text = facts.text.unwrap();
        assert_eq!(text.index, 2);
        assert_eq!(text.insets, [0, 45_720, 91_440, 12_700]);
        assert!(text.fit_shape);
    }

    /// A freeform FOPT: bounds, vertices and optional segments as complex
    /// arrays (MS-ODRAW 2.3.6.1-2.3.6.9).
    fn freeform(points: &[[i32; 2]], segments: &[u16], extra: &[(u16, u32)]) -> Vec<u8> {
        let array = |count: usize, size: u16, body: Vec<u8>| {
            [
                (count as u16).to_le_bytes().as_slice(),
                &(count as u16).to_le_bytes(),
                &size.to_le_bytes(),
                &body,
            ]
            .concat()
        };
        let vertices = array(
            points.len(),
            8,
            points
                .iter()
                .flatten()
                .flat_map(|v| v.to_le_bytes())
                .collect(),
        );
        let segments = array(
            segments.len(),
            2,
            segments.iter().flat_map(|v| v.to_le_bytes()).collect(),
        );
        let mut properties: Vec<(u16, u32)> = vec![(0x142, 100), (0x143, 50), (0x144, 4)];
        properties.extend_from_slice(extra);
        let mut body = Vec::new();
        for (key, value) in &properties {
            body.extend(key.to_le_bytes());
            body.extend(value.to_le_bytes());
        }
        for (key, data) in [(0xc145u16, &vertices), (0xc146, &segments)] {
            body.extend(key.to_le_bytes());
            body.extend((data.len() as u32).to_le_bytes());
        }
        body.extend(&vertices);
        body.extend(&segments);
        let count = properties.len() + 2;
        record(
            0xf004,
            15,
            &record(0xf00b, ((count as u16) << 4) | 3, &body),
        )
    }

    #[test]
    fn freeform_paths_normalize_to_the_shape_box() {
        // Move, two lines, a cubic, close, end: a filled and stroked outline.
        let points = [[0, 0], [100, 0], [100, 50], [50, 50], [25, 50], [0, 25]];
        let segments = [0x4000, 0x0002, 0x2001, 0x6001, 0x8000];
        let facts = read(0, 0xa00, &freeform(&points, &segments, &[]), [9000, 3000]).unwrap();
        assert_eq!(facts.preset, None);
        assert_eq!(facts.fill.as_deref(), Some("FFFFFF"));
        assert!(facts.line.is_some());
        let [path] = facts.subpaths.as_slice() else {
            panic!("one path")
        };
        assert!(matches!(path[0], PathCmd::MoveTo { x, y } if x == 0.0 && y == 0.0));
        assert!(matches!(path[2], PathCmd::LineTo { x, y } if x == 1.0 && y == 1.0));
        assert!(matches!(
            path[3],
            PathCmd::CubicBezTo { x1, y2, x, y, .. }
                if x1 == 0.5 && y2 == 1.0 && x == 0.0 && y == 0.5
        ));
        assert!(matches!(path[4], PathCmd::Close));
        // An open path is never filled; a uniform noFill escape removes fill.
        let open = [0x4000, 0x0002, 0x8000];
        let facts = read(0, 0xa00, &freeform(&points[..3], &open, &[]), [9, 9]).unwrap();
        assert!(facts.fill.is_none() && facts.line.is_some());
        let no_fill = [0x4000, 0xaa00, 0x0002, 0x8000];
        let facts = read(0, 0xa00, &freeform(&points[..3], &no_fill, &[]), [9, 9]).unwrap();
        assert!(facts.fill.is_none());
        // Adjust values and guides only feed formulas that the path does not use.
        let adjusted = freeform(&points[..3], &open, &[(0x147, 5000)]);
        assert!(read(0, 0xa00, &adjusted, [9, 9]).is_ok());
    }

    #[test]
    fn freeform_paint_and_formula_gaps_fail_closed() {
        let points = [[0, 0], [100, 0], [100, 50], [0, 50]];
        // Differing per-path fill flags cannot be expressed by one shape.
        let mixed = [
            0x4000, 0x0001, 0x6001, 0x8000, 0xaa00, 0x4000, 0x0001, 0x6001, 0x8000,
        ];
        assert!(read(0, 0xa00, &freeform(&points, &mixed, &[]), [9, 9])
            .err()
            .is_some_and(|error| error.contains("differing")));
        // Guide-referencing vertices and arc escapes are not decoded.
        let guided = [[0, 0], [i32::MIN, 0]];
        assert!(read(
            0,
            0xa00,
            &freeform(&guided, &[0x4000, 0x0001, 0x8000], &[]),
            [9, 9]
        )
        .is_err());
        let arc = [0x4000, 0xa301, 0x8000];
        assert!(read(0, 0xa00, &freeform(&points, &arc, &[]), [9, 9]).is_err());
        // Geometry properties stay rejected on preset shapes.
        assert!(read(
            1,
            0xa00,
            &freeform(&points, &[0x4000, 0x0003, 0x8000], &[]),
            [9, 9]
        )
        .is_err());
    }

    #[test]
    fn pseudo_inline_shapes_are_reported_to_the_caller() {
        let bytes = container(&[], &[(0x53f, 0x0001_0001), (0x390, 3), (0x392, 3)], &[]);
        assert!(read(1, 0xa00, &bytes, [9, 9]).unwrap().pseudo_inline);
        let bytes = container(&[], &[(0x53f, 0x0001_0000)], &[]);
        assert!(!read(1, 0xa00, &bytes, [9, 9]).unwrap().pseudo_inline);
    }

    #[test]
    fn relative_sizes_follow_word_drawingml_relative_size() {
        let bytes = container(&[], &[(0x7c1, 200), (0x7c5, 0), (0x7c4, 0)], &[]);
        let facts = read(202, 0xa00, &bytes, [9, 9]).unwrap();
        assert_eq!(facts.relative_size, [None, Some((0.2, "margin"))]);
        // Absent origins default to the page; zero percentages use the extent.
        let bytes = container(&[], &[(0x7c0, 1000), (0x7c1, 0)], &[]);
        let facts = read(202, 0xa00, &bytes, [9, 9]).unwrap();
        assert_eq!(facts.relative_size, [Some((1.0, "page")), None]);
    }

    #[test]
    fn word_window_system_colors_resolve_to_their_observed_values() {
        let bytes = container(&[(0x181, 0x1000_0011), (0x1c0, 0x1000_0001)], &[], &[]);
        let facts = read(1, 0xa00, &bytes, [9, 9]).unwrap();
        assert_eq!(facts.fill.as_deref(), Some("FFFFFF"));
        assert_eq!(facts.line.unwrap().color, "000000");
    }

    #[test]
    fn stretched_picture_fills_reference_the_drawing_store() {
        let fill = |extra: &[(u16, u32)]| {
            let mut properties = vec![(0x180u16, 3u32), (0x4186, 1)];
            properties.extend_from_slice(extra);
            read(1, 0xa00, &container(&properties, &[], &[]), [9, 9])
        };
        let facts = fill(&[]).unwrap();
        assert_eq!(facts.fill_picture, Some((0, false)));
        assert!(facts.fill.is_none());
        // fUseShapeAnchor (use bit 21, value bit 5) rotates the fill.
        assert_eq!(
            fill(&[(0x1bf, 0x0060_0060)]).unwrap().fill_picture,
            Some((0, true))
        );
        // Opacity, fill rectangles, tiling origins and missing BLIPs fail.
        for extra in [
            (0x182u16, 0x8000u32),
            (0x1bf, 0x0002_0002),
            (0x198, 1),
            (0x195, 1),
        ] {
            assert!(fill(&[extra]).is_err(), "{extra:x?}");
        }
        assert!(read(1, 0xa00, &container(&[(0x180, 3)], &[], &[]), [9, 9]).is_err());
        // Other BLIP-valued properties stay rejected.
        assert!(read(1, 0xa00, &container(&[(0x4104, 1)], &[], &[]), [9, 9]).is_err());
    }

    #[test]
    fn explicit_no_fill_no_line_and_line_shapes() {
        let bytes = container(&[(0x1bf, 0x100000), (0x1ff, 0x80008)], &[], &[]);
        let facts = read(1, 0xa00, &bytes, [9000, 3000]).unwrap();
        assert!(facts.fill.is_none() && facts.line.is_some() && facts.text.is_none());
        // Lines keep a zero-width axis and never acquire a fill.
        let bytes = container(&[(0x1d1, 1), (0x1ff, 0x180018)], &[], &[]);
        let facts = read(20, 0xa80, &bytes, [9000, 0]).unwrap();
        assert_eq!(facts.preset, Some("line"));
        assert!(facts.fill.is_none());
        assert_eq!(facts.line.unwrap().ends[1].unwrap().kind, "triangle");
        let facts = read(32, 0xb00, &container(&[], &[], &[]), [0, 900]).unwrap();
        assert_eq!(facts.preset, Some("straightConnector1"));
        assert!(read(20, 0xa00, &container(&[(0x1ff, 0x80000)], &[], &[]), [9, 0]).is_err());
        assert!(read(1, 0xa00, &container(&[], &[], &[]), [9000, 0]).is_err());
    }

    #[test]
    #[allow(clippy::type_complexity)]
    fn unrepresented_display_properties_fail_closed() {
        let rejected: &[(u16, &[(u16, u32)], &[(u16, u32)])] = &[
            // Unsupported shape type, adjusted preset geometry and a freeform
            // without a decodable path.
            (51, &[], &[]),
            (0, &[], &[]),
            (0, &[(0x144, 4)], &[]),
            (1, &[(0x147, 100)], &[]),
            // System/palette/scheme colors and non-solid paint.
            (1, &[(0x1c0, 0x1000_0002)], &[]),
            (1, &[(0x1c0, 0x1001_0011)], &[]),
            (1, &[(0x181, 0x0800_0001)], &[]),
            (1, &[(0x180, 4)], &[]),
            (1, &[(0x1c4, 1)], &[]),
            (1, &[(0x1cd, 1)], &[]),
            // Shadows, auto margins, no-wrap text, flow and linked chains.
            (1, &[(0x23f, 0x20002)], &[]),
            (1, &[(0xbf, 0x80008)], &[]),
            (1, &[(0x85, 2)], &[]),
            (1, &[(0x88, 1)], &[]),
            (1, &[(0x80, 0x10001)], &[]),
            (1, &[(0x80, 0x10000), (0x8a, 0x806)], &[]),
            // Relative sizing, pseudo-inline, horizontal rule, unknown ids.
            (1, &[], &[(0x7c3, 0xc8)]),
            (1, &[], &[(0x7c1, 10_001)]),
            (1, &[], &[(0x7c1, 5), (0x7c5, 6)]),
            (1, &[(0x3bf, 0x0800_0800)], &[]),
            (1, &[(0x2ff, 0)], &[]),
            // Conflicting scalars and Boolean members between tables.
            (1, &[(0x181, 1)], &[(0x181, 2)]),
            (1, &[(0x1ff, 0x80008)], &[(0x1ff, 0x80000)]),
        ];
        for (kind, primary, tertiary) in rejected {
            let bytes = container(primary, tertiary, &[]);
            assert!(
                read(*kind, 0xa00, &bytes, [9000, 3000])
                    .err()
                    .is_some_and(|error| error.starts_with("UNSUPPORTED:")),
                "{kind} {primary:x?} {tertiary:x?}"
            );
        }
        // Group members, OLE, master-linked and non-straight connectors.
        for flags in [0xa02, 0xa10, 0xa20, 0xb00] {
            assert!(read(1, flags, &container(&[], &[], &[]), [9, 9]).is_err());
        }
        // A textbox reference must agree with lTxid.
        let mismatch = container(
            &[(0x80, 0x10000)],
            &[],
            &record(0xf00d, 0, &0x20000u32.to_le_bytes()),
        );
        assert!(read(202, 0xa00, &mismatch, [9, 9]).is_err());
        // Group-relative anchors belong to nested groups.
        let child = container(&[], &[], &record(0xf00f, 0, &[0; 16]));
        assert!(read(1, 0xa00, &child, [9, 9]).is_err());
    }

    #[test]
    fn editing_identification_and_print_mode_properties_are_ignored() {
        let name = [0x61u8, 0, 0, 0];
        let mut body = Vec::new();
        for (key, value) in [
            (0x7fu16, 0x01e1_0080u32),
            (0x304, 9),
            (0x33f, 0x0010_0010),
            (0x388, 0),
            (0x3aa, 5),
            (0x7c4, 0),
            (0xc380, name.len() as u32),
        ] {
            body.extend(key.to_le_bytes());
            body.extend(value.to_le_bytes());
        }
        body.extend(name);
        let bytes = record(0xf004, 15, &record(0xf00b, (7 << 4) | 3, &body));
        assert!(read(1, 0xa00, &bytes, [9, 9]).is_ok());
    }
}
