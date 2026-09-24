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
    paint::Paint,
    properties::{self, Property},
    stroke::LineEnd,
    Record,
};
use std::collections::BTreeMap;

pub(in crate::doc) struct Facts {
    pub preset: &'static str,
    /// Solid fill color, `RRGGBB` or `RRGGBBAA`.
    pub fill: Option<String>,
    pub line: Option<Line>,
    pub text: Option<Text>,
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
    pub fn read(
        kind: u16,
        flags: u32,
        shape: Record<'_>,
        extent: [i64; 2],
        budget: &mut usize,
    ) -> Result<Self, String> {
        let preset = preset(kind).ok_or_else(|| {
            unsupported(format!("Word drawing shape type {kind} is not supported"))
        })?;
        // MS-ODRAW 2.2.40 FSP flags: group members, patriarchs, deleted,
        // OLE and master-linked shapes need facts this projection lacks.
        // fConnector is accepted only for the straight connector preset,
        // whose static path is kept without endpoint rerouting.
        if flags & 0x43f != 0 || (flags & 0x100 != 0 && kind != 32) {
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
        table.facts(kind, preset, client_text)
    }
}

/// Primary and tertiary FOPT values. Scalars must agree when repeated;
/// Boolean property sets merge by their use bits (MS-ODRAW 2.3.1).
#[derive(Default)]
struct Table {
    values: BTreeMap<u16, u32>,
}

impl Table {
    fn add(&mut self, property: Property<'_>) -> Result<(), String> {
        let id = property.opid & 0x3fff;
        if property.complex.is_some() {
            // Name, description and the alternate metro XML blob identify or
            // duplicate the shape; they never change its binary rendering.
            // Word sets fBid on these complex strings as well.
            return match id {
                0x380 | 0x381 | 0x3a9 => Ok(()),
                _ => Err(unsupported(format!(
                    "Word drawing shape complex property {id:#06x} is not supported"
                ))),
            };
        }
        if property.opid & 0x4000 != 0 {
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
        preset: &'static str,
        client_text: Option<u32>,
    ) -> Result<Facts, String> {
        let mut paint = Paint::default();
        let mut insets = [0x16530, 0xb298, 0x16530, 0xb298];
        let mut text_id = None;
        for (&id, &value) in &self.values {
            match id {
                // Transform: only an unrotated shape is projected.
                0x4 if value == 0 => {}
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
                0x17f | 0x180..=0x1bf | 0x1c0..=0x1d7 | 0x1ff => paint.property(id, value)?,
                0x23f => {}
                // Black-and-white display modes only affect B/W output.
                0x304..=0x306 => {}
                0x303 if value == 0 => {}
                0x33f => {}
                0x384..=0x387 | 0x388 | 0x38f..=0x392 | 0x3aa | 0x3bf => {}
                0x53f => {}
                0x7c0..=0x7c3 if value == 0 => {}
                // MS-ODRAW 2.3.5 <34>-<37>: Word 2007+ honors relative size
                // and position; their interaction with SPA/fFitShapeToText
                // needs an Office control.
                0x7c0..=0x7c3 => {
                    return Err(unsupported(
                        "Word relative drawing size or position is not supported",
                    ))
                }
                0x7c4 | 0x7c5 => {}
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
            (0x53f, 0, "Word pseudo-inline drawings are not supported"),
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

        let line_shape = is_line(kind);
        let fill = if let Some((color, alpha)) = paint.solid_fill_or_default(!line_shape) {
            Some(rgb(color, alpha)?)
        } else {
            None
        };
        // An enabled but non-solid fill (gradient, pattern, texture,
        // picture or custom fill rectangle) must not become "no fill".
        if !line_shape
            && fill.is_none()
            && paint.filled.unwrap_or(true)
            && paint.fill_ok.unwrap_or(true)
        {
            return Err(unsupported(
                "Word non-solid drawing fills are not supported",
            ));
        }
        let line = match paint.solid_line_or_default(true) {
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
                if paint.lined.unwrap_or(true) && paint.line_ok.unwrap_or(true) {
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
            fill,
            line,
            text,
        })
    }
}

/// OfficeArtCOLORREF (MS-ODRAW 2.2.2) literal colors only. fSystemRGB is an
/// ordinary solid RGB color; palette, scheme and system color indices depend
/// on the rendering host and are rejected.
fn rgb(color: u32, alpha: u32) -> Result<String, String> {
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
        Facts::read(kind, flags, shape, extent, &mut 1000)
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
        assert_eq!(facts.preset, "rect");
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
        assert_eq!(facts.fill.as_deref(), Some("4F81BD80"));
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

    #[test]
    fn explicit_no_fill_no_line_and_line_shapes() {
        let bytes = container(&[(0x1bf, 0x100000), (0x1ff, 0x80008)], &[], &[]);
        let facts = read(1, 0xa00, &bytes, [9000, 3000]).unwrap();
        assert!(facts.fill.is_none() && facts.line.is_some() && facts.text.is_none());
        // Lines keep a zero-width axis and never acquire a fill.
        let bytes = container(&[(0x1d1, 1), (0x1ff, 0x180018)], &[], &[]);
        let facts = read(20, 0xa80, &bytes, [9000, 0]).unwrap();
        assert_eq!(facts.preset, "line");
        assert!(facts.fill.is_none());
        assert_eq!(facts.line.unwrap().ends[1].unwrap().kind, "triangle");
        let facts = read(32, 0xb00, &container(&[], &[], &[]), [0, 900]).unwrap();
        assert_eq!(facts.preset, "straightConnector1");
        assert!(read(20, 0xa00, &container(&[(0x1ff, 0x80000)], &[], &[]), [9, 0]).is_err());
        assert!(read(1, 0xa00, &container(&[], &[], &[]), [9000, 0]).is_err());
    }

    #[test]
    #[allow(clippy::type_complexity)]
    fn unrepresented_display_properties_fail_closed() {
        let rejected: &[(u16, &[(u16, u32)], &[(u16, u32)])] = &[
            // Unsupported shape type and adjusted/custom geometry.
            (51, &[], &[]),
            (1, &[(0x147, 100)], &[]),
            (1, &[(0x4, 0x5a_0000)], &[]),
            // System/palette/scheme colors and non-solid paint.
            (1, &[(0x1c0, 0x1000_0001)], &[]),
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
            (1, &[], &[(0x7c1, 0xc8)]),
            (1, &[], &[(0x53f, 0x10001)]),
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
