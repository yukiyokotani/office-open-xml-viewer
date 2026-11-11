//! StarView metafiles (`VCLMTF`, LibreOffice's native vector picture format,
//! stored as `Pictures/*.svm` and as object replacement pictures) converted
//! to SVG, which the DOCX and PPTX renderers draw natively. Shared by the
//! ODT and ODP readers.
//!
//! The record layout follows LibreOffice's `SvmReader`: a header (map mode,
//! preferred size, action count) and actions of `u16 type`, a version/length
//! compat header and a payload. The converter draws the geometry and text
//! actions LibreOffice writes for vector pictures (map modes including
//! relative ones, push/pop, line/fill/text colours, fonts, text alignment,
//! text with character positions, polylines, polygons with Bézier flags,
//! poly-polygons, rectangles, ellipses, lines and transparent polygons).
//! Raster, gradient, hatch and clip actions are skipped (a bounded,
//! documented subset of the format; every action is length-delimited, so a
//! skipped action never desynchronizes the stream). EMF+ comment records
//! that LibreOffice keeps beside the converted actions are not replayed.

use std::fmt::Write as _;

/// Implementation bounds.
const MAX_ACTIONS: u32 = 2_000_000;
const MAX_SVG_BYTES: usize = 64 * 1024 * 1024;
const MAX_STACK: usize = 1024;

struct Reader<'a> {
    b: &'a [u8],
    p: usize,
}

/// A polygon's points and optional point flags.
type Polygon = (Vec<(f64, f64)>, Option<Vec<u8>>);

impl<'a> Reader<'a> {
    fn u8(&mut self) -> Option<u8> {
        let v = *self.b.get(self.p)?;
        self.p += 1;
        Some(v)
    }
    fn u16(&mut self) -> Option<u16> {
        let v = u16::from_le_bytes(self.b.get(self.p..self.p + 2)?.try_into().ok()?);
        self.p += 2;
        Some(v)
    }
    fn u32(&mut self) -> Option<u32> {
        let v = u32::from_le_bytes(self.b.get(self.p..self.p + 4)?.try_into().ok()?);
        self.p += 4;
        Some(v)
    }
    fn i32(&mut self) -> Option<i32> {
        self.u32().map(|v| v as i32)
    }
    fn point(&mut self) -> Option<(f64, f64)> {
        Some((self.i32()? as f64, self.i32()? as f64))
    }
    fn compat(&mut self) -> Option<(u16, usize)> {
        let version = self.u16()?;
        let len = self.u32()? as usize;
        let end = self.p.checked_add(len)?;
        (end <= self.b.len()).then_some((version, end))
    }
    /// A nested compat record: a reader confined to its declared payload,
    /// with this reader moved past it (never backwards, so a record cannot
    /// be read twice).
    fn nested(&mut self) -> Option<(u16, Reader<'a>)> {
        let (version, end) = self.compat()?;
        let inner = Reader {
            b: &self.b[..end],
            p: self.p,
        };
        self.p = end;
        Some((version, inner))
    }
    /// A colour: 0xTTRRGGBB (transparency in the high byte).
    fn color(&mut self) -> Option<u32> {
        self.u32()
    }
    /// A string in the stream charset: UTF-16 with a u32 length for the
    /// Unicode charset, otherwise u16 length + bytes (read as Latin-1).
    fn string(&mut self, unicode: bool) -> Option<String> {
        if unicode {
            let n = self.u32()? as usize;
            self.utf16(n)
        } else {
            let n = self.u16()? as usize;
            let bytes = self.b.get(self.p..self.p.checked_add(n)?)?;
            self.p += n;
            Some(bytes.iter().map(|&b| b as char).collect())
        }
    }
    fn utf16(&mut self, n: usize) -> Option<String> {
        let bytes = self.b.get(self.p..self.p.checked_add(n.checked_mul(2)?)?)?;
        self.p += n * 2;
        let units: Vec<u16> = bytes
            .chunks(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        Some(String::from_utf16_lossy(&units))
    }
    fn polygon(&mut self) -> Option<Vec<(f64, f64)>> {
        let n = self.u16()? as usize;
        let mut out = Vec::with_capacity(n.min(65_536));
        for _ in 0..n {
            out.push(self.point()?);
        }
        Some(out)
    }
    /// `tools::Polygon::Read`: compat, points, and optional point flags.
    fn polygon_with_flags(&mut self) -> Option<Polygon> {
        let (_, mut r) = self.nested()?;
        let points = r.polygon()?;
        let flags = if r.u8()? != 0 {
            Some(r.b.get(r.p..r.p.checked_add(points.len())?)?.to_vec())
        } else {
            None
        };
        Some((points, flags))
    }
}

#[derive(Clone)]
struct MapMode {
    /// Unit size in 1/100 mm.
    unit: f64,
    origin: (f64, f64),
    scale: (f64, f64),
}

fn unit_hmm(unit: u16) -> f64 {
    match unit {
        0 => 1.0,
        1 => 10.0,
        2 => 100.0,
        3 => 1000.0,
        4 => 2.54,
        5 => 25.4,
        6 => 254.0,
        7 => 2540.0,
        8 => 2540.0 / 72.0,
        9 => 2540.0 / 1440.0,
        _ => 2540.0 / 96.0,
    }
}

fn read_map_mode(r: &mut Reader<'_>, current: Option<&MapMode>) -> Option<MapMode> {
    let (_, mut inner) = r.nested()?;
    let r = &mut inner;
    let unit = r.u16()?;
    let origin = r.point()?;
    let fraction = |r: &mut Reader<'_>| -> Option<f64> {
        let n = r.i32()? as f64;
        let d = r.i32()? as f64;
        Some(if d == 0.0 { 1.0 } else { n / d })
    };
    let sx = fraction(r)?;
    let sy = fraction(r)?;
    // MapUnit::MapRelative (13) moves the origin and multiplies the scale of
    // the current map mode.
    if unit == 13 {
        let base = current.cloned().unwrap_or(MapMode {
            unit: 1.0,
            origin: (0.0, 0.0),
            scale: (1.0, 1.0),
        });
        return Some(MapMode {
            unit: base.unit,
            origin: (base.origin.0 + origin.0, base.origin.1 + origin.1),
            scale: (base.scale.0 * sx, base.scale.1 * sy),
        });
    }
    Some(MapMode {
        unit: unit_hmm(unit),
        origin,
        scale: (sx, sy),
    })
}

#[derive(Clone)]
struct Font {
    /// Shared: pushed states copy it (`MAX_STACK` deep).
    family: std::rc::Rc<str>,
    height: f64,
    weight: u16,
    italic: bool,
    underline: bool,
    strike: bool,
    orientation: f64,
    unicode: bool,
}

#[derive(Clone)]
struct State {
    map: MapMode,
    line: Option<u32>,
    fill: Option<u32>,
    text: u32,
    align: u16,
    font: Font,
}

fn css(color: u32) -> String {
    format!("#{:06x}", color & 0xFF_FFFF)
}

fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for c in text.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            c if (c as u32) < 0x20 => {}
            c => out.push(c),
        }
    }
    out
}

/// The natural size of a metafile in EMU (its preferred size).
pub struct Converted {
    pub svg: Vec<u8>,
    #[allow(dead_code)] // the ODP reader sizes pictures by it
    pub size_emu: (f64, f64),
}

pub fn to_svg(bytes: &[u8]) -> Option<Converted> {
    if !bytes.starts_with(b"VCLMTF") {
        return None;
    }
    let mut r = Reader { b: bytes, p: 6 };
    let (_, header_end) = r.compat()?;
    let _compression = r.u32()?;
    let map = read_map_mode(&mut r, None)?;
    let (pw, ph) = r.point()?;
    let count = r.u32()?;
    r.p = header_end;
    if count > MAX_ACTIONS {
        return None;
    }
    let header_map = map.clone();
    // Output coordinates: the header map mode's logical units.
    let to_out = |m: &MapMode, x: f64, y: f64| -> (f64, f64) {
        let k = m.unit / header_map.unit;
        let hx = header_map.scale.0.max(1e-12);
        let hy = header_map.scale.1.max(1e-12);
        (
            ((x + m.origin.0) * m.scale.0 * k) / hx - header_map.origin.0,
            ((y + m.origin.1) * m.scale.1 * k) / hy - header_map.origin.1,
        )
    };
    let mut state = State {
        map,
        line: Some(0),
        fill: Some(0xFF_FFFF),
        text: 0,
        align: 0,
        font: Font {
            family: std::rc::Rc::from(""),
            height: 0.0,
            weight: 5,
            italic: false,
            underline: false,
            strike: false,
            orientation: 0.0,
            unicode: false,
        },
    };
    let mut stack: Vec<State> = Vec::new();
    let mut body = String::new();
    let width = pw.abs().max(1.0);
    let height = ph.abs().max(1.0);
    for _ in 0..count {
        let Some(kind) = r.u16() else { break };
        let Some((version, end)) = r.compat() else {
            break;
        };
        let mut a = Reader {
            b: &bytes[..end],
            p: r.p,
        };
        let map = state.map.clone();
        let pt = |p: (f64, f64)| to_out(&map, p.0, p.1);
        let len_scale = |v: f64| {
            let (x0, _) = to_out(&map, 0.0, 0.0);
            let (x1, _) = to_out(&map, v, 0.0);
            (x1 - x0).abs()
        };
        let paint = |s: &State, fill: bool, stroke: bool, width: f64| -> String {
            let mut out = String::new();
            match (fill, s.fill) {
                (true, Some(c)) => {
                    let _ = write!(out, r#" fill="{}""#, css(c));
                }
                _ => out.push_str(r#" fill="none""#),
            }
            if let (true, Some(c)) = (stroke, s.line) {
                let _ = write!(
                    out,
                    r#" stroke="{}" stroke-width="{:.2}""#,
                    css(c),
                    width.max(len_scale(1.0)).max(width)
                );
            }
            out
        };
        // Path data written straight into the output, within its bound
        // (checked as each command is written; `None` past it).
        let path_data = |d: &mut String,
                         points: &[(f64, f64)],
                         flags: Option<&[u8]>,
                         close: bool|
         -> Option<()> {
            let mut i = 0;
            while i < points.len() {
                let (x, y) = pt(points[i]);
                let control = |k: usize| flags.and_then(|f| f.get(k)).copied() == Some(2);
                if i == 0 {
                    let _ = write!(d, "M{x:.2} {y:.2}");
                    i += 1;
                } else if control(i) && i + 2 < points.len() {
                    let (x1, y1) = pt(points[i]);
                    let (x2, y2) = pt(points[i + 1]);
                    let (x3, y3) = pt(points[i + 2]);
                    let _ = write!(d, "C{x1:.2} {y1:.2} {x2:.2} {y2:.2} {x3:.2} {y3:.2}");
                    i += 3;
                } else {
                    let _ = write!(d, "L{x:.2} {y:.2}");
                    i += 1;
                }
                if d.len() > MAX_SVG_BYTES {
                    return None;
                }
            }
            if close && !points.is_empty() {
                d.push('Z');
            }
            Some(())
        };
        match kind {
            // LINE
            102 => {
                if let (Some(p0), Some(p1)) = (a.point(), a.point()) {
                    let (x0, y0) = pt(p0);
                    let (x1, y1) = pt(p1);
                    let w = if version >= 2 {
                        line_width(&mut a).map(len_scale).unwrap_or(0.0)
                    } else {
                        0.0
                    };
                    let _ = write!(
                        body,
                        r#"<path d="M{x0:.2} {y0:.2}L{x1:.2} {y1:.2}"{}/>"#,
                        paint(&state, false, true, w)
                    );
                }
            }
            // RECT
            103 => {
                if let (Some(p0), Some(p1)) = (a.point(), a.point()) {
                    let (x0, y0) = pt(p0);
                    let (x1, y1) = pt(p1);
                    let _ = write!(
                        body,
                        r#"<rect x="{:.2}" y="{:.2}" width="{:.2}" height="{:.2}"{}/>"#,
                        x0.min(x1),
                        y0.min(y1),
                        (x1 - x0).abs(),
                        (y1 - y0).abs(),
                        paint(&state, true, true, 0.0)
                    );
                }
            }
            // ROUNDRECT
            104 => {
                if let (Some(p0), Some(p1), Some(rx), Some(ry)) =
                    (a.point(), a.point(), a.u32(), a.u32())
                {
                    let (x0, y0) = pt(p0);
                    let (x1, y1) = pt(p1);
                    let _ = write!(
                        body,
                        r#"<rect x="{:.2}" y="{:.2}" width="{:.2}" height="{:.2}" rx="{:.2}" ry="{:.2}"{}/>"#,
                        x0.min(x1),
                        y0.min(y1),
                        (x1 - x0).abs(),
                        (y1 - y0).abs(),
                        len_scale(rx as f64) / 2.0,
                        len_scale(ry as f64) / 2.0,
                        paint(&state, true, true, 0.0)
                    );
                }
            }
            // ELLIPSE
            105 => {
                if let (Some(p0), Some(p1)) = (a.point(), a.point()) {
                    let (x0, y0) = pt(p0);
                    let (x1, y1) = pt(p1);
                    let _ = write!(
                        body,
                        r#"<ellipse cx="{:.2}" cy="{:.2}" rx="{:.2}" ry="{:.2}"{}/>"#,
                        (x0 + x1) / 2.0,
                        (y0 + y1) / 2.0,
                        (x1 - x0).abs() / 2.0,
                        (y1 - y0).abs() / 2.0,
                        paint(&state, true, true, 0.0)
                    );
                }
            }
            // POLYLINE
            109 => {
                if let Some(mut points) = a.polygon() {
                    let w = if version >= 2 {
                        line_width(&mut a).map(len_scale).unwrap_or(0.0)
                    } else {
                        0.0
                    };
                    let mut flags = None;
                    if version >= 3 && a.u8() == Some(1) {
                        if let Some((p, f)) = a.polygon_with_flags() {
                            points = p;
                            flags = f;
                        }
                    }
                    body.push_str(r#"<path d=""#);
                    path_data(&mut body, &points, flags.as_deref(), false)?;
                    let _ = write!(body, r#""{}/>"#, paint(&state, false, true, w));
                }
            }
            // POLYGON
            110 => {
                if let Some(mut points) = a.polygon() {
                    let mut flags = None;
                    if version >= 2 && a.u8() == Some(1) {
                        if let Some((p, f)) = a.polygon_with_flags() {
                            points = p;
                            flags = f;
                        }
                    }
                    body.push_str(r#"<path d=""#);
                    path_data(&mut body, &points, flags.as_deref(), true)?;
                    let _ = write!(
                        body,
                        r#"" fill-rule="evenodd"{}/>"#,
                        paint(&state, true, true, 0.0)
                    );
                }
            }
            // POLYPOLYGON, TRANSPARENT
            111 | 142 => {
                let Some(n) = a.u16() else {
                    r.p = end;
                    continue;
                };
                let mut polys: Vec<Polygon> = Vec::new();
                for _ in 0..n {
                    match a.polygon() {
                        Some(p) => polys.push((p, None)),
                        None => break,
                    }
                }
                let mut opacity = 1.0;
                if kind == 142 {
                    opacity = 1.0 - a.u16().unwrap_or(0).min(100) as f64 / 100.0;
                } else if version >= 2 {
                    if let Some(complex) = a.u16() {
                        for _ in 0..complex {
                            let Some(index) = a.u16() else { break };
                            let Some(poly) = a.polygon_with_flags() else {
                                break;
                            };
                            if let Some(slot) = polys.get_mut(index as usize) {
                                *slot = poly;
                            }
                        }
                    }
                }
                let alpha = if opacity < 1.0 {
                    format!(r#" opacity="{opacity:.3}""#)
                } else {
                    String::new()
                };
                body.push_str(r#"<path d=""#);
                for (p, f) in &polys {
                    path_data(&mut body, p, f.as_deref(), true)?;
                }
                let _ = write!(
                    body,
                    r#"" fill-rule="evenodd"{}{alpha}/>"#,
                    paint(&state, true, kind == 111, 0.0)
                );
            }
            // TEXT, TEXTARRAY, STRETCHTEXT
            112..=114 => {
                let Some(origin) = a.point() else {
                    r.p = end;
                    continue;
                };
                let Some(mut text) = a.string(state.font.unicode) else {
                    r.p = end;
                    continue;
                };
                if kind == 114 {
                    let _ = a.u32();
                }
                let index = a.u16().unwrap_or(0) as usize;
                let len = a.u16().unwrap_or(u16::MAX) as usize;
                let mut dx: Vec<i32> = Vec::new();
                if kind == 113 {
                    let n = a.u32().unwrap_or(0).min(1 << 20);
                    for _ in 0..n {
                        match a.i32() {
                            Some(v) => dx.push(v),
                            None => break,
                        }
                    }
                }
                if version >= 2 && !state.font.unicode {
                    if let Some(n) = a.u16() {
                        if let Some(u) = a.utf16(n as usize) {
                            if !u.is_empty() {
                                text = u;
                            }
                        }
                    }
                }
                // The shown text, escaped (at most 6 bytes per input byte) with
                // the family, must fit the output bound before it is copied.
                let escaped_bound = text
                    .len()
                    .min(len.saturating_mul(4))
                    .saturating_add(state.font.family.len())
                    .saturating_mul(6);
                if body.len().saturating_add(escaped_bound) > MAX_SVG_BYTES {
                    return None;
                }
                let shown: Vec<char> = text.chars().skip(index).take(len).collect();
                if shown.iter().all(|c| c.is_whitespace()) {
                    r.p = end;
                    continue;
                }
                let (x, y) = pt(origin);
                let size = len_scale(state.font.height.abs().max(1.0));
                // Text alignment: 0 top, 1 baseline, 2 bottom.
                let baseline_shift = match state.align {
                    0 => size * 0.8,
                    2 => -size * 0.2,
                    _ => 0.0,
                };
                let angle = state.font.orientation;
                let (sin, cos) = (-angle.to_radians()).sin_cos();
                let (bx, by) = (x - sin * baseline_shift, y + cos * baseline_shift);
                let mut xs = String::new();
                if !dx.is_empty() {
                    for (i, _) in shown.iter().enumerate() {
                        let offset = if i == 0 {
                            0.0
                        } else {
                            dx.get(i - 1).copied().unwrap_or(0) as f64
                        };
                        let (px, _) = pt((origin.0 + offset, origin.1));
                        let _ = write!(xs, "{:.2} ", px + (bx - x));
                    }
                }
                let weight = match state.font.weight {
                    0 | 5 => 400,
                    1 => 100,
                    2 => 200,
                    3 => 300,
                    4 => 350,
                    6 => 500,
                    7 => 600,
                    8 => 700,
                    9 => 800,
                    _ => 900,
                };
                let mut deco = String::new();
                if state.font.underline {
                    deco.push_str("underline ");
                }
                if state.font.strike {
                    deco.push_str("line-through");
                }
                let transform = if angle != 0.0 {
                    format!(r#" transform="rotate({:.3} {bx:.2} {by:.2})""#, -angle)
                } else {
                    String::new()
                };
                let positions = if xs.is_empty() || angle != 0.0 {
                    format!(r#"x="{bx:.2}""#)
                } else {
                    format!(r#"x="{}""#, xs.trim_end())
                };
                let _ = write!(
                    body,
                    r#"<text {positions} y="{by:.2}" font-family="{}" font-size="{size:.2}" font-weight="{weight}"{}{} fill="{}" xml:space="preserve"{transform}>{}</text>"#,
                    escape(&state.font.family),
                    if state.font.italic {
                        r#" font-style="italic""#
                    } else {
                        ""
                    },
                    if deco.is_empty() {
                        String::new()
                    } else {
                        format!(r#" text-decoration="{}""#, deco.trim_end())
                    },
                    css(state.text),
                    escape(&shown.iter().collect::<String>())
                );
            }
            // LINECOLOR, FILLCOLOR
            132 | 133 => {
                if let (Some(c), Some(set)) = (a.color(), a.u8()) {
                    let value = (set != 0).then_some(c);
                    if kind == 132 {
                        state.line = value;
                    } else {
                        state.fill = value;
                    }
                }
            }
            // TEXTCOLOR
            134 => {
                if let Some(c) = a.color() {
                    state.text = c;
                }
            }
            // TEXTALIGN
            136 => {
                if let Some(v) = a.u16() {
                    state.align = v;
                }
            }
            // MAPMODE
            137 => {
                if let Some(m) = read_map_mode(&mut a, Some(&state.map)) {
                    state.map = m;
                }
            }
            // FONT
            138 => {
                if let Some(font) = read_font(&mut a) {
                    state.font = font;
                }
            }
            // PUSH, POP
            139 => {
                if stack.len() < MAX_STACK {
                    stack.push(state.clone());
                }
            }
            140 => {
                if let Some(s) = stack.pop() {
                    state = s;
                }
            }
            _ => {}
        }
        r.p = end;
        if body.len() > MAX_SVG_BYTES {
            return None;
        }
    }
    let svg = format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{:.3}mm" height="{:.3}mm" viewBox="0 0 {width} {height}" preserveAspectRatio="none">{body}</svg>"#,
        width * header_map.unit * header_map.scale.0.abs() / 100.0,
        height * header_map.unit * header_map.scale.1.abs() / 100.0,
    );
    let emu = |v: f64, s: f64| v * header_map.unit * s.abs() * 360.0;
    Some(Converted {
        size_emu: (
            emu(width, header_map.scale.0),
            emu(height, header_map.scale.1),
        ),
        svg: svg.into_bytes(),
    })
}

/// `LineInfo` (compat; style; width), returning the width.
fn line_width(r: &mut Reader<'_>) -> Option<f64> {
    let (_, mut r) = r.nested()?;
    let _style = r.u16()?;
    let width = r.i32()? as f64;
    Some(width)
}

/// `vcl::Font` as SvmReader reads it.
fn read_font(r: &mut Reader<'_>) -> Option<Font> {
    let (version, mut inner) = r.nested()?;
    let r = &mut inner;
    // Names are byte strings in the font's own charset; the family is
    // written before the charset is known, so it is read as a byte string
    // (LibreOffice writes ASCII family names here) unless the length prefix
    // spells a UTF-16 string.
    let family = r.string(false)?;
    let _style = r.string(false)?;
    let _width = r.i32()?;
    let height = r.i32()? as f64;
    let charset = r.u16()?;
    let _family_kind = r.u16()?;
    let _pitch = r.u16()?;
    let weight = r.u16()?;
    let underline = r.u16()?;
    let strike = r.u16()?;
    let italic = r.u16()?;
    let _language = r.u16()?;
    let _width_type = r.u16()?;
    let orientation = r.u16()? as i16 as f64 / 10.0;
    let _ = version;
    Some(Font {
        family: std::rc::Rc::from(family),
        height,
        weight,
        italic: italic != 0,
        underline: underline != 0,
        strike: strike != 0,
        orientation,
        // RTL_TEXTENCODING_UNICODE: strings are UTF-16 with a u32 length.
        unicode: charset == 0xFFFF,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn compat(version: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = version.to_le_bytes().to_vec();
        out.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }
    fn map_mode() -> Vec<u8> {
        let mut m = 0u16.to_le_bytes().to_vec();
        for v in [0i32, 0, 1, 1, 1, 1] {
            m.extend_from_slice(&v.to_le_bytes());
        }
        m.push(1);
        compat(1, &m)
    }
    fn action(kind: u16, version: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = kind.to_le_bytes().to_vec();
        out.extend(compat(version, payload));
        out
    }

    #[test]
    fn converts_polygons_and_unicode_text() {
        let mut header = 0u32.to_le_bytes().to_vec();
        header.extend(map_mode());
        header.extend_from_slice(&1000i32.to_le_bytes());
        header.extend_from_slice(&500i32.to_le_bytes());
        header.extend_from_slice(&4u32.to_le_bytes());
        let mut file = b"VCLMTF".to_vec();
        file.extend(compat(1, &header));
        file.extend(action(133, 1, &[0xBD, 0x81, 0x4F, 0, 1]));
        let mut poly = 3u16.to_le_bytes().to_vec();
        for v in [0i32, 0, 100, 0, 100, 100] {
            poly.extend_from_slice(&v.to_le_bytes());
        }
        file.extend(action(110, 1, &poly));
        let mut font = 4u16.to_le_bytes().to_vec();
        font.extend_from_slice(b"Sans");
        font.extend_from_slice(&0u16.to_le_bytes());
        font.extend_from_slice(&0i32.to_le_bytes());
        font.extend_from_slice(&200i32.to_le_bytes());
        for v in [0xFFFFu16, 0, 0, 8, 0, 0, 0, 0, 0, 0] {
            font.extend_from_slice(&v.to_le_bytes());
        }
        let mut text = Vec::new();
        for v in [10i32, 300] {
            text.extend_from_slice(&v.to_le_bytes());
        }
        text.extend_from_slice(&2u32.to_le_bytes());
        text.extend_from_slice(&[b'H', 0, b'i', 0]);
        text.extend_from_slice(&0u16.to_le_bytes());
        text.extend_from_slice(&2u16.to_le_bytes());
        let mut with_font = action(138, 1, &compat(1, &font));
        with_font.extend(action(112, 1, &text));
        file.extend(with_font);
        file.extend(action(133, 1, &[0, 0, 0, 0, 1]));
        let svg = String::from_utf8(to_svg(&file).unwrap().svg).unwrap();
        assert!(svg.contains(r##"<path d="M0.00 0.00L100.00 0.00L100.00 100.00Z" fill-rule="evenodd" fill="#4f81bd""##), "{svg}");
        assert!(
            svg.contains(r#"font-family="Sans" font-size="200.00" font-weight="700""#),
            "{svg}"
        );
        assert!(svg.contains(">Hi</text>"), "{svg}");
        // The action count bounds the stream: the trailing action is not read.
        assert_eq!(svg.matches("<path").count(), 1);
    }
}
