//! Drawing geometry of the ODP reader: `draw:transform` (Part 3 §19.229),
//! SVG path data (`svg:d`, §19.554, and `draw:points`), and the enhanced
//! geometry of custom shapes (`draw:enhanced-geometry`, §19.145 path
//! commands and §19.171 formulas).
//!
//! Geometry is produced in shape-local EMU (origin at the shape box's
//! top-left corner) as absolute segments; the reader normalizes them into
//! the PPTX custom-geometry model. Arcs are converted to cubic Béziers here
//! so the model never needs an arc whose radii depend on a zero-sized box.

use crate::error::invalid;
use crate::odp_units::{length_emu, EMU_PER_HMM};
use std::collections::HashMap;

pub const NS_DRAWOOO: &str = "http://openoffice.org/2010/draw";

/// Implementation bound on the segments one shape may emit (all paths).
pub const MAX_SEGMENTS_PER_SHAPE: usize = 200_000;
/// Implementation bound on formula evaluation steps for one shape.
const MAX_FORMULA_WORK: usize = 2_000_000;
/// Implementation bound on formula nesting (parser recursion).
const MAX_FORMULA_DEPTH: usize = 64;
/// Implementation bound on the nodes of one parsed formula. It also bounds
/// the depth of the expression tree (a flat `a+b+…` chain is left-deep), so
/// evaluating and dropping a tree never recurses deeper than this.
const MAX_FORMULA_NODES: usize = 512;
/// Implementation bound on nested evaluation across equation references
/// (`?f1` → `?f2` → …): deeper chains evaluate to 0 like a cycle.
const MAX_EVAL_DEPTH: usize = 512;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Seg {
    Move(f64, f64),
    Line(f64, f64),
    Cubic(f64, f64, f64, f64, f64, f64),
    Quad(f64, f64, f64, f64),
    Close,
}

impl Seg {
    pub fn end(&self) -> Option<(f64, f64)> {
        match *self {
            Seg::Move(x, y) | Seg::Line(x, y) => Some((x, y)),
            Seg::Cubic(_, _, _, _, x, y) | Seg::Quad(_, _, x, y) => Some((x, y)),
            Seg::Close => None,
        }
    }
}

/// One painted path: a set of subpaths filled together (§19.145 `N`).
#[derive(Clone, Debug, Default)]
pub struct GeoPath {
    pub segs: Vec<Seg>,
    pub fill: bool,
    pub stroke: bool,
    /// `H`/`I`/`J`/`K` fill modifiers as the DrawingML path fill mode.
    pub fill_mode: Option<&'static str>,
}

#[derive(Clone, Debug, Default)]
pub struct Geometry {
    pub paths: Vec<GeoPath>,
    /// The first `draw:text-areas` rectangle in shape-local EMU.
    pub text_area: Option<(f64, f64, f64, f64)>,
}

impl Geometry {
    pub fn bounds(&self) -> Option<(f64, f64, f64, f64)> {
        let mut b: Option<(f64, f64, f64, f64)> = None;
        let mut add = |x: f64, y: f64| {
            b = Some(match b {
                None => (x, y, x, y),
                Some((a, c, d, e)) => (a.min(x), c.min(y), d.max(x), e.max(y)),
            });
        };
        for path in &self.paths {
            for seg in &path.segs {
                match *seg {
                    Seg::Move(x, y) | Seg::Line(x, y) => add(x, y),
                    Seg::Cubic(a, b2, c, d, x, y) => {
                        add(a, b2);
                        add(c, d);
                        add(x, y)
                    }
                    Seg::Quad(a, b2, x, y) => {
                        add(a, b2);
                        add(x, y)
                    }
                    Seg::Close => {}
                }
            }
        }
        b
    }

    pub fn translate(&mut self, dx: f64, dy: f64) {
        for path in &mut self.paths {
            for seg in &mut path.segs {
                *seg = map_seg(*seg, |x, y| (x + dx, y + dy));
            }
        }
        if let Some(t) = &mut self.text_area {
            t.0 += dx;
            t.1 += dy;
        }
    }
}

pub fn map_seg(seg: Seg, f: impl Fn(f64, f64) -> (f64, f64)) -> Seg {
    match seg {
        Seg::Move(x, y) => {
            let (x, y) = f(x, y);
            Seg::Move(x, y)
        }
        Seg::Line(x, y) => {
            let (x, y) = f(x, y);
            Seg::Line(x, y)
        }
        Seg::Cubic(a, b, c, d, x, y) => {
            let (a, b) = f(a, b);
            let (c, d) = f(c, d);
            let (x, y) = f(x, y);
            Seg::Cubic(a, b, c, d, x, y)
        }
        Seg::Quad(a, b, x, y) => {
            let (a, b) = f(a, b);
            let (x, y) = f(x, y);
            Seg::Quad(a, b, x, y)
        }
        Seg::Close => Seg::Close,
    }
}

// ---------------------------------------------------------------------------
// Affine transforms (draw:transform)
// ---------------------------------------------------------------------------

/// 2×3 affine matrix: x' = a·x + c·y + e, y' = b·x + d·y + f.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Affine {
    pub a: f64,
    pub b: f64,
    pub c: f64,
    pub d: f64,
    pub e: f64,
    pub f: f64,
}

impl Affine {
    pub const IDENTITY: Affine = Affine {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: 0.0,
        f: 0.0,
    };

    /// `self` followed by `next`.
    pub fn then(&self, next: &Affine) -> Affine {
        Affine {
            a: next.a * self.a + next.c * self.b,
            b: next.b * self.a + next.d * self.b,
            c: next.a * self.c + next.c * self.d,
            d: next.b * self.c + next.d * self.d,
            e: next.a * self.e + next.c * self.f + next.e,
            f: next.b * self.e + next.d * self.f + next.f,
        }
    }

    pub fn apply(&self, x: f64, y: f64) -> (f64, f64) {
        (
            self.a * x + self.c * y + self.e,
            self.b * x + self.d * y + self.f,
        )
    }

    pub fn translate(x: f64, y: f64) -> Affine {
        Affine {
            e: x,
            f: y,
            ..Affine::IDENTITY
        }
    }
}

/// Parse `draw:transform` (§19.229). The listed transforms apply left to
/// right (the first one is applied to the shape first): LibreOffice writes
/// `rotate (r) translate (x y)` for a shape rotated about its origin and
/// then placed. `rotate` takes radians and turns counter-clockwise on the
/// page (a PowerPoint 90° clockwise rotation is written `rotate (-1.5708)`).
/// Translations are lengths.
pub fn parse_transform(value: &str) -> Option<Affine> {
    let mut m = Affine::IDENTITY;
    let mut rest = value.trim();
    while !rest.is_empty() {
        let open = rest.find('(')?;
        let close = rest[open..].find(')')? + open;
        let name = rest[..open].trim();
        let args: Vec<&str> = rest[open + 1..close]
            .split(|c: char| c == ',' || c.is_whitespace())
            .filter(|s| !s.is_empty())
            .collect();
        let num = |i: usize| -> Option<f64> { args.get(i)?.parse::<f64>().ok() };
        let len = |i: usize| -> Option<f64> {
            let arg = args.get(i)?;
            // A plain number in a translate is 1/100 mm (LibreOffice writes
            // units, other producers may not).
            length_emu(arg)
        };
        let step = match name {
            "rotate" => {
                let r = num(0)?;
                let (s, c) = r.sin_cos();
                Affine {
                    a: c,
                    b: -s,
                    c: s,
                    d: c,
                    e: 0.0,
                    f: 0.0,
                }
            }
            "translate" => Affine::translate(len(0)?, len(1).unwrap_or(0.0)),
            "scale" => {
                let sx = num(0)?;
                let sy = num(1).unwrap_or(sx);
                Affine {
                    a: sx,
                    d: sy,
                    ..Affine::IDENTITY
                }
            }
            "skewX" => Affine {
                c: num(0)?.tan(),
                ..Affine::IDENTITY
            },
            "skewY" => Affine {
                b: num(0)?.tan(),
                ..Affine::IDENTITY
            },
            "matrix" => Affine {
                a: num(0)?,
                b: num(1)?,
                c: num(2)?,
                d: num(3)?,
                e: len(4)?,
                f: len(5)?,
            },
            _ => return None,
        };
        m = m.then(&step);
        rest = rest[close + 1..].trim_start_matches(|c: char| c == ',' || c.is_whitespace());
    }
    [m.a, m.b, m.c, m.d, m.e, m.f]
        .iter()
        .all(|v| v.is_finite())
        .then_some(m)
}

/// A shape box mapped by an affine transform, in the PPTX transform model:
/// the axis-aligned box (x, y, w, h) before rotation, a clockwise rotation
/// about its centre in degrees, and a vertical mirror for a reflecting
/// transform.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Placed {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    pub rot: f64,
    pub flip_v: bool,
}

pub fn place(w: f64, h: f64, m: &Affine) -> Placed {
    let sx = m.a.hypot(m.b);
    let det = m.a * m.d - m.b * m.c;
    let sy = if sx > 0.0 { det / sx } else { m.c.hypot(m.d) };
    let (cx, cy) = m.apply(w / 2.0, h / 2.0);
    let rot = if sx > 0.0 {
        m.b.atan2(m.a).to_degrees()
    } else {
        0.0
    };
    let nw = w * sx;
    let nh = h * sy.abs();
    Placed {
        x: cx - nw / 2.0,
        y: cy - nh / 2.0,
        w: nw,
        h: nh,
        rot: normalize_degrees(rot),
        flip_v: sy < 0.0,
    }
}

pub fn normalize_degrees(value: f64) -> f64 {
    let mut v = value % 360.0;
    if v < 0.0 {
        v += 360.0;
    }
    if (v - 360.0).abs() < 1e-9 || v.abs() < 1e-9 {
        0.0
    } else {
        v
    }
}

// ---------------------------------------------------------------------------
// Arcs
// ---------------------------------------------------------------------------

/// Cubic Béziers for an elliptical arc in y-down parametric angles
/// (radians): P(t) = (cx + rx·cos t, cy + ry·sin t), from `start` sweeping
/// `sweep` (positive = clockwise on the page).
pub fn arc_segments(
    cx: f64,
    cy: f64,
    rx: f64,
    ry: f64,
    start: f64,
    sweep: f64,
    out: &mut Vec<Seg>,
) {
    if !(rx.is_finite() && ry.is_finite() && start.is_finite() && sweep.is_finite()) {
        return;
    }
    let pieces = ((sweep.abs() / (std::f64::consts::FRAC_PI_2)).ceil() as usize).clamp(1, 8);
    let step = sweep / pieces as f64;
    let alpha = 4.0 / 3.0 * (step / 4.0).tan();
    let point = |t: f64| (cx + rx * t.cos(), cy + ry * t.sin());
    let deriv = |t: f64| (-rx * t.sin(), ry * t.cos());
    for i in 0..pieces {
        let a = start + step * i as f64;
        let b = a + step;
        let (ax, ay) = point(a);
        let (bx, by) = point(b);
        let (dax, day) = deriv(a);
        let (dbx, dby) = deriv(b);
        out.push(Seg::Cubic(
            ax + alpha * dax,
            ay + alpha * day,
            bx - alpha * dbx,
            by - alpha * dby,
            bx,
            by,
        ));
    }
}

// ---------------------------------------------------------------------------
// SVG path data (svg:d, §19.554) and point lists (draw:points, §19.208)
// ---------------------------------------------------------------------------

/// Parse SVG path data into subpaths of absolute segments in the path's own
/// coordinates. Every subpath belongs to one painted path (SVG fills the
/// whole path together).
pub fn parse_svg_path(d: &str) -> Option<Vec<Seg>> {
    let mut out = Vec::new();
    let mut tokens = SvgTokens::new(d);
    let (mut cx, mut cy) = (0.0f64, 0.0f64);
    let (mut sx, mut sy) = (0.0f64, 0.0f64);
    let mut last_ctrl: Option<(f64, f64, char)> = None;
    let mut command: Option<char> = None;
    loop {
        let cmd = match tokens.peek_command() {
            Some(c) => {
                tokens.next_command();
                c
            }
            None => {
                if !tokens.has_number() {
                    break;
                }
                // Implicit repetition (after M, repeated coordinates are L).
                match command? {
                    'M' => 'L',
                    'm' => 'l',
                    'Z' | 'z' => return None,
                    c => c,
                }
            }
        };
        command = Some(cmd);
        let rel = cmd.is_ascii_lowercase();
        let (ox, oy) = if rel { (cx, cy) } else { (0.0, 0.0) };
        match cmd.to_ascii_uppercase() {
            'M' => {
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                out.push(Seg::Move(x, y));
                cx = x;
                cy = y;
                sx = x;
                sy = y;
                last_ctrl = None;
            }
            'L' => {
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                out.push(Seg::Line(x, y));
                cx = x;
                cy = y;
                last_ctrl = None;
            }
            'H' => {
                let x = tokens.num()? + ox;
                out.push(Seg::Line(x, cy));
                cx = x;
                last_ctrl = None;
            }
            'V' => {
                let y = tokens.num()? + if rel { cy } else { 0.0 };
                out.push(Seg::Line(cx, y));
                cy = y;
                last_ctrl = None;
            }
            'C' => {
                let (x1, y1) = (tokens.num()? + ox, tokens.num()? + oy);
                let (x2, y2) = (tokens.num()? + ox, tokens.num()? + oy);
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                out.push(Seg::Cubic(x1, y1, x2, y2, x, y));
                last_ctrl = Some((x2, y2, 'C'));
                cx = x;
                cy = y;
            }
            'S' => {
                let (x1, y1) = match last_ctrl {
                    Some((px, py, 'C')) => (2.0 * cx - px, 2.0 * cy - py),
                    _ => (cx, cy),
                };
                let (x2, y2) = (tokens.num()? + ox, tokens.num()? + oy);
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                out.push(Seg::Cubic(x1, y1, x2, y2, x, y));
                last_ctrl = Some((x2, y2, 'C'));
                cx = x;
                cy = y;
            }
            'Q' => {
                let (x1, y1) = (tokens.num()? + ox, tokens.num()? + oy);
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                out.push(Seg::Quad(x1, y1, x, y));
                last_ctrl = Some((x1, y1, 'Q'));
                cx = x;
                cy = y;
            }
            'T' => {
                let (x1, y1) = match last_ctrl {
                    Some((px, py, 'Q')) => (2.0 * cx - px, 2.0 * cy - py),
                    _ => (cx, cy),
                };
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                out.push(Seg::Quad(x1, y1, x, y));
                last_ctrl = Some((x1, y1, 'Q'));
                cx = x;
                cy = y;
            }
            'A' => {
                let rx = tokens.num()?.abs();
                let ry = tokens.num()?.abs();
                let phi = tokens.num()?.to_radians();
                let large = tokens.flag()?;
                let sweep = tokens.flag()?;
                let (x, y) = (tokens.num()? + ox, tokens.num()? + oy);
                svg_arc(cx, cy, rx, ry, phi, large, sweep, x, y, &mut out);
                cx = x;
                cy = y;
                last_ctrl = None;
            }
            'Z' => {
                out.push(Seg::Close);
                cx = sx;
                cy = sy;
                last_ctrl = None;
            }
            _ => return None,
        }
        if out.len() > MAX_SEGMENTS_PER_SHAPE {
            return None;
        }
    }
    Some(out)
}

/// SVG endpoint arc → centre parameterization (SVG 1.1 F.6.5), rotated by
/// `phi`, emitted as Béziers.
#[allow(clippy::too_many_arguments)]
fn svg_arc(
    x1: f64,
    y1: f64,
    mut rx: f64,
    mut ry: f64,
    phi: f64,
    large: bool,
    sweep: bool,
    x2: f64,
    y2: f64,
    out: &mut Vec<Seg>,
) {
    if rx == 0.0 || ry == 0.0 || (x1 == x2 && y1 == y2) {
        out.push(Seg::Line(x2, y2));
        return;
    }
    let (s, c) = phi.sin_cos();
    let dx = (x1 - x2) / 2.0;
    let dy = (y1 - y2) / 2.0;
    let x1p = c * dx + s * dy;
    let y1p = -s * dx + c * dy;
    let lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
    if lambda > 1.0 {
        rx *= lambda.sqrt();
        ry *= lambda.sqrt();
    }
    let num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
    let den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
    let mut coef = if den == 0.0 {
        0.0
    } else {
        (num / den).max(0.0).sqrt()
    };
    if large == sweep {
        coef = -coef;
    }
    let cxp = coef * rx * y1p / ry;
    let cyp = -coef * ry * x1p / rx;
    let cx = c * cxp - s * cyp + (x1 + x2) / 2.0;
    let cy = s * cxp + c * cyp + (y1 + y2) / 2.0;
    let angle = |ux: f64, uy: f64, vx: f64, vy: f64| {
        let dot = ux * vx + uy * vy;
        let len = (ux * ux + uy * uy).sqrt() * (vx * vx + vy * vy).sqrt();
        let mut a = (dot / len).clamp(-1.0, 1.0).acos();
        if ux * vy - uy * vx < 0.0 {
            a = -a;
        }
        a
    };
    let theta1 = angle(1.0, 0.0, (x1p - cxp) / rx, (y1p - cyp) / ry);
    let mut dtheta = angle(
        (x1p - cxp) / rx,
        (y1p - cyp) / ry,
        (-x1p - cxp) / rx,
        (-y1p - cyp) / ry,
    );
    if !sweep && dtheta > 0.0 {
        dtheta -= std::f64::consts::TAU;
    } else if sweep && dtheta < 0.0 {
        dtheta += std::f64::consts::TAU;
    }
    let mut local = Vec::new();
    arc_segments(0.0, 0.0, rx, ry, theta1, dtheta, &mut local);
    for seg in local {
        out.push(map_seg(seg, |x, y| {
            (c * x - s * y + cx, s * x + c * y + cy)
        }));
    }
}

struct SvgTokens<'s> {
    bytes: &'s [u8],
    pos: usize,
}

impl<'s> SvgTokens<'s> {
    fn new(text: &'s str) -> Self {
        Self {
            bytes: text.as_bytes(),
            pos: 0,
        }
    }
    fn skip(&mut self) {
        while self.pos < self.bytes.len()
            && (self.bytes[self.pos].is_ascii_whitespace() || self.bytes[self.pos] == b',')
        {
            self.pos += 1;
        }
    }
    fn peek_command(&mut self) -> Option<char> {
        self.skip();
        let b = *self.bytes.get(self.pos)?;
        (b.is_ascii_alphabetic() && b != b'e' && b != b'E').then_some(b as char)
    }
    fn next_command(&mut self) {
        self.pos += 1;
    }
    fn has_number(&mut self) -> bool {
        self.skip();
        self.bytes
            .get(self.pos)
            .is_some_and(|b| b.is_ascii_digit() || matches!(b, b'-' | b'+' | b'.'))
    }
    fn flag(&mut self) -> Option<bool> {
        self.skip();
        let b = *self.bytes.get(self.pos)?;
        self.pos += 1;
        match b {
            b'0' => Some(false),
            b'1' => Some(true),
            _ => None,
        }
    }
    fn num(&mut self) -> Option<f64> {
        self.skip();
        let start = self.pos;
        let mut seen_dot = false;
        let mut seen_exp = false;
        if matches!(self.bytes.get(self.pos), Some(b'-' | b'+')) {
            self.pos += 1;
        }
        while let Some(&b) = self.bytes.get(self.pos) {
            if b.is_ascii_digit() {
                self.pos += 1;
            } else if b == b'.' && !seen_dot && !seen_exp {
                seen_dot = true;
                self.pos += 1;
            } else if (b == b'e' || b == b'E') && !seen_exp {
                seen_exp = true;
                self.pos += 1;
                if matches!(self.bytes.get(self.pos), Some(b'-' | b'+')) {
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
        let text = std::str::from_utf8(&self.bytes[start..self.pos]).ok()?;
        let v: f64 = text.parse().ok()?;
        v.is_finite().then_some(v)
    }
}

/// `draw:points` (§19.208): "x,y x,y …".
pub fn parse_points(value: &str) -> Option<Vec<(f64, f64)>> {
    let mut out = Vec::new();
    for pair in value.split_whitespace() {
        let (x, y) = pair.split_once(',')?;
        let x: f64 = x.trim().parse().ok()?;
        let y: f64 = y.trim().parse().ok()?;
        if !x.is_finite() || !y.is_finite() {
            return None;
        }
        out.push((x, y));
        if out.len() > MAX_SEGMENTS_PER_SHAPE {
            return None;
        }
    }
    Some(out)
}

/// `svg:viewBox` (§19.570): "x y w h".
pub fn parse_view_box(value: &str) -> Option<(f64, f64, f64, f64)> {
    let v: Vec<f64> = value
        .split(|c: char| c == ',' || c.is_whitespace())
        .filter(|s| !s.is_empty())
        .map(|s| s.parse::<f64>())
        .collect::<Result<_, _>>()
        .ok()?;
    (v.len() == 4 && v.iter().all(|x| x.is_finite())).then(|| (v[0], v[1], v[2], v[3]))
}

// ---------------------------------------------------------------------------
// Enhanced geometry (§19.145 draw:enhanced-path, §19.171 draw:formula)
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
enum Expr {
    Num(f64),
    Eq(usize),
    Modifier(usize),
    Const(Constant),
    Neg(Box<Expr>),
    Bin(char, Box<Expr>, Box<Expr>),
    Func(Func, Vec<Expr>),
}

#[derive(Clone, Copy, Debug)]
enum Constant {
    Pi,
    Left,
    Top,
    Right,
    Bottom,
    XStretch,
    YStretch,
    HasStroke,
    HasFill,
    Width,
    Height,
    LogWidth,
    LogHeight,
}

#[derive(Clone, Copy, Debug)]
enum Func {
    Abs,
    Sqrt,
    Sin,
    Cos,
    Tan,
    Atan,
    Atan2,
    Min,
    Max,
    If,
}

struct FormulaParser<'s, 'n> {
    bytes: &'s [u8],
    pos: usize,
    names: &'n HashMap<&'n str, usize>,
    depth: usize,
    nodes: usize,
}

impl<'s, 'n> FormulaParser<'s, 'n> {
    fn skip(&mut self) {
        while self.pos < self.bytes.len() && self.bytes[self.pos].is_ascii_whitespace() {
            self.pos += 1;
        }
    }
    fn peek(&mut self) -> Option<u8> {
        self.skip();
        self.bytes.get(self.pos).copied()
    }
    /// Count one expression node against the formula's node bound.
    fn node(&mut self) -> Option<()> {
        self.nodes += 1;
        (self.nodes <= MAX_FORMULA_NODES).then_some(())
    }
    fn expr(&mut self) -> Option<Expr> {
        self.depth += 1;
        if self.depth > MAX_FORMULA_DEPTH {
            return None;
        }
        let mut left = self.term()?;
        while let Some(op @ (b'+' | b'-')) = self.peek() {
            self.pos += 1;
            self.node()?;
            let right = self.term()?;
            left = Expr::Bin(op as char, Box::new(left), Box::new(right));
        }
        self.depth -= 1;
        Some(left)
    }
    fn term(&mut self) -> Option<Expr> {
        let mut left = self.unary()?;
        while let Some(op @ (b'*' | b'/')) = self.peek() {
            self.pos += 1;
            self.node()?;
            let right = self.unary()?;
            left = Expr::Bin(op as char, Box::new(left), Box::new(right));
        }
        Some(left)
    }
    fn unary(&mut self) -> Option<Expr> {
        match self.peek()? {
            b'-' => {
                self.pos += 1;
                self.node()?;
                self.depth += 1;
                if self.depth > MAX_FORMULA_DEPTH {
                    return None;
                }
                let inner = self.unary()?;
                self.depth -= 1;
                Some(Expr::Neg(Box::new(inner)))
            }
            b'+' => {
                self.pos += 1;
                self.unary()
            }
            _ => self.primary(),
        }
    }
    fn primary(&mut self) -> Option<Expr> {
        self.node()?;
        let b = self.peek()?;
        if b == b'(' {
            self.pos += 1;
            let e = self.expr()?;
            if self.peek()? != b')' {
                return None;
            }
            self.pos += 1;
            return Some(e);
        }
        if b == b'?' {
            self.pos += 1;
            let name = self.ident();
            return self.names.get(name.as_str()).map(|i| Expr::Eq(*i));
        }
        if b == b'$' {
            self.pos += 1;
            let start = self.pos;
            while self.pos < self.bytes.len() && self.bytes[self.pos].is_ascii_digit() {
                self.pos += 1;
            }
            let n: usize = std::str::from_utf8(&self.bytes[start..self.pos])
                .ok()?
                .parse()
                .ok()?;
            return Some(Expr::Modifier(n));
        }
        if b.is_ascii_digit() || b == b'.' {
            return number(self.bytes, &mut self.pos).map(Expr::Num);
        }
        if b.is_ascii_alphabetic() {
            let name = self.ident();
            let constant = match name.as_str() {
                "pi" => Some(Constant::Pi),
                "left" => Some(Constant::Left),
                "top" => Some(Constant::Top),
                "right" => Some(Constant::Right),
                "bottom" => Some(Constant::Bottom),
                "xstretch" => Some(Constant::XStretch),
                "ystretch" => Some(Constant::YStretch),
                "hasstroke" => Some(Constant::HasStroke),
                "hasfill" => Some(Constant::HasFill),
                "width" => Some(Constant::Width),
                "height" => Some(Constant::Height),
                "logwidth" => Some(Constant::LogWidth),
                "logheight" => Some(Constant::LogHeight),
                _ => None,
            };
            if let Some(c) = constant {
                return Some(Expr::Const(c));
            }
            let func = match name.as_str() {
                "abs" => Func::Abs,
                "sqrt" => Func::Sqrt,
                "sin" => Func::Sin,
                "cos" => Func::Cos,
                "tan" => Func::Tan,
                "atan" => Func::Atan,
                "atan2" => Func::Atan2,
                "min" => Func::Min,
                "max" => Func::Max,
                "if" => Func::If,
                _ => return None,
            };
            if self.peek()? != b'(' {
                return None;
            }
            self.pos += 1;
            let mut args = vec![self.expr()?];
            while self.peek()? == b',' {
                self.pos += 1;
                args.push(self.expr()?);
            }
            if self.peek()? != b')' {
                return None;
            }
            self.pos += 1;
            let arity = match func {
                Func::Atan2 | Func::Min | Func::Max => 2,
                Func::If => 3,
                _ => 1,
            };
            return (args.len() == arity).then_some(Expr::Func(func, args));
        }
        None
    }
    fn ident(&mut self) -> String {
        let start = self.pos;
        while self.pos < self.bytes.len()
            && (self.bytes[self.pos].is_ascii_alphanumeric() || self.bytes[self.pos] == b'_')
        {
            self.pos += 1;
        }
        String::from_utf8_lossy(&self.bytes[start..self.pos]).into_owned()
    }
}

fn number(bytes: &[u8], pos: &mut usize) -> Option<f64> {
    let start = *pos;
    let mut seen_exp = false;
    while let Some(&b) = bytes.get(*pos) {
        if b.is_ascii_digit() || b == b'.' {
            *pos += 1;
        } else if (b == b'e' || b == b'E')
            && !seen_exp
            && bytes
                .get(*pos + 1)
                .is_some_and(|n| n.is_ascii_digit() || *n == b'-' || *n == b'+')
        {
            seen_exp = true;
            *pos += 2;
        } else {
            break;
        }
    }
    let v: f64 = std::str::from_utf8(&bytes[start..*pos])
        .ok()?
        .parse()
        .ok()?;
    v.is_finite().then_some(v)
}

/// The evaluation context of one custom shape.
pub struct ShapeFrame {
    /// Shape box in EMU.
    pub width: f64,
    pub height: f64,
    pub has_fill: bool,
    pub has_stroke: bool,
}

struct Evaluator {
    exprs: Vec<Option<Expr>>,
    values: Vec<Option<f64>>,
    busy: Vec<bool>,
    modifiers: Vec<f64>,
    view: (f64, f64, f64, f64),
    log: (f64, f64),
    has_fill: bool,
    has_stroke: bool,
    work: usize,
    /// Current evaluation nesting (expression levels and equation hops).
    depth: usize,
}

impl Evaluator {
    fn eval(&mut self, e: &Expr) -> f64 {
        self.work += 1;
        if self.work > MAX_FORMULA_WORK || self.depth >= MAX_EVAL_DEPTH {
            return 0.0;
        }
        self.depth += 1;
        let v = self.eval_inner(e);
        self.depth -= 1;
        v
    }

    fn eval_inner(&mut self, e: &Expr) -> f64 {
        let v = match e {
            Expr::Num(n) => *n,
            Expr::Eq(i) => self.equation(*i),
            Expr::Modifier(i) => self.modifiers.get(*i).copied().unwrap_or(0.0),
            Expr::Const(c) => match c {
                Constant::Pi => std::f64::consts::PI,
                Constant::Left => self.view.0,
                Constant::Top => self.view.1,
                Constant::Right => self.view.0 + self.view.2,
                Constant::Bottom => self.view.1 + self.view.3,
                Constant::XStretch | Constant::YStretch => 0.0,
                Constant::HasStroke => f64::from(u8::from(self.has_stroke)),
                Constant::HasFill => f64::from(u8::from(self.has_fill)),
                Constant::Width => self.view.2,
                Constant::Height => self.view.3,
                Constant::LogWidth => self.log.0,
                Constant::LogHeight => self.log.1,
            },
            Expr::Neg(inner) => -self.eval(inner),
            Expr::Bin(op, a, b) => {
                let a = self.eval(a);
                let b = self.eval(b);
                match op {
                    '+' => a + b,
                    '-' => a - b,
                    '*' => a * b,
                    _ => {
                        if b == 0.0 {
                            0.0
                        } else {
                            a / b
                        }
                    }
                }
            }
            Expr::Func(f, args) => {
                let x = self.eval(&args[0]);
                match f {
                    Func::Abs => x.abs(),
                    Func::Sqrt => x.max(0.0).sqrt(),
                    Func::Sin => x.sin(),
                    Func::Cos => x.cos(),
                    Func::Tan => x.tan(),
                    Func::Atan => x.atan(),
                    Func::Atan2 => x.atan2(self.eval(&args[1])),
                    Func::Min => x.min(self.eval(&args[1])),
                    Func::Max => x.max(self.eval(&args[1])),
                    Func::If => {
                        if x > 0.0 {
                            self.eval(&args[1])
                        } else {
                            self.eval(&args[2])
                        }
                    }
                }
            }
        };
        if v.is_finite() {
            v
        } else {
            0.0
        }
    }

    fn equation(&mut self, i: usize) -> f64 {
        if let Some(Some(v)) = self.values.get(i) {
            return *v;
        }
        if i >= self.exprs.len() || self.busy[i] || self.depth >= MAX_EVAL_DEPTH {
            // A cyclic reference evaluates to 0 (LibreOffice does the same);
            // so does a chain of references beyond the nesting bound.
            return 0.0;
        }
        self.busy[i] = true;
        let expr = self.exprs[i].take();
        let v = expr.as_ref().map(|e| self.eval(e)).unwrap_or(0.0);
        self.exprs[i] = expr;
        self.busy[i] = false;
        self.values[i] = Some(v);
        v
    }
}

#[derive(Clone, Debug)]
enum Param {
    Num(f64),
    Eq(usize),
    Modifier(usize),
    Const(Constant),
}

/// Evaluate a custom shape's enhanced geometry into shape-local EMU paths.
pub fn enhanced_geometry(
    node: roxmltree::Node<'_, '_>,
    frame: &ShapeFrame,
) -> Result<Geometry, String> {
    use crate::xml::{NS_DRAW, NS_SVG};
    let log = (frame.width / EMU_PER_HMM, frame.height / EMU_PER_HMM);
    let mut view = node
        .attribute((NS_SVG, "viewBox"))
        .and_then(parse_view_box)
        .unwrap_or((0.0, 0.0, 21600.0, 21600.0));
    // A zero-sized view box selects the shape's own logical size (1/100 mm),
    // as LibreOffice writes for shapes imported from DrawingML.
    if view.2 == 0.0 || view.3 == 0.0 {
        view = (0.0, 0.0, log.0, log.1);
    }
    let mut names: HashMap<&str, usize> = HashMap::new();
    let equations: Vec<_> = node
        .children()
        .filter(|c| crate::xml::is(*c, NS_DRAW, "equation"))
        .collect();
    for (i, eq) in equations.iter().enumerate() {
        if let Some(name) = eq.attribute((NS_DRAW, "name")) {
            names.insert(name, i);
        }
    }
    let exprs: Vec<Option<Expr>> = equations
        .iter()
        .map(|eq| {
            let formula = eq.attribute((NS_DRAW, "formula"))?;
            let mut p = FormulaParser {
                bytes: formula.as_bytes(),
                pos: 0,
                names: &names,
                depth: 0,
                nodes: 0,
            };
            let e = p.expr()?;
            (p.peek().is_none()).then_some(e)
        })
        .collect();
    let modifiers: Vec<f64> = node
        .attribute((NS_DRAW, "modifiers"))
        .map(|m| {
            m.split_whitespace()
                .map(|s| {
                    s.parse::<f64>()
                        .ok()
                        .filter(|v| v.is_finite())
                        .unwrap_or(0.0)
                })
                .collect()
        })
        .unwrap_or_default();
    let count = exprs.len();
    let mut ev = Evaluator {
        exprs,
        values: vec![None; count],
        busy: vec![false; count],
        modifiers,
        view,
        log,
        has_fill: frame.has_fill,
        has_stroke: frame.has_stroke,
        work: 0,
        depth: 0,
    };
    let path = node
        .attribute((NS_DRAWOOO, "enhanced-path"))
        .or_else(|| node.attribute((NS_DRAW, "enhanced-path")))
        .unwrap_or("");
    let sub_views: Vec<f64> = node
        .attribute((NS_DRAWOOO, "sub-view-size"))
        .map(|s| {
            s.split_whitespace()
                .filter_map(|v| v.parse().ok())
                .collect()
        })
        .unwrap_or_default();
    let tokens =
        tokenize_path(path, &names).ok_or_else(|| invalid("an enhanced path is malformed"))?;
    let mut geometry = Geometry::default();
    let base_scale = (
        if view.2 != 0.0 {
            frame.width / view.2
        } else {
            0.0
        },
        if view.3 != 0.0 {
            frame.height / view.3
        } else {
            0.0
        },
    );
    build_paths(
        &tokens,
        &mut ev,
        view,
        base_scale,
        &sub_views,
        frame,
        &mut geometry,
    )?;
    if let Some(areas) = node.attribute((NS_DRAW, "text-areas")) {
        if let Some(params) = tokenize_path(areas, &names) {
            let nums: Vec<f64> = params
                .iter()
                .filter_map(|t| match t {
                    PathToken::Param(p) => Some(eval_param(p, &mut ev)),
                    _ => None,
                })
                .collect();
            if nums.len() >= 4 {
                let x = |v: f64| (v - view.0) * base_scale.0;
                let y = |v: f64| (v - view.1) * base_scale.1;
                let (l, t, r, b) = (x(nums[0]), y(nums[1]), x(nums[2]), y(nums[3]));
                geometry.text_area = Some((l.min(r), t.min(b), (r - l).abs(), (b - t).abs()));
            }
        }
    }
    Ok(geometry)
}

#[derive(Debug)]
enum PathToken {
    Cmd(char),
    Param(Param),
}

fn tokenize_path(path: &str, names: &HashMap<&str, usize>) -> Option<Vec<PathToken>> {
    let bytes = path.as_bytes();
    let mut pos = 0;
    let mut out = Vec::new();
    while pos < bytes.len() {
        let b = bytes[pos];
        if b.is_ascii_whitespace() || b == b',' {
            pos += 1;
            continue;
        }
        if b == b'?' {
            pos += 1;
            let start = pos;
            while pos < bytes.len() && (bytes[pos].is_ascii_alphanumeric() || bytes[pos] == b'_') {
                pos += 1;
            }
            let name = std::str::from_utf8(&bytes[start..pos]).ok()?;
            // An unknown equation evaluates to 0 (LibreOffice ignores it).
            out.push(PathToken::Param(
                names.get(name).map_or(Param::Num(0.0), |i| Param::Eq(*i)),
            ));
            continue;
        }
        if b == b'$' {
            pos += 1;
            let start = pos;
            while pos < bytes.len() && bytes[pos].is_ascii_digit() {
                pos += 1;
            }
            let n: usize = std::str::from_utf8(&bytes[start..pos]).ok()?.parse().ok()?;
            out.push(PathToken::Param(Param::Modifier(n)));
            continue;
        }
        if b.is_ascii_digit() || b == b'.' || b == b'-' || b == b'+' {
            let negative = b == b'-';
            if b == b'-' || b == b'+' {
                pos += 1;
            }
            let v = number(bytes, &mut pos)?;
            out.push(PathToken::Param(Param::Num(if negative { -v } else { v })));
            continue;
        }
        if b.is_ascii_alphabetic() {
            let start = pos;
            while pos < bytes.len() && bytes[pos].is_ascii_alphabetic() {
                pos += 1;
            }
            let word = std::str::from_utf8(&bytes[start..pos]).ok()?;
            let constant = match word {
                "pi" => Some(Constant::Pi),
                "left" => Some(Constant::Left),
                "top" => Some(Constant::Top),
                "right" => Some(Constant::Right),
                "bottom" => Some(Constant::Bottom),
                "xstretch" => Some(Constant::XStretch),
                "ystretch" => Some(Constant::YStretch),
                "hasstroke" => Some(Constant::HasStroke),
                "hasfill" => Some(Constant::HasFill),
                "width" => Some(Constant::Width),
                "height" => Some(Constant::Height),
                "logwidth" => Some(Constant::LogWidth),
                "logheight" => Some(Constant::LogHeight),
                _ => None,
            };
            if let Some(c) = constant {
                out.push(PathToken::Param(Param::Const(c)));
            } else {
                // Commands are single letters and may be written together.
                for c in word.chars() {
                    out.push(PathToken::Cmd(c));
                }
            }
            continue;
        }
        return None;
    }
    if out.len() > MAX_SEGMENTS_PER_SHAPE * 8 {
        return None;
    }
    Some(out)
}

fn eval_param(p: &Param, ev: &mut Evaluator) -> f64 {
    match p {
        Param::Num(n) => *n,
        Param::Eq(i) => ev.equation(*i),
        Param::Modifier(i) => ev.modifiers.get(*i).copied().unwrap_or(0.0),
        Param::Const(c) => ev.eval(&Expr::Const(*c)),
    }
}

/// Build the painted paths. Each `N` ends one painted path (its subpaths
/// fill together); `F`/`S` suppress fill/stroke of the current one.
fn build_paths(
    tokens: &[PathToken],
    ev: &mut Evaluator,
    view: (f64, f64, f64, f64),
    base_scale: (f64, f64),
    sub_views: &[f64],
    frame: &ShapeFrame,
    geometry: &mut Geometry,
) -> Result<(), String> {
    let mut i = 0usize;
    let mut path_index = 0usize;
    let mut current = GeoPath {
        fill: true,
        stroke: true,
        ..GeoPath::default()
    };
    let mut pen: Option<(f64, f64)> = None;
    let mut start: (f64, f64) = (0.0, 0.0);
    let mut command = 'M';
    let mut quadrant_x = true;
    // Segments of the paths already finished (running total for the bound).
    let mut finished = 0usize;
    let scale_for = |index: usize| -> (f64, f64, f64, f64) {
        if let (Some(&w), Some(&h)) = (sub_views.get(index * 2), sub_views.get(index * 2 + 1)) {
            if w > 0.0 && h > 0.0 {
                return (0.0, 0.0, frame.width / w, frame.height / h);
            }
        }
        (view.0, view.1, base_scale.0, base_scale.1)
    };
    let mut scale = scale_for(0);
    macro_rules! params {
        ($n:expr) => {{
            let mut values = [0.0f64; 8];
            let mut ok = true;
            for slot in values.iter_mut().take($n) {
                match tokens.get(i) {
                    Some(PathToken::Param(p)) => {
                        *slot = eval_param(p, ev);
                        i += 1;
                    }
                    _ => {
                        ok = false;
                        break;
                    }
                }
            }
            if ok {
                Some(values)
            } else {
                None
            }
        }};
    }
    while i < tokens.len() {
        let step_start = i;
        let cmd = match &tokens[i] {
            PathToken::Cmd(c) => {
                i += 1;
                *c
            }
            PathToken::Param(_) => {
                // Repeated parameter groups continue the last command; a
                // moveto continues as lineto. A command without parameters
                // (Z N F S H I J K) does not repeat: a stray parameter after
                // it is skipped, so every step consumes a token.
                if matches!(command, 'Z' | 'N' | 'F' | 'S' | 'H' | 'I' | 'J' | 'K') {
                    i += 1;
                    continue;
                }
                match command {
                    'M' => 'L',
                    'U' => 'T',
                    'B' => 'A',
                    'V' => 'W',
                    c => c,
                }
            }
        };

        command = cmd;
        let (ox, oy, sx, sy) = scale;
        let px = |v: f64| (v - ox) * sx;
        let py = |v: f64| (v - oy) * sy;
        match cmd {
            'M' => {
                let Some(v) = params!(2) else { break };
                let p = (px(v[0]), py(v[1]));
                current.segs.push(Seg::Move(p.0, p.1));
                pen = Some(p);
                start = p;
                quadrant_x = true;
            }
            'L' => {
                let Some(v) = params!(2) else { break };
                let p = (px(v[0]), py(v[1]));
                ensure_pen(&mut current, &mut pen, &mut start, p);
                current.segs.push(Seg::Line(p.0, p.1));
                pen = Some(p);
            }
            'C' => {
                let Some(v) = params!(6) else { break };
                let p = (px(v[4]), py(v[5]));
                ensure_pen(&mut current, &mut pen, &mut start, (px(v[0]), py(v[1])));
                current
                    .segs
                    .push(Seg::Cubic(px(v[0]), py(v[1]), px(v[2]), py(v[3]), p.0, p.1));
                pen = Some(p);
            }
            'Q' => {
                let Some(v) = params!(4) else { break };
                let p = (px(v[2]), py(v[3]));
                ensure_pen(&mut current, &mut pen, &mut start, (px(v[0]), py(v[1])));
                current.segs.push(Seg::Quad(px(v[0]), py(v[1]), p.0, p.1));
                pen = Some(p);
            }
            'Z' => {
                current.segs.push(Seg::Close);
                pen = Some(start);
            }
            'N' => {
                if !current.segs.is_empty() {
                    finished += current.segs.len();
                    geometry.paths.push(std::mem::take(&mut current));
                }
                current = GeoPath {
                    fill: true,
                    stroke: true,
                    ..GeoPath::default()
                };
                pen = None;
                path_index += 1;
                scale = scale_for(path_index);
            }
            'F' => current.fill = false,
            'S' => current.stroke = false,
            'H' => current.fill_mode = Some("darken"),
            'I' => current.fill_mode = Some("darkenLess"),
            'J' => current.fill_mode = Some("lighten"),
            'K' => current.fill_mode = Some("lightenLess"),
            'T' | 'U' => {
                let Some(v) = params!(6) else { break };
                let (cx, cy) = (px(v[0]), py(v[1]));
                let (rx, ry) = ((v[2] * sx).abs(), (v[3] * sy).abs());
                let (t0, t1) = (v[4], v[5]);
                let full = ((t0 as i64).rem_euclid(360)) == ((t1 as i64).rem_euclid(360))
                    && (t0.fract() == t1.fract());
                // Angles are counter-clockwise on the page; the parametric
                // y-down angle is their negation.
                let start_t = -t0.to_radians();
                let mut sweep = -(t1 - t0).to_radians();
                if full {
                    sweep = -std::f64::consts::TAU;
                } else if sweep > 0.0 {
                    sweep -= std::f64::consts::TAU;
                }
                let first = (cx + rx * start_t.cos(), cy + ry * start_t.sin());
                if cmd == 'U' || pen.is_none() {
                    current.segs.push(Seg::Move(first.0, first.1));
                    start = first;
                } else {
                    current.segs.push(Seg::Line(first.0, first.1));
                }
                arc_segments(cx, cy, rx, ry, start_t, sweep, &mut current.segs);
                let end = start_t + sweep;
                pen = Some((cx + rx * end.cos(), cy + ry * end.sin()));
            }
            'A' | 'B' | 'W' | 'V' => {
                let Some(v) = params!(8) else { break };
                let (x1, y1, x2, y2) = (px(v[0]), py(v[1]), px(v[2]), py(v[3]));
                let (cx, cy) = ((x1 + x2) / 2.0, (y1 + y2) / 2.0);
                let (rx, ry) = (((x2 - x1) / 2.0).abs(), ((y2 - y1) / 2.0).abs());
                let angle = |x: f64, y: f64| {
                    let dx = if rx > 0.0 { (x - cx) / rx } else { 0.0 };
                    let dy = if ry > 0.0 { (y - cy) / ry } else { 0.0 };
                    dy.atan2(dx)
                };
                let a0 = angle(px(v[4]), py(v[5]));
                let a1 = angle(px(v[6]), py(v[7]));
                let clockwise = cmd == 'W' || cmd == 'V';
                let mut sweep = a1 - a0;
                if clockwise {
                    if sweep <= 0.0 {
                        sweep += std::f64::consts::TAU;
                    }
                } else if sweep >= 0.0 {
                    sweep -= std::f64::consts::TAU;
                }
                let first = (cx + rx * a0.cos(), cy + ry * a0.sin());
                if cmd == 'B' || cmd == 'V' || pen.is_none() {
                    current.segs.push(Seg::Move(first.0, first.1));
                    start = first;
                } else {
                    current.segs.push(Seg::Line(first.0, first.1));
                }
                arc_segments(cx, cy, rx, ry, a0, sweep, &mut current.segs);
                let end = a0 + sweep;
                pen = Some((cx + rx * end.cos(), cy + ry * end.sin()));
            }
            'X' | 'Y' => {
                if matches!(tokens.get(i.wrapping_sub(1)), Some(PathToken::Cmd(_))) {
                    quadrant_x = cmd == 'X';
                }
                let Some(v) = params!(2) else { break };
                let p = (px(v[0]), py(v[1]));
                let s = pen.unwrap_or(p);
                ensure_pen(&mut current, &mut pen, &mut start, s);
                const K: f64 = 0.552_284_749_830_793_4;
                let seg = if quadrant_x {
                    Seg::Cubic(
                        s.0 + K * (p.0 - s.0),
                        s.1,
                        p.0,
                        p.1 + K * (s.1 - p.1),
                        p.0,
                        p.1,
                    )
                } else {
                    Seg::Cubic(
                        s.0,
                        s.1 + K * (p.1 - s.1),
                        p.0 + K * (s.0 - p.0),
                        p.1,
                        p.0,
                        p.1,
                    )
                };
                current.segs.push(seg);
                pen = Some(p);
                quadrant_x = !quadrant_x;
            }
            'G' => {
                let Some(v) = params!(4) else { break };
                let (rx, ry) = ((v[0] * sx).abs(), (v[1] * sy).abs());
                let (st, sw) = (v[2].to_radians(), v[3].to_radians());
                let p = pen.unwrap_or((0.0, 0.0));
                ensure_pen(&mut current, &mut pen, &mut start, p);
                let cx = p.0 - rx * st.cos();
                let cy = p.1 - ry * st.sin();
                arc_segments(cx, cy, rx, ry, st, sw, &mut current.segs);
                let end = st + sw;
                pen = Some((cx + rx * end.cos(), cy + ry * end.sin()));
            }
            _ => {
                // An unknown command ends the path (LibreOffice stops too).
                break;
            }
        }
        if finished + current.segs.len() > MAX_SEGMENTS_PER_SHAPE {
            return Err("OpenDocument shape exceeds the geometry limit".to_string());
        }
        // Every step consumes at least one token (its command letter or its
        // parameters); a step that did not ends the path.
        if i == step_start {
            break;
        }
    }
    if !current.segs.is_empty() {
        geometry.paths.push(current);
    }
    Ok(())
}

fn ensure_pen(
    path: &mut GeoPath,
    pen: &mut Option<(f64, f64)>,
    start: &mut (f64, f64),
    p: (f64, f64),
) {
    if pen.is_none() {
        path.segs.push(Seg::Move(p.0, p.1));
        *pen = Some(p);
        *start = p;
    }
}

// ---------------------------------------------------------------------------
// Path measurement and trimming (line ends under markers)
// ---------------------------------------------------------------------------

fn lerp(a: (f64, f64), b: (f64, f64), t: f64) -> (f64, f64) {
    (a.0 + (b.0 - a.0) * t, a.1 + (b.1 - a.1) * t)
}

/// Split a cubic at `t`, returning the two halves' control points.
type Cubic = [(f64, f64); 4];
fn split_cubic(c: Cubic, t: f64) -> (Cubic, Cubic) {
    let ab = lerp(c[0], c[1], t);
    let bc = lerp(c[1], c[2], t);
    let cd = lerp(c[2], c[3], t);
    let abc = lerp(ab, bc, t);
    let bcd = lerp(bc, cd, t);
    let m = lerp(abc, bcd, t);
    ([c[0], ab, abc, m], [m, bcd, cd, c[3]])
}

fn cubic_length(c: Cubic) -> f64 {
    let mut len = 0.0;
    let mut prev = c[0];
    for i in 1..=16 {
        let t = i as f64 / 16.0;
        let (_, rest) = split_cubic(c, t);
        let p = rest[0];
        len += (p.0 - prev.0).hypot(p.1 - prev.1);
        prev = p;
    }
    len
}

fn as_cubic(from: (f64, f64), seg: &Seg) -> Option<Cubic> {
    match *seg {
        Seg::Line(x, y) => Some([
            from,
            lerp(from, (x, y), 1.0 / 3.0),
            lerp(from, (x, y), 2.0 / 3.0),
            (x, y),
        ]),
        Seg::Cubic(a, b, c, d, x, y) => Some([from, (a, b), (c, d), (x, y)]),
        Seg::Quad(a, b, x, y) => Some([
            from,
            (
                from.0 + 2.0 / 3.0 * (a - from.0),
                from.1 + 2.0 / 3.0 * (b - from.1),
            ),
            (x + 2.0 / 3.0 * (a - x), y + 2.0 / 3.0 * (b - y)),
            (x, y),
        ]),
        _ => None,
    }
}

/// Shorten an open subpath at its start by `len` along the first drawing
/// segment. Returns false when the subpath is too short.
pub fn trim_start(segs: &mut [Seg], len: f64) {
    if len <= 0.0 || segs.len() < 2 {
        return;
    }
    let Some(Seg::Move(x0, y0)) = segs.first().copied() else {
        return;
    };
    let seg = segs[1];
    match seg {
        Seg::Line(x, y) => {
            let total = (x - x0).hypot(y - y0);
            if total <= 0.0 {
                return;
            }
            let t = (len / total).min(1.0);
            let p = lerp((x0, y0), (x, y), t);
            segs[0] = Seg::Move(p.0, p.1);
        }
        _ => {
            let Some(c) = as_cubic((x0, y0), &seg) else {
                return;
            };
            let total = cubic_length(c);
            if total <= 0.0 {
                return;
            }
            let (_, rest) = split_cubic(c, (len / total).min(1.0));
            segs[0] = Seg::Move(rest[0].0, rest[0].1);
            segs[1] = Seg::Cubic(
                rest[1].0, rest[1].1, rest[2].0, rest[2].1, rest[3].0, rest[3].1,
            );
        }
    }
}

/// Shorten an open subpath at its end by `len`.
pub fn trim_end(segs: &mut [Seg], len: f64) {
    if len <= 0.0 || segs.len() < 2 {
        return;
    }
    let n = segs.len();
    let Some(from) = segs[n - 2].end() else {
        return;
    };
    let seg = segs[n - 1];
    match seg {
        Seg::Line(x, y) => {
            let total = (x - from.0).hypot(y - from.1);
            if total <= 0.0 {
                return;
            }
            let t = 1.0 - (len / total).min(1.0);
            let p = lerp(from, (x, y), t);
            segs[n - 1] = Seg::Line(p.0, p.1);
        }
        _ => {
            let Some(c) = as_cubic(from, &seg) else {
                return;
            };
            let total = cubic_length(c);
            if total <= 0.0 {
                return;
            }
            let (head, _) = split_cubic(c, 1.0 - (len / total).min(1.0));
            segs[n - 1] = Seg::Cubic(
                head[1].0, head[1].1, head[2].0, head[2].1, head[3].0, head[3].1,
            );
        }
    }
}

/// The start point of an open subpath and the outward direction there
/// (unit vector pointing away from the path).
pub fn start_tangent(segs: &[Seg]) -> Option<((f64, f64), (f64, f64))> {
    let Seg::Move(x0, y0) = *segs.first()? else {
        return None;
    };
    for seg in &segs[1..] {
        let toward = match *seg {
            Seg::Line(x, y) => (x, y),
            Seg::Cubic(a, b, c, d, x, y) => [(a, b), (c, d), (x, y)]
                .into_iter()
                .find(|p| (p.0 - x0).hypot(p.1 - y0) > 1e-6)
                .unwrap_or((x, y)),
            Seg::Quad(a, b, x, y) => {
                if (a - x0).hypot(b - y0) > 1e-6 {
                    (a, b)
                } else {
                    (x, y)
                }
            }
            _ => continue,
        };
        let (dx, dy) = (x0 - toward.0, y0 - toward.1);
        let len = dx.hypot(dy);
        if len > 1e-6 {
            return Some(((x0, y0), (dx / len, dy / len)));
        }
    }
    None
}

pub fn end_tangent(segs: &[Seg]) -> Option<((f64, f64), (f64, f64))> {
    let n = segs.len();
    if n < 2 {
        return None;
    }
    let end = segs[n - 1].end()?;
    let prev_end = segs[n - 2].end();
    let toward = match segs[n - 1] {
        Seg::Cubic(a, b, c, d, _, _) => [(c, d), (a, b)]
            .into_iter()
            .chain(prev_end)
            .find(|p| (p.0 - end.0).hypot(p.1 - end.1) > 1e-6)?,
        Seg::Quad(a, b, _, _) => [(a, b)]
            .into_iter()
            .chain(prev_end)
            .find(|p| (p.0 - end.0).hypot(p.1 - end.1) > 1e-6)?,
        _ => prev_end?,
    };
    let (dx, dy) = (end.0 - toward.0, end.1 - toward.1);
    let len = dx.hypot(dy);
    (len > 1e-6).then(|| (end, (dx / len, dy / len)))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn doc(xml: &str) -> roxmltree::Document<'_> {
        roxmltree::Document::parse(xml).unwrap()
    }

    const NS: &str = r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:drawooo="http://openoffice.org/2010/draw""#;

    fn frame() -> ShapeFrame {
        ShapeFrame {
            width: 360_000.0,
            height: 360_000.0,
            has_fill: true,
            has_stroke: true,
        }
    }

    #[test]
    fn parameters_after_parameterless_commands_are_skipped() {
        for path in ["F 0", "M 0 0 L 1 1 N 5 6 7", "M 0 0 L 10 10 Z 1 2 S 3"] {
            let xml = format!(r#"<draw:enhanced-geometry {NS} draw:enhanced-path="{path}"/>"#);
            let d = doc(&xml);
            let g = enhanced_geometry(d.root_element(), &frame()).unwrap();
            assert!(g.paths.iter().all(|p| p.segs.len() <= 3), "{path}: {g:?}");
        }
    }

    #[test]
    fn long_equation_chains_and_flat_sums_stay_bounded() {
        // A chain of references deeper than the evaluation bound, and a flat
        // sum longer than the node bound: both evaluate (to bounded values)
        // without exhausting the stack.
        let mut equations = String::from(r#"<draw:equation draw:name="f0" draw:formula="1"/>"#);
        for i in 1..5000 {
            equations.push_str(&format!(
                r#"<draw:equation draw:name="f{i}" draw:formula="?f{} +1"/>"#,
                i - 1
            ));
        }
        let flat = vec!["1"; 100_000].join("+");
        equations.push_str(&format!(
            r#"<draw:equation draw:name="flat" draw:formula="{flat}"/>"#
        ));
        let xml = format!(
            r#"<draw:enhanced-geometry {NS} draw:enhanced-path="M 0 0 L ?f4999 ?flat N">{equations}</draw:enhanced-geometry>"#
        );
        let d = doc(&xml);
        let g = enhanced_geometry(d.root_element(), &frame()).unwrap();
        let Seg::Line(x, y) = g.paths[0].segs[1] else {
            panic!("{g:?}");
        };
        assert!(x.is_finite() && x < 5000.0 * 360_000.0 / 21_600.0, "{x}");
        // The over-long sum does not parse and evaluates to 0.
        assert_eq!(y, 0.0);
    }

    #[test]
    fn many_short_paths_are_counted_incrementally() {
        let path = "M 0 0 L 1 1 N ".repeat(60_000);
        let xml = format!(r#"<draw:enhanced-geometry {NS} draw:enhanced-path="{path}"/>"#);
        let d = doc(&xml);
        assert!(enhanced_geometry(d.root_element(), &frame()).is_ok());
        let path = "M 0 0 L 1 1 N ".repeat(110_000);
        let xml = format!(r#"<draw:enhanced-geometry {NS} draw:enhanced-path="{path}"/>"#);
        let d = doc(&xml);
        assert!(enhanced_geometry(d.root_element(), &frame()).is_err());
    }

    #[test]
    fn transforms_compose_left_to_right_with_counter_clockwise_rotation() {
        // A 2×1 cm box rotated 90° clockwise (ODF −π/2) and placed.
        let m = parse_transform("rotate (-1.5707963267949) translate (3cm 4cm)").unwrap();
        let p = place(720_000.0, 360_000.0, &m);
        assert!((p.rot - 90.0).abs() < 1e-6, "{p:?}");
        // The centre (1cm, 0.5cm) maps to (−0.5cm, 1cm) + (3cm, 4cm).
        assert!((p.x + p.w / 2.0 - 2.5 * 360_000.0).abs() < 1.0);
        assert!((p.y + p.h / 2.0 - 5.0 * 360_000.0).abs() < 1.0);
        assert!(!p.flip_v);
    }

    #[test]
    fn evaluates_formulas_and_draws_ooxml_arcs() {
        let xml = format!(
            r#"<draw:enhanced-geometry {NS} svg:viewBox="0 0 0 0" draw:modifiers="25000" draw:text-areas="0 0 ?f1 logheight" drawooo:enhanced-path="M 0 ?f0 G ?f0 ?f0 180 90 L ?f1 0 Z N" draw:enhanced-path="M 0 0 Z N"><draw:equation draw:name="f0" draw:formula="min(logwidth,logheight)*$0 /100000"/><draw:equation draw:name="f1" draw:formula="logwidth-?f0 "/></draw:enhanced-geometry>"#
        );
        let d = doc(&xml);
        // 4×2 cm shape: logwidth 4000, logheight 2000, f0 = 500 (0.5 cm).
        let g = enhanced_geometry(
            d.root_element(),
            &ShapeFrame {
                width: 1_440_000.0,
                height: 720_000.0,
                has_fill: true,
                has_stroke: true,
            },
        )
        .unwrap();
        assert_eq!(g.paths.len(), 1);
        let segs = &g.paths[0].segs;
        assert_eq!(segs[0], Seg::Move(0.0, 180_000.0));
        // The quarter arc ends at (r, 0): the top-left rounded corner.
        let Seg::Cubic(_, _, _, _, x, y) = segs[1] else {
            panic!("{segs:?}")
        };
        assert!((x - 180_000.0).abs() < 1e-6 && y.abs() < 1e-6);
        assert_eq!(segs[2], Seg::Line(1_260_000.0, 0.0));
        let area = g.text_area.unwrap();
        assert!((area.2 - 1_260_000.0).abs() < 1e-6 && (area.3 - 720_000.0).abs() < 1e-6);
    }

    #[test]
    fn view_box_scaling_subpath_flags_and_angle_ellipse() {
        let xml = format!(
            r#"<draw:enhanced-geometry {NS} svg:viewBox="0 0 21600 21600" draw:enhanced-path="U 10800 10800 10800 10800 0 360 Z N F M 0 0 L 21600 21600 N"/>"#
        );
        let d = doc(&xml);
        let g = enhanced_geometry(
            d.root_element(),
            &ShapeFrame {
                width: 2000.0,
                height: 1000.0,
                has_fill: true,
                has_stroke: true,
            },
        )
        .unwrap();
        assert_eq!(g.paths.len(), 2);
        assert_eq!(g.paths[0].segs[0], Seg::Move(2000.0, 500.0));
        let (x0, y0, x1, y1) = Geometry {
            paths: vec![g.paths[0].clone()],
            text_area: None,
        }
        .bounds()
        .unwrap();
        assert!(
            x0.abs() < 1.0
                && y0.abs() < 1.0
                && (x1 - 2000.0).abs() < 1.0
                && (y1 - 1000.0).abs() < 1.0
        );
        assert!(!g.paths[1].fill && g.paths[1].stroke);
        assert_eq!(g.paths[1].segs[1], Seg::Line(2000.0, 1000.0));
    }

    #[test]
    fn parses_svg_paths_with_relative_and_arc_commands() {
        let segs = parse_svg_path("M10 0l-10 30h20z").unwrap();
        assert_eq!(
            segs,
            vec![
                Seg::Move(10.0, 0.0),
                Seg::Line(0.0, 30.0),
                Seg::Line(20.0, 30.0),
                Seg::Close
            ]
        );
        let arc = parse_svg_path("M0 0 A10 10 0 0 1 20 0").unwrap();
        let Seg::Cubic(_, _, _, _, x, y) = *arc.last().unwrap() else {
            panic!()
        };
        assert!((x - 20.0).abs() < 1e-9 && y.abs() < 1e-9);
        assert!(parse_svg_path("M0 0 L").is_none());
    }

    #[test]
    fn trims_open_path_ends() {
        let mut segs = vec![Seg::Move(0.0, 0.0), Seg::Line(10.0, 0.0)];
        trim_start(&mut segs, 2.0);
        trim_end(&mut segs, 3.0);
        assert_eq!(segs, vec![Seg::Move(2.0, 0.0), Seg::Line(7.0, 0.0)]);
        let (p, d) = start_tangent(&[Seg::Move(0.0, 0.0), Seg::Line(0.0, 5.0)]).unwrap();
        assert_eq!((p, d), ((0.0, 0.0), (0.0, -1.0)));
    }
}
