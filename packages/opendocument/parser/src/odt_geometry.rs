//! Custom-shape geometry (ODF Part 3 §10.6.2 `draw:enhanced-geometry`):
//! `draw:equation` formulas (§19.171), modifiers (§19.196) and the enhanced
//! path (§19.145, Table 10), evaluated into normalized [0, 1] DOCX path
//! commands. Arcs become cubic Béziers of at most 90°.

use crate::budget;
use crate::error::resource;
use docx_model::{PathCmd, PathPaint};
use std::collections::HashMap;

/// Implementation bounds: equation definitions of one shape (enforced where
/// they are collected), path parameters and formula depth. A shape beyond
/// the first two is rejected.
pub(super) const MAX_EQUATIONS: usize = 4096;
const MAX_PATH_TOKENS: usize = 1 << 20;
const MAX_DEPTH: usize = 64;
/// Implementation bound on the recursion of one evaluation: referenced
/// equations, parentheses and function arguments all count, so no formula
/// shape nests the parser deeper than this.
const MAX_NESTING: usize = 128;

pub(super) struct Geometry {
    pub view_box: (f64, f64, f64, f64),
    pub equations: HashMap<String, String>,
    pub modifiers: Vec<f64>,
    /// Shape size in 1/100 mm (logwidth/logheight).
    pub log_size: (f64, f64),
    pub has_fill: bool,
    pub has_stroke: bool,
}

/// A formula or path parameter evaluator.
pub(super) struct Evaluator<'g> {
    geometry: &'g Geometry,
    cache: HashMap<String, f64>,
    depth: usize,
    /// Current parser recursion (equations, parentheses, arguments).
    nesting: usize,
    /// Equations being evaluated: a reference back to one is a cycle and
    /// reads as 0, so no equation is evaluated twice at once.
    active: std::collections::HashSet<&'g str>,
}

impl<'g> Evaluator<'g> {
    pub fn new(geometry: &'g Geometry) -> Self {
        Self {
            geometry,
            cache: HashMap::new(),
            depth: 0,
            nesting: 0,
            active: std::collections::HashSet::new(),
        }
    }

    fn identifier(&self, name: &str) -> Option<f64> {
        let (x, y, w, h) = self.geometry.view_box;
        Some(match name {
            "pi" => std::f64::consts::PI,
            "left" => x,
            "top" => y,
            "right" => x + w,
            "bottom" => y + h,
            "width" => w,
            "height" => h,
            "xstretch" | "ystretch" => 0.0,
            "hasfill" => f64::from(u8::from(self.geometry.has_fill)),
            "hasstroke" => f64::from(u8::from(self.geometry.has_stroke)),
            "logwidth" => self.geometry.log_size.0,
            "logheight" => self.geometry.log_size.1,
            _ => return None,
        })
    }

    pub fn equation(&mut self, name: &str) -> f64 {
        if let Some(value) = self.cache.get(name) {
            return *value;
        }
        if self.depth > MAX_DEPTH {
            return 0.0;
        }
        // The formula is borrowed from the geometry (never copied per
        // evaluation).
        let geometry = self.geometry;
        let Some((key, formula)) = geometry.equations.get_key_value(name) else {
            return 0.0;
        };
        if !self.active.insert(key.as_str()) {
            return 0.0;
        }
        self.depth += 1;
        let mut parser = Parser {
            text: formula.as_bytes(),
            at: 0,
        };
        let value = parser.additive(self);
        self.depth -= 1;
        self.active.remove(key.as_str());
        let value = if value.is_finite() { value } else { 0.0 };
        self.cache.insert(name.to_string(), value);
        value
    }

    fn modifier(&self, index: usize) -> f64 {
        self.geometry.modifiers.get(index).copied().unwrap_or(0.0)
    }

    /// A path parameter: a number, `?name` or `$index`.
    pub fn parameter(&mut self, token: &str) -> f64 {
        if let Some(name) = token.strip_prefix('?') {
            return self.equation(name);
        }
        if let Some(index) = token.strip_prefix('$') {
            return index.parse::<usize>().map_or(0.0, |i| self.modifier(i));
        }
        if let Some(value) = self.identifier(token) {
            return value;
        }
        token.parse::<f64>().unwrap_or(0.0)
    }
}

struct Parser<'t> {
    text: &'t [u8],
    at: usize,
}

impl Parser<'_> {
    fn skip(&mut self) {
        while self.at < self.text.len() && matches!(self.text[self.at], b' ' | b'\t') {
            self.at += 1;
        }
    }

    fn peek(&mut self) -> Option<u8> {
        self.skip();
        self.text.get(self.at).copied()
    }

    fn additive(&mut self, eval: &mut Evaluator<'_>) -> f64 {
        let mut value = self.multiplicative(eval);
        while let Some(op @ (b'+' | b'-')) = self.peek() {
            self.at += 1;
            let rhs = self.multiplicative(eval);
            value = if op == b'+' { value + rhs } else { value - rhs };
        }
        value
    }

    fn multiplicative(&mut self, eval: &mut Evaluator<'_>) -> f64 {
        let mut value = self.unary(eval);
        while let Some(op @ (b'*' | b'/')) = self.peek() {
            self.at += 1;
            let rhs = self.unary(eval);
            value = if op == b'*' {
                value * rhs
            } else if rhs == 0.0 {
                0.0
            } else {
                value / rhs
            };
        }
        value
    }

    fn unary(&mut self, eval: &mut Evaluator<'_>) -> f64 {
        if self.peek() == Some(b'-') {
            self.at += 1;
            return -self.basic(eval);
        }
        self.basic(eval)
    }

    fn arguments(&mut self, eval: &mut Evaluator<'_>, count: usize) -> Vec<f64> {
        let mut out = Vec::with_capacity(count);
        if self.peek() != Some(b'(') {
            return vec![0.0; count];
        }
        self.at += 1;
        for index in 0..count {
            out.push(self.additive(eval));
            if index + 1 < count && self.peek() == Some(b',') {
                self.at += 1;
            }
        }
        if self.peek() == Some(b')') {
            self.at += 1;
        }
        out
    }

    /// One primary expression. Every recursive path of the grammar passes
    /// through here, so the nesting bound covers parentheses, function
    /// arguments, unary minus chains and referenced equations alike; past
    /// the bound the rest of the formula is dropped and reads as 0.
    fn basic(&mut self, eval: &mut Evaluator<'_>) -> f64 {
        if eval.nesting >= MAX_NESTING {
            self.at = self.text.len();
            return 0.0;
        }
        eval.nesting += 1;
        let value = self.primary(eval);
        eval.nesting -= 1;
        value
    }

    fn primary(&mut self, eval: &mut Evaluator<'_>) -> f64 {
        match self.peek() {
            Some(b'(') => {
                self.at += 1;
                let value = self.additive(eval);
                if self.peek() == Some(b')') {
                    self.at += 1;
                }
                value
            }
            Some(b'?') => {
                self.at += 1;
                let start = self.at;
                while self.at < self.text.len()
                    && !matches!(
                        self.text[self.at],
                        b' ' | b'\t' | b'+' | b'-' | b'*' | b'/' | b')' | b','
                    )
                {
                    self.at += 1;
                }
                let name = String::from_utf8_lossy(&self.text[start..self.at]).to_string();
                eval.equation(&name)
            }
            Some(b'$') => {
                self.at += 1;
                let start = self.at;
                while self.at < self.text.len() && self.text[self.at].is_ascii_digit() {
                    self.at += 1;
                }
                std::str::from_utf8(&self.text[start..self.at])
                    .ok()
                    .and_then(|v| v.parse::<usize>().ok())
                    .map_or(0.0, |i| eval.modifier(i))
            }
            Some(c) if c.is_ascii_digit() || c == b'.' || c == b'+' => {
                let start = self.at;
                if c == b'+' {
                    self.at += 1;
                }
                while self.at < self.text.len()
                    && (self.text[self.at].is_ascii_digit() || self.text[self.at] == b'.')
                {
                    self.at += 1;
                }
                if self.at < self.text.len() && matches!(self.text[self.at], b'e' | b'E') {
                    self.at += 1;
                    if self.at < self.text.len() && matches!(self.text[self.at], b'+' | b'-') {
                        self.at += 1;
                    }
                    while self.at < self.text.len() && self.text[self.at].is_ascii_digit() {
                        self.at += 1;
                    }
                }
                std::str::from_utf8(&self.text[start..self.at])
                    .ok()
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(0.0)
            }
            Some(c) if c.is_ascii_alphabetic() => {
                let start = self.at;
                while self.at < self.text.len() && self.text[self.at].is_ascii_alphanumeric() {
                    self.at += 1;
                }
                let name = String::from_utf8_lossy(&self.text[start..self.at]).to_string();
                match name.as_str() {
                    "abs" | "sqrt" | "sin" | "cos" | "tan" | "atan" => {
                        let a = self.arguments(eval, 1)[0];
                        match name.as_str() {
                            "abs" => a.abs(),
                            "sqrt" => a.max(0.0).sqrt(),
                            "sin" => a.sin(),
                            "cos" => a.cos(),
                            "tan" => a.tan(),
                            _ => a.atan(),
                        }
                    }
                    "min" | "max" | "atan2" => {
                        let args = self.arguments(eval, 2);
                        match name.as_str() {
                            "min" => args[0].min(args[1]),
                            "max" => args[0].max(args[1]),
                            _ => args[0].atan2(args[1]),
                        }
                    }
                    "if" => {
                        let args = self.arguments(eval, 3);
                        if args[0] > 0.0 {
                            args[1]
                        } else {
                            args[2]
                        }
                    }
                    other => eval.identifier(other).unwrap_or(0.0),
                }
            }
            _ => {
                self.at = self.text.len();
                0.0
            }
        }
    }
}

/// Tokens of an enhanced path: command letters and parameters.
fn tokens(path: &str) -> Result<Vec<String>, String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let flush = |current: &mut String, out: &mut Vec<String>| {
        if !current.is_empty() {
            out.push(std::mem::take(current));
        }
    };
    for c in path.chars() {
        if out.len() > MAX_PATH_TOKENS {
            return Err(resource("shape path"));
        }
        match c {
            ' ' | ',' | '\t' | '\n' | '\r' => flush(&mut current, &mut out),
            'A'..='Z' if current.is_empty() || !current.starts_with('?') => {
                flush(&mut current, &mut out);
                out.push(c.to_string());
            }
            '?' | '$' => {
                flush(&mut current, &mut out);
                current.push(c);
            }
            '-' if !current.is_empty()
                && !current.starts_with('?')
                && !current.ends_with(['e', 'E']) =>
            {
                flush(&mut current, &mut out);
                current.push(c);
            }
            _ => current.push(c),
        }
    }
    flush(&mut current, &mut out);
    Ok(out)
}

/// The evaluated path: normalized sub-paths and their paint.
pub(super) struct Path {
    pub subpaths: Vec<Vec<PathCmd>>,
    pub paint: Vec<PathPaint>,
}

struct Builder {
    view_box: (f64, f64, f64, f64),
    subpaths: Vec<Vec<PathCmd>>,
    paint: Vec<PathPaint>,
    current: Vec<PathCmd>,
    point: (f64, f64),
    start: (f64, f64),
    fill: bool,
    stroke: bool,
    /// Sub-paths of the current set (one N-terminated group).
    group_start: usize,
}

impl Builder {
    fn norm(&self, (x, y): (f64, f64)) -> (f64, f64) {
        let (vx, vy, vw, vh) = self.view_box;
        (
            if vw != 0.0 { (x - vx) / vw } else { 0.0 },
            if vh != 0.0 { (y - vy) / vh } else { 0.0 },
        )
    }

    fn finish_subpath(&mut self) {
        if self.current.len() > 1 {
            self.subpaths.push(std::mem::take(&mut self.current));
            self.paint.push(PathPaint {
                fill_rule: None,
                fill: None,
                stroke: true,
            });
        } else {
            self.current.clear();
        }
    }

    fn move_to(&mut self, p: (f64, f64)) {
        self.finish_subpath();
        let (x, y) = self.norm(p);
        self.current.push(PathCmd::MoveTo { x, y });
        self.point = p;
        self.start = p;
    }

    fn line_to(&mut self, p: (f64, f64)) {
        if self.current.is_empty() {
            self.move_to(self.point);
        }
        let (x, y) = self.norm(p);
        self.current.push(PathCmd::LineTo { x, y });
        self.point = p;
    }

    fn cubic(&mut self, c1: (f64, f64), c2: (f64, f64), p: (f64, f64)) {
        if self.current.is_empty() {
            self.move_to(self.point);
        }
        let (x1, y1) = self.norm(c1);
        let (x2, y2) = self.norm(c2);
        let (x, y) = self.norm(p);
        self.current.push(PathCmd::CubicBezTo {
            x1,
            y1,
            x2,
            y2,
            x,
            y,
        });
        self.point = p;
    }

    fn close(&mut self) {
        if !self.current.is_empty() {
            self.current.push(PathCmd::Close);
            self.point = self.start;
        }
    }

    /// End of a path set (N): apply the set's fill and stroke flags.
    fn end_set(&mut self) {
        self.finish_subpath();
        for paint in &mut self.paint[self.group_start..] {
            if !self.fill {
                paint.fill = Some("none".into());
            }
            paint.stroke = self.stroke;
        }
        self.group_start = self.paint.len();
        self.fill = true;
        self.stroke = true;
    }

    /// An elliptical arc around `center` with radii `r`, from parameter
    /// angle `start` sweeping `sweep` (radians, screen coordinates: positive
    /// is clockwise), as cubic Béziers.
    fn arc(&mut self, center: (f64, f64), r: (f64, f64), start: f64, sweep: f64, line: bool) {
        let point = |a: f64| (center.0 + r.0 * a.cos(), center.1 + r.1 * a.sin());
        let first = point(start);
        if line && !self.current.is_empty() {
            self.line_to(first);
        } else {
            self.move_to(first);
        }
        let segments = (sweep.abs() / std::f64::consts::FRAC_PI_2).ceil().max(1.0) as usize;
        let step = sweep / segments as f64;
        let k = 4.0 / 3.0 * (step / 4.0).tan();
        let mut a = start;
        for _ in 0..segments.min(64) {
            let b = a + step;
            let p0 = point(a);
            let p1 = point(b);
            let c1 = (p0.0 - k * r.0 * a.sin(), p0.1 + k * r.1 * a.cos());
            let c2 = (p1.0 + k * r.0 * b.sin(), p1.1 - k * r.1 * b.cos());
            self.cubic(c1, c2, p1);
            a = b;
        }
    }
}

/// Evaluate an enhanced path (Table 10) into normalized commands. Each
/// token takes a step and adds a bounded number of commands (an arc is at
/// most 64 curves).
pub(super) fn evaluate(path: &str, geometry: &Geometry) -> Result<Path, String> {
    let mut eval = Evaluator::new(geometry);
    let tokens = tokens(path)?;
    let mut builder = Builder {
        view_box: geometry.view_box,
        subpaths: Vec::new(),
        paint: Vec::new(),
        current: Vec::new(),
        point: (0.0, 0.0),
        start: (0.0, 0.0),
        fill: true,
        stroke: true,
        group_start: 0,
    };
    let mut command = 'M';
    let mut index = 0;
    // X/Y alternate their tangent axis per quarter.
    let mut quadrant_x = true;
    while index < tokens.len() {
        budget::tick(1)?;
        let token = &tokens[index];
        if token.len() == 1 && token.chars().next().is_some_and(|c| c.is_ascii_uppercase()) {
            command = token.chars().next().unwrap_or('M');
            index += 1;
            match command {
                'Z' => {
                    builder.close();
                    continue;
                }
                'N' => {
                    builder.end_set();
                    continue;
                }
                'F' => {
                    builder.fill = false;
                    continue;
                }
                'S' => {
                    builder.stroke = false;
                    continue;
                }
                'H' | 'I' | 'J' | 'K' => continue,
                'X' => quadrant_x = true,
                'Y' => quadrant_x = false,
                _ => {}
            }
        }
        let arity = match command {
            'M' | 'L' | 'X' | 'Y' => 2,
            'C' => 6,
            'Q' => 4,
            'T' | 'U' => 6,
            'A' | 'B' | 'V' | 'W' => 8,
            'G' => 4,
            _ => {
                index += 1;
                continue;
            }
        };
        if index + arity > tokens.len() {
            break;
        }
        let values: Vec<f64> = tokens[index..index + arity]
            .iter()
            .map(|t| eval.parameter(t))
            .collect();
        index += arity;
        match command {
            'M' => {
                builder.move_to((values[0], values[1]));
                // Further pairs are line-tos.
                command = 'L';
            }
            'L' => builder.line_to((values[0], values[1])),
            'C' => builder.cubic(
                (values[0], values[1]),
                (values[2], values[3]),
                (values[4], values[5]),
            ),
            'Q' => {
                let p0 = builder.point;
                let c = (values[0], values[1]);
                let p = (values[2], values[3]);
                builder.cubic(
                    (
                        p0.0 + 2.0 / 3.0 * (c.0 - p0.0),
                        p0.1 + 2.0 / 3.0 * (c.1 - p0.1),
                    ),
                    (p.0 + 2.0 / 3.0 * (c.0 - p.0), p.1 + 2.0 / 3.0 * (c.1 - p.1)),
                    p,
                );
            }
            'T' | 'U' => {
                // Angles in degrees, counter-clockwise on screen from the x
                // axis (LibreOffice EnhancedCustomShape2d).
                let (cx, cy, rx, ry) = (values[0], values[1], values[2], values[3]);
                let mut start = values[4];
                let mut end = values[5];
                if end <= start {
                    end += 360.0;
                }
                if (end - start - 360.0).abs() < 1e-9 || end - start > 360.0 {
                    start = 0.0;
                    end = 360.0;
                }
                builder.arc(
                    (cx, cy),
                    (rx, ry),
                    -start.to_radians(),
                    -(end - start).to_radians(),
                    command == 'T',
                );
            }
            'A' | 'B' | 'V' | 'W' => {
                let (x1, y1, x2, y2) = (values[0], values[1], values[2], values[3]);
                let center = ((x1 + x2) / 2.0, (y1 + y2) / 2.0);
                let r = ((x2 - x1).abs() / 2.0, (y2 - y1).abs() / 2.0);
                if r.0 == 0.0 || r.1 == 0.0 {
                    builder.line_to((values[6], values[7]));
                    continue;
                }
                let angle = |x: f64, y: f64| ((y - center.1) / r.1).atan2((x - center.0) / r.0);
                let start = angle(values[4], values[5]);
                let end = angle(values[6], values[7]);
                let clockwise = matches!(command, 'V' | 'W');
                let tau = std::f64::consts::TAU;
                let mut sweep = end - start;
                if clockwise {
                    while sweep <= 0.0 {
                        sweep += tau;
                    }
                } else {
                    while sweep >= 0.0 {
                        sweep -= tau;
                    }
                }
                builder.arc(center, r, start, sweep, matches!(command, 'A' | 'W'));
            }
            'G' => {
                // OOXML arcTo (LibreOffice extension): radii, start and
                // swing angles in degrees, clockwise on screen.
                let (rx, ry) = (values[0], values[1]);
                let start = values[2].to_radians();
                let sweep = values[3].to_radians();
                // The parametric angle of the start point on the ellipse.
                let parametric = (start.sin() * rx).atan2(start.cos() * ry);
                let center = (
                    builder.point.0 - rx * parametric.cos(),
                    builder.point.1 - ry * parametric.sin(),
                );
                let end = start + sweep;
                let end_parametric = (end.sin() * rx).atan2(end.cos() * ry);
                let mut parametric_sweep = end_parametric - parametric;
                if sweep > 0.0 && parametric_sweep < 0.0 {
                    parametric_sweep += std::f64::consts::TAU;
                } else if sweep < 0.0 && parametric_sweep > 0.0 {
                    parametric_sweep -= std::f64::consts::TAU;
                }
                builder.arc(center, (rx, ry), parametric, parametric_sweep, true);
            }
            'X' | 'Y' => {
                let p0 = builder.point;
                let p = (values[0], values[1]);
                let k = 0.552_284_75;
                if quadrant_x {
                    builder.cubic(
                        (p0.0 + k * (p.0 - p0.0), p0.1),
                        (p.0, p.1 - k * (p.1 - p0.1)),
                        p,
                    );
                } else {
                    builder.cubic(
                        (p0.0, p0.1 + k * (p.1 - p0.1)),
                        (p.0 - k * (p.0 - p0.0), p.1),
                        p,
                    );
                }
                quadrant_x = !quadrant_x;
            }
            _ => {}
        }
    }
    builder.end_set();
    Ok(Path {
        subpaths: builder.subpaths,
        paint: builder.paint,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn expression_nesting_is_bounded() {
        // Parentheses, unary minus over parentheses and function arguments
        // nest the parser; past the bound the formula reads as 0.
        for formula in [
            format!("{}1{}", "(".repeat(200_000), ")".repeat(200_000)),
            format!("{}1{}", "abs(".repeat(200_000), ")".repeat(200_000)),
            format!("{}1{}", "-(".repeat(200_000), ")".repeat(200_000)),
        ] {
            let geometry = geometry(&[("f0", formula.as_str())], &[]);
            let mut eval = Evaluator::new(&geometry);
            assert_eq!(eval.parameter("?f0"), 0.0);
        }
        // A cycle reads as 0 without re-entering the equation: a
        // self-doubling reference stays linear.
        let cyclic = geometry(&[("f", "?f+?f+1")], &[]);
        assert_eq!(Evaluator::new(&cyclic).parameter("?f"), 1.0);
        let geometry = geometry(&[("f0", "((1+2))*abs(-3)")], &[]);
        assert_eq!(Evaluator::new(&geometry).parameter("?f0"), 9.0);
    }

    fn geometry(equations: &[(&str, &str)], modifiers: &[f64]) -> Geometry {
        Geometry {
            view_box: (0.0, 0.0, 21600.0, 21600.0),
            equations: equations
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect(),
            modifiers: modifiers.to_vec(),
            log_size: (1000.0, 1000.0),
            has_fill: true,
            has_stroke: true,
        }
    }

    #[test]
    fn evaluates_formulas_with_references_functions_and_identifiers() {
        let g = geometry(
            &[
                ("f0", "$0 * 2"),
                ("f1", "?f0 + width / 2"),
                ("f2", "if(?f0 - 100, max(3, 7), -1)"),
                ("f3", "sqrt(16) + abs(-2)"),
            ],
            &[100.0],
        );
        let mut eval = Evaluator::new(&g);
        assert_eq!(eval.equation("f1"), 200.0 + 10800.0);
        assert_eq!(eval.equation("f2"), 7.0);
        assert_eq!(eval.equation("f3"), 6.0);
    }

    #[test]
    fn evaluates_paths_into_normalized_commands() {
        let g = geometry(&[("f0", "21600")], &[]);
        let path = evaluate("M 0 0 L ?f0 0 ?f0 21600 Z N", &g).unwrap();
        assert_eq!(path.subpaths.len(), 1);
        match path.subpaths[0][2] {
            PathCmd::LineTo { x, y } => assert_eq!((x, y), (1.0, 1.0)),
            _ => panic!(),
        }
        let ellipse = evaluate("U 10800 10800 10800 10800 0 360 Z N", &g).unwrap();
        assert_eq!(ellipse.subpaths.len(), 1);
        assert!(ellipse.subpaths[0].len() >= 5);
    }
}
