//! `draw:enhanced-geometry` paths (ODF Part 3 §19.145 `draw:enhanced-path`,
//! §19.141 `draw:equation`) as the workbook model's custom geometry.
//!
//! Path parameters are numbers, equation results (`?fN`), modifier values
//! (`$N`) or the §19.145 keywords; equations are evaluated with the §19.141
//! grammar. LibreOffice writes imported OOXML custom geometry as
//! `draw:type="ooxml-non-primitive"` with one coordinate space per subpath
//! (`drawooo:sub-view-size`), otherwise the `svg:viewBox` is the space.
//! Elliptic arcs (`T`, `U`, `A`, `B`, `V`, `W`, `X`, `Y`) are flattened to
//! line segments of at most 5 degrees.

use crate::xml::{self, NS_DRAW, NS_SVG};
use xlsx_model::{PathCmd, PathInfo};

const NS_DRAWOOO: &str = "http://openoffice.org/2010/draw";
/// Implementation bounds: equations per shape, recursion depth, commands.
const MAX_EQUATIONS: usize = 4096;
const MAX_DEPTH: usize = 64;
const MAX_COMMANDS: usize = 200_000;
const ARC_STEP_DEGREES: f64 = 5.0;

struct Env<'a> {
    equations: Vec<&'a str>,
    cache: Vec<Option<f64>>,
    modifiers: Vec<f64>,
    view: (f64, f64, f64, f64),
    log: (f64, f64),
}

impl Env<'_> {
    fn equation(&mut self, index: usize, depth: usize) -> f64 {
        if let Some(Some(value)) = self.cache.get(index) {
            return *value;
        }
        if depth > MAX_DEPTH || index >= self.equations.len() {
            return 0.0;
        }
        let formula = self.equations[index];
        let value = Parser {
            text: formula.as_bytes(),
            at: 0,
        }
        .expression(self, depth + 1);
        let value = if value.is_finite() { value } else { 0.0 };
        self.cache[index] = Some(value);
        value
    }

    fn keyword(&self, name: &str) -> Option<f64> {
        let (x, y, w, h) = self.view;
        Some(match name {
            "pi" => std::f64::consts::PI,
            "left" => x,
            "top" => y,
            "right" => x + w,
            "bottom" => y + h,
            "width" => w,
            "height" => h,
            "xstretch" | "ystretch" => 0.0,
            "hasstroke" | "hasfill" => 1.0,
            "logwidth" => self.log.0,
            "logheight" => self.log.1,
            _ => return None,
        })
    }
}

/// A recursive-descent reader of the §19.141 formula grammar.
struct Parser<'t> {
    text: &'t [u8],
    at: usize,
}

impl Parser<'_> {
    fn skip(&mut self) {
        while self.at < self.text.len() && self.text[self.at].is_ascii_whitespace() {
            self.at += 1;
        }
    }

    fn peek(&mut self) -> Option<u8> {
        self.skip();
        self.text.get(self.at).copied()
    }

    fn expression(&mut self, env: &mut Env<'_>, depth: usize) -> f64 {
        let mut value = self.term(env, depth);
        while let Some(op @ (b'+' | b'-')) = self.peek() {
            self.at += 1;
            let rhs = self.term(env, depth);
            value = if op == b'+' { value + rhs } else { value - rhs };
        }
        value
    }

    fn term(&mut self, env: &mut Env<'_>, depth: usize) -> f64 {
        let mut value = self.unary(env, depth);
        while let Some(op @ (b'*' | b'/')) = self.peek() {
            self.at += 1;
            let rhs = self.unary(env, depth);
            value = if op == b'*' {
                value * rhs
            } else if rhs != 0.0 {
                value / rhs
            } else {
                0.0
            };
        }
        value
    }

    fn unary(&mut self, env: &mut Env<'_>, depth: usize) -> f64 {
        match self.peek() {
            Some(b'-') => {
                self.at += 1;
                -self.unary(env, depth)
            }
            Some(b'+') => {
                self.at += 1;
                self.unary(env, depth)
            }
            _ => self.primary(env, depth),
        }
    }

    fn primary(&mut self, env: &mut Env<'_>, depth: usize) -> f64 {
        match self.peek() {
            Some(b'(') => {
                self.at += 1;
                let value = self.expression(env, depth);
                if self.peek() == Some(b')') {
                    self.at += 1;
                }
                value
            }
            Some(b'?') => {
                self.at += 1;
                let start = self.at;
                while self.at < self.text.len() && self.text[self.at].is_ascii_alphanumeric() {
                    self.at += 1;
                }
                let name = std::str::from_utf8(&self.text[start..self.at]).unwrap_or("");
                name.strip_prefix('f')
                    .and_then(|n| n.parse::<usize>().ok())
                    .map_or(0.0, |n| env.equation(n, depth))
            }
            Some(b'$') => {
                self.at += 1;
                let index = self.number().unwrap_or(0.0) as usize;
                env.modifiers.get(index).copied().unwrap_or(0.0)
            }
            Some(c) if c.is_ascii_digit() || c == b'.' => self.number().unwrap_or(0.0),
            Some(c) if c.is_ascii_alphabetic() => {
                let start = self.at;
                while self.at < self.text.len() && self.text[self.at].is_ascii_alphanumeric() {
                    self.at += 1;
                }
                let name = std::str::from_utf8(&self.text[start..self.at])
                    .unwrap_or("")
                    .to_string();
                if self.peek() == Some(b'(') {
                    self.at += 1;
                    let mut args = Vec::new();
                    loop {
                        args.push(self.expression(env, depth));
                        match self.peek() {
                            Some(b',') => self.at += 1,
                            Some(b')') => {
                                self.at += 1;
                                break;
                            }
                            _ => break,
                        }
                    }
                    let a = |i: usize| args.get(i).copied().unwrap_or(0.0);
                    match name.as_str() {
                        "abs" => a(0).abs(),
                        "sqrt" => a(0).max(0.0).sqrt(),
                        "sin" => a(0).sin(),
                        "cos" => a(0).cos(),
                        "tan" => a(0).tan(),
                        "atan" => a(0).atan(),
                        "atan2" => a(0).atan2(a(1)),
                        "min" => a(0).min(a(1)),
                        "max" => a(0).max(a(1)),
                        "if" => {
                            if a(0) > 0.0 {
                                a(1)
                            } else {
                                a(2)
                            }
                        }
                        _ => 0.0,
                    }
                } else {
                    env.keyword(&name).unwrap_or(0.0)
                }
            }
            _ => {
                self.at = self.text.len();
                0.0
            }
        }
    }

    fn number(&mut self) -> Option<f64> {
        self.skip();
        let start = self.at;
        while self.at < self.text.len()
            && (self.text[self.at].is_ascii_digit()
                || self.text[self.at] == b'.'
                || (self.at > start && matches!(self.text[self.at], b'e' | b'E')))
        {
            self.at += 1;
        }
        std::str::from_utf8(&self.text[start..self.at])
            .ok()?
            .parse()
            .ok()
    }
}

/// One path token: a command letter or a parameter value.
enum Token {
    Command(char),
    Value(f64),
}

fn tokens(path: &str, env: &mut Env<'_>) -> Vec<Token> {
    let mut out = Vec::new();
    for word in path
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|w| !w.is_empty())
    {
        if out.len() > MAX_COMMANDS * 8 {
            break;
        }
        let first = word.chars().next().unwrap_or(' ');
        if word.len() == 1 && first.is_ascii_alphabetic() {
            out.push(Token::Command(first));
            continue;
        }
        let value = if let Some(rest) = word.strip_prefix('?') {
            rest.strip_prefix('f')
                .and_then(|n| n.parse::<usize>().ok())
                .map_or(0.0, |n| env.equation(n, 0))
        } else if let Some(rest) = word.strip_prefix('$') {
            rest.parse::<usize>()
                .ok()
                .and_then(|i| env.modifiers.get(i).copied())
                .unwrap_or(0.0)
        } else if let Ok(number) = word.parse::<f64>() {
            number
        } else if let Some(value) = env.keyword(word) {
            value
        } else {
            // A command letter fused to its first number (`M0 0`).
            if first.is_ascii_alphabetic() {
                out.push(Token::Command(first));
                if let Ok(number) = word[1..].parse::<f64>() {
                    out.push(Token::Value(number));
                }
            }
            continue;
        };
        out.push(Token::Value(value));
    }
    out
}

/// Points of an elliptic arc around (cx, cy), angles in degrees, y down.
fn arc_points(cx: f64, cy: f64, rx: f64, ry: f64, start: f64, sweep: f64) -> Vec<(f64, f64)> {
    let steps = ((sweep.abs() / ARC_STEP_DEGREES).ceil() as usize).clamp(1, 720);
    (0..=steps)
        .map(|i| {
            let angle = (start + sweep * i as f64 / steps as f64).to_radians();
            (cx + rx * angle.cos(), cy - ry * angle.sin())
        })
        .collect()
}

/// The angle (degrees, counter-clockwise, y up) of a point on an ellipse.
fn angle_of(cx: f64, cy: f64, rx: f64, ry: f64, x: f64, y: f64) -> f64 {
    let dx = if rx != 0.0 { (x - cx) / rx } else { 0.0 };
    let dy = if ry != 0.0 { (cy - y) / ry } else { 0.0 };
    dy.atan2(dx).to_degrees()
}

/// Read a custom shape's geometry. `None` when it has no enhanced path.
pub fn custom_geometry(
    shape: roxmltree::Node<'_, '_>,
    size_pt: (f64, f64),
) -> Option<Vec<PathInfo>> {
    let geometry = shape
        .children()
        .find(|n| xml::is(*n, NS_DRAW, "enhanced-geometry"))?;
    let path = geometry.attribute((NS_DRAW, "enhanced-path"))?;
    let view: Vec<f64> = geometry
        .attribute((NS_SVG, "viewBox"))
        .unwrap_or("0 0 21600 21600")
        .split_whitespace()
        .filter_map(|v| v.parse().ok())
        .collect();
    let mut view = (
        view.first().copied().unwrap_or(0.0),
        view.get(1).copied().unwrap_or(0.0),
        view.get(2).copied().unwrap_or(21600.0),
        view.get(3).copied().unwrap_or(21600.0),
    );
    let sub_views: Vec<(f64, f64)> = geometry
        .attribute((NS_DRAWOOO, "sub-view-size"))
        .map(|v| {
            let numbers: Vec<f64> = v
                .split_whitespace()
                .filter_map(|n| n.parse().ok())
                .collect();
            numbers
                .chunks(2)
                .filter(|c| c.len() == 2)
                .map(|c| (c[0], c[1]))
                .collect()
        })
        .unwrap_or_default();
    if view.2 <= 0.0 || view.3 <= 0.0 {
        if let Some((w, h)) = sub_views.first() {
            view = (0.0, 0.0, *w, *h);
        }
    }
    let equations: Vec<(&str, &str)> = geometry
        .children()
        .filter(|n| xml::is(*n, NS_DRAW, "equation"))
        .take(MAX_EQUATIONS)
        .filter_map(|n| {
            Some((
                n.attribute((NS_DRAW, "name"))?,
                n.attribute((NS_DRAW, "formula"))?,
            ))
        })
        .collect();
    // `?fN` names the N-th equation by its `draw:name` "fN" (§19.141).
    let mut ordered = vec![""; equations.len()];
    for (name, formula) in &equations {
        if let Some(index) = name.strip_prefix('f').and_then(|n| n.parse::<usize>().ok()) {
            if index < ordered.len() {
                ordered[index] = formula;
            }
        }
    }
    let modifiers = geometry
        .attribute((NS_DRAW, "modifiers"))
        .map(|m| {
            m.split_whitespace()
                .filter_map(|v| v.parse().ok())
                .collect()
        })
        .unwrap_or_default();
    let mut env = Env {
        cache: vec![None; ordered.len()],
        equations: ordered,
        modifiers,
        view,
        log: (size_pt.0 * 2540.0 / 72.0, size_pt.1 * 2540.0 / 72.0),
    };
    let tokens = tokens(path, &mut env);
    let mut paths: Vec<PathInfo> = Vec::new();
    let new_path = |index: usize| PathInfo {
        w: sub_views.get(index).map_or(view.2, |v| v.0).max(1.0),
        h: sub_views.get(index).map_or(view.3, |v| v.1).max(1.0),
        fill: None,
        stroke: true,
        commands: Vec::new(),
    };
    let mut current = new_path(0);
    let (ox, oy) = (view.0, view.1);
    let mut command = 'M';
    let mut values: Vec<f64> = Vec::new();
    let mut last = (0.0f64, 0.0f64);
    let mut emitted = 0usize;
    let flush = |command: char,
                 values: &mut Vec<f64>,
                 current: &mut PathInfo,
                 last: &mut (f64, f64),
                 emitted: &mut usize| {
        let p = |x: f64, y: f64| (x - ox, y - oy);
        let mut push = |cmd: PathCmd, current: &mut PathInfo| {
            if *emitted < MAX_COMMANDS {
                current.commands.push(cmd);
                *emitted += 1;
            }
        };
        match command {
            'M' => {
                for (i, pair) in values.chunks(2).filter(|c| c.len() == 2).enumerate() {
                    let (x, y) = p(pair[0], pair[1]);
                    push(
                        if i == 0 {
                            PathCmd::MoveTo { x, y }
                        } else {
                            PathCmd::LineTo { x, y }
                        },
                        current,
                    );
                    *last = (x, y);
                }
            }
            'L' => {
                for pair in values.chunks(2).filter(|c| c.len() == 2) {
                    let (x, y) = p(pair[0], pair[1]);
                    push(PathCmd::LineTo { x, y }, current);
                    *last = (x, y);
                }
            }
            'C' => {
                for c in values.chunks(6).filter(|c| c.len() == 6) {
                    let (x1, y1) = p(c[0], c[1]);
                    let (x2, y2) = p(c[2], c[3]);
                    let (x3, y3) = p(c[4], c[5]);
                    push(
                        PathCmd::CubicBezTo {
                            x1,
                            y1,
                            x2,
                            y2,
                            x3,
                            y3,
                        },
                        current,
                    );
                    *last = (x3, y3);
                }
            }
            'Q' => {
                for c in values.chunks(4).filter(|c| c.len() == 4) {
                    let (x1, y1) = p(c[0], c[1]);
                    let (x2, y2) = p(c[2], c[3]);
                    push(PathCmd::QuadBezTo { x1, y1, x2, y2 }, current);
                    *last = (x2, y2);
                }
            }
            'T' | 'U' => {
                // angle-ellipse(to): centre, radii, start and end angle.
                for (i, c) in values.chunks(6).filter(|c| c.len() == 6).enumerate() {
                    let (cx, cy) = p(c[0], c[1]);
                    let (start, end) = (c[4], c[5]);
                    let mut sweep = end - start;
                    if sweep <= 0.0 {
                        sweep += 360.0;
                    }
                    for (j, (x, y)) in arc_points(cx, cy, c[2], c[3], start, sweep)
                        .into_iter()
                        .enumerate()
                    {
                        let first_move = command == 'U' && i == 0 && j == 0;
                        push(
                            if first_move {
                                PathCmd::MoveTo { x, y }
                            } else {
                                PathCmd::LineTo { x, y }
                            },
                            current,
                        );
                        *last = (x, y);
                    }
                }
            }
            'A' | 'B' | 'V' | 'W' => {
                // arc(to): bounding box, start and end points (counter-
                // clockwise for A/B, clockwise for V/W).
                for (i, c) in values.chunks(8).filter(|c| c.len() == 8).enumerate() {
                    let (x1, y1) = p(c[0], c[1]);
                    let (x2, y2) = p(c[2], c[3]);
                    let (cx, cy, rx, ry) = (
                        (x1 + x2) / 2.0,
                        (y1 + y2) / 2.0,
                        (x2 - x1).abs() / 2.0,
                        (y2 - y1).abs() / 2.0,
                    );
                    let (sx, sy) = p(c[4], c[5]);
                    let (ex, ey) = p(c[6], c[7]);
                    let start = angle_of(cx, cy, rx, ry, sx, sy);
                    let end = angle_of(cx, cy, rx, ry, ex, ey);
                    let clockwise = matches!(command, 'V' | 'W');
                    let mut sweep = end - start;
                    if clockwise {
                        if sweep >= 0.0 {
                            sweep -= 360.0;
                        }
                    } else if sweep <= 0.0 {
                        sweep += 360.0;
                    }
                    let move_first = matches!(command, 'B' | 'V');
                    for (j, (x, y)) in arc_points(cx, cy, rx, ry, start, sweep)
                        .into_iter()
                        .enumerate()
                    {
                        let first = move_first && i == 0 && j == 0;
                        push(
                            if first {
                                PathCmd::MoveTo { x, y }
                            } else {
                                PathCmd::LineTo { x, y }
                            },
                            current,
                        );
                        *last = (x, y);
                    }
                }
            }
            'X' | 'Y' => {
                // Elliptical quadrants, alternating x-first and y-first.
                let mut x_first = command == 'X';
                for pair in values.chunks(2).filter(|c| c.len() == 2) {
                    let (x, y) = p(pair[0], pair[1]);
                    let (lx, ly) = *last;
                    let (cx, cy) = if x_first { (lx, y) } else { (x, ly) };
                    let (rx, ry) = ((x - lx).abs(), (y - ly).abs());
                    let start = angle_of(cx, cy, rx, ry, lx, ly);
                    let end = angle_of(cx, cy, rx, ry, x, y);
                    let mut sweep = end - start;
                    if sweep > 180.0 {
                        sweep -= 360.0;
                    } else if sweep < -180.0 {
                        sweep += 360.0;
                    }
                    for (px, py) in arc_points(cx, cy, rx, ry, start, sweep).into_iter().skip(1) {
                        push(PathCmd::LineTo { x: px, y: py }, current);
                    }
                    *last = (x, y);
                    x_first = !x_first;
                }
            }
            _ => {}
        }
        values.clear();
    };
    let mut path_index = 0usize;
    for token in tokens {
        match token {
            Token::Value(v) => values.push(v),
            Token::Command(c) => {
                flush(command, &mut values, &mut current, &mut last, &mut emitted);
                match c {
                    'Z' => current.commands.push(PathCmd::Close),
                    'N' => {
                        let next = new_path(path_index + 1);
                        paths.push(std::mem::replace(&mut current, next));
                        path_index += 1;
                    }
                    'F' => current.fill = Some("none".into()),
                    'S' => current.stroke = false,
                    _ => {}
                }
                command = c;
            }
        }
    }
    flush(command, &mut values, &mut current, &mut last, &mut emitted);
    if !current.commands.is_empty() {
        paths.push(current);
    }
    paths.retain(|p| !p.commands.is_empty());
    (!paths.is_empty()).then_some(paths)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn evaluates_equations_and_subpaths() {
        let xml = format!(
            r#"<draw:custom-shape xmlns:draw="{NS_DRAW}" xmlns:svg="{NS_SVG}" xmlns:drawooo="{NS_DRAWOOO}"><draw:enhanced-geometry svg:viewBox="0 0 0 0" drawooo:sub-view-size="100 50 10 10" draw:enhanced-path="M 0 0 L ?f0 50 F N M 0 0 L 10 10 Z N"><draw:equation draw:name="f0" draw:formula="width*2 - logwidth*0 + max(1, 3)"/></draw:enhanced-geometry></draw:custom-shape>"#
        );
        let document = roxmltree::Document::parse(&xml).unwrap();
        let paths = custom_geometry(document.root_element(), (72.0, 36.0)).unwrap();
        assert_eq!(paths.len(), 2);
        assert_eq!((paths[0].w, paths[0].h), (100.0, 50.0));
        assert_eq!(paths[0].fill.as_deref(), Some("none"));
        match paths[0].commands[1] {
            PathCmd::LineTo { x, y } => assert_eq!((x, y), (203.0, 50.0)),
            _ => panic!(),
        }
        assert_eq!((paths[1].w, paths[1].h), (10.0, 10.0));
        assert!(matches!(paths[1].commands.last(), Some(PathCmd::Close)));
    }
}
