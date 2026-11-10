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
/// Implementation bounds: equations per shape, nesting depth of one
/// evaluation (parentheses, signs, function arguments and equation
/// references together), path commands per shape. A shape over any of them
/// rejects the document.
const MAX_EQUATIONS: usize = 4096;
const MAX_DEPTH: usize = 256;
const MAX_COMMANDS: usize = 200_000;
const ARC_STEP_DEGREES: f64 = 5.0;

/// The rejection of a shape over the geometry bounds.
fn over_limit() -> String {
    crate::error::resource("shape geometry")
}

#[derive(Clone, Copy, PartialEq)]
enum Slot {
    Pending,
    Evaluating,
    Done(f64),
}

struct Env<'a> {
    equations: Vec<&'a str>,
    /// `draw:name` → equation index (§19.141: `?name` references).
    names: std::collections::HashMap<&'a str, usize>,
    cache: Vec<Slot>,
    modifiers: Vec<f64>,
    view: (f64, f64, f64, f64),
    log: (f64, f64),
    /// Set when an evaluation exceeded [`MAX_DEPTH`].
    overflow: bool,
}

impl Env<'_> {
    fn equation(&mut self, index: usize, depth: usize) -> f64 {
        match self.cache.get(index) {
            Some(Slot::Done(value)) => return *value,
            // A reference into its own evaluation (a cycle) has no value.
            Some(Slot::Evaluating) | None => return 0.0,
            Some(Slot::Pending) => {}
        }
        if depth > MAX_DEPTH {
            self.overflow = true;
            return 0.0;
        }
        self.cache[index] = Slot::Evaluating;
        let formula = self.equations[index];
        let value = Parser {
            text: formula.as_bytes(),
            at: 0,
        }
        .expression(self, depth + 1);
        let value = if value.is_finite() { value } else { 0.0 };
        self.cache[index] = Slot::Done(value);
        value
    }

    /// The value of the equation named `name` (0 when none has the name).
    fn named(&mut self, name: &str, depth: usize) -> f64 {
        match self.names.get(name) {
            Some(index) => self.equation(*index, depth),
            None => 0.0,
        }
    }

    /// Whether an evaluation at `depth` is within the bound (and records
    /// the overflow otherwise).
    fn enter(&mut self, depth: usize) -> bool {
        if depth > MAX_DEPTH {
            self.overflow = true;
        }
        !self.overflow
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
        if !env.enter(depth) {
            self.at = self.text.len();
            return 0.0;
        }
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
        if !env.enter(depth) {
            self.at = self.text.len();
            return 0.0;
        }
        match self.peek() {
            Some(b'-') => {
                self.at += 1;
                -self.unary(env, depth + 1)
            }
            Some(b'+') => {
                self.at += 1;
                self.unary(env, depth + 1)
            }
            _ => self.primary(env, depth),
        }
    }

    fn primary(&mut self, env: &mut Env<'_>, depth: usize) -> f64 {
        match self.peek() {
            Some(b'(') => {
                self.at += 1;
                let value = self.expression(env, depth + 1);
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
                env.named(name, depth)
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
                        args.push(self.expression(env, depth + 1));
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

fn tokens(path: &str, env: &mut Env<'_>) -> Result<Vec<Token>, String> {
    let mut out = Vec::new();
    for word in path
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|w| !w.is_empty())
    {
        if out.len() > MAX_COMMANDS * 8 {
            return Err(over_limit());
        }
        let first = word.chars().next().unwrap_or(' ');
        if word.len() == 1 && first.is_ascii_alphabetic() {
            out.push(Token::Command(first));
            continue;
        }
        let value = if let Some(rest) = word.strip_prefix('?') {
            env.named(rest, 0)
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
    if env.overflow {
        return Err(over_limit());
    }
    Ok(out)
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

/// Read a custom shape's geometry. `Ok(None)` when it has no enhanced path;
/// a shape over the geometry bounds rejects the document.
pub fn custom_geometry(
    shape: roxmltree::Node<'_, '_>,
    size_pt: (f64, f64),
) -> Result<Option<Vec<PathInfo>>, String> {
    let Some(geometry) = shape
        .children()
        .find(|n| xml::is(*n, NS_DRAW, "enhanced-geometry"))
    else {
        return Ok(None);
    };
    let Some(path) = geometry.attribute((NS_DRAW, "enhanced-path")) else {
        return Ok(None);
    };
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
    let equation_nodes: Vec<_> = geometry
        .children()
        .filter(|n| xml::is(*n, NS_DRAW, "equation"))
        .collect();
    if equation_nodes.len() > MAX_EQUATIONS {
        return Err(over_limit());
    }
    // `?name` references the equation whose `draw:name` is `name` (§19.141).
    let mut equations: Vec<&str> = Vec::new();
    let mut names = std::collections::HashMap::new();
    for node in equation_nodes {
        let (Some(name), Some(formula)) = (
            node.attribute((NS_DRAW, "name")),
            node.attribute((NS_DRAW, "formula")),
        ) else {
            continue;
        };
        names.entry(name).or_insert(equations.len());
        equations.push(formula);
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
        cache: vec![Slot::Pending; equations.len()],
        equations,
        names,
        modifiers,
        view,
        log: (size_pt.0 * 2540.0 / 72.0, size_pt.1 * 2540.0 / 72.0),
        overflow: false,
    };
    // Equations mostly reference earlier ones: evaluating them in order
    // keeps each evaluation's reference chain short.
    for index in 0..env.equations.len() {
        env.equation(index, 0);
        if env.overflow {
            return Err(over_limit());
        }
    }
    let tokens = tokens(path, &mut env)?;
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
        // Commands past the bound are counted, not stored: the shape is
        // rejected once the path is read.
        let mut push = |cmd: PathCmd, current: &mut PathInfo| {
            *emitted += 1;
            if *emitted <= MAX_COMMANDS {
                current.commands.push(cmd);
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
                if matches!(c, 'Z' | 'N') {
                    emitted += 1;
                }
                if emitted > MAX_COMMANDS {
                    return Err(over_limit());
                }
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
    if emitted > MAX_COMMANDS {
        return Err(over_limit());
    }
    if !current.commands.is_empty() {
        paths.push(current);
    }
    paths.retain(|p| !p.commands.is_empty());
    Ok((!paths.is_empty()).then_some(paths))
}

/// `svg:d` path data (SVG 1.1 §8.3, the geometry of ODF `draw:path`,
/// Part 3 §10.3.7) in the coordinate space of `view` (`svg:viewBox`):
/// moves, lines, horizontal/vertical lines, cubic and quadratic Béziers
/// with their smooth forms, elliptical arcs (flattened like the enhanced
/// path's arcs) and close-paths, absolute and relative. A path over the
/// geometry bounds rejects the document; malformed data ends the path at
/// the error, as SVG renders the path up to its first error (§F.6).
pub fn svg_path(d: &str, view: (f64, f64, f64, f64)) -> Result<Option<PathInfo>, String> {
    let (ox, oy) = (view.0, view.1);
    let mut commands: Vec<PathCmd> = Vec::new();
    let push = |commands: &mut Vec<PathCmd>, cmd: PathCmd| -> Result<(), String> {
        if commands.len() >= MAX_COMMANDS {
            return Err(over_limit());
        }
        commands.push(cmd);
        Ok(())
    };
    let mut reader = SvgReader {
        text: d.as_bytes(),
        at: 0,
    };
    let (mut cx, mut cy) = (0.0f64, 0.0f64);
    let (mut sx, mut sy) = (0.0f64, 0.0f64);
    // The previous command's second control point, for S / T.
    let mut last_cubic: Option<(f64, f64)> = None;
    let mut last_quad: Option<(f64, f64)> = None;
    let mut command: Option<u8> = None;
    loop {
        reader.skip_separators();
        let Some(next) = reader.peek() else {
            break;
        };
        if next.is_ascii_alphabetic() {
            reader.at += 1;
            command = Some(next);
        } else if command.is_none() {
            break;
        }
        let Some(letter) = command else {
            break;
        };
        let relative = letter.is_ascii_lowercase();
        let (bx, by) = if relative { (cx, cy) } else { (0.0, 0.0) };
        let p = |x: f64, y: f64| (x - ox, y - oy);
        let upper = letter.to_ascii_uppercase();
        let ok = match upper {
            b'Z' => {
                push(&mut commands, PathCmd::Close)?;
                cx = sx;
                cy = sy;
                command = None;
                last_cubic = None;
                last_quad = None;
                continue;
            }
            b'M' => match reader.numbers::<2>() {
                Some([x, y]) => {
                    cx = bx + x;
                    cy = by + y;
                    sx = cx;
                    sy = cy;
                    let (x, y) = p(cx, cy);
                    push(&mut commands, PathCmd::MoveTo { x, y })?;
                    // Further pairs are implicit line-tos.
                    command = Some(if relative { b'l' } else { b'L' });
                    true
                }
                None => false,
            },
            b'L' => match reader.numbers::<2>() {
                Some([x, y]) => {
                    cx = bx + x;
                    cy = by + y;
                    let (x, y) = p(cx, cy);
                    push(&mut commands, PathCmd::LineTo { x, y })?;
                    true
                }
                None => false,
            },
            b'H' => match reader.numbers::<1>() {
                Some([x]) => {
                    cx = bx + x;
                    let (x, y) = p(cx, cy);
                    push(&mut commands, PathCmd::LineTo { x, y })?;
                    true
                }
                None => false,
            },
            b'V' => match reader.numbers::<1>() {
                Some([y]) => {
                    cy = if relative { cy + y } else { y };
                    let (x, y) = p(cx, cy);
                    push(&mut commands, PathCmd::LineTo { x, y })?;
                    true
                }
                None => false,
            },
            b'C' | b'S' => {
                let points = if upper == b'C' {
                    reader.numbers::<6>().map(|[x1, y1, x2, y2, x, y]| {
                        (bx + x1, by + y1, bx + x2, by + y2, bx + x, by + y)
                    })
                } else {
                    reader.numbers::<4>().map(|[x2, y2, x, y]| {
                        let (x1, y1) =
                            last_cubic.map_or((cx, cy), |(lx, ly)| (2.0 * cx - lx, 2.0 * cy - ly));
                        (x1, y1, bx + x2, by + y2, bx + x, by + y)
                    })
                };
                match points {
                    Some((x1, y1, x2, y2, x, y)) => {
                        let (x1, y1) = p(x1, y1);
                        let (px2, py2) = p(x2, y2);
                        let (x3, y3) = p(x, y);
                        push(
                            &mut commands,
                            PathCmd::CubicBezTo {
                                x1,
                                y1,
                                x2: px2,
                                y2: py2,
                                x3,
                                y3,
                            },
                        )?;
                        cx = x;
                        cy = y;
                        last_cubic = Some((x2, y2));
                        last_quad = None;
                        continue;
                    }
                    None => false,
                }
            }
            b'Q' | b'T' => {
                let points = if upper == b'Q' {
                    reader
                        .numbers::<4>()
                        .map(|[x1, y1, x, y]| (bx + x1, by + y1, bx + x, by + y))
                } else {
                    reader.numbers::<2>().map(|[x, y]| {
                        let (x1, y1) =
                            last_quad.map_or((cx, cy), |(lx, ly)| (2.0 * cx - lx, 2.0 * cy - ly));
                        (x1, y1, bx + x, by + y)
                    })
                };
                match points {
                    Some((x1, y1, x, y)) => {
                        let (px1, py1) = p(x1, y1);
                        let (x2, y2) = p(x, y);
                        push(
                            &mut commands,
                            PathCmd::QuadBezTo {
                                x1: px1,
                                y1: py1,
                                x2,
                                y2,
                            },
                        )?;
                        cx = x;
                        cy = y;
                        last_quad = Some((x1, y1));
                        last_cubic = None;
                        continue;
                    }
                    None => false,
                }
            }
            b'A' => match reader.arc() {
                Some((rx, ry, rotation, large, sweep, x, y)) => {
                    let (x, y) = (bx + x, by + y);
                    for (px, py) in svg_arc(cx, cy, rx, ry, rotation, large, sweep, x, y) {
                        let (px, py) = p(px, py);
                        push(&mut commands, PathCmd::LineTo { x: px, y: py })?;
                    }
                    cx = x;
                    cy = y;
                    true
                }
                None => false,
            },
            _ => false,
        };
        last_cubic = None;
        last_quad = None;
        if !ok {
            break;
        }
    }
    if commands.is_empty() {
        return Ok(None);
    }
    Ok(Some(PathInfo {
        w: view.2.max(1.0),
        h: view.3.max(1.0),
        fill: None,
        stroke: true,
        commands,
    }))
}

struct SvgReader<'t> {
    text: &'t [u8],
    at: usize,
}

impl SvgReader<'_> {
    fn skip_separators(&mut self) {
        while self
            .text
            .get(self.at)
            .is_some_and(|c| c.is_ascii_whitespace() || *c == b',')
        {
            self.at += 1;
        }
    }

    fn peek(&self) -> Option<u8> {
        self.text.get(self.at).copied()
    }

    /// One SVG number (§8.3.9 grammar: sign, digits, fraction, exponent).
    fn number(&mut self) -> Option<f64> {
        self.skip_separators();
        let start = self.at;
        let bytes = self.text;
        let mut at = self.at;
        if matches!(bytes.get(at), Some(b'+' | b'-')) {
            at += 1;
        }
        let digits_start = at;
        while bytes.get(at).is_some_and(u8::is_ascii_digit) {
            at += 1;
        }
        if bytes.get(at) == Some(&b'.') {
            at += 1;
            while bytes.get(at).is_some_and(u8::is_ascii_digit) {
                at += 1;
            }
        }
        if at == digits_start || (at == digits_start + 1 && bytes[digits_start] == b'.') {
            return None;
        }
        if matches!(bytes.get(at), Some(b'e' | b'E')) {
            let mut exponent = at + 1;
            if matches!(bytes.get(exponent), Some(b'+' | b'-')) {
                exponent += 1;
            }
            if bytes.get(exponent).is_some_and(u8::is_ascii_digit) {
                while bytes.get(exponent).is_some_and(u8::is_ascii_digit) {
                    exponent += 1;
                }
                at = exponent;
            }
        }
        let value = std::str::from_utf8(&bytes[start..at])
            .ok()?
            .parse::<f64>()
            .ok()?;
        if !value.is_finite() {
            return None;
        }
        self.at = at;
        Some(value)
    }

    fn numbers<const N: usize>(&mut self) -> Option<[f64; N]> {
        let mut out = [0.0; N];
        for slot in &mut out {
            *slot = self.number()?;
        }
        Some(out)
    }

    /// A flag (`0` / `1`) of an arc, which may abut the next number.
    fn flag(&mut self) -> Option<bool> {
        self.skip_separators();
        let value = match self.peek()? {
            b'0' => false,
            b'1' => true,
            _ => return None,
        };
        self.at += 1;
        Some(value)
    }

    #[allow(clippy::type_complexity)]
    fn arc(&mut self) -> Option<(f64, f64, f64, bool, bool, f64, f64)> {
        let [rx, ry, rotation] = self.numbers::<3>()?;
        let large = self.flag()?;
        let sweep = self.flag()?;
        let [x, y] = self.numbers::<2>()?;
        Some((rx, ry, rotation, large, sweep, x, y))
    }
}

/// The points after the start of an SVG elliptical arc (SVG 1.1 §F.6.5
/// endpoint-to-centre conversion, §F.6.6 radius correction), flattened in
/// steps of at most [`ARC_STEP_DEGREES`].
#[allow(clippy::too_many_arguments)]
fn svg_arc(
    x1: f64,
    y1: f64,
    rx: f64,
    ry: f64,
    rotation: f64,
    large: bool,
    sweep: bool,
    x2: f64,
    y2: f64,
) -> Vec<(f64, f64)> {
    let (mut rx, mut ry) = (rx.abs(), ry.abs());
    if (x1 == x2 && y1 == y2) || rx == 0.0 || ry == 0.0 {
        // §F.6.2: a zero radius is a straight line; equal ends draw nothing.
        return if x1 == x2 && y1 == y2 {
            Vec::new()
        } else {
            vec![(x2, y2)]
        };
    }
    let phi = rotation.to_radians();
    let (sin_phi, cos_phi) = phi.sin_cos();
    let dx = (x1 - x2) / 2.0;
    let dy = (y1 - y2) / 2.0;
    let x1p = cos_phi * dx + sin_phi * dy;
    let y1p = -sin_phi * dx + cos_phi * dy;
    let lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
    if lambda > 1.0 {
        let scale = lambda.sqrt();
        rx *= scale;
        ry *= scale;
    }
    let numerator = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
    let denominator = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
    let mut coefficient = if denominator == 0.0 {
        0.0
    } else {
        (numerator / denominator).max(0.0).sqrt()
    };
    if large == sweep {
        coefficient = -coefficient;
    }
    let cxp = coefficient * rx * y1p / ry;
    let cyp = -coefficient * ry * x1p / rx;
    let cx = cos_phi * cxp - sin_phi * cyp + (x1 + x2) / 2.0;
    let cy = sin_phi * cxp + cos_phi * cyp + (y1 + y2) / 2.0;
    let angle = |ux: f64, uy: f64, vx: f64, vy: f64| (ux * vy - uy * vx).atan2(ux * vx + uy * vy);
    let theta1 = angle(1.0, 0.0, (x1p - cxp) / rx, (y1p - cyp) / ry);
    let mut delta = angle(
        (x1p - cxp) / rx,
        (y1p - cyp) / ry,
        (-x1p - cxp) / rx,
        (-y1p - cyp) / ry,
    );
    if !sweep && delta > 0.0 {
        delta -= std::f64::consts::TAU;
    } else if sweep && delta < 0.0 {
        delta += std::f64::consts::TAU;
    }
    let steps = ((delta.abs().to_degrees() / ARC_STEP_DEGREES).ceil() as usize).max(1);
    (1..=steps)
        .map(|i| {
            if i == steps {
                return (x2, y2);
            }
            let t = theta1 + delta * i as f64 / steps as f64;
            let (sin_t, cos_t) = t.sin_cos();
            (
                cx + rx * cos_phi * cos_t - ry * sin_phi * sin_t,
                cy + rx * sin_phi * cos_t + ry * cos_phi * sin_t,
            )
        })
        .collect()
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
        let paths = custom_geometry(document.root_element(), (72.0, 36.0))
            .unwrap()
            .unwrap();
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

    fn shape(equations: &str, path: &str) -> String {
        format!(
            r#"<draw:custom-shape xmlns:draw="{NS_DRAW}" xmlns:svg="{NS_SVG}"><draw:enhanced-geometry svg:viewBox="0 0 100 100" draw:enhanced-path="{path}">{equations}</draw:enhanced-geometry></draw:custom-shape>"#
        )
    }

    fn geometry(xml: &str) -> Result<Option<Vec<PathInfo>>, String> {
        let document = roxmltree::Document::parse(xml).unwrap();
        custom_geometry(document.root_element(), (72.0, 72.0))
    }

    #[test]
    fn deeply_nested_equations_reject_instead_of_overflowing() {
        // Nesting is bounded per evaluation, whatever the XML depth.
        let deep = format!("{}1{}", "(".repeat(100_000), ")".repeat(100_000));
        let xml = shape(
            &format!(r#"<draw:equation draw:name="f0" draw:formula="{deep}"/>"#),
            "M 0 0 L ?f0 0",
        );
        assert_eq!(geometry(&xml).err(), Some(over_limit()));
        let signs = shape(
            &format!(
                r#"<draw:equation draw:name="f0" draw:formula="{}1"/>"#,
                "-".repeat(100_000)
            ),
            "M 0 0 L ?f0 0",
        );
        assert_eq!(geometry(&signs).err(), Some(over_limit()));
        let nested_calls = shape(
            &format!(
                r#"<draw:equation draw:name="f0" draw:formula="{}1{}"/>"#,
                "abs(".repeat(10_000),
                ")".repeat(10_000)
            ),
            "M 0 0 L ?f0 0",
        );
        assert_eq!(geometry(&nested_calls).err(), Some(over_limit()));
        // Within the bound, nesting evaluates.
        let fine = shape(
            &format!(
                r#"<draw:equation draw:name="f0" draw:formula="{}7{}"/>"#,
                "(".repeat(50),
                ")".repeat(50)
            ),
            "M 0 0 L ?f0 0",
        );
        let paths = geometry(&fine).unwrap().unwrap();
        assert!(matches!(paths[0].commands[1], PathCmd::LineTo { x, .. } if x == 7.0));
    }

    #[test]
    fn equations_resolve_by_name_and_cycles_are_zero() {
        // A long backward chain evaluates (equations run in order), a
        // cycle has no value, and names need not be `fN`.
        let chain: String = (0..2000)
            .map(|i| {
                let formula = if i == 0 {
                    "1".to_string()
                } else {
                    format!("?f{} + 1", i - 1)
                };
                format!(r#"<draw:equation draw:name="f{i}" draw:formula="{formula}"/>"#)
            })
            .collect();
        let xml = shape(
            &format!(
                r#"{chain}<draw:equation draw:name="loop" draw:formula="?loop + 5"/><draw:equation draw:name="half" draw:formula="width / 2"/>"#
            ),
            "M ?loop 0 L ?f1999 ?half",
        );
        let paths = geometry(&xml).unwrap().unwrap();
        assert!(matches!(paths[0].commands[0], PathCmd::MoveTo { x, .. } if x == 5.0));
        assert!(
            matches!(paths[0].commands[1], PathCmd::LineTo { x, y } if x == 2000.0 && y == 50.0)
        );
    }

    #[test]
    fn too_many_commands_or_equations_reject() {
        let path = "L 1 1 ".repeat(MAX_COMMANDS + 1);
        assert_eq!(
            geometry(&shape("", &format!("M 0 0 {path}"))).err(),
            Some(over_limit())
        );
        let equations: String = (0..=MAX_EQUATIONS)
            .map(|i| format!(r#"<draw:equation draw:name="f{i}" draw:formula="1"/>"#))
            .collect();
        assert_eq!(
            geometry(&shape(&equations, "M 0 0 L 1 1")).err(),
            Some(over_limit())
        );
    }

    #[test]
    fn svg_path_data() {
        let path = svg_path(
            "M10 20 l5,5 H30 v-10 C 30 0 40 0 40 10 s10 10 20 0 Q 70 0 80 10 t 10 0 z m 1 1 2 2",
            (10.0, 0.0, 100.0, 50.0),
        )
        .unwrap()
        .unwrap();
        assert_eq!((path.w, path.h), (100.0, 50.0));
        let c = &path.commands;
        assert!(matches!(c[0], PathCmd::MoveTo { x, y } if (x, y) == (0.0, 20.0)));
        assert!(matches!(c[1], PathCmd::LineTo { x, y } if (x, y) == (5.0, 25.0)));
        assert!(matches!(c[2], PathCmd::LineTo { x, y } if (x, y) == (20.0, 25.0)));
        assert!(matches!(c[3], PathCmd::LineTo { x, y } if (x, y) == (20.0, 15.0)));
        assert!(matches!(c[4], PathCmd::CubicBezTo { x3, y3, .. } if (x3, y3) == (30.0, 10.0)));
        // `s`: the first control point reflects (40, 0) about (40, 10).
        assert!(matches!(c[5], PathCmd::CubicBezTo { x1, y1, x3, y3, .. }
            if (x1, y1) == (30.0, 20.0) && (x3, y3) == (50.0, 10.0)));
        assert!(matches!(c[6], PathCmd::QuadBezTo { x2, y2, .. } if (x2, y2) == (70.0, 10.0)));
        assert!(matches!(c[7], PathCmd::QuadBezTo { x1, y1, x2, y2 }
            if (x1, y1) == (80.0, 20.0) && (x2, y2) == (80.0, 10.0)));
        assert!(matches!(c[8], PathCmd::Close));
        // `m` after `z` is relative to the subpath start; its extra pair is
        // an implicit relative line.
        assert!(matches!(c[9], PathCmd::MoveTo { x, y } if (x, y) == (1.0, 21.0)));
        assert!(matches!(c[10], PathCmd::LineTo { x, y } if (x, y) == (3.0, 23.0)));
        // A half-circle arc ends exactly at its end point.
        let arc = svg_path("M 0 0 A 5 5 0 0 1 10 0", (0.0, 0.0, 10.0, 10.0))
            .unwrap()
            .unwrap();
        assert!(arc.commands.len() > 3);
        assert!(
            matches!(arc.commands.last(), Some(PathCmd::LineTo { x, y }) if (*x, *y) == (10.0, 0.0))
        );
        assert!(svg_path("", (0.0, 0.0, 1.0, 1.0)).unwrap().is_none());
        let long = format!("M 0 0 {}", "L 1 1 ".repeat(MAX_COMMANDS));
        assert_eq!(
            svg_path(&long, (0.0, 0.0, 1.0, 1.0)).err(),
            Some(over_limit())
        );
    }
}
