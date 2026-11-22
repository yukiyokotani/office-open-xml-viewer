//! ECMA-376 Part 4 §19.1.2.4/.14: acquire VML formulas and all path commands
//! into the same normalized geometry used by DrawingML. No painter parses VML.
use super::*;

pub(super) struct Geometry {
    pub paths: Vec<Vec<PathCmd>>,
    pub paint: Vec<PathPaint>,
}

pub(super) fn pair(raw: &str) -> Option<(f64, f64)> {
    let mut values = raw
        .split(|c: char| c == ',' || c.is_whitespace())
        .filter(|v| !v.is_empty());
    let (x, y) = (
        values.next()?.parse::<f64>().ok()?,
        values.next()?.parse::<f64>().ok()?,
    );
    (x.is_finite() && y.is_finite() && values.next().is_none()).then_some((x, y))
}

pub(super) fn attr<'a>(
    shape: roxmltree::Node<'a, '_>,
    definition: Option<roxmltree::Node<'a, '_>>,
    name: &str,
) -> Option<&'a str> {
    shape
        .attribute(name)
        .or_else(|| definition.and_then(|n| n.attribute(name)))
}

pub(super) fn path_attr<'a>(
    shape: roxmltree::Node<'a, '_>,
    definition: Option<roxmltree::Node<'a, '_>>,
    name: &str,
) -> Option<&'a str> {
    vml_direct_child(shape, "path")
        .and_then(|n| n.attribute(name))
        .or_else(|| {
            definition
                .and_then(|n| vml_direct_child(n, "path"))
                .and_then(|n| n.attribute(name))
        })
}

pub(super) fn path_source<'a>(
    shape: roxmltree::Node<'a, '_>,
    definition: Option<roxmltree::Node<'a, '_>>,
) -> Option<&'a str> {
    // An instance path (attribute or child) replaces the inherited path.
    shape
        .attribute("path")
        .or_else(|| vml_direct_child(shape, "path").and_then(|n| n.attribute("v")))
        .or_else(|| {
            definition.and_then(|n| {
                n.attribute("path")
                    .or_else(|| vml_direct_child(n, "path").and_then(|p| p.attribute("v")))
            })
        })
}

struct Values {
    guides: Vec<f64>,
    adjustments: Vec<f64>,
    named: HashMap<&'static str, f64>,
}
impl Values {
    fn get(&self, token: &str) -> Option<f64> {
        let value = if let Some(index) = token.strip_prefix('@') {
            *self.guides.get(index.parse::<usize>().ok()?)?
        } else if let Some(index) = token.strip_prefix('#') {
            *self.adjustments.get(index.parse::<usize>().ok()?)?
        } else if let Some(value) = self.named.get(token) {
            *value
        } else {
            token.parse().ok()?
        };
        value.is_finite().then_some(value)
    }
    fn formula(&self, eqn: &str) -> Option<f64> {
        let mut words = eqn.split_whitespace();
        let op = words.next()?;
        let args = words
            .take(4)
            .map(|s| self.get(s))
            .collect::<Option<Vec<_>>>()?;
        let v = *args.first()?;
        let p = *args.get(1).unwrap_or(&0.0);
        let q = *args.get(2).unwrap_or(&0.0);
        let angle = p / 65536.0 * std::f64::consts::PI / 180.0;
        let (n, arity) = match op {
            "val" => (v, 1),
            "sum" => (v + p - q, 3),
            "prod" | "product" => (((v * p / q) + 0.5).floor(), 3),
            "mid" => (((v + p) / 2.0).trunc(), 2),
            "abs" => (v.abs(), 1),
            "min" => (v.min(p), 2),
            "max" => (v.max(p), 2),
            "if" => (if v > 0.0 { p } else { q }, 3),
            "mod" => (v.hypot(p).hypot(q), 3),
            "atan2" => (p.atan2(v).to_degrees() * 65536.0, 2),
            "sin" => (v * angle.sin(), 2),
            "cos" => (v * angle.cos(), 2),
            "tan" => (v * angle.tan(), 2),
            "cosatan2" => (v * q.atan2(p).cos(), 3),
            "sinatan2" => (v * q.atan2(p).sin(), 3),
            "sqrt" => (v.sqrt(), 1),
            "sumangle" => (v + p * 65536.0 - q * 65536.0, 3),
            "ellipse" => (q * (1.0 - (v / p).powi(2)).sqrt(), 3),
            _ => return None,
        };
        // §19.1.2.4 requires signed 32-bit integer results, product ties toward
        // +infinity, mid toward zero, other inexact results toward -infinity.
        (args.len() == arity && n.is_finite() && n >= i32::MIN as f64 && n <= i32::MAX as f64)
            .then(|| n.floor())
    }
}

// Keep comma omissions (including leading/trailing omissions) distinct from
// whitespace. @/# and signs delimit adjacent parameters, e.g. @0@1 or 1-2.
fn parameters(raw: &str, values: &Values) -> Option<Vec<f64>> {
    let mut result = Vec::new();
    for part in raw.split(',') {
        let mut i = 0;
        let bytes = part.as_bytes();
        let before = result.len();
        while i < bytes.len() {
            if bytes[i].is_ascii_whitespace() {
                i += 1;
                continue;
            }
            let start = i;
            if matches!(bytes[i], b'@' | b'#' | b'+' | b'-') {
                i += 1;
            }
            while i < bytes.len() && (bytes[i].is_ascii_digit() || bytes[i] == b'.') {
                i += 1;
            }
            if start == i || (i == start + 1 && matches!(bytes[start], b'@' | b'#' | b'+' | b'-')) {
                return None;
            }
            result.push(values.get(&part[start..i])?);
        }
        if before == result.len() {
            result.push(0.0);
        }
    }
    Some(result)
}

struct Builder {
    paths: Vec<Vec<PathCmd>>,
    paint: Vec<PathPaint>,
    current: Vec<PathCmd>,
    fill: bool,
    stroke: bool,
    cp: (f64, f64),
    start: (f64, f64),
    origin: (f64, f64),
    size: (f64, f64),
}
impl Builder {
    fn point(&self, p: (f64, f64)) -> (f64, f64) {
        (
            (p.0 - self.origin.0) / self.size.0,
            (p.1 - self.origin.1) / self.size.1,
        )
    }
    fn move_to(&mut self, p: (f64, f64)) {
        let (x, y) = self.point(p);
        self.current.push(PathCmd::MoveTo { x, y });
        self.cp = p;
        self.start = p;
    }
    fn line_to(&mut self, p: (f64, f64)) {
        let (x, y) = self.point(p);
        self.current.push(PathCmd::LineTo { x, y });
        self.cp = p;
    }
    fn end(&mut self) {
        if !self.current.is_empty() {
            self.paths.push(std::mem::take(&mut self.current));
            self.paint.push(PathPaint {
                // W3C NOTE-VML, path command e: eofill within one end-delimited
                // set; separate e sets overpaint instead of cancelling ink.
                fill_rule: Some("evenodd".into()),
                fill: (!self.fill).then(|| "none".into()),
                stroke: self.stroke,
            });
        }
        self.fill = true;
        self.stroke = true;
    }
    fn arc(
        &mut self,
        center: (f64, f64),
        radii: (f64, f64),
        start: f64,
        sweep: f64,
        connect: bool,
    ) -> Option<()> {
        if radii.0 <= 0.0
            || radii.1 <= 0.0
            || ![center.0, center.1, radii.0, radii.1, start, sweep]
                .iter()
                .all(|n| n.is_finite())
        {
            return None;
        }
        let p = (
            center.0 + radii.0 * start.to_radians().cos(),
            center.1 + radii.1 * start.to_radians().sin(),
        );
        if connect && !self.current.is_empty() {
            self.line_to(p);
        } else {
            self.move_to(p);
        }
        self.current.push(PathCmd::ArcTo {
            wr: radii.0 / self.size.0,
            hr: radii.1 / self.size.1,
            st_ang: start,
            sw_ang: sweep,
        });
        self.cp = (
            center.0 + radii.0 * (start + sweep).to_radians().cos(),
            center.1 + radii.1 * (start + sweep).to_radians().sin(),
        );
        Some(())
    }
    fn command(&mut self, cmd: &str, a: &[f64]) -> Option<()> {
        let arity = match cmd {
            "m" | "t" | "l" | "r" | "qx" | "qy" => 2,
            "c" | "v" | "ae" | "al" => 6,
            "at" | "ar" | "wa" | "wr" => 8,
            "qb" => 2,
            "x" | "e" | "nf" | "ns" | "ha" | "hb" | "hc" | "hd" | "he" | "hf" | "hg" | "hh"
            | "hi" => 0,
            _ => return None,
        };
        if arity == 0 {
            if !a.is_empty() {
                return None;
            }
            match cmd {
                "x" => {
                    self.current.push(PathCmd::Close);
                    self.cp = self.start;
                }
                "e" => self.end(),
                "nf" => self.fill = false,
                "ns" => self.stroke = false,
                _ => {}
            }
            return Some(());
        }
        if a.is_empty() || !a.len().is_multiple_of(arity) || (cmd == "m" && a.len() != 2) {
            return None;
        }
        if cmd == "qb" {
            if a.len() < 4 {
                return None;
            }
            let closed = self.current.is_empty();
            if closed {
                self.move_to((a[a.len() - 2], a[a.len() - 1]));
            }
            for i in (0..a.len() - 2).step_by(2) {
                let end = if i + 4 == a.len() {
                    (a[i + 2], a[i + 3])
                } else {
                    ((a[i] + a[i + 2]) / 2.0, (a[i + 1] + a[i + 3]) / 2.0)
                };
                let (x1, y1) = self.point((a[i], a[i + 1]));
                let (x, y) = self.point(end);
                self.current.push(PathCmd::QuadBezTo { x1, y1, x, y });
                self.cp = end;
            }
            if closed {
                self.current.push(PathCmd::Close);
                self.cp = self.start;
            }
            return Some(());
        }
        for a in a.chunks_exact(arity) {
            let relative = matches!(cmd, "t" | "r" | "v");
            let p = |i: usize| {
                (
                    a[i] + if relative { self.cp.0 } else { 0.0 },
                    a[i + 1] + if relative { self.cp.1 } else { 0.0 },
                )
            };
            match cmd {
                "m" | "t" => self.move_to(p(0)),
                "l" | "r" => self.line_to(p(0)),
                "c" | "v" => {
                    let end = p(4);
                    let (x1, y1) = self.point(p(0));
                    let (x2, y2) = self.point(p(2));
                    let (x, y) = self.point(end);
                    self.current.push(PathCmd::CubicBezTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    });
                    self.cp = end;
                }
                // §19.1.2.14 supplies start/end angles; MS-OI29500
                // §2.1.1803(b) documents Office's 1/65536-degree units.
                // VML mathematical angles run opposite Canvas's y-down angles.
                "ae" | "al" => {
                    self.arc(
                        (a[0], a[1]),
                        (a[2], a[3]),
                        -a[4] / 65536.0,
                        -(a[5] - a[4]) / 65536.0,
                        cmd == "ae",
                    )?;
                }
                "at" | "ar" | "wa" | "wr" => {
                    let center = ((a[0] + a[2]) / 2.0, (a[1] + a[3]) / 2.0);
                    let radii = ((a[2] - a[0]).abs() / 2.0, (a[3] - a[1]).abs() / 2.0);
                    let start = ((a[5] - center.1) / radii.1)
                        .atan2((a[4] - center.0) / radii.0)
                        .to_degrees();
                    let end = ((a[7] - center.1) / radii.1)
                        .atan2((a[6] - center.0) / radii.0)
                        .to_degrees();
                    let sweep = if matches!(cmd, "wa" | "wr") {
                        (end - start).rem_euclid(360.0)
                    } else {
                        -(start - end).rem_euclid(360.0)
                    };
                    self.arc(
                        center,
                        radii,
                        start,
                        if sweep == 0.0 {
                            if matches!(cmd, "wa" | "wr") {
                                360.0
                            } else {
                                -360.0
                            }
                        } else {
                            sweep
                        },
                        matches!(cmd, "at" | "wa"),
                    )?;
                }
                "qx" | "qy" => {
                    let end = (a[0], a[1]);
                    let dx = end.0 - self.cp.0;
                    let dy = end.1 - self.cp.1;
                    if dx == 0.0 || dy == 0.0 {
                        self.line_to(end);
                        continue;
                    }
                    let center = if cmd == "qx" {
                        (self.cp.0, end.1)
                    } else {
                        (end.0, self.cp.1)
                    };
                    let start = ((self.cp.1 - center.1) / dy.abs())
                        .atan2((self.cp.0 - center.0) / dx.abs())
                        .to_degrees();
                    let sweep = 90.0 * (dx * dy).signum() * if cmd == "qx" { 1.0 } else { -1.0 };
                    self.arc(center, (dx.abs(), dy.abs()), start, sweep, true)?;
                }
                _ => return None,
            }
        }
        Some(())
    }
}

type GeometryContext = (Values, (f64, f64), (f64, f64));

fn context(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    width: f64,
    height: f64,
) -> Option<GeometryContext> {
    let size = pair(attr(shape, definition, "coordsize").unwrap_or("1000,1000"))?;
    let origin = pair(attr(shape, definition, "coordorigin").unwrap_or("0,0"))?;
    if size.0 <= 0.0 || size.1 <= 0.0 {
        return None;
    }
    let limo = path_attr(shape, definition, "limo")
        .and_then(pair)
        .unwrap_or((0.0, 0.0));
    let mut values = Values {
        guides: Vec::new(),
        adjustments: attr(shape, definition, "adj")
            .unwrap_or("")
            .split(',')
            .map(|v| {
                if v.trim().is_empty() {
                    Some(0.0)
                } else {
                    v.trim().parse().ok()
                }
            })
            .collect::<Option<_>>()?,
        named: HashMap::from([
            ("width", size.0),
            ("height", size.1),
            ("xcenter", origin.0 + size.0 / 2.0),
            ("ycenter", origin.1 + size.1 / 2.0),
            ("xlimo", limo.0),
            ("ylimo", limo.1),
            (
                "hasfill",
                if vml_fill::enabled(shape, definition) {
                    1.0
                } else {
                    0.0
                },
            ),
            (
                "hasstroke",
                if resolve_vml_stroke(shape, definition).color.is_some() {
                    1.0
                } else {
                    0.0
                },
            ),
            ("pixelwidth", width * 96.0 / 72.0),
            ("pixelheight", height * 96.0 / 72.0),
            (
                "pixellinewidth",
                resolve_vml_stroke(shape, definition).width_pt * 96.0 / 72.0,
            ),
            ("emuwidth", width * 12700.0),
            ("emuheight", height * 12700.0),
            ("emuwidth2", width * 6350.0),
            ("emuheight2", height * 6350.0),
        ]),
    };
    if let Some(formulas) = vml_direct_child(shape, "formulas")
        .or_else(|| definition.and_then(|n| vml_direct_child(n, "formulas")))
    {
        for f in formulas.children().filter(|n| n.is_element()) {
            values.guides.push(values.formula(f.attribute("eqn")?)?);
        }
    }
    Some((values, size, origin))
}

/// Part 4 §19.1.2.14: retain all authored text rectangles in normalized
/// geometry coordinates, including adjustment/formula references.
pub(super) fn text_rects(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    width: f64,
    height: f64,
) -> Option<Vec<[f64; 4]>> {
    let raw = path_attr(shape, definition, "textboxrect")?;
    let (values, size, origin) = context(shape, definition, width, height)?;
    raw.split(';')
        .map(|rect| {
            let p = parameters(rect, &values)?;
            if p.len() != 4 || p[2] < p[0] || p[3] < p[1] {
                return None;
            }
            Some([
                (p[0] - origin.0) / size.0,
                (p[1] - origin.1) / size.1,
                (p[2] - origin.0) / size.0,
                (p[3] - origin.1) / size.1,
            ])
        })
        .collect()
}

pub(super) fn acquire(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    width: f64,
    height: f64,
) -> Option<Geometry> {
    let (values, size, origin) = context(shape, definition, width, height)?;
    let mut b = Builder {
        paths: Vec::new(),
        paint: Vec::new(),
        current: Vec::new(),
        fill: true,
        stroke: true,
        cp: (0.0, 0.0),
        start: (0.0, 0.0),
        origin,
        size,
    };
    let path = path_source(shape, definition)?;
    let bytes = path.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i].is_ascii_whitespace() {
            i += 1;
            continue;
        }
        let start = i;
        while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
            i += 1;
        }
        if i == start {
            return None;
        }
        // Adjacent zero-argument commands (xe, nfns) need individual tokens.
        let end = if matches!(
            bytes[start],
            b'x' | b'e' | b'm' | b'l' | b'c' | b't' | b'r' | b'v'
        ) {
            start + 1
        } else {
            (start + 2).min(i)
        };
        i = end;
        let start_args = i;
        while i < bytes.len() && !bytes[i].is_ascii_alphabetic() {
            i += 1;
        }
        let args = path[start_args..i].trim();
        let args = if args.is_empty() {
            Vec::new()
        } else {
            parameters(args, &values)?
        };
        b.command(&path[start..end], &args)?;
    }
    b.end();
    Some(Geometry {
        paths: b.paths,
        paint: b.paint,
    })
}

/// Curves and polylines ignore style extents and use the parent's coordinate
/// space (§19.1.2.3/.15). A control-point envelope is a coordinate frame only;
/// the cubic remains a cubic, with no sampling or guessed shape dimensions.
pub(super) fn primitive(shape: roxmltree::Node) -> Option<(Geometry, (f64, f64, f64, f64))> {
    let curve = shape.tag_name().name() == "curve";
    let points = if curve {
        vec![
            parse_vml_point_pt(shape.attribute("from").unwrap_or("0,0"))?,
            parse_vml_point_pt(shape.attribute("control1").unwrap_or("10,10"))?,
            parse_vml_point_pt(shape.attribute("control2").unwrap_or("20,0"))?,
            parse_vml_point_pt(shape.attribute("to").unwrap_or("30,20"))?,
        ]
    } else {
        let a = shape
            .attribute("points")
            .unwrap_or("0,0 10,10")
            .split(|c: char| c == ',' || c.is_whitespace())
            .filter(|s| !s.is_empty())
            .map(parse_vml_length_pt)
            .collect::<Option<Vec<_>>>()?;
        if a.len() < 4 || a.len() % 2 != 0 {
            return None;
        }
        a.as_chunks::<2>().0.iter().map(|&[x, y]| (x, y)).collect()
    };
    let x = points.iter().map(|p| p.0).fold(f64::INFINITY, f64::min);
    let y = points.iter().map(|p| p.1).fold(f64::INFINITY, f64::min);
    let w = points.iter().map(|p| p.0).fold(f64::NEG_INFINITY, f64::max) - x;
    let h = points.iter().map(|p| p.1).fold(f64::NEG_INFINITY, f64::max) - y;
    // Unit normalization on a degenerate axis preserves a horizontal/vertical
    // path; its physical extent remains zero, so no phantom area is introduced.
    let mut b = Builder {
        paths: Vec::new(),
        paint: Vec::new(),
        current: Vec::new(),
        fill: true,
        stroke: true,
        cp: (0.0, 0.0),
        start: (0.0, 0.0),
        origin: (x, y),
        size: (
            if w == 0.0 { 1.0 } else { w },
            if h == 0.0 { 1.0 } else { h },
        ),
    };
    b.move_to(points[0]);
    if curve {
        b.command(
            "c",
            &[
                points[1].0,
                points[1].1,
                points[2].0,
                points[2].1,
                points[3].0,
                points[3].1,
            ],
        )?;
    } else {
        for p in &points[1..] {
            b.line_to(*p);
        }
    }
    b.end();
    Some((
        Geometry {
            paths: b.paths,
            paint: b.paint,
        },
        (x, y, w, h),
    ))
}

pub(super) fn arc(shape: roxmltree::Node) -> Option<Geometry> {
    let start = shape
        .attribute("startAngle")
        .or_else(|| shape.attribute("startangle"))
        .unwrap_or("0")
        .parse::<f64>()
        .ok()?
        - 90.0;
    let end = shape
        .attribute("endAngle")
        .or_else(|| shape.attribute("endangle"))
        .unwrap_or("90")
        .parse::<f64>()
        .ok()?
        - 90.0;
    let mut b = Builder {
        paths: Vec::new(),
        paint: Vec::new(),
        current: Vec::new(),
        fill: true,
        stroke: true,
        cp: (0.0, 0.0),
        start: (0.0, 0.0),
        origin: (0.0, 0.0),
        size: (1.0, 1.0),
    };
    b.arc(
        (0.5, 0.5),
        (0.5, 0.5),
        start,
        (end - start).rem_euclid(360.0),
        false,
    )?;
    b.end();
    Some(Geometry {
        paths: b.paths,
        paint: b.paint,
    })
}

fn component(raw: &str) -> Option<(f64, f64)> {
    let raw = raw.trim();
    if let Some(n) = raw.strip_suffix('%') {
        Some((0.0, n.parse::<f64>().ok()? / 100.0))
    } else if let Some(n) = raw.strip_suffix('f') {
        Some((0.0, n.parse::<f64>().ok()? / 65536.0))
    } else if raw.chars().any(|c| c.is_ascii_alphabetic()) {
        Some((parse_vml_length_pt(raw)?, 0.0))
    } else {
        Some((0.0, raw.parse().ok()?))
    }
}
fn offset(raw: &str) -> Option<VmlShadowOffset> {
    let (x, y) = raw.split_once(',')?;
    let (x_pt, x_fraction) = component(x)?;
    let (y_pt, y_fraction) = component(y)?;
    [x_pt, y_pt, x_fraction, y_fraction]
        .iter()
        .all(|v| v.is_finite())
        .then_some(VmlShadowOffset {
            x_pt,
            y_pt,
            x_fraction,
            y_fraction,
        })
}

pub(super) fn shadow(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
) -> Option<VmlShadow> {
    let direct = vml_direct_child(shape, "shadow");
    let inherited = definition.and_then(|n| vml_direct_child(n, "shadow"));
    if direct.is_none() && inherited.is_none() {
        return None;
    }
    let attr = |name| {
        direct
            .and_then(|n| n.attribute(name))
            .or_else(|| inherited.and_then(|n| n.attribute(name)))
    };
    if attr("on").and_then(parse_vml_true_false) == Some(false)
        || path_attr(shape, definition, "shadowok").and_then(parse_vml_true_false) == Some(false)
    {
        return None;
    }
    let origin = attr("origin").unwrap_or("0,0").split_once(',')?;
    let origin = (component(origin.0)?.1, component(origin.1)?.1);
    let mut matrix: [f64; 6] = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
    if let Some(raw) = attr("matrix") {
        let parts = raw.split(',').collect::<Vec<_>>();
        if parts.len() != 6 {
            return None;
        }
        for (i, s) in parts.iter().enumerate() {
            if !s.trim().is_empty() {
                matrix[i] = parse_vml_opacity(s.trim())?;
            }
        }
    }
    // MS-OI29500 §2.1.1807(d): Office normalizes the two relative
    // spellings to perspective. They share the same acquired matrix facts.
    let kind = match attr("type").unwrap_or("single") {
        "shapeRelative" | "drawingRelative" | "shaperelative" | "drawingrelative" => "perspective",
        kind => kind,
    };
    if !matches!(kind, "single" | "double" | "emboss" | "perspective")
        || !matrix.iter().all(|v| v.is_finite())
    {
        return None;
    }
    let primary_offset = offset(attr("offset").unwrap_or("2pt,2pt"))?;
    let secondary_offset = offset(attr("offset2").unwrap_or("-2pt,-2pt"))?;
    if !origin.0.is_finite() || !origin.1.is_finite() {
        return None;
    }
    let absolute_perspective = attr("offset").unwrap_or("2pt,2pt").split(',').any(|s| {
        s.trim()
            .chars()
            .any(|c| c.is_ascii_alphabetic() && c != 'f')
    });
    Some(VmlShadow {
        kind: kind.into(),
        color: attr("color")
            .and_then(vml_color_hex6)
            .unwrap_or_else(|| "808080".into()),
        color2: attr("color2")
            .and_then(vml_color_hex6)
            .unwrap_or_else(|| "CBCBCB".into()),
        offset: primary_offset,
        offset2: secondary_offset,
        opacity: attr("opacity")
            .and_then(parse_vml_opacity)
            .filter(|v| v.is_finite())
            .unwrap_or(1.0)
            .clamp(0.0, 1.0),
        obscured: attr("obscured")
            .and_then(parse_vml_true_false)
            .unwrap_or(false),
        origin,
        matrix,
        absolute_perspective,
    })
}

/// Part 4 §19.2.2.31 and MS-OI29500 §2.1.1842: the matrix is
/// sxx,sxy,syx,syy,px,py. Image data is expressly outside this effect.
pub(super) fn skew(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> Option<VmlSkew> {
    let attr = |name| vml_decoration_attr(shape, definition, "skew", name);
    if attr("on").and_then(parse_vml_true_false) != Some(true) {
        return None;
    }
    let origin = attr("origin").unwrap_or("0,0").split_once(',')?;
    let origin = (component(origin.0)?.1, component(origin.1)?.1);
    let mut matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
    if let Some(raw) = attr("matrix") {
        let values = raw.split(',').collect::<Vec<_>>();
        if values.len() != 6 {
            return None;
        }
        for (i, value) in values.iter().enumerate() {
            if !value.trim().is_empty() {
                matrix[i] = parse_vml_opacity(value.trim())?;
            }
        }
    }
    // Part 4 calls p reciprocal units, while MS-OI29500 §2.1.1842(b)
    // pictures a reciprocal-factor expression. Affine zero-p matrices are
    // unambiguous. Preserve a diagnostic for projective factors until those
    // interpretations can be reconciled; do not guess a transfer function.
    if matrix[4] != 0.0 || matrix[5] != 0.0 {
        return None;
    }
    let raw_offset = attr("offset").unwrap_or("2pt,2pt");
    let offset = offset(raw_offset)?;
    if !matrix
        .iter()
        .chain([&origin.0, &origin.1])
        .all(|v| v.is_finite())
    {
        return None;
    }
    Some(VmlSkew {
        matrix,
        origin,
        offset,
        absolute_perspective: raw_offset.split(',').any(|s| {
            s.trim()
                .chars()
                .any(|c| c.is_ascii_alphabetic() && c != 'f')
        }),
    })
}

/// Part 4 §§19.2.2.1/.6/.17/.29/.32: the edge element supersedes
/// its parent pen completely. An instance edge still inherits the same edge
/// from its shapetype per property; explicit on=false clears inherited ink.
pub(super) fn edge_strokes(
    shape: roxmltree::Node,
    definition: Option<roxmltree::Node>,
    media: &HashMap<String, String>,
) -> Vec<VmlEdgeStroke> {
    ["left", "top", "right", "bottom", "column"]
        .into_iter()
        .filter_map(|edge| {
            let edge_node =
                |host| vml_direct_child(host, "stroke").and_then(|n| vml_direct_child(n, edge));
            let direct = edge_node(shape);
            let inherited = definition.and_then(edge_node);
            let attr = |name| {
                direct
                    .and_then(|n| n.attributes().find(|a| a.name() == name).map(|a| a.value()))
                    .or_else(|| {
                        inherited.and_then(|n| {
                            n.attributes().find(|a| a.name() == name).map(|a| a.value())
                        })
                    })
            };
            if attr("on").and_then(parse_vml_true_false) != Some(true) {
                return None;
            }
            let color = attr("color")
                .and_then(vml_color_hex6)
                .unwrap_or_else(|| "000000".into());
            let image_source = direct
                .filter(|n| n.attribute("src").is_some() || vml_image_relationship(*n).is_some())
                .or(inherited);
            let image = vml_fill::pen_image(attr, image_source, media);
            let pattern =
                (image.is_some() && attr("filltype") == Some("pattern")).then(|| Duotone {
                    clr1: attr("color2")
                        .and_then(vml_color_hex6)
                        .unwrap_or_else(|| "000000".into()),
                    clr2: color.clone(),
                });
            Some(VmlEdgeStroke {
                edge: edge.into(),
                color: vml_color_alpha(color, attr("opacity").and_then(parse_vml_opacity)),
                width_pt: attr("weight")
                    .and_then(parse_vml_stroke_weight_pt)
                    .unwrap_or(1.0),
                dash: attr("dashstyle").and_then(normalize_vml_dashstyle),
                cap: match attr("endcap").unwrap_or("flat") {
                    "round" => Some("round".into()),
                    "square" => Some("square".into()),
                    _ => Some("butt".into()),
                },
                join: attr("joinstyle").unwrap_or("round").into(),
                miter_limit: attr("miterlimit")
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(8.0),
                compound: match attr("linestyle").unwrap_or("single") {
                    "thinThin" => Some("dbl".into()),
                    "thinThick" => Some("thinThick".into()),
                    "thickThin" => Some("thickThin".into()),
                    "thickBetweenThin" => Some("tri".into()),
                    _ => None,
                },
                image,
                pattern,
            })
        })
        .collect()
}

/// Part 4 §19.1.2.14: an authored path defaults arrowok=false;
/// native primitives without a path decoration retain their own arrow ends.
pub(super) fn arrows_enabled(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    let has_path = path_source(shape, definition).is_some()
        || vml_direct_child(shape, "path").is_some()
        || definition.is_some_and(|n| vml_direct_child(n, "path").is_some());
    !has_path
        || path_attr(shape, definition, "arrowok").and_then(parse_vml_true_false) == Some(true)
}

pub(super) fn incomplete_host(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    // Part 4 §19.1 attribute sweep: editor identifiers, hyperlinks and
    // connection/handle metadata have no screen ink. Do not silently accept
    // other authored host/path modes just because the primitive is supported.
    // These retained diagnostics describe implementation limits, not Office
    // compatibility rules; available geometry and sibling content still paint.
    for host in std::iter::once(shape).chain(definition) {
        if host.attribute("style").is_some_and(|style| {
            style
                .split(';')
                .filter_map(|decl| decl.split_once(':'))
                .any(|(key, _)| {
                    !matches!(
                        key.trim(),
                        "position"
                            | "left"
                            | "top"
                            | "width"
                            | "height"
                            | "margin-left"
                            | "margin-top"
                            | "rotation"
                            | "flip"
                            | "visibility"
                            | "z-index"
                            | "mso-position-horizontal"
                            | "mso-position-vertical"
                            | "mso-position-horizontal-relative"
                            | "mso-position-vertical-relative"
                            | "mso-wrap-style"
                            | "mso-wrap-distance-left"
                            | "mso-wrap-distance-right"
                            | "mso-wrap-distance-top"
                            | "mso-wrap-distance-bottom"
                            | "mso-wrap-edited"
                    )
                })
        }) {
            return true;
        }
        for a in host.attributes() {
            if host != shape
                && shape
                    .attributes()
                    .any(|local| local.name() == a.name() && local.namespace() == a.namespace())
            {
                continue;
            }
            if !matches!(a.name(),
                "id" | "style" | "href" | "target" | "title" | "alt" | "coordsize"
                | "coordorigin" | "print" | "type" | "adj" | "path" | "filled"
                | "fillcolor" | "opacity" | "stroked" | "strokecolor" | "strokeweight"
                | "chromakey" | "src" | "cropleft" | "croptop" | "cropright" | "cropbottom"
                | "gain" | "blacklevel" | "gamma" | "grayscale" | "bilevel"
                | "from" | "to" | "control1" | "control2" | "points" | "startAngle"
                | "endAngle" | "startangle" | "endangle" | "arcsize" | "spt"
                | "spid" | "oned" | "regroupid" | "doubleclicknotify" | "button"
                | "userhidden" | "userdrawn" | "preferrelative" | "ole" | "oleicon"
                | "allowincell" | "allowoverlap" | "insetmode" | "master" | "gfxdata"
                | "equationxml" | "connectortype" | "dgmlayoutmru" | "dgmnodekind"
                | "editas" | "tableproperties" | "tablelimits"
                // Tight/through wrap polygon: anchor acquisition consumes it;
                // parse_vml_shape discloses it on unacquired legacy placement.
                | "wrapcoords")
                // Inactive optional modes do not change screen output. An
                // enabled inset pen is handled below after the path override.
                && !(matches!(a.name(), "hr" | "hrstd" | "hrnoshade" | "bullet"
                    | "forcedash" | "clip" | "cliptowrap" | "insetpen")
                    && parse_vml_true_false(a.value()) == Some(false))
                && a.name() != "insetpen"
            {
                return true;
            }
        }
    }
    false
}

/// CT_Shape/CT_Shapetype drawables are examined even when the host geometry is
/// supported. Editor handles, connection sites and locks have no independent
/// ink. Unknown extensions and paint modes not acquired by this adapter must
/// reach the existing incomplete-story mechanism, never vanish inside a leaf.
pub(super) fn incomplete_leaf(shape: roxmltree::Node, definition: Option<roxmltree::Node>) -> bool {
    if vml_image::incomplete(shape, definition) {
        return true;
    }
    if incomplete_host(shape, definition) {
        return true;
    }
    for host in std::iter::once(shape).chain(definition) {
        if let Some(path) = vml_direct_child(host, "path") {
            if path.attributes().any(|a| {
                !matches!(
                    a.name(),
                    "id" | "v"
                        | "limo"
                        | "textboxrect"
                        | "fillok"
                        | "strokeok"
                        | "shadowok"
                        | "arrowok"
                        | "gradientshapeok"
                        | "textpathok"
                        | "insetpenok"
                        | "connecttype"
                        | "connectlocs"
                        | "connectangles"
                        | "extrusionok"
                )
            }) {
                return true;
            }
        }
    }
    let effective = |name| {
        vml_direct_child(shape, name).or_else(|| definition.and_then(|n| vml_direct_child(n, name)))
    };
    for host in std::iter::once(shape).chain(definition) {
        for child in host.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "path" | "formulas" | "handles" | "fill" | "stroke" | "shadow" | "textbox"
                | "textpath" | "imagedata" | "skew" | "lock" | "wrap" | "anchorlock"
                | "ClientData" => {}
                _ => return true,
            }
        }
    }
    // Audit authored paint attributes, not just element names. Unacquired
    // image filters/pen textures and edge-specific Office strokes are real
    // ink; metadata such as title/relationship IDs has no independent paint.
    if vml_decoration_attr(shape, definition, "skew", "on").and_then(parse_vml_true_false)
        == Some(true)
        && skew(shape, definition).is_none()
    {
        return true;
    }
    for (element, allowed) in [
        (
            "skew",
            &["id", "ext", "on", "matrix", "offset", "origin"][..],
        ),
        (
            "textpath",
            &[
                "id", "style", "on", "fitshape", "fitpath", "trim", "xscale", "string",
            ][..],
        ),
        (
            "shadow",
            &[
                "id", "on", "type", "obscured", "color", "opacity", "offset", "color2", "offset2",
                "origin", "matrix",
            ][..],
        ),
        (
            "fill",
            &[
                "id",
                "type",
                "on",
                "color",
                "color2",
                "opacity",
                "opacity2",
                "src",
                "href",
                "althref",
                "size",
                "origin",
                "position",
                "aspect",
                "colors",
                "angle",
                "alignshape",
                "focus",
                "focussize",
                "focusposition",
                "method",
                "detectmouseclick",
                "title",
                "recolor",
                "rotate",
                "relid",
            ][..],
        ),
        (
            "stroke",
            &[
                "id",
                "on",
                "color",
                "weight",
                "opacity",
                "dashstyle",
                "endcap",
                "joinstyle",
                "miterlimit",
                "linestyle",
                "startarrow",
                "startarrowwidth",
                "startarrowlength",
                "endarrow",
                "endarrowwidth",
                "endarrowlength",
                "href",
                "althref",
                "title",
                "detectmouseclick",
                "insetpen",
                "filltype",
                "src",
                "relid",
                "imageaspect",
                "imagesize",
                "imagealignshape",
                "color2",
            ][..],
        ),
        (
            "imagedata",
            &[
                "id",
                "src",
                "relid",
                "pict",
                "href",
                "althref",
                "title",
                "oleid",
                "detectmouseclick",
                "cropleft",
                "croptop",
                "cropright",
                "cropbottom",
                "grayscale",
                "chromakey",
                // Effective numeric defaults and unavailable image filters
                // are handled together by vml_image::incomplete above.
                "gain",
                "gamma",
                "blacklevel",
                "bilevel",
            ][..],
        ),
    ] {
        for host in std::iter::once(shape).chain(definition) {
            if let Some(node) = vml_direct_child(host, element) {
                if node.attributes().any(|a| {
                    !allowed.contains(&a.name())
                        && !matches!(
                            (a.name(), a.value()),
                            ("filltype", "solid")
                                | ("insetpen", "f" | "false" | "0")
                                | ("gain" | "gamma", "1" | "1.0" | "65536f")
                                | ("blacklevel", "0" | "0.0" | "0f")
                                | ("grayscale" | "bilevel", "f" | "false" | "0")
                        )
                        && (if element == "imagedata" {
                            vml_image::property(shape, definition, a.name())
                        } else {
                            vml_decoration_attr(shape, definition, element, a.name())
                        }) == Some(a.value())
                }) || node.children().any(|n| {
                    if !n.is_element() {
                        return false;
                    }
                    if element != "stroke"
                        || !matches!(
                            n.tag_name().name(),
                            "left" | "top" | "right" | "bottom" | "column"
                        )
                    {
                        return true;
                    }
                    // Office border pens have the same attribute vocabulary,
                    // but do not inherit parent defaults. Audit the enabled
                    // child as well; recognizing its tag is not paint support.
                    n.attributes()
                        .any(|a| !allowed.contains(&a.name()) && a.name() != "ext")
                        || n.children().any(|c| c.is_element())
                }) {
                    return true;
                }
            }
        }
    }
    for host in std::iter::once(shape).chain(definition) {
        if let Some(text) = vml_direct_child(host, "textpath") {
            // Part 4 §19.1.2.23 also defines decorations and per-letter
            // rotation. Neither standalone nor grouped glyph acquisition
            // retains these modes yet; keep active values on the existing
            // incomplete-content path rather than claiming font-only parity
            // implements those additional text effects.
            if text.attribute("style").is_some_and(|style| {
                style
                    .split(';')
                    .filter_map(|decl| decl.split_once(':'))
                    .any(|(key, value)| {
                        !matches!(
                            key.trim(),
                            "font" | "font-family" | "font-size" | "font-style" | "font-weight"
                        ) && !matches!(
                            (key.trim(), value.trim()),
                            ("font-variant", "normal")
                                | ("text-decoration", "none")
                                | (
                                    "mso-text-shadow"
                                        | "v-rotate-letters"
                                        | "v-same-letter-heights"
                                        | "v-text-kern"
                                        | "v-text-reverse",
                                    "false" | "f" | "0"
                                )
                                | ("v-text-align", "left")
                                | ("v-text-spacing-mode", "tightening")
                                | ("v-text-spacing", "100")
                        )
                    })
            }) {
                return true;
            }
        }
    }
    if let Some(kind) = vml_fill::property(shape, definition, "type") {
        if !matches!(
            kind,
            "solid" | "gradient" | "gradientradial" | "tile" | "pattern" | "frame"
        ) {
            return true;
        }
    }
    effective("shadow").is_some()
        && vml_decoration_attr(shape, definition, "shadow", "on").and_then(parse_vml_true_false)
            != Some(false)
        && path_attr(shape, definition, "shadowok").and_then(parse_vml_true_false) != Some(false)
        && shadow(shape, definition).is_none()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn geometry(path: &str) -> Geometry {
        let xml = format!(
            r#"<v:shape xmlns:v="urn:schemas-microsoft-com:vml" coordsize="100,100" path="{path}"/>"#
        );
        let doc = XmlDoc::parse(&xml).unwrap();
        acquire(doc.root_element(), None, 100.0, 100.0).unwrap()
    }

    #[test]
    fn relative_paths_comma_omissions_and_independent_paint_sets_keep_their_meaning() {
        let g = geometry("m10,20r20,10v10,10,,,20,30xnfe t5,5l0,0ns e");
        assert_eq!(g.paths.len(), 2);
        assert!(matches!(g.paths[0][1], PathCmd::LineTo { x: 0.3, y: 0.3 }));
        assert!(matches!(
            g.paths[0][2],
            PathCmd::CubicBezTo {
                x1: 0.4,
                y1: 0.4,
                x2: 0.3,
                y2: 0.3,
                x: 0.5,
                y: 0.6
            }
        ));
        assert_eq!(g.paint[0].fill.as_deref(), Some("none"));
        assert!(g.paint[0].stroke);
        assert!(g.paint[1].fill.is_none());
        assert!(!g.paint[1].stroke);
        assert!(matches!(
            g.paths[1][0],
            PathCmd::MoveTo { x: 0.15, y: 0.25 }
        ));
    }

    #[test]
    fn arc_families_keep_direction_and_quadrant_tangents() {
        for (command, clockwise) in [("at", false), ("ar", false), ("wa", true), ("wr", true)] {
            let g = geometry(&format!("m0,0{command}0,0,100,100,100,50,50,100e"));
            let PathCmd::ArcTo {
                wr,
                hr,
                st_ang,
                sw_ang,
            } = g.paths[0][2]
            else {
                panic!("arc command")
            };
            assert_eq!((wr, hr, st_ang), (0.5, 0.5, 0.0));
            assert_eq!(sw_ang, if clockwise { 90.0 } else { -270.0 });
        }
        for (command, start, sweep) in [("qx", -90.0, 90.0), ("qy", 180.0, -90.0)] {
            let g = geometry(&format!("m0,0{command}100,100e"));
            assert!(
                matches!(g.paths[0][2], PathCmd::ArcTo { st_ang, sw_ang, .. } if st_ang == start && sw_ang == sweep)
            );
        }
        let g = geometry("al50,50,50,50,2949120,8847360e");
        assert!(matches!(
            g.paths[0][1],
            PathCmd::ArcTo {
                st_ang: -45.0,
                sw_ang: -90.0,
                ..
            }
        ));
    }

    #[test]
    fn quadratic_splines_interpolate_control_midpoints_and_close_implicit_paths() {
        let g = geometry("qb0,100,100,100,100,0e");
        assert!(matches!(g.paths[0][0], PathCmd::MoveTo { x: 1.0, y: 0.0 }));
        assert!(matches!(
            g.paths[0][1],
            PathCmd::QuadBezTo {
                x1: 0.0,
                y1: 1.0,
                x: 0.5,
                y: 1.0
            }
        ));
        assert!(matches!(
            g.paths[0][2],
            PathCmd::QuadBezTo {
                x1: 1.0,
                y1: 1.0,
                x: 1.0,
                y: 0.0
            }
        ));
        assert!(matches!(g.paths[0][3], PathCmd::Close));
    }

    #[test]
    fn formula_operations_use_specified_rounding_and_angle_units() {
        let values = Values {
            guides: vec![7.0],
            adjustments: vec![10.0],
            named: HashMap::from([("width", 100.0)]),
        };
        for (eqn, expected) in [
            ("val @0", 7.0),
            ("sum #0 width 7", 103.0),
            ("product 3 1 2", 2.0),
            ("prod -3 1 2", -1.0),
            ("mid -3 0", -1.0),
            ("abs -4", 4.0),
            ("min 3 5", 3.0),
            ("max 3 5", 5.0),
            ("if 0 3 5", 5.0),
            ("mod 3 4 0", 5.0),
            ("atan2 0 1", 5898240.0),
            ("sin 100 5898240", 100.0),
            ("cos 100 0", 100.0),
            ("cosatan2 5 4 3", 4.0),
            ("sinatan2 5 4 3", 3.0),
            ("sqrt 15", 3.0),
            ("sumangle 65536 3 1", 196608.0),
            ("ellipse 0 1 7", 7.0),
            ("tan 1 0", 0.0),
        ] {
            assert_eq!(values.formula(eqn), Some(expected), "{eqn}");
        }
        assert_eq!(values.formula("prod 1 1 0"), None);
        assert_eq!(values.formula("val @1"), None);
    }
}
