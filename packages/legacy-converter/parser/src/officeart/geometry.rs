//! Explicit OfficeArt paths, independent of the source Office application.
//! MS-ODRAW 2.2.51/53-55, 2.3.6.1-9, 2.4.9/30-31, decoded into the path
//! commands of ECMA-376 20.1.9.6-8/13-16 custom geometry.
use super::unsupported;

#[derive(Clone, Copy)]
pub(crate) struct GeometryStorage<T> {
    bounds: [Option<i32>; 4],
    path: Option<u32>,
    vertices: Option<T>,
    segments: Option<T>,
}
pub(crate) type Geometry<'a> = GeometryStorage<&'a [u8]>;
#[cfg(any(test, feature = "direct-ppt"))]
pub(crate) type SpannedGeometry = GeometryStorage<super::ByteSpan>;

impl<T> Default for GeometryStorage<T> {
    fn default() -> Self {
        Self {
            bounds: [None; 4],
            path: None,
            vertices: None,
            segments: None,
        }
    }
}

impl<T: Default + Clone> GeometryStorage<T> {
    pub fn scalar(&mut self, id: u16, value: u32) -> Result<(), String> {
        match id {
            0x140..=0x143 => self.bounds[usize::from(id - 0x140)] = Some(value as i32),
            0x144 => self.path = Some(value),
            0x145 | 0x146 => {
                if value != 0 {
                    return Err(unsupported("nonzero scalar OfficeArt geometry array"));
                }
                if id == 0x145 {
                    self.vertices = Some(T::default());
                } else {
                    self.segments = Some(T::default());
                }
            }
            _ => {}
        }
        Ok(())
    }
    pub fn complex(&mut self, id: u16, data: T) {
        match id {
            0x145 => self.vertices = Some(data),
            0x146 => self.segments = Some(data),
            _ => {}
        }
    }
    /// OfficeArt property inheritance preserves explicit empty arrays as resets.
    /// Borrowed arrays never expand until a visible destination needs the path.
    #[cfg(any(test, feature = "direct-ppt"))]
    pub fn inherit(&self, parent: &Self) -> Self {
        Self {
            bounds: std::array::from_fn(|i| self.bounds[i].or(parent.bounds[i])),
            path: self.path.or(parent.path),
            vertices: self.vertices.as_ref().or(parent.vertices.as_ref()).cloned(),
            segments: self.segments.as_ref().or(parent.segments.as_ref()).cloned(),
        }
    }
}

#[cfg(any(test, feature = "direct-ppt"))]
impl SpannedGeometry {
    /// Materialize only borrowed views over the owning session's immutable
    /// backing. Bounds/inheritance/path decoding are shared with borrowed users;
    /// retaining or inheriting this metadata never copies or expands arrays.
    pub fn view<'a>(&self, backing: &'a [u8]) -> Result<Geometry<'a>, String> {
        Ok(Geometry {
            bounds: self.bounds,
            path: self.path,
            vertices: self
                .vertices
                .as_ref()
                .map(|span| span.view(backing))
                .transpose()?,
            segments: self
                .segments
                .as_ref()
                .map(|span| span.view(backing))
                .transpose()?,
        })
    }
}

impl Geometry<'_> {
    /// Borrow source arrays until a visible shape needs them. Expanded point and
    /// segment work is charged per occurrence, including repeated master layers.
    pub fn decode(&self, budget: &mut usize) -> Result<Option<Decoded>, String> {
        let Some(vertices) = self.vertices.filter(|v| !v.is_empty()) else {
            return Ok(None);
        };
        let (count, size, bytes) = array(vertices)?;
        // Compact/truncated points and guide coordinates need separate decoding;
        // do not infer their values from a different element representation.
        if size != 8 || count == 0 {
            return Ok(None);
        }
        charge(budget, count)?;
        let bounds = [
            self.bounds[0].unwrap_or(0),
            self.bounds[1].unwrap_or(0),
            self.bounds[2].unwrap_or(21600),
            self.bounds[3].unwrap_or(21600),
        ]
        .map(i64::from);
        let dx = bounds[2] - bounds[0];
        let dy = bounds[3] - bounds[1];
        let mut points = Vec::with_capacity(count);
        for point in bytes.chunks_exact(8) {
            let x = u32::from_le_bytes(point[..4].try_into().unwrap());
            let y = u32::from_le_bytes(point[4..].try_into().unwrap());
            if [x, y].iter().any(|v| (0x80000000..=0x8000007f).contains(v)) {
                return Ok(None);
            }
            let x = i64::from(x as i32) - bounds[0];
            let y = i64::from(y as i32) - bounds[1];
            // A degenerate axis is representable only when all its points are
            // constant. A unit path axis then maps every point to the same zero
            // coordinate; it is not an estimated width or visual scale factor.
            if (dx == 0 && x != 0) || (dy == 0 && y != 0) {
                return Ok(None);
            }
            points.push([
                x * if dx < 0 { -1 } else { 1 },
                y * if dy < 0 { -1 } else { 1 },
            ]);
        }
        let mut result = Decoded {
            width: dx.abs().max(1),
            height: dy.abs().max(1),
            paths: Vec::new(),
        };
        let mut reader = PathReader {
            points: &points,
            position: 0,
            current: Path::default(),
            open: false,
            result: &mut result,
        };
        let segments = self
            .segments
            .filter(|s| !s.is_empty())
            .map(array)
            .transpose()?;
        if let Some((count, size, bytes)) = segments.filter(|(n, _, _)| *n != 0) {
            if size != 2 {
                return Ok(None);
            }
            charge(budget, count)?;
            for bytes in bytes.chunks_exact(2) {
                let word = u16::from_le_bytes(bytes.try_into().unwrap());
                let count = usize::from(word & 0x1fff);
                match word >> 13 {
                    0 => reader.lines(count)?,
                    1 => reader.curves(count)?,
                    2 if count == 0 => reader.move_to()?,
                    3 if count == 1 => reader.close()?,
                    4 if count == 0 => reader.end(),
                    5 => match ((word >> 8) & 31, word & 255) {
                        (10, 0) => reader.current.fill = false,
                        (11, 0) => reader.current.stroke = false,
                        // Arc/guide/editing escapes are not guessed as lines.
                        _ => return Ok(None),
                    },
                    6 => return Ok(None),
                    _ => return Err(unsupported("invalid OfficeArt path segment")),
                }
            }
            if !reader.current.commands.is_empty() {
                return Err(unsupported("OfficeArt complex path lacks end marker"));
            }
        } else {
            match self.path.unwrap_or(1) {
                0..=3 => {
                    reader.move_to()?;
                    let count = points.len() - 1;
                    if self.path.unwrap_or(1) < 2 {
                        reader.lines(count)?;
                    } else {
                        if count % 3 != 0 {
                            return Err(unsupported("incomplete OfficeArt cubic path"));
                        }
                        reader.curves(count / 3)?;
                    }
                    if self.path.unwrap_or(1) & 1 != 0 {
                        reader.close()?;
                    }
                    reader.end();
                }
                4 => return Err(unsupported("OfficeArt complex path lacks segments")),
                _ => return Err(unsupported("invalid OfficeArt shape path")),
            }
        }
        if reader.position != points.len() {
            return Err(unsupported("unused OfficeArt path vertices"));
        }
        if result.paths.is_empty() {
            return Ok(None);
        }
        Ok(Some(result))
    }
}

type Point = [i64; 2];
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum DecodedCommand {
    Move(Point),
    Line(Point),
    Cubic([Point; 3]),
    Close,
}
struct Path {
    commands: Vec<DecodedCommand>,
    fill: bool,
    stroke: bool,
    has_open_subpath: bool,
    /// The fill flag as authored, before the open-path veto in `end`.
    authored_fill: bool,
}
impl Default for Path {
    fn default() -> Self {
        Self {
            commands: Vec::new(),
            fill: true,
            stroke: true,
            has_open_subpath: false,
            authored_fill: true,
        }
    }
}
pub(crate) struct Decoded {
    width: i64,
    height: i64,
    paths: Vec<Path>,
}

/// Decoded-path accessors for the direct-model projections.
#[cfg(any(test, feature = "direct-ppt", feature = "direct-xls"))]
pub(crate) struct DecodedPath<'a> {
    path: &'a Path,
}

#[cfg(any(test, feature = "direct-ppt", feature = "direct-xls"))]
impl Decoded {
    pub(crate) fn width(&self) -> i64 {
        self.width
    }

    pub(crate) fn height(&self) -> i64 {
        self.height
    }

    pub(crate) fn paths(&self) -> impl ExactSizeIterator<Item = DecodedPath<'_>> {
        self.paths.iter().map(|path| DecodedPath { path })
    }
}

#[cfg(any(test, feature = "direct-ppt", feature = "direct-xls"))]
impl DecodedPath<'_> {
    #[cfg(any(feature = "direct-ppt", feature = "direct-xls"))]
    pub(crate) fn fill(&self) -> bool {
        self.path.fill
    }

    #[cfg(any(feature = "direct-ppt", feature = "direct-xls"))]
    pub(crate) fn stroke(&self) -> bool {
        self.path.stroke
    }

    /// The fill flag as authored (segment escapes only), before the
    /// PowerPoint open-path veto applied by `fill`.
    #[cfg(any(feature = "direct-ppt", feature = "direct-xls"))]
    pub(crate) fn authored_fill(&self) -> bool {
        self.path.authored_fill
    }

    pub(crate) fn commands(&self) -> impl ExactSizeIterator<Item = DecodedCommand> + '_ {
        self.path.commands.iter().copied()
    }
}
struct PathReader<'a, 'b> {
    points: &'a [Point],
    position: usize,
    current: Path,
    open: bool,
    result: &'b mut Decoded,
}
impl PathReader<'_, '_> {
    fn point(&mut self) -> Result<Point, String> {
        let point = *self
            .points
            .get(self.position)
            .ok_or_else(|| unsupported("OfficeArt path vertex underflow"))?;
        self.position += 1;
        Ok(point)
    }
    fn move_to(&mut self) -> Result<(), String> {
        self.current.has_open_subpath |= self.open;
        let point = self.point()?;
        self.current.commands.push(DecodedCommand::Move(point));
        self.open = true;
        Ok(())
    }
    fn lines(&mut self, count: usize) -> Result<(), String> {
        if !self.open {
            return Err(unsupported("OfficeArt line without current point"));
        }
        for _ in 0..count {
            let point = self.point()?;
            self.current.commands.push(DecodedCommand::Line(point));
        }
        Ok(())
    }
    fn curves(&mut self, count: usize) -> Result<(), String> {
        if !self.open {
            return Err(unsupported("OfficeArt curve without current point"));
        }
        for _ in 0..count {
            let points = [self.point()?, self.point()?, self.point()?];
            self.current.commands.push(DecodedCommand::Cubic(points));
        }
        Ok(())
    }
    fn close(&mut self) -> Result<(), String> {
        if !self.open {
            return Err(unsupported("OfficeArt close without current subpath"));
        }
        self.current.commands.push(DecodedCommand::Close);
        self.open = false;
        Ok(())
    }
    fn end(&mut self) {
        if !self.current.commands.is_empty() {
            self.current.authored_fill = self.current.fill;
            self.current.fill &= !self.open && !self.current.has_open_subpath;
            self.result.paths.push(std::mem::take(&mut self.current));
        }
        self.current = Path::default();
        self.open = false;
    }
}
fn charge(budget: &mut usize, count: usize) -> Result<(), String> {
    *budget = budget
        .checked_sub(count)
        .ok_or_else(|| unsupported("OfficeArt geometry work budget exceeded"))?;
    Ok(())
}
fn array(bytes: &[u8]) -> Result<(usize, usize, &[u8]), String> {
    if bytes.len() < 6 {
        return Err(unsupported("truncated OfficeArt geometry array"));
    }
    let n = usize::from(u16::from_le_bytes(bytes[..2].try_into().unwrap()));
    let alloc = usize::from(u16::from_le_bytes(bytes[2..4].try_into().unwrap()));
    let size = u16::from_le_bytes(bytes[4..6].try_into().unwrap());
    let size = if size == 0xfff0 { 4 } else { usize::from(size) };
    if alloc < n || size == 0 || bytes.len() != 6 + n * size {
        return Err(unsupported("invalid OfficeArt geometry array dimensions"));
    }
    Ok((n, size, &bytes[6..]))
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn spanned_geometry_reuses_decode_after_move_and_preserves_empty_reset() {
        let bytes = vertices(&[[10, 20], [30, 40]]);
        let mut parent = SpannedGeometry::default();
        parent.complex(
            0x145,
            super::super::ByteSpan::new(0..bytes.len(), bytes.len(), "geometry").unwrap(),
        );
        let moved = bytes;
        let mut inherited = SpannedGeometry::default().inherit(&parent);
        let mut work = 100;
        let decoded = inherited
            .view(&moved)
            .unwrap()
            .decode(&mut work)
            .unwrap()
            .unwrap();
        assert_eq!(work, 98);
        // The original decoder remains the sole geometry interpretation.
        let expected = geometry(&moved, None).decode(&mut 100).unwrap().unwrap();
        assert_eq!(describe(&decoded), describe(&expected));
        inherited.scalar(0x145, 0).unwrap();
        assert!(inherited
            .inherit(&parent)
            .view(&moved)
            .unwrap()
            .decode(&mut 0)
            .unwrap()
            .is_none());
        assert!(parent.view(&moved[..moved.len() - 1]).is_err());
    }
    fn array_bytes(size: u16, bytes: Vec<u8>) -> Vec<u8> {
        let n = (bytes.len() / usize::from(size)) as u16;
        [
            n.to_le_bytes().to_vec(),
            n.to_le_bytes().to_vec(),
            size.to_le_bytes().to_vec(),
            bytes,
        ]
        .concat()
    }
    fn vertices(points: &[[i32; 2]]) -> Vec<u8> {
        array_bytes(
            8,
            points
                .iter()
                .flatten()
                .flat_map(|v| v.to_le_bytes())
                .collect(),
        )
    }
    fn segments(words: &[u16]) -> Vec<u8> {
        array_bytes(2, words.iter().flat_map(|v| v.to_le_bytes()).collect())
    }
    fn geometry<'a>(v: &'a [u8], s: Option<&'a [u8]>) -> Geometry<'a> {
        Geometry {
            vertices: Some(v),
            segments: s,
            ..Geometry::default()
        }
    }
    /// The decoded path: `w=W h=H`, then per path its fill/stroke flags and
    /// commands (M/L move and line, C cubic, Z close) with their points.
    fn describe(d: &Decoded) -> String {
        let mut out = format!("w={} h={}", d.width, d.height);
        for path in &d.paths {
            out.push_str(&format!(" | fill={} stroke={}:", path.fill, path.stroke));
            for command in &path.commands {
                match command {
                    DecodedCommand::Move([x, y]) => out.push_str(&format!(" M({x},{y})")),
                    DecodedCommand::Line([x, y]) => out.push_str(&format!(" L({x},{y})")),
                    DecodedCommand::Cubic(points) => {
                        out.push_str(" C");
                        for [x, y] in points {
                            out.push_str(&format!("({x},{y})"));
                        }
                    }
                    DecodedCommand::Close => out.push_str(" Z"),
                }
            }
        }
        out
    }

    /// (fill, stroke) of each decoded path.
    fn flags(d: &Decoded) -> Vec<(bool, bool)> {
        d.paths
            .iter()
            .map(|path| (path.fill, path.stroke))
            .collect()
    }
    #[test]
    fn explicit_cubic_geometry_translates_its_signed_coordinate_space() {
        let v = vertices(&[[-20, 10], [-10, 20], [30, 40], [60, 70], [-20, 10]]);
        let s = segments(&[0x4000, 0x2001, 1, 0x6001, 0x8000]);
        let mut g = geometry(&v, Some(&s));
        for (id, value) in [(0x140, -20i32), (0x141, 10), (0x142, 80), (0x143, 110)] {
            g.scalar(id, value as u32).unwrap();
        }
        // A nonempty segment list overrides shapePath, including an unused value.
        g.scalar(0x144, 99).unwrap();
        let d = g.decode(&mut 100).unwrap().unwrap();
        assert_eq!(
            describe(&d),
            "w=100 h=100 | fill=true stroke=true: M(0,0) C(10,10)(50,30)(80,60) L(0,0) Z"
        );
        assert_eq!(flags(&d), [(true, true)]);
    }
    #[test]
    fn absent_or_empty_segments_use_the_declared_polygon_or_curve_path() {
        let v = vertices(&[[0, 0], [1, 2], [3, 4], [5, 6]]);
        let empty = segments(&[]);
        for s in [None, Some(empty.as_slice())] {
            for path in 0..4 {
                let mut g = geometry(&v, s);
                g.scalar(0x144, path).unwrap();
                let d = g.decode(&mut 100).unwrap().unwrap();
                let out = describe(&d);
                assert_eq!(out.contains(" C("), path >= 2);
                assert_eq!(out.ends_with(" Z"), path & 1 != 0);
                assert_eq!(flags(&d), [(path & 1 != 0, true)]);
            }
        }
        let mut g = geometry(&v, None);
        g.scalar(0x144, 4).unwrap();
        assert!(g.decode(&mut 100).is_err());
    }
    #[test]
    fn paint_escapes_and_path_ends_preserve_separate_path_policy() {
        let v = vertices(&[[0, 0], [10, 0], [0, 0], [0, 10]]);
        let s = segments(&[
            0x4000, 1, 0x6001, 0xaa00, 0x8000, 0x4000, 1, 0x6001, 0xab00, 0x8000,
        ]);
        let d = geometry(&v, Some(&s)).decode(&mut 100).unwrap().unwrap();
        assert_eq!(flags(&d), [(false, true), (true, false)]);
    }
    #[test]
    fn rejects_truncated_arrays_and_inconsistent_allocation_headers() {
        let v = vertices(&[[0, 0], [1, 1]]);
        for n in 0..v.len() {
            if n == 0 {
                continue;
            }
            assert!(
                geometry(&v[..n], None).decode(&mut 100).is_err(),
                "length {n}"
            );
        }
        let mut bad = v.clone();
        bad[2..4].copy_from_slice(&1u16.to_le_bytes());
        assert!(geometry(&bad, None).decode(&mut 100).is_err());
        let mut bad = v.clone();
        bad[4..6].fill(0);
        assert!(geometry(&bad, None).decode(&mut 100).is_err());
        let mut bad = v.clone();
        bad.push(0);
        assert!(geometry(&bad, None).decode(&mut 100).is_err());
    }
    #[test]
    fn unsupported_compact_points_guides_and_arc_escapes_are_not_guessed() {
        let compact = [1, 0, 1, 0, 0xf0, 0xff, 0, 0, 0, 0];
        assert!(geometry(&compact, None).decode(&mut 100).unwrap().is_none());
        let guide = vertices(&[[i32::MIN, 0]]);
        assert!(geometry(&guide, None).decode(&mut 100).unwrap().is_none());
        let v = vertices(&[[0, 0]]);
        let arc = segments(&[0x4000, 0xa304, 0x8000]);
        assert!(geometry(&v, Some(&arc)).decode(&mut 100).unwrap().is_none());
    }
    #[test]
    fn rejects_vertex_underflow_unused_points_and_malformed_segment_state() {
        let v = vertices(&[[0, 0], [1, 1]]);
        for words in [
            &[0x4000, 0x2001, 0x8000][..],
            &[0x4000, 0x8000],
            &[1, 0x8000],
            &[0x4001, 1, 0x8000],
            &[0x4000, 1, 0x6000, 0x8000],
            &[0x4000, 1],
            &[0x4000, 1, 0x8001],
            &[0xe000],
        ] {
            let s = segments(words);
            assert!(
                geometry(&v, Some(&s)).decode(&mut 100).is_err(),
                "{words:?}"
            );
        }
    }
    #[test]
    fn degenerate_axes_and_reversed_bounds_are_exact_not_estimated() {
        let v = vertices(&[[10, 20], [10, 40]]);
        let mut g = geometry(&v, None);
        g.bounds = [Some(10), Some(40), Some(10), Some(20)];
        g.path = Some(0);
        let out = describe(&g.decode(&mut 100).unwrap().unwrap());
        assert!(out.starts_with("w=1 h=20"), "{out}");
        assert!(out.contains("(0,20)"), "{out}");
        assert!(out.contains("(0,0)"), "{out}");
        g.bounds[0] = Some(0);
        g.bounds[2] = Some(0);
        assert!(g.decode(&mut 100).unwrap().is_none());
    }
    #[test]
    fn scalar_array_reset_and_expansion_budgets_remain_bounded() {
        let v = vertices(&[[0, 0], [1, 1]]);
        let s = segments(&[0x4000, 1, 0x8000]);
        let mut g = geometry(&v, Some(&s));
        let mut budget = 4;
        assert!(g.decode(&mut budget).is_err());
        assert!(g.decode(&mut 5).unwrap().is_some());
        assert!(g.scalar(0x145, 4).is_err());
        g.scalar(0x145, 0).unwrap();
        assert!(g.decode(&mut 0).unwrap().is_none());
    }
    #[test]
    fn master_properties_inherit_without_copying_arrays_and_explicit_zero_resets() {
        let v = vertices(&[[10, 20], [30, 40]]);
        let mut parent = geometry(&v, None);
        parent.bounds = [Some(10), Some(20), Some(30), Some(40)];
        let mut child = Geometry::default();
        child.scalar(0x142, 50).unwrap();
        let inherited = child.inherit(&parent);
        assert_eq!(inherited.vertices.unwrap().as_ptr(), v.as_ptr());
        assert!(describe(&inherited.decode(&mut 100).unwrap().unwrap()).starts_with("w=40 h=20"));
        child.scalar(0x145, 0).unwrap();
        assert!(child.inherit(&parent).decode(&mut 0).unwrap().is_none());
    }
}
