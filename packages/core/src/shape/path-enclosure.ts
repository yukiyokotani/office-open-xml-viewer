import type { GeometryPath } from './path-data';
import type { ResolvedStrokeGeometry } from './stroke-geometry';
export interface GeometryMatrix { readonly a: number; readonly b: number; readonly c: number; readonly d: number; readonly e: number; readonly f: number }
export interface GeometryBox { readonly x: number; readonly y: number; readonly w: number; readonly h: number }
type Point = [number, number];
type Segment = { points: Point[]; start: Point; end: Point; curved: boolean };
type Subpath = { origin: Point; segments: Segment[]; closed: boolean };
export type PathState = { paths: Subpath[]; current?: Subpath; pen?: Point };
export function appendPathStateCommand(state: PathState, matrix: GeometryMatrix, name: string, args: readonly number[]): void {
  if (!['beginPath', 'moveTo', 'lineTo', 'closePath', 'rect', 'roundRect', 'quadraticCurveTo', 'bezierCurveTo', 'arc', 'ellipse'].includes(name)) return;
  const values = (name === 'roundRect' ? args.slice(0, 4) : args).map((v) => typeof v === 'boolean' ? Number(v) : v);
  if (!values.every(Number.isFinite) || !Object.values(matrix).every(Number.isFinite)) throw new RangeError('Non-finite path enclosure input');
  const point = (x: number, y: number): Point => {
    const m = matrix; return [m.a * x + m.c * y + m.e, m.b * x + m.d * y + m.f];
  };
  const vector = (x: number, y: number): Point => {
    const m = matrix; return [m.a * x + m.c * y, m.b * x + m.d * y];
  };
  const move = (p: Point) => {
    const path: Subpath = { origin: p, segments: [], closed: false };
    state.paths.push(path); state.current = path; state.pen = p;
  };
  const line = (p: Point) => {
    if (!state.pen || !state.current) { move(p); return; }
    const tangent: Point = [p[0] - state.pen[0], p[1] - state.pen[1]];
    state.current.segments.push({ points: [state.pen, p], start: tangent, end: tangent, curved: false });
    state.pen = p;
  };
  const close = () => {
    if (!state.current) return;
    line(state.current.origin); state.current.closed = true;
  };
  const ellipse = (cx: number, cy: number, rx: number, ry: number, rotation: number, start: number, end: number, ccw: boolean) => {
    const c = Math.cos(rotation); const s = Math.sin(rotation);
    const at = (angle: number) => point(cx + rx * Math.cos(angle) * c - ry * Math.sin(angle) * s,
      cy + rx * Math.cos(angle) * s + ry * Math.sin(angle) * c);
    const tangent = (angle: number) => {
      const direction = ccw ? -1 : 1;
      return vector(direction * (-rx * Math.sin(angle) * c - ry * Math.cos(angle) * s),
        direction * (-rx * Math.sin(angle) * s + ry * Math.cos(angle) * c));
    };
    const first = at(start); if (!state.current) move(first); else line(first);
    const center = point(cx, cy); const vx = vector(rx * c, rx * s); const vy = vector(-ry * s, ry * c);
    const hx = Math.hypot(vx[0], vy[0]); const hy = Math.hypot(vx[1], vy[1]);
    state.current?.segments.push({ points: [first, at(end), [center[0] - hx, center[1] - hy], [center[0] + hx, center[1] + hy]],
      start: tangent(start), end: tangent(end), curved: true });
    state.pen = at(end);
  };
  const observe = (name: string, a: readonly number[]) => {
    switch (name) {
      case 'beginPath': state.paths = []; state.current = undefined; state.pen = undefined; break;
      case 'moveTo': move(point(a[0], a[1])); break;
      case 'lineTo': line(point(a[0], a[1])); break;
      case 'closePath': close(); break;
      case 'rect':
      case 'roundRect':
        // A round rectangle is contained by the rectangle; its smooth corners
        // cannot add miters. Keeping the rectangle is conservative for both.
        move(point(a[0], a[1])); line(point(a[0] + a[2], a[1]));
        line(point(a[0] + a[2], a[1] + a[3])); line(point(a[0], a[1] + a[3])); close();
        if (name === 'roundRect' && state.current) {
          for (const segment of state.current.segments) segment.curved = true;
        }
        break;
      case 'quadraticCurveTo':
      case 'bezierCurveTo': {
        const controls: Point[] = [];
        for (let i = 0; i < a.length; i += 2) controls.push(point(a[i], a[i + 1]));
        if (!state.pen) move(controls[0]);
        const points = [state.pen as Point, ...controls]; const last = points.length - 1;
        const tangent = (from: number, step: number): Point => {
          for (let i = from + step; i >= 0 && i <= last; i += step) {
            const v: Point = [(points[i][0] - points[from][0]) * step, (points[i][1] - points[from][1]) * step];
            if (v[0] !== 0 || v[1] !== 0) return v;
          }
          return [0, 0];
        };
        state.current?.segments.push({ points, start: tangent(0, 1), end: tangent(last, -1), curved: true });
        state.pen = points[last]; break;
      }
      case 'arc': ellipse(a[0], a[1], a[2], a[2], 0, a[3], a[4], Boolean(a[5])); break;
      case 'ellipse': ellipse(a[0], a[1], a[2], a[3], a[4], a[5], a[6], Boolean(a[7])); break;
    }
  };
  observe(name, values);
}
export function geometryPathBounds(path: GeometryPath, matrix: GeometryMatrix, stroke: ResolvedStrokeGeometry): GeometryBox | undefined {
  const state: PathState = { paths: [] };
  for (const command of path) appendPathStateCommand(state, matrix, command.op, command.args);
  return strokePathStateBounds(state, matrix, stroke);
}

export function strokePathStateBounds(state: PathState, matrix: GeometryMatrix, stroke: ResolvedStrokeGeometry): GeometryBox | undefined {
  if (state.paths.length === 0) return undefined;
  const m = matrix; const det = m.a * m.d - m.b * m.c;
  if (!Number.isFinite(det) || det === 0) throw new RangeError('Invalid path enclosure transform');
  const r = stroke.lineWidth / 2;
  const hx = r * Math.hypot(m.a, m.c); const hy = r * Math.hypot(m.b, m.d);
  const dashed = stroke.dash.some(length => length > 0);
  const squareDash = dashed && stroke.lineCap === 'square';
  let left = Infinity; let right = -Infinity; let top = Infinity; let bottom = -Infinity;
  const add = (p: Point, x = 0, y = 0) => {
    left = Math.min(left, p[0] - x); right = Math.max(right, p[0] + x);
    top = Math.min(top, p[1] - y); bottom = Math.max(bottom, p[1] + y);
  };
  const localUnit = (p: Point): Point => {
    const x = (m.d * p[0] - m.c * p[1]) / det; const y = (-m.b * p[0] + m.a * p[1]) / det;
    const length = Math.hypot(x, y); return length > 0 ? [x / length, y / length] : [0, 0];
  };
  const offset = (p: Point, x: number, y: number) => {
    add([p[0] + m.a * x + m.c * y, p[1] + m.b * x + m.d * y]);
  };
  for (const path of state.paths) {
    const segments = path.segments.filter(segment => segment.start[0] !== 0 || segment.start[1] !== 0);
    if (segments.length === 0 && path.segments.length > 0 && !path.closed && stroke.lineCap === 'round') {
      add(path.origin, hx, hy);
    }
    for (const segment of segments) {
      if (segment.curved) {
        for (const p of segment.points) add(p, squareDash ? Math.SQRT2 * hx : hx, squareDash ? Math.SQRT2 * hy : hy);
      }
      else {
        const [x, y] = localUnit(segment.start);
        for (const p of segment.points) {
          offset(p, -y * r, x * r); offset(p, y * r, -x * r);
          if (squareDash) {
            for (const direction of [-1, 1]) {
              offset(p, r * (direction * x - y), r * (direction * y + x));
              offset(p, r * (direction * x + y), r * (direction * y - x));
            }
          } else if (dashed && stroke.lineCap === 'round') add(p, hx, hy);
        }
      }
    }
    for (let i = 0; i < segments.length; i++) {
      const incoming = segments[i]; const outgoing = segments[(i + 1) % segments.length];
      if (!path.closed && i === segments.length - 1) break;
      const p = incoming.points[incoming.points.length - 1];
      if (stroke.lineJoin === 'round') { add(p, hx, hy); continue; }
      if (stroke.lineJoin !== 'miter') continue;
      const a = localUnit(incoming.end); const b = localUnit(outgoing.start);
      const divisor = 1 + a[0] * b[0] + a[1] * b[1];
      if (divisor <= 0 || Math.sqrt(2 / divisor) > stroke.miterLimit) continue;
      const x = -r * (a[1] + b[1]) / divisor; const y = r * (a[0] + b[0]) / divisor;
      offset(p, x, y); offset(p, -x, -y);
    }
    if (!path.closed && segments.length > 0) {
      const first = segments[0]; const last = segments[segments.length - 1];
      for (const [p, tangent, direction] of [[first.points[0], first.start, -1], [last.points[last.points.length - 1], last.end, 1]] as const) {
        if (stroke.lineCap === 'round') add(p, hx, hy);
        else if (stroke.lineCap === 'square') {
          const [x, y] = localUnit(tangent);
          offset(p, r * (direction * x - y), r * (direction * y + x));
          offset(p, r * (direction * x + y), r * (direction * y - x));
        }
      }
    }
  }
  if (left === Infinity) return undefined;
  const value = { x: left, y: top, w: right - left, h: bottom - top };
  if (!Object.values(value).every(Number.isFinite)) throw new RangeError('Non-finite path enclosure');
  return value;
}
