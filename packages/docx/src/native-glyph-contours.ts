import { canvasFontString, createAuxCanvas, GeometryWorkBudgetError } from '@silurus/ooxml-core';
import type { GlyphContourRequest } from './layout/text.js';
import { restoreCanvasLanguage, selectCanvasLanguage } from './canvas-language.js';

type Point = Readonly<{x: number; y: number}>;

/** Canvas exposes native shaped glyph ink but no vector outline. Acquire its
 * half-coverage contours here, through the font service, before paint. Grid
 * precision targets 1/32 final pt; a 2048² surface and one million boundary
 * edges bound memory/work. This numerical raster-to-vector acquisition
 * reuses the resolved font route; complete service-owned spans preserve joining and ligatures.
 * The resulting polygons permit a physical pen AFTER arbitrary VML fitting,
 * including independent dash/join/compound/image pens. */
export function acquireCanvasGlyphContours(request: GlyphContourRequest): readonly (readonly Point[])[] | undefined {
  let ink = request.ink;
  // A Canvas implementation may quantize measured ink differently from its
  // scaled raster. Detect clipping at the scratch edge and expand acquisition
  // bounds, without changing retained metrics. Four bounded attempts cap work;
  // unresolved clipping is unavailable contour precision, never invented ink.
  for (let attempt = 0; attempt < 4; attempt++) {
    const width = ink.xMaxPt - ink.xMinPt, height = ink.ascentPt + ink.descentPt;
    if (!(width > 0 && height > 0 && request.maximumScale > 0)) return [];
    const scale = Math.min(32 * request.maximumScale, 2046 / Math.max(width, height));
    if (!Number.isFinite(scale) || scale <= 0) throw new RangeError('Invalid glyph contour scale');
    const canvas = createAuxCanvas(width * scale + 2, height * scale + 2);
    if (!canvas) return undefined;
    try {
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
      if (!ctx) throw new Error('Canvas cannot acquire native glyph contours');
      const ox = ink.xMinPt - 1 / scale, oy = -ink.ascentPt - 1 / scale;
      ctx.scale(scale, scale);
      ctx.translate(-ox, -oy);
      ctx.fillStyle = 'white';
      ctx.textAlign = 'left'; ctx.textBaseline = 'alphabetic';
      let x = 0;
      for (const span of request.spans) {
        ctx.font = canvasFontString(span.fontRoute, request.fontSizePt, span.font.weight, span.font.style);
        const previousLang = selectCanvasLanguage(ctx, span.lang);
        ctx.fillText(span.text, x, 0);
        restoreCanvasLanguage(ctx, previousLang);
        x += span.advancePt;
      }
      const w = canvas.width, h = canvas.height, stride = w + 1;
      const alpha = ctx.getImageData(0, 0, w, h).data;
      const filled = (x: number, y: number) => x >= 0 && y >= 0 && x < w && y < h && alpha[(y * w + x) * 4 + 3]! >= 128;
      const left = Array.from({length: h}, (_, y) => y).some(y => filled(0, y));
      const right = Array.from({length: h}, (_, y) => y).some(y => filled(w - 1, y));
      const top = Array.from({length: w}, (_, x) => x).some(x => filled(x, 0));
      const bottom = Array.from({length: w}, (_, x) => x).some(x => filled(x, h - 1));
      if (left || right || top || bottom) {
        ink = { xMinPt: ink.xMinPt - (left ? width / 2 : 0),
          xMaxPt: ink.xMaxPt + (right ? width / 2 : 0),
          ascentPt: ink.ascentPt + (top ? height / 2 : 0),
          descentPt: ink.descentPt + (bottom ? height / 2 : 0) };
        continue;
      }
      const next = new Map<number, number[]>();
      let work = 0;
      const budget = Math.min(request.maximumWork, 1_048_576);
      const edge = (x: number, y: number, x2: number, y2: number) => {
        if (++work > budget) throw new GeometryWorkBudgetError('Native glyph contour work budget exceeded');
        const start = y * stride + x, end = y2 * stride + x2;
        const values = next.get(start);
        if (values) values.push(end); else next.set(start, [end]);
      };
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (filled(x, y)) {
        if (!filled(x, y - 1)) edge(x, y, x + 1, y);
        if (!filled(x + 1, y)) edge(x + 1, y, x + 1, y + 1);
        if (!filled(x, y + 1)) edge(x + 1, y + 1, x, y + 1);
        if (!filled(x - 1, y)) edge(x, y + 1, x, y);
      }
      const contours: Point[][] = [];
      const direction = (a: number, b: number) => b === a + 1 ? 0 : b === a + stride ? 1 : b === a - 1 ? 2 : 3;
      while (next.size) {
        const start = next.keys().next().value as number;
        let at = start, incoming = 0;
        const vertices: number[] = [];
        do {
          vertices.push(at);
          const choices = next.get(at);
          if (!choices?.length) throw new Error('Native glyph contour is not closed');
          // At a diagonally touching vertex, follow the boundary with the ink
          // on its right; do not merge distinct glyph components or holes.
          choices.sort((a, b) => ((direction(at, a) - incoming + 3) % 4) - ((direction(at, b) - incoming + 3) % 4));
          const end = choices.shift() as number;
          if (!choices.length) next.delete(at);
          incoming = direction(at, end); at = end;
        } while (at !== start);
        const points = vertices.filter((v, i) => direction(vertices[(i + vertices.length - 1) % vertices.length]!, v)
          !== direction(v, vertices[(i + 1) % vertices.length]!))
          .map(v => ({ x: ox + (v % stride) / scale, y: oy + Math.floor(v / stride) / scale }));
        if (points.length >= 3) {
          // Pixel stair steps have Manhattan arclength even at arbitrarily fine
          // resolution. Remove that quantization before dashing/compound pens,
          // within one raster-cell diagonal (the quantization envelope).
          const simplify = (path: Point[]): Point[] => {
            const keep = new Set([0, path.length - 1]);
            const pending: Array<[number, number]> = [[0, path.length - 1]];
            while (pending.length) {
              const [a, b] = pending.pop()!;
              const p = path[a]!, q = path[b]!, dx = q.x - p.x, dy = q.y - p.y;
              let far = -1, error = Math.SQRT2 / scale;
              for (let i = a + 1; i < b; i++) {
                if (++work > budget) throw new GeometryWorkBudgetError('Native glyph contour simplification work budget exceeded');
                const r = path[i]!, t = Math.max(0, Math.min(1, ((r.x - p.x) * dx + (r.y - p.y) * dy) / (dx * dx + dy * dy || 1)));
                const distance = Math.hypot(r.x - p.x - t * dx, r.y - p.y - t * dy);
                if (distance > error) { far = i; error = distance; }
              }
              if (far >= 0) { keep.add(far); pending.push([a, far], [far, b]); }
            }
            return [...keep].sort((a, b) => a - b).map(i => path[i]!);
          };
          const simplified = simplify([...points, points[0]!]);
          const area = simplified.reduce((sum, p, i) => {
            const q = simplified[(i + 1) % simplified.length]!;
            return sum + p.x * q.y - p.y * q.x;
          }, 0);
          // A sub-cell component can legitimately be smaller than the
          // simplification tolerance. Keep its original boundary; reporting a
          // degenerate two-point loop as an acquired dot would lose its pen.
          contours.push(area !== 0 ? simplified : [...points, points[0]!]);
        }
      }
      // Nonempty measured ink can disappear below this bounded raster's
      // half-coverage threshold. That is acquisition failure, not empty text.
      return contours.length ? contours : undefined;
    } finally { canvas.width = 1; canvas.height = 1; }
  }
  return undefined;
}
