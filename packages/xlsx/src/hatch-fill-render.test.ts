import { describe, expect, it, vi } from 'vitest';
import { renderViewport } from './renderer.js';
import type { ShapeFill, ShapeHatchFill, Styles, Worksheet } from './types.js';

// An OpenDocument hatch (ODF 1.3 Part 3 §16.42) is stroked directly inside
// the shape's clip: no pattern bitmap is built, and the number of lines is
// bounded by the visible canvas, whatever the authored spacing and zoom.

/** A 2D affine matrix with the DOMMatrix members the painter reads. */
function matrix(a: number, b: number, c: number, d: number, e: number, f: number) {
  return {
    a, b, c, d, e, f,
    inverse() {
      const det = a * d - b * c;
      return matrix(d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det);
    },
    transformPoint(p: { x: number; y: number }) {
      return { x: a * p.x + c * p.y + e, y: b * p.x + d * p.y + f };
    },
    multiply(m: { a: number; b: number; c: number; d: number; e: number; f: number }) {
      return matrix(
        a * m.a + c * m.b, b * m.a + d * m.b,
        a * m.c + c * m.d, b * m.c + d * m.d,
        a * m.e + c * m.f + e, b * m.e + d * m.f + f,
      );
    },
  };
}
type Matrix = ReturnType<typeof matrix>;

/** The hatch colour, distinct from every grid and header colour. */
const HATCH = '#123456';

/** Paint one `size` × `size` px rectangle with `fill` on a 400 × 300 canvas. */
function paint(
  fill: ShapeFill | undefined, size = 100, images?: Map<string, CanvasImageSource | null>, preset = 'rect',
  more: Record<string, unknown> = {}, height = size,
) {
  const calls: string[] = [];
  /** Rotations applied while the hatch colour is set, in degrees. */
  const rotations: number[] = [];
  /** The transform each created pattern was given. */
  const patterns: { repetition: string; transform?: DOMMatrix2DInit }[] = [];
  /** Rectangles filled with a pattern: its offset then, and the rectangle. */
  const bands: { e?: number; f?: number; rect: number[] }[] = [];
  /** Layers drawn onto the canvas: their opacity and device position. */
  const layers: { alpha: number; x: number; y: number }[] = [];
  /** The opacity of each fill with a pattern. */
  const alphas: number[] = [];
  let lines = 0;
  const stack: Matrix[] = [];
  let current = matrix(1, 0, 0, 1, 0, 0);
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (prop === 'canvas') return { width: 400, height: 300 };
      if (prop === 'globalAlpha') return target.globalAlpha;
      if (prop === 'drawImage') {
        return (_image: unknown, x: number, y: number) => {
          layers.push({ alpha: Number(target.globalAlpha), x, y });
          calls.push('drawImage');
        };
      }
      if (prop === 'fill') {
        return () => {
          const style = target.fillStyle;
          if (style && typeof style === 'object' && 'repetition' in style) alphas.push(Number(target.globalAlpha));
          calls.push('fill');
        };
      }
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'getTransform') return () => current;
      if (prop === 'save') return () => { stack.push(current); };
      if (prop === 'restore') return () => { current = stack.pop() ?? current; };
      if (prop === 'translate') return (x: number, y: number) => { current = current.multiply(matrix(1, 0, 0, 1, x, y)); };
      if (prop === 'scale') return (x: number, y: number) => { current = current.multiply(matrix(x, 0, 0, y, 0, 0)); };
      if (prop === 'setTransform') {
        return (a = 1, b = 0, c = 0, d = 1, e = 0, f = 0) => { current = matrix(a, b, c, d, e, f); };
      }
      if (prop === 'resetTransform') return () => { current = matrix(1, 0, 0, 1, 0, 0); };
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
      if (prop === 'rotate') {
        return (r: number) => {
          if (String(target.fillStyle).toLowerCase() === HATCH) rotations.push(Math.round((r * 180) / Math.PI) || 0);
          current = current.multiply(matrix(Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0));
        };
      }
      if (prop === 'createPattern') {
        return (_image: unknown, repetition: string) => {
          const pattern: { repetition: string; transform?: DOMMatrix2DInit; setTransform(t: DOMMatrix2DInit): void } = {
            repetition,
            setTransform(t) { pattern.transform = t; },
          };
          patterns.push(pattern);
          calls.push('createPattern');
          return pattern;
        };
      }
      if (prop === 'fillRect') {
        return (...rect: number[]) => {
          if (String(target.fillStyle).toLowerCase() === HATCH) lines++;
          const style = target.fillStyle as { transform?: DOMMatrix2DInit } | undefined;
          if (style && typeof style === 'object' && 'repetition' in style) {
            bands.push({ e: style.transform?.e, f: style.transform?.f, rect });
          }
          calls.push('fillRect');
        };
      }
      if (typeof prop === 'string') return (...args: unknown[]) => { calls.push(prop); void args; };
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const ws = {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
    defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], images: [],
    shapeGroups: [{
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 0, toColOff: size * 9525, toRow: 0, toRowOff: height * 9525,
      nativeExtCx: 0, nativeExtCy: 0,
      shapes: [{
        x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
        fill,
        geom: { type: 'preset', name: preset },
        ...more,
      }],
    }],
  } as unknown as Worksheet;
  renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
    { row: 1, col: 1, rows: 10, cols: 10 }, images ? { loadedImages: images } : {});
  return { calls, lines, rotations, patterns, bands, alphas, layers };
}

const hatch = (distancePt: number, lines = 1, angle = 0): ShapeHatchFill => ({
  fillType: 'hatch', color: HATCH, angle, distancePt, lines,
});

describe('OpenDocument hatch fills', () => {
  it('draws lines at the authored spacing, one family per hatch style', () => {
    // 7.5 pt = 10 px: lines at 0, 10, …, 100 across the 100 px box.
    expect(paint(hatch(7.5)).lines).toBe(11);
    expect(paint(hatch(7.5, 2)).lines).toBe(22);
    expect(paint(hatch(7.5, 3)).lines).toBeGreaterThan(22);
  });

  it('turns the families counter-clockwise from the horizontal, as LibreOffice draws them', () => {
    // Canvas rotations are clockwise (y down): a 30° hatch has its lines at
    // -30°, the perpendicular family at 60° and the third 45° further
    // counter-clockwise, at -75° (LibreOffice 26.8: 30°, 120° and 75°).
    expect(paint(hatch(7.5, 1, 0)).rotations).toEqual([0]);
    expect(paint(hatch(7.5, 3, 30)).rotations).toEqual([-30, 60, -75]);
  });

  it('never builds a bitmap, however wide the authored spacing', () => {
    // 10000 pt is a 13 333 px period: one line, no pattern, no canvas.
    const { calls, lines } = paint(hatch(10_000, 3, 30));
    expect(lines).toBeLessThanOrEqual(3);
    expect(calls).not.toContain('createPattern');
    expect(calls).not.toContain('drawImage');
  });

  it('bounds the lines by the visible canvas at any spacing and zoom', () => {
    // A 0.01 pt spacing would be 10 000 lines over the box; lines closer
    // than two device pixels merge, so the 100 px box holds at most ~51.
    expect(paint(hatch(0.01)).lines).toBeLessThanOrEqual(52);
    // A shape far larger than the canvas (as when zoomed far in): only the
    // visible part of the 400 × 300 canvas is hatched, the lines again two
    // device pixels apart at least. Covering the whole box would take
    // millions of lines.
    const bound = 2 * (Math.ceil(Math.hypot(400, 300) / 2) + 2);
    const huge = paint(hatch(0.01, 2), 10_000_000).lines;
    expect(huge).toBeGreaterThan(100);
    expect(huge).toBeLessThanOrEqual(bound);
  });
});

describe('lines without area', () => {
  it('strokes a horizontal line, and draws nothing unstroked or without extent', () => {
    const stroke = { strokeColor: '#000000', strokeWidth: 9525 };
    const strokes = (more: Record<string, unknown>, height: number, width = 100) =>
      paint(undefined, width, undefined, 'line', more, height).calls.filter(c => c === 'stroke').length;
    // The grid strokes too: count what the shape adds.
    const grid = strokes({}, 50);
    expect(strokes(stroke, 50)).toBeGreaterThan(grid);
    expect(strokes(stroke, 0)).toBe(strokes(stroke, 50));
    expect(strokes({}, 0)).toBe(grid);
    expect(strokes(stroke, 0, 0)).toBe(strokes({}, 0, 0));
  });
});

describe('translucent hatches', () => {
  it('paints the hatch opaque on a layer and draws the layer once at its opacity', () => {
    // Lines crossing at 50% must not show 75% at their crossings: the
    // lines go to a layer (here a recording stand-in), not to the canvas.
    let layerLines = 0;
    const sizes: number[][] = [];
    class Layer {
      constructor(public width: number, public height: number) { sizes.push([width, height]); }
      getContext() {
        const state: Record<string, unknown> = {};
        return new Proxy(state, {
          get: (target, prop) => {
            if (prop === 'canvas') return this;
            if (prop === 'getTransform') return undefined;
            if (prop === 'fillRect') return () => { if (String(target.fillStyle).toLowerCase() === HATCH) layerLines++; };
            if (typeof prop === 'string') return () => {};
          },
          set: (target, prop, value) => { target[String(prop)] = value; return true; },
        });
      }
    }
    vi.stubGlobal('OffscreenCanvas', Layer);
    try {
      const painted = paint({ ...hatch(7.5, 2), alpha: 0.5 });
      expect(painted.lines).toBe(0);
      expect(layerLines).toBe(22);
      expect(painted.layers).toHaveLength(1);
      expect(painted.layers[0].alpha).toBe(0.5);
      expect(sizes).toEqual([[100, 100]]);
    } finally {
      vi.unstubAllGlobals();
    }
    // Without a layer the lines are painted directly, at the opacity.
    expect(paint({ ...hatch(7.5, 2), alpha: 0.5 }).lines).toBe(22);
  });
});

describe('OpenDocument picture fills', () => {
  // A 20 × 10 px picture in the 100 px box.
  const images = new Map<string, CanvasImageSource | null>([
    ['Pictures/p.png', { width: 20, height: 10 } as unknown as CanvasImageSource],
  ]);
  const picture = (
    layout: 'stretch' | 'tile' | 'once', align: [number, number] = [0.5, 0.5],
    more: Partial<Extract<ShapeFill, { fillType: 'image' }>> = {},
  ): ShapeFill => ({
    fillType: 'image', imagePath: 'Pictures/p.png', mimeType: 'image/png', layout, align, offset: [0, 0], ...more,
  });

  it('stretches over the shape', () => {
    expect(paint(picture('stretch'), 100, images).patterns).toEqual([
      { repetition: 'no-repeat', setTransform: expect.any(Function), transform: { a: 5, b: 0, c: 0, d: 10, e: 0, f: 0 } },
    ]);
  });

  it('paints a translucent picture at its opacity, in presets too', () => {
    expect(paint(picture('stretch'), 100, images).alphas).toEqual([1]);
    expect(paint(picture('stretch', [0.5, 0.5], { alpha: 0.5 }), 100, images).alphas).toEqual([0.5]);
    expect(paint(picture('stretch', [0.5, 0.5], { alpha: 0.5 }), 100, images, 'ellipse').alphas).toEqual([0.5]);
  });

  it('draws a single picture once at its own size, where it is aligned', () => {
    // style:repeat="no-repeat": not stretched; centred unless aligned.
    const [centred] = paint(picture('once'), 100, images).patterns;
    expect(centred.repetition).toBe('no-repeat');
    expect(centred.transform).toEqual({ a: 1, b: 0, c: 0, d: 1, e: 40, f: 45 });
    // 30 pt = 40 px wide and half the box high: flush right, at the top.
    // The offset moves repeated tiles only.
    const [corner] = paint(
      picture('once', [1, 0], { width: { pt: 30 }, height: { fraction: 0.5 }, offset: [0.25, 0.25] }), 100, images,
    ).patterns;
    expect(corner.transform).toEqual({ a: 2, b: 0, c: 0, d: 5, e: 60, f: 0 });
  });

  it('repeats tiles from the aligned one, moved by the offset', () => {
    const [tiled] = paint(picture('tile'), 100, images).patterns;
    expect(tiled.repetition).toBe('repeat');
    expect(tiled.transform).toEqual({ a: 1, b: 0, c: 0, d: 1, e: 40, f: 45 });
    // A quarter of the 20 × 10 px tile right and down from the top-left.
    const [moved] = paint(picture('tile', [0, 0], { offset: [0.25, 0.5] }), 100, images).patterns;
    expect(moved.transform).toEqual({ a: 1, b: 0, c: 0, d: 1, e: 5, f: 5 });
  });

  it('moves every second row of staggered tiles', () => {
    // 25 px tiles over the 100 px box: four row bands, the second and
    // fourth moved half a tile to the right.
    const staggered = picture('tile', [0, 0], {
      width: { fraction: 0.25 }, height: { fraction: 0.25 }, shift: { rows: 0.5 },
    });
    const { patterns, bands } = paint(staggered, 100, images);
    expect(patterns).toHaveLength(1);
    expect(bands).toEqual([
      { e: 0, f: 0, rect: [0, 0, 100, 25] },
      { e: 12.5, f: 0, rect: [0, 25, 100, 25] },
      { e: 0, f: 0, rect: [0, 50, 100, 25] },
      { e: 12.5, f: 0, rect: [0, 75, 100, 25] },
    ]);
    // A preset shape (an ellipse) staggers through the same painter.
    expect(paint(staggered, 100, images, 'ellipse').bands).toHaveLength(4);
  });

  it('paints nothing it cannot place, and never loops on a far-away origin', () => {
    // Tiles 2^54 tile heights away: band numbers have no parity there.
    const far = picture('tile', [0, 0], {
      width: { fraction: 0.25 }, height: { fraction: 0.25 }, offset: [0, 2 ** 54], shift: { rows: 0.5 },
    });
    expect(paint(far, 100, images).bands).toEqual([]);
    const unplaced = picture('tile', [0, 0], { width: { pt: Number.POSITIVE_INFINITY } });
    expect(paint(unplaced, 100, images).patterns).toEqual([]);
  });
});
