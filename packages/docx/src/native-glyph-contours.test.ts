import { Canvas } from 'skia-canvas';
import { afterEach, expect, it, vi } from 'vitest';
import { acquireCanvasGlyphContours } from './native-glyph-contours.js';
import type { GlyphContourRequest } from './layout/text.js';

afterEach(() => vi.unstubAllGlobals());
const request: GlyphContourRequest = {
  spans: [{text: 'N', advancePt: 30,
    fontRoute: {familyList: 'Arial', scope: 'native', fingerprint: 'native-arial'},
    font: {weight: 400, style: 'normal'}}],
  fontSizePt: 24, ink: {xMinPt: 0, xMaxPt: 30, ascentPt: 30, descentPt: 5},
  maximumScale: 1, maximumWork: 1_048_576,
};

it('acquires diagonal glyph edges without pixel-staircase dash arclength', () => {
  vi.stubGlobal('OffscreenCanvas', Canvas);
  const contours = acquireCanvasGlyphContours(request);
  expect(contours?.some(contour => contour.some((p, i) => {
    const q = contour[(i+1)%contour.length]!;
    return Math.abs(p.x-q.x) > 3 && Math.abs(p.y-q.y) > 10;
  }))).toBe(true);
});

it('releases its bounded surface on contour-work failure and reports unresolvable ink', () => {
  const surfaces: Canvas[] = [];
  class TrackedCanvas extends Canvas {
    constructor(w: number, h: number) { super(w, h); surfaces.push(this); }
  }
  vi.stubGlobal('OffscreenCanvas', TrackedCanvas);
  expect(() => acquireCanvasGlyphContours({...request, maximumWork: 8})).toThrow(/budget/);
  expect(surfaces[0]?.width).toBe(1); expect(surfaces[0]?.height).toBe(1);
  expect(acquireCanvasGlyphContours({...request, ink: {...request.ink, xMaxPt: 1e12}})).toBeUndefined();
  expect(surfaces[1]?.width).toBe(1); expect(surfaces[1]?.height).toBe(1);
});

it('retains a surviving sub-cell glyph component when simplification would collapse its outline', () => {
  vi.stubGlobal('OffscreenCanvas', Canvas);
  const contours = acquireCanvasGlyphContours({...request,
    spans: [{...request.spans[0]!, text: '■'}],
    ink: {xMinPt: 0, xMaxPt: 30, ascentPt: 40, descentPt: 0}, maximumScale: .0016,
  });
  expect(contours?.length).toBeGreaterThan(0);
  expect(contours?.every(contour => Math.abs(contour.reduce((area, p, i) => {
    const q = contour[(i+1)%contour.length]!;
    return area + p.x*q.y - p.y*q.x;
  }, 0)) > 0)).toBe(true);
});


it('uses the acquired native base on its owned contour surface and disposes it', () => {
  const directions: string[] = [];
  const surfaces: Canvas[] = [];
  class DirectedCanvas extends Canvas {
    constructor(w: number, h: number) {
      super(w, h); surfaces.push(this);
      const context = this.getContext('2d');
      const fill = context.fillText.bind(context);
      context.fillText = (text: string, x: number, y: number) => {
        directions.push(context.direction);
        fill(text, x, y);
      };
    }
  }
  vi.stubGlobal('OffscreenCanvas', DirectedCanvas);
  expect(acquireCanvasGlyphContours({ ...request, baseDirection: 'rtl' })?.length).toBeGreaterThan(0);
  expect(directions.length).toBeGreaterThan(0);
  expect(directions.every(direction => direction === 'rtl')).toBe(true);
  expect(surfaces.every(surface => surface.width === 1 && surface.height === 1)).toBe(true);
});

it('keeps the route language across the reference-em draw when save() resets it', () => {
  const drawn: string[] = [];
  class ChromeCanvas {
    width: number; height: number;
    constructor(w: number, h: number) { this.width = w; this.height = h; }
    getContext() {
      const saved: string[] = [];
      const context = {
        lang: 'inherit', font: '', letterSpacing: '0px', fillStyle: '', textAlign: '', textBaseline: '', direction: 'ltr',
        scale() {}, translate() {},
        // Chrome resets lang on the state that stays active; restore() brings the saved one back.
        save() { saved.push(context.lang); context.lang = 'inherit'; },
        restore() { context.lang = saved.pop() ?? 'inherit'; },
        fillText() { drawn.push(context.lang); },
        getImageData: (_x: number, _y: number, w: number, h: number) => ({ data: new Uint8ClampedArray(w * h * 4) }),
      };
      return context;
    }
  }
  vi.stubGlobal('OffscreenCanvas', ChromeCanvas);
  const advances = { plan: (font: string) => ({ font, scale: 0.5 }) } as never;
  acquireCanvasGlyphContours({
    ...request, spans: [{ ...request.spans[0]!, fontRoute: { ...request.spans[0]!.fontRoute, lang: 'sr-rs' } }],
  }, advances);
  expect(drawn.length).toBeGreaterThan(0);
  expect(drawn.every(lang => lang === 'sr-rs')).toBe(true);
});
