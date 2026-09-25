import { describe, expect, it } from 'vitest';
import { renderSlide } from './renderer';
import type { ShapeElement, Slide } from './types';

/**
 * ECMA-376 §20.1.9.15 `a:path@fill` / `@stroke`: each custom-geometry path
 * decides whether it is filled (and with which shading mode) and stroked.
 */

type Op = { op: 'fill' | 'stroke'; style: unknown; path: number };

function recordingCanvas(): { canvas: HTMLCanvasElement; ops: Op[] } {
  const ops: Op[] = [];
  let pathIndex = 0;
  let fillStyle: unknown = '';
  const canvas = {
    width: 960,
    height: 540,
    style: {} as CSSStyleDeclaration,
    offsetWidth: 960,
  } as HTMLCanvasElement;
  const base = {
    canvas,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1 }),
    measureText: (text: string) => ({ width: text.length * 6 }),
    beginPath: () => { pathIndex += 1; },
    fill: () => { ops.push({ op: 'fill', style: fillStyle, path: pathIndex }); },
    stroke: () => { ops.push({ op: 'stroke', style: null, path: pathIndex }); },
    get fillStyle() { return fillStyle; },
    set fillStyle(value: unknown) { fillStyle = value; },
  };
  const ctx = new Proxy(base as unknown as CanvasRenderingContext2D, {
    get(target, property, receiver) {
      if (property in target) return Reflect.get(target, property, receiver);
      return () => undefined;
    },
    set(target, property, value, receiver) {
      return Reflect.set(target, property, value, receiver);
    },
  });
  canvas.getContext = (() => ctx) as unknown as HTMLCanvasElement['getContext'];
  return { canvas, ops };
}

const square = [
  { cmd: 'moveTo', x: 0, y: 0 },
  { cmd: 'lineTo', x: 1, y: 0 },
  { cmd: 'lineTo', x: 1, y: 1 },
  { cmd: 'close' },
] as ShapeElement['custGeom'] extends (infer P)[] | null ? P : never;

function shape(custGeomPaint?: ShapeElement['custGeomPaint']): ShapeElement {
  return {
    type: 'shape',
    x: 914_400,
    y: 914_400,
    width: 914_400,
    height: 914_400,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'custGeom',
    fill: { fillType: 'solid', color: '336699' },
    stroke: { color: '000000', width: 12_700 },
    textBody: null,
    defaultTextColor: null,
    custGeom: [square, square, square],
    ...(custGeomPaint ? { custGeomPaint } : {}),
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
  } as ShapeElement;
}

async function render(element: ShapeElement): Promise<Op[]> {
  const { canvas, ops } = recordingCanvas();
  const slide: Slide = { index: 0, slideNumber: 1, background: null, elements: [element] };
  await renderSlide(canvas, slide, 9_144_000, 6_858_000, { width: 960, dpr: 1 });
  return ops;
}

describe('custom geometry per-path paint', () => {
  it('fills and strokes the combined path once without per-path flags', async () => {
    const ops = await render(shape());
    expect(ops.filter((o) => o.op === 'fill')).toHaveLength(1);
    expect(ops.filter((o) => o.op === 'stroke')).toHaveLength(1);
  });

  it('honours unfilled, unstroked and shaded paths individually', async () => {
    const ops = await render(shape([
      { fill: null, stroke: false },
      { fill: 'none', stroke: true },
      { fill: 'darken', stroke: true },
    ]));
    const fills = ops.filter((o) => o.op === 'fill');
    const strokes = ops.filter((o) => o.op === 'stroke');
    // Path 1: base fill only. Path 2: stroke only. Path 3: base fill, a
    // darkening overlay on the same path, then its stroke.
    expect(strokes).toHaveLength(2);
    expect(fills).toHaveLength(3);
    expect(fills[1].path).toBe(fills[2].path);
    expect(String(fills[2].style)).toContain('rgba(0,0,0');
    expect(strokes[0].path).not.toBe(fills[0].path);
    expect(strokes[1].path).toBe(fills[2].path);
  });
});
