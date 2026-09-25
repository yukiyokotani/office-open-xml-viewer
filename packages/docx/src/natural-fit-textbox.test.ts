import { describe, expect, it } from 'vitest';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types.js';

interface DrawnText { text: string; x: number; y: number }

function recordingContext() {
  let font = '10px serif';
  const drawn: DrawnText[] = [];
  const px = () => Number(/([\d.]+)px/u.exec(font)?.[1] ?? 10);
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px', fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * px(),
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
    }) as TextMetrics,
    fillText(text: string, x: number, y: number) { drawn.push({ text, x, y }); },
    strokeText(text: string, x: number, y: number) { drawn.push({ text, x, y }); },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, drawImage() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  } as unknown as CanvasRenderingContext2D;
  return { ctx, drawn };
}

function shape(alignment: ShapeText['alignment']): ShapeRun {
  const text = 'AAAA BBBB CCCC'; // natural advance: 140pt at 10pt per code point
  const block = {
    text, fontSizePt: 10, fontFamily: 'serif', alignment,
    runs: [{ text, fontSizePt: 10, fontFamily: 'serif' } as ShapeTextRun],
  } as ShapeText;
  return {
    type: 'shape', presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: [block],
  } as unknown as ShapeRun;
}

function lines(alignment: ShapeText['alignment'], width: number, scale = 1): DrawnText[][] {
  const { ctx, drawn } = recordingContext();
  acquireAndPaintShapeTextBox(shape(alignment), 100 * scale, 50 * scale,
    width * scale, 400 * scale, ctx, scale);
  const byY = new Map<number, DrawnText[]>();
  for (const event of drawn) {
    const line = byY.get(event.y) ?? [];
    line.push(event);
    byY.set(event.y, line);
  }
  return [...byY.entries()].sort((a, b) => a[0] - b[0])
    .map(([, events]) => events.sort((a, b) => a.x - b.x));
}

describe('natural word fit in shape text', () => {
  it.each(['left', 'center', 'right'] as const)(
    'wraps a marginal word before paint for %s alignment',
    (alignment) => {
      const result = lines(alignment, 138);
      expect(result.map((line) => line.map((part) => part.text.trimEnd()).join(' ')))
        .toEqual(['AAAA BBBB', 'CCCC']);
    },
  );

  it('keeps the same break at non-unit paint scale', () => {
    const result = lines('center', 138, 0.75);
    expect(result.map((line) => line.map((part) => part.text.trimEnd()).join(' ')))
      .toEqual(['AAAA BBBB', 'CCCC']);
  });

  it('preserves natural spacing when the full line fits', () => {
    const [line] = lines('left', 200);
    expect(line).toHaveLength(3);
    expect(line![1]!.x - (line![0]!.x + 40)).toBeCloseTo(10, 3);
    expect(line![2]!.x - (line![1]!.x + 40)).toBeCloseTo(10, 3);
  });
});
