import { describe, it, expect } from 'vitest';
import { PT_TO_PX } from '@silurus/ooxml-core';
import { drawShapeText } from './renderer.js';
import type { ShapeParagraph, ShapeText, ShapeTextRun } from './types.js';

// ECMA-376 §21.1.2.2.5 <a:lnSpc> + §21.1.2.1.3 normAutofit lnSpcReduction for
// shape text bodies (drawShapeText). No real xlsx sample carries lnSpc or an
// applied autofit, so the feature is inert on the VRT corpus; these tests drive
// it directly with a mock CanvasRenderingContext2D, mirroring the (already
// verified) pptx model. The mock reports the same metrics for each family.

interface FillTextCall {
  text: string;
  x: number;
  y: number;
}

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y });
    },
    drawImage() {},
    fillStyle: '#000' as string | CanvasGradient | CanvasPattern,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function textRun(text: string, size = 20, fontFace?: string): ShapeTextRun {
  return { type: 'text', text, bold: false, italic: false, size, fontFace };
}

function para(text: string, spaceLine?: ShapeParagraph['spaceLine'], fontFace?: string): ShapeParagraph {
  return { align: 'l', runs: [textRun(text, 20, fontFace)], spaceLine };
}

/** A paragraph whose single run declares only an East-Asian face (`<a:ea>`). */
function paraEa(text: string, fontFaceEa: string): ShapeParagraph {
  return { align: 'l', runs: [{ type: 'text', text, bold: false, italic: false, size: 20, fontFaceEa }] };
}

/** A paragraph whose single run declares only a complex-script face (`<a:cs>`). */
function paraCs(text: string, fontFaceCs: string): ShapeParagraph {
  return { align: 'l', runs: [{ type: 'text', text, bold: false, italic: false, size: 20, fontFaceCs }] };
}

/** Top-anchored, wrap:none, so a line's baseline is the cumulative sum of the
 *  heights of the lines above it. With two single-line paragraphs the A→B
 *  baseline gap equals the applied per-line height of the (identical) first line
 *  (textBaseline='middle': drawY = lineTop + height/2, so gap = h0/2 + h1/2). */
function gap(paragraphs: ShapeParagraph[], overrides: Partial<ShapeText> = {}): number {
  const { ctx, calls } = makeRecordingCtx();
  const txt: ShapeText = {
    anchor: 't',
    wrap: 'none',
    lIns: 91440,
    tIns: 45720,
    rIns: 91440,
    bIns: 45720,
    paragraphs,
    ...overrides,
  };
  drawShapeText(ctx, txt, 400, 400, 1);
  const a = calls.find((c) => c.text === 'A');
  const b = calls.find((c) => c.text === 'B');
  expect(a).toBeDefined();
  expect(b).toBeDefined();
  return b!.y - a!.y;
}

describe('shape-text line spacing (§21.1.2.2.5 <a:lnSpc>) + normAutofit lnSpcReduction', () => {
  const cs = 1;
  const naturalSingle = 20 * PT_TO_PX * cs * 1.2;

  it('baseline: unspaced line height is the natural 1.2×em single line', () => {
    expect(gap([para('A'), para('B')])).toBeCloseTo(naturalSingle, 5);
  });

  it('spcPct 200% doubles the natural single-line height', () => {
    const spaced = gap([para('A', { type: 'pct', val: 200000 }), para('B', { type: 'pct', val: 200000 })]);
    expect(spaced).toBeCloseTo(naturalSingle * 2, 5);
    // And exactly 2× the unspaced reference.
    expect(spaced).toBeCloseTo(gap([para('A'), para('B')]) * 2, 5);
  });

  it.each(['Meiryo UI', 'Sakkal Majalla'])(
    'does not let the %s family change an explicit spcPct',
    (fontFace) => {
      const pct100 = { type: 'pct', val: 100000 } as const;
      const explicit = gap([
        para('A', pct100, fontFace),
        para('B', pct100, fontFace),
      ]);
      expect(explicit).toBeCloseTo(naturalSingle, 5);
    },
  );

  it('spcPts 40 makes each line an absolute 40 pt (cs-scaled) height', () => {
    const spaced = gap([para('A', { type: 'pts', val: 40 }), para('B', { type: 'pts', val: 40 })]);
    expect(spaced).toBeCloseTo(40 * PT_TO_PX * cs, 5);
  });

  it('normAutofit lnSpcReduction 0.2 scales each line to 80% of natural', () => {
    const reduced = gap([para('A'), para('B')], { autoFit: 'norm', lnSpcReduction: 0.2 });
    expect(reduced).toBeCloseTo(naturalSingle * 0.8, 5);
  });

  it('spAutoFit / noAutofit leave the natural line height unchanged (not applied)', () => {
    expect(gap([para('A'), para('B')], { autoFit: 'sp' })).toBeCloseTo(naturalSingle, 5);
    expect(gap([para('A'), para('B')], { autoFit: 'none' })).toBeCloseTo(naturalSingle, 5);
    // A stored fontScale is modeled but intentionally NOT applied to layout.
    expect(gap([para('A'), para('B')], { autoFit: 'norm', fontScale: 0.5 })).toBeCloseTo(naturalSingle, 5);
  });

  it('lnSpcReduction does NOT reduce absolute spcPts spacing (§21.1.2.1.3 note)', () => {
    // The spec: lnSpcReduction "applies only to paragraphs with percentage line
    // spacing." An absolute spcPts line height must be left as-is; only pct and
    // the implicit single (100 % percentage) are reduced.
    const pts = [para('A', { type: 'pts', val: 40 }), para('B', { type: 'pts', val: 40 })];
    const reduced = gap(pts, { autoFit: 'norm', lnSpcReduction: 0.2 });
    // Still the absolute 40 pt — NOT 40 × 0.8.
    expect(reduced).toBeCloseTo(40 * PT_TO_PX * cs, 5);
    // A pct paragraph in the same body IS reduced (control), proving the gate is
    // on the spacing type, not on the reduction being ignored entirely.
    const pct = [para('A', { type: 'pct', val: 100000 }), para('B', { type: 'pct', val: 100000 })];
    expect(gap(pct, { autoFit: 'norm', lnSpcReduction: 0.2 })).toBeCloseTo(naturalSingle * 0.8, 5);
  });
});

// Names alone cannot establish the selected font's bytes or line geometry.
// With identical Canvas measurements, all font slots keep the same line height.
describe('shape-text font slot identity', () => {
  const em = 20 * PT_TO_PX;

  it('does not assign numeric corrections to authored latin, East Asian, or complex-script names', () => {
    const expected = em * 1.2;
    expect(gap([para('A', undefined, 'Meiryo'), para('B', undefined, 'Meiryo')])).toBeCloseTo(expected, 5);
    expect(gap([paraEa('A', 'Meiryo'), paraEa('B', 'Meiryo')])).toBeCloseTo(expected, 5);
    expect(gap([paraCs('A', 'Meiryo'), paraCs('B', 'Meiryo')])).toBeCloseTo(expected, 5);
  });
});
