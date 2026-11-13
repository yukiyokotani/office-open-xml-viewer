import { describe, expect, it } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types.js';
import type { TextRunData } from '@silurus/ooxml-core';

const SCALE = 1 / 12700; // one canvas unit per point

function context() {
  const fonts: string[] = [];
  let font = '';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get font() { return font; }, set font(value: string) { font = value; },
    get fillStyle() { return fillStyle; }, set fillStyle(value: string) { fillStyle = value; },
    get direction() { return direction; }, set direction(value: CanvasDirection) { direction = value; },
    measureText: (value: string) => {
      const size = Number(/([\d.]+)px/.exec(font)?.[1] ?? 24);
      return {
        width: value.length * size * 0.55,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
        fontBoundingBoxAscent: size * 0.9,
        fontBoundingBoxDescent: size * 0.3,
      };
    },
    fillText: () => fonts.push(font),
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fonts };
}

function body(overrides: Partial<TextBody> = {}): TextBody {
  const run: TextRunData = {
    type: 'text', text: 'Hamburger', fontSize: 24, fontFamily: 'Arial',
    bold: null, italic: null, underline: false, strikethrough: false, color: '000000',
  };
  const paragraph = (): Paragraph => ({
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null,
    lvl: 0, bullet: { type: 'none' },
    defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], eaLnBrk: true, runs: [run],
  } as Paragraph);
  return {
    verticalAnchor: 't', paragraphs: [paragraph(), paragraph(), paragraph()],
    defaultFontSize: 24, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'none', vert: 'horz', autoFit: 'norm',
    ...overrides,
  } as TextBody;
}

function paintedSizes(text: TextBody, boxHeight: number, boxWidth = 400): number[] {
  const { ctx, fonts } = context();
  renderTextBody(ctx, text, 0, 0, boxWidth, boxHeight, SCALE);
  return fonts.map((font) => Number(/([\d.]+)px/.exec(font)?.[1]));
}

describe('normAutofit without stored fontScale (PowerPoint PDF controls)', () => {
  it('keeps authored glyph size through slight and large overflow', () => {
    for (const height of [140, 76, 44]) {
      expect(paintedSizes(body(), height)).toEqual([24, 24, 24]);
    }
  });

  it('does not derive a scale from percentage line or paragraph spacing', () => {
    const text = body();
    text.paragraphs[0]!.spaceLine = { type: 'pct', val: 120000 };
    text.paragraphs[2]!.spaceAfterPct = 50000;
    text.spcFirstLastPara = true;
    expect(paintedSizes(text, 44)).toEqual([24, 24, 24]);
  });

  it('model sources with shrinkOnOverflow take the first LibreOffice level that fits', () => {
    // Three 24 pt lines (28.8 pt each, 86.4 pt): 80 pt fits at spacing 0.9
    // alone; 60 pt needs a font level (0.85); ample room keeps 24 pt.
    expect(paintedSizes(body({ shrinkOnOverflow: true }), 140)).toEqual([24, 24, 24]);
    expect(paintedSizes(body({ shrinkOnOverflow: true }), 80)).toEqual([24, 24, 24]);
    expect(paintedSizes(body({ shrinkOnOverflow: true }), 60)).toEqual([20.4, 20.4, 20.4]);
  });

  it('shrink-to-fit fits the text columns, not one column', () => {
    // 86.4 pt in one column overflows 60 pt; two columns hold two lines
    // (57.6 pt) and one, so nothing shrinks.
    const text = body({ shrinkOnOverflow: true, numCol: 2, spcCol: 0 });
    expect(paintedSizes(text, 60)).toEqual([24, 24, 24]);
  });

  it('anchors a columned shrink-to-fit body by its tallest column', () => {
    // Two columns of 57.6 pt (two lines) and 28.8 pt in a 60 pt frame,
    // bottom-anchored: the first baseline sits in the frame, not above it.
    const text = body({ shrinkOnOverflow: true, numCol: 2, spcCol: 0, verticalAnchor: 'b' });
    const { ctx } = context();
    const baselines: number[] = [];
    (ctx as unknown as { fillText: (t: string, x: number, y: number) => void }).fillText =
      (_t, _x, y) => baselines.push(y);
    renderTextBody(ctx, text, 0, 0, 400, 60, SCALE);
    expect(baselines.length).toBe(3);
    expect(Math.min(...baselines)).toBeGreaterThan(0);
    expect(Math.max(...baselines)).toBeLessThanOrEqual(60);
  });

  it('shrink-to-fit shrinks wrapped text wider than its box', () => {
    // A 9.6 pt wide box: no single 24 pt glyph (13.2 pt) fits, however the
    // line breaks; at the 0.7 level (16.8 pt, 9.24 pt) it does.
    const text = body({ shrinkOnOverflow: true, wrap: 'square' });
    const sizes = paintedSizes(text, 2000, 9.6);
    expect(sizes.length).toBeGreaterThan(0);
    expect(new Set(sizes.map((s) => Math.round(s * 10) / 10))).toEqual(new Set([16.8]));
  });

  it('uses a stored scale even when it does not match the content', () => {
    expect(paintedSizes(body({ fontScale: 0.5 }), 140)).toEqual([12, 12, 12]);
  });
});
