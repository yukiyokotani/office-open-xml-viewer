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

function paintedSizes(text: TextBody, boxHeight: number): number[] {
  const { ctx, fonts } = context();
  renderTextBody(ctx, text, 0, 0, 400, boxHeight, SCALE);
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

  it('uses a stored scale even when it does not match the content', () => {
    expect(paintedSizes(body({ fontScale: 0.5 }), 140)).toEqual([12, 12, 12]);
  });
});
