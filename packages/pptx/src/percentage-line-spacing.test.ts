import { describe, expect, it } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { Bullet, Paragraph, TextBody } from './types';
import type { TextRun } from '@silurus/ooxml-core';

/**
 * DrawingML percentage line spacing is authored explicitly by `a:lnSpc >
 * a:spcPct` (ECMA-376 §21.1.2.2.5 / §21.1.2.2.11). It is based on the
 * largest text size on the line. The same natural single-line base is used
 * when `a:lnSpc` is omitted; a font's design box is not the baseline pitch.
 *
 * The regression was visible with Meiryo UI: its 1.596-em design-height floor
 * expanded the pitch, while PowerPoint kept the same 120%-of-point-size natural
 * pitch as Arial for fixed and spAutoFit shapes.
 */

const SCALE = 1 / 12700; // 1 pt => 1 px

function recordingCtx(): {
  ctx: CanvasRenderingContext2D;
  texts: Array<{ text: string; y: number }>;
} {
  const texts: Array<{ text: string; y: number }> = [];
  let fillStyle = '';
  let font = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get font() { return font; },
    set font(v: string) { font = v; },
    get direction() { return direction; },
    set direction(v: CanvasDirection) { direction = v; },
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 16,
      fontBoundingBoxDescent: 4,
    }),
    fillText: (text: string, _x: number, y: number) => texts.push({ text, y }),
    fillRect: () => {},
    drawImage: () => {},
    save: () => {},
    restore: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    beginPath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    stroke: () => {},
    clip: () => {},
    rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function textRun(text: string, fontFamily: string): TextRun {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily,
    fontFamilyEa: fontFamily,
  };
}

function textBody(
  fontFamily: string,
  spaceLine: Paragraph['spaceLine'],
  bullet: Bullet = { type: 'none' },
): TextBody {
  const paragraph = (text: string): Paragraph => ({
    alignment: 'l',
    marL: bullet.type === 'none' ? 0 : 342900,
    marR: 0,
    indent: bullet.type === 'none' ? 0 : -190500,
    spaceBefore: 0,
    spaceAfter: 0,
    spaceLine,
    lvl: 0,
    bullet,
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [textRun(text, fontFamily)],
  } as Paragraph);
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph('第一段落'), paragraph('第二段落')],
    defaultFontSize: 20,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
  } as TextBody;
}

function baselinePitch(
  fontFamily: string,
  spaceLine: Paragraph['spaceLine'],
  bullet: Bullet = { type: 'none' },
): number {
  const { ctx, texts } = recordingCtx();
  renderTextBody(ctx, textBody(fontFamily, spaceLine, bullet), 0, 0, 400, 300, SCALE);
  const ys = [...new Set(texts.filter(({ text }) => text !== '•').map(({ y }) => y))]
    .sort((a, b) => a - b);
  expect(ys).toHaveLength(2);
  return ys[1] - ys[0];
}

describe('pptx DrawingML percentage line spacing', () => {
  const pct100: Paragraph['spaceLine'] = { type: 'pct', val: 100000 };
  const bullet: Bullet = {
    type: 'char',
    char: '•',
    color: null,
    sizePct: null,
    fontFamily: null,
  };

  it.each(['Meiryo UI', 'Sakkal Majalla'])(
    'does not let the %s design-height floor override an explicit 100%',
    (fontFamily) => {
      expect(baselinePitch(fontFamily, pct100))
        .toBeCloseTo(baselinePitch('Arial', pct100), 5);
    },
  );

  it('applies the same explicit-percentage rule to bulleted paragraphs', () => {
    expect(baselinePitch('Meiryo UI', pct100, bullet))
      .toBeCloseTo(baselinePitch('Arial', pct100, bullet), 5);
  });

  it.each(['Meiryo UI', 'Arial'])(
    'uses the natural 120%% pitch for omitted line spacing in %s',
    (fontFamily) => {
      expect(baselinePitch(fontFamily, null)).toBeCloseTo(20 * 1.2, 5);
    },
  );

  it('keeps absolute point line spacing independent of the font design height', () => {
    const pts18: Paragraph['spaceLine'] = { type: 'pts', val: 18 };
    expect(baselinePitch('Meiryo UI', pts18)).toBeCloseTo(18, 5);
    expect(baselinePitch('Meiryo UI', pts18))
      .toBeCloseTo(baselinePitch('Arial', pts18), 5);
  });
});
