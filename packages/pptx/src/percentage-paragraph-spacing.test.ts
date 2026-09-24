import { describe, expect, it } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types';
import type { TextRun } from '@silurus/ooxml-core';

/**
 * `a:spcBef` / `a:spcAft` may hold `a:spcPct` instead of `a:spcPts`
 * (ECMA-376 §21.1.2.2.9-.10, §21.1.2.3.11): a percentage of the text size,
 * 100000 being one line. It is measured on the same single-line base as a
 * percentage `a:lnSpc`, on the line the spacing is attached to.
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

function textRun(text: string, fontSize: number): TextRun {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize,
    color: '000000',
    fontFamily: 'Arial',
  };
}

function paragraph(text: string, fontSize: number, spacing: Partial<Paragraph>): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: { type: 'pct', val: 100000 },
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [textRun(text, fontSize)],
    ...spacing,
  } as Paragraph;
}

function baselines(
  paragraphs: Paragraph[],
  anchor = 't',
  spcFirstLastPara?: boolean,
): number[] {
  const { ctx, texts } = recordingCtx();
  const body = {
    verticalAnchor: anchor,
    spcFirstLastPara,
    paragraphs,
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
  renderTextBody(ctx, body, 0, 0, 400, 400, SCALE);
  return texts.map(({ y }) => y);
}

describe('pptx DrawingML percentage paragraph spacing', () => {
  const plain = baselines([paragraph('A', 20, {}), paragraph('B', 20, {})]);
  const pitch = plain[1] - plain[0]; // one 100% line of 20 pt text

  it('adds a percentage of the following line before a paragraph', () => {
    const ys = baselines([
      paragraph('A', 20, {}),
      paragraph('B', 20, { spaceBeforePct: 50000 }),
    ]);
    expect(ys[1] - ys[0]).toBeCloseTo(pitch * 1.5, 5);
  });

  it('measures space before on the text size of the line it precedes', () => {
    const large = baselines([paragraph('A', 20, {}), paragraph('B', 40, {})]);
    const ys = baselines([
      paragraph('A', 20, {}),
      paragraph('B', 40, { spaceBeforePct: 100000 }),
    ]);
    expect(ys[1] - ys[0] - (large[1] - large[0])).toBeCloseTo(40 * 1.2, 5);
  });

  it('adds a percentage of the last line after a paragraph', () => {
    const ys = baselines([
      paragraph('A', 20, { spaceAfterPct: 20000 }),
      paragraph('B', 20, {}),
    ]);
    expect(ys[1] - ys[0]).toBeCloseTo(pitch * 1.2, 5);
  });

  it('keeps suppressing space before on the first paragraph', () => {
    const ys = baselines([paragraph('A', 20, { spaceBeforePct: 100000 })]);
    expect(ys[0]).toBeCloseTo(plain[0], 5);
  });

  it('keeps absolute point spacing unchanged', () => {
    const ys = baselines([
      paragraph('A', 20, {}),
      paragraph('B', 20, { spaceBefore: 1200 }),
    ]);
    expect(ys[1] - ys[0]).toBeCloseTo(pitch + 12, 5);
  });

  // ECMA-376 §21.1.2.1.1 bodyPr@spcFirstLastPara (default false).
  it('suppresses the last paragraph space after unless spcFirstLastPara is set', () => {
    const bottom = baselines([paragraph('A', 20, {})], 'b');
    const suppressed = baselines([paragraph('A', 20, { spaceAfterPct: 50000 })], 'b');
    expect(suppressed[0]).toBeCloseTo(bottom[0], 5);
    const points = baselines([paragraph('A', 20, { spaceAfter: 1200 })], 'b');
    expect(points[0]).toBeCloseTo(bottom[0], 5);
    const respected = baselines([paragraph('A', 20, { spaceAfter: 1200 })], 'b', true);
    expect(bottom[0] - respected[0]).toBeCloseTo(12, 5);
  });

  it('applies the first paragraph space before when spcFirstLastPara is set', () => {
    const ys = baselines([paragraph('A', 20, { spaceBefore: 1200 })], 't', true);
    expect(ys[0] - plain[0]).toBeCloseTo(12, 5);
  });
});
