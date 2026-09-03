import { describe, expect, it } from 'vitest';
import { intendedSingleLinePx, type TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types.js';

const SCALE = 1 / 12700; // 1 point => 1 canvas unit

function recordingContext(ascent: number, descent: number, exposeFontMetrics = true) {
  const draws: Array<{ text: string; y: number }> = [];
  let font = '';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get fillStyle() { return fillStyle; },
    set fillStyle(value: string) { fillStyle = value; },
    get direction() { return direction; },
    set direction(value: CanvasDirection) { direction = value; },
    measureText: (text: string) => ({
      width: [...text].length * 20,
      actualBoundingBoxAscent: ascent,
      actualBoundingBoxDescent: descent,
      ...(exposeFontMetrics
        ? { fontBoundingBoxAscent: ascent, fontBoundingBoxDescent: descent }
        : {}),
    }),
    fillText: (text: string, _x: number, y: number) => draws.push({ text, y }),
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
  return { ctx: ctx as unknown as CanvasRenderingContext2D, draws };
}

function body(fontFamily: string, fontFamilyEa: string, fontSize: number): TextBody {
  const run: TextRunData = {
    type: 'text',
    text: '見出し',
    bold: true,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize,
    color: '000000',
    fontFamily,
    fontFamilyEa,
  };
  const paragraph: Paragraph = {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [run],
  } as Paragraph;
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph],
    defaultFontSize: fontSize,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'none',
    vert: 'horz',
    autoFit: 'sp',
  } as TextBody;
}

describe('pptx spAutoFit top anchoring', () => {
  it.each([
    ['theme-resolved heading face', 'Meiryo', 'Meiryo', 32, 46],
    ['explicit Meiryo face', 'Meiryo', 'Meiryo', 36, 50.9],
  ])('uses the measured glyph ascent for %s instead of pushing the first line down', (
    _label,
    latin,
    eastAsian,
    fontSize,
    storedHeight,
  ) => {
    // The resolved browser font can have a shorter font box than Meiryo's
    // 1.596em saved-design line. spAutoFit must recalculate from those live
    // metrics instead of reusing the document font's stale design-height floor.
    const ascent = fontSize * 0.98;
    const descent = fontSize * 0.37;
    const { ctx, draws } = recordingContext(ascent, descent);
    renderTextBody(
      ctx,
      body(latin as string, eastAsian as string, fontSize as number),
      0,
      0,
      400,
      storedHeight as number,
      SCALE,
    );

    expect(draws).toHaveLength(1);
    expect(draws[0]!.y).toBeCloseTo(ascent, 5);

    const neededHeight = renderTextBody(
      ctx,
      body(latin as string, eastAsian as string, fontSize as number),
      0,
      0,
      400,
      storedHeight as number,
      SCALE,
      null,
      0,
      false,
      false,
      '#000000',
      undefined,
      undefined,
      undefined,
      true,
    );
    expect(neededHeight).toBeCloseTo(ascent + descent, 5);
  });

  it('falls back to the authored font design box when Canvas has no font metrics', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.82, fontSize * 0.18, false);
    const textBody = body('Meiryo', 'Meiryo', fontSize);
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);
    const designedLineHeight = intendedSingleLinePx('Meiryo', fontSize);

    expect(draws[0]!.y).toBeCloseTo(designedLineHeight * 0.8, 5);
    expect(renderTextBody(
      ctx, textBody, 0, 0, 400, 46, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, undefined, true,
    )).toBeCloseTo(designedLineHeight, 5);
  });

  it('does not apply live spAutoFit metrics to fixed text', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.98, fontSize * 0.37);
    const textBody = { ...body('Meiryo', 'Meiryo', fontSize), autoFit: 'none' as const };
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(draws[0]!.y).toBeCloseTo(intendedSingleLinePx('Meiryo', fontSize) * 0.8, 5);
  });

  it('keeps the established natural line box for spAutoFit fonts without a taller design floor', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.98, fontSize * 0.37);
    const textBody = body('Arial', 'Arial', fontSize);
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(intendedSingleLinePx('Arial', fontSize)).toBeLessThan(fontSize * 1.2);
    expect(draws[0]!.y).toBeCloseTo(fontSize * 0.98, 5);
  });
});
