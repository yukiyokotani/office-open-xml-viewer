import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types.js';

const SCALE = 1 / 12700; // 1 point => 1 canvas unit

function recordingContext(
  actualAscent: number,
  actualDescent: number,
  exposeFontMetrics = true,
  fontAscent = actualAscent,
  fontDescent = actualDescent,
) {
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
      actualBoundingBoxAscent: actualAscent,
      actualBoundingBoxDescent: actualDescent,
      ...(exposeFontMetrics
        ? { fontBoundingBoxAscent: fontAscent, fontBoundingBoxDescent: fontDescent }
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

function twoLineBody(fontFamily: string, fontFamilyEa: string, fontSize: number): TextBody {
  const textBody = body(fontFamily, fontFamilyEa, fontSize);
  const first = textBody.paragraphs[0]!.runs[0] as TextRunData;
  textBody.paragraphs[0]!.runs = [
    { ...first, text: '一行目' },
    { type: 'break' } as unknown as TextRunData,
    { ...first, text: '二行目' },
  ];
  return textBody;
}

describe('pptx spAutoFit top anchoring', () => {
  it('keeps implicit multi-line pitch separate from the taller resolved font box (#1473)', () => {
    const fontSize = 32;
    const actualAscent = fontSize * 0.78;
    const actualDescent = fontSize * 0.18;
    const fontAscent = fontSize * 0.98;
    const fontDescent = fontSize * 0.37;
    const { ctx, draws } = recordingContext(
      actualAscent,
      actualDescent,
      true,
      fontAscent,
      fontDescent,
    );
    const textBody = twoLineBody('Meiryo', 'Meiryo', fontSize);

    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(draws.map(({ text }) => text)).toEqual(['一行目', '二行目']);
    expect(draws[0]!.y).toBeCloseTo(actualAscent, 5);
    // A resolved Canvas design box may be useful for required bounds, but it is
    // not PowerPoint's implicit baseline pitch when a:lnSpc is omitted.
    expect(draws[1]!.y - draws[0]!.y).toBeCloseTo(fontSize * 1.2, 5);

    const neededHeight = renderTextBody(
      ctx, textBody, 0, 0, 400, 46, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, undefined, true,
    );
    // The final line still reserves the live font box below the last baseline:
    // one baseline pitch plus one full resolved line box.
    expect(neededHeight).toBeCloseTo(fontSize * 1.2 + fontAscent + fontDescent, 5);
  });

  it('preserves explicit percentage line spacing under spAutoFit (#1473)', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(
      fontSize * 0.78,
      fontSize * 0.18,
      true,
      fontSize * 0.98,
      fontSize * 0.37,
    );
    const textBody = twoLineBody('Meiryo', 'Meiryo', fontSize);
    textBody.paragraphs[0]!.spaceLine = { type: 'pct', val: 150000 };

    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    // a:spcPct is based on PowerPoint's natural single-line box; only the
    // omitted a:lnSpc path is recalculated by this fix.
    expect(draws[1]!.y - draws[0]!.y).toBeCloseTo(fontSize * 1.2 * 1.5, 5);
  });

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
    const actualAscent = fontSize * 0.78;
    const actualDescent = fontSize * 0.18;
    const fontAscent = fontSize * 0.98;
    const fontDescent = fontSize * 0.37;
    const { ctx, draws } = recordingContext(
      actualAscent,
      actualDescent,
      true,
      fontAscent,
      fontDescent,
    );
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
    // PowerPoint PDF output places the visible glyph top at the top inset.
    // Canvas therefore needs the actual glyph ascent for the first baseline,
    // not the larger font-box ascent that includes leading above the ink.
    expect(draws[0]!.y).toBeCloseTo(actualAscent, 5);

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
    expect(neededHeight).toBeCloseTo(fontAscent + fontDescent, 5);
  });

  it.each(['ctr', 'b'] as const)(
    'keeps the font-box baseline for explicitly %s-anchored spAutoFit text',
    (verticalAnchor) => {
      const fontSize = 32;
      const actualAscent = fontSize * 0.78;
      const fontAscent = fontSize * 0.98;
      const fontDescent = fontSize * 0.37;
      const { ctx, draws } = recordingContext(
        actualAscent,
        fontSize * 0.18,
        true,
        fontAscent,
        fontDescent,
      );
      const textBody = { ...body('Meiryo', 'Meiryo', fontSize), verticalAnchor };
      renderTextBody(ctx, textBody, 0, 0, 400, fontAscent + fontDescent, SCALE);

      expect(draws[0]!.y).toBeCloseTo(fontAscent, 5);
    },
  );

  it('falls back to the font-box baseline when Canvas exposes no glyph ascent', () => {
    const fontSize = 32;
    const fontAscent = fontSize * 0.98;
    const fontDescent = fontSize * 0.37;
    const { ctx, draws } = recordingContext(0, 0, true, fontAscent, fontDescent);

    renderTextBody(ctx, body('Meiryo', 'Meiryo', fontSize), 0, 0, 400, 46, SCALE);

    expect(draws[0]!.y).toBeCloseTo(fontAscent, 5);
  });

  it('falls back to the natural line box when Canvas has no font metrics', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.82, fontSize * 0.18, false);
    const textBody = body('Meiryo', 'Meiryo', fontSize);
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);
    const naturalLineHeight = fontSize * 1.2;

    expect(draws[0]!.y).toBeCloseTo(naturalLineHeight * 0.8, 5);
    expect(renderTextBody(
      ctx, textBody, 0, 0, 400, 46, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, undefined, true,
    )).toBeCloseTo(naturalLineHeight, 5);
  });

  it('keeps fixed text independent of the authored font design box', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.98, fontSize * 0.37);
    const textBody = { ...body('Meiryo', 'Meiryo', fontSize), autoFit: 'none' as const };
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(draws[0]!.y).toBeCloseTo(fontSize * 0.98, 5);
  });

  it('uses live containment metrics for any tall resolved spAutoFit face', () => {
    const fontSize = 32;
    const { ctx, draws } = recordingContext(fontSize * 0.98, fontSize * 0.37);
    const textBody = body('Arial', 'Arial', fontSize);
    renderTextBody(ctx, textBody, 0, 0, 400, 46, SCALE);

    expect(draws[0]!.y).toBeCloseTo(fontSize * 0.98, 5);
  });
});
