import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types';

const SCALE = 1 / 12700; // 1pt rPr@spc becomes 1px.
const RC = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };

function clusterCount(text: string): number {
  return [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].length;
}

type LetterSpacingMode = 'native' | 'inert' | 'absent';

function mockCtx(mode: LetterSpacingMode = 'native') {
  let letterSpacing = '0px';
  let font = '20px serif';
  const fills: Array<{ text: string; x: number; letterSpacing: string }> = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      const clusters = clusterCount(text);
      const spacing = mode === 'native' ? Number.parseFloat(letterSpacing) || 0 : 0;
      return {
        width: clusters * 10 + clusters * spacing,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      } as TextMetrics;
    },
    fillText: (text: string, x: number) => fills.push({
      text,
      x,
      letterSpacing: mode === 'absent' ? '<absent>' : letterSpacing,
    }),
    strokeText: () => {},
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  if (mode !== 'absent') {
    Object.defineProperty(ctx, 'letterSpacing', {
      configurable: true,
      get: () => letterSpacing,
      set: (value: string) => { letterSpacing = value; },
    });
  }
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills };
}

function run(
  text: string,
  letterSpacing: number,
  color: string,
  fontFamilyEa?: string,
): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color, fontFamily: 'Arial', letterSpacing,
    ...(fontFamilyEa ? { fontFamilyEa } : {}),
  };
}

function body(
  runs: TextRunData[],
  alignment: Paragraph['alignment'] = 'l',
  rtl: boolean = false,
): TextBody {
  const paragraph = {
    alignment, marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true, rtl, runs,
  } as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [paragraph], defaultFontSize: 20,
    defaultBold: null, defaultItalic: null, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'none', vert: 'horz', autoFit: 'none',
  } as TextBody;
}

describe('pptx rPr@spc uses Canvas shaping-cluster advances', () => {
  it('applies positive and negative spacing only at internal cluster boundaries', () => {
    for (const spacing of [4, -4]) {
      const { ctx, fills } = mockCtx();
      const seen: Array<{ text: string; x: number; w: number }> = [];
      renderTextBody(
        ctx,
        body([run('AB', spacing, '000000'), run('X', 0, 'FF0000')]),
        0, 0, 400, 100, SCALE,
        null, 0, false, false, '#000000', undefined, RC,
        (info) => seen.push({ text: info.text, x: info.inShapeX, w: info.w }),
      );

      const tracked = seen.find((info) => info.text === 'AB')!;
      const following = seen.find((info) => info.text === 'X')!;
      const trackedPaint = fills.find((call) => call.text === 'AB')!;
      const followingPaint = fills.find((call) => call.text === 'X')!;
      const expectedAdvance = 20 + spacing;

      expect(tracked.w).toBeCloseTo(expectedAdvance, 6);
      expect(following.x).toBeCloseTo(tracked.x + expectedAdvance, 6);
      expect(followingPaint.x).toBeCloseTo(trackedPaint.x + expectedAdvance, 6);
      expect(trackedPaint.letterSpacing).toBe(`${spacing}px`);
    }
  });

  it('preserves the spacing seam when one authored run switches to an EA font', () => {
    for (const spacing of [4, -4]) {
      for (const alignment of ['l', 'ctr'] as const) {
        const { ctx, fills } = mockCtx();
        const seen: Array<{ text: string; x: number; w: number }> = [];
        renderTextBody(
          ctx,
          body([run('Aあ', spacing, '000000', 'Meiryo')], alignment),
          0, 0, 100, 100, SCALE,
          null, 0, false, false, '#000000', undefined, RC,
          (info) => seen.push({ text: info.text, x: info.inShapeX, w: info.w }),
        );

        const latin = fills.find((call) => call.text === 'A')!;
        const cjk = fills.find((call) => call.text === 'あ')!;
        expect(cjk.x).toBeCloseTo(latin.x + 10 + spacing, 6);
        const first = seen.find((info) => info.text === 'A')!;
        const second = seen.find((info) => info.text === 'あ')!;
        expect(second.x + second.w - first.x).toBeCloseTo(20 + spacing, 6);
      }
    }
  });

  it('preserves an EA-font spacing seam after a tab stop', () => {
    const { ctx, fills } = mockCtx();
    const tabbed = body([run('\tAあ', 4, '000000', 'Meiryo')]);
    tabbed.paragraphs[0].tabStops = [{ pos: 30 * 12700, algn: 'l' }];
    renderTextBody(
      ctx, tabbed,
      0, 0, 100, 100, SCALE,
      null, 0, false, false, '#000000', undefined, RC,
    );
    const latin = fills.find((call) => call.text === 'A')!;
    const cjk = fills.find((call) => call.text === 'あ')!;
    expect(latin.x).toBeCloseTo(30, 6);
    expect(cjk.x).toBeCloseTo(latin.x + 14, 6);
  });

  it('does not insert spacing between distinct authored runs', () => {
    const { ctx, fills } = mockCtx();
    renderTextBody(
      ctx,
      body([run('A', 4, '000000'), run('あ', 4, 'FF0000', 'Meiryo')]),
      0, 0, 100, 100, SCALE,
      null, 0, false, false, '#000000', undefined, RC,
    );
    const latin = fills.find((call) => call.text === 'A')!;
    const cjk = fills.find((call) => call.text === 'あ')!;
    expect(cjk.x).toBeCloseTo(latin.x + 10, 6);
  });

  it('keeps an EA-font spacing seam between visually reordered RTL segments', () => {
    for (const spacing of [4, -4]) {
      const { ctx, fills } = mockCtx();
      renderTextBody(
        ctx,
        body([run('אבあ', spacing, '000000', 'Meiryo')], 'r', true),
        0, 0, 100, 100, SCALE,
        null, 0, false, false, '#000000', undefined, RC,
      );
      const cjk = fills.find((call) => call.text === 'あ')!;
      const hebrew = fills.find((call) => call.text === 'אב')!;
      // Visual order is [あ, אב]. The spacing belongs at that visual seam,
      // never before the first glyph at the line edge.
      expect(hebrew.x).toBeCloseTo(cjk.x + 10 + spacing, 6);
    }
  });

  for (const [name, text, display] of [
    ['combining sequence', 'e\u0301', 'é'],
    ['emoji ZWJ sequence', '👨‍👩‍👧‍👦', '👨‍👩‍👧‍👦'],
    ['regional-indicator flag', '🇯🇵', '🇯🇵'],
  ] as const) {
    it(`keeps ${name} layout and paint aligned for positive and negative spacing`, () => {
      for (const spacing of [4, -4]) {
        const { ctx, fills } = mockCtx();
        const seen: Array<{ text: string; x: number; w: number }> = [];
        renderTextBody(
          ctx,
          body([run(text, spacing, '000000'), run('X', 0, 'FF0000')]),
          0, 0, 400, 100, SCALE,
          null, 0, false, false, '#000000', undefined, RC,
          (info) => seen.push({ text: info.text, x: info.inShapeX, w: info.w }),
        );

        const tracked = seen.find((info) => info.text === display)!;
        const following = seen.find((info) => info.text === 'X')!;
        const trackedPaint = fills.find((call) => call.text === display)!;
        const followingPaint = fills.find((call) => call.text === 'X')!;
        // Each case is one browser shaping cluster, so DrawingML has no
        // internal character boundary at which to apply spacing.
        const expectedAdvance = 10;

        expect(tracked.w).toBeCloseTo(expectedAdvance, 6);
        expect(following.x).toBeCloseTo(tracked.x + expectedAdvance, 6);
        expect(followingPaint.x).toBeCloseTo(trackedPaint.x + expectedAdvance, 6);
      }
    });
  }

  for (const mode of ['inert', 'absent'] as const) {
    it(`keeps manual layout and paint aligned when Canvas letterSpacing is ${mode}`, () => {
      const { ctx, fills } = mockCtx(mode);
      const seen: Array<{ text: string; x: number; w: number }> = [];
      renderTextBody(
        ctx,
        body([run('AB', 4, '000000'), run('X', 0, 'FF0000')]),
        0, 0, 400, 100, SCALE,
        null, 0, false, false, '#000000', undefined, RC,
        (info) => seen.push({ text: info.text, x: info.inShapeX, w: info.w }),
      );

      const tracked = seen.find((info) => info.text === 'AB')!;
      const following = seen.find((info) => info.text === 'X')!;
      const a = fills.find((call) => call.text === 'A')!;
      const b = fills.find((call) => call.text === 'B')!;
      const x = fills.find((call) => call.text === 'X')!;

      expect(fills.some((call) => call.text === 'AB')).toBe(false);
      expect(tracked.w).toBeCloseTo(24, 6);
      expect(b.x).toBeCloseTo(a.x + 14, 6);
      expect(following.x).toBeCloseTo(tracked.x + 24, 6);
      expect(x.x).toBeCloseTo(a.x + 24, 6);
    });

    it(`preserves cluster font ownership in manual tracking when Canvas letterSpacing is ${mode}`, () => {
      const { ctx, fills } = mockCtx(mode);
      const text = 'A\u0301B';
      const seen: Array<{ text: string; x: number; w: number }> = [];
      renderTextBody(ctx, body([run(text, 4, '000000'), run('X', 0, 'FF0000')]),
        0, 0, 400, 100, SCALE, null, 0, false, false, '#000000', undefined, RC,
        (info) => seen.push({ text: info.text, x: info.inShapeX, w: info.w }));
      expect(fills.map((call) => call.text)).toEqual(['Á', 'B', 'X']);
      expect(fills[1].x - fills[0].x).toBeCloseTo(14, 6);
      expect(fills[2].x - fills[0].x).toBeCloseTo(24, 6);
      expect(seen.find((info) => info.text === 'ÁB')?.w).toBeCloseTo(24, 6);
    });

    it(`manually paints fully-distributed glyph positions when Canvas letterSpacing is ${mode}`, () => {
      const { ctx, fills } = mockCtx(mode);
      renderTextBody(
        ctx,
        body([run('あいう', 4, '000000')], 'dist'),
        0, 0, 100, 100, SCALE,
        null, 0, false, false, '#000000', undefined, RC,
      );

      const glyphs = fills.filter((call) => ['あ', 'い', 'う'].includes(call.text));
      expect(fills.some((call) => call.text === 'あいう')).toBe(false);
      expect(glyphs).toHaveLength(3);
      expect(glyphs[1].x).toBeGreaterThan(glyphs[0].x + 14);
      expect(glyphs[2].x - glyphs[1].x).toBeCloseTo(glyphs[1].x - glyphs[0].x, 6);
    });
  }
});
