import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { attachBarTabRules } from './layout/paragraph.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.3.1.37 (tabs) + §17.15.1.25 (defaultTabStop): a paragraph's
// effective tab stops are the explicit (custom) stops PLUS automatic stops at
// every multiple of `defaultTabStop` that occur AFTER the last custom stop, all
// measured from the TEXT MARGIN (the same origin as custom stops, NOT the
// paragraph indent). A tab advances the pen to the next effective stop greater
// than the current position.
//
// sample-16 p.2 (判型/余白 block) relies on this: rows align their value column
// by using however many tabs reach a shared stop, regardless of label width. A
// labeled row ("○判型"\t\t value) and a leading-tab row (\t\t\t\t value) must
// land on the SAME column. Word's PDF puts that column at 72pt from the margin.
//
// This reproduces that with the same metrics: indent 9pt, defaultTabStop 18pt,
// custom left tab at 18pt, body font 9pt (each glyph 9px in the mock canvas).

interface FillCall {
  text: string;
  x: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number) { fills.push({ text, x }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

/** One paragraph carrying a whole "row" of the tab-aligned block. `text` uses
 *  "\t" for each <w:tab/>. Indent 9pt, one custom left tab at 18pt unless
 *  `opts` overrides the indent / custom stops. */
function rowPara(
  text: string,
  opts: { indentLeft?: number; tabStops?: { pos: number; alignment: string; leader: string }[] } = {},
): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: opts.indentLeft ?? 9, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null,
    tabStops: opts.tabStops ?? [{ pos: 18, alignment: 'left', leader: 'none' }],
    runs: [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 9, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 9, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function splitRowPara(
  texts: string[],
  tabStops: { pos: number; alignment: string; leader: string }[],
): DocParagraph {
  return {
    ...rowPara('', { indentLeft: 0, tabStops }),
    runs: texts.map((text) => ({
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 9, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: 'Times New Roman', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    })),
  } as unknown as DocParagraph;
}

function doc(paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 300, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    // §17.15.1.25 defaultTabStop = 360 twips = 18pt.
    settings: { defaultTabStop: 18 },
    body: paras.map((p) => ({ type: 'paragraph', ...p })),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('tab-stop grid (§17.15.1.25 automatic stops, margin-relative)', () => {
  // scale = width/pageWidth = 300/300 = 1 px/pt; each glyph = fontSize px.
  const COL = 72; // pt from the text margin (Word PDF: x=123.1 => 72pt from margin)

  it('aligns a labeled row and a leading-tab row on the same column', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(
      doc([
        // "abc"(3 glyphs=27px) starts at margin 9 -> ends 36; \t\t -> 54 -> 72.
        rowPara('abc\t\tV1'),
        // no label: \t\t\t\t from margin 9 -> 18 -> 36 -> 54 -> 72.
        rowPara('\t\t\t\tV2'),
      ]),
      canvas, 0, { dpr: 1, width: 300 },
    );

    const v1 = fills.find((f) => f.text === 'V1');
    const v2 = fills.find((f) => f.text === 'V2');
    expect(v1, '"V1" must be drawn').toBeDefined();
    expect(v2, '"V2" must be drawn').toBeDefined();
    // Both value tokens converge on the same column (the user-reported bug: the
    // leading-tab row shot ~36pt further right than the labeled row).
    expect(v2!.x).toBeCloseTo(v1!.x, 3);
    // ...and that column is the spec-correct stop, 72pt from the margin.
    expect(v1!.x).toBeCloseTo(COL, 3);
  });

  it('honors the document defaultTabStop interval (not a hardcoded 36pt)', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    // No custom stops, indent 0 -> pure automatic grid from the margin. With
    // defaultTabStop=18 two tabs land at 18 then 36; a hardcoded 36pt grid would
    // instead land at 36 then 72.
    await renderDocumentToCanvas(
      doc([rowPara('\t\tW', { indentLeft: 0, tabStops: [] })]),
      canvas, 0, { dpr: 1, width: 300 },
    );
    const w = fills.find((f) => f.text === 'W');
    expect(w, '"W" must be drawn').toBeDefined();
    // Two automatic tabs at the 18pt interval: 18 -> 36.
    expect(w!.x).toBeCloseTo(36, 3);
  });

  it('aligns the first halfwidth period across following runs and falls back to end', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const decimal = [{ pos: 144, alignment: 'decimal', leader: 'none' }];
    await renderDocumentToCanvas(
      doc([
        splitRowPara(['A', '\t', '12', '.', '34'], decimal),
        splitRowPara(['B', '\t', '1234.56'], decimal),
        splitRowPara(['C', '\t', '1234'], decimal),
      ]),
      canvas, 0, { dpr: 1, width: 300 },
    );

    const splitPeriod = fills.find((fill) => fill.text === '12.34');
    const whole = fills.find((fill) => fill.text === '1234.56');
    const integer = fills.find((fill) => fill.text === '1234');
    expect((splitPeriod?.x ?? 0) + 2 * 9).toBeCloseTo(144, 6);
    expect((whole?.x ?? 0) + 4 * 9).toBeCloseTo(144, 6);
    expect((integer?.x ?? 0) + 4 * 9).toBeCloseTo(144, 6);
  });

  it('uses the implicit decimal separator after the first number before a suffix', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    const decimal = [{ pos: 144, alignment: 'decimal', leader: 'none' }];
    await renderDocumentToCanvas(
      doc([
        splitRowPara(['D', '\t', '12', '3', 'kg'], decimal),
      ]),
      canvas, 0, { dpr: 1, width: 300 },
    );

    const value = fills.find((fill) => fill.text === '123kg');
    expect(value?.x).toBeCloseTo(117, 6);
    expect((value?.x ?? 0) + 3 * 9).toBeCloseTo(144, 6);
  });

  it('retains bar rules at the logical leading-edge coordinate without creating a stop', () => {
    const line = {
      range: { start: 0, end: 1 },
      bounds: { xPt: 20, yPt: 30, widthPt: 10, heightPt: 18 },
      baselinePt: 42,
      advancePt: 18,
      placements: [],
    };
    const stops = [{ pos: 72, alignment: 'bar' as const, leader: 'none' as const }];
    const ltr = attachBarTabRules([line], 20, 300, false, stops)[0];
    const rtl = attachBarTabRules([line], 20, 300, true, stops)[0];

    expect(ltr?.barTabRules?.[0]).toMatchObject({
      from: { xPt: 92, yPt: 30 },
      to: { xPt: 92, yPt: 48 },
      widthPt: 0,
    });
    expect(rtl?.barTabRules?.[0]).toMatchObject({
      from: { xPt: 248, yPt: 30 },
      to: { xPt: 248, yPt: 48 },
    });
  });
});

// #2056: a tab-stop interval or position that is positive but cannot index a
// floating-point stop grid (tiny intervals overflow the stop index; huge ones
// overflow the stop position) must not reach layout as a non-finite stop.
describe('degenerate tab-stop geometry (#2056)', () => {
  function docWith(defaultTabStop: number, paras: DocParagraph[]): DocxDocumentModel {
    return { ...doc(paras), settings: { defaultTabStop } } as DocxDocumentModel;
  }
  function numbered(text: string): DocParagraph {
    return {
      ...rowPara(text, { tabStops: [] }),
      indentLeft: 36, indentFirst: -18,
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 36,
        tab: null, suff: 'tab', jc: 'left', fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      },
    } as unknown as DocParagraph;
  }

  it.each([1e-307, 1e-300, 1e-200, 1e300, Number.MAX_VALUE])(
    'lays out plain and numbered tabs with defaultTabStop %s pt', async (value) => {
      const { canvas, fills } = makeRecordingCanvas();
      // Pens past 180 pt: 180 / 1e-307 overflows the stop index.
      const lead = 'a'.repeat(20);
      await renderDocumentToCanvas(docWith(value, [
        rowPara('a\tb', { tabStops: [] }), rowPara(`${lead}\tb`, { tabStops: [] }),
        { ...numbered('c\td'), indentLeft: 230, indentFirst: -18 } as DocParagraph,
      ]),
        canvas, 0, { dpr: 1, width: 300 });
      for (const text of ['b', 'c', 'd']) {
        const fill = fills.find((f) => f.text.includes(text));
        expect(fill, text).toBeDefined();
        expect(Number.isFinite(fill!.x)).toBe(true);
      }
      // A huge interval's first stop lies past the line, so the text after
      // the tab wraps; a tiny one must not move it back over earlier text.
      if (value > 1) return;
      const [a, b, leadFill, b2, , c, d] = fills;
      expect(b!.x).toBeGreaterThanOrEqual(a!.x + 9 - 1e-6);
      expect(b2!.x).toBeGreaterThanOrEqual(leadFill!.x + 180 - 1e-6);
      expect(d!.x).toBeGreaterThanOrEqual(c!.x + 9 - 1e-6);
    });

  it.each([1e300, Number.MAX_VALUE])('lays out a custom tab stop at %s pt', async (pos) => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(docWith(18, [
      rowPara('a\tb', { tabStops: [{ pos, alignment: 'left', leader: 'none' }] }),
      rowPara('a\tb', { tabStops: [{ pos, alignment: 'right', leader: 'none' }] }),
    ]), canvas, 0, { dpr: 1, width: 300 });
    for (const fill of fills) expect(Number.isFinite(fill.x)).toBe(true);
  });
});
