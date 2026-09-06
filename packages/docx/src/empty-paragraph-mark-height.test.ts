import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import { paragraphMarkBelowBaselinePt, paragraphMarkLineHeight } from './line-layout.js';
import { paragraphMarkShapeInput } from './parser-model.js';
import { canvasFontString } from '@silurus/ooxml-core';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.33 (line height) + §17.3.1.29 (paragraph mark): an EMPTY
// paragraph still occupies one paragraph-mark line, whose height is the mark
// font's single-line height — the SAME computation a text line of that font and
// size uses (layoutLines → fontBoundingBox). The renderer must not size the
// empty mark line from a synthetic 1-em box while text lines use the font's real
// (often substituted) metrics; that asymmetry under-measured every empty
// paragraph. In sample-12's figure block a run of nine empty "spacer" paragraphs
// fell ~1.65 pt short per line, so the following centered caption rose into the
// image float's wrap band and its first line wrapped beside the figure instead
// of clearing it (the whole caption sits below the figure in Word).
//
// This stub reports an ASYMMETRIC ~1.15-em font box (asc 0.95 + desc 0.20),
// mimicking a substituted Latin font (a real browser/skia "Times New Roman"
// fallback reports ~1.15 em, NOT the synthetic 1.0 em). With a 1.0-em stub the
// two code paths would coincide and the regression would be invisible.

interface FillTextCall { text: string; x: number; y: number; }
interface FillRectCall { x: number; y: number; width: number; height: number; }

const ASC_RATIO = 0.95;
const DESC_RATIO = 0.2; // sum = 1.15 em ≠ the old synthetic 0.8 + 0.2 = 1.0 em

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: FillTextCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * ASC_RATIO,
        fontBoundingBoxDescent: p * DESC_RATIO,
        actualBoundingBoxAscent: p * ASC_RATIO,
        actualBoundingBoxDescent: p * DESC_RATIO,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function makeResolvedMetricCanvas(): {
  canvas: HTMLCanvasElement;
  textCalls: FillTextCall[];
  rectCalls: FillRectCall[];
} {
  let font = '10px serif';
  const textCalls: FillTextCall[] = [];
  const rectCalls: FillRectCall[] = [];
  const fontPx = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const px = fontPx();
      const local = font.includes('__ooxml_local_exact');
      const ascent = px * (local ? 1.2 : 0.8);
      const descent = px * (local ? 0.3 : 0.2);
      return {
        width: [...s].length * px * 0.5,
        fontBoundingBoxAscent: ascent,
        fontBoundingBoxDescent: descent,
        actualBoundingBoxAscent: ascent,
        actualBoundingBoxDescent: descent,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, strokeRect() {},
    clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillRect(x: number, y: number, width: number, height: number) {
      rectCalls.push({ x, y, width, height });
    },
    fillText(text: string, x: number, y: number) { textCalls.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, textCalls, rectCalls };
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 11, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocxTextRun;
}

function para(text: string): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function docOf(bodyParas: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 800,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: bodyParas as unknown as BodyElement[],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function baselines(bodyParas: DocParagraph[]): Promise<FillTextCall[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(docOf(bodyParas), canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('empty paragraph mark line height (§17.3.1.29 / §17.3.1.33)', () => {
  it('an empty paragraph advances the cursor by the SAME single-line height as a text paragraph', async () => {
    // Reference: [A, M, B] — three single-line text paragraphs. The gap between
    // A and B baselines = lineHeight(M) + lineHeight(A) (one full intervening line).
    const refCalls = await baselines([para('A'), para('M'), para('B')]);
    const refA = refCalls.find((c) => c.text === 'A')!;
    const refB = refCalls.find((c) => c.text === 'B')!;
    expect(refA).toBeDefined();
    expect(refB).toBeDefined();
    const refGap = refB.y - refA.y;

    // Subject: [A, <empty>, B] — the middle paragraph is empty. Its mark line must
    // reserve the same single-line height, so the A→B gap is identical.
    const subjCalls = await baselines([para('A'), para(''), para('B')]);
    const subjA = subjCalls.find((c) => c.text === 'A')!;
    const subjB = subjCalls.find((c) => c.text === 'B')!;
    expect(subjA).toBeDefined();
    expect(subjB).toBeDefined();
    const subjGap = subjB.y - subjA.y;

    // Before the fix subjGap was short by (1.15 − 1.0) × 11 = 1.65 pt because the
    // empty mark line used a synthetic 1-em box while the text lines used the
    // 1.15-em stub box.
    expect(subjGap).toBeCloseTo(refGap, 2);
  });

  it('uses the paragraph mark eastAsia font axis in East Asian document-grid context', () => {
    const p = {
      ...para(''),
      defaultFontFamily: 'Century',
      defaultFontFamilyEastAsia: 'Test East Asian Serif',
    } as DocParagraph;
    let font = '';
    const ctx = {
      get font() { return font; },
      set font(v: string) { font = v; },
      measureText: () => {
        const ea = font.includes('Test East Asian Serif');
        const total = ea ? 30 : 10;
        return {
          width: 0,
          fontBoundingBoxAscent: total * 0.8,
          fontBoundingBoxDescent: total * 0.2,
          actualBoundingBoxAscent: total * 0.8,
          actualBoundingBoxDescent: total * 0.2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(p, 1, { type: null, linePitchPt: null }, false, false, ctx, {})).toBe(10);
    expect(paragraphMarkLineHeight(p, 1, { type: null, linePitchPt: null }, false, true, ctx, {})).toBe(30);
  });

  it('uses an arbitrary resolved font resource metric for an empty East Asian mark', () => {
    const family = 'Arbitrary Embedded EA';
    const p = {
      ...para(''),
      defaultFontSize: 12,
      defaultFontFamily: 'Century',
      defaultFontFamilyEastAsia: family,
    } as DocParagraph;
    const ctx = {
      font: '',
      measureText: () => ({
        width: 12,
        fontBoundingBoxAscent: 10.3125,
        fontBoundingBoxDescent: 1.6875,
        actualBoundingBoxAscent: 10.3125,
        actualBoundingBoxDescent: 1.6875,
      } as TextMetrics),
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(
      p,
      1,
      { type: null, linePitchPt: null },
      false,
      false,
      ctx,
      {},
      null,
      {
        'arbitrary embedded ea': {
          family,
          eastAsianLineHeightRatio: 1.3,
        },
      },
      undefined,
      {
        fontSizePt: 12,
        fonts: { complexScript: 'Times New Roman' },
        themeFonts: {
          ascii: family,
          highAnsi: family,
          eastAsia: family,
        },
        themeFontPresence: {
          ascii: true,
          highAnsi: true,
          eastAsia: true,
          complexScript: false,
        },
        weight: 400,
        style: 'normal',
        complexScript: false,
        fontHint: 'eastAsia',
        eastAsiaLanguage: 'ja-jp',
      },
    )).toBeCloseTo(15.6, 5);
  });

  it('retains the observed MS Mincho height for an unresolved empty East Asian mark', () => {
    const p = {
      ...para(''),
      defaultFontSize: 12,
      defaultFontFamily: 'Century',
      defaultFontFamilyEastAsia: 'ＭＳ 明朝',
    } as DocParagraph;
    const ctx = {
      font: '',
      measureText: () => ({
        width: 12,
        fontBoundingBoxAscent: 10.3125,
        fontBoundingBoxDescent: 1.6875,
        actualBoundingBoxAscent: 10.3125,
        actualBoundingBoxDescent: 1.6875,
      } as TextMetrics),
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(
      p,
      1,
      { type: null, linePitchPt: null },
      false,
      false,
      ctx,
      {},
      null,
      {},
      undefined,
      {
        fontSizePt: 12,
        fonts: { complexScript: 'Times New Roman' },
        themeFonts: {
          ascii: 'ＭＳ 明朝',
          highAnsi: 'ＭＳ 明朝',
          eastAsia: 'ＭＳ 明朝',
        },
        themeFontPresence: {
          ascii: true,
          highAnsi: true,
          eastAsia: true,
          complexScript: false,
        },
        weight: 400,
        style: 'normal',
        complexScript: false,
        fontHint: 'eastAsia',
        eastAsiaLanguage: 'ja-jp',
      },
    )).toBeCloseTo(15.6, 5);
  });

  it('uses the same arbitrary resource metric for pagination and paint', async () => {
    const family = 'Arbitrary Embedded EA';
    const observedParagraph = (text: string, shading?: string): DocParagraph => ({
      ...para(text),
      runs: text === ''
        ? []
        : [{
            type: 'text',
            ...textRun(text),
            fontSize: 12,
            fontFamily: 'Century',
            fontFamilyEastAsia: family,
          } as DocParagraph['runs'][number]],
      defaultFontSize: 12,
      defaultFontFamily: 'Century',
      defaultFontFamilyEastAsia: family,
      shading,
      paragraphMarkFontFacts: {
        fontSize: 12,
        fontFamily: 'Century',
        fontFamilyEastAsia: family,
        fontHint: 'eastAsia',
        langEastAsia: 'ja-jp',
      },
    } as unknown as DocParagraph);
    const model = docOf([
      observedParagraph('甲'),
      observedParagraph('', 'c0ffee'),
      observedParagraph('乙'),
    ]);
    model.section.pageWidth = 200;
    model.section.pageHeight = 39;
    model.fontFamilyClasses = {};
    const measurement = makeResolvedMetricCanvas();
    const resolved = {
      'arbitrary embedded ea': {
        family,
        eastAsianLineHeightRatio: 1.3,
      },
    };
    const services = createLayoutServices(model, {
      measureContext: measurement.canvas.getContext('2d') as CanvasRenderingContext2D,
      fontMetrics: resolved,
    });

    // The resolved 1.3-em East-Asian mark plus the two 12pt visible lines is
    // 39.6pt, so the final line starts on page two. Omitting the resource metric
    // would incorrectly fit all three lines into the 39pt first-page band.
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);

    const painted = [] as ReturnType<typeof makeResolvedMetricCanvas>[];
    for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex++) {
      const recording = makeResolvedMetricCanvas();
      painted.push(recording);
      await renderDocumentToCanvas(model, recording.canvas, pageIndex, {
        dpr: 1,
        width: 200,
        layoutServices: services,
      });
    }
    expect(painted[0].textCalls.map((call) => call.text)).toContain('甲');
    expect(painted[0].textCalls.map((call) => call.text)).not.toContain('乙');
    expect(painted[1].textCalls.map((call) => call.text)).toContain('乙');
  });

  it('counts an untabled 20pt East Asian mark as two cells on a 20pt grid', () => {
    const p = {
      ...para(''),
      defaultFontSize: 20,
      defaultFontFamilyEastAsia: 'ＭＳ 明朝',
    } as DocParagraph;
    let font = '';
    const ctx = {
      get font() { return font; },
      set font(v: string) { font = v; },
      measureText: () => ({
        width: 0,
        fontBoundingBoxAscent: 18,
        fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 18,
        actualBoundingBoxDescent: 2,
      } as TextMetrics),
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(
      p,
      1,
      { type: 'lines', linePitchPt: 20 },
      false,
      true,
      ctx,
      {},
    )).toBe(40);
  });

  it('uses an eastAsia paragraph-mark hint for face selection without forcing Latin text into two grid cells', () => {
    const p = {
      ...para(''),
      defaultFontSize: 16,
      defaultFontFamily: 'Georgia',
      defaultFontFamilyEastAsia: 'ＭＳ 明朝',
    } as DocParagraph;
    const ctx = {
      font: '',
      measureText: () => ({
        width: 0,
        fontBoundingBoxAscent: 14,
        fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 14,
        actualBoundingBoxDescent: 2,
      } as TextMetrics),
    } as unknown as CanvasRenderingContext2D;

    expect(paragraphMarkLineHeight(
      p,
      1,
      { type: 'lines', linePitchPt: 18 },
      false,
      false,
      ctx,
      {},
      null,
      {},
      undefined,
      {
        fontSizePt: 16,
        fonts: {
          ascii: 'Georgia',
          highAnsi: 'Georgia',
          eastAsia: 'ＭＳ 明朝',
          complexScript: 'Georgia',
        },
        weight: 400,
        style: 'normal',
        complexScript: false,
        fontHint: 'eastAsia',
      },
    )).toBe(18);
  });

  it('uses the resolved local alias and line ratio for the mark advance and baseline split', () => {
    const p = {
      ...para(''),
      defaultFontSize: 10,
      defaultFontFamily: 'Authored Family',
    } as DocParagraph;
    let font = '';
    const ctx = {
      get font() { return font; },
      set font(v: string) { font = v; },
      measureText: () => {
        const local = font.includes('__ooxml_local_exact');
        return {
          width: 0,
          fontBoundingBoxAscent: local ? 12 : 8,
          fontBoundingBoxDescent: local ? 3 : 2,
          actualBoundingBoxAscent: local ? 12 : 8,
          actualBoundingBoxDescent: local ? 3 : 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const resolved = {
      'authored family': { family: '__ooxml_local_exact', lineHeightRatio: 1.5 },
    };

    expect(paragraphMarkLineHeight(
      p, 1, { type: null, linePitchPt: null }, false, false, ctx, {}, null, resolved,
    )).toBe(15);
    expect(paragraphMarkBelowBaselinePt(
      p, { type: null, linePitchPt: null }, false, false, ctx, {}, null, resolved,
    )).toBe(3);
    expect(font).toBe('');
  });

  it('resolves an empty mark probe through the same registered substitute route in main and worker services', () => {
    const makeContext = () => {
      let font = '';
      const measured: Array<{ text: string; font: string }> = [];
      const ctx = {
        get font() { return font; },
        set font(value: string) { font = value; },
        letterSpacing: '0px',
        fontKerning: 'auto' as CanvasFontKerning,
        measureText(text: string) {
          measured.push({ text, font });
          const substitute = font.includes('Carlito');
          return {
            width: 0,
            fontBoundingBoxAscent: substitute ? 12 : 8,
            fontBoundingBoxDescent: substitute ? 3 : 2,
            actualBoundingBoxAscent: substitute ? 12 : 8,
            actualBoundingBoxDescent: substitute ? 3 : 2,
          } as TextMetrics;
        },
      } as unknown as CanvasRenderingContext2D;
      return { ctx, measured };
    };
    const p = {
      ...para(''),
      defaultFontSize: 10,
      defaultFontFamily: 'Legacy Mark',
      paragraphMarkFontFacts: {
        fontFamily: 'Legacy Mark',
        fontSlots: {
          direct: { ascii: 'Legacy Direct' },
          theme: { ascii: 'Calibri' },
          themePresent: { ascii: true, highAnsi: false, eastAsia: false, complexScript: false },
        },
      },
    } as unknown as DocParagraph;
    const model = {
      ...docOf([]),
      majorFont: 'Calibri',
    };
    const mainContext = makeContext();
    const workerContext = makeContext();
    const options = (ctx: CanvasRenderingContext2D) => ({
      useGoogleFonts: true,
      googleFaces: [{
        family: 'Carlito', weight: '400', style: 'normal', status: 'loaded',
      } as FontFace],
      measureContext: ctx,
    });
    const main = createLayoutServices(model, options(mainContext.ctx));
    const worker = createLayoutServices(model, options(workerContext.ctx));
    const measureMark = (
      ctx: CanvasRenderingContext2D,
      service: typeof main.text,
    ) => paragraphMarkLineHeight(
      p, 1, { type: null, linePitchPt: null }, false, false,
      ctx, {}, null, service.localMetrics, service, paragraphMarkShapeInput(p),
    );

    expect(main.text.fingerprint).toBe(worker.text.fingerprint);
    expect(measureMark(mainContext.ctx, main.text)).toBe(15);
    expect(measureMark(workerContext.ctx, worker.text)).toBe(15);
    const expectedRoute = main.text.shape({
      text: 'x', fontSizePt: 10,
      fonts: { ascii: 'Legacy Direct' },
      themeFonts: { ascii: 'Calibri' },
      themeFontPresence: { ascii: true },
    }).spans[0]!.fontRoute;
    const expectedFont = canvasFontString(expectedRoute, 10, 400, 'normal');
    expect(mainContext.measured.filter(({ text }) => text === 'x').map(({ font }) => font))
      .toContain(expectedFont);
    expect(workerContext.measured.filter(({ text }) => text === 'x').map(({ font }) => font))
      .toContain(expectedFont);
  });

  it('uses the same resolved local mark metrics for pagination and paint', async () => {
    const authoredFamily = 'Authored Family';
    const localPara = (text: string, shading?: string): DocParagraph => ({
      ...para(text),
      runs: text === ''
        ? []
        : [{ type: 'text', ...textRun(text), fontSize: 10, fontFamily: authoredFamily } as DocParagraph['runs'][number]],
      defaultFontSize: 10,
      defaultFontFamily: authoredFamily,
      shading,
    } as DocParagraph);
    const model = docOf([
      localPara('A'),
      localPara('', 'c0ffee'),
      localPara('B'),
    ]);
    model.section.pageWidth = 200;
    model.section.pageHeight = 40;
    model.fontFamilyClasses = {};
    const resolved = {
      'authored family': { family: '__ooxml_local_exact', lineHeightRatio: 1.5 },
    };
    const globalWithCanvas = globalThis as unknown as { OffscreenCanvas?: unknown };
    const previousOffscreenCanvas = globalWithCanvas.OffscreenCanvas;
    globalWithCanvas.OffscreenCanvas = class {
      getContext() { return makeResolvedMetricCanvas().canvas.getContext('2d'); }
    };
    const services = createLayoutServices(model, { localMetrics: resolved });

    try {
      // Three 15pt local-face line boxes do not fit in the 40pt content band.
      // Before the fix the empty mark used the authored fallback's 10pt box, so
      // all three paragraphs were incorrectly packed onto one page.
      const layout = layoutDocument(model, services, { currentDateMs: 0 });
      expect(layout.pages).toHaveLength(2);
      expect(layout.pages[0].layers.body).toHaveLength(2);
      expect(layout.pages[1].layers.body).toHaveLength(1);

      const painted = [] as ReturnType<typeof makeResolvedMetricCanvas>[];
      for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex++) {
        const recording = makeResolvedMetricCanvas();
        painted.push(recording);
        await renderDocumentToCanvas(model, recording.canvas, pageIndex, {
          dpr: 1,
          width: 200,
          layoutServices: services,
        });
      }

      expect(painted[0].textCalls.map((call) => call.text)).toContain('A');
      expect(painted[0].textCalls.map((call) => call.text)).not.toContain('B');
      expect(painted[1].textCalls.map((call) => call.text)).toContain('B');
      expect(painted[0].rectCalls.some((call) => Math.abs(call.height - 15) < 1e-6)).toBe(true);
    } finally {
      if (previousOffscreenCanvas === undefined) delete globalWithCanvas.OffscreenCanvas;
      else globalWithCanvas.OffscreenCanvas = previousOffscreenCanvas;
    }
  });
});
