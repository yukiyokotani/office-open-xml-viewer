import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import {
  acquireAndPaintShapeTextBox,
  shapeAcquisitionState,
} from './retained-shape-textbox.test-support.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
  NumberingInfo,
  ShapeRun,
  ShapeText,
} from './types';

// ECMA-376 §17.3.2.26 (rFonts ascii/eastAsia axes) + §17.9.6 (numbering level
// rPr). End-to-end renderer verification of the sample-10 heading "1 原稿の体裁":
// the auto-number "1" (Latin) must draw with the ASCII font (Times / serif) and
// the CJK title "原稿の体裁" with the EASTASIA font (MS Gothic / sans). Before the
// fix the renderer drew the number with a hardcoded sans-serif and routed the CJK
// glyphs through the ascii font's serif-mincho fallback — the exact inverse.

// fontTable §17.8.3.10 classes drive serif vs sans deterministically: 'roman' →
// serif tail, 'swiss' → sans tail. Times = serif (roman), MS Gothic = sans
// (swiss), MS Mincho = serif (roman) for the no-regression case.
const FONT_CLASSES: Record<string, string> = {
  'Times New Roman': 'roman',
  'ＭＳ ゴシック': 'swiss',
  'ＭＳ 明朝': 'roman',
};

/** Recording 2D context. Records each fillText WITH the active font, so the
 *  numbering marker (drawn via fillText, NOT reported through onTextRun) can be
 *  inspected. Glyph advance = charCount × fontPx; font box 0.8/0.2 em. */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number; font: string }[];
} {
  let font = '16px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
  const fillTextCalls: { text: string; x: number; font: string }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
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
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, _y: number) { fillTextCalls.push({ text, x, font }); },
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
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

/** A run whose ascii face is `fontFamily` and CJK (eastAsia) face is
 *  `fontFamilyEastAsia` — the resolved sample-10 heading run. */
function run(text: string, fontFamily: string, fontFamilyEastAsia: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 16, color: null, fontFamily, fontFamilyEastAsia,
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  };
}

function numbering(): NumberingInfo {
  // suff:'space' so the marker is measured/drawn AND the body abuts it (no tab
  // dependency). The decimal marker "1" → ascii axis. Marker fonts resolved by
  // the parser (ascii=Times, eastAsia=MS Gothic).
  return {
    numId: 1, level: 0, format: 'decimal', text: '1',
    indentLeft: 0, tab: 18, suff: 'space',
    fontFamily: 'Times New Roman',
    fontFamilyEastAsia: 'ＭＳ ゴシック',
  };
}

function headingDoc(num: NumberingInfo | null): DocxDocumentModel {
  const p: DocParagraph = {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: num, tabStops: [],
    // ascii=Times (serif), eastAsia=MS Gothic (sans) — the resolved heading run.
    runs: [{ type: 'text', ...run('原稿の体裁', 'Times New Roman', 'ＭＳ ゴシック') } as DocParagraph['runs'][number]],
    defaultFontSize: 16, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  };
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    // fontTable §17.8.3.10 classes — the authoritative serif/sans source the
    // renderer reads from the MODEL (doc.fontFamilyClasses), not the options.
    fontFamilyClasses: FONT_CLASSES,
  } as unknown as DocxDocumentModel;
}

async function render(num: NumberingInfo | null) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  const model = headingDoc(num);
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1,
    width: 400, // scale = 1 (px per pt)
    onTextRun: (r) => runs.push(r),
    layoutServices: createLayoutServices(model, {
      localMetrics: testFontSnapshot([
        { family: 'Times New Roman' },
        { family: 'ＭＳ ゴシック' },
        { family: 'ＭＳ 明朝' },
      ]),
      measureContext: canvas.getContext('2d'),
    }),
  });
  return { runs, fillTextCalls };
}

// The requested family is always the FIRST quoted token in the CSS font string
// (`<style> <weight> <size>px "<family>", <fallbacks>`).
const headFamily = (font: string) => /"([^"]+)"/.exec(font)?.[1] ?? font;

describe('numbering marker + body eastAsia font routing (§17.3.2.26 / §17.9.6)', () => {
  it('routes the CJK title to the eastAsia (MS Gothic / sans) face', async () => {
    const { runs } = await render(numbering());
    const title = runs.find((r) => r.text === '原稿の体裁');
    expect(title).toBeDefined();
    // The whole title is one CJK segment → drawn with the eastAsia family.
    expect(headFamily(title!.font)).toBe('ＭＳ ゴシック');
    // …and that family resolves to a SANS tail (fontTable 'swiss').
    expect(title!.font.endsWith('sans-serif')).toBe(true);
  });

  it('draws the auto-number "1" with the ascii (Times / serif) face, not sans', async () => {
    const { fillTextCalls } = await render(numbering());
    const marker = fillTextCalls.find((c) => c.text === '1');
    expect(marker).toBeDefined();
    // The marker is a Latin digit → ascii axis → Times (serif), NOT the old
    // hardcoded "sans-serif".
    expect(headFamily(marker!.font)).toBe('Times New Roman');
    expect(marker!.font.endsWith('serif')).toBe(true); // serif tail
    expect(marker!.font.endsWith('sans-serif')).toBe(false); // NOT the old sans hardcode
  });

  it('keeps the title and marker on DIFFERENT faces (the bug was the inverse)', async () => {
    const { runs, fillTextCalls } = await render(numbering());
    const title = runs.find((r) => r.text === '原稿の体裁')!;
    const marker = fillTextCalls.find((c) => c.text === '1')!;
    expect(headFamily(marker.font)).toBe('Times New Roman'); // serif number
    expect(headFamily(title.font)).toBe('ＭＳ ゴシック'); // sans title
    expect(headFamily(marker.font)).not.toBe(headFamily(title.font));
  });

  it('shapes a mixed 第1章 marker per scalar from retained four-slot theme facts', async () => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const num = {
      ...numbering(),
      text: '第1章',
      fontFamily: 'Legacy ASCII',
      fontFamilyEastAsia: 'Legacy EA',
      fontFacts: {
        fontFamily: 'Legacy ASCII',
        fontFamilyHighAnsi: 'Legacy HANSI',
        fontFamilyEastAsia: 'Legacy EA',
        fontSlots: {
          direct: {
            ascii: 'Direct ASCII', highAnsi: 'Direct HANSI',
            eastAsia: 'Direct EA', complexScript: 'Direct CS',
          },
          theme: { ascii: 'Theme ASCII', eastAsia: 'Theme EA' },
          themePresent: { ascii: true, highAnsi: false, eastAsia: true, complexScript: false },
        },
      },
    } as unknown as NumberingInfo;
    const model = headingDoc(num);
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1,
      width: 400,
      layoutServices: createLayoutServices(model, {
        localMetrics: testFontSnapshot([
          { family: 'Theme ASCII' },
          { family: 'Theme EA' },
        ]),
        measureContext: canvas.getContext('2d'),
      }),
    });

    const marker = fillTextCalls.filter((call) => ['第', '1', '章'].includes(call.text));
    expect(marker.map((call) => [call.text, headFamily(call.font)])).toEqual([
      ['第', 'Theme EA'],
      ['1', 'Theme ASCII'],
      ['章', 'Theme EA'],
    ]);
    expect(fillTextCalls.some((call) => call.text === '第1章')).toBe(false);
  });

  it('routes U+2022 through highAnsi theme presence and the registered substitute face', async () => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const num = {
      ...numbering(),
      format: 'bullet',
      text: '•',
      fontFamily: 'Legacy ASCII',
      fontFamilyEastAsia: 'Legacy EA',
      fontFacts: {
        fontFamily: 'Legacy ASCII',
        fontFamilyHighAnsi: 'Direct HANSI',
        fontFamilyEastAsia: 'Legacy EA',
        fontSlots: {
          direct: { ascii: 'Legacy ASCII', highAnsi: 'Direct HANSI', eastAsia: 'Legacy EA' },
          theme: { highAnsi: 'Calibri' },
          themePresent: { ascii: false, highAnsi: true, eastAsia: false, complexScript: false },
        },
      },
    } as unknown as NumberingInfo;
    const model = { ...headingDoc(num), majorFont: 'Calibri' };
    const carlito = {
      family: 'Carlito', weight: '400', style: 'normal', status: 'loaded',
    } as FontFace;
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1,
      width: 400,
      layoutServices: createLayoutServices(model, {
        providerRoutes: {
          calibri: { family: carlito.family.replaceAll('"', ''), source: 'substitute' },
        },
        measureContext: canvas.getContext('2d'),
      }),
    });

    const marker = fillTextCalls.find((call) => call.text === '•');
    expect(marker).toBeDefined();
    expect(headFamily(marker!.font)).toBe('Carlito');
  });

  it('no-regression: a CJK run with NO eastAsia axis falls back to the ascii face', async () => {
    // Field-run / legacy single-axis output: fontFamilyEastAsia absent → the CJK
    // glyphs keep the ascii family, exactly as before this change.
    const { canvas } = makeRecordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    const p: DocParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [{
        type: 'text',
        ...run('原稿', 'ＭＳ 明朝', '' as unknown as string),
      } as DocParagraph['runs'][number]],
      defaultFontSize: 16, defaultFontFamily: 'ＭＳ 明朝', widowControl: false,
    };
    // Drop the eastAsia axis entirely (simulate older parser output).
    delete (p.runs[0] as Partial<DocxTextRun>).fontFamilyEastAsia;
    const model = {
      section: {
        pageWidth: 400, pageHeight: 400, marginTop: 0, marginRight: 0,
        marginBottom: 0, marginLeft: 0, headerDistance: 0, footerDistance: 0,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [{ type: 'paragraph', ...p }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: FONT_CLASSES,
    } as unknown as DocxDocumentModel;
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 400, onTextRun: (r) => runs.push(r),
      layoutServices: createLayoutServices(model, {
        localMetrics: testFontSnapshot([{ family: 'Times New Roman' }, { family: 'ＭＳ 明朝' }]), measureContext: canvas.getContext('2d'),
      }),
    });
    const seg = runs.find((r) => r.text === '原稿');
    expect(seg).toBeDefined();
    expect(headFamily(seg!.font)).toBe('ＭＳ 明朝'); // ascii fallback, unchanged
  });

  it('no-regression: the common mincho case (eastAsia=明朝) stays serif', async () => {
    // eastAsia = MS Mincho (serif), same CLASS as the ascii fallback → output
    // unchanged from before the per-script split.
    const { canvas } = makeRecordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    const p: DocParagraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [{
        type: 'text', ...run('本文', 'Times New Roman', 'ＭＳ 明朝'),
      } as DocParagraph['runs'][number]],
      defaultFontSize: 16, defaultFontFamily: 'ＭＳ 明朝', widowControl: false,
    };
    const model = {
      section: {
        pageWidth: 400, pageHeight: 400, marginTop: 0, marginRight: 0,
        marginBottom: 0, marginLeft: 0, headerDistance: 0, footerDistance: 0,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [{ type: 'paragraph', ...p }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: FONT_CLASSES,
    } as unknown as DocxDocumentModel;
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 400, onTextRun: (r) => runs.push(r),
      layoutServices: createLayoutServices(model, {
        localMetrics: testFontSnapshot([{ family: 'Times New Roman' }, { family: 'ＭＳ 明朝' }]), measureContext: canvas.getContext('2d'),
      }),
    });
    const seg = runs.find((r) => r.text === '本文')!;
    expect(headFamily(seg.font)).toBe('ＭＳ 明朝'); // eastAsia mincho
    expect(seg.font).toMatch(/serif/); // still serif — no visible change
  });

  const textBoxShape = (num: NumberingInfo): ShapeRun => ({
    type: 'shape', presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: [{
      text: 'item', fontSizePt: 10, alignment: 'left',
      runs: [{ text: 'item', fontSizePt: 10 }],
      numbering: num,
    } as unknown as ShapeText],
  }) as unknown as ShapeRun;

  it('text boxes retain the same per-scalar 第1章 routes for width and paint', () => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d')!;
    const num = {
      ...numbering(),
      text: '第1章',
      fontFacts: {
        fontFamily: 'Legacy ASCII',
        fontFamilyHighAnsi: 'Legacy HANSI',
        fontFamilyEastAsia: 'Legacy EA',
        fontSlots: {
          direct: { ascii: 'Direct ASCII', highAnsi: 'Direct HANSI', eastAsia: 'Direct EA' },
          theme: { ascii: 'Theme ASCII', eastAsia: 'Theme EA' },
          themePresent: { ascii: true, highAnsi: false, eastAsia: true, complexScript: false },
        },
      },
    } as unknown as NumberingInfo;
    const doc = headingDoc(num);
    const services = createLayoutServices(doc, {
      localMetrics: testFontSnapshot([{ family: 'Theme ASCII' }, { family: 'Theme EA' }]),
      measureContext: ctx,
    });
    const state = shapeAcquisitionState(ctx, {});
    state.layoutServices = services;

    acquireAndPaintShapeTextBox(textBoxShape(num), 0, 0, 200, 100, ctx, 1, {}, new Map(), state);

    const marker = fillTextCalls.filter((call) => ['第', '1', '章'].includes(call.text));
    expect(marker.map((call) => [call.text, headFamily(call.font)])).toEqual([
      ['第', 'Theme EA'],
      ['1', 'Theme ASCII'],
      ['章', 'Theme EA'],
    ]);
    expect(marker[1]!.x).toBeGreaterThan(marker[0]!.x);
    expect(marker[2]!.x).toBeGreaterThan(marker[1]!.x);
    expect(fillTextCalls.some((call) => call.text === '第1章')).toBe(false);
  });

  it('text boxes route U+2022 through highAnsi theme substitution for measure and paint', () => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d')!;
    const num = {
      ...numbering(),
      format: 'bullet',
      text: '•',
      fontFacts: {
        fontFamily: 'Legacy ASCII',
        fontFamilyHighAnsi: 'Direct HANSI',
        fontSlots: {
          direct: { ascii: 'Legacy ASCII', highAnsi: 'Direct HANSI' },
          theme: { highAnsi: 'Calibri' },
          themePresent: { ascii: false, highAnsi: true, eastAsia: false, complexScript: false },
        },
      },
    } as unknown as NumberingInfo;
    const doc = { ...headingDoc(num), majorFont: 'Calibri' };
    const services = createLayoutServices(doc, {
      providerRoutes: {
        calibri: { family: 'Carlito', source: 'substitute' },
      },
      measureContext: ctx,
    });
    const state = shapeAcquisitionState(ctx, {});
    state.layoutServices = services;

    acquireAndPaintShapeTextBox(textBoxShape(num), 0, 0, 200, 100, ctx, 1, {}, new Map(), state);

    const marker = fillTextCalls.find((call) => call.text === '•');
    expect(marker).toBeDefined();
    expect(headFamily(marker!.font)).toBe('Carlito');
  });
});
