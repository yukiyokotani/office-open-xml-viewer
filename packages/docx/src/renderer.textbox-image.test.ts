import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { preloadImages } from './test-support/renderer-internals.test-support.js';
import {
  acquireAndPaintShapeTextBox,
  acquireShapeTextBoxForTest,
  type ShapeAcquisitionTestState,
} from './retained-shape-textbox.test-support.js';
import type { DocxDocumentModel, ShapeRun, ShapeText, ShapeTextRun } from './types';
import { canvasFontString } from '@silurus/ooxml-core';

/**
 * Inline images living INSIDE a DOCX text box (`<wps:txbx>`) ride on the
 * shape's `textBlocks[i].imagePath` rather than a top-level `image` run. Two
 * things must hold end-to-end:
 *   1. `collectImagePairs` (exercised through `preloadImages`, exactly as
 *      renderer.image.test.ts drives it) must surface those textbox images so
 *      their bytes reach the decode pipeline (WMF decoder included).
 *   2. retained text-box paint must draw the decoded bitmap fitted to the inner
 *      width, and must NOT throw / draw when the bitmap is missing.
 */

// Recording mock canvas context (extends pagination.test.ts's makeCtx with a
// drawImage spy and a measureText stub: glyph advance = charCount × fontPx).
interface DrawImageCall {
  bmp: unknown;
  x: number;
  y: number;
  w: number;
  h: number;
}
function makeRecordingCtx(): {
  ctx: CanvasRenderingContext2D;
  drawImageCalls: DrawImageCall[];
  fillTextCalls: { text: string; x: number; y: number; font: string; fillStyle: string }[];
  fontCalls: string[];
} {
  let font = '10px serif';
  let fillStyle = '#000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const drawImageCalls: DrawImageCall[] = [];
  const fillTextCalls: { text: string; x: number; y: number; font: string; fillStyle: string }[] = [];
  const fontCalls: string[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; fontCalls.push(v); },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
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
    save() {}, restore() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {},
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, font, fillStyle });
    },
    strokeText() {},
    drawImage(bmp: unknown, x: number, y: number, w: number, h: number) {
      drawImageCalls.push({ bmp, x, y, w, h });
    },
    strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, drawImageCalls, fillTextCalls, fontCalls };
}

/** Build a shape (ShapeRun) carrying a single rich-text block whose `runs`
 *  describe per-run formatting. Insets default to 0 so layout math is easy; the
 *  box width is supplied to retained text-box acquisition at each call. */
function richTextbox(runs: ShapeTextRun[], alignment = 'left'): ShapeRun {
  const block: ShapeText = {
    text: runs.map((r) => r.text).join(''),
    // Single block-level fields come from the first run (parser backward compat).
    fontSizePt: runs[0]?.fontSizePt ?? 10,
    bold: runs[0]?.bold,
    alignment,
    runs,
  };
  return {
    type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

/** A text box (ShapeRun) whose first text block is an inline image and whose
 *  second is a caption — the sample-10 Fig.1 layout. Insets default to 0 so the
 *  fit math is easy to assert. */
function textboxWithImage(overrides: Partial<ShapeText> = {}): ShapeRun {
  const imageBlock: ShapeText = {
    text: '',
    fontSizePt: 10,
    alignment: 'center',
    imagePath: 'word/media/image1.emf',
    mimeType: 'image/x-wmf',
    imageWidthPt: 100,
    imageHeightPt: 50,
    ...overrides,
  };
  const captionBlock: ShapeText = {
    text: 'Fig. 1: A sample figure.',
    fontSizePt: 10,
    alignment: 'center',
  };
  return {
    type: 'shape',
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: null,
    stroke: null,
    textBlocks: [imageBlock, captionBlock],
    textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

describe('textbox inline images — collection', () => {
  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_src: unknown) => ({ width: 4, height: 2, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('collectImagePairs (via preloadImages) surfaces a textbox shape image', async () => {
    const fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
    );
    // A shape run (not an image run) carrying an inline image on its text block.
    const doc = {
      body: [
        { type: 'paragraph', runs: [textboxWithImage()] },
      ],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel;

    const map = await preloadImages(doc, fetchImage);

    // The textbox image must have been fetched + decoded and keyed by its path.
    expect(fetchImage).toHaveBeenCalledWith('word/media/image1.emf', 'image/x-wmf');
    expect(map.has('word/media/image1.emf')).toBe(true);
  });
});

describe('textbox inline images — rendering', () => {
  // A bitmap-like stub: drawImage only needs width/height-bearing source.
  const fakeBmp = { width: 200, height: 100, close: () => {} } as unknown as ImageBitmap;

  it('draws the bitmap fitted to the inner width', () => {
    const { ctx, drawImageCalls } = makeRecordingCtx();
    const shape = textboxWithImage(); // natural 100×50 pt
    const images = new Map<string, DecodedImage>([['word/media/image1.emf', fakeBmp]]);

    // Box: 80pt wide so the 100pt-wide image must scale DOWN to innerW=80,
    // height 50 × (80/100) = 40. scale=1 → px == pt.
    const scale = 1;
    acquireAndPaintShapeTextBox(shape, /*x*/ 0, /*y*/ 0, /*w*/ 80, /*h*/ 200, ctx, scale, {}, images);

    expect(drawImageCalls).toHaveLength(1);
    const call = drawImageCalls[0];
    expect(call.bmp).toBe(fakeBmp);
    expect(call.w).toBeCloseTo(80, 5);   // scaled to innerW
    expect(call.h).toBeCloseTo(40, 5);   // aspect preserved
    // innerW == fitW ⇒ centered draw sits flush at x=0.
    expect(call.x).toBeCloseTo(0, 5);
    // Image is the first block ⇒ drawn at the top of the inner box (anchor 't').
    expect(call.y).toBeCloseTo(0, 5);
  });

  it('draws an image smaller than innerW at natural size, centered', () => {
    const { ctx, drawImageCalls } = makeRecordingCtx();
    const shape = textboxWithImage(); // natural 100×50 pt
    const images = new Map<string, DecodedImage>([['word/media/image1.emf', fakeBmp]]);

    // innerW=200 > natural 100 ⇒ keep natural 100×50, centered ⇒ x=(200-100)/2=50.
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 200, ctx, 1, {}, images);

    expect(drawImageCalls).toHaveLength(1);
    expect(drawImageCalls[0].w).toBeCloseTo(100, 5);
    expect(drawImageCalls[0].h).toBeCloseTo(50, 5);
    expect(drawImageCalls[0].x).toBeCloseTo(50, 5);
  });

  it('with a missing bitmap draws nothing and does not throw', () => {
    const { ctx, drawImageCalls, fillTextCalls } = makeRecordingCtx();
    const shape = textboxWithImage();
    const images = new Map<string, DecodedImage>(); // bitmap NOT present

    expect(() => acquireAndPaintShapeTextBox(shape, 0, 0, 80, 200, ctx, 1, {}, images)).not.toThrow();
    // No image drawn …
    expect(drawImageCalls).toHaveLength(0);
    // … but the caption text block still renders (height was still reserved).
    // The 80px-wide box wraps the caption across several lines, so assert the
    // concatenated wrapped lines reproduce the caption (ignoring the spaces
    // dropped at wrap points) rather than a single full-string fillText.
    const captionInk = fillTextCalls.map((c) => c.text).join('').replace(/\s/g, '');
    expect(captionInk).toContain('Fig.1:Asamplefigure.');
  });

  it('wraps a long text block to multiple lines within the inner width', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    // Single text block (no image). At fontSizePt 10 / scale 1 the mock advance
    // is 10px per char, so a 100px-wide box holds ~10 chars per line.
    const shape = {
      type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
      textBlocks: [{ text: 'aaaa bbbb cccc dddd eeee', fontSizePt: 10, alignment: 'left' }],
      textAnchor: 't', textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    } as unknown as ShapeRun;

    acquireAndPaintShapeTextBox(shape, 0, 0, 100, 400, ctx, 1, {}, new Map());

    // The 24-char run cannot fit on one 100px line ⇒ it wraps.
    expect(fillTextCalls.length).toBeGreaterThan(1);
    // No drawn line exceeds the inner width (10px/char).
    for (const c of fillTextCalls) {
      expect(c.text.length * 10).toBeLessThanOrEqual(100 + 1e-6);
    }
    // All ink preserved (only wrap-point spaces dropped).
    const ink = fillTextCalls.map((c) => c.text).join('').replace(/\s/g, '');
    expect(ink).toBe('aaaabbbbccccddddeeee');
  });

  it('spAutoFit text boxes use serialized bodyPr horizontal insets once', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const shape = {
      type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
      textBlocks: [{ text: 'fit', fontSizePt: 10, alignment: 'left' }],
      textAnchor: 't', textInsetL: 7.2, textInsetT: 0, textInsetR: 7.2, textInsetB: 0,
      textAutofit: 'sp',
    } as unknown as ShapeRun;

    acquireAndPaintShapeTextBox(shape, 10, 0, 200, 400, ctx, 1, {}, new Map());

    expect(fillTextCalls[0]?.x).toBeCloseTo(17.2, 5);
  });

  it('retained spAutoFit acquisition returns content height plus bodyPr insets', () => {
    const { ctx } = makeRecordingCtx();
    const shape = {
      type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
      textBlocks: [{ text: 'fit', fontSizePt: 10, alignment: 'left' }],
      textAnchor: 't', textInsetL: 2, textInsetT: 3, textInsetR: 2, textInsetB: 4,
      textAutofit: 'sp',
    } as unknown as ShapeRun;

    const layout = acquireShapeTextBoxForTest(shape, 0, 0, 200, 400, ctx, 1);

    expect(layout?.flowBounds.heightPt).toBeCloseTo(17, 5);
  });
});

describe('textbox rich text — per-run formatting', () => {
  it('uses one byte-identical service route for auto-fit, line metrics, and paint', () => {
    const { ctx, fontCalls } = makeRecordingCtx();
    const doc = {
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: { 'Roman Face': 'roman' },
    } as DocxDocumentModel;
    const services = createLayoutServices(doc, { measureContext: ctx });
    const route = services.text.resolve({
      fonts: { ascii: 'Roman Face', highAnsi: 'Roman Face' },
      slot: 'ascii',
    }).route;
    const expected = canvasFontString(route, 10, 400, 'normal');
    const state = {
      pageIndex: 0,
      totalPages: 1,
      layoutServices: services,
      resolvedLocalFonts: services.text.localMetrics,
    } satisfies ShapeAcquisitionTestState;
    const shape = richTextbox([{ text: 'route', fontSizePt: 10, fontFamily: 'Roman Face' }]);

    fontCalls.length = 0;
    acquireShapeTextBoxForTest(
      shape, 0, 0, 200, 100, ctx, 1, doc.fontFamilyClasses, state,
    );
    // Native line-metric probes use larger sizes; the authored 10px route must
    // still be identical for text acquisition and paint.
    const authoredCalls = fontCalls.filter((font) => font.includes('Roman Face') && font.includes(' 10px '));
    expect(authoredCalls).toContain(expected);
    expect(authoredCalls.every((font) => font === expected)).toBe(true);

    fontCalls.length = 0;
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, doc.fontFamilyClasses, new Map(), state);
    const paintedCalls = fontCalls.filter((font) => font.includes('Roman Face') && font.includes(' 10px '));
    expect(paintedCalls).toContain(expected);
    expect(paintedCalls.every((font) => font === expected)).toBe(true);
  });

  /** Tokens belonging to a substring, with the font each was drawn with. */
  function tokensFor(
    calls: { text: string; font: string }[],
    needle: string,
  ): { text: string; font: string }[] {
    // The renderer draws one fillText per token (Latin word incl. trailing
    // space, or a single CJK char). A token "belongs" to `needle` when it is a
    // non-empty substring of it.
    const probe = needle.replace(/\s/g, '');
    return calls.filter((c) => {
      const t = c.text.replace(/\s/g, '');
      return t.length > 0 && probe.includes(t);
    });
  }
  const isBoldFont = (f: string) => /\bbold\b|\b700\b/i.test(f);

  it('draws the bold label run bold and the non-bold body run non-bold', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    // sample-10 Abstract: "Abstract－ " bold, body NOT bold. Wide box ⇒ 1 line.
    const shape = richTextbox([
      { text: 'Abstract－ ', fontSizePt: 10, bold: true },
      { text: 'This document.', fontSizePt: 10, bold: false },
    ]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 2000, 400, ctx, 1, {}, new Map());

    // All ink preserved (wrap-point spaces aside).
    const ink = fillTextCalls.map((c) => c.text).join('').replace(/\s/g, '');
    expect(ink).toBe('Abstract－Thisdocument.');

    // The "Abstract" tokens are bold; the body tokens are not.
    const labelToks = tokensFor(fillTextCalls, 'Abstract－');
    const bodyToks = tokensFor(fillTextCalls, 'Thisdocument.');
    expect(labelToks.length).toBeGreaterThan(0);
    expect(bodyToks.length).toBeGreaterThan(0);
    for (const t of labelToks) expect(isBoldFont(t.font)).toBe(true);
    for (const t of bodyToks) expect(isBoldFont(t.font)).toBe(false);
  });

  it('draws ruby annotations above text-box base glyphs', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const shape = richTextbox([
      {
        text: '根号',
        fontSizePt: 22,
        fontFamily: 'MS Mincho',
        fontFamilyEastAsia: 'MS Mincho',
        ruby: { text: 'こんごう', fontSizePt: 5 },
      },
      { text: 'を含む式の', fontSizePt: 22 },
      { text: '加', fontSizePt: 22, ruby: { text: 'か', fontSizePt: 5 } },
      { text: '減', fontSizePt: 22, ruby: { text: 'げん', fontSizePt: 5 } },
    ]);

    acquireAndPaintShapeTextBox(shape, 0, 0, 500, 120, ctx, 1, {}, new Map());

    const root = fillTextCalls.find((c) => c.text === '根号');
    const rootRuby = fillTextCalls.find((c) => c.text === 'こ');
    const add = fillTextCalls.find((c) => c.text === '加');
    const addRuby = fillTextCalls.find((c) => c.text === 'か');
    const sub = fillTextCalls.find((c) => c.text === '減');
    const subRuby = fillTextCalls.find((c) => c.text === 'げ');

    expect(root).toBeDefined();
    expect(rootRuby).toBeDefined();
    expect(add).toBeDefined();
    expect(addRuby).toBeDefined();
    expect(sub).toBeDefined();
    expect(subRuby).toBeDefined();
    expect(rootRuby!.y).toBeLessThan(root!.y);
    expect(addRuby!.y).toBeLessThan(add!.y);
    expect(subRuby!.y).toBeLessThan(sub!.y);
    expect(fillTextCalls.map((c) => c.text).join('')).toBe('根号こんごうを含む式の加か減げん');
  });

  it('uses hpsRaise for the horizontal text-box ruby baseline and preserves zero', () => {
    const distance = (hpsRaisePt: number): number => {
      const { ctx, fillTextCalls } = makeRecordingCtx();
      const shape = richTextbox([{
        text: '漢', fontSizePt: 12, fontFamily: 'NotInMetrics',
        ruby: { text: 'かん', fontSizePt: 8, hpsRaisePt },
      }]);
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 2, {}, new Map());
      const base = fillTextCalls.find((call) => call.text === '漢');
      const ruby = fillTextCalls.find((call) => /[かん]/.test(call.text));
      expect(base, 'base glyph drawn').toBeDefined();
      expect(ruby, 'ruby glyph drawn').toBeDefined();
      return base!.y - ruby!.y;
    };

    expect(distance(14)).toBeCloseTo(28, 8);
    expect(distance(0)).toBe(0);
  });

  it('draws a numbered marker for text-box paragraphs', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const shape = richTextbox([{ text: '加法、減法の言葉に合った数式を生徒に考えさせる。', fontSizePt: 10, color: '000000' }]);
    shape.defaultTextColor = 'FFFFFF';
    shape.textBlocks![0].indentLeft = 36;
    shape.textBlocks![0].indentFirst = -36;
    shape.textBlocks![0].paragraphMarkColor = '000000';
    shape.textBlocks![0].numbering = {
      numId: 5,
      level: 0,
      format: 'bullet',
      text: '※',
      indentLeft: 36,
      tab: 36,
      suff: 'tab',
      jc: 'left',
      fontFamily: 'MS Gothic',
      fontFamilyEastAsia: 'MS Gothic',
    };

    acquireAndPaintShapeTextBox(shape, 0, 0, 260, 120, ctx, 1, {}, new Map());

    const marker = fillTextCalls.find((c) => c.text === '※');
    const body = fillTextCalls.find((c) => c.text.includes('加法'));
    expect(marker).toBeDefined();
    expect(body).toBeDefined();
    expect(marker!.fillStyle).toBe('#000000');
    expect(marker!.x).toBeLessThan(body!.x);
    expect(marker!.y).toBe(body!.y);
  });

  it('keeps each run its own font when mixed-run text wraps across lines', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    // 10px/char mock advance, 100px box ⇒ ~10 chars/line ⇒ forces a wrap.
    // First run bold, second run non-bold; the wrap falls inside the body run.
    const shape = richTextbox([
      { text: 'aaaa ', fontSizePt: 10, bold: true },
      { text: 'bbbb cccc dddd eeee', fontSizePt: 10, bold: false },
    ]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 100, 400, ctx, 1, {}, new Map());

    // It wrapped.
    expect(fillTextCalls.length).toBeGreaterThan(1);
    // Bold run token "aaaa" is bold; every body token is non-bold — regardless
    // of which line it landed on.
    const boldToks = tokensFor(fillTextCalls, 'aaaa');
    expect(boldToks.length).toBeGreaterThan(0);
    for (const t of boldToks) expect(isBoldFont(t.font)).toBe(true);
    for (const t of tokensFor(fillTextCalls, 'bbbbccccddddeeee')) {
      expect(isBoldFont(t.font)).toBe(false);
    }
    // All ink preserved.
    const ink = fillTextCalls.map((c) => c.text).join('').replace(/\s/g, '');
    expect(ink).toBe('aaaabbbbccccddddeeee');
  });

  it('applies each run its own color', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const shape = richTextbox([
      { text: 'red ', fontSizePt: 10, color: 'ff0000' },
      { text: 'plain.', fontSizePt: 10 },
    ]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 2000, 400, ctx, 1, {}, new Map());

    const redToks = fillTextCalls.filter((c) => c.text.replace(/\s/g, '') === 'red');
    const plainToks = fillTextCalls.filter((c) => c.text.replace(/\s/g, '') === 'plain.');
    expect(redToks.length).toBeGreaterThan(0);
    expect(plainToks.length).toBeGreaterThan(0);
    for (const t of redToks) expect(t.fillStyle.toLowerCase()).toBe('#ff0000');
    for (const t of plainToks) expect(t.fillStyle.toLowerCase()).toBe('#000000');
  });

  // ECMA-376 §17.3.2.26: within ONE run each character picks the ascii font
  // (Latin/digits) or the eastAsia font (CJK) by its Unicode block. A text-box
  // run carries BOTH axes (`fontFamily` = ascii, `fontFamilyEastAsia`). sample-10's
  // title run "第11回…" has eastAsia="ＭＳ ゴシック" (a gothic/sans) and falls through
  // to the docDefault ascii "Century" (a serif) for the embedded digits — Word
  // draws "第","回" sans and "11" serif. `fontFamilyClasses` drives the font CLASS
  // (fontTable §17.8.3.10): Century→roman (serif tail), ＭＳ ゴシック→swiss (sans tail).
  it('picks the eastAsia font for CJK chars and the ascii font for digits in one run', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const fontFamilyClasses = { Century: 'roman', 'ＭＳ ゴシック': 'swiss' };
    const shape = richTextbox([
      { text: '第11回', fontSizePt: 10, fontFamily: 'Century', fontFamilyEastAsia: 'ＭＳ ゴシック' },
    ]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 2000, 400, ctx, 1, fontFamilyClasses, new Map());

    // All ink preserved.
    const ink = fillTextCalls.map((c) => c.text).join('');
    expect(ink).toBe('第11回');

    // CJK tokens ("第","回") draw with the eastAsia family (ＭＳ ゴシック → sans tail);
    // the digit token ("11") draws with the ascii family (Century → serif tail).
    const cjkToks = fillTextCalls.filter((c) => c.text === '第' || c.text === '回');
    const digitToks = fillTextCalls.filter((c) => c.text.replace(/\s/g, '') === '11');
    expect(cjkToks.length).toBe(2);
    expect(digitToks.length).toBeGreaterThan(0);
    for (const t of cjkToks) {
      expect(t.font).toContain('"ＭＳ ゴシック"');
      expect(t.font).toContain('sans-serif');
      expect(t.font).not.toContain('"Century"');
    }
    for (const t of digitToks) {
      expect(t.font).toContain('"Century"');
      expect(t.font).toContain('serif');
      expect(t.font).not.toContain('"ＭＳ ゴシック"');
    }
  });

  // Back-compat: a run with only `fontFamily` (no eastAsia axis) keeps drawing
  // every token — CJK included — with that single family (the renderer falls back
  // `fontFamilyEastAsia ?? fontFamily`).
  it('falls back to the ascii font for CJK when no eastAsia axis is present', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const fontFamilyClasses = { 'Yu Mincho': 'roman' };
    const shape = richTextbox([
      { text: '本文 text', fontSizePt: 10, fontFamily: 'Yu Mincho' },
    ]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 2000, 400, ctx, 1, fontFamilyClasses, new Map());

    for (const c of fillTextCalls.filter((t) => t.text.trim().length > 0)) {
      expect(c.font).toContain('"Yu Mincho"');
    }
  });
});

// A rich (per-run) text box draws one token per fillText, so — unlike the
// single-fillText plain path, where the canvas reorders internally — the tokens
// must be reordered by the UAX#9 visual pass (rule L2, the same one body
// paragraphs use). Without it, RTL/mixed text drew in logical order and Word's
// word order was reversed (sample-8's yellow text box).
describe('retained text-box paint — rich-text RTL visual reordering (UAX#9 L2)', () => {
  const run = (text: string): ShapeTextRun =>
    ({ text, fontSizePt: 10 }) as unknown as ShapeTextRun;
  // Drawn texts sorted by x = visual left-to-right reading order.
  const visualOrder = (calls: { text: string; x: number }[]) =>
    [...calls].sort((a, b) => a.x - b.x).map((c) => c.text.trim()).filter(Boolean);

  it('reverses pure-RTL words into visual order', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    // Logical Arabic "one two three"; under an RTL base the visual L→R order is
    // the reverse: three, two, one.
    const shape = richTextbox([run('واحد اثنان ثلاثة')], 'right');
    acquireAndPaintShapeTextBox(shape, 0, 0, 300, 100, ctx, 1);
    expect(visualOrder(fillTextCalls)).toEqual(['ثلاثة', 'اثنان', 'واحد']);
  });

  it('keeps a Latin/number island left-to-right inside an RTL line', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    // "value 2025 end" in Arabic with an embedded number: words reverse, but the
    // number stays a single LTR island in the middle.
    const shape = richTextbox([run('قيمة 2025 نهاية')], 'right');
    acquireAndPaintShapeTextBox(shape, 0, 0, 300, 100, ctx, 1);
    expect(visualOrder(fillTextCalls)).toEqual(['نهاية', '2025', 'قيمة']);
  });

  it('leaves pure-LTR rich text in logical order', () => {
    const { ctx, fillTextCalls } = makeRecordingCtx();
    const shape = richTextbox([run('alpha beta gamma')], 'left');
    acquireAndPaintShapeTextBox(shape, 0, 0, 300, 100, ctx, 1);
    expect(visualOrder(fillTextCalls)).toEqual(['alpha', 'beta', 'gamma']);
  });
});

// Local alias matching renderer.ts's internal DecodedImage union.
type DecodedImage = ImageBitmap | HTMLImageElement;
