import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock ONLY core's path-keyed bitmap cache so it resolves to `null` — the
// legitimate "no drawable output" a true EMF or geometry-less metafile produces
// (see core getCachedBitmapByPath's contract). Everything else in core is the
// real module (importOriginal). This file is deliberately separate from
// renderer.image.test.ts, which needs the REAL cache to count fetch/decode
// calls; a hoisted vi.mock is file-scoped, so keeping the two suites apart lets
// each see the core it needs.
vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return {
    ...actual,
    getCachedBitmapByPath: vi.fn(async () => null),
    getCachedSvgImageByPath: vi.fn(async () => {
      throw new Error('corrupt SVG payload');
    }),
  };
});

import { preloadImages } from './test-support/renderer-internals.test-support.js';
import { renderDocumentToCanvas } from './renderer';
import { getCachedBitmapByPath } from '@silurus/ooxml-core';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  ImageRun,
  SectionProps,
} from './types';

/**
 * A raster/metafile blip can legitimately produce NO drawable output — a true
 * EMF, or a WMF with no geometry — in which case core's shared
 * `getCachedBitmapByPath` resolves to `null` (not an error). docx must treat
 * that as "missing image": drop it from the preloaded map and skip it at every
 * draw site, exactly like pptx's `if (!bitmap) return` and xlsx's falsy-skip.
 * The whole page/document render must NOT throw.
 *
 * These tests pin that contract at two levels:
 *   1. `preloadImages` returns a map that simply OMITS the undrawable blip
 *      (never throws / rejects, never stores an `undefined` value).
 *   2. `renderDocumentToCanvas` renders a page whose only picture is undrawable
 *      without throwing, and without emitting a `drawImage` for the missing
 *      bitmap.
 */

const fetchImage = vi.fn(
  async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
);

function imageRun(imagePath: string, extra: Partial<ImageRun> = {}): ImageRun {
  return {
    type: 'image',
    imagePath,
    mimeType: 'image/emf', // sniffed to a true EMF upstream → null base bitmap
    widthPt: 40,
    heightPt: 30,
    // inline (wp:inline): flows with text and reaches renderInlineImage's draw
    anchor: false,
    ...extra,
  } as unknown as ImageRun;
}

function imageDoc(runs: ImageRun[]): DocxDocumentModel {
  const para: DocParagraph = {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: runs as unknown as DocParagraph['runs'],
    defaultFontSize: 16, defaultFontFamily: 'Arial',
    widowControl: false,
  };
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...para } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

/** Recording 2D context that spies on drawImage; mirrors run-inline-formatting's
 *  makeRecordingCanvas (synthetic 0.8/0.2-em metrics, charCount × fontPx advance). */
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  drawImageCount: () => number;
} {
  let font = '16px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
  let drawImageCalls = 0;
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
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, clip() {}, rect() {},
    scale() {}, translate() {}, setLineDash() {}, clearRect() {}, fillRect() {}, strokeRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    drawImage() { drawImageCalls++; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, drawImageCount: () => drawImageCalls };
}

describe('docx undrawable blip (null base bitmap) is skipped, never crashes', () => {
  beforeEach(() => {
    // createImageBitmap is absent in the node test env; a sentinel is enough for
    // any path that DID decode (none should, since the cache returns null).
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 2, height: 2, close: () => {} }) as unknown as ImageBitmap),
    );
    fetchImage.mockClear();
    // The late-rejection test waits for the call started by that test. Source
    // inspection adds an await before bitmap resolution, so cumulative calls
    // from earlier cases must not make the readiness assertion pass early.
    vi.mocked(getCachedBitmapByPath).mockClear();
  });
  afterEach(() => vi.unstubAllGlobals());

  it('preloadImages omits a blip whose base bitmap decodes to null (no throw, no undefined value)', async () => {
    const doc = imageDoc([imageRun('word/media/image1.emf')]);
    // Must resolve — never reject — even though the only image is undrawable.
    const map = await preloadImages(doc, fetchImage);
    expect(map.has('word/media/image1.emf')).toBe(false);
    expect(map.size).toBe(0);
    // The map never stores an undefined/null value under any key.
    for (const v of map.values()) expect(v).toBeTruthy();
  });

  it('preloadImages drops every undrawable blip and still resolves to a well-formed (empty) map', async () => {
    // With the cache stubbed to null for ALL paths, every image drops; the point
    // is the map is well-formed (empty) and the call resolved rather than rejected.
    const doc = imageDoc([
      imageRun('word/media/image1.emf'),
      imageRun('word/media/image2.emf'),
    ]);
    const map = await preloadImages(doc, fetchImage);
    expect(map.size).toBe(0);
  });

  it('renderDocumentToCanvas renders a page whose only picture is undrawable without throwing (and draws no image)', async () => {
    const { canvas, drawImageCount } = makeRecordingCanvas();
    const doc = imageDoc([imageRun('word/media/image1.emf')]);
    await expect(
      renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400, fetchImage }),
    ).resolves.toBeUndefined();
    // The missing bitmap must not reach the draw path.
    expect(drawImageCount()).toBe(0);
  });

  it('renderDocumentToCanvas rejects a corrupt image instead of treating it as unsupported', async () => {
    vi.mocked(getCachedBitmapByPath).mockRejectedValueOnce(new Error('corrupt image payload'));
    const { canvas } = makeRecordingCanvas();
    const doc = imageDoc([imageRun('word/media/corrupt.emf')]);

    await expect(
      renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400, fetchImage }),
    ).rejects.toThrow('corrupt image payload');
  });

  it('preserves a corrupt SVG error when its raster fallback is legitimately undrawable', async () => {
    const doc = imageDoc([imageRun('word/media/fallback.emf', {
      svgImagePath: 'word/media/corrupt.svg',
    })]);

    await expect(preloadImages(doc, fetchImage)).rejects.toThrow('corrupt SVG payload');
  });

  it('accepts a drawable raster fallback after the preferred SVG decode fails', async () => {
    const raster = { width: 2, height: 2 } as unknown as ImageBitmap;
    vi.mocked(getCachedBitmapByPath).mockResolvedValueOnce(raster);
    const path = 'word/media/fallback.png';
    const doc = imageDoc([imageRun(path, {
      mimeType: 'image/png',
      svgImagePath: 'word/media/corrupt.svg',
    })]);

    await expect(preloadImages(doc, fetchImage)).resolves.toEqual(new Map([[path, raster]]));
  });

  it('suppresses a late corrupt-image rejection after a newer render owns the canvas', async () => {
    let rejectDecode: ((reason?: unknown) => void) | undefined;
    vi.mocked(getCachedBitmapByPath).mockImplementationOnce(() => new Promise((_, reject) => {
      rejectDecode = reject;
    }));
    const { canvas } = makeRecordingCanvas();
    const stale = renderDocumentToCanvas(
      imageDoc([imageRun('word/media/late-corrupt.emf')]),
      canvas,
      0,
      { dpr: 1, width: 400, fetchImage },
    );
    await vi.waitFor(() => expect(getCachedBitmapByPath).toHaveBeenCalled());

    await renderDocumentToCanvas(imageDoc([]), canvas, 0, { dpr: 1, width: 400, fetchImage });
    rejectDecode?.(new Error('late corrupt image payload'));

    await expect(stale).resolves.toBeUndefined();
  });
});
