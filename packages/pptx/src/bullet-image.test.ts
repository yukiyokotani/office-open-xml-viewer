import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  renderSlide,
  renderSlideWithEmbeddedFonts,
  renderTextBody,
  getCachedBitmap,
  dropImageBitmapCache,
} from './renderer.js';
import type { Slide, TextBody, Paragraph, BlipBullet } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

/**
 * Picture bullets (`<a:buBlip>`, ECMA-376 §21.1.2.4.2) are drawn inside the
 * synchronous text-body layout. The renderer warms the image cache up front
 * (renderSlide's prefetch pass), then the draw reads the prepared source (or a
 * directly warmed bitmap in focused tests) and paints it with `ctx.drawImage`.
 * These tests drive
 * `renderTextBody` directly against a mock 2D context that records `drawImage`,
 * mirroring the mock-ctx approach in text-highlight.test.ts / tabular-text.test.ts.
 *
 * EMU_PER_PT = 12700 and emuToPx(emu, scale) = emu * scale, so scale = 1/12700
 * makes "1pt → 1px": a 20pt run yields a 20px bullet box (× buSzPct).
 */
const SCALE = 1 / 12700;

// A sentinel ImageBitmap the stubbed createImageBitmap returns, so we can assert
// the exact object reaches drawImage. Intentionally NON-square (16×8, ratio 2:1)
// so the tests catch a forced-square draw: PowerPoint scales a picture bullet to
// the text height while preserving the bitmap's aspect ratio (§21.1.2.4.2).
const SENTINEL_W = 16;
const SENTINEL_H = 8;
const SENTINEL_RATIO = SENTINEL_W / SENTINEL_H;
const SENTINEL = {
  width: SENTINEL_W,
  height: SENTINEL_H,
  close: () => {},
} as unknown as ImageBitmap;

function mockCtx() {
  const draws: Array<{ img: unknown; x: number; y: number; w: number; h: number }> = [];
  const texts: Array<{ text: string; x: number; y: number }> = [];
  let fillStyle = '';
  let font = '';
  let direction: CanvasDirection = 'ltr';
  const ctx = {
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    get direction() {
      return direction;
    },
    set direction(v: CanvasDirection) {
      direction = v;
    },
    // Every glyph advances 10px so the line has a measurable, predictable width.
    measureText: (s: string) => ({
      width: s.length * 10,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
    fillText: (t: string, x: number, y: number) => texts.push({ text: t, x, y }),
    fillRect: () => {},
    drawImage: (img: unknown, x: number, y: number, w: number, h: number) =>
      draws.push({ img, x, y, w, h }),
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
  return { ctx: ctx as unknown as CanvasRenderingContext2D, draws, texts };
}

function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily: 'Arial',
    ...over,
  };
}

function bodyWithBullet(bullet: Paragraph['bullet'], runs: TextRunData[] = [run('Item')]): TextBody {
  const para: Paragraph = {
    alignment: 'l',
    marL: 457200, // a normal hanging-indent list metric so the bullet has a gutter
    marR: 0,
    indent: -457200,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet,
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs,
  } as Paragraph;
  return {
    verticalAnchor: 't',
    paragraphs: [para],
    defaultFontSize: 20,
    defaultBold: null,
    defaultItalic: null,
    lIns: 91440,
    rIns: 91440,
    tIns: 45720,
    bIns: 45720,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
  };
}

// A picture-bullet variant. Cast through the PPTX Bullet union (the parser emits
// `type: "blip"`; the statically-narrower core Bullet doesn't list it).
function blipBullet(over: Partial<BlipBullet> = {}): Paragraph['bullet'] {
  const b: BlipBullet = {
    type: 'blip',
    imagePath: 'ppt/media/bullet-img.png',
    mimeType: 'image/png',
    sizePct: null,
    ...over,
  };
  return b as unknown as Paragraph['bullet'];
}

type FetchImageFn = (path: string, mime: string) => Promise<Blob>;

/** Minimal PNG signature + IHDR prefix for the core dimension sniffer. */
function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  new DataView(bytes.buffer).setUint32(16, width);
  new DataView(bytes.buffer).setUint32(20, height);
  return bytes;
}

function slideCanvas(draws: unknown[]): HTMLCanvasElement {
  const canvas = {
    width: 0,
    height: 0,
    style: {} as CSSStyleDeclaration,
    offsetWidth: 960,
  } as HTMLCanvasElement;
  const state: Record<string, unknown> = {
    canvas,
    fillStyle: '',
    strokeStyle: '',
    globalAlpha: 1,
    lineWidth: 1,
    direction: 'ltr',
    measureText: (text: string) => ({
      width: text.length * 10,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    }),
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    drawImage: (image: unknown) => draws.push(image),
  };
  const context = new Proxy(state, {
    get(target, property: string) {
      if (property in target) return target[property];
      return () => undefined;
    },
    set(target, property: string, value) {
      target[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  canvas.getContext = (() => context) as unknown as HTMLCanvasElement['getContext'];
  return canvas;
}

describe('renderTextBody — picture bullet (buBlip) draws the bitmap', () => {
  let fetchImage: FetchImageFn;

  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (_blob: Blob) => SENTINEL),
    );
    fetchImage = vi.fn(
      async (_path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    ) as FetchImageFn;
  });
  afterEach(() => {
    dropImageBitmapCache(fetchImage);
    vi.unstubAllGlobals();
  });

  it('draws the warmed bullet bitmap sized to the text height with the bitmap aspect ratio, at the bullet gutter', async () => {
    const path = 'ppt/media/bullet-warmed.png';
    // Warm the cache the way renderSlide's prefetch pass does, then await it so
    // the settled bitmap is visible to the synchronous draw.
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path })),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1);
    const d = draws[0];
    expect(d.img).toBe(SENTINEL);
    // Height = 20pt run × scale(1/12700) × 12700 = 20px (default buSzPct = 100%).
    expect(d.h).toBeCloseTo(20, 6);
    // Width preserves the bitmap aspect ratio (16:8 = 2:1), not forced square.
    expect(d.w).toBeCloseTo(20 * SENTINEL_RATIO, 6);
    expect(d.w).toBeCloseTo(d.h * SENTINEL_RATIO, 6);
  });

  it('preserves a safe native cache variant prepared by renderSlide', async () => {
    const path = 'ppt/media/large-picture-bullet.png';
    const largePng = pngHeader(4096, 2048);
    fetchImage = vi.fn(async (_path: string, mime: string) =>
      new Blob([largePng as BlobPart], { type: mime })) as FetchImageFn;
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'shape',
        x: 0,
        y: 0,
        width: 4_572_000,
        height: 2_286_000,
        rotation: 0,
        flipH: false,
        flipV: false,
        geometry: 'rect',
        fill: null,
        stroke: null,
        textBody: bodyWithBullet(blipBullet({ imagePath: path })),
      }],
    } as Slide;
    const draws: unknown[] = [];

    await renderSlide(slideCanvas(draws), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage,
    });

    // This 8 MP source fits both the per-surface and aggregate budgets, so the
    // renderer preserves the established native decode. The synchronous bullet
    // paint must use that exact prepared result.
    expect(draws).toContain(SENTINEL);
    expect(globalThis.createImageBitmap).toHaveBeenCalledWith(expect.any(Blob));
  });

  it('routes SVG picture bullets through the worker bridge at their marker height', async () => {
    const path = 'ppt/media/svg-picture-bullet.svg';
    const svg = new Blob([
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 8"><rect width="16" height="8"/></svg>',
    ], { type: 'image/svg+xml' });
    fetchImage = vi.fn(async () => svg) as FetchImageFn;
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      throw new Error('Chromium workers cannot decode this SVG Blob');
    }));
    let bridgedBitmap: ImageBitmap | undefined;
    const svgDecoder = vi.fn(async (
      _blob: Blob,
      target: { targetWidthPx?: number; targetHeightPx?: number } = {},
    ) => {
      const height = target.targetHeightPx ?? 0;
      bridgedBitmap = {
        width: height * 2,
        height,
        close: () => {},
      } as unknown as ImageBitmap;
      return bridgedBitmap;
    });
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'shape',
        x: 0,
        y: 0,
        width: 4_572_000,
        height: 2_286_000,
        rotation: 0,
        flipH: false,
        flipV: false,
        geometry: 'rect',
        fill: null,
        stroke: null,
        textBody: bodyWithBullet(blipBullet({ imagePath: path, mimeType: 'image/svg+xml' })),
      }],
    } as Slide;
    const draws: unknown[] = [];

    await renderSlideWithEmbeddedFonts(slideCanvas(draws), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
      svgDecoder,
    });

    expect(svgDecoder).toHaveBeenCalledWith(svg, {
      targetWidthPx: 1,
      targetHeightPx: 54,
    });
    expect(bridgedBitmap).toMatchObject({ width: 108, height: 54 });
    expect(draws).toContain(bridgedBitmap);
  });

  it('loads an SVG picture bullet through HTMLImageElement in Window mode', async () => {
    const path = 'ppt/media/window-svg-picture-bullet.svg';
    const svg = new Blob([
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 8"><rect width="16" height="8"/></svg>',
    ], { type: 'image/svg+xml' });
    fetchImage = vi.fn(async () => svg) as FetchImageFn;
    const createImageBitmap = vi.fn(async () => {
      throw new DOMException('The source image could not be decoded', 'InvalidStateError');
    });
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    vi.stubGlobal('URL', {
      createObjectURL: vi.fn(() => 'blob:window-svg-picture-bullet'),
      revokeObjectURL: vi.fn(),
    });
    class FakeSvgImage {
      width = 16;
      height = 8;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      decode = vi.fn(async () => undefined);
      set src(_value: string) { queueMicrotask(() => this.onload?.()); }
    }
    vi.stubGlobal('Image', FakeSvgImage);
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'shape',
        x: 0,
        y: 0,
        width: 4_572_000,
        height: 2_286_000,
        rotation: 0,
        flipH: false,
        flipV: false,
        geometry: 'rect',
        fill: null,
        stroke: null,
        textBody: bodyWithBullet(blipBullet({ imagePath: path, mimeType: 'image/svg+xml' })),
      }],
    } as Slide;
    const draws: unknown[] = [];

    await renderSlide(slideCanvas(draws), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage,
    });

    expect(draws).toHaveLength(1);
    expect(draws[0]).toBeInstanceOf(FakeSvgImage);
    expect(createImageBitmap).not.toHaveBeenCalled();
  });

  it('dedupes a shared SVG bullet at the largest authored marker height', async () => {
    const path = 'ppt/media/shared-svg-picture-bullet.svg';
    const svg = new Blob([
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 8"><rect width="16" height="8"/></svg>',
    ], { type: 'image/svg+xml' });
    fetchImage = vi.fn(async () => svg) as FetchImageFn;
    const svgDecoder = vi.fn(async () => SENTINEL);
    const shape = (y: number, sizePts?: number) => ({
      type: 'shape' as const,
      x: 0,
      y,
      width: 4_572_000,
      height: 2_286_000,
      rotation: 0,
      flipH: false,
      flipV: false,
      geometry: 'rect',
      fill: null,
      stroke: null,
      textBody: bodyWithBullet(blipBullet({
        imagePath: path,
        mimeType: 'image/svg+xml',
        ...(sizePts === undefined ? {} : { sizePts }),
      })),
    });
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [shape(0), shape(2_286_000, 40)],
    } as Slide;
    const draws: unknown[] = [];

    await renderSlideWithEmbeddedFonts(slideCanvas(draws), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
      svgDecoder,
    });

    expect(svgDecoder).toHaveBeenCalledOnce();
    expect(svgDecoder).toHaveBeenCalledWith(svg, {
      targetWidthPx: 1,
      // 40pt × 96 CSS px/in ÷ 72pt/in × DPR 2, rounded up by the cache.
      targetHeightPx: 107,
    });
    expect(draws.filter((image) => image === SENTINEL)).toHaveLength(2);
  });

  it('scales the bullet by buSzPct (§21.1.2.4.9)', async () => {
    const path = 'ppt/media/bullet-sized.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path, sizePct: 50 })),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1);
    // Height = 20px text × 50% = 10px; width preserves the 2:1 aspect ratio.
    expect(draws[0].h).toBeCloseTo(10, 6);
    expect(draws[0].w).toBeCloseTo(10 * SENTINEL_RATIO, 6);
  });

  it('sizes the bullet by buSzPts absolutely (§21.1.2.4.10), independent of the run', async () => {
    const path = 'ppt/media/bullet-pts.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      // Run is 20pt; buSzPts = 40pt → the bullet box is 40px (40pt at this
      // scale), NOT the 20px run height — the size is absolute, not relative.
      bodyWithBullet(blipBullet({ imagePath: path, sizePts: 40 })),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1);
    expect(draws[0].h).toBeCloseTo(40, 6);
    expect(draws[0].w).toBeCloseTo(40 * SENTINEL_RATIO, 6);
  });

  it('prefers buSzPts over a co-present buSzPct on a picture bullet', async () => {
    const path = 'ppt/media/bullet-pts-over-pct.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      // Absolute 30pt → 30px wins over 200% × 20pt run (= 40px): the two are the
      // one EG_TextBulletSize choice, and the absolute size takes precedence.
      bodyWithBullet(blipBullet({ imagePath: path, sizePts: 30, sizePct: 200 })),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1);
    expect(draws[0].h).toBeCloseTo(30, 6);
    expect(draws[0].w).toBeCloseTo(30 * SENTINEL_RATIO, 6);
  });

  it('draws nothing (no throw) when the bullet image is not yet decoded', () => {
    // No getCachedBitmap warm-up → peekCachedBitmap returns undefined.
    const { ctx, draws } = mockCtx();
    expect(() =>
      renderTextBody(
        ctx,
        bodyWithBullet(blipBullet({ imagePath: 'ppt/media/cold.png' })),
        0, 0, 4000, 2000,
        SCALE,
        null, 0, false, false, '#000000', 1,
        { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
        undefined,
        false,
        fetchImage,
      ),
    ).not.toThrow();
    expect(draws).toHaveLength(0);
  });

  it('does not draw a picture bullet on an empty paragraph', async () => {
    const path = 'ppt/media/bullet-empty.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws } = mockCtx();
    // Empty paragraph (no runs) — PowerPoint draws no marker.
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path }), []),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(0);
  });

  it('does not call drawImage for a char bullet (regression guard)', async () => {
    // A char bullet must still go through fillText, never drawImage.
    const { ctx, draws } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet({ type: 'char', char: '•', color: null, sizePct: null, fontFamily: 'Arial' }),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );
    expect(draws).toHaveLength(0);
  });

  it('first-line text does not overlap the picture bullet (hanging indent)', async () => {
    // Regression guard for the first-line hanging-indent overlap. A picture
    // bullet occupies the gutter exactly like a char/autoNum marker, so the
    // first line's hanging indent MUST be suppressed. Before the fix `hasBullet`
    // excluded `blip`, so the first line started at the bullet's x and rendered
    // ON TOP of the image (same class as docx PR #476). The mock records fillText
    // x, which the other tests don't, so this is the case they couldn't catch.
    const path = 'ppt/media/bullet-hang.png';
    await getCachedBitmap(path, 'image/png', fetchImage);

    const { ctx, draws, texts } = mockCtx();
    renderTextBody(
      ctx,
      bodyWithBullet(blipBullet({ imagePath: path }), [run('Item')]),
      0, 0, 4000, 2000,
      SCALE,
      null, 0, false, false, '#000000', 1,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
      undefined,
      false,
      fetchImage,
    );

    expect(draws).toHaveLength(1); // the bullet image was drawn
    const bulletX = draws[0].x;
    expect(texts.length).toBeGreaterThan(0); // the run text was drawn
    const firstTextX = texts[0].x;
    // First-line text starts to the RIGHT of the bullet by the hanging gap
    // (|indent| = 457200 EMU = 36px at this scale), NOT at the bullet's x.
    // Pre-fix this difference was 0 (overlap) → this assertion fails without it.
    expect(firstTextX).toBeGreaterThan(bulletX);
    expect(firstTextX - bulletX).toBeCloseTo(457200 * SCALE, 4);
  });
});
