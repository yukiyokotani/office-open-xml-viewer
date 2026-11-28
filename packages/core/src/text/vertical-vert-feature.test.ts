import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  measureVerticalVertGlyph,
  verticalVertGlyphReachable,
  withVertFeature,
} from './vertical-vert-feature.js';

type Listener = () => void;

interface ProbeFixture {
  ctx: CanvasRenderingContext2D;
  featuresAtPaint: string[];
  fontAssignments: () => number;
  paints: () => number;
  listeners: Map<string, Listener>;
  variants: Map<number, 'wide-to-tall' | 'placement' | 'tall-to-wide' | 'outline' | 'metrics'>;
}

function fakeHtmlCanvasProbe(
  initialFeature = 'normal',
  options: { computedFeature?: string; jitter?: boolean } = {},
): ProbeFixture {
  const listeners = new Map<string, Listener>();
  const variants = new Map<number, 'wide-to-tall' | 'placement' | 'tall-to-wide' | 'outline' | 'metrics'>();
  const featuresAtPaint: string[] = [];
  let paints = 0;
  let fontAssignments = 0;
  let rasterReads = 0;

  const effectiveFeature = (canvas: FakeCanvas): string =>
    canvas.style.fontFeatureSettings || options.computedFeature || 'normal';

  class FakeCanvas {
    width = 1;
    height = 1;
    isConnected = false;
    ownerDocument: typeof document;
    style = { fontFeatureSettings: initialFeature };
    private context: CanvasRenderingContext2D | null = null;

    constructor(ownerDocument: typeof document) {
      this.ownerDocument = ownerDocument;
    }

    setAttribute() {}
    remove() { this.isConnected = false; }

    getContext(): CanvasRenderingContext2D {
      if (this.context !== null) return this.context;
      const canvas = this;
      let font = '';
      let resolvedFeature = effectiveFeature(canvas);
      let drawnCp = 0;
      this.context = {
        canvas,
        get font() { return font; },
        set font(value: string) {
          font = value;
          resolvedFeature = effectiveFeature(canvas);
          fontAssignments += 1;
        },
        fillStyle: '#000',
        textAlign: 'center',
        textBaseline: 'middle',
        clearRect() {},
        fillText(text: string) {
          drawnCp = text.codePointAt(0) ?? 0;
          featuresAtPaint.push(resolvedFeature);
          paints += 1;
        },
        measureText() {
          const featured = resolvedFeature.includes('"vert" 1');
          const metricsChanged = featured && variants.get(drawnCp) === 'metrics';
          return {
            width: 200,
            actualBoundingBoxLeft: 20,
            actualBoundingBoxRight: 20,
            actualBoundingBoxAscent: metricsChanged ? 90 : 80,
            actualBoundingBoxDescent: metricsChanged ? -30 : -20,
          } as TextMetrics;
        },
        getImageData() {
          const data = new Uint8ClampedArray(canvas.width * canvas.height * 4);
          const featured = resolvedFeature.includes('"vert" 1');
          const variant = variants.get(drawnCp);
          const tallToWide = variant === 'tall-to-wide';
          const inkWidth = featured && variant === 'wide-to-tall'
            ? 3
            : featured && tallToWide
              ? 9
              : tallToWide
                ? 3
                : 9;
          const inkHeight = featured && variant === 'wide-to-tall'
            ? 9
            : featured && tallToWide
              ? 3
              : tallToWide
                ? 9
                : variant === 'outline'
                  ? 9
                  : 3;
          const xOffset = (featured && variant === 'placement' ? 4 : 0)
            + (options.jitter ? rasterReads : 0);
          rasterReads += 1;
          for (let y = 0; y < inkHeight; y += 1) {
            for (let x = 0; x < inkWidth; x += 1) {
              if (featured && variant === 'outline' && x === 4 && y === 4) continue;
              data[(y * canvas.width + x + xOffset) * 4 + 3] = 255;
            }
          }
          return { data } as ImageData;
        },
      } as unknown as CanvasRenderingContext2D;
      return this.context;
    }
  }

  const parent = {
    appendChild(canvas: FakeCanvas) { canvas.isConnected = true; },
  };
  const fakeDocument = {
    body: parent,
    documentElement: parent,
    fonts: {
      addEventListener(name: string, listener: Listener) { listeners.set(name, listener); },
    },
    createElement: () => new FakeCanvas(fakeDocument as unknown as Document),
  } as unknown as Document;
  vi.stubGlobal('HTMLCanvasElement', FakeCanvas);
  vi.stubGlobal('document', fakeDocument);
  vi.stubGlobal('getComputedStyle', (element: FakeCanvas) => ({
    fontFeatureSettings: effectiveFeature(element),
  }));

  const canvas = new FakeCanvas(fakeDocument);
  let sourceFont = '16px "Probe Mincho", serif';
  const ctx = {
    canvas,
    get font() { return sourceFont; },
    set font(value: string) { sourceFont = value; },
  } as unknown as CanvasRenderingContext2D;
  return {
    ctx,
    featuresAtPaint,
    fontAssignments: () => fontAssignments,
    paints: () => paints,
    listeners,
    variants,
  };
}

describe('vertical OpenType vert feature', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reports unsupported for a context whose canvas has no element style', () => {
    const ctx = {
      canvas: { width: 1, height: 1 },
      font: '16px serif',
    } as unknown as OffscreenCanvasRenderingContext2D;

    expect(verticalVertGlyphReachable(ctx, 0x30fc)).toBe(false);
  });

  it('detects placement-only, tall-to-wide, wide-to-tall, and same-box outline alternates', () => {
    const fixture = fakeHtmlCanvasProbe();
    fixture.variants.set(0x3001, 'placement');
    fixture.variants.set(0x300c, 'tall-to-wide');
    fixture.variants.set(0x30fc, 'wide-to-tall');
    fixture.variants.set(0x3041, 'outline');
    fixture.variants.set(0x3002, 'metrics');

    expect(verticalVertGlyphReachable(fixture.ctx, 0x3001)).toBe(true);
    expect(verticalVertGlyphReachable(fixture.ctx, 0x300c)).toBe(true);
    expect(verticalVertGlyphReachable(fixture.ctx, 0x30fc)).toBe(true);
    expect(verticalVertGlyphReachable(fixture.ctx, 0x3041)).toBe(true);
    expect(verticalVertGlyphReachable(fixture.ctx, 0x3002)).toBe(true);
    expect(verticalVertGlyphReachable(fixture.ctx, 0xff1b)).toBe(false);
    expect(verticalVertGlyphReachable(fixture.ctx, 0xff01)).toBe(false);
  });

  it('forces vert off and on when the source already enables it', () => {
    const fixture = fakeHtmlCanvasProbe('"kern" 1, "vert" 1, "ss01" 1');
    fixture.variants.set(0x30fc, 'wide-to-tall');

    expect(verticalVertGlyphReachable(fixture.ctx, 0x30fc)).toBe(true);
    expect(fixture.featuresAtPaint).toEqual([
      '"kern" 1, "ss01" 1, "vert" 0',
      '"kern" 1, "ss01" 1, "vert" 0',
      '"kern" 1, "ss01" 1, "vert" 1',
      '"kern" 1, "ss01" 1, "vert" 1',
    ]);
  });

  it('uses computed feature settings as the probe source', () => {
    const fixture = fakeHtmlCanvasProbe('', {
      computedFeature: '"kern" 1, "vert" 1',
    });
    fixture.variants.set(0x300c, 'tall-to-wide');

    expect(verticalVertGlyphReachable(fixture.ctx, 0x300c)).toBe(true);
    expect(fixture.featuresAtPaint).toEqual([
      '"kern" 1, "vert" 0',
      '"kern" 1, "vert" 0',
      '"kern" 1, "vert" 1',
      '"kern" 1, "vert" 1',
    ]);
  });

  it('rejects raster jitter when vert produces no repeatable change', () => {
    const fixture = fakeHtmlCanvasProbe('normal', { jitter: true });

    expect(verticalVertGlyphReachable(fixture.ctx, 0x30fc)).toBe(false);
  });

  it('probes the requested glyph, refonts after each feature change, and caches per code point', () => {
    const fixture = fakeHtmlCanvasProbe('"kern" 1, "ss01" 1');
    fixture.variants.set(0x30fc, 'wide-to-tall');

    expect(verticalVertGlyphReachable(fixture.ctx, 0x30fc)).toBe(true);
    expect(fixture.featuresAtPaint).toEqual([
      '"kern" 1, "ss01" 1, "vert" 0',
      '"kern" 1, "ss01" 1, "vert" 0',
      '"kern" 1, "ss01" 1, "vert" 1',
      '"kern" 1, "ss01" 1, "vert" 1',
    ]);
    expect(fixture.fontAssignments()).toBeGreaterThanOrEqual(4);
    expect(verticalVertGlyphReachable(fixture.ctx, 0x30fc)).toBe(true);
    expect(fixture.paints()).toBe(4);

    expect(verticalVertGlyphReachable(fixture.ctx, 0x301c)).toBe(false);
    expect(fixture.paints()).toBe(8);
  });

  it('takes the route language as an argument, never from the caller context', () => {
    const fixture = fakeHtmlCanvasProbe();
    fixture.variants.set(0x30fc, 'wide-to-tall');
    const accesses: string[] = [];
    const ctx = new Proxy(fixture.ctx, {
      has(object, key) {
        if (key === 'lang') accesses.push('has');
        return Reflect.has(object, key);
      },
      get(object, key) {
        if (key === 'lang') accesses.push('get');
        return Reflect.get(object, key);
      },
    });
    expect(verticalVertGlyphReachable(ctx, 0x30fc)).toBe(true);
    expect(fixture.paints()).toBe(4);
    // A language selects forms of its own: it is probed and cached apart.
    expect(verticalVertGlyphReachable(ctx, 0x30fc, 'ja-jp')).toBe(true);
    expect(fixture.paints()).toBe(8);
    expect(verticalVertGlyphReachable(ctx, 0x30fc, 'ja-jp')).toBe(true);
    expect(verticalVertGlyphReachable(ctx, 0x30fc)).toBe(true);
    expect(fixture.paints()).toBe(8);
    expect(accesses).toEqual([]);
  });

  it('invalidates each glyph result after the FontFace epoch changes', () => {
    const fixture = fakeHtmlCanvasProbe();

    expect(verticalVertGlyphReachable(fixture.ctx, 0xff5e)).toBe(false);
    fixture.variants.set(0xff5e, 'wide-to-tall');
    expect(verticalVertGlyphReachable(fixture.ctx, 0xff5e)).toBe(false);
    expect(fixture.paints()).toBe(4);

    fixture.listeners.get('loadingdone')?.();
    expect(verticalVertGlyphReachable(fixture.ctx, 0xff5e)).toBe(true);
    expect(fixture.paints()).toBe(8);
  });

  it('keeps the featured advance origin fixed when designed ink pokes before the cell', () => {
    let font = '10px "Probe Mincho"';
    let feature = 'normal';
    const ctx = {
      canvas: { style: { get fontFeatureSettings() { return feature; }, set fontFeatureSettings(v: string) { feature = v; } } },
      get font() { return font; },
      set font(value: string) { font = value; },
      textAlign: 'left',
      textBaseline: 'alphabetic',
      measureText() {
        return {
          width: 10,
          actualBoundingBoxAscent: feature.includes('"vert" 1') ? 8 : 2,
          actualBoundingBoxDescent: feature.includes('"vert" 1') ? 5 : 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;

    expect(measureVerticalVertGlyph(ctx, 'ー')).toEqual({
      advancePx: 10,
      inkBeforePx: 8,
      inkAfterPx: 5,
      cellAdvancePx: 10,
      originInCellPx: 5,
    });
    expect(ctx.textAlign).toBe('left');
    expect(ctx.textBaseline).toBe('alphabetic');
    expect(feature).toBe('normal');
  });

  it('falls back to the featured advance when A/D metrics are unavailable', () => {
    let font = '10px serif';
    const style = { fontFeatureSettings: '' };
    const ctx = {
      canvas: { style },
      get font() { return font; },
      set font(value: string) { font = value; },
      textAlign: 'left',
      textBaseline: 'alphabetic',
      measureText: () => ({ width: 12 }) as TextMetrics,
    } as unknown as CanvasRenderingContext2D;

    expect(measureVerticalVertGlyph(ctx, '「')).toEqual({
      advancePx: 12,
      inkBeforePx: 0,
      inkAfterPx: 0,
      cellAdvancePx: 12,
      originInCellPx: 6,
    });
  });

  it('sets the canvas feature before refonting and restores both after drawing', () => {
    const events: string[] = [];
    let font = 'italic 700 16px "Hiragino Mincho ProN"';
    let feature = '"kern" 1';
    const style = {
      get fontFeatureSettings() {
        return feature;
      },
      set fontFeatureSettings(value: string) {
        feature = value;
        events.push(`feature:${value}`);
      },
    };
    const ctx = {
      canvas: { style },
      get font() {
        return font;
      },
      set font(value: string) {
        font = value;
        events.push(`font:${value}`);
      },
    } as unknown as CanvasRenderingContext2D;

    const result = withVertFeature(ctx, () => {
      events.push(`draw:${style.fontFeatureSettings}`);
      return 47;
    });

    expect(result).toBe(47);
    expect(events).toEqual([
      'feature:"kern" 1, "vert" 1',
      'font:italic 700 16px "Hiragino Mincho ProN"',
      'draw:"kern" 1, "vert" 1',
      'feature:"kern" 1',
      'font:italic 700 16px "Hiragino Mincho ProN"',
    ]);
  });

  it('briefly connects a detached element so measurement and paint resolve the same feature', () => {
    const fixture = fakeHtmlCanvasProbe();
    const canvas = fixture.ctx.canvas as HTMLCanvasElement;
    expect(canvas.isConnected).toBe(false);

    withVertFeature(fixture.ctx, () => {
      expect(canvas.isConnected).toBe(true);
    });

    expect(canvas.isConnected).toBe(false);
  });

  it('preserves computed features while restoring the exact inline value', () => {
    let font = '16px serif';
    const style = { fontFeatureSettings: '' };
    const ownerDocument = {
      defaultView: {
        getComputedStyle: () => ({ fontFeatureSettings: '"kern" 1, "ss01" 1' }),
      },
    } as unknown as Document;
    const ctx = {
      canvas: { style, ownerDocument },
      get font() { return font; },
      set font(value: string) { font = value; },
    } as unknown as CanvasRenderingContext2D;

    withVertFeature(ctx, () => {
      expect(style.fontFeatureSettings).toBe('"kern" 1, "ss01" 1, "vert" 1');
    });

    expect(style.fontFeatureSettings).toBe('');
  });

  it('restores the prior feature and refonts when drawing throws', () => {
    let fontAssignments = 0;
    let font = '16px serif';
    const style = { fontFeatureSettings: '' };
    const ctx = {
      canvas: { style },
      get font() {
        return font;
      },
      set font(value: string) {
        font = value;
        fontAssignments += 1;
      },
    } as unknown as CanvasRenderingContext2D;

    expect(() =>
      withVertFeature(ctx, () => {
        throw new Error('paint failed');
      }),
    ).toThrow('paint failed');
    expect(style.fontFeatureSettings).toBe('');
    expect(fontAssignments).toBe(2);
  });
});
