import { describe, expect, it, vi } from 'vitest';
import type { ChartExElementStyle } from '../types/chart.js';
import {
  chartStyleEffectRecipe,
  chartStyleEffectOwner,
  chartEffectRasterWorkUpperBound,
  paintChartStyleEffects,
  paintChartStyleOuterEffectBehind,
  withChartEffectBudget,
} from './style-effects.js';
import { withSparseStyleIndexCache } from './sparse-style-index.js';

const SHADOW = {
  color: '112233', alpha: 0.5, blur: 12_700, dist: 25_400, dir: 0,
};

describe('chart style effect precedence', () => {
  it('uses direct effects before the resolved linked/numeric role', () => {
    const fallback: ChartExElementStyle = { shadows: [SHADOW], effectAuthored: true };
    const direct = {
      glows: [{ color: '445566', alpha: 0.75, radius: 6_350 }],
      effectAuthored: true,
    } satisfies ChartExElementStyle;
    expect(chartStyleEffectRecipe(direct, fallback, 0)).toEqual({
      glow: direct.glows![0]!,
    });
  });

  it('treats authored empty and unsupported direct effects as suppression', () => {
    const fallback: ChartExElementStyle = { shadows: [SHADOW], effectAuthored: true };
    expect(chartStyleEffectRecipe({ effectAuthored: true }, fallback, 0)).toBeUndefined();
    expect(chartStyleEffectRecipe({ effectUnsupported: true }, fallback, 0)).toBeUndefined();
  });

  it('skips a more-specific style that authors only a different component', () => {
    const series = { shadows: [SHADOW], effectAuthored: true } satisfies ChartExElementStyle;
    expect(chartStyleEffectOwner({ fillHidden: true }, series)).toBe(series);
  });

  it('uses the fixed effect style color index for every effect palette', () => {
    const second = { ...SHADOW, color: 'AABBCC' };
    expect(chartStyleEffectRecipe({
      shadows: [SHADOW, second], effectAuthored: true, effectColorIndex: 1,
    }, undefined, 42)?.shadow).toEqual(second);
  });

  it('maps sparse source formatting indexes to compact effect slots', () => {
    const second = { ...SHADOW, color: 'DDEEFF' };
    expect(chartStyleEffectRecipe({
      shadows: [SHADOW, second],
      effectAuthored: true,
      effectFormattingIndices: [3, 9],
    }, undefined, 9)?.shadow).toEqual(second);
  });

  it('materializes a large sparse effect-index domain only once', () => {
    const source = Array.from({ length: 4_096 }, (_, index) => index * 3);
    let indexedReads = 0;
    const observed = new Proxy(source, {
      get(target, property, receiver) {
        if (typeof property === 'string' && /^\d+$/.test(property)) indexedReads++;
        return Reflect.get(target, property, receiver);
      },
    });
    const style = {
      shadows: [SHADOW],
      effectAuthored: true,
      effectFormattingIndices: observed,
    } satisfies ChartExElementStyle;
    withSparseStyleIndexCache(() => {
      expect(chartStyleEffectRecipe(style, undefined, source.at(-1)!)).toBeDefined();
      expect(indexedReads).toBe(4_096);
      expect(chartStyleEffectRecipe(style, undefined, source.at(-2)!)).toBeDefined();
    });
    expect(indexedReads).toBe(4_096);
  });
});

function recordingContext() {
  const stack: Array<Record<string, unknown>> = [];
  const ctx = {
    canvas: { width: 0, height: 0 },
    shadowColor: 'transparent',
    shadowBlur: 0,
    shadowOffsetX: 0,
    shadowOffsetY: 0,
    globalCompositeOperation: 'source-over',
    save: vi.fn(function (this: Record<string, unknown>) { stack.push({
      shadowColor: this.shadowColor,
      shadowBlur: this.shadowBlur,
      shadowOffsetX: this.shadowOffsetX,
      shadowOffsetY: this.shadowOffsetY,
      globalCompositeOperation: this.globalCompositeOperation,
    }); }),
    restore: vi.fn(function (this: Record<string, unknown>) { Object.assign(this, stack.pop()); }),
    setTransform: vi.fn(),
    translate: vi.fn(),
    scale: vi.fn(),
    transform: vi.fn(),
    drawImage: vi.fn(),
    getTransform: vi.fn(() => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 })),
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

describe('chart style effect painting', () => {
  it('charges a reflection as one full-canvas surface', () => {
    const reflection = {
      blur: 0, dist: 0, dir: 90, stA: 1, stPos: 0,
      endA: 0, endPos: 1, sx: 1, sy: -1,
    };
    expect(chartEffectRasterWorkUpperBound(
      { reflection }, { x: 5, y: 5, w: 10, h: 10 }, 1, 100, 80,
    )).toBe(8_000);
  });

  it('bounds full-canvas reflection and uses one composed outer-shadow surface', () => {
    class FakeAuxContext {
      globalCompositeOperation = 'source-over';
      fillStyle: unknown = '#000';
      save() {}
      restore() {}
      setTransform() {}
      fillRect() {}
      createLinearGradient() { return { addColorStop() {} }; }
    }
    class FakeOffscreenCanvas {
      constructor(public width: number, public height: number) {}
      getContext() { return new FakeAuxContext(); }
    }
    vi.stubGlobal('OffscreenCanvas', FakeOffscreenCanvas);
    const reflection = {
      blur: 0, dist: 0, dir: 90, stA: 1, stPos: 0,
      endA: 0, endPos: 1, sx: 1, sy: -1,
    };
    try {
      const admitted = recordingContext();
      admitted.canvas.width = 100;
      admitted.canvas.height = 80;
      const admittedBody = vi.fn();
      withChartEffectBudget(admitted, () => {
        paintChartStyleEffects(
          admitted,
          undefined,
          { reflections: [reflection], effectAuthored: true },
          0,
          { x: 5, y: 5, w: 10, h: 10 },
          1,
          admittedBody,
        );
        paintChartStyleEffects(
          admitted,
          undefined,
          { reflections: [reflection], effectAuthored: true },
          1,
          { x: 20, y: 5, w: 10, h: 10 },
          1,
          admittedBody,
        );
      }, 16_000, 2);
      // Equivalent consumers are admitted together; paint order cannot select
      // a styled prefix of the series.
      expect(admittedBody).toHaveBeenCalledTimes(4);
      expect(admitted.drawImage).toHaveBeenCalledTimes(2);

      const rejected = recordingContext();
      rejected.canvas.width = 100;
      rejected.canvas.height = 80;
      const rejectedBody = vi.fn();
      withChartEffectBudget(rejected, () => {
        for (let index = 0; index < 2; index++) {
          paintChartStyleEffects(
            rejected,
            undefined,
            { reflections: [reflection], effectAuthored: true },
            index,
            { x: 5 + index * 15, y: 5, w: 10, h: 10 },
            1,
            rejectedBody,
          );
        }
      }, 15_999, 2);
      expect(rejectedBody).toHaveBeenCalledTimes(2);
      expect(rejected.drawImage).not.toHaveBeenCalled();

      const composedShadow = recordingContext();
      composedShadow.canvas.width = 100;
      composedShadow.canvas.height = 80;
      const shadowBody = vi.fn();
      withChartEffectBudget(composedShadow, () => paintChartStyleEffects(
        composedShadow,
        undefined,
        { shadows: [SHADOW], effectAuthored: true },
        0,
        { x: 5, y: 5, w: 10, h: 10 },
        1,
        shadowBody,
      ), 8_000);
      expect(shadowBody).toHaveBeenCalledTimes(2);
      expect(composedShadow.drawImage).toHaveBeenCalledTimes(1);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('renders an outer shadow through the allocation-free Canvas fallback', () => {
    const ctx = recordingContext();
    const observations: Array<[string, number, number]> = [];
    paintChartStyleEffects(
      ctx,
      undefined,
      { shadows: [SHADOW], effectAuthored: true },
      0,
      { x: 10, y: 20, w: 30, h: 40 },
      1,
      target => observations.push([
        target.shadowColor as string,
        target.shadowOffsetX,
        target.shadowOffsetY,
      ]),
    );
    expect(observations).toEqual([['rgba(17,34,51,0.5)', 2, 0]]);
    expect(ctx.shadowColor).toBe('transparent');
  });

  it('paints the body once with no effects and under a zero raster budget', () => {
    const ctx = recordingContext();
    const body = vi.fn();
    withChartEffectBudget(ctx, () => {
      for (let index = 0; index < 2; index++) {
        paintChartStyleEffects(
          ctx,
          undefined,
          { shadows: [SHADOW], effectAuthored: true },
          index,
          { x: index * 10, y: 0, w: 10, h: 10 },
          1,
          body,
        );
      }
    }, 0);
    expect(body).toHaveBeenCalledTimes(2);

    body.mockClear();
    paintChartStyleEffects(ctx, undefined, undefined, 0, { x: 0, y: 0, w: 1, h: 1 }, 1, body);
    expect(body).toHaveBeenCalledTimes(1);
  });

  it('composites only allocation-free outer effects behind a completed 3-D scene', () => {
    const ctx = recordingContext();
    const observations: Array<[string, string]> = [];
    paintChartStyleOuterEffectBehind(
      ctx,
      undefined,
      { shadows: [SHADOW], effectAuthored: true },
      0,
      1,
      target => observations.push([
        target.globalCompositeOperation,
        target.shadowColor as string,
      ]),
    );
    expect(observations).toEqual([['destination-over', 'rgba(17,34,51,0.5)']]);
    expect(ctx.globalCompositeOperation).toBe('source-over');

    const rasterOnlyBody = vi.fn();
    paintChartStyleOuterEffectBehind(
      ctx,
      { reflections: [{
        blur: 0, dist: 0, dir: 90, stA: 1, stPos: 0,
        endA: 0, endPos: 1, sx: 1, sy: -1,
      }], effectAuthored: true },
      undefined,
      0,
      1,
      rasterOnlyBody,
    );
    expect(rasterOnlyBody).not.toHaveBeenCalled();
  });

  it('selects a linked 3-D effect in its own series/point formatting domain', () => {
    const ctx = recordingContext();
    const observations: string[] = [];
    paintChartStyleOuterEffectBehind(
      ctx,
      undefined,
      {
        shadows: [SHADOW, { ...SHADOW, color: 'AABBCC' }],
        effectAuthored: true,
      },
      7,
      1,
      target => observations.push(target.shadowColor as string),
      1,
    );
    expect(observations).toEqual(['rgba(170,187,204,0.5)']);
  });
});
