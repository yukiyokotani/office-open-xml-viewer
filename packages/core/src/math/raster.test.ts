import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  drawMathJaxSvg,
  rasterizeMathSvg,
  sizeMathSvgForRaster,
} from './raster.js';
import { MAX_CANVAS_AREA } from '../canvas/clamp.js';
import { mathMLToSvg } from './engine.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('worker-safe math SVG rasterization', () => {
  it('pins MathJax SVGs to a deterministic high-resolution pixel size', () => {
    expect(sizeMathSvgForRaster(
      '<svg width="2ex" height="1ex" viewBox="0 -800 2000 1000"></svg>',
      2,
      1,
    )).toContain('<svg viewBox="0 -800 2000 1000" width="512" height="256">');
  });

  it('draws MathJax paths and rules through the same Canvas path in workers', async () => {
    const path = { rect: vi.fn(), moveTo: vi.fn(), lineTo: vi.fn() };
    const context = {
      save: vi.fn(), restore: vi.fn(), setTransform: vi.fn(), transform: vi.fn(),
      translate: vi.fn(), scale: vi.fn(), rotate: vi.fn(), fill: vi.fn(), stroke: vi.fn(),
      drawImage: vi.fn(),
      fillStyle: '', strokeStyle: '', lineWidth: 1, lineCap: 'butt', lineJoin: 'miter',
      globalAlpha: 1,
    } as unknown as OffscreenCanvasRenderingContext2D;
    class TestPath2D {
      constructor(readonly data?: string) {}
      rect = path.rect;
      moveTo = path.moveTo;
      lineTo = path.lineTo;
    }
    class TestOffscreenCanvas {
      readonly width: number;
      readonly height: number;
      constructor(width: number, height: number) { this.width = width; this.height = height; }
      getContext() { return context; }
    }
    vi.stubGlobal('Path2D', TestPath2D);
    vi.stubGlobal('OffscreenCanvas', TestOffscreenCanvas);

    const raster = await rasterizeMathSvg({
      svg: '<svg viewBox="0 -800 2000 1000"><g fill="currentColor" transform="scale(1,-1)"><path d="M0 0L2 2Z"/><rect x="1" y="2" width="3" height="4"/></g></svg>',
      widthEm: 2,
      ascentEm: 0.8,
      descentEm: 0.2,
    }, '#123456');

    expect(raster.widthPx).toBe(128);
    expect(raster.heightPx).toBe(64);
    expect(context.drawImage).toHaveBeenCalledTimes(2);
    expect(context.imageSmoothingEnabled).toBe(true);
    expect(context.imageSmoothingQuality).toBe('high');
    expect(context.setTransform).toHaveBeenCalledWith(0.256, 0, 0, 0.256, -0, 204.8);
    expect(context.scale).toHaveBeenCalledWith(1, -1);
    expect(context.fill).toHaveBeenCalledTimes(2);
    expect(path.rect).toHaveBeenCalledWith(1, 2, 3, 4);
    expect(context.fillStyle).toBe('#123456');
  });

  it('paints MathJax fallback text for an excluded font range', async () => {
    const output = await mathMLToSvg(
      '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>Ж</mi></math>',
    );
    const context = {
      save: vi.fn(), restore: vi.fn(), setTransform: vi.fn(), scale: vi.fn(),
      fillText: vi.fn(), fillStyle: '', strokeStyle: '', lineWidth: 1,
      lineCap: 'butt', lineJoin: 'miter', globalAlpha: 1, font: '',
    } as unknown as OffscreenCanvasRenderingContext2D;

    drawMathJaxSvg(context, output.svg, 600, 950);

    expect(context.fillText).toHaveBeenCalledWith('Ж', 0, 0);
    expect(context.font).toMatch(/italic .*px serif/);
  });

  it('rejects malformed view boxes before painting', () => {
    const context = { save: vi.fn() } as unknown as CanvasRenderingContext2D;
    expect(() => drawMathJaxSvg(context, '<svg></svg>', 10, 10)).toThrow(
      'must contain a finite viewBox',
    );
  });

  it('bounds pathological equation rasters to the shared Canvas pixel budget', async () => {
    const context = {
      save: vi.fn(), restore: vi.fn(), setTransform: vi.fn(), transform: vi.fn(),
      translate: vi.fn(), scale: vi.fn(), rotate: vi.fn(), fill: vi.fn(), stroke: vi.fn(),
      drawImage: vi.fn(),
      fillStyle: '', strokeStyle: '', lineWidth: 1, lineCap: 'butt', lineJoin: 'miter',
      globalAlpha: 1,
    } as unknown as OffscreenCanvasRenderingContext2D;
    class TestPath2D {}
    class TestOffscreenCanvas {
      constructor(readonly width: number, readonly height: number) {}
      getContext() { return context; }
    }
    vi.stubGlobal('Path2D', TestPath2D);
    vi.stubGlobal('OffscreenCanvas', TestOffscreenCanvas);

    const raster = await rasterizeMathSvg({
      svg: '<svg viewBox="0 0 1000 1000"><path d="M0 0L1 1Z"/></svg>',
      widthEm: 100_000,
      ascentEm: 50_000,
      descentEm: 50_000,
    }, '#000000');

    expect(raster.widthPx * raster.heightPx).toBeLessThanOrEqual(MAX_CANVAS_AREA);
    expect(raster.widthPx).toBe(raster.heightPx);
  });
});
