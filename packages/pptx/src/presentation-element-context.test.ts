import { describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import type { ShapeElement, Slide } from './types.js';

function shape(): ShapeElement {
  return {
    type: 'shape', id: '7', name: 'Title', x: 0, y: 0, width: 100, height: 50,
    rotation: 0, flipH: false, flipV: false, geometry: 'rect', fill: null,
    stroke: null, textBody: null, defaultTextColor: null, custGeom: null,
    adj: null, adj2: null, adj3: null, adj4: null, adj5: null, adj6: null,
    adj7: null, adj8: null, shadow: null,
  };
}

function slide(): Slide {
  return {
    index: 0, slideNumber: 1, background: null, elements: [shape()],
    elementSources: [{ origin: 'slide' }],
  };
}

function presentation(mode: 'main' | 'worker') {
  const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
  instance._mode = mode;
  instance._preflight = { slideCount: 1 };
  instance._resourceFailure = null;
  return instance;
}

describe('PptxPresentation.getElementContextAt', () => {
  it('uses the bounded slide repository in main mode', async () => {
    const instance = presentation('main');
    instance._slides = { withSlide: vi.fn(async (_index, consume) => consume(slide())) };

    await expect((instance as unknown as PptxPresentation).getElementContextAt(
      0, { x: 10, y: 10 },
    )).resolves.toMatchObject({ kind: 'element', shapeId: '7', origin: 'slide' });
  });

  it('uses the worker-owned slide model in worker mode', async () => {
    const instance = presentation('worker');
    const context = {
      format: 'pptx' as const, kind: 'element' as const, slideIndex: 0, elementIndex: 0,
      origin: 'slide' as const, elementType: 'shape' as const, point: { x: 10, y: 10 },
      bounds: { x: 0, y: 0, width: 100, height: 50, rotation: 0, flipH: false, flipV: false },
      shapeId: '7', geometry: 'rect', truncated: false,
    };
    const request = vi.fn(async (build: (id: number) => unknown) => {
      expect(build(12)).toMatchObject({ kind: 'hitTestElement', id: 12, slideIndex: 0 });
      return { kind: 'elementHit', id: 12, context };
    });
    instance._bridge = { request };

    await expect((instance as unknown as PptxPresentation).getElementContextAt(
      0, { x: 10, y: 10 },
    )).resolves.toEqual(context);
  });
});

describe('PptxPresentation.getElementBoundsByIds', () => {
  it('resolves authored ids without hit-testing in main mode', async () => {
    const instance = presentation('main');
    instance._slides = { withSlide: vi.fn(async (_index, consume) => consume(slide())) };

    await expect((instance as unknown as PptxPresentation).getElementBoundsByIds(
      0, ['7'],
    )).resolves.toEqual([{
      elementId: '7', elementIndex: 0, origin: 'slide', elementType: 'shape',
      bounds: { x: 0, y: 0, width: 100, height: 50, rotation: 0, flipH: false, flipV: false },
    }]);
  });

  it('prefers the slide-authored id when composed master content reuses it', async () => {
    const instance = presentation('main');
    const master = { ...shape(), x: 1, width: 10 };
    const authored = { ...shape(), x: 40, width: 60 };
    instance._slides = { withSlide: vi.fn(async (_index, consume) => consume({
      ...slide(),
      elements: [master, authored],
      elementSources: [{ origin: 'master' }, { origin: 'slide' }],
    })) };

    await expect((instance as unknown as PptxPresentation).getElementBoundsByIds(
      0, ['7'],
    )).resolves.toMatchObject([{
      elementId: '7', elementIndex: 1, origin: 'slide',
      bounds: { x: 40, width: 60 },
    }]);
  });

  it('uses one worker request for every requested id', async () => {
    const instance = presentation('worker');
    const bounds = [{
      elementId: '7', elementIndex: 0, origin: 'slide' as const, elementType: 'shape' as const,
      bounds: { x: 0, y: 0, width: 100, height: 50, rotation: 0, flipH: false, flipV: false },
    }];
    instance._bridge = { request: vi.fn(async (build: (id: number) => unknown) => {
      expect(build(13)).toEqual({
        kind: 'resolveElementBounds', id: 13, slideIndex: 0, elementIds: ['7', '8'],
      });
      return { kind: 'elementBoundsResolved', id: 13, bounds };
    }) };

    await expect((instance as unknown as PptxPresentation).getElementBoundsByIds(
      0, ['7', '8'],
    )).resolves.toEqual(bounds);
  });
});
