import { describe, expect, it } from 'vitest';
import { PptxPresentation } from './presentation';
import type { Presentation, ShapeElement } from './types';

function shape(): ShapeElement {
  return {
    type: 'shape',
    id: '7',
    x: 0,
    y: 0,
    width: 100,
    height: 50,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    fill: null,
    stroke: null,
    textBody: null,
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
  };
}

function makePresentation(mode: 'main' | 'worker' = 'main') {
  const target = shape();
  const model: Presentation = {
    slideWidth: 1000,
    slideHeight: 750,
    slides: [{
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [target],
    }],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
  const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
  instance._mode = mode;
  instance._presentation = mode === 'main' ? model : null;
  instance._meta = mode === 'worker'
    ? { slideCount: 1, slideWidth: 1000, slideHeight: 750 }
    : null;
  return {
    presentation: instance as unknown as PptxPresentation,
    target,
  };
}

describe('PptxPresentation.hitTestShape', () => {
  it('returns a detached shape from the current main-thread model', () => {
    const { presentation, target } = makePresentation();

    const hit = presentation.hitTestShape(0, { x: 20, y: 20 });

    expect(hit).toMatchObject({
      slideIndex: 0,
      shapeId: '7',
      point: { x: 20, y: 20 },
    });
    expect(hit?.shape).toEqual(target);
    expect(hit?.shape).not.toBe(target);
  });

  it('rejects an invalid slide index', () => {
    const { presentation } = makePresentation();
    expect(() => presentation.hitTestShape(1, { x: 0, y: 0 })).toThrow(
      /slide index 1 out of range/i,
    );
  });

  it('requires main mode because worker mode does not retain the shape model', () => {
    const { presentation } = makePresentation('worker');
    expect(() => presentation.hitTestShape(0, { x: 0, y: 0 })).toThrow(/mode.*main/i);
  });
});
