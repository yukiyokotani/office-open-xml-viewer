import { describe, expect, it } from 'vitest';
import {
  hitTestShapeElement,
  hitTestSlideShape,
  type PptxSlidePoint,
} from './shape-hit-test';
import type { ShapeElement, Slide } from './types';

function shape(
  id: string | undefined,
  overrides: Partial<ShapeElement> = {},
): ShapeElement {
  return {
    type: 'shape',
    id,
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
    ...overrides,
  };
}

function slide(elements: Slide['elements']): Slide {
  return {
    index: 0,
    slideNumber: 1,
    background: null,
    elements,
  };
}

describe('PPTX shape hit testing', () => {
  it('hits an ordinary shape by its bounding frame', () => {
    expect(hitTestShapeElement(shape('1'), { x: 50, y: 25 })).toBe(true);
    expect(hitTestShapeElement(shape('1'), { x: 101, y: 25 })).toBe(false);
  });

  it('undoes rotation before testing the local frame', () => {
    const rotated = shape('1', { width: 100, height: 20, rotation: 90 });
    expect(hitTestShapeElement(rotated, { x: 50, y: 50 })).toBe(true);
    expect(hitTestShapeElement(rotated, { x: 90, y: 10 })).toBe(false);
  });

  it('uses a configurable tolerance for line shapes', () => {
    const line = shape('1', {
      width: 100,
      height: 100,
      geometry: 'line',
    });
    expect(hitTestShapeElement(line, { x: 50, y: 54 }, 5)).toBe(true);
    expect(hitTestShapeElement(line, { x: 50, y: 60 }, 5)).toBe(false);
  });

  it('returns the topmost identified shape and a detached snapshot', () => {
    const bottom = shape('bottom');
    const top = shape('top');
    const point: PptxSlidePoint = { x: 20, y: 20 };

    const hit = hitTestSlideShape(3, slide([bottom, top]), point);

    expect(hit).toMatchObject({
      slideIndex: 3,
      shapeId: 'top',
      point,
    });
    expect(hit?.shape).toEqual(top);
    expect(hit?.shape).not.toBe(top);
    expect(hit?.point).not.toBe(point);
  });

  it('skips parser-synthesized shapes without an id', () => {
    expect(hitTestSlideShape(0, slide([shape(undefined)]), { x: 20, y: 20 })).toBeNull();
  });

  it('returns null for non-finite or empty-space points', () => {
    const target = slide([shape('1')]);
    expect(hitTestSlideShape(0, target, { x: Number.NaN, y: 20 })).toBeNull();
    expect(hitTestSlideShape(0, target, { x: 200, y: 200 })).toBeNull();
  });
});
