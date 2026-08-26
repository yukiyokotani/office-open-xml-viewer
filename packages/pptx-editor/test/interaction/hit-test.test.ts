import { describe, expect, it } from 'vitest';

import type { MediaElement, PictureElement, Presentation } from '@maxgent/ooxml/pptx';

import {
  clientPointToSlidePoint,
  hitTestSlideElement,
  hitTestSlideShape,
  resolveElementSelection,
} from '../../src/interaction/hit-test';
import { deck, shape } from '../fixtures/presentation';

describe('PPTX editor shape hit testing', () => {
  it('maps browser client coordinates into slide EMUs independently of CSS scale', () => {
    const presentation = {
      ...deck([]),
      slideWidth: 1_000,
      slideHeight: 500,
    };
    const canvas = canvasWithRect({ left: 10, top: 20, width: 200, height: 100 });

    expect(clientPointToSlidePoint(canvas, presentation, {
      clientX: 110,
      clientY: 70,
    })).toEqual({ x: 500, y: 250 });
    expect(clientPointToSlidePoint(canvas, presentation, {
      clientX: 211,
      clientY: 70,
    })).toBeUndefined();
  });

  it('selects the topmost direct slide shape and returns a stable ElementRef', () => {
    const bottom = shape('7', 'bottom');
    const top = shape('8', 'top');
    const presentation = deck([bottom, top]);

    expect(hitTestSlideShape(presentation, 0, { x: 5, y: 5 })).toMatchObject({
      target: {
        origin: 'slide',
        slideId: 'ppt/slides/slide1.xml',
        elementId: '8',
      },
      presentationElementIndex: 1,
      element: top,
      isOfficeCliTargetable: true,
    });
  });

  it('stops at a topmost non-shape slide element instead of selecting a covered shape', () => {
    const bottom = shape('7', 'covered');
    const top: PictureElement = {
      type: 'picture',
      x: 0,
      y: 0,
      width: 10,
      height: 10,
      rotation: 0,
      flipH: false,
      flipV: false,
      imagePath: 'ppt/media/image1.png',
      mimeType: 'image/png',
      stroke: null,
    };
    const presentation = deck([bottom, top]);

    expect(hitTestSlideShape(presentation, 0, { x: 5, y: 5 })).toBeUndefined();
  });

  it('selects a topmost picture for element-level actions', () => {
    const picture: PictureElement = {
      type: 'picture',
      id: '8',
      x: 0,
      y: 0,
      width: 10,
      height: 10,
      rotation: 0,
      flipH: false,
      flipV: false,
      imagePath: 'ppt/media/image1.png',
      mimeType: 'image/png',
      stroke: null,
    };
    const presentation = deck([picture]);

    expect(hitTestSlideElement(presentation, 0, { x: 5, y: 5 })).toMatchObject({
      target: { elementId: '8' },
      element: picture,
      isOfficeCliTargetable: true,
    });
  });

  it('does not select media or elements covered by media', () => {
    const media: MediaElement = {
      type: 'media',
      id: '12',
      x: 0,
      y: 0,
      width: 10,
      height: 10,
      rotation: 0,
      flipH: false,
      flipV: false,
      mediaKind: 'video',
      posterPath: '',
      posterMimeType: '',
      mediaPath: 'ppt/media/video1.mp4',
      mimeType: 'video/mp4',
    };
    const presentation = deck([shape('7', 'covered'), media]);

    expect(hitTestSlideElement(presentation, 0, { x: 5, y: 5 })).toBeUndefined();
    expect(resolveElementSelection(presentation, {
      origin: 'slide',
      slideId: 'ppt/slides/slide1.xml',
      elementId: '12',
    })).toBeUndefined();
  });

  it('ignores inherited layout/master elements and non-shape elements', () => {
    const inherited = shape('9', 'layout');
    const direct = shape('7', 'direct');
    const presentation: Presentation = {
      ...deck([direct, inherited]),
      slides: [{
        ...deck([direct, inherited]).slides[0],
        elementSources: [
          { origin: 'slide' },
          { origin: 'layout' },
        ],
      }],
    };

    expect(hitTestSlideShape(presentation, 0, { x: 5, y: 5 })?.element).toBe(direct);
  });

  it('inverse-rotates the pointer before testing the shape bounds', () => {
    const rotated = shape('7', 'rotated', {
      x: 40,
      y: 40,
      width: 20,
      height: 10,
      rotation: 90,
    });
    const presentation = {
      ...deck([rotated]),
      slideWidth: 100,
      slideHeight: 100,
    };

    expect(hitTestSlideShape(presentation, 0, { x: 50, y: 54 })?.element).toBe(rotated);
    expect(hitTestSlideShape(presentation, 0, { x: 61, y: 45 })).toBeUndefined();
  });

  it('uses hit slop for zero-width shapes and marks positional ids as non-targetable', () => {
    const line = shape(undefined, '', {
      x: 50,
      y: 20,
      width: 0,
      height: 40,
    });
    const presentation = {
      ...deck([line]),
      slideWidth: 100,
      slideHeight: 100,
    };

    expect(hitTestSlideShape(
      presentation,
      0,
      { x: 54, y: 40 },
      { hitSlop: 5 },
    )).toMatchObject({
      target: { elementId: 'index:0' },
      isOfficeCliTargetable: false,
    });
  });
});

function canvasWithRect(
  rect: Pick<DOMRect, 'left' | 'top' | 'width' | 'height'>,
): Pick<HTMLCanvasElement, 'getBoundingClientRect'> {
  return {
    getBoundingClientRect: () => rect as DOMRect,
  };
}
