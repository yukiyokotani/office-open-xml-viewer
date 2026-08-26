import type { Presentation, SlideElement } from '@maxgent/ooxml/pptx';

import {
  createElementRef,
  getElementSources,
  resolveElementRef,
} from '../adapters/pptx-json-adapter';
import { ELEMENT_ORIGINS } from '../domain/element-origin';
import type { ElementRef } from '../domain/mutation';
import type {
  ClientPoint,
  ElementHitTestOptions,
  PptxEditorElementSelection,
  PptxEditorSelectableElement,
  PptxEditorShapeSelection,
  SlidePoint,
} from './types';

export function clientPointToSlidePoint(
  canvas: Pick<HTMLCanvasElement, 'getBoundingClientRect'>,
  presentation: Presentation,
  point: ClientPoint,
): SlidePoint | undefined {
  const rect = canvas.getBoundingClientRect();
  if (
    !Number.isFinite(point.clientX)
    || !Number.isFinite(point.clientY)
    || !Number.isFinite(presentation.slideWidth)
    || !Number.isFinite(presentation.slideHeight)
    || presentation.slideWidth <= 0
    || presentation.slideHeight <= 0
    || rect.width <= 0
    || rect.height <= 0
  ) {
    return undefined;
  }

  const offsetX = point.clientX - rect.left;
  const offsetY = point.clientY - rect.top;
  if (offsetX < 0 || offsetY < 0 || offsetX > rect.width || offsetY > rect.height) {
    return undefined;
  }

  return Object.freeze({
    x: (offsetX / rect.width) * presentation.slideWidth,
    y: (offsetY / rect.height) * presentation.slideHeight,
  });
}

export function hitTestSlideShape(
  presentation: Presentation,
  slideIndex: number,
  point: SlidePoint,
  options: ElementHitTestOptions = {},
): PptxEditorShapeSelection | undefined {
  const selection = hitTestSlideElement(presentation, slideIndex, point, options);
  return selection?.element.type === 'shape'
    ? selection as PptxEditorShapeSelection
    : undefined;
}

export function hitTestSlideElement(
  presentation: Presentation,
  slideIndex: number,
  point: SlidePoint,
  options: ElementHitTestOptions = {},
): PptxEditorElementSelection | undefined {
  if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= presentation.slides.length) {
    return undefined;
  }
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return undefined;

  const slide = presentation.slides[slideIndex];
  const sources = getElementSources(slide);
  if (!sources) return undefined;
  const hitSlop = normalizedHitSlop(options.hitSlop);

  for (let index = slide.elements.length - 1; index >= 0; index -= 1) {
    const element = slide.elements[index];
    const source = sources[index];
    if (source.origin !== ELEMENT_ORIGINS.SLIDE) continue;
    if (!containsPoint(element, point, hitSlop)) continue;
    if (element.type === 'media') return undefined;
    return createElementSelection(
      createElementRef(slide, element, index),
      slideIndex,
      index,
      element,
    );
  }
  return undefined;
}

export function resolveShapeSelection(
  presentation: Presentation,
  target: ElementRef,
): PptxEditorShapeSelection | undefined {
  const selection = resolveElementSelection(presentation, target);
  return selection?.element.type === 'shape'
    ? selection as PptxEditorShapeSelection
    : undefined;
}

export function resolveElementSelection(
  presentation: Presentation,
  target: ElementRef,
): PptxEditorElementSelection | undefined {
  const resolved = resolveElementRef(presentation, target);
  if (
    !resolved
    || resolved.element.type === 'media'
    || resolved.source.origin !== ELEMENT_ORIGINS.SLIDE
  ) {
    return undefined;
  }
  return createElementSelection(
    target,
    resolved.slideIndex,
    resolved.presentationElementIndex,
    resolved.element,
  );
}

function createElementSelection<Element extends PptxEditorSelectableElement>(
  target: ElementRef,
  slideIndex: number,
  presentationElementIndex: number,
  element: Element,
): PptxEditorElementSelection<Element> {
  return Object.freeze({
    target,
    slideIndex,
    presentationElementIndex,
    element,
    isOfficeCliTargetable: /^\d+$/.test(target.elementId),
  });
}

function normalizedHitSlop(value: number | undefined): number {
  return value !== undefined && Number.isFinite(value) && value > 0 ? value : 0;
}

function containsPoint(
  element: SlideElement,
  point: SlidePoint,
  hitSlop: number,
): boolean {
  const values = [
    element.x,
    element.y,
    element.width,
    element.height,
    element.rotation,
  ];
  if (values.some((value) => !Number.isFinite(value))) return false;

  const halfWidth = Math.abs(element.width) / 2;
  const halfHeight = Math.abs(element.height) / 2;
  const centerX = element.x + element.width / 2;
  const centerY = element.y + element.height / 2;
  const dx = point.x - centerX;
  const dy = point.y - centerY;
  const angle = (-element.rotation * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  let localX = dx * cos - dy * sin;
  let localY = dx * sin + dy * cos;
  if (element.flipH) localX = -localX;
  if (element.flipV) localY = -localY;

  return Math.abs(localX) <= halfWidth + hitSlop
    && Math.abs(localY) <= halfHeight + hitSlop;
}
