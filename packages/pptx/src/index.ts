export { PptxViewer, type PptxViewerOptions } from './viewer';
export { PptxPresentation, type LoadOptions, type RenderSlideOptions } from './presentation';
export { renderSlide, type RenderOptions, type TextRunInfo, type TextRunCallback } from './renderer';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
export type {
  Presentation,
  Slide,
  SlideElement,
  ShapeElement,
  PictureElement,
  Fill,
  SolidFill,
  NoFill,
  Stroke,
  TextBody,
  Paragraph,
  TextRun,
  TextRunData,
  LineBreak,
} from './types';
