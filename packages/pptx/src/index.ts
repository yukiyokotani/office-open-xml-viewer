export {
  PptxViewer,
  type PptxViewerOptions,
  type PptxShapeClickEvent,
  type HiddenSlideMode,
} from './viewer';
export { PptxScrollViewer, type PptxScrollViewerOptions } from './scroll-viewer';
export {
  PptxPresentation,
  type LoadOptions,
  type PptxShapeChangesResult,
  type RenderSlideOptions,
  type RenderSlideToBitmapOptions,
} from './presentation';
export {
  PptxShapeChangeType,
  type PptxShapeChange,
  type PptxShapeProperties,
  type PptxOptionalShapeProperty,
  type PptxApplyShapeChangesRequest,
} from './shape-changes';
export {
  type PptxSlidePoint,
  type PptxShapeHit,
  type PptxShapeHitTestOptions,
} from './shape-hit-test';
export { renderSlide, type RenderOptions, type PptxTextRunInfo, type TextRunCallback } from './renderer';
export { buildPptxTextLayer } from './text-layer';
// IX2 find-in-document: the highlight overlay builder + the pptx match-location
// shape. `FindMatch` / `FindMatchesOptions` come from core (shared across formats).
export {
  buildPptxHighlightLayer,
  type PptxHighlightMatch,
  type PptxHighlightColors,
} from './find-highlight-layer';
export type { PptxMatchLocation } from './find';
export type { FindMatch, FindMatchesOptions } from '@silurus/ooxml-core';
export type { PresentationHandle } from './presentation-handle';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
// IX1 — the shared hyperlink target shape surfaced by `PptxViewerOptions.
// onHyperlinkClick`, `PptxTextRunInfo.hyperlink`, and the 5th arg of
// `buildPptxTextLayer`, plus the default "open in a new tab, sanitised" helper.
export { type HyperlinkTarget, openExternalHyperlink } from '@silurus/ooxml-core';
// Typed load-time error surfaced by PptxPresentation.load (e.g. a
// password-protected or legacy-binary .ppt file). Re-exported so
// `@silurus/ooxml/pptx` consumers can narrow on `err.code`.
export { OoxmlError, type OoxmlErrorCode } from '@silurus/ooxml-core';
export type {
  Presentation,
  Slide,
  // Reachable via Slide.comments — exported so consumers reading legacy slide
  // comments have a name for the element type.
  PptxComment,
  SlideElement,
  ShapeElement,
  PictureElement,
  // SlideElement union members — reachable via Slide.elements; exported so a
  // consumer that narrows on `el.type` has a name for every variant.
  TableElement,
  TableRow,
  TableCell,
  ChartElement,
  MediaElement,
  TextRect,
  // 3D scene types (reachable via ShapeElement.scene3d / .sp3d).
  Scene3d,
  Camera3d,
  Rot3d,
  LightRig,
  Sp3d,
  Bevel3d,
  // Fill / stroke variants (reachable via ShapeElement.fill / .stroke etc.).
  Fill,
  SolidFill,
  NoFill,
  GradientFill,
  GradientStop,
  ImageFill,
  FillRect,
  TileInfo,
  Stroke,
  // Effect types (reachable via ShapeElement.shadow / .glow / …).
  Shadow,
  Glow,
  SoftEdge,
  Reflection,
  // Geometry + text run sub-types (reachable via ShapeElement / TextBody).
  PathCmd,
  Bullet,
  // Picture-bullet variant (`<a:buBlip>`, §21.1.2.4.2) — part of the PPTX
  // `Bullet` union; exported so consumers narrowing on `bullet.type === 'blip'`
  // have a name for the shape.
  BlipBullet,
  SpaceLine,
  TabStop,
  TextBody,
  Paragraph,
  TextRun,
  TextRunData,
  LineBreak,
  // Chart model (reachable via ChartElement.series and renderChart).
  ChartModel,
  ChartSeries,
  // Hidden-slide dimming options (reachable via RenderSlideOptions.dim /
  // RenderSlideToBitmapOptions.dim) — the translucent overlay mechanism.
  DimOptions,
} from './types';
