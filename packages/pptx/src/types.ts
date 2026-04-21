// ===== Shared types re-exported from @silurus/ooxml-core =====
export type {
  PathCmd,
  Fill, SolidFill, NoFill, GradientFill, GradientStop,
  Shadow,
  Stroke,
  TextBody,
  SpaceLine,
  Bullet,
  TabStop,
  Paragraph,
  TextRun, TextRunData, LineBreak,
  RenderOptions,
  ChartModel, ChartSeries,
} from '@silurus/ooxml-core';

// ===== Presentation data model =====
// All positions and sizes are in EMUs (English Metric Units).
// 914400 EMU = 1 inch, 12700 EMU = 1 pt

import type { Fill, Stroke, TextBody, Shadow, PathCmd, ChartSeries } from '@silurus/ooxml-core';

export interface Presentation {
  slideWidth: number;
  slideHeight: number;
  slides: Slide[];
  /** Theme dk1 color (e.g. "383838"). Used as fallback text color when no explicit color is set. */
  defaultTextColor: string | null;
  /** Theme major (heading) font family name (e.g. "Aptos Display", "Nunito Sans"). Null if not set. */
  majorFont: string | null;
  /** Theme minor (body) font family name (e.g. "Aptos", "Nunito Sans"). Null if not set. */
  minorFont: string | null;
}

export interface Slide {
  index: number;
  /** 1-based slide number (index + 1); used to render slidenum fields */
  slideNumber: number;
  background: Fill | null;
  elements: SlideElement[];
}

export type SlideElement = ShapeElement | PictureElement | TableElement | ChartElement | MediaElement;

export interface MediaElement {
  type: 'media';
  x: number;
  y: number;
  width: number;
  height: number;
  /** "audio" or "video" */
  mediaKind: 'audio' | 'video';
  /** Poster image zip path (e.g. "ppt/media/image2.png"). Empty when no poster. */
  posterPath: string;
  /** Poster image MIME type (empty when no poster). */
  posterMimeType: string;
  /** Path inside the pptx zip (e.g. "ppt/media/media2.mp4"). Used by getMedia. */
  mediaPath: string;
  /** MIME type of the underlying media (e.g. "audio/mpeg", "video/mp4"). */
  mimeType: string;
}

export interface ShapeElement {
  type: 'shape';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Rotation in degrees, clockwise */
  rotation: number;
  /** Horizontal mirror (a:xfrm flipH) */
  flipH: boolean;
  /** Vertical mirror (a:xfrm flipV) */
  flipV: boolean;
  /** OOXML preset name or "custGeom" when custom paths are used */
  geometry: string;
  fill: Fill | null;
  stroke: Stroke | null;
  textBody: TextBody | null;
  /** Default text color from p:style > fontRef (hex). Used when run/para has no explicit color. */
  defaultTextColor: string | null;
  /** Custom geometry sub-paths (set only when geometry === "custGeom").
   *  Outer array: one entry per <a:path>; inner: path commands with coords in [0,1]. */
  custGeom: PathCmd[][] | null;
  /** First adjustment value from prstGeom avLst (e.g. trapezoid inset). Range 0–100000. */
  adj: number | null;
  /** Second adjustment value from prstGeom avLst (e.g. arrow head width). Range 0–100000. */
  adj2: number | null;
  /** Third adjustment value from prstGeom avLst (e.g. callout tip x). Range 0–100000. */
  adj3: number | null;
  /** Fourth adjustment value from prstGeom avLst (e.g. callout tip y). Range 0–100000. */
  adj4: number | null;
  /** Drop shadow from effectLst > outerShdw (null if not present). */
  shadow: Shadow | null;
}

export interface TableElement {
  type: 'table';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Column widths in EMU */
  cols: number[];
  rows: TableRow[];
}

export interface TableRow {
  /** Row height in EMU */
  height: number;
  cells: TableCell[];
}

export interface TableCell {
  textBody: TextBody | null;
  fill: Fill | null;
  borderL: Stroke | null;
  borderR: Stroke | null;
  borderT: Stroke | null;
  borderB: Stroke | null;
  /** Diagonal from top-left to bottom-right */
  diagonalTL?: Stroke | null;
  /** Diagonal from top-right to bottom-left */
  diagonalTR?: Stroke | null;
  /** Column span */
  gridSpan: number;
  /** Row span */
  rowSpan: number;
  /** Horizontal merge continuation */
  hMerge: boolean;
  /** Vertical merge continuation */
  vMerge: boolean;
}

/**
 * PPTX chart element. The Rust parser emits ChartModel fields flat at the
 * top level, alongside the element position (x/y/width/height in EMU).
 * Pass this straight to `renderChart` from `@silurus/ooxml-core`.
 */
export interface ChartElement {
  type: 'chart';
  x: number;
  y: number;
  width: number;
  height: number;
  chartType: string;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
  valMax: number | null;
  valMin: number | null;
  subtotalIndices: number[];
  showDataLabels: boolean;
  catAxisHidden: boolean;
  valAxisHidden: boolean;
  plotAreaBg: string | null;
  /** Outer chartSpace background (hex without '#'). null when noFill/absent. */
  chartBg: string | null;
  /** True when <c:legend> is declared; false suppresses the legend entirely. */
  showLegend: boolean;
  /** catAx crossBetween: "between" (default, 0.5-step padding) or "midCat". */
  catAxisCrossBetween: 'between' | 'midCat' | string;
  /** `<c:valAx><c:majorTickMark>`. "cross" (default) | "out" | "in" | "none". */
  valAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** `<c:catAx><c:majorTickMark>`. */
  catAxisMajorTickMark: 'cross' | 'out' | 'in' | 'none' | string;
  /** Title font size in OOXML hundredths of a point (1600 = 16pt). null = default. */
  titleFontSizeHpt: number | null;
  /** Title font color as a hex string without '#'. null = default/theme. */
  titleFontColor?: string | null;
  /** Title font family (`<a:latin typeface>`). null = default/theme. */
  titleFontFace?: string | null;
  /** `<c:catAx><c:txPr>` font size (hpt). null = proportional default. */
  catAxisFontSizeHpt: number | null;
  /** `<c:valAx><c:txPr>` font size (hpt). null = proportional default. */
  valAxisFontSizeHpt: number | null;
  /** `<c:dLbls><c:txPr>` font size (hpt) for data-point value labels. */
  dataLabelFontSizeHpt: number | null;
}

export interface PictureElement {
  type: 'picture';
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /** Data URL, e.g. "data:image/png;base64,..." */
  dataUrl: string;
  /** OOXML adj value (0–100000) for roundRect clip, null = plain rectangle */
  clipAdjust: number | null;
  /**
   * ECMA-376 a:srcRect — source image crop as fractions (0..1) of the source
   * width/height. Omitted when the image is not cropped.
   */
  srcRect?: { l?: number; t?: number; r?: number; b?: number };
  /** a:blip > a:alphaModFix@amt as 0..1. Undefined = fully opaque. */
  alpha?: number;
}

// ===== Worker message protocol =====

export type WorkerRequest =
  | { kind: 'init'; wasmUrl: string }
  | { kind: 'parse'; id: number; buffer: ArrayBuffer }
  | { kind: 'extractMedia'; id: number; path: string };

export type WorkerResponse =
  | { kind: 'ready' }
  | { kind: 'parsed'; id: number; presentation: Presentation }
  | { kind: 'mediaExtracted'; id: number; bytes: ArrayBuffer }
  | { kind: 'error'; id: number; message: string };
