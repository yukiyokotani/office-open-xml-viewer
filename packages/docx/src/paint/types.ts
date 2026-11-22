import type {
  PaintResourceDescriptorKind,
  LayoutRect,
  PaintResourceKind,
  Matrix2DData,
} from '../layout/types.js';
import type { ResolvedPaintResource } from './resource-session.js';
import type { TextAdvanceLookup } from '@silurus/ooxml-core';

export interface PaintPageOptions {
  readonly scale: number;
  readonly dpr: number;
}

export interface PaintCanvas2D {
  globalAlpha: number;
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  direction: CanvasDirection;
  letterSpacing: string;
  fontKerning: CanvasFontKerning;
  fillRect(x: number, y: number, width: number, height: number): void;
  /** Native Canvas pattern source; optional on test recording contexts. */
  createPattern?(image: CanvasImageSource, repetition: string | null): CanvasPattern | null;
  strokeRect(x: number, y: number, width: number, height: number): void;
  setLineDash(segments: number[]): void;
  fillText(text: string, x: number, y: number): void;
  translate(x: number, y: number): void;
  rotate(angle: number): void;
  scale(x: number, y: number): void;
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  drawImage(image: CanvasImageSource, ...coordinates: number[]): void;
  save(): void;
  restore(): void;
  beginPath(): void;
  rect(x: number, y: number, width: number, height: number): void;
  clip(): void;
  moveTo(x: number, y: number): void;
  lineTo(x: number, y: number): void;
  stroke(): void;
  fill(): void;
}

export type CanvasPaintResourceHandler<K extends PaintResourceDescriptorKind> = (
  resource: ResolvedPaintResource<K>,
  bounds: LayoutRect,
  ctx: PaintCanvas2D,
  coverageBounds?: LayoutRect,
) => void;

export type CanvasPaintResourceHandlers = Readonly<{
  [K in PaintResourceDescriptorKind]: CanvasPaintResourceHandler<K>;
}>;

export interface CanvasPaintResourcePainter {
  paint(
    resourceKey: string,
    kind: PaintResourceKind,
    bounds: LayoutRect,
    ctx: PaintCanvas2D,
    coverageBounds?: LayoutRect,
  ): void;
}

export interface CanvasPaintContext {
  readonly ctx: PaintCanvas2D;
  /** Painting already happens inside a line's retained frame (layout/types.ts
   * ParagraphLayout.lineFrame): nested stories, such as the text of an inline
   * shape, must not apply their own line frame again. */
  readonly lineFramed?: true;
  readonly scale: number;
  readonly dpr: number;
  /** Page-to-device frame for DrawingML pattern fills. Chart and VML hosts
   * keep their own fill-coordinate policy. */
  readonly patternPageToDevice?: DOMMatrix2DInit;
  /** Document-level fallback retained independently from any enclosing
   * DrawingML text-box `fontRef` color override. */
  readonly documentDefaultTextColor?: string;
  readonly defaultTextColor?: string;
  readonly resources: CanvasPaintResourcePainter;
  /**
   * Affine map from the current retained point-space to final logical CSS pixels
   * (device-pixel scaling is represented separately by `dpr`). It mirrors every
   * production Canvas transform that can affect point geometry.
   */
  readonly pointToCss?: Matrix2DData;
  readonly layoutTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  readonly textBoxVerticalMode?: 'vert' | 'vert270' | 'eaVert' | 'mongolianVert';
  /** The page paint plan owns anchored drawings for this retained root. */
  readonly omitAnchoredDrawings?: boolean;
  /** WORD_CELL_CONTENT_CLIP: the cell content clip, in the coordinates of
   * the paragraph being painted, for a cell paragraph whose VML pictures or
   * anchored drawings paint outside it. Never inherited by nested stories. */
  readonly cellContentClip?: LayoutRect;
  /** Layout measurement's reference-advance decisions (core
   * reference-advances.ts): a retained shorthand layout measured at the
   * reference em is drawn there under a linear scale, so the shaped glyph
   * advances are the retained ones. Read-only; paint never probes. */
  readonly textAdvances?: TextAdvanceLookup;
}
