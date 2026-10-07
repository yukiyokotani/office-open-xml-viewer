import { PT_TO_PX } from '@silurus/ooxml-core';
import { composeAffine, mapAffinePoint, scaleAffine } from './layout/affine.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { paragraphGeometryForPage } from './layout/text-index.js';
import type { DocumentLayout, LayoutRect, LayoutServices, Matrix2DData } from './layout/types.js';
import type { DocxStorySource } from './types.js';
import type { SelectedTextRunsForPageOptions, TextRunsForPageOptions } from './text-run-projection.js';

/** CSS page coordinates at the requested width, independently of DPR. */
export interface DocxLayoutPoint { readonly x: number; readonly y: number }
export interface DocxLayoutRect {
  readonly x: number; readonly y: number; readonly width: number; readonly height: number;
}
export interface DocxParagraphLineInfo {
  /** Half-open UTF-16 range in the source paragraph. */
  readonly range: Readonly<{ start: number; end: number }>;
  /** Axis-aligned envelope; corners preserve rotation/vertical writing geometry. */
  readonly bounds: DocxLayoutRect;
  readonly corners: readonly DocxLayoutPoint[];
  readonly baseline: DocxLayoutPoint;
  readonly markOnly: boolean;
}
/** A source paragraph occurrence on one page, without internal layout/paint objects. */
export interface DocxPageParagraphInfo {
  readonly source: Readonly<DocxStorySource>;
  readonly paragraphId?: string;
  readonly pageIndex: number;
  readonly lines: readonly DocxParagraphLineInfo[];
  /** Ancestor clip quadrilaterals in the same page coordinates as lines. */
  readonly clips: readonly (readonly DocxLayoutPoint[])[];
}

function point(matrix: Matrix2DData, xPt: number, yPt: number): DocxLayoutPoint {
  const mapped = mapAffinePoint(matrix, { xPt, yPt });
  return { x: mapped.xPt, y: mapped.yPt };
}
function corners(matrix: Matrix2DData, bounds: LayoutRect): DocxLayoutPoint[] {
  const { xPt, yPt, widthPt, heightPt } = bounds;
  return [point(matrix, xPt, yPt), point(matrix, xPt + widthPt, yPt),
    point(matrix, xPt + widthPt, yPt + heightPt), point(matrix, xPt, yPt + heightPt)];
}
function envelope(points: readonly DocxLayoutPoint[]): DocxLayoutRect {
  const x = Math.min(...points.map(point => point.x));
  const y = Math.min(...points.map(point => point.y));
  return { x, y, width: Math.max(...points.map(point => point.x)) - x,
    height: Math.max(...points.map(point => point.y)) - y };
}

export function paragraphsForPage(
  layout: DocumentLayout, pageIndex: number, options: TextRunsForPageOptions,
): DocxPageParagraphInfo[] {
  if (!Number.isFinite(options.scale) || options.scale <= 0)
    throw new RangeError(`Paragraph projection scale must be positive: ${options.scale}`);
  const scale = scaleAffine(options.scale);
  return paragraphGeometryForPage(layout, pageIndex).map(({ paragraph, pointToPage, clips }) => {
    const matrix = composeAffine(scale, pointToPage);
    const mark = paragraph.paragraphMark;
    const lines = paragraph.lines.length ? paragraph.lines.map(line => ({ line, markOnly: false }))
      : mark && !mark.hidden && mark.line ? [{ line: mark.line, markOnly: true }] : [];
    return {
      source: { story: paragraph.source.story, storyInstance: paragraph.source.storyInstance,
        path: [...paragraph.source.path] },
      ...(paragraph.paragraphId === undefined ? {} : { paragraphId: paragraph.paragraphId }),
      pageIndex,
      lines: lines.map(({ line, markOnly }) => {
        const points = corners(matrix, line.bounds);
        return { range: { ...line.range }, bounds: envelope(points), corners: points,
          baseline: point(matrix, line.bounds.xPt, line.baselinePt), markOnly };
      }),
      clips: clips.map(clip => corners(composeAffine(scale, clip.pointToPage), clip.bounds)),
    };
  });
}

/** Select the exact layout variant used by paint before projecting page-local geometry. */
export function paragraphsForSelectedPage(
  services: LayoutServices, pageIndex: number, options: SelectedTextRunsForPageOptions,
): DocxPageParagraphInfo[] {
  const selected = selectDocumentLayoutPage(services, {
    currentDate: options.currentDate, defaultCurrentDateMs: options.defaultCurrentDateMs,
    showTrackedChanges: options.showTrackedChanges,
  }, pageIndex);
  const scale = (options.width ?? selected.page.geometry.widthPt * PT_TO_PX) / selected.page.geometry.widthPt;
  return paragraphsForPage(selected.layout, pageIndex, { scale });
}
