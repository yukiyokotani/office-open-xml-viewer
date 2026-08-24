import { canvasFontString, PT_TO_PX } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer.js';
import {
  composeAffine,
  mapAffinePoint,
  scaleAffine,
} from './layout/affine.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import type { TextRunGeometry } from './layout/text-index.js';
import type { DocumentLayout, LayoutServices, Matrix2DData } from './layout/types.js';
import { cssTransformFor } from './paint/affine.js';

export interface TextRunsForPageOptions {
  readonly scale: number;
}

export interface SelectedTextRunsForPageOptions {
  readonly defaultCurrentDateMs: number;
  readonly currentDate?: Date | number;
  readonly width?: number;
}

function projectTextRun(
  geometry: TextRunGeometry,
  pointToCss: Matrix2DData,
): DocxTextRunInfo {
  const { placement } = geometry;
  const origin = mapAffinePoint(pointToCss, placement.bounds);
  const inlineScale = Math.hypot(pointToCss.a, pointToCss.b);
  const blockScale = Math.hypot(pointToCss.c, pointToCss.d);
  const transform = cssTransformFor(pointToCss);
  const letterSpacingPt = placement.paintOps[0]?.letterSpacingPt ?? 0;
  return {
    source: {
      story: geometry.source.story,
      storyInstance: geometry.source.storyInstance,
      path: [...geometry.source.path],
    },
    ...(geometry.paragraphId !== undefined
      ? { paragraphId: geometry.paragraphId }
      : {}),
    ...(placement.sourceRunIndex !== undefined
      ? { sourceRunIndex: placement.sourceRunIndex }
      : {}),
    text: placement.text,
    x: origin.xPt,
    y: origin.yPt,
    w: placement.bounds.widthPt * inlineScale,
    h: placement.bounds.heightPt * blockScale,
    fontSize: placement.fontSizePt * blockScale,
    font: canvasFontString(
      placement.fontRoute,
      placement.fontSizePt * blockScale,
      placement.fontWeight,
      placement.fontStyle,
    ),
    ...(letterSpacingPt !== 0
      ? { letterSpacingPx: letterSpacingPt * inlineScale }
      : {}),
    ...(transform ? { transform } : {}),
    ...(placement.hyperlink ? { hyperlink: placement.hyperlink } : {}),
    ...(placement.tateChuYoko ? { eastAsianVert: true } : {}),
  };
}

export function textRunsForPage(
  layout: DocumentLayout,
  pageIndex: number,
  options: TextRunsForPageOptions,
): DocxTextRunInfo[] {
  if (!Number.isFinite(options.scale) || options.scale <= 0) {
    throw new RangeError(`Text projection scale must be positive: ${options.scale}`);
  }
  const displayScale = scaleAffine(options.scale);
  return textRunGeometryForPage(layout, pageIndex).map((geometry) => (
    projectTextRun(
      geometry,
      composeAffine(displayScale, geometry.pointToPage),
    )
  ));
}

/** Select the same keyed layout variant as paint, then project its retained text. */
export function textRunsForSelectedPage(
  services: LayoutServices,
  pageIndex: number,
  options: SelectedTextRunsForPageOptions,
): DocxTextRunInfo[] {
  const selected = selectDocumentLayoutPage(services, {
    currentDate: options.currentDate,
    defaultCurrentDateMs: options.defaultCurrentDateMs,
  }, pageIndex);
  const scale = (
    options.width ?? selected.page.geometry.widthPt * PT_TO_PX
  ) / selected.page.geometry.widthPt;
  return textRunsForPage(selected.layout, pageIndex, { scale });
}
