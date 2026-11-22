import { canvasFontString, PT_TO_PX, referenceTextFont } from '@silurus/ooxml-core';
import type { TextAdvanceLookup } from '@silurus/ooxml-core';
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
  /** Layout measurement's reference-advance decisions (core
   * reference-advances.ts). A run whose shorthand layout planned at the
   * reference em is projected with `referenceFont`; absent, runs keep the
   * plain `font` contract. */
  readonly textAdvances?: TextAdvanceLookup;
}

export interface SelectedTextRunsForPageOptions {
  readonly defaultCurrentDateMs: number;
  readonly currentDate?: Date | number;
  /** ECMA-376 §17.13.5 tracked-change view — selects the same keyed layout
   *  variant as paint (see RenderPageOptions.showTrackedChanges). */
  readonly showTrackedChanges?: boolean;
  readonly width?: number;
  readonly textAdvances?: TextAdvanceLookup;
}

function projectTextRun(
  geometry: TextRunGeometry,
  pointToCss: Matrix2DData,
  textAdvances: TextAdvanceLookup | undefined,
  displayScale: Matrix2DData,
): DocxTextRunInfo {
  const { placement } = geometry;
  const origin = mapAffinePoint(pointToCss, placement.bounds);
  const highlightOrigin = placement.highlightBounds
    ? mapAffinePoint(pointToCss, placement.highlightBounds)
    : undefined;
  const inlineScale = Math.hypot(pointToCss.a, pointToCss.b);
  const blockScale = Math.hypot(pointToCss.c, pointToCss.d);
  const transform = cssTransformFor(pointToCss);
  // Every paint clip, placed like highlightBounds with its own orientation.
  const clips = geometry.clips.map((paintClip) => {
    const clipToCss = composeAffine(displayScale, paintClip.pointToPage);
    const clipCorner = mapAffinePoint(clipToCss, paintClip.bounds);
    const clipTransform = cssTransformFor(clipToCss);
    const clipInline = Math.hypot(clipToCss.a, clipToCss.b);
    const clipBlock = Math.hypot(clipToCss.c, clipToCss.d);
    return Object.freeze({
      x: clipCorner.xPt,
      y: clipCorner.yPt,
      width: paintClip.bounds.widthPt * clipInline,
      height: paintClip.bounds.heightPt * clipBlock,
      ...(clipTransform ? { transform: clipTransform } : {}),
      // A union clip's rectangles, as offsets along its axes from its corner.
      ...(paintClip.rects ? {
        rects: Object.freeze(paintClip.rects.map((rect) => Object.freeze({
          x: (rect.xPt - paintClip.bounds.xPt) * clipInline,
          y: (rect.yPt - paintClip.bounds.yPt) * clipBlock,
          width: rect.widthPt * clipInline,
          height: rect.heightPt * clipBlock,
        }))),
      } : {}),
    });
  });
  const letterSpacingPt = placement.letterSpacingPt;
  const font = canvasFontString(
    placement.fontRoute,
    placement.fontSizePt * blockScale,
    placement.fontWeight,
    placement.fontStyle,
  );
  // A run layout measured and paint drew at the reference em (core
  // reference-advances.ts) is laid out by overlays at that em too: the host
  // grid-fits this face, so the browser's own layout of `font` would not
  // follow the drawn advances. Other runs keep the plain `font` contract.
  const reference = textAdvances?.planned(canvasFontString(
    placement.fontRoute,
    placement.fontSizePt,
    placement.fontWeight,
    placement.fontStyle,
  )) ? referenceTextFont(font) : null;
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
    direction: placement.direction,
    // ECMA-376 §17.3.3.29: the conditional glyph is a layout opportunity,
    // not an added source character. Preserve its owner/geometry for callers
    // while copy/find receive logical text. Ordinary '-' remains unchanged.
    ...(placement.optionalHyphenGlyph ? { optionalHyphenGlyph: true as const } : {}),
    text: placement.optionalHyphenGlyph ? '' : placement.text,
    x: origin.xPt,
    y: origin.yPt,
    w: placement.bounds.widthPt * inlineScale,
    ...(placement.trailingSpaceCompressionPt !== undefined ? {
      trailingSpaceCompressionPx: placement.trailingSpaceCompressionPt * inlineScale,
    } : {}),
    h: placement.bounds.heightPt * blockScale,
    ...(placement.highlightBounds && highlightOrigin ? {
      highlightBounds: Object.freeze({
        x: highlightOrigin.xPt,
        y: highlightOrigin.yPt,
        width: placement.highlightBounds.widthPt * inlineScale,
        height: placement.highlightBounds.heightPt * blockScale,
      }),
    } : {}),
    fontSize: placement.fontSizePt * blockScale,
    font,
    ...(reference ? { referenceFont: reference } : {}),
    ...(letterSpacingPt !== 0
      ? { letterSpacingPx: letterSpacingPt * inlineScale }
      : {}),
    ...(placement.optionalHyphenGlyph ? {} : {
      clusterBoxes: geometry.clusterBoxes.map((box) => ({
        start: box.start, end: box.end, x: box.xPt * inlineScale, width: box.widthPt * inlineScale, rtl: box.rtl,
      })),
    }),
    ...(transform ? { transform } : {}),
    ...(placement.hyperlink ? { hyperlink: placement.hyperlink } : {}),
    ...(placement.tateChuYoko ? { eastAsianVert: true } : {}),
    ...(clips.length > 0 ? { clips: Object.freeze(clips) } : {}),
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
      options.textAdvances,
      displayScale,
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
    showTrackedChanges: options.showTrackedChanges,
  }, pageIndex);
  const scale = (
    options.width ?? selected.page.geometry.widthPt * PT_TO_PX
  ) / selected.page.geometry.widthPt;
  return textRunsForPage(selected.layout, pageIndex, {
    scale,
    ...(options.textAdvances ? { textAdvances: options.textAdvances } : {}),
  });
}
