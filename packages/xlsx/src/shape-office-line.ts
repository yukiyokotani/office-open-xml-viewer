import { findReferenceFontMetrics } from '@silurus/ooxml-core';
import { officeOpenTypeAutoLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';
import type { OfficeFontFallbackRequest, OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import type { ShapeText, ShapeTextRun } from './types.js';

type TextRun = Extract<ShapeTextRun, { type: 'text' }>;

/** Excel evidence covers one natural DrawingML line with omitted <a:lnSpc>.
 * A body that can wrap is admitted only after the renderer measures one line. */
export function singleNaturalShapeRun(text: ShapeText): TextRun | undefined {
  if ((text.autoFit && text.autoFit !== 'none') || text.paragraphs.length !== 1) return undefined;
  const paragraph = text.paragraphs[0];
  if (paragraph.spaceLine != null || paragraph.runs.length !== 1) return undefined;
  const run = paragraph.runs[0];
  if (run.type !== 'text' || !run.text || run.text.includes('\n') || !run.fontFace?.trim()) return undefined;
  // The existing shape paint path selects a:latin. A distinct East Asian or
  // complex-script face would require script-run routing before its metrics
  // could own the complete line box.
  if (run.fontFaceEa && run.fontFaceEa.toLocaleLowerCase('en-US') !== run.fontFace.toLocaleLowerCase('en-US')) return undefined;
  if (run.fontFaceCs && run.fontFaceCs.toLocaleLowerCase('en-US') !== run.fontFace.toLocaleLowerCase('en-US')) return undefined;
  return run;
}

/** One key rule for workbook, worker, and synchronous shape paint. */
export function officeRequestKey(request: OfficeFontFallbackRequest): string {
  const family = request.family.trim().toLowerCase();
  const weight = request.weight ?? 400;
  const style = request.style ?? 'normal';
  return weight === 400 && style === 'normal' ? family : `${family}:${weight}:${style}`;
}

export function shapeOfficeRouteKey(run: TextRun): string {
  return officeRequestKey({ family: run.fontFace!, weight: run.bold ? 700 : 400,
    style: run.italic ? 'italic' : 'normal' });
}

/**
 * The static catalog is reference geometry, not proof of bytes behind local().
 * Admit it only after an exact style has loaded and every known profile for
 * that authored tuple agrees on the Office projection. Unknown OS/2 class or
 * conflicting Office/macOS versions leave the ordinary Canvas line height.
 */
export function shapeOfficeNaturalLineRatio(
  run: TextRun,
  route: OfficeFontFallbackRoute | undefined,
): number | undefined {
  if (!route || route.source !== 'local' || route.metric.synthesized) return undefined;
  const family = run.fontFace!.trim();
  if (route.requestedFamily.toLocaleLowerCase('en-US') !== family.toLocaleLowerCase('en-US')
    || route.weight !== (run.bold ? 700 : 400)
    || route.style !== (run.italic ? 'italic' : 'normal')) return undefined;
  const profiles = findReferenceFontMetrics(family, { weight: route.weight, style: route.style });
  if (profiles.length === 0) return undefined;
  let ratio: number | undefined;
  for (const profile of profiles) {
    if (profile.farEastCodePage == null) return undefined;
    const projected = officeOpenTypeAutoLineRatios({
      unitsPerEm: profile.unitsPerEm,
      hheaAscent: profile.hhea[0],
      hheaDescent: profile.hhea[1],
      hheaLineGap: profile.hhea[2],
      farEastCodePage: profile.farEastCodePage,
    })?.lineHeightRatio;
    if (projected == null || (ratio !== undefined && projected !== ratio)) return undefined;
    ratio = projected;
  }
  return ratio;
}
