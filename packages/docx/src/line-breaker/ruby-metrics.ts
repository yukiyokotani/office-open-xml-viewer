import type { MeasurementTextContext } from '../layout/measurement-capabilities.js';
import { calcEffectiveFontPx, independentTextShapeRequest } from '../layout/text.js';
import { type LayoutTextSeg } from './model.js';
import { buildFont } from './font-routes.js';

/** ECMA-376 §17.3.3.12 defines `hpsRaise` as the “distance [...] between the
 * phonetic guide base text and the phonetic guide text.” The absent case needs
 * selected-face ink and is therefore resolved by retainedRubyAscentReservePx. */
export function rubyAscentReservePx(
  rubySizePt: number,
  hpsRaisePt: number | undefined,
  scale: number,
  segment?: LayoutTextSeg,
  ctx?: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
): number {
  if (hpsRaisePt != null) return hpsRaisePt * scale;
  if (!segment?.ruby || !ctx) {
    throw new Error(`Ruby at ${rubySizePt}pt without hpsRaise requires retained base and guide ink`);
  }
  if (segment.textLayoutService && segment.textShapeRequest) {
    const base = segment.textLayoutService.shape({
      ...segment.textShapeRequest,
      fontSizePt: calcEffectiveFontPx(segment, scale),
      measure: true,
      clusterGeometry: false,
    });
    const guide = segment.textLayoutService.shape({
      ...independentTextShapeRequest(segment.textShapeRequest, segment.ruby.text),
      fontSizePt: segment.ruby.fontSizePt * scale,
      measure: true,
      clusterGeometry: false,
    });
    if (base.inkBounds && guide.inkBounds) {
      return base.inkBounds.ascentPt + guide.inkBounds.descentPt;
    }
  }
  // Isolated line-layout callers may not carry a service snapshot. Canvas
  // actual ink under the same selected route is still authoritative geometry;
  // retain it here instead of restoring the former font-size ratio.
  const previousFont = ctx.font;
  try {
    ctx.font = buildFont(
      segment.bold, segment.italic, calcEffectiveFontPx(segment, scale),
      segment.fontFamily, fontFamilyClasses, segment.fontRoute,
    );
    const base = ctx.measureText(segment.text);
    ctx.font = buildFont(
      segment.bold, segment.italic, rubySizePt * scale,
      segment.fontFamily, fontFamilyClasses, segment.fontRoute,
    );
    const guide = ctx.measureText(segment.ruby.text);
    if (
      Number.isFinite(base.actualBoundingBoxAscent)
      && Number.isFinite(guide.actualBoundingBoxDescent)
    ) {
      return base.actualBoundingBoxAscent + guide.actualBoundingBoxDescent;
    }
  } finally {
    ctx.font = previousFont;
  }
  throw new Error('Ruby without hpsRaise requires retained base and guide ink');
}
