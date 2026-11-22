import type { MeasurementTextContext } from '../layout/measurement-capabilities.js';
import { calcEffectiveFontPx, independentTextShapeRequest } from '../layout/text.js';
import { type LayoutTextSeg } from './model.js';
import { buildFont } from './font-routes.js';
import { wordRubyBaseOffset } from '../layout/line-compatibility.js';

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

/** Natural advance of `text` shaped through `segment`'s request (its fonts by
 * slot) at `sizePt` × `scale`, or Canvas when no shaping service is present. */
function shapedAdvancePx(
  segment: LayoutTextSeg,
  text: string,
  sizePt: number,
  scale: number,
  ctx: MeasurementTextContext,
  fontFamilyClasses: Record<string, string>,
): number {
  if (segment.textLayoutService && segment.textShapeRequest) {
    return segment.textLayoutService.shape({
      ...independentTextShapeRequest(segment.textShapeRequest, text),
      fontSizePt: sizePt * scale,
      measure: true,
      clusterGeometry: false,
    }).advancePt;
  }
  const previousFont = ctx.font;
  try {
    ctx.font = buildFont(segment.bold, segment.italic, sizePt * scale,
      segment.fontFamily, fontFamilyClasses, segment.fontRoute);
    return ctx.measureText(text).width;
  } finally {
    ctx.font = previousFont;
  }
}

/** WORD_RUBY_ALIGN_PLACEMENT: the cell of a ruby run with an authored
 * alignment — the guide's and the whole base's natural advances, the space a
 * longer guide adds and the base's offset inside the widened cell. Every
 * piece of a split base measures the same complete texts. */
export function rubyCellGeometryPx(
  segment: LayoutTextSeg,
  ruby: NonNullable<LayoutTextSeg['ruby']>,
  scale: number,
  ctx: MeasurementTextContext,
  fontFamilyClasses: Record<string, string> = {},
  /** The laid-out advance of an unsplit base, which already holds its
   * spacing and grid allocation. */
  ownBasePx?: number,
): Readonly<{ basePx: number; guidePx: number; overhangPx: number; baseOffsetPx: number }> {
  const guidePx = shapedAdvancePx(segment, ruby.text, ruby.fontSizePt, scale, ctx, fontFamilyClasses);
  const basePx = ownBasePx
    ?? shapedAdvancePx(segment, ruby.baseText ?? segment.text, segment.fontSize, scale, ctx, fontFamilyClasses);
  const overhangPx = ruby.align ? Math.max(0, guidePx - basePx) : 0;
  return { basePx, guidePx, overhangPx, baseOffsetPx: wordRubyBaseOffset(ruby.align, overhangPx) };
}
