import type { TextLayoutService, TextShapeRequest, TextShapeResult } from '../layout/text.js';
import { wordArabicTextKerned } from '../layout/line-compatibility.js';
import type { LayoutTextSeg } from './model.js';

export function hasNativeUnitReceipt(segment: Pick<LayoutTextSeg, 'text' | 'textShapeRequest' | 'nativeLevelUnit'>): boolean {
  const unit = segment.nativeLevelUnit, request = segment.textShapeRequest, scope = request?.substituteContext;
  return !!unit && request?.baseDirection === unit.baseDirection && scope?.text === unit.sourceText
    && scope.offset >= unit.start && scope.offset + segment.text.length <= unit.end
    && scope.text.slice(scope.offset, scope.offset + segment.text.length) === segment.text;
}

export function sameNativeLevelUnit(left: LayoutTextSeg, right: LayoutTextSeg): boolean {
  const a = left.nativeLevelUnit, b = right.nativeLevelUnit;
  return !a && !b || !!a && !!b && hasNativeUnitReceipt(left) && hasNativeUnitReceipt(right)
    && a.sourceText === b.sourceText && a.start === b.start && a.end === b.end && a.baseDirection === b.baseDirection;
}

/** One selected request may retain only its exact actually measured prefix.
 * Generated/replaced text loses its base, while an unchanged source suffix
 * keeps the same word-only native request and source boundary. */
export function hasNativeVisiblePrefix(segment: LayoutTextSeg): boolean {
  const prefix = segment.nativeVisiblePrefix, request = segment.textShapeRequest;
  if (!prefix || !request || !hasNativeUnitReceipt(segment) || request.baseDirection !== 'rtl'
    || !segment.text.endsWith(' ')) return false;
  const probe = prefix.request;
  return prefix.endUTF16 === segment.text.replace(/ +$/u, '').length
    && probe.text === segment.text.slice(0, prefix.endUTF16)
    && probe.baseDirection === 'rtl' && prefix.measurement.nativeBaseDirection === 'rtl'
    && prefix.measurement.horizontalInkBoundsAreTight === true
    && probe.fontRoute.fingerprint === segment.fontRoute?.fingerprint
    && probe.fontRoute.familyList === segment.fontRoute?.familyList && probe.fontRoute.scope === segment.fontRoute?.scope
    && probe.fontSizePt === request.fontSizePt && probe.weight === (request.weight ?? 400)
    && probe.style === (request.style ?? 'normal') && probe.letterSpacingPt === (request.letterSpacingPt ?? 0)
    && probe.kerning === (wordArabicTextKerned(segment.text) ? true : request.kerning);
}

/** A preflight receipt does not certify later geometry/cache acquisitions.
 * Preserve the ordinary service call for absent base; opted-in directed source
 * requests must honor their current result, without an ordinary retry. */
export function shapeNativeSourceRequest(
  service: TextLayoutService, request: Readonly<TextShapeRequest>,
): TextShapeResult {
  if (request.baseDirection === undefined) return service.shape(request);
  if (!service.shapeDirected) throw new TypeError('Directed source request lost its directed provider');
  const shaped = service.shapeDirected(request as Readonly<TextShapeRequest & { baseDirection: 'ltr' | 'rtl' }>);
  if (shaped.nativeBaseDirection !== request.baseDirection) {
    throw new TypeError('Directed source result receipt does not honor its requested base');
  }
  return shaped;
}
