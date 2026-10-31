import {
  findReferenceFontMetrics,
  loadNativeRouteSentinels,
  measureNativeRouteNormalLines,
  verifyNativeRouteNormalLines,
  type NativeRouteLineRequest,
  type NativeRouteNormalLine,
  unloadOfficeFontFallbacks,
} from '@silurus/ooxml-core';
import { nullRecord } from '@silurus/ooxml-core/internal/own-key';
import { docxRenderedTextUsages } from './document-content.js';
import { nativeRouteMetricKey, type NativeRouteLineMetric } from './layout/font-service.js';
import type { TextLayoutService } from './layout/text.js';
import type { DocxDocumentModel } from './types.js';

/**
 * Line geometry of an absent authored face (LIBRARY_ABSENT_FACE_FALLBACK_LINE_BOX).
 *
 * A native route (`source: 'native'` without `installedFamily`) paints with
 * the first available face of its CSS family list. Its span is sized from that
 * face's own normal line: the Canvas font box plus the leading Canvas omits,
 * measured with a CSS `line-height: normal` strut in the document realm before
 * layout (core native-route-line-metrics.ts). The pipeline is:
 *
 * 1. {@link docxNativeRouteRequests} enumerates every native family list layout
 *    can request, by resolving the document's rendered font usages through the
 *    production text service itself (language, script slot, generic family,
 *    CJK fallback and scoped substitutes are the resolver's own inputs).
 * 2. The document realm measures each list once ({@link measureDocxNativeRouteLines});
 *    a worker receives the plain numbers over its existing host channel.
 * 3. The layout realm admits a measurement only after re-identifying the same
 *    first available face with the same Canvas box ({@link admitDocxNativeRouteLines}),
 *    so main-thread and worker layout resolve one table and paginate alike.
 *
 * Policy for the first available face (this predicate): CSS cannot read
 * OS/2, so the reference-font-metrics catalog supplies the code-page class.
 * Only a named family whose every catalogued profile is non-Far-East is
 * measured. Far East faces keep today's Canvas box: their line allocation
 * is the usWin projection of #1832, which a CSS normal line does not expose;
 * unclassified families and generic keywords (an unnamed face) decline for the
 * same reason. A declined route keeps the v0.88 Canvas-box behaviour.
 */
export function docxNativeRoutePrimaryAdmitted(primaryFamily: string): boolean {
  const profiles = findReferenceFontMetrics(primaryFamily);
  return profiles.length > 0 && profiles.every((profile) => profile.farEastCodePage === false);
}

const ENUMERATION_HAN_PROBE = '漢';

/** Every native route family list (with its tuple) that layout can request
 * for this document. A list missing here is only ever declined (Canvas box),
 * identically in every realm, because admission reads one shared table. */
export function docxNativeRouteRequests(
  doc: DocxDocumentModel,
  text: Pick<TextLayoutService, 'resolve'>,
): NativeRouteLineRequest[] {
  const requests = new Map<string, NativeRouteLineRequest>();
  const visited = new Set<string>();
  const visit = (
    family: string | null | undefined, bold: boolean, italic: boolean,
    eastAsiaLanguage: string | undefined,
  ) => {
    const name = family?.trim();
    if (!name) return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' as const : 'normal' as const;
    const key = JSON.stringify([name, weight, style, eastAsiaLanguage ?? null]);
    if (visited.has(key)) return;
    visited.add(key);
    const fonts = { ascii: name, highAnsi: name, eastAsia: name, complexScript: name };
    // Regional lists apply only to Han spans (language only in the eastAsia
    // slot). Every usage is also resolved as Han: generated field results
    // (number formats, dates) can produce Han that the cached text lacks.
    for (const slot of ['ascii', 'eastAsia', 'complexScript'] as const) {
      for (const probe of ['', ENUMERATION_HAN_PROBE]) {
        const resolved = text.resolve({ text: probe, eastAsiaLanguage, fonts, slot, weight, style });
        if (resolved.source !== 'native' || resolved.installedFamily) continue;
        const request = { familyList: resolved.route.familyList, weight: resolved.weight, style: resolved.style };
        requests.set(nativeRouteMetricKey(request.familyList, request.weight, request.style), request);
      }
    }
  };
  visit(doc.majorFont, false, false, undefined);
  visit(doc.minorFont, false, false, undefined);
  for (const usage of docxRenderedTextUsages(doc, 'google-preload')) {
    const bold = usage.bold === true;
    const italic = usage.italic === true;
    for (const family of usage.fontFamilies) visit(family, bold, italic, usage.eastAsiaLanguage);
    if (usage.text && (usage.latinFontFamily === null
      || (usage.latinFontFamily === undefined && !usage.fontFamilies.some(Boolean)))) {
      visit(doc.minorFont, bold, italic, usage.eastAsiaLanguage);
    }
  }
  return [...requests.values()];
}

/** Measure in the document realm. Sentinels are loaded for the measurement;
 * a main-thread layout keeps them (returned for its release list) because it
 * also judges glyph ownership in this realm. */
export async function measureDocxNativeRouteLines(
  requests: readonly NativeRouteLineRequest[],
  retainSentinels: boolean,
): Promise<{ lines: (NativeRouteNormalLine | null)[]; faces: FontFace[] }> {
  const owner = typeof document === 'undefined' ? null : document;
  if (requests.length === 0 || !owner?.body) return { lines: requests.map(() => null), faces: [] };
  const faces = await loadNativeRouteSentinels(owner.fonts);
  if (!faces) return { lines: requests.map(() => null), faces: [] };
  let lines: (NativeRouteNormalLine | null)[] = requests.map(() => null);
  try {
    lines = measureNativeRouteNormalLines(requests, owner, docxNativeRoutePrimaryAdmitted);
  } catch { /* Declined: Canvas boxes remain. */ }
  if (retainSentinels) return { lines, faces };
  unloadOfficeFontFallbacks(faces);
  return { lines, faces: [] };
}

/** Project admitted measurements into resolver metrics for the layout realm.
 * The layout realm's sentinels must already be loaded (they judge ownership
 * during layout). The measured normal line follows the exact-local
 * projection (host-font-line-metrics.ts): leading above the baseline, the
 * Canvas descent below it. */
export function admitDocxNativeRouteLines(
  lines: readonly (NativeRouteNormalLine | null)[],
  fontSet: FontFaceSet | null,
): Record<string, NativeRouteLineMetric> {
  const measured = lines.filter((line): line is NativeRouteNormalLine => line !== null
    && docxNativeRoutePrimaryAdmitted(line.primaryFamily));
  const verified = verifyNativeRouteNormalLines(measured, fontSet);
  // Keys embed authored family lists; keep every key an own entry.
  const metrics = nullRecord<NativeRouteLineMetric>();
  measured.forEach((line, index) => {
    if (!verified[index]) return;
    const ascent = line.ascentRatio + line.lineGapRatio;
    const descent = line.descentRatio;
    if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent >= 0 && descent >= 0
      && ascent + descent > 0)) return;
    metrics[nativeRouteMetricKey(line.familyList, line.weight, line.style)] = Object.freeze({
      family: line.primaryFamily,
      requestedFamily: line.primaryFamily,
      weight: line.weight,
      style: line.style,
      // Admission above requires a catalogued non-Far-East class.
      farEastCodePage: false,
      requiresGlyphCoverage: true,
      nativeRoutePrimaryFamily: line.primaryFamilyText,
      lineHeightRatio: ascent + descent,
      designAscentRatio: ascent,
      designDescentRatio: descent,
      lineGapRatio: line.lineGapRatio,
    });
  });
  return metrics;
}
