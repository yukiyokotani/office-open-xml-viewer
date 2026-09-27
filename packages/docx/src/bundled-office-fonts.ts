import {
  parseOpenTypeResourceMetrics,
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  type OfficeFontFallbackRequest,
  type OfficeFontFallbackRoute,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';

const FAMILY = '__ooxml_docx_bundled_carlito';
const MAX_FONT_BYTES = 4 * 1024 * 1024;

type Slot = 'regular' | 'bold' | 'italic' | 'boldItalic';
const SLOT = {
  regular: { weight: 400, style: 'normal' },
  bold: { weight: 700, style: 'normal' },
  italic: { weight: 400, style: 'italic' },
  boldItalic: { weight: 700, style: 'italic' },
} as const;

export interface LoadedBundledOfficeFonts {
  readonly faces: readonly FontFace[];
  readonly routes: readonly OfficeFontFallbackRoute[];
}

/** Calibri is absent on many hosts. Carlito is a published metric-compatible
 * substitute under the SIL OFL. Only document-used, unresolved Calibri tuples
 * fetch their packaged face. An embedded or exact local face stays authoritative.
 * The private CSS family avoids collisions with another document's font bytes;
 * core's font registry retains and releases the four bounded resources. */
export async function loadBundledCalibri(
  requests: readonly OfficeFontFallbackRequest[],
  resolvedTuples: ReadonlySet<string>,
): Promise<LoadedBundledOfficeFonts> {
  const slots = new Set<Slot>();
  for (const request of requests) {
    if (request.family.trim().toLowerCase() !== 'calibri') continue;
    const weight = request.weight ?? 400;
    const style = request.style ?? 'normal';
    if ((weight !== 400 && weight !== 700) || (style !== 'normal' && style !== 'italic')) continue;
    if (resolvedTuples.has(`calibri:${weight}:${style}`)) continue;
    slots.add(weight === 700 ? (style === 'italic' ? 'boldItalic' : 'bold')
      : (style === 'italic' ? 'italic' : 'regular'));
  }
  if (slots.size === 0) return { faces: [], routes: [] };
  const { CARLITO_URLS } = await import('./assets/carlito/urls.js');
  const prepared: Array<{ slot: Slot; bytes: Uint8Array; metric: ResolvedFontMetric }> = [];
  // Four fixed assets form the complete budget. Sequential fetch bounds retained
  // decode memory without making document startup depend on a network service.
  for (const slot of slots) {
    try {
      const response = await fetch(CARLITO_URLS[slot]);
      if (!response.ok) continue;
      const bytes = new Uint8Array(await response.arrayBuffer());
      if (bytes.length === 0 || bytes.length > MAX_FONT_BYTES) continue;
      const openType = parseOpenTypeResourceMetrics(bytes);
      const line = openType?.farEastCodePage == null ? null
        : wordOpenTypeAutoLineRatios({ ...openType, farEastCodePage: openType.farEastCodePage });
      const tuple = SLOT[slot];
      prepared.push({
        slot, bytes,
        metric: {
          family: FAMILY, requestedFamily: 'Calibri', ...tuple,
          sourceIdentity: `bundled:carlito:${slot}`, synthesized: false,
          ...(line ? {
            lineHeightRatio: line.lineHeightRatio,
            designAscentRatio: line.designAscentRatio,
            designDescentRatio: line.designDescentRatio,
          } : {}),
          ...(openType?.averageCharWidthRatio == null ? {}
            : { averageCharWidthRatio: openType.averageCharWidthRatio }),
          unicodeRanges: openType?.unicodeRanges ?? [],
        },
      });
    } catch {
      // An optional packaged face can fail to load in a restrictive host;
      // the authored family then retains the renderer's ordinary fallback.
    }
  }
  if (prepared.length === 0) return { faces: [], routes: [] };
  const faces = await registerEmbeddedFonts(prepared.map(({ slot, bytes }) => ({
    family: FAMILY, bytes, odttf: false,
    weight: SLOT[slot].weight === 700 ? 'bold' : 'normal',
    style: SLOT[slot].style,
  })));
  const loaded = new Set(faces.map((face) => `${face.weight}:${face.style}`));
  const routes = prepared.filter(({ slot }) => loaded.has(
    `${SLOT[slot].weight === 700 ? 'bold' : 'normal'}:${SLOT[slot].style}`,
  )).map(({ slot, metric }): OfficeFontFallbackRoute => ({
    requestedFamily: 'Calibri', family: FAMILY, source: 'substitute',
    resourceIdentity: `bundled:carlito:${slot}`,
    weight: SLOT[slot].weight, style: SLOT[slot].style,
    metric,
  }));
  return { faces, routes };
}

export function unloadBundledOfficeFonts(faces: Iterable<FontFace>): void {
  unregisterEmbeddedFonts(faces);
}
