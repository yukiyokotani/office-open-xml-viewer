import type { OpenTypeLineMetrics, ResolvedFontMetric } from '@silurus/ooxml-core';
import { fontFamilyKey, parseUnicodeRange, type LoadedFontSources } from '@silurus/ooxml-core/internal/font-sources';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';
import { cssMatchedFace } from './line-breaker/font-metrics.js';

/** Resource identity of the faces a font source serves under one family. The
 * resolver gives the served route this identity, and only metrics built here
 * from that family's own loaded faces carry it. */
export function sourceFontResourceIdentity(servedFamily: string): string {
  return `font-source:${fontFamilyKey(servedFamily)}`;
}

/** Word's automatic single line for one face's own tables: the projection a
 * document-embedded face gets (embedded-fonts.ts) and the one the catalog
 * applies to an installed face with the same tables
 * (reference-font-line-metrics.ts). Undefined when the tables cannot support
 * it (no declared OS/2 code-page class, or a Far East face without usWin). */
export function wordResourceLineMetric(openType: OpenTypeLineMetrics): {
  lineHeightRatio: number; designAscentRatio: number; designDescentRatio: number;
} | undefined {
  if (openType.farEastCodePage == null) return undefined;
  return wordOpenTypeAutoLineRatios({ ...openType, farEastCodePage: openType.farEastCodePage }) ?? undefined;
}

type Range = readonly [number, number];
type SourceFace = LoadedFontSources['faces'][number];

function normalizedRanges(ranges: readonly Range[]): [number, number][] {
  const merged: [number, number][] = [];
  for (const [lo, hi] of [...ranges].sort((a, b) => a[0] - b[0])) {
    const last = merged[merged.length - 1];
    if (last && lo <= last[1] + 1) last[1] = Math.max(last[1], hi);
    else merged.push([lo, hi]);
  }
  return merged;
}

function intersectRanges(a: readonly Range[], b: readonly Range[]): Range[] {
  const out: Range[] = [];
  let i = 0;
  let j = 0;
  while (i < a.length && j < b.length) {
    const lo = Math.max(a[i]![0], b[j]![0]);
    const hi = Math.min(a[i]![1], b[j]![1]);
    if (lo <= hi) out.push(Object.freeze([lo, hi] as const));
    if (a[i]![1] < b[j]![1]) i++;
    else j++;
  }
  return out;
}

function descriptorWeight(value: string): number | undefined {
  const weight = value.trim().toLowerCase();
  if (weight === 'normal') return 400;
  if (weight === 'bold') return 700;
  if (!/^\d+$/.test(weight)) return undefined;
  const numeric = Number(weight);
  return numeric >= 1 && numeric <= 1000 ? numeric : undefined;
}

function descriptorStyle(value: string): 'normal' | 'italic' | undefined {
  const style = value.trim().toLowerCase();
  return style === 'normal' || style === 'italic' ? style : undefined;
}

const displayFamily = (family: string) => family.trim().replace(/^(['"])(.*)\1$/, '$2');

/**
 * Vertical geometry of the faces the `fonts` option served, from each face's
 * own OpenType tables. Layout follows the face that paints: a served face is
 * sized exactly as an installed face with the same tables (catalog projection)
 * or a document-embedded face (embedded-fonts.ts), never from the Canvas font
 * box, which omits the line gap.
 *
 * For each WordprocessingML tuple (regular/bold × upright/italic) the served
 * family's loaded faces are narrowed by CSS Fonts 4 §5.2 matching; every face
 * of the matched descriptor set contributes one metric limited to the scalars
 * it both declares (`unicode-range`) and maps (cmap), so a subset lends
 * geometry only to spans it paints. A tuple is declined (its spans keep the
 * Canvas box) when the tables of any face of the matched set are unavailable
 * (a face the browser fetched itself, WOFF/WOFF2, variable fonts). A family is
 * declined when a face's descriptors are outside what this matching models, or
 * when another loaded face in the painting FontFaceSet declares the family.
 */
export function sourceFontLineMetrics(
  loaded: LoadedFontSources,
  fontSet: FontFaceSet | null | undefined,
): Record<string, ResolvedFontMetric> {
  const byFamily = new Map<string, { display: string; faces: SourceFace[] }>();
  for (const face of loaded.faces) {
    const key = fontFamilyKey(face.family);
    if (!key) continue;
    const group = byFamily.get(key);
    if (group) group.faces.push(face);
    else byFamily.set(key, { display: displayFamily(face.family), faces: [face] });
  }
  const sessionFaces = new Set(loaded.faces.map((face) => face.face));
  const foreignFamilies = new Set<string>();
  if (fontSet && typeof fontSet[Symbol.iterator] === 'function') {
    for (const face of fontSet) {
      if (face.status === 'loaded' && !sessionFaces.has(face)) foreignFamilies.add(fontFamilyKey(displayFamily(face.family)));
    }
  }
  const metrics: Record<string, ResolvedFontMetric> = Object.create(null);
  for (const [key, { display, faces }] of byFamily) {
    if (foreignFamilies.has(key)) continue;
    const tuples: { face: SourceFace; weight: number; style: 'normal' | 'italic' }[] = [];
    let modelled = true;
    for (const face of faces) {
      const weight = descriptorWeight(face.weight);
      const style = descriptorStyle(face.style);
      if (weight === undefined || style === undefined || face.stretch.trim().toLowerCase() !== 'normal') {
        modelled = false;
        break;
      }
      tuples.push({ face, weight, style });
    }
    if (!modelled) continue;
    const identity = sourceFontResourceIdentity(display);
    for (const weight of [400, 700]) {
      for (const style of ['normal', 'italic'] as const) {
        const matched = cssMatchedFace(tuples, weight, style);
        if (!matched) continue;
        const projected = tuples
          .filter((tuple) => tuple.weight === matched.weight && tuple.style === matched.style)
          .map(({ face }) => {
            const openType = face.resourceMetrics;
            const line = openType && wordResourceLineMetric(openType);
            const declared = parseUnicodeRange(face.unicodeRange);
            if (!openType?.unicodeRanges || !line || !declared) return undefined;
            return {
              line,
              ranges: Object.freeze(intersectRanges(normalizedRanges(declared), normalizedRanges(openType.unicodeRanges))),
              eastAsian: openType.hasEastAsianCmap,
            };
          });
        if (projected.some((entry) => entry === undefined)) continue;
        projected.forEach((entry, index) => {
          const { line, ranges, eastAsian } = entry!;
          metrics[`${identity}\u0000${weight}\u0000${style}\u0000${index}`] = Object.freeze({
            family: display,
            weight,
            style,
            sourceIdentity: identity,
            synthesized: false,
            lineHeightRatio: line.lineHeightRatio,
            designAscentRatio: line.designAscentRatio,
            designDescentRatio: line.designDescentRatio,
            unicodeRanges: ranges,
            ...(eastAsian ? { eastAsianLineHeightRatio: line.lineHeightRatio } : {}),
          });
        });
      }
    }
  }
  return metrics;
}
