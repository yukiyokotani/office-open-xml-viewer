import {
  findReferenceFontMetrics,
  type LocalFontNormalLine,
  type OpenTypeLineMetrics,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import {
  fontFamilyKey,
  parseUnicodeRange,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';
import { NO_SERVED_FACE_LINES, WML_TUPLES, tupleKey, type PreparedServedFaceLines, type ServedFaceLine } from './served-face-lines.js';

/** Resource identity of the faces a font source serves under one family. The
 * resolver gives the served route this identity, and only metrics built here
 * from that family's own loaded faces carry it. */
export function sourceFontResourceIdentity(servedFamily: string): string {
  return `font-source:${fontFamilyKey(servedFamily)}`;
}

/** Resource identity of an application-registered FontFace family. */
export function appFontResourceIdentity(family: string): string {
  return `app-face:${fontFamilyKey(family)}`;
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

export function normalizedRanges(ranges: readonly Range[]): [number, number][] {
  const merged: [number, number][] = [];
  for (const [lo, hi] of [...ranges].sort((a, b) => a[0] - b[0])) {
    const last = merged[merged.length - 1];
    if (last && lo <= last[1] + 1) last[1] = Math.max(last[1], hi);
    else merged.push([lo, hi]);
  }
  return merged;
}

export const frozenRanges = (ranges: readonly Range[]) =>
  Object.freeze(normalizedRanges(ranges).map((range) => Object.freeze(range as Range)));

export const displayFamily = (family: string) => family.trim().replace(/^(['"])(.*)\1$/, '$2');

/** The measured line of one tuple as resource geometry (the projection of
 * host-font-line-metrics.ts: leading above the baseline, Canvas descent
 * below), sized only where the measured face is proven to draw. A CSS normal
 * line shows neither the OS/2 code pages nor usWin, so the face's Far East
 * class is unknown: a family the catalogue lists with any Far East (or
 * unknown) profile is declined, as for installed faces, and the metric
 * carries no class. */
export function measuredMetric(
  served: string,
  line: LocalFontNormalLine | null | undefined, ownership: ServedFaceLine,
): Omit<ResolvedFontMetric, 'family' | 'weight' | 'style' | 'sourceIdentity'> | undefined {
  if (!line) return undefined;
  if (findReferenceFontMetrics(served).some((profile) => profile.farEastCodePage !== false)) return undefined;
  const ascent = line.ascentRatio + line.lineGapRatio;
  const descent = line.descentRatio;
  if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent >= 0 && descent >= 0 && ascent + descent > 0)) return undefined;
  return {
    requiresGlyphCoverage: true,
    nativeRoutePrimaryFamily: quoteCssFontFamily(ownership.ownershipFamily),
    lineHeightRatio: ascent + descent,
    designAscentRatio: ascent,
    designDescentRatio: descent,
    lineGapRatio: line.lineGapRatio,
  };
}

/** The served faces' part of {@link sourceFontLineMetrics}
 * (served-font-metrics.ts), loaded with the font-source session and carried
 * with the loaded sources, so a document without sources never loads it. */
export type ServedFontLineMetrics = (
  loaded: LoadedFontSources,
  prepared: PreparedServedFaceLines,
) => Record<string, ResolvedFontMetric>;

/** One document's loaded font sources, as layout reads them. */
export interface DocxSourceFonts {
  /** Demanded names (authored, then fallback) the inventory iterates. */
  readonly names: readonly string[];
  readonly loaded: LoadedFontSources;
  readonly lineMetrics: ServedFontLineMetrics;
}

/**
 * Vertical geometry of the faces the `fonts` option served, and of
 * application-registered FontFaces of authored families. Layout follows the
 * face that paints: a served face is sized exactly as an installed or embedded
 * face with the same tables, never from the Canvas font box, which omits the
 * line gap.
 *
 * - Tables (`resourceMetrics`): Word's projection of the face's own tables,
 *   at the variable instance CSS selects for the tuple (`wght` = the weight
 *   clamped to the face's weight descriptor, `ital` = 1 for an italic face,
 *   other axes at their defaults), limited to declared unicode-range ∩ cmap.
 *   The projection is wordOpenTypeAutoLineRatios with the inputs the
 *   installed-face catalogue (reference-font-line-metrics.ts) and embedded
 *   faces (embedded-fonts.ts) pass: hhea, usWin and the parsed code-page
 *   class. OS/2 USE_TYPO_METRICS and the typo values are an input of none of
 *   these paths.
 * - No tables: the face's measured normal line (served-face-lines.ts),
 *   limited to its declared unicode-range and to clusters the route
 *   sentinels prove it draws.
 *
 * For each WordprocessingML tuple the served family's faces are narrowed by
 * CSS Fonts 4 §5.2 matching; every face of the matched set contributes one
 * metric. A tuple is declined (its spans keep the Canvas box) when a face of
 * the matched set has neither; a family is declined when a face's descriptors
 * are outside this matching (oblique, non-normal stretch). The served faces
 * are registered under the document's private families (core
 * source-session.ts), so no face of another document or of the application
 * can join a matched set: the metrics describe exactly the faces that paint.
 */
export function sourceFontLineMetrics(
  sources: Pick<DocxSourceFonts, 'loaded' | 'lineMetrics'> | undefined,
  fontSet: FontFaceSet | null | undefined,
  prepared: PreparedServedFaceLines = NO_SERVED_FACE_LINES,
): Record<string, ResolvedFontMetric> {
  const metrics: Record<string, ResolvedFontMetric> = sources?.loaded.faces.length
    ? sources.lineMetrics(sources.loaded, prepared) : Object.create(null);
  // Families with a served face are sized from those faces only.
  const servedFamilies = new Set<string>();
  for (const face of sources?.loaded.faces ?? []) {
    const key = fontFamilyKey(face.family);
    if (key) servedFamilies.add(key);
  }
  // Application-registered faces (served-face-lines.ts appRegisteredFaces).
  for (const [key, ownership] of prepared.familyLines) {
    if (servedFamilies.has(key)) continue;
    const family = ownership.ownershipFamily;
    const identity = appFontResourceIdentity(family);
    const face = fontSet && typeof fontSet[Symbol.iterator] === 'function'
      ? [...fontSet].find((candidate) => fontFamilyKey(displayFamily(candidate.family)) === key) : undefined;
    const declared = parseUnicodeRange(face?.unicodeRange) ?? [[0, 0x10ffff] as const];
    for (const { weight, style } of WML_TUPLES) {
      const measured = measuredMetric(family, ownership.lines[tupleKey(weight, style)], ownership);
      if (!measured) continue;
      metrics[`${identity}\u0000${weight}\u0000${style}`] = Object.freeze({
        family, weight, style, sourceIdentity: identity, synthesized: false, ...measured, unicodeRanges: frozenRanges(declared),
      });
    }
  }
  return metrics;
}
