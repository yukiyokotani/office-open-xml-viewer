import {
  findReferenceFontMetrics,
  type LocalFontNormalLine,
  type OpenTypeLineMetrics,
  type ReferenceFontMetricProfile,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import {
  fontFamilyKey,
  instanceOpenTypeLineMetrics,
  parseUnicodeRange,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';
import { cssMatchedFace } from './line-breaker/font-metrics.js';
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
type SourceFace = LoadedFontSources['faces'][number];
type Style = 'normal' | 'italic';

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

const frozenRanges = (ranges: readonly Range[]) =>
  Object.freeze(normalizedRanges(ranges).map((range) => Object.freeze(range as Range)));

/** CSS `font-weight` descriptor: one value or a `min max` range. */
function descriptorWeight(value: string): readonly [number, number] | undefined {
  const parts = value.trim().toLowerCase().split(/\s+/);
  const one = (part: string) => part === 'normal' ? 400 : part === 'bold' ? 700
    : /^\d+(?:\.\d+)?$/.test(part) && Number(part) >= 1 && Number(part) <= 1000 ? Number(part) : undefined;
  const values = parts.map(one);
  if (parts.length > 2 || values.some((value) => value === undefined)) return undefined;
  const low = values[0]!;
  const high = values[1] ?? low;
  return low <= high ? [low, high] : [high, low];
}

function descriptorStyle(value: string): Style | undefined {
  const style = value.trim().toLowerCase();
  return style === 'normal' || style === 'italic' ? style : undefined;
}

const displayFamily = (family: string) => family.trim().replace(/^(['"])(.*)\1$/, '$2');
const clamp = (value: number, [low, high]: readonly [number, number]) => Math.min(high, Math.max(low, value));

/** CSS Fonts 4 §5.2 measures a weight range by its nearest value: the face
 * matches as if its weight were the request clamped to the range. */
function matchedSet<T extends { range: readonly [number, number]; style: Style }>(
  faces: readonly T[], weight: number, style: Style,
): T[] {
  const candidates = faces.map((face) => ({ face, weight: clamp(weight, face.range), style: face.style }));
  const matched = cssMatchedFace(candidates, weight, style);
  return matched ? candidates.filter((candidate) => candidate.weight === matched.weight
    && candidate.style === matched.style && candidate.face.range[0] <= matched.weight
    && matched.weight <= candidate.face.range[1]).map(({ face }) => face) : [];
}

/** Two quantized measurements per em (the strut height is an integer CSS
 * pixel, the Canvas box carries sub-pixel rounding) at the probe em. */
const PROFILE_MATCH_TOLERANCE_EM = 2 / 1024;

/** The vertical boxes a rasterizer may report for a catalogued face as its
 * CSS normal line: hhea, usWin, or typo when USE_TYPO_METRICS is set. */
function profileBoxes(profile: ReferenceFontMetricProfile): (readonly [number, number, number])[] {
  const em = profile.unitsPerEm;
  const boxes: (readonly [number, number, number])[] = [[profile.hhea[0] / em, -profile.hhea[1] / em, profile.hhea[2] / em]];
  if (profile.win) boxes.push([profile.win[0] / em, profile.win[1] / em, 0]);
  if (profile.typoMetrics) boxes.push([profile.typoMetrics[0] / em, -profile.typoMetrics[1] / em, profile.typoMetrics[2] / em]);
  return boxes;
}

/**
 * The auto line of a measured face the catalogue classifies as Far East. A
 * CSS normal line cannot show OS/2 code pages or usWin (the Far East line is
 * 1.3 × usWin, #1832), but it can identify the data: when the measured box
 * equals a box the catalogue records for this family and tuple, the
 * catalogue's tables are this face's, and wordOpenTypeAutoLineRatios is
 * applied to them exactly as for an installed face. Every matching profile must project to the same line; no match, or
 * disagreeing matches, decline.
 */
function catalogueFarEastLine(
  served: string, weight: number, style: Style, line: LocalFontNormalLine,
): { lineHeightRatio: number; designAscentRatio: number; designDescentRatio: number } | undefined {
  const matching = findReferenceFontMetrics(served, { weight, style }).filter((profile) =>
    profileBoxes(profile).some(([ascent, descent, gap]) =>
      Math.abs(line.ascentRatio - ascent) <= PROFILE_MATCH_TOLERANCE_EM
      && Math.abs(line.descentRatio - descent) <= PROFILE_MATCH_TOLERANCE_EM
      && Math.abs(line.lineGapRatio - gap) <= PROFILE_MATCH_TOLERANCE_EM));
  let result: { lineHeightRatio: number; designAscentRatio: number; designDescentRatio: number } | undefined;
  for (const profile of matching) {
    if (profile.farEastCodePage !== true || !profile.win) return undefined;
    const projected = wordOpenTypeAutoLineRatios({
      unitsPerEm: profile.unitsPerEm,
      hheaAscent: profile.hhea[0], hheaDescent: profile.hhea[1], hheaLineGap: profile.hhea[2],
      winAscent: profile.win[0], winDescent: profile.win[1], farEastCodePage: true,
    });
    if (!projected || (result && (projected.lineHeightRatio !== result.lineHeightRatio
      || projected.designAscentRatio !== result.designAscentRatio
      || projected.designDescentRatio !== result.designDescentRatio))) return undefined;
    result = projected;
  }
  return result;
}

/** The measured line of one tuple as resource geometry (the projection of
 * host-font-line-metrics.ts: leading above the baseline, Canvas descent
 * below), sized only where the measured face is proven to draw. A family the
 * catalogue classifies as Far East takes the catalogue's Word line when the
 * measurement identifies the catalogued data (catalogueFarEastLine), and is
 * declined otherwise, as for installed faces. */
function measuredMetric(
  served: string, weight: number, style: Style,
  line: LocalFontNormalLine | null | undefined, ownership: ServedFaceLine,
): Omit<ResolvedFontMetric, 'family' | 'weight' | 'style' | 'sourceIdentity'> | undefined {
  if (!line) return undefined;
  const ownedBy = {
    requiresGlyphCoverage: true,
    nativeRoutePrimaryFamily: quoteCssFontFamily(ownership.ownershipFamily),
  } as const;
  if (findReferenceFontMetrics(served).some((profile) => profile.farEastCodePage !== false)) {
    const farEast = catalogueFarEastLine(served, weight, style, line);
    return farEast && { farEastCodePage: true, ...ownedBy, ...farEast };
  }
  const ascent = line.ascentRatio + line.lineGapRatio;
  const descent = line.descentRatio;
  if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent >= 0 && descent >= 0 && ascent + descent > 0)) return undefined;
  return {
    farEastCodePage: false,
    ...ownedBy,
    lineHeightRatio: ascent + descent,
    designAscentRatio: ascent,
    designDescentRatio: descent,
    lineGapRatio: line.lineGapRatio,
  };
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
 * - No tables: the face's measured normal line (served-face-lines.ts),
 *   limited to its declared unicode-range and to clusters the route
 *   sentinels prove it draws.
 *
 * For each WordprocessingML tuple the served family's faces are narrowed by
 * CSS Fonts 4 §5.2 matching; every face of the matched set contributes one
 * metric. A tuple is declined (its spans keep the Canvas box) when a face of
 * the matched set has neither; a family is declined when a face's descriptors
 * are outside this matching (oblique, non-normal stretch), or when another
 * loaded face in the painting FontFaceSet declares the family.
 */
export function sourceFontLineMetrics(
  loaded: LoadedFontSources | undefined,
  fontSet: FontFaceSet | null | undefined,
  prepared: PreparedServedFaceLines = NO_SERVED_FACE_LINES,
): Record<string, ResolvedFontMetric> {
  const metrics: Record<string, ResolvedFontMetric> = Object.create(null);
  const byFamily = new Map<string, { display: string; faces: { face: SourceFace; index: number }[] }>();
  (loaded?.faces ?? []).forEach((face, index) => {
    const key = fontFamilyKey(face.family);
    if (!key) return;
    const group = byFamily.get(key);
    if (group) group.faces.push({ face, index });
    else byFamily.set(key, { display: displayFamily(face.family), faces: [{ face, index }] });
  });
  const sessionFaces = new Set(loaded?.faces.map((face) => face.face) ?? []);
  const foreignFamilies = new Set<string>();
  if (fontSet && typeof fontSet[Symbol.iterator] === 'function') {
    for (const face of fontSet) {
      if (face.status === 'loaded' && !sessionFaces.has(face)) foreignFamilies.add(fontFamilyKey(displayFamily(face.family)));
    }
  }
  for (const [key, { display, faces }] of byFamily) {
    if (foreignFamilies.has(key)) continue;
    const described: { face: SourceFace; index: number; range: readonly [number, number]; style: Style }[] = [];
    let modelled = true;
    for (const { face, index } of faces) {
      const range = descriptorWeight(face.weight);
      const style = descriptorStyle(face.style);
      if (!range || !style || face.stretch.trim().toLowerCase() !== 'normal') { modelled = false; break; }
      described.push({ face, index, range, style });
    }
    if (!modelled) continue;
    const identity = sourceFontResourceIdentity(display);
    for (const { weight, style } of WML_TUPLES) {
      const set = matchedSet(described, weight, style);
      const projected = set.map(({ face, index, range, style: faceStyle }) => {
        const declared = parseUnicodeRange(face.unicodeRange);
        if (!declared) return undefined;
        if (face.resourceMetrics) {
          const instance = instanceOpenTypeLineMetrics(face.resourceMetrics, {
            // An upright face drawn for an italic request is slanted by
            // synthesis, not by its ital axis.
            wght: clamp(weight, range), ital: faceStyle === 'italic' ? 1 : 0,
          });
          const line = instance && wordResourceLineMetric(instance);
          if (!instance?.unicodeRanges || !line) return undefined;
          return {
            lineHeightRatio: line.lineHeightRatio,
            designAscentRatio: line.designAscentRatio,
            designDescentRatio: line.designDescentRatio,
            unicodeRanges: Object.freeze(intersectRanges(normalizedRanges(declared), normalizedRanges(instance.unicodeRanges))),
            ...(instance.hasEastAsianCmap ? { eastAsianLineHeightRatio: line.lineHeightRatio } : {}),
          };
        }
        const ownership = prepared.faceLines.get(index);
        const measured = ownership && measuredMetric(display, weight, style, ownership.lines[tupleKey(weight, style)], ownership);
        return measured && { ...measured, unicodeRanges: frozenRanges(declared) };
      });
      if (projected.length === 0 || projected.some((entry) => entry === undefined)) continue;
      projected.forEach((entry, position) => {
        metrics[`${identity}\u0000${weight}\u0000${style}\u0000${position}`] = Object.freeze({
          family: display, weight, style, sourceIdentity: identity, synthesized: false, ...entry!,
        });
      });
    }
  }
  // Application-registered faces (served-face-lines.ts appRegisteredFaces).
  for (const [key, ownership] of prepared.familyLines) {
    if (byFamily.has(key)) continue;
    const family = ownership.ownershipFamily;
    const identity = appFontResourceIdentity(family);
    const face = fontSet && typeof fontSet[Symbol.iterator] === 'function'
      ? [...fontSet].find((candidate) => fontFamilyKey(displayFamily(candidate.family)) === key) : undefined;
    const declared = parseUnicodeRange(face?.unicodeRange) ?? [[0, 0x10ffff] as const];
    for (const { weight, style } of WML_TUPLES) {
      const measured = measuredMetric(family, weight, style, ownership.lines[tupleKey(weight, style)], ownership);
      if (!measured) continue;
      metrics[`${identity}\u0000${weight}\u0000${style}`] = Object.freeze({
        family, weight, style, sourceIdentity: identity, synthesized: false, ...measured, unicodeRanges: frozenRanges(declared),
      });
    }
  }
  return metrics;
}
