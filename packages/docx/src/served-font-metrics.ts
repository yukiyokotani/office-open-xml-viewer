import type { OpenTypeLineMetrics, ResolvedFontMetric } from '@silurus/ooxml-core';
import {
  fontFamilyKey,
  instanceOpenTypeLineMetrics,
  parseUnicodeRange,
  type LoadedFontSources,
} from '@silurus/ooxml-core/internal/font-sources';
import { cssMatchedFace } from './line-breaker/font-metrics.js';
import { libreOfficeProducerLineMetric } from './layout/line-compatibility.js';
import { WML_TUPLES, tupleKey, type PreparedServedFaceLines } from './served-face-lines.js';
import {
  displayFamily,
  frozenRanges,
  measuredMetric,
  normalizedRanges,
  sourceFontResourceIdentity,
  wordResourceLineMetric,
} from './source-font-metrics.js';

/** The MVAR tags of the values LIBREOFFICE_FONT_LINE_BOX reads: the OS/2
 * typo values with USE_TYPO_METRICS, usWin for an empty hhea, and nothing
 * else (hhea has no MVAR tag). */
function normalLineMvarTags(tables: OpenTypeLineMetrics): ('hasc' | 'hdsc' | 'hlgp' | 'hcla' | 'hcld')[] {
  if (tables.useTypoMetrics) return ['hasc', 'hdsc', 'hlgp'];
  return tables.hheaAscent === 0 && tables.hheaDescent === 0 ? ['hcla', 'hcld'] : [];
}

// The served faces' part of sourceFontLineMetrics (source-font-metrics.ts,
// whose comment states the rules): loaded only with the font-source session
// (font-source-load.ts), so a document without `LoadOptions.fonts` loads
// neither this module nor the variable-font instancing it uses.

type Range = readonly [number, number];
type SourceFace = LoadedFontSources['faces'][number];
type Style = 'normal' | 'italic';

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

/** Declared ranges of each loaded face: snapshot faces are frozen, and
 * layout asks for the metrics more than once per snapshot. */
const facesDeclared = new WeakMap<SourceFace, ReturnType<typeof parseUnicodeRange>>();

/** The declared ranges of `face`, one parse per value within one call
 * (`values`: the faces of one file at two weights declare the same range). */
function declaredRanges(face: SourceFace, values: Map<string, ReturnType<typeof parseUnicodeRange>>): ReturnType<typeof parseUnicodeRange> {
  if (facesDeclared.has(face)) return facesDeclared.get(face)!;
  let ranges = values.get(face.unicodeRange);
  if (ranges === undefined) values.set(face.unicodeRange, ranges = parseUnicodeRange(face.unicodeRange));
  facesDeclared.set(face, ranges);
  return ranges;
}

/** {@link frozenRanges} of each declared range list ({@link declaredRanges}
 * hands the faces of one value the same list). */
const frozenByDeclared = new WeakMap<object, ReturnType<typeof frozenRanges>>();

function frozenDeclared(declared: NonNullable<ReturnType<typeof parseUnicodeRange>>): ReturnType<typeof frozenRanges> {
  let frozen = frozenByDeclared.get(declared);
  if (!frozen) frozenByDeclared.set(declared, frozen = frozenRanges(declared));
  return frozen;
}

/** Metrics of the faces `loaded` served, by tuple (see sourceFontLineMetrics). */
export function servedFontLineMetrics(
  loaded: LoadedFontSources,
  prepared: PreparedServedFaceLines,
  libreOfficeLineBoxes = false,
  libreOfficeWordFarEastBoxes = false,
): Record<string, ResolvedFontMetric> {
  const metrics: Record<string, ResolvedFontMetric> = Object.create(null);
  const byFamily = new Map<string, { display: string; faces: { face: SourceFace; index: number }[] }>();
  const values = new Map<string, ReturnType<typeof parseUnicodeRange>>();
  loaded.faces.forEach((face, index) => {
    const key = fontFamilyKey(face.family);
    if (!key) return;
    const group = byFamily.get(key);
    if (group) group.faces.push({ face, index });
    else byFamily.set(key, { display: displayFamily(face.family), faces: [{ face, index }] });
  });
  for (const { display, faces } of byFamily.values()) {
    const described: {
      face: SourceFace; index: number; range: readonly [number, number]; style: Style;
      declared: ReturnType<typeof parseUnicodeRange>;
    }[] = [];
    let modelled = true;
    for (const { face, index } of faces) {
      const range = descriptorWeight(face.weight);
      const style = descriptorStyle(face.style);
      if (!range || !style || face.stretch.trim().toLowerCase() !== 'normal') { modelled = false; break; }
      described.push({ face, index, range, style, declared: declaredRanges(face, values) });
    }
    if (!modelled) continue;
    const identity = sourceFontResourceIdentity(display);
    for (const { weight, style } of WML_TUPLES) {
      const set = matchedSet(described, weight, style);
      const projected = set.map(({ face, index, range, style: faceStyle, declared }) => {
        if (!declared) return undefined;
        if (face.resourceMetrics && libreOfficeLineBoxes
          && !(libreOfficeWordFarEastBoxes && face.resourceMetrics.farEastCodePage === true)) {
          // LIBREOFFICE_FONT_LINE_BOX reads only the face's normal line; the
          // instance must be size-free in exactly the values it reads.
          const instance = instanceOpenTypeLineMetrics(face.resourceMetrics, {
            wght: clamp(weight, range), ital: faceStyle === 'italic' ? 1 : 0,
          }, normalLineMvarTags(face.resourceMetrics));
          if (!instance?.unicodeRanges) return undefined;
          return libreOfficeProducerLineMetric(instance,
            Object.freeze(intersectRanges(normalizedRanges(declared), normalizedRanges(instance.unicodeRanges))));
        }
        if (face.resourceMetrics) {
          const instance = instanceOpenTypeLineMetrics(face.resourceMetrics, {
            // An upright face drawn for an italic request is slanted by
            // synthesis, not by its ital axis.
            wght: clamp(weight, range), ital: faceStyle === 'italic' ? 1 : 0,
          // The projection reads usWin only for a Far East face, and hhea
          // has no MVAR tag: only those usWin values must be size-free.
          }, face.resourceMetrics.farEastCodePage ? ['hcla', 'hcld'] : []);
          const line = instance && wordResourceLineMetric(instance);
          if (!instance?.unicodeRanges || !line || instance.farEastCodePage == null) return undefined;
          return {
            // The parsed class, as the installed-face catalogue carries it
            // (reference-font-line-metrics.ts): grid allocation reads it.
            farEastCodePage: instance.farEastCodePage,
            lineHeightRatio: line.lineHeightRatio,
            designAscentRatio: line.designAscentRatio,
            designDescentRatio: line.designDescentRatio,
            unicodeRanges: Object.freeze(intersectRanges(normalizedRanges(declared), normalizedRanges(instance.unicodeRanges))),
            ...(instance.hasEastAsianCmap ? { eastAsianLineHeightRatio: line.lineHeightRatio } : {}),
          };
        }
        const ownership = prepared.faceLines.get(index);
        const measured = ownership && measuredMetric(display, ownership.lines[tupleKey(weight, style)], ownership);
        return measured && { ...measured, unicodeRanges: frozenDeclared(declared) };
      });
      if (projected.length === 0 || projected.some((entry) => entry === undefined)) continue;
      projected.forEach((entry, position) => {
        metrics[`${identity}\u0000${weight}\u0000${style}\u0000${position}`] = Object.freeze({
          family: display, weight, style, sourceIdentity: identity, synthesized: false, ...entry!,
        });
      });
    }
  }
  return metrics;
}
