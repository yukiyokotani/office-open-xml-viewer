import type { CanvasFontRoute, ResolvedFontMetric } from '@silurus/ooxml-core';
import { localFontMetricCoversText, normalizeFontMetricFamily } from '@silurus/ooxml-core';
import type { FontResolution } from '../layout/font-service.js';
import type { TextLayoutService } from '../layout/text.js';
import { referenceFontLineMetrics, type ReferenceFontLineMetrics } from '../reference-font-line-metrics.js';
import { LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY } from '../layout/line-compatibility.js';

/** Text that proves which face a native route paints (#1848). */
export interface SynthesizedFaceProbe {
  readonly text: TextLayoutService;
  readonly probe: string;
  readonly baseDirection?: 'ltr' | 'rtl';
}

const EAST_ASIAN_PROBE_SCRIPT = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}]/u;

type FaceTuple = Readonly<{ weight: number; style: 'normal' | 'italic' }>;

/** CSS Fonts Level 4 §5.2 step 4 over the family's loaded faces: style
 * narrows first (italic, then normal; no oblique faces are loaded), then
 * weight: above 500 heavier ascending then lighter descending; below 400
 * lighter descending then heavier ascending; 400–500 heavier up to 500, then
 * lighter descending, then above 500 ascending. */
function cssMatchedFace<T extends FaceTuple>(faces: readonly T[], weight: number, style: 'normal' | 'italic'): T | undefined {
  for (const candidateStyle of style === 'italic' ? ['italic', 'normal'] as const : ['normal', 'italic'] as const) {
    const sameStyle = faces.filter((face) => face.style === candidateStyle);
    if (sameStyle.length === 0) continue;
    const lighter = sameStyle.filter((face) => face.weight < weight).sort((a, b) => b.weight - a.weight);
    const heavier = sameStyle.filter((face) => face.weight > weight).sort((a, b) => a.weight - b.weight);
    const exact = sameStyle.filter((face) => face.weight === weight);
    const order = weight > 500 ? [...exact, ...heavier, ...lighter]
      : weight < 400 ? [...exact, ...lighter, ...heavier]
      : [...exact, ...heavier.filter((face) => face.weight <= 500), ...lighter,
        ...heavier.filter((face) => face.weight > 500)];
    return order[0];
  }
  return undefined;
}

/**
 * Catalog vertical geometry for a selected face: the exact tuple, or, for a
 * native route of an installed family whose tuple has no catalog entry (bold
 * of a face catalogued only at 400, such as MS Mincho), the geometry of the
 * loaded exact local() face that the route is PROVED to paint. The proof
 * rasterizes probe glyphs of the span's script class through the native route
 * and through a private
 * family holding only that loaded face; CSS Fonts Level 4 §5.2 makes the
 * private family synthesize the missing tuple from that one face. Pixel
 * identity shows the native route paints the same synthesized face, whose
 * vertical tables (and so its Word 1.3 × usWin Far-East line, #1832) are the
 * loaded face's. A difference means another face paints (a genuine styled face
 * absent from the catalog, or a fallback face), and the span keeps its own
 * Canvas box. Catalog absence alone is never treated as synthesis.
 */
export function authoredReferenceVerticalMetric(
  selected: FontResolution,
  synthesis?: SynthesizedFaceProbe,
  plain = false,
): ReferenceFontLineMetrics | undefined {
  void LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY;
  const exact = referenceFontLineMetrics(selected.requestedFamily, selected.weight, selected.style, plain);
  if (exact || !synthesis?.text.sameGlyphRaster
    || selected.source !== 'native' || selected.installedFamily !== true) return exact;
  const faces = synthesis.text.exactLocalFaces?.(selected.requestedFamily) ?? [];
  const matched = cssMatchedFace(faces, selected.weight, selected.style);
  // Glyph ownership: the span's own graphemes (and the raster probe's) must
  // come from the matched exact source, not from a fallback face.
  if (!matched?.coverageProbeFamily || !synthesis.text.ownsGlyphs) return undefined;
  const reference = referenceFontLineMetrics(selected.requestedFamily, matched.weight, matched.style, plain);
  // Face identity is a property of the route and tuple, as for the regular
  // catalog tuple this extends (which admits a native installed route without
  // per-glyph proof). One fixed probe per script class bounds the work to a
  // few rasterizations per document instead of one per distinct span.
  const probe = EAST_ASIAN_PROBE_SCRIPT.test(synthesis.probe) ? '国あ' : 'Hxg';
  const owned = (text: string) => synthesis.text.ownsGlyphs!({
    text, weight: selected.weight, style: selected.style, coverageProbeFamily: matched.coverageProbeFamily!,
    ...(synthesis.baseDirection !== undefined ? { baseDirection: synthesis.baseDirection } : {}),
  });
  return reference && owned(probe) && owned(synthesis.probe) && synthesis.text.sameGlyphRaster({
    text: probe,
    weight: selected.weight,
    style: selected.style,
    route: selected.route,
    faceFamily: matched.resolvedFamily,
    ...(synthesis.baseDirection !== undefined ? { baseDirection: synthesis.baseDirection } : {}),
  }) ? reference : undefined;
}

/** Vertical admission shared by ordinary text and numbering markers. Resource
 * identity, tuple ambiguity and glyph coverage are resolved before this gate.
 * Canvas sides remain the fallback; exact spacing suppresses design geometry. */
export function selectedFontLineMetric(
  selected: FontResolution | undefined,
  localFont: ResolvedFontMetric | undefined,
  naturalMetricAllowed: boolean,
  referenceMetricAllowed = true,
  synthesis?: SynthesizedFaceProbe,
  plain = false,
) {
  const resourceMetric = (naturalMetricAllowed || localFont?.designAscentRatio == null)
    && (localFont?.lineHeightRatio != null || localFont?.designAscentRatio != null
      || localFont?.eastAsianLineHeightRatio != null) ? localFont : undefined;
  const referenceMetric = naturalMetricAllowed && referenceMetricAllowed && !resourceMetric
    && mayUseAuthoredReferenceVerticalMetric(selected)
    ? authoredReferenceVerticalMetric(selected, synthesis, plain)
    : undefined;
  return { resourceMetric, referenceMetric, lineMetric: resourceMetric ?? referenceMetric };
}

/** A subset face lends design metrics only to spans whose every scalar it owns.
 * A FontFace tuple can win CSS selection while Canvas silently paints missing
 * glyphs from another face. Ranges are normalized once at the service boundary.
 *
 * `paintsGlyphs` is false only for a probe that selects a face without painting
 * it: the paragraph mark (§17.3.1.29). Painted-glyph ownership exists to keep a
 * span from borrowing the geometry of a face Canvas did not draw it with; a
 * mark draws nothing, so its box is the selected resource's own line box even
 * when that face lacks the script-matched probe glyph (for example an
 * East Asian slot authored with a Latin face). Unicode-range subsets still
 * gate the match, because they decide which registered resource is selected. */
export function metricCoversText(
  metric: ResolvedFontMetric, text: string, route?: CanvasFontRoute, paintsGlyphs = true,
): boolean {
  if (paintsGlyphs && metric.requiresGlyphCoverage
    && !localFontMetricCoversText(metric, text, route)) return false;
  const ranges = metric.unicodeRanges;
  if (ranges === undefined || text.length === 0) return true;
  for (const scalar of text) {
    const codePoint = scalar.codePointAt(0)!;
    let low = 0;
    let high = ranges.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (ranges[middle][1] < codePoint) low = middle + 1;
      else high = middle;
    }
    if (low >= ranges.length || ranges[low][0] > codePoint) return false;
  }
  return true;
}


/** An authored-family key alone cannot prove the selected route. Match the
 * chosen face tuple and acquisition source before admitting its metric;
 * a substitute may preserve the request while painting another face. */
export function metricMatchesSelectedFont(
  metric: ResolvedFontMetric,
  selected: FontResolution,
  text: string,
  paintsGlyphs = true,
): boolean {
  if (resourceMetricKey(metric) !== resolutionMetricKey(selected)
    || !metricCoversText(metric, text, selected.route, paintsGlyphs)) return false;
  const identity = metric.sourceIdentity;
  if (selected.source === 'embedded') {
    // The production loader routes Canvas to a resource-derived private family.
    // Its metric must come from that same registration, including when another
    // document embeds different bytes under the same authored family/style.
    return selected.resourceIdentity === undefined
      ? identity?.startsWith('embedded:') === true
      : identity === selected.resourceIdentity;
  }
  if (selected.source === 'local' || selected.source === 'substitute') {
    // The same CSS tuple may be supplied by an unrelated caller or another
    // loaded face. Only the resource that created this resolver route may lend
    // vertical geometry; a matching family name is insufficient.
    return identity !== undefined && identity === selected.resourceIdentity;
  }
  return false;
}


/** Multiple resources may be registered under one CSS family/style. The
 * authored alias does not control Canvas's choice among overlapping faces;
 * only geometry shared by every covering resource is safe to admit. */
export function sameResourceGeometry(a: ResolvedFontMetric, b: ResolvedFontMetric): boolean {
  return a.farEastCodePage === b.farEastCodePage
    && a.lineHeightRatio === b.lineHeightRatio
    && a.designAscentRatio === b.designAscentRatio
    && a.designDescentRatio === b.designDescentRatio
    && a.lineGapRatio === b.lineGapRatio
    && a.eastAsianLineHeightRatio === b.eastAsianLineHeightRatio
    && a.fontBoxRatio === b.fontBoxRatio;
}


export type MetricTupleIndex = ReadonlyMap<string, readonly ResolvedFontMetric[]>;

export const metricTupleIndexes = new WeakMap<object, MetricTupleIndex>();


export function selectedMetricKey(
  family: string,
  weight: number,
  style: string,
): string {
  return `${normalizeFontMetricFamily(family)}\0${weight}\0${style}`;
}


const resourceMetricKeys = new WeakMap<ResolvedFontMetric, string>();

function resourceMetricKey(metric: ResolvedFontMetric): string {
  const cached = resourceMetricKeys.get(metric);
  if (cached !== undefined) return cached;
  const key = selectedMetricKey(metric.family, metric.weight ?? 400, metric.style ?? 'normal');
  if (Object.isFrozen(metric)) resourceMetricKeys.set(metric, key);
  return key;
}

const resolutionMetricKeys = new WeakMap<FontResolution, string>();

/** Retained resolutions are immutable and shared by every span of a tuple.
 * Local aliases can be long: normalize once rather than at every line-metric
 * and catalog-admission lookup. Mutable custom resolutions stay uncached. */
function resolutionMetricKey(selected: FontResolution): string {
  const cached = resolutionMetricKeys.get(selected);
  if (cached !== undefined) return cached;
  const key = selectedMetricKey(selected.resolvedFamily, selected.weight, selected.style);
  if (Object.isFrozen(selected)) resolutionMetricKeys.set(selected, key);
  return key;
}


/** A frozen service snapshot is immutable for the document lifetime, so one
 * tuple/source index can serve every paragraph. A custom mutable service gets
 * one temporary index per synchronous buildSegments call, then can change its
 * entries before the next call. */
export function indexedFontMetrics(
  metrics: Readonly<Record<string, ResolvedFontMetric>>,
): MetricTupleIndex {
  const frozenRoot = Object.isFrozen(metrics);
  if (frozenRoot) {
    const retained = metricTupleIndexes.get(metrics);
    if (retained) return retained;
  }
  const values = Object.values(metrics);
  const cacheable = frozenRoot && values.every((metric) => Object.isFrozen(metric));
  const mutable = new Map<string, ResolvedFontMetric[]>();
  for (const metric of values) {
    const key = resourceMetricKey(metric);
    const candidates = mutable.get(key);
    if (candidates) candidates.push(metric);
    else mutable.set(key, [metric]);
  }
  for (const candidates of mutable.values()) Object.freeze(candidates);
  if (cacheable) metricTupleIndexes.set(metrics, mutable);
  return mutable;
}


export function selectResourceMetric(
  index: MetricTupleIndex,
  selected: FontResolution,
  probeText: string,
  paintsGlyphs = true,
): ResolvedFontMetric | undefined {
  if (selected.source === 'native') {
    // The measured line describes the route's first available face, not the
    // absent authored one; it lends geometry only to clusters that face draws
    // (or, for a paragraph mark that draws nothing, to the face it selects).
    const metric = selected.routeLineMetric;
    return metric && metricCoversText(metric, probeText, selected.route, paintsGlyphs)
      ? metric : undefined;
  }
  if (selected.source !== 'embedded' && selected.source !== 'local'
    && selected.source !== 'substitute') return undefined;
  const candidates = index.get(resolutionMetricKey(selected)) ?? [];
  let first: ResolvedFontMetric | undefined;
  for (const metric of candidates) {
    if (!metricMatchesSelectedFont(metric, selected, probeText, paintsGlyphs)) continue;
    first ??= metric;
  }
  if (!first) return undefined;
  // Source identity proves which resource was requested, but not which of the
  // overlapping CSS faces Canvas actually paints. Compare every covering
  // resource in the tuple, including resources acquired from other sources.
  for (const metric of candidates) {
    if (metricCoversText(metric, probeText, selected.route, paintsGlyphs)
      && !sameResourceGeometry(first, metric)) return undefined;
  }
  return first;
}


/** Width inference has a stricter, independent projection. Conflicting xAvg
 * metadata cannot revoke a valid vertical metric for the same selected face. */
export function selectResourceAverageWidthRatio(
  index: MetricTupleIndex,
  selected: FontResolution,
  probeText: string,
): number | undefined {
  if (selected.source !== 'embedded' && selected.source !== 'local'
    && selected.source !== 'substitute') return undefined;
  const candidates = index.get(resolutionMetricKey(selected)) ?? [];
  const withSpace = `${probeText} `;
  const first = candidates.find((metric) => metricMatchesSelectedFont(metric, selected, withSpace));
  const ratio = first?.averageCharWidthRatio;
  if (ratio == null || ratio <= 0) return undefined;
  return candidates.every((metric) => !metricCoversText(metric, withSpace, selected.route)
    || metric.averageCharWidthRatio === ratio) ? ratio : undefined;
}


export function mayUseExactLocalReferenceWidthMetric(
  selected: FontResolution | undefined,
): selected is FontResolution {
  // The width floor requires an exact selected face. A native CSS family list
  // cannot establish which fallback Canvas painted or measured the glyph.
  return selected?.source === 'local'
    && selected.resourceIdentity?.startsWith('office-local:') === true;
}


export function mayUseAuthoredReferenceVerticalMetric(
  selected: FontResolution | undefined,
): selected is FontResolution {
  // Layout follows the face that actually paints on the host. A native CSS
  // family list whose authored face is not installed paints with a browser
  // fallback, so the authored face's catalog line box would size lines for
  // glyphs drawn by a different face; that span keeps its own Canvas box, or
  // the painting face's measured normal line (FontResolution.routeLineMetric,
  // admitted by selectResourceMetric for the clusters that face draws).
  // See LIBRARY_ABSENT_FACE_FALLBACK_LINE_BOX (layout/line-compatibility.ts).
  // A loaded local() tuple uses the catalog only when neither parsed nor
  // prepared host measurements supply the required geometry. A registered
  // substitute is a different selected face: if its resource metric is
  // unavailable, use its Canvas line box in both axes instead of importing the
  // authored font's vertical geometry.
  // A native resolution of an installed family is present: the installed
  // family paints (CSS font matching), so its catalog geometry may apply.
  return (selected?.source === 'native' && selected.installedFamily === true)
    || mayUseExactLocalReferenceWidthMetric(selected);
}

/** A prepared host face that failed coverage must fall through to the span's
 * native measurement, never re-enter through its authored catalog profile. */
export function allowsReferenceFontMetric(
  index: MetricTupleIndex, selected: FontResolution | undefined,
): boolean {
  return !selected || !(index.get(resolutionMetricKey(selected)) ?? []).some(metric => metric.requiresGlyphCoverage);
}
