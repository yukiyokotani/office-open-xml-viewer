import {
  activeFontSet,
  findReferenceFontMetrics,
  loadOfficeFontFallbacks,
  measureLocalFontNormalLine,
  loadLocalFontCoverageProbe,
  measureLocalFontSourceNormalLine,
  type LocalFontNormalLine,
  type OfficeFontFallbackRoute,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import { nullRecord } from '@silurus/ooxml-core/internal/own-key';

/** Resource-aware vertical admission, independent of viewer OS and family
 * spelling. ECMA-376 §17.3.1.33 defines spacing in normal lines, not a catalog
 * version of an installed face. For ordinary local faces retain the host's
 * measured normal box, including signed leading above the baseline (the same
 * partition as wordOpenTypeAutoLineRatios). Catalog profiles are fallback data.
 *
 * CSS normal-line measurement cannot expose OS/2 class/usWin/typo tables.
 * Classification remains catalog fallback metadata, not proof of host OS/2
 * bytes; unknown or disagreeing classes decline this projection.
 * Known Far East tuples therefore decline this measurement and keep the
 * catalog usWin fallback from #1832; parsed resource metrics, including a
 * different host usWin box, always win. This is a measurement capability gate,
 * not a platform selector or a claim that Canvas's box is usWin. Substitutes
 * use their own measured normal box and never the authored face's catalog. */
export function hostFontLineMetric(
  route: OfficeFontFallbackRoute,
  measured: LocalFontNormalLine | null,
): ResolvedFontMetric {
  if (route.metric.lineHeightRatio != null || !measured) return route.metric;
  const profiles = findReferenceFontMetrics(route.requestedFamily, {
    weight: route.weight, style: route.style,
  });
  if (route.source !== 'substitute'
    && profiles.some((profile) => profile.farEastCodePage !== false)) return route.metric;
  const ascent = measured.ascentRatio + measured.lineGapRatio;
  const descent = measured.descentRatio;
  if (!(Number.isFinite(ascent) && Number.isFinite(descent)
    && ascent >= 0 && descent >= 0 && ascent + descent > 0)) return route.metric;
  return Object.freeze({
    ...route.metric,
    requiresGlyphCoverage: true,
    lineHeightRatio: ascent + descent,
    designAscentRatio: ascent,
    designDescentRatio: descent,
    lineGapRatio: measured.lineGapRatio,
    ...(profiles.length > 0 && route.source === 'local' ? { farEastCodePage: false } : {}),
  });
}

/** The DOM supplies leading unavailable on OffscreenCanvas. A worker asks its
 * owning DOM realm for numeric measurements of its positively resolved source;
 * the response is admitted only after matching resource identity and tuple.
 * No parser model or font bytes cross this measurement boundary. */
export interface HostFontLineMeasurement {
  resourceIdentity: string;
  weight: number;
  style: 'normal' | 'italic';
  measured: LocalFontNormalLine | null;
}

function probeEm(route: OfficeFontFallbackRoute): number | undefined {
  if (route.metric.lineHeightRatio != null) return undefined;
  const profiles = findReferenceFontMetrics(route.requestedFamily, {
    weight: route.weight, style: route.style,
  });
  return route.source === 'substitute' || profiles.every(profile => profile.farEastCodePage === false)
    ? profiles[0]?.unitsPerEm : undefined;
}

export async function measureHostFontLines(
  routes: readonly OfficeFontFallbackRoute[],
): Promise<HostFontLineMeasurement[]> {
  const owner = typeof document === 'undefined' ? null : document;
  return Promise.all(routes.map(async route => {
    const em = probeEm(route);
    const source = localSource(route);
    const measured = owner?.body && source && em
      ? await measureLocalFontSourceNormalLine(source, route.family, route.weight, route.style, em, owner)
      : null;
    return { resourceIdentity: route.resourceIdentity, weight: route.weight, style: route.style, measured };
  }));
}

function synthesizesStyledFaces(route: OfficeFontFallbackRoute): boolean {
  if (route.source !== 'local' || route.weight !== 400 || route.style !== 'normal') return false;
  const faces = findReferenceFontMetrics(route.requestedFamily);
  return !faces.some((face) => face.weight >= 600 && face.style === 'normal')
    || !faces.some((face) => face.style === 'italic');
}

function localSource(route: OfficeFontFallbackRoute): string | undefined {
  return route.source === 'local' && route.resourceIdentity.startsWith('office-local:local(')
    ? route.resourceIdentity.slice('office-local:'.length) : undefined;
}

/** Retain a native coverage probe alongside each measured paint resource.
 * Headless/Node consumers without Canvas keep the existing catalog fallback.
 * All font loads and DOM reads finish before immutable layout; coverage uses
 * a bounded cache of native shaping results, never per-segment DOM probes. */
export async function prepareDocxHostFontMetrics(
  loaded: Awaited<ReturnType<typeof loadOfficeFontFallbacks>>,
  fontSet: FontFaceSet | null,
  hostMeasurements?: readonly HostFontLineMeasurement[],
): Promise<Awaited<ReturnType<typeof loadOfficeFontFallbacks>>> {
  const owner = typeof document === 'undefined' ? null : document;
  if (!fontSet) return loaded;
  const context = owner?.body && owner.fonts === fontSet
    ? owner.createElement('canvas').getContext('2d') : null;
  // Measurements are synchronous; the coverage-face loads they need run in
  // parallel so preparation waits for the slowest load, not their sum.
  // Authored aliases can resolve to one paint face: retain its coverage probe
  // once per preparation, matching the single release on destroy.
  const probes = new Map<FontFace, Promise<FontFace | null>>();
  const entries = await Promise.all(Object.entries(loaded.routes).map(async ([key, route]) => {
    const face = loaded.faces.find(candidate => candidate.family === route.family);
    const em = probeEm(route);
    const transferred = hostMeasurements?.find(measurement =>
      measurement.resourceIdentity === route.resourceIdentity
      && measurement.weight === route.weight && measurement.style === route.style);
    const measured = face && em && context && owner?.body
      ? measureLocalFontNormalLine(face, fontSet, context, em, owner)
      : transferred?.measured ?? null;
    let metric = hostFontLineMetric(route, measured);
    const source = localSource(route);
    let probe: FontFace | null = null;
    // LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY needs glyph ownership for a
    // regular face whose family the catalog holds without a bold or italic
    // face: those styles paint as synthesized faces of this exact source.
    let coverageProbeFamily: string | undefined;
    if (face && source && synthesizesStyledFaces(route)) {
      let pending = probes.get(face);
      if (!pending) {
        pending = loadLocalFontCoverageProbe(face, source, fontSet);
        probes.set(face, pending);
      }
      coverageProbeFamily = (await pending)?.family;
    }
    if (metric !== route.metric) {
      if (face && source) {
        let pending = probes.get(face);
        if (!pending) {
          pending = loadLocalFontCoverageProbe(face, source, fontSet);
          probes.set(face, pending);
        }
        probe = await pending;
      }
      metric = probe
        ? Object.freeze({ ...metric, glyphCoverageProbeFamily: probe.family })
        // Without proof the new measured resource projection is unavailable.
        : route.metric;
    }
    return { key, route: { ...route, metric, ...(coverageProbeFamily ? { coverageProbeFamily } : {}) }, probe };
  }));
  // Route keys are authored face names; a null-prototype record keeps a face
  // named `__proto__` as an own entry instead of invoking the setter.
  const routes: typeof loaded.routes = nullRecord();
  for (const entry of entries) routes[entry.key] = entry.route;
  for (const pending of probes.values()) {
    const probe = await pending;
    if (probe) loaded.faces.push(probe);
  }
  return { ...loaded, routes };
}

export async function loadDocxOfficeFontFallbacks(
  ...args: Parameters<typeof loadOfficeFontFallbacks>
): ReturnType<typeof loadOfficeFontFallbacks> {
  const loaded = await loadOfficeFontFallbacks(...args);
  return prepareDocxHostFontMetrics(loaded, args[1] === undefined ? activeFontSet() : args[1]);
}
