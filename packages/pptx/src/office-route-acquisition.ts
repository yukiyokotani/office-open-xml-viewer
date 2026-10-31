import {
  normalizeLocalFontMetricFamily,
  type OfficeFontFallbackRequest,
  type OfficeFontFallbackRoute,
} from '@silurus/ooxml-core';
import { nullRecord, ownValue } from '@silurus/ooxml-core/internal/own-key';

type LoadedOfficeFontFallbacks = {
  faces: FontFace[];
  routes: Record<string, OfficeFontFallbackRoute>;
  installed?: string[];
};

export interface AcquiredOfficeRoutes {
  routes: Record<string, OfficeFontFallbackRoute>;
  /**
   * Normalized requested families whose installed face paints the authored
   * name: proven by a presence probe or a loaded route of any tuple, and not
   * shadowed by an application FontFace declaring the same family name.
   */
  installed: string[];
}

export type OfficeRouteCache = Map<string, Promise<LoadedOfficeFontFallbacks>>;

function requestKey(request: OfficeFontFallbackRequest): string {
  const family = normalizeLocalFontMetricFamily(request.family);
  return request.presenceOnly ? `presence:${family}` : `${family}:${request.weight ?? 400}:${request.style ?? 'normal'}`;
}

/**
 * Exact local() routes for the requested authored tuples, shared across slides.
 * Each tuple is acquired once per cache (family, weight and style all key it).
 * Tuples not yet cached are loaded together in one loader call, so a slide keeps
 * the loader's per-call source, concurrency and deadline bounds instead of one
 * unbounded call per tuple.
 */
export async function acquireOfficeRoutes(
  cache: OfficeRouteCache,
  requests: readonly OfficeFontFallbackRequest[],
  load: (requests: OfficeFontFallbackRequest[]) => Promise<LoadedOfficeFontFallbacks>,
  fontSet: FontFaceSet | null = null,
): Promise<AcquiredOfficeRoutes> {
  const fresh = new Map<string, OfficeFontFallbackRequest>();
  for (const request of requests) {
    const key = requestKey(request);
    if (!cache.has(key) && !fresh.has(key)) fresh.set(key, request);
  }
  if (fresh.size > 0) {
    const batch = load([...fresh.values()]);
    for (const key of fresh.keys()) cache.set(key, batch);
  }
  const pending = new Set(requests.map((request) => cache.get(requestKey(request))!));
  const loaded = await Promise.all(pending);
  // Route keys are authored face names (`__proto__`, `constructor`, ...): merge
  // into and collect from null-prototype records and read own entries only, so
  // no name is answered by, or rewrites, Object.prototype.
  const available: Record<string, OfficeFontFallbackRoute> = Object.assign(
    nullRecord<OfficeFontFallbackRoute>(), ...loaded.map((entry) => entry.routes));
  // A batch may hold routes for tuples another slide requested; return only
  // this call's tuples so what the renderer knows never depends on history.
  const installedAnywhere = new Set(loaded.flatMap((entry) => entry.installed ?? []));
  const routes = nullRecord<OfficeFontFallbackRoute>();
  const installed = new Set<string>();
  for (const request of requests) {
    const family = normalizeLocalFontMetricFamily(request.family);
    if (request.presenceOnly) {
      if (installedAnywhere.has(family)) installed.add(family);
      continue;
    }
    const key = routeKeyFor(request);
    const route = ownValue(available, key);
    if (route) {
      routes[key] = route;
      installed.add(family);
    }
  }
  // An application FontFace with the authored family name shadows the system
  // family for every tuple (CSS font matching), so system presence cannot
  // describe the face that paints. Checked per call: faces can be added later.
  for (const family of declaredFamilies(fontSet)) installed.delete(family);
  return { routes, installed: [...installed].sort() };
}

/** Families named by FontFaces declared in the set. Office routes register
 * under private aliases, so they never shadow an authored family here. */
function declaredFamilies(fontSet: FontFaceSet | null): Set<string> {
  if (!fontSet || typeof fontSet[Symbol.iterator] !== 'function') return new Set();
  return new Set([...fontSet].map((face) =>
    normalizeLocalFontMetricFamily(face.family.replace(/^(['"])(.*)\1$/u, '$2'))));
}

/** The core office-fallback route key (regular tuples key by family alone). */
function routeKeyFor(request: OfficeFontFallbackRequest): string {
  const family = normalizeLocalFontMetricFamily(request.family);
  const weight = request.weight ?? 400;
  const style = request.style ?? 'normal';
  return weight === 400 && style === 'normal' ? family : `${family}:${weight}:${style}`;
}

/** Every distinct load in the cache (one batch backs several tuples). */
export function distinctOfficeRouteLoads(cache: OfficeRouteCache): Promise<LoadedOfficeFontFallbacks>[] {
  return [...new Set(cache.values())];
}
