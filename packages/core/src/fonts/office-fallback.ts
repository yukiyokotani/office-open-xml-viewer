import { hadLocalFontProbeTimeout, loadLocalFontMetrics, normalizeLocalFontMetricFamily, unloadLocalFontMetrics } from './local-metrics.js';
import { activeFontSet } from './preload.js';
import { findReferenceFontMetrics } from './reference-font-metrics.js';
import type { ResolvedFontMetric } from './resource-metrics.js';

/** A document requests an authored face and style. This API retains its name
 * for compatibility, but library policy no longer supplies font bytes. */
export interface OfficeFontFallbackRequest {
  family: string;
  weight?: number;
  style?: 'normal' | 'italic';
}

export interface OfficeFontFallbackRoute {
  requestedFamily: string;
  family: string;
  /** `substitute` remains in the public type for existing callers; this loader
   * now emits only positively loaded local faces. */
  source: 'local' | 'substitute';
  resourceIdentity: string;
  weight: 400 | 700;
  style: 'normal' | 'italic';
  metric: ResolvedFontMetric;
}

export interface LoadedOfficeFontFallbacks {
  /** One retention per loaded local source (aliases may share it). Release when
   * the owning document closes. */
  faces: FontFace[];
  /** Normalized family key, with :weight:style for non-regular tuples. */
  routes: Record<string, OfficeFontFallbackRoute>;
  /** Tuples whose local() preflight completed without hitting a timeout. A
   * missing route means the attempted sources did not load, not proof that no
   * system installation exists. Budget/deadline omissions are absent. */
  checked: string[];
}

type Tuple = Readonly<{
  family: string;
  weight: 400 | 700;
  style: 'normal' | 'italic';
  localNames: readonly string[];
}>;

// Resource-governance limits for optional local() preflight, not Office layout
// coefficients. A document with hundreds of authored faces must not block its
// first paint on one 15-second FontFace.load() ceiling per four-face batch.
const MAX_PREFLIGHT_SOURCES = 32;
const PREFLIGHT_DEADLINE_MS = 8_000;

function tupleFor(request: OfficeFontFallbackRequest): Tuple | undefined {
  const family = request.family.trim();
  if (!family) return undefined;
  const weight = request.weight ?? 400;
  const style = request.style ?? 'normal';
  if ((weight !== 400 && weight !== 700) || (style !== 'normal' && style !== 'italic')) return undefined;
  const profiles = findReferenceFontMetrics(family, { weight, style });
  if (profiles.length === 0) return undefined;
  // An alias shared by regular and bold is a family name, not proof of a
  // styled face. Prefer full/PostScript aliases unique to this tuple. For the
  // regular face only, a family name is also an exact local() candidate.
  const otherAliases = new Set(findReferenceFontMetrics(family)
    .filter((profile) => profile.weight !== weight || profile.style !== style)
    .flatMap((profile) => profile.aliases.map(normalizeLocalFontMetricFamily)));
  const distinct = [...new Set(profiles.flatMap((profile) => profile.aliases))];
  const localNames = distinct.filter((alias) => !otherAliases.has(normalizeLocalFontMetricFamily(alias)));
  if (weight === 400 && style === 'normal') {
    for (const profile of profiles) {
      if (!localNames.some((alias) => normalizeLocalFontMetricFamily(alias)
        === normalizeLocalFontMetricFamily(profile.family))) localNames.push(profile.family);
    }
  }
  if (localNames.length === 0) return undefined;
  return { family, weight, style, localNames };
}

function routeKey(tuple: Tuple): string {
  const family = normalizeLocalFontMetricFamily(tuple.family);
  return tuple.weight === 400 && tuple.style === 'normal'
    ? family : `${family}:${tuple.weight}:${tuple.style}`;
}

function loadedFaceCoversTuple(face: FontFace, tuple: Tuple): boolean {
  if (face.status !== 'loaded'
    || normalizeLocalFontMetricFamily(face.family.replace(/^(['"])(.*)\1$/u, '$2'))
      !== normalizeLocalFontMetricFamily(tuple.family)
    || face.style.trim().toLowerCase() !== tuple.style) return false;
  const descriptor = face.weight.trim().toLowerCase();
  if (descriptor === 'normal') return tuple.weight === 400;
  if (descriptor === 'bold') return tuple.weight === 700;
  const range = /^(\d+)(?:\s+(\d+))?$/u.exec(descriptor);
  if (!range) return false;
  const lower = Number(range[1]);
  const upper = Number(range[2] ?? range[1]);
  return lower <= tuple.weight && tuple.weight <= upper;
}

/** ECMA-376 font names identify requested families, not transferable font
 * resources. Probe only catalogued exact local() names for document-used
 * tuples. A regular face does not establish bold or italic. CSS local() exposes
 * no installed bytes, so the route records identity without claiming resource
 * metrics. Missing tuples keep the authored name and generic fallback; this
 * path neither packages fonts nor makes a network request. */
export async function loadOfficeFontFallbacks(
  requests: readonly OfficeFontFallbackRequest[],
  targetFontSet: FontFaceSet | null = activeFontSet(),
): Promise<LoadedOfficeFontFallbacks> {
  if (!targetFontSet || typeof FontFace === 'undefined') return { faces: [], routes: {}, checked: [] };
  // A loaded application face wins only its declared style/weight tuple. A
  // regular face must not suppress exact-local Bold or Italic, and a face still
  // loading cannot supply stable geometry to this document's layout snapshot.
  const declared = typeof targetFontSet[Symbol.iterator] === 'function'
    ? [...targetFontSet] : [];
  const tuples = [...new Map(requests.map(tupleFor).filter((tuple): tuple is Tuple => !!tuple)
    .filter((tuple) => !declared.some((face) => loadedFaceCoversTuple(face, tuple)))
    .map((tuple) => [routeKey(tuple), tuple])).values()];
  if (tuples.length === 0) return { faces: [], routes: {}, checked: [] };
  // A document can name many catalogued faces. Keep failed local() loads from
  // serializing startup by probing at most four independent source tuples at a
  // time. Aliases of one source share one registration/refcount and one load.
  const groups = new Map<string, Tuple[]>();
  for (const tuple of tuples) {
    const signature = JSON.stringify([tuple.localNames, tuple.weight, tuple.style]);
    const group = groups.get(signature) ?? [];
    group.push(tuple);
    groups.set(signature, group);
  }
  const jobs = [...groups.values()].slice(0, MAX_PREFLIGHT_SOURCES);
  const loaded = new Array<Awaited<ReturnType<typeof loadLocalFontMetrics>>>(jobs.length);
  let nextJob = 0;
  let accepting = true;
  const workers = Array.from({ length: Math.min(4, jobs.length) }, async () => {
    while (accepting && nextJob < jobs.length) {
      const index = nextJob++;
      const result = await loadLocalFontMetrics(jobs[index].map((tuple) => ({
        family: tuple.family, localNames: tuple.localNames,
        weight: tuple.weight, style: tuple.style,
      })), targetFontSet);
      if (accepting) loaded[index] = result;
      else unloadLocalFontMetrics(result.faces);
    }
  });
  let deadline: ReturnType<typeof setTimeout> | undefined;
  const settled = await Promise.race([
    Promise.allSettled(workers),
    new Promise<null>((resolve) => {
      deadline = setTimeout(() => resolve(null), PREFLIGHT_DEADLINE_MS);
    }),
  ]);
  if (deadline !== undefined) clearTimeout(deadline);
  accepting = false;
  const failure = settled?.find((result): result is PromiseRejectedResult => result.status === 'rejected');
  if (failure) {
    unloadLocalFontMetrics(loaded.flatMap((result) => result?.faces ?? []));
    throw failure.reason;
  }
  const metrics = Object.assign({}, ...loaded.flatMap((result) => result ? [result.metrics] : [])) as Record<string, ResolvedFontMetric>;
  const routes: Record<string, OfficeFontFallbackRoute> = {};
  for (const tuple of tuples) {
    const key = routeKey(tuple);
    const metric = metrics[key];
    if (!metric) continue;
    const resourceIdentity = `office-local:${metric.sourceIdentity ?? tuple.localNames.join(',')}`;
    routes[key] = {
      requestedFamily: tuple.family, family: metric.family, source: 'local',
      resourceIdentity, weight: tuple.weight, style: tuple.style,
      metric: { ...metric, sourceIdentity: resourceIdentity },
    };
  }
  return {
    faces: loaded.flatMap((result) => result?.faces ?? []), routes,
    checked: loaded.flatMap((result, index) => result && !hadLocalFontProbeTimeout(result)
      ? jobs[index].map(routeKey) : []),
  };
}

export function unloadOfficeFontFallbacks(faces: Iterable<FontFace>): void {
  unloadLocalFontMetrics(faces);
}
