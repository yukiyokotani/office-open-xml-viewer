import { hadLocalFontProbeTimeout, loadLocalFontMetrics, normalizeLocalFontMetricFamily, unloadLocalFontMetrics } from './local-metrics.js';
import { activeFontSet } from './font-registry.js';
import { findReferenceFontMetrics } from './reference-font-metrics.js';
import type { ResolvedFontMetric } from './resource-metrics.js';

/** A document requests an authored face and style. This API retains its name
 * for compatibility, but library policy no longer supplies font bytes. */
export interface OfficeFontFallbackRequest {
  family: string;
  weight?: number;
  style?: 'normal' | 'italic';
  /**
   * Only report whether ANY face of this family is installed, independent of
   * the styles a document uses (exact local() probe of the family name and
   * every catalogued alias). No route or metric is created; see
   * {@link LoadedOfficeFontFallbacks.installed}.
   */
  presenceOnly?: boolean;
}

export interface OfficeFontFallbackRoute {
  requestedFamily: string;
  family: string;
  /** `substitute` remains in the public type for existing callers; this loader
   * now emits only positively loaded local faces. */
  source: 'local' | 'substitute';
  resourceIdentity: string;
  weight: number;
  style: 'normal' | 'italic';
  metric: ResolvedFontMetric;
  /** Family of this exact source's size-adjust:0% registration, when a format
   * prepared one: a scalar it owns shapes with zero advance. */
  coverageProbeFamily?: string;
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
  /** Normalized families of `presenceOnly` requests whose face is installed. */
  installed?: string[];
  /** Normalized families of `presenceOnly` requests whose presence probe
   * completed (installed or not). A family past the source cap or the
   * deadline is absent: it is unproven, not missing. */
  presenceChecked?: string[];
}

type Tuple = Readonly<{
  family: string;
  weight: number;
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
  // Any OpenType usWeightClass value; the catalog lookup below admits only
  // tuples with a catalogued face (named cuts such as a 300 Light face too).
  if (!Number.isInteger(weight) || weight < 1 || weight > 1000
    || (style !== 'normal' && style !== 'italic')) return undefined;
  const profiles = findReferenceFontMetrics(family, { weight, style });
  if (profiles.length === 0) return undefined;
  // An alias shared by regular and bold is a family name, not proof of a
  // styled face. Prefer full/PostScript aliases unique to this tuple. For the
  // regular face only, a family name is also an exact local() candidate.
  // A named cut ("Arial Black") can resolve to a profile of another catalog
  // family ("Arial" 900), whose aliases include that family's own name. Collect
  // the other tuples of every matched catalog family too, so local("Arial") can
  // never stand in for the 900 cut and load the regular face.
  const catalogFamilies = new Set([family, ...profiles.map((profile) => profile.family)]);
  const otherAliases = new Set([...catalogFamilies].flatMap((name) => findReferenceFontMetrics(name))
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

/** Whether a loaded CSS face advertises the requested numeric weight. A
 * variable face may cover several WML tuples, while a scalar face covers one. */
export function fontFaceWeightCovers(descriptor: string, weight: number): boolean {
  const normalized = descriptor.trim().toLowerCase();
  if (normalized === 'normal') return weight === 400;
  if (normalized === 'bold') return weight === 700;
  const range = /^(\d+)(?:\s+(\d+))?$/u.exec(normalized);
  if (!range) return false;
  const lower = Number(range[1]);
  const upper = Number(range[2] ?? range[1]);
  return lower <= weight && weight <= upper;
}

function loadedFaceCoversTuple(face: FontFace, tuple: Tuple): boolean {
  if (face.status !== 'loaded'
    || normalizeLocalFontMetricFamily(face.family.replace(/^(['"])(.*)\1$/u, '$2'))
      !== normalizeLocalFontMetricFamily(tuple.family)
    || face.style.trim().toLowerCase() !== tuple.style) return false;
  return fontFaceWeightCovers(face.weight, tuple.weight);
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
  const presence = requests.filter((request) => request.presenceOnly).map((request) => request.family);
  const [loaded, installed] = await Promise.all([
    loadOfficeFontRoutes(requests.filter((request) => !request.presenceOnly), targetFontSet),
    presence.length > 0
      ? probeInstalledFontFamilies(presence, targetFontSet)
      : Promise.resolve(undefined),
  ]);
  if (!installed) return loaded;
  return {
    ...loaded,
    ...(installed.installed.size > 0 ? { installed: [...installed.installed].sort() } : {}),
    presenceChecked: [...installed.checked].sort(),
  };
}

async function loadOfficeFontRoutes(
  requests: readonly OfficeFontFallbackRequest[],
  targetFontSet: FontFaceSet | null,
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
  // Null prototype: keys are document font names (e.g. `__proto__`).
  const metrics = Object.assign(Object.create(null), ...loaded.flatMap((result) => result ? [result.metrics] : [])) as Record<string, ResolvedFontMetric>;
  const routes = Object.create(null) as Record<string, OfficeFontFallbackRoute>;
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

/** Exact local() names that identify ANY installed face of a family: the
 * family name (a regular face's full name in practice) plus every catalogued
 * full/PostScript alias of every weight and style. Presence of the family is a
 * question separate from which styles a document uses. */
function familyPresenceLocalNames(family: string): string[] {
  // A named cut ("Arial Black") can resolve to profiles of a broader catalog
  // family ("Arial" 900). Those profiles' aliases may name the broader family
  // or its other cuts, which prove only that family installed, never the cut.
  // As in tupleFor, keep a borrowed profile's aliases only when no other tuple
  // of its catalog family shares them.
  const requested = normalizeLocalFontMetricFamily(family);
  const names = findReferenceFontMetrics(family).flatMap((profile) => {
    if (normalizeLocalFontMetricFamily(profile.family) === requested) return profile.aliases;
    const others = new Set(findReferenceFontMetrics(profile.family)
      .filter((other) => other.weight !== profile.weight || other.style !== profile.style)
      .flatMap((other) => other.aliases.map(normalizeLocalFontMetricFamily)));
    return profile.aliases.filter((alias) => !others.has(normalizeLocalFontMetricFamily(alias)));
  });
  return [...new Set([family, ...names])];
}

/**
 * Which authored families the host has installed: a family counts when any of
 * its faces loads through an exact local() name (see
 * {@link familyPresenceLocalNames}). Each family is an independent probe. Probe
 * faces are released as soon as each probe completes, and no route or metric
 * is created, so the authored family keeps its ordinary CSS resolution.
 *
 * Callers use this only to decline an optional web substitute for a family
 * that is actually present. Library policy: the authored family names the
 * face to use, so an installed authored face is never displaced by a
 * substitute. The probes share the catalogued preflight's bounds: at most four
 * run at once, no probe starts after the document deadline, and the result is
 * returned at the deadline with the probes completed so far. A probe that
 * completes later only releases its face. A timeout or a missing face is not
 * proof of absence; it only leaves the substitute enabled.
 */
/** Outcome of one bounded family presence probe: normalized families found
 * installed, and every normalized family whose probe completed (installed or
 * not). A family over the source cap or past the deadline is in neither. */
export interface FontFamilyPresence {
  readonly installed: ReadonlySet<string>;
  readonly checked: ReadonlySet<string>;
}

async function probeInstalledFontFamilies(
  families: readonly string[],
  targetFontSet: FontFaceSet | null,
): Promise<FontFamilyPresence> {
  const unique = [...new Map(families
    .map((family) => family.trim())
    .filter(Boolean)
    .map((family) => [normalizeLocalFontMetricFamily(family), family] as const)).values()]
    .slice(0, MAX_PREFLIGHT_SOURCES);
  const installed = new Set<string>();
  const checked = new Set<string>();
  if (!targetFontSet || typeof FontFace === 'undefined' || unique.length === 0) return { installed, checked };
  let next = 0;
  let accepting = true;
  const workers = Array.from({ length: Math.min(4, unique.length) }, async () => {
    while (accepting && next < unique.length) {
      const family = unique[next++];
      const result = await loadLocalFontMetrics(
        [{ family, localNames: familyPresenceLocalNames(family) }],
        targetFontSet,
      );
      unloadLocalFontMetrics(result.faces);
      if (!accepting) continue;
      checked.add(normalizeLocalFontMetricFamily(family));
      if (Object.keys(result.metrics).length > 0) installed.add(normalizeLocalFontMetricFamily(family));
    }
  });
  let deadline: ReturnType<typeof setTimeout> | undefined;
  await Promise.race([
    Promise.allSettled(workers),
    new Promise<void>((resolve) => { deadline = setTimeout(resolve, PREFLIGHT_DEADLINE_MS); }),
  ]);
  if (deadline !== undefined) clearTimeout(deadline);
  accepting = false;
  return { installed: new Set(installed), checked: new Set(checked) };
}

/**
 * Presence of every family in `families`, probed in batches of the per-call
 * source cap (the same `local()` probe as presence-only fallback requests).
 * Used where a font source must not shadow an installed face, so every
 * demanded family is probed rather than only the first batch. Each batch keeps
 * its own deadline; a family whose probe did not complete is absent from
 * `checked`, and callers must not treat it as proven missing.
 */
export async function probeFontFamilyPresence(
  families: readonly string[],
  targetFontSet: FontFaceSet | null = activeFontSet(),
): Promise<FontFamilyPresence> {
  const unique = [...new Map(families
    .map((family) => family.trim())
    .filter(Boolean)
    .map((family) => [normalizeLocalFontMetricFamily(family), family] as const)).values()];
  const installed = new Set<string>();
  const checked = new Set<string>();
  for (let start = 0; start < unique.length; start += MAX_PREFLIGHT_SOURCES) {
    const batch = await probeInstalledFontFamilies(unique.slice(start, start + MAX_PREFLIGHT_SOURCES), targetFontSet);
    for (const family of batch.installed) installed.add(family);
    for (const family of batch.checked) checked.add(family);
  }
  return { installed, checked };
}
