import { cjkLangFromLanguage, type CjkLang, type ResolvedFontMetric } from '@silurus/ooxml-core';
import { quoteCssFontFamily } from '@silurus/ooxml-core/internal/css-font-family';
import type { LayoutDiagnostic } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import { ownValue } from '@silurus/ooxml-core/internal/own-key';
import {
  createCanvasFontRoute,
  type CanvasFontRoute,
  type FontSubstituteScript,
} from '@silurus/ooxml-core';

export type FontResolutionSource = 'embedded' | 'local' | 'css' | 'google' | 'substitute' | 'native' | 'generic';
export type FontStyle = 'normal' | 'italic';

export interface FontRequest {
  readonly cjkFallback?: CjkLang;
  readonly language?: string;
  readonly requestedFamily?: string | null;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly weight?: number;
  readonly style?: FontStyle;
  /** Script that the requested text belongs to, when it belongs to one that a
   * scoped visual substitute covers. A scoped inventory face answers only such
   * a request; any other request resolves as if that face did not exist. */
  readonly script?: FontSubstituteScript;
}

export interface FontResolution {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly route: CanvasFontRoute;
  readonly source: FontResolutionSource;
  /** Identity of the registered resource that supplied this face, when known. */
  readonly resourceIdentity?: string;
  readonly weight: number;
  readonly style: FontStyle;
  readonly diagnostics: readonly LayoutDiagnostic[];
  readonly genericFamily: 'serif' | 'sans-serif' | 'monospace';
  /** Native resolution of a family that has an installed local face in another
   * tuple: CSS font matching paints that installed family (synthesizing
   * weight/style), so the authored face is present on the host, not absent. */
  readonly installedFamily?: true;
  /** Native route of an absent authored face: the measured normal line of the
   * route's first available face (native-route-line-metrics.ts). Admitted
   * only for clusters that face is proven to draw. */
  readonly routeLineMetric?: NativeRouteLineMetric;
}

/** Plain, array-free projection of a measured native route line, so a
 * resolution stays structured-clone/deep-readonly plain data. */
export type NativeRouteLineMetric = Readonly<Pick<ResolvedFontMetric,
  'family' | 'requestedFamily' | 'weight' | 'style' | 'farEastCodePage' | 'requiresGlyphCoverage'
  | 'nativeRoutePrimaryFamily' | 'lineHeightRatio' | 'designAscentRatio' | 'designDescentRatio'
  | 'lineGapRatio'>>;

export interface FontResolver {
  readonly fingerprint: string;
  resolve(request: Readonly<FontRequest>): FontResolution;
  /** Registry scope used only to delimit the shared run-context proof rule.
   * It does not authorize changing a scalar's slot or selected face. */
  configuredSubstituteScript?(requestedFamily: string | null | undefined): FontSubstituteScript | undefined;
  /** Script of the scoped substitute that actually wins this resource tuple.
   * Authored embedded/local/installed faces retain §17.3.2.26 scalar slots. */
  scopedSubstituteScript?(requestedFamily: string | null | undefined, weight?: number, style?: FontStyle): FontSubstituteScript | undefined;
  /** Installed Office faces of an authored family loaded through an exact
   * local() name, each under its own private single-face family. */
  exactLocalFaces?(requestedFamily: string): readonly Readonly<{ resolvedFamily: string; weight: number; style: FontStyle; coverageProbeFamily?: string }>[];
}

export interface FontInventoryFace {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly source: Exclude<FontResolutionSource, 'generic' | 'native'>;
  readonly resourceIdentity?: string;
  readonly weight?: number;
  readonly style?: FontStyle;
  /** Visual substitute limited to one script (core `substitute-script.ts`). */
  readonly script?: FontSubstituteScript;
  /** Size-adjust:0% registration of this exact source (glyph ownership). */
  readonly coverageProbeFamily?: string;
}

export interface FontResolverOptions {
  /** Immutable regional routes; their concrete outputs participate in cache identity. */
  readonly regionalFamilyLists?: Partial<Record<CjkLang, Readonly<Record<string, string>>>>;
  /** Stable DOCX fallback routes derived from document metadata and rendered faces. */
  readonly nativeFamilyLists?: Readonly<Record<string, string>>;
  /**
   * Authored families (normalized) whose web substitute is script-scoped (core
   * substitute-script.ts), with the substitute families of that script. For a
   * request not in that script, those families are removed from EVERY route
   * of the family (explicit, native and regional CSS lists), so Canvas can
   * never select them for other characters, even as a CSS fallback.
   */
  /**
   * Normalized authored families proven installed on the host by the family
   * presence probe (any face of the family), independent of which styled
   * tuples the document requests. CSS font matching paints an installed family
   * for every tuple, synthesizing weight/style.
   */
  readonly installedFamilies?: readonly string[];
  /**
   * Normalized authored families declared by a loaded application FontFace,
   * in any descriptor. Such a face shadows the system family for every tuple,
   * so a system installation never describes what paints.
   */
  readonly shadowedFamilies?: readonly string[];
  readonly scriptScopedFamilies?: Readonly<Record<string, Readonly<{
    script: FontSubstituteScript;
    substituteFamilies: readonly string[];
  }>>>;
  /** Measured native route lines keyed by {@link nativeRouteMetricKey}. They
   * participate in cache identity: main and worker resolve the same table. */
  readonly nativeRouteMetrics?: Readonly<Record<string, NativeRouteLineMetric>>;
}

/** Exact native route identity: the complete CSS family list Canvas receives
 * and the normalized tuple. */
export function nativeRouteMetricKey(familyList: string, weight: number, style: FontStyle): string {
  return JSON.stringify([familyList, weight, style]);
}

// DOCX family keys use the locale-independent Unicode lowercase mapping. It is
// identical to toLocaleLowerCase('en-US') for every input (only lt/tr/az tailor
// lowercasing), but avoids per-call locale resolution on this per-segment path.
function normalizeFamily(value: string): string {
  return value.trim().toLowerCase();
}

function normalizedWeight(value: number | undefined): number {
  if (value == null || !Number.isFinite(value)) return 400;
  return Math.min(900, Math.max(100, Math.round(value / 100) * 100));
}

function freezeResolution(value: FontResolution): FontResolution {
  return Object.freeze({ ...value, diagnostics: Object.freeze([...value.diagnostics]) });
}

function cssFamilyList(family: string, generic: FontResolution['genericFamily']): string {
  return `${quoteCssFontFamily(family)}, ${generic}`;
}

/**
 * Snapshot the font inventory used by one document. ECMA-376 §17.8.2 leaves
 * the substitution algorithm implementation-defined, so a substituted or
 * generic result is carried as an explicit diagnostic instead of being hidden
 * in paragraph geometry.
 */
export function createFontResolver(
  inventory: readonly FontInventoryFace[],
  options: Readonly<FontResolverOptions> = {},
): FontResolver {
  const sourcePriority: Readonly<Record<FontInventoryFace['source'], number>> = {
    embedded: 0,
    local: 1,
    css: 2,
    google: 3,
    substitute: 4,
  };
  const faces = inventory
    .filter((face) => face.requestedFamily.trim() && face.resolvedFamily.trim())
    .map((face) => Object.freeze({
      ...face,
      weight: normalizedWeight(face.weight),
      style: face.style ?? 'normal',
    }))
    .sort((a, b) => {
      const family = normalizeFamily(a.requestedFamily).localeCompare(normalizeFamily(b.requestedFamily));
      return family || sourcePriority[a.source] - sourcePriority[b.source]
        || a.resolvedFamily.localeCompare(b.resolvedFamily)
        || a.weight - b.weight
        || a.style.localeCompare(b.style);
    });
  const byFamily = new Map<string, (typeof faces)[number][]>();
  for (const face of faces) {
    const key = normalizeFamily(face.requestedFamily);
    byFamily.set(key, [...(byFamily.get(key) ?? []), face]);
  }
  const nativeFamilyLists = Object.freeze(Object.fromEntries(
    Object.entries(options.nativeFamilyLists ?? {})
      .filter(([family, familyList]) => family.trim() && familyList.trim())
      .map(([family, familyList]) => [normalizeFamily(family), familyList] as const)
      .sort(([a], [b]) => a.localeCompare(b)),
  ));
  const regionalFamilyLists = Object.freeze(Object.fromEntries(
    Object.entries(options.regionalFamilyLists ?? {}).map(([region, lists]) => [
      region,
      Object.freeze(Object.fromEntries(Object.entries(lists)
        .map(([family, list]) => [normalizeFamily(family), list])
        .sort(([a], [b]) => a.localeCompare(b)))),
    ]).sort(([a], [b]) => String(a).localeCompare(String(b))),
  )) as Readonly<Partial<Record<CjkLang, Readonly<Record<string, string>>>>>;
  const familyListFor = (
    family: string,
    language: string | undefined,
    fallback: CjkLang | undefined,
  ): string | undefined => {
    const region = cjkLangFromLanguage(language) ?? fallback;
    // Both maps are keyed by authored face names: read own entries only.
    const regional = region ? regionalFamilyLists[region] : undefined;
    return (regional ? ownValue(regional, normalizeFamily(family)) : undefined)
      ?? ownValue(nativeFamilyLists, normalizeFamily(family));
  };
  const scriptScopedFamilies = Object.freeze(Object.fromEntries(
    Object.entries(options.scriptScopedFamilies ?? {})
      .map(([family, scope]) => [normalizeFamily(family), Object.freeze({
        script: scope.script,
        substituteFamilies: Object.freeze([...scope.substituteFamilies]
          .map((name) => normalizeFamily(name)).sort()),
      })] as const)
      .sort(([a], [b]) => a.localeCompare(b)),
  ));
  // Remove a scoped substitute's families from a CSS family list. Names never
  // contain commas in these generated lists; each entry is a quoted family or
  // a CSS generic keyword.
  const withoutScopedSubstitutes = (
    familyList: string,
    requestedFamily: string,
    script: FontSubstituteScript | undefined,
  ): string => {
    const scope = ownValue(scriptScopedFamilies, normalizeFamily(requestedFamily));
    if (!scope || scope.script === script) return familyList;
    return familyList.split(',').map((entry) => entry.trim())
      .filter((entry) => !scope.substituteFamilies.includes(
        normalizeFamily(entry.replace(/^"(.*)"$/u, '$1').replaceAll('\\"', '"').replaceAll('\\\\', '\\')),
      ))
      .join(', ');
  };
  const installedFamilies = [...new Set((options.installedFamilies ?? []).map(normalizeFamily))].sort();
  const installedFamilySet = new Set(installedFamilies);
  const shadowedFamilies = [...new Set((options.shadowedFamilies ?? []).map(normalizeFamily))].sort();
  const shadowedFamilySet = new Set(shadowedFamilies);
  const nativeRouteMetrics = Object.freeze(Object.fromEntries(
    Object.entries(options.nativeRouteMetrics ?? {}).sort(([a], [b]) => a.localeCompare(b)),
  ));
  const fingerprint = stableFingerprint('fonts', {
    faces, nativeFamilyLists, regionalFamilyLists, scriptScopedFamilies,
    ...(installedFamilies.length > 0 ? { installedFamilies } : {}),
    ...(shadowedFamilies.length > 0 ? { shadowedFamilies } : {}),
    ...(Object.keys(nativeRouteMetrics).length > 0 ? { nativeRouteMetrics } : {}),
  });

  // Every shaped script span, line segment and east-Asian line floor asks for a
  // resolution, and each fresh answer carried its own copy of the complete CSS
  // family list plus a percent-encoded route fingerprint (together ~1.7 KB for
  // a registered local face). A long document retained hundreds of thousands
  // of value-identical copies across the text caches and the retained layout.
  // `resolve` is a pure function of the request fields below and this
  // resolver's immutable snapshot, and every resolution is deep-frozen, so
  // equal requests share one answer. The raw fields are the key (not their
  // normalized forms), so each entry is exactly what that request computed.
  // The bound only caps pathological documents; a miss recomputes.
  const resolutionMemoLimit = 4096;
  const resolutions = new Map<string, FontResolution>();
  // One resource selector serves resolution and the scope gate. Inventory
  // priority, exact tuple matching and script admission cannot diverge.
  const selectedFace = (family: string, weight: number, style: FontStyle, script?: FontSubstituteScript) =>
    (byFamily.get(normalizeFamily(family)) ?? []).find((candidate) =>
      candidate.weight === weight && candidate.style === style
      && (candidate.script === undefined || candidate.script === script));
  const resolveUncached = (request: Readonly<FontRequest>): FontResolution => {
    const requestedFamily = request.requestedFamily?.trim() || request.genericFamily || 'sans-serif';
    const weight = normalizedWeight(request.weight);
    const style = request.style ?? 'normal';
    const face = selectedFace(requestedFamily, weight, style, request.script);
    if (face) {
      const diagnostics: LayoutDiagnostic[] = face.source === 'substitute'
        ? [{
            code: 'UNSUPPORTED_FEATURE',
            severity: 'warning',
            message: `ECMA-376 §17.8.2 implementation-dependent font substitution: ${requestedFamily} resolved to ${face.resolvedFamily}`,
          }]
        : [];
      const fallbackList = familyListFor(requestedFamily, request.language, request.cjkFallback);
      const familyList = fallbackList
        ? `${quoteCssFontFamily(face.resolvedFamily)}, ${fallbackList}`
        : cssFamilyList(face.resolvedFamily, request.genericFamily ?? 'sans-serif');
      return freezeResolution({
        requestedFamily,
        resolvedFamily: face.resolvedFamily,
        route: createCanvasFontRoute(
          withoutScopedSubstitutes(familyList, requestedFamily, request.script),
          'registered',
        ),
        source: face.source,
        ...(face.resourceIdentity === undefined ? {} : { resourceIdentity: face.resourceIdentity }),
        weight,
        style,
        diagnostics,
        genericFamily: request.genericFamily ?? 'sans-serif',
      });
    }

    const generic = request.genericFamily ?? 'sans-serif';
    const authored = request.requestedFamily?.trim();
    if (authored) {
      const familyList = familyListFor(authored, request.language, request.cjkFallback)
        ?? cssFamilyList(authored, generic);
      // Only a system installation (the family presence probe, or a face found
      // by the office local() loader) shows that CSS matching paints this
      // family; an application resource under a private alias does not. An
      // application @font-face declaring the authored family shadows the
      // system family for every tuple, so the system face cannot describe it.
      const candidates = byFamily.get(normalizeFamily(authored)) ?? [];
      const installedFamily = !shadowedFamilySet.has(normalizeFamily(authored))
        && !candidates.some((candidate) => candidate.source === 'css')
        && (installedFamilySet.has(normalizeFamily(authored))
          || candidates.some((candidate) => candidate.source === 'local'
            && candidate.resourceIdentity?.startsWith('office-local:') === true));
      const route = createCanvasFontRoute(
        withoutScopedSubstitutes(familyList, authored, request.script),
        'native',
      );
      // An installed authored face paints this route itself (catalog path);
      // otherwise the route's first available face paints, and its measured
      // normal line (when one was admitted for this exact list) describes it.
      const routeLineMetric = installedFamily
        ? undefined
        : ownValue(nativeRouteMetrics, nativeRouteMetricKey(route.familyList, weight, style));
      return freezeResolution({
        ...(installedFamily ? { installedFamily: true as const } : {}),
        ...(routeLineMetric ? { routeLineMetric } : {}),
        requestedFamily,
        resolvedFamily: authored,
        route,
        source: 'native',
        weight,
        style,
        diagnostics: [],
        genericFamily: generic,
      });
    }
    return freezeResolution({
      requestedFamily,
      resolvedFamily: generic,
      route: createCanvasFontRoute(
        familyListFor(generic, request.language, request.cjkFallback) ?? generic,
        'generic',
      ),
      source: 'generic',
      weight,
      style,
      diagnostics: [],
      genericFamily: generic,
    });
  };

  return Object.freeze({
    fingerprint,
    configuredSubstituteScript(requestedFamily: string | null | undefined): FontSubstituteScript | undefined {
      return requestedFamily ? ownValue(scriptScopedFamilies, normalizeFamily(requestedFamily))?.script : undefined;
    },
    exactLocalFaces(requestedFamily: string) {
      return (byFamily.get(normalizeFamily(requestedFamily)) ?? [])
        .filter((face) => face.source === 'local' && face.resourceIdentity?.startsWith('office-local:') === true)
        .map((face) => ({
          resolvedFamily: face.resolvedFamily, weight: face.weight, style: face.style,
          ...(face.coverageProbeFamily ? { coverageProbeFamily: face.coverageProbeFamily } : {}),
        }));
    },
    scopedSubstituteScript(requestedFamily: string | null | undefined, weight?: number, style?: FontStyle): FontSubstituteScript | undefined {
      const family = requestedFamily?.trim();
      if (!family) return undefined;
      const script = ownValue(scriptScopedFamilies, normalizeFamily(family))?.script;
      if (!script) return undefined;
      // Same priority and tuple selection as resolveUncached: a configured
      // substitute is insufficient when an authored resource outranks it.
      const face = selectedFace(family, normalizedWeight(weight), style ?? 'normal', script);
      return face?.source === 'substitute' && face.script === script ? script : undefined;
    },
    resolve(request: Readonly<FontRequest>): FontResolution {
      const key = JSON.stringify([
        request.requestedFamily ?? null,
        request.genericFamily ?? null,
        request.weight ?? null,
        request.style ?? null,
        request.language ?? null,
        request.cjkFallback ?? null,
        request.script ?? null,
      ]);
      const retained = resolutions.get(key);
      if (retained) return retained;
      const resolution = resolveUncached(request);
      resolutions.set(key, resolution);
      if (resolutions.size > resolutionMemoLimit) {
        resolutions.delete(resolutions.keys().next().value as string);
      }
      return resolution;
    },
  });
}
