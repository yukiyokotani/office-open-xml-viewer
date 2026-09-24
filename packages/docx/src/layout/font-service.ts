import { cjkLangFromLanguage, type CjkLang } from '@silurus/ooxml-core';
import type { LayoutDiagnostic } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import { createCanvasFontRoute, type CanvasFontRoute } from '@silurus/ooxml-core';

export type FontResolutionSource = 'embedded' | 'local' | 'google' | 'substitute' | 'native' | 'generic';
export type FontStyle = 'normal' | 'italic';

export interface FontRequest {
  readonly cjkFallback?: CjkLang;
  readonly language?: string;
  readonly requestedFamily?: string | null;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly weight?: number;
  readonly style?: FontStyle;
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
}

export interface FontResolver {
  readonly fingerprint: string;
  resolve(request: Readonly<FontRequest>): FontResolution;
}

export interface FontInventoryFace {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly source: Exclude<FontResolutionSource, 'generic' | 'native'>;
  readonly resourceIdentity?: string;
  readonly weight?: number;
  readonly style?: FontStyle;
}

export interface FontResolverOptions {
  /** Immutable regional routes; their concrete outputs participate in cache identity. */
  readonly regionalFamilyLists?: Partial<Record<CjkLang, Readonly<Record<string, string>>>>;
  /** Stable DOCX fallback routes derived from document metadata and rendered faces. */
  readonly nativeFamilyLists?: Readonly<Record<string, string>>;
}

function normalizeFamily(value: string): string {
  return value.trim().toLocaleLowerCase('en-US');
}

function normalizedWeight(value: number | undefined): number {
  if (value == null || !Number.isFinite(value)) return 400;
  return Math.min(900, Math.max(100, Math.round(value / 100) * 100));
}

function freezeResolution(value: FontResolution): FontResolution {
  return Object.freeze({ ...value, diagnostics: Object.freeze([...value.diagnostics]) });
}

function quoteCssFamily(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`;
}

function cssFamilyList(family: string, generic: FontResolution['genericFamily']): string {
  return `${quoteCssFamily(family)}, ${generic}`;
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
    google: 2,
    substitute: 3,
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
    return (region ? regionalFamilyLists[region]?.[normalizeFamily(family)] : undefined)
      ?? nativeFamilyLists[normalizeFamily(family)];
  };
  const fingerprint = stableFingerprint('fonts', { faces, nativeFamilyLists, regionalFamilyLists });

  return Object.freeze({
    fingerprint,
    resolve(request: Readonly<FontRequest>): FontResolution {
      const requestedFamily = request.requestedFamily?.trim() || request.genericFamily || 'sans-serif';
      const weight = normalizedWeight(request.weight);
      const style = request.style ?? 'normal';
      const candidates = byFamily.get(normalizeFamily(requestedFamily)) ?? [];
      const face = candidates.find((candidate) => candidate.weight === weight && candidate.style === style);
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
          ? `${quoteCssFamily(face.resolvedFamily)}, ${fallbackList}`
          : cssFamilyList(face.resolvedFamily, request.genericFamily ?? 'sans-serif');
        return freezeResolution({
          requestedFamily,
          resolvedFamily: face.resolvedFamily,
          route: createCanvasFontRoute(familyList, 'registered'),
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
        return freezeResolution({
          requestedFamily,
          resolvedFamily: authored,
          route: createCanvasFontRoute(familyList, 'native'),
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
    },
  });
}
