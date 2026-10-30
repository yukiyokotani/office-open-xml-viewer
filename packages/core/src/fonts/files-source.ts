/**
 * Family matching over a known declaration list, and the `font-files` source.
 *
 * A `font-files` source ({@link fontFiles}) declares its faces up front, so it
 * resolves synchronously and fetches nothing at resolution: URL faces are
 * fetched later by `FontFace.load()` in the painting realm, only when their
 * `unicode-range` intersects the demanded scalars, and `data` faces are
 * registered from the owned bytes. The same index matching answers a
 * `font-css` stylesheet once it is parsed (`css-source.ts`).
 *
 * Matching is the single rule of `substitutions.ts`: declarations of the
 * demanded family itself win; otherwise the first of the source's declared
 * substitutes (declaration order) that the source serves claims the family.
 */
import type { FamilyResolution, SourceCache } from './source-cache.js';
import type { FontFaceDeclaration, FontFilesSource } from './source.js';
import { fontFamilyKey, substitutionsFor, type FontSubstituteScripts, type FontSubstitution } from './substitutions.js';

/** Declarations grouped by lower-cased family, each group in declaration
 * order (overlapping `unicode-range` precedence is order-dependent). */
export function indexFontDeclarations(
  declarations: readonly FontFaceDeclaration[],
): ReadonlyMap<string, readonly FontFaceDeclaration[]> {
  const index = new Map<string, FontFaceDeclaration[]>();
  for (const declaration of declarations) {
    const key = fontFamilyKey(declaration.family);
    const group = index.get(key);
    if (group) group.push(declaration);
    else index.set(key, [declaration]);
  }
  return index;
}

/** Claim `family` from an index, directly or through a declared substitute;
 * `absent` when the index serves neither. */
export function claimFromIndex(
  index: ReadonlyMap<string, readonly FontFaceDeclaration[]>,
  family: string,
  substitutes: readonly FontSubstitution[] | undefined,
  scripts: FontSubstituteScripts = 'all',
): FamilyResolution {
  const direct = index.get(fontFamilyKey(family));
  if (direct?.length) return { kind: 'claimed', served: direct[0].family, declarations: direct };
  for (const substitution of substitutionsFor(substitutes, family, scripts)) {
    const served = index.get(fontFamilyKey(substitution.use));
    if (served?.length) {
      return { kind: 'claimed', served: served[0].family, substitution, declarations: served };
    }
  }
  return { kind: 'absent' };
}

/** Resolve demanded families against a `font-files` source. */
export function resolveFontFiles(
  source: FontFilesSource,
  cache: SourceCache,
  families: readonly string[],
  scriptsOf: (family: string) => FontSubstituteScripts = () => 'all',
): ReadonlyMap<string, FamilyResolution> {
  cache.index ??= indexFontDeclarations(source.faces);
  const index = cache.index;
  return new Map(families.map((family) => [fontFamilyKey(family),
    claimFromIndex(index, family, source.substitutes, scriptsOf(family))]));
}
