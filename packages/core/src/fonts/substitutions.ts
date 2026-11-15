/**
 * Vendor-neutral font substitution data and matching.
 *
 * A substitution says that a family a source serves may stand in for an
 * authored family the host lacks. ECMA-376 leaves font substitution
 * implementation-defined (§17.8.2 for WordprocessingML; DrawingML and
 * SpreadsheetML name faces without a substitution rule), so this is library
 * policy, not an Office layout rule. It contains no URLs and no provider
 * naming: any font source may adopt {@link OFFICE_FONT_SUBSTITUTIONS}, and the
 * renderers consume only the substitutions whose served face actually loaded.
 *
 * Keep the list to published advance-width alternatives or clearly identified
 * visual substitutes. Adding a pair changes how documents using the authored
 * name measure (from "whatever the host falls back to" to the substitute), so
 * do not add speculative names or claim exact Office layout.
 */
import type { FontSubstituteScript } from './substitute-script.js';

export interface FontSubstitution {
  /** Authored family this substitution may stand in for (case-insensitive). */
  readonly for: string;
  /** Family the source actually serves. */
  readonly use: string;
  /** Visual substitute limited to one script (substitute-script.ts). */
  readonly script?: FontSubstituteScript;
}

function freezeAll(entries: readonly FontSubstitution[]): readonly FontSubstitution[] {
  return Object.freeze(entries.map((entry) => Object.freeze({ ...entry })));
}

/**
 * Office faces with an openly licensed substitute.
 *
 * - Calibri → Carlito and Cambria → Caladea are the published
 *   advance-width-compatible pairings. Their vertical metrics need not match.
 *   Carlito has no Calibri Light (weight 300) counterpart and Caladea has
 *   neither Cambria's vertical metrics nor Cambria Math's MATH table, so those
 *   distinct faces are deliberately absent.
 * - Libre Franklin is the open Franklin-family substitute. Office templates
 *   name Book (regular) and Medium in the family rather than with bold, so both
 *   names are listed.
 * - Arabic faces hosts rarely ship map to Noto Naskh Arabic (traditional,
 *   serif-like) or Noto Sans Arabic (modern geometric). They are visual
 *   substitutes for Arabic script only: the Noto faces' Latin glyphs and tall
 *   line boxes do not resemble the authored faces' own Latin, so they are
 *   scoped to Arabic characters and other text from the same family falls back
 *   as if no substitute existed.
 */
export const OFFICE_FONT_SUBSTITUTIONS: readonly FontSubstitution[] = freezeAll([
  { for: 'Calibri', use: 'Carlito' },
  { for: 'Cambria', use: 'Caladea' },
  { for: 'Franklin Gothic Book', use: 'Libre Franklin' },
  { for: 'Franklin Gothic Medium', use: 'Libre Franklin' },
  { for: 'Sakkal Majalla', use: 'Noto Naskh Arabic', script: 'arabic' },
  { for: 'Traditional Arabic', use: 'Noto Naskh Arabic', script: 'arabic' },
  { for: 'Simplified Arabic', use: 'Noto Naskh Arabic', script: 'arabic' },
  { for: 'Arabic Typesetting', use: 'Noto Naskh Arabic', script: 'arabic' },
  { for: 'Univers Next Arabic', use: 'Noto Sans Arabic', script: 'arabic' },
]);

/** Case-insensitive comparison key of a CSS family name: trimmed, one pair of
 * surrounding quotes removed (as `FontFace.family` may report it). */
export function fontFamilyKey(family: string): string {
  return family.trim().replace(/^(['"])(.*)\1$/s, '$2').trim().toLocaleLowerCase('en-US');
}

/** The substitute scripts one demanded family's text proves
 * (`FontDemand.substituteScripts`), or `'all'` when demand does not restrict
 * scoped substitutions. */
export type FontSubstituteScripts = 'all' | readonly FontSubstituteScript[];

/** Substitutions a source declares for one authored family, in declaration
 * order (the first is preferred). A script-scoped substitution is admissible
 * only when `scripts` proves its script: it limits when the substitute is
 * requested, never whether the authored family itself is. */
export function substitutionsFor(
  substitutions: readonly FontSubstitution[] | undefined,
  authoredFamily: string,
  scripts: FontSubstituteScripts = 'all',
): readonly FontSubstitution[] {
  if (!substitutions || substitutions.length === 0) return [];
  const key = fontFamilyKey(authoredFamily);
  return substitutions.filter((entry) => fontFamilyKey(entry.for) === key
    && (entry.script === undefined || scripts === 'all' || scripts.includes(entry.script)));
}

/** Resolution-cache key of one family under one substitution scope: a family
 * resolved without its scoped substitutes is a different answer from one
 * resolved with them, so they never share a cached resolution. */
export function scopedFamilyKey(family: string, scripts: FontSubstituteScripts): string {
  const key = fontFamilyKey(family);
  return scripts === 'all' ? key : `${key}\u0000${[...scripts].sort().join(',')}`;
}

/** How a declared face family satisfies one demanded family. */
export type FontDeclarationMatch =
  | { readonly kind: 'family' }
  | { readonly kind: 'substitute'; readonly substitution: FontSubstitution };

/**
 * The single matching rule for every source kind: a declaration whose CSS
 * family equals the demanded family satisfies it directly; otherwise it
 * satisfies it as a substitute when it equals the `use` of one of the source's
 * substitutions whose `for` equals the demanded family. A direct match is
 * preferred over a substitute (an authored face beats a stand-in).
 */
export function matchFontDeclaration(
  declarationFamily: string,
  demandedFamily: string,
  substitutions: readonly FontSubstitution[] | undefined,
): FontDeclarationMatch | null {
  const declared = fontFamilyKey(declarationFamily);
  if (declared === fontFamilyKey(demandedFamily)) return { kind: 'family' };
  const substitution = substitutionsFor(substitutions, demandedFamily)
    .find((entry) => fontFamilyKey(entry.use) === declared);
  return substitution ? { kind: 'substitute', substitution } : null;
}

/** The face facts {@link loadedRegularAliases} reads (a `FontFace` fits). */
export interface LoadedFaceFacts {
  readonly family: string;
  readonly weight: string;
  readonly style: string;
  readonly status: FontFaceLoadStatus;
}

/**
 * Some Office files record a regular face's full name (for example "Lato
 * Regular") while CSS registers the family as "Lato" with weight 400. Offer
 * that family for the full name only when a weight-400 normal face of it
 * actually loaded from a font source, so an authored or installed full-name
 * face stays first and a failed or absent source never changes the fallback
 * chain. Callers pass the faces their font sources loaded, not the whole
 * `FontFaceSet` (an application's own `@font-face` is not a source face).
 *
 * Returns lower-cased `"<family> regular"` → the loaded family's spelling.
 */
export function loadedRegularAliases(
  loadedFaces: Iterable<LoadedFaceFacts>,
): ReadonlyMap<string, string> {
  const aliases = new Map<string, string>();
  for (const face of loadedFaces) {
    if (face.status !== 'loaded' || face.style !== 'normal') continue;
    if (face.weight !== '400' && face.weight !== 'normal') continue;
    const family = face.family.trim().replace(/^(['"])(.*)\1$/s, '$2').trim();
    if (!family) continue;
    const key = `${family.toLocaleLowerCase('en-US')} regular`;
    if (!aliases.has(key)) aliases.set(key, family);
  }
  return aliases;
}
