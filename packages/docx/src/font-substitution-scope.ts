import type { FontSubstituteScript } from '@silurus/ooxml-core';
import {
  fontFamilyKey,
  type FontSourceDescriptor,
  type FontSubstitution,
} from '@silurus/ooxml-core/internal/font-sources';

// Substitution scope of the configured font sources (LoadOptions.fonts):
// pure data over the descriptors, free of parser-model facts so layout font
// routing may read it.

/** Every substitution the configured sources declare, in option order. Their
 * script scopes are policy (core `substitute-script.ts`): they delimit where a
 * scoped substitute may paint whether or not its face loads. */
export function declaredFontSubstitutions(
  sources: readonly FontSourceDescriptor[] | undefined,
): readonly FontSubstitution[] {
  return Object.freeze((sources ?? []).flatMap((source) => source.substitutes ?? []));
}

/** Authored families (normalised) every declared substitution of which is
 * scoped to one script, mapped to that script. A family with any unscoped
 * substitution, or with substitutions of different scripts, is not scoped. */
export function scriptScopedSubstitutedFamilies(
  substitutes: readonly FontSubstitution[],
): ReadonlyMap<string, FontSubstituteScript> {
  const scripts = new Map<string, FontSubstituteScript | null>();
  for (const substitution of substitutes) {
    const key = fontFamilyKey(substitution.for);
    const script = substitution.script ?? null;
    const prior = scripts.get(key);
    scripts.set(key, prior === undefined || prior === script ? script : null);
  }
  const scoped = new Map<string, FontSubstituteScript>();
  for (const [key, script] of scripts) if (script !== null) scoped.set(key, script);
  return scoped;
}
