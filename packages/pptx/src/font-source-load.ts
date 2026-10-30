/**
 * PPTX font-source loading (`LoadOptions.fonts`): the demand one presentation
 * (or one progressive slide delta) asks the configured sources for. Library
 * resource policy, not OOXML behaviour: DrawingML names faces (ECMA-376
 * §21.1.2.3) and leaves their provision to the host.
 *
 * The main thread imports this module lazily, only when sources are
 * configured, so the default offline path never loads the session or the CSS
 * adapters. The main thread and the render worker build the same demand from
 * the same preflight facts with the same descriptors, so both realms demand,
 * register and load identical faces.
 */
import type { FontPreloadDemand } from '@silurus/ooxml-core';
import type { FontSessionDemand } from '@silurus/ooxml-core/internal/font-sources';
import { excludeEmbeddedFontFamilies } from './embedded-fonts.js';
import type { PptxFontDemandFacts } from './font-requests.js';

export { FontSourceSession } from '@silurus/ooxml-core/internal/font-sources';

const CSS_GENERIC_FAMILY_NAMES: ReadonlySet<string> = new Set([
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui',
]);

/** A preflight name that names a face a source could serve: not a DrawingML
 * theme token (`+mj-lt`, `+mn-ea`, …, §20.1.4.1.17; its theme font is
 * demanded by name) and not a CSS generic family. */
function demandableFamily(name: string | null): name is string {
  const family = name?.trim();
  return !!family && !family.startsWith('+mj-') && !family.startsWith('+mn-')
    && !CSS_GENERIC_FAMILY_NAMES.has(family.toLowerCase());
}

/**
 * One presentation's font demand (or one progressive delta's: families
 * already handled by the session are not resolved again, and `scalars` is the
 * newly demanded scalar delta).
 *
 * - `families` / `fallbackFamilies`: the preflight's authored families and
 *   library script fallbacks (`fontPreloadNames`, split at
 *   `facts.familyCount`), minus families an embedded font part
 *   (ECMA-376 Part 1 §19.2.1.9 `p:embeddedFont`) already serves;
 * - `scalars`: the collected Unicode scalars (`'all'` when the collector
 *   declined or the preflight omitted them, the conservative answer);
 * - `substituteScripts`: per family, the scoped-substitute scripts its painted
 *   text proves, so a script-scoped substitute is never requested for text it
 *   would not paint.
 */
export function pptxFontDemand(
  names: readonly (string | null)[],
  facts: PptxFontDemandFacts | undefined,
  scalars: FontPreloadDemand | undefined,
  embeddedAliases: ReadonlyMap<string, string>,
): FontSessionDemand {
  const familyCount = facts?.familyCount ?? names.length;
  const families = excludeEmbeddedFontFamilies(names.slice(0, familyCount), embeddedAliases).filter(demandableFamily);
  const fallbackFamilies = excludeEmbeddedFontFamilies(names.slice(familyCount), embeddedAliases).filter(demandableFamily);
  return Object.freeze({
    families: Object.freeze(families),
    fallbackFamilies: Object.freeze(fallbackFamilies),
    scalars: scalars ?? 'all',
    ...(facts ? { substituteScripts: facts.substituteScripts } : {}),
  });
}

/** Installed authored families are probed before any source face of a delta
 * registers, and never demanded (core `installed-families.ts`). */
export { InstalledFamilyExclusion as PptxInstalledFamilies } from '@silurus/ooxml-core/internal/font-sources';
