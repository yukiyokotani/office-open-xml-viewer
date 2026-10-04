import type { ReferenceFontMetricProfile } from './reference-font-metrics.js';
/** One alias normalization authority for reference metrics and generated PPTX
 * route descriptors. Caller queries are not retained in any cache. */
export function normalizeReferenceFamily(value: string): string {
  return value.normalize('NFKC').trim().replace(/\s+/gu, ' ').toLocaleLowerCase('en-US');
}
/** Exhaustive family-level route projection, separate from per-cut ownership.
 * Preserve profilesOf's Office-first source preference, OR classification and
 * three-valued basic-CJK absence. Generation consumes every alias and row;
 * no selected font list or independent renderer classification policy exists. */
export function deriveReferenceFontRoutes(profiles: readonly (ReferenceFontMetricProfile & { cjkUnifiedIdeographs?: boolean | null })[]): readonly (readonly [string, number])[] {
  const buckets = new Map<string, typeof profiles[number][]>();
  for (const profile of profiles) for (const alias of new Set(profile.aliases.map(normalizeReferenceFamily))) {
    const rows = buckets.get(alias) ?? []; rows.push(profile); buckets.set(alias, rows);
  }
  return [...buckets].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0).map(([alias, rows]) => {
    const office = rows.filter(row => row.source === 'office-mac');
    const chosen = office.length ? office : rows;
    const eastAsian = chosen.some(row => row.farEastCodePage === true);
    // Existing PPTX compatibility classification: Latin PANOSE serif styles
    // 1–10 select serif tiers; style 0 and non-Latin kinds remain unknown/sans.
    // This family aggregation selects a chain, never a concrete glyph owner.
    const serif = chosen.some(row => row.panose?.[0] === 2 && row.panose[1] >= 1 && row.panose[1] <= 10);
    const cjk = chosen.some(row => row.cjkUnifiedIdeographs === true) ? 2
      : chosen.every(row => row.cjkUnifiedIdeographs === false) ? 1 : 0;
    return [alias, (eastAsian ? 1 : 0) | (serif ? 2 : 0) | cjk << 2] as const;
  });
}
