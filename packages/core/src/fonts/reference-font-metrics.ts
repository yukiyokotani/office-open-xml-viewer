import referenceData from './reference-font-metrics-data.json';
import { OPEN_FONT_REFERENCE_PROFILES } from './reference-font-metrics-open.js';
import { normalizeReferenceFontFamilyName } from './reference-font-name.js';

export type ReferenceFontSource = 'office-mac' | 'macos-system' | 'macos-supplemental' | 'published-open-font';
export type ReferenceFontStyle = 'normal' | 'italic';
export interface ReferenceFontMetricProfile {
  readonly source: ReferenceFontSource;
  readonly family: string;
  readonly aliases: readonly string[];
  readonly weight: number;
  readonly style: ReferenceFontStyle;
  readonly unitsPerEm: number;
  /** Signed OS/2 xAvgCharWidth design units, when declared. */
  readonly xAvgCharWidth?: number | null;
  readonly hhea: readonly [ascender: number, descender: number, lineGap: number];
  /** OS/2 [usWinAscent, usWinDescent] design units. Null means the face has no
   * OS/2 table; undefined means this source did not record the field. */
  readonly win?: readonly [ascent: number, descent: number] | null;
  /** OS/2 [sTypoAscender, sTypoDescender, sTypoLineGap], present only when the
   * face sets fsSelection USE_TYPO_METRICS (OS/2 v4+). */
  readonly typoMetrics?: readonly [ascender: number, descender: number, lineGap: number];
  /** GSUB vert single-substitution inputs within the generated stacked-sideways
   * domain; vrt2 is excluded. Omitted when no coverage is recorded. */
  readonly stackedVerticalGlyphs?: string;
  /** GSUB vrt2 single-substitution inputs within the same domain that vert
   * does not cover. Omitted when there are none. */
  readonly stackedVrt2OnlyGlyphs?: string;
  /** Derived OS/2 code-page class. Null means this source did not provide the
   * code-page field needed to classify Word's auto-line allocation. */
  readonly farEastCodePage: boolean | null;
  /** OS/2 PANOSE [bFamilyType, bSerifStyle]. Null means the face has no OS/2
   * table; undefined means this source did not record the field. */
  readonly panose?: readonly [familyType: number, serifStyle: number] | null;
  /** post [underlinePosition, underlineThickness] and OS/2
   * [yStrikeoutPosition, yStrikeoutSize] design units (positions are the
   * rule's top edge, the underline's negative below the baseline). Null when
   * the face lacks post or OS/2; undefined when the source did not record it. */
  readonly decoration?: readonly [
    underlinePosition: number,
    underlineThickness: number,
    strikeoutPosition: number,
    strikeoutSize: number,
  ] | null;
}

export interface FindReferenceFontMetricsOptions {
  readonly source?: ReferenceFontSource;
  readonly weight?: number;
  readonly style?: ReferenceFontStyle;
}

function freezeProfile(profile: ReferenceFontMetricProfile): ReferenceFontMetricProfile {
  Object.freeze(profile.aliases);
  Object.freeze(profile.hhea);
  if (profile.win) Object.freeze(profile.win);
  if (profile.typoMetrics) Object.freeze(profile.typoMetrics);
  if (profile.panose) Object.freeze(profile.panose);
  if (profile.decoration) Object.freeze(profile.decoration);
  return Object.freeze(profile);
}

// The generator includes every supported static face from the Office and macOS
// catalogs; independently verified, pinned open-font profiles follow them.
let profiles: readonly ReferenceFontMetricProfile[] | undefined;

function getProfiles(): readonly ReferenceFontMetricProfile[] {
  // The JSON payload is statically imported and parsed with the module. The
  // alias index is deferred until the first lookup, and each profile is frozen
  // when a lookup first buckets it (getOptionBuckets), so every profile a
  // caller can observe is frozen without freezing the whole catalogue.
  return profiles ??= Object.freeze(
    [...referenceData.profiles as unknown as ReferenceFontMetricProfile[],
      ...OPEN_FONT_REFERENCE_PROFILES],
  );
}

type OptionBuckets = ReadonlyMap<string, readonly ReferenceFontMetricProfile[]>;
let aliasProfiles: ReadonlyMap<string, readonly ReferenceFontMetricProfile[]> | undefined;
// Option buckets are built per alias on its first lookup. Documents name a few
// families, so first paint no longer pays for the whole catalogue's buckets.
const aliasBuckets = new Map<string, OptionBuckets>();
const EMPTY_RESULTS = Object.freeze([]) as readonly ReferenceFontMetricProfile[];

const normalizeFamilyName = normalizeReferenceFontFamilyName;

function optionKey(options: FindReferenceFontMetricsOptions): string {
  return `${options.source ?? '*'}|${options.weight ?? '*'}|${options.style ?? '*'}`;
}

function getAliasProfiles(): ReadonlyMap<string, readonly ReferenceFontMetricProfile[]> {
  if (aliasProfiles !== undefined) return aliasProfiles;
  const mutable = new Map<string, ReferenceFontMetricProfile[]>();
  for (const profile of getProfiles()) {
    for (const alias of new Set(profile.aliases.map(normalizeFamilyName))) {
      const matches = mutable.get(alias);
      if (matches === undefined) mutable.set(alias, [profile]);
      else matches.push(profile);
    }
  }
  aliasProfiles = mutable;
  return aliasProfiles;
}

function getOptionBuckets(alias: string): OptionBuckets | undefined {
  const cached = aliasBuckets.get(alias);
  if (cached !== undefined) return cached;
  const matches = getAliasProfiles().get(alias);
  if (matches === undefined) return undefined;
  const buckets = new Map<string, ReferenceFontMetricProfile[]>();
  for (const profile of matches) {
    freezeProfile(profile);
    const optionKeys = new Set<string>();
    for (const source of [undefined, profile.source] as const) {
      for (const weight of [undefined, profile.weight] as const) {
        for (const style of [undefined, profile.style] as const) {
          optionKeys.add(optionKey({ source, weight, style }));
        }
      }
    }
    for (const key of optionKeys) {
      const bucket = buckets.get(key);
      if (bucket === undefined) buckets.set(key, [profile]);
      else bucket.push(profile);
    }
  }
  for (const bucket of buckets.values()) Object.freeze(bucket);
  // Bounded solely by aliases in the immutable generated catalogue: unknown
  // caller-provided names return above and are never memoized.
  aliasBuckets.set(alias, buckets);
  return buckets;
}

/**
 * Finds metadata-only reference profiles without claiming which font Canvas or
 * the host selected. Multiple results intentionally preserve same-name metric
 * conflicts across and within source catalogs.
 */
export function findReferenceFontMetrics(
  familyOrAlias: string,
  options: FindReferenceFontMetricsOptions = {},
): readonly ReferenceFontMetricProfile[] {
  const query = normalizeFamilyName(familyOrAlias);
  if (!query) return EMPTY_RESULTS;
  return getOptionBuckets(query)?.get(optionKey(options)) ?? EMPTY_RESULTS;
}
