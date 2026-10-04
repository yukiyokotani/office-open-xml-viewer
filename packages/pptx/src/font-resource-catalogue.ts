import { referenceFontMetricOrdinal, REFERENCE_METRIC_GENERATION, type ReferenceFontMetricProfile } from '@silurus/ooxml-core/internal/reference-font-metrics';
import { decodeFontRanges, decodeFontBitmap } from '@silurus/ooxml-core/internal/font-data-codec';
import type { FontSupportFacts } from '@silurus/ooxml-core/internal/font-support-registry';
import data from './font-resource-catalogue-data.json';
/** PPTX alone consumes this exact per-cut catalogue. The shared metric owner
 * stores a private physical-row ordinal; equal tables/names never join cuts.
 * Companion generation IDs prevent mismatched data from lending certainty.
 * Published-open/foreign profiles without a row retain unknown support.
 * Decode only requested repertoires/certificates, bounded by immutable table
 * counts; no query-key cache, font bytes, glyph graphs or worker payload exists. */
const coherent = data.generation === REFERENCE_METRIC_GENERATION && data.schemaVersion === 1
  && data.supportSchema === 'ot-definedness-1' && data.supportProfile === 'canonical-static-v1' && data.unicode === '17.0.0';
// Generated physical records are 13 bytes: five nullable uint16 index+1
// fields, one nullable uint16 glyph count, and basic-CJK three-valued coverage.
// Decode once on demand, then materialize only queried physical rows. Both
// retained caches are bounded by the generated catalogue, never query strings.
let packedRows: string | undefined;
const rows = new Map<number, readonly (number | null)[]>();
function row(profile: ReferenceFontMetricProfile): readonly (number | null)[] | undefined {
  const ordinal = referenceFontMetricOrdinal(profile);
  if (!coherent || ordinal === undefined || ordinal >= data.rowCount) return undefined;
  const cached = rows.get(ordinal); if (cached) return cached;
  if (packedRows === undefined) {
    try { packedRows = atob(data.rowsEncoded); } catch { return undefined; }
    if (packedRows.length !== data.rowCount * 13) { packedRows = undefined; return undefined; }
  }
  const offset = ordinal * 13, values: (number | null)[] = [];
  for (let i = 0; i < 6; i++) {
    const value = packedRows.charCodeAt(offset + i * 2) * 256 + packedRows.charCodeAt(offset + i * 2 + 1);
    values.push(value === 0 ? null : value - (i < 5 ? 1 : 0));
  }
  const basic = packedRows.charCodeAt(offset + 12); if (basic > 2) return undefined;
  values.push(basic);
  const result = Object.freeze(values); rows.set(ordinal, result); return result;
}
const symbolRanges = new Map<number, readonly number[]>();
function ranges(index: number | null | undefined): readonly number[] | undefined {
  if (index == null || !Number.isInteger(index) || index < 0 || index >= data.symbolCoverages.length) return undefined;
  const cached = symbolRanges.get(index); if (cached) return cached;
  const decoded = decodeFontRanges(data.symbolCoverages[index]);
  if (decoded) symbolRanges.set(index, decoded);
  return decoded;
}
function includes(ranges: readonly number[], cp: number): boolean {
  let lo = 0, hi = ranges.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (cp < ranges[mid * 2]) hi = mid - 1;
    else if (cp > ranges[mid * 2 + 1]) lo = mid + 1;
    else return true;
  }
  return false;
}
export function isReferenceSymbolCodePoint(cp: number): boolean {
  return Number.isInteger(cp) && data.symbolCoverageRanges.some(([lo, hi]) => cp >= lo && cp <= hi);
}
export function referenceFontCoversSymbol(profile: ReferenceFontMetricProfile, cp: number): boolean | undefined {
  if (!isReferenceSymbolCodePoint(cp)) return undefined;
  const record = row(profile), known = ranges(record?.[0]);
  if (!known) return undefined;
  if (includes(known, cp)) return true;
  const possible = ranges(record?.[1]);
  return !possible || includes(possible, cp) ? undefined : false;
}
const cjkBitmaps = new Map<number, Uint8Array>();
const CJK_BYTES = Math.ceil(data.cjkCoverageRanges.reduce((sum, [lo, hi]) => sum + hi - lo + 1, 0) / 8);
function cjkBitmap(index: number | null | undefined): Uint8Array | undefined {
  if (index == null || !Number.isInteger(index) || index < 0 || index >= data.cjkCoverages.length) return undefined;
  const cached = cjkBitmaps.get(index); if (cached) return cached;
  const bits = decodeFontBitmap(data.cjkCoverages[index], CJK_BYTES);
  if (!bits) return undefined;
  cjkBitmaps.set(index, bits); return bits;
}
export function referenceFontCoversCjk(profile: ReferenceFontMetricProfile, cp: number): boolean | undefined {
  if (!Number.isInteger(cp)) return undefined;
  const record = row(profile), known = cjkBitmap(record?.[2]);
  if (!known) return undefined;
  let offset = 0;
  for (const [lo, hi] of data.cjkCoverageRanges) {
    if (cp >= lo && cp <= hi) {
      const index = offset + cp - lo, mask = 1 << (index & 7);
      if (known[index >>> 3] & mask) return true;
      const possible = cjkBitmap(record?.[3]);
      return !possible || possible[index >>> 3] & mask ? undefined : false;
    }
    offset += hi - lo + 1;
  }
  return undefined;
}
const certificates = new Map<number, FontSupportFacts>();
const ternary = (n: number | null | undefined): boolean | undefined => n === 0 || n == null ? undefined : n === 2;
/** Materialization preserves optional/unknown fields and erasure projection.
 * Every enum is generic representation, never a font-selection rule. */
export function referenceFontSupportFacts(profile: ReferenceFontMetricProfile): FontSupportFacts | undefined {
  const ordinal = referenceFontMetricOrdinal(profile), record = row(profile);
  if (ordinal === undefined || !record || record[4] === null) return undefined;
  const cached = certificates.get(ordinal); if (cached) return cached;
  const t = data.templates[record[4] as number]; if (!t) return undefined;
  const decoded = t[7] === null ? undefined : decodeFontRanges(data.erasureRanges[t[7] as number]);
  if (t[7] !== null && !decoded) return undefined;
  const erasureSafeRanges = decoded && Object.freeze(Array.from({length: decoded.length / 2}, (_, i) => Object.freeze([decoded[i * 2], decoded[i * 2 + 1]] as const)));
  const result: FontSupportFacts = Object.freeze({ schema: 'ot-definedness-1',
    glyphCount: record[5] ?? undefined, nonzeroPreserved: ternary(t[0]), missingIsolated: ternary(t[1]),
    ...(t[2] ? {noErasure: ternary(t[2])} : {}), ...(t[3] ? {anyIndic3ScriptPresent: ternary(t[3])} : {}),
    ...(t[4] !== null ? {gsubLookupCount: t[4]} : {}),
    ...(t[5] ? {gsubDisposition: ['identity', 'active-open-type', 'profile-inactive-major'][(t[5] as number) - 1] as FontSupportFacts['gsubDisposition']} : {}),
    ...(t[6] ? {reason: ['glyph-domain', 'unsupported', 'malformed', 'budget', 'cycle'][(t[6] as number) - 1] as FontSupportFacts['reason']} : {}),
    ...(t[8] ? {profile: 'canonical-static-v1' as const} : {}), ...(erasureSafeRanges ? {erasureSafeRanges} : {}),
  });
  certificates.set(ordinal, result); return result;
}
