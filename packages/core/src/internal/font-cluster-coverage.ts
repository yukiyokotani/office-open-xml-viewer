import { SHAPING_SCRIPT_RANGES, SIMPLE_SYLLABLE_ROLES } from './font-shaping-profile-data.js';
import { canonicalCombiningClass } from './canonical-combining-class.js';
import { CANONICAL_UNICODE_VERSION } from './canonical-font-data.js';
import { canonicalComposition, canonicalDecomposition, canonicalFontClusterText, canonicalOrder,
  canonicalUnicodeAssigned, canonicalUnicodeIgnorable, canonicalUnicodeMark } from './canonical-font-text.js';
import { fontSupportFacts, type FontSupportFacts } from './font-support-registry.js';
export { canonicalFontClusterText, fontSupportFacts };
export type { FontSupportFacts };

/** Fixed INTERNAL library attribution policy, active in window and workers.
 * It adopts the source-backed HarfBuzz default canonical normalizer and modern
 * Hangul path plus Chromium's atomic base/mark missing-glyph fallback rule:
 * hb-ot-shape-normalize.cc, hb-ot-shaper-hangul.cc, harfbuzz_shaper.cc.
 * Browser probes independently distinguish recursive Greek cuts, unavailable
 * alternative Latin cuts, ordinary base/marks, and mixed Hangul tails.
 * This is model-relative evidence, not detected engine capability or a portable
 * claim that Canvas exposes actual resource identity. No UA/geometry oracle.
 * Unsupported script transforms, partitioning and font mechanisms stay unknown;
 * no inference changes the accepted Office line-box/fallback formula. */
export const FONT_SUPPORT_PROFILE = 'canonical-static-v1';
type Proof = Readonly<{ profile: typeof FONT_SUPPORT_PROFILE; unicode: typeof CANONICAL_UNICODE_VERSION; atoms: number }>;
export type ResourceSupport =
  | Readonly<{ kind: 'complete' | 'absent' | 'partial'; proof: Proof }>
  | Readonly<{ kind: 'unknown'; reason: 'sequence' | 'unicode-version' | 'font-transform' | 'coverage' | 'partition' | 'missing-isolation' }>;
type Covers = (cp: number) => boolean | undefined;

function profileProperty(ranges: readonly number[], cp: number): number {
  let lo = 0, hi = ranges.length / 3 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1, at = mid * 3;
    if (cp < ranges[at]) hi = mid - 1;
    else if (cp > ranges[at + 1]) lo = mid + 1;
    else return ranges[at + 2];
  }
  return 0;
}
function erasureSafe(facts: FontSupportFacts, cp: number): boolean {
  if (facts.noErasure === true) return true;
  const ranges = facts.erasureSafeRanges;
  if (!ranges) return false;
  let lo = 0, hi = ranges.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (cp < ranges[mid][0]) hi = mid - 1;
    else if (cp > ranges[mid][1]) lo = mid + 1;
    else return true;
  }
  return false;
}
function hangul(cp: number): boolean { return cp >= 0xac00 && cp <= 0xd7a3; }
function jamo(cp: number): boolean { return cp >= 0x1100 && cp <= 0x11ff || cp >= 0xa960 && cp <= 0xa97f || cp >= 0xd7b0 && cp <= 0xd7ff; }

/** Execute downward direct-decomposition edges from the actual NFC input.
 * b must be mapped before descending into a; a supported parent can be used
 * when deeper descent fails. This follows decompose()/shortest in HarfBuzz's
 * normalizer, not UAX #15 or a search for any equivalent covered spelling.
 * Greek 1F8C → 1F0C+0345 does not require 1F08. Conversely 1EA0+030A cannot
 * reach 00C5+0323 unless A is mapped. Unknown branch facts never choose a path. */
function normalizedAtom(points: readonly number[], covers: Covers): number[] | undefined {
  let unknown = false;
  const mapped = (cp: number) => { const answer = covers(cp); if (answer === undefined) unknown = true; return answer; };
  const descend = (cp: number, shortest: boolean): number[] | null => {
    const direct = canonicalDecomposition(cp);
    if (!direct || (direct[1] && mapped(direct[1]) !== true)) return null;
    const hasA = mapped(direct[0]);
    if (shortest && hasA === true) return direct[1] ? [direct[0], direct[1]] : [direct[0]];
    const deeper = descend(direct[0], shortest);
    if (deeper) { if (direct[1]) deeper.push(direct[1]); return deeper; }
    return hasA === true ? direct[1] ? [direct[0], direct[1]] : [direct[0]] : null;
  };
  const result: number[] = [];
  // A simple supported scalar short-circuits. Bases followed by Unicode marks
  // use the long downward path; mark-category and CCC are distinct properties.
  const shortest = points.length === 1;
  for (const cp of points) {
    if (shortest && mapped(cp) === true) { result.push(cp); continue; }
    const decomposition = descend(cp, shortest);
    if (decomposition) for (const d of decomposition) result.push(d);
    else { mapped(cp); result.push(cp); }
  }
  if (unknown) return undefined;
  const ordered = canonicalOrder(result), composed: number[] = [];
  let starter = -1, previousClass = 0;
  for (const cp of ordered) {
    const cls = canonicalCombiningClass(cp);
    const pair = starter >= 0 && canonicalUnicodeMark(cp)
      && (starter === composed.length - 1 || previousClass < cls) ? canonicalComposition(composed[starter], cp) : undefined;
    if (pair !== undefined && mapped(pair) === true) { composed[starter] = pair; continue; }
    if (cls === 0) starter = composed.length;
    composed.push(cp); previousClass = cls;
  }
  return unknown ? undefined : composed;
}

function normalizedHangulAtom(points: readonly number[], covers: Covers): number[] | undefined {
  const cp = points[0];
  const coverage = covers(cp);
  if (coverage === undefined) return undefined;
  const shaped: number[] = [];
  if (coverage) shaped.push(cp);
  else {
    // Modern Hangul prefers a whole supported syllable; otherwise it uses the
    // complete L/V[/T] path. A partial LV+T canonical cut is not this profile.
    const index = cp - 0xac00;
    shaped.push(0x1100 + Math.floor(index / 588), 0x1161 + Math.floor(index % 588 / 28));
    if (index % 28) shaped.push(0x11a7 + index % 28);
  }
  for (let i = 1; i < points.length; i++) shaped.push(points[i]);
  return shaped;
}

/** Analyze one display unit or final shaping span under canonical-static-v1.
 * Complete definedness requires nonzero preservation, not missing isolation.
 * Absence is stronger: every fallback atom must be rejected and zero must stay
 * isolated through every potentially enabled lookup. A mixed atom result is a
 * barrier, never permission to borrow the next resource's sole metric.
 * Default-profile admission is scalar or one-base+mark units. Simple Indic KA
 * plus one vowel mark and Myanmar KA+E remain admitted; multi-base conjuncts,
 * virama/kinzi sequences, old Hangul and tone marks need profiles we do not implement.
 * Context-safe certificates quantify over every lookup/input, so coalescing or
 * line splitting cannot create missing glyphs in an otherwise complete span.
 * Unsafe/contextual GSUB is unknown for the entire span, including singletons. */
export function analyzeFontResourceSupport(display: string, covers: Covers, facts: FontSupportFacts | undefined): ResourceSupport {
  if (!facts || facts.nonzeroPreserved !== true) return { kind: 'unknown', reason: 'font-transform' };
  if (!display) return { kind: 'unknown', reason: 'partition' };
  const points = [...display].map((ch) => ch.codePointAt(0) as number);
  if (points.some(canonicalUnicodeIgnorable)) return { kind: 'unknown', reason: 'sequence' };
  if (points.some((cp) => !canonicalUnicodeAssigned(cp))) return { kind: 'unknown', reason: 'unicode-version' };
  // Script dispatch and syllable categories are pinned source facts. The
  // default normalizer does not authorize Thai AM, Hebrew composition, USE,
  // Khmer/Sinhala repair, Arabic fallback or arbitrary Indic conjuncts.
  if (points.some((cp) => profileProperty(SHAPING_SCRIPT_RANGES, cp) === 12 || cp === 0x302e || cp === 0x302f)) return { kind: 'unknown', reason: 'partition' };
  const atoms: number[][] = [];
  for (const cp of points) {
    if (canonicalUnicodeMark(cp) && atoms.length) atoms[atoms.length - 1].push(cp);
    else atoms.push([cp]);
  }
  // Tail admission depends on span-wide membership, not on the current atom.
  // Resolve it once so an ordinary prefix plus many tails stays linear.
  const hasModernHangul = atoms.some((atom) => hangul(atom[0]));
  let complete = 0, rejected = 0;
  for (const atom of atoms) {
    if (jamo(atom[0])) {
      // Modern NFC syllables followed by an extra trailing jamo are separate
      // browser fallback atoms even though Intl.Segmenter gives one EGC.
      // Only those modern tails have an admitted scalar atom here.
      if (!(atom[0] >= 0x11a8 && atom[0] <= 0x11c2 && hasModernHangul)) return { kind: 'unknown', reason: 'partition' };
    }
    if (canonicalUnicodeMark(atom[0]) && facts.noErasure !== true) {
      const dot = covers(0x25cc);
      if (dot === undefined || dot && !erasureSafe(facts, 0x25cc)) return { kind: 'unknown', reason: 'font-transform' };
    }
    const script = profileProperty(SHAPING_SCRIPT_RANGES, atom[0]);
    let shaped: number[] | undefined;
    if (script >= 1 && script <= 10) {
      // HarfBuzz 11.0.0 Indic/Myanmar grammar accepts C[VD]? without
      // vowel-repair insertion. Roles intersect Unicode 16/17 ISC with the
      // exact generated C/Ra and M/MPst or Myanmar vowel categories. Indic3
      // tags can select USE instead, so that unresolved route is excluded.
      if (facts.anyIndic3ScriptPresent !== false) return { kind: 'unknown', reason: 'partition' };
      const role = profileProperty(SIMPLE_SYLLABLE_ROLES, atom[0]);
      const simple = role === 1 && (atom.length === 1 || atom.length === 2
        && profileProperty(SHAPING_SCRIPT_RANGES, atom[1]) === script
        && profileProperty(SIMPLE_SYLLABLE_ROLES, atom[1]) === 2);
      const singletonMark = atom.length === 1 && canonicalUnicodeMark(atom[0]) && !canonicalDecomposition(atom[0]);
      if (!simple && !singletonMark) return { kind: 'unknown', reason: 'partition' };
      if (atom.every((cp) => covers(cp) === true) && facts.noErasure === true) shaped = atom;
      else {
        // Generic and specialized decomposition hooks need no inference for
        // a nondecomposable simple unit; reorder only preserves glyphs. An
        // unsupported split-vowel/custom-decomposition route remains unknown.
        if (atom.some((cp) => canonicalDecomposition(cp))) return { kind: 'unknown', reason: 'partition' };
        shaped = atom;
      }
    } else shaped = hangul(atom[0]) ? normalizedHangulAtom(atom, covers) : normalizedAtom(atom, covers);
    if (!shaped) return { kind: 'unknown', reason: 'coverage' };
    let missing = false;
    for (const cp of shaped) {
      const presence = covers(cp);
      if (presence === undefined) return { kind: 'unknown', reason: 'coverage' };
      if (presence && !erasureSafe(facts, cp)) {
        return { kind: 'unknown', reason: 'font-transform' };
      }
      if (!presence) missing = true;
    }
    if (missing) rejected++; else complete++;
  }
  const proof: Proof = { profile: FONT_SUPPORT_PROFILE, unicode: CANONICAL_UNICODE_VERSION, atoms: atoms.length };
  if (!rejected) return { kind: 'complete', proof };
  if (facts.missingIsolated !== true) return { kind: 'unknown', reason: 'missing-isolation' };
  return { kind: complete ? 'partial' : 'absent', proof };
}

/** Normalization-profile compatibility helper. This tests a supplied scalar
 * oracle under identity transformation; production must pass parsed facts to
 * analyzeFontResourceSupport. It cannot authorize a font resource by itself. */
export function fontResourceCoversCluster(cluster: string, covers: Covers): boolean | undefined {
  const support = analyzeFontResourceSupport(canonicalFontClusterText(cluster), covers,
    { schema: 'ot-definedness-1', glyphCount: 65535, nonzeroPreserved: true, missingIsolated: true, noErasure: true, anyIndic3ScriptPresent: false });
  return support.kind === 'complete' ? true : support.kind === 'absent' ? false : undefined;
}
