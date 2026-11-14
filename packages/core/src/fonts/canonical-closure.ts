/**
 * The scalars a canonically equivalent rendering of one text cluster can
 * paint. Text shaping may decompose a precomposed character the face lacks,
 * or compose a base and marks the face holds precomposed (HarfBuzz
 * normalization; UAX #15 canonical equivalence), so a text that spells `が`
 * (U+304C) can paint `か` U+304B and the voiced mark U+3099, and one that
 * spells `e` U+0301 can paint `é` U+00E9. The closure of a cluster (a
 * starter and the joiners after it) under canonical decomposition and
 * pairwise canonical composition (with the composition-excluded composites
 * a shaper may still compose) is a superset of every such rendering: a
 * superset only loads more `unicode-range` subsets.
 *
 * The decompositions and compositions are the host's `normalize()`, used
 * only for scalars the host knows are assigned: canonical decompositions and
 * compositions of assigned characters never change across Unicode versions
 * (UAX #15 §12.1 normalization stability), so for those the host's answer is
 * every engine's. A scalar the host does not know proves nothing (the caller
 * answers `'all'`).
 */

const ASSIGNED = /\p{Assigned}/u;

/** A cluster whose closure grows past this many scalars proves nothing. */
const MAX_CLUSTER_CLOSURE = 64;

/** Whether the host knows `cp` as an assigned scalar (not a surrogate), so
 * its canonical decomposition and compositions are the stable ones. */
export function isCanonicallyKnown(cp: number): boolean {
  if (cp >= 0xd800 && cp <= 0xdfff) return false;
  return ASSIGNED.test(String.fromCodePoint(cp));
}

const MARK = /\p{M}/u;

/** Whether a known scalar that is not normalization-inert attaches to the
 * cluster before it: a combining mark (General_Category M, with or without a
 * decomposition: U+0340 paints as U+0300 on its base), or a scalar with no
 * canonical decomposition, which is then not normalization-inert because it
 * can combine (a nonzero combining class, the second operand of a canonical
 * pair, a Hangul jamo). */
export function isClusterJoiner(cp: number): boolean {
  const text = String.fromCodePoint(cp);
  return MARK.test(text) || text.normalize('NFD') === text;
}

/**
 * The full composition exclusions whose decomposition has two or more
 * scalars (Unicode 17 Full_Composition_Exclusion without the singletons,
 * which no shaper composes to), as inclusive ranges. NFC never composes
 * them, but a shaper may (HarfBuzz's Indic shaper recomposes U+09AF U+09BC
 * to U+09DF), so a cluster holding one's whole decomposition demands it too.
 * Their decompositions are the host's NFD (stable, UAX #15 §12.1).
 */
const EXCLUDED_COMPOSITES: readonly number[] = [
  0x344, 0x344, 0x958, 0x95f, 0x9dc, 0x9dd, 0x9df, 0x9df, 0xa33, 0xa33, 0xa36, 0xa36,
  0xa59, 0xa5b, 0xa5e, 0xa5e, 0xb5c, 0xb5d, 0xf43, 0xf43, 0xf4d, 0xf4d, 0xf52, 0xf52,
  0xf57, 0xf57, 0xf5c, 0xf5c, 0xf69, 0xf69, 0xf73, 0xf73, 0xf75, 0xf76, 0xf78, 0xf78,
  0xf81, 0xf81, 0xf93, 0xf93, 0xf9d, 0xf9d, 0xfa2, 0xfa2, 0xfa7, 0xfa7, 0xfac, 0xfac,
  0xfb9, 0xfb9, 0x2adc, 0x2adc, 0xfb1d, 0xfb1d, 0xfb1f, 0xfb1f, 0xfb2a, 0xfb36,
  0xfb38, 0xfb3c, 0xfb3e, 0xfb3e, 0xfb40, 0xfb41, 0xfb43, 0xfb44, 0xfb46, 0xfb4e,
  0x1d15e, 0x1d164, 0x1d1bb, 0x1d1c0,
];

type Decomposed = readonly [composite: number, parts: readonly number[]];
let excludedDecompositions: readonly Decomposed[] | undefined;

/** Each excluded composite with its full canonical decomposition (first use
 * allocates). */
function excludedComposites(): readonly Decomposed[] {
  if (!excludedDecompositions) {
    const list: Decomposed[] = [];
    for (let at = 0; at < EXCLUDED_COMPOSITES.length; at += 2) {
      for (let cp = EXCLUDED_COMPOSITES[at]!; cp <= EXCLUDED_COMPOSITES[at + 1]!; cp++) {
        list.push([cp, [...String.fromCodePoint(cp).normalize('NFD')].map((part) => part.codePointAt(0) as number)]);
      }
    }
    excludedDecompositions = list;
  }
  return excludedDecompositions;
}

function decompose(cp: number, into: Set<number>): void {
  into.add(cp);
  for (const part of String.fromCodePoint(cp).normalize('NFD')) into.add(part.codePointAt(0) as number);
}

/**
 * Every scalar a canonically equivalent rendering of `cluster` (a starter, or
 * nothing, then joiners) can paint: its scalars, their full decompositions
 * and every pairwise canonical composition of those, closed. Undefined when
 * the closure outgrows a sane cluster.
 */
export function canonicalClusterScalars(cluster: readonly number[]): readonly number[] | undefined {
  const scalars = new Set<number>();
  for (const cp of cluster) decompose(cp, scalars);
  if (scalars.size > MAX_CLUSTER_CLOSURE) return undefined;
  for (let grown = true; grown;) {
    grown = false;
    for (const [cp, parts] of excludedComposites()) {
      if (scalars.has(cp) || !parts.every((part) => scalars.has(part))) continue;
      scalars.add(cp);
      grown = true;
      if (scalars.size > MAX_CLUSTER_CLOSURE) return undefined;
    }
    const current = [...scalars];
    for (const a of current) {
      for (const b of current) {
        const composed = (String.fromCodePoint(a) + String.fromCodePoint(b)).normalize('NFC');
        const cp = composed.codePointAt(0) as number;
        if (composed.length !== String.fromCodePoint(cp).length || scalars.has(cp)) continue;
        decompose(cp, scalars);
        grown = true;
        if (scalars.size > MAX_CLUSTER_CLOSURE) return undefined;
      }
    }
  }
  return [...scalars];
}
