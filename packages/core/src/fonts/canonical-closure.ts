/**
 * The scalars a canonically equivalent rendering of one text cluster can
 * paint. Text shaping may decompose a precomposed character the face lacks,
 * or compose a base and marks the face holds precomposed (HarfBuzz
 * normalization; UAX #15 canonical equivalence), so a text that spells `が`
 * (U+304C) can paint `か` U+304B and the voiced mark U+3099, and one that
 * spells `e` U+0301 can paint `é` U+00E9. The closure of a cluster (a
 * starter and the joiners after it) under canonical decomposition and
 * pairwise canonical composition is a superset of every such rendering: a
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

/** Whether a known scalar that is not normalization-inert attaches to the
 * cluster before it: it has no canonical decomposition, so it is not
 * normalization-inert because it can combine (a nonzero combining class, the
 * second operand of a canonical pair, a Hangul jamo). */
export function isClusterJoiner(cp: number): boolean {
  const text = String.fromCodePoint(cp);
  return text.normalize('NFD') === text;
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
  for (let grown = true; grown;) {
    grown = false;
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
