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
 * The blocks holding every full composition exclusion whose decomposition
 * has two or more scalars (Unicode Full_Composition_Exclusion without the
 * singletons, which no shaper composes to: U+0344, the Indic and Tibetan
 * nukta and vowel forms, U+2ADC, the Hebrew presentation forms and the
 * musical symbols), as inclusive ranges. NFC never composes them, but a
 * shaper may (HarfBuzz's Indic shaper recomposes U+09AF U+09BC to U+09DF),
 * so a cluster holding one's whole decomposition demands it too. Which
 * scalars of these blocks they are, and their decompositions, are the
 * host's (stable, UAX #15 §12.1): a multi-scalar NFD that NFC does not
 * recompose to the scalar.
 */
const EXCLUSION_BLOCKS = [0x344, 0x344, 0x958, 0xfb9, 0x2adc, 0x2adc, 0xfb1d, 0xfb4e, 0x1d15e, 0x1d1c0];

type Decomposed = readonly [composite: number, parts: readonly number[]];
let excludedDecompositions: readonly Decomposed[] | undefined;
let excludedPartSet: Set<number> | undefined;

/** Each excluded composite with its full canonical decomposition (first use
 * scans {@link EXCLUSION_BLOCKS}). */
function excludedComposites(): readonly Decomposed[] {
  if (!excludedDecompositions) {
    const list: Decomposed[] = [];
    excludedPartSet = new Set();
    for (let at = 0; at < EXCLUSION_BLOCKS.length; at += 2) {
      for (let cp = EXCLUSION_BLOCKS[at]!; cp <= EXCLUSION_BLOCKS[at + 1]!; cp++) {
        const text = String.fromCodePoint(cp);
        const nfd = text.normalize('NFD');
        const parts = [...nfd].map((part) => part.codePointAt(0) as number);
        if (parts.length < 2 || nfd.normalize('NFC') === text) continue;
        list.push([cp, parts]);
        for (const part of parts) excludedPartSet.add(part);
      }
    }
    excludedDecompositions = list;
  }
  return excludedDecompositions;
}

/** Whether `starter` followed by `joiner` may paint a composite: their
 * canonical forms compose (NFC holds fewer scalars than NFD), or `joiner` is
 * part of an excluded composite a shaper may still compose. */
export function mayCompose(starter: number, joiner: number): boolean {
  const text = String.fromCodePoint(starter, joiner);
  if ([...text.normalize('NFC')].length !== [...text.normalize('NFD')].length) return true;
  excludedComposites();
  return excludedPartSet!.has(joiner);
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
