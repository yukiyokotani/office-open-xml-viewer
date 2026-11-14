/**
 * The scalars a canonically equivalent rendering of one text cluster can
 * paint. Text shaping may decompose a precomposed character the face lacks,
 * or compose a base and marks the face holds precomposed (HarfBuzz
 * normalization; UAX #15 canonical equivalence), so a text that spells `が`
 * (U+304C) can paint `か` U+304B and the voiced mark U+3099, and one that
 * spells `e` U+0301 can paint `é` U+00E9. The closure of a cluster (a
 * starter and the joiners after it) under one-level canonical decomposition
 * and pairwise composition is a superset of every such rendering: a superset
 * only loads more `unicode-range` subsets.
 *
 * Tables: `canonical-decompositions.ts` (Unicode 17.0.0, generated); Hangul
 * syllables are algorithmic (Unicode Standard §3.12). Host `normalize()` is
 * never used: the painter's Unicode data is not the host's.
 */
import { CANONICAL_DECOMPOSITIONS, CLUSTER_JOINERS } from './canonical-decompositions.js';

const S_BASE = 0xac00, L_BASE = 0x1100, V_BASE = 0x1161, T_BASE = 0x11a7;
const L_COUNT = 19, V_COUNT = 21, T_COUNT = 28, N_COUNT = V_COUNT * T_COUNT, S_COUNT = L_COUNT * N_COUNT;

/** A cluster whose closure grows past this many scalars proves nothing. */
const MAX_CLUSTER_CLOSURE = 64;

interface Tables {
  readonly decompositions: ReadonlyMap<number, readonly number[]>;
  readonly compositions: ReadonlyMap<number, number>;
  /** Sorted inclusive joiner spans, flattened. */
  readonly joiners: readonly number[];
}

let tables: Tables | undefined;

function loadTables(): Tables {
  if (tables) return tables;
  const decompositions = new Map<number, readonly number[]>();
  const compositions = new Map<number, number>();
  let cp = 0;
  for (const entry of CANONICAL_DECOMPOSITIONS.split(';')) {
    const [delta, parts] = entry.split(':') as [string, string];
    cp += parseInt(delta, 36);
    const mapping = parts.split(',').map((part) => parseInt(part, 36));
    decompositions.set(cp, mapping);
    // Every canonical pair, composition exclusions included (a superset).
    if (mapping.length === 2) compositions.set(pairKey(mapping[0]!, mapping[1]!), cp);
  }
  const joiners: number[] = [];
  let end = 0;
  for (const span of CLUSTER_JOINERS.split(',')) {
    const [delta, length = '0'] = span.split('-') as [string, string?];
    const lo = end + parseInt(delta, 36);
    end = lo + parseInt(length, 36);
    joiners.push(lo, end);
  }
  return (tables = { decompositions, compositions, joiners });
}

const pairKey = (a: number, b: number) => a * 0x110000 + b;

/** Whether `cp` attaches to the cluster before it: a nonzero canonical
 * combining class, the second operand of a canonical pair, or a Hangul vowel
 * or trailing jamo. */
export function isClusterJoiner(cp: number): boolean {
  const { joiners } = loadTables();
  let lo = 0, hi = joiners.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (cp < joiners[mid * 2]!) hi = mid - 1;
    else if (cp > joiners[mid * 2 + 1]!) lo = mid + 1;
    else return true;
  }
  return false;
}

/** Whether `cp` is an assigned scalar the tables describe: a canonical
 * decomposition key, a joiner, or a Hangul leading jamo or syllable. Every
 * assigned scalar that is not normalization-inert is one of these. */
export function isCanonicallyDescribed(cp: number): boolean {
  return loadTables().decompositions.has(cp) || isClusterJoiner(cp)
    || (cp >= L_BASE && cp < L_BASE + L_COUNT) || (cp >= S_BASE && cp < S_BASE + S_COUNT);
}

function decompose(cp: number, into: Set<number>): void {
  into.add(cp);
  if (cp >= S_BASE && cp < S_BASE + S_COUNT) {
    const index = cp - S_BASE;
    const t = index % T_COUNT;
    // LV + T, and the LV syllable into L + V.
    if (t) decompose(cp - t, into);
    else {
      into.add(L_BASE + Math.floor(index / N_COUNT));
      into.add(V_BASE + Math.floor((index % N_COUNT) / T_COUNT));
    }
    if (t) into.add(T_BASE + t);
    return;
  }
  for (const part of loadTables().decompositions.get(cp) ?? []) decompose(part, into);
}

function compose(a: number, b: number): number | undefined {
  if (a >= L_BASE && a < L_BASE + L_COUNT && b >= V_BASE && b < V_BASE + V_COUNT) {
    return S_BASE + ((a - L_BASE) * V_COUNT + (b - V_BASE)) * T_COUNT;
  }
  if (a >= S_BASE && a < S_BASE + S_COUNT && (a - S_BASE) % T_COUNT === 0 && b > T_BASE && b < T_BASE + T_COUNT) {
    return a + (b - T_BASE);
  }
  return loadTables().compositions.get(pairKey(a, b));
}

/**
 * Every scalar a canonically equivalent rendering of `cluster` (a starter, or
 * nothing, then joiners) can paint: its scalars, their full decompositions
 * and every pairwise composition of those, closed. Undefined when the closure
 * outgrows a sane cluster.
 */
export function canonicalClusterScalars(cluster: readonly number[]): readonly number[] | undefined {
  const scalars = new Set<number>();
  for (const cp of cluster) decompose(cp, scalars);
  for (let grown = true; grown;) {
    grown = false;
    for (const a of [...scalars]) {
      for (const b of [...scalars]) {
        const composed = compose(a, b);
        if (composed === undefined || scalars.has(composed)) continue;
        decompose(composed, scalars);
        grown = true;
        if (scalars.size > MAX_CLUSTER_CLOSURE) return undefined;
      }
    }
  }
  return [...scalars];
}
