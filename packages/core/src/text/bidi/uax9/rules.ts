// From-scratch implementation of the Unicode Bidirectional Algorithm (UAX#9),
// "isolates" revision. Works in CODE-POINT space and returns per-code-point
// embedding levels (REMOVED for characters deleted by rule X9). Reordering
// (rule L2) is exposed separately as `reorderByLevels`.
//
// Rule references (e.g. "X5a", "N0") are to https://www.unicode.org/reports/tr9/.
// Conformance is verified against Unicode's BidiCharacterTest.txt / BidiTest.txt.

import { bidiClass, bracket } from '../char-data.js';
import type { BaseDirection, BidiClass } from '../types.js';

export const MAX_DEPTH = 125;
/** Level sentinel for code points removed by rule X9 (RLE/LRE/RLO/LRO/PDF/BN). */
export const REMOVED = -1;

type Dir = 'L' | 'R';
type Override = 'neutral' | 'L' | 'R';

const isRemovedType = (t: BidiClass): boolean =>
  t === 'RLE' || t === 'LRE' || t === 'RLO' || t === 'LRO' || t === 'PDF' || t === 'BN';
const isIsolateInitiator = (t: BidiClass): boolean =>
  t === 'LRI' || t === 'RLI' || t === 'FSI';
const isNI = (t: BidiClass): boolean =>
  t === 'B' || t === 'S' || t === 'WS' || t === 'ON' ||
  t === 'FSI' || t === 'LRI' || t === 'RLI' || t === 'PDI';

const leastGreaterOdd = (x: number): number => (x & 1 ? x + 2 : x + 1);
const leastGreaterEven = (x: number): number => (x & 1 ? x + 1 : x + 2);

/** UAX#9 P2-P3 over [start,end): first strong (L/AL/R), skipping isolated runs. */
function firstStrongLevel(types: BidiClass[], start: number, end: number): number {
  let depth = 0;
  for (let i = start; i < end; i++) {
    const t = types[i];
    if (isIsolateInitiator(t)) depth++;
    else if (t === 'PDI') {
      if (depth > 0) depth--;
    } else if (depth === 0) {
      if (t === 'L') return 0;
      if (t === 'R' || t === 'AL') return 1;
    }
  }
  return 0;
}

/** BD9: pair isolate initiators with their matching PDIs. */
function computeMatching(types: BidiClass[]): { pdiOf: Int32Array; initOf: Int32Array } {
  const n = types.length;
  const pdiOf = new Int32Array(n).fill(n); // initiator -> matching PDI (n = none)
  const initOf = new Int32Array(n).fill(-1); // PDI -> matching initiator (-1 = none)
  const stack: number[] = [];
  for (let i = 0; i < n; i++) {
    const t = types[i];
    if (isIsolateInitiator(t)) stack.push(i);
    else if (t === 'PDI' && stack.length) {
      const j = stack.pop()!;
      pdiOf[j] = i;
      initOf[i] = j;
    }
  }
  return { pdiOf, initOf };
}

/** Rules X1-X8: explicit embedding levels + override application. */
function determineExplicitLevels(
  origTypes: BidiClass[],
  paragraphLevel: number,
  pdiOf: Int32Array,
): { levels: number[]; types: BidiClass[] } {
  const n = origTypes.length;
  const levels = new Array<number>(n).fill(paragraphLevel);
  const types = origTypes.slice();

  interface Entry {
    level: number;
    override: Override;
    isolate: boolean;
  }
  const stack: Entry[] = [{ level: paragraphLevel, override: 'neutral', isolate: false }];
  let overflowIsolate = 0;
  let overflowEmbedding = 0;
  let validIsolate = 0;
  const top = (): Entry => stack[stack.length - 1];

  for (let i = 0; i < n; i++) {
    const t = origTypes[i];
    switch (t) {
      case 'RLE':
      case 'LRE':
      case 'RLO':
      case 'LRO': {
        levels[i] = top().level; // removed by X9
        const isRTL = t === 'RLE' || t === 'RLO';
        const newLevel = isRTL ? leastGreaterOdd(top().level) : leastGreaterEven(top().level);
        if (newLevel <= MAX_DEPTH && overflowIsolate === 0 && overflowEmbedding === 0) {
          stack.push({
            level: newLevel,
            override: t === 'RLO' ? 'R' : t === 'LRO' ? 'L' : 'neutral',
            isolate: false,
          });
        } else if (overflowIsolate === 0) {
          overflowEmbedding++;
        }
        break;
      }
      case 'RLI':
      case 'LRI':
      case 'FSI': {
        levels[i] = top().level;
        const ov = top().override;
        if (ov !== 'neutral') types[i] = ov;
        let dir: Dir;
        if (t === 'RLI') dir = 'R';
        else if (t === 'LRI') dir = 'L';
        else dir = firstStrongLevel(origTypes, i + 1, pdiOf[i]) === 1 ? 'R' : 'L';
        const newLevel = dir === 'R' ? leastGreaterOdd(top().level) : leastGreaterEven(top().level);
        if (newLevel <= MAX_DEPTH && overflowIsolate === 0 && overflowEmbedding === 0) {
          validIsolate++;
          stack.push({ level: newLevel, override: 'neutral', isolate: true });
        } else {
          overflowIsolate++;
        }
        break;
      }
      case 'PDI': {
        if (overflowIsolate > 0) {
          overflowIsolate--;
        } else if (validIsolate > 0) {
          overflowEmbedding = 0;
          while (!top().isolate) stack.pop();
          stack.pop();
          validIsolate--;
        }
        {
          levels[i] = top().level;
          const ov = top().override;
          if (ov !== 'neutral') types[i] = ov;
        }
        break;
      }
      case 'PDF': {
        levels[i] = top().level; // removed by X9
        if (overflowIsolate > 0) {
          /* no-op */
        } else if (overflowEmbedding > 0) {
          overflowEmbedding--;
        } else if (!top().isolate && stack.length >= 2) {
          stack.pop();
        }
        break;
      }
      case 'B': {
        // X8: a paragraph separator terminates all embeddings/overrides/isolates.
        stack.length = 1;
        overflowIsolate = 0;
        overflowEmbedding = 0;
        validIsolate = 0;
        levels[i] = paragraphLevel;
        break;
      }
      case 'BN': {
        levels[i] = top().level; // removed by X9
        break;
      }
      default: {
        // X6: all other types.
        levels[i] = top().level;
        const ov = top().override;
        if (ov !== 'neutral') types[i] = ov;
        break;
      }
    }
  }
  return { levels, types };
}

/** Canonical-equivalence fold for BD16 bracket matching (only two pairs apply). */
function canonBracket(cp: number): number {
  if (cp === 0x2329) return 0x3008;
  if (cp === 0x232a) return 0x3009;
  return cp;
}

interface Sequence {
  indices: number[]; // original code-point indices, logical order
  level: number;
  sos: Dir;
  eos: Dir;
}

/** X10 + BD13: build isolating run sequences with their sos/eos directions. */
function buildSequences(
  origTypes: BidiClass[],
  levels: number[],
  removed: boolean[],
  paragraphLevel: number,
  pdiOf: Int32Array,
  initOf: Int32Array,
): Sequence[] {
  const n = origTypes.length;
  const kept: number[] = [];
  for (let i = 0; i < n; i++) if (!removed[i]) kept.push(i);

  // Level runs: maximal same-level spans of the reduced (X9) text.
  const runs: number[][] = [];
  for (let k = 0; k < kept.length; k++) {
    const idx = kept[k];
    if (k === 0 || levels[idx] !== levels[kept[k - 1]]) runs.push([idx]);
    else runs[runs.length - 1].push(idx);
  }
  const runByFirst = new Map<number, number[]>();
  for (const r of runs) runByFirst.set(r[0], r);

  const dirOf = (level: number): Dir => (level & 1 ? 'R' : 'L');
  const sequences: Sequence[] = [];

  for (const run of runs) {
    const firstIdx = run[0];
    // A run beginning with a PDI that matches an initiator is appended to that
    // initiator's sequence, not a sequence start of its own.
    if (origTypes[firstIdx] === 'PDI' && initOf[firstIdx] !== -1) continue;

    const indices: number[] = [];
    let cur = run;
    for (;;) {
      for (const idx of cur) indices.push(idx);
      const lastIdx = cur[cur.length - 1];
      if (isIsolateInitiator(origTypes[lastIdx]) && pdiOf[lastIdx] !== n) {
        const next = runByFirst.get(pdiOf[lastIdx]);
        if (next) {
          cur = next;
          continue;
        }
      }
      break;
    }

    const seqLevel = levels[indices[0]];
    // sos: compare seq level with the level of the preceding non-removed char.
    let prevLevel = paragraphLevel;
    for (let j = indices[0] - 1; j >= 0; j--) {
      if (!removed[j]) {
        prevLevel = levels[j];
        break;
      }
    }
    const sos = dirOf(Math.max(seqLevel, prevLevel));

    // eos: paragraph level if the sequence ends in an isolate initiator with no
    // matching PDI; otherwise the level of the following non-removed char.
    const lastIdx = indices[indices.length - 1];
    let nextLevel = paragraphLevel;
    if (!(isIsolateInitiator(origTypes[lastIdx]) && pdiOf[lastIdx] === n)) {
      for (let j = lastIdx + 1; j < n; j++) {
        if (!removed[j]) {
          nextLevel = levels[j];
          break;
        }
      }
    }
    const eos = dirOf(Math.max(seqLevel, nextLevel));

    sequences.push({ indices, level: seqLevel, sos, eos });
  }
  return sequences;
}

const strongDirOf = (t: BidiClass): Dir | null =>
  t === 'L' ? 'L' : t === 'R' ? 'R' : t === 'EN' || t === 'AN' ? 'R' : null;

/** Rules W1-W7, N0-N2, I1-I2 over one isolating run sequence (mutates `wt`/`levels`). */
function resolveSequence(
  seq: Sequence,
  cps: number[],
  origTypes: BidiClass[],
  wt: BidiClass[],
  levels: number[],
): void {
  const ind = seq.indices;
  const m = ind.length;
  const { sos, eos, level } = seq;

  // W1: NSM -> type of previous char (ON if previous is isolate initiator/PDI; sos if first).
  for (let p = 0; p < m; p++) {
    const idx = ind[p];
    if (wt[idx] === 'NSM') {
      if (p === 0) wt[idx] = sos;
      else {
        const pv = wt[ind[p - 1]];
        wt[idx] = pv === 'LRI' || pv === 'RLI' || pv === 'FSI' || pv === 'PDI' ? 'ON' : pv;
      }
    }
  }

  // UAX#9 Unicode 17, W2 then W3: retain the pre-W3 strong type in
  // lastStrong, so converting a stored AL to R cannot hide it from later EN.
  {
    let lastStrong: BidiClass = sos;
    for (let p = 0; p < m; p++) {
      const t = wt[ind[p]];
      if (t === 'R' || t === 'L' || t === 'AL') lastStrong = t;
      else if (t === 'EN' && lastStrong === 'AL') wt[ind[p]] = 'AN';
      if (t === 'AL') wt[ind[p]] = 'R';
    }
  }

  // W4: a single ES between EN EN, or CS between EN EN / AN AN, becomes that number type.
  for (let p = 1; p < m - 1; p++) {
    const t = wt[ind[p]];
    const a = wt[ind[p - 1]];
    const b = wt[ind[p + 1]];
    if (t === 'ES' && a === 'EN' && b === 'EN') wt[ind[p]] = 'EN';
    else if (t === 'CS' && a === 'EN' && b === 'EN') wt[ind[p]] = 'EN';
    else if (t === 'CS' && a === 'AN' && b === 'AN') wt[ind[p]] = 'AN';
  }

  // W5: a sequence of ET adjacent to EN becomes EN.
  for (let p = 0; p < m; p++) {
    if (wt[ind[p]] !== 'ET') continue;
    let q = p;
    while (q < m && wt[ind[q]] === 'ET') q++;
    const before = p > 0 ? wt[ind[p - 1]] : sos;
    const after = q < m ? wt[ind[q]] : eos;
    if (before === 'EN' || after === 'EN') for (let r = p; r < q; r++) wt[ind[r]] = 'EN';
    p = q - 1;
  }

  // W6 then W7 share a traversal: W6 changes no strong L/R context.
  {
    let lastStrong: BidiClass = sos;
    for (let p = 0; p < m; p++) {
      const t = wt[ind[p]];
      if (t === 'ES' || t === 'ET' || t === 'CS') wt[ind[p]] = 'ON';
      if (t === 'R' || t === 'L') lastStrong = t;
      else if (t === 'EN' && lastStrong === 'L') wt[ind[p]] = 'L';
    }
  }

  // N0: paired brackets (BD16 pairing + resolution).
  const e: Dir = level & 1 ? 'R' : 'L';
  const o: Dir = e === 'R' ? 'L' : 'R';
  const pairs: { open: number; close: number }[] = [];
  {
    const stack: { expect: number; pos: number }[] = [];
    outer: for (let p = 0; p < m; p++) {
      const idx = ind[p];
      if (wt[idx] !== 'ON') continue;
      const b = bracket(cps[idx]);
      if (!b) continue;
      if (b.type === 'o') {
        if (stack.length === 63) break outer; // BD16 stack limit -> stop
        stack.push({ expect: canonBracket(b.pair), pos: p });
      } else {
        const cc = canonBracket(cps[idx]);
        for (let s = stack.length - 1; s >= 0; s--) {
          if (stack[s].expect === cc) {
            pairs.push({ open: stack[s].pos, close: p });
            stack.length = s;
            break;
          }
        }
      }
    }
    pairs.sort((x, y) => x.open - y.open);
  }
  const applyNsmAfter = (p: number, dir: Dir): void => {
    for (let k = p + 1; k < m && origTypes[ind[k]] === 'NSM'; k++) wt[ind[k]] = dir;
  };
  for (const { open, close } of pairs) {
    let foundE = false;
    let foundO = false;
    for (let k = open + 1; k < close; k++) {
      const d = strongDirOf(wt[ind[k]]);
      if (d === e) foundE = true;
      else if (d === o) foundO = true;
    }
    let setDir: Dir | null = null;
    if (foundE) setDir = e;
    else if (foundO) {
      let ctx: Dir = sos;
      for (let k = open - 1; k >= 0; k--) {
        const d = strongDirOf(wt[ind[k]]);
        if (d) {
          ctx = d;
          break;
        }
      }
      setDir = ctx === o ? o : e;
    }
    if (setDir) {
      wt[ind[open]] = setDir;
      wt[ind[close]] = setDir;
      applyNsmAfter(open, setDir);
      applyNsmAfter(close, setDir);
    }
  }

  // N1 then N2, after all N0 bracket/NSM updates: each NI run takes equal
  // surrounding directions (EN/AN count as R), otherwise its embedding direction.
  for (let p = 0; p < m; p++) {
    if (!isNI(wt[ind[p]])) continue;
    let q = p;
    while (q < m && isNI(wt[ind[q]])) q++;
    const before = p > 0 ? strongDirOf(wt[ind[p - 1]]) : sos;
    const after = q < m ? strongDirOf(wt[ind[q]]) : eos;
    const direction = before && after && before === after ? before : e;
    for (let r = p; r < q; r++) wt[ind[r]] = direction;
    p = q - 1;
  }

  // I1/I2: implicit levels.
  for (let p = 0; p < m; p++) {
    const idx = ind[p];
    const t = wt[idx];
    if ((level & 1) === 0) {
      if (t === 'R') levels[idx] += 1;
      else if (t === 'AN' || t === 'EN') levels[idx] += 2;
    } else {
      if (t === 'L' || t === 'EN' || t === 'AN') levels[idx] += 1;
    }
  }
}

/**
 * Resolve per-code-point embedding levels for one paragraph (rules P2-I2 + L1).
 * `levels[i]` is REMOVED for code points deleted by rule X9.
 */
export function resolveLevels(
  cps: number[],
  base: BaseDirection,
  classOverride?: ReadonlyArray<BidiClass | null | undefined>,
): { levels: number[]; paragraphLevel: number } {
  const n = cps.length;
  const origTypes: BidiClass[] = new Array(n);
  // UAX#9 §4.3 HL1 higher-level protocol: a caller may override the assigned
  // Bidi_Class of selected code points BEFORE the algorithm runs (e.g. Word
  // classifying European digits as AN inside Arabic-language complex-script
  // context — see W2). The pure algorithm is unchanged when `classOverride` is
  // absent, so the UAX#9 conformance suite (which passes no override) still
  // drives `resolveLevels` exactly as the standard specifies.
  for (let i = 0; i < n; i++) {
    origTypes[i] = classOverride?.[i] ?? bidiClass(cps[i]);
  }

  const paragraphLevel =
    base === 'rtl' ? 1 : base === 'ltr' ? 0 : firstStrongLevel(origTypes, 0, n);

  const { pdiOf, initOf } = computeMatching(origTypes);
  const { levels, types } = determineExplicitLevels(origTypes, paragraphLevel, pdiOf);

  const removed: boolean[] = new Array(n);
  for (let i = 0; i < n; i++) removed[i] = isRemovedType(origTypes[i]);

  const sequences = buildSequences(origTypes, levels, removed, paragraphLevel, pdiOf, initOf);
  // determineExplicitLevels already owns a fresh mutable copy; origTypes
  // remains unchanged for N0's original NSM facts and L1 resets.
  const wt = types;
  for (const seq of sequences) resolveSequence(seq, cps, origTypes, wt, levels);

  // L1: reset segment/paragraph separators and trailing/leading-to-separator
  // whitespace + isolate formatting to the paragraph level (using ORIGINAL types).
  const isResettable = (i: number): boolean => {
    const t = origTypes[i];
    return t === 'WS' || isIsolateInitiator(t) || t === 'PDI' || removed[i];
  };
  for (let i = 0; i < n; i++) {
    const t = origTypes[i];
    if (t === 'B' || t === 'S') {
      levels[i] = paragraphLevel;
      for (let j = i - 1; j >= 0 && isResettable(j); j--) {
        if (!removed[j]) levels[j] = paragraphLevel;
      }
    }
  }
  for (let j = n - 1; j >= 0 && isResettable(j); j--) {
    if (!removed[j]) levels[j] = paragraphLevel;
  }

  for (let i = 0; i < n; i++) if (removed[i]) levels[i] = REMOVED;

  return { levels, paragraphLevel };
}

/**
 * Rule L2: given resolved levels for the range [start, end), return the visual
 * order as a permutation of logical indices (REMOVED code points are skipped).
 * Works for any level array whose removed sentinel is < 0 or > MAX_DEPTH + 1
 * (so both the code-point REMOVED=-1 and the code-unit 255 are excluded).
 */
export function reorderByLevels(
  levels: ArrayLike<number>,
  start: number,
  end: number,
): number[] {
  const order: number[] = [];
  for (let i = start; i < end; i++) {
    const l = levels[i];
    // Valid resolved levels reach MAX_DEPTH + 1: I1/I2 add up to +2 on top of
    // the explicit-level cap (UAX#9 §3.3.4). Anything above is a removed
    // sentinel (code-point REMOVED=-1 maps to <0; code-unit sentinel is 255).
    if (l >= 0 && l <= MAX_DEPTH + 1) order.push(i);
  }
  if (order.length === 0) return order;

  let highest = 0;
  let lowestOdd = MAX_DEPTH + 2;
  for (const idx of order) {
    const l = levels[idx];
    if (l > highest) highest = l;
    if (l & 1 && l < lowestOdd) lowestOdd = l;
  }
  for (let lvl = highest; lvl >= lowestOdd; lvl--) {
    let i = 0;
    while (i < order.length) {
      if (levels[order[i]] >= lvl) {
        let j = i + 1;
        while (j < order.length && levels[order[j]] >= lvl) j++;
        for (let a = i, b = j - 1; a < b; a++, b--) {
          const tmp = order[a];
          order[a] = order[b];
          order[b] = tmp;
        }
        i = j;
      } else {
        i++;
      }
    }
  }
  return order;
}
