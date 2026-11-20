import { describe, expect, it } from 'vitest';
import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import { createFontResolver } from '../layout/font-service.js';
import { createTextLayoutService, sliceTextShapeRequest } from '../layout/text.js';
import {
  CONDENSED_CONTEXT_GRAPHEMES,
  condensedFloorAdvancesPx,
  condensedUnitAdvancesPx,
  contextualGraphemeAdvances,
  segAdvanceWidth,
  type CondensedAdvanceUnits,
  type NaturalClusterAdvance,
} from './advance.js';
import type { LayoutTextSeg } from './model.js';

// #1986: the pair-context floor screen must reproduce the exact K-context
// floor (#1975) bit for bit. `exactAdvances` is the pre-#1986
// contextualGraphemeAdvances verbatim (eager, charged per window).

function exactAdvances(
  segment: LayoutTextSeg, start: number, end: number, fontSizePx: number,
  reserveWork: (units: number) => void,
): NaturalClusterAdvance[] {
  const offsets = [...new Set([
    start, ...graphemeClusterOffsets(segment.text.slice(start, end)).map((o) => start + o), end,
  ])].sort((a, b) => a - b);
  const width = (from: number, to: number): number => {
    if (to <= from) return 0;
    reserveWork(to - from);
    return segment.textLayoutService!.shape({
      ...sliceTextShapeRequest(segment.textShapeRequest!, from, to),
      fontSizePt: fontSizePx, measure: true, clusterGeometry: false,
    }).advancePt;
  };
  const clusters: NaturalClusterAdvance[] = [];
  for (let index = 0; index < offsets.length - 1; index += 1) {
    const contextStart = offsets[Math.max(0, index - CONDENSED_CONTEXT_GRAPHEMES)]!;
    clusters.push({
      text: segment.text.slice(offsets[index]!, offsets[index + 1]!),
      advancePx: width(contextStart, offsets[index + 1]!) - width(contextStart, offsets[index]!),
    });
  }
  return clusters;
}

function rng(seed: number): () => number {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const EM: Record<string, number> = {
  i: 0.22, l: 0.22, j: 0.22, f: 0.3, t: 0.33, '.': 0.25, ',': 0.25, ' ': 0.25,
  A: 0.67, V: 0.67, W: 0.94, T: 0.61, o: 0.56, y: 0.5, M: 0.83, e: 0.56,
  '́': 0, '‍': 0, 'あ': 1, '漢': 1,
};
const em = (ch: string): number => EM[ch] ?? 0.5;
const KERN: Record<string, number> = { AV: -0.08, VA: -0.08, To: -0.12, Ty: -0.1, 'f.': -0.06, 'y.': -0.07, WA: -0.06, 'l.': -0.02 };

/** Glyph models (em units of the requested size). The realistic ones change a
 * grapheme's increment beyond its preceding grapheme only by small amounts
 * (ligature components, a 0.01 em per-'l' long context); `adversarial` has a
 * 0.6 em trigram effect, outside the screen's stated bound. */
const MEASURERS: Record<string, (text: string) => number> = {
  table: (s) => [...s].reduce((sum, ch) => sum + em(ch), 0),
  pairKerning: (s) => [...s].reduce((sum, ch, i, all) => sum + em(ch) + (KERN[ch + (all[i + 1] ?? '')] ?? 0), 0),
  ligatures: (s) => {
    let sum = MEASURERS.pairKerning!(s);
    for (const [lig, delta] of [['ffi', -0.08], ['ff', -0.04], ['fi', -0.05], ['fl', -0.05]] as const) {
      sum += delta * s.split(lig).length - delta;
      s = s.split(lig).join('#');
    }
    return sum + [...s].reduce((n, ch, i, all) => n + (ch === 'l' ? 0.01 * all.slice(0, i).filter((c) => c === 'l').length : 0), 0);
  },
  adversarial: (s) => MEASURERS.pairKerning!(s) - 0.6 * (s.split('Toy').length - 1),
};

function makeSegment(measure: (text: string) => number, text: string, fontSizePt: number,
  charSpacing: number, charScale: number): LayoutTextSeg {
  const textLayoutService = createTextLayoutService({
    fonts: createFontResolver([]),
    measurer: {
      fingerprint: 'condensed-floor-screen',
      measure: (request) => ({ advancePt: measure(request.text) * request.fontSizePt, ascentPt: 8, descentPt: 2 }),
    },
  });
  return {
    text, fontSize: fontSizePt, charSpacing, charScale, textLayoutService,
    textShapeRequest: { text, fontSizePt, fonts: { ascii: 'Stub', highAnsi: 'Stub', eastAsia: 'Stub', complexScript: 'Stub' } },
  } as unknown as LayoutTextSeg;
}

const ALPHABET = ['A', 'V', 'T', 'o', 'y', 'f', 'i', 'l', '.', ',', ' ', 'W', 'M', 'e', 'é', 'ff', 'ffi', 'Toy', 'あ', '漢', 'l‍l'];

function* cases(seed: number, count: number) {
  const random = rng(seed);
  const pick = <T,>(list: readonly T[]): T => list[Math.floor(random() * list.length)]!;
  for (let n = 0; n < count; n += 1) {
    let text = '';
    const length = 1 + Math.floor(random() * 14);
    for (let k = 0; k < length; k += 1) text += pick(ALPHABET);
    const fontSizePt = pick([8, 10.5, 11, 12, 20, 24]);
    // Negative spacing from a hairline to past the widest glyph.
    const charSpacing = -Math.round(random() * fontSizePt * 1.1 * 20) / 20 - 0.05;
    yield { text, fontSizePt, charSpacing, charScale: pick([1, 1, 0.5, 0.9, 1.5, 2]) };
  }
}

interface Outcome {
  width: number;
  floored: boolean;
  advances: readonly number[];
  work: number[];
}

function outcome(seg: LayoutTextSeg, start: number, end: number, exact: boolean): Outcome {
  const work: number[] = [];
  const clusters = exact
    ? exactAdvances(seg, start, end, seg.fontSize, (units) => work.push(units))
    : contextualGraphemeAdvances(seg, start, end, seg.fontSize, () => { throw new Error('unshaped'); },
      (units) => work.push(units));
  const slice = start === 0 && end === seg.text.length ? seg : { ...seg, text: seg.text.slice(start, end),
    textShapeRequest: { ...seg.textShapeRequest!, text: seg.text.slice(start, end) } } as LayoutTextSeg;
  const units: CondensedAdvanceUnits = { kind: 'natural-clusters', clusters };
  const natural = slice.textLayoutService!.shape({ ...slice.textShapeRequest!, measure: true }).advancePt;
  // Pre-#1986 decision: any signed unit advance below zero (exact units carry
  // no estimate, so condensedFloorAdvancesPx takes the exact path for them).
  const floored = exact
    ? condensedUnitAdvancesPx(slice, undefined, 1, units).some((advance) => advance < 0)
    : condensedFloorAdvancesPx(slice, undefined, 1, units) !== undefined;
  return {
    width: segAdvanceWidth(slice, natural, undefined, 1, units),
    floored,
    work,
    // Read last: the exact values whenever a consumer asks for them.
    advances: clusters.map((cluster) => cluster.advancePx),
  };
}

/** Largest scaled amount by which context before the preceding grapheme
 * lowers a cluster's increment, against the segment's condensing. */
function beyondPairExcess(seg: LayoutTextSeg, start: number, end: number): number {
  const clusters = contextualGraphemeAdvances(seg, start, end, seg.fontSize, () => 0, () => {});
  return Math.max(0, ...clusters.map((c) => ((c.pairAdvancePx ?? c.advancePx) - c.advancePx) * (seg.charScale ?? 1)))
    + (seg.charSpacing ?? 0);
}

describe('#1986 pair-context condensed floor screen', () => {
  for (const [name, measure] of Object.entries(MEASURERS)) {
    it(`matches the exact K-context floor (${name})`, () => {
      let floors = 0;
      let checked = 0;
      for (const { text, fontSizePt, charSpacing, charScale } of cases(name.length * 7919, 1500)) {
        const seg = makeSegment(measure, text, fontSizePt, charSpacing, charScale);
        const offsets = [0, ...graphemeClusterOffsets(text), text.length];
        for (const start of new Set([0, offsets[Math.floor(offsets.length / 2)]!])) {
          const reference = outcome(seg, start, text.length, true);
          const screened = outcome(makeSegment(measure, text, fontSizePt, charSpacing, charScale), start, text.length, false);
          // The screen's bound: identity is claimed whenever no cluster's
          // beyond-pair change exceeds the condensing (always, for the
          // realistic models).
          if (name === 'adversarial' && beyondPairExcess(seg, start, text.length) > 0) continue;
          checked += 1;
          if (reference.floored) floors += 1;
          expect(screened.floored, `${JSON.stringify(text)} ${charSpacing}`).toBe(reference.floored);
          expect(Object.is(screened.width, reference.width), `${JSON.stringify(text)} ${charSpacing}`).toBe(true);
          expect(screened.advances).toEqual(reference.advances);
          expect(screened.work).toEqual(reference.work);
        }
      }
      // Both sides of the threshold are exercised.
      expect(floors).toBeGreaterThan(100);
      expect(checked - floors).toBeGreaterThan(100);
    });
  }

  it('measures no K-context window for a run the screen clears', () => {
    const shaped: string[] = [];
    const seg = makeSegment((text) => { shaped.push(text); return MEASURERS.pairKerning!(text); },
      'AVAVAVAVAVAV', 20, -1, 1);
    const clusters = contextualGraphemeAdvances(seg, 0, seg.text.length, 20, () => 0, () => {});
    expect(condensedFloorAdvancesPx(seg, undefined, 1, { kind: 'natural-clusters', clusters })).toBeUndefined();
    expect(Math.max(...shaped.map((text) => text.length))).toBe(2);
    // Read on demand, the exact values are those of the pre-#1986 path.
    expect(clusters.map((c) => c.advancePx))
      .toEqual(exactAdvances(seg, 0, seg.text.length, 20, () => {}).map((c) => c.advancePx));
  });
});
