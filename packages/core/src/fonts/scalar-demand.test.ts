import { describe, expect, it } from 'vitest';
import { HARD_MAX_FONT_DEMAND_SCALARS } from '../worker/resource-policy.generated.js';
import { FontDemandScalarCollector, collectFontDemandScalars } from './scalar-demand.js';

const SENTINELS = [0x4d, 0x69, 0x20]; // M, i, space
const BUDGET = 1 << 20;

describe('collectFontDemandScalars', () => {
  it('returns first-seen scalars after the measurement sentinels, deduplicated', () => {
    expect(collectFontDemandScalars(['Ab', 'b漢𠀋'], { budget: BUDGET }))
      .toEqual([...SENTINELS, 0x41, 0x62, 0x6f22, 0x2000b]);
  });

  it('answers all for every text whose painted scalars its spelling cannot prove', () => {
    for (const text of ['\ud800', '͸', '', 'שלום', '‮(']) {
      expect(collectFontDemandScalars(['A', text], { budget: BUDGET }), JSON.stringify(text)).toBe('all');
    }
  });

  it('demands every scalar a canonically equivalent rendering can paint', () => {
    const demanded = (text: string) => {
      const result = collectFontDemandScalars([text], { budget: BUDGET });
      if (result === 'all') throw new Error(`all for ${JSON.stringify(text)}`);
      return new Set(result);
    };
    const covers = (text: string, scalars: readonly number[]) => {
      const set = demanded(text);
      for (const cp of scalars) expect(set.has(cp), `${JSON.stringify(text)} U+${cp.toString(16)}`).toBe(true);
    };
    // Precomposed: itself and its decomposition.
    covers('\u00e9', [0xe9, 0x65, 0x301]);
    // Decomposed: the spelled scalars and the composition.
    covers('e\u0301', [0x65, 0x301, 0xe9]);
    // Voiced kana: the base kana and the combining voiced marks.
    covers('\u304c', [0x304c, 0x304b, 0x3099]);
    covers('\u30d7', [0x30d7, 0x30d5, 0x309a]);
    // Two marks: each partial composition a shaper may stop at.
    covers('A\u0323\u0302', [0x41, 0x323, 0x302, 0x1ea0, 0xc2, 0x1eac]);
    // Hangul: the syllable and its jamo, both ways.
    covers('\ud55c', [0xd55c, 0x1112, 0x1161, 0x11ab, 0xd558]);
    covers('\u1112\u1161\u11ab', [0xd55c, 0xd558]);
    // A singleton decomposition (Angstrom sign) paints U+00C5 or A + ring.
    covers('\u212b', [0x212b, 0xc5, 0x41, 0x30a]);
    // A mark without a base may be painted on a dotted circle.
    covers('\u0301', [0x301, 0x25cc]);
    // Inert text is unchanged.
    expect([...demanded('AB')].filter((cp) => !SENTINELS.includes(cp))).toEqual([0x41, 0x42]);
  });

  it('charges only uncommitted scalars and degrades to all beyond the JSON budget', () => {
    const committed = new Set([...SENTINELS, 0x41]);
    expect(collectFontDemandScalars(['AB'], { budget: BUDGET, committed })).toEqual([0x42]);
    // "[]" alone needs two bytes.
    expect(collectFontDemandScalars([], { budget: 1, committed })).toBe('all');
    expect(collectFontDemandScalars([], { budget: 2, committed })).toEqual([]);
    // "[66]" is four bytes; "[66,67]" seven.
    expect(collectFontDemandScalars(['B'], { budget: 5, committed })).toEqual([0x42]);
    expect(collectFontDemandScalars(['BC'], { budget: 6, committed })).toBe('all');
  });

  it('degrades to all beyond HARD_MAX_FONT_DEMAND_SCALARS, counting committed scalars', () => {
    const committedOf = (size: number) => new Set([...SENTINELS,
      ...Array.from({ length: size - SENTINELS.length }, (_, i) => 0x4e00 + i)]);
    expect(collectFontDemandScalars(['A'], { budget: BUDGET, committed: committedOf(HARD_MAX_FONT_DEMAND_SCALARS - 1) }))
      .toEqual([0x41]);
    expect(collectFontDemandScalars(['A'], { budget: BUDGET, committed: committedOf(HARD_MAX_FONT_DEMAND_SCALARS) }))
      .toBe('all');
  });

  it('stays all once declined, so later texts cannot resurrect a partial list', () => {
    const collector = new FontDemandScalarCollector({ budget: BUDGET });
    expect(collector.add('A')).toBe(true);
    expect(collector.decline()).toBe(false);
    expect(collector.add('B')).toBe(false);
    expect(collector.result()).toBe('all');
  });
});
