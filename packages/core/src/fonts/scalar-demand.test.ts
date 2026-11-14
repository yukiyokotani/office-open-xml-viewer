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
    // A combining mark with a decomposition still attaches to its base.
    covers('e\u0340', [0x65, 0x340, 0x300, 0xe8]);
    covers('A\u0344', [0x41, 0x344, 0x308, 0x301, 0xc4]);
    // A composition exclusion a shaper may still compose (HarfBuzz Indic).
    covers('\u09af\u09bc', [0x9af, 0x9bc, 0x9df]);
    covers('\u0915\u093c', [0x958]);
    // Inert text is unchanged.
    expect([...demanded('AB')].filter((cp) => !SENTINELS.includes(cp))).toEqual([0x41, 0x42]);
  });

  it('closes a leading joiner run with every demanded scalar layout may shape it after', () => {
    const scalars = (texts: string[]) => {
      const result = collectFontDemandScalars(texts, { budget: BUDGET });
      if (result === 'all') throw new Error(`all for ${JSON.stringify(texts)}`);
      return result;
    };
    // Adjacent runs of one format.
    for (const cp of [0x65, 0x301, 0xe9, 0x25cc]) expect(scalars(['e', '\u0301'])).toContain(cp);
    // A case-mapped run visited in another order, or a deleted run between.
    for (const cp of [0xc1, 0xe1]) expect(scalars(['\u0301', 'a', 'A'])).toContain(cp);
    for (const cp of [0x103, 0x1eaf]) expect(scalars(['a', 'x', '\u0306\u0301'])).toContain(cp);
    // A leading run followed by a starter in the same text.
    expect(scalars(['e', '\u0301x'])).toContain(0xe9);
    expect(scalars(['\u1112', '\u1161\u11abx'])).toContain(0xd55c);
    // Runs that follow one another, to a fixed point.
    expect(scalars(['a', '\u0306', '\u0301'])).toContain(0x1eaf);
    expect(scalars(['\u1112', '\u1161', '\u11ab'])).toContain(0xd55c);
    // An excluded composition a shaper may still make across runs.
    expect(scalars(['\u09af', '\u09bc'])).toContain(0x9df);
    // A precomposed starter decomposes and recomposes with the run.
    expect(scalars(['\u00e9', '\u0323'])).toContain(0x1eb9);
  });

  it('demands the dotted circle for any mark or joiner', () => {
    for (const text of [' \u09cb', '\uac00\u302e\u302e', '\u0915\u093e', '\u1100\u1161']) {
      expect(collectFontDemandScalars([text], { budget: BUDGET }), JSON.stringify(text)).toContain(0x25cc);
    }
    expect(collectFontDemandScalars(['AB'], { budget: BUDGET })).not.toContain(0x25cc);
  });

  it('demands every multi-scalar composition exclusion from its decomposition', () => {
    // Unicode 17 Full_Composition_Exclusion, singletons left out.
    const excluded = ('344 958-95f 9dc-9dd 9df a33 a36 a59-a5b a5e b5c-b5d f43 f4d f52 f57 f5c f69 f73 f75-f76 f78 f81 '
      + 'f93 f9d fa2 fa7 fac fb9 2adc fb1d fb1f fb2a-fb36 fb38-fb3c fb3e fb40-fb41 fb43-fb44 fb46-fb4e 1d15e-1d164 1d1bb-1d1c0')
      .split(' ').flatMap((range) => {
        const [lo, hi = lo] = range.split('-').map((hex) => parseInt(hex, 16)) as [number, number?];
        return Array.from({ length: hi! - lo + 1 }, (_, at) => lo + at);
      });
    expect(excluded).toHaveLength(85);
    // The Hebrew presentation forms are strong RTL text, which declines.
    for (const cp of excluded.filter((cp) => cp < 0xfb1d || cp > 0xfb4e)) {
      const text = String.fromCodePoint(cp).normalize('NFD');
      expect(collectFontDemandScalars([text], { budget: BUDGET }), `U+${cp.toString(16)}`).toContain(cp);
    }
  });

  it('declines, without throwing, a long leading run of one mark', () => {
    expect(collectFontDemandScalars(['\u0301'.repeat(150_000)], { budget: Number.POSITIVE_INFINITY })).not.toBe('all');
    const marks = String.fromCodePoint(...[...Array(80)].map((_, i) => 0x300 + i));
    expect(collectFontDemandScalars(['a', marks], { budget: Number.POSITIVE_INFINITY })).toBe('all');
  });

  it('declines a cluster too large to close before composing it', () => {
    const marks = String.fromCodePoint(...[...Array(64)].map((_, i) => 0x300 + i), ...[...Array(32)].map((_, i) => 0x350 + i));
    const text = (' ' + marks).repeat(1000);
    expect(collectFontDemandScalars([text], { budget: BUDGET })).toBe('all');
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
