import { describe, expect, it } from 'vitest';
import { createNamedFamilyOracle } from './native-route-line-metrics.js';
import { PROBE_EM, SENTINELS, SENTINEL_UPM } from './route-sentinels.js';

// A structural Canvas double for the route sentinels: the named family
// "Covering" draws every cluster except those containing 'ж'; anything it does
// not draw falls through to the sentinel the font string ends with.
function fakeOwner(options: { lang?: boolean; sentinels?: boolean } = {}) {
  const calls: string[] = [];
  const scale = PROBE_EM / SENTINEL_UPM;
  const sentinelMetrics = (index: number) => {
    const sentinel = SENTINELS[index]!;
    return {
      width: sentinel.advance * scale, actualBoundingBoxLeft: 0, actualBoundingBoxRight: sentinel.advance * scale,
      actualBoundingBoxAscent: sentinel.inkUp * scale, actualBoundingBoxDescent: sentinel.inkDown * scale,
      fontBoundingBoxAscent: sentinel.ascent * scale, fontBoundingBoxDescent: sentinel.descent * scale,
    } as TextMetrics;
  };
  const context: Record<string, unknown> = {
    font: '10px serif',
    measureText(text: string) {
      calls.push(text);
      const font = context.font as string;
      const sentinel = SENTINELS.findIndex((candidate) => font.includes(candidate.family));
      if (sentinel < 0 || options.sentinels === false) {
        return { width: 0, actualBoundingBoxLeft: 0, actualBoundingBoxRight: 0, actualBoundingBoxAscent: 0,
          actualBoundingBoxDescent: 0, fontBoundingBoxAscent: 0, fontBoundingBoxDescent: 0 } as TextMetrics;
      }
      const covered = font.includes('Covering') && !text.includes('ж');
      return covered
        ? { width: 500, actualBoundingBoxLeft: 0, actualBoundingBoxRight: 500, actualBoundingBoxAscent: 700,
          actualBoundingBoxDescent: 200, fontBoundingBoxAscent: 800, fontBoundingBoxDescent: 200 } as TextMetrics
        : sentinelMetrics(sentinel);
    },
  };
  if (options.lang !== false) context.lang = 'inherit';
  const factory = () => ({ getContext: () => context }) as unknown as HTMLCanvasElement;
  return { factory, calls, owner: context };
}

describe('named-family oracle (Canvas lang admission)', () => {
  it('admits text only named families draw, cut at the first generic', () => {
    const { factory, owner } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    expect(oracle.available).toBe(true);
    expect(oracle.draws('"Covering", serif', 400, 'normal', 'где')).toBe(true);
    expect(oracle.draws('"Covering", serif', 400, 'normal', 'гж')).toBe(false);
    // A generic before the named family: nothing named is reached first.
    expect(oracle.draws('serif, "Covering"', 400, 'normal', 'где')).toBe(false);
    // Presentation-selected clusters are never admitted.
    expect(oracle.draws('"Covering"', 400, 'normal', '◼︎')).toBe(false);
    expect(oracle.draws('"Covering"', 400, 'normal', 'ᠠ᠋')).toBe(false);
  });

  it('cannot judge without Canvas lang or working sentinels', () => {
    const noLang = fakeOwner({ lang: false });
    expect(createNamedFamilyOracle(noLang.owner, noLang.factory).available).toBe(false);
    const noSentinels = fakeOwner({ sentinels: false });
    const broken = createNamedFamilyOracle(noSentinels.owner, noSentinels.factory);
    expect(broken.available).toBe(false);
    expect(broken.draws('"Covering"', 400, 'normal', 'где')).toBe(false);
  });

  it('remembers short clusters only, and stays bounded over high-cardinality text', () => {
    const { factory, calls, owner } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    expect(oracle.draws('"Covering"', 400, 'normal', 'г')).toBe(true);
    const afterFirst = calls.length;
    expect(oracle.draws('"Covering"', 400, 'normal', 'г')).toBe(true);
    expect(calls.length).toBe(afterFirst);
    // A cluster longer than the cached bound is proved again each time.
    const long = `a${'́'.repeat(40)}`;
    expect(oracle.draws('"Covering"', 400, 'normal', long)).toBe(true);
    const afterLong = calls.length;
    expect(oracle.draws('"Covering"', 400, 'normal', long)).toBe(true);
    expect(calls.length).toBeGreaterThan(afterLong);
    // Twenty thousand distinct ideographs: every one is still judged.
    for (let code = 0x4e00; code < 0x4e00 + 20000; code += 1) {
      expect(oracle.draws('"Covering"', 400, 'normal', String.fromCodePoint(code))).toBe(true);
    }
  });
});
