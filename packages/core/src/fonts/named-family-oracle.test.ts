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
  // Each canvas gets its own context, as in a browser.
  const makeContext = (): Record<string, unknown> => {
    const context: Record<string, unknown> = {
      font: '10px serif',
      measureText(text: string) {
        calls.push(text);
        const font = context.font as string;
        const sentinel = SENTINELS.findIndex((candidate) => font.includes(candidate.family));
        if ((globalThis as { __nanControls?: boolean }).__nanControls && !font.includes('Covering') && sentinel >= 0) {
          return { width: NaN, actualBoundingBoxLeft: NaN, actualBoundingBoxRight: NaN, actualBoundingBoxAscent: NaN,
            actualBoundingBoxDescent: NaN, fontBoundingBoxAscent: NaN, fontBoundingBoxDescent: NaN } as TextMetrics;
        }
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
    return context;
  };
  const context = makeContext();
  let fontAssignments = 0;
  const factory = () => {
    const created = makeContext();
    let font = created.font as string;
    Object.defineProperty(created, 'font', {
      get: () => font,
      set: (value: string) => { fontAssignments += 1; font = value; },
    });
    return { getContext: () => created } as unknown as HTMLCanvasElement;
  };
  return { factory, calls, owner: context, fontAssignments: () => fontAssignments };
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

  it('assigns each proof context its font once over many cold clusters of one route', () => {
    const { factory, owner, fontAssignments } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    expect(oracle.available).toBe(true);
    const before = fontAssignments();
    for (let code = 0x4e00; code < 0x4e00 + 9000; code += 1) {
      expect(oracle.draws('"Covering"', 400, 'normal', String.fromCodePoint(code))).toBe(true);
    }
    // Two route contexts and two control contexts, each set once.
    expect(fontAssignments() - before).toBeLessThanOrEqual(4);
  });

  it('takes no ownership evidence from non-finite control geometry', () => {
    const { factory, owner } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    expect(oracle.available).toBe(true);
    // Every proof context now reports NaN advances for the sentinel-only
    // controls while the route pair stays equal and finite.
    const broken = (globalThis as { __nanControls?: boolean });
    broken.__nanControls = true;
    try {
      expect(oracle.draws('"Covering"', 400, 'normal', 'й')).toBe(false);
    } finally {
      delete broken.__nanControls;
    }
  });

  it('cannot judge without Canvas lang or working sentinels', () => {
    const noLang = fakeOwner({ lang: false });
    expect(createNamedFamilyOracle(noLang.owner, noLang.factory).available).toBe(false);
    const noSentinels = fakeOwner({ sentinels: false });
    const broken = createNamedFamilyOracle(noSentinels.owner, noSentinels.factory);
    expect(broken.available).toBe(false);
    expect(broken.draws('"Covering"', 400, 'normal', 'где')).toBe(false);
  });

  it('verifies the sentinels on the measuring context and on both proof contexts', () => {
    // Registered and applied where layout measures, but not where the proof
    // runs: the capability is absent, for a run language and for a worker.
    const measuring = fakeOwner();
    const broken = fakeOwner({ sentinels: false });
    const split = createNamedFamilyOracle(measuring.owner, broken.factory);
    expect([split.sentinelsApply, split.available]).toEqual([false, false]);
    expect(split.draws('"Covering"', 400, 'normal', 'где')).toBe(false);
    // Sentinels applied without Canvas `lang`: the realm can prove ownership
    // (so a worker may proceed) and no run language is admitted.
    const noLang = fakeOwner({ lang: false });
    const oracle = createNamedFamilyOracle(noLang.owner, noLang.factory);
    expect([oracle.sentinelsApply, oracle.available]).toEqual([true, false]);
    // The measuring context's own font is restored.
    expect(createNamedFamilyOracle(measuring.owner, fakeOwner().factory).available).toBe(true);
    expect(measuring.owner.font).toBe('10px serif');
  });

  it('remembers short clusters only: native probes per request and per cache state', () => {
    const { factory, calls, owner } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    // The capability is judged on first use (six sentinel measurements).
    expect(oracle.available).toBe(true);
    const probes = (text: string, list = '"Covering"') => {
      const before = calls.length;
      expect(oracle.draws(list, 400, 'normal', text)).toBe(!text.includes('ж'));
      return calls.length - before;
    };
    // A cold cluster costs the sentinel-only control pair and the route pair.
    expect(probes('г')).toBe(4);
    expect(probes('г')).toBe(0);
    // Another route reuses the control result: only its own pair.
    expect(probes('г', '"Covering", "Other"')).toBe(2);
    // A cluster longer than the cached bound is proved again each time, in
    // full: neither its result nor its control is retained.
    const long = `a${'\u0301'.repeat(40)}`;
    expect(probes(long)).toBe(4);
    expect(probes(long)).toBe(4);
    expect(probes(`${long}г`)).toBe(4);
    // A rejected cluster ends the request: later clusters are not probed.
    expect(probes('жд')).toBe(4);
    expect(probes('жд')).toBe(0);
    expect(probes('д')).toBe(4);
  });

  it('evicts instead of growing over high-cardinality text', () => {
    const { factory, calls, owner } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    // The capability is judged on first use (six sentinel measurements).
    expect(oracle.available).toBe(true);
    const probes = (text: string, list = '"Covering"') => {
      const before = calls.length;
      expect(oracle.draws(list, 400, 'normal', text)).toBe(true);
      return calls.length - before;
    };
    const ideograph = (index: number) => String.fromCodePoint(0x4e00 + index);
    // Twenty thousand distinct ideographs: each is judged exactly once.
    let cold = 0;
    for (let index = 0; index < 20000; index += 1) cold += probes(ideograph(index));
    expect(cold).toBe(20000 * 4);
    // The newest results are still held; the oldest were evicted (8192
    // results, 4096 controls) and are proved again.
    expect(probes(ideograph(19999))).toBe(0);
    expect(probes(ideograph(20000 - 8192))).toBe(0);
    expect(probes(ideograph(20000 - 8193))).toBe(4);
    expect(probes(ideograph(0))).toBe(4);
    // Evicted result, control still held: only the route pair.
    expect(probes(ideograph(20000 - 4000), '"Covering", "Second"')).toBe(2);
  });

  it('never interns a family list larger than the whole list budget', () => {
    const { factory, calls, owner } = fakeOwner();
    const oracle = createNamedFamilyOracle(owner, factory);
    // The capability is judged on first use (six sentinel measurements).
    expect(oracle.available).toBe(true);
    const probes = (list: string, text: string) => {
      const before = calls.length;
      expect(oracle.draws(list, 400, 'normal', text)).toBe(true);
      return calls.length - before;
    };
    expect(probes('"Covering"', 'г')).toBe(4);
    const oversized = `"Covering", ${Array.from({ length: 9000 }, (_, index) => `"Family ${index}"`).join(', ')}`;
    expect(oversized.length).toBeGreaterThan(65536);
    // Judged (the control is shared), never remembered, and the lists and
    // results already interned survive it.
    expect(probes(oversized, 'г')).toBe(2);
    expect(probes(oversized, 'г')).toBe(2);
    expect(probes('"Covering"', 'г')).toBe(0);
    // Lists that fit are interned until their total reaches the budget; the
    // table then starts over (results with it), never exceeding it.
    const list = (index: number) => `"Covering", "${`Family ${index} `.repeat(400)}"`;
    expect(probes(list(0), 'г')).toBe(2);
    expect(probes(list(0), 'г')).toBe(0);
    for (let index = 1; index < 24; index += 1) probes(list(index), 'г');
    expect(probes(list(23), 'г')).toBe(0);
    expect(probes(list(0), 'г')).toBe(2);
  });
});
