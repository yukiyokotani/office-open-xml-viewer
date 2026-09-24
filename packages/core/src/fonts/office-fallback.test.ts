import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { _resetFontRegistryForTests } from './font-registry.js';
import { loadOfficeFontFallbacks, unloadOfficeFontFallbacks } from './office-fallback.js';
import referenceData from './reference-font-metrics-data.json';

const globals = globalThis as unknown as Record<string, unknown>;
const originals = Object.fromEntries(
  ['FontFace', 'fetch', 'document', 'self'].map((name) => [name, globals[name]]),
);

type Face = { family: string; source: string | ArrayBuffer; status: FontFaceLoadStatus };

function fontSet(installed: readonly string[], delayMs = 0, declared: readonly string[] = []) {
  const added: Face[] = [];
  const deleted: Face[] = [];
  let active = 0;
  let peak = 0;
  const set = {
    add(face: Face) { added.push(face); },
    delete(face: Face) { deleted.push(face); return true; },
    *[Symbol.iterator]() {
      for (const family of declared) yield { family, source: '', status: 'loaded' };
      for (const face of added) if (!deleted.includes(face)) yield face;
    },
  } as unknown as FontFaceSet;
  class FakeFace implements Face {
    status: FontFaceLoadStatus = 'unloaded';
    constructor(readonly family: string, readonly source: string | ArrayBuffer) {}
    async load(): Promise<this> {
      active++;
      peak = Math.max(peak, active);
      if (delayMs) await new Promise((resolve) => setTimeout(resolve, delayMs));
      active--;
      const source = this.source;
      if (typeof source !== 'string' ||
          !installed.some((name) => source.includes(`local("${name}")`))) {
        this.status = 'error';
        throw new Error('missing local face');
      }
      this.status = 'loaded';
      return this;
    }
  }
  globals.FontFace = FakeFace;
  return { set, added, deleted, peak: () => peak };
}

beforeEach(() => {
  _resetFontRegistryForTests();
});

afterEach(() => {
  vi.useRealTimers();
  for (const [name, value] of Object.entries(originals)) {
    if (value === undefined) delete globals[name];
    else globals[name] = value;
  }
  _resetFontRegistryForTests();
});

describe('loadOfficeFontFallbacks', () => {
  it('leaves an unavailable Calibri face unresolved without a network request', async () => {
    const { set, added, deleted } = fontSet([]);
    let fetchCount = 0;
    globals.fetch = async () => { fetchCount++; throw new Error('network must remain unused'); };

    const loaded = await loadOfficeFontFallbacks([{ family: 'Calibri', weight: 700 }], set);

    expect(loaded).toEqual({ faces: [], routes: {} });
    expect(fetchCount).toBe(0);
    expect(added).toHaveLength(1);
    expect(deleted).toEqual(added);
  });

  it('uses an exact local styled face without borrowing a missing style', async () => {
    const { set, deleted } = fontSet(['Calibri-Bold']);
    globals.fetch = async () => { throw new Error('unexpected fetch'); };

    const result = await loadOfficeFontFallbacks([
      { family: 'Calibri', weight: 700 },
      { family: 'Calibri', weight: 700 },
      { family: 'Calibri' },
      { family: 'Calibri Light' },
    ], set);

    expect(result.faces).toHaveLength(1);
    expect(result.routes.calibri).toBeUndefined();
    expect(result.routes['calibri:700:normal']).toMatchObject({
      source: 'local', weight: 700, style: 'normal',
      resourceIdentity: 'office-local:local("Calibri-Bold")',
      metric: { sourceIdentity: 'office-local:local("Calibri-Bold")' },
    });
    expect(result.routes['calibri:700:normal'].metric.lineHeightRatio).toBeUndefined();
    unloadOfficeFontFallbacks(result.faces);
    expect(deleted).toContain(result.faces[0]);
  });

  it('retains one local face across concurrent holders in the same font set', async () => {
    const { set, added, deleted } = fontSet(['Calibri']);
    const [first, second] = await Promise.all([
      loadOfficeFontFallbacks([{ family: 'Calibri' }], set),
      loadOfficeFontFallbacks([{ family: 'Calibri' }], set),
    ]);
    expect(added).toHaveLength(1);
    expect(first.faces[0]).toBe(second.faces[0]);
    unloadOfficeFontFallbacks(first.faces);
    expect(deleted).toHaveLength(0);
    unloadOfficeFontFallbacks(second.faces);
    expect(deleted).toEqual(added);
  });

  it('registers an exact tuple independently in distinct main and worker font sets', async () => {
    const firstSet = fontSet(['Calibri']);
    const secondSet = fontSet(['Calibri']);

    const [main, worker] = await Promise.all([
      loadOfficeFontFallbacks([{ family: 'Calibri' }], firstSet.set),
      loadOfficeFontFallbacks([{ family: 'Calibri' }], secondSet.set),
    ]);

    expect(main.faces).toHaveLength(1);
    expect(worker.faces).toHaveLength(1);
    expect(main.faces[0]).not.toBe(worker.faces[0]);
    unloadOfficeFontFallbacks(main.faces);
    expect(firstSet.deleted).toContain(main.faces[0]);
    expect(secondSet.deleted).not.toContain(worker.faces[0]);
    unloadOfficeFontFallbacks(worker.faces);
  });

  it('resolves a catalogued authored family through a present exact face and isolates its style', async () => {
    const { set } = fontSet(['TimesNewRomanPSMT', 'TimesNewRomanPS-BoldMT']);
    const result = await loadOfficeFontFallbacks([
      { family: 'Times New Roman' },
      { family: 'Times New Roman', weight: 700 },
      { family: 'Times New Roman', style: 'italic' },
    ], set);
    expect(Object.keys(result.routes).sort()).toEqual([
      'times new roman', 'times new roman:700:normal',
    ]);
    expect(result.routes['times new roman'].resourceIdentity).toContain('TimesNewRomanPSMT');
    expect(result.routes['times new roman:700:normal'].resourceIdentity).toContain('TimesNewRomanPS-BoldMT');
    unloadOfficeFontFallbacks(result.faces);
  });

  it('resolves a translated authored family by its catalogued local name', async () => {
    const { set } = fontSet(['Meiryo']);
    const result = await loadOfficeFontFallbacks([
      { family: 'Meiryo' }, { family: 'メイリオ' },
    ], set);
    expect(result.faces).toHaveLength(1);
    expect(result.routes.meiryo).toBeDefined();
    expect(result.routes['メイリオ']).toMatchObject({
      requestedFamily: 'メイリオ', source: 'local', weight: 400,
    });
    expect(result.routes['メイリオ'].resourceIdentity).toContain('local("Meiryo")');
    unloadOfficeFontFallbacks(result.faces);
  });

  it('bounds independent local probes and cleans up failed faces', async () => {
    const { set, added, deleted, peak } = fontSet([], 3);
    const result = await loadOfficeFontFallbacks([
      'Calibri', 'Times New Roman', 'Meiryo', 'Arial', 'Courier New', 'Georgia', 'Cambria',
    ].map((family) => ({ family })), set);
    expect(result.routes).toEqual({});
    expect(added.length).toBeGreaterThan(4);
    expect(peak()).toBeGreaterThan(1);
    expect(peak()).toBeLessThanOrEqual(4);
    expect(deleted).toEqual(added);
  });

  it('leaves an application-declared authored family to normal CSS resolution', async () => {
    const { set, added } = fontSet(['TimesNewRomanPSMT'], 0, ['"Times New Roman"']);
    const result = await loadOfficeFontFallbacks([{ family: 'Times New Roman' }], set);
    expect(result).toEqual({ faces: [], routes: {} });
    expect(added).toEqual([]);
  });

  it('caps the number of local source probes for a font-heavy document', async () => {
    const { set, added, deleted } = fontSet([]);
    const families = [...new Set(referenceData.profiles
      .filter((profile) => profile.weight === 400 && profile.style === 'normal')
      .map((profile) => profile.family))].slice(0, 80);
    await loadOfficeFontFallbacks(families.map((family) => ({ family })), set);
    expect(families.length).toBe(80);
    expect(added.length).toBeGreaterThan(0);
    expect(added.length).toBeLessThanOrEqual(32);
    expect(deleted).toEqual(added);
  });

  it('returns at the document deadline and releases a face that finishes later', async () => {
    vi.useFakeTimers();
    const added: Face[] = [];
    const deleted: Face[] = [];
    let finish!: () => void;
    class DeferredFace implements Face {
      status: FontFaceLoadStatus = 'unloaded';
      constructor(readonly family: string, readonly source: string | ArrayBuffer) {}
      load(): Promise<this> {
        return new Promise((resolve) => {
          finish = () => { this.status = 'loaded'; resolve(this); };
        });
      }
    }
    globals.FontFace = DeferredFace;
    const set = {
      add(face: Face) { added.push(face); },
      delete(face: Face) { deleted.push(face); return true; },
    } as unknown as FontFaceSet;
    const pending = loadOfficeFontFallbacks([{ family: 'Calibri' }], set);
    await vi.advanceTimersByTimeAsync(8_000);
    expect(await pending).toEqual({ faces: [], routes: {} });
    expect(added).toHaveLength(1);
    finish();
    await vi.advanceTimersByTimeAsync(0);
    expect(deleted).toEqual(added);
  });
});
