import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { LoadedFontSources } from '@silurus/ooxml-core/internal/font-sources';

// Route sentinels need a real engine; the probe wait is what is tested here.
vi.mock('@silurus/ooxml-core', async (original) => ({
  ...await original<typeof import('@silurus/ooxml-core')>(),
  loadNativeRouteSentinels: async () => [],
}));

const { prepareServedFaceLines, servedFaceLineRequests, SERVED_FACE_PROBE_TIMEOUT_MS } = await import('./served-face-lines.js');
const { unloadOfficeFontFallbacks } = await import('@silurus/ooxml-core');

const G = globalThis as { FontFace?: unknown };
const ORIGINAL = G.FontFace;

/** A FontFace whose load never settles (a stalled second request). */
class StalledFontFace {
  status = 'unloaded';
  constructor(public family: string, public source: unknown) {}
  load(): Promise<this> { return new Promise(() => undefined); }
}

function fakeSet() {
  const faces: StalledFontFace[] = [];
  return {
    faces,
    set: {
      add: (face: StalledFontFace) => { faces.push(face); },
      delete: (face: StalledFontFace) => { const at = faces.indexOf(face); if (at >= 0) faces.splice(at, 1); return at >= 0; },
      [Symbol.iterator]: () => faces[Symbol.iterator](),
    } as unknown as FontFaceSet,
  };
}

const sources = {
  faces: [{
    family: 'Unreadable', weight: '400', style: 'normal', stretch: 'normal', unicodeRange: 'U+0-10FFFF',
    sourceIndex: 0, face: {} as FontFace, probeSource: new ArrayBuffer(8),
  }],
  substitutions: new Map(),
  regularAliases: new Map(),
  families: new Map([['unreadable', 'ooxml-font-source-test-0']]),
} as unknown as LoadedFontSources;

beforeEach(() => { G.FontFace = StalledFontFace; });
afterEach(() => { G.FontFace = ORIGINAL; vi.useRealTimers(); });

describe('served-face probe waits', () => {
  it('end when the document aborts, releasing the probe registration', async () => {
    const { faces, set } = fakeSet();
    const controller = new AbortController();
    const pending = prepareServedFaceLines(sources, set, [], [], controller.signal);
    await Promise.resolve();
    await Promise.resolve();
    expect(faces).toHaveLength(1);
    controller.abort();
    const prepared = await pending;
    expect(prepared.faces).toEqual([]);
    expect(prepared.faceLines.size).toBe(0);
    expect(faces).toHaveLength(0);
  });

  it('end at the probe ceiling when a load never settles', async () => {
    vi.useFakeTimers();
    const { faces, set } = fakeSet();
    const pending = prepareServedFaceLines(sources, set, [], []);
    await vi.advanceTimersByTimeAsync(SERVED_FACE_PROBE_TIMEOUT_MS);
    const prepared = await pending;
    expect(prepared.faces).toEqual([]);
    expect(faces).toHaveLength(0);
  });
});

/** A FontFace that loads at once (the bytes are already in memory). */
class LoadingFontFace {
  status = 'unloaded';
  constructor(public family: string, public source: unknown, public descriptors: FontFaceDescriptors) {}
  load(): Promise<this> { this.status = 'loaded'; return Promise.resolve(this); }
}

function servedSources(buffer: ArrayBuffer): LoadedFontSources {
  const face = (weight: string, unicodeRange: string) => ({
    family: 'Sliced', weight, style: 'normal', stretch: 'normal', unicodeRange,
    sourceIndex: 0, face: {} as FontFace, probeSource: buffer,
  });
  return {
    // Two declarations of one file with equal descriptors, and the same file
    // declared bold (a variable font served for two weights).
    faces: [face('400', 'U+0-FF'), face('400', 'U+100-24F'), face('700', 'U+0-FF')],
    substitutions: new Map(),
    regularAliases: new Map(),
    families: new Map([['sliced', 'ooxml-font-source-test-1']]),
  } as unknown as LoadedFontSources;
}

const LINES = { '400:normal': { ascentRatio: 0.8, descentRatio: 0.2, lineGapRatio: 0.3 } };

describe('served-face probe registrations', () => {
  it('register one probe per source and descriptors, shared by the documents of a realm', async () => {
    G.FontFace = LoadingFontFace;
    const { faces, set } = fakeSet();
    const buffer = new ArrayBuffer(8);
    const measured = [0, 1, 2].map((index) => ({ index, lines: LINES }));
    const first = await prepareServedFaceLines(servedSources(buffer), set, [], measured);
    const second = await prepareServedFaceLines(servedSources(buffer), set, [], measured);
    // Regular and bold of the one file: two probes for six snapshot faces.
    expect(faces).toHaveLength(2);
    expect(new Set((faces as unknown as LoadingFontFace[]).map((face) => face.descriptors.weight))).toEqual(new Set(['400', '700']));
    for (const prepared of [first, second]) {
      expect(prepared.faceLines.get(0)?.ownershipFamily).toBe(prepared.faceLines.get(1)?.ownershipFamily);
      expect(prepared.faceLines.get(0)?.ownershipFamily).not.toBe(prepared.faceLines.get(2)?.ownershipFamily);
      expect(prepared.faceLines.get(2)?.lines).toEqual(LINES);
    }
    expect(second.faceLines.get(2)?.ownershipFamily).toBe(first.faceLines.get(2)?.ownershipFamily);
    // Releasing one document keeps the other's probes; the last release
    // removes them.
    unloadOfficeFontFallbacks(first.faces);
    expect(faces).toHaveLength(2);
    unloadOfficeFontFallbacks(second.faces);
    expect(faces).toHaveLength(0);
  });

  it('register the layout realm probes while the DOM realm measures', async () => {
    G.FontFace = LoadingFontFace;
    const { faces, set } = fakeSet();
    let answer: (measurements: { index: number; lines: typeof LINES }[]) => void = () => undefined;
    const measured = new Promise<{ index: number; lines: typeof LINES }[]>((resolve) => { answer = resolve; });
    const pending = prepareServedFaceLines(servedSources(new ArrayBuffer(8)), set, [], measured);
    for (let turn = 0; turn < 5; turn++) await Promise.resolve();
    expect(faces).toHaveLength(2);
    answer([0, 1, 2].map((index) => ({ index, lines: LINES })));
    const prepared = await pending;
    expect([...prepared.faceLines.keys()]).toEqual([0, 1, 2]);
    unloadOfficeFontFallbacks(prepared.faces);
    expect(faces).toHaveLength(0);
  });
});

describe('subsets of a shared-metric style', () => {
  function slicedSources(): LoadedFontSources {
    const face = (family: string, weight: string, unicodeRange: string, shared: boolean) => ({
      family, weight, style: 'normal', stretch: 'normal', unicodeRange,
      sourceIndex: 0, face: {} as FontFace, probeSource: new ArrayBuffer(8),
      ...(shared ? { sharedSubsetMetrics: true } : {}),
    });
    return {
      faces: [
        face('Sliced', '400', 'U+4E00-9FFF', true),
        face('Sliced', '400', 'U+0-FF', true),
        face('Sliced', '400', 'U+3000-30FF', true),
        face('Sliced', '700', 'U+0-FF', true),
        // A style none of whose loaded subsets holds U+0020: per face.
        face('Sliced', '400', 'U+AC00-D7AF', false),
        face('Kanji Only', '400', 'U+4E00-9FFF', true),
        face('Kanji Only', '400', 'U+3000-30FF', true),
      ],
      substitutions: new Map(),
      regularAliases: new Map(),
      families: new Map([['sliced', 'ooxml-font-source-test-2'], ['kanji only', 'ooxml-font-source-test-3']]),
    } as unknown as LoadedFontSources;
  }

  it('are measured once, through the subset holding U+0020, and judged on the painting family', async () => {
    const sources = slicedSources();
    const requests = servedFaceLineRequests(sources);
    expect(requests.map((request) => [request.index, request.members])).toEqual([
      [1, [0, 1, 2]], [3, [3]], [4, undefined], [5, undefined], [6, undefined],
    ]);
    G.FontFace = LoadingFontFace;
    const { faces, set } = fakeSet();
    const measured = requests.map((request) => ({ index: request.index, lines: LINES }));
    const prepared = await prepareServedFaceLines(sources, set, [], measured);
    // The layout realm registers probes only for faces judged on their own.
    expect(faces).toHaveLength(3);
    for (const index of [0, 1, 2]) {
      expect(prepared.faceLines.get(index)).toEqual({ ownershipFamily: 'ooxml-font-source-test-2', lines: LINES });
    }
    expect(prepared.faceLines.get(3)?.ownershipFamily).toBe('ooxml-font-source-test-2');
    for (const index of [4, 5, 6]) expect(prepared.faceLines.get(index)?.ownershipFamily).toMatch(/^__ooxml_served_face_/);
    unloadOfficeFontFallbacks(prepared.faces);
    expect(faces).toHaveLength(0);
  });
});
