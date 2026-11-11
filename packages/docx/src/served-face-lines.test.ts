import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { LoadedFontSources } from '@silurus/ooxml-core/internal/font-sources';

// Route sentinels need a real engine; the probe wait is what is tested here.
vi.mock('@silurus/ooxml-core', async (original) => ({
  ...await original<typeof import('@silurus/ooxml-core')>(),
  loadNativeRouteSentinels: async () => [],
}));

const { prepareServedFaceLines, SERVED_FACE_PROBE_TIMEOUT_MS } = await import('./served-face-lines.js');

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
