import { vi, afterEach, beforeEach, describe, expect, it } from 'vitest';
let { loadLocalFontMetrics, unloadLocalFontMetrics } = await import('./local-metrics.js');

beforeEach(async () => {
  vi.resetModules();
  ({ loadLocalFontMetrics, unloadLocalFontMetrics } = await import('./local-metrics.js'));
});

const G = globalThis as unknown as Record<string, unknown>;
const ORIGINALS = {
  document: G.document,
  self: G.self,
  FontFace: G.FontFace,
  OffscreenCanvas: G.OffscreenCanvas,
};


afterEach(() => {
  for (const [key, value] of Object.entries(ORIGINALS)) {
    if (value === undefined) delete G[key];
    else G[key] = value;
  }
});

function installFontEnvironment(widthOf: (text: string) => number = () => 50) {
  const added: FakeFontFace[] = [];
  const deleted: FakeFontFace[] = [];
  class FakeFontFace {
    status: FontFaceLoadStatus = 'unloaded';
    constructor(
      readonly family: string,
      readonly source: string,
      readonly descriptors: FontFaceDescriptors = {},
    ) {}
    async load(): Promise<this> {
      if (this.source.includes('Missing Face')) {
        this.status = 'error';
        throw new DOMException('missing', 'NetworkError');
      }
      this.status = 'loaded';
      return this;
    }
  }
  const fonts = {
    add(face: FakeFontFace) { added.push(face); },
    delete(face: FakeFontFace) { deleted.push(face); return true; },
  } as unknown as FontFaceSet;
  class FakeOffscreenCanvas {
    constructor(_width: number, _height: number) {}
    getContext() {
      return {
        font: '',
        measureText: (text: string) => ({
          width: widthOf(text),
          fontBoundingBoxAscent: 106,
          fontBoundingBoxDescent: 44,
        }),
      };
    }
  }
  G.document = { fonts };
  G.FontFace = FakeFontFace;
  G.OffscreenCanvas = FakeOffscreenCanvas;
  return { added, deleted };
}

describe('loadLocalFontMetrics', () => {
  it('loads an exact local face under an isolated alias and derives its line ratio', async () => {
    const { added, deleted } = installFontEnvironment();

    const loaded = await loadLocalFontMetrics([{
      family: 'メイリオ',
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    }]);

    expect(loaded.faces).toHaveLength(1);
    expect(added).toHaveLength(1);
    expect(added[0].source).toBe('local("Meiryo")');
    expect(loaded.metrics['メイリオ'].family).toBe(added[0].family);
    expect(loaded.metrics['メイリオ'].lineHeightRatio).toBeCloseTo(1.95, 8);

    unloadLocalFontMetrics(loaded.faces);
    expect(deleted).toEqual([added[0]]);
  });

  it('keeps a family named like an Object.prototype member as an own entry', async () => {
    installFontEnvironment();
    const loaded = await loadLocalFontMetrics([
      { family: '__proto__', localNames: ['__proto__'] },
      { family: 'constructor', localNames: ['constructor'] },
    ]);
    expect(Object.keys(loaded.metrics).sort()).toEqual(['__proto__', 'constructor']);
    expect(Object.prototype.hasOwnProperty.call(loaded.metrics, '__proto__')).toBe(true);
    unloadLocalFontMetrics(loaded.faces);
  });

  it('does not mistake a missing local font for a fallback face', async () => {
    const { added, deleted } = installFontEnvironment();

    const loaded = await loadLocalFontMetrics([{
      family: 'Missing',
      localNames: ['Missing Face'],
      lineHeightMultiplier: 1.3,
    }]);

    expect(loaded.faces).toEqual([]);
    expect(loaded.metrics).toEqual({});
    expect(added).toHaveLength(1);
    expect(deleted).toEqual(added);
  });

  it('loads exact regular and bold face names as separate declared capabilities', async () => {
    const { added } = installFontEnvironment();
    const loaded = await loadLocalFontMetrics([
      { family: 'Authored Serif', localNames: ['Authored Serif Regular'] },
      { family: 'Authored Serif', localNames: ['Authored Serif Bold'], weight: 700 },
    ]);

    expect(added).toHaveLength(2);
    expect(added.map((face) => face.source)).toEqual([
      'local("Authored Serif Regular")', 'local("Authored Serif Bold")',
    ]);
    expect(added.map((face) => face.descriptors)).toEqual([
      { weight: '400', style: 'normal' }, { weight: '700', style: 'normal' },
    ]);
    expect(loaded.metrics['authored serif'].family)
      .not.toBe(loaded.metrics['authored serif:700:normal'].family);
    expect(loaded.metrics['authored serif:700:normal']).toMatchObject({
      requestedFamily: 'Authored Serif', weight: 700, style: 'normal',
      sourceIdentity: 'local("Authored Serif Bold")',
    });
    expect(loaded.metrics['authored serif:700:normal'].lineHeightRatio).toBeUndefined();
  });

  it('does not sample Canvas geometry into exact-source identity', async () => {
    const request = [{ family: 'Authored Serif', localNames: ['Authored Serif Regular'] }];
    installFontEnvironment((text) => text === 'Q' ? 51 : 50);
    const first = await loadLocalFontMetrics(request);
    unloadLocalFontMetrics(first.faces);

    installFontEnvironment((text) => text === 'Q' ? 52 : 50);
    const second = await loadLocalFontMetrics(request);

    expect(first.metrics).toEqual(second.metrics);
    expect(first.metrics['authored serif']).not.toHaveProperty('geometrySignature');
  });

  it('does not invent styled capabilities from an exact regular face', async () => {
    installFontEnvironment();
    const loaded = await loadLocalFontMetrics([
      { family: 'Authored Serif', localNames: ['Authored Serif Regular'] },
    ]);

    expect(loaded.metrics).toHaveProperty('authored serif');
    expect(loaded.metrics).not.toHaveProperty('authored serif:700:normal');
    expect(loaded.metrics).not.toHaveProperty('authored serif:400:italic');
  });

  it('waits for a shared face that another document is still loading', async () => {
    const added: Array<{ family: string; source: string; status: FontFaceLoadStatus }> = [];
    let finishLoad!: () => void;
    const loaded = new Promise<void>((resolve) => { finishLoad = resolve; });
    class DeferredFontFace {
      status: FontFaceLoadStatus = 'unloaded';
      constructor(readonly family: string, readonly source: string) {}
      async load(): Promise<this> {
        if (this.status === 'loaded') return this;
        this.status = 'loading';
        await loaded;
        this.status = 'loaded';
        return this;
      }
    }
    const fonts = {
      add(face: DeferredFontFace) { added.push(face); },
      delete() { return true; },
    } as unknown as FontFaceSet;
    class DeferredCanvas {
      getContext() {
        return {
          font: '',
          measureText: () => added[0]?.status === 'loaded'
            ? { width: 50, fontBoundingBoxAscent: 106, fontBoundingBoxDescent: 44 }
            : { width: Number.NaN, fontBoundingBoxAscent: Number.NaN, fontBoundingBoxDescent: Number.NaN },
        };
      }
    }
    G.document = { fonts };
    G.FontFace = DeferredFontFace;
    G.OffscreenCanvas = DeferredCanvas;
    const request = [{
      family: 'Meiryo',
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    }];

    const first = loadLocalFontMetrics(request);
    await Promise.resolve();
    const second = loadLocalFontMetrics(request);
    await Promise.resolve();
    finishLoad();

    const [a, b] = await Promise.all([first, second]);
    expect(added).toHaveLength(1);
    expect(a.metrics.meiryo.lineHeightRatio).toBeCloseTo(1.95, 8);
    expect(b.metrics.meiryo.lineHeightRatio).toBeCloseTo(1.95, 8);
  });

  it('is a no-op when the current context has no FontFaceSet', async () => {
    delete G.document;
    delete G.self;
    const loaded = await loadLocalFontMetrics([{
      family: 'Meiryo',
      localNames: ['Meiryo'],
      lineHeightMultiplier: 1.3,
    }]);
    expect(loaded).toEqual({ faces: [], metrics: {} });
  });
});
