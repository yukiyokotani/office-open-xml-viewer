import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  FontProvider,
  FontProviderSession,
  providerFontFamily,
  type FontAsset,
} from './provider.js';
import { _resetFontRegistryForTests } from './font-registry.js';

const G = globalThis as unknown as Record<string, unknown>;
const ORIGINALS = { document: G.document, FontFace: G.FontFace, fetch: G.fetch };

class TestProvider extends FontProvider {
  readonly resolve = vi.fn(async (families: readonly string[]): Promise<readonly FontAsset[]> =>
    families.map((family) => ({
      family,
      source: { data: new Uint8Array([1, 2, 3, family.length]).buffer },
      descriptors: { weight: '400', style: 'normal' },
    })),
  );
}

function installFonts(fail = false, waitForLoad?: () => Promise<void>) {
  const added: FakeFace[] = [];
  const deleted: FakeFace[] = [];
  class FakeFace {
    status: FontFaceLoadStatus = 'unloaded';
    constructor(
      readonly family: string,
      readonly source: string | ArrayBuffer,
      readonly descriptors: FontFaceDescriptors = {},
    ) {}
    async load(): Promise<this> {
      if (fail) throw new Error('invalid font');
      await waitForLoad?.();
      this.status = 'loaded';
      return this;
    }
  }
  const fonts = {
    add(face: FakeFace) { added.push(face); },
    delete(face: FakeFace) { deleted.push(face); return true; },
  } as unknown as FontFaceSet;
  G.document = { fonts };
  G.FontFace = FakeFace;
  return { fonts, added, deleted };
}

beforeEach(() => _resetFontRegistryForTests());
afterEach(() => {
  for (const [key, value] of Object.entries(ORIGINALS)) {
    if (value === undefined) delete G[key];
    else G[key] = value;
  }
  _resetFontRegistryForTests();
  vi.restoreAllMocks();
});

describe('FontProviderSession', () => {
  it('rejects an invalid failure policy', () => {
    expect(() => new FontProviderSession(new TestProvider(), 'invalid' as never))
      .toThrow('invalid fontFailure');
  });

  it('registers an isolated fallback and keeps the authored family first', async () => {
    const { fonts, added } = installFonts();
    const provider = new TestProvider();
    const session = new FontProviderSession(provider);

    const resolved = await session.ensure(['Calibri'], fonts);

    expect(provider.resolve).toHaveBeenCalledWith(['Calibri'], expect.anything());
    expect(added).toHaveLength(1);
    expect(added[0].family).not.toBe('Calibri');
    expect(providerFontFamily(resolved.routes, 'Calibri')).toBe(added[0].family);
  });

  it('deduplicates family resolution and registration across incremental requests', async () => {
    const { fonts, added } = installFonts();
    const provider = new TestProvider();
    const session = new FontProviderSession(provider);

    await session.ensure(['Calibri'], fonts);
    await session.ensure(['calibri', 'Cambria'], fonts);

    expect(provider.resolve.mock.calls).toEqual([
      [['Calibri'], expect.anything()],
      [['Cambria'], expect.anything()],
    ]);
    expect(added).toHaveLength(2);
  });

  it('deduplicates concurrent registration and releases its only hold', async () => {
    let loadCount = 0;
    let finishLoad!: () => void;
    const loadGate = new Promise<void>((resolve) => { finishLoad = resolve; });
    const { fonts, added, deleted } = installFonts(
      false,
      async () => {
        loadCount += 1;
        await loadGate;
      },
    );
    const session = new FontProviderSession(new TestProvider());

    const first = session.ensure(['Calibri'], fonts);
    await vi.waitFor(() => expect(loadCount).toBe(1));
    const second = session.ensure(['calibri'], fonts);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const concurrentLoads = loadCount;
    finishLoad();
    await Promise.all([first, second]);
    session.destroy();

    expect(concurrentLoads).toBe(1);
    expect(added).toHaveLength(1);
    expect(deleted).toEqual(added);
  });

  it('returns independent worker buffers and releases registered faces', async () => {
    const { fonts, added, deleted } = installFonts();
    const session = new FontProviderSession(new TestProvider());
    const resolved = await session.ensure(['Calibri'], fonts);

    expect(resolved.faces).toHaveLength(1);
    expect(resolved.faces[0].data).not.toBe(added[0].source);
    new Uint8Array(resolved.faces[0].data)[0] = 9;
    expect(new Uint8Array(added[0].source as ArrayBuffer)[0]).toBe(1);

    session.destroy();
    expect(deleted).toEqual(added);
  });

  it('releases a viewer FontFaceSet after its last holder', async () => {
    const { fonts, added, deleted } = installFonts();
    const session = new FontProviderSession(new TestProvider());
    await session.ensure(['Calibri'], fonts);
    const releaseFirst = session.retain(fonts);
    const releaseSecond = session.retain(fonts);

    releaseFirst();
    expect(deleted).toEqual([]);
    releaseSecond();
    expect(deleted).toEqual(added);
  });

  it('supports fallback and strict missing-family policies', async () => {
    installFonts();
    class EmptyProvider extends FontProvider {
      async resolve(): Promise<readonly FontAsset[]> { return []; }
    }
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);

    await expect(new FontProviderSession(new EmptyProvider()).ensure(['Missing']))
      .resolves.toMatchObject({ faces: [], routes: {} });
    expect(warn).toHaveBeenCalledOnce();

    await expect(new FontProviderSession(new EmptyProvider(), 'error').ensure(['Missing']))
      .rejects.toThrow('Missing');
  });

  it('does not route main-thread layout to a face that failed to load', async () => {
    const { fonts, added, deleted } = installFonts(true);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);

    const resolved = await new FontProviderSession(new TestProvider()).ensure(['Calibri'], fonts);

    expect(resolved.routes).toEqual({});
    expect(added).toHaveLength(1);
    expect(deleted).toEqual(added);
    expect(warn).toHaveBeenCalledOnce();
  });
});
