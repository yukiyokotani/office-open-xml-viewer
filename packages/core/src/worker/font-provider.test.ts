import { afterEach, describe, expect, it, vi } from 'vitest';
import { FontProvider, FontProviderSession, type FontAsset } from '../fonts/provider.js';
import { _resetFontRegistryForTests } from '../fonts/font-registry.js';
import {
  FONT_PROVIDER_PROTOCOL,
  FontProviderClient,
  FontProviderHost,
  isWorkerFontRequest,
} from './font-provider.js';

const G = globalThis as unknown as Record<string, unknown>;
const ORIGINALS = { document: G.document, FontFace: G.FontFace };

afterEach(() => {
  for (const [key, value] of Object.entries(ORIGINALS)) {
    if (value === undefined) delete G[key];
    else G[key] = value;
  }
  _resetFontRegistryForTests();
});

describe('font provider worker protocol', () => {
  it('resolves a worker request through the main-thread provider', async () => {
    class Provider extends FontProvider {
      async resolve(families: readonly string[]): Promise<readonly FontAsset[]> {
        return families.map((family) => ({
          family,
          source: { data: new Uint8Array([1, 2, 3]).buffer },
        }));
      }
    }
    const post = vi.fn();
    const host = new FontProviderHost(new FontProviderSession(new Provider()), post);
    const request = {
      protocol: FONT_PROVIDER_PROTOCOL,
      kind: 'resolve' as const,
      fontRequestId: 4,
      generation: 8,
      families: ['Calibri'],
    };

    expect(isWorkerFontRequest(request)).toBe(true);
    await expect(host.accept(request)).resolves.toBe(true);
    expect(post).toHaveBeenCalledOnce();
    expect(post.mock.calls[0][0]).toMatchObject({
      protocol: FONT_PROVIDER_PROTOCOL,
      kind: 'resolved',
      fontRequestId: 4,
      generation: 8,
      resolved: { routes: { calibri: expect.stringContaining('__ooxml_provider_') } },
    });
  });

  it('rejects strict requests when the worker cannot load a resolved face', async () => {
    class BrokenFace {
      constructor(
        readonly family: string,
        readonly source: string | ArrayBuffer,
      ) {}
      async load(): Promise<this> { throw new Error('invalid font'); }
    }
    G.document = { fonts: { add() {}, delete() { return true; } } };
    G.FontFace = BrokenFace;
    const post = vi.fn();
    const client = new FontProviderClient(post);
    const pending = client.resolve(['Calibri'], 3);
    await Promise.resolve();
    const request = post.mock.calls[0][0] as { fontRequestId: number };

    await client.accept({
      protocol: FONT_PROVIDER_PROTOCOL,
      kind: 'resolved',
      fontRequestId: request.fontRequestId,
      generation: 3,
      strict: true,
      resolved: {
        routes: { calibri: '__private_calibri' },
        faces: [{
          family: 'Calibri',
          alias: '__private_calibri',
          data: new Uint8Array([1]).buffer,
          descriptors: {},
        }],
      },
    });

    await expect(pending).rejects.toThrow('failed to load');
  });

  it('does not ask the host for the same family twice', async () => {
    const post = vi.fn();
    const client = new FontProviderClient(post);
    const first = client.resolve(['Calibri'], 1);
    await Promise.resolve();
    const request = post.mock.calls[0][0] as { fontRequestId: number };
    await client.accept({
      protocol: FONT_PROVIDER_PROTOCOL,
      kind: 'resolved',
      fontRequestId: request.fontRequestId,
      generation: 1,
      resolved: { routes: {}, faces: [] },
    });
    await first;

    await expect(client.resolve(['calibri'], 1)).resolves.toEqual({});
    expect(post).toHaveBeenCalledOnce();
  });
});
