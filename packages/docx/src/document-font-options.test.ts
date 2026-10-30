import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkerLike } from '@silurus/ooxml-core';
import {
  USE_GOOGLE_FONTS_DEPRECATION_MESSAGE,
  googleFonts,
  resetFontDeprecationWarningsForTests,
} from '@silurus/ooxml-core/internal/font-sources';
import { DocxDocument } from './document';

/** `DocxDocument.load` normalises the deprecated `useGoogleFonts` before it
 * reads any option (design §9.1), and a worker load posts the normalised
 * descriptors, never the key. `_parse` is stubbed to observe what crosses the
 * worker boundary and to end the load there. */

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void {}
}

const G = globalThis as Record<string, unknown>;
const ORIGINAL = { Worker: G.Worker, OffscreenCanvas: G.OffscreenCanvas, location: G.location };
afterEach(() => {
  G.Worker = ORIGINAL.Worker;
  G.OffscreenCanvas = ORIGINAL.OffscreenCanvas;
  G.location = ORIGINAL.location;
  resetFontDeprecationWarningsForTests();
  vi.restoreAllMocks();
});

const stop = new Error('parse stub');
function stubParse() {
  G.Worker = SilentWorker;
  G.OffscreenCanvas = class {};
  G.location = { href: 'http://localhost/' };
  return vi.spyOn(
    DocxDocument.prototype as unknown as { _parse(...args: unknown[]): Promise<void> },
    '_parse',
  ).mockRejectedValue(stop);
}

describe('DocxDocument.load font options', () => {
  it('posts useGoogleFonts: true to the worker as fonts: [googleFonts()] and warns once', async () => {
    const parse = stubParse();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    await expect(DocxDocument.load(new ArrayBuffer(0), { mode: 'worker', useGoogleFonts: true })).rejects.toBe(stop);
    await expect(DocxDocument.load(new ArrayBuffer(0), { mode: 'worker', useGoogleFonts: true })).rejects.toBe(stop);
    expect(parse.mock.calls[0][2]).toEqual([googleFonts()]);
    expect(warn.mock.calls.filter(([message]) => message === USE_GOOGLE_FONTS_DEPRECATION_MESSAGE)).toHaveLength(1);
  });

  it('posts application sources unchanged and nothing without sources', async () => {
    const parse = stubParse();
    const fonts = [googleFonts({ cssOrigin: 'https://fonts.internal.example:8443' })];
    await expect(DocxDocument.load(new ArrayBuffer(0), { mode: 'worker', fonts })).rejects.toBe(stop);
    await expect(DocxDocument.load(new ArrayBuffer(0), { mode: 'worker' })).rejects.toBe(stop);
    await expect(DocxDocument.load(new ArrayBuffer(0), { fonts })).rejects.toBe(stop);
    expect(parse.mock.calls.map((call) => call[2])).toEqual([fonts, undefined, undefined]);
  });

  it('rejects useGoogleFonts: true with fonts, and malformed sources, before any worker starts', async () => {
    const parse = stubParse();
    vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    await expect(DocxDocument.load(new ArrayBuffer(0), { useGoogleFonts: true, fonts: [] }))
      .rejects.toThrow(TypeError);
    await expect(DocxDocument.load(new ArrayBuffer(0), { fonts: [{ kind: 'font-css-api' }] as never }))
      .rejects.toThrow(TypeError);
    expect(parse).not.toHaveBeenCalled();
  });
});
