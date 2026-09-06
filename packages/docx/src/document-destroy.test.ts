import { describe, it, expect, afterEach, vi } from 'vitest';
import { BoundedRawPartCache } from '@silurus/ooxml-core/internal/bounded-raw-part-cache';
import {
  WorkerBridge,
  registerEmbeddedFonts,
  preloadGoogleFonts,
  type WorkerLike,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import { DocxDocument } from './document';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';

/**
 * `DocxDocument.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight —
 * otherwise a `load()` / image extraction awaiting the worker would hang
 * forever after the document is disposed. This pins that delegation using a
 * real {@link WorkerBridge} over an in-memory worker (no real Worker: the
 * constructor opens one, so we build the instance off-prototype and inject the
 * bridge — the established pattern from `document.image.test.ts`).
 */

/** In-memory Worker stand-in that never answers, so requests stay pending until
 *  the bridge is terminated. */
class SilentWorker implements WorkerLike {
  static instances: SilentWorker[] = [];

  constructor() {
    SilentWorker.instances.push(this);
  }

  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminated = false;
  terminate(): void {
    this.terminated = true;
  }
}

// ── Fake FontFaceSet so destroy()'s embedded-font / Google-Fonts release is
// observable ──────────────────────────────────────────────────────────────
const G = globalThis as Record<string, unknown>;
const ORIG_FONTS = {
  document: G.document,
  self: G.self,
  fetch: G.fetch,
  FontFace: G.FontFace,
  OffscreenCanvas: G.OffscreenCanvas,
  Worker: G.Worker,
  location: G.location,
};
afterEach(() => {
  G.document = ORIG_FONTS.document;
  G.self = ORIG_FONTS.self;
  G.fetch = ORIG_FONTS.fetch;
  G.FontFace = ORIG_FONTS.FontFace;
  G.OffscreenCanvas = ORIG_FONTS.OffscreenCanvas;
  G.Worker = ORIG_FONTS.Worker;
  G.location = ORIG_FONTS.location;
  SilentWorker.instances = [];
  vi.restoreAllMocks();
});

interface FakeFace { family: string }
function installFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: ArrayBuffer, public descriptors?: object) {}
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  const set = {
    faces: added,
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  delete G.self;
  return { added };
}

// ── Google-Fonts flavored fake: `preloadGoogleFonts` needs `fetch` (to pull
// the CSS) and a string-`src` `FontFace` constructor, unlike the ArrayBuffer
// source used by the embedded-font fake above. Mirrors the fake used by
// `presentation-destroy.test.ts` / `workbook-destroy.test.ts`. ──────────────
const GOOGLE_CSS = `@font-face { font-family: 'Carlito'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2'); }`;
function installGoogleFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: string, public descriptors?: object) {}
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  const set = {
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  G.fetch = async () => ({ ok: true, text: async () => GOOGLE_CSS });
  delete G.self;
  return { added };
}

const GOOGLE_FONT_MAP: Record<string, FontPreloadEntry> = {
  calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
};

/** A minimal valid sfnt header so registerEmbeddedFonts accepts the face. */
const validHeader = (): Uint8Array =>
  new Uint8Array([
    0x00, 0x01, 0x00, 0x00, 0x00, 0x10, 0x01, 0x00, 0x00, 0x40, 0x00, 0x30,
    0x47, 0x53, 0x55, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  ]);

interface DestroyProbe {
  destroy(): void;
  getBookmarkPage(bookmarkName: string): number | undefined;
}

describe('DocxDocument.destroy() — rejects in-flight worker requests', () => {
  function makeDocument() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
    attachDocumentLayoutRuntime(instance, 0);
    instance._bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance._rawParts = new BoundedRawPartCache({ maxEntries: 4, maxBytes: 1024 });
    instance._embeddedFontFaces = [];
    instance._googleFontFaces = [];
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { doc: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { doc, bridge, worker } = makeDocument();
    // A request the worker will never answer.
    const inFlight = bridge.request((id) => ({ id }));
    doc.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice (second terminate has nothing pending)', () => {
    const { doc } = makeDocument();
    doc.destroy();
    expect(() => doc.destroy()).not.toThrow();
  });

  it('terminates the owned worker when a partially initialized load rejects', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      DocxDocument.prototype as unknown as {
        _parse(
          buffer: ArrayBuffer,
          resourcePolicy: object,
          useGoogleFonts?: boolean,
          timeoutMs?: number,
        ): Promise<void>;
      },
      '_parse',
    ).mockRejectedValueOnce(failure);

    await expect(DocxDocument.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('preserves the load error and terminates directly when destroy throws', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const failure = new Error('injected load failure');
    vi.spyOn(
      DocxDocument.prototype as unknown as {
        _parse(
          buffer: ArrayBuffer,
          resourcePolicy: object,
          useGoogleFonts?: boolean,
          timeoutMs?: number,
        ): Promise<void>;
      },
      '_parse',
    ).mockRejectedValueOnce(failure);
    vi.spyOn(DocxDocument.prototype, 'destroy').mockImplementationOnce(() => {
      throw new Error('cleanup failure');
    });

    await expect(DocxDocument.load(new ArrayBuffer(0))).rejects.toBe(failure);
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('terminates directly when construction fails before the factory owns an instance', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'not a valid base URL' };

    await expect(
      DocxDocument.load(new ArrayBuffer(0), { wasmUrl: 'relative.wasm' }),
    ).rejects.toThrow();
    expect(SilentWorker.instances).toHaveLength(1);
    expect(SilentWorker.instances[0].terminated).toBe(true);
  });

  it('rejects invalid resource options before fetch or worker creation', async () => {
    G.Worker = SilentWorker;
    G.location = { href: 'http://localhost/' };
    const fetch = vi.fn();
    G.fetch = fetch;

    await expect(
      DocxDocument.load('/document.docx', {
        resourceLimits: { maxArchiveEntryBytes: 0 },
      }),
    ).rejects.toThrow(/resourceLimits\.maxArchiveEntryBytes/);
    expect(fetch).not.toHaveBeenCalled();
    expect(SilentWorker.instances).toHaveLength(0);
  });

  it('returns no bookmark before load or after destroy without poisoning loaded lookup', () => {
    const { doc } = makeDocument();

    expect(doc.getBookmarkPage('loaded')).toBeUndefined();
    (doc as unknown as { _meta: { bookmarkPages: [string, number][] } })._meta = {
      bookmarkPages: [['loaded', 3]],
    };
    expect(doc.getBookmarkPage('loaded')).toBe(3);

    doc.destroy();
    expect(doc.getBookmarkPage('loaded')).toBeUndefined();
  });

  // Wiring guard: destroy() must actually release the embedded fonts the document
  // registered into the shared FontFaceSet. The other tests set
  // `_embeddedFontFaces = []`, so they never exercise the unregister branch — a
  // dropped call (or a wrong field name) would go unnoticed. Register a real face
  // through core, hand it to the document, then assert destroy() removes it from
  // the (fake) FontFaceSet and clears the held array.
  it('destroy() releases the document’s embedded fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await registerEmbeddedFonts([
      { family: 'DocxEmbedded', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(added).toHaveLength(1); // the face is in the shared set

    const { doc } = makeDocument();
    (doc as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces = held;
    doc.destroy();

    // destroy() called unregisterEmbeddedFonts(held): last holder gone → the face
    // left the FontFaceSet, and the held array was cleared.
    expect(added).toHaveLength(0);
    expect((doc as unknown as { _embeddedFontFaces: FontFace[] })._embeddedFontFaces).toHaveLength(0);
  });

  // Wiring guard: destroy() must actually release the Google-Fonts substitutes
  // the document preloaded into the shared FontFaceSet. The other tests set
  // `_googleFontFaces = []`, so they never exercise the unload branch — a
  // dropped call (or a wrong field name) would go unnoticed. Preload a real
  // face through core, hand it to the document, then assert destroy() removes
  // it from the (fake) FontFaceSet and clears the held array. Twin of the
  // embedded-fonts guard above; same shape as
  // `presentation-destroy.test.ts` / `workbook-destroy.test.ts`.
  it('destroy() releases the document’s Google fonts from the FontFaceSet', async () => {
    const { added } = installGoogleFontFaceSet();
    const held = await preloadGoogleFonts(['Calibri'], GOOGLE_FONT_MAP);
    expect(added).toHaveLength(1); // the web font is in the shared set

    const { doc } = makeDocument();
    (doc as unknown as { _googleFontFaces: FontFace[] })._googleFontFaces = held;
    doc.destroy();

    // destroy() called unloadGoogleFonts(held): last holder gone → the face
    // left the FontFaceSet, and the held array was cleared.
    expect(added).toHaveLength(0);
    expect((doc as unknown as { _googleFontFaces: FontFace[] })._googleFontFaces).toHaveLength(0);
  });

});
