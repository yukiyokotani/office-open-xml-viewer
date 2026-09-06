import { describe, it, expect, afterEach, vi } from 'vitest';
import {
  deobfuscateOdttf,
  embeddedFontBytesAreWithinLimit,
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  _resetEmbeddedRegistryForTests,
  type EmbeddedFontFace,
} from './embedded.js';

const G = globalThis as Record<string, unknown>;
const ORIG = { document: G.document, self: G.self, FontFace: G.FontFace };

afterEach(() => {
  G.document = ORIG.document;
  G.self = ORIG.self;
  G.FontFace = ORIG.FontFace;
  _resetEmbeddedRegistryForTests(); // dedup registry is module-global
  vi.restoreAllMocks();
});

describe('deobfuscateOdttf (ECMA-376 §17.8.1)', () => {
  // Synthetic: obfuscate a known 32-byte header with a GUID, then de-obfuscate
  // and assert we recover the original. Because XOR is its own inverse, the
  // obfuscate step is the same routine — so a round-trip proving recovery also
  // proves the direction the viewer needs (obfuscated part → plaintext font).
  it('is an involution: obfuscate then de-obfuscate recovers the bytes', () => {
    const guid = '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}';
    // A plausible sfnt header: 00 01 00 00 (TrueType) + arbitrary tail.
    const plain = new Uint8Array(64);
    plain.set([0x00, 0x01, 0x00, 0x00], 0);
    for (let i = 4; i < 64; i++) plain[i] = (i * 37) & 0xff;

    const obf = deobfuscateOdttf(plain, guid); // XOR masks first 32 bytes
    // First 32 bytes changed, the rest untouched.
    expect(Array.from(obf.slice(0, 4))).not.toEqual([0x00, 0x01, 0x00, 0x00]);
    expect(Array.from(obf.slice(32))).toEqual(Array.from(plain.slice(32)));

    const back = deobfuscateOdttf(obf, guid);
    expect(Array.from(back)).toEqual(Array.from(plain));
  });

  it('recovers the TrueType magic 00 01 00 00 from a real ODTTF header slice', () => {
    // First 32 bytes of sample-11.docx word/fonts/font1.odttf (Ubuntu Regular),
    // whose w:embedRegular fontKey is {3EEE3167-...}. De-obfuscating must yield
    // the sfnt version 0x00010000 in the first four bytes.
    const header = new Uint8Array([
      0x4d, 0x1c, 0xe3, 0x71, 0x6b, 0xff, 0x49, 0xae, 0x98, 0x43, 0xb8, 0xb5,
      0x23, 0x62, 0xa7, 0x79, 0x15, 0x12, 0x90, 0x39, 0x6b, 0xef, 0x04, 0x5e,
      0x98, 0x47, 0xa1, 0xd5, 0x20, 0x61, 0xa1, 0x6d,
    ]);
    const out = deobfuscateOdttf(header, '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}');
    expect(Array.from(out.slice(0, 4))).toEqual([0x00, 0x01, 0x00, 0x00]);
  });

  it('accepts a GUID without braces/hyphens', () => {
    const header = new Uint8Array([
      0x4d, 0x1c, 0xe3, 0x71, 0x6b, 0xff, 0x49, 0xae, 0x98, 0x43, 0xb8, 0xb5,
      0x23, 0x62, 0xa7, 0x79, 0x15, 0x12, 0x90, 0x39, 0x6b, 0xef, 0x04, 0x5e,
      0x98, 0x47, 0xa1, 0xd5, 0x20, 0x61, 0xa1, 0x6d,
    ]);
    const out = deobfuscateOdttf(header, '3EEE3167E5B84798AE48EA6B71E31D4D');
    expect(Array.from(out.slice(0, 4))).toEqual([0x00, 0x01, 0x00, 0x00]);
  });

  it('leaves bytes past the first 32 untouched even for a short buffer', () => {
    const short = new Uint8Array([1, 2, 3, 4, 5]);
    const out = deobfuscateOdttf(short, '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}');
    expect(out.length).toBe(5);
    // Round trip still recovers.
    expect(Array.from(deobfuscateOdttf(out, '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}'))).toEqual([
      1, 2, 3, 4, 5,
    ]);
  });

  it('throws on a malformed fontKey (not 32 hex digits)', () => {
    expect(() => deobfuscateOdttf(new Uint8Array(32), 'not-a-guid')).toThrow(/invalid fontKey/);
    expect(() => deobfuscateOdttf(new Uint8Array(32), '{ABCD}')).toThrow(/invalid fontKey/);
  });

  it('does not mutate the input', () => {
    const input = new Uint8Array([0xaa, 0xbb, 0xcc, 0xdd]);
    const copy = input.slice();
    deobfuscateOdttf(input, '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}');
    expect(Array.from(input)).toEqual(Array.from(copy));
  });
});

// ---- registerEmbeddedFonts -------------------------------------------------

interface FakeFace {
  family: string;
  source: ArrayBuffer;
  descriptors?: object;
  loadCalls: number;
  load: () => Promise<FakeFace>;
}

function installFontFaceSet(opts: { failLoad?: (family: string) => boolean } = {}) {
  const added: FakeFace[] = [];
  class FakeFontFace implements FakeFace {
    family: string;
    source: ArrayBuffer;
    loadCalls = 0;
    constructor(family: string, source: ArrayBuffer, public descriptors?: object) {
      this.family = family;
      this.source = source;
    }
    load(): Promise<FakeFace> {
      this.loadCalls++;
      return opts.failLoad?.(this.family)
        ? Promise.reject(new Error('load failed'))
        : Promise.resolve(this);
    }
  }
  const set = {
    faces: added,
    add: (f: FakeFace) => {
      added.push(f);
    },
    delete: (f: FakeFace) => {
      const i = added.indexOf(f);
      if (i >= 0) added.splice(i, 1);
      return i >= 0;
    },
    [Symbol.iterator]() {
      return added[Symbol.iterator]();
    },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  delete G.self;
  return { set, added };
}

const validHeader = () =>
  new Uint8Array([
    0x00, 0x01, 0x00, 0x00, 0x00, 0x10, 0x01, 0x00, 0x00, 0x40, 0x00, 0x30,
    0x47, 0x53, 0x55, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
  ]);

describe('registerEmbeddedFonts', () => {
  it('shares the exact admission boundary with pre-registration inspectors', () => {
    expect(embeddedFontBytesAreWithinLimit(new Uint8Array(10), 10)).toBe(true);
    expect(embeddedFontBytesAreWithinLimit(new Uint8Array(0), 10)).toBe(false);
    expect(embeddedFontBytesAreWithinLimit(new Uint8Array(11), 10)).toBe(false);
  });

  it('registers a raw (pptx) face and force-loads it', async () => {
    const { added } = installFontFaceSet();
    const faces: EmbeddedFontFace[] = [
      { family: 'MyFont', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ];
    await registerEmbeddedFonts(faces);
    expect(added).toHaveLength(1);
    expect(added[0].family).toBe('MyFont');
    expect(added[0].loadCalls).toBe(1);
    expect(added[0].descriptors).toMatchObject({ weight: 'normal', style: 'normal' });
  });

  it('de-obfuscates an odttf face before registering (bytes passed to FontFace are plaintext)', async () => {
    const { added } = installFontFaceSet();
    // Obfuscate a valid header so registration must de-obfuscate to a valid sfnt.
    const guid = '{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}';
    const obf = deobfuscateOdttf(validHeader(), guid);
    await registerEmbeddedFonts([
      { family: 'Ubuntu', bytes: obf, odttf: true, fontKey: guid, weight: 'bold', style: 'italic' },
    ]);
    expect(added).toHaveLength(1);
    const src = new Uint8Array(added[0].source);
    expect(Array.from(src.slice(0, 4))).toEqual([0x00, 0x01, 0x00, 0x00]);
    expect(added[0].descriptors).toMatchObject({ weight: 'bold', style: 'italic' });
  });

  it('maps all four style slots (regular/bold/italic/boldItalic) to one family', async () => {
    const { added } = installFontFaceSet();
    const faces: EmbeddedFontFace[] = [
      { family: 'Ubuntu', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
      { family: 'Ubuntu', bytes: validHeader(), odttf: false, weight: 'bold', style: 'normal' },
      { family: 'Ubuntu', bytes: validHeader(), odttf: false, weight: 'normal', style: 'italic' },
      { family: 'Ubuntu', bytes: validHeader(), odttf: false, weight: 'bold', style: 'italic' },
    ];
    await registerEmbeddedFonts(faces);
    expect(added).toHaveLength(4);
    expect(added.every((f) => f.family === 'Ubuntu')).toBe(true);
  });

  it('skips a malformed odttf fontKey without aborting other faces', async () => {
    const { added } = installFontFaceSet();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await registerEmbeddedFonts([
      { family: 'Bad', bytes: validHeader(), odttf: true, fontKey: 'nope', weight: 'normal', style: 'normal' },
      { family: 'Good', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(added.map((f) => f.family)).toEqual(['Good']);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('Bad'));
  });

  it('skips an oversized face (memory / zip-bomb safety net)', async () => {
    const { added } = installFontFaceSet();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await registerEmbeddedFonts(
      [{ family: 'Huge', bytes: new Uint8Array(100), odttf: false, weight: 'normal', style: 'normal' }],
      10, // maxBytes = 10
    );
    expect(added).toHaveLength(0);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('Huge'));
  });

  it('skips an empty face', async () => {
    const { added } = installFontFaceSet();
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    await registerEmbeddedFonts([
      { family: 'Empty', bytes: new Uint8Array(0), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(added).toHaveLength(0);
  });

  it('warns once when a registered face fails to load', async () => {
    const { added } = installFontFaceSet({ failLoad: (f) => f === 'Flaky' });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const held = await registerEmbeddedFonts([
      { family: 'Flaky', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(added).toHaveLength(0);
    expect(held).toEqual([]);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('Flaky'));
  });

  it('no-ops (returns []) when no FontFaceSet exists', async () => {
    delete G.document;
    delete G.self;
    delete G.FontFace;
    await expect(
      registerEmbeddedFonts([
        { family: 'X', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
      ]),
    ).resolves.toEqual([]);
  });

  it('registers into self.fonts when there is no document (worker context)', async () => {
    const added: FakeFace[] = [];
    class FakeFontFace {
      constructor(public family: string, public source: ArrayBuffer, public descriptors?: object) {}
      load() {
        return Promise.resolve(this);
      }
    }
    delete G.document;
    G.self = {
      fonts: {
        add: (f: FakeFace) => added.push(f),
        ready: Promise.resolve(),
      },
    };
    G.FontFace = FakeFontFace;
    await registerEmbeddedFonts([
      { family: 'WorkerFont', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(added).toHaveLength(1);
    expect(added[0].family).toBe('WorkerFont');
  });

  it('returns the added FontFace objects for the caller to hold', async () => {
    installFontFaceSet();
    const held = await registerEmbeddedFonts([
      { family: 'Held', bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal' },
    ]);
    expect(held).toHaveLength(1);
    expect((held[0] as unknown as FakeFace).family).toBe('Held');
  });
});

// ---- dedup + destroy() cleanup (#762) --------------------------------------
//
// document.fonts is a process-global singleton, so opening the same embedded
// font twice (two documents, or one document reopened in an SPA) must NOT add a
// second byte-identical FontFace; and a document's destroy() must remove the
// faces it registered so they don't accumulate forever.
describe('registerEmbeddedFonts — dedup + unregisterEmbeddedFonts (#762)', () => {
  const font = (family = 'Ubuntu'): EmbeddedFontFace => ({
    family, bytes: validHeader(), odttf: false, weight: 'normal', style: 'normal',
  });

  it('does not add a byte-identical face twice (dedup by content signature)', async () => {
    const { added } = installFontFaceSet();
    // Simulate opening the same embedded font in two documents.
    const a = await registerEmbeddedFonts([font()]);
    const b = await registerEmbeddedFonts([font()]);
    // Only ONE FontFace was added to the set; the second registration reused it.
    expect(added).toHaveLength(1);
    // Both callers hold the SAME shared FontFace reference.
    expect(a).toHaveLength(1);
    expect(b).toHaveLength(1);
    expect(a[0]).toBe(b[0]);
    // Each holder observes the idempotent load result so a concurrent/reused
    // registration cannot mistake a failed first holder for a usable face.
    expect((added[0] as FakeFace).loadCalls).toBe(2);
  });

  it('distinct faces (different bytes) are NOT deduped', async () => {
    const { added } = installFontFaceSet();
    const other = new Uint8Array(validHeader());
    other[8] ^= 0xff; // perturb a byte → different content signature
    await registerEmbeddedFonts([font()]);
    await registerEmbeddedFonts([{ ...font(), bytes: other }]);
    expect(added).toHaveLength(2);
  });

  it('destroy() removes the document’s faces from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await registerEmbeddedFonts([font()]);
    expect(added).toHaveLength(1);
    unregisterEmbeddedFonts(held); // what DocxDocument.destroy() calls
    expect(added).toHaveLength(0); // face left the set
  });

  it('a shared font survives until the LAST holder releases it (refcount)', async () => {
    const { added } = installFontFaceSet();
    const a = await registerEmbeddedFonts([font()]); // document A
    const b = await registerEmbeddedFonts([font()]); // document B (same font)
    expect(added).toHaveLength(1);

    unregisterEmbeddedFonts(a); // document A destroyed
    // Still referenced by B → the face stays in the set.
    expect(added).toHaveLength(1);

    unregisterEmbeddedFonts(b); // document B destroyed
    // Last holder gone → the face is removed.
    expect(added).toHaveLength(0);
  });

  it('re-registering after a full release adds a fresh face (no stale registry entry)', async () => {
    const { added } = installFontFaceSet();
    const a = await registerEmbeddedFonts([font()]);
    unregisterEmbeddedFonts(a);
    expect(added).toHaveLength(0);
    // Opening the same font again after everyone released it registers anew.
    const b = await registerEmbeddedFonts([font()]);
    expect(added).toHaveLength(1);
    expect(b).toHaveLength(1);
  });

  it('unregisterEmbeddedFonts is a no-op for unknown / already-released faces', () => {
    installFontFaceSet();
    // A FontFace never registered here → nothing to remove, must not throw.
    const stray = { family: 'Stray' } as unknown as FontFace;
    expect(() => unregisterEmbeddedFonts([stray])).not.toThrow();
  });

  // Double-release safety: a stray double-release of one document's faces must
  // never evict a font a SECOND document is still using (no over-decrement).
  it('the same face passed twice in ONE call is decremented at most once (no over-decrement)', async () => {
    const { added } = installFontFaceSet();
    const a = await registerEmbeddedFonts([font()]); // document A
    const b = await registerEmbeddedFonts([font()]); // document B (shares the face)
    expect(added).toHaveLength(1);
    expect(a[0]).toBe(b[0]); // same shared FontFace, refs === 2

    // Document A releases, but its held list mistakenly contains the face twice.
    // Naively that decrements refs 2 → 0 and evicts the shared face while B still
    // uses it. The per-call `seen` guard must collapse the duplicate to a single
    // decrement (refs 2 → 1), so the face STAYS in the set.
    unregisterEmbeddedFonts([a[0], a[0]]);
    expect(added).toHaveLength(1); // survived — B still holds it

    // B releasing once now drops the last reference and removes the face.
    unregisterEmbeddedFonts(b);
    expect(added).toHaveLength(0);
  });

  it('re-releasing a FULLY-released face does not disturb a fresh re-registration of the same font', async () => {
    const { added } = installFontFaceSet();
    const a = await registerEmbeddedFonts([font()]); // document A, refs === 1
    expect(added).toHaveLength(1);

    unregisterEmbeddedFonts(a); // last (only) holder releases → face removed, entry gone
    expect(added).toHaveLength(0);

    // A different document opens the same font again — a BRAND-NEW FontFace object
    // and registry entry (refs === 1).
    const b = await registerEmbeddedFonts([font()]);
    expect(added).toHaveLength(1);
    expect(b[0]).not.toBe(a[0]); // distinct FontFace object (a's was deleted)

    // A stray second release of A's OLD, already-removed face must be a no-op —
    // it must NOT find and decrement the new registration (different identity) and
    // so must not evict B's font.
    unregisterEmbeddedFonts(a);
    expect(added).toHaveLength(1); // B's fresh registration untouched
  });
});
