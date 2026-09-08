import { describe, it, expect, afterEach, vi } from 'vitest';
import { preloadGoogleFonts, unloadGoogleFonts, parseFontFaceRules, _resetCssCacheForTests, type FontPreloadEntry } from './preload.js';
import { _resetFontRegistryForTests } from './font-registry.js';
import { SCRIPT_GOOGLE_FONTS } from './scripts.js';

const G = globalThis as Record<string, unknown>;
const ORIG = { document: G.document, self: G.self, fetch: G.fetch, FontFace: G.FontFace };

afterEach(() => {
  G.document = ORIG.document;
  G.self = ORIG.self;
  G.fetch = ORIG.fetch;
  G.FontFace = ORIG.FontFace;
  _resetCssCacheForTests();
  _resetFontRegistryForTests(); // the FontFace refcount registry is module-global
  vi.restoreAllMocks();
});

const CSS = `
/* latin */
@font-face {
  font-family: 'Carlito';
  font-style: italic;
  font-weight: 700;
  src: url(https://fonts.gstatic.com/s/carlito/x.woff2) format('woff2');
  unicode-range: U+0000-00FF;
}
@font-face {
  font-family: 'Carlito';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2');
}
`;

describe('parseFontFaceRules', () => {
  it('extracts family, src and descriptors from @font-face blocks', () => {
    const faces = parseFontFaceRules(CSS);
    expect(faces).toHaveLength(2);
    expect(faces[0].family).toBe('Carlito');
    expect(faces[0].src).toContain('woff2');
    expect(faces[0].descriptors).toMatchObject({
      style: 'italic',
      weight: '700',
      unicodeRange: 'U+0000-00FF',
    });
    expect(faces[1].descriptors.style).toBe('normal');
  });
});

interface FakeFace {
  family: string;
  source: string;
  loadCalls: number;
  descriptors?: object;
  load: () => Promise<FakeFace>;
}

function installFakes(opts: { failLoad?: boolean; quoteFamily?: boolean } = {}) {
  const added: FakeFace[] = [];
  class FakeFontFace implements FakeFace {
    family: string; source: string; loadCalls = 0;
    constructor(family: string, source: string, public descriptors?: object) {
      // Chrome serializes a multi-word FontFace.family back WITH quotes
      // (e.g. `"Nunito Sans"`). quoteFamily reproduces that so the loader is
      // verified NOT to depend on matching the set by family string.
      this.family = opts.quoteFamily ? `"${family}"` : family;
      this.source = source;
    }
    load(): Promise<FakeFace> {
      this.loadCalls++;
      return opts.failLoad ? Promise.reject(new Error('net')) : Promise.resolve(this);
    }
  }
  const set = {
    faces: added,
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.fetch = vi.fn(async () => ({ ok: true, text: async () => CSS }));
  return { set, added };
}

const MAP: Record<string, FontPreloadEntry> = {
  calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
};

describe('preloadGoogleFonts', () => {
  it('trims a native Noto CJK name and fetches the aliased Google family', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const notoCss = CSS.replaceAll('Carlito', 'Noto Sans SC');
    G.fetch = vi.fn(async () => ({ ok: true, text: async () => notoCss }));

    await preloadGoogleFonts(['  NoTo SaNs CjK Sc  '], SCRIPT_GOOGLE_FONTS);

    expect(G.fetch).toHaveBeenCalledWith(expect.stringContaining('family=Noto+Sans+SC'));
    expect(added.map((face) => face.family)).toEqual(['Noto Sans SC', 'Noto Sans SC']);
    expect(added.every((face) => face.loadCalls === 1)).toBe(true);
  });

  it('fetches and registers the Google Fonts HK sans family for its native alias', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const notoCss = CSS.replaceAll('Carlito', 'Noto Sans HK');
    G.fetch = vi.fn(async () => ({ ok: true, text: async () => notoCss }));

    await preloadGoogleFonts(['Noto Sans CJK HK'], SCRIPT_GOOGLE_FONTS);

    expect(G.fetch).toHaveBeenCalledWith(expect.stringContaining('family=Noto+Sans+HK'));
    expect(added.map((face) => face.family)).toEqual(['Noto Sans HK', 'Noto Sans HK']);
  });

  it('does not request a Google Fonts substitute for the native HK serif family', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const fetch = vi.fn();
    G.fetch = fetch;

    await preloadGoogleFonts(['Noto Serif CJK HK'], SCRIPT_GOOGLE_FONTS);

    expect(fetch).not.toHaveBeenCalled();
  });

  it('fetches the CSS once, registers FontFaces and force-loads them (document.fonts)', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    await preloadGoogleFonts(['Calibri'], MAP);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
    expect(added.map((f) => f.family)).toEqual(['Carlito', 'Carlito']);
    expect(added.every((f) => f.loadCalls === 1)).toBe(true);
  });

  it('force-loads faces even when FontFace.family serializes with quotes (Chrome multi-word)', async () => {
    // Regression: the loader must load the FontFace objects it created, not
    // re-select them from the set by family string. Chrome returns a quoted
    // `.family` for multi-word names, so a family-string filter matches nothing
    // and the fonts silently never load (worker OffscreenCanvas falls back).
    const { added } = installFakes({ quoteFamily: true });
    G.document = { fonts: { faces: added, add: (f: FakeFace) => added.push(f), [Symbol.iterator]() { return added[Symbol.iterator](); }, ready: Promise.resolve() } };
    delete G.self;
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(added.length).toBe(2);
    expect(added.every((f) => f.loadCalls === 1)).toBe(true); // loaded despite quoted .family
  });

  it('uses self.fonts when document is undefined (worker)', async () => {
    const { set, added } = installFakes();
    delete G.document;
    G.self = { fonts: set };
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(2);
  });

  it('registers faces in an explicitly selected popup FontFaceSet', async () => {
    const { set: openerSet, added: openerFaces } = installFakes();
    const popupFaces: FakeFace[] = [];
    const popupSet = {
      add: (face: FakeFace) => popupFaces.push(face),
      delete: (face: FakeFace) => {
        const index = popupFaces.indexOf(face);
        if (index >= 0) popupFaces.splice(index, 1);
        return index >= 0;
      },
      ready: Promise.resolve(),
    } as unknown as FontFaceSet;
    G.document = { fonts: openerSet };

    const openerHeld = await preloadGoogleFonts(['Calibri'], MAP);
    const popupHeld = await preloadGoogleFonts(['Calibri'], MAP, popupSet);
    expect(openerFaces).toHaveLength(2);
    expect(popupFaces).toHaveLength(2);
    expect(popupHeld[0]).not.toBe(openerHeld[0]);

    unloadGoogleFonts(popupHeld);
    expect(popupFaces).toHaveLength(0);
    expect(openerFaces).toHaveLength(2);
    unloadGoogleFonts(openerHeld);
  });

  it('no-ops (returns []) without any FontFaceSet', async () => {
    delete G.document;
    delete G.self;
    await expect(preloadGoogleFonts(['Calibri'], MAP)).resolves.toEqual([]);
  });

  it('does not refetch a CSS url already registered in this context', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    await preloadGoogleFonts(['Calibri'], MAP);
    await preloadGoogleFonts(['Calibri'], MAP);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
  });

  it('warns once when faces fail to load, and still resolves', async () => {
    const { set } = installFakes({ failLoad: true });
    G.document = { fonts: set };
    delete G.self;
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toContain('carlito');
  });

  it('warns once on a non-ok fetch and frees the slot so a later call retries', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    // Override fetch with a 404 so registration fails (res.ok === false).
    G.fetch = vi.fn(async () => ({ ok: false, status: 404, text: async () => '' }));
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await preloadGoogleFonts(['Calibri'], MAP);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0][0]).toContain('carlito');
    // Failed url is removed from the cache, so a second call re-fetches.
    await preloadGoogleFonts(['Calibri'], MAP);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(2);
  });

  it('joins a concurrent same-url call: second does not resolve before the first registers', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    // Defer the fetch so both calls overlap with the registration in flight.
    let resolveFetch!: (v: { ok: boolean; text: () => Promise<string> }) => void;
    const fetchPromise = new Promise<{ ok: boolean; text: () => Promise<string> }>((r) => {
      resolveFetch = r;
    });
    G.fetch = vi.fn(() => fetchPromise);

    let aDone = false;
    let bDone = false;
    const a = preloadGoogleFonts(['Calibri'], MAP).then(() => { aDone = true; });
    const b = preloadGoogleFonts(['Calibri'], MAP).then(() => { bDone = true; });

    // Let microtasks flush; neither call may settle while the fetch is pending.
    await Promise.resolve();
    await Promise.resolve();
    expect(aDone).toBe(false);
    expect(bDone).toBe(false);

    resolveFetch({ ok: true, text: async () => CSS });
    await Promise.all([a, b]);
    expect(aDone).toBe(true);
    expect(bDone).toBe(true);
    // Faces are registered exactly once (the joined call reuses them via the
    // shared refcount registry, not re-adding) and the url is fetched once. The
    // first holder force-loads each face; the second reuses it (already loaded).
    expect(added.map((f) => f.family)).toEqual(['Carlito', 'Carlito']);
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
    expect(added.every((f) => f.loadCalls >= 1)).toBe(true);
  });
});

// ---- dedup + destroy() cleanup (Google Fonts preload leak) ------------------
//
// document.fonts is a process-global singleton, so preloading the same Google
// Font twice (two documents, or one document reopened in an SPA) must NOT add a
// second identical FontFace; and a document's destroy() must remove the faces it
// preloaded so they don't accumulate forever. Same refcount contract as the
// embedded-font loader — both share the core FontFace registry.
describe('preloadGoogleFonts — dedup + unloadGoogleFonts (SPA leak)', () => {
  it('does not add the same web font twice (dedup by signature)', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    // Simulate preloading the same font in two documents.
    const a = await preloadGoogleFonts(['Calibri'], MAP);
    const b = await preloadGoogleFonts(['Calibri'], MAP);
    // Only the two @font-face rules were added ONCE; the second preload reused them.
    expect(added.map((f) => f.family)).toEqual(['Carlito', 'Carlito']);
    // Both callers hold the SAME shared FontFace references.
    expect(a).toHaveLength(2);
    expect(b).toHaveLength(2);
    expect(a[0]).toBe(b[0]);
    expect(a[1]).toBe(b[1]);
    // Each shared face is force-loaded exactly once (not per preload).
    expect(added.every((f) => f.loadCalls === 1)).toBe(true);
    // And the stylesheet was fetched once.
    expect((G.fetch as ReturnType<typeof vi.fn>)).toHaveBeenCalledTimes(1);
  });

  it('destroy() removes the document’s web fonts from the FontFaceSet', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const held = await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(2);
    unloadGoogleFonts(held); // what DocxDocument/PptxPresentation/XlsxWorkbook.destroy() calls
    expect(added).toHaveLength(0); // faces left the set
  });

  it('a shared web font survives until the LAST holder releases it (refcount)', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const a = await preloadGoogleFonts(['Calibri'], MAP); // document A
    const b = await preloadGoogleFonts(['Calibri'], MAP); // document B (same font)
    expect(added).toHaveLength(2);

    unloadGoogleFonts(a); // document A destroyed
    expect(added).toHaveLength(2); // still referenced by B → faces stay

    unloadGoogleFonts(b); // document B destroyed
    expect(added).toHaveLength(0); // last holder gone → faces removed
  });

  it('re-preloading after a full release adds fresh faces (no stale registry entry)', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const a = await preloadGoogleFonts(['Calibri'], MAP);
    unloadGoogleFonts(a);
    expect(added).toHaveLength(0);
    // Opening the same font again after everyone released it registers anew.
    const b = await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(2);
    expect(b[0]).not.toBe(a[0]); // distinct FontFace object (a's was deleted)
  });

  it('unloadGoogleFonts is a no-op for unknown / already-released faces', () => {
    installFakes();
    G.document = { fonts: undefined };
    const stray = { family: 'Stray' } as unknown as FontFace;
    expect(() => unloadGoogleFonts([stray])).not.toThrow();
  });

  it('the same face passed twice in ONE release is decremented at most once (no over-decrement)', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const a = await preloadGoogleFonts(['Calibri'], MAP); // document A
    const b = await preloadGoogleFonts(['Calibri'], MAP); // document B (shares faces)
    expect(added).toHaveLength(2);
    expect(a[0]).toBe(b[0]); // same shared FontFace, refs === 2

    // Document A releases, but its held list mistakenly contains a face twice.
    // The per-call `seen` guard collapses the duplicate to a single decrement,
    // so the face STAYS in the set (B still uses it).
    unloadGoogleFonts([a[0], a[0], a[1]]);
    expect(added).toHaveLength(2); // survived — B still holds them

    unloadGoogleFonts(b); // B releasing drops the last references
    expect(added).toHaveLength(0);
  });

  it('shares one FontFace across two formats requesting the same stylesheet (cross-loader dedup)', async () => {
    // Two DIFFERENT preload maps (e.g. docx and xlsx) that point at the same
    // Google Fonts CSS url for the same substitute compute the same signature,
    // so the FontFace is added once and refcounted across both. This proves the
    // registry is genuinely shared, not per-map.
    const { set, added } = installFakes();
    G.document = { fonts: set };
    delete G.self;
    const MAP_A: Record<string, FontPreloadEntry> = {
      calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
    };
    const MAP_B: Record<string, FontPreloadEntry> = {
      // A different Office name mapping to the SAME substitute + url.
      'calibri light': { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
    };
    const a = await preloadGoogleFonts(['Calibri'], MAP_A);
    const b = await preloadGoogleFonts(['Calibri Light'], MAP_B);
    expect(added).toHaveLength(2); // added once despite two maps
    expect(a[0]).toBe(b[0]);

    unloadGoogleFonts(a);
    expect(added).toHaveLength(2); // B still holds
    unloadGoogleFonts(b);
    expect(added).toHaveLength(0);
  });
});



describe('custom Google Fonts CSS origins', () => {
  it.each(['document', 'worker'])('uses a custom CSS origin in %s', async (context) => {
    const { set, added } = installFakes();
    delete G.document;
    delete G.self;
    G[context === 'document' ? 'document' : 'self'] = { fonts: set };
    const css = CSS.replaceAll('https://fonts.gstatic.com', 'https://fonts.internal.example');
    G.fetch = vi.fn(async () => ({ ok: true, text: async () => css }));

    await preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.internal.example/');

    expect(G.fetch).toHaveBeenCalledWith('https://fonts.internal.example/css2?family=Carlito');
    expect(added).toHaveLength(2);
    expect(added.every((face) => face.source.includes('https://fonts.internal.example/'))).toBe(true);
    expect(added.every((face) => face.loadCalls === 1)).toBe(true);
    expect(added[0].descriptors).toMatchObject({ style: 'italic', weight: '700', unicodeRange: 'U+0000-00FF' });
    expect(MAP.calibri.url).toBe('https://fonts.googleapis.com/css2?family=Carlito');
  });

  it('preserves the endpoint path and family query with an HTTP origin and port', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    const map = { calibri: { ...MAP.calibri, url: MAP.calibri.url + ':ital,wght@0,400;1,700&display=swap' } };
    await preloadGoogleFonts(['Calibri'], map, undefined, 'http://fonts.internal.example:8080');
    expect(G.fetch).toHaveBeenCalledWith('http://fonts.internal.example:8080/css2?family=Carlito:ital,wght@0,400;1,700&display=swap');
  });

  it('uses the font URLs returned by the China CSS service', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    G.fetch = vi.fn(async () => ({ ok: true, text: async () => CSS.replaceAll('fonts.gstatic.com', 'fonts.gstatic.cn') }));
    await preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.googleapis.cn');
    expect(G.fetch).toHaveBeenCalledWith('https://fonts.googleapis.cn/css2?family=Carlito');
    expect(added[0].source).toContain('https://fonts.gstatic.cn/');
  });

  it('isolates concurrent origins and deduplicates normalized origins independently', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    const [globalFaces, internalFaces, internalAgain] = await Promise.all([
      preloadGoogleFonts(['Calibri'], MAP),
      preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.internal.example'),
      preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.internal.example/'),
    ]);
    expect(G.fetch).toHaveBeenCalledTimes(2);
    expect(globalFaces[0]).not.toBe(internalFaces[0]);
    expect(internalFaces[0]).toBe(internalAgain[0]);
    unloadGoogleFonts(internalFaces);
    unloadGoogleFonts(internalAgain);
    expect(set.faces).toEqual(globalFaces);
    unloadGoogleFonts(globalFaces);
    expect(set.faces).toHaveLength(0);
  });

  it.each([
    ['url("../files/a.woff2")', 'url("https://fonts.internal.example/files/a.woff2")'],
    ["url('/files/a.woff2')", 'url("https://fonts.internal.example/files/a.woff2")'],
    ['url(a.woff2)', 'url("https://fonts.internal.example/styles/a.woff2")'],
    ['url(//assets.internal.example/a.woff2)', 'url("https://assets.internal.example/a.woff2")'],
    ['url(https://fonts.gstatic.com/s/a.woff2)', 'url(https://fonts.gstatic.com/s/a.woff2)'],
  ])('resolves font sources relative to the final stylesheet URL: %s', async (src, expected) => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    G.fetch = vi.fn(async () => ({
      ok: true, url: 'https://fonts.internal.example/styles/redirected.css',
      text: async () => `@font-face { font-family: Carlito; src: ${src}; }`,
    }));
    await preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.internal.example');
    expect(added[0].source).toBe(expected);
  });

  it('resolves a relative font source against the requested URL when response.url is unavailable', async () => {
    const { set, added } = installFakes();
    G.document = { fonts: set };
    G.fetch = vi.fn(async () => ({ ok: true, text: async () => '@font-face { font-family: Carlito; src: local("Carlito"), url(/fonts/a.woff2) format("woff2"); }' }));
    await preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.internal.example');
    expect(added[0].source).toBe('local("Carlito"), url("https://fonts.internal.example/fonts/a.woff2") format("woff2")');
  });

  it('leaves unrelated stylesheet hosts unchanged', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    const map = { calibri: { ...MAP.calibri, url: 'https://fonts.googleapis.com.example.com/css2?family=Carlito' } };
    await preloadGoogleFonts(['Calibri'], map, undefined, 'https://fonts.internal.example');
    expect(G.fetch).toHaveBeenCalledExactlyOnceWith(map.calibri.url);
  });

  it.each(['', '/fonts', 'file:///fonts', 'ftp://fonts.example', 'https://fonts.example/proxy', 'https://fonts.example?key=1', 'https://fonts.example#fragment', 'https://user:pass@fonts.example'])('rejects an invalid origin before fetching: %s', async (origin) => {
    const { set } = installFakes();
    G.document = { fonts: set };
    await expect(preloadGoogleFonts(['Calibri'], MAP, undefined, origin)).rejects.toThrow(/googleFontsCssOrigin/);
    expect(G.fetch).not.toHaveBeenCalled();
  });

  it('does not fall back to a public origin after a custom service fails', async () => {
    const { set } = installFakes();
    G.document = { fonts: set };
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    G.fetch = vi.fn(async () => ({ ok: false, status: 503 }));
    expect(await preloadGoogleFonts(['Calibri'], MAP, undefined, 'https://fonts.internal.example')).toEqual([]);
    expect(G.fetch).toHaveBeenCalledExactlyOnceWith('https://fonts.internal.example/css2?family=Carlito');
  });
});
