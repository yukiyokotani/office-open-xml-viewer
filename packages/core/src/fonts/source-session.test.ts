import { setFlagsFromString } from 'node:v8';
import { runInNewContext } from 'node:vm';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { OoxmlError } from '../errors/ooxml-error.js';
import { HARD_MAX_FONT_SOURCE_FACE_BYTES, HARD_MAX_FONT_SOURCE_FAMILIES } from '../worker/resource-policy.generated.js';
import { googleFonts } from './presets/google-fonts.js';
import { liveSourceCacheCount, sourceCacheEntryCount } from './source-cache.js';
import { FontSourceSession, retainedFontSourceBytes, type FontSessionDemand } from './source-session.js';
import {
  fontCssApi,
  fontCss,
  fontFiles,
  fontResolver,
  type FontDemand,
  type FontFaceDeclaration,
  type FontResolveFn,
  type FontSource,
} from './source.js';
import { OFFICE_FONT_SUBSTITUTIONS } from './substitutions.js';

// ---- shims ---------------------------------------------------------------------------

const G = globalThis as Record<string, unknown>;
const ORIGINAL = { FontFace: G.FontFace, fetch: G.fetch, document: G.document, self: G.self };

class FakeFontFace {
  static loadImpl: (face: FakeFontFace) => Promise<void> = () => Promise.resolve();
  status: FontFaceLoadStatus = 'unloaded';
  loadCalls = 0;
  readonly weight: string;
  readonly style: string;
  readonly unicodeRange: string;
  private loading?: Promise<FakeFontFace>;
  constructor(readonly family: string, readonly source: string | ArrayBuffer, descriptors: FontFaceDescriptors = {}) {
    this.weight = descriptors.weight ?? 'normal';
    this.style = descriptors.style ?? 'normal';
    this.unicodeRange = descriptors.unicodeRange ?? 'U+0-10FFFF';
  }
  load(): Promise<FakeFontFace> {
    this.loadCalls++;
    this.loading ??= FakeFontFace.loadImpl(this).then(
      () => { this.status = 'loaded'; return this; },
      (error: unknown) => { this.status = 'error'; throw error; },
    );
    this.status = this.status === 'unloaded' ? 'loading' : this.status;
    return this.loading;
  }
}

function fakeSet() {
  const faces: FakeFontFace[] = [];
  return {
    faces,
    add: (face: FakeFontFace) => { faces.push(face); },
    delete: (face: FakeFontFace) => { const at = faces.indexOf(face); if (at >= 0) faces.splice(at, 1); return at >= 0; },
  };
}
type FakeSet = ReturnType<typeof fakeSet>;
const asSet = (set: FakeSet) => set as unknown as FontFaceSet;
const urlOf = (face: FakeFontFace) => (typeof face.source === 'string' ? /url\("(.*)"\)/.exec(face.source)?.[1] : undefined);

interface Route {
  css?: string;
  status?: number;
  /** Final URL after redirects. */
  url?: string;
  /** Resolves when the test lets the response through. */
  gate?: Promise<void>;
  /** Body size of a font file response. */
  bytes?: number;
}

function installFetch(routes: Record<string, Route>) {
  const calls: { url: string; init?: RequestInit }[] = [];
  const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
    calls.push({ url, init });
    const route = routes[url];
    if (route?.gate) {
      await new Promise<void>((resolve, reject) => {
        route.gate?.then(resolve);
        init?.signal?.addEventListener('abort', () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' })));
      });
    }
    if (!route) return { ok: false, status: 404, url, text: async (): Promise<string> => '' };
    return {
      ok: (route.status ?? 200) < 400, status: route.status ?? 200, url: route.url ?? url,
      text: async (): Promise<string> => route.css ?? '',
      arrayBuffer: async (): Promise<ArrayBuffer> => new Uint8Array(route.bytes ?? 0).fill(7).buffer,
    };
  });
  G.fetch = fetchMock;
  return { calls, fetchMock };
}

const rule = (family: string, file: string, extra = '') =>
  `@font-face { font-family: '${family}'; font-weight: 400; src: url(${file}) format('woff2'); ${extra} }\n`;

const LATIN = 'unicode-range: U+0000-00FF;';
const KANA = 'unicode-range: U+3000-30FF;';
const CJK = 'unicode-range: U+4E00-9FFF;';

function demand(families: string[], scalars: FontSessionDemand['scalars'] = 'all', fallbackFamilies: string[] = []): FontSessionDemand {
  return { families, fallbackFamilies, scalars };
}

const sessions: FontSourceSession[] = [];
function open(sources: readonly FontSource[], set: FakeSet, signal?: AbortSignal): FontSourceSession {
  const session = new FontSourceSession({ sources, fontSet: asSet(set), signal });
  sessions.push(session);
  return session;
}

let warn: ReturnType<typeof vi.spyOn>;
beforeEach(() => {
  G.FontFace = FakeFontFace;
  FakeFontFace.loadImpl = () => Promise.resolve();
  warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(() => {
  for (const session of sessions.splice(0)) session.release();
  expect(liveSourceCacheCount()).toBe(0);
  expect(retainedFontSourceBytes()).toBe(0);
  Object.assign(G, ORIGINAL);
  vi.restoreAllMocks();
  vi.useRealTimers();
});

// ---- tests ------------------------------------------------------------------------------

describe('FontSourceSession: lazy unicode-range loading', () => {
  const SHEET_URL = 'https://fonts.example/noto.css';
  const SHEET = rule('Noto Sans JP', 'latin.woff2', LATIN) + rule('Noto Sans JP', 'kana.woff2', KANA)
    + rule('Noto Sans JP', 'cjk.woff2', CJK);

  it('registers every subset in order but loads only those the scalars intersect', async () => {
    installFetch({ [SHEET_URL]: { css: SHEET } });
    const set = fakeSet();
    const session = open([fontCss({ url: SHEET_URL })], set);

    const first = await session.ensure(demand([], [0x4d, 0x41, 0x41], ['Noto Sans JP']));
    expect(set.faces.map(urlOf)).toEqual([
      'https://fonts.example/latin.woff2', 'https://fonts.example/kana.woff2', 'https://fonts.example/cjk.woff2',
    ]);
    expect(set.faces.map((face) => face.loadCalls)).toEqual([1, 0, 0]);
    expect(first.faces.map((face) => face.unicodeRange)).toEqual(['U+0000-00FF']);

    // A later delta loads only the newly intersected subset.
    const second = await session.ensure(demand([], [0x3042]));
    expect(set.faces.map((face) => face.loadCalls)).toEqual([1, 1, 0]);
    expect(second.faces).toHaveLength(2);

    await session.ensure(demand([], 'all'));
    expect(set.faces.map((face) => face.loadCalls)).toEqual([1, 1, 1]);
  });

  it('applies the same laziness to font-css-api presets and to font files', async () => {
    const preset = googleFonts();
    const carlito = 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap';
    installFetch({ [carlito]: { css: rule('Carlito', 'https://fonts.gstatic.example/c-latin.woff2', LATIN)
      + rule('Carlito', 'https://fonts.gstatic.example/c-greek.woff2', 'unicode-range: U+0370-03FF;') } });
    const set = fakeSet();
    const files = fontFiles({ faces: [
      { family: 'Bundled', url: 'https://app.example/b-latin.woff2', unicodeRange: 'U+0-FF' },
      { family: 'Bundled', url: 'https://app.example/b-jp.woff2', unicodeRange: 'U+3000-30FF' },
    ] });
    const loaded = await open([files, preset], set).ensure(demand(['Calibri', 'Bundled'], [0x61]));
    expect(set.faces.map((face) => [urlOf(face), face.loadCalls])).toEqual([
      ['https://fonts.gstatic.example/c-latin.woff2', 1],
      ['https://fonts.gstatic.example/c-greek.woff2', 0],
      ['https://app.example/b-latin.woff2', 1],
      ['https://app.example/b-jp.woff2', 0],
    ]);
    expect(loaded.substitutions.get('calibri')).toEqual({ family: 'Carlito' });
    expect(loaded.regularAliases.get('carlito regular')).toBe('Carlito');
  });
});

describe('FontSourceSession: stylesheet URLs', () => {
  it('resolves relative font URLs against the final stylesheet URL after a redirect', async () => {
    const { calls } = installFetch({
      'https://fonts.internal.example:8443/kit.css': {
        url: 'https://cdn.internal.example/v7/kit/kit.css',
        css: rule('Brand', '../files/brand.woff2') + rule('Brand Serif', '/abs/serif.woff2'),
      },
    });
    const set = fakeSet();
    await open([fontCss({ url: 'https://fonts.internal.example:8443/kit.css' })], set)
      .ensure(demand(['Brand', 'Brand Serif']));
    expect(set.faces.map(urlOf)).toEqual([
      'https://cdn.internal.example/v7/files/brand.woff2', 'https://cdn.internal.example/abs/serif.woff2',
    ]);
    expect(calls).toHaveLength(1);
    expect(calls[0].init?.credentials).toBe('same-origin');
  });
});

describe('FontSourceSession: credentials omit', () => {
  it('downloads the claimed faces of every source kind without credentials', async () => {
    const { calls } = installFetch({
      'https://fonts.internal.example/kit.css': { css: rule('Brand', 'brand.woff2') },
      'https://fonts.internal.example/brand.woff2': { bytes: 12 },
      'https://fonts.internal.example/css?family=Api': { css: rule('Api', 'api.woff2') },
      'https://fonts.internal.example/api.woff2': { bytes: 10 },
      'https://app.example/file.woff2': { bytes: 8 },
    });
    const set = fakeSet();
    const loaded = await open([
      fontCss({ url: 'https://fonts.internal.example/kit.css', credentials: 'omit' }),
      fontCssApi({ urlTemplate: 'https://fonts.internal.example/css?family={family}', families: { Api: true }, credentials: 'omit' }),
      fontFiles({ faces: [{ family: 'File', url: 'https://app.example/file.woff2' }], credentials: 'omit' }),
    ], set).ensure(demand(['Brand', 'Api', 'File']));
    expect(loaded.faces.map((face) => face.family)).toEqual(['Brand', 'Api', 'File']);
    // Bytes, not URLs: FontFace never fetches with its own credentials.
    expect(set.faces.map((face) => (face.source as ArrayBuffer).byteLength)).toEqual([12, 10, 8]);
    expect(calls).toHaveLength(5);
    for (const call of calls) expect(call.init?.credentials).toBe('omit');
    expect(retainedFontSourceBytes()).toBe(30);
  });

  it('keeps the default to the browser, and fails only the faces whose download fails', async () => {
    const { calls } = installFetch({ 'https://app.example/a.woff2': { bytes: 4 } });
    const lazy = fakeSet();
    await open([fontFiles({ faces: [{ family: 'File', url: 'https://app.example/a.woff2' }], credentials: 'same-origin' })], lazy)
      .ensure(demand(['File']));
    expect(lazy.faces.map(urlOf)).toEqual(['https://app.example/a.woff2']);
    expect(calls).toHaveLength(0);

    const set = fakeSet();
    const loaded = await open([fontFiles({
      faces: [
        { family: 'File', url: 'https://app.example/a.woff2' },
        { family: 'File', url: 'https://app.example/missing.woff2', weight: '700' },
        { family: 'Gone', url: 'https://app.example/missing.woff2' },
      ],
      credentials: 'omit',
    })], set).ensure(demand(['File', 'Gone']));
    expect(loaded.faces.map((face) => [face.family, face.weight])).toEqual([['File', 'normal']]);
    // One request per URL.
    expect(calls.map((call) => call.url)).toEqual(['https://app.example/a.woff2', 'https://app.example/missing.woff2']);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('File (HTTP 404), Gone (HTTP 404)');
  });

  it('rejects a downloaded face beyond the face ceiling', async () => {
    installFetch({ 'https://app.example/huge.woff2': { bytes: HARD_MAX_FONT_SOURCE_FACE_BYTES + 1 } });
    const session = open([fontFiles({ faces: [{ family: 'Huge', url: 'https://app.example/huge.woff2' }], credentials: 'omit' })], fakeSet());
    await expect(session.ensure(demand(['Huge']))).rejects.toMatchObject({ code: 'font-source-limit', violation: { quota: 'face' } });
  });
});

describe('FontSourceSession: isolation and sharing', () => {
  const MIRROR = 'https://fonts.mirror.example';
  const INTERNAL = 'https://fonts.internal.example:8443';
  const carlitoAt = (origin: string) =>
    `${origin}/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap`;

  it('keeps two origins apart and never falls back to another origin', async () => {
    const { calls } = installFetch({
      [carlitoAt(MIRROR)]: { status: 503 },
      [carlitoAt(INTERNAL)]: { css: rule('Carlito', 'carlito.woff2') },
    });
    const mirrorSet = fakeSet();
    const internalSet = fakeSet();
    const [mirror, internal] = await Promise.all([
      open([googleFonts({ cssOrigin: MIRROR })], mirrorSet).ensure(demand(['Calibri'])),
      open([googleFonts({ cssOrigin: INTERNAL })], internalSet).ensure(demand(['Calibri'])),
    ]);
    expect(calls.map((call) => new URL(call.url).origin).sort()).toEqual([INTERNAL, MIRROR]);
    expect(mirror.faces).toEqual([]);
    expect(mirror.substitutions.size).toBe(0);
    expect(internal.faces.map((face) => urlOf(face.face as unknown as FakeFontFace)))
      .toEqual([`${INTERNAL}/carlito.woff2`]);
    expect(warn).toHaveBeenCalledTimes(1);
    expect(String(warn.mock.calls[0][0])).toContain('Calibri (HTTP 503)');
  });

  it('shares one stylesheet fetch and one face per realm across documents', async () => {
    const { fetchMock } = installFetch({ 'https://fonts.example/a.css': { css: rule('Brand', 'brand.woff2') } });
    const set = fakeSet();
    // Equal configurations built separately share one identity.
    const a = open([fontCss({ url: 'https://fonts.example/a.css' })], set);
    const b = open([fontCss({ url: 'https://fonts.example/a.css' })], set);
    const [loadedA, loadedB] = await Promise.all([a.ensure(demand(['Brand'])), b.ensure(demand(['Brand']))]);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(set.faces).toHaveLength(1);
    expect(loadedA.faces[0].face).toBe(loadedB.faces[0].face);

    a.release();
    expect(set.faces).toHaveLength(1);
    b.release();
    expect(set.faces).toHaveLength(0);
    expect(liveSourceCacheCount()).toBe(0);
  });

  it('gives each FontFaceSet its own registration', async () => {
    installFetch({ 'https://fonts.example/a.css': { css: rule('Brand', 'brand.woff2') } });
    const main = fakeSet();
    const popup = fakeSet();
    const source = fontCss({ url: 'https://fonts.example/a.css' });
    await open([source], main).ensure(demand(['Brand']));
    await open([source], popup).ensure(demand(['Brand']));
    expect(main.faces).toHaveLength(1);
    expect(popup.faces).toHaveLength(1);
    expect(main.faces[0]).not.toBe(popup.faces[0]);
  });

  // #1497 regressions. Covered elsewhere, not repeated here: request URL
  // composition per origin (`presets/google-fonts.test.ts`), relative asset
  // URLs behind a redirect ('FontSourceSession: stylesheet URLs' above and
  // `css-source.test.ts`), a failing origin keeping host fallback with no
  // request elsewhere (the first test of this block) and failed families never
  // retrying a later source ('claim-once chaining' below).
  it('sends every stylesheet request of a custom-origin preset to that origin only (#1497)', async () => {
    const CN = 'https://fonts.googleapis.cn';
    const { calls } = installFetch({
      [carlitoAt(CN)]: { css: rule('Carlito', 'https://fonts.gstatic.cn/s/carlito/latin.woff2', LATIN) },
      [`${CN}/css2?family=Noto+Sans+JP:wght@400;700&display=swap`]: {
        css: rule('Noto Sans JP', 'https://fonts.gstatic.cn/s/notosansjp/kana.woff2', KANA),
      },
    });
    const set = fakeSet();
    const loaded = await open([googleFonts({ cssOrigin: CN })], set)
      .ensure(demand(['Calibri'], [0x61, 0x3042], ['Noto Sans JP']));
    expect(calls).toHaveLength(2);
    expect(calls.every((call) => new URL(call.url).origin === CN)).toBe(true);
    const contacted = [...calls.map((call) => call.url), ...set.faces.map(urlOf)];
    expect(contacted.filter((url) => /fonts\.(googleapis|gstatic)\.com/.test(String(url)))).toEqual([]);
    expect(loaded.faces.map((face) => face.family)).toEqual(['Carlito', 'Noto Sans JP']);
    expect(loaded.substitutions.get('calibri')).toEqual({ family: 'Carlito' });
  });

  it('keeps two concurrent documents on two origins apart within one realm (#1497)', async () => {
    const OTHER = 'https://fonts.other-mirror.example';
    const { calls } = installFetch({
      [carlitoAt(MIRROR)]: { css: rule('Carlito', 'carlito.woff2') },
      [carlitoAt(OTHER)]: { css: rule('Carlito', 'carlito.woff2') },
    });
    const set = fakeSet();
    const first = open([googleFonts({ cssOrigin: MIRROR })], set);
    const second = open([googleFonts({ cssOrigin: OTHER })], set);
    const [loadedFirst, loadedSecond] = await Promise.all([
      first.ensure(demand(['Calibri'])), second.ensure(demand(['Calibri'])),
    ]);
    // One stylesheet request and one registered face per origin, although the
    // family and descriptors are identical.
    expect(calls.map((call) => call.url).sort()).toEqual([carlitoAt(MIRROR), carlitoAt(OTHER)].sort());
    expect(set.faces.map(urlOf).sort()).toEqual([`${MIRROR}/carlito.woff2`, `${OTHER}/carlito.woff2`]);
    expect(urlOf(loadedFirst.faces[0].face as unknown as FakeFontFace)).toBe(`${MIRROR}/carlito.woff2`);
    expect(urlOf(loadedSecond.faces[0].face as unknown as FakeFontFace)).toBe(`${OTHER}/carlito.woff2`);
    // Releasing one document leaves the other's registration alone.
    first.release();
    expect(set.faces.map(urlOf)).toEqual([`${OTHER}/carlito.woff2`]);
    expect(warn).not.toHaveBeenCalled();
  });

  it('does not cache failures: a later document retries the same source', async () => {
    const routes: Record<string, Route> = { 'https://fonts.example/a.css': { status: 500 } };
    const { fetchMock } = installFetch(routes);
    const source = fontCss({ url: 'https://fonts.example/a.css' });
    const first = open([source], fakeSet());
    expect((await first.ensure(demand(['Brand']))).faces).toEqual([]);
    routes['https://fonts.example/a.css'] = { css: rule('Brand', 'brand.woff2') };
    // Same session: the family is final (no second network wave).
    await first.ensure(demand(['Brand']));
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect((await open([source], fakeSet()).ensure(demand(['Brand']))).faces).toHaveLength(1);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
});

describe('FontSourceSession: claim-once chaining', () => {
  it('lets the first source that declares a family claim it, and passes on the rest', async () => {
    const { calls } = installFetch({
      'https://fonts.example/all.css': { css: rule('Carlito', 'css-carlito.woff2') + rule('Other', 'other.woff2') },
    });
    const set = fakeSet();
    const loaded = await open([
      fontFiles({ faces: [{ family: 'Carlito', url: 'https://app.example/carlito.woff2' }] }),
      fontCss({ url: 'https://fonts.example/all.css', substitutes: OFFICE_FONT_SUBSTITUTIONS }),
    ], set).ensure(demand(['Calibri', 'Carlito', 'Other']));
    // Calibri: neither source serves it directly; the files source has no
    // substitutes, so the stylesheet claims it through Calibri → Carlito.
    expect(set.faces.map(urlOf)).toEqual([
      'https://fonts.example/css-carlito.woff2', 'https://app.example/carlito.woff2', 'https://fonts.example/other.woff2',
    ]);
    expect(loaded.faces.map((face) => face.sourceIndex)).toEqual([1, 0, 1]);
    expect(loaded.substitutions.get('calibri')).toEqual({ family: 'Carlito' });
    expect(calls).toHaveLength(1);
  });

  it('never consults a later source for a family whose source failed', async () => {
    const { calls } = installFetch({ 'https://fonts.example/b.css': { css: rule('Brand', 'b.woff2') } });
    const loaded = await open([
      fontCss({ url: 'https://fonts.example/missing.css' }),
      fontCss({ url: 'https://fonts.example/b.css' }),
    ], fakeSet()).ensure(demand(['Brand']));
    expect(calls.map((call) => call.url)).toEqual(['https://fonts.example/missing.css']);
    expect(loaded.faces).toEqual([]);
  });

  it('never re-enters the chain when a claimed face fails to load', async () => {
    FakeFontFace.loadImpl = (face) => urlOf(face)?.includes('broken') ? Promise.reject(new Error('net')) : Promise.resolve();
    const { calls } = installFetch({ 'https://fonts.example/b.css': { css: rule('Brand', 'b.woff2') } });
    const session = open([
      fontFiles({ faces: [{ family: 'Brand', url: 'https://app.example/broken.woff2' }] }),
      fontCss({ url: 'https://fonts.example/b.css' }),
    ], fakeSet());
    const loaded = await session.ensure(demand(['Brand']));
    expect(loaded.faces).toEqual([]);
    expect(calls).toHaveLength(0);
    expect(warn).toHaveBeenCalledTimes(1);
    // One warning per family per session.
    await session.ensure(demand(['Brand'], [0x41]));
    expect(warn).toHaveBeenCalledTimes(1);
  });

  it('chains a resolver before a CSS API and keeps resolver failures final', async () => {
    const carlito = 'https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&display=swap';
    const caladea = 'https://fonts.googleapis.com/css2?family=Caladea:ital,wght@0,400;0,700;1,400;1,700&display=swap';
    const { calls } = installFetch({
      [carlito]: { css: rule('Carlito', 'carlito.woff2') },
      [caladea]: { css: rule('Caladea', 'caladea.woff2') },
    });
    const seen: FontDemand[] = [];
    const resolve: FontResolveFn = (request, { signal }) => {
      seen.push(request);
      expect(signal).toBeInstanceOf(AbortSignal);
      return request.families.includes('Private') ? [{ family: 'Private', data: new Uint8Array([1, 2, 3]).buffer }] : [];
    };
    const set = fakeSet();
    const loaded = await open([fontResolver(resolve), googleFonts()], set)
      .ensure(demand(['Private', 'Calibri'], [0x62, 0x61, 0x62], ['Noto Sans JP']));
    expect(seen).toEqual([{
      families: ['Private', 'Calibri'], fallbackFamilies: ['Noto Sans JP'],
    }]);
    expect(loaded.faces.map((face) => face.family)).toEqual(['Private', 'Carlito']);
    // The cached resolution and the registered face are both real copies.
    expect(retainedFontSourceBytes()).toBe(6);
    // Noto Sans JP passed through the resolver to the preset.
    expect(calls.map((call) => new URL(call.url).searchParams.get('family')))
      .toEqual(['Carlito:ital,wght@0,400;0,700;1,400;1,700', 'Noto Sans JP:wght@400;700']);

    const failing = fontResolver(() => { throw new Error('token expired'); });
    const failed = await open([failing, googleFonts()], fakeSet()).ensure(demand(['Cambria']));
    expect(failed.faces).toEqual([]);
    expect(calls.some((call) => call.url === caladea)).toBe(false);
    expect(String(warn.mock.calls.at(-1)?.[0])).toContain('Cambria (token expired)');
  });
});

describe('FontSourceSession: limits fail closed', () => {
  function expectLimit(error: unknown, quota: string) {
    expect(error).toBeInstanceOf(OoxmlError);
    expect(error).toMatchObject({ code: 'font-source-limit', violation: { quota } });
  }

  it('rejects more demanded families than the ceiling, then stays failed', async () => {
    installFetch({});
    const set = fakeSet();
    const session = open([fontFiles({ faces: [{ family: 'F0', url: 'https://app.example/f0.woff2' }] })], set);
    await session.ensure(demand(['F0']));
    expect(set.faces).toHaveLength(1);
    const many = Array.from({ length: HARD_MAX_FONT_SOURCE_FAMILIES + 1 }, (_, index) => `F${index}`);
    const error = await session.ensure(demand(many)).catch((caught: unknown) => caught);
    expectLimit(error, 'families');
    // The session released everything and rethrows the same breach.
    expect(set.faces).toHaveLength(0);
    await expect(session.ensure(demand(['F0']))).rejects.toBe(error);
  });

  it('rejects more declarations per session than the ceiling', async () => {
    const faces = (family: string) => Array.from({ length: 1100 }, (_, index) => ({
      family, url: `https://app.example/${family}-${index}.woff2`, unicodeRange: `U+${(index + 0x100).toString(16)}`,
    }));
    const set = fakeSet();
    const session = open([fontFiles({ faces: faces('A') }), fontFiles({ faces: faces('B') })], set);
    expectLimit(await session.ensure(demand(['A', 'B'], [])).catch((caught: unknown) => caught), 'declarations');
    expect(set.faces).toHaveLength(0);
  });

  it('rejects a stylesheet beyond the byte ceiling', async () => {
    installFetch({ 'https://fonts.example/huge.css': { css: `/*${'x'.repeat(1024 * 1024)}*/` } });
    const error = await open([fontCss({ url: 'https://fonts.example/huge.css' })], fakeSet())
      .ensure(demand(['Brand'])).catch((caught: unknown) => caught);
    expectLimit(error, 'stylesheet');
  });

  it('charges data bytes once per realm and rejects beyond the retained ceiling', async () => {
    vi.resetModules();
    vi.doMock('../worker/resource-policy.generated.js', async (importOriginal) => ({
      ...await importOriginal<Record<string, unknown>>(),
      HARD_MAX_FONT_SOURCE_RETAINED_BYTES: 100,
    }));
    try {
      const fresh = await import('./source-session.js');
      const sourceModule = await import('./source.js');
      const bytes = (length: number) => new Uint8Array(length).buffer;
      const shared = sourceModule.fontFiles({ faces: [{ family: 'Shared', data: bytes(60) }] });
      const set = fakeSet();
      const a = new fresh.FontSourceSession({ sources: [shared], fontSet: asSet(set) });
      const b = new fresh.FontSourceSession({ sources: [shared], fontSet: asSet(set) });
      await a.ensure(demand(['Shared']));
      await b.ensure(demand(['Shared']));
      expect(fresh.retainedFontSourceBytes()).toBe(60);
      const over = new fresh.FontSourceSession({
        sources: [sourceModule.fontFiles({ faces: [{ family: 'Big', url: `data:font/ttf;base64,${'AAAA'.repeat(14)}` }] })],
        fontSet: asSet(set),
      });
      const error = await over.ensure(demand(['Big'])).catch((caught: unknown) => caught);
      expect(error).toMatchObject({ code: 'font-source-limit', violation: { quota: 'retained', observed: 102 } });
      a.release();
      b.release();
      expect(fresh.retainedFontSourceBytes()).toBe(0);
      expect(set.faces).toHaveLength(0);
    } finally {
      vi.doUnmock('../worker/resource-policy.generated.js');
      vi.resetModules();
    }
  });
});

describe('FontSourceSession: retained bytes held by shared caches', () => {
  it('charges cached resolver bytes while the requesting session holds them', async () => {
    vi.resetModules();
    vi.doMock('../worker/resource-policy.generated.js', async (importOriginal) => ({
      ...await importOriginal<Record<string, unknown>>(),
      HARD_MAX_FONT_SOURCE_RETAINED_BYTES: 100,
    }));
    try {
      const fresh = await import('./source-session.js');
      const sourceModule = await import('./source.js');
      const source = sourceModule.fontResolver((request) =>
        request.families.map((family) => ({ family, data: new Uint8Array(30).buffer })));
      const set = fakeSet();
      const idle = new fresh.FontSourceSession({ sources: [source], fontSet: asSet(set) });
      const cycle = async (family: string) => {
        const session = new fresh.FontSourceSession({ sources: [source], fontSet: asSet(set) });
        try { await session.ensure(demand([family])); } finally { session.release(); }
      };
      // A released session's cached resolutions go with it, even while an
      // idle session keeps the source's cache alive.
      await cycle('A');
      await cycle('B');
      expect(fresh.retainedFontSourceBytes()).toBe(0);
      // While held, the cached copy and the registered copy both count:
      // 30 + 30 for one open session; a second one exceeds 100.
      const open1 = new fresh.FontSourceSession({ sources: [source], fontSet: asSet(set) });
      await open1.ensure(demand(['C']));
      expect(fresh.retainedFontSourceBytes()).toBe(60);
      const open2 = new fresh.FontSourceSession({ sources: [source], fontSet: asSet(set) });
      await expect(open2.ensure(demand(['D']))).rejects.toMatchObject({ code: 'font-source-limit', violation: { quota: 'retained' } });
      open1.release();
      idle.release();
      expect(fresh.retainedFontSourceBytes()).toBe(0);
    } finally {
      vi.doUnmock('../worker/resource-policy.generated.js');
      vi.resetModules();
    }
  });
});

describe('FontSourceSession: caches keep only accounted bytes', () => {
  it('lets unclaimed resolver bytes of a cached batch be collected', async () => {
    setFlagsFromString('--expose-gc');
    const gc = runInNewContext('gc') as () => void;
    vi.resetModules();
    vi.doMock('../worker/resource-policy.generated.js', async (importOriginal) => ({
      ...await importOriginal<Record<string, unknown>>(),
      HARD_MAX_FONT_SOURCE_RETAINED_BYTES: 100,
    }));
    try {
      const fresh = await import('./source-session.js');
      const sourceModule = await import('./source.js');
      let live = 0;
      const registry = new FinalizationRegistry<number>((bytes) => { live -= bytes; });
      // Answers every demanded family except "Absent*" with 30 bytes.
      const source = sourceModule.fontResolver((request) => request.families
        .filter((family) => !family.startsWith('Absent'))
        .map((family) => {
          const data = new ArrayBuffer(30);
          registry.register(data, 30);
          live += 30;
          return { family, data };
        }));
      const set = fakeSet();
      const idle = new fresh.FontSourceSession({ sources: [source], fontSet: asSet(set) });
      for (let batch = 0; batch < 5; batch++) {
        const session = new fresh.FontSourceSession({ sources: [source], fontSet: asSet(set) });
        // The cached absent family must not pin its batch (and the claimed
        // faces the ceiling rejected).
        await session.ensure(demand([`Absent${batch}`, ...[0, 1, 2, 3].map((index) => `F${batch}-${index}`)]))
          .catch(() => undefined);
        session.release();
      }
      const accounted = fresh.retainedFontSourceBytes();
      expect(accounted).toBeLessThanOrEqual(100);
      for (let round = 0; round < 50 && live > accounted; round++) {
        gc();
        await new Promise((resolve) => setTimeout(resolve, 5));
      }
      expect(live).toBe(accounted);
      idle.release();
    } finally {
      vi.doUnmock('../worker/resource-policy.generated.js');
      vi.resetModules();
    }
  });
});

describe('FontSourceSession: charge and refund pair per cache entry', () => {
  it('refunds a charge whose task was abandoned before it settled, across retries', async () => {
    const source = fontResolver((request) => request.families.map((family) => ({ family, data: new Uint8Array(10).buffer })));
    const set = fakeSet();
    const idle = open([source], set);
    for (let ticks = 0; ticks < 40; ticks++) {
      const family = `F${ticks}`;
      const early = open([source], set);
      const pending = early.ensure(demand([family])).catch(() => undefined);
      for (let tick = 0; tick < ticks; tick++) await Promise.resolve();
      early.release();
      await pending;
      // A retry by another document for the same family.
      const retry = open([source], set);
      await retry.ensure(demand([family]));
      retry.release();
    }
    idle.release();
    expect(retainedFontSourceBytes()).toBe(0);
  });
});

describe('FontSourceSession: registration order', () => {
  it('keeps the last occurrence of a repeated declaration, as CSS precedence does', async () => {
    const a = { family: 'Brand', url: 'https://app.example/a.woff2' };
    const b = { family: 'Brand', url: 'https://app.example/b.woff2' };
    const set = fakeSet();
    await open([fontFiles({ faces: [a, b, a] })], set).ensure(demand(['Brand']));
    expect(set.faces.map(urlOf)).toEqual(['https://app.example/b.woff2', 'https://app.example/a.woff2']);
  });
});

describe('FontSourceSession: released sessions keep no FontFace alive', () => {
  it('lets every face of a released session be collected without another ensure', async () => {
    setFlagsFromString('--expose-gc');
    const gc = runInNewContext('gc') as () => void;
    let live = 0;
    const registry = new FinalizationRegistry<undefined>(() => { live--; });
    class TrackedFontFace extends FakeFontFace {
      constructor(family: string, source: string | ArrayBuffer, descriptors?: FontFaceDescriptors) {
        super(family, source, descriptors);
        registry.register(this, undefined);
        live++;
      }
    }
    G.FontFace = TrackedFontFace;
    const set = fakeSet();
    const kept: FontSourceSession[] = [];
    for (let index = 0; index < 5; index++) {
      const session = new FontSourceSession({
        sources: [fontFiles({ faces: [{ family: `F${index}`, url: `https://app.example/f${index}.woff2` }] })],
        fontSet: asSet(set),
      });
      await session.ensure(demand([`F${index}`]));
      session.release();
      kept.push(session);
    }
    expect(set.faces).toHaveLength(0);
    for (let round = 0; round < 50 && live > 0; round++) {
      gc();
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
    expect(live).toBe(0);
    expect(kept).toHaveLength(5);
  });
});

describe('FontSourceSession: shared resolver work survives a released waiter', () => {
  it.each([0, 1, 2, 3])('serves the surviving session when the creating one is released %i microtasks in', async (ticks) => {
    let calls = 0;
    const source = fontResolver((request) => {
      calls++;
      return request.families.map((family) => ({ family, url: 'https://app.example/shared.woff2' }));
    });
    const set = fakeSet();
    const first = open([source], set);
    const second = open([source], set);
    const released = first.ensure(demand(['Shared']));
    const surviving = second.ensure(demand(['Shared']));
    for (let tick = 0; tick < ticks; tick++) await Promise.resolve();
    first.release();
    expect((await released).faces).toEqual([]);
    expect((await surviving).faces.map((face) => face.family)).toEqual(['Shared']);
    expect(set.faces).toHaveLength(1);
    expect(calls).toBe(1);
  });
});

describe('FontSourceSession: cached resolutions are bounded by live requesters', () => {
  it.each([
    ['font-css', () => fontCss({ url: 'https://fonts.example/kit.css' })],
    ['font-css-api allow-list', () => fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}', families: { Brand: true } })],
  ] as const)('evicts what released sessions requested while another session holds the source (%s)', async (_kind, make) => {
    installFetch({
      'https://fonts.example/kit.css': { css: rule('Brand', 'brand.woff2') },
      'https://fonts.example/css?family=Brand': { css: rule('Brand', 'brand.woff2') },
    });
    const source = make();
    const set = fakeSet();
    const idle = open([source], set);
    const missing = (batch: number) => Array.from({ length: HARD_MAX_FONT_SOURCE_FAMILIES - 1 }, (_, index) => `Missing ${batch}-${index}`);
    for (let batch = 0; batch < 100; batch++) {
      const session = open([source], set);
      await session.ensure(demand(['Brand', ...missing(batch)], [0x41]));
      // While open: its families (present and absent) plus at most one stylesheet.
      expect(sourceCacheEntryCount()).toBeLessThanOrEqual(HARD_MAX_FONT_SOURCE_FAMILIES + 1);
      session.release();
      expect(sourceCacheEntryCount()).toBe(0);
    }
    // A session still open keeps exactly what it requested.
    await idle.ensure(demand(['Brand', 'Missing']));
    expect(sourceCacheEntryCount()).toBe(3);
    idle.release();
    expect(sourceCacheEntryCount()).toBe(0);
  });
});

describe('FontSourceSession: released sessions keep no source bytes', () => {
  it('drops its source descriptors (and their buffers) on release', async () => {
    setFlagsFromString('--expose-gc');
    const gc = runInNewContext('gc') as () => void;
    let live = 0;
    const registry = new FinalizationRegistry<undefined>(() => { live--; });
    const set = fakeSet();
    const kept: FontSourceSession[] = [];
    // Each descriptor is only reachable through its session.
    const openInline = async (index: number) => {
      const data = new ArrayBuffer(1024 * 1024);
      registry.register(data, undefined);
      live++;
      const session = new FontSourceSession({
        sources: [fontFiles({ faces: [{ family: `Inline${index}`, data }] })], fontSet: asSet(set),
      });
      await session.ensure(demand([`Inline${index}`]));
      session.release();
      return session;
    };
    for (let index = 0; index < 5; index++) kept.push(await openInline(index));
    expect(retainedFontSourceBytes()).toBe(0);
    expect(set.faces).toHaveLength(0);
    for (let round = 0; round < 50 && live > 0; round++) {
      gc();
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
    expect(live).toBe(0);
    expect(kept).toHaveLength(5);
  });
});

describe('FontSourceSession: released sessions keep no font content', () => {
  it('drops the stylesheet family index with the last release, even while session objects live on', async () => {
    setFlagsFromString('--expose-gc');
    const gc = runInNewContext('gc') as () => void;
    let live = 0;
    const registry = new FinalizationRegistry<undefined>(() => { live--; });
    vi.resetModules();
    vi.doMock('./files-source.js', async (importOriginal) => {
      const original = await importOriginal<typeof import('./files-source.js')>();
      return {
        ...original,
        indexFontDeclarations: (declarations: readonly FontFaceDeclaration[]) => {
          for (const declaration of declarations) { registry.register(declaration, undefined); live++; }
          return original.indexFontDeclarations(declarations);
        },
      };
    });
    try {
      const fresh = await import('./source-session.js');
      const sourceModule = await import('./source.js');
      const dataFace = `data:font/woff2;base64,${'AAAA'.repeat(10)}`;
      const routes: Record<string, Route> = {};
      for (let index = 0; index < 5; index++) {
        routes[`https://fonts.example/${index}.css`] = { css: rule('Inline', `"${dataFace}"`) };
      }
      installFetch(routes);
      const kept: InstanceType<typeof fresh.FontSourceSession>[] = [];
      for (let index = 0; index < 5; index++) {
        const session = new fresh.FontSourceSession({
          sources: [sourceModule.fontCss({ url: `https://fonts.example/${index}.css` })], fontSet: asSet(fakeSet()),
        });
        await session.ensure(demand(['Inline']));
        session.release();
        kept.push(session);
      }
      expect(live).toBe(5);
      expect(fresh.retainedFontSourceBytes()).toBe(0);
      for (let round = 0; round < 50 && live > 0; round++) {
        gc();
        await new Promise((resolve) => setTimeout(resolve, 5));
      }
      expect(live).toBe(0);
      expect(kept).toHaveLength(5);
    } finally {
      vi.doUnmock('./files-source.js');
      vi.resetModules();
    }
  });
});

describe('FontSourceSession: release balance and cancellation', () => {
  it('accepts an already-aborted signal: inert, cache released, ensures reject', async () => {
    const { fetchMock } = installFetch({});
    const session = new FontSourceSession({
      sources: [fontCss({ url: 'https://fonts.example/a.css' })], fontSet: asSet(fakeSet()), signal: AbortSignal.abort(),
    });
    expect(liveSourceCacheCount()).toBe(0);
    await expect(session.ensure(demand(['Brand']))).rejects.toMatchObject({ name: 'AbortError' });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it.each([0, 1, 2, 3, 4])('starts no live fetch when released %i microtasks into an ensure', async (ticks) => {
    const { calls } = installFetch({ 'https://fonts.example/a.css': { css: rule('Brand', 'b.woff2') } });
    const set = fakeSet();
    const session = open([fontCss({ url: 'https://fonts.example/a.css' })], set);
    const pending = session.ensure(demand(['Brand']));
    for (let tick = 0; tick < ticks; tick++) await Promise.resolve();
    session.release();
    expect((await pending).faces).toEqual([]);
    // A request may only run while a waiter wants it.
    expect(calls.filter((call) => !call.init?.signal?.aborted)).toEqual([]);
    expect(set.faces).toHaveLength(0);
  });

  it('releases exactly what it retained, once, including faces two families share', async () => {
    installFetch({});
    const set = fakeSet();
    const other = open([fontFiles({ faces: [{ family: 'Carlito', url: 'https://app.example/c.woff2' }] })], set);
    await other.ensure(demand(['Carlito']));
    const session = open([fontFiles({
      faces: [{ family: 'Carlito', url: 'https://app.example/c.woff2' }],
      substitutes: OFFICE_FONT_SUBSTITUTIONS,
    })], set);
    const loaded = await session.ensure(demand(['Calibri', 'Carlito']));
    expect(loaded.faces).toHaveLength(1);
    // Different source identities never share a registration.
    expect(set.faces).toHaveLength(2);
    session.release();
    session.release();
    expect(set.faces).toHaveLength(1);
    other.release();
    expect(set.faces).toHaveLength(0);
    expect(await session.ensure(demand(['Calibri']))).toEqual(expect.objectContaining({ faces: [] }));
  });

  it('retains nothing when released while a stylesheet is in flight', async () => {
    let open_!: () => void;
    const gate = new Promise<void>((resolve) => { open_ = resolve; });
    const { calls } = installFetch({ 'https://fonts.example/a.css': { css: rule('Brand', 'b.woff2'), gate } });
    const set = fakeSet();
    const session = open([fontCss({ url: 'https://fonts.example/a.css' })], set);
    const pending = session.ensure(demand(['Brand']));
    await vi.waitFor(() => expect(calls).toHaveLength(1));
    session.release();
    expect((await pending).faces).toEqual([]);
    // The sole waiter left, so the shared fetch was aborted.
    expect(calls[0].init?.signal?.aborted).toBe(true);
    open_();
    expect(set.faces).toHaveLength(0);
  });

  it('rejects with AbortError on caller abort without aborting a fetch another document joined', async () => {
    let open_!: () => void;
    const gate = new Promise<void>((resolve) => { open_ = resolve; });
    const { calls } = installFetch({ 'https://fonts.example/a.css': { css: rule('Brand', 'b.woff2'), gate } });
    const set = fakeSet();
    const controller = new AbortController();
    const source = fontCss({ url: 'https://fonts.example/a.css' });
    const aborted = open([source], set, controller.signal).ensure(demand(['Brand']));
    const survivor = open([source], set).ensure(demand(['Brand']));
    await vi.waitFor(() => expect(calls).toHaveLength(1));
    controller.abort();
    await expect(aborted).rejects.toMatchObject({ name: 'AbortError' });
    expect(calls[0].init?.signal?.aborted).toBe(false);
    open_();
    expect((await survivor).faces).toHaveLength(1);
    expect(set.faces).toHaveLength(1);
  });

  it('gives up after the ceiling, warns, and aborts the abandoned fetch', async () => {
    vi.useFakeTimers();
    const { calls } = installFetch({ 'https://fonts.example/a.css': { gate: new Promise(() => undefined) } });
    const session = open([fontCss({ url: 'https://fonts.example/a.css' })], fakeSet());
    let settled: unknown;
    void session.ensure(demand(['Brand'])).then((value) => { settled = value; });
    await vi.advanceTimersByTimeAsync(14_999);
    expect(settled).toBeUndefined();
    await vi.advanceTimersByTimeAsync(1);
    expect(settled).toMatchObject({ faces: [] });
    expect(calls[0].init?.signal?.aborted).toBe(true);
    expect(String(warn.mock.calls[0][0])).toContain('Brand (timed out)');
  });

  it('passes an abort signal to a resolver and drops its late answer after release', async () => {
    let finish!: (faces: readonly FontFaceDeclaration[]) => void;
    let received: AbortSignal | undefined;
    const source = fontResolver((_demand, { signal }) => {
      received = signal;
      return new Promise((resolve) => { finish = resolve; });
    });
    const set = fakeSet();
    const session = open([source], set);
    const pending = session.ensure(demand(['Private']));
    await vi.waitFor(() => expect(received).toBeDefined());
    session.release();
    expect(received?.aborted).toBe(true);
    finish([{ family: 'Private', data: new Uint8Array(4).buffer }]);
    expect((await pending).faces).toEqual([]);
    expect(set.faces).toHaveLength(0);
  });
});

describe('FontSourceSession: demand contract', () => {
  it('rejects malformed demand and treats an oversized scalar list as all', async () => {
    const session = open([fontFiles({ faces: [{ family: 'A', url: 'https://app.example/a.woff2', unicodeRange: 'U+10FFFF' }] })], fakeSet());
    await expect(session.ensure(demand(['A'], [0x110000]))).rejects.toBeInstanceOf(TypeError);
    await expect(session.ensure({ ...demand(['A']), families: 'A' as unknown as string[] })).rejects.toBeInstanceOf(TypeError);
    const many = Array.from({ length: 70_000 }, (_, index) => index);
    expect((await session.ensure(demand(['A'], many))).faces).toHaveLength(1);
  });

  it('is inert without a FontFaceSet (Node without a shim)', async () => {
    const fetchMock = installFetch({}).fetchMock;
    const session = new FontSourceSession({ sources: [fontCssApi({ urlTemplate: 'https://fonts.example/css?family={family}' })], fontSet: null });
    expect((await session.ensure(demand(['Calibri']))).faces).toEqual([]);
    expect(fetchMock).not.toHaveBeenCalled();
    session.release();
  });
});

describe('FontSourceSession: script-scoped substitution proof (FontSessionDemand.substituteScripts)', () => {
  const NASKH = 'https://fonts.googleapis.com/css2?family=Noto+Naskh+Arabic:wght@400;700&display=swap';

  it('consults a scoped substitute only for families whose text proves its script, caching the scopes apart', async () => {
    const { calls } = installFetch({ [NASKH]: { css: rule('Noto Naskh Arabic', 'https://fonts.gstatic.example/n.woff2') } });
    const preset = googleFonts();
    const set = fakeSet();
    const latin = await open([preset], set).ensure({
      ...demand(['Sakkal Majalla'], [0x61]), substituteScripts: { 'sakkal majalla': [] },
    });
    expect(latin.substitutions.size).toBe(0);
    expect(calls).toEqual([]);
    // Same source identity, Arabic proven: the earlier unscoped answer is not reused.
    const arabic = await open([preset], set).ensure({
      ...demand(['Sakkal Majalla'], [0x645]), substituteScripts: { 'sakkal majalla': ['arabic'] },
    });
    expect(arabic.substitutions.get('sakkal majalla')).toEqual({ family: 'Noto Naskh Arabic', script: 'arabic' });
    expect(calls.map(({ url }) => url)).toEqual([NASKH]);
    // Unscoped substitutions and the family's own faces are never restricted.
    const files = fontFiles({ faces: [{ family: 'Sakkal Majalla', url: 'https://app.example/s.woff2' }] });
    const own = await open([files, preset], set).ensure({
      ...demand(['Sakkal Majalla'], [0x61]), substituteScripts: {},
    });
    expect(own.faces.map((face) => face.family)).toEqual(['Sakkal Majalla']);
  });

  it('resolves an absent family again when a later delta widens its proven scope', async () => {
    // Progressive loading: the family's first text (Latin) proves no script,
    // a later slide's Arabic text in the same family proves 'arabic'.
    const files = fontFiles({
      faces: [{ family: 'Served', url: 'https://app.example/served.woff2' }],
      substitutes: [{ for: 'Authored', use: 'Served', script: 'arabic' }],
    });
    const set = fakeSet();
    const session = open([files], set);
    const first = await session.ensure({ ...demand(['Authored'], [0x61]), substituteScripts: { authored: [] } });
    expect(first.substitutions.size).toBe(0);
    expect(set.faces).toHaveLength(0);
    const second = await session.ensure({ ...demand(['Authored'], [0x645]), substituteScripts: { authored: ['arabic'] } });
    expect(second.substitutions.get('authored')).toEqual({ family: 'Served', script: 'arabic' });
    expect(set.faces.map((face) => face.family)).toEqual(['Served']);
    // A repeated demand with the same scope resolves nothing further.
    await session.ensure({ ...demand(['Authored'], [0x646]), substituteScripts: { authored: ['arabic'] } });
    expect(set.faces).toHaveLength(1);
    session.release();
  });
});
