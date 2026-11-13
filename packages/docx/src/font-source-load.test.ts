import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  fontCssApi as publicFontCssApi,
  fontFiles as publicFontFiles,
  fontResolver as publicFontResolver,
  fontSourceDescriptors,
  OFFICE_FONT_SUBSTITUTIONS,
  type FontSource,
} from '@silurus/ooxml-core/internal/font-sources';

// The descriptors behind the opaque factory results (what a load hands the
// internal font-source pipeline).
const descriptor = (source: FontSource) => fontSourceDescriptors([source])[0]!;
const fontCssApi = (...args: Parameters<typeof publicFontCssApi>) => descriptor(publicFontCssApi(...args));
const fontFiles = (...args: Parameters<typeof publicFontFiles>) => descriptor(publicFontFiles(...args));
const fontResolver = (...args: Parameters<typeof publicFontResolver>) => descriptor(publicFontResolver(...args));
import { docxFontFamilyPlan } from './font-demand-plan.js';
import { formatOrdinalNumber } from '@silurus/ooxml-core';
import { ASCII_NUMBER_FORMATS, docxFontDemandScalars, loadDocxFontSources } from './font-source-load.js';
import { LAYOUT_SYNTHESIZED_SCALARS } from './line-layout.js';
import type { DocxDocumentModel } from './types.js';

function docWith(runs: object[], extra: Partial<DocxDocumentModel> = {}): DocxDocumentModel {
  return {
    section: {} as DocxDocumentModel['section'],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    majorFont: 'Calibri',
    minorFont: 'Calibri',
    body: [{ type: 'paragraph', runs } as never],
    ...extra,
  } as DocxDocumentModel;
}
const text = (value: string, props: object = {}) => ({ type: 'text', text: value, ...props });
const scalarsOf = (doc: DocxDocumentModel) => docxFontDemandScalars(doc);

describe('docxFontDemandScalars (design §6)', () => {
  it('collects painted text plus the layout-synthesized scalars, nothing else', () => {
    const scalars = scalarsOf(docWith([text('Hello 日本')]));
    expect(scalars).not.toBe('all');
    const set = new Set(scalars as number[]);
    for (const cp of LAYOUT_SYNTHESIZED_SCALARS) expect(set.has(cp)).toBe(true);
    expect(set.has('日'.codePointAt(0)!)).toBe(true);
    expect(set.has('本'.codePointAt(0)!)).toBe(true);
    expect(set.has('語'.codePointAt(0)!)).toBe(false);
  });

  it('adds what layout paints but the text does not spell', () => {
    const set = (doc: DocxDocumentModel) => new Set(scalarsOf(doc) as number[]);
    // §17.3.2.5 allCaps paints upper case.
    expect(set(docWith([text('α', { allCaps: true })])).has(0x0391)).toBe(true);
    expect(set(docWith([text('α')])).has(0x0391)).toBe(false);
    // §17.3.2.12 emphasis marks paint their own glyph.
    expect(set(docWith([text('漢', { emphasisMark: 'circle' })])).has(0x25cb)).toBe(true);
    // §17.16.4.1 DATE picture literals.
    const date = { type: 'field', fieldType: 'date', instruction: 'DATE \\@ "yyyy\'年\'M\'月\'"', fallbackText: '' };
    const dated = set(docWith([date]));
    expect(dated.has('年'.codePointAt(0)!)).toBe(true);
    expect(dated.has('月'.codePointAt(0)!)).toBe(true);
  });

  it.each([
    ['strong RTL text', docWith([text('مرحبا')])],
    ['an RTL run', docWith([text('abc', { rtl: true })])],
    ['ruby', docWith([text('漢', { ruby: { text: 'かん' } })])],
    ['vertical text', docWith([text('a')], { section: { textDirection: 'tbRl' } as never })],
    ['a chart', docWith([{ type: 'chart', chart: {} }])],
    ['an equation', docWith([{ type: 'math', nodes: [], display: false, fontSize: 10 }])],
    ['a metafile picture', docWith([{ type: 'image', imagePath: 'word/media/image1.emf', mimeType: 'image/emf' }])],
    ['a PUA symbol code', docWith([text('')])],
    ['a non-ASCII page number format', docWith([text('a')], { section: { pageNumType: { fmt: 'hebrew1' } } as never })],
    ['a non-ASCII field switch', docWith([{ type: 'field', fieldType: 'page', instruction: 'PAGE \\* HEBREW1', fallbackText: '1' }])],
  ])('answers all for %s', (_label, doc) => {
    expect(scalarsOf(doc)).toBe('all');
  });

  it('demands only what a symbol font may paint, never every subset', () => {
    // A Symbol bullet (U+F0B7) and a Wingdings run: the authored codes, their
    // F0-shifted forms, their conversion targets and VS15, nothing else.
    const doc = docWith([text('l', { fontFamily: 'Wingdings' })]);
    (doc.body[0] as { numbering?: object }).numbering = { text: '\uF0B7', fontFamily: 'Symbol' };
    const scalars = scalarsOf(doc);
    expect(scalars).not.toBe('all');
    const set = new Set(scalars as number[]);
    for (const cp of [0xf0b7, 0x2022, 0x6c, 0xf06c, 0x25cf, 0xfe0e]) expect(set.has(cp), cp.toString(16)).toBe(true);
    // Bounded by the tables (layout-generated text counts as painted in each
    // symbol family too): a few hundred scalars, not every subset.
    expect(set.size).toBeLessThan(1024);
    // §17.3.2.5 caps paints the upper case, converted in the run's symbol face (Α).
    expect(new Set(scalarsOf(docWith([text('a', { fontFamily: 'Symbol', allCaps: true })])) as number[]).has(0x391)).toBe(true);
    // A generated field result (PAGE \* ROMAN may paint V) is converted too (ς).
    expect(new Set(scalarsOf(docWith([{ type: 'field', fieldType: 'page', instruction: 'PAGE \\* ROMAN',
      fallbackText: 'I', fontFamily: 'Symbol' }])) as number[]).has(0x3c2)).toBe(true);
    // A PUA code outside a symbol family still proves nothing.
    expect(scalarsOf(docWith([text('\uF0B7', { fontFamily: 'Arial' })]))).toBe('all');
  });

  it('treats as ASCII only number formats whose every output is ASCII', () => {
    for (const format of ASCII_NUMBER_FORMATS) {
      for (let n = -2; n <= 5000; n++) expect(formatOrdinalNumber(n, format), format).toMatch(/^[\x20-\x7e]*$/u);
    }
  });

  it('keeps ASCII page number formats lazy', () => {
    expect(scalarsOf(docWith([{ type: 'field', fieldType: 'page', instruction: 'PAGE \\* ROMAN', fallbackText: 'I' }],
      { section: { pageNumType: { fmt: 'lowerLetter' } } as never }))).not.toBe('all');
  });
});

// ---- subset laziness in the painting realm ----------------------------------------

const G = globalThis as Record<string, unknown>;
const ORIGINAL = { FontFace: G.FontFace, fetch: G.fetch };

class FakeFontFace {
  status: FontFaceLoadStatus = 'unloaded';
  readonly weight: string;
  readonly style: string;
  readonly unicodeRange: string;
  constructor(readonly family: string, readonly source: string, descriptors: FontFaceDescriptors = {}) {
    this.weight = descriptors.weight ?? 'normal';
    this.style = descriptors.style ?? 'normal';
    this.unicodeRange = descriptors.unicodeRange ?? 'U+0-10FFFF';
  }
  load(): Promise<this> {
    this.status = 'loaded';
    return Promise.resolve(this);
  }
}

const SUBSETS = { latin: 'U+0-FF', kana: 'U+3000-30FF', kanji: 'U+6000-6FFF' } as const;
const NOTO_CSS = Object.entries(SUBSETS).map(([name, range]) =>
  `@font-face { font-family: 'Noto Sans JP'; font-weight: 400; src: url(https://fonts.example/jp-${name}.woff2); unicode-range: ${range}; }`).join('\n');
const CARLITO_CSS = `@font-face { font-family: 'Carlito'; font-weight: 400; src: url(https://fonts.example/carlito.woff2); unicode-range: ${SUBSETS.latin}; }`;

function install() {
  const faces: FakeFontFace[] = [];
  const set = {
    add: (face: FakeFontFace) => { faces.push(face); },
    delete: (face: FakeFontFace) => { const at = faces.indexOf(face); if (at >= 0) faces.splice(at, 1); return at >= 0; },
  } as unknown as FontFaceSet;
  G.FontFace = FakeFontFace;
  const fetched: string[] = [];
  G.fetch = vi.fn(async (url: string) => {
    fetched.push(url);
    // A face file the library fetches itself: its body names the file.
    if (/\.woff2$/.test(url)) {
      return { ok: true, status: 200, url, arrayBuffer: async () => new TextEncoder().encode(url).buffer };
    }
    const css = url.includes('Noto') ? NOTO_CSS : url.includes('Carlito') ? CARLITO_CSS : '';
    return { ok: css !== '', status: css ? 200 : 404, url, text: async () => css };
  });
  return { faces, set, fetched };
}

afterEach(() => {
  G.FontFace = ORIGINAL.FontFace;
  G.fetch = ORIGINAL.fetch;
});

const SOURCE = fontCssApi({
  urlTemplate: 'https://fonts.example/css?family={family}',
  families: { 'Noto Sans JP': true, Carlito: true },
  query: 'x',
  substitutes: OFFICE_FONT_SUBSTITUTIONS,
});

describe('loadDocxFontSources', () => {
  it('loads only the subsets whose ranges intersect the demand', async () => {
    const { faces, set, fetched } = install();
    // Latin text, CJK fallback demanded: the kanji subset never loads.
    const latin = docWith([text('Hello')]);
    const plan = { ...docxFontFamilyPlan(latin, [SOURCE], undefined, { routes: {} }), fallbackFamilies: ['Noto Sans JP'] };
    const loaded = await loadDocxFontSources(latin, plan, [SOURCE], { fontSet: set });
    // A demanded subset is registered from the bytes the library fetched;
    // the others are neither fetched nor registered.
    const sourceName = (source: unknown) => typeof source === 'string' ? source
      : `bytes:${new TextDecoder().decode(source as ArrayBuffer)}`;
    const status = () => Object.fromEntries(faces.filter((face) => face.family === loaded.loaded.families.get('noto sans jp'))
      .map((face) => [sourceName(face.source), face.status]));
    expect(status()).toEqual({
      'bytes:https://fonts.example/jp-latin.woff2': 'loaded',
      // U+3042 and U+4E00, layout probes (LAYOUT_SYNTHESIZED_SCALARS).
      'bytes:https://fonts.example/jp-kana.woff2': 'loaded',
    });
    expect(loaded.loaded.substitutions.get('calibri')).toEqual({ family: 'Carlito' });
    expect(loaded.names).toEqual(['Calibri', 'Noto Sans JP']);
    // One stylesheet per served family (Calibri is asked for through
    // Carlito), and one request per demanded subset.
    expect(fetched.sort()).toEqual([
      'https://fonts.example/carlito.woff2',
      'https://fonts.example/css?family=Carlito',
      'https://fonts.example/css?family=Noto+Sans+JP',
      'https://fonts.example/jp-kana.woff2',
      'https://fonts.example/jp-latin.woff2',
    ]);
    loaded.session.release();
    expect(faces).toHaveLength(0);

    // Mixed Latin + kanji: the kanji subset loads too.
    const mixed = docWith([text('Hello 本')]);
    const mixedPlan = docxFontFamilyPlan(mixed, [SOURCE], undefined, { routes: {} });
    expect(mixedPlan.fallbackFamilies).toContain('Noto Sans JP');
    const second = await loadDocxFontSources(mixed, mixedPlan, [SOURCE], { fontSet: set });
    const jp = second.loaded.families.get('noto sans jp');
    expect(faces.filter((face) => face.family === jp).every((face) => face.status === 'loaded')).toBe(true);
    second.session.release();
  });

  it('keeps a Symbol-bullet document lazy: only the subsets of its painted scalars load', async () => {
    const { faces, set } = install();
    const doc = docWith([text('Hello')]);
    (doc.body[0] as { numbering?: object }).numbering = { text: '\uF0B7', fontFamily: 'Symbol' };
    const plan = { ...docxFontFamilyPlan(doc, [SOURCE], undefined, { routes: {} }), fallbackFamilies: ['Noto Sans JP'] };
    const loaded = await loadDocxFontSources(doc, plan, [SOURCE], { fontSet: set });
    const jp = loaded.loaded.families.get('noto sans jp');
    // The kanji subset holds no scalar the bullet or its conversion paints.
    expect(faces.filter((face) => face.family === jp).map((face) => [face.unicodeRange, face.status]))
      .toEqual([[SUBSETS.latin, 'loaded'], [SUBSETS.kana, 'loaded']]);
    loaded.session.release();
  });

  it('forwards the worker bridge options: claimed resolver URL faces go through the acquirer', async () => {
    const { faces, set, fetched } = install();
    const url = 'https://fonts.app.example/brand.woff2';
    const source = fontResolver((request) => request.families.map((family) => ({ family, url })));
    const bytes = new Uint8Array(12).buffer;
    const acquireResolverFaces = vi.fn(async (_request: { readonly sourceIndex: number }) => new Map([[url, { data: bytes }]]));
    const doc = docWith([text('Hello', { fontFamily: 'Brand' })]);
    const plan = { ...docxFontFamilyPlan(doc, [source], undefined, { routes: {} }), families: ['Brand'] };
    const loaded = await loadDocxFontSources(doc, plan, [source], {
      fontSet: set, acquireResolverFaces, acquiredFacesOwnedBySession: true,
    });
    expect(acquireResolverFaces).toHaveBeenCalledTimes(1);
    expect(acquireResolverFaces.mock.calls[0][0]).toMatchObject({ sourceIndex: 0, families: [{ family: 'Brand', urls: [url] }] });
    expect(faces.filter((face) => face.family === loaded.loaded.families.get('brand')).map((face) => face.source)).toEqual([bytes]);
    expect(fetched).toEqual([]);
    loaded.session.release();
  });

  it('never demands an installed authored family', async () => {
    const { set, fetched } = install();
    const doc = docWith([text('Hello')]);
    const plan = docxFontFamilyPlan(doc, [SOURCE], undefined, { routes: {}, installed: ['calibri'] });
    expect(plan.families).toEqual([]);
    expect(plan.installedSubstituteFamilies).toEqual(['calibri']);
    const loaded = await loadDocxFontSources(doc, plan, [SOURCE], { fontSet: set });
    expect(fetched).toEqual([]);
    loaded.session.release();
  });

  it('lets an earlier source serve a scoped-substitute family by name; the later substitute is never requested', async () => {
    const { set, fetched } = install();
    const doc = docWith([text('Latin only', { fontFamily: 'Sakkal Majalla' })]);
    const own = fontFiles({ faces: [{ family: 'Sakkal Majalla', data: new Uint8Array([0, 1, 0, 0]).buffer }] });
    const naskh = fontCssApi({
      urlTemplate: 'https://fonts.example/css?family={family}',
      families: { 'Noto Naskh Arabic': true },
      query: 'x',
      substitutes: OFFICE_FONT_SUBSTITUTIONS,
    });
    for (const sources of [[own, naskh], [own]]) {
      const plan = docxFontFamilyPlan(doc, sources, undefined, { routes: {} });
      expect(plan.families).toContain('Sakkal Majalla');
      const loaded = await loadDocxFontSources(doc, plan, sources, { fontSet: set });
      expect(loaded.loaded.faces.map((face) => face.family)).toContain('Sakkal Majalla');
      loaded.session.release();
    }
    // Without an own face, the Arabic-scoped substitute is not requested for
    // Latin-only text.
    const plan = docxFontFamilyPlan(doc, [naskh], undefined, { routes: {} });
    const loaded = await loadDocxFontSources(doc, plan, [naskh], { fontSet: set });
    expect(loaded.loaded.substitutions.size).toBe(0);
    expect(fetched).toEqual([]);
    loaded.session.release();
  });
});
