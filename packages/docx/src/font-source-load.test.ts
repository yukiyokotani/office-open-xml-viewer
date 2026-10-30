import { afterEach, describe, expect, it, vi } from 'vitest';
import { fontCssApi, fontFiles, fontResolver, OFFICE_FONT_SUBSTITUTIONS } from '@silurus/ooxml-core/internal/font-sources';
import { docxFontFamilyPlan } from './font-plan.js';
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
    ['a symbol font', docWith([text('l', { fontFamily: 'Wingdings' })])],
    ['a PUA symbol code', docWith([text('')])],
    ['a non-ASCII page number format', docWith([text('a')], { section: { pageNumType: { fmt: 'hebrew1' } } as never })],
    ['a non-ASCII field switch', docWith([{ type: 'field', fieldType: 'page', instruction: 'PAGE \\* HEBREW1', fallbackText: '1' }])],
  ])('answers all for %s', (_label, doc) => {
    expect(scalarsOf(doc)).toBe('all');
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
    const status = () => Object.fromEntries(faces.filter((face) => face.family === 'Noto Sans JP')
      .map((face) => [face.source, face.status]));
    expect(status()).toEqual({
      'url("https://fonts.example/jp-latin.woff2")': 'loaded',
      // U+3042 and U+4E00, layout probes (LAYOUT_SYNTHESIZED_SCALARS).
      'url("https://fonts.example/jp-kana.woff2")': 'loaded',
      'url("https://fonts.example/jp-kanji.woff2")': 'unloaded',
    });
    expect(loaded.loaded.substitutions.get('calibri')).toEqual({ family: 'Carlito' });
    expect(loaded.names).toEqual(['Calibri', 'Noto Sans JP']);
    // One stylesheet per served family; Calibri is asked for through Carlito.
    expect(fetched.sort()).toEqual([
      'https://fonts.example/css?family=Carlito',
      'https://fonts.example/css?family=Noto+Sans+JP',
    ]);
    loaded.session.release();
    expect(faces).toHaveLength(0);

    // Mixed Latin + kanji: the kanji subset loads too.
    const mixed = docWith([text('Hello 本')]);
    const mixedPlan = docxFontFamilyPlan(mixed, [SOURCE], undefined, { routes: {} });
    expect(mixedPlan.fallbackFamilies).toContain('Noto Sans JP');
    const second = await loadDocxFontSources(mixed, mixedPlan, [SOURCE], { fontSet: set });
    expect(faces.filter((face) => face.family === 'Noto Sans JP').every((face) => face.status === 'loaded')).toBe(true);
    second.session.release();
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
    expect(faces.filter((face) => face.family === 'Brand').map((face) => face.source)).toEqual([bytes]);
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
