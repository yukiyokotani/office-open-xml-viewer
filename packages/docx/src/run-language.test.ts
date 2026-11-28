import { describe, expect, it, vi } from 'vitest';

// Node has no Canvas realm to run the route sentinels in; production judges
// named-family ownership there (core createNamedFamilyOracle). Rendering tests
// below use faces their recording canvas draws itself.
// 'ж' stands for text only a generic fallback draws.
// Text units the face-resolution predicates scan (see the faceText cost test).
const scanned = vi.hoisted(() => ({ units: 0 }));
vi.mock('@silurus/ooxml-core/internal/script-preload-accumulator', async (importOriginal) => {
  const original = await importOriginal<typeof import('@silurus/ooxml-core/internal/script-preload-accumulator')>();
  return {
    ...original,
    containsHanScript: (text: string) => { scanned.units += text.length; return original.containsHanScript(text); },
  };
});
vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const original = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return {
    ...original,
    fontSubstituteScriptCoversText: (...args: Parameters<typeof original.fontSubstituteScriptCoversText>) => {
      scanned.units += args[1].length;
      return original.fontSubstituteScriptCoversText(...args);
    },
    createNamedFamilyOracle: () => ({
      sentinelsApply: true,
      available: true,
      // U+0640 (a generated justification kashida) is drawn by named families
      // of "Kashida Face" only.
      // '〓' likewise stands for East Asian text only a fallback draws.
      // "Fallback Face" stands for a route no named family of which draws.
      draws: (list: string, _w: number, _s: string, text: string) => !/[ж〓]/u.test(text)
        && !list.includes('Fallback Face')
        && (!text.includes('\u0640') || list.includes('Kashida Face')),
    }),
  };
});
import { compatProfile } from './test-support/compat-environment.test-support.js';
import { createFontResolver } from './layout/font-service.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import { createLayoutServices } from './layout-runtime.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import { runOverlayMeasure } from './run-overlay-measure.js';
import { canvasLanguageTag, createTextLayoutService, rubyGuideShapeRequest, sliceTextShapeRequest, type GlyphMeasureRequest } from './layout/text.js';
import { renderDocumentToCanvas } from './renderer.js';
import { buildSegments } from './line-layout.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, NumberingInfo, SectionProps } from './types';

// ECMA-376 §17.3.2.20 w:lang — w:val governs text that is neither East Asian
// nor complex script, w:eastAsia and w:bidi the other two §17.3.2.26 slots.
// The run language reaches Canvas `lang`, which selects language-specific
// OpenType forms (locl), e.g. Serbian italic Cyrillic (#2001). It reaches
// Canvas only for text a named face of the route draws: generic and
// last-resort fallback faces are chosen per language (CSS Fonts 4 §5.1/§5.4).

describe('text service: §17.3.2.20 language per font slot', () => {
  const measured: GlyphMeasureRequest[] = [];
  const service = createTextLayoutService({
    fonts: createFontResolver([
      { requestedFamily: 'Face', resolvedFamily: 'Face', source: 'local' },
    ]),
    measurer: {
      fingerprint: 'run-language',
      // '中' stands for text no named family draws (generic fallback).
      namedFamiliesDraw: (request) => !request.text.includes('中'),
      measure: (request) => {
        measured.push(request);
        // A language-sensitive face: Serbian forms are narrower.
        const em = request.fontRoute.lang?.startsWith('sr') ? 4 : 5;
        return { advancePt: [...request.text].length * em, ascentPt: 8, descentPt: 2 };
      },
    },
  });
  const fonts = { ascii: 'Face', highAnsi: 'Face', eastAsia: 'Face', complexScript: 'Face' };
  // The fingerprint of the route 'Face' resolves to here (a pin's route).
  const ROUTE = service.shape({ text: 'b', fontSizePt: 10, fonts, clusterGeometry: false }).spans[0]!.fontRoute.fingerprint;

  it('measures each slot with its own language and reports it on the span', () => {
    measured.length = 0;
    const languages = { language: 'sr-latn-rs', eastAsiaLanguage: 'ja-jp', complexScriptLanguage: 'ar-sa' };
    const mixed = service.shape({ text: 'ж中', fontSizePt: 10, fonts, clusterGeometry: false, ...languages });
    // §17.3.2.7 w:cs puts the whole run on the complex-script slot.
    const cs = service.shape({
      text: 'بa', fontSizePt: 10, fonts, clusterGeometry: false, complexScript: true, ...languages,
    });
    const kana = service.shape({ text: 'か', fontSizePt: 10, fonts, clusterGeometry: false, ...languages });
    expect([...mixed.spans, ...cs.spans, ...kana.spans].map((span) => [span.script, span.fontRoute.lang])).toEqual([
      // Generic fallback draws '中': its language would choose the face.
      ['highAnsi', 'sr-latn-rs'], ['eastAsia', undefined], ['complexScript', 'ar-sa'], ['eastAsia', 'ja-jp'],
    ]);
    expect(measured.map((request) => [request.text, request.fontRoute.lang])).toEqual([
      ['ж', 'sr-latn-rs'], ['中', undefined], ['بa', 'ar-sa'], ['か', 'ja-jp'],
    ]);
    // A span without a language keeps the resolved route itself, so its
    // identity, cache keys and output are those of untagged text.
    expect(mixed.spans[1]!.fontRoute).toBe(mixed.spans[1]!.font.route);
    expect('lang' in mixed.spans[1]!.fontRoute).toBe(false);
    // An admitted language is part of the route identity; the resolved faces
    // (font.route, the line-metric identity) are unchanged.
    expect(mixed.spans[0]!.fontRoute.fingerprint).toBe(`${mixed.spans[0]!.font.route.fingerprint}:lang:sr-latn-rs`);
    expect(mixed.spans[0]!.fontRoute.familyList).toBe(mixed.spans[0]!.font.route.familyList);
    expect(kana.spans[0]!.font.route.fingerprint).toBe(mixed.spans[1]!.fontRoute.fingerprint);
    expect(kana.spans[0]!.fontRoute.fingerprint).not.toBe(mixed.spans[1]!.fontRoute.fingerprint);
  });

  it('keys measurement by the route, whose identity has a language only when one is set', () => {
    measured.length = 0;
    const serbian = service.shape({ text: 'бгд', fontSizePt: 10, fonts, language: 'sr-rs', clusterGeometry: false });
    const plain = service.shape({ text: 'бгд', fontSizePt: 10, fonts, clusterGeometry: false });
    // Word's "x-none" names no language: same measurement as untagged text.
    const none = service.shape({ text: 'бгд', fontSizePt: 10, fonts, language: 'x-none', clusterGeometry: false });
    expect([serbian.advancePt, plain.advancePt, none.advancePt]).toEqual([12, 15, 15]);
    expect(measured.map((request) => request.fontRoute.lang ?? 'absent')).toEqual(['sr-rs', 'absent']);
    expect(plain.spans[0]!.fontRoute).toBe(none.spans[0]!.fontRoute);
    expect(serbian.spans[0]!.fontRoute.fingerprint).not.toBe(plain.spans[0]!.fontRoute.fingerprint);
    // Not a language tag (and Canvas's own keyword): no language either.
    for (const language of ['inherit', 'not a tag', 'x'.repeat(200)]) {
      expect(service.shape({ text: 'бгд', fontSizePt: 10, fonts, language, clusterGeometry: false }).spans[0]!.fontRoute)
        .toBe(plain.spans[0]!.fontRoute);
    }
  });

  it('sets no language without named-family ownership evidence', () => {
    const unjudged = createTextLayoutService({
      fonts: createFontResolver([{ requestedFamily: 'Face', resolvedFamily: 'Face', source: 'local' }]),
      measurer: { fingerprint: 'unjudged', measure: request => {
        expect(request.fontRoute.lang).toBeUndefined();
        return { advancePt: 5, ascentPt: 8, descentPt: 2 };
      } },
    });
    const shaped = unjudged.shape({ text: 'б', fontSizePt: 10, fonts, language: 'sr-rs', clusterGeometry: false });
    expect(shaped.spans[0]!.fontRoute.lang).toBeUndefined();
  });

  it('reuses a piece\'s acquired route language in every derived request of that slot group', () => {
    const withheld = { text: 'бгд', fontSizePt: 10, fonts, language: 'sr-rs', clusterGeometry: false as const,
      canvasLanguagePin: { group: 'latin' as const, route: ROUTE, lang: null } };
    expect(service.shape(withheld).spans[0]!.fontRoute.lang).toBeUndefined();
    expect(service.shape(sliceTextShapeRequest(withheld, 0, 1)).spans[0]!.fontRoute.lang).toBeUndefined();
    // A fragment of an accepted piece shapes through the piece's route: the
    // retained decision is not judged again for a slice of the same text.
    const accepted = { ...withheld, canvasLanguagePin: { group: 'latin' as const, route: ROUTE, lang: 'sr-rs' } };
    const whole = service.shape(accepted).spans[0]!.fontRoute;
    expect(whole.lang).toBe('sr-rs');
    expect(service.shape(sliceTextShapeRequest(accepted, 1, 3)).spans[0]!.fontRoute).toBe(whole);
    // Another slot group is decided on its own.
    expect(service.shape({ ...accepted, text: 'か', eastAsiaLanguage: 'ja-jp' }).spans.map(span => span.fontRoute.lang))
      .toEqual(['ja-jp']);
    // A decision proved on another route does not carry over: proved afresh.
    expect(service.shape({ ...withheld, canvasLanguagePin: { group: 'latin' as const, route: 'other', lang: null } })
      .spans[0]!.fontRoute.lang).toBe('sr-rs');
  });

  it('shapes a ruby guide without a Canvas language, keeping the base regional font input', () => {
    const base = { text: '漢', fontSizePt: 10, fonts, language: 'en-us', eastAsiaLanguage: 'zh-cn',
      complexScriptLanguage: 'ar-sa', canvasLanguagePin: { group: 'eastAsia' as const, route: ROUTE, lang: 'zh-cn' } };
    const guide = rubyGuideShapeRequest(base, 'かんab');
    expect([guide.language, guide.complexScriptLanguage, guide.eastAsiaLanguage, guide.canvasEastAsiaLanguage,
      guide.canvasLanguagePin]).toEqual([undefined, undefined, 'zh-cn', null, undefined]);
    expect(service.shape({ ...guide, clusterGeometry: false }).spans.map(span => span.fontRoute.lang))
      .toEqual([undefined, undefined]);
    // An untagged base gives an untagged guide request, as before languages.
    const untagged = rubyGuideShapeRequest({ text: '漢', fontSizePt: 10, fonts }, 'かん');
    expect('canvasEastAsiaLanguage' in untagged).toBe(false);
  });

  it('shapes a fragment through the route its piece was resolved and painted with', () => {
    // Regional fallback is chosen only for text containing Han: a fragment
    // without Han would otherwise resolve another route than its piece.
    const regional = createTextLayoutService({
      fonts: createFontResolver([], {
        regionalFamilyLists: { sc: { Absent: '"SC Face", serif' } },
        nativeFamilyLists: { Absent: '"Plain Face", serif' },
      }),
      measurer: { fingerprint: 'regional', namedFamiliesDraw: () => true,
        measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }) },
    });
    const piece = { text: 'か漢', fontSizePt: 10, fonts: { eastAsia: 'Absent' }, eastAsiaLanguage: 'zh-cn',
      clusterGeometry: false as const, faceText: 'か漢' };
    const whole = regional.shape(piece).spans[0]!.fontRoute;
    expect([whole.familyList, whole.lang]).toEqual(['"SC Face", serif', 'zh-cn']);
    // The fragment keeps the piece's route and language.
    expect(regional.shape(sliceTextShapeRequest(piece, 0, 1)).spans[0]!.fontRoute).toBe(whole);
    // Without the piece (an independent probe), the fragment resolves its own.
    expect(regional.shape({ ...piece, text: 'か', faceText: undefined }).spans[0]!.fontRoute.familyList)
      .toBe('"Plain Face", serif');
  });

  it('keys a fragment by what a resolution reads from its piece, never by the piece text', () => {
    const regional = createTextLayoutService({
      fonts: createFontResolver([], {
        regionalFamilyLists: { sc: { Absent: '"SC Face", serif' } },
        nativeFamilyLists: { Absent: '"Plain Face", serif' },
      }),
      measurer: { fingerprint: 'regional-keys', namedFamiliesDraw: () => true,
        measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }) },
    });
    const serialized = (count: number) => {
      const text = Array.from({ length: count }, (_, index) => String.fromCharCode(0x4e00 + index)).join('');
      const piece = { text, fontSizePt: 10, fonts: { eastAsia: 'Absent' }, eastAsiaLanguage: 'zh-cn',
        clusterGeometry: false as const, measure: false as const, faceText: text };
      const route = regional.shape(piece).spans[0]!.fontRoute;
      const stringify = JSON.stringify;
      let units = 0;
      let carriesPiece = false;
      const spy = vi.spyOn(JSON, 'stringify').mockImplementation((...args: Parameters<typeof JSON.stringify>) => {
        const out = stringify(...args);
        units += out?.length ?? 0;
        carriesPiece ||= out?.includes(text) === true;
        // The piece enters a fragment's key as its Han presence alone,
        // whatever the piece's length.
        const part = /\["faceText",("[^"]*")\]/u.exec(out ?? '');
        if (part) expect(part[1]).toBe('"H"');
        return out;
      });
      try {
        // Every one-character fragment still shapes through the piece's route.
        for (let index = 0; index < count; index += 1) {
          expect(regional.shape(sliceTextShapeRequest(piece, index, index + 1)).spans[0]!.fontRoute).toBe(route);
        }
      } finally {
        spy.mockRestore();
      }
      expect(carriesPiece).toBe(false);
      return units;
    };
    // Linear in the number of fragments: no key repeats the piece's text.
    const small = serialized(128);
    const large = serialized(512);
    expect(large / small).toBeLessThan(4.5);
    // 128 one-character fragments: well under the 135,424 units their keys
    // took when each carried the piece.
    expect(small).toBeLessThan(128 * 400);
    // Pieces that resolve alike share their fragments' shapes.
    const shapes: string[] = [];
    const shared = createTextLayoutService({
      fonts: createFontResolver([], { nativeFamilyLists: { Absent: '"Plain Face", serif' } }),
      measurer: { fingerprint: 'shared-keys', namedFamiliesDraw: () => true,
        measure: (request) => {
          shapes.push(request.text);
          return { advancePt: 10, ascentPt: 8, descentPt: 2 };
        } },
    });
    for (const text of ['\u4e00\u4e8c', '\u4e00\u4e09']) {
      shared.shape(sliceTextShapeRequest({ text, fontSizePt: 10, fonts: { eastAsia: 'Absent' },
        eastAsiaLanguage: 'zh-cn', clusterGeometry: false as const, faceText: text }, 0, 1));
    }
    expect(shapes).toEqual(['\u4e00']);
  });

  it('scans a piece once for its face facts, however many fragments resolve for it', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver([
        { requestedFamily: 'Scoped Face', resolvedFamily: 'Arabic Substitute', source: 'substitute', script: 'arabic' },
      ], {
        scriptScopedFamilies: { 'scoped face': { script: 'arabic', substituteFamilies: ['Arabic Substitute'] } },
        regionalFamilyLists: { sc: { Absent: '"SC Face", serif' } },
        nativeFamilyLists: { Absent: '"Plain Face", serif' },
      }),
      measurer: { fingerprint: 'face-scan', namedFamiliesDraw: () => true,
        measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }) },
    });
    const unitsFor = (count: number) => {
      // Han with an Arabic letter in the middle: every slot names the scoped
      // face, so each slot's coverage is a fact of the piece.
      const text = Array.from({ length: count }, (_, index) =>
        index === count >> 1 ? '\u0628' : String.fromCharCode(0x4e00 + count + index)).join('');
      const piece = { text, fontSizePt: 10, eastAsiaLanguage: 'zh-cn', clusterGeometry: false as const,
        measure: false as const, faceText: text,
        fonts: { ascii: 'Scoped Face', highAnsi: 'Scoped Face', eastAsia: 'Scoped Face', complexScript: 'Scoped Face' } };
      scanned.units = 0;
      const whole = service.shape(piece);
      for (let index = 0; index < count; index += 1) service.shape(sliceTextShapeRequest(piece, index, index + 1));
      expect(whole.spans.length).toBeGreaterThan(0);
      return scanned.units;
    };
    // One Han scan and at most one coverage scan per slot of the piece, plus
    // the fragments' own (one-unit) texts: linear, not count × count (a
    // rescan per fragment took 17,024 units for 128 fragments, 264,704 for
    // 512).
    for (const count of [128, 512]) expect(unitsFor(count)).toBeLessThanOrEqual(8 * count);
  });

  it('resolves untagged text directly, consulting no table keyed by its text', () => {
    const service = createTextLayoutService({
      fonts: createFontResolver([
        { requestedFamily: 'Scoped Face', resolvedFamily: 'Arabic Substitute', source: 'substitute', script: 'arabic' },
      ], {
        scriptScopedFamilies: { 'scoped face': { script: 'arabic', substituteFamilies: ['Arabic Substitute'] } },
        regionalFamilyLists: { sc: { Absent: '"SC Face", serif' } },
        nativeFamilyLists: { Absent: '"Plain Face", serif' },
      }),
      measurer: { fingerprint: 'untagged-resolve', namedFamiliesDraw: () => true,
        measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }) },
    });
    const text = '\u4e00\u0628\u4e8c-untagged';
    const lookups = { count: 0 };
    const tableGet = Map.prototype.get;
    const tableHas = Map.prototype.has;
    const spyGet = vi.spyOn(Map.prototype, 'get').mockImplementation(function (this: Map<unknown, unknown>, key: unknown) {
      if (key === text) lookups.count += 1;
      return tableGet.call(this, key);
    });
    const spyHas = vi.spyOn(Map.prototype, 'has').mockImplementation(function (this: Map<unknown, unknown>, key: unknown) {
      if (key === text) lookups.count += 1;
      return tableHas.call(this, key);
    });
    try {
      for (const slot of ['ascii', 'eastAsia', 'complexScript'] as const) {
        for (const fonts of [{ [slot]: 'Absent' }, { [slot]: 'Scoped Face' }]) {
          scanned.units = 0;
          service.resolve({ text, slot, fonts, eastAsiaLanguage: 'zh-cn', weight: 400, style: 'normal' });
          // Re-resolving the text scans it again: no retained fact answers it.
          const first = scanned.units;
          service.resolve({ text, slot, fonts, eastAsiaLanguage: 'zh-cn', weight: 400, style: 'normal' });
          expect(scanned.units).toBe(2 * first);
        }
      }
      // Untagged shaping (no faceText) resolves its groups the same way. (A
      // scoped face would add main's per-run scope cache, keyed by its run.)
      service.shape({ text, fontSizePt: 10, eastAsiaLanguage: 'zh-cn', clusterGeometry: false as const,
        measure: false as const, fonts: { eastAsia: 'Absent', complexScript: 'Absent' } });
      expect(lookups.count).toBe(0);
      // The probe sees the table a tagged piece's fragments use.
      service.shape(sliceTextShapeRequest({ text, fontSizePt: 10, eastAsiaLanguage: 'zh-cn',
        clusterGeometry: false as const, measure: false as const, faceText: text,
        fonts: { eastAsia: 'Absent' } }, 0, 1));
      expect(lookups.count).toBeGreaterThan(0);
    } finally {
      spyGet.mockRestore();
      spyHas.mockRestore();
    }
  });

  it('selects the Canvas language of a script-class run independently of the slot table', () => {
    const regional = createTextLayoutService({
      fonts: createFontResolver([], {
        regionalFamilyLists: { sc: { Absent: '"SC Face", serif' } },
        nativeFamilyLists: { Absent: '"Plain Face", serif' },
      }),
      measurer: { fingerprint: 'script-class', namedFamiliesDraw: () => true,
        measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }) },
    });
    const base = { fontSizePt: 10, fonts: { ascii: 'Absent', highAnsi: 'Absent', eastAsia: 'Absent', complexScript: 'Absent' },
      language: 'fr-fr', eastAsiaLanguage: 'zh-cn', complexScriptLanguage: 'ar-sa', clusterGeometry: false as const };
    const languages = (request: Parameters<typeof regional.shape>[0]) =>
      regional.shape(request).spans.map(span => [span.text, span.script, span.fontRoute.lang]);
    // U+1D11E takes the East Asian slot of §17.3.2.26, yet it is text of the
    // run's one (Latin-class) language; without the fact the slot decides.
    expect(languages({ ...base, text: 'a\u{1d11e}' })).toEqual([['a', 'ascii', 'fr-fr'], ['\u{1d11e}', 'eastAsia', 'zh-cn']]);
    expect(languages({ ...base, text: 'a\u{1d11e}', scriptClassLanguage: 'fr-FR' }).map(span => span[2]))
      .toEqual(['fr-fr', 'fr-fr']);
    // Present and empty: the class names no language, whatever the axes say.
    expect(languages({ ...base, text: 'a\u{1d11e}\u5b57', scriptClassLanguage: '' })
      .every(span => span[2] === undefined)).toBe(true);
    // The Asian axis stays the regional font-selection input: Han of a
    // Latin-class run still resolves the zh-CN regional route.
    const han = regional.shape({ ...base, text: '\u5b57', scriptClassLanguage: 'fr-FR' }).spans[0]!.fontRoute;
    expect([han.familyList, han.lang]).toEqual(['"SC Face", serif', 'fr-fr']);
    // A ruby guide takes no language from a script-class base either.
    const guide = rubyGuideShapeRequest({ ...base, text: '\u5b57', scriptClassLanguage: 'zh-CN' }, '\u3058');
    expect(guide.scriptClassLanguage).toBe('');
    expect(regional.shape(guide).spans.map(span => span.fontRoute.lang)).toEqual([undefined]);
  });

  it('keeps independent languages when scoped Arabic shares a physical face across slots', () => {
    const scoped = createTextLayoutService({
      fonts: createFontResolver([
        { requestedFamily: 'Scoped Face', resolvedFamily: 'Arabic Substitute', source: 'substitute', script: 'arabic' },
      ], { scriptScopedFamilies: { 'scoped face': { script: 'arabic', substituteFamilies: ['Arabic Substitute'] } } }),
      measurer: { fingerprint: 'scoped-language', namedFamiliesDraw: () => true, measure: request => ({
        advancePt: [...request.text].length * (request.fontRoute.lang === 'ar' ? 4 : 5), ascentPt: 8, descentPt: 2,
      }) },
    });
    const request = {
      // BMP Arabic is ASCII-slot; supplementary Arabic mathematical letters
      // are East-Asian-slot in §17.3.2.26, even with the same scoped face.
      text: 'ا\u{1EE00}', fontSizePt: 10, clusterGeometry: false as const,
      fonts: { ascii: 'Scoped Face', eastAsia: 'Scoped Face' },
      language: 'ar', eastAsiaLanguage: 'ur',
    };
    const split = scoped.shape(request);
    expect(split.spans.map(span => [span.text, span.fontRoute.lang])).toEqual([['ا', 'ar'], ['\u{1EE00}', 'ur']]);
    expect(split.advancePt).toBe(9);
    // Equal languages retain the existing joining-safe physical shape.
    expect(scoped.shape({ ...request, eastAsiaLanguage: 'ar' }).spans.map(span => span.text))
      .toEqual([request.text]);
  });
});

describe('Canvas lang at measure and paint', () => {
  const FONT_PX = 20;
  function recordingCanvas() {
    let font = `${FONT_PX}px serif`;
    const draws: Array<{ text: string; x: number; lang: string }> = [];
    const measures: Array<{ text: string; lang: string }> = [];
    const ctx = {
      lang: 'inherit',
      get font() { return font; },
      set font(v: string) { font = v; },
      letterSpacing: '0px',
      fontKerning: 'auto',
      measureText(this: { lang: string }, s: string) {
        measures.push({ text: s, lang: this.lang });
        const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
        // Serbian forms are narrower, so fitting depends on the language.
        const em = this.lang.startsWith('sr') ? 0.4 * p : 0.5 * p;
        // An 'fi' ligature makes advances non-additive across a seam.
        const ligatures = s.split('fi').length - 1;
        return {
          width: ([...s].length - 0.5 * ligatures) * em,
          fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
          actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
        } as TextMetrics;
      },
      fillText(this: { lang: string }, text: string, x: number) { draws.push({ text, x, lang: this.lang }); },
      save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
      stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
      setLineDash() {}, drawImage() {}, clearRect() {}, arc() {},
      quadraticCurveTo() {}, bezierCurveTo() {},
      createLinearGradient() { return { addColorStop() {} }; },
      scale() {}, translate() {}, rotate() {}, transform() {}, setTransform() {}, resetTransform() {},
      strokeText() {},
      fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
      textAlign: 'left' as CanvasTextAlign, textBaseline: 'alphabetic' as CanvasTextBaseline,
      direction: 'ltr' as CanvasDirection,
      globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
    };
    const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
    return { canvas: canvas as unknown as HTMLCanvasElement, ctx, draws, measures };
  }

  function run(text: string, extra: Record<string, unknown> = {}): DocxTextRun {
    return {
      type: 'text', text, bold: false, italic: true, underline: false, strikethrough: false,
      fontSize: FONT_PX, color: null, fontFamily: 'Times New Roman', isLink: false,
      background: null, vertAlign: null, hyperlink: null, ...extra,
    } as unknown as DocxTextRun;
  }

  function documentOf(runs: DocxTextRun[]): DocxDocumentModel {
    const paragraph = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs, defaultFontSize: FONT_PX, defaultFontFamily: 'Times New Roman', widowControl: false,
    } as unknown as DocParagraph;
    const section = {
      pageWidth: 600, pageHeight: 600, marginTop: 0, marginRight: 0, marginBottom: 0,
      marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
      evenAndOddHeaders: false,
    } as SectionProps;
    return {
      section, body: [paragraph as unknown as BodyElement],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;
  }

  it('preserves numbering-level language through marker measurement and paint', async () => {
    const { canvas, ctx, draws, measures } = recordingCanvas();
    const document = documentOf([run('Body')]);
    const paragraph = document.body[0] as DocParagraph;
    paragraph.numbering = {
      numId: 1, level: 0, format: 'bullet', text: 'где', indentLeft: 0,
      tab: 18, suff: 'space', fontFamily: 'Times New Roman',
      fontFacts: { lang: 'sr-rs', fontSize: FONT_PX, italic: true },
    } as unknown as NumberingInfo;
    await renderDocumentToCanvas(document, canvas, 0, { dpr: 1, width: 600 });
    expect(draws.filter(draw => draw.text === 'где').map(draw => draw.lang)).toEqual(['sr-rs']);
    expect(measures.filter(measure => measure.text === 'где').map(measure => measure.lang))
      .toContain('sr-rs');
    expect(draws.filter(draw => draw.text === 'Body').map(draw => draw.lang)).toEqual(['inherit']);
    expect(ctx.lang).toBe('inherit');
  });

  it('draws a tagged run with the language it was measured with, and restores it', async () => {
    const { canvas, ctx, draws, measures } = recordingCanvas();
    await renderDocumentToCanvas(documentOf([
      run('где', { lang: 'sr-rs' }),
      run('где'),
    ]), canvas, 0, { dpr: 1, width: 600 });
    const glyphs = draws.filter((draw) => draw.text === 'где');
    expect(glyphs.map((draw) => draw.lang)).toEqual(['sr-rs', 'inherit']);
    // The second run starts where the Serbian-measured first run ends.
    expect(glyphs[1]!.x - glyphs[0]!.x).toBeCloseTo(3 * 0.4 * FONT_PX, 6);
    const textMeasures = measures.filter((measure) => measure.text === 'где');
    expect(new Set(textMeasures.map((measure) => measure.lang))).toEqual(new Set(['sr-rs', 'inherit']));
    expect(ctx.lang).toBe('inherit');
  });

  it('keeps a grapheme cluster whole across a language-only run seam', async () => {
    const { canvas, draws } = recordingCanvas();
    // A base and its variation selector authored in runs that differ only in
    // w:val stay one cluster, drawn once in the language of the cluster's
    // first unit; the rest of the second run keeps its own language.
    await renderDocumentToCanvas(documentOf([
      run('\u25fc', { lang: 'fr-ch' }),
      run('\ufe0e', { lang: 'en-us' }),
      run('ab', { lang: 'en-us' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([['\u25fc\ufe0e', 'fr-ch'], ['ab', 'en-us']]);
  });

  it('transfers only the crossing cluster to its base language', async () => {
    const { canvas, draws } = recordingCanvas();
    // 'я' + U+0301 share one font slot and route: the cluster is one draw, in
    // Serbian (its base), and every other cluster keeps its own run language.
    await renderDocumentToCanvas(documentOf([
      run('гя', { lang: 'sr-rs' }),
      run('\u0301бв', { lang: 'en-us' }),
      run('г', { lang: 'en-us' }),
      run('д', { lang: 'sr-rs' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([
      ['гя\u0301', 'sr-rs'], ['бвг', 'en-us'], ['д', 'sr-rs'],
    ]);
    // Each piece starts where the previous one, measured in its own
    // language, ends (Serbian forms are narrower).
    expect(draws[1]!.x - draws[0]!.x).toBeCloseTo(3 * 0.4 * FONT_PX, 6);
    expect(draws[2]!.x - draws[1]!.x).toBeCloseTo(3 * 0.5 * FONT_PX, 6);

    // The language seam inside the word is no line-break opportunity: a word
    // that does not fit after 'гггг ' moves to the next line whole.
    const wrapped = recordingCanvas();
    const narrow = documentOf([
      run('гггг гя', { lang: 'sr-rs' }),
      run('\u0301бв', { lang: 'en-us' }),
    ]);
    (narrow.section as SectionProps).pageWidth = 70;
    await renderDocumentToCanvas(narrow, wrapped.canvas, 0, { dpr: 1, width: 70 });
    expect(wrapped.draws.filter(draw => draw.text.trim())
      .map(draw => [draw.text.trim(), Math.round(draw.x * 1e6) / 1e6, draw.lang]))
      .toEqual([['гггг', 0, 'sr-rs'], ['гя\u0301', 0, 'sr-rs'], ['бв', 24, 'en-us']]);
  });

  it('confines the language transfer to the cluster in emoji and multi-run sequences', async () => {
    const { canvas, draws } = recordingCanvas();
    await renderDocumentToCanvas(documentOf([
      run('\u{1f469}', { lang: 'fr-ch' }),
      run('\u200d', { lang: 'en-us' }),
      run('\u{1f467} ', { lang: 'de-de' }),
      run('x', { lang: 'sr-rs' }),
      run('a', { lang: 'sr-rs' }),
      run('\u0301b', { lang: 'en-us' }),
      run('c', { lang: 'en-us' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    // §17.3.2.26 font slots still delimit shaped spans (an emoji is East
    // Asian text, which these runs give no language; U+0301 is not in the
    // slot of its ASCII base), exactly as inside one run. The western
    // language of every unit of a crossing cluster is its base run's; the
    // text after the cluster keeps its own.
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([
      ['\u{1f469}', 'inherit'], ['\u200d', 'fr-ch'], ['\u{1f467}', 'inherit'], [' ', 'de-de'],
      ['xa', 'sr-rs'], ['\u0301', 'sr-rs'], ['bc', 'en-us'],
    ]);
  });

  it('draws a cluster spread over three runs as one string in its base language', async () => {
    const { canvas, draws } = recordingCanvas();
    await renderDocumentToCanvas(documentOf([
      run('\u0433', { lang: 'sr-rs' }), run('\u0301', { lang: 'en-us' }), run('\u0323', { lang: 'sr-rs' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([['\u0433\u0301\u0323', 'sr-rs']]);
  });

  it('admits only well-formed BCP 47 tags', () => {
    expect(['sr-Latn-RS', 'i-klingon', 'zh-min-nan', 'x-private', 'iw'].map(canvasLanguageTag))
      .toEqual(['sr-latn-rs', 'i-klingon', 'zh-min-nan', 'x-private', 'he']);
    // Extlang tags need the registry's prefix pairs: none is admitted.
    expect(['zh-cmn-Hans-CN', 'en-cmn-Hans-CN', 'en-eng'].map(canvasLanguageTag)).toEqual([undefined, undefined, undefined]);
    expect(['en-a', 'en-us-gb', 'x-none', 'inherit', '', `en-${'a'.repeat(70)}`].map(canvasLanguageTag))
      .toEqual([undefined, undefined, undefined, undefined, undefined, undefined]);
  });

  it('shapes x-none and ill-formed tags exactly as untagged text', async () => {
    const drawn = async (lang?: string) => {
      const { canvas, draws } = recordingCanvas();
      await renderDocumentToCanvas(documentOf([run('f'), run('i', lang ? { lang } : {})]), canvas, 0, { dpr: 1, width: 600 });
      return draws.map(draw => [draw.text, draw.x, draw.lang]);
    };
    const untagged = await drawn();
    expect(untagged).toEqual([['fi', 0, 'inherit']]);
    expect(await drawn('x-none')).toEqual(untagged);
    expect(await drawn('en-a')).toEqual(untagged);
  });

  it('draws a cluster of runs naming one language on every axis whole in its base language', async () => {
    const { canvas, draws } = recordingCanvas();
    const uniform = (lang: string) => ({ lang, langEastAsia: lang, langBidi: lang });
    await renderDocumentToCanvas(documentOf([
      run('\u0433', uniform('sr-rs')), run('\u0301', uniform('en-us')),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([['\u0433\u0301', 'sr-rs']]);
  });

  it('shapes every range of a mixed sequence in its own run\'s language tuple', async () => {
    const { canvas, draws } = recordingCanvas();
    const all = (lang: string) => ({ lang, langEastAsia: lang, langBidi: lang });
    // One sequence: a Serbian base, a mark and kana in ja-JP on every axis,
    // then a mark and Han with en-US on w:val and ja-JP on the other axes.
    await renderDocumentToCanvas(documentOf([
      run('\u0433', all('sr-rs')), run('\u0301\u304b', all('ja-jp')),
      run('\u0301\u6f22', { lang: 'en-us', langEastAsia: 'ja-jp', langBidi: 'ja-jp' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => draw.text).join('')).toBe('\u0433\u0301\u304b\u0301\u6f22');
    // No text takes another run's East Asian language.
    const language = (text: string) => draws.find(draw => draw.text.includes(text))?.lang;
    expect([language('\u0433'), language('\u304b'), language('\u6f22')]).toEqual(['sr-rs', 'ja-jp', 'ja-jp']);
  });

  it('takes East Asian and complex-script Canvas languages by w:lang element presence', async () => {
    // ECMA-376 §17.3.2.20: a style naming w:eastAsia="ja-JP" and w:bidi,
    // then a run whose <w:lang> names only w:val. The parser keeps the
    // per-attribute langEastAsia/langBidi (regional font selection, bidi
    // digit classes) and records the element-presence Canvas axes apart.
    const style = { lang: 'en-us', langEastAsia: 'ja-jp', langBidi: 'ar-sa' };
    const partial = { ...style, langEastAsiaAutomatic: true, langBidiAutomatic: true };
    const { canvas, draws } = recordingCanvas();
    await renderDocumentToCanvas(documentOf([run('\u6f22', partial), run('x\u5b57', style)]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([
      ['\u6f22', 'inherit'], ['x', 'en-us'], ['\u5b57', 'ja-jp'],
    ]);
    // The regional font-selection input keeps the inherited w:eastAsia.
    const text = createLayoutServices(documentOf([]), { measureContext: recordingCanvas().ctx as never }).text;
    const [han] = buildSegments([run('\u6f22', partial)] as unknown as Parameters<typeof buildSegments>[0], { pageIndex: 0, totalPages: 1,
      compat: compatProfile({ compatibilityMode: 15 }), layoutServices: { text } as never });
    expect(han && 'textShapeRequest' in han ? han.textShapeRequest : undefined).toMatchObject({
      eastAsiaLanguage: 'ja-jp', canvasEastAsiaLanguage: null,
    });
    expect(han && 'textShapeRequest' in han ? han.textShapeRequest?.complexScriptLanguage : 'unset').toBeUndefined();
    // Complex-script text: no Canvas language, while digits stay AN under
    // the per-attribute w:bidi.
    const arabic = recordingCanvas();
    await renderDocumentToCanvas(documentOf([run('\u0633 12-34', { cs: true, ...partial })]), arabic.canvas, 0, { dpr: 1, width: 600 });
    expect(arabic.draws.map(draw => draw.text)).toEqual(expect.arrayContaining(['-', '12', '34', '\u0633']));
    expect(new Set(arabic.draws.map(draw => draw.lang))).toEqual(new Set(['inherit']));
    const inherited = recordingCanvas();
    await renderDocumentToCanvas(documentOf([run('\u0633', { cs: true, ...style })]), inherited.canvas, 0, { dpr: 1, width: 600 });
    expect(inherited.draws.map(draw => draw.lang)).toEqual(['ar-sa']);
  });

  it('draws script-class runs in their class language, cluster seams included', async () => {
    const { canvas, draws } = recordingCanvas();
    // OpenDocument runs: the three axes are the style's own values; the
    // script-class language names each run's language.
    const axes = { lang: 'fr-fr', langEastAsia: 'ja-jp', langBidi: 'ar-sa' };
    await renderDocumentToCanvas(documentOf([
      run('a\u{1d11e}', { ...axes, langScriptClass: 'fr-FR' }),
      run('\u0301\u5b57', { ...axes, langScriptClass: 'ja-JP' }),
      run('x', { ...axes, langScriptClass: '' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.map(draw => draw.text).join('')).toBe('a\u{1d11e}\u0301\u5b57x');
    // U+1D11E and the mark completing its cluster are French (§17.3.2.26
    // font slots still delimit shaped spans, as inside one run); the class
    // without a language draws untagged.
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([
      ['a', 'fr-fr'], ['\u{1d11e}', 'fr-fr'], ['\u0301', 'fr-fr'], ['\u5b57', 'ja-jp'], ['x', 'inherit'],
    ]);
  });

  it('classes digits by the w:bidi language of their own range', async () => {
    const drawn = async (runs: DocxTextRun[]) => {
      const { canvas, draws } = recordingCanvas();
      await renderDocumentToCanvas(documentOf(runs), canvas, 0, { dpr: 1, width: 600 });
      return draws.map(draw => draw.text).sort();
    };
    // Digits of a complex-script run are AN under an RTL w:bidi language (segmented
    // into digit groups) and EN otherwise; runs differing only in w:bidi are
    // one sequence, and each part keeps its own class.
    const arabic = run('12-34', { cs: true, langBidi: 'ar-sa' });
    const english = run('56-78', { cs: true, langBidi: 'en-us' });
    expect(await drawn([arabic, english]))
      .toEqual([...await drawn([arabic]), ...await drawn([english])].sort());
    expect(await drawn([arabic])).toEqual(['-', '12', '34']);
    expect(await drawn([english])).toEqual(['56-78']);
  });

  it('keeps Arabic joining across runs whose unused w:val differs', async () => {
    const drawn = async (secondLang: string) => {
      const { canvas, draws } = recordingCanvas();
      await renderDocumentToCanvas(documentOf([
        run('\u0633\u0644', { cs: true, lang: 'en-us', langBidi: 'ar-sa' }),
        run('\u0627\u0645', { cs: true, lang: secondLang, langBidi: 'ar-sa' }),
      ]), canvas, 0, { dpr: 1, width: 600 });
      return draws.map(draw => draw.text);
    };
    const same = await drawn('en-us');
    expect(same).toEqual(['\u0633\u0644\u0627\u0645']);
    expect(await drawn('fr-fr')).toEqual(same);
  });

  it('keeps Arabic joining across script-class tags that differ only in spelling', async () => {
    const drawn = async (secondClass: string) => {
      const { canvas, draws } = recordingCanvas();
      await renderDocumentToCanvas(documentOf([
        run('\u0633\u0644', { lang: 'ar-sa', langBidi: 'ar-sa', langScriptClass: 'ar-SA' }),
        run('\u0627\u0645', { lang: 'ar-sa', langBidi: 'ar-sa', langScriptClass: secondClass }),
      ]), canvas, 0, { dpr: 1, width: 600 });
      return draws.map(draw => draw.text);
    };
    expect(await drawn('ar-SA')).toEqual(['\u0633\u0644\u0627\u0645']);
    expect(await drawn('AR-sa')).toEqual(['\u0633\u0644\u0627\u0645']);
  });

  it('measures overlay slices in the run language', () => {
    const calls: Array<[string, string | undefined]> = [];
    const measureForFont = (font: string, lang?: string) => {
      calls.push([font, lang]);
      return (s: string) => s.length;
    };
    const base = { text: 'где', x: 0, y: 0, w: 1, h: 1, fontSize: 10, font: '10px A' };
    runOverlayMeasure({ ...base, lang: 'sr-rs' }, measureForFont);
    runOverlayMeasure({ ...base, lang: 'sr-rs', referenceFont: { font: '64px A', scale: 0.25 } }, measureForFont);
    runOverlayMeasure(base, measureForFont);
    expect(calls).toEqual([['10px A', 'sr-rs'], ['64px A', 'sr-rs'], ['10px A', undefined]]);
  });

  it('calls the measurer factory with the font alone for untagged runs', () => {
    const argumentCounts: number[] = [];
    const measureForFont = function (_font: string, _lang?: string) {
      argumentCounts.push(arguments.length);
      return (s: string) => s.length;
    };
    const base = { text: 'abc', x: 0, y: 0, w: 1, h: 1, fontSize: 10, font: '10px A' };
    runOverlayMeasure(base, measureForFont);
    runOverlayMeasure({ ...base, referenceFont: { font: '64px A', scale: 0.25 } }, measureForFont);
    runOverlayMeasure({ ...base, lang: 'sr-rs' }, measureForFont);
    expect(argumentCounts).toEqual([1, 1, 2]);
  });

  it('measures every fragment of a wrapped piece in the language it is painted with', async () => {
    const { canvas, draws, measures } = recordingCanvas();
    // The piece contains text no named family draws, so it is untagged as a
    // whole; its fragments must not regain the language when re-measured.
    await renderDocumentToCanvas(documentOf([
      run('гдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдегдеж', { lang: 'sr-rs' }),
    ]), canvas, 0, { dpr: 1, width: 200 });
    expect(draws.length).toBeGreaterThan(1);
    expect(new Set(draws.map(draw => draw.lang))).toEqual(new Set(['inherit']));
    expect(measures.filter(measure => measure.text.includes('г')).every(measure => measure.lang === 'inherit')).toBe(true);
  });

  it('resolves fragments of text without a language for their own text, as before languages', async () => {
    const keys = async (extra: Record<string, unknown>) => {
      const { canvas, draws } = recordingCanvas();
      const stringify = JSON.stringify;
      let pieceFaces = 0;
      const spy = vi.spyOn(JSON, 'stringify').mockImplementation((...args: Parameters<typeof JSON.stringify>) => {
        const out = stringify(...args);
        if (out?.includes('"faceText"')) pieceFaces += 1;
        return out;
      });
      try {
        await renderDocumentToCanvas(documentOf([
          run('где'.repeat(90), extra),
        ]), canvas, 0, { dpr: 1, width: 200 });
      } finally {
        spy.mockRestore();
      }
      expect(draws.length).toBeGreaterThan(1);
      return pieceFaces;
    };
    // Only a piece that retains a route language resolves its fragments for
    // the whole piece; untagged text has no such request or key.
    expect(await keys({})).toBe(0);
    expect(await keys({ lang: 'x-none' })).toBe(0);
    expect(await keys({ lang: 'sr-rs' })).toBeGreaterThan(0);
  });

  it('keeps the measured language on text drawn inside a save/restore pair', async () => {
    const { canvas, ctx, draws } = recordingCanvas();
    // Chrome resets the active state's lang on save(); restore() brings back
    // the saved one (core saveCanvasState).
    const stack: string[] = [];
    ctx.save = function (this: { lang: string }) { stack.push(this.lang); this.lang = 'inherit'; };
    ctx.restore = function (this: { lang: string }) { this.lang = stack.pop() ?? this.lang; };
    await renderDocumentToCanvas(documentOf([
      run('где', { lang: 'sr-rs', charScale: 150 }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    // Scaled glyphs are drawn one by one, each inside its own save/restore.
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([['г', 'sr-rs'], ['д', 'sr-rs'], ['е', 'sr-rs']]);
    expect(ctx.lang).toBe('inherit');
  });
  it('plans reference-em advances in a caller context\'s ambient language', async () => {
    // A caller canvas set to ja-jp whose double grid-fits small sizes only in
    // that language (HTML canvas text styles: lang is caller state).
    const planned = async (extra: Record<string, unknown>) => {
      const { canvas, ctx } = recordingCanvas();
      const target = ctx as unknown as Record<PropertyKey, unknown>;
      let lang = 'ja-jp';
      const accesses: string[] = [];
      const measuredFonts: string[] = [];
      Object.defineProperty(target, 'lang', {
        get: () => { accesses.push('get'); return lang; },
        set: (value: string) => { accesses.push('set'); lang = value; },
      });
      // The double's own recording of drawn languages is not an access.
      target.fillText = () => {};
      target.measureText = (s: string) => {
        const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(ctx.font)?.[1] ?? String(FONT_PX));
        if (s === 'ab') measuredFonts.push(ctx.font);
        const half = lang === 'ja-jp' && p < 64 ? Math.floor(p / 2 + 0.5) : p / 2;
        return {
          width: [...s].length * half,
          fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
          actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
        } as TextMetrics;
      };
      const proxy = new Proxy(target, {
        has(object, key) {
          if (key === 'lang') accesses.push('has');
          return Reflect.has(object, key);
        },
      });
      (canvas as unknown as { getContext: () => unknown }).getContext = () => proxy;
      // 8.25pt = 11px: grid-fitted in the ambient ja-jp state.
      await renderDocumentToCanvas(documentOf([run('ab', { fontSize: 8.25, italic: false, ...extra })]),
        canvas, 0, { dpr: 1, width: 600 });
      expect(lang).toBe('ja-jp');
      return { accesses: accesses.length, measuredFonts };
    };
    const untagged = await planned({});
    // The reference em (core reference-advances.ts) is chosen, as before
    // languages, with no access to the caller's language.
    expect(untagged.measuredFonts.length).toBeGreaterThan(0);
    expect(untagged.measuredFonts.every(font => font.includes('256px'))).toBe(true);
    expect(untagged.accesses).toBe(0);
    // A run language does not change the per-font decision.
    const tagged = await planned({ lang: 'sr-rs' });
    expect(tagged.measuredFonts.length).toBeGreaterThan(0);
    expect(tagged.measuredFonts.every(font => font.includes('256px'))).toBe(true);
  });

  it('paints untagged text with plain save() and no Canvas language access', async () => {
    const paintAccesses = async (extra: Record<string, unknown>) => {
      const { canvas, ctx } = recordingCanvas();
      const events: string[] = [];
      const target = ctx as unknown as Record<PropertyKey, unknown>;
      target.fillText = () => { events.push('fill'); };
      target.save = () => { events.push('save'); };
      const proxy = new Proxy(target, {
        has(object, key) {
          if (key === 'lang') events.push('lang');
          return Reflect.has(object, key);
        },
        get(object, key) {
          if (key === 'lang') events.push('lang');
          return Reflect.get(object, key);
        },
        set(object, key, value) {
          if (key === 'lang') events.push('lang');
          return Reflect.set(object, key, value);
        },
      });
      (canvas as unknown as { getContext: () => unknown }).getContext = () => proxy;
      // Scaled glyphs are drawn one by one, each inside its own save().
      await renderDocumentToCanvas(documentOf([
        run('где', { charScale: 150, ...extra }),
      ]), canvas, 0, { dpr: 1, width: 600 });
      const painted = events.slice(events.indexOf('fill') === -1 ? events.length : events.lastIndexOf('save', events.indexOf('fill')));
      expect(painted.filter(event => event === 'fill')).toHaveLength(3);
      return painted.filter(event => event === 'lang').length;
    };
    expect(await paintAccesses({})).toBe(0);
    expect(await paintAccesses({ lang: 'x-none' })).toBe(0);
    // The tagged control: the route language is selected and kept per save.
    expect(await paintAccesses({ lang: 'sr-rs' })).toBeGreaterThan(0);
  });

  it('does not merge AutoFit intrinsic pieces painted in different language states', () => {
    const context: ParagraphLayoutContext = {
      lineGrid: { active: false, pitchPt: null },
      characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
      rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
      physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
      lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
      isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false, hasEastAsianText: false,
      kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() }, defaultTabPt: 36,
    };
    const maxContent = (runs: DocxTextRun[]) => {
      const { ctx } = recordingCanvas();
      const document = documentOf(runs);
      return measureParagraphIntrinsicWidths(
        document.body[0] as unknown as Parameters<typeof measureParagraphIntrinsicWidths>[0], context, 1000,
        { context: ctx as unknown as CanvasRenderingContext2D, fontFamilyClasses: {} }, {
          pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
          documentHasEastAsianText: false, compat: compatProfile({ compatibilityMode: 15 }),
          layoutServices: createLayoutServices(document, { measureContext: ctx as unknown as CanvasRenderingContext2D }),
        },
      ).maxWidthPt;
    };
    const serbian = 0.4 * FONT_PX;
    const other = 0.5 * FONT_PX;
    // One language: one merged string, as before.
    expect(maxContent([run('гдегде', { lang: 'sr-rs' }), run('гдегде', { lang: 'sr-rs' })])).toBeCloseTo(12 * serbian, 6);
    // Serbian then Russian: each piece is measured in the language it paints in.
    expect(maxContent([run('гдегде', { lang: 'sr-rs' }), run('гдегде', { lang: 'ru-ru' })]))
      .toBeCloseTo(6 * serbian + 6 * other, 6);
    // Accepted then withheld (text no named family draws) in one language.
    expect(maxContent([run('гдегде', { lang: 'sr-rs', bold: true }), run('гдеж', { lang: 'sr-rs' })]))
      .toBeCloseTo(6 * serbian + 4 * other, 6);
    // Tagged-but-withheld next to untagged text is one state: still merged.
    expect(maxContent([run('гдеж', { lang: 'sr-rs' }), run('гдеж')])).toBeCloseTo(8 * other, 6);
    // A ligature across a language seam: 'ж' (highAnsi, drawn only by a
    // fallback) and 'f'/'i' (ascii) are separate slot spans. Different
    // authored languages are separate sequences, so 'f' and 'i' are shaped
    // apart, each in its own admitted language, and never form 'fi'.
    expect(maxContent([run('жf', { lang: 'sr-rs' }), run('iж', { lang: 'ru-ru' })]))
      .toBeCloseTo(other + serbian + other + other, 6);
    // One language: one sequence, one 'fi' span with its ligature.
    expect(maxContent([run('жf', { lang: 'sr-rs' }), run('iж', { lang: 'sr-rs' })]))
      .toBeCloseTo(other + 1.5 * serbian + other, 6);
    // Both languages withheld (no named family draws the route): the two
    // sequences are still shaped apart, as paint draws them.
    const fallback = { fontFamily: 'Fallback Face' };
    expect(maxContent([run('f', { lang: 'sr-rs', ...fallback }), run('i', { lang: 'ru-ru', ...fallback })]))
      .toBeCloseTo(2 * other, 6);
    expect(maxContent([run('f', { lang: 'sr-rs', ...fallback }), run('i', { lang: 'sr-rs', ...fallback })]))
      .toBeCloseTo(1.5 * other, 6);
  });

  it('declines a kashida whose elongated text the route language was not admitted for', async () => {
    const beh = '\u0628';
    const paragraphOf = (extra: Record<string, unknown>) => {
      const document = documentOf([run(`${beh.repeat(4)} ${beh.repeat(4)} ${'w'.repeat(28)}`, { italic: false, ...extra })]);
      (document.body[0] as DocParagraph).alignment = 'lowKashida' as DocParagraph['alignment'];
      (document.section as SectionProps).pageWidth = 30 * 0.5 * FONT_PX;
      return document;
    };
    const tatweels = async (extra: Record<string, unknown>) => {
      const { canvas, draws } = recordingCanvas();
      await renderDocumentToCanvas(paragraphOf(extra), canvas, 0, { dpr: 1, width: 30 * 0.5 * FONT_PX });
      return draws.filter(draw => draw.text.includes('\u0640')).map(draw => draw.lang);
    };
    // Untagged text is elongated as before.
    expect((await tatweels({})).length).toBeGreaterThan(0);
    // The language was admitted for the authored text. Named families of this
    // route do not draw U+0640, so the elongation is declined rather than
    // drawn by a fallback face in that language.
    expect(await tatweels({ lang: 'ar-sa' })).toEqual([]);
    // Where they draw it, the elongated text keeps the piece's language.
    const kept = await tatweels({ lang: 'ar-sa', fontFamily: 'Kashida Face' });
    expect(kept.length).toBeGreaterThan(0);
    expect(new Set(kept)).toEqual(new Set(['ar-sa']));
  });

  it('draws a ruby guide without a language over a tagged base', async () => {
    const { canvas, draws } = recordingCanvas();
    await renderDocumentToCanvas(documentOf([run('ab', {
      lang: 'sr-rs',
      ruby: { text: 'гд', fontSizePt: FONT_PX / 2 },
    })]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.filter(draw => draw.text === 'ab').map(draw => draw.lang)).toEqual(['sr-rs']);
    expect([...new Set(draws.filter(draw => /[гд]/u.test(draw.text)).map(draw => draw.lang))]).toEqual(['inherit']);
  });
  it('probes punctuation compression and width balance in the acquired language state', async () => {
    const settings = { characterSpacingControl: 'compressPunctuation', balanceSingleByteDoubleByteWidth: true };
    const rendered = async (runs: DocxTextRun[]) => {
      const { canvas, measures, draws } = recordingCanvas();
      const document = documentOf(runs);
      (document as unknown as { settings: unknown }).settings = settings;
      await renderDocumentToCanvas(document, canvas, 0, { dpr: 1, width: 600 });
      return { measures, draws };
    };
    // The piece is withheld as a whole ('〓'): its punctuation, which a named
    // family does draw, is not probed in the language on its own.
    const withheld = await rendered([run('漢、〓', { langEastAsia: 'ja-jp' })]);
    expect(withheld.measures.filter(measure => measure.lang !== 'inherit')).toEqual([]);
    expect(new Set(withheld.draws.map(draw => draw.lang))).toEqual(new Set(['inherit']));
    // An admitted piece probes it in the piece's language.
    const admitted = await rendered([run('漢、漢', { langEastAsia: 'ja-jp' })]);
    expect(new Set(admitted.measures.filter(measure => measure.text === '、').map(measure => measure.lang)))
      .toEqual(new Set(['ja-jp']));
    // Width balance: the half ideographic cell of a space sequence is probed
    // in the run's own East Asian language, whatever an earlier run cached.
    const gap = (draws: Array<{ text: string; x: number }>, from: string, to: string) =>
      draws.find(draw => draw.text.startsWith(to))!.x - draws.find(draw => draw.text.startsWith(from))!.x;
    const alone = await rendered([run('c  d', { langEastAsia: 'ru-ru' })]);
    const after = await rendered([run('a  b ', { langEastAsia: 'sr-rs' }), run('c  d', { langEastAsia: 'ru-ru' })]);
    expect(gap(after.draws, 'a', 'b')).not.toBeCloseTo(gap(alone.draws, 'c', 'd'), 6);
    expect(gap(after.draws, 'c', 'd')).toBeCloseTo(gap(alone.draws, 'c', 'd'), 6);
    // An automatic w:eastAsia keeps its regional tag but shapes untagged:
    // its half cell is not the explicitly tagged run's, in either order.
    const automatic = { langEastAsia: 'sr-rs', langEastAsiaAutomatic: true } as Partial<DocxTextRun>;
    const explicit = { langEastAsia: 'sr-rs' } as Partial<DocxTextRun>;
    const automaticAlone = gap((await rendered([run('c  d', automatic)])).draws, 'c', 'd');
    const explicitAlone = gap((await rendered([run('a  b', explicit)])).draws, 'a', 'b');
    expect(automaticAlone).not.toBeCloseTo(explicitAlone, 6);
    const explicitFirst = await rendered([run('a  b ', explicit), run('c  d', automatic)]);
    expect(gap(explicitFirst.draws, 'a', 'b')).toBeCloseTo(explicitAlone, 6);
    expect(gap(explicitFirst.draws, 'c', 'd')).toBeCloseTo(automaticAlone, 6);
    const automaticFirst = await rendered([run('c  d ', automatic), run('a  b', explicit)]);
    expect(gap(automaticFirst.draws, 'c', 'd')).toBeCloseTo(automaticAlone, 6);
    expect(gap(automaticFirst.draws, 'a', 'b')).toBeCloseTo(explicitAlone, 6);
  });
});
