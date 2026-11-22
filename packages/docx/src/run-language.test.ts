import { describe, expect, it, vi } from 'vitest';

// Node has no Canvas realm to run the route sentinels in; production judges
// named-family ownership there (core namedFamiliesDrawText). Rendering tests
// below use faces their recording canvas draws itself.
// 'ж' stands for text only a generic fallback draws.
vi.mock('@silurus/ooxml-core', async (importOriginal) => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  createNamedFamilyOracle: () => ({
    available: true,
    draws: (_list: string, _w: number, _s: string, text: string) => !text.includes('ж'),
  }),
}));
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService, rubyGuideShapeRequest, sliceTextShapeRequest, type GlyphMeasureRequest } from './layout/text.js';
import { renderDocumentToCanvas } from './renderer.js';
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
        const em = request.lang?.startsWith('sr') ? 4 : 5;
        return { advancePt: [...request.text].length * em, ascentPt: 8, descentPt: 2 };
      },
    },
  });
  const fonts = { ascii: 'Face', highAnsi: 'Face', eastAsia: 'Face', complexScript: 'Face' };

  it('measures each slot with its own language and reports it on the span', () => {
    measured.length = 0;
    const languages = { language: 'sr-latn-rs', eastAsiaLanguage: 'ja-jp', complexScriptLanguage: 'ar-sa' };
    const mixed = service.shape({ text: 'ж中', fontSizePt: 10, fonts, clusterGeometry: false, ...languages });
    // §17.3.2.7 w:cs puts the whole run on the complex-script slot.
    const cs = service.shape({
      text: 'بa', fontSizePt: 10, fonts, clusterGeometry: false, complexScript: true, ...languages,
    });
    const kana = service.shape({ text: 'か', fontSizePt: 10, fonts, clusterGeometry: false, ...languages });
    expect([...mixed.spans, ...cs.spans, ...kana.spans].map((span) => [span.script, span.lang])).toEqual([
      // Generic fallback draws '中': its language would choose the face.
      ['highAnsi', 'sr-latn-rs'], ['eastAsia', undefined], ['complexScript', 'ar-sa'], ['eastAsia', 'ja-jp'],
    ]);
    expect(measured.map((request) => [request.text, request.lang])).toEqual([
      ['ж', 'sr-latn-rs'], ['中', undefined], ['بa', 'ar-sa'], ['か', 'ja-jp'],
    ]);
    expect('lang' in mixed.spans[1]!).toBe(false);
  });

  it('keys measurement by language only when one is set', () => {
    measured.length = 0;
    const serbian = service.shape({ text: 'бгд', fontSizePt: 10, fonts, language: 'sr-rs', clusterGeometry: false });
    const plain = service.shape({ text: 'бгд', fontSizePt: 10, fonts, clusterGeometry: false });
    // Word's "x-none" names no language: same measurement as untagged text.
    const none = service.shape({ text: 'бгд', fontSizePt: 10, fonts, language: 'x-none', clusterGeometry: false });
    expect([serbian.advancePt, plain.advancePt, none.advancePt]).toEqual([12, 15, 15]);
    expect(measured.map((request) => 'lang' in request ? request.lang : 'absent')).toEqual(['sr-rs', 'absent']);
    expect('lang' in plain.spans[0]!).toBe(false);
  });

  it('sets no language without named-family ownership evidence', () => {
    const unjudged = createTextLayoutService({
      fonts: createFontResolver([{ requestedFamily: 'Face', resolvedFamily: 'Face', source: 'local' }]),
      measurer: { fingerprint: 'unjudged', measure: request => {
        expect('lang' in request).toBe(false);
        return { advancePt: 5, ascentPt: 8, descentPt: 2 };
      } },
    });
    const shaped = unjudged.shape({ text: 'б', fontSizePt: 10, fonts, language: 'sr-rs', clusterGeometry: false });
    expect('lang' in shaped.spans[0]!).toBe(false);
  });

  it('reuses a piece\'s pinned language in every derived request of that slot group', () => {
    const withheld = { text: 'бгд', fontSizePt: 10, fonts, language: 'sr-rs', clusterGeometry: false as const,
      canvasLanguagePin: { group: 'latin' as const, lang: null } };
    expect('lang' in service.shape(withheld).spans[0]!).toBe(false);
    expect('lang' in service.shape(sliceTextShapeRequest(withheld, 0, 1)).spans[0]!).toBe(false);
    // An accepted language holds for derived text the proof would refuse
    // (a kashida or probe glyph drawn by fallback): paint draws it that way.
    const accepted = { ...withheld, text: 'бж', canvasLanguagePin: { group: 'latin' as const, lang: 'sr-rs' } };
    expect(service.shape(accepted).spans.map(span => span.lang)).toEqual(['sr-rs']);
    // Another slot group is decided on its own.
    expect(service.shape({ ...accepted, text: 'か', eastAsiaLanguage: 'ja-jp' }).spans.map(span => span.lang))
      .toEqual(['ja-jp']);
  });

  it('shapes a ruby guide in its own languages', () => {
    const base = { text: '漢', fontSizePt: 10, fonts, language: 'en-us', eastAsiaLanguage: 'zh-cn',
      canvasLanguagePin: { group: 'eastAsia' as const, lang: null } };
    const guide = rubyGuideShapeRequest(base, 'かん', { guide: 'ja-jp', latin: 'en-gb' });
    expect([guide.language, guide.eastAsiaLanguage, guide.complexScriptLanguage, guide.canvasLanguagePin])
      .toEqual(['en-gb', 'ja-jp', 'ja-jp', undefined]);
    const inherited = rubyGuideShapeRequest(base, 'かん', undefined);
    expect([inherited.language, inherited.eastAsiaLanguage, inherited.canvasLanguagePin])
      .toEqual(['en-us', 'zh-cn', undefined]);
  });

  it('keeps independent languages when scoped Arabic shares a physical face across slots', () => {
    const scoped = createTextLayoutService({
      fonts: createFontResolver([
        { requestedFamily: 'Scoped Face', resolvedFamily: 'Arabic Substitute', source: 'substitute', script: 'arabic' },
      ], { scriptScopedFamilies: { 'scoped face': { script: 'arabic', substituteFamilies: ['Arabic Substitute'] } } }),
      measurer: { fingerprint: 'scoped-language', namedFamiliesDraw: () => true, measure: request => ({
        advancePt: [...request.text].length * (request.lang === 'ar' ? 4 : 5), ascentPt: 8, descentPt: 2,
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
    expect(split.spans.map(span => [span.text, span.lang])).toEqual([['ا', 'ar'], ['\u{1EE00}', 'ur']]);
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
        return {
          width: [...s].length * em,
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
    // w:val stay one cluster; disagreeing, they carry no language.
    await renderDocumentToCanvas(documentOf([
      run('\u25fc', { lang: 'fr-ch' }),
      run('\ufe0e', { lang: 'en-us' }),
      run('ab', { lang: 'en-us' }),
    ]), canvas, 0, { dpr: 1, width: 600 });
    expect(draws.filter(draw => draw.text.includes('\u25fc')).map(draw => [draw.text, draw.lang]))
      .toEqual([['\u25fc\ufe0e', 'inherit']]);
    expect(draws.filter(draw => draw.text.includes('ab')).map(draw => [draw.text, draw.lang]))
      .toEqual([['ab', 'en-us']]);
  });

  it('joins runs differing only in w:val inside a grapheme cluster and gives them no language', async () => {
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
    // The recording canvas draws one grapheme at a time. Runs joined to keep
    // a cluster whole disagree on w:val, so none of them is drawn in another
    // run's language; the next run keeps its own.
    expect(draws.map(draw => [draw.text, draw.lang])).toEqual([
      ['\u{1f469}', 'inherit'], ['\u200d', 'inherit'], ['\u{1f467}', 'inherit'], [' ', 'inherit'],
      // 'x' and 'a' share a run language but join the accent's sequence.
      ['xa', 'inherit'], ['\u0301', 'inherit'], ['b', 'inherit'], ['c', 'en-us'],
    ]);
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
});
