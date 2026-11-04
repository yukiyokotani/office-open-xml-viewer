import { describe, expect, it } from 'vitest';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService, type GlyphMeasureRequest } from './layout/text.js';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, SectionProps } from './types';

// ECMA-376 §17.3.2.20 w:lang — w:val governs text that is neither East Asian
// nor complex script, w:eastAsia and w:bidi the other two §17.3.2.26 slots.
// The run language reaches Canvas `lang`, which selects language-specific
// OpenType forms (locl), e.g. Serbian italic Cyrillic (#2001).

describe('text service: §17.3.2.20 language per font slot', () => {
  const measured: GlyphMeasureRequest[] = [];
  const service = createTextLayoutService({
    fonts: createFontResolver([
      { requestedFamily: 'Face', resolvedFamily: 'Face', source: 'local' },
    ]),
    measurer: {
      fingerprint: 'run-language',
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
    expect([...mixed.spans, ...cs.spans].map((span) => [span.script, span.lang])).toEqual([
      ['highAnsi', 'sr-latn-rs'], ['eastAsia', 'ja-jp'], ['complexScript', 'ar-sa'],
    ]);
    expect(measured.map((request) => [request.text, request.lang])).toEqual([
      ['ж', 'sr-latn-rs'], ['中', 'ja-jp'], ['بa', 'ar-sa'],
    ]);
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
});
