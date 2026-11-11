import { describe, expect, it } from 'vitest';
import { symbolCodeSet, symbolTextToUnicodeSegments } from '@silurus/ooxml-core';
import type { TextLayoutService } from './text.js';
import { resolveNumberingMarkerGeometry } from './numbering-marker.js';
import { createTextLayoutService } from './text.js';
import { createFontResolver } from './font-service.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { makeRecordingCanvas } from '../test-support/recording-canvas.test-support.js';
import type { DocxDocumentModel } from '../types.js';

const route = {
  familyList: 'serif', scope: 'generic', fingerprint: 'numbering-marker-test-route',
} as const;

function textService(advancePt: number): TextLayoutService {
  return {
    fingerprint: 'numbering-marker-test-text',
    localMetrics: {},
    resolve(request) {
      return {
        requestedFamily: request.fonts.ascii ?? 'serif',
        resolvedFamily: request.fonts.ascii ?? 'serif',
        route,
        source: 'generic',
        weight: request.weight ?? 400,
        style: request.style ?? 'normal',
        diagnostics: [],
        genericFamily: 'serif',
      };
    },
    shape(request) {
      return {
        text: request.text,
        spans: [{
          text: request.text,
          start: 0,
          end: request.text.length,
          script: 'eastAsia',
          breakBefore: true,
          font: {
            requestedFamily: 'serif', resolvedFamily: 'serif', route,
            source: 'generic', weight: request.weight ?? 400,
            style: request.style ?? 'normal', diagnostics: [], genericFamily: 'serif',
          },
          fontRoute: route,
          advancePt,
          ascentPt: 8,
          descentPt: 2,
        }],
        advancePt,
        ascentPt: 8,
        descentPt: 2,
        graphemeBoundaries: [0, request.text.length],
        clusters: [{
          range: { start: 0, end: request.text.length }, offsetPt: 0, advancePt,
        }],
        diagnostics: [],
      };
    },
  };
}

describe('resolveNumberingMarkerGeometry', () => {
  // Deliberately synthetic advances: the service's Arial substitute has a
  // 0.3125em space, distinct from both Courier and any fitted Arial fraction.
  const suffixService = createTextLayoutService({
    fonts: createFontResolver([
      { requestedFamily: 'Arial', resolvedFamily: 'Suffix Sans', source: 'substitute' },
      { requestedFamily: 'Courier New', resolvedFamily: 'Courier New', source: 'local' },
    ]),
    measurer: { fingerprint: 'suffix-font-advances', measure(request) {
      const em = request.fontRoute.familyList.includes('Suffix Sans') ? 0.3125 : 0.625;
      return { advancePt: request.text.length * request.fontSizePt * em,
        ascentPt: request.fontSizePt * 0.8, descentPt: request.fontSizePt * 0.2 };
    } },
  });

  // #2177 review: a converted marker is Unicode text. Shaping it with the
  // marker's Symbol family would let a symbol-encoded Symbol re-read U+00A9
  // as its own code 0xA9 (a heart), so its slots are cleared. A code the face
  // draws keeps the Symbol slots.
  it.each([
    ['lacks', [] as number[], '©\uFE0E', null],
    ['draws', [0xd3], '\uF0D3', 'Symbol'],
  ])('shapes a Symbol marker the face %s with the matching slots', (_case, owned, text, ascii) => {
    const requests: { text: string; ascii: string | null | undefined }[] = [];
    const base = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: { fingerprint: 'marker-symbol', measure: (request) => ({
        advancePt: request.text.length * 6, ascentPt: 8, descentPt: 2,
      }) },
      symbolText: (markerText, family) => symbolTextToUnicodeSegments(markerText, family, {
        encoding: { kind: 'known', pua: symbolCodeSet(owned), bare: symbolCodeSet(owned) },
        monochrome: () => true,
      }),
    });
    const service: TextLayoutService = { ...base, shape(request) {
      requests.push({ text: request.text, ascii: request.fonts.ascii });
      return base.shape(request);
    } };
    const geometry = resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'bullet', text: '\uF0D3', suff: 'tab', indentLeft: 0, tab: 36, fontFamily: 'Symbol' },
      { fontSizePt: 12, fonts: { ascii: 'Symbol', highAnsi: 'Symbol', eastAsia: 'Symbol', complexScript: 'Symbol' },
        weight: 400, style: 'normal', complexScript: false },
      { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0, tabStops: [], defaultTabPt: 36 }, service,
    );
    expect(geometry.markerText).toBe(text);
    expect(requests.find((request) => request.text === text)?.ascii).toBe(ascii);
  });

  it('shapes a partly native Symbol marker piece by piece, in order', () => {
    const requests: { text: string; ascii: string | null | undefined }[] = [];
    const base = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: { fingerprint: 'marker-symbol-mixed', measure: (request) => ({
        advancePt: request.text.length * 6, ascentPt: 8, descentPt: 2,
      }) },
      symbolText: (markerText, family) => symbolTextToUnicodeSegments(markerText, family, {
        encoding: { kind: 'known', pua: symbolCodeSet([0x61]), bare: symbolCodeSet([0x61]) },
      }),
    });
    const service: TextLayoutService = { ...base, shape(request) {
      requests.push({ text: request.text, ascii: request.fonts.ascii });
      return base.shape(request);
    } };
    const geometry = resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'bullet', text: '\uF061\uF062', suff: 'tab', indentLeft: 0, tab: 36, fontFamily: 'Symbol' },
      { fontSizePt: 12, fonts: { ascii: 'Symbol', highAnsi: 'Symbol', eastAsia: 'Symbol', complexScript: 'Symbol' },
        weight: 400, style: 'normal', complexScript: false },
      { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0, tabStops: [], defaultTabPt: 36, baseRtl: false }, service,
    );
    expect(geometry.markerText).toBe('\uF061\u03b2');
    expect(requests.slice(0, 2)).toEqual([{ text: '\uF061', ascii: 'Symbol' }, { text: '\u03b2', ascii: null }]);
    expect(geometry.shape?.spans.map((span) => [span.text, span.start, span.end])).toEqual([['\uF061', 0, 1], ['\u03b2', 1, 2]]);
    expect(geometry.markerWidthPt).toBe(geometry.shape?.advancePt);
  });

  it.each([
    [undefined, 3.75], [14, 3.75], [15, 3.75],
    [11, 7.5], [12, 7.5], [13, 7.5],
  ])('uses the suffix font advance at marker size in mode %s', (compatibilityMode, suffixPt) => {
    const input = { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0,
      tabStops: [], defaultTabPt: 36, compatibilityMode, baseRtl: false };
    const geometry = resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'decimal', text: '1.', suff: 'space', indentLeft: 0, tab: 36 },
      { fontSizePt: 12, fonts: { ascii: 'Courier New', eastAsia: 'Courier New', complexScript: 'Courier New' },
        weight: 400, style: 'normal', complexScript: false },
      input, suffixService,
    );
    expect(geometry.markerWidthPt).toBe(15);
    expect(geometry.bodyOffsetPt - geometry.markerWidthPt).toBe(suffixPt);
  });

  // WORD_NUMBERING_SPACE_SUFFIX_ARIAL is measured for regular Latin / East
  // Asian markers in left-to-right paragraphs only: every other class keeps
  // the previous space in the marker's own face (7.5pt here) byte-identically.
  it.each([
    ['a right-to-left paragraph', { baseRtl: true }, {}, {}],
    ['an unknown paragraph direction', { baseRtl: undefined }, {}, {}],
    ['a bold marker', {}, { weight: 700 }, {}],
    ['an italic marker', {}, { style: 'italic' as const }, {}],
    ['a complex-script marker', {}, { complexScript: true }, {}],
    ['a Private Use Area symbol marker', {}, {}, { text: '\uF0B7', fontFamily: 'Symbol' }],
    ['a remapped symbol-font marker', {}, {}, { text: '\u00B7', fontFamily: 'Symbol' }],
  ])('keeps the marker-face space for %s (unmeasured)', (_label, inputPatch, markerPatch, numberingPatch) => {
    const geometry = resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'decimal', text: '1.', suff: 'space', indentLeft: 0, tab: 36, ...numberingPatch },
      { fontSizePt: 12, fonts: { ascii: 'Courier New', eastAsia: 'Courier New', complexScript: 'Courier New' },
        weight: 400, style: 'normal', complexScript: false, ...markerPatch },
      { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0, tabStops: [], defaultTabPt: 36,
        compatibilityMode: 15, baseRtl: false, ...inputPatch },
      suffixService,
    );
    expect(geometry.bodyOffsetPt - geometry.markerWidthPt).toBe(7.5);
  });

  it.each([
    [15, false, 18, 3.75],
    [15, false, 24, 3.75],
    [12, false, 18, 7.5],
    [15, true, 18, 7.5],
    [12, true, 18, 7.5],
  ] as const)('retains the marker-sized suffix gap through layout (mode %s, RTL %s, run %spt)',
    (compatibilityMode, bidi, runSize, suffixPt) => {
      const { canvas } = makeRecordingCanvas();
      const ctx = canvas.getContext('2d')!;
      const measure = ctx.measureText.bind(ctx);
      ctx.measureText = (text) => ({ ...measure(text), width: [...text].length
        * parseFloat(/([\d.]+)px/.exec(ctx.font)![1]!)
        * (ctx.font.includes('"Arial"') ? 0.3125 : 0.625) });
      const model = {
        section: { pageWidth: 400, pageHeight: 400, marginTop: 0, marginRight: 0,
          marginBottom: 0, marginLeft: 0, headerDistance: 0, footerDistance: 0 },
        settings: { compatibilityMode },
        body: [{ type: 'paragraph', alignment: bidi ? 'right' : 'left', bidi,
          indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
          lineSpacing: null, tabStops: [], defaultFontSize: 16, defaultFontFamily: 'Verdana',
          numbering: { numId: 1, level: 0, format: 'decimal', text: '1.', suff: 'space',
            indentLeft: 0, tab: 36, fontFamily: 'Courier New',
            fontFacts: { fontFamily: 'Courier New', fontSize: 12 } },
          runs: [{ type: 'text', text: 'Body', fontSize: runSize, fontFamily: 'Georgia',
            bold: false, italic: false, underline: false, strikethrough: false }],
          widowControl: false }],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
      } as unknown as DocxDocumentModel;
      const layout = layoutDocument(model, createLayoutServices(model, { measureContext: ctx }));
      const paragraph = layout.pages[0]!.layers.body.find((node) => node.kind === 'paragraph')!;
      if (paragraph.kind !== 'paragraph') throw new Error('Expected a paragraph');
      const placements = paragraph.lines[0]!.placements.filter((node) => node.kind === 'text');
      const marker = placements.find((node) => node.role === 'numbering-marker')!;
      const body = placements.find((node) => node.role !== 'numbering-marker')!;
      expect(marker.fontSizePt).toBe(12);
      const gap = bidi
        ? marker.bounds.xPt - (body.bounds.xPt + body.advancePt)
        : body.bounds.xPt - (marker.bounds.xPt + marker.advancePt);
      expect(gap).toBeCloseTo(suffixPt, 8);
    });

  it('allocates a registered marker from the same precise font box as body text', () => {
    // Canvas rounds both sides independently. Those device-pixel sides must
    // not enlarge a numbered line whose selected resource owns precise sides.
    const serviceWithCoverage = (unicodeRanges?: Array<[number, number]>) => createTextLayoutService({
      fonts: createFontResolver([{ requestedFamily: 'Fixture Serif', resolvedFamily: 'Fixture Serif',
        source: 'embedded', resourceIdentity: 'embedded:marker', weight: 400, style: 'normal' }]),
      fontMetrics: { marker: { family: 'Fixture Serif', sourceIdentity: 'embedded:marker',
        weight: 400, style: 'normal', lineHeightRatio: 1,
        designAscentRatio: .8, designDescentRatio: .2, unicodeRanges } },
      measurer: { fingerprint: 'rounded-marker-box', measure(request) {
        return { advancePt: request.text.length * 5, ascentPt: 9, descentPt: 3 };
      } },
    });
    const service = serviceWithCoverage();
    const acquire = (fontSizePt: number, textService = service) => resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'decimal', text: '1.', suff: 'nothing', indentLeft: 0, tab: 36 },
      { fontSizePt, fonts: { ascii: 'Fixture Serif' }, weight: 400, style: 'normal', complexScript: false },
      { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0, tabStops: [], defaultTabPt: 36 }, textService,
    );
    const sameSize = acquire(10);
    expect(sameSize.lineBox).toEqual({ ascentPt: 8, descentPt: 2, intendedSinglePt: 10 });
    expect(acquire(20).lineBox).toEqual({ ascentPt: 16, descentPt: 4, intendedSinglePt: 20 });
    // Geometry for painting and horizontal fitting still belongs to Canvas.
    expect(sameSize.markerWidthPt).toBe(10);
    expect(sameSize.shape?.ascentPt).toBe(9);
    expect(sameSize.shape?.descentPt).toBe(3);
    // A selected subset lacking marker digits cannot lend its design geometry
    // to Canvas fallback glyphs, even when it covers the following body text.
    expect(acquire(10, serviceWithCoverage([[65, 90]])).lineBox)
      .toEqual({ ascentPt: 9, descentPt: 3, intendedSinglePt: 0 });
  });
  function geometryWithSuffix(suff: 'nothing' | 'space', authoredFirstIndentPt: number) {
    return resolveNumberingMarkerGeometry(
      {
        numId: 1,
        level: 0,
        format: 'decimal',
        text: '1.',
        indentLeft: 36,
        tab: 36,
        suff,
        jc: 'left',
      },
      {
        fontSizePt: 10,
        fonts: { ascii: 'serif', eastAsia: 'serif' },
        weight: 400,
        style: 'normal',
        complexScript: false,
      },
      {
        authoredFirstIndentPt,
        physicalIndentLeftPt: 36,
        tabStops: [],
        defaultTabPt: 36,
      },
      textService(10),
    );
  }

  function geometryAtCoincidentStop(alignment: 'left' | 'num') {
    return resolveNumberingMarkerGeometry(
      {
        numId: 1,
        level: 0,
        format: 'bullet',
        text: '\u30fb',
        indentLeft: 18,
        tab: 18,
        suff: 'tab',
        jc: 'left',
      },
      {
        fontSizePt: 10,
        fonts: { ascii: 'serif', eastAsia: 'serif' },
        weight: 400,
        style: 'normal',
        complexScript: false,
      },
      {
        // Marker interval relative to the 18pt paragraph indent is [-8, 2].
        // Its absolute end is therefore exactly the authored 20pt list tab.
        authoredFirstIndentPt: -8,
        physicalIndentLeftPt: 18,
        tabStops: [{ pos: 20, alignment, leader: 'none' }],
        defaultTabPt: 42.55,
      },
      textService(10),
    );
  }

  it('keeps a suffix tab on the list stop coincident with the marker end', () => {
    const geometry = geometryAtCoincidentStop('num');

    expect(geometry.markerWidthPt).toBe(10);
    expect(geometry.bodyOffsetPt).toBe(2);
  });

  it('keeps ordinary coincident tabs strictly forward', () => {
    expect(geometryAtCoincidentStop('left').bodyOffsetPt).toBeCloseTo(24.55, 10);
  });

  it.each([
    ['nothing', -36],
    ['space', -36],
  ] as const)(
    'keeps a %s suffix body at the paragraph start when its marker ends inside the hanging region',
    (suffix, authoredFirstIndentPt) => {
      expect(geometryWithSuffix(suffix, authoredFirstIndentPt).bodyOffsetPt).toBe(0);
    },
  );

  it.each([
    ['nothing', 2],
    ['space', 12],
  ] as const)(
    'keeps a positive %s suffix advance beyond the paragraph start',
    (suffix, expectedBodyOffsetPt) => {
      expect(geometryWithSuffix(suffix, -8).bodyOffsetPt).toBe(expectedBodyOffsetPt);
    },
  );
});
