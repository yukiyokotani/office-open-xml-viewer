import { describe, expect, it } from 'vitest';
import { symbolTextToUnicodeSegments } from '@silurus/ooxml-core';
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

  // #2177: a symbol marker is judged like run text and shaped as one run.
  // `native` lists the authored codes (bare or U+F0xx) the marker font draws
  // as its own symbol glyphs.
  const symbolMarker = (text: string, native: readonly string[], options: Readonly<{
    complexScriptFace?: string; baseRtl?: boolean;
    /** Families whose face paints text in its own colours. */
    colourFamilies?: readonly string[];
  }> = {}) => {
    const requests: { text: string; ascii: string | null | undefined; cs: string | null | undefined }[] = [];
    const base = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: { fingerprint: 'marker-symbol', measure: (request) => ({
        advancePt: request.text.length * 6, ascentPt: 8, descentPt: 2,
      }) },
      symbolText: (markerText, family) => symbolTextToUnicodeSegments(markerText, family, {
        native: (code) => native.includes(code),
        monochrome: () => true,
      }),
      symbolDraws: (code) => native.includes(code),
      symbolMonochrome: (_text, resolution) => !options.colourFamilies?.includes(resolution.requestedFamily),
    });
    const service: TextLayoutService = { ...base, shape(request) {
      requests.push({ text: request.text, ascii: request.fonts.ascii, cs: request.fonts.complexScript });
      return base.shape(request);
    } };
    const geometry = resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'bullet', text, suff: 'tab', indentLeft: 0, tab: 36, fontFamily: 'Symbol' },
      { fontSizePt: 12, fonts: { ascii: 'Symbol', highAnsi: 'Symbol', eastAsia: 'Symbol',
        complexScript: options.complexScriptFace ?? 'Symbol' },
        weight: 400, style: 'normal', complexScript: options.complexScriptFace !== undefined },
      { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0, tabStops: [], defaultTabPt: 36, baseRtl: options.baseRtl ?? false },
      service,
    );
    return { geometry, requests };
  };

  it('keeps a marker code its font draws, in the symbol slots', () => {
    const { geometry, requests } = symbolMarker('', ['']);
    expect(geometry.markerText).toBe('');
    expect(requests).toEqual([{ text: '', ascii: 'Symbol', cs: 'Symbol' }]);
  });

  it('shapes a kept bare code that bidi mirroring would replace in its F0-shifted form', () => {
    const { geometry, requests } = symbolMarker('\u00AB', ['\u00AB'], { baseRtl: true });
    expect(geometry.markerText).toBe('\uF0AB');
    expect(requests).toEqual([{ text: '\uF0AB', ascii: 'Symbol', cs: 'Symbol' }]);
  });

  it('converts a kept mirrored code whose F0-shifted form shaping would give another slot family', () => {
    // U+00AB takes highAnsi (Symbol) and is native there; U+F0AB would take
    // eastAsia under the hint, another font: it converts instead of staying
    // bare, where right-to-left shaping would mirror it to another glyph.
    const service = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: { fingerprint: 'symbol-shift', measure: (request) => ({ advancePt: request.text.length * 6, ascentPt: 8, descentPt: 2 }) },
      symbolText: (text, family, _resolution, _route, keepsShifted) => symbolTextToUnicodeSegments(text, family, {
        native: () => true, monochrome: () => true, ...(keepsShifted ? { keepsShifted } : {}),
      }),
    });
    const slots = (eastAsia: string) => ({
      text: '«', fonts: { ascii: 'Symbol', highAnsi: 'Symbol', eastAsia, complexScript: 'Symbol' },
      weight: 400, style: 'normal' as const, fontHint: 'eastAsia' as const,
    });
    expect(service.symbolText!(slots('MS Mincho'))).toEqual([{ text: '↔︎', mapped: true }]);
    expect(service.symbolText!(slots('Symbol'))).toEqual([{ text: '', mapped: false }]);
  });

  it('keeps the conversion of a native mirrored marker code whose shifted form would change slot family', () => {
    // As above, for a one-code marker: the code is native, yet it converted;
    // that is not a partly native marker, and the bare code must not return.
    const requests: string[] = [];
    const base = createTextLayoutService({
      fonts: createFontResolver([]),
      measurer: { fingerprint: 'marker-shift', measure: (request) => ({ advancePt: request.text.length * 6, ascentPt: 8, descentPt: 2 }) },
      symbolText: (text, family, _resolution, _route, keepsShifted) => symbolTextToUnicodeSegments(text, family, {
        native: () => true, monochrome: () => true, ...(keepsShifted ? { keepsShifted } : {}),
      }),
      symbolDraws: () => true,
    });
    const service: TextLayoutService = { ...base, shape(request) { requests.push(request.text); return base.shape(request); } };
    const geometry = resolveNumberingMarkerGeometry(
      { numId: 1, level: 0, format: 'bullet', text: '«', suff: 'tab', indentLeft: 0, tab: 36, fontFamily: 'Symbol' },
      { fontSizePt: 12, fonts: { ascii: 'Symbol', highAnsi: 'Symbol', eastAsia: 'MS Mincho', complexScript: 'Symbol' },
        weight: 400, style: 'normal', complexScript: false, fontHint: 'eastAsia' },
      { authoredFirstIndentPt: 0, physicalIndentLeftPt: 0, tabStops: [], defaultTabPt: 36, baseRtl: true }, service,
    );
    expect(geometry.markerText).toBe('↔︎');
    expect(requests).toEqual(['↔︎']);
  });

  it('keeps the marker slots for text converted from a font that draws no symbol code', () => {
    // Apple's Symbol is Unicode-encoded: U+F0B7 becomes "•", which that same
    // face draws, as in 0.89.
    const { geometry, requests } = symbolMarker('', []);
    expect(geometry.markerText).toBe('•');
    expect(requests).toEqual([{ text: '•', ascii: 'Symbol', cs: 'Symbol' }]);
  });

  it('gives up the authored slots where they would paint the converted text in colour', () => {
    // The conversion's colour proof was made in the cleared route.
    const { geometry, requests } = symbolMarker('\uF0B7', [], { colourFamilies: ['Symbol'] });
    expect(geometry.markerText).toBe('\u2022');
    expect(requests).toEqual([{ text: '\u2022', ascii: null, cs: null }]);
  });

  it('clears the slots of converted text a symbol cmap could read as its own codes', () => {
    // Symbol 0xB8 is "÷" (U+00F7). A symbol-encoded face that lacks U+F0B8
    // (say one admitted for U+0000-00FF only) would re-read the converted
    // character as its code 0xF7, and nothing can show that a face reads a
    // Latin-1 character as itself: such text never keeps the symbol slots.
    const { geometry, requests } = symbolMarker('', []);
    expect(geometry.markerText).toBe('÷');
    expect(requests).toEqual([{ text: '÷', ascii: null, cs: null }]);
    // A target above U+00FF next to one below it: cleared as well.
    expect(symbolMarker('', []).requests).toEqual([{ text: '•÷', ascii: null, cs: null }]);
  });

  it('shapes an emoji-capable target in the cleared route its monochrome proof was made in', () => {
    const { geometry, requests } = symbolMarker('', []);
    expect(geometry.markerText).toBe('©︎');
    expect(requests[0]).toEqual({ text: '©︎', ascii: null, cs: null });
  });

  it.each([false, true])('keeps a marker as authored when its font draws only some of its codes (rtl: %s)', (baseRtl) => {
    const { geometry, requests } = symbolMarker('', [''], { baseRtl });
    expect(geometry.markerText).toBe('');
    expect(requests).toEqual([{ text: '', ascii: 'Symbol', cs: 'Symbol' }]);
    expect(geometry.markerWidthPt).toBe(geometry.shape?.advancePt);
  });

  it('keeps the mirror-safe form of the native codes of a partly native marker', () => {
    // The face draws a bare U+00AB but lacks U+F0B7: the marker stays as
    // authored, with the kept code F0-shifted as in body text, so
    // right-to-left shaping cannot mirror it to another glyph.
    const { geometry, requests } = symbolMarker('«', ['«', ''], { baseRtl: true });
    expect(geometry.markerText).toBe('');
    expect(requests).toEqual([{ text: '', ascii: 'Symbol', cs: 'Symbol' }]);
  });

  it('judges a complex-script symbol marker in its symbol family, not its cs text face', () => {
    // The cs slot names a text face; Word still draws the lvlText code in
    // Symbol. A kept code is shaped in that family alone; converted text is
    // ordinary Unicode in the authored slots.
    const kept = symbolMarker('', [''], { complexScriptFace: 'Times New Roman', baseRtl: true });
    expect(kept.geometry.markerText).toBe('');
    expect(kept.requests).toEqual([{ text: '', ascii: 'Symbol', cs: 'Symbol' }]);
    const converted = symbolMarker('', [], { complexScriptFace: 'Times New Roman', baseRtl: true });
    expect(converted.geometry.markerText).toBe('•');
    expect(converted.requests).toEqual([{ text: '•', ascii: 'Symbol', cs: 'Times New Roman' }]);
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
