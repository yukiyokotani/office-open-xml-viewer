import { describe, expect, it, vi } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';
import {
  buildSegments,
  layoutLines,
  paragraphMarkLineMetrics,
  lineBoxHeight,
  segmentEastAsiaFloorSingleLinePx,
  segmentIntendedSingleLinePx,
  type LayoutTextSeg,
} from './line-layout.js';
import type { ResolvedFontMetric } from './layout/text.js';
import type { DocParagraph, DocRun, DocxDocumentModel } from './types.js';

const context = {
  font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
  measureText: (text: string) => ({
    width: [...text].length * 5,
    actualBoundingBoxAscent: 6,
    actualBoundingBoxDescent: 1,
    fontBoundingBoxAscent: 8,
    fontBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

function services(
  localMetrics?: Readonly<Record<string, ResolvedFontMetric>>,
  measureContext: CanvasRenderingContext2D = context,
  majorFont?: string,
) {
  const empty = { default: null, first: null, even: null };
  // The fake Canvas face inventory must carry the same explicit resource
  // identity as production's loaded FontFace records.
  const identified = localMetrics && Object.fromEntries(Object.entries(localMetrics).map(
    ([key, metric]) => [key, {
      ...metric,
      sourceIdentity: metric.sourceIdentity ?? `test-resource:${key}`,
    }],
  ));
  return createLayoutServices({
    section: {
      pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [], headers: empty, footers: empty, majorFont,
  } as DocxDocumentModel, { measureContext, localMetrics: identified });
}

function verifiedReferenceServices(
  families: readonly { family: string; weight?: number; style?: 'normal' | 'italic' }[],
  localMetrics?: Readonly<Record<string, ResolvedFontMetric>>,
  measureContext: CanvasRenderingContext2D = context,
) {
  const base = services(localMetrics, measureContext);
  const text = createTextLayoutService({
    fonts: createFontResolver(families.map(({ family, weight, style }) => ({
      requestedFamily: family,
      resolvedFamily: family,
      source: 'local' as const,
      resourceIdentity: `office-local:local("${family}")`,
      weight,
      style,
    }))),
    measurer: {
      fingerprint: 'verified-local-reference-fixture',
      measure: (request) => ({
        advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
      }),
    },
    localMetrics,
  });
  return { ...base, text };
}

function textLine(family: string) {
  const layoutServices = verifiedReferenceServices([{ family }]);
  const segments = buildSegments([{
    type: 'text', text: 'A', fontFamily: family, fontFamilyEastAsia: family,
    fontSize: 10, bold: false, italic: false, underline: false,
    strikethrough: false,
  }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
  return {
    segment: segments[0] as LayoutTextSeg,
    line: layoutLines(context, segments, 100, 0, 1)[0]!,
  };
}

function substitutedMeiryoService(metric: ResolvedFontMetric, resolvedFamily: string) {
  return createTextLayoutService({
    fonts: createFontResolver([{
      requestedFamily: 'Meiryo', resolvedFamily, source: 'substitute',
    }]),
    measurer: {
      fingerprint: 'substituted-meiryo-test',
      measure: (request) => ({
        advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
      }),
    },
    fontMetrics: { meiryo: metric },
  });
}

describe('native reference font vertical layout', () => {
  it('keeps a loaded application CSS face on its measured line box', () => {
    const loadedFace = {
      family: 'Calibri', weight: '400', style: 'normal', status: 'loaded',
    } as FontFace;
    vi.stubGlobal('document', {
      fonts: { [Symbol.iterator]: function* () { yield loadedFace; } },
    });
    try {
      const layoutServices = services(undefined, context, 'Calibri');
      const selected = layoutServices.text.resolve({
        fonts: { ascii: 'Calibri' }, slot: 'ascii', weight: 400, style: 'normal',
      });
      expect(selected.source).toBe('css');
      const segment = buildSegments([{
        type: 'text', text: 'A', fontFamily: 'Calibri', fontSize: 10,
        bold: false, italic: false, underline: false, strikethrough: false,
      }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices })[0] as LayoutTextSeg;
      expect(segment.referenceFontVerticalMetric).toBeUndefined();
      expect(segment.latinSpaceAverageWidthRatio).toBeUndefined();
      const line = layoutLines(context, [segment], 100, 0, 1)[0]!;
      expect(line.ascent).toBe(8);
      expect(line.descent).toBe(2);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('does not treat an unloaded CSS declaration as a selected face', () => {
    const unloadedFace = {
      family: 'Calibri', weight: '400', style: 'normal', status: 'unloaded',
    } as FontFace;
    vi.stubGlobal('document', {
      fonts: { [Symbol.iterator]: function* () { yield unloadedFace; } },
    });
    try {
      const layoutServices = services(undefined, context, 'Calibri');
      expect(layoutServices.text.resolve({
        fonts: { ascii: 'Calibri' }, slot: 'ascii', weight: 400, style: 'normal',
      }).source).toBe('native');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('uses authored Calibri vertical geometry with a native CSS fallback, without borrowing widths', () => {
    // A CSS fallback does not prove that Calibri paints the glyph. Its catalog
    // sides still define Word pagination; glyph advances remain measured.
    const layoutServices = services();
    const segment = buildSegments([{
      type: 'text', text: 'A B', fontFamily: 'Calibri', fontFamilyEastAsia: 'Calibri',
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    }] as DocRun[], {
      pageIndex: 0, totalPages: 1, layoutServices,
      characterSpacingControl: 'compressPunctuation',
    })[0] as LayoutTextSeg;
    expect(segment.referenceFontVerticalMetric).toBe(true);
    expect(segment.resolvedLineHeightRatio).toBe(2500 / 2048);
    expect(segment.latinSpaceAverageWidthRatio).toBeUndefined();
    const line = layoutLines(context, [segment], 100, 0, 1)[0]!;
    expect(line.intendedSingle).toBeCloseTo(10 * 2500 / 2048, 8);
    expect(line.segments[0]?.measuredWidth).toBe(10);
  });

  it('admits Latin xAvg only from the selected covering resource without changing vertical selection', () => {
    const sourceIdentity = 'provided-sfnt:selected';
    const fontMetrics: Record<string, ResolvedFontMetric> = {
      selected: {
        family: 'Registered Face', requestedFamily: 'Authored Face', sourceIdentity,
        lineHeightRatio: 1.2, averageCharWidthRatio: 0.4,
        unicodeRanges: [[32, 32], [65, 66]],
      },
    };
    const text = createTextLayoutService({
      fonts: createFontResolver([{
        requestedFamily: 'Authored Face', resolvedFamily: 'Registered Face',
        source: 'local', resourceIdentity: sourceIdentity,
      }]),
      measurer: {
        fingerprint: 'latin-xavg-identity',
        measure: (request) => ({ advancePt: [...request.text].length * 5,
          ascentPt: 8, descentPt: 2 }),
      },
      fontMetrics,
    });
    const run = {
      type: 'text', text: 'A B', fontFamily: 'Authored Face', fontFamilyEastAsia: null,
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    } as DocRun;
    const segment = (metrics: Record<string, ResolvedFontMetric>, word6 = false) => buildSegments([run], {
      pageIndex: 0, totalPages: 1, characterSpacingControl: 'compressPunctuation',
      lineWrapLikeWord6: word6,
      layoutServices: { ...services(), text: { ...text, fontMetrics: metrics } },
    })[0] as LayoutTextSeg;
    expect(segment(fontMetrics).latinSpaceAverageWidthRatio).toBe(0.4);
    expect(segment(fontMetrics, true).latinSpaceAverageWidthRatio).toBeUndefined();
    const conflict = {
      ...fontMetrics,
      overlapping: { ...fontMetrics.selected, sourceIdentity: 'provided-sfnt:other',
        averageCharWidthRatio: 0.6 },
    };
    expect(segment(conflict).latinSpaceAverageWidthRatio).toBeUndefined();
    expect(segment(conflict).resolvedLineHeightRatio).toBe(1.2);
    const widthOnly = { selected: { ...fontMetrics.selected, lineHeightRatio: undefined } };
    expect(segment(widthOnly).resolvedLineHeightRatio).toBeUndefined();
    expect(segment(widthOnly).latinSpaceAverageWidthRatio).toBe(0.4);
    expect(segment({ selected: { ...fontMetrics.selected,
      sourceIdentity: 'provided-sfnt:unrelated' } }).latinSpaceAverageWidthRatio)
      .toBeUndefined();
  });
  it('keeps an untabled East Asian grid count independent of the native Canvas box', () => {
    const nativeContext = {
      font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
      measureText(this: { font: string }, text: string) {
        const size = Number(this.font.match(/([\d.]+)px/u)?.[1] ?? 10);
        const native = this.font.includes('Uncatalogued Native');
        return {
          width: [...text].length * (native ? 7 : 5),
          actualBoundingBoxAscent: native ? 6 : 5,
          actualBoundingBoxDescent: 1,
          fontBoundingBoxAscent: size * (native ? 0.85 : 0.9),
          fontBoundingBoxDescent: size * (native ? 0.15 : 0.25),
        };
      },
    } as unknown as CanvasRenderingContext2D;
    const layoutServices = services(undefined, nativeContext);
    const segments = buildSegments([{
      type: 'text', text: '資料を確認します。',
      fontFamily: 'Uncatalogued Native', fontFamilyEastAsia: 'Uncatalogued Native',
      fontSize: 14, bold: true, italic: false, underline: false,
      strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
    const line = layoutLines(nativeContext, segments, 200, 0, 1)[0]!;

    // The 1.0em native Canvas box is valid for ordinary automatic line height,
    // but cannot establish the unavailable face's Word Far-East grid design.
    expect(line.intendedSingle).toBe(14);
    expect(line.gridCountSingle).toBeCloseTo(14 * 1.3, 8);
    expect(lineBoxHeight(null, line.ascent, line.descent, 1,
      { linePitchPt: 18, type: 'lines' }, false,
      line.intendedSingle, true, line.gridCountSingle)).toBe(36);

    // An empty mark using the same selected bold face must reserve the same
    // Far-East grid cells as the visible run, even with a 1.0em Canvas box.
    const paragraph = {
      runs: [], defaultFontFamily: 'Uncatalogued Native',
      defaultFontFamilyEastAsia: 'Uncatalogued Native',
      defaultFontSize: 14, lineSpacing: null,
    } as unknown as DocParagraph;
    const markShape = {
      fontSizePt: 14,
      fonts: { ascii: 'Uncatalogued Native', highAnsi: 'Uncatalogued Native',
        eastAsia: 'Uncatalogued Native', complexScript: 'Uncatalogued Native' },
      weight: 700, style: 'normal' as const, complexScript: false,
      fontHint: 'eastAsia' as const,
    };
    const mark = paragraphMarkLineMetrics(
      paragraph, 1, { linePitchPt: 18, type: 'lines' }, false, true,
      nativeContext, {}, null, layoutServices.text.localMetrics,
      layoutServices.text, markShape,
    );
    expect(mark.advancePx).toBe(36);
    expect(mark.advancePx).toBe(lineBoxHeight(null, line.ascent, line.descent, 1,
      { linePitchPt: 18, type: 'lines' }, false,
      line.intendedSingle, true, line.gridCountSingle));
    expect(paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, true,
      nativeContext, {}, null, layoutServices.text.localMetrics,
      layoutServices.text, markShape,
    ).advancePx).toBe(14);
  });

  it('recovers a native face line floor from a large Canvas probe when no catalog entry exists', () => {
    const quantizedContext = {
      font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
      measureText(this: { font: string }, text: string) {
        const size = Number(this.font.match(/([\d.]+)px/u)?.[1] ?? 10);
        const native = this.font.includes('Uncatalogued Native');
        // A 14px Canvas box loses almost half a pixel of the source face's
        // 1.396em design line. The same selected face at 1000px retains it.
        const ascent = native ? Math.round(size * 0.884) : Math.round(size * 0.9);
        const descent = native ? Math.round(size * 0.513) : Math.round(size * 0.25);
        return {
          width: text.length * (native ? 7 : 5),
          actualBoundingBoxAscent: native ? 5 : 4,
          actualBoundingBoxDescent: native ? 2 : 1,
          fontBoundingBoxAscent: ascent,
          fontBoundingBoxDescent: descent,
        };
      },
    } as unknown as CanvasRenderingContext2D;
    const layoutServices = services(undefined, quantizedContext);
    const segments = buildSegments([{
      type: 'text', text: 'ABC', fontFamily: 'Uncatalogued Native',
      fontFamilyEastAsia: null, fontSize: 14, bold: false, italic: false,
      underline: false, strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
    const line = layoutLines(quantizedContext, segments, 200, 0, 1)[0]!;
    expect(line.intendedSingle).toBeCloseTo(14 * 1.397, 8);
    expect(lineBoxHeight({ rule: 'auto', value: 276 / 240 },
      line.ascent, line.descent, 1, undefined, false, line.intendedSingle))
      .toBeCloseTo(14 * 1.397 * 276 / 240, 8);

    const mark = { runs: [], defaultFontFamily: 'Uncatalogued Native',
      defaultFontSize: 14, lineSpacing: null } as unknown as DocParagraph;
    expect(paragraphMarkLineMetrics(mark, 1, undefined, false, false,
      quantizedContext, {}, null, layoutServices.text.fontMetrics, layoutServices.text)
      .advancePx).toBeCloseTo(14 * 1.397, 8);

    // A positively loaded exact local face admits reference geometry even
    // when its Canvas box is rounded differently at this size.
    const verifiedTimes = verifiedReferenceServices(
      [{ family: 'Times New Roman' }], undefined, quantizedContext,
    );
    const known = buildSegments([{
      type: 'text', text: 'A', fontFamily: 'Times New Roman', fontSize: 14,
      bold: false, italic: false, underline: false, strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices: verifiedTimes });
    expect(layoutLines(quantizedContext, known, 200, 0, 1)[0]?.intendedSingle)
      .toBeCloseTo(14 * referenceFontLineMetrics('Times New Roman')!.lineHeightRatio, 8);
  });
  it('keeps Japanese rows separated using the verified BIZ UDMincho normal-line box', () => {
    for (const bold of [false, true]) {
      const layoutServices = verifiedReferenceServices([
        { family: 'BIZ UDMincho', weight: bold ? 700 : 400 },
      ]);
      const segments = buildSegments([{
        type: 'text', text: '資料を確認します。',
        fontFamily: 'BIZ UDMincho', fontFamilyEastAsia: 'BIZ UDMincho',
        fontSize: 10.5, bold, italic: false, underline: false,
        strikethrough: false,
      }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
      const segment = segments[0] as LayoutTextSeg;
      const line = layoutLines(context, segments, 200, 0, 1)[0]!;
      expect(segment.referenceFontVerticalMetric).toBe(true);
      expect(segment.resolvedLineHeightRatio).toBeCloseTo(1.3, 8);
      expect(line.ascent + line.descent).toBeCloseTo(13.65, 8);
      expect(line.segments[0]?.measuredWidth).toBe(45);
    }
    expect(referenceFontLineMetrics('BIZ UDMincho', 400, 'italic')).toBeUndefined();
  });

  it('uses positively loaded Calibri reference sides without changing its measured width route', () => {
    const { segment, line } = textLine('Calibri');
    expect(segment.referenceFontVerticalMetric).toBe(true);
    expect(segment.resolvedLineHeightRatio).toBe(2500 / 2048);
    expect(line.segments[0]?.measuredWidth).toBe(5);
    expect(line.ascent).toBeCloseTo(10 * 1950 / 2048, 8);
    expect(line.descent).toBeCloseTo(10 * 550 / 2048, 8);
  });

  it('retains measured sides for a family absent from the generated catalog', () => {
    const { segment, line } = textLine('Uncatalogued Browser Face');
    expect(segment.referenceFontVerticalMetric).toBeUndefined();
    expect(line.ascent).toBe(8);
    expect(line.descent).toBe(2);
  });

  it('uses the selected exact local reference for an empty paragraph mark', () => {
    const layoutServices = verifiedReferenceServices([{ family: 'Calibri' }]);
    const paragraph = {
      runs: [], defaultFontFamily: 'Calibri', defaultFontSize: 10,
      lineSpacing: null,
    } as unknown as DocParagraph;
    const measured = paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, false, context, {}, null,
      layoutServices.text.fontMetrics, layoutServices.text,
    );
    expect(measured.ascentPx).toBeCloseTo(10 * 1950 / 2048, 8);
    expect(measured.descentPx).toBeCloseTo(10 * 550 / 2048, 8);
    expect(measured.advancePx).toBeCloseTo(10 * 2500 / 2048, 8);
  });

  it('keeps a loaded East Asian floor under reference geometry authority', () => {
    const layoutServices = verifiedReferenceServices([{ family: 'Meiryo' }]);
    const segments = buildSegments([{
      type: 'text', text: 'A', fontFamily: 'Calibri', fontFamilyEastAsia: 'Meiryo',
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false, textBoxLineFloor: true,
    }] as unknown as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
    const segment = segments[0] as LayoutTextSeg;

    // Meiryo advertises a Far-East code page, so Word allocates 1.3 times
    // its hhea box rather than using the signed hhea line gap.
    expect(segment.resolvedEaFloorLineHeightRatio).toBeCloseTo(1.3 * 3072 / 2048, 12);
    expect(segmentEastAsiaFloorSingleLinePx(segment, 10, true)).toBeCloseTo(10 * 1.3 * 3072 / 2048, 8);
  });

  it('uses a selected East Asian subset resource for a textbox floor even when the run is Latin', () => {
    const metric: ResolvedFontMetric = {
      family: 'Embedded CJK', requestedFamily: 'Embedded CJK',
      weight: 400, style: 'normal', sourceIdentity: 'embedded:font-part',
      lineHeightRatio: 1.5, unicodeRanges: [[0x3042, 0x3042]],
    };
    const text = createTextLayoutService({
      fonts: createFontResolver([{
        requestedFamily: 'Embedded CJK', resolvedFamily: 'Embedded CJK', source: 'embedded',
      }]),
      measurer: {
        fingerprint: 'textbox-east-asian-subset',
        measure: (request) => ({
          advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
        }),
      },
      fontMetrics: { 'embedded cjk': metric },
    });
    const layoutServices = { ...services(), text };
    const segments = buildSegments([{
      type: 'text', text: 'ABC', fontFamily: 'Uncatalogued Latin',
      fontFamilyEastAsia: 'Embedded CJK', fontSize: 10,
      bold: false, italic: false, underline: false, strikethrough: false,
      textBoxLineFloor: true,
    }] as unknown as DocRun[], {
      pageIndex: 0, totalPages: 1, layoutServices,
    });
    const segment = segments[0] as LayoutTextSeg;
    expect(segment.resolvedLineHeightRatio).toBeUndefined();
    expect(segment.resolvedEaFloorLineHeightRatio).toBe(1.5);
    expect(segmentEastAsiaFloorSingleLinePx(segment, 10, false)).toBe(15);
  });

  it('declines conflicting geometry and unproven fallback for shared CSS faces', () => {
    const aliases = {
      'alias a': {
        family: 'Shared Face', requestedFamily: 'Alias A',
        sourceIdentity: 'provided-sfnt:a', lineHeightRatio: 1.1,
        unicodeRanges: [[0x41, 0x41], [0x78, 0x78]],
      },
      'alias b': {
        family: 'Shared Face', requestedFamily: 'Alias B',
        sourceIdentity: 'provided-sfnt:b', lineHeightRatio: 1.5,
        unicodeRanges: [[0x41, 0x41], [0x78, 0x78]],
      },
    } satisfies Record<string, ResolvedFontMetric>;
    const text = createTextLayoutService({
      fonts: createFontResolver([{
        requestedFamily: 'Alias B', resolvedFamily: 'Shared Face', source: 'local',
        resourceIdentity: 'provided-sfnt:b',
      }]),
      measurer: {
        fingerprint: 'shared-face-alias',
        measure: (request) => ({
          advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
        }),
      },
      fontMetrics: aliases,
    });
    const run = {
      type: 'text', text: 'A', fontFamily: 'Alias B', fontFamilyEastAsia: null,
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    } as DocRun;
    const environment = { pageIndex: 0, totalPages: 1, layoutServices: { ...services(), text } };
    // The two resources share a Canvas family/style and both cover A. The
    // authored alias does not prove which FontFace Canvas selected.
    expect((buildSegments([run], environment)[0] as LayoutTextSeg).resolvedLineHeightRatio).toBeUndefined();
    const mark = { runs: [], defaultFontFamily: 'Alias B', defaultFontSize: 10,
      lineSpacing: null } as unknown as DocParagraph;
    expect(paragraphMarkLineMetrics(mark, 1, undefined, false, false,
      context, {}, null, aliases, text).advancePx).toBe(10);
    const equalGeometry = {
      ...aliases,
      'alias b': { ...aliases['alias b'], lineHeightRatio: 1.1 },
    };
    expect((buildSegments([run], {
      ...environment,
      layoutServices: {
        ...environment.layoutServices,
        text: { ...text, fontMetrics: equalGeometry },
      },
    })[0] as LayoutTextSeg).resolvedLineHeightRatio).toBe(1.1);
    expect(paragraphMarkLineMetrics(mark, 1, undefined, false, false,
      context, {}, null, equalGeometry, { ...text, fontMetrics: equalGeometry }).advancePx).toBe(11);
    const uncovered = createTextLayoutService({
      fonts: createFontResolver([{
        requestedFamily: 'Alias B', resolvedFamily: 'Shared Face', source: 'local',
        resourceIdentity: 'provided-sfnt:b',
      }]),
      measurer: {
        fingerprint: 'shared-face-alias-uncovered',
        measure: (request) => ({
          advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
        }),
      },
      fontMetrics: {
        ...aliases,
        'alias b': { ...aliases['alias b'], unicodeRanges: [[0x42, 0x42]] },
      },
    });
    expect((buildSegments([run], {
      ...environment, layoutServices: { ...environment.layoutServices, text: uncovered },
    })[0] as LayoutTextSeg).resolvedLineHeightRatio).toBeUndefined();
  });

  it('does not borrow geometry from a different resource under the selected local CSS tuple', () => {
    const selectedIdentity = 'local("Selected Face")';
    const font = createFontResolver([{
      requestedFamily: 'Authored Face', resolvedFamily: 'Registered Face',
      source: 'local', resourceIdentity: selectedIdentity,
    }]);
    const run = {
      type: 'text', text: 'A', fontFamily: 'Authored Face', fontFamilyEastAsia: null,
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    } as DocRun;
    const shaped = (sourceIdentity?: string) => {
      const metric = {
        family: 'Registered Face', requestedFamily: 'Authored Face',
        lineHeightRatio: 1.5, sourceIdentity,
      };
      const text = createTextLayoutService({
        fonts: font,
        measurer: {
          fingerprint: 'resource-identity',
          measure: () => ({ advancePt: 5, ascentPt: 8, descentPt: 2 }),
        },
        fontMetrics: { face: metric },
      });
      return (buildSegments([run], {
        pageIndex: 0, totalPages: 1,
        layoutServices: { ...services(), text },
      })[0] as LayoutTextSeg).resolvedLineHeightRatio;
    };
    expect(shaped('provided-sfnt:unrelated-face')).toBeUndefined();
    expect(shaped()).toBeUndefined();
    expect(shaped(selectedIdentity)).toBe(1.5);
  });

  it('declines conflicting embedded and local resources under one CSS face tuple', () => {
    const fontMetrics = {
      embedded: {
        family: 'Shared Face', requestedFamily: 'Embedded Alias',
        sourceIdentity: 'embedded:one', lineHeightRatio: 1.2,
        unicodeRanges: [[0x41, 0x41], [0x78, 0x78]],
      },
      local: {
        family: 'Shared Face', requestedFamily: 'Local Alias',
        sourceIdentity: 'provided-sfnt:two', lineHeightRatio: 1.5,
        unicodeRanges: [[0x41, 0x41], [0x78, 0x78]],
      },
    } satisfies Record<string, ResolvedFontMetric>;
    const text = createTextLayoutService({
      fonts: createFontResolver([{
        requestedFamily: 'Embedded Alias', resolvedFamily: 'Shared Face', source: 'embedded',
      }]),
      measurer: {
        fingerprint: 'cross-source-face-collision',
        measure: (request) => ({
          advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
        }),
      },
      fontMetrics,
    });
    const run = {
      type: 'text', text: 'A', fontFamily: 'Embedded Alias', fontFamilyEastAsia: null,
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    } as DocRun;
    expect((buildSegments([run], {
      pageIndex: 0, totalPages: 1,
      layoutServices: { ...services(), text },
    })[0] as LayoutTextSeg).resolvedLineHeightRatio).toBeUndefined();
    const mark = { runs: [], defaultFontFamily: 'Embedded Alias', defaultFontSize: 10,
      lineSpacing: null } as unknown as DocParagraph;
    expect(paragraphMarkLineMetrics(mark, 1, undefined, false, false,
      context, {}, null, fontMetrics, text).advancePx).toBe(10);
  });

  it('rebuilds the temporary metric index after a mutable service changes between paragraphs', () => {
    const mutableMetrics: Record<string, ResolvedFontMetric> = {
      face: {
        family: 'Selected Face', requestedFamily: 'Authored Face',
        sourceIdentity: 'provided-sfnt:one', lineHeightRatio: 1.2,
        unicodeRanges: [[0x41, 0x41]],
      },
    };
    const baseText = createTextLayoutService({
      fonts: createFontResolver([{
        requestedFamily: 'Authored Face', resolvedFamily: 'Selected Face', source: 'local',
        resourceIdentity: 'provided-sfnt:one',
      }]),
      measurer: {
        fingerprint: 'mutable-font-metric-service',
        measure: (request) => ({
          advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2,
        }),
      },
    });
    const layoutServices = { ...services(), text: { ...baseText, fontMetrics: mutableMetrics } };
    const run = {
      type: 'text', text: 'A', fontFamily: 'Authored Face', fontFamilyEastAsia: null,
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    } as DocRun;
    const environment = { pageIndex: 0, totalPages: 1, layoutServices };
    expect((buildSegments([run], environment)[0] as LayoutTextSeg).resolvedLineHeightRatio).toBe(1.2);
    mutableMetrics.face = { ...mutableMetrics.face, lineHeightRatio: 1.6 };
    expect((buildSegments([run], environment)[0] as LayoutTextSeg).resolvedLineHeightRatio).toBe(1.6);
  });

  it('keeps a resolved local East Asian resource ahead of reference lookup', () => {
    const layoutServices = services({
      meiryo: { family: '__embedded_meiryo', lineHeightRatio: 1.2 },
    });
    const segments = buildSegments([{
      type: 'text', text: 'A', fontFamily: 'Calibri', fontFamilyEastAsia: 'Meiryo',
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false, textBoxLineFloor: true,
    }] as unknown as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
    const segment = segments[0] as LayoutTextSeg;

    expect(segment.resolvedEaFloorLineHeightRatio).toBe(1.2);
    expect(segmentEastAsiaFloorSingleLinePx(segment, 10, true)).toBeCloseTo(12, 8);
  });

  it('keeps exact resource geometry ahead of family compatibility metrics in a text line', () => {
    const resource = { meiryo: { family: 'Meiryo', lineHeightRatio: 1.2 } };
    const layoutServices = services(resource);
    const segments = buildSegments([{
      type: 'text', text: 'A', fontFamily: 'Meiryo', fontFamilyEastAsia: 'Meiryo',
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
    const segment = segments[0] as LayoutTextSeg;
    const line = layoutLines(context, segments, 100, 0, 1)[0]!;

    expect(segment.resolvedLineHeightRatio).toBe(1.2);
    expect(segmentIntendedSingleLinePx(segment, 10)).toBe(12);
    expect(line.ascent).toBe(8);
    expect(line.descent).toBe(2);

    const tallContext = {
      ...context,
      measureText: () => ({
        width: 5, fontBoundingBoxAscent: 15, fontBoundingBoxDescent: 7,
      }),
    } as unknown as CanvasRenderingContext2D;
    const directSegments = buildSegments([{
      type: 'text', text: 'A', fontFamily: 'Meiryo', fontFamilyEastAsia: 'Meiryo',
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    }] as DocRun[], {
      pageIndex: 0, totalPages: 1,
      resolvedLocalFonts: { meiryo: { family: 'Meiryo', lineHeightRatio: 1.2 } },
    });
    const tallLine = layoutLines(tallContext, directSegments, 100, 0, 1)[0]!;
    expect(tallLine.ascent).toBe(15);
    expect(tallLine.descent).toBe(7);
  });

  it('carries selected resource ascent and descent through the layout snapshot into text and marks', () => {
    const resource = {
      family: 'Meiryo', lineHeightRatio: 1.3,
      designAscentRatio: 1.05, designDescentRatio: 0.25,
    };
    const layoutServices = services({ meiryo: resource });
    const segments = buildSegments([{
      type: 'text', text: 'A', fontFamily: 'Meiryo', fontFamilyEastAsia: 'Meiryo',
      fontSize: 10, bold: false, italic: false, underline: false,
      strikethrough: false,
    }] as DocRun[], { pageIndex: 0, totalPages: 1, layoutServices });
    const segment = segments[0] as LayoutTextSeg;
    const line = layoutLines(context, segments, 100, 0, 1)[0]!;
    expect(segment.resolvedResourceVerticalMetric).toBe(true);
    expect(segment.resolvedDesignAscentRatio).toBe(1.05);
    expect(segment.resolvedDesignDescentRatio).toBe(0.25);
    expect(line.ascent).toBeCloseTo(10.5, 8);
    expect(line.descent).toBeCloseTo(2.5, 8);

    const mark = paragraphMarkLineMetrics({
      runs: [], defaultFontFamily: 'Meiryo', defaultFontSize: 10,
      lineSpacing: null,
    } as unknown as DocParagraph, 1, undefined, false, false,
    context, {}, null, layoutServices.text.localMetrics, layoutServices.text);
    expect(mark.ascentPx).toBeCloseTo(10.5, 8);
    expect(mark.descentPx).toBeCloseTo(2.5, 8);
  });

  it('keeps the normal font line box in atLeast spacing for text and a paragraph mark', () => {
      const spacing = { rule: 'atLeast' as const, value: 9, explicit: true };
      const resource = {
        meiryo: {
          family: 'Meiryo', lineHeightRatio: 1.3,
          designAscentRatio: 1.05, designDescentRatio: 0.25,
          sourceIdentity: 'office-local:local("Meiryo")',
        },
      };
      const layoutServices = verifiedReferenceServices([
        { family: 'Calibri' }, { family: 'Meiryo' },
      ], resource);
      for (const [family, expectedRatio] of [['Calibri', 2500 / 2048], ['Meiryo', 1.3]] as const) {
        const segments = buildSegments([{
          type: 'text', text: 'A', fontFamily: family, fontFamilyEastAsia: family,
          fontSize: 10, bold: false, italic: false, underline: false,
          strikethrough: false,
        }] as DocRun[], {
          pageIndex: 0, totalPages: 1, layoutServices,
          lineSpacing: spacing,
        });
        const segment = segments[0] as LayoutTextSeg;
        expect(segment.resolvedLineHeightRatio).toBeCloseTo(expectedRatio, 8);
        expect(segment.resolvedDesignAscentRatio).toBeDefined();
        expect(lineBoxHeight(spacing, 8, 2, 1, undefined, false,
          segmentIntendedSingleLinePx(segment, 10))).toBeCloseTo(expectedRatio * 10, 8);
      }
      const mark = paragraphMarkLineMetrics({
        runs: [], defaultFontFamily: 'Meiryo', defaultFontSize: 10,
        lineSpacing: spacing,
      } as unknown as DocParagraph, 1, undefined, false, false,
      context, {}, spacing, layoutServices.text.localMetrics, layoutServices.text);
      expect(mark.ascentPx).toBeCloseTo(10.5, 8);
      expect(mark.descentPx).toBeCloseTo(2.5, 8);
      expect(mark.advancePx).toBeCloseTo(13, 8);
  });

  it('keeps the Word-observed Calibri normal line until an atLeast floor exceeds it', () => {
    const metric = referenceFontLineMetrics('Calibri', 400, 'normal', 'other');
    expect(metric).toBeDefined();
    const natural = 11 * metric!.lineHeightRatio;
    const minimum = (twips: number) => lineBoxHeight(
      { rule: 'atLeast', value: twips / 20, explicit: true },
      8, 3, 1, undefined, false, natural,
    );
    // Controlled Word-for-Mac PDF: 11pt Calibri lines are 13.44pt at
    // atLeast 0–269 twips; 360/480 twips produce 18/24pt. The catalog
    // predicts 13.4277pt before Word's page-position quantization.
    expect(minimum(0)).toBeCloseTo(13.44, 1);
    expect(minimum(269)).toBeCloseTo(13.44, 1);
    expect(minimum(360)).toBe(18);
    expect(minimum(480)).toBe(24);
  });

  it('preserves intended natural height above an explicit Latin grid pitch', () => {
    expect(lineBoxHeight(
      { rule: 'auto', value: 1, explicit: true },
      8, 2, 1, { type: 'lines', linePitchPt: 12 }, false, 13, false,
    )).toBe(13);
  });

  it('keeps Word automatic OpenType projection out of exact spacing', () => {
      const spacing = { rule: 'exact' as const, value: 9, explicit: true };
      const layoutServices = services({
        meiryo: {
          family: 'Meiryo', lineHeightRatio: 1.3,
          designAscentRatio: 1.05, designDescentRatio: 0.25,
        },
      });
      for (const family of ['Calibri', 'Meiryo']) {
        const segments = buildSegments([{
          type: 'text', text: 'A', fontFamily: family, fontFamilyEastAsia: family,
          fontSize: 10, bold: false, italic: false, underline: false,
          strikethrough: false,
        }] as DocRun[], {
          pageIndex: 0, totalPages: 1, layoutServices,
          lineSpacing: spacing,
        });
        const segment = segments[0] as LayoutTextSeg;
        const line = layoutLines(context, segments, 100, 0, 1)[0]!;
        expect(segment.resolvedLineHeightRatio).toBeUndefined();
        expect(segment.resolvedDesignAscentRatio).toBeUndefined();
        expect(line.ascent).toBe(8);
        expect(line.descent).toBe(2);
      }
      const mark = paragraphMarkLineMetrics({
        runs: [], defaultFontFamily: 'Meiryo', defaultFontSize: 10,
        lineSpacing: spacing,
      } as unknown as DocParagraph, 1, undefined, false, false,
      context, {}, spacing, layoutServices.text.localMetrics, layoutServices.text);
      expect(mark.ascentPx).toBe(8);
      expect(mark.descentPx).toBe(2);
      expect(mark.advancePx).toBe(9);
  });

  it('keeps exact resource geometry ahead of family compatibility metrics for an empty mark', () => {
    const paragraph = {
      runs: [], defaultFontFamily: 'Meiryo', defaultFontSize: 10,
      lineSpacing: null,
    } as unknown as DocParagraph;
    const resource = { meiryo: { family: 'Meiryo', lineHeightRatio: 1.2 } };
    const layoutServices = services(resource);
    const mark = paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, false, context, {}, null,
      layoutServices.text.localMetrics, layoutServices.text,
    );

    expect(mark.ascentPx).toBe(8);
    expect(mark.descentPx).toBe(2);
    expect(mark.advancePx).toBe(12);

    const tallContext = {
      ...context,
      measureText: () => ({
        width: 5, fontBoundingBoxAscent: 15, fontBoundingBoxDescent: 7,
      }),
    } as unknown as CanvasRenderingContext2D;
    const tallServices = services(resource, tallContext);
    const tallMark = paragraphMarkLineMetrics(
      paragraph, 1, undefined, false, false, tallContext, {}, null,
      tallServices.text.localMetrics, tallServices.text,
    );
    expect(tallMark.ascentPx).toBe(15);
    expect(tallMark.descentPx).toBe(7);
  });

  it.each(['Fallback Face', 'Meiryo'])(
    'keeps authored vertical geometry while refusing an unrelated resource metric for substitute %s',
    (resolvedFamily) => {
      const metric = { family: 'Meiryo', lineHeightRatio: 1.2 };
      const text = substitutedMeiryoService(metric, resolvedFamily);
      const layoutServices = { ...services(), text };
      const segments = buildSegments([{
        type: 'text', text: 'A', fontFamily: 'Meiryo', fontFamilyEastAsia: 'Meiryo',
        fontSize: 10, bold: false, italic: false, underline: false,
        strikethrough: false, textBoxLineFloor: true,
      }] as unknown as DocRun[], {
        pageIndex: 0, totalPages: 1, layoutServices,
        resolvedLocalFonts: { meiryo: metric },
      });
      const segment = segments[0] as LayoutTextSeg;
      expect(segment.fontFamily).toBe(resolvedFamily);
      const authoredRatio = referenceFontLineMetrics('Meiryo')!.lineHeightRatio;
      expect(segment.resolvedLineHeightRatio).toBe(authoredRatio);
      expect(segment.resolvedEaFloorLineHeightRatio).toBe(authoredRatio);
      expect(segment.resolvedLineHeightRatio).not.toBe(metric.lineHeightRatio);

      const paragraph = {
        runs: [], defaultFontFamily: 'Meiryo', defaultFontSize: 10,
        lineSpacing: null,
      } as unknown as DocParagraph;
      const mark = paragraphMarkLineMetrics(
        paragraph, 1, undefined, false, false, context, {}, null,
        { meiryo: metric }, text,
      );
      expect(mark.advancePx).toBeCloseTo(10 * authoredRatio, 8);
    },
  );
});
