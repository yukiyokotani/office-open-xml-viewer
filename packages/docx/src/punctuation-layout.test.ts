import { DEFAULT_KINSOKU_RULES, graphemeClusterOffsets } from '@silurus/ooxml-core';
import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  layoutLines,
  punctuationCompressionTotalPt,
  type DocGridCtx,
  type LayoutLine,
  type LayoutTextSeg,
} from './line-layout.js';
import { createFontResolver } from './layout/font-service.js';
import {
  createTextLayoutService,
  type TextLayoutService,
} from './layout/text.js';
import { wordIsOverflowPunctuation } from './layout/line-compatibility.js';
import type { LayoutServices } from './layout/types.js';
import type { DocParagraph, DocxTextRun } from './types.js';
import type { VerticalGlyphMeasurementService } from './layout/measurement-capabilities.js';

const compressionPt = (segment: LayoutTextSeg): number | undefined => {
  const value = punctuationCompressionTotalPt(segment);
  return value === 0 ? undefined : value;
};

const VERTICAL_MEASUREMENT = {
  fingerprint: 'punctuation-layout:test',
  measureRunInkExtra: () => 0,
  planRun: () => [],
} as VerticalGlyphMeasurementService;

function context(): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = () => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    fontKerning: 'auto',
    letterSpacing: '0px',
    measureText: (text: string) => ({
      // Full-width synthetic face: exactly one em per scalar.
      width: [...text].length * size(),
      fontBoundingBoxAscent: size() * 0.8,
      fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8,
      actualBoundingBoxDescent: size() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function textRun(
  text: string,
  eastAsiaLanguage: string | null = 'ja-jp',
): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Synthetic CJK',
    fontFamilyEastAsia: 'Synthetic CJK',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  };
  return {
    type: 'text',
    ...run,
    // Parser-only effective language input consumed by the isolated
    // overflow-punctuation compatibility projection.
    langEastAsia: eastAsiaLanguage ?? undefined,
  } as DocParagraph['runs'][number];
}

function lines(
  segments: ReturnType<typeof buildSegments>,
  width: number,
  overflowPunct: boolean,
  characterGrid?: DocGridCtx,
): LayoutLine[] {
  return layoutLines(
    context(),
    segments,
    width,
    0,
    1,
    [],
    undefined,
    {},
    undefined,
    DEFAULT_KINSOKU_RULES,
    characterGrid,
    36,
    width,
    false,
    false,
    false,
    undefined,
    undefined,
    segments.some((segment) => 'text' in segment && segment.verticalRun)
      ? VERTICAL_MEASUREMENT
      : undefined,
    overflowPunct,
  );
}

const textOf = (line: LayoutLine): string =>
  line.segments
    .filter((segment): segment is LayoutTextSeg => 'text' in segment)
    .map((segment) => segment.text)
    .join('');

function punctuationMetricsServices(options: Readonly<{
  punctuationAdvancePt?: number;
  ideographicCellAdvancePt?: number;
  leadingInkStartPt?: number;
}> = {}): LayoutServices {
  const text: TextLayoutService = createTextLayoutService({
    fonts: createFontResolver([]),
    measurer: {
      fingerprint: 'punctuation-layout:tight-ink',
      measure(request) {
        const scalarCount = [...request.text].length;
        const tightCharacters = new Set([
          '、', '。', '．', '」', '）', '！', '？', '：', '；',
          'あ', 'ア', 'ー', 'か\u3099', 'し', 'な', 'い',
        ]);
        const tightTrailingWhitespace = tightCharacters.has(request.text);
        const graphemeBoundaries = [0, ...graphemeClusterOffsets(request.text), request.text.length];
        const graphemes = graphemeBoundaries.slice(0, -1)
          .map((start, index) => request.text.slice(start, graphemeBoundaries[index + 1]));
        const advancePt = request.text === '\u4e00'
          ? (options.ideographicCellAdvancePt ?? 10)
          : graphemes.reduce((sum, grapheme) =>
              sum + (tightCharacters.has(grapheme)
                ? (options.punctuationAdvancePt ?? 10)
                : 10),
            0);
        return {
          advancePt,
          ascentPt: 8,
          descentPt: 2,
          ...(tightTrailingWhitespace
            ? {
                inkBounds: {
                  xMinPt: 1,
                  xMaxPt: request.text === '．'
                    ? 3
                    : request.text === '、'
                      ? 1
                      : advancePt - 5,
                  ascentPt: 2,
                  descentPt: 0,
                },
                horizontalInkBoundsAreTight: true,
              }
            : {
                inkBounds: {
                  xMinPt: options.leadingInkStartPt ?? 0,
                  xMaxPt: advancePt,
                  ascentPt: 8,
                  descentPt: 2,
                },
                horizontalInkBoundsAreTight: true,
              }),
        };
      },
    },
  });
  return { text } as unknown as LayoutServices;
}

describe('ECMA-376 East-Asian punctuation fit', () => {
  it('does not apply document-level punctuation compression over authored run spacing', () => {
    const authoredText = '独自文面の配置を確認します）。';
    const run = {
      ...textRun(authoredText),
      charSpacing: 0.4,
    } as DocParagraph['runs'][number];
    const segments = buildSegments([run], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments).toHaveLength(1);
    expect(segments[0].punctuationCompressions).toBeUndefined();
    expect(segments[0].charSpacing).toBe(0.4);
    const authoredWidthPt = [...authoredText].length * 10.4;
    expect(lines(segments, authoredWidthPt, false)[0].segments[0].measuredWidth)
      .toBeCloseTo(authoredWidthPt, 5);
  });

  it('retains at least a half-width cell around compressed full-width punctuation', () => {
    const segments = buildSegments([textRun('甲乙．丙')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    });

    expect(segments.map((segment) => 'text' in segment ? segment.text : ''))
      .toEqual(['甲乙．丙']);
    const punctuation = segments[0] as LayoutTextSeg;
    expect(punctuation.charSpacing).toBeUndefined();
    expect(compressionPt(punctuation)).toBe(-5);

    const laidOut = lines(segments, 35, false);
    expect(laidOut).toHaveLength(1);
    expect(textOf(laidOut[0])).toBe('甲乙．丙');
    expect((laidOut[0].segments[0] as LayoutTextSeg).measuredWidth).toBe(35);
  });

  it('does not trim a Japanese comma to its tiny ink bounds', () => {
    const segments = buildSegments([textRun('甲、乙')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments.map((segment) => segment.text)).toEqual(['甲、乙']);
    // The comma ink ends at 1pt in a 10pt cell. Compressing all 9pt of trailing
    // whitespace makes the following ideograph touch the comma; Word's
    // punctuation compression retains the half-em punctuation cell.
    expect(compressionPt(segments[0])).toBe(-5);
    expect(lines(segments, 25, false)).toHaveLength(1);
    expect(lines(segments, 25, false)[0].segments.map((segment) =>
      'text' in segment ? segment.measuredWidth : undefined,
    )).toEqual([25]);
  });

  it.each([
    ['inside one run', [textRun('）次')]],
    ['across a source-run boundary', [textRun('）'), textRun('次')]],
  ])('keeps compressed closing punctuation clear of following-glyph left overhang %s', (
    _case,
    runs,
  ) => {
    const segments = buildSegments(runs, {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices({ leadingInkStartPt: -2 }),
    }) as LayoutTextSeg[];

    const closing = segments[0]!;
    // The closing parenthesis ink ends at 5pt. The next glyph reaches 2pt left
    // of its origin, so its origin must remain at or beyond 7pt. A 10pt
    // full-width cell may therefore be compressed by at most 3pt.
    expect(closing.punctuationCompressions).toEqual([
      { end: 1, adjustmentPt: -3 },
    ]);
  });

  it('does not compress punctuation that the selected face already renders proportionally', () => {
    const segments = buildSegments([textRun('甲、乙')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices({
        punctuationAdvancePt: 8,
        ideographicCellAdvancePt: 10,
      }),
    }) as LayoutTextSeg[];

    // §17.15.1.18 permits compression of full-width punctuation. The selected
    // face has already placed the comma on an 8pt proportional advance while
    // its ideographic cell is 10pt, so applying another trim would double-count
    // the font's own spacing choice.
    expect(compressionPt(segments[0])).toBeUndefined();
    expect(lines(segments, 28, false)[0].segments.map((segment) =>
      'text' in segment ? segment.measuredWidth : undefined,
    )).toEqual([28]);
  });

  it('adds the character-grid pitch once without re-compressing proportional punctuation', () => {
    const segments = buildSegments([textRun('甲、'), textRun('E')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices({
        punctuationAdvancePt: 8,
        ideographicCellAdvancePt: 10,
      }),
    }) as LayoutTextSeg[];

    const [line] = lines(segments, 100, false, {
      type: 'linesAndChars',
      linePitchPt: 15,
      charSpacePt: -0.4,
    });
    expect(line.segments.map((segment) =>
      'text' in segment ? [segment.text, segment.measuredWidth] : undefined,
    )).toEqual([
      ['甲、', 17.2],
      ['E', 9.6],
    ]);
  });

  it('does not invent a fixed trim when tight horizontal ink bounds are unavailable', () => {
    const segments = buildSegments([textRun('甲乙．丙')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    const punctuation = segments[0] as LayoutTextSeg;
    expect(punctuation.text).toBe('甲乙．丙');
    expect(compressionPt(punctuation)).toBeUndefined();
    expect(lines(segments, 35, false)).toHaveLength(2);
  });

  it('scales punctuation sidebearing trim with the authored glyph width', () => {
    const run = {
      ...textRun('．甲'),
      charScale: 0.5,
    } as DocParagraph['runs'][number];
    const [punctuation] = buildSegments([run], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    // The retained half-cell and its 5pt trim are both scaled to 50%.
    expect(compressionPt(punctuation)).toBe(-2.5);
    expect(lines([punctuation], 7.5, false)).toHaveLength(1);
  });

  it('allows one eligible punctuation character past the line extent by default policy', () => {
    const segments = buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
    });

    const hanging = lines(segments, 20, true);
    expect(hanging).toHaveLength(1);
    expect(textOf(hanging[0])).toBe('甲乙．');

    const disabled = lines(buildSegments([textRun('甲乙．')], {
      pageIndex: 0,
      totalPages: 1,
    }), 20, false);
    expect(disabled).toHaveLength(2);
    expect(disabled.map(textOf)).toEqual(['甲', '乙．']);
  });

  it('does not uniformly compress a Japanese middle dot', () => {
    const segments = buildSegments([textRun('甲乙・丙丁')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    expect(segments).toHaveLength(1);
    expect((segments[0] as LayoutTextSeg).text).toBe('甲乙・丙丁');
    expect(compressionPt(segments[0] as LayoutTextSeg)).toBeUndefined();
    const [line] = lines(segments, 100, false);
    expect((line.segments[0] as LayoutTextSeg).measuredWidth).toBe(50);
  });

  it('dispatches exact ST_CharacterSpacing values and compresses kana only in the combined mode', () => {
    const punctuationOnly = buildSegments([textRun('あアー．漢')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(punctuationOnly.map((segment) => segment.text)).toEqual(['あアー．漢']);
    expect(punctuationOnly.map(compressionPt)).toEqual([-5]);

    const punctuationAndKana = buildSegments([textRun('あアー・漢．乙')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationAndJapaneseKana',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(punctuationAndKana.map((segment) => segment.text)).toEqual([
      'あアー・漢．乙',
    ]);
    expect(punctuationAndKana[0].punctuationCompressions).toEqual([
      { end: 1, adjustmentPt: -5 },
      { end: 2, adjustmentPt: -5 },
      { end: 3, adjustmentPt: -5 },
      { end: 6, adjustmentPt: -5 },
    ]);

    const unknownPrefix = buildSegments([textRun('あ．')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationFuture',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(unknownPrefix).toHaveLength(1);
    expect(compressionPt(unknownPrefix[0])).toBeUndefined();
  });

  it('does not compress halfwidth punctuation as full-width punctuation', () => {
    const segments = buildSegments([textRun('｡､')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    }) as LayoutTextSeg[];

    expect(segments).toHaveLength(1);
    expect(segments[0].text).toBe('｡､');
    expect(compressionPt(segments[0])).toBeUndefined();
  });

  it('keeps a decomposed kana base and combining mark in one compressed grapheme', () => {
    const decomposedGa = 'か\u3099';
    const segments = buildSegments([textRun(`甲${decomposedGa}乙`)], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationAndJapaneseKana',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments.map((segment) => segment.text)).toEqual([`甲${decomposedGa}乙`]);
    expect(compressionPt(segments[0])).toBe(-5);
    expect(segments.every((segment) => segment.text !== '\u3099')).toBe(true);
  });

  it('keeps full-width exclamation and question marks on their authored cell', () => {
    const segments = buildSegments([textRun('！甲？乙：丙；丁')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments.map((segment) => segment.text)).toEqual([
      '！甲？乙：丙；丁',
    ]);
    expect(segments[0].punctuationCompressions).toBeUndefined();
  });

  it('applies characterSpacingControl from selected glyph metrics to MS Mincho', () => {
    const minchoRun = {
      ...textRun('甲、乙）丙'),
      fontFamily: 'MS Mincho',
      fontFamilyEastAsia: 'MS Mincho',
    } as DocParagraph['runs'][number];
    const segments = buildSegments([minchoRun], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments.map((segment) => segment.text)).toEqual(['甲、乙）丙']);
    expect(segments[0].punctuationCompressions).toEqual([
      { end: 2, adjustmentPt: -5 },
      { end: 4, adjustmentPt: -5 },
    ]);
    expect(lines(segments, 50, false)[0].segments.reduce(
      (sum, segment) => sum + segment.measuredWidth,
      0,
    )).toBe(40);
  });

  it('keeps U+3017 and full-width !/? uncompressed while compressing the observed punctuation subset', () => {
    const segments = buildSegments([textRun('、。〗！？」')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];

    expect(segments).toHaveLength(1);
    expect(segments[0].punctuationCompressions).toEqual([
      { end: 1, adjustmentPt: -5 },
      { end: 2, adjustmentPt: -5 },
      { end: 6, adjustmentPt: -5 },
    ]);
  });

  it('retains consecutive punctuation and kana as per-cluster adjustments in one shape', () => {
    const punctuation = buildSegments([textRun('文章。」次')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(punctuation).toHaveLength(1);
    expect(punctuation[0].text).toBe('文章。」次');
    expect(punctuation[0].punctuationCompressions).toEqual([
      { end: 3, adjustmentPt: -5 },
      { end: 4, adjustmentPt: -5 },
    ]);
    expect(lines(punctuation, 40, false).map(textOf)).toEqual(['文章。」次']);

    const kana = buildSegments([textRun('希望しない）')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuationAndJapaneseKana',
      layoutServices: punctuationMetricsServices(),
    }) as LayoutTextSeg[];
    expect(kana).toHaveLength(1);
    expect(kana[0].text).toBe('希望しない）');
    expect(kana[0].punctuationCompressions).toEqual([
      { end: 3, adjustmentPt: -5 },
      { end: 4, adjustmentPt: -5 },
      { end: 5, adjustmentPt: -5 },
      { end: 6, adjustmentPt: -5 },
    ]);
  });

  it('preserves cross-segment kinsoku ownership at a compressed punctuation boundary', () => {
    const segments = buildSegments([textRun('一二三四五六、七')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
    });

    expect(segments.map((segment) => 'text' in segment ? segment.text : ''))
      .toEqual(['一二三四五六、七']);
    const wrapped = lines(segments, 60, false);
    expect(wrapped.map(textOf)).toEqual(['一二三四五', '六、七']);
  });

  it('keeps the one-time punctuation trim on the wrapped tail, never the prefix', () => {
    const segments = buildSegments([textRun('甲乙丙）')], {
      pageIndex: 0,
      totalPages: 1,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    });

    const wrapped = lines(segments, 20, false);
    expect(wrapped.map(textOf)).toEqual(['甲乙', '丙）']);
    const first = wrapped[0].segments[0] as LayoutTextSeg;
    const second = wrapped[1].segments[0] as LayoutTextSeg;
    expect(compressionPt(first)).toBeUndefined();
    expect(first.measuredWidth).toBe(20);
    expect(compressionPt(second)).toBe(-5);
    expect(second.measuredWidth).toBe(15);
  });

  it('applies the document-wide punctuation compression in vertical CJK columns', () => {
    const text = `${'甲'.repeat(36)}。甲`;
    const compressed = buildSegments([textRun(text)], {
      pageIndex: 0,
      totalPages: 1,
      verticalCJK: true,
      characterSpacingControl: 'compressPunctuation',
      layoutServices: punctuationMetricsServices(),
    });
    // 37 full-em ideographs + the selected period's retained half-em cell = 375pt.
    expect(lines(compressed, 375, false)).toHaveLength(1);
    const compressedMark = compressed[0] as LayoutTextSeg;
    expect(compressedMark.text).toBe(text);
    expect(compressedMark.verticalRun).toBe(true);
    expect(compressionPt(compressedMark)).toBe(-5);

    const uncompressed = buildSegments([textRun(text)], {
      pageIndex: 0,
      totalPages: 1,
      verticalCJK: true,
      characterSpacingControl: 'doNotCompress',
    });
    expect(lines(uncompressed, 375, false)).toHaveLength(2);

    const middleDot = buildSegments([textRun(`${'甲'.repeat(36)}・甲`)], {
      pageIndex: 0,
      totalPages: 1,
      verticalCJK: true,
      characterSpacingControl: 'compressPunctuation',
    });
    expect(middleDot).toHaveLength(1);
    expect(compressionPt(middleDot[0] as LayoutTextSeg)).toBeUndefined();
    expect(lines(middleDot, 375, false)).toHaveLength(2);
  });

  it('uses the bidi language boundary for complex-script closing punctuation', () => {
    const latin = buildSegments([textRun('A B.', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    });

    // Office boundary controls in Latin Calibri/Arial keep . and , but wrap
    // the complete word ending in ) or } once its advance exceeds the line.
    expect(lines(latin, 30, true).map(textOf)).toEqual(['A B.']);
    expect(lines(buildSegments([textRun('A B,', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['A B,']);
    expect(lines(buildSegments([textRun('A B)', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['A ', 'B)']);
    expect(lines(buildSegments([textRun('A B}', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['A ', 'B}']);
    expect(lines(buildSegments([textRun('A B.', null)], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['A B.']);
    expect(lines(buildSegments([textRun('A B.', 'ja-jp')], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['A B.']);
    expect(lines(buildSegments([textRun('甲 A)', null)], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['甲 A)']);
    expect(lines(buildSegments([textRun('甲 A)', 'en-us')], {
      pageIndex: 0,
      totalPages: 1,
    }), 30, true).map(textOf)).toEqual(['甲 A)']);
    expect(wordIsOverflowPunctuation('.', 'en-us', false, true)).toBe(true);
    expect(wordIsOverflowPunctuation(',', 'en-us', false, true)).toBe(true);
    // Office-produced Calibri/Arial boundary controls wrap these Latin words
    // at their normal advance; the older ASCII-union projection overhung them.
    expect(wordIsOverflowPunctuation('}', undefined, false, true)).toBe(false);
    expect(wordIsOverflowPunctuation(')', 'en-us', false, true)).toBe(false);
    expect(wordIsOverflowPunctuation(':', 'en-us', false, true)).toBe(false);
    expect(wordIsOverflowPunctuation('>', 'en-us', false, true)).toBe(false);
    expect(wordIsOverflowPunctuation('.', 'ar-sa', false, false)).toBe(false);
    expect(wordIsOverflowPunctuation('.', 'en-us', false, false, true, undefined)).toBe(true);
    expect(wordIsOverflowPunctuation(':', 'en-us', false, false, true, 'en-us')).toBe(true);
    expect(wordIsOverflowPunctuation(')', 'en-us', false, false, true, 'ar-sa')).toBe(false);

    const complex = (bidiLanguage?: string) => buildSegments([{
      ...textRun('اب.', 'en-us'),
      rtl: true,
      langBidi: bidiLanguage,
    } as DocParagraph['runs'][number]], { pageIndex: 0, totalPages: 1 });
    expect(lines(complex(), 20, true).map(textOf)).toEqual(['اب.']);
    expect(lines(complex('ar-sa'), 20, true).map(textOf)).toEqual(['اب', '.']);
  });

  it('finds the final visible punctuation before a collapsible separator', () => {
    const enabled = buildSegments([textRun('A B. C', 'ja-jp')], {
      pageIndex: 0,
      totalPages: 1,
    });
    const disabled = buildSegments([textRun('A B. C', 'ja-jp')], {
      pageIndex: 0,
      totalPages: 1,
    });

    expect(lines(enabled, 30, true).map(textOf)).toEqual(['A B. ', 'C']);
    expect(lines(disabled, 30, false).map(textOf)).toEqual(['A ', 'B. ', 'C']);
  });

  it('limits U+3017 to the Simplified Chinese overflow-punctuation set', () => {
    expect(wordIsOverflowPunctuation('〗', 'zh-cn')).toBe(true);
    expect(wordIsOverflowPunctuation('〗', 'zh-hans')).toBe(true);
    expect(wordIsOverflowPunctuation('〗', 'zh-tw')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'zh-hant')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'zh-mo')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'ja-jp')).toBe(false);
    expect(wordIsOverflowPunctuation('〗', 'ko-kr')).toBe(false);
  });
});
