import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph, renderTable, type PptxTextRunInfo } from './renderer.js';
import type { Paragraph, TableCell, TableElement, TextBody } from './types.js';

// ECMA-376 §21.1.2.2.7 a:pPr@eaLnBrk. PowerPoint controls E00–E12 (issue
// #1562) show that eaLnBrk="0" still breaks ideographs per character; it lifts
// the East Asian line-start/line-end rules, so a CJK bracket may end a line
// and a break may follow an opening bracket before Latin text. Every glyph
// measures 10 px in these probes.
function measuringContext(): CanvasRenderingContext2D {
  let font = '';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
    canvas: { width: 1000, height: 1000 },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {}, transform() {},
    setLineDash() {}, getLineDash() { return []; }, drawImage() {},
    fillText() {}, strokeText() {},
    fillStyle: '', strokeStyle: '', lineWidth: 1, globalAlpha: 1,
    textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px',
  } as unknown as CanvasRenderingContext2D;
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000',
    fontFamily: 'Arial', fontFamilyEa: 'Meiryo',
  };
}

function paragraph(runs: TextRunData[], eaLnBrk: boolean): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk, runs,
  } as Paragraph;
}

function lines(runs: TextRunData[], width: number, eaLnBrk: boolean): string[] {
  return layoutParagraph(measuringContext(), paragraph(runs, eaLnBrk), width, 20, '000000', 1, 0)
    .map((line) => line.segments.map((segment) => segment.text).join(''));
}

describe('pptx eaLnBrk (§21.1.2.2.7)', () => {
  it('still breaks ideographs per character when eaLnBrk is false (control E00)', () => {
    expect(lines([run('日本語')], 10, false)).toEqual(['日', '本', '語']);
  });

  it('lets a CJK opening bracket end a line only when eaLnBrk is false (E01)', () => {
    expect(lines([run('日本語「'), run('日本語')], 20, false)).toEqual(['日本', '語「', '日本', '語']);
    expect(lines([run('日本語「'), run('日本語')], 20, true)).toEqual(['日本', '語', '「日', '本語']);
  });

  it('breaks after an opening bracket before Latin text only when eaLnBrk is false (E08, E09)', () => {
    expect(lines([run('「'), run('abc')], 30, false)).toEqual(['「', 'abc']);
    expect(lines([run('「'), run('abc')], 30, true)).toEqual(['「ab', 'c']);
  });

  it('honours eaLnBrk in table cell text', () => {
    const EMU = 12_700;
    const body = (eaLnBrk: boolean) => ({
      verticalAnchor: 't', paragraphs: [{ ...paragraph([run('「'), run('abc')], eaLnBrk) }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
    }) as unknown as TextBody;
    const cellTexts = (eaLnBrk: boolean): string[] => {
      const cell = {
        textBody: body(eaLnBrk), fill: null,
        borderL: null, borderR: null, borderT: null, borderB: null,
        diagonalTL: null, diagonalTR: null,
        gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false,
      } as TableCell;
      const table: TableElement = {
        type: 'table', x: 0, y: 0, width: 30 * EMU, height: 90 * EMU,
        rotation: 0, flipH: false, flipV: false,
        cols: [30 * EMU], rows: [{ height: 90 * EMU, cells: [cell] }],
      };
      const runs: PptxTextRunInfo[] = [];
      renderTable(measuringContext(), table, 1 / EMU, undefined,
        { themeMajorFont: null, themeMinorFont: null, dpr: 1 }, (info) => runs.push(info));
      // Group the emitted runs into visual lines by their vertical position.
      const byLine = new Map<number, string>();
      for (const info of runs) byLine.set(info.inShapeY, (byLine.get(info.inShapeY) ?? '') + info.text);
      return [...byLine.values()];
    };
    expect(cellTexts(false)).toEqual(['「', 'abc']);
    expect(cellTexts(true)).toEqual(['「ab', 'c']);
  });
});

describe('pptx kinsoku across run seams (issue #1653 controls K00–K23)', () => {
  // Each Office two-run case breaks exactly like its one-run twin. Curly quotes
  // paint in the Latin slot here, so their slot seam is a run seam as well.
  const lang = (text: string, value: string): TextRunData => ({ ...run(text), lang: value });

  it('keeps Japanese line-start and line-end rules continuous across authored and slot seams', () => {
    for (const parts of [['漢漢', '”漢'], ['漢漢”漢']]) {
      expect(lines(parts.map((text) => lang(text, 'ja-JP')), 20, true)).toEqual(['漢', '漢”', '漢']);
    }
    for (const parts of [['漢', '“漢漢'], ['漢“漢漢']]) {
      expect(lines(parts.map((text) => lang(text, 'ja-JP')), 20, true)).toEqual(['漢', '“漢', '漢']);
    }
  });

  it('leaves the curly quote at the seam under en-US', () => {
    expect(lines([lang('漢漢', 'en-US'), lang('”漢', 'en-US')], 20, true)).toEqual(['漢漢', '”漢']);
    expect(lines([lang('漢', 'en-US'), lang('“漢漢', 'en-US')], 20, true)).toEqual(['漢“', '漢漢']);
  });

  it('decides by language, not by whether the Latin and East Asian faces coincide', () => {
    // The controls' own setup: Arial in every slot, the quote inside one run.
    const arial = (text: string, value: string): TextRunData => ({ ...lang(text, value), fontFamilyEa: 'Arial' });
    expect(lines([arial('漢漢”漢', 'en-US')], 20, true)).toEqual(['漢漢', '”漢']);
    expect(lines([arial('漢“漢漢', 'en-US')], 20, true)).toEqual(['漢“', '漢漢']);
    expect(lines([arial('漢漢”漢', 'ja-JP')], 20, true)).toEqual(['漢', '漢”', '漢']);
    expect(lines([arial('漢“漢漢', 'ja-JP')], 20, true)).toEqual(['漢', '“漢', '漢']);
    // The seam is a break boundary only; one face still paints as one segment.
    const laid = layoutParagraph(measuringContext(), paragraph([arial('漢漢”漢', 'ja-JP')], true), 20, 20,
      '000000', 1, 0);
    expect(laid[1].segments.map((segment) => segment.text)).toEqual(['漢”']);
  });

  it('adds the one-face slot seam only for the measured en-US and ja-JP runs', () => {
    const arial = (text: string, value?: string): TextRunData =>
      ({ ...run(text), fontFamilyEa: 'Arial', ...(value ? { lang: value } : {}) });
    expect(lines([arial('漢漢”漢', 'EN-us')], 20, true)).toEqual(['漢漢', '”漢']);
    // Unmeasured languages, including other en/ja tags, keep one segment.
    for (const value of [undefined, 'ko-KR', 'zh-CN', 'fr-FR', 'en-GB']) {
      expect(lines([arial('漢漢”漢', value)], 20, true)).toEqual(['漢', '漢”', '漢']);
    }
  });

  it('keeps unobserved punctuation and mixed-language seams outside the new rule', () => {
    expect(lines([lang('漢漢', 'ja-JP'), lang('、漢', 'ja-JP')], 20, true)).toEqual(['漢漢', '、漢']);
    expect(lines([lang('漢“', 'en-US'), lang('）', 'ja-JP'), lang('”漢', 'ja-JP')], 30, true))
      .toEqual(['漢“）', '”漢']);
    expect(lines([lang('漢“', 'en-US'), lang('漢', 'ja-JP'), lang('”漢', 'ja-JP')], 30, true))
      .toEqual(['漢“漢', '”漢']);
  });
});

describe('pptx line feed inside a:t (controls L00–L08)', () => {
  it('breaks at the line feed; the next line takes the containing run size (L00)', () => {
    const laid = layoutParagraph(measuringContext(), paragraph([
      { ...run('Ab\nCd'), fontSize: 40 }, { ...run('Ef'), fontSize: 14 },
    ], true), 1000, 20, '000000', 1, 0);
    expect(laid.map((line) => line.segments.map((segment) => segment.text).join(''))).toEqual(['Ab', 'CdEf']);
    expect(Math.max(...laid[1].segments.map((segment) => segment.sizePx))).toBe(40 * 12_700);
  });

  it('does not size the next line by a run that ends at the line feed (L02)', () => {
    // Office L02: 40 pt `Ab⏎` then 14 pt `Cd`; the second line is a 14 pt line.
    const laid = layoutParagraph(measuringContext(), paragraph([
      { ...run('Ab\n'), fontSize: 40 }, { ...run('Cd'), fontSize: 14 },
    ], true), 1000, 20, '000000', 1, 0);
    expect(laid.map((line) => line.segments.map((segment) => segment.text).join(''))).toEqual(['Ab', 'Cd']);
    expect(laid[1].segments.map((segment) => segment.sizePx)).toEqual([14 * 12_700]);
  });

  it('sizes an empty line opened by a line feed by that run (L07, L08)', () => {
    const laid = layoutParagraph(measuringContext(), paragraph([{ ...run('Ab\n\nCd'), fontSize: 40 }], true),
      1000, 20, '000000', 1, 0);
    expect(laid.map((line) => line.segments.map((segment) => segment.text).join(''))).toEqual(['Ab', '', 'Cd']);
    expect(laid[1].segments.map((segment) => segment.sizePx)).toEqual([40 * 12_700]);
  });
});
